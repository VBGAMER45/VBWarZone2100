Attribute VB_Name = "silk_dec_API_c"
' libopus 1.5.2 (ref/opus/silk): dec_API.c (and API.h's decoder declarations), the float build: no OSCE, no
' DEEP_PLC (silk_LoadOSCEModels does nothing). silk_get_TOC is #if 0 in the C.
' (VB) void *decState is a silk_decoder (the opus decoder holds one). The output buffers are an array and an
' offset. C keeps the decoded frames before resampling either in the output buffer or in a stack buffer
' (delay_stack_alloc, a stack-saving choice with the same result): here always in samplesOut1_tmp_storage.
Option Explicit

'/************************/
'/* Decoder Super Struct */
'/************************/
Public Type silk_decoder
    channel_state(0 To 1) As silk_decoder_state         ' [ DECODER_NUM_CHANNELS ]
    sStereo As stereo_dec_state
    nChannelsAPI As Long
    nChannelsInternal As Long
    prev_decode_only_middle As Long
End Type

' (VB) silk_Decode's stack buffers
Private pulses(0 To 319) As Integer                     ' [ MAX_FRAME_LENGTH ]
Private samplesOut1_tmp_storage(0 To 643) As Integer    ' nChannelsInternal * ( frame_length + 2 )
Private samplesOut2_tmp(0 To 959) As Integer            ' *nSamplesOut (20 ms at 48 kHz)

'/*********************/
'/* Decoder functions */
'/*********************/

Public Function silk_LoadOSCEModels(decState As silk_decoder, data() As Byte, ByVal len_ As Long) As Long
    silk_LoadOSCEModels = SILK_NO_ERROR
End Function

Public Function silk_Get_Decoder_Size(decSizeBytes As Long) As Long  '/* O    Number of bytes in SILK decoder state           */
    Dim d As silk_decoder
    Dim ret As Long
    ret = SILK_NO_ERROR

    decSizeBytes = LenB(d)

    silk_Get_Decoder_Size = ret
End Function

'/* Reset decoder state */
Public Function silk_ResetDecoder(decState As silk_decoder) As Long   '/* I/O  State                                           */
    Dim n As Long, ret As Long
    Static blank As stereo_dec_state
    ret = SILK_NO_ERROR

    For n = 0 To DECODER_NUM_CHANNELS - 1
        ret = silk_reset_decoder(decState.channel_state(n))
    Next
    decState.sStereo = blank
    '/* Not strictly needed, but it's cleaner that way */
    decState.prev_decode_only_middle = 0

    silk_ResetDecoder = ret
End Function

Public Function silk_InitDecoder(decState As silk_decoder) As Long    '/* I/O  State                                           */
    Dim n As Long, ret As Long
    Static blank As stereo_dec_state
    Dim noData() As Byte
    ret = SILK_NO_ERROR
    silk_tables_init                                    ' (VB) the ROM tables, once
    '/* load osce models */
    silk_LoadOSCEModels decState, noData, 0             ' (data, len) = (NULL, 0)

    For n = 0 To DECODER_NUM_CHANNELS - 1
        ret = silk_init_decoder(decState.channel_state(n))
    Next
    decState.sStereo = blank
    '/* Not strictly needed, but it's cleaner that way */
    decState.prev_decode_only_middle = 0

    silk_InitDecoder = ret
End Function

'/* Decode a frame */
Public Function silk_Decode(decState As silk_decoder, decControl As silk_DecControlStruct, ByVal lostFlag As Long, _
        ByVal newPacketFlag As Long, psRangeDec As ec_ctx, samplesOut() As Integer, ByVal samplesOutOffset As Long, _
        nSamplesOut As Long, ByVal arch As Long) As Long
    Dim i As Long, n As Long, decode_only_middle As Long, ret As Long
    Dim nSamplesOutDec As Long, LBRR_symbol As Long
    Dim samplesOut1_tmp(0 To 1) As Long                 ' opus_int16 *[ 2 ]: offsets into samplesOut1_tmp_storage
    Dim MS_pred_Q13(0 To 1) As Long
    Dim has_side As Long
    Dim stereo_to_mono As Long
    Dim fs_kHz_dec As Long
    Dim condCoding As Long
    Dim FrameIndex As Long
    Dim mult_tab(0 To 2) As Long
    Dim apiToStereo As Boolean

    ret = SILK_NO_ERROR
    ' celt_assert( decControl->nChannelsInternal == 1 || decControl->nChannelsInternal == 2 );

    With decState
    '/**********************************/
    '/* Test if first frame in payload */
    '/**********************************/
    If newPacketFlag <> 0 Then
        For n = 0 To decControl.nChannelsInternal - 1
            .channel_state(n).nFramesDecoded = 0        '/* Used to count frames in packet */
        Next
    End If

    '/* If Mono -> Stereo transition in bitstream: init state of second channel */
    If decControl.nChannelsInternal > .nChannelsInternal Then
        ret = ret + silk_init_decoder(.channel_state(1))
    End If

    If decControl.nChannelsInternal = 1 And .nChannelsInternal = 2 And _
        (decControl.internalSampleRate = 1000 * .channel_state(0).fs_kHz) Then stereo_to_mono = 1

    If .channel_state(0).nFramesDecoded = 0 Then
        For n = 0 To decControl.nChannelsInternal - 1
            If decControl.payloadSize_ms = 0 Then
                '/* Assuming packet loss, use 10 ms */
                .channel_state(n).nFramesPerPacket = 1
                .channel_state(n).nb_subfr = 2
            ElseIf decControl.payloadSize_ms = 10 Then
                .channel_state(n).nFramesPerPacket = 1
                .channel_state(n).nb_subfr = 2
            ElseIf decControl.payloadSize_ms = 20 Then
                .channel_state(n).nFramesPerPacket = 1
                .channel_state(n).nb_subfr = 4
            ElseIf decControl.payloadSize_ms = 40 Then
                .channel_state(n).nFramesPerPacket = 2
                .channel_state(n).nb_subfr = 4
            ElseIf decControl.payloadSize_ms = 60 Then
                .channel_state(n).nFramesPerPacket = 3
                .channel_state(n).nb_subfr = 4
            Else
                ' celt_assert( 0 );
                silk_Decode = SILK_DEC_INVALID_FRAME_SIZE
                Exit Function
            End If
            fs_kHz_dec = silk_RSHIFT(decControl.internalSampleRate, 10) + 1
            If fs_kHz_dec <> 8 And fs_kHz_dec <> 12 And fs_kHz_dec <> 16 Then
                ' celt_assert( 0 );
                silk_Decode = SILK_DEC_INVALID_SAMPLING_FREQUENCY
                Exit Function
            End If
            ret = ret + silk_decoder_set_fs(.channel_state(n), fs_kHz_dec, decControl.API_sampleRate)
        Next
    End If

    If decControl.nChannelsAPI = 2 And decControl.nChannelsInternal = 2 And (.nChannelsAPI = 1 Or .nChannelsInternal = 1) Then
        .sStereo.pred_prev_Q13(0) = 0: .sStereo.pred_prev_Q13(1) = 0
        .sStereo.sSide(0) = 0: .sStereo.sSide(1) = 0
        .channel_state(1).resampler_state = .channel_state(0).resampler_state
    End If
    .nChannelsAPI = decControl.nChannelsAPI
    .nChannelsInternal = decControl.nChannelsInternal

    If decControl.API_sampleRate > MAX_API_FS_KHZ * 1000 Or decControl.API_sampleRate < 8000 Then
        ret = SILK_DEC_INVALID_SAMPLING_FREQUENCY
        silk_Decode = ret
        Exit Function
    End If

    If lostFlag <> FLAG_PACKET_LOST And .channel_state(0).nFramesDecoded = 0 Then
        '/* First decoder call for this payload */
        '/* Decode VAD flags and LBRR flag */
        For n = 0 To decControl.nChannelsInternal - 1
            For i = 0 To .channel_state(n).nFramesPerPacket - 1
                .channel_state(n).VAD_flags(i) = ec_dec_bit_logp(psRangeDec, 1)
            Next
            .channel_state(n).LBRR_flag = ec_dec_bit_logp(psRangeDec, 1)
        Next
        '/* Decode LBRR flags */
        For n = 0 To decControl.nChannelsInternal - 1
            .channel_state(n).LBRR_flags(0) = 0: .channel_state(n).LBRR_flags(1) = 0: .channel_state(n).LBRR_flags(2) = 0
            If .channel_state(n).LBRR_flag <> 0 Then
                If .channel_state(n).nFramesPerPacket = 1 Then
                    .channel_state(n).LBRR_flags(0) = 1
                Else
                    LBRR_symbol = ec_dec_icdf(psRangeDec, silk_rom_u8, silk_LBRR_flags_iCDF_ptr(.channel_state(n).nFramesPerPacket - 2), 8) + 1
                    For i = 0 To .channel_state(n).nFramesPerPacket - 1
                        .channel_state(n).LBRR_flags(i) = silk_RSHIFT(LBRR_symbol, i) And 1
                    Next
                End If
            End If
        Next

        If lostFlag = FLAG_DECODE_NORMAL Then
            '/* Regular decoding: skip all LBRR data */
            For i = 0 To .channel_state(0).nFramesPerPacket - 1
                For n = 0 To decControl.nChannelsInternal - 1
                    If .channel_state(n).LBRR_flags(i) <> 0 Then
                        If decControl.nChannelsInternal = 2 And n = 0 Then
                            silk_stereo_decode_pred psRangeDec, MS_pred_Q13
                            If .channel_state(1).LBRR_flags(i) = 0 Then
                                silk_stereo_decode_mid_only psRangeDec, decode_only_middle
                            End If
                        End If
                        '/* Use conditional coding if previous frame available */
                        condCoding = CODE_INDEPENDENTLY
                        If i > 0 Then
                            If .channel_state(n).LBRR_flags(i - 1) <> 0 Then condCoding = CODE_CONDITIONALLY
                        End If
                        silk_decode_indices .channel_state(n), psRangeDec, i, 1, condCoding
                        silk_decode_pulses psRangeDec, pulses, .channel_state(n).indices.signalType, _
                            .channel_state(n).indices.quantOffsetType, .channel_state(n).frame_length
                    End If
                Next
            Next
        End If
    End If

    '/* Get MS predictor index */
    If decControl.nChannelsInternal = 2 Then
        If lostFlag = FLAG_DECODE_NORMAL Or _
            (lostFlag = FLAG_DECODE_LBRR And .channel_state(0).LBRR_flags(.channel_state(0).nFramesDecoded) = 1) Then
            silk_stereo_decode_pred psRangeDec, MS_pred_Q13
            '/* For LBRR data, decode mid-only flag only if side-channel's LBRR flag is false */
            If (lostFlag = FLAG_DECODE_NORMAL And .channel_state(1).VAD_flags(.channel_state(0).nFramesDecoded) = 0) Or _
                (lostFlag = FLAG_DECODE_LBRR And .channel_state(1).LBRR_flags(.channel_state(0).nFramesDecoded) = 0) Then
                silk_stereo_decode_mid_only psRangeDec, decode_only_middle
            Else
                decode_only_middle = 0
            End If
        Else
            For n = 0 To 1
                MS_pred_Q13(n) = .sStereo.pred_prev_Q13(n)
            Next
        End If
    End If

    '/* Reset side channel decoder prediction memory for first frame with side coding */
    If decControl.nChannelsInternal = 2 And decode_only_middle = 0 And .prev_decode_only_middle = 1 Then
        silk_memset_i16 .channel_state(1).outBuf, 0, MAX_FRAME_LENGTH + 2 * MAX_SUB_FRAME_LENGTH
        silk_memset_i32 .channel_state(1).sLPC_Q14_buf, 0, MAX_LPC_ORDER
        .channel_state(1).lagPrev = 100
        .channel_state(1).LastGainIndex = 10
        .channel_state(1).prevSignalType = TYPE_NO_VOICE_ACTIVITY
        .channel_state(1).first_frame_after_reset = 1
    End If

    '/* Check if the temp buffer fits into the output PCM buffer. If it fits,
    '   we can delay allocating the temp buffer until after the SILK peak stack
    '   usage. We need to use a < and not a <= because of the two extra samples. */
    ' (VB) always the stack buffer (samplesOut1_tmp_storage): the same samples either way
    samplesOut1_tmp(0) = 0
    samplesOut1_tmp(1) = .channel_state(0).frame_length + 2

    If lostFlag = FLAG_DECODE_NORMAL Then
        If decode_only_middle = 0 Then has_side = 1
    Else
        If .prev_decode_only_middle = 0 Then
            has_side = 1
        ElseIf decControl.nChannelsInternal = 2 And lostFlag = FLAG_DECODE_LBRR Then
            If .channel_state(1).LBRR_flags(.channel_state(1).nFramesDecoded) = 1 Then has_side = 1
        End If
    End If
    .channel_state(0).sPLC.enable_deep_plc = decControl.enable_deep_plc
    '/* Call decoder for one frame */
    For n = 0 To decControl.nChannelsInternal - 1
        If n = 0 Or has_side <> 0 Then
            FrameIndex = .channel_state(0).nFramesDecoded - n
            '/* Use independent coding if no previous frame available */
            If FrameIndex <= 0 Then
                condCoding = CODE_INDEPENDENTLY
            ElseIf lostFlag = FLAG_DECODE_LBRR Then
                If .channel_state(n).LBRR_flags(FrameIndex - 1) <> 0 Then condCoding = CODE_CONDITIONALLY Else condCoding = CODE_INDEPENDENTLY
            ElseIf n > 0 And .prev_decode_only_middle <> 0 Then
                '/* If we skipped a side frame in this packet, we don't
                '   need LTP scaling; the LTP state is well-defined. */
                condCoding = CODE_INDEPENDENTLY_NO_LTP_SCALING
            Else
                condCoding = CODE_CONDITIONALLY
            End If
            ret = ret + silk_decode_frame(.channel_state(n), psRangeDec, samplesOut1_tmp_storage, samplesOut1_tmp(n) + 2, _
                nSamplesOutDec, lostFlag, condCoding, arch)
        Else
            silk_memset_i16 samplesOut1_tmp_storage, samplesOut1_tmp(n) + 2, nSamplesOutDec
        End If
        .channel_state(n).nFramesDecoded = .channel_state(n).nFramesDecoded + 1
    Next

    If decControl.nChannelsAPI = 2 And decControl.nChannelsInternal = 2 Then
        '/* Convert Mid/Side to Left/Right */
        silk_stereo_MS_to_LR .sStereo, samplesOut1_tmp_storage, samplesOut1_tmp(0), samplesOut1_tmp_storage, samplesOut1_tmp(1), _
            MS_pred_Q13, .channel_state(0).fs_kHz, nSamplesOutDec
    Else
        '/* Buffering */
        samplesOut1_tmp_storage(samplesOut1_tmp(0)) = .sStereo.sMid(0)
        samplesOut1_tmp_storage(samplesOut1_tmp(0) + 1) = .sStereo.sMid(1)
        .sStereo.sMid(0) = samplesOut1_tmp_storage(samplesOut1_tmp(0) + nSamplesOutDec)
        .sStereo.sMid(1) = samplesOut1_tmp_storage(samplesOut1_tmp(0) + nSamplesOutDec + 1)
    End If

    '/* Number of output samples */
    nSamplesOut = (nSamplesOutDec * decControl.API_sampleRate) \ silk_SMULBB(.channel_state(0).fs_kHz, 1000)

    '/* Set up pointers to temp buffers */
    ' (VB) resample_out_ptr is samplesOut2_tmp when nChannelsAPI == 2, else samplesOut (from samplesOutOffset)
    apiToStereo = (decControl.nChannelsAPI = 2)

    For n = 0 To silk_min(decControl.nChannelsAPI, decControl.nChannelsInternal) - 1

        '/* Resample decoded signal to API_sampleRate */
        If apiToStereo Then
            ret = ret + silk_resampler(.channel_state(n).resampler_state, samplesOut2_tmp, 0, samplesOut1_tmp_storage, samplesOut1_tmp(n) + 1, nSamplesOutDec)
        Else
            ret = ret + silk_resampler(.channel_state(n).resampler_state, samplesOut, samplesOutOffset, samplesOut1_tmp_storage, samplesOut1_tmp(n) + 1, nSamplesOutDec)
        End If

        '/* Interleave if stereo output and stereo stream */
        If apiToStereo Then
            For i = 0 To nSamplesOut - 1
                samplesOut(samplesOutOffset + n + 2 * i) = samplesOut2_tmp(i)
            Next
        End If
    Next

    '/* Create two channel output from mono stream */
    If decControl.nChannelsAPI = 2 And decControl.nChannelsInternal = 1 Then
        If stereo_to_mono <> 0 Then
            '/* Resample right channel for newly collapsed stereo just in case
            '   we weren't doing collapsing when switching to mono */
            ret = ret + silk_resampler(.channel_state(1).resampler_state, samplesOut2_tmp, 0, samplesOut1_tmp_storage, samplesOut1_tmp(0) + 1, nSamplesOutDec)

            For i = 0 To nSamplesOut - 1
                samplesOut(samplesOutOffset + 1 + 2 * i) = samplesOut2_tmp(i)
            Next
        Else
            For i = 0 To nSamplesOut - 1
                samplesOut(samplesOutOffset + 1 + 2 * i) = samplesOut(samplesOutOffset + 0 + 2 * i)
            Next
        End If
    End If

    '/* Export pitch lag, measured at 48 kHz sampling rate */
    If .channel_state(0).prevSignalType = TYPE_VOICED Then
        mult_tab(0) = 6: mult_tab(1) = 4: mult_tab(2) = 3
        decControl.prevPitchLag = .channel_state(0).lagPrev * mult_tab(silk_RSHIFT(.channel_state(0).fs_kHz - 8, 2))
    Else
        decControl.prevPitchLag = 0
    End If

    If lostFlag = FLAG_PACKET_LOST Then
       '/* On packet loss, remove the gain clamping to prevent having the energy "bounce back"
       '   if we lose packets when the energy is going down */
       For i = 0 To .nChannelsInternal - 1
          .channel_state(i).LastGainIndex = 10
       Next
    Else
       .prev_decode_only_middle = decode_only_middle
    End If
    End With
    silk_Decode = ret
End Function
