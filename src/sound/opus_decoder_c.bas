Attribute VB_Name = "opus_decoder_c"
' libopus 1.5.2 (ref/opus/src): opus_decoder.c, the float build (FIXED_POINT undefined: opus_decode_float is
' the native decoder, opus_decode converts with FLOAT2INT16 after a soft clip), without DEEP_PLC, DRED or OSCE
' (the OpusDRED functions are not ported; opus_decode_native's dred and dred_offset arguments are dropped),
' no hardening (VALIDATE_OPUS_DECODER is empty) and no assertions (celt_assert is a comment; MUST_SUCCEED returns
' OPUS_INTERNAL_ERROR, as arch.h has it then).
' (VB) An OpusDecoder * is a handle: an index of this module's pool (0 = NULL). The C carves the SILK decoder and
' the CELT decoder out of the OpusDecoder's memory (silk_dec_offset, celt_dec_offset): here the OpusDecoder
' holds its silk_decoder and O2's CELTDecoder handle. A packet is an array and an offset; C's data == NULL is
' dataIsNull (the public functions: len == 0); PCM is an array and an offset. The stack buffers are local
' dynamic arrays (opus_decode_frame calls itself). The float arithmetic that must round as C's is float_vb's
' (smooth_fade, the decode gain); the SILK samples are added as C does in one rounding (an exact product by
' 2^-15 and one sum, stored as a Single).
Option Explicit

Public Type OpusDecoder
    celt_dec As Long                                    ' CELTDecoder * (C: int celt_dec_offset)
    silk_dec As silk_decoder                            ' (C: int silk_dec_offset: the state that follows the struct)
    channels As Long
    Fs As Long                                          '/** Sampling rate (at the API level) */
    DecControl As silk_DecControlStruct
    decode_gain As Long
    complexity As Long
    arch As Long

    '/* Everything beyond this point gets cleared on a reset */
    ' #define OPUS_DECODER_RESET_START stream_channels
    stream_channels As Long

    bandwidth As Long
    mode As Long
    prev_mode As Long
    frame_size As Long
    prev_redundancy As Long
    last_packet_duration As Long
    softclip_mem(0 To 1) As Single                      ' opus_val16 (#ifndef FIXED_POINT)

    rangeFinal As Long                                  ' opus_uint32
    vbInUse As Boolean                                  ' (VB) the pool slot is taken
End Type

Private decs() As OpusDecoder
Private decs_count As Long

' (VB) a pool slot (opus_alloc): its index, with its CELT decoder
Private Function opus_decoder_vbAlloc() As Long
    Dim i As Long
    For i = 1 To decs_count
        If Not decs(i).vbInUse Then Exit For
    Next
    If i > decs_count Then
        decs_count = decs_count + 4
        ReDim Preserve decs(1 To decs_count)
    End If
    decs(i).vbInUse = True
    If decs(i).celt_dec = 0 Then decs(i).celt_dec = celt_decoder_vbNew()
    opus_decoder_vbAlloc = i
End Function

' (VB) opus_free
Private Sub opus_decoder_vbFree(ByVal st As Long)
    If st < 1 Or st > decs_count Then Exit Sub
    If decs(st).celt_dec <> 0 Then celt_decoder_vbFree decs(st).celt_dec
    decs(st).celt_dec = 0
    decs(st).vbInUse = False
End Sub

Public Function opus_decoder_get_size(ByVal channels As Long) As Long
    Dim silkDecSizeBytes As Long, celtDecSizeBytes As Long
    Dim ret As Long
    Dim d As OpusDecoder
    If channels < 1 Or channels > 2 Then
        Exit Function
    End If
    ret = silk_Get_Decoder_Size(silkDecSizeBytes)
    If ret <> 0 Then
        Exit Function
    End If
    silkDecSizeBytes = align(silkDecSizeBytes)
    celtDecSizeBytes = celt_decoder_get_size(channels)
    opus_decoder_get_size = align(LenB(d) - silkDecSizeBytes) + silkDecSizeBytes + celtDecSizeBytes
End Function

Public Function opus_decoder_init(ByVal st As Long, ByVal Fs As Long, ByVal channels As Long) As Long
    Dim ret As Long, silkDecSizeBytes As Long
    Dim celt_dec As Long
    Static blank As OpusDecoder

    If (Fs <> 48000 And Fs <> 24000 And Fs <> 16000 And Fs <> 12000 And Fs <> 8000) _
        Or (channels <> 1 And channels <> 2) Then
        opus_decoder_init = OPUS_BAD_ARG
        Exit Function
    End If
    If st < 1 Or st > decs_count Then                   ' (VB) not a decoder of this pool
        opus_decoder_init = OPUS_BAD_ARG
        Exit Function
    End If

    ' OPUS_CLEAR((char*)st, opus_decoder_get_size(channels)); (VB) the CELT decoder is kept: celt_decoder_init clears it
    celt_dec = decs(st).celt_dec
    decs(st) = blank
    decs(st).celt_dec = celt_dec
    decs(st).vbInUse = True
    '/* Initialize SILK decoder */
    ret = silk_Get_Decoder_Size(silkDecSizeBytes)
    If ret <> 0 Then
        opus_decoder_init = OPUS_INTERNAL_ERROR
        Exit Function
    End If

    decs(st).channels = channels
    decs(st).stream_channels = channels
    decs(st).complexity = 0

    decs(st).Fs = Fs
    decs(st).DecControl.API_sampleRate = decs(st).Fs
    decs(st).DecControl.nChannelsAPI = decs(st).channels

    '/* Reset decoder */
    ret = silk_InitDecoder(decs(st).silk_dec)
    If ret <> 0 Then
        opus_decoder_init = OPUS_INTERNAL_ERROR
        Exit Function
    End If

    '/* Initialize CELT decoder */
    ret = celt_decoder_init(celt_dec, Fs, channels)
    If ret <> OPUS_OK Then
        opus_decoder_init = OPUS_INTERNAL_ERROR
        Exit Function
    End If

    celt_decoder_ctl celt_dec, CELT_SET_SIGNALLING_REQUEST, 0

    decs(st).prev_mode = 0
    decs(st).frame_size = Fs \ 400
    decs(st).arch = 0                                   ' opus_select_arch(): the C paths only
    opus_decoder_init = OPUS_OK
End Function

Public Function opus_decoder_create(ByVal Fs As Long, ByVal channels As Long, error_ As Long) As Long
    Dim ret As Long
    Dim st As Long
    If (Fs <> 48000 And Fs <> 24000 And Fs <> 16000 And Fs <> 12000 And Fs <> 8000) _
        Or (channels <> 1 And channels <> 2) Then
        error_ = OPUS_BAD_ARG
        opus_decoder_create = 0
        Exit Function
    End If
    st = opus_decoder_vbAlloc()
    If st = 0 Then
        error_ = OPUS_ALLOC_FAIL
        Exit Function
    End If
    ret = opus_decoder_init(st, Fs, channels)
    error_ = ret
    If ret <> OPUS_OK Then
        opus_decoder_vbFree st
        st = 0
    End If
    opus_decoder_create = st
End Function

Private Sub smooth_fade(in1() As Single, ByVal in1Offset As Long, in2() As Single, ByVal in2Offset As Long, _
        out_() As Single, ByVal outOffset As Long, ByVal overlap As Long, ByVal channels As Long, _
        window() As Single, ByVal Fs As Long)
    Dim i As Long, c As Long
    Dim inc As Long
    Dim w As Single
    inc = 48000 \ Fs
    For c = 0 To channels - 1
        For i = 0 To overlap - 1
            w = fmul(window(i * inc), window(i * inc))  ' MULT16_16_Q15(window[i*inc], window[i*inc])
            ' SHR32(MAC16_16(MULT16_16(w,in2[i*channels+c]), Q15ONE-w, in1[i*channels+c]), 15)
            out_(outOffset + i * channels + c) = fadd(fmul(w, in2(in2Offset + i * channels + c)), _
                                                      fmul(fsub(1!, w), in1(in1Offset + i * channels + c)))
        Next
    Next
End Sub

Private Function opus_packet_get_mode(data() As Byte, ByVal dataOffset As Long) As Long
    Dim mode As Long
    If (data(dataOffset) And &H80) <> 0 Then
        mode = MODE_CELT_ONLY
    ElseIf (data(dataOffset) And &H60) = &H60 Then
        mode = MODE_HYBRID
    Else
        mode = MODE_SILK_ONLY
    End If
    opus_packet_get_mode = mode
End Function

Private Function opus_decode_frame(ByVal st As Long, data() As Byte, ByVal dataOffset As Long, ByVal dataIsNull As Boolean, _
        ByVal len_ As Long, pcm() As Single, ByVal pcmOffset As Long, ByVal frame_size As Long, ByVal decode_fec As Long) As Long
    Dim celt_dec As Long
    Dim i As Long, silk_ret As Long, celt_ret As Long
    Dim dec As ec_ctx
    Dim silk_frame_size As Long
    Dim pcm_silk_size As Long
    Dim pcm_silk() As Integer
    Dim pcm_transition_silk_size As Long
    Dim pcm_transition_silk() As Single
    Dim pcm_transition_celt_size As Long
    Dim pcm_transition_celt() As Single
    Dim pcm_transition As Long                          ' opus_val16 *: 0 = NULL, 1 = pcm_transition_celt, 2 = pcm_transition_silk
    Dim redundant_audio_size As Long
    Dim redundant_audio() As Single

    Dim audiosize As Long
    Dim mode As Long
    Dim bandwidth As Long
    Dim transition As Long
    Dim start_band As Long
    Dim redundancy As Long
    Dim redundancy_bytes As Long
    Dim celt_to_silk As Long
    Dim c As Long
    Dim F2_5 As Long, F5 As Long, F10 As Long, F20 As Long
    Dim window() As Single                              ' const opus_val16 *window (celt_mode->window)
    Dim overlap As Long
    Dim redundant_rng As Long                           ' opus_uint32
    Dim celt_accum As Long
    Dim ret As Long
    Dim lost_flag As Long, decoded_samples As Long
    Dim pcm_ptr As Long                                 ' opus_int16 *: an offset into pcm_silk
    Dim first_frame As Long
    Dim endband As Long
    Dim celt_frame_size As Long
    Dim silence(0 To 1) As Byte
    Dim celt_mode As Long                               ' const CELTMode *: O2's mode handle
    Dim gain As Single
    Dim noDec As ec_ctx

    celt_dec = decs(st).celt_dec
    F20 = decs(st).Fs \ 50
    F10 = F20 \ 2
    F5 = F10 \ 2
    F2_5 = F5 \ 2
    If frame_size < F2_5 Then
        opus_decode_frame = OPUS_BUFFER_TOO_SMALL
        Exit Function
    End If
    '/* Limit frame_size to avoid excessive stack allocations. */
    If (decs(st).Fs \ 25) * 3 < frame_size Then frame_size = (decs(st).Fs \ 25) * 3     ' IMIN(frame_size, st->Fs/25*3)
    '/* Payloads of 1 (2 including ToC) or 0 trigger the PLC/DTX */
    If len_ <= 1 Then
        dataIsNull = True
        '/* In that case, don't conceal more than what the ToC says */
        If decs(st).frame_size < frame_size Then frame_size = decs(st).frame_size        ' IMIN(frame_size, st->frame_size)
    End If
    If Not dataIsNull Then
        audiosize = decs(st).frame_size
        mode = decs(st).mode
        bandwidth = decs(st).bandwidth
        ec_dec_init dec, data, dataOffset, len_
    Else
        audiosize = frame_size
        '/* Run PLC using last used mode (CELT if we ended with CELT redundancy) */
        If decs(st).prev_redundancy <> 0 Then mode = MODE_CELT_ONLY Else mode = decs(st).prev_mode
        bandwidth = 0

        If mode = 0 Then
            '/* If we haven't got any packet yet, all we can do is return zeros */
            For i = 0 To audiosize * decs(st).channels - 1
                pcm(pcmOffset + i) = 0
            Next
            opus_decode_frame = audiosize
            Exit Function
        End If

        '/* Avoids trying to run the PLC on sizes other than 2.5 (CELT), 5 (CELT),
        '   10, or 20 (e.g. 12.5 or 30 ms). */
        If audiosize > F20 Then
            Do
                If audiosize < F20 Then ret = audiosize Else ret = F20                 ' IMIN(audiosize, F20)
                ret = opus_decode_frame(st, data, 0, True, 0, pcm, pcmOffset, ret, 0)
                If ret < 0 Then
                    opus_decode_frame = ret
                    Exit Function
                End If
                pcmOffset = pcmOffset + ret * decs(st).channels
                audiosize = audiosize - ret
            Loop While audiosize > 0
            opus_decode_frame = frame_size
            Exit Function
        ElseIf audiosize < F20 Then
            If audiosize > F10 Then
                audiosize = F10
            ElseIf mode <> MODE_SILK_ONLY And audiosize > F5 And audiosize < F10 Then
                audiosize = F5
            End If
        End If
    End If

    '/* In fixed-point, we can tell CELT to do the accumulation on top of the
    '   SILK PCM buffer. This saves some stack space. */
    celt_accum = 0

    pcm_transition_silk_size = 0                        ' ALLOC_NONE
    pcm_transition_celt_size = 0                        ' ALLOC_NONE
    If Not dataIsNull And decs(st).prev_mode > 0 Then
        If (mode = MODE_CELT_ONLY And decs(st).prev_mode <> MODE_CELT_ONLY And decs(st).prev_redundancy = 0) _
            Or (mode <> MODE_CELT_ONLY And decs(st).prev_mode = MODE_CELT_ONLY) Then
            transition = 1
            '/* Decide where to allocate the stack memory for pcm_transition */
            If mode = MODE_CELT_ONLY Then
                pcm_transition_celt_size = F5 * decs(st).channels
            Else
                pcm_transition_silk_size = F5 * decs(st).channels
            End If
        End If
    End If
    If pcm_transition_celt_size > 0 Then ReDim pcm_transition_celt(0 To pcm_transition_celt_size - 1)
    If transition <> 0 And mode = MODE_CELT_ONLY Then
        pcm_transition = 1
        If F5 < audiosize Then ret = F5 Else ret = audiosize                            ' IMIN(F5, audiosize)
        opus_decode_frame st, data, 0, True, 0, pcm_transition_celt, 0, ret, 0
    End If
    If audiosize > frame_size Then
        '/*fprintf(stderr, "PCM buffer too small: %d vs %d (mode = %d)\n", audiosize, frame_size, mode);*/
        opus_decode_frame = OPUS_BAD_ARG
        Exit Function
    Else
        frame_size = audiosize
    End If

    '/* Don't allocate any memory when in CELT-only mode */
    If mode <> MODE_CELT_ONLY And celt_accum = 0 Then
        If F10 > frame_size Then pcm_silk_size = F10 * decs(st).channels Else pcm_silk_size = frame_size * decs(st).channels    ' IMAX(F10, frame_size)*st->channels
    End If
    If pcm_silk_size > 0 Then ReDim pcm_silk(0 To pcm_silk_size - 1)

    '/* SILK processing */
    If mode <> MODE_CELT_ONLY Then
        pcm_ptr = 0

        If decs(st).prev_mode = MODE_CELT_ONLY Then
            silk_ResetDecoder decs(st).silk_dec
        End If

        '/* The SILK PLC cannot produce frames of less than 10 ms */
        decs(st).DecControl.payloadSize_ms = (1000 * audiosize) \ decs(st).Fs
        If decs(st).DecControl.payloadSize_ms < 10 Then decs(st).DecControl.payloadSize_ms = 10   ' IMAX(10, 1000 * audiosize / st->Fs)

        If Not dataIsNull Then
            decs(st).DecControl.nChannelsInternal = decs(st).stream_channels
            If mode = MODE_SILK_ONLY Then
                If bandwidth = OPUS_BANDWIDTH_NARROWBAND Then
                    decs(st).DecControl.internalSampleRate = 8000
                ElseIf bandwidth = OPUS_BANDWIDTH_MEDIUMBAND Then
                    decs(st).DecControl.internalSampleRate = 12000
                ElseIf bandwidth = OPUS_BANDWIDTH_WIDEBAND Then
                    decs(st).DecControl.internalSampleRate = 16000
                Else
                    decs(st).DecControl.internalSampleRate = 16000
                    ' celt_assert( 0 );
                End If
            Else
                '/* Hybrid mode */
                decs(st).DecControl.internalSampleRate = 16000
            End If
        End If
        If decs(st).complexity >= 5 Then decs(st).DecControl.enable_deep_plc = 1 Else decs(st).DecControl.enable_deep_plc = 0

        If dataIsNull Then
            lost_flag = 1
        ElseIf decode_fec <> 0 Then
            lost_flag = 2                               ' 2 * !!decode_fec
        Else
            lost_flag = 0
        End If
        decoded_samples = 0
        Do
            '/* Call SILK decoder */
            If decoded_samples = 0 Then first_frame = 1 Else first_frame = 0
            silk_ret = silk_Decode(decs(st).silk_dec, decs(st).DecControl, lost_flag, first_frame, dec, pcm_silk, pcm_ptr, _
                silk_frame_size, decs(st).arch)
            If silk_ret <> 0 Then
                If lost_flag <> 0 Then
                    '/* PLC failure should not be fatal */
                    silk_frame_size = frame_size
                    For i = 0 To frame_size * decs(st).channels - 1
                        pcm_silk(pcm_ptr + i) = 0
                    Next
                Else
                    opus_decode_frame = OPUS_INTERNAL_ERROR
                    Exit Function
                End If
            End If
            pcm_ptr = pcm_ptr + silk_frame_size * decs(st).channels
            decoded_samples = decoded_samples + silk_frame_size
        Loop While decoded_samples < frame_size
    End If

    start_band = 0
    If decode_fec = 0 And mode <> MODE_CELT_ONLY And Not dataIsNull Then
        If mode = MODE_HYBRID Then c = 20 Else c = 0
        If ec_tell(dec) + 17 + c <= 8 * len_ Then
            '/* Check if we have a redundant 0-8 kHz band */
            If mode = MODE_HYBRID Then
                redundancy = ec_dec_bit_logp(dec, 12)
            Else
                redundancy = 1
            End If
            If redundancy <> 0 Then
                celt_to_silk = ec_dec_bit_logp(dec, 1)
                '/* redundancy_bytes will be at least two, in the non-hybrid
                '   case due to the ec_tell() check above */
                If mode = MODE_HYBRID Then
                    redundancy_bytes = ec_dec_uint(dec, 256) + 2
                Else
                    redundancy_bytes = len_ - ((ec_tell(dec) + 7) \ 8)
                End If
                len_ = len_ - redundancy_bytes
                '/* This is a sanity check. It should never happen for a valid
                '   packet, so the exact behaviour is not normative. */
                If len_ * 8 < ec_tell(dec) Then
                    len_ = 0
                    redundancy_bytes = 0
                    redundancy = 0
                End If
                '/* Shrink decoder because of raw bits */
                dec.storage = dec.storage - redundancy_bytes
            End If
        End If
    End If
    If mode <> MODE_CELT_ONLY Then
        start_band = 17
    End If

    If redundancy <> 0 Then
        transition = 0
        pcm_transition_silk_size = 0                    ' ALLOC_NONE
    End If

    If pcm_transition_silk_size > 0 Then ReDim pcm_transition_silk(0 To pcm_transition_silk_size - 1)

    If transition <> 0 And mode <> MODE_CELT_ONLY Then
        pcm_transition = 2
        If F5 < audiosize Then ret = F5 Else ret = audiosize                            ' IMIN(F5, audiosize)
        opus_decode_frame st, data, 0, True, 0, pcm_transition_silk, 0, ret, 0
    End If

    If bandwidth <> 0 Then
        endband = 21

        Select Case bandwidth
        Case OPUS_BANDWIDTH_NARROWBAND
            endband = 13
        Case OPUS_BANDWIDTH_MEDIUMBAND, OPUS_BANDWIDTH_WIDEBAND
            endband = 17
        Case OPUS_BANDWIDTH_SUPERWIDEBAND
            endband = 19
        Case OPUS_BANDWIDTH_FULLBAND
            endband = 21
        Case Else
            ' celt_assert(0);
        End Select
        If celt_decoder_ctl(celt_dec, CELT_SET_END_BAND_REQUEST, endband) <> OPUS_OK Then GoTo must_succeed_failed
    End If
    If celt_decoder_ctl(celt_dec, CELT_SET_CHANNELS_REQUEST, decs(st).stream_channels) <> OPUS_OK Then GoTo must_succeed_failed

    '/* Only allocation memory for redundancy if/when needed */
    If redundancy <> 0 Then redundant_audio_size = F5 * decs(st).channels Else redundant_audio_size = 0
    If redundant_audio_size > 0 Then ReDim redundant_audio(0 To redundant_audio_size - 1)

    '/* 5 ms redundant frame for CELT->SILK*/
    If redundancy <> 0 And celt_to_silk <> 0 Then
        '/* If the previous frame did not use CELT (the first redundancy frame in
        '   a transition from SILK may have been lost) then the CELT decoder is
        '   stale at this point and the redundancy audio is not useful, however
        '   the final range is still needed (for testing), so the redundancy is
        '   always decoded but the decoded audio may not be used */
        If celt_decoder_ctl(celt_dec, CELT_SET_START_BAND_REQUEST, 0) <> OPUS_OK Then GoTo must_succeed_failed
        celt_decode_with_ec celt_dec, data, dataOffset + len_, False, redundancy_bytes, _
                            redundant_audio, 0, F5, noDec, True, 0
        If celt_decoder_ctl(celt_dec, OPUS_GET_FINAL_RANGE_REQUEST, redundant_rng) <> OPUS_OK Then GoTo must_succeed_failed
    End If

    '/* MUST be after PLC */
    If celt_decoder_ctl(celt_dec, CELT_SET_START_BAND_REQUEST, start_band) <> OPUS_OK Then GoTo must_succeed_failed

    If mode <> MODE_SILK_ONLY Then
        If F20 < frame_size Then celt_frame_size = F20 Else celt_frame_size = frame_size  ' IMIN(F20, frame_size)
        '/* Make sure to discard any previous CELT state */
        If mode <> decs(st).prev_mode And decs(st).prev_mode > 0 And decs(st).prev_redundancy = 0 Then
            If celt_decoder_ctl(celt_dec, OPUS_RESET_STATE) <> OPUS_OK Then GoTo must_succeed_failed
        End If
        '/* Decode CELT */
        ' celt_decode_with_ec_dred(celt_dec, decode_fec ? NULL : data, len, pcm, celt_frame_size, &dec, celt_accum)
        celt_ret = celt_decode_with_ec_dred(celt_dec, data, dataOffset, (decode_fec <> 0) Or dataIsNull, _
                                            len_, pcm, pcmOffset, celt_frame_size, dec, False, celt_accum)
    Else
        silence(0) = &HFF: silence(1) = &HFF
        If celt_accum = 0 Then
            For i = 0 To frame_size * decs(st).channels - 1
                pcm(pcmOffset + i) = 0
            Next
        End If
        '/* For hybrid -> SILK transitions, we let the CELT MDCT
        '   do a fade-out by decoding a silence frame */
        If decs(st).prev_mode = MODE_HYBRID And Not (redundancy <> 0 And celt_to_silk <> 0 And decs(st).prev_redundancy <> 0) Then
            If celt_decoder_ctl(celt_dec, CELT_SET_START_BAND_REQUEST, 0) <> OPUS_OK Then GoTo must_succeed_failed
            celt_decode_with_ec celt_dec, silence, 0, False, 2, pcm, pcmOffset, F2_5, noDec, True, celt_accum
        End If
    End If

    If mode <> MODE_CELT_ONLY And celt_accum = 0 Then
        For i = 0 To frame_size * decs(st).channels - 1
            ' pcm[i] = pcm[i] + (opus_val16)((1.f/32768.f)*pcm_silk[i]): the product is exact, the sum rounded once
            pcm(pcmOffset + i) = pcm(pcmOffset + i) + CSng(pcm_silk(i)) * 3.0517578125E-05
        Next
    End If

    If celt_decoder_ctl(celt_dec, CELT_GET_MODE_REQUEST, celt_mode) <> OPUS_OK Then GoTo must_succeed_failed
    celt_vbModeWindow celt_mode, window, overlap        ' window = celt_mode->window;

    '/* 5 ms redundant frame for SILK->CELT */
    If redundancy <> 0 And celt_to_silk = 0 Then
        If celt_decoder_ctl(celt_dec, OPUS_RESET_STATE) <> OPUS_OK Then GoTo must_succeed_failed
        If celt_decoder_ctl(celt_dec, CELT_SET_START_BAND_REQUEST, 0) <> OPUS_OK Then GoTo must_succeed_failed

        celt_decode_with_ec celt_dec, data, dataOffset + len_, False, redundancy_bytes, redundant_audio, 0, F5, noDec, True, 0
        If celt_decoder_ctl(celt_dec, OPUS_GET_FINAL_RANGE_REQUEST, redundant_rng) <> OPUS_OK Then GoTo must_succeed_failed
        smooth_fade pcm, pcmOffset + decs(st).channels * (frame_size - F2_5), redundant_audio, decs(st).channels * F2_5, _
                    pcm, pcmOffset + decs(st).channels * (frame_size - F2_5), F2_5, decs(st).channels, window, decs(st).Fs
    End If
    '/* 5ms redundant frame for CELT->SILK; ignore if the previous frame did not
    '   use CELT (the first redundancy frame in a transition from SILK may have
    '   been lost) */
    If redundancy <> 0 And celt_to_silk <> 0 And (decs(st).prev_mode <> MODE_SILK_ONLY Or decs(st).prev_redundancy <> 0) Then
        For c = 0 To decs(st).channels - 1
            For i = 0 To F2_5 - 1
                pcm(pcmOffset + decs(st).channels * i + c) = redundant_audio(decs(st).channels * i + c)
            Next
        Next
        smooth_fade redundant_audio, decs(st).channels * F2_5, pcm, pcmOffset + decs(st).channels * F2_5, _
                    pcm, pcmOffset + decs(st).channels * F2_5, F2_5, decs(st).channels, window, decs(st).Fs
    End If
    If transition <> 0 Then
        If audiosize >= F5 Then
            If pcm_transition = 1 Then
                For i = 0 To decs(st).channels * F2_5 - 1
                    pcm(pcmOffset + i) = pcm_transition_celt(i)
                Next
                smooth_fade pcm_transition_celt, decs(st).channels * F2_5, pcm, pcmOffset + decs(st).channels * F2_5, _
                            pcm, pcmOffset + decs(st).channels * F2_5, F2_5, decs(st).channels, window, decs(st).Fs
            Else
                For i = 0 To decs(st).channels * F2_5 - 1
                    pcm(pcmOffset + i) = pcm_transition_silk(i)
                Next
                smooth_fade pcm_transition_silk, decs(st).channels * F2_5, pcm, pcmOffset + decs(st).channels * F2_5, _
                            pcm, pcmOffset + decs(st).channels * F2_5, F2_5, decs(st).channels, window, decs(st).Fs
            End If
        Else
            '/* Not enough time to do a clean transition, but we do it anyway
            '   This will not preserve amplitude perfectly and may introduce
            '   a bit of temporal aliasing, but it shouldn't be too bad and
            '   that's pretty much the best we can do. In any case, generating this
            '   transition it pretty silly in the first place */
            If pcm_transition = 1 Then
                smooth_fade pcm_transition_celt, 0, pcm, pcmOffset, pcm, pcmOffset, F2_5, decs(st).channels, window, decs(st).Fs
            Else
                smooth_fade pcm_transition_silk, 0, pcm, pcmOffset, pcm, pcmOffset, F2_5, decs(st).channels, window, decs(st).Fs
            End If
        End If
    End If

    If decs(st).decode_gain <> 0 Then
        gain = celt_exp2(fmul(6.48814081E-04!, CSng(decs(st).decode_gain)))   ' MULT16_16_P15(QCONST16(6.48814081e-4f, 25), st->decode_gain)
        For i = 0 To frame_size * decs(st).channels - 1
            pcm(pcmOffset + i) = fmul(pcm(pcmOffset + i), gain)        ' x = MULT16_32_P16(pcm[i],gain); pcm[i] = SATURATE(x, 32767)
        Next
    End If

    If len_ <= 1 Then
        decs(st).rangeFinal = 0
    Else
        decs(st).rangeFinal = dec.rng Xor redundant_rng
    End If

    decs(st).prev_mode = mode
    If redundancy <> 0 And celt_to_silk = 0 Then decs(st).prev_redundancy = 1 Else decs(st).prev_redundancy = 0

    ' if (celt_ret>=0) { if (OPUS_CHECK_ARRAY(pcm, audiosize*st->channels)) OPUS_PRINT_INT(audiosize); } (FUZZING only)

    If celt_ret < 0 Then opus_decode_frame = celt_ret Else opus_decode_frame = audiosize
    Exit Function

must_succeed_failed:
    ' MUST_SUCCEED(call): do {if((call) != OPUS_OK) {RESTORE_STACK; return OPUS_INTERNAL_ERROR;} } while (0)
    opus_decode_frame = OPUS_INTERNAL_ERROR
End Function

Public Function opus_decode_native(ByVal st As Long, data() As Byte, ByVal dataOffset As Long, ByVal dataIsNull As Boolean, _
        ByVal len_ As Long, pcm() As Single, ByVal pcmOffset As Long, ByVal frame_size As Long, ByVal decode_fec As Long, _
        ByVal self_delimited As Long, packet_offset As Long, ByVal soft_clip As Long) As Long
    Dim i As Long, nb_samples As Long
    Dim count As Long, offset As Long
    Dim toc As Long
    Dim packet_frame_size As Long, packet_bandwidth As Long, packet_mode As Long, packet_stream_channels As Long
    '/* 48 x 2.5 ms = 120 ms */
    Dim size(0 To 47) As Integer
    Dim frames(0 To 47) As Long                         ' (VB) the frames NULL
    Dim padding As Long, padding_len As Long            ' (VB) NULL
    Dim pcm_count As Long
    Dim ret As Long
    Dim duration_copy As Long
    Dim dummy As Long

    ' VALIDATE_OPUS_DECODER(st);
    If decode_fec < 0 Or decode_fec > 1 Then
        opus_decode_native = OPUS_BAD_ARG
        Exit Function
    End If
    '/* For FEC/PLC, frame_size has to be to have a multiple of 2.5 ms */
    If decode_fec <> 0 Or len_ = 0 Or dataIsNull Then
        If frame_size Mod (decs(st).Fs \ 400) <> 0 Then
            opus_decode_native = OPUS_BAD_ARG
            Exit Function
        End If
    End If
    If len_ = 0 Or dataIsNull Then
        pcm_count = 0
        Do
            ret = opus_decode_frame(st, data, 0, True, 0, pcm, pcmOffset + pcm_count * decs(st).channels, frame_size - pcm_count, 0)
            If ret < 0 Then
                opus_decode_native = ret
                Exit Function
            End If
            pcm_count = pcm_count + ret
        Loop While pcm_count < frame_size
        ' celt_assert(pcm_count == frame_size);
        decs(st).last_packet_duration = pcm_count
        opus_decode_native = pcm_count
        Exit Function
    ElseIf len_ < 0 Then
        opus_decode_native = OPUS_BAD_ARG
        Exit Function
    End If

    packet_mode = opus_packet_get_mode(data, dataOffset)
    packet_bandwidth = opus_packet_get_bandwidth(data, dataOffset)
    packet_frame_size = opus_packet_get_samples_per_frame(data, dataOffset, decs(st).Fs)
    packet_stream_channels = opus_packet_get_nb_channels(data, dataOffset)

    count = opus_packet_parse_impl(data, dataOffset, len_, self_delimited, toc, frames, _
                                   size, offset, packet_offset, padding, padding_len)
    If count < 0 Then
        opus_decode_native = count
        Exit Function
    End If

    dataOffset = dataOffset + offset

    If decode_fec <> 0 Then
        '/* If no FEC can be present, run the PLC (recursive call) */
        If frame_size < packet_frame_size Or packet_mode = MODE_CELT_ONLY Or decs(st).mode = MODE_CELT_ONLY Then
            opus_decode_native = opus_decode_native(st, data, 0, True, 0, pcm, pcmOffset, frame_size, 0, 0, dummy, soft_clip)
            Exit Function
        End If
        '/* Otherwise, run the PLC on everything except the size for which we might have FEC */
        duration_copy = decs(st).last_packet_duration
        If frame_size - packet_frame_size <> 0 Then
            ret = opus_decode_native(st, data, 0, True, 0, pcm, pcmOffset, frame_size - packet_frame_size, 0, 0, dummy, soft_clip)
            If ret < 0 Then
                decs(st).last_packet_duration = duration_copy
                opus_decode_native = ret
                Exit Function
            End If
            ' celt_assert(ret==frame_size-packet_frame_size);
        End If
        '/* Complete with FEC */
        decs(st).mode = packet_mode
        decs(st).bandwidth = packet_bandwidth
        decs(st).frame_size = packet_frame_size
        decs(st).stream_channels = packet_stream_channels
        ret = opus_decode_frame(st, data, dataOffset, False, size(0), pcm, pcmOffset + decs(st).channels * (frame_size - packet_frame_size), _
              packet_frame_size, 1)
        If ret < 0 Then
            opus_decode_native = ret
            Exit Function
        Else
            decs(st).last_packet_duration = frame_size
            opus_decode_native = frame_size
            Exit Function
        End If
    End If

    If count * packet_frame_size > frame_size Then
        opus_decode_native = OPUS_BUFFER_TOO_SMALL
        Exit Function
    End If

    '/* Update the state as the last step to avoid updating it on an invalid packet */
    decs(st).mode = packet_mode
    decs(st).bandwidth = packet_bandwidth
    decs(st).frame_size = packet_frame_size
    decs(st).stream_channels = packet_stream_channels

    nb_samples = 0
    For i = 0 To count - 1
        ret = opus_decode_frame(st, data, dataOffset, False, size(i), pcm, pcmOffset + nb_samples * decs(st).channels, frame_size - nb_samples, 0)
        If ret < 0 Then
            opus_decode_native = ret
            Exit Function
        End If
        ' celt_assert(ret==packet_frame_size);
        dataOffset = dataOffset + size(i)
        nb_samples = nb_samples + ret
    Next
    decs(st).last_packet_duration = nb_samples
    If soft_clip <> 0 Then
        opus_pcm_soft_clip pcm, pcmOffset, nb_samples, decs(st).channels, decs(st).softclip_mem, 0
    Else
        decs(st).softclip_mem(0) = 0: decs(st).softclip_mem(1) = 0
    End If
    opus_decode_native = nb_samples
End Function

' (VB) float_cast.h's FLOAT2INT16 (float2int is lrintf: VB's CLng rounds half to even, as lrintf does)
Private Function FLOAT2INT16(ByVal x As Single) As Integer
    x = x * 32768!                                      ' x = x*CELT_SIG_SCALE (exact)
    If x < -32768! Then x = -32768!                     ' MAX32(x, -32768)
    If x > 32767! Then x = 32767!                       ' MIN32(x, 32767)
    FLOAT2INT16 = CLng(x)
End Function

Public Function opus_decode(ByVal st As Long, data() As Byte, ByVal dataOffset As Long, ByVal len_ As Long, _
        pcm() As Integer, ByVal pcmOffset As Long, ByVal frame_size As Long, ByVal decode_fec As Long) As Long
    Dim out_() As Single
    Dim ret As Long, i As Long
    Dim nb_samples As Long
    Dim dummy As Long

    If frame_size <= 0 Then
        opus_decode = OPUS_BAD_ARG
        Exit Function
    End If

    If len_ > 0 And decode_fec = 0 Then                 ' data != NULL && len > 0 && !decode_fec
        nb_samples = opus_decoder_get_nb_samples(st, data, dataOffset, len_)
        If nb_samples > 0 Then
            If nb_samples < frame_size Then frame_size = nb_samples                     ' IMIN(frame_size, nb_samples)
        Else
            opus_decode = OPUS_INVALID_PACKET
            Exit Function
        End If
    End If
    ' celt_assert(st->channels == 1 || st->channels == 2);
    ReDim out_(0 To frame_size * decs(st).channels - 1)

    ret = opus_decode_native(st, data, dataOffset, (len_ = 0), len_, out_, 0, frame_size, decode_fec, 0, dummy, 1)
    If ret > 0 Then
        For i = 0 To ret * decs(st).channels - 1
            pcm(pcmOffset + i) = FLOAT2INT16(out_(i))
        Next
    End If
    opus_decode = ret
End Function

Public Function opus_decode_float(ByVal st As Long, data() As Byte, ByVal dataOffset As Long, ByVal len_ As Long, _
        pcm() As Single, ByVal pcmOffset As Long, ByVal frame_size As Long, ByVal decode_fec As Long) As Long
    Dim dummy As Long
    If frame_size <= 0 Then
        opus_decode_float = OPUS_BAD_ARG
        Exit Function
    End If
    opus_decode_float = opus_decode_native(st, data, dataOffset, (len_ = 0), len_, pcm, pcmOffset, frame_size, decode_fec, 0, dummy, 0)
End Function

' (VB) the request's argument: value (int32 or opus_uint32 *, a GET writes it). The C's NULL checks of a GET's
' pointer (goto bad_arg) have no counterpart: an omitted Optional is a 0 written to nothing.
Public Function opus_decoder_ctl(ByVal st As Long, ByVal request As Long, Optional ByRef value As Long) As Long
    Dim ret As Long
    Dim celt_dec As Long

    ret = OPUS_OK
    celt_dec = decs(st).celt_dec

    Select Case request
    Case OPUS_GET_BANDWIDTH_REQUEST
        value = decs(st).bandwidth
    Case OPUS_SET_COMPLEXITY_REQUEST
        If value < 0 Or value > 10 Then
            GoTo bad_arg
        End If
        decs(st).complexity = value
        celt_decoder_ctl celt_dec, OPUS_SET_COMPLEXITY_REQUEST, value
    Case OPUS_GET_COMPLEXITY_REQUEST
        value = decs(st).complexity
    Case OPUS_GET_FINAL_RANGE_REQUEST
        value = decs(st).rangeFinal
    Case OPUS_RESET_STATE
        ' OPUS_CLEAR((char*)&st->OPUS_DECODER_RESET_START, sizeof(OpusDecoder)- ((char*)&st->OPUS_DECODER_RESET_START - (char*)st));
        decs(st).stream_channels = 0
        decs(st).bandwidth = 0
        decs(st).mode = 0
        decs(st).prev_mode = 0
        decs(st).frame_size = 0
        decs(st).prev_redundancy = 0
        decs(st).last_packet_duration = 0
        decs(st).softclip_mem(0) = 0: decs(st).softclip_mem(1) = 0
        decs(st).rangeFinal = 0

        celt_decoder_ctl celt_dec, OPUS_RESET_STATE
        silk_ResetDecoder decs(st).silk_dec
        decs(st).stream_channels = decs(st).channels
        decs(st).frame_size = decs(st).Fs \ 400
    Case OPUS_GET_SAMPLE_RATE_REQUEST
        value = decs(st).Fs
    Case OPUS_GET_PITCH_REQUEST
        If decs(st).prev_mode = MODE_CELT_ONLY Then
            ret = celt_decoder_ctl(celt_dec, OPUS_GET_PITCH_REQUEST, value)
        Else
            value = decs(st).DecControl.prevPitchLag
        End If
    Case OPUS_GET_GAIN_REQUEST
        value = decs(st).decode_gain
    Case OPUS_SET_GAIN_REQUEST
        If value < -32768 Or value > 32767 Then
            GoTo bad_arg
        End If
        decs(st).decode_gain = value
    Case OPUS_GET_LAST_PACKET_DURATION_REQUEST
        value = decs(st).last_packet_duration
    Case OPUS_SET_PHASE_INVERSION_DISABLED_REQUEST
        If value < 0 Or value > 1 Then
            GoTo bad_arg
        End If
        ret = celt_decoder_ctl(celt_dec, OPUS_SET_PHASE_INVERSION_DISABLED_REQUEST, value)
    Case OPUS_GET_PHASE_INVERSION_DISABLED_REQUEST
        ret = celt_decoder_ctl(celt_dec, OPUS_GET_PHASE_INVERSION_DISABLED_REQUEST, value)
    Case Else
        '/*fprintf(stderr, "unknown opus_decoder_ctl() request: %d", request);*/
        ret = OPUS_UNIMPLEMENTED
    End Select

    opus_decoder_ctl = ret
    Exit Function
bad_arg:
    opus_decoder_ctl = OPUS_BAD_ARG
End Function

Public Sub opus_decoder_destroy(ByVal st As Long)
    opus_decoder_vbFree st
End Sub

Public Function opus_packet_get_bandwidth(data() As Byte, ByVal dataOffset As Long) As Long
    Dim bandwidth As Long
    If (data(dataOffset) And &H80) <> 0 Then
        bandwidth = OPUS_BANDWIDTH_MEDIUMBAND + ((data(dataOffset) \ 32) And &H3)
        If bandwidth = OPUS_BANDWIDTH_MEDIUMBAND Then
            bandwidth = OPUS_BANDWIDTH_NARROWBAND
        End If
    ElseIf (data(dataOffset) And &H60) = &H60 Then
        If (data(dataOffset) And &H10) <> 0 Then bandwidth = OPUS_BANDWIDTH_FULLBAND Else bandwidth = OPUS_BANDWIDTH_SUPERWIDEBAND
    Else
        bandwidth = OPUS_BANDWIDTH_NARROWBAND + ((data(dataOffset) \ 32) And &H3)
    End If
    opus_packet_get_bandwidth = bandwidth
End Function

Public Function opus_packet_get_nb_channels(data() As Byte, ByVal dataOffset As Long) As Long
    If (data(dataOffset) And &H4) <> 0 Then opus_packet_get_nb_channels = 2 Else opus_packet_get_nb_channels = 1
End Function

Public Function opus_packet_get_nb_frames(packet() As Byte, ByVal packetOffset As Long, ByVal len_ As Long) As Long
    Dim count As Long
    If len_ < 1 Then
        opus_packet_get_nb_frames = OPUS_BAD_ARG
        Exit Function
    End If
    count = packet(packetOffset) And &H3
    If count = 0 Then
        opus_packet_get_nb_frames = 1
    ElseIf count <> 3 Then
        opus_packet_get_nb_frames = 2
    ElseIf len_ < 2 Then
        opus_packet_get_nb_frames = OPUS_INVALID_PACKET
    Else
        opus_packet_get_nb_frames = packet(packetOffset + 1) And &H3F
    End If
End Function

Public Function opus_packet_get_nb_samples(packet() As Byte, ByVal packetOffset As Long, ByVal len_ As Long, ByVal Fs As Long) As Long
    Dim samples As Long
    Dim count As Long
    count = opus_packet_get_nb_frames(packet, packetOffset, len_)

    If count < 0 Then
        opus_packet_get_nb_samples = count
        Exit Function
    End If

    samples = count * opus_packet_get_samples_per_frame(packet, packetOffset, Fs)
    '/* Can't have more than 120 ms */
    If samples * 25 > Fs * 3 Then
        opus_packet_get_nb_samples = OPUS_INVALID_PACKET
    Else
        opus_packet_get_nb_samples = samples
    End If
End Function

Public Function opus_packet_has_lbrr(packet() As Byte, ByVal packetOffset As Long, ByVal len_ As Long) As Long
    Dim ret As Long
    Dim frames(0 To 47) As Long
    Dim size(0 To 47) As Integer
    Dim packet_mode As Long, packet_frame_size As Long, packet_stream_channels As Long
    Dim nb_frames As Long
    Dim lbrr As Long
    Dim toc As Long, payload_offset As Long

    nb_frames = 1
    packet_mode = opus_packet_get_mode(packet, packetOffset)
    If packet_mode = MODE_CELT_ONLY Then
        Exit Function
    End If
    packet_frame_size = opus_packet_get_samples_per_frame(packet, packetOffset, 48000)
    If packet_frame_size > 960 Then
        nb_frames = packet_frame_size \ 960
    End If
    packet_stream_channels = opus_packet_get_nb_channels(packet, packetOffset)
    ret = opus_packet_parse(packet, packetOffset, len_, toc, frames, size, payload_offset)
    If ret <= 0 Then
        opus_packet_has_lbrr = ret
        Exit Function
    End If
    ' (VB) a byte shifted right without silk_RSHIFT (its table is filled by a decoder's init; this runs without one)
    lbrr = (packet(frames(0)) \ (2 ^ (7 - nb_frames))) And &H1
    If packet_stream_channels = 2 Then
        If lbrr = 0 Then
            If ((packet(frames(0)) \ (2 ^ (6 - 2 * nb_frames))) And &H1) <> 0 Then lbrr = 1
        End If
    End If
    opus_packet_has_lbrr = lbrr
End Function

Public Function opus_decoder_get_nb_samples(ByVal dec As Long, packet() As Byte, ByVal packetOffset As Long, ByVal len_ As Long) As Long
    opus_decoder_get_nb_samples = opus_packet_get_nb_samples(packet, packetOffset, len_, decs(dec).Fs)
End Function

' (VB) for opus_multistream_decoder_c: the channels of a decoder (its st->channels)
Public Function opus_decoder_vbChannels(ByVal st As Long) As Long
    opus_decoder_vbChannels = decs(st).channels
End Function
