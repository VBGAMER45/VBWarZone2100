Attribute VB_Name = "opus_multistream_decoder_c"
' libopus 1.5.2 (ref/opus/src): opus_multistream_decoder.c, the float build (opus_multistream_decode soft-clips
' and converts with FLOAT2INT16), no hardening (VALIDATE_MS_DECODER is empty).
' (VB) An OpusMSDecoder * is a handle: an index of this module's pool (0 = NULL). The C lays the streams'
' decoders out after the struct (the ptr arithmetic with align(coupled_size) and align(mono_size)): here the
' struct holds their OpusDecoder handles, made by opus_multistream_decoder_init. The opus_copy_channel_out_func
' function pointer is a number (COPY_CHANNEL_OUT_FLOAT, COPY_CHANNEL_OUT_SHORT) and the void *pcm is the two
' arrays (only the one it names is used). A packet is an array and an offset (len == 0: data NULL).
Option Explicit

Public Type OpusMSDecoder
    layout As ChannelLayout
    '/* Decoder states go here */
    dec() As Long                                       ' (VB) the streams' OpusDecoder handles, 0 To nb_streams - 1
    vbInUse As Boolean                                  ' (VB) the pool slot is taken
End Type

' opus_copy_channel_out_func (VB: its number)
Private Const COPY_CHANNEL_OUT_FLOAT As Long = 1        ' opus_copy_channel_out_float
Private Const COPY_CHANNEL_OUT_SHORT As Long = 2        ' opus_copy_channel_out_short

Private msdecs() As OpusMSDecoder
Private msdecs_count As Long
Private buf() As Single                                 ' opus_multistream_decode_native's stack buffer
Private buf_size As Long

'/* DECODER */

Public Function opus_multistream_decoder_get_size(ByVal nb_streams As Long, ByVal nb_coupled_streams As Long) As Long
    Dim coupled_size As Long
    Dim mono_size As Long
    Dim d As OpusMSDecoder

    If nb_streams < 1 Or nb_coupled_streams > nb_streams Or nb_coupled_streams < 0 Then Exit Function
    coupled_size = opus_decoder_get_size(2)
    mono_size = opus_decoder_get_size(1)
    opus_multistream_decoder_get_size = align(LenB(d.layout)) _
         + nb_coupled_streams * align(coupled_size) _
         + (nb_streams - nb_coupled_streams) * align(mono_size)
End Function

' (VB) the streams' decoders of a pool slot, freed
Private Sub opus_multistream_vbFreeStreams(ByVal st As Long)
    Dim s As Long
    If msdecs(st).layout.nb_streams > 0 Then
        For s = 0 To UBound(msdecs(st).dec)
            If msdecs(st).dec(s) <> 0 Then opus_decoder_destroy msdecs(st).dec(s)
            msdecs(st).dec(s) = 0
        Next
    End If
End Sub

Public Function opus_multistream_decoder_init(ByVal st As Long, ByVal Fs As Long, ByVal channels As Long, ByVal streams As Long, _
        ByVal coupled_streams As Long, mapping() As Byte, ByVal mappingOffset As Long) As Long
    Dim i As Long, ret As Long

    If (channels > 255) Or (channels < 1) Or (coupled_streams > streams) Or _
        (streams < 1) Or (coupled_streams < 0) Or (streams > 255 - coupled_streams) Then
        opus_multistream_decoder_init = OPUS_BAD_ARG
        Exit Function
    End If

    opus_multistream_vbFreeStreams st
    msdecs(st).layout.nb_channels = channels
    msdecs(st).layout.nb_streams = streams
    msdecs(st).layout.nb_coupled_streams = coupled_streams
    ReDim msdecs(st).dec(0 To streams - 1)

    For i = 0 To msdecs(st).layout.nb_channels - 1
        msdecs(st).layout.mapping(i) = mapping(mappingOffset + i)
    Next
    If validate_layout(msdecs(st).layout) = 0 Then
        opus_multistream_decoder_init = OPUS_BAD_ARG
        Exit Function
    End If

    ' (VB) opus_decoder_init on the states after the struct: a decoder made for each stream
    For i = 0 To msdecs(st).layout.nb_coupled_streams - 1
        msdecs(st).dec(i) = opus_decoder_create(Fs, 2, ret)
        If ret <> OPUS_OK Then
            opus_multistream_decoder_init = ret
            Exit Function
        End If
    Next
    For i = msdecs(st).layout.nb_coupled_streams To msdecs(st).layout.nb_streams - 1
        msdecs(st).dec(i) = opus_decoder_create(Fs, 1, ret)
        If ret <> OPUS_OK Then
            opus_multistream_decoder_init = ret
            Exit Function
        End If
    Next
    opus_multistream_decoder_init = OPUS_OK
End Function

Public Function opus_multistream_decoder_create(ByVal Fs As Long, ByVal channels As Long, ByVal streams As Long, _
        ByVal coupled_streams As Long, mapping() As Byte, ByVal mappingOffset As Long, error_ As Long) As Long
    Dim ret As Long
    Dim st As Long
    Static blank As OpusMSDecoder

    If (channels > 255) Or (channels < 1) Or (coupled_streams > streams) Or _
        (streams < 1) Or (coupled_streams < 0) Or (streams > 255 - coupled_streams) Then
        error_ = OPUS_BAD_ARG
        Exit Function
    End If
    ' st = (OpusMSDecoder *)opus_alloc(opus_multistream_decoder_get_size(streams, coupled_streams));
    For st = 1 To msdecs_count
        If Not msdecs(st).vbInUse Then Exit For
    Next
    If st > msdecs_count Then
        msdecs_count = msdecs_count + 4
        ReDim Preserve msdecs(1 To msdecs_count)
    End If
    msdecs(st) = blank
    msdecs(st).vbInUse = True
    ret = opus_multistream_decoder_init(st, Fs, channels, streams, coupled_streams, mapping, mappingOffset)
    error_ = ret
    If ret <> OPUS_OK Then
        opus_multistream_decoder_destroy st
        st = 0
    End If
    opus_multistream_decoder_create = st
End Function

Private Function opus_multistream_packet_validate(data() As Byte, ByVal dataOffset As Long, ByVal len_ As Long, _
        ByVal nb_streams As Long, ByVal Fs As Long) As Long
    Dim s As Long
    Dim count As Long
    Dim toc As Long
    Dim size(0 To 47) As Integer
    Dim samples As Long
    Dim packet_offset As Long
    Dim tmp_samples As Long
    Dim frames(0 To 47) As Long, payload_offset As Long, padding As Long, padding_len As Long   ' (VB) NULL

    For s = 0 To nb_streams - 1
        If len_ <= 0 Then
            opus_multistream_packet_validate = OPUS_INVALID_PACKET
            Exit Function
        End If
        count = opus_packet_parse_impl(data, dataOffset, len_, IIf(s <> nb_streams - 1, 1, 0), toc, frames, _
                                       size, payload_offset, packet_offset, padding, padding_len)
        If count < 0 Then
            opus_multistream_packet_validate = count
            Exit Function
        End If
        tmp_samples = opus_packet_get_nb_samples(data, dataOffset, packet_offset, Fs)
        If s <> 0 And samples <> tmp_samples Then
            opus_multistream_packet_validate = OPUS_INVALID_PACKET
            Exit Function
        End If
        samples = tmp_samples
        dataOffset = dataOffset + packet_offset
        len_ = len_ - packet_offset
    Next
    opus_multistream_packet_validate = samples
End Function

Public Function opus_multistream_decode_native(ByVal st As Long, data() As Byte, ByVal dataOffset As Long, ByVal len_ As Long, _
        pcmF() As Single, pcmS() As Integer, ByVal pcmOffset As Long, ByVal copy_channel_out As Long, _
        ByVal frame_size As Long, ByVal decode_fec As Long, ByVal soft_clip As Long) As Long
    Dim Fs As Long
    Dim s As Long, c As Long
    Dim do_plc As Long
    Dim dec As Long                                     ' OpusDecoder *
    Dim packet_offset As Long
    Dim ret As Long
    Dim chan As Long, prev As Long

    ' VALIDATE_MS_DECODER(st);
    If frame_size <= 0 Then
        opus_multistream_decode_native = OPUS_BAD_ARG
        Exit Function
    End If
    '/* Limit frame_size to avoid excessive stack allocations. */
    If opus_multistream_decoder_ctl(st, OPUS_GET_SAMPLE_RATE_REQUEST, Fs) <> OPUS_OK Then      ' MUST_SUCCEED
        opus_multistream_decode_native = OPUS_INTERNAL_ERROR
        Exit Function
    End If
    If (Fs \ 25) * 3 < frame_size Then frame_size = (Fs \ 25) * 3                              ' IMIN(frame_size, Fs/25*3)
    ' ALLOC(buf, 2*frame_size, opus_val16);
    If buf_size < 2 * frame_size Then
        buf_size = 2 * frame_size
        ReDim buf(0 To buf_size - 1)
    End If

    If len_ = 0 Then
        do_plc = 1
    End If
    If len_ < 0 Then
        opus_multistream_decode_native = OPUS_BAD_ARG
        Exit Function
    End If
    If do_plc = 0 And len_ < 2 * msdecs(st).layout.nb_streams - 1 Then
        opus_multistream_decode_native = OPUS_INVALID_PACKET
        Exit Function
    End If
    If do_plc = 0 Then
        ret = opus_multistream_packet_validate(data, dataOffset, len_, msdecs(st).layout.nb_streams, Fs)
        If ret < 0 Then
            opus_multistream_decode_native = ret
            Exit Function
        ElseIf ret > frame_size Then
            opus_multistream_decode_native = OPUS_BUFFER_TOO_SMALL
            Exit Function
        End If
    End If
    For s = 0 To msdecs(st).layout.nb_streams - 1
        dec = msdecs(st).dec(s)

        If do_plc = 0 And len_ <= 0 Then
            opus_multistream_decode_native = OPUS_INTERNAL_ERROR
            Exit Function
        End If
        packet_offset = 0
        ret = opus_decode_native(dec, data, dataOffset, (do_plc <> 0), len_, buf, 0, frame_size, decode_fec, _
            IIf(s <> msdecs(st).layout.nb_streams - 1, 1, 0), packet_offset, soft_clip)
        If do_plc = 0 Then
            dataOffset = dataOffset + packet_offset
            len_ = len_ - packet_offset
        End If
        If ret <= 0 Then
            opus_multistream_decode_native = ret
            Exit Function
        End If
        frame_size = ret
        If s < msdecs(st).layout.nb_coupled_streams Then
            prev = -1
            '/* Copy "left" audio to the channel(s) where it belongs */
            chan = get_left_channel(msdecs(st).layout, s, prev)
            Do While chan <> -1
                copy_channel_out_ copy_channel_out, pcmF, pcmS, pcmOffset, msdecs(st).layout.nb_channels, chan, _
                    False, 0, 2, frame_size
                prev = chan
                chan = get_left_channel(msdecs(st).layout, s, prev)
            Loop
            prev = -1
            '/* Copy "right" audio to the channel(s) where it belongs */
            chan = get_right_channel(msdecs(st).layout, s, prev)
            Do While chan <> -1
                copy_channel_out_ copy_channel_out, pcmF, pcmS, pcmOffset, msdecs(st).layout.nb_channels, chan, _
                    False, 1, 2, frame_size
                prev = chan
                chan = get_right_channel(msdecs(st).layout, s, prev)
            Loop
        Else
            prev = -1
            '/* Copy audio to the channel(s) where it belongs */
            chan = get_mono_channel(msdecs(st).layout, s, prev)
            Do While chan <> -1
                copy_channel_out_ copy_channel_out, pcmF, pcmS, pcmOffset, msdecs(st).layout.nb_channels, chan, _
                    False, 0, 1, frame_size
                prev = chan
                chan = get_mono_channel(msdecs(st).layout, s, prev)
            Loop
        End If
    Next
    '/* Handle muted channels */
    For c = 0 To msdecs(st).layout.nb_channels - 1
        If msdecs(st).layout.mapping(c) = 255 Then
            copy_channel_out_ copy_channel_out, pcmF, pcmS, pcmOffset, msdecs(st).layout.nb_channels, c, _
                True, 0, 0, frame_size
        End If
    Next
    opus_multistream_decode_native = frame_size
End Function

' (VB) (*copy_channel_out)(dst, dst_stride, dst_channel, src, src_stride, frame_size, user_data): src is buf from
' srcOffset (NULL: srcIsNull)
Private Sub copy_channel_out_(ByVal copy_channel_out As Long, pcmF() As Single, pcmS() As Integer, ByVal pcmOffset As Long, _
        ByVal dst_stride As Long, ByVal dst_channel As Long, ByVal srcIsNull As Boolean, ByVal srcOffset As Long, _
        ByVal src_stride As Long, ByVal frame_size As Long)
    If copy_channel_out = COPY_CHANNEL_OUT_FLOAT Then
        opus_copy_channel_out_float pcmF, pcmOffset, dst_stride, dst_channel, srcIsNull, srcOffset, src_stride, frame_size
    Else
        opus_copy_channel_out_short pcmS, pcmOffset, dst_stride, dst_channel, srcIsNull, srcOffset, src_stride, frame_size
    End If
End Sub

Private Sub opus_copy_channel_out_float(float_dst() As Single, ByVal dstOffset As Long, ByVal dst_stride As Long, _
        ByVal dst_channel As Long, ByVal srcIsNull As Boolean, ByVal srcOffset As Long, ByVal src_stride As Long, _
        ByVal frame_size As Long)
    Dim i As Long
    If Not srcIsNull Then
        For i = 0 To frame_size - 1
            float_dst(dstOffset + i * dst_stride + dst_channel) = buf(srcOffset + i * src_stride)
        Next
    Else
        For i = 0 To frame_size - 1
            float_dst(dstOffset + i * dst_stride + dst_channel) = 0
        Next
    End If
End Sub

' (VB) float_cast.h's FLOAT2INT16 (float2int is lrintf: VB's CLng rounds half to even, as lrintf does)
Private Function FLOAT2INT16(ByVal x As Single) As Integer
    x = x * 32768!                                      ' x = x*CELT_SIG_SCALE (exact)
    If x < -32768! Then x = -32768!                     ' MAX32(x, -32768)
    If x > 32767! Then x = 32767!                       ' MIN32(x, 32767)
    FLOAT2INT16 = CLng(x)
End Function

Private Sub opus_copy_channel_out_short(short_dst() As Integer, ByVal dstOffset As Long, ByVal dst_stride As Long, _
        ByVal dst_channel As Long, ByVal srcIsNull As Boolean, ByVal srcOffset As Long, ByVal src_stride As Long, _
        ByVal frame_size As Long)
    Dim i As Long
    If Not srcIsNull Then
        For i = 0 To frame_size - 1
            short_dst(dstOffset + i * dst_stride + dst_channel) = FLOAT2INT16(buf(srcOffset + i * src_stride))
        Next
    Else
        For i = 0 To frame_size - 1
            short_dst(dstOffset + i * dst_stride + dst_channel) = 0
        Next
    End If
End Sub

Public Function opus_multistream_decode(ByVal st As Long, data() As Byte, ByVal dataOffset As Long, ByVal len_ As Long, _
        pcm() As Integer, ByVal pcmOffset As Long, ByVal frame_size As Long, ByVal decode_fec As Long) As Long
    Dim noFloat() As Single
    opus_multistream_decode = opus_multistream_decode_native(st, data, dataOffset, len_, _
        noFloat, pcm, pcmOffset, COPY_CHANNEL_OUT_SHORT, frame_size, decode_fec, 1)
End Function

Public Function opus_multistream_decode_float(ByVal st As Long, data() As Byte, ByVal dataOffset As Long, _
        ByVal len_ As Long, pcm() As Single, ByVal pcmOffset As Long, ByVal frame_size As Long, ByVal decode_fec As Long) As Long
    Dim noShort() As Integer
    opus_multistream_decode_float = opus_multistream_decode_native(st, data, dataOffset, len_, _
        pcm, noShort, pcmOffset, COPY_CHANNEL_OUT_FLOAT, frame_size, decode_fec, 0)
End Function

' (VB) va_list: value is the request's argument (int32 or a GET's pointer); OPUS_MULTISTREAM_GET_DECODER_STATE
' takes two: value is the stream_id, value2 gets the OpusDecoder handle. The C's NULL pointer checks of a GET
' (goto bad_arg) have no counterpart.
Public Function opus_multistream_decoder_ctl_va_list(ByVal st As Long, ByVal request As Long, value As Long, value2 As Long) As Long
    Dim ret As Long
    Dim s As Long
    Dim tmp As Long
    Dim stream_id As Long

    ret = OPUS_OK
    Select Case request
        Case OPUS_GET_BANDWIDTH_REQUEST, OPUS_GET_SAMPLE_RATE_REQUEST, OPUS_GET_GAIN_REQUEST, _
             OPUS_GET_LAST_PACKET_DURATION_REQUEST, OPUS_GET_PHASE_INVERSION_DISABLED_REQUEST
            '/* For int32* GET params, just query the first stream */
            ret = opus_decoder_ctl(msdecs(st).dec(0), request, value)
        Case OPUS_GET_FINAL_RANGE_REQUEST
            value = 0
            For s = 0 To msdecs(st).layout.nb_streams - 1
                ret = opus_decoder_ctl(msdecs(st).dec(s), request, tmp)
                If ret <> OPUS_OK Then Exit For
                value = value Xor tmp
            Next
        Case OPUS_RESET_STATE
            For s = 0 To msdecs(st).layout.nb_streams - 1
                ret = opus_decoder_ctl(msdecs(st).dec(s), OPUS_RESET_STATE)
                If ret <> OPUS_OK Then
                    Exit For
                End If
            Next
        Case OPUS_MULTISTREAM_GET_DECODER_STATE_REQUEST
            stream_id = value
            If stream_id < 0 Or stream_id >= msdecs(st).layout.nb_streams Then
                GoTo bad_arg
            End If
            value2 = msdecs(st).dec(stream_id)
        Case OPUS_SET_GAIN_REQUEST, OPUS_SET_PHASE_INVERSION_DISABLED_REQUEST
            '/* This works for int32 params */
            For s = 0 To msdecs(st).layout.nb_streams - 1
                ret = opus_decoder_ctl(msdecs(st).dec(s), request, value)
                If ret <> OPUS_OK Then
                    Exit For
                End If
            Next
        Case Else
            ret = OPUS_UNIMPLEMENTED
    End Select
    opus_multistream_decoder_ctl_va_list = ret
    Exit Function
bad_arg:
    opus_multistream_decoder_ctl_va_list = OPUS_BAD_ARG
End Function

Public Function opus_multistream_decoder_ctl(ByVal st As Long, ByVal request As Long, Optional ByRef value As Long, _
        Optional ByRef value2 As Long) As Long
    opus_multistream_decoder_ctl = opus_multistream_decoder_ctl_va_list(st, request, value, value2)
End Function

Public Sub opus_multistream_decoder_destroy(ByVal st As Long)
    Static blank As OpusMSDecoder
    If st < 1 Or st > msdecs_count Then Exit Sub
    opus_multistream_vbFreeStreams st
    msdecs(st) = blank                                  ' opus_free(st)
End Sub
