Attribute VB_Name = "opus_c"
' libopus 1.5.2 (ref/opus/src): opus.c whole (opus_pcm_soft_clip, encode_size, parse_size,
' opus_packet_get_samples_per_frame, opus_packet_parse_impl, opus_packet_parse), with celt/celt.c's
' opus_strerror and opus_get_version_string (the public API's, here per K_DESIGN; the float build's version
' string has no "-fixed").
' (VB) A const unsigned char * into a packet is the packet array and an index; frames[] holds indexes into it.
' The C's optional out pointers (out_toc, frames, payload_offset, packet_offset, padding, padding_len) are always
' written: a caller that passes NULL in C passes a dummy. opus_pcm_soft_clip's float arithmetic is float_vb's
' (each operation rounded to float, as C's; no NaN handling: the decoder's output has none).
Option Explicit

Public Sub opus_pcm_soft_clip(pcm() As Single, ByVal pcmOffset As Long, ByVal N As Long, ByVal C As Long, _
        declip_mem() As Single, ByVal memOffset As Long)
    Dim cc As Long
    Dim i As Long
    Dim x As Long                                       ' float *: an offset into pcm
    Dim a As Single
    Dim x0 As Single
    Dim curr As Long
    Dim start As Long, end_ As Long
    Dim maxval As Single
    Dim special As Long
    Dim peak_pos As Long
    Dim delta As Single
    Dim offset As Single
    Dim v As Single

    If C < 1 Or N < 1 Then Exit Sub                     ' || !_x || !declip_mem: (VB) arrays

    '/* First thing: saturate everything to +/- 2 which is the highest level our
    '   non-linearity can handle. At the point where the signal reaches +/-2,
    '   the derivative will be zero anyway, so this doesn't introduce any
    '   discontinuity in the derivative. */
    For i = 0 To N * C - 1
        v = pcm(pcmOffset + i)
        If 2! < v Then v = 2!                           ' MIN16(2.f, _x[i])
        If -2! > v Then v = -2!                         ' MAX16(-2.f, ...)
        pcm(pcmOffset + i) = v
    Next
    For cc = 0 To C - 1
        x = pcmOffset + cc
        a = declip_mem(memOffset + cc)
        '/* Continue applying the non-linearity from the previous frame to avoid
        '   any discontinuity. */
        For i = 0 To N - 1
            If fmul(pcm(x + i * C), a) >= 0 Then
                Exit For
            End If
            pcm(x + i * C) = fadd(pcm(x + i * C), fmul(fmul(a, pcm(x + i * C)), pcm(x + i * C)))
        Next

        curr = 0
        x0 = pcm(x)
        Do
            special = 0
            For i = curr To N - 1
                If pcm(x + i * C) > 1 Or pcm(x + i * C) < -1 Then
                    Exit For
                End If
            Next
            If i = N Then
                a = 0
                Exit Do
            End If
            peak_pos = i
            start = i: end_ = i
            maxval = Abs(pcm(x + i * C))
            '/* Look for first zero crossing before clipping */
            Do While start > 0
                If Not (fmul(pcm(x + i * C), pcm(x + (start - 1) * C)) >= 0) Then Exit Do
                start = start - 1
            Loop
            '/* Look for first zero crossing after clipping */
            Do While end_ < N
                If Not (fmul(pcm(x + i * C), pcm(x + end_ * C)) >= 0) Then Exit Do
                '/* Look for other peaks until the next zero-crossing. */
                If Abs(pcm(x + end_ * C)) > maxval Then
                    maxval = Abs(pcm(x + end_ * C))
                    peak_pos = end_
                End If
                end_ = end_ + 1
            Loop
            '/* Detect the special case where we clip before the first zero crossing */
            If start = 0 Then
                If fmul(pcm(x + i * C), pcm(x)) >= 0 Then special = 1
            End If

            '/* Compute a such that maxval + a*maxval^2 = 1 */
            a = fdiv(fsub(maxval, 1!), fmul(maxval, maxval))
            '/* Slightly boost "a" by 2^-22. This is just enough to ensure -ffast-math
            '   does not cause output values larger than +/-1, but small enough not
            '   to matter even for 24-bit output.  */
            a = fadd(a, fmul(a, 2.4E-07!))
            If pcm(x + i * C) > 0 Then
                a = -a
            End If
            '/* Apply soft clipping */
            For i = start To end_ - 1
                pcm(x + i * C) = fadd(pcm(x + i * C), fmul(fmul(a, pcm(x + i * C)), pcm(x + i * C)))
            Next

            If special <> 0 And peak_pos >= 2 Then
                '/* Add a linear ramp from the first sample to the signal peak.
                '   This avoids a discontinuity at the beginning of the frame. */
                offset = fsub(x0, pcm(x))
                delta = fdiv(offset, CSng(peak_pos))
                For i = curr To peak_pos - 1
                    offset = fsub(offset, delta)
                    pcm(x + i * C) = fadd(pcm(x + i * C), offset)
                    v = pcm(x + i * C)
                    If 1! < v Then v = 1!               ' MIN16(1.f, x[i*C])
                    If -1! > v Then v = -1!             ' MAX16(-1.f, ...)
                    pcm(x + i * C) = v
                Next
            End If
            curr = end_
            If curr = N Then
                Exit Do
            End If
        Loop
        declip_mem(memOffset + cc) = a
    Next
End Sub

Public Function encode_size(ByVal size As Long, data() As Byte, ByVal dataOffset As Long) As Long
    If size < 252 Then
        data(dataOffset) = size
        encode_size = 1
    Else
        data(dataOffset) = 252 + (size And &H3)
        data(dataOffset + 1) = (size - data(dataOffset)) \ 4
        encode_size = 2
    End If
End Function

Private Function parse_size(data() As Byte, ByVal dataOffset As Long, ByVal len_ As Long, size() As Integer, ByVal sizeIndex As Long) As Long
    If len_ < 1 Then
        size(sizeIndex) = -1
        parse_size = -1
    ElseIf data(dataOffset) < 252 Then
        size(sizeIndex) = data(dataOffset)
        parse_size = 1
    ElseIf len_ < 2 Then
        size(sizeIndex) = -1
        parse_size = -1
    Else
        size(sizeIndex) = 4 * CLng(data(dataOffset + 1)) + data(dataOffset)
        parse_size = 2
    End If
End Function

' (VB) x << n for the TOC's 0 <= n <= 3. The packet functions run before any decoder exists (opusfile calls them
' while it opens a file), so they use no silk_SigProc_FIX_h macro: those read tables that silk_InitDecoder fills.
Private Function shl_small(ByVal x As Long, ByVal n As Long) As Long
    Select Case n
        Case 0: shl_small = x
        Case 1: shl_small = x * 2
        Case 2: shl_small = x * 4
        Case Else: shl_small = x * 8
    End Select
End Function

Public Function opus_packet_get_samples_per_frame(data() As Byte, ByVal dataOffset As Long, ByVal Fs As Long) As Long
    Dim audiosize As Long
    If (data(dataOffset) And &H80) <> 0 Then
        audiosize = (data(dataOffset) \ 8) And &H3
        audiosize = shl_small(Fs, audiosize) \ 400      ' (Fs<<audiosize)/400
    ElseIf (data(dataOffset) And &H60) = &H60 Then
        If (data(dataOffset) And &H8) <> 0 Then audiosize = Fs \ 50 Else audiosize = Fs \ 100
    Else
        audiosize = (data(dataOffset) \ 8) And &H3
        If audiosize = 3 Then
            audiosize = (Fs * 60) \ 1000
        Else
            audiosize = shl_small(Fs, audiosize) \ 100  ' (Fs<<audiosize)/100
        End If
    End If
    opus_packet_get_samples_per_frame = audiosize
End Function

Public Function opus_packet_parse_impl(data() As Byte, ByVal dataOffset As Long, ByVal len_ As Long, ByVal self_delimited As Long, _
        out_toc As Long, frames() As Long, size() As Integer, payload_offset As Long, packet_offset As Long, _
        padding As Long, padding_len As Long) As Long
    Dim i As Long, bytes As Long
    Dim count As Long
    Dim cbr As Long
    Dim ch As Long, toc As Long
    Dim framesize As Long
    Dim last_size As Long
    Dim pad As Long
    Dim data0 As Long                                   ' const unsigned char *data0 = data
    Dim p As Long, tmp As Long

    data0 = dataOffset

    If len_ < 0 Then                                    ' size==NULL || len<0: (VB) size is an array
        opus_packet_parse_impl = OPUS_BAD_ARG
        Exit Function
    End If
    If len_ = 0 Then
        opus_packet_parse_impl = OPUS_INVALID_PACKET
        Exit Function
    End If

    framesize = opus_packet_get_samples_per_frame(data, dataOffset, 48000)

    cbr = 0
    toc = data(dataOffset): dataOffset = dataOffset + 1
    len_ = len_ - 1
    last_size = len_
    Select Case toc And &H3
    '/* One frame */
    Case 0
        count = 1
    '/* Two CBR frames */
    Case 1
        count = 2
        cbr = 1
        If self_delimited = 0 Then
            If (len_ And &H1) <> 0 Then
                opus_packet_parse_impl = OPUS_INVALID_PACKET
                Exit Function
            End If
            last_size = len_ \ 2
            '/* If last_size doesn't fit in size[0], we'll catch it later */
            size(0) = last_size
        End If
    '/* Two VBR frames */
    Case 2
        count = 2
        bytes = parse_size(data, dataOffset, len_, size, 0)
        len_ = len_ - bytes
        If size(0) < 0 Or size(0) > len_ Then
            opus_packet_parse_impl = OPUS_INVALID_PACKET
            Exit Function
        End If
        dataOffset = dataOffset + bytes
        last_size = len_ - size(0)
    '/* Multiple CBR/VBR frames (from 0 to 120 ms) */
    Case Else '/*case 3:*/
        If len_ < 1 Then
            opus_packet_parse_impl = OPUS_INVALID_PACKET
            Exit Function
        End If
        '/* Number of frames encoded in bits 0 to 5 */
        ch = data(dataOffset): dataOffset = dataOffset + 1
        count = ch And &H3F
        If count <= 0 Or framesize * count > 5760 Then
            opus_packet_parse_impl = OPUS_INVALID_PACKET
            Exit Function
        End If
        len_ = len_ - 1
        '/* Padding flag is bit 6 */
        If (ch And &H40) <> 0 Then
            Do
                If len_ <= 0 Then
                    opus_packet_parse_impl = OPUS_INVALID_PACKET
                    Exit Function
                End If
                p = data(dataOffset): dataOffset = dataOffset + 1
                len_ = len_ - 1
                If p = 255 Then tmp = 254 Else tmp = p
                len_ = len_ - tmp
                pad = pad + tmp
            Loop While p = 255
        End If
        If len_ < 0 Then
            opus_packet_parse_impl = OPUS_INVALID_PACKET
            Exit Function
        End If
        '/* VBR flag is bit 7 */
        If (ch And &H80) = 0 Then cbr = 1 Else cbr = 0
        If cbr = 0 Then
            '/* VBR case */
            last_size = len_
            For i = 0 To count - 2
                bytes = parse_size(data, dataOffset, len_, size, i)
                len_ = len_ - bytes
                If size(i) < 0 Or size(i) > len_ Then
                    opus_packet_parse_impl = OPUS_INVALID_PACKET
                    Exit Function
                End If
                dataOffset = dataOffset + bytes
                last_size = last_size - (bytes + size(i))
            Next
            If last_size < 0 Then
                opus_packet_parse_impl = OPUS_INVALID_PACKET
                Exit Function
            End If
        ElseIf self_delimited = 0 Then
            '/* CBR case */
            last_size = len_ \ count
            If last_size * count <> len_ Then
                opus_packet_parse_impl = OPUS_INVALID_PACKET
                Exit Function
            End If
            For i = 0 To count - 2
                size(i) = last_size
            Next
        End If
    End Select
    '/* Self-delimited framing has an extra size for the last frame. */
    If self_delimited <> 0 Then
        bytes = parse_size(data, dataOffset, len_, size, count - 1)
        len_ = len_ - bytes
        If size(count - 1) < 0 Or size(count - 1) > len_ Then
            opus_packet_parse_impl = OPUS_INVALID_PACKET
            Exit Function
        End If
        dataOffset = dataOffset + bytes
        '/* For CBR packets, apply the size to all the frames. */
        If cbr <> 0 Then
            If CLng(size(count - 1)) * count > len_ Then
                opus_packet_parse_impl = OPUS_INVALID_PACKET
                Exit Function
            End If
            For i = 0 To count - 2
                size(i) = size(count - 1)
            Next
        ElseIf bytes + size(count - 1) > last_size Then
            opus_packet_parse_impl = OPUS_INVALID_PACKET
            Exit Function
        End If
    Else
        '/* Because it's not encoded explicitly, it's possible the size of the
        '   last packet (or all the packets, for the CBR case) is larger than
        '   1275. Reject them here.*/
        If last_size > 1275 Then
            opus_packet_parse_impl = OPUS_INVALID_PACKET
            Exit Function
        End If
        size(count - 1) = last_size
    End If

    payload_offset = dataOffset - data0

    For i = 0 To count - 1
        frames(i) = dataOffset
        dataOffset = dataOffset + size(i)
    Next

    padding = dataOffset
    padding_len = pad
    packet_offset = pad + (dataOffset - data0)

    out_toc = toc

    opus_packet_parse_impl = count
End Function

Public Function opus_packet_parse(data() As Byte, ByVal dataOffset As Long, ByVal len_ As Long, out_toc As Long, _
        frames() As Long, size() As Integer, payload_offset As Long) As Long
    Dim packet_offset As Long, padding As Long, padding_len As Long
    opus_packet_parse = opus_packet_parse_impl(data, dataOffset, len_, 0, out_toc, frames, size, payload_offset, _
        packet_offset, padding, padding_len)
End Function

' celt/celt.c
Public Function opus_strerror(ByVal error_ As Long) As String
    ' static const char * const error_strings[8]
    If error_ > 0 Or error_ < -7 Then
        opus_strerror = "unknown error"
    Else
        Select Case -error_
            Case 0: opus_strerror = "success"
            Case 1: opus_strerror = "invalid argument"
            Case 2: opus_strerror = "buffer too small"
            Case 3: opus_strerror = "internal error"
            Case 4: opus_strerror = "corrupted stream"
            Case 5: opus_strerror = "request not implemented"
            Case 6: opus_strerror = "invalid state"
            Case 7: opus_strerror = "memory allocation failed"
        End Select
    End If
End Function

Public Function opus_get_version_string() As String
    opus_get_version_string = "libopus " & "1.5.2"     ' "libopus " PACKAGE_VERSION
    '/* Applications may rely on the presence of this substring in the version
    '   string to determine if they have a fixed-point or floating-point build
    '   at runtime. */
End Function
