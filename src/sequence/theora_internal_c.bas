Attribute VB_Name = "theora_internal_c"
' ref/theora (libtheora 1.2.0) lib/internal.c with internal.h and ocintrin.h. The tables of internal.c
' (OC_FZIG_ZAG, OC_IZIG_ZAG, OC_MB_MAP, OC_MB_MAP_IDXS, OC_MB_MAP_NIDXS, OC_DCT_TOKEN_EXTRA_BITS) are in the
' generated theora_tables_vb. Not ported: oc_aligned_malloc, oc_aligned_free, oc_malloc_2d, oc_calloc_2d,
' oc_free_2d (VB arrays; the reference frames' 16-byte alignment is kept as offsets from the start of their
' storage). ocintrin.h's macros are functions here (OC_MAXI, OC_MINI, OC_CLAMPI, OC_CLAMP255, OC_SIGNMASK) for
' the code that is not hot; the hot loops write them out.
' (VB) oc_vb_init fills the tables once; the library's entry points call it.
Option Explicit

' This library's version.
Public Const OC_VENDOR_STRING As String = "Xiph.Org libtheora 1.2.0 20250329 (Ptalarbvorm)"

' Theora bitstream version.
Public Const TH_VERSION_MAJOR As Long = 3
Public Const TH_VERSION_MINOR As Long = 2
Public Const TH_VERSION_SUB As Long = 1

Private vbReady As Boolean

' (VB) the tables of the generated module and of the bit reader, filled once
Public Sub oc_vb_init()
    If vbReady Then Exit Sub
    theora_tables_init
    oc_pack_vbInit
    vbReady = True
End Sub

' TH_VERSION_CHECK(_info,_maj,_min,_sub)
Public Function TH_VERSION_CHECK(info As th_info, ByVal maj As Long, ByVal min_ As Long, ByVal sub_ As Long) As Long
    If info.version_major > maj Then
        TH_VERSION_CHECK = 1
    ElseIf info.version_major = maj Then
        If info.version_minor > min_ Then
            TH_VERSION_CHECK = 1
        ElseIf info.version_minor = min_ Then
            If info.version_subminor >= sub_ Then TH_VERSION_CHECK = 1
        End If
    End If
End Function

' ocintrin.h: OC_MAXI(_a,_b)
Public Function OC_MAXI(ByVal a As Long, ByVal b As Long) As Long
    If b > a Then OC_MAXI = b Else OC_MAXI = a
End Function

' ocintrin.h: OC_MINI(_a,_b)
Public Function OC_MINI(ByVal a As Long, ByVal b As Long) As Long
    If b < a Then OC_MINI = b Else OC_MINI = a
End Function

' ocintrin.h: Clamps an integer into the given range. If _a>_c, then the lower bound _a is respected over the
' upper bound _c (this behavior is required to meet our documented API behavior).
'  _a: The lower bound.
'  _b: The value to clamp.
'  _c: The upper boud.
Public Function OC_CLAMPI(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    OC_CLAMPI = OC_MAXI(a, OC_MINI(b, c))
End Function

' ocintrin.h: OC_CLAMP255(_x)
Public Function OC_CLAMP255(ByVal x As Long) As Long
    If x < 0 Then
        OC_CLAMP255 = 0
    ElseIf x > 255 Then
        OC_CLAMP255 = 255
    Else
        OC_CLAMP255 = x
    End If
End Function

' ocintrin.h: Slightly more portable than relying on a sign-extended right-shift (which is not guaranteed by
' ANSI), and just as fast, since gcc (3.x and 4.x both) compile it into the right-shift anyway.
Public Function OC_SIGNMASK(ByVal a As Long) As Long
    If a < 0 Then OC_SIGNMASK = -1
End Function

' ocintrin.h: Right shifts _dividend by _shift, adding _rval, and subtracting one for negative dividends
' first. When _rval is (1<<_shift-1), this is equivalent to division with rounding ties away from zero.
Public Function OC_DIV_ROUND_POW2(ByVal dividend As Long, ByVal shift As Long, ByVal rval As Long) As Long
    OC_DIV_ROUND_POW2 = SAR(dividend + OC_SIGNMASK(dividend) + rval, shift)
End Function

Public Function oc_ilog(ByVal v As Long) As Long
    Dim ret As Long
    ' for(ret=0;_v;ret++)_v>>=1; (unsigned)
    ret = 0
    Do While v <> 0
        If v < 0 Then v = ((v And &H7FFFFFFF) \ 2) Or &H40000000 Else v = v \ 2
        ret = ret + 1
    Loop
    oc_ilog = ret
End Function

' Fills in a Y'CbCr buffer with a pointer to the image data in the first buffer, but with the opposite
' vertical orientation.
'  _dst: The destination buffer. This can be the same as _src.
'  _src: The source buffer.
Public Sub oc_ycbcr_buffer_flip(dst As th_ycbcr_buffer, src As th_ycbcr_buffer)
    Dim pli As Long
    For pli = 0 To 2
        dst.p(pli).width = src.p(pli).width
        dst.p(pli).height = src.p(pli).height
        dst.p(pli).stride = -src.p(pli).stride
        dst.p(pli).data = src.p(pli).data + (1 - dst.p(pli).height) * dst.p(pli).stride
    Next
End Sub

Public Function th_version_string() As String
    th_version_string = OC_VENDOR_STRING
End Function

Public Function th_version_number() As Long
    th_version_number = TH_VERSION_MAJOR * &H10000 + TH_VERSION_MINOR * &H100& + TH_VERSION_SUB
End Function

' Determines the packet type. Note that this correctly interprets a 0-byte packet as a video data packet.
' Return: 1 for a header packet, 0 for a data packet.
Public Function th_packet_isheader(op As ogg_packet) As Long
    If op.bytes > 0 Then th_packet_isheader = op.packet(0) \ 128 Else th_packet_isheader = 0
End Function

' Determines the frame type of a video data packet. Note that this correctly interprets a 0-byte packet as a
' delta frame.
' Return: 1 for a key frame, 0 for a delta frame, and -1 for a header packet.
Public Function th_packet_iskeyframe(op As ogg_packet) As Long
    If op.bytes <= 0 Then
        th_packet_iskeyframe = 0
    ElseIf (op.packet(0) And &H80) <> 0 Then
        th_packet_iskeyframe = -1
    ElseIf (op.packet(0) And &H40) = 0 Then
        th_packet_iskeyframe = 1
    Else
        th_packet_iskeyframe = 0
    End If
End Function
