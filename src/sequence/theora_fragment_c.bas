Attribute VB_Name = "theora_fragment_c"
' ref/theora (libtheora 1.2.0) lib/fragment.c: copying and reconstructing 8x8 fragments (the C versions of
' the accelerated functions).
' (VB) A pointer into a frame is an offset into the frame storage pix (oc_theora_state.ref_frame_handle); the
' residue (ogg_int16_t _residue[64]) is residue(resOffset) of a Long array (the pipeline's coefficients).
Option Explicit

Public Sub oc_frag_copy_c(pix() As Byte, ByVal dst As Long, ByVal src As Long, ByVal ystride As Long)
    Dim i As Long
    For i = 7 To 0 Step -1
        ' memcpy(_dst,_src,8*sizeof(*_dst))
        pix(dst) = pix(src)
        pix(dst + 1) = pix(src + 1)
        pix(dst + 2) = pix(src + 2)
        pix(dst + 3) = pix(src + 3)
        pix(dst + 4) = pix(src + 4)
        pix(dst + 5) = pix(src + 5)
        pix(dst + 6) = pix(src + 6)
        pix(dst + 7) = pix(src + 7)
        dst = dst + ystride
        src = src + ystride
    Next
End Sub

' Copies the fragments specified by the lists of fragment indices from one frame to another.
'  _dst_frame:     The reference frame to copy to.
'  _src_frame:     The reference frame to copy from.
'  _ystride:       The row stride of the reference frames.
'  _fragis:        A pointer to a list of fragment indices. (VB) fragis(fragisOffset)
'  _nfragis:       The number of fragment indices to copy.
'  _frag_buf_offs: The offsets of fragments in the reference frames.
Public Sub oc_frag_copy_list_c(pix() As Byte, ByVal dst_frame As Long, ByVal src_frame As Long, ByVal ystride As Long, _
    fragis() As Long, ByVal fragisOffset As Long, ByVal nfragis As Long, frag_buf_offs() As Long)
    Dim fragii As Long
    Dim frag_buf_off As Long
    For fragii = 0 To nfragis - 1
        frag_buf_off = frag_buf_offs(fragis(fragisOffset + fragii))
        oc_frag_copy_c pix, dst_frame + frag_buf_off, src_frame + frag_buf_off, ystride
    Next
End Sub

Public Sub oc_frag_recon_intra_c(pix() As Byte, ByVal dst As Long, ByVal ystride As Long, residue() As Long, _
    ByVal resOffset As Long)
    Dim i As Long
    Dim j As Long
    Dim v As Long
    For i = 0 To 7
        For j = 0 To 7
            v = residue(resOffset + i * 8 + j) + 128
            If v < 0 Then
                v = 0
            ElseIf v > 255 Then
                v = 255
            End If
            pix(dst + j) = v
        Next
        dst = dst + ystride
    Next
End Sub

Public Sub oc_frag_recon_inter_c(pix() As Byte, ByVal dst As Long, ByVal src As Long, ByVal ystride As Long, _
    residue() As Long, ByVal resOffset As Long)
    Dim i As Long
    Dim j As Long
    Dim v As Long
    For i = 0 To 7
        For j = 0 To 7
            v = residue(resOffset + i * 8 + j) + pix(src + j)
            If v < 0 Then
                v = 0
            ElseIf v > 255 Then
                v = 255
            End If
            pix(dst + j) = v
        Next
        dst = dst + ystride
        src = src + ystride
    Next
End Sub

Public Sub oc_frag_recon_inter2_c(pix() As Byte, ByVal dst As Long, ByVal src1 As Long, ByVal src2 As Long, _
    ByVal ystride As Long, residue() As Long, ByVal resOffset As Long)
    Dim i As Long
    Dim j As Long
    Dim v As Long
    For i = 0 To 7
        For j = 0 To 7
            v = residue(resOffset + i * 8 + j) + ((CLng(pix(src1 + j)) + pix(src2 + j)) \ 2)
            If v < 0 Then
                v = 0
            ElseIf v > 255 Then
                v = 255
            End If
            pix(dst + j) = v
        Next
        dst = dst + ystride
        src1 = src1 + ystride
        src2 = src2 + ystride
    Next
End Sub

Public Sub oc_restore_fpu_c()
End Sub
