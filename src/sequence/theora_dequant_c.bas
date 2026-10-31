Attribute VB_Name = "theora_dequant_c"
' ref/theora (libtheora 1.2.0) lib/dequant.c with dequant.h: unpacking the quantization parameters of the
' setup header.
' (VB) The sizes lists and base matrices the ranges point to are stored in the th_quant_info itself
' (theora_codec_h: vb_sizes, vb_matrices); a range's pointers are indices into them, so a copied range shares
' its lists as in the C. Freeing them is erasing that storage.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Function oc_quant_params_unpack(opb As oc_pack_buf, qinfo As th_quant_info) As Long
    Dim base_mats() As Byte     ' th_quant_base *base_mats
    Dim val_ As Long
    Dim nbase_mats As Long
    Dim sizes(0 To 63) As Long
    Dim indices(0 To 63) As Long
    Dim nbits As Long
    Dim bmi As Long
    Dim ci As Long
    Dim qti As Long
    Dim pli As Long
    Dim qri As Long
    Dim qi As Long
    Dim i As Long
    Dim qtj As Long
    Dim plj As Long
    Dim qrsizes As Long
    Dim qrbms As Long
    val_ = oc_pack_read(opb, 3)
    nbits = val_
    For qi = 0 To 63
        val_ = oc_pack_read(opb, nbits)
        qinfo.loop_filter_limits(qi) = val_ And 255
    Next
    val_ = oc_pack_read(opb, 4)
    nbits = val_ + 1
    For qi = 0 To 63
        val_ = oc_pack_read(opb, nbits)
        qinfo.ac_scale(qi) = val_ And &HFFFF&
    Next
    val_ = oc_pack_read(opb, 4)
    nbits = val_ + 1
    For qi = 0 To 63
        val_ = oc_pack_read(opb, nbits)
        qinfo.dc_scale(qi) = val_ And &HFFFF&
    Next
    val_ = oc_pack_read(opb, 9)
    nbase_mats = val_ + 1
    ReDim base_mats(0 To nbase_mats * 64 - 1)
    For bmi = 0 To nbase_mats - 1
        For ci = 0 To 63
            val_ = oc_pack_read(opb, 8)
            base_mats(bmi * 64 + ci) = val_ And 255
        Next
    Next
    nbits = oc_ilog(nbase_mats - 1)
    For i = 0 To 5
        qti = i \ 3
        pli = i Mod 3
        ' qranges=_qinfo->qi_ranges[qti]+pli: (VB) qinfo.qi_ranges(qti * 3 + pli)
        If i > 0 Then
            val_ = oc_pack_read1(opb)
            If val_ = 0 Then
                If qti > 0 Then
                    val_ = oc_pack_read1(opb)
                    If val_ <> 0 Then
                        qtj = qti - 1
                        plj = pli
                    Else
                        qtj = (i - 1) \ 3
                        plj = (i - 1) Mod 3
                    End If
                Else
                    qtj = (i - 1) \ 3
                    plj = (i - 1) Mod 3
                End If
                qinfo.qi_ranges(qti * 3 + pli) = qinfo.qi_ranges(qtj * 3 + plj)
                GoTo continue_
            End If
        End If
        val_ = oc_pack_read(opb, nbits)
        indices(0) = val_
        qi = 0
        qri = 0
        Do While qi < 63
            val_ = oc_pack_read(opb, oc_ilog(62 - qi))
            sizes(qri) = val_ + 1
            qi = qi + val_ + 1
            val_ = oc_pack_read(opb, nbits)
            qri = qri + 1
            indices(qri) = val_
        Loop
        ' Note: The caller is responsible for cleaning up any partially constructed qinfo.
        If qi > 63 Then
            oc_quant_params_unpack = TH_EBADHEADER
            Exit Function
        End If
        qinfo.qi_ranges(qti * 3 + pli).nranges = qri
        ' qranges->sizes=qrsizes=(int *)_ogg_malloc(qri*sizeof(qrsizes[0])): (VB) the next qri entries
        qrsizes = qinfo.vb_nsizes
        If qinfo.vb_nsizes = 0 Then
            ReDim qinfo.vb_sizes(0 To qri - 1)
        Else
            ReDim Preserve qinfo.vb_sizes(0 To qinfo.vb_nsizes + qri - 1)
        End If
        qinfo.vb_nsizes = qinfo.vb_nsizes + qri
        qinfo.qi_ranges(qti * 3 + pli).sizes = qrsizes
        CopyMemory qinfo.vb_sizes(qrsizes), sizes(0), qri * 4
        ' qrbms=(th_quant_base *)_ogg_malloc((qri+1)*sizeof(qrbms[0])): (VB) the next qri+1 matrices
        qrbms = qinfo.vb_nmatrices
        If qinfo.vb_nmatrices = 0 Then
            ReDim qinfo.vb_matrices(0 To (qri + 1) * 64 - 1)
        Else
            ReDim Preserve qinfo.vb_matrices(0 To (qinfo.vb_nmatrices + qri + 1) * 64 - 1)
        End If
        qinfo.vb_nmatrices = qinfo.vb_nmatrices + qri + 1
        qinfo.qi_ranges(qti * 3 + pli).base_matrices = qrbms
        Do
            bmi = indices(qri)
            ' Note: The caller is responsible for cleaning up any partially constructed qinfo.
            If bmi >= nbase_mats Then
                oc_quant_params_unpack = TH_EBADHEADER
                Exit Function
            End If
            CopyMemory qinfo.vb_matrices((qrbms + qri) * 64), base_mats(bmi * 64), 64
            qri = qri - 1
        Loop While qri + 1 > 0          ' while(qri-->0)
continue_:
    Next
    oc_quant_params_unpack = 0
End Function

Public Sub oc_quant_params_clear(qinfo As th_quant_info)
    Dim i As Long
    Dim qti As Long
    Dim pli As Long
    Dim qtj As Long
    Dim plj As Long
    For i = 5 To 0 Step -1
        qti = i \ 3
        pli = i Mod 3
        ' Clear any duplicate pointer references.
        If i > 0 Then
            qtj = (i - 1) \ 3
            plj = (i - 1) Mod 3
            If qinfo.qi_ranges(qti * 3 + pli).sizes = qinfo.qi_ranges(qtj * 3 + plj).sizes Then
                qinfo.qi_ranges(qti * 3 + pli).sizes = -1
            End If
            If qinfo.qi_ranges(qti * 3 + pli).base_matrices = qinfo.qi_ranges(qtj * 3 + plj).base_matrices Then
                qinfo.qi_ranges(qti * 3 + pli).base_matrices = -1
            End If
        End If
        If qti > 0 Then
            If qinfo.qi_ranges(3 + pli).sizes = qinfo.qi_ranges(pli).sizes Then
                qinfo.qi_ranges(3 + pli).sizes = -1
            End If
            If qinfo.qi_ranges(3 + pli).base_matrices = qinfo.qi_ranges(pli).base_matrices Then
                qinfo.qi_ranges(3 + pli).base_matrices = -1
            End If
        End If
        ' Now free all the non-duplicate storage. (VB) all of it is erased below
        qinfo.qi_ranges(qti * 3 + pli).sizes = -1
        qinfo.qi_ranges(qti * 3 + pli).base_matrices = -1
    Next
    Erase qinfo.vb_sizes
    qinfo.vb_nsizes = 0
    Erase qinfo.vb_matrices
    qinfo.vb_nmatrices = 0
End Sub
