Attribute VB_Name = "theora_quant_c"
' ref/theora (libtheora 1.2.0) lib/quant.c with quant.h: the dequantization tables. OC_DC_QUANT_MIN and
' OC_AC_QUANT_MIN are in the generated theora_tables_vb.
' (VB) ogg_uint16_t *_dequant[64][3][2] is an array of offsets into the tables' storage, dequant((qi * 3 +
' pli) * 2 + qti); a table (oc_quant_table, ogg_uint16_t[64]) is 64 Longs of that storage.
Option Explicit

' typedef ogg_uint16_t oc_quant_table[64]: (VB) 64 Longs

' Maximum scaled quantizer value.
Public Const OC_QUANT_MAX As Long = 1024 * 4

' The maximum output of the DCT with +/- 255 inputs is +/- 8157. These minimum quantizers ensure the result
' after quantization (and after prediction for DC) will be no more than +/- 510. The tokenization system can
' handle values up to +/- 580, so there is no need to do any coefficient clamping. I would rather have
' allowed smaller quantizers and had to clamp, but these minimums were required when constructing the
' original VP3 matrices and have been formalized in the spec.
' (OC_DC_QUANT_MIN[2]={4<<2,8<<2}; OC_AC_QUANT_MIN[2]={2<<2,4<<2}: theora_tables_vb)

' Initializes the dequantization tables from a set of quantizer info. Currently the dequantizer (and
' elsewhere enquantizer) tables are expected to be initialized as pointing to the storage reserved for them
' in the oc_theora_state (resp. oc_enc_ctx) structure. If some tables are duplicates of others, the pointers
' will be adjusted to point to a single copy of the tables, but the storage for them will not be freed. If
' you're concerned about the memory footprint, the obvious thing to do is to move the storage out of its
' fixed place in the structures and allocate it on demand. However, a much, much better option is to only
' store the quantization matrices being used for the current frame, and to recalculate these as the qi
' values change between frames (this is what VP3 did).
' (VB) dequant: the pointers (offsets into data); pp_dc_scale NULL <=> ppIsNull
Public Sub oc_dequant_tables_init(dequant() As Long, data() As Long, pp_dc_scale() As Long, _
    ByVal ppIsNull As Boolean, qinfo As th_quant_info)
    ' Coding mode: intra or inter.
    Dim qti As Long
    ' Y', C_b, C_r
    Dim pli As Long
    ' Quality index.
    Dim qi As Long
    ' Range iterator.
    Dim qri As Long
    Dim basem(0 To 63) As Long      ' th_quant_base base
    Dim q As Long                   ' ogg_uint32_t
    Dim qi_start As Long
    Dim qi_end As Long
    Dim qfac As Long                ' ogg_uint32_t
    Dim zzi As Long
    Dim ci As Long
    Dim dupe As Long
    Dim qtj As Long
    Dim plj As Long
    Dim plEnd As Long
    Dim r As Long                   ' (VB) qinfo.qi_ranges(qti * 3 + pli)
    Dim bm As Long                  ' (VB) the range's base_matrices
    Dim t As Long                   ' (VB) _dequant[qi][pli][qti]
    Dim t2 As Long
    Dim k As Long
    For qti = 0 To 1
        For pli = 0 To 2
            r = qti * 3 + pli
            bm = qinfo.qi_ranges(r).base_matrices
            qi = 0
            For qri = 0 To qinfo.qi_ranges(r).nranges
                ' memcpy(base,_qinfo->qi_ranges[qti][pli].base_matrices[qri],sizeof(base))
                For ci = 0 To 63
                    basem(ci) = qinfo.vb_matrices((bm + qri) * 64 + ci)
                Next
                qi_start = qi
                If qri = qinfo.qi_ranges(r).nranges Then
                    qi_end = qi + 1
                Else
                    qi_end = qi + qinfo.vb_sizes(qinfo.qi_ranges(r).sizes + qri)
                End If
                ' Iterate over quality indices in this range.
                Do
                    ' In the original VP3.2 code, the rounding offset and the size of the dead zone around 0
                    ' were controlled by a "sharpness" parameter. The size of our dead zone is now controlled
                    ' by the per-coefficient quality thresholds returned by our HVS module. We round down from
                    ' a more accurate value when the quality of the reconstruction does not fall below our
                    ' threshold and it saves bits. Hence, all of that VP3.2 code is gone from here, and the
                    ' remaining floating point code has been implemented as equivalent integer code with exact
                    ' precision.
                    qfac = qinfo.dc_scale(qi) * basem(0)
                    ' For postprocessing, not dequantization.
                    If Not ppIsNull Then pp_dc_scale(qi) = qfac \ 160
                    ' Scale DC the coefficient from the proper table.
                    q = (qfac \ 100) * 4
                    q = OC_CLAMPI(OC_DC_QUANT_MIN(qti), q, OC_QUANT_MAX)
                    t = dequant((qi * 3 + pli) * 2 + qti)
                    data(t) = q
                    ' Now scale AC coefficients from the proper table.
                    For zzi = 1 To 63
                        q = (qinfo.ac_scale(qi) * basem(OC_FZIG_ZAG(zzi)) \ 100) * 4
                        q = OC_CLAMPI(OC_AC_QUANT_MIN(qti), q, OC_QUANT_MAX)
                        data(t + zzi) = q
                    Next
                    ' If this is a duplicate of a previous matrix, use that instead. This simple check helps us
                    ' improve cache coherency later.
                    dupe = 0
                    For qtj = 0 To qti
                        If qtj < qti Then plEnd = 3 Else plEnd = pli
                        For plj = 0 To plEnd - 1
                            ' !memcmp(_dequant[qi][pli][qti],_dequant[qi][plj][qtj],sizeof(oc_quant_table))
                            t2 = dequant((qi * 3 + plj) * 2 + qtj)
                            dupe = 1
                            For k = 0 To 63
                                If data(t + k) <> data(t2 + k) Then dupe = 0: Exit For
                            Next
                            If dupe <> 0 Then Exit For
                        Next
                        If dupe <> 0 Then Exit For
                    Next
                    If dupe <> 0 Then dequant((qi * 3 + pli) * 2 + qti) = dequant((qi * 3 + plj) * 2 + qtj)
                    qi = qi + 1
                    If qi >= qi_end Then Exit Do
                    ' Interpolate the next base matrix.
                    For ci = 0 To 63
                        basem(ci) = ((2 * ((qi_end - qi) * qinfo.vb_matrices((bm + qri) * 64 + ci) + _
                            (qi - qi_start) * qinfo.vb_matrices((bm + qri + 1) * 64 + ci)) + _
                            qinfo.vb_sizes(qinfo.qi_ranges(r).sizes + qri)) \ _
                            (2 * qinfo.vb_sizes(qinfo.qi_ranges(r).sizes + qri))) And 255
                    Next
                Loop
            Next
        Next
    Next
End Sub
