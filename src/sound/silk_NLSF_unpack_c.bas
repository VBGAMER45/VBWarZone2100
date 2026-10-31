Attribute VB_Name = "silk_NLSF_unpack_c"
' libopus 1.5.2 (ref/opus/silk): NLSF_unpack.c. (VB) psNLSF_CB is an index of silk_NLSF_CBs.
Option Explicit

'/* Unpack predictor values and indices for entropy coding tables */
Public Sub silk_NLSF_unpack(ec_ix() As Integer, pred_Q8() As Byte, ByVal psNLSF_CB As Long, ByVal CB1_index As Long)
    Dim i As Long
    Dim entry As Long                                   ' opus_uint8
    Dim ec_sel_ptr As Long                              ' const opus_uint8 *: an offset into silk_rom_u8
    Dim order As Long, pred As Long

    order = silk_NLSF_CBs(psNLSF_CB).order
    pred = silk_NLSF_CBs(psNLSF_CB).pred_Q8
    ec_sel_ptr = silk_NLSF_CBs(psNLSF_CB).ec_sel + (CB1_index * order) \ 2
    For i = 0 To order - 1 Step 2
        entry = silk_rom_u8(ec_sel_ptr)
        ec_sel_ptr = ec_sel_ptr + 1
        ec_ix(i) = silk_SMULBB(silk_RSHIFT(entry, 1) And 7, 2 * NLSF_QUANT_MAX_AMPLITUDE + 1)
        pred_Q8(i) = silk_rom_u8(pred + i + (entry And 1) * (order - 1))
        ec_ix(i + 1) = silk_SMULBB(silk_RSHIFT(entry, 5) And 7, 2 * NLSF_QUANT_MAX_AMPLITUDE + 1)
        pred_Q8(i + 1) = silk_rom_u8(pred + i + (silk_RSHIFT(entry, 4) And 1) * (order - 1) + 1)
    Next
End Sub
