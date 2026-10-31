Attribute VB_Name = "theora_decint_h"
' ref/theora (libtheora 1.2.0) lib/decint.h: the decoder's types (th_setup_info, oc_dec_pipeline_state,
' th_dec_ctx) and (VB) their pools. The C paths only (no OC_DEC_USE_VTABLE, no HAVE_CAIRO telemetry).
' (VB) A th_setup_info * is a handle into oc_setups (theora_decinfo_c allocates them), a th_dec_ctx * a handle
' into oc_decs (theora_decode_c); 0 = NULL. The procedures take the pool element ByRef. The C's pointers
' into the decoder's arrays are indices; a set of Huffman trees is offsets into a Long array
' (theora_huffdec_c).
Option Explicit

' Constants for the packet-in state machine specific to the decoder.
' Next packet to read: Data packet.
Public Const OC_PACKET_DATA As Long = 0

Public Type th_setup_info
    ' The Huffman codes.
    huff_tables() As Long       ' ogg_int16_t *huff_tables[TH_NHUFFMAN_TABLES]: (VB) offsets into vb_huff
    vb_huff() As Long           ' (VB) the trees' storage
    vb_nhuff As Long
    ' The quantization parameters.
    qinfo As th_quant_info
    vbUsed As Boolean           ' (VB) the pool entry is allocated
End Type

' Decoder specific functions with accelerated variants (oc_dec_opt_vtable): (VB) not used (the C paths)

Public Type oc_dec_pipeline_state
    ' Decoded DCT coefficients. These are placed here instead of on the stack so that they can persist
    ' between blocks, which makes clearing them back to zero much faster when only a few non-zero
    ' coefficients were decoded. It requires at least 65 elements because the zig-zag index array uses the
    ' 65th element as a dumping ground for out-of-range indices to protect us from buffer overflow. We make it
    ' fully twice as large so that the second half can serve as the reconstruction buffer, which saves passing
    ' another parameter to all the acceleration functions. It also solves problems with 16-byte alignment for
    ' NEON on ARM. gcc (as of 4.2.1) only seems to be able to give stack variables 8-byte alignment, and
    ' silently produces incorrect results if you ask for 16. Finally, keeping it off the stack means there's
    ' less likely to be a data hazard between the NEON co-processor and the regular ARM core, which avoids
    ' unnecessary stalls.
    dct_coeffs() As Long        ' OC_ALIGN16(ogg_int16_t dct_coeffs[128]): (VB) 128 Longs holding int16 values
    bounding_values() As Long   ' OC_ALIGN16(signed char bounding_values[256])
    ti(0 To 191) As Long        ' ptrdiff_t ti[3][64]: (VB) ti(pli * 64 + zzi)
    ' ptrdiff_t ebi[3][64]: unused
    eob_runs(0 To 191) As Long  ' ptrdiff_t eob_runs[3][64]: (VB) eob_runs(pli * 64 + zzi)
    coded_fragis(0 To 2) As Long    ' const ptrdiff_t *: (VB) indices into state.coded_fragis
    uncoded_fragis(0 To 2) As Long  ' const ptrdiff_t *: (VB) indices into state.coded_fragis
    ncoded_fragis(0 To 2) As Long
    nuncoded_fragis(0 To 2) As Long
    dequant(0 To 17) As Long    ' const ogg_uint16_t *dequant[3][3][2]: (VB) dequant((pli * 3 + qii) * 2 + qti),
                                ' offsets into state.dequant_table_data
    fragy0(0 To 2) As Long
    fragy_end(0 To 2) As Long
    pred_last(0 To 11) As Long  ' int pred_last[3][4]: (VB) pred_last(pli * 4 + refi)
    mcu_nvfrags As Long
    loop_filter As Long
    pp_level As Long
End Type

Public Type th_dec_ctx
    ' Shared encoder/decoder state.
    state As oc_theora_state
    ' Whether or not packets are ready to be emitted. This takes on negative values while there are remaining
    ' header packets to be emitted, reaches 0 when the codec is ready for input, and goes to 1 when a frame
    ' has been processed and a data packet is ready.
    packet_state As Long
    ' Buffer in which to assemble packets.
    opb As oc_pack_buf
    ' Huffman decode trees.
    huff_tables() As Long       ' ogg_int16_t *huff_tables[TH_NHUFFMAN_TABLES]: (VB) offsets into vb_huff
    vb_huff() As Long           ' (VB) the trees' storage
    vb_nhuff As Long
    ' The index of the first token in each plane for each coefficient.
    ti0(0 To 191) As Long       ' ptrdiff_t ti0[3][64]: (VB) ti0(pli * 64 + zzi)
    ' The number of outstanding EOB runs at the start of each coefficient in each plane.
    eob_runs(0 To 191) As Long  ' ptrdiff_t eob_runs[3][64]
    ' The DCT token lists.
    dct_tokens() As Byte
    ' The extra bits associated with DCT tokens. (unsigned char *extra_bits: unused)
    ' The number of dct tokens unpacked so far.
    dct_tokens_count As Long
    ' The out-of-loop post-processing level.
    pp_level As Long
    ' The DC scale used for out-of-loop deblocking.
    pp_dc_scale() As Long       ' int pp_dc_scale[64]
    ' The sharpen modifier used for out-of-loop deringing.
    pp_sharp_mod(0 To 63) As Long
    ' The DC quantization index of each block.
    dc_qis() As Byte
    vbDcQisAllocated As Boolean     ' (VB) dc_qis != NULL
    ' The variance of each block.
    variances() As Long
    vbVariancesAllocated As Boolean ' (VB) variances != NULL
    ' The storage for the post-processed frame buffer.
    pp_frame_data As Long       ' unsigned char *: (VB) an offset into state.ref_frame_handle, 0 = NULL
    ' Whether or not the post-processsed frame buffer has space for chroma.
    pp_frame_state As Long
    ' The buffer used for the post-processed frame. Note that this is _not_ guaranteed to have the same strides
    ' and offsets as the reference frame buffers.
    pp_frame_buf As th_ycbcr_buffer
    ' The striped decode callback function. (th_stripe_callback stripe_cb: not ported, no callback is called)
    pipe As oc_dec_pipeline_state
    vbUsed As Boolean           ' (VB) the pool entry is allocated
End Type

' (VB) the pools: index = handle (1-based)
Public oc_setups() As th_setup_info
Public oc_nsetups As Long
Public oc_decs() As th_dec_ctx
Public oc_ndecs As Long
