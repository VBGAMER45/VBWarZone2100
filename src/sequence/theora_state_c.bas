Attribute VB_Name = "theora_state_c"
' ref/theora (libtheora 1.2.0) lib/state.c with state.h: the state shared by the encoder and the decoder
' (fragment maps, reference frames, motion vectors, the loop filter, granule positions). The C paths only
' (no OC_STATE_USE_VTABLE: oc_frag_copy, oc_idct8x8, ... are the _c functions); oc_state_dump_frame
' (OC_DUMP_IMAGES) is not ported. SB_MAP, OC_MVMAP and OC_MVMAP2 are in the generated theora_tables_vb.
' (VB) The C's pointers into the state's arrays are indices: an oc_sb_map (ptrdiff_t[4][4]) is sb_maps(sbi *
' 16 + quadi * 4 + bi), an oc_mb_map (ptrdiff_t[3][4]) mb_maps(mbi * 12 + pli * 4 + bi). The reference
' frames live in one Byte array, ref_frame_handle, and every frame pointer (th_img_plane.data,
' ref_frame_data) is an offset into it, with the C's arithmetic (a frame is stored bottom-up, its pointers
' flipped as in the C); the decoder's post-processing buffer follows the reference frames in it
' (theora_decode_c). An oc_mv (ogg_int16_t) is a Long holding the sign-extended value. A function pointer
' table (OC_SET_CHROMA_MVS_TABLE, OC_MB_FILL_CMAPPING_TABLE) is a Select Case on the pixel format.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' A keyframe.
Public Const OC_INTRA_FRAME As Long = 0
' A predicted frame.
Public Const OC_INTER_FRAME As Long = 1
' A frame of unknown type (frame type decision has not yet been made).
Public Const OC_UNKWN_FRAME As Long = -1

' The amount of padding to add to the reconstructed frame buffers on all sides. This is used to allow
' unrestricted motion vectors without special casing. This must be a multiple of 2.
Public Const OC_UMV_PADDING As Long = 16

' Frame classification indices.
' The previous golden frame.
Public Const OC_FRAME_GOLD As Long = 0
' The previous frame.
Public Const OC_FRAME_PREV As Long = 1
' The current frame.
Public Const OC_FRAME_SELF As Long = 2
' Used to mark uncoded fragments (for DC prediction).
Public Const OC_FRAME_NONE As Long = 3

' The input or output buffer.
Public Const OC_FRAME_IO As Long = 3
' Uncompressed prev golden frame.
Public Const OC_FRAME_GOLD_ORIG As Long = 4
' Uncompressed previous frame.
Public Const OC_FRAME_PREV_ORIG As Long = 5

' Macroblock modes.
' Macro block is invalid: It is never coded.
Public Const OC_MODE_INVALID As Long = -1
' Encoded difference from the same macro block in the previous frame.
Public Const OC_MODE_INTER_NOMV As Long = 0
' Encoded with no motion compensated prediction.
Public Const OC_MODE_INTRA As Long = 1
' Encoded difference from the previous frame offset by the given motion vector.
Public Const OC_MODE_INTER_MV As Long = 2
' Encoded difference from the previous frame offset by the last coded motion vector.
Public Const OC_MODE_INTER_MV_LAST As Long = 3
' Encoded difference from the previous frame offset by the second to last coded motion vector.
Public Const OC_MODE_INTER_MV_LAST2 As Long = 4
' Encoded difference from the same macro block in the previous golden frame.
Public Const OC_MODE_GOLDEN_NOMV As Long = 5
' Encoded difference from the previous golden frame offset by the given motion vector.
Public Const OC_MODE_GOLDEN_MV As Long = 6
' Encoded difference from the previous frame offset by the individual motion vectors given for each block.
Public Const OC_MODE_INTER_MV_FOUR As Long = 7
' The number of (coded) modes.
Public Const OC_NMODES As Long = 8

' Constants for the packet state machine common between encoder and decoder.
' Next packet to emit/read: Codec info header.
Public Const OC_PACKET_INFO_HDR As Long = -3
' Next packet to emit/read: Comment header.
Public Const OC_PACKET_COMMENT_HDR As Long = -2
' Next packet to emit/read: Codec setup header.
Public Const OC_PACKET_SETUP_HDR As Long = -1
' No more packets to emit/read.
Public Const OC_PACKET_DONE As Long = &H7FFFFFFF

' Super blocks are 32x32 segments of pixels in a single color plane indexed in image order. Internally, super
' blocks are broken up into four quadrants, each of which contains a 2x2 pattern of blocks, each of which is
' an 8x8 block of pixels. Quadrants, and the blocks within them, are indexed in a special order called a
' "Hilbert curve" within the super block.
' In order to differentiate between the Hilbert-curve indexing strategy and the regular image order indexing
' strategy, blocks indexed in image order are called "fragments". Fragments are indexed in image order, left
' to right, then bottom to top, from Y' plane to Cb plane to Cr plane.
' The co-located fragments in all image planes corresponding to the location of a single quadrant of a luma
' plane super block form a macro block. Thus there is only a single set of macro blocks for all planes, each
' of which contains between 6 and 12 fragments, depending on the pixel format. Therefore macro block
' information is kept in a separate set of arrays from super blocks to avoid unused space in the other
' planes. The lists are indexed in super block order. That is, the macro block corresponding to the macro
' block mbi in (luma plane) super block sbi is at index (sbi<<2|mbi). Thus the number of macro blocks in each
' dimension is always twice the number of super blocks, even when only an odd number fall inside the coded
' frame. These "extra" macro blocks are just an artifact of our internal data layout, and not part of the
' coded stream; they are flagged with a negative MB mode.

' Super block information.
Public Type oc_sb_flags
    coded_fully As Byte         ' :1
    coded_partially As Byte     ' :1
    quad_valid As Byte          ' :4
End Type

' Information about a fragment which intersects the border of the displayable region. This marks which
' pixels belong to the displayable region.
Public Type oc_border_info
    ' A bit mask marking which pixels are in the displayable region. Pixel (x,y) corresponds to bit (y<<3|x).
    mask_lo As Long         ' ogg_int64_t mask: (VB) bits 0-31
    mask_hi As Long         ' (VB) bits 32-63
    ' The number of pixels in the displayable region. This is always positive, and always less than 64.
    npixels As Long
End Type

' Fragment information.
Public Type oc_fragment
    ' A flag indicating whether or not this fragment is coded.
    coded As Byte           ' :1
    ' A flag indicating that this entire fragment lies outside the displayable region of the frame. Note the
    ' contrast with an invalid macro block, which is outside the coded frame, not just the displayable one.
    ' There are no fragments outside the coded frame by construction.
    invalid As Byte         ' :1
    ' The index of the quality index used for this fragment's AC coefficients.
    qii As Byte             ' :4
    ' The index of the reference frame this fragment is predicted from.
    refi As Byte            ' :2
    ' The mode of the macroblock this fragment belongs to.
    mb_mode As Byte         ' :3
    ' The index of the associated border information for fragments which lie partially outside the
    ' displayable region. For fragments completely inside or outside this region, this is -1. Note that the C
    ' standard requires an explicit signed keyword for bitfield types, since some compilers may treat them as
    ' unsigned without it.
    borderi As Integer      ' signed int borderi:5
    ' The prediction-corrected DC component. Note that the C standard requires an explicit signed keyword for
    ' bitfield types, since some compilers may treat them as unsigned without it.
    dc As Long              ' signed int dc:16: (VB) the code that stores it keeps it in 16 bits
End Type

' A description of each fragment plane.
Public Type oc_fragment_plane
    ' The number of fragments in the horizontal direction.
    nhfrags As Long
    ' The number of fragments in the vertical direction.
    nvfrags As Long
    ' The offset of the first fragment in the plane.
    froffset As Long
    ' The total number of fragments in the plane.
    nfrags As Long
    ' The number of super blocks in the horizontal direction.
    nhsbs As Long
    ' The number of super blocks in the vertical direction.
    nvsbs As Long
    ' The offset of the first super block in the plane.
    sboffset As Long
    ' The total number of super blocks in the plane.
    nsbs As Long
End Type

' State information common to both the encoder and decoder.
Public Type oc_theora_state
    ' The stream information.
    info As th_info
    ' Table for shared data used by accelerated functions: opt_data.dct_fzig_zag (VB: always OC_FZIG_ZAG)
    ' CPU flags to detect the presence of extended instruction sets.
    cpu_flags As Long
    ' The fragment plane descriptions.
    fplanes(0 To 2) As oc_fragment_plane
    ' The list of fragments, indexed in image order.
    frags() As oc_fragment
    ' The the offset into the reference frame buffer to the upper-left pixel of each fragment.
    frag_buf_offs() As Long
    ' The motion vector for each fragment.
    frag_mvs() As Long          ' oc_mv
    ' The total number of fragments in a single frame.
    nfrags As Long
    ' The list of super block maps, indexed in image order.
    sb_maps() As Long           ' oc_sb_map: (VB) sb_maps(sbi * 16 + quadi * 4 + bi)
    ' The list of super block flags, indexed in image order.
    sb_flags() As oc_sb_flags
    ' The total number of super blocks in a single frame.
    nsbs As Long
    ' The fragments from each color plane that belong to each macro block. Fragments are stored in image
    ' order (left to right then top to bottom). When chroma components are decimated, the extra fragments
    ' have an index of -1.
    mb_maps() As Long           ' oc_mb_map: (VB) mb_maps(mbi * 12 + pli * 4 + bi)
    ' The list of macro block modes. A negative number indicates the macro block lies entirely outside the
    ' coded frame.
    mb_modes() As Long          ' signed char
    ' The number of macro blocks in the X direction.
    nhmbs As Long
    ' The number of macro blocks in the Y direction.
    nvmbs As Long
    ' The total number of macro blocks.
    nmbs As Long
    ' The list of coded fragments, in coded order. Uncoded fragments are stored in reverse order from the end
    ' of the list.
    coded_fragis() As Long
    ' The number of coded fragments in each plane.
    ncoded_fragis(0 To 2) As Long
    ' The total number of coded fragments.
    ntotal_coded_fragis As Long
    ' The actual buffers used for the reference frames.
    ref_frame_bufs(0 To 5) As th_ycbcr_buffer
    ' The index of the buffers being used for each OC_FRAME_* reference frame.
    ref_frame_idx(0 To 5) As Long
    ' The storage for the reference frame buffers. This is just ref_frame_bufs[ref_frame_idx[i]][0].data, but
    ' is cached here for faster look-up.
    ref_frame_data(0 To 5) As Long  ' unsigned char *: (VB) offsets into ref_frame_handle, 0 = NULL
    ' The handle used to allocate the reference frame buffers.
    ref_frame_handle() As Byte
    ' The strides for each plane in the reference frames.
    ref_ystride(0 To 2) As Long
    ' The number of unique border patterns.
    nborders As Long
    ' The unique border patterns for all border fragments. The borderi field of fragments which straddle the
    ' border indexes this list.
    borders(0 To 15) As oc_border_info
    ' The frame number of the last keyframe.
    keyframe_num As Double      ' ogg_int64_t
    ' The frame number of the current frame.
    curframe_num As Double      ' ogg_int64_t
    ' The granpos of the current frame.
    granpos As Double           ' ogg_int64_t
    ' The type of the current frame.
    frame_type As Long          ' signed char
    ' The bias to add to the frame count when computing granule positions.
    granpos_bias As Long        ' unsigned char
    ' The number of quality indices used in the current frame.
    nqis As Long                ' unsigned char
    ' The quality indices of the current frame.
    qis(0 To 2) As Long         ' unsigned char
    ' The dequantization tables, stored in zig-zag order, and indexed by qi, pli, qti, and zzi.
    dequant_tables() As Long    ' ogg_uint16_t *dequant_tables[64][3][2]: (VB) offsets into dequant_table_data,
                                ' dequant_tables((qi * 3 + pli) * 2 + qti)
    dequant_table_data() As Long    ' OC_ALIGN16(oc_quant_table dequant_table_data[64][3][2]): (VB) 64 * 3 * 2
                                ' tables of 64, then 128 zeroes read by the out-of-range zero runs of bogus
                                ' streams (the C reads the memory after the table there)
    ' Loop filter strength parameters.
    loop_filter_limits(0 To 63) As Byte
    vb_ref_frame_sz As Long     ' (VB) the reference frames' bytes in ref_frame_handle (what follows is the
                                ' decoder's post-processing buffer)
End Type

' The function used to fill in the chroma plane motion vectors for a macro block when 4 different motion
' vectors are specified in the luma plane. This version is for use with chroma decimated in the X and Y
' directions (4:2:0).
'  _cbmvs: The chroma block-level motion vectors to fill in.
'  _lbmvs: The luma block-level motion vectors.
Private Sub oc_set_chroma_mvs00(cbmvs() As Long, lbmvs() As Long)
    Dim dx As Long
    Dim dy As Long
    dx = OC_MV_X(lbmvs(0)) + OC_MV_X(lbmvs(1)) + OC_MV_X(lbmvs(2)) + OC_MV_X(lbmvs(3))
    dy = OC_MV_Y(lbmvs(0)) + OC_MV_Y(lbmvs(1)) + OC_MV_Y(lbmvs(2)) + OC_MV_Y(lbmvs(3))
    cbmvs(0) = OC_MV(OC_DIV_ROUND_POW2(dx, 2, 2), OC_DIV_ROUND_POW2(dy, 2, 2))
End Sub

' The function used to fill in the chroma plane motion vectors for a macro block when 4 different motion
' vectors are specified in the luma plane. This version is for use with chroma decimated in the Y direction.
'  _cbmvs: The chroma block-level motion vectors to fill in.
'  _lbmvs: The luma block-level motion vectors.
Private Sub oc_set_chroma_mvs01(cbmvs() As Long, lbmvs() As Long)
    Dim dx As Long
    Dim dy As Long
    dx = OC_MV_X(lbmvs(0)) + OC_MV_X(lbmvs(2))
    dy = OC_MV_Y(lbmvs(0)) + OC_MV_Y(lbmvs(2))
    cbmvs(0) = OC_MV(OC_DIV_ROUND_POW2(dx, 1, 1), OC_DIV_ROUND_POW2(dy, 1, 1))
    dx = OC_MV_X(lbmvs(1)) + OC_MV_X(lbmvs(3))
    dy = OC_MV_Y(lbmvs(1)) + OC_MV_Y(lbmvs(3))
    cbmvs(1) = OC_MV(OC_DIV_ROUND_POW2(dx, 1, 1), OC_DIV_ROUND_POW2(dy, 1, 1))
End Sub

' The function used to fill in the chroma plane motion vectors for a macro block when 4 different motion
' vectors are specified in the luma plane. This version is for use with chroma decimated in the X direction
' (4:2:2).
'  _cbmvs: The chroma block-level motion vectors to fill in.
'  _lbmvs: The luma block-level motion vectors.
Private Sub oc_set_chroma_mvs10(cbmvs() As Long, lbmvs() As Long)
    Dim dx As Long
    Dim dy As Long
    dx = OC_MV_X(lbmvs(0)) + OC_MV_X(lbmvs(1))
    dy = OC_MV_Y(lbmvs(0)) + OC_MV_Y(lbmvs(1))
    cbmvs(0) = OC_MV(OC_DIV_ROUND_POW2(dx, 1, 1), OC_DIV_ROUND_POW2(dy, 1, 1))
    dx = OC_MV_X(lbmvs(2)) + OC_MV_X(lbmvs(3))
    dy = OC_MV_Y(lbmvs(2)) + OC_MV_Y(lbmvs(3))
    cbmvs(2) = OC_MV(OC_DIV_ROUND_POW2(dx, 1, 1), OC_DIV_ROUND_POW2(dy, 1, 1))
End Sub

' The function used to fill in the chroma plane motion vectors for a macro block when 4 different motion
' vectors are specified in the luma plane. This version is for use with no chroma decimation (4:4:4).
'  _cbmvs: The chroma block-level motion vectors to fill in.
'  _lmbmv: The luma macro-block level motion vector to fill in for use in prediction.
'  _lbmvs: The luma block-level motion vectors.
Private Sub oc_set_chroma_mvs11(cbmvs() As Long, lbmvs() As Long)
    cbmvs(0) = lbmvs(0)
    cbmvs(1) = lbmvs(1)
    cbmvs(2) = lbmvs(2)
    cbmvs(3) = lbmvs(3)
End Sub

' A table of functions used to fill in the chroma plane motion vectors for a macro block when 4 different
' motion vectors are specified in the luma plane (OC_SET_CHROMA_MVS_TABLE[TH_PF_NFORMATS]).
' (VB) (*OC_SET_CHROMA_MVS_TABLE[pixel_fmt])(cbmvs,lbmvs)
Public Sub OC_SET_CHROMA_MVS(ByVal pixel_fmt As Long, cbmvs() As Long, lbmvs() As Long)
    Select Case pixel_fmt
        Case 0: oc_set_chroma_mvs00 cbmvs, lbmvs
        Case 1: oc_set_chroma_mvs01 cbmvs, lbmvs
        Case 2: oc_set_chroma_mvs10 cbmvs, lbmvs
        Case Else: oc_set_chroma_mvs11 cbmvs, lbmvs
    End Select
End Sub

' OC_MV(_x,_y): ((oc_mv)((_x)&0xFF|(_y)*256))
Public Function OC_MV(ByVal x As Long, ByVal y As Long) As Long
    OC_MV = (((x And &HFF&) Or (y * 256)) + 32768 And &HFFFF&) - 32768
End Function

' OC_MV_X(_mv): ((signed char)(_mv))
Public Function OC_MV_X(ByVal mv As Long) As Long
    OC_MV_X = ((mv And &HFF&) Xor &H80&) - &H80&
End Function

' OC_MV_Y(_mv): ((_mv)>>8)
Public Function OC_MV_Y(ByVal mv As Long) As Long
    OC_MV_Y = (mv And &HFFFFFF00) \ 256
End Function

' Determines the reference frame used for a given MB mode (OC_FRAME_FOR_MODE(_x)).
Public Function OC_FRAME_FOR_MODE(ByVal x As Long) As Long
    ' OC_UNIBBLE_TABLE32(OC_FRAME_PREV,OC_FRAME_SELF,OC_FRAME_PREV,OC_FRAME_PREV,OC_FRAME_PREV,
    '  OC_FRAME_GOLD,OC_FRAME_GOLD,OC_FRAME_PREV,(_x))
    Select Case x
        Case 1: OC_FRAME_FOR_MODE = OC_FRAME_SELF
        Case 5, 6: OC_FRAME_FOR_MODE = OC_FRAME_GOLD
        Case Else: OC_FRAME_FOR_MODE = OC_FRAME_PREV
    End Select
End Function

' Returns the fragment index of the top-left block in a macro block. This can be used to test whether or not
' the whole macro block is valid.
'  _sb_map: The super block map.
'  _quadi:  The quadrant number.
'  Return: The index of the fragment of the upper left block in the macro block, or -1 if the block lies
'   outside the coded frame.
Private Function oc_sb_quad_top_left_frag(st As oc_theora_state, ByVal sbi As Long, ByVal quadi As Long) As Long
    ' It so happens that under the Hilbert curve ordering described below, the upper-left block in each macro
    ' block is at index 0, except in macro block 3, where it is at index 2.
    oc_sb_quad_top_left_frag = st.sb_maps(sbi * 16 + quadi * 4 + (quadi And (quadi * 2)))
End Function

' Fills in the mapping from block positions to fragment numbers for a single color plane. This function also
' fills in the "valid" flag of each quadrant in the super block flags.
'  _sb_maps:  The array of super block maps for the color plane. (VB) st.sb_maps from super block sb0
'  _sb_flags: The array of super block flags for the color plane. (VB) st.sb_flags from sb0
'  _frag0:    The index of the first fragment in the plane.
'  _hfrags:   The number of horizontal fragments in a coded frame.
'  _vfrags:   The number of vertical fragments in a coded frame.
Private Sub oc_sb_create_plane_mapping(st As oc_theora_state, ByVal sb0 As Long, ByVal frag0 As Long, _
    ByVal hfrags As Long, ByVal vfrags As Long)
    ' Contains the (macro_block,block) indices for a 4x4 grid of fragments. The pattern is a 4x4 Hilbert
    ' space-filling curve. A Hilbert curve has the nice property that as the curve grows larger, its fractal
    ' dimension approaches 2. The intuition is that nearby blocks in the curve are also close spatially, with
    ' the previous element always an immediate neighbor, so that runs of blocks should be well correlated.
    ' (SB_MAP: OC_SB_MAP in theora_tables_vb, SB_MAP[i][j][k] = OC_SB_MAP(i * 8 + j * 2 + k))
    Dim yfrag As Long
    Dim sbi As Long
    Dim y As Long
    Dim imax As Long
    Dim x As Long
    Dim xfrag As Long
    Dim jmax As Long
    Dim quadi As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    sbi = 0
    yfrag = frag0
    y = 0
    Do
        ' Figure out how many columns of blocks in this super block lie within the image.
        imax = vfrags - y
        If imax > 4 Then
            imax = 4
        ElseIf imax <= 0 Then
            Exit Do
        End If
        x = 0
        Do
            ' Figure out how many rows of blocks in this super block lie within the image.
            jmax = hfrags - x
            If jmax > 4 Then
                jmax = 4
            ElseIf jmax <= 0 Then
                Exit Do
            End If
            ' By default, set all fragment indices to -1.
            For k = 0 To 15
                st.sb_maps((sb0 + sbi) * 16 + k) = -1
            Next
            ' Fill in the fragment map for this super block.
            xfrag = yfrag + x
            For i = 0 To imax - 1
                For j = 0 To jmax - 1
                    st.sb_maps((sb0 + sbi) * 16 + OC_SB_MAP(i * 8 + j * 2) * 4 + OC_SB_MAP(i * 8 + j * 2 + 1)) = xfrag + j
                Next
                xfrag = xfrag + hfrags
            Next
            ' Mark which quadrants of this super block lie within the image.
            For quadi = 0 To 3
                If oc_sb_quad_top_left_frag(st, sb0 + sbi, quadi) >= 0 Then
                    st.sb_flags(sb0 + sbi).quad_valid = st.sb_flags(sb0 + sbi).quad_valid Or OC_PB_POW2(quadi)
                End If
            Next
            x = x + 4
            sbi = sbi + 1
        Loop
        yfrag = yfrag + hfrags * 4
        y = y + 4
    Loop
End Sub

' Fills in the Y plane fragment map for a macro block given the fragment coordinates of its upper-left hand
' corner.
'  _mb_map:    The macro block map to fill. (VB) st.mb_maps of mbi
'  _fplane: The description of the Y plane.
'  _xfrag0: The X location of the upper-left hand fragment in the luma plane.
'  _yfrag0: The Y location of the upper-left hand fragment in the luma plane.
Private Sub oc_mb_fill_ymapping(st As oc_theora_state, ByVal mbi As Long, ByVal xfrag0 As Long, ByVal yfrag0 As Long)
    Dim i As Long
    Dim j As Long
    For i = 0 To 1
        For j = 0 To 1
            st.mb_maps(mbi * 12 + (i * 2 Or j)) = (yfrag0 + i) * st.fplanes(0).nhfrags + xfrag0 + j
        Next
    Next
End Sub

' Fills in the chroma plane fragment maps for a macro block. This version is for use with chroma decimated in
' the X and Y directions (4:2:0).
'  _mb_map:  The macro block map to fill.
'  _fplanes: The descriptions of the fragment planes.
'  _xfrag0:  The X location of the upper-left hand fragment in the luma plane.
'  _yfrag0:  The Y location of the upper-left hand fragment in the luma plane.
Private Sub oc_mb_fill_cmapping00(st As oc_theora_state, ByVal mbi As Long, ByVal xfrag0 As Long, ByVal yfrag0 As Long)
    Dim fragi As Long
    xfrag0 = xfrag0 \ 2
    yfrag0 = yfrag0 \ 2
    fragi = yfrag0 * st.fplanes(1).nhfrags + xfrag0
    st.mb_maps(mbi * 12 + 4) = fragi + st.fplanes(1).froffset
    st.mb_maps(mbi * 12 + 8) = fragi + st.fplanes(2).froffset
End Sub

' Fills in the chroma plane fragment maps for a macro block. This version is for use with chroma decimated in
' the Y direction.
'  _mb_map:  The macro block map to fill.
'  _fplanes: The descriptions of the fragment planes.
'  _xfrag0:  The X location of the upper-left hand fragment in the luma plane.
'  _yfrag0:  The Y location of the upper-left hand fragment in the luma plane.
Private Sub oc_mb_fill_cmapping01(st As oc_theora_state, ByVal mbi As Long, ByVal xfrag0 As Long, ByVal yfrag0 As Long)
    Dim fragi As Long
    Dim j As Long
    yfrag0 = yfrag0 \ 2
    fragi = yfrag0 * st.fplanes(1).nhfrags + xfrag0
    For j = 0 To 1
        st.mb_maps(mbi * 12 + 4 + j) = fragi + st.fplanes(1).froffset
        st.mb_maps(mbi * 12 + 8 + j) = fragi + st.fplanes(2).froffset
        fragi = fragi + 1
    Next
End Sub

' Fills in the chroma plane fragment maps for a macro block. This version is for use with chroma decimated in
' the X direction (4:2:2).
'  _mb_map:  The macro block map to fill.
'  _fplanes: The descriptions of the fragment planes.
'  _xfrag0:  The X location of the upper-left hand fragment in the luma plane.
'  _yfrag0:  The Y location of the upper-left hand fragment in the luma plane.
Private Sub oc_mb_fill_cmapping10(st As oc_theora_state, ByVal mbi As Long, ByVal xfrag0 As Long, ByVal yfrag0 As Long)
    Dim fragi As Long
    Dim i As Long
    xfrag0 = xfrag0 \ 2
    fragi = yfrag0 * st.fplanes(1).nhfrags + xfrag0
    For i = 0 To 1
        st.mb_maps(mbi * 12 + 4 + i * 2) = fragi + st.fplanes(1).froffset
        st.mb_maps(mbi * 12 + 8 + i * 2) = fragi + st.fplanes(2).froffset
        fragi = fragi + st.fplanes(1).nhfrags
    Next
End Sub

' Fills in the chroma plane fragment maps for a macro block. This version is for use with no chroma
' decimation (4:4:4). This uses the already filled-in luma plane values.
'  _mb_map:  The macro block map to fill.
'  _fplanes: The descriptions of the fragment planes.
'  _xfrag0:  The X location of the upper-left hand fragment in the luma plane.
'  _yfrag0:  The Y location of the upper-left hand fragment in the luma plane.
Private Sub oc_mb_fill_cmapping11(st As oc_theora_state, ByVal mbi As Long)
    Dim k As Long
    For k = 0 To 3
        st.mb_maps(mbi * 12 + 4 + k) = st.mb_maps(mbi * 12 + k) + st.fplanes(1).froffset
        st.mb_maps(mbi * 12 + 8 + k) = st.mb_maps(mbi * 12 + k) + st.fplanes(2).froffset
    Next
End Sub

' Fills in the mapping from macro blocks to their corresponding fragment numbers in each plane.
'  _mb_maps:   The list of macro block maps.
'  _mb_modes:  The list of macro block modes; macro blocks completely outside the coded region are marked
'               invalid.
'  _fplanes:   The descriptions of the fragment planes.
'  _pixel_fmt: The chroma decimation type.
' (VB) mb_fill_cmapping=OC_MB_FILL_CMAPPING_TABLE[_pixel_fmt]: a Select Case on pixel_fmt
Private Sub oc_mb_create_mapping(st As oc_theora_state, ByVal pixel_fmt As Long)
    Dim sbi As Long
    Dim y As Long
    Dim x As Long
    Dim ymb As Long
    Dim xmb As Long
    Dim mbi As Long
    Dim mbx As Long
    Dim mby As Long
    Dim k As Long
    ' Loop through the luma plane super blocks.
    sbi = 0
    For y = 0 To st.fplanes(0).nvfrags - 1 Step 4
        For x = 0 To st.fplanes(0).nhfrags - 1 Step 4
            ' Loop through the macro blocks in each super block in display order.
            For ymb = 0 To 1
                For xmb = 0 To 1
                    mbi = sbi * 4 Or OC_MB_MAP(ymb * 2 + xmb)
                    mbx = x Or xmb * 2
                    mby = y Or ymb * 2
                    ' Initialize fragment indices to -1.
                    For k = 0 To 11
                        st.mb_maps(mbi * 12 + k) = -1
                    Next
                    ' Make sure this macro block is within the encoded region.
                    If mbx >= st.fplanes(0).nhfrags Or mby >= st.fplanes(0).nvfrags Then
                        st.mb_modes(mbi) = OC_MODE_INVALID
                    Else
                        ' Fill in the fragment indices for the luma plane.
                        oc_mb_fill_ymapping st, mbi, mbx, mby
                        ' Fill in the fragment indices for the chroma planes.
                        Select Case pixel_fmt
                            Case 0: oc_mb_fill_cmapping00 st, mbi, mbx, mby
                            Case 1: oc_mb_fill_cmapping01 st, mbi, mbx, mby
                            Case 2: oc_mb_fill_cmapping10 st, mbi, mbx, mby
                            Case Else: oc_mb_fill_cmapping11 st, mbi
                        End Select
                    End If
                Next
            Next
            sbi = sbi + 1
        Next
    Next
End Sub

' Marks the fragments which fall all or partially outside the displayable region of the frame.
'  _state: The Theora state containing the fragments to be marked.
Private Sub oc_state_border_init(st As oc_theora_state)
    Dim frag As Long            ' oc_fragment *: (VB) an index into st.frags
    Dim yfrag_end As Long
    Dim xfrag_end As Long
    Dim crop_x0 As Long
    Dim crop_y0 As Long
    Dim crop_xf As Long
    Dim crop_yf As Long
    Dim pli As Long
    Dim y As Long
    Dim x As Long
    Dim mask_lo As Long         ' ogg_uint64_t mask: (VB) two halves
    Dim mask_hi As Long
    Dim npixels As Long
    Dim i As Long
    Dim j As Long
    Dim bit As Long
    ' The method we use here is slow, but the code is dead simple and handles all the special cases easily.
    ' We only ever need to do it once.
    ' Loop through the fragments, marking those completely outside the displayable region and constructing a
    ' border mask for those that straddle the border.
    st.nborders = 0
    frag = 0
    yfrag_end = 0
    For pli = 0 To 2
        ' Set up the cropping rectangle for this plane.
        crop_x0 = st.info.pic_x
        crop_xf = st.info.pic_x + st.info.pic_width
        crop_y0 = st.info.pic_y
        crop_yf = st.info.pic_y + st.info.pic_height
        If pli > 0 Then
            If (st.info.pixel_fmt And 1) = 0 Then
                crop_x0 = crop_x0 \ 2
                crop_xf = (crop_xf + 1) \ 2
            End If
            If (st.info.pixel_fmt And 2) = 0 Then
                crop_y0 = crop_y0 \ 2
                crop_yf = (crop_yf + 1) \ 2
            End If
        End If
        y = 0
        yfrag_end = yfrag_end + st.fplanes(pli).nfrags
        Do While frag < yfrag_end
            x = 0
            xfrag_end = frag + st.fplanes(pli).nhfrags
            Do While frag < xfrag_end
                ' First check to see if this fragment is completely outside the displayable region.
                ' Note the special checks for an empty cropping rectangle. This guarantees that if we count a
                ' fragment as straddling the border below, at least one pixel in the fragment will be inside
                ' the displayable region.
                If x + 8 <= crop_x0 Or crop_xf <= x Or y + 8 <= crop_y0 Or crop_yf <= y Or crop_x0 >= crop_xf Or _
                    crop_y0 >= crop_yf Then
                    st.frags(frag).invalid = 1
                ' Otherwise, check to see if it straddles the border.
                ElseIf (x < crop_x0 And crop_x0 < x + 8) Or (x < crop_xf And crop_xf < x + 8) Or _
                    (y < crop_y0 And crop_y0 < y + 8) Or (y < crop_yf And crop_yf < y + 8) Then
                    mask_lo = 0
                    mask_hi = 0
                    npixels = 0
                    For i = 0 To 7
                        For j = 0 To 7
                            If x + j >= crop_x0 And x + j < crop_xf And y + i >= crop_y0 And y + i < crop_yf Then
                                ' mask|=(ogg_uint64_t)1<<(i<<3|j)
                                bit = i * 8 Or j
                                If bit < 31 Then
                                    mask_lo = mask_lo Or OC_PB_POW2(bit)
                                ElseIf bit = 31 Then
                                    mask_lo = mask_lo Or &H80000000
                                ElseIf bit < 63 Then
                                    mask_hi = mask_hi Or OC_PB_POW2(bit - 32)
                                Else
                                    mask_hi = mask_hi Or &H80000000
                                End If
                                npixels = npixels + 1
                            End If
                        Next
                    Next
                    ' Search the fragment array for border info with the same pattern. In general, there will be
                    ' at most 8 different patterns (per plane).
                    i = 0
                    Do
                        If i >= st.nborders Then
                            st.nborders = st.nborders + 1
                            st.borders(i).mask_lo = mask_lo
                            st.borders(i).mask_hi = mask_hi
                            st.borders(i).npixels = npixels
                        ElseIf st.borders(i).mask_lo <> mask_lo Or st.borders(i).mask_hi <> mask_hi Then
                            GoTo continue_
                        End If
                        st.frags(frag).borderi = i
                        Exit Do
continue_:
                        i = i + 1
                    Loop
                Else
                    st.frags(frag).borderi = -1
                End If
                frag = frag + 1
                x = x + 8
            Loop
            y = y + 8
        Loop
    Next
End Sub

Private Function oc_state_frarray_init(st As oc_theora_state) As Long
    Dim yhfrags As Long
    Dim yvfrags As Long
    Dim chfrags As Long
    Dim cvfrags As Long
    Dim yfrags As Long
    Dim cfrags As Long
    Dim nfrags As Long
    Dim yhsbs As Long
    Dim yvsbs As Long
    Dim chsbs As Long
    Dim cvsbs As Long
    Dim ysbs As Long
    Dim csbs As Long
    Dim nsbs As Long
    Dim nmbs As Long
    Dim hdec As Long
    Dim vdec As Long
    Dim pli As Long
    ' Figure out the number of fragments in each plane.
    ' These parameters have already been validated to be multiples of 16.
    yhfrags = st.info.frame_width \ 8
    yvfrags = st.info.frame_height \ 8
    If (st.info.pixel_fmt And 1) = 0 Then hdec = 1 Else hdec = 0
    If (st.info.pixel_fmt And 2) = 0 Then vdec = 1 Else vdec = 0
    chfrags = (yhfrags + hdec) \ OC_PB_POW2(hdec)
    cvfrags = (yvfrags + vdec) \ OC_PB_POW2(vdec)
    yfrags = yhfrags * yvfrags
    cfrags = chfrags * cvfrags
    nfrags = yfrags + 2 * cfrags
    ' Figure out the number of super blocks in each plane.
    yhsbs = (yhfrags + 3) \ 4
    yvsbs = (yvfrags + 3) \ 4
    chsbs = (chfrags + 3) \ 4
    cvsbs = (cvfrags + 3) \ 4
    ysbs = yhsbs * yvsbs
    csbs = chsbs * cvsbs
    nsbs = ysbs + 2 * csbs
    nmbs = ysbs * 4
    ' Check for overflow. We support the ridiculous upper limits of the specification (1048560 by 1048560, or
    ' 3 TB frames) if the target architecture has 64-bit pointers, but for those with 32-bit pointers (or
    ' smaller!) we have to check. If the caller wants to prevent denial-of-service by imposing a more
    ' reasonable upper limit on the size of attempted allocations, they must do so themselves; we have no
    ' platform independent way to determine how much system memory there is nor an application-independent
    ' way to decide what a "reasonable" allocation is.
    ' (VB) the products in Double: a Long would overflow before the C's wrapped value could be compared
    If CDbl(yhfrags) * yvfrags + 2# * CDbl(chfrags) * cvfrags > 2147483647# / 16 Or _
        CDbl(yhsbs) * yvsbs * 4 + 2# * CDbl(chsbs) * cvsbs > 2147483647# / 16 Then
        oc_state_frarray_init = TH_EIMPL
        Exit Function
    End If
    ' Initialize the fragment array.
    st.fplanes(0).nhfrags = yhfrags
    st.fplanes(0).nvfrags = yvfrags
    st.fplanes(0).froffset = 0
    st.fplanes(0).nfrags = yfrags
    st.fplanes(0).nhsbs = yhsbs
    st.fplanes(0).nvsbs = yvsbs
    st.fplanes(0).sboffset = 0
    st.fplanes(0).nsbs = ysbs
    st.fplanes(1).nhfrags = chfrags
    st.fplanes(2).nhfrags = chfrags
    st.fplanes(1).nvfrags = cvfrags
    st.fplanes(2).nvfrags = cvfrags
    st.fplanes(1).froffset = yfrags
    st.fplanes(2).froffset = yfrags + cfrags
    st.fplanes(1).nfrags = cfrags
    st.fplanes(2).nfrags = cfrags
    st.fplanes(1).nhsbs = chsbs
    st.fplanes(2).nhsbs = chsbs
    st.fplanes(1).nvsbs = cvsbs
    st.fplanes(2).nvsbs = cvsbs
    st.fplanes(1).sboffset = ysbs
    st.fplanes(2).sboffset = ysbs + csbs
    st.fplanes(1).nsbs = csbs
    st.fplanes(2).nsbs = csbs
    st.nfrags = nfrags
    ReDim st.frags(0 To nfrags - 1)
    ReDim st.frag_mvs(0 To nfrags - 1)
    st.nsbs = nsbs
    ReDim st.sb_maps(0 To nsbs * 16 - 1)
    ReDim st.sb_flags(0 To nsbs - 1)
    st.nhmbs = yhsbs * 2
    st.nvmbs = yvsbs * 2
    st.nmbs = nmbs
    ReDim st.mb_maps(0 To nmbs * 12 - 1)
    ReDim st.mb_modes(0 To nmbs - 1)
    ReDim st.coded_fragis(0 To nfrags - 1)
    ' Create the mapping from super blocks to fragments.
    For pli = 0 To 2
        oc_sb_create_plane_mapping st, st.fplanes(pli).sboffset, st.fplanes(pli).froffset, _
            st.fplanes(pli).nhfrags, st.fplanes(pli).nvfrags
    Next
    ' Create the mapping from macro blocks to fragments.
    oc_mb_create_mapping st, st.info.pixel_fmt
    ' Initialize the invalid and borderi fields of each fragment.
    oc_state_border_init st
    oc_state_frarray_init = 0
End Function

Private Sub oc_state_frarray_clear(st As oc_theora_state)
    Erase st.coded_fragis
    Erase st.mb_modes
    Erase st.mb_maps
    Erase st.sb_flags
    Erase st.sb_maps
    Erase st.frag_mvs
    Erase st.frags
End Sub

' Initializes the buffers used for reconstructed frames. These buffers are padded with 16 extra pixels on each
' side, to allow unrestricted motion vectors without special casing the boundary. If chroma is decimated in
' either direction, the padding is reduced by a factor of 2 on the appropriate sides.
'  _nrefs: The number of reference buffers to init; must be in the range 3...6.
Private Function oc_state_ref_bufs_init(st As oc_theora_state, ByVal nrefs As Long) As Long
    Dim ref_frame_data As Long      ' unsigned char *: (VB) an offset into st.ref_frame_handle
    Dim ref_frame_data_sz As Long
    Dim ref_frame_sz As Long
    Dim yplane_sz As Long
    Dim cplane_sz As Long
    Dim yhstride As Long
    Dim yheight As Long
    Dim chstride As Long
    Dim cheight As Long
    Dim align As Long
    Dim yoffset As Long
    Dim coffset As Long
    Dim fragi As Long
    Dim hdec As Long
    Dim vdec As Long
    Dim rfi As Long
    Dim pli As Long
    Dim vpix As Long
    Dim stride As Long
    Dim vfragi_end As Long
    Dim nhfrags As Long
    Dim hfragi_end As Long
    Dim hpix As Long
    If nrefs < 3 Or nrefs > 6 Then oc_state_ref_bufs_init = TH_EINVAL: Exit Function
    ' Compute the image buffer parameters for each plane.
    If (st.info.pixel_fmt And 1) = 0 Then hdec = 1 Else hdec = 0
    If (st.info.pixel_fmt And 2) = 0 Then vdec = 1 Else vdec = 0
    yhstride = st.info.frame_width + 2 * OC_UMV_PADDING
    yheight = st.info.frame_height + 2 * OC_UMV_PADDING
    ' Require 16-byte aligned rows in the chroma planes.
    chstride = ((yhstride \ OC_PB_POW2(hdec)) + 15) And Not 15
    cheight = yheight \ OC_PB_POW2(vdec)
    yplane_sz = yhstride * yheight
    cplane_sz = chstride * cheight
    yoffset = OC_UMV_PADDING + OC_UMV_PADDING * yhstride
    coffset = (OC_UMV_PADDING \ OC_PB_POW2(hdec)) + (OC_UMV_PADDING \ OC_PB_POW2(vdec)) * chstride
    ' Although we guarantee the rows of the chroma planes are a multiple of 16 bytes, the initial padding on
    ' the first row may only be 8 bytes. Compute the offset needed to the actual image data to a multiple of
    ' 16.
    align = -coffset And 15
    ref_frame_sz = yplane_sz + 2 * cplane_sz + 16
    ref_frame_data_sz = nrefs * ref_frame_sz
    ' Check for overflow. The same caveats apply as for oc_state_frarray_init(). (VB) checked there
    ' ref_frame_data=oc_aligned_malloc(ref_frame_data_sz,16): (VB) offsets from 0 of ref_frame_handle
    ReDim st.ref_frame_handle(0 To ref_frame_data_sz - 1)
    st.vb_ref_frame_sz = ref_frame_data_sz
    ReDim st.frag_buf_offs(0 To st.nfrags - 1)
    ' Set up the width, height and stride for the image buffers.
    st.ref_frame_bufs(0).p(0).width = st.info.frame_width
    st.ref_frame_bufs(0).p(0).height = st.info.frame_height
    st.ref_frame_bufs(0).p(0).stride = yhstride
    st.ref_frame_bufs(0).p(1).width = st.info.frame_width \ OC_PB_POW2(hdec)
    st.ref_frame_bufs(0).p(2).width = st.ref_frame_bufs(0).p(1).width
    st.ref_frame_bufs(0).p(1).height = st.info.frame_height \ OC_PB_POW2(vdec)
    st.ref_frame_bufs(0).p(2).height = st.ref_frame_bufs(0).p(1).height
    st.ref_frame_bufs(0).p(1).stride = chstride
    st.ref_frame_bufs(0).p(2).stride = chstride
    For rfi = 1 To nrefs - 1
        st.ref_frame_bufs(rfi) = st.ref_frame_bufs(0)
    Next
    ' _state->ref_frame_handle=ref_frame_data
    ref_frame_data = 0
    ' Set up the data pointers for the image buffers.
    For rfi = 0 To nrefs - 1
        st.ref_frame_bufs(rfi).p(0).data = ref_frame_data + yoffset
        ref_frame_data = ref_frame_data + yplane_sz + align
        st.ref_frame_bufs(rfi).p(1).data = ref_frame_data + coffset
        ref_frame_data = ref_frame_data + cplane_sz
        st.ref_frame_bufs(rfi).p(2).data = ref_frame_data + coffset
        ref_frame_data = ref_frame_data + cplane_sz + (16 - align)
        ' Flip the buffer upside down. This allows us to decode Theora's bottom-up frames in their natural
        ' order, yet return a top-down buffer with a positive stride to the user.
        oc_ycbcr_buffer_flip st.ref_frame_bufs(rfi), st.ref_frame_bufs(rfi)
    Next
    st.ref_ystride(0) = -yhstride
    st.ref_ystride(1) = -chstride
    st.ref_ystride(2) = -chstride
    ' Initialize the fragment buffer offsets.
    ref_frame_data = st.ref_frame_bufs(0).p(0).data
    fragi = 0
    For pli = 0 To 2
        vpix = st.ref_frame_bufs(0).p(pli).data
        vfragi_end = st.fplanes(pli).froffset + st.fplanes(pli).nfrags
        nhfrags = st.fplanes(pli).nhfrags
        stride = st.ref_frame_bufs(0).p(pli).stride
        Do While fragi < vfragi_end
            hpix = vpix
            hfragi_end = fragi + nhfrags
            Do While fragi < hfragi_end
                st.frag_buf_offs(fragi) = hpix - ref_frame_data
                hpix = hpix + 8
                fragi = fragi + 1
            Loop
            vpix = vpix + stride * 8
        Loop
    Next
    ' Initialize the reference frame pointers and indices.
    st.ref_frame_idx(OC_FRAME_GOLD) = -1
    st.ref_frame_idx(OC_FRAME_PREV) = -1
    st.ref_frame_idx(OC_FRAME_GOLD_ORIG) = -1
    st.ref_frame_idx(OC_FRAME_PREV_ORIG) = -1
    st.ref_frame_idx(OC_FRAME_SELF) = -1
    st.ref_frame_idx(OC_FRAME_IO) = -1
    st.ref_frame_data(OC_FRAME_GOLD) = 0
    st.ref_frame_data(OC_FRAME_PREV) = 0
    st.ref_frame_data(OC_FRAME_GOLD_ORIG) = 0
    st.ref_frame_data(OC_FRAME_PREV_ORIG) = 0
    st.ref_frame_data(OC_FRAME_SELF) = 0
    st.ref_frame_data(OC_FRAME_IO) = 0
    oc_state_ref_bufs_init = 0
End Function

Private Sub oc_state_ref_bufs_clear(st As oc_theora_state)
    Erase st.frag_buf_offs
    Erase st.ref_frame_handle
End Sub

Public Sub oc_state_accel_init_c(st As oc_theora_state)
    st.cpu_flags = 0
    ' _state->opt_data.dct_fzig_zag=OC_FZIG_ZAG: (VB) the procedures use OC_FZIG_ZAG
End Sub

Public Function oc_state_init(st As oc_theora_state, info As th_info, ByVal nrefs As Long) As Long
    Dim ret As Long
    Dim blank As oc_theora_state
    ' First validate the parameters.
    ' The width and height of the encoded frame must be multiples of 16. They must also, when divided by 16,
    ' fit into a 16-bit unsigned integer. The displayable frame offset coordinates must fit into an 8-bit
    ' unsigned integer. Note that the offset Y in the API is specified on the opposite side from how it is
    ' specified in the bitstream, because the Y axis is flipped in the bitstream. The displayable frame must
    ' fit inside the encoded frame. The color space must be one known by the encoder. The framerate ratio
    ' must not contain a zero value.
    If (info.frame_width And &HF) <> 0 Or (info.frame_height And &HF) <> 0 Or _
        info.frame_width <= 0 Or info.frame_width >= &H100000 Or _
        info.frame_height <= 0 Or info.frame_height >= &H100000 Then
        oc_state_init = TH_EINVAL
        Exit Function
    End If
    If info.pic_x + info.pic_width > info.frame_width Or info.pic_y + info.pic_height > info.frame_height Or _
        info.pic_x > 255 Or info.frame_height - info.pic_height - info.pic_y > 255 Or _
        info.colorspace < 0 Or info.colorspace >= TH_CS_NSPACES Or _
        info.pixel_fmt < 0 Or info.pixel_fmt >= TH_PF_NFORMATS Or _
        info.fps_numerator = 0 Or info.fps_denominator = 0 Then
        ' Note: the following <0 comparisons may generate spurious warnings on platforms where enums are
        ' unsigned. (the fps are ogg_uint32_t: <1 is 0)
        oc_state_init = TH_EINVAL
        Exit Function
    End If
    ' memset(_state,0,sizeof(*_state))
    st = blank
    st.info = info
    ' Invert the sense of pic_y to match Theora's right-handed coordinate system.
    st.info.pic_y = info.frame_height - info.pic_height - info.pic_y
    st.frame_type = OC_UNKWN_FRAME
    ' (VB) the C's arrays inside the struct
    ReDim st.dequant_tables(0 To 64 * 3 * 2 - 1)
    ReDim st.dequant_table_data(0 To 64 * 3 * 2 * 64 + 127)
    oc_state_accel_init_c st
    ret = oc_state_frarray_init(st)
    If ret >= 0 Then ret = oc_state_ref_bufs_init(st, nrefs)
    If ret < 0 Then
        oc_state_frarray_clear st
        oc_state_init = ret
        Exit Function
    End If
    ' If the keyframe_granule_shift is out of range, use the maximum allowable value.
    If info.keyframe_granule_shift < 0 Or info.keyframe_granule_shift > 31 Then
        st.info.keyframe_granule_shift = 31
    End If
    st.keyframe_num = 0
    st.curframe_num = -1
    ' 3.2.0 streams mark the frame index instead of the frame count. This was changed with stream version
    ' 3.2.1 to conform to other Ogg codecs. We add an extra bias when computing granule positions for new
    ' streams.
    st.granpos_bias = TH_VERSION_CHECK(info, 3, 2, 1)
    oc_state_init = 0
End Function

Public Sub oc_state_clear(st As oc_theora_state)
    oc_state_ref_bufs_clear st
    oc_state_frarray_clear st
End Sub

' Duplicates the pixels on the border of the image plane out into the surrounding padding for use by
' unrestricted motion vectors. This function only adds the left and right borders, and only for the fragment
' rows specified.
'  _refi: The index of the reference buffer to pad.
'  _pli:  The color plane.
'  _y0:   The Y coordinate of the first row to pad.
'  _yend: The Y coordinate of the row to stop padding at.
' (VB) pix: st.ref_frame_handle
Public Sub oc_state_borders_fill_rows(st As oc_theora_state, pix() As Byte, ByVal refi As Long, ByVal pli As Long, _
    ByVal y0 As Long, ByVal yend As Long)
    Dim apix As Long
    Dim bpix As Long
    Dim epix As Long
    Dim stride As Long
    Dim hpadding As Long
    Dim k As Long
    Dim a As Byte
    Dim b As Byte
    hpadding = OC_UMV_PADDING
    If pli <> 0 Then
        If (st.info.pixel_fmt And 1) = 0 Then hpadding = OC_UMV_PADDING \ 2
    End If
    stride = st.ref_frame_bufs(refi).p(pli).stride
    apix = st.ref_frame_bufs(refi).p(pli).data + y0 * stride
    bpix = apix + st.ref_frame_bufs(refi).p(pli).width - 1
    epix = st.ref_frame_bufs(refi).p(pli).data + yend * stride
    ' Note the use of != instead of <, which allows the stride to be negative.
    Do While apix <> epix
        ' memset(apix-hpadding,apix[0],hpadding); memset(bpix+1,bpix[0],hpadding);
        a = pix(apix)
        b = pix(bpix)
        For k = 1 To hpadding
            pix(apix - k) = a
            pix(bpix + k) = b
        Next
        apix = apix + stride
        bpix = bpix + stride
    Loop
End Sub

' Duplicates the pixels on the border of the image plane out into the surrounding padding for use by
' unrestricted motion vectors. This function only adds the top and bottom borders, and must be called after
' the left and right borders are added.
'  _refi:      The index of the reference buffer to pad.
'  _pli:       The color plane.
Public Sub oc_state_borders_fill_caps(st As oc_theora_state, ByVal refi As Long, ByVal pli As Long)
    Dim apix As Long
    Dim bpix As Long
    Dim epix As Long
    Dim stride As Long
    Dim hpadding As Long
    Dim vpadding As Long
    Dim fullw As Long
    hpadding = OC_UMV_PADDING
    vpadding = OC_UMV_PADDING
    If pli <> 0 Then
        If (st.info.pixel_fmt And 1) = 0 Then hpadding = OC_UMV_PADDING \ 2
        If (st.info.pixel_fmt And 2) = 0 Then vpadding = OC_UMV_PADDING \ 2
    End If
    stride = st.ref_frame_bufs(refi).p(pli).stride
    fullw = st.ref_frame_bufs(refi).p(pli).width + hpadding * 2
    apix = st.ref_frame_bufs(refi).p(pli).data - hpadding
    bpix = st.ref_frame_bufs(refi).p(pli).data + (st.ref_frame_bufs(refi).p(pli).height - 1) * stride - hpadding
    epix = apix - stride * vpadding
    Do While apix <> epix
        CopyMemory st.ref_frame_handle(apix - stride), st.ref_frame_handle(apix), fullw
        CopyMemory st.ref_frame_handle(bpix + stride), st.ref_frame_handle(bpix), fullw
        apix = apix - stride
        bpix = bpix + stride
    Loop
End Sub

' Duplicates the pixels on the border of the given reference image out into the surrounding padding for use
' by unrestricted motion vectors.
'  _state: The context containing the reference buffers.
'  _refi:  The index of the reference buffer to pad.
Public Sub oc_state_borders_fill(st As oc_theora_state, ByVal refi As Long)
    Dim pli As Long
    For pli = 0 To 2
        oc_state_borders_fill_rows st, st.ref_frame_handle, refi, pli, 0, st.ref_frame_bufs(refi).p(pli).height
        oc_state_borders_fill_caps st, refi, pli
    Next
End Sub

' Determines the offsets in an image buffer to use for motion compensation.
'  _state:   The Theora state the offsets are to be computed with.
'  _offsets: Returns the offset for the buffer(s). _offsets[0] is always set. _offsets[1] is set if the
'             motion vector has non-zero fractional components.
'  _pli:     The color plane index.
'  _mv:      The motion vector.
'  Return: The number of offsets returned: 1 or 2.
Public Function oc_state_get_mv_offsets(st As oc_theora_state, offsets() As Long, ByVal pli As Long, _
    ByVal mv As Long) As Long
    ' Here is a brief description of how Theora handles motion vectors: Motion vector components are
    ' specified to half-pixel accuracy in undecimated directions of each plane, and quarter-pixel accuracy in
    ' decimated directions. Integer parts are extracted by dividing (not shifting) by the appropriate amount,
    ' with truncation towards zero. These integer values are used to calculate the first offset.
    ' If either of the fractional parts are non-zero, then a second offset is computed. No third or fourth
    ' offsets are computed, even if both components have non-zero fractional parts. The second offset is
    ' computed by dividing (not shifting) by the appropriate amount, always truncating _away_ from zero.
    ' (#if 0: a version without tables, not ported.) Using tables simplifies the code, and there's enough
    ' arithmetic to hide the latencies of the memory references. (OC_MVMAP[2][64], OC_MVMAP2[2][64]:
    ' theora_tables_vb)
    Dim ystride As Long
    Dim qpx As Long
    Dim qpy As Long
    Dim mx As Long
    Dim my As Long
    Dim mx2 As Long
    Dim my2 As Long
    Dim offs As Long
    Dim dx As Long
    Dim dy As Long
    ystride = st.ref_ystride(pli)
    qpy = 0
    qpx = 0
    If pli <> 0 Then
        If (st.info.pixel_fmt And 2) = 0 Then qpy = 1
        If (st.info.pixel_fmt And 1) = 0 Then qpx = 1
    End If
    dx = ((mv And &HFF&) Xor &H80&) - &H80&     ' OC_MV_X(_mv)
    dy = (mv And &HFFFFFF00) \ 256              ' OC_MV_Y(_mv)
    my = OC_MVMAP(qpy * 64 + dy + 31)
    my2 = OC_MVMAP2(qpy * 64 + dy + 31)
    mx = OC_MVMAP(qpx * 64 + dx + 31)
    mx2 = OC_MVMAP2(qpx * 64 + dx + 31)
    offs = my * ystride + mx
    If mx2 <> 0 Or my2 <> 0 Then
        offsets(1) = offs + my2 * ystride + mx2
        offsets(0) = offs
        oc_state_get_mv_offsets = 2
        Exit Function
    End If
    offsets(0) = offs
    oc_state_get_mv_offsets = 1
End Function

' (VB) dct_coeffs: the pipeline's 128 coefficients (the reconstruction buffer is its second half)
Public Sub oc_state_frag_recon_c(st As oc_theora_state, ByVal fragi As Long, ByVal pli As Long, _
    dct_coeffs() As Long, ByVal last_zzi As Long, ByVal dc_quant As Long)
    Dim dst As Long
    Dim frag_buf_off As Long
    Dim ystride As Long
    Dim refi As Long
    Dim p As Long
    Dim ci As Long
    Dim ref As Long
    Static mvoffsets(0 To 1) As Long    ' (VB) Static: no array set up per call
    ' Apply the inverse transform.
    ' Special case only having a DC component.
    If last_zzi < 2 Then
        ' We round this dequant product (and not any of the others) because there's no iDCT rounding.
        ' p=(ogg_int16_t)(_dct_coeffs[0]*(ogg_int32_t)_dc_quant+15>>5)
        p = dct_coeffs(0) * dc_quant + 15
        p = ((((p And &HFFFFFFE0) \ 32) + 32768) And &HFFFF&) - 32768
        ' LOOP VECTORIZES.
        For ci = 0 To 63
            dct_coeffs(64 + ci) = p
        Next
    Else
        ' First, dequantize the DC coefficient.
        dct_coeffs(0) = ((dct_coeffs(0) * dc_quant + 32768) And &HFFFF&) - 32768
        oc_idct8x8_c dct_coeffs, 64, dct_coeffs, 0, last_zzi
    End If
    ' Fill in the target buffer.
    frag_buf_off = st.frag_buf_offs(fragi)
    refi = st.frags(fragi).refi
    ystride = st.ref_ystride(pli)
    dst = st.ref_frame_data(OC_FRAME_SELF) + frag_buf_off
    If refi = OC_FRAME_SELF Then
        oc_frag_recon_intra_c st.ref_frame_handle, dst, ystride, dct_coeffs, 64
    Else
        ref = st.ref_frame_data(refi) + frag_buf_off
        If oc_state_get_mv_offsets(st, mvoffsets, pli, st.frag_mvs(fragi)) > 1 Then
            oc_frag_recon_inter2_c st.ref_frame_handle, dst, ref + mvoffsets(0), ref + mvoffsets(1), ystride, _
                dct_coeffs, 64
        Else
            oc_frag_recon_inter_c st.ref_frame_handle, dst, ref + mvoffsets(0), ystride, dct_coeffs, 64
        End If
    End If
End Sub

' (VB) _pix: pix(p); _bv: bv(bvOffset) (the C's &_bvarray[127])
Private Sub loop_filter_h(pix() As Byte, ByVal p As Long, ByVal ystride As Long, bv() As Long, ByVal bvOffset As Long)
    Dim y As Long
    Dim f As Long
    Dim v As Long
    p = p - 2
    For y = 0 To 7
        f = CLng(pix(p)) - pix(p + 3) + 3 * (CLng(pix(p + 2)) - pix(p + 1))
        ' The _bv array is used to compute the function
        '   f=OC_CLAMPI(OC_MINI(-_2flimit-f,0),f,OC_MAXI(_2flimit-f,0));
        '   where _2flimit=_state->loop_filter_limits[_state->qis[0]]<<1;
        f = bv(bvOffset + ((f + 4) And &HFFFFFFF8) \ 8)
        v = pix(p + 1) + f
        If v < 0 Then v = 0 Else If v > 255 Then v = 255
        pix(p + 1) = v
        v = pix(p + 2) - f
        If v < 0 Then v = 0 Else If v > 255 Then v = 255
        pix(p + 2) = v
        p = p + ystride
    Next
End Sub

Private Sub loop_filter_v(pix() As Byte, ByVal p As Long, ByVal ystride As Long, bv() As Long, ByVal bvOffset As Long)
    Dim x As Long
    Dim f As Long
    Dim v As Long
    p = p - ystride * 2
    For x = 0 To 7
        f = CLng(pix(p + x)) - pix(p + ystride * 3 + x) + 3 * (CLng(pix(p + ystride * 2 + x)) - pix(p + ystride + x))
        ' The _bv array is used to compute the function
        '   f=OC_CLAMPI(OC_MINI(-_2flimit-f,0),f,OC_MAXI(_2flimit-f,0));
        '   where _2flimit=_state->loop_filter_limits[_state->qis[0]]<<1;
        f = bv(bvOffset + ((f + 4) And &HFFFFFFF8) \ 8)
        v = pix(p + ystride + x) + f
        If v < 0 Then v = 0 Else If v > 255 Then v = 255
        pix(p + ystride + x) = v
        v = pix(p + ystride * 2 + x) - f
        If v < 0 Then v = 0 Else If v > 255 Then v = 255
        pix(p + ystride * 2 + x) = v
    Next
End Sub

' Initialize the bounding values array used by the loop filter.
'  _bv: Storage for the array.
'  _flimit: The filter limit as defined in Section 7.10 of the spec.
Public Sub oc_loop_filter_init_c(bv() As Long, ByVal flimit As Long)
    Dim i As Long
    For i = 0 To 255
        bv(i) = 0
    Next
    For i = 0 To flimit - 1
        If 127 - i - flimit >= 0 Then bv(127 - i - flimit) = ToInt8(i - flimit)
        bv(127 - i) = ToInt8(-i)
        bv(127 + i) = ToInt8(i)
        If 127 + i + flimit < 256 Then bv(127 + i + flimit) = ToInt8(flimit - i)
    Next
End Sub

' Apply the loop filter to a given set of fragment rows in the given plane. The filter may be run on the
' bottom edge, affecting pixels in the next row of fragments, so this row also needs to be available.
'  _bv:        The bounding values array.
'  _refi:      The index of the frame buffer to filter.
'  _pli:       The color plane to filter.
'  _fragy0:    The Y coordinate of the first fragment row to filter.
'  _fragy_end: The Y coordinate of the fragment row to stop filtering at.
' (VB) pix: st.ref_frame_handle
Public Sub oc_state_loop_filter_frag_rows_c(st As oc_theora_state, pix() As Byte, bvarray() As Long, _
    ByVal refi As Long, ByVal pli As Long, ByVal fragy0 As Long, ByVal fragy_end As Long)
    Dim fragi_top As Long
    Dim fragi_bot As Long
    Dim fragi0 As Long
    Dim fragi0_end As Long
    Dim ystride As Long
    Dim nhfrags As Long
    Dim ref_frame_data As Long
    Dim fragi As Long
    Dim fragi_end As Long
    Dim ref As Long
    ' signed char *_bv = &_bvarray[127]: (VB) bvarray from 127
    nhfrags = st.fplanes(pli).nhfrags
    fragi_top = st.fplanes(pli).froffset
    fragi_bot = fragi_top + st.fplanes(pli).nfrags
    fragi0 = fragi_top + fragy0 * nhfrags
    fragi0_end = fragi_top + fragy_end * nhfrags
    ystride = st.ref_ystride(pli)
    ref_frame_data = st.ref_frame_data(refi)
    ' The following loops are constructed somewhat non-intuitively on purpose. The main idea is: if a block
    ' boundary has at least one coded fragment on it, the filter is applied to it. However, the order that the
    ' filters are applied in matters, and VP3 chose the somewhat strange ordering used below.
    Do While fragi0 < fragi0_end
        fragi = fragi0
        fragi_end = fragi + nhfrags
        Do While fragi < fragi_end
            If st.frags(fragi).coded <> 0 Then
                ref = ref_frame_data + st.frag_buf_offs(fragi)
                If fragi > fragi0 Then loop_filter_h pix, ref, ystride, bvarray, 127
                If fragi0 > fragi_top Then loop_filter_v pix, ref, ystride, bvarray, 127
                If fragi + 1 < fragi_end Then
                    If st.frags(fragi + 1).coded = 0 Then loop_filter_h pix, ref + 8, ystride, bvarray, 127
                End If
                If fragi + nhfrags < fragi_bot Then
                    If st.frags(fragi + nhfrags).coded = 0 Then loop_filter_v pix, ref + ystride * 8, ystride, bvarray, 127
                End If
            End If
            fragi = fragi + 1
        Loop
        fragi0 = fragi0 + nhfrags
    Loop
End Sub

' (VB) _encdec is a decoder handle (theora_decode_c): the C casts the th_dec_ctx * to its leading
' oc_theora_state
Public Function th_granule_frame(ByVal encdec As Long, ByVal granpos As Double) As Double
    Dim iframe As Double
    Dim pframe As Double
    Dim shift As Double
    If granpos >= 0 Then
        shift = 2# ^ oc_decs(encdec).state.info.keyframe_granule_shift
        iframe = Int(granpos / shift)
        pframe = granpos - iframe * shift
        ' 3.2.0 streams store the frame index in the granule position. 3.2.1 and later store the frame count.
        ' We return the index, so adjust the value if we have a 3.2.1 or later stream.
        th_granule_frame = iframe + pframe - TH_VERSION_CHECK(oc_decs(encdec).state.info, 3, 2, 1)
        Exit Function
    End If
    th_granule_frame = -1
End Function

Public Function th_granule_time(ByVal encdec As Long, ByVal granpos As Double) As Double
    If granpos >= 0 Then
        th_granule_time = (th_granule_frame(encdec, granpos) + 1) * _
            (U32(oc_decs(encdec).state.info.fps_denominator) / U32(oc_decs(encdec).state.info.fps_numerator))
        Exit Function
    End If
    th_granule_time = -1
End Function
