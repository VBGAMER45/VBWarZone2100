Attribute VB_Name = "theora_codec_h"
' ref/theora (libtheora 1.2.0) include/theora/codec.h and theoradec.h: the th_ API's types and constants
' (the decoder's half; th_huff_code, the encoder's, is left out). The legacy API's (theora.h) are in theora_h.
' (VB) th_comment is theora_h's theora_comment: the C relies on the two structs being identical (decapiwrapper.c
' casts one to the other), so the th_comment functions (theora_info_c) take a theora_comment. A
' th_ycbcr_buffer (th_img_plane[3]) is a UDT holding the three planes, so that the C's array copies are
' assignments. ogg_uint32_t is a Long holding its bits.
Option Explicit

' Return codes
' An invalid pointer was provided.
Public Const TH_EFAULT As Long = -1
' An invalid argument was provided.
Public Const TH_EINVAL As Long = -10
' The contents of the header were incomplete, invalid, or unexpected.
Public Const TH_EBADHEADER As Long = -20
' The header does not belong to a Theora stream.
Public Const TH_ENOTFORMAT As Long = -21
' The bitstream version is too high.
Public Const TH_EVERSION As Long = -22
' The specified function is not implemented.
Public Const TH_EIMPL As Long = -23
' There were errors in the video data packet.
Public Const TH_EBADPACKET As Long = -24
' The decoded packet represented a dropped frame. The player can continue to display the current frame, as
' the contents of the decoded frame buffer have not changed.
Public Const TH_DUPFRAME As Long = 1

' The currently defined color space tags (th_colorspace).
' The color space was not specified at the encoder. It may be conveyed by an external means.
Public Const TH_CS_UNSPECIFIED As Long = 0
' A color space designed for NTSC content.
Public Const TH_CS_ITU_REC_470M As Long = 1
' A color space designed for PAL/SECAM content.
Public Const TH_CS_ITU_REC_470BG As Long = 2
' The total number of currently defined color spaces.
Public Const TH_CS_NSPACES As Long = 3

' The currently defined pixel format tags (th_pixel_fmt).
' Chroma decimation by 2 in both the X and Y directions (4:2:0). The Cb and Cr chroma planes are half the
' width and half the height of the luma plane.
Public Const TH_PF_420 As Long = 0
' Currently reserved.
Public Const TH_PF_RSVD As Long = 1
' Chroma decimation by 2 in the X direction (4:2:2). The Cb and Cr chroma planes are half the width of the
' luma plane, but full height.
Public Const TH_PF_422 As Long = 2
' No chroma decimation (4:4:4). The Cb and Cr chroma planes are full width and full height.
Public Const TH_PF_444 As Long = 3
' The total number of currently defined pixel formats.
Public Const TH_PF_NFORMATS As Long = 4

' A buffer for a single color plane in an uncompressed image. This contains the image data in a left-to-right,
' top-down format. Each row of pixels is stored contiguously in memory, but successive rows need not be.
' Use stride to compute the offset of the next row. The width value gives the number of bytes of image data
' per row; the stride may be negative.
Public Type th_img_plane
    width As Long           ' The width of this plane.
    height As Long          ' The height of this plane.
    stride As Long          ' The offset in bytes between successive rows.
    data As Long            ' unsigned char *: A pointer to the beginning of the first row. (VB) an offset into
                            ' the decoder's frame memory (oc_theora_state.ref_frame_handle), 0 = NULL
End Type

' A complete image buffer for an uncompressed frame. The chroma planes may be decimated by a factor of two in
' either direction, as indicated by th_info#pixel_fmt. The width parameter of the Y' plane must be a multiple
' of 16. The height parameter of the Y' plane must be a multiple of 16 for TH_PF_420, and a multiple of 8
' otherwise.
' (VB) the C's th_img_plane[3]: p(0) Y', p(1) Cb, p(2) Cr
Public Type th_ycbcr_buffer
    p(0 To 2) As th_img_plane
End Type

' Theora bitstream information. This contains the basic playback parameters for a stream, and corresponds to
' the initial 'info' header packet. To initialize an encoder, the application fills in this structure and
' passes it to th_encode_alloc(). A default encoding mode is chosen based on the values of the quality and
' target_bitrate fields. On decode, it is filled in by th_decode_headerin(), and then passed to
' th_decode_alloc().
' Encoded Theora frames must be a multiple of 16 in size; this is what the frame_width and frame_height
' members represent. To handle arbitrary picture sizes, a crop rectangle is specified in the pic_x, pic_y,
' pic_width and pic_height members.
' All frame buffers contain pointers to the full, padded frame. However, the current encoder will not
' reference pixels outside of the cropped picture region, and the application does not need to fill them
' in. The decoder will allocate storage for a full frame, but the application should not rely on the padding
' containing sensible data.
' It is also generally recommended that the offsets and sizes should still be multiples of 2 to avoid
' chroma sampling shifts when chroma is sub-sampled. See the Theora specification, Section 4.4, for more
' details.
' Frame rate, in frames per second, is stored as a rational fraction, as is the pixel aspect ratio. Note
' that this refers to the aspect ratio of the individual pixels, not of the overall frame itself. The frame
' aspect ratio can be computed from pixel aspect ratio using the image dimensions.
Public Type th_info
    ' Theora version
    ' Bitstream version information.
    version_major As Byte
    version_minor As Byte
    version_subminor As Byte
    ' The encoded frame width. This must be a multiple of 16, and less than 1048576.
    frame_width As Long                 ' ogg_uint32_t
    ' The encoded frame height. This must be a multiple of 16, and less than 1048576.
    frame_height As Long
    ' The displayed picture width. This must be no larger than width.
    pic_width As Long
    ' The displayed picture height. This must be no larger than height.
    pic_height As Long
    ' The X offset of the displayed picture. This must be no larger than frame_width-pic_width or 255,
    ' whichever is smaller.
    pic_x As Long
    ' The Y offset of the displayed picture. This must be no larger than frame_height-pic_height, and
    ' frame_height-pic_height-pic_y must be no larger than 255. This slightly funny restriction is due to the
    ' fact that the offset is specified from the top of the image for consistency with the standard graphics
    ' left-handed coordinate system used throughout this API, while it is stored in the encoded stream as an
    ' offset from the bottom.
    pic_y As Long
    ' Frame rate
    ' The frame rate, as a fraction. If either is 0, the frame rate is undefined.
    fps_numerator As Long
    fps_denominator As Long
    ' Aspect ratio
    ' The aspect ratio of the pixels. If either value is zero, the aspect ratio is undefined. If not
    ' specified by any external means, 1:1 should be assumed. The aspect ratio of the full picture can be
    ' computed as aspect_numerator*pic_width/(aspect_denominator*pic_height).
    aspect_numerator As Long
    aspect_denominator As Long
    ' The color space.
    colorspace As Long                  ' th_colorspace
    ' The pixel format.
    pixel_fmt As Long                   ' th_pixel_fmt
    ' The target bit-rate in bits per second. If initializing an encoder with this struct, set this field to
    ' a non-zero value to activate CBR encoding by default.
    target_bitrate As Long
    ' The target quality level. Valid values range from 0 to 63, inclusive, with higher values giving
    ' higher quality. If initializing an encoder with this struct, and target_bitrate is set to zero, VBR
    ' encoding at this quality will be activated by default.
    quality As Long
    ' The amount to shift to extract the last keyframe number from the granule position. This can be at most
    ' 31. th_info_init() will set this to a default value (currently 6, which is good for streaming
    ' applications), but you can set it to 0 to make every frame a keyframe. The maximum distance between
    ' key frames is 1<<keyframe_granule_shift. The keyframe frequency can be more finely controlled with
    ' TH_ENCCTL_SET_KEYFRAME_FREQUENCY_FORCE, which can also be adjusted during encoding (for example, to
    ' force the next frame to be a keyframe), but it cannot be set larger than the amount permitted by this
    ' field after the headers have been output.
    keyframe_granule_shift As Long
End Type

' A single base matrix (th_quant_base: unsigned char[64]).
' (VB) a matrix is 64 bytes of th_quant_info.vb_matrices.

' A set of qi ranges.
Public Type th_quant_ranges
    ' The number of ranges in the set.
    nranges As Long
    ' The size of each of the nranges ranges. These must sum to 63.
    sizes As Long           ' const int *: (VB) the index of sizes[0] in the th_quant_info's vb_sizes, -1 = NULL
    ' nranges+1 base matrices. Matrices i and i+1 form the endpoints of range i.
    base_matrices As Long   ' const th_quant_base *: (VB) the number of the first matrix in the th_quant_info's
                            ' vb_matrices (matrix m is vb_matrices(m * 64 + ci)), -1 = NULL
End Type

' A complete set of quantization parameters. The quantizer for each coefficient is calculated as:
'   Q=MAX(MIN(qmin[qti][ci!=0],scale[ci!=0][qi]*base[qti][pli][qi][ci]/100), 1024).
' qti is the quantization type index: 0 for intra, 1 for inter. ci!=0 is 0 for the DC coefficient and 1 for
' AC coefficients. qi is the quality index, ranging between 0 (low quality) and 63 (high quality). pli is the
' color plane index: 0 for Y', 1 for Cb, 2 for Cr. ci is the DCT coefficient index. Coefficient indices
' correspond to the normal 2D DCT basis functions, in standard raster order (not zig-zag order).
' qmin is an encoder-specific per-plane minimum quantizer. base is interpolated from a set of base matrices
' given by the ranges, which are specified in qi_ranges. qi_ranges[qti][pli] are the ranges for the given
' type of block in the given color plane.
Public Type th_quant_info
    ' The DC scaling factors.
    dc_scale(0 To 63) As Long           ' ogg_uint16_t
    ' The AC scaling factors.
    ac_scale(0 To 63) As Long           ' ogg_uint16_t
    ' The loop filter limit values.
    loop_filter_limits(0 To 63) As Byte
    ' The qi ranges for each ci and pli.
    qi_ranges(0 To 5) As th_quant_ranges    ' [2][3]: (VB) qi_ranges(qti * 3 + pli)
    ' (VB) the storage the ranges point into (the C mallocs each list): the sizes, the base matrices
    vb_sizes() As Long
    vb_nsizes As Long
    vb_matrices() As Byte
    vb_nmatrices As Long
End Type

' The number of Huffman tables used by Theora.
Public Const TH_NHUFFMAN_TABLES As Long = 80
' The number of DCT token values in each table.
Public Const TH_NDCT_TOKENS As Long = 32

' theoradec.h: the decoder control codes (TH_DECCTL_GET_PPLEVEL_MAX (1), TH_DECCTL_SET_PPLEVEL (3) and
' TH_DECCTL_SET_GRANPOS (5) are theora_h's: both headers define them with the same values)
' Sets the striped decode callback function. If set, this function will be called as each piece of a frame
' is fully decoded in th_decode_packetin(). (VB) not ported: a function pointer the legacy API cannot pass;
' th_decode_ctl answers TH_EIMPL.
Public Const TH_DECCTL_SET_STRIPE_CB As Long = 7
' Enables telemetry and sets the macroblock display mode (built only with HAVE_CAIRO: not ported)
Public Const TH_DECCTL_SET_TELEMETRY_MBMODE As Long = 9
' Enables telemetry and sets the motion vector display mode
Public Const TH_DECCTL_SET_TELEMETRY_MV As Long = 11
' Enables telemetry and sets the adaptive quantization display mode
Public Const TH_DECCTL_SET_TELEMETRY_QI As Long = 13
' Enables telemetry and sets the bitstream breakdown visualization mode
Public Const TH_DECCTL_SET_TELEMETRY_BITS As Long = 15
