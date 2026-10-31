Attribute VB_Name = "theora_h"
' ref/theora (libtheora 1.2.0) include/theora/theora.h: the legacy pre-1.0 API's types and constants, which
' lib/sequence/sequence.cpp uses. The functions are in the modules of their C files:
'   theora_apiwrapper_c (lib/apiwrapper.c): theora_info_init, theora_info_clear, theora_clear, theora_control,
'     theora_granule_frame, theora_granule_time, theora_packet_isheader, theora_packet_iskeyframe,
'     theora_granule_shift, theora_comment_init, theora_comment_query, theora_comment_query_count,
'     theora_comment_clear, theora_comment_add, theora_comment_add_tag, theora_version_string,
'     theora_version_number
'   theora_decapiwrapper_c (lib/decapiwrapper.c): theora_decode_init, theora_decode_header,
'     theora_decode_packetin, theora_decode_YUVout
' Their VB signatures:
'   Public Sub theora_info_init(ci As theora_info)
'   Public Sub theora_info_clear(ci As theora_info)
'   Public Sub theora_comment_init(tc As theora_comment)
'   Public Sub theora_comment_clear(tc As theora_comment)
'   Public Function theora_decode_header(ci As theora_info, cc As theora_comment, op As ogg_packet) As Long
'   Public Function theora_decode_init(th As theora_state, c As theora_info) As Long
'   Public Function theora_decode_packetin(th As theora_state, op As ogg_packet) As Long
'   Public Function theora_decode_YUVout(th As theora_state, yuv As yuv_buffer) As Long
'   Public Function theora_granule_frame(th As theora_state, ByVal granulepos As Double) As Double
'   Public Function theora_granule_time(th As theora_state, ByVal granulepos As Double) As Double
'   Public Function theora_control(th As theora_state, ByVal req As Long, buf As Long, ByVal buf_sz As Long) As Long
'   Public Sub theora_clear(th As theora_state)
' (VB) Pointers are handles (0 = NULL): theora_info.codec_setup is a th_api_wrapper of
' theora_decapiwrapper_c's pool, theora_state.i the th_api_info whose theora_info the state points at.
' ogg_int64_t is a Double, ogg_uint32_t a Long holding its bits. theora_control's void *buf is a Long (the
' int of TH_DECCTL_GET_PPLEVEL_MAX and TH_DECCTL_SET_PPLEVEL; buf_sz is then 4, sizeof(int)).
' yuv_buffer's planes are copies (below), made by theora_decode_YUVout.
Option Explicit

' A YUV buffer for passing uncompressed frames to and from the codec.
' This holds a Y'CbCr frame in planar format. The CbCr planes can be subsampled and have their own separate
' dimensions and row stride offsets. Note that the strides may be negative in some configurations. For
' theora the width and height of the largest plane must be a multiple of 16. The actual meaningful picture
' size and offset are stored in the theora_info structure; frames returned by the decoder may need to be
' cropped for display.
' All samples are 8 bits. Within each plane samples are ordered by row from the top of the frame to the
' bottom. Within each row samples are ordered from left to right.
' (VB) y(), u() and v() are copies of the planes the C's pointers point at (the decoder's reference frame,
' read bottom-up with a negative stride): the whole encoded frame (y_width x y_height, uv_width x
' uv_height; the picture region is theora_info's frame_width x frame_height at offset_x, offset_y), top row
' first, index 0 = the C's y[0] (the top left pixel), and the strides are the planes' widths, so the C's
' y[row * y_stride + x] is the same element. theora_decode_YUVout keeps the arrays when their sizes fit.
Public Type yuv_buffer
    y_width As Long         ' Width of the Y' luminance plane
    y_height As Long        ' Height of the luminance plane
    y_stride As Long        ' Offset in bytes between successive rows

    uv_width As Long        ' Width of the Cb and Cr chroma planes
    uv_height As Long       ' Height of the chroma planes
    uv_stride As Long       ' Offset between successive chroma rows

    y() As Byte             ' Pointer to start of luminance data
    u() As Byte             ' Pointer to start of Cb data
    v() As Byte             ' Pointer to start of Cr data
End Type

' A Colorspace (theora_colorspace)
Public Const OC_CS_UNSPECIFIED As Long = 0      ' The colorspace is unknown or unspecified
Public Const OC_CS_ITU_REC_470M As Long = 1     ' This is the best option for 'NTSC' content
Public Const OC_CS_ITU_REC_470BG As Long = 2    ' This is the best option for 'PAL' content
Public Const OC_CS_NSPACES As Long = 3          ' This marks the end of the defined colorspaces

' A Chroma subsampling (theora_pixelformat)
' These enumerate the available chroma subsampling options supported by the theora format. See Section 4.4
' of the specification for exact definitions.
Public Const OC_PF_420 As Long = 0      ' Chroma subsampling by 2 in each direction (4:2:0)
Public Const OC_PF_RSVD As Long = 1     ' Reserved value
Public Const OC_PF_422 As Long = 2      ' Horizonatal chroma subsampling by 2 (4:2:2)
Public Const OC_PF_444 As Long = 3      ' No chroma subsampling at all (4:4:4)

' Theora bitstream info.
' Contains the basic playback parameters for a stream, corresponding to the initial 'info' header packet.
' Encoded theora frames must be a multiple of 16 in width and height. To handle other frame sizes, a crop
' rectangle is specified in frame_height and frame_width, offset_x and offset_y. The offset and size should
' still be a multiple of 2 to avoid chroma sampling shifts. Offset values in this structure are measured
' from the upper left of the image.
' Frame rate, in frames per second, is stored as a rational fraction. Aspect ratio is also stored as a
' rational fraction, and refers to the aspect ratio of the frame pixels, not of the overall frame itself.
' See [`examples/encoder_example.c`](https://gitlab.xiph.org/xiph/theora/-/blob/main/examples/encoder_example.c)
' for usage examples of the other parameters and good default settings for the encoder parameters.
Public Type theora_info
    width As Long               ' ogg_uint32_t: encoded frame width
    height As Long              ' encoded frame height
    frame_width As Long         ' display frame width
    frame_height As Long        ' display frame height
    offset_x As Long            ' horizontal offset of the displayed frame
    offset_y As Long            ' vertical offset of the displayed frame
    fps_numerator As Long       ' frame rate numerator
    fps_denominator As Long     ' frame rate denominator
    aspect_numerator As Long    ' pixel aspect ratio numerator
    aspect_denominator As Long  ' pixel aspect ratio denominator
    colorspace As Long          ' theora_colorspace: colorspace
    target_bitrate As Long      ' nominal bitrate in bits per second
    quality As Long             ' Nominal quality setting, 0-63
    quick_p As Long             ' Quick encode/decode

    ' decode only
    version_major As Byte
    version_minor As Byte
    version_subminor As Byte

    codec_setup As Long         ' void *: (VB) a th_api_wrapper handle (theora_decapiwrapper_c), 0 = NULL

    ' encode only
    dropframes_p As Long
    keyframe_auto_p As Long
    keyframe_frequency As Long          ' ogg_uint32_t
    keyframe_frequency_force As Long    ' ogg_uint32_t: also used for decode init to get granpos shift correct
    keyframe_data_target_bitrate As Long    ' ogg_uint32_t
    keyframe_auto_threshold As Long     ' ogg_int32_t
    keyframe_mindistance As Long        ' ogg_uint32_t
    noise_sensitivity As Long           ' ogg_int32_t
    sharpness As Long                   ' ogg_int32_t

    pixelformat As Long         ' theora_pixelformat: chroma subsampling mode to expect
End Type

' Codec internal state and context.
Public Type theora_state
    i As Long                   ' theora_info *: (VB) the th_api_info handle whose info this points at, 0 = NULL
    granulepos As Double        ' ogg_int64_t
    internal_encode As Long     ' void *: (VB) 0 (no encoder)
    internal_decode As Long     ' void *: (VB) 1 = &OC_DEC_DISPATCH_VTBL, 0 = NULL
End Type

' Comment header metadata.
' This structure holds the in-stream metadata corresponding to the 'comment' header packet.
' Meta data is stored as a series of (tag, value) pairs, in length-encoded string vectors. The first
' occurence of the '=' character delimits the tag and value. A particular tag may occur more than once. The
' character set encoding for the strings is always UTF-8, but the tag names are limited to case-insensitive
' ASCII. See the spec for details.
' In filling in this structure, theora_decode_header() will null-terminate the user_comment strings for
' safety. However, the bitstream format itself treats them as 8-bit clean, and so the length array should
' be treated as authoritative for their length.
' (VB) The strings hold the bytes as characters 0-255 (no UTF-8 decoding, so the lengths in bytes are the
' strings' lengths); user_comments(0 To comments - 1).
Public Type theora_comment
    user_comments() As String   ' char **: An array of comment string vectors
    comment_lengths() As Long   ' int *: An array of corresponding string vector lengths in bytes
    comments As Long            ' The total number of comment string vectors
    vendor As String            ' char *: The vendor string identifying the encoder, null terminated
End Type

' theora_control() codes
' These are the available request codes for theora_control() when called with a decoder instance. By
' convention, these are odd, to distinguish them from the encoder control codes. Keep any experimental or
' vendor-specific values above 0x8000.
' Get the maximum post-processing level. The decoder supports a post-processing filter that can improve the
' appearance of the decoded images. This returns the highest level setting for this post-processor,
' corresponding to maximum improvement and computational expense.
Public Const TH_DECCTL_GET_PPLEVEL_MAX As Long = 1
' Set the post-processing level. Sets the level of post-processing to use when decoding the compressed
' stream. This must be a value between zero (off) and the maximum returned by TH_DECCTL_GET_PPLEVEL_MAX.
Public Const TH_DECCTL_SET_PPLEVEL As Long = 3
' Sets the maximum distance between key frames. (encoder)
Public Const TH_ENCCTL_SET_KEYFRAME_FREQUENCY_FORCE As Long = 4
' Set the granule position. Call this after a seek, to update the internal granulepos in the decoder, to
' insure that subsequent frames are marked properly. If you track timestamps yourself and do not use the
' granule postion returned by the decoder, then you do not need to use this control.
Public Const TH_DECCTL_SET_GRANPOS As Long = 5
' Sets the quantization parameters to use. (encoder)
Public Const TH_ENCCTL_SET_QUANT_PARAMS As Long = 2
' Disables any encoder features that would prevent lossless transcoding back to VP3. (encoder)
Public Const TH_ENCCTL_SET_VP3_COMPATIBLE As Long = 10
' Gets the maximum speed level. (encoder)
Public Const TH_ENCCTL_GET_SPLEVEL_MAX As Long = 12
' Sets the speed level. (encoder)
Public Const TH_ENCCTL_SET_SPLEVEL As Long = 14

' Error codes
Public Const OC_FAULT As Long = -1          ' General failure
Public Const OC_EINVAL As Long = -10        ' Library encountered invalid internal data
Public Const OC_DISABLED As Long = -11      ' Requested action is disabled
Public Const OC_BADHEADER As Long = -20     ' Header packet was corrupt/invalid
Public Const OC_NOTFORMAT As Long = -21     ' Packet is not a theora packet
Public Const OC_VERSION As Long = -22       ' Bitstream version is not handled
Public Const OC_IMPL As Long = -23          ' Feature or action not implemented
Public Const OC_BADPACKET As Long = -24     ' Packet is corrupt
Public Const OC_NEWPACKET As Long = -25     ' Packet is an (ignorable) unhandled extension
Public Const OC_DUPFRAME As Long = 1        ' Packet is a dropped frame
