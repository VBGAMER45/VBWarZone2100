Attribute VB_Name = "opusfile_h"
' opusfile (lib/sound/3rdparty/opusfile at 9d718345, 4.7.0's submodule), include/opusfile.h: the error codes,
' the header structs and the stream callbacks. The functions are opusfile_c's (opusfile.c), opusfile_info_c's
' (info.c) and opusfile_stream_c's (stream.c). 4.7.0 builds it without http.c (OP_DISABLE_HTTP): the URL
' options, OpusServerInfo and the op_url_* functions are not ported.
' (VB) A char * of these structs is a String holding one byte per character (the bytes of the stream as they
' are, UTF-8 or binary: the format is 8-bit clean), vbNullString for NULL. An OggOpusFile * is opusfile_c's
' handle (0 = NULL). A function pointer of OpusFileCallbacks is a number (OPFN_*, 0 = NULL) that
' opusfile_stream_c's dispatch calls; the void *_stream is a Variant (a Long handle or an object).
Option Explicit

' stdio.h: the seek origins
Public Const SEEK_SET As Long = 0
Public Const SEEK_CUR As Long = 1
Public Const SEEK_END As Long = 2

' \defgroup error_codes Error Codes
' A request did not succeed.
Public Const OP_FALSE As Long = -1
' Currently not used externally.
Public Const OP_EOF As Long = -2
' There was a hole in the page sequence numbers (e.g., a page was corrupt or
'    missing).
Public Const OP_HOLE As Long = -3
' An underlying read, seek, or tell operation failed when it should have
'    succeeded.
Public Const OP_EREAD As Long = -128
' A <code>NULL</code> pointer was passed where one was unexpected, or an
'    internal memory allocation failed, or an internal library error was
'    encountered.
Public Const OP_EFAULT As Long = -129
' The stream used a feature that is not implemented, such as an unsupported
'    channel family.
Public Const OP_EIMPL As Long = -130
' One or more parameters to a function were invalid.
Public Const OP_EINVAL As Long = -131
' A purported Ogg Opus stream did not begin with an Ogg page, a purported
'    header packet did not start with one of the required strings, "OpusHead" or
'    "OpusTags", or a link in a chained file was encountered that did not
'    contain any logical Opus streams.
Public Const OP_ENOTFORMAT As Long = -132
' A required header packet was not properly formatted, contained illegal
'    values, or was missing altogether.
Public Const OP_EBADHEADER As Long = -133
' The ID header contained an unrecognized version number.
Public Const OP_EVERSION As Long = -134
' Currently not used at all.
Public Const OP_ENOTAUDIO As Long = -135
' An audio packet failed to decode properly.
'   This is usually caused by a multistream Ogg packet where the durations of
'    the individual Opus packets contained in it are not all the same.
Public Const OP_EBADPACKET As Long = -136
' We failed to find data we had seen before, or the bitstream structure was
'    sufficiently malformed that seeking to the target destination was
'    impossible.
Public Const OP_EBADLINK As Long = -137
' An operation that requires seeking was requested on an unseekable stream.
Public Const OP_ENOSEEK As Long = -138
' The first or last granule position of a link failed basic validity checks.
Public Const OP_EBADTIMESTAMP As Long = -139

' \defgroup header_info Header Information
' The maximum number of channels in an Ogg Opus stream.
Public Const OPUS_CHANNEL_COUNT_MAX As Long = 255

' Ogg Opus bitstream information.
'   This contains the basic playback parameters for a stream, and corresponds to
'    the initial ID header packet of an Ogg Opus stream.
Public Type OpusHead
    ' The Ogg Opus format version, in the range 0...255.
    '   The top 4 bits represent a "major" version, and the bottom four bits
    '    represent backwards-compatible "minor" revisions.
    version As Long
    ' The number of channels, in the range 1...255.
    channel_count As Long
    ' The number of samples that should be discarded from the beginning of the
    '    stream.
    pre_skip As Long                ' unsigned
    ' The sampling rate of the original input.
    '   All Opus audio is coded at 48 kHz, and should also be decoded at 48 kHz
    '    for playback (unless the target hardware does not support this sampling
    '    rate).
    input_sample_rate As Long       ' opus_uint32 ((VB) the bits)
    ' The gain to apply to the decoded output, in dB, as a Q8 value in the range
    '    -32768...32767.
    '   The <tt>libopusfile</tt> API will automatically apply this gain to the
    '    decoded output before returning it, scaling it by
    '    <code>pow(10,output_gain/(20.0*256))</code>.
    output_gain As Long
    ' The channel mapping family, in the range 0...255.
    '   Channel mapping family 0 covers mono or stereo in a single stream.
    '   Channel mapping family 1 covers 1 to 8 channels in one or more streams,
    '    using the Vorbis speaker assignments.
    '   Channel mapping family 255 covers 1 to 255 channels in one or more
    '    streams, but without any defined speaker assignment.
    mapping_family As Long
    ' The number of Opus streams in each Ogg packet, in the range 1...255.
    stream_count As Long
    ' The number of coupled Opus streams in each Ogg packet, in the range
    '    0...127.
    coupled_count As Long
    ' The mapping from coded stream channels to output channels.
    '   Let <code>index=mapping[k]</code> be the value for channel <code>k</code>.
    '   If <code>index<2*coupled_count</code>, then it refers to the left channel
    '    from stream <code>(index/2)</code> if even, and the right channel from
    '    stream <code>(index/2)</code> if odd.
    '   Otherwise, it refers to the output of the uncoupled stream
    '    <code>(index-coupled_count)</code>.
    mapping(0 To OPUS_CHANNEL_COUNT_MAX - 1) As Byte
End Type

' The metadata from an Ogg Opus stream.
'   This structure holds the in-stream metadata corresponding to the 'comment'
'    header packet of an Ogg Opus stream.
'   In filling in this structure, <tt>libopusfile</tt> will null-terminate the
'    #user_comments strings for safety.
'   However, the bitstream format itself treats them as 8-bit clean vectors,
'    possibly containing NUL characters, so the #comment_lengths array should be
'    treated as their authoritative length.
Public Type OpusTags
    ' The array of comment string vectors. (VB) user_comments(comments) holds the binary suffix data, if any.
    user_comments() As String
    ' An array of the corresponding length of each vector, in bytes.
    comment_lengths() As Long
    ' The total number of comment streams.
    comments As Long
    ' The null-terminated vendor string.
    '   This identifies the software used to encode the stream.
    vendor As String
    vbAllocated As Boolean          ' (VB) user_comments and comment_lengths != NULL
End Type

' \name Picture tag image formats
' The MIME type was not recognized, or the image data did not match the
'    declared MIME type.
Public Const OP_PIC_FORMAT_UNKNOWN As Long = -1
' The MIME type indicates the image data is really a URL.
Public Const OP_PIC_FORMAT_URL As Long = 0
' The image is a JPEG.
Public Const OP_PIC_FORMAT_JPEG As Long = 1
' The image is a PNG.
Public Const OP_PIC_FORMAT_PNG As Long = 2
' The image is a GIF.
Public Const OP_PIC_FORMAT_GIF As Long = 3

' The contents of a METADATA_BLOCK_PICTURE tag.
Public Type OpusPictureTag
    ' The picture type according to the ID3v2 APIC frame
    '   Others are reserved and should not be used.
    '   There may only be one each of picture type 1 and 2 in a file.
    typ As Long                     ' opus_int32 type
    ' The MIME type of the picture, in printable ASCII characters 0x20-0x7E.
    '   The MIME type may also be <code>"-->"</code> to signify that the data part
    '    is a URL pointing to the picture instead of the picture data itself.
    '   In this case, a terminating NUL is appended to the URL string in #data,
    '    but #data_length is set to the length of the string excluding that
    '    terminating NUL.
    mime_type As String
    ' The description of the picture, in UTF-8.
    description As String
    ' The width of the picture in pixels.
    width As Long
    ' The height of the picture in pixels.
    height As Long
    ' The color depth of the picture in bits-per-pixel (<em>not</em>
    '    bits-per-channel).
    depth As Long
    ' For indexed-color pictures (e.g., GIF), the number of colors used, or 0
    '    for non-indexed pictures.
    colors As Long
    ' The length of the picture data in bytes.
    data_length As Long
    ' The binary picture data.
    data() As Byte
    ' The format of the picture data, if known.
    format_ As Long                 ' format
End Type

' \defgroup stream_callbacks Abstract Stream Reading Interface
' The callbacks used to access non-<code>FILE</code> stream resources.
'   The function prototypes are basically the same as for the stdio functions
'    <code>fread()</code>, <code>fseek()</code>, <code>ftell()</code>, and
'    <code>fclose()</code>.
'   The differences are that the <code>FILE *</code> arguments have been
'    replaced with a <code>void *</code>, which is to be used as a pointer to
'    whatever internal data these functions might need, that #seek and #tell
'    take and return 64-bit offsets, and that #seek <em>must</em> return -1 if
'    the stream is unseekable.
' (VB) each one an OPFN_* number of opusfile_stream_c's dispatch (0 = NULL)
Public Type OpusFileCallbacks
    ' Used to read data from the stream.
    '   This must not be <code>NULL</code>.
    read_ As Long                   ' op_read_func read
    ' Used to seek in the stream.
    '   This may be <code>NULL</code> if seeking is not implemented.
    seek_ As Long                   ' op_seek_func seek
    ' Used to return the current read position in the stream.
    '   This may be <code>NULL</code> if seeking is not implemented.
    tell As Long                    ' op_tell_func tell
    ' Used to close the stream when the decoder is freed.
    '   This may be <code>NULL</code> to leave the stream open.
    close_ As Long                  ' op_close_func close
End Type

' \name Decoding
' Indicates that the decoding callback should produce signed 16-bit
'    native-endian output samples.
Public Const OP_DEC_FORMAT_SHORT As Long = 7008
' Indicates that the decoding callback should produce 32-bit native-endian
'    float samples.
Public Const OP_DEC_FORMAT_FLOAT As Long = 7040
' Indicates that the decoding callback did not decode anything, and that
'    <tt>libopusfile</tt> should decode normally instead.
Public Const OP_DEC_USE_DEFAULT As Long = 6720

' Gain offset type that indicates that the provided offset is relative to the
'    header gain.
'   This is the default.
Public Const OP_HEADER_GAIN As Long = 0
' Gain offset type that indicates that the provided offset is relative to the
'    R128_ALBUM_GAIN value (if any), in addition to the header gain.
Public Const OP_ALBUM_GAIN As Long = 3007
' Gain offset type that indicates that the provided offset is relative to the
'    R128_TRACK_GAIN value (if any), in addition to the header gain.
Public Const OP_TRACK_GAIN As Long = 3008
' Gain offset type that indicates that the provided offset should be used as
'    the gain directly, without applying any the header or track gains.
Public Const OP_ABSOLUTE_GAIN As Long = 3009
