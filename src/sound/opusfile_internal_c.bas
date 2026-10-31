Attribute VB_Name = "opusfile_internal_c"
' opusfile (lib/sound/3rdparty/opusfile at 9d718345), src/internal.h and src/internal.c: the private structs
' (OggOpusLink, OggOpusFile), the limits and op_strncasecmp. The float build (OP_FIXED_POINT undefined, as
' 4.7.0 builds it), so op_sample is float and OP_SOFT_CLIP is defined (libopus 1.5.2 has
' OPUS_GET_EXPERT_FRAME_DURATION_REQUEST). OP_ENABLE_ASSERTIONS is not defined: OP_ASSERT is nothing,
' OP_ALWAYS_TRUE evaluates its argument (op_fatal_impl is not compiled). OP_LIKELY and OP_UNLIKELY are their
' argument. The (VB) types: ogg_int64_t and opus_int64 are Doubles (exact to 2^53; OP_INT64_MAX is 2^63 as
' the nearest Double), ogg_uint32_t and opus_uint32 Longs holding the bits; a pointer of OggOpusFile is a
' dynamic array (with a flag where the C tests it against NULL); an OggOpusFile * is a handle of opusfile_c's
' pool, an OpusMSDecoder * a handle of the multistream decoder (opus_multistream_decoder_c, 0 = NULL).
Option Explicit

Public Const OP_INT64_MAX As Double = 4294967296# * 2147483648#     ' (2*(((ogg_int64_t)1<<62)-1)|1): (VB) 2^63
Public Const OP_INT64_MIN As Double = -4294967296# * 2147483648#    ' (-OP_INT64_MAX-1)
Public Const OP_INT32_MAX As Long = 2147483647                      ' (2*(((ogg_int32_t)1<<30)-1)|1)
Public Const OP_INT32_MIN As Long = &H80000000                      ' (-OP_INT32_MAX-1)

' The maximum channel count for any mapping we'll actually decode.
Public Const OP_NCHANNELS_MAX As Long = 8

' Initial state.
Public Const OP_NOTOPEN As Long = 0
' We've found the first Opus stream in the first link.
Public Const OP_PARTOPEN As Long = 1
Public Const OP_OPENED As Long = 2
' We've found the first Opus stream in the current link.
Public Const OP_STREAMSET As Long = 3
' We've initialized the decoder for the chosen Opus stream in the current
'   link.
Public Const OP_INITSET As Long = 4

' Information cached for a single link in a chained Ogg Opus file.
'  We choose the first Opus stream encountered in each link to play back (and
'   require at least one).
Public Type OggOpusLink
    ' The byte offset of the first header page in this link.
    offset As Double                ' opus_int64
    ' The byte offset of the first data page from the chosen Opus stream in this
    '   link (after the headers).
    data_offset As Double
    ' The byte offset of the last page from the chosen Opus stream in this link.
    '  This is used when seeking to ensure we find a page before the last one, so
    '   that end-trimming calculations work properly.
    '  This is only valid for seekable sources.
    end_offset As Double
    ' The total duration of all prior links.
    '  This is always zero for non-seekable sources.
    pcm_file_offset As Double       ' ogg_int64_t
    ' The granule position of the last sample.
    '  This is only valid for seekable sources.
    pcm_end As Double
    ' The granule position before the first sample.
    pcm_start As Double
    ' The serial number.
    serialno As Long                ' ogg_uint32_t
    ' The contents of the info header.
    head As OpusHead
    ' The contents of the comment header.
    tags As OpusTags
End Type

Public Type OggOpusFile
    ' The callbacks used to access the stream.
    callbacks As OpusFileCallbacks
    ' A FILE *, memory buffer, etc.
    stream As Variant               ' void *
    ' Whether or not we can seek with this stream.
    seekable As Long
    ' The number of links in this chained Ogg Opus file.
    nlinks As Long
    ' The cached information from each link in a chained Ogg Opus file.
    '  If stream isn't seekable (e.g., it's a pipe), only the current link
    '   appears.
    links() As OggOpusLink
    links_allocated As Boolean      ' (VB) links != NULL
    ' The number of serial numbers from a single link.
    nserialnos As Long
    ' The capacity of the list of serial numbers from a single link.
    cserialnos As Long
    ' Storage for the list of serial numbers from a single link.
    '  This is a scratch buffer used when scanning the BOS pages at the start of
    '   each link.
    serialnos() As Long             ' ogg_uint32_t *: (VB) allocated when cserialnos > 0
    ' This is the current offset of the data processed by the ogg_sync_state.
    '  After a seek, this should be set to the target offset so that we can track
    '   the byte offsets of subsequent pages.
    '  After a call to op_get_next_page(), this will point to the first byte after
    '   that page.
    offset As Double                ' opus_int64
    ' The total size of this stream, or -1 if it's unseekable.
    end_ As Double                  ' opus_int64 end
    ' Used to locate pages in the stream.
    oy As ogg_sync_state
    ' One of OP_NOTOPEN, OP_PARTOPEN, OP_OPENED, OP_STREAMSET, OP_INITSET.
    ready_state As Long
    ' The current link being played back.
    cur_link As Long
    ' The number of decoded samples to discard from the start of decoding.
    cur_discard_count As Long
    ' The granule position of the previous packet (current packet start time).
    prev_packet_gp As Double
    ' The stream offset of the most recent page with completed packets, or -1.
    '  This is only needed to recover continued packet data in the seeking logic,
    '   when we use the current position as one of our bounds, only to later
    '   discover it was the correct starting point.
    prev_page_offset As Double
    ' The number of bytes read since the last bitrate query, including framing.
    bytes_tracked As Double
    ' The number of samples decoded since the last bitrate query.
    samples_tracked As Double
    ' Takes physical pages and welds them into a logical stream of packets.
    os As ogg_stream_state
    ' Re-timestamped packets from a single page.
    '  Buffering these relies on the undocumented libogg behavior that ogg_packet
    '   pointers remain valid until the next page is submitted to the
    '   ogg_stream_state they came from. ((VB) the packets are copies)
    op(0 To 254) As ogg_packet
    ' The index of the next packet to return.
    op_pos As Long
    ' The total number of packets available.
    op_count As Long
    ' Central working state for the packet-to-PCM decoder.
    od As Long                      ' OpusMSDecoder *
    ' The application-provided packet decode callback.
    decode_cb As Long               ' op_decode_cb_func: (VB) a number, 0 = NULL
    ' The application-provided packet decode callback context.
    decode_cb_ctx As Variant        ' void *
    ' The stream count used to initialize the decoder.
    od_stream_count As Long
    ' The coupled stream count used to initialize the decoder.
    od_coupled_count As Long
    ' The channel count used to initialize the decoder.
    od_channel_count As Long
    ' The channel mapping used to initialize the decoder.
    od_mapping(0 To OP_NCHANNELS_MAX - 1) As Byte
    ' The buffered data for one decoded packet.
    od_buffer() As Single           ' op_sample *
    od_buffer_allocated As Boolean  ' (VB) od_buffer != NULL
    ' The current position in the decoded buffer.
    od_buffer_pos As Long
    ' The number of valid samples in the decoded buffer.
    od_buffer_size As Long
    ' The type of gain offset to apply.
    '  One of OP_HEADER_GAIN, OP_ALBUM_GAIN, OP_TRACK_GAIN, or OP_ABSOLUTE_GAIN.
    gain_type As Long
    ' The offset to apply to the gain.
    gain_offset_q8 As Long
    ' Internal state for soft clipping and dithering float->short output.
    ' (OP_SOFT_CLIP)
    clip_state(0 To OP_NCHANNELS_MAX - 1) As Single
    dither_a(0 To OP_NCHANNELS_MAX * 4 - 1) As Single
    dither_b(0 To OP_NCHANNELS_MAX * 4 - 1) As Single
    dither_seed As Long             ' opus_uint32
    dither_mute As Long
    dither_disabled As Long
    ' The number of channels represented by the internal state.
    '  This gets set to 0 whenever anything that would prevent state propagation
    '   occurs (switching between the float/short APIs, or between the
    '   stereo/multistream APIs).
    state_channel_count As Long
End Type

' (VB) OP_ADV_OFFSET: Advance a file offset by the given amount, clamping against OP_INT64_MAX.
'  This is used to advance a known offset by things like OP_CHUNK_SIZE or
'   OP_PAGE_SIZE_MAX, while making sure to avoid signed overflow.
'  It assumes that both _offset and _amount are non-negative.
Public Function OP_ADV_OFFSET(ByVal offset_ As Double, ByVal amount_ As Double) As Double
    If offset_ < OP_INT64_MAX - amount_ Then OP_ADV_OFFSET = offset_ + amount_ Else OP_ADV_OFFSET = OP_INT64_MAX - amount_ + amount_
End Function

' A version of strncasecmp() that is guaranteed to only ignore the case of
'   ASCII characters.
' (VB) a_ and b_ from aOff and bOff; a character past a String's end is its NUL. char is signed (x86).
Public Function op_strncasecmp(ByVal a_ As String, ByVal aOff As Long, ByVal b_ As String, ByVal bOff As Long, ByVal n_ As Long) As Long
    Dim i As Long, a As Long, b As Long, d As Long
    For i = 0 To n_ - 1
        a = op_vbChar(a_, aOff + i)
        b = op_vbChar(b_, bOff + i)
        If a >= 97 And a <= 122 Then a = a - (97 - 65)          ' if(a>='a'&&a<='z')a-='a'-'A'
        If b >= 97 And b <= 122 Then b = b - (97 - 65)
        d = a - b
        If d Then op_strncasecmp = d: Exit Function
    Next
    op_strncasecmp = 0
End Function

' (VB) the C's s[i] for a char * string held one byte per character: 0 past the end (the terminating NUL), a
' signed char (-128 ... 127) as on x86
Public Function op_vbChar(s As String, ByVal i As Long) As Long
    If i < 0 Or i >= Len(s) Then
        op_vbChar = 0
    Else
        op_vbChar = AscW(Mid$(s, i + 1, 1)) And &HFF&
        If op_vbChar >= 128 Then op_vbChar = op_vbChar - 256
    End If
End Function

' (VB) a String of n bytes of b from off, one byte per character (op_strdup_with_len's copy)
Public Function op_vbBytesToStr(b() As Byte, ByVal off As Long, ByVal n As Long) As String
    Dim w() As Byte, i As Long
    If n <= 0 Then op_vbBytesToStr = "": Exit Function
    ReDim w(0 To 2 * n - 1)
    For i = 0 To n - 1
        w(2 * i) = b(off + i)
    Next
    op_vbBytesToStr = w
End Function

' (VB) memcmp(data + off, s, Len(s)) == 0 for an ASCII s
Public Function op_vbMatch(data_() As Byte, ByVal off As Long, ByVal s As String) As Boolean
    Dim i As Long
    For i = 0 To Len(s) - 1
        If data_(off + i) <> Asc(Mid$(s, i + 1, 1)) Then Exit Function
    Next
    op_vbMatch = True
End Function
