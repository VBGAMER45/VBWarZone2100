Attribute VB_Name = "opusfile_c"
' opusfile (lib/sound/3rdparty/opusfile at 9d718345, 4.7.0's submodule), src/opusfile.c: stdio-based
' convenience library for opening/seeking/decoding; whole, as 4.7.0 builds it: the float build
' (OP_FIXED_POINT undefined: its #if branches, op_read's fixed-point filters and OP_STEREO_DOWNMIX_Q14, are not
' compiled), OP_HAVE_LRINTF (op_float2int is lrintf: VB's CLng rounds to even as lrintf does), OP_SOFT_CLIP,
' OP_SMALL_FOOTPRINT undefined, no OP_ENABLE_ASSERTIONS (OP_ASSERT is a comment).
' (VB) An OggOpusFile * is a handle into this module's pool (1-based, 0 = NULL; _ogg_malloc and _ogg_free of
' the struct are op_vbAlloc and op_vbFree); the static functions take the struct ByRef. A callback of
' OpusFileCallbacks is called through opusfile_stream_c's dispatch (op_vbCallRead, ...); op_decode_cb_func and
' op_read_filter_func are numbers. A pointer into a sample buffer is the array and an offset. The float
' arithmetic of the 16-bit conversion (op_float2short_filter: noise shaping, dither) is done one operation per
' statement into Single variables, as C's float rounds each one. The 64-bit arithmetic is Double (exact to 2^53;
' op_rescale64 keeps its 63-bit fraction as bits).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' This implementation is largely based off of libvorbisfile.
'  All of the Ogg bits work roughly the same, though I have made some
'   "improvements" that have not been folded back there, yet.

' A 'chained bitstream' is an Ogg Opus bitstream that contains more than one
'   logical bitstream arranged end to end (the only form of Ogg multiplexing
'   supported by this library.
'  Grouping (parallel multiplexing) is not supported, except to the extent that
'   if there are multiple logical Ogg streams in a single link of the chain, we
'   will ignore all but the first Opus stream we find.

' An Ogg Opus file can be played beginning to end (streamed) without worrying
'   ahead of time about chaining (see opusdec from the opus-tools package).
'  If we have the whole file, however, and want random access
'   (seeking/scrubbing) or desire to know the total length/time of a file, we
'   need to account for the possibility of chaining.

' We can handle things a number of ways.
'  We can determine the entire bitstream structure right off the bat, or find
'   pieces on demand.
'  This library determines and caches structure for the entire bitstream, but
'   builds a virtual decoder on the fly when moving between links in the chain.

' There are also different ways to implement seeking.
'  Enough information exists in an Ogg bitstream to seek to sample-granularity
'   positions in the output.
'  Or, one can seek by picking some portion of the stream roughly in the desired
'   area if we only want coarse navigation through the stream.
'  We implement and expose both strategies.

' The maximum number of bytes in a page (including the page headers).
Private Const OP_PAGE_SIZE_MAX As Long = 65307
' The default amount to seek backwards per step when trying to find the
'   previous page.
'  This must be at least as large as the maximum size of a page.
Private Const OP_CHUNK_SIZE As Long = 65536
' The maximum amount to seek backwards per step when trying to find the
'   previous page.
Private Const OP_CHUNK_SIZE_MAX As Long = 1024& * 1024&
' A smaller read size is needed for low-rate streaming.
Private Const OP_READ_SIZE As Long = 2048

' The minimum granule position spacing allowed for making predictions.
'  This corresponds to about 1 second of audio at 48 kHz for both Opus and
'   Vorbis, or one keyframe interval in Theora with the default keyframe spacing
'   of 256.
Private Const OP_GP_SPACING_MIN As Long = 48000

' This controls how close the target has to be to use the current stream
'   position to subdivide the initial range.
'  Two minutes seems to be a good default.
Private Const OP_CUR_TIME_THRESH As Long = 120& * 48& * 1000&

' The dithering code here is adapted from opusdec, part of opus-tools.
'  It was originally written by Greg Maxwell.
Private Const OP_GAIN As Single = 32753!
Private Const OP_PRNG_GAIN As Single = 1# / 4294967296#     ' (1.0F/(float)0xFFFFFFFF): (float)0xFFFFFFFF is 2^32

' (VB) the op_read_filter_func numbers
Private Const FILTER_op_float2short_filter As Long = 1
Private Const FILTER_op_stereo_filter As Long = 2
Private Const FILTER_op_float2short_stereo_filter As Long = 3

' We use this to remember the pages we found while enumerating the links of a
'   chained stream.
'  We keep track of the starting and ending offsets, as well as the point we
'   started searching from, so we know where to bisect.
'  We also keep the serial number, so we can tell if the page belonged to the
'   current link or not, as well as the granule position, to aid in estimating
'   the start of the link.
Private Type OpusSeekRecord
    ' The earliest byte we know of such that reading forward from it causes
    '   capture to be regained at this page.
    search_start As Double          ' opus_int64
    ' The offset of this page.
    offset As Double
    ' The size of this page.
    size As Long
    ' The serial number of this page.
    serialno As Long                ' ogg_uint32_t
    ' The granule position of this page.
    gp As Double                    ' ogg_int64_t
End Type

' (VB) the OggOpusFile structs (_ogg_malloc'd in the C); ofUsed marks the live ones
Private ofPool() As OggOpusFile
Private ofUsed() As Boolean
Private ofPoolSize As Long

' Matrices for downmixing from the supported channel counts to stereo.
'  The matrices with 5 or more channels are normalized to a total volume of 2.0,
'   since most mixes sound too quiet if normalized to 1.0 (as there is generally
'   little volume in the side/rear channels).
' (VB) OP_STEREO_DOWNMIX[OP_NCHANNELS_MAX-2][OP_NCHANNELS_MAX][2] is OP_STEREO_DOWNMIX(a * 16 + ci * 2 + k)
Private OP_STEREO_DOWNMIX(0 To (OP_NCHANNELS_MAX - 2) * OP_NCHANNELS_MAX * 2 - 1) As Single
' 48 kHz noise shaping filter, sd=2.34.
Private OP_FCOEF_B(0 To 3) As Single
Private OP_FCOEF_A(0 To 3) As Single
Private tablesReady As Boolean

' (VB) the static const tables (VB has no array initialisers)
Private Sub op_vbInitTables()
    Dim t As Variant, i As Long
    ' 3.0
    t = Array(0.5858!, 0!, 0.4142!, 0.4142!, 0!, 0.5858!)
    For i = 0 To UBound(t): OP_STEREO_DOWNMIX(0 * 16 + i) = t(i): Next
    ' quadrophonic
    t = Array(0.4226!, 0!, 0!, 0.4226!, 0.366!, 0.2114!, 0.2114!, 0.336!)
    For i = 0 To UBound(t): OP_STEREO_DOWNMIX(1 * 16 + i) = t(i): Next
    ' 5.0
    t = Array(0.651!, 0!, 0.46!, 0.46!, 0!, 0.651!, 0.5636!, 0.3254!, 0.3254!, 0.5636!)
    For i = 0 To UBound(t): OP_STEREO_DOWNMIX(2 * 16 + i) = t(i): Next
    ' 5.1
    t = Array(0.529!, 0!, 0.3741!, 0.3741!, 0!, 0.529!, 0.4582!, 0.2645!, 0.2645!, 0.4582!, 0.3741!, 0.3741!)
    For i = 0 To UBound(t): OP_STEREO_DOWNMIX(3 * 16 + i) = t(i): Next
    ' 6.1
    t = Array(0.4553!, 0!, 0.322!, 0.322!, 0!, 0.4553!, 0.3943!, 0.2277!, 0.2277!, 0.3943!, 0.2788!, 0.2788!, 0.322!, 0.322!)
    For i = 0 To UBound(t): OP_STEREO_DOWNMIX(4 * 16 + i) = t(i): Next
    ' 7.1
    t = Array(0.3886!, 0!, 0.2748!, 0.2748!, 0!, 0.3886!, 0.3366!, 0.1943!, 0.1943!, 0.3366!, 0.3366!, 0.1943!, 0.1943!, 0.3366!, 0.2748!, 0.2748!)
    For i = 0 To UBound(t): OP_STEREO_DOWNMIX(5 * 16 + i) = t(i): Next
    OP_FCOEF_B(0) = 2.2374!: OP_FCOEF_B(1) = -0.7339!: OP_FCOEF_B(2) = -0.1251!: OP_FCOEF_B(3) = -0.6033!
    OP_FCOEF_A(0) = 0.903!: OP_FCOEF_A(1) = 0.0116!: OP_FCOEF_A(2) = -0.5853!: OP_FCOEF_A(3) = -0.2571!
    tablesReady = True
End Sub

' (VB) _ogg_malloc(sizeof(OggOpusFile)): a free slot of the pool
Private Function op_vbAlloc() As Long
    Dim i As Long
    For i = 1 To ofPoolSize
        If Not ofUsed(i) Then Exit For
    Next
    If i > ofPoolSize Then
        ofPoolSize = ofPoolSize + 4
        ReDim Preserve ofPool(1 To ofPoolSize)
        ReDim Preserve ofUsed(1 To ofPoolSize)
    End If
    ofUsed(i) = True
    op_vbAlloc = i
End Function

' (VB) _ogg_free(of)
Private Sub op_vbFree(ByVal of_ As Long)
    Dim blank As OggOpusFile
    ofPool(of_) = blank
    ofUsed(of_) = False
End Sub

' (VB) whether a void * stream is NULL (a Long 0, Empty or Nothing)
Private Function op_vbIsNull(stream_ As Variant) As Boolean
    If IsObject(stream_) Then
        op_vbIsNull = (stream_ Is Nothing)
    ElseIf IsEmpty(stream_) Then
        op_vbIsNull = True
    Else
        op_vbIsNull = (stream_ = 0)
    End If
End Function

' (VB) a = b for a void * (an object is Set)
Private Sub op_vbSetVoid(a As Variant, b As Variant)
    If IsObject(b) Then Set a = b Else a = b
End Sub

Public Function op_test(head_ As OpusHead, initial_data() As Byte, ByVal initial_bytes As Long, _
                        Optional ByVal headIsNull As Boolean = False) As Long
    Dim oy As ogg_sync_state, data_ As Long, err_ As Long, os As ogg_stream_state, og As ogg_page, ret As Long
    Dim op As ogg_packet
    ' The first page of a normal Opus file will be at most 57 bytes (27 Ogg
    '   page header bytes + 1 lacing value + 21 Opus header bytes + 8 channel
    '   mapping bytes).
    '  It will be at least 47 bytes (27 Ogg page header bytes + 1 lacing value +
    '   19 Opus header bytes using channel mapping family 0).
    '  If we don't have at least that much data, give up now.
    If initial_bytes < 47 Then op_test = OP_FALSE: Exit Function
    ' Only proceed if we start with the magic OggS string.
    '  This is to prevent us spending a lot of time allocating memory and looking
    '   for Ogg pages in non-Ogg files.
    If initial_data(0) <> 79 Or initial_data(1) <> 103 Or initial_data(2) <> 103 Or initial_data(3) <> 83 Then
        op_test = OP_ENOTFORMAT
        Exit Function
    End If
    ' (if(OP_UNLIKELY(_initial_bytes>(size_t)LONG_MAX))return OP_EFAULT: a Long)
    ogg_sync_init oy
    data_ = ogg_sync_buffer(oy, initial_bytes)
    If data_ >= 0 Then
        CopyMemory oy.data(data_), initial_data(0), initial_bytes
        ogg_sync_wrote oy, initial_bytes
        ogg_stream_init os, -1
        err_ = OP_FALSE
        Do
            ret = ogg_sync_pageout(oy, og)
            ' Ignore holes.
            If ret < 0 Then GoTo continue_
            ' Stop if we run out of data.
            If ret = 0 Then Exit Do
            ogg_stream_reset_serialno os, ogg_page_serialno(og)
            ogg_stream_pagein os, og
            ' Only process the first packet on this page (if it's a BOS packet,
            '   it's required to be the only one).
            If ogg_stream_packetout(os, op) = 1 Then
                If op.b_o_s Then
                    ret = opus_head_parse(head_, op.packet, 0, op.bytes, headIsNull)
                    ' If this didn't look like Opus, keep going.
                    If ret = OP_ENOTFORMAT Then GoTo continue_
                    ' Otherwise we're done, one way or another.
                    err_ = ret
                Else
                    ' We finished parsing the headers.
                    '  There is no Opus to be found.
                    err_ = OP_ENOTFORMAT
                End If
            End If
continue_:
        Loop While err_ = OP_FALSE
        ogg_stream_clear os
    Else
        err_ = OP_EFAULT
    End If
    ogg_sync_clear oy
    op_test = err_
End Function

' Many, many internal helpers.
'  The intention is not to be confusing.
'  Rampant duplication and monolithic function implementation (though we do have
'   some large, omnibus functions still) would be harder to understand anyway.
'  The high level functions are last.
'  Begin grokking near the end of the file if you prefer to read things
'   top-down.

' The read/seek functions track absolute position within the stream.

' Read a little more data from the file/pipe into the ogg_sync framer.
'  _nbytes: The maximum number of bytes to read.
'  Return: A positive number of bytes read on success, 0 on end-of-file, or a
'           negative value on failure.
Private Function op_get_data(of_ As OggOpusFile, ByVal nbytes_ As Long) As Long
    Dim buffer As Long, nbytes As Long
    ' (OP_ASSERT(_nbytes>0))
    buffer = ogg_sync_buffer(of_.oy, nbytes_)
    If buffer < 0 Then op_get_data = OP_EFAULT: Exit Function
    nbytes = op_vbCallRead(of_.callbacks.read_, of_.stream, of_.oy.data, buffer, nbytes_)
    ' (OP_ASSERT(nbytes<=_nbytes))
    If nbytes > 0 Then ogg_sync_wrote of_.oy, nbytes
    op_get_data = nbytes
End Function

' Save a tiny smidge of verbosity to make the code more readable.
Private Function op_seek_helper(of_ As OggOpusFile, ByVal offset_ As Double) As Long
    If offset_ = of_.offset Then op_seek_helper = 0: Exit Function
    If of_.callbacks.seek_ = 0 Then op_seek_helper = OP_EREAD: Exit Function
    If op_vbCallSeek(of_.callbacks.seek_, of_.stream, offset_, SEEK_SET) Then op_seek_helper = OP_EREAD: Exit Function
    of_.offset = offset_
    ogg_sync_reset of_.oy
    op_seek_helper = 0
End Function

' Get the current position indicator of the underlying stream.
'  This should be the same as the value reported by tell().
Private Function op_position(of_ As OggOpusFile) As Double
    ' The current position indicator is _not_ simply offset.
    '  We may also have unprocessed, buffered data in the sync state.
    op_position = of_.offset + of_.oy.fill - of_.oy.returned
End Function

' From the head of the stream, get the next page.
'  _boundary specifies if the function is allowed to fetch more data from the
'   stream (and how much) or only use internally buffered data.
'  _boundary: -1: Unbounded search.
'              0: Read no additional data.
'                 Use only cached data.
'              n: Search for the start of a new page up to file position n.
'  Return: n>=0:       Found a page at absolute offset n.
'          OP_FALSE:   Hit the _boundary limit.
'          OP_EREAD:   An underlying read operation failed.
'          OP_BADLINK: We hit end-of-file before reaching _boundary.
Private Function op_get_next_page(of_ As OggOpusFile, og_ As ogg_page, ByVal boundary_ As Double) As Double
    Dim more As Long, read_nbytes As Long, ret As Long, position As Double, page_offset As Double
    Do While boundary_ <= 0 Or of_.offset < boundary_
        more = ogg_sync_pageseek(of_.oy, og_)
        ' Skipped (-more) bytes.
        If more < 0 Then
            of_.offset = of_.offset - more
        ElseIf more = 0 Then
            ' Send more paramedics.
            If boundary_ = 0 Then op_get_next_page = OP_FALSE: Exit Function
            If boundary_ < 0 Then
                read_nbytes = OP_READ_SIZE
            Else
                position = op_position(of_)
                If position >= boundary_ Then op_get_next_page = OP_FALSE: Exit Function
                If boundary_ - position < OP_READ_SIZE Then read_nbytes = boundary_ - position Else read_nbytes = OP_READ_SIZE
            End If
            ret = op_get_data(of_, read_nbytes)
            If ret < 0 Then op_get_next_page = OP_EREAD: Exit Function
            If ret = 0 Then
                ' Only fail cleanly on EOF if we didn't have a known boundary.
                '  Otherwise, we should have been able to reach that boundary, and this
                '   is a fatal error.
                If boundary_ < 0 Then op_get_next_page = OP_FALSE Else op_get_next_page = OP_EBADLINK
                Exit Function
            End If
        Else
            ' Got a page.
            '  Return the page start offset and advance the internal offset past the
            '   page end.
            page_offset = of_.offset
            of_.offset = of_.offset + more
            ' (OP_ASSERT(page_offset>=0))
            op_get_next_page = page_offset
            Exit Function
        End If
    Loop
    op_get_next_page = OP_FALSE
End Function

Private Function op_add_serialno(og_ As ogg_page, serialnos_() As Long, nserialnos_ As Long, cserialnos_ As Long) As Long
    Dim nserialnos As Long, cserialnos As Long, s As Long
    s = ogg_page_serialno(og_)
    nserialnos = nserialnos_
    cserialnos = cserialnos_
    If nserialnos >= cserialnos Then
        If cserialnos > (2147483647 \ 4 - 1) \ 2 Then op_add_serialno = OP_EFAULT: Exit Function
        cserialnos = 2 * cserialnos + 1
        ' (OP_ASSERT(nserialnos<cserialnos))
        ' serialnos=(ogg_uint32_t *)_ogg_realloc(serialnos,sizeof(*serialnos)*cserialnos)
        If cserialnos_ > 0 Then ReDim Preserve serialnos_(0 To cserialnos - 1) Else ReDim serialnos_(0 To cserialnos - 1)
    End If
    serialnos_(nserialnos) = s
    nserialnos = nserialnos + 1
    nserialnos_ = nserialnos
    cserialnos_ = cserialnos
    op_add_serialno = 0
End Function

' Returns nonzero if found.
Private Function op_lookup_serialno(ByVal s As Long, serialnos_() As Long, ByVal nserialnos_ As Long) As Boolean
    Dim i As Long
    For i = 0 To nserialnos_ - 1
        If serialnos_(i) = s Then Exit For
    Next
    op_lookup_serialno = (i < nserialnos_)
End Function

Private Function op_lookup_page_serialno(og_ As ogg_page, serialnos_() As Long, ByVal nserialnos_ As Long) As Boolean
    op_lookup_page_serialno = op_lookup_serialno(ogg_page_serialno(og_), serialnos_, nserialnos_)
End Function

' Find the last page beginning before _offset with a valid granule position.
'  There is no '_boundary' parameter as it will always have to read more data.
'  This is much dirtier than the above, as Ogg doesn't have any backward search
'   linkage.
'  This search prefers pages of the specified serial number.
'  If a page of the specified serial number is spotted during the
'   seek-back-and-read-forward, it will return the info of last page of the
'   matching serial number, instead of the very last page, unless the very last
'   page belongs to a different link than preferred serial number.
'  If no page of the specified serial number is seen, it will return the info of
'   the last page.
'  [out] _sr:   Returns information about the page that was found on success.
'  _offset:     The _offset before which to find a page.
'               Any page returned will consist of data entirely before _offset.
'  _serialno:   The preferred serial number.
'               If a page with this serial number is found, it will be returned
'                even if another page in the same link is found closer to
'                _offset.
'               This is purely opportunistic: there is no guarantee such a page
'                will be found if it exists.
'  _serialnos:  The list of serial numbers in the link that contains the
'                preferred serial number.
'  _nserialnos: The number of serial numbers in the current link.
'  Return: 0 on success, or a negative value on failure.
'          OP_EREAD:    Failed to read more data (error or EOF).
'          OP_EBADLINK: We couldn't find a page even after seeking back to the
'                        start of the stream.
Private Function op_get_prev_page_serial(of_ As OggOpusFile, sr_ As OpusSeekRecord, ByVal offset_ As Double, _
                                         ByVal serialno_ As Long, serialnos_() As Long, ByVal nserialnos_ As Long) As Long
    Dim preferred_sr As OpusSeekRecord, og As ogg_page, begin As Double, end_ As Double, original_end As Double
    Dim chunk_size As Long, preferred_found As Long, search_start As Double, ret As Long, llret As Double
    Dim serialno As Long
    end_ = offset_
    begin = offset_
    original_end = offset_
    preferred_found = 0
    offset_ = -1
    chunk_size = OP_CHUNK_SIZE
    Do
        ' (OP_ASSERT(chunk_size>=OP_PAGE_SIZE_MAX))
        If begin - chunk_size > 0 Then begin = begin - chunk_size Else begin = 0
        ret = op_seek_helper(of_, begin)
        If ret < 0 Then op_get_prev_page_serial = ret: Exit Function
        search_start = begin
        Do While of_.offset < end_
            llret = op_get_next_page(of_, og, end_)
            If llret < OP_FALSE Then
                op_get_prev_page_serial = llret
                Exit Function
            ElseIf llret = OP_FALSE Then
                Exit Do
            End If
            serialno = ogg_page_serialno(og)
            ' Save the information for this page.
            '  We're not interested in the page itself... just the serial number, byte
            '   offset, page size, and granule position.
            sr_.search_start = search_start
            offset_ = llret
            sr_.offset = offset_
            sr_.serialno = serialno
            ' (OP_ASSERT(_of->offset-_offset>=0); OP_ASSERT(_of->offset-_offset<=OP_PAGE_SIZE_MAX))
            sr_.size = of_.offset - offset_
            sr_.gp = ogg_page_granulepos(og)
            ' If this page is from the stream we're looking for, remember it.
            If serialno = serialno_ Then
                preferred_found = 1
                preferred_sr = sr_
            End If
            If Not op_lookup_serialno(serialno, serialnos_, nserialnos_) Then
                ' We fell off the end of the link, which means we seeked back too far
                '   and shouldn't have been looking in that link to begin with.
                '  If we found the preferred serial number, forget that we saw it.
                preferred_found = 0
            End If
            search_start = llret + 1
        Loop
        ' We started from the beginning of the stream and found nothing.
        '  This should be impossible unless the contents of the stream changed out
        '   from under us after we read from it.
        If begin = 0 And offset_ < 0 Then op_get_prev_page_serial = OP_EBADLINK: Exit Function
        ' Bump up the chunk size.
        '  This is mildly helpful when seeks are very expensive (http).
        If 2 * chunk_size < OP_CHUNK_SIZE_MAX Then chunk_size = 2 * chunk_size Else chunk_size = OP_CHUNK_SIZE_MAX
        ' Avoid quadratic complexity if we hit an invalid patch of the file.
        If begin + OP_PAGE_SIZE_MAX - 1 < original_end Then end_ = begin + OP_PAGE_SIZE_MAX - 1 Else end_ = original_end
    Loop While offset_ < 0
    If preferred_found Then sr_ = preferred_sr
    op_get_prev_page_serial = 0
End Function

' Find the last page beginning before _offset with the given serial number and
'   a valid granule position.
'  Unlike the above search, this continues until it finds such a page, but does
'   not stray outside the current link.
'  We could implement it (inefficiently) by calling op_get_prev_page_serial()
'   repeatedly until it returned a page that had both our preferred serial
'   number and a valid granule position, but doing it with a separate function
'   allows us to avoid repeatedly re-scanning valid pages from other streams as
'   we seek-back-and-read-forward.
'  [out] _gp:   Returns the granule position of the page that was found on
'                success.
'  _offset:     The _offset before which to find a page.
'               Any page returned will consist of data entirely before _offset.
'  _serialno:   The target serial number.
'  _serialnos:  The list of serial numbers in the link that contains the
'                preferred serial number.
'  _nserialnos: The number of serial numbers in the current link.
'  Return: The offset of the page on success, or a negative value on failure.
'          OP_EREAD:    Failed to read more data (error or EOF).
'          OP_EBADLINK: We couldn't find a page even after seeking back past the
'                        beginning of the link.
Private Function op_get_last_page(of_ As OggOpusFile, gp_ As Double, ByVal offset_ As Double, ByVal serialno_ As Long, _
                                  serialnos_() As Long, ByVal nserialnos_ As Long) As Double
    Dim og As ogg_page, gp As Double, begin As Double, end_ As Double, original_end As Double, chunk_size As Long
    Dim left_link As Long, ret As Long, llret As Double, serialno As Long, page_gp As Double
    ' The target serial number must belong to the current link.
    ' (OP_ASSERT(op_lookup_serialno(_serialno,_serialnos,_nserialnos)))
    end_ = offset_
    begin = offset_
    original_end = offset_
    offset_ = -1
    ' We shouldn't have to initialize gp, but gcc is too dumb to figure out that
    '   ret>=0 implies we entered the if(page_gp!=-1) block at least once.
    gp = -1
    chunk_size = OP_CHUNK_SIZE
    Do
        ' (OP_ASSERT(chunk_size>=OP_PAGE_SIZE_MAX))
        If begin - chunk_size > 0 Then begin = begin - chunk_size Else begin = 0
        ret = op_seek_helper(of_, begin)
        If ret < 0 Then op_get_last_page = ret: Exit Function
        left_link = 0
        Do While of_.offset < end_
            llret = op_get_next_page(of_, og, end_)
            If llret < OP_FALSE Then
                op_get_last_page = llret
                Exit Function
            ElseIf llret = OP_FALSE Then
                Exit Do
            End If
            serialno = ogg_page_serialno(og)
            If serialno = serialno_ Then
                ' The page is from the right stream...
                page_gp = ogg_page_granulepos(og)
                If page_gp <> -1 Then
                    ' And has a valid granule position.
                    '  Let's remember it.
                    offset_ = llret
                    gp = page_gp
                End If
            ElseIf Not op_lookup_serialno(serialno, serialnos_, nserialnos_) Then
                ' We fell off the start of the link, which means we don't need to keep
                '   seeking any farther back.
                left_link = 1
            End If
        Loop
        ' We started from at or before the beginning of the link and found nothing.
        '  This should be impossible unless the contents of the stream changed out
        '   from under us after we read from it.
        If (left_link <> 0 Or begin = 0) And offset_ < 0 Then
            op_get_last_page = OP_EBADLINK
            Exit Function
        End If
        ' Bump up the chunk size.
        '  This is mildly helpful when seeks are very expensive (http).
        If 2 * chunk_size < OP_CHUNK_SIZE_MAX Then chunk_size = 2 * chunk_size Else chunk_size = OP_CHUNK_SIZE_MAX
        ' Avoid quadratic complexity if we hit an invalid patch of the file.
        If begin + OP_PAGE_SIZE_MAX - 1 < original_end Then end_ = begin + OP_PAGE_SIZE_MAX - 1 Else end_ = original_end
    Loop While offset_ < 0
    gp_ = gp
    op_get_last_page = offset_
End Function

' Uses the local ogg_stream storage in _of.
'  This is important for non-streaming input sources.
' (VB) _serialnos NULL is serialnosIsNull
Private Function op_fetch_headers_impl(of_ As OggOpusFile, head_ As OpusHead, tags_ As OpusTags, serialnos_() As Long, _
                                       nserialnos_ As Long, cserialnos_ As Long, ByVal serialnosIsNull As Boolean, _
                                       og_ As ogg_page) As Long
    Dim op As ogg_packet, ret As Long
    If Not serialnosIsNull Then nserialnos_ = 0
    ' Extract the serialnos of all BOS pages plus the first set of Opus headers
    '   we see in the link.
    Do While ogg_page_bos(og_)
        If Not serialnosIsNull Then
            If op_lookup_page_serialno(og_, serialnos_, nserialnos_) Then
                ' A dupe serialnumber in an initial header packet set==invalid stream.
                op_fetch_headers_impl = OP_EBADHEADER
                Exit Function
            End If
            ret = op_add_serialno(og_, serialnos_, nserialnos_, cserialnos_)
            If ret < 0 Then op_fetch_headers_impl = ret: Exit Function
        End If
        If of_.ready_state < OP_STREAMSET Then
            ' We don't have an Opus stream in this link yet, so begin prospective
            '   stream setup.
            '  We need a stream to get packets.
            ogg_stream_reset_serialno of_.os, ogg_page_serialno(og_)
            ogg_stream_pagein of_.os, og_
            If ogg_stream_packetout(of_.os, op) > 0 Then
                ret = opus_head_parse(head_, op.packet, 0, op.bytes)
                ' Found a valid Opus header.
                '  Continue setup.
                If ret >= 0 Then
                    of_.ready_state = OP_STREAMSET
                ' If it's just a stream type we don't recognize, ignore it.
                '  Everything else is fatal.
                ElseIf ret <> OP_ENOTFORMAT Then
                    op_fetch_headers_impl = ret
                    Exit Function
                End If
            End If
            ' TODO: Should a BOS page with no packets be an error?
        End If
        ' Get the next page.
        '  No need to clamp the boundary offset against _of->end, as all errors
        '   become OP_ENOTFORMAT or OP_EBADHEADER.
        If op_get_next_page(of_, og_, OP_ADV_OFFSET(of_.offset, OP_CHUNK_SIZE)) < 0 Then
            If of_.ready_state < OP_STREAMSET Then op_fetch_headers_impl = OP_ENOTFORMAT Else op_fetch_headers_impl = OP_EBADHEADER
            Exit Function
        End If
    Loop
    If of_.ready_state <> OP_STREAMSET Then op_fetch_headers_impl = OP_ENOTFORMAT: Exit Function
    ' If the first non-header page belonged to our Opus stream, submit it.
    If of_.os.serialno = ogg_page_serialno(og_) Then ogg_stream_pagein of_.os, og_
    ' Loop getting packets.
    Do
        Select Case ogg_stream_packetout(of_.os, op)
        Case 0
            ' Loop getting pages.
            Do
                ' No need to clamp the boundary offset against _of->end, as all
                '   errors become OP_EBADHEADER.
                If op_get_next_page(of_, og_, OP_ADV_OFFSET(of_.offset, OP_CHUNK_SIZE)) < 0 Then
                    op_fetch_headers_impl = OP_EBADHEADER
                    Exit Function
                End If
                ' If this page belongs to the correct stream, go parse it.
                If of_.os.serialno = ogg_page_serialno(og_) Then
                    ogg_stream_pagein of_.os, og_
                    Exit Do
                End If
                ' If the link ends before we see the Opus comment header, abort.
                If ogg_page_bos(og_) Then op_fetch_headers_impl = OP_EBADHEADER: Exit Function
                ' Otherwise, keep looking.
            Loop
        ' We shouldn't get a hole in the headers!
        Case -1
            op_fetch_headers_impl = OP_EBADHEADER
            Exit Function
        Case Else
            ' Got a packet.
            '  It should be the comment header.
            ret = opus_tags_parse(tags_, op.packet, 0, op.bytes)
            If ret < 0 Then op_fetch_headers_impl = ret: Exit Function
            ' Make sure the page terminated at the end of the comment header.
            '  If there is another packet on the page, or part of a packet, then
            '   reject the stream.
            '  Otherwise seekable sources won't be able to seek back to the start
            '   properly.
            ret = ogg_stream_packetout(of_.os, op)
            If ret <> 0 Or og_.header(og_.header_len - 1) = 255 Then
                ' If we fail, the caller assumes our tags are uninitialized.
                opus_tags_clear tags_
                op_fetch_headers_impl = OP_EBADHEADER
                Exit Function
            End If
            op_fetch_headers_impl = 0
            Exit Function
        End Select
    Loop
End Function

' (VB) _og NULL is ogIsNull
Private Function op_fetch_headers(of_ As OggOpusFile, head_ As OpusHead, tags_ As OpusTags, serialnos_() As Long, _
                                  nserialnos_ As Long, cserialnos_ As Long, ByVal serialnosIsNull As Boolean, _
                                  og_ As ogg_page, ByVal ogIsNull As Boolean) As Long
    Dim og As ogg_page, ret As Long
    If ogIsNull Then
        ' No need to clamp the boundary offset against _of->end, as all errors
        '   become OP_ENOTFORMAT.
        If op_get_next_page(of_, og, OP_ADV_OFFSET(of_.offset, OP_CHUNK_SIZE)) < 0 Then
            op_fetch_headers = OP_ENOTFORMAT
            Exit Function
        End If
        ' (_og=&og)
        of_.ready_state = OP_OPENED
        ret = op_fetch_headers_impl(of_, head_, tags_, serialnos_, nserialnos_, cserialnos_, serialnosIsNull, og)
    Else
        of_.ready_state = OP_OPENED
        ret = op_fetch_headers_impl(of_, head_, tags_, serialnos_, nserialnos_, cserialnos_, serialnosIsNull, og_)
    End If
    ' Revert back from OP_STREAMSET to OP_OPENED on failure, to prevent
    '   double-free of the tags in an unseekable stream.
    If ret < 0 Then of_.ready_state = OP_OPENED
    op_fetch_headers = ret
End Function

' Granule position manipulation routines.
'  A granule position is defined to be an unsigned 64-bit integer, with the
'   special value -1 in two's complement indicating an unset or invalid granule
'   position.
'  We are not guaranteed to have an unsigned 64-bit type, so we construct the
'   following routines that
'   a) Properly order negative numbers as larger than positive numbers, and
'   b) Check for underflow or overflow past the special -1 value.
'  This lets us operate on the full, valid range of granule positions in a
'   consistent and safe manner.
'  This full range is organized into distinct regions:
'   [ -1 (invalid) ][ 0 ... OP_INT64_MAX ][ OP_INT64_MIN ... -2 ][-1 (invalid) ]
'
'  No one should actually use granule positions so large that they're negative,
'   even if they are technically valid, as very little software handles them
'   correctly (including most of Xiph.Org's).
'  This library also refuses to support durations so large they won't fit in a
'   signed 64-bit integer (to avoid exposing this mess to the application, and
'   to simplify a good deal of internal arithmetic), so the only way to use them
'   successfully is if pcm_start is very large.
'  This means there isn't anything you can do with negative granule positions
'   that you couldn't have done with purely non-negative ones.
'  The main purpose of these routines is to allow us to think very explicitly
'   about the possible failure cases of all granule position manipulations.

' Safely adds a small signed integer to a valid (not -1) granule position.
'  The result can use the full 64-bit range of values (both positive and
'   negative), but will fail on overflow (wrapping past -1; wrapping past
'   OP_INT64_MAX is explicitly okay).
'  [out] _dst_gp: The resulting granule position.
'                 Only modified on success.
'  _src_gp:       The granule position to add to.
'                 This must not be -1.
'  _delta:        The amount to add.
'                 This is allowed to be up to 32 bits to support the maximum
'                  duration of a single Ogg page (255 packets * 120 ms per
'                  packet == 1,468,800 samples at 48 kHz).
'  Return: 0 on success, or OP_EINVAL if the result would wrap around past -1.
Private Function op_granpos_add(dst_gp_ As Double, ByVal src_gp_ As Double, ByVal delta_ As Double) As Long
    ' The code below handles this case correctly, but there's no reason we
    '   should ever be called with these values, so make sure we aren't.
    ' (OP_ASSERT(_src_gp!=-1))
    If delta_ > 0 Then
        ' Adding this amount to the granule position would overflow its 64-bit
        '   range.
        If src_gp_ < 0 And src_gp_ >= -1 - delta_ Then op_granpos_add = OP_EINVAL: Exit Function
        If src_gp_ > OP_INT64_MAX - delta_ Then
            ' Adding this amount to the granule position would overflow the positive
            '   half of its 64-bit range.
            '  Since signed overflow is undefined in C, do it in a way the compiler
            '   isn't allowed to screw up.
            delta_ = delta_ - ((OP_INT64_MAX - src_gp_) + 1)
            src_gp_ = OP_INT64_MIN
        End If
    ElseIf delta_ < 0 Then
        ' Subtracting this amount from the granule position would underflow its
        '   64-bit range.
        If src_gp_ >= 0 And src_gp_ < -delta_ Then op_granpos_add = OP_EINVAL: Exit Function
        If src_gp_ < OP_INT64_MIN - delta_ Then
            ' Subtracting this amount from the granule position would underflow the
            '   negative half of its 64-bit range.
            '  Since signed underflow is undefined in C, do it in a way the compiler
            '   isn't allowed to screw up.
            delta_ = delta_ + ((src_gp_ - OP_INT64_MIN) + 1)
            src_gp_ = OP_INT64_MAX
        End If
    End If
    dst_gp_ = src_gp_ + delta_
    op_granpos_add = 0
End Function

' Safely computes the difference between two granule positions.
'  The difference must fit in a signed 64-bit integer, or the function fails.
'  It correctly handles the case where the granule position has wrapped around
'   from positive values to negative ones.
'  [out] _delta: The difference between the granule positions.
'                Only modified on success.
'  _gp_a:        The granule position to subtract from.
'                This must not be -1.
'  _gp_b:        The granule position to subtract.
'                This must not be -1.
'  Return: 0 on success, or OP_EINVAL if the result would not fit in a signed
'           64-bit integer.
Private Function op_granpos_diff(delta_ As Double, ByVal gp_a_ As Double, ByVal gp_b_ As Double) As Long
    Dim gp_a_negative As Boolean, gp_b_negative As Boolean, da As Double, db As Double
    ' The code below handles these cases correctly, but there's no reason we
    '   should ever be called with these values, so make sure we aren't.
    ' (OP_ASSERT(_gp_a!=-1); OP_ASSERT(_gp_b!=-1))
    gp_a_negative = (gp_a_ < 0)
    gp_b_negative = (gp_b_ < 0)
    If gp_a_negative Xor gp_b_negative Then
        If gp_a_negative Then
            ' _gp_a has wrapped to a negative value but _gp_b hasn't: the difference
            '   should be positive.
            ' Step 1: Handle wrapping.
            ' _gp_a < 0 => da < 0.
            da = (OP_INT64_MIN - gp_a_) - 1
            ' _gp_b >= 0  => db >= 0.
            db = OP_INT64_MAX - gp_b_
            ' Step 2: Check for overflow.
            If OP_INT64_MAX + da < db Then op_granpos_diff = OP_EINVAL: Exit Function
            delta_ = db - da
        Else
            ' _gp_b has wrapped to a negative value but _gp_a hasn't: the difference
            '   should be negative.
            ' Step 1: Handle wrapping.
            ' _gp_a >= 0 => da <= 0
            da = gp_a_ + OP_INT64_MIN
            ' _gp_b < 0 => db <= 0
            db = OP_INT64_MIN - gp_b_
            ' Step 2: Check for overflow.
            If da < OP_INT64_MIN - db Then op_granpos_diff = OP_EINVAL: Exit Function
            delta_ = da + db
        End If
    Else
        delta_ = gp_a_ - gp_b_
    End If
    op_granpos_diff = 0
End Function

Private Function op_granpos_cmp(ByVal gp_a_ As Double, ByVal gp_b_ As Double) As Long
    ' The invalid granule position -1 should behave like NaN: neither greater
    '   than nor less than any other granule position, nor equal to any other
    '   granule position, including itself.
    '  However, that means there isn't anything we could sensibly return from this
    '   function for it.
    ' (OP_ASSERT(_gp_a!=-1); OP_ASSERT(_gp_b!=-1))
    ' Handle the wrapping cases.
    If gp_a_ < 0 Then
        If gp_b_ >= 0 Then op_granpos_cmp = 1: Exit Function
        ' Else fall through.
    ElseIf gp_b_ < 0 Then
        op_granpos_cmp = -1
        Exit Function
    End If
    ' No wrapping case.
    op_granpos_cmp = -(gp_a_ > gp_b_) - -(gp_b_ > gp_a_)
End Function

' Returns the duration of the packet (in samples at 48 kHz), or a negative
'   value on error.
Private Function op_get_packet_duration(data_() As Byte, ByVal dataOff As Long, ByVal len_ As Long) As Long
    Dim nframes As Long, frame_size As Long, nsamples As Long
    nframes = opus_packet_get_nb_frames(data_, dataOff, len_)
    If nframes < 0 Then op_get_packet_duration = OP_EBADPACKET: Exit Function
    frame_size = opus_packet_get_samples_per_frame(data_, dataOff, 48000)
    nsamples = nframes * frame_size
    If nsamples > 120 * 48 Then op_get_packet_duration = OP_EBADPACKET: Exit Function
    op_get_packet_duration = nsamples
End Function

' This function more properly belongs in info.c, but we define it here to allow
'   the static granule position manipulation functions to remain static.
Public Function opus_granule_sample(head_ As OpusHead, ByVal gp_ As Double) As Double
    Dim pre_skip As Long
    pre_skip = head_.pre_skip
    If gp_ <> -1 Then
        If op_granpos_add(gp_, gp_, -pre_skip) Then gp_ = -1
    End If
    opus_granule_sample = gp_
End Function

' Grab all the packets currently in the stream state, and compute their
'   durations.
'  _of->op_count is set to the number of packets collected.
'  [out] _durations: Returns the durations of the individual packets.
'  Return: The total duration of all packets, or OP_HOLE if there was a hole.
Private Function op_collect_audio_packets(of_ As OggOpusFile, durations_() As Long) As Long
    Dim total_duration As Long, op_count As Long, ret As Long
    ' Count the durations of all packets in the page.
    op_count = 0
    total_duration = 0
    Do
        ' This takes advantage of undocumented libogg behavior that returned
        '   ogg_packet buffers are valid at least until the next page is
        '   submitted.
        '  Relying on this is not too terrible, as _none_ of the Ogg memory
        '   ownership/lifetime rules are well-documented.
        '  But I can read its code and know this will work.
        ' (VB) a page completes at most 255 packets (op_count < 255 below), so _of->op+op_count stays in
        ' the array whenever there is a packet to return
        If op_count >= 255 Then Exit Do
        ret = ogg_stream_packetout(of_.os, of_.op(op_count))
        If ret = 0 Then Exit Do
        If ret < 0 Then
            ' We shouldn't get holes in the middle of pages.
            ' (OP_ASSERT(op_count==0))
            ' Set the return value and break out of the loop.
            '  We want to make sure op_count gets set to 0, because we've ingested a
            '   page, so any previously loaded packets are now invalid.
            total_duration = OP_HOLE
            Exit Do
        End If
        ' Unless libogg is broken, we can't get more than 255 packets from a
        '   single page.
        ' (OP_ASSERT(op_count<255))
        durations_(op_count) = op_get_packet_duration(of_.op(op_count).packet, 0, of_.op(op_count).bytes)
        If durations_(op_count) > 0 Then
            ' With at most 255 packets on a page, this can't overflow.
            total_duration = total_duration + durations_(op_count)
            op_count = op_count + 1
        ' Ignore packets with an invalid TOC sequence.
        ElseIf op_count > 0 Then
            ' But save the granule position, if there was one.
            of_.op(op_count - 1).granulepos = of_.op(op_count).granulepos
        End If
    Loop
    of_.op_pos = 0
    of_.op_count = op_count
    op_collect_audio_packets = total_duration
End Function

' Starting from current cursor position, get the initial PCM offset of the next
'   page.
'  This also validates the granule position on the first page with a completed
'   audio data packet, as required by the spec.
'  If this link is completely empty (no pages with completed packets), then this
'   function sets pcm_start=pcm_end=0 and returns the BOS page of the next link
'   (if any).
'  In the seekable case, we initialize pcm_end=-1 before calling this function,
'   so that later we can detect that the link was empty before calling
'   op_find_final_pcm_offset().
'  [inout] _link: The link for which to find pcm_start.
'  [out] _og:     Returns the BOS page of the next link if this link was empty.
'                 In the unseekable case, we can then feed this to
'                  op_fetch_headers() to start the next link.
'                 The caller may pass NULL (e.g., for seekable streams), in
'                  which case this page will be discarded.
'                 ((VB) such a caller passes a page of its own)
'  Return: 0 on success, 1 if there is a buffered BOS page available, or a
'           negative value on unrecoverable error.
Private Function op_find_initial_pcm_offset(of_ As OggOpusFile, link_ As OggOpusLink, og_ As ogg_page) As Long
    Dim page_offset As Double, pcm_start As Double, prev_packet_gp As Double, cur_page_gp As Double
    Dim serialno As Long, total_duration As Long, durations(0 To 254) As Long, cur_page_eos As Long
    Dim op_count As Long, pi As Long, diff As Double
    serialno = of_.os.serialno
    op_count = 0
    ' We shouldn't have to initialize total_duration, but gcc is too dumb to
    '   figure out that op_count>0 implies we've been through the whole loop at
    '   least once.
    total_duration = 0
    Do
        page_offset = op_get_next_page(of_, og_, of_.end_)
        ' We should get a page unless the file is truncated or mangled.
        '  Otherwise there are no audio data packets in the whole logical stream.
        If page_offset < 0 Then
            ' Fail if there was a read error.
            If page_offset < OP_FALSE Then op_find_initial_pcm_offset = page_offset: Exit Function
            ' Fail if the pre-skip is non-zero, since it's asking us to skip more
            '   samples than exist.
            If link_.head.pre_skip > 0 Then op_find_initial_pcm_offset = OP_EBADTIMESTAMP: Exit Function
            link_.pcm_file_offset = 0
            ' Set pcm_end and end_offset so we can skip the call to
            '   op_find_final_pcm_offset().
            link_.pcm_end = 0
            link_.pcm_start = 0
            link_.end_offset = link_.data_offset
            op_find_initial_pcm_offset = 0
            Exit Function
        End If
        ' Similarly, if we hit the next link in the chain, we've gone too far.
        If ogg_page_bos(og_) Then
            If link_.head.pre_skip > 0 Then op_find_initial_pcm_offset = OP_EBADTIMESTAMP: Exit Function
            ' Set pcm_end and end_offset so we can skip the call to
            '   op_find_final_pcm_offset().
            link_.pcm_file_offset = 0
            link_.pcm_end = 0
            link_.pcm_start = 0
            link_.end_offset = link_.data_offset
            ' Tell the caller we've got a buffered page for them.
            op_find_initial_pcm_offset = 1
            Exit Function
        End If
        ' Ignore pages from other streams (not strictly necessary, because of the
        '   checks in ogg_stream_pagein(), but saves some work).
        If serialno <> ogg_page_serialno(og_) Then GoTo continue_
        ogg_stream_pagein of_.os, og_
        ' Bitrate tracking: add the header's bytes here.
        '  The body bytes are counted when we consume the packets.
        of_.bytes_tracked = of_.bytes_tracked + og_.header_len
        ' Count the durations of all packets in the page.
        Do
            total_duration = op_collect_audio_packets(of_, durations)
        ' Ignore holes.
        Loop While total_duration < 0
        op_count = of_.op_count
continue_:
    Loop While op_count <= 0
    ' We found the first page with a completed audio data packet: actually look
    '   at the granule position.
    '  RFC 3533 says, "A special value of -1 (in two's complement) indicates that
    '   no packets finish on this page," which does not say that a granule
    '   position that is NOT -1 indicates that some packets DO finish on that page
    '   (even though this was the intention, libogg itself violated this intention
    '   for years before we fixed it).
    '  The Ogg Opus specification only imposes its start-time requirements
    '   on the granule position of the first page with completed packets,
    '   so we ignore any set granule positions until then.
    cur_page_gp = of_.op(op_count - 1).granulepos
    ' But getting a packet without a valid granule position on the page is not
    '   okay.
    If cur_page_gp = -1 Then op_find_initial_pcm_offset = OP_EBADTIMESTAMP: Exit Function
    cur_page_eos = of_.op(op_count - 1).e_o_s
    If cur_page_eos = 0 Then
        ' The EOS flag wasn't set.
        '  Work backwards from the provided granule position to get the starting PCM
        '   offset.
        If op_granpos_add(pcm_start, cur_page_gp, -total_duration) < 0 Then
            ' The starting granule position MUST not be smaller than the amount of
            '   audio on the first page with completed packets.
            op_find_initial_pcm_offset = OP_EBADTIMESTAMP
            Exit Function
        End If
    Else
        ' The first page with completed packets was also the last.
        If op_granpos_add(pcm_start, cur_page_gp, -total_duration) < 0 Then
            ' If there's less audio on the page than indicated by the granule
            '   position, then we're doing end-trimming, and the starting PCM offset
            '   is zero by spec mandate.
            pcm_start = 0
            ' However, the end-trimming MUST not ask us to trim more samples than
            '   exist after applying the pre-skip.
            If op_granpos_cmp(cur_page_gp, link_.head.pre_skip) < 0 Then
                op_find_initial_pcm_offset = OP_EBADTIMESTAMP
                Exit Function
            End If
        End If
    End If
    ' Timestamp the individual packets.
    prev_packet_gp = pcm_start
    For pi = 0 To op_count - 1
        If cur_page_eos Then
            op_granpos_diff diff, cur_page_gp, prev_packet_gp       ' (OP_ALWAYS_TRUE)
            diff = durations(pi) - diff
            ' If we have samples to trim...
            If diff > 0 Then
                ' If we trimmed the entire packet, stop (the spec says encoders
                '   shouldn't do this, but we support it anyway).
                If diff > durations(pi) Then Exit For
                prev_packet_gp = cur_page_gp
                of_.op(pi).granulepos = prev_packet_gp
                ' Move the EOS flag to this packet, if necessary, so we'll trim the
                '   samples.
                of_.op(pi).e_o_s = 1
                GoTo next_pi
            End If
        End If
        ' Update the granule position as normal.
        op_granpos_add of_.op(pi).granulepos, prev_packet_gp, durations(pi)  ' (OP_ALWAYS_TRUE)
        prev_packet_gp = of_.op(pi).granulepos
next_pi:
    Next
    ' Update the packet count after end-trimming.
    of_.op_count = pi
    of_.cur_discard_count = link_.head.pre_skip
    link_.pcm_file_offset = 0
    link_.pcm_start = pcm_start
    of_.prev_packet_gp = pcm_start
    of_.prev_page_offset = page_offset
    op_find_initial_pcm_offset = 0
End Function

' Starting from current cursor position, get the final PCM offset of the
'   previous page.
'  This also validates the duration of the link, which, while not strictly
'   required by the spec, we need to ensure duration calculations don't
'   overflow.
'  This is only done for seekable sources.
'  We must validate that op_find_initial_pcm_offset() succeeded for this link
'   before calling this function, otherwise it will scan the entire stream
'   backwards until it reaches the start, and then fail.
Private Function op_find_final_pcm_offset(of_ As OggOpusFile, serialnos_() As Long, ByVal nserialnos_ As Long, _
                                          link_ As OggOpusLink, ByVal offset_ As Double, ByVal end_serialno_ As Long, _
                                          ByVal end_gp_ As Double, total_duration_ As Double) As Long
    Dim total_duration As Double, duration As Double, cur_serialno As Long
    ' For the time being, fetch end PCM offset the simple way.
    cur_serialno = link_.serialno
    If end_serialno_ <> cur_serialno Or end_gp_ = -1 Then
        offset_ = op_get_last_page(of_, end_gp_, offset_, cur_serialno, serialnos_, nserialnos_)
        If offset_ < 0 Then op_find_final_pcm_offset = offset_: Exit Function
    End If
    ' At worst we should have found the first page with completed packets.
    If offset_ < link_.data_offset Then op_find_final_pcm_offset = OP_EBADLINK: Exit Function
    ' This implementation requires that the difference between the first and last
    '   granule positions in each link be representable in a signed, 64-bit
    '   number, and that each link also have at least as many samples as the
    '   pre-skip requires.
    If op_granpos_diff(duration, end_gp_, link_.pcm_start) < 0 Then op_find_final_pcm_offset = OP_EBADTIMESTAMP: Exit Function
    If duration < link_.head.pre_skip Then op_find_final_pcm_offset = OP_EBADTIMESTAMP: Exit Function
    ' We also require that the total duration be representable in a signed,
    '   64-bit number.
    duration = duration - link_.head.pre_skip
    total_duration = total_duration_
    If OP_INT64_MAX - duration < total_duration Then op_find_final_pcm_offset = OP_EBADTIMESTAMP: Exit Function
    total_duration_ = total_duration + duration
    link_.pcm_end = end_gp_
    link_.end_offset = offset_
    op_find_final_pcm_offset = 0
End Function

' Rescale the number _x from the range [0,_from] to [0,_to].
'  _from and _to must be positive.
' (VB) frac's 63 bits are bits(); _x, _from, _to, ret are Doubles (x stays below from)
Private Function op_rescale64(ByVal x_ As Double, ByVal from_ As Double, ByVal to_ As Double) As Double
    Dim bits(0 To 62) As Boolean, ret As Double, i As Long
    If x_ >= from_ Then op_rescale64 = to_: Exit Function
    If x_ <= 0 Then op_rescale64 = 0: Exit Function
    ' frac=0
    For i = 0 To 62
        ' frac<<=1
        ' (OP_ASSERT(_x<=_from))
        If x_ >= Int(from_ / 2) Then
            x_ = x_ - (from_ - x_)
            bits(i) = True                  ' frac|=1
        Else
            x_ = x_ * 2
        End If
    Next
    ret = 0
    For i = 0 To 62
        ' if(frac&1): frac's lowest bit is the one set last
        If bits(62 - i) Then
            ' ret=(ret&_to&1)+(ret>>1)+(_to>>1)
            If (ret - Int(ret / 2) * 2) = 1 And (to_ - Int(to_ / 2) * 2) = 1 Then
                ret = 1 + Int(ret / 2) + Int(to_ / 2)
            Else
                ret = Int(ret / 2) + Int(to_ / 2)
            End If
        Else
            ret = Int(ret / 2)
        End If
        ' frac>>=1
    Next
    op_rescale64 = ret
End Function

' Try to estimate the location of the next link using the current seek
'   records, assuming the initial granule position of any streams we've found is
'   0.
Private Function op_predict_link_start(sr_() As OpusSeekRecord, ByVal nsr_ As Long, ByVal searched_ As Double, _
                                       ByVal end_searched_ As Double, ByVal bias_ As Long) As Double
    Dim bisect As Double, sri As Long, srj As Long, gp1 As Double, gp2_min As Double, serialno1 As Long
    Dim offset1 As Double, gp2 As Double, offset2 As Double, num As Double, den As Double, ipart As Double
    ' Require that we be at least OP_CHUNK_SIZE from the end.
    '  We don't require that we be at least OP_CHUNK_SIZE from the beginning,
    '   because if we are we'll just scan forward without seeking.
    end_searched_ = end_searched_ - OP_CHUNK_SIZE
    If searched_ >= end_searched_ Then op_predict_link_start = -1: Exit Function
    bisect = end_searched_
    For sri = 0 To nsr_ - 1
        ' If the granule position is negative, either it's invalid or we'd cause
        '   overflow.
        '  If it is larger than OP_INT64_MAX-OP_GP_SPACING_MIN, then no positive
        '   granule position would satisfy our minimum spacing requirements below.
        gp1 = sr_(sri).gp
        If gp1 < 0 Or gp1 > OP_INT64_MAX - OP_GP_SPACING_MIN Then GoTo next_sri
        ' We require some minimum distance between granule positions to make an
        '   estimate.
        '  We don't actually know what granule position scheme is being used,
        '   because we have no idea what kind of stream these came from.
        '  Therefore we require a minimum spacing between them, with the
        '   expectation that while bitrates and granule position increments might
        '   vary locally in quite complex ways, they are globally smooth.
        gp2_min = gp1 + OP_GP_SPACING_MIN
        offset1 = sr_(sri).offset
        serialno1 = sr_(sri).serialno
        For srj = sri - 1 To 0 Step -1
            gp2 = sr_(srj).gp
            If gp2 < gp2_min Then GoTo next_srj
            ' Oh, and also make sure these came from the same stream.
            If sr_(srj).serialno <> serialno1 Then GoTo next_srj
            offset2 = sr_(srj).offset
            ' For once, we can subtract with impunity.
            den = gp2 - gp1
            ipart = Fix(gp2 / den)
            num = offset2 - offset1
            ' (OP_ASSERT(num>0))
            If ipart > 0 Then
                If Fix((offset2 - searched_) / ipart) < num Then GoTo next_srj
            End If
            offset2 = offset2 - ipart * num
            gp2 = gp2 - ipart * den
            offset2 = offset2 - (op_rescale64(gp2, den, num) - bias_)
            If offset2 < searched_ Then GoTo next_srj
            If offset2 < bisect Then bisect = offset2
            Exit For
next_srj:
        Next
next_sri:
    Next
    If bisect >= end_searched_ Then op_predict_link_start = -1 Else op_predict_link_start = bisect
End Function

' Finds each bitstream link, one at a time, using a bisection search.
'  This has to begin by knowing the offset of the first link's initial page.
' (VB) _serialnos, _nserialnos, _cserialnos are _of's: serialnos and nserialnos below are the arrays themselves
' (the C's local copies are refreshed at the top of each pass, before any use)
Private Function op_bisect_forward_serialno(of_ As OggOpusFile, ByVal searched_ As Double, sr_() As OpusSeekRecord, _
                                            ByVal csr_ As Long, serialnos_() As Long, nserialnos_ As Long, cserialnos_ As Long) As Long
    Dim og As ogg_page, ogNone As ogg_page, nlinks As Long, clinks As Long, total_duration As Double, nsr As Long, ret As Long
    Dim end_searched As Double, bisect As Double, nxt As Double, last As Double, end_offset As Double, end_gp As Double
    Dim sri As Long, last_offset As Double, avg_link_size As Double, upper_limit As Double, next_bias As Long
    Dim serialno As Long, gp As Double
    nlinks = of_.nlinks
    clinks = of_.nlinks
    total_duration = 0
    ' We start with one seek record, for the last page in the file.
    '  We build up a list of records for places we seek to during link
    '   enumeration.
    '  This list is kept sorted in reverse order.
    '  We only care about seek locations that were _not_ in the current link,
    '   therefore we can add them one at a time to the end of the list as we
    '   improve the lower bound on the location where the next link starts.
    nsr = 1
    Do
        If nlinks >= clinks Then
            If clinks > (2147483647 - 1) \ 2 Then op_bisect_forward_serialno = OP_EFAULT: Exit Function
            clinks = 2 * clinks + 1
            ' (OP_ASSERT(nlinks<clinks))
            ReDim Preserve of_.links(0 To clinks - 1)   ' links=(OggOpusLink *)_ogg_realloc(links,sizeof(*links)*clinks)
        End If
        ' Invariants:
        '  We have the headers and serial numbers for the link beginning at 'begin'.
        '  We have the offset and granule position of the last page in the file
        '   (potentially not a page we care about).
        ' Scan the seek records we already have to save us some bisection.
        For sri = 0 To nsr - 1
            If op_lookup_serialno(sr_(sri).serialno, serialnos_, nserialnos_) Then Exit For
        Next
        ' Is the last page in our current list of serial numbers?
        If sri <= 0 Then Exit Do
        ' Last page wasn't found.
        '  We have at least one more link.
        last = -1
        end_searched = sr_(sri - 1).search_start
        nxt = sr_(sri - 1).offset
        end_gp = -1
        If sri < nsr Then
            searched_ = sr_(sri).offset + sr_(sri).size
            If sr_(sri).serialno = of_.links(nlinks - 1).serialno Then
                end_gp = sr_(sri).gp
                end_offset = sr_(sri).offset
            End If
        End If
        nsr = sri
        bisect = -1
        ' If we've already found the end of at least one link, try to pick the
        '   first bisection point at twice the average link size.
        '  This is a good choice for files with lots of links that are all about the
        '   same size.
        If nlinks > 1 Then
            last_offset = of_.links(nlinks - 1).offset
            avg_link_size = Fix(last_offset / (nlinks - 1))
            upper_limit = end_searched - OP_CHUNK_SIZE - avg_link_size
            If last_offset > searched_ - avg_link_size And last_offset < upper_limit Then
                bisect = last_offset + avg_link_size
                If bisect < upper_limit Then bisect = bisect + avg_link_size
            End If
        End If
        ' We guard against garbage separating the last and first pages of two
        '   links below.
        Do While searched_ < end_searched
            ' If we don't have a better estimate, use simple bisection.
            If bisect = -1 Then bisect = searched_ + Int((end_searched - searched_) / 2)
            ' If we're within OP_CHUNK_SIZE of the start, scan forward.
            If bisect - searched_ < OP_CHUNK_SIZE Then
                bisect = searched_
            ' Otherwise we're skipping data.
            '  Forget the end page, if we saw one, as we might miss a later one.
            Else
                end_gp = -1
            End If
            ret = op_seek_helper(of_, bisect)
            If ret < 0 Then op_bisect_forward_serialno = ret: Exit Function
            last = op_get_next_page(of_, og, sr_(nsr - 1).offset)
            If last < OP_FALSE Then op_bisect_forward_serialno = last: Exit Function
            next_bias = 0
            If last = OP_FALSE Then
                end_searched = bisect
            Else
                serialno = ogg_page_serialno(og)
                gp = ogg_page_granulepos(og)
                If Not op_lookup_serialno(serialno, serialnos_, nserialnos_) Then
                    end_searched = bisect
                    nxt = last
                    ' In reality we should always have enough room, but be paranoid.
                    If nsr < csr_ Then
                        sr_(nsr).search_start = bisect
                        sr_(nsr).offset = last
                        ' (OP_ASSERT(_of->offset-last>=0); OP_ASSERT(_of->offset-last<=OP_PAGE_SIZE_MAX))
                        sr_(nsr).size = of_.offset - last
                        sr_(nsr).serialno = serialno
                        sr_(nsr).gp = gp
                        nsr = nsr + 1
                    End If
                Else
                    searched_ = of_.offset
                    next_bias = OP_CHUNK_SIZE
                    If serialno = of_.links(nlinks - 1).serialno Then
                        ' This page was from the stream we want, remember it.
                        '  If it's the last such page in the link, we won't have to go back
                        '   looking for it later.
                        end_gp = gp
                        end_offset = last
                    End If
                End If
            End If
            bisect = op_predict_link_start(sr_, nsr, searched_, end_searched, next_bias)
        Loop
        ' Bisection point found.
        '  Get the final granule position of the previous link, assuming
        '   op_find_initial_pcm_offset() didn't already determine the link was
        '   empty.
        If of_.links(nlinks - 1).pcm_end = -1 Then
            If end_gp = -1 Then
                ' If we don't know where the end page is, we'll have to seek back and
                '   look for it, starting from the end of the link.
                end_offset = nxt
                ' Also forget the last page we read.
                '  It won't be available after the seek.
                last = -1
            End If
            ret = op_find_final_pcm_offset(of_, serialnos_, nserialnos_, of_.links(nlinks - 1), end_offset, _
                                           of_.links(nlinks - 1).serialno, end_gp, total_duration)
            If ret < 0 Then op_bisect_forward_serialno = ret: Exit Function
        End If
        If last <> nxt Then
            ' The last page we read was not the first page the next link.
            '  Move the cursor position to the offset of that first page.
            '  This only performs an actual seek if the first page of the next link
            '   does not start at the end of the last page from the current Opus
            '   stream with a valid granule position.
            ret = op_seek_helper(of_, nxt)
            If ret < 0 Then op_bisect_forward_serialno = ret: Exit Function
        End If
        ret = op_fetch_headers(of_, of_.links(nlinks).head, of_.links(nlinks).tags, serialnos_, nserialnos_, cserialnos_, _
                               False, og, last <> nxt)
        If ret < 0 Then op_bisect_forward_serialno = ret: Exit Function
        ' Mark the current link count so it can be cleaned up on error.
        of_.nlinks = nlinks + 1
        of_.links(nlinks).offset = nxt
        of_.links(nlinks).data_offset = of_.offset
        of_.links(nlinks).serialno = of_.os.serialno
        of_.links(nlinks).pcm_end = -1
        ' This might consume a page from the next link, however the next bisection
        '   always starts with a seek.
        ret = op_find_initial_pcm_offset(of_, of_.links(nlinks), ogNone)     ' (_og NULL)
        If ret < 0 Then op_bisect_forward_serialno = ret: Exit Function
        of_.links(nlinks).pcm_file_offset = total_duration
        searched_ = of_.offset
        nlinks = nlinks + 1
    Loop
    ' Last page is in the starting serialno list, so we've reached the last link.
    '  Now find the last granule position for it (if we didn't the first time we
    '   looked at the end of the stream, and if op_find_initial_pcm_offset()
    '   didn't already determine the link was empty).
    If of_.links(nlinks - 1).pcm_end = -1 Then
        ret = op_find_final_pcm_offset(of_, serialnos_, nserialnos_, of_.links(nlinks - 1), sr_(0).offset, _
                                       sr_(0).serialno, sr_(0).gp, total_duration)
        If ret < 0 Then op_bisect_forward_serialno = ret: Exit Function
    End If
    ' Trim back the links array if necessary.
    ReDim Preserve of_.links(0 To nlinks - 1)
    ' We also don't need these anymore.
    Erase serialnos_                            ' _ogg_free(*_serialnos); *_serialnos=NULL
    cserialnos_ = 0
    nserialnos_ = 0
    op_bisect_forward_serialno = 0
End Function

Private Sub op_update_gain(of_ As OggOpusFile)
    Dim gain_q8 As Long, li As Long, album_gain_q8 As Long, track_gain_q8 As Long
    ' If decode isn't ready, then we'll apply the gain when we initialize the
    '   decoder.
    If of_.ready_state < OP_INITSET Then Exit Sub
    gain_q8 = of_.gain_offset_q8
    If of_.seekable Then li = of_.cur_link Else li = 0
    ' (head=&_of->links[li].head)
    ' We don't have to worry about overflow here because the header gain and
    '   track gain must lie in the range [-32768,32767], and the user-supplied
    '   offset has been pre-clamped to [-98302,98303].
    Select Case of_.gain_type
    Case OP_ALBUM_GAIN
        album_gain_q8 = 0
        opus_tags_get_album_gain of_.links(li).tags, album_gain_q8
        gain_q8 = gain_q8 + album_gain_q8
        gain_q8 = gain_q8 + of_.links(li).head.output_gain
    Case OP_TRACK_GAIN
        track_gain_q8 = 0
        opus_tags_get_track_gain of_.links(li).tags, track_gain_q8
        gain_q8 = gain_q8 + track_gain_q8
        gain_q8 = gain_q8 + of_.links(li).head.output_gain
    Case OP_HEADER_GAIN
        gain_q8 = gain_q8 + of_.links(li).head.output_gain
    Case OP_ABSOLUTE_GAIN
    Case Else
        ' (OP_ASSERT(0))
    End Select
    ' OP_CLAMP(-32768,gain_q8,32767)
    If gain_q8 > 32767 Then gain_q8 = 32767
    If gain_q8 < -32768 Then gain_q8 = -32768
    ' (OP_ASSERT(_of->od!=NULL))
    ' (OPUS_SET_GAIN is defined)
    opus_multistream_decoder_ctl of_.od, OPUS_SET_GAIN_REQUEST, gain_q8
End Sub

Private Function op_make_decode_ready(of_ As OggOpusFile) As Long
    Dim li As Long, stream_count As Long, coupled_count As Long, channel_count As Long, same As Boolean
    Dim i As Long, err_ As Long
    If of_.ready_state > OP_STREAMSET Then op_make_decode_ready = 0: Exit Function
    If of_.ready_state < OP_STREAMSET Then op_make_decode_ready = OP_EFAULT: Exit Function
    If of_.seekable Then li = of_.cur_link Else li = 0
    ' (head=&_of->links[li].head)
    stream_count = of_.links(li).head.stream_count
    coupled_count = of_.links(li).head.coupled_count
    channel_count = of_.links(li).head.channel_count
    ' Check to see if the current decoder is compatible with the current link.
    same = False
    If of_.od <> 0 And of_.od_stream_count = stream_count And of_.od_coupled_count = coupled_count _
       And of_.od_channel_count = channel_count Then
        same = True                         ' memcmp(_of->od_mapping,head->mapping,channel_count)==0
        For i = 0 To channel_count - 1
            If of_.od_mapping(i) <> of_.links(li).head.mapping(i) Then same = False: Exit For
        Next
    End If
    If same Then
        opus_multistream_decoder_ctl of_.od, OPUS_RESET_STATE
    Else
        If of_.od <> 0 Then opus_multistream_decoder_destroy of_.od      ' ((VB) destroy(NULL) does nothing)
        of_.od = opus_multistream_decoder_create(48000, channel_count, stream_count, coupled_count, _
                                                 of_.links(li).head.mapping, 0, err_)
        If of_.od = 0 Then op_make_decode_ready = OP_EFAULT: Exit Function
        of_.od_stream_count = stream_count
        of_.od_coupled_count = coupled_count
        of_.od_channel_count = channel_count
        For i = 0 To channel_count - 1
            of_.od_mapping(i) = of_.links(li).head.mapping(i)
        Next
    End If
    of_.ready_state = OP_INITSET
    of_.bytes_tracked = 0
    of_.samples_tracked = 0
    ' (!OP_FIXED_POINT)
    of_.state_channel_count = 0
    ' Use the serial number for the PRNG seed to get repeatable output for
    '   straight play-throughs.
    of_.dither_seed = of_.links(li).serialno
    op_update_gain of_
    op_make_decode_ready = 0
End Function

Private Function op_open_seekable2_impl(of_ As OggOpusFile) As Long
    Dim sr(0 To 63) As OpusSeekRecord, data_offset As Double, ret As Long
    ' 64 seek records should be enough for anybody.
    '  Actually, with a bisection search in a 63-bit range down to OP_CHUNK_SIZE
    '   granularity, much more than enough.
    ' We can seek, so set out learning all about this file.
    op_vbCallSeek of_.callbacks.seek_, of_.stream, 0, SEEK_END
    of_.end_ = op_vbCallTell(of_.callbacks.tell, of_.stream)
    of_.offset = of_.end_
    If of_.end_ < 0 Then op_open_seekable2_impl = OP_EREAD: Exit Function
    data_offset = of_.links(0).data_offset
    If of_.end_ < data_offset Then op_open_seekable2_impl = OP_EBADLINK: Exit Function
    ' Get the offset of the last page of the physical bitstream, or, if we're
    '   lucky, the last Opus page of the first link, as most Ogg Opus files will
    '   contain a single logical bitstream.
    ret = op_get_prev_page_serial(of_, sr(0), of_.end_, of_.links(0).serialno, of_.serialnos, of_.nserialnos)
    If ret < 0 Then op_open_seekable2_impl = ret: Exit Function
    ' If there's any trailing junk, forget about it.
    of_.end_ = sr(0).offset + sr(0).size
    If of_.end_ < data_offset Then op_open_seekable2_impl = OP_EBADLINK: Exit Function
    ' Now enumerate the bitstream structure.
    op_open_seekable2_impl = op_bisect_forward_serialno(of_, data_offset, sr, 64, of_.serialnos, of_.nserialnos, of_.cserialnos)
End Function

Private Function op_open_seekable2(of_ As OggOpusFile) As Long
    Dim oy_start As ogg_sync_state, os_start As ogg_stream_state, op_start() As ogg_packet
    Dim prev_page_offset As Double, start_offset As Double, start_op_count As Long, ret As Long, i As Long
    ' We're partially open and have a first link header state in storage in _of.
    '  Save off that stream state so we can come back to it.
    '  It would be simpler to just dump all this state and seek back to
    '   links[0].data_offset when we're done.
    '  But we do the extra work to allow us to seek back to _exactly_ the same
    '   stream position we're at now.
    '  This allows, e.g., the HTTP backend to continue reading from the original
    '   connection (if it's still available), instead of opening a new one.
    '  This means we can open and start playing a normal Opus file with a single
    '   link and reasonable packet sizes using only two HTTP requests.
    start_op_count = of_.op_count
    ' This is a bit too large to put on the stack unconditionally.
    If start_op_count > 0 Then ReDim op_start(0 To start_op_count - 1)
    oy_start = of_.oy
    os_start = of_.os
    prev_page_offset = of_.prev_page_offset
    start_offset = of_.offset
    For i = 0 To start_op_count - 1             ' memcpy(op_start,_of->op,sizeof(*op_start)*start_op_count)
        op_start(i) = of_.op(i)
    Next
    ' (OP_ASSERT((*_of->callbacks.tell)(_of->stream)==op_position(_of)))
    ogg_sync_init of_.oy
    ogg_stream_init of_.os, -1
    ret = op_open_seekable2_impl(of_)
    ' Restore the old stream state.
    ogg_stream_clear of_.os
    ogg_sync_clear of_.oy
    of_.oy = oy_start
    of_.os = os_start
    of_.offset = start_offset
    of_.op_count = start_op_count
    For i = 0 To start_op_count - 1
        of_.op(i) = op_start(i)
    Next
    Erase op_start                              ' _ogg_free(op_start)
    of_.prev_packet_gp = of_.links(0).pcm_start
    of_.prev_page_offset = prev_page_offset
    of_.cur_discard_count = of_.links(0).head.pre_skip
    If ret < 0 Then op_open_seekable2 = ret: Exit Function
    ' And restore the position indicator.
    ret = op_vbCallSeek(of_.callbacks.seek_, of_.stream, op_position(of_), SEEK_SET)
    If ret < 0 Then op_open_seekable2 = OP_EREAD Else op_open_seekable2 = 0
End Function

' Clear out the current logical bitstream decoder.
Private Sub op_decode_clear(of_ As OggOpusFile)
    ' We don't actually free the decoder.
    '  We might be able to re-use it for the next link.
    of_.op_count = 0
    of_.od_buffer_size = 0
    of_.prev_packet_gp = -1
    of_.prev_page_offset = -1
    If of_.seekable = 0 Then
        ' (OP_ASSERT(_of->ready_state>=OP_INITSET))
        opus_tags_clear of_.links(0).tags
    End If
    of_.ready_state = OP_OPENED
End Sub

Private Sub op_clear(of_ As OggOpusFile)
    Dim nlinks As Long, link As Long
    Erase of_.od_buffer                         ' _ogg_free(_of->od_buffer)
    of_.od_buffer_allocated = False
    If of_.od <> 0 Then opus_multistream_decoder_destroy of_.od
    ' (links=_of->links)
    If of_.seekable = 0 Then
        If of_.ready_state > OP_OPENED Or of_.ready_state = OP_PARTOPEN Then
            opus_tags_clear of_.links(0).tags
        End If
    ElseIf of_.links_allocated Then
        nlinks = of_.nlinks
        For link = 0 To nlinks - 1
            opus_tags_clear of_.links(link).tags
        Next
    End If
    Erase of_.links                             ' _ogg_free(links)
    of_.links_allocated = False
    Erase of_.serialnos                         ' _ogg_free(_of->serialnos)
    ogg_stream_clear of_.os
    ogg_sync_clear of_.oy
    If of_.callbacks.close_ <> 0 Then op_vbCallClose of_.callbacks.close_, of_.stream
End Sub

' (VB) _initial_data is initial_data(0 ...)
Private Function op_open1(of_ As OggOpusFile, stream_ As Variant, cb_ As OpusFileCallbacks, initial_data() As Byte, _
                          ByVal initial_bytes As Long) As Long
    Dim og As ogg_page, pogIsNull As Boolean, seekable As Long, ret As Long, blank As OggOpusFile, buffer As Long
    Dim pos As Double
    of_ = blank                                 ' memset(_of,0,sizeof(*_of))
    ' (if(OP_UNLIKELY(_initial_bytes>(size_t)LONG_MAX))return OP_EFAULT: a Long)
    of_.end_ = -1
    op_vbSetVoid of_.stream, stream_
    of_.callbacks = cb_
    ' At a minimum, we need to be able to read data.
    If of_.callbacks.read_ = 0 Then op_open1 = OP_EREAD: Exit Function
    ' Initialize the framing state.
    ogg_sync_init of_.oy
    ' Perhaps some data was previously read into a buffer for testing against
    '   other stream types.
    '  Allow initialization from this previously read data (especially as we may
    '   be reading from a non-seekable stream).
    '  This requires copying it into a buffer allocated by ogg_sync_buffer() and
    '   doesn't support seeking, so this is not a good mechanism to use for
    '   decoding entire files from RAM.
    If initial_bytes > 0 Then
        buffer = ogg_sync_buffer(of_.oy, initial_bytes)
        If buffer < 0 Then op_open1 = OP_EFAULT: Exit Function
        CopyMemory of_.oy.data(buffer), initial_data(0), initial_bytes
        ogg_sync_wrote of_.oy, initial_bytes
    End If
    ' Can we seek?
    '  Stevens suggests the seek test is portable.
    '  It's actually not for files on win32, but we address that by fixing it in
    '   our callback implementation (see stream.c).
    seekable = 0
    If cb_.seek_ <> 0 Then
        If op_vbCallSeek(cb_.seek_, stream_, 0, SEEK_CUR) <> -1 Then seekable = 1
    End If
    ' If seek is implemented, tell must also be implemented.
    If seekable Then
        If of_.callbacks.tell = 0 Then op_open1 = OP_EINVAL: Exit Function
        pos = op_vbCallTell(of_.callbacks.tell, of_.stream)
        ' If the current position is not equal to the initial bytes consumed,
        '   absolute seeking will not work.
        If pos <> initial_bytes Then op_open1 = OP_EINVAL: Exit Function
    End If
    of_.seekable = seekable
    ' Don't seek yet.
    '  Set up a 'single' (current) logical bitstream entry for partial open.
    ReDim of_.links(0 To 0)                     ' _of->links=(OggOpusLink *)_ogg_malloc(sizeof(*_of->links))
    of_.links_allocated = True
    ' The serialno gets filled in later by op_fetch_headers().
    ogg_stream_init of_.os, -1
    pogIsNull = True
    Do
        ' Fetch all BOS pages, store the Opus header and all seen serial numbers,
        '  and load subsequent Opus setup headers.
        ret = op_fetch_headers(of_, of_.links(0).head, of_.links(0).tags, of_.serialnos, of_.nserialnos, of_.cserialnos, _
                               False, og, pogIsNull)
        If ret < 0 Then Exit Do
        of_.nlinks = 1
        of_.links(0).offset = 0
        of_.links(0).data_offset = of_.offset
        of_.links(0).pcm_end = -1
        of_.links(0).serialno = of_.os.serialno
        ' Fetch the initial PCM offset.
        ret = op_find_initial_pcm_offset(of_, of_.links(0), og)
        If seekable <> 0 Or ret <= 0 Then Exit Do
        ' This link was empty, but we already have the BOS page for the next one in
        '   og.
        '  We can't seek, so start processing the next link right now.
        opus_tags_clear of_.links(0).tags
        of_.nlinks = 0
        If seekable = 0 Then of_.cur_link = of_.cur_link + 1
        pogIsNull = False                       ' pog=&og
    Loop
    If ret >= 0 Then of_.ready_state = OP_PARTOPEN
    op_open1 = ret
End Function

Private Function op_open2(of_ As OggOpusFile) As Long
    Dim ret As Long
    ' (OP_ASSERT(_of->ready_state==OP_PARTOPEN))
    If of_.seekable Then
        of_.ready_state = OP_OPENED
        ret = op_open_seekable2(of_)
    Else
        ret = 0
    End If
    If ret >= 0 Then
        ' We have buffered packets from op_find_initial_pcm_offset().
        '  Move to OP_INITSET so we can use them.
        of_.ready_state = OP_STREAMSET
        ret = op_make_decode_ready(of_)
        If ret >= 0 Then op_open2 = 0: Exit Function
    End If
    ' Don't auto-close the stream on failure.
    of_.callbacks.close_ = 0
    op_clear of_
    op_open2 = ret
End Function

' (VB) _initial_data NULL is an empty array with _initial_bytes 0; returns the OggOpusFile handle (0 = NULL)
Public Function op_test_callbacks(stream_ As Variant, cb_ As OpusFileCallbacks, initial_data() As Byte, _
                                  ByVal initial_bytes As Long, error_ As Long) As Long
    Dim of_ As Long, ret As Long
    of_ = op_vbAlloc()
    ret = OP_EFAULT
    ' (if(OP_LIKELY(of!=NULL)): VB raises an error where malloc returns NULL)
    ret = op_open1(ofPool(of_), stream_, cb_, initial_data, initial_bytes)
    If ret >= 0 Then
        error_ = 0
        op_test_callbacks = of_
        Exit Function
    End If
    ' Don't auto-close the stream on failure.
    ofPool(of_).callbacks.close_ = 0
    op_clear ofPool(of_)
    op_vbFree of_
    error_ = ret
    op_test_callbacks = 0
End Function

Public Function op_open_callbacks(stream_ As Variant, cb_ As OpusFileCallbacks, initial_data() As Byte, _
                                  ByVal initial_bytes As Long, error_ As Long) As Long
    Dim of_ As Long, ret As Long
    of_ = op_test_callbacks(stream_, cb_, initial_data, initial_bytes, error_)
    If of_ <> 0 Then
        ret = op_open2(ofPool(of_))
        If ret >= 0 Then op_open_callbacks = of_: Exit Function
        error_ = ret
        op_vbFree of_
    End If
    op_open_callbacks = 0
End Function

' Convenience routine to clean up from failure for the open functions that
'   create their own streams.
Private Function op_open_close_on_failure(stream_ As Variant, cb_ As OpusFileCallbacks, error_ As Long) As Long
    Dim of_ As Long, noData() As Byte
    If op_vbIsNull(stream_) Then
        error_ = OP_EFAULT
        op_open_close_on_failure = 0
        Exit Function
    End If
    of_ = op_open_callbacks(stream_, cb_, noData, 0, error_)
    If of_ = 0 Then op_vbCallClose cb_.close_, stream_
    op_open_close_on_failure = of_
End Function

Public Function op_open_file(ByVal path_ As String, error_ As Long) As Long
    Dim cb As OpusFileCallbacks, stream_ As Variant
    stream_ = op_fopen(cb, path_, "rb")
    op_open_file = op_open_close_on_failure(stream_, cb, error_)
End Function

' (VB) the stream reads a copy of data_(0 .. size_ - 1)
Public Function op_open_memory(data_() As Byte, ByVal size_ As Long, error_ As Long) As Long
    Dim cb As OpusFileCallbacks, stream_ As Variant
    stream_ = op_mem_stream_create(cb, data_, size_)
    op_open_memory = op_open_close_on_failure(stream_, cb, error_)
End Function

' Convenience routine to clean up from failure for the open functions that
'   create their own streams.
Private Function op_test_close_on_failure(stream_ As Variant, cb_ As OpusFileCallbacks, error_ As Long) As Long
    Dim of_ As Long, noData() As Byte
    If op_vbIsNull(stream_) Then
        error_ = OP_EFAULT
        op_test_close_on_failure = 0
        Exit Function
    End If
    of_ = op_test_callbacks(stream_, cb_, noData, 0, error_)
    If of_ = 0 Then op_vbCallClose cb_.close_, stream_
    op_test_close_on_failure = of_
End Function

Public Function op_test_file(ByVal path_ As String, error_ As Long) As Long
    Dim cb As OpusFileCallbacks, stream_ As Variant
    stream_ = op_fopen(cb, path_, "rb")
    op_test_file = op_test_close_on_failure(stream_, cb, error_)
End Function

Public Function op_test_memory(data_() As Byte, ByVal size_ As Long, error_ As Long) As Long
    Dim cb As OpusFileCallbacks, stream_ As Variant
    stream_ = op_mem_stream_create(cb, data_, size_)
    op_test_memory = op_test_close_on_failure(stream_, cb, error_)
End Function

Public Function op_test_open(ByVal of_ As Long) As Long
    Dim ret As Long, blank As OggOpusFile
    If ofPool(of_).ready_state <> OP_PARTOPEN Then op_test_open = OP_EINVAL: Exit Function
    ret = op_open2(ofPool(of_))
    ' op_open2() will clear this structure on failure.
    '  Reset its contents to prevent double-frees in op_free().
    If ret < 0 Then ofPool(of_) = blank
    op_test_open = ret
End Function

Public Sub op_free(ByVal of_ As Long)
    If of_ <> 0 Then
        op_clear ofPool(of_)
        op_vbFree of_
    End If
End Sub

Public Function op_seekable(ByVal of_ As Long) As Long
    op_seekable = ofPool(of_).seekable
End Function

Public Function op_link_count(ByVal of_ As Long) As Long
    op_link_count = ofPool(of_).nlinks
End Function

Public Function op_serialno(ByVal of_ As Long, ByVal li_ As Long) As Long
    If li_ >= ofPool(of_).nlinks Then li_ = ofPool(of_).nlinks - 1
    If ofPool(of_).seekable = 0 Then li_ = 0
    If li_ < 0 Then li_ = ofPool(of_).cur_link
    op_serialno = ofPool(of_).links(li_).serialno
End Function

Public Function op_channel_count(ByVal of_ As Long, ByVal li_ As Long) As Long
    op_channel_count = op_head(of_, li_).channel_count
End Function

Public Function op_raw_total(ByVal of_ As Long, ByVal li_ As Long) As Double
    Dim a As Double, b As Double
    If ofPool(of_).ready_state < OP_OPENED Or ofPool(of_).seekable = 0 Or li_ >= ofPool(of_).nlinks Then
        op_raw_total = OP_EINVAL
        Exit Function
    End If
    If li_ < 0 Then op_raw_total = ofPool(of_).end_: Exit Function
    If li_ + 1 >= ofPool(of_).nlinks Then a = ofPool(of_).end_ Else a = ofPool(of_).links(li_ + 1).offset
    If li_ > 0 Then b = ofPool(of_).links(li_).offset Else b = 0
    op_raw_total = a - b
End Function

Public Function op_pcm_total(ByVal of_ As Long, ByVal li_ As Long) As Double
    Dim pcm_total As Double, diff As Double, nlinks As Long
    nlinks = ofPool(of_).nlinks
    If ofPool(of_).ready_state < OP_OPENED Or ofPool(of_).seekable = 0 Or li_ >= nlinks Then
        op_pcm_total = OP_EINVAL
        Exit Function
    End If
    ' (links=_of->links)
    ' We verify that the granule position differences are larger than the
    '   pre-skip and that the total duration does not overflow during link
    '   enumeration, so we don't have to check here.
    pcm_total = 0
    If li_ < 0 Then
        pcm_total = ofPool(of_).links(nlinks - 1).pcm_file_offset
        li_ = nlinks - 1
    End If
    op_granpos_diff diff, ofPool(of_).links(li_).pcm_end, ofPool(of_).links(li_).pcm_start  ' (OP_ALWAYS_TRUE)
    op_pcm_total = pcm_total + (diff - ofPool(of_).links(li_).head.pre_skip)
End Function

' (VB) returns a copy of the header (the C's pointer is never NULL)
Public Function op_head(ByVal of_ As Long, ByVal li_ As Long) As OpusHead
    If li_ >= ofPool(of_).nlinks Then li_ = ofPool(of_).nlinks - 1
    If ofPool(of_).seekable = 0 Then li_ = 0
    If li_ < 0 Then li_ = ofPool(of_).cur_link
    op_head = ofPool(of_).links(li_).head
End Function

' (VB) returns false for NULL, else a copy of the tags in tags_
Public Function op_tags(ByVal of_ As Long, ByVal li_ As Long, tags_ As OpusTags) As Boolean
    If li_ >= ofPool(of_).nlinks Then li_ = ofPool(of_).nlinks - 1
    If ofPool(of_).seekable = 0 Then
        If ofPool(of_).ready_state < OP_STREAMSET And ofPool(of_).ready_state <> OP_PARTOPEN Then
            op_tags = False
            Exit Function
        End If
        li_ = 0
    ElseIf li_ < 0 Then
        If ofPool(of_).ready_state >= OP_STREAMSET Then li_ = ofPool(of_).cur_link Else li_ = 0
    End If
    tags_ = ofPool(of_).links(li_).tags
    op_tags = True
End Function

Public Function op_current_link(ByVal of_ As Long) As Long
    If ofPool(of_).ready_state < OP_OPENED Then op_current_link = OP_EINVAL: Exit Function
    op_current_link = ofPool(of_).cur_link
End Function

' Compute an average bitrate given a byte and sample count.
'  Return: The bitrate in bits per second.
Private Function op_calc_bitrate(ByVal bytes_ As Double, ByVal samples_ As Double) As Long
    Dim den As Double, r As Double
    If samples_ <= 0 Then op_calc_bitrate = OP_INT32_MAX: Exit Function
    ' These rates are absurd, but let's handle them anyway.
    If bytes_ > Int((OP_INT64_MAX - Int(samples_ / 2)) / (48000# * 8)) Then
        If Int(bytes_ / (OP_INT32_MAX \ (48000& * 8))) >= samples_ Then
            op_calc_bitrate = OP_INT32_MAX
            Exit Function
        End If
        den = Int(samples_ / (48000# * 8))
        op_calc_bitrate = Int((bytes_ + Int(den / 2)) / den)
        Exit Function
    End If
    ' This can't actually overflow in normal operation: even with a pre-skip of
    '   545 2.5 ms frames with 8 streams running at 1282*8+1 bytes per packet
    '   (1275 byte frames + Opus framing overhead + Ogg lacing values), that all
    '   produce a single sample of decoded output, we still don't top 45 Mbps.
    '  The only way to get bitrates larger than that is with excessive Opus
    '   padding, more encoded streams than output channels, or lots and lots of
    '   Ogg pages with no packets on them.
    r = Int((bytes_ * 48000# * 8 + Int(samples_ / 2)) / samples_)
    If r < OP_INT32_MAX Then op_calc_bitrate = r Else op_calc_bitrate = OP_INT32_MAX
End Function

Public Function op_bitrate(ByVal of_ As Long, ByVal li_ As Long) As Long
    If ofPool(of_).ready_state < OP_OPENED Or ofPool(of_).seekable = 0 Or li_ >= ofPool(of_).nlinks Then
        op_bitrate = OP_EINVAL
        Exit Function
    End If
    op_bitrate = op_calc_bitrate(op_raw_total(of_, li_), op_pcm_total(of_, li_))
End Function

Public Function op_bitrate_instant(ByVal of_ As Long) As Long
    Dim samples_tracked As Double, ret As Long
    If ofPool(of_).ready_state < OP_OPENED Then op_bitrate_instant = OP_EINVAL: Exit Function
    samples_tracked = ofPool(of_).samples_tracked
    If samples_tracked = 0 Then op_bitrate_instant = OP_FALSE: Exit Function
    ret = op_calc_bitrate(ofPool(of_).bytes_tracked, samples_tracked)
    ofPool(of_).bytes_tracked = 0
    ofPool(of_).samples_tracked = 0
    op_bitrate_instant = ret
End Function

' Given a serialno, find a link with a corresponding Opus stream, if it exists.
'  Return: The index of the link to which the page belongs, or a negative number
'           if it was not a desired Opus bitstream section.
Private Function op_get_link_from_serialno(of_ As OggOpusFile, ByVal cur_link_ As Long, ByVal page_offset_ As Double, _
                                           ByVal serialno_ As Long) As Long
    Dim nlinks As Long, li_lo As Long, li_hi As Long
    ' (OP_ASSERT(_of->seekable))
    nlinks = of_.nlinks
    li_lo = 0
    ' Start off by guessing we're just a multiplexed page in the current link.
    li_hi = nlinks
    If cur_link_ + 1 < nlinks Then
        If page_offset_ < of_.links(cur_link_ + 1).offset Then li_hi = cur_link_ + 1
    End If
    Do
        If page_offset_ >= of_.links(cur_link_).offset Then li_lo = cur_link_ Else li_hi = cur_link_
        cur_link_ = li_lo + (li_hi - li_lo) \ 2
    Loop While li_hi - li_lo > 1
    ' We've identified the link that should contain this page.
    '  Make sure it's a page we care about.
    If of_.links(cur_link_).serialno <> serialno_ Then op_get_link_from_serialno = OP_FALSE: Exit Function
    op_get_link_from_serialno = cur_link_
End Function

' Fetch and process a page.
'  This handles the case where we're at a bitstream boundary and dumps the
'   decoding machine.
'  If the decoding machine is unloaded, it loads it.
'  It also keeps prev_packet_gp up to date (seek and read both use this).
'  Return: <0) Error, OP_HOLE (lost packet), or OP_EOF.
'           0) Got at least one audio data packet.
' (VB) _og NULL is ogIsNull
Private Function op_fetch_and_process_page(of_ As OggOpusFile, og_ As ogg_page, ByVal ogIsNull As Boolean, _
                                           ByVal page_offset_ As Double, ByVal spanp_ As Long, ByVal ignore_holes_ As Long) As Long
    Dim cur_serialno As Long, seekable As Long, cur_link As Long, ret As Long, og As ogg_page, serialno As Long
    Dim new_link As Long, total_duration As Long, durations(0 To 254) As Long, op_count As Long, report_hole As Long
    Dim diff As Double, prev_packet_gp As Double, cur_packet_gp As Double, cur_page_gp As Double, cur_page_eos As Long
    Dim pi As Long, cur_discard_count As Long, pre_skip As Long, doTrim As Boolean, noSerialnos() As Long, noCount As Long
    Dim mx As Long
    ' We shouldn't get here if we have unprocessed packets.
    ' (OP_ASSERT(_of->ready_state<OP_INITSET||_of->op_pos>=_of->op_count))
    seekable = of_.seekable
    ' (links=_of->links)
    If seekable Then cur_link = of_.cur_link Else cur_link = 0
    cur_serialno = of_.links(cur_link).serialno
    ' Handle one page.
    Do
        ' (OP_ASSERT(_of->ready_state>=OP_OPENED))
        ' If we were given a page to use, use it.
        If Not ogIsNull Then
            og = og_
            ogIsNull = True
        ' Keep reading until we get a page with the correct serialno.
        Else
            page_offset_ = op_get_next_page(of_, og, of_.end_)
        End If
        ' EOF: Leave uninitialized.
        If page_offset_ < 0 Then
            If page_offset_ < OP_FALSE Then op_fetch_and_process_page = page_offset_ Else op_fetch_and_process_page = OP_EOF
            Exit Function
        End If
        If of_.ready_state >= OP_STREAMSET And cur_serialno <> ogg_page_serialno(og) Then
            ' Two possibilities:
            '   1) Another stream is multiplexed into this logical section, or
            If ogg_page_bos(og) = 0 Then GoTo continue_
            '  2) Our decoding just traversed a bitstream boundary.
            If spanp_ = 0 Then op_fetch_and_process_page = OP_EOF: Exit Function
            If of_.ready_state >= OP_INITSET Then op_decode_clear of_
        ' Bitrate tracking: add the header's bytes here.
        '  The body bytes are counted when we consume the packets.
        Else
            of_.bytes_tracked = of_.bytes_tracked + og.header_len
        End If
        ' Do we need to load a new machine before submitting the page?
        '  This is different in the seekable and non-seekable cases.
        '  In the seekable case, we already have all the header information loaded
        '   and cached.
        '  We just initialize the machine with it and continue on our merry way.
        '  In the non-seekable (streaming) case, we'll only be at a boundary if we
        '   just left the previous logical bitstream, and we're now nominally at the
        '   header of the next bitstream.
        If of_.ready_state < OP_STREAMSET Then
            If seekable Then
                serialno = ogg_page_serialno(og)
                ' Match the serialno to bitstream section.
                ' (OP_ASSERT(cur_link>=0&&cur_link<_of->nlinks))
                If of_.links(cur_link).serialno <> serialno Then
                    ' It wasn't a page from the current link.
                    '  Is it from the next one?
                    new_link = -1
                    If cur_link + 1 < of_.nlinks Then
                        If of_.links(cur_link + 1).serialno = serialno Then new_link = cur_link + 1
                    End If
                    If new_link < 0 Then
                        new_link = op_get_link_from_serialno(of_, cur_link, page_offset_, serialno)
                        ' Not a desired Opus bitstream section.
                        '  Keep trying.
                        If new_link < 0 Then GoTo continue_
                    End If
                    cur_link = new_link
                End If
                cur_serialno = serialno
                of_.cur_link = cur_link
                ogg_stream_reset_serialno of_.os, serialno
                of_.ready_state = OP_STREAMSET
                ' If we're at the start of this link, initialize the granule position
                '   and pre-skip tracking.
                If page_offset_ <= of_.links(cur_link).data_offset Then
                    of_.prev_packet_gp = of_.links(cur_link).pcm_start
                    of_.prev_page_offset = -1
                    of_.cur_discard_count = of_.links(cur_link).head.pre_skip
                    ' Ignore a hole at the start of a new link (this is common for
                    '   streams joined in the middle) or after seeking.
                    ignore_holes_ = 1
                End If
            Else
                Do
                    ' We're streaming.
                    '  Fetch the two header packets, build the info struct.
                    ret = op_fetch_headers(of_, of_.links(0).head, of_.links(0).tags, noSerialnos, noCount, noCount, True, og, False)
                    If ret < 0 Then op_fetch_and_process_page = ret: Exit Function
                    ' op_find_initial_pcm_offset() will suppress any initial hole for us,
                    '   so no need to set _ignore_holes.
                    ret = op_find_initial_pcm_offset(of_, of_.links(0), og)
                    If ret < 0 Then op_fetch_and_process_page = ret: Exit Function
                    cur_serialno = of_.os.serialno
                    of_.links(0).serialno = cur_serialno
                    of_.cur_link = of_.cur_link + 1
                ' If the link was empty, keep going, because we already have the
                '   BOS page of the next one in og.
                Loop While ret > 0
                ' If we didn't get any packets out of op_find_initial_pcm_offset(),
                '   keep going (this is possible if end-trimming trimmed them all).
                If of_.op_count <= 0 Then GoTo continue_
                ' Otherwise, we're done.
                '  TODO: This resets bytes_tracked, which misses the header bytes
                '   already processed by op_find_initial_pcm_offset().
                ret = op_make_decode_ready(of_)
                If ret < 0 Then op_fetch_and_process_page = ret: Exit Function
                op_fetch_and_process_page = 0
                Exit Function
            End If
        End If
        ' The buffered page is the data we want, and we're ready for it.
        '  Add it to the stream state.
        If of_.ready_state = OP_STREAMSET Then
            ret = op_make_decode_ready(of_)
            If ret < 0 Then op_fetch_and_process_page = ret: Exit Function
        End If
        ' Extract all the packets from the current page.
        ogg_stream_pagein of_.os, og
        If of_.ready_state >= OP_INITSET Then
            report_hole = 0
            total_duration = op_collect_audio_packets(of_, durations)
            If total_duration < 0 Then
                ' libogg reported a hole (a gap in the page sequence numbers).
                '  Drain the packets from the page anyway.
                '  If we don't, they'll still be there when we fetch the next page.
                '  Then, when we go to pull out packets, we might get more than 255,
                '   which would overrun our packet buffer.
                '  We repeat this call until we get any actual packets, since we might
                '   have buffered multiple out-of-sequence pages with no packets on
                '   them.
                Do
                    total_duration = op_collect_audio_packets(of_, durations)
                Loop While total_duration < 0
                If ignore_holes_ = 0 Then
                    ' Report the hole to the caller after we finish timestamping the
                    '   packets.
                    report_hole = 1
                    ' We had lost or damaged pages, so reset our granule position
                    '   tracking.
                    '  This makes holes behave the same as a small raw seek.
                    '  If the next page is the EOS page, we'll discard it (because we
                    '   can't perform end trimming properly), and we'll always discard at
                    '   least 80 ms of audio (to allow decoder state to re-converge).
                    '  We could try to fill in the gap with PLC by looking at timestamps
                    '   in the non-EOS case, but that's complicated and error prone and we
                    '   can't rely on the timestamps being valid.
                    of_.prev_packet_gp = -1
                End If
            End If
            op_count = of_.op_count
            ' If we found at least one audio data packet, compute per-packet granule
            '   positions for them.
            If op_count > 0 Then
                cur_page_gp = of_.op(op_count - 1).granulepos
                cur_page_eos = of_.op(op_count - 1).e_o_s
                prev_packet_gp = of_.prev_packet_gp
                If prev_packet_gp = -1 Then
                    ' This is the first call after a raw seek.
                    '  Try to reconstruct prev_packet_gp from scratch.
                    ' (OP_ASSERT(seekable))
                    If cur_page_eos Then
                        ' If the first page we hit after our seek was the EOS page, and
                        '   we didn't start from data_offset or before, we don't have
                        '   enough information to do end-trimming.
                        '  Proceed to the next link, rather than risk playing back some
                        '   samples that shouldn't have been played.
                        of_.op_count = 0
                        If report_hole Then op_fetch_and_process_page = OP_HOLE: Exit Function
                        GoTo continue_
                    End If
                    ' By default discard 80 ms of data after a seek, unless we seek
                    '   into the pre-skip region.
                    cur_discard_count = 80 * 48
                    cur_page_gp = of_.op(op_count - 1).granulepos
                    ' Try to initialize prev_packet_gp.
                    '  If the current page had packets but didn't have a granule
                    '   position, or the granule position it had was too small (both
                    '   illegal), just use the starting granule position for the link.
                    prev_packet_gp = of_.links(cur_link).pcm_start
                    If cur_page_gp <> -1 Then
                        op_granpos_add prev_packet_gp, cur_page_gp, -total_duration
                    End If
                    If op_granpos_diff(diff, prev_packet_gp, of_.links(cur_link).pcm_start) = 0 Then
                        ' If we start at the beginning of the pre-skip region, or we're
                        '   at least 80 ms from the end of the pre-skip region, we discard
                        '   to the end of the pre-skip region.
                        '  Otherwise, we still use the 80 ms default, which will discard
                        '   past the end of the pre-skip region.
                        pre_skip = of_.links(cur_link).head.pre_skip
                        mx = pre_skip - 80 * 48                 ' OP_MAX(0,pre_skip-80*48)
                        If 0 > mx Then mx = 0
                        If diff >= 0 And diff <= mx Then
                            cur_discard_count = pre_skip - diff
                        End If
                    End If
                    of_.cur_discard_count = cur_discard_count
                End If
                If cur_page_gp = -1 Then
                    ' This page had completed packets but didn't have a valid granule
                    '   position.
                    '  This is illegal, but we'll try to handle it by continuing to count
                    '   forwards from the previous page.
                    If op_granpos_add(cur_page_gp, prev_packet_gp, total_duration) < 0 Then
                        ' The timestamp for this page overflowed.
                        cur_page_gp = of_.links(cur_link).pcm_end
                    End If
                End If
                ' If we hit the last page, handle end-trimming.
                doTrim = False
                If cur_page_eos Then
                    If op_granpos_diff(diff, cur_page_gp, prev_packet_gp) = 0 Then
                        If diff < total_duration Then doTrim = True
                    End If
                End If
                If doTrim Then
                    cur_packet_gp = prev_packet_gp
                    For pi = 0 To op_count - 1
                        ' Check for overflow.
                        If diff < 0 And OP_INT64_MAX + diff < durations(pi) Then
                            diff = durations(pi) + 1
                        Else
                            diff = durations(pi) - diff
                        End If
                        ' If we have samples to trim...
                        If diff > 0 Then
                            ' If we trimmed the entire packet, stop (the spec says encoders
                            '   shouldn't do this, but we support it anyway).
                            If diff > durations(pi) Then Exit For
                            cur_packet_gp = cur_page_gp
                            ' Move the EOS flag to this packet, if necessary, so we'll trim
                            '   the samples during decode.
                            of_.op(pi).e_o_s = 1
                        Else
                            ' Update the granule position as normal.
                            op_granpos_add cur_packet_gp, cur_packet_gp, durations(pi)  ' (OP_ALWAYS_TRUE)
                        End If
                        of_.op(pi).granulepos = cur_packet_gp
                        op_granpos_diff diff, cur_page_gp, cur_packet_gp                ' (OP_ALWAYS_TRUE)
                    Next
                Else
                    ' Propagate timestamps to earlier packets.
                    '  op_granpos_add(&prev_packet_gp,prev_packet_gp,total_duration)
                    '   should succeed and give prev_packet_gp==cur_page_gp.
                    '  But we don't bother to check that, as there isn't much we can do
                    '   if it's not true, and it actually will not be true on the first
                    '   page after a seek, if there was a continued packet.
                    '  The only thing we guarantee is that the start and end granule
                    '   positions of the packets are valid, and that they are monotonic
                    '   within a page.
                    '  They might be completely out of range for this link (we'll check
                    '   that elsewhere), or non-monotonic between pages.
                    If op_granpos_add(prev_packet_gp, cur_page_gp, -total_duration) < 0 Then
                        ' The starting timestamp for the first packet on this page
                        '   underflowed.
                        '  This is illegal, but we ignore it.
                        prev_packet_gp = 0
                    End If
                    For pi = 0 To op_count - 1
                        If op_granpos_add(cur_packet_gp, cur_page_gp, -total_duration) < 0 Then
                            ' The start timestamp for this packet underflowed.
                            '  This is illegal, but we ignore it.
                            cur_packet_gp = 0
                        End If
                        total_duration = total_duration - durations(pi)
                        ' (OP_ASSERT(total_duration>=0))
                        op_granpos_add cur_packet_gp, cur_packet_gp, durations(pi)      ' (OP_ALWAYS_TRUE)
                        of_.op(pi).granulepos = cur_packet_gp
                    Next
                    ' (OP_ASSERT(total_duration==0))
                End If
                of_.prev_packet_gp = prev_packet_gp
                of_.prev_page_offset = page_offset_
                op_count = pi
                of_.op_count = op_count
            End If
            If report_hole Then op_fetch_and_process_page = OP_HOLE: Exit Function
            ' If end-trimming didn't trim all the packets, we're done.
            If op_count > 0 Then op_fetch_and_process_page = 0: Exit Function
        End If
continue_:
    Loop
End Function

Public Function op_raw_seek(ByVal of_ As Long, ByVal pos_ As Double) As Long
    Dim ret As Long, cur_link As Long, noPage As ogg_page
    If ofPool(of_).ready_state < OP_OPENED Then op_raw_seek = OP_EINVAL: Exit Function
    ' Don't dump the decoder state if we can't seek.
    If ofPool(of_).seekable = 0 Then op_raw_seek = OP_ENOSEEK: Exit Function
    If pos_ < 0 Or pos_ > ofPool(of_).end_ Then op_raw_seek = OP_EINVAL: Exit Function
    ' Clear out any buffered, decoded data.
    op_decode_clear ofPool(of_)
    ofPool(of_).bytes_tracked = 0
    ofPool(of_).samples_tracked = 0
    ret = op_seek_helper(ofPool(of_), pos_)
    If ret < 0 Then op_raw_seek = OP_EREAD: Exit Function
    ret = op_fetch_and_process_page(ofPool(of_), noPage, True, -1, 1, 1)
    ' If we hit EOF, op_fetch_and_process_page() leaves us uninitialized.
    '  Instead, jump to the end.
    If ret = OP_EOF Then
        op_decode_clear ofPool(of_)
        cur_link = ofPool(of_).nlinks - 1
        ofPool(of_).cur_link = cur_link
        ofPool(of_).prev_packet_gp = ofPool(of_).links(cur_link).pcm_end
        ofPool(of_).cur_discard_count = 0
        ret = 0
    End If
    op_raw_seek = ret
End Function

' Convert a PCM offset relative to the start of the whole stream to a granule
'   position in an individual link.
Private Function op_get_granulepos(of_ As OggOpusFile, ByVal pcm_offset_ As Double, li_ As Long) As Double
    Dim duration As Double, pcm_start As Double, pre_skip As Long, nlinks As Long, li_lo As Long, li_hi As Long
    Dim li As Long
    ' (OP_ASSERT(_pcm_offset>=0))
    nlinks = of_.nlinks
    ' (links=_of->links)
    li_lo = 0
    li_hi = nlinks
    Do
        li = li_lo + (li_hi - li_lo) \ 2
        If of_.links(li).pcm_file_offset <= pcm_offset_ Then li_lo = li Else li_hi = li
    Loop While li_hi - li_lo > 1
    pcm_offset_ = pcm_offset_ - of_.links(li_lo).pcm_file_offset
    pcm_start = of_.links(li_lo).pcm_start
    pre_skip = of_.links(li_lo).head.pre_skip
    op_granpos_diff duration, of_.links(li_lo).pcm_end, pcm_start      ' (OP_ALWAYS_TRUE)
    duration = duration - pre_skip
    If pcm_offset_ >= duration Then op_get_granulepos = -1: Exit Function
    pcm_offset_ = pcm_offset_ + pre_skip
    If pcm_start > OP_INT64_MAX - pcm_offset_ Then
        ' Adding this amount to the granule position would overflow the positive
        '   half of its 64-bit range.
        '  Since signed overflow is undefined in C, do it in a way the compiler
        '   isn't allowed to screw up.
        pcm_offset_ = pcm_offset_ - (OP_INT64_MAX - pcm_start + 1)
        pcm_start = OP_INT64_MIN
    End If
    pcm_start = pcm_start + pcm_offset_
    li_ = li_lo
    op_get_granulepos = pcm_start
End Function

' A small helper to determine if an Ogg page contains data that continues onto
'   a subsequent page.
Private Function op_page_continues(og_ As ogg_page) As Boolean
    Dim nlacing As Long
    ' (OP_ASSERT(_og->header_len>=27))
    nlacing = og_.header(26)
    ' (OP_ASSERT(_og->header_len>=27+nlacing))
    ' This also correctly handles the (unlikely) case of nlacing==0, because
    '   0!=255.
    op_page_continues = (og_.header(27 + nlacing - 1) = 255)
End Function

' A small helper to buffer the continued packet data from a page.
Private Sub op_buffer_continued_data(of_ As OggOpusFile, og_ As ogg_page)
    Dim op As ogg_packet
    ogg_stream_pagein of_.os, og_
    ' Drain any packets that did end on this page (and ignore holes).
    '  We only care about the continued packet data.
    Do While ogg_stream_packetout(of_.os, op)
    Loop
End Sub

' Note: The OP_SMALL_FOOTPRINT #define doesn't (currently) save much code size,
'   but it's meant to serve as documentation for portions of the seeking
'   algorithm that are purely optional, to aid others learning from/porting this
'   code to other contexts.
' #define OP_SMALL_FOOTPRINT (1)

' Search within link _li for the page with the highest granule position
'   preceding (or equal to) _target_gp.
'  There is a danger here: missing pages or incorrect frame number information
'   in the bitstream could make our task impossible.
'  Account for that (and report it as an error condition).
Private Function op_pcm_seek_page(of_ As OggOpusFile, ByVal target_gp_ As Double, ByVal li_ As Long) As Long
    Dim og As ogg_page, pcm_pre_skip As Double, pcm_start As Double, pcm_end As Double, best_gp As Double
    Dim diff As Double, serialno As Long, pre_skip As Long, begin As Double, end_ As Double, boundary As Double
    Dim best As Double, best_start As Double, page_offset As Double, d0 As Double, d1 As Double, d2 As Double
    Dim force_bisect As Long, buffering As Long, ret As Long, offset As Double, op_count As Long, gp As Double
    Dim prev_page_gp As Double, bisect As Double, next_boundary As Double, chunk_size As Long, diff2 As Double
    Dim has_packets As Long, useCur As Boolean
    of_.bytes_tracked = 0
    of_.samples_tracked = 0
    ' (link=_of->links+_li)
    pcm_start = of_.links(li_).pcm_start
    best_gp = pcm_start
    pcm_end = of_.links(li_).pcm_end
    serialno = of_.links(li_).serialno
    begin = of_.links(li_).data_offset
    best_start = begin
    best = begin
    page_offset = -1
    buffering = 0
    ' We discard the first 80 ms of data after a seek, so seek back that much
    '   farther.
    '  If we can't, simply seek to the beginning of the link.
    If op_granpos_add(target_gp_, target_gp_, -80 * 48) < 0 Then
        target_gp_ = pcm_start
    ElseIf op_granpos_cmp(target_gp_, pcm_start) < 0 Then
        target_gp_ = pcm_start
    End If
    ' Special case seeking to the start of the link.
    pre_skip = of_.links(li_).head.pre_skip
    op_granpos_add pcm_pre_skip, pcm_start, pre_skip                    ' (OP_ALWAYS_TRUE)
    If op_granpos_cmp(target_gp_, pcm_pre_skip) < 0 Then
        boundary = begin
        end_ = boundary
    Else
        boundary = of_.links(li_).end_offset
        end_ = boundary
        ' (!OP_SMALL_FOOTPRINT)
        ' If we were decoding from this link, we can narrow the range a bit.
        If li_ = of_.cur_link And of_.ready_state >= OP_INITSET Then
            op_count = of_.op_count
            ' The offset can be out of range if we were reading through the stream
            '   and encountered a page with the granule position for another link
            '   outside of the data range identified during link enumeration when we
            '   were opening the file.
            '  We will just ignore the current position in that case.
            '  The only way the offset can be valid _and_ we can fail the granule
            '   position checks below is if someone changed the contents of the last
            '   page since we read it.
            '  We'd be within our rights to just return OP_EBADLINK, but instead we'll
            '   simply ignore the current position in that case, too.
            offset = of_.offset
            If op_count > 0 And begin <= offset And offset <= end_ Then
                ' Make sure the timestamp is valid.
                '  The granule position might be -1 if we collected the packets from a
                '   page without a granule position after reporting a hole.
                gp = of_.op(op_count - 1).granulepos
                useCur = False
                If gp <> -1 Then
                    If op_granpos_cmp(pcm_start, gp) < 0 Then
                        If op_granpos_cmp(pcm_end, gp) > 0 Then useCur = True
                    End If
                End If
                If useCur Then
                    op_granpos_diff diff, gp, target_gp_                    ' (OP_ALWAYS_TRUE)
                    ' We only actually use the current time if either
                    '  a) We can cut off at least half the range, or
                    '  b) We're seeking sufficiently close to the current position that
                    '      it's likely to be informative.
                    '  Otherwise it appears using the whole link range to estimate the
                    '   first seek location gives better results, on average.
                    If diff < 0 Then
                        If offset - begin >= Int((end_ - begin) / 2) Or diff > -OP_CUR_TIME_THRESH Then
                            begin = offset
                            best = offset
                            pcm_start = gp
                            best_gp = gp
                            ' If we have buffered data from a continued packet, remember the
                            '   offset of the previous page's start, so that if we do wind up
                            '   having to seek back here later, we can prime the stream with
                            '   the continued packet data.
                            '  With no continued packet, we remember the end of the page.
                            If of_.os.body_returned < of_.os.body_fill Then best_start = of_.prev_page_offset Else best_start = best
                            ' If there's completed packets and data in the stream state,
                            '   prev_page_offset should always be set.
                            ' (OP_ASSERT(best_start>=0))
                            ' Buffer any continued packet data starting from here.
                            buffering = 1
                        End If
                    Else
                        ' We might get lucky and already have the packet with the target
                        '   buffered.
                        '  Worth checking.
                        '  For very small files (with all of the data in a single page,
                        '   generally 1 second or less), we can loop them continuously
                        '   without seeking at all.
                        If op_granpos_add(prev_page_gp, of_.op(0).granulepos, _
                                          -op_get_packet_duration(of_.op(0).packet, 0, of_.op(0).bytes)) < 0 Then
                            ' We validate/sanitize the per-packet timestamps, so the only way
                            '   we should fail to calculate a granule position for the
                            '   previous page is if the first page with completed packets in
                            '   the stream is also the last, and end-trimming causes the
                            '   apparent granule position preceding the first sample in the
                            '   first packet to underflow.
                            '  The starting PCM offset is then 0 by spec mandate.
                            ' (OP_ASSERT(_of->op[0].e_o_s))
                            prev_page_gp = 0
                        End If
                        If op_granpos_cmp(prev_page_gp, target_gp_) <= 0 Then
                            ' Don't call op_decode_clear(), because it will dump our
                            '   packets.
                            of_.op_pos = 0
                            of_.od_buffer_size = 0
                            of_.prev_packet_gp = prev_page_gp
                            ' _of->prev_page_offset already points to the right place.
                            of_.ready_state = OP_STREAMSET
                            op_pcm_seek_page = op_make_decode_ready(of_)
                            Exit Function
                        End If
                        ' No such luck.
                        '  Check if we can cut off at least half the range, though.
                        If offset - begin <= Int((end_ - begin) / 2) Or diff < OP_CUR_TIME_THRESH Then
                            ' We really want the page start here, but this will do.
                            boundary = offset
                            end_ = boundary
                            pcm_end = gp
                        End If
                    End If
                End If
            End If
        End If
    End If
    ' This code was originally based on the "new search algorithm by HB (Nicholas
    '   Vinen)" from libvorbisfile.
    '  It has been modified substantially since.
    op_decode_clear of_
    If buffering = 0 Then ogg_stream_reset_serialno of_.os, serialno
    of_.cur_link = li_
    of_.ready_state = OP_STREAMSET
    ' Initialize the interval size history.
    d0 = end_ - begin
    d1 = d0
    d2 = d0
    force_bisect = 0
    Do While begin < end_
        If end_ - begin < OP_CHUNK_SIZE Then
            bisect = begin
        Else
            ' Update the interval size history.
            d0 = Int(d1 / 2)
            d1 = Int(d2 / 2)
            d2 = Int((end_ - begin) / 2)
            If force_bisect Then
                bisect = begin + Int((end_ - begin) / 2)
            Else
                op_granpos_diff diff, target_gp_, pcm_start             ' (OP_ALWAYS_TRUE)
                op_granpos_diff diff2, pcm_end, pcm_start               ' (OP_ALWAYS_TRUE)
                ' Take a (pretty decent) guess.
                bisect = begin + op_rescale64(diff, diff2, end_ - begin) - OP_CHUNK_SIZE
            End If
            If bisect - OP_CHUNK_SIZE < begin Then bisect = begin
            force_bisect = 0
        End If
        If bisect <> of_.offset Then
            ' Discard any buffered continued packet data.
            If buffering Then ogg_stream_reset of_.os
            buffering = 0
            page_offset = -1
            ret = op_seek_helper(of_, bisect)
            If ret < 0 Then op_pcm_seek_page = ret: Exit Function
        End If
        chunk_size = OP_CHUNK_SIZE
        next_boundary = boundary
        ' Now scan forward and figure out where we landed.
        '  In the ideal case, we will see a page with a granule position at or
        '   before our target, followed by a page with a granule position after our
        '   target (or the end of the search interval).
        '  Then we can just drop out and will have all of the data we need with no
        '   additional seeking.
        '  If we landed too far before, or after, we'll break out and do another
        '   bisection.
        Do While begin < end_
            page_offset = op_get_next_page(of_, og, boundary)
            If page_offset < 0 Then
                If page_offset < OP_FALSE Then op_pcm_seek_page = page_offset: Exit Function
                ' There are no more pages in our interval from our stream with a valid
                '   timestamp that start at position bisect or later.
                ' If we scanned the whole interval, we're done.
                If bisect <= begin + 1 Then
                    end_ = begin
                Else
                    ' Otherwise, back up one chunk.
                    '  First, discard any data from a continued packet.
                    If buffering Then ogg_stream_reset of_.os
                    buffering = 0
                    If bisect - chunk_size > begin Then bisect = bisect - chunk_size Else bisect = begin
                    ret = op_seek_helper(of_, bisect)
                    If ret < 0 Then op_pcm_seek_page = ret: Exit Function
                    ' Bump up the chunk size.
                    If 2 * chunk_size < OP_CHUNK_SIZE_MAX Then chunk_size = 2 * chunk_size Else chunk_size = OP_CHUNK_SIZE_MAX
                    ' If we did find a page from another stream or without a timestamp,
                    '   don't read past it.
                    boundary = next_boundary
                End If
            Else
                ' Save the offset of the first page we found after the seek, regardless
                '   of the stream it came from or whether or not it has a timestamp.
                If page_offset < next_boundary Then next_boundary = page_offset
                If serialno <> ogg_page_serialno(og) Then GoTo continue_inner
                has_packets = -(ogg_page_packets(og) > 0)
                ' Force the gp to -1 (as it should be per spec) if no packets end on
                '   this page.
                '  Otherwise we might get confused when we try to pull out a packet
                '   with that timestamp and can't find it.
                If has_packets Then gp = ogg_page_granulepos(og) Else gp = -1
                If gp = -1 Then
                    If buffering Then
                        If has_packets = 0 Then
                            ogg_stream_pagein of_.os, og
                        Else
                            ' If packets did end on this page, but we still didn't have a
                            '   valid granule position (in violation of the spec!), stop
                            '   buffering continued packet data.
                            '  Otherwise we might continue past the packet we actually
                            '   wanted.
                            ogg_stream_reset of_.os
                            buffering = 0
                        End If
                    End If
                    GoTo continue_inner
                End If
                If op_granpos_cmp(gp, target_gp_) < 0 Then
                    ' We found a page that ends before our target.
                    '  Advance to the raw offset of the next page.
                    begin = of_.offset
                    If op_granpos_cmp(pcm_start, gp) > 0 Or op_granpos_cmp(pcm_end, gp) < 0 Then
                        ' Don't let pcm_start get out of range!
                        '  That could happen with an invalid timestamp.
                        Exit Do
                    End If
                    ' Save the byte offset of the end of the page with this granule
                    '   position.
                    best = begin
                    best_start = begin
                    ' Buffer any data from a continued packet, if necessary.
                    '  This avoids the need to seek back here if the next timestamp we
                    '   encounter while scanning forward lies after our target.
                    If buffering Then ogg_stream_reset of_.os
                    If op_page_continues(og) Then
                        op_buffer_continued_data of_, og
                        ' If we have a continued packet, remember the offset of this
                        '   page's start, so that if we do wind up having to seek back here
                        '   later, we can prime the stream with the continued packet data.
                        '  With no continued packet, we remember the end of the page.
                        best_start = page_offset
                    End If
                    ' Then force buffering on, so that if a packet starts (but does not
                    '   end) on the next page, we still avoid the extra seek back.
                    buffering = 1
                    pcm_start = gp
                    best_gp = gp
                    op_granpos_diff diff, target_gp_, pcm_start         ' (OP_ALWAYS_TRUE)
                    ' If we're more than a second away from our target, break out and
                    '   do another bisection.
                    If diff > 48000 Then Exit Do
                    ' Otherwise, keep scanning forward (do NOT use begin+1).
                    bisect = begin
                Else
                    ' We found a page that ends after our target.
                    ' If we scanned the whole interval before we found it, we're done.
                    If bisect <= begin + 1 Then
                        end_ = begin
                    Else
                        end_ = bisect
                        ' In later iterations, don't read past the first page we found.
                        boundary = next_boundary
                        ' If we're not making much progress shrinking the interval size,
                        '   start forcing straight bisection to limit the worst case.
                        force_bisect = -(end_ - begin > d0 * 2)
                        ' Don't let pcm_end get out of range!
                        '  That could happen with an invalid timestamp.
                        If op_granpos_cmp(pcm_end, gp) > 0 And op_granpos_cmp(pcm_start, gp) <= 0 Then
                            pcm_end = gp
                        End If
                        Exit Do
                    End If
                End If
            End If
continue_inner:
        Loop
    Loop
    ' Found our page.
    ' (OP_ASSERT(op_granpos_cmp(best_gp,pcm_start)>=0))
    ' Seek, if necessary.
    '  If we were buffering data from a continued packet, we should be able to
    '   continue to scan forward to get the rest of the data (even if
    '   page_offset==-1).
    '  Otherwise, we need to seek back to best_start.
    If buffering = 0 Then
        If best_start <> page_offset Then
            page_offset = -1
            ret = op_seek_helper(of_, best_start)
            If ret < 0 Then op_pcm_seek_page = ret: Exit Function
        End If
        If best_start < best Then
            ' Retrieve the page at best_start, if we do not already have it.
            If page_offset < 0 Then
                page_offset = op_get_next_page(of_, og, of_.links(li_).end_offset)
                If page_offset < OP_FALSE Then op_pcm_seek_page = page_offset: Exit Function
                If page_offset <> best_start Then op_pcm_seek_page = OP_EBADLINK: Exit Function
            End If
            op_buffer_continued_data of_, og
            page_offset = -1
        End If
    End If
    ' Update prev_packet_gp to allow per-packet granule position assignment.
    of_.prev_packet_gp = best_gp
    of_.prev_page_offset = best_start
    ret = op_fetch_and_process_page(of_, og, page_offset < 0, page_offset, 0, 1)
    If ret < 0 Then op_pcm_seek_page = OP_EBADLINK: Exit Function
    ' Verify result.
    If op_granpos_cmp(of_.prev_packet_gp, target_gp_) > 0 Then op_pcm_seek_page = OP_EBADLINK: Exit Function
    ' Our caller will set cur_discard_count to handle pre-roll.
    op_pcm_seek_page = 0
End Function

Public Function op_pcm_seek(ByVal of_ As Long, ByVal pcm_offset_ As Double) As Long
    Dim pcm_start As Double, target_gp As Double, prev_packet_gp As Double, skip As Double, diff As Double
    Dim op_count As Long, op_pos As Long, ret As Long, li As Long, gp As Double, discard_count As Double
    Dim nbuffered As Long, cur_packet_gp As Double, noPage As ogg_page, found As Boolean
    If ofPool(of_).ready_state < OP_OPENED Then op_pcm_seek = OP_EINVAL: Exit Function
    If ofPool(of_).seekable = 0 Then op_pcm_seek = OP_ENOSEEK: Exit Function
    If pcm_offset_ < 0 Then op_pcm_seek = OP_EINVAL: Exit Function
    target_gp = op_get_granulepos(ofPool(of_), pcm_offset_, li)
    If target_gp = -1 Then op_pcm_seek = OP_EINVAL: Exit Function
    ' (link=_of->links+li)
    pcm_start = ofPool(of_).links(li).pcm_start
    op_granpos_diff pcm_offset_, target_gp, pcm_start                   ' (OP_ALWAYS_TRUE)
    ' (!OP_SMALL_FOOTPRINT)
    ' For small (90 ms or less) forward seeks within the same link, just decode
    '   forward.
    '  This also optimizes the case of seeking to the current position.
    If li = ofPool(of_).cur_link And ofPool(of_).ready_state >= OP_INITSET Then
        gp = ofPool(of_).prev_packet_gp
        If gp <> -1 Then
            nbuffered = ofPool(of_).od_buffer_size - ofPool(of_).od_buffer_pos
            If nbuffered < 0 Then nbuffered = 0
            op_granpos_add gp, gp, -nbuffered                           ' (OP_ALWAYS_TRUE)
            ' We do _not_ add cur_discard_count to gp.
            '  Otherwise the total amount to discard could grow without bound, and it
            '   would be better just to do a full seek.
            If op_granpos_diff(discard_count, target_gp, gp) = 0 Then
                ' We use a threshold of 90 ms instead of 80, since 80 ms is the
                '   _minimum_ we would have discarded after a full seek.
                '  Assuming 20 ms frames (the default), we'd discard 90 ms on average.
                If discard_count >= 0 And discard_count < 90 * 48 Then
                    ofPool(of_).cur_discard_count = discard_count
                    op_pcm_seek = 0
                    Exit Function
                End If
            End If
        End If
    End If
    ret = op_pcm_seek_page(ofPool(of_), target_gp, li)
    If ret < 0 Then op_pcm_seek = ret: Exit Function
    ' Now skip samples until we actually get to our target.
    ' Figure out where we should skip to.
    If pcm_offset_ <= ofPool(of_).links(li).head.pre_skip Then
        skip = 0
    ElseIf pcm_offset_ - 80 * 48 > 0 Then
        skip = pcm_offset_ - 80 * 48
    Else
        skip = 0
    End If
    ' (OP_ASSERT(_pcm_offset-skip>=0); OP_ASSERT(_pcm_offset-skip<OP_INT32_MAX-120*48))
    ' Skip packets until we find one with samples past our skip target.
    Do
        op_count = ofPool(of_).op_count
        prev_packet_gp = ofPool(of_).prev_packet_gp
        For op_pos = ofPool(of_).op_pos To op_count - 1
            cur_packet_gp = ofPool(of_).op(op_pos).granulepos
            found = False
            If op_granpos_diff(diff, cur_packet_gp, pcm_start) = 0 Then
                If diff > skip Then found = True
            End If
            If found Then Exit For
            prev_packet_gp = cur_packet_gp
        Next
        ofPool(of_).prev_packet_gp = prev_packet_gp
        ofPool(of_).op_pos = op_pos
        If op_pos < op_count Then Exit Do
        ' We skipped all the packets on this page.
        '  Fetch another.
        ret = op_fetch_and_process_page(ofPool(of_), noPage, True, -1, 0, 1)
        If ret < 0 Then op_pcm_seek = OP_EBADLINK: Exit Function
    Loop
    ' We skipped too far, or couldn't get within 2 billion samples of the target.
    '  Either the timestamps were illegal or there was a hole in the data.
    If op_granpos_diff(diff, prev_packet_gp, pcm_start) Then op_pcm_seek = OP_EBADLINK: Exit Function
    If diff > skip Or pcm_offset_ - diff >= OP_INT32_MAX Then op_pcm_seek = OP_EBADLINK: Exit Function
    ' TODO: If there are further holes/illegal timestamps, we still won't decode
    '   to the correct sample.
    '  However, at least op_pcm_tell() will report the correct value immediately
    '   after returning.
    ofPool(of_).cur_discard_count = pcm_offset_ - diff
    op_pcm_seek = 0
End Function

Public Function op_raw_tell(ByVal of_ As Long) As Double
    If ofPool(of_).ready_state < OP_OPENED Then op_raw_tell = OP_EINVAL: Exit Function
    op_raw_tell = ofPool(of_).offset
End Function

' Convert a granule position from a given link to a PCM offset relative to the
'   start of the whole stream.
'  For unseekable sources, this gets reset to 0 at the beginning of each link.
Private Function op_get_pcm_offset(of_ As OggOpusFile, ByVal gp_ As Double, ByVal li_ As Long) As Double
    Dim pcm_offset As Double, delta As Double
    ' (links=_of->links)
    ' (OP_ASSERT(_li>=0&&_li<_of->nlinks))
    pcm_offset = of_.links(li_).pcm_file_offset
    If of_.seekable Then
        If op_granpos_cmp(gp_, of_.links(li_).pcm_end) > 0 Then gp_ = of_.links(li_).pcm_end
    End If
    If op_granpos_cmp(gp_, of_.links(li_).pcm_start) > 0 Then
        If op_granpos_diff(delta, gp_, of_.links(li_).pcm_start) < 0 Then
            ' This means an unseekable stream claimed to have a page from more than
            '   2 billion days after we joined.
            ' (OP_ASSERT(!_of->seekable))
            op_get_pcm_offset = OP_INT64_MAX
            Exit Function
        End If
        If delta < of_.links(li_).head.pre_skip Then delta = 0 Else delta = delta - of_.links(li_).head.pre_skip
        ' In the seekable case, _gp was limited by pcm_end.
        '  In the unseekable case, pcm_offset should be 0.
        ' (OP_ASSERT(pcm_offset<=OP_INT64_MAX-delta))
        pcm_offset = pcm_offset + delta
    End If
    op_get_pcm_offset = pcm_offset
End Function

Public Function op_pcm_tell(ByVal of_ As Long) As Double
    Dim gp As Double, nbuffered As Long, li As Long
    If ofPool(of_).ready_state < OP_OPENED Then op_pcm_tell = OP_EINVAL: Exit Function
    gp = ofPool(of_).prev_packet_gp
    If gp = -1 Then op_pcm_tell = 0: Exit Function
    nbuffered = ofPool(of_).od_buffer_size - ofPool(of_).od_buffer_pos
    If nbuffered < 0 Then nbuffered = 0
    op_granpos_add gp, gp, -nbuffered                                   ' (OP_ALWAYS_TRUE)
    If ofPool(of_).seekable Then li = ofPool(of_).cur_link Else li = 0
    If op_granpos_add(gp, gp, ofPool(of_).cur_discard_count) < 0 Then
        gp = ofPool(of_).links(li).pcm_end
    End If
    op_pcm_tell = op_get_pcm_offset(ofPool(of_), gp, li)
End Function

' (VB) _decode_cb is a number of op_vbCallDecode (0 = NULL)
Public Sub op_set_decode_callback(ByVal of_ As Long, ByVal decode_cb_ As Long, ctx_ As Variant)
    ofPool(of_).decode_cb = decode_cb_
    op_vbSetVoid ofPool(of_).decode_cb_ctx, ctx_
End Sub

Public Function op_set_gain_offset(ByVal of_ As Long, ByVal gain_type_ As Long, ByVal gain_offset_q8_ As Long) As Long
    If gain_type_ <> OP_HEADER_GAIN And gain_type_ <> OP_ALBUM_GAIN _
       And gain_type_ <> OP_TRACK_GAIN And gain_type_ <> OP_ABSOLUTE_GAIN Then
        op_set_gain_offset = OP_EINVAL
        Exit Function
    End If
    ofPool(of_).gain_type = gain_type_
    ' The sum of header gain and track gain lies in the range [-65536,65534].
    '  These bounds allow the offset to set the final value to anywhere in the
    '   range [-32768,32767], which is what we'll clamp it to before applying.
    If gain_offset_q8_ > 98303 Then gain_offset_q8_ = 98303
    If gain_offset_q8_ < -98302 Then gain_offset_q8_ = -98302
    ofPool(of_).gain_offset_q8 = gain_offset_q8_
    op_update_gain ofPool(of_)
    op_set_gain_offset = 0
End Function

Public Sub op_set_dither_enabled(ByVal of_ As Long, ByVal enabled_ As Long)
    ' (!OP_FIXED_POINT)
    ofPool(of_).dither_disabled = -(enabled_ = 0)
    If enabled_ = 0 Then ofPool(of_).dither_mute = 65
End Sub

' Allocate the decoder scratch buffer.
'  This is done lazily, since if the user provides large enough buffers, we'll
'   never need it.
Private Function op_init_buffer(of_ As OggOpusFile) As Long
    Dim nchannels_max As Long, nlinks As Long, li As Long
    If of_.seekable Then
        ' (links=_of->links)
        nlinks = of_.nlinks
        nchannels_max = 1
        For li = 0 To nlinks - 1
            If of_.links(li).head.channel_count > nchannels_max Then nchannels_max = of_.links(li).head.channel_count
        Next
    Else
        nchannels_max = OP_NCHANNELS_MAX
    End If
    ReDim of_.od_buffer(0 To nchannels_max * 120 * 48 - 1)
    of_.od_buffer_allocated = True
    op_init_buffer = 0
End Function

' (VB) the application's decode callback by its number: none is registered in the port (WZ sets none), so
' every number lets libopusfile decode
Private Function op_vbCallDecode(ByVal decode_cb_ As Long, ctx_ As Variant, ByVal od_ As Long, pcm_() As Single, _
                                 ByVal pcmOff As Long, op_ As ogg_packet, ByVal nsamples_ As Long, ByVal nchannels_ As Long, _
                                 ByVal format_ As Long, ByVal li_ As Long) As Long
    op_vbCallDecode = OP_DEC_USE_DEFAULT
End Function

' Decode a single packet into the target buffer.
Private Function op_decode(of_ As OggOpusFile, pcm_() As Single, ByVal pcmOff As Long, op_ As ogg_packet, _
                           ByVal nsamples_ As Long, ByVal nchannels_ As Long) As Long
    Dim ret As Long
    ' First we try using the application-provided decode callback.
    If of_.decode_cb <> 0 Then
        ' (!OP_FIXED_POINT)
        ret = op_vbCallDecode(of_.decode_cb, of_.decode_cb_ctx, of_.od, pcm_, pcmOff, op_, nsamples_, nchannels_, _
                              OP_DEC_FORMAT_FLOAT, of_.cur_link)
    Else
        ret = OP_DEC_USE_DEFAULT
    End If
    ' If the application didn't want to handle decoding, do it ourselves.
    If ret = OP_DEC_USE_DEFAULT Then
        ' (!OP_FIXED_POINT)
        ret = opus_multistream_decode_float(of_.od, op_.packet, 0, op_.bytes, pcm_, pcmOff, nsamples_, 0)
        ' (OP_ASSERT(ret<0||ret==_nsamples))
    ' If the application returned a positive value other than 0 or
    '   OP_DEC_USE_DEFAULT, fail.
    ElseIf ret > 0 Then
        op_decode = OP_EBADPACKET
        Exit Function
    End If
    If ret < 0 Then op_decode = OP_EBADPACKET: Exit Function
    op_decode = ret
End Function

' Read more samples from the stream, using the same API as op_read() or
'   op_read_float().
' (VB) _pcm is pcm_(pcmOff ...); _li NULL: li_ is set anyway
Private Function op_read_native(of_ As OggOpusFile, pcm_() As Single, ByVal pcmOff As Long, ByVal buf_size_ As Long, _
                                li_ As Long) As Long
    Dim ret As Long, nchannels As Long, od_buffer_pos As Long, nsamples As Long, op_pos As Long, popI As Long
    Dim diff As Double, cur_discard_count As Long, duration As Long, trimmed_duration As Long, noPage As ogg_page
    If of_.ready_state < OP_OPENED Then op_read_native = OP_EINVAL: Exit Function
    Do
        If of_.ready_state >= OP_INITSET Then
            If of_.seekable Then nchannels = of_.links(of_.cur_link).head.channel_count Else nchannels = of_.links(0).head.channel_count
            od_buffer_pos = of_.od_buffer_pos
            nsamples = of_.od_buffer_size - od_buffer_pos
            ' If we have buffered samples, return them.
            If nsamples > 0 Then
                If nsamples * nchannels > buf_size_ Then nsamples = buf_size_ \ nchannels
                ' (OP_ASSERT(_pcm!=NULL||nsamples<=0))
                ' Check nsamples again so we don't pass NULL to memcpy() if _buf_size
                '   is zero.
                '  That would technically be undefined behavior, even if the number of
                '   bytes to copy were zero.
                If nsamples > 0 Then
                    CopyMemory pcm_(pcmOff), of_.od_buffer(nchannels * od_buffer_pos), 4 * nchannels * nsamples
                    od_buffer_pos = od_buffer_pos + nsamples
                    of_.od_buffer_pos = od_buffer_pos
                End If
                li_ = of_.cur_link
                op_read_native = nsamples
                Exit Function
            End If
            ' If we have buffered packets, decode one.
            op_pos = of_.op_pos
            If op_pos < of_.op_count Then
                popI = op_pos                           ' pop=_of->op+op_pos++
                op_pos = op_pos + 1
                of_.op_pos = op_pos
                cur_discard_count = of_.cur_discard_count
                duration = op_get_packet_duration(of_.op(popI).packet, 0, of_.op(popI).bytes)
                ' We don't buffer packets with an invalid TOC sequence.
                ' (OP_ASSERT(duration>0))
                trimmed_duration = duration
                ' Perform end-trimming.
                If of_.op(popI).e_o_s Then
                    If op_granpos_cmp(of_.op(popI).granulepos, of_.prev_packet_gp) <= 0 Then
                        trimmed_duration = 0
                    ElseIf op_granpos_diff(diff, of_.op(popI).granulepos, of_.prev_packet_gp) = 0 Then
                        If diff < trimmed_duration Then trimmed_duration = diff
                    End If
                End If
                of_.prev_packet_gp = of_.op(popI).granulepos
                If duration * nchannels > buf_size_ Then
                    ' If the user's buffer is too small, decode into a scratch buffer.
                    ' (buf=_of->od_buffer)
                    If Not of_.od_buffer_allocated Then
                        ret = op_init_buffer(of_)
                        If ret < 0 Then op_read_native = ret: Exit Function
                    End If
                    ret = op_decode(of_, of_.od_buffer, 0, of_.op(popI), duration, nchannels)
                    If ret < 0 Then op_read_native = ret: Exit Function
                    ' Perform pre-skip/pre-roll.
                    If trimmed_duration < cur_discard_count Then od_buffer_pos = trimmed_duration Else od_buffer_pos = cur_discard_count
                    cur_discard_count = cur_discard_count - od_buffer_pos
                    of_.cur_discard_count = cur_discard_count
                    of_.od_buffer_pos = od_buffer_pos
                    of_.od_buffer_size = trimmed_duration
                    ' Update bitrate tracking based on the actual samples we used from
                    '   what was decoded.
                    of_.bytes_tracked = of_.bytes_tracked + of_.op(popI).bytes
                    of_.samples_tracked = of_.samples_tracked + (trimmed_duration - od_buffer_pos)
                Else
                    ' (OP_ASSERT(_pcm!=NULL))
                    ' Otherwise decode directly into the user's buffer.
                    ret = op_decode(of_, pcm_, pcmOff, of_.op(popI), duration, nchannels)
                    If ret < 0 Then op_read_native = ret: Exit Function
                    If trimmed_duration > 0 Then
                        ' Perform pre-skip/pre-roll.
                        If trimmed_duration < cur_discard_count Then od_buffer_pos = trimmed_duration Else od_buffer_pos = cur_discard_count
                        cur_discard_count = cur_discard_count - od_buffer_pos
                        of_.cur_discard_count = cur_discard_count
                        trimmed_duration = trimmed_duration - od_buffer_pos
                        If trimmed_duration > 0 And od_buffer_pos > 0 Then
                            CopyMemory pcm_(pcmOff), pcm_(pcmOff + od_buffer_pos * nchannels), 4 * trimmed_duration * nchannels
                        End If
                        ' Update bitrate tracking based on the actual samples we used from
                        '   what was decoded.
                        of_.bytes_tracked = of_.bytes_tracked + of_.op(popI).bytes
                        of_.samples_tracked = of_.samples_tracked + trimmed_duration
                        If trimmed_duration > 0 Then
                            li_ = of_.cur_link
                            op_read_native = trimmed_duration
                            Exit Function
                        End If
                    End If
                End If
                ' Don't grab another page yet.
                '  This one might have more packets, or might have buffered data now.
                GoTo continue_
            End If
        End If
        ' Suck in another page.
        ret = op_fetch_and_process_page(of_, noPage, True, -1, 1, 0)
        If ret = OP_EOF Then
            li_ = of_.cur_link
            op_read_native = 0
            Exit Function
        End If
        If ret < 0 Then op_read_native = ret: Exit Function
continue_:
    Loop
End Function

' A generic filter to apply to the decoded audio data.
'  _src is non-const because we will destructively modify the contents of the
'   source buffer that we consume in some cases.
' (VB) typedef int (*op_read_filter_func)(...): the FILTER_* numbers; _dst is dstS (opus_int16) or dstF (float)

' Decode some samples and then apply a custom filter to them.
'  This is used to convert to different output formats.
Private Function op_filter_read_native(of_ As OggOpusFile, dstS_() As Integer, dstF_() As Single, ByVal dstOff As Long, _
                                       ByVal dst_sz_ As Long, ByVal filter_ As Long, li_ As Long) As Long
    Static noPcm() As Single
    Dim ret As Long, od_buffer_pos As Long, nchannels As Long
    ' Ensure we have some decoded samples in our buffer.
    ret = op_read_native(of_, noPcm, 0, 0, li_)
    ' Now apply the filter to them.
    If ret >= 0 And of_.ready_state >= OP_INITSET Then
        od_buffer_pos = of_.od_buffer_pos
        ret = of_.od_buffer_size - od_buffer_pos
        If ret > 0 Then
            If of_.seekable Then nchannels = of_.links(of_.cur_link).head.channel_count Else nchannels = of_.links(0).head.channel_count
            Select Case filter_
            Case FILTER_op_float2short_filter
                ret = op_float2short_filter(of_, dstS_, dstOff, dst_sz_, of_.od_buffer, nchannels * od_buffer_pos, ret, nchannels)
            Case FILTER_op_float2short_stereo_filter
                ret = op_float2short_stereo_filter(of_, dstS_, dstOff, dst_sz_, of_.od_buffer, nchannels * od_buffer_pos, ret, nchannels)
            Case FILTER_op_stereo_filter
                ret = op_stereo_filter(of_, dstF_, dstOff, dst_sz_, of_.od_buffer, nchannels * od_buffer_pos, ret, nchannels)
            End Select
            ' (OP_ASSERT(ret>=0); OP_ASSERT(ret<=_of->od_buffer_size-od_buffer_pos))
            od_buffer_pos = od_buffer_pos + ret
            of_.od_buffer_pos = od_buffer_pos
        End If
    End If
    op_filter_read_native = ret
End Function

' (!OP_FIXED_POINT: the float build; op_float2int is lrintf (OP_HAVE_LRINTF))

' (VB) op_rand: static opus_uint32 op_rand(opus_uint32 _seed){ return _seed*96314165+907633515&0xFFFFFFFFU; }
' is written out in op_float2short_filter (twice per sample): the product modulo 2^32 from the seed's 16-bit
' halves (96314165 = 1469 * 65536 + 41781, 41781 = 163 * 256 + 53: the middle term modulo 2^16 in Longs), the
' sum below 2^33 in a Double

' This implements 16-bit quantization with full triangular dither and IIR noise
'   shaping.
'  The noise shaping filters were designed by Sebastian Gesemann, and are based
'   on the LAME ATH curves with flattening to limit their peak gain to 20 dB.
'  Everyone else's noise shaping filters are mildly crazy.
'  The 48 kHz version of this filter is just a warped version of the 44.1 kHz
'   filter and probably could be improved by shifting the HF shelf up in
'   frequency a little bit, since 48 kHz has a bit more room and being more
'   conservative against bat-ears is probably more important than more noise
'   suppression.
'  This process can increase the peak level of the signal (in theory by the peak
'   error of 1.5 +20 dB, though that is unobservably rare).
'  To avoid clipping, the signal is attenuated by a couple thousandths of a dB.
'  Initially, the approach taken here was to only attenuate by the 99.9th
'   percentile, making clipping rare but not impossible (like SoX), but the
'   limited gain of the filter means that the worst case was only two
'   thousandths of a dB more, so this just uses the worst case.
'  The attenuation is probably also helpful to prevent clipping in the DAC
'   reconstruction filters or downstream resampling, in any case.

' (VB) dst is dst_(dstOff ...), _src src_(srcOff ...)
Private Function op_float2short_filter(of_ As OggOpusFile, dst_() As Integer, ByVal dstOff As Long, ByVal dst_sz_ As Long, _
                                       src_() As Single, ByVal srcOff As Long, ByVal nsamples_ As Long, ByVal nchannels_ As Long) As Long
    Dim ci As Long, i As Long, seed As Long, mute As Long, silent As Long, r As Single, s As Single, err_ As Single
    Dim si As Long, j As Long, t1 As Single, t2 As Single, x As Single, n As Long, k As Long
    Dim hi As Long, lo As Long, m As Long, d As Double, rf As Single
    If Not tablesReady Then op_vbInitTables
    ' (dst=(opus_int16 *)_dst)
    If nsamples_ * nchannels_ > dst_sz_ Then nsamples_ = dst_sz_ \ nchannels_
    ' (OP_SOFT_CLIP)
    If of_.state_channel_count <> nchannels_ Then
        For ci = 0 To nchannels_ - 1
            of_.clip_state(ci) = 0
        Next
    End If
    opus_pcm_soft_clip src_, srcOff, nsamples_, nchannels_, of_.clip_state, 0
    If of_.dither_disabled Then
        n = nchannels_ * nsamples_
        For i = 0 To n - 1
            ' op_float2int(OP_CLAMP(-32768,32768.0F*_src[i],32767))
            x = 32768! * src_(srcOff + i)
            If Not (x < 32767!) Then x = 32767!
            If -32768! > x Then x = -32768!
            dst_(dstOff + i) = CLng(x)
        Next
    Else
        seed = of_.dither_seed
        mute = of_.dither_mute
        If of_.state_channel_count <> nchannels_ Then mute = 65
        ' In order to avoid replacing digital silence with quiet dither noise, we
        '   mute if the output has been silent for a while.
        If mute > 64 Then
            For k = 0 To 4 * nchannels_ - 1     ' memset(_of->dither_a,0,sizeof(*_of->dither_a)*4*_nchannels)
                of_.dither_a(k) = 0
            Next
        End If
        For i = 0 To nsamples_ - 1
            silent = 1
            For ci = 0 To nchannels_ - 1
                s = src_(srcOff + nchannels_ * i + ci)
                If s <> 0 Then silent = 0                       ' silent&=s==0
                s = s * OP_GAIN
                err_ = 0
                For j = 0 To 3
                    ' err+=OP_FCOEF_B[j]*_of->dither_b[ci*4+j]-OP_FCOEF_A[j]*_of->dither_a[ci*4+j]
                    t1 = OP_FCOEF_B(j) * of_.dither_b(ci * 4 + j)
                    t2 = OP_FCOEF_A(j) * of_.dither_a(ci * 4 + j)
                    t1 = t1 - t2
                    err_ = err_ + t1
                Next
                For j = 2 To 0 Step -1                          ' for(j=3;j-->0;)
                    of_.dither_a(ci * 4 + j + 1) = of_.dither_a(ci * 4 + j)
                Next
                For j = 2 To 0 Step -1
                    of_.dither_b(ci * 4 + j + 1) = of_.dither_b(ci * 4 + j)
                Next
                of_.dither_a(ci * 4) = err_
                s = s - err_
                If mute > 16 Then
                    r = 0
                Else
                    ' seed=op_rand(seed)
                    lo = seed And &HFFFF&
                    hi = (seed And &H7FFF0000) \ &H10000
                    If seed < 0 Then hi = hi + &H8000&
                    m = (hi * 53 + ((hi * 163) And &HFF) * 256) And &HFFFF&     ' hi * 41781 modulo 2^16
                    m = (m + lo * 1469) And &HFFFF&
                    d = lo * 41781# + m * 65536# + 907633515#                   ' < 2^33
                    If d >= 4294967296# Then d = d - 4294967296#
                    If d >= 2147483648# Then seed = d - 4294967296# Else seed = d
                    ' r=seed*OP_PRNG_GAIN: (float)seed, times 2^-32
                    If seed < 0 Then rf = seed + 4294967296# Else rf = seed
                    r = rf * OP_PRNG_GAIN
                    ' seed=op_rand(seed)
                    lo = seed And &HFFFF&
                    hi = (seed And &H7FFF0000) \ &H10000
                    If seed < 0 Then hi = hi + &H8000&
                    m = (hi * 53 + ((hi * 163) And &HFF) * 256) And &HFFFF&     ' hi * 41781 modulo 2^16
                    m = (m + lo * 1469) And &HFFFF&
                    d = lo * 41781# + m * 65536# + 907633515#                   ' < 2^33
                    If d >= 4294967296# Then d = d - 4294967296#
                    If d >= 2147483648# Then seed = d - 4294967296# Else seed = d
                    ' r-=seed*OP_PRNG_GAIN
                    If seed < 0 Then rf = seed + 4294967296# Else rf = seed
                    t1 = rf * OP_PRNG_GAIN
                    r = r - t1
                End If
                ' Clamp in float out of paranoia that the input will be > 96 dBFS and
                '   wrap if the integer is clamped.
                ' si=op_float2int(OP_CLAMP(-32768,s+r,32767))
                x = s + r
                If Not (x < 32767!) Then x = 32767!
                If -32768! > x Then x = -32768!
                si = CLng(x)
                dst_(dstOff + nchannels_ * i + ci) = si
                ' Including clipping in the noise shaping is generally disastrous: the
                '   futile effort to restore the clipped energy results in more clipping.
                '  However, small amounts---at the level which could normally be created
                '   by dither and rounding---are harmless and can even reduce clipping
                '   somewhat due to the clipping sometimes reducing the dither + rounding
                '   error.
                If mute > 16 Then
                    of_.dither_b(ci * 4) = 0
                Else
                    ' OP_CLAMP(-1.5F,si-s,1.5F)
                    x = si - s
                    If Not (x < 1.5!) Then x = 1.5!
                    If -1.5! > x Then x = -1.5!
                    of_.dither_b(ci * 4) = x
                End If
            Next
            mute = mute + 1
            If silent = 0 Then mute = 0
        Next
        If mute < 65 Then of_.dither_mute = mute Else of_.dither_mute = 65
        of_.dither_seed = seed
    End If
    of_.state_channel_count = nchannels_
    op_float2short_filter = nsamples_
End Function

Public Function op_read(ByVal of_ As Long, pcm_() As Integer, ByVal pcmOff As Long, ByVal buf_size_ As Long, li_ As Long) As Long
    Static noF() As Single
    op_read = op_filter_read_native(ofPool(of_), pcm_, noF, pcmOff, buf_size_, FILTER_op_float2short_filter, li_)
End Function

Public Function op_read_float(ByVal of_ As Long, pcm_() As Single, ByVal pcmOff As Long, ByVal buf_size_ As Long, li_ As Long) As Long
    ofPool(of_).state_channel_count = 0
    op_read_float = op_read_native(ofPool(of_), pcm_, pcmOff, buf_size_, li_)
End Function

' (VB) _dst is dst_(dstOff ...), floats; the in-place downmix passes the source array as dst_ too
Private Function op_stereo_filter(of_ As OggOpusFile, dst_() As Single, ByVal dstOff As Long, ByVal dst_sz_ As Long, _
                                  src_() As Single, ByVal srcOff As Long, ByVal nsamples_ As Long, ByVal nchannels_ As Long) As Long
    Dim i As Long, l As Single, r As Single, ci As Long, t As Single, base As Long
    If Not tablesReady Then op_vbInitTables
    If dst_sz_ \ 2 < nsamples_ Then nsamples_ = dst_sz_ \ 2      ' _nsamples=OP_MIN(_nsamples,_dst_sz>>1)
    If nchannels_ = 2 Then
        If nsamples_ > 0 Then CopyMemory dst_(dstOff), src_(srcOff), nsamples_ * 2 * 4
    Else
        ' (dst=(float *)_dst)
        If nchannels_ = 1 Then
            For i = 0 To nsamples_ - 1
                dst_(dstOff + 2 * i + 1) = src_(srcOff + i)
                dst_(dstOff + 2 * i + 0) = dst_(dstOff + 2 * i + 1)
            Next
        Else
            base = (nchannels_ - 3) * 16
            For i = 0 To nsamples_ - 1
                l = 0
                r = 0
                For ci = 0 To nchannels_ - 1
                    t = OP_STEREO_DOWNMIX(base + ci * 2 + 0) * src_(srcOff + nchannels_ * i + ci)
                    l = l + t
                    t = OP_STEREO_DOWNMIX(base + ci * 2 + 1) * src_(srcOff + nchannels_ * i + ci)
                    r = r + t
                Next
                dst_(dstOff + 2 * i + 0) = l
                dst_(dstOff + 2 * i + 1) = r
            Next
        End If
    End If
    op_stereo_filter = nsamples_
End Function

Private Function op_float2short_stereo_filter(of_ As OggOpusFile, dst_() As Integer, ByVal dstOff As Long, ByVal dst_sz_ As Long, _
                                              src_() As Single, ByVal srcOff As Long, ByVal nsamples_ As Long, ByVal nchannels_ As Long) As Long
    Dim i As Long
    ' (dst=(opus_int16 *)_dst)
    If nchannels_ = 1 Then
        nsamples_ = op_float2short_filter(of_, dst_, dstOff, dst_sz_ \ 2, src_, srcOff, nsamples_, 1)
        For i = nsamples_ - 1 To 0 Step -1
            dst_(dstOff + 2 * i + 1) = dst_(dstOff + i)
            dst_(dstOff + 2 * i + 0) = dst_(dstOff + 2 * i + 1)
        Next
    Else
        If nchannels_ > 2 Then
            If dst_sz_ \ 2 < nsamples_ Then nsamples_ = dst_sz_ \ 2
            nsamples_ = op_stereo_filter(of_, src_, srcOff, nsamples_ * 2, src_, srcOff, nsamples_, nchannels_)
        End If
        nsamples_ = op_float2short_filter(of_, dst_, dstOff, dst_sz_, src_, srcOff, nsamples_, 2)
    End If
    op_float2short_stereo_filter = nsamples_
End Function

Public Function op_read_stereo(ByVal of_ As Long, pcm_() As Integer, ByVal pcmOff As Long, ByVal buf_size_ As Long) As Long
    Static noF() As Single
    Dim li As Long
    op_read_stereo = op_filter_read_native(ofPool(of_), pcm_, noF, pcmOff, buf_size_, FILTER_op_float2short_stereo_filter, li)
End Function

Public Function op_read_float_stereo(ByVal of_ As Long, pcm_() As Single, ByVal pcmOff As Long, ByVal buf_size_ As Long) As Long
    Static noS() As Integer
    Dim li As Long
    ofPool(of_).state_channel_count = 0
    op_read_float_stereo = op_filter_read_native(ofPool(of_), noS, pcm_, pcmOff, buf_size_, FILTER_op_stereo_filter, li)
End Function
