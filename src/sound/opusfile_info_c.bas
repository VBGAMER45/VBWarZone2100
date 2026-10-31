Attribute VB_Name = "opusfile_info_c"
' opusfile (lib/sound/3rdparty/opusfile at 9d718345), src/info.c: the ID header, the comment header (tags) and
' the METADATA_BLOCK_PICTURE tag; whole. opus_granule_sample is opusfile_c's (opusfile.c defines it).
' (VB) a const unsigned char * argument is an array and the offset of the C's pointer in it; a NULL OpusHead *
' or OpusTags * is the Optional headIsNull / tagsIsNull. A char * is a String of one byte per character
' (opusfile_h); a char * returned into a tag (opus_tags_query) is the C string from there, a copy.
Option Explicit

Private Const INT_MAX As Long = 2147483647

Private Function op_parse_uint16le(data_() As Byte, ByVal off As Long) As Long
    op_parse_uint16le = data_(off) Or (CLng(data_(off + 1)) * &H100&)
End Function

Private Function op_parse_int16le(data_() As Byte, ByVal off As Long) As Long
    Dim ret As Long
    ret = data_(off) Or (CLng(data_(off + 1)) * &H100&)
    op_parse_int16le = (ret Xor &H8000&) - &H8000&
End Function

' (VB) the opus_uint32's bits
Private Function op_parse_uint32le(data_() As Byte, ByVal off As Long) As Long
    Dim hi As Long
    If data_(off + 3) And &H80 Then hi = ((data_(off + 3) And &H7F) * &H1000000) Or &H80000000 Else hi = data_(off + 3) * &H1000000
    op_parse_uint32le = data_(off) Or (CLng(data_(off + 1)) * &H100&) Or (CLng(data_(off + 2)) * &H10000) Or hi
End Function

Private Function op_parse_uint32be(data_() As Byte, ByVal off As Long) As Long
    Dim hi As Long
    If data_(off) And &H80 Then hi = ((data_(off) And &H7F) * &H1000000) Or &H80000000 Else hi = data_(off) * &H1000000
    op_parse_uint32be = data_(off + 3) Or (CLng(data_(off + 2)) * &H100&) Or (CLng(data_(off + 1)) * &H10000) Or hi
End Function

' (VB) an opus_uint32's value
Private Function U32_(ByVal x As Long) As Double
    If x < 0 Then U32_ = x + 4294967296# Else U32_ = x
End Function

Public Function opus_head_parse(head_ As OpusHead, data_() As Byte, ByVal dataOff As Long, ByVal len_ As Long, _
                                Optional ByVal headIsNull As Boolean = False) As Long
    Dim head As OpusHead, size As Long, ci As Long
    If len_ < 8 Then opus_head_parse = OP_ENOTFORMAT: Exit Function
    ' if(memcmp(_data,"OpusHead",8)!=0)return OP_ENOTFORMAT
    If Not op_vbMatch(data_, dataOff, "OpusHead") Then opus_head_parse = OP_ENOTFORMAT: Exit Function
    If len_ < 9 Then opus_head_parse = OP_EBADHEADER: Exit Function
    head.version = data_(dataOff + 8)
    If head.version > 15 Then opus_head_parse = OP_EVERSION: Exit Function
    If len_ < 19 Then opus_head_parse = OP_EBADHEADER: Exit Function
    head.channel_count = data_(dataOff + 9)
    head.pre_skip = op_parse_uint16le(data_, dataOff + 10)
    head.input_sample_rate = op_parse_uint32le(data_, dataOff + 12)
    head.output_gain = op_parse_int16le(data_, dataOff + 16)
    head.mapping_family = data_(dataOff + 18)
    If head.mapping_family = 0 Then
        If head.channel_count < 1 Or head.channel_count > 2 Then opus_head_parse = OP_EBADHEADER: Exit Function
        If head.version <= 1 And len_ > 19 Then opus_head_parse = OP_EBADHEADER: Exit Function
        head.stream_count = 1
        head.coupled_count = head.channel_count - 1
        If Not headIsNull Then
            head_.mapping(0) = 0
            head_.mapping(1) = 1
        End If
    ElseIf head.mapping_family = 1 Then
        If head.channel_count < 1 Or head.channel_count > 8 Then opus_head_parse = OP_EBADHEADER: Exit Function
        size = 21 + head.channel_count
        If len_ < size Or (head.version <= 1 And len_ > size) Then opus_head_parse = OP_EBADHEADER: Exit Function
        head.stream_count = data_(dataOff + 19)
        If head.stream_count < 1 Then opus_head_parse = OP_EBADHEADER: Exit Function
        head.coupled_count = data_(dataOff + 20)
        If head.coupled_count > head.stream_count Then opus_head_parse = OP_EBADHEADER: Exit Function
        For ci = 0 To head.channel_count - 1
            If data_(dataOff + 21 + ci) >= head.stream_count + head.coupled_count _
               And data_(dataOff + 21 + ci) <> 255 Then
                opus_head_parse = OP_EBADHEADER
                Exit Function
            End If
        Next
        If Not headIsNull Then
            For ci = 0 To head.channel_count - 1        ' memcpy(_head->mapping,_data+21,head.channel_count)
                head_.mapping(ci) = data_(dataOff + 21 + ci)
            Next
        End If
    ' General purpose players should not attempt to play back content with
    '   channel mapping family 255.
    ElseIf head.mapping_family = 255 Then
        opus_head_parse = OP_EIMPL
        Exit Function
    ' No other channel mapping families are currently defined.
    Else
        opus_head_parse = OP_EBADHEADER
        Exit Function
    End If
    If Not headIsNull Then
        ' memcpy(_head,&head,head.mapping-(unsigned char *)&head): the fields before mapping
        head_.version = head.version
        head_.channel_count = head.channel_count
        head_.pre_skip = head.pre_skip
        head_.input_sample_rate = head.input_sample_rate
        head_.output_gain = head.output_gain
        head_.mapping_family = head.mapping_family
        head_.stream_count = head.stream_count
        head_.coupled_count = head.coupled_count
    End If
    opus_head_parse = 0
End Function

Public Sub opus_tags_init(tags_ As OpusTags)
    Dim blank As OpusTags
    tags_ = blank                               ' memset(_tags,0,sizeof(*_tags))
End Sub

Public Sub opus_tags_clear(tags_ As OpusTags)
    Dim ncomments As Long
    ncomments = tags_.comments
    If tags_.vbAllocated Then ncomments = ncomments + 1
    ' (else OP_ASSERT(ncomments==0))
    ' for(ci=ncomments;ci-->0;)_ogg_free(_tags->user_comments[ci]); _ogg_free(_tags->user_comments);
    ' _ogg_free(_tags->comment_lengths); _ogg_free(_tags->vendor)
    Erase tags_.user_comments
    Erase tags_.comment_lengths
    tags_.vbAllocated = False
    tags_.vendor = vbNullString
End Sub

' Ensure there's room for up to _ncomments comments.
Private Function op_tags_ensure_capacity(tags_ As OpusTags, ByVal ncomments_ As Long) As Long
    Dim cur_ncomments As Long, wasNull As Boolean
    If ncomments_ < 0 Or ncomments_ >= INT_MAX Then op_tags_ensure_capacity = OP_EFAULT: Exit Function
    ' size=sizeof(*_tags->comment_lengths)*(_ncomments+1) (it cannot overflow here)
    cur_ncomments = tags_.comments
    ' We only support growing.
    '  Trimming requires cleaning up the allocated strings in the old space, and
    '   is best handled separately if it's ever needed.
    ' (OP_ASSERT(_ncomments>=(size_t)cur_ncomments))
    wasNull = Not tags_.vbAllocated
    ' comment_lengths=(int *)_ogg_realloc(_tags->comment_lengths,size)
    If wasNull Then
        ReDim tags_.comment_lengths(0 To ncomments_)
        ' (OP_ASSERT(cur_ncomments==0))
        tags_.comment_lengths(cur_ncomments) = 0
    Else
        ReDim Preserve tags_.comment_lengths(0 To ncomments_)
    End If
    tags_.comment_lengths(ncomments_) = tags_.comment_lengths(cur_ncomments)
    ' user_comments=(char **)_ogg_realloc(_tags->user_comments,size)
    If wasNull Then
        ReDim tags_.user_comments(0 To ncomments_)
        ' (OP_ASSERT(cur_ncomments==0))
        tags_.user_comments(cur_ncomments) = vbNullString
    Else
        ReDim Preserve tags_.user_comments(0 To ncomments_)
    End If
    tags_.user_comments(ncomments_) = tags_.user_comments(cur_ncomments)
    tags_.vbAllocated = True
    op_tags_ensure_capacity = 0
End Function

' Duplicate a (possibly non-NUL terminated) string with a known length.
' (VB) the first len_ characters of s (a String of one byte per character), NULs past its end
Private Function op_strdup_with_len(s As String, ByVal len_ As Long) As String
    If Len(s) >= len_ Then
        op_strdup_with_len = Left$(s, len_)
    Else
        op_strdup_with_len = s & String$(len_ - Len(s), vbNullChar)
    End If
End Function

' (VB) strlen(s + off) of a String of one byte per character: up to its first NUL
Private Function op_vbStrlen(s As String, ByVal off As Long) As Long
    Dim p As Long
    If off >= Len(s) Then Exit Function
    p = InStr(off + 1, s, vbNullChar)
    If p = 0 Then op_vbStrlen = Len(s) - off Else op_vbStrlen = p - 1 - off
End Function

' The actual implementation of opus_tags_parse().
'  Unlike the public API, this function requires _tags to already be
'   initialized, modifies its contents before success is guaranteed, and assumes
'   the caller will clear it on error.
Private Function opus_tags_parse_impl(tags_ As OpusTags, ByVal tagsIsNull As Boolean, data_() As Byte, _
                                      ByVal off As Long, ByVal len0 As Long) As Long
    Dim count As Long, len_ As Long, ncomments As Long, ci As Long, ret As Long
    len_ = len0
    If len_ < 8 Then opus_tags_parse_impl = OP_ENOTFORMAT: Exit Function
    If Not op_vbMatch(data_, off, "OpusTags") Then opus_tags_parse_impl = OP_ENOTFORMAT: Exit Function
    If len_ < 16 Then opus_tags_parse_impl = OP_EBADHEADER: Exit Function
    off = off + 8
    len_ = len_ - 8
    count = op_parse_uint32le(data_, off)
    off = off + 4
    len_ = len_ - 4
    If U32_(count) > len_ Then opus_tags_parse_impl = OP_EBADHEADER: Exit Function
    If Not tagsIsNull Then
        tags_.vendor = op_vbBytesToStr(data_, off, count)  ' op_strdup_with_len((char *)_data,count)
        ' (if(_tags->vendor==NULL)return OP_EFAULT: VB raises an error where malloc returns NULL)
    End If
    off = off + count
    len_ = len_ - count
    If len_ < 4 Then opus_tags_parse_impl = OP_EBADHEADER: Exit Function
    count = op_parse_uint32le(data_, off)
    off = off + 4
    len_ = len_ - 4
    ' Check to make sure there's minimally sufficient data left in the packet.
    If U32_(count) > (len_ \ 4) Then opus_tags_parse_impl = OP_EBADHEADER: Exit Function
    ' Check for overflow (the API limits this to an int).
    If U32_(count) > INT_MAX - 1 Then opus_tags_parse_impl = OP_EFAULT: Exit Function
    If Not tagsIsNull Then
        ret = op_tags_ensure_capacity(tags_, count)
        If ret < 0 Then opus_tags_parse_impl = ret: Exit Function
    End If
    ncomments = count
    For ci = 0 To ncomments - 1
        ' Check to make sure there's minimally sufficient data left in the packet.
        If ncomments - ci > (len_ \ 4) Then opus_tags_parse_impl = OP_EBADHEADER: Exit Function
        count = op_parse_uint32le(data_, off)
        off = off + 4
        len_ = len_ - 4
        If U32_(count) > len_ Then opus_tags_parse_impl = OP_EBADHEADER: Exit Function
        ' Check for overflow (the API limits this to an int).
        If U32_(count) > INT_MAX Then opus_tags_parse_impl = OP_EFAULT: Exit Function
        If Not tagsIsNull Then
            tags_.user_comments(ci) = op_vbBytesToStr(data_, off, count)
            tags_.comment_lengths(ci) = count
            tags_.comments = ci + 1
            ' Needed by opus_tags_clear() if we fail before parsing the (optional)
            '   binary metadata.
            tags_.user_comments(ci + 1) = vbNullString
        End If
        off = off + count
        len_ = len_ - count
    Next
    If len_ > 0 Then
        If data_(off) And 1 Then
            ' (if(len>(opus_uint32)INT_MAX)return OP_EFAULT: len_ is a Long)
            If Not tagsIsNull Then
                tags_.user_comments(ncomments) = op_vbBytesToStr(data_, off, len_)
                tags_.comment_lengths(ncomments) = len_
            End If
        End If
    End If
    opus_tags_parse_impl = 0
End Function

Public Function opus_tags_parse(tags_ As OpusTags, data_() As Byte, ByVal dataOff As Long, ByVal len_ As Long, _
                                Optional ByVal tagsIsNull As Boolean = False) As Long
    Dim tags As OpusTags, ret As Long
    If Not tagsIsNull Then
        opus_tags_init tags
        ret = opus_tags_parse_impl(tags, False, data_, dataOff, len_)
        If ret < 0 Then
            opus_tags_clear tags
        Else
            tags_ = tags
        End If
        opus_tags_parse = ret
    Else
        opus_tags_parse = opus_tags_parse_impl(tags, True, data_, dataOff, len_)
    End If
End Function

' The actual implementation of opus_tags_copy().
'  Unlike the public API, this function requires _dst to already be
'   initialized, modifies its contents before success is guaranteed, and assumes
'   the caller will clear it on error.
Private Function opus_tags_copy_impl(dst_ As OpusTags, src_ As OpusTags) As Long
    Dim vendor As String, ncomments As Long, ret As Long, ci As Long, len_ As Long
    vendor = src_.vendor
    dst_.vendor = op_strdup_with_len(vendor, op_vbStrlen(vendor, 0))
    ncomments = src_.comments
    ret = op_tags_ensure_capacity(dst_, ncomments)
    If ret < 0 Then opus_tags_copy_impl = ret: Exit Function
    For ci = 0 To ncomments - 1
        len_ = src_.comment_lengths(ci)
        ' (OP_ASSERT(len>=0))
        dst_.user_comments(ci) = op_strdup_with_len(src_.user_comments(ci), len_)
        dst_.comment_lengths(ci) = len_
        dst_.comments = ci + 1
    Next
    If src_.vbAllocated Then
        len_ = src_.comment_lengths(ncomments)
        If len_ > 0 Then
            dst_.user_comments(ncomments) = op_strdup_with_len(src_.user_comments(ncomments), len_)
            dst_.comment_lengths(ncomments) = len_
        End If
    End If
    opus_tags_copy_impl = 0
End Function

Public Function opus_tags_copy(dst_ As OpusTags, src_ As OpusTags) As Long
    Dim dst As OpusTags, ret As Long
    opus_tags_init dst
    ret = opus_tags_copy_impl(dst, src_)
    If ret < 0 Then
        opus_tags_clear dst
    Else
        dst_ = dst
    End If
    opus_tags_copy = ret
End Function

Public Function opus_tags_add(tags_ As OpusTags, ByVal tag_ As String, ByVal value_ As String) As Long
    Dim comment As String, tag_len As Long, value_len As Long, ncomments As Long, ret As Long
    ncomments = tags_.comments
    ret = op_tags_ensure_capacity(tags_, ncomments + 1)
    If ret < 0 Then opus_tags_add = ret: Exit Function
    tag_len = op_vbStrlen(tag_, 0)
    value_len = op_vbStrlen(value_, 0)
    ' +2 for '=' and '\0'.
    If tag_len + value_len > INT_MAX - 2 Then opus_tags_add = OP_EFAULT: Exit Function
    comment = Left$(tag_, tag_len) & "=" & Left$(value_, value_len)
    tags_.user_comments(ncomments) = comment
    tags_.comment_lengths(ncomments) = tag_len + value_len + 1
    tags_.comments = ncomments + 1
    opus_tags_add = 0
End Function

Public Function opus_tags_add_comment(tags_ As OpusTags, ByVal comment_ As String) As Long
    Dim comment As String, comment_len As Long, ncomments As Long, ret As Long
    ncomments = tags_.comments
    ret = op_tags_ensure_capacity(tags_, ncomments + 1)
    If ret < 0 Then opus_tags_add_comment = ret: Exit Function
    comment_len = op_vbStrlen(comment_, 0)
    comment = op_strdup_with_len(comment_, comment_len)
    tags_.user_comments(ncomments) = comment
    tags_.comment_lengths(ncomments) = comment_len
    tags_.comments = ncomments + 1
    opus_tags_add_comment = 0
End Function

' (VB) _data NULL is dataIsNull
Public Function opus_tags_set_binary_suffix(tags_ As OpusTags, data_() As Byte, ByVal len_ As Long, _
                                            Optional ByVal dataIsNull As Boolean = False) As Long
    Dim ncomments As Long, ret As Long
    If len_ < 0 Then opus_tags_set_binary_suffix = OP_EINVAL: Exit Function
    If len_ > 0 Then
        If dataIsNull Then opus_tags_set_binary_suffix = OP_EINVAL: Exit Function
        If (data_(0) And 1) = 0 Then opus_tags_set_binary_suffix = OP_EINVAL: Exit Function
    End If
    ncomments = tags_.comments
    ret = op_tags_ensure_capacity(tags_, ncomments)
    If ret < 0 Then opus_tags_set_binary_suffix = ret: Exit Function
    ' binary_suffix_data=(unsigned char *)_ogg_realloc(_tags->user_comments[ncomments],_len); memcpy(...)
    ' ((VB) a length of 0 gives an empty suffix)
    tags_.user_comments(ncomments) = op_vbBytesToStr(data_, 0, len_)
    tags_.comment_lengths(ncomments) = len_
    opus_tags_set_binary_suffix = 0
End Function

Public Function opus_tagcompare(ByVal tag_name As String, ByVal comment_ As String) As Long
    Dim tag_len As Long
    tag_len = op_vbStrlen(tag_name, 0)
    ' (if(OP_UNLIKELY(tag_len>(size_t)INT_MAX))return -1: a Long)
    opus_tagcompare = opus_tagncompare(tag_name, tag_len, comment_)
End Function

Public Function opus_tagncompare(ByVal tag_name As String, ByVal tag_len As Long, ByVal comment_ As String) As Long
    Dim ret As Long
    ' (OP_ASSERT(_tag_len>=0))
    ret = op_strncasecmp(tag_name, 0, comment_, 0, tag_len)
    If ret Then opus_tagncompare = ret Else opus_tagncompare = Asc("=") - op_vbChar(comment_, tag_len)
End Function

' (VB) returns the value as a String (the C string the C's pointer points at), vbNullString for NULL
Public Function opus_tags_query(tags_ As OpusTags, ByVal tag_ As String, ByVal count_ As Long) As String
    Dim tag_len As Long, found As Long, ncomments As Long, ci As Long, s As String
    tag_len = op_vbStrlen(tag_, 0)
    ncomments = tags_.comments
    found = 0
    For ci = 0 To ncomments - 1
        If opus_tagncompare(tag_, tag_len, tags_.user_comments(ci)) = 0 Then
            ' We return a pointer to the data, not a copy.
            If count_ = found Then
                s = tags_.user_comments(ci)
                opus_tags_query = Mid$(s, tag_len + 2, op_vbStrlen(s, tag_len + 1))
                Exit Function
            End If
            found = found + 1
        End If
    Next
    ' Didn't find anything.
    opus_tags_query = vbNullString
End Function

Public Function opus_tags_query_count(tags_ As OpusTags, ByVal tag_ As String) As Long
    Dim tag_len As Long, found As Long, ncomments As Long, ci As Long
    tag_len = op_vbStrlen(tag_, 0)
    ncomments = tags_.comments
    found = 0
    For ci = 0 To ncomments - 1
        If opus_tagncompare(tag_, tag_len, tags_.user_comments(ci)) = 0 Then found = found + 1
    Next
    opus_tags_query_count = found
End Function

' (VB) returns the suffix as a String of one byte per character, vbNullString for NULL
Public Function opus_tags_get_binary_suffix(tags_ As OpusTags, len_ As Long) As String
    Dim ncomments As Long, n As Long
    ncomments = tags_.comments
    If tags_.vbAllocated Then n = tags_.comment_lengths(ncomments) Else n = 0
    len_ = n
    ' (OP_ASSERT(len==0||_tags->user_comments!=NULL))
    If n > 0 Then opus_tags_get_binary_suffix = tags_.user_comments(ncomments) Else opus_tags_get_binary_suffix = vbNullString
End Function

Private Function opus_tags_get_gain(tags_ As OpusTags, gain_q8_ As Long, ByVal tag_name As String, ByVal tag_len As Long) As Long
    Dim ncomments As Long, ci As Long, p As Long, s As String, gain_q8 As Long, negative As Long, c As Long
    ncomments = tags_.comments
    ' Look for the first valid tag with the name _tag_name and use that.
    For ci = 0 To ncomments - 1
        ' (OP_ASSERT(_tag_len<=(size_t)INT_MAX))
        If opus_tagncompare(tag_name, tag_len, tags_.user_comments(ci)) = 0 Then
            s = tags_.user_comments(ci)
            p = tag_len + 1                         ' p=comments[ci]+_tag_len+1
            negative = 0
            If op_vbChar(s, p) = Asc("-") Then
                negative = -1
                p = p + 1
            ElseIf op_vbChar(s, p) = Asc("+") Then
                p = p + 1
            End If
            gain_q8 = 0
            Do
                c = op_vbChar(s, p)
                If c < Asc("0") Or c > Asc("9") Then Exit Do
                gain_q8 = 10 * gain_q8 + c - Asc("0")
                If gain_q8 > 32767 - negative Then Exit Do
                p = p + 1
            Loop
            ' This didn't look like a signed 16-bit decimal integer.
            '  Not a valid gain tag.
            If op_vbChar(s, p) = 0 Then
                gain_q8_ = (gain_q8 + negative) Xor negative
                opus_tags_get_gain = 0
                Exit Function
            End If
        End If
    Next
    opus_tags_get_gain = OP_FALSE
End Function

Public Function opus_tags_get_album_gain(tags_ As OpusTags, gain_q8_ As Long) As Long
    opus_tags_get_album_gain = opus_tags_get_gain(tags_, gain_q8_, "R128_ALBUM_GAIN", 15)
End Function

Public Function opus_tags_get_track_gain(tags_ As OpusTags, gain_q8_ As Long) As Long
    opus_tags_get_track_gain = opus_tags_get_gain(tags_, gain_q8_, "R128_TRACK_GAIN", 15)
End Function

' (VB) _buf is buf(off ...), _buf_sz bytes
Private Function op_is_jpeg(buf_() As Byte, ByVal off As Long, ByVal buf_sz As Long) As Boolean
    If buf_sz >= 3 Then
        op_is_jpeg = (buf_(off) = &HFF And buf_(off + 1) = &HD8 And buf_(off + 2) = &HFF)
    End If
End Function

' Tries to extract the width, height, bits per pixel, and palette size of a
'   JPEG.
'  On failure, simply leaves its outputs unmodified.
Private Sub op_extract_jpeg_params(buf_() As Byte, ByVal off As Long, ByVal buf_sz As Long, _
                                   width_ As Long, height_ As Long, depth_ As Long, colors_ As Long, has_palette_ As Long)
    Dim offs As Long, segment_len As Long, marker As Long
    If op_is_jpeg(buf_, off, buf_sz) Then
        offs = 2
        Do
            Do While offs < buf_sz
                If buf_(off + offs) = &HFF Then Exit Do
                offs = offs + 1
            Loop
            Do While offs < buf_sz
                If buf_(off + offs) <> &HFF Then Exit Do
                offs = offs + 1
            Loop
            marker = buf_(off + offs)
            offs = offs + 1
            ' If we hit EOI* (end of image), or another SOI* (start of image),
            '   or SOS (start of scan), then stop now.
            If offs >= buf_sz Or (marker >= &HD8 And marker <= &HDA) Then
                Exit Do
            ' RST* (restart markers): skip (no segment length).
            ElseIf marker >= &HD0 And marker <= &HD7 Then
                GoTo continue_
            End If
            ' Read the length of the marker segment.
            If buf_sz - offs < 2 Then Exit Do
            segment_len = CLng(buf_(off + offs)) * &H100& Or buf_(off + offs + 1)
            If segment_len < 2 Or buf_sz - offs < segment_len Then Exit Do
            If marker = &HC0 Or (marker > &HC0 And marker < &HD0 And (marker And 3) <> 0) Then
                ' Found a SOFn (start of frame) marker segment:
                If segment_len >= 8 Then
                    height_ = CLng(buf_(off + offs + 3)) * &H100& Or buf_(off + offs + 4)
                    width_ = CLng(buf_(off + offs + 5)) * &H100& Or buf_(off + offs + 6)
                    depth_ = CLng(buf_(off + offs + 2)) * buf_(off + offs + 7)
                    colors_ = 0
                    has_palette_ = 0
                End If
                Exit Do
            End If
            ' Other markers: skip the whole marker segment.
            offs = offs + segment_len
continue_:
        Loop
    End If
End Sub

Private Function op_is_png(buf_() As Byte, ByVal off As Long, ByVal buf_sz As Long) As Boolean
    ' memcmp(_buf,"\x89PNG\x0D\x0A\x1A\x0A",8)==0
    If buf_sz >= 8 Then
        If buf_(off) = &H89 And buf_(off + 1) = &H50 And buf_(off + 2) = &H4E And buf_(off + 3) = &H47 And _
           buf_(off + 4) = &HD And buf_(off + 5) = &HA And buf_(off + 6) = &H1A And buf_(off + 7) = &HA Then
            op_is_png = True
        End If
    End If
End Function

' Tries to extract the width, height, bits per pixel, and palette size of a
'   PNG.
'  On failure, simply leaves its outputs unmodified.
Private Sub op_extract_png_params(buf_() As Byte, ByVal off As Long, ByVal buf_sz As Long, _
                                  width_ As Long, height_ As Long, depth_ As Long, colors_ As Long, has_palette_ As Long)
    Dim offs As Long, chunk_len As Long, color_type As Long, sample_depth As Long
    If op_is_png(buf_, off, buf_sz) Then
        offs = 8
        Do While buf_sz - offs >= 12
            chunk_len = op_parse_uint32be(buf_, off + offs)
            If U32_(chunk_len) > buf_sz - (offs + 12) Then
                Exit Do
            ElseIf chunk_len = 13 And op_vbMatch(buf_, off + offs + 4, "IHDR") Then
                width_ = op_parse_uint32be(buf_, off + offs + 8)
                height_ = op_parse_uint32be(buf_, off + offs + 12)
                color_type = buf_(off + offs + 17)
                If color_type = 3 Then
                    depth_ = 24
                    has_palette_ = 1
                Else
                    sample_depth = buf_(off + offs + 16)
                    If color_type = 0 Then
                        depth_ = sample_depth
                    ElseIf color_type = 2 Then
                        depth_ = sample_depth * 3
                    ElseIf color_type = 4 Then
                        depth_ = sample_depth * 2
                    ElseIf color_type = 6 Then
                        depth_ = sample_depth * 4
                    End If
                    colors_ = 0
                    has_palette_ = 0
                    Exit Do
                End If
            ElseIf has_palette_ > 0 And op_vbMatch(buf_, off + offs + 4, "PLTE") Then
                colors_ = chunk_len \ 3
                Exit Do
            End If
            offs = offs + 12 + chunk_len
        Loop
    End If
End Sub

Private Function op_is_gif(buf_() As Byte, ByVal off As Long, ByVal buf_sz As Long) As Boolean
    If buf_sz >= 6 Then
        op_is_gif = op_vbMatch(buf_, off, "GIF87a") Or op_vbMatch(buf_, off, "GIF89a")
    End If
End Function

' Tries to extract the width, height, bits per pixel, and palette size of a
'   GIF.
'  On failure, simply leaves its outputs unmodified.
Private Sub op_extract_gif_params(buf_() As Byte, ByVal off As Long, ByVal buf_sz As Long, _
                                  width_ As Long, height_ As Long, depth_ As Long, colors_ As Long, has_palette_ As Long)
    If op_is_gif(buf_, off, buf_sz) And buf_sz >= 14 Then
        width_ = buf_(off + 6) Or CLng(buf_(off + 7)) * &H100&
        height_ = buf_(off + 8) Or CLng(buf_(off + 9)) * &H100&
        ' libFLAC hard-codes the depth to 24.
        depth_ = 24
        colors_ = BIT_((buf_(off + 10) And 7) + 1)
        has_palette_ = 1
    End If
End Sub

' The actual implementation of opus_picture_tag_parse().
'  Unlike the public API, this function requires _pic to already be
'   initialized, modifies its contents before success is guaranteed, and assumes
'   the caller will clear it on error.
' (VB) _tag is tag_ from tagOff; _buf is buf_ (buf_sz + 1 bytes)
Private Function opus_picture_tag_parse_impl(pic_ As OpusPictureTag, tag_ As String, ByVal tagOff As Long, _
                                             buf_() As Byte, ByVal buf_sz As Long, ByVal base64_sz As Long) As Long
    Dim picture_type As Long, mime_type_length As Long, mime_type As String, description_length As Long
    Dim description As String, width As Long, height As Long, depth As Long, colors As Long, data_length As Long
    Dim file_width As Long, file_height As Long, file_depth As Long, file_colors As Long, format_ As Long
    Dim has_palette As Long, colors_set As Boolean, i As Long, value As Long, j As Long, c As Long, d As Long
    Dim dat() As Byte
    ' Decode the BASE64 data.
    ' (OP_ASSERT(_base64_sz>=11))
    For i = 0 To base64_sz - 1
        value = 0
        For j = 0 To 3
            c = op_vbChar(tag_, tagOff + 4 * i + j) And &HFF
            If c = Asc("+") Then
                d = 62
            ElseIf c = Asc("/") Then
                d = 63
            ElseIf c >= Asc("0") And c <= Asc("9") Then
                d = 52 + c - Asc("0")
            ElseIf c >= Asc("a") And c <= Asc("z") Then
                d = 26 + c - Asc("a")
            ElseIf c >= Asc("A") And c <= Asc("Z") Then
                d = c - Asc("A")
            ElseIf c = Asc("=") And 3 * i + j > buf_sz Then
                d = 0
            Else
                opus_picture_tag_parse_impl = OP_ENOTFORMAT
                Exit Function
            End If
            value = value * 64 Or d                 ' value=value<<6|d (24 bits)
        Next
        buf_(3 * i) = (value \ &H10000) And &HFF
        If 3 * i + 1 < buf_sz Then
            buf_(3 * i + 1) = (value \ &H100&) And &HFF
            If 3 * i + 2 < buf_sz Then buf_(3 * i + 2) = value And &HFF
        End If
    Next
    i = 0
    picture_type = op_parse_uint32be(buf_, i)
    i = i + 4
    ' Extract the MIME type.
    mime_type_length = op_parse_uint32be(buf_, i)
    i = i + 4
    If U32_(mime_type_length) > buf_sz - 32 Then opus_picture_tag_parse_impl = OP_ENOTFORMAT: Exit Function
    mime_type = op_vbBytesToStr(buf_, i, mime_type_length)
    pic_.mime_type = mime_type
    i = i + mime_type_length
    ' Extract the description string.
    description_length = op_parse_uint32be(buf_, i)
    i = i + 4
    If U32_(description_length) > buf_sz - mime_type_length - 32 Then opus_picture_tag_parse_impl = OP_ENOTFORMAT: Exit Function
    description = op_vbBytesToStr(buf_, i, description_length)
    pic_.description = description
    i = i + description_length
    ' Extract the remaining fields.
    width = op_parse_uint32be(buf_, i)
    i = i + 4
    height = op_parse_uint32be(buf_, i)
    i = i + 4
    depth = op_parse_uint32be(buf_, i)
    i = i + 4
    colors = op_parse_uint32be(buf_, i)
    i = i + 4
    ' If one of these is set, they all must be, but colors==0 is a valid value.
    colors_set = (width <> 0 Or height <> 0 Or depth <> 0 Or colors <> 0)
    If (width = 0 Or height = 0 Or depth = 0) And colors_set Then opus_picture_tag_parse_impl = OP_ENOTFORMAT: Exit Function
    data_length = op_parse_uint32be(buf_, i)
    i = i + 4
    If U32_(data_length) > buf_sz - i Then opus_picture_tag_parse_impl = OP_ENOTFORMAT: Exit Function
    ' Trim extraneous data so we don't copy it below.
    buf_sz = i + data_length
    ' Attempt to determine the image format.
    format_ = OP_PIC_FORMAT_UNKNOWN
    If mime_type_length = 3 And mime_type = "-->" Then
        format_ = OP_PIC_FORMAT_URL
        ' Picture type 1 must be a 32x32 PNG.
        If picture_type = 1 And (width <> 0 Or height <> 0) And (width <> 32 Or height <> 32) Then
            opus_picture_tag_parse_impl = OP_ENOTFORMAT
            Exit Function
        End If
        ' Append a terminating NUL for the convenience of our callers.
        buf_(buf_sz) = 0
        buf_sz = buf_sz + 1
    Else
        If mime_type_length = 10 And op_strncasecmp(mime_type, 0, "image/jpeg", 0, mime_type_length) = 0 Then
            If op_is_jpeg(buf_, i, data_length) Then format_ = OP_PIC_FORMAT_JPEG
        ElseIf mime_type_length = 9 And op_strncasecmp(mime_type, 0, "image/png", 0, mime_type_length) = 0 Then
            If op_is_png(buf_, i, data_length) Then format_ = OP_PIC_FORMAT_PNG
        ElseIf mime_type_length = 9 And op_strncasecmp(mime_type, 0, "image/gif", 0, mime_type_length) = 0 Then
            If op_is_gif(buf_, i, data_length) Then format_ = OP_PIC_FORMAT_GIF
        ElseIf mime_type_length = 0 Or (mime_type_length = 6 And op_strncasecmp(mime_type, 0, "image/", 0, mime_type_length) = 0) Then
            If op_is_jpeg(buf_, i, data_length) Then
                format_ = OP_PIC_FORMAT_JPEG
            ElseIf op_is_png(buf_, i, data_length) Then
                format_ = OP_PIC_FORMAT_PNG
            ElseIf op_is_gif(buf_, i, data_length) Then
                format_ = OP_PIC_FORMAT_GIF
            End If
        End If
        file_width = 0: file_height = 0: file_depth = 0: file_colors = 0
        has_palette = -1
        Select Case format_
        Case OP_PIC_FORMAT_JPEG
            op_extract_jpeg_params buf_, i, data_length, file_width, file_height, file_depth, file_colors, has_palette
        Case OP_PIC_FORMAT_PNG
            op_extract_png_params buf_, i, data_length, file_width, file_height, file_depth, file_colors, has_palette
        Case OP_PIC_FORMAT_GIF
            op_extract_gif_params buf_, i, data_length, file_width, file_height, file_depth, file_colors, has_palette
        End Select
        If has_palette >= 0 Then
            ' If we successfully extracted these parameters from the image, override
            '   any declared values.
            width = file_width
            height = file_height
            depth = file_depth
            colors = file_colors
        End If
        ' Picture type 1 must be a 32x32 PNG.
        If picture_type = 1 And (format_ <> OP_PIC_FORMAT_PNG Or width <> 32 Or height <> 32) Then
            opus_picture_tag_parse_impl = OP_ENOTFORMAT
            Exit Function
        End If
    End If
    ' Adjust _buf_sz instead of using data_length to capture the terminating NUL
    '   for URLs.
    buf_sz = buf_sz - i
    ' memmove(_buf,_buf+i,sizeof(*_buf)*_buf_sz); _buf=(unsigned char *)_ogg_realloc(_buf,_buf_sz)
    If buf_sz > 0 Then
        ReDim dat(0 To buf_sz - 1)
        For j = 0 To buf_sz - 1
            dat(j) = buf_(i + j)
        Next
        pic_.data = dat
    Else
        Erase pic_.data
    End If
    pic_.typ = picture_type
    pic_.width = width
    pic_.height = height
    pic_.depth = depth
    pic_.colors = colors
    pic_.data_length = data_length
    pic_.format_ = format_
    opus_picture_tag_parse_impl = 0
End Function

Public Function opus_picture_tag_parse(pic_ As OpusPictureTag, ByVal tag_ As String) As Long
    Dim pic As OpusPictureTag, buf() As Byte, base64_sz As Long, buf_sz As Long, tag_length As Long
    Dim ret As Long, tagOff As Long
    tagOff = 0
    If opus_tagncompare("METADATA_BLOCK_PICTURE", 22, tag_) = 0 Then tagOff = tagOff + 23
    ' Figure out how much BASE64-encoded data we have.
    tag_length = op_vbStrlen(tag_, tagOff)
    If tag_length And 3 Then opus_picture_tag_parse = OP_ENOTFORMAT: Exit Function
    base64_sz = tag_length \ 4
    buf_sz = 3 * base64_sz
    If buf_sz < 32 Then opus_picture_tag_parse = OP_ENOTFORMAT: Exit Function
    If op_vbChar(tag_, tagOff + tag_length - 1) = Asc("=") Then buf_sz = buf_sz - 1
    If op_vbChar(tag_, tagOff + tag_length - 2) = Asc("=") Then buf_sz = buf_sz - 1
    If buf_sz < 32 Then opus_picture_tag_parse = OP_ENOTFORMAT: Exit Function
    ' Allocate an extra byte to allow appending a terminating NUL to URL data.
    ReDim buf(0 To buf_sz)
    opus_picture_tag_init pic
    ret = opus_picture_tag_parse_impl(pic, tag_, tagOff, buf, buf_sz, base64_sz)
    If ret < 0 Then
        opus_picture_tag_clear pic
        Erase buf                               ' _ogg_free(buf)
    Else
        pic_ = pic
    End If
    opus_picture_tag_parse = ret
End Function

Public Sub opus_picture_tag_init(pic_ As OpusPictureTag)
    Dim blank As OpusPictureTag
    pic_ = blank                                ' memset(_pic,0,sizeof(*_pic))
End Sub

Public Sub opus_picture_tag_clear(pic_ As OpusPictureTag)
    pic_.description = vbNullString
    pic_.mime_type = vbNullString
    Erase pic_.data
End Sub
