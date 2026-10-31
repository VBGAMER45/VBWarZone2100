Attribute VB_Name = "theora_decinfo_c"
' ref/theora (libtheora 1.2.0) lib/decinfo.c: the three header packets (HAVE_MEMORY_CONSTRAINT, used only for
' fuzzing, is not defined).
' (VB) A th_setup_info * is a handle into theora_decint_h's oc_setups, allocated and freed here
' (oc_setup_vbAlloc, oc_setup_vbFree); 0 = NULL. th_comment is theora_comment (theora_codec_h). The C's NULL
' checks of _info, _tc, _setup and _op have no VB counterpart (ByRef arguments are never NULL).
Option Explicit

' (VB) _ogg_calloc(1,sizeof(oc_setup_info)): a free entry of oc_setups, zeroed; its handle
Private Function oc_setup_vbAlloc() As Long
    Dim h As Long
    Dim blank As th_setup_info
    Dim i As Long
    For h = 1 To oc_nsetups
        If Not oc_setups(h).vbUsed Then Exit For
    Next
    If h > oc_nsetups Then
        oc_nsetups = oc_nsetups + 1
        ReDim Preserve oc_setups(1 To oc_nsetups)
        h = oc_nsetups
    End If
    oc_setups(h) = blank
    ReDim oc_setups(h).huff_tables(0 To TH_NHUFFMAN_TABLES - 1)
    For i = 0 To 5
        oc_setups(h).qinfo.qi_ranges(i).sizes = -1
        oc_setups(h).qinfo.qi_ranges(i).base_matrices = -1
    Next
    oc_setups(h).vbUsed = True
    oc_setup_vbAlloc = h
End Function

' (VB) _ogg_free of a setup
Private Sub oc_setup_vbFree(ByVal h As Long)
    Dim blank As th_setup_info
    oc_setups(h) = blank
End Sub

' Unpacks a series of octets from a given byte array into the pack buffer. No checking is done to ensure the
' buffer contains enough data.
'  _opb: The pack buffer to read the octets from.
'  _buf: The byte array to store the unpacked bytes in. (VB) a String of characters 0-255
'  _len: The number of octets to unpack.
Private Function oc_unpack_octets(opb As oc_pack_buf, ByVal len_ As Long) As String
    Dim s As String
    Dim i As Long
    Dim val_ As Long
    s = String$(len_, 0)
    For i = 1 To len_
        val_ = oc_pack_read(opb, 8)
        Mid$(s, i, 1) = ChrW$(val_)
    Next
    oc_unpack_octets = s
End Function

' Unpacks a 32-bit integer encoded by octets in little-endian form.
Private Function oc_unpack_length(opb As oc_pack_buf) As Long
    Dim ret(0 To 3) As Long
    Dim i As Long
    For i = 0 To 3
        ret(i) = oc_pack_read(opb, 8)
    Next
    ' ret[0]|ret[1]<<8|ret[2]<<16|ret[3]<<24 (a 32-bit long: the top bit makes it negative)
    If ret(3) >= 128 Then
        oc_unpack_length = ((ret(3) - 128) * &H1000000) Or &H80000000 Or ret(2) * &H10000 Or ret(1) * &H100& Or ret(0)
    Else
        oc_unpack_length = ret(3) * &H1000000 Or ret(2) * &H10000 Or ret(1) * &H100& Or ret(0)
    End If
End Function

Private Function oc_info_unpack(opb As oc_pack_buf, info As th_info) As Long
    Dim val_ As Long
    ' Check the codec bitstream version.
    val_ = oc_pack_read(opb, 8)
    info.version_major = val_
    val_ = oc_pack_read(opb, 8)
    info.version_minor = val_
    val_ = oc_pack_read(opb, 8)
    info.version_subminor = val_
    ' verify we can parse this bitstream version. We accept earlier minors and all subminors, by spec
    If info.version_major > TH_VERSION_MAJOR Or _
        (info.version_major = TH_VERSION_MAJOR And info.version_minor > TH_VERSION_MINOR) Then
        oc_info_unpack = TH_EVERSION
        Exit Function
    End If
    ' Read the encoded frame description.
    val_ = oc_pack_read(opb, 16)
    info.frame_width = val_ * 16
    val_ = oc_pack_read(opb, 16)
    info.frame_height = val_ * 16
    val_ = oc_pack_read(opb, 24)
    info.pic_width = val_
    val_ = oc_pack_read(opb, 24)
    info.pic_height = val_
    val_ = oc_pack_read(opb, 8)
    info.pic_x = val_
    val_ = oc_pack_read(opb, 8)
    info.pic_y = val_
    val_ = oc_pack_read(opb, 32)
    info.fps_numerator = val_
    val_ = oc_pack_read(opb, 32)
    info.fps_denominator = val_
    If info.frame_width = 0 Or info.frame_height = 0 Or info.pic_width + info.pic_x > info.frame_width Or _
        info.pic_height + info.pic_y > info.frame_height Or info.fps_numerator = 0 Or info.fps_denominator = 0 Then
        oc_info_unpack = TH_EBADHEADER
        Exit Function
    End If
    ' Note: The sense of pic_y is inverted in what we pass back to the application compared to how it is
    ' stored in the bitstream. This is because the bitstream uses a right-handed coordinate system, while
    ' applications expect a left-handed one.
    info.pic_y = info.frame_height - info.pic_height - info.pic_y
    val_ = oc_pack_read(opb, 24)
    info.aspect_numerator = val_
    val_ = oc_pack_read(opb, 24)
    info.aspect_denominator = val_
    val_ = oc_pack_read(opb, 8)
    info.colorspace = val_
    val_ = oc_pack_read(opb, 24)
    info.target_bitrate = val_
    val_ = oc_pack_read(opb, 6)
    info.quality = val_
    val_ = oc_pack_read(opb, 5)
    info.keyframe_granule_shift = val_
    val_ = oc_pack_read(opb, 2)
    info.pixel_fmt = val_
    If info.pixel_fmt = TH_PF_RSVD Then oc_info_unpack = TH_EBADHEADER: Exit Function
    val_ = oc_pack_read(opb, 3)
    If val_ <> 0 Or oc_pack_bytes_left(opb) < 0 Then oc_info_unpack = TH_EBADHEADER: Exit Function
    oc_info_unpack = 0
End Function

Private Function oc_comment_unpack(opb As oc_pack_buf, tc As theora_comment) As Long
    Dim len_ As Long
    Dim i As Long
    ' Read the vendor string.
    len_ = oc_unpack_length(opb)
    If len_ < 0 Or len_ > oc_pack_bytes_left(opb) Then oc_comment_unpack = TH_EBADHEADER: Exit Function
    ' _tc->vendor=_ogg_malloc((size_t)len+1); oc_unpack_octets(_opb,_tc->vendor,len); _tc->vendor[len]='\0';
    tc.vendor = oc_unpack_octets(opb, len_)
    ' Read the user comments.
    tc.comments = oc_unpack_length(opb)
    len_ = tc.comments
    If len_ < 0 Or len_ > (&H7FFFFFFF \ 4) Then
        tc.comments = 0
        oc_comment_unpack = TH_EBADHEADER
        Exit Function
    End If
    If len_ * 4 > oc_pack_bytes_left(opb) Then
        tc.comments = 0
        oc_comment_unpack = TH_EBADHEADER
        Exit Function
    End If
    If 0 < tc.comments Then
        ReDim tc.comment_lengths(0 To tc.comments - 1)
        ReDim tc.user_comments(0 To tc.comments - 1)
        For i = 0 To tc.comments - 1
            len_ = oc_unpack_length(opb)
            If len_ < 0 Or len_ > oc_pack_bytes_left(opb) Then
                tc.comments = i
                oc_comment_unpack = TH_EBADHEADER
                Exit Function
            End If
            tc.comment_lengths(i) = len_
            tc.user_comments(i) = oc_unpack_octets(opb, len_)
        Next
    Else
        Erase tc.comment_lengths
        Erase tc.user_comments
    End If
    If oc_pack_bytes_left(opb) < 0 Then oc_comment_unpack = TH_EBADHEADER Else oc_comment_unpack = 0
End Function

Private Function oc_setup_unpack(opb As oc_pack_buf, setup As th_setup_info) As Long
    Dim ret As Long
    ' Read the quantizer tables.
    ret = oc_quant_params_unpack(opb, setup.qinfo)
    If ret < 0 Then oc_setup_unpack = ret: Exit Function
    ' Read the Huffman trees.
    oc_setup_unpack = oc_huff_trees_unpack(opb, setup.huff_tables, setup.vb_huff, setup.vb_nhuff)
End Function

Private Sub oc_setup_clear(setup As th_setup_info)
    oc_quant_params_clear setup.qinfo
    oc_huff_trees_clear setup.huff_tables, setup.vb_huff, setup.vb_nhuff
End Sub

' (VB) _setup: th_setup_info ** (a handle, ByRef)
Private Function oc_dec_headerin(opb As oc_pack_buf, info As th_info, tc As theora_comment, setup As Long, _
    op As ogg_packet) As Long
    Dim buffer As String
    Dim val_ As Long
    Dim packtype As Long
    Dim ret As Long
    Dim s As Long
    val_ = oc_pack_read(opb, 8)
    packtype = val_
    ' If we're at a data packet...
    If (packtype And &H80) = 0 Then
        ' Check to make sure we received all three headers... If we haven't seen any valid headers, assume
        ' this is not actually Theora.
        If info.frame_width <= 0 Then oc_dec_headerin = TH_ENOTFORMAT: Exit Function
        ' Follow our documentation, which says we'll return TH_EFAULT if this are NULL (_info was checked by
        ' our caller). (VB) _tc is never NULL
        ' And if any other headers were missing, declare this packet "out of sequence" instead.
        If StrPtr(tc.vendor) = 0 Then oc_dec_headerin = TH_EBADHEADER: Exit Function
        ' Don't check this until it's needed, since we allow passing NULL for the arguments that we're not
        ' expecting the next header to fill in yet. (VB) _setup is never NULL
        If setup = 0 Then oc_dec_headerin = TH_EBADHEADER: Exit Function
        ' If we got everything, we're done.
        oc_dec_headerin = 0
        Exit Function
    End If
    ' Check the codec string.
    buffer = oc_unpack_octets(opb, 6)
    If buffer <> "theora" Then oc_dec_headerin = TH_ENOTFORMAT: Exit Function
    Select Case packtype
        ' Codec info header.
        Case &H80
            ' This should be the first packet, and we should not already be initialized.
            If op.b_o_s = 0 Or info.frame_width > 0 Then oc_dec_headerin = TH_EBADHEADER: Exit Function
            ret = oc_info_unpack(opb, info)
            If ret < 0 Then th_info_clear info Else ret = 3
        ' Comment header.
        Case &H81
            ' We should have already decoded the info header, and should not yet have decoded the comment
            ' header.
            If info.frame_width = 0 Or StrPtr(tc.vendor) <> 0 Then oc_dec_headerin = TH_EBADHEADER: Exit Function
            ret = oc_comment_unpack(opb, tc)
            If ret < 0 Then th_comment_clear tc Else ret = 2
        ' Codec setup header.
        Case &H82
            ' We should have already decoded the info header and the comment header, and should not yet have
            ' decoded the setup header.
            If info.frame_width = 0 Or StrPtr(tc.vendor) = 0 Or setup <> 0 Then
                oc_dec_headerin = TH_EBADHEADER
                Exit Function
            End If
            s = oc_setup_vbAlloc()
            ret = oc_setup_unpack(opb, oc_setups(s))
            If ret < 0 Then
                oc_setup_clear oc_setups(s)
                oc_setup_vbFree s
            Else
                setup = s
                ret = 1
            End If
        Case Else
            ' We don't know what this header is.
            oc_dec_headerin = TH_EBADHEADER
            Exit Function
    End Select
    oc_dec_headerin = ret
End Function

' Decodes one header packet. This should be called repeatedly with the packets at the beginning of the
' stream until it returns 0.
' (VB) _setup: th_setup_info ** (a handle, ByRef)
Public Function th_decode_headerin(info As th_info, tc As theora_comment, setup As Long, op As ogg_packet) As Long
    Dim opb As oc_pack_buf
    ' if(_op==NULL)return TH_EBADHEADER; if(_info==NULL)return TH_EFAULT: (VB) never NULL
    oc_vb_init
    oc_pack_readinit opb, op.packet, 0, op.bytes
    th_decode_headerin = oc_dec_headerin(opb, info, tc, setup, op)
End Function

Public Sub th_setup_free(ByVal setup As Long)
    If setup <> 0 Then
        oc_setup_clear oc_setups(setup)
        oc_setup_vbFree setup
    End If
End Sub
