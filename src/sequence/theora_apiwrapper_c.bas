Attribute VB_Name = "theora_apiwrapper_c"
' ref/theora (libtheora 1.2.0) lib/apiwrapper.c with apiwrapper.h: the legacy theora.h API over the th_ one
' (the parts shared by the encoder and the decoder; the decoder's are in theora_decapiwrapper_c).
' (VB) A th_api_wrapper * is a handle into oc_apis (0 = NULL), allocated by theora_decapiwrapper_c and freed
' here; a th_api_info (a wrapper followed by a theora_info, one malloc'd block in the C) is an entry whose
' info is used (vbHasInfo), and the C's &apiinfo->api and &apiinfo->info are both its handle. The function
' pointers (th_api_wrapper.clear, the oc_state_dispatch_vtable of theora_state.internal_decode) are numbers:
' clear 1 is th_dec_api_clear, internal_decode 1 is decapiwrapper.c's OC_DEC_DISPATCH_VTBL. The encoder's
' half (internal_encode, th_enc_ctx) is not ported: internal_encode stays 0.
Option Explicit

' Provide an entry point for the codec setup to clear itself in case we ever want to break pieces off into a
' common base library shared by encoder and decoder. In addition, this makes several other pieces of the API
' wrapper cleaner. (oc_setup_clear_func: (VB) a number, 1 = th_dec_api_clear)
Public Const OC_VB_DEC_API_CLEAR As Long = 1
' (VB) theora_state.internal_decode = &OC_DEC_DISPATCH_VTBL
Public Const OC_VB_DEC_DISPATCH_VTBL As Long = 1

' Generally only one of these pointers will be non-NULL in any given instance. Technically we do not even
' really need this struct, since we should be able to figure out which one from "context", but doing it this
' way makes sure we don't flub it up.
' (VB) th_api_wrapper and th_api_info in one type: the info is used when vbHasInfo
Public Type th_api_info
    ' th_api_wrapper api:
    clear As Long               ' oc_setup_clear_func: (VB) OC_VB_DEC_API_CLEAR or 0
    setup As Long               ' th_setup_info *: (VB) a theora_decinfo_c handle
    decode As Long              ' th_dec_ctx *: (VB) a theora_decode_c handle
    encode As Long              ' th_enc_ctx *: (VB) always 0
    info As theora_info         ' (th_api_info only)
    vbHasInfo As Boolean        ' (VB) allocated as a th_api_info
    vbUsed As Boolean           ' (VB) the pool entry is allocated
End Type

Public oc_apis() As th_api_info
Public oc_napis As Long

' (VB) _ogg_calloc of a th_api_wrapper (hasInfo False) or a th_api_info (True): its handle
Public Function oc_api_vbAlloc(ByVal hasInfo As Boolean) As Long
    Dim h As Long
    Dim blank As th_api_info
    For h = 1 To oc_napis
        If Not oc_apis(h).vbUsed Then Exit For
    Next
    If h > oc_napis Then
        oc_napis = oc_napis + 1
        ReDim Preserve oc_apis(1 To oc_napis)
        h = oc_napis
    End If
    oc_apis(h) = blank
    oc_apis(h).vbHasInfo = hasInfo
    oc_apis(h).vbUsed = True
    oc_api_vbAlloc = h
End Function

' (VB) _ogg_free of a wrapper
Public Sub oc_api_vbFree(ByVal h As Long)
    Dim blank As th_api_info
    oc_apis(h) = blank
End Sub

Public Function theora_version_string() As String
    theora_version_string = th_version_string()
End Function

Public Function theora_version_number() As Long
    theora_version_number = th_version_number()
End Function

Public Sub theora_info_init(ci As theora_info)
    Dim blank As theora_info
    ci = blank
End Sub

Public Sub theora_info_clear(ci As theora_info)
    Dim api As Long
    Dim blank As theora_info
    api = ci.codec_setup
    ci = blank
    If api <> 0 Then
        If oc_apis(api).clear <> 0 Then th_dec_api_clear api     ' (*api->clear)(api)
        oc_api_vbFree api
    End If
End Sub

Public Sub theora_clear(th As theora_state)
    Dim blank As theora_state
    ' Provide compatibility with mixed encoder and decoder shared lib versions.
    If th.internal_decode <> 0 Then theora_decode_clear th          ' (*vtable->clear)(_th)
    ' if(_th->internal_encode!=NULL): (VB) no encoder
    If th.i <> 0 Then theora_info_clear oc_apis(th.i).info
    th = blank
End Sub

Public Function theora_control(th As theora_state, ByVal req As Long, buf As Long, ByVal buf_sz As Long) As Long
    ' Provide compatibility with mixed encoder and decoder shared lib versions.
    If th.internal_decode <> 0 Then
        theora_control = theora_decode_control(th, req, buf, buf_sz)
    Else
        ' else if(_th->internal_encode!=NULL): (VB) no encoder
        theora_control = TH_EINVAL
    End If
End Function

Public Function theora_granule_frame(th As theora_state, ByVal gp As Double) As Double
    ' Provide compatibility with mixed encoder and decoder shared lib versions.
    If th.internal_decode <> 0 Then
        theora_granule_frame = theora_decode_granule_frame(th, gp)
    Else
        theora_granule_frame = -1
    End If
End Function

Public Function theora_granule_time(th As theora_state, ByVal gp As Double) As Double
    ' Provide compatibility with mixed encoder and decoder shared lib versions.
    If th.internal_decode <> 0 Then
        theora_granule_time = theora_decode_granule_time(th, gp)
    Else
        theora_granule_time = -1
    End If
End Function

Public Sub oc_theora_info2th_info(info As th_info, ci As theora_info)
    info.version_major = ci.version_major
    info.version_minor = ci.version_minor
    info.version_subminor = ci.version_subminor
    info.frame_width = ci.width
    info.frame_height = ci.height
    info.pic_width = ci.frame_width
    info.pic_height = ci.frame_height
    info.pic_x = ci.offset_x
    info.pic_y = ci.offset_y
    info.fps_numerator = ci.fps_numerator
    info.fps_denominator = ci.fps_denominator
    info.aspect_numerator = ci.aspect_numerator
    info.aspect_denominator = ci.aspect_denominator
    Select Case ci.colorspace
        Case OC_CS_ITU_REC_470M: info.colorspace = TH_CS_ITU_REC_470M
        Case OC_CS_ITU_REC_470BG: info.colorspace = TH_CS_ITU_REC_470BG
        Case Else: info.colorspace = TH_CS_UNSPECIFIED
    End Select
    Select Case ci.pixelformat
        Case OC_PF_420: info.pixel_fmt = TH_PF_420
        Case OC_PF_422: info.pixel_fmt = TH_PF_422
        Case OC_PF_444: info.pixel_fmt = TH_PF_444
        Case Else: info.pixel_fmt = TH_PF_RSVD
    End Select
    info.target_bitrate = ci.target_bitrate
    info.quality = ci.quality
    ' _ci->keyframe_frequency_force>0?OC_MINI(31,oc_ilog(_ci->keyframe_frequency_force-1)):0 (ogg_uint32_t)
    If ci.keyframe_frequency_force <> 0 Then
        info.keyframe_granule_shift = OC_MINI(31, oc_ilog(ci.keyframe_frequency_force - 1))
    Else
        info.keyframe_granule_shift = 0
    End If
End Sub

Public Function theora_packet_isheader(op As ogg_packet) As Long
    theora_packet_isheader = th_packet_isheader(op)
End Function

Public Function theora_packet_iskeyframe(op As ogg_packet) As Long
    theora_packet_iskeyframe = th_packet_iskeyframe(op)
End Function

Public Function theora_granule_shift(ci As theora_info) As Long
    ' This breaks when keyframe_frequency_force is not positive or is larger than 2**31 (if your int is more
    ' than 32 bits), but that's what the original function does.
    theora_granule_shift = oc_ilog(ci.keyframe_frequency_force - 1)
End Function

Public Sub theora_comment_init(tc As theora_comment)
    th_comment_init tc
End Sub

Public Function theora_comment_query(tc As theora_comment, ByVal tag As String, ByVal count As Long) As String
    theora_comment_query = th_comment_query(tc, tag, count)
End Function

Public Function theora_comment_query_count(tc As theora_comment, ByVal tag As String) As Long
    theora_comment_query_count = th_comment_query_count(tc, tag)
End Function

Public Sub theora_comment_clear(tc As theora_comment)
    th_comment_clear tc
End Sub

Public Sub theora_comment_add(tc As theora_comment, ByVal comment As String)
    th_comment_add tc, comment
End Sub

Public Sub theora_comment_add_tag(tc As theora_comment, ByVal tag As String, ByVal value As String)
    th_comment_add_tag tc, tag, value
End Sub
