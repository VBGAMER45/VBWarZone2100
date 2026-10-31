Attribute VB_Name = "theora_decapiwrapper_c"
' ref/theora (libtheora 1.2.0) lib/decapiwrapper.c: the decoder's half of the legacy theora.h API (the calls
' lib/sequence/sequence.cpp makes: theora_decode_header, theora_decode_init, theora_decode_packetin,
' theora_decode_YUVout; theora_control, theora_granule_time and theora_clear dispatch here through
' theora_apiwrapper_c).
' (VB) The wrappers are theora_apiwrapper_c's pool entries (handles); OC_DEC_DISPATCH_VTBL is the number
' OC_VB_DEC_DISPATCH_VTBL in theora_state.internal_decode, its four functions the procedures below. The
' yuv_buffer gets copies of the planes (theora_h's header).
Option Explicit

' (VB) _api: a theora_apiwrapper_c handle
Public Sub th_dec_api_clear(ByVal api As Long)
    If oc_apis(api).setup <> 0 Then th_setup_free oc_apis(api).setup
    If oc_apis(api).decode <> 0 Then th_decode_free oc_apis(api).decode
    ' memset(_api,0,sizeof(*_api)): the th_api_wrapper part
    oc_apis(api).clear = 0
    oc_apis(api).setup = 0
    oc_apis(api).decode = 0
    oc_apis(api).encode = 0
End Sub

Public Sub theora_decode_clear(td As theora_state)
    Dim blank As theora_state
    If td.i <> 0 Then theora_info_clear oc_apis(td.i).info
    td = blank
End Sub

Public Function theora_decode_control(td As theora_state, ByVal req As Long, buf As Long, ByVal buf_sz As Long) As Long
    theora_decode_control = th_decode_ctl(oc_apis(oc_apis(td.i).info.codec_setup).decode, req, buf, buf_sz)
End Function

Public Function theora_decode_granule_frame(td As theora_state, ByVal gp As Double) As Double
    theora_decode_granule_frame = th_granule_frame(oc_apis(oc_apis(td.i).info.codec_setup).decode, gp)
End Function

Public Function theora_decode_granule_time(td As theora_state, ByVal gp As Double) As Double
    theora_decode_granule_time = th_granule_time(oc_apis(oc_apis(td.i).info.codec_setup).decode, gp)
End Function

' static const oc_state_dispatch_vtable OC_DEC_DISPATCH_VTBL={theora_decode_clear, theora_decode_control,
'  theora_decode_granule_frame, theora_decode_granule_time}: (VB) OC_VB_DEC_DISPATCH_VTBL (theora_apiwrapper_c)

Private Sub th_info2theora_info(ci As theora_info, info As th_info)
    ci.version_major = info.version_major
    ci.version_minor = info.version_minor
    ci.version_subminor = info.version_subminor
    ci.width = info.frame_width
    ci.height = info.frame_height
    ci.frame_width = info.pic_width
    ci.frame_height = info.pic_height
    ci.offset_x = info.pic_x
    ci.offset_y = info.pic_y
    ci.fps_numerator = info.fps_numerator
    ci.fps_denominator = info.fps_denominator
    ci.aspect_numerator = info.aspect_numerator
    ci.aspect_denominator = info.aspect_denominator
    Select Case info.colorspace
        Case TH_CS_ITU_REC_470M: ci.colorspace = OC_CS_ITU_REC_470M
        Case TH_CS_ITU_REC_470BG: ci.colorspace = OC_CS_ITU_REC_470BG
        Case Else: ci.colorspace = OC_CS_UNSPECIFIED
    End Select
    Select Case info.pixel_fmt
        Case TH_PF_420: ci.pixelformat = OC_PF_420
        Case TH_PF_422: ci.pixelformat = OC_PF_422
        Case TH_PF_444: ci.pixelformat = OC_PF_444
        Case Else: ci.pixelformat = OC_PF_RSVD
    End Select
    ci.target_bitrate = info.target_bitrate
    ci.quality = info.quality
    ' 1<<_info->keyframe_granule_shift
    If info.keyframe_granule_shift >= 31 Then
        ci.keyframe_frequency_force = &H80000000
    Else
        ci.keyframe_frequency_force = OC_PB_POW2(info.keyframe_granule_shift)
    End If
End Sub

Public Function theora_decode_init(td As theora_state, ci As theora_info) As Long
    Dim apiinfo As Long         ' th_api_info *
    Dim api As Long             ' th_api_wrapper *
    Dim info As th_info
    oc_vb_init
    api = ci.codec_setup
    ' Allocate our own combined API wrapper/theora_info struct. We put them both in one malloc'd block so that
    ' when the API wrapper is freed, the info struct goes with it. This avoids having to figure out whether or
    ' not we need to free the info struct in either theora_info_clear() or theora_clear().
    apiinfo = oc_api_vbAlloc(True)
    ' Make our own copy of the info struct, since its lifetime should be independent of the one we were passed
    ' in.
    oc_apis(apiinfo).info = ci
    ' Convert the info struct now instead of saving the the one we decoded with theora_decode_header(), since
    ' the user might have modified values (i.e., color space, aspect ratio, etc. can be specified from a higher
    ' level). The user also might be doing something "clever" with the header packets if they are not using an
    ' Ogg encapsulation.
    oc_theora_info2th_info info, ci
    ' Don't bother to copy the setup info; th_decode_alloc() makes its own copy of the stuff it needs.
    ' (VB) a NULL api (no header decoded) passes a NULL setup, where the C would crash
    If api <> 0 Then oc_apis(apiinfo).decode = th_decode_alloc(info, oc_apis(api).setup)
    If oc_apis(apiinfo).decode = 0 Then
        oc_api_vbFree apiinfo
        theora_decode_init = OC_EINVAL
        Exit Function
    End If
    oc_apis(apiinfo).clear = OC_VB_DEC_API_CLEAR
    td.internal_encode = 0
    ' Provide entry points for ABI compatibility with old decoder shared libs.
    td.internal_decode = OC_VB_DEC_DISPATCH_VTBL
    td.granulepos = 0
    td.i = apiinfo
    oc_apis(apiinfo).info.codec_setup = apiinfo
    theora_decode_init = 0
End Function

Public Function theora_decode_header(ci As theora_info, cc As theora_comment, op As ogg_packet) As Long
    Dim api As Long             ' th_api_wrapper *
    Dim info As th_info
    Dim ret As Long
    Dim setup As Long
    api = ci.codec_setup
    ' Allocate an API wrapper struct on demand, since it will not also include a theora_info struct like the
    ' ones that are used in a theora_state struct.
    If api = 0 Then
        ci.codec_setup = oc_api_vbAlloc(False)
        api = ci.codec_setup
        oc_apis(api).clear = OC_VB_DEC_API_CLEAR
    End If
    ' Convert from the theora_info struct instead of saving our own th_info struct between calls. The user
    ' might be doing something "clever" with the header packets if they are not using an Ogg encapsulation, and
    ' we don't want to break this.
    oc_theora_info2th_info info, ci
    ' We rely on the fact that theora_comment and th_comment structures are actually identical. Take care not
    ' to change this fact unless you change the code here as well!
    setup = oc_apis(api).setup
    ret = th_decode_headerin(info, cc, setup, op)
    oc_apis(api).setup = setup
    ' We also rely on the fact that the error return code values are the same, and that the implementations of
    ' these two functions return the same set of them. Note that theora_decode_header() really can return
    ' OC_NOTFORMAT, even though it is not currently documented to do so.
    If ret < 0 Then theora_decode_header = ret: Exit Function
    th_info2theora_info ci, info
    theora_decode_header = 0
End Function

Public Function theora_decode_packetin(td As theora_state, op As ogg_packet) As Long
    Dim api As Long             ' th_api_wrapper *
    Dim gp As Double
    Dim ret As Long
    If td.i = 0 Then theora_decode_packetin = OC_FAULT: Exit Function
    If oc_apis(td.i).info.codec_setup = 0 Then theora_decode_packetin = OC_FAULT: Exit Function
    api = oc_apis(td.i).info.codec_setup
    ret = th_decode_packetin(oc_apis(api).decode, op, gp)
    If ret < 0 Then theora_decode_packetin = OC_BADPACKET: Exit Function
    td.granulepos = gp
    theora_decode_packetin = 0
End Function

' (VB) the number of elements of a plane copy (0 when it was never allocated)
Private Function oc_vbPlaneSize(b() As Byte) As Long
    On Error GoTo none
    oc_vbPlaneSize = UBound(b) + 1
    Exit Function
none:
    oc_vbPlaneSize = 0
End Function

Public Function theora_decode_YUVout(td As theora_state, yuv As yuv_buffer) As Long
    Dim api As Long             ' th_api_wrapper *
    Dim decode As Long          ' th_dec_ctx *
    Dim buf As th_ycbcr_buffer
    Dim ret As Long
    If td.i = 0 Then theora_decode_YUVout = OC_FAULT: Exit Function
    If oc_apis(td.i).info.codec_setup = 0 Then theora_decode_YUVout = OC_FAULT: Exit Function
    api = oc_apis(td.i).info.codec_setup
    decode = oc_apis(api).decode
    If decode = 0 Then theora_decode_YUVout = OC_FAULT: Exit Function
    ret = th_decode_ycbcr_out(decode, buf)
    If ret >= 0 Then
        yuv.y_width = buf.p(0).width
        yuv.y_height = buf.p(0).height
        yuv.uv_width = buf.p(1).width
        yuv.uv_height = buf.p(1).height
        ' _yuv->y_stride=buf[0].stride; _yuv->uv_stride=buf[1].stride; _yuv->y=buf[0].data; _yuv->u=buf[1].data;
        ' _yuv->v=buf[2].data: (VB) copies of the planes, their strides their widths
        yuv.y_stride = buf.p(0).width
        yuv.uv_stride = buf.p(1).width
        th_decode_vbCopyPlane decode, buf.p(0), yuv.y, oc_vbPlaneSize(yuv.y)
        th_decode_vbCopyPlane decode, buf.p(1), yuv.u, oc_vbPlaneSize(yuv.u)
        th_decode_vbCopyPlane decode, buf.p(2), yuv.v, oc_vbPlaneSize(yuv.v)
    End If
    theora_decode_YUVout = ret
End Function
