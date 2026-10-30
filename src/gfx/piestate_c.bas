Attribute VB_Name = "piestate_c"
' lib/ivis_opengl/piestate.cpp: renderer setup and state control routines for
' 3D rendering; piestate.h's RENDER_STATE and pal_PIELIGHTtoVec4, and
' pielight_convert.h's pielightToRGBAVec4. A PIELIGHT is a Long: its rgba()
' (r in the low byte, as the union's byte order); PIELIGHT_r ... PIELIGHT_a
' read its bytes.
Option Explicit

Public Type RENDER_STATE
    fogEnabled As Boolean
    fog As Boolean
    fogColour As Long                   ' PIELIGHT
    fogBegin As Single
    fogEnd As Single
    texPage As Long
    rendMode As REND_MODE
End Type

' gfx_api::buffer* pie_internal::rectBuffer = nullptr;
Public pie_internal_rectBuffer As gfx_api_buffer
Private rendStates As RENDER_STATE
Private timeState As Single              ' gfx_api::gfxFloat = 0.0f

'--------------------------------------------------------------------------
' (VB) a PIELIGHT's bytes

Public Function PIELIGHT_r(ByVal c As Long) As Long
    PIELIGHT_r = c And &HFF&
End Function

Public Function PIELIGHT_g(ByVal c As Long) As Long
    PIELIGHT_g = (c And &HFF00&) \ &H100&
End Function

Public Function PIELIGHT_b(ByVal c As Long) As Long
    PIELIGHT_b = (c And &HFF0000) \ &H10000
End Function

Public Function PIELIGHT_a(ByVal c As Long) As Long
    PIELIGHT_a = ((c And &H7F000000) \ &H1000000) Or ((c < 0) And &H80&)
End Function

' (VB) a PIELIGHT with alpha replaced (colour.byte.a = a)
Public Function PIELIGHT_setA(ByVal c As Long, ByVal a As Long) As Long
    PIELIGHT_setA = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' piestate.h: (1 / 255.0f) * glm::vec4{r, g, b, a}
Public Function pal_PIELIGHTtoVec4(ByVal rgba As Long) As Vector4f
    pal_PIELIGHTtoVec4.x = (1! / 255!) * PIELIGHT_r(rgba)
    pal_PIELIGHTtoVec4.y = (1! / 255!) * PIELIGHT_g(rgba)
    pal_PIELIGHTtoVec4.z = (1! / 255!) * PIELIGHT_b(rgba)
    pal_PIELIGHTtoVec4.w = (1! / 255!) * PIELIGHT_a(rgba)
End Function

' pielight_convert.h
Public Function pielightToRGBAVec4(ByVal color As Long) As Vector4f
    pielightToRGBAVec4.x = PIELIGHT_r(color) / 255!
    pielightToRGBAVec4.y = PIELIGHT_g(color) / 255!
    pielightToRGBAVec4.z = PIELIGHT_b(color) / 255!
    pielightToRGBAVec4.w = PIELIGHT_a(color) / 255!
End Function

'--------------------------------------------------------------------------

' const Vector3f defaultSunPosition(225.0f, -600.0f, 450.0f);
Private Function defaultSunPosition() As Vector3f
    defaultSunPosition.x = 225!
    defaultSunPosition.y = -600!
    defaultSunPosition.z = 450!
End Function

Public Sub rendStatesRendModeHack()
    rendStates.rendMode = REND_ALPHA
End Sub

'
'	Source
'

Public Sub pie_SetDefaultStates() 'Sets all states
    Dim black As Long
    black = PIELIGHT_setA(0, 255)

    'fog off
    rendStates.fogEnabled = False ' enable fog before renderer
    rendStates.fog = False 'to force reset to false
    pie_SetFogStatus False
    pie_SetFogColour black

    rendStates.rendMode = REND_ALPHA    ' to force reset to REND_OPAQUE
End Sub

'***************************************************************************
'
' pie_EnableFog(bool val)
'
' Global enable/disable fog to allow fog to be turned of ingame
'
'***************************************************************************
Public Sub pie_EnableFog(ByVal val As Boolean)
    If rendStates.fogEnabled <> val Then
        If val Then
            debug_ 87, LOG_FOG, "pie_EnableFog", "pie_EnableFog: Setting fog to %s", "ON"
        Else
            debug_ 87, LOG_FOG, "pie_EnableFog", "pie_EnableFog: Setting fog to %s", "OFF"
        End If
        rendStates.fogEnabled = val
    End If
End Sub

Public Function pie_GetFogEnabled() As Boolean
    pie_GetFogEnabled = rendStates.fogEnabled
End Function

Public Function pie_GetFogStatus() As Boolean
    pie_GetFogStatus = rendStates.fog
End Function

Public Sub pie_SetFogColour(ByVal colour As Long)
    rendStates.fogColour = colour
End Sub

Public Function pie_GetFogColour() As Long
    pie_GetFogColour = rendStates.fogColour
End Function

Public Sub pie_FreeShaders()
    ' no-op, currently
End Sub

'static float fogBegin;
'static float fogEnd;

' Run from screen.c on init.
Public Function pie_LoadShaders(ByVal shadowFilterSize As Long, ByVal pointLightEnabled As Boolean) As Boolean
    Dim shadowConstants As lighting_constants, supported As Boolean
    Dim rect(0 To 15) As Byte
    ' note: actual loading of shaders now occurs in gfx_api

    ' initialize gfx context shadow constants (must happen after context is initialized)
    If Not (gfx_api_context_isInitialized()) Then ASSERT_FAILURE "gfx_api::context::isInitialized()", "lib/ivis_opengl/piestate.cpp:126", "pie_LoadShaders", "gfx context isn't initialized?"
    shadowConstants = gfx_api_context_getShadowConstants()
    shadowConstants.shadowFilterSize = shadowFilterSize
    shadowConstants.isPointLightPerPixelEnabled = pointLightEnabled
    gfx_api_context_setShadowConstants shadowConstants

    ' pie_supportsShadowMapping().value_or(false)
    If Not pie_supportsShadowMapping(supported) Then supported = False
    If Not supported Then
        pie_setShadowMode ShadowMode_Fallback_Stencil_Shadows
    Else
        pie_setShadowMode ShadowMode_Shadow_Mapping
    End If

    ' gfx_api::gfxUByte rect[] { 0, 255, 0, 255,  0, 0, 0, 255,  255, 255, 0, 255,  255, 0, 0, 255 };
    rect(0) = 0: rect(1) = 255: rect(2) = 0: rect(3) = 255
    rect(4) = 0: rect(5) = 0: rect(6) = 0: rect(7) = 255
    rect(8) = 255: rect(9) = 255: rect(10) = 0: rect(11) = 255
    rect(12) = 255: rect(13) = 0: rect(14) = 0: rect(15) = 255
    If pie_internal_rectBuffer Is Nothing Then Set pie_internal_rectBuffer = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_static_draw, "rectBuffer")
    pie_internal_rectBuffer.upload 16, VarPtr(rect(0))

    pie_LoadShaders = True
End Function

Public Sub pie_SetShaderTime(ByVal shaderTime As Long)
    Dim base As Long
    base = UMOD(shaderTime, 1000)
    If base > 500 Then
        base = 1000 - base  ' cycle
    End If
    timeState = CSng(base) / 1000!
End Sub

Public Function pie_GetShaderTime() As Single
    pie_GetShaderTime = timeState
End Function

'/ Set the OpenGL fog start and end
Public Sub pie_UpdateFogDistance(ByVal begin As Single, ByVal end_ As Single)
    rendStates.fogBegin = begin
    rendStates.fogEnd = end_
End Sub

Public Sub pie_SetFogStatus(ByVal val As Boolean)
    If rendStates.fogEnabled Then
        'fog enabled so toggle if required
        rendStates.fog = val
    Else
        rendStates.fog = False
    End If
End Sub

Public Function getCurrentRenderState() As RENDER_STATE
    getCurrentRenderState = rendStates
End Function

' (VB) getCurrentRenderState() for the exe: the fields of the RENDER_STATE
Public Sub getCurrentRenderState_fields(fogEnabled As Boolean, fog As Boolean, fogColour As Long, fogBegin As Single, fogEnd As Single, texPage As Long, rendMode As REND_MODE)
    fogEnabled = rendStates.fogEnabled
    fog = rendStates.fog
    fogColour = rendStates.fogColour
    fogBegin = rendStates.fogBegin
    fogEnd = rendStates.fogEnd
    texPage = rendStates.texPage
    rendMode = rendStates.rendMode
End Sub

Public Function pie_GetMaxAntialiasing() As Long
    Dim maxSamples As Long
    maxSamples = gfx_api_context_get_context_value(context_value_MAX_SAMPLES)
    pie_GetMaxAntialiasing = maxSamples
End Function

Public Function getDefaultSunPosition() As Vector3f
    getDefaultSunPosition = defaultSunPosition()
End Function

' (VB) uint32_t % for a value held in a Long
Private Function UMOD(ByVal x As Long, ByVal m As Long) As Long
    If x >= 0 Then
        UMOD = x Mod m
    Else
        UMOD = CLng((CDbl(x) + 4294967296#) - Fix((CDbl(x) + 4294967296#) / m) * m)
    End If
End Function
