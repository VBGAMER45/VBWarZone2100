Attribute VB_Name = "piemode_c"
' lib/ivis_opengl/piemode.cpp: renderer control for pumpkin library
' functions; ivisdef.h's iClip and iSurface (the screen surface structure).
Option Explicit

'*************************************************************************
'
' screen surface structure
'
'*************************************************************************
Public Type iClip
    left_ As Long
    top As Long
    right_ As Long
    bottom As Long
End Type

Public Type iSurface
    xcentre As Long
    ycentre As Long
    clip As iClip

    width As Long
    height As Long
End Type

'***************************************************************************
'
'	Source
'
'***************************************************************************

Public rendSurface As iSurface

' (VB) !renderingFrame: static bool renderingFrame = true; // starts off true (a VB module has no initialisers)
Private renderingFrameEnded As Boolean

Public Sub pie_UpdateSurfaceGeometry()
    rendSurface.width = pie_GetVideoBufferWidth()
    rendSurface.height = pie_GetVideoBufferHeight()
    rendSurface.xcentre = pie_GetVideoBufferWidth() \ 2
    rendSurface.ycentre = pie_GetVideoBufferHeight() \ 2
    rendSurface.clip.left_ = 0
    rendSurface.clip.top = 0
    rendSurface.clip.right_ = pie_GetVideoBufferWidth()
    rendSurface.clip.bottom = pie_GetVideoBufferHeight()
End Sub

Public Function pie_Initialise() As Boolean
    pie_TexInit

    ' Find texture compression extension
    If get_wz_texture_compression() Then
        debug_ 70, LOG_TEXTURE, "pie_Initialise", "Texture compression: Yes"
    Else
        debug_ 74, LOG_TEXTURE, "pie_Initialise", "Texture compression: No"
    End If

    pie_UpdateSurfaceGeometry

    pie_SetDefaultStates
    debug_ 80, LOG_3D, "pie_Initialise", "xcentre %d; ycentre %d", rendSurface.xcentre, rendSurface.ycentre

    pie_InitializeInstancedRenderer

    pie_Initialise = True
End Function


Public Sub pie_ShutDown()
    pie_CleanUp
End Sub

'***************************************************************************

Public Sub pie_ScreenFrameRenderBegin()
    If Not renderingFrameEnded Then
        debug_ 101, LOG_WZ, "pie_ScreenFrameRenderBegin", "Call to pie_ScreenFrameRenderBegin when previous frame render hasn't ended yet"
        Exit Sub
    End If
    renderingFrameEnded = False
    gfx_api_context_beginRenderPass
    If screen_GetBackDrop() Then
        screen_Display
    End If
End Sub

Public Sub pie_ScreenFrameRenderEnd()
    If renderingFrameEnded Then
        debug_ 116, LOG_WZ, "pie_ScreenFrameRenderEnd", "Call to pie_ScreenFrameRenderEnd without matching pie_ScreenFrameRenderBegin"
        Exit Sub
    End If
    renderingFrameEnded = True
    screenDoDumpToDiskIfRequired
    gfx_api_context_endRenderPass
    wzPerfFrame
End Sub

'***************************************************************************
Public Function pie_GetResScalingFactor() As Long
    Dim result As Long
    result = 0
    If pie_GetVideoBufferWidth() * 4 > pie_GetVideoBufferHeight() * 5 Then
        result = pie_GetVideoBufferHeight() * 5 \ 4 \ 6
    Else
        result = pie_GetVideoBufferWidth() \ 6
    End If
    If result < 1 Then result = 1       ' std::max<UDWORD>(result, 1)
    pie_GetResScalingFactor = result
End Function
