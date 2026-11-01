VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlatformGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The game side of WzPlatform.dll's IPlatformGame: each method calls the
' game function it stands for.
Option Explicit

Implements IPlatformGame

Private Sub IPlatformGame_mainLoop()
    main_c.mainLoop
End Sub

Private Sub IPlatformGame_mainShutdown()
    main_c.mainShutdown
End Sub

Private Function IPlatformGame_realTime() As Long
    IPlatformGame_realTime = realTime
End Function

Private Sub IPlatformGame_gameScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    gameScreenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
End Sub

Private Sub IPlatformGame_saveGfxConfig()
    saveGfxConfig
End Sub

Private Function IPlatformGame_warGetWidth() As Long
    IPlatformGame_warGetWidth = war_GetWidth()
End Function

Private Sub IPlatformGame_warSetWidth(ByVal width As Long)
    war_SetWidth width
End Sub

Private Function IPlatformGame_warGetHeight() As Long
    IPlatformGame_warGetHeight = war_GetHeight()
End Function

Private Sub IPlatformGame_warSetHeight(ByVal height As Long)
    war_SetHeight height
End Sub

Private Function IPlatformGame_warGetScreen() As Long
    IPlatformGame_warGetScreen = war_GetScreen()
End Function

Private Sub IPlatformGame_warSetScreen(ByVal screen As Long)
    war_SetScreen screen
End Sub

Private Function IPlatformGame_warGetDisplayScale() As Long
    IPlatformGame_warGetDisplayScale = war_GetDisplayScale()
End Function

Private Sub IPlatformGame_warSetDisplayScale(ByVal scale_ As Long)
    war_SetDisplayScale scale_
End Sub

Private Function IPlatformGame_warGetColouredCursor() As Boolean
    IPlatformGame_warGetColouredCursor = war_GetColouredCursor()
End Function

Private Function IPlatformGame_warGetCursorScale() As Long
    IPlatformGame_warGetCursorScale = war_getCursorScale()
End Function

Private Sub IPlatformGame_warSetCursorScale(ByVal scale_ As Long)
    war_setCursorScale scale_
End Sub

Private Function IPlatformGame_warGetFullscreenModeScreen() As Long
    IPlatformGame_warGetFullscreenModeScreen = war_GetFullscreenModeScreen()
End Function

Private Sub IPlatformGame_warSetFullscreenModeScreen(ByVal screen As Long)
    war_SetFullscreenModeScreen screen
End Sub

Private Function IPlatformGame_warGetFullscreenModeWidth() As Long
    IPlatformGame_warGetFullscreenModeWidth = war_GetFullscreenModeWidth()
End Function

Private Sub IPlatformGame_warSetFullscreenModeWidth(ByVal width As Long)
    war_SetFullscreenModeWidth width
End Sub

Private Function IPlatformGame_warGetFullscreenModeHeight() As Long
    IPlatformGame_warGetFullscreenModeHeight = war_GetFullscreenModeHeight()
End Function

Private Sub IPlatformGame_warSetFullscreenModeHeight(ByVal height As Long)
    war_SetFullscreenModeHeight height
End Sub

Private Function IPlatformGame_warGetFullscreenModePixelDensity() As Single
    IPlatformGame_warGetFullscreenModePixelDensity = war_GetFullscreenModePixelDensity()
End Function

Private Sub IPlatformGame_warSetFullscreenModePixelDensity(ByVal pixelDensity As Single)
    war_SetFullscreenModePixelDensity pixelDensity
End Sub

Private Function IPlatformGame_warGetFullscreenModeRefreshRate() As Single
    IPlatformGame_warGetFullscreenModeRefreshRate = war_GetFullscreenModeRefreshRate()
End Function

Private Sub IPlatformGame_warSetFullscreenModeRefreshRate(ByVal refreshRate As Single)
    war_SetFullscreenModeRefreshRate refreshRate
End Sub

Private Function IPlatformGame_warGetMinimizeOnFocusLoss() As Long
    IPlatformGame_warGetMinimizeOnFocusLoss = war_getMinimizeOnFocusLoss()
End Function

Private Function IPlatformGame_warGetAutoAdjustDisplayScale() As Boolean
    IPlatformGame_warGetAutoAdjustDisplayScale = war_getAutoAdjustDisplayScale()
End Function

Private Sub IPlatformGame_warSetWindowMode(ByVal mode As Long)
    war_setWindowMode mode
End Sub

Private Sub IPlatformGame_gameDisplayScaleFactorDidChange(ByVal newDisplayScaleFactor As Single)
    gameDisplayScaleFactorDidChange newDisplayScaleFactor
End Sub
