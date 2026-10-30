VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPlatformGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' What lib/sdl calls in the game (src, lib/gamelib): the exe implements this
' interface (PlatformGame.cls) and hands it to PlatformSetGame at startup.
' Each method stands for the C++ function of the same name without its
' underscores (VB's Implements refuses them): warGetWidth is war_GetWidth.
' More arrive with the clusters that port the rest of main_sdl.cpp.
Option Explicit

' src/main.cpp: mainLoop(), once per frame
Public Sub mainLoop()
End Sub

' src/main.cpp: mainShutdown(), the onShutdown of wzMainEventLoop
Public Sub mainShutdown()
End Sub

' lib/gamelib/gtime.h: realTime
Public Function realTime() As Long
End Function

' src/game.cpp: gameScreenSizeDidChange
Public Sub gameScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
End Sub

' src/configuration.cpp
Public Sub saveGfxConfig()
End Sub

' src/warzoneconfig.cpp
Public Function warGetWidth() As Long
End Function

Public Sub warSetWidth(ByVal width As Long)
End Sub

Public Function warGetHeight() As Long
End Function

Public Sub warSetHeight(ByVal height As Long)
End Sub

Public Function warGetScreen() As Long
End Function

Public Sub warSetScreen(ByVal screen As Long)
End Sub

Public Function warGetDisplayScale() As Long
End Function

Public Sub warSetDisplayScale(ByVal scale_ As Long)
End Sub
