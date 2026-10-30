VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzTitleUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/titleui.h: class WzTitleUI, the interface of the title
' screens (WzOldTitleUI in titleui/old.cpp; the multiplayer options, game
' finder, limits, message box and the other screens come with cluster F).
' Each screen implements it; std::shared_ptr<WzTitleUI> is a reference to
' it. run returns a TITLECODE. Regarding construction vs. start(): this
' allows a reference to the parent to be held for a stack-like effect.
Option Explicit

Public Sub start()
End Sub

' NOTE! When porting, add screen_disableMapPreview(); if relevant!
Public Function run() As Long
End Function

' (WzTitleUI::screenSizeDidChange: no-op by default)
Public Sub screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
End Sub
