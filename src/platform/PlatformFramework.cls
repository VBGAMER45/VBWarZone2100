VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlatformFramework"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' WzPlatform's answers to the window functions lib/framework calls
' (WzFramework's IFrameworkPlatform); registered by wzMain.
Option Explicit

Implements IFrameworkPlatform

Private Function IFrameworkPlatform_hasBackend() As Boolean
    IFrameworkPlatform_hasBackend = main_win.WZbackendHasValue
End Function

Private Function IFrameworkPlatform_windowHandle() As Long
    IFrameworkPlatform_windowHandle = main_win.WZwindow
End Function

Private Function IFrameworkPlatform_wzIsFullscreen() As Boolean
    IFrameworkPlatform_wzIsFullscreen = main_win.wzIsFullscreen()
End Function

Private Function IFrameworkPlatform_wzChangeWindowMode(ByVal mode As WINDOW_MODE, ByVal silent As Boolean) As Boolean
    IFrameworkPlatform_wzChangeWindowMode = main_win.wzChangeWindowMode(mode, silent)
End Function

Private Sub IFrameworkPlatform_inputInitialise()
    main_win.inputInitialise
End Sub
