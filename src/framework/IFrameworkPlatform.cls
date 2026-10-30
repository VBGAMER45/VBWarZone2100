VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFrameworkPlatform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' What lib/framework calls in lib/sdl (WzPlatform): WzPlatform implements this
' interface and hands it to frameSetPlatform. Method names keep the C++ ones,
' without underscores (VB's Implements refuses them).
Option Explicit

' main_sdl.cpp: WZbackend.has_value() (false in headless mode)
Public Function hasBackend() As Boolean
End Function

' main_sdl.cpp: WZwindow, as a window handle (0 = nullptr)
Public Function windowHandle() As Long
End Function

Public Function wzIsFullscreen() As Boolean
End Function

Public Function wzChangeWindowMode(ByVal mode As WINDOW_MODE, ByVal silent As Boolean) As Boolean
End Function

Public Sub inputInitialise()
End Sub
