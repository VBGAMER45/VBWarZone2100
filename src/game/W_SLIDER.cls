VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_SLIDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_SLIDER
' lib/widget/slider.h: W_SLIDER's fields (slider_c). optional<WidgetHelp> help is Nothing when it has no
' value; std::vector<SliderOnChangeFunc> onChangeFuncs is a Collection of WzFuncs.
Option Explicit

Public orientation As WSLD_ORIENTATION  ' The orientation of the slider
Public numStops As Long                 ' UWORD: Number of stop positions on the slider
Public barSize As Long                  ' UWORD: Thickness of slider bar
Public pos As Long                      ' UWORD: Current stop position of the slider
Public state As Long                    ' UWORD: Slider state
Public pTip As String                   ' std::string: Tool tip
Public help As WidgetHelp               ' optional<WidgetHelp>
Public isHandlingDrag As Boolean        ' = false
Public onChangeFuncs As Collection

Private Sub Class_Initialize()
    Set onChangeFuncs = New Collection
End Sub
