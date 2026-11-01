VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGameGuideScreen_CLICKFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGameGuideScreen_CLICKFORM
' src/screens/guidescreen.cpp: WzGameGuideScreen_CLICKFORM's fields (guidescreen_c); the initialisers are WzGameGuideScreen_CLICKFORM_construct's
' and Class_Initialize's.
Option Explicit

Public backgroundColor As Long               ' PIELIGHT = pal_RGBA(0, 0, 0, 80)
Public onClickedFunc As WzFunc               ' std::function<void ()>
Public onCancelPressed As WzFunc             ' std::function<void ()>
Public guideForm As WIDGET                   ' std::shared_ptr<WzGuideForm>
Public edgeMargin As Long                    ' const int = 10
Public weakSerial As Long                    ' (VB) its widgweak_vb serial (std::weak_ptr<WzGameGuideScreen_CLICKFORM>)

Private Sub Class_Initialize()
    edgeMargin = 10
End Sub
