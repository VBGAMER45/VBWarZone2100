VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectToForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ConnectToForm
' src/titleui/widgets/gamebrowserform.cpp: the fields of class ConnectToForm (gamebrowserform_c); its consts are the
' module's.
Option Explicit

Public onSubmitFunc As WzFunc           ' OnSubmitFunc: void (WzString inputStr, bool asSpectator)
Public has_frontendImgID As Boolean     ' optional<UWORD> frontendImgID = nullopt;
Public frontendImgID As Long
Public joinTextStr As String            ' WzString joinTextStr;

Public inputBoxBackground As Long       ' PIELIGHT
Public inputBoxBackgroundHighlight As Long ' PIELIGHT

Public inputBox As WIDGET               ' std::shared_ptr<W_EDITBOX>
Public wzText As WzText                 ' WzText wzText;

Public optionsPopover As WIDGET         ' std::shared_ptr<PopoverWidget>

Public has_lastFrameMouseIsOverRowOrChildren As Boolean ' optional<UDWORD> lastFrameMouseIsOverRowOrChildren = nullopt;
Public lastFrameMouseIsOverRowOrChildren As Long

Public weakSerial As Long               ' (VB) its std::weak_ptr<ConnectToForm>s: the widgweak_vb serial

Private Sub Class_Initialize()
    Set wzText = New WzText
End Sub
