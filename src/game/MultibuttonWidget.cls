VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultibuttonWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class MultibuttonWidget
' lib/widget/multibutform.h: MultibuttonWidget's fields (multibutform_c). std::vector<std::pair<
' std::shared_ptr<W_BUTTON>, int>> buttons is the Collections buttons and buttonValues.
Option Explicit

Public label As WIDGET
Public buttons As Collection
Public buttonValues As Collection
Public currentValue_ As Long
Public disabled As Boolean
Public gap_ As Long
Public outerPaddingLeft_ As Long
Public outerPaddingRight_ As Long
Public lockCurrent As Boolean
Public butAlign As MultibuttonWidget_ButtonAlignment ' = ButtonAlignment::RIGHT_ALIGN
Public onChooseHandlers As Collection   ' std::vector<W_ON_CHOOSE_FUNC>
Public canChooseHandler As WzFunc       ' W_CAN_CHOOSE_FUNC

Private Sub Class_Initialize()
    Set buttons = New Collection
    Set buttonValues = New Collection
    Set onChooseHandlers = New Collection
End Sub
