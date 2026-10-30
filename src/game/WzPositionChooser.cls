VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzPositionChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multiint.cpp: the fields of class WzPositionChooser (multiint_lobby_c)
' (VB) data of widget class WzPositionChooser
Option Explicit

Public switcherPlayerIdx As Long        ' uint32_t
Public psWeakTitleUI As Long            ' std::weak_ptr<WzMultiplayerOptionsTitleUI>: its address
Public weakSelf As Long                 ' (VB) its std::weak_ptr<WzPositionChooser>s: the widgweak_vb serial

' std::vector<PlayerButtonRow> positionSelectionRows (struct PlayerButtonRow {widget, targetPlayerIdx})
Private rows_widget() As WIDGET         ' std::shared_ptr<W_BUTTON> widget
Private rows_targetPlayerIdx() As Long  ' uint32_t targetPlayerIdx = 0
Private rows_size As Long

Public Sub positionSelectionRows_push_back(ByVal widget As WIDGET, ByVal targetPlayerIdx As Long)
    ReDim Preserve rows_widget(0 To rows_size)
    ReDim Preserve rows_targetPlayerIdx(0 To rows_size)
    Set rows_widget(rows_size) = widget
    rows_targetPlayerIdx(rows_size) = targetPlayerIdx
    rows_size = rows_size + 1
End Sub

Public Function positionSelectionRows_size() As Long
    positionSelectionRows_size = rows_size
End Function

Public Function positionSelectionRows_widget(ByVal i As Long) As WIDGET
    Set positionSelectionRows_widget = rows_widget(i)
End Function

Public Function positionSelectionRows_targetPlayerIdx(ByVal i As Long) As Long
    positionSelectionRows_targetPlayerIdx = rows_targetPlayerIdx(i)
End Function
