VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzQuickChatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzQuickChatForm
' src/hci/quickchat.cpp: the fields of class WzQuickChatForm (quickchat_c); panels is a Collection (element n is item n + 1); weakSerial: (VB) its std::weak_ptrs; the initialisers are its _construct's.
Option Explicit

Public weakSerial As Long
Public onQuickChatSent As WzFunc           ' std::function<void ()>
Public quickChatLabel As WIDGET            ' std::shared_ptr<W_LABEL>
Public panelSwitcher As WIDGET             ' std::shared_ptr<WzQuickChatPanelTabs>
Public panels As Collection                ' std::vector<std::shared_ptr<WzQuickChatPanel>>
Public currentlyAttachedPanel As WIDGET    ' std::shared_ptr<WzQuickChatPanel>
Public showHighlighted As Boolean          ' = false

Private Sub Class_Initialize()
    Set panels = New Collection
End Sub
