VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzPlayerBoxTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/lobbyplayerboxtabs.h: the fields of class WzPlayerBoxTabs
' (VB) data of widget class WzPlayerBoxTabs
Option Explicit

Public weakTitleUI As Long              ' std::weak_ptr<WzMultiplayerOptionsTitleUI>: its address
Private tabButtons_() As WIDGET         ' std::vector<std::shared_ptr<WzPlayerBoxTabButton>>: tabButtons_(0 .. tabButtons_n - 1)
Private tabButtons_n As Long
Public optionsButton As WIDGET          ' std::shared_ptr<WzPlayerBoxOptionsButton>
Public optionsOverlayScreen As W_SCREEN ' std::shared_ptr<W_SCREEN>
Public playerDisplayView_ As Long       ' PlayerDisplayView playerDisplayView = PlayerDisplayView::Players;

' tabButtons[i]
Public Property Get tabButtons(ByVal i As Long) As WIDGET
    Set tabButtons = tabButtons_(i)
End Property

' tabButtons.size()
Public Function tabButtons_size() As Long
    tabButtons_size = tabButtons_n
End Function

' tabButtons.push_back(w)
Public Sub tabButtons_push_back(ByVal w As WIDGET)
    If tabButtons_n = 0 Then
        ReDim tabButtons_(0 To 1)
    ElseIf tabButtons_n > UBound(tabButtons_) Then
        ReDim Preserve tabButtons_(0 To 2 * tabButtons_n - 1)
    End If
    Set tabButtons_(tabButtons_n) = w
    tabButtons_n = tabButtons_n + 1
End Sub
