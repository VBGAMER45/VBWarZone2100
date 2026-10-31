VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzQuickChatPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzQuickChatPanel
' src/hci/quickchat.cpp: the fields of class WzQuickChatPanel (quickchat_c); sectionRows is a Collection of Collections (std::vector<Row>), sections a Collection (element n is item n + 1); optional<WzQuickChatSendToMode> is has_sendToMode and sendToMode; the initialisers are its _construct's.
Option Explicit

Public panelName As String                 ' WzString
Public sectionRows As Collection           ' std::vector<Row>, typedef std::vector<std::shared_ptr<WzQuickChatSection>> Row
Public sections As Collection              ' std::vector<std::shared_ptr<WzQuickChatSection>>
Public sendToWidget As WIDGET              ' std::shared_ptr<WzQuickChatSendToSelector>
Public has_sendToMode As Boolean           ' optional<WzQuickChatSendToMode> sendToMode
Public sendToMode As WzQuickChatPanel_WzQuickChatSendToMode
Public layoutMinSectionsPerRow As Long     ' size_t = 1
Public hasFocus As Boolean                 ' = false

Private Sub Class_Initialize()
    Set sectionRows = New Collection
    Set sections = New Collection
End Sub
