VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideSidebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideSidebar
' src/screens/guidescreen.cpp: WzGuideSidebar's fields (guidescreen_c); the initialisers are WzGuideSidebar_construct's
' and Class_Initialize's.
Option Explicit

Public m_parentGuideForm As Long             ' std::weak_ptr<WzGuideForm>: the form's widgweak_vb serial
Public m_topicButtons As Collection          ' std::vector<std::shared_ptr<WzGuideSidebarTopicButton>> (1-based)
Public m_topicsList As WIDGET                ' std::shared_ptr<ScrollableListWidget>
Public m_topicIdentifierToListIdMap As Collection ' std::unordered_map<std::string, size_t>: "k" & id -> list index
Public m_currentSelectedTopicIdentifier As String ' std::string
Public weakSerial As Long                    ' (VB) its widgweak_vb serial (std::weak_ptr<WzGuideSidebar>)

Private Sub Class_Initialize()
    Set m_topicButtons = New Collection
    Set m_topicIdentifierToListIdMap = New Collection
End Sub
