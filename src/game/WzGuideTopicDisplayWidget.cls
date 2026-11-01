VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideTopicDisplayWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideTopicDisplayWidget
' src/screens/guidescreen.cpp: WzGuideTopicDisplayWidget's fields (guidescreen_c); the initialisers are WzGuideTopicDisplayWidget_construct's
' and Class_Initialize's.
Option Explicit

Public m_topicID As String                   ' std::string
Public m_parentNavPath As WIDGET             ' std::shared_ptr<W_LABEL>
Public m_displayName As WIDGET               ' std::shared_ptr<W_LABEL>
Public m_visuals As Collection               ' std::vector<std::shared_ptr<WIDGET>>
Public m_contents As WIDGET                  ' std::shared_ptr<Paragraph>
Public m_scrollableContentsContainer As WIDGET ' std::shared_ptr<ScrollableListWidget>
Public m_linkedTopicsWidget As WIDGET        ' std::shared_ptr<WzGuideLinkedTopicsWidget>
Public maxLinkedTopicsHeight As Long         ' int = 28
Public weakSerial As Long                    ' (VB) its widgweak_vb serial (std::weak_ptr<WzGuideTopicDisplayWidget>)

Private Sub Class_Initialize()
    Set m_visuals = New Collection
    maxLinkedTopicsHeight = 28
End Sub
