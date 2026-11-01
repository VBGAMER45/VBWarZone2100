VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideLinkedTopicsWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideLinkedTopicsWidget
' src/screens/guidescreen.cpp: WzGuideLinkedTopicsWidget's fields (guidescreen_c); the initialisers are WzGuideLinkedTopicsWidget_construct's
' and Class_Initialize's.
Option Explicit

Public m_seeAlsoLabel As WIDGET              ' std::shared_ptr<W_LABEL>
Public m_linkedTopics As WIDGET              ' std::shared_ptr<Paragraph>
Public m_scrollableLinkedTopicsContainer As WIDGET ' std::shared_ptr<ScrollableListWidget>
Public m_onClickLinkedTopicHandler As WzFunc ' OnClickLinkedTopicFunc
Public weakSerial As Long                    ' (VB) its widgweak_vb serial (std::weak_ptr<WzGuideLinkedTopicsWidget>)
