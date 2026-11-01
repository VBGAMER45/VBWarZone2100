VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideForm
' src/screens/guidescreen.cpp: WzGuideForm's fields (guidescreen_c); the initialisers are WzGuideForm_construct's
' and Class_Initialize's.
Option Explicit

Public m_navBar As WIDGET                    ' std::shared_ptr<WzGuideTopicsNavBar>
Public m_sideBar As WIDGET                   ' std::shared_ptr<WzGuideSidebar>
Public m_guideTopic As WIDGET                ' std::shared_ptr<WzGuideTopicDisplayWidget>
Public m_sidebarVisible As Boolean           ' = false
Public m_queuedTopicsListRefresh As Boolean  ' = false
Public m_closeHandler As WzFunc              ' OnCloseFunc
Public weakSerial As Long                    ' (VB) its widgweak_vb serial (std::weak_ptr<WzGuideForm>)
