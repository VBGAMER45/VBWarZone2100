VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideTopicsNavBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGuideTopicsNavBar
' src/screens/guidescreen.cpp: WzGuideTopicsNavBar's fields (guidescreen_c); the initialisers are WzGuideTopicsNavBar_construct's
' and Class_Initialize's.
Option Explicit

Public m_titleLabel As WIDGET                ' std::shared_ptr<W_LABEL>
Public m_backButton As WIDGET                ' std::shared_ptr<WzGuideTopicsNavButton>
Public m_nextButton As WIDGET                ' std::shared_ptr<WzGuideTopicsNavButton>
Public m_sidebarButton As WIDGET             ' std::shared_ptr<W_BUTTON>
Public m_closeButton As WIDGET               ' std::shared_ptr<W_BUTTON>
Public m_optionsButton As WIDGET             ' std::shared_ptr<W_BUTTON>
Public optionsOverlayScreen As W_SCREEN      ' std::shared_ptr<W_SCREEN>
Public m_stateHistory As Collection          ' std::vector<std::string> (1-based)
Public m_currentStateIndex As Long           ' size_t = 0
Public m_popStateHandler As WzFunc           ' OnPopStateFunction
Public m_onSidebarToggleHandler As WzFunc    ' OnSidebarToggleFunc
Public m_onCloseHandler As WzFunc            ' OnCloseFunc
Public m_topicDisplayWidth As Long           ' int32_t = 0
Public weakSerial As Long                    ' (VB) its widgweak_vb serial (std::weak_ptr<WzGuideTopicsNavBar>)

Private Sub Class_Initialize()
    Set m_stateHistory = New Collection
End Sub
