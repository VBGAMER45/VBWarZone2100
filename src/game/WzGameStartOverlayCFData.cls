VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGameStartOverlayCFData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzGameStartOverlayScreen_CLICKFORM
' src/screens/netpregamescreen.cpp: the fields of class WzGameStartOverlayScreen_CLICKFORM (netpregamescreen_c); the initialisers are its _construct's.
Option Explicit

Public backgroundColor As Long          ' PIELIGHT = pal_RGBA(0, 0, 0, 200)
Public onClickedFunc As WzFunc          ' std::function<void ()>
Public onCancelPressed As WzFunc        ' std::function<void ()>
Public topLabel As WIDGET               ' std::shared_ptr<W_LABEL>
Public countdownLabel As WIDGET         ' std::shared_ptr<WzCountdownLabel>
Public handleConsoleMessage As WzFunc   ' std::shared_ptr<CONSOLE_MESSAGE_LISTENER>
Public teamStrategyView As WIDGET       ' std::shared_ptr<WIDGET>
Public strategySideTextLabel As WIDGET  ' std::shared_ptr<W_LABEL>
Public playersLoadingStatus As WIDGET   ' std::shared_ptr<WzLoadingPlayersStatusForm>
Public has_currentCountdownStart As Boolean ' optional<std::chrono::steady_clock::time_point> currentCountdownStart: has_value()
Public currentCountdownStart As Currency    ' its value: a QueryPerformanceCounter count
Public bHasTeamPlanningStage As Boolean ' = false
Public currentState As Long             ' PregameWaitingState = PregameWaitingState::Uninitialized
Public currentConsoleMessages As Collection ' std::deque<CachedConsoleMessage>: CachedConsoleMessage objects

Private Sub Class_Initialize()
    backgroundColor = pal_RGBA(0, 0, 0, 200)
    Set currentConsoleMessages = New Collection
End Sub
