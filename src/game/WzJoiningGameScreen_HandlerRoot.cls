VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzJoiningGameScreen_HandlerRoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzJoiningGameScreen_HandlerRoot
' src/screens/joiningscreen.cpp: the fields of class WzJoiningGameScreen_HandlerRoot (joiningscreen_c); the UDT
' fields are Private with Friend accessors, the connection list is connection(0 .. connectionListCount - 1).
Option Explicit

Public onSuccessFunc As WzFunc          ' OnSuccessFunc: void (const JoinConnectionDescription&) (host, port, type)
Public onFailureFunc As WzFunc          ' OnFailureFunc: void ()
Public onClickedFunc As WzFunc          ' std::function<void ()>
Public onCancelPressed As WzFunc        ' std::function<void ()>
Public backgroundColor As Long          ' PIELIGHT = pal_RGBA(0, 0, 0, 80)

' Input
Public playerName As String             ' char playerName[64]
Private playerIdentity_ As EcKey        ' EcKey playerIdentity;
Public asSpectator As Boolean           ' = false
Private expectedHostProps_ As ExpectedHostProperties
Public numJoinRedirects As Long         ' size_t = 0

' Internal display
Public joiningProgressForm As WIDGET    ' std::shared_ptr<WzJoiningStatusForm>

' Internal state
Public currentJoiningState As Long      ' JoiningState = JoiningState::AwaitingConnection
Private connectionList_() As InternalConnectionInfo ' std::vector<InternalConnectionInfo> connectionList;
Public connectionListCount As Long
Public currentConnectionIdx As Long     ' size_t = 0
Public timeStarted As Double            ' std::chrono::steady_clock::time_point (widgSteadyNowMs)
Public gamePassword As String           ' char gamePassword[password_string_size] = {}

' Connection state
Public client_transient_socket As IClientConnection ' IClientConnection* = nullptr
Public tmp_joining_socket_set As IConnectionPollGroup ' IConnectionPollGroup* = nullptr
Public tmpJoiningQUEUE As Long          ' NETQUEUE tmpJoiningQUEUE = {};
Public has_tmpJoiningQueuePair As Boolean ' NetQueuePair* tmpJoiningQueuePair = nullptr (WzNet's: QUEUE_TRANSIENT_JOIN)
Private initialAckBuffer_(0 To 3) As Byte ' char initialAckBuffer[expectedInitialAckSize] = {}
Public usedInitialAckBuffer As Long     ' size_t = 0
Public startTime As Long                ' uint32_t = 0 (wzGetTicks)
Private hostIdentity_ As EcKey          ' EcKey hostIdentity;
Private challengeForHost_() As Byte     ' std::vector<uint8_t> challengeForHost;
Private connectionAuthSessionKeys_ As SessionKeys ' std::unique_ptr<SessionKeys> connectionAuthSessionKeys;

Public weakSerial As Long               ' (VB) its std::weak_ptr<WzJoiningGameScreen_HandlerRoot>s: the widgweak_vb serial

Friend Property Get playerIdentity() As EcKey
    playerIdentity = playerIdentity_
End Property

Friend Property Let playerIdentity(v As EcKey)
    playerIdentity_ = v
End Property

Friend Property Get expectedHostProps() As ExpectedHostProperties
    expectedHostProps = expectedHostProps_
End Property

Friend Property Let expectedHostProps(v As ExpectedHostProperties)
    expectedHostProps_ = v
End Property

Friend Property Get hostIdentity() As EcKey
    hostIdentity = hostIdentity_
End Property

Friend Property Let hostIdentity(v As EcKey)
    hostIdentity_ = v
End Property

Friend Property Get connectionAuthSessionKeys() As SessionKeys
    connectionAuthSessionKeys = connectionAuthSessionKeys_
End Property

Friend Property Let connectionAuthSessionKeys(v As SessionKeys)
    connectionAuthSessionKeys_ = v
End Property

Friend Property Get connection(ByVal i As Long) As InternalConnectionInfo
    connection = connectionList_(i)
End Property

Friend Sub setConnection(ByVal i As Long, v As InternalConnectionInfo)
    connectionList_(i) = v
End Sub

' connectionList = list(0 .. n - 1)
Friend Sub setConnectionList(list() As InternalConnectionInfo, ByVal n As Long)
    Dim i As Long
    connectionListCount = n
    If n > 0 Then
        ReDim connectionList_(0 To n - 1)
        For i = 0 To n - 1
            connectionList_(i) = list(i)
        Next
    Else
        Erase connectionList_
    End If
End Sub

Friend Function getInitialAckByte(ByVal i As Long) As Long
    getInitialAckByte = initialAckBuffer_(i)
End Function

Friend Sub setInitialAckByte(ByVal i As Long, ByVal v As Long)
    initialAckBuffer_(i) = v
End Sub

Friend Sub getChallengeForHost(out() As Byte)
    out = challengeForHost_
End Sub

Friend Sub setChallengeForHost(v() As Byte)
    challengeForHost_ = v
End Sub
