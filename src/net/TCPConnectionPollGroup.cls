VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCPConnectionPollGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/tcp/tcp_connection_poll_group.cpp, tcp_connection_poll_group.h: class tcp::TCPConnectionPollGroup, the
' TCP part of an IConnectionPollGroup. std::vector<IClientConnection*> conns_ is a Collection.
Option Explicit

Private conns_ As Collection
Private connProviderSet_ As Boolean             ' WzConnectionProvider* connProvider_ (only its descriptor set is used)
' Pre-allocated descriptor set for `checkConnectionsReadable` operation to avoid extra memory allocations.
Private readableSet_ As SelectDescriptorSet

' explicit TCPConnectionPollGroup(WzConnectionProvider& connProvider): makes its base, the IConnectionPollGroup
Friend Function construct(ByVal connProvider As WzConnectionProvider) As IConnectionPollGroup
    Dim base As IConnectionPollGroup
    Set conns_ = New Collection
    connProviderSet_ = True
    Set readableSet_ = connProvider.newDescriptorSet(0)       ' PollEventType::READABLE
    Set base = New IConnectionPollGroup
    base.construct Me
    Set construct = base
End Function

' virtual net::result<int> checkConnectionsReadable(std::chrono::milliseconds timeout) override
Friend Function checkConnectionsReadable(ByVal timeoutMs As Long, ec As Long) As Long
    checkConnectionsReadable = polling_util_c.checkConnectionsReadable(conns_, readableSet_, timeoutMs, ec)
End Function

' virtual void add(IClientConnection* conn) override
Friend Sub add(ByVal conn As IClientConnection)
    Dim tcpConn As TCPClientConnection
    Set tcpConn = conn.tcpImpl                  ' dynamic_cast<TCPClientConnection*>(conn)
    If Not (Not tcpConn Is Nothing) Then ASSERT_FAILURE "tcpConn != nullptr", "lib/netplay/tcp/tcp_connection_poll_group.cpp:46", "add", "Expected to have TCPClientConnection instance": Exit Sub
    conns_.Add conn
    If Not (readableSet_.add(conn)) Then ASSERT_FAILURE "readableSet_->add(conn)", "lib/netplay/tcp/tcp_connection_poll_group.cpp:49", "add", "Failed to add connection to internal descriptor set"
End Sub

' virtual void remove(IClientConnection* conn) override
Friend Sub remove(ByVal conn As IClientConnection)
    Dim tcpConn As TCPClientConnection, i As Long
    Set tcpConn = conn.tcpImpl
    If Not (Not tcpConn Is Nothing) Then ASSERT_FAILURE "tcpConn != nullptr", "lib/netplay/tcp/tcp_connection_poll_group.cpp:55", "remove", "Expected to have TCPClientConnection instance": Exit Sub
    ' std::find + erase: the first one
    For i = 1 To conns_.Count
        If conns_.Item(i) Is conn Then
            conns_.remove i
            Exit For
        End If
    Next
    If Not (readableSet_.remove(conn)) Then ASSERT_FAILURE "readableSet_->remove(conn)", "lib/netplay/tcp/tcp_connection_poll_group.cpp:61", "remove", "Failed to remove connection from internal descriptor set"
End Sub
