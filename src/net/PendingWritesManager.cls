VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PendingWritesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/pending_writes_manager.cpp, pending_writes_manager.h: class PendingWritesManager, every write of one
' connection provider's connections queued and sent when the sockets take it.
' No threads in the port (cluster J's Contract C): the "WzPendingWrites" thread's loop body is threadPass, one pass
' over the queued connections with a zero select timeout (4.7.0's thread waits up to 50 ms for a writable socket);
' append runs it, as the semaphore post wakes the thread, and PendingWritesManager_pollAll runs it for every
' manager (netplay calls it each frame). The mutex and the semaphore are nothing.
' This file is not one that includes tcp/sock_error.h: EBADF and ECONNRESET are MSVC's errno values.
' Private to WzNet (netplay and the connection classes use it; the exe has no use for it).
Option Explicit

Private Const EBADF As Long = 9
Private Const ECONNRESET As Long = 108

Private pendingWrites_ As Collection            ' ConnectionThreadWriteMap: ConnectionWriteQueue items keyed by ObjPtr
Private writableSet_ As SelectDescriptorSet
Private initialized_ As Boolean                 ' thread_ != nullptr
Private stopRequested_ As Boolean

Private Sub Class_Initialize()
    Set pendingWrites_ = New Collection
End Sub

' ~PendingWritesManager()
Private Sub Class_Terminate()
    deinitialize
End Sub

Private Function keyOf(ByVal conn As IClientConnection) As String
    keyOf = CStr(ObjPtr(conn))
End Function

' (VB) pendingWrites_.find(conn) (Nothing: end())
Private Function findQueue(ByVal conn As IClientConnection) As ConnectionWriteQueue
    On Error Resume Next
    Set findQueue = pendingWrites_.Item(keyOf(conn))
End Function

' void deinitialize()
Friend Sub deinitialize()
    If Not initialized_ Then
        ' No-op in case of a repeated `deinitialize()` call
        Exit Sub
    End If
    stopRequested_ = True
    Set pendingWrites_ = New Collection          ' pendingWrites_.clear()
    Set writableSet_ = Nothing
    ' (the thread is woken and joined; the mutex and semaphore destroyed)
    initialized_ = False
End Sub

' void initialize(WzConnectionProvider& connProvider)
Friend Sub initialize(ByVal connProvider As WzConnectionProvider)
    If initialized_ Then
        ' No-op in case of a repeated `initialize()` call
        Exit Sub
    End If
    Set writableSet_ = connProvider.newDescriptorSet(1)      ' PollEventType::WRITABLE
    stopRequested_ = False
    ' (the mutex, the semaphore and the "WzPendingWrites" thread: threadPass)
    initialized_ = True
End Sub

' template <typename AppendToWriteQueueFn> void append(IClientConnection* conn, AppendToWriteQueueFn&& appendFn): the
' appendFn of both callers inserts bytes at the queue's end, so it takes them
Friend Sub append(ByVal conn As IClientConnection, buf() As Byte, ByVal ofs As Long, ByVal n As Long)
    Dim writeQueue As ConnectionWriteQueue
    ' executeUnderLock: if (pendingWrites_.empty()) wzSemaphorePost(sema_) wakes the thread
    Set writeQueue = findQueue(conn)
    If writeQueue Is Nothing Then
        Set writeQueue = New ConnectionWriteQueue
        writeQueue.construct conn
        pendingWrites_.Add writeQueue, keyOf(conn)
    End If
    writeQueue.append buf, ofs, n
    ' (VB) the woken thread's pass
    If initialized_ And Not stopRequested_ Then threadPass
End Sub

' void clearPendingWrites(IClientConnection* conn)
Friend Sub clearPendingWrites(ByVal conn As IClientConnection)
    If Not findQueue(conn) Is Nothing Then pendingWrites_.remove keyOf(conn)
End Sub

' net::result<int> checkConnectionsWritable(IDescriptorSet& writableSet, std::chrono::milliseconds timeout)
Private Function checkConnectionsWritable(ByVal writableSet As SelectDescriptorSet, ByVal timeoutMs As Long, ec As Long) As Long
    Dim pollRes As Long, msg As String
    ec = 0
    If writableSet.empty_() Then
        checkConnectionsWritable = 0
        Exit Function
    End If

    pollRes = writableSet.poll(timeoutMs, ec)

    If ec <> 0 Then
        msg = net_errorMessage(ec)
        debug_ 87, LOG_ERROR, "checkConnectionsWritable", "poll failed: %s", msg
        Exit Function
    End If

    If pollRes = 0 Then
        debug_ 93, LOG_WARNING, "checkConnectionsWritable", "poll timed out after waiting for %u milliseconds", timeoutMs
        checkConnectionsWritable = 0
        Exit Function
    End If

    checkConnectionsWritable = pollRes
End Function

' void populateWritableSet(IDescriptorSet& writableSet)
Private Sub populateWritableSet(ByVal writableSet As SelectDescriptorSet)
    Dim items() As ConnectionWriteQueue, n As Long, i As Long, conn As IClientConnection
    n = snapshot(items)
    For i = 0 To n - 1
        If Not items(i).empty_() Then
            writableSet.add items(i).conn
        Else
            ASSERT_FAILURE "false", "lib/netplay/pending_writes_manager.cpp:112", "populateWritableSet", "Empty buffer for pending socket writes"   ' This shouldn't happen!
            Set conn = items(i).conn
            pendingWrites_.remove keyOf(conn)
            If conn.deleteLaterRequested() Then
                conn.destroy_
            End If
        End If
    Next
End Sub

' (VB) the map's entries as they are now, for loops that erase
Private Function snapshot(items() As ConnectionWriteQueue) As Long
    Dim q As ConnectionWriteQueue, n As Long
    If pendingWrites_.Count = 0 Then Exit Function
    ReDim items(0 To pendingWrites_.Count - 1)
    For Each q In pendingWrites_
        Set items(n) = q
        n = n + 1
    Next
    snapshot = n
End Function

' void threadImplFunction(): one pass of its `while (!stopRequested_)` loop
Friend Sub threadPass()
    Const WRITABLE_CHECK_TIMEOUT As Long = 0     ' (VB) 4.7.0's thread waits 50 ms; the game's thread must not
    Dim checkWritableRes As Long, ec As Long, items() As ConnectionWriteQueue, n As Long, i As Long
    Dim conn As IClientConnection, writeQueue As ConnectionWriteQueue, isSetResult As Boolean, errState As Long
    Dim retSent As Long, ecSent As Long, connStatus As Long, errMsg As String

    If stopRequested_ Or Not initialized_ Then Exit Sub

    ' Check if we can write to some connections.
    writableSet_.clear
    populateWritableSet writableSet_
    If Not (Not writableSet_.empty_() Or pendingWrites_.Count = 0) Then ASSERT_FAILURE "!writableSet_->empty() || pendingWrites_.empty()", "lib/netplay/pending_writes_manager.cpp:132", "threadImplFunction", "writableSet must not be empty if there are pending writes."

    checkWritableRes = checkConnectionsWritable(writableSet_, WRITABLE_CHECK_TIMEOUT, ec)

    If ec = 0 And checkWritableRes <> 0 Then
        n = snapshot(items)
        For i = 0 To n - 1
            Set writeQueue = items(i)
            Set conn = writeQueue.conn
            If Not (Not writeQueue.empty_()) Then ASSERT_FAILURE "!writeQueue.empty()", "lib/netplay/pending_writes_manager.cpp:145", "threadImplFunction", "writeQueue[sock] must not be empty."

            isSetResult = writableSet_.isSet(conn, errState)
            If errState <> 0 Then
                ' Poll returned some fatal error state for the connection
                Select Case errState
                Case 1                           ' IDescriptorSet::ErroredState::InvalidConn
                    ' Connection is somehow invalid?
                    debug_ 155, LOG_ERROR, "threadImplFunction", "Socket error: PollInvalidConn?"
                    conn.setWriteErrorCode True, make_network_error_code(EBADF)
                Case 2                           ' IDescriptorSet::ErroredState::Error
                    debug_ 159, LOG_INFO, "threadImplFunction", "Socket error: PollError"
                    conn.setWriteErrorCode True, make_network_error_code(EBADF)
                Case 3                           ' IDescriptorSet::ErroredState::HangUp
                    debug_ 163, LOG_NET, "threadImplFunction", "Socket error: PollHangUp"
                    conn.setWriteErrorCode True, make_network_error_code(ECONNRESET)
                End Select

                pendingWrites_.remove keyOf(conn)    ' Connection broken, don't try writing to it again.
                If conn.deleteLaterRequested() Then
                    conn.destroy_
                End If
                GoTo continue_
            End If

            If Not isSetResult Or writeQueue.empty_() Then
                GoTo continue_                   ' This connection is not ready for writing, or we don't have anything to write.
            End If

            ' Write data.
            retSent = writeQueue.sendTo(conn, ecSent)
            If ecSent = 0 Then
                ' Erase as much data as written.
                writeQueue.eraseFront retSent
                If writeQueue.empty_() Then
                    pendingWrites_.remove keyOf(conn)   ' Nothing left to write, delete from pending list.
                    If conn.deleteLaterRequested() Then
                        conn.destroy_
                    End If
                End If
            Else
                If net_errorEqualsErrc(ecSent, errc_interrupted) Then
                    ' Not an actual error, just try to send again later
                    GoTo continue_
                End If
                connStatus = conn.connectionStatus()
                If Not conn.isValid() Or connStatus <> 0 Then    ' Check if the connection is still open
                    If connStatus <> 0 Then
                        errMsg = net_errorMessage(connStatus)
                        debug_ 209, LOG_NET, "threadImplFunction", "Socket error: %s", errMsg
                        conn.setWriteErrorCode True, connStatus
                    Else
                        debug_ 214, LOG_NET, "threadImplFunction", "Socket error: connection is not valid"
                        conn.setWriteErrorCode True, make_network_error_code(ECONNRESET)
                    End If
                    pendingWrites_.remove keyOf(conn)   ' Connection broken, don't try writing to it again.
                    If conn.deleteLaterRequested() Then
                        conn.destroy_
                    End If
                End If
            End If
continue_:
        Next
    End If
    ' if (pendingWrites_.empty()) the thread waits on the semaphore: the next append or pollAll
End Sub

' void safeDispose(IClientConnection* conn)
Friend Sub safeDispose(ByVal conn As IClientConnection)
    Dim connProvider As WzConnectionProvider
    ' executeUnderLock
    If Not findQueue(conn) Is Nothing Then
        ' Wait until the data is written, then delete the socket.
        conn.requestDeleteLater
    Else
        ' Notify the owning connection provider to properly dispose of the connection object.
        Set connProvider = conn.connProviderLock()
        If Not connProvider Is Nothing Then
            connProvider.disposeConnection conn
        Else
            ASSERT_FAILURE "false", "lib/netplay/pending_writes_manager.cpp:256", "safeDispose", "IClientConnection::connProvider_ has gone away before safeDispose!"
        End If
        ' Delete the socket and destroy the connection right away.
        conn.destroy_
    End If
End Sub
