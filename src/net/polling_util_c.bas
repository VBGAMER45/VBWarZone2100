Attribute VB_Name = "polling_util_c"
' lib/netplay/polling_util.cpp, polling_util.h. A const std::vector<IClientConnection*>& is a Collection of
' IClientConnection objects.
Option Explicit

' void resetDescriptorSet(const std::vector<IClientConnection*>& conns, IDescriptorSet& dset)
Private Sub resetDescriptorSet(ByVal conns As Collection, ByVal dset As SelectDescriptorSet)
    Dim conn As IClientConnection
    dset.clear
    For Each conn In conns
        If Not (dset.add(conn)) Then ASSERT_FAILURE "dset.add(conn)", "lib/netplay/polling_util.cpp:36", "resetDescriptorSet", "Failed to add connection to descriptor set"
    Next
End Sub

' net::result<int> checkConnectionsReadable(const std::vector<IClientConnection*>& conns, IDescriptorSet& readableSet,
'     std::chrono::milliseconds timeout)
Public Function checkConnectionsReadable(ByVal conns As Collection, ByVal readableSet As SelectDescriptorSet, _
        ByVal timeoutMs As Long, ec As Long) As Long
    Dim conn As IClientConnection, compressedReady As Boolean, pollRes As Long, msg As String, errState As Long
    ec = 0
    If conns.Count = 0 Then
        checkConnectionsReadable = 0
        Exit Function
    End If

    compressedReady = False
    For Each conn In conns
        If Not (conn.isValid()) Then ASSERT_FAILURE "conn->isValid()", "lib/netplay/polling_util.cpp:53", "checkConnectionsReadable", "Invalid connection!"
        If conn.isCompressed() Then
            If Not conn.compressionAdapter.decompressionNeedInput() Then
                compressedReady = True
                Exit For
            End If
        End If
    Next

    If compressedReady Then
        ' A socket already has some data ready. Don't really poll the sockets.
        For Each conn In conns
            If conn.isCompressed() Then
                conn.setReadReady Not conn.compressionAdapter.decompressionNeedInput()
            Else
                conn.setReadReady False
            End If
        Next
        checkConnectionsReadable = conns.Count
        Exit Function
    End If

    resetDescriptorSet conns, readableSet
    pollRes = readableSet.poll(timeoutMs, ec)
    If ec <> 0 Then
        msg = net_errorMessage(ec)
        debug_ 77, LOG_ERROR, "checkConnectionsReadable", "poll failed: %s", msg
        Exit Function
    End If

    If pollRes = 0 Then
        debug_ 83, LOG_WARNING, "checkConnectionsReadable", "poll timed out after waiting for %u milliseconds", timeoutMs
        checkConnectionsReadable = 0
        Exit Function
    End If

    For Each conn In conns
        conn.setReadReady readableSet.isSet(conn, errState)    ' isSet(conn).value_or(false): an error is false
    Next
    checkConnectionsReadable = pollRes
End Function
