VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/netqueue.h, netqueue.cpp: class NetQueue, a queue of
' NetMessages. A NetQueue can convert the messages into a stream of bytes,
' which can be sent over the network, and converted back into a queue of
' NetMessages by the NetQueue at the other end.
'
' (VB) The std::list: messages are added to the front and read from the
' back. Each message gets a serial number as it is pushed (newer ones
' larger), kept at msgs(serial - baseSerial); the list holds serials oldest ..
' nextSerial - 1. An iterator is a serial, or END_ for messages.end();
' begin() is the newest serial (END_ when empty), and --i goes one message
' newer (from end(), to the oldest). Like a std::list iterator, a serial
' stays on its message when others are pushed (a saved begin() goes stale).
' The message type numbers (GAME_GAME_TIME 120, NET_SECURED_NET_MESSAGE
' 72) are written out: classes do not use NetApi's enums.
Option Explicit

Private Const END_ As Long = -1             ' messages.end()

Private canGetMessagesForNet As Boolean     ' True if we will send the messages over the network, false if we don't.
Private canGetMessages As Boolean           ' True if we will get the messages, false if we don't use them ourselves.

Private dataPos As Long                     ' List::iterator: Last message which was sent over the network.
Private messagePos As Long                  ' List::iterator: Last message which was popped.
Private msgs() As NetMessage                ' List of messages. Messages are added to the front and read from the back.
Private msgsCapacity As Long                ' (VB) msgs()'s size
Private baseSerial As Long                        ' (VB) the serial at msgs(0)
Private oldest As Long                      ' (VB) the serial of the back (oldest) message
Private nextSerial As Long                  ' (VB) the newest serial + 1
Private incompleteReceivedMessageData() As Byte  ' Data from network which has not yet formed an entire message.
Private incompleteSize As Long              ' (VB) its size
Private pendingGameTimeUpdateMessages As Long    ' size_t: Pending GAME_GAME_TIME messages added to this queue
Private bCurrentMessageWasDecrypted As Boolean

Private Sub Class_Initialize()
    canGetMessagesForNet = True
    canGetMessages = True
    pendingGameTimeUpdateMessages = 0
    bCurrentMessageWasDecrypted = False
    dataPos = END_
    messagePos = END_
End Sub

'--------------------------------------------------------------------------
' (VB) the std::list

' messages.begin()
Private Function begin_() As Long
    If nextSerial = oldest Then begin_ = END_ Else begin_ = nextSerial - 1
End Function

' --i
Private Function prev_(ByVal i As Long) As Long
    If i = END_ Then prev_ = oldest Else prev_ = i + 1
End Function

' *i
Private Function at_(ByVal i As Long) As NetMessage
    Set at_ = msgs(i - baseSerial)
End Function

' messages.emplace_front(m)
Private Sub push_front(ByVal m As NetMessage)
    Dim i As Long, n As Long
    If nextSerial - baseSerial >= msgsCapacity Then
        n = nextSerial - oldest
        If msgsCapacity > 0 And oldest - baseSerial >= msgsCapacity \ 2 Then
            ' the erased half at the start makes room
            For i = 0 To n - 1
                Set msgs(i) = msgs(oldest - baseSerial + i)
            Next
            For i = n To msgsCapacity - 1
                Set msgs(i) = Nothing
            Next
            baseSerial = oldest
        ElseIf msgsCapacity = 0 Then
            msgsCapacity = 16
            ReDim msgs(0 To msgsCapacity - 1)
        Else
            msgsCapacity = msgsCapacity * 2
            ReDim Preserve msgs(0 To msgsCapacity - 1)
        End If
    End If
    Set msgs(nextSerial - baseSerial) = m
    nextSerial = nextSerial + 1
End Sub

' messages.erase(i, messages.end()): i and every older message
Private Sub erase_to_end(ByVal i As Long)
    Dim s As Long
    If i = END_ Then Exit Sub
    For s = oldest To i
        Set msgs(s - baseSerial) = Nothing
    Next
    oldest = i + 1
    If oldest = nextSerial Then baseSerial = oldest   ' empty: msgs() is used from the start again
End Sub

'--------------------------------------------------------------------------

' Network related, receiving
' Inserts data from the network into the NetQueue: netData(0 .. netLen - 1).
Public Sub writeRawData(netData() As Byte, ByVal netLen As Long)
    Dim used As Long, typ As Long, length As Long, msg As NetMessageBuilder, i As Long
    used = 0
    ' Short alias: incompleteReceivedMessageData.

    ' Insert the data.
    If netLen > 0 Then
        If incompleteSize = 0 Then
            ReDim incompleteReceivedMessageData(0 To netLen - 1)
        ElseIf UBound(incompleteReceivedMessageData) < incompleteSize + netLen - 1 Then
            ReDim Preserve incompleteReceivedMessageData(0 To incompleteSize + netLen - 1)
        End If
        For i = 0 To netLen - 1
            incompleteReceivedMessageData(incompleteSize + i) = netData(i)
        Next
        incompleteSize = incompleteSize + netLen
    End If

    ' Extract the messages.
    Do While incompleteSize - used >= NetMessage_HEADER_LENGTH
        typ = incompleteReceivedMessageData(used)
        ' Load payload length from uint16_t (network byte order) starting at the second byte of the data buffer.
        length = CLng(incompleteReceivedMessageData(used + 1)) * 256 + incompleteReceivedMessageData(used + 2)

        If incompleteSize - used - NetMessage_HEADER_LENGTH < length Then
            Exit Do                             ' Don't have a whole message ready yet.
        End If

        Set msg = NetMessageBuilder_new(typ, length)
        msg.append_bytes incompleteReceivedMessageData, used + NetMessage_HEADER_LENGTH, length
        push_front msg.build()

        If typ = 120 Then                       ' GAME_GAME_TIME
            pendingGameTimeUpdateMessages = pendingGameTimeUpdateMessages + 1
        End If
        used = used + NetMessage_HEADER_LENGTH + length
    Loop

    ' Recycle old data.
    If used > 0 Then
        For i = 0 To incompleteSize - used - 1
            incompleteReceivedMessageData(i) = incompleteReceivedMessageData(used + i)
        Next
        incompleteSize = incompleteSize - used
    End If
End Sub

Public Function currentIncompleteDataBuffered() As Long
    currentIncompleteDataBuffered = incompleteSize
End Function

' Network related, sending
' Marks that we will not be sending this data over the network.
Public Sub setWillNeverGetMessagesForNet()
    canGetMessagesForNet = False
End Sub

' Checks that we didn't mark that we will not be sending this data over the network (returns 0), and returns the number of messages to be sent.
Public Function numMessagesForNet() As Long
    Dim count As Long, i As Long
    count = 0
    If canGetMessagesForNet Then
        i = dataPos
        Do While i <> begin_()
            count = count + 1
            i = prev_(i)
        Loop
    End If

    numMessagesForNet = count
End Function

' Extracts data from the NetQueue to send over the network.
Public Function getMessageForNet() As NetMessage
    If Not canGetMessagesForNet Then ASSERT_FAILURE "canGetMessagesForNet", "lib/netplay/netqueue.cpp:248", "getMessageForNet", "Wrong NetQueue type for getMessageForNet."
    If Not (dataPos <> begin_()) Then ASSERT_FAILURE "dataPos != messages.begin()", "lib/netplay/netqueue.cpp:249", "getMessageForNet", "No message to get!"

    ' Return the message.
    Set getMessageForNet = internal_getMessageForNet()
End Function

' Pops the extracted data, so that future getMessageForNet calls do not return that data.
Public Sub popMessageForNet()
    If Not canGetMessagesForNet Then ASSERT_FAILURE "canGetMessagesForNet", "lib/netplay/netqueue.cpp:257", "popMessageForNet", "Wrong NetQueue type for popMessageForNet."
    If Not (dataPos <> begin_()) Then ASSERT_FAILURE "dataPos != messages.begin()", "lib/netplay/netqueue.cpp:258", "popMessageForNet", "No message to pop!"

    If messagePos <> begin_() Then
        If internal_getMessageForNet().type_() = 120 Then  ' GAME_GAME_TIME
            If pendingGameTimeUpdateMessages > 0 Then
                pendingGameTimeUpdateMessages = pendingGameTimeUpdateMessages - 1
            End If
        End If
    End If

    ' Pop the message.
    dataPos = prev_(dataPos)

    ' Recycle old data.
    popOldMessages
End Sub

' All game clients should check game messages from all queues, including their own, and only the net messages sent to them.
' Message related, storing.
' Adds a message to the queue.
Public Sub pushMessage(ByVal message As NetMessage)
    If message.type_() = 120 Then               ' GAME_GAME_TIME
        pendingGameTimeUpdateMessages = pendingGameTimeUpdateMessages + 1
    End If
    push_front message
End Sub

' Message related, extracting.
' Marks that we will not be reading any of the messages (only sending over the network).
Public Sub setWillNeverGetMessages()
    canGetMessages = False
End Sub

' Return true if we have a message ready to return.
Public Function haveMessage() As Boolean
    If Not canGetMessages Then ASSERT_FAILURE "canGetMessages", "lib/netplay/netqueue.cpp:291", "haveMessage", "Wrong NetQueue type for haveMessage."
    haveMessage = messagePos <> begin_()
End Function

' Returns a message.
Public Function getMessage() As NetMessage
    If Not canGetMessages Then ASSERT_FAILURE "canGetMessages", "lib/netplay/netqueue.cpp:297", "getMessage", "Wrong NetQueue type for getMessage."
    If Not (messagePos <> begin_()) Then ASSERT_FAILURE "messagePos != messages.begin()", "lib/netplay/netqueue.cpp:298", "getMessage", "No message to get!"

    ' Return the message.
    Set getMessage = internal_getConstMessage()
End Function

Public Function currentMessageWasDecrypted() As Boolean
    currentMessageWasDecrypted = bCurrentMessageWasDecrypted
End Function

Public Function replaceCurrentWithDecrypted(ByVal decryptedMessage As NetMessage) As Boolean
    Dim currentMessage As NetMessage
    If Not canGetMessages Then ASSERT_FAILURE "canGetMessages", "lib/netplay/netqueue.cpp:311", "replaceCurrentWithDecrypted", "Wrong NetQueue type for getMessage.": replaceCurrentWithDecrypted = False: Exit Function
    If Not (messagePos <> begin_()) Then ASSERT_FAILURE "messagePos != messages.begin()", "lib/netplay/netqueue.cpp:312", "replaceCurrentWithDecrypted", "No message to get!": replaceCurrentWithDecrypted = False: Exit Function

    Set currentMessage = internal_getMessage()
    If Not (currentMessage.type_() = 72) Then ASSERT_FAILURE "currentMessage.type() == NET_SECURED_NET_MESSAGE", "lib/netplay/netqueue.cpp:315", "replaceCurrentWithDecrypted", "Current message is not a secured message!": replaceCurrentWithDecrypted = False: Exit Function

    Set msgs(prev_(messagePos) - baseSerial) = decryptedMessage    ' currentMessage = std::move(decryptedMessage)
    bCurrentMessageWasDecrypted = True
    replaceCurrentWithDecrypted = True
End Function

' Pops the last returned message.
Public Sub popMessage()
    If Not canGetMessages Then ASSERT_FAILURE "canGetMessages", "lib/netplay/netqueue.cpp:324", "popMessage", "Wrong NetQueue type for popMessage."
    If Not (messagePos <> begin_()) Then ASSERT_FAILURE "messagePos != messages.begin()", "lib/netplay/netqueue.cpp:325", "popMessage", "No message to pop!"

    If messagePos <> begin_() Then
        If internal_getConstMessage().type_() = 120 Then  ' GAME_GAME_TIME
            If pendingGameTimeUpdateMessages > 0 Then
                pendingGameTimeUpdateMessages = pendingGameTimeUpdateMessages - 1
            End If
        End If
    End If

    ' Pop the message.
    messagePos = prev_(messagePos)
    bCurrentMessageWasDecrypted = False

    ' Recycle old data.
    popOldMessages
End Sub

Public Function numPendingGameTimeUpdateMessages() As Long
    numPendingGameTimeUpdateMessages = pendingGameTimeUpdateMessages
End Function

' Pops any messages that are no longer needed.
Private Sub popOldMessages()
    Dim i As Long
    If Not canGetMessagesForNet Then
        dataPos = begin_()
    End If
    If Not canGetMessages Then
        messagePos = begin_()
    End If

    i = END_
    Do While i <> dataPos And i <> messagePos
        i = prev_(i)
    Loop

    If i = dataPos Then
        dataPos = END_                          ' Old iterator will become invalid.
    End If
    If i = messagePos Then
        messagePos = END_                       ' Old iterator will become invalid.
    End If

    erase_to_end i
End Sub

Private Function internal_getMessageForNet() As NetMessage
    ' Return the message.
    Set internal_getMessageForNet = at_(prev_(dataPos))
End Function

Private Function internal_getConstMessage() As NetMessage
    ' Return the message.
    Set internal_getConstMessage = at_(prev_(messagePos))
End Function

Private Function internal_getMessage() As NetMessage
    ' Return the message.
    Set internal_getMessage = at_(prev_(messagePos))
End Function
