Attribute VB_Name = "chat_c"
' src/chat.cpp, chat.h: struct InGameChatMessage, whole. Its members are
' InGameChatMessage_<member> taking the message (this) first.
' std::set<uint32_t> toPlayers is toPlayers(0 .. toPlayers_size - 1), kept
' in ascending (unsigned) order without duplicates, as the set iterates;
' getReceivers' std::vector<uint32_t> is an array and its count. The
' std::chrono::steady_clock::time_point of playerSpamMutedUntil is a
' QueryPerformanceCounter count.
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' chat.h
Public Type InGameChatMessage
    sender As Long                      ' uint32_t
    text As String                      ' char const *
    toAllies As Boolean
    toPlayers() As Long                 ' std::set<uint32_t> (private)
    toPlayers_size As Long
End Type

Public Function InGameChatMessage_new(ByVal messageSender As Long, ByVal messageText As String) As InGameChatMessage
    InGameChatMessage_new.sender = messageSender
    InGameChatMessage_new.text = messageText
End Function

' (VB) toPlayers.insert(playerIndex)
Private Sub InGameChatMessage_toPlayers_insert(m As InGameChatMessage, ByVal playerIndex As Long)
    Dim i As Long, k As Long
    k = 0
    Do While k < m.toPlayers_size
        If m.toPlayers(k) = playerIndex Then Exit Sub
        If ULT(playerIndex, m.toPlayers(k)) Then Exit Do
        k = k + 1
    Loop
    ReDim Preserve m.toPlayers(0 To m.toPlayers_size)
    For i = m.toPlayers_size To k + 1 Step -1
        m.toPlayers(i) = m.toPlayers(i - 1)
    Next
    m.toPlayers(k) = playerIndex
    m.toPlayers_size = m.toPlayers_size + 1
End Sub

' (VB) toPlayers.find(playerIndex) != toPlayers.end()
Private Function InGameChatMessage_toPlayers_contains(m As InGameChatMessage, ByVal playerIndex As Long) As Boolean
    Dim i As Long
    For i = 0 To m.toPlayers_size - 1
        If m.toPlayers(i) = playerIndex Then
            InGameChatMessage_toPlayers_contains = True
            Exit Function
        End If
    Next
End Function

Private Function InGameChatMessage_isGlobal(m As InGameChatMessage) As Boolean
    InGameChatMessage_isGlobal = Not m.toAllies And m.toPlayers_size = 0
End Function

Private Function InGameChatMessage_shouldReceive(m As InGameChatMessage, ByVal playerIndex As Long) As Boolean
    ' (&& and ||: NetPlay.players[playerIndex] only for an index within it)
    If (playerIndex >= game.maxPlayers) Then
        If (playerIndex >= NetPlay_players_size()) Then
            InGameChatMessage_shouldReceive = False
            Exit Function
        ElseIf (Not NetPlay_players_allocated(playerIndex)) Then
            InGameChatMessage_shouldReceive = False
            Exit Function
        End If
    End If
    ' (||, &&: aiCheckAlliances only for two players)
    If InGameChatMessage_isGlobal(m) Or InGameChatMessage_toPlayers_contains(m, playerIndex) Then
        InGameChatMessage_shouldReceive = True
    ElseIf m.toAllies And ULT(m.sender, MAX_PLAYERS) And ULT(playerIndex, MAX_PLAYERS) Then
        InGameChatMessage_shouldReceive = aiCheckAlliances(m.sender, playerIndex)
    End If
End Function

Private Sub InGameChatMessage_getReceivers(m As InGameChatMessage, receivers() As Long, receivers_size As Long)
    Dim playerIndex As Long

    receivers_size = 0
    ReDim receivers(0 To MAX_CONNECTED_PLAYERS - 1)

    For playerIndex = 0 To MAX_CONNECTED_PLAYERS - 1
        If InGameChatMessage_shouldReceive(m, playerIndex) And openchannels(playerIndex) Then
            receivers(receivers_size) = playerIndex
            receivers_size = receivers_size + 1
        End If
    Next
End Sub

Private Function InGameChatMessage_formatReceivers(m As InGameChatMessage) As String
    Dim directs As Long, ss As String, nextName As String

    If InGameChatMessage_isGlobal(m) Then
        InGameChatMessage_formatReceivers = "Global"
        Exit Function
    End If

    If m.toAllies And m.toPlayers_size = 0 Then
        InGameChatMessage_formatReceivers = "Allies"
        Exit Function
    End If

    directs = 0
    If m.toAllies Then
        ss = ss & "Allies"
    Else
        ss = ss & "private to "
        ss = ss & getPlayerName(m.toPlayers(directs))
        directs = directs + 1
    End If

    Do While directs <> m.toPlayers_size
        nextName = getPlayerName(m.toPlayers(directs))
        directs = directs + 1
        If StrPtr(nextName) = 0 Then
            GoTo continue_
        End If
        If directs = m.toPlayers_size Then ss = ss & " and " Else ss = ss & ", "
        ss = ss & nextName
continue_:
    Loop

    InGameChatMessage_formatReceivers = ss
End Function

Private Sub InGameChatMessage_sendToHumanPlayers(m As InGameChatMessage)
    Dim formatted As String, message As NetworkTextMessage, receivers() As Long, receivers_size As Long, i As Long
    formatted = Left$(va("%s (%s): %s", getPlayerName(m.sender), InGameChatMessage_formatReceivers(m), m.text), MAX_CONSOLE_STRING_LENGTH - 1)   ' char formatted[MAX_CONSOLE_STRING_LENGTH]

    message = NetworkTextMessage_new(m.sender, formatted)
    message.teamSpecific = m.toAllies And m.toPlayers_size = 0

    If m.sender = selectedPlayer Or InGameChatMessage_shouldReceive(m, selectedPlayer) Then
        printInGameTextMessage message
    End If

    If InGameChatMessage_isGlobal(m) Then
        NetworkTextMessage_enqueue message, NETbroadcastQueue()
        Exit Sub
    End If

    InGameChatMessage_getReceivers m, receivers, receivers_size
    For i = 0 To receivers_size - 1
        If isHumanPlayer(receivers(i)) Then
            NetworkTextMessage_enqueue message, NETnetQueue(receivers(i))
        End If
    Next
End Sub

Private Sub InGameChatMessage_sendToAiPlayer(m As InGameChatMessage, ByVal receiver As Long)
    Dim responsiblePlayer As Long, w As MessageWriter

    If Not ingame.localOptionsReceived Then
        Exit Sub
    End If

    responsiblePlayer = whosResponsible(receiver)

    If responsiblePlayer >= MAX_PLAYERS And responsiblePlayer <> NetPlay_hostPlayer Then
        debug_ 131, LOG_ERROR, "sendToAiPlayer", "sendToAiPlayer() - responsiblePlayer >= MAX_PLAYERS"
        Exit Sub
    End If

    If Not isHumanPlayer(responsiblePlayer) Then
        debug_ 137, LOG_ERROR, "sendToAiPlayer", "sendToAiPlayer() - responsiblePlayer is not human."
        Exit Sub
    End If

    Set w = NETbeginEncode(NETnetQueue(responsiblePlayer), NET_AITEXTMSG)
    NETuint32_t w, m.sender
    NETuint32_t w, receiver
    NETstring w, m.text, MAX_CONSOLE_STRING_LENGTH
    NETend w
End Sub

Private Sub InGameChatMessage_sendToAiPlayers(m As InGameChatMessage)
    Dim receivers() As Long, receivers_size As Long, i As Long
    InGameChatMessage_getReceivers m, receivers, receivers_size
    For i = 0 To receivers_size - 1
        If Not isHumanPlayer(receivers(i)) Then
            If myResponsibility(receivers(i)) Then
                triggerEventChat m.sender, receivers(i), m.text
            Else
                InGameChatMessage_sendToAiPlayer m, receivers(i)
            End If
        End If
    Next
End Sub

Private Sub InGameChatMessage_sendToSpectators(m As InGameChatMessage)
    Dim formatted As String, message As NetworkTextMessage, receivers() As Long, receivers_size As Long, i As Long

    If Not ingame.localOptionsReceived Then
        Exit Sub
    End If

    formatted = Left$(va("%s (%s): %s", getPlayerName(m.sender), "Spectators", m.text), MAX_CONSOLE_STRING_LENGTH - 1)   ' char formatted[MAX_CONSOLE_STRING_LENGTH]

    If (m.sender = selectedPlayer Or InGameChatMessage_shouldReceive(m, selectedPlayer)) And NetPlay_players_isSpectator(selectedPlayer) Then
        message = NetworkTextMessage_new(m.sender, formatted)
        printInGameTextMessage message
    End If

    InGameChatMessage_getReceivers m, receivers, receivers_size
    For i = 0 To receivers_size - 1
        If isHumanPlayer(receivers(i)) And NetPlay_players_isSpectator(receivers(i)) And receivers(i) <> selectedPlayer Then
            If Not (Not myResponsibility(receivers(i))) Then ASSERT_FAILURE "!myResponsibility(receiver)", "src/chat.cpp:185", "sendToSpectators", "Should not be my responsibility..."
            InGameChatMessage_enqueueSpectatorMessage m, NETnetQueue(receivers(i)), formatted
        End If
    Next
End Sub

Private Sub InGameChatMessage_enqueueSpectatorMessage(m As InGameChatMessage, ByVal queue As Long, ByVal formattedMsg As String)
    Dim w As MessageWriter
    Set w = NETbeginEncode(queue, NET_SPECTEXTMSG)
    NETuint32_t w, m.sender
    NETstring w, formattedMsg, MAX_CONSOLE_STRING_LENGTH
    NETend w
End Sub

Public Sub InGameChatMessage_addReceiverByPosition(m As InGameChatMessage, ByVal playerPosition As Long)
    Dim playerIndex As Long
    playerIndex = findPlayerIndexByPosition(playerPosition)
    If playerIndex >= 0 Then
        InGameChatMessage_toPlayers_insert m, playerIndex
    End If
End Sub

Public Sub InGameChatMessage_addReceiverByIndex(m As InGameChatMessage, ByVal playerIndex As Long)
    InGameChatMessage_toPlayers_insert m, playerIndex
End Sub

Public Sub InGameChatMessage_send(m As InGameChatMessage)
    Dim mutedUntil As Currency, currentTime As Currency, freq As Currency
    If NetPlay_bComms Then
        If playerSpamMutedUntil(m.sender, mutedUntil) Then     ' auto mutedUntil = ...; mutedUntil.has_value()
            QueryPerformanceCounter currentTime
            QueryPerformanceFrequency freq
            addConsoleMessage "You have sent too many messages in the last few seconds. Please wait and try again.", DEFAULT_JUSTIFY, INFO_MESSAGE, False, FromU32(Fix(CDbl(mutedUntil - currentTime) * 1000# / CDbl(freq)))
            Exit Sub
        End If
        recordPlayerMessageSent m.sender
    End If

    If NetPlay_players_isSpectator(selectedPlayer) And Not NetPlay_isHost Then
        InGameChatMessage_sendToSpectators m
    Else
        InGameChatMessage_sendToHumanPlayers m
        If NetPlay_isHost And NetPlay_players_isSpectator(selectedPlayer) Then
            ' spectator hosts do get to send messages visible to all players,
            ' but not AI / scripts
            Exit Sub
        End If
        InGameChatMessage_sendToAiPlayers m
        triggerEventChat m.sender, m.sender, m.text
    End If
End Sub
