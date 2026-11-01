Attribute VB_Name = "multilobbycommands_c"
' src/multilobbycommands.cpp, multilobbycommands.h: the lobby's slash commands (processChatLobbySlashCommands, run
' by the host for its own and its players' chat) and the admin identities (the --addlobbyadminhash /
' --addlobbyadminpublickey lists; identityMatchesAdmin). The HostLobbyOperationsInterface& is multiint.cpp's
' static WzHostLobbyOperationsInterface cmdInterface, its only implementation here (stdinreader's command
' interface is not ported): its methods are multiint_lobby_c's cmdInterface_<method>. cmdInterfaceLogChatMsg
' (the command interface's output) stays a stand-in; it gets the message and the prefix (the sender's hash and
' key it would print are the command interface's).
' The std::unordered_set<std::string>s are String arrays (insert keeps one copy; the order is not used);
' getConnectedAdminPlayerIndexes' std::unordered_set<size_t> is the indexes in ascending order (MSVC's
' unordered_set iterates in insertion order while it has fewer elements than buckets).
Option Explicit
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' multilobbycommands.h
Private Const LOBBY_COMMAND_PREFIX As String = "/"
Private Const LOBBY_COMMAND_PREFIX_LENGTH As Long = 1

Private lobbyAdminPublicHashStrings() As String
Private lobbyAdminPublicHashStringsCount As Long
Private lobbyAdminPublicKeys() As String
Private lobbyAdminPublicKeysCount As Long

' (VB) unordered_set<std::string>::count / insert / erase over a String array
Private Function strSetFind(set_() As String, ByVal n As Long, ByVal s As String) As Long
    Dim i As Long
    For i = 0 To n - 1
        If set_(i) = s Then
            strSetFind = i
            Exit Function
        End If
    Next
    strSetFind = -1
End Function

Private Sub strSetInsert(set_() As String, n As Long, ByVal s As String)
    If strSetFind(set_, n, s) >= 0 Then Exit Sub
    ReDim Preserve set_(0 To n)
    set_(n) = s
    n = n + 1
End Sub

Private Function strSetErase(set_() As String, n As Long, ByVal s As String) As Long
    Dim i As Long
    i = strSetFind(set_, n, s)
    If i < 0 Then Exit Function
    set_(i) = set_(n - 1)
    n = n - 1
    If n = 0 Then Erase set_ Else ReDim Preserve set_(0 To n - 1)
    strSetErase = 1
End Function

Public Function addLobbyAdminIdentityHash(ByVal playerIdentityHash As String) As Boolean
    strSetInsert lobbyAdminPublicHashStrings, lobbyAdminPublicHashStringsCount, playerIdentityHash
    addLobbyAdminIdentityHash = True
End Function

Public Function removeLobbyAdminIdentityHash(ByVal playerIdentityHash As String) As Boolean
    removeLobbyAdminIdentityHash = strSetErase(lobbyAdminPublicHashStrings, lobbyAdminPublicHashStringsCount, playerIdentityHash) > 0
End Function

Public Function addLobbyAdminPublicKey(ByVal publicKeyB64Str As String) As Boolean
    strSetInsert lobbyAdminPublicKeys, lobbyAdminPublicKeysCount, publicKeyB64Str
    addLobbyAdminPublicKey = True
End Function

Public Function removeLobbyAdminPublicKey(ByVal publicKeyB64Str As String) As Boolean
    removeLobbyAdminPublicKey = strSetErase(lobbyAdminPublicKeys, lobbyAdminPublicKeysCount, publicKeyB64Str) > 0
End Function

' (VB) base64Encode(identity.toBytes(EcKey::Public))
Private Function publicKeyB64(identity As EcKey) As String
    Dim key() As Byte, n As Long
    n = EcKey_toBytes(identity, EcKey_Public, key)
    publicKeyB64 = base64Encode(key, n)
End Function

' checks for specific identity being an admin
Public Function identityMatchesAdmin(identity As EcKey) As Boolean
    Dim senderIdentityHash As String, senderPublicKeyB64 As String
    senderIdentityHash = EcKey_publicHashString(identity)
    senderPublicKeyB64 = publicKeyB64(identity)
    identityMatchesAdmin = strSetFind(lobbyAdminPublicKeys, lobbyAdminPublicKeysCount, senderPublicKeyB64) >= 0 _
        Or strSetFind(lobbyAdminPublicHashStrings, lobbyAdminPublicHashStringsCount, senderIdentityHash) >= 0
End Function

' **THIS** is the function that should be used to determine whether a sender currently has permission to execute admin commands
Private Function senderHasLobbyCommandAdminPrivs(ByVal playerIdx As Long, Optional ByVal quiet As Boolean = False) As Boolean
    Dim trueId As TrueIdentity, identity As EcKey, senderIdentityHash As String, senderPublicKeyB64 As String
    If Not ULT(playerIdx, MAX_CONNECTED_PLAYERS) Then
        senderHasLobbyCommandAdminPrivs = False
        Exit Function
    End If
    If playerIdx = NetPlay_hostPlayer And NetPlay_isHost Then
        ' the host always has permissions
        senderHasLobbyCommandAdminPrivs = True
        Exit Function
    End If

    trueId = getTruePlayerIdentity(playerIdx)
    If EcKey_empty(trueId.identity) Then
        senderHasLobbyCommandAdminPrivs = False
        Exit Function
    End If
    If Not identityMatchesAdmin(trueId.identity) Then
        ' identity is not in permitted list
        senderHasLobbyCommandAdminPrivs = False
        Exit Function
    End If

    ' Verify the player's identity has been verified
    If Not trueId.verified Then
        ' While this player claims to have an identity that matches an admin,
        ' they have not yet verified it by responding to a NET_PING with a valid signature
        If Not quiet Then
            identity = getOutputPlayerIdentity(playerIdx)
            senderIdentityHash = EcKey_publicHashString(identity)
            senderPublicKeyB64 = publicKeyB64(identity)
            sendRoomSystemMessageToSingleReceiver "Waiting for sync (admin privileges not yet enabled)", playerIdx, True
            If strSetFind(lobbyAdminPublicKeys, lobbyAdminPublicKeysCount, senderPublicKeyB64) >= 0 Then
                debug_ 105, LOG_INFO, "senderHasLobbyCommandAdminPrivs", "Received an admin check for player %u that passed (public key: %s), but they have not yet verified their identity", playerIdx, senderPublicKeyB64
            Else
                debug_ 109, LOG_INFO, "senderHasLobbyCommandAdminPrivs", "Received an admin check for player %u that passed (public identity: %s), but they have not yet verified their identity", playerIdx, senderIdentityHash
            End If
        End If
        senderHasLobbyCommandAdminPrivs = False
        Exit Function
    End If
    senderHasLobbyCommandAdminPrivs = True
End Function

Private Sub lobbyCommand_PrintHelp(ByVal receiver As Long)
    sendRoomSystemMessageToSingleReceiver "Command list:", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "help - Get this message", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "admin - Display currently-connected admin players", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "me - Display your information", receiver, True
    If Not senderHasLobbyCommandAdminPrivs(receiver, True) Then
        sendRoomSystemMessageToSingleReceiver "(Additional commands are available for admins)", receiver, True
        Exit Sub
    End If
    ' admin-only commands
    sendRoomSystemMessageToSingleReceiver "Admin-only commands: (All slots count from 0)", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "swap <slot-from> <slot-to> - Swap player/slot positions", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "makespec <slot> - Move a player to spectators", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "makeplayer s<slot> - Request to move a spectator to players", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "kick <slot> - Kick a player; (or s<slot> for spectator - ex. s0)", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "team <slot> <team> - Change team for player/slot", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "base <base level> - Change base level (0, 1, 2)", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "alliance <alliance type> - Change alliance setting (0, 1, 2, 3)", receiver, True
    sendRoomSystemMessageToSingleReceiver LOBBY_COMMAND_PREFIX & "scav <scav level> - Change scav setting (0=off, 1=on, 2=ultimate)", receiver, True
End Sub

' (VB) std::unordered_set<size_t>: the indexes in idx(0 .. n - 1), n returned
Private Function getConnectedAdminPlayerIndexes(idx() As Long) As Long
    Dim playerIdx As Long, lim As Long, n As Long
    lim = NetPlay_players_size()
    If lim > MAX_CONNECTED_PLAYERS Then lim = MAX_CONNECTED_PLAYERS
    For playerIdx = 0 To lim - 1
        If senderHasLobbyCommandAdminPrivs(playerIdx, True) Then
            ReDim Preserve idx(0 To n)
            idx(n) = playerIdx
            n = n + 1
        End If
    Next
    getConnectedAdminPlayerIndexes = n
End Function

Private Sub lobbyCommand_Admin()
    Dim adminPlayerIndexes() As Long, n As Long, msg As String, currNum As Long, i As Long, adminPlayerIdx As Long
    ' get list of connected admins
    n = getConnectedAdminPlayerIndexes(adminPlayerIndexes)
    If n = 0 Then
        sendRoomSystemMessage "No room admins currently connected"
        Exit Sub
    End If
    msg = "Currently connected room admins:"
    currNum = 0
    For i = 0 To n - 1
        adminPlayerIdx = adminPlayerIndexes(i)
        If adminPlayerIdx > MAX_CONNECTED_PLAYERS Then
            debug_ 168, LOG_ERROR, "lobbyCommand_Admin", "Invalid admin index?? %zu", adminPlayerIdx
            GoTo continue_
        End If
        If currNum > 0 Then
            msg = msg & ","
        End If
        msg = msg & " ["
        msg = msg & CStr(adminPlayerIdx) & "] "
        msg = msg & getPlayerName(adminPlayerIdx, True)
        currNum = currNum + 1
continue_:
    Next
    sendRoomSystemMessage msg
End Sub

' #define ADMIN_REQUIRED_FOR_COMMAND(command): False (the caller returns false) when the sender is no admin
Private Function adminRequiredForCommand(message As NetworkTextMessage, ByVal command As String) As Boolean
    If Not senderHasLobbyCommandAdminPrivs(message.sender) Then
        sendRoomSystemMessageToSingleReceiver "Only admin can use the """ & command & """ command", message.sender, True
        adminRequiredForCommand = False
        Exit Function
    End If
    adminRequiredForCommand = True
End Function

' the lambda posToNetPlayer: the player index at map position a, else a
Private Function posToNetPlayer(ByVal a As Long) As Long
    Dim i As Long
    For i = 0 To MAX_PLAYERS - 1
        If NetPlay_players_position(i) = a Then
            posToNetPlayer = i
            Exit Function
        End If
    Next
    posToNetPlayer = a
End Function

' (VB) strncmp(cmd, prefix, Len(prefix)) == 0
Private Function startsWith(ByVal cmd As String, ByVal prefix As String) As Boolean
    startsWith = (Left$(cmd, Len(prefix)) = prefix)
End Function

Public Function processChatLobbySlashCommands(message As NetworkTextMessage) As Boolean
    Dim startingCommandPosition As Long, identity As EcKey, senderhash As String, senderPublicKeyB64 As String
    Dim cmd As String, msg As String, r As Long, s1 As Long, s2 As Long, isBan As Boolean, command As String
    Dim playerPos As Long, playerIdx As Long, playerIdxA As Long, alliancesType As Long, scavsValue As Long
    Dim isMute As Boolean

    If message.sender = SYSTEM_MESSAGE Or message.sender = NOTIFY_MESSAGE Or message.sender < 0 Then
        processChatLobbySlashCommands = False
        Exit Function
    End If

    If Left$(message.text, 1) <> LOBBY_COMMAND_PREFIX Then
        ' Message does not start with command prefix
        processChatLobbySlashCommands = False
        Exit Function
    End If

    If Not NetPlay_isHost Then
        ' only the host should process lobby chat commands
        processChatLobbySlashCommands = False
        Exit Function
    End If

    If Not (message.sender < MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "message.sender < MAX_CONNECTED_PLAYERS", "src/multilobbycommands.cpp:237", "processChatLobbySlashCommands", "Out of bounds message sender?: %i", message.sender: processChatLobbySlashCommands = False: Exit Function

    ' strip /hostmsg off the front of the command, if present (i.e. if starts with "<LOBBY_COMMAND_PREFIX>hostmsg <LOBBY_COMMAND_PREFIX>")
    startingCommandPosition = LOBBY_COMMAND_PREFIX_LENGTH
    If Mid$(message.text, startingCommandPosition + 1, 7 + 1 + LOBBY_COMMAND_PREFIX_LENGTH) = "hostmsg " & LOBBY_COMMAND_PREFIX Then
        startingCommandPosition = startingCommandPosition + 7 + 1 + LOBBY_COMMAND_PREFIX_LENGTH
    End If
    cmd = Mid$(message.text, startingCommandPosition + 1)       ' &message.text[startingCommandPosition]

    identity = getOutputPlayerIdentity(message.sender)
    senderhash = EcKey_publicHashString(identity, 64)
    senderPublicKeyB64 = publicKeyB64(identity)
    debug_ 257, LOG_INFO, "processChatLobbySlashCommands", "message [%s] [%s]", senderhash, message.text
    cmdInterfaceLogChatMsg message, "WZCHATCMD"     ' (senderhash, senderPublicKeyB64: the command interface's)
    If cmd = "help" Then
        lobbyCommand_PrintHelp message.sender
    ElseIf cmd = "admin" Or cmd = "admins" Then
        lobbyCommand_Admin
    ElseIf cmd = "me" Then
        msg = va("Your player information:" & vbLf & "identity: %s" & vbLf & "hash: %s" & vbLf & "sender [%d] position [%d] name [%s]", _
                 senderPublicKeyB64, senderhash, message.sender, NetPlay_players_position(message.sender), getPlayerName(message.sender, True))
        sendRoomSystemMessageToSingleReceiver msg, message.sender, True
    ElseIf startsWith(cmd, "team ") Then
        If Not adminRequiredForCommand(message, "team") Then Exit Function
        s1 = 0: s2 = 0
        r = sscanf(cmd, "team %u %u", s1, s2)
        If r <> 2 Or Not ULT(s1, MAX_PLAYERS) Or Not ULT(s2, MAX_PLAYERS) Then
            sendRoomNotifyMessage "Usage: " & LOBBY_COMMAND_PREFIX & "team <slot> <team>"
            processChatLobbySlashCommands = False
            Exit Function
        End If
        If Not cmdInterface_changeTeam(posToNetPlayer(s1), s2, message.sender) Then
            msg = va("Unable to change player %u team to %u", s1, s2)
            sendRoomNotifyMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If

        msg = va("Changed player %u team to %u", s1, s2)
        sendRoomSystemMessage msg
    ElseIf cmd = "hostexit" Then
        If Not adminRequiredForCommand(message, "hostexit") Then Exit Function
        cmdInterface_quitGame 5
    ElseIf startsWith(cmd, "kick ") Or startsWith(cmd, "ban ") Then
        isBan = startsWith(cmd, "ban")
        If Not isBan Then
            If Not adminRequiredForCommand(message, "kick") Then Exit Function
        Else
            If Not adminRequiredForCommand(message, "ban") Then Exit Function
        End If
        If Not isBan Then command = "kick" Else command = "ban"
        playerPos = MAX_PLAYERS + 1
        playerIdx = MAX_CONNECTED_PLAYERS + 1
        r = sscanf(cmd, command & " %u", playerPos)
        If r = 1 Then
            playerIdx = posToNetPlayer(playerPos)
            If Not ULT(playerIdx, MAX_PLAYERS) Then
                msg = "Usage: " & LOBBY_COMMAND_PREFIX & command & " <slot>"
                sendRoomNotifyMessage msg
                processChatLobbySlashCommands = False
                Exit Function
            End If
        Else
            r = sscanf(cmd, command & " s%u", playerPos)
            If r <> 1 Or Not ULT(playerPos, MAX_SPECTATOR_SLOTS) Then
                msg = "Usage: " & LOBBY_COMMAND_PREFIX & command & " <slot>"
                sendRoomNotifyMessage msg
                processChatLobbySlashCommands = False
                Exit Function
            End If
            playerIdx = MAX_PLAYER_SLOTS + playerPos
            If Not (playerIdx < MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/multilobbycommands.cpp:339", "processChatLobbySlashCommands", "Invalid index: %u", playerIdx: processChatLobbySlashCommands = False: Exit Function
        End If
        If playerIdx = NetPlay_hostPlayer Then
            ' Can't kick the host...
            sendRoomSystemMessage "Can't kick the host."
            processChatLobbySlashCommands = False
            Exit Function
        End If
        If Not cmdInterface_kickPlayer(playerIdx, "Administrator has kicked you from the game.", isBan, message.sender) Then
            msg = va("Failed to kick %s: %u", IIf(playerIdx < MAX_PLAYER_SLOTS, "player", "spectator"), playerPos)
            sendRoomSystemMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If
    ElseIf startsWith(cmd, "swap ") Then
        If Not adminRequiredForCommand(message, "swap") Then Exit Function
        r = sscanf(cmd, "swap %u %u", s1, s2)
        playerIdxA = posToNetPlayer(s1)
        If r <> 2 Then
            sendRoomNotifyMessage "Usage: " & LOBBY_COMMAND_PREFIX & "swap <slot-from> <slot-to>"
            processChatLobbySlashCommands = False
            Exit Function
        End If

        If playerIdxA >= IIf(game.maxPlayers < MAX_PLAYERS, game.maxPlayers, MAX_PLAYERS) Then
            msg = va("Swap - invalid player 1: %u", playerIdxA)
            sendRoomNotifyMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If
        If Not ULT(s2, game.maxPlayers) Then
            msg = va("Swap - invalid player 2: %u", s2)
            sendRoomNotifyMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If

        If Not cmdInterface_changePosition(playerIdxA, s2, message.sender) Then
            ' (PRIu8 is "hhu": the low 8 bits of each)
            msg = va("Unable to swap players %u and %u", s1 And &HFF&, s2 And &HFF&)
            sendRoomNotifyMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If

        sendRoomSystemMessage "Swapping player " & va("%u", s1) & " and " & va("%u", s2)
    ElseIf startsWith(cmd, "base ") Then
        If Not adminRequiredForCommand(message, "base") Then Exit Function
        r = sscanf(cmd, "base %d", s1)
        If r <> 1 Then
            sendRoomNotifyMessage "Usage: " & LOBBY_COMMAND_PREFIX & "base <base level>"
            processChatLobbySlashCommands = False
            Exit Function
        End If

        If s1 < 0 Or s1 > 2 Then
            sendRoomNotifyMessage "Base level must be between [0 - 2]"
            processChatLobbySlashCommands = False
            Exit Function
        End If
        If Not cmdInterface_changeBase(s1) Then
            msg = va("Unable to change base to: %u", s1 And &HFF&)
            sendRoomNotifyMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If

        sendRoomSystemMessage "Starting base set to " & CStr(s1)
    ElseIf startsWith(cmd, "alliance ") Then
        If Not adminRequiredForCommand(message, "alliance") Then Exit Function
        r = sscanf(cmd, "alliance %d", s1)
        If r <> 1 Then
            sendRoomNotifyMessage "Usage: " & LOBBY_COMMAND_PREFIX & "alliance <alliance type>"
            processChatLobbySlashCommands = False
            Exit Function
        Else
            If s1 < 0 Or s1 > 3 Then
                sendRoomNotifyMessage "Alliance type must be in range [0 - 3]"
                processChatLobbySlashCommands = False
                Exit Function
            End If

            alliancesType = s1
            ' flip 2nd and 3rd values to match the order displayed in UI
            If alliancesType = 3 Then
                alliancesType = 2
            ElseIf alliancesType = 2 Then
                alliancesType = 3
            End If
            If Not cmdInterface_changeAlliances(alliancesType) Then
                msg = va("Unable to set alliances to: %u", alliancesType)
                sendRoomNotifyMessage msg
                processChatLobbySlashCommands = False
                Exit Function
            End If

            sendRoomSystemMessage "Alliance type set to " & CStr(s1)
        End If
    ElseIf startsWith(cmd, "scav ") Then
        If Not adminRequiredForCommand(message, "scav") Then Exit Function
        r = sscanf(cmd, "scav %d", scavsValue)
        If r <> 1 Then
            sendRoomNotifyMessage "Usage: " & LOBBY_COMMAND_PREFIX & "scav <0/1/2>"
            processChatLobbySlashCommands = False
            Exit Function
        End If

        If Not (scavsValue = NO_SCAVENGERS Or scavsValue = SCAVENGERS Or scavsValue = ULTIMATE_SCAVENGERS) Then
            sendRoomNotifyMessage "Scavs must be off, on, or ultimate: (0, 1, or 2)"
            processChatLobbySlashCommands = False
            Exit Function
        End If

        If Not cmdInterface_changeScavengers(scavsValue) Then
            msg = va("Unable to set scavs to: %d", scavsValue)
            sendRoomNotifyMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If

        sendRoomSystemMessage "Scavangers set to " & CStr(scavsValue)
    ElseIf startsWith(cmd, "makespec ") Then
        If Not adminRequiredForCommand(message, "makespec") Then Exit Function
        playerPos = MAX_PLAYERS + 1
        r = sscanf(cmd, "makespec %u", playerPos)
        playerIdx = posToNetPlayer(playerPos)
        If r <> 1 Or Not ULT(playerPos, MAX_PLAYERS) Then
            sendRoomNotifyMessage "Usage: " & LOBBY_COMMAND_PREFIX & "makespec <slot>"
            processChatLobbySlashCommands = False
            Exit Function
        End If
        If playerIdx = NetPlay_hostPlayer Then
            ' Can't move the host...
            sendRoomSystemMessage "Can't move the host."
            processChatLobbySlashCommands = False
            Exit Function
        End If
        If playerIdx = message.sender Then
            ' Can't move self this way (or it'll prevent us from moving back) - use the UI!
            sendRoomSystemMessageToSingleReceiver "Use the UI to move yourself.", message.sender, True
            processChatLobbySlashCommands = False
            Exit Function
        End If
        If Not cmdInterface_movePlayerToSpectators(playerIdx) Then
            msg = va("Failed to move player to spectators: %u", playerPos)
            sendRoomSystemMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If
    ElseIf startsWith(cmd, "makeplayer ") Then
        If Not adminRequiredForCommand(message, "makeplayer") Then Exit Function
        playerPos = MAX_PLAYERS + 1
        r = sscanf(cmd, "makeplayer s%u", playerPos)
        If r <> 1 Or Not ULT(playerPos, MAX_SPECTATOR_SLOTS) Then
            sendRoomNotifyMessage "Usage: " & LOBBY_COMMAND_PREFIX & "makeplayer s<spectator slot>"
            processChatLobbySlashCommands = False
            Exit Function
        End If
        playerIdx = MAX_PLAYER_SLOTS + playerPos
        If Not (playerIdx < MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/multilobbycommands.cpp:518", "processChatLobbySlashCommands", "Invalid index: %u", playerIdx: processChatLobbySlashCommands = False: Exit Function
        If playerIdx = NetPlay_hostPlayer Then
            ' Can't move the host...
            sendRoomSystemMessage "Can't move the host."
            processChatLobbySlashCommands = False
            Exit Function
        End If
        If Not cmdInterface_requestMoveSpectatorToPlayers(playerIdx) Then
            msg = va("Failed to move player to spectators: %u", playerPos)
            sendRoomSystemMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If
    ElseIf startsWith(cmd, "mute ") Or startsWith(cmd, "unmute ") Then
        isMute = startsWith(cmd, "mute")
        If Not isMute Then
            If Not adminRequiredForCommand(message, "mute") Then Exit Function
        Else
            If Not adminRequiredForCommand(message, "unmute") Then Exit Function
        End If
        If isMute Then command = "mute" Else command = "unmute"
        playerPos = MAX_PLAYERS + 1
        playerIdx = MAX_CONNECTED_PLAYERS + 1
        r = sscanf(cmd, command & " %u", playerPos)
        If r = 1 Then
            playerIdx = posToNetPlayer(playerPos)
            If Not ULT(playerIdx, MAX_PLAYERS) Then
                msg = "Usage: " & LOBBY_COMMAND_PREFIX & command & " <slot>"
                sendRoomNotifyMessage msg
                processChatLobbySlashCommands = False
                Exit Function
            End If
        Else
            r = sscanf(cmd, command & " s%u", playerPos)
            If r <> 1 Or Not ULT(playerPos, MAX_SPECTATOR_SLOTS) Then
                msg = "Usage: " & LOBBY_COMMAND_PREFIX & command & " <slot>"
                sendRoomNotifyMessage msg
                processChatLobbySlashCommands = False
                Exit Function
            End If
            playerIdx = MAX_PLAYER_SLOTS + playerPos
            If Not (playerIdx < MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "playerIdx < MAX_CONNECTED_PLAYERS", "src/multilobbycommands.cpp:569", "processChatLobbySlashCommands", "Invalid index: %u", playerIdx: processChatLobbySlashCommands = False: Exit Function
        End If
        If playerIdx = NetPlay_hostPlayer Then
            ' Can't mute the host...
            sendRoomSystemMessage "Can't mute the host."
            processChatLobbySlashCommands = False
            Exit Function
        End If
        If Not cmdInterface_changeHostChatPermissions(playerIdx, Not isMute) Then
            msg = va("Failed to mute / unmute %s: %u", IIf(playerIdx < MAX_PLAYER_SLOTS, "player", "spectator"), playerPos)
            sendRoomSystemMessage msg
            processChatLobbySlashCommands = False
            Exit Function
        End If
    Else
        ' unrecognized command - not handled
        processChatLobbySlashCommands = False
        Exit Function
    End If

    processChatLobbySlashCommands = True
End Function
