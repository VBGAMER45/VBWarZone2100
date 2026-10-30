Attribute VB_Name = "multijoin_c"
' src/multijoin.cpp, multijoin.h: so far clearPlayer and
' destroyPlayerResources (a player's objects removed at the game's start or
' when the player leaves). NetPlay.playerReferences (lib/netplay's
' PlayerReference objects) is not ported yet (cluster J): clearPlayer's
' disconnect of the old reference is left out. Not ported yet: the players
' leaving and joining during a game (splitResourcesAmongTeam,
' handlePlayerLeftInGame, recvPlayerLeft, MultiPlayerLeave, MultiPlayerJoin,
' the data checks, setupNewPlayer, ShowLobbyStatusMessage): cluster J.
Option Explicit

' lib/netplay/nettypes.h
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

'//////////////////////////////////////////////////////////////////////////////
'
' when a remote player leaves an arena game do this!
'
' @param player -- the one we need to clear
' @param quietly -- true means without any visible effects
'
Public Sub clearPlayer(ByVal player As Long, ByVal quietly As Boolean)
    Dim q As String
    If Not (ULT(player, MAX_CONNECTED_PLAYERS)) Then ASSERT_FAILURE "player < MAX_CONNECTED_PLAYERS", "src/multijoin.cpp:94", "clearPlayer", "Invalid player: %u", player: Exit Sub

    ' (NetPlay.playerReferences: cluster J)
    ' ASSERT(player < NetPlay.playerReferences.size(), "Invalid player: %" PRIu32 "", player);
    ' NetPlay.playerReferences[player]->disconnect();
    ' NetPlay.playerReferences[player] = std::make_shared<PlayerReference>(player);

    If quietly Then q = "true" Else q = "false"
    ' (VB) the macro evaluates getPlayerName only with LOG_NET on (it asserts for the empty slots past MAX_PLAYERS)
    If debugPartEnabled(LOG_NET) Then debug_ 100, LOG_NET, "clearPlayer", "R.I.P. %s (%u). quietly is %s", getPlayerName(player), player, q

    ingame.LagCounter(player) = 0
    ingame.DesyncCounter(player) = 0
    ingame.JoiningInProgress(player) = False   ' if they never joined, reset the flag
    ingame.DataIntegrity(player) = False
    ingame.hostChatPermissions(player) = False
    ingame.has_lastSentPlayerDataCheck2(player) = False

    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub ' no more to do
    End If

    destroyPlayerResources player, quietly
End Sub

Public Sub destroyPlayerResources(ByVal player As Long, ByVal quietly As Boolean)
    Dim i As Long, it As Long, nxt As Long, psStruct As Long

    If Not ULT(player, MAX_PLAYERS) Then
        Exit Sub ' no more to do
    End If

    For i = 0 To MAX_PLAYERS - 1                ' remove alliances
        ' Never remove a player's self-alliance, as the player can be selected and units added via the debug menu
        ' even after they have left, and this would lead to them firing on each other.
        If i <> player Then
            alliances(player, i) = ALLIANCE_BROKEN
            alliances(i, player) = ALLIANCE_BROKEN
            alliancebits(i) = alliancebits(i) And Not BIT_(player)
            alliancebits(player) = alliancebits(player) And Not BIT_(i)
        End If
    Next

    debug_ 139, LOG_DEATH, "destroyPlayerResources", "killing off all droids for player %d", player
    ' delete all droids
    ' mutating_list_iterate
    it = apsDroidLists(player).head
    Do While it <> 0
        nxt = lnNext(it)
        If quietly Then             ' don't show effects
            killDroid lnObj(it)
        Else                        ' show effects
            destroyDroid lnObj(it), gameTime
        End If
        it = nxt
    Loop

    debug_ 154, LOG_DEATH, "destroyPlayerResources", "killing off all structures for player %d", player
    ' delete all structs
    ' mutating_list_iterate
    it = apsStructLists(player).head
    Do While it <> 0
        nxt = lnNext(it)
        psStruct = lnObj(it)
        ' FIXME: look why destroyStruct() doesn't put back the feature like removeStruct() does
        If quietly Or asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then       ' don't show effects
            removeStruct psStruct, True
        Else                        ' show effects
            destroyStruct psStruct, gameTime
        End If
        it = nxt
    Loop

    Exit Sub
End Sub
