Attribute VB_Name = "input_debugmappings_c"
' src/input/debugmappings.cpp, src/input/debugmappings.h: whether the debug key mappings (the cheats)
' may run: every human player must want them.
' (VB) The one DebugInputManager is display.cpp's gInputManager's (InputManager's debugManager()): its
' methods are DebugInputManager_<method>. The callers' earlier spellings forward to them:
' debugMappingsAllowed(), gInputManager_debugManager_debugMappingsAllowed() and
' dbgInputManager_setPlayerWantsDebugMappings.
Option Explicit

Private bDebugPrioritized As Boolean
Private bDoingDebugMappings As Boolean
Private playerWantsDebugMappings() As Boolean   ' std::vector<bool>
Private playerWantsDebugMappings_size As Long

' DebugInputManager::DebugInputManager(const unsigned int maxPlayers)
Public Sub DebugInputManager_construct(ByVal maxPlayers As Long)
    bDoingDebugMappings = False
    playerWantsDebugMappings_size = maxPlayers
    If maxPlayers > 0 Then
        ReDim playerWantsDebugMappings(0 To maxPlayers - 1)
    Else
        Erase playerWantsDebugMappings
    End If
End Sub

' bool DebugInputManager::debugMappingsAllowed() const
Public Function DebugInputManager_debugMappingsAllowed() As Boolean
    DebugInputManager_debugMappingsAllowed = bDoingDebugMappings
End Function

' bool DebugInputManager::getPlayerWantsDebugMappings(const unsigned int playerIndex) const
Public Function DebugInputManager_getPlayerWantsDebugMappings(ByVal playerIndex As Long) As Boolean
    If Not (ULT(playerIndex, playerWantsDebugMappings_size)) Then ASSERT_FAILURE "playerIndex < playerWantsDebugMappings.size()", "src/input/debugmappings.cpp:39", "getPlayerWantsDebugMappings", "Tried to get debug mapping status for playerIndex > maxPlayers": DebugInputManager_getPlayerWantsDebugMappings = False: Exit Function

    DebugInputManager_getPlayerWantsDebugMappings = playerWantsDebugMappings(playerIndex)
End Function

' void DebugInputManager::setPlayerWantsDebugMappings(const unsigned int playerIndex, const bool bWants)
Public Sub DebugInputManager_setPlayerWantsDebugMappings(ByVal playerIndex As Long, ByVal bWants As Boolean)
    Dim bIsTrueMultiplayerGame As Boolean, n As Long
    Dim bIsEmptySlot As Boolean, bIsSpectatorSlot As Boolean, bPlayerNWantsDebugMappings As Boolean
    If Not (ULT(playerIndex, playerWantsDebugMappings_size)) Then ASSERT_FAILURE "playerIndex < playerWantsDebugMappings.size()", "src/input/debugmappings.cpp:46", "setPlayerWantsDebugMappings", "playerIndex is invalid: %u", playerIndex: Exit Sub
    playerWantsDebugMappings(playerIndex) = bWants
    bDoingDebugMappings = True
    bIsTrueMultiplayerGame = bMultiPlayer And NetPlay_bComms
    For n = 0 To playerWantsDebugMappings_size - 1
        If Not (n < MAX_PLAYERS) Then ASSERT_FAILURE "n < MAX_PLAYERS", "src/input/debugmappings.cpp:52", "setPlayerWantsDebugMappings", "playerWantsDebugMappings has more entries than MAX_PLAYERS": Exit Sub

        bIsEmptySlot = Not NetPlay_players_allocated(n)
        bIsSpectatorSlot = NetPlay_players_isSpectator(n)
        bPlayerNWantsDebugMappings = bIsEmptySlot Or (bIsSpectatorSlot And bIsTrueMultiplayerGame) Or playerWantsDebugMappings(n)
        bDoingDebugMappings = bDoingDebugMappings And bPlayerNWantsDebugMappings
    Next
End Sub

' bool DebugInputManager::toggleDebugMappingPriority()
Public Function DebugInputManager_toggleDebugMappingPriority() As Boolean
    bDebugPrioritized = Not bDebugPrioritized
    DebugInputManager_toggleDebugMappingPriority = bDebugPrioritized
End Function

' bool DebugInputManager::isDebugPrioritized() const
Public Function DebugInputManager_isDebugPrioritized() As Boolean
    DebugInputManager_isDebugPrioritized = bDebugPrioritized
End Function

' (VB) the callers' spellings of gInputManager.debugManager().debugMappingsAllowed()
Public Function debugMappingsAllowed() As Boolean
    debugMappingsAllowed = bDoingDebugMappings
End Function

Public Function gInputManager_debugManager_debugMappingsAllowed() As Boolean
    gInputManager_debugManager_debugMappingsAllowed = bDoingDebugMappings
End Function

' (VB) init.cpp's dbgInputManager.setPlayerWantsDebugMappings(n, false)
Public Sub dbgInputManager_setPlayerWantsDebugMappings(ByVal playerIdx As Long, ByVal bWants As Boolean)
    DebugInputManager_setPlayerWantsDebugMappings playerIdx, bWants
End Sub
