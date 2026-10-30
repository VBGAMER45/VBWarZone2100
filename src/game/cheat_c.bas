Attribute VB_Name = "cheat_c"
' src/cheat.cpp, cheat.h: the cheat codes typed in the chat. CHEAT_ENTRY's function pointer is the entry's
' index, run by cheatCodes_function (VB: Select Case); strcasecmp is a compare of the lower-cased texts
' (the codes are ASCII). _attemptCheatCode is attemptCheatCode_ (CONVENTIONS: a leading underscore moves
' to the end).
'
' Handles cheat codes for Warzone.
' Alex M 19th - Jan. 1999
Option Explicit

Public Cheated As Boolean               ' bool Cheated = false;

' static CHEAT_ENTRY cheatCodes[]: the names (pName); the functions in cheatCodes_function
Private Const NUM_CHEAT_CODES As Long = 47     ' ARRAY_SIZE(cheatCodes)

Private Function cheatCodes_pName(ByVal i As Long) As String
    Select Case i
    Case 0: cheatCodes_pName = "templates"          ' print templates
    Case 1: cheatCodes_pName = "jsdebug"            ' show scripting states
    Case 2: cheatCodes_pName = "teach us"           ' give experience to selected units
    Case 3: cheatCodes_pName = "makemehero"         ' make selected units Heros
    Case 4: cheatCodes_pName = "untouchable"        ' make selected droids unselectable
    Case 5: cheatCodes_pName = "clone wars"         ' clone selected units
    Case 6: cheatCodes_pName = "clone wars!"        ' clone selected units
    Case 7: cheatCodes_pName = "clone wars!!"       ' clone selected units
    Case 8: cheatCodes_pName = "noassert"           ' turn off asserts
    Case 9: cheatCodes_pName = "count me"           ' give a count of objects in the world
    Case 10: cheatCodes_pName = "give all"          ' give all
    Case 11: cheatCodes_pName = "research all"      ' research everything at once
    Case 12: cheatCodes_pName = "superpower"        ' get tons of power
    Case 13: cheatCodes_pName = "more power"        ' get tons of power
    Case 14: cheatCodes_pName = "deity"             ' from above
    Case 15: cheatCodes_pName = "droidinfo"         ' show unit stats
    Case 16: cheatCodes_pName = "sensors"           ' show sensor ranges
    Case 17: cheatCodes_pName = "timedemo"          ' timedemo
    Case 18: cheatCodes_pName = "kill"              ' kill selected
    Case 19: cheatCodes_pName = "john kettley"      ' john kettley
    Case 20: cheatCodes_pName = "mouseflip"         ' mouseflip
    Case 21: cheatCodes_pName = "biffer baker"      ' almost invincible units
    Case 22: cheatCodes_pName = "supereasy"         ' supereasy
    Case 23: cheatCodes_pName = "easy"              ' easy
    Case 24: cheatCodes_pName = "normal"            ' normal
    Case 25: cheatCodes_pName = "hard"              ' hard
    Case 26: cheatCodes_pName = "insane"            ' insane
    Case 27: cheatCodes_pName = "double up"         ' your units take half the damage
    Case 28: cheatCodes_pName = "whale fin"         ' turns on/off infinte power
    Case 29: cheatCodes_pName = "get off my land"   ' kills all enemy units and structures
    Case 30: cheatCodes_pName = "build info"        ' tells you when the game was built
    Case 31: cheatCodes_pName = "time toggle"
    Case 32: cheatCodes_pName = "work harder"
    Case 33: cheatCodes_pName = "tileinfo"          ' output debug info about a tile
    Case 34: cheatCodes_pName = "showfps"           ' displays your average FPS
    Case 35: cheatCodes_pName = "showunits"         ' displays unit count information
    Case 36: cheatCodes_pName = "showsamples"       ' displays the # of Sound samples in Queue & List
    Case 37: cheatCodes_pName = "showorders"        ' displays unit order/action state.
    Case 38: cheatCodes_pName = "pause"             ' Pause the game.
    Case 39: cheatCodes_pName = "power info"
    Case 40: cheatCodes_pName = "reload me"         ' reload selected weapons immediately
    Case 41: cheatCodes_pName = "desync me"
    Case 42: cheatCodes_pName = "damage me"
    Case 43: cheatCodes_pName = "autogame on"
    Case 44: cheatCodes_pName = "autogame off"
    Case 45: cheatCodes_pName = "shakey"            ' shakey
    Case 46: cheatCodes_pName = "list droids"
    Case Else: cheatCodes_pName = ""
    End Select
End Function

' curCheat->function(): the entry's function
Private Sub cheatCodes_function(ByVal i As Long)
    Select Case i
    Case 0: listTemplates
    Case 1: jsShowDebug
    Case 2: kf_TeachSelected
    Case 3: kf_MakeMeHero
    Case 4: kf_Unselectable
    Case 5: kf_CloneSelected 10             ' []{ kf_CloneSelected(10); }
    Case 6: kf_CloneSelected 40             ' []{ kf_CloneSelected(40); }
    Case 7: kf_CloneSelected 135            ' []{ kf_CloneSelected(135); }
    Case 8: kf_NoAssert
    Case 9: kf_ShowNumObjects
    Case 10: kf_AllAvailable
    Case 11: kf_FinishAllResearch
    Case 12: kf_MaxPower
    Case 13: kf_UpThePower
    Case 14: kf_ToggleGodMode
    Case 15: kf_DebugDroidInfo
    Case 16: kf_ToggleSensorDisplay
    Case 17: kf_FrameRate
    Case 18: kf_KillSelected
    Case 19: kf_ToggleWeather
    Case 20: kf_ToggleMouseInvert
    Case 21: kf_BifferBaker
    Case 22: kf_SetDifficultyLevel DL_SUPER_EASY
    Case 23: kf_SetDifficultyLevel DL_EASY
    Case 24: kf_SetDifficultyLevel DL_NORMAL
    Case 25: kf_SetDifficultyLevel DL_HARD
    Case 26: kf_SetDifficultyLevel DL_INSANE
    Case 27: kf_DoubleUp
    Case 28: kf_TogglePower
    Case 29: kf_KillEnemy
    Case 30: kf_BuildInfo
    Case 31: kf_ToggleMissionTimer
    Case 32: kf_FinishResearch
    Case 33: kf_TileInfo
    Case 34: kf_ToggleFPS
    Case 35: kf_ToggleUnitCount
    Case 36: kf_ToggleSamples
    Case 37: kf_ToggleOrders
    Case 38: kf_TogglePauseMode
    Case 39: kf_PowerInfo
    Case 40: kf_Reload
    Case 41: kf_ForceDesync
    Case 42: kf_DamageMe
    Case 43: kf_AutoGame
    Case 44: kf_AutoGame
    Case 45: kf_ToggleShakeStatus
    Case 46: kf_ListDroids
    End Select
End Sub

' bool _attemptCheatCode(const char *cheat_name)
Private Function attemptCheatCode_(ByVal cheat_name As String) As Boolean
    Dim curCheat As Long, buf As String

    ' there is no reason to make people enter "cheat mode" to enter following commands
    If LCase$("showfps") = LCase$(cheat_name) Then
        kf_ToggleFPS
        attemptCheatCode_ = True
        Exit Function
    End If
    If LCase$("showunits") = LCase$(cheat_name) Then
        kf_ToggleUnitCount
        attemptCheatCode_ = True
        Exit Function
    End If
    If LCase$("specstats") = LCase$(cheat_name) Then
        kf_ToggleSpecOverlays
        attemptCheatCode_ = True
        Exit Function
    End If

    If LCase$("help") = LCase$(cheat_name) Then
        intShowWidgetHelp
        attemptCheatCode_ = True
        Exit Function
    End If

    ' (const DebugInputManager& dbgInputManager = gInputManager.debugManager(): input_debugmappings_c's module)
    If cheat_name = "cheat on" Or cheat_name = "debug" Then
        If Not DebugInputManager_debugMappingsAllowed() Then
            kf_ToggleDebugMappings
        End If
        attemptCheatCode_ = True
        Exit Function
    End If
    If cheat_name = "cheat off" Then
        If DebugInputManager_debugMappingsAllowed() Then
            kf_ToggleDebugMappings
            attemptCheatCode_ = True
            Exit Function
        End If
    End If
    If Not DebugInputManager_debugMappingsAllowed() Then
        attemptCheatCode_ = False
        Exit Function
    End If

    For curCheat = 0 To NUM_CHEAT_CODES - 1     ' for (curCheat = cheatCodes; curCheat != EndCheat; ++curCheat)
        If LCase$(cheat_name) = LCase$(cheatCodes_pName(curCheat)) Then
            ' We've got our man...
            cheatCodes_function curCheat    ' run it

            ' Copy this info to be used by the crash handler for the dump file
            buf = Left$("User has used cheat code: " & cheatCodes_pName(curCheat), 255)  ' ssprintf(buf, ...): char buf[256]
            addDumpInfo buf

            ' And get out of here
            Cheated = True
            attemptCheatCode_ = True
            Exit Function
        End If
    Next

    attemptCheatCode_ = False
End Function

Public Function attemptCheatCode(ByVal cheat_name As String) As Boolean
    Dim result As Boolean
    result = attemptCheatCode_(cheat_name)
    If result Then
        ActivityManager_cheatUsed cheat_name
    End If
    attemptCheatCode = result
End Function

Public Sub sendProcessDebugMappings(ByVal val As Boolean)
    Dim w As MessageWriter
    If NETisReplay() Then
        Exit Sub
    End If
    If selectedPlayer >= MAX_PLAYERS Then
        Exit Sub
    End If
    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_DEBUG_MODE)
    NETbool w, val
    NETend w
End Sub

' static std::string getWantedDebugMappingStatuses(const DebugInputManager& dbgInputManager, bool bStatus):
' the positions of the players whose wish is bStatus, as digits, sorted
Private Function getWantedDebugMappingStatuses(ByVal bStatus As Boolean) As String
    Dim ret() As Long, p As Long, n As Long, i As Long, j As Long, t As Long, s As String
    ReDim ret(0 To MAX_PLAYERS)             ' char ret[MAX_PLAYERS + 1] = "\0";
    p = 0
    For n = 0 To MAX_PLAYERS - 1
        If NetPlay_players_allocated(n) And Not NetPlay_players_isSpectator(n) And (DebugInputManager_getPlayerWantsDebugMappings(n) = bStatus) Then
            ret(p) = AscW("0") + NetPlay_players_position(n)    ' *p++ = '0' + NetPlay.players[n].position;
            p = p + 1
        End If
    Next
    ' std::sort(ret, p);
    For i = 1 To p - 1
        t = ret(i)
        j = i - 1
        Do While j >= 0
            If ret(j) <= t Then Exit Do
            ret(j + 1) = ret(j)
            j = j - 1
        Loop
        ret(j + 1) = t
    Next
    For i = 0 To p - 1
        s = s & ChrW$(ret(i) And &HFF&)     ' char
    Next
    getWantedDebugMappingStatuses = s
End Function

Public Sub recvProcessDebugMappings(ByVal queue As Long)
    Dim val As Boolean, r As MessageReader, oldDebugMode As Boolean, newDebugMode As Boolean, cmsg As String
    val = False
    Set r = NETbeginDecode(queue, GAME_DEBUG_MODE)
    NETbool r, val
    NETend r

    If getLockedOptions().cheats And val Then
        addConsoleMessage "The host has disabled debug mode / cheats for this game.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        Exit Sub
    End If

    ' (DebugInputManager& dbgInputManager = gInputManager.debugManager(): input_debugmappings_c's module)
    oldDebugMode = DebugInputManager_debugMappingsAllowed()
    DebugInputManager_setPlayerWantsDebugMappings NETQUEUE_index(queue), val
    newDebugMode = DebugInputManager_debugMappingsAllowed()

    If val Then
        cmsg = va("%s wants to enable debug mode. Enabled: %s, Disabled: %s.", _
            getPlayerName(NETQUEUE_index(queue)), _
            getWantedDebugMappingStatuses(True), _
            getWantedDebugMappingStatuses(False))
    Else
        cmsg = va("%s wants to disable debug mode. Enabled: %s, Disabled: %s.", _
            getPlayerName(NETQUEUE_index(queue)), _
            getWantedDebugMappingStatuses(True), _
            getWantedDebugMappingStatuses(False))
    End If
    addConsoleMessage cmsg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE

    If Not oldDebugMode And newDebugMode Then
        addConsoleMessage "Debug mode now enabled!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        Cheated = True
        gInputManager_contexts_set InputContext_DEBUG_MISC, InputContext_State_ACTIVE
        triggerEventCheatMode True
    ElseIf oldDebugMode And Not newDebugMode Then
        addConsoleMessage "Debug mode now disabled!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        If Not NETisReplay() Then
            gInputManager_contexts_set InputContext_DEBUG_MISC, InputContext_State_INACTIVE
        End If
        triggerEventCheatMode False
    End If
End Sub
