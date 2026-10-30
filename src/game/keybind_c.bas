Attribute VB_Name = "keybind_c"
' src/keybind.cpp, src/keybind.h (whole): all the functions that can be mapped to a key (kf_*), and
' enableGodMode, which makePlayerSpectator calls for the local spectator (it sees the whole map).
' Also names src/template.cpp (kf_CloneSelected's enumerateTemplates).
' (VB) A MappableFunction (std::function<void()>) is a WzFunc: a plain kf_ function the key table names is
' FN_kf_<name> (10000 up), a factory's lambda (kf_Zoom(multiplier), ...) FN_kf_<factory> (10100 up) with its
' captures; keybind_c_fn runs them. SPECTATOR_NO_OP() is a function: True when the caller returns. A
' Rational is its numerator and denominator (gtime_c's gameTimeGetMod / gameTimeSetMod). The #ifndef DEBUG
' parts are kept (a release build). _("...") is the English text; ngettext a Private copy. keybind_c_init is
' the globals' C++ initialisers (Sub Main runs it).
Option Explicit

Public Const MAP_ZOOM_RATE_MAX As Long = 1000
Public Const MAP_ZOOM_RATE_MIN As Long = 200
Public Const MAP_ZOOM_RATE_DEFAULT As Long = 500
Public Const MAP_ZOOM_RATE_STEP As Long = 50

Public Const SPIN_SCALING As Long = 360& * 182       ' (360*DEG_1): DEG_1 is (65536 / 360) = 182 in C
Public Const SECS_PER_SPIN As Long = 2
Public Const MAP_SPIN_RATE As Long = SPIN_SCALING \ SECS_PER_SPIN
Public Const MAP_PITCH_RATE As Long = SPIN_SCALING \ SECS_PER_SPIN

Private Const QUICKSAVE_CAM_FOLDER As String = "savegames/campaign/QuickSave"
Private Const QUICKSAVE_SKI_FOLDER As String = "savegames/skirmish/QuickSave"
Private Const QUICKSAVE_CAM_FILENAME As String = "savegames/campaign/QuickSave.gam"
Private Const QUICKSAVE_SKI_FILENAME As String = "savegames/skirmish/QuickSave.gam"
Private Const QUICKSAVE_CAM_JSON_FILENAME As String = QUICKSAVE_CAM_FOLDER & "/gam.json"
Private Const QUICKSAVE_SKI_JSON_FILENAME As String = QUICKSAVE_SKI_FOLDER & "/gam.json"

' (VB) __TIME__: VB has no compile time
Private Const COMPILE_TIME As String = "00:00:00"

' the plain key functions as MappableFunctions (keyconfig.cpp's table names them)
Public Const FN_kf_ChooseManufacture As Long = 10000
Public Const FN_kf_ChooseResearch As Long = 10001
Public Const FN_kf_ChooseBuild As Long = 10002
Public Const FN_kf_ChooseDesign As Long = 10003
Public Const FN_kf_ChooseIntelligence As Long = 10004
Public Const FN_kf_ChooseCommand As Long = 10005
Public Const FN_kf_QuickSave As Long = 10006
Public Const FN_kf_ToggleRadar As Long = 10007
Public Const FN_kf_QuickLoad As Long = 10008
Public Const FN_kf_ToggleConsole As Long = 10009
Public Const FN_kf_ToggleEnergyBars As Long = 10010
Public Const FN_kf_ScreenDump As Long = 10011
Public Const FN_kf_ToggleFormationSpeedLimiting As Long = 10012
Public Const FN_kf_MoveToLastMessagePos As Long = 10013
Public Const FN_kf_ToggleSensorDisplay As Long = 10014
Public Const FN_kf_addMultiMenu As Long = 10015
Public Const FN_kf_SeekNorth As Long = 10016
Public Const FN_kf_ToggleCamera As Long = 10017
Public Const FN_kf_addInGameOptions As Long = 10018
Public Const FN_kf_PitchForward As Long = 10019
Public Const FN_kf_RotateLeft As Long = 10020
Public Const FN_kf_ResetPitch As Long = 10021
Public Const FN_kf_RotateRight As Long = 10022
Public Const FN_kf_PitchBack As Long = 10023
Public Const FN_kf_RightOrderMenu As Long = 10024
Public Const FN_kf_SlowDown As Long = 10025
Public Const FN_kf_SpeedUp As Long = 10026
Public Const FN_kf_NormalSpeed As Long = 10027
Public Const FN_kf_FaceNorth As Long = 10028
Public Const FN_kf_FaceSouth As Long = 10029
Public Const FN_kf_FaceEast As Long = 10030
Public Const FN_kf_FaceWest As Long = 10031
Public Const FN_kf_JumpToResourceExtractor As Long = 10032
Public Const FN_kf_ToggleOverlays As Long = 10033
Public Const FN_kf_ToggleConsoleDrop As Long = 10034
Public Const FN_kf_ToggleTeamChat As Long = 10035
Public Const FN_kf_RotateBuildingCW As Long = 10036
Public Const FN_kf_RotateBuildingACW As Long = 10037
Public Const FN_kf_CentreOnBase As Long = 10038
Public Const FN_kf_JumpToUnassignedUnits As Long = 10039
Public Const FN_kf_SendGlobalMessage As Long = 10040
Public Const FN_kf_SendTeamMessage As Long = 10041
Public Const FN_kf_SendGlobalQuickChat As Long = 10042
Public Const FN_kf_SendTeamQuickChat As Long = 10043
Public Const FN_kf_AddHelpBlip As Long = 10044
Public Const FN_kf_ToggleShadows As Long = 10045
Public Const FN_kf_toggleTrapCursor As Long = 10046
Public Const FN_kf_ToggleRadarTerrain As Long = 10047
Public Const FN_kf_ToggleRadarAllyEnemy As Long = 10048
Public Const FN_kf_ShowMappings As Long = 10049
Public Const FN_kf_FrameRate As Long = 10050
Public Const FN_kf_ToggleDebugMappings As Long = 10051
Public Const FN_kf_PrioritizeDebugMappings As Long = 10052
Public Const FN_kf_ToggleLevelEditor As Long = 10053
Public Const FN_kf_ToggleShowPath As Long = 10054
Public Const FN_kf_ToggleShowGateways As Long = 10055
Public Const FN_kf_ToggleVisibility As Long = 10056
Public Const FN_kf_RaiseTile As Long = 10057
Public Const FN_kf_LowerTile As Long = 10058
Public Const FN_kf_ToggleFog As Long = 10059
Public Const FN_kf_ToggleWeather As Long = 10060
Public Const FN_kf_TriFlip As Long = 10061
Public Const FN_kf_PerformanceSample As Long = 10062
Public Const FN_kf_AllAvailable As Long = 10063
Public Const FN_kf_KillSelected As Long = 10064
Public Const FN_kf_ToggleGodMode As Long = 10065
Public Const FN_kf_ChooseOptions As Long = 10066
Public Const FN_kf_FinishResearch As Long = 10067
Public Const FN_kf_RevealMapAtPos As Long = 10068
Public Const FN_kf_TraceObject As Long = 10069
Public Const FN_kf_ToggleFullscreen As Long = 10070

' the factories' lambdas
Public Const FN_kf_Zoom As Long = 10100                 ' kf_Zoom: [multiplier]()
Public Const FN_kf_RadarZoom As Long = 10101            ' kf_RadarZoom: [multiplier]()
Public Const FN_kf_SelectGrouping_N As Long = 10102     ' [n]()
Public Const FN_kf_AssignGrouping_N As Long = 10103     ' [n]()
Public Const FN_kf_AddGrouping_N As Long = 10104        ' [n]()
Public Const FN_kf_RemoveFromGrouping As Long = 10105   ' []()
Public Const FN_kf_SelectCommander_N As Long = 10106    ' [n]()
Public Const FN_kf_JumpToMapMarker As Long = 10107      ' [x, z, yaw]()
Public Const FN_kf_ScrollCamera As Long = 10108         ' [horizontal, vertical]()
Public Const FN_kf_ShowMappings_onClose As Long = 10109 ' kf_ShowMappings's showInGameOptionsScreen onCloseFunc: []()
Public Const FN_kf_JumpToUnits As Long = 10110          ' [droidType]()
Public Const FN_kf_SelectNextFactory As Long = 10111    ' [factoryType, bJumpToSelected]()
Public Const FN_kf_SelectNextResearch As Long = 10112   ' [bJumpToSelected]()
Public Const FN_kf_SelectNextPowerStation As Long = 10113 ' [bJumpToSelected]()
Public Const FN_kf_SelectUnits As Long = 10114          ' [selectionClass, selectionType, bOnScreen]()
Public Const FN_kf_SelectNoGroupUnits As Long = 10115   ' [selectionClass, selectionType, bOnScreen]()
Public Const FN_kf_SetDroid As Long = 10116             ' [order, state]()
Public Const FN_kf_OrderDroid As Long = 10117           ' [order]()

Public bMovePause As Boolean                    ' = false
Public bAllowOtherKeyPresses As Boolean         ' = true
Private beaconMsg() As String                   ' char beaconMsg[MAX_PLAYERS][MAX_CONSOLE_STRING_LENGTH]: beacon msg for each player (VB: Private, only keybind.cpp uses it; multiplay_c's parameter of that name)

' using ExtractorListIter = optional<ExtractorList::const_iterator>; static ExtractorListIter psOldRE = {};
' (an apsExtractorLists node, 0 = end(); psOldRE_has: has_value)
Private psOldRE As Long
Private psOldRE_has As Boolean
Private sCurrentConsoleText As String           ' char[MAX_CONSOLE_STRING_LENGTH]: remember what user types in console for beacon msg

' kf_ToggleGodMode's static bool pastReveal = true;
Private kf_ToggleGodMode_pastReveal As Boolean

Private oldTrapCursorEnabledValue As Long       ' static TrapCursorMode = TrapCursorMode::Enabled

' static const Rational available_speed[]; const unsigned nb_available_speeds = ARRAY_SIZE(available_speed);
Private Const nb_available_speeds As Long = 20
Private available_speed_n(0 To nb_available_speeds - 1) As Long
Private available_speed_d(0 To nb_available_speeds - 1) As Long

' (VB) the C++ initialisers of keybind.cpp's globals and statics
Public Sub keybind_c_init()
    bMovePause = False
    bAllowOtherKeyPresses = True
    psOldRE_has = False
    psOldRE = 0
    kf_ToggleGodMode_pastReveal = True
    oldTrapCursorEnabledValue = TrapCursorMode_Enabled
    ReDim beaconMsg(0 To MAX_PLAYERS - 1)
    available_speed_init
End Sub

' (VB) lib/framework/gettext.h's ngettext without translations: the singular for 1, else the plural
Private Function ngettext(ByVal msgid1 As String, ByVal msgid2 As String, ByVal n As Long) As String
    If n = 1 Then ngettext = msgid1 Else ngettext = msgid2
End Function

' #define SPECTATOR_NO_OP() do { if (selectedPlayer >= MAX_PLAYERS || NetPlay.players[selectedPlayer].isSpectator) { return; } } while (0):
' True when the caller returns
Private Function SPECTATOR_NO_OP() As Boolean
    SPECTATOR_NO_OP = True
    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then Exit Function
    If NetPlay_players_isSpectator(selectedPlayer) Then Exit Function
    SPECTATOR_NO_OP = False
End Function

' A function to determine whether we're running a multiplayer game, not just a
' single player campaign or a skirmish game.
' \return false if this is a skirmish or single player game, true if it is a
'         multiplayer game.
Public Function runningMultiplayer() As Boolean
    runningMultiplayer = bMultiPlayer And NetPlay_bComms
End Function

Private Sub noMPCheatMsg()
    addConsoleMessage "Sorry, that cheat is disabled in multiplayer games.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
End Sub

Public Function shouldTrapCursor() As Boolean
    Dim result As Long
    result = war_GetTrapCursor()
    Select Case result
    Case TrapCursorMode_Disabled
        shouldTrapCursor = False
        Exit Function
    Case TrapCursorMode_Enabled
        shouldTrapCursor = True
        Exit Function
    Case TrapCursorMode_Automatic
        ' automatic mode - if in fullscreen mode (including desktop full), or maximized, trap the cursor
        shouldTrapCursor = (wzIsFullscreen() Or wzIsMaximized())
        Exit Function
    End Select
    shouldTrapCursor = False ' silence compiler warning
End Function

' --------------------------------------------------------------------------
Public Sub kf_AutoGame()
    Dim i As Long
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If game.typ = LEVEL_TYPE_CAMPAIGN Then
        CONPRINTF "%s", "Not possible with the campaign!"
        Exit Sub
    End If
    ' Notify all human players that we are trying to enable autogame
    For i = 0 To MAX_PLAYERS - 1
        If NetPlay_players_allocated(i) Then
            sendGift AUTOGAME_GIFT, i
        End If
    Next

    CONPRINTF "%s", "autogame request has been sent to all players. AI script *must* support this command!"
End Sub

Public Sub kf_ToggleMissionTimer()
    addConsoleMessage "Warning! This cheat is buggy.  We recommend to NOT use it.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    setMissionCheatTime (mission_cheatTime = 0)
End Sub

Public Sub kf_ToggleShowGateways()
    addConsoleMessage "Gateways toggled.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    showGateways = Not showGateways
End Sub

Public Sub kf_ToggleShowPath()
    addConsoleMessage "Path display toggled.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    showPath = Not showPath
End Sub

Public Sub kf_PerformanceSample()
    wzPerfStart
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleRadarJump()
    war_SetRadarJump Not war_GetRadarJump()
End Sub

' --------------------------------------------------------------------------

Public Sub kf_ForceDesync()
    syncDebug "kf_ForceDesync", "Oh no!!! I went out of sync!!!"
End Sub

Public Sub kf_PowerInfo()
    Dim i As Long
    For i = 0 To game.maxPlayers - 1
        console "Player %d: %d power", i, getPower(i)
    Next
End Sub

Public Sub kf_DamageMe()
    Dim it As Long, psDroid As Long, psStruct As Long, val_ As Long
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If
    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then
        Exit Sub ' no-op
    End If
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected Then
            val_ = aDroid(psDroid).body - ((aDroid(psDroid).originalBody \ 100) * 20)
            If val_ > 0 Then
                aDroid(psDroid).body = val_
                addConsoleMessage "Ouch! Droid's health is down 20%!", LEFT_JUSTIFY, SYSTEM_MESSAGE
            Else
                aDroid(psDroid).body = 0
            End If
        End If
        it = lnNext(it)
    Loop
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        psStruct = lnObj(it)
        If aStruct(psStruct).selected Then
            val_ = aStruct(psStruct).body - ((STRUCTURE_structureBody(psStruct) \ 100) * 20)
            If val_ > 0 Then
                aStruct(psStruct).body = val_
                addConsoleMessage "Ouch! Structure's health is down 20%!", LEFT_JUSTIFY, SYSTEM_MESSAGE
            Else
                aStruct(psStruct).body = 0
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

Public Sub kf_TraceObject()
    Dim it As Long, psCDroid As Long, psCStruct As Long
    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then
        Exit Sub ' no-op
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psCDroid = lnObj(it)
        If aDroid(psCDroid).selected Then
            objTraceEnable aDroid(psCDroid).id
            CONPRINTF "Tracing droid %d", aDroid(psCDroid).id
            Exit Sub
        End If
        it = lnNext(it)
    Loop
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        psCStruct = lnObj(it)
        If aStruct(psCStruct).selected Then
            objTraceEnable aStruct(psCStruct).id
            CONPRINTF "Tracing structure %d", aStruct(psCStruct).id
            Exit Sub
        End If
        it = lnNext(it)
    Loop
    objTraceDisable
    CONPRINTF "%s", "No longer tracing anything."
End Sub

'===================================================
Public Sub kf_ToggleSensorDisplay()
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    rangeOnScreen = Not rangeOnScreen

    If rangeOnScreen Then
        addConsoleMessage "Lets us see what you see!", LEFT_JUSTIFY, SYSTEM_MESSAGE    'added this message... Yeah, its lame. :)
    Else
        addConsoleMessage "Fine, weapon & sensor display is off!", LEFT_JUSTIFY, SYSTEM_MESSAGE    'added this message... Yeah, its lame. :)
    End If
End Sub
'===================================================
' Halves all the heights of the map tiles
Public Sub kf_HalveHeights()
    Dim psTile As Long, i As Long, j As Long
    For j = 0 To mapHeight - 1
        For i = 0 To mapWidth - 1
            psTile = mapTile(i, j)
            psMapTiles(psTile).height = psMapTiles(psTile).height \ 2
        Next
    Next
End Sub

' --------------------------------------------------------------------------
Public Sub kf_FaceNorth()
    kf_SeekNorth
End Sub
' --------------------------------------------------------------------------
Public Sub kf_FaceSouth()
    playerPos.r.y = DEG(180)
    If getWarCamStatus() Then
        camToggleStatus
    End If
    CONPRINTF "%s", "View Aligned to South"
End Sub
' --------------------------------------------------------------------------
Public Sub kf_FaceEast()
    playerPos.r.y = DEG(90)
    If getWarCamStatus() Then
        camToggleStatus
    End If
    CONPRINTF "%s", "View Aligned to East"
End Sub
' --------------------------------------------------------------------------
Public Sub kf_FaceWest()
    playerPos.r.y = DEG(270)
    If getWarCamStatus() Then
        camToggleStatus
    End If
    CONPRINTF "%s", "View Aligned to West"
End Sub
' --------------------------------------------------------------------------

' Writes out debug info about all the selected droids
Public Sub kf_DebugDroidInfo()
    Dim it As Long, psDroid As Long
    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then
        Exit Sub ' no-op
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected Then
            printDroidInfo psDroid
        End If
        it = lnNext(it)
    Loop
End Sub

Public Sub kf_CloneSelected(ByVal limit As Long)
    Dim sTemplate As Long, selectedAnything As Boolean, droidToClone As Long, it As Long, d As Long, t As Long, psTempl As Long
    Dim i As Long, pos As Vector2i, psNewDroid As Long, msg As String, noOrders As INITIAL_DROID_ORDERS, rot As Rotation
    sTemplate = 0
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then
        Exit Sub ' no-op
    End If

    selectedAnything = False
    droidToClone = 0

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        d = lnObj(it)
        If Not aDroid(d).selected Then
            GoTo continue_
        End If
        selectedAnything = True
        ' enumerateTemplates(selectedPlayer, [psDroid = d, &sTemplate](DROID_TEMPLATE* psTempl) { ... })
        If Not (selectedPlayer >= 0 And selectedPlayer < MAX_PLAYERS) Then
            ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:659", "enumerateTemplates", "Invalid player: %u", selectedPlayer
        Else
            For t = 0 To templateCount(selectedPlayer) - 1
                psTempl = droidTemplates_at(selectedPlayer, t)
                If apsTemplates(psTempl).name_ = aDroid(d).aName Then
                    sTemplate = psTempl
                    Exit For ' stop enumerating
                End If
            Next
        End If
        If sTemplate = 0 Then
            debug_ 429, LOG_ERROR, "kf_CloneSelected", "Cloning vat has been destroyed. We can't find the template for this droid: %s, id:%u, type:%d!", aDroid(d).aName, aDroid(d).id, aDroid(d).droidType
            Exit Do
        End If
        droidToClone = d
        ' Break out of the loop if we've successfully found the associated droid template.
        Exit Do
continue_:
        it = lnNext(it)
    Loop

    If Not selectedAnything Then
        debug_ 438, LOG_INFO, "kf_CloneSelected", "Nothing was selected?"
        Exit Sub
    End If
    If sTemplate = 0 Then
        Exit Sub
    End If
    If Not (droidToClone <> 0) Then ASSERT_FAILURE "droidToClone != nullptr", "src/keybind.cpp:445", "kf_CloneSelected", "No droid to clone found"

    ' create a new droid army
    For i = 0 To limit - 1
        ' 40503 = 65536/phi (A bit more than a right angle)
        pos.x = aDroid(droidToClone).pos.x + iSinR((40503 * i) And &HFFFF&, iSqrt(50& * 50 * (i + 1)))
        pos.y = aDroid(droidToClone).pos.y + iCosR((40503 * i) And &HFFFF&, iSqrt(50& * 50 * (i + 1)))
        psNewDroid = buildDroid(sTemplate, pos.x, pos.y, aDroid(droidToClone).player, False, False, noOrders, rot)
        If psNewDroid <> 0 Then
            addDroid psNewDroid, apsDroidLists
            triggerEventDroidBuilt psNewDroid, 0
        ElseIf Not bMultiMessages Then
            debug_ 458, LOG_ERROR, "kf_CloneSelected", "Cloning has failed for template:%s id:%d", apsTemplates(sTemplate).id, apsTemplates(sTemplate).multiPlayerID
        End If
    Next

    msg = va("Player %u is cheating a new droid army of: %d " & ChrW$(215) & " %s.", selectedPlayer, limit, aDroid(droidToClone).aName)
    sendInGameSystemMessage msg
    Cheated = True
    audio_PlayTrack ID_SOUND_NEXUS_LAUGH1
End Sub

Public Sub kf_MakeMeHero()
    Dim it As Long, psDroid As Long
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then
        Exit Sub ' no-op
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected And (aDroid(psDroid).droidType = DROID_COMMAND Or aDroid(psDroid).droidType = DROID_SENSOR) Then
            aDroid(psDroid).experience = 16& * 65536 * 128
        ElseIf aDroid(psDroid).selected Then
            aDroid(psDroid).experience = 4& * 65536 * 128
        End If
        it = lnNext(it)
    Loop
End Sub

Public Sub kf_TeachSelected()
    Dim it As Long, psDroid As Long
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then
        Exit Sub ' no-op
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected Then
            aDroid(psDroid).experience = aDroid(psDroid).experience + 4& * 65536
        End If
        it = lnNext(it)
    Loop
End Sub

Public Sub kf_Unselectable()
    Dim it As Long, psDroid As Long, psStruct As Long
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then
        Exit Sub ' no-op
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected Then
            aDroid(psDroid).flags = aDroid(psDroid).flags Or BIT_(OBJECT_FLAG_UNSELECTABLE)
            aDroid(psDroid).selected = False
        End If
        it = lnNext(it)
    Loop
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        psStruct = lnObj(it)
        If aStruct(psStruct).selected Then
            aStruct(psStruct).flags = aStruct(psStruct).flags Or BIT_(OBJECT_FLAG_UNSELECTABLE)
            aStruct(psStruct).selected = False
        End If
        it = lnNext(it)
    Loop
End Sub

' --------------------------------------------------------------------------
'
'/* Prints out the date and time of the build of the game */
Public Sub kf_BuildInfo()
    CONPRINTF "Built: %s %s", getCompileDate(), COMPILE_TIME
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleConsoleDrop()
    If Not bInTutorial Then
        setHistoryMode False
        toggleConsoleDrop
    End If
End Sub

Public Sub kf_ToggleTeamChat()
    If Not bInTutorial Then
        setHistoryMode True
        toggleConsoleDrop
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_BifferBaker()
    Dim cmsg As String
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    ' player deals far more damage, and the enemy far less
    setDamageModifiers 999, 1
    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, "Hard as nails!!!")
    sendInGameSystemMessage cmsg
End Sub
' --------------------------------------------------------------------------
Public Sub kf_UpThePower()
    Dim cmsg As String
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    addPower selectedPlayer, 1000
    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, "1000 big ones!!!")
    sendInGameSystemMessage cmsg
End Sub

' --------------------------------------------------------------------------
Public Sub kf_MaxPower()
    Dim cmsg As String
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    setPower selectedPlayer, 100000
    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, "Power overwhelming")
    sendInGameSystemMessage cmsg
End Sub

' --------------------------------------------------------------------------
Public Sub kf_SetDifficultyLevel(ByVal level As DIFFICULTY_LEVEL)
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    setDifficultyLevel level
    Select Case level
    Case DL_SUPER_EASY: addConsoleMessage "A power fantasy!", LEFT_JUSTIFY, SYSTEM_MESSAGE
    Case DL_EASY: addConsoleMessage "Taking things easy!", LEFT_JUSTIFY, SYSTEM_MESSAGE
    Case DL_NORMAL: addConsoleMessage "Back to normality!", LEFT_JUSTIFY, SYSTEM_MESSAGE
    Case DL_HARD: addConsoleMessage "Getting tricky!", LEFT_JUSTIFY, SYSTEM_MESSAGE
    Case DL_INSANE: addConsoleMessage "In a nightmare!", LEFT_JUSTIFY, SYSTEM_MESSAGE
    Case Else
    End Select
End Sub

' --------------------------------------------------------------------------
Public Sub kf_DoubleUp()
    Dim cmsg As String
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If
    setDamageModifiers 100, 50 ' enemy damage halved
    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, "Twice as nice!")
    sendInGameSystemMessage cmsg
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleFPS() 'This shows *just FPS* and is always visible (when active) -Q.
    ' Toggle the boolean value of showFPS
    showFPS = Not showFPS

    If showFPS Then
        CONPRINTF "%s", "FPS display is enabled."
    Else
        CONPRINTF "%s", "FPS display is disabled."
    End If
End Sub

Public Sub kf_ToggleUnitCount()     ' Display units built / lost / produced counter
    ' Toggle the boolean value of showUNITCOUNT
    showUNITCOUNT = Not showUNITCOUNT

    If showUNITCOUNT Then
        CONPRINTF "%s", "Unit count display is enabled."
    Else
        CONPRINTF "%s", "Unit count display is disabled."
    End If
End Sub

Public Sub kf_ToggleSamples() 'Displays number of sound sample in the sound queues & lists.
    ' Toggle the boolean value of showSAMPLES
    showSAMPLES = Not showSAMPLES

    CONPRINTF "Sound Samples displayed is %s", IIf(showSAMPLES, "Enabled", "Disabled")
End Sub

Public Sub kf_ToggleOrders()  ' Displays orders & action of currently selected unit.
    ' Toggle the boolean value of showORDERS
    showORDERS = Not showORDERS
    CONPRINTF "Unit Order/Action displayed is %s", IIf(showORDERS, "Enabled", "Disabled")
End Sub

' Writes out the frame rate
Public Sub kf_FrameRate()
    CONPRINTF "FPS %d; PIEs %zu; polys %zu", _
              frameRate(), loopPieCount, loopPolyCount
    If runningMultiplayer() Then
        CONPRINTF "NETWORK:  Bytes: s-%zu r-%zu  Uncompressed Bytes: s-%zu r-%zu  Packets: s-%zu r-%zu", _
                  NETgetStatistic(NetStatisticRawBytes, True), _
                  NETgetStatistic(NetStatisticRawBytes, False), _
                  NETgetStatistic(NetStatisticUncompressedBytes, True), _
                  NETgetStatistic(NetStatisticUncompressedBytes, False), _
                  NETgetStatistic(NetStatisticPackets, True), _
                  NETgetStatistic(NetStatisticPackets, False)
    End If
    gameStats = Not gameStats
    CONPRINTF "Built: %s %s", getCompileDate(), COMPILE_TIME
End Sub

' --------------------------------------------------------------------------

' display the total number of objects in the world
Public Sub kf_ShowNumObjects()
    Dim droids As Long, structures As Long, features As Long, cmsg As String
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    objCount droids, structures, features
    cmsg = va("(Player %u) is using a cheat :Num Droids: %d  Num Structures: %d  Num Features: %d", _
              selectedPlayer, droids, structures, features)
    sendInGameSystemMessage cmsg
End Sub

Public Sub kf_ListDroids()
    Dim i As Long, it As Long, psDroid As Long, x As Long, y As Long
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    debug_ 749, LOG_INFO, "kf_ListDroids", "list droids:"
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If
    For i = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            x = map_coord(aDroid(psDroid).pos.x)
            y = map_coord(aDroid(psDroid).pos.y)
            debug_ 761, LOG_INFO, "kf_ListDroids", "droid %i;%s;%i;%i;%i", i, aDroid(psDroid).aName, aDroid(psDroid).droidType, x, y
            it = lnNext(it)
        Loop
    Next
End Sub

' --------------------------------------------------------------------------
' Toggles radar on off
Public Sub kf_ToggleRadar()
    radarOnScreen = Not radarOnScreen
End Sub

' --------------------------------------------------------------------------
' Toggles infinite power on/off
Public Sub kf_TogglePower()
    Dim cmsg As String
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    powerCalculated = Not powerCalculated
    If powerCalculated Then
        powerCalc True
    End If

    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, IIf(powerCalculated, "Infinite power disabled", "Infinite power enabled"))
    sendInGameSystemMessage cmsg
End Sub

' --------------------------------------------------------------------------
' Recalculates the lighting values for a tile
Public Sub kf_RecalcLighting()
    initLighting 0, 0, mapWidth, mapHeight
    addConsoleMessage "Lighting values for all tiles recalculated", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
End Sub

' --------------------------------------------------------------------------
' Sends the screen buffer to disk
Public Sub kf_ScreenDump()
    screenDumpToDisk ScreenDumpPath, getLevelName()
End Sub

' --------------------------------------------------------------------------
' Make all functions available
Public Sub kf_AllAvailable()
    Dim cmsg As String
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    makeAllAvailable
    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, "All items made available")
    sendInGameSystemMessage cmsg
End Sub

' --------------------------------------------------------------------------
' Flips the cut of a tile
Public Sub kf_TriFlip()
    Dim psTile As Long
    psTile = mapTile(mouseTileX, mouseTileY)
    TOGGLE_TRIFLIP psTile
'   addConsoleMessage(_("Triangle flip status toggled"),DEFAULT_JUSTIFY,SYSTEM_MESSAGE);
End Sub

' --------------------------------------------------------------------------
' Debug info about a map tile
Public Sub kf_TileInfo()
    Dim psTile As Long
    psTile = mapTile(mouseTileX, mouseTileY)

    debug_ 861, LOG_ERROR, "kf_TileInfo", "Tile position=(%d, %d) Terrain=%d Texture=%u Height=%d Illumination=%u", _
          mouseTileX, mouseTileY, terrainType(psTile), TileNumber_tile(psMapTiles(psTile).texture), psMapTiles(psTile).height, _
          psMapTiles(psTile).illumination
    addConsoleMessage "Tile info dumped into log", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
End Sub

' Toggles fog on/off
Public Sub kf_ToggleFog()
    Dim cmsg As String
    If pie_GetFogEnabled() Then
        pie_SetFogStatus False
        pie_EnableFog False
    Else
        pie_EnableFog True
    End If

    If pie_GetFogEnabled() Then cmsg = "Fog on" Else cmsg = "Fog off"
    sendInGameSystemMessage cmsg
End Sub

' --------------------------------------------------------------------------

' Toggle camera on/off
Public Sub kf_ToggleCamera()
    If Not getWarCamStatus() Then
        shakeStop ' Ensure screen shake stopped before starting camera mode.
    End If
    camToggleStatus
End Sub

Public Sub kf_RevealMapAtPos()
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then Exit Sub

    addSpotter mouseTileX, mouseTileY, selectedPlayer, 1024, False, gameTime + 2000
End Sub

' --------------------------------------------------------------------------
Public Sub kf_MapCheck()
    Dim it As Long, psDroid As Long
    ' Bail out if we're running a _true_ multiplayer game (to prevent desyncs)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then Exit Sub

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        aDroid(psDroid).pos.z = map_Height(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)
        it = lnNext(it)
    Loop

    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        alignStructure lnObj(it)
        it = lnNext(it)
    Loop

    it = apsFlagPosLists(selectedPlayer).head
    Do While it <> 0
        With aFlagPos(lnObj(it))
            .coords.z = map_Height(.coords.x, .coords.y) + ASSEMBLY_POINT_Z_PADDING
        End With
        it = lnNext(it)
    Loop
End Sub

' Raises the tile under the mouse
Public Sub kf_RaiseTile()
    If runningMultiplayer() Then
        Exit Sub  ' Don't desynch if pressing 'W'...
    End If
    raiseTile mouseTileX, mouseTileY
End Sub

' --------------------------------------------------------------------------
' Lowers the tile under the mouse
Public Sub kf_LowerTile()
    If runningMultiplayer() Then
        Exit Sub  ' Don't desynch if pressing 'A'...
    End If
    lowerTile mouseTileX, mouseTileY
End Sub

' --------------------------------------------------------------------------
' Zooms in/out from display
Public Function kf_Zoom(ByVal multiplier As Long) As WzFunc
    Set kf_Zoom = WzFunc_new(FN_kf_Zoom, multiplier)
End Function

' Zooms in/out from radar
Public Function kf_RadarZoom(ByVal multiplier As Long) As WzFunc
    Set kf_RadarZoom = WzFunc_new(FN_kf_RadarZoom, multiplier)
End Function

' kf_RadarZoom's [multiplier]() { ... } (uint8_t arithmetic: the sum wraps at 256)
Private Sub kf_RadarZoom_lambda(ByVal multiplier As Long)
    Dim oldZoomLevel As Long, newZoomLevel As Long
    oldZoomLevel = GetRadarZoom() And &HFF&
    newZoomLevel = (oldZoomLevel + (RADARZOOM_STEP * multiplier)) And &HFF&
    newZoomLevel = IIf(newZoomLevel > MAX_RADARZOOM, MAX_RADARZOOM, newZoomLevel)
    newZoomLevel = IIf(newZoomLevel < MIN_RADARZOOM, MIN_RADARZOOM, newZoomLevel)

    If newZoomLevel <> oldZoomLevel Then
        CONPRINTF "Setting radar zoom to %u", newZoomLevel
        SetRadarZoom newZoomLevel
        war_SetRadarZoom GetRadarZoom() ' persist changed setting to config
        audio_PlayTrack ID_SOUND_BUTTON_CLICK_5
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_MaxScrollLimits()
    scrollMinX = 0: scrollMinY = 0
    scrollMaxX = mapWidth
    scrollMaxY = mapHeight
End Sub

' --------------------------------------------------------------------------
' Spins the world round left
Public Sub kf_RotateLeft()
    Dim rotAmount As Long
    rotAmount = ToInt(realTimeAdjustedIncrement(MAP_SPIN_RATE))

    playerPos.r.y = playerPos.r.y + rotAmount
End Sub

' --------------------------------------------------------------------------
' Spins the world right
Public Sub kf_RotateRight()
    Dim rotAmount As Long
    rotAmount = ToInt(realTimeAdjustedIncrement(MAP_SPIN_RATE))

    playerPos.r.y = playerPos.r.y - rotAmount
    If playerPos.r.y < 0 Then
        playerPos.r.y = playerPos.r.y + DEG(360)
    End If
End Sub

' --------------------------------------------------------------------------
' Rotate editing building direction clockwise
Public Sub kf_RotateBuildingCW()
    incrementBuildingDirection DEG(-90)
End Sub

' --------------------------------------------------------------------------
' Rotate editing building direction anticlockwise
Public Sub kf_RotateBuildingACW()
    incrementBuildingDirection DEG(90)
End Sub

' --------------------------------------------------------------------------
' Pitches camera back
Public Sub kf_PitchBack()
    Dim pitchAmount As Long
    pitchAmount = ToInt(realTimeAdjustedIncrement(MAP_PITCH_RATE))

    playerPos.r.x = playerPos.r.x + pitchAmount

    If playerPos.r.x > DEG(360 + MAX_PLAYER_X_ANGLE) Then
        playerPos.r.x = DEG(360 + MAX_PLAYER_X_ANGLE)
    End If
End Sub

' --------------------------------------------------------------------------
' Pitches camera forward
Public Sub kf_PitchForward()
    Dim pitchAmount As Long
    pitchAmount = ToInt(realTimeAdjustedIncrement(MAP_PITCH_RATE))

    playerPos.r.x = playerPos.r.x - pitchAmount
    If playerPos.r.x < DEG(360 + MIN_PLAYER_X_ANGLE) Then
        playerPos.r.x = DEG(360 + MIN_PLAYER_X_ANGLE)
    End If
End Sub

' --------------------------------------------------------------------------
' Resets pitch to default
Public Sub kf_ResetPitch()
    playerPos.r.x = DEG(360 + INITIAL_STARTING_PITCH)
    setViewDistance STARTDISTANCE
End Sub

' --------------------------------------------------------------------------
' Quickly access the in-game keymap
Public Sub kf_ShowMappings()
    Dim optionsBrowser As WIDGET
    If Not InGameOpUp And Not isInGamePopupUp Then
        ' Open new Options screen and jump to Controls section
        Set optionsBrowser = createOptionsBrowser(True)
        If Not OptionsBrowserForm_switchToOptionsForm(optionsBrowser, OptionsBrowserForm_Modes_Controls) Then
            debug_ 1078, LOG_INFO, "kf_ShowMappings", "Failed to open Controls options form?"
            Exit Sub
        End If
        showInGameOptionsScreen psWScreen, optionsBrowser, WzFunc_new(FN_kf_ShowMappings_onClose)
    End If
End Sub

' --------------------------------------------------------------------------
' Selects the player's groups 1..9
Public Sub kf_SelectGrouping(ByVal groupNumber As Long)
    Dim bAlreadySelected As Boolean, Selected As Boolean, it As Long, psDroid As Long
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    bAlreadySelected = False
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        ' Wipe out the ones in the wrong group
        If aDroid(psDroid).selected And aDroid(psDroid).group <> groupNumber Then
            aDroid(psDroid).selected = False
        End If
        ' Get the right ones
        If aDroid(psDroid).group = groupNumber Then
            If aDroid(psDroid).selected Then
                bAlreadySelected = True
            End If
        End If
        it = lnNext(it)
    Loop
    If bAlreadySelected Then
        Selected = activateGroupAndMove(selectedPlayer, groupNumber)
    Else
        Selected = activateGroup(selectedPlayer, groupNumber)
    End If

    intGroupsChanged groupNumber And &HFF&

    ' play group audio but only if they weren't already selected - AM
    If Selected And Not bAlreadySelected And war_getPlayAudioCue_GroupReporting() Then
        audio_QueueTrack ID_SOUND_GROUP_0 + groupNumber
        audio_QueueTrack ID_SOUND_REPORTING
        audio_QueueTrack ID_SOUND_RADIOCLICK_1 + (rand() Mod 6)
    End If
    triggerEventSelected
End Sub

' --------------------------------------------------------------------------
Public Function kf_SelectGrouping_N(ByVal n As Long) As WzFunc
    Set kf_SelectGrouping_N = WzFunc_new(FN_kf_SelectGrouping_N, n)
End Function

Public Function kf_AssignGrouping_N(ByVal n As Long) As WzFunc
    Set kf_AssignGrouping_N = WzFunc_new(FN_kf_AssignGrouping_N, n)
End Function

Public Function kf_AddGrouping_N(ByVal n As Long) As WzFunc
    Set kf_AddGrouping_N = WzFunc_new(FN_kf_AddGrouping_N, n)
End Function

Public Function kf_RemoveFromGrouping() As WzFunc
    Set kf_RemoveFromGrouping = WzFunc_new(FN_kf_RemoveFromGrouping)
End Function

Public Function kf_SelectCommander_N(ByVal n As Long) As WzFunc
    Set kf_SelectCommander_N = WzFunc_new(FN_kf_SelectCommander_N, n)
End Function

' --------------------------------------------------------------------------
Public Sub kf_ToggleDroidInfo()
    camToggleInfo
End Sub

Public Sub kf_addInGameOptions()
    setWidgetsStatus True
    If Not isInGamePopupUp Then ' they can *only* quit when popup is up.
        intResetScreen False
        intAddInGameOptions
    End If
End Sub

' --------------------------------------------------------------------------
' Display multiplayer guff.
Public Sub kf_addMultiMenu()
    If bMultiPlayer And (intMode <> INT_INTELMAP) Then
        intAddMultiMenu
    End If
End Sub

' --------------------------------------------------------------------------
Public Function kf_JumpToMapMarker(ByVal x As Long, ByVal z As Long, ByVal yaw As Long) As WzFunc
    Set kf_JumpToMapMarker = WzFunc_new(FN_kf_JumpToMapMarker, x, z, yaw)
End Function

' --------------------------------------------------------------------------
' Toggles the power bar display on and off
Public Sub kf_TogglePowerBar()
    togglePowerBar
End Sub
' --------------------------------------------------------------------------
' Toggles whether we process debug key mappings
Public Sub kf_ToggleDebugMappings()
    sendProcessDebugMappings Not DebugInputManager_getPlayerWantsDebugMappings(selectedPlayer)
End Sub

' Toggles the local debug mapping context prioritization status
Public Sub kf_PrioritizeDebugMappings()
    Dim status As String
    If DebugInputManager_toggleDebugMappingPriority() Then status = "TRUE" Else status = "FALSE"
    CONPRINTF "%s%s", "Toggling debug mapping priority: ", status
End Sub

Public Sub kf_ToggleLevelEditor()
    Dim bIsActive As Boolean
    bIsActive = gInputManager_contexts_isActive(InputContext_DEBUG_LEVEL_EDITOR)

    If bIsActive Then
        CONPRINTF "%s", "Disabling level editor"
        gInputManager_contexts_set InputContext_DEBUG_LEVEL_EDITOR, InputContext_State_INACTIVE
    Else
        CONPRINTF "%s", "Enabling level editor"
        gInputManager_contexts_set InputContext_DEBUG_LEVEL_EDITOR, InputContext_State_ACTIVE
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub enableGodMode()
    If godMode Then
        Exit Sub
    End If

    ' Bail out if we're running a _true_ multiplayer game and we aren't a spectator (to prevent MP cheating)
    If runningMultiplayer() And Not NetPlay_players_isSpectator(selectedPlayer) Then
        noMPCheatMsg
        Exit Sub
    End If

    godMode = True ' view all structures and droids
    revealAll selectedPlayer
    setRevealStatus True ' view the entire map
    radarPermitted = True 'add minimap without CC building

    preProcessVisibility
End Sub

Public Sub kf_ToggleGodMode()
    Dim it As Long, playerId As Long, cmsg As String
    ' static bool pastReveal = true; (kf_ToggleGodMode_pastReveal)

    ' not supported if a spectator - since they automatically get full vis
    If SPECTATOR_NO_OP() Then Exit Sub

    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If godMode Then
        If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/keybind.cpp:1316", "kf_ToggleGodMode", "Cannot disable godMode for spectator-only slots": Exit Sub
        godMode = False
        setRevealStatus kf_ToggleGodMode_pastReveal
        ' now hide the features
        it = apsFeatureList(0).head
        Do While it <> 0
            aFeature(lnObj(it)).visible(selectedPlayer) = 0
            it = lnNext(it)
        Loop

        ' and the structures
        For playerId = 0 To MAX_PLAYERS - 1
            If playerId <> selectedPlayer Then
                it = apsStructLists(playerId).head
                Do While it <> 0
                    aStruct(lnObj(it)).visible(selectedPlayer) = 0
                    it = lnNext(it)
                Loop
            End If
        Next
        ' remove all proximity messages
        releaseAllProxDisp
        radarPermitted = structureExists(selectedPlayer, REF_HQ, True, False) Or structureExists(selectedPlayer, REF_HQ, True, True)
    Else
        kf_ToggleGodMode_pastReveal = getRevealStatus()
        enableGodMode
    End If

    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, IIf(godMode, "God Mode ON", "God Mode OFF"))
    sendInGameSystemMessage cmsg
End Sub
' --------------------------------------------------------------------------
' Aligns the view to north - some people can't handle the world spinning
Public Sub kf_SeekNorth()
    playerPos.r.y = 0
    If getWarCamStatus() Then
        camToggleStatus
    End If
    CONPRINTF "%s", "View Aligned to North"
End Sub

Public Function kf_ScrollCamera(ByVal horizontal As Long, ByVal vertical As Long) As WzFunc
    Set kf_ScrollCamera = WzFunc_new(FN_kf_ScrollCamera, horizontal, vertical)
End Function

Public Sub kf_toggleTrapCursor()
    Dim value As Long, pValueStr As String, msg As String
    value = war_GetTrapCursor()
    If value = TrapCursorMode_Disabled Then
        value = oldTrapCursorEnabledValue
    ElseIf value = TrapCursorMode_Automatic And Not shouldTrapCursor() Then
        ' set to auto, but will not automatically trap the cursor - toggle to enabled
        value = TrapCursorMode_Enabled
    Else
        oldTrapCursorEnabledValue = value
        value = TrapCursorMode_Disabled
    End If
    war_SetTrapCursor value
    If shouldTrapCursor() Then wzGrabMouse Else wzReleaseMouse

    pValueStr = ""
    Select Case value
    Case TrapCursorMode_Disabled: pValueStr = "OFF"
    Case TrapCursorMode_Enabled: pValueStr = "ON"
    Case TrapCursorMode_Automatic: pValueStr = "AUTO"
    End Select

    msg = va("Trap cursor %s", pValueStr)
    addConsoleMessage msg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
End Sub

' --------------------------------------------------------------------------
Public Sub kf_TogglePauseMode()
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    Dim campaignTrue As Boolean, gameMode As Long, lev As DIFFICULTY_LEVEL
    ' Bail out if we're running a _true_ multiplayer game (which cannot be paused)
    If runningMultiplayer() Then
        Exit Sub
    End If

    ' Is the game running?
    If Not gamePaused() Then
        ' Then pause it
        setGamePauseStatus True
        setConsolePause True
        setScriptPause True
        setAudioPause True
        setScrollPause True

        ' If cursor trapping is enabled, allow the cursor to leave the window when paused
        wzReleaseMouse

        ' And stop the clock
        gameTimeStop
        addConsoleMessage "PAUSED", CENTRE_JUSTIFY, SYSTEM_MESSAGE

        ' check if campaign and display gamemode
        campaignTrue = False
        gameMode = ActivityManager_getCurrentGameMode()
        Select Case gameMode
        Case ActivitySink_GameMode_CAMPAIGN
            addConsoleMessage "CAMPAIGN", CENTRE_JUSTIFY, SYSTEM_MESSAGE
            campaignTrue = True
        Case ActivitySink_GameMode_CHALLENGE
            addConsoleMessage "CHALLENGE", CENTRE_JUSTIFY, SYSTEM_MESSAGE
        Case ActivitySink_GameMode_SKIRMISH
            addConsoleMessage "SKIRMISH", CENTRE_JUSTIFY, SYSTEM_MESSAGE
        Case Else
        End Select

        ' display level info
        addConsoleMessage getLevelName(), CENTRE_JUSTIFY, SYSTEM_MESSAGE

        ' if we are playing campaign, display difficulty
        If campaignTrue Then
            lev = getDifficultyLevel()
            Select Case lev
            Case DL_SUPER_EASY
                addConsoleMessage "DIFFICULTY: SUPER EASY", CENTRE_JUSTIFY, SYSTEM_MESSAGE
            Case DL_EASY
                addConsoleMessage "DIFFICULTY: EASY", CENTRE_JUSTIFY, SYSTEM_MESSAGE
            Case DL_NORMAL
                addConsoleMessage "DIFFICULTY: NORMAL", CENTRE_JUSTIFY, SYSTEM_MESSAGE
            Case DL_HARD
                addConsoleMessage "DIFFICULTY: HARD", CENTRE_JUSTIFY, SYSTEM_MESSAGE
            Case DL_INSANE
                addConsoleMessage "DIFFICULTY: INSANE", CENTRE_JUSTIFY, SYSTEM_MESSAGE
            End Select
        End If

        ' display cheats status
        If DebugInputManager_debugMappingsAllowed() Then
            addConsoleMessage "CHEATS: ENABLED", CENTRE_JUSTIFY, SYSTEM_MESSAGE
        Else
            addConsoleMessage "CHEATS: DISABLED", CENTRE_JUSTIFY, SYSTEM_MESSAGE
        End If
    Else
        ' Else get it going again
        setGamePauseStatus False
        setConsolePause False
        setScriptPause False
        setAudioPause False
        setScrollPause False

        ' Re-enable cursor trapping if it is enabled
        If shouldTrapCursor() Then
            wzGrabMouse
        End If

        clearActiveConsole

        ' And start the clock again
        gameTimeStart
    End If
End Sub

' --------------------------------------------------------------------------
' finish all the research for the selected player
Public Sub kf_FinishAllResearch()
    Dim j As Long, cmsg As String
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    For j = 0 To asResearch_size - 1
        If Not IsResearchCompleted(asPlayerResList(selectedPlayer, j)) Then
            If asResearch(j).excludeFromCheats <> 0 Then
                GoTo continue_
            End If
            If bMultiMessages Then
                SendResearch selectedPlayer, j, False
                ' Wait for our message before doing anything.
            Else
                researchResult j, selectedPlayer, False, 0, False
            End If
        End If
continue_:
    Next
    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, "Researched EVERYTHING for you!")
    sendInGameSystemMessage cmsg
End Sub

Public Sub kf_Reload()
    Dim intStrList As OBJ_LIST, it As Long, psCurr As Long, firePause As Long
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If interfaceStructList(intStrList) Then
        it = intStrList.head
        Do While it <> 0
            psCurr = lnObj(it)
            If isLasSat(aStruct(psCurr).pStructureType) And aStruct(psCurr).selected Then
                firePause = weaponFirePause(asWeaponStats(aStruct(psCurr).asWeaps(0).nStat), aStruct(psCurr).player)

                aStruct(psCurr).asWeaps(0).lastFired = FromU32(U32(aStruct(psCurr).asWeaps(0).lastFired) - U32(firePause))
                CONPRINTF "%s", "Selected buildings instantly recharged!"
            End If
            it = lnNext(it)
        Loop
    End If
End Sub

' --------------------------------------------------------------------------
' finish all the research for the selected player
Public Sub kf_FinishResearch()
    Dim intStrList As OBJ_LIST, it As Long, psCurr As Long, pSubject As Long, rindex As Long, cmsg As String
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If interfaceStructList(intStrList) Then
        it = intStrList.head
        Do While it <> 0
            psCurr = lnObj(it)
            If asStructureStats(aStruct(psCurr).pStructureType).typ = REF_RESEARCH Then
                ' find out what we are researching here
                pSubject = aStruct(psCurr).functionality.researchFacility.psSubject
                If pSubject <> -1 Then
                    rindex = asResearch(pSubject).index
                    If Not IsResearchCompleted(asPlayerResList(selectedPlayer, rindex)) Then
                        If bMultiMessages Then
                            SendResearch selectedPlayer, rindex, True
                            ' Wait for our message before doing anything.
                        Else
                            researchResult rindex, selectedPlayer, True, psCurr, True
                        End If
                        cmsg = va("(Player %u) is using cheat :%s %s", selectedPlayer, "Researched", asResearch(pSubject).name_)
                        sendInGameSystemMessage cmsg
                        intResearchFinished psCurr
                    Else
                        debug_ 1632, LOG_ERROR, "kf_FinishResearch", "Research already completed for player %u?: %s", selectedPlayer, asResearch(pSubject).name_
                        GoTo continue_
                    End If
                End If
            End If
continue_:
            it = lnNext(it)
        Loop
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleEnergyBars()
    Select Case toggleEnergyBars()
    Case BAR_SELECTED
        addConsoleMessage "Only displaying energy bars when selected", LEFT_JUSTIFY, SYSTEM_MESSAGE
    Case BAR_DROIDS
        addConsoleMessage "Always displaying energy bars for units", LEFT_JUSTIFY, SYSTEM_MESSAGE
    Case BAR_DROIDS_AND_STRUCTURES
        addConsoleMessage "Always displaying energy bars for units and structures", LEFT_JUSTIFY, SYSTEM_MESSAGE
    Case Else
        ASSERT_FAILURE "false", "src/keybind.cpp:1656", "kf_ToggleEnergyBars", "Bad energy bar status"
    End Select
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ChooseOptions()
    Dim cmsg As String
    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, "Debug menu is Open")
    sendInGameSystemMessage cmsg
    jsShowDebug
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleProximitys()
    setProximityDraw Not doWeDrawProximitys()
End Sub
' --------------------------------------------------------------------------
Public Sub kf_JumpToResourceExtractor()
    Dim bReset As Boolean
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    If apsExtractorLists(selectedPlayer).head = 0 Then
        addConsoleMessage "Unable to locate any oil derricks!", LEFT_JUSTIFY, SYSTEM_MESSAGE
        Exit Sub
    End If

    ' !psOldRE.has_value() || std::next(*psOldRE) == apsExtractorLists[selectedPlayer].end() (end(): node 0)
    bReset = Not psOldRE_has
    If Not bReset Then bReset = lnNext(psOldRE) = 0
    If bReset Then
        ' Reset the pointer if `psOldRE` is either not initialized yet or points to the last element.
        psOldRE = apsExtractorLists(selectedPlayer).head
        psOldRE_has = True
    Else
        psOldRE = lnNext(psOldRE)
    End If

    If lnObj(psOldRE) <> 0 Then
        playerPos.r.y = 0 ' face north
        setViewPos map_coord(aStruct(lnObj(psOldRE)).pos.x), map_coord(aStruct(lnObj(psOldRE)).pos.y), True
    Else
        addConsoleMessage "Unable to locate any oil derricks!", LEFT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub

Public Sub keybindInformResourceExtractorRemoved(ByVal psResourceExtractor As Long)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    If psOldRE_has Then
        If psOldRE <> 0 Then
            If lnObj(psOldRE) = psResourceExtractor Then
                psOldRE_has = False
                psOldRE = 0
            End If
        End If
    End If
End Sub

' --------------------------------------------------------------------------
Public Function kf_JumpToUnits(ByVal droidType As Long) As WzFunc
    Set kf_JumpToUnits = WzFunc_new(FN_kf_JumpToUnits, droidType)
End Function

' --------------------------------------------------------------------------
Public Sub kf_JumpToUnassignedUnits()
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    selNextUnassignedUnit
End Sub
' --------------------------------------------------------------------------

Public Sub kf_ToggleOverlays()
    Dim hide As Boolean
    hide = False
    If getWidgetsStatus() Then
        If Not isChatUp() Then
            hide = Not gamePaused()
        End If
    End If
    If hide Then
        setWidgetsStatus False
    Else
        setWidgetsStatus True
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ChooseCommand()
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    If intCheckReticuleButEnabled(IDRET_COMMAND) Then
        setKeyButtonMapping IDRET_COMMAND
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ChooseManufacture()
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    If intCheckReticuleButEnabled(IDRET_MANUFACTURE) Then
        setKeyButtonMapping IDRET_MANUFACTURE
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ChooseResearch()
    If intCheckReticuleButEnabled(IDRET_RESEARCH) Then
        setKeyButtonMapping IDRET_RESEARCH
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ChooseBuild()
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    If intCheckReticuleButEnabled(IDRET_BUILD) Then
        setKeyButtonMapping IDRET_BUILD
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ChooseDesign()
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    If intCheckReticuleButEnabled(IDRET_DESIGN) Then
        setKeyButtonMapping IDRET_DESIGN
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ChooseIntelligence()
    If intCheckReticuleButEnabled(IDRET_INTEL_MAP) Then
        setKeyButtonMapping IDRET_INTEL_MAP
    End If
End Sub

' --------------------------------------------------------------------------

Public Sub kf_ChooseCancel()
    setKeyButtonMapping IDRET_CANCEL
End Sub

' --------------------------------------------------------------------------
Public Sub kf_MovePause()
    If Not bMultiPlayer Then ' can't do it in multiplay
        If Not bMovePause Then
            ' Then pause it
            setGamePauseStatus True
            setConsolePause True
            setScriptPause True
            setAudioPause True
            ' And stop the clock
            gameTimeStop
            setWidgetsStatus False
            bMovePause = True
        Else
            setWidgetsStatus True
            ' Else get it going again
            setGamePauseStatus False
            setConsolePause False
            setScriptPause False
            setAudioPause False
            ' And start the clock again
            gameTimeStart
            bMovePause = False
        End If
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_MoveToLastMessagePos()
    Dim iX As Long, iY As Long, iZ As Long

    If Not audio_GetPreviousQueueTrackPos(iX, iY, iZ) Then
        Exit Sub
    End If
    If iX <> 0 And iY <> 0 Then
        setViewPos map_coord(iX), map_coord(iY), True
    End If
End Sub
' --------------------------------------------------------------------------
' Makes it snow if it's not snowing and stops it if it is
Public Sub kf_ToggleWeather()
    If atmosGetWeatherType() = WT_NONE Then
        atmosSetWeatherType WT_SNOWING
        addConsoleMessage "Oh, the weather outside is frightful... SNOW", LEFT_JUSTIFY, SYSTEM_MESSAGE
    ElseIf atmosGetWeatherType() = WT_SNOWING Then
        atmosSetWeatherType WT_RAINING
        addConsoleMessage "Singing in the rain, I'm singing in the rain... RAIN", LEFT_JUSTIFY, SYSTEM_MESSAGE
    Else
        atmosInitSystem
        atmosSetWeatherType WT_NONE
        addConsoleMessage "Forecast : Clear skies for all areas... NO WEATHER", LEFT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub

' --------------------------------------------------------------------------
Public Function kf_SelectNextFactory(ByVal factoryType As Long, Optional ByVal bJumpToSelected As Boolean = False) As WzFunc
    ' static const std::vector<STRUCTURE_TYPE> FACTORY_TYPES = { REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY } (FACTORY_TYPES_contains)
    Set kf_SelectNextFactory = WzFunc_new(FN_kf_SelectNextFactory, factoryType, bJumpToSelected)
End Function

' kf_SelectNextFactory's FACTORY_TYPES: std::any_of(FACTORY_TYPES, [currentType](const STRUCTURE_TYPE type) { return currentType == type; })
Private Function FACTORY_TYPES_contains(ByVal currentType As Long) As Boolean
    FACTORY_TYPES_contains = currentType = REF_FACTORY Or currentType = REF_CYBORG_FACTORY Or currentType = REF_VTOL_FACTORY
End Function

' kf_SelectNextFactory's [factoryType, bJumpToSelected]() { ... }
Private Sub kf_SelectNextFactory_lambda(ByVal factoryType As Long, ByVal bJumpToSelected As Boolean)
    Dim it As Long, psCurrent As Long, currentType As Long, bIsAnotherTypeOfFactory As Boolean
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    selNextSpecifiedBuilding factoryType, bJumpToSelected

    'deselect factories of other types
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        psCurrent = lnObj(it)
        currentType = asStructureStats(aStruct(psCurrent).pStructureType).typ
        bIsAnotherTypeOfFactory = currentType <> factoryType And FACTORY_TYPES_contains(currentType)
        If aStruct(psCurrent).selected And bIsAnotherTypeOfFactory Then
            aStruct(psCurrent).selected = False
        End If
        it = lnNext(it)
    Loop

    If intCheckReticuleButEnabled(IDRET_MANUFACTURE) Then
        setKeyButtonMapping IDRET_MANUFACTURE
    End If

    triggerEventSelected
End Sub

' --------------------------------------------------------------------------
Public Function kf_SelectNextResearch(Optional ByVal bJumpToSelected As Boolean = False) As WzFunc
    Set kf_SelectNextResearch = WzFunc_new(FN_kf_SelectNextResearch, bJumpToSelected)
End Function

' --------------------------------------------------------------------------
Public Function kf_SelectNextPowerStation(Optional ByVal bJumpToSelected As Boolean = False) As WzFunc
    Set kf_SelectNextPowerStation = WzFunc_new(FN_kf_SelectNextPowerStation, bJumpToSelected)
End Function

' --------------------------------------------------------------------------
Public Sub kf_KillEnemy()
    Dim cmsg As String, playerId As Long, it As Long
    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    If ULT(MAX_PLAYERS - 1, selectedPlayer) Then Exit Sub

    debug_ 1970, LOG_DEATH, "kf_KillEnemy", "Destroying enemy droids and structures"
    CONPRINTF "%s", "Warning! This can have drastic consequences if used incorrectly in missions."
    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, "All enemies destroyed by cheating!")
    sendInGameSystemMessage cmsg
    Cheated = True

    For playerId = 0 To MAX_PLAYERS - 1
        If playerId <> selectedPlayer And Not aiCheckAlliances(selectedPlayer, playerId) Then
            ' wipe out all the droids
            it = apsDroidLists(playerId).head
            Do While it <> 0
                SendDestroyDroid lnObj(it)
                it = lnNext(it)
            Loop
            ' wipe out all their structures
            it = apsStructLists(playerId).head
            Do While it <> 0
                SendDestroyStructure lnObj(it)
                it = lnNext(it)
            Loop
        End If
    Next
End Sub

' kill all the selected objects
Public Sub kf_KillSelected()
    Dim cmsg As String, it As Long
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    ' Bail out if we're running a _true_ multiplayer game (to prevent MP cheating)
    If runningMultiplayer() Then
        noMPCheatMsg
        Exit Sub
    End If

    cmsg = va("(Player %u) is using cheat :%s", _
              selectedPlayer, "Destroying selected droids and structures!")
    sendInGameSystemMessage cmsg

    debug_ 2014, LOG_DEATH, "kf_KillSelected", "Destroying selected droids and structures"
    audio_PlayTrack ID_SOUND_COLL_DIE
    Cheated = True

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        If aDroid(lnObj(it)).selected Then
            SendDestroyDroid lnObj(it)
        End If
        it = lnNext(it)
    Loop
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        If aStruct(lnObj(it)).selected Then
            SendDestroyStructure lnObj(it)
        End If
        it = lnNext(it)
    Loop
End Sub

' --------------------------------------------------------------------------
' Chat message. NOTE THIS FUNCTION CAN DISABLE ALL OTHER KEYPRESSES
Private Sub OpenChatUI(ByVal mode As Long, ByVal startWithQuickChatFocused As Boolean)
    If mode = CHAT_TEAM Then
        ' not supported if a spectator
        If SPECTATOR_NO_OP() Then Exit Sub
    End If

    If Not getWidgetsStatus() Then
        Exit Sub
    End If

    If bAllowOtherKeyPresses And Not gamePaused() Then ' just starting.
        sCurrentConsoleText = ""          'for beacons
        inputClearBuffer
        chatDialog mode, startWithQuickChatFocused  ' throw up the dialog
    End If
End Sub

' Chat message. NOTE THIS FUNCTION CAN DISABLE ALL OTHER KEYPRESSES
Public Sub kf_SendTeamMessage()
    OpenChatUI CHAT_TEAM, False
End Sub

' Chat message. NOTE THIS FUNCTION CAN DISABLE ALL OTHER KEYPRESSES
Public Sub kf_SendGlobalMessage()
    OpenChatUI CHAT_GLOB, False
End Sub

' Chat message. NOTE THIS FUNCTION CAN DISABLE ALL OTHER KEYPRESSES
Public Sub kf_SendTeamQuickChat()
    OpenChatUI CHAT_TEAM, True
End Sub

' Chat message. NOTE THIS FUNCTION CAN DISABLE ALL OTHER KEYPRESSES
Public Sub kf_SendGlobalQuickChat()
    OpenChatUI CHAT_GLOB, True
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleConsole()
    If getConsoleDisplayStatus() Then
        enableConsoleDisplay False
    Else
        enableConsoleDisplay True
    End If
End Sub

' --------------------------------------------------------------------------
Public Function kf_SelectUnits(ByVal selectionType_ As SELECTIONTYPE, Optional ByVal selectionClass As SELECTION_CLASS = DS_BY_TYPE, Optional ByVal bOnScreen As Boolean = False) As WzFunc
    Set kf_SelectUnits = WzFunc_new(FN_kf_SelectUnits, selectionClass, selectionType_, bOnScreen)
End Function

' --------------------------------------------------------------------------
Public Function kf_SelectNoGroupUnits(ByVal selectionType_ As SELECTIONTYPE, Optional ByVal selectionClass As SELECTION_CLASS = DS_BY_TYPE, Optional ByVal bOnScreen As Boolean = False) As WzFunc
    Set kf_SelectNoGroupUnits = WzFunc_new(FN_kf_SelectNoGroupUnits, selectionClass, selectionType_, bOnScreen)
End Function

' --------------------------------------------------------------------------
Public Function kf_SetDroid(ByVal order As SECONDARY_ORDER, ByVal state As SECONDARY_STATE) As WzFunc
    Set kf_SetDroid = WzFunc_new(FN_kf_SetDroid, order, state)
End Function

' --------------------------------------------------------------------------
Public Function kf_OrderDroid(ByVal order As DroidOrderType) As WzFunc
    Set kf_OrderDroid = WzFunc_new(FN_kf_OrderDroid, order)
End Function

' kf_OrderDroid's [order]() { ... }
Private Sub kf_OrderDroid_lambda(ByVal order As Long)
    Dim it As Long, psDroid As Long
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected Then
            orderDroid psDroid, order, ModeQueue
        End If
        it = lnNext(it)
    Loop
    intRefreshOrder
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleVisibility()
    If getRevealStatus() Then
        CONPRINTF "%s", "Reveal OFF"
        setRevealStatus False
    Else
        CONPRINTF "%s", "Reveal ON"
        setRevealStatus True
    End If
End Sub

' --------------------------------------------------------------------------
Private Sub kfsf_SetSelectedDroidsState(ByVal sec As SECONDARY_ORDER, ByVal state As SECONDARY_STATE)
    Dim it As Long, psDroid As Long
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    ' NOT A CHEAT CODE
    ' This is a function to set unit orders via keyboard shortcuts. It should
    ' _not_ be disallowed in multiplayer games.

    ' This code is similar to SetSecondaryState() in intorder.cpp. Unfortunately, it seems hard to un-duplicate the code.
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        ' Only set the state if it's not a transporter.
        If aDroid(psDroid).selected And Not DROID_isTransporter(psDroid) Then
            secondarySetState psDroid, sec, state
        End If
        it = lnNext(it)
    Loop
    intRefreshOrder
End Sub

' --------------------------------------------------------------------------
Public Sub kf_TriggerRayCast()
    Dim it As Long
    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    ' std::find_if(apsDroidLists[selectedPlayer], [](DROID* d) { return d->selected; })
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        If aDroid(lnObj(it)).selected Then Exit Do
        it = lnNext(it)
    Loop

    If it <> 0 Then
'       getBlockHeightDirToEdgeOfGrid(UDWORD x, UDWORD y, UBYTE direction, UDWORD *height, UDWORD *dist)
'       getBlockHeightDirToEdgeOfGrid(psOther->pos.x,psOther->pos.y,psOther->direction,&height,&dist);
'       getBlockHeightDirToEdgeOfGrid(mouseTileX*TILE_UNITS,mouseTileY*TILE_UNITS,getTestAngle(),&height,&dist);
    End If
End Sub

' --------------------------------------------------------------------------

Public Sub kf_CentreOnBase()
    Dim bGotHQ As Boolean, xJump As Long, yJump As Long, it As Long, pStruct As Long
    bGotHQ = False
    xJump = 0
    yJump = 0

    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    ' Got through our buildings
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        pStruct = lnObj(it)
        ' Have we got a HQ?
        If asStructureStats(aStruct(pStruct).pStructureType).typ = REF_HQ Then
            bGotHQ = True
            xJump = aStruct(pStruct).pos.x
            yJump = aStruct(pStruct).pos.y
            Exit Do
        End If
        it = lnNext(it)
    Loop

    ' If we found it, then jump to it!
    If bGotHQ Then
        addConsoleMessage "Centered on player HQ, direction NORTH", LEFT_JUSTIFY, SYSTEM_MESSAGE
        playerPos.p.x = xJump
        playerPos.p.z = yJump
        playerPos.r.y = 0 ' face north
        ' A fix to stop the camera continuing when marker code is called
        If getWarCamStatus() Then
            camToggleStatus
        End If
    Else
        addConsoleMessage "Unable to locate HQ!", LEFT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleFormationSpeedLimiting()
    Dim resultingValue As Boolean
    resultingValue = False
    If moveToggleFormationSpeedLimiting(selectedPlayer, True, resultingValue) Then
        If resultingValue Then
            addConsoleMessage "Formation speed limiting ON", LEFT_JUSTIFY, SYSTEM_MESSAGE
        Else
            addConsoleMessage "Formation speed limiting OFF", LEFT_JUSTIFY, SYSTEM_MESSAGE
        End If
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_RightOrderMenu()
    Dim psGotOne As Long, bFound As Boolean, it As Long, psDroid As Long
    psGotOne = 0
    bFound = False

    ' if menu open, then close it!
    If Not widgGetFromID(psWScreen, IDORDER_FORM) Is Nothing Then
        intRemoveOrder  ' close the screen.
        Exit Sub
    End If

    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected Then ' && droidOnScreen(psDroid,0))
            bFound = True
            psGotOne = psDroid
            Exit Do
        End If
        it = lnNext(it)
    Loop
    If bFound Then
        intResetScreen True
        intObjectSelected psGotOne
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleMouseInvert()
    If getInvertMouseStatus() Then
        setInvertMouseStatus False
        CONPRINTF "%s", "Vertical rotation direction: Normal"
    Else
        setInvertMouseStatus True
        CONPRINTF "%s", "Vertical rotation direction: Flipped"
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleShakeStatus()
    If getShakeStatus() Then
        setShakeStatus False
        CONPRINTF "%s", "Screen shake when things die: Off"
    Else
        setShakeStatus True
        CONPRINTF "%s", "Screen shake when things die: On"
    End If
End Sub

' --------------------------------------------------------------------------
Public Sub kf_ToggleShadows()
    If getDrawShadows() Then
        setDrawShadows False
    Else
        setDrawShadows True
    End If
End Sub
' --------------------------------------------------------------------------

' static const Rational available_speed[] (its numerators and denominators, Rational's normalised form)
Private Sub available_speed_init()
    Dim i As Long
    i = 0
' p = pumpkin allowed, n = new entries allowed in debug mode only.
' Since some of these values can ruin a SP game, we disallow them in normal mode.
    available_speed_n(i) = 0: available_speed_d(i) = 1: i = i + 1        ' n
    available_speed_n(i) = 1: available_speed_d(i) = 1000: i = i + 1     ' n
    available_speed_n(i) = 1: available_speed_d(i) = 100: i = i + 1      ' n
    available_speed_n(i) = 1: available_speed_d(i) = 40: i = i + 1       ' n
    available_speed_n(i) = 1: available_speed_d(i) = 8: i = i + 1        ' n
    available_speed_n(i) = 1: available_speed_d(i) = 5: i = i + 1        ' n
    available_speed_n(i) = 1: available_speed_d(i) = 3: i = i + 1        ' p
    available_speed_n(i) = 1: available_speed_d(i) = 2: i = i + 1
    available_speed_n(i) = 3: available_speed_d(i) = 4: i = i + 1        ' p
    available_speed_n(i) = 1: available_speed_d(i) = 1: i = i + 1        ' p
    available_speed_n(i) = 5: available_speed_d(i) = 4: i = i + 1        ' p
    available_speed_n(i) = 3: available_speed_d(i) = 2: i = i + 1        ' p
    available_speed_n(i) = 2: available_speed_d(i) = 1: i = i + 1        ' p (in debug mode only)
    available_speed_n(i) = 5: available_speed_d(i) = 2: i = i + 1        ' n
    available_speed_n(i) = 3: available_speed_d(i) = 1: i = i + 1        ' n
    available_speed_n(i) = 10: available_speed_d(i) = 1: i = i + 1       ' n
    available_speed_n(i) = 20: available_speed_d(i) = 1: i = i + 1       ' n
    available_speed_n(i) = 30: available_speed_d(i) = 1: i = i + 1       ' n
    available_speed_n(i) = 60: available_speed_d(i) = 1: i = i + 1       ' n
    available_speed_n(i) = 100: available_speed_d(i) = 1: i = i + 1      ' n
End Sub

' (VB) Rational's operator < (n * b.d < b.n * d)
Private Function Rational_less(ByVal an As Long, ByVal ad As Long, ByVal bn As Long, ByVal bd As Long) As Boolean
    Rational_less = an * bd < bn * ad
End Function

Private Sub tryChangeSpeed(ByVal newMod_n As Long, ByVal newMod_d As Long, ByVal oldMod_n As Long, ByVal oldMod_d As Long)
    Dim modString As String, speed As Single
    ' Bail out if we're running a _true_ multiplayer game or are playing a tutorial
    If runningMultiplayer() Or bInTutorial Then
        If Not bInTutorial Then
            addConsoleMessage "Sorry, but game speed cannot be changed in multiplayer.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        End If
        Exit Sub
    End If

    ' only in debug/cheat mode do we enable all time compression speeds.
    ' newMod > 2 || newMod <= 0 (Rational(2), Rational(0))
    If Not DebugInputManager_debugMappingsAllowed() And (Rational_less(2, 1, newMod_n, newMod_d) Or Not Rational_less(0, 1, newMod_n, newMod_d)) Then ' 2 = max officially allowed time compression
        Exit Sub
    End If

    ' char modString[30];
    If newMod_d <> 1 Then
        speed = CSng(newMod_n) / CSng(newMod_d)
        If CDbl(speed) > 0.4 Then
            modString = va("%.1f", CDbl(speed))
        ElseIf CDbl(speed) > 0.15 Then
            modString = va("%.2f", CDbl(speed))
        Else
            modString = va("%.3f", CDbl(speed))
        End If
    Else
        modString = va("%d", newMod_n)
    End If

    If newMod_n * 1 = 1 * newMod_d Then ' newMod == 1
        CONPRINTF "%s", "Game Speed Reset"
    ElseIf Rational_less(oldMod_n, oldMod_d, newMod_n, newMod_d) Then ' newMod > oldMod
        CONPRINTF "Game Speed Increased to %s", modString
    Else
        CONPRINTF "Game Speed Reduced to %s", modString
    End If
    gameTimeSetMod newMod_n, newMod_d
End Sub

Public Sub kf_SpeedUp()
    Dim mod_n As Long, mod_d As Long, newMod As Long
    ' get the current modifier
    gameTimeGetMod mod_n, mod_d

    ' std::upper_bound(available_speed, available_speed + nb_available_speeds, mod): the first speed > mod
    newMod = 0
    Do While newMod < nb_available_speeds
        If Rational_less(mod_n, mod_d, available_speed_n(newMod), available_speed_d(newMod)) Then Exit Do
        newMod = newMod + 1
    Loop
    If newMod = nb_available_speeds Then
        Exit Sub  ' Already at maximum speed.
    End If

    tryChangeSpeed available_speed_n(newMod), available_speed_d(newMod), mod_n, mod_d
End Sub

Public Sub kf_SlowDown()
    Dim mod_n As Long, mod_d As Long, newMod As Long
    ' get the current modifier
    gameTimeGetMod mod_n, mod_d

    ' std::lower_bound(available_speed, available_speed + nb_available_speeds, mod): the first speed >= mod
    newMod = 0
    Do While newMod < nb_available_speeds
        If Not Rational_less(available_speed_n(newMod), available_speed_d(newMod), mod_n, mod_d) Then Exit Do
        newMod = newMod + 1
    Loop
    If newMod = 0 Then
        Exit Sub  ' Already at minimum speed.
    End If

    newMod = newMod - 1  ' Point to lower speed instead of current speed.
    tryChangeSpeed available_speed_n(newMod), available_speed_d(newMod), mod_n, mod_d
End Sub

Public Sub kf_NormalSpeed()
    ' Bail out if we're running a _true_ multiplayer game or are playing a tutorial
    If runningMultiplayer() Or bInTutorial Then
        If Not bInTutorial Then
            noMPCheatMsg
        End If
        Exit Sub
    End If

    CONPRINTF "%s", "Game Speed Reset"
    gameTimeResetMod
End Sub

' --------------------------------------------------------------------------

Public Sub kf_ToggleRadarAllyEnemy()
    bEnemyAllyRadarColor = Not bEnemyAllyRadarColor

    If bEnemyAllyRadarColor Then
        CONPRINTF "%s", "Radar showing friend-foe colors"
    Else
        CONPRINTF "%s", "Radar showing player colors"
    End If
    resizeRadar
End Sub

Public Sub kf_ToggleRadarTerrain()
    radarDrawMode = radarDrawMode + 1
    If radarDrawMode >= NUM_RADAR_MODES Then
        radarDrawMode = 0
    End If
    Select Case radarDrawMode
    Case RADAR_MODE_NO_TERRAIN
        CONPRINTF "%s", "Radar showing only objects"
    Case RADAR_MODE_COMBINED
        CONPRINTF "%s", "Radar blending terrain and height"
    Case RADAR_MODE_TERRAIN
        CONPRINTF "%s", "Radar showing terrain"
    Case RADAR_MODE_HEIGHT_MAP
        CONPRINTF "%s", "Radar showing height"
    Case NUM_RADAR_MODES
        ' assert(false); (release: nothing)
    End Select
End Sub

'Add a beacon (blip)
Public Sub kf_AddHelpBlip()
    Dim worldX As Long, worldY As Long, i As Long, tempStr As String, x As Long, y As Long
    worldX = -1
    worldY = -1

    ' not needed in campaign
    If Not bMultiPlayer Then
        Exit Sub
    End If

    ' not supported if a spectator
    If SPECTATOR_NO_OP() Then Exit Sub

    debug_ 2517, LOG_WZ, "kf_AddHelpBlip", "Adding beacon='%s'", sCurrentConsoleText

    ' check if clicked on radar
    x = mouseX()
    y = mouseY()
    If isMouseOverRadar() Then
        CalcRadarPosition x, y, worldX, worldY
        worldX = worldX * TILE_UNITS + TILE_UNITS \ 2
        worldY = worldY * TILE_UNITS + TILE_UNITS \ 2
    ' convert screen to world
    Else
        worldX = mouseTileX * TILE_UNITS + TILE_UNITS \ 2
        worldY = mouseTileY * TILE_UNITS + TILE_UNITS \ 2
    End If

    tempStr = Left$(getPlayerName(selectedPlayer), 254)     'temporary solution

    ' add beacon for the sender
    beaconMsg(selectedPlayer) = Left$(tempStr, MAX_CONSOLE_STRING_LENGTH - 1)
    sendBeaconToPlayer worldX, worldY, selectedPlayer, selectedPlayer, beaconMsg(selectedPlayer)

    ' send beacon to other players
    For i = 0 To game.maxPlayers - 1
        If openchannels(i) And (i <> selectedPlayer) Then
            beaconMsg(i) = Left$(tempStr, MAX_CONSOLE_STRING_LENGTH - 1)
            sendBeaconToPlayer worldX, worldY, i, selectedPlayer, beaconMsg(i)
        End If
    Next
End Sub

Public Sub kf_NoAssert()
    debugDisableAssert
    console "Asserts turned off"
    debug_ 2556, LOG_ERROR, "kf_NoAssert", "Asserts turned off"
End Sub

' routine to decrement the tab-scroll 'buttons'
Public Sub kf_BuildPrevPage()
    Dim psTForm As WIDGET
    If Not (Not psWScreen Is Nothing) Then ASSERT_FAILURE "psWScreen != nullptr", "src/keybind.cpp:2562", "kf_BuildPrevPage", " Invalid screen pointer!": Exit Sub
    Set psTForm = widgGetFromID(psWScreen, IDSTAT_TABFORM)

    If psTForm Is Nothing Then
        Exit Sub
    End If

    If ListTabWidget_currentPage(psTForm) = 0 Then
        audio_PlayBuildFailedOnce
        Exit Sub
    ElseIf Not ListTabWidget_setCurrentPage(psTForm, ListTabWidget_currentPage(psTForm) - 1) Then
        audio_PlayBuildFailedOnce
        Exit Sub
    End If

    audio_PlayTrack ID_SOUND_BUTTON_CLICK_5
End Sub

' routine to advance the tab-scroll 'buttons'
Public Sub kf_BuildNextPage()
    Dim psTForm As WIDGET
    If Not (Not psWScreen Is Nothing) Then ASSERT_FAILURE "psWScreen != nullptr", "src/keybind.cpp:2581", "kf_BuildNextPage", " Invalid screen pointer!": Exit Sub
    Set psTForm = widgGetFromID(psWScreen, IDSTAT_TABFORM)

    If psTForm Is Nothing Then
        Exit Sub
    End If

    If Not ListTabWidget_setCurrentPage(psTForm, ListTabWidget_currentPage(psTForm) + 1) Then
        ' went over max
        audio_PlayBuildFailedOnce
        Exit Sub
    End If

    audio_PlayTrack ID_SOUND_BUTTON_CLICK_5
End Sub

Public Sub kf_QuickSave()
    Dim filename As String, quickSaveFolder As String
    ' Bail out if we're running a _true_ multiplayer game or are playing a tutorial
    If runningMultiplayer() Or bInTutorial Then
        console "QuickSave not allowed for multiplayer or tutorial games"
        Exit Sub
    End If
    If InGameOpUp Or isInGamePopupUp Then
        Exit Sub
    End If
    If getCamTweakOption_AutosavesOnly() Then
        console "QuickSave not allowed in Autosaves-Only mode"
        Exit Sub
    End If
    If NETisReplay() Then
        Exit Sub ' Bail out if we're running a replay
    End If

    If bMultiPlayer Then filename = QUICKSAVE_SKI_FILENAME Else filename = QUICKSAVE_CAM_FILENAME
    If bMultiPlayer Then quickSaveFolder = QUICKSAVE_SKI_FOLDER Else quickSaveFolder = QUICKSAVE_CAM_FOLDER
    If WZ_PHYSFS_isDirectory(quickSaveFolder) <> 0 Then
        deleteSaveGame quickSaveFolder
    End If
    If saveGame(filename, GTYPE_SAVE_MIDMISSION) Then ' still expects a .gam filename... TODO: FIX
        console "QuickSave"
    Else
        console "QuickSave failed"
    End If
End Sub

Public Sub kf_QuickLoad()
    Dim filename As String, jsonName As String
    ' Bail out if we're running a _true_ multiplayer game or are playing a tutorial
    If runningMultiplayer() Or bInTutorial Then
        console "QuickLoad not allowed for multiplayer or tutorial games"
        Exit Sub
    End If
    If InGameOpUp Or isInGamePopupUp Then
        Exit Sub
    End If
    If getCamTweakOption_AutosavesOnly() Then
        console "QuickLoad not allowed in Autosaves-Only mode"
        Exit Sub
    End If

    If bMultiPlayer Then filename = QUICKSAVE_SKI_FILENAME Else filename = QUICKSAVE_CAM_FILENAME
    If bMultiPlayer Then jsonName = QUICKSAVE_SKI_JSON_FILENAME Else jsonName = QUICKSAVE_CAM_JSON_FILENAME
    ' check for .json version, because that's what going to be loaded anyway
    If PHYSFS_exists(filename) <> 0 Or PHYSFS_exists(jsonName) <> 0 Then
        console "QuickLoad"
        audio_StopAll
        'clear out any mission widgets - timers etc that may be on the screen
        clearMissionWidgets
        setWidgetsStatus True
        intResetScreen False
        wzSetCursor CURSOR_DEFAULT
        loopMissionState = LMS_LOADGAME
        saveGameName = Left$(filename, 255)
    Else
        console "QuickSave file does not exist yet"
    End If
End Sub

Public Sub kf_ToggleFullscreen()
    war_setWindowMode wzAltEnterToggleFullscreen()
End Sub

Public Sub kf_ToggleSpecOverlays()
    If Not ULT(realSelectedPlayer, NetPlay_players_size()) Then
        Exit Sub
    End If
    If Not NetPlay_players_isSpectator(realSelectedPlayer) Then
        Exit Sub
    End If
    specToggleOverlays
End Sub

Public Sub keybindShutdown()
    psOldRE_has = False
    psOldRE = 0
End Sub

' MARK: - the MappableFunctions

Public Function keybind_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_kf_ChooseManufacture: kf_ChooseManufacture
    Case FN_kf_ChooseResearch: kf_ChooseResearch
    Case FN_kf_ChooseBuild: kf_ChooseBuild
    Case FN_kf_ChooseDesign: kf_ChooseDesign
    Case FN_kf_ChooseIntelligence: kf_ChooseIntelligence
    Case FN_kf_ChooseCommand: kf_ChooseCommand
    Case FN_kf_QuickSave: kf_QuickSave
    Case FN_kf_ToggleRadar: kf_ToggleRadar
    Case FN_kf_QuickLoad: kf_QuickLoad
    Case FN_kf_ToggleConsole: kf_ToggleConsole
    Case FN_kf_ToggleEnergyBars: kf_ToggleEnergyBars
    Case FN_kf_ScreenDump: kf_ScreenDump
    Case FN_kf_ToggleFormationSpeedLimiting: kf_ToggleFormationSpeedLimiting
    Case FN_kf_MoveToLastMessagePos: kf_MoveToLastMessagePos
    Case FN_kf_ToggleSensorDisplay: kf_ToggleSensorDisplay
    Case FN_kf_addMultiMenu: kf_addMultiMenu
    Case FN_kf_SeekNorth: kf_SeekNorth
    Case FN_kf_ToggleCamera: kf_ToggleCamera
    Case FN_kf_addInGameOptions: kf_addInGameOptions
    Case FN_kf_PitchForward: kf_PitchForward
    Case FN_kf_RotateLeft: kf_RotateLeft
    Case FN_kf_ResetPitch: kf_ResetPitch
    Case FN_kf_RotateRight: kf_RotateRight
    Case FN_kf_PitchBack: kf_PitchBack
    Case FN_kf_RightOrderMenu: kf_RightOrderMenu
    Case FN_kf_SlowDown: kf_SlowDown
    Case FN_kf_SpeedUp: kf_SpeedUp
    Case FN_kf_NormalSpeed: kf_NormalSpeed
    Case FN_kf_FaceNorth: kf_FaceNorth
    Case FN_kf_FaceSouth: kf_FaceSouth
    Case FN_kf_FaceEast: kf_FaceEast
    Case FN_kf_FaceWest: kf_FaceWest
    Case FN_kf_JumpToResourceExtractor: kf_JumpToResourceExtractor
    Case FN_kf_ToggleOverlays: kf_ToggleOverlays
    Case FN_kf_ToggleConsoleDrop: kf_ToggleConsoleDrop
    Case FN_kf_ToggleTeamChat: kf_ToggleTeamChat
    Case FN_kf_RotateBuildingCW: kf_RotateBuildingCW
    Case FN_kf_RotateBuildingACW: kf_RotateBuildingACW
    Case FN_kf_CentreOnBase: kf_CentreOnBase
    Case FN_kf_JumpToUnassignedUnits: kf_JumpToUnassignedUnits
    Case FN_kf_SendGlobalMessage: kf_SendGlobalMessage
    Case FN_kf_SendTeamMessage: kf_SendTeamMessage
    Case FN_kf_SendGlobalQuickChat: kf_SendGlobalQuickChat
    Case FN_kf_SendTeamQuickChat: kf_SendTeamQuickChat
    Case FN_kf_AddHelpBlip: kf_AddHelpBlip
    Case FN_kf_ToggleShadows: kf_ToggleShadows
    Case FN_kf_toggleTrapCursor: kf_toggleTrapCursor
    Case FN_kf_ToggleRadarTerrain: kf_ToggleRadarTerrain
    Case FN_kf_ToggleRadarAllyEnemy: kf_ToggleRadarAllyEnemy
    Case FN_kf_ShowMappings: kf_ShowMappings
    Case FN_kf_FrameRate: kf_FrameRate
    Case FN_kf_ToggleDebugMappings: kf_ToggleDebugMappings
    Case FN_kf_PrioritizeDebugMappings: kf_PrioritizeDebugMappings
    Case FN_kf_ToggleLevelEditor: kf_ToggleLevelEditor
    Case FN_kf_ToggleShowPath: kf_ToggleShowPath
    Case FN_kf_ToggleShowGateways: kf_ToggleShowGateways
    Case FN_kf_ToggleVisibility: kf_ToggleVisibility
    Case FN_kf_RaiseTile: kf_RaiseTile
    Case FN_kf_LowerTile: kf_LowerTile
    Case FN_kf_ToggleFog: kf_ToggleFog
    Case FN_kf_ToggleWeather: kf_ToggleWeather
    Case FN_kf_TriFlip: kf_TriFlip
    Case FN_kf_PerformanceSample: kf_PerformanceSample
    Case FN_kf_AllAvailable: kf_AllAvailable
    Case FN_kf_KillSelected: kf_KillSelected
    Case FN_kf_ToggleGodMode: kf_ToggleGodMode
    Case FN_kf_ChooseOptions: kf_ChooseOptions
    Case FN_kf_FinishResearch: kf_FinishResearch
    Case FN_kf_RevealMapAtPos: kf_RevealMapAtPos
    Case FN_kf_TraceObject: kf_TraceObject
    Case FN_kf_ToggleFullscreen: kf_ToggleFullscreen

    Case FN_kf_Zoom                     ' [multiplier]() { incrementViewDistance(war_GetMapZoomRate() * multiplier); }
        incrementViewDistance war_GetMapZoomRate() * CLng(f.cap(0))
    Case FN_kf_RadarZoom                ' [multiplier]()
        kf_RadarZoom_lambda CLng(f.cap(0))
    Case FN_kf_SelectGrouping_N         ' [n]() { SPECTATOR_NO_OP(); kf_SelectGrouping(n); }
        ' not supported if a spectator
        If Not SPECTATOR_NO_OP() Then kf_SelectGrouping CLng(f.cap(0))
    Case FN_kf_AssignGrouping_N         ' [n]() { SPECTATOR_NO_OP(); assignObjectToGroup(selectedPlayer, n, true); }
        If Not SPECTATOR_NO_OP() Then assignObjectToGroup selectedPlayer, CLng(f.cap(0)), True
    Case FN_kf_AddGrouping_N            ' [n]() { SPECTATOR_NO_OP(); assignObjectToGroup(selectedPlayer, n, false); }
        If Not SPECTATOR_NO_OP() Then assignObjectToGroup selectedPlayer, CLng(f.cap(0)), False
    Case FN_kf_RemoveFromGrouping       ' []() { SPECTATOR_NO_OP(); removeObjectFromGroup(selectedPlayer); }
        If Not SPECTATOR_NO_OP() Then removeObjectFromGroup selectedPlayer
    Case FN_kf_SelectCommander_N        ' [n]() { SPECTATOR_NO_OP(); selCommander(n); }
        If Not SPECTATOR_NO_OP() Then selCommander CLng(f.cap(0))
    Case FN_kf_JumpToMapMarker          ' [x, z, yaw]()
        If Not getRadarTrackingStatus() Then
            playerPos.p.x = CLng(f.cap(0))
            playerPos.p.z = CLng(f.cap(1))
            playerPos.r.y = CLng(f.cap(2))
            ' A fix to stop the camera continuing when marker code is called
            If getWarCamStatus() Then
                camToggleStatus
            End If
        End If
    Case FN_kf_ScrollCamera             ' [horizontal, vertical]()
        scrollDirLeftRight = scrollDirLeftRight + CLng(f.cap(0))
        scrollDirUpDown = scrollDirUpDown + CLng(f.cap(1))
    Case FN_kf_ShowMappings_onClose     ' []() { intShowGroupSelectionMenu(); }
        ' the setting for group menu display may have been modified
        intShowGroupSelectionMenu
    Case FN_kf_JumpToUnits              ' [droidType]() { SPECTATOR_NO_OP(); selNextSpecifiedUnit(droidType); }
        If Not SPECTATOR_NO_OP() Then selNextSpecifiedUnit CLng(f.cap(0))
    Case FN_kf_SelectNextFactory        ' [factoryType, bJumpToSelected]()
        kf_SelectNextFactory_lambda CLng(f.cap(0)), CBool(f.cap(1))
    Case FN_kf_SelectNextResearch       ' [bJumpToSelected]()
        ' not supported if a spectator
        If Not SPECTATOR_NO_OP() Then
            selNextSpecifiedBuilding REF_RESEARCH, CBool(f.cap(0))
            If intCheckReticuleButEnabled(IDRET_RESEARCH) Then
                setKeyButtonMapping IDRET_RESEARCH
            End If
            triggerEventSelected
        End If
    Case FN_kf_SelectNextPowerStation   ' [bJumpToSelected]()
        ' not supported if a spectator
        If Not SPECTATOR_NO_OP() Then
            selNextSpecifiedBuilding REF_POWER_GEN, CBool(f.cap(0))
            triggerEventSelected
        End If
    Case FN_kf_SelectUnits              ' [selectionClass, selectionType, bOnScreen]()
        selDroidSelection selectedPlayer, CLng(f.cap(0)), CLng(f.cap(1)), CBool(f.cap(2))
    Case FN_kf_SelectNoGroupUnits       ' [selectionClass, selectionType, bOnScreen]()
        ' not supported if a spectator
        If Not SPECTATOR_NO_OP() Then activateNoGroup selectedPlayer, CLng(f.cap(1)), CLng(f.cap(0)), CBool(f.cap(2))
    Case FN_kf_SetDroid                 ' [order, state]() { kfsf_SetSelectedDroidsState(order, state); }
        kfsf_SetSelectedDroidsState CLng(f.cap(0)), CLng(f.cap(1))
    Case FN_kf_OrderDroid               ' [order]()
        kf_OrderDroid_lambda CLng(f.cap(0))
    End Select
End Function
