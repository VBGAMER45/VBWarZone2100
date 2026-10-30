Attribute VB_Name = "mission_c"
' src/mission.cpp, mission.h, missiondef.h: the missions (whole). struct MISSION is the mission_*
' globals (mission.apsDroidLists is mission_apsDroidLists, and so on: a list array inside a UDT could not
' be passed as a PerPlayer...Lists& argument). The lists are OBJ_LISTs (objlist_vb): a list's copy
' (operator=) gets nodes of its own (copyObjList), a std::move takes the nodes over (moveObjList),
' std::swap swaps the heads (swapObjList); the map's arrays (psMapTiles, psBlockMap, psAuxMap) move and
' swap by their SAFEARRAY pointers (swapArrayPointers), with psMapTiles_allocated; the gateway list is
' gateway_c's (gwSwapGateways). The mission timer, the transporter timer and the mission result screen
' are widgets (docs/CONVENTIONS.md, "Widgets"; the lambdas' ids are 11030 - 11034 and 13000 - 13499).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

' mission.h
Public Const MAX_NOGO_AREAS As Long = MAX_PLAYERS + 1
Public Const LIMBO_LANDING As Long = MAX_PLAYERS    ' used for Transporter Landing Zones 0-7 are for players, 8 = LIMBO_LANDING

Public Const INVALID_XY As Long = 512& * 127

' missiondef.h
Public Const SCR_LZ_COMPROMISED_TIME As Long = 999990
Public Const LZ_COMPROMISED_TIME As Long = 99999000
Public Const JS_LZ_COMPROMISED_TIME As Long = 99999

Public Type LANDING_ZONE
    x1 As Long                          ' uint8_t
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Public mission_type As Long             ' LEVEL_TYPE: defines which start and end functions to use - see levels_type in levels.h
Public mission_startTime As Long        ' UDWORD: time the mission started
Public mission_time As Long             ' SDWORD: how long the mission can last
' < 0 = no limit
Public mission_ETA As Long              ' SDWORD: time taken for reinforcements to arrive
' < 0 = none allowed
Public mission_cheatTime As Long        ' UDWORD: time the cheating started (mission time-wise!)

'original object lists
Public mission_apsStructLists(0 To MAX_PLAYERS - 1) As OBJ_LIST
Public mission_apsExtractorLists(0 To MAX_PLAYERS - 1) As OBJ_LIST
Public mission_apsDroidLists(0 To MAX_PLAYERS - 1) As OBJ_LIST
Public mission_apsFeatureList(0 To 0) As OBJ_LIST
Public mission_apsSensorList(0 To 0) As OBJ_LIST
Public mission_apsOilList(0 To 0) As OBJ_LIST
Public mission_apsFlagPosLists(0 To MAX_PLAYERS - 1) As OBJ_LIST
Public mission_asCurrentPower(0 To MAX_PLAYERS - 1) As Long  ' int32_t

Public apsLimboDroids(0 To MAX_PLAYERS - 1) As OBJ_LIST

Public offWorldKeepLists As Boolean

' the rest of struct MISSION (missiondef.h)
Public mission_psMapTiles() As MAPTILE_ ' std::unique_ptr<MAPTILE[]> psMapTiles: the original mapTiles
Public mission_psMapTiles_allocated As Boolean  ' (VB) mission.psMapTiles != nullptr
Public mission_mapWidth As Long         ' int32_t: the original mapWidth
Public mission_mapHeight As Long        ' int32_t: the original mapHeight
Public mission_psBlockMap() As Byte     ' psBlockMap[AUX_MAX], as map_c's psBlockMap (Erase = nullptr)
Public mission_psAuxMap() As Byte       ' psAuxMap[MAX_PLAYERS + AUX_MAX], as map_c's psAuxMap
Public mission_psGateways() As GATEWAY  ' GATEWAY_LIST psGateways: the gateway list (as gateway_c's)
Public mission_psGateways_size As Long
Public mission_scrollMinX As Long       ' int32_t: scroll coords for original map
Public mission_scrollMinY As Long
Public mission_scrollMaxX As Long
Public mission_scrollMaxY As Long
Public mission_homeLZ_X As Long         ' UWORD: selectedPlayer's LZ x and y
Public mission_homeLZ_Y As Long
Public mission_playerX As Long          ' SDWORD: original view position
Public mission_playerY As Long
' transporter entry/exit tiles (UWORD)
Public mission_iTranspEntryTileX(0 To MAX_PLAYERS - 1) As Long
Public mission_iTranspEntryTileY(0 To MAX_PLAYERS - 1) As Long
Public mission_iTranspExitTileX(0 To MAX_PLAYERS - 1) As Long
Public mission_iTranspExitTileY(0 To MAX_PLAYERS - 1) As Long

' the timers (mission.h)
Public Const IDTIMER_FORM As Long = 11000
Public Const IDTIMER_DISPLAY As Long = 11001
Public Const IDTRANTIMER_DISPLAY As Long = 11010

' Timer position defines
Public Const TIMER_Y As Long = 22

' Mission timer label position
Private Const TIMER_LABELX As Long = 15
Private Const TIMER_LABELY As Long = 0

' Transporter Timer form position
Private Const TRAN_FORM_X As Long = STAT_X
Private Const TRAN_FORM_Y As Long = TIMER_Y

' Transporter Timer position
Private Const TRAN_TIMER_X As Long = 4
Private Const TRAN_TIMER_Y As Long = TIMER_LABELY
Private Const TRAN_TIMER_WIDTH As Long = 25

' These are used for mission countdown
Private Const TEN_MINUTES As Long = 10& * 60 * GAME_TICKS_PER_SEC
Private Const FIVE_MINUTES As Long = 5& * 60 * GAME_TICKS_PER_SEC
Private Const FOUR_MINUTES As Long = 4& * 60 * GAME_TICKS_PER_SEC
Private Const THREE_MINUTES As Long = 3& * 60 * GAME_TICKS_PER_SEC
Private Const TWO_MINUTES As Long = 2& * 60 * GAME_TICKS_PER_SEC
Private Const ONE_MINUTE As Long = 1& * 60 * GAME_TICKS_PER_SEC
Private Const NOT_PLAYED_ONE As Long = &H1
Private Const NOT_PLAYED_TWO As Long = &H2
Private Const NOT_PLAYED_THREE As Long = &H4
Private Const NOT_PLAYED_FIVE As Long = &H8
Private Const NOT_PLAYED_TEN As Long = &H10
Private Const NOT_PLAYED_ACTIVATED As Long = &H20

Private Const OFFSCREEN_HEIGHT As Long = 600
Private Const EDGE_SIZE As Long = 1

Private Const MAXLIMBODROIDS As Long = 999

Private Const TRANSPORTER_REINFORCE_LEADIN As Long = 10& * GAME_TICKS_PER_SEC

'Where the Transporter lands for player 0 (sLandingZone[0]), and the rest are
'a list of areas that cannot be built on, used for landing the enemy transporters
Private sLandingZone(0 To MAX_NOGO_AREAS - 1) As LANDING_ZONE

'flag to indicate when the droids in a Transporter are flown to safety and not the next mission
Private bDroidsToSafety As Boolean

Private missionCountDown As Long        ' UBYTE
'flag to indicate whether the coded mission countdown is played
Private bPlayCountDown As Long          ' UBYTE

Public MissionResUp As Boolean          ' = false

Private g_iReinforceTime As Long        ' SDWORD = 0

' the mission result screen (mission.h, mission.cpp)
Public Const IDMISSIONRES_FORM As Long = 11002
Public Const IDMISSIONRES_QUIT As Long = 11007
Public Const IDMISSIONRES_SAVE As Long = 11006

Public Const MISSIONRES_X As Long = 20  ' pos & size of box.
Public Const MISSIONRES_Y As Long = 380
Public Const MISSIONRES_W As Long = 600
Public Const MISSIONRES_H As Long = 80

Public Const MISSIONRES_TITLE_X As Long = 20
Public Const MISSIONRES_TITLE_Y As Long = 20
Public Const MISSIONRES_TITLE_W As Long = 600
Public Const MISSIONRES_TITLE_H As Long = 40

Private Const IDMISSIONRES_TXT As Long = 11004
Private Const IDMISSIONRES_LOAD As Long = 11005
Private Const IDMISSIONRES_CONTINUE As Long = 11008
Private Const IDMISSIONRES_BACKFORM As Long = 11013
Private Const IDMISSIONRES_TITLE As Long = 11014

Private Const MISSION_1_X As Long = 5   ' pos of text options in box.
Private Const MISSION_1_Y As Long = 15
Private Const MISSION_2_X As Long = 5
Private Const MISSION_2_Y As Long = 35
Private Const MISSION_3_X As Long = 5
Private Const MISSION_3_Y As Long = 55

Private Const MISSION_TEXT_W As Long = MISSIONRES_W - 10
Private Const MISSION_TEXT_H As Long = 16

' (VB) the mission result screen's lambdas (cluster G: fork G5)
Public Const FN_intDisplayMissionBackDrop As Long = 11030
Public Const FN_deleteScoreDataToScreenCache As Long = 11031
Public Const FN_missionResBackForm_layout As Long = 11032
Public Const FN_missionResTitle_layout As Long = 11033
Public Const FN_missionResForm_layout As Long = 11034
' (VB) the timers' lambdas and callbacks
Public Const FN_intAddMissionTimer_calcLayout As Long = 13000
Public Const FN_intUpdateMissionTimer As Long = 13001
Public Const FN_intUpdateTransporterTimer As Long = 13002

'--------------------------------------------------------------------------
' (VB) the lists and the map's arrays, as the C++ moves, copies and swaps them

' (VB) dst = src (std::list's copy assignment): dst gets nodes of its own
Private Sub copyObjList(dst As OBJ_LIST, src As OBJ_LIST)
    Dim it As Long
    objlist_clear dst
    it = src.head
    Do While it <> 0
        objlist_push_back dst, lnObj(it)
        it = lnNext(it)
    Loop
End Sub

' (VB) dst = std::move(src): dst's nodes freed, src's taken over, src left empty
Private Sub moveObjList(dst As OBJ_LIST, src As OBJ_LIST)
    Dim emptyList As OBJ_LIST
    objlist_clear dst
    dst = src
    src = emptyList
End Sub

' (VB) std::swap(a, b) of two lists
Private Sub swapObjList(a As OBJ_LIST, b As OBJ_LIST)
    Dim t As OBJ_LIST
    t = a
    a = b
    b = t
End Sub

' (VB) list.reverse()
Private Sub reverseObjList(l As OBJ_LIST)
    Dim r As OBJ_LIST, it As Long
    it = l.head
    Do While it <> 0
        objlist_push_front r, lnObj(it)
        it = lnNext(it)
    Loop
    moveObjList l, r
End Sub

' (VB) std::swap of two dynamic arrays: their SAFEARRAY pointers (ppA, ppB: VarPtrArray of each)
Private Sub swapArrayPointers(ByVal ppA As Long, ByVal ppB As Long)
    Dim t As Long
    CopyMemory t, ByVal ppA, 4
    CopyMemory ByVal ppA, ByVal ppB, 4
    CopyMemory ByVal ppB, t, 4
End Sub

Private Sub swapLongs(a As Long, b As Long)
    Dim t As Long
    t = a
    a = b
    b = t
End Sub

' (VB) psMapTiles = std::move(mission.psMapTiles)
Private Sub mapTilesFromMission()
    swapArrayPointers VarPtrArray(psMapTiles), VarPtrArray(mission_psMapTiles)
    Erase mission_psMapTiles
    psMapTiles_allocated = mission_psMapTiles_allocated
    mission_psMapTiles_allocated = False
End Sub

' (VB) for (i ...) psBlockMap[i] = std::move(mission.psBlockMap[i]); the same for psAuxMap
Private Sub auxMapsFromMission()
    swapArrayPointers VarPtrArray(psBlockMap), VarPtrArray(mission_psBlockMap)
    Erase mission_psBlockMap
    swapArrayPointers VarPtrArray(psAuxMap), VarPtrArray(mission_psAuxMap)
    Erase mission_psAuxMap
End Sub

' (VB) the droid handle a widget's pUserData holds (void *: Empty and 0 are nullptr)
Private Function widgetUserDroid(ByVal psWidget As WIDGET) As Long
    Dim v As Variant
    WzFunc_assign v, psWidget.pUserData
    If IsObject(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    widgetUserDroid = CLng(v)
End Function

'--------------------------------------------------------------------------

'Remove soon-to-be illegal references to objects for some structures before going offWorld.
Private Sub resetHomeStructureObjects()
    Dim i As Long, it As Long, psStruct As Long
    For i = 0 To MAX_PLAYERS - 1
        it = apsStructLists(i).head
        Do While it <> 0
            psStruct = lnObj(it)
            ' (!psStruct->pFunctionality || !psStruct->pStructureType: continue)
            If aStruct(psStruct).pFunctionality <> 0 And aStruct(psStruct).pStructureType <> -1 Then
                If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REPAIR_FACILITY Then
                    ' REPAIR_FACILITY *psRepairFac = &psStruct->pFunctionality->repairFacility;
                    If aStruct(psStruct).functionality.repairFacility.psObj <> 0 Then
                        If aStruct(psStruct).functionality.repairFacility.state = RepairState_Repairing Then
                            droidRepairStopped castDroid(aStruct(psStruct).functionality.repairFacility.psObj), psStruct
                        End If
                        aStruct(psStruct).functionality.repairFacility.psObj = 0
                        aStruct(psStruct).functionality.repairFacility.state = RepairState_Idle
                    End If
                ElseIf asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REARM_PAD Then
                    ' REARM_PAD *psReArmPad = &psStruct->pFunctionality->rearmPad;
                    If aStruct(psStruct).functionality.rearmPad.psObj <> 0 Then
                        aStruct(psStruct).functionality.rearmPad.psObj = 0
                    End If
                End If
            End If
            it = lnNext(it)
        Loop
    Next
End Sub

'returns true if on an off world mission
Public Function missionIsOffworld() As Boolean
    missionIsOffworld = ((mission_type = LEVEL_TYPE_LDS_MKEEP) _
                         Or (mission_type = LEVEL_TYPE_LDS_MCLEAR) _
                         Or (mission_type = LEVEL_TYPE_LDS_MKEEP_LIMBO) _
                        )
End Function

'returns true if the correct type of mission for reinforcements
Public Function missionForReInforcements() As Boolean
    If mission_type = LEVEL_TYPE_LDS_CAMSTART Or missionIsOffworld() Or mission_type = LEVEL_TYPE_LDS_CAMCHANGE Then
        missionForReInforcements = True
    Else
        missionForReInforcements = False
    End If
End Function

'returns true if the correct type of mission and a reinforcement time has been set
Public Function missionCanReEnforce() As Boolean
    If mission_ETA >= 0 Then
        If missionForReInforcements() Then
            missionCanReEnforce = True
            Exit Function
        End If
    End If
    missionCanReEnforce = False
End Function

'returns true if the mission is a Limbo Expand mission
Public Function missionLimboExpand() As Boolean
    missionLimboExpand = (mission_type = LEVEL_TYPE_LDS_EXPAND_LIMBO)
End Function

' mission initialisation game code
Public Sub initMission()
    Dim inc As Long

    debug_ 272, LOG_SAVE, "initMission", "*** Init Mission ***"
    mission_type = LEVEL_TYPE_LDS_NONE
    For inc = 0 To MAX_PLAYERS - 1
        objlist_clear mission_apsStructLists(inc)
        objlist_clear mission_apsDroidLists(inc)
        objlist_clear mission_apsFlagPosLists(inc)
        objlist_clear mission_apsExtractorLists(inc)
        objlist_clear apsLimboDroids(inc)
    Next
    objlist_clear mission_apsFeatureList(0)
    objlist_clear mission_apsSensorList(0)
    objlist_clear mission_apsOilList(0)
    offWorldKeepLists = False
    mission_time = -1
    setMissionCountDown

    mission_ETA = -1
    mission_startTime = 0
    Erase mission_psGateways                ' just in case
    mission_psGateways_size = 0
    mission_mapHeight = 0
    mission_mapWidth = 0
    Erase mission_psBlockMap                ' for (auto &i : mission.psBlockMap) i.reset();
    Erase mission_psAuxMap                  ' for (auto &i : mission.psAuxMap) i.reset();

    'init all the landing zones
    initNoGoAreas

    bDroidsToSafety = False
    setPlayCountDown 1                      ' true

    'start as not cheating!
    mission_cheatTime = 0
End Sub

'this is called everytime the game is quit
Public Sub releaseMission()
    ' mission.apsDroidLists may contain some droids that have been transferred from one campaign to the next
    freeAllMissionDroids

    ' apsLimboDroids may contain some droids that have been saved at the end of one mission and not yet used
    freeAllLimboDroids
End Sub

'called to shut down when mid-mission on an offWorld map
Public Function missionShutDown() As Boolean
    Dim inc As Long

    debug_ 326, LOG_SAVE, "missionShutDown", "called, mission is %s", IIf(missionIsOffworld(), "off-world", "main map")
    If missionIsOffworld() Then
        'clear out the audio
        audio_StopAll

        freeAllDroids
        freeAllStructs
        freeAllFeatures
        freeAllFlagPositions
        releaseAllProxDisp
        gwShutDown

        For inc = 0 To MAX_PLAYERS - 1
            moveObjList apsDroidLists(inc), mission_apsDroidLists(inc)      ' (and mission.apsDroidLists[inc].clear())
            moveObjList apsStructLists(inc), mission_apsStructLists(inc)
            moveObjList apsFlagPosLists(inc), mission_apsFlagPosLists(inc)
            moveObjList apsExtractorLists(inc), mission_apsExtractorLists(inc)
        Next
        moveObjList apsFeatureList(0), mission_apsFeatureList(0)
        moveObjList apsSensorList(0), mission_apsSensorList(0)
        moveObjList apsOilList(0), mission_apsOilList(0)
        ' (mission.apsFeatureList[0].clear() and the others: empty after the move)

        mapTilesFromMission                 ' psMapTiles = std::move(mission.psMapTiles);
        mapWidth = mission_mapWidth
        mapHeight = mission_mapHeight
        auxMapsFromMission                  ' psBlockMap[i] = std::move(mission.psBlockMap[i]); psAuxMap[i] too
        gwSwapGateways mission_psGateways, mission_psGateways_size  ' std::swap(mission.psGateways, gwGetGateways());
    End If
    keybindShutdown
    ' sorry if this breaks something - but it looks like it's what should happen - John
    mission_type = LEVEL_TYPE_LDS_NONE

    missionShutDown = True
End Function


'on the PC - sets the countdown played flag
Public Sub setMissionCountDown()
    Dim timeRemaining As Long

    timeRemaining = mission_time - (gameTime - mission_startTime)
    If timeRemaining < 0 Then
        timeRemaining = 0
    End If

    ' Need to init the countdown played each time the mission time is changed
    missionCountDown = NOT_PLAYED_ONE Or NOT_PLAYED_TWO Or NOT_PLAYED_THREE Or NOT_PLAYED_FIVE Or NOT_PLAYED_TEN Or NOT_PLAYED_ACTIVATED

    If timeRemaining < TEN_MINUTES - 1 Then
        missionCountDown = missionCountDown And Not NOT_PLAYED_TEN
    End If
    If timeRemaining < FIVE_MINUTES - 1 Then
        missionCountDown = missionCountDown And Not NOT_PLAYED_FIVE
    End If
    If timeRemaining < THREE_MINUTES - 1 Then
        missionCountDown = missionCountDown And Not NOT_PLAYED_THREE
    End If
    If timeRemaining < TWO_MINUTES - 1 Then
        missionCountDown = missionCountDown And Not NOT_PLAYED_TWO
    End If
    If timeRemaining < ONE_MINUTE - 1 Then
        missionCountDown = missionCountDown And Not NOT_PLAYED_ONE
    End If
End Sub


Public Function startMission(ByVal missionType As Long, ByVal gameDetails As GameLoadDetails) As Boolean
    Dim loaded As Boolean

    loaded = True

    debug_ 419, LOG_SAVE, "startMission", "type %d", missionType

    ' Player has (obviously) not failed at the start
    setPlayerHasLost False
    setPlayerHasWon False

    ' and win/lose video is obviously not playing
    setScriptWinLoseVideo PLAY_NONE

    ' this inits the flag so that 'reinforcements have arrived' message isn't played for the first transporter load
    initFirstTransporterFlag

    If mission_type <> LEVEL_TYPE_LDS_NONE Then
        'mission type gets set to none when you have returned from a mission
        'so don't want to go another mission when already on one! - so ignore
        debug_ 435, LOG_SAVE, "startMission", "Already on a mission"
        startMission = True
        Exit Function
    End If

    initEffectsSystem

    'load the game file for all types of mission except a Between Mission
    If missionType <> LEVEL_TYPE_LDS_BETWEEN Then
        loadGameInit gameDetails
    End If

    'all proximity messages are removed between missions now
    releaseAllProxDisp

    Select Case missionType
    Case LEVEL_TYPE_LDS_CAMSTART
        If Not startMissionCampaignStart(gameDetails) Then
            loaded = False
        End If
    Case LEVEL_TYPE_LDS_MKEEP, LEVEL_TYPE_LDS_MKEEP_LIMBO
        If Not startMissionOffKeep(gameDetails) Then
            loaded = False
        End If
    Case LEVEL_TYPE_LDS_BETWEEN
        'do anything?
        If Not startMissionBetween() Then
            loaded = False
        End If
    Case LEVEL_TYPE_LDS_CAMCHANGE
        If Not startMissionCampaignChange(gameDetails) Then
            loaded = False
        End If
    Case LEVEL_TYPE_LDS_EXPAND
        If Not startMissionCampaignExpand(gameDetails) Then
            loaded = False
        End If
    Case LEVEL_TYPE_LDS_EXPAND_LIMBO
        If Not startMissionCampaignExpandLimbo(gameDetails) Then
            loaded = False
        End If
    Case LEVEL_TYPE_LDS_MCLEAR
        If Not startMissionOffClear(gameDetails) Then
            loaded = False
        End If
    Case Else
        'error!
        debug_ 498, LOG_ERROR, "startMission", "Unknown Mission Type"

        loaded = False
    End Select

    If Not loaded Then
        debug_ 505, LOG_ERROR, "startMission", "Failed to start mission, missiontype = %d, game, %s", missionType, gameDetails.filePath
        startMission = False
        Exit Function
    End If

    mission_type = missionType

    If missionIsOffworld() Then
        'add what power have got from the home base
        adjustMissionPower
    End If

    If missionCanReEnforce() Then
        'add mission timer - if necessary
        addMissionTimerInterface

        'add Transporter Timer
        addTransporterTimerInterface
    End If

    scoreInitSystem

    startMission = True
End Function


' initialise the mission stuff for a save game
Public Function startMissionSave(ByVal missionType As Long) As Boolean
    mission_type = missionType

    startMissionSave = True
End Function


'checks the time has been set and then adds the timer if not already on
'the display
Public Sub addMissionTimerInterface()
    'don't add if the timer hasn't been set
    If mission_time < 0 And Not challengeActive Then
        Exit Sub
    End If

    'check timer is not already on the screen
    If widgGetFromID(psWScreen, IDTIMER_FORM) Is Nothing Then
        intAddMissionTimer
    End If
End Sub

'checks that the timer has been set and that a Transporter exists before
'adding the timer button
Public Sub addTransporterTimerInterface()
    Dim psTransporter As Long, bAddInterface As Boolean, psForm As WIDGET, it As Long, psDroid As Long

    psTransporter = 0
    bAddInterface = False

    'check if reinforcements are allowed
    If mission_ETA >= 0 And ULT(selectedPlayer, MAX_PLAYERS) Then
        'check the player has at least one Transporter back at base
        it = mission_apsDroidLists(selectedPlayer).head
        Do While it <> 0
            psDroid = lnObj(it)
            If DROID_isTransporter(psDroid) Then
                psTransporter = psDroid
                Exit Do
            End If
            it = lnNext(it)
        Loop
        If psTransporter <> 0 Then
            bAddInterface = True

            ' Check that neither the timer nor the launch button are on screen
            If widgGetFromID(psWScreen, IDTRANTIMER_BUTTON) Is Nothing And widgGetFromID(psWScreen, IDTRANS_LAUNCH) Is Nothing Then
                intAddTransporterTimer
            End If

            'set the data for the transporter timer
            widgSetUserData psWScreen, IDTRANTIMER_DISPLAY, psTransporter

            ' lock the button if necessary
            If transporterFlying(psTransporter) Then
                ' disable the form so can't add any more droids into the transporter
                Set psForm = widgGetFromID(psWScreen, IDTRANTIMER_BUTTON)
                If Not psForm Is Nothing Then
                    psForm.setState WBUT_LOCK
                End If
            End If
        End If
    End If
    ' if criteria not met
    If Not bAddInterface Then
        ' make sure its not there!
        intRemoveTransporterTimer
    End If
End Sub

' fly in transporters at start of level
Public Sub missionFlyTransportersIn(ByVal iPlayer As Long, ByVal bTrackTransporter As Boolean)
    Dim iX As Long, iY As Long, iZ As Long, iLandX As Long, iLandY As Long
    Dim it As Long, nxt As Long, psTransporter As Long, iDx As Long, iDy As Long

    If Not (iPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "iPlayer < MAX_PLAYERS", "src/mission.cpp:620", "missionFlyTransportersIn", "Flying nonexistent player %d's transporters in", iPlayer: Exit Sub

    iLandX = getLandingX(iPlayer)
    iLandY = getLandingY(iPlayer)
    missionGetTransporterEntry iPlayer, iX, iY
    iZ = (map_Height(iX, iY) + OFFSCREEN_HEIGHT) And &HFFFF&

    'get the droids for the mission
    ' (mutating_list_iterate: the next node is taken before the handler)
    it = mission_apsDroidLists(iPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psTransporter = lnObj(it)
        If aDroid(psTransporter).droidType = DROID_SUPERTRANSPORTER Then
            ' Check that this transporter actually contains some droids
            If aDroid(psTransporter).psGroup <> 0 Then
                If aGroup(aDroid(psTransporter).psGroup).refCount > 1 Then
                    ' Remove map information from previous map
                    Erase aDroid(psTransporter).watchedTiles     ' watchedTiles.clear()
                    aDroid(psTransporter).watchedTiles_size = 0

                    ' Remove out of stored list and add to current Droid list
                    If droidRemove(psTransporter, mission_apsDroidLists) Then
                        ' Do not want to add it unless managed to remove it from the previous list
                        addDroid psTransporter, apsDroidLists
                    End If

                    ' set start position
                    aDroid(psTransporter).pos.x = iX
                    aDroid(psTransporter).pos.y = iY
                    aDroid(psTransporter).pos.z = iZ

                    ' set start direction
                    iDx = iLandX - iX
                    iDy = iLandY - iY

                    aDroid(psTransporter).rot.direction = iAtan2(iDx, iDy)

                    ' Camera track requested and it's the selected player.
                    If (bTrackTransporter = True) And (iPlayer = selectedPlayer) Then
                        ' deselect all droids
                        selDroidDeselect selectedPlayer

                        If getWarCamStatus() Then
                            camToggleStatus
                        End If

                        ' select transporter
                        aDroid(psTransporter).selected = 1
                        camToggleStatus
                    End If

                    'little hack to ensure all Transporters are fully repaired by time enter world
                    aDroid(psTransporter).body = aDroid(psTransporter).originalBody

                    ' set fly-in order
                    orderDroidLoc psTransporter, DORDER_TRANSPORTIN, iLandX, iLandY, ModeImmediate

                    audio_PlayObjDynamicTrack psTransporter, ID_SOUND_BLIMP_FLIGHT, CB_moveCheckDroidMovingAndVisible

                    'only want to fly one transporter in at a time now - AB 14/01/99
                    Exit Do
                End If
            End If
        End If
        it = nxt
    Loop
End Sub

' Saves the necessary data when moving from a home base Mission to an OffWorld mission
Public Sub saveMissionData()
    Dim bRepairExists As Boolean, bRearmPadExists As Boolean, it As Long, nxt As Long, it2 As Long
    Dim psStruct As Long, psDroid As Long, psStructBeingBuilt As Long, vtolAndPadsExist As Boolean, inc As Long

    bRepairExists = False
    bRearmPadExists = False

    debug_ 696, LOG_SAVE, "saveMissionData", "called"

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:698", "saveMissionData", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer

    'clear out the audio
    audio_StopAll

    'set any structures currently being built to completed for the selected player
    ' (mutating_list_iterate)
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psStruct = lnObj(it)
        If aStruct(psStruct).status = SS_BEING_BUILT Then
            'find a droid working on it
            it2 = apsDroidLists(selectedPlayer).head
            Do While it2 <> 0
                psDroid = lnObj(it2)
                psStructBeingBuilt = orderStateObj(psDroid, DORDER_BUILD)
                If psStructBeingBuilt <> 0 Then
                    If psStructBeingBuilt = psStruct Then
                        ' just give it all its build points
                        structureBuild psStruct, 0, structureBuildPointsToCompletion(psStruct)
                        'don't bother looking for any other droids working on it
                        Exit Do
                    End If
                End If
                it2 = lnNext(it2)
            Loop
        End If
        'check if have a completed repair facility or rearming pad on home world
        If aStruct(psStruct).status = SS_BUILT And aStruct(psStruct).pStructureType <> -1 Then
            If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REPAIR_FACILITY Then
                bRepairExists = True
            ElseIf asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REARM_PAD Then
                bRearmPadExists = True
            End If
        End If
        it = nxt
    Loop

    'repair and rearm all droids back at home base if have a repair facility or rearming pad
    If bRepairExists Or bRearmPadExists Then
        it = apsDroidLists(selectedPlayer).head
        Do While it <> 0
            psDroid = lnObj(it)
            vtolAndPadsExist = (DROID_isVtol(psDroid) And bRearmPadExists)
            If (bRepairExists Or vtolAndPadsExist) And DROID_isDamaged(psDroid) Then
                aDroid(psDroid).body = aDroid(psDroid).originalBody
            End If
            If vtolAndPadsExist Then
                fillVtolDroid psDroid
            End If
            it = lnNext(it)
        Loop
    End If

    'clear droid orders for all droids except constructors still building
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        psStructBeingBuilt = orderStateObj(psDroid, DORDER_BUILD)
        If psStructBeingBuilt <> 0 Then
            If aStruct(psStructBeingBuilt).status = SS_BUILT Then
                orderDroid psDroid, DORDER_STOP, ModeImmediate
            End If
        Else
            orderDroid psDroid, DORDER_STOP, ModeImmediate
        End If
        resetObjectAnimationState psDroid
        it = lnNext(it)
    Loop

    resetHomeStructureObjects 'get rid of soon-to-be illegal references of droids in repair facilities and rearming pads.

    'save the mission data
    swapArrayPointers VarPtrArray(mission_psMapTiles), VarPtrArray(psMapTiles)   ' mission.psMapTiles = std::move(psMapTiles);
    Erase psMapTiles
    mission_psMapTiles_allocated = psMapTiles_allocated
    psMapTiles_allocated = False
    mission_mapWidth = mapWidth
    mission_mapHeight = mapHeight
    ' for (...) mission.psBlockMap[i] = std::move(psBlockMap[i]); and psAuxMap
    swapArrayPointers VarPtrArray(mission_psBlockMap), VarPtrArray(psBlockMap)
    Erase psBlockMap
    swapArrayPointers VarPtrArray(mission_psAuxMap), VarPtrArray(psAuxMap)
    Erase psAuxMap
    mission_scrollMinX = scrollMinX
    mission_scrollMinY = scrollMinY
    mission_scrollMaxX = scrollMaxX
    mission_scrollMaxY = scrollMaxY
    gwSwapGateways mission_psGateways, mission_psGateways_size  ' std::swap(mission.psGateways, gwGetGateways());
    ' save the selectedPlayer's LZ
    mission_homeLZ_X = getLandingX(selectedPlayer)
    mission_homeLZ_Y = getLandingY(selectedPlayer)

    For inc = 0 To MAX_PLAYERS - 1
        copyObjList mission_apsStructLists(inc), apsStructLists(inc)
        copyObjList mission_apsDroidLists(inc), apsDroidLists(inc)
        copyObjList mission_apsFlagPosLists(inc), apsFlagPosLists(inc)
        copyObjList mission_apsExtractorLists(inc), apsExtractorLists(inc)
    Next
    copyObjList mission_apsFeatureList(0), apsFeatureList(0)
    copyObjList mission_apsSensorList(0), apsSensorList(0)
    copyObjList mission_apsOilList(0), apsOilList(0)

    mission_playerX = playerPos.p.x
    mission_playerY = playerPos.p.z

    'save the power settings
    saveMissionPower

    'init before loading in the new game
    initFactoryNumFlag

    'clear all the effects from the map
    initEffectsSystem

    resizeRadar
End Sub

'
'    This routine frees the memory for the offworld mission map (in the call to mapShutdown)
'
'    - so when this routine is called we must still be set to the offworld map data
'    i.e. We shoudn't have called SwapMissionPointers()
'
Private Sub restoreMissionData()
    Dim inc As Long, it As Long

    debug_ 832, LOG_SAVE, "restoreMissionData", "called"

    'clear out the audio
    audio_StopAll

    'clear all the lists
    proj_FreeAllProjectiles
    freeAllDroids
    freeAllStructs
    freeAllFeatures
    freeAllFlagPositions
    gwShutDown
    If game.typ <> LEVEL_TYPE_CAMPAIGN Then
        ASSERT_FAILURE "false", "src/mission.cpp:846", "restoreMissionData", "game type isn't campaign, but we are in a campaign game!"
        game.typ = LEVEL_TYPE_CAMPAIGN  ' fix the issue, since it is obviously a bug
    End If
    'restore the game pointers
    For inc = 0 To MAX_PLAYERS - 1
        moveObjList apsDroidLists(inc), mission_apsDroidLists(inc)      ' (and mission.apsDroidLists[inc].clear())
        it = apsDroidLists(inc).head
        Do While it <> 0
            aDroid(lnObj(it)).died = 0  'make sure the died flag is not set
            it = lnNext(it)
        Loop

        moveObjList apsStructLists(inc), mission_apsStructLists(inc)

        moveObjList apsFlagPosLists(inc), mission_apsFlagPosLists(inc)

        moveObjList apsExtractorLists(inc), mission_apsExtractorLists(inc)
    Next
    moveObjList apsFeatureList(0), mission_apsFeatureList(0)
    moveObjList apsSensorList(0), mission_apsSensorList(0)
    moveObjList apsOilList(0), mission_apsOilList(0)
    objlist_clear mission_apsFeatureList(0)
    objlist_clear mission_apsSensorList(0)
    objlist_clear apsOilList(0)             ' (the C++ clears apsOilList[0] here, not mission.apsOilList[0])
    'swap mission data over

    mapTilesFromMission                     ' psMapTiles = std::move(mission.psMapTiles);

    mapWidth = mission_mapWidth
    mapHeight = mission_mapHeight
    auxMapsFromMission                      ' psBlockMap[i] = std::move(mission.psBlockMap[i]); psAuxMap[i] too
    scrollMinX = mission_scrollMinX
    scrollMinY = mission_scrollMinY
    scrollMaxX = mission_scrollMaxX
    scrollMaxY = mission_scrollMaxY
    gwSwapGateways mission_psGateways, mission_psGateways_size  ' std::swap(mission.psGateways, gwGetGateways());
    'and clear the mission pointers
    ' (mission.psMapTiles = nullptr: moved out already)
    mission_mapWidth = 0
    mission_mapHeight = 0
    mission_scrollMinX = 0
    mission_scrollMinY = 0
    mission_scrollMaxX = 0
    mission_scrollMaxY = 0
    Erase mission_psGateways                ' mission.psGateways.clear();
    mission_psGateways_size = 0

    'reset the current structure lists
    setCurrentStructQuantity False

    initFactoryNumFlag
    resetFactoryNumFlag

    offWorldKeepLists = False

    resizeRadar
End Sub

'Saves the necessary data when moving from one mission to a limbo expand Mission
Public Sub saveMissionLimboData()
    Dim it As Long, nxt As Long, psDroid As Long, psStruct As Long

    debug_ 917, LOG_SAVE, "saveMissionLimboData", "called"

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:919", "saveMissionLimboData", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer

    'clear out the audio
    audio_StopAll

    ' before copy the pointers over check selectedPlayer's mission.droids since
    ' there might be some from the previous camapign
    processPreviousCampDroids

    ' move droids properly - does all the clean up code
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psDroid = lnObj(it)
        If droidRemove(psDroid, apsDroidLists) Then
            addDroid psDroid, mission_apsDroidLists
        End If
        it = nxt
    Loop
    objlist_clear apsDroidLists(selectedPlayer)

    ' any selectedPlayer's factories/research need to be put on holdProduction/holdresearch
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        psStruct = lnObj(it)
        If psStruct <> 0 Then
            If STRUCTURE_isFactory(psStruct) Then
                holdProduction psStruct, ModeImmediate
            ElseIf asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESEARCH Then
                holdResearch psStruct, ModeImmediate
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

'this is called via a script function to place the Limbo droids once the mission has started
Public Sub placeLimboDroids()
    Dim it As Long, nxt As Long, psDroid As Long, droidX As Long, droidY As Long, pickRes As PICKTILE

    debug_ 960, LOG_SAVE, "placeLimboDroids", "called"

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:962", "placeLimboDroids", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer

    ' Copy the droids across for the selected Player
    it = apsLimboDroids(selectedPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psDroid = lnObj(it)
        If droidRemove(psDroid, apsLimboDroids) Then
            addDroid psDroid, apsDroidLists
            'KILL OFF TRANSPORTER - should never be one but....
            If DROID_isTransporter(psDroid) Then
                vanishDroid psDroid
                GoTo continue_
            End If
            'set up location for each of the droids
            droidX = map_coord(getLandingX(LIMBO_LANDING))
            droidY = map_coord(getLandingY(LIMBO_LANDING))
            pickRes = pickHalfATile(droidX, droidY, LOOK_FOR_EMPTY_TILE)
            If pickRes = NO_FREE_TILE Then
                ASSERT_FAILURE "false", "src/mission.cpp:985", "placeLimboDroids", "placeLimboUnits: Unable to find a free location"
            End If
            aDroid(psDroid).pos.x = world_coord(droidX) And &HFFFF&
            aDroid(psDroid).pos.y = world_coord(droidY) And &HFFFF&
            If Not (worldOnMap(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)) Then ASSERT_FAILURE "worldOnMap(psDroid->pos.x, psDroid->pos.y)", "src/mission.cpp:989", "placeLimboDroids", "limbo droid is not on the map"
            aDroid(psDroid).pos.z = map_Height(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)
            updateDroidOrientation psDroid
            aDroid(psDroid).selected = 0
            'this is mainly for VTOLs
            setDroidBase psDroid, 0
            'initialise the movement data
            initDroidMovement psDroid
            'make sure the died flag is not set
            aDroid(psDroid).died = 0
            'update visibility
            visTilesUpdate psDroid
        Else
            ASSERT_FAILURE "false", "src/mission.cpp:1004", "placeLimboDroids", "placeLimboUnits: Unable to remove unit from Limbo list"
        End If
continue_:
        it = nxt
    Loop
End Sub

'restores the necessary data on completion of a Limbo Expand mission
Private Sub restoreMissionLimboData()
    Dim it As Long, nxt As Long, psDroid As Long

    debug_ 1013, LOG_SAVE, "restoreMissionLimboData", "called"

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:1015", "restoreMissionLimboData", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer

    'the droids stored in the mission droid list need to be added back
    'into the current droid list
    it = mission_apsDroidLists(selectedPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psDroid = lnObj(it)
        'remove out of stored list and add to current Droid list
        If droidRemove(psDroid, mission_apsDroidLists) Then
            addDroid psDroid, apsDroidLists
            'reset droid orders
            orderDroid psDroid, DORDER_STOP, ModeImmediate
            'the location of the droid should be valid!
            If aDroid(psDroid).pos.x <> INVALID_XY And aDroid(psDroid).pos.y <> INVALID_XY Then
                visTilesUpdate psDroid  'update visibility
            End If
        End If
        it = nxt
    Loop
    If Not (mission_apsDroidLists(selectedPlayer).size = 0) Then ASSERT_FAILURE "mission.apsDroidLists[selectedPlayer].empty()", "src/mission.cpp:1035", "restoreMissionLimboData", "list should be empty"
End Sub

'Saves the necessary data when moving from one campaign to the start of the
'next - saves out the list of droids for the selected player
Public Sub saveCampaignData()
    Dim it As Long, nxt As Long, it2 As Long, nxt2 As Long, psDroid As Long, psCurr As Long, psSafeDroid As Long, inc As Long

    debug_ 1042, LOG_SAVE, "saveCampaignData", "called"

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:1044", "saveCampaignData", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer

    ' If the droids have been moved to safety then get any Transporters that exist
    If getDroidsToSafetyFlag() Then
        ' Move any Transporters into the mission list
        it = apsDroidLists(selectedPlayer).head
        Do While it <> 0
            nxt = lnNext(it)
            psDroid = lnObj(it)
            If DROID_isTransporter(psDroid) Then
                ' Empty the transporter into the mission list
                If Not (aDroid(psDroid).psGroup <> 0) Then ASSERT_FAILURE "psDroid->psGroup != nullptr", "src/mission.cpp:1055", "saveCampaignData", "Transporter does not have a group": GoTo continue1

                it2 = aGroup(aDroid(psDroid).psGroup).psList.head
                Do While it2 <> 0
                    nxt2 = lnNext(it2)
                    psCurr = lnObj(it2)
                    If psCurr = psDroid Then
                        Exit Do
                    End If
                    ' Remove it from the transporter group
                    DROID_GROUP_remove aDroid(psDroid).psGroup, psCurr
                    ' Cam change add droid
                    aDroid(psCurr).pos.x = INVALID_XY
                    aDroid(psCurr).pos.y = INVALID_XY
                    ' Add it back into current droid lists
                    addDroid psCurr, mission_apsDroidLists
                    it2 = nxt2
                Loop
                ' Remove the transporter from the current list
                If droidRemove(psDroid, apsDroidLists) Then
                    'cam change add droid
                    aDroid(psDroid).pos.x = INVALID_XY
                    aDroid(psDroid).pos.y = INVALID_XY
                    addDroid psDroid, mission_apsDroidLists
                End If
            End If
continue1:
            it = nxt
        Loop
    Else
        ' Reserve the droids for selected player for start of next campaign
        moveObjList mission_apsDroidLists(selectedPlayer), apsDroidLists(selectedPlayer)  ' (and apsDroidLists[selectedPlayer].clear())
        it = mission_apsDroidLists(selectedPlayer).head
        Do While it <> 0
            psDroid = lnObj(it)
            'cam change add droid
            aDroid(psDroid).pos.x = INVALID_XY
            aDroid(psDroid).pos.y = INVALID_XY
            it = lnNext(it)
        Loop
    End If

    'if the droids have been moved to safety then get any Transporters that exist
    If getDroidsToSafetyFlag() Then
        'now that every unit for the selected player has been moved into the
        'mission list - reverse it and fill the transporter with the first ten units
        reverseObjList mission_apsDroidLists(selectedPlayer)

        'find the *first* transporter
        it = mission_apsDroidLists(selectedPlayer).head
        Do While it <> 0
            nxt = lnNext(it)
            psDroid = lnObj(it)
            If DROID_isTransporter(psDroid) Then
                'fill it with droids from the mission list
                it2 = mission_apsDroidLists(selectedPlayer).head
                Do While it2 <> 0
                    nxt2 = lnNext(it2)
                    psSafeDroid = lnObj(it2)
                    If psSafeDroid <> psDroid Then
                        'add to the Transporter, checking for when full
                        If checkTransporterSpace(psDroid, psSafeDroid) Then
                            If droidRemove(psSafeDroid, mission_apsDroidLists) Then
                                DROID_GROUP_add aDroid(psDroid).psGroup, psSafeDroid
                            End If
                        Else
                            Exit Do
                        End If
                    End If
                    it2 = nxt2
                Loop
                'only want to fill one transporter
                Exit Do
            End If
            it = nxt
        Loop
    End If

    'clear all other droids
    For inc = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(inc).head
        Do While it <> 0
            nxt = lnNext(it)
            vanishDroid lnObj(it)
            it = nxt
        Loop
    Next

    'clear out the audio
    audio_StopAll

    'clear all other memory
    freeAllStructs
    freeAllFeatures
End Sub


'start an off world mission - clearing the object lists
Private Function startMissionOffClear(ByVal gameToLoad As GameLoadDetails) As Boolean
    debug_ 1159, LOG_SAVE, "startMissionOffClear", "called for %s", gameToLoad.filePath

    saveMissionData

    'load in the new game clearing the lists
    If Not loadGame(gameToLoad, Not KEEPOBJECTS, Not FREEMEM) Then
        startMissionOffClear = False
        Exit Function
    End If

    offWorldKeepLists = False

    ' The message should have been played at the between stage
    missionCountDown = missionCountDown And Not NOT_PLAYED_ACTIVATED

    startMissionOffClear = True
End Function

'start an off world mission - keeping the object lists
Private Function startMissionOffKeep(ByVal gameToLoad As GameLoadDetails) As Boolean
    debug_ 1180, LOG_SAVE, "startMissionOffKeep", "called for %s", gameToLoad.filePath
    saveMissionData

    'load in the new game clearing the lists
    If Not loadGame(gameToLoad, Not KEEPOBJECTS, Not FREEMEM) Then
        startMissionOffKeep = False
        Exit Function
    End If

    offWorldKeepLists = True

    ' The message should have been played at the between stage
    missionCountDown = missionCountDown And Not NOT_PLAYED_ACTIVATED

    startMissionOffKeep = True
End Function

Private Function startMissionCampaignStart(ByVal gameToLoad As GameLoadDetails) As Boolean
    debug_ 1199, LOG_SAVE, "startMissionCampaignStart", "called for %s", gameToLoad.filePath

    ' Clear out all intelligence screen messages
    freeMessages

    ' Check no units left with any settings that are invalid
    clearCampaignUnits

    ' Load in the new game details
    If Not loadGame(gameToLoad, Not KEEPOBJECTS, FREEMEM) Then
        startMissionCampaignStart = False
        Exit Function
    End If

    offWorldKeepLists = False

    startMissionCampaignStart = True
End Function

Private Function startMissionCampaignChange(ByVal gameToLoad As GameLoadDetails) As Boolean
    ' Clear out all intelligence screen messages
    freeMessages

    ' Cancel any research active in a lab
    CancelAllResearch selectedPlayer

    ' Check no units left with any settings that are invalid
    clearCampaignUnits

    ' Clear out the production run between campaigns
    changeProductionPlayer selectedPlayer And &HFF&

    saveCampaignData

    'load in the new game details
    If Not loadGame(gameToLoad, Not KEEPOBJECTS, Not FREEMEM) Then
        startMissionCampaignChange = False
        Exit Function
    End If

    offWorldKeepLists = False

    startMissionCampaignChange = True
End Function

Private Function startMissionCampaignExpand(ByVal gameToLoad As GameLoadDetails) As Boolean
    'load in the new game details
    If Not loadGame(gameToLoad, KEEPOBJECTS, Not FREEMEM) Then
        startMissionCampaignExpand = False
        Exit Function
    End If

    offWorldKeepLists = False
    startMissionCampaignExpand = True
End Function

Private Function startMissionCampaignExpandLimbo(ByVal gameToLoad As GameLoadDetails) As Boolean
    saveMissionLimboData

    'load in the new game details
    If Not loadGame(gameToLoad, KEEPOBJECTS, Not FREEMEM) Then
        startMissionCampaignExpandLimbo = False
        Exit Function
    End If

    offWorldKeepLists = False

    startMissionCampaignExpandLimbo = True
End Function

Private Function startMissionBetween() As Boolean
    offWorldKeepLists = False

    startMissionBetween = True
End Function

'check no units left with any settings that are invalid
Private Sub clearCampaignUnits()
    Dim it As Long, psDroid As Long

    If selectedPlayer >= MAX_PLAYERS Then Exit Sub

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        orderDroid psDroid, DORDER_STOP, ModeImmediate
        setDroidBase psDroid, 0
        visRemoveVisibilityOffWorld psDroid
        checkDroid psDroid, "src/mission.cpp:1289", "clearCampaignUnits", max_check_object_recursion
        it = lnNext(it)
    Loop
End Sub

'This deals with droids at the end of an offworld mission
Private Sub processMission()
    Dim it As Long, nxt As Long, psDroid As Long, droidX As Long, droidY As Long, pickRes As PICKTILE, x As Long, y As Long

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:1296", "processMission", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer

    'and the rest on the mission map  - for now?
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psDroid = lnObj(it)
        'reset order - do this to all the droids that are returning from offWorld
        orderDroid psDroid, DORDER_STOP, ModeImmediate
        ' clean up visibility
        visRemoveVisibility psDroid
        'remove out of stored list and add to current Droid list
        If droidRemove(psDroid, apsDroidLists) Then
            addDroid psDroid, mission_apsDroidLists
            droidX = getHomeLandingX()
            droidY = getHomeLandingY()
            ' Swap the droid and map pointers
            swapMissionPointers

            pickRes = pickHalfATile(droidX, droidY, LOOK_FOR_EMPTY_TILE)
            If Not (pickRes <> NO_FREE_TILE) Then ASSERT_FAILURE "pickRes != NO_FREE_TILE", "src/mission.cpp:1319", "processMission", "processMission: Unable to find a free location"
            x = world_coord(droidX) And &HFFFF&
            y = world_coord(droidY) And &HFFFF&
            droidSetPosition psDroid, x, y
            If Not (worldOnMap(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)) Then ASSERT_FAILURE "worldOnMap(psDroid->pos.x, psDroid->pos.y)", "src/mission.cpp:1323", "processMission", "the droid is not on the map"
            updateDroidOrientation psDroid
            ' Swap the droid and map pointers back again
            swapMissionPointers
            aDroid(psDroid).selected = 0
            ' This is mainly for VTOLs
            setDroidBase psDroid, 0
        End If
        it = nxt
    Loop
End Sub


'This deals with droids at the end of an offworld Limbo mission
Private Sub processMissionLimbo()
    Dim numDroidsAddedToLimboList As Long, it As Long, nxt As Long, psDroid As Long
    numDroidsAddedToLimboList = 0

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:1343", "processMissionLimbo", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer

    'all droids (for selectedPlayer only) are placed into the limbo list
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psDroid = lnObj(it)
        'KILL OFF TRANSPORTER - should never be one but....
        If DROID_isTransporter(psDroid) Then
            vanishDroid psDroid
        Else
            If numDroidsAddedToLimboList >= MAXLIMBODROIDS Then     ' any room in limbo list
                vanishDroid psDroid
            Else
                ' Remove out of stored list and add to current Droid list
                If droidRemove(psDroid, apsDroidLists) Then
                    ' Limbo list invalidate XY
                    aDroid(psDroid).pos.x = INVALID_XY
                    aDroid(psDroid).pos.y = INVALID_XY
                    addDroid psDroid, apsLimboDroids
                    ' This is mainly for VTOLs
                    setDroidBase psDroid, 0
                    orderDroid psDroid, DORDER_STOP, ModeImmediate
                    numDroidsAddedToLimboList = numDroidsAddedToLimboList + 1
                End If
            End If
        End If
        it = nxt
    Loop
End Sub

'switch the pointers for the map and droid lists so that droid placement
' and orientation can occur on the map they will appear on
' NOTE: This is one huge hack for campaign games!
' Pay special attention on what is getting swapped!
Public Sub swapMissionPointers()
    Dim inc As Long, b As Boolean

    debug_ 1385, LOG_SAVE, "swapMissionPointers", "called"

    swapArrayPointers VarPtrArray(psMapTiles), VarPtrArray(mission_psMapTiles)
    b = psMapTiles_allocated
    psMapTiles_allocated = mission_psMapTiles_allocated
    mission_psMapTiles_allocated = b
    swapLongs mapWidth, mission_mapWidth
    swapLongs mapHeight, mission_mapHeight
    ' for (...) std::swap(psBlockMap[i], mission.psBlockMap[i]); and psAuxMap
    swapArrayPointers VarPtrArray(psBlockMap), VarPtrArray(mission_psBlockMap)
    swapArrayPointers VarPtrArray(psAuxMap), VarPtrArray(mission_psAuxMap)
    'swap gateway zones
    gwSwapGateways mission_psGateways, mission_psGateways_size  ' std::swap(mission.psGateways, gwGetGateways());
    swapLongs scrollMinX, mission_scrollMinX
    swapLongs scrollMinY, mission_scrollMinY
    swapLongs scrollMaxX, mission_scrollMaxX
    swapLongs scrollMaxY, mission_scrollMaxY
    For inc = 0 To MAX_PLAYERS - 1
        swapObjList apsDroidLists(inc), mission_apsDroidLists(inc)
        swapObjList apsStructLists(inc), mission_apsStructLists(inc)
        swapObjList apsFlagPosLists(inc), mission_apsFlagPosLists(inc)
        swapObjList apsExtractorLists(inc), mission_apsExtractorLists(inc)
    Next
    swapObjList apsFeatureList(0), mission_apsFeatureList(0)
    swapObjList apsSensorList(0), mission_apsSensorList(0)
    swapObjList apsOilList(0), mission_apsOilList(0)
End Sub

Public Sub endMission()
    If mission_type <> LEVEL_TYPE_LDS_BETWEEN Then
        releaseAllFlicMessages apsMessages 'Needed to remove mission objectives from offworld missions
        releaseObjectives = True
    Else
        releaseObjectives = False
    End If

    If mission_type = LEVEL_TYPE_LDS_NONE Then
        'can't go back any further!!
        debug_ 1431, LOG_SAVE, "endMission", "Already returned from mission"
        Exit Sub
    End If

    Select Case mission_type
    Case LEVEL_TYPE_LDS_CAMSTART
        'any transporters that are flying in need to be emptied
        emptyTransporters False
        'when loading in a save game mid cam2a or cam3a it is loaded as a camstart
        endMissionCamChange
        '
        '    This is called at the end of every campaign mission
        '
    Case LEVEL_TYPE_LDS_MKEEP
        'any transporters that are flying in need to be emptied
        emptyTransporters True
        endMissionOffKeep
    Case LEVEL_TYPE_LDS_EXPAND, LEVEL_TYPE_LDS_BETWEEN
        '
        '    This is called at the end of every campaign mission
        '

    Case LEVEL_TYPE_LDS_CAMCHANGE
        'any transporters that are flying in need to be emptied
        emptyTransporters False
        endMissionCamChange
    ' left in so can skip the mission for testing...
    Case LEVEL_TYPE_LDS_EXPAND_LIMBO
        'shouldn't be any transporters on this mission but...who knows?
        endMissionExpandLimbo
    Case LEVEL_TYPE_LDS_MCLEAR
        'any transporters that are flying in need to be emptied
        emptyTransporters True
        endMissionOffClear
    Case LEVEL_TYPE_LDS_MKEEP_LIMBO
        'any transporters that are flying in need to be emptied
        emptyTransporters True
        endMissionOffKeepLimbo
    Case Else
        'error!
        debug_ 1480, LOG_FATAL, "endMission", "Unknown Mission Type"
        exit_ 3                             ' abort()
    End Select

    intRemoveMissionTimer
    intRemoveTransporterTimer

    'at end of mission always do this
    intRemoveTransporterLaunch

    'and this...
    'make sure the cheat time is not set for the next mission
    mission_cheatTime = 0


    'reset the bSetPlayCountDown flag
    setPlayCountDown 1                      ' true


    'mission.type = MISSION_NONE;
    mission_type = LEVEL_TYPE_LDS_NONE

    ' reset the transporters
    initTransporters
End Sub

Private Sub endMissionCamChange()
    'get any droids remaining from the previous campaign
    processPreviousCampDroids
End Sub

Private Sub endMissionOffClear()
    processMission
    restoreMissionData

    'reset variables in Droids
    missionResetDroids
End Sub

Private Sub endMissionOffKeep()
    processMission
    restoreMissionData

    'reset variables in Droids
    missionResetDroids
End Sub

'In this case any droids remaining (for selectedPlayer) go into a limbo list
'for use in a future mission (expand type)
Private Sub endMissionOffKeepLimbo()
    ' Save any droids left 'alive'
    processMissionLimbo

    ' Set the lists back to the home base lists
    restoreMissionData

    'reset variables in Droids
    missionResetDroids
End Sub

'This happens MID_MISSION now! but is left here in case the scripts fail but somehow get here...?
'The selectedPlayer's droids which were separated at the start of the
'mission need to merged back into the list
Private Sub endMissionExpandLimbo()
    restoreMissionLimboData
End Sub


'this is called mid Limbo mission via the script
Public Sub resetLimboMission()
    'add the units that were moved into the mission list at the start of the mission
    restoreMissionLimboData
    'set the mission type to plain old expand...
    mission_type = LEVEL_TYPE_LDS_EXPAND
End Sub

' The update routine for all droids left back at home base
' Only interested in Transporters at present
Public Sub missionDroidUpdate(ByVal psDroid As Long)
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/mission.cpp:1566", "missionDroidUpdate", "Invalid unit pointer": Exit Sub

    'This is required for Transporters that are moved offWorld so the
    'saveGame doesn't try to set their position in the map - especially important
    'for endCam2 where there isn't a valid map!
    If DROID_isTransporter(psDroid) Then
        aDroid(psDroid).pos.x = INVALID_XY
        aDroid(psDroid).pos.y = INVALID_XY
    End If

    'ignore all droids except Transporters
    If Not DROID_isTransporter(psDroid) Then Exit Sub
    If Not (orderState(psDroid, DORDER_TRANSPORTOUT) Or _
            orderState(psDroid, DORDER_TRANSPORTIN) Or _
            orderState(psDroid, DORDER_TRANSPORTRETURN)) Then
        Exit Sub
    End If

    ' NO ai update droid

    ' update the droids order
    If Not orderUpdateDroid(psDroid) Then
        ASSERT_FAILURE "false", "src/mission.cpp:1591", "missionDroidUpdate", "orderUpdateDroid returned false?"
    End If

    ' update the action of the droid
    actionUpdateDroid psDroid

    'NO move update
End Sub

' Reset variables in Droids such as order and position
Private Sub missionResetDroids()
    Dim player As Long, it As Long, nxt As Long, d As Long, psDroid As Long, placed As Boolean
    Dim psStruct As Long, hasFactory As Boolean, pickRes As PICKTILE, x As Long, y As Long, wx As Long, wy As Long
    Dim it2 As Long, psStructure As Long

    debug_ 1603, LOG_SAVE, "missionResetDroids", "called"

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:1605", "missionResetDroids", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer

    For player = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(player).head
        Do While it <> 0
            nxt = lnNext(it)
            d = lnObj(it)
            ' Reset order - unless constructor droid that is mid-build
            If (aDroid(d).droidType = DROID_CONSTRUCT _
                Or aDroid(d).droidType = DROID_CYBORG_CONSTRUCT) _
               And orderStateObj(d, DORDER_BUILD) <> 0 Then
                ' Need to set the action time to ignore the previous mission time
                aDroid(d).actionStarted = gameTime
            Else
                orderDroid d, DORDER_STOP, ModeImmediate
            End If

            'KILL OFF TRANSPORTER
            If DROID_isTransporter(d) Then
                vanishDroid d
            Else
                If aDroid(d).pos.x <> INVALID_XY And aDroid(d).pos.y <> INVALID_XY Then
                    ' update visibility
                    visTilesUpdate d
                End If
            End If
            it = nxt
        Loop
    Next

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psDroid = lnObj(it)
        placed = False

        'for all droids that have never left home base
        If aDroid(psDroid).pos.x = INVALID_XY And aDroid(psDroid).pos.y = INVALID_XY Then
            psStruct = aDroid(psDroid).psBaseStruct
            hasFactory = False                  ' FACTORY* psFactory = nullptr;

            If psStruct <> 0 Then
                If STRUCTURE_isFactory(psStruct) Then
                    hasFactory = (aStruct(psStruct).pFunctionality <> 0)    ' psFactory = (FACTORY*)psStruct->pFunctionality;
                End If
            End If
            'find a location next to the factory
            If hasFactory Then
                ' Use factory DP if one
                If aStruct(psStruct).functionality.factory.psAssemblyPoint <> 0 Then
                    x = map_coord(aFlagPos(aStruct(psStruct).functionality.factory.psAssemblyPoint).coords.x)
                    y = map_coord(aFlagPos(aStruct(psStruct).functionality.factory.psAssemblyPoint).coords.y)
                Else
                    x = map_coord(aStruct(psStruct).pos.x)
                    y = map_coord(aStruct(psStruct).pos.y)
                End If
                pickRes = pickHalfATile(x, y, LOOK_FOR_EMPTY_TILE)
                If pickRes = NO_FREE_TILE Then
                    ASSERT_FAILURE "false", "src/mission.cpp:1675", "missionResetDroids", "missionResetUnits: Unable to find a free location"
                Else
                    wx = world_coord(x)
                    wy = world_coord(y)

                    droidSetPosition psDroid, wx, wy
                    placed = True
                End If
            Else ' if couldn't find the factory - try to place near HQ instead
                it2 = apsStructLists(aDroid(psDroid).player).head
                Do While it2 <> 0
                    psStructure = lnObj(it2)
                    If asStructureStats(aStruct(psStructure).pStructureType).typ = REF_HQ Then
                        x = map_coord(aStruct(psStructure).pos.x)
                        y = map_coord(aStruct(psStructure).pos.y)
                        pickRes = pickHalfATile(x, y, LOOK_FOR_EMPTY_TILE)

                        If pickRes = NO_FREE_TILE Then
                            ASSERT_FAILURE "false", "src/mission.cpp:1698", "missionResetDroids", "missionResetUnits: Unable to find a free location"
                        Else
                            wx = world_coord(x)
                            wy = world_coord(y)

                            droidSetPosition psDroid, wx, wy
                            placed = True
                        End If
                        Exit Do
                    End If
                    it2 = lnNext(it2)
                Loop
            End If
            If placed Then
                ' Do all the things in build droid that never did when it was built!
                ' check the droid is a reasonable distance from the edge of the map
                If aDroid(psDroid).pos.x <= world_coord(EDGE_SIZE) Or _
                   aDroid(psDroid).pos.y <= world_coord(EDGE_SIZE) Or _
                   aDroid(psDroid).pos.x >= world_coord(mapWidth - EDGE_SIZE) Or _
                   aDroid(psDroid).pos.y >= world_coord(mapHeight - EDGE_SIZE) Then
                    debug_ 1721, LOG_ERROR, "missionResetDroids", "missionResetUnits: unit too close to edge of map - removing"
                    vanishDroid psDroid
                    GoTo continue2
                End If

                ' People always stand upright
                If aDroid(psDroid).droidType <> DROID_PERSON And Not DROID_isCyborg(psDroid) Then
                    updateDroidOrientation psDroid
                End If
                ' Reset the selected flag
                aDroid(psDroid).selected = 0

                ' update visibility
                visTilesUpdate psDroid
            Else
                'can't put it down so get rid of this droid!!
                ASSERT_FAILURE "false", "src/mission.cpp:1740", "missionResetDroids", "missionResetUnits: can't place unit - cancel to continue"
                vanishDroid psDroid
            End If
        End If
continue2:
        it = nxt
    Loop
End Sub

'unloads the Transporter passed into the mission at the specified x/y
Public Sub unloadTransporter(ByVal psTransporter As Long, ByVal x As Long, ByVal y As Long)
    Dim disembarkedDroids() As Long, disembarkedDroids_size As Long   ' static std::vector<DROID *> disembarkedDroids
    Dim droidX As Long, droidY As Long, has_lastDroid As Boolean, lastDroidX As Long, lastDroidY As Long
    Dim psGroup As Long, it As Long, nxt As Long, psDroid As Long, iX As Long, iY As Long, bFound As Boolean, k As Long

    If Not (psTransporter <> 0) Then ASSERT_FAILURE "psTransporter != nullptr", "src/mission.cpp:1757", "unloadTransporter", "Invalid transporter": Exit Sub
    ' (ppCurrentList = &apsDroidLists)

    disembarkedDroids_size = 0              ' disembarkedDroids.clear()

    'unload all the droids from within the current Transporter
    If DROID_isTransporter(psTransporter) Then
        If Not (aDroid(psTransporter).psGroup <> 0) Then ASSERT_FAILURE "psTransporter->psGroup != nullptr", "src/mission.cpp:1765", "unloadTransporter", "psTransporter->psGroup is null??"
        it = aGroup(aDroid(psTransporter).psGroup).psList.head
        Do While it <> 0
            psDroid = lnObj(it)
            If psDroid = psTransporter Then
                Exit Do
            End If

            'starting point...based around the value passed in
            droidX = map_coord(x)
            droidY = map_coord(y)
            If Not pickATileGen(droidX, droidY, LOOK_FOR_EMPTY_TILE, PAT_zonedPAT) Then
                If Not bMultiPlayer Then
                    ' try to avoid failing to unload a transporter in campaign mode - stack units on top of each other if needed
                    If has_lastDroid Then
                        debug_ 1783, LOG_INFO, "unloadTransporter", "unloadTransporter: Stacking units due to lack of space"
                        droidX = lastDroidX
                        droidY = lastDroidY
                    Else
                        ASSERT_FAILURE "false", "src/mission.cpp:1789", "unloadTransporter", "unloadTransporter: Unable to find a valid location"
                        Exit Do
                    End If
                Else
                    ' in skirmish / MP, a full or partial failure to unload may occur - log and continue with possibly partial disembark
                    debug_ 1796, LOG_INFO, "unloadTransporter", "unloadTransporter: Only room to unload %zu units", disembarkedDroids_size
                    Exit Do
                End If
            End If
            has_lastDroid = True
            lastDroidX = droidX
            lastDroidY = droidY

            'add it back into current droid lists
            addDroid psDroid, apsDroidLists

            droidSetPosition psDroid, world_coord(droidX), world_coord(droidY)
            updateDroidOrientation psDroid

            'reset droid orders
            orderDroid psDroid, DORDER_STOP, ModeImmediate
            aDroid(psDroid).selected = 0
            If Not bMultiPlayer Then
                ' So VTOLs don't try to rearm on another map
                setDroidBase psDroid, 0
            End If

            ReDim Preserve disembarkedDroids(0 To disembarkedDroids_size)  ' disembarkedDroids.push_back(psDroid)
            disembarkedDroids(disembarkedDroids_size) = psDroid
            disembarkedDroids_size = disembarkedDroids_size + 1
            it = lnNext(it)
        Loop

        If Not bMultiPlayer Or disembarkedDroids_size <> 0 Then
            ' trigger script callback detailing group about to disembark
            triggerEvent TRIGGER_TRANSPORTER_LANDED, psTransporter
        End If

        ' remove droids from transporter group if not already transferred to script group
        it = aGroup(aDroid(psTransporter).psGroup).psList.head
        Do While it <> 0
            nxt = lnNext(it)
            psDroid = lnObj(it)
            If psDroid = psTransporter Then
                Exit Do
            End If

            bFound = False                   ' std::find(disembarkedDroids.begin(), disembarkedDroids.end(), psDroid)
            For k = 0 To disembarkedDroids_size - 1
                If disembarkedDroids(k) = psDroid Then
                    bFound = True
                    Exit For
                End If
            Next
            If Not bFound Then
                ' not found in list of disembarked droids - skip removing from the transporter group
            ElseIf aDroid(psDroid).droidType = DROID_COMMAND Then
                ' a commander needs to get it's group back
                psGroup = grpCreate()
                DROID_GROUP_add psGroup, psDroid
                clearCommandDroidFactory psDroid
            Else
                DROID_GROUP_remove aDroid(psTransporter).psGroup, psDroid
            End If
            it = nxt
        Loop
    End If

    ' Don't do this in multiPlayer
    If Not bMultiPlayer Then
        ' Send all transporter Droids back to home base if off world
        ' Stop the camera following the transporter
        aDroid(psTransporter).selected = 0

        ' Send transporter offworld
        iX = 0
        iY = 0
        missionGetTransporterExit aDroid(psTransporter).player, iX, iY
        orderDroidLoc psTransporter, DORDER_TRANSPORTRETURN, iX, iY, ModeImmediate

        ' Set the launch time so the transporter doesn't just disappear for CAMSTART/CAMCHANGE
        transporterSetLaunchTime gameTime
    End If

    disembarkedDroids_size = 0              ' disembarkedDroids.clear()
End Sub

Public Sub missionMoveTransporterOffWorld(ByVal psTransporter As Long)
    Dim psForm As WIDGET, i As Long, it As Long, bFound As Boolean

    If aDroid(psTransporter).droidType = DROID_SUPERTRANSPORTER Then
        ' trigger script callback
        triggerEvent TRIGGER_TRANSPORTER_EXIT, psTransporter

        If droidRemove(psTransporter, apsDroidLists) Then
            addDroid psTransporter, mission_apsDroidLists
        End If

        'stop the droid moving - the moveUpdate happens AFTER the orderUpdate and can cause problems if the droid moves from one tile to another
        moveReallyStopDroid psTransporter

        ' clear targets / action targets
        setDroidTarget psTransporter, 0
        For i = 0 To MAX_WEAPONS - 1
            setDroidActionTarget psTransporter, 0, i
        Next

        'if offworld mission, then add the timer
        'if (mission.type == LDS_MKEEP || mission.type == LDS_MCLEAR)
        If missionCanReEnforce() And aDroid(psTransporter).player = selectedPlayer Then
            addTransporterTimerInterface
            'set the data for the transporter timer label
            widgSetUserData psWScreen, IDTRANTIMER_DISPLAY, psTransporter

            'make sure the button is enabled
            Set psForm = widgGetFromID(psWScreen, IDTRANTIMER_BUTTON)
            If Not psForm Is Nothing Then
                psForm.setState WBUT_PLAIN
            End If
        End If
        'need a callback for when all the selectedPlayers' reinforcements have been delivered
        If aDroid(psTransporter).player = selectedPlayer Then
            If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:1918", "missionMoveTransporterOffWorld", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer
            ' std::find_if(mission.apsDroidLists[selectedPlayer], !d->isTransporter())
            bFound = False
            it = mission_apsDroidLists(selectedPlayer).head
            Do While it <> 0
                If Not DROID_isTransporter(lnObj(it)) Then
                    bFound = True
                    Exit Do
                End If
                it = lnNext(it)
            Loop
            If Not bFound Then
                triggerEvent TRIGGER_TRANSPORTER_DONE, psTransporter
            End If
        End If
    Else
        debug_ 1931, LOG_SAVE, "missionMoveTransporterOffWorld", "droid type not transporter!"
    End If
End Sub


'add the Mission timer into the top  right hand corner of the screen
Private Function intAddMissionTimer() As Boolean
    Dim sFormInit As W_FORMINIT, sLabInit As W_LABINIT

    'check to see if it exists already
    If Not widgGetFromID(psWScreen, IDTIMER_FORM) Is Nothing Then
        intAddMissionTimer = True
        Exit Function
    End If

    ' Add the background
    W_FORMINIT_construct sFormInit

    sFormInit.formID = 0
    sFormInit.id = IDTIMER_FORM
    sFormInit.style = WFORM_PLAIN

    sFormInit.width = iV_GetImageWidth(IntImages, IMAGE_MISSION_CLOCK) 'TIMER_WIDTH;
    sFormInit.height = iV_GetImageHeight(IntImages, IMAGE_MISSION_CLOCK) 'TIMER_HEIGHT;
    sFormInit.x = ToInt16(RADTLX() + RADWIDTH - sFormInit.width)   ' (SWORD)
    sFormInit.y = TIMER_Y                   ' (SWORD)
    Set sFormInit.calcLayout = WzFunc_new(FN_intAddMissionTimer_calcLayout)
    sFormInit.UserData = PACKDWORD_TRI(0, IMAGE_MISSION_CLOCK, IMAGE_MISSION_CLOCK_UP)
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayMissionClock)

    If widgAddForm(psWScreen, sFormInit) Is Nothing Then
        intAddMissionTimer = False
        Exit Function
    End If

    'add labels for the time display
    W_LABINIT_construct sLabInit
    sLabInit.formID = IDTIMER_FORM
    sLabInit.id = IDTIMER_DISPLAY
    sLabInit.style = WLAB_PLAIN Or WIDG_HIDDEN
    sLabInit.x = TIMER_LABELX
    sLabInit.y = TIMER_LABELY
    sLabInit.width = sFormInit.width       'TIMER_WIDTH;
    sLabInit.height = sFormInit.height     'TIMER_HEIGHT;
    sLabInit.pText = "00:00:00"
    Set sLabInit.pCallback = WzFunc_new(FN_intUpdateMissionTimer)

    If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
        intAddMissionTimer = False
        Exit Function
    End If

    intAddMissionTimer = True
End Function

'add the Transporter timer into the top left hand corner of the screen
Private Function intAddTransporterTimer() As Boolean
    Dim sFormInit As W_FORMINIT, pForm As WIDGET, help As WidgetHelp, sLabInit As W_LABINIT, psCapacityLabel As WIDGET
    Dim emptyLabInit As W_LABINIT

    ' Make sure that Transporter Launch button isn't up as well
    intRemoveTransporterLaunch

    'check to see if it exists already
    If Not widgGetFromID(psWScreen, IDTRANTIMER_BUTTON) Is Nothing Then
        intAddTransporterTimer = True
        Exit Function
    End If

    ' Add the button form - clickable
    W_FORMINIT_construct sFormInit
    sFormInit.formID = 0
    sFormInit.id = IDTRANTIMER_BUTTON
    sFormInit.style = WFORM_CLICKABLE Or WFORM_NOCLICKMOVE
    sFormInit.x = TRAN_FORM_X
    sFormInit.y = TRAN_FORM_Y
    sFormInit.width = iV_GetImageWidth(IntImages, IMAGE_TRANSETA_UP)
    sFormInit.height = iV_GetImageHeight(IntImages, IMAGE_TRANSETA_UP)
    sFormInit.pTip = "Load Transport"
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    sFormInit.UserData = PACKDWORD_TRI(0, IMAGE_TRANSETA_DOWN, IMAGE_TRANSETA_UP)

    Set pForm = widgAddForm(psWScreen, sFormInit)
    If pForm Is Nothing Then
        intAddTransporterTimer = False
        Exit Function
    End If

    Set help = New WidgetHelp
    help.setTitle "Load Transport"
    help.setDescription "Shows the number of units currently loaded into the mission transporter, and the total capacity."
    help.addInteraction WidgetHelp_triggers(InteractionTriggers_PrimaryClick), "Open the Transporter Load Menu"
    pForm.setHelp help

    'add labels for the time display
    W_LABINIT_construct sLabInit
    sLabInit.formID = IDTRANTIMER_BUTTON
    sLabInit.id = IDTRANTIMER_DISPLAY
    sLabInit.style = WIDG_HIDDEN
    sLabInit.x = TRAN_TIMER_X
    sLabInit.y = TRAN_TIMER_Y
    sLabInit.width = TRAN_TIMER_WIDTH
    sLabInit.height = sFormInit.height
    Set sLabInit.pCallback = WzFunc_new(FN_intUpdateTransporterTimer)
    If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
        intAddTransporterTimer = False
        Exit Function
    End If

    'add the capacity label
    sLabInit = emptyLabInit                 ' sLabInit = W_LABINIT();
    W_LABINIT_construct sLabInit
    sLabInit.formID = IDTRANTIMER_BUTTON
    sLabInit.id = IDTRANS_CAPACITY
    sLabInit.x = 70
    sLabInit.y = 1
    sLabInit.width = 16
    sLabInit.height = 16
    sLabInit.pText = "00/10"
    Set sLabInit.pCallback = WzFunc_new(FN_intUpdateTransCapacity)
    Set psCapacityLabel = widgAddLabel(psWScreen, sLabInit)
    If psCapacityLabel Is Nothing Then
        intAddTransporterTimer = False
        Exit Function
    End If
    psCapacityLabel.setTransparentToMouse True

    intAddTransporterTimer = True
End Function

Public Sub missionSetReinforcementTime(ByVal iTime As Long)
    g_iReinforceTime = iTime
End Sub

Public Function missionGetReinforcementTime() As Long
    missionGetReinforcementTime = g_iReinforceTime
End Function

'fills in a hours(if bHours = true), minutes and seconds display for a given time in 1000th sec
Private Sub fillTimeDisplay(ByVal Label As WIDGET, ByVal time_ As Long, ByVal bHours As Boolean)
    Dim psText As String, tmp As tm
    'this is only for the transporter timer - never have hours!
    If time_ = LZ_COMPROMISED_TIME Then
        psText = "--:--"
    Else
        tmp = getUtcTime(Int(U32(time_) / GAME_TICKS_PER_SEC))  ' static_cast<time_t>(time / GAME_TICKS_PER_SEC)
        psText = strftime(IIf(bHours, "%H:%M:%S", "%M:%S"), tmp)
    End If
    Label.setString psText
End Sub


'update function for the mission timer
Private Sub intUpdateMissionTimer(ByVal psWidget As WIDGET)
    Dim Label As WIDGET, timeElapsed As Long, timeRemaining As Long
    Set Label = psWidget                    ' W_LABEL *Label = (W_LABEL *)psWidget;

    ' If the cheatTime has been set, then don't want the timer to countdown until stop cheating
    If mission_cheatTime <> 0 Then
        timeElapsed = mission_cheatTime - mission_startTime
    Else
        timeElapsed = gameTime - mission_startTime
    End If

    If Not challengeActive Then
        timeRemaining = mission_time - timeElapsed
        If timeRemaining < 0 Then
            timeRemaining = 0
        End If
    Else
        timeRemaining = timeElapsed
    End If

    fillTimeDisplay Label, timeRemaining, True
    Label.show                              ' Make sure its visible

    If challengeActive Then
        Exit Sub                            ' all done
    End If

    'make timer flash if time remaining < 5 minutes
    If timeRemaining < FIVE_MINUTES Then
        flashMissionButton IDTIMER_FORM
    End If
    'stop timer from flashing when gets to < 4 minutes
    If timeRemaining < FOUR_MINUTES Then
        stopMissionButtonFlash IDTIMER_FORM
    End If
    'play audio the first time the timed mission is started
    If timeRemaining <> 0 And (missionCountDown And NOT_PLAYED_ACTIVATED) <> 0 Then
        audio_QueueTrack ID_SOUND_MISSION_TIMER_ACTIVATED
        missionCountDown = missionCountDown And Not NOT_PLAYED_ACTIVATED
    End If
    'play some audio for mission countdown - start at 10 minutes remaining
    If getPlayCountDown() And timeRemaining < TEN_MINUTES Then
        If timeRemaining < TEN_MINUTES And (missionCountDown And NOT_PLAYED_TEN) <> 0 Then
            audio_QueueTrack ID_SOUND_10_MINUTES_REMAINING
            missionCountDown = missionCountDown And Not NOT_PLAYED_TEN
        ElseIf timeRemaining < FIVE_MINUTES And (missionCountDown And NOT_PLAYED_FIVE) <> 0 Then
            audio_QueueTrack ID_SOUND_5_MINUTES_REMAINING
            missionCountDown = missionCountDown And Not NOT_PLAYED_FIVE
        ElseIf timeRemaining < THREE_MINUTES And (missionCountDown And NOT_PLAYED_THREE) <> 0 Then
            audio_QueueTrack ID_SOUND_3_MINUTES_REMAINING
            missionCountDown = missionCountDown And Not NOT_PLAYED_THREE
        ElseIf timeRemaining < TWO_MINUTES And (missionCountDown And NOT_PLAYED_TWO) <> 0 Then
            audio_QueueTrack ID_SOUND_2_MINUTES_REMAINING
            missionCountDown = missionCountDown And Not NOT_PLAYED_TWO
        ElseIf timeRemaining < ONE_MINUTE And (missionCountDown And NOT_PLAYED_ONE) <> 0 Then
            audio_QueueTrack ID_SOUND_1_MINUTE_REMAINING
            missionCountDown = missionCountDown And Not NOT_PLAYED_ONE
        End If
    End If
End Sub

'update function for the transporter timer
Private Sub intUpdateTransporterTimer(ByVal psWidget As WIDGET)
    Dim Label As WIDGET, psTransporter As Long, timeRemaining As Long, ETA As Long
    Set Label = psWidget                    ' W_LABEL *Label = (W_LABEL *)psWidget;

    ETA = mission_ETA
    If ETA < 0 Then
        ETA = 0
    End If

    ' Get the object associated with this widget.
    psTransporter = widgetUserDroid(Label)  ' (DROID *)Label->pUserData
    If psTransporter <> 0 Then
        If Not (psTransporter <> 0) Then ASSERT_FAILURE "psTransporter != nullptr", "src/mission.cpp:2191", "intUpdateTransporterTimer", "intUpdateTransporterTimer: invalid Droid pointer"

        If aDroid(psTransporter).action = DACTION_TRANSPORTIN Or _
           aDroid(psTransporter).action = DACTION_TRANSPORTWAITTOFLYIN Then
            If mission_ETA = LZ_COMPROMISED_TIME Then
                timeRemaining = LZ_COMPROMISED_TIME
            Else
                timeRemaining = mission_ETA - (gameTime - g_iReinforceTime)
                If timeRemaining < 0 Then
                    timeRemaining = 0
                End If
                If timeRemaining < TRANSPORTER_REINFORCE_LEADIN Then
                    ' arrived: tell the transporter to move to the new onworld
                    ' location if not already doing so
                    If aDroid(psTransporter).action = DACTION_TRANSPORTWAITTOFLYIN Then
                        missionFlyTransportersIn selectedPlayer, False
                        ' executeFnAndProcessScriptQueuedRemovals([psTransporter]() { triggerEvent(TRIGGER_TRANSPORTER_ARRIVED, psTransporter); });
                        triggerEvent TRIGGER_TRANSPORTER_ARRIVED, psTransporter
                        wzapi_processScriptQueuedObjectRemovals
                    End If
                End If
            End If
            fillTimeDisplay Label, timeRemaining, False
        Else
            fillTimeDisplay Label, ETA, False
        End If
    Else
        If missionCanReEnforce() Then  ' ((mission.type == LDS_MKEEP) || (mission.type == LDS_MCLEAR)) & (mission.ETA >= 0) ) {
            fillTimeDisplay Label, ETA, False
        Else

            fillTimeDisplay Label, 0, False

        End If
    End If

    'minutes
    '/*calcTime = timeRemaining / (60*GAME_TICKS_PER_SEC);
    'Label->aText[0] = (UBYTE)('0'+ calcTime / 10);
    'Label->aText[1] = (UBYTE)('0'+ calcTime % 10);
    'timeElapsed -= calcTime * (60*GAME_TICKS_PER_SEC);
    '//separator
    'Label->aText[3] = (UBYTE)(':');
    '//seconds
    'calcTime = timeRemaining / GAME_TICKS_PER_SEC;
    'Label->aText[3] = (UBYTE)('0'+ calcTime / 10);
    'Label->aText[4] = (UBYTE)('0'+ calcTime % 10);*/

    Label.show
End Sub

' Remove the Mission Timer widgets from the screen
Public Sub intRemoveMissionTimer()
    ' Check it's up.
    If Not widgGetFromID(psWScreen, IDTIMER_FORM) Is Nothing Then
        'and remove it.
        widgDelete psWScreen, IDTIMER_FORM
    End If
End Sub

' Remove the Transporter Timer widgets from the screen
Public Sub intRemoveTransporterTimer()

    'remove main screen
    If Not widgGetFromID(psWScreen, IDTRANTIMER_BUTTON) Is Nothing Then
        widgDelete psWScreen, IDTRANTIMER_BUTTON
    End If
End Sub



' ////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////
' mission result functions for the interface.



Private Sub intDisplayMissionBackDrop(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim cache As ScoreDataToScreenCache
    ' Any widget using intDisplayMissionBackDrop must have its pUserData initialized to a (ScoreDataToScreenCache*)
    ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
    Set cache = psWidget.pUserData

    scoreDataToScreen psWidget, cache
End Sub

Private Sub missionResetInGameState()
    ' Add the background
    ' get rid of reticule etc..
    intResetScreen False

    'stop the game if in single player mode
    setMissionPauseState

    ' reset the input state
    resetInput

    forceHidePowerBar
    intRemoveReticule
    intRemoveMissionTimer
    intRemoveTransporterTimer
    intHideInGameOptionsButton
    intHideGroupSelectionMenu
End Sub

Private Sub intDestroyMissionResultWidgets()
    widgDelete psWScreen, IDMISSIONRES_TITLE
    widgDelete psWScreen, IDMISSIONRES_FORM
    widgDelete psWScreen, IDMISSIONRES_BACKFORM
End Sub

Private Function intAddMissionResult_(ByVal result As Boolean, ByVal bPlaySuccess As Boolean, ByVal showBackDrop As Boolean) As Boolean ' static bool _intAddMissionResult
    Dim sFormInit As W_FORMINIT, missionResBackForm As WIDGET, missionResTitle As WIDGET, missionResForm As WIDGET
    Dim sLabInit As W_LABINIT, sButInit As W_BUTINIT
    ' ensure the guide screen is closed
    closeGuideScreen

    missionResetInGameState
    scoreUpdateVar WD_MISSION_ENDED     'Store completion time for this mission

    W_FORMINIT_construct sFormInit

    ' add some funky beats
    cdAudio_PlayTrack SONG_FRONTEND

    If Not bMultiPlayer And result And showBackDrop Then
        If Not screen_GetBackDrop() Then
            pie_LoadBackDrop SCREEN_MISSIONEND
        End If
        screen_RestartBackDrop
    End If

    ' ensure these widgets are deleted before attempting to create
    intDestroyMissionResultWidgets

    sFormInit.formID = 0
    sFormInit.id = IDMISSIONRES_BACKFORM
    sFormInit.style = WFORM_PLAIN
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayMissionBackDrop)
    Set sFormInit.pUserData = New ScoreDataToScreenCache
    Set sFormInit.onDelete = WzFunc_new(FN_deleteScoreDataToScreenCache)
    Set missionResBackForm = widgAddForm(psWScreen, sFormInit)
    If missionResBackForm Is Nothing Then ASSERT_FAILURE "missionResBackForm != nullptr", "src/mission.cpp:2360", "_intAddMissionResult", "Failed to create IDMISSIONRES_BACKFORM": intAddMissionResult_ = False: Exit Function
    missionResBackForm.setCalcLayout WzFunc_new(FN_missionResBackForm_layout)

    ' TITLE
    Set missionResTitle = IntFormAnimated_new()
    missionResBackForm.attach missionResTitle
    missionResTitle.id = IDMISSIONRES_TITLE
    missionResTitle.setCalcLayout WzFunc_new(FN_missionResTitle_layout)

    ' add form
    Set missionResForm = IntFormAnimated_new()
    missionResBackForm.attach missionResForm
    missionResForm.id = IDMISSIONRES_FORM
    missionResForm.setCalcLayout WzFunc_new(FN_missionResForm_layout)

    ' description of success/fail
    W_LABINIT_construct sLabInit
    sLabInit.formID = IDMISSIONRES_TITLE
    sLabInit.id = IDMISSIONRES_TXT
    sLabInit.style = WLAB_ALIGNCENTRE
    sLabInit.x = 0
    sLabInit.y = 12
    sLabInit.width = MISSIONRES_TITLE_W
    sLabInit.height = 16
    If result Then

        'don't bother adding the text if haven't played the audio
        If bPlaySuccess Then
            If Cheated Then sLabInit.pText = "OBJECTIVE ACHIEVED by cheating!" Else sLabInit.pText = "OBJECTIVE ACHIEVED" '"Objective Achieved";
        End If

    Else
        If Cheated Then sLabInit.pText = "OBJECTIVE FAILED--and you cheated!" Else sLabInit.pText = "OBJECTIVE FAILED" '"Objective Failed;
    End If
    sLabInit.FontID = font_regular
    If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
        intAddMissionResult_ = False
        Exit Function
    End If
    ' options.
    W_BUTINIT_construct sButInit
    sButInit.formID = IDMISSIONRES_FORM
    sButInit.style = WBUT_TXTCENTRE
    sButInit.width = MISSION_TEXT_W
    sButInit.height = MISSION_TEXT_H
    Set sButInit.pDisplay = WzFunc_new(FN_displayTextOption)
    Set sButInit.initPUserDataFunc = WzFunc_new(FN_newDisplayTextOptionCache)
    Set sButInit.onDelete = WzFunc_new(FN_deleteDisplayTextOptionCache)
    'if won
    If result Or bMultiPlayer Then
        ' Finished the mission, so display "Continue Game"
        If Not testPlayerHasWon() Or bMultiPlayer Then
            sButInit.x = MISSION_1_X
            sButInit.y = MISSION_1_Y
            sButInit.id = IDMISSIONRES_CONTINUE
            sButInit.pText = "Continue Game" '"Continue Game";
            widgAddButton psWScreen, sButInit
        End If

        ' Won the game, so display "Quit to main menu"
        If bMultiPlayer Or (testPlayerHasWon() And Not bMultiPlayer) Then
            sButInit.x = MISSION_2_X
            sButInit.y = MISSION_2_Y
            sButInit.id = IDMISSIONRES_QUIT
            sButInit.pText = "Quit To Main Menu"
            widgAddButton psWScreen, sButInit
        End If

        ' FIXME, We got serious issues with savegames at the *END* of some missions, and while they
        ' will load, they don't have the correct state information or other settings.
        ' See transition from CAM2->CAM3 for a example.
        ' (the commented out save option and Autosave.gam)
    Else
        'load
        sButInit.id = IDMISSIONRES_LOAD
        sButInit.x = MISSION_1_X
        sButInit.y = MISSION_1_Y
        sButInit.pText = "Load Saved Game" '"Load Saved Game";
        widgAddButton psWScreen, sButInit
        'quit
        sButInit.id = IDMISSIONRES_QUIT
        sButInit.x = MISSION_2_X
        sButInit.y = MISSION_2_Y
        sButInit.pText = "Quit To Main Menu" '"Quit to Main Menu";
        widgAddButton psWScreen, sButInit
    End If

    intMode = INT_MISSIONRES
    MissionResUp = True

    ' play result audio
    If result = True And bPlaySuccess Then
        audio_QueueTrack ID_SOUND_OBJECTIVE_ACCOMPLISHED
    End If

    intAddMissionResult_ = True
End Function


Public Function intAddMissionResult(ByVal result As Boolean, ByVal bPlaySuccess As Boolean, ByVal showBackDrop As Boolean) As Boolean
    ActivityManager_completedMission result, Cheated ' (and collectEndGameStatsData(): the end of game statistics, cluster J)
    intAddMissionResult = intAddMissionResult_(result, bPlaySuccess, showBackDrop)
End Function

Public Sub intRemoveMissionResultNoAnim()
    intDestroyMissionResultWidgets

    cdAudio_Stop

    MissionResUp = False
    intMode = INT_NORMAL

    'reset the pauses
    resetMissionPauseState

    ' add back the reticule and power bar.
    intAddReticule

    intShowPowerBar
    intShowGroupSelectionMenu
End Sub

Public Sub intRunMissionResult()
    Dim msg As String
    wzSetCursor CURSOR_DEFAULT

    If bLoadSaveUp Then
        If runLoadSave(False) Then      ' check for file name.
            If Len(sRequestResult) <> 0 Then
                debug_ 2529, LOG_SAVE, "intRunMissionResult", "Returned %s", sRequestResult

                If Not bRequestLoad Then
                    msg = ""            ' char msg[256] = {'\0'};

                    saveGame sRequestResult, GTYPE_SAVE_START
                    msg = "GAME SAVED :"
                    msg = Left$(msg & savegameWithoutExtension(sRequestResult), 255) ' sstrcat(msg, ...): 256 with the '\0'
                    addConsoleMessage msg, LEFT_JUSTIFY, NOTIFY_MESSAGE
                End If
            End If
        End If
    End If
End Sub

Private Sub missionContineButtonPressed()
    If nextMissionType = LEVEL_TYPE_LDS_CAMSTART _
       Or nextMissionType = LEVEL_TYPE_LDS_BETWEEN _
       Or nextMissionType = LEVEL_TYPE_LDS_EXPAND _
       Or nextMissionType = LEVEL_TYPE_LDS_EXPAND_LIMBO Then
        launchMission
    End If
    widgDelete psWScreen, IDMISSIONRES_FORM ' close option box.

    If bMultiPlayer Then
        intRemoveMissionResultNoAnim
    End If
End Sub

Public Sub intProcessMissionResult(ByVal id As Long)
    Dim sButInit As W_BUTINIT
    Select Case id
    Case IDMISSIONRES_LOAD
        ' throw up some filerequester
        addLoadSave LOAD_MISSIONEND, "Load Saved Game"
    Case IDMISSIONRES_SAVE
        addLoadSave SAVE_MISSIONEND, "Save Game"
        If widgGetFromID(psWScreen, IDMISSIONRES_QUIT) Is Nothing Then
            'Add Quit Button now save has been pressed
            W_BUTINIT_construct sButInit
            sButInit.formID = IDMISSIONRES_FORM
            sButInit.style = WBUT_TXTCENTRE
            sButInit.width = MISSION_TEXT_W
            sButInit.height = MISSION_TEXT_H
            Set sButInit.pDisplay = WzFunc_new(FN_displayTextOption)
            Set sButInit.pUserData = New DisplayTextOptionCache
            Set sButInit.onDelete = WzFunc_new(FN_deleteDisplayTextOptionCache)
            sButInit.id = IDMISSIONRES_QUIT
            sButInit.x = MISSION_3_X
            sButInit.y = MISSION_3_Y
            sButInit.pText = "Quit To Main Menu"
            widgAddButton psWScreen, sButInit
        End If
    Case IDMISSIONRES_QUIT
        ' catered for by hci.c.
    Case IDMISSIONRES_CONTINUE
        If bLoadSaveUp Then
            closeLoadSave               ' close save interface if it's up.
        End If
        missionContineButtonPressed

    Case Else
    End Select
End Sub

' end of interface stuff.
' ////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////

'builds a droid back at the home base whilst on a mission - stored in a list made
'available to the transporter interface
Public Function buildMissionDroid(ByVal psTempl As Long, ByVal x As Long, ByVal y As Long, ByVal player As Long) As Long
    Dim psNewDroid As Long, noOrders As INITIAL_DROID_ORDERS, rot As Rotation

    psNewDroid = buildDroid(psTempl, world_coord(x), world_coord(y), player, True, False, noOrders, rot)
    If psNewDroid = 0 Then
        buildMissionDroid = 0
        Exit Function
    End If
    addDroid psNewDroid, mission_apsDroidLists
    'set its x/y to impossible values so can detect when return from mission
    aDroid(psNewDroid).pos.x = INVALID_XY
    aDroid(psNewDroid).pos.y = INVALID_XY

    'set all the droids to selected from when return
    aDroid(psNewDroid).selected = -CLng(isSelectable(psNewDroid))

    ' Set died parameter correctly
    aDroid(psNewDroid).died = NOT_CURRENT_LIST

    buildMissionDroid = psNewDroid
End Function

'this causes the new mission data to be loaded up - only if startMission has been called
Public Sub launchMission()
    'if (mission.type == MISSION_NONE)
    If mission_type = LEVEL_TYPE_LDS_NONE Then
        ' tell the loop that a new level has to be loaded up
        loopMissionState = LMS_NEWLEVEL
    Else
        debug_ 2650, LOG_SAVE, "launchMission", "Start Mission has not been called"
    End If
End Sub


'sets up the game to start a new mission
Public Function setUpMission(ByVal typ As Long) As Boolean
    Dim bPlaySuccess As Boolean

    ' Close the interface
    intResetScreen True

    ' The last mission must have been successful otherwise endgame would have been called
    endMission

    'release the level data for the previous mission
    If Not levReleaseMissionData() Then
        setUpMission = False
        Exit Function
    End If

    If typ = LEVEL_TYPE_LDS_CAMSTART Then
        ' Another one of those lovely hacks!!
        bPlaySuccess = True

        ' We don't want the 'mission accomplished' audio/text message at end of cam1
        If getCampaignNumber() = 2 Then
            bPlaySuccess = False
        End If
        ' Give the option of save/continue
        If Not intAddMissionResult(True, bPlaySuccess, True) Then
            setUpMission = False
            Exit Function
        End If
        clearCampaignName
        loopMissionState = LMS_SAVECONTINUE
    ElseIf typ = LEVEL_TYPE_LDS_MKEEP _
           Or typ = LEVEL_TYPE_LDS_MCLEAR _
           Or typ = LEVEL_TYPE_LDS_MKEEP_LIMBO Then
        launchMission
    Else
        If Not getWidgetsStatus() Then
            setWidgetsStatus True
            intResetScreen False
        End If

        ' Give the option of save / continue
        If Not intAddMissionResult(True, True, True) Then
            setUpMission = False
            Exit Function
        End If
        loopMissionState = LMS_SAVECONTINUE
    End If

    setUpMission = True
End Function

'save the power settings before loading in the new map data
Private Sub saveMissionPower()
    Dim inc As Long

    For inc = 0 To MAX_PLAYERS - 1
        mission_asCurrentPower(inc) = getPower(inc)
    Next
End Sub

'add the power from the home base to the current power levels for the mission map
Public Sub adjustMissionPower()
    Dim inc As Long

    For inc = 0 To MAX_PLAYERS - 1
        addPower inc, mission_asCurrentPower(inc)
    Next
End Sub

'sets the appropriate pause states for when the interface is up but the
'game needs to be paused
Public Sub setMissionPauseState()
    If Not bMultiPlayer Then
        gameTimeStop
        setGameUpdatePause True
        setAudioPause True
        setScriptPause True
        setConsolePause True
    End If
End Sub

'resets the pause states
Public Sub resetMissionPauseState()
    If Not bMultiPlayer Then
        setGameUpdatePause False
        setAudioPause False
        setScriptPause False
        setConsolePause False
        gameTimeStart
    End If
End Sub

'gets the coords for a no go area
' (VB) LANDING_ZONE *: the zone's copy (game.cpp only reads it)
Public Function getLandingZone(ByVal i As Long) As LANDING_ZONE
    If Not ((i >= 0) And (i < MAX_NOGO_AREAS)) Then ASSERT_FAILURE "(i >= 0) && (i < (11 + 1))", "src/mission.cpp:2765", "getLandingZone", "getLandingZone out of range."
    getLandingZone = sLandingZone(i)
End Function

'Initialises all the nogo areas to 0
Public Sub initNoGoAreas()
    Dim i As Long
    For i = 0 To MAX_NOGO_AREAS - 1
        sLandingZone(i).x1 = 0
        sLandingZone(i).y1 = 0
        sLandingZone(i).x2 = 0
        sLandingZone(i).y2 = 0
    Next
End Sub

'sets the coords for a no go area
Public Sub setNoGoArea(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal area As Long)
    Dim t As Long
    ' make sure that x2 > x1 and y2 > y1
    If x2 < x1 Then
        t = x1: x1 = x2: x2 = t         ' std::swap(x1, x2)
    End If
    If y2 < y1 Then
        t = y1: y1 = y2: y2 = t         ' std::swap(y1, y2)
    End If

    sLandingZone(area).x1 = x1
    sLandingZone(area).x2 = x2
    sLandingZone(area).y1 = y1
    sLandingZone(area).y2 = y2

    If area = 0 And x1 <> 0 And y1 <> 0 Then
        addLandingLights getLandingX(area) + 64, getLandingY(area) + 64
    End If
End Sub

Private Sub addLandingLight(ByVal x As Long, ByVal y As Long, ByVal spec As LAND_LIGHT_SPEC, ByVal lit As Boolean)
    ' The height the landing lights should be above the ground
    Const AboveGround As Long = 16
    Dim pos As Vector3i

    If x < 0 Or y < 0 Then
        Exit Sub
    End If

    pos.x = x
    pos.z = y
    pos.y = map_Height(x, y) + AboveGround

    effectSetLandLightSpec spec

    addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_LAND_LIGHT, False, 0, -CLng(lit), graphicsTime   ' (the overload without effectTime: graphicsTime)
End Sub

Private Sub addLandingLights(ByVal x As Long, ByVal y As Long)
    addLandingLight x, y, LL_MIDDLE, True                       ' middle

    addLandingLight x + 128, y + 128, LL_OUTER, True            ' outer
    addLandingLight x + 128, y - 128, LL_OUTER, True
    addLandingLight x - 128, y + 128, LL_OUTER, True
    addLandingLight x - 128, y - 128, LL_OUTER, True

    addLandingLight x + 64, y + 64, LL_INNER, True              ' inner
    addLandingLight x + 64, y - 64, LL_INNER, True
    addLandingLight x - 64, y + 64, LL_INNER, True
    addLandingLight x - 64, y - 64, LL_INNER, True
End Sub

'   checks the x,y passed in are not within the boundary of any Landing Zone
'   x and y in tile coords
Public Function withinLandingZone(ByVal x As Long, ByVal y As Long) As Boolean
    Dim inc As Long

    If Not (ULT(x, mapWidth)) Then ASSERT_FAILURE "x < mapWidth", "src/mission.cpp:2843", "withinLandingZone", "withinLandingZone: x coord bigger than mapWidth"
    If Not (ULT(y, mapHeight)) Then ASSERT_FAILURE "y < mapHeight", "src/mission.cpp:2844", "withinLandingZone", "withinLandingZone: y coord bigger than mapHeight"

    For inc = 0 To MAX_NOGO_AREAS - 1
        If Not ULT(x, sLandingZone(inc).x1) And Not ULT(sLandingZone(inc).x2, x) And _
           Not ULT(y, sLandingZone(inc).y1) And Not ULT(sLandingZone(inc).y2, y) Then
            withinLandingZone = True
            Exit Function
        End If
    Next
    withinLandingZone = False
End Function

'returns the x coord for where the Transporter can land (for player 0)
Public Function getLandingX(ByVal iPlayer As Long) As Long
    If Not (iPlayer < MAX_NOGO_AREAS) Then ASSERT_FAILURE "iPlayer < MAX_NOGO_AREAS", "src/mission.cpp:2861", "getLandingX", "getLandingX: player %d out of range", iPlayer: getLandingX = 0: Exit Function
    getLandingX = world_coord((sLandingZone(iPlayer).x1 + (sLandingZone(iPlayer).x2 - _
                              sLandingZone(iPlayer).x1) \ 2)) And &HFFFF&
End Function

'returns the y coord for where the Transporter can land
Public Function getLandingY(ByVal iPlayer As Long) As Long
    If Not (iPlayer < MAX_NOGO_AREAS) Then ASSERT_FAILURE "iPlayer < MAX_NOGO_AREAS", "src/mission.cpp:2869", "getLandingY", "getLandingY: player %d out of range", iPlayer: getLandingY = 0: Exit Function
    getLandingY = world_coord((sLandingZone(iPlayer).y1 + (sLandingZone(iPlayer).y2 - _
                              sLandingZone(iPlayer).y1) \ 2)) And &HFFFF&
End Function

'returns the x coord for where the Transporter can land back at home base
Private Function getHomeLandingX() As Long
    getHomeLandingX = map_coord(mission_homeLZ_X)
End Function

'returns the y coord for where the Transporter can land back at home base
Private Function getHomeLandingY() As Long
    getHomeLandingY = map_coord(mission_homeLZ_Y)
End Function

Public Sub missionSetTransporterEntry(ByVal iPlayer As Long, ByVal iEntryTileX As Long, ByVal iEntryTileY As Long)
    If Not (iPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "iPlayer < MAX_PLAYERS", "src/mission.cpp:2888", "missionSetTransporterEntry", "missionSetTransporterEntry: player %i too high", iPlayer: Exit Sub

    If (iEntryTileX > scrollMinX) And (iEntryTileX < scrollMaxX) Then
        mission_iTranspEntryTileX(iPlayer) = iEntryTileX And &HFFFF&
    Else
        debug_ 2896, LOG_SAVE, "missionSetTransporterEntry", "entry point x %i outside scroll limits %i->%i", iEntryTileX, scrollMinX, scrollMaxX
        mission_iTranspEntryTileX(iPlayer) = (scrollMinX + EDGE_SIZE) And &HFFFF&
    End If

    If (iEntryTileY > scrollMinY) And (iEntryTileY < scrollMaxY) Then
        mission_iTranspEntryTileY(iPlayer) = iEntryTileY And &HFFFF&
    Else
        debug_ 2906, LOG_SAVE, "missionSetTransporterEntry", "entry point y %i outside scroll limits %i->%i", iEntryTileY, scrollMinY, scrollMaxY
        mission_iTranspEntryTileY(iPlayer) = (scrollMinY + EDGE_SIZE) And &HFFFF&
    End If
End Sub

Public Sub missionSetTransporterExit(ByVal iPlayer As Long, ByVal iExitTileX As Long, ByVal iExitTileY As Long)
    If Not (iPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "iPlayer < MAX_PLAYERS", "src/mission.cpp:2913", "missionSetTransporterExit", "missionSetTransporterExit: player %i too high", iPlayer: Exit Sub

    If (iExitTileX > scrollMinX) And (iExitTileX < scrollMaxX) Then
        mission_iTranspExitTileX(iPlayer) = iExitTileX And &HFFFF&
    Else
        debug_ 2921, LOG_SAVE, "missionSetTransporterExit", "entry point x %i outside scroll limits %i->%i", iExitTileX, scrollMinX, scrollMaxX
        mission_iTranspExitTileX(iPlayer) = (scrollMinX + EDGE_SIZE) And &HFFFF&
    End If

    If (iExitTileY > scrollMinY) And (iExitTileY < scrollMaxY) Then
        mission_iTranspExitTileY(iPlayer) = iExitTileY And &HFFFF&
    Else
        debug_ 2931, LOG_SAVE, "missionSetTransporterExit", "entry point y %i outside scroll limits %i->%i", iExitTileY, scrollMinY, scrollMaxY
        mission_iTranspExitTileY(iPlayer) = (scrollMinY + EDGE_SIZE) And &HFFFF&
    End If
End Sub

' iX, iY: UWORD *
Public Sub missionGetTransporterEntry(ByVal iPlayer As Long, iX As Long, iY As Long)
    If Not (iPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "iPlayer < MAX_PLAYERS", "src/mission.cpp:2938", "missionGetTransporterEntry", "missionGetTransporterEntry: player %i too high", iPlayer: Exit Sub

    iX = world_coord(mission_iTranspEntryTileX(iPlayer)) And &HFFFF&
    iY = world_coord(mission_iTranspEntryTileY(iPlayer)) And &HFFFF&
End Sub

' iX, iY: UDWORD *
Public Sub missionGetTransporterExit(ByVal iPlayer As Long, iX As Long, iY As Long)
    If Not (iPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "iPlayer < MAX_PLAYERS", "src/mission.cpp:2946", "missionGetTransporterExit", "missionGetTransporterExit: player %i too high", iPlayer: Exit Sub

    iX = world_coord(mission_iTranspExitTileX(iPlayer))
    iY = world_coord(mission_iTranspExitTileY(iPlayer))
End Sub

'update routine for mission details
Public Sub missionTimerUpdate()
    'don't bother with the time check if have 'cheated'
    If mission_cheatTime = 0 Then
        'Want a mission timer on all types of missions now - AB 26/01/99
        'only interested in off world missions (so far!) and if timer has been set
        If mission_time >= 0 Then  '&& (
            'mission.type == LDS_MKEEP || mission.type == LDS_MKEEP_LIMBO ||
            'mission.type == LDS_MCLEAR || mission.type == LDS_BETWEEN))
            'check if time is up
            If (gameTime - mission_startTime) > mission_time Then
                'the script can call the end game cos have failed!
                ' executeFnAndProcessScriptQueuedRemovals([]() { triggerEvent(TRIGGER_MISSION_TIMEOUT); });
                triggerEvent TRIGGER_MISSION_TIMEOUT
                wzapi_processScriptQueuedObjectRemovals
            End If
        End If
    End If
End Sub


' Remove any objects left ie walls,structures and droids that are not the selected player.
'
Public Sub missionDestroyObjects()
    Dim Player As Long, i As Long, it As Long, nxt As Long, psDroid As Long, psStruct As Long

    debug_ 2981, LOG_SAVE, "missionDestroyObjects", "called"
    proj_FreeAllProjectiles
    For Player = 0 To MAX_PLAYERS - 1
        If Player <> selectedPlayer Then
            ' AI player, clear out old data

            it = apsDroidLists(Player).head
            Do While it <> 0
                nxt = lnNext(it)
                removeDroidBase lnObj(it)
                it = nxt
            Loop

            'clear out the mission lists as well to make sure no Transporters exist
            moveObjList apsDroidLists(Player), mission_apsDroidLists(Player)

            it = apsDroidLists(Player).head
            Do While it <> 0
                nxt = lnNext(it)
                psDroid = lnObj(it)
                'make sure its died flag is not set since we've swapped the apsDroidList pointers over
                aDroid(psDroid).died = 0
                removeDroidBase psDroid
                it = nxt
            Loop
            objlist_clear mission_apsDroidLists(Player)

            it = apsStructLists(Player).head
            Do While it <> 0
                nxt = lnNext(it)
                removeStruct lnObj(it), True
                it = nxt
            Loop
        End If
    Next

    ' human player, check that we do not reference the cleared out data
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:3016", "missionDestroyObjects", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer
    Player = selectedPlayer

    it = apsDroidLists(Player).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).psBaseStruct <> 0 Then
            If aStruct(aDroid(psDroid).psBaseStruct).died <> 0 Then
                setDroidBase psDroid, 0
            End If
        End If
        For i = 0 To MAX_WEAPONS - 1
            If aDroid(psDroid).psActionTarget(i) <> 0 Then
                If SIMPLE_OBJECT_died(aDroid(psDroid).psActionTarget(i)) <> 0 Then
                    setDroidActionTarget psDroid, 0, i
                    ' Clear action too if this requires a valid first action target
                    If i = 0 _
                       And aDroid(psDroid).action <> DACTION_MOVEFIRE _
                       And aDroid(psDroid).action <> DACTION_TRANSPORTIN _
                       And aDroid(psDroid).action <> DACTION_TRANSPORTOUT Then
                        aDroid(psDroid).action = DACTION_NONE
                    End If
                End If
            End If
        Next
        If aDroid(psDroid).order.psObj <> 0 Then
            If SIMPLE_OBJECT_died(aDroid(psDroid).order.psObj) <> 0 Then
                setDroidTarget psDroid, 0
            End If
        End If
        it = lnNext(it)
    Loop

    it = apsStructLists(Player).head
    Do While it <> 0
        psStruct = lnObj(it)
        For i = 0 To MAX_WEAPONS - 1
            If aStruct(psStruct).psTarget(i) <> 0 Then
                If SIMPLE_OBJECT_died(aStruct(psStruct).psTarget(i)) <> 0 Then
                    setStructureTarget psStruct, 0, i, ORIGIN_UNKNOWN
                End If
            End If
        Next
        it = lnNext(it)
    Loop

    ' FIXME: check that orders do not reference anything bad?

    If psDestroyedObj.size <> 0 Then
        debug_ 3061, LOG_INFO, "missionDestroyObjects", "%zu destroyed objects", psDestroyedObj.size
    End If
    gameTime = gameTime + 1 ' Wonderful hack to ensure objects destroyed above get free'ed up by objmemUpdate.
    objmemUpdate ' Actually free objects removed above
End Sub

Private Sub processPreviousCampDroids()
    Dim it As Long, nxt As Long, psDroid As Long

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/mission.cpp:3069", "processPreviousCampDroids", "selectedPlayer %u exceeds MAX_PLAYERS", selectedPlayer

    ' See if any are left
    If mission_apsDroidLists(selectedPlayer).size <> 0 Then
        it = mission_apsDroidLists(selectedPlayer).head
        Do While it <> 0
            nxt = lnNext(it)
            psDroid = lnObj(it)
            ' We want to kill off all droids now! - AB 27/01/99
            ' KILL OFF TRANSPORTER
            If droidRemove(psDroid, mission_apsDroidLists) Then
                addDroid psDroid, apsDroidLists
                vanishDroid psDroid
            End If
            it = nxt
        Loop
    End If
End Sub

'access functions for droidsToSafety flag - so we don't have to end the mission when a Transporter fly's off world
Public Sub setDroidsToSafetyFlag(ByVal set_ As Boolean)
    bDroidsToSafety = set_
End Sub

Public Function getDroidsToSafetyFlag() As Boolean
    getDroidsToSafetyFlag = bDroidsToSafety
End Function


'access functions for bPlayCountDown flag - true = play coded mission count down
Public Sub setPlayCountDown(ByVal set_ As Long)
    bPlayCountDown = set_ And &HFF&
End Sub

Public Function getPlayCountDown() As Boolean
    getPlayCountDown = (bPlayCountDown <> 0)
End Function


'checks to see if the player has any droids (except Transporters left)
Public Function missionDroidsRemaining(ByVal player As Long) As Boolean
    Dim it As Long
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/mission.cpp:3115", "missionDroidsRemaining", "invalid player: %u", player: missionDroidsRemaining = False: Exit Function
    it = apsDroidLists(player).head
    Do While it <> 0
        If Not DROID_isTransporter(lnObj(it)) Then
            missionDroidsRemaining = True
            Exit Function
        End If
        it = lnNext(it)
    Loop
    missionDroidsRemaining = False
End Function

'called when a Transporter gets to the edge of the world and the droids are
'being flown to safety. The droids inside the Transporter are placed into the
'mission list for later use
Public Sub moveDroidsToSafety(ByVal psTransporter As Long)
    Dim it As Long, nxt As Long, psDroid As Long

    If Not (DROID_isTransporter(psTransporter)) Then ASSERT_FAILURE "psTransporter->isTransporter()", "src/mission.cpp:3131", "moveDroidsToSafety", "unit not a Transporter": Exit Sub

    If aDroid(psTransporter).psGroup <> 0 Then
        'move droids out of Transporter into mission list
        it = aGroup(aDroid(psTransporter).psGroup).psList.head
        Do While it <> 0
            nxt = lnNext(it)
            psDroid = lnObj(it)
            If psDroid = psTransporter Then
                Exit Do
            End If
            DROID_GROUP_remove aDroid(psTransporter).psGroup, psDroid
            'cam change add droid
            aDroid(psDroid).pos.x = INVALID_XY
            aDroid(psDroid).pos.y = INVALID_XY
            addDroid psDroid, mission_apsDroidLists
            it = nxt
        Loop
    End If

    'move the transporter into the mission list also
    If droidRemove(psTransporter, apsDroidLists) Then
        addDroid psTransporter, mission_apsDroidLists
    End If
End Sub

Public Sub clearMissionWidgets()
    'remove any widgets that are up due to the missions
    If mission_time > 0 Then
        intRemoveMissionTimer
    End If

    If mission_ETA >= 0 Then
        intRemoveTransporterTimer
    End If

    intRemoveTransporterLaunch
End Sub

'
' Try to find a transporter among the player's droids, or in the mission list (transporter waiting to come back).
'
Private Function find_transporter() As Long
    Dim it As Long

    If selectedPlayer >= MAX_PLAYERS Then
        find_transporter = 0
        Exit Function
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        If DROID_isTransporter(lnObj(it)) Then
            find_transporter = lnObj(it)
            Exit Function
        End If
        it = lnNext(it)
    Loop
    it = mission_apsDroidLists(selectedPlayer).head
    Do While it <> 0
        If DROID_isTransporter(lnObj(it)) Then
            find_transporter = lnObj(it)
            Exit Function
        End If
        it = lnNext(it)
    Loop

    find_transporter = 0
End Function

Public Sub resetMissionWidgets()
    Dim transporter As Long, showLaunch As Boolean

    If mission_type = LEVEL_TYPE_LDS_NONE Then
        Exit Sub
    End If

    'add back any widgets that should be up due to the missions
    If mission_time > 0 Then
        intAddMissionTimer
        'make sure its not flashing when added
        stopMissionButtonFlash IDTIMER_FORM
    End If

    transporter = find_transporter()

    ' Check if not a reinforcement type mission and we have an transporter
    ' (&&: transporterFlying only for a transporter)
    If Not missionForReInforcements() Then
        If transporter <> 0 Then
            showLaunch = Not transporterFlying(transporter)
        End If
    End If
    If showLaunch Then
        ' Show launch button if the transporter has not already been launched
        intAddTransporterLaunch transporter
    ElseIf mission_ETA >= 0 Then
        addTransporterTimerInterface
    End If

End Sub

'deals with any selectedPlayer's transporters that are flying in when the
'mission ends. bOffWorld is true if the Mission is currently offWorld
Private Sub emptyTransporters(ByVal bOffWorld As Boolean)
    Dim it As Long, nxt As Long, it2 As Long, nxt2 As Long, psTransporter As Long, psDroid As Long

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/mission.cpp:3237", "emptyTransporters", "selectedPlayer %u >= MAX_PLAYERS", selectedPlayer: Exit Sub

    'see if there are any Transporters in the world
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psTransporter = lnObj(it)
        If DROID_isTransporter(psTransporter) Then
            'if flying in, empty the contents
            If orderState(psTransporter, DORDER_TRANSPORTIN) Then
                ' if we're offWorld, all we need to do is put the droids into the apsDroidList
                ' and processMission() will assign them a location etc
                If bOffWorld Then
                    it2 = aGroup(aDroid(psTransporter).psGroup).psList.head
                    Do While it2 <> 0
                        nxt2 = lnNext(it2)
                        psDroid = lnObj(it2)
                        If psDroid = psTransporter Then
                            Exit Do
                        End If
                        'take it out of the Transporter group
                        DROID_GROUP_remove aDroid(psTransporter).psGroup, psDroid
                        'add it back into current droid lists
                        addDroid psDroid, apsDroidLists
                        it2 = nxt2
                    Loop
                ' we're not offWorld so add to mission.apsDroidList to be
                ' processed by the endMission function
                Else
                    it2 = aGroup(aDroid(psTransporter).psGroup).psList.head
                    Do While it2 <> 0
                        nxt2 = lnNext(it2)
                        psDroid = lnObj(it2)
                        If psDroid = psTransporter Then
                            Exit Do
                        End If
                        'take it out of the Transporter group
                        DROID_GROUP_remove aDroid(psTransporter).psGroup, psDroid
                        'add it back into current droid lists
                        addDroid psDroid, mission_apsDroidLists
                        it2 = nxt2
                    Loop
                End If
                'now kill off the Transporter
                vanishDroid psTransporter
            ElseIf Not bOffWorld And orderState(psTransporter, DORDER_TRANSPORTRETURN) Then
                'also destroy transporters in the process of flying back and we're not offWorld
                vanishDroid psTransporter
            End If
        End If
        it = nxt
    Loop

    'deal with any transporters that are waiting to come over
    it = mission_apsDroidLists(selectedPlayer).head
    Do While it <> 0
        nxt = lnNext(it)
        psTransporter = lnObj(it)
        If DROID_isTransporter(psTransporter) Then
            'for each droid within the transporter...
            it2 = aGroup(aDroid(psTransporter).psGroup).psList.head
            Do While it2 <> 0
                nxt2 = lnNext(it2)
                psDroid = lnObj(it2)
                If psDroid = psTransporter Then
                    Exit Do
                End If
                'take it out of the Transporter group
                DROID_GROUP_remove aDroid(psTransporter).psGroup, psDroid
                'add it back into mission droid lists
                addDroid psDroid, mission_apsDroidLists
                it2 = nxt2
            Loop
        End If
        'don't need to destroy the transporter here - it is dealt with by the endMission process
        it = nxt
    Loop
End Sub

'bCheating = true == start of cheat
'bCheating = false == end of cheat
Public Sub setMissionCheatTime(ByVal bCheating As Boolean)
    If bCheating Then
        mission_cheatTime = gameTime
    Else
        'adjust the mission start time for the duration of the cheat!
        mission_startTime = FromU32(U32(mission_startTime) + U32(gameTime) - U32(mission_cheatTime))
        mission_cheatTime = 0
    End If
End Sub

'--------------------------------------------------------------------------
' (VB) the lambdas and callbacks of the mission result screen and the timers

Public Function mission_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET
    Select Case f.id
    Case FN_intDisplayMissionBackDrop
        intDisplayMissionBackDrop args(0), args(1), args(2)

    Case FN_deleteScoreDataToScreenCache
        ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
        Set psWidget = args(0)
        psWidget.pUserData = Empty      ' delete static_cast<ScoreDataToScreenCache *>(psWidget->pUserData); psWidget->pUserData = nullptr;

    Case FN_missionResBackForm_layout
        Set psWidget = args(0)
        psWidget.setGeometry 0 + D_W, 0 + D_H, 640, 480

    Case FN_missionResTitle_layout
        Set psWidget = args(0)
        psWidget.setGeometry MISSIONRES_TITLE_X, MISSIONRES_TITLE_Y, MISSIONRES_TITLE_W, MISSIONRES_TITLE_H

    Case FN_missionResForm_layout
        Set psWidget = args(0)
        psWidget.setGeometry MISSIONRES_X, MISSIONRES_Y, MISSIONRES_W, MISSIONRES_H

    Case FN_intAddMissionTimer_calcLayout
        ' LAMBDA_CALCLAYOUT_SIMPLE: psWidget->move((SWORD)(RADTLX + RADWIDTH - psWidget->width() - 18), TIMER_Y);
        Set psWidget = args(0)
        psWidget.move ToInt16(RADTLX() + RADWIDTH - psWidget.width() - 18), TIMER_Y

    Case FN_intUpdateMissionTimer
        ' (psWidget, the W_CONTEXT's xOffset, yOffset, mx, my)
        intUpdateMissionTimer args(0)

    Case FN_intUpdateTransporterTimer
        intUpdateTransporterTimer args(0)
    End Select
End Function
