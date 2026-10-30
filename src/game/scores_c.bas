Attribute VB_Name = "scores_c"
' src/scores.cpp, scores.h: the present mission's statistics (MISSION_DATA missionData), the end of mission
' screen's statistics bars (scoreDataToScreen, with the ScoreDataToScreenCache class), the end game
' statistics, the save game's score file and the autogame's summary. getAsciiTime (scores.h's template) is
' getAsciiTimeN with the callers' char[255]: both return the text. cheat.cpp's Cheated is cheat_c's.
Option Explicit

Public Enum DATA_INDEX
    WD_UNITS_BUILT
    WD_UNITS_KILLED
    WD_UNITS_LOST
    WD_STR_BUILT
    WD_STR_KILLED
    WD_STR_LOST
    WD_ARTEFACTS_FOUND
    WD_MISSION_STARTED
    WD_MISSION_ENDED
    WD_SHOTS_ON_TARGET
    WD_SHOTS_OFF_TARGET
    WD_BARBARIANS_MOWED_DOWN
End Enum

' --------------------------------------------------------------------
' The mission results data
Public Type MISSION_DATA
    unitsBuilt As Long                  ' uint32_t: How many units were built
    unitsKilled As Long                 ' How many enemy units you blew up
    unitsLost As Long                   ' How many units were lost
    strBuilt As Long                    ' How many structures we built
    strKilled As Long                   ' How many enemy structures you blew up
    strLost As Long                     ' How many structures were lost
    artefactsFound As Long              ' How many artefacts were found
    missionStarted As Long              ' When was the mission started
    missionEnded As Long                ' When was the mission ended
    shotsOnTarget As Long               ' How many hits
    shotsOffTarget As Long              ' How many misses
    babasMowedDown As Long              ' How many barbarians did we mow down?
End Type

' Could use widgets, but hey.....
Private Type STAT_BAR
    topX As Long                        ' UDWORD topX, topY: Obvious
    topY As Long
    width As Long                       ' UDWORD width, height: Height down screen and width _unfilled_
    height As Long
    percent As Long                     ' What percentage full is it?
    stringID As Long                    ' String resource name to stick next to it.
    queTime As Long                     ' How many game ticks before it's active?
    bQueued As Boolean                  ' Already fired off?
    bActive As Boolean                  ' Is this one active?
    number As Long                      ' %d string for the associated text string.
End Type

Public Type END_GAME_STATS_DATA
    missionData As MISSION_DATA
    maxDroidsPerLevel As Long           ' uint32_t = 0
    numDroidsPerLevel() As Long         ' std::vector<uint32_t>: numDroidsPerLevel(0 .. numDroidsPerLevel_size - 1)
    numDroidsPerLevel_size As Long
    numUnits As Long                    ' uint32_t = 0
End Type

' enum (the infoBars' indices)
Private Const STAT_UNIT_LOST As Long = 0
Private Const STAT_UNIT_KILLED As Long = 1
Private Const STAT_STR_LOST As Long = 2
Private Const STAT_STR_BLOWN_UP As Long = 3
Private Const STAT_UNITS_BUILT As Long = 4
Private Const STAT_UNITS_NOW As Long = 5
Private Const STAT_STR_BUILT As Long = 6
Private Const STAT_STR_NOW As Long = 7
Private Const STAT_ROOKIE As Long = 8
Private Const STAT_GREEN As Long = 9
Private Const STAT_TRAINED As Long = 10
Private Const STAT_REGULAR As Long = 11
Private Const STAT_VETERAN As Long = 12
Private Const STAT_CRACK As Long = 13
Private Const STAT_ELITE As Long = 14
Private Const STAT_SPECIAL As Long = 15
Private Const STAT_ACE As Long = 16

Private Const BAR_CRAWL_TIME As Long = GAME_TICKS_PER_SEC * 3

Private Const DROID_LEVELS As Long = 9

Private Const LC_X As Long = 32
Private Const RC_X As Long = 320 + 32
Private Const RANK_BAR_WIDTH As Long = 100
Private Const STAT_BAR_WIDTH As Long = 100

' enum MR_STRING
Private Const STR_MR_UNITS_LOST As Long = 0
Private Const STR_MR_UNITS_KILLED As Long = 1
Private Const STR_MR_STR_LOST As Long = 2
Private Const STR_MR_STR_BLOWN_UP As Long = 3
Private Const STR_MR_UNITS_BUILT As Long = 4
Private Const STR_MR_UNITS_NOW As Long = 5
Private Const STR_MR_STR_BUILT As Long = 6
Private Const STR_MR_STR_NOW As Long = 7

Private Const STR_MR_LEVEL_ROOKIE As Long = 8
Private Const STR_MR_LEVEL_GREEN As Long = 9
Private Const STR_MR_LEVEL_TRAINED As Long = 10
Private Const STR_MR_LEVEL_REGULAR As Long = 11
Private Const STR_MR_LEVEL_VETERAN As Long = 12
Private Const STR_MR_LEVEL_CRACK As Long = 13
Private Const STR_MR_LEVEL_ELITE As Long = 14
Private Const STR_MR_LEVEL_SPECIAL As Long = 15
Private Const STR_MR_LEVEL_ACE As Long = 16

Private infoBars(0 To 17) As STAT_BAR   ' STAT_BAR infoBars[] (scores_c_init)

' The present mission data
Public missionData As MISSION_DATA
Private dispST As Long                  ' static UDWORD dispST;
Private bDispStarted As Boolean         ' static bool bDispStarted = false;
' (static char text[255]; static char text2[255]: locals of scoreDataToScreen in VB)

' stdOutGameSummary's static UDWORD lastOutputRealTime = 0;
Private lastOutputRealTime As Long

' #define MT_Y_POS (MISSIONRES_TITLE_Y  + D_H + 80)
Private Function MT_Y_POS() As Long
    MT_Y_POS = MISSIONRES_TITLE_Y + D_H() + 80
End Function

' return translated string
Private Function getDescription(ByVal id As Long) As String
    Select Case id
    Case STR_MR_UNITS_LOST: getDescription = "Own Units: %u"
    Case STR_MR_UNITS_KILLED: getDescription = "Enemy Units: %u"
    Case STR_MR_STR_LOST: getDescription = "Own Structures: %u"
    Case STR_MR_STR_BLOWN_UP: getDescription = "Enemy Structures: %u"
    Case STR_MR_UNITS_BUILT: getDescription = "Units Manufactured: %u"
    Case STR_MR_UNITS_NOW: getDescription = "Total Units: %u"
    Case STR_MR_STR_BUILT: getDescription = "Structures Built: %u"
    Case STR_MR_STR_NOW: getDescription = "Total Structures: %u"

    Case STR_MR_LEVEL_ROOKIE: getDescription = "Rookie: %u"
    Case STR_MR_LEVEL_GREEN: getDescription = "Green: %u"     ' P_("rank", "Green: %u")
    Case STR_MR_LEVEL_TRAINED: getDescription = "Trained: %u"
    Case STR_MR_LEVEL_REGULAR: getDescription = "Regular: %u"
    Case STR_MR_LEVEL_VETERAN: getDescription = "Professional: %u"
    Case STR_MR_LEVEL_CRACK: getDescription = "Veteran: %u"
    Case STR_MR_LEVEL_ELITE: getDescription = "Elite: %u"
    Case STR_MR_LEVEL_SPECIAL: getDescription = "Special: %u"
    Case STR_MR_LEVEL_ACE: getDescription = "Hero: %u"
    Case Else
        ' make compiler shut up
        getDescription = ""
    End Select
End Function

Private Function getColour(ByVal id As Long) As Long
    Select Case id
    Case STR_MR_STR_LOST, STR_MR_UNITS_LOST
        getColour = psPalette(WZCOL_MENU_SCORE_LOSS)
    Case STR_MR_STR_BLOWN_UP, STR_MR_UNITS_KILLED
        getColour = psPalette(WZCOL_MENU_SCORE_DESTROYED)
    Case STR_MR_UNITS_BUILT, STR_MR_UNITS_NOW, STR_MR_STR_BUILT, STR_MR_STR_NOW
        getColour = psPalette(WZCOL_MENU_SCORE_BUILT)
    Case STR_MR_LEVEL_ROOKIE, STR_MR_LEVEL_GREEN, STR_MR_LEVEL_TRAINED, STR_MR_LEVEL_REGULAR, STR_MR_LEVEL_VETERAN, _
         STR_MR_LEVEL_CRACK, STR_MR_LEVEL_ELITE, STR_MR_LEVEL_SPECIAL, STR_MR_LEVEL_ACE
        getColour = psPalette(WZCOL_MENU_SCORE_RANK)
    Case Else
        getColour = psPalette(WZCOL_BLACK)
    End Select
End Function

' (VB) STAT_BAR infoBars[]'s initialiser
Private Sub infoBar(ByVal i As Long, ByVal topX As Long, ByVal topY As Long, ByVal width As Long, ByVal height As Long, ByVal percent As Long, _
                    ByVal stringID As Long, ByVal queTime As Long, ByVal bQueued As Boolean, ByVal bActive As Boolean, ByVal number As Long)
    infoBars(i).topX = topX
    infoBars(i).topY = topY
    infoBars(i).width = width
    infoBars(i).height = height
    infoBars(i).percent = percent
    infoBars(i).stringID = stringID
    infoBars(i).queTime = queTime
    infoBars(i).bQueued = bQueued
    infoBars(i).bActive = bActive
    infoBars(i).number = number
End Sub

' (VB) the statics' initialisers (exemain_vb calls it)
Public Sub scores_c_init()
    infoBar 0, LC_X, 100, STAT_BAR_WIDTH, 16, 10, STR_MR_UNITS_LOST, 0, False, True, 0      ' left column		STAT_UNIT_LOST
    infoBar 1, LC_X, 120, STAT_BAR_WIDTH, 16, 20, STR_MR_UNITS_KILLED, 0, False, True, 0    ' STAT_UNIT_KILLED
    infoBar 2, LC_X, 160, STAT_BAR_WIDTH, 16, 30, STR_MR_STR_LOST, 0, False, True, 0        ' STAT_STR_LOST
    infoBar 3, LC_X, 180, STAT_BAR_WIDTH, 16, 40, STR_MR_STR_BLOWN_UP, 0, False, True, 0    ' STAT_STR_BLOWN_UP
    infoBar 4, LC_X, 220, STAT_BAR_WIDTH, 16, 50, STR_MR_UNITS_BUILT, 0, False, True, 0     ' STAT_UNITS_BUILT
    infoBar 5, LC_X, 240, STAT_BAR_WIDTH, 16, 60, STR_MR_UNITS_NOW, 0, False, True, 0       ' STAT_UNITS_NOW
    infoBar 6, LC_X, 260, STAT_BAR_WIDTH, 16, 70, STR_MR_STR_BUILT, 0, False, True, 0       ' STAT_STR_BUILT
    infoBar 7, LC_X, 280, STAT_BAR_WIDTH, 16, 80, STR_MR_STR_NOW, 0, False, False, 0        ' STAT_STR_NOW

    infoBar 8, RC_X, 100, RANK_BAR_WIDTH, 16, 10, STR_MR_LEVEL_ROOKIE, 0, False, True, 0    ' right column	//	STAT_ROOKIE
    infoBar 9, RC_X, 120, RANK_BAR_WIDTH, 16, 20, STR_MR_LEVEL_GREEN, 0, False, True, 0     ' STAT_GREEN
    infoBar 10, RC_X, 140, RANK_BAR_WIDTH, 16, 30, STR_MR_LEVEL_TRAINED, 0, False, True, 0  ' STAT_TRAINED
    infoBar 11, RC_X, 160, RANK_BAR_WIDTH, 16, 40, STR_MR_LEVEL_REGULAR, 0, False, True, 0  ' STAT_REGULAR
    infoBar 12, RC_X, 180, RANK_BAR_WIDTH, 16, 50, STR_MR_LEVEL_VETERAN, 0, False, True, 0  ' STAT_VETERAN
    infoBar 13, RC_X, 200, RANK_BAR_WIDTH, 16, 60, STR_MR_LEVEL_CRACK, 0, False, True, 0    ' STAT_CRACK
    infoBar 14, RC_X, 220, RANK_BAR_WIDTH, 16, 70, STR_MR_LEVEL_ELITE, 0, False, True, 0    ' STAT_ELITE
    infoBar 15, RC_X, 240, RANK_BAR_WIDTH, 16, 80, STR_MR_LEVEL_SPECIAL, 0, False, True, 0  ' STAT_SPECIAL
    infoBar 16, RC_X, 260, RANK_BAR_WIDTH, 16, 90, STR_MR_LEVEL_ACE, 0, False, True, 0      ' STAT_ACE

    infoBar 17, 0, 0, 0, 0, 0, 0, 0, False, False, 0
End Sub

' --------------------------------------------------------------------
' Initialise the mission data info - done before each mission
Public Function scoreInitSystem() As Boolean
    missionData.unitsBuilt = 0
    missionData.unitsKilled = 0
    missionData.unitsLost = 0

    missionData.strBuilt = 0
    missionData.strKilled = 0
    missionData.strLost = 0

    missionData.artefactsFound = 0
    scoreUpdateVar WD_MISSION_STARTED   ' total game time is just gameTime
    missionData.missionEnded = 0
    missionData.shotsOnTarget = 0
    missionData.shotsOffTarget = 0
    missionData.babasMowedDown = 0
    Cheated = False
    bDispStarted = False
    scoreInitSystem = True
End Function

' --------------------------------------------------------------------
' Updates a game statistic - more can be added if we need 'em
Public Sub scoreUpdateVar(ByVal var As DATA_INDEX)
    Select Case var
    Case WD_UNITS_BUILT
        missionData.unitsBuilt = missionData.unitsBuilt + 1         ' We've built another unit
    Case WD_UNITS_KILLED
        missionData.unitsKilled = missionData.unitsKilled + 1       ' We've destroyed an enemy unit
    Case WD_UNITS_LOST
        missionData.unitsLost = missionData.unitsLost + 1           ' We've lost a unit
    Case WD_STR_BUILT
        missionData.strBuilt = missionData.strBuilt + 1             ' Built a structure
    Case WD_STR_KILLED
        missionData.strKilled = missionData.strKilled + 1           ' Destroyed an enemy structure
    Case WD_STR_LOST
        missionData.strLost = missionData.strLost + 1               ' Lost a structure
    Case WD_ARTEFACTS_FOUND
        missionData.artefactsFound = missionData.artefactsFound + 1 ' Got an artefact
    Case WD_MISSION_STARTED
        missionData.missionStarted = gameTime   ' Init the mission start time
                                                ' Should be called once per mission
    Case WD_MISSION_ENDED
        missionData.missionEnded = gameTime - missionData.missionStarted    ' Mission ended
    Case WD_SHOTS_ON_TARGET
        missionData.shotsOnTarget = missionData.shotsOnTarget + 1   ' We hit something
    Case WD_SHOTS_OFF_TARGET
        missionData.shotsOffTarget = missionData.shotsOffTarget + 1 ' Missed something
    Case WD_BARBARIANS_MOWED_DOWN
        missionData.babasMowedDown = missionData.babasMowedDown + 1 ' Ran over a barbarian
    Case Else
        debug_ 259, LOG_FATAL, "scoreUpdateVar", "Weirdy variable request from scoreUpdateVar"
        exit_ 3                         ' abort()
    End Select
End Sub

' Builds an ascii string for the passed in components 4:02:23 for example.
' void getAsciiTimeN(char *pDest, size_t destSize, unsigned time, bool showMs): (VB) the text, returned
Public Function getAsciiTimeN(ByVal destSize As Long, ByVal time_ As Long, Optional ByVal showMs As Boolean = False) As String
    Dim hours As Long, minutes As Long, seconds As Long, milliseconds As Long, hourColon As String, hoursText As String, s As String
    getTimeComponents time_, hours, minutes, seconds, milliseconds

    If hours <> 0 Then hourColon = ":" Else hourColon = ""
    ' "%.0d": no digits for a 0
    If hours <> 0 Then hoursText = CStr(hours) Else hoursText = ""

    If showMs Then
        s = hoursText & hourColon & va("%02d:%02d.%03d", minutes, seconds, milliseconds)
    Else
        s = hoursText & hourColon & va("%02d:%02d", minutes, seconds)
    End If
    getAsciiTimeN = Left$(s, destSize - 1)  ' snprintf(pDest, destSize, ...)
End Function

' template <unsigned N> getAsciiTime(char (&dest)[N], unsigned time, bool showMs = false): (VB) the text, the
' callers' buffers being char[255]
Public Function getAsciiTime(ByVal time_ As Long, Optional ByVal showMs As Boolean = False) As String
    getAsciiTime = getAsciiTimeN(255, time_, showMs)
End Function

Public Function collectEndGameStatsData() As END_GAME_STATS_DATA
    Dim fullStats As END_GAME_STATS_DATA, i As Long, idx As Long, it As Long, it2 As Long, psDroid As Long, psCurr As Long
    Dim dList As Long, haveList As Boolean
    fullStats.missionData = missionData

    ReDim fullStats.numDroidsPerLevel(0 To DROID_LEVELS - 1)
    For i = 0 To DROID_LEVELS - 1
        fullStats.numDroidsPerLevel(i) = getNumDroidsForLevel(selectedPlayer, i)
    Next
    fullStats.numDroidsPerLevel_size = DROID_LEVELS

    fullStats.numUnits = 0
    If ULT(selectedPlayer, MAX_PLAYERS) Then
        idx = 0
        Do
            ' (VB) DroidList* dList: 0 apsDroidLists, 1 mission.apsDroidLists, 2 apsLimboDroids (nullptr: not taken)
            haveList = False
            Select Case idx
            Case 0: it = apsDroidLists(selectedPlayer).head: haveList = True
            Case 1: it = mission_apsDroidLists(selectedPlayer).head: haveList = True
            Case 2
                If prevMissionType = LEVEL_TYPE_LDS_MKEEP_LIMBO Then
                    it = apsLimboDroids(selectedPlayer).head: haveList = True
                End If
            End Select
            If Not haveList Then
                GoTo continue_
            End If
            Do While it <> 0
                psDroid = lnObj(it)
                it = lnNext(it)
                If psDroid = 0 Then GoTo continueDroid
                If isDead(psDroid) Then GoTo continueDroid
                fullStats.numUnits = fullStats.numUnits + 1
                If DROID_isTransporter(psDroid) Then
                    If aDroid(psDroid).psGroup = 0 Then
                        GoTo continueDroid
                    End If
                    it2 = aGroup(aDroid(psDroid).psGroup).psList.head
                    Do While it2 <> 0
                        psCurr = lnObj(it2)
                        it2 = lnNext(it2)
                        If psCurr <> psDroid Then
                            fullStats.numUnits = fullStats.numUnits + 1
                        End If
                    Loop
                End If
continueDroid:
            Loop
continue_:
            idx = idx + 1
        Loop While idx < 3
    End If

    collectEndGameStatsData = fullStats
End Function

' void scoreDataToScreen(WIDGET *psWidget, ScoreDataToScreenCache& cache)
Public Sub scoreDataToScreen(ByVal psWidget As WIDGET, ByVal cache As ScoreDataToScreenCache)
    Dim index As Long, x As Long, y As Long, width As Long, height As Long, bMoreBars As Boolean
    Dim mul As Single, length As Long, text As String, text2 As String, cheatedTextColor As Long, t As WzText

    If Not bDispStarted Then
        bDispStarted = True
        dispST = realTime
        audio_PlayTrack ID_SOUND_BUTTON_CLICK_5
    End If

    fillUpStats collectEndGameStatsData()

    pie_UniTransBoxFill 16 + D_W(), MT_Y_POS() - 16, pie_GetVideoBufferWidth() - D_W() - 16, MT_Y_POS() + 256 + 16, psPalette(WZCOL_SCORE_BOX)
    iV_Box 16 + D_W(), MT_Y_POS() - 16, pie_GetVideoBufferWidth() - D_W() - 16, MT_Y_POS() + 256 + 16, psPalette(WZCOL_SCORE_BOX_BORDER)

    cache.wzLabelText_UnitLosses.setText "Unit Losses", font_regular
    cache.wzLabelText_UnitLosses.render LC_X + D_W(), 80 + 16 + D_H(), psPalette(WZCOL_FORM_TEXT)
    cache.wzLabelText_StructureLosses.setText "Structure Losses", font_regular
    cache.wzLabelText_StructureLosses.render LC_X + D_W(), 140 + 16 + D_H(), psPalette(WZCOL_FORM_TEXT)
    cache.wzLabelText_ForceInformation.setText "Force Information", font_regular
    cache.wzLabelText_ForceInformation.render LC_X + D_W(), 200 + 16 + D_H(), psPalette(WZCOL_FORM_TEXT)

    index = 0
    bMoreBars = True
    Do While bMoreBars
        ' Is it time to display this bar?
        If infoBars(index).bActive Then
            ' Has it been queued before?
            If infoBars(index).bQueued = False Then
                ' Don't do this next time...!
                infoBars(index).bQueued = True

                ' Play a sound
                audio_PlayTrack ID_SOUND_BUTTON_CLICK_5
            End If
            x = infoBars(index).topX + D_W()
            y = infoBars(index).topY + D_H()
            width = infoBars(index).width
            height = infoBars(index).height

            iV_Box x, y, x + width, y + height, psPalette(WZCOL_BLACK)

            ' Draw the background border box
            pie_BoxFill x - 1, y - 1, x + width + 1, y + height + 1, psPalette(WZCOL_MENU_BACKGROUND)

            ' Draw the interior grey
            pie_BoxFill x, y, x + width, y + height, psPalette(WZCOL_MENU_SCORES_INTERIOR)

            If ((U32(realTime - dispST) > U32(infoBars(index).queTime))) Then
                ' Now draw amount filled
                If U32(realTime - dispST) < BAR_CRAWL_TIME Then
                    mul = CSng(U32(realTime - dispST)) / CSng(BAR_CRAWL_TIME)
                Else
                    mul = 1!
                End If

                length = Fix(CSng(CSng(CSng(U32(infoBars(index).percent)) / 100!) * CSng(U32(infoBars(index).width))) * mul)

                If length > 4 Then

                    ' Black shadow
                    pie_BoxFill x + 1, y + 3, x + length - 1, y + height - 1, psPalette(WZCOL_MENU_SHADOW)

                    ' Solid coloured bit
                    pie_BoxFill x + 1, y + 2, x + length - 4, y + height - 4, getColour(index)
                End If
            End If
            ' Now render the text by the bar
            text = Left$(va(getDescription(infoBars(index).stringID), infoBars(index).number), 254)
            Do While index >= cache.wzInfoBarText.Count     ' cache.wzInfoBarText.resize(index + 1);
                cache.wzInfoBarText.Add New WzText
            Loop
            Set t = cache.wzInfoBarText(index + 1)
            t.setText text, font_regular
            t.render x + width + 16, y + 12, psPalette(WZCOL_FORM_TEXT)

            ' If we're beyond STAT_ROOKIE, then we're on rankings
            If index >= STAT_GREEN And index <= STAT_ACE Then
                iV_DrawImage IntImages, (IMAGE_LEV_0 + (index - STAT_GREEN)) And &HFFFF&, x - 8, y + 2
            End If
        End If
        ' Move onto the next bar
        index = index + 1
        If index > STAT_ACE Then
            bMoreBars = False
        ElseIf infoBars(index).topX = 0 And infoBars(index).topY = 0 Then
            bMoreBars = False
        End If
    Loop

    ' We now need to display the mission time, game time, average unit experience level an number of artefacts found

    ' Firstly, top of the screen, number of artefacts found
    text = Left$(va("ARTIFACTS RECOVERED: %d", missionData.artefactsFound), 254)
    cache.wzInfoText_ArtifactsFound.setText text, font_regular
    cache.wzInfoText_ArtifactsFound.render (pie_GetVideoBufferWidth() - cache.wzInfoText_ArtifactsFound.width()) \ 2, 300 + D_H(), psPalette(WZCOL_FORM_TEXT)

    ' Get the mission result time in a string - and write it out
    text2 = getAsciiTime(missionData.missionEnded)
    text = Left$(va("Mission Time - %s", text2), 254)
    cache.wzInfoText_MissionTime.setText text, font_regular
    cache.wzInfoText_MissionTime.render (pie_GetVideoBufferWidth() - cache.wzInfoText_MissionTime.width()) \ 2, 320 + D_H(), psPalette(WZCOL_FORM_TEXT)

    ' Write out total game time so far
    text2 = getAsciiTime(gameTime)
    text = Left$(va("Total Game Time - %s", text2), 254)
    cache.wzInfoText_TotalGameTime.setText text, font_regular
    cache.wzInfoText_TotalGameTime.render (pie_GetVideoBufferWidth() - cache.wzInfoText_TotalGameTime.width()) \ 2, 340 + D_H(), psPalette(WZCOL_FORM_TEXT)
    If Cheated Then
        ' A quick way to flash the text
        If ((U32(realTime) \ 250) Mod 2) <> 0 Then cheatedTextColor = psPalette(WZCOL_RED) Else cheatedTextColor = psPalette(WZCOL_YELLOW)
        text = Left$("You cheated!", 254)
        cache.wzInfoText_Cheated.setText text, font_regular
        cache.wzInfoText_Cheated.render (pie_GetVideoBufferWidth() - cache.wzInfoText_Cheated.width()) \ 2, 360 + D_H(), cheatedTextColor
    End If
End Sub

' (VB) PERCENT(a, b): (a) * 100 / (b), in UDWORD arithmetic
Private Function PERCENT_(ByVal a As Long, ByVal b As Long) As Long
    PERCENT_ = FromU32(Fix((U32(a) * 100) / U32(b)))
End Function

' -----------------------------------------------------------------------------------
Private Sub fillUpStats(stats As END_GAME_STATS_DATA)
    Dim i As Long, maxi As Long, scaleFactor As Single, length As Long

    ' Do rankings first cos they're easier
    maxi = 0
    For i = 0 To DROID_LEVELS - 1
        If U32(stats.numDroidsPerLevel(i)) > U32(maxi) Then maxi = stats.numDroidsPerLevel(i)
    Next

    ' Make sure we got something
    If maxi = 0 Then
        scaleFactor = 0!
    Else
        scaleFactor = CSng(RANK_BAR_WIDTH) / CSng(U32(maxi))
    End If

    ' Scale for percent
    For i = 0 To DROID_LEVELS - 1
        length = FromU32(Fix(CSng(scaleFactor * CSng(U32(stats.numDroidsPerLevel(i))))))
        infoBars(STAT_ROOKIE + i).percent = PERCENT_(length, RANK_BAR_WIDTH)
        infoBars(STAT_ROOKIE + i).number = stats.numDroidsPerLevel(i)
    Next

    ' Now do the other stuff...
    ' Units killed and lost...
    maxi = MAX_U32(stats.missionData.unitsLost, stats.missionData.unitsKilled)
    If maxi = 0 Then
        scaleFactor = 0!
    Else
        scaleFactor = CSng(STAT_BAR_WIDTH) / CSng(U32(maxi))
    End If

    length = FromU32(Fix(CSng(scaleFactor * CSng(U32(stats.missionData.unitsLost)))))
    infoBars(STAT_UNIT_LOST).percent = PERCENT_(length, STAT_BAR_WIDTH)
    length = FromU32(Fix(CSng(scaleFactor * CSng(U32(stats.missionData.unitsKilled)))))
    infoBars(STAT_UNIT_KILLED).percent = PERCENT_(length, STAT_BAR_WIDTH)

    ' Now do the structure losses
    maxi = MAX_U32(stats.missionData.strLost, stats.missionData.strKilled)
    If maxi = 0 Then
        scaleFactor = 0!
    Else
        scaleFactor = CSng(STAT_BAR_WIDTH) / CSng(U32(maxi))
    End If

    length = FromU32(Fix(CSng(scaleFactor * CSng(U32(stats.missionData.strLost)))))
    infoBars(STAT_STR_LOST).percent = PERCENT_(length, STAT_BAR_WIDTH)
    length = FromU32(Fix(CSng(scaleFactor * CSng(U32(stats.missionData.strKilled)))))
    infoBars(STAT_STR_BLOWN_UP).percent = PERCENT_(length, STAT_BAR_WIDTH)

    ' Finally the force information - need amount of droids as well

    maxi = MAX_U32(stats.missionData.unitsBuilt, stats.missionData.strBuilt)
    maxi = MAX_U32(maxi, stats.numUnits)

    If maxi = 0 Then
        scaleFactor = 0!
    Else
        scaleFactor = CSng(STAT_BAR_WIDTH) / CSng(U32(maxi))
    End If

    length = FromU32(Fix(CSng(scaleFactor * CSng(U32(stats.missionData.unitsBuilt)))))
    infoBars(STAT_UNITS_BUILT).percent = PERCENT_(length, STAT_BAR_WIDTH)
    length = FromU32(Fix(CSng(scaleFactor * CSng(U32(stats.numUnits)))))
    infoBars(STAT_UNITS_NOW).percent = PERCENT_(length, STAT_BAR_WIDTH)
    length = FromU32(Fix(CSng(scaleFactor * CSng(U32(stats.missionData.strBuilt)))))
    infoBars(STAT_STR_BUILT).percent = PERCENT_(length, STAT_BAR_WIDTH)

    ' Finally the numbers themselves
    infoBars(STAT_UNIT_LOST).number = stats.missionData.unitsLost
    infoBars(STAT_UNIT_KILLED).number = stats.missionData.unitsKilled
    infoBars(STAT_STR_LOST).number = stats.missionData.strLost
    infoBars(STAT_STR_BLOWN_UP).number = stats.missionData.strKilled
    infoBars(STAT_UNITS_BUILT).number = stats.missionData.unitsBuilt
    infoBars(STAT_UNITS_NOW).number = stats.numUnits
    infoBars(STAT_STR_BUILT).number = stats.missionData.strBuilt
End Sub

' (VB) MAX(a, b) of two uint32_t
Private Function MAX_U32(ByVal a As Long, ByVal b As Long) As Long
    If U32(a) > U32(b) Then MAX_U32 = a Else MAX_U32 = b
End Function

' -----------------------------------------------------------------------------------
' This will save out the score data
Public Function writeScoreData(ByVal fileName As String) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadAndWrite)

    ' Dump the scores for the current player
    ini.setValue "unitsBuilt", json_value(U32(missionData.unitsBuilt))
    ini.setValue "unitsKilled", json_value(U32(missionData.unitsKilled))
    ini.setValue "unitsLost", json_value(U32(missionData.unitsLost))
    ini.setValue "strBuilt", json_value(U32(missionData.strBuilt))
    ini.setValue "strKilled", json_value(U32(missionData.strKilled))
    ini.setValue "strLost", json_value(U32(missionData.strLost))
    ini.setValue "artefactsFound", json_value(U32(missionData.artefactsFound))
    ini.setValue "missionStarted", json_value(U32(missionData.missionStarted))
    ini.setValue "missionEnded", json_value(U32(missionData.missionEnded))
    ini.setValue "shotsOnTarget", json_value(U32(missionData.shotsOnTarget))
    ini.setValue "shotsOffTarget", json_value(U32(missionData.shotsOffTarget))
    ini.setValue "babasMowedDown", json_value(U32(missionData.babasMowedDown))

    ' Everything is just fine!
    writeScoreData = True
End Function

' -----------------------------------------------------------------------------------
' This will read in the score data
Public Function readScoreData(ByVal fileName As String) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnly)

    ' Retrieve the score data for the current player
    missionData.unitsBuilt = ini.value("unitsBuilt").toInt()
    missionData.unitsKilled = ini.value("unitsKilled").toInt()
    missionData.unitsLost = ini.value("unitsLost").toInt()
    missionData.strBuilt = ini.value("strBuilt").toInt()
    missionData.strKilled = ini.value("strKilled").toInt()
    missionData.strLost = ini.value("strLost").toInt()
    missionData.artefactsFound = ini.value("artefactsFound").toInt()
    missionData.missionStarted = ini.value("missionStarted").toInt()
    missionData.missionEnded = ini.value("missionEnded").toInt()
    missionData.shotsOnTarget = ini.value("shotsOnTarget").toInt()
    missionData.shotsOffTarget = ini.value("shotsOffTarget").toInt()
    missionData.babasMowedDown = ini.value("babasMowedDown").toInt()

    ' Hopefully everything's just fine by now
    readScoreData = True
End Function

Public Sub stdOutGameSummary(Optional ByVal realTimeThrottleSeconds As Long = 5, Optional ByVal flush_output As Boolean = True)
    Dim n As Long, nMax As Long, unitsKilled As Long, numUnits As Long, numStructs As Long, numFactories As Long, numResearch As Long
    Dim numFactoriesThatCanProduceConstructionUnits As Long, it As Long, psStruct As Long, structInfoString As String
    Dim playerCantDoAnything As Boolean, deadStatus As String, t As STRUCTURE_TYPE
    If realTimeThrottleSeconds > 0 And (U32(realTime - lastOutputRealTime) < U32(realTimeThrottleSeconds) * GAME_TICKS_PER_SEC) Then
        Exit Sub
    End If
    fprintf STDOUT_FILE, "Game State [gameTime: %u]\n", gameTime
    fprintf STDOUT_FILE, "--------------------------------------------------------------------------------------\n"
    If ActivityManager_getCurrentGameMode() <> ActivitySink_GameMode_CAMPAIGN Then
        fprintf STDOUT_FILE, " # | Player Name | Extrct Pwr | Units Killed | Structs (F/R) | Units Alive |  Power  |\n"
        fprintf STDOUT_FILE, "-- | ----------- | ---------- | ------------ | ------------- | ----------- | ------- |\n"
        nMax = MAX_PLAYERS
        If game.maxPlayers < nMax Then nMax = game.maxPlayers
        For n = 0 To nMax - 1
            If NetPlay_players_ai(n) < 0 Then
                ' skip closed slots / open / empty slots
                GoTo continue_
            End If
            unitsKilled = getMultiPlayUnitsKilled(n)
            numUnits = apsDroidLists(n).size
            numStructs = apsStructLists(n).size
            numFactories = 0
            numResearch = 0
            numFactoriesThatCanProduceConstructionUnits = 0
            it = apsStructLists(n).head
            Do While it <> 0
                psStruct = lnObj(it)
                it = lnNext(it)
                If aStruct(psStruct).status <> SS_BUILT Or aStruct(psStruct).died <> 0 Then
                    GoTo continueStruct     ' ignore structures that aren't completely built, or are "dead"
                End If
                t = asStructureStats(aStruct(psStruct).pStructureType).typ
                If STRUCTURE_isFactory(psStruct) Then
                    numFactories = numFactories + 1
                    If t = REF_FACTORY Or t = REF_CYBORG_FACTORY Then
                        numFactoriesThatCanProduceConstructionUnits = numFactoriesThatCanProduceConstructionUnits + 1
                    End If
                ElseIf t = REF_RESEARCH Then
                    numResearch = numResearch + 1
                End If
continueStruct:
            Loop
            structInfoString = CStr(numStructs) & " (" & CStr(numFactories) & "/" & CStr(numResearch) & ")"
            ' NOTE: This duplicates the logic in rules.js - checkEndConditions()
            playerCantDoAnything = (numFactoriesThatCanProduceConstructionUnits = 0) And (numUnits = 0)
            If playerCantDoAnything Then deadStatus = "x" Else deadStatus = ""
            fprintf STDOUT_FILE, "%2u | %11.11s | %10lld | %12d | %13.13s | %11d | %7d | %s\n", n, getPlayerName(n), getExtractedPower(n), unitsKilled, structInfoString, numUnits, getPower(n), deadStatus
continue_:
        Next
    End If
    fprintf STDOUT_FILE, "--------------------------------------------------------------------------------------\n"
    If flush_output Then
        fflush STDOUT_FILE
    End If
    lastOutputRealTime = realTime
End Sub
