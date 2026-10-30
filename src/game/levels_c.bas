Attribute VB_Name = "levels_c"
' src/levels.cpp, levels.h, levelint.h: the level descriptions (gamedesc.lev,
' addon.lev: the lexer is level_lexer_c) and the loading of a level's data
' (levLoadData). A LEVEL_DATASET* is a LEVEL_DATASET object (Nothing:
' nullptr); LEVEL_LIST psLevels is psLevels(0 .. psLevelsCount - 1); a
' Sha256 const *hash is the hash with a Boolean (hasHash false: nullptr);
' a char* that may be nullptr is vbNullString for it.
' gamedef.h's enum GAME_TYPE is here too.
' levLoadData preloads the faction IMDs in the order the models were loaded
' (enumerateLoadedModels; C++: hash order).
' Missing: levParse_JSON and levAddWzMap (wzmaplib's level details and map
' packages, not ported yet); the ASSERT on the C library's LC_NUMERIC locale
' (the port parses numbers without it).
Option Explicit

' levels.h: maximum number of data files
Public Const LEVEL_MAXFILES As Long = 9

' levels.h: enum class LEVEL_TYPE : uint8_t (types of level datasets)
Public Const LEVEL_TYPE_LDS_COMPLETE As Long = 0           ' all data required for a stand alone level
Public Const LEVEL_TYPE_LDS_CAMPAIGN As Long = 1           ' the data set for a campaign (no map data)
Public Const LEVEL_TYPE_LDS_CAMSTART As Long = 2           ' mapdata for the start of a campaign
Public Const LEVEL_TYPE_LDS_CAMCHANGE As Long = 3          ' data for changing between levels
Public Const LEVEL_TYPE_LDS_EXPAND As Long = 4             ' extra data for expanding a campaign map
Public Const LEVEL_TYPE_LDS_BETWEEN As Long = 5            ' pause between missions
Public Const LEVEL_TYPE_LDS_MKEEP As Long = 6              ' off map mission (extra map data)
Public Const LEVEL_TYPE_LDS_MCLEAR As Long = 7             ' off map mission (extra map data)
Public Const LEVEL_TYPE_LDS_EXPAND_LIMBO As Long = 8       ' expand campaign map using droids held in apsLimboDroids
Public Const LEVEL_TYPE_LDS_MKEEP_LIMBO As Long = 9        ' off map saving any droids (selectedPlayer) at end into apsLimboDroids
Public Const LEVEL_TYPE_LDS_NONE As Long = 10              ' flags when not got a mission to go back to or when
                                                           ' already on one - ****LEAVE AS LAST ONE****
Public Const LEVEL_TYPE_LDS_MULTI_TYPE_START As Long = 11  ' Start number for custom type numbers (as used by a `type` instruction)
Public Const LEVEL_TYPE_CAMPAIGN As Long = 12
Public Const LEVEL_TYPE_SKIRMISH As Long = 14
Public Const LEVEL_TYPE_MULTI_SKIRMISH2 As Long = 18
Public Const LEVEL_TYPE_MULTI_SKIRMISH3 As Long = 19
Public Const LEVEL_TYPE_MULTI_SKIRMISH4 As Long = 20

' gamedef.h: enum GAME_TYPE
Public Const GTYPE_SCENARIO_START As Long = 0   ' Initial scenario state.
Public Const GTYPE_SCENARIO_EXPAND As Long = 1  ' Scenario scroll area expansion.
Public Const GTYPE_MISSION As Long = 2          ' Stand alone mission.
Public Const GTYPE_SAVE_START As Long = 3       ' User saved game - at the start of a level.
Public Const GTYPE_SAVE_MIDMISSION As Long = 4  ' User saved game - in the middle of a level

' block ID number start for the current level data (as opposed to a dataset)
Private Const CURRENT_DATAID As Long = LEVEL_MAXFILES

' modes for the parser (enum LEVELPARSER_STATE)
Private Const LP_START As Long = 0      ' no input received
Private Const LP_LEVEL As Long = 1      ' level token received
Private Const LP_LEVELDONE As Long = 2  ' defined a level waiting for players/type/data
Private Const LP_PLAYERS As Long = 3    ' players token received
Private Const LP_TYPE As Long = 4       ' type token received
Private Const LP_DATASET As Long = 5    ' dataset token received
Private Const LP_WAITDATA As Long = 6   ' defining level data, waiting for data token
Private Const LP_DATA As Long = 7       ' data token received
Private Const LP_GAME As Long = 8       ' game token received

Private currentLevelName As String      ' char[32], "main"

' the current level descriptions
Public psLevels() As LEVEL_DATASET
Public psLevelsCount As Long

' the currently loaded data set
Private psBaseData As LEVEL_DATASET
Private psCurrLevel As LEVEL_DATASET

' dummy level data for single WRF loads
Private sSingleWRF As LEVEL_DATASET

' return values from the lexer
Public pLevToken As String
Public levVal As Long                   ' LEVEL_TYPE
Private levelLoadType As Long           ' GAME_TYPE

Private statics As Boolean

Private Sub InitStatics()
    currentLevelName = "main"
    Set sSingleWRF = New LEVEL_DATASET
    statics = True
End Sub

' psLevels.push_back(level)
Private Sub PushLevel(ByVal level As LEVEL_DATASET)
    If psLevelsCount = 0 Then
        ReDim psLevels(0 To 63)
    ElseIf psLevelsCount > UBound(psLevels) Then
        ReDim Preserve psLevels(0 To 2 * psLevelsCount - 1)
    End If
    Set psLevels(psLevelsCount) = level
    psLevelsCount = psLevelsCount + 1
End Sub

' psLevels.erase(psLevels.begin() + i)
Private Sub EraseLevel(ByVal i As Long)
    Dim k As Long
    For k = i To psLevelsCount - 2
        Set psLevels(k) = psLevels(k + 1)
    Next
    psLevelsCount = psLevelsCount - 1
    Set psLevels(psLevelsCount) = Nothing
End Sub

' initialise the level system
Public Function levInitialise() As Boolean
    If Not statics Then InitStatics
    Erase psLevels
    psLevelsCount = 0
    Set psBaseData = Nothing
    Set psCurrLevel = Nothing

    levInitialise = True
End Function

Public Function getLevelLoadType() As Long
    getLevelLoadType = levelLoadType
End Function

Private Sub freeLevel(ByVal toDelete As LEVEL_DATASET)
    toDelete.reset
End Sub

' shutdown the level system
Public Sub levShutDown()
    Dim i As Long
    For i = 0 To psLevelsCount - 1
        freeLevel psLevels(i)
    Next
    Erase psLevels
    psLevelsCount = 0
End Sub

' error report function for the level parser
Public Sub lev_error(ByVal msg As String)
    debug_ 161, LOG_ERROR, "lev_error", "Level File parse error: `%s` at line `%d` text `%s`", msg, lev_get_lineno(), lev_get_text()
End Sub

' Find a level dataset with the given name (hash nullptr).
' @return a dataset with associated with the given @c name, or Nothing if none
'         could be found.
Public Function levFindDataSet(ByVal name_ As String) As LEVEL_DATASET
    Dim zero As Sha256
    Set levFindDataSet = levFindDataSet_hash(name_, zero, False)
End Function

' levFindDataSet(name, hash): hasHash false for nullptr
Public Function levFindDataSet_hash(ByVal name_ As String, hash As Sha256, ByVal hasHash As Boolean) As LEVEL_DATASET
    Dim i As Long, psNewLevel As LEVEL_DATASET, h As Sha256
    If hasHash Then
        If Sha256_isZero(hash) Then
            hasHash = False  ' Don't check hash if it's just 0000000000000000000000000000000000000000000000000000000000000000. Assuming real map files probably won't have that particular SHA-256 hash.
        End If
    End If

    For i = 0 To psLevelsCount - 1
        Set psNewLevel = psLevels(i)
        If Len(psNewLevel.pName) > 0 And psNewLevel.pName = name_ Then
            If Not hasHash Then
                Set levFindDataSet_hash = psNewLevel
                Exit Function
            End If
            h = levGetFileHash(psNewLevel)
            If Sha256_equals(h, hash) Then
                Set levFindDataSet_hash = psNewLevel
                Exit Function
            End If
        End If
    Next
End Function

Public Function levFindDataSetByRealFileName(ByVal realFileName As String, hash As Sha256, ByVal hasHash As Boolean) As LEVEL_DATASET
    Dim i As Long, psNewLevel As LEVEL_DATASET, h As Sha256
    If hasHash Then
        If Sha256_isZero(hash) Then
            hasHash = False  ' Don't check hash if it's just 0000000000000000000000000000000000000000000000000000000000000000. Assuming real map files probably won't have that particular SHA-256 hash.
        End If
    End If

    For i = 0 To psLevelsCount - 1
        Set psNewLevel = psLevels(i)
        If StrPtr(psNewLevel.realFileName) <> 0 Then
            If psNewLevel.realFileName = realFileName Then
                If Not hasHash Then
                    Set levFindDataSetByRealFileName = psNewLevel
                    Exit Function
                End If
                h = levGetFileHash(psNewLevel)
                If Sha256_equals(h, hash) Then
                    Set levFindDataSetByRealFileName = psNewLevel
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function levRemoveDataSetByRealFileName(ByVal realFileName As String, hash As Sha256, ByVal hasHash As Boolean) As Boolean
    Dim numRemoved As Long, i As Long, level As LEVEL_DATASET, h As Sha256, matches As Boolean
    If hasHash Then
        If Sha256_isZero(hash) Then
            hasHash = False  ' Don't check hash if it's just 0000000000000000000000000000000000000000000000000000000000000000. Assuming real map files probably won't have that particular SHA-256 hash.
        End If
    End If

    numRemoved = 0
    i = 0
    Do While i < psLevelsCount
        Set level = psLevels(i)
        If Not level Is Nothing Then
            If StrPtr(level.realFileName) <> 0 Then
                If level.realFileName = realFileName Then
                    matches = Not hasHash
                    If Not matches Then
                        h = levGetFileHash(level)
                        matches = Sha256_equals(h, hash)
                    End If
                    If matches Then
                        If psCurrLevel Is psLevels(i) Then
                            ASSERT_FAILURE "false", "src/levels.cpp:228", "levRemoveDataSetByRealFileName", "Trying to remove what is still the current level"
                            GoTo continue_  ' (the C++ loop does not advance here)
                        End If

                        EraseLevel i
                        freeLevel level
                        numRemoved = numRemoved + 1
                        GoTo continue_
                    End If
                End If
            End If
        End If
        i = i + 1
continue_:
    Loop

    levRemoveDataSetByRealFileName = numRemoved > 0
End Function

Public Function levGetFileHash(ByVal level As LEVEL_DATASET) As Sha256
    Dim h As Sha256
    h = level.realFileHash
    If StrPtr(level.realFileName) <> 0 And Sha256_isZero(h) Then
        findHashOfFile level.realFileName, h
        level.realFileHash = h
        debug_ 250, LOG_WZ, "levGetFileHash", "Hash of file ""%s"" is %s.", level.realFileName, Sha256_toString(h)
    End If
    levGetFileHash = level.realFileHash
End Function

Public Function levSetFileHashByRealFileName(ByVal realFileName As String, hash As Sha256) As Boolean
    Dim numAffected As Long, i As Long, psNewLevel As LEVEL_DATASET, h As Sha256
    numAffected = 0
    For i = 0 To psLevelsCount - 1
        Set psNewLevel = psLevels(i)
        If Not psNewLevel Is Nothing Then
            If StrPtr(psNewLevel.realFileName) <> 0 Then
                If psNewLevel.realFileName = realFileName Then
                    h = psNewLevel.realFileHash
                    If Not (Sha256_isZero(h)) Then ASSERT_FAILURE "psNewLevel->realFileHash.isZero()", "src/levels.cpp:262", "levSetFileHashByRealFileName", "Level already has a hash??"
                    psNewLevel.realFileHash = hash
                    numAffected = numAffected + 1
                End If
            End If
        End If
    Next
    levSetFileHashByRealFileName = numAffected > 0
End Function

Public Function levGetMapNameHash(ByVal mapName As String) As Sha256
    Dim level As LEVEL_DATASET, zero As Sha256
    Set level = levFindDataSet(mapName)
    If level Is Nothing Then
        debug_ 275, LOG_WARNING, "levGetMapNameHash", "Couldn't find map ""%s"" to hash.", mapName
        Sha256_setZero zero
        levGetMapNameHash = zero
        Exit Function
    End If
    levGetMapNameHash = levGetFileHash(level)
End Function

Public Function levFindBaseTileset(ByVal tileset As MAP_TILESET) As LEVEL_DATASET
    Select Case tileset
    Case MAP_TILESET_ARIZONA
        Set levFindBaseTileset = levFindDataSet("MULTI_CAM_1")
    Case MAP_TILESET_URBAN
        Set levFindBaseTileset = levFindDataSet("MULTI_CAM_2")
    Case MAP_TILESET_ROCKIES
        Set levFindBaseTileset = levFindDataSet("MULTI_CAM_3")
    End Select
End Function

' parse a level description data file
' the ignoreWrf hack is for compatibility with old maps that try to link in various
' data files that we have removed
' (buffer(0 .. size - 1); realFileName vbNullString: nullptr)
Public Function levParse(buffer() As Byte, ByVal size As Long, ByVal pathMode As Long, ByVal ignoreWrf As Boolean, ByVal realFileName As String) As Boolean
    Dim state As Long, token As Long, currData As Long, psDataSet As LEVEL_DATASET, psFoundData As LEVEL_DATASET
    Dim zero As Sha256
    currData = -1

    lev_set_extra buffer, size

    state = LP_START

    token = lev_lex()
    Do While token <> 0
        Select Case token
        Case LTK_LEVEL, LTK_CAMPAIGN, LTK_CAMSTART, LTK_CAMCHANGE, LTK_EXPAND, LTK_BETWEEN, LTK_MKEEP, LTK_MCLEAR, LTK_EXPAND_LIMBO, LTK_MKEEP_LIMBO
            If state = LP_START Or state = LP_WAITDATA Then
                If Not psDataSet Is Nothing Then
                    ' push the previous level onto the level list
                    PushLevel psDataSet
                    Set psDataSet = Nothing
                End If

                ' start a new level data set
                Set psDataSet = New LEVEL_DATASET
                psDataSet.players = 1
                psDataSet.game_ = -1
                psDataSet.dataDir_ = pathMode
                psDataSet.realFileName = realFileName
                Sha256_setZero zero
                psDataSet.realFileHash = zero   ' The hash is only calculated on demand; for example, if the map name matches.
                currData = 0

                ' set the dataset type
                Select Case token
                Case LTK_LEVEL
                    psDataSet.typ = LEVEL_TYPE_LDS_COMPLETE
                Case LTK_CAMPAIGN
                    psDataSet.typ = LEVEL_TYPE_LDS_CAMPAIGN
                Case LTK_CAMSTART
                    psDataSet.typ = LEVEL_TYPE_LDS_CAMSTART
                Case LTK_BETWEEN
                    psDataSet.typ = LEVEL_TYPE_LDS_BETWEEN
                Case LTK_MKEEP
                    psDataSet.typ = LEVEL_TYPE_LDS_MKEEP
                Case LTK_CAMCHANGE
                    psDataSet.typ = LEVEL_TYPE_LDS_CAMCHANGE
                Case LTK_EXPAND
                    psDataSet.typ = LEVEL_TYPE_LDS_EXPAND
                Case LTK_MCLEAR
                    psDataSet.typ = LEVEL_TYPE_LDS_MCLEAR
                Case LTK_EXPAND_LIMBO
                    psDataSet.typ = LEVEL_TYPE_LDS_EXPAND_LIMBO
                Case LTK_MKEEP_LIMBO
                    psDataSet.typ = LEVEL_TYPE_LDS_MKEEP_LIMBO
                Case Else
                    ASSERT_FAILURE "false", "src/levels.cpp:487", "levParse", "eh?"
                End Select
            Else
                lev_error "Syntax Error"
                GoTo fail
            End If
            state = LP_LEVEL
        Case LTK_PLAYERS
            If state = LP_LEVELDONE And (psDataSet.typ = LEVEL_TYPE_LDS_COMPLETE Or psDataSet.typ >= LEVEL_TYPE_LDS_MULTI_TYPE_START) Then
                state = LP_PLAYERS
            Else
                lev_error "Syntax Error"
                GoTo fail
            End If
        Case LTK_TYPE
            If state = LP_LEVELDONE And psDataSet.typ = LEVEL_TYPE_LDS_COMPLETE Then
                state = LP_TYPE
            Else
                lev_error "Syntax Error"
                GoTo fail
            End If
        Case LTK_INTEGER
            If state = LP_PLAYERS Then
                psDataSet.players = levVal      ' (SWORD)levVal
            ElseIf state = LP_TYPE Then
                If levVal < LEVEL_TYPE_LDS_MULTI_TYPE_START Then
                    lev_error "invalid type number"
                    GoTo fail
                End If

                psDataSet.typ = levVal
            Else
                lev_error "Syntax Error"
                GoTo fail
            End If
            state = LP_LEVELDONE
        Case LTK_DATASET
            If state = LP_LEVELDONE And psDataSet.typ <> LEVEL_TYPE_LDS_COMPLETE Then
                state = LP_DATASET
            Else
                lev_error "Syntax Error"
                GoTo fail
            End If
        Case LTK_DATA
            If state = LP_WAITDATA Then
                state = LP_DATA
            ElseIf state = LP_LEVELDONE Then
                Select Case psDataSet.typ
                Case LEVEL_TYPE_LDS_CAMSTART, LEVEL_TYPE_LDS_MKEEP, LEVEL_TYPE_LDS_CAMCHANGE, LEVEL_TYPE_LDS_EXPAND, LEVEL_TYPE_LDS_MCLEAR, LEVEL_TYPE_LDS_EXPAND_LIMBO, LEVEL_TYPE_LDS_MKEEP_LIMBO
                    lev_error "Missing dataset command"
                    GoTo fail
                End Select
                state = LP_DATA
            Else
                lev_error "Syntax Error"
                GoTo fail
            End If
        Case LTK_GAME
            If (state = LP_WAITDATA Or state = LP_LEVELDONE) And psDataSet.game_ = -1 And psDataSet.typ <> LEVEL_TYPE_LDS_CAMPAIGN Then
                state = LP_GAME
            Else
                lev_error "Syntax Error"
                GoTo fail
            End If
        Case LTK_IDENT
            If state = LP_LEVEL Then
                If psDataSet.typ = LEVEL_TYPE_LDS_CAMCHANGE Then
                    ' This is a campaign change dataset, we need to find the full data set.
                    Set psFoundData = levFindDataSet(pLevToken)

                    If psFoundData Is Nothing Then
                        lev_error "Cannot find full data set for camchange"
                        GoTo fail
                    End If

                    If psFoundData.typ <> LEVEL_TYPE_LDS_CAMSTART Then
                        lev_error "Invalid data set name for cam change"
                        GoTo fail
                    End If
                    Set psFoundData.psChange = psDataSet
                End If
                ' store the level name
                psDataSet.pName = pLevToken

                state = LP_LEVELDONE
            ElseIf state = LP_DATASET Then
                ' find the dataset
                Set psDataSet.psBaseData = levFindDataSet(pLevToken)

                If psDataSet.psBaseData Is Nothing Then
                    lev_error "Unknown dataset"
                    GoTo fail
                End If
                state = LP_WAITDATA
            Else
                lev_error "Syntax Error"
                GoTo fail
            End If
        Case LTK_STRING
            If state = LP_DATA Or state = LP_GAME Then
                If currData >= LEVEL_MAXFILES Then
                    lev_error "Too many data files"
                    GoTo fail
                End If

                ' note the game index if necessary
                If state = LP_GAME Then
                    psDataSet.game_ = currData
                ElseIf ignoreWrf Then
                    state = LP_WAITDATA
                    GoTo next_token ' ignore this wrf line
                End If

                ' store the data name
                psDataSet.apDataFiles(currData) = pLevToken

                resToLower pLevToken

                currData = currData + 1
                state = LP_WAITDATA
            Else
                lev_error "Syntax Error"
                GoTo fail
            End If
        Case Else
            lev_error "Unexpected token"
        End Select
next_token:
        token = lev_lex()
    Loop

    ' Accept empty files when parsing (indicated by currData < 0)
    If currData >= 0 And (state <> LP_WAITDATA Or currData = 0) Then
        lev_error "Unexpected end of file"
        lev_lex_destroy
        levParse = False
        Exit Function
    End If

    If Not psDataSet Is Nothing Then
        PushLevel psDataSet
        Set psDataSet = Nothing
    End If

    lev_lex_destroy
    levParse = True
    Exit Function
fail:
    ' if (psDataSet) { delete psDataSet; }
    If Not psDataSet Is Nothing Then psDataSet.reset
    lev_lex_destroy
    levParse = False
End Function

' free the data for the current mission
Public Function levReleaseMissionData() As Boolean
    Dim i As Long
    ' release old data if any was loaded
    If Not psCurrLevel Is Nothing Then
        If Not stageThreeShutDown() Then
            levReleaseMissionData = False
            Exit Function
        End If

        resDoResLoadCallback       ' do callback.

        ' free up the old data
        For i = LEVEL_MAXFILES - 1 To 0 Step -1
            If i = psCurrLevel.game_ Then
                If psCurrLevel.psBaseData Is Nothing Then
                    If Not stageTwoShutDown() Then
                        levReleaseMissionData = False
                        Exit Function
                    End If
                End If
            Else ' if (psCurrLevel->apDataFiles[i])
                resReleaseBlockData i + CURRENT_DATAID
            End If
        Next
    End If
    releaseObjectives = True ' allow releasing mission objectives after quitting / saveload
    levReleaseMissionData = True
End Function

' free the currently loaded dataset
Public Function levReleaseAll(Optional ByVal forceOnError As Boolean = False) As Boolean
    Dim didLoadLevel As Boolean, didLoadBaseLevel As Boolean, i As Long, releaseBase As Boolean
    ' clear out old effect data first
    initEffectsSystem

    ' release old data if any was loaded
    didLoadLevel = Not psCurrLevel Is Nothing
    didLoadBaseLevel = Not psBaseData Is Nothing

    If didLoadLevel Then
        resDoResLoadCallback       ' do callback.

        If Not levReleaseMissionData() Then
            debug_ 764, LOG_ERROR, "levReleaseAll", "Failed to unload mission data"
            levReleaseAll = False
            Exit Function
        End If
    End If

    If didLoadLevel Or didLoadBaseLevel Or forceOnError Then
        resDoResLoadCallback       ' do callback.

        ' release the base game data
        releaseBase = Not psBaseData Is Nothing
        If Not releaseBase And Not psCurrLevel Is Nothing Then releaseBase = Not psCurrLevel.psBaseData Is Nothing
        If releaseBase Then
            If Not stageTwoShutDown() Then
                debug_ 778, LOG_ERROR, "levReleaseAll", "Failed stage two shutdown"
                levReleaseAll = False
                Exit Function
            End If
        End If

        If Not psBaseData Is Nothing Then
            For i = LEVEL_MAXFILES - 1 To 0 Step -1
                If Len(psBaseData.apDataFiles(i)) > 0 Then
                    resReleaseBlockData i
                End If
            Next
        End If

        If Not psCurrLevel Is Nothing Then
            If Not psCurrLevel.psBaseData Is Nothing Then
                If Not psCurrLevel.psBaseData Is psBaseData Then ' can this even happen?
                    debug_ 796, LOG_INFO, "levReleaseAll", "psCurrLevel->psBaseData != psBaseData"
                    For i = LEVEL_MAXFILES - 1 To 0 Step -1
                        If Len(psCurrLevel.psBaseData.apDataFiles(i)) > 0 Then
                            resReleaseBlockData i
                        End If
                    Next
                End If
            End If
        End If

        Set psBaseData = Nothing
        Set psCurrLevel = Nothing
    End If

    If didLoadLevel Or didLoadBaseLevel Or forceOnError Then
        If Not stageOneShutDown() Then
            debug_ 814, LOG_ERROR, "levReleaseAll", "Failed stage one shutdown"
            levReleaseAll = False
            Exit Function
        End If
    End If

    levReleaseAll = True
End Function

' load up a single wrf file
Private Function levLoadSingleWRF(ByVal name_ As String) As Boolean
    ' free the old data
    If Not levReleaseAll() Then
        levLoadSingleWRF = False
        Exit Function
    End If

    ' create the dummy level data
    sSingleWRF.reset
    sSingleWRF.pName = name_

    ' load up the WRF
    If Not stageOneInitialise() Then
        levLoadSingleWRF = False
        Exit Function
    End If

    ' load the data
    debug_ 842, LOG_WZ, "levLoadSingleWRF", "Loading %s ...", name_
    If Not resLoad(name_, 0) Then
        levLoadSingleWRF = False
        Exit Function
    End If

    If Not stageThreeInitialise() Then
        levLoadSingleWRF = False
        Exit Function
    End If

    Set psCurrLevel = sSingleWRF

    levLoadSingleWRF = True
End Function

Public Function getLevelName() As String
    If Not statics Then InitStatics
    getLevelName = currentLevelName
End Function

Private Function levConstructGameLoadDetails(ByVal psNewLevel As LEVEL_DATASET, ByVal scenarioIndex As Long) As GameLoadDetails
    If StrPtr(psNewLevel.realFileName) <> 0 Then
        Set levConstructGameLoadDetails = GameLoadDetails_makeMapPackageLoad(psNewLevel.realFileName)
        Exit Function
    End If
    Set levConstructGameLoadDetails = GameLoadDetails_makeLevelFileLoad(psNewLevel.apDataFiles(scenarioIndex))
End Function

' load up the data for a level (hasHash false: hash nullptr; pSaveName vbNullString: nullptr)
Public Function levLoadData(ByVal name_ As String, hash As Sha256, ByVal hasHash As Boolean, ByVal pSaveName As String, ByVal saveType As Long) As Boolean
    Dim psNewLevel As LEVEL_DATASET, psChangeLevel As LEVEL_DATASET, bCamChangeSaveGame As Boolean
    Dim i As Long, hasSaveName As Boolean, hashStr As String, buf As String, baseDiffers As Boolean
    Dim enabledNonNormalFactions() As Long, numEnabledNonNormalFactions As Long
    Dim factionModels() As String, normalModels() As String, nFactionModels As Long, f As Long, modelName_ As String, factionModel As String, retval As Long
    If Not statics Then InitStatics
    hasSaveName = (StrPtr(pSaveName) <> 0)
    If hasHash Then hashStr = Sha256_toString(hash) Else hashStr = "builtin"
    debug_ 875, LOG_WZ, "levLoadData", "Loading level %s hash %s (%s, type %d)", name_, hashStr, IIf(hasSaveName, pSaveName, "(null)"), saveType
    If saveType = GTYPE_SAVE_START Or saveType = GTYPE_SAVE_MIDMISSION Then
        If Not levReleaseAll() Then
            debug_ 880, LOG_ERROR, "levLoadData", "Failed to unload old data"
            levLoadData = False
            Exit Function
        End If
    End If

    ' Ensure that the LC_NUMERIC locale setting is "C" (the port's number parsing does not use the C locale)

    levelLoadType = saveType

    ' find the level dataset
    Set psNewLevel = levFindDataSet_hash(name_, hash, hasHash)
    If psNewLevel Is Nothing Then
        debug_ 894, LOG_INFO, "levLoadData", "Dataset %s not found - trying to load as WRF", name_
        levLoadData = levLoadSingleWRF(name_)
        Exit Function
    End If
    debug_ 897, LOG_WZ, "levLoadData", "** Data set found is %s type %d", psNewLevel.pName, psNewLevel.typ

    ' Keep a copy of the present level name
    If Len(name_) > 31 Then currentLevelName = Left$(name_, 31) Else currentLevelName = name_

    bCamChangeSaveGame = hasSaveName And saveType = GTYPE_SAVE_START And Not psNewLevel.psChange Is Nothing
    If bCamChangeSaveGame Then
        debug_ 905, LOG_WZ, "levLoadData", "** CAMCHANGE FOUND"
    End If

    ' select the change dataset if there is one
    Set psChangeLevel = Nothing
    If (Not psNewLevel.psChange Is Nothing And Not psCurrLevel Is Nothing) Or bCamChangeSaveGame Then
        ' store the level name
        debug_ 913, LOG_WZ, "levLoadData", "Found CAMCHANGE dataset"
        Set psChangeLevel = psNewLevel
        Set psNewLevel = psNewLevel.psChange
    End If

    ' ensure the correct dataset is loaded
    If psNewLevel.typ = LEVEL_TYPE_LDS_CAMPAIGN Then
        debug_ 921, LOG_ERROR, "levLoadData", "Cannot load a campaign dataset (%s)", psNewLevel.pName
        levLoadData = False
        Exit Function
    Else
        If Not psCurrLevel Is Nothing Then
            baseDiffers = Not (psCurrLevel.psBaseData Is psNewLevel.psBaseData)
            If baseDiffers Or (psCurrLevel.typ < LEVEL_TYPE_LDS_NONE And psNewLevel.typ >= LEVEL_TYPE_LDS_NONE) Or (psCurrLevel.typ >= LEVEL_TYPE_LDS_NONE And psNewLevel.typ < LEVEL_TYPE_LDS_NONE) Then
                ' there is a dataset loaded but it isn't the correct one
                debug_ 933, LOG_WZ, "levLoadData", "Incorrect base dataset loaded (%p != %p, %d - %d)", ObjPtr(psCurrLevel.psBaseData), ObjPtr(psNewLevel.psBaseData), psCurrLevel.typ, psNewLevel.typ
                If Not levReleaseAll() Then ' this sets psCurrLevel to NULL
                    debug_ 937, LOG_ERROR, "levLoadData", "Failed to release old data"
                    levLoadData = False
                    Exit Function
                End If
            Else
                debug_ 943, LOG_WZ, "levLoadData", "Correct base dataset already loaded."
            End If
        End If

        ' setup the correct dataset to load if necessary
        If psCurrLevel Is Nothing Then
            If Not psNewLevel.psBaseData Is Nothing Then
                debug_ 952, LOG_WZ, "levLoadData", "Setting base dataset to load: %s", psNewLevel.psBaseData.pName
            End If
            Set psBaseData = psNewLevel.psBaseData
        Else
            debug_ 958, LOG_WZ, "levLoadData", "No base dataset to load"
            Set psBaseData = Nothing
        End If
    End If

    If Not rebuildSearchPath(psNewLevel.dataDir_, True) Then
        debug_ 965, LOG_ERROR, "levLoadData", "Failed to rebuild search path"
        levLoadData = False
        Exit Function
    End If

    ' reset the old mission data if necessary
    If Not psCurrLevel Is Nothing Then
        debug_ 972, LOG_WZ, "levLoadData", "Resetting old mission data"
        If Not levReleaseMissionData() Then
            debug_ 975, LOG_ERROR, "levLoadData", "Failed to unload old mission data"
            levLoadData = False
            Exit Function
        End If
    End If

    ' need to free the current map and droids etc for a save game
    If psBaseData Is Nothing And hasSaveName Then
        If Not saveGameReset() Then
            debug_ 985, LOG_ERROR, "levLoadData", "Failed to saveGameReset()!"
            levLoadData = False
            Exit Function
        End If
    End If

    ' initialise if necessary
    If psNewLevel.typ = LEVEL_TYPE_LDS_COMPLETE Or Not psBaseData Is Nothing Then
        debug_ 993, LOG_WZ, "levLoadData", "Calling stageOneInitialise!"
        If Not stageOneInitialise() Then
            debug_ 996, LOG_ERROR, "levLoadData", "Failed stageOneInitialise!"
            levLoadData = False
            Exit Function
        End If
    End If

    ' load up a base dataset if necessary
    If Not psBaseData Is Nothing Then
        debug_ 1004, LOG_WZ, "levLoadData", "Loading base dataset %s", psBaseData.pName
        For i = 0 To LEVEL_MAXFILES - 1
            If Len(psBaseData.apDataFiles(i)) > 0 Then
                ' load the data
                debug_ 1010, LOG_WZ, "levLoadData", "Loading [directory: %s] %s ...", PHYSFS_getRealDir(psBaseData.apDataFiles(i)) & "", psBaseData.apDataFiles(i)
                If Not resLoad(psBaseData.apDataFiles(i), i) Then
                    debug_ 1013, LOG_ERROR, "levLoadData", "Failed resLoad(%s)!", psBaseData.apDataFiles(i)
                    levLoadData = False
                    Exit Function
                End If
            End If
        Next
    End If
    ' preload faction IMDs
    numEnabledNonNormalFactions = getEnabledFactions(True, enabledNonNormalFactions)
    If numEnabledNonNormalFactions > 0 Then
        ' factionModelsToLoad: { factionModel, normalModel } (enumerateLoadedModels: in the order the models were loaded)
        nFactionModels = 0
        For i = 0 To enumerateLoadedModels_count() - 1
            enumerateLoadedModels_at i, modelName_
            For f = 0 To numEnabledNonNormalFactions - 1
                If getFactionModelName(enabledNonNormalFactions(f), modelName_, factionModel) Then
                    ReDim Preserve factionModels(0 To nFactionModels)
                    ReDim Preserve normalModels(0 To nFactionModels)
                    factionModels(nFactionModels) = factionModel
                    normalModels(nFactionModels) = modelName_
                    nFactionModels = nFactionModels + 1
                End If
            Next
        Next
        For i = 0 To nFactionModels - 1
            retval = modelGet(factionModels(i))
            If Not (retval <> 0) Then ASSERT_FAILURE "retval != nullptr", "src/levels.cpp:1043", "levLoadData", "Cannot find the faction PIE model %s (for normal model: %s)", _
                factionModels(i), normalModels(i)
        Next
        resDoResLoadCallback       ' do callback.
    End If

    If psNewLevel.typ = LEVEL_TYPE_LDS_CAMCHANGE Then
        If Not campaignReset() Then
            debug_ 1053, LOG_ERROR, "levLoadData", "Failed campaignReset()!"
            levLoadData = False
            Exit Function
        End If
    End If
    If psNewLevel.game_ = -1 Then  'no .gam file to load - BETWEEN missions (for Editor games only)
        If Not (psNewLevel.typ = LEVEL_TYPE_LDS_BETWEEN) Then ASSERT_FAILURE "psNewLevel->type == LEVEL_TYPE::LDS_BETWEEN", "src/levels.cpp:1059", "levLoadData", "Only BETWEEN missions do not need a .gam file"
        debug_ 1060, LOG_WZ, "levLoadData", "No .gam file for level: BETWEEN mission"
        If hasSaveName Then
            If Not psBaseData Is Nothing Then
                If Not stageTwoInitialise() Then
                    debug_ 1067, LOG_ERROR, "levLoadData", "Failed stageTwoInitialise()!"
                    levLoadData = False
                    Exit Function
                End If
            End If

            'set the mission type before the saveGame data is loaded
            If saveType = GTYPE_SAVE_MIDMISSION Then
                debug_ 1075, LOG_WZ, "levLoadData", "Init mission stuff"
                If Not startMissionSave(psNewLevel.typ) Then
                    debug_ 1078, LOG_ERROR, "levLoadData", "Failed startMissionSave(%d)!", ToInt8(psNewLevel.typ)
                    levLoadData = False
                    Exit Function
                End If

                debug_ 1082, LOG_NEVER, "levLoadData", "dataSetSaveFlag"
                dataSetSaveFlag
            End If

            debug_ 1086, LOG_NEVER, "levLoadData", "Loading savegame: %s", pSaveName
            If Not loadGame(GameLoadDetails_makeUserSaveGameLoad(pSaveName), False, True) Then
                debug_ 1089, LOG_ERROR, "levLoadData", "Failed loadGame(%s)!", pSaveName
                levLoadData = False
                Exit Function
            End If
        End If

        If Not hasSaveName Or saveType = GTYPE_SAVE_START Then
            debug_ 1096, LOG_NEVER, "levLoadData", "Start mission - no .gam"
            If Not startMission(psNewLevel.typ, GameLoadDetails_makeLevelFileLoad("")) Then
                debug_ 1099, LOG_ERROR, "levLoadData", "Failed startMission(%d)!", ToInt8(psNewLevel.typ)
                levLoadData = False
                Exit Function
            End If
        End If
    End If

    'we need to load up the save game data here for a camchange
    If bCamChangeSaveGame Then
        If hasSaveName Then
            If Not psBaseData Is Nothing Then
                If Not stageTwoInitialise() Then
                    debug_ 1114, LOG_ERROR, "levLoadData", "Failed stageTwoInitialise() [camchange]!"
                    levLoadData = False
                    Exit Function
                End If
            End If

            debug_ 1119, LOG_NEVER, "levLoadData", "loading savegame: %s", pSaveName
            If Not loadGame(GameLoadDetails_makeUserSaveGameLoad(pSaveName), False, True) Then
                debug_ 1122, LOG_ERROR, "levLoadData", "Failed loadGame(%s)!", pSaveName
                levLoadData = False
                Exit Function
            End If

            campaignReset
        End If
    End If

    ' load the new data
    debug_ 1132, LOG_NEVER, "levLoadData", "Loading mission dataset: %s", psNewLevel.pName
    For i = 0 To LEVEL_MAXFILES - 1
        If psNewLevel.game_ = i Then
            ' do some more initialising if necessary
            If psNewLevel.typ = LEVEL_TYPE_LDS_COMPLETE Or psNewLevel.typ >= LEVEL_TYPE_LDS_MULTI_TYPE_START Or (Not psBaseData Is Nothing And Not bCamChangeSaveGame) Then
                If Not stageTwoInitialise() Then
                    debug_ 1142, LOG_ERROR, "levLoadData", "Failed stageTwoInitialise() [newdata]!"
                    levLoadData = False
                    Exit Function
                End If
            End If

            ' load a savegame if there is one - but not if already done so
            If hasSaveName And Not bCamChangeSaveGame Then
                'set the mission type before the saveGame data is loaded
                If saveType = GTYPE_SAVE_MIDMISSION Then
                    debug_ 1153, LOG_WZ, "levLoadData", "Init mission stuff"
                    If Not startMissionSave(psNewLevel.typ) Then
                        debug_ 1156, LOG_ERROR, "levLoadData", "Failed startMissionSave(%d)!", psNewLevel.typ
                        levLoadData = False
                        Exit Function
                    End If

                    debug_ 1160, LOG_NEVER, "levLoadData", "dataSetSaveFlag"
                    dataSetSaveFlag
                End If

                debug_ 1164, LOG_NEVER, "levLoadData", "Loading save game %s", pSaveName
                If Not loadGame(GameLoadDetails_makeUserSaveGameLoad(pSaveName), False, True) Then
                    debug_ 1167, LOG_ERROR, "levLoadData", "Failed loadGame(%s)!", pSaveName
                    levLoadData = False
                    Exit Function
                End If
            End If

            If Not hasSaveName Or saveType = GTYPE_SAVE_START Then
                ' load the game
                debug_ 1175, LOG_WZ, "levLoadData", "Loading scenario file %s", psNewLevel.apDataFiles(i)
                Select Case psNewLevel.typ
                Case LEVEL_TYPE_LDS_COMPLETE, LEVEL_TYPE_LDS_CAMSTART
                    debug_ 1180, LOG_WZ, "levLoadData", "LDS_COMPLETE / LDS_CAMSTART"
                    If Not startMission(LEVEL_TYPE_LDS_CAMSTART, levConstructGameLoadDetails(psNewLevel, i)) Then
                        debug_ 1183, LOG_ERROR, "levLoadData", "Failed startMission(%d, %s)!", LEVEL_TYPE_LDS_CAMSTART, psNewLevel.apDataFiles(i)
                        levLoadData = False
                        Exit Function
                    End If
                Case LEVEL_TYPE_LDS_BETWEEN
                    debug_ 1188, LOG_WZ, "levLoadData", "LDS_BETWEEN"
                    If Not startMission(LEVEL_TYPE_LDS_BETWEEN, levConstructGameLoadDetails(psNewLevel, i)) Then
                        debug_ 1191, LOG_ERROR, "levLoadData", "Failed startMission(%d, %s)!", LEVEL_TYPE_LDS_BETWEEN, psNewLevel.apDataFiles(i)
                        levLoadData = False
                        Exit Function
                    End If

                Case LEVEL_TYPE_LDS_MKEEP
                    debug_ 1197, LOG_WZ, "levLoadData", "LDS_MKEEP"
                    If Not startMission(LEVEL_TYPE_LDS_MKEEP, levConstructGameLoadDetails(psNewLevel, i)) Then
                        debug_ 1200, LOG_ERROR, "levLoadData", "Failed startMission(%d, %s)!", LEVEL_TYPE_LDS_MKEEP, psNewLevel.apDataFiles(i)
                        levLoadData = False
                        Exit Function
                    End If
                Case LEVEL_TYPE_LDS_CAMCHANGE
                    debug_ 1205, LOG_WZ, "levLoadData", "LDS_CAMCHANGE"
                    If Not startMission(LEVEL_TYPE_LDS_CAMCHANGE, levConstructGameLoadDetails(psNewLevel, i)) Then
                        debug_ 1208, LOG_ERROR, "levLoadData", "Failed startMission(%d, %s)!", LEVEL_TYPE_LDS_CAMCHANGE, psNewLevel.apDataFiles(i)
                        levLoadData = False
                        Exit Function
                    End If

                Case LEVEL_TYPE_LDS_EXPAND
                    debug_ 1214, LOG_WZ, "levLoadData", "LDS_EXPAND"
                    If Not startMission(LEVEL_TYPE_LDS_EXPAND, levConstructGameLoadDetails(psNewLevel, i)) Then
                        debug_ 1217, LOG_ERROR, "levLoadData", "Failed startMission(%d, %s)!", LEVEL_TYPE_LDS_EXPAND, psNewLevel.apDataFiles(i)
                        levLoadData = False
                        Exit Function
                    End If
                Case LEVEL_TYPE_LDS_EXPAND_LIMBO
                    debug_ 1222, LOG_WZ, "levLoadData", "LDS_LIMBO"
                    If Not startMission(LEVEL_TYPE_LDS_EXPAND_LIMBO, levConstructGameLoadDetails(psNewLevel, i)) Then
                        debug_ 1225, LOG_ERROR, "levLoadData", "Failed startMission(%d, %s)!", LEVEL_TYPE_LDS_EXPAND_LIMBO, psNewLevel.apDataFiles(i)
                        levLoadData = False
                        Exit Function
                    End If

                Case LEVEL_TYPE_LDS_MCLEAR
                    debug_ 1231, LOG_WZ, "levLoadData", "LDS_MCLEAR"
                    If Not startMission(LEVEL_TYPE_LDS_MCLEAR, levConstructGameLoadDetails(psNewLevel, i)) Then
                        debug_ 1234, LOG_ERROR, "levLoadData", "Failed startMission(%d, %s)!", LEVEL_TYPE_LDS_MCLEAR, psNewLevel.apDataFiles(i)
                        levLoadData = False
                        Exit Function
                    End If
                Case LEVEL_TYPE_LDS_MKEEP_LIMBO
                    debug_ 1239, LOG_WZ, "levLoadData", "LDS_MKEEP_LIMBO"
                    If Not startMission(LEVEL_TYPE_LDS_MKEEP_LIMBO, levConstructGameLoadDetails(psNewLevel, i)) Then
                        debug_ 1242, LOG_ERROR, "levLoadData", "Failed startMission(%d, %s)!", LEVEL_TYPE_LDS_MKEEP_LIMBO, psNewLevel.apDataFiles(i)
                        levLoadData = False
                        Exit Function
                    End If
                Case Else
                    If Not (psNewLevel.typ >= LEVEL_TYPE_LDS_MULTI_TYPE_START) Then ASSERT_FAILURE "psNewLevel->type >= LEVEL_TYPE::LDS_MULTI_TYPE_START", "src/levels.cpp:1247", "levLoadData", "Unexpected mission type"
                    debug_ 1248, LOG_WZ, "levLoadData", "default (MULTIPLAYER)"
                    If Not startMission(LEVEL_TYPE_LDS_CAMSTART, levConstructGameLoadDetails(psNewLevel, i)) Then
                        debug_ 1251, LOG_ERROR, "levLoadData", "Failed startMission(%d, %s) (default)!", LEVEL_TYPE_LDS_CAMSTART, psNewLevel.apDataFiles(i)
                        levLoadData = False
                        Exit Function
                    End If
                End Select
            End If
        ElseIf Len(psNewLevel.apDataFiles(i)) > 0 Then
            ' load the data
            debug_ 1261, LOG_WZ, "levLoadData", "Loading %s", psNewLevel.apDataFiles(i)
            If Not resLoad(psNewLevel.apDataFiles(i), i + CURRENT_DATAID) Then
                debug_ 1264, LOG_ERROR, "levLoadData", "Failed resLoad(%s, %d) (default)!", psNewLevel.apDataFiles(i), i + CURRENT_DATAID
                levLoadData = False
                Exit Function
            End If
        End If
    Next

    If bMultiPlayer Then
        ' This calls resLoadFile("SMSG", "multiplay.txt"). Must be before loadMissionExtras, which calls loadSaveMessage, which calls getViewData.
        loadMultiScripts
    End If

    If hasSaveName Then
        'load MidMission Extras
        If Not loadMissionExtras(pSaveName, psNewLevel.typ) Then
            debug_ 1281, LOG_ERROR, "levLoadData", "Failed loadMissionExtras(%s, %d)!", pSaveName, ToInt8(psNewLevel.typ)
            levLoadData = False
            Exit Function
        End If
    End If

    If hasSaveName And saveType = GTYPE_SAVE_MIDMISSION Then
        'load script stuff
        ' load the event system state here for a save game
        debug_ 1290, LOG_SAVE, "levLoadData", "Loading script system state"
        If Not loadScriptState(pSaveName) Then
            debug_ 1293, LOG_ERROR, "levLoadData", "Failed loadScriptState(%s)!", pSaveName
            levLoadData = False
            Exit Function
        End If
    End If
    ' this will trigger upgrades
    If Not stageThreeInitialise() Then
        debug_ 1300, LOG_ERROR, "levLoadData", "Failed stageThreeInitialise()!"
        levLoadData = False
        Exit Function
    End If

    dataClearSaveFlag

    'restore the level name for comparisons on next mission load up
    If psChangeLevel Is Nothing Then
        Set psCurrLevel = psNewLevel
    Else
        Set psCurrLevel = psChangeLevel
    End If

    ' Copy this info to be used by the crash handler for the dump file
    buf = "Current Level/map is " & psCurrLevel.pName
    If Len(buf) > 255 Then buf = Left$(buf, 255)
    addDumpInfo buf

    If autogame_enabled() And getHostLaunch() <> HostLaunch_LoadReplay Then
        gameTimeSetMod 500, 1                   ' Rational(500)
        If getHostLaunch() <> HostLaunch_Skirmish Then ' tests will specify the AI manually
            If ULT(selectedPlayer, MAX_PLAYERS) And Not NetPlay_players_isSpectator(selectedPlayer) Then
                jsAutogameSpecific "multiplay/skirmish/semperfi.js", selectedPlayer, AIDifficulty_DEFAULT
            Else
                debug_ 1333, LOG_INFO, "levLoadData", "Skipping autogame auto-AI for selectedPlayer %u", selectedPlayer
            End If
        End If
    End If

    ActivityManager_loadedLevel psCurrLevel.typ, mapNameWithoutTechlevel(getLevelName())

    levLoadData = True
End Function

Public Function mapNameWithoutTechlevel(ByVal mapName As String) As String
    Dim result As String
    If Not (StrPtr(mapName) <> 0) Then
        ASSERT_FAILURE "mapName != nullptr", "src/levels.cpp:1345", "mapNameWithoutTechlevel", "null mapName provided"
        mapNameWithoutTechlevel = ""
        Exit Function
    End If
    result = mapName
    trimTechLevelFromMapName result
    mapNameWithoutTechlevel = result
End Function

' returns maps of the right 'type' (LEVEL_LIST: list(0 .. count - 1))
Public Function enumerateMultiMaps(ByVal camToUse As Long, ByVal numPlayers As Long, list() As LEVEL_DATASET) As Long
    Dim n As Long, i As Long, k As Long, lev As LEVEL_DATASET, cam As Long, already_added As Boolean
    Dim levelBaseName As String, mapBaseName As String

    If game.typ = LEVEL_TYPE_SKIRMISH Then
        ' Add maps with exact match to type
        For i = 0 To psLevelsCount - 1
            Set lev = psLevels(i)
            cam = 1
            If lev.typ = LEVEL_TYPE_MULTI_SKIRMISH2 Then
                cam = 2
            ElseIf lev.typ = LEVEL_TYPE_MULTI_SKIRMISH3 Then
                cam = 3
            ElseIf lev.typ = LEVEL_TYPE_MULTI_SKIRMISH4 Then
                cam = 4
            End If

            If (lev.typ = LEVEL_TYPE_SKIRMISH Or lev.typ = LEVEL_TYPE_MULTI_SKIRMISH2 Or lev.typ = LEVEL_TYPE_MULTI_SKIRMISH3 Or lev.typ = LEVEL_TYPE_MULTI_SKIRMISH4) _
               And (numPlayers = 0 Or numPlayers = lev.players) And cam = camToUse Then
                PushList list, n, lev
            End If
        Next
        ' Also add maps where only the tech level is different, if a more specific map has not been added
        For i = 0 To psLevelsCount - 1
            Set lev = psLevels(i)
            If (lev.typ = LEVEL_TYPE_SKIRMISH Or lev.typ = LEVEL_TYPE_MULTI_SKIRMISH2 Or lev.typ = LEVEL_TYPE_MULTI_SKIRMISH3 Or lev.typ = LEVEL_TYPE_MULTI_SKIRMISH4) _
               And (numPlayers = 0 Or numPlayers = lev.players) And Len(lev.pName) > 0 Then
                already_added = False
                For k = 0 To n - 1
                    If Len(list(k).pName) = 0 Then
                        GoTo continue_
                    End If
                    levelBaseName = mapNameWithoutTechlevel(lev.pName)
                    mapBaseName = mapNameWithoutTechlevel(list(k).pName)
                    If levelBaseName = mapBaseName Then
                        already_added = True
                        Exit For
                    End If
continue_:
                Next
                If Not already_added Then
                    PushList list, n, lev
                End If
            End If
        Next
    End If

    enumerateMultiMaps = n
End Function

Private Sub PushList(list() As LEVEL_DATASET, n As Long, ByVal lev As LEVEL_DATASET)
    If n = 0 Then
        ReDim list(0 To 15)
    ElseIf n > UBound(list) Then
        ReDim Preserve list(0 To 2 * n - 1)
    End If
    Set list(n) = lev
    n = n + 1
End Sub
