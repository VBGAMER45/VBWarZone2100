Attribute VB_Name = "loadsave_c"
' src/loadsave.cpp, loadsave.h, whole, and main.cpp / main.h's SaveGamePath_t (SAVEGAME_LOC, SaveGameLocToPath,
' toPath). The slots' pUserData is a LoadSaveDisplayLoadSlotCache. The lambdas given to WZ_PHYSFS_enumerateFolders
' and WZ_PHYSFS_enumerateFiles are the callers' loops over the names those give (physfs_ext_c); a
' nlohmann::json::exception thrown in them is a VB error (Err.Description its what()), caught by a _try function
' where the C++'s try is. TagVer is version(0 .. 2) (version_c); the comparison is TagVer_lessOrEqual.
' wzAsyncExecOnMainThread is a stand-in (platform's, not ported). strftime's "%F" is written "%Y-%m-%d"
' (ctime_win's strftime has no %F). Translations (_()) are not ported: the English text is used.
'
' loadsave.c
' load and save Popup screens.
'
' these don't actually do any loading or saving, but just
' return a filename to use for the ops.
Option Explicit

Private Const totalslots As Long = 36           ' saves slots
Private Const slotsInColumn As Long = 12        ' # of slots in a column
Private Const totalslotspace As Long = 64       ' guessing 64 max chars for filename.

Private Const LOADSAVE_W As Long = 610
Private Const LOADSAVE_H As Long = 215

Private Const MAX_SAVE_NAME As Long = 60

Private Const LOADSAVE_HGAP As Long = 9
Private Const LOADSAVE_VGAP As Long = 9
Private Const LOADSAVE_BANNER_DEPTH As Long = 40    ' top banner which displays either load or save

Private Const LOADENTRY_W As Long = ((LOADSAVE_W \ 3) - (3 * LOADSAVE_HGAP))
Private Const LOADENTRY_H As Long = (LOADSAVE_H - (5 * LOADSAVE_VGAP) - (LOADSAVE_BANNER_DEPTH + LOADSAVE_VGAP)) \ 5

Private Const ID_LOADSAVE As Long = 21000
Private Const LOADSAVE_FORM As Long = ID_LOADSAVE + 1       ' back form.
Private Const LOADSAVE_CANCEL As Long = ID_LOADSAVE + 2     ' cancel but.
Private Const LOADSAVE_LABEL As Long = ID_LOADSAVE + 3      ' load/save
Private Const LOADSAVE_BANNER As Long = ID_LOADSAVE + 4     ' banner.

Private Const LOADENTRY_START As Long = ID_LOADSAVE + 10                ' each of the buttons.
Private Const LOADENTRY_END As Long = ID_LOADSAVE + 10 + totalslots     ' must have unique ID hmm -Q

Private Const SAVEENTRY_EDIT As Long = ID_LOADSAVE + totalslots + totalslots    ' save edit box. must be highest value possible I guess. -Q

Private Const INT64_MAX_D As Double = 9.22337203685478E+18     ' INT64_MAX (as the int64_t epochs' Double)

' the lambdas (WzFunc: loadsave_c_fn)
Public Const FN_loadsave_displayLoadBanner As Long = 14100      ' (widget, xOffset, yOffset)
Public Const FN_loadsave_displayLoadSlot As Long = 14101        ' (widget, xOffset, yOffset)
Public Const FN_loadSaveForm_calcLayout As Long = 14102         ' (widget)
Public Const FN_newLoadSaveDisplayLoadSlotCache As Long = 14103 ' []() -> void * { return new LoadSaveDisplayLoadSlotCache(); }
Public Const FN_deleteLoadSaveDisplayLoadSlotCache As Long = 14104  ' (widget): the slots' onDelete
Public Const FN_autoSave_doAutoSave As Long = 14105             ' (): autoSave's wzAsyncExecOnMainThread lambda

' main.h
Public Enum SAVEGAME_LOC
    SAVEGAME_LOC_CAM
    SAVEGAME_LOC_CAM_AUTO
    SAVEGAME_LOC_SKI
    SAVEGAME_LOC_SKI_AUTO
End Enum

' main.h: struct SaveGamePath_t's enum class Extension
Public Enum SaveGamePath_t_Extension
    SaveGamePath_t_Extension_GAM
    SaveGamePath_t_Extension_JSON
End Enum

' main.h: struct SaveGamePath_t (its toPath: SaveGamePath_t_toPath)
Public Type SaveGamePath_t
    loc As SAVEGAME_LOC
    ' not a path, just the name
    gameName As String
End Type

'**************************************************************************
'
'	Global Definitions
'
'**************************************************************************

' loadsave.h
Public Enum LOADSAVE_MODE
    LOAD_FRONTEND_MISSION
    LOAD_MISSIONEND
    LOAD_INGAME_MISSION
    LOAD_FRONTEND_SKIRMISH
    LOAD_INGAME_SKIRMISH
    SAVE_MISSIONEND
    SAVE_INGAME_MISSION
    SAVE_INGAME_SKIRMISH
    LOADREPLAY_FRONTEND_SKIRMISH
    LOADREPLAY_FRONTEND_MULTI
    LOAD_FRONTEND_MISSION_AUTO = 16     ' internal use only
    LOAD_MISSIONEND_AUTO                ' internal use only
    LOAD_INGAME_MISSION_AUTO            ' internal use only
    LOAD_FRONTEND_SKIRMISH_AUTO         ' internal use only
    LOAD_INGAME_SKIRMISH_AUTO           ' internal use only
End Enum

' ////////////////////////////////////////////////////////////////////////////
Private psRequestScreen As W_SCREEN     ' static std::shared_ptr<W_SCREEN> psRequestScreen = nullptr; // Widget screen for requester
Private modeLoad As Boolean
Private modeReplay As Boolean
Private modeIngame As Boolean
Private chosenSlotId As Long            ' UDWORD

Public bLoadSaveUp As Boolean           ' true when interface is up and should be run.
Public saveGameName As String           ' char[256]: the name of the save game to load from the front end
Public sRequestResult As String         ' char[PATH_MAX]: filename returned;
Public bRequestLoad As Boolean
Public autosaveEnabled As Boolean       ' = true (loadsave_c_init)
Public bRequestLoadReplay As Boolean
Public bLoadSaveMode As LOADSAVE_MODE
Private savedTitle As String            ' static const char *savedTitle;
Private Const sSaveGameExtension As String = ".gam"
Private Const sSaveReplayExtension As String = ".wzrp"

' addLoadSave's static char sSlotCaps[totalslots][totalslotspace], sSlotTips[totalslots][totalslotspace]
Private sSlotCaps(0 To totalslots - 1) As String
Private sSlotTips(0 To totalslots - 1) As String

Public lastSavePath As SaveGamePath_t
Public lastSaveMP As Boolean
Private lastSaveTime As Double          ' static int64_t lastSaveTime;

' runLoadSave's static char sDelete[PATH_MAX]
Private sDelete As String

' the C++ static initialisers
Public Sub loadsave_c_init()
    autosaveEnabled = True
End Sub

' main.cpp: const char *SaveGameLocToPath[] = { SAVEGAME_CAM, SAVEGAME_CAM_AUTO, SAVEGAME_SKI, SAVEGAME_SKI_AUTO, };
Public Function SaveGameLocToPath(ByVal loc_ As SAVEGAME_LOC) As String
    Select Case loc_
    Case SAVEGAME_LOC_CAM: SaveGameLocToPath = SAVEGAME_CAM
    Case SAVEGAME_LOC_CAM_AUTO: SaveGameLocToPath = SAVEGAME_CAM_AUTO
    Case SAVEGAME_LOC_SKI: SaveGameLocToPath = SAVEGAME_SKI
    Case SAVEGAME_LOC_SKI_AUTO: SaveGameLocToPath = SAVEGAME_SKI_AUTO
    End Select
End Function

' main.cpp: std::string SaveGamePath_t::toPath(SaveGamePath_t::Extension ext)
Public Function SaveGamePath_t_toPath(self As SaveGamePath_t, ByVal ext As SaveGamePath_t_Extension) As String
    Dim out As String
    Select Case ext
    Case SaveGamePath_t_Extension_GAM
        out = SaveGameLocToPath(self.loc) & "/" & self.gameName & ".gam"
    Case SaveGamePath_t_Extension_JSON
        out = SaveGameLocToPath(self.loc) & "/" & self.gameName & ".json"
    End Select
    SaveGamePath_t_toPath = out
End Function

' lastSavePath.toPath(SaveGamePath_t::Extension::GAM) (frontend.cpp)
Public Function lastSavePath_toPath_GAM() As String
    lastSavePath_toPath_GAM = SaveGamePath_t_toPath(lastSavePath, SaveGamePath_t_Extension_GAM)
End Function

' #define LOADSAVE_X D_W + 16
Private Function LOADSAVE_X() As Long
    LOADSAVE_X = D_W() + 16
End Function

' #define LOADSAVE_Y D_H + 5
Private Function LOADSAVE_Y() As Long
    LOADSAVE_Y = D_H() + 5
End Function

' version.h: bool TagVer::operator <= (const TagVer &other) const: *this < other || *this == other
' (operator <: std::lexicographical_compare of the three versions; operator ==: the three equal)
Private Function TagVer_lessOrEqual(a() As Long, b() As Long) As Boolean
    Dim i As Long
    For i = 0 To 2
        If a(i) < b(i) Then
            TagVer_lessOrEqual = True
            Exit Function
        End If
        If b(i) < a(i) Then
            TagVer_lessOrEqual = False
            Exit Function
        End If
    Next
    TagVer_lessOrEqual = True   ' equal
End Function

' physfs_ext.h: time_t WZ_PHYSFS_getLastModTime(const char *filename) (Private in physfs_ext_c: its body)
Private Function WZ_PHYSFS_getLastModTime(ByVal filename As String) As Double
    Dim filesize As Double, modtime As Double, filetype As Long, readonly_ As Long
    PHYSFS_stat filename, filesize, modtime, filetype, readonly_
    WZ_PHYSFS_getLastModTime = modtime
End Function

' ////////////////////////////////////////////////////////////////////////////
' return whether the specified filename looks like a saved game file, i.e. ends with .gam
' Note: this is left for backward compatibiiliy. Remove it later.
Private Function isASavedGamefile(ByVal filename As String, ByVal extension As String) As Boolean
    Dim filenameLength As Long, extensionLength As Long
    If StrPtr(filename) = 0 Then
        isASavedGamefile = False
        Exit Function
    End If

    filenameLength = Len(filename)
    extensionLength = Len(extension)
    ' reject filename of insufficient length to contain "<anything>.gam"
    If extensionLength < filenameLength Then
        isASavedGamefile = StrComp(Right$(filename, extensionLength), extension, vbBinaryCompare) = 0
    End If
End Function

' ////////////////////////////////////////////////////////////////////////////
' return whether the save screen was displayed in the mission results screen
Public Function saveInMissionRes() As Boolean
    saveInMissionRes = bLoadSaveMode = SAVE_MISSIONEND
End Function

' ////////////////////////////////////////////////////////////////////////////
' return whether the save screen was displayed in the middle of a mission
Public Function saveMidMission() As Boolean
    saveMidMission = bLoadSaveMode = SAVE_INGAME_MISSION
End Function

Public Sub loadSaveScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If psRequestScreen Is Nothing Then Exit Sub
    psRequestScreen.screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
End Sub

Private Function getNewSaveGamePathFromMode(ByVal savemode As LOADSAVE_MODE) As String
    Dim result As String

    Select Case savemode
    Case LOAD_FRONTEND_MISSION, LOAD_INGAME_MISSION, LOAD_MISSIONEND
        result = SaveGamePath & "campaign" & "/"
    Case LOAD_FRONTEND_SKIRMISH, LOAD_INGAME_SKIRMISH
        result = SaveGamePath & "skirmish" & "/"
    Case LOAD_FRONTEND_MISSION_AUTO, LOAD_INGAME_MISSION_AUTO, LOAD_MISSIONEND_AUTO
        result = SaveGamePath & "campaign/auto" & "/"
    Case LOAD_FRONTEND_SKIRMISH_AUTO, LOAD_INGAME_SKIRMISH_AUTO
        result = SaveGamePath & "skirmish/auto" & "/"
    Case SAVE_MISSIONEND, SAVE_INGAME_MISSION
        result = SaveGamePath & "campaign" & "/"
    Case SAVE_INGAME_SKIRMISH
        result = SaveGamePath & "skirmish" & "/"
    Case LOADREPLAY_FRONTEND_SKIRMISH
        result = ReplayPath & "skirmish" & "/"
    Case LOADREPLAY_FRONTEND_MULTI
        result = ReplayPath & "multiplay" & "/"
    Case Else
        ' (ASSERT("Invalid load/save mode!", "Invalid load/save mode!"): its expression, a string literal, is never
        ' false, so it never fires)
        result = SaveGamePath & "campaign" & "/"
    End Select

    getNewSaveGamePathFromMode = result
End Function

' ////////////////////////////////////////////////////////////////////////////
Public Function addLoadSave(ByVal savemode As LOADSAVE_MODE, ByVal title_ As String) As Boolean
    Dim NewSaveGamePath As String, slotCount As Long, bLoad As Boolean, bReplay As Boolean, bIngame As Boolean
    Dim parent As WIDGET, loadSaveForm As WIDGET, sFormInit As W_FORMINIT, sLabInit As W_LABINIT, sButInit As W_BUTINIT
    Dim button As WIDGET
    ' (std::vector<SaveGameNamesAndTimes> saveGameNamesAndTimes: its name and epoch in sgName/sgEpoch(0 .. sgCount - 1))
    Dim sgName() As String, sgEpoch() As Double, sgCount As Long
    Dim buildTagVer() As Long, buildQualifier As String, latestTagHasValue As Boolean, errDescription As String
    Dim extension As String, extensionLen As Long, files() As String, nFiles As Long, k As Long, i As String
    Dim savefile_ As String, savetime As Double, lenIWithoutExtension As Long, fileNameWithoutExtension As String
    Dim e As Long, j As Long, tName As String, tEpoch As Double, newtime As tm

    bLoadSaveMode = savemode
    savedTitle = title_

    ' Static as these are assigned to the widget buttons by reference
    ' (sSlotCaps, sSlotTips: the module's)

    bLoad = True
    bReplay = False
    NewSaveGamePath = getNewSaveGamePathFromMode(savemode)
    Select Case savemode
    Case SAVE_MISSIONEND, SAVE_INGAME_MISSION, SAVE_INGAME_SKIRMISH
        bLoad = False
    Case LOADREPLAY_FRONTEND_SKIRMISH, LOADREPLAY_FRONTEND_MULTI
        bReplay = True
    Case Else
    End Select

    bIngame = bLoadSaveMode = LOAD_INGAME_MISSION Or bLoadSaveMode = SAVE_INGAME_MISSION _
        Or bLoadSaveMode = LOAD_INGAME_SKIRMISH Or bLoadSaveMode = SAVE_INGAME_SKIRMISH _
        Or bLoadSaveMode = LOAD_INGAME_MISSION_AUTO Or bLoadSaveMode = LOAD_INGAME_SKIRMISH_AUTO

    modeLoad = bLoad
    modeIngame = bIngame
    modeReplay = bReplay
    debug_ 239, LOG_SAVE, "addLoadSave", "called (%d, %s)", Abs(bLoad), title_

    If bIngame Then
        If Not bMultiPlayer Or Not NetPlay_bComms Then
            gameTimeStop
            If GetGameMode() = GS_NORMAL Then
                ' just display the 3d, no interface
                displayWorld
            End If

            setGamePauseStatus True
            setAllPauseStates True
        End If

        forceHidePowerBar
        intRemoveReticule
        intHideInGameOptionsButton
        intHideGroupSelectionMenu
    End If

    Set psRequestScreen = W_SCREEN_make()

    Set parent = psRequestScreen.psForm

    ' add a form to place the tabbed form on
    ' we need the form to be long enough for all resolutions, so we take the total number of items * height
    ' and * the gaps, add the banner, and finally, the fudge factor ;)
    Set loadSaveForm = IntFormAnimated_new()
    parent.attach loadSaveForm
    loadSaveForm.id = LOADSAVE_FORM
    loadSaveForm.setCalcLayout WzFunc_new(FN_loadSaveForm_calcLayout)

    ' Add Banner
    W_FORMINIT_construct sFormInit
    sFormInit.formID = LOADSAVE_FORM
    sFormInit.id = LOADSAVE_BANNER
    sFormInit.x = LOADSAVE_HGAP
    sFormInit.y = LOADSAVE_VGAP
    sFormInit.width = LOADSAVE_W - (2 * LOADSAVE_HGAP)
    sFormInit.height = LOADSAVE_BANNER_DEPTH
    Set sFormInit.pDisplay = WzFunc_new(FN_loadsave_displayLoadBanner)
    sFormInit.UserData = Abs(bLoad)
    widgAddForm psRequestScreen, sFormInit

    ' Add Banner Label
    W_LABINIT_construct sLabInit
    sLabInit.formID = LOADSAVE_BANNER
    sLabInit.FontID = font_large
    sLabInit.id = LOADSAVE_LABEL
    sLabInit.style = WLAB_ALIGNCENTRE
    sLabInit.x = 0
    sLabInit.y = 0
    sLabInit.width = LOADSAVE_W - (2 * LOADSAVE_HGAP)      ' LOADSAVE_W;
    sLabInit.height = LOADSAVE_BANNER_DEPTH                 ' This looks right -Q
    sLabInit.pText = title_
    widgAddLabel psRequestScreen, sLabInit

    ' add cancel.
    W_BUTINIT_construct sButInit
    sButInit.formID = LOADSAVE_BANNER
    sButInit.x = 8
    sButInit.y = 10
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_NRUTER)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_NRUTER)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_NRUTER, IMAGE_NRUTER)

    sButInit.id = LOADSAVE_CANCEL
    sButInit.style = WBUT_PLAIN
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    widgAddButton psRequestScreen, sButInit

    ' add slots
    W_BUTINIT_construct sButInit                ' sButInit = W_BUTINIT();
    sButInit.formID = LOADSAVE_FORM
    sButInit.style = WBUT_PLAIN
    sButInit.width = LOADENTRY_W
    sButInit.height = LOADENTRY_H
    Set sButInit.pDisplay = WzFunc_new(FN_loadsave_displayLoadSlot)
    Set sButInit.initPUserDataFunc = WzFunc_new(FN_newLoadSaveDisplayLoadSlotCache)
    Set sButInit.onDelete = WzFunc_new(FN_deleteLoadSaveDisplayLoadSlotCache)

    For slotCount = 0 To totalslots - 1
        sButInit.id = slotCount + LOADENTRY_START

        If slotCount < slotsInColumn Then
            sButInit.x = 22 + LOADSAVE_HGAP
            sButInit.y = ToInt16((LOADSAVE_BANNER_DEPTH + (2 * LOADSAVE_VGAP)) + ( _
                                  slotCount * (LOADSAVE_VGAP + LOADENTRY_H)))
        ElseIf slotCount >= slotsInColumn And (slotCount < (slotsInColumn * 2)) Then
            sButInit.x = 22 + (2 * LOADSAVE_HGAP + LOADENTRY_W)
            sButInit.y = ToInt16((LOADSAVE_BANNER_DEPTH + (2 * LOADSAVE_VGAP)) + ( _
                                  (slotCount Mod slotsInColumn) * (LOADSAVE_VGAP + LOADENTRY_H)))
        Else
            sButInit.x = 22 + (3 * LOADSAVE_HGAP + (2 * LOADENTRY_W))
            sButInit.y = ToInt16((LOADSAVE_BANNER_DEPTH + (2 * LOADSAVE_VGAP)) + ( _
                                  (slotCount Mod slotsInColumn) * (LOADSAVE_VGAP + LOADENTRY_H)))
        End If
        widgAddButton psRequestScreen, sButInit
    Next

    ' fill slots.
    slotCount = 0

    ' The first slot is for [auto] or [..]
    Set button = widgGetFromID(psRequestScreen, LOADENTRY_START + slotCount)
    If bLoadSaveMode >= LOAD_FRONTEND_MISSION_AUTO Then
        W_BUTTON_(button).pTip = "Parent directory"
        W_BUTTON_(button).pText = "[..]"
    Else
        If modeLoad Then
            W_BUTTON_(button).pTip = "Autosave directory"
        Else
            W_BUTTON_(button).pTip = "Autosave directory (not allowed for saving)"
        End If
        W_BUTTON_(button).pText = "[auto]"
    End If
    slotCount = slotCount + 1

    debug_ 376, LOG_SAVE, "addLoadSave", "Searching ""%s"" for savegames", NewSaveGamePath

    ' add savegame filenames minus extensions to buttons

    sgCount = 0
    latestTagHasValue = version_extractVersionNumberFromTag(version_getLatestTag(), buildTagVer, buildQualifier)
    If Not (latestTagHasValue) Then ASSERT_FAILURE "latestTagResult.has_value()", "src/loadsave.cpp:393", "addLoadSave", "No extractable latest tag?? - Please try re-downloading the latest official source bundle"
    If Not latestTagHasValue Then ReDim buildTagVer(0 To 2)    ' const TagVer buildTagVer = latestTagResult.value_or(TagVer());
    If Not addLoadSave_enumerateFolders_try(NewSaveGamePath, buildTagVer, sgName, sgEpoch, sgCount, errDescription) Then
        ' catch( nlohmann::json::exception &e)
        debug_ 429, LOG_ERROR, "addLoadSave", "can't load game: %s", errDescription
        ' continue, because still may find old .gam to load
    End If

    If bReplay Then
        extension = sSaveReplayExtension
        extensionLen = Len(extension)
        ' Note: this is left for backward compatibility reasons.
        ' Previously added old .gam saves. Now it exists to look for replay files.
        If WZ_PHYSFS_enumerateFiles(NewSaveGamePath, files, nFiles) Then
            For k = 0 To nFiles - 1
                i = files(k)
                ' (char savefile[256]; time_t savetime;)

                If Not isASavedGamefile(i, extension) Then
                    ' If it doesn't, move on to the next filename
                    GoTo continue_replay        ' return true;
                End If

                debug_ 449, LOG_SAVE, "addLoadSave", "We found [%s]", i

                ' Figure save-time
                savefile_ = Left$(NewSaveGamePath & "/" & i, 255)  ' snprintf(savefile, sizeof(savefile), "%s/%s", ...)
                savetime = WZ_PHYSFS_getLastModTime(savefile_)

                lenIWithoutExtension = Len(i) - extensionLen
                fileNameWithoutExtension = Left$(i, lenIWithoutExtension)
                For e = 0 To sgCount - 1
                    ' only add if doesn't exist yet
                    If StrComp(sgName(e), fileNameWithoutExtension, vbBinaryCompare) = 0 Then
                        GoTo continue_replay    ' return true; // move to next
                    End If
                Next
                ReDim Preserve sgName(0 To sgCount)
                ReDim Preserve sgEpoch(0 To sgCount)
                sgName(sgCount) = fileNameWithoutExtension
                sgEpoch(sgCount) = savetime
                sgCount = sgCount + 1
continue_replay:
            Next
        End If
    End If

    ' Sort the save games so that the most recent one appears first
    ' (std::sort with [](a, b) { return a.epoch > b.epoch; }: an insertion sort)
    For k = 1 To sgCount - 1
        tName = sgName(k)
        tEpoch = sgEpoch(k)
        j = k - 1
        Do While j >= 0
            If Not (tEpoch > sgEpoch(j)) Then Exit Do
            sgName(j + 1) = sgName(j)
            sgEpoch(j + 1) = sgEpoch(j)
            j = j - 1
        Loop
        sgName(j + 1) = tName
        sgEpoch(j + 1) = tEpoch
    Next

    ' Now store the sorted save game names to the buttons
    slotCount = 1
    ' ((void)std::all_of(saveGameNamesAndTimes..., [&](SaveGameNamesAndTimes& saveGameNameAndTime) {...}))
    For k = 0 To sgCount - 1
        ' Set the button-text and tip text (the save time) into static storage
        sSlotCaps(slotCount) = Left$(sgName(k), totalslotspace - 1)    ' sstrcpy
        newtime = getLocalTime(sgEpoch(k))
        sSlotTips(slotCount) = Left$(strftime("%Y-%m-%d %H:%M:%S", newtime), totalslotspace - 1)  ' "%F %H:%M:%S"

        ' Add a button that references the static strings
        Set button = widgGetFromID(psRequestScreen, LOADENTRY_START + slotCount)
        W_BUTTON_(button).pTip = sSlotTips(slotCount)
        W_BUTTON_(button).pText = sSlotCaps(slotCount)
        slotCount = slotCount + 1

        If Not (slotCount < totalslots) Then Exit For
    Next

    bLoadSaveUp = True
    addLoadSave = True
End Function

' (VB) addLoadSave's try block: WZ_PHYSFS_enumerateFolders(NewSaveGamePath, [NewSaveGamePath, &buildTagVer,
' &saveGameNamesAndTimes](const char* dirName) {...}): the lambda over the folders, adding to
' sgName/sgEpoch(0 .. sgCount - 1); False with the exception's what() in errDescription when it throws
Private Function addLoadSave_enumerateFolders_try(ByVal NewSaveGamePath As String, buildTagVer() As Long, sgName() As String, sgEpoch() As Double, sgCount As Long, errDescription As String) As Boolean
    Dim folders() As String, nFolders As Long, k As Long, dirName As String, saveInfoFilename As String
    Dim saveInfoData As nlohmann_json, saveLatestTagArray As nlohmann_json, saveTagVer(0 To 2) As Long, v As Long
    Dim epoch As Double
    On Error GoTo threw
    If WZ_PHYSFS_enumerateFolders(NewSaveGamePath, folders, nFolders) Then
        For k = 0 To nFolders - 1
            dirName = folders(k)
            If dirName = "auto" Then
                GoTo continue_                  ' return true; // continue
            End If
            saveInfoFilename = NewSaveGamePath & dirName & "/save-info.json"

            ' avoid spamming stdout if doesn't exist
            If PHYSFS_exists(saveInfoFilename) = 0 Then GoTo continue_

            If parseJsonFile(saveInfoFilename, saveInfoData) Then
                ' decide what savegames are viewable/loadable
                ' assume that we can safely load games older than current version
                Set saveLatestTagArray = saveInfoData.at("latestTagArray")
                ' (get<std::vector<uint16_t>>(), then TagVer saveTagVer(saveLatestTagArray, ""): its first three)
                For v = 0 To 2
                    saveTagVer(v) = saveLatestTagArray.at_idx(v).get_uint() And &HFFFF&
                Next
                If TagVer_lessOrEqual(saveTagVer, buildTagVer) Then
                    ' looks like loadable
                    epoch = saveInfoData.at("epoch").get_int64()
                    ReDim Preserve sgName(0 To sgCount)
                    ReDim Preserve sgEpoch(0 To sgCount)
                    sgName(sgCount) = dirName
                    sgEpoch(sgCount) = epoch
                    sgCount = sgCount + 1
                Else
                    debug_ 422, LOG_SAVEGAME, "addLoadSave", "not showing savegame '%s', because version is higher than this game", saveInfoFilename
                End If
            End If
continue_:
        Next
    End If
    addLoadSave_enumerateFolders_try = True
    Exit Function
threw:
    errDescription = Err.Description
    addLoadSave_enumerateFolders_try = False
End Function

' ////////////////////////////////////////////////////////////////////////////
Public Function closeLoadSave(Optional ByVal goBack As Boolean = False) As Boolean
    bLoadSaveUp = False

    If modeIngame Then
        If goBack Then
            intReopenMenuWithoutUnPausing
        End If
        If Not bMultiPlayer Or Not NetPlay_bComms Then
            setGameUpdatePause False
            If Not goBack Then
                gameTimeStart
                setGamePauseStatus False
                setAllPauseStates False
            End If
        End If

        intAddReticule
        intShowPowerBar
        intShowGroupSelectionMenu
    End If

    Set psRequestScreen = Nothing
    ' need to "eat" up the return key so it don't pass back to game.
    inputLoseFocus
    closeLoadSave = True
End Function

Public Function closeLoadSaveOnShutdown() As Boolean
    ' Workaround for ensuring psRequestScreen is freed on shutdown - ideally this whole screen could probably be implemented as an overlay screen now and simplified
    bLoadSaveUp = False
    Set psRequestScreen = Nothing
    closeLoadSaveOnShutdown = True
End Function

'**************************************************************************
'	Delete a savegame.  fileName should be a .gam extension save game
'	filename reference.  We delete this file, any .es file with the same
'	name, and any files in the directory with the same name.
'**************************************************************************
' TODO: Replace the old uses of .gam files with the savegame folder and call deleteSaveGame directly?
Public Sub deleteSaveGame_classic(ByVal gamFileName As String)
    Dim gamFolderPath As String
    If Not (StrPtr(gamFileName) <> 0) Then ASSERT_FAILURE "gamFileName != nullptr", "src/loadsave.cpp:547", "deleteSaveGame_classic", "Null gamFileName": Exit Sub
    gamFolderPath = gamFileName

    ' Old method of stripping file extension - makes lots of assumptions (like there *is* a file extension and it has a length of 4 including the period)
    If Not (Len(gamFolderPath) > 4) Then ASSERT_FAILURE "gamFolderPath.size() > 4", "src/loadsave.cpp:551", "deleteSaveGame_classic", "Input gamFileName too short?: %s", gamFileName: Exit Sub
    gamFolderPath = Left$(gamFolderPath, Len(gamFolderPath) - 4)

    deleteSaveGame gamFolderPath
End Sub

Public Sub deleteSaveGame(ByVal saveGameFolderPath As String)
    Dim oldParentFile As String, files() As String, nFiles As Long, k As Long, del_file As String

    ' Remove any trailing path separators (/)
    Do While Len(saveGameFolderPath) > 0
        If InStrRev(saveGameFolderPath, "/") <> Len(saveGameFolderPath) Then Exit Do
        saveGameFolderPath = Left$(saveGameFolderPath, Len(saveGameFolderPath) - 1) ' Remove trailing path separators
    Loop

    ' First, check if some old files in the parent directory (named the same as this directory, but with specific extensions) exist
    ' (since we removed trailing directory separators above, we can just append the appropriate file extensions)
    oldParentFile = saveGameFolderPath & ".gam"
    If PHYSFS_exists(oldParentFile) <> 0 Then
        PHYSFS_delete oldParentFile
    End If
    oldParentFile = saveGameFolderPath & ".es"  ' ancient script data file
    If PHYSFS_exists(oldParentFile) <> 0 Then
        PHYSFS_delete oldParentFile
    End If

    ' check for a directory and remove that too.
    If WZ_PHYSFS_enumerateFiles(saveGameFolderPath, files, nFiles) Then
        For k = 0 To nFiles - 1
            ' Construct the full path to the file by appending the
            ' filename to the directory it is in.
            del_file = saveGameFolderPath & "/" & files(k)

            debug_ 584, LOG_SAVE, "deleteSaveGame", "Deleting [%s].", del_file

            ' Delete the file
            If PHYSFS_delete(del_file) = 0 Then
                debug_ 589, LOG_ERROR, "deleteSaveGame", "Warning [%s] could not be deleted due to PhysicsFS error: %s", del_file, WZ_PHYSFS_getLastError()
            End If
        Next                                    ' return true; // continue
    End If

    If WZ_PHYSFS_isDirectory(saveGameFolderPath) <> 0 Then
        If PHYSFS_delete(saveGameFolderPath) = 0 Then   ' now (should be)empty directory
            debug_ 596, LOG_ERROR, "deleteSaveGame", "Warning directory[%s] could not be deleted because %s", saveGameFolderPath, WZ_PHYSFS_getLastError()
        End If
    End If
End Sub

Private Function findLastSaveFrom(ByVal loc_ As SAVEGAME_LOC) As Boolean
    Dim found_ As Boolean, path As String, pathToCommonSaveDir As String, errDescription As String
    found_ = False
    path = SaveGameLocToPath(loc_)
    debug_ 608, LOG_SAVEGAME, "findLastSaveFrom", "looking for last save in %s", path
    pathToCommonSaveDir = path
    If Not findLastSaveFrom_try(pathToCommonSaveDir, loc_, found_, errDescription) Then
        ' catch (nlohmann::json::exception &e)
        debug_ 642, LOG_ERROR, "findLastSaveFrom", "find last save failed:%s", errDescription
        findLastSaveFrom = False
        Exit Function
    End If

    findLastSaveFrom = found_
End Function

' (VB) findLastSaveFrom's try block: WZ_PHYSFS_enumerateFolders(pathToCommonSaveDir, [&found, &pathToCommonSaveDir,
' loc] (const char * dirName) {...}); False with the exception's what() in errDescription when it throws
Private Function findLastSaveFrom_try(ByVal pathToCommonSaveDir As String, ByVal loc_ As SAVEGAME_LOC, found_ As Boolean, errDescription As String) As Boolean
    Dim folders() As String, nFolders As Long, k As Long, dirName As String, pathToSaveInfo As String
    Dim saveInfo As nlohmann_json, epoch As Double
    On Error GoTo threw
    If WZ_PHYSFS_enumerateFolders(pathToCommonSaveDir, folders, nFolders) Then
        For k = 0 To nFolders - 1
            dirName = folders(k)
            If dirName = "auto" Then
                ' skip "auto" folder, it will be iterated on its own
                GoTo continue_
            End If
            pathToSaveInfo = pathToCommonSaveDir & "/" & dirName & "/save-info.json"
            If PHYSFS_exists(pathToSaveInfo) <> 0 Then
                If Not parseJsonFile(pathToSaveInfo, saveInfo) Then
                    debug_ 624, LOG_SAVEGAME, "findLastSaveFrom", "weird directory without save-info.json: %s", dirName
                    GoTo continue_
                End If
                epoch = saveInfo.at("epoch").get_int64()
                If epoch > lastSaveTime Then
                    lastSaveTime = epoch
                    found_ = True
                    lastSavePath.loc = loc_
                    lastSavePath.gameName = dirName
                    debug_ 635, LOG_SAVEGAME, "findLastSaveFrom", "found last saved game: %s%s", pathToCommonSaveDir, lastSavePath.gameName
                End If
            End If
continue_:
        Next                                    ' return true;
    End If
    findLastSaveFrom_try = True
    Exit Function
threw:
    errDescription = Err.Description
    findLastSaveFrom_try = False
End Function

Public Function findLastSave() As Boolean
    Dim foundMP As Boolean, foundCAM As Boolean
    lastSaveTime = 0
    foundCAM = findLastSaveFrom(SAVEGAME_LOC_CAM)
    foundCAM = foundCAM Or findLastSaveFrom(SAVEGAME_LOC_CAM_AUTO)
    foundMP = findLastSaveFrom(SAVEGAME_LOC_SKI)
    foundMP = foundMP Or findLastSaveFrom(SAVEGAME_LOC_SKI_AUTO)
    lastSaveMP = foundMP
    findLastSave = foundMP Or foundCAM
End Function

Private Function suggestSaveName(ByVal saveGamePath_ As String) As String
    Dim levelName As String, levelNameStr As String, cheatedSuffix As String, saveNamePartial As String
    Dim campaignName As String, missionName As String, endName As String, levelNames() As String, nLevelNames As Long
    Dim i As Long, fullMapName As String, endsWithS As Boolean, saveName As String, similarSaveGames As Long
    Dim suffixArr() As Long, nSuffix As Long, folders() As String, nFolders As Long, k As Long, j As Long, t As Long
    Dim dirNameStr As String, saveNameStr As String, pos As Long, restOfSaveName As String, lastSpace As Long
    Dim converted As Long, curSaveNum As Long, foundStart As Boolean, it As Long

    levelName = getLevelName()
    levelNameStr = levelName
    If Cheated Then cheatedSuffix = "cheated" Else cheatedSuffix = ""
    saveNamePartial = ""                        ' char saveNamePartial[64] = "\0";

    If Not bMultiPlayer Then
        missionName = levelName
        endName = "End"                         ' last mission for one the campaigns
        nLevelNames = 0                         ' std::vector<std::string> levelNames = {};

        If Left$(levelName, 5) = "CAM_1" Or Left$(levelName, 5) = "SUB_1" Then
            campaignName = "Alpha"
            levelNames = Split("CAM_1A,CAM_1B,SUB_1_1,SUB_1_2,SUB_1_3,CAM_1C," & _
                               "CAM_1CA,SUB_1_4A,SUB_1_5,CAM_1A-C,SUB_1_7,SUB_1_D,CAM_1END", ",")
            nLevelNames = UBound(levelNames) + 1
        ElseIf Left$(levelName, 5) = "CAM_2" Or Left$(levelName, 5) = "SUB_2" Then
            campaignName = "Beta"
            levelNames = Split("CAM_2A,SUB_2_1,CAM_2B,SUB_2_2,CAM_2C,SUB_2_5," & _
                               "SUB_2D,SUB_2_6,SUB_2_7,SUB_2_8,CAM_2END", ",")
            nLevelNames = UBound(levelNames) + 1
        ElseIf Left$(levelName, 5) = "CAM_3" Or Left$(levelName, 4) = "CAM3" Or Left$(levelName, 5) = "SUB_3" Then
            campaignName = "Gamma"
            levelNames = Split("CAM_3A,SUB_3_1,CAM_3B,SUB_3_2,CAM3A-B," & _
                               "CAM3C,CAM3A-D1,CAM3A-D2,CAM_3_4", ",")
            nLevelNames = UBound(levelNames) + 1
        End If

        For i = 0 To nLevelNames - 1
            fullMapName = levelNames(i)
            If InStr(1, levelNameStr, fullMapName, vbBinaryCompare) <> 0 Then
                endsWithS = Right$(levelNameStr, 1) = "S"
                If Not endsWithS And (StrComp(levelNameStr, fullMapName, vbBinaryCompare) <> 0) Then
                    GoTo continue_level
                End If
                missionName = CStr(i + 1)
                If (i + 1) = nLevelNames Then
                    missionName = endName
                End If
                If (StrComp(missionName, endName, vbBinaryCompare) = 0) And Not (Left$(levelName, 5) = "CAM_3") Then ' Gamma end is an exception
                    Exit For
                End If
                If endsWithS Then
                    missionName = missionName & " Base" ' The Project's base for the campaign map.
                End If
                Exit For
            End If
continue_level:
        Next

        saveNamePartial = Left$(campaignName & " " & missionName & " " & cheatedSuffix, 63)    ' ssprintf(saveNamePartial, "%s %s %s", ...)
    Else
        saveNamePartial = Left$(mapNameWithoutTechlevel(levelNameStr) & " " & cheatedSuffix, 63)    ' ssprintf(saveNamePartial, "%s %s", ...)
    End If

    saveName = WzString_trimmed(saveNamePartial)
    similarSaveGames = 0
    nSuffix = 0                                 ' std::vector<int> suffixArr;
    If WZ_PHYSFS_enumerateFolders(saveGamePath_, folders, nFolders) Then
        For k = 0 To nFolders - 1
            dirNameStr = folders(k)
            saveNameStr = saveName
            pos = InStr(1, dirNameStr, saveNameStr, vbBinaryCompare)

            If pos = 0 Then
                GoTo continue_folder            ' return true; //An original save name
            End If

            restOfSaveName = Mid$(dirNameStr, pos)
            If StrComp(restOfSaveName, saveNameStr, vbBinaryCompare) = 0 Then
                similarSaveGames = similarSaveGames + 1 ' direct match
                GoTo continue_folder            ' return true;
            End If

            ' chop off at the last space+number, if it exists, and compare again
            lastSpace = InStrRev(restOfSaveName, " ")
            If lastSpace <> 0 Then
                ' Get the suffix number
                converted = atoi(Mid$(restOfSaveName, lastSpace))
                If converted <> 0 Then
                    ' Suffix is a number
                    ReDim Preserve suffixArr(0 To nSuffix)
                    suffixArr(nSuffix) = converted
                    nSuffix = nSuffix + 1
                End If
            End If
continue_folder:
        Next                                    ' return true;
    End If
    If similarSaveGames > 0 Then
        ' Sorting the suffix numbers and determining missing saves
        curSaveNum = 2
        ' (std::sort(suffixArr.begin(), suffixArr.end()): an insertion sort)
        For k = 1 To nSuffix - 1
            t = suffixArr(k)
            j = k - 1
            Do While j >= 0
                If Not (t < suffixArr(j)) Then Exit Do
                suffixArr(j + 1) = suffixArr(j)
                j = j - 1
            Loop
            suffixArr(j + 1) = t
        Next
        foundStart = False
        For it = 0 To nSuffix - 1
            If suffixArr(it) = 2 Then foundStart = True
            If Not foundStart And suffixArr(it) <> 2 Then GoTo continue_suffix
            If suffixArr(it) = curSaveNum Then curSaveNum = curSaveNum + 1
continue_suffix:
        Next
        saveName = saveName & " " & CStr(curSaveNum)
    End If

    suggestSaveName = saveName
End Function

Private Sub runLoadSaveCleanup(ByVal resetWidgets As Boolean, ByVal goBack As Boolean)
    closeLoadSave goBack
    bRequestLoad = False
    bRequestLoadReplay = False
    If Not goBack And resetWidgets Then
        If widgGetFromID(psWScreen, IDMISSIONRES_FORM) Is Nothing Then
            resetMissionWidgets
        End If
    End If
End Sub

Private Sub runLoadCleanup()
    bRequestLoadReplay = modeReplay
    bRequestLoad = Not modeReplay
    If modeReplay Then
        ' Load a replay.
        closeLoadSave
        Exit Sub
    End If

    closeLoadSave
End Sub

' ////////////////////////////////////////////////////////////////////////////
' Returns true if cancel pressed or a valid game slot was selected.
' if when returning true strlen(sRequestResult) != 0 then a valid game slot was selected
' otherwise cancel was selected..
Public Function runLoadSave(ByVal bResetMissionWidgets As Boolean) As Boolean
    Dim NewSaveGamePath As String, triggers() As WidgetTrigger, nTriggers As Long, id As Long, cancel As Boolean
    Dim iLoadSaveMode As Long, slotButton As WIDGET, parent As WIDGET, saveEntryEdit As WIDGET, r As WzRect
    Dim suggestedSaveName As String, context As W_CONTEXT, sTemp As String, i As Long, ext As String

    nTriggers = widgRunScreen(psRequestScreen, triggers)
    If nTriggers = 0 Then id = 0 Else id = triggers(0).widget.id   ' Just use first click here, since the next click could be on another menu.

    sRequestResult = ""                         ' set returned filename to null;

    If id = LOADSAVE_CANCEL Then cancel = True Else cancel = CancelPressed()   ' (||: CancelPressed only when not cancel)
    If cancel Then
        runLoadSaveCleanup bResetMissionWidgets, True
        runLoadSave = True
        Exit Function
    End If
    NewSaveGamePath = getNewSaveGamePathFromMode(bLoadSaveMode)
    If id = LOADENTRY_START And modeLoad And Not modeReplay Then  ' [auto] or [..], ignore click for saves
        iLoadSaveMode = bLoadSaveMode           ' for evil integer arithmetics
        bLoadSaveMode = iLoadSaveMode Xor 16    ' toggle _AUTO bit
        closeLoadSave                           ' decrement game time / pause counters
        addLoadSave bLoadSaveMode, savedTitle   ' restart in other directory
        runLoadSave = False
        Exit Function
    ' clicked a load entry
    ElseIf id > LOADENTRY_START And id <= LOADENTRY_END Then
        Set slotButton = widgGetFromID(psRequestScreen, id)

        If modeLoad Then                        ' Loading, return that entry.
            If Len(W_BUTTON_(slotButton).pText) = 0 Then
                runLoadSave = False             ' clicked on an empty box
                Exit Function
            End If
            If modeReplay Then ext = sSaveReplayExtension Else ext = sSaveGameExtension
            sRequestResult = NewSaveGamePath & W_BUTTON_(widgGetFromID(psRequestScreen, id)).pText & ext

            runLoadCleanup
            runLoadSave = True
            Exit Function
        Else                                    '  SAVING!add edit box at that position.

            If widgGetFromID(psRequestScreen, SAVEENTRY_EDIT) Is Nothing Then
                Set parent = widgGetFromID(psRequestScreen, LOADSAVE_FORM)

                ' add blank box.
                Set saveEntryEdit = W_EDITBOX_new()
                parent.attach saveEntryEdit
                saveEntryEdit.id = SAVEENTRY_EDIT
                r = slotButton.geometry()
                saveEntryEdit.setGeometryRect r
                saveEntryEdit.setString slotButton.getString()
                W_EDITBOX_setBoxColours saveEntryEdit, psPalette(WZCOL_MENU_LOAD_BORDER), psPalette(WZCOL_MENU_LOAD_BORDER), psPalette(WZCOL_MENU_BACKGROUND)

                If Len(W_BUTTON_(slotButton).pText) <> 0 Then
                    sDelete = NewSaveGamePath & W_BUTTON_(slotButton).pText & sSaveGameExtension
                Else
                    suggestedSaveName = suggestSaveName(NewSaveGamePath)
                    saveEntryEdit.setString suggestedSaveName
                    sDelete = ""
                End If

                slotButton.hide                 ' hide the old button
                chosenSlotId = id

                ' auto click in the edit box we just made.
                ' (W_CONTEXT context = W_CONTEXT::ZeroContext(): a new W_CONTEXT is zeroed)
                context.mx = mouseX()
                context.my = mouseY()
                saveEntryEdit.clicked context
            End If
        End If
    End If

    ' finished entering a name.
    If id = SAVEENTRY_EDIT Then
        ' (char sTemp[MAX_STR_LENGTH])

        If Not keyPressed(KEY_RETURN) And Not keyPressed(KEY_KPENTER) Then  ' enter was not pushed, so not a vaild entry.
            widgDelete psRequestScreen, SAVEENTRY_EDIT  ' unselect this box, and go back ..
            widgReveal psRequestScreen, chosenSlotId
            runLoadSave = True
            Exit Function
        End If


        ' scan to see if that game exists in another slot, if so then fail.
        sTemp = Left$(widgGetString(psRequestScreen, id), MAX_STR_LENGTH - 1)

        For i = LOADENTRY_START To LOADENTRY_END - 1
            If i <> chosenSlotId Then

                If Len(W_BUTTON_(widgGetFromID(psRequestScreen, i)).pText) <> 0 Then
                    If StrComp(sTemp, W_BUTTON_(widgGetFromID(psRequestScreen, i)).pText, vbBinaryCompare) = 0 Then
                        widgDelete psRequestScreen, SAVEENTRY_EDIT  ' unselect this box, and go back ..
                        widgReveal psRequestScreen, chosenSlotId
                        audio_PlayBuildFailedOnce
                        runLoadSave = True
                        Exit Function
                    End If
                End If
            End If
        Next


        ' return with this name, as we've edited it.
        If Len(widgGetString(psRequestScreen, id)) <> 0 Then
            sTemp = Left$(widgGetString(psRequestScreen, id), MAX_STR_LENGTH - 1)
            removeWildcards sTemp
            sRequestResult = NewSaveGamePath & sTemp & sSaveGameExtension
            If Len(sDelete) <> 0 Then
                deleteSaveGame_classic sDelete  ' only delete game if a new game fills the slot
            End If
        End If

        runLoadSaveCleanup bResetMissionWidgets, False
        runLoadSave = True
        Exit Function
    End If

    runLoadSave = False
End Function

' ////////////////////////////////////////////////////////////////////////////
' should be done when drawing the other widgets.
Public Function displayLoadSave() As Boolean
    widgDisplayScreen psRequestScreen   ' display widgets.
    displayLoadSave = True
End Function


' ////////////////////////////////////////////////////////////////////////////
' char HANDLER, replaces dos wildcards in a string with harmless chars.
' (pStr: the char[] changed in place; a String's characters >= 0x80 stand for the UTF-8 bytes >= 0x80, and its
' length for strlen)
Public Sub removeWildcards(pStr As String)
    Dim i As Long, c As Long, prev As Long

    ' Remember never to allow: < > : " / \ | ? *

    ' Whitelist: Get rid of any characters except:
    ' a-z A-Z 0-9 - + ! , = ^ @ # $ % & ' ( ) [ ] (and space and unicode characters >= 0x80)
    For i = 0 To Len(pStr) - 1
        c = AscW(Mid$(pStr, i + 1, 1)) And &HFFFF&
        If i > 0 Then prev = AscW(Mid$(pStr, i, 1)) And &HFFFF&
        If Not ((c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)) _
           And (c <> 32 Or i = 0 Or prev = 32) _
           And c <> 45 _
           And c <> 43 _
           And c <> 91 And c <> 93 _
           And c < &H80& Then
            ' (!isalnum; we allow spaces as long as they aren't the first char, or two spaces in a row; '-' '+'
            ' '[' ']'; (pStr[i] & 0x80) != 0x80: the accented and Greek letters, a e i o u alpha beta ...)
            Mid$(pStr, i + 1, 1) = "_"
        End If
    Next

    If Len(pStr) >= MAX_SAVE_NAME Then
        pStr = Left$(pStr, MAX_SAVE_NAME - 1)
    ElseIf Len(pStr) = 0 Then
        pStr = "!"
        Exit Sub
    End If
    ' Trim trailing spaces
    i = Len(pStr)
    Do While i > 0
        If Mid$(pStr, i, 1) <> " " Then Exit Do
        i = i - 1
    Loop
    pStr = Left$(pStr, i)

    ' If that leaves us with a blank string, replace with '!'
    If Len(pStr) = 0 Then
        pStr = "!"
    End If
End Sub

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' DISPLAY FUNCTIONS

Private Sub displayLoadBanner(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim col As Long, x As Long, y As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    If Not IsEmpty(psWidget.pUserData) Then
        col = psPalette(WZCOL_GREEN)
    Else
        col = psPalette(WZCOL_MENU_LOAD_BORDER)
    End If

    pie_BoxFill x, y, x + psWidget.width(), y + psWidget.height(), col
    pie_BoxFill x + 2, y + 2, x + psWidget.width() - 2, y + psWidget.height() - 2, psPalette(WZCOL_MENU_BACKGROUND)
End Sub

' ////////////////////////////////////////////////////////////////////////////
Private Sub displayLoadSlot(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim cache As LoadSaveDisplayLoadSlotCache, x As Long, y As Long, butString As String
    ' Any widget using displayLoadSlot must have its pUserData initialized to a (LoadSaveDisplayLoadSlotCache*)
    ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
    Set cache = psWidget.pUserData

    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    ' (char butString[64])

    drawBlueBox x, y, psWidget.width(), psWidget.height()   ' draw box

    If Len(W_BUTTON_(psWidget).pText) <> 0 Then
        butString = Left$(W_BUTTON_(psWidget).pText, 63)    ' sstrcpy(butString, ...)

        If cache.fullText <> butString Then
            ' Update cache
            Do While iV_GetTextWidth(butString, font_regular) > psWidget.width()
                butString = Left$(butString, Len(butString) - 1)
            Loop
            cache.wzText_.setText butString, font_regular
            cache.fullText = butString
        End If

        cache.wzText_.render x + 4, y + 17, psPalette(WZCOL_FORM_TEXT)
    End If
End Sub

Public Sub drawBlueBoxInset(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    pie_BoxFill x, y, x + w, y + h, psPalette(WZCOL_MENU_BORDER)
    pie_BoxFill x + 1, y + 1, x + w - 1, y + h - 1, psPalette(WZCOL_MENU_BACKGROUND)
End Sub

'*
' * Same as drawBlueBoxInset, but the rectangle is overflown by one pixel.
' *
Public Sub drawBlueBox(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    drawBlueBoxInset x - 1, y - 1, w + 2, h + 2
End Sub

' Note: remove later at some point
' returns true if something was deleted
' (WZ_PHYSFS_cleanupOldFilesInFolder(path, sSaveGameExtension, -1, [](const char *fileName){
' deleteSaveGame_classic(fileName); return true; }) > 0: physfs_ext_c's, with its deleting loop here)
Private Function freeAutoSaveSlot_old(ByVal path As String) As Boolean
    Const fileLimit As Long = -1
    Dim result As Long, nfiles As Long, oldest() As String, count As Long, k As Long, numFilesDeleted As Long
    result = WZ_PHYSFS_cleanupOldFilesInFolder(path, sSaveGameExtension, fileLimit, nfiles, oldest, count)
    If result = 1 Then
        ' now delete the oldest
        numFilesDeleted = 0
        k = 0
        Do While nfiles > fileLimit And k < count
            deleteSaveGame_classic oldest(k)    ' (deleteFileFunction: true)
            numFilesDeleted = numFilesDeleted + 1
            nfiles = nfiles - 1
            k = k + 1                           ' fileTimes.erase(fileTimes.begin());
            If fileLimit < 0 Then
                Exit Do
            End If
        Loop
        result = numFilesDeleted
    End If
    freeAutoSaveSlot_old = result > 0
End Function

Private Sub freeAutoSaveSlot(ByVal loc_ As SAVEGAME_LOC)
    Dim path As String, oldestEpoch As Double, oldestKey As String, count As Long, errDescription As String
    Dim savefile_ As String
    path = SaveGameLocToPath(loc_)
    oldestEpoch = INT64_MAX_D
    oldestKey = ""
    count = 0
    If Not freeAutoSaveSlot_try(path, oldestKey, oldestEpoch, count, errDescription) Then
        ' catch (nlohmann::json::exception &e)
        debug_ 1119, LOG_ERROR, "freeAutoSaveSlot", "failed to remove an autosave game: %s", errDescription
        Exit Sub
    End If

    If count < totalslots Then
        Exit Sub
    End If
    If Not (Len(oldestKey) > 0) Then ASSERT_FAILURE "oldestKey.size() > 0", "src/loadsave.cpp:1127", "freeAutoSaveSlot", "Bug: oldestKey can't be empty here": Exit Sub
    ' (char savefile[PATH_MAX])
    savefile_ = path & "/" & oldestKey & ".gam"
    debug_ 1130, LOG_SAVEGAME, "freeAutoSaveSlot", "deleting the oldest autosave file %s", savefile_
    deleteSaveGame_classic savefile_
End Sub

' (VB) freeAutoSaveSlot's try block: WZ_PHYSFS_enumerateFolders(path, [path, &oldestKey, &oldestEpoch, &count](const
' char* dirName) {...}); False with the exception's what() in errDescription when it throws
Private Function freeAutoSaveSlot_try(ByVal path As String, oldestKey As String, oldestEpoch As Double, count As Long, errDescription As String) As Boolean
    Dim folders() As String, nFolders As Long, k As Long, dirName As String, saveInfoFilename As String
    Dim saveInfoDataOpt As Boolean, saveInfoData As nlohmann_json, epoch As Double
    On Error GoTo threw
    If WZ_PHYSFS_enumerateFolders(path, folders, nFolders) Then
        For k = 0 To nFolders - 1
            dirName = folders(k)
            ' (if (!dirName) { return true; }: a name is never a nullptr here)
            If dirName = "auto" Then
                GoTo continue_                  ' return true; // continue
            End If
            count = count + 1
            saveInfoFilename = path & "/" & dirName & "/save-info.json"
            saveInfoDataOpt = parseJsonFile(saveInfoFilename, saveInfoData)
            If Not saveInfoDataOpt Or PHYSFS_exists(saveInfoFilename) = 0 Then
                ' nothing to do, this has been handled by old routine
                GoTo continue_
            End If
            If Not saveInfoData.is_object() Then
                ' just overwrite this (apparently) corrupt autosave
                oldestEpoch = 0
                oldestKey = dirName
                GoTo continue_
            End If
            epoch = saveInfoData.at("epoch").get_int64()
            If epoch < oldestEpoch Then
                oldestEpoch = epoch
                oldestKey = dirName
            End If
continue_:
        Next                                    ' return true;
    End If
    freeAutoSaveSlot_try = True
    Exit Function
threw:
    errDescription = Err.Description
    freeAutoSaveSlot_try = False
End Function

Private Function doAutoSave() As Boolean
    Dim dir_ As String, now_ As Double, timeinfo As tm, savedate As String, suggestedName As String, savefile_ As String
    If bMultiPlayer Then dir_ = SAVEGAME_SKI_AUTO Else dir_ = SAVEGAME_CAM_AUTO
    ' Backward compatibility: remove later
    If Not freeAutoSaveSlot_old(dir_) Then
        ' no old .gam found: check for new saves
        If bMultiPlayer Then
            freeAutoSaveSlot SAVEGAME_LOC_SKI_AUTO
        Else
            freeAutoSaveSlot SAVEGAME_LOC_CAM_AUTO
        End If
    End If

    now_ = time_()
    timeinfo = getLocalTime(now_)
    savedate = strftime("%Y-%m-%d_%H%M%S", timeinfo)   ' "%F_%H%M%S" (char savedate[PATH_MAX])

    suggestedName = suggestSaveName(dir_)
    savefile_ = dir_ & "/" & suggestedName & "_" & savedate & ".gam"   ' (char savefile[PATH_MAX])
    If saveGame(savefile_, GTYPE_SAVE_MIDMISSION, True) Then
        console "AutoSave %s", savegameWithoutExtension(savefile_)
        doAutoSave = True
    Else
        console "AutoSave %s failed", savegameWithoutExtension(savefile_)
        doAutoSave = False
    End If
End Function

Public Function autoSave(ByVal force As Boolean) As Boolean
    ' Bail out if we're running a _true_ multiplayer game or are playing a tutorial/debug/cheating/autogames
    ' (const DebugInputManager& dbgInputManager = gInputManager.debugManager();)
    If Not autosaveEnabled Or runningMultiplayer() Or bInTutorial Or DebugInputManager_debugMappingsAllowed() Or Cheated Or autogame_enabled() Then
        autoSave = False
        Exit Function
    End If
    ' Bail out if we're running a replay
    If NETisReplay() Then
        autoSave = False
        Exit Function
    End If
    ' Bail out if autosaves only mode and not forced
    If getCamTweakOption_AutosavesOnly() And Not force Then
        autoSave = False
        Exit Function
    End If

    console "AutoSaving..."
    ' queue for next main loop run, so we have a chance to render the "AutoSaving..." message before a potential delay while autosaving
    wzAsyncExecOnMainThread WzFunc_new(FN_autoSave_doAutoSave)

    autoSave = True ' autosave queued
End Function

' the lambdas
Public Function loadsave_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET
    Select Case f.id
    Case FN_loadsave_displayLoadBanner
        displayLoadBanner args(0), args(1), args(2)
    Case FN_loadsave_displayLoadSlot
        displayLoadSlot args(0), args(1), args(2)
    Case FN_loadSaveForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry LOADSAVE_X(), LOADSAVE_Y(), LOADSAVE_W, slotsInColumn * (LOADENTRY_H + LOADSAVE_HGAP) + LOADSAVE_BANNER_DEPTH + 20
    Case FN_newLoadSaveDisplayLoadSlotCache
        Set loadsave_c_fn = New LoadSaveDisplayLoadSlotCache
    Case FN_deleteLoadSaveDisplayLoadSlotCache
        Set psWidget = args(0)
        ' (assert(psWidget->pUserData != nullptr): the C assert)
        psWidget.pUserData = Empty          ' delete static_cast<LoadSaveDisplayLoadSlotCache *>(psWidget->pUserData); psWidget->pUserData = nullptr;
    Case FN_autoSave_doAutoSave
        If Not (gameInitialised) Then ASSERT_FAILURE "gameInitialised", "src/loadsave.cpp:1186", "autoSave", "Skipping autosave - game no longer initialized": Exit Function
        doAutoSave
    End Select
End Function
