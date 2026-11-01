Attribute VB_Name = "init_c"
' src/init.cpp (ported in part): the search path. Ported: versionedModsPath
' (getFullModPath, buildFullModsPaths), the search path registry
' (cleanSearchPath, registerSearchPath, unregisterSearchPath),
' getPhysFSSearchPathsAsStr, debugOutputSearchPaths,
' debugOutputSearchPathMountErrors, clearAllPhysFSSearchPaths,
' WZ_PHYSFS_MountSearchPathWrapper, rebuildExistingSearchPathWithGraphicsOptionChange,
' getTerrainOverrideBaseSourcePath, getCurrentTerrainOverrideBaseSourcePath,
' rebuildSearchPath, loadLevFile, the map list (listMapFiles, findMap,
' CheckForMod, CheckForRandom, buildMapList), systemInitialise,
' systemShutdown. WZ_Maps (std::unordered_map<std::string, WZmapInfo>) is a
' Scripting.Dictionary of name -> Array(isMapMod, isRandom).
' buildMapList reads the map packages in maps/*.wz (WzMapZipIO, wzmaplib's
' MapPackage). Missing: the in-memory map archive, loadLevFile_JSON (no
' caller). Also ported:
' getAvailableTerrainShaderQualityTextures (with its cache), InitialiseGlobals,
' frontendInitialise, frontendShutdown, the three stages' initialisation and
' shutdown (stageOneInitialise ... stageThreeShutDown, displayLoadingErrors),
' campaignReset, saveGameReset and initMiscVars; their calls into the
' display, interface, sound and scripts are stand-ins until those clusters.
' executeFnAndProcessScriptQueuedRemovals([]{ f(); }) is f followed by
' wzapi_processScriptQueuedObjectRemovals (as in loop_c). FrontImages
' (IMAGEFILE *) is resGetData's Variant.
Option Explicit

' init.h: enum searchPathMode
Public Const mod_clean As Long = 0
Public Const mod_campaign As Long = 1
Public Const mod_multiplay As Long = 2

' init.h: enum MODS_PATHS
Public Const MODS_MUSIC As Long = 0
Public Const MODS_GLOBAL As Long = 1
Public Const MODS_AUTOLOAD As Long = 2
Public Const MODS_CAMPAIGN As Long = 3
Public Const MODS_MULTIPLAY As Long = 4
Public Const MODS_PATHS_MAX As Long = 5

' init.h
Public Const FILE_LOAD_BUFFER_SIZE As Long = 1024& * 1024 * 4

Public fileLoadBuffer() As Byte         ' char fileLoadBuffer[FILE_LOAD_BUFFER_SIZE]

Public FrontImages As Variant           ' IMAGEFILE *: resGetData's data (Empty: nullptr)

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

' init.h: struct wzSearchPath
Private Type wzSearchPath
    path As String
    priority As Long                    ' unsigned int
End Type

' C: std::vector<std::unique_ptr<wzSearchPath>>, sorted by priority
Private searchPathRegistry() As wzSearchPath
Private searchPathRegistrySize As Long
Private searchPathMountErrors As StringVector

' static optional<std::vector<TerrainShaderQuality>> cachedAvailableTerrainShaderQualityTextures
Private cachedAvailableTerrainShaderQualityTextures() As Long
Private cachedAvailableTerrainShaderQualityTexturesSize As Long
Private cachedAvailableTerrainShaderQualityTexturesHasValue As Boolean

' RebuildSearchPathCommand
Private Type RebuildSearchPathCommand
    mode As Long                        ' searchPathMode (mod_clean)
    force As Boolean
    lastTerrainShaderQuality As Long    ' TerrainShaderQuality (MEDIUM)
End Type

Private lastCommand As RebuildSearchPathCommand
Private current_mode As Long            ' rebuildSearchPath's static (mod_clean)
Private cachedFullModsPaths(0 To MODS_PATHS_MAX - 1) As String  ' versionedModsPath's static
Private cachedFullModsPathsReady As Boolean
Private staticsReady As Boolean

Private Const UserMusicPath As String = "music"

Private WZ_Maps As Object               ' Scripting.Dictionary (BinaryCompare): MapName -> Array(isMapMod, isRandom)

' struct MapFileListPath
Private Type MapFileListPath
    platformIndependent As String
    platformDependent As String
End Type

' (VB) the C++ static array: fileLoadBuffer
Public Sub init_c_init()
    ReDim fileLoadBuffer(0 To FILE_LOAD_BUFFER_SIZE - 1)
End Sub

' the C++ initialisers of the statics
Private Sub InitStatics()
    If staticsReady Then Exit Sub
    staticsReady = True
    lastCommand.mode = mod_clean
    lastCommand.force = False
    lastCommand.lastTerrainShaderQuality = MEDIUM
    current_mode = mod_clean
End Sub

Private Function getFullModPath(ByVal typ As Long) As String
    Select Case typ
    Case MODS_MUSIC
        getFullModPath = version_getVersionedModsFolderPath("music")
        Exit Function
    Case MODS_GLOBAL
        getFullModPath = version_getVersionedModsFolderPath("global")
        Exit Function
    Case MODS_AUTOLOAD
        getFullModPath = version_getVersionedModsFolderPath("autoload")
        Exit Function
    Case MODS_CAMPAIGN
        ' With the new campaign mod packaging format, we no longer need a versioned path (as the mods contain version compatibility info)
        getFullModPath = "mods/campaign"
        Exit Function
    Case MODS_MULTIPLAY
        getFullModPath = version_getVersionedModsFolderPath("multiplay")
        Exit Function
    Case MODS_PATHS_MAX
    End Select
    getFullModPath = version_getVersionedModsFolderPath()
End Function

Private Sub buildFullModsPaths()
    Dim i As Long
    For i = 0 To MODS_PATHS_MAX - 1
        cachedFullModsPaths(i) = getFullModPath(i)
    Next
End Sub

Public Function versionedModsPath(ByVal typ As Long) As String
    If Not cachedFullModsPathsReady Then
        buildFullModsPaths
        cachedFullModsPathsReady = True
    End If
    versionedModsPath = cachedFullModsPaths(typ)
End Function

Private Sub cleanSearchPath()
    searchPathRegistrySize = 0
    Erase searchPathRegistry
    cachedAvailableTerrainShaderQualityTexturesHasValue = False    ' .reset()
End Sub

'!
' Register searchPath above the path with next lower priority
' For information about what can be a search path, refer to PhysFS documentation
Public Sub registerSearchPath(ByVal newPath As String, ByVal priority As Long)
    Dim tmpSearchPath As wzSearchPath, insert_pos As Long, i As Long

    If Not (Len(newPath) > 0) Then
        ASSERT_FAILURE "!newPath.empty()", "src/init.cpp:273", "registerSearchPath", "Calling registerSearchPath with empty path, priority %u", priority
        Exit Sub
    End If
    tmpSearchPath.path = newPath
    If Not strEndsWith(tmpSearchPath.path, PHYSFS_getDirSeparator()) Then
        tmpSearchPath.path = tmpSearchPath.path & PHYSFS_getDirSeparator()
    End If
    tmpSearchPath.priority = priority

    debug_ 282, LOG_WZ, "registerSearchPath", "registerSearchPath: Registering %s at priority %i", newPath, priority

    ' insert in sorted order (std::upper_bound: after the entries of the same priority)
    insert_pos = searchPathRegistrySize
    For i = 0 To searchPathRegistrySize - 1
        If priority < searchPathRegistry(i).priority Then
            insert_pos = i
            Exit For
        End If
    Next
    If searchPathRegistrySize = 0 Then
        ReDim searchPathRegistry(0 To 7)
    ElseIf searchPathRegistrySize > UBound(searchPathRegistry) Then
        ReDim Preserve searchPathRegistry(0 To 2 * searchPathRegistrySize - 1)
    End If
    For i = searchPathRegistrySize To insert_pos + 1 Step -1
        searchPathRegistry(i) = searchPathRegistry(i - 1)
    Next
    searchPathRegistry(insert_pos) = tmpSearchPath
    searchPathRegistrySize = searchPathRegistrySize + 1

    cachedAvailableTerrainShaderQualityTexturesHasValue = False    ' .reset()
End Sub

Public Sub unregisterSearchPath(ByVal path As String)
    Dim i As Long, j As Long
    i = 0
    Do While i < searchPathRegistrySize
        If searchPathRegistry(i).path = path Then
            ' ignore this notification - remove from the list
            For j = i To searchPathRegistrySize - 2
                searchPathRegistry(j) = searchPathRegistry(j + 1)
            Next
            searchPathRegistrySize = searchPathRegistrySize - 1
        Else
            i = i + 1
        End If
    Loop

    cachedAvailableTerrainShaderQualityTexturesHasValue = False    ' .reset()
End Sub

' C: std::list<std::string> getPhysFSSearchPathsAsStr(): the paths in results(0 ..), and their count
Private Function getPhysFSSearchPathsAsStr(results() As String) As Long
    getPhysFSSearchPathsAsStr = PHYSFS_getSearchPath(results)
End Function

Public Sub debugOutputSearchPaths()
    Dim idx As Long, searchPaths() As String, n As Long
    debug_ 329, LOG_INFO, "debugOutputSearchPaths", "Search path registry:"
    idx = 0
    For idx = 0 To searchPathRegistrySize - 1
        debug_ 333, LOG_INFO, "debugOutputSearchPaths", "- [%zu] [priority: %u]: %s", idx, searchPathRegistry(idx).priority, searchPathRegistry(idx).path
    Next
    debug_ 336, LOG_INFO, "debugOutputSearchPaths", "Enumerated search paths:"
    n = getPhysFSSearchPathsAsStr(searchPaths)
    For idx = 0 To n - 1
        debug_ 341, LOG_INFO, "debugOutputSearchPaths", "- [%zu]: %s", idx, searchPaths(idx)
    Next
End Sub

Public Sub debugOutputSearchPathMountErrors()
    Dim i As Long
    If searchPathMountErrors.size = 0 Then
        Exit Sub
    End If
    debug_ 352, LOG_INFO, "debugOutputSearchPathMountErrors", "Search path mount errors:"
    For i = 0 To searchPathMountErrors.size - 1
        debug_ 355, LOG_INFO, "debugOutputSearchPathMountErrors", "%s", searchPathMountErrors.items(i)
    Next
End Sub

Private Sub clearAllPhysFSSearchPaths()
    Dim searchPaths() As String, n As Long, i As Long, j As Long, searchPathsRemoved As Long
    Dim errorPaths As StringVector, errorReasons As StringVector, pErrorStr As String, reason As String

    n = getPhysFSSearchPathsAsStr(searchPaths)
    If n = 0 Then
        Exit Sub
    End If

    ' mountPathToErrorUnmounting: errorPaths(k) -> errorReasons(k)
    Do
        searchPathsRemoved = 0
        StringVector_clear errorPaths
        StringVector_clear errorReasons
        i = 0
        Do While i < n
            If PHYSFS_unmount(searchPaths(i)) <> 0 Then     ' WZ_PHYSFS_unmount
                searchPathsRemoved = searchPathsRemoved + 1
                For j = i To n - 2
                    searchPaths(j) = searchPaths(j + 1)
                Next
                n = n - 1
            Else
                pErrorStr = WZ_PHYSFS_getLastError()
                If StrPtr(pErrorStr) = 0 Then pErrorStr = "<unknown>"
                StringVector_push_back errorPaths, searchPaths(i)
                StringVector_push_back errorReasons, pErrorStr
                i = i + 1
            End If
        Loop
    Loop While n > 0 And searchPathsRemoved > 0

    For i = 0 To n - 1
        reason = ""
        For j = 0 To errorPaths.size - 1
            If errorPaths.items(j) = searchPaths(i) Then reason = errorReasons.items(j)
        Next
        debug_ 393, LOG_WZ, "clearAllPhysFSSearchPaths", "Unable to unmount search path: %s; (Reason: %s)", searchPaths(i), reason
    Next
End Sub

Private Function WZ_PHYSFS_MountSearchPathWrapper(ByVal newDir As String, ByVal mountPoint As String, ByVal appendToPath As Long) As Boolean
    Dim errorCode As Long, errorStr As String, mp As String
    If Not (StrPtr(newDir) <> 0) Then
        ASSERT_FAILURE "newDir != nullptr", "src/init.cpp:412", "WZ_PHYSFS_MountSearchPathWrapper", "Null newDir"
        Exit Function
    End If
    If PHYSFS_mount(newDir, mountPoint, appendToPath) <> 0 Then
        WZ_PHYSFS_MountSearchPathWrapper = True
    Else
        ' mount failed
        errorCode = PHYSFS_getLastErrorCode()
        If errorCode <> PHYSFS_ERR_NOT_FOUND Then
            errorStr = PHYSFS_getErrorByCode(errorCode)
            If StrPtr(mountPoint) <> 0 Then mp = mountPoint Else mp = "/"
            If StrPtr(errorStr) = 0 Then errorStr = "<no details available?>"
            StringVector_push_back searchPathMountErrors, va("Failed to mount ""%s"" @ ""%s"": %s", newDir, mp, errorStr)
        End If
        WZ_PHYSFS_MountSearchPathWrapper = False
    End If
End Function

Public Function rebuildExistingSearchPathWithGraphicsOptionChange() As Boolean
    InitStatics
    rebuildExistingSearchPathWithGraphicsOptionChange = rebuildSearchPath(lastCommand.mode, False) ' do not force
End Function

' C: optional<std::string>: has_value, and the path ByRef
Public Function getTerrainOverrideBaseSourcePath(ByVal quality As Long, path As String) As Boolean
    Select Case quality
    Case CLASSIC
        path = "terrain_overrides/classic"
        getTerrainOverrideBaseSourcePath = True
    Case MEDIUM
        getTerrainOverrideBaseSourcePath = False
    Case NORMAL_MAPPING
        path = "terrain_overrides/high"
        getTerrainOverrideBaseSourcePath = True
    Case UNINITIALIZED_PICK_DEFAULT
        getTerrainOverrideBaseSourcePath = False
    Case Else
        getTerrainOverrideBaseSourcePath = False ' silence compiler warning
    End Select
End Function

Public Function getCurrentTerrainOverrideBaseSourcePath(path As String) As Boolean
    getCurrentTerrainOverrideBaseSourcePath = getTerrainOverrideBaseSourcePath(getTerrainShaderQuality(), path)
End Function

' getAvailableTerrainShaderQualityTextures's checkValidTerrainOverridesPath lambda
Private Function checkValidTerrainOverridesPath(ByVal terrainOverridesPlatformDependentPath As String) As Boolean
    Dim mountPoint As String, actualMountPoint As String, validTerrainOverridesPackage As Boolean, checkFolder As String
    If PHYSFS_mount(terrainOverridesPlatformDependentPath, "WZTerrainOverrideTest", PHYSFS_APPEND) = 0 Then
        checkValidTerrainOverridesPath = False
        Exit Function
    End If

    ' Get the actual mount point for this archive
    ' (Why? Because it could have been mounted *before* this check at a different location, if it's the currently-selected terrain override pack and is present...)
    mountPoint = PHYSFS_getMountPoint(terrainOverridesPlatformDependentPath)
    If StrPtr(mountPoint) <> 0 Then actualMountPoint = mountPoint Else actualMountPoint = "WZTerrainOverrideTest"
    If Not strEndsWith(actualMountPoint, "/") Then
        actualMountPoint = actualMountPoint & "/"
    End If

    validTerrainOverridesPackage = False

    ' Sanity check: For texpages and/or tileset folder(s)
    checkFolder = actualMountPoint & "texpages"
    If WZ_PHYSFS_isDirectory(checkFolder) <> 0 Then
        validTerrainOverridesPackage = True
    End If
    checkFolder = actualMountPoint & "tileset"
    If Not validTerrainOverridesPackage And WZ_PHYSFS_isDirectory(checkFolder) <> 0 Then
        validTerrainOverridesPackage = True
    End If

    If PHYSFS_unmount(terrainOverridesPlatformDependentPath) = 0 Then     ' WZ_PHYSFS_unmount
        debug_ 507, LOG_WZ, "getAvailableTerrainShaderQualityTextures", "* Failed to remove path %s again", terrainOverridesPlatformDependentPath
    End If

    checkValidTerrainOverridesPath = validTerrainOverridesPackage
End Function

' std::vector<TerrainShaderQuality> getAvailableTerrainShaderQualityTextures(): the values in result (0-based), their
' number returned
Public Function getAvailableTerrainShaderQualityTextures(result() As Long) As Long
    Dim allQualities() As Long, allQualities_size As Long, q As Long, i As Long, n As Long, tmpstr As String
    Dim textureOverrideBaseSourcePath As String, foundTexturePack As Boolean

    ' Cached result, which is invalidated with any calls to: cleanSearchPath / registerSearchPath / unregisterSearchPath
    If cachedAvailableTerrainShaderQualityTexturesHasValue Then
        n = cachedAvailableTerrainShaderQualityTexturesSize
        If n > 0 Then
            ReDim result(0 To n - 1)
            For i = 0 To n - 1
                result(i) = cachedAvailableTerrainShaderQualityTextures(i)
            Next
        Else
            Erase result
        End If
        getAvailableTerrainShaderQualityTextures = n
        Exit Function
    End If

    allQualities_size = getAllTerrainShaderQualityOptions(allQualities)
    Erase result
    n = 0

    For q = 0 To allQualities_size - 1
        foundTexturePack = False
        If getTerrainOverrideBaseSourcePath(allQualities(q), textureOverrideBaseSourcePath) Then
            ' Check all search paths for this folder or .wz file
            For i = 0 To searchPathRegistrySize - 1
                ' Check <path> (folder)
                tmpstr = searchPathRegistry(i).path & textureOverrideBaseSourcePath
                If checkValidTerrainOverridesPath(tmpstr) Then
                    foundTexturePack = True
                    Exit For
                End If

                ' Check <path>.wz
                tmpstr = tmpstr & ".wz"
                If checkValidTerrainOverridesPath(tmpstr) Then
                    foundTexturePack = True
                    Exit For
                End If
            Next
        Else
            ' quality doesn't need a separate texture pack - always available
            foundTexturePack = True
        End If

        If foundTexturePack Then
            If n = 0 Then ReDim result(0 To 0) Else ReDim Preserve result(0 To n)
            result(n) = allQualities(q)
            n = n + 1
        End If
    Next

    cachedAvailableTerrainShaderQualityTexturesSize = n
    If n > 0 Then
        ReDim cachedAvailableTerrainShaderQualityTextures(0 To n - 1)
        For i = 0 To n - 1
            cachedAvailableTerrainShaderQualityTextures(i) = result(i)
        Next
    End If
    cachedAvailableTerrainShaderQualityTexturesHasValue = True

    getAvailableTerrainShaderQualityTextures = n
End Function

'!
' \brief Rebuilds the PHYSFS searchPath with mode specific subdirs
'
' Priority:
' maps > mods > base > base.wz
Public Function rebuildSearchPath(ByVal mode As Long, ByVal force As Boolean) As Boolean
    Dim tmpstr As String, currentTerrainShaderQuality As Long, i As Long, k As Long, curPath As String
    Dim hasTerrainOverride As Boolean, terrainQualityOverrideBasePath As String
    Dim loadedTerrainTextureOverrides As Boolean, noList As StringVector, hashList As StringVector

    InitStatics
    currentTerrainShaderQuality = getTerrainShaderQuality()

    If mode <> current_mode _
        Or force _
        Or lastCommand.lastTerrainShaderQuality <> currentTerrainShaderQuality _
        Or (use_override_mods And override_mod_list <> getModList()) Then

        If mode <> mod_clean Then
            rebuildSearchPath mod_clean, False
        End If

        current_mode = mode
        lastCommand.mode = mode ' store this separately, so it doesn't get overridden by anything below
        lastCommand.force = force
        lastCommand.lastTerrainShaderQuality = currentTerrainShaderQuality

        hasTerrainOverride = getCurrentTerrainOverrideBaseSourcePath(terrainQualityOverrideBasePath)
        loadedTerrainTextureOverrides = False

        Select Case mode
        Case mod_clean
            debug_ 590, LOG_WZ, "rebuildSearchPath", "Cleaning up"
            clearLoadedMods

            For i = 0 To searchPathRegistrySize - 1
                curPath = searchPathRegistry(i).path
                ' Remove maps and mods
                removeSubdirs curPath, "maps"
                removeSubdirs curPath, versionedModsPath(MODS_MUSIC)
                removeSubdirs curPath, versionedModsPath(MODS_GLOBAL)
                removeSubdirs curPath, versionedModsPath(MODS_AUTOLOAD)
                removeSubdirs curPath, versionedModsPath(MODS_CAMPAIGN)
                removeSubdirs curPath, versionedModsPath(MODS_MULTIPLAY)
                removeSubdirs curPath, "mods/downloads" ' not versioned

                ' Remove multiplay patches
                tmpstr = curPath & "mp"
                PHYSFS_unmount tmpstr
                tmpstr = curPath & "mp.wz"
                PHYSFS_unmount tmpstr

                ' Remove plain dir
                PHYSFS_unmount curPath

                ' Remove base files
                tmpstr = curPath & "base"
                PHYSFS_unmount tmpstr
                tmpstr = curPath & "base.wz"
                PHYSFS_unmount tmpstr

                ' remove video search path as well
                tmpstr = curPath & "sequences.wz"
                PHYSFS_unmount tmpstr
            Next

            ' This should properly remove all paths, but testing is needed to ensure that all supported versions of PhysFS behave as expected
            ' For now, keep the old code above as well as this new method
            clearAllPhysFSSearchPaths
            StringVector_clear searchPathMountErrors

        Case mod_campaign
            debug_ 633, LOG_WZ, "rebuildSearchPath", "*** Switching to campaign mods ***"
            clearLoadedMods

            For i = 0 To searchPathRegistrySize - 1
                ' make sure videos override included files
                tmpstr = searchPathRegistry(i).path & "sequences.wz"
                WZ_PHYSFS_MountSearchPathWrapper tmpstr, vbNullString, PHYSFS_APPEND
            Next

            For i = 0 To searchPathRegistrySize - 1
                curPath = searchPathRegistry(i).path
                ' Add global and campaign mods
                WZ_PHYSFS_MountSearchPathWrapper curPath, vbNullString, PHYSFS_APPEND

                addSubdirs curPath, versionedModsPath(MODS_MUSIC), True, noList, False, False
                If use_override_mods Then
                    addSubdirs curPath, versionedModsPath(MODS_GLOBAL), True, override_mods, True, True
                    addSubdirs curPath, versionedModsPath(MODS_AUTOLOAD), True, override_mods, True, True
                    addSubdirs curPath, versionedModsPath(MODS_CAMPAIGN), True, override_mods, True, True
                Else
                    addSubdirs curPath, versionedModsPath(MODS_GLOBAL), True, global_mods, True, True
                    addSubdirs curPath, versionedModsPath(MODS_AUTOLOAD), True, noList, False, True
                    addSubdirs curPath, versionedModsPath(MODS_CAMPAIGN), True, campaign_mods, True, True
                End If
                If PHYSFS_unmount(curPath) = 0 Then
                    debug_ 657, LOG_WZ, "rebuildSearchPath", "* Failed to remove path %s again", curPath
                End If

                ' Add plain dir
                WZ_PHYSFS_MountSearchPathWrapper curPath, vbNullString, PHYSFS_APPEND

                If hasTerrainOverride And Not loadedTerrainTextureOverrides Then
                    ' Add terrain quality override files
                    tmpstr = curPath & terrainQualityOverrideBasePath
                    loadedTerrainTextureOverrides = WZ_PHYSFS_MountSearchPathWrapper(tmpstr, vbNullString, PHYSFS_APPEND) Or loadedTerrainTextureOverrides
                    tmpstr = tmpstr & ".wz"
                    loadedTerrainTextureOverrides = WZ_PHYSFS_MountSearchPathWrapper(tmpstr, vbNullString, PHYSFS_APPEND) Or loadedTerrainTextureOverrides
                    If loadedTerrainTextureOverrides Then
                        debug_ 672, LOG_INFO, "rebuildSearchPath", "Loaded terrain overrides from: %s", curPath
                    End If
                End If

                ' Add base files
                tmpstr = curPath & "base"
                WZ_PHYSFS_MountSearchPathWrapper tmpstr, vbNullString, PHYSFS_APPEND
                tmpstr = curPath & "base.wz"
                WZ_PHYSFS_MountSearchPathWrapper tmpstr, vbNullString, PHYSFS_APPEND
            Next

        Case mod_multiplay
            debug_ 684, LOG_WZ, "rebuildSearchPath", "*** Switching to multiplay mods ***"
            clearLoadedMods

            For i = 0 To searchPathRegistrySize - 1
                ' make sure videos override included files
                tmpstr = searchPathRegistry(i).path & "sequences.wz"
                WZ_PHYSFS_MountSearchPathWrapper tmpstr, vbNullString, PHYSFS_APPEND
            Next
            For i = 0 To searchPathRegistrySize - 1
                curPath = searchPathRegistry(i).path
                ' Add global and multiplay mods
                WZ_PHYSFS_MountSearchPathWrapper curPath, vbNullString, PHYSFS_APPEND
                addSubdirs curPath, versionedModsPath(MODS_MUSIC), True, noList, False, False

                ' Only load if we are host or singleplayer (Initial mod load relies on this, too)
                If ingame.side = InGameSide_HOST_OR_SINGLEPLAYER Or Not NetPlay_bComms Then
                    If use_override_mods Then
                        addSubdirs curPath, versionedModsPath(MODS_GLOBAL), True, override_mods, True, True
                        addSubdirs curPath, versionedModsPath(MODS_AUTOLOAD), True, override_mods, True, True
                        addSubdirs curPath, versionedModsPath(MODS_MULTIPLAY), True, override_mods, True, True
                    Else
                        addSubdirs curPath, versionedModsPath(MODS_GLOBAL), True, global_mods, True, True
                        addSubdirs curPath, versionedModsPath(MODS_AUTOLOAD), True, noList, False, True
                        addSubdirs curPath, versionedModsPath(MODS_MULTIPLAY), True, multiplay_mods, True, True
                    End If
                Else
                    For k = 0 To game.modHashCount - 1
                        StringVector_clear hashList
                        StringVector_push_back hashList, Sha256_toString(game.modHashes(k))
                        addSubdirs curPath, "mods/downloads", True, hashList, True, True ' not versioned
                    Next
                End If
                PHYSFS_unmount curPath

                ' Add multiplay patches
                tmpstr = curPath & "mp"
                WZ_PHYSFS_MountSearchPathWrapper tmpstr, vbNullString, PHYSFS_APPEND
                tmpstr = curPath & "mp.wz"
                WZ_PHYSFS_MountSearchPathWrapper tmpstr, vbNullString, PHYSFS_APPEND

                ' Add plain dir
                WZ_PHYSFS_MountSearchPathWrapper curPath, vbNullString, PHYSFS_APPEND

                If hasTerrainOverride And Not loadedTerrainTextureOverrides Then
                    ' Add terrain quality override files
                    tmpstr = curPath & terrainQualityOverrideBasePath
                    loadedTerrainTextureOverrides = WZ_PHYSFS_MountSearchPathWrapper(tmpstr, vbNullString, PHYSFS_APPEND) Or loadedTerrainTextureOverrides
                    tmpstr = tmpstr & ".wz"
                    loadedTerrainTextureOverrides = WZ_PHYSFS_MountSearchPathWrapper(tmpstr, vbNullString, PHYSFS_APPEND) Or loadedTerrainTextureOverrides
                    If loadedTerrainTextureOverrides Then
                        debug_ 738, LOG_INFO, "rebuildSearchPath", "Loaded terrain overrides from: %s", curPath
                    End If
                End If

                ' Add base files
                tmpstr = curPath & "base"
                WZ_PHYSFS_MountSearchPathWrapper tmpstr, vbNullString, PHYSFS_APPEND
                tmpstr = curPath & "base.wz"
                WZ_PHYSFS_MountSearchPathWrapper tmpstr, vbNullString, PHYSFS_APPEND
            Next

        Case Else
            debug_ 750, LOG_ERROR, "rebuildSearchPath", "Can't switch to unknown mods %i", mode
            rebuildSearchPath = False
            Exit Function
        End Select
        If use_override_mods And mode <> mod_clean Then
            If getModList() <> override_mod_list Then
                debug_ 757, LOG_POPUP, "rebuildSearchPath", "The required mod could not be loaded: %s\n\nWarzone will try to load the game without it.", override_mod_list
            End If
        End If

        ' User's home dir must be first so we always see what we write
        PHYSFS_unmount PHYSFS_getWriteDir()
        WZ_PHYSFS_MountSearchPathWrapper PHYSFS_getWriteDir(), vbNullString, PHYSFS_PREPEND

        If mode <> mod_clean Then
            gfx_api_loadTextureCompressionOverrides ' reload these as mods may have overridden the file!
        End If

        If hasTerrainOverride And (mode = mod_campaign Or mode = mod_multiplay) Then
            If Not loadedTerrainTextureOverrides Then
                debug_ 778, LOG_INFO, "rebuildSearchPath", "Failed to load expected terrain quality overrides: %s", terrainQualityOverrideBasePath
            End If
        End If

        ActivityManager_rebuiltSearchPath
    End If
    rebuildSearchPath = True
End Function

' (realFileName vbNullString: nullptr)
Public Function loadLevFile(ByVal filename As String, ByVal pathMode As Long, ByVal ignoreWrf As Boolean, ByVal realFileName As String) As Boolean
    Dim pBuffer() As Byte, size As Long

    If StrPtr(realFileName) = 0 Then
        debug_ 216, LOG_WZ, "loadLevFile", "Loading lev file: ""%s"", builtin\n", filename
    Else
        debug_ 220, LOG_WZ, "loadLevFile", "Loading lev file: ""%s"" from ""%s""\n", filename, realFileName
    End If

    If PHYSFS_exists(filename) = 0 Then
        debug_ 225, LOG_ERROR, "loadLevFile", "File not found: %s\n", filename
        loadLevFile = False ' only in NDEBUG case
        Exit Function
    End If
    If Not loadFile(filename, pBuffer, size) Then
        debug_ 225, LOG_ERROR, "loadLevFile", "File not found: %s\n", filename
        loadLevFile = False ' only in NDEBUG case
        Exit Function
    End If
    If Not levParse(pBuffer, size, pathMode, ignoreWrf, realFileName) Then
        debug_ 230, LOG_ERROR, "loadLevFile", "Parse error in %s\n", filename
        loadLevFile = False
        Exit Function
    End If

    loadLevFile = True
End Function

' maps/*.wz in files(0 .. count - 1)
Private Function listMapFiles(ret() As MapFileListPath) As Long
    Dim names() As String, n As Long, i As Long, count As Long, wzfile As String, dotPos As Long
    WZ_PHYSFS_enumerateFiles "maps", names, n
    For i = 0 To n - 1
        wzfile = names(i)
        dotPos = InStrRev(wzfile, ".")
        If Left$(wzfile, 1) = "." Or Mid$(wzfile, dotPos + 1) <> "wz" Then
            GoTo continue_ ' continue;
        End If

        If count = 0 Then ReDim ret(0 To n - 1)
        ret(count).platformIndependent = "maps" & "/" & names(i)
        ret(count).platformDependent = "maps" & PHYSFS_getDirSeparator() & names(i)
        count = count + 1
continue_:
    Next

    listMapFiles = count
End Function

' Map processing

Private Sub InitWZMaps()
    If WZ_Maps Is Nothing Then
        Set WZ_Maps = CreateObject("Scripting.Dictionary")
        WZ_Maps.CompareMode = 0         ' BinaryCompare
    End If
End Sub

' (mapFile vbNullString: nullptr)
Public Function CheckForMod(ByVal mapFile As String) As Boolean
    InitWZMaps
    If StrPtr(mapFile) <> 0 Then
        If WZ_Maps.Exists(mapFile) Then
            CheckForMod = WZ_Maps(mapFile)(0)
            Exit Function
        End If
    End If
    If StrPtr(mapFile) <> 0 Then
        debug_ 835, LOG_ERROR, "CheckForMod", "Couldn't find map %s", mapFile
    End If

    CheckForMod = False
End Function

' (mapFile, mapDataFile0 vbNullString: nullptr)
Public Function CheckForRandom(ByVal mapFile As String, ByVal mapDataFile0 As String) As Boolean
    Dim fn As String
    InitWZMaps
    If StrPtr(mapFile) <> 0 Then
        If WZ_Maps.Exists(mapFile) Then
            CheckForRandom = WZ_Maps(mapFile)(1)
            Exit Function
        End If
    End If

    If StrPtr(mapFile) <> 0 Then
        debug_ 850, LOG_ERROR, "CheckForRandom", "Couldn't find map %s", mapFile
    ElseIf StrPtr(mapDataFile0) <> 0 Then
        If Len(mapDataFile0) > 4 Then
            fn = mapDataFile0
            fn = Left$(fn, Len(fn) - 4)
            fn = fn & "/game.js"
            CheckForRandom = (PHYSFS_exists(fn) <> 0)
            Exit Function
        End If
    End If

    CheckForRandom = False
End Function

' static inline WZmapInfo CheckInMap(WzMap::MapPackage& mapPackage): Array(isMapMod, isRandom)
Private Function CheckInMap(ByVal mapPackage_ As MapPackage) As Variant
    Dim mapmod As Boolean, isRandom As Boolean
    mapmod = mapPackage_.packageType() = MapPackageType_Map_Mod
    isRandom = mapPackage_.isScriptGeneratedMap()
    CheckInMap = Array(mapmod, isRandom)
End Function

Public Function buildMapList(Optional ByVal campaignOnly As Boolean = False) As Boolean
    Dim realFileNames() As MapFileListPath, n As Long, i As Long, pRealDirStr As String
    Dim realFilePathAndName As String, zipReadSource As WzZipIOPHYSFSSourceReadProvider
    Dim debugLoggerInstance As WzMapLoadDebugLogger, mapZipIO As WzMapZipIO, mapPackage_ As MapPackage
    Dim WZmapInfoResult As Variant
    InitWZMaps
    If Not loadLevFile("gamedesc.lev", mod_campaign, False, vbNullString) Then
        buildMapList = False
        Exit Function
    End If
    If Not campaignOnly Then
        loadLevFile "addon.lev", mod_multiplay, False, vbNullString
    End If
    WZ_Maps.RemoveAll
    ' (the in-memory map archive is not ported: there is none to clear)
    If campaignOnly Then
        buildMapList = True
        Exit Function
    End If
    n = listMapFiles(realFileNames)
    Set debugLoggerInstance = New WzMapLoadDebugLogger
    For i = 0 To n - 1
        pRealDirStr = PHYSFS_getRealDir(realFileNames(i).platformIndependent)
        If StrPtr(pRealDirStr) = 0 Then
            debug_ 1011, LOG_ERROR, "buildMapList", "Failed to find realdir for: %s", realFileNames(i).platformIndependent
            GoTo continue_ ' skip
        End If
        realFilePathAndName = pRealDirStr & realFileNames(i).platformDependent

        Set zipReadSource = WzZipIOPHYSFSSourceReadProvider_make(realFileNames(i).platformIndependent)
        If zipReadSource Is Nothing Then
            debug_ 1019, LOG_ERROR, "buildMapList", "Failed to open: %s", realFileNames(i).platformIndependent
            GoTo continue_
        End If

        debugLoggerInstance.setLogErrors True
        Set mapZipIO = WzMapZipIO_openZipArchiveReadIOProvider(zipReadSource, debugLoggerInstance)
        If mapZipIO Is Nothing Then
            debug_ 1027, LOG_INFO, "buildMapList", "Failed to open archive: %s." & vbLf & "Please delete or move the file specified.", realFilePathAndName
            GoTo continue_
        End If
        debugLoggerInstance.setLogErrors False
        Set mapPackage_ = MapPackage_loadPackage("", debugLoggerInstance, mapZipIO)
        If mapPackage_ Is Nothing Then
            debug_ 1034, LOG_INFO, "buildMapList", "Failed to load %s." & vbLf & "Please delete or move the file specified.", realFilePathAndName
            GoTo continue_
        End If

        If Not levAddWzMap(mapPackage_.levelDetails(), mod_multiplay, realFileNames(i).platformIndependent) Then
            debug_ 1040, LOG_ERROR, "buildMapList", "Corrupt / invalid map file: %s", realFilePathAndName
            GoTo continue_
        End If

        WZmapInfoResult = CheckInMap(mapPackage_)
        If Not WZ_Maps.Exists(realFileNames(i).platformIndependent) Then           ' std::map::insert keeps an existing key
            WZ_Maps.Add realFileNames(i).platformIndependent, WZmapInfoResult
        End If
continue_:
        Set mapPackage_ = Nothing
        Set mapZipIO = Nothing                  ' (VB) the archive unmounted
        Set zipReadSource = Nothing
    Next

    buildMapList = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' Called once on program startup.
'
Public Function systemInitialise(ByVal horizScalePercentage As Long, ByVal vertScalePercentage As Long) As Boolean
    Dim configDefaultAI As String
    If Not widgInitialise() Then
        systemInitialise = False
        Exit Function
    End If

    If Not notificationsInitialize() Then
        systemInitialise = False
        Exit Function
    End If

    buildMapList

    ' Initialize render engine
    If Not pie_Initialise() Then
        debug_ 1072, LOG_ERROR, "systemInitialise", "Unable to initialise renderer"
        systemInitialise = False
        Exit Function
    End If

    If Not audio_Init(CB_droidAudioTrackStopped, war_GetHRTFMode(), war_getSoundEnabled()) Then
        debug_ 1078, LOG_SOUND, "systemInitialise", "Continuing without audio"
    End If
    If war_getSoundEnabled() And war_GetMusicEnabled() Then
        cdAudio_Open UserMusicPath
    Else
        debug_ 1086, LOG_SOUND, "systemInitialise", "Music disabled"
    End If

    If Not dataInitLoadFuncs() Then ' Pass all the data loading functions to the framework library
        systemInitialise = False
        Exit Function
    End If

    ' Initialize the iVis text rendering module
    wzSceneBegin "Main menu loop"
    iV_TextInit horizScalePercentage, vertScalePercentage

    pie_InitRadar

    readAIs
    ' Reset config and order if default not found (user messed with config or maybe a mod AI is now missing).
    configDefaultAI = defaultSkirmishAI
    If configDefaultAI <> getDefaultSkirmishAI() Then
        readAIs
    End If

    initTerrainShaderType

    systemInitialise = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' Called once at program shutdown.
'
Public Sub systemShutdown()
    If bLoadSaveUp Then
        closeLoadSaveOnShutdown ' TODO: Ideally this would not be required here (refactor loadsave.cpp / frontend.cpp?)
    End If

    NETclose

    urlRequestShutdown ' MUST come after NETclose(), as hosts need a chance to inform lobby they are gone

    seqReleaseAll

    pie_ShutdownRadar
    clearLoadedMods
    shutdownConsoleMessages

    shutdownEffectsSystem
    wzSceneEnd vbNullString  ' Might want to end the "Main menu loop" or "Main game loop".
    gInputManager_shutdown

    ' free up all the load functions (all the data should already have been freed)
    resReleaseAll

    If Not multiShutdown() Then ' ajl. init net stuff
        debug_ 1143, LOG_FATAL, "systemShutdown", "Unable to multiShutdown() cleanly!"
        exit_ 3                 ' abort()
    End If

    ' shut down various databases
    shutdownKnownPlayers

    debug_ 1150, LOG_MAIN, "systemShutdown", "shutting down audio subsystems"

    debug_ 1152, LOG_MAIN, "systemShutdown", "shutting down CD audio"
    cdAudio_Close

    If Not audio_Disabled() Then
        If Not audio_Shutdown() Then
            debug_ 1157, LOG_FATAL, "systemShutdown", "Unable to audio_Shutdown() cleanly!"
            exit_ 3             ' abort()
        End If
    End If

    debug_ 1161, LOG_MAIN, "systemShutdown", "shutting down graphics subsystem"
    levShutDown
    notificationsShutDown
    widgShutDown
    fpathShutdown
    mapShutdown
    modelShutdown
    debug_ 1168, LOG_MAIN, "systemShutdown", "shutting down everything else"
    pal_ShutDown        ' currently unused stub
    frameShutDown       ' close screen / SDL / resources / cursors / trig
    screenShutDown
    netplayShutDown     ' MUST come after widgShutDown (as widget screens might have connections, etc)
    gfx_api_context_shutdown
    cleanSearchPath     ' clean PHYSFS search paths
    debug_exit          ' cleanup debug routines
    PHYSFS_deinit       ' cleanup PHYSFS (If failure, state of PhysFS is undefined, and probably badly screwed up.)
    ' NOTE: Exception handler is cleaned via atexit(ExchndlShutdown);
End Sub

' Each module in the game should have a call from here to initialise
' any globals and statics to there default values each time the game
' or frontend restarts.
'
Private Function InitialiseGlobals() As Boolean
    frontendInitVars    ' Initialise frontend globals and statics.
    structureInitVars
    If Not messageInitVars() Then
        InitialiseGlobals = False
        Exit Function
    End If
    If Not researchInitVars() Then
        InitialiseGlobals = False
        Exit Function
    End If
    featureInitVars
    radarInitVars
    Edit3DInitVars

    InitialiseGlobals = True
End Function

' (VB) dst = src for a Variant that may hold an object (resGetData's IMAGEFILE *)
Private Sub assignData(dst As Variant, src As Variant)
    If IsObject(src) Then Set dst = src Else dst = src
End Sub

' (VB) FrontImages == nullptr: resGetData's Empty, or a loader's Nothing
Private Function isNullData(v As Variant) As Boolean
    If IsObject(v) Then
        isNullData = (v Is Nothing)
    Else
        isNullData = IsEmpty(v)
    End If
End Function

'/***************************************************************************/

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' Called At Frontend Startup.

Public Function frontendInitialise(ByVal ResourceFile As String) As Boolean
    Dim errorMessage As String, mods() As WzMods_LoadedMod
    frontendIsShuttingDown

    debug_ 1190, LOG_WZ, "frontendInitialise", "== Initializing frontend == : %s", ResourceFile

    If Not InitialiseGlobals() Then              ' Initialise all globals and statics everywhere.
        frontendInitialise = False
        Exit Function
    End If

    If Not stringsInitialise() Then              ' Initialise the string system
        frontendInitialise = False
        Exit Function
    End If

    If Not objInitialise() Then                  ' Initialise the object system
        frontendInitialise = False
        Exit Function
    End If

    If Not allocPlayerPower() Then   'set up the PlayerPower for each player - this should only be done ONCE now
        frontendInitialise = False
        Exit Function
    End If

    debug_ 1212, LOG_MAIN, "frontendInitialise", "frontEndInitialise: loading resource file ....."
    If Not resLoad(ResourceFile, 0) Then
        'need the object heaps to have been set up before loading in the save game
        frontendInitialise = False
        Exit Function
    End If

    If Not dispInitialise() Then                 ' Initialise the display system
        frontendInitialise = False
        Exit Function
    End If

    assignData FrontImages, resGetData("IMG", "frontend.img")
    If isNullData(FrontImages) Then
        errorMessage = va("Unable to load: %s.", "frontend.img")
        If getLoadedMods(mods) > 0 Then
            errorMessage = errorMessage & " "
            errorMessage = errorMessage & "Please remove all incompatible mods."
        End If
        debug_ 1233, LOG_FATAL, "frontendInitialise", "%s", errorMessage
        frontendInitialise = False
        Exit Function
    End If

    ' Shift the interface initialisation here temporarily so that it
    ' can pick up the stats after they have been loaded
    If Not intInitialise() Then
        frontendInitialise = False
        Exit Function
    End If

    ' reinitialise key mappings
    gInputManager_resetMappings False

    ' Set the default uncoloured cursor here, since it looks slightly
    ' better for menus and such.
    wzSetCursor CURSOR_DEFAULT

    SetFormAudioIDs -1, ID_SOUND_WINDOWCLOSE         ' disable the open noise since distorted in 3dfx builds.

    initMiscVars

    gameTimeInit

    ' hit me with some funky beats....
    cdAudio_PlayTrack SONG_FRONTEND

    frontendInitialise = True
End Function


Public Function frontendShutdown() As Boolean
    debug_ 1266, LOG_WZ, "frontendShutdown", "== Shutting down frontend =="

    saveConfig ' save settings to registry.

    If Not mechanicsShutdown() Then
        frontendShutdown = False
        Exit Function
    End If

    changeTitleUI Nothing
    If challengesUp Then
        closeChallenges ' TODO: Ideally this would not be required here (refactor challenge.cpp / frontend.cpp?)
    End If
    If bLoadSaveUp Then
        closeLoadSaveOnShutdown ' TODO: Ideally this would not be required here (refactor loadsave.cpp / frontend.cpp?)
    End If
    interfaceShutDown
    shutdownGameGuide

    'do this before shutting down the iV library
    resReleaseAllData
    FrontImages = Empty
    frontendIsShuttingDown

    If Not objShutdown() Then
        frontendShutdown = False
        Exit Function
    End If

    ResearchRelease

    debug_ 1299, LOG_TEXTURE, "frontendShutdown", "=== frontendShutdown ==="
    modelShutdown
    pie_TexShutDown
    pie_TexInit ' ready for restart
    freeComponentLists
    statsShutDown

    frontendShutdown = True
End Function


'/******************************************************************************/
'/*                       Initialisation before data is loaded                 */



Public Function stageOneInitialise() As Boolean
    debug_ 1317, LOG_WZ, "stageOneInitialise", "== stageOneInitialise =="
    wzSceneEnd "Main menu loop"
    wzSceneBegin "Main game loop"

    ' Initialise all globals and statics everywhere.
    If Not InitialiseGlobals() Then
        stageOneInitialise = False
        Exit Function
    End If

    If Not stringsInitialise() Then  ' Initialise the string system
        stageOneInitialise = False
        Exit Function
    End If

    If Not objInitialise() Then      ' Initialise the object system
        stageOneInitialise = False
        Exit Function
    End If

    If Not droidInit() Then
        stageOneInitialise = False
        Exit Function
    End If

    If Not initViewData() Then
        stageOneInitialise = False
        Exit Function
    End If

    If Not grpInitialise() Then
        stageOneInitialise = False
        Exit Function
    End If

    If Not aiInitialise() Then       ' Initialise the AI system  ' pregame
        stageOneInitialise = False
        Exit Function
    End If

    If Not allocPlayerPower() Then   'set up the PlayerPower for each player - this should only be done ONCE now
        stageOneInitialise = False
        Exit Function
    End If

    If Not formationInitialise() Then        ' Initialise the formation system
        stageOneInitialise = False
        Exit Function
    End If

    ' initialise the visibility stuff
    If Not visInitialise() Then
        stageOneInitialise = False
        Exit Function
    End If

    If Not proj_InitSystem() Then
        stageOneInitialise = False
        Exit Function
    End If

    If Not gridInitialise() Then
        stageOneInitialise = False
        Exit Function
    End If

    initMission
    initTransporters
    scriptInit

    gameTimeInit
    transitionInit
    resetScroll

    strategyPlansInit

    stageOneInitialise = True
End Function


'/******************************************************************************/
'/*                       Shutdown after data is released                      */

Public Function stageOneShutDown() As Boolean
    Dim needsLevReload As Boolean
    debug_ 1402, LOG_WZ, "stageOneShutDown", "== stageOneShutDown =="

    atmosSetWeatherType WT_NONE ' reset weather and free its data
    wzPerfShutdown

    pie_FreeShaders

    If Not audio_Disabled() Then
        sound_CheckAllUnloaded
    End If

    shutdownGameGuide

    proj_Shutdown

    releaseMission

    If Not aiShutdown() Then
        stageOneShutDown = False
        Exit Function
    End If

    If Not objShutdown() Then
        stageOneShutDown = False
        Exit Function
    End If

    grpShutDown

    formationShutDown

    ResearchRelease

    'free up the gateway stuff?
    gwShutDown

    shutdownTerrain

    If Not mapShutdown() Then
        stageOneShutDown = False
        Exit Function
    End If

    gridShutDown

    debug_ 1448, LOG_TEXTURE, "stageOneShutDown", "== stageOneShutDown =="
    modelShutdown
    pie_TexShutDown

    needsLevReload = (hasOverrideMods() Or hasCampaignMods()) And (game.typ = LEVEL_TYPE_CAMPAIGN)
    clearOverrideMods
    clearCampaignMods
    clearCamTweakOptions

    ' Use mod_multiplay as the default (campaign might have set it to mod_singleplayer)
    rebuildSearchPath mod_multiplay, True

    If needsLevReload Then
        ' Clear & reload level list
        levShutDown
        levInitialise
        pal_Init ' Update palettes.
        If Not buildMapList(False) Then
            debug_ 1468, LOG_FATAL, "stageOneShutDown", "Failed to rebuild map / level list?"
        End If
    End If

    clearSpecialInMemoryMap

    pie_TexInit ' restart it

    initMiscVars
    reinitFactionsMapping
    wzSceneEnd "Main game loop"
    wzSceneBegin "Main menu loop"

    stageOneShutDown = True
End Function


' ////////////////////////////////////////////////////////////////////////////
' Initialise after the base data is loaded but before final level data is loaded

Public Function stageTwoInitialise() As Boolean
    Dim i As Long

    debug_ 1492, LOG_WZ, "stageTwoInitialise", "== stageTwoInitialise =="

    ' prevent "double up" and "biffer baker" cheats from messing up damage modifiers
    resetDamageModifiers

    ' make sure we clear on loading; this a bad hack to fix a bug when
    ' loading a savegame where we are building a lassat
    For i = 0 To MAX_PLAYERS - 1
        setLasSatExists False, i
    Next

    If Not dispInitialise() Then     ' Initialise the display system
        stageTwoInitialise = False
        Exit Function
    End If

    If Not initMiscImds() Then           ' Set up the explosions
        debug_ 1511, LOG_FATAL, "stageTwoInitialise", "Can't find all the explosions graphics?"
        exit_ 3                 ' abort()
        stageTwoInitialise = False
        Exit Function
    End If

    If Not cmdDroidInit() Then
        stageTwoInitialise = False
        Exit Function
    End If

    ' Shift the interface initialisation here temporarily so that it
    ' can pick up the stats after they have been loaded

    If Not intInitialise() Then
        stageTwoInitialise = False
        Exit Function
    End If

    If Not initMessage() Then            ' Initialise the message heaps
        stageTwoInitialise = False
        Exit Function
    End If

    If Not gwInitialise() Then
        stageTwoInitialise = False
        Exit Function
    End If

    If Not initScripts() Then        ' Initialise the new javascript system
        stageTwoInitialise = False
        Exit Function
    End If

    ' reinitialise key mappings
    gInputManager_resetMappings False

    If NETisReplay() Then
        ' Debug keybinds start enabled for replays
        gInputManager_contexts_set InputContext_DEBUG_MISC, InputContext_State_ACTIVE
    End If

    ' Set the default uncoloured cursor here, since it looks slightly
    ' better for menus and such.
    wzSetCursor CURSOR_DEFAULT

    SetFormAudioIDs ID_SOUND_WINDOWOPEN, ID_SOUND_WINDOWCLOSE

    ' Setup game queues.
    ' Don't ask why this doesn't go in stage three. In fact, don't even ask me what stage one/two/three is supposed to mean, it seems about as descriptive as stage doStuff, stage doMoreStuff and stage doEvenMoreStuff...
    debug_ 1561, LOG_MAIN, "stageTwoInitialise", "Init game queues, I am %d.", selectedPlayer
    sendQueuedDroidInfo  ' Discard any pending orders which could later get flushed into the game queue.
    If Not NETisReplay() Then
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            NETinitQueue NETgameQueue(i)

            If Not myResponsibility(i) Or Not NetPlay_bComms Then
                NETsetNoSendOverNetwork NETgameQueue(i)
            End If
        Next
    End If

    ' NOTE:
    ' - Loading savegames calls `fPathDroidRoute` (from `loadSaveDroid`) before `stageThreeInitialise`
    '   is called, hence removing this call to `fpathInitialise` will currently break loading certain
    '   savegames (if they interact with the fPath system).
    If Not fpathInitialise() Then
        stageTwoInitialise = False
        Exit Function
    End If

    debug_ 1585, LOG_MAIN, "stageTwoInitialise", "stageTwoInitialise: done"

    stageTwoInitialise = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
' Free up after level specific data has been released but before base data is released
'
Public Function stageTwoShutDown() As Boolean
    debug_ 1596, LOG_WZ, "stageTwoShutDown", "== stageTwoShutDown =="

    shutdown3DView_FullReset

    fpathShutdown

    cdAudio_Stop

    freeAllStructs
    freeAllDroids
    freeAllFeatures
    freeAllFlagPositions

    If Not messageShutdown() Then
        stageTwoShutDown = False
        Exit Function
    End If

    If Not mechanicsShutdown() Then
        stageTwoShutDown = False
        Exit Function
    End If

    interfaceShutDown

    cmdDroidShutDown

    'free up the gateway stuff?
    gwShutDown

    If Not mapShutdown() Then
        stageTwoShutDown = False
        Exit Function
    End If

    clearCampaignName

    stageTwoShutDown = True
End Function

Private Sub displayLoadingErrors()
    Dim loadingErrors As StringVector, modelLoadingErrorCount As Long, modelTextureLoadingFailures As Long
    Dim statModelLoadingFailures As Long, modelErrorDescription As String, i As Long, mods() As WzMods_LoadedMod

    ' warn the user if a mod might have corrupted models / files
    modelLoadingErrorCount = getModelLoadingErrorCount()
    If modelLoadingErrorCount > 0 Then
        debug_ 1644, LOG_INFO, "displayLoadingErrors", "Failure to load %zu game model(s) - please remove any outdated or broken mods", modelLoadingErrorCount
        StringVector_push_back loadingErrors, va("Failure to load %zu game model(s)", modelLoadingErrorCount)
    End If
    modelTextureLoadingFailures = getModelTextureLoadingFailuresCount()
    If modelTextureLoadingFailures > 0 Then
        debug_ 1650, LOG_INFO, "displayLoadingErrors", "Failed to load textures for %zu models - please remove any outdated or broken mods", modelTextureLoadingFailures
        StringVector_push_back loadingErrors, va("Failure to load textures for %zu model(s)", modelTextureLoadingFailures)
    End If
    statModelLoadingFailures = getStatModelLoadingFailures()
    If statModelLoadingFailures > 0 Then
        debug_ 1656, LOG_INFO, "displayLoadingErrors", "Failed to load model for %zu stats entries - please remove any outdated or broken mods", statModelLoadingFailures
        StringVector_push_back loadingErrors, va("Failure to load model for %zu stat(s)", statModelLoadingFailures)
    End If

    If loadingErrors.size > 0 Then
        modelErrorDescription = "Warzone 2100 encountered error(s) loading game data:"
        For i = 0 To loadingErrors.size - 1
            modelErrorDescription = modelErrorDescription & vbLf
            modelErrorDescription = modelErrorDescription & "- "
            modelErrorDescription = modelErrorDescription & loadingErrors.items(i)
        Next
        modelErrorDescription = modelErrorDescription & vbLf & vbLf
        If getLoadedMods(mods) > 0 Then
            modelErrorDescription = modelErrorDescription & "Please try removing any new mods - they may have issues or be incompatible with this version."
            modelErrorDescription = modelErrorDescription & vbLf & vbLf
            modelErrorDescription = modelErrorDescription & "Loaded mod(s):"
            modelErrorDescription = modelErrorDescription & vbLf
            modelErrorDescription = modelErrorDescription & "- " & getModList()
        Else
            modelErrorDescription = modelErrorDescription & "Base game files may be corrupt or outdated - please try reinstalling the game."
        End If
        wzDisplayDialog Dialog_Error, "Error Loading Game Data", modelErrorDescription
    End If
End Sub

Public Function stageThreeInitialise() As Boolean
    Dim fromSave As Boolean, modelTilesetIdx As Long

    fromSave = (getSaveGameType() = GTYPE_SAVE_START Or getSaveGameType() = GTYPE_SAVE_MIDMISSION)

    debug_ 1690, LOG_WZ, "stageThreeInitialise", "== stageThreeInitialise =="

    loopMissionState = LMS_NORMAL

    ' preload model textures for current tileset
    modelTilesetIdx = currentMapTileset
    modelUpdateTilesetIdx modelTilesetIdx
    If Not headlessGameMode() Then
        ' (enumerateLoadedModels(... pDisplayShape->getTextures() for each display model ...): the models' textures,
        ' cluster D)
    End If
    resDoResLoadCallback       ' do callback.

    If Not InitRadar() Then  ' After resLoad cause it needs the game palette initialised.
        stageThreeInitialise = False
        Exit Function
    End If

    ' reset the clock to normal speed
    gameTimeResetMod

    If Not init3DView() Then ' Initialise 3d view stuff. After resLoad cause it needs the game palette initialised.
        stageThreeInitialise = False
        Exit Function
    End If

    effectResetUpdates
    initLighting 0, 0, mapWidth, mapHeight
    pie_InitLighting
    LightMap_reset mapWidth, mapHeight

    If fromSave Then
        ' these two lines are the biggest hack in the world.
        ' the reticule seems to get detached from 'reticuleup'
        ' this forces it back in sync...
        intRemoveReticule
        intAddReticule
    End If

    If bMultiPlayer Then
        multiGameInit
        initTemplates
    End If

    preProcessVisibility

    prepareScripts (getLevelLoadType() = GTYPE_SAVE_MIDMISSION Or getLevelLoadType() = GTYPE_SAVE_START)

    If Not fpathInitialise() Then
        stageThreeInitialise = False
        Exit Function
    End If

    mapInit
    gridReset

    'if mission screen is up, close it.
    If MissionResUp Then
        intRemoveMissionResultNoAnim
    End If

    ' Re-inititialise some static variables.

    bInTutorial = False
    rangeOnScreen = False

    If fromSave And ActivityManager_getCurrentGameMode() = ActivitySink_GameMode_CHALLENGE Then
        challengeActive = True
    End If

    ' add radar to interface screen, and resize
    intAddRadarWidget
    resizeRadar

    setAllPauseStates False

    countUpdate False

    If getLevelLoadType() = GTYPE_SAVE_MIDMISSION Or getLevelLoadType() = GTYPE_SAVE_START Then
        triggerEvent TRIGGER_GAME_LOADED
        wzapi_processScriptQueuedObjectRemovals
    Else
        initNoGoAreas

        If debugMappingsAllowed() Then
            Cheated = True
            triggerEventCheatMode True
            gInputManager_contexts_set InputContext_DEBUG_MISC, InputContext_State_ACTIVE
        End If

        triggerEvent TRIGGER_GAME_INIT
        wzapi_processScriptQueuedObjectRemovals
        playerBuiltHQ = structureExists(selectedPlayer, REF_HQ, True, False)
    End If

    ' Start / randomize in-game music
    cdAudio_PlayTrack SONG_INGAME

    ' always start off with a refresh of the groups UI data
    intGroupsChanged

    ' warn the user if a mod might have corrupted models / files
    displayLoadingErrors

    If bMultiPlayer Then
        cameraToHome selectedPlayer, False, fromSave
        playerResponding ' say howdy!
        If NetPlay_bComms And Not NETisReplay() Then
            multiStartScreenInit
        End If
    End If

    ' Call once again to update counts (if modified by earlier wzapi events)
    countUpdate False

    stageThreeInitialise = True
End Function

'/*****************************************************************************/
'/*      Shutdown before any data is released                                 */

Public Function stageThreeShutDown() As Boolean
    debug_ 1827, LOG_WZ, "stageThreeShutDown", "== stageThreeShutDown =="

    resetHostLaunch

    removeSpotters

    ' There is an asymmetry in scripts initialization and destruction, due
    ' the many different ways scripts get loaded.
    If Not shutdownScripts() Then
        stageThreeShutDown = False
        Exit Function
    End If

    specStatsViewShutdown

    challengesUp = False
    challengeActive = False
    isInGamePopupUp = False
    InGameOpUp = False
    bInTutorial = False

    shutdownTemplates

    ' make sure any button tips are gone.
    widgReset

    If Not ShutdownRadar() Then
        debug_ 1855, LOG_FATAL, "stageThreeShutDown", "Failed to shutdown radar"
    End If

    audio_StopAll

    If bMultiPlayer Then
        multiGameShutdown
    End If

    'call this here before mission data is released
    If Not missionShutDown() Then
        stageThreeShutDown = False
        Exit Function
    End If

    setScriptWinLoseVideo PLAY_NONE

    shutdown3DView

    stageThreeShutDown = True
End Function

' Reset the game between campaigns
Public Function campaignReset() As Boolean
    debug_ 1881, LOG_MAIN, "campaignReset", "campaignReset"
    gwShutDown
    mapShutdown
    shutdownTerrain
    ' when the terrain textures are reloaded we need to reset the radar
    ' otherwise it will end up as a terrain texture somehow
    ShutdownRadar
    InitRadar
    campaignReset = True
End Function

' Reset the game when loading a save game
Public Function saveGameReset() As Boolean
    debug_ 1895, LOG_MAIN, "saveGameReset", "saveGameReset"

    cdAudio_Stop

    freeAllStructs
    freeAllDroids
    freeAllFeatures
    freeAllFlagPositions
    initMission
    initTransporters

    'free up the gateway stuff?
    gwShutDown
    intResetScreen True

    If Not mapShutdown() Then
        saveGameReset = False
        Exit Function
    End If

    freeMessages

    saveGameReset = True
End Function

' --- Miscellaneous Initialisation stuff that really should be done each restart
Private Sub initMiscVars()
    Dim n As Long
    selectedPlayer = 0
    realSelectedPlayer = 0
    godMode = False

    radarOnScreen = True
    radarPermitted = True
    allowDesign = True
    includeRedundantDesigns = False
    enableConsoleDisplay True

    For n = 0 To MAX_PLAYERS - 1
        dbgInputManager_setPlayerWantsDebugMappings n, False
    Next
End Sub
