Attribute VB_Name = "modding_c"
' src/modding.cpp. WZ_PHYSFS_enumerateFiles's callback is a loop over the
' list PHYSFS_enumerateFiles returns. The mods' hashes (getModHashList,
' WzMods::LoadedMod::getHash) come from findHashOfFile, a stand-in that
' gives the zero hash until SHA-256 is ported. Missing: getModFilename.
Option Explicit

Private Const PATH_MAX As Long = 260            ' MAX_PATH on Windows
Private Const modlist_string_size As Long = 255 ' netplay.h: For a concatenated list of mods

' modding.h: struct WzMods::LoadedMod
Public Type WzMods_LoadedMod
    name_ As String
    filename As String
    fullRealPath As String
    fileHash As Sha256
End Type

Public global_mods As StringVector
Public campaign_mods As StringVector
Public multiplay_mods As StringVector

Public override_mods As StringVector
Public override_mod_list As String
Public use_override_mods As Boolean

Private loaded_mods() As WzMods_LoadedMod
Private loaded_mods_size As Long
Private mod_list As String
Private mod_names_list As StringVector
Private mod_hash_list() As Sha256       ' std::vector<Sha256>: mod_hash_list(0 .. mod_hash_list_size - 1)
Private mod_hash_list_size As Long

Private Function split_(ByVal str_ As String, ByVal sep As String) As StringVector
    Dim strs As StringVector, begin_ As Long, end_ As Long
    If Len(str_) = 0 Then
        split_ = strs
        Exit Function
    End If
    begin_ = 1
    end_ = InStr(str_, sep)
    Do While end_ > 0
        StringVector_push_back strs, Mid$(str_, begin_, end_ - begin_)
        begin_ = end_ + Len(sep)
        end_ = InStr(begin_, str_, sep)
    Loop
    StringVector_push_back strs, Mid$(str_, begin_)
    split_ = strs
End Function

Public Function convertToPlatformDependentPath(ByVal platformIndependentPath As String) As String
    Dim path As String
    path = platformIndependentPath

    ' platform-independent notation uses "/" as the path separator
    If PHYSFS_getDirSeparator() <> "/" Then
        ' replace all "/" with PHYSFS_getDirSeparator()
        path = Replace(path, "/", PHYSFS_getDirSeparator())
    End If

    convertToPlatformDependentPath = path
End Function

'!
' Tries to mount a list of directories, found in /basedir/subdir/<list>.
' \param basedir Base directory (in platform-dependent notation)
' \param subdir A subdirectory of basedir (in platform-independent notation - i.e. with "/" as the path separator)
' \param appendToPath Whether to append or prepend
' \param checkList List of directories to check. NULL means any (here: useCheckList = False).
Public Function addSubdirs(ByVal basedir As String, ByVal subdir As String, ByVal appendToPath As Boolean, checkList As StringVector, ByVal useCheckList As Boolean, ByVal addToModList As Boolean) As Long
    Dim numAddedMods As Long, subdir_platformDependent As String, tmpFullModRealPath As String
    Dim list() As String, n As Long, k As Long, i As String, filename As String, buf As String
    Dim log_level As Long, errorCode As Long, pErrStr As String, mountOrder As Long

    subdir_platformDependent = convertToPlatformDependentPath(subdir)
    n = PHYSFS_enumerateFiles(subdir, list)             ' WZ_PHYSFS_enumerateFiles
    For k = 0 To n - 1
        i = list(k)
        If Left$(i, 1) <> "." Then
            If Not useCheckList Or StringVector_contains(checkList, i) Then
                ' platform-dependent notation
                tmpFullModRealPath = basedir
                tmpFullModRealPath = tmpFullModRealPath & subdir_platformDependent
                tmpFullModRealPath = tmpFullModRealPath & PHYSFS_getDirSeparator()
                tmpFullModRealPath = tmpFullModRealPath & i
                If hasLoadedModRealPath(tmpFullModRealPath) Then
                    debug_ 112, LOG_INFO, "addSubdirs", "Already loaded: %s, skipping", tmpFullModRealPath
                    GoTo continue_
                End If
                If appendToPath Then mountOrder = PHYSFS_APPEND Else mountOrder = PHYSFS_PREPEND
                If PHYSFS_mount(tmpFullModRealPath, vbNullString, mountOrder) <> 0 Then ' platform-dependent notation
                    numAddedMods = numAddedMods + 1
                    If addToModList Then
                        filename = va("%s/%s", subdir, i) ' platform-independent notation
                        addLoadedMod i, filename, tmpFullModRealPath
                        buf = Left$(va("mod: %s", i), 255)
                        addDumpInfo buf
                    End If
                Else
                    ' failed to mount mod
                    log_level = LOG_WZ
                    errorCode = PHYSFS_getLastErrorCode()
                    Select Case errorCode
                    Case PHYSFS_ERR_CORRUPT
                        log_level = LOG_ERROR
                    Case Else
                        log_level = LOG_WZ
                    End Select
                    pErrStr = PHYSFS_getErrorByCode(errorCode)
                    If StrPtr(pErrStr) = 0 Then
                        pErrStr = "<unknown error>"
                    End If
                    debug_ 147, log_level, "addSubdirs", "Failed to load mod from path: %s, %s", tmpFullModRealPath, pErrStr
                End If
            End If
        End If
continue_:
    Next
    addSubdirs = numAddedMods
End Function

Public Sub removeSubdirs(ByVal basedir As String, ByVal subdir As String)
    Dim subdir_platformDependent As String, tmpstr As String, list() As String, n As Long, k As Long, i As String
    If StrPtr(basedir) = 0 Then
        ASSERT_FAILURE "basedir", "src/modding.cpp:158", "removeSubdirs", "basedir is null"
    End If
    If StrPtr(subdir) = 0 Then
        ASSERT_FAILURE "subdir", "src/modding.cpp:159", "removeSubdirs", "subdir is null"
    End If
    subdir_platformDependent = convertToPlatformDependentPath(subdir)
    n = PHYSFS_enumerateFiles(subdir, list)             ' WZ_PHYSFS_enumerateFiles
    For k = 0 To n - 1
        i = list(k)
        tmpstr = Left$(va("%s%s%s%s", basedir, subdir_platformDependent, PHYSFS_getDirSeparator(), i), PATH_MAX - 1)
        If PHYSFS_unmount(tmpstr) = 0 Then ' platform-dependent notation
            ' (DEBUG builds log the reason)
        End If
    Next
End Sub

Public Sub printSearchPath()
    Dim searchPath() As String, n As Long, i As Long
    debug_ 182, LOG_WZ, "printSearchPath", "Search paths:"
    n = PHYSFS_getSearchPath(searchPath)
    For i = 0 To n - 1
        debug_ 186, LOG_WZ, "printSearchPath", "    [%s]", searchPath(i)
    Next
End Sub

Public Sub setOverrideMods(ByVal modlist As String)
    override_mods = split_(modlist, ", ")
    override_mod_list = modlist
    use_override_mods = True
End Sub

Public Sub clearOverrideMods()
    StringVector_clear override_mods
    override_mod_list = ""
    use_override_mods = False
End Sub

Public Function hasOverrideMods() As Boolean
    hasOverrideMods = (override_mods.size <> 0)
End Function

Public Sub clearCampaignMods()
    StringVector_clear campaign_mods
End Sub

Public Function hasCampaignMods() As Boolean
    hasCampaignMods = (campaign_mods.size <> 0)
End Function

Private Sub addLoadedMod(ByVal modname As String, ByVal filename As String, ByVal fullRealPath As String)
    ' Note, findHashOfFile won't work right now, since the search paths aren't set up until after all calls to addSubdirs, see rebuildSearchPath in init.cpp.
    If loaded_mods_size = 0 Then
        ReDim loaded_mods(0 To 7)
    ElseIf loaded_mods_size > UBound(loaded_mods) Then
        ReDim Preserve loaded_mods(0 To 2 * loaded_mods_size - 1)
    End If
    loaded_mods(loaded_mods_size).name_ = modname
    loaded_mods(loaded_mods_size).filename = filename
    loaded_mods(loaded_mods_size).fullRealPath = fullRealPath
    Sha256_setZero loaded_mods(loaded_mods_size).fileHash
    loaded_mods_size = loaded_mods_size + 1
    mod_list = ""
    StringVector_clear mod_names_list
    Erase mod_hash_list
    mod_hash_list_size = 0
End Sub

Private Function hasLoadedModRealPath(ByVal fullRealPath As String) As Boolean
    Dim i As Long
    For i = 0 To loaded_mods_size - 1
        If loaded_mods(i).fullRealPath = fullRealPath Then
            hasLoadedModRealPath = True
            Exit Function
        End If
    Next
End Function

Public Sub clearLoadedMods()
    loaded_mods_size = 0
    Erase loaded_mods
    mod_list = ""
    StringVector_clear mod_names_list
    Erase mod_hash_list
    mod_hash_list_size = 0
End Sub

' C: std::vector<WzMods::LoadedMod> const &getLoadedMods(): the mods in out(0 ..), and their count
Public Function getLoadedMods(out() As WzMods_LoadedMod) As Long
    If loaded_mods_size > 0 Then out = loaded_mods
    getLoadedMods = loaded_mods_size
End Function

Public Function getModList() As String
    Dim mods As StringVector, i As Long
    If Len(mod_list) = 0 Then
        ' Construct mod list.
        For i = 0 To loaded_mods_size - 1
            StringVector_push_back mods, loaded_mods(i).name_
        Next
        StringVector_sort mods
        StringVector_unique mods
        mod_list = strJoin(mods, ", ")
        If Len(mod_list) > modlist_string_size - 1 Then mod_list = Left$(mod_list, modlist_string_size - 1) ' Probably not needed.
    End If
    getModList = mod_list
End Function

Public Function getModNamesList() As StringVector
    Dim i As Long
    If mod_names_list.size = 0 Then
        For i = 0 To loaded_mods_size - 1
            StringVector_push_back mod_names_list, loaded_mods(i).name_
        Next
        StringVector_sort mod_names_list
        StringVector_unique mod_names_list
    End If
    getModNamesList = mod_names_list
End Function

' Sha256& WzMods::LoadedMod::getHash()
Private Function LoadedMod_getHash(m As WzMods_LoadedMod) As Sha256
    If Sha256_isZero(m.fileHash) Then
        findHashOfFile m.filename, m.fileHash
    End If
    LoadedMod_getHash = m.fileHash
End Function

' C: std::vector<Sha256> const &getModHashList(): the hashes in out(0 ..), and their count
Public Function getModHashList(out() As Sha256) As Long
    Dim i As Long, hash As Sha256
    If mod_hash_list_size = 0 Then
        For i = 0 To loaded_mods_size - 1
            hash = LoadedMod_getHash(loaded_mods(i))
            debug_ 303, LOG_WZ, "getModHashList", "Mod[%s]: %s\n", Sha256_toString(hash), loaded_mods(i).filename
            ' mod_hash_list.push_back(hash)
            If mod_hash_list_size = 0 Then
                ReDim mod_hash_list(0 To 7)
            ElseIf mod_hash_list_size > UBound(mod_hash_list) Then
                ReDim Preserve mod_hash_list(0 To 2 * mod_hash_list_size - 1)
            End If
            mod_hash_list(mod_hash_list_size) = hash
            mod_hash_list_size = mod_hash_list_size + 1
        Next
    End If
    If mod_hash_list_size > 0 Then out = mod_hash_list
    getModHashList = mod_hash_list_size
End Function
