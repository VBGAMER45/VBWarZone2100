Attribute VB_Name = "playlist_c"
' lib/sound/playlist.cpp, playlist.h: the music's albums (<path>/albums/<dir>/album.json) and tracks, the
' user's choice of tracks per game mode (music.json: WZ_Playlist_Preferences, a class here) and the current
' song. A std::shared_ptr<const WZ_TRACK> is a WZ_TRACK (Nothing = nullptr); albumList and fullTrackList
' are Collections; TRACK_SETTINGS_MAP (keyed by the track pointer) is two arrays searched with Is;
' currentFilterFunc (only PlayList_FilterByMusicMode's lambda sets one) is the mode it filters by;
' optional<size_t> currentSong is hasCurrentSong and a 0-based index. MusicGameMode is SoundApi's enum.
' The globals' initialisers run on the first call (InitGlobals).
Option Explicit

Private Const NUM_MUSICGAMEMODES As Long = 4

' MARK: - WZ_TRACK_SETTINGS

Public Type WZ_TRACK_SETTINGS
    music_modes(0 To NUM_MUSICGAMEMODES - 1) As Boolean ' = {false}
End Type

Private Type WZ_TRACK_SETTINGS_PAIR
    default_settings As WZ_TRACK_SETTINGS
    user_settings As WZ_TRACK_SETTINGS
End Type

' MARK: - Globals

Private albumList As New Collection             ' std::vector<std::shared_ptr<WZ_ALBUM>>
Private fullTrackList As New Collection         ' std::vector<std::shared_ptr<const WZ_TRACK>>
Private trackToSettingsMap_keys() As WZ_TRACK    ' TRACK_SETTINGS_MAP trackToSettingsMap: the keys,
Private trackToSettingsMap_values() As WZ_TRACK_SETTINGS_PAIR ' their values
Private trackToSettingsMap_size As Long
Private hasCurrentSong As Boolean               ' optional<size_t> currentSong = nullopt
Private currentSong As Long
Private lastFilteredMode As MusicGameMode       ' = MusicGameMode::MENUS
Private hasCurrentFilterFunc As Boolean         ' std::function currentFilterFunc: not empty,
Private currentFilterFunc_mode As MusicGameMode ' the mode its lambda captured
Private playlistPrefs As WZ_Playlist_Preferences ' = nullptr

Private globalsReady As Boolean

' (VB) the globals' initialisers
Private Sub InitGlobals()
    If globalsReady Then Exit Sub
    globalsReady = True
    lastFilteredMode = MusicGameMode_MENUS
End Sub

' WZ_TRACK_SETTINGS::setMusicMode
Private Sub WZ_TRACK_SETTINGS_setMusicMode(p As WZ_TRACK_SETTINGS, ByVal mode As MusicGameMode, ByVal enabled As Boolean)
    p.music_modes(mode) = enabled
End Sub

' inline void to_json(nlohmann::json& j, const WZ_TRACK_SETTINGS& p)
Public Function WZ_TRACK_SETTINGS_to_json(p As WZ_TRACK_SETTINGS) As nlohmann_json
    Dim j As nlohmann_json, music_modes As nlohmann_json, i As Long
    Set j = json_object()
    Set music_modes = json_array()
    For i = 0 To NUM_MUSICGAMEMODES - 1
        music_modes.set_item_idx i, json_value(p.music_modes(i))
    Next
    j.set_item "music_modes", music_modes
    Set WZ_TRACK_SETTINGS_to_json = j
End Function

' inline void from_json(const nlohmann::json& j, WZ_TRACK_SETTINGS& p): throws nlohmann::json::type_error
Public Sub WZ_TRACK_SETTINGS_from_json(ByVal j As nlohmann_json, p As WZ_TRACK_SETTINGS)
    Dim it As nlohmann_json, i As Long, k As Long
    If Not j.is_object() Then
        Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] type must be an object, but is " & j.type_name()
    End If
    Set it = j.find("music_modes")
    If Not it Is Nothing Then
        If Not it.is_array() Then
            Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] music_modes type must be an array, but is " & j.type_name()
        End If
        i = 0
        For k = 0 To it.size() - 1
            If i >= NUM_MUSICGAMEMODES Then
                Exit For
            End If
            p.music_modes(i) = it.value_at(k).get_bool()
            i = i + 1
        Next
    End If
End Sub

' MARK: - WZ_Playlist_Preferences

Private Function WZ_Playlist_Preferences_new(ByVal fileName As String) As WZ_Playlist_Preferences
    Dim p As WZ_Playlist_Preferences
    Set p = New WZ_Playlist_Preferences
    p.construct fileName
    Set WZ_Playlist_Preferences_new = p
End Function

' (VB) trackToSettingsMap.find(track): the index, or -1 for end()
Private Function settingsMapFind(ByVal track As WZ_TRACK) As Long
    Dim i As Long
    For i = 0 To trackToSettingsMap_size - 1
        If trackToSettingsMap_keys(i) Is track Then
            settingsMapFind = i
            Exit Function
        End If
    Next
    settingsMapFind = -1
End Function

' MARK: - Core PlayList Functions

Private Sub PlayList_ClearAll()
    Dim a As Long, t As Long, album As WZ_ALBUM, track As WZ_TRACK
    ' (VB) the tracks' weak_ptr album: dropped, so that each album and its tracks go
    For a = 1 To albumList.Count
        Set album = albumList(a)
        For t = 1 To album.tracks.Count
            Set track = album.tracks(t)
            Set track.album = Nothing
        Next
    Next
    Set albumList = New Collection
    Set fullTrackList = New Collection
    Erase trackToSettingsMap_keys          ' trackToSettingsMap.clear()
    Erase trackToSettingsMap_values
    trackToSettingsMap_size = 0
    hasCurrentFilterFunc = False            ' currentFilterFunc = nullptr
    hasCurrentSong = False                  ' currentSong = nullopt
    If Not playlistPrefs Is Nothing Then
        Set playlistPrefs = Nothing         ' delete playlistPrefs
    End If
End Sub

Public Sub PlayList_Init()
    InitGlobals
    PlayList_ClearAll
    Set playlistPrefs = WZ_Playlist_Preferences_new("music.json")
End Sub

Public Sub PlayList_Quit()
    InitGlobals
    If Not playlistPrefs Is Nothing Then
        playlistPrefs.savePreferences
    End If
    PlayList_ClearAll
End Sub

' mode: a MusicGameMode
Public Function PlayList_IsTrackEnabledForMusicMode(ByVal track As WZ_TRACK, ByVal mode As Long) As Boolean
    Dim it As Long
    InitGlobals
    it = settingsMapFind(track)
    If it = -1 Then
        PlayList_IsTrackEnabledForMusicMode = False
        Exit Function
    End If
    PlayList_IsTrackEnabledForMusicMode = trackToSettingsMap_values(it).user_settings.music_modes(mode)
End Function

' mode: a MusicGameMode
Public Sub PlayList_SetTrackMusicMode(ByVal track As WZ_TRACK, ByVal mode As Long, ByVal enabled As Boolean)
    Dim it As Long, oldValue As Boolean
    InitGlobals
    it = settingsMapFind(track)
    If it = -1 Then
        Exit Sub
    End If
    oldValue = trackToSettingsMap_values(it).user_settings.music_modes(mode)
    If oldValue <> enabled Then
        trackToSettingsMap_values(it).user_settings.music_modes(mode) = enabled
        If Not playlistPrefs Is Nothing Then
            playlistPrefs.setTrackSettings track, trackToSettingsMap_values(it).user_settings
        End If
    End If
End Sub

Private Sub PlayList_SetTrackDefaultSettings(ByVal track As WZ_TRACK, default_settings As WZ_TRACK_SETTINGS)
    Dim it As Long, user_settings As WZ_TRACK_SETTINGS, empty_ As WZ_TRACK_SETTINGS_PAIR
    it = settingsMapFind(track)
    If it = -1 Then
        ' trackToSettingsMap.insert(TRACK_SETTINGS_MAP::value_type(track, WZ_TRACK_SETTINGS_PAIR()))
        it = trackToSettingsMap_size
        ReDim Preserve trackToSettingsMap_keys(0 To it)
        ReDim Preserve trackToSettingsMap_values(0 To it)
        Set trackToSettingsMap_keys(it) = track
        trackToSettingsMap_values(it) = empty_
        trackToSettingsMap_size = it + 1
    End If
    trackToSettingsMap_values(it).default_settings = default_settings
    If playlistPrefs.getTrackSettings(track, user_settings) Then
        ' load the saved user settings
        trackToSettingsMap_values(it).user_settings = user_settings
    Else
        ' initialize the user settings to the defaults
        trackToSettingsMap_values(it).user_settings = default_settings
        playlistPrefs.setTrackSettings track, trackToSettingsMap_values(it).user_settings
    End If
End Sub

' const std::vector<std::shared_ptr<const WZ_TRACK>>& PlayList_GetFullTrackList(): (VB) the number of
' tracks, and each one by its index (0-based)
Public Function PlayList_GetFullTrackList_size() As Long
    InitGlobals
    PlayList_GetFullTrackList_size = fullTrackList.Count
End Function

Public Function PlayList_GetFullTrackList_at(ByVal i As Long) As WZ_TRACK
    InitGlobals
    Set PlayList_GetFullTrackList_at = fullTrackList(i + 1)
End Function

' size_t PlayList_SetTrackFilter(const std::function<bool (const std::shared_ptr<const WZ_TRACK>& track)>& filterFunc):
' (VB) the filter is PlayList_FilterByMusicMode's lambda, given its captured mode
Private Function PlayList_SetTrackFilter(ByVal filterFunc_mode As MusicGameMode) As Long
    Dim tracksEnabled As Long, a As Long, t As Long, album As WZ_ALBUM
    tracksEnabled = 0
    For a = 1 To albumList.Count
        Set album = albumList(a)
        For t = 1 To album.tracks.Count
            If PlayList_IsTrackEnabledForMusicMode(album.tracks(t), filterFunc_mode) Then   ' filterFunc(track)
                tracksEnabled = tracksEnabled + 1
            End If
        Next
    Next
    hasCurrentFilterFunc = True         ' currentFilterFunc = filterFunc
    currentFilterFunc_mode = filterFunc_mode
    PlayList_SetTrackFilter = tracksEnabled
End Function

' (VB) currentFilterFunc(track): PlayList_FilterByMusicMode's lambda
Private Function currentFilterFunc(ByVal track As WZ_TRACK) As Boolean
    currentFilterFunc = PlayList_IsTrackEnabledForMusicMode(track, currentFilterFunc_mode)
End Function

' currentMode: a MusicGameMode
Public Function PlayList_FilterByMusicMode(ByVal currentMode As Long) As Long
    Dim result As Long
    InitGlobals
    result = 0
    If currentMode <> MusicGameMode_MENUS Then
        ' [currentMode](const std::shared_ptr<const WZ_TRACK>& track) -> bool {
        '     return PlayList_IsTrackEnabledForMusicMode(track, currentMode); }
        result = PlayList_SetTrackFilter(currentMode)
    Else
        hasCurrentFilterFunc = False    ' currentFilterFunc = nullptr
    End If
    If lastFilteredMode <> currentMode Then
        hasCurrentSong = False          ' currentSong = nullopt
    End If
    lastFilteredMode = currentMode
    PlayList_FilterByMusicMode = result
End Function

Public Function PlayList_GetCurrentMusicMode() As MusicGameMode
    InitGlobals
    PlayList_GetCurrentMusicMode = lastFilteredMode
End Function

' (VB) json[#KEY].get<std::string>(): False (and e.what()) where it throws
Private Function jsonGetString(ByVal j As nlohmann_json, value As String, what As String) As Boolean
    On Error GoTo failed
    value = j.get_string()
    jsonGetString = True
    Exit Function
failed:
    what = Err.Description
End Function

' (VB) json[#KEY].get<unsigned int>(): False (and e.what()) where it throws
Private Function jsonGetUInt(ByVal j As nlohmann_json, value As Long, what As String) As Boolean
    On Error GoTo failed
    value = j.get_uint()
    jsonGetUInt = True
    Exit Function
failed:
    what = Err.Description
End Function

Private Function PlayList_LoadAlbum(ByVal json As nlohmann_json, ByVal sourcePath As String, ByVal albumDir As String) As WZ_ALBUM
    Dim album As WZ_ALBUM, track As WZ_TRACK, trackJson As nlohmann_json, k As Long, m As Long
    Dim s As String, u As Long, what As String, default_settings As WZ_TRACK_SETTINGS, fresh As WZ_TRACK_SETTINGS
    Dim musicModeJSON As nlohmann_json, musicMode As String, tracksJson As nlohmann_json, modesJson As nlohmann_json

    If Not (Not json.is_null()) Then ASSERT_FAILURE "!json.is_null()", "lib/sound/playlist.cpp:390", "PlayList_LoadAlbum", "JSON album from %s is null", sourcePath
    If Not (json.is_object()) Then ASSERT_FAILURE "json.is_object()", "lib/sound/playlist.cpp:391", "PlayList_LoadAlbum", "JSON album from %s is not an object", sourcePath

    Set album = New WZ_ALBUM            ' std::make_shared<WZ_ALBUM>()

    ' GET_REQUIRED_ALBUM_KEY_STRVAL(title)
    If Not json.contains("title") Then
        debug_ 408, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required `title` key", sourcePath
        Exit Function
    End If
    If Not jsonGetString(json.item("title"), s, what) Then
        debug_ 408, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get `title` key, error: %s", sourcePath, what
        Exit Function
    End If
    album.title = s
    ' GET_REQUIRED_ALBUM_KEY_STRVAL(author)
    If Not json.contains("author") Then
        debug_ 409, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required `author` key", sourcePath
        Exit Function
    End If
    If Not jsonGetString(json.item("author"), s, what) Then
        debug_ 409, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get `author` key, error: %s", sourcePath, what
        Exit Function
    End If
    album.author = s
    ' GET_REQUIRED_ALBUM_KEY_STRVAL(date)
    If Not json.contains("date") Then
        debug_ 410, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required `date` key", sourcePath
        Exit Function
    End If
    If Not jsonGetString(json.item("date"), s, what) Then
        debug_ 410, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get `date` key, error: %s", sourcePath, what
        Exit Function
    End If
    album.date_ = s
    ' GET_REQUIRED_ALBUM_KEY_STRVAL(description)
    If Not json.contains("description") Then
        debug_ 411, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required `description` key", sourcePath
        Exit Function
    End If
    If Not jsonGetString(json.item("description"), s, what) Then
        debug_ 411, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get `description` key, error: %s", sourcePath, what
        Exit Function
    End If
    album.description = s
    ' GET_REQUIRED_ALBUM_KEY_STRVAL(album_cover_filename)
    If Not json.contains("album_cover_filename") Then
        debug_ 412, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required `album_cover_filename` key", sourcePath
        Exit Function
    End If
    If Not jsonGetString(json.item("album_cover_filename"), s, what) Then
        debug_ 412, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get `album_cover_filename` key, error: %s", sourcePath, what
        Exit Function
    End If
    album.album_cover_filename = s
    If Len(album.album_cover_filename) <> 0 Then
        album.album_cover_filename = albumDir & "/" & album.album_cover_filename
    End If

    If Not json.contains("tracks") Then
        debug_ 420, LOG_ERROR, "PlayList_LoadAlbum", "Missing required `tracks` key"
        Exit Function
    End If
    Set tracksJson = json.item("tracks")
    If Not tracksJson.is_array() Then
        debug_ 425, LOG_ERROR, "PlayList_LoadAlbum", "Required key `tracks` should be an array"
        Exit Function
    End If

    For k = 0 To tracksJson.size() - 1
        Set trackJson = tracksJson.value_at(k)
        Set track = New WZ_TRACK        ' std::make_shared<WZ_TRACK>()
        ' GET_REQUIRED_TRACK_KEY(filename, std::string)
        If Not trackJson.contains("filename") Then
            debug_ 445, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required track key: `filename`", sourcePath
            Exit Function
        End If
        If Not jsonGetString(trackJson.item("filename"), s, what) Then
            debug_ 445, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get track key `filename`, error: %s", sourcePath, what
            Exit Function
        End If
        track.filename = s
        If Len(track.filename) = 0 Then
            ' Track must have a filename
            debug_ 449, LOG_ERROR, "PlayList_LoadAlbum", "%s: Empty or invalid filename for track", sourcePath
            Exit Function
        End If
        track.filename = albumDir & "/" & track.filename
        ' GET_REQUIRED_TRACK_KEY(title, std::string)
        If Not trackJson.contains("title") Then
            debug_ 453, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required track key: `title`", sourcePath
            Exit Function
        End If
        If Not jsonGetString(trackJson.item("title"), s, what) Then
            debug_ 453, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get track key `title`, error: %s", sourcePath, what
            Exit Function
        End If
        track.title = s
        ' GET_REQUIRED_TRACK_KEY(author, std::string)
        If Not trackJson.contains("author") Then
            debug_ 454, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required track key: `author`", sourcePath
            Exit Function
        End If
        If Not jsonGetString(trackJson.item("author"), s, what) Then
            debug_ 454, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get track key `author`, error: %s", sourcePath, what
            Exit Function
        End If
        track.author = s
        ' GET_REQUIRED_TRACK_KEY(base_volume, unsigned int)
        If Not trackJson.contains("base_volume") Then
            debug_ 455, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required track key: `base_volume`", sourcePath
            Exit Function
        End If
        If Not jsonGetUInt(trackJson.item("base_volume"), u, what) Then
            debug_ 455, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get track key `base_volume`, error: %s", sourcePath, what
            Exit Function
        End If
        track.base_volume = u
        ' GET_REQUIRED_TRACK_KEY(bpm, unsigned int)
        If Not trackJson.contains("bpm") Then
            debug_ 456, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required track key: `bpm`", sourcePath
            Exit Function
        End If
        If Not jsonGetUInt(trackJson.item("bpm"), u, what) Then
            debug_ 456, LOG_ERROR, "PlayList_LoadAlbum", "%s: Failed to get track key `bpm`, error: %s", sourcePath, what
            Exit Function
        End If
        track.bpm = u
        Set track.album = album         ' std::weak_ptr<WZ_ALBUM>(album)

        default_settings = fresh        ' WZ_TRACK_SETTINGS default_settings;
        ' default_music_modes
        If Not trackJson.contains("default_music_modes") Then
            debug_ 463, LOG_ERROR, "PlayList_LoadAlbum", "%s: Missing required track key: `default_music_modes`", sourcePath
            Exit Function
        End If
        Set modesJson = trackJson.item("default_music_modes")
        If Not modesJson.is_array() Then
            debug_ 468, LOG_ERROR, "PlayList_LoadAlbum", "%s: Required track key `default_music_modes` should be an array", sourcePath
            Exit Function
        End If
        For m = 0 To modesJson.size() - 1
            Set musicModeJSON = modesJson.value_at(m)
            If Not musicModeJSON.is_string() Then
                debug_ 475, LOG_ERROR, "PlayList_LoadAlbum", "%s: `default_music_modes` array item should be a string", sourcePath
                GoTo continue_
            End If
            musicMode = musicModeJSON.get_string()
            If musicMode = "campaign" Then
                WZ_TRACK_SETTINGS_setMusicMode default_settings, MusicGameMode_CAMPAIGN, True
            ElseIf musicMode = "challenge" Then
                WZ_TRACK_SETTINGS_setMusicMode default_settings, MusicGameMode_CHALLENGE, True
            ElseIf musicMode = "skirmish" Then
                WZ_TRACK_SETTINGS_setMusicMode default_settings, MusicGameMode_SKIRMISH, True
            ElseIf musicMode = "multiplayer" Then
                WZ_TRACK_SETTINGS_setMusicMode default_settings, MusicGameMode_MULTIPLAYER, True
            Else
                debug_ 497, LOG_WARNING, "PlayList_LoadAlbum", "%s: `default_music_modes` array item is unknown value: %s", sourcePath, musicMode
                GoTo continue_
            End If
continue_:
        Next
        PlayList_SetTrackDefaultSettings track, default_settings
        album.tracks.Add track          ' album->tracks.push_back(std::move(track))
    Next

    Set PlayList_LoadAlbum = album
End Function

' (VB) json::parse(data, data + size): 0, or 1 for a std::exception (what: e.what()), 2 for another error
Private Function parseAlbumJson(data() As Byte, ByVal size As Long, tmpJson As nlohmann_json, what As String) As Long
    On Error GoTo failed
    Set tmpJson = json_parse_bytes(data, 0, size)
    Exit Function
failed:
    what = Err.Description
    If Err.Source = "nlohmann::json" Then parseAlbumJson = 1 Else parseAlbumJson = 2
End Function

Public Function PlayList_Read(ByVal path As String) As Boolean
    Dim size As Long, data() As Byte, albumsPath As String, folders() As String, nFolders As Long, f As Long
    Dim albumDir As String, str_ As String, tmpJson As nlohmann_json, album As WZ_ALBUM, what As String
    Dim sorted As Collection, a As Long, t As Long
    InitGlobals
    size = 0
    albumsPath = path & "/albums"       ' astringf("%s/albums", path)

    ' WZ_PHYSFS_enumerateFolders(albumsPath.c_str(), [&](const char *i) -> bool { ... })
    WZ_PHYSFS_enumerateFolders albumsPath, folders, nFolders
    For f = 0 To nFolders - 1
        albumDir = albumsPath & "/" & folders(f)
        str_ = albumDir & "/album.json"
        If PHYSFS_exists(str_) = 0 Then
            GoTo continue_ ' continue;
        End If
        If Not loadFile(str_, data, size) Then
            debug_ 523, LOG_ERROR, "PlayList_Read", "album JSON file ""%s"" could not be opened!", str_
            GoTo continue_ ' continue;
        End If
        Set tmpJson = json_value(Empty)  ' nlohmann::json tmpJson
        Select Case parseAlbumJson(data, size, tmpJson, what)
        Case 1
            debug_ 531, LOG_ERROR, "PlayList_Read", "album JSON file %s is invalid: %s", str_, what
            GoTo continue_ ' continue;
        Case 2
            debug_ 535, LOG_FATAL, "PlayList_Read", "Unexpected exception parsing album JSON from %s", str_
        End Select
        If Not (Not tmpJson.is_null()) Then ASSERT_FAILURE "!tmpJson.is_null()", "lib/sound/playlist.cpp:537", "PlayList_Read", "JSON album from %s is null", str_
        If Not (tmpJson.is_object()) Then ASSERT_FAILURE "tmpJson.is_object()", "lib/sound/playlist.cpp:538", "PlayList_Read", "JSON album from %s is not an object. Read: \n%s", str_, utf8_decode(data, 0, size)
        ' (free(data))

        ' load album data
        Set album = PlayList_LoadAlbum(tmpJson, str_, albumDir)
        If album Is Nothing Then
            debug_ 545, LOG_ERROR, "PlayList_Read", "Failed to load album JSON: %s", str_
            GoTo continue_ ' continue;
        End If
        albumList.Add album             ' albumList.push_back(std::move(album))
continue_:
    Next

    ' Ensure that "Warzone 2100 OST" is always the first album in the list
    ' (std::stable_sort with a < b when a's title is the OST's: those first, each part in its order)
    Set sorted = New Collection
    For a = 1 To albumList.Count
        Set album = albumList(a)
        If album.title = "Warzone 2100 OST" Then sorted.Add album
    Next
    For a = 1 To albumList.Count
        Set album = albumList(a)
        If album.title <> "Warzone 2100 OST" Then sorted.Add album
    Next
    Set albumList = sorted

    ' generate full track list
    Set fullTrackList = New Collection
    For a = 1 To albumList.Count
        Set album = albumList(a)
        For t = 1 To album.tracks.Count
            fullTrackList.Add album.tracks(t)
        Next
    Next
    hasCurrentSong = False              ' currentSong = nullopt

    PlayList_Read = albumList.Count <> 0
End Function

' optional<size_t> PlayList_FindNextMatchingTrack(size_t startingSongIdx): has_value returned, the index in idxOut
Private Function PlayList_FindNextMatchingTrack(ByVal startingSongIdx As Long, idxOut As Long) As Boolean
    Dim idx As Long, size As Long
    If Not hasCurrentFilterFunc Then PlayList_FindNextMatchingTrack = False: Exit Function
    size = fullTrackList.Count
    If size = 0 Then PlayList_FindNextMatchingTrack = False: Exit Function
    If ULT(startingSongIdx, size - 1) Then idx = startingSongIdx + 1 Else idx = 0
    Do While idx <> startingSongIdx
        If currentFilterFunc(fullTrackList(idx + 1)) Then
            idxOut = idx
            PlayList_FindNextMatchingTrack = True
            Exit Function
        End If
        If idx < size - 1 Then idx = idx + 1 Else idx = 0
    Loop
    PlayList_FindNextMatchingTrack = False
End Function

Public Function PlayList_CurrentSong() As WZ_TRACK
    InitGlobals
    If fullTrackList.Count = 0 Then
        Set PlayList_CurrentSong = Nothing
        Exit Function
    End If
    If Not hasCurrentSong Or (currentSong >= fullTrackList.Count) Then
        hasCurrentSong = PlayList_FindNextMatchingTrack(fullTrackList.Count - 1, currentSong)
        If Not hasCurrentSong Then
            Set PlayList_CurrentSong = Nothing
            Exit Function
        End If
    End If
    Set PlayList_CurrentSong = fullTrackList(currentSong + 1)
End Function

Public Function PlayList_NextSong() As WZ_TRACK
    Dim currentSongIdx As Long
    InitGlobals
    If hasCurrentSong Then currentSongIdx = currentSong Else currentSongIdx = fullTrackList.Count - 1
    hasCurrentSong = PlayList_FindNextMatchingTrack(currentSongIdx, currentSong)
    Set PlayList_NextSong = PlayList_CurrentSong()
End Function

Public Function PlayList_RandomizeCurrentSong() As WZ_TRACK
    Dim incrementByNum As Long, currentSongIdx As Long, i As Long
    InitGlobals
    If fullTrackList.Count = 0 Then
        Set PlayList_RandomizeCurrentSong = Nothing
        Exit Function
    End If
    incrementByNum = rand() Mod fullTrackList.Count  ' ((size_t)rand() % fullTrackList.size())
    If hasCurrentSong Then currentSongIdx = currentSong Else currentSongIdx = fullTrackList.Count - 1
    For i = 0 To incrementByNum - 1
        hasCurrentSong = PlayList_FindNextMatchingTrack(currentSongIdx, currentSong)
        If Not hasCurrentSong Then
            Exit For
        End If
        currentSongIdx = currentSong
    Next
    Set PlayList_RandomizeCurrentSong = PlayList_CurrentSong()
End Function

Public Function PlayList_SetCurrentSong(ByVal track As WZ_TRACK) As Boolean
    Dim i As Long, hasTrackIdx As Boolean, trackIdx As Long
    InitGlobals
    If track Is Nothing Then PlayList_SetCurrentSong = False: Exit Function
    hasTrackIdx = False                 ' optional<size_t> trackIdx = nullopt
    For i = 0 To fullTrackList.Count - 1
        If track Is fullTrackList(i + 1) Then
            hasTrackIdx = True
            trackIdx = i
            Exit For
        End If
    Next
    hasCurrentSong = hasTrackIdx        ' currentSong = trackIdx
    currentSong = trackIdx
    PlayList_SetCurrentSong = hasTrackIdx
End Function
