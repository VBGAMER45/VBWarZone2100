VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZ_Playlist_Preferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sound/playlist.cpp: class WZ_Playlist_Preferences, the user's music settings per track (music.json in
' the config folder). Made by playlist_c's WZ_Playlist_Preferences_new(fileName), which calls construct.
' optional<WZ_TRACK_SETTINGS> is a Boolean (has_value) and the settings ByRef; the template getValueForTrack
' is the copy for WZ_TRACK_SETTINGS. A std::exception is an error with Err.Source "nlohmann::json"
' (CONVENTIONS); typeid(WZ_TRACK_SETTINGS).name() is MSVC's.
Option Explicit

Private mRoot As nlohmann_json
Private mFilename As String

' (VB) json::parse(data, data + size) into mRoot: 0, or 1 for a std::exception (what: e.what()), 2 for another error
Private Function parseRoot(data() As Byte, ByVal size As Long, what As String) As Long
    On Error GoTo failed
    Set mRoot = json_parse_bytes(data, 0, size)
    Exit Function
failed:
    what = Err.Description
    If Err.Source = "nlohmann::json" Then parseRoot = 1 Else parseRoot = 2
End Function

' WZ_Playlist_Preferences(const std::string &fileName)
Friend Sub construct(ByVal fileName As String)
    Dim size As Long, data() As Byte, what As String, notificationsObject As nlohmann_json, isObj As Boolean

    mFilename = fileName
    Set mRoot = json_value(Empty)       ' nlohmann::json mRoot: null
    If PHYSFS_exists(fileName) <> 0 Then
        If loadFile(fileName, data, size) Then
            Select Case parseRoot(data, size, what)
            Case 1
                ASSERT_FAILURE "false", "lib/sound/playlist.cpp:155", "WZ_Playlist_Preferences", "JSON document from %s is invalid: %s", fileName, what
            Case 2
                debug_ 158, LOG_ERROR, "WZ_Playlist_Preferences", "Unexpected exception parsing JSON %s", fileName
            End Select
            If Not (Not mRoot.is_null()) Then ASSERT_FAILURE "!mRoot.is_null()", "lib/sound/playlist.cpp:160", "WZ_Playlist_Preferences", "JSON document from %s is null", fileName
            If Not (mRoot.is_object()) Then ASSERT_FAILURE "mRoot.is_object()", "lib/sound/playlist.cpp:161", "WZ_Playlist_Preferences", "JSON document from %s is not an object. Read: \n%s", fileName, utf8_decode(data, 0, size)
            ' (free(data))
        Else
            debug_ 166, LOG_ERROR, "WZ_Playlist_Preferences", "Could not open ""%s""", fileName
            ' treat as if no preferences exist yet
            Set mRoot = json_object()
        End If
    Else
        ' no preferences exist yet
        Set mRoot = json_object()
    End If

    ' always ensure there's a "notifications" dictionary in the root object
    Set notificationsObject = mRoot.find("tracks")
    If Not notificationsObject Is Nothing Then isObj = notificationsObject.is_object()
    If notificationsObject Is Nothing Or Not isObj Then
        ' create a dictionary object
        mRoot.set_item "tracks", json_object()
    End If
End Sub

Private Function keyForTrack(ByVal track As WZ_TRACK) As String
    Dim result As String
    ' remove the "music/albums/" prefix
    result = track.filename
    If Left$(result, 13) = "music/albums/" Then     ' result.rfind("music/albums/", 0) == 0
        result = Mid$(result, 14)                   ' result.substr(13)
    End If
    keyForTrack = result
End Function

Private Sub storeValueForTrack(ByVal track As WZ_TRACK, ByVal key As String, ByVal value As nlohmann_json)
    Dim trackSettingsObj As nlohmann_json, trackKey As String, trackData As nlohmann_json, isObj As Boolean
    Set trackSettingsObj = mRoot.item("tracks")
    trackKey = keyForTrack(track)
    Set trackData = trackSettingsObj.find(trackKey)
    If Not trackData Is Nothing Then isObj = trackData.is_object()
    If trackData Is Nothing Or Not isObj Then
        trackSettingsObj.set_item trackKey, json_object()
        Set trackData = trackSettingsObj.find(trackKey)
    End If
    trackData.set_item key, value
End Sub

' optional<nlohmann::json> getJSONValueForTrack(track, key): has_value returned, the value in result
Private Function getJSONValueForTrack(ByVal track As WZ_TRACK, ByVal key As String, result As nlohmann_json) As Boolean
    Dim n As Long, src As String, desc As String
    On Error GoTo catch_
    Set result = mRoot.at("tracks").at(keyForTrack(track)).at(key)
    getJSONValueForTrack = True
    Exit Function
catch_:
    n = Err.Number: src = Err.Source: desc = Err.Description
    If src = "nlohmann::json" And InStr(1, desc, "[json.exception.out_of_range.", vbBinaryCompare) = 1 Then
        ' some part of the path doesn't exist yet
        getJSONValueForTrack = False
        Exit Function
    End If
    Err.Raise n, src, desc              ' (not nlohmann::json::out_of_range: thrown on)
End Function

' template<typename T> optional<T> getValueForTrack(track, key), T = WZ_TRACK_SETTINGS
Private Function getValueForTrack_settings(ByVal track As WZ_TRACK, ByVal key As String, result As WZ_TRACK_SETTINGS) As Boolean
    Dim jsonResult As nlohmann_json, fresh As WZ_TRACK_SETTINGS
    On Error GoTo catch_
    If Not getJSONValueForTrack(track, key, jsonResult) Then
        getValueForTrack_settings = False
        Exit Function
    End If
    result = fresh                      ' jsonResult.value().get<T>(): a T made, then from_json
    WZ_TRACK_SETTINGS_from_json jsonResult, result
    getValueForTrack_settings = True
    Exit Function
catch_:
    If Err.Source = "nlohmann::json" Then
        debug_ 128, LOG_WARNING, "getValueForTrack", "Failed to convert json_variant to %s because of error: %s", "struct WZ_TRACK_SETTINGS", Err.Description
    Else
        debug_ 132, LOG_FATAL, "getValueForTrack", "Unexpected exception encountered: json_variant::toType<%s>", "struct WZ_TRACK_SETTINGS"
    End If
    getValueForTrack_settings = False
End Function

' optional<WZ_TRACK_SETTINGS> getTrackSettings(track): has_value returned, the settings in settings
Friend Function getTrackSettings(ByVal track As WZ_TRACK, settings As WZ_TRACK_SETTINGS) As Boolean
    getTrackSettings = getValueForTrack_settings(track, "settings", settings)
End Function

Friend Sub setTrackSettings(ByVal track As WZ_TRACK, settings As WZ_TRACK_SETTINGS)
    storeValueForTrack track, "settings", WZ_TRACK_SETTINGS_to_json(settings)
End Sub

Friend Sub clearAllPreferences()
    mRoot.set_item "tracks", json_object()
End Sub

Friend Function savePreferences() As Boolean
    Dim jsonString As String, b() As Byte, n As Long
    jsonString = mRoot.dump(4)
    n = utf8_encode(jsonString, b)      ' (VB) the bytes the file gets: std::string's size
    ' (#if SIZE_MAX >= UDWORD_MAX)
    If Not (U32(n) <= 4294967295#) Then ASSERT_FAILURE "jsonString.size() <= static_cast<size_t>(std::numeric_limits<UDWORD>::max())", "lib/sound/playlist.cpp:241", "savePreferences", "jsonString.size (%zu) exceeds UDWORD::max", n: savePreferences = False: Exit Function
    saveFile mFilename, b, n
    savePreferences = True
End Function
