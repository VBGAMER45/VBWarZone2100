Attribute VB_Name = "cdaudio_c"
' lib/sound/cdaudio.cpp, cdaudio.h: the music: the menu's track, the playlist's (playlist_c) in a game,
' played as an AUDIO_STREAM (openal_track_c's sound_PlayStream; its onFinished: a CdAudioTrackFinished that
' holds the track). A std::shared_ptr<const WZ_TRACK> is a WZ_TRACK (Nothing = nullptr); the event sinks
' are ICDAudioEventSink objects (NOTIFY_MUSIC_EVENT and NOTIFY_MUSIC_EVENT_TRACK are notifyMusicEvent).
' The enums SONG_CONTEXT and MusicGameMode are SoundApi's; to_string(MusicGameMode) is
' MusicGameMode_to_string. The globals' initialisers run on the first call (InitGlobals). No music ships
' with the game data, and Opus (music/menu.opus, the albums' tracks) has no decoder here: sound_PlayStream
' fails as for a missing file.
Option Explicit

' (VB) the events of NOTIFY_MUSIC_EVENT and NOTIFY_MUSIC_EVENT_TRACK
Private Enum MUSIC_EVENT
    ME_startedPlayingTrack
    ME_trackEnded
    ME_musicStopped
    ME_musicPaused
    ME_musicResumed
End Enum

' MARK: - Globals

Private music_volume As Single              ' = 0.5
Private music_initialized As Boolean        ' = false
Private stopping As Boolean                 ' = true
Private is_opening_new_track As Boolean     ' = false
Private queued_play_track_while_loading As Boolean ' = false
Private cdStream As AUDIO_STREAM            ' = nullptr

Private Const MENU_MUSIC As String = "music/menu.opus"
Private currentSongContext As SONG_CONTEXT  ' = SONG_FRONTEND
Private currentTrack As WZ_TRACK
Private registeredEventSinks As New Collection ' std::vector<std::shared_ptr<CDAudioEventSink>>

Private globalsReady As Boolean

' (VB) the globals' initialisers
Private Sub InitGlobals()
    If globalsReady Then Exit Sub
    globalsReady = True
    music_volume = 0.5!
    stopping = True
    currentSongContext = SONG_FRONTEND
End Sub

' lib/framework/math_ext.h
Private Function clipf(ByVal x As Single, ByVal min_ As Single, ByVal max_ As Single) As Single
    If x < min_ Then
        clipf = min_
    ElseIf x > max_ Then
        clipf = max_
    Else
        clipf = x
    End If
End Function

' MARK: - Helpers

Public Function MusicGameMode_to_string(ByVal mode As MusicGameMode) As String
    Select Case mode
    Case MusicGameMode_MENUS
        MusicGameMode_to_string = "Menu"
    Case MusicGameMode_CAMPAIGN
        MusicGameMode_to_string = "Campaign"
    Case MusicGameMode_CHALLENGE
        MusicGameMode_to_string = "Challenge"
    Case MusicGameMode_SKIRMISH
        MusicGameMode_to_string = "Skirmish"
    Case MusicGameMode_MULTIPLAYER
        MusicGameMode_to_string = "Multiplayer"
    Case Else
        MusicGameMode_to_string = "" ' silence warning
    End Select
End Function

' MARK: - CDAudioEventSink

Private Sub EventSinkCleanup()
    Dim i As Long, a As ICDAudioEventSink
    If registeredEventSinks.Count = 0 Then Exit Sub
    i = 1
    Do While i <= registeredEventSinks.Count
        Set a = registeredEventSinks(i)
        If a Is Nothing Then
            registeredEventSinks.Remove i
        ElseIf a.unregisterEventSink() Then
            registeredEventSinks.Remove i
        Else
            i = i + 1
        End If
    Loop
End Sub

' NOTIFY_MUSIC_EVENT(event), NOTIFY_MUSIC_EVENT_TRACK(event, currenttrack)
Private Sub notifyMusicEvent(ByVal event_ As MUSIC_EVENT, Optional ByVal track As WZ_TRACK)
    Dim i As Long, sink As ICDAudioEventSink
    EventSinkCleanup
    For i = 1 To registeredEventSinks.Count
        Set sink = registeredEventSinks(i)
        Select Case event_
        Case ME_startedPlayingTrack: sink.startedPlayingTrack track
        Case ME_trackEnded: sink.trackEnded track
        Case ME_musicStopped: sink.musicStopped
        Case ME_musicPaused: sink.musicPaused track
        Case ME_musicResumed: sink.musicResumed track
        End Select
    Next
End Sub

' MARK: - Core cdAudio functions

Public Function cdAudio_Open(ByVal user_musicdir As String) As Boolean
    InitGlobals
    PlayList_Init

    If StrPtr(user_musicdir) = 0 Then
        cdAudio_Open = False
        Exit Function
    End If
    If Not PlayList_Read(user_musicdir) Then
        cdAudio_Open = False
        Exit Function
    End If

    debug_ 122, LOG_SOUND, "cdAudio_Open", "called(%s)", user_musicdir

    music_initialized = True
    stopping = True

    cdAudio_Open = True
End Function

Public Sub cdAudio_Close()
    InitGlobals
    debug_ 132, LOG_SOUND, "cdAudio_Close", "called"
    cdAudio_Stop
    PlayList_Quit

    music_initialized = False
    stopping = True
End Sub

Private Sub cdAudio_Stop_Internal(Optional ByVal suppressEvent As Boolean = False)
    stopping = True
    debug_ 143, LOG_SOUND, "cdAudio_Stop_Internal", "called, cdStream=%p", ObjPtr(cdStream)

    If Not cdStream Is Nothing Then
        sound_StopStream cdStream
        Set cdStream = Nothing
        Set currentTrack = Nothing
        sound_Update
        If Not suppressEvent Then
            notifyMusicEvent ME_musicStopped
        End If
    End If
End Sub

Private Function cdAudio_GetMenuTrack() As WZ_TRACK
    Dim pTrack As WZ_TRACK
    Set pTrack = New WZ_TRACK
    pTrack.filename = MENU_MUSIC
    pTrack.title = "Menu Music"
    pTrack.author = "Martin Severn"
    pTrack.base_volume = 100
    pTrack.bpm = 80
    Set cdAudio_GetMenuTrack = pTrack
End Function

Private Function cdAudio_CalculateTrackVolume(ByVal track As WZ_TRACK) As Single
    Dim track_volume As Single
    If track Is Nothing Then cdAudio_CalculateTrackVolume = music_volume: Exit Function
    If track.base_volume = 100 Then
        cdAudio_CalculateTrackVolume = music_volume
        Exit Function
    End If
    track_volume = music_volume * (CSng(track.base_volume) / CSng(100))
    ' Keep volume in the range of 0.0 - 1.0
    track_volume = clipf(track_volume, 0!, 1!)
    cdAudio_CalculateTrackVolume = track_volume
End Function

' (VB) the lambda [track](const AUDIO_STREAM*, const void*) { cdAudio_TrackFinished(track);}
Private Function CdAudioTrackFinished_new(ByVal track As WZ_TRACK) As IStreamFinished
    Dim f As CdAudioTrackFinished
    Set f = New CdAudioTrackFinished
    Set f.track = track
    Set CdAudioTrackFinished_new = f
End Function

Private Function cdAudio_OpenTrack(ByVal track As WZ_TRACK) As Boolean
    Dim filename As String, bufferEntireStream As Boolean
    filename = track.filename
    If Not music_initialized Then
        cdAudio_OpenTrack = False
        Exit Function
    End If

    bufferEntireStream = (currentSongContext = SONG_FRONTEND)

    debug_ 194, LOG_SOUND, "cdAudio_OpenTrack", "called(%s)", filename
    is_opening_new_track = True
    cdAudio_Stop_Internal True
    is_opening_new_track = False

    PlayList_SetCurrentSong track

    Set cdStream = sound_PlayStream(filename, bufferEntireStream, cdAudio_CalculateTrackVolume(track), _
                        CdAudioTrackFinished_new(track), Empty)
    If cdStream Is Nothing Then
        debug_ 205, LOG_ERROR, "cdAudio_OpenTrack", "Failed creating audio stream for %s", filename
        notifyMusicEvent ME_musicStopped
        cdAudio_OpenTrack = False
        Exit Function
    End If
    Set currentTrack = track
    notifyMusicEvent ME_startedPlayingTrack, track

    debug_ 212, LOG_SOUND, "cdAudio_OpenTrack", "successful(%s)", filename
    stopping = False
    cdAudio_OpenTrack = True
End Function

Private Function cdAudio_GetNextTrack() As WZ_TRACK
    Dim nextTrack As WZ_TRACK
    Select Case currentSongContext
    Case SONG_FRONTEND
        Set nextTrack = cdAudio_GetMenuTrack()

    Case SONG_INGAME
        Set nextTrack = PlayList_NextSong()
    End Select
    Set cdAudio_GetNextTrack = nextTrack
End Function

' (VB) Public: CdAudioTrackFinished calls it
Public Sub cdAudio_TrackFinished(ByVal track As WZ_TRACK)
    Dim nextTrack As WZ_TRACK
    InitGlobals
    notifyMusicEvent ME_trackEnded, track

    ' This pointer is now officially invalidated; so set it to NULL
    Set cdStream = Nothing
    Set currentTrack = Nothing

    If is_opening_new_track Then
        Exit Sub ' shortcut further processing
    End If
    Set nextTrack = cdAudio_GetNextTrack()

    If nextTrack Is Nothing Then
        If Not stopping Then
            debug_ 251, LOG_SOUND, "cdAudio_TrackFinished", "Out of playlist - was playing %s", track.filename
            notifyMusicEvent ME_musicStopped
        End If
        Exit Sub
    End If

    If Not stopping Then
        If cdAudio_OpenTrack(nextTrack) Then
            debug_ 259, LOG_SOUND, "cdAudio_TrackFinished", "Now playing %s (was playing %s)", nextTrack.filename, track.filename
        End If
    End If
End Sub

' context: a SONG_CONTEXT
Public Function cdAudio_PlayTrack(ByVal context As Long) As Boolean
    Dim nextTrack As WZ_TRACK
    InitGlobals
    debug_ 265, LOG_SOUND, "cdAudio_PlayTrack", "called(%d)", context
    currentSongContext = context

    Select Case context
    Case SONG_FRONTEND
        cdAudio_PlayTrack = cdAudio_OpenTrack(cdAudio_GetMenuTrack())
        Exit Function

    Case SONG_INGAME
        If PlayList_GetCurrentMusicMode() = MusicGameMode_MENUS Then
            ' Likely caused by loading a saved game, but we don't (yet) have the full game mode
            ' As a workaround, queue playing the track
            queued_play_track_while_loading = True
            cdAudio_PlayTrack = True
            Exit Function
        End If

        Set nextTrack = PlayList_RandomizeCurrentSong()

        If nextTrack Is Nothing Then
            cdAudio_Stop
            cdAudio_PlayTrack = False
            Exit Function
        End If

        cdAudio_PlayTrack = cdAudio_OpenTrack(nextTrack)
        Exit Function
    End Select

    ASSERT_FAILURE "!""Invalid songcontext""", "lib/sound/cdaudio.cpp:295", "cdAudio_PlayTrack", "Invalid song context specified for playing: %u", context

    cdAudio_PlayTrack = False
End Function

Public Function cdAudio_PlaySpecificTrack(ByVal track As WZ_TRACK) As Boolean
    InitGlobals
    If track Is Nothing Then cdAudio_PlaySpecificTrack = False: Exit Function
    cdAudio_PlaySpecificTrack = cdAudio_OpenTrack(track)
End Function

Public Function cdAudio_CurrentSongContext() As SONG_CONTEXT
    InitGlobals
    cdAudio_CurrentSongContext = currentSongContext
End Function

' mode: a MusicGameMode
Public Sub cdAudio_SetGameMode(ByVal mode As Long)
    Dim oldMode As Long, numTracks As Long
    InitGlobals
    oldMode = PlayList_GetCurrentMusicMode()
    If oldMode <> mode Then
        numTracks = PlayList_FilterByMusicMode(mode)
        If (numTracks = 0) And (mode <> MusicGameMode_MENUS) Then
            ' no music configured for this game mode...
            debug_ 320, LOG_WARNING, "cdAudio_SetGameMode", "No music configured for current game mode"
        End If
        If queued_play_track_while_loading Then
            ' Workaround for lack of proper game type before savegame is fully loaded
            queued_play_track_while_loading = False
            cdAudio_PlayTrack currentSongContext
        End If
    End If
End Sub

Public Sub cdAudio_NextTrack()
    InitGlobals
    cdAudio_OpenTrack cdAudio_GetNextTrack()
End Sub

Public Sub cdAudio_Stop()
    InitGlobals
    cdAudio_Stop_Internal
End Sub

Public Sub cdAudio_Pause()
    InitGlobals
    If Not cdStream Is Nothing Then
        sound_PauseStream cdStream
        notifyMusicEvent ME_musicPaused, currentTrack
    End If
End Sub

Public Sub cdAudio_Resume()
    InitGlobals
    If Not cdStream Is Nothing Then
        sound_ResumeStream cdStream
        notifyMusicEvent ME_musicResumed, currentTrack
    End If
End Sub

Public Function sound_GetMusicVolume() As Single
    InitGlobals
    sound_GetMusicVolume = music_volume
End Function

Public Sub sound_SetMusicVolume(ByVal volume As Single)
    InitGlobals
    ' Keep volume in the range of 0.0 - 1.0
    music_volume = clipf(volume, 0!, 1!)

    ' Change the volume of the current stream as well (if any)
    If Not cdStream Is Nothing Then
        sound_SetStreamVolume cdStream, cdAudio_CalculateTrackVolume(currentTrack)
    End If
End Sub

Public Function cdAudio_GetCurrentTrack() As WZ_TRACK
    InitGlobals
    Set cdAudio_GetCurrentTrack = currentTrack
End Function

Public Sub cdAudio_RegisterForEvents(ByVal musicEventSink As ICDAudioEventSink)
    InitGlobals
    If musicEventSink Is Nothing Then Exit Sub
    If musicEventSink.unregisterEventSink() Then Exit Sub
    registeredEventSinks.Add musicEventSink
End Sub
