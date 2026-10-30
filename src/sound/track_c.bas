Attribute VB_Name = "track_c"
' lib/sound/track.cpp, track.h: the loaded sound effects (g_apTrack: the TRACKs by ID number, their loop,
' volume and audible radius from the audio effect file) and the samples' start, stop and end. A TRACK * is
' a TRACK, an AUDIO_SAMPLE * an AUDIO_SAMPLE (Nothing = nullptr); std::vector<TRACK *> g_apTrack is an
' array of TRACKs and its size. The stopped callback is the game's number, called through ISoundGame
' (soundgame_vb's audio_CallBack), and so is droid.cpp's droidAudioTrackStopped. The samples themselves are
' played by openal_track_c (sound_Play2DSample, sound_Play3DSample, sound_StopSample, sound_FreeTrack).
Option Explicit

'*
'
' defines
Private Const MAX_TRACKS As Long = 65536

'*
'
' static global variables
' array of pointers to sound effects
Private g_apTrack() As TRACK            ' std::vector<TRACK *>: g_apTrack(0 .. g_apTrack_size - 1)
Private g_apTrack_size As Long

' number of tracks loaded
Private g_iCurTracks As Long            ' = 0

' flag set when system is active (for callbacks etc)
Private g_bSystemActive As Boolean      ' = false
Private g_pStopTrackCallback As Long    ' AUDIO_CALLBACK = nullptr: the game's number (0 = nullptr)

' (VB) g_apTrack.resize(n): the new slots are nullptr
Private Sub g_apTrack_resize(ByVal n As Long)
    If n = 0 Then
        Erase g_apTrack
    ElseIf g_apTrack_size = 0 Then
        ReDim g_apTrack(0 To n - 1)
    Else
        ReDim Preserve g_apTrack(0 To n - 1)
    End If
    g_apTrack_size = n
End Sub

' (VB) (TRACK *)resGetData("WAV", fileName): the resource's data; Nothing when it holds no TRACK (the sound
' was off when it loaded: nullptr)
Private Function trackOfResource(ByVal data As Variant) As TRACK
    If IsObject(data) Then Set trackOfResource = data
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_Init(ByVal hrtf As Long) As Boolean
    g_iCurTracks = 0
    If Not sound_InitLibrary(hrtf) Then
        debug_ 57, LOG_ERROR, "sound_Init", "Cannot init sound library"
        sound_Init = False
        Exit Function
    End If

    If ID_MAX_SOUND > 512 Then g_apTrack_resize ID_MAX_SOUND Else g_apTrack_resize 512 ' std::max<size_t>(ID_MAX_SOUND, 512)

    ' set system active flag for callbacks
    g_bSystemActive = True
    sound_Init = True
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_Shutdown() As Boolean
    ' set inactive flag to prevent callbacks coming after shutdown
    g_bSystemActive = False
    sound_ShutdownLibrary
    sound_Shutdown = True
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_GetSystemActive() As Boolean
    sound_GetSystemActive = g_bSystemActive
End Function

' Retrieves loaded audio files and retrieves a TRACK from them on which some values are set and returns their respective ID numbers
'  \param fileName the filename of the track
'  \param loop whether the track should be looped until explicitly stopped
'  \param volume the volume this track should be played on (range is 0-100)
'  \param audibleRadius the radius from the source of sound where it can be heard
'  \return a non-zero value representing the track id number when successful, zero when the file is not found or no more tracks can be loaded (i.e. the limit is reached)
Public Function sound_SetTrackVals(ByVal fileName As String, ByVal loop_ As Boolean, ByVal volume As Long, ByVal audibleRadius As Long) As Long
    Dim trackID As Long, psTrack As TRACK, s As String

    If StrPtr(fileName) = 0 Or Len(fileName) = 0 Then ' check for empty filename.  This is a non fatal error.
        If StrPtr(fileName) = 0 Then s = "a NULL pointer" Else s = "empty"
        debug_ 103, LOG_WARNING, "sound_SetTrackVals", "fileName is %s", s
        sound_SetTrackVals = 0
        Exit Function
    End If

    Set psTrack = trackOfResource(resGetData("WAV", fileName))
    If psTrack Is Nothing Then
        debug_ 110, LOG_WARNING, "sound_SetTrackVals", "track %s resource not found", fileName
        sound_SetTrackVals = 0
        Exit Function
    End If

    ' get pre-assigned ID or produce one
    trackID = audio_GetIDFromStr(fileName)
    If trackID = NO_SOUND Then
        ' No pre-assigned ID available, produce one
        trackID = sound_GetAvailableID()
        If trackID >= MAX_TRACKS Then
            If Not (trackID < MAX_TRACKS) Then ASSERT_FAILURE "trackID < ( 65536 )", "lib/sound/track.cpp:122", "sound_SetTrackVals", "sound_GetTrackID: unused track not found / available!"
            sound_SetTrackVals = 0
            Exit Function
        End If
    End If

    If Not (ULT(trackID, g_apTrack_size)) Then ASSERT_FAILURE "trackID < g_apTrack.size()", "lib/sound/track.cpp:127", "sound_SetTrackVals", "Invalid trackID: %u", trackID: sound_SetTrackVals = 0: Exit Function
    If Not g_apTrack(trackID) Is Nothing Then
        debug_ 130, LOG_SOUND, "sound_SetTrackVals", "sound_SetTrackVals: track %i already set (filename: ""%s"") - overriding", trackID, g_apTrack(trackID).fileName
        If Not (psTrack Is g_apTrack(trackID)) Then ASSERT_FAILURE "psTrack == g_apTrack[trackID]", "lib/sound/track.cpp:131", "sound_SetTrackVals", "Data pointer not equal?": sound_SetTrackVals = 0: Exit Function
    End If

    ' set track members
    psTrack.bLoop = loop_
    psTrack.iVol = volume
    psTrack.iAudibleRadius = audibleRadius

    ' RAII: Make sure to initialize all values (we don't want undefined values)!
    psTrack.iTime = 0
    psTrack.iTimeLastFinished = 0
    psTrack.iNumPlaying = 0

    ' set global
    Set g_apTrack(trackID) = psTrack

    ' increment counter for amount of loaded tracks
    g_iCurTracks = g_iCurTracks + 1

    sound_SetTrackVals = trackID
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_ReleaseTrack(ByVal psTrack As TRACK)
    Dim i As Long
    ' This is here to save CPU wasted by the loop below;
    ' Calling this function with psTrack = NULL is perfectly legal,
    ' with or without this check
    If psTrack Is Nothing Then
        Exit Sub
    End If

    ' Run through the list of tracks and set the pointer to the track which is to be released to NULL
    For i = 0 To g_apTrack_size - 1
        If g_apTrack(i) Is psTrack Then
            Set g_apTrack(i) = Nothing
        End If
    Next

    sound_FreeTrack psTrack
    ' (free(psTrack): the last reference goes with the caller's)
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_CheckAllUnloaded()
    Dim i As Long
    For i = 0 To g_apTrack_size - 1
        If Not (g_apTrack(i) Is Nothing) Then ASSERT_FAILURE "pTrack == nullptr", "lib/sound/track.cpp:188", "sound_CheckAllUnloaded", "A track is not unloaded yet (%s); check audio.json for duplicate IDs", g_apTrack(i).fileName
    Next
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_TrackLooped(ByVal iTrack As Long) As Boolean
    If Not sound_CheckTrack(iTrack) Then
        sound_TrackLooped = False
        Exit Function
    End If
    sound_TrackLooped = g_apTrack(iTrack).bLoop
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_GetNumPlaying(ByVal iTrack As Long) As Long
    If Not sound_CheckTrack(iTrack) Then
        sound_GetNumPlaying = 0
        Exit Function
    End If
    sound_GetNumPlaying = g_apTrack(iTrack).iNumPlaying
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_CheckTrack(ByVal iTrack As Long) As Boolean
    If iTrack < 0 Or iTrack > g_iCurTracks - 1 Or iTrack >= g_apTrack_size Then
        debug_ 226, LOG_SOUND, "sound_CheckTrack", "Track number %i outside max %i (size: %zu)\n", iTrack, g_iCurTracks, g_apTrack_size
        sound_CheckTrack = False
        Exit Function
    End If

    If g_apTrack(iTrack) Is Nothing Then
        debug_ 232, LOG_SOUND, "sound_CheckTrack", "Track %i NULL\n", iTrack
        sound_CheckTrack = False
        Exit Function
    End If

    sound_CheckTrack = True
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_GetTrackTime(ByVal iTrack As Long) As Long
    If Not sound_CheckTrack(iTrack) Then
        sound_GetTrackTime = 0
        Exit Function
    End If
    sound_GetTrackTime = g_apTrack(iTrack).iTime
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_GetTrackVolume(ByVal iTrack As Long) As Long
    If Not sound_CheckTrack(iTrack) Then
        sound_GetTrackVolume = 0
        Exit Function
    End If
    sound_GetTrackVolume = g_apTrack(iTrack).iVol
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_GetTrackAudibleRadius(ByVal iTrack As Long) As Long
    If Not sound_CheckTrack(iTrack) Then
        sound_GetTrackAudibleRadius = 0
        Exit Function
    End If
    sound_GetTrackAudibleRadius = g_apTrack(iTrack).iAudibleRadius
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
' (VB) vbNullString for nullptr
Public Function sound_GetTrackName(ByVal iTrack As Long) As String
    ' If we get an invalid track ID or there are
    ' currently no tracks loaded then return NULL
    If iTrack <= 0 Or iTrack >= g_apTrack_size Or iTrack = SAMPLE_NOT_FOUND Then
        sound_GetTrackName = vbNullString
        Exit Function
    End If
    If g_apTrack(iTrack) Is Nothing Then
        sound_GetTrackName = vbNullString
        Exit Function
    End If

    sound_GetTrackName = g_apTrack(iTrack).fileName
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_Play2DTrack(ByVal psSample As AUDIO_SAMPLE, ByVal bQueued As Boolean) As Boolean
    Dim psTrack As TRACK

    ' Check to make sure the requested track is loaded
    If Not sound_CheckTrack(psSample.iTrack) Then
        sound_Play2DTrack = False
        Exit Function
    End If

    Set psTrack = g_apTrack(psSample.iTrack)
    sound_Play2DTrack = sound_Play2DSample(psTrack, psSample, bQueued)
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_Play3DTrack(ByVal psSample As AUDIO_SAMPLE) As Boolean
    Dim psTrack As TRACK

    ' Check to make sure the requested track is loaded
    If Not sound_CheckTrack(psSample.iTrack) Then
        sound_Play3DTrack = False
        Exit Function
    End If

    Set psTrack = g_apTrack(psSample.iTrack)
    sound_Play3DTrack = sound_Play3DSample(psTrack, psSample)
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_StopTrack(ByVal psSample As AUDIO_SAMPLE)
    If Not (Not psSample Is Nothing) Then ASSERT_FAILURE "psSample != nullptr", "lib/sound/track.cpp:339", "sound_StopTrack", "Sample pointer invalid": Exit Sub

    sound_StopSample psSample

    ' do stopped callback
    If g_pStopTrackCallback <> 0 And psSample.psObj <> 0 Then
        audio_CallBack g_pStopTrackCallback, psSample.psObj
    End If
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_PauseTrack(ByVal psSample As AUDIO_SAMPLE)
    sound_StopSample psSample
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_FinishedCallback(ByVal psSample As AUDIO_SAMPLE)
    If Not (Not psSample Is Nothing) Then ASSERT_FAILURE "psSample != nullptr", "lib/sound/track.cpp:365", "sound_FinishedCallback", "sound_FinishedCallback: sample pointer invalid\n": Exit Sub

    ' Check to make sure the requested track is loaded
    If Not sound_CheckTrack(psSample.iTrack) Then
        Exit Sub
    End If

    If Not g_apTrack(psSample.iTrack) Is Nothing Then
        g_apTrack(psSample.iTrack).iTimeLastFinished = sound_GetGameTime()
    End If

    ' call user callback if specified
    If psSample.pCallback <> 0 Then
        audio_CallBack psSample.pCallback, psSample.psObj
        ' NOTE: we must invalidate the iAudioID (iTrack) so the game knows to add the
        ' sample back into the queue.  This is a bit of a hacky way to handle "looped" samples,
        ' since currently, the game don't have a decent way to handle this.
        If psSample.psObj <> 0 Then
            droidAudioTrackStopped psSample.psObj
        End If
    End If

    ' set finished flag
    psSample.bFinishedPlaying = True
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_GetTrackID(ByVal psTrack As TRACK) As Long
    Dim i As Long
    If psTrack Is Nothing Then
        sound_GetTrackID = SAMPLE_NOT_FOUND
        Exit Function
    End If

    ' find matching track
    For i = 0 To g_apTrack_size - 1
        If g_apTrack(i) Is psTrack Then
            sound_GetTrackID = i
            Exit Function
        End If
    Next

    ' no matching track found
    sound_GetTrackID = SAMPLE_NOT_FOUND
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_GetAvailableID() As Long
    Dim i As Long

    ' Iterate through the list of tracks until we find an unused ID slot
    For i = ID_SOUND_NEXT To g_apTrack_size - 1
        If g_apTrack(i) Is Nothing Then
            sound_GetAvailableID = i
            Exit Function
        End If
    Next                                ' (VB) i is where C's loop ends: g_apTrack.size(), or ID_SOUND_NEXT when it never ran

    ' add a new entry
    g_apTrack_resize g_apTrack_size + 1 ' g_apTrack.push_back(nullptr)
    sound_GetAvailableID = i
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_SetStoppedCallback(ByVal pStopTrackCallback As Long)
    g_pStopTrackCallback = pStopTrackCallback
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_GetTrackTimeLastFinished(ByVal iTrack As Long) As Long
    If Not sound_CheckTrack(iTrack) Then
        sound_GetTrackTimeLastFinished = 0
        Exit Function
    End If
    sound_GetTrackTimeLastFinished = g_apTrack(iTrack).iTimeLastFinished
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_SetTrackTimeLastFinished(ByVal iTrack As Long, ByVal iTime As Long)
    If Not sound_CheckTrack(iTrack) Then
        Exit Sub
    End If
    g_apTrack(iTrack).iTimeLastFinished = iTime
End Sub
