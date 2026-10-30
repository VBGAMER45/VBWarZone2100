Attribute VB_Name = "audio_c"
' lib/sound/audio.cpp, audio.h: the sound effects the game plays (2D; 3D at a place or on an object; queued),
' the sample queue and the list of the samples playing, the audio effect file (loadAudioEffectFileData).
' The two std::list<AUDIO_SAMPLE *> are Collections (the front is item 1); free(psSample) is the sample
' dropped from its list. mutating_list_iterate runs over a copy of the list's items (the body may remove
' the current one). aud.h's functions, realTime and sound_GetGameTime go to the game through ISoundGame
' (soundgame_vb). WZAudioDataResourceInterface and WzAudioDataPhysFSHandle (oggopus.cpp's audio data, not
' ported) are left out. A UDWORD time difference is udiff's (mod 2^32).
Option Explicit

' defines
Private Const NO_SAMPLE As Long = -2
Private Const MAX_SAME_SAMPLES As Long = 2

Private Const GAME_TICKS_PER_SEC As Long = 1000         ' lib/gamelib/gtime.h
Private Const ATTENUATION_FACTOR As Single = 0.0003!    ' track.h
Private Const REPEAT_AUDIO_CUE_THRESHOLD As Long = 100

' (VB) the fields of the static AUDIO_SAMPLE g_sPreviousSample that audio.cpp uses (zeroed, as a static)
Private Type PREVIOUS_SAMPLE
    iTrack As Long
    x As Long
    y As Long
    z As Long
End Type

' global variables
Private g_psSampleList As New Collection    ' std::list<AUDIO_SAMPLE *>
Private g_psSampleQueue As New Collection   ' std::list<AUDIO_SAMPLE *>
Private g_bAudioEnabled As Boolean          ' = false
Private g_bAudioPaused As Boolean           ' = false
Private g_sPreviousSample As PREVIOUS_SAMPLE
Private g_iPreviousSampleTime As Long       ' = 0

Private lastTimeBuildFailedPlayed As Long   ' static UDWORD = 0 (before audio_PlayBuildFailedOnce)

' (VB) a - b of two UDWORDs, wrapped as C's unsigned subtraction (0 .. 2^32 - 1)
Private Function udiff(ByVal a As Long, ByVal b As Long) As Double
    udiff = U32(a) - U32(b)
    If udiff < 0# Then udiff = udiff + 4294967296#
End Function

' (VB) an int32 from a sum that may overflow it, wrapped as x86 does
Private Function wrap32(ByVal d As Double) As Long
    d = d - Int(d / 4294967296#) * 4294967296#
    wrap32 = FromU32(d)
End Function

' (VB) the items of a sample list, copied (mutating_list_iterate's iteration); the number of them
Private Function listItems(ByVal list As Collection, items() As AUDIO_SAMPLE) As Long
    Dim i As Long, n As Long
    n = list.Count
    If n > 0 Then
        ReDim items(1 To n)
        For i = 1 To n
            Set items(i) = list(i)
        Next
    End If
    listItems = n
End Function

Public Function loadAudioEffectFileData(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, i As Long, array_ As nlohmann_json, a As nlohmann_json, k As Long
    Dim fname As String, loop_ As Boolean, volume As Long, range_ As Long

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "lib/sound/audio.cpp:70", "loadAudioEffectFileData", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        Set array_ = ini.json("data")
        If array_.is_null() Then
            GoTo continue_              ' (C++ continues without ini.endGroup())
        End If
        If Not (array_.is_array()) Then ASSERT_FAILURE "array.is_array()", "lib/sound/audio.cpp:80", "loadAudioEffectFileData", "data is not an array"
        For k = 0 To array_.size() - 1
            Set a = array_.value_at(k)
            fname = a.item("fileName").get_string()
            loop_ = a.item("loop").get_bool()
            volume = a.item("volume").get_uint()
            range_ = a.item("range").get_uint()
            audio_SetTrackVals fname, loop_, volume, range_
        Next
        ini.endGroup
continue_:
    Next
    loadAudioEffectFileData = True
End Function

' Counts the number of samples in the SampleQueue
'  \return the number of samples in the SampleQueue
Public Function audio_GetSampleQueueCount() As Long
    audio_GetSampleQueueCount = g_psSampleQueue.Count
End Function

' Counts the number of samples in the SampleList
'  \return the number of samples in the SampleList
Public Function audio_GetSampleListCount() As Long
    audio_GetSampleListCount = g_psSampleList.Count
End Function
'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function audio_Disabled() As Boolean
    audio_Disabled = Not g_bAudioEnabled
End Function

Public Function audio_Paused() As Boolean
    audio_Paused = g_bAudioPaused
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
' pStopTrackCallback: the game's number for its AUDIO_CALLBACK; hrtf: an HRTFMode
Public Function audio_Init(ByVal pStopTrackCallback As Long, ByVal hrtf As Long, ByVal really_enable As Boolean) As Boolean
    ' init audio system
    g_sPreviousSample.iTrack = NO_SAMPLE
    g_sPreviousSample.x = SAMPLE_COORD_INVALID
    g_sPreviousSample.y = SAMPLE_COORD_INVALID
    g_sPreviousSample.z = SAMPLE_COORD_INVALID
    g_bAudioEnabled = really_enable
    If g_bAudioEnabled Then
        g_bAudioEnabled = sound_Init(hrtf)
    End If
    If g_bAudioEnabled Then
        sound_SetStoppedCallback pStopTrackCallback
    End If
    audio_Init = g_bAudioEnabled
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function audio_Shutdown() As Boolean
    Dim bOK As Boolean

    ' if audio not enabled return true to carry on game without audio
    If audio_Disabled() Then
        audio_Shutdown = True
        Exit Function
    End If

    bOK = sound_Shutdown()

    ' empty sample list
    ' empty sample queue
    ' free sample heap
    Set g_psSampleList = New Collection
    Set g_psSampleQueue = New Collection

    audio_Shutdown = bOK
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function audio_GetPreviousQueueTrackPos(iX As Long, iY As Long, iZ As Long) As Boolean
    If g_sPreviousSample.x = SAMPLE_COORD_INVALID _
        Or g_sPreviousSample.y = SAMPLE_COORD_INVALID _
        Or g_sPreviousSample.z = SAMPLE_COORD_INVALID Then
        audio_GetPreviousQueueTrackPos = False
        Exit Function
    End If

    iX = g_sPreviousSample.x
    iY = g_sPreviousSample.y
    iZ = g_sPreviousSample.z

    audio_GetPreviousQueueTrackPos = True
End Function

Public Function audio_GetPreviousQueueTrackRadarBlipPos(iX As Long, iY As Long) As Boolean
    If g_sPreviousSample.x = SAMPLE_COORD_INVALID _
        Or g_sPreviousSample.y = SAMPLE_COORD_INVALID Then
        audio_GetPreviousQueueTrackRadarBlipPos = False
        Exit Function
    End If

    If g_sPreviousSample.iTrack <> ID_SOUND_STRUCTURE_UNDER_ATTACK And g_sPreviousSample.iTrack <> ID_SOUND_UNIT_UNDER_ATTACK And _
        g_sPreviousSample.iTrack <> ID_SOUND_LASER_SATELLITE_FIRING And g_sPreviousSample.iTrack <> ID_SOUND_INCOMING_LASER_SAT_STRIKE Then
        audio_GetPreviousQueueTrackRadarBlipPos = False
        Exit Function
    End If

    ' realTime > g_iPreviousSampleTime + 5 * GAME_TICKS_PER_SEC: the int converts to unsigned
    If ULT(g_iPreviousSampleTime + 5 * GAME_TICKS_PER_SEC, realTime()) Then
        audio_GetPreviousQueueTrackRadarBlipPos = False
        Exit Function
    End If

    iX = g_sPreviousSample.x
    iY = g_sPreviousSample.y

    audio_GetPreviousQueueTrackRadarBlipPos = True
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Private Sub audio_AddSampleToHead(ByVal ppsSampleList As Collection, ByVal psSample As AUDIO_SAMPLE)
    If ppsSampleList.Count = 0 Then ppsSampleList.Add psSample Else ppsSampleList.Add psSample, , 1 ' emplace_front
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Private Sub audio_AddSampleToTail(ByVal ppsSampleList As Collection, ByVal psSample As AUDIO_SAMPLE)
    ppsSampleList.Add psSample          ' emplace_back
End Sub

'*
'
' audio_RemoveSample Removes sample from list but doesn't free its memory

'*
' =======================================================================================================================
' =======================================================================================================================
'
Private Sub audio_RemoveSample(ByVal ppsSampleList As Collection, ByVal psSample As AUDIO_SAMPLE)
    Dim i As Long
    If psSample Is Nothing Then
        Exit Sub
    End If

    For i = 1 To ppsSampleList.Count    ' std::find
        If ppsSampleList(i) Is psSample Then Exit For
    Next
    ' (VB) erase(end()) is undefined: nothing is removed
    If Not (i <= ppsSampleList.Count) Then ASSERT_FAILURE "it != ppsSampleList.end()", "lib/sound/audio.cpp:262", "audio_RemoveSample", "audio_RemoveSample: sample not found in list": Exit Sub
    ppsSampleList.Remove i
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Private Function audio_CheckTooManySameQueueTracksPlaying(ByVal iTrack As Long) As Boolean
    Dim iCount As Long, isTooMany As Boolean, psSample As AUDIO_SAMPLE, i As Long

    ' return if audio not enabled
    If audio_Disabled() Or audio_Paused() Then
        audio_CheckTooManySameQueueTracksPlaying = True
        Exit Function
    End If

    ' loop through queue sounds and check whether too many already in it
    For i = 1 To g_psSampleQueue.Count
        Set psSample = g_psSampleQueue(i)
        If psSample.iTrack = iTrack Then
            iCount = iCount + 1
        End If

        If iCount > MAX_SAME_SAMPLES Then
            isTooMany = True
            Exit For
        End If
    Next

    audio_CheckTooManySameQueueTracksPlaying = isTooMany
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Private Function audio_QueueSample(ByVal iTrack As Long) As AUDIO_SAMPLE
    Dim psSample As AUDIO_SAMPLE

    ' return if audio not enabled
    If audio_Disabled() Or audio_Paused() Then
        Exit Function
    End If

    If Not (sound_CheckTrack(iTrack) = True) Then ASSERT_FAILURE "sound_CheckTrack(iTrack) == true", "lib/sound/audio.cpp:317", "audio_QueueSample", "audio_QueueSample: track %i outside limits\n", iTrack

    ' reject track if too many of same ID already in queue
    If audio_CheckTooManySameQueueTracksPlaying(iTrack) Then
        Exit Function
    End If

    Set psSample = New AUDIO_SAMPLE     ' calloc(1, sizeof(AUDIO_SAMPLE)) (VB: no out of memory)

    psSample.iTrack = iTrack
    psSample.x = SAMPLE_COORD_INVALID
    psSample.y = SAMPLE_COORD_INVALID
    psSample.z = SAMPLE_COORD_INVALID
    psSample.bFinishedPlaying = False

    ' add to queue
    audio_AddSampleToTail g_psSampleQueue, psSample

    Set audio_QueueSample = psSample
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub audio_QueueTrack(ByVal iTrack As Long)
    ' return if audio not enabled
    If audio_Disabled() Or audio_Paused() Then
        Exit Sub
    End If

    audio_QueueSample iTrack
End Sub

'*
'
'
' * audio_QueueTrackMinDelay Will only play track if iMinDelay has elapsed since
' * track last finished
'

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub audio_QueueTrackMinDelay(ByVal iTrack As Long, ByVal iMinDelay As Long)
    Dim psSample As AUDIO_SAMPLE, iDelay As Double

    ' return if audio not enabled
    If audio_Disabled() Or audio_Paused() Then
        Exit Sub
    End If

    ' Determine if at least iMinDelay time has passed since the last time this track was played
    iDelay = udiff(sound_GetGameTime(), sound_GetTrackTimeLastFinished(iTrack))
    If iDelay < U32(iMinDelay) Then
        Exit Sub
    End If

    ' Construct an audio sample from requested track
    Set psSample = audio_QueueSample(iTrack)
    If psSample Is Nothing Then
        Exit Sub
    End If

    ' Set last finished tracktime to current time to prevent parallel playing of this track
    sound_SetTrackTimeLastFinished iTrack, sound_GetGameTime()
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub audio_QueueTrackMinDelayPos(ByVal iTrack As Long, ByVal iMinDelay As Long, ByVal iX As Long, ByVal iY As Long, ByVal iZ As Long)
    Dim psSample As AUDIO_SAMPLE, iDelay As Double

    ' return if audio not enabled
    If audio_Disabled() Or audio_Paused() Then
        Exit Sub
    End If

    ' Determine if at least iMinDelay time has passed since the last time this track was played
    iDelay = udiff(sound_GetGameTime(), sound_GetTrackTimeLastFinished(iTrack))
    If iDelay < U32(iMinDelay) Then
        Exit Sub
    End If

    ' Construct an audio sample from requested track
    Set psSample = audio_QueueSample(iTrack)
    If psSample Is Nothing Then
        Exit Sub
    End If

    psSample.x = iX
    psSample.y = iY
    psSample.z = iZ

    ' Set last finished tracktime to current time to prevent parallel playing of this track
    sound_SetTrackTimeLastFinished iTrack, sound_GetGameTime()
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub audio_QueueTrackPos(ByVal iTrack As Long, ByVal iX As Long, ByVal iY As Long, ByVal iZ As Long)
    Dim psSample As AUDIO_SAMPLE

    ' return if audio not enabled
    If audio_Disabled() Or audio_Paused() Then
        Exit Sub
    End If

    ' Construct an audio sample from requested track
    Set psSample = audio_QueueSample(iTrack)
    If psSample Is Nothing Then
        Exit Sub
    End If

    psSample.x = iX
    psSample.y = iY
    psSample.z = iZ
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Private Sub audio_UpdateQueue()
    Dim psSample As AUDIO_SAMPLE

    ' return if audio not enabled
    If audio_Disabled() Or audio_Paused() Then
        Exit Sub
    End If

    If sound_QueueSamplePlaying() = True Then
        Exit Sub
    End If

    ' check queue for members
    If g_psSampleQueue.Count = 0 Then
        Exit Sub
    End If

    ' remove queue head
    Set psSample = g_psSampleQueue(1)
    audio_RemoveSample g_psSampleQueue, psSample

    ' add sample to list if able to play
    If Not sound_Play2DTrack(psSample, True) Then
        debug_ 503, LOG_NEVER, "audio_UpdateQueue", "audio_UpdateQueue: couldn't play sample\n"
        ' (free(psSample): the last reference goes)
        Exit Sub
    End If

    audio_AddSampleToHead g_psSampleList, psSample

    ' update last queue sound coords
    If psSample.x <> SAMPLE_COORD_INVALID And psSample.y <> SAMPLE_COORD_INVALID _
        And psSample.z <> SAMPLE_COORD_INVALID Then
        g_sPreviousSample.x = psSample.x
        g_sPreviousSample.y = psSample.y
        g_sPreviousSample.z = psSample.z
        g_sPreviousSample.iTrack = psSample.iTrack
        g_iPreviousSampleTime = realTime()
    End If
End Sub

Public Sub audio_Update()
    Dim playerPos_x As Single, playerPos_y As Single, playerPos_z As Single, angle As Single
    Dim items() As AUDIO_SAMPLE, n As Long, i As Long, psSample As AUDIO_SAMPLE, finished As Boolean
    Dim x As Long, y As Long, z As Long

    ' if audio not enabled return true to carry on game without audio
    If audio_Disabled() Then
        Exit Sub
    End If

    alGetError                          ' clear error codes
    audio_UpdateQueue
    alGetError                          ' clear error codes

    ' get player position
    audio_GetPlayerPos playerPos_x, playerPos_y, playerPos_z     ' playerPos = audio_GetPlayerPos()
    audio_Get3DPlayerRotAboutVerticalAxis angle
    sound_SetPlayerPos playerPos_x, playerPos_y, playerPos_z
    sound_SetPlayerOrientation angle

    ' loop through 3D sounds and remove if finished or update position
    n = listItems(g_psSampleList, items)
    For i = 1 To n
        Set psSample = items(i)
        ' remove finished samples from list
        If psSample.bFinishedPlaying = True Then
            audio_RemoveSample g_psSampleList, psSample
            ' (free(psSample): the last reference goes)
        Else ' check looping sound callbacks for finished condition
            If psSample.psObj <> 0 Then
                ' (VB) audio_ObjectDead(...) || (pCallback != nullptr && pCallback(psObj) == false), in that order
                finished = audio_ObjectDead(psSample.psObj)
                If Not finished Then
                    If psSample.pCallback <> 0 Then finished = (audio_CallBack(psSample.pCallback, psSample.psObj) = False)
                End If
                If finished Then
                    sound_StopTrack psSample
                    psSample.psObj = 0
                Else    ' update sample position
                    audio_GetObjectPos psSample.psObj, x, y, z
                    psSample.x = x
                    psSample.y = y
                    psSample.z = z
                    sound_SetObjectPosition psSample
                End If
            End If
        End If
    Next

    sound_Update
End Sub


' Retrieves loaded audio files and retrieves a TRACK from them on which some values are set and returns their respective ID numbers
'  \param fileName the filename of the track
'  \param loop whether the track should be looped until explicitly stopped
'  \param volume the volume this track should be played on (range is 0-100)
'  \param audibleRadius the radius from the source of sound where it can be heard
'  \return a non-zero value when successful or audio is disabled, zero when the file is not found or no more tracks can be loaded (i.e. the limit is reached)
Public Function audio_SetTrackVals(ByVal fileName As String, ByVal loop_ As Boolean, ByVal volume As Long, ByVal audibleRadius As Long) As Long
    ' if audio not enabled return a random non-zero value to carry on game without audio
    If audio_Disabled() Then
        audio_SetTrackVals = 1
        Exit Function
    End If

    debug_ 594, LOG_NEVER, "audio_SetTrackVals", "File: %s, loop: %d, volume: %d, radius: %d", fileName, -CLng(loop_), volume, audibleRadius
    audio_SetTrackVals = sound_SetTrackVals(fileName, loop_, volume, audibleRadius)
End Function

'*
'
'
' * audio_CheckTooManySame3DTracksPlaying Reject samples if too many already playing in
' * same area
'

'*
' =======================================================================================================================
' =======================================================================================================================
'
Private Function audio_CheckTooManySame3DTracksPlaying(ByVal iTrack As Long, ByVal iX As Long, ByVal iY As Long, ByVal iZ As Long) As Boolean
    Dim iCount As Long, iDx As Long, iDy As Long, iDz As Long, iDistSq As Long, iMaxDistSq As Long, iRad As Long
    Dim isTooMany As Boolean, psSample As AUDIO_SAMPLE, i As Long

    ' return if audio not enabled
    If audio_Disabled() Or audio_Paused() Then
        audio_CheckTooManySame3DTracksPlaying = True
        Exit Function
    End If

    ' loop through 3D sounds and check whether too many already in earshot
    For i = 1 To g_psSampleList.Count
        Set psSample = g_psSampleList(i)
        If psSample.iTrack = iTrack Then
            iDx = wrap32(CDbl(iX) - psSample.x)
            iDy = wrap32(CDbl(iY) - psSample.y)
            iDz = wrap32(CDbl(iZ) - psSample.z)
            iDistSq = wrap32((CDbl(iDx) * iDx) + (CDbl(iDy) * iDy) + (CDbl(iDz) * iDz))
            iRad = sound_GetTrackAudibleRadius(iTrack)
            iMaxDistSq = wrap32(CDbl(iRad) * iRad)
            If iDistSq < iMaxDistSq Then
                iCount = iCount + 1
            End If

            If iCount > MAX_SAME_SAMPLES Then
                isTooMany = True
                Exit For
            End If
        End If
    Next

    audio_CheckTooManySame3DTracksPlaying = isTooMany
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
' psObj: the game's object handle (0 = nullptr); pUserCallback: the game's number for the AUDIO_CALLBACK
Private Function audio_Play3DTrack(ByVal iX As Long, ByVal iY As Long, ByVal iZ As Long, ByVal iTrack As Long, ByVal psObj As Long, ByVal pUserCallback As Long) As Boolean
    Dim psSample As AUDIO_SAMPLE
    ' coordinates
    Dim listenerX As Single, listenerY As Single, listenerZ As Single, dX As Single, dY As Single, dZ As Single
    ' calculation results
    Dim distance As Single, gain As Single
    Dim err_ As Long

    ' if audio not enabled return true to carry on game without audio
    If audio_Disabled() Or audio_Paused() Then
        audio_Play3DTrack = False
        Exit Function
    End If

    If audio_CheckTooManySame3DTracksPlaying(iTrack, iX, iY, iZ) Then
        audio_Play3DTrack = False
        Exit Function
    End If

    ' compute distance
    ' NOTE, if this call fails, expect garbage
    alGetListener3f AL_POSITION, listenerX, listenerY, listenerZ
    err_ = sound_GetError("lib/sound/audio.cpp:676") ' sound_GetError(): __sound_GetError(AT_MACRO)
    If err_ <> AL_NO_ERROR Then
        audio_Play3DTrack = False
        Exit Function
    End If
    dX = CSng(iX) - listenerX ' distances on all axis
    dY = CSng(iY) - listenerY
    dZ = CSng(iZ) - listenerZ
    distance = CSng(Sqr(dX * dX + dY * dY + dZ * dZ)) ' Pythagorean theorem

    ' compute gain
    gain = (1! - (distance * ATTENUATION_FACTOR)) ' * 1.0f * sfx3d_volume
    If gain > 1! Then
        gain = 1!
    End If
    If gain < 0! Then
        gain = 0!
    End If
    ' don't bother adding samples that we can't hear
    If gain = 0! Then
        audio_Play3DTrack = False
        Exit Function
    End If

    Set psSample = New AUDIO_SAMPLE     ' malloc(sizeof(AUDIO_SAMPLE)) (VB: no out of memory)

    ' setup sample
    ' (memset(psSample, 0, sizeof(AUDIO_SAMPLE)): a new AUDIO_SAMPLE is zeroed)   ' [check] -Q
    psSample.iTrack = iTrack
    psSample.x = iX
    psSample.y = iY
    psSample.z = iZ
    psSample.bFinishedPlaying = False
    psSample.psObj = psObj
    psSample.pCallback = pUserCallback

    ' add sample to list if able to play
    If Not sound_Play3DTrack(psSample) Then
        debug_ 722, LOG_NEVER, "audio_Play3DTrack", "audio_Play3DTrack: couldn't play sample\n"
        ' (free(psSample): the last reference goes)
        audio_Play3DTrack = False
        Exit Function
    End If

    audio_AddSampleToHead g_psSampleList, psSample
    audio_Play3DTrack = True
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function audio_PlayStaticTrack(ByVal iMapX As Long, ByVal iMapY As Long, ByVal iTrack As Long) As Boolean
    Dim iX As Long, iY As Long, iZ As Long

    ' if audio not enabled return true to carry on game without audio
    If audio_Disabled() Or audio_Paused() Then
        audio_PlayStaticTrack = False
        Exit Function
    End If

    audio_GetStaticPos iMapX, iMapY, iX, iY, iZ
    audio_PlayStaticTrack = audio_Play3DTrack(iX, iY, iZ, iTrack, 0, 0)
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function audio_PlayObjStaticTrack(ByVal psObj As Long, ByVal iTrack As Long) As Boolean
    Dim iX As Long, iY As Long, iZ As Long

    ' if audio not enabled return true to carry on game without audio
    If audio_Disabled() Or audio_Paused() Then
        audio_PlayObjStaticTrack = False
        Exit Function
    End If

    audio_GetObjectPos psObj, iX, iY, iZ
    audio_PlayObjStaticTrack = audio_Play3DTrack(iX, iY, iZ, iTrack, psObj, 0)
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function audio_PlayObjStaticTrackCallback(ByVal psObj As Long, ByVal iTrack As Long, ByVal pUserCallback As Long) As Boolean
    Dim iX As Long, iY As Long, iZ As Long

    ' if audio not enabled return true to carry on game without audio
    If audio_Disabled() Or audio_Paused() Then
        audio_PlayObjStaticTrackCallback = False
        Exit Function
    End If

    audio_GetObjectPos psObj, iX, iY, iZ
    audio_PlayObjStaticTrackCallback = audio_Play3DTrack(iX, iY, iZ, iTrack, psObj, pUserCallback)
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function audio_PlayObjDynamicTrack(ByVal psObj As Long, ByVal iTrack As Long, ByVal pUserCallback As Long) As Boolean
    Dim iX As Long, iY As Long, iZ As Long

    ' if audio not enabled return true to carry on game without audio
    If audio_Disabled() Or audio_Paused() Then
        audio_PlayObjDynamicTrack = False
        Exit Function
    End If

    audio_GetObjectPos psObj, iX, iY, iZ
    audio_PlayObjDynamicTrack = audio_Play3DTrack(iX, iY, iZ, iTrack, psObj, pUserCallback)
End Function

' Plays the given audio file as a stream and reports back when it has finished
'  playing.
'  \param fileName the (OggVorbis|OggOpus) file to play from
'  \param volume the volume to use while playing this file (in the range of
'         0.0 - 1.0)
'  \param onFinished a callback function to invoke when playing of this stream
'         has been finished. You can use NULL to specifiy no callback function.
'  \param user_data a pointer to contain some user data to pass along to the
'         finished callback.
'  \return a pointer to the currently playing stream when the stream is playing
'          (and as such the callback will be invoked some time in the future),
'          NULL when the stream didn't start playing (and the callback won't be
'          invoked).
'  \note The returned pointer will become invalid/dangling immediately after
'        the \c onFinished callback is invoked.
'  \note You must _never_ manually free() the memory used by the returned
'        pointer.
Public Function audio_PlayStream(ByVal fileName As String, ByVal volume As Single, ByVal onFinished As IStreamFinished, ByVal user_data As Variant) As AUDIO_STREAM
    Dim stream As AUDIO_STREAM
    ' If audio is not enabled return false to indicate that the given callback
    ' will not be invoked.
    If audio_Disabled() Then
        Set audio_PlayStream = Nothing
        Exit Function
    End If
    Set stream = sound_PlayStream(fileName, False, volume, onFinished, user_data)
    Set audio_PlayStream = stream
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub audio_StopObjTrack(ByVal psObj As Long, ByVal iTrack As Long)
    Dim i As Long, psSample As AUDIO_SAMPLE

    ' return if audio not enabled
    If audio_Disabled() Then
        Exit Sub
    End If

    ' find sample
    For i = 1 To g_psSampleList.Count
        Set psSample = g_psSampleList(i)
        ' If track has been found stop it and return
        If psSample.psObj = psObj And psSample.iTrack = iTrack Then
            sound_StopTrack psSample
            Exit Sub
        End If
    Next
End Sub

' Play annoying sound only once per tick
Public Sub audio_PlayBuildFailedOnce()
    If udiff(realTime(), lastTimeBuildFailedPlayed) >= REPEAT_AUDIO_CUE_THRESHOLD Then
        audio_PlayTrack ID_SOUND_BUILD_FAIL
        lastTimeBuildFailedPlayed = realTime()
    End If
End Sub

'*
'
' audio_PlayTrack Play immediate 2D FX track

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub audio_PlayTrack(ByVal iTrack As Long, Optional ByVal playIfPaused As Boolean = True)
    Dim psSample As AUDIO_SAMPLE

    ' return if audio not enabled
    If audio_Disabled() Or (audio_Paused() And Not playIfPaused) Then
        Exit Sub
    End If

    ' Allocate a sample
    Set psSample = New AUDIO_SAMPLE     ' malloc(sizeof(AUDIO_SAMPLE)) (VB: no out of memory)

    ' setup/initialize sample
    psSample.iTrack = iTrack
    psSample.bFinishedPlaying = False

    ' Zero callback stuff since we don't need/want it
    psSample.pCallback = 0
    psSample.psObj = 0

    ' iSample, psPrev, and psNext will be initialized by the
    ' following functions, and x, y and z will be completely
    ' ignored so we don't need to bother about it

    ' add sample to list if able to play
    If Not sound_Play2DTrack(psSample, False) Then
        debug_ 920, LOG_NEVER, "audio_PlayTrack", "audio_PlayTrack: couldn't play sample\n"
        ' (free(psSample): the last reference goes)
        Exit Sub
    End If

    audio_AddSampleToHead g_psSampleList, psSample
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub audio_setPause(ByVal state As Boolean)
    ' return if audio not enabled
    If audio_Disabled() Then
        Exit Sub
    End If

    g_bAudioPaused = state
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub audio_StopAll()
    Dim i As Long, psSample As AUDIO_SAMPLE

    ' return if audio not enabled
    If audio_Disabled() Then
        Exit Sub
    End If

    '
    ' * empty list - audio_Update will free samples because callbacks have to come in first
    '
    For i = 1 To g_psSampleList.Count
        Set psSample = g_psSampleList(i)
        ' Stop this sound sample
        sound_StopTrack psSample

        ' HACK:
        ' Make sure to set psObj to NULL, since sometimes it becomes
        ' invalidated, i.e. dangling, before audio_Update() gets the
        ' chance to clean it up.
        '
        ' Meaning at this place in the code audio_ObjectDead(psObj)
        ' would indicate that psObj is still valid and will remain
        ' such. While in reality, before audio_Update() can get the
        ' chance to check again, this pointer will have become
        ' dangling.
        '
        ' NOTE: This would always happen when audio_StopAll() had been
        ' invoked by stageThreeShutDown().
        psSample.psObj = 0
    Next

    ' empty sample queue
    ' Destroy the sample
    Set g_psSampleQueue = New Collection
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function audio_GetTrackID(ByVal fileName As String) As Long
    Dim psTrack As TRACK, s As String

    ' return if audio not enabled
    If audio_Disabled() Then
        audio_GetTrackID = SAMPLE_NOT_FOUND
        Exit Function
    End If

    If StrPtr(fileName) = 0 Or Len(fileName) = 0 Then
        If StrPtr(fileName) = 0 Then s = "a NULL pointer" Else s = "empty"
        debug_ 1006, LOG_WARNING, "audio_GetTrackID", "fileName is %s", s
        audio_GetTrackID = SAMPLE_NOT_FOUND
        Exit Function
    End If

    Set psTrack = trackOfResource(resGetData("WAV", fileName))
    If psTrack Is Nothing Then
        audio_GetTrackID = SAMPLE_NOT_FOUND
        Exit Function
    End If

    audio_GetTrackID = sound_GetTrackID(psTrack)
End Function

' (VB) (TRACK *)resGetData("WAV", fileName): Nothing when the resource holds no TRACK
Private Function trackOfResource(ByVal data As Variant) As TRACK
    If IsObject(data) Then Set trackOfResource = data
End Function

' Loop through the list of playing and queued audio samples, and destroy any
'  of them that refer to the given object.
'  \param psObj pointer to the object for which we must destroy all of its
'               outstanding audio samples.
Public Sub audio_RemoveObj(ByVal psObj As Long)
    Dim count As Long, items() As AUDIO_SAMPLE, n As Long, i As Long, psSample As AUDIO_SAMPLE, toRemove As AUDIO_SAMPLE

    ' loop through queued sounds and check if a sample needs to be removed
    n = listItems(g_psSampleQueue, items)
    For i = 1 To n
        Set psSample = items(i)
        If psSample.psObj = psObj Then
            ' The current audio sample seems to refer to an object
            ' that is about to be destroyed. So destroy this
            ' sample as well.
            Set toRemove = psSample

            debug_ 1040, LOG_MEMORY, "audio_RemoveObj", "audio_RemoveObj: callback %p sample %d\n", toRemove.pCallback, toRemove.iTrack
            ' Remove sound from global active list
            sound_RemoveActiveSample toRemove   'remove from global active list.

            ' Perform the actual task of destroying this sample
            audio_RemoveSample g_psSampleQueue, toRemove
            ' (free(toRemove): the last reference goes)

            ' Increment the deletion count
            count = count + 1
        End If
    Next

    If count <> 0 Then
        debug_ 1056, LOG_MEMORY, "audio_RemoveObj", "audio_RemoveObj: BASE_OBJECT* 0x%p was found %u times in the audio sample queue", psObj, count
    End If

    ' Reset the deletion count
    count = 0

    ' loop through list of currently playing sounds and check if a sample needs to be removed
    n = listItems(g_psSampleList, items)
    For i = 1 To n
        Set psSample = items(i)
        If psSample.psObj = psObj Then
            ' The current audio sample seems to refer to an object
            ' that is about to be destroyed. So destroy this
            ' sample as well.
            Set toRemove = psSample

            debug_ 1074, LOG_MEMORY, "audio_RemoveObj", "audio_RemoveObj: callback %p sample %d\n", toRemove.pCallback, toRemove.iTrack
            ' Stop this sound sample
            sound_RemoveActiveSample toRemove   'remove from global active list.

            ' Perform the actual task of destroying this sample
            audio_RemoveSample g_psSampleList, toRemove
            ' (free(toRemove): the last reference goes)

            ' Increment the deletion count
            count = count + 1
        End If
    Next

    If count <> 0 Then
        debug_ 1090, LOG_MEMORY, "audio_RemoveObj", "audio_RemoveObj: ***Warning! psOBJ %p was found %u times in the list of playing audio samples", psObj, count
    End If
End Sub
