Attribute VB_Name = "openal_track_c"
' lib/sound/openal_track.cpp: the sound library's OpenAL side (tracklib.h), over al_win (the subset of
' OpenAL WZ uses, on waveOut). A TRACK *, AUDIO_SAMPLE *, AUDIO_STREAM * is the class (Nothing =
' nullptr); the std::lists of active samples and streams are Collections (emplace_front adds before the
' first); mutating_list_iterate is a loop that takes the next element before its body runs (the body
' may remove the current one). sound_GetError() and sound_GetContextError(device) pass their
' "file:line" (openal_error_c). ALC_SOFT_HRTF is defined, as in 4.7's builds with OpenAL Soft's headers:
' al_win has no extension, so the HRTF is Unsupported. WZ_OS_UNIX and DEBUG are not defined (AL_PITCH
' 1.0; a release build). A WZDecoder * is the interface WZDecoder (WZVorbisDecoder, WZOpusDecoder).
' The statics' initialisers run at the first call (staticsInit). sound_vbSetCapture (VB): a test aid.
Option Explicit

' track.h (the macros this file uses)
Private Const ATTENUATION_FACTOR As Single = 0.0003!
' (SAMPLE_NOT_ALLOCATED: SoundApi's)

Private current_queue_sample As Long            ' ALuint = -1

Private openal_initialized As Boolean           ' = false
Private Const bufferSize As Long = 16 * 1024
Private Const buffer_count As Long = 32

' (struct AUDIO_STREAM: AUDIO_STREAM.cls)

Private active_samples As Collection            ' std::list<AUDIO_SAMPLE *>

' actives openAL-Sources
Private active_streams As Collection            ' std::list<AUDIO_STREAM *>

Private sfx_volume As Single                    ' = 1.0
Private sfx3d_volume As Single                  ' = 1.0

Private device As Long                          ' ALCdevice * = nullptr
Private context As Long                         ' ALCcontext * = nullptr

' (ALC_SOFT_HRTF) LPALCGETSTRINGISOFT alcGetStringiSOFT, LPALCRESETDEVICESOFT alcResetDeviceSOFT:
' function pointers (al_win's alcGetProcAddress finds none: 0)
Private alcGetStringiSOFT As Long
Private alcResetDeviceSOFT As Long

Private staticsReady As Boolean

' (VB) the statics' initialisers
Private Sub staticsInit()
    If staticsReady Then Exit Sub
    staticsReady = True
    current_queue_sample = -1
    Set active_samples = New Collection
    Set active_streams = New Collection
    sfx_volume = 1!
    sfx3d_volume = 1!
End Sub

' lib/exceptionhandler/dumpinfo.h (not ported): the crash dump's extra lines
Private Sub addDumpInfo(ByVal inbuffer As String)
End Sub

' (VB) the std::list helpers: the first element (Nothing if none), the one after obj (Nothing at the end),
' obj's removal, obj added at the front
Private Function listFirst(ByVal col As Collection) As Object
    If col.Count > 0 Then Set listFirst = col(1)
End Function

Private Function listAfter(ByVal col As Collection, ByVal obj As Object) As Object
    Dim i As Long
    For i = 1 To col.Count
        If col(i) Is obj Then
            If i < col.Count Then Set listAfter = col(i + 1)
            Exit Function
        End If
    Next
End Function

Private Sub listRemove(ByVal col As Collection, ByVal obj As Object)
    Dim i As Long
    For i = 1 To col.Count
        If col(i) Is obj Then
            col.Remove i
            Exit Sub
        End If
    Next
End Sub

Private Sub listPushFront(ByVal col As Collection, ByVal obj As Object)
    If col.Count = 0 Then col.Add obj Else col.Add obj, , 1
End Sub

' (VB) const void *user_data == nullptr
Private Function isNullUserData(user_data As Variant) As Boolean
    If IsObject(user_data) Then
        isNullUserData = (user_data Is Nothing)
    ElseIf IsEmpty(user_data) Then
        isNullUserData = True
    ElseIf IsNumeric(user_data) Then
        isNullUserData = (user_data = 0)
    End If
End Function

Private Function HRTFModeToALCint(ByVal mode As Long) As Long
    ' (ALC_SOFT_HRTF)
    Select Case mode
    Case HRTFMode_Unsupported
        ' should never be called with unsupported, log it and fall through to disabled
        debug_ 100, LOG_ERROR, "HRTFModeToALCint", "HRTFModeToALCint called with HRTFMode::Unsupported"
        ' fallthrough
        HRTFModeToALCint = ALC_FALSE
        Exit Function
    Case HRTFMode_Disabled
        HRTFModeToALCint = ALC_FALSE
        Exit Function
    Case HRTFMode_Enabled
        HRTFModeToALCint = ALC_TRUE
        Exit Function
    Case HRTFMode_Auto
        HRTFModeToALCint = ALC_DONT_CARE_SOFT
        Exit Function
    End Select
    HRTFModeToALCint = ALC_FALSE
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_InitLibrary(ByVal hrtf As Long) As Boolean
    Dim err_ As Long, listenerVel(0 To 2) As Single, listenerOri(0 To 5) As Single, buf As String
    Dim deviceName As String, info As String, hrtfStatus As Long, hrtfStatusString As String
    staticsInit
    ' const ALfloat listenerVel[3] = { 0.0, 0.0, 0.0 };
    ' const ALfloat listenerOri[6] = { 0.0, 0.0, 1.0, 0.0, 1.0, 0.0 };
    listenerOri(2) = 1!
    listenerOri(4) = 1!

    ' (#if 0: the enumeration of the devices, disabled because it apparently crashes PulseAudio on Fedora12)

    If debugPartEnabled(LOG_SOUND) Then
        info = Output_PlaybackDevices()
        addDumpInfo info
        If debugPartEnabled(LOG_SOUND) Then
            debug_multiline 0, LOG_SOUND, "sound", info
        End If
    End If

    ' Open default device
    device = alcOpenDevice(vbNullString)

    If device = 0 Then
        debug_ 156, LOG_ERROR, "sound_InitLibrary", "Couldn't open audio device."
        sound_InitLibrary = False
        Exit Function
    End If

    If debugPartEnabled(LOG_SOUND) Then
        info = Output_ALCInfo(device)
        ' addDumpInfo(output.c_str());
        'if (enabled_debug[LOG_SOUND])
        '{
        debug_multiline 0, LOG_SOUND, "sound", info
        '}
    End If

    ' Load some extensions from OpenAL-Soft (if available)
    alcGetStringiSOFT = alcGetProcAddress(device, "alcGetStringiSOFT")
    alcResetDeviceSOFT = alcGetProcAddress(device, "alcResetDeviceSOFT")

    ' Print current device name and add it to dump info
    deviceName = alcGetString(device, ALC_DEVICE_SPECIFIER)
    debug_ 180, LOG_SOUND, "sound_InitLibrary", "Current audio device: %s", deviceName
    buf = Left$("OpenAL Device Name: " & deviceName, 511)   ' ssprintf(buf, ...): char buf[512]
    addDumpInfo buf

    context = alcCreateContext(device, 0)       'NULL was contextAttributes
    If context = 0 Then
        debug_ 187, LOG_ERROR, "sound_InitLibrary", "Couldn't open audio context."

        debug_ 189, LOG_SOUND, "sound_InitLibrary", "close device"
        If alcCloseDevice(device) = ALC_FALSE Then
            debug_ 192, LOG_SOUND, "sound_InitLibrary", "OpenAl could not close the audio device."
        End If
        device = 0
        sound_InitLibrary = False
        Exit Function
    End If

    alcMakeContextCurrent context

    err_ = sound_GetContextError(device, "lib/sound/openal_track.cpp:200")
    If err_ <> ALC_NO_ERROR Then
        debug_ 203, LOG_ERROR, "sound_InitLibrary", "Couldn't initialize audio context: %s", alcGetString(device, err_)

        alcGetError device      ' clear error codes

        alcMakeContextCurrent 0
        sound_GetContextError device, "lib/sound/openal_track.cpp:208"
        alcDestroyContext context ' this gives a long delay on some impl.
        context = 0
        sound_GetContextError device, "lib/sound/openal_track.cpp:211"

        debug_ 213, LOG_SOUND, "sound_InitLibrary", "close device"
        If alcCloseDevice(device) = ALC_FALSE Then
            debug_ 216, LOG_SOUND, "sound_InitLibrary", "OpenAl could not close the audio device."
        End If
        device = 0
        sound_InitLibrary = False
        Exit Function
    End If

    ' Dump Open AL device info (depends on context)
    ' to the crash handler for the dump file and debug log
    buf = Left$("OpenAL Vendor: " & alGetString(AL_VENDOR), 511)
    addDumpInfo buf
    debug_ 226, LOG_SOUND, "sound_InitLibrary", "%s", buf

    buf = Left$("OpenAL Version: " & alGetString(AL_VERSION), 511)
    addDumpInfo buf
    debug_ 230, LOG_SOUND, "sound_InitLibrary", "%s", buf

    buf = Left$("OpenAL Renderer: " & alGetString(AL_RENDERER), 511)
    addDumpInfo buf
    debug_ 234, LOG_SOUND, "sound_InitLibrary", "%s", buf

    buf = Left$("OpenAL Extensions: " & alGetString(AL_EXTENSIONS), 511)
    addDumpInfo buf
    debug_ 238, LOG_SOUND, "sound_InitLibrary", "%s", buf

    openal_initialized = True

    ' (ALC_SOFT_HRTF)
    If alcIsExtensionPresent(device, "ALC_SOFT_HRTF") <> ALC_FALSE Then
        ' Set desired HRTF mode
        sound_SetHRTFMode hrtf

        ' Get current HRTF status
        alcGetIntegerv device, ALC_HRTF_STATUS_SOFT, 1, hrtfStatus

        hrtfStatusString = vbNullString
        Select Case hrtfStatus
        Case ALC_HRTF_DISABLED_SOFT
            hrtfStatusString = "ALC_HRTF_DISABLED_SOFT: HRTF is disabled"
        Case ALC_HRTF_ENABLED_SOFT
            hrtfStatusString = "ALC_HRTF_ENABLED_SOFT: HRTF is enabled"
        Case ALC_HRTF_DENIED_SOFT
            ' This may be caused by invalid resource permissions, or other user configuration that disallows HRTF.
            hrtfStatusString = "ALC_HRTF_DENIED_SOFT: HRTF is disabled because it's not allowed on the device."
        Case ALC_HRTF_REQUIRED_SOFT
            ' This may be caused by a device that can only use HRTF, or other user configuration that forces HRTF to be used.
            hrtfStatusString = "ALC_HRTF_REQUIRED_SOFT: HRTF is enabled because it must be used on the device."
        Case ALC_HRTF_HEADPHONES_DETECTED_SOFT
            hrtfStatusString = "ALC_HRTF_HEADPHONES_DETECTED_SOFT: HRTF is enabled automatically because the device reported itself as headphones."
        Case ALC_HRTF_UNSUPPORTED_FORMAT_SOFT
            ' HRTF is disabled because the device does not support it with the current format.
            ' Typically this is caused by non-stereo output or an incompatible output frequency.
            hrtfStatusString = "ALC_HRTF_UNSUPPORTED_FORMAT_SOFT: HRTF is disabled because the device does not support it with the current format."
        Case Else
            hrtfStatusString = vbNullString
        End Select

        If StrPtr(hrtfStatusString) <> 0 Then
            debug_ 284, LOG_SOUND, "sound_InitLibrary", "%s", hrtfStatusString
        Else
            debug_ 288, LOG_SOUND, "sound_InitLibrary", "OpenAL-Soft returned an unknown ALC_HRTF_STATUS_SOFT result: %d", hrtfStatus
        End If
    Else
        debug_ 293, LOG_SOUND, "sound_InitLibrary", "alcIsExtensionPresent(..., ""ALC_SOFT_HRTF"") returned false"
    End If

    ' Clear Error Codes
    alGetError
    alcGetError device

    alListener3f AL_POSITION, 0!, 0!, 0!
    alListenerfv AL_VELOCITY, listenerVel
    alListenerfv AL_ORIENTATION, listenerOri
    alDistanceModel AL_NONE
    sound_GetError "lib/sound/openal_track.cpp:307"

    sound_InitLibrary = True
End Function

Public Function sound_GetHRTFMode() As Long
    Dim hrtfStatus As Long
    staticsInit
    ' (ALC_SOFT_HRTF)
    If alcIsExtensionPresent(device, "ALC_SOFT_HRTF") <> ALC_FALSE Then
        alcGetIntegerv device, ALC_HRTF_SOFT, 1, hrtfStatus
        Select Case hrtfStatus
        Case ALC_TRUE
            sound_GetHRTFMode = HRTFMode_Enabled
            Exit Function
        Case ALC_FALSE
            sound_GetHRTFMode = HRTFMode_Disabled
            Exit Function
        Case Else
            debug_ 326, LOG_SOUND, "sound_GetHRTFMode", "OpenAL-Soft returned an unexpected ALC_HRTF_SOFT result: %d", hrtfStatus
        End Select
    End If
    sound_GetHRTFMode = HRTFMode_Unsupported
End Function

Public Function sound_SetHRTFMode(ByVal mode As Long) As Boolean
    Dim hrtfSetting As Long
    staticsInit
    ' (ALC_SOFT_HRTF)
    If alcIsExtensionPresent(device, "ALC_SOFT_HRTF") <> ALC_FALSE Then
        hrtfSetting = HRTFModeToALCint(mode)

        ' ALCint attrs[] = { ALC_HRTF_SOFT, hrtfSetting, /* configure HRTF */ 0 /* end of list */ };
        If alcResetDeviceSOFT <> 0 Then
            If Not (device <> 0) Then ASSERT_FAILURE "device", "lib/sound/openal_track.cpp:346", "sound_SetHRTFMode", "device is null"
            ' (VB) alcResetDeviceSOFT(device, attrs): a function pointer al_win never hands out (it has no
            ' extension), so this branch is never taken; on its failure the C++ logs "Failed to reset
            ' device: %s\n" (line 349) and returns false
            sound_SetHRTFMode = False
            Exit Function
        Else
            debug_ 356, LOG_ERROR, "sound_SetHRTFMode", "ALC_SOFT_HRTF extension is available, but alcResetDeviceSOFT is null"
        End If
    End If
    sound_SetHRTFMode = False
End Function

Public Sub sound_ShutdownLibrary()
    Dim it As AUDIO_STREAM
    staticsInit
    If Not openal_initialized Then
        Exit Sub
    End If
    debug_ 371, LOG_SOUND, "sound_ShutdownLibrary", "starting shutdown"

    ' Stop all streams, sound_UpdateStreams() will deallocate all stopped streams
    For Each it In active_streams
        sound_StopStream it
    Next
    sound_UpdateStreams

    alcGetError device      ' clear error codes

    ' On Linux since this caused some versions of OpenAL to hang on exit. - Per
    debug_ 383, LOG_SOUND, "sound_ShutdownLibrary", "make default context NULL"
    alcMakeContextCurrent 0
    sound_GetContextError device, "lib/sound/openal_track.cpp:385"

    debug_ 387, LOG_SOUND, "sound_ShutdownLibrary", "destroy previous context"
    alcDestroyContext context ' this gives a long delay on some impl.
    context = 0
    sound_GetContextError device, "lib/sound/openal_track.cpp:390"

    debug_ 392, LOG_SOUND, "sound_ShutdownLibrary", "close device"
    If alcCloseDevice(device) = ALC_FALSE Then
        debug_ 395, LOG_SOUND, "sound_ShutdownLibrary", "OpenAl could not close the audio device."
    End If
    device = 0

    Set active_samples = New Collection     ' active_samples.clear()
End Sub

' Deletes the given sample and performs additional OpenAL cleanup procedures.
' \param sample iterator to the current sample in the list which you want to delete, invalidated upon return
' (VB) the sample itself: removed from active_samples by identity
Private Sub sound_DestroyIteratedSample(ByVal sample As AUDIO_SAMPLE)
    Dim ids(0 To 0) As Long
    ' If an OpenAL source is associated with this sample, release it
    If sample.iSample <> AL_INVALID Then
        ids(0) = sample.iSample
        alDeleteSources 1, ids
        sound_GetError "lib/sound/openal_track.cpp:412"
    End If

    ' Do the cleanup of this sample
    sound_FinishedCallback sample

    ' Remove the sample from the list
    listRemove active_samples, sample
End Sub

' Counts the number of samples in active_samples
' \return the number of actively playing sound samples
Public Function sound_GetActiveSamplesCount() As Long
    staticsInit
    sound_GetActiveSamplesCount = active_samples.Count
End Function

' gets called in audio.cpp: audio_update(), which gets called in renderLoop()
Public Sub sound_Update()
    Dim it As AUDIO_SAMPLE, nxt As AUDIO_SAMPLE, state As Long, err_ As Long, gain As Single
    staticsInit

    If Not openal_initialized Then
        Exit Sub
    End If

    ' Update all streaming audio
    sound_UpdateStreams

    Set it = listFirst(active_samples)
    Do While Not it Is Nothing
        Set nxt = listAfter(active_samples, it)

        ' query what the gain is for this sample
        gain = -1!                      ' (VB) the C's gain is undefined when alGetSourcef fails: not 0 here
        alGetSourcef it.iSample, AL_GAIN, gain
        err_ = sound_GetError("lib/sound/openal_track.cpp:451")

        ' if gain is 0, then we can't hear it, so we kill it.
        If gain = 0! Then
            sound_DestroyIteratedSample it
        Else
            'ASSERT(alIsSource(node->curr->iSample), "Not a valid source!");
            alGetSourcei it.iSample, AL_SOURCE_STATE, state

            ' Check whether an error occurred while retrieving the state.
            ' If one did, the state returned is useless. So instead of
            ' using it continue with the next sample.
            err_ = sound_GetError("lib/sound/openal_track.cpp:466")
            If err_ <> AL_NO_ERROR Then
                ' Make sure to invoke the "finished" callback
                sound_FinishedCallback it

                ' Destroy this object and move to the next object
                sound_DestroyIteratedSample it
            Else
                Select Case state
                Case AL_PLAYING, AL_PAUSED
                    ' If we haven't finished playing yet, just
                    ' continue with the next item in the list.

                    ' sound_SetObjectPosition(i->curr->iSample, i->curr->x, i->curr->y, i->curr->z);

                    ' Move to the next object

                    ' NOTE: if it isn't playing | paused, then it is most likely either done
                    ' or a error.  In either case, we want to kill the sample in question.

                Case Else
                    sound_DestroyIteratedSample it
                End Select
            End If
        End If
        Set it = nxt
    Loop

    ' Reset the current error state
    alcGetError device

    alcProcessContext context

    err_ = sound_GetContextError(device, "lib/sound/openal_track.cpp:504")
    If err_ <> ALC_NO_ERROR Then
        debug_ 507, LOG_ERROR, "sound_Update", "Error while processing audio context: %s", alGetString(err_)
    End If
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_QueueSamplePlaying() As Boolean
    Dim state As Long, it As AUDIO_SAMPLE, nxt As AUDIO_SAMPLE, sampleFound As Boolean
    staticsInit

    If Not openal_initialized Then
        sound_QueueSamplePlaying = False
        Exit Function
    End If
    If current_queue_sample = AL_INVALID Then
        sound_QueueSamplePlaying = False
        Exit Function
    End If

    alGetSourcei current_queue_sample, AL_SOURCE_STATE, state

    ' Check whether an error occurred while retrieving the state.
    ' If one did, the state returned is useless. So instead of
    ' using it return false.
    If sound_GetError("lib/sound/openal_track.cpp:533") <> AL_NO_ERROR Then
        sound_QueueSamplePlaying = False
        Exit Function
    End If

    If state = AL_PLAYING Then
        sound_QueueSamplePlaying = True
        Exit Function
    End If

    If current_queue_sample <> AL_INVALID Then
        sampleFound = False
        Set it = listFirst(active_samples)
        Do While Not it Is Nothing
            Set nxt = listAfter(active_samples, it)
            If it.iSample = current_queue_sample Then
                sound_DestroyIteratedSample it
                current_queue_sample = AL_INVALID
                sampleFound = True
                Exit Do                 ' IterationResult::BREAK_ITERATION
            End If
            Set it = nxt
        Loop
        If sampleFound Then
            sound_QueueSamplePlaying = False
            Exit Function
        End If

        debug_ 562, LOG_ERROR, "sound_QueueSamplePlaying", "Sample %u not deleted because it wasn't in the active queue!", current_queue_sample
        current_queue_sample = AL_INVALID
    End If
    sound_QueueSamplePlaying = False
End Function

' Decodes *entirely* an opened OggVorbis file into an OpenAL buffer.
' This is used to play sound effects, not "music". Assumes .ogg file.
'
' \param psTrack pointer to object which will contain the final buffer
' \param PHYSFS_fileHandle file handle given by PhysicsFS to the opened file
' \return true on success
Private Function sound_DecodeOggVorbisTrack(ByVal psTrack As TRACK, ByVal fileName As String) As Boolean
    Dim decoder As WZVorbisDecoder, estimate As Long, buffer() As Byte, written As Long, fmt_ As Long, alBuffer(0 To 0) As Long

    If Not openal_initialized Then
        sound_DecodeOggVorbisTrack = False
        Exit Function
    End If

    Set decoder = WZVorbisDecoder_fromFilename(fileName)
    If decoder Is Nothing Then
        debug_ 585, LOG_ERROR, "sound_DecodeOggVorbisTrack", "couldn't allocate decoder for %s", fileName
        sound_DecodeOggVorbisTrack = False
        Exit Function
    End If
    estimate = decoder.totalSamples() * decoder.channels() * 2
    ' uint8_t* buffer = (uint8_t*) malloc(estimate); memset(buffer, 0, estimate);
    ' (VB: ReDim raises an error instead of the C's nullptr; a VB array has one element at least)
    If estimate > 0 Then ReDim buffer(0 To estimate - 1) Else ReDim buffer(0 To 0)
    If Not decoder.decode(buffer, estimate, written) Then
        debug_ 600, LOG_ERROR, "sound_DecodeOggVorbisTrack", "failed decoding %s", fileName
        Set decoder = Nothing
        sound_DecodeOggVorbisTrack = False
        Exit Function
    End If

    ' Determine PCM data format
    If decoder.channels() = 1 Then fmt_ = AL_FORMAT_MONO16 Else fmt_ = AL_FORMAT_STEREO16
    ' Create an OpenAL buffer and fill it with the decoded data
    alGenBuffers 1, alBuffer
    sound_GetError "lib/sound/openal_track.cpp:611"
    If Not (CDbl(estimate) <= 2147483647#) Then ASSERT_FAILURE "estimate <= static_cast<size_t>(std::numeric_limits<ALsizei>::max())", "lib/sound/openal_track.cpp:612", "sound_DecodeOggVorbisTrack", "soundBuffer->size (%u) exceeds ALsizei::max", estimate
    If Not (CDbl(decoder.frequency()) <= 2147483647#) Then ASSERT_FAILURE "decoder->frequency() <= static_cast<size_t>(std::numeric_limits<ALsizei>::max())", "lib/sound/openal_track.cpp:613", "sound_DecodeOggVorbisTrack", "decoder->frequency() (%zu) exceeds ALsizei::max ??", decoder.frequency()
    alBufferData alBuffer(0), fmt_, buffer, estimate, decoder.frequency()
    sound_GetError "lib/sound/openal_track.cpp:615"

    ' save buffer name in track
    psTrack.iBufferName = alBuffer(0)
    Set decoder = Nothing
    sound_DecodeOggVorbisTrack = True
End Function

' This is used to play sound effets (not "music"). Assumes .ogg file.
' \param [in] fileName: <soundeffect>.ogg
' \returns Track pointer, or nullptr on failure
Public Function sound_LoadTrackFromFile(ByVal fileName As String) As TRACK
    Dim pTrack As TRACK, filename_size As Long, track_name As String, lastName As String
    staticsInit

    lastName = GetLastResourceFilename()
    If StrPtr(lastName) = 0 Then
        ' This is a non fatal error.  We just can't find filename for some reason.
        debug_ 637, LOG_WARNING, "sound_LoadTrackFromFile", "sound_LoadTrackFromFile: missing resource filename?"
        filename_size = 0
    Else
        filename_size = Len(lastName) + 1
    End If

    ' allocate track, plus the memory required to contain the filename
    ' one malloc call ensures only one free call is required
    ' (VB: New raises an error where malloc returns nullptr: no "couldn't allocate memory" here)
    Set pTrack = New TRACK

    ' Initialize everything (except for the filename) to zero (a new TRACK's fields are)

    ' Set filename pointer; if the filename (as returned by
    ' GetLastResourceFilename()) is a NULL pointer, then this will be a
    ' NULL pointer as well.
    ' Copy the filename into the struct, if we don't have a NULL pointer
    If filename_size <> 0 Then track_name = lastName Else track_name = vbNullString
    pTrack.fileName = track_name

    ' Now use sound_ReadTrackFromBuffer to decode the file's contents
    If Not sound_DecodeOggVorbisTrack(pTrack, fileName) Then
        Set pTrack = Nothing    ' free(pTrack)
        Set sound_LoadTrackFromFile = Nothing
        Exit Function
    End If

    Set sound_LoadTrackFromFile = pTrack
End Function

Public Sub sound_FreeTrack(ByVal psTrack As TRACK)
    Dim ids(0 To 0) As Long
    staticsInit
    ids(0) = psTrack.iBufferName
    alDeleteBuffers 1, ids
    sound_GetError "lib/sound/openal_track.cpp:683"
End Sub

Private Sub sound_AddActiveSample(ByVal psSample As AUDIO_SAMPLE)
    ' Prepend the given sample to our list of active samples
    listPushFront active_samples, psSample
End Sub

' Routine gets rid of the psObj's sound sample and reference in active_samples.
Public Sub sound_RemoveActiveSample(ByVal psSample As AUDIO_SAMPLE)
    Dim it As AUDIO_SAMPLE, nxt As AUDIO_SAMPLE
    staticsInit
    Set it = listFirst(active_samples)
    Do While Not it Is Nothing
        Set nxt = listAfter(active_samples, it)
        If it.psObj = psSample.psObj Then
            debug_ 704, LOG_MEMORY, "sound_RemoveActiveSample", "Removing object 0x%p from active_samples list" & vbLf, psSample.psObj

            ' Buginator: should we wait for it to finish, or just stop it?
            sound_StopSample it

            sound_FinishedCallback it   'tell the callback it is finished.

            sound_DestroyIteratedSample it
        End If
        ' (else: Move to the next sample object)
        Set it = nxt
    Loop
End Sub

Private Function sound_SetupChannel(ByVal psSample As AUDIO_SAMPLE) As Boolean
    sound_AddActiveSample psSample

    sound_SetupChannel = sound_TrackLooped(psSample.iTrack)
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_Play2DSample(ByVal psTrack As TRACK, ByVal psSample As AUDIO_SAMPLE, ByVal bQueued As Boolean) As Boolean
    Dim zero(0 To 2) As Single, volume As Single, error_ As Long, ids(0 To 0) As Long
    staticsInit

    If sfx_volume = 0! Then
        sound_Play2DSample = False
        Exit Function
    End If
    volume = CSng(psTrack.iVol) / 100!      ' each object can have OWN volume!
    psSample.fVol = volume                  ' save computed volume
    volume = volume * sfx_volume            ' and now take into account the Users sound Prefs.

    ' We can't hear it, so don't bother creating it.
    If volume = 0! Then
        sound_Play2DSample = False
        Exit Function
    End If

    ' Clear error codes
    alGetError

    ids(0) = psSample.iSample               ' (alGenSources writes &psSample->iSample: kept when it fails)
    alGenSources 1, ids
    psSample.iSample = ids(0)

    error_ = sound_GetError("lib/sound/openal_track.cpp:753")
    If error_ <> AL_NO_ERROR Then
        ' FIXME: We run out of OpenAL sources very quickly, so we
        ' should handle the case where we've ran out of them.
        ' Currently we don't do this, causing some unpleasant side
        ' effects, e.g. crashing...
    End If

    alSourcef psSample.iSample, AL_PITCH, 1!
    alSourcef psSample.iSample, AL_GAIN, volume
    alSourcefv psSample.iSample, AL_POSITION, zero
    alSourcefv psSample.iSample, AL_VELOCITY, zero
    alSourcei psSample.iSample, AL_BUFFER, psTrack.iBufferName
    alSourcei psSample.iSample, AL_SOURCE_RELATIVE, AL_TRUE
    If sound_SetupChannel(psSample) Then
        alSourcei psSample.iSample, AL_LOOPING, AL_TRUE
    Else
        alSourcei psSample.iSample, AL_LOOPING, AL_FALSE
    End If

    ' NOTE: this is only useful for debugging. (#ifdef DEBUG: is3d, isLooping, filename)
    ' Clear error codes
    alGetError

    alSourcePlay psSample.iSample
    sound_GetError "lib/sound/openal_track.cpp:782"

    If bQueued Then
        current_queue_sample = psSample.iSample
    ElseIf current_queue_sample = psSample.iSample Then
        current_queue_sample = -1
    End If

    sound_Play2DSample = True
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_Play3DSample(ByVal psTrack As TRACK, ByVal psSample As AUDIO_SAMPLE) As Boolean
    Dim zero(0 To 2) As Single, volume As Single, error_ As Long, ids(0 To 0) As Long
    staticsInit

    If sfx3d_volume = 0! Then
        sound_Play3DSample = False
        Exit Function
    End If

    volume = CSng(psTrack.iVol) / 100!      ' max range is 0-100
    psSample.fVol = volume                  ' store results for later

    ' If we can't hear it, then don't bother playing it.
    If volume = 0! Then
        sound_Play3DSample = False
        Exit Function
    End If
    ' Clear error codes
    alGetError

    ids(0) = psSample.iSample
    alGenSources 1, ids
    psSample.iSample = ids(0)

    error_ = sound_GetError("lib/sound/openal_track.cpp:824")
    If error_ <> AL_NO_ERROR Then
        ' FIXME: We run out of OpenAL sources very quickly, so we
        ' should handle the case where we've ran out of them.
        ' Currently we don't do this, causing some unpleasant side
        ' effects, e.g. crashing...
    End If

    ' (WZ_OS_UNIX: AL_PITCH 1.001f, a workaround for 64-bit OpenAL on GNU/Linux)
    alSourcef psSample.iSample, AL_PITCH, 1!

    sound_SetObjectPosition psSample
    alSourcefv psSample.iSample, AL_VELOCITY, zero
    alSourcei psSample.iSample, AL_BUFFER, psTrack.iBufferName
    If sound_SetupChannel(psSample) Then
        alSourcei psSample.iSample, AL_LOOPING, AL_TRUE
    Else
        alSourcei psSample.iSample, AL_LOOPING, AL_FALSE
    End If

    ' NOTE: this is only useful for debugging. (#ifdef DEBUG: is3d, isLooping, filename)

    ' Clear error codes
    alGetError

    alSourcePlay psSample.iSample
    sound_GetError "lib/sound/openal_track.cpp:859"

    sound_Play3DSample = True
End Function

' Fills N buffers, each of buffSize
' Only frees whatever was allocated by itself.
' \returns nb of buffers copied (but not enqueued yet) to openAL
' (VB) alBuffersIds(first ...): the C's alBuffersIds pointer
Private Function sound_fillNBuffers(alBuffersIds() As Long, ByVal first As Long, ByVal decoder As WZDecoder, ByVal n As Long, ByVal buffSize As Long) As Long
    Static pcm() As Byte, pcmSize As Long
    Dim fmt_ As Long, i As Long, i_max As Long, written As Long

    If Not (CDbl(n) <= 2147483647#) Then ASSERT_FAILURE "n <= static_cast<size_t>(std::numeric_limits<int>::max())", "lib/sound/openal_track.cpp:870", "sound_fillNBuffers", "number of buffers (%zu) exceeds int::max", n: sound_fillNBuffers = -1: Exit Function
    ' static uint8_t *pcm = (uint8_t*) malloc(buffSize): allocated at the first call
    ' (VB: ReDim raises an error where malloc returns nullptr: no "can't allocate buff of size %zu" here)
    If pcmSize = 0 Then
        pcmSize = buffSize
        ReDim pcm(0 To pcmSize - 1)
    End If
    ' Determine PCM data format
    If decoder.channels() = 1 Then fmt_ = AL_FORMAT_MONO16 Else fmt_ = AL_FORMAT_STEREO16
    i = 0
    i_max = n
    Do While i < i_max
        ReDim pcm(0 To pcmSize - 1)     ' memset(pcm, 0, buffSize)
        ' Decode some audio data
        If Not decoder.decode(pcm, buffSize, written) Then
            'free(pcm);
            sound_GetError "lib/sound/openal_track.cpp:890"
            sound_fillNBuffers = -1
            Exit Function
        End If

        ' If we actually decoded some data
        If written > 0 Then
            If Not (CDbl(written) <= 2147483647#) Then ASSERT_FAILURE "res.value() <= static_cast<size_t>(std::numeric_limits<ALint>::max())", "lib/sound/openal_track.cpp:897", "sound_fillNBuffers", "read size (%zu) exceeds ALint::max", written
            alBufferData alBuffersIds(first + i), fmt_, pcm, written, decoder.frequency()
            sound_GetError "lib/sound/openal_track.cpp:899"
        Else ' if (res.value() == 0)
            ' If no data has been decoded we're probably at the end of our
            ' stream. So cleanup the excess stuff here.
            ' First remove the data buffer itself
            'free(pcm);
            sound_GetError "lib/sound/openal_track.cpp:907"
            Exit Do
        End If
        i = i + 1
    Loop
    ' all good: return how many buffers were actually (at least partially) filled
    sound_fillNBuffers = i
End Function

' Plays the audio data from the given file
' \param volume the volume to play the audio at (in a range of 0.0 to 1.0)
' \param onFinished callback to invoke when we're finished playing
' \param user_data user-data pointer to pass to the \c onFinished callback
' \return a pointer to the currently playing stream when playing started
'         successfully, NULL otherwise.
' \note The returned pointer will become invalid/dangling immediately after
'       the \c onFinished callback is invoked.
' \note You must _never_ manually free() the memory used by the returned
'       pointer.
Public Function sound_PlayStream(ByVal fileName As String, ByVal bufferEntireStream As Boolean, ByVal volume As Single, _
                                 ByVal onFinished As IStreamFinished, ByVal user_data As Variant) As AUDIO_STREAM
    Static alBuffersIds() As Long, idsReady As Boolean
    Dim decoder As WZDecoder, length As Long, stream As AUDIO_STREAM, ids(0 To 0) As Long, res As Long, i As Long
    staticsInit

    If Not openal_initialized Then
        debug_ 934, LOG_WARNING, "sound_PlayStream", "OpenAL isn't initialized, not creating an audio stream"
        Set sound_PlayStream = Nothing
        Exit Function
    End If
    ' Clean errors
    alGetError
    Set decoder = Nothing
    length = Len(fileName)
    If length > 4 And LCase$(Right$(fileName, 4)) = ".ogg" Then
        Set decoder = WZVorbisDecoder_fromFilename(fileName)
    ElseIf length > 5 And LCase$(Right$(fileName, 5)) = ".opus" Then
        Set decoder = WZOpusDecoder_fromFilename(fileName, bufferEntireStream)
    End If
    If decoder Is Nothing Then
        debug_ 951, LOG_ERROR, "sound_PlayStream", "couldn't allocate decoder for %s", fileName
        Set sound_PlayStream = Nothing
        Exit Function
    End If

    ' (VB: New raises an error where new returns nullptr: no "sound_PlayStream: Out of memory" here)
    Set stream = New AUDIO_STREAM
    Set stream.decoder = decoder
    ' Retrieve an OpenAL sound source
    ids(0) = stream.source
    alGenSources 1, ids
    stream.source = ids(0)
    If sound_GetError("lib/sound/openal_track.cpp:965") <> AL_NO_ERROR Then
        ' Failed to create OpenAL sound source, so bail out...
        debug_ 968, LOG_SOUND, "sound_PlayStream", "alGenSources failed, most likely out of sound sources"
        ' (delete stream; delete decoder)
        Set sound_PlayStream = Nothing
        Exit Function
    End If

    stream.volume = volume
    alSourcef stream.source, AL_GAIN, stream.volume

    ' (WZ_OS_UNIX: AL_PITCH 1.001f, a workaround for 64-bit OpenAL on GNU/Linux)
    alSourcef stream.source, AL_PITCH, 1!
    If sound_GetError("lib/sound/openal_track.cpp:984") <> AL_NO_ERROR Then
        ' (delete decoder; delete stream)
        Set sound_PlayStream = Nothing
        Exit Function
    End If
    res = 0
    ' Create some OpenAL buffers to store the decoded data in
    ' static ALuint *alBuffersIds = (ALuint *) malloc(buffer_count * sizeof(ALuint)): at the first call
    If Not idsReady Then
        ReDim alBuffersIds(0 To buffer_count - 1)
        idsReady = True
    End If
    For i = 0 To buffer_count - 1       ' memset(alBuffersIds, 0, ...)
        alBuffersIds(i) = 0
    Next
    alGenBuffers buffer_count, alBuffersIds
    If sound_GetError("lib/sound/openal_track.cpp:995") <> AL_NO_ERROR Then GoTo error_
    ' Copy the audio data into one of OpenAL's own buffers
    If Not (CDbl(bufferSize) <= 2147483647#) Then ASSERT_FAILURE "bufferSize <= static_cast<size_t>(std::numeric_limits<ALsizei>::max())", "lib/sound/openal_track.cpp:997", "sound_PlayStream", "soundBuffer->size (%zu) exceeds ALsizei::max", bufferSize

    res = sound_fillNBuffers(alBuffersIds, 0, decoder, buffer_count, bufferSize)
    ' Bail out if we didn't fill any buffers
    If res <= 0 Then
        debug_ 1003, LOG_ERROR, "sound_PlayStream", "Failed to fill buffers with decoded audio data!"
        GoTo error_with_albuffers
    End If
    ' Attach the OpenAL buffers to our OpenAL source
    alGetError
    If res < buffer_count Then
        ' free unused buffers
        debug_ 1011, LOG_SOUND, "sound_PlayStream", "freeing unused %i buffers", buffer_count - res
        alDeleteBuffers buffer_count - res, alBuffersIds, res
        If sound_GetError("lib/sound/openal_track.cpp:1013") <> AL_NO_ERROR Then GoTo error_with_albuffers
    End If

    alSourceQueueBuffers stream.source, res, alBuffersIds
    If sound_GetError("lib/sound/openal_track.cpp:1017") <> AL_NO_ERROR Then GoTo error_with_albuffers

    ' Start playing the source
    alGetError
    alSourcePlay stream.source
    If sound_GetError("lib/sound/openal_track.cpp:1022") <> AL_NO_ERROR Then GoTo error_with_albuffers

    ' Set callback info
    Set stream.onFinished = onFinished
    If IsObject(user_data) Then Set stream.user_data = user_data Else stream.user_data = user_data

    ' Prepend this stream to the linked list
    listPushFront active_streams, stream

    Set sound_PlayStream = stream
    Exit Function

error_with_albuffers:
    alDeleteBuffers buffer_count, alBuffersIds

error_:
    ids(0) = stream.source
    alDeleteSources 1, ids
    ' (delete stream; delete decoder)
    Set sound_PlayStream = Nothing
End Function

' Checks if the stream is playing.
' \param stream the stream to check
' \post true if playing, false otherwise.
Public Function sound_isStreamPlaying(ByVal stream As AUDIO_STREAM) As Boolean
    Dim state As Long
    staticsInit
    alGetError
    If Not stream Is Nothing Then
        alGetSourcei stream.source, AL_SOURCE_STATE, state
        sound_GetError "lib/sound/openal_track.cpp:1055"
        If state = AL_PLAYING Then
            sound_isStreamPlaying = True
            Exit Function
        End If
    End If
    sound_isStreamPlaying = False
End Function

' Stops the current stream from playing.
' \param stream the stream to stop
' \post The stopped stream will be destroyed on the next invocation of
'       sound_UpdateStreams(). So calling this function will result in the
'       \c onFinished callback being called and the \c stream pointer becoming
'       invalid.
Public Sub sound_StopStream(ByVal stream As AUDIO_STREAM)
    staticsInit
    ' assert(stream != nullptr): the C assert, off in a release build

    alGetError  ' clear error codes
    stream.queuedStop = True
    ' Tell OpenAL to stop playing on the given source
    alSourceStop stream.source
    sound_GetError "lib/sound/openal_track.cpp:1079"
End Sub

' Pauses playing of this stream until playing is resumed with
' sound_ResumeStream() or completely stopped with sound_StopStream().
' \param stream the stream to pause playing for
Public Sub sound_PauseStream(ByVal stream As AUDIO_STREAM)
    Dim state As Long
    staticsInit

    ' To be sure we won't go mutilating this OpenAL source, check whether
    ' it's playing first.
    alGetError
    alGetSourcei stream.source, AL_SOURCE_STATE, state
    sound_GetError "lib/sound/openal_track.cpp:1094"

    If state <> AL_PLAYING Then
        Exit Sub
    End If

    ' Pause playing of this OpenAL source
    alSourcePause stream.source
    sound_GetError "lib/sound/openal_track.cpp:1103"
End Sub

' Resumes playing of a stream that's paused by means of sound_PauseStream().
' \param stream the stream to resume playing for
Public Sub sound_ResumeStream(ByVal stream As AUDIO_STREAM)
    Dim state As Long
    staticsInit

    ' To be sure we won't go mutilating this OpenAL source, check whether
    ' it's paused first.
    alGetError
    alGetSourcei stream.source, AL_SOURCE_STATE, state
    sound_GetError "lib/sound/openal_track.cpp:1117"

    If state <> AL_PAUSED Then
        Exit Sub
    End If

    ' Resume playing of this OpenAL source
    alSourcePlay stream.source
    sound_GetError "lib/sound/openal_track.cpp:1126"
End Sub

' Retrieve the playing volume of the given stream.
'
' @param stream the stream to retrieve the volume for.
'
' @return a floating point value between 0.f and 1.f, representing this
'         stream's volume.
Public Function sound_GetStreamVolume(ByVal stream As AUDIO_STREAM) As Single
    Dim volume As Single
    staticsInit
    alGetError
    alGetSourcef stream.source, AL_GAIN, volume
    sound_GetError "lib/sound/openal_track.cpp:1141"

    sound_GetStreamVolume = volume
End Function

' Set the playing volume of the given stream.
'
' @param stream the stream to change the volume for.
' @param volume a floating point value between 0.f and 1.f, to use as this
'               @c stream's volume.
Public Sub sound_SetStreamVolume(ByVal stream As AUDIO_STREAM, ByVal volume As Single)
    staticsInit
    alGetError
    stream.volume = volume
    alSourcef stream.source, AL_GAIN, stream.volume
    sound_GetError "lib/sound/openal_track.cpp:1157"
End Sub

Public Function sound_GetStreamTotalTime(ByVal stream As AUDIO_STREAM) As Double
    Dim decoder As WZDecoder
    Set decoder = stream.decoder
    sound_GetStreamTotalTime = decoder.totalTime()
End Function

' Update the given stream (="alSource" in openAL parlance) by making sure its buffers remain full
' \param stream the stream to update
' \return true when the stream is still playing, false when it has stopped
Private Function sound_UpdateStream(ByVal stream As AUDIO_STREAM) As Boolean
    Dim state As Long, buffers_processed_count As Long, alBuffersIds() As Long, res As Long, decoder As WZDecoder
    alGetError
    alGetSourcei stream.source, AL_SOURCE_STATE, state
    sound_GetError "lib/sound/openal_track.cpp:1174"

    If state <> AL_PLAYING And state <> AL_PAUSED And (state <> AL_STOPPED Or stream.queuedStop) Then
        sound_UpdateStream = False
        Exit Function
    End If

    ' Retrieve the amount of buffers which were processed and need refilling
    alGetSourcei stream.source, AL_BUFFERS_PROCESSED, buffers_processed_count
    If sound_GetError("lib/sound/openal_track.cpp:1183") <> AL_NO_ERROR Then sound_UpdateStream = False: Exit Function
    If buffers_processed_count = 0 Then sound_UpdateStream = True: Exit Function

    ' Determine PCM data format
    ReDim alBuffersIds(0 To buffers_processed_count - 1)
    alSourceUnqueueBuffers stream.source, buffers_processed_count, alBuffersIds
    If sound_GetError("lib/sound/openal_track.cpp:1189") <> AL_NO_ERROR Then sound_UpdateStream = False: Exit Function

    Set decoder = stream.decoder
    res = sound_fillNBuffers(alBuffersIds, 0, decoder, buffers_processed_count, bufferSize)
    If res = 0 Then
        ' nothing more to read and queue - will be deleted with sound_DestroyStream (later, when done playing)
        sound_GetError "lib/sound/openal_track.cpp:1202"
        freeUnusedALBuffers alBuffersIds, buffers_processed_count, 0
        ' (free(alBuffersIds))
        If state <> AL_STOPPED Then
            sound_UpdateStream = True ' must return true here - don't shortcut playing the remaining buffers!
        Else
            ' no more buffers to read, and the existing audio has stopped
            sound_UpdateStream = False
        End If
        Exit Function
    End If
    If res < 0 Then
        debug_ 1217, LOG_ERROR, "sound_UpdateStream", "bailing out"
        freeUnusedALBuffers alBuffersIds, buffers_processed_count, 0
        ' (free(alBuffersIds))
        sound_UpdateStream = False
        Exit Function
    End If
    ' Reattach the filled buffers to the source
    alSourceQueueBuffers stream.source, res, alBuffersIds
    ' Delete any unused unqueued buffers
    freeUnusedALBuffers alBuffersIds, buffers_processed_count, res
    sound_GetError "lib/sound/openal_track.cpp:1226"
    ' (free(alBuffersIds))

    If state = AL_STOPPED And Not stream.queuedStop Then
        ' Resume playing of this OpenAL source
        debug_ 1232, LOG_SOUND, "sound_UpdateStream", "Auto-resuming play of stream"
        alSourcePlay stream.source
        If sound_GetError("lib/sound/openal_track.cpp:1234") <> AL_NO_ERROR Then sound_UpdateStream = False: Exit Function
    End If

    sound_UpdateStream = True
End Function

' sound_UpdateStream's lambda freeUnusedALBuffers [alBuffersIds, buffers_processed_count](ALint startingIdx)
Private Sub freeUnusedALBuffers(alBuffersIds() As Long, ByVal buffers_processed_count As Long, ByVal startingIdx As Long)
    If startingIdx < buffers_processed_count Then
        alDeleteBuffers buffers_processed_count - startingIdx, alBuffersIds, startingIdx
    End If
End Sub

' Destroy the given stream (="alSource" in openAL parlance) and release its associated resources. This function
' also handles calling of the \c onFinished callback function and closing of
' the PhysicsFS file handle.
' \param stream the stream to destroy
' \note  we are not exiting on errors, we don't want to leak memory...
Private Sub sound_DestroyStream(ByVal stream As AUDIO_STREAM)
    Dim buffers_processed_count As Long, buffers() As Long, ids(0 To 0) As Long

    ' Stop the OpenAL source from playing
    alGetError
    alSourceStop stream.source

    sound_GetError "lib/sound/openal_track.cpp:1255"

    ' Retrieve the amount of buffers which were processed
    alGetSourcei stream.source, AL_BUFFERS_PROCESSED, buffers_processed_count
    If sound_GetError("lib/sound/openal_track.cpp:1259") <> AL_NO_ERROR Then
        buffers_processed_count = 0
        ' proceed as if nothing happened
    End If

    ' Detach all buffers and retrieve their ID numbers
    If buffers_processed_count > 0 Then
        ReDim buffers(0 To buffers_processed_count - 1)     ' malloc and memset
        alSourceUnqueueBuffers stream.source, buffers_processed_count, buffers
        sound_GetError "lib/sound/openal_track.cpp:1272"

        ' Destroy all of these buffers
        alDeleteBuffers buffers_processed_count, buffers
        sound_GetError "lib/sound/openal_track.cpp:1276"
        ' (free(buffers))
    Else
        debug_ 1281, LOG_SOUND, "sound_DestroyStream", "alGetSourcei(AL_BUFFERS_PROCESSED) returned count: %d", buffers_processed_count
    End If

    ' Destroy the OpenAL source
    ids(0) = stream.source
    alDeleteSources 1, ids
    sound_GetError "lib/sound/openal_track.cpp:1286"

    ' Destroy the sound decoder
    Set stream.decoder = Nothing

    ' Now call the finished callback
    If Not stream.onFinished Is Nothing Then
        If Not (isNullUserData(stream.user_data)) Then ASSERT_FAILURE "stream->user_data == nullptr", "lib/sound/openal_track.cpp:1294", "sound_DestroyStream", "user_data was not null!"
        stream.onFinished.streamFinished stream, stream.user_data
    End If

    ' Free the memory used by this stream (delete stream: the caller's reference is the last)
End Sub

' Update all currently running streams(="alSource"s) and destroy them when they're finished.
Private Sub sound_UpdateStreams()
    Dim it As AUDIO_STREAM, nxt As AUDIO_STREAM
    Set it = listFirst(active_streams)
    Do While Not it Is Nothing
        Set nxt = listAfter(active_streams, it)
        ' Attempt to update the current stream, if we find that impossible,
        ' destroy it instead.
        If Not sound_UpdateStream(it) Then
            ' First remove our current stream from the linked list
            listRemove active_streams, it
            ' Now actually destroy the current stream
            sound_DestroyStream it
        End If
        Set it = nxt
    Loop
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_StopSample(ByVal psSample As AUDIO_SAMPLE)
    staticsInit
    If psSample.iSample = SAMPLE_NOT_ALLOCATED Then
        debug_ 1330, LOG_SOUND, "sound_StopSample", "sound_StopSample: sample number (%u) out of range, we probably have run out of available OpenAL sources", psSample.iSample
        Exit Sub
    End If
    alGetError  ' clear error codes
    ' Tell OpenAL to stop playing the given sample
    alSourceStop psSample.iSample
    sound_GetError "lib/sound/openal_track.cpp:1336"
End Sub

' Vector3f pos
Public Sub sound_SetPlayerPos(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    staticsInit
    alListener3f AL_POSITION, x, y, z
    sound_GetError "lib/sound/openal_track.cpp:1342"
End Sub

'
' Sets the player's orientation to use for sound
' \param angle the angle in radians
' @NOTE the up vector is swapped because of qsound idiosyncrasies
' @FIXME we don't use qsound, but it still is in qsound 'format'...
Public Sub sound_SetPlayerOrientation(ByVal angle As Single)
    Dim ori(0 To 5) As Single
    staticsInit
    ori(0) = -CSng(Sin(angle))          ' forward (at) vector
    ori(1) = CSng(Cos(angle))
    ori(2) = 0!
    ori(3) = 0!                         ' up vector
    ori(4) = 0!
    ori(5) = 1!
    alGetError
    alListenerfv AL_ORIENTATION, ori
    sound_GetError "lib/sound/openal_track.cpp:1360"
End Sub

'
' Sets the player's orientation to use for sound
' \param forward forward pointing vector
' \param up      upward pointing vector
Public Sub sound_SetPlayerOrientationVector(ByVal fx As Single, ByVal fy As Single, ByVal fz As Single, _
                                            ByVal ux As Single, ByVal uy As Single, ByVal uz As Single)
    Dim ori(0 To 5) As Single
    staticsInit
    ori(0) = fx: ori(1) = fy: ori(2) = fz
    ori(3) = ux: ori(4) = uy: ori(5) = uz
    alGetError
    alListenerfv AL_ORIENTATION, ori
    sound_GetError "lib/sound/openal_track.cpp:1377"
End Sub

'*
' =======================================================================================================================
' Compute the sample's volume relative to AL_POSITION.
' =======================================================================================================================
'
Public Sub sound_SetObjectPosition(ByVal psSample As AUDIO_SAMPLE)
    ' coordinates
    Dim listenerX As Single, listenerY As Single, listenerZ As Single, dX As Single, dY As Single, dZ As Single

    ' calculation results
    Dim distance As Single, gain As Single
    staticsInit

    ' only set it when we have a valid sample
    If psSample Is Nothing Then
        Exit Sub
    End If
    alGetError
    ' compute distance
    alGetListener3f AL_POSITION, listenerX, listenerY, listenerZ
    sound_GetError "lib/sound/openal_track.cpp:1403"
    dX = CSng(psSample.x) - listenerX   ' distances on all axis
    dY = CSng(psSample.y) - listenerY
    dZ = CSng(psSample.z) - listenerZ
    distance = CSng(Sqr(dX * dX + dY * dY + dZ * dZ)) ' Pythagorean theorem

    ' compute gain
    gain = (1! - (distance * ATTENUATION_FACTOR)) * psSample.fVol * sfx3d_volume
    ' max volume
    If gain > 1! Then
        gain = 1!
    End If
    If gain < 0! Then
        ' this sample can't be heard right now
        gain = 0!
    End If
    alSourcef psSample.iSample, AL_GAIN, gain

    ' the alSource3i variant would be better, if it wouldn't provide linker errors however
    alSource3f psSample.iSample, AL_POSITION, CSng(psSample.x), CSng(psSample.y), CSng(psSample.z)
    sound_GetError "lib/sound/openal_track.cpp:1425"
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_PauseSample(ByVal psSample As AUDIO_SAMPLE)
    staticsInit
    alGetError
    alSourcePause psSample.iSample
    sound_GetError "lib/sound/openal_track.cpp:1438"
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Sub sound_ResumeSample(ByVal psSample As AUDIO_SAMPLE)
    staticsInit
    alGetError
    alSourcePlay psSample.iSample
    sound_GetError "lib/sound/openal_track.cpp:1449"
End Sub

'*
' =======================================================================================================================
' =======================================================================================================================
'
Public Function sound_SampleIsFinished(ByVal psSample As AUDIO_SAMPLE) As Boolean
    Dim state As Long, ids(0 To 0) As Long
    staticsInit
    alGetError
    alGetSourcei psSample.iSample, AL_SOURCE_STATE, state
    sound_GetError "lib/sound/openal_track.cpp:1461" ' check for an error and clear the error state for later on in this function
    If state = AL_PLAYING Or state = AL_PAUSED Then
        sound_SampleIsFinished = False
        Exit Function
    End If

    If psSample.iSample <> AL_INVALID Then
        ids(0) = psSample.iSample
        alDeleteSources 1, ids
        sound_GetError "lib/sound/openal_track.cpp:1470"
        psSample.iSample = AL_INVALID
    End If

    sound_SampleIsFinished = True
End Function

'*
' =======================================================================================================================
' =======================================================================================================================
'

Public Function sound_GetUIVolume() As Single
    staticsInit
    sound_GetUIVolume = sfx_volume
End Function

Public Sub sound_SetUIVolume(ByVal volume As Single)
    staticsInit
    sfx_volume = volume

    ' Keep volume in the range of 0.0 - 1.0
    If sfx_volume < 0# Then
        sfx_volume = 0#
    ElseIf sfx_volume > 1# Then
        sfx_volume = 1#
    End If
End Sub

Public Function sound_GetEffectsVolume() As Single
    staticsInit
    sound_GetEffectsVolume = sfx3d_volume
End Function

Public Sub sound_SetEffectsVolume(ByVal volume As Single)
    staticsInit
    sfx3d_volume = volume

    ' Keep volume in the range of 0.0 - 1.0
    If sfx3d_volume < 0# Then
        sfx3d_volume = 0#
    ElseIf sfx3d_volume > 1# Then
        sfx3d_volume = 1#
    End If
End Sub

' (VB) a test aid: before sound_InitLibrary, the mix goes to this WAV file and no device opens ("": the
' device again); the file grows as fast as the real time passes (al_win)
Public Sub sound_vbSetCapture(ByVal wavPath As String)
    al_vbSetCapture wavPath
End Sub
