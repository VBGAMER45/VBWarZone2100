Attribute VB_Name = "al_win"
' (VB) OpenAL (al.h, alc.h, OpenAL Soft's alext.h names) as far as lib/sound/openal_track.cpp and
' openal_info.cpp call it, over the Windows waveOut device: one device, one context. WZ 4.7 sets
' alDistanceModel(AL_NONE) and computes every gain itself (sound_SetObjectPosition), so this mixes,
' resamples and pans:
' * the output: 44100 Hz, stereo, 16 bit; a ring of BLOCKS blocks of BLOCK_FRAMES frames (about 93
'   ms), polled as QuakeVB's snd_win.bas polls it: alcProcessContext (sound_Update calls it every
'   frame) mixes and writes again each block the device has played. A source advances as it is
'   mixed, ahead of what is heard by up to the ring's length; while nothing calls
'   alcProcessContext (a long load) the device runs dry and the sources wait.
' * the resampler: linear interpolation (OpenAL Soft's "linear"), with AL_PITCH;
' * a mono source is panned from its direction in the listener's frame, with equal power: left =
'   sqrt((1 - pan) / 2), right = sqrt((1 + pan) / 2), pan = the unit direction's component along the
'   listener's right (sin(azimuth) * cos(elevation): OpenAL Soft's first-order panning on a stereo
'   device); a source at the listener (AL_SOURCE_RELATIVE at 0, 0, 0) is centred; a relative
'   source's position is already in the listener's frame (right +x, up +y, front -z); stereo
'   buffers are not panned;
' * the effective gain is clamped to 0 .. 1 (AL_MIN_GAIN, AL_MAX_GAIN); AL_INVERSE_DISTANCE_CLAMPED
'   (the default, before WZ's alDistanceModel(AL_NONE)) has reference distance 1 and rolloff 1;
' * no extensions: no ALC_SOFT_HRTF (sound_GetHRTFMode: Unsupported), no ALC_ENUMERATE_ALL_EXT;
'   alcGetProcAddress finds nothing.
' Test aids (no device): al_vbSetCapture(path) before alcOpenDevice writes the mix to a WAV file
' instead, as much of it as the real time since the device opened asks for at each
' alcProcessContext; al_vbCaptureMix mixes a given number of frames into it now (src/mixertests);
' al_vbMixBench mixes into nothing (the mixer's cost).
' ALuint, ALint, ALenum, ALsizei, ALCdevice * and ALCcontext * are Long (the device and the
' context are 1, 0 = NULL), ALfloat Single; a C pointer to names or values is a VB array and the
' index of the first element. A name is its slot's index + 1 (0 is AL_NONE).
Option Explicit

' al.h
Public Const AL_NONE As Long = 0
Public Const AL_FALSE As Long = 0
Public Const AL_TRUE As Long = 1
Public Const AL_SOURCE_RELATIVE As Long = &H202
Public Const AL_PITCH As Long = &H1003
Public Const AL_POSITION As Long = &H1004
Public Const AL_DIRECTION As Long = &H1005
Public Const AL_VELOCITY As Long = &H1006
Public Const AL_LOOPING As Long = &H1007
Public Const AL_BUFFER As Long = &H1009
Public Const AL_GAIN As Long = &H100A
Public Const AL_ORIENTATION As Long = &H100F
Public Const AL_SOURCE_STATE As Long = &H1010
Public Const AL_INITIAL As Long = &H1011
Public Const AL_PLAYING As Long = &H1012
Public Const AL_PAUSED As Long = &H1013
Public Const AL_STOPPED As Long = &H1014
Public Const AL_BUFFERS_QUEUED As Long = &H1015
Public Const AL_BUFFERS_PROCESSED As Long = &H1016
Public Const AL_SOURCE_TYPE As Long = &H1027
Public Const AL_STATIC As Long = &H1028
Public Const AL_STREAMING As Long = &H1029
Public Const AL_UNDETERMINED As Long = &H1030
Public Const AL_FORMAT_MONO8 As Long = &H1100
Public Const AL_FORMAT_MONO16 As Long = &H1101
Public Const AL_FORMAT_STEREO8 As Long = &H1102
Public Const AL_FORMAT_STEREO16 As Long = &H1103
Public Const AL_NO_ERROR As Long = 0
Public Const AL_INVALID As Long = -1
Public Const AL_INVALID_NAME As Long = &HA001&
Public Const AL_INVALID_ENUM As Long = &HA002&
Public Const AL_INVALID_VALUE As Long = &HA003&
Public Const AL_INVALID_OPERATION As Long = &HA004&
Public Const AL_OUT_OF_MEMORY As Long = &HA005&
Public Const AL_VENDOR As Long = &HB001&
Public Const AL_VERSION As Long = &HB002&
Public Const AL_RENDERER As Long = &HB003&
Public Const AL_EXTENSIONS As Long = &HB004&
Public Const AL_INVERSE_DISTANCE As Long = &HD001&
Public Const AL_INVERSE_DISTANCE_CLAMPED As Long = &HD002&

' alc.h
Public Const ALC_FALSE As Long = 0
Public Const ALC_TRUE As Long = 1
Public Const ALC_FREQUENCY As Long = &H1007
Public Const ALC_REFRESH As Long = &H1008
Public Const ALC_NO_ERROR As Long = 0
Public Const ALC_INVALID_DEVICE As Long = &HA001&
Public Const ALC_INVALID_CONTEXT As Long = &HA002&
Public Const ALC_INVALID_ENUM As Long = &HA003&
Public Const ALC_INVALID_VALUE As Long = &HA004&
Public Const ALC_OUT_OF_MEMORY As Long = &HA005&
Public Const ALC_MAJOR_VERSION As Long = &H1000
Public Const ALC_MINOR_VERSION As Long = &H1001
Public Const ALC_DEFAULT_DEVICE_SPECIFIER As Long = &H1004
Public Const ALC_DEVICE_SPECIFIER As Long = &H1005
Public Const ALC_EXTENSIONS As Long = &H1006
Public Const ALC_DEFAULT_ALL_DEVICES_SPECIFIER As Long = &H1012
Public Const ALC_ALL_DEVICES_SPECIFIER As Long = &H1013

' alext.h: ALC_SOFT_HRTF (never present here)
Public Const ALC_DONT_CARE_SOFT As Long = &H2
Public Const ALC_HRTF_SOFT As Long = &H1992
Public Const ALC_HRTF_STATUS_SOFT As Long = &H1993
Public Const ALC_HRTF_DISABLED_SOFT As Long = &H0
Public Const ALC_HRTF_ENABLED_SOFT As Long = &H1
Public Const ALC_HRTF_DENIED_SOFT As Long = &H2
Public Const ALC_HRTF_REQUIRED_SOFT As Long = &H3
Public Const ALC_HRTF_HEADPHONES_DETECTED_SOFT As Long = &H4
Public Const ALC_HRTF_UNSUPPORTED_FORMAT_SOFT As Long = &H5

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    reserved As Long
End Type

Private Type WAVEOUTCAPSW
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname(0 To 63) As Byte                ' WCHAR[32]
    dwFormats As Long
    wChannels As Integer
    wReserved1 As Integer
    dwSupport As Long
End Type

Private Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, _
    lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hwo As Long, ByVal pwh As Long, ByVal cbwh As Long) As Long
Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hwo As Long, ByVal pwh As Long, ByVal cbwh As Long) As Long
Private Declare Function waveOutWrite Lib "winmm.dll" (ByVal hwo As Long, ByVal pwh As Long, ByVal cbwh As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hwo As Long) As Long
Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hwo As Long) As Long
Private Declare Function waveOutGetDevCapsW Lib "winmm.dll" (ByVal uDeviceID As Long, pwoc As WAVEOUTCAPSW, ByVal cbwoc As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, _
    ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const WAVE_MAPPER As Long = -1
Private Const WAVE_FORMAT_PCM As Long = 1
Private Const CALLBACK_NULL As Long = 0
Private Const MMSYSERR_NOERROR As Long = 0
Private Const WHDR_DONE As Long = 1
Private Const GENERIC_WRITE As Long = &H40000000
Private Const CREATE_ALWAYS As Long = 2
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_BEGIN As Long = 0
Private Const FILE_END As Long = 2

Private Const MIX_RATE As Long = 44100      ' ALC_FREQUENCY
Private Const BLOCK_FRAMES As Long = 512    ' stereo frames of a waveOut block (11.6 ms)
Private Const BLOCKS As Long = 8            ' the ring: 93 ms
Private Const MAX_SOURCES As Long = 256     ' OpenAL Soft's default source limit
Private Const AL_DEVICE As Long = 1         ' the one ALCdevice *
Private Const AL_CONTEXT As Long = 1        ' the one ALCcontext *

' a buffer: 16-bit PCM, interleaved when stereo
Private Type AL_BUFFER_T
    used As Boolean
    channels As Long
    freq As Long
    frames As Long
    data() As Integer
    refs As Long                            ' the queue entries that hold it (in use when > 0)
End Type

Private Type AL_SOURCE_T
    used As Boolean
    state As Long                           ' AL_INITIAL, AL_PLAYING, AL_PAUSED, AL_STOPPED
    gain As Single
    pitch As Single
    pos(0 To 2) As Single
    vel(0 To 2) As Single
    relative As Boolean
    looping As Boolean
    sourceType As Long                      ' AL_UNDETERMINED, AL_STATIC, AL_STREAMING
    queue() As Long                         ' buffer names: queue(0 .. queueCount - 1)
    queueCount As Long
    current As Long                         ' the queue entry playing: those before it are processed
    offset As Double                        ' the frame in the current buffer (fractional: the resampler)
End Type

Private alError As Long                     ' the context's error, kept until alGetError reads it
Private alcError As Long                    ' the device's
Private deviceOpen As Boolean
Private contextAlive As Boolean
Private contextCurrent As Boolean
Private openedName As String

Private bufs() As AL_BUFFER_T
Private bufCap As Long
Private srcs(0 To MAX_SOURCES - 1) As AL_SOURCE_T

Private listenerPos(0 To 2) As Single
Private listenerVel(0 To 2) As Single
Private listenerAt(0 To 2) As Single
Private listenerUp(0 To 2) As Single
Private listenerGain As Single
Private distModel As Long

' the waveOut ring
Private hWaveOut As Long
Private ringHdr(0 To BLOCKS - 1) As WAVEHDR
Private ringPcm() As Integer
Private nextBlock As Long                   ' the block the device plays next

' the capture
Private capturePath As String
Private captureFile As Long                 ' the WAV file's handle (0: none)
Private captureFrames As Double             ' frames written
Private captureStart As Currency            ' QueryPerformanceCounter when the device opened
Private captureBuf() As Integer

' the mixer's accumulators
Private accL() As Single
Private accR() As Single
Private accSize As Long

'--------------------------------------------------------------------------
' errors

Private Sub setError(ByVal e As Long)
    If alError = AL_NO_ERROR Then alError = e
End Sub

Private Sub setContextError(ByVal e As Long)
    If alcError = ALC_NO_ERROR Then alcError = e
End Sub

Public Function alGetError() As Long
    If Not contextCurrent Then
        alGetError = AL_INVALID_OPERATION   ' OpenAL Soft: no current context
        Exit Function
    End If
    alGetError = alError
    alError = AL_NO_ERROR
End Function

Public Function alGetString(ByVal param As Long) As String
    Select Case param
    Case AL_NO_ERROR: alGetString = "No Error"
    Case AL_INVALID_NAME: alGetString = "Invalid Name"
    Case AL_INVALID_ENUM: alGetString = "Invalid Enum"
    Case AL_INVALID_VALUE: alGetString = "Invalid Value"
    Case AL_INVALID_OPERATION: alGetString = "Invalid Operation"
    Case AL_OUT_OF_MEMORY: alGetString = "Out of Memory"
    Case AL_VENDOR: alGetString = "WarzoneVB"
    Case AL_VERSION: alGetString = "1.1 WarzoneVB waveOut mixer"
    Case AL_RENDERER: alGetString = "WarzoneVB waveOut mixer (44100 Hz stereo, linear resampler)"
    Case AL_EXTENSIONS: alGetString = ""
    Case Else
        setError AL_INVALID_ENUM
        alGetString = vbNullString
    End Select
End Function

'--------------------------------------------------------------------------
' names

' the slot of a source name, -1 if it is not one (AL_INVALID_NAME)
Private Function srcSlot(ByVal name_ As Long) As Long
    srcSlot = -1
    If name_ < 1 Or name_ > MAX_SOURCES Then Exit Function
    If srcs(name_ - 1).used Then srcSlot = name_ - 1
End Function

Private Function bufSlot(ByVal name_ As Long) As Long
    bufSlot = -1
    If name_ < 1 Or name_ > bufCap Then Exit Function
    If bufs(name_ - 1).used Then bufSlot = name_ - 1
End Function

Public Function alIsSource(ByVal source As Long) As Long
    If Not contextCurrent Then Exit Function
    If srcSlot(source) >= 0 Then alIsSource = AL_TRUE Else alIsSource = AL_FALSE
End Function

'--------------------------------------------------------------------------
' sources

Public Sub alGenSources(ByVal n As Long, sources() As Long, Optional ByVal first As Long = 0)
    Dim i As Long, s As Long, free_ As Long
    If Not contextCurrent Then Exit Sub
    If n < 0 Then setError AL_INVALID_VALUE: Exit Sub
    For s = 0 To MAX_SOURCES - 1
        If Not srcs(s).used Then free_ = free_ + 1
    Next
    If free_ < n Then setError AL_OUT_OF_MEMORY: Exit Sub
    s = 0
    For i = 0 To n - 1
        Do While srcs(s).used
            s = s + 1
        Loop
        resetSource s
        srcs(s).used = True
        sources(first + i) = s + 1
    Next
End Sub

Private Sub resetSource(ByVal s As Long)
    srcs(s).state = AL_INITIAL
    srcs(s).gain = 1!
    srcs(s).pitch = 1!
    srcs(s).pos(0) = 0!: srcs(s).pos(1) = 0!: srcs(s).pos(2) = 0!
    srcs(s).vel(0) = 0!: srcs(s).vel(1) = 0!: srcs(s).vel(2) = 0!
    srcs(s).relative = False
    srcs(s).looping = False
    srcs(s).sourceType = AL_UNDETERMINED
    Erase srcs(s).queue
    srcs(s).queueCount = 0
    srcs(s).current = 0
    srcs(s).offset = 0#
End Sub

' drops the source's queue (the buffers are no longer in use by it)
Private Sub clearQueue(ByVal s As Long)
    Dim i As Long, b As Long
    For i = 0 To srcs(s).queueCount - 1
        b = bufSlot(srcs(s).queue(i))
        If b >= 0 Then bufs(b).refs = bufs(b).refs - 1
    Next
    Erase srcs(s).queue
    srcs(s).queueCount = 0
    srcs(s).current = 0
    srcs(s).offset = 0#
End Sub

Public Sub alDeleteSources(ByVal n As Long, sources() As Long, Optional ByVal first As Long = 0)
    Dim i As Long, s As Long
    If Not contextCurrent Then Exit Sub
    If n < 0 Then setError AL_INVALID_VALUE: Exit Sub
    For i = 0 To n - 1
        If srcSlot(sources(first + i)) < 0 Then setError AL_INVALID_NAME: Exit Sub
    Next
    For i = 0 To n - 1
        s = srcSlot(sources(first + i))
        If s >= 0 Then
            clearQueue s
            srcs(s).used = False
        End If
    Next
End Sub

Public Sub alSourcef(ByVal source As Long, ByVal param As Long, ByVal value As Single)
    Dim s As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    Select Case param
    Case AL_GAIN
        If value < 0! Then setError AL_INVALID_VALUE: Exit Sub
        srcs(s).gain = value
    Case AL_PITCH
        If value < 0! Then setError AL_INVALID_VALUE: Exit Sub
        srcs(s).pitch = value
    Case Else
        setError AL_INVALID_ENUM
    End Select
End Sub

Public Sub alSource3f(ByVal source As Long, ByVal param As Long, ByVal v1 As Single, ByVal v2 As Single, ByVal v3 As Single)
    Dim s As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    Select Case param
    Case AL_POSITION
        srcs(s).pos(0) = v1: srcs(s).pos(1) = v2: srcs(s).pos(2) = v3
    Case AL_VELOCITY
        srcs(s).vel(0) = v1: srcs(s).vel(1) = v2: srcs(s).vel(2) = v3
    Case Else
        setError AL_INVALID_ENUM
    End Select
End Sub

Public Sub alSource3i(ByVal source As Long, ByVal param As Long, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    alSource3f source, param, CSng(v1), CSng(v2), CSng(v3)
End Sub

Public Sub alSourcefv(ByVal source As Long, ByVal param As Long, values() As Single, Optional ByVal first As Long = 0)
    Select Case param
    Case AL_POSITION, AL_VELOCITY
        alSource3f source, param, values(first), values(first + 1), values(first + 2)
    Case Else
        alSourcef source, param, values(first)
    End Select
End Sub

Public Sub alSourcei(ByVal source As Long, ByVal param As Long, ByVal value As Long)
    Dim s As Long, b As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    Select Case param
    Case AL_BUFFER
        If srcs(s).state = AL_PLAYING Or srcs(s).state = AL_PAUSED Then setError AL_INVALID_OPERATION: Exit Sub
        If value = 0 Then
            clearQueue s
            srcs(s).sourceType = AL_UNDETERMINED
        Else
            b = bufSlot(value)
            If b < 0 Then setError AL_INVALID_VALUE: Exit Sub
            clearQueue s
            ReDim srcs(s).queue(0 To 0)
            srcs(s).queue(0) = value
            srcs(s).queueCount = 1
            bufs(b).refs = bufs(b).refs + 1
            srcs(s).sourceType = AL_STATIC
        End If
    Case AL_SOURCE_RELATIVE
        If value <> AL_FALSE And value <> AL_TRUE Then setError AL_INVALID_VALUE: Exit Sub
        srcs(s).relative = (value = AL_TRUE)
    Case AL_LOOPING
        If value <> AL_FALSE And value <> AL_TRUE Then setError AL_INVALID_VALUE: Exit Sub
        srcs(s).looping = (value = AL_TRUE)
    Case AL_PITCH, AL_GAIN
        alSourcef source, param, CSng(value)
    Case Else
        setError AL_INVALID_ENUM
    End Select
End Sub

Public Sub alGetSourcef(ByVal source As Long, ByVal param As Long, value As Single)
    Dim s As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    Select Case param
    Case AL_GAIN: value = srcs(s).gain
    Case AL_PITCH: value = srcs(s).pitch
    Case Else: setError AL_INVALID_ENUM
    End Select
End Sub

Public Sub alGetSourcei(ByVal source As Long, ByVal param As Long, value As Long)
    Dim s As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    Select Case param
    Case AL_SOURCE_STATE: value = srcs(s).state
    Case AL_BUFFERS_QUEUED: value = srcs(s).queueCount
    Case AL_BUFFERS_PROCESSED: value = processedCount(s)
    Case AL_BUFFER
        If srcs(s).current < srcs(s).queueCount Then value = srcs(s).queue(srcs(s).current) Else value = 0
    Case AL_LOOPING: If srcs(s).looping Then value = AL_TRUE Else value = AL_FALSE
    Case AL_SOURCE_RELATIVE: If srcs(s).relative Then value = AL_TRUE Else value = AL_FALSE
    Case AL_SOURCE_TYPE: value = srcs(s).sourceType
    Case Else: setError AL_INVALID_ENUM
    End Select
End Sub

' AL_BUFFERS_PROCESSED: none of a looping or AL_INITIAL source, all of a stopped one, else those
' before the one playing
Private Function processedCount(ByVal s As Long) As Long
    If srcs(s).looping Or srcs(s).state = AL_INITIAL Then Exit Function
    If srcs(s).state = AL_STOPPED Then
        processedCount = srcs(s).queueCount
    Else
        processedCount = srcs(s).current
    End If
End Function

Public Sub alSourcePlay(ByVal source As Long)
    Dim s As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    If srcs(s).queueCount = 0 Then
        srcs(s).state = AL_STOPPED          ' nothing to play
        Exit Sub
    End If
    If srcs(s).state <> AL_PAUSED Then
        ' (a playing source restarts; an initial or stopped one starts from its first buffer)
        srcs(s).current = 0
        srcs(s).offset = 0#
    End If
    srcs(s).state = AL_PLAYING
End Sub

Public Sub alSourceStop(ByVal source As Long)
    Dim s As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    If srcs(s).state <> AL_INITIAL Then
        srcs(s).state = AL_STOPPED
        srcs(s).current = srcs(s).queueCount
        srcs(s).offset = 0#
    End If
End Sub

Public Sub alSourcePause(ByVal source As Long)
    Dim s As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    If srcs(s).state = AL_PLAYING Then srcs(s).state = AL_PAUSED
End Sub

Public Sub alSourceQueueBuffers(ByVal source As Long, ByVal nb As Long, buffers() As Long, Optional ByVal first As Long = 0)
    Dim s As Long, i As Long, b As Long, b0 As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    If nb < 0 Then setError AL_INVALID_VALUE: Exit Sub
    If nb = 0 Then Exit Sub
    If srcs(s).sourceType = AL_STATIC Then setError AL_INVALID_OPERATION: Exit Sub
    ' the queue's format: its first buffer's
    b0 = -1
    If srcs(s).queueCount > 0 Then b0 = bufSlot(srcs(s).queue(0))
    For i = 0 To nb - 1
        b = bufSlot(buffers(first + i))
        If b < 0 Then setError AL_INVALID_NAME: Exit Sub
        If b0 < 0 Then b0 = b
        If bufs(b).channels <> bufs(b0).channels Or bufs(b).freq <> bufs(b0).freq Then setError AL_INVALID_VALUE: Exit Sub
    Next
    ReDim Preserve srcs(s).queue(0 To srcs(s).queueCount + nb - 1)
    For i = 0 To nb - 1
        b = bufSlot(buffers(first + i))
        srcs(s).queue(srcs(s).queueCount + i) = buffers(first + i)
        bufs(b).refs = bufs(b).refs + 1
    Next
    srcs(s).queueCount = srcs(s).queueCount + nb
    srcs(s).sourceType = AL_STREAMING
End Sub

Public Sub alSourceUnqueueBuffers(ByVal source As Long, ByVal nb As Long, buffers() As Long, Optional ByVal first As Long = 0)
    Dim s As Long, i As Long, b As Long
    If Not contextCurrent Then Exit Sub
    s = srcSlot(source)
    If s < 0 Then setError AL_INVALID_NAME: Exit Sub
    If nb < 0 Then setError AL_INVALID_VALUE: Exit Sub
    If nb = 0 Then Exit Sub
    If srcs(s).sourceType <> AL_STREAMING Then setError AL_INVALID_VALUE: Exit Sub
    If nb > processedCount(s) Then setError AL_INVALID_VALUE: Exit Sub
    For i = 0 To nb - 1
        buffers(first + i) = srcs(s).queue(i)
        b = bufSlot(srcs(s).queue(i))
        If b >= 0 Then bufs(b).refs = bufs(b).refs - 1
    Next
    For i = nb To srcs(s).queueCount - 1
        srcs(s).queue(i - nb) = srcs(s).queue(i)
    Next
    srcs(s).queueCount = srcs(s).queueCount - nb
    srcs(s).current = srcs(s).current - nb
    If srcs(s).current < 0 Then srcs(s).current = 0
    If srcs(s).queueCount = 0 Then
        Erase srcs(s).queue
        srcs(s).sourceType = AL_UNDETERMINED
    Else
        ReDim Preserve srcs(s).queue(0 To srcs(s).queueCount - 1)
    End If
End Sub

'--------------------------------------------------------------------------
' buffers

Public Sub alGenBuffers(ByVal n As Long, buffers() As Long, Optional ByVal first As Long = 0)
    Dim i As Long, b As Long
    If Not contextCurrent Then Exit Sub
    If n < 0 Then setError AL_INVALID_VALUE: Exit Sub
    For i = 0 To n - 1
        b = 0
        Do While b < bufCap
            If Not bufs(b).used Then Exit Do
            b = b + 1
        Loop
        If b = bufCap Then
            bufCap = bufCap + 64
            ReDim Preserve bufs(0 To bufCap - 1)
        End If
        bufs(b).used = True
        bufs(b).channels = 1
        bufs(b).freq = 0
        bufs(b).frames = 0
        bufs(b).refs = 0
        Erase bufs(b).data
        buffers(first + i) = b + 1
    Next
End Sub

Public Sub alDeleteBuffers(ByVal n As Long, buffers() As Long, Optional ByVal first As Long = 0)
    Dim i As Long, b As Long
    If Not contextCurrent Then Exit Sub
    If n < 0 Then setError AL_INVALID_VALUE: Exit Sub
    For i = 0 To n - 1
        If buffers(first + i) <> 0 Then    ' (0, AL_NONE, is ignored)
            b = bufSlot(buffers(first + i))
            If b < 0 Then setError AL_INVALID_NAME: Exit Sub
            If bufs(b).refs > 0 Then setError AL_INVALID_OPERATION: Exit Sub
        End If
    Next
    For i = 0 To n - 1
        b = bufSlot(buffers(first + i))
        If b >= 0 Then
            bufs(b).used = False
            bufs(b).frames = 0
            Erase bufs(b).data
        End If
    Next
End Sub

Public Sub alBufferData(ByVal buffer As Long, ByVal fmt_ As Long, data() As Byte, ByVal size As Long, ByVal freq As Long)
    Dim b As Long, chans As Long, bytesPerSample As Long, i As Long, n As Long
    If Not contextCurrent Then Exit Sub
    b = bufSlot(buffer)
    If b < 0 Then setError AL_INVALID_NAME: Exit Sub
    If bufs(b).refs > 0 Then setError AL_INVALID_OPERATION: Exit Sub
    Select Case fmt_
    Case AL_FORMAT_MONO8: chans = 1: bytesPerSample = 1
    Case AL_FORMAT_MONO16: chans = 1: bytesPerSample = 2
    Case AL_FORMAT_STEREO8: chans = 2: bytesPerSample = 1
    Case AL_FORMAT_STEREO16: chans = 2: bytesPerSample = 2
    Case Else: setError AL_INVALID_ENUM: Exit Sub
    End Select
    If size < 0 Or freq < 1 Then setError AL_INVALID_VALUE: Exit Sub
    If size Mod (chans * bytesPerSample) <> 0 Then setError AL_INVALID_VALUE: Exit Sub
    n = size \ bytesPerSample
    bufs(b).channels = chans
    bufs(b).freq = freq
    bufs(b).frames = n \ chans
    If n = 0 Then
        Erase bufs(b).data
        Exit Sub
    End If
    ReDim bufs(b).data(0 To n - 1)
    If bytesPerSample = 2 Then
        CopyMemory bufs(b).data(0), data(0), size
    Else
        For i = 0 To n - 1
            bufs(b).data(i) = (CLng(data(i)) - 128) * 256   ' unsigned 8 bit
        Next
    End If
End Sub

'--------------------------------------------------------------------------
' the listener and the distance model

Public Sub alListener3f(ByVal param As Long, ByVal v1 As Single, ByVal v2 As Single, ByVal v3 As Single)
    If Not contextCurrent Then Exit Sub
    Select Case param
    Case AL_POSITION
        listenerPos(0) = v1: listenerPos(1) = v2: listenerPos(2) = v3
    Case AL_VELOCITY
        listenerVel(0) = v1: listenerVel(1) = v2: listenerVel(2) = v3
    Case Else
        setError AL_INVALID_ENUM
    End Select
End Sub

Public Sub alListenerfv(ByVal param As Long, values() As Single, Optional ByVal first As Long = 0)
    If Not contextCurrent Then Exit Sub
    Select Case param
    Case AL_POSITION, AL_VELOCITY
        alListener3f param, values(first), values(first + 1), values(first + 2)
    Case AL_ORIENTATION
        listenerAt(0) = values(first): listenerAt(1) = values(first + 1): listenerAt(2) = values(first + 2)
        listenerUp(0) = values(first + 3): listenerUp(1) = values(first + 4): listenerUp(2) = values(first + 5)
    Case AL_GAIN
        If values(first) < 0! Then setError AL_INVALID_VALUE: Exit Sub
        listenerGain = values(first)
    Case Else
        setError AL_INVALID_ENUM
    End Select
End Sub

Public Sub alGetListener3f(ByVal param As Long, v1 As Single, v2 As Single, v3 As Single)
    If Not contextCurrent Then Exit Sub
    Select Case param
    Case AL_POSITION
        v1 = listenerPos(0): v2 = listenerPos(1): v3 = listenerPos(2)
    Case AL_VELOCITY
        v1 = listenerVel(0): v2 = listenerVel(1): v3 = listenerVel(2)
    Case Else
        setError AL_INVALID_ENUM
    End Select
End Sub

Public Sub alDistanceModel(ByVal value As Long)
    If Not contextCurrent Then Exit Sub
    Select Case value
    Case AL_NONE, AL_INVERSE_DISTANCE, AL_INVERSE_DISTANCE_CLAMPED
        distModel = value
    Case Else
        setError AL_INVALID_VALUE
    End Select
End Sub

'--------------------------------------------------------------------------
' the device and the context (alc.h)

Public Function alcGetError(ByVal device As Long) As Long
    alcGetError = alcError
    alcError = ALC_NO_ERROR
End Function

' the waveOut device's name (the wave mapper's)
Private Function mapperName() As String
    Dim caps As WAVEOUTCAPSW, s As String, n As Long
    If waveOutGetDevCapsW(WAVE_MAPPER, caps, LenB(caps)) <> MMSYSERR_NOERROR Then
        mapperName = "waveOut"
        Exit Function
    End If
    s = String$(32, vbNullChar)
    CopyMemory ByVal StrPtr(s), caps.szPname(0), 64
    n = InStr(s, vbNullChar)
    If n > 0 Then s = Left$(s, n - 1)
    mapperName = "waveOut: " & s
End Function

Private Function defaultDeviceName() As String
    If Len(capturePath) > 0 Then
        defaultDeviceName = "WarzoneVB capture: " & capturePath
    Else
        defaultDeviceName = mapperName()
    End If
End Function

Public Function alcGetString(ByVal device As Long, ByVal param As Long) As String
    Select Case param
    Case ALC_NO_ERROR: alcGetString = "No Error"
    Case ALC_INVALID_DEVICE: alcGetString = "Invalid Device"
    Case ALC_INVALID_CONTEXT: alcGetString = "Invalid Context"
    Case ALC_INVALID_ENUM: alcGetString = "Invalid Enum"
    Case ALC_INVALID_VALUE: alcGetString = "Invalid Value"
    Case ALC_OUT_OF_MEMORY: alcGetString = "Out of Memory"
    Case ALC_DEFAULT_DEVICE_SPECIFIER
        alcGetString = defaultDeviceName()
    Case ALC_DEVICE_SPECIFIER
        If device = 0 Then
            ' the list of the devices: each name followed by a NUL, the list by another
            alcGetString = defaultDeviceName() & vbNullChar & vbNullChar
        ElseIf device = AL_DEVICE And deviceOpen Then
            alcGetString = openedName
        Else
            setContextError ALC_INVALID_DEVICE
            alcGetString = vbNullString
        End If
    Case ALC_EXTENSIONS
        If device = AL_DEVICE And deviceOpen Then
            alcGetString = ""
        Else
            setContextError ALC_INVALID_DEVICE
            alcGetString = vbNullString
        End If
    Case Else
        setContextError ALC_INVALID_ENUM
        alcGetString = vbNullString
    End Select
End Function

Public Sub alcGetIntegerv(ByVal device As Long, ByVal param As Long, ByVal size As Long, data As Long)
    If size < 1 Then setContextError ALC_INVALID_VALUE: Exit Sub
    Select Case param
    Case ALC_MAJOR_VERSION: data = 1
    Case ALC_MINOR_VERSION: data = 1
    Case ALC_FREQUENCY, ALC_REFRESH
        If device <> AL_DEVICE Or Not deviceOpen Then setContextError ALC_INVALID_DEVICE: Exit Sub
        If param = ALC_FREQUENCY Then data = MIX_RATE Else data = MIX_RATE \ BLOCK_FRAMES
    Case Else
        setContextError ALC_INVALID_ENUM
    End Select
End Sub

Public Function alcIsExtensionPresent(ByVal device As Long, ByVal extName As String) As Long
    alcIsExtensionPresent = ALC_FALSE   ' no extensions
End Function

Public Function alcGetProcAddress(ByVal device As Long, ByVal funcName As String) As Long
    alcGetProcAddress = 0               ' no extension functions
End Function

Public Function alcOpenDevice(ByVal devicename As String) As Long
    Dim fmt As WAVEFORMATEX, i As Long
    If deviceOpen Then setContextError ALC_INVALID_VALUE: Exit Function   ' (one device)
    If Len(capturePath) > 0 Then
        captureFile = CreateFileW(StrPtr(capturePath), GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        If captureFile = INVALID_HANDLE_VALUE Then
            captureFile = 0
            Exit Function
        End If
        captureFrames = 0#
        writeWavHeader 0
        QueryPerformanceCounter captureStart
    Else
        fmt.wFormatTag = WAVE_FORMAT_PCM
        fmt.nChannels = 2
        fmt.wBitsPerSample = 16
        fmt.nSamplesPerSec = MIX_RATE
        fmt.nBlockAlign = 4
        fmt.nAvgBytesPerSec = MIX_RATE * 4
        fmt.cbSize = 0
        If waveOutOpen(hWaveOut, WAVE_MAPPER, fmt, 0, 0, CALLBACK_NULL) <> MMSYSERR_NOERROR Then
            hWaveOut = 0
            Exit Function
        End If
        ReDim ringPcm(0 To BLOCKS * BLOCK_FRAMES * 2 - 1)    ' (never resized while the device is open)
        For i = 0 To BLOCKS - 1
            ringHdr(i).lpData = VarPtr(ringPcm(i * BLOCK_FRAMES * 2))
            ringHdr(i).dwBufferLength = BLOCK_FRAMES * 4
            ringHdr(i).dwBytesRecorded = 0
            ringHdr(i).dwUser = 0
            ringHdr(i).dwFlags = 0
            ringHdr(i).dwLoops = 0
            If waveOutPrepareHeader(hWaveOut, VarPtr(ringHdr(i)), LenB(ringHdr(i))) <> MMSYSERR_NOERROR Then
                closeWave
                Exit Function
            End If
        Next
        ' start the device on silence: each block played is mixed and written again by alcProcessContext
        For i = 0 To BLOCKS - 1
            waveOutWrite hWaveOut, VarPtr(ringHdr(i)), LenB(ringHdr(i))
        Next
        nextBlock = 0
    End If
    openedName = defaultDeviceName()
    deviceOpen = True
    alcError = ALC_NO_ERROR
    alcOpenDevice = AL_DEVICE
End Function

Private Sub closeWave()
    Dim i As Long
    If hWaveOut = 0 Then Exit Sub
    waveOutReset hWaveOut
    For i = 0 To BLOCKS - 1
        waveOutUnprepareHeader hWaveOut, VarPtr(ringHdr(i)), LenB(ringHdr(i))
    Next
    waveOutClose hWaveOut
    hWaveOut = 0
    Erase ringPcm
End Sub

Public Function alcCloseDevice(ByVal device As Long) As Long
    Dim b As Long
    If device <> AL_DEVICE Or Not deviceOpen Then setContextError ALC_INVALID_DEVICE: Exit Function
    If contextAlive Then destroyContext
    If captureFile <> 0 Then
        CloseHandle captureFile
        captureFile = 0
    End If
    closeWave
    ' the device's buffers
    For b = 0 To bufCap - 1
        bufs(b).used = False
        Erase bufs(b).data
    Next
    deviceOpen = False
    alcCloseDevice = ALC_TRUE
End Function

Public Function alcCreateContext(ByVal device As Long, ByVal attrlist As Long) As Long
    Dim s As Long
    If device <> AL_DEVICE Or Not deviceOpen Then setContextError ALC_INVALID_DEVICE: Exit Function
    If contextAlive Then setContextError ALC_INVALID_VALUE: Exit Function    ' (one context)
    For s = 0 To MAX_SOURCES - 1
        srcs(s).used = False
        Erase srcs(s).queue
        srcs(s).queueCount = 0
    Next
    ' the listener's defaults
    listenerPos(0) = 0!: listenerPos(1) = 0!: listenerPos(2) = 0!
    listenerVel(0) = 0!: listenerVel(1) = 0!: listenerVel(2) = 0!
    listenerAt(0) = 0!: listenerAt(1) = 0!: listenerAt(2) = -1!
    listenerUp(0) = 0!: listenerUp(1) = 1!: listenerUp(2) = 0!
    listenerGain = 1!
    distModel = AL_INVERSE_DISTANCE_CLAMPED
    alError = AL_NO_ERROR
    contextAlive = True
    alcCreateContext = AL_CONTEXT
End Function

Public Function alcMakeContextCurrent(ByVal context As Long) As Long
    If context = 0 Then
        contextCurrent = False
        alcMakeContextCurrent = ALC_TRUE
    ElseIf context = AL_CONTEXT And contextAlive Then
        contextCurrent = True
        alcMakeContextCurrent = ALC_TRUE
    Else
        setContextError ALC_INVALID_CONTEXT
        alcMakeContextCurrent = ALC_FALSE
    End If
End Function

Private Sub destroyContext()
    Dim s As Long
    For s = 0 To MAX_SOURCES - 1
        If srcs(s).used Then
            clearQueue s
            srcs(s).used = False
        End If
    Next
    contextAlive = False
    contextCurrent = False
End Sub

Public Sub alcDestroyContext(ByVal context As Long)
    If context <> AL_CONTEXT Or Not contextAlive Then setContextError ALC_INVALID_CONTEXT: Exit Sub
    destroyContext
End Sub

' the device's work: the blocks played are mixed and written again (the capture: the frames the
' real time asks for)
Public Sub alcProcessContext(ByVal context As Long)
    Dim now As Currency, freq As Currency, due As Double
    If context <> AL_CONTEXT Or Not contextAlive Then setContextError ALC_INVALID_CONTEXT: Exit Sub
    If captureFile <> 0 Then
        QueryPerformanceCounter now
        QueryPerformanceFrequency freq
        due = Int(CDbl(now - captureStart) / CDbl(freq) * MIX_RATE) - captureFrames
        Do While due > 0#
            If due > 4096# Then
                al_vbCaptureMix 4096
                due = due - 4096#
            Else
                al_vbCaptureMix CLng(due)
                due = 0#
            End If
        Loop
    ElseIf hWaveOut <> 0 Then
        Do While (ringHdr(nextBlock).dwFlags And WHDR_DONE) <> 0
            mixFrames ringPcm, nextBlock * BLOCK_FRAMES * 2, BLOCK_FRAMES
            ringHdr(nextBlock).dwFlags = ringHdr(nextBlock).dwFlags And Not WHDR_DONE
            If waveOutWrite(hWaveOut, VarPtr(ringHdr(nextBlock)), LenB(ringHdr(nextBlock))) <> MMSYSERR_NOERROR Then Exit Do
            nextBlock = (nextBlock + 1) Mod BLOCKS
        Loop
    End If
End Sub

'--------------------------------------------------------------------------
' the capture (test aids)

' (VB) before alcOpenDevice: no device, the mix goes to this WAV file ("" opens the device again)
Public Sub al_vbSetCapture(ByVal wavPath As String)
    capturePath = wavPath
End Sub

Private Sub writeLE(ByVal v As Long)
    Dim written As Long
    WriteFile captureFile, v, 4, written, 0
End Sub

' the WAV header (44100 Hz, 16 bit stereo) with dataBytes of samples, at the file's start
Private Sub writeWavHeader(ByVal dataBytes As Long)
    SetFilePointer captureFile, 0, 0, FILE_BEGIN
    writeLE &H46464952                  ' "RIFF"
    writeLE 36 + dataBytes
    writeLE &H45564157                  ' "WAVE"
    writeLE &H20746D66                  ' "fmt "
    writeLE 16
    writeLE WAVE_FORMAT_PCM + 2 * &H10000
    writeLE MIX_RATE
    writeLE MIX_RATE * 4
    writeLE 4 + 16 * &H10000
    writeLE &H61746164                  ' "data"
    writeLE dataBytes
    SetFilePointer captureFile, 0, 0, FILE_END
End Sub

' (VB) mixes frames frames now and appends them to the capture file
Public Sub al_vbCaptureMix(ByVal frames As Long)
    Dim written As Long
    If captureFile = 0 Or frames <= 0 Then Exit Sub
    If Not contextAlive Then Exit Sub
    ReDim captureBuf(0 To frames * 2 - 1)
    mixFrames captureBuf, 0, frames
    WriteFile captureFile, captureBuf(0), frames * 4, written, 0
    captureFrames = captureFrames + frames
    writeWavHeader CLng(captureFrames * 4#)
End Sub

' (VB) mixes frames frames into nothing: the mixer's cost
Public Sub al_vbMixBench(ByVal frames As Long)
    If Not contextAlive Or frames <= 0 Then Exit Sub
    ReDim captureBuf(0 To frames * 2 - 1)
    mixFrames captureBuf, 0, frames
End Sub

'--------------------------------------------------------------------------
' the mixer

' mixes frames stereo frames of the playing sources into out(first ...), interleaved, 16 bit
Private Sub mixFrames(out() As Integer, ByVal first As Long, ByVal frames As Long)
    Dim i As Long, s As Long, v As Long, j As Long
    If accSize < frames Then
        accSize = frames
        ReDim accL(0 To accSize - 1)
        ReDim accR(0 To accSize - 1)
    Else
        For i = 0 To frames - 1
            accL(i) = 0!
            accR(i) = 0!
        Next
    End If
    For s = 0 To MAX_SOURCES - 1
        If srcs(s).used Then
            If srcs(s).state = AL_PLAYING Then mixSource s, frames
        End If
    Next
    j = first
    For i = 0 To frames - 1
        If accL(i) >= 32767! Then
            v = 32767
        ElseIf accL(i) <= -32768! Then
            v = -32768
        Else
            v = CLng(accL(i))
        End If
        out(j) = v
        If accR(i) >= 32767! Then
            v = 32767
        ElseIf accR(i) <= -32768! Then
            v = -32768
        Else
            v = CLng(accR(i))
        End If
        out(j + 1) = v
        j = j + 2
    Next
End Sub

' the source's effective gains on the left and right channels
Private Sub sourceGains(ByVal s As Long, gl As Single, gr As Single)
    Dim g As Single, d(0 To 2) As Single, dist As Single, b As Long
    Dim rx As Single, ry As Single, rz As Single, rl As Single, pan As Single, k As Long
    g = srcs(s).gain * listenerGain
    ' the direction from the listener (a relative source: already in the listener's frame)
    For k = 0 To 2
        If srcs(s).relative Then d(k) = srcs(s).pos(k) Else d(k) = srcs(s).pos(k) - listenerPos(k)
    Next
    dist = Sqr(d(0) * d(0) + d(1) * d(1) + d(2) * d(2))
    If distModel = AL_INVERSE_DISTANCE_CLAMPED Or distModel = AL_INVERSE_DISTANCE Then
        ' reference distance 1, rolloff 1: 1 / (1 + (dist - 1))
        If distModel = AL_INVERSE_DISTANCE_CLAMPED Then
            If dist < 1! Then g = g / 1! Else g = g / dist
        ElseIf dist > 0! Then
            g = g / dist
        End If
    End If
    If g > 1! Then g = 1!
    If g < 0! Then g = 0!
    b = -1
    If srcs(s).current < srcs(s).queueCount Then b = bufSlot(srcs(s).queue(srcs(s).current))
    If b >= 0 Then
        If bufs(b).channels = 2 Then        ' stereo: not panned
            gl = g
            gr = g
            Exit Sub
        End If
    End If
    pan = 0!
    If dist > 0.000001! Then
        If srcs(s).relative Then
            pan = d(0) / dist               ' the listener's frame: right is +x
        Else
            ' right = at x up
            rx = listenerAt(1) * listenerUp(2) - listenerAt(2) * listenerUp(1)
            ry = listenerAt(2) * listenerUp(0) - listenerAt(0) * listenerUp(2)
            rz = listenerAt(0) * listenerUp(1) - listenerAt(1) * listenerUp(0)
            rl = Sqr(rx * rx + ry * ry + rz * rz)
            If rl > 0.000001! Then pan = (d(0) * rx + d(1) * ry + d(2) * rz) / (rl * dist)
        End If
    End If
    If pan > 1! Then pan = 1!
    If pan < -1! Then pan = -1!
    gl = g * Sqr((1! - pan) / 2!)
    gr = g * Sqr((1! + pan) / 2!)
End Sub

' the first frame's sample (channel ch) of what follows the current buffer: the next queued one, the
' first again when looping, else silence
Private Function nextFirstSample(ByVal s As Long, ByVal ch As Long) As Single
    Dim nq As Long, b As Long
    nq = srcs(s).current + 1
    If nq >= srcs(s).queueCount Then
        If Not srcs(s).looping Then Exit Function
        nq = 0
    End If
    b = bufSlot(srcs(s).queue(nq))
    If b < 0 Then Exit Function
    If bufs(b).frames = 0 Then Exit Function
    If ch < bufs(b).channels Then nextFirstSample = bufs(b).data(ch)
End Function

' adds the source's next frames frames to the accumulators, advancing it (and stopping it at the end)
Private Sub mixSource(ByVal s As Long, ByVal frames As Long)
    Dim gl As Single, gr As Single, b As Long, nframes As Long, stp As Double, pos As Double
    Dim i As Long, ip As Long, fr As Single, s0 As Single, s1 As Single, v As Single, emptyRun As Long
    Dim t0 As Single, t1 As Single
    sourceGains s, gl, gr
    i = 0
    Do While i < frames
        If srcs(s).current >= srcs(s).queueCount Then
            srcs(s).state = AL_STOPPED
            srcs(s).current = srcs(s).queueCount
            srcs(s).offset = 0#
            Exit Do
        End If
        b = bufSlot(srcs(s).queue(srcs(s).current))
        If b < 0 Then nframes = 0 Else nframes = bufs(b).frames
        If srcs(s).offset >= nframes Then
            ' this buffer is done: the next one
            srcs(s).offset = srcs(s).offset - nframes
            If nframes = 0 Then
                srcs(s).offset = 0#
                emptyRun = emptyRun + 1
                If emptyRun > srcs(s).queueCount Then    ' nothing but empty buffers
                    srcs(s).state = AL_STOPPED
                    srcs(s).current = srcs(s).queueCount
                    Exit Do
                End If
            Else
                emptyRun = 0
            End If
            srcs(s).current = srcs(s).current + 1
            If srcs(s).current >= srcs(s).queueCount And srcs(s).looping Then srcs(s).current = 0
        Else
            stp = CDbl(bufs(b).freq) * srcs(s).pitch / MIX_RATE
            pos = srcs(s).offset
            If bufs(b).channels = 1 Then
                Do While i < frames
                    If pos >= nframes Then Exit Do
                    ip = Int(pos)
                    fr = pos - ip
                    s0 = bufs(b).data(ip)
                    If ip + 1 < nframes Then s1 = bufs(b).data(ip + 1) Else s1 = nextFirstSample(s, 0)
                    v = s0 + (s1 - s0) * fr
                    accL(i) = accL(i) + v * gl
                    accR(i) = accR(i) + v * gr
                    pos = pos + stp
                    i = i + 1
                Loop
            Else
                Do While i < frames
                    If pos >= nframes Then Exit Do
                    ip = Int(pos)
                    fr = pos - ip
                    s0 = bufs(b).data(2 * ip)
                    t0 = bufs(b).data(2 * ip + 1)
                    If ip + 1 < nframes Then
                        s1 = bufs(b).data(2 * ip + 2)
                        t1 = bufs(b).data(2 * ip + 3)
                    Else
                        s1 = nextFirstSample(s, 0)
                        t1 = nextFirstSample(s, 1)
                    End If
                    accL(i) = accL(i) + (s0 + (s1 - s0) * fr) * gl
                    accR(i) = accR(i) + (t0 + (t1 - t0) * fr) * gr
                    pos = pos + stp
                    i = i + 1
                Loop
            End If
            srcs(s).offset = pos
        End If
    Loop
End Sub
