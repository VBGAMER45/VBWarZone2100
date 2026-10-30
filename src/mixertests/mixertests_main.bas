Attribute VB_Name = "mixertests_main"
' MixerTests.exe (cluster H, tools/mixer_test.py): the OpenAL subset of src/sound/al_win.bas, without
' WzSound: synthetic buffers (sines) through sources in capture mode, one WAV file per scenario, and
' results.txt with the states, counts and error codes the scenarios read back ("name value" lines).
' tools/mixer_test.py measures the WAV files (levels, balance, frequencies, clipping) and checks the
' results. usage: MixerTests.exe <output folder>
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const PI As Double = 3.14159265358979

Private outDir As String
Private resultsFile As Integer

Private Sub result(ByVal name_ As String, ByVal value As Variant)
    Print #resultsFile, name_ & " " & CStr(value)
End Sub

' a 16-bit sine of frames frames at rate Hz (stereo: on the left channel only when leftOnly)
Private Function sine(ByVal hz As Double, ByVal rate As Long, ByVal frames As Long, ByVal amp As Double, _
                      ByVal channels As Long, Optional ByVal leftOnly As Boolean = False) As Byte()
    Dim b() As Byte, i As Long, c As Long, v As Integer, k As Long
    ReDim b(0 To frames * channels * 2 - 1)
    k = 0
    For i = 0 To frames - 1
        For c = 0 To channels - 1
            If c = 1 And leftOnly Then
                v = 0
            Else
                v = CInt(amp * Sin(2# * PI * hz * i / rate))
            End If
            b(k) = v And &HFF
            b(k + 1) = ((v And &HFF00&) \ &H100&) And &HFF
            k = k + 2
        Next
    Next
    sine = b
End Function

Private Function makeBuffer(data() As Byte, ByVal channels As Long, ByVal rate As Long) As Long
    Dim ids(0 To 0) As Long, fmt_ As Long
    alGenBuffers 1, ids
    If channels = 1 Then fmt_ = AL_FORMAT_MONO16 Else fmt_ = AL_FORMAT_STEREO16
    alBufferData ids(0), fmt_, data, UBound(data) + 1, rate
    makeBuffer = ids(0)
End Function

Private Function makeSource() As Long
    Dim ids(0 To 0) As Long
    alGenSources 1, ids
    makeSource = ids(0)
End Function

Private Function state(ByVal src As Long) As Long
    alGetSourcei src, AL_SOURCE_STATE, state
End Function

Private Function processed(ByVal src As Long) As Long
    alGetSourcei src, AL_BUFFERS_PROCESSED, processed
End Function

' opens the capture device on <out>\<name>.wav with a current context
Private Sub openScenario(ByVal name_ As String)
    Dim dev As Long, ctx As Long
    al_vbSetCapture outDir & "\" & name_ & ".wav"
    dev = alcOpenDevice(vbNullString)
    ctx = alcCreateContext(dev, 0)
    alcMakeContextCurrent ctx
    alGetError
End Sub

Private Sub closeScenario()
    alcMakeContextCurrent 0
    alcDestroyContext 1
    alcCloseDevice 1
    al_vbSetCapture ""
End Sub

' a mono 1 kHz sine source at 44100 Hz (amplitude 16000, 2 s)
Private Function toneSource(ByVal hz As Double) As Long
    Dim d() As Byte, buf As Long, src As Long
    d = sine(hz, 44100, 88200, 16000#, 1)
    buf = makeBuffer(d, 1, 44100)
    src = makeSource()
    alSourcei src, AL_BUFFER, buf
    toneSource = src
End Function

Private Sub scenarioGain()
    Dim src As Long
    openScenario "gain"
    src = toneSource(1000#)
    alSourcei src, AL_SOURCE_RELATIVE, AL_TRUE
    alSourcef src, AL_GAIN, 0.5!
    alSourcePlay src
    al_vbCaptureMix 22050
    result "gain_error", alGetError()
    closeScenario
End Sub

' a source at x, y, z (not relative) with the listener's orientation at / up, distance model AL_NONE
' unless distance
Private Sub scenarioPan(ByVal name_ As String, ByVal x As Single, ByVal y As Single, ByVal z As Single, _
                        ByVal ax As Single, ByVal ay As Single, ByVal az As Single, _
                        ByVal ux As Single, ByVal uy As Single, ByVal uz As Single, ByVal distance As Boolean)
    Dim src As Long, ori(0 To 5) As Single
    openScenario name_
    ori(0) = ax: ori(1) = ay: ori(2) = az: ori(3) = ux: ori(4) = uy: ori(5) = uz
    alListenerfv AL_ORIENTATION, ori
    alListener3f AL_POSITION, 100!, 200!, 300!
    If Not distance Then alDistanceModel AL_NONE
    src = toneSource(1000#)
    alSource3f src, AL_POSITION, 100! + x, 200! + y, 300! + z
    alSourcePlay src
    al_vbCaptureMix 22050
    result name_ & "_error", alGetError()
    closeScenario
End Sub

Private Sub scenarioResample()
    Dim d() As Byte, buf As Long, src As Long
    openScenario "resample"
    d = sine(441#, 11025, 11025, 16000#, 1)
    buf = makeBuffer(d, 1, 11025)
    src = makeSource()
    alSourcei src, AL_BUFFER, buf
    alSourcePlay src
    al_vbCaptureMix 44100
    result "resample_state_1s", state(src)
    al_vbCaptureMix 4410
    result "resample_state_end", state(src)
    closeScenario
End Sub

Private Sub scenarioPitch()
    Dim src As Long
    openScenario "pitch"
    src = toneSource(1000#)
    alSourcef src, AL_PITCH, 2!
    alSourcePlay src
    al_vbCaptureMix 22050
    result "pitch_state_1s", state(src)
    al_vbCaptureMix 22100                   ' (the 2 s buffer ends at 44100 frames: its end is found after)
    result "pitch_state_end", state(src)
    closeScenario
End Sub

Private Sub scenarioLoop()
    Dim d() As Byte, buf As Long, src As Long
    openScenario "loop"
    d = sine(1000#, 44100, 1000, 16000#, 1)
    buf = makeBuffer(d, 1, 44100)
    src = makeSource()
    alSourcei src, AL_BUFFER, buf
    alSourcei src, AL_LOOPING, AL_TRUE
    alSourcePlay src
    al_vbCaptureMix 44100
    result "loop_state", state(src)
    result "loop_processed", processed(src)
    closeScenario
End Sub

Private Sub scenarioQueue()
    Dim d() As Byte, bufs(0 To 2) As Long, src As Long, i As Long, out(0 To 2) As Long, extra As Long
    openScenario "queue"
    For i = 0 To 2
        d = sine(500# * (i + 1), 44100, 4410, 16000#, 1)
        bufs(i) = makeBuffer(d, 1, 44100)
    Next
    src = makeSource()
    alSourceQueueBuffers src, 3, bufs
    alGetSourcei src, AL_BUFFERS_QUEUED, out(0)
    result "queue_queued", out(0)
    alSourcePlay src
    al_vbCaptureMix 2205
    result "queue_processed_0.05s", processed(src)
    al_vbCaptureMix 4410
    result "queue_processed_0.15s", processed(src)
    alSourceUnqueueBuffers src, 1, out
    result "queue_unqueued_first", IIf(out(0) = bufs(0), 1, 0)
    alGetSourcei src, AL_BUFFERS_QUEUED, out(1)
    result "queue_queued_after_unqueue", out(1)
    alSourceUnqueueBuffers src, 2, out
    result "queue_unqueue_too_many_error", alGetError()
    al_vbCaptureMix 13230
    result "queue_state_end", state(src)
    result "queue_processed_end", processed(src)
    ' refilled after the end and played again: plays the new buffer (the WAV's last 0.1 s: 500 Hz)
    alSourceUnqueueBuffers src, 2, out
    d = sine(500#, 44100, 4410, 16000#, 1)
    alBufferData out(0), AL_FORMAT_MONO16, d, UBound(d) + 1, 44100
    alSourceQueueBuffers src, 1, out
    alSourcePlay src
    result "queue_state_replay", state(src)
    al_vbCaptureMix 4410
    result "queue_error", alGetError()
    closeScenario
End Sub

Private Sub scenarioStates()
    Dim src As Long
    openScenario "states"
    src = toneSource(1000#)
    alSourcePlay src
    al_vbCaptureMix 4410                    ' 0.0 - 0.1 s: playing
    alSourcePause src
    result "states_paused", state(src)
    al_vbCaptureMix 4410                    ' 0.1 - 0.2 s: silent
    alSourcePlay src
    result "states_resumed", state(src)
    al_vbCaptureMix 4410                    ' 0.2 - 0.3 s: playing
    alSourceStop src
    result "states_stopped", state(src)
    result "states_processed_stopped", processed(src)
    al_vbCaptureMix 4410                    ' 0.3 - 0.4 s: silent
    alSourcePlay src
    result "states_replayed", state(src)
    al_vbCaptureMix 4410                    ' 0.4 - 0.5 s: playing
    result "states_error", alGetError()
    closeScenario
End Sub

Private Sub scenarioClamp()
    Dim a As Long, b As Long, d() As Byte, buf As Long
    openScenario "clamp"
    d = sine(1000#, 44100, 44100, 30000#, 1)
    buf = makeBuffer(d, 1, 44100)
    a = makeSource()
    b = makeSource()
    alSourcei a, AL_BUFFER, buf
    alSourcei b, AL_BUFFER, buf
    alSourcei a, AL_SOURCE_RELATIVE, AL_TRUE
    alSourcei b, AL_SOURCE_RELATIVE, AL_TRUE
    alSourcePlay a
    alSourcePlay b
    al_vbCaptureMix 4410
    result "clamp_error", alGetError()
    closeScenario
End Sub

Private Sub scenarioStereo()
    Dim d() As Byte, buf As Long, src As Long
    openScenario "stereo"
    alDistanceModel AL_NONE
    d = sine(1000#, 44100, 44100, 16000#, 2, True)
    buf = makeBuffer(d, 2, 44100)
    src = makeSource()
    alSourcei src, AL_BUFFER, buf
    alSource3f src, AL_POSITION, 10!, 0!, 0!
    alSourcePlay src
    al_vbCaptureMix 22050
    result "stereo_error", alGetError()
    closeScenario
End Sub

Private Sub scenarioErrors()
    Dim src As Long, d() As Byte, buf As Long, ids(0 To 0) As Long
    openScenario "errors"
    alSourcef 999, AL_GAIN, 1!
    result "error_invalid_name", alGetError()
    src = toneSource(1000#)
    alSourcef src, AL_GAIN, -1!
    result "error_negative_gain", alGetError()
    alGetSourcei src, AL_BUFFER, buf
    d = sine(1000#, 44100, 100, 1000#, 1)
    alBufferData buf, AL_FORMAT_MONO16, d, UBound(d) + 1, 44100
    result "error_buffer_in_use", alGetError()
    ids(0) = buf
    alDeleteBuffers 1, ids
    result "error_delete_in_use", alGetError()
    alSourcef src, 12345, 1!
    result "error_bad_enum", alGetError()
    result "error_none_left", alGetError()
    closeScenario
End Sub

' the waveOut device itself, silent (a source at gain 0): alcProcessContext cycles its blocks in real time,
' so a 0.1 s buffer has ended after 0.4 s (the ring is 93 ms); no device: device_open 0
Private Sub scenarioDevice()
    Dim dev As Long, ctx As Long, src As Long, t0 As Currency, t As Currency, f As Currency, d() As Byte, buf As Long
    al_vbSetCapture ""
    dev = alcOpenDevice(vbNullString)
    result "device_open", IIf(dev <> 0, 1, 0)
    If dev = 0 Then Exit Sub
    ctx = alcCreateContext(dev, 0)
    alcMakeContextCurrent ctx
    d = sine(1000#, 44100, 4410, 16000#, 1)
    buf = makeBuffer(d, 1, 44100)
    src = makeSource()
    alSourcei src, AL_BUFFER, buf
    alSourcef src, AL_GAIN, 0!
    alSourcePlay src
    result "device_state_start", state(src)
    QueryPerformanceFrequency f
    QueryPerformanceCounter t0
    Do
        alcProcessContext ctx
        Sleep 5
        QueryPerformanceCounter t
    Loop While CDbl(t - t0) / CDbl(f) < 0.4
    result "device_state_0.4s", state(src)
    result "device_error", alGetError()
    closeScenario
End Sub

' the cost of mixing n looping 11025 Hz mono sources at different places: ms per second of audio
Private Sub benchmark(ByVal n As Long)
    Dim d() As Byte, buf As Long, i As Long, src As Long, t0 As Currency, t1 As Currency, f As Currency
    openScenario "bench"
    alDistanceModel AL_NONE
    d = sine(440#, 11025, 11025, 8000#, 1)
    buf = makeBuffer(d, 1, 11025)
    For i = 1 To n
        src = makeSource()
        alSourcei src, AL_BUFFER, buf
        alSourcei src, AL_LOOPING, AL_TRUE
        alSource3f src, AL_POSITION, CSng(i * 37 Mod 200) - 100!, 0!, CSng(i * 53 Mod 200) - 100!
        alSourcePlay src
    Next
    QueryPerformanceCounter t0
    For i = 1 To 861                    ' 10 s of 512-frame blocks
        al_vbMixBench 512
    Next
    QueryPerformanceCounter t1
    QueryPerformanceFrequency f
    result "bench_" & n & "_ms_per_s", Format$(CDbl(t1 - t0) / CDbl(f) * 1000# / 10#, "0.00")
    closeScenario
End Sub

Sub Main()
    outDir = Trim$(Command$)
    If Left$(outDir, 1) = """" Then outDir = Mid$(outDir, 2, Len(outDir) - 2)
    If Len(outDir) = 0 Then outDir = CurDir$
    resultsFile = FreeFile
    Open outDir & "\results.txt" For Output As #resultsFile
    scenarioGain
    ' default orientation (at 0 0 -1, up 0 1 0): right is +x
    scenarioPan "pan_right", 10!, 0!, 0!, 0!, 0!, -1!, 0!, 1!, 0!, False
    scenarioPan "pan_left", -10!, 0!, 0!, 0!, 0!, -1!, 0!, 1!, 0!, False
    scenarioPan "pan_45", 10!, 0!, -10!, 0!, 0!, -1!, 0!, 1!, 0!, False
    ' WZ's orientation for a camera turned by pi / 2 (sound_SetPlayerOrientation: at -sin, cos, 0; up 0 0 1):
    ' right is +y
    scenarioPan "pan_wz", 0!, 10!, 0!, -1!, 0!, 0!, 0!, 0!, 1!, False
    ' the default distance model (inverse, clamped; reference 1, rolloff 1): 4 in front, gain 1 / 4
    scenarioPan "distance", 0!, 0!, -4!, 0!, 0!, -1!, 0!, 1!, 0!, True
    scenarioResample
    scenarioPitch
    scenarioLoop
    scenarioQueue
    scenarioStates
    scenarioClamp
    scenarioStereo
    scenarioErrors
    scenarioDevice
    benchmark 1
    benchmark 8
    benchmark 32
    result "done", 1
    Close #resultsFile
End Sub
