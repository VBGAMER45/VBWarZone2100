Attribute VB_Name = "soundgame_vb"
' (VB) lib/sound's calls into the game: the ISoundGame the exe hands to SoundSetGame (SoundApi), and
' procedures with the C++ names (aud.h's, tracklib.h's sound_GetGameTime, gtime's realTime, droid.h's
' droidAudioTrackStopped) that call it. An AUDIO_CALLBACK is the game's number; audio_CallBack calls it as
' C++'s psSample->pCallback(psSample->psObj).
Option Explicit

Public soundGame As ISoundGame

' aud.h
Public Sub audio_GetObjectPos(ByVal psObj As Long, piX As Long, piY As Long, piZ As Long)
    soundGame.audioGetObjectPos psObj, piX, piY, piZ
End Sub

Public Sub audio_GetStaticPos(ByVal iWorldX As Long, ByVal iWorldY As Long, piX As Long, piY As Long, piZ As Long)
    soundGame.audioGetStaticPos iWorldX, iWorldY, piX, piY, piZ
End Sub

Public Function audio_ObjectDead(ByVal psObj As Long) As Boolean
    audio_ObjectDead = soundGame.audioObjectDead(psObj)
End Function

' Vector3f audio_GetPlayerPos(): the vector's fields
Public Sub audio_GetPlayerPos(x As Single, y As Single, z As Single)
    soundGame.audioGetPlayerPos x, y, z
End Sub

Public Sub audio_Get3DPlayerRotAboutVerticalAxis(angle As Single)
    soundGame.audioGet3DPlayerRotAboutVerticalAxis angle
End Sub

' tracklib.h
Public Function sound_GetGameTime() As Long
    sound_GetGameTime = soundGame.soundGetGameTime()
End Function

' lib/gamelib/gtime.h: UDWORD realTime
Public Function realTime() As Long
    realTime = soundGame.realTime()
End Function

' an AUDIO_CALLBACK pCallback(psObj) (pCallback <> 0)
Public Function audio_CallBack(ByVal pCallback As Long, ByVal psObj As Long) As Boolean
    audio_CallBack = soundGame.audioCallback(pCallback, psObj)
End Function

' droid.h
Public Function droidAudioTrackStopped(ByVal psObj As Long) As Boolean
    droidAudioTrackStopped = soundGame.droidAudioTrackStopped(psObj)
End Function
