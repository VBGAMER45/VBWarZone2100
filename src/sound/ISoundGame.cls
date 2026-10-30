VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISoundGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' What lib/sound calls in the game (src, lib/gamelib): the exe implements this interface (SoundGame.cls)
' and hands it to SoundSetGame at startup. aud.h's functions (src/aud.cpp), tracklib.h's sound_GetGameTime
' (aud.cpp), gtime's realTime, droid.cpp's droidAudioTrackStopped (track.cpp calls it directly) and the
' AUDIO_CALLBACKs, which are the game's numbers. A SIMPLE_OBJECT * is the game's Long handle (0 =
' nullptr). VB refuses underscores in the method names of an interface used with Implements, so each name
' drops its underscores and capitalises the letter after each one.
Option Explicit

' aud.h: void audio_GetObjectPos(const SIMPLE_OBJECT *psObj, SDWORD *piX, SDWORD *piY, SDWORD *piZ)
Public Sub audioGetObjectPos(ByVal psObj As Long, piX As Long, piY As Long, piZ As Long)
End Sub

' aud.h: void audio_GetStaticPos(SDWORD iWorldX, SDWORD iWorldY, SDWORD *piX, SDWORD *piY, SDWORD *piZ)
Public Sub audioGetStaticPos(ByVal iWorldX As Long, ByVal iWorldY As Long, piX As Long, piY As Long, piZ As Long)
End Sub

' aud.h: bool audio_ObjectDead(const SIMPLE_OBJECT *psObj)
Public Function audioObjectDead(ByVal psObj As Long) As Boolean
End Function

' aud.h: Vector3f audio_GetPlayerPos() (the vector's fields)
Public Sub audioGetPlayerPos(x As Single, y As Single, z As Single)
End Sub

' aud.h: void audio_Get3DPlayerRotAboutVerticalAxis(float *angle)
Public Sub audioGet3DPlayerRotAboutVerticalAxis(angle As Single)
End Sub

' tracklib.h: UDWORD sound_GetGameTime() (aud.cpp: gameTime)
Public Function soundGetGameTime() As Long
End Function

' lib/gamelib/gtime.h: UDWORD realTime
Public Function realTime() As Long
End Function

' an AUDIO_CALLBACK, bool (*)(void *psObj): the game's number for the function, called with the object
Public Function audioCallback(ByVal fn As Long, ByVal psObj As Long) As Boolean
End Function

' droid.h: bool droidAudioTrackStopped(void *psObj) (track.cpp's sound_FinishedCallback calls it)
Public Function droidAudioTrackStopped(ByVal psObj As Long) As Boolean
End Function
