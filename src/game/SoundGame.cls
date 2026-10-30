VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The game side of WzSound.dll's ISoundGame: aud.cpp's functions (aud_c), the game clock, and the
' AUDIO_CALLBACKs by their numbers (CB_*: droid_c, move_c).
Option Explicit

Implements ISoundGame

Private Sub ISoundGame_audioGetObjectPos(ByVal psObj As Long, piX As Long, piY As Long, piZ As Long)
    audio_GetObjectPos psObj, piX, piY, piZ
End Sub

Private Sub ISoundGame_audioGetStaticPos(ByVal iWorldX As Long, ByVal iWorldY As Long, piX As Long, piY As Long, piZ As Long)
    audio_GetStaticPos iWorldX, iWorldY, piX, piY, piZ
End Sub

Private Function ISoundGame_audioObjectDead(ByVal psObj As Long) As Boolean
    ISoundGame_audioObjectDead = audio_ObjectDead(psObj)
End Function

Private Sub ISoundGame_audioGetPlayerPos(x As Single, y As Single, z As Single)
    audio_GetPlayerPos x, y, z
End Sub

Private Sub ISoundGame_audioGet3DPlayerRotAboutVerticalAxis(angle As Single)
    audio_Get3DPlayerRotAboutVerticalAxis angle
End Sub

Private Function ISoundGame_soundGetGameTime() As Long
    ISoundGame_soundGetGameTime = sound_GetGameTime()
End Function

Private Function ISoundGame_realTime() As Long
    ISoundGame_realTime = realTime
End Function

' an AUDIO_CALLBACK by its number
Private Function ISoundGame_audioCallback(ByVal fn As Long, ByVal psObj As Long) As Boolean
    Select Case fn
    Case CB_droidAudioTrackStopped: ISoundGame_audioCallback = droidAudioTrackStopped(psObj)
    Case CB_moveCheckDroidMovingAndVisible: ISoundGame_audioCallback = moveCheckDroidMovingAndVisible(psObj)
    Case CB_moveDroidStartCallback: ISoundGame_audioCallback = moveDroidStartCallback(psObj)
    Case CB_droidCheckBuildStillInProgress: ISoundGame_audioCallback = droidCheckBuildStillInProgress(psObj)
    Case CB_droidBuildStartAudioCallback: ISoundGame_audioCallback = droidBuildStartAudioCallback(psObj)
    Case Else
        debug__ 0, LOG_ERROR, "audioCallback", "(VB) unknown AUDIO_CALLBACK %d", fn
    End Select
End Function

Private Function ISoundGame_droidAudioTrackStopped(ByVal psObj As Long) As Boolean
    ISoundGame_droidAudioTrackStopped = droidAudioTrackStopped(psObj)
End Function
