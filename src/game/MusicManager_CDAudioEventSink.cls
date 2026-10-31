VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicManager_CDAudioEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/musicmanager.cpp: class MusicManager_CDAudioEventSink (a CDAudioEventSink: WzSound's ICDAudioEventSink), what
' the Music Manager form hears of the music. std::weak_ptr<MusicManagerForm> weakMusicManagerForm is the form's
' widgweak_vb serial. Made by musicmanager_c's MusicManagerForm_initialize.
Option Explicit

Implements ICDAudioEventSink

Private weakMusicManagerForm As Long       ' std::weak_ptr<MusicManagerForm>
Private shouldUnregisterEventSink As Boolean ' = false

' MusicManager_CDAudioEventSink(const std::shared_ptr<MusicManagerForm>& musicManagerForm)
Friend Sub construct(ByVal musicManagerFormSerial As Long)
    weakMusicManagerForm = musicManagerFormSerial
    shouldUnregisterEventSink = False
End Sub

Friend Sub setUnregisterEventSink()
    shouldUnregisterEventSink = True
End Sub

Private Function ICDAudioEventSink_unregisterEventSink() As Boolean
    ICDAudioEventSink_unregisterEventSink = shouldUnregisterEventSink
End Function

Private Sub CDAudioEvent_UpdateCurrentTrack(ByVal track As WZ_TRACK)
    Dim psMusicManagerForm As WIDGET
    Set psMusicManagerForm = WeakWidget_lock(weakMusicManagerForm)
    If Not psMusicManagerForm Is Nothing Then
        MusicManagerForm_updateActiveTrack psMusicManagerForm, track
    End If
End Sub

Private Sub ICDAudioEventSink_startedPlayingTrack(ByVal track As WZ_TRACK)
    CDAudioEvent_UpdateCurrentTrack track
End Sub

Private Sub ICDAudioEventSink_trackEnded(ByVal track As WZ_TRACK)
    ' currently a no-op
End Sub

Private Sub ICDAudioEventSink_musicStopped()
    CDAudioEvent_UpdateCurrentTrack Nothing
End Sub

Private Sub ICDAudioEventSink_musicPaused(ByVal track As WZ_TRACK)
    CDAudioEvent_UpdateCurrentTrack track
End Sub

Private Sub ICDAudioEventSink_musicResumed(ByVal track As WZ_TRACK)
    CDAudioEvent_UpdateCurrentTrack track
End Sub
