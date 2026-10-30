VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CdAudioTrackFinished"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cdaudio.cpp's cdAudio_OpenTrack lambda [track](const AUDIO_STREAM*, const void*) { cdAudio_TrackFinished(track); }:
' the onFinished of the music's stream (cdaudio_c.CdAudioTrackFinished_new).
Option Explicit

Implements IStreamFinished

Public track As WZ_TRACK

Private Sub IStreamFinished_streamFinished(ByVal stream As AUDIO_STREAM, ByVal userData As Variant)
    cdAudio_TrackFinished track
End Sub
