VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICDAudioEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' cdaudio.h: class CDAudioEventSink, what cdAudio_RegisterForEvents tells of the music (the game's music
' manager implements it). The base class's empty bodies: an implementation leaves those methods empty.
' track: a std::shared_ptr<const WZ_TRACK> (Nothing = nullptr).
Option Explicit

Public Sub startedPlayingTrack(ByVal track As WZ_TRACK)
End Sub

Public Sub trackEnded(ByVal track As WZ_TRACK)
End Sub

Public Sub musicStopped()
End Sub

Public Sub musicPaused(ByVal track As WZ_TRACK)
End Sub

Public Sub musicResumed(ByVal track As WZ_TRACK)
End Sub

' virtual bool unregisterEventSink() const = 0
Public Function unregisterEventSink() As Boolean
End Function
