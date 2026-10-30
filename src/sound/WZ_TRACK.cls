VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZ_TRACK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' cdaudio.h: struct WZ_TRACK, a music track of an album. PublicNotCreatable: the game's music manager reads
' them (cdAudio_GetCurrentTrack, PlayList_GetFullTrackList). A std::shared_ptr<const WZ_TRACK> is a WZ_TRACK
' (Nothing = nullptr). album is the std::weak_ptr<WZ_ALBUM>: a reference that playlist_c's PlayList_ClearAll
' clears, so that the album and its tracks don't keep each other alive.
Option Explicit

Public filename As String
Public title As String
Public author As String
Public base_volume As Long              ' unsigned int = 100
Public bpm As Long                      ' unsigned int = 80
Public album As WZ_ALBUM                ' std::weak_ptr<WZ_ALBUM>

Private Sub Class_Initialize()
    base_volume = 100
    bpm = 80
End Sub
