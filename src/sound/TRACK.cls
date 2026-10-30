VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRACK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' track.h: struct TRACK, a loaded sound effect. PublicNotCreatable: the game's resource module holds the
' TRACK that sound_LoadTrackFromFile returns (data.cpp's dataAudioLoad) and hands it back to
' sound_ReleaseTrack. A TRACK * is a TRACK (Nothing = nullptr).
Option Explicit

Public bLoop As Boolean
Public iVol As Long
Public iAudibleRadius As Long
Public iTime As Long                    ' duration in milliseconds
Public iTimeLastFinished As Long        ' UDWORD: time last finished in ms
Public iNumPlaying As Long              ' UDWORD
Public iBufferName As Long              ' ALuint: OpenAL name of the buffer
Public fileName As String               ' const char *: vbNullString = nullptr
