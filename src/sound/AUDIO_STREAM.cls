VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AUDIO_STREAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' openal_track.cpp: struct AUDIO_STREAM, a sound streamed from its file (the music, the briefings' voice).
' PublicNotCreatable: audio_PlayStream returns it to the game (an AUDIO_STREAM * is an AUDIO_STREAM,
' Nothing = nullptr); only WzSound reads its fields.
Option Explicit

Public source As Long                   ' ALuint = -1: OpenAL name of the sound source
Public decoder As Object                ' WZDecoder * = nullptr: a WZDecoder (WZVorbisDecoder or WZOpusDecoder; Nothing = nullptr)
Public fileHandle As Long               ' PHYSFS_file * = nullptr (0)
Public volume As Single                 ' = 0.f
Public queuedStop As Boolean            ' = false: when sound_StopStream has been called on the stream

' Callbacks
Public onFinished As IStreamFinished    ' std::function: Nothing when empty
Public user_data As Variant             ' const void * = nullptr (Empty)

Private Sub Class_Initialize()
    source = -1
End Sub
