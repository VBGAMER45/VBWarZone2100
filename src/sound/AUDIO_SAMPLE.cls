VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AUDIO_SAMPLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' track.h: struct AUDIO_SAMPLE, a sound being played (audio.cpp's queue and sample list, openal_track.cpp's
' active samples). WzSound only: an AUDIO_SAMPLE * is an AUDIO_SAMPLE (Nothing = nullptr).
Option Explicit

Public iTrack As Long                   ' SDWORD: ID number identifying a specific sound; currently (r1182) mapped in audio_id.c
Public iSample As Long                  ' ALuint: OpenAL name of the sound source
Public isLooping As Boolean             ' ALboolean: if sample loops
Public is3d As Boolean                  ' ALboolean: if sample is 3d (as opposed to 2d)
Public filename As String               ' char[256]: actual filename of sample
Public x As Long                        ' SDWORD x, y, z
Public y As Long
Public z As Long
Public fVol As Single                   ' computed volume of sample
Public bFinishedPlaying As Boolean
Public pCallback As Long                ' AUDIO_CALLBACK: the game's number for it (0 = nullptr), called through ISoundGame
Public psObj As Long                    ' SIMPLE_OBJECT *: the game's object handle (0 = nullptr)
