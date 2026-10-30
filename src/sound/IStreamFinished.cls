VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStreamFinished"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' (VB) the std::function<void (const AUDIO_STREAM *, const void *)> onFinished of sound_PlayStream and
' audio_PlayStream: an object implementing this (Nothing = an empty std::function). cdaudio.cpp's
' cdAudio_TrackFinished is a private class of WzSound; the game's lambdas (intelmap.cpp) are an exe class
' that calls its WzFunc. user_data is the const void * handed to sound_PlayStream, as it was given.
Option Explicit

Public Sub streamFinished(ByVal stream As AUDIO_STREAM, ByVal userData As Variant)
End Sub
