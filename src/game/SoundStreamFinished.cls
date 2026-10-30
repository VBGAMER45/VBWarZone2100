VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundStreamFinished"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) a std::function<void (const AUDIO_STREAM *, const void *)> the game hands to audio_PlayStream (WzSound's
' IStreamFinished): its WzFunc, called with the stream and the user data (aud_c's audio_StreamFinished_new).
Option Explicit

Implements IStreamFinished

Public f As WzFunc

Private Sub IStreamFinished_streamFinished(ByVal stream As AUDIO_STREAM, ByVal userData As Variant)
    WzFunc_call f, stream, userData
End Sub
