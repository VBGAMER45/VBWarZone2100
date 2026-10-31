VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZAudioDataResourceInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sound/codecs.h: class WZAudioDataResourceInterface, the audio data a WZOpusDecoder owns (its
' std::unique_ptr<WZAudioDataResourceInterface> m_data): an interface with no member (the C++ class has only
' its constructor and virtual destructor, audio.cpp's, which do nothing), implemented by WZAudioDataMemoryBuffer
' and WzAudioDataPhysFSHandle. The last reference going is unique_ptr's delete.
Option Explicit
