VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sound/codecs.h: class WZDecoder, the abstract decoder openal_track.cpp streams from: an interface,
' implemented by WZVorbisDecoder (oggvorbis.cpp) and WZOpusDecoder (oggopus.cpp). A private class of WzSound
' (nothing outside it holds a decoder; AUDIO_STREAM.decoder is an Object). The C++ destructor is the
' implementing class's Class_Terminate. decode's optional<size_t>: false for nullopt, the bytes in written.
Option Explicit

' Decode all the data into a buffer.
' Note that the API is different from oggvorbis: opus takes in, and gives out *samples* not bytes
' and we have to convert to little endian ourselves
'
' \param [out] buffer: preallocated buffer to hold decoded data, little-endian. Will *not* be freed, even on error.
' \param [in] bufferSize: takes buffer size in *bytes* as input, to be consistent with oggvorbis,
' \returns how much bytes have been placed into buffer, or nullopt  if error
Public Function decode(buffer() As Byte, ByVal bufferSize As Long, written As Long) As Boolean
End Function

Public Function totalTime() As Double
End Function

Public Function channels() As Long
End Function

' Sampling rate of the bitstream
Public Function frequency() As Long
End Function
