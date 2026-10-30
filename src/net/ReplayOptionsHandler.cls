VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReplayOptionsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/nettypes.h: class ReplayOptionsHandler, what a replay file keeps
' of the game's options besides the messages: the exe's
' WZGameReplayOptionsHandler (src/multiint.cpp) and the tests' handlers
' Implement it (object is object_: Object is a VB keyword). The JSON values are WzFramework's nlohmann_json objects (the
' callee fills or reads the object it is given); struct EmbeddedMapData is
' its fields: dataVersion and mapBinaryData(0 .. mapBinaryDataSize - 1). The
' size_t results are Doubles (std::numeric_limits<size_t>::max() is
' 18446744073709551615#). VB refuses underscores in the method names of an
' interface used with Implements.
Option Explicit

Public Function saveOptions(ByVal object_ As nlohmann_json) As Boolean
End Function

' EmbeddedMapData& mapData: filled
Public Function saveMap(dataVersion As Long, mapBinaryData() As Byte, mapBinaryDataSize As Long) As Boolean
End Function

Public Function optionsUpdatePlayerInfo(ByVal object_ As nlohmann_json) As Boolean
End Function

' EmbeddedMapData&& embeddedMapData: given
Public Function restoreOptions(ByVal object_ As nlohmann_json, ByVal dataVersion As Long, mapBinaryData() As Byte, ByVal mapBinaryDataSize As Long, ByVal replayNetcodeMajor As Long, ByVal replayNetcodeMinor As Long) As Boolean
End Function

Public Function desiredBufferSize() As Double
End Function

Public Function maximumEmbeddedMapBufferSize() As Double
End Function
