VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SequenceApi"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' WzSequence.dll's public interface: lib/sequence's sequence.h. GlobalMultiUse: the exe calls these
' procedures as if they were its own. lib/sequence is ported into this DLL's modules (sequence_c, timer_c;
' libtheora's decoder in the theora_* modules); this class only forwards to them. A
' std::shared_ptr<VideoProvider> is a VideoProvider (Nothing = nullptr).
Option Explicit

' sequence.h: SCANLINE_MODE
Public Enum SCANLINE_MODE
    SCANLINES_OFF
    SCANLINES_50
    SCANLINES_BLACK
End Enum

' Name and version of this DLL and of the WzFramework it uses, for the log header.
Public Function SequenceBuildInfo() As String
    SequenceBuildInfo = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision & " (uses " & FrameworkBuildInfo() & ")"
End Function

' Abstraction for handling loading videos from files, or from memory buffer, etc
' std::shared_ptr<VideoProvider> makeVideoProvider(PHYSFS_file *in, const WzString& filename): fileHandle is
' WzFramework's PHYSFS_file handle; the provider takes ownership of it (closes it when it goes)
Public Function makeVideoProvider_file(ByVal fileHandle As Long, ByVal filename As String) As VideoProvider
    Set makeVideoProvider_file = sequence_c.makeVideoProvider_file(fileHandle, filename)
End Function

' std::shared_ptr<VideoProvider> makeVideoProvider(std::shared_ptr<const std::vector<char>> memoryBuffer, const
' WzString& filename): the provider keeps a copy of data (all its elements)
Public Function makeVideoProvider_data(data() As Byte, ByVal filename As String) As VideoProvider
    Set makeVideoProvider_data = sequence_c.makeVideoProvider_data(data, filename)
End Function

Public Function seq_Play(ByVal video As VideoProvider) As Boolean
    seq_Play = sequence_c.seq_Play(video)
End Function

Public Function seq_Playing() As Boolean
    seq_Playing = sequence_c.seq_Playing()
End Function

Public Function seq_Update() As Boolean
    seq_Update = sequence_c.seq_Update()
End Function

Public Sub seq_Shutdown()
    sequence_c.seq_Shutdown
End Sub

Public Function seq_GetFrameNumber() As Long
    seq_GetFrameNumber = sequence_c.seq_GetFrameNumber()
End Function

Public Sub seq_SetDisplaySize(ByVal sizeX As Long, ByVal sizeY As Long, ByVal posX As Long, ByVal posY As Long)
    sequence_c.seq_SetDisplaySize sizeX, sizeY, posX, posY
End Sub

Public Sub seq_setScanlinesDisabled(ByVal flag As Boolean)
    sequence_c.seq_setScanlinesDisabled flag
End Sub

Public Function seq_getScanlinesDisabled() As Boolean
    seq_getScanlinesDisabled = sequence_c.seq_getScanlinesDisabled()
End Function

' mode: a SCANLINE_MODE
Public Sub seq_setScanlineMode(ByVal mode As Long)
    sequence_c.seq_setScanlineMode mode
End Sub

' a SCANLINE_MODE
Public Function seq_getScanlineMode() As Long
    seq_getScanlineMode = sequence_c.seq_getScanlineMode()
End Function

Public Function seq_GetFrameTime() As Double
    seq_GetFrameTime = sequence_c.seq_GetFrameTime()
End Function
