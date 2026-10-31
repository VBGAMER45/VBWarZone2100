VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/sequence/sequence.cpp: class VideoProvider, the abstraction for handling loading videos from files, or
' from a memory buffer. PublicNotCreatable: makeVideoProvider_file / _data (SequenceApi) give it to the game
' (a std::shared_ptr<VideoProvider> is a VideoProvider, Nothing = nullptr; VB's reference count is the
' shared_ptr's). The derived classes are WzSequence's private PhysFSVideoProvider and MemoryBufferVideoProvider,
' which this object holds (base behaviour by containment): the pure virtual methods are a choice of the one it
' holds. Their buffer_data takes the ogg_sync_state UDT, hence Friend methods (only sequence_c calls them).
Option Explicit

Private m_filename As String
Private mPhysFS As PhysFSVideoProvider                  ' (VB) the derived object: one of the two
Private mMemory As MemoryBufferVideoProvider

' VideoProvider(const WzString& filename) : m_filename(filename), as the base of a PhysFSVideoProvider
Friend Sub construct_physfs(ByVal derived As PhysFSVideoProvider, ByVal filename As String)
    m_filename = filename
    Set mPhysFS = derived
End Sub

' VideoProvider(const WzString& filename) : m_filename(filename), as the base of a MemoryBufferVideoProvider
Friend Sub construct_memory(ByVal derived As MemoryBufferVideoProvider, ByVal filename As String)
    m_filename = filename
    Set mMemory = derived
End Sub

' virtual const WzString& filename() const
Public Function filename() As String
    filename = m_filename
End Function

' virtual bool seek_begin() = 0
Friend Function seek_begin() As Boolean
    If Not mPhysFS Is Nothing Then
        seek_begin = mPhysFS.seek_begin()
    ElseIf Not mMemory Is Nothing Then
        seek_begin = mMemory.seek_begin()
    End If
End Function

' virtual int buffer_data(ogg_sync_state *oy) = 0
Friend Function buffer_data(oy As ogg_sync_state) As Long
    If Not mPhysFS Is Nothing Then
        buffer_data = mPhysFS.buffer_data(oy)
    ElseIf Not mMemory Is Nothing Then
        buffer_data = mMemory.buffer_data(oy)
    End If
End Function

' virtual bool end_of_data() const = 0
Friend Function end_of_data() As Boolean
    If Not mPhysFS Is Nothing Then
        end_of_data = mPhysFS.end_of_data()
    ElseIf Not mMemory Is Nothing Then
        end_of_data = mMemory.end_of_data()
    Else
        end_of_data = True
    End If
End Function

' virtual ~VideoProvider(): the derived object goes with the last reference
Private Sub Class_Terminate()
    Set mPhysFS = Nothing
    Set mMemory = Nothing
End Sub
