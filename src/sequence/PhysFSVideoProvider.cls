VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PhysFSVideoProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sequence/sequence.cpp: class PhysFSVideoProvider : public VideoProvider, a video read from a PhysFS file
' (WzFramework's PHYSFS_file handle, 0 = nullptr). Private: a VideoProvider holds it (makeVideoProvider_file).
' It takes ownership of the file: closed when the object goes (Class_Terminate, the C++ destructor).
' PHYSFS_isInit() is not in FrameworkApi: after PHYSFS_deinit, physfs_win's close of the handle fails harmlessly.
Option Explicit

Private in_ As Long                     ' PHYSFS_file *in = nullptr

' Takes ownership of the PHYSFS_file*
' PhysFSVideoProvider(PHYSFS_file *in, const WzString& filename) : VideoProvider(filename), in(in)
Friend Sub construct(ByVal fileIn As Long)
    in_ = fileIn
End Sub

' ~PhysFSVideoProvider()
Private Sub Class_Terminate()
    If in_ <> 0 Then                    ' (in && PHYSFS_isInit())
        PHYSFS_close in_
        in_ = 0
    End If
End Sub

Friend Function seek_begin() As Boolean
    seek_begin = PHYSFS_seek(in_, 0) <> 0
End Function

Friend Function buffer_data(oy As ogg_sync_state) As Long
    Const size As Long = 262144
    Dim buffer As Long, bytes As Long
    If in_ = 0 Then buffer_data = 0: Exit Function

    ' read in 256K chunks
    buffer = ogg_sync_buffer(oy, size)  ' (VB) the offset into oy.data
    bytes = CLng(WZ_PHYSFS_readBytes(in_, oy.data, buffer, size))

    ogg_sync_wrote oy, bytes
    buffer_data = (bytes)
End Function

' bool PhysFSVideoProvider::PhysFSVideoProvider::end_of_data() const
Friend Function end_of_data() As Boolean
    end_of_data = PHYSFS_eof(in_) <> 0
End Function
