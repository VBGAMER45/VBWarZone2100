VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzAudioDataPhysFSHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sound/codecs.h: class WzAudioDataPhysFSHandle : public WZAudioDataResourceInterface, a PhysFS file the
' decoder reads (oggopus.cpp's wz_opus_physfs_* callbacks); its destructor is lib/sound/audio.cpp's
' (~WzAudioDataPhysFSHandle: closes the file). The constructor is construct (after New).
Option Explicit

Implements WZAudioDataResourceInterface

Private m_fileHandle As Long                    ' PHYSFS_file * = nullptr
Private m_fileName As String

' WzAudioDataPhysFSHandle(PHYSFS_file *pFileHandleToOwn, const std::string& fileName)
Friend Sub construct(ByVal pFileHandleToOwn As Long, ByVal fileName As String)
    m_fileHandle = pFileHandleToOwn
    m_fileName = fileName
End Sub

Public Function fileHandle() As Long
    fileHandle = m_fileHandle
End Function

' ~WzAudioDataPhysFSHandle() (audio.cpp)
Private Sub Class_Terminate()
    debug_ 60, LOG_WZ, "~WzAudioDataPhysFSHandle", "Closing: %s", m_fileName
    If m_fileHandle <> 0 Then
        PHYSFS_close m_fileHandle
        m_fileHandle = 0
    End If
End Sub
