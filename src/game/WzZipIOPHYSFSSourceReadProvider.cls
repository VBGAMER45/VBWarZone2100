VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzZipIOPHYSFSSourceReadProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/wzphysfszipioprovider.cpp, .h: class WzZipIOPHYSFSSourceReadProvider, a zip archive's bytes read from a
' PhysFS file (the source WzMapZipIO reads the archive from). Made by wzphysfszipioprovider_c's
' WzZipIOPHYSFSSourceReadProvider_make (Nothing: nullptr). The optional<uint64_t> results are a Boolean with the
' value (a Double). (VB) WzMapZipIO hands the open file to PhysFS's zip archiver (PHYSFS_mountHandle), which then
' owns it: detachFile.
Option Explicit

Private fp As Long                              ' PHYSFS_File *fp = nullptr
Private fileLength As Double                    ' uint64_t
Private metaData_modtime As Double              ' PHYSFS_Stat metaData (its modtime: all that is read)
Private m_path As String                        ' (VB) the path it was opened with

Private Sub Class_Terminate()
    If fp <> 0 Then
        PHYSFS_close fp
        fp = 0
    End If
End Sub

' (VB) make's work on the new object; False: return nullptr
Friend Function init(ByVal path As String) As Boolean
    Dim f As Long, length As Double, st_filesize As Double, st_modtime As Double, filetype As Long, readonly_ As Long
    f = PHYSFS_openRead(path)
    If f = 0 Then
        init = False
        Exit Function
    End If
    fp = f
    m_path = path
    length = PHYSFS_fileLength(fp)
    If length >= 0 Then
        fileLength = length
    Else
        init = False
        Exit Function
    End If
    If PHYSFS_stat(path, st_filesize, st_modtime, filetype, readonly_) = 0 Then
        ' stat failed - non-fatal error, but make sure metaData is reinitialized
        metaData_modtime = 0
    Else
        metaData_modtime = st_modtime
    End If
    init = True
End Function

Public Function tell(value As Double) As Boolean
    Dim currentOffset As Double
    currentOffset = PHYSFS_tell(fp)
    If currentOffset < 0 Then
        tell = False
        Exit Function
    End If
    value = currentOffset
    tell = True
End Function

' seek (VB: Seek is a statement)
Public Function seek_(ByVal pos As Double) As Boolean
    seek_ = PHYSFS_seek(fp, pos) <> 0
End Function

' buffer(ofs .. ofs + len_ - 1)
Public Function readBytes(buffer() As Byte, ByVal ofs As Long, ByVal len_ As Long, value As Double) As Boolean
    Dim bytesRead As Long
    bytesRead = PHYSFS_readBytes(fp, buffer, ofs, len_)
    If bytesRead < 0 Then
        readBytes = False
        Exit Function
    End If
    value = bytesRead
    readBytes = True
End Function

Public Function fileSize(value As Double) As Boolean
    value = fileLength
    fileSize = True
End Function

Public Function modTime(value As Double) As Boolean
    If metaData_modtime > 0 Then
        value = metaData_modtime
        modTime = True
    End If
End Function

' (VB) the path it was opened with
Public Function path() As String
    path = m_path
End Function

' (VB) the PHYSFS_File, given away (WzMapZipIO: PHYSFS_mountHandle owns it once mounted)
Friend Function detachFile() As Long
    detachFile = fp
    fp = 0
End Function

' (VB) a file given back (the mount failed: PhysFS did not take it)
Friend Sub reattachFile(ByVal f As Long)
    fp = f
End Sub
