VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMapZipIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/wzmaplib/plugins/ZipIOProvider: class WzMapZipIO, the WzMap::IOProvider over a zip archive (a .wz map
' package), read-only as openZipArchiveReadIOProvider opens it. 4.7.0 reads the archive with libzip from a
' WzZipIOSourceReadProvider; (VB) libzip is not ported: the source's PhysFS file is mounted with PhysFS's zip
' archiver (PHYSFS_mountHandle, under a name and a mount point no search path entry has: nothing else sees it)
' and read through a WzMapPhysFSIO based at that mount point; the archive is unmounted (and the file closed)
' with the object. Made by wzphysfszipioprovider_c's WzMapZipIO_openZipArchiveReadIOProvider.
Option Explicit

Implements IOProvider

Private m_io As IOProvider                      ' (VB) a WzMapPhysFSIO at m_mountPoint
Private m_mountName As String                   ' (VB) the archive's name in the search path
Private m_mountPoint As String
Private m_mounted As Boolean

Private Sub Class_Terminate()
    If m_mounted Then
        If PHYSFS_unmount(m_mountName) = 0 Then
            debug_ 0, LOG_ERROR, "~WzMapZipIO", "(VB) Failed to unmount %s: %s", m_mountName, WZ_PHYSFS_getLastError()
        End If
        m_mounted = False
    End If
End Sub

' (VB) the archive in zipSource mounted; False: it is not a zip archive PhysFS reads (the source keeps its file)
Friend Function openFrom(ByVal zipSource As WzZipIOPHYSFSSourceReadProvider, ByVal uid As Long) As Boolean
    Dim f As Long, physIO As WzMapPhysFSIO
    m_mountName = "WzMapZipIO:" & uid & ":" & zipSource.path()
    m_mountPoint = "WzMapZipIO/" & uid
    f = zipSource.detachFile()
    If PHYSFS_mountHandle(f, m_mountName, m_mountPoint, PHYSFS_APPEND) = 0 Then
        zipSource.reattachFile f
        openFrom = False
        Exit Function
    End If
    m_mounted = True
    Set physIO = New WzMapPhysFSIO
    physIO.construct m_mountPoint
    Set m_io = physIO
    openFrom = True
End Function

Private Function IOProvider_openBinaryStream(ByVal filename As String, ByVal mode As OpenMode) As BinaryIOStream
    If mode <> OpenMode_READ Then
        ' (a read-only archive)
        Set IOProvider_openBinaryStream = Nothing
        Exit Function
    End If
    Set IOProvider_openBinaryStream = m_io.openBinaryStream(filename, mode)
End Function

Private Function IOProvider_loadFullFile(ByVal filename As String, fileData() As Byte, fileSize As Long, ByVal maxFileSize As Long, ByVal appendNullCharacter As Boolean) As LoadFullFileResult
    IOProvider_loadFullFile = m_io.loadFullFile(filename, fileData, fileSize, maxFileSize, appendNullCharacter)
End Function

Private Function IOProvider_writeFullFile(ByVal filename As String, fileData() As Byte, ByVal fileSize As Long) As Boolean
    IOProvider_writeFullFile = False                ' (a read-only archive)
End Function

Private Function IOProvider_makeDirectory(ByVal directoryPath As String) As Boolean
    IOProvider_makeDirectory = False                ' (a read-only archive)
End Function

Private Function IOProvider_pathSeparator() As String
    IOProvider_pathSeparator = "/"
End Function

Private Function IOProvider_fileExists(ByVal filePath As String) As Boolean
    IOProvider_fileExists = m_io.fileExists(filePath)
End Function

Private Function IOProvider_enumerateFiles(ByVal basePath As String, files() As String, count As Long) As Boolean
    IOProvider_enumerateFiles = m_io.enumerateFiles(basePath, files, count)
End Function

Private Function IOProvider_enumerateFolders(ByVal basePath As String, folders() As String, count As Long) As Boolean
    IOProvider_enumerateFolders = m_io.enumerateFolders(basePath, folders, count)
End Function
