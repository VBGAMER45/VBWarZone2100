VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMapPhysFSIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/map.cpp, map.h: class WzMapPhysFSIO, the WzMap::IOProvider the game
' gives wzmaplib: its files are PhysFS's. New WzMapPhysFSIO is the default
' constructor; construct(baseMountPath) the other. The base class's
' pathJoin is MapApi's IOProvider_pathJoin.
Option Explicit

Implements IOProvider

Private m_basePath As String

Friend Sub construct(ByVal baseMountPath As String)
    m_basePath = baseMountPath
End Sub

Private Function IOProvider_openBinaryStream(ByVal filename As String, ByVal mode As OpenMode) As BinaryIOStream
    Dim pStream As WzMapBinaryPhysFSStream
    Set pStream = New WzMapBinaryPhysFSStream
    If Len(m_basePath) = 0 Then
        pStream.construct filename, mode
    Else
        pStream.construct IOProvider_pathJoin(Me, m_basePath, filename), mode
    End If
    If Not pStream.openedFile() Then
        Set pStream = Nothing
        Set IOProvider_openBinaryStream = Nothing
        Exit Function
    End If
    Set IOProvider_openBinaryStream = pStream
End Function

Private Function IOProvider_loadFullFile(ByVal filename As String, fileData() As Byte, fileSize As Long, ByVal maxFileSize As Long, ByVal appendNullCharacter As Boolean) As LoadFullFileResult
    Dim filenameFull As String
    If Len(m_basePath) = 0 Then filenameFull = filename Else filenameFull = IOProvider_pathJoin(Me, m_basePath, filename)
    If PHYSFS_exists(filenameFull) = 0 Then
        IOProvider_loadFullFile = LoadFullFileResult_FAILURE_OPEN
        Exit Function
    End If
    If Not loadFileToBufferVector(filenameFull, fileData, fileSize, True, appendNullCharacter) Then
        IOProvider_loadFullFile = LoadFullFileResult_FAILURE_OPEN
        Exit Function
    End If
    IOProvider_loadFullFile = LoadFullFileResult_SUCCESS
End Function

Private Function IOProvider_writeFullFile(ByVal filename As String, fileData() As Byte, ByVal fileSize As Long) As Boolean
    Dim filenameFull As String
    If Len(m_basePath) = 0 Then filenameFull = filename Else filenameFull = IOProvider_pathJoin(Me, m_basePath, filename)
    IOProvider_writeFullFile = saveFile(filenameFull, fileData, fileSize)
End Function

Private Function IOProvider_makeDirectory(ByVal directoryPath As String) As Boolean
    Dim directoryPathFull As String
    If Len(m_basePath) = 0 Then directoryPathFull = directoryPath Else directoryPathFull = IOProvider_pathJoin(Me, m_basePath, directoryPath)
    IOProvider_makeDirectory = PHYSFS_mkdir(directoryPath) <> 0    ' (C++: directoryPath, not directoryPathFull)
End Function

Private Function IOProvider_pathSeparator() As String
    IOProvider_pathSeparator = "/" ' the platform-independent PhysFS path separator
End Function

Private Function IOProvider_fileExists(ByVal filename As String) As Boolean
    Dim filenameFull As String, filesize As Double, modtime As Double, filetype As Long, readonly_ As Long
    If Len(m_basePath) = 0 Then filenameFull = filename Else filenameFull = IOProvider_pathJoin(Me, m_basePath, filename)
    If PHYSFS_stat(filenameFull, filesize, modtime, filetype, readonly_) = 0 Then
        ' PHYSFS_stat failed
        IOProvider_fileExists = False
        Exit Function
    End If
    IOProvider_fileExists = (filetype = PHYSFS_FILETYPE_REGULAR)
End Function

Public Function folderExists(ByVal dirPath As String) As Boolean
    Dim dirPathFull As String
    If Len(m_basePath) = 0 Then dirPathFull = dirPath Else dirPathFull = IOProvider_pathJoin(Me, m_basePath, dirPath)
    folderExists = WZ_PHYSFS_isDirectory(dirPathFull) <> 0
End Function

' (the C++ callback skips empty names and directories, and passes the rest on)
Private Function IOProvider_enumerateFiles(ByVal basePath As String, files() As String, count As Long) As Boolean
    Dim basePathFull As String, all() As String, n As Long, i As Long, fullPath As String
    If Len(m_basePath) = 0 Then basePathFull = basePath Else basePathFull = IOProvider_pathJoin(Me, m_basePath, basePath)
    IOProvider_enumerateFiles = WZ_PHYSFS_enumerateFiles(basePathFull, all, n)
    count = 0
    If n > 0 Then ReDim files(0 To n - 1)
    For i = 0 To n - 1
        If Len(all(i)) = 0 Then GoTo continue_
        fullPath = basePathFull & "/" & all(i)
        If WZ_PHYSFS_isDirectory(fullPath) <> 0 Then
            GoTo continue_ ' skip and continue
        End If
        files(count) = all(i)
        count = count + 1
continue_:
    Next
End Function

Private Function IOProvider_enumerateFolders(ByVal basePath As String, folders() As String, count As Long) As Boolean
    Dim basePathFull As String
    If Len(m_basePath) = 0 Then basePathFull = basePath Else basePathFull = IOProvider_pathJoin(Me, m_basePath, basePath)
    IOProvider_enumerateFolders = WZ_PHYSFS_enumerateFolders(basePathFull, folders, count)
End Function
