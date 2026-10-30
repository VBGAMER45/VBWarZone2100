VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StdIOProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/src/map_io.cpp: class WzMap::StdIOProvider, the default
' IOProvider (the _WIN32 branch: _wfopen, _wstat, _wmkdir and
' FindFirstFileW, with paths in UTF-16 as a VB String already is).
' StdIOProvider_new (MapApi) makes one. The helpers the C++ keeps static in
' map_io.cpp (stdIOFolderExistsInternal, enumerateDirInternal, and _wstat
' as wstat_) are in map_io_c.
Option Explicit

Implements IOProvider

Private Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long

Private Function IOProvider_openBinaryStream(ByVal filename As String, ByVal mode As OpenMode) As BinaryIOStream
    Dim pStream As WzMapBinaryStdIOStream
    Set pStream = New WzMapBinaryStdIOStream
    pStream.construct filename, modeToStdIOMode(mode)
    If Not pStream.openedFile() Then
        Set pStream = Nothing
        Set IOProvider_openBinaryStream = Nothing
        Exit Function
    End If
    Set IOProvider_openBinaryStream = pStream
End Function

Private Function IOProvider_loadFullFile(ByVal filename As String, fileData() As Byte, fileSize As Long, ByVal maxFileSize As Long, ByVal appendNullCharacter As Boolean) As LoadFullFileResult
    Dim pStream As BinaryIOStream, chunkSize As Long, data() As Byte, dataSize As Long, readPos As Long
    Dim bytesRead As Long, hasBytesRead As Boolean
    Set pStream = IOProvider_openBinaryStream(filename, OpenMode_READ)
    If pStream Is Nothing Then
        IOProvider_loadFullFile = LoadFullFileResult_FAILURE_OPEN
        Exit Function
    End If
    chunkSize = 512& * 1024&
    readPos = 0
    Do
        dataSize = dataSize + chunkSize
        ReDim Preserve data(0 To dataSize - 1)
        bytesRead = 0
        hasBytesRead = pStream.readBytes(data, readPos, chunkSize, bytesRead)
        If Not hasBytesRead Then bytesRead = 0          ' value_or(0)
        readPos = readPos + bytesRead
        If maxFileSize <> 0 And ULT(maxFileSize, readPos) Then
            ' read bytes exceeds specified maxFileSize
            IOProvider_loadFullFile = LoadFullFileResult_FAILURE_EXCEEDS_MAXFILESIZE
            Exit Function
        End If
    Loop While hasBytesRead And bytesRead = chunkSize
    If Not hasBytesRead Then
        ' failed reading
        IOProvider_loadFullFile = LoadFullFileResult_FAILURE_READ
        Exit Function
    End If
    ' truncate to exact length read (optionally with extra byte for appending null character)
    dataSize = readPos
    If appendNullCharacter Then dataSize = dataSize + 1
    If appendNullCharacter Then
        ReDim Preserve data(0 To dataSize - 1)
        data(dataSize - 1) = 0
    ElseIf dataSize > 0 Then
        ReDim Preserve data(0 To dataSize - 1)
    End If
    If dataSize > 0 Then
        fileData = data
    Else
        Erase fileData
    End If
    fileSize = dataSize
    IOProvider_loadFullFile = LoadFullFileResult_SUCCESS
End Function

Private Function IOProvider_writeFullFile(ByVal filename As String, fileData() As Byte, ByVal fileSize As Long) As Boolean
    Dim pStream As BinaryIOStream, bytesWritten As Long
    Set pStream = IOProvider_openBinaryStream(filename, OpenMode_WRITE)
    If pStream Is Nothing Then
        Exit Function
    End If
    If Not pStream.writeBytes(fileData, 0, fileSize, bytesWritten) Then
        ' failed writing
        Exit Function
    End If
    If bytesWritten <> fileSize Then
        ' failed writing all of the data
        Exit Function
    End If
    IOProvider_writeFullFile = True
End Function

Private Function internalMakeDir(ByVal dirPath As String) As Boolean
    internalMakeDir = (CreateDirectoryW(StrPtr(dirPath), 0) <> 0)    ' _wmkdir(...) == 0
End Function

' Creates a directory (and any required parent directories) and returns true on success (or if the directory already exists)
Private Function IOProvider_makeDirectory(ByVal directoryPath As String) As Boolean
    Dim result As Boolean, currSeparatorPos As Long, currentSubpath As String, sep As String
    sep = IOProvider_pathSeparator()
    currSeparatorPos = 0
    currSeparatorPos = InStr(currSeparatorPos + 2, directoryPath, sep, vbBinaryCompare) - 1   ' find(sep, currSeparatorPos + 1)
    Do
        If currSeparatorPos = -1 Then
            currentSubpath = directoryPath
        Else
            currentSubpath = Left$(directoryPath, currSeparatorPos)
        End If
        result = stdIOFolderExistsInternal(currentSubpath)
        If Not result Then
            result = internalMakeDir(currentSubpath)
            If Not result Then
                IOProvider_makeDirectory = False
                Exit Function
            End If
        End If
        If currSeparatorPos = -1 Then
            IOProvider_makeDirectory = result
            Exit Function
        End If
        currSeparatorPos = InStr(currSeparatorPos + 2, directoryPath, sep, vbBinaryCompare) - 1
    Loop
End Function

Private Function IOProvider_pathSeparator() As String
    IOProvider_pathSeparator = "\"
End Function

Private Function IOProvider_fileExists(ByVal filePath As String) As Boolean
    Dim st_mode As Long
    If wstat_(filePath, st_mode) <> 0 Then
        Exit Function
    End If
    IOProvider_fileExists = ((st_mode And S_IFREG_) = S_IFREG_)
End Function

' Check if a folder exists at a path
' NOTE: dirPath should *not* contain a trailing pathSeparator (or this may always return false)
Public Function folderExists(ByVal dirPath As String) As Boolean
    folderExists = stdIOFolderExistsInternal(dirPath)
End Function

' enumFunc receives each enumerated file, and returns true to continue enumeration, or false to shortcut / stop enumeration
Private Function IOProvider_enumerateFiles(ByVal basePath As String, files() As String, count As Long) As Boolean
    count = 0
    IOProvider_enumerateFiles = enumerateDirInternal(basePath, "", ENUM_FILES, files, count)
End Function

' enumFunc receives each enumerated subfolder, and returns true to continue enumeration, or false to shortcut / stop enumeration
Private Function IOProvider_enumerateFolders(ByVal basePath As String, folders() As String, count As Long) As Boolean
    count = 0
    IOProvider_enumerateFolders = enumerateDirInternal(basePath, "", ENUM_FOLDERS, folders, count)
End Function
