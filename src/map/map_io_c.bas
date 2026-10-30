Attribute VB_Name = "map_io_c"
' lib/wzmaplib/src/map_io.cpp: the methods that class BinaryIOStream and
' class IOProvider implement themselves (BinaryIOStream_readULE16(stream,
' v), IOProvider_pathJoin(io, a, b)...), and StdIOProvider's static
' helpers (the _WIN32 branch). The classes: BinaryIOStream.cls,
' IOProvider.cls, StdIOProvider.cls, WzMapBinaryStdIOStream.cls.
' A uint8_t / uint16_t / uint32_t value is a Long (a uint32_t holds its 32
' bits); SDL_SwapLE16 / SDL_SwapLE32 are the byte order of the reads and
' writes. The enumeration callbacks are the lists returned
' (IOProvider.cls). _wstat is wstat_, for what StdIOProvider asks of it:
' 0 and st_mode (_S_IFDIR or _S_IFREG) for an existing path, -1 otherwise,
' and -1 for a path with a trailing separator (other than a root) or with
' a wildcard, as the C library's.
Option Explicit

Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATAW) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATAW) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATAW
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName(0 To 519) As Byte             ' WCHAR[MAX_PATH]
    cAlternateFileName(0 To 27) As Byte     ' WCHAR[14]
End Type

Public Enum EnumDirFlags
    ENUM_RECURSE = &H1
    ENUM_FILES = &H2
    ENUM_FOLDERS = &H4
End Enum

Public Const S_IFDIR_ As Long = &H4000
Public Const S_IFREG_ As Long = &H8000&

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const INVALID_FILE_ATTRIBUTES As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

'--------------------------------------------------------------------------
' BinaryIOStream

Public Function BinaryIOStream_readULE8(ByVal s As BinaryIOStream, pVal As Long) As Boolean
    Dim val_(0 To 0) As Byte, result As Long
    If Not s.readBytes(val_, 0, 1, result) Then BinaryIOStream_readULE8 = False: Exit Function
    pVal = val_(0)
    BinaryIOStream_readULE8 = (result = 1)
End Function

Public Function BinaryIOStream_readULE16(ByVal s As BinaryIOStream, pVal As Long) As Boolean
    Dim val_(0 To 1) As Byte, result As Long
    If Not s.readBytes(val_, 0, 2, result) Then BinaryIOStream_readULE16 = False: Exit Function
    pVal = val_(0) Or (CLng(val_(1)) * &H100&)
    BinaryIOStream_readULE16 = (result = 2)
End Function

Public Function BinaryIOStream_readULE32(ByVal s As BinaryIOStream, pVal As Long) As Boolean
    Dim val_(0 To 3) As Byte, result As Long
    If Not s.readBytes(val_, 0, 4, result) Then BinaryIOStream_readULE32 = False: Exit Function
    CopyMemory pVal, val_(0), 4
    BinaryIOStream_readULE32 = (result = 4)
End Function

Public Function BinaryIOStream_readSLE8(ByVal s As BinaryIOStream, pVal As Long) As Boolean
    Dim val_(0 To 0) As Byte, result As Long
    If Not s.readBytes(val_, 0, 1, result) Then BinaryIOStream_readSLE8 = False: Exit Function
    pVal = val_(0)
    If pVal >= 128 Then pVal = pVal - 256
    BinaryIOStream_readSLE8 = (result = 1)
End Function

Public Function BinaryIOStream_readSLE16(ByVal s As BinaryIOStream, pVal As Long) As Boolean
    Dim val_(0 To 1) As Byte, result As Long
    If Not s.readBytes(val_, 0, 2, result) Then BinaryIOStream_readSLE16 = False: Exit Function
    pVal = val_(0) Or (CLng(val_(1)) * &H100&)
    If pVal >= 32768 Then pVal = pVal - 65536
    BinaryIOStream_readSLE16 = (result = 2)
End Function

Public Function BinaryIOStream_readSLE32(ByVal s As BinaryIOStream, pVal As Long) As Boolean
    Dim val_(0 To 3) As Byte, result As Long
    If Not s.readBytes(val_, 0, 4, result) Then BinaryIOStream_readSLE32 = False: Exit Function
    CopyMemory pVal, val_(0), 4
    BinaryIOStream_readSLE32 = (result = 4)
End Function

' n bytes of v (little-endian, or big-endian), written
Private Function WriteBytesOf(ByVal s As BinaryIOStream, ByVal v As Long, ByVal n As Long, ByVal bigEndian As Boolean) As Boolean
    Dim le(0 To 3) As Byte, val_(0 To 3) As Byte, i As Long, result As Long
    CopyMemory le(0), v, 4
    For i = 0 To n - 1
        If bigEndian Then val_(i) = le(n - 1 - i) Else val_(i) = le(i)
    Next
    If Not s.writeBytes(val_, 0, n, result) Then WriteBytesOf = False: Exit Function
    WriteBytesOf = (result = n)
End Function

Public Function BinaryIOStream_writeULE8(ByVal s As BinaryIOStream, ByVal val_ As Long) As Boolean
    BinaryIOStream_writeULE8 = WriteBytesOf(s, val_, 1, False)
End Function

Public Function BinaryIOStream_writeULE16(ByVal s As BinaryIOStream, ByVal val_ As Long) As Boolean
    BinaryIOStream_writeULE16 = WriteBytesOf(s, val_, 2, False)
End Function

Public Function BinaryIOStream_writeULE32(ByVal s As BinaryIOStream, ByVal val_ As Long) As Boolean
    BinaryIOStream_writeULE32 = WriteBytesOf(s, val_, 4, False)
End Function

Public Function BinaryIOStream_writeSLE8(ByVal s As BinaryIOStream, ByVal val_ As Long) As Boolean
    BinaryIOStream_writeSLE8 = WriteBytesOf(s, val_, 1, False)
End Function

Public Function BinaryIOStream_writeSLE16(ByVal s As BinaryIOStream, ByVal val_ As Long) As Boolean
    BinaryIOStream_writeSLE16 = WriteBytesOf(s, val_, 2, False)
End Function

Public Function BinaryIOStream_writeSLE32(ByVal s As BinaryIOStream, ByVal val_ As Long) As Boolean
    BinaryIOStream_writeSLE32 = WriteBytesOf(s, val_, 4, False)
End Function

Public Function BinaryIOStream_writeUBE32(ByVal s As BinaryIOStream, ByVal val_ As Long) As Boolean
    BinaryIOStream_writeUBE32 = WriteBytesOf(s, val_, 4, True)
End Function

Public Function BinaryIOStream_writeSBE32(ByVal s As BinaryIOStream, ByVal val_ As Long) As Boolean
    BinaryIOStream_writeSBE32 = WriteBytesOf(s, val_, 4, True)
End Function

'--------------------------------------------------------------------------
' IOProvider

' std::string::rfind(subStr, pos): the index (from 0) of the last subStr
' that starts at or before pos; pos < 0 is npos (and so is the result when
' there is none)
Private Function rfind_(ByVal s As String, ByVal subStr As String, ByVal pos As Long) As Long
    If Len(subStr) > Len(s) Then
        rfind_ = -1
        Exit Function
    End If
    If pos < 0 Or pos > Len(s) - Len(subStr) Then pos = Len(s) - Len(subStr)
    If Len(subStr) = 0 Then
        rfind_ = pos
        Exit Function
    End If
    rfind_ = InStrRev(s, subStr, pos + Len(subStr), vbBinaryCompare) - 1
End Function

Private Sub AddName(names() As String, count As Long, ByVal s As String)
    If count = 0 Then
        ReDim names(0 To 15)
    ElseIf count > UBound(names) Then
        ReDim Preserve names(0 To 2 * count - 1)
    End If
    names(count) = s
    count = count + 1
End Sub

Public Function IOProvider_pathJoin(ByVal io As IOProvider, ByVal a As String, ByVal b As String) As String
    Dim fullPath As String, separator As String
    fullPath = a
    If Len(b) = 0 Then
        IOProvider_pathJoin = fullPath
        Exit Function
    End If
    separator = io.pathSeparator()
    If Len(fullPath) > 0 Then
        If Not strEndsWith(fullPath, separator) Then
            fullPath = fullPath & separator
        End If
    End If
    fullPath = fullPath & b
    IOProvider_pathJoin = fullPath
End Function

Public Function IOProvider_pathDirName(ByVal io As IOProvider, ByVal fullPath As String) As String
    Dim pathSep As String, dirnameLength As Long, lastPos As Long
    ' Split the path at the path separator
    pathSep = io.pathSeparator()
    If Len(pathSep) = 0 Then
        IOProvider_pathDirName = fullPath
        Exit Function
    End If
    dirnameLength = Len(fullPath)
    lastPos = dirnameLength
    Do
        lastPos = dirnameLength - Len(pathSep)          ' below 0: size_t wraps (npos)
        dirnameLength = rfind_(fullPath, pathSep, lastPos)
    Loop While dirnameLength <> -1 And dirnameLength = lastPos
    If dirnameLength <> -1 Then IOProvider_pathDirName = Left$(fullPath, dirnameLength) Else IOProvider_pathDirName = ""
End Function

Public Function IOProvider_pathBaseName(ByVal io As IOProvider, ByVal fullPath As String) As String
    Dim pathSep As String, dirnameLength As Long, lastPos As Long
    ' Split the path at the path separator
    pathSep = io.pathSeparator()
    If Len(pathSep) = 0 Then
        IOProvider_pathBaseName = fullPath
        Exit Function
    End If
    dirnameLength = Len(fullPath)
    lastPos = dirnameLength
    Do
        lastPos = dirnameLength - Len(pathSep)
        dirnameLength = rfind_(fullPath, pathSep, lastPos)
    Loop While dirnameLength <> -1 And dirnameLength = lastPos
    If dirnameLength <> -1 Then IOProvider_pathBaseName = Mid$(fullPath, dirnameLength + 2) Else IOProvider_pathBaseName = fullPath
End Function

' A base implementation of recursive enumeration that utilizes the regular enumerateFiles/Folders functions.
' The files, with their paths relative to basePath, in files(0 .. count - 1)
Public Function IOProvider_enumerateFilesRecursive(ByVal io As IOProvider, ByVal basePath As String, files() As String, count As Long) As Boolean
    Dim names() As String, n As Long, i As Long, folders() As String, nf As Long, f As Long
    Dim enumSuccess As Boolean, enumFilesSuccess As Boolean
    count = 0
    ' Step 1: Enumerate all the files in this folder
    enumSuccess = io.enumerateFiles(basePath, names, n)
    For i = 0 To n - 1
        AddName files, count, names(i)
    Next
    If Not enumSuccess Then
        IOProvider_enumerateFilesRecursive = False
        Exit Function
    End If
    ' Step 2: Enumerate all the folders in this folder, enumerating all the files in each
    enumSuccess = IOProvider_enumerateFoldersRecursive(io, basePath, folders, nf)
    For f = 0 To nf - 1
        enumFilesSuccess = io.enumerateFiles(IOProvider_pathJoin(io, basePath, folders(f)), names, n)
        For i = 0 To n - 1
            AddName files, count, IOProvider_pathJoin(io, folders(f), names(i))
        Next
        If Not enumFilesSuccess Then
            ' stop enumerating folders (enumerateFoldersRecursive then returns true)
            IOProvider_enumerateFilesRecursive = True
            Exit Function
        End If
    Next
    If Not enumSuccess Then
        IOProvider_enumerateFilesRecursive = False
        Exit Function
    End If
    IOProvider_enumerateFilesRecursive = True
End Function

' A base implementation of recursive enumeration that utilizes the regular enumerateFolders function.
' The folders, with their paths relative to basePath, in folders(0 .. count - 1)
Public Function IOProvider_enumerateFoldersRecursive(ByVal io As IOProvider, ByVal basePath As String, folders() As String, count As Long) As Boolean
    Dim foldersList As StringVector, currentPath As String, subfolders() As String, n As Long, i As Long
    Dim fullRelativePath As String, it As Long, enumSuccess As Boolean
    count = 0
    StringVector_push_back foldersList, ""      ' "" is used for basePath
    Do While foldersList.size > 0
        ' enumerate the last folder in the foldersList
        it = foldersList.size - 1
        currentPath = foldersList.items(it)
        enumSuccess = io.enumerateFolders(IOProvider_pathJoin(io, basePath, currentPath), subfolders, n)
        For i = 0 To n - 1
            fullRelativePath = currentPath & subfolders(i)
            ' call enumFunc with the enumerated subfolder
            AddName folders, count, fullRelativePath
            ' push this subfolder onto the stack
            StringVector_push_back foldersList, fullRelativePath
        Next
        ' remove the item just processed from the stack
        StringVector_erase foldersList, it
        If Not enumSuccess Then
            IOProvider_enumerateFoldersRecursive = False
            Exit Function
        End If
    Loop
    IOProvider_enumerateFoldersRecursive = True
End Function

'--------------------------------------------------------------------------
' StdIOProvider's helpers

' C: _wstat(path, &buf): 0 and buf.st_mode's file type, or -1
Public Function wstat_(ByVal path As String, st_mode As Long) As Long
    Dim attr As Long, lastCh As String
    st_mode = 0
    wstat_ = -1
    If Len(path) = 0 Then Exit Function
    If InStr(path, "*") > 0 Or InStr(path, "?") > 0 Then Exit Function
    lastCh = Right$(path, 1)
    If Len(path) > 1 And (lastCh = "\" Or lastCh = "/") Then
        If Not (Len(path) = 3 And Mid$(path, 2, 1) = ":") Then Exit Function
    End If
    attr = GetFileAttributesW(StrPtr(path))
    If attr = INVALID_FILE_ATTRIBUTES Then Exit Function
    If (attr And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then st_mode = S_IFDIR_ Else st_mode = S_IFREG_
    wstat_ = 0
End Function

Public Function stdIOFolderExistsInternal(ByVal dirPath As String) As Boolean
    Dim st_mode As Long
    If wstat_(dirPath, st_mode) <> 0 Then
        Exit Function
    End If
    stdIOFolderExistsInternal = ((st_mode And S_IFDIR_) = S_IFDIR_)
End Function

' The names enumFunc would receive (basePath & "\" & name), appended to names(0 .. count - 1)
Public Function enumerateDirInternal(ByVal rootBasePath As String, ByVal basePath As String, ByVal dirFlags As Long, names() As String, count As Long) As Boolean
    Dim findFileStr As String, ffd As WIN32_FIND_DATAW, hFind As Long, fileName_ As String, fullFilePath As String
    ' List files inside the basePath directory (recursing into subdirectories)
    findFileStr = rootBasePath
    If Len(findFileStr) > 0 Then
        If Right$(findFileStr, 1) <> "\" Then findFileStr = findFileStr & "\"
    End If
    If Len(basePath) > 0 Then
        findFileStr = findFileStr & basePath
        If Right$(basePath, 1) <> "\" Then findFileStr = findFileStr & "\"
    End If
    findFileStr = findFileStr & "*"
    hFind = FindFirstFileW(StrPtr(findFileStr), ffd)
    If hFind = INVALID_HANDLE_VALUE Then
        enumerateDirInternal = False
        Exit Function
    End If
    Do
        fileName_ = String$(lstrlenW(VarPtr(ffd.cFileName(0))), 0)
        If Len(fileName_) > 0 Then CopyMemory ByVal StrPtr(fileName_), ffd.cFileName(0), Len(fileName_) * 2
        If fileName_ = "." Or fileName_ = ".." Then GoTo continue_
        fullFilePath = basePath & "\" & fileName_
        If (ffd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
            ' found a file
            If (dirFlags And ENUM_FILES) = ENUM_FILES Then
                AddName names, count, fullFilePath
            End If
        Else
            If (dirFlags And ENUM_FOLDERS) = ENUM_FOLDERS Then
                AddName names, count, fullFilePath
            End If
            If (dirFlags And ENUM_RECURSE) = ENUM_RECURSE Then
                ' recurse into subfolder
                enumerateDirInternal rootBasePath, fullFilePath, dirFlags, names, count
            End If
        End If
continue_:
    Loop While FindNextFileW(hFind, ffd) <> 0
    FindClose hFind
    enumerateDirInternal = True
End Function

Public Function modeToStdIOMode(ByVal mode As OpenMode) As String
    Select Case mode
    Case OpenMode_READ
        modeToStdIOMode = "r"
        Exit Function
    Case OpenMode_WRITE
        modeToStdIOMode = "w"
        Exit Function
    End Select
    modeToStdIOMode = "r"   ' silence warning
End Function
