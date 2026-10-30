Attribute VB_Name = "physfs_win"
' PhysicsFS for the port, on Win32 (no C++ counterpart): the part of the
' PHYSFS_ API Warzone uses, over plain folders: init, the write directory,
' mkdir and delete, the search path (mount, unmount, mount points), opening,
' reading, writing and seeking files, stat, getRealDir, enumerateFiles, the
' base and pref dirs, and WZ_PHYSFS_getLastError (physfs_ext.h). .wz archives
' (zip) come later. A PHYSFS_File* is an index into files() (0 = NULL).
' Names are PhysFS's platform-independent notation ("dir/file"), checked as
' PhysFS does (no "." or ".." elements, no ":" or "\"); folders on disk use
' "\". enumerateFiles returns its list sorted (strcmp order) and without
' duplicates, as PhysFS 2.1 does.
Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FILE_ATTRIBUTE_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
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
    cFileName(0 To 519) As Byte
    cAlternateFileName(0 To 27) As Byte
End Type

Private Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileAttributesExW Lib "kernel32" (ByVal lpFileName As Long, ByVal fInfoLevelId As Long, lpFileInformation As WIN32_FILE_ATTRIBUTE_DATA) As Long
Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFilename As Long, ByVal nSize As Long) As Long
Private Declare Function SHGetFolderPathW Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATAW) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATAW) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const INVALID_FILE_ATTRIBUTES As Long = -1
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1
Private Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400
Private Const CSIDL_APPDATA As Long = &H1A
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const OPEN_ALWAYS As Long = 4
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_BEGIN As Long = 0
Private Const FILE_END As Long = 2

' PhysFS's messages for the error codes used here
Private Const ERR_NOT_INITIALIZED As String = "not initialized"
Private Const ERR_NOT_FOUND As String = "not found"
Private Const ERR_NO_WRITE_DIR As String = "write directory is not set"
Private Const ERR_IO As String = "i/o error"
Private Const ERR_BAD_FILENAME As String = "filename is illegal or insecure"
Private Const ERR_UNSUPPORTED As String = "unsupported"
Private Const ERR_INVALID_ARGUMENT As String = "invalid argument"
Private Const ERR_OPEN_FOR_READING As String = "file open for reading"
Private Const ERR_OPEN_FOR_WRITING As String = "file open for writing"
Private Const ERR_PAST_EOF As String = "past end of file"
Private Const ERR_NOT_MOUNTED As String = "not mounted"

Private Type PHYSFS_FileHandle
    handle As Long
    forReading As Boolean
    inUse As Boolean
End Type

Private initialised As Boolean
Private baseDir As String
Private writeDir As String
Private searchPath() As String          ' the folders as mounted
Private searchMount() As String         ' their mount points ("" = the root)
Private searchCount As Long
Private files() As PHYSFS_FileHandle
Private filesReady As Boolean
Private lastError As String
Private permitSymLinks As Boolean

'--------------------------------------------------------------------------
' Helpers

Private Function IsDirectory(ByVal path As String) As Boolean
    Dim a As Long
    a = GetFileAttributesW(StrPtr(path))
    If a = INVALID_FILE_ATTRIBUTES Then Exit Function
    IsDirectory = (a And FILE_ATTRIBUTE_DIRECTORY) <> 0
End Function

' Creates a folder and its missing parents.
Private Function MakePath(ByVal path As String) As Boolean
    Dim p As Long
    If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
    If IsDirectory(path) Then
        MakePath = True
        Exit Function
    End If
    p = InStrRev(path, "\")
    If p > 3 Then
        If Not MakePath(Left$(path, p - 1)) Then Exit Function
    End If
    If CreateDirectoryW(StrPtr(path), 0) = 0 Then
        MakePath = IsDirectory(path)
    Else
        MakePath = True
    End If
End Function

Private Function WithSeparator(ByVal dirPath As String) As String
    If Right$(dirPath, 1) = "\" Then WithSeparator = dirPath Else WithSeparator = dirPath & "\"
End Function

' PhysFS's sanitizePlatformIndependentPath: leading and repeated "/" go, "." and ".." elements,
' ":" and "\" are refused. False with lastError set on a bad name.
Private Function Sanitize(ByVal fname As String, out As String) As Boolean
    Dim parts() As String, i As Long, result As String
    If InStr(fname, ":") > 0 Or InStr(fname, "\") > 0 Then
        lastError = ERR_BAD_FILENAME
        Exit Function
    End If
    parts = Split(fname, "/")
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then
            If parts(i) = "." Or parts(i) = ".." Then
                lastError = ERR_BAD_FILENAME
                Exit Function
            End If
            If Len(result) > 0 Then result = result & "/"
            result = result & parts(i)
        End If
    Next
    out = result
    Sanitize = True
End Function

' The folder on disk for name in search path entry i, or "" if the entry's mount point does not cover it.
Private Function EntryPath(ByVal i As Long, ByVal nm As String) As String
    Dim mp As String, rest As String
    mp = searchMount(i)
    If Len(mp) = 0 Then
        rest = nm
    ElseIf nm = mp Then
        rest = ""
    ElseIf Left$(nm, Len(mp) + 1) = mp & "/" Then
        rest = Mid$(nm, Len(mp) + 2)
    Else
        Exit Function
    End If
    EntryPath = WithSeparator(searchPath(i)) & Replace(rest, "/", "\")
End Function

' The first search path entry that has name (file or folder), or -1.
Private Function FindEntry(ByVal nm As String, path As String) As Long
    Dim i As Long, p As String
    For i = 0 To searchCount - 1
        p = EntryPath(i, nm)
        If Len(p) > 0 Then
            If GetFileAttributesW(StrPtr(p)) <> INVALID_FILE_ATTRIBUTES Then
                path = p
                FindEntry = i
                Exit Function
            End If
        End If
    Next
    FindEntry = -1
End Function

Private Function NewHandle(ByVal h As Long, ByVal forReading As Boolean) As Long
    Dim i As Long
    If Not filesReady Then
        ReDim files(1 To 16)
        filesReady = True
    End If
    For i = 1 To UBound(files)
        If Not files(i).inUse Then Exit For
    Next
    If i > UBound(files) Then ReDim Preserve files(1 To i + 15)
    files(i).handle = h
    files(i).forReading = forReading
    files(i).inUse = True
    NewHandle = i
End Function

Private Function ValidHandle(ByVal f As Long) As Boolean
    If Not filesReady Then Exit Function
    If f < 1 Or f > UBound(files) Then Exit Function
    ValidHandle = files(f).inUse
End Function

Private Function Unsigned(ByVal l As Long) As Double
    If l < 0 Then Unsigned = l + 4294967296# Else Unsigned = l
End Function

Private Function FileTimeToUnix(ft As FILETIME) As Double
    FileTimeToUnix = Int((Unsigned(ft.dwHighDateTime) * 4294967296# + Unsigned(ft.dwLowDateTime)) / 10000000#) - 11644473600#
End Function

'--------------------------------------------------------------------------
' Initialisation, the write directory, the search path

Public Function PHYSFS_init(ByVal argv0 As String) As Long
    Dim buf As String, n As Long
    buf = String$(1024, 0)
    n = GetModuleFileNameW(0, StrPtr(buf), 1024)
    buf = Left$(buf, n)
    baseDir = Left$(buf, InStrRev(buf, "\"))
    searchCount = 0
    initialised = True
    PHYSFS_init = 1
End Function

' Closes the open files and forgets the search path and the write dir.
Public Function PHYSFS_deinit() As Long
    Dim i As Long
    If Not initialised Then
        lastError = ERR_NOT_INITIALIZED
        Exit Function
    End If
    If filesReady Then
        For i = 1 To UBound(files)
            If files(i).inUse Then
                CloseHandle files(i).handle
                files(i).inUse = False
                files(i).handle = 0
            End If
        Next
    End If
    searchCount = 0
    writeDir = vbNullString
    baseDir = vbNullString
    initialised = False
    PHYSFS_deinit = 1
End Function

Public Function PHYSFS_setWriteDir(ByVal newDir As String) As Long
    If Not initialised Then
        lastError = ERR_NOT_INITIALIZED
        Exit Function
    End If
    If StrPtr(newDir) = 0 Then
        writeDir = vbNullString
        PHYSFS_setWriteDir = 1
        Exit Function
    End If
    If Not IsDirectory(newDir) Then
        lastError = ERR_NOT_FOUND
        Exit Function
    End If
    writeDir = newDir
    PHYSFS_setWriteDir = 1
End Function

' vbNullString when no write directory is set (C: NULL)
Public Function PHYSFS_getWriteDir() As String
    PHYSFS_getWriteDir = writeDir
End Function

' Creates dirName and its parents in the write directory.
Public Function PHYSFS_mkdir(ByVal dirName As String) As Long
    Dim nm As String
    If Len(writeDir) = 0 Then
        lastError = ERR_NO_WRITE_DIR
        Exit Function
    End If
    If Not Sanitize(dirName, nm) Then Exit Function
    If MakePath(WithSeparator(writeDir) & Replace(nm, "/", "\")) Then
        PHYSFS_mkdir = 1
    Else
        lastError = ERR_IO
    End If
End Function

' Deletes a file or an empty folder in the write directory.
Public Function PHYSFS_delete(ByVal fname As String) As Long
    Dim nm As String, path As String
    If Len(writeDir) = 0 Then
        lastError = ERR_NO_WRITE_DIR
        Exit Function
    End If
    If Not Sanitize(fname, nm) Then Exit Function
    path = WithSeparator(writeDir) & Replace(nm, "/", "\")
    If IsDirectory(path) Then
        If RemoveDirectoryW(StrPtr(path)) <> 0 Then PHYSFS_delete = 1 Else lastError = ERR_IO
    ElseIf DeleteFileW(StrPtr(path)) <> 0 Then
        PHYSFS_delete = 1
    ElseIf GetFileAttributesW(StrPtr(path)) = INVALID_FILE_ATTRIBUTES Then
        lastError = ERR_NOT_FOUND
    Else
        lastError = ERR_IO
    End If
End Function

' appendToPath: PHYSFS_APPEND (1) or PHYSFS_PREPEND (0). mountPoint: vbNullString or "" = root.
Public Function PHYSFS_mount(ByVal newDir As String, ByVal mountPoint As String, ByVal appendToPath As Long) As Long
    Dim i As Long, mp As String
    If Not initialised Then
        lastError = ERR_NOT_INITIALIZED
        Exit Function
    End If
    If Not IsDirectory(newDir) Then
        If GetFileAttributesW(StrPtr(newDir)) <> INVALID_FILE_ATTRIBUTES Then
            lastError = ERR_UNSUPPORTED             ' a file: an archive (.wz, zip) - not yet
        Else
            lastError = ERR_NOT_FOUND
        End If
        Exit Function
    End If
    If Not Sanitize(mountPoint & "", mp) Then Exit Function
    For i = 0 To searchCount - 1
        If StrComp(searchPath(i), newDir, vbTextCompare) = 0 Then
            PHYSFS_mount = 1                    ' already mounted: PhysFS does nothing
            Exit Function
        End If
    Next
    ReDim Preserve searchPath(0 To searchCount)
    ReDim Preserve searchMount(0 To searchCount)
    If appendToPath <> 0 Then
        searchPath(searchCount) = newDir
        searchMount(searchCount) = mp
    Else
        For i = searchCount To 1 Step -1
            searchPath(i) = searchPath(i - 1)
            searchMount(i) = searchMount(i - 1)
        Next
        searchPath(0) = newDir
        searchMount(0) = mp
    End If
    searchCount = searchCount + 1
    PHYSFS_mount = 1
End Function

Public Function PHYSFS_unmount(ByVal oldDir As String) As Long
    Dim i As Long, j As Long
    For i = 0 To searchCount - 1
        If StrComp(searchPath(i), oldDir, vbTextCompare) = 0 Then
            For j = i To searchCount - 2
                searchPath(j) = searchPath(j + 1)
                searchMount(j) = searchMount(j + 1)
            Next
            searchCount = searchCount - 1
            PHYSFS_unmount = 1
            Exit Function
        End If
    Next
    lastError = ERR_NOT_MOUNTED
End Function

' The search path in order; returns the count.
Public Function PHYSFS_getSearchPath(out() As String) As Long
    Dim i As Long
    If searchCount = 0 Then
        Erase out
        Exit Function
    End If
    ReDim out(0 To searchCount - 1)
    For i = 0 To searchCount - 1
        out(i) = searchPath(i)
    Next
    PHYSFS_getSearchPath = searchCount
End Function

' The program's folder, with a trailing separator.
Public Function PHYSFS_getBaseDir() As String
    PHYSFS_getBaseDir = baseDir
End Function

Public Function PHYSFS_getDirSeparator() As String
    PHYSFS_getDirSeparator = "\"
End Function

Public Sub PHYSFS_permitSymbolicLinks(ByVal allow As Long)
    permitSymLinks = (allow <> 0)
End Sub

' %APPDATA%\org\app\, created if missing, with a trailing separator (as PhysFS does on Windows).
Public Function PHYSFS_getPrefDir(ByVal org As String, ByVal app As String) As String
    Dim buf As String, path As String
    buf = String$(260, 0)
    If SHGetFolderPathW(0, CSIDL_APPDATA, 0, 0, StrPtr(buf)) <> 0 Then
        lastError = ERR_NOT_FOUND
        PHYSFS_getPrefDir = vbNullString
        Exit Function
    End If
    path = Left$(buf, InStr(buf, vbNullChar) - 1) & "\" & org & "\" & app & "\"
    If Not MakePath(path) Then
        lastError = ERR_IO
        PHYSFS_getPrefDir = vbNullString
        Exit Function
    End If
    PHYSFS_getPrefDir = path
End Function

'--------------------------------------------------------------------------
' Files

Public Function PHYSFS_openRead(ByVal fname As String) As Long
    Dim nm As String, path As String, h As Long
    If Not Sanitize(fname, nm) Then Exit Function
    If FindEntry(nm, path) < 0 Then
        lastError = ERR_NOT_FOUND
        Exit Function
    End If
    If IsDirectory(path) Then
        lastError = ERR_NOT_FOUND
        Exit Function
    End If
    h = CreateFileW(StrPtr(path), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If h = INVALID_HANDLE_VALUE Then
        lastError = ERR_IO
        Exit Function
    End If
    PHYSFS_openRead = NewHandle(h, True)
End Function

Private Function OpenInWriteDir(ByVal fname As String, ByVal disposition As Long, ByVal append As Boolean) As Long
    Dim nm As String, path As String, h As Long, hi As Long
    If Len(writeDir) = 0 Then
        lastError = ERR_NO_WRITE_DIR
        Exit Function
    End If
    If Not Sanitize(fname, nm) Then Exit Function
    path = WithSeparator(writeDir) & Replace(nm, "/", "\")
    h = CreateFileW(StrPtr(path), GENERIC_WRITE, FILE_SHARE_READ, 0, disposition, FILE_ATTRIBUTE_NORMAL, 0)
    If h = INVALID_HANDLE_VALUE Then
        lastError = ERR_IO
        Exit Function
    End If
    If append Then SetFilePointer h, 0, hi, FILE_END
    OpenInWriteDir = NewHandle(h, False)
End Function

Public Function PHYSFS_openWrite(ByVal fname As String) As Long
    PHYSFS_openWrite = OpenInWriteDir(fname, CREATE_ALWAYS, False)
End Function

Public Function PHYSFS_openAppend(ByVal fname As String) As Long
    PHYSFS_openAppend = OpenInWriteDir(fname, OPEN_ALWAYS, True)
End Function

' Reads length bytes into buffer(ofs ..). The count read, or -1.
Public Function PHYSFS_readBytes(ByVal f As Long, buffer() As Byte, ByVal ofs As Long, ByVal length As Long) As Long
    Dim got As Long
    If Not ValidHandle(f) Then
        lastError = ERR_INVALID_ARGUMENT
        PHYSFS_readBytes = -1
        Exit Function
    End If
    If Not files(f).forReading Then
        lastError = ERR_OPEN_FOR_WRITING
        PHYSFS_readBytes = -1
        Exit Function
    End If
    If length = 0 Then Exit Function
    If ReadFile(files(f).handle, VarPtr(buffer(ofs)), length, got, 0) = 0 Then
        lastError = ERR_IO
        PHYSFS_readBytes = -1
        Exit Function
    End If
    If got < length Then lastError = ERR_PAST_EOF
    PHYSFS_readBytes = got
End Function

' Writes length bytes from buffer(ofs ..). The count written, or -1.
Public Function PHYSFS_writeBytes(ByVal f As Long, buffer() As Byte, ByVal ofs As Long, ByVal length As Long) As Long
    Dim done As Long
    If Not ValidHandle(f) Then
        lastError = ERR_INVALID_ARGUMENT
        PHYSFS_writeBytes = -1
        Exit Function
    End If
    If files(f).forReading Then
        lastError = ERR_OPEN_FOR_READING
        PHYSFS_writeBytes = -1
        Exit Function
    End If
    If length = 0 Then Exit Function
    If WriteFile(files(f).handle, VarPtr(buffer(ofs)), length, done, 0) = 0 Then
        lastError = ERR_IO
        PHYSFS_writeBytes = -1
        Exit Function
    End If
    PHYSFS_writeBytes = done
End Function

' The file's length in bytes, or -1.
Public Function PHYSFS_fileLength(ByVal f As Long) As Double
    Dim lo As Long, hi As Long
    If Not ValidHandle(f) Then
        lastError = ERR_INVALID_ARGUMENT
        PHYSFS_fileLength = -1
        Exit Function
    End If
    lo = GetFileSize(files(f).handle, hi)
    PHYSFS_fileLength = Unsigned(hi) * 4294967296# + Unsigned(lo)
End Function

Public Function PHYSFS_tell(ByVal f As Long) As Double
    Dim lo As Long, hi As Long
    If Not ValidHandle(f) Then
        lastError = ERR_INVALID_ARGUMENT
        PHYSFS_tell = -1
        Exit Function
    End If
    hi = 0
    lo = SetFilePointer(files(f).handle, 0, hi, 1)          ' FILE_CURRENT
    PHYSFS_tell = Unsigned(hi) * 4294967296# + Unsigned(lo)
End Function

Public Function PHYSFS_seek(ByVal f As Long, ByVal pos As Double) As Long
    Dim lo As Double, hi As Long, l As Long
    If Not ValidHandle(f) Then
        lastError = ERR_INVALID_ARGUMENT
        Exit Function
    End If
    hi = Int(pos / 4294967296#)
    lo = pos - CDbl(hi) * 4294967296#
    If lo >= 2147483648# Then l = lo - 4294967296# Else l = lo
    SetFilePointer files(f).handle, l, hi, FILE_BEGIN
    PHYSFS_seek = 1
End Function

Public Function PHYSFS_eof(ByVal f As Long) As Long
    If Not ValidHandle(f) Then Exit Function
    If PHYSFS_tell(f) >= PHYSFS_fileLength(f) Then PHYSFS_eof = 1
End Function

Public Function PHYSFS_flush(ByVal f As Long) As Long
    If Not ValidHandle(f) Then Exit Function
    If Not files(f).forReading Then FlushFileBuffers files(f).handle
    PHYSFS_flush = 1
End Function

Public Function PHYSFS_close(ByVal f As Long) As Long
    If Not ValidHandle(f) Then
        lastError = ERR_INVALID_ARGUMENT
        Exit Function
    End If
    CloseHandle files(f).handle
    files(f).inUse = False
    files(f).handle = 0
    PHYSFS_close = 1
End Function

'--------------------------------------------------------------------------
' Byte order helpers (physfs.h): 32-bit values in big-endian order; nonzero on
' success. A read that gets fewer than 4 bytes fails and leaves value as it was.

Public Function PHYSFS_writeSBE32(ByVal f As Long, ByVal value As Long) As Long
    Dim b(0 To 3) As Byte
    b(0) = (value And &H7F000000) \ &H1000000
    If value < 0 Then b(0) = b(0) Or &H80
    b(1) = (value And &HFF0000) \ &H10000
    b(2) = (value And &HFF00&) \ &H100&
    b(3) = value And &HFF&
    If PHYSFS_writeBytes(f, b, 0, 4) = 4 Then PHYSFS_writeSBE32 = 1
End Function

Public Function PHYSFS_writeUBE32(ByVal f As Long, ByVal value As Long) As Long
    PHYSFS_writeUBE32 = PHYSFS_writeSBE32(f, value)
End Function

Public Function PHYSFS_readSBE32(ByVal f As Long, value As Long) As Long
    Dim b(0 To 3) As Byte
    If PHYSFS_readBytes(f, b, 0, 4) <> 4 Then Exit Function
    value = (CLng(b(0) And &H7F) * &H1000000) Or (CLng(b(1)) * &H10000) Or (CLng(b(2)) * &H100&) Or b(3)
    If (b(0) And &H80) <> 0 Then value = value Or &H80000000
    PHYSFS_readSBE32 = 1
End Function

Public Function PHYSFS_readUBE32(ByVal f As Long, value As Long) As Long
    PHYSFS_readUBE32 = PHYSFS_readSBE32(f, value)
End Function

'--------------------------------------------------------------------------
' Queries

Public Function PHYSFS_exists(ByVal fname As String) As Long
    Dim nm As String, path As String
    If Not Sanitize(fname, nm) Then Exit Function
    If Len(nm) = 0 Then
        PHYSFS_exists = 1                     ' the root
    ElseIf FindEntry(nm, path) >= 0 Then
        PHYSFS_exists = 1
    End If
End Function

' PHYSFS_stat: 1 and the fields, or 0 if fname is in no search path entry. Times are seconds since 1970.
Public Function PHYSFS_stat(ByVal fname As String, filesize As Double, modtime As Double, filetype As Long, readonly_ As Long) As Long
    Dim nm As String, path As String, d As WIN32_FILE_ATTRIBUTE_DATA
    ' set some sane defaults... (PhysFS does, so a failed stat leaves them)
    filesize = -1
    modtime = -1
    filetype = 3                            ' PHYSFS_FILETYPE_OTHER
    readonly_ = 1
    If Not Sanitize(fname, nm) Then Exit Function
    If Len(nm) = 0 Then
        filesize = 0
        modtime = 0
        filetype = PHYSFS_FILETYPE_DIRECTORY
        readonly_ = 1
        PHYSFS_stat = 1
        Exit Function
    End If
    If FindEntry(nm, path) < 0 Then
        lastError = ERR_NOT_FOUND
        Exit Function
    End If
    If GetFileAttributesExW(StrPtr(path), 0, d) = 0 Then
        lastError = ERR_IO
        Exit Function
    End If
    If (d.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
        filetype = PHYSFS_FILETYPE_DIRECTORY
        filesize = 0
    ElseIf (d.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT) <> 0 Then
        filetype = PHYSFS_FILETYPE_SYMLINK
        filesize = Unsigned(d.nFileSizeHigh) * 4294967296# + Unsigned(d.nFileSizeLow)
    Else
        filetype = PHYSFS_FILETYPE_REGULAR
        filesize = Unsigned(d.nFileSizeHigh) * 4294967296# + Unsigned(d.nFileSizeLow)
    End If
    modtime = FileTimeToUnix(d.ftLastWriteTime)
    If (d.dwFileAttributes And FILE_ATTRIBUTE_READONLY) <> 0 Then readonly_ = 1 Else readonly_ = 0
    PHYSFS_stat = 1
End Function

' physfs_ext.h: WZ_PHYSFS_isDirectory
Public Function WZ_PHYSFS_isDirectory(ByVal fname As String) As Long
    Dim filesize As Double, modtime As Double, filetype As Long, ro As Long
    If PHYSFS_stat(fname, filesize, modtime, filetype, ro) = 0 Then Exit Function
    If filetype = PHYSFS_FILETYPE_DIRECTORY Then WZ_PHYSFS_isDirectory = 1
End Function

' The search path entry (as mounted) that has fname; vbNullString if none (C: NULL).
Public Function PHYSFS_getRealDir(ByVal fname As String) As String
    Dim nm As String, path As String, i As Long
    PHYSFS_getRealDir = vbNullString
    If Not Sanitize(fname, nm) Then Exit Function
    i = FindEntry(nm, path)
    If i < 0 Then
        lastError = ERR_NOT_FOUND
        Exit Function
    End If
    PHYSFS_getRealDir = searchPath(i)
End Function

' Inserts s into the sorted list (strcmp order), unless it is there already.
Private Sub AddSorted(list() As String, n As Long, ByVal s As String)
    Dim lo As Long, hi As Long, m As Long, c As Long, i As Long
    lo = 0
    hi = n - 1
    Do While lo <= hi
        m = (lo + hi) \ 2
        c = StrComp(list(m), s, vbBinaryCompare)
        If c = 0 Then Exit Sub
        If c < 0 Then lo = m + 1 Else hi = m - 1
    Loop
    If n > UBound(list) Then ReDim Preserve list(0 To n * 2)
    For i = n To lo + 1 Step -1
        list(i) = list(i - 1)
    Next
    list(lo) = s
    n = n + 1
End Sub

' The names in folder dir, over the whole search path; returns the count.
Public Function PHYSFS_enumerateFiles(ByVal dirPath As String, out() As String) As Long
    Dim nm As String, i As Long, path As String, h As Long, fd As WIN32_FIND_DATAW
    Dim n As Long, entry As String, k As Long, list() As String, mp As String
    ReDim list(0 To 15)
    If Not Sanitize(dirPath, nm) Then
        Erase out
        Exit Function
    End If
    For i = 0 To searchCount - 1
        path = EntryPath(i, nm)
        If Len(path) > 0 Then
            h = FindFirstFileW(StrPtr(WithSeparator(path) & "*"), fd)
            If h <> INVALID_HANDLE_VALUE Then
                Do
                    entry = fd.cFileName
                    k = InStr(entry, vbNullChar)
                    If k > 0 Then entry = Left$(entry, k - 1)
                    If entry <> "." And entry <> ".." Then AddSorted list, n, entry
                Loop While FindNextFileW(h, fd) <> 0
                FindClose h
            End If
        Else
            ' a mount point inside dir shows up as a folder
            mp = searchMount(i)
            If Len(nm) = 0 Then
                k = InStr(mp, "/")
                If k > 0 Then AddSorted list, n, Left$(mp, k - 1) Else AddSorted list, n, mp
            ElseIf Left$(mp, Len(nm) + 1) = nm & "/" Then
                entry = Mid$(mp, Len(nm) + 2)
                k = InStr(entry, "/")
                If k > 0 Then entry = Left$(entry, k - 1)
                AddSorted list, n, entry
            End If
        End If
    Next
    If n = 0 Then
        Erase out
        Exit Function
    End If
    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        out(i) = list(i)
    Next
    PHYSFS_enumerateFiles = n
End Function

' physfs_ext.h: the message of PhysFS's last error.
' PhysFS 3's PHYSFS_ErrorCode values (their order) and messages
Public Function PHYSFS_getErrorByCode(ByVal code As Long) As String
    Select Case code
    Case 0: PHYSFS_getErrorByCode = "no error"
    Case 1: PHYSFS_getErrorByCode = "unknown error"
    Case 2: PHYSFS_getErrorByCode = "out of memory"
    Case 3: PHYSFS_getErrorByCode = ERR_NOT_INITIALIZED
    Case 4: PHYSFS_getErrorByCode = "already initialized"
    Case 5: PHYSFS_getErrorByCode = "argv[0] is NULL"
    Case 6: PHYSFS_getErrorByCode = ERR_UNSUPPORTED
    Case 7: PHYSFS_getErrorByCode = ERR_PAST_EOF
    Case 8: PHYSFS_getErrorByCode = "files still open"
    Case 9: PHYSFS_getErrorByCode = ERR_INVALID_ARGUMENT
    Case 10: PHYSFS_getErrorByCode = ERR_NOT_MOUNTED
    Case 11: PHYSFS_getErrorByCode = ERR_NOT_FOUND
    Case 12: PHYSFS_getErrorByCode = "symlinks are forbidden"
    Case 13: PHYSFS_getErrorByCode = ERR_NO_WRITE_DIR
    Case 14: PHYSFS_getErrorByCode = ERR_OPEN_FOR_READING
    Case 15: PHYSFS_getErrorByCode = ERR_OPEN_FOR_WRITING
    Case 16: PHYSFS_getErrorByCode = "not a file"
    Case 17: PHYSFS_getErrorByCode = "read-only filesystem"
    Case 18: PHYSFS_getErrorByCode = "corrupted"
    Case 19: PHYSFS_getErrorByCode = "infinite symbolic link loop"
    Case 20: PHYSFS_getErrorByCode = ERR_IO
    Case 21: PHYSFS_getErrorByCode = "permission denied"
    Case 22: PHYSFS_getErrorByCode = "no space available for writing"
    Case 23: PHYSFS_getErrorByCode = ERR_BAD_FILENAME
    Case 24: PHYSFS_getErrorByCode = "tried to modify a file the OS needs"
    Case 25: PHYSFS_getErrorByCode = "directory isn't empty"
    Case 26: PHYSFS_getErrorByCode = "OS reported an error"
    Case 27: PHYSFS_getErrorByCode = "duplicate resource"
    Case 28: PHYSFS_getErrorByCode = "bad password"
    Case 29: PHYSFS_getErrorByCode = "app callback reported error"
    Case Else: PHYSFS_getErrorByCode = vbNullString
    End Select
End Function

' The code of the last error, which is then cleared (as PhysFS does)
Public Function PHYSFS_getLastErrorCode() As Long
    Dim code As Long
    If Len(lastError) > 0 Then
        For code = 1 To 29
            If PHYSFS_getErrorByCode(code) = lastError Then Exit For
        Next
        If code > 29 Then code = 1                  ' PHYSFS_ERR_OTHER_ERROR
    End If
    lastError = ""
    PHYSFS_getLastErrorCode = code
End Function

Public Function WZ_PHYSFS_getLastError() As String
    WZ_PHYSFS_getLastError = lastError
End Function
