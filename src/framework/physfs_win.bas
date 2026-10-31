Attribute VB_Name = "physfs_win"
' PhysicsFS 3.2.0 for the port, on Win32: the part of physfs.c Warzone uses, with its archiver dispatch (a
' search path entry is a DirHandle: a directory opened by the DIR archiver, physfs_archiver_dir.c, or an
' archive opened by an archiver; mounting a file tries the archivers, of which the ZIP one is ported:
' physfs_zip_c), the PHYSFS_Io objects (the native file io and the archivers' ios), the Windows platform
' layer they call (physfs_platform_windows.c's __PHYSFS_platform* functions: PHYSFS_platformStat__ ...), and
' WZ_PHYSFS_getLastError and WZ_PHYSFS_isDirectory (physfs_ext.h). Ported from physfs.c: the error state
' (PHYSFS_setErrorCode, PHYSFS_getLastErrorCode, PHYSFS_getErrorByCode), the native io, openDirectory and
' tryOpenDir, sanitizePlatformIndependentPath, partOfMountPoint, createDirHandle, freeDirHandle, PHYSFS_init,
' PHYSFS_deinit (doDeinit, closeFileHandleList, freeSearchPath), PHYSFS_getPrefDir, PHYSFS_getBaseDir,
' PHYSFS_getWriteDir, PHYSFS_setWriteDir, doMount, PHYSFS_mount, PHYSFS_unmount, PHYSFS_getSearchPath,
' PHYSFS_permitSymbolicLinks, verifyPath, doMkdir, PHYSFS_mkdir, doDelete, PHYSFS_delete, getRealDirHandle,
' PHYSFS_getRealDir, locateInStringList, enumFilesCallback, PHYSFS_enumerateFiles, enumerateFromMountPoint,
' enumCallbackFilterSymLinks, PHYSFS_enumerate, PHYSFS_exists, doOpenWrite, PHYSFS_openWrite,
' PHYSFS_openAppend, PHYSFS_openRead, closeHandleInOpenList, PHYSFS_close, PHYSFS_readBytes,
' PHYSFS_writeBytes, PHYSFS_eof, PHYSFS_tell, PHYSFS_seek, PHYSFS_fileLength, PHYSFS_flush, PHYSFS_stat,
' __PHYSFS_readAll; physfs_byteorder.c's 32-bit big-endian helpers. Not ported (Warzone does not call them):
' the other archivers (7z, grp, ...: mounting such a file gives "unsupported", as a file no archiver claims),
' the memory and handle ios (PHYSFS_mountMemory, PHYSFS_mountIo, PHYSFS_mountHandle), PHYSFS_setRoot (no
' DirHandle has a root), PHYSFS_setBuffer (no FileHandle has a buffer: WZ_PHYSFS_SETBUFFER is nothing, see
' physfs_ext_c), the allocator, the mutexes (one thread), the CD-ROM and user dirs.
' (VB) A PHYSFS_File * is an index into files() (0 = NULL); a PHYSFS_Io * an index into ios() (0 = NULL),
' whose kind is its function table (io->read(io, ...) is PHYSFS_Io_read(io, ...)); a DirHandle's funcs is
' its archiver's number and its opaque the archiver's handle (DIR: dirOpaques(), its folder with a separator
' at its end). The search path (the C's linked list) is searchPath(0 .. searchCount - 1) in order; a
' FileHandle's dirHandle is its DirHandle's id. Names are VB Strings where PhysFS has UTF-8 (strcmp is a
' binary StrComp); the enumeration callbacks are numbers (PHYSFS_EnumerateCallback_call), their data the
' module's variables. File times are seconds since 1970 (FileTimeToPhysfsTime's result: its round trip
' through the local time is left out, so a time in the hour a DST change repeats can differ by 3600 s).
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
Private Declare Function GetFileAttributesExW Lib "kernel32" (ByVal lpFileName As Long, ByVal fInfoLevelId As Long, lpFileInformation As WIN32_FILE_ATTRIBUTE_DATA) As Long
Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFilename As Long, ByVal nSize As Long) As Long
Private Declare Function SHGetFolderPathW Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
' (a LARGE_INTEGER by value: its low and high halves)
Private Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, ByVal liDistanceToMoveLow As Long, ByVal liDistanceToMoveHigh As Long, ByVal lpNewFilePointer As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, ByVal lpFileSize As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATAW) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATAW) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal src As Long, ByVal n As Long)

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_DEVICE As Long = &H40
Private Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400          ' PHYSFS_FILE_ATTRIBUTE_REPARSE_POINT
Private Const FILE_ATTRIBUTE_OFFLINE As Long = &H1000
Private Const IO_REPARSE_TAG_SYMLINK As Long = &HA000000C           ' PHYSFS_IO_REPARSE_TAG_SYMLINK
Private Const CSIDL_APPDATA As Long = &H1A
Private Const CSIDL_FLAG_CREATE As Long = &H8000&
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const OPEN_ALWAYS As Long = 4
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_BEGIN As Long = 0
Private Const FILE_CURRENT As Long = 1
Private Const FILE_END As Long = 2

' Win32 error codes (errcodeFromWinApiError)
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_NOT_ENOUGH_MEMORY As Long = 8
Private Const ERROR_OUTOFMEMORY As Long = 14
Private Const ERROR_INVALID_DRIVE As Long = 15
Private Const ERROR_CURRENT_DIRECTORY As Long = 16
Private Const ERROR_WRITE_PROTECT As Long = 19
Private Const ERROR_NOT_READY As Long = 21
Private Const ERROR_CRC As Long = 23
Private Const ERROR_SEEK As Long = 25
Private Const ERROR_NOT_DOS_DISK As Long = 26
Private Const ERROR_SECTOR_NOT_FOUND As Long = 27
Private Const ERROR_WRITE_FAULT As Long = 29
Private Const ERROR_READ_FAULT As Long = 30
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Const ERROR_LOCK_VIOLATION As Long = 33
Private Const ERROR_HANDLE_DISK_FULL As Long = 39
Private Const ERROR_DEV_NOT_EXIST As Long = 55
Private Const ERROR_NETWORK_ACCESS_DENIED As Long = 65
Private Const ERROR_DRIVE_LOCKED As Long = 108
Private Const ERROR_BUFFER_OVERFLOW As Long = 111
Private Const ERROR_DISK_FULL As Long = 112
Private Const ERROR_INVALID_NAME As Long = 123
Private Const ERROR_DIR_NOT_EMPTY As Long = 145
Private Const ERROR_PATH_BUSY As Long = 148
Private Const ERROR_BAD_PATHNAME As Long = 161
Private Const ERROR_BUSY As Long = 170
Private Const ERROR_DIRECTORY As Long = 267
Private Const ERROR_DELETE_PENDING As Long = 303

' physfs.h: struct PHYSFS_Stat ((VB) PHYSFS_Stat_: the function PHYSFS_stat keeps the name)
Public Type PHYSFS_Stat_
    filesize As Double                  ' size in bytes, -1 for non-files and unknown
    modtime As Double                   ' last modification time
    createtime As Double                ' like modtime, but for file creation time
    accesstime As Double                ' like modtime, but for file access time
    filetype As Long                    ' File? Directory? Symlink? (PHYSFS_FileType)
    readonly_ As Long                   ' non-zero if read only, zero if writable.
End Type

' physfs.h: PHYSFS_EnumerateCallbackResult
Public Const PHYSFS_ENUM_ERROR As Long = -1     ' Stop enumerating, report error to app.
Public Const PHYSFS_ENUM_STOP As Long = 0       ' Stop enumerating, report success to app.
Public Const PHYSFS_ENUM_OK As Long = 1         ' Keep enumerating, no problems

' (VB) the PHYSFS_Io function tables: __PHYSFS_nativeIoInterface and physfs_zip_c's ZIP_Io
Public Const PHYSFS_IO_NATIVE As Long = 1
Public Const PHYSFS_IO_ZIP As Long = 2

' (VB) the archivers (const PHYSFS_Archiver *): __PHYSFS_Archiver_DIR and the static ones (initStaticArchivers)
Private Const ARCHIVER_DIR As Long = 1
Private Const ARCHIVER_ZIP As Long = 2
Private Const numArchivers As Long = 1          ' archivers[]: ZIP

' (VB) the enumeration callbacks (PHYSFS_EnumerateCallback)
Private Const ENUMCB_FILES As Long = 1          ' enumFilesCallback
Private Const ENUMCB_FILTER_SYMLINKS As Long = 2 ' enumCallbackFilterSymLinks

Private Const MODE_R As Long = 114               ' 'r'
Private Const MODE_W As Long = 119              ' 'w'
Private Const MODE_A As Long = 97               ' 'a'

' PHYSFS_Io (physfs.h) with the native io's NativeIoInfo
Private Type PHYSFS_Io
    kind As Long                        ' (VB) its function table: PHYSFS_IO_NATIVE, PHYSFS_IO_ZIP (0: free)
    opaque As Long                      ' Instance data for this struct: (VB) the archiver's io's (ZIP: a ZIPfileinfo)
    handle As Long                      ' NativeIoInfo: void *handle
    path As String                      ' const char *path
    mode As Long                        ' 'r', 'w', or 'a'
End Type

Private Type DirHandle
    opaque As Long                      ' Instance data unique to the archiver.
    dirName As String                   ' Path to archive in platform-dependent notation.
    mountPoint As String                ' Mountpoint in virtual file tree. (vbNullString: NULL)
    funcs As Long                       ' Ptr to archiver info for this handle. (VB: its number)
    id As Long                          ' (VB) what a FileHandle's dirHandle names
End Type

Private Type FileHandle
    io As Long                          ' Instance data unique to the archiver for this file.
    forReading As Boolean               ' Non-zero if reading, zero if write/append
    dirHandle As Long                   ' Archiver instance that created this ((VB) its id)
    inUse As Boolean                    ' (VB) in openReadList or openWriteList
End Type

' General PhysicsFS state ...
Private initialized As Boolean
Private errorCode As Long                       ' ErrState's code (one thread)
Private searchPath() As DirHandle               ' the search path, in order (the C's list)
Private searchCount As Long
Private writeDir As DirHandle
Private hasWriteDir As Boolean                  ' writeDir != NULL
Private files() As FileHandle                   ' openWriteList and openReadList
Private filesCount As Long
Private baseDir As String
Private prefDir As String
Private allowSymLinks As Boolean
Private nextDirHandleId As Long                 ' (VB)

Private ios() As PHYSFS_Io                      ' (VB) the PHYSFS_Io allocations
Private iosCount As Long
Private dirOpaques() As String                  ' (VB) the DIR archiver's instances
Private dirOpaqueUsed() As Boolean
Private dirOpaquesCount As Long

' EnumStringListCallbackData (enumFilesCallback's data)
Private enumList() As String
Private enumListSize As Long
Private enumErrcode As Long
' SymlinkFilterData (enumCallbackFilterSymLinks's data)
Private filterCallback As Long
Private filterFuncs As Long                     ' the dirhandle's funcs and opaque
Private filterOpaque As Long
Private filterArcfname As String
Private filterErrcode As Long

'--------------------------------------------------------------------------
' (VB) helpers

Private Function Unsigned(ByVal l As Long) As Double
    If l < 0 Then Unsigned = l + 4294967296# Else Unsigned = l
End Function

' FileTimeToPhysfsTime: seconds since 1970 (see the header)
Private Function FileTimeToPhysfsTime(ft As FILETIME) As Double
    FileTimeToPhysfsTime = Int((Unsigned(ft.dwHighDateTime) * 4294967296# + Unsigned(ft.dwLowDateTime)) / 10000000#) - 11644473600#
End Function

Private Function newFileHandle() As Long
    Dim i As Long
    For i = 1 To filesCount
        If Not files(i).inUse Then Exit For
    Next
    If i > filesCount Then
        filesCount = filesCount + 16
        ReDim Preserve files(1 To filesCount)
    End If
    files(i).inUse = True
    newFileHandle = i
End Function

Private Function ValidHandle(ByVal f As Long) As Boolean
    If f < 1 Or f > filesCount Then Exit Function
    ValidHandle = files(f).inUse
End Function

'--------------------------------------------------------------------------
' Errors

Public Sub PHYSFS_setErrorCode(ByVal errcode As Long)
    If errcode = 0 Then
        Exit Sub
    End If
    errorCode = errcode
End Sub ' PHYSFS_setErrorCode

Private Function currentErrorCode() As Long
    currentErrorCode = errorCode
End Function ' currentErrorCode

' The code of the last error, which is then cleared
Public Function PHYSFS_getLastErrorCode() As Long
    PHYSFS_getLastErrorCode = errorCode
    errorCode = PHYSFS_ERR_OK
End Function ' PHYSFS_getLastErrorCode

' PHYSFS_ErrorCode's messages (vbNullString: NULL, an unknown code)
Public Function PHYSFS_getErrorByCode(ByVal code As Long) As String
    Select Case code
    Case PHYSFS_ERR_OK: PHYSFS_getErrorByCode = "no error"
    Case PHYSFS_ERR_OTHER_ERROR: PHYSFS_getErrorByCode = "unknown error"
    Case PHYSFS_ERR_OUT_OF_MEMORY: PHYSFS_getErrorByCode = "out of memory"
    Case PHYSFS_ERR_NOT_INITIALIZED: PHYSFS_getErrorByCode = "not initialized"
    Case PHYSFS_ERR_IS_INITIALIZED: PHYSFS_getErrorByCode = "already initialized"
    Case PHYSFS_ERR_ARGV0_IS_NULL: PHYSFS_getErrorByCode = "argv[0] is NULL"
    Case PHYSFS_ERR_UNSUPPORTED: PHYSFS_getErrorByCode = "unsupported"
    Case PHYSFS_ERR_PAST_EOF: PHYSFS_getErrorByCode = "past end of file"
    Case PHYSFS_ERR_FILES_STILL_OPEN: PHYSFS_getErrorByCode = "files still open"
    Case PHYSFS_ERR_INVALID_ARGUMENT: PHYSFS_getErrorByCode = "invalid argument"
    Case PHYSFS_ERR_NOT_MOUNTED: PHYSFS_getErrorByCode = "not mounted"
    Case PHYSFS_ERR_NOT_FOUND: PHYSFS_getErrorByCode = "not found"
    Case PHYSFS_ERR_SYMLINK_FORBIDDEN: PHYSFS_getErrorByCode = "symlinks are forbidden"
    Case PHYSFS_ERR_NO_WRITE_DIR: PHYSFS_getErrorByCode = "write directory is not set"
    Case PHYSFS_ERR_OPEN_FOR_READING: PHYSFS_getErrorByCode = "file open for reading"
    Case PHYSFS_ERR_OPEN_FOR_WRITING: PHYSFS_getErrorByCode = "file open for writing"
    Case PHYSFS_ERR_NOT_A_FILE: PHYSFS_getErrorByCode = "not a file"
    Case PHYSFS_ERR_READ_ONLY: PHYSFS_getErrorByCode = "read-only filesystem"
    Case PHYSFS_ERR_CORRUPT: PHYSFS_getErrorByCode = "corrupted"
    Case PHYSFS_ERR_SYMLINK_LOOP: PHYSFS_getErrorByCode = "infinite symbolic link loop"
    Case PHYSFS_ERR_IO: PHYSFS_getErrorByCode = "i/o error"
    Case PHYSFS_ERR_PERMISSION: PHYSFS_getErrorByCode = "permission denied"
    Case PHYSFS_ERR_NO_SPACE: PHYSFS_getErrorByCode = "no space available for writing"
    Case PHYSFS_ERR_BAD_FILENAME: PHYSFS_getErrorByCode = "filename is illegal or insecure"
    Case PHYSFS_ERR_BUSY: PHYSFS_getErrorByCode = "tried to modify a file the OS needs"
    Case PHYSFS_ERR_DIR_NOT_EMPTY: PHYSFS_getErrorByCode = "directory isn't empty"
    Case PHYSFS_ERR_OS_ERROR: PHYSFS_getErrorByCode = "OS reported an error"
    Case PHYSFS_ERR_DUPLICATE: PHYSFS_getErrorByCode = "duplicate resource"
    Case PHYSFS_ERR_BAD_PASSWORD: PHYSFS_getErrorByCode = "bad password"
    Case PHYSFS_ERR_APP_CALLBACK: PHYSFS_getErrorByCode = "app callback reported error"
    Case Else: PHYSFS_getErrorByCode = vbNullString  ' don't know this error code.
    End Select
End Function ' PHYSFS_getErrorByCode

' physfs_ext.h: #define WZ_PHYSFS_getLastError() PHYSFS_getErrorByCode(PHYSFS_getLastErrorCode())
Public Function WZ_PHYSFS_getLastError() As String
    WZ_PHYSFS_getLastError = PHYSFS_getErrorByCode(PHYSFS_getLastErrorCode())
End Function

'--------------------------------------------------------------------------
' physfs_platform_windows.c

Private Function errcodeFromWinApiError(ByVal winErr As Long) As Long
    '
    ' win32 error codes are sort of a tricky thing; Microsoft intentionally
    '  doesn't list which ones a given API might trigger, there are several
    '  with overlapping and unclear meanings...and there's 16 thousand of
    '  them in Windows 7. It looks like the ones we care about are in the
    '  first 500, but I can't say this list is perfect; we might miss
    '  important values or misinterpret others.
    '
    ' Don't treat this list as anything other than a work in progress.
    '
    Select Case winErr
    Case ERROR_SUCCESS: errcodeFromWinApiError = PHYSFS_ERR_OK
    Case ERROR_ACCESS_DENIED: errcodeFromWinApiError = PHYSFS_ERR_PERMISSION
    Case ERROR_NETWORK_ACCESS_DENIED: errcodeFromWinApiError = PHYSFS_ERR_PERMISSION
    Case ERROR_NOT_READY: errcodeFromWinApiError = PHYSFS_ERR_IO
    Case ERROR_CRC: errcodeFromWinApiError = PHYSFS_ERR_IO
    Case ERROR_SEEK: errcodeFromWinApiError = PHYSFS_ERR_IO
    Case ERROR_SECTOR_NOT_FOUND: errcodeFromWinApiError = PHYSFS_ERR_IO
    Case ERROR_NOT_DOS_DISK: errcodeFromWinApiError = PHYSFS_ERR_IO
    Case ERROR_WRITE_FAULT: errcodeFromWinApiError = PHYSFS_ERR_IO
    Case ERROR_READ_FAULT: errcodeFromWinApiError = PHYSFS_ERR_IO
    Case ERROR_DEV_NOT_EXIST: errcodeFromWinApiError = PHYSFS_ERR_IO
    Case ERROR_BUFFER_OVERFLOW: errcodeFromWinApiError = PHYSFS_ERR_BAD_FILENAME
    Case ERROR_INVALID_NAME: errcodeFromWinApiError = PHYSFS_ERR_BAD_FILENAME
    Case ERROR_BAD_PATHNAME: errcodeFromWinApiError = PHYSFS_ERR_BAD_FILENAME
    Case ERROR_DIRECTORY: errcodeFromWinApiError = PHYSFS_ERR_BAD_FILENAME
    Case ERROR_FILE_NOT_FOUND: errcodeFromWinApiError = PHYSFS_ERR_NOT_FOUND
    Case ERROR_PATH_NOT_FOUND: errcodeFromWinApiError = PHYSFS_ERR_NOT_FOUND
    Case ERROR_DELETE_PENDING: errcodeFromWinApiError = PHYSFS_ERR_NOT_FOUND
    Case ERROR_INVALID_DRIVE: errcodeFromWinApiError = PHYSFS_ERR_NOT_FOUND
    Case ERROR_HANDLE_DISK_FULL: errcodeFromWinApiError = PHYSFS_ERR_NO_SPACE
    Case ERROR_DISK_FULL: errcodeFromWinApiError = PHYSFS_ERR_NO_SPACE
    Case ERROR_WRITE_PROTECT: errcodeFromWinApiError = PHYSFS_ERR_READ_ONLY
    Case ERROR_LOCK_VIOLATION: errcodeFromWinApiError = PHYSFS_ERR_BUSY
    Case ERROR_SHARING_VIOLATION: errcodeFromWinApiError = PHYSFS_ERR_BUSY
    Case ERROR_CURRENT_DIRECTORY: errcodeFromWinApiError = PHYSFS_ERR_BUSY
    Case ERROR_DRIVE_LOCKED: errcodeFromWinApiError = PHYSFS_ERR_BUSY
    Case ERROR_PATH_BUSY: errcodeFromWinApiError = PHYSFS_ERR_BUSY
    Case ERROR_BUSY: errcodeFromWinApiError = PHYSFS_ERR_BUSY
    Case ERROR_NOT_ENOUGH_MEMORY: errcodeFromWinApiError = PHYSFS_ERR_OUT_OF_MEMORY
    Case ERROR_OUTOFMEMORY: errcodeFromWinApiError = PHYSFS_ERR_OUT_OF_MEMORY
    Case ERROR_DIR_NOT_EMPTY: errcodeFromWinApiError = PHYSFS_ERR_DIR_NOT_EMPTY
    Case Else: errcodeFromWinApiError = PHYSFS_ERR_OS_ERROR
    End Select
End Function ' errcodeFromWinApiError

' (VB) errcodeFromWinApi(): GetLastError() is VB's Err.LastDllError of the last Declare'd call
Private Function errcodeFromWinApi() As Long
    errcodeFromWinApi = errcodeFromWinApiError(Err.LastDllError)
End Function ' errcodeFromWinApi

' PHYSFS_Stat st: (VB) the fields; follow: 0 reports a symlink as one
Private Function PHYSFS_platformStat__(ByVal filename As String, st As PHYSFS_Stat_, ByVal follow As Long) As Long
    Dim winstat As WIN32_FILE_ATTRIBUTE_DATA, errNum As Long, rc As Long, issymlink As Boolean, w32dw As WIN32_FIND_DATAW, h As Long

    rc = GetFileAttributesExW(StrPtr(filename), 0, winstat)     ' GetFileExInfoStandard

    If rc = 0 Then
        errNum = Err.LastDllError
    ElseIf follow = 0 Then   ' check for symlink while wstr is still available (isSymlink)
        If (winstat.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT) <> 0 Then
            h = FindFirstFileW(StrPtr(filename), w32dw)
            If h <> INVALID_HANDLE_VALUE Then
                FindClose h
                issymlink = (w32dw.dwReserved0 = IO_REPARSE_TAG_SYMLINK)
            End If
        End If
    End If

    If rc = 0 Then
        PHYSFS_setErrorCode errcodeFromWinApiError(errNum)
        PHYSFS_platformStat__ = 0
        Exit Function
    End If

    st.modtime = FileTimeToPhysfsTime(winstat.ftLastWriteTime)
    st.accesstime = FileTimeToPhysfsTime(winstat.ftLastAccessTime)
    st.createtime = FileTimeToPhysfsTime(winstat.ftCreationTime)

    If issymlink Then
        st.filetype = PHYSFS_FILETYPE_SYMLINK
        st.filesize = 0

    ElseIf (winstat.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
        st.filetype = PHYSFS_FILETYPE_DIRECTORY
        st.filesize = 0

    ElseIf (winstat.dwFileAttributes And (FILE_ATTRIBUTE_OFFLINE Or FILE_ATTRIBUTE_DEVICE)) <> 0 Then
        st.filetype = PHYSFS_FILETYPE_OTHER
        st.filesize = Unsigned(winstat.nFileSizeHigh) * 4294967296# + Unsigned(winstat.nFileSizeLow)

    Else
        st.filetype = PHYSFS_FILETYPE_REGULAR
        st.filesize = Unsigned(winstat.nFileSizeHigh) * 4294967296# + Unsigned(winstat.nFileSizeLow)
    End If

    If (winstat.dwFileAttributes And FILE_ATTRIBUTE_READONLY) <> 0 Then st.readonly_ = 1 Else st.readonly_ = 0

    PHYSFS_platformStat__ = 1
End Function ' __PHYSFS_platformStat

' (VB) callback: PHYSFS_EnumerateCallback_call's number
Private Function PHYSFS_platformEnumerate__(ByVal dirname As String, ByVal callback As Long, ByVal origdir As String) As Long
    Dim retval As Long, dir_ As Long, entw As WIN32_FIND_DATAW, searchPath_ As String, fn As String, k As Long
    retval = PHYSFS_ENUM_OK

    ' Copy current dirname
    searchPath_ = dirname

    ' if there's no '\\' at the end of the path, stick one in there.
    If Right$(searchPath_, 1) <> "\" Then
        searchPath_ = searchPath_ & "\"
    End If

    ' Append the "*" to the end of the string
    searchPath_ = searchPath_ & "*"

    dir_ = FindFirstFileW(StrPtr(searchPath_), entw)
    If dir_ = INVALID_HANDLE_VALUE Then
        PHYSFS_setErrorCode errcodeFromWinApi()
        PHYSFS_platformEnumerate__ = PHYSFS_ENUM_ERROR
        Exit Function
    End If

    Do
        fn = entw.cFileName
        k = InStr(fn, vbNullChar)
        If k > 0 Then fn = Left$(fn, k - 1)

        If fn = "." Or fn = ".." Then  ' ignore "." and ".."
            GoTo continue_
        End If

        retval = PHYSFS_EnumerateCallback_call(callback, origdir, fn)
        If retval = PHYSFS_ENUM_ERROR Then
            PHYSFS_setErrorCode PHYSFS_ERR_APP_CALLBACK
        End If
continue_:
        If retval <> PHYSFS_ENUM_OK Then Exit Do
    Loop While FindNextFileW(dir_, entw) <> 0

    FindClose dir_

    PHYSFS_platformEnumerate__ = retval
End Function ' __PHYSFS_platformEnumerate

Private Function PHYSFS_platformMkDir__(ByVal path As String) As Long
    Dim rc As Long
    rc = CreateDirectoryW(StrPtr(path), 0)
    If rc = 0 Then
        PHYSFS_setErrorCode errcodeFromWinApi()
        PHYSFS_platformMkDir__ = 0
        Exit Function
    End If
    PHYSFS_platformMkDir__ = 1
End Function ' __PHYSFS_platformMkDir

' doOpen: the handle, or INVALID_HANDLE_VALUE with the error set
Private Function platformDoOpen(ByVal fname As String, ByVal mode As Long, ByVal creation As Long) As Long
    Dim fileh As Long
    fileh = CreateFileW(StrPtr(fname), mode, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, creation, FILE_ATTRIBUTE_NORMAL, 0)  ' winCreateFileW

    If fileh = INVALID_HANDLE_VALUE Then
        PHYSFS_setErrorCode errcodeFromWinApi()
        platformDoOpen = INVALID_HANDLE_VALUE
        Exit Function
    End If

    platformDoOpen = fileh
End Function ' doOpen

' (VB) a handle, 0 = NULL
Private Function PHYSFS_platformOpenRead__(ByVal filename As String) As Long
    Dim h As Long
    h = platformDoOpen(filename, GENERIC_READ, OPEN_EXISTING)
    If h = INVALID_HANDLE_VALUE Then PHYSFS_platformOpenRead__ = 0 Else PHYSFS_platformOpenRead__ = h
End Function ' __PHYSFS_platformOpenRead

Private Function PHYSFS_platformOpenWrite__(ByVal filename As String) As Long
    Dim h As Long
    h = platformDoOpen(filename, GENERIC_WRITE, CREATE_ALWAYS)
    If h = INVALID_HANDLE_VALUE Then PHYSFS_platformOpenWrite__ = 0 Else PHYSFS_platformOpenWrite__ = h
End Function ' __PHYSFS_platformOpenWrite

Private Function PHYSFS_platformOpenAppend__(ByVal filename As String) As Long
    Dim h As Long, errc As Long
    h = platformDoOpen(filename, GENERIC_WRITE, OPEN_ALWAYS)
    If h = INVALID_HANDLE_VALUE Then Exit Function

    If SetFilePointerEx(h, 0, 0, 0, FILE_END) = 0 Then  ' winSetFilePointer
        errc = errcodeFromWinApi()
        CloseHandle h
        PHYSFS_setErrorCode errc
        Exit Function
    End If

    PHYSFS_platformOpenAppend__ = h
End Function ' __PHYSFS_platformOpenAppend

' (VB) buf(ofs .. ofs + len - 1)
Private Function PHYSFS_platformRead__(ByVal opaque As Long, buf() As Byte, ByVal ofs As Long, ByVal len_ As Long) As Long
    Dim totalRead As Long, numRead As Long
    totalRead = 0

    ' (if (!__PHYSFS_ui64FitsAddressSpace(len)) ...: a Long fits)

    Do While len_ > 0
        numRead = 0
        If ReadFile(opaque, VarPtr(buf(ofs + totalRead)), len_, numRead, 0) = 0 Then
            PHYSFS_setErrorCode errcodeFromWinApi()
            PHYSFS_platformRead__ = -1
            Exit Function
        End If
        len_ = len_ - numRead
        totalRead = totalRead + numRead
        If numRead <> len_ + numRead Then           ' (numRead != thislen)
            Exit Do
        End If
    Loop

    PHYSFS_platformRead__ = totalRead
End Function ' __PHYSFS_platformRead

Private Function PHYSFS_platformWrite__(ByVal opaque As Long, buffer() As Byte, ByVal ofs As Long, ByVal len_ As Long) As Long
    Dim totalWritten As Long, numWritten As Long
    totalWritten = 0

    Do While len_ > 0
        numWritten = 0
        If WriteFile(opaque, VarPtr(buffer(ofs + totalWritten)), len_, numWritten, 0) = 0 Then
            PHYSFS_setErrorCode errcodeFromWinApi()
            PHYSFS_platformWrite__ = -1
            Exit Function
        End If
        len_ = len_ - numWritten
        totalWritten = totalWritten + numWritten
        If numWritten <> len_ + numWritten Then     ' (numWritten != thislen)
            Exit Do
        End If
    Loop

    PHYSFS_platformWrite__ = totalWritten
End Function ' __PHYSFS_platformWrite

Private Function PHYSFS_platformSeek__(ByVal opaque As Long, ByVal pos As Double) As Long
    Dim hi As Double, lo As Double, l As Long
    hi = Int(pos / 4294967296#)
    lo = pos - hi * 4294967296#
    If lo >= 2147483648# Then l = lo - 4294967296# Else l = lo
    If SetFilePointerEx(opaque, l, CLng(hi), 0, FILE_BEGIN) = 0 Then
        PHYSFS_setErrorCode errcodeFromWinApi()
        PHYSFS_platformSeek__ = 0
        Exit Function
    End If
    PHYSFS_platformSeek__ = 1  ' No error occured
End Function ' __PHYSFS_platformSeek

Private Function PHYSFS_platformTell__(ByVal opaque As Long) As Double
    Dim pos(0 To 1) As Long
    If SetFilePointerEx(opaque, 0, 0, VarPtr(pos(0)), FILE_CURRENT) = 0 Then
        PHYSFS_setErrorCode errcodeFromWinApi()
        PHYSFS_platformTell__ = -1
        Exit Function
    End If
    PHYSFS_platformTell__ = CDbl(pos(1)) * 4294967296# + Unsigned(pos(0))
End Function ' __PHYSFS_platformTell

Private Function PHYSFS_platformFileLength__(ByVal opaque As Long) As Double
    Dim size(0 To 1) As Long, retval As Double
    If GetFileSizeEx(opaque, VarPtr(size(0))) = 0 Then  ' winGetFileSize
        retval = -1
    Else
        retval = CDbl(size(1)) * 4294967296# + Unsigned(size(0))
    End If
    If retval < 0 Then
        PHYSFS_setErrorCode errcodeFromWinApi()
        PHYSFS_platformFileLength__ = -1
        Exit Function
    End If
    PHYSFS_platformFileLength__ = retval
End Function ' __PHYSFS_platformFileLength

Private Function PHYSFS_platformFlush__(ByVal opaque As Long) As Long
    If FlushFileBuffers(opaque) = 0 Then
        PHYSFS_setErrorCode errcodeFromWinApi()
        PHYSFS_platformFlush__ = 0
        Exit Function
    End If
    PHYSFS_platformFlush__ = 1
End Function ' __PHYSFS_platformFlush

Private Sub PHYSFS_platformClose__(ByVal opaque As Long)
    CloseHandle opaque  ' ignore errors. You should have flushed!
End Sub ' __PHYSFS_platformClose

Private Function PHYSFS_platformDelete__(ByVal path As String) As Long
    Dim info As WIN32_FILE_ATTRIBUTE_DATA, isdir As Boolean, rc As Long
    If GetFileAttributesExW(StrPtr(path), 0, info) = 0 Then   ' doPlatformDelete
        PHYSFS_setErrorCode errcodeFromWinApi()
        PHYSFS_platformDelete__ = 0
        Exit Function
    Else
        isdir = (info.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0
        If isdir Then rc = RemoveDirectoryW(StrPtr(path)) Else rc = DeleteFileW(StrPtr(path))
        If rc = 0 Then
            PHYSFS_setErrorCode errcodeFromWinApi()
            PHYSFS_platformDelete__ = 0
            Exit Function
        End If
    End If
    PHYSFS_platformDelete__ = 1    ' if you made it here, it worked.
End Function ' __PHYSFS_platformDelete

'--------------------------------------------------------------------------
' PHYSFS_Io implementation for i/o to physical filesystem... and the dispatch on an io's function table

' (VB) a new PHYSFS_Io of a kind (allocator.Malloc(sizeof (PHYSFS_Io)) and the table's memcpy)
Public Function PHYSFS_Io_new(ByVal kind As Long, ByVal opaque As Long) As Long
    Dim i As Long, blank As PHYSFS_Io
    For i = 1 To iosCount
        If ios(i).kind = 0 Then Exit For
    Next
    If i > iosCount Then
        iosCount = iosCount + 16
        ReDim Preserve ios(1 To iosCount)
    End If
    ios(i) = blank
    ios(i).kind = kind
    ios(i).opaque = opaque
    PHYSFS_Io_new = i
End Function

' io->opaque
Public Function PHYSFS_Io_opaque(ByVal io As Long) As Long
    PHYSFS_Io_opaque = ios(io).opaque
End Function

Private Sub PHYSFS_Io_free(ByVal io As Long)
    Dim blank As PHYSFS_Io
    ios(io) = blank                                     ' allocator.Free(io)
End Sub

' io->read(io, buf, len): (VB) into buf(ofs ..)
Public Function PHYSFS_Io_read(ByVal io As Long, buf() As Byte, ByVal ofs As Long, ByVal len_ As Long) As Long
    Select Case ios(io).kind
    Case PHYSFS_IO_NATIVE: PHYSFS_Io_read = PHYSFS_platformRead__(ios(io).handle, buf, ofs, len_)    ' nativeIo_read
    Case PHYSFS_IO_ZIP: PHYSFS_Io_read = ZIP_read(io, buf, ofs, len_)
    End Select
End Function

Public Function PHYSFS_Io_write(ByVal io As Long, buf() As Byte, ByVal ofs As Long, ByVal len_ As Long) As Long
    Select Case ios(io).kind
    Case PHYSFS_IO_NATIVE: PHYSFS_Io_write = PHYSFS_platformWrite__(ios(io).handle, buf, ofs, len_)  ' nativeIo_write
    Case PHYSFS_IO_ZIP: PHYSFS_Io_write = ZIP_write(io, buf, ofs, len_)
    End Select
End Function

Public Function PHYSFS_Io_seek(ByVal io As Long, ByVal offset As Double) As Long
    Select Case ios(io).kind
    Case PHYSFS_IO_NATIVE: PHYSFS_Io_seek = PHYSFS_platformSeek__(ios(io).handle, offset)            ' nativeIo_seek
    Case PHYSFS_IO_ZIP: PHYSFS_Io_seek = ZIP_seek(io, offset)
    End Select
End Function

Public Function PHYSFS_Io_tell(ByVal io As Long) As Double
    Select Case ios(io).kind
    Case PHYSFS_IO_NATIVE: PHYSFS_Io_tell = PHYSFS_platformTell__(ios(io).handle)                    ' nativeIo_tell
    Case PHYSFS_IO_ZIP: PHYSFS_Io_tell = ZIP_tell(io)
    End Select
End Function

Public Function PHYSFS_Io_length(ByVal io As Long) As Double
    Select Case ios(io).kind
    Case PHYSFS_IO_NATIVE: PHYSFS_Io_length = PHYSFS_platformFileLength__(ios(io).handle)            ' nativeIo_length
    Case PHYSFS_IO_ZIP: PHYSFS_Io_length = ZIP_length(io)
    End Select
End Function

Public Function PHYSFS_Io_duplicate(ByVal io As Long) As Long
    Select Case ios(io).kind
    Case PHYSFS_IO_NATIVE: PHYSFS_Io_duplicate = PHYSFS_createNativeIo__(ios(io).path, ios(io).mode) ' nativeIo_duplicate
    Case PHYSFS_IO_ZIP: PHYSFS_Io_duplicate = ZIP_duplicate(io)
    End Select
End Function

Public Function PHYSFS_Io_flush(ByVal io As Long) As Long
    Select Case ios(io).kind
    Case PHYSFS_IO_NATIVE: PHYSFS_Io_flush = PHYSFS_platformFlush__(ios(io).handle)                  ' nativeIo_flush
    Case PHYSFS_IO_ZIP: PHYSFS_Io_flush = ZIP_flush(io)
    End Select
End Function

Public Sub PHYSFS_Io_destroy(ByVal io As Long)
    Select Case ios(io).kind
    Case PHYSFS_IO_NATIVE                                                                               ' nativeIo_destroy
        PHYSFS_platformClose__ ios(io).handle
        ' (allocator.Free((void *) info->path); allocator.Free(info);)
        PHYSFS_Io_free io
    Case PHYSFS_IO_ZIP
        ZIP_destroy io
        PHYSFS_Io_free io                               ' (ZIP_destroy's allocator.Free(io))
    End Select
End Sub

' (VB) mode: 'r', 'w' or 'a' (MODE_R ...); the io, 0 = NULL
Public Function PHYSFS_createNativeIo__(ByVal path As String, ByVal mode As Long) As Long
    Dim io As Long, handle As Long

    ' (assert((mode == 'r') || (mode == 'w') || (mode == 'a'));)

    If mode = MODE_R Then
        handle = PHYSFS_platformOpenRead__(path)
    ElseIf mode = MODE_W Then
        handle = PHYSFS_platformOpenWrite__(path)
    ElseIf mode = MODE_A Then
        handle = PHYSFS_platformOpenAppend__(path)
    End If

    If handle = 0 Then Exit Function

    io = PHYSFS_Io_new(PHYSFS_IO_NATIVE, 0)             ' memcpy(io, &__PHYSFS_nativeIoInterface, sizeof (*io))
    ios(io).handle = handle
    ios(io).path = path
    ios(io).mode = mode
    PHYSFS_createNativeIo__ = io
End Function ' __PHYSFS_createNativeIo

Public Function PHYSFS_readAll__(ByVal io As Long, buf() As Byte, ByVal ofs As Long, ByVal len_ As Long) As Long
    If len_ < 0 Then Exit Function                      ' (VB) a size_t past the address space: the read fails
    If PHYSFS_Io_read(io, buf, ofs, len_) = len_ Then PHYSFS_readAll__ = 1
End Function ' __PHYSFS_readAll

'--------------------------------------------------------------------------
' physfs_archiver_dir.c: Standard directory I/O support routines for PhysicsFS.

Private Function cvtToDependent(ByVal prepend As String, ByVal path As String) As String
    cvtToDependent = prepend & Replace(path, "/", "\")
End Function ' cvtToDependent

Private Function DIR_openArchive(ByVal io As Long, ByVal name_ As String, ByVal forWriting As Long, claimed As Long) As Long
    Dim st As PHYSFS_Stat_, retval As String, i As Long

    ' (assert(io == NULL);  /* shouldn't create an Io for these. */)
    If PHYSFS_platformStat__(name_, st, 1) = 0 Then Exit Function

    If st.filetype <> PHYSFS_FILETYPE_DIRECTORY Then
        PHYSFS_setErrorCode PHYSFS_ERR_UNSUPPORTED
        Exit Function
    End If

    claimed = 1
    retval = name_

    ' make sure there's a dir separator at the end of the string
    If Right$(retval, 1) <> "\" Then
        retval = retval & "\"
    End If

    ' (VB) allocator.Malloc: the pool
    For i = 1 To dirOpaquesCount
        If Not dirOpaqueUsed(i) Then Exit For
    Next
    If i > dirOpaquesCount Then
        dirOpaquesCount = dirOpaquesCount + 8
        ReDim Preserve dirOpaques(1 To dirOpaquesCount)
        ReDim Preserve dirOpaqueUsed(1 To dirOpaquesCount)
    End If
    dirOpaques(i) = retval
    dirOpaqueUsed(i) = True
    DIR_openArchive = i
End Function ' DIR_openArchive

Private Function DIR_enumerate(ByVal opaque As Long, ByVal dname As String, ByVal cb As Long, ByVal origdir As String) As Long
    DIR_enumerate = PHYSFS_platformEnumerate__(cvtToDependent(dirOpaques(opaque), dname), cb, origdir)
End Function ' DIR_enumerate

Private Function DIR_doOpen(ByVal opaque As Long, ByVal name_ As String, ByVal mode As Long) As Long
    Dim io As Long, f As String, errc As Long, statbuf As PHYSFS_Stat_

    f = cvtToDependent(dirOpaques(opaque), name_)

    io = PHYSFS_createNativeIo__(f, mode)
    If io = 0 Then
        errc = PHYSFS_getLastErrorCode()
        PHYSFS_platformStat__ f, statbuf, 0  ' !!! FIXME: why are we stating here?
        PHYSFS_setErrorCode errc
    End If

    DIR_doOpen = io
End Function ' doOpen

Private Function DIR_openRead(ByVal opaque As Long, ByVal filename As String) As Long
    DIR_openRead = DIR_doOpen(opaque, filename, MODE_R)
End Function ' DIR_openRead

Private Function DIR_openWrite(ByVal opaque As Long, ByVal filename As String) As Long
    DIR_openWrite = DIR_doOpen(opaque, filename, MODE_W)
End Function ' DIR_openWrite

Private Function DIR_openAppend(ByVal opaque As Long, ByVal filename As String) As Long
    DIR_openAppend = DIR_doOpen(opaque, filename, MODE_A)
End Function ' DIR_openAppend

Private Function DIR_remove(ByVal opaque As Long, ByVal name_ As String) As Long
    DIR_remove = PHYSFS_platformDelete__(cvtToDependent(dirOpaques(opaque), name_))
End Function ' DIR_remove

Private Function DIR_mkdir(ByVal opaque As Long, ByVal name_ As String) As Long
    DIR_mkdir = PHYSFS_platformMkDir__(cvtToDependent(dirOpaques(opaque), name_))
End Function ' DIR_mkdir

Private Sub DIR_closeArchive(ByVal opaque As Long)
    dirOpaques(opaque) = vbNullString                   ' allocator.Free(opaque)
    dirOpaqueUsed(opaque) = False
End Sub ' DIR_closeArchive

Private Function DIR_stat(ByVal opaque As Long, ByVal name_ As String, stat As PHYSFS_Stat_) As Long
    DIR_stat = PHYSFS_platformStat__(cvtToDependent(dirOpaques(opaque), name_), stat, 0)
End Function ' DIR_stat

'--------------------------------------------------------------------------
' (VB) the archivers' function tables: h->funcs->openArchive(...) is arcOpenArchive(h.funcs, ...)

Private Function arcOpenArchive(ByVal funcs As Long, ByVal io As Long, ByVal d As String, ByVal forWriting As Long, claimed As Long) As Long
    Select Case funcs
    Case ARCHIVER_DIR: arcOpenArchive = DIR_openArchive(io, d, forWriting, claimed)
    Case ARCHIVER_ZIP: arcOpenArchive = ZIP_openArchive(io, d, forWriting, claimed)
    End Select
End Function

Private Function arcEnumerate(ByVal funcs As Long, ByVal opaque As Long, ByVal dname As String, ByVal cb As Long, ByVal origdir As String) As Long
    Select Case funcs
    Case ARCHIVER_DIR: arcEnumerate = DIR_enumerate(opaque, dname, cb, origdir)
    Case ARCHIVER_ZIP: arcEnumerate = ZIP_enumerate(opaque, dname, cb, origdir)        ' __PHYSFS_DirTreeEnumerate
    End Select
End Function

Private Function arcOpenRead(ByVal funcs As Long, ByVal opaque As Long, ByVal fnm As String) As Long
    Select Case funcs
    Case ARCHIVER_DIR: arcOpenRead = DIR_openRead(opaque, fnm)
    Case ARCHIVER_ZIP: arcOpenRead = ZIP_openRead(opaque, fnm)
    End Select
End Function

Private Function arcOpenWrite(ByVal funcs As Long, ByVal opaque As Long, ByVal fnm As String) As Long
    Select Case funcs
    Case ARCHIVER_DIR: arcOpenWrite = DIR_openWrite(opaque, fnm)
    Case ARCHIVER_ZIP: arcOpenWrite = ZIP_openWrite(opaque, fnm)
    End Select
End Function

Private Function arcOpenAppend(ByVal funcs As Long, ByVal opaque As Long, ByVal fnm As String) As Long
    Select Case funcs
    Case ARCHIVER_DIR: arcOpenAppend = DIR_openAppend(opaque, fnm)
    Case ARCHIVER_ZIP: arcOpenAppend = ZIP_openAppend(opaque, fnm)
    End Select
End Function

Private Function arcRemove(ByVal funcs As Long, ByVal opaque As Long, ByVal fnm As String) As Long
    Select Case funcs
    Case ARCHIVER_DIR: arcRemove = DIR_remove(opaque, fnm)
    Case ARCHIVER_ZIP: arcRemove = ZIP_remove(opaque, fnm)
    End Select
End Function

Private Function arcMkdir(ByVal funcs As Long, ByVal opaque As Long, ByVal fnm As String) As Long
    Select Case funcs
    Case ARCHIVER_DIR: arcMkdir = DIR_mkdir(opaque, fnm)
    Case ARCHIVER_ZIP: arcMkdir = ZIP_mkdir(opaque, fnm)
    End Select
End Function

Private Function arcStat(ByVal funcs As Long, ByVal opaque As Long, ByVal fnm As String, stat As PHYSFS_Stat_) As Long
    Select Case funcs
    Case ARCHIVER_DIR: arcStat = DIR_stat(opaque, fnm, stat)
    Case ARCHIVER_ZIP: arcStat = ZIP_stat(opaque, fnm, stat)
    End Select
End Function

Private Sub arcCloseArchive(ByVal funcs As Long, ByVal opaque As Long)
    Select Case funcs
    Case ARCHIVER_DIR: DIR_closeArchive opaque
    Case ARCHIVER_ZIP: ZIP_closeArchive opaque
    End Select
End Sub

' funcs->info.extension ("" for DIR)
Private Function arcExtension(ByVal funcs As Long) As String
    If funcs = ARCHIVER_ZIP Then arcExtension = "ZIP"
End Function

' funcs->info.supportsSymlinks (both say 1)
Private Function arcSupportsSymlinks(ByVal funcs As Long) As Boolean
    arcSupportsSymlinks = True
End Function

' archivers[i] (initStaticArchivers' order: ZIP first; 7Z, GRP, ... are not ported)
Private Function archiverAt(ByVal i As Long) As Long
    If i = 0 Then archiverAt = ARCHIVER_ZIP
End Function

'--------------------------------------------------------------------------
' physfs.c

' (VB) the extension in ext; False: NULL
Private Function find_filename_extension(ByVal fname As String, ext As String) As Boolean
    Dim p As Long
    If StrPtr(fname) <> 0 Then
        p = InStrRev(fname, ".")                        ' the last '.'
        If p > 0 Then
            ext = Mid$(fname, p + 1)                    ' skip '.'
            find_filename_extension = True
        End If
    End If
End Function ' find_filename_extension

' (VB) retval: the new DirHandle (its opaque and funcs); False: NULL
Private Function tryOpenDir(ByVal io As Long, ByVal funcs As Long, ByVal d As String, ByVal forWriting As Long, claimed As Long, retval As DirHandle) As Boolean
    Dim opaque As Long, blank As DirHandle

    If io <> 0 Then
        If PHYSFS_Io_seek(io, 0) = 0 Then Exit Function
    End If

    opaque = arcOpenArchive(funcs, io, d, forWriting, claimed)
    If opaque <> 0 Then
        retval = blank                                  ' memset(retval, '\0', sizeof (DirHandle))
        retval.mountPoint = vbNullString
        retval.funcs = funcs
        retval.opaque = opaque
        tryOpenDir = True
    End If
End Function ' tryOpenDir

Private Function openDirectory(ByVal io As Long, ByVal d As String, ByVal forWriting As Long, retval As DirHandle) As Boolean
    Dim i As Long, ext As String, created_io As Long, claimed As Long, errcode As Long, statbuf As PHYSFS_Stat_, found As Boolean

    ' (assert((io != NULL) || (d != NULL));)

    If io = 0 Then
        ' file doesn't exist, etc? Just fail out.
        If PHYSFS_platformStat__(d, statbuf, 1) = 0 Then Exit Function

        ' DIR gets first shot (unlike the rest, it doesn't deal with files).
        If statbuf.filetype = PHYSFS_FILETYPE_DIRECTORY Then
            found = tryOpenDir(io, ARCHIVER_DIR, d, forWriting, claimed, retval)
            If found Or (claimed <> 0) Then
                openDirectory = found
                Exit Function
            End If
        End If

        If forWriting <> 0 Then
            io = PHYSFS_createNativeIo__(d, MODE_W)
        Else
            io = PHYSFS_createNativeIo__(d, MODE_R)
        End If
        If io = 0 Then Exit Function
        created_io = 1
    End If

    If find_filename_extension(d, ext) Then
        ' Look for archivers with matching file extensions first...
        i = 0
        Do While (i < numArchivers) And (Not found) And (claimed = 0)
            If StrComp(ext, arcExtension(archiverAt(i)), vbTextCompare) = 0 Then     ' PHYSFS_utf8stricmp
                found = tryOpenDir(io, archiverAt(i), d, forWriting, claimed, retval)
            End If
            i = i + 1
        Loop

        ' failing an exact file extension match, try all the others...
        i = 0
        Do While (i < numArchivers) And (Not found) And (claimed = 0)
            If StrComp(ext, arcExtension(archiverAt(i)), vbTextCompare) <> 0 Then
                found = tryOpenDir(io, archiverAt(i), d, forWriting, claimed, retval)
            End If
            i = i + 1
        Loop

    Else  ' no extension? Try them all.
        i = 0
        Do While (i < numArchivers) And (Not found) And (claimed = 0)
            found = tryOpenDir(io, archiverAt(i), d, forWriting, claimed, retval)
            i = i + 1
        Loop
    End If

    If claimed <> 0 Then errcode = currentErrorCode() Else errcode = PHYSFS_ERR_UNSUPPORTED

    If (Not found) And (created_io <> 0) Then
        PHYSFS_Io_destroy io
    End If

    If Not found Then
        PHYSFS_setErrorCode errcode
        Exit Function
    End If
    openDirectory = True
End Function ' openDirectory

'
' Make a platform-independent path string sane. Doesn't actually check the
'  file hierarchy, it just cleans up the string.
'  (dst) must be a buffer at least as big as (src), as this is where the
'  cleaned up string is deposited.
' If there are illegal bits in the path (".." entries, etc) then we
'  return zero and (dst) is undefined. Non-zero if the path was sanitized.
'
Private Function sanitizePlatformIndependentPath(ByVal src As String, dst As String) As Boolean
    Dim s() As Integer, o() As Integer, n As Long, si As Long, di As Long, prev As Long, ch As Long

    n = Len(src)
    ReDim s(0 To n)                                     ' (VB) the characters, with C's terminating zero
    If n > 0 Then CopyMemory VarPtr(s(0)), StrPtr(src), n * 2
    ReDim o(0 To n)

    Do While s(si) = 47                                 ' skip initial '/' chars...
        si = si + 1
    Loop

    ' Make sure the entire string isn't "." or ".."
    If s(si) = 46 Then
        If s(si + 1) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_BAD_FILENAME: Exit Function
        If s(si + 1) = 46 Then
            If s(si + 2) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_BAD_FILENAME: Exit Function
        End If
    End If

    prev = 0                                            ' prev = dst
    di = 0
    Do
        ch = s(si)
        si = si + 1

        If (ch = 58) Or (ch = 92) Then                  ' illegal chars in a physfs path. (':' '\\')
            PHYSFS_setErrorCode PHYSFS_ERR_BAD_FILENAME
            Exit Function
        End If

        If ch = 47 Then     ' path separator.
            ' *dst = '\0';  "." and ".." are illegal pathnames.
            If di - prev = 1 Then
                If o(prev) = 46 Then PHYSFS_setErrorCode PHYSFS_ERR_BAD_FILENAME: Exit Function
            ElseIf di - prev = 2 Then
                If o(prev) = 46 And o(prev + 1) = 46 Then PHYSFS_setErrorCode PHYSFS_ERR_BAD_FILENAME: Exit Function
            End If

            Do While s(si) = 47     ' chop out doubles...
                si = si + 1
            Loop

            If s(si) = 0 Then       ' ends with a pathsep?
                Exit Do             ' we're done, don't add final pathsep to dst.
            End If

            prev = di + 1
        End If

        o(di) = ch                                      ' *(dst++) = ch
        di = di + 1
    Loop While ch <> 0

    If di > 0 Then
        If o(di - 1) = 0 Then di = di - 1               ' (VB) the terminating zero
    End If
    dst = String$(di, 0)
    If di > 0 Then CopyMemory StrPtr(dst), VarPtr(o(0)), di * 2
    sanitizePlatformIndependentPath = True
End Function ' sanitizePlatformIndependentPath

'
' Figure out if (fname) is part of (h)'s mountpoint. (fname) must be an
'  output from sanitizePlatformIndependentPath(), so that it is in a known
'  state.
'
' This only finds legitimate segments of a mountpoint. If the mountpoint is
'  "/a/b/c" and (fname) is "/a/b/c", "/", or "/a/b/c/d", then the results are
'  all zero. "/a/b" will succeed, though.
'
Private Function partOfMountPoint(h As DirHandle, ByVal fname As String) As Boolean
    Dim len_ As Long, mntpntlen As Long

    If StrPtr(h.mountPoint) = 0 Then
        Exit Function
    ElseIf Len(fname) = 0 Then
        partOfMountPoint = True
        Exit Function
    End If

    len_ = Len(fname)
    mntpntlen = Len(h.mountPoint)
    If len_ > mntpntlen Then    ' can't be a subset of mountpoint.
        Exit Function
    End If

    ' if true, must be not a match or a complete match, but not a subset.
    If (len_ + 1) = mntpntlen Then
        Exit Function
    End If

    If StrComp(fname, Left$(h.mountPoint, len_), vbBinaryCompare) <> 0 Then    ' !!! FIXME: case insensitive?
        Exit Function   ' not a match.
    End If

    ' make sure /a/b matches /a/b/ and not /a/bc ...
    partOfMountPoint = (Mid$(h.mountPoint, len_ + 1, 1) = "/")
End Function ' partOfMountPoint

' (VB) dirHandle: the new DirHandle; False: NULL
Private Function createDirHandle(ByVal io As Long, ByVal newDir As String, ByVal mountPoint As String, ByVal forWriting As Long, dh As DirHandle) As Boolean
    Dim tmpmntpnt As String

    ' (assert(newDir != NULL);  /* should have caught this higher up. */)

    If StrPtr(mountPoint) <> 0 Then
        If Not sanitizePlatformIndependentPath(mountPoint, tmpmntpnt) Then
            Exit Function                               ' goto badDirHandle
        End If
        mountPoint = tmpmntpnt  ' sanitized version.
    End If

    If Not openDirectory(io, newDir, forWriting, dh) Then Exit Function

    dh.dirName = newDir

    If StrPtr(mountPoint) <> 0 Then
        If Len(mountPoint) > 0 Then
            dh.mountPoint = mountPoint & "/"
        End If
    End If

    nextDirHandleId = nextDirHandleId + 1
    dh.id = nextDirHandleId
    createDirHandle = True
End Function ' createDirHandle

' (VB) forWritingList: the open list to check (openWriteList: True)
Private Function freeDirHandle(dh As DirHandle, ByVal forWritingList As Boolean) As Boolean
    Dim i As Long

    For i = 1 To filesCount
        If files(i).inUse And (files(i).forReading <> forWritingList) Then
            If files(i).dirHandle = dh.id Then
                PHYSFS_setErrorCode PHYSFS_ERR_FILES_STILL_OPEN
                Exit Function
            End If
        End If
    Next

    arcCloseArchive dh.funcs, dh.opaque

    ' (allocator.Free(dh->dirName); allocator.Free(dh->mountPoint); allocator.Free(dh);)
    freeDirHandle = True
End Function ' freeDirHandle

Public Function PHYSFS_init(ByVal argv0 As String) As Long
    Dim buf As String, n As Long
    If initialized Then
        PHYSFS_setErrorCode PHYSFS_ERR_IS_INITIALIZED
        Exit Function
    End If

    ' (the allocator, __PHYSFS_platformInit, the mutexes)

    ' calculateBaseDir: __PHYSFS_platformCalcBaseDir (the program's folder)
    buf = String$(1024, 0)
    n = GetModuleFileNameW(0, StrPtr(buf), 1024)
    buf = Left$(buf, n)
    baseDir = Left$(buf, InStrRev(buf, "\"))

    ' (userDir = __PHYSFS_platformCalcUserDir(): not kept)

    ' (initStaticArchivers: the ZIP archiver)

    initialized = True

    ' This makes sure that the error subsystem is initialized.
    PHYSFS_setErrorCode PHYSFS_getLastErrorCode()

    PHYSFS_init = 1
End Function ' PHYSFS_init

' (VB) forWritingList: openWriteList (True) or openReadList
Private Function closeFileHandleList(ByVal forWritingList As Boolean) As Long
    Dim i As Long, io As Long

    For i = 1 To filesCount
        If files(i).inUse And (files(i).forReading <> forWritingList) Then
            io = files(i).io

            If PHYSFS_Io_flush(io) = 0 Then
                closeFileHandleList = 0
                Exit Function
            End If

            PHYSFS_Io_destroy io
            files(i).inUse = False                      ' allocator.Free(i)
            files(i).io = 0
        End If
    Next

    closeFileHandleList = 1
End Function ' closeFileHandleList

Private Sub freeSearchPath()
    Dim i As Long

    closeFileHandleList False

    If searchCount > 0 Then
        For i = 0 To searchCount - 1
            freeDirHandle searchPath(i), False
        Next
        searchCount = 0                                 ' searchPath = NULL
        Erase searchPath
    End If
End Sub ' freeSearchPath

Private Function doDeinit() As Long
    closeFileHandleList True
    If PHYSFS_setWriteDir(vbNullString) = 0 Then
        PHYSFS_setErrorCode PHYSFS_ERR_FILES_STILL_OPEN
        Exit Function
    End If

    freeSearchPath
    ' (freeArchivers(); freeErrorStates();)

    baseDir = vbNullString
    prefDir = vbNullString

    allowSymLinks = False
    initialized = False

    doDeinit = 1
End Function ' doDeinit

Public Function PHYSFS_deinit() As Long
    If Not initialized Then
        PHYSFS_setErrorCode PHYSFS_ERR_NOT_INITIALIZED
        Exit Function
    End If
    PHYSFS_deinit = doDeinit()
End Function ' PHYSFS_deinit

Public Function PHYSFS_getDirSeparator() As String
    PHYSFS_getDirSeparator = "\"                        ' __PHYSFS_platformDirSeparator
End Function ' PHYSFS_getDirSeparator

' (vbNullString: NULL)
Public Function PHYSFS_getPrefDir(ByVal org As String, ByVal app As String) As String
    Dim statbuf As PHYSFS_Stat_, ptr As Long, path As String, buf As String

    PHYSFS_getPrefDir = vbNullString
    If Not initialized Then PHYSFS_setErrorCode PHYSFS_ERR_NOT_INITIALIZED: Exit Function
    If StrPtr(org) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function
    If Len(org) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function
    If StrPtr(app) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function
    If Len(app) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function

    ' prefDir = __PHYSFS_platformCalcPrefDir(org, app);
    buf = String$(260, 0)
    If SHGetFolderPathW(0, CSIDL_APPDATA Or CSIDL_FLAG_CREATE, 0, 0, StrPtr(buf)) < 0 Then    ' !SUCCEEDED
        prefDir = vbNullString
        PHYSFS_setErrorCode PHYSFS_ERR_OS_ERROR
        Exit Function
    End If
    prefDir = Left$(buf, InStr(buf, vbNullChar) - 1) & "\" & org & "\" & app & "\"

    ' (assert(strlen(prefDir) > 0); the final dirsep masked out for now)
    path = Left$(prefDir, Len(prefDir) - 1)

    If PHYSFS_platformStat__(path, statbuf, 1) = 0 Then
        ptr = InStr(path, "\")
        Do While ptr > 0
            PHYSFS_platformMkDir__ Left$(path, ptr - 1)
            ptr = InStr(ptr + 1, path, "\")
        Loop

        If PHYSFS_platformMkDir__(path) = 0 Then
            prefDir = vbNullString
            Exit Function
        End If
    End If

    ' (readd the final dirsep.)
    PHYSFS_getPrefDir = prefDir
End Function ' PHYSFS_getPrefDir

Public Function PHYSFS_getBaseDir() As String
    PHYSFS_getBaseDir = baseDir   ' this is calculated in PHYSFS_init()...
End Function ' PHYSFS_getBaseDir

' (vbNullString: NULL)
Public Function PHYSFS_getWriteDir() As String
    PHYSFS_getWriteDir = vbNullString
    If hasWriteDir Then
        PHYSFS_getWriteDir = writeDir.dirName
    End If
End Function ' PHYSFS_getWriteDir

' (newDir vbNullString: NULL)
Public Function PHYSFS_setWriteDir(ByVal newDir As String) As Long
    Dim retval As Long
    retval = 1

    If hasWriteDir Then
        If Not freeDirHandle(writeDir, True) Then Exit Function
        hasWriteDir = False
    End If

    If StrPtr(newDir) <> 0 Then
        hasWriteDir = createDirHandle(0, newDir, vbNullString, 1, writeDir)
        If hasWriteDir Then retval = 1 Else retval = 0
    End If

    PHYSFS_setWriteDir = retval
End Function ' PHYSFS_setWriteDir

' (VB) io: 0 (PHYSFS_mount's: no io)
Private Function doMount(ByVal io As Long, ByVal fname As String, ByVal mountPoint As String, ByVal appendToPath As Long) As Long
    Dim dh As DirHandle, i As Long

    If StrPtr(fname) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function

    If StrPtr(mountPoint) = 0 Then
        mountPoint = "/"
    End If

    For i = 0 To searchCount - 1
        ' already in search path?
        If StrComp(fname, searchPath(i).dirName, vbBinaryCompare) = 0 Then
            doMount = 1
            Exit Function
        End If
    Next

    If Not createDirHandle(io, fname, mountPoint, 0, dh) Then Exit Function

    ReDim Preserve searchPath(0 To searchCount)
    If appendToPath <> 0 Then
        searchPath(searchCount) = dh
    Else
        For i = searchCount To 1 Step -1
            searchPath(i) = searchPath(i - 1)
        Next
        searchPath(0) = dh
    End If
    searchCount = searchCount + 1

    doMount = 1
End Function ' doMount

' appendToPath: PHYSFS_APPEND (1) or PHYSFS_PREPEND (0); mountPoint vbNullString: NULL ("/")
Public Function PHYSFS_mount(ByVal newDir As String, ByVal mountPoint As String, ByVal appendToPath As Long) As Long
    If StrPtr(newDir) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function
    PHYSFS_mount = doMount(0, newDir, mountPoint, appendToPath)
End Function ' PHYSFS_mount

Public Function PHYSFS_unmount(ByVal oldDir As String) As Long
    Dim i As Long, j As Long

    If StrPtr(oldDir) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function

    For i = 0 To searchCount - 1
        If StrComp(searchPath(i).dirName, oldDir, vbBinaryCompare) = 0 Then
            If Not freeDirHandle(searchPath(i), False) Then Exit Function

            For j = i To searchCount - 2
                searchPath(j) = searchPath(j + 1)
            Next
            searchCount = searchCount - 1

            PHYSFS_unmount = 1
            Exit Function
        End If
    Next

    PHYSFS_setErrorCode PHYSFS_ERR_NOT_MOUNTED
End Function ' PHYSFS_unmount

' The search path's dirNames in order in out(0 ..); returns the count.
Public Function PHYSFS_getSearchPath(out() As String) As Long
    Dim i As Long
    If searchCount = 0 Then
        Erase out
        Exit Function
    End If
    ReDim out(0 To searchCount - 1)
    For i = 0 To searchCount - 1
        out(i) = searchPath(i).dirName
    Next
    PHYSFS_getSearchPath = searchCount
End Function ' PHYSFS_getSearchPath (PHYSFS_getSearchPathCallback)

Public Sub PHYSFS_permitSymbolicLinks(ByVal allow As Long)
    allowSymLinks = (allow <> 0)
End Sub ' PHYSFS_permitSymbolicLinks

'
' Verify that (fname) (in platform-independent notation), in relation
'  to (h) is secure. That means that each element of fname is checked
'  for symlinks (if they aren't permitted). This also allows for quick
'  rejection of files that exist outside an archive's mountpoint.
'
' With some exceptions (like PHYSFS_mkdir(), which builds multiple subdirs
'  at a time), you should always pass zero for "allowMissing" for efficiency.
'
' (fname) must point to an output from sanitizePlatformIndependentPath(),
'  since it will make sure that path names are in the right format for
'  passing certain checks. It will also do checks for "insecure" pathnames
'  like ".." which should be done once instead of once per archive. This also
'  gives us license to treat (fname) as scratch space in this function.
'
' Returns non-zero if string is safe, zero if there's a security issue.
'  PHYSFS_getLastError() will specify what was wrong. (*fname) will be
'  updated to point past any mount point elements so it is prepared to
'  be used with the archiver directly.
'
Private Function verifyPath(h As DirHandle, fname As String, ByVal allowMissing As Boolean) As Boolean
    Dim retval As Boolean, start As Long, end_ As Long, mntpntlen As Long, len_ As Long, statbuf As PHYSFS_Stat_, rc As Long
    retval = True

    If Len(fname) = 0 Then  ' quick rejection. (no root)
        verifyPath = True
        Exit Function
    End If

    ' !!! FIXME: This codeblock sucks.
    If StrPtr(h.mountPoint) <> 0 Then  ' NULL mountpoint means "/".
        mntpntlen = Len(h.mountPoint)
        len_ = Len(fname)
        ' (assert(mntpntlen > 1); /* root mount points should be NULL. */)
        ' not under the mountpoint, so skip this archive.
        If len_ < mntpntlen - 1 Then PHYSFS_setErrorCode PHYSFS_ERR_NOT_FOUND: Exit Function
        ' !!! FIXME: Case insensitive?
        If StrComp(Left$(h.mountPoint, mntpntlen - 1), Left$(fname, mntpntlen - 1), vbBinaryCompare) <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_NOT_FOUND: Exit Function
        If len_ > mntpntlen - 1 Then   ' corner case...
            If Mid$(fname, mntpntlen, 1) <> "/" Then PHYSFS_setErrorCode PHYSFS_ERR_NOT_FOUND: Exit Function
        End If
        fname = Mid$(fname, mntpntlen)  ' move to start of actual archive path.
        If Left$(fname, 1) = "/" Then
            fname = Mid$(fname, 2)
        End If
        retval = True  ' may be reset, below.
    End If

    ' (prepend the root directory, if any: no DirHandle has one)

    start = 1
    If Not allowSymLinks Then
        Do
            end_ = InStr(start, fname, "/")

            If end_ <> 0 Then                           ' (*end = '\0')
                rc = arcStat(h.funcs, h.opaque, Left$(fname, end_ - 1), statbuf)
            Else
                rc = arcStat(h.funcs, h.opaque, fname, statbuf)
            End If
            If rc <> 0 Then
                If statbuf.filetype = PHYSFS_FILETYPE_SYMLINK Then rc = 1 Else rc = 0
            ElseIf currentErrorCode() = PHYSFS_ERR_NOT_FOUND Then
                retval = False
            End If

            ' insecure path (has a disallowed symlink in it)?
            If rc <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_SYMLINK_FORBIDDEN: Exit Function

            ' break out early if path element is missing.
            If Not retval Then
                '
                ' We need to clear it if it's the last element of the path,
                '  since this might be a non-existant file we're opening
                '  for writing...
                '
                If (end_ = 0) Or allowMissing Then
                    retval = True
                End If
                Exit Do
            End If

            If end_ = 0 Then
                Exit Do
            End If

            start = end_ + 1
        Loop
    End If

    verifyPath = retval
End Function ' verifyPath

Private Function doMkdir(ByVal dname_ As String) As Long
    Dim dname As String, start As Long, end_ As Long, retval As Long, exists As Boolean, statbuf As PHYSFS_Stat_, rc As Long
    Dim part As String
    retval = 0
    exists = True   ' force existance check on first path element.

    ' (assert(h != NULL);)

    If Not sanitizePlatformIndependentPath(dname_, dname) Then Exit Function
    If Not verifyPath(writeDir, dname, True) Then Exit Function

    start = 1
    Do
        end_ = InStr(start, dname, "/")
        If end_ <> 0 Then
            part = Left$(dname, end_ - 1)
        Else
            part = dname
        End If

        ' only check for existance if all parent dirs existed, too...
        If exists Then
            rc = arcStat(writeDir.funcs, writeDir.opaque, part, statbuf)
            If (rc = 0) And (currentErrorCode() = PHYSFS_ERR_NOT_FOUND) Then
                exists = False
            End If
            ' verifyPath made sure that (dname) doesn't have symlinks if they aren't
            '   allowed, but it's possible the mounted writeDir itself has symlinks in it,
            '   (for example "/var" on iOS is a symlink, and the prefpath will be somewhere
            '   under that)...if we mounted that writeDir, we must allow those symlinks here
            '   unconditionally.
            retval = 0
            If rc <> 0 Then
                If (statbuf.filetype = PHYSFS_FILETYPE_DIRECTORY) Or (statbuf.filetype = PHYSFS_FILETYPE_SYMLINK) Then retval = 1
            End If
        End If

        If Not exists Then
            retval = arcMkdir(writeDir.funcs, writeDir.opaque, part)
        End If

        If retval = 0 Then
            Exit Do
        End If

        If end_ = 0 Then
            Exit Do
        End If

        start = end_ + 1
    Loop

    doMkdir = retval
End Function ' doMkdir

Public Function PHYSFS_mkdir(ByVal dname_ As String) As Long
    If StrPtr(dname_) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function

    If Not hasWriteDir Then PHYSFS_setErrorCode PHYSFS_ERR_NO_WRITE_DIR: Exit Function
    PHYSFS_mkdir = doMkdir(dname_)
End Function ' PHYSFS_mkdir

Private Function doDelete(ByVal fname_ As String) As Long
    Dim fname As String
    If Not sanitizePlatformIndependentPath(fname_, fname) Then Exit Function
    If Not verifyPath(writeDir, fname, False) Then Exit Function
    doDelete = arcRemove(writeDir.funcs, writeDir.opaque, fname)
End Function ' doDelete

Public Function PHYSFS_delete(ByVal fname_ As String) As Long
    If Not hasWriteDir Then PHYSFS_setErrorCode PHYSFS_ERR_NO_WRITE_DIR: Exit Function
    PHYSFS_delete = doDelete(fname_)
End Function ' PHYSFS_delete

' (VB) the index of the search path entry, or -1 (NULL)
Private Function getRealDirHandle(ByVal fname_ As String) As Long
    Dim retval As Long, fname As String, i As Long, arcfname As String, statbuf As PHYSFS_Stat_
    retval = -1

    If StrPtr(fname_) = 0 Then
        PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT
        getRealDirHandle = -1
        Exit Function
    End If

    If sanitizePlatformIndependentPath(fname_, fname) Then
        For i = 0 To searchCount - 1
            arcfname = fname
            If partOfMountPoint(searchPath(i), arcfname) Then
                retval = i
                Exit For
            ElseIf verifyPath(searchPath(i), arcfname, False) Then
                If arcStat(searchPath(i).funcs, searchPath(i).opaque, arcfname, statbuf) <> 0 Then
                    retval = i
                    Exit For
                End If
            End If
        Next
    End If

    getRealDirHandle = retval
End Function ' getRealDirHandle

' The search path entry (as mounted) that has fname; vbNullString if none (C: NULL).
Public Function PHYSFS_getRealDir(ByVal fname As String) As String
    Dim dh As Long
    dh = getRealDirHandle(fname)
    If dh >= 0 Then PHYSFS_getRealDir = searchPath(dh).dirName Else PHYSFS_getRealDir = vbNullString
End Function ' PHYSFS_getRealDir

' (VB) the list is enumList(0 .. pos - 1)
Private Function locateInStringList(ByVal str_ As String, pos As Long) As Boolean
    Dim len_ As Long, half_len As Long, lo As Long, middle As Long, cmp As Long
    len_ = pos
    lo = 0

    Do While len_ > 0
        half_len = len_ \ 2
        middle = lo + half_len
        cmp = StrComp(enumList(middle), str_, vbBinaryCompare)

        If cmp = 0 Then  ' it's in the list already.
            locateInStringList = True
            Exit Function
        ElseIf cmp > 0 Then
            len_ = half_len
        Else
            lo = middle + 1
            len_ = len_ - (half_len + 1)
        End If
    Loop

    pos = lo
End Function ' locateInStringList

Private Function enumFilesCallback(ByVal origdir As String, ByVal str_ As String) As Long
    Dim pos As Long, zero As Long

    '
    ' See if file is in the list already, and if not, insert it in there
    '  alphabetically...
    '
    pos = enumListSize
    If locateInStringList(str_, pos) Then
        enumFilesCallback = PHYSFS_ENUM_OK  ' already in the list, but keep going.
        Exit Function
    End If

    ' (allocator.Realloc(pecd->list, (pecd->size + 2) * sizeof (char *)))
    If enumListSize + 2 > UBound(enumList) + 1 Then
        ReDim Preserve enumList(0 To 2 * (enumListSize + 2) - 1)
    End If

    If pos <> enumListSize Then
        ' memmove(&pecd->list[pos+1], &pecd->list[pos], sizeof (char *) * ((pecd->size) - pos)): (VB) the
        ' string pointers move; the slot left at pos is emptied without freeing what moved on
        CopyMemory VarPtr(enumList(pos + 1)), VarPtr(enumList(pos)), 4 * (enumListSize - pos)
        CopyMemory VarPtr(enumList(pos)), VarPtr(zero), 4
    End If

    enumList(pos) = str_
    enumListSize = enumListSize + 1

    enumFilesCallback = PHYSFS_ENUM_OK
End Function ' enumFilesCallback

' The names in folder path, over the whole search path, sorted, in out(0 ..); returns the count, -1 on an error
' (C: NULL).
Public Function PHYSFS_enumerateFiles(ByVal path As String, out() As String) As Long
    Dim errcode As Long

    ' (EnumStringListCallbackData ecd; memset(&ecd, '\0', sizeof (ecd)); ecd.list = allocator.Malloc(...))
    ReDim enumList(0 To 0)
    enumListSize = 0
    enumErrcode = PHYSFS_ERR_OK
    If PHYSFS_enumerate(path, ENUMCB_FILES) = 0 Then
        errcode = currentErrorCode()
        Erase enumList
        enumListSize = 0
        If errcode = PHYSFS_ERR_APP_CALLBACK Then
            PHYSFS_setErrorCode enumErrcode
        End If
        Erase out
        PHYSFS_enumerateFiles = -1
        Exit Function
    End If

    If enumListSize = 0 Then
        Erase out
    Else
        ReDim Preserve enumList(0 To enumListSize - 1)
        out = enumList
    End If
    PHYSFS_enumerateFiles = enumListSize
    Erase enumList
    enumListSize = 0
End Function ' PHYSFS_enumerateFiles

'
' Broke out to seperate function so we can use stack allocation gratuitously.
'
Private Function enumerateFromMountPoint(i As DirHandle, ByVal arcfname As String, ByVal callback As Long, ByVal fname_ As String) As Long
    Dim retval As Long, len_ As Long, ptr As Long, end_ As Long, mountPoint As String
    len_ = Len(arcfname)

    mountPoint = i.mountPoint
    If len_ <> 0 Then ptr = len_ + 2 Else ptr = 1    ' mountPoint + ((len) ? len + 1 : 0)
    end_ = InStr(ptr, mountPoint, "/")
    ' (assert(end);  /* should always find a terminating '/'. */)
    retval = PHYSFS_EnumerateCallback_call(callback, fname_, Mid$(mountPoint, ptr, end_ - ptr))

    If retval = PHYSFS_ENUM_ERROR Then PHYSFS_setErrorCode PHYSFS_ERR_APP_CALLBACK
    enumerateFromMountPoint = retval
End Function ' enumerateFromMountPoint

Private Function enumCallbackFilterSymLinks(ByVal origdir As String, ByVal fname As String) As Long
    Dim statbuf As PHYSFS_Stat_, trimmedDir As String, path As String, retval As Long
    If Left$(filterArcfname, 1) = "/" Then trimmedDir = Mid$(filterArcfname, 2) Else trimmedDir = filterArcfname
    retval = PHYSFS_ENUM_OK

    If Len(trimmedDir) > 0 Then                         ' snprintf(path, slen, "%s%s%s", trimmedDir, *trimmedDir ? "/" : "", fname)
        path = trimmedDir & "/" & fname
    Else
        path = fname
    End If

    If arcStat(filterFuncs, filterOpaque, path, statbuf) = 0 Then
        filterErrcode = currentErrorCode()
        retval = PHYSFS_ENUM_ERROR
    Else
        ' Pass it on to the application if it's not a symlink.
        If statbuf.filetype <> PHYSFS_FILETYPE_SYMLINK Then
            retval = PHYSFS_EnumerateCallback_call(filterCallback, origdir, fname)
            If retval = PHYSFS_ENUM_ERROR Then
                filterErrcode = PHYSFS_ERR_APP_CALLBACK
            End If
        End If
    End If

    enumCallbackFilterSymLinks = retval
End Function ' enumCallbackFilterSymLinks

' (VB) cb(data, origdir, fname): the callbacks by number (their data: this module's variables)
Public Function PHYSFS_EnumerateCallback_call(ByVal cb As Long, ByVal origdir As String, ByVal fname As String) As Long
    Select Case cb
    Case ENUMCB_FILES: PHYSFS_EnumerateCallback_call = enumFilesCallback(origdir, fname)
    Case ENUMCB_FILTER_SYMLINKS: PHYSFS_EnumerateCallback_call = enumCallbackFilterSymLinks(origdir, fname)
    End Select
End Function

Private Function PHYSFS_enumerate(ByVal fn_ As String, ByVal cb As Long) As Long
    Dim retval As Long, fname As String, i As Long, arcfname As String, statbuf As PHYSFS_Stat_, blank As PHYSFS_Stat_
    retval = PHYSFS_ENUM_OK

    If StrPtr(fn_) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function
    ' (BAIL_IF(!cb, PHYSFS_ERR_INVALID_ARGUMENT, 0);)

    If Not sanitizePlatformIndependentPath(fn_, fname) Then
        retval = PHYSFS_ENUM_STOP
    Else
        ' (SymlinkFilterData filterdata: memset, its callback and callbackData: filterCallback)
        If Not allowSymLinks Then
            filterCallback = cb
        End If

        i = 0
        Do While (retval = PHYSFS_ENUM_OK) And (i < searchCount)
            arcfname = fname

            If partOfMountPoint(searchPath(i), arcfname) Then
                retval = enumerateFromMountPoint(searchPath(i), arcfname, cb, fn_)

            ElseIf verifyPath(searchPath(i), arcfname, False) Then
                statbuf = blank                         ' (VB) the C's is left as it was when stat fails
                If arcStat(searchPath(i).funcs, searchPath(i).opaque, arcfname, statbuf) = 0 Then
                    If currentErrorCode() = PHYSFS_ERR_NOT_FOUND Then
                        GoTo continue_  ' no such dir in this archive, skip it.
                    End If
                End If

                If statbuf.filetype <> PHYSFS_FILETYPE_DIRECTORY Then
                    GoTo continue_  ' not a directory in this archive, skip it.

                ElseIf (Not allowSymLinks) And arcSupportsSymlinks(searchPath(i).funcs) Then
                    filterFuncs = searchPath(i).funcs   ' filterdata.dirhandle = i
                    filterOpaque = searchPath(i).opaque
                    filterArcfname = arcfname
                    filterErrcode = PHYSFS_ERR_OK
                    retval = arcEnumerate(searchPath(i).funcs, searchPath(i).opaque, arcfname, ENUMCB_FILTER_SYMLINKS, fn_)
                    If retval = PHYSFS_ENUM_ERROR Then
                        If currentErrorCode() = PHYSFS_ERR_APP_CALLBACK Then
                            PHYSFS_setErrorCode filterErrcode
                        End If
                    End If
                Else
                    retval = arcEnumerate(searchPath(i).funcs, searchPath(i).opaque, arcfname, cb, fn_)
                End If
            End If
continue_:
            i = i + 1
        Loop
    End If

    If retval = PHYSFS_ENUM_ERROR Then PHYSFS_enumerate = 0 Else PHYSFS_enumerate = 1
End Function ' PHYSFS_enumerate

Public Function PHYSFS_exists(ByVal fname As String) As Long
    If getRealDirHandle(fname) >= 0 Then PHYSFS_exists = 1
End Function ' PHYSFS_exists

' (VB) forWriting: openWrite (False) or openAppend (True)
Private Function doOpenWrite(ByVal fname_ As String, ByVal appending As Boolean) As Long
    Dim fh As Long, fname As String, io As Long, arcfname As String

    If StrPtr(fname_) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function

    If Not hasWriteDir Then PHYSFS_setErrorCode PHYSFS_ERR_NO_WRITE_DIR: Exit Function

    If sanitizePlatformIndependentPath(fname_, fname) Then
        arcfname = fname
        If verifyPath(writeDir, arcfname, False) Then
            If appending Then
                io = arcOpenAppend(writeDir.funcs, writeDir.opaque, arcfname)
            Else
                io = arcOpenWrite(writeDir.funcs, writeDir.opaque, arcfname)
            End If

            If io <> 0 Then
                fh = newFileHandle()
                files(fh).io = io
                files(fh).forReading = False
                files(fh).dirHandle = writeDir.id
            End If
        End If
    End If

    doOpenWrite = fh
End Function ' doOpenWrite

Public Function PHYSFS_openWrite(ByVal filename As String) As Long
    PHYSFS_openWrite = doOpenWrite(filename, False)
End Function ' PHYSFS_openWrite

Public Function PHYSFS_openAppend(ByVal filename As String) As Long
    PHYSFS_openAppend = doOpenWrite(filename, True)
End Function ' PHYSFS_openAppend

Public Function PHYSFS_openRead(ByVal fname_ As String) As Long
    Dim fh As Long, fname As String, io As Long, i As Long, arcfname As String

    If StrPtr(fname_) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function

    If searchCount = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_NOT_FOUND: Exit Function

    If sanitizePlatformIndependentPath(fname_, fname) Then
        For i = 0 To searchCount - 1
            arcfname = fname
            If verifyPath(searchPath(i), arcfname, False) Then
                io = arcOpenRead(searchPath(i).funcs, searchPath(i).opaque, arcfname)
                If io <> 0 Then
                    Exit For
                End If
            End If
        Next

        If io <> 0 Then
            fh = newFileHandle()
            files(fh).io = io
            files(fh).forReading = True
            files(fh).dirHandle = searchPath(i).id
        End If
    End If

    PHYSFS_openRead = fh
End Function ' PHYSFS_openRead

' -1 == close failure. 0 == not found. 1 == success.
Private Function closeHandleInOpenList(ByVal forWritingList As Boolean, ByVal handle As Long) As Long
    Dim io As Long

    If Not ValidHandle(handle) Then Exit Function
    If files(handle).forReading = forWritingList Then Exit Function    ' handle is in this list?

    io = files(handle).io

    ' send our buffer to io...
    If Not files(handle).forReading Then
        If PHYSFS_flush(handle) = 0 Then
            closeHandleInOpenList = -1
            Exit Function

        ' ...then have io send it to the disk...
        ElseIf PHYSFS_Io_flush(io) = 0 Then
            closeHandleInOpenList = -1
            Exit Function
        End If
    End If

    ' ...then close the underlying file.
    PHYSFS_Io_destroy io

    ' (if (tmp != NULL) allocator.Free(tmp);: no buffer)

    files(handle).inUse = False                         ' allocator.Free(handle)
    files(handle).io = 0
    closeHandleInOpenList = 1
End Function ' closeHandleInOpenList

Public Function PHYSFS_close(ByVal handle As Long) As Long
    Dim rc As Long

    ' -1 == close failure. 0 == not found. 1 == success.
    rc = closeHandleInOpenList(False, handle)
    If rc = -1 Then Exit Function
    If rc = 0 Then
        rc = closeHandleInOpenList(True, handle)
        If rc = -1 Then Exit Function
    End If

    If rc = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function
    PHYSFS_close = 1
End Function ' PHYSFS_close

' Reads length bytes into buffer(ofs ..). The count read, or -1.
Public Function PHYSFS_readBytes(ByVal handle As Long, buffer() As Byte, ByVal ofs As Long, ByVal length As Long) As Long
    If Not ValidHandle(handle) Then                     ' (VB) the C does not check its handle
        PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT
        PHYSFS_readBytes = -1
        Exit Function
    End If

    If length < 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: PHYSFS_readBytes = -1: Exit Function   ' (_len > maxlen)
    If Not files(handle).forReading Then PHYSFS_setErrorCode PHYSFS_ERR_OPEN_FOR_WRITING: PHYSFS_readBytes = -1: Exit Function
    If length = 0 Then Exit Function
    ' (if (fh->buffer) return doBufferedRead(fh, buffer, len);: no buffer)

    PHYSFS_readBytes = PHYSFS_Io_read(files(handle).io, buffer, ofs, length)
End Function ' PHYSFS_readBytes

' Writes length bytes from buffer(ofs ..). The count written, or -1.
Public Function PHYSFS_writeBytes(ByVal handle As Long, buffer() As Byte, ByVal ofs As Long, ByVal length As Long) As Long
    If Not ValidHandle(handle) Then                     ' (VB) the C does not check its handle
        PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT
        PHYSFS_writeBytes = -1
        Exit Function
    End If

    If length < 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: PHYSFS_writeBytes = -1: Exit Function  ' (_len > maxlen)
    If files(handle).forReading Then PHYSFS_setErrorCode PHYSFS_ERR_OPEN_FOR_READING: PHYSFS_writeBytes = -1: Exit Function
    If length = 0 Then Exit Function
    ' (if (fh->buffer) return doBufferedWrite(handle, buffer, len);: no buffer)

    PHYSFS_writeBytes = PHYSFS_Io_write(files(handle).io, buffer, ofs, length)
End Function ' PHYSFS_writeBytes

Public Function PHYSFS_eof(ByVal handle As Long) As Long
    Dim io As Long, pos As Double, len_ As Double
    If Not ValidHandle(handle) Then Exit Function       ' (VB)

    If Not files(handle).forReading Then    ' never EOF on files opened for write/append.
        Exit Function
    End If

    ' can't be eof if buffer isn't empty: (no buffer)
    ' check the Io.
    io = files(handle).io
    pos = PHYSFS_Io_tell(io)
    len_ = PHYSFS_Io_length(io)
    If (pos < 0) Or (len_ < 0) Then
        Exit Function   ' beats me.
    End If
    If pos >= len_ Then PHYSFS_eof = 1
End Function ' PHYSFS_eof

Public Function PHYSFS_tell(ByVal handle As Long) As Double
    Dim pos As Double
    If Not ValidHandle(handle) Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: PHYSFS_tell = -1: Exit Function  ' (VB)
    pos = PHYSFS_Io_tell(files(handle).io)
    ' (forReading ? (pos - fh->buffill) + fh->bufpos : (pos + fh->buffill): no buffer)
    PHYSFS_tell = pos
End Function ' PHYSFS_tell

Public Function PHYSFS_seek(ByVal handle As Long, ByVal pos As Double) As Long
    If Not ValidHandle(handle) Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function  ' (VB)
    If PHYSFS_flush(handle) = 0 Then Exit Function

    ' (if (fh->buffer && fh->forReading) ...: no buffer)

    ' we have to fall back to a 'raw' seek.
    PHYSFS_seek = PHYSFS_Io_seek(files(handle).io, pos)
End Function ' PHYSFS_seek

' The file's length in bytes, or -1.
Public Function PHYSFS_fileLength(ByVal handle As Long) As Double
    If Not ValidHandle(handle) Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: PHYSFS_fileLength = -1: Exit Function  ' (VB)
    PHYSFS_fileLength = PHYSFS_Io_length(files(handle).io)
End Function ' PHYSFS_filelength

Public Function PHYSFS_flush(ByVal handle As Long) As Long
    If Not ValidHandle(handle) Then Exit Function       ' (VB)
    ' open for read or buffer empty are successful no-ops: (no buffer, so always)
    PHYSFS_flush = 1
End Function ' PHYSFS_flush

' PHYSFS_stat: 1 and the fields, or 0. Times are seconds since 1970.
Public Function PHYSFS_stat(ByVal fname_ As String, filesize As Double, modtime As Double, filetype As Long, readonly_ As Long) As Long
    Dim stat As PHYSFS_Stat_, retval As Long, fname As String, i As Long, exists As Boolean, arcfname As String

    If StrPtr(fname_) = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_INVALID_ARGUMENT: Exit Function

    ' set some sane defaults...
    stat.filesize = -1
    stat.modtime = -1
    stat.createtime = -1
    stat.accesstime = -1
    stat.filetype = PHYSFS_FILETYPE_OTHER
    stat.readonly_ = 1

    If sanitizePlatformIndependentPath(fname_, fname) Then
        If Len(fname) = 0 Then
            stat.filetype = PHYSFS_FILETYPE_DIRECTORY
            If hasWriteDir Then stat.readonly_ = 0 Else stat.readonly_ = 1   ' Writeable if we have a writeDir
            retval = 1
        Else
            i = 0
            Do While (i < searchCount) And (Not exists)
                arcfname = fname
                exists = partOfMountPoint(searchPath(i), arcfname)
                If exists Then
                    stat.filetype = PHYSFS_FILETYPE_DIRECTORY
                    stat.readonly_ = 1
                    retval = 1
                ElseIf verifyPath(searchPath(i), arcfname, False) Then
                    retval = arcStat(searchPath(i).funcs, searchPath(i).opaque, arcfname, stat)
                    If (retval <> 0) Or (currentErrorCode() <> PHYSFS_ERR_NOT_FOUND) Then
                        exists = True
                    End If
                End If
                i = i + 1
            Loop
        End If
    End If

    filesize = stat.filesize
    modtime = stat.modtime
    filetype = stat.filetype
    readonly_ = stat.readonly_
    PHYSFS_stat = retval
End Function ' PHYSFS_stat

' physfs_ext.h: WZ_PHYSFS_isDirectory
Public Function WZ_PHYSFS_isDirectory(ByVal fname As String) As Long
    Dim filesize As Double, modtime As Double, filetype As Long, ro As Long
    If PHYSFS_stat(fname, filesize, modtime, filetype, ro) = 0 Then Exit Function
    If filetype = PHYSFS_FILETYPE_DIRECTORY Then WZ_PHYSFS_isDirectory = 1
End Function

'--------------------------------------------------------------------------
' physfs_byteorder.c: 32-bit values in big-endian order; nonzero on success. A read that gets fewer than 4
' bytes fails and leaves value as it was.

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
