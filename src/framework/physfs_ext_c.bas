Attribute VB_Name = "physfs_ext_c"
' lib/framework/physfs_ext.cpp, physfs_ext.h (ported in part). The enumeration
' callback (enumFunc, which can stop the enumeration) is the caller's loop over
' the names returned here, and so is WZ_PHYSFS_cleanupOldFilesInFolder's
' deleteFileFunction (see there). Ported: WZ_PHYSFS_readBytes,
' WZ_PHYSFS_writeBytes (the header's inline wrappers), WZ_PHYSFS_enumerateFiles,
' WZ_PHYSFS_enumerateFolders, WZ_PHYSFS_cleanupOldFilesInFolder (the overload
' with an extension and the one it calls), filenameEndWithExtension.
' WZ_PHYSFS_SETBUFFER is nothing: this PhysFS writes unbuffered, and a buffer
' does not change the bytes written. Missing: WZ_PHYSFS_enumerateFilesEx,
' WZ_PHYSFS_createPlatformIndependentPath, the other helpers.
Option Explicit

' physfs_ext.h: PHYSFS_readBytes's count (-1 on an error), as a PHYSFS_sint64
Public Function WZ_PHYSFS_readBytes(ByVal handle As Long, buf() As Byte, ByVal first As Long, ByVal size As Long) As Double
    WZ_PHYSFS_readBytes = PHYSFS_readBytes(handle, buf, first, size)
End Function

' physfs_ext.h: PHYSFS_writeBytes's count (-1 on an error), as a PHYSFS_sint64
Public Function WZ_PHYSFS_writeBytes(ByVal handle As Long, buf() As Byte, ByVal first As Long, ByVal size As Long) As Double
    WZ_PHYSFS_writeBytes = PHYSFS_writeBytes(handle, buf, first, size)
End Function

' physfs_ext.h: the modification time PHYSFS_stat gives
Private Function WZ_PHYSFS_getLastModTime(ByVal filename As String) As Double
    Dim filesize As Double, modtime As Double, filetype As Long, readonly_ As Long
    PHYSFS_stat filename, filesize, modtime, filetype, readonly_
    WZ_PHYSFS_getLastModTime = modtime
End Function

Public Function filenameEndWithExtension(ByVal filename As String, ByVal extension As String) As Boolean
    Dim filenameLength As Long, extensionLength As Long
    If StrPtr(filename) = 0 Then
        filenameEndWithExtension = False
        Exit Function
    End If

    filenameLength = Len(filename)
    extensionLength = Len(extension)
    If extensionLength < filenameLength Then
        filenameEndWithExtension = StrComp(Right$(filename, extensionLength), extension, vbBinaryCompare) = 0
    End If
End Function

' int WZ_PHYSFS_cleanupOldFilesInFolder(path, extension, fileLimit, deleteFileFunction), with the overload taking
' CleanupFileEnumFilterFunctions that it calls (a file name filter: filenameEndWithExtension; no last-modified filter).
' (VB) deleteFileFunction (std::function) is the caller's: the deleting loop is the caller's too, over what this gives.
' It returns -1 (an assert failed) or 0 (no more files than fileLimit: the C++ returns 0 there), or 1 with nfiles (the
' matching files' count) and oldest(0 .. count - 1), the paths "path/name" of the matching files, oldest first
' (std::multiset by modification time: equal times keep their enumeration order). The C++ then runs, returning
' numFilesDeleted:
'     while (nfiles > fileLimit && !fileTimes.empty()) { if (deleteFileFunction(oldest)) { ++numFilesDeleted; --nfiles; }
'                                                       erase the oldest; if (fileLimit < 0) break; }
Public Function WZ_PHYSFS_cleanupOldFilesInFolder(ByVal path As String, ByVal extension As String, ByVal fileLimit As Long, nfiles As Long, oldest() As String, count As Long) As Long
    Dim files() As String, n As Long, i As Long, k As Long, j As Long, savefile_ As String
    Dim names() As String, times() As Double, savetime As Double
    count = 0
    nfiles = 0
    If Not (StrPtr(extension) <> 0) Then ASSERT_FAILURE "extension != nullptr", "lib/framework/physfs_ext.cpp:213", "WZ_PHYSFS_cleanupOldFilesInFolder", "Null extension": WZ_PHYSFS_cleanupOldFilesInFolder = -1: Exit Function
    If Not (StrPtr(path) <> 0) Then ASSERT_FAILURE "path != nullptr", "lib/framework/physfs_ext.cpp:223", "WZ_PHYSFS_cleanupOldFilesInFolder", "Null path": WZ_PHYSFS_cleanupOldFilesInFolder = -1: Exit Function
    n = PHYSFS_enumerateFiles(path, files)
    If Not (n >= 0) Then ASSERT_FAILURE "files", "lib/framework/physfs_ext.cpp:227", "WZ_PHYSFS_cleanupOldFilesInFolder", "PHYSFS_enumerateFiles(""%s"") failed: %s", path, WZ_PHYSFS_getLastError(): WZ_PHYSFS_cleanupOldFilesInFolder = -1: Exit Function
    For i = 0 To n - 1
        If Not filenameEndWithExtension(files(i), extension) Then
            GoTo continue_count
        End If
        nfiles = nfiles + 1
continue_count:
    Next
    If nfiles <= fileLimit Or nfiles <= 0 Then
        WZ_PHYSFS_cleanupOldFilesInFolder = 0
        Exit Function
    End If

    ' too many files
    debug_ 244, LOG_SAVE, "WZ_PHYSFS_cleanupOldFilesInFolder", "found %i matching files in %s, limit is %i", nfiles, path, fileLimit

    ' build a sorted list of file + save time
    ReDim names(0 To n - 1)
    ReDim times(0 To n - 1)
    For i = 0 To n - 1
        If Not filenameEndWithExtension(files(i), extension) Then
            GoTo continue_times
        End If
        ' Gather save-time
        savefile_ = path & "/" & files(i)
        savetime = WZ_PHYSFS_getLastModTime(savefile_)
        ' (multiset insert: after the equal times already there)
        k = count
        Do While k > 0
            If Not (savetime < times(k - 1)) Then Exit Do
            k = k - 1
        Loop
        For j = count To k + 1 Step -1
            names(j) = names(j - 1)
            times(j) = times(j - 1)
        Next
        names(k) = files(i)
        times(k) = savetime
        count = count + 1
continue_times:
    Next

    ' now delete the oldest (the caller's loop)
    If count > 0 Then
        ReDim oldest(0 To count - 1)
        For i = 0 To count - 1
            oldest(i) = path & "/" & names(i)
        Next
    End If
    WZ_PHYSFS_cleanupOldFilesInFolder = 1
End Function

' The names in dir (files and folders), in files(0 .. count - 1); false when the enumeration fails
Public Function WZ_PHYSFS_enumerateFiles(ByVal dirPath As String, files() As String, count As Long) As Boolean
    count = PHYSFS_enumerateFiles(dirPath, files)
    If count < 0 Then
        debug_ 34, LOG_ERROR, "WZ_PHYSFS_enumerateFiles", "PHYSFS_enumerateFiles(""%s"") failed: %s", dirPath, WZ_PHYSFS_getLastError()
        count = 0
        Exit Function
    End If
    WZ_PHYSFS_enumerateFiles = True
End Function

' The folders in dir, in folders(0 .. count - 1); false when the enumeration fails
Public Function WZ_PHYSFS_enumerateFolders(ByVal dirPath As String, folders() As String, count As Long) As Boolean
    Dim files() As String, n As Long, i As Long, baseDir As String, isDir As String
    count = 0
    n = PHYSFS_enumerateFiles(dirPath, files)
    If n < 0 Then
        debug_ 117, LOG_ERROR, "WZ_PHYSFS_enumerateFolders", "PHYSFS_enumerateFiles(""%s"") failed: %s", dirPath, WZ_PHYSFS_getLastError()
        Exit Function
    End If
    baseDir = dirPath
    If Not strEndsWith(baseDir, "/") Then
        baseDir = baseDir & "/"
    End If
    For i = 0 To n - 1
        isDir = baseDir
        isDir = isDir & files(i)
        If WZ_PHYSFS_isDirectory(isDir) <> 0 Then
            If count = 0 Then ReDim folders(0 To n - 1)
            folders(count) = files(i)
            count = count + 1
        End If
    Next
    WZ_PHYSFS_enumerateFolders = True
End Function
