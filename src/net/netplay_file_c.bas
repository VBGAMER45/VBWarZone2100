Attribute VB_Name = "netplay_file_c"
' lib/netplay/netplay.cpp (4 of 4): the file transfers (the maps and mods a joining player lacks):
' NETsendFile, validateReceivedFile, markAsDownloadedFile, NETrecvFile, NETgetDownloadProgress and the list
' of downloads (DownloadingWzFiles: a Collection of WZFile objects; a PLAYER's wzFiles likewise).
' validateReceivedFile hashes the file 4 MiB at a time with libsodium's crypto_hash_sha256_init / update / final
' (src/shared/sodium).
Option Explicit

Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const FILE_SHARE_DELETE As Long = 4
Private Const OPEN_ALWAYS As Long = 4
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1

Private Const MAX_FILE_TRANSFER_PACKET As Long = 4096

Private DownloadingWzFiles As New Collection        ' static std::vector<WZFile> DownloadingWzFiles

' const std::vector<WZFile>& NET_getDownloadingWzFiles(): its size and its items (0-based)
Public Function NET_getDownloadingWzFiles_size() As Long
    NET_getDownloadingWzFiles_size = DownloadingWzFiles.Count
End Function

Public Function NET_getDownloadingWzFiles_item(ByVal i As Long) As WZFile
    Set NET_getDownloadingWzFiles_item = DownloadingWzFiles.Item(i + 1)
End Function

' (VB) any_of(DownloadingWzFiles, file.hash == hash): the hash's 32 bytes
Public Function NET_getDownloadingWzFiles_hasHash(hashBytes() As Byte) As Boolean
    Dim hash As Sha256, h As Sha256, f As WZFile
    Sha256_fromBytes hash, hashBytes
    For Each f In DownloadingWzFiles
        h = f.hash()
        If Sha256_equals(h, hash) Then
            NET_getDownloadingWzFiles_hasHash = True
            Exit Function
        End If
    Next
    NET_getDownloadingWzFiles_hasHash = False
End Function

Public Sub NET_addDownloadingWZFile(ByVal newFile As WZFile)
    DownloadingWzFiles.Add newFile
End Sub

Public Sub NET_clearDownloadingWZFiles()
    Dim file As WZFile
    ' if we were in a middle of transferring a file, then close the file handle
    For Each file In DownloadingWzFiles
        debug_ 362, LOG_NET, "NET_clearDownloadingWZFiles", "closing aborted file"
        file.closeFile
        If file.filename = "" Then ASSERT_FAILURE "!file.filename.empty()", "lib/netplay/netplay.cpp:364", "NET_clearDownloadingWZFiles", "filename must not be empty"
        PHYSFS_delete file.filename         ' delete incomplete (map) file
    Next
    Set DownloadingWzFiles = New Collection
End Sub

' (VB) DownloadingWzFiles.clear(): without closing anything (NET_InitPlayers)
Public Sub NET_clearDownloadingWzFilesList()
    Set DownloadingWzFiles = New Collection
End Sub

' ////////////////////////////////////////////////////////////////////////
' File Transfer programs.
' Send file. It returns % of file sent when 100 it's complete. Call until it returns 100.
'  @TODO: more error checking (?) different file types (?)
'          Maybe should close file handle, and seek each time?
'
'  @NOTE: MAX_FILE_TRANSFER_PACKET is set to 4k per packet since 7*4 = 28K which is pretty
'         much our limit.  Don't screw with that without having a bigger buffer!
'         NET_BUFFER_SIZE is at 256k.  (also remember text chat, plus all the other cruff)
Public Function NETsendFile(ByVal file As WZFile, ByVal player As Long) As Long
    Dim inBuff(0 To MAX_FILE_TRANSFER_PACKET - 1) As Byte, readBytesResult As Double, bytesToRead As Long, w As MessageWriter
    Dim hashBytes() As Byte
    If Not NetPlay.isHost Then ASSERT_FAILURE "NetPlay.isHost", "lib/netplay/netplay.cpp:2975", "NETsendFile", "Trying to send a file and we are not the host!": NETsendFile = 100: Exit Function
    If file.handle() = 0 Then ASSERT_FAILURE "file.handle() != nullptr", "lib/netplay/netplay.cpp:2976", "NETsendFile", "Null file handle": NETsendFile = 100: Exit Function

    ' read some bytes.
    readBytesResult = WZ_PHYSFS_readBytes(file.handle(), inBuff, 0, MAX_FILE_TRANSFER_PACKET)
    If readBytesResult < 0 Then
        ASSERT_FAILURE "readBytesResult >= 0", "lib/netplay/netplay.cpp:2985", "NETsendFile", "Error reading file."
        file.closeFile
        NETsendFile = 100
        Exit Function
    End If
    bytesToRead = readBytesResult

    hashBytes = file.hashBytes()
    Set w = NETbeginEncode(NETnetQueue(player), NET_FILE_PAYLOAD)
    NETbin_w w, hashBytes, 32
    NETuint32_t_w w, file.size      ' total bytes in this file. (we don't support 64bit yet)
    NETuint32_t_w w, file.pos       ' start byte
    NETuint32_t_w w, bytesToRead    ' bytes in this packet
    NETbin_w w, inBuff, bytesToRead
    NETend_w w

    file.pos = file.pos + bytesToRead  ' update position!
    If file.pos = file.size Then
        file.closeFile ' We are done sending to this client.
    End If

    NETsendFile = Int(U32(file.pos) * 100# / U32(file.size))
End Function

Public Function validateReceivedFile(ByVal file As WZFile) As Boolean
    Dim fileHandle As Long, actualFileSize64 As Double, actualFileSize As Long, actualFileHash As Sha256, fileChunkBuffer() As Byte
    Dim lengthRead As Double, expected As Sha256, state As crypto_hash_sha256_state, bufferSize As Long
    fileHandle = PHYSFS_openRead(file.filename)
    If fileHandle = 0 Then ASSERT_FAILURE "fileHandle != nullptr", "lib/netplay/netplay.cpp:3011", "validateReceivedFile", "Could not open downloaded file %s for reading: %s", file.filename, WZ_PHYSFS_getLastError(): validateReceivedFile = False: Exit Function

    actualFileSize64 = PHYSFS_fileLength(fileHandle)
    If actualFileSize64 < 0 Then
        debug_ 3016, LOG_ERROR, "validateReceivedFile", "Failed to determine file size of the downloaded file!"
        PHYSFS_close fileHandle
        validateReceivedFile = False
        Exit Function
    End If
    If actualFileSize64 > 2147483647# Then
        debug_ 3022, LOG_ERROR, "validateReceivedFile", "Downloaded file is too large!"
        PHYSFS_close fileHandle
        validateReceivedFile = False
        Exit Function
    End If

    actualFileSize = actualFileSize64
    If actualFileSize <> file.size Then
        debug_ 3030, LOG_ERROR, "validateReceivedFile", "Downloaded map unexpected size! Got %u, expected %u!", actualFileSize, file.size
        PHYSFS_close fileHandle
        validateReceivedFile = False
        Exit Function
    End If

    ' verify actual downloaded file hash matches expected hash

    crypto_hash_sha256_init state
    bufferSize = actualFileSize
    If bufferSize > 4 * 1024& * 1024& Then bufferSize = 4 * 1024& * 1024&
    If bufferSize > 0 Then ReDim fileChunkBuffer(0 To bufferSize - 1)
    Do
        lengthRead = 0
        If bufferSize > 0 Then lengthRead = WZ_PHYSFS_readBytes(fileHandle, fileChunkBuffer, 0, bufferSize)
        If lengthRead <> bufferSize Then
            If lengthRead < 0 Or PHYSFS_eof(fileHandle) = 0 Then
                ' did not read expected amount, but did not reach end of file - some other error reading the file occurred
                debug_ 3050, LOG_ERROR, "validateReceivedFile", "Failed to read downloaded file: %s", WZ_PHYSFS_getLastError()
                PHYSFS_close fileHandle
                validateReceivedFile = False
                Exit Function
            End If
        End If
        crypto_hash_sha256_update state, fileChunkBuffer, 0, lengthRead
    Loop While lengthRead = bufferSize And bufferSize > 0
    crypto_hash_sha256_final state, actualFileHash.bytes, 0
    Erase fileChunkBuffer

    expected = file.hash()
    If Not Sha256_equals(actualFileHash, expected) Then
        debug_ 3062, LOG_ERROR, "validateReceivedFile", "Downloaded file hash (%s) does not match requested file hash (%s)", Sha256_toString(actualFileHash), Sha256_toString(expected)
        PHYSFS_close fileHandle
        validateReceivedFile = False
        Exit Function
    End If

    PHYSFS_close fileHandle
    validateReceivedFile = True
End Function

Public Function markAsDownloadedFile(ByVal filename As String) As Boolean
    Dim current_writeDir As String, fullFilePath As String, fullFilePathUTF16 As String, hStream As Long
    Dim zone() As Byte, dwNumberOfBytesWritten As Long
    ' Files are written to the PhysFS writeDir
    current_writeDir = PHYSFS_getWriteDir()
    If current_writeDir = "" Then ASSERT_FAILURE "current_writeDir != nullptr", "lib/netplay/netplay.cpp:3075", "markAsDownloadedFile", "Failed to get PhysFS writeDir: %s", WZ_PHYSFS_getLastError()
    fullFilePath = current_writeDir & PHYSFS_getDirSeparator() & filename

    ' On Windows:
    '   - Create the Alternate Data Stream required to set the Internet Zone identifier
    ' (a VB String is already UTF-16)
    fullFilePathUTF16 = fullFilePath & ":Zone.Identifier"

    hStream = CreateFileW(StrPtr(fullFilePathUTF16), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hStream = INVALID_HANDLE_VALUE Then
        ' Failed to open stream
        debug_ 3103, LOG_ERROR, "markAsDownloadedFile", "Could not open stream; failed with error %lu: %s" & vbLf, Err.LastDllError, fullFilePath
        markAsDownloadedFile = False
        Exit Function
    End If

    ' Set it to "downloaded from the Internet Zone" (ZoneId 3)
    zone = StrConv("[ZoneTransfer]" & vbCrLf & "ZoneId=3" & vbCrLf, vbFromUnicode)
    If WriteFile(hStream, zone(0), UBound(zone) + 1, dwNumberOfBytesWritten, 0) = 0 Then
        debug_ 3112, LOG_ERROR, "markAsDownloadedFile", "Failed to write to stream with error %lu: %s" & vbLf, Err.LastDllError, fullFilePath
        CloseHandle hStream
        markAsDownloadedFile = False
        Exit Function
    End If

    FlushFileBuffers hStream
    CloseHandle hStream

    markAsDownloadedFile = True
End Function

' (VB) std::find_if over DownloadingWzFiles by hash: the 1-based index, 0 when none
Private Function findDownloadingFile(hash As Sha256) As Long
    Dim i As Long, h As Sha256, f As WZFile
    For i = 1 To DownloadingWzFiles.Count
        Set f = DownloadingWzFiles.Item(i)      ' (VB) typed: hash() is a Friend member
        h = f.hash()
        If Sha256_equals(h, hash) Then
            findDownloadingFile = i
            Exit Function
        End If
    Next
    findDownloadingFile = 0
End Function

Private Sub sendCancelFileDownload(hash As Sha256)
    Dim w As MessageWriter
    Set w = NETbeginEncode(NETnetQueue(NetPlay.hostPlayer), NET_FILE_CANCELLED)
    NETbin_w w, hash.bytes, 32
    NETend_w w
End Sub

Private Sub terminateFileDownload(ByVal fileIdx As Long)
    Dim file As WZFile, h As Sha256
    Set file = DownloadingWzFiles.Item(fileIdx)
    If Not file.closeFile() Then
        debug_ 3172, LOG_ERROR, "NETrecvFile", "Could not close file handle after trying to terminate download: %s", WZ_PHYSFS_getLastError()
    End If
    PHYSFS_delete file.filename
    h = file.hash()
    sendCancelFileDownload h
    DownloadingWzFiles.remove fileIdx
End Sub

' recv file. it returns % of the file so far recvd.
Public Function NETrecvFile(ByVal queue As Long) As Long
    Dim hash As Sha256, size As Long, pos As Long, bytesToRead As Long, buf() As Byte
    Dim r As MessageReader, fileIdx As Long, file As WZFile, newPos As Long, hb() As Byte
    Sha256_setZero hash
    size = 0
    pos = 0
    bytesToRead = 0
    ReDim buf(0 To MAX_FILE_TRANSFER_PACKET - 1)       ' uint8_t buf[MAX_FILE_TRANSFER_PACKET]

    'read incoming bytes.
    Set r = NETbeginDecode(queue, NET_FILE_PAYLOAD)
    ReDim hb(0 To 31)
    NETbin_r r, hb, 32
    Sha256_fromBytes hash, hb
    NETuint32_t_r r, size  ' total bytes in this file. (we don't support 64bit yet)
    NETuint32_t_r r, pos  ' start byte
    NETuint32_t_r r, bytesToRead  ' bytes in this packet
    If ULT(MAX_FILE_TRANSFER_PACKET, bytesToRead) Then ASSERT_FAILURE "bytesToRead <= sizeof(buf)", "lib/netplay/netplay.cpp:3148", "NETrecvFile", "Bad value.": NETrecvFile = 100: Exit Function
    NETbin_r r, buf, bytesToRead
    NETend_r r

    debug_ 3152, LOG_NET, "NETrecvFile", "New file position is %u", pos

    fileIdx = findDownloadingFile(hash)

    If fileIdx = 0 Then
        debug_ 3164, LOG_WARNING, "NETrecvFile", "Receiving file data we didn't request."
        sendCancelFileDownload hash
        NETrecvFile = 100
        Exit Function
    End If
    Set file = DownloadingWzFiles.Item(fileIdx)

    'sanity checks
    If file.size <> size Then
        If file.size = 0 Then
            ' host does not send the file size until the first recvFile packet
            file.size = size
        Else
            ' host sent a different file size for this file with this chunk (vs the first chunk)
            ' this should not happen!
            debug_ 3191, LOG_ERROR, "NETrecvFile", "Host sent a different file size for this file; (original size: %u, new size: %u)", file.size, size
            terminateFileDownload fileIdx ' 'file' is now an invalidated iterator.
            NETrecvFile = 100
            Exit Function
        End If
    End If

    If ULT(MAX_NET_TRANSFERRABLE_FILE_SIZE, size) Then
        ' file size is too large
        debug_ 3200, LOG_ERROR, "NETrecvFile", "Downloaded filesize is too large; (size: %u)", size
        terminateFileDownload fileIdx ' 'file' is now an invalidated iterator.
        NETrecvFile = 100
        Exit Function
    End If

    If PHYSFS_tell(file.handle()) <> U32(pos) Then
        ' actual position in file does not equal the expected position in the file (sent by the host)
        debug_ 3208, LOG_ERROR, "NETrecvFile", "Invalid file position in downloaded file; (desired: %u)", pos
        terminateFileDownload fileIdx ' 'file' is now an invalidated iterator.
        NETrecvFile = 100
        Exit Function
    End If

    ' Write packet to the file.
    WZ_PHYSFS_writeBytes file.handle(), buf, 0, bytesToRead

    newPos = pos + bytesToRead
    file.pos = newPos

    If Not ULT(newPos, size) Then  ' last packet
        If Not file.closeFile() Then
            debug_ 3223, LOG_ERROR, "NETrecvFile", "Could not close file handle after trying to save map: %s", WZ_PHYSFS_getLastError()
        End If

        If Not validateReceivedFile(file) Then
            ' Delete the (invalid) downloaded file
            PHYSFS_delete file.filename
        Else
            ' Attach Quarantine / "downloaded" file attribute to file
            markAsDownloadedFile file.filename
        End If

        DownloadingWzFiles.remove fileIdx
    End If
    ' 'file' may now be an invalidated iterator.

    'return the percentage count
    If size <> 0 Then
        NETrecvFile = Int(U32(newPos) * 100# / U32(size))
        Exit Function
    End If
    debug_ 3246, LOG_ERROR, "NETrecvFile", "Received 0 byte file from host?"
    NETrecvFile = 100     ' file is nullbyte, so we are done.
End Function

Public Function NETgetDownloadProgress(ByVal player As Long) As Long
    Dim files As Collection, progress As Double, file As WZFile, p As Double, sz As Double
    If player = selectedPlayer Then
        Set files = DownloadingWzFiles      ' Check our own download progress.
    Else
        Set files = NetPlay.players(player).wzFiles  ' Check their download progress (currently only works if we are the host).
    End If

    If files Is Nothing Then ASSERT_FAILURE "files != nullptr", "lib/netplay/netplay.cpp:3256", "NETgetDownloadProgress", "Uninitialized wzFiles? (Player %u)", player: NETgetDownloadProgress = 100: Exit Function

    progress = 100
    For Each file In files
        sz = U32(file.size)
        If sz < 1 Then sz = 1
        p = Int(U32(file.pos) * 100# / sz)
        If p < progress Then progress = p
    Next
    NETgetDownloadProgress = progress
End Function

' (VB) NetPlay.players[i].wzFiles: its size, an item (0-based), push_back, clear, erase by hash
Public Function NetPlay_players_wzFiles_size(ByVal i As Long) As Long
    NetPlay_players_wzFiles_size = NetPlay.players(i).wzFiles.Count
End Function

Public Function NetPlay_players_wzFiles_item(ByVal i As Long, ByVal n As Long) As WZFile
    Set NetPlay_players_wzFiles_item = NetPlay.players(i).wzFiles.Item(n + 1)
End Function

Public Sub NetPlay_players_wzFiles_push_back(ByVal i As Long, ByVal file As WZFile)
    NetPlay.players(i).wzFiles.Add file
End Sub

Public Sub NetPlay_players_wzFiles_clear(ByVal i As Long)
    Set NetPlay.players(i).wzFiles = New Collection
End Sub

' wzFiles->erase(std::remove_if(..., [&hash](WZFile const &file) { return file.hash == hash; }), wzFiles->end())
Public Sub NetPlay_players_wzFiles_eraseHash(ByVal i As Long, hashBytes() As Byte)
    Dim n As Long, hash As Sha256, h As Sha256, f As WZFile
    Sha256_fromBytes hash, hashBytes
    For n = NetPlay.players(i).wzFiles.Count To 1 Step -1
        Set f = NetPlay.players(i).wzFiles.Item(n)
        h = f.hash()
        If Sha256_equals(h, hash) Then NetPlay.players(i).wzFiles.remove n
    Next
End Sub

' std::any_of(wzFiles->begin(), wzFiles->end(), [&hash](WZFile const &file) { return file.hash == hash; })
Public Function NetPlay_players_wzFiles_hasHash(ByVal i As Long, hashBytes() As Byte) As Boolean
    Dim n As Long, hash As Sha256, h As Sha256, f As WZFile
    Sha256_fromBytes hash, hashBytes
    For n = 1 To NetPlay.players(i).wzFiles.Count
        Set f = NetPlay.players(i).wzFiles.Item(n)
        h = f.hash()
        If Sha256_equals(h, hash) Then
            NetPlay_players_wzFiles_hasHash = True
            Exit Function
        End If
    Next
    NetPlay_players_wzFiles_hasHash = False
End Function

' wzFiles->emplace_back(pFileHandle, filename, hash, fileSize)
Public Sub NetPlay_players_wzFiles_emplace_back(ByVal i As Long, ByVal pFileHandle As Long, ByVal filename As String, hashBytes() As Byte, ByVal fileSize As Long)
    NetPlay.players(i).wzFiles.Add WZFile_new(pFileHandle, filename, hashBytes, fileSize)
End Sub

' wzFiles->erase(std::remove_if(..., [](WZFile const &file) { return file.handle() == nullptr; }), wzFiles->end())
Public Sub NetPlay_players_wzFiles_eraseClosed(ByVal i As Long)
    Dim n As Long, f As WZFile
    For n = NetPlay.players(i).wzFiles.Count To 1 Step -1
        Set f = NetPlay.players(i).wzFiles.Item(n)
        If f.handle() = 0 Then NetPlay.players(i).wzFiles.remove n
    Next
End Sub

' NETsendFile((*NetPlay.players[player].wzFiles)[fileIdx], player)
Public Function NETsendFile_idx(ByVal player As Long, ByVal fileIdx As Long) As Long
    NETsendFile_idx = NETsendFile(NetPlay.players(player).wzFiles.Item(fileIdx + 1), player)
End Function

' (VB) Sha256's bytes from b(0 .. 31)
Public Sub Sha256_fromBytes(h As Sha256, b() As Byte)
    Dim k As Long
    For k = 0 To 31
        h.bytes(k) = b(k)
    Next
End Sub
