Attribute VB_Name = "netreplay_c"
' lib/netplay/netreplay.cpp, netreplay.h: replay files. A replay is the magic
' number, the settings (JSON: the format version, the netcode version and the
' replay options handler's game options) with their size, the embedded map
' (its data version and size, then its bytes), every game message the game
' processed (its player's byte, then its raw data), REPLAY_ENDED from the host,
' then the end-of-game JSON with its size before and after it.
' The writer thread (replaySaveThreadFunc on a
' moodycamel::BlockingReaderWriterQueue of SerializedNetMessagesBuffer chunks)
' is plain writes: while "the thread" runs, a chunk queued is written at once
' and the empty end chunk stops it; without it (a multiplayer game's large
' buffer: the settings are written last, before the chunks), the queue keeps
' the chunks until NETreplaySaveStop writes them. Joining is nothing; the
' file's bytes are the same. WZ_PHYSFS_SETBUFFER is nothing (see physfs_ext_c).
' WZ_PHYSFS_cleanupOldFilesInFolder's deleting loop, with this file's lambda
' (deleteOldReplayFile), is here (see physfs_ext_c). The onFail lambda is
' replayLoadFail. The lambdas log under their enclosing function's name (GCC's
' __FUNCTION__ in a lambda is "operator()").
Option Explicit

Private Const SIZE_MAX_D As Double = 18446744073709551615#  ' std::numeric_limits<size_t>::max() (64-bit)

Private replaySaveHandle As Long            ' PHYSFS_file *
Private replayLoadHandle As Long            ' PHYSFS_file *

Private Const magicReplayNumber As Long = &H575A7270  ' "WZrp"
Private Const currentReplayFormatVer As Long = 3
Private Const minReplayFormatVerSupported As Long = 3
Private Const DefaultReplayBufferSize As Long = 32768
Private Const MaxReplayBufferSize As Long = 2& * 1024 * 1024

' typedef std::vector<uint8_t> SerializedNetMessagesBuffer: the queue's chunks
Private queueChunks As Collection           ' serializedBufferWriteQueue (each a Byte array; an empty chunk is an empty String)
Private queuedSaveSettings As nlohmann_json
Private latestWriteBuffer() As Byte         ' SerializedNetMessagesBuffer latestWriteBuffer: (0 .. latestWriteBufferSize - 1)
Private latestWriteBufferSize As Long
Private minBufferSizeToQueue As Double      ' size_t (DefaultReplayBufferSize until a replay is started)
Private saveThread As Boolean               ' WZ_THREAD *saveThread: True while the writer "thread" runs

' (VB) serializedBufferWriteQueue.enqueue(chunk): the running writer thread writes it at once (and stops at the empty
' chunk); else the queue keeps it
Private Sub enqueueChunk(chunk() As Byte, ByVal chunkSize As Long)
    Dim item As Variant
    If saveThread Then
        If chunkSize = 0 Then
            ' end chunk - the thread is done
            saveThread = False
            Exit Sub
        End If
        WZ_PHYSFS_writeBytes replaySaveHandle, chunk, 0, chunkSize
        Exit Sub
    End If
    If queueChunks Is Nothing Then Set queueChunks = New Collection
    If chunkSize = 0 Then
        queueChunks.Add ""
    Else
        item = chunk
        If UBound(chunk) <> chunkSize - 1 Then
            ReDim Preserve chunk(0 To chunkSize - 1)
            item = chunk
        End If
        queueChunks.Add item
    End If
End Sub

' This function is run in its own thread! Do not call any non-threadsafe functions!
' (VB) run where the thread would be, on the chunks the queue kept
Private Function replaySaveThreadFunc(ByVal pSaveHandle As Long) As Long
    Dim item As Variant, data() As Byte
    If pSaveHandle = 0 Then
        replaySaveThreadFunc = 1
        Exit Function
    End If
    If queueChunks Is Nothing Then Set queueChunks = New Collection
    Do While queueChunks.Count > 0
        item = queueChunks(1)
        queueChunks.Remove 1
        If VarType(item) = vbString Then
            ' end chunk - we're done
            Exit Do
        End If
        data = item
        WZ_PHYSFS_writeBytes pSaveHandle, data, 0, UBound(data) + 1
    Loop
    replaySaveThreadFunc = 0
End Function

' (VB) a String as its UTF-8 bytes (std::string's data()), their count returned
Private Function stringBytes(ByVal s As String, b() As Byte) As Long
    stringBytes = utf8_encode(s, b)
End Function

Private Function NETreplaySaveWritePreamble(ByVal settings As nlohmann_json, ByVal optionsHandler As ReplayOptionsHandler) As Boolean
    Dim data As String, bytes() As Byte, n As Long, dataVersion As Long, mapBinaryData() As Byte, mapBinaryDataSize As Long
    If replaySaveHandle = 0 Then
        NETreplaySaveWritePreamble = False
        Exit Function
    End If

    data = settings.dump(-1, " ", False, error_handler_t_replace)
    n = stringBytes(data, bytes)
    PHYSFS_writeUBE32 replaySaveHandle, n
    If n > 0 Then WZ_PHYSFS_writeBytes replaySaveHandle, bytes, 0, n

    ' Save extra map data (if present)
    dataVersion = 0
    mapBinaryDataSize = 0
    If Not optionsHandler.saveMap(dataVersion, mapBinaryData, mapBinaryDataSize) Then
        ' Failed to save map data - just empty it out for now
        mapBinaryDataSize = 0
    End If
    PHYSFS_writeUBE32 replaySaveHandle, dataVersion
    ' (SIZE_MAX > UINT32_MAX: the official builds are 64-bit)
    If Not (U32(mapBinaryDataSize) <= 4294967295#) Then ASSERT_FAILURE "embeddedMapData.mapBinaryData.size() <= static_cast<size_t>(std::numeric_limits<uint32_t>::max())", "lib/netplay/netreplay.cpp:108", "NETreplaySaveWritePreamble", "Embedded map data is way too big": NETreplaySaveWritePreamble = False: Exit Function
    PHYSFS_writeUBE32 replaySaveHandle, mapBinaryDataSize
    If mapBinaryDataSize > 0 Then
        WZ_PHYSFS_writeBytes replaySaveHandle, mapBinaryData, 0, mapBinaryDataSize
    End If

    NETreplaySaveWritePreamble = True
End Function

' NETreplaySaveStart's lambda for WZ_PHYSFS_cleanupOldFilesInFolder
Private Function deleteOldReplayFile(ByVal fileName As String) As Boolean
    If PHYSFS_delete(fileName) = 0 Then
        debug_ 137, LOG_ERROR, "NETreplaySaveStart", "Failed to delete old replay file: %s", fileName
        deleteOldReplayFile = False
        Exit Function
    End If
    deleteOldReplayFile = True
End Function

Public Function NETreplaySaveStart(ByVal subdir As String, ByVal optionsHandler As ReplayOptionsHandler, ByVal maxReplaysSaved As Long, ByVal appendPlayerToFilename As Boolean) As String
    Dim replayFullDir As String, aclock As Double, newtime As tm, filename As String, settings As nlohmann_json
    Dim gameOptions As nlohmann_json, desiredBufferSize As Double, r As Long, nfiles As Long, oldest() As String
    Dim count As Long, k As Long, fileLimit As Long
    If NETisReplay() Then
        ' Have already loaded and will be running a replay - don't bother saving another
        debug_ 124, LOG_WZ, "NETreplaySaveStart", "Replay loaded - skip recording of new replay"
        NETreplaySaveStart = ""
        Exit Function
    End If

    If Not (Len(subdir) <> 0) Then ASSERT_FAILURE "!subdir.empty()", "lib/netplay/netreplay.cpp:128", "NETreplaySaveStart", "Must provide a valid subdir": NETreplaySaveStart = "": Exit Function

    If maxReplaysSaved > 0 Then
        ' clean up old replay files
        replayFullDir = "replay/" & subdir
        ' WZ_PHYSFS_cleanupOldFilesInFolder(replayFullDir, ".wzrp", maxReplaysSaved - 1, deleteOldReplayFile): its
        ' deleting loop (physfs_ext.cpp) over the files it gives
        fileLimit = maxReplaysSaved - 1
        r = WZ_PHYSFS_cleanupOldFilesInFolder(replayFullDir, ".wzrp", fileLimit, nfiles, oldest, count)
        If r > 0 Then
            k = 0
            Do While nfiles > fileLimit And k < count
                If deleteOldReplayFile(oldest(k)) Then
                    nfiles = nfiles - 1
                End If
                k = k + 1
                If fileLimit < 0 Then
                    Exit Do
                End If
            Loop
        End If
    End If

    aclock = time_()                    ' Get time in seconds
    localtime_s newtime, aclock         ' Convert time to struct

    If appendPlayerToFilename Then
        filename = va("replay/%s/%04d%02d%02d_%02d%02d%02d_%s_p%u.wzrp", subdir, newtime.tm_year + 1900, newtime.tm_mon + 1, newtime.tm_mday, newtime.tm_hour, newtime.tm_min, newtime.tm_sec, subdir, selectedPlayer)
    Else
        filename = va("replay/%s/%04d%02d%02d_%02d%02d%02d_%s.wzrp", subdir, newtime.tm_year + 1900, newtime.tm_mon + 1, newtime.tm_mday, newtime.tm_hour, newtime.tm_min, newtime.tm_sec, subdir)
    End If
    replaySaveHandle = PHYSFS_openWrite(filename)  ' open the file
    If replaySaveHandle = 0 Then
        debug_ 160, LOG_ERROR, "NETreplaySaveStart", "Could not create replay file %s: %s", filename, WZ_PHYSFS_getLastError()
        NETreplaySaveStart = ""
        Exit Function
    End If

    ' WZ_PHYSFS_SETBUFFER(replaySaveHandle, 1024 * 32): nothing (see physfs_ext_c)

    PHYSFS_writeSBE32 replaySaveHandle, magicReplayNumber

    ' Save map name or map data and game settings and list of players in game and stuff.
    Set settings = json_object()

    ' Save "replay file format version"
    settings.set_item "replayFormatVer", json_value_unsigned(currentReplayFormatVer)

    ' Save Netcode version
    settings.set_item "major", json_value_unsigned(U32(NETGetMajorVersion()))
    settings.set_item "minor", json_value_unsigned(U32(NETGetMinorVersion()))

    ' Save desired info from optionsHandler
    Set gameOptions = json_object()
    optionsHandler.saveOptions gameOptions
    settings.set_item "gameOptions", gameOptions

    ' determine best buffer size
    desiredBufferSize = optionsHandler.desiredBufferSize()
    If desiredBufferSize = 0 Then
        ' use default
        minBufferSizeToQueue = DefaultReplayBufferSize
    ElseIf desiredBufferSize >= MaxReplayBufferSize Then
        minBufferSizeToQueue = MaxReplayBufferSize
    Else
        minBufferSizeToQueue = desiredBufferSize
    End If

    debug_ 199, LOG_INFO, "NETreplaySaveStart", "Started writing replay file ""%s"".", filename

    ' Create a background thread and hand off all responsibility for writing to the file handle to it
    If saveThread Then ASSERT_FAILURE "saveThread == nullptr", "lib/netplay/netreplay.cpp:202", "NETreplaySaveStart", "Failed to release prior thread"
    ' latestWriteBuffer.reserve(minBufferSizeToQueue)
    If desiredBufferSize <> SIZE_MAX_D Then
        ' Write the preamble immediately (settings, etc)
        NETreplaySaveWritePreamble settings, optionsHandler

        ' use a background thread
        Set queueChunks = New Collection
        saveThread = True
    Else
        ' Do not immediately write settings out - instead, queue them for later writing
        Set queuedSaveSettings = settings

        ' don't use a background thread
        Set queueChunks = New Collection
        saveThread = False
    End If

    NETreplaySaveStart = filename
End Function

Public Function NETreplaySaveStop(ByVal optionsHandler As ReplayOptionsHandler) As Boolean
    Dim replayEndedMessage As NetMessage, endOfGameInfo As nlohmann_json, data As String, bytes() As Byte, n As Long
    Dim emptyChunk() As Byte, hadThread As Boolean
    If replaySaveHandle = 0 Then
        NETreplaySaveStop = False
        Exit Function
    End If

    ' v2: Append the "REPLAY_ENDED" message (from hostPlayer)
    Set replayEndedMessage = NetMessageBuilder_new(REPLAY_ENDED, 0).build()
    appendToLatestWriteBuffer NetPlay.hostPlayer
    replayEndedMessage.rawDataAppendToVector latestWriteBuffer, latestWriteBufferSize

    ' Queue the last chunk for writing
    If latestWriteBufferSize <> 0 Then
        enqueueChunk latestWriteBuffer, latestWriteBufferSize
    End If

    ' Then push one empty chunk to signify "we're done!"
    latestWriteBufferSize = 0
    Erase latestWriteBuffer
    hadThread = saveThread
    enqueueChunk emptyChunk, 0

    ' Wait for writing thread to finish
    If hadThread Then
        ' wzThreadJoin(saveThread): the chunks were written as they came
        saveThread = False
    Else
        ' update the queued settings (ex. might have revealed player identities in a blind game)
        optionsHandler.optionsUpdatePlayerInfo queuedSaveSettings.at("gameOptions")

        ' write the preamble
        NETreplaySaveWritePreamble queuedSaveSettings, optionsHandler

        ' do the writing now on the main thread
        replaySaveThreadFunc replaySaveHandle
    End If

    ' v2: Write the "end of game info" chunk
    ' (this is JSON that is preceded *and* followed by its size - so it should be possible to seek to the end of the file, read the last uint32_t, and then back up and grab the JSON without processing the whole file)
    Set endOfGameInfo = json_object()
    endOfGameInfo.set_item "gameTimeElapsed", json_value_unsigned(U32(netGame.gameTime()))
    ' FUTURE TODO: Could save things like the game results / winners + losers

    data = endOfGameInfo.dump()
    n = stringBytes(data, bytes)
    PHYSFS_writeUBE32 replaySaveHandle, n
    If n > 0 Then WZ_PHYSFS_writeBytes replaySaveHandle, bytes, 0, n
    PHYSFS_writeUBE32 replaySaveHandle, n  ' should also end with the json size for easy reading from end of file

    If PHYSFS_close(replaySaveHandle) = 0 Then
        debug_ 278, LOG_ERROR, "NETreplaySaveStop", "Could not close replay file: %s", WZ_PHYSFS_getLastError()
        NETreplaySaveStop = False
        Exit Function
    End If
    replaySaveHandle = 0

    NETreplaySaveStop = True
End Function

' (VB) latestWriteBuffer.push_back(b)
Private Sub appendToLatestWriteBuffer(ByVal b As Long)
    If latestWriteBufferSize = 0 Then
        ReDim latestWriteBuffer(0 To 1023)
    ElseIf latestWriteBufferSize > UBound(latestWriteBuffer) Then
        ReDim Preserve latestWriteBuffer(0 To latestWriteBufferSize * 2 - 1)
    End If
    latestWriteBuffer(latestWriteBufferSize) = b And &HFF&
    latestWriteBufferSize = latestWriteBufferSize + 1
End Sub

Public Sub NETreplaySaveNetMessage(ByVal message As NetMessage, ByVal player As Long)
    If replaySaveHandle = 0 Then
        Exit Sub
    End If

    If message.type_() > GAME_MIN_TYPE And message.type_() < GAME_MAX_TYPE Then
        appendToLatestWriteBuffer player
        message.rawDataAppendToVector latestWriteBuffer, latestWriteBufferSize

        If latestWriteBufferSize >= minBufferSizeToQueue Then
            enqueueChunk latestWriteBuffer, latestWriteBufferSize
            Erase latestWriteBuffer
            latestWriteBufferSize = 0
            ' latestWriteBuffer.reserve(minBufferSizeToQueue)
        End If
    End If
End Sub

' NETreplayLoadStart's onFail lambda
Private Function replayLoadFail(ByVal filename As String, ByVal reason As String) As Boolean
    debug_ 310, LOG_ERROR, "NETreplayLoadStart", "Could not load replay file %s: %s", filename, reason
    If replayLoadHandle <> 0 Then
        PHYSFS_close replayLoadHandle
        replayLoadHandle = 0
    End If
    replayLoadFail = False
End Function

Public Function NETreplayLoadStart(ByVal filename As String, ByVal optionsHandler As ReplayOptionsHandler, output_replayFormatVer As Long) As Boolean
    Dim replayNumber As Long, dataSize As Long, data() As Byte, dataRead As Double, settings As nlohmann_json
    Dim replayFormatVer As Long, mismatchVersionDescription As String, failLogStr As String
    Dim replay_netcodeMajor As Long, replay_netcodeMinor As Long, dataVersion As Long, binaryDataSize As Long
    Dim mapBinaryData() As Byte, mapBinaryDataSize As Long, binaryDataRead As Double, filePos As Double, afterMapDataPos As Double
    Dim parseError As String

    replayLoadHandle = PHYSFS_openRead(filename)
    If replayLoadHandle = 0 Then
        NETreplayLoadStart = replayLoadFail(filename, WZ_PHYSFS_getLastError())
        Exit Function
    End If

    replayNumber = 0
    PHYSFS_readSBE32 replayLoadHandle, replayNumber
    If replayNumber <> magicReplayNumber Then
        NETreplayLoadStart = replayLoadFail(filename, "bad header")
        Exit Function
    End If

    dataSize = 0
    PHYSFS_readUBE32 replayLoadHandle, dataSize
    dataRead = 0
    If dataSize <> 0 Then
        ReDim data(0 To dataSize - 1)
        dataRead = WZ_PHYSFS_readBytes(replayLoadHandle, data, 0, dataSize)
    End If
    If dataRead <> U32(dataSize) Then
        NETreplayLoadStart = replayLoadFail(filename, "truncated header")
        Exit Function
    End If

    ' Restore map name or map data and game settings and list of players in game and stuff.
    On Error GoTo json_failed
    If dataSize <> 0 Then Set settings = json_parse_bytes(data, 0, dataSize) Else Set settings = json_parse("")

    replayFormatVer = settings.at("replayFormatVer").get_uint()
    output_replayFormatVer = replayFormatVer
    If ULT(currentReplayFormatVer, replayFormatVer) Then
        mismatchVersionDescription = "The replay file format is newer than this version of Warzone 2100 can support."
        mismatchVersionDescription = mismatchVersionDescription & vbLf & vbLf
        mismatchVersionDescription = mismatchVersionDescription & va("Replay Format Version: %u", replayFormatVer)
        wzDisplayDialog Dialog_Error, "Replay File Format Unsupported", mismatchVersionDescription

        failLogStr = "Replay file format is newer than this version of Warzone 2100 can support: " & Format$(U32(replayFormatVer))
        NETreplayLoadStart = replayLoadFail(filename, failLogStr)
        Exit Function
    End If

    If ULT(replayFormatVer, minReplayFormatVerSupported) Then
        mismatchVersionDescription = "The replay file format is older than this version of Warzone 2100 can support."
        mismatchVersionDescription = mismatchVersionDescription & vbLf & vbLf
        mismatchVersionDescription = mismatchVersionDescription & va("Replay Format Version: %u", replayFormatVer)
        wzDisplayDialog Dialog_Error, "Replay File Format Unsupported", mismatchVersionDescription

        failLogStr = "Replay file format is older than this version of Warzone 2100 can support: " & Format$(U32(replayFormatVer))
        NETreplayLoadStart = replayLoadFail(filename, failLogStr)
        Exit Function
    End If

    replay_netcodeMajor = settings.at("major").get_uint()
    replay_netcodeMinor = settings.at("minor").get_uint()
    If Not NETisCorrectVersion(replay_netcodeMajor, replay_netcodeMinor) Then
        debug_ 375, LOG_INFO, "NETreplayLoadStart", "NetCode Version mismatch: (replay file: 0x%x, 0x%x) - (current: 0x%x, 0x%x)", replay_netcodeMajor, replay_netcodeMinor, NETGetMajorVersion(), NETGetMinorVersion()
        ' do not immediately fail out - restoreOptions handles displaying a nicer warning popup
    End If

    ' ReplayOptionsHandler::EmbeddedMapData embeddedMapData;
    dataVersion = 0
    mapBinaryDataSize = 0
    If ULT(1, replayFormatVer) Then
        PHYSFS_readUBE32 replayLoadHandle, dataVersion
        binaryDataSize = 0
        PHYSFS_readUBE32 replayLoadHandle, binaryDataSize
        If binaryDataSize <> 0 Then
            If U32(binaryDataSize) <= optionsHandler.maximumEmbeddedMapBufferSize() Then
                ReDim mapBinaryData(0 To binaryDataSize - 1)
                mapBinaryDataSize = binaryDataSize
                binaryDataRead = WZ_PHYSFS_readBytes(replayLoadHandle, mapBinaryData, 0, mapBinaryDataSize)
                If binaryDataRead <> U32(binaryDataSize) Then
                    NETreplayLoadStart = replayLoadFail(filename, "truncated embedded map data")
                    Exit Function
                End If
            Else
                ' don't even bother trying to load this - it's too big
                ' just attempt to skip to where it claims the map data ends
                filePos = PHYSFS_tell(replayLoadHandle)
                If filePos < 0 Then
                    NETreplayLoadStart = replayLoadFail(filename, "error getting current file position")
                    Exit Function
                End If
                afterMapDataPos = filePos + U32(binaryDataSize)
                If PHYSFS_seek(replayLoadHandle, afterMapDataPos) = 0 Then
                    NETreplayLoadStart = replayLoadFail(filename, "failed to seek after map data")
                    Exit Function
                End If
            End If
        End If
    End If

    ' Load game options using optionsHandler
    If Not optionsHandler.restoreOptions(settings.at("gameOptions"), dataVersion, mapBinaryData, mapBinaryDataSize, replay_netcodeMajor, replay_netcodeMinor) Then
        NETreplayLoadStart = replayLoadFail(filename, "invalid options")
        Exit Function
    End If
    On Error GoTo 0

    debug_ 427, LOG_INFO, "NETreplayLoadStart", "Started reading replay file ""%s"".", filename
    NETreplayLoadStart = True
    Exit Function

json_failed:
    ' catch (const std::exception& e): Failed to parse or find a key in the json
    If Err.Source <> "nlohmann::json" Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    parseError = "Error parsing info JSON (""" & Err.Description & """)"
    On Error GoTo 0
    NETreplayLoadStart = replayLoadFail(filename, parseError)
End Function

Public Function NETreplayLoadNetMessage(message As NetMessage, player As Long) As Boolean
    Dim b1(0 To 0) As Byte, typ As Long, b(0 To 1) As Byte, rd As Boolean, length As Long, replayData() As Byte
    Dim messageRead As Double, msgBuilder As NetMessageBuilder
    If replayLoadHandle = 0 Then
        NETreplayLoadNetMessage = False
        Exit Function
    End If

    If WZ_PHYSFS_readBytes(replayLoadHandle, b1, 0, 1) = 1 Then player = b1(0)

    typ = 0
    If WZ_PHYSFS_readBytes(replayLoadHandle, b1, 0, 1) = 1 Then typ = b1(0)

    rd = WZ_PHYSFS_readBytes(replayLoadHandle, b, 0, 2) <> 0
    If Not rd Then
        NETreplayLoadNetMessage = False
        Exit Function
    End If
    ' Load payload length from uint16_t (network byte order) starting at the second byte of the data buffer.
    length = CLng(b(0)) * 256 + b(1)

    messageRead = 0
    If length > 0 Then
        ReDim replayData(0 To length - 1)
        messageRead = WZ_PHYSFS_readBytes(replayLoadHandle, replayData, 0, length)
    End If

    If messageRead <> length Then
        NETreplayLoadNetMessage = False
        Exit Function
    End If

    Set msgBuilder = NetMessageBuilder_new(typ, length)
    If length > 0 Then msgBuilder.append_bytes replayData, 0, length

    Set message = msgBuilder.build()

    NETreplayLoadNetMessage = (message.type_() > GAME_MIN_TYPE And message.type_() < GAME_MAX_TYPE) Or message.type_() = REPLAY_ENDED
End Function

Public Function NETreplayLoadStop() As Boolean
    If replayLoadHandle = 0 Then
        NETreplayLoadStop = False
        Exit Function
    End If

    If PHYSFS_close(replayLoadHandle) = 0 Then
        debug_ 478, LOG_ERROR, "NETreplayLoadStop", "Could not close replay file: %s", WZ_PHYSFS_getLastError()
        NETreplayLoadStop = False
        Exit Function
    End If
    replayLoadHandle = 0

    NETreplayLoadStop = True
End Function
