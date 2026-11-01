Attribute VB_Name = "sync_debug_c"
' lib/netplay/sync_debug.cpp, sync_debug.h: the sync log. Each game tick's
' syncDebug entries and their CRC are kept for the last MAX_SYNC_HISTORY
' ticks; dumpLocalDebugSyncLog prints one, and the verbose mode writes each
' tick's to a file. checkDebugSync compares a player's CRC with ours and on a
' mismatch dumps our log (DesyncLogOutputter) and broadcasts it
' (sendDebugSync); recvDebugSync writes the logs other players send.
Option Explicit

Private verboseDebugOutputAllSyncLogsUntilGameTime As Long     ' uint32_t

' struct SyncDebugLog (the type gets the underscore: the array syncDebugLog
' is used far more). The C++ keeps the strings, value changes and int
' lists in three vectors, the strings' characters and the lists' ints in
' two more, and the order in log; here each entry's fields are at its index
' (logType: "s", "v" or "i"), and the ints are in one array as in C++.
Private Type SyncDebugLog_
    logType() As Byte               ' std::vector<char> log
    entryFunction() As String       ' SyncDebugEntry::function
    entryString() As String         ' SyncDebugString's chars, SyncDebugValueChange::variableName, SyncDebugIntList::string
    entryValue() As Long            ' SyncDebugValueChange::newValue, SyncDebugIntList::numInts
    entryId() As Long               ' SyncDebugValueChange::id
    numEntries As Long              ' log.size()
    capacity As Long                ' (VB) the entry arrays' size
    ints() As Long                  ' std::vector<int> ints
    numInts As Long
    intsCapacity As Long            ' (VB) ints()'s size
    time_ As Long                   ' uint32_t
    crc As Long                     ' uint32_t
End Type

Private Const MAX_LEN_LOG_LINE As Long = 512    ' From debug.c - no use printing something longer.
Private Const MAX_SYNC_HISTORY As Long = 12

Private syncDebugNext As Long                   ' unsigned
Private syncDebugLog(0 To MAX_SYNC_HISTORY - 1) As SyncDebugLog_
Private syncDebugExtraGameTime As Long          ' uint32_t
Private syncDebugExtraCrc As Long               ' uint32_t

Private syncDebugNumDumps As Long               ' uint32_t

' 5 MiB seems to be a reasonable size limit (at the moment)
Private Const MaxDesyncLogSize As Long = 5000000

Private Const kDefaultDebugSyncBufferSize As Long = 4000000
Private Const kMaxDebugSyncBufferSize As Long = MaxDesyncLogSize
Private targetDebugSyncTmpBufSize As Long       ' size_t (kDefaultDebugSyncBufferSize; 0 until resetSyncDebug)

Private crcBuf() As Byte                        ' (VB) the bytes crcUpdateString and crcUpdateInt pass to wz_crc_update
Private crcBufSize As Long

Private Const MaxPlayerSyncDebugDumps As Long = 2
Private Const NetMessageOverheadMax As Long = 1 + 4         ' 1 byte for message type, 4 bytes for max uint32_t
Private Const MAX_DESYNC_LOG_TRANSFER_PACKET As Long = 32768 - NetMessageOverheadMax - 8    ' MaxMsgSize - overhead

' class DesyncLogOutputter: PlayerDesyncLogOutput's outputs (a
' std::unordered_map<uint32_t, LogOutputter> by time) as an array of
' LogOutputters with their times
Private Type LogOutputter_
    time_ As Long                               ' (VB) the map's key
    fp As Long                                  ' PHYSFS_file* fp = nullptr
    noFurtherWrites As Boolean
    writtenSize As Long                         ' size_t
End Type

Private Type PlayerDesyncLogOutput_
    outputs() As LogOutputter_                  ' DesyncOutputs outputs
    numOutputs As Long
End Type

' static DesyncLogOutputter playerDesyncLogOutput: its std::array<PlayerDesyncLogOutput, MAX_GAMEQUEUE_SLOTS> players
Private playerDesyncLogOutput(0 To MAX_GAMEQUEUE_SLOTS - 1) As PlayerDesyncLogOutput_

'--------------------------------------------------------------------------
' (VB) wz::crc_update over a C string and its terminating NUL: the
' characters' bytes (UTF-8, as the C++ strings are)
Private Function crcUpdateString(ByVal crc As Long, ByVal s As String) As Long
    Dim b() As Byte, n As Long, i As Long, ascii As Boolean
    n = Len(s)
    If n + 1 > crcBufSize Then
        crcBufSize = (n + 1) * 2
        ReDim crcBuf(0 To crcBufSize - 1)
    End If
    If n > 0 Then
        b = s                                   ' UTF-16LE
        ascii = True
        For i = 0 To n - 1
            If b(2 * i + 1) <> 0 Or b(2 * i) >= 128 Then
                ascii = False
                Exit For
            End If
            crcBuf(i) = b(2 * i)
        Next
        If Not ascii Then
            n = utf8_encode(s, b)
            If n + 1 > crcBufSize Then
                crcBufSize = (n + 1) * 2
                ReDim crcBuf(0 To crcBufSize - 1)
            End If
            For i = 0 To n - 1
                crcBuf(i) = b(i)
            Next
        End If
    End If
    crcBuf(n) = 0
    crcUpdateString = wz_crc_update(crc, crcBuf, 0, n + 1)
End Function

' (VB) wz::crc_update over wz_htonl(value)'s 4 bytes
Private Function crcUpdateInt(ByVal crc As Long, ByVal value As Long) As Long
    crcUpdateInt = wz_crc_update_be32(crc, value)
End Function

'--------------------------------------------------------------------------
' SyncDebugLog's members

Private Sub SyncDebugLog_clear(l As SyncDebugLog_)
    l.numEntries = 0
    l.time_ = 0
    l.crc = wz_crc_init()
    'printf("Freeing %d strings, %d valueChanges, %d intLists, %d chars, %d ints\n", (int)strings.size(), (int)valueChanges.size(), (int)intLists.size(), (int)chars.size(), (int)ints.size());
    l.numInts = 0
End Sub

' (VB) room for one more entry
Private Sub SyncDebugLog_grow(l As SyncDebugLog_)
    If l.numEntries < l.capacity Then Exit Sub
    If l.capacity = 0 Then l.capacity = 256 Else l.capacity = l.capacity * 2
    ReDim Preserve l.logType(0 To l.capacity - 1)
    ReDim Preserve l.entryFunction(0 To l.capacity - 1)
    ReDim Preserve l.entryString(0 To l.capacity - 1)
    ReDim Preserve l.entryValue(0 To l.capacity - 1)
    ReDim Preserve l.entryId(0 To l.capacity - 1)
End Sub

' SyncDebugLog::string, with SyncDebugString::set
Private Sub SyncDebugLog_string(l As SyncDebugLog_, ByVal f As String, ByVal s As String)
    SyncDebugLog_grow l
    l.entryFunction(l.numEntries) = f
    l.entryString(l.numEntries) = s
    l.crc = crcUpdateString(l.crc, f)
    l.crc = crcUpdateString(l.crc, s)

    l.logType(l.numEntries) = 115               ' 's'
    l.numEntries = l.numEntries + 1
End Sub

' SyncDebugLog::valueChange, with SyncDebugValueChange::set
Private Sub SyncDebugLog_valueChange(l As SyncDebugLog_, ByVal f As String, ByVal vn As String, ByVal nv As Long, ByVal i As Long)
    SyncDebugLog_grow l
    l.entryFunction(l.numEntries) = f
    l.entryString(l.numEntries) = vn
    l.entryValue(l.numEntries) = nv
    l.entryId(l.numEntries) = i
    l.crc = crcUpdateString(l.crc, f)
    l.crc = crcUpdateString(l.crc, vn)
    l.crc = crcUpdateInt(l.crc, nv)
    l.logType(l.numEntries) = 118               ' 'v'
    l.numEntries = l.numEntries + 1
End Sub

' SyncDebugLog::intList, with SyncDebugIntList::set: ints(first .. first + num - 1)
Private Sub SyncDebugLog_intList(l As SyncDebugLog_, ByVal f As String, ByVal s As String, ints() As Long, ByVal first As Long, ByVal num As Long)
    Dim n As Long, numInts As Long
    If l.numInts + num > l.intsCapacity Then
        If l.intsCapacity = 0 Then l.intsCapacity = 1024 Else l.intsCapacity = l.intsCapacity * 2
        If l.intsCapacity < l.numInts + num Then l.intsCapacity = l.numInts + num
        ReDim Preserve l.ints(0 To l.intsCapacity - 1)
    End If
    For n = 0 To num - 1
        l.ints(l.numInts + n) = ints(first + n)
    Next

    SyncDebugLog_grow l
    l.entryFunction(l.numEntries) = f
    l.entryString(l.numEntries) = s
    numInts = num                               ' std::min(num, ARRAY_SIZE(valueBytes))
    If numInts > 40 Then numInts = 40
    l.entryValue(l.numEntries) = numInts
    For n = 0 To numInts - 1
        l.crc = crcUpdateInt(l.crc, ints(first + n))
    Next
    l.numInts = l.numInts + num
    l.logType(l.numEntries) = 105               ' 'i'
    l.numEntries = l.numEntries + 1
End Sub

' SyncDebugLog::snprint (with each entry's snprint): the log as text, cut
' after bufSize characters as the C++ loop stops; the length it would have
' had is returned in printed
Private Function SyncDebugLog_snprint(l As SyncDebugLog_, ByVal bufSize As Long, printed As Long) As String
    Dim n As Long, intPtr As Long, k As Long, s As String, a As Variant
    Dim parts() As String, numParts As Long
    If l.numEntries > 0 Then ReDim parts(0 To l.numEntries - 1)
    For n = 0 To l.numEntries - 1
        If printed >= bufSize Then Exit For
        Select Case l.logType(n)
        Case 115                                ' 's'
            s = "[" & l.entryFunction(n) & "] " & l.entryString(n) & vbLf
        Case 118                                ' 'v'
            If l.entryId(n) <> -1 Then
                s = "[" & l.entryFunction(n) & "] " & CStr(l.entryId(n)) & " " & l.entryString(n) & " = " & CStr(l.entryValue(n)) & vbLf
            Else
                s = "[" & l.entryFunction(n) & "] " & l.entryString(n) & " = " & CStr(l.entryValue(n)) & vbLf
            End If
        Case 105                                ' 'i'
            s = "[" & l.entryFunction(n) & "] "
            If l.entryValue(n) = 0 Then
                s = s & l.entryString(n)
            ElseIf l.entryValue(n) <= 40 Then
                ReDim a(0 To l.entryValue(n) - 1)
                For k = 0 To l.entryValue(n) - 1
                    a(k) = l.ints(intPtr + k)
                Next
                s = s & vsprintf(l.entryString(n), a)
            Else
                s = s & "Too many ints in intlist."
            End If
            s = s & vbLf
            intPtr = intPtr + l.entryValue(n)
        End Select
        parts(numParts) = s
        numParts = numParts + 1
        printed = printed + Len(s)
    Next
    If numParts > 0 Then
        ReDim Preserve parts(0 To numParts - 1)
        SyncDebugLog_snprint = Join(parts, "")
    End If
End Function

'--------------------------------------------------------------------------

' _syncDebug(function, str, ...): the syncDebug(...) macro; VB passes the
' arguments as one array (args), as vsprintf takes them
Public Sub syncDebugV(ByVal function_ As String, ByVal str_ As String, args As Variant)
    Dim outputBuffer As String
    outputBuffer = vsprintf(str_, args)
    If Len(outputBuffer) > MAX_LEN_LOG_LINE - 1 Then outputBuffer = Left$(outputBuffer, MAX_LEN_LOG_LINE - 1)
    SyncDebugLog_string syncDebugLog(syncDebugNext), function_, outputBuffer
End Sub

' _syncDebugIntList(function, str, ints, numInts)
Public Sub syncDebugIntList_(ByVal function_ As String, ByVal str_ As String, ints() As Long, ByVal numInts As Long)
    SyncDebugLog_intList syncDebugLog(syncDebugNext), function_, str_, ints, LBound(ints), numInts
End Sub

' _syncDebugBacktrace(function): not the glibc branch
Public Sub syncDebugBacktrace_(ByVal function_ As String)
    Dim backupCrc As Long
    backupCrc = SyncDebugLog_getCrc(syncDebugLog(syncDebugNext))    ' Ignore CRC changes from _syncDebug(), since identical backtraces can be printed differently.

    syncDebugV "BT", "Sorry, syncDebugBacktrace() not implemented on your system. Called from %s.", Array(function_)

    ' Use CRC of something platform-independent, to avoid false positive desynchs.
    backupCrc = Not crcUpdateString(Not backupCrc, function_)
    SyncDebugLog_setCrc syncDebugLog(syncDebugNext), backupCrc
End Sub

Private Function SyncDebugLog_getCrc(l As SyncDebugLog_) As Long
    SyncDebugLog_getCrc = Not l.crc             ' Invert bits, since everyone else seems to do that with CRCs...
End Function

Private Sub SyncDebugLog_setCrc(l As SyncDebugLog_, ByVal newCrc As Long)
    l.crc = Not newCrc                          ' Invert bits, since everyone else seems to do that with CRCs...
End Sub

Public Function syncDebugGetCrc() As Long
    syncDebugGetCrc = SyncDebugLog_getCrc(syncDebugLog(syncDebugNext))
End Function

Public Sub syncDebugSetCrc(ByVal crc As Long)
    SyncDebugLog_setCrc syncDebugLog(syncDebugNext), crc
End Sub

Public Sub resetSyncDebug()
    Dim i As Long
    For i = 0 To MAX_SYNC_HISTORY - 1
        SyncDebugLog_clear syncDebugLog(i)
    Next

    syncDebugExtraGameTime = 0
    syncDebugExtraCrc = &HFFFFFFFF

    syncDebugNext = 0

    syncDebugNumDumps = 0
    DesyncLogOutputter_clear
    ' (debugSyncTmpBuf: VB keeps no buffer)

    targetDebugSyncTmpBufSize = kDefaultDebugSyncBufferSize
End Sub

' nextDebugSync(): returns the GameCrcType (uint16_t); VB passes gtime's
' gameTime, which lives in the exe
Public Function nextDebugSync(ByVal gameTime_ As Long) As Long
    Dim ret As Long
    ret = SyncDebugLog_getCrc(syncDebugLog(syncDebugNext))

    ' Save gameTime, so we know which CRC to compare with, later.
    syncDebugLog(syncDebugNext).time_ = gameTime_

    ' Go to next position, and free it ready for use.
    syncDebugNext = (syncDebugNext + 1) Mod MAX_SYNC_HISTORY
    SyncDebugLog_clear syncDebugLog(syncDebugNext)

    nextDebugSync = ret And &HFFFF&
End Function

' dumpLocalDebugSyncLog(logIndex): the text it puts in debugSyncTmpBuf
Public Function dumpLocalDebugSyncLog(ByVal logIndex As Long) As String
    Dim bufSizeLimit As Long, bufIndex As Long, s As String, out As String, printed As Long
    If targetDebugSyncTmpBufSize = 0 Then targetDebugSyncTmpBufSize = kDefaultDebugSyncBufferSize
    bufSizeLimit = targetDebugSyncTmpBufSize

    ' Dump our version, and also erase it, so we only dump it at most once.
    s = va("===== BEGIN gameTime=%u, %u entries, CRC 0x%08X =====\n", syncDebugLog(logIndex).time_, syncDebugLog(logIndex).numEntries, _
        SyncDebugLog_getCrc(syncDebugLog(logIndex)))
    bufIndex = Len(s)
    out = s
    enforceLimit bufIndex, bufSizeLimit         ' snprintf will not overflow debugSyncTmpBuf, but returns as much as it would have printed if possible.
    printed = bufIndex
    s = SyncDebugLog_snprint(syncDebugLog(logIndex), bufSizeLimit, printed)
    bufIndex = printed
    out = out & s
    enforceLimit bufIndex, bufSizeLimit
    s = va("===== END gameTime=%u, %u entries, CRC 0x%08X =====\n", syncDebugLog(logIndex).time_, syncDebugLog(logIndex).numEntries, _
        SyncDebugLog_getCrc(syncDebugLog(logIndex)))
    out = out & s
    bufIndex = bufIndex + Len(s)
    enforceLimit bufIndex, bufSizeLimit
    ' (VB) what snprintf could not write: the buffer holds bufSizeLimit - 1 characters and the NUL
    If Len(out) > bufIndex Then out = Left$(out, bufIndex)
    If bufIndex = bufSizeLimit And Len(out) = bufSizeLimit Then out = Left$(out, bufSizeLimit - 1) & vbNullChar
    dumpLocalDebugSyncLog = out
End Function

' (VB, for the tests) dumpLocalDebugSyncLog of the log being written
Public Function dumpCurrentDebugSyncLog() As String
    dumpCurrentDebugSyncLog = dumpLocalDebugSyncLog(syncDebugNext)
End Function

' (VB, for the tests) syncDebugNext: the index of the log being written
Public Function syncDebugNextIndex() As Long
    syncDebugNextIndex = syncDebugNext
End Function

' dumpLocalDebugSyncLog's lambda enforceLimit
Private Sub enforceLimit(bufIndex As Long, ByVal bufSizeLimit As Long)
    If bufIndex > bufSizeLimit Then
        bufIndex = bufSizeLimit
        ' grow buffer slightly for next debug sync log send (up to max)
        targetDebugSyncTmpBufSize = targetDebugSyncTmpBufSize + 500000
        If targetDebugSyncTmpBufSize > kMaxDebugSyncBufferSize Then targetDebugSyncTmpBufSize = kMaxDebugSyncBufferSize
    End If
End Sub

' dumpLocalDebugSyncLogByTime(time): "" where C++ returns 0
Private Function dumpLocalDebugSyncLogByTime(ByVal time_ As Long) As String
    Dim logIndex As Long
    For logIndex = 0 To MAX_SYNC_HISTORY - 1
        If syncDebugLog(logIndex).time_ = time_ Then
            Exit For                            ' found the log index for this time
        End If
    Next
    If logIndex >= MAX_SYNC_HISTORY Then
        ' did not find it
        debug_ 723, LOG_WARNING, "dumpLocalDebugSyncLogByTime", "Couldn't find gameTime: %u in history", time_
        Exit Function
    End If

    dumpLocalDebugSyncLogByTime = dumpLocalDebugSyncLog(logIndex)
End Function

' Set whether verbose debug mode - outputting the current player's sync log for every single game tick - is enabled
' WARNING: This may significantly impact performance *and* will fill up your drive with a lot of logs data!
' It is only intended to be used for debugging issues such as: replays desyncing when gameplay does not, etc. (And don't let the game run too long!)
Public Sub NET_setDebuggingModeVerboseOutputAllSyncLogs(ByVal untilGameTime As Long)
    verboseDebugOutputAllSyncLogsUntilGameTime = untilGameTime
    If verboseDebugOutputAllSyncLogsUntilGameTime <> 0 Then
        debug_ 796, LOG_INFO, "NET_setDebuggingModeVerboseOutputAllSyncLogs", "WARNING: Verbose sync log output is enabled!"
    End If
End Sub

' debugVerboseLogSyncIfNeeded(): VB passes gtime's gameTime
Public Sub debugVerboseLogSyncIfNeeded(ByVal gameTime_ As Long)
    Dim buf As String, fname As String, fp As Long, b() As Byte, n As Long
    If verboseDebugOutputAllSyncLogsUntilGameTime = 0 Or ULT(verboseDebugOutputAllSyncLogsUntilGameTime, gameTime_) Then
        Exit Sub
    End If

    buf = dumpLocalDebugSyncLogByTime(gameTime_)
    If Len(buf) > 0 Then
        fname = va("logs/sync%u_p%u.txt", gameTime_, selectedPlayer)
        fp = openSaveFile(fname)
        If fp = 0 Then
            debug_ 815, LOG_ERROR, "debugVerboseLogSyncIfNeeded", "Failed to open file for writing: %s", fname
            Exit Sub
        End If
        n = utf8_encode(buf, b)
        PHYSFS_writeBytes fp, b, 0, n
        PHYSFS_close fp
    End If
End Sub

'--------------------------------------------------------------------------
' DesyncLogOutputter

Private Function DesyncLogOutputter_getFilename(ByVal time_ As Long, ByVal player As Long) As String
    DesyncLogOutputter_getFilename = va("logs/desync%u_p%u.txt", time_, player)
End Function

' (VB) outputs.find(time): the index, -1 for end()
Private Function PlayerDesyncLogOutput_find(p As PlayerDesyncLogOutput_, ByVal time_ As Long) As Long
    Dim i As Long
    For i = 0 To p.numOutputs - 1
        If p.outputs(i).time_ = time_ Then
            PlayerDesyncLogOutput_find = i
            Exit Function
        End If
    Next
    PlayerDesyncLogOutput_find = -1
End Function

' (VB) outputs.insert(DesyncOutputs::value_type(time, {})).first: a new LogOutputter unless there is one for time
Private Function PlayerDesyncLogOutput_insert(p As PlayerDesyncLogOutput_, ByVal time_ As Long) As Long
    Dim i As Long, blank As LogOutputter_
    i = PlayerDesyncLogOutput_find(p, time_)
    If i >= 0 Then
        PlayerDesyncLogOutput_insert = i
        Exit Function
    End If
    If p.numOutputs = 0 Then
        ReDim p.outputs(0 To 3)
    ElseIf p.numOutputs > UBound(p.outputs) Then
        ReDim Preserve p.outputs(0 To p.numOutputs * 2 - 1)
    End If
    p.outputs(p.numOutputs) = blank
    p.outputs(p.numOutputs).time_ = time_
    PlayerDesyncLogOutput_insert = p.numOutputs
    p.numOutputs = p.numOutputs + 1
End Function

Private Sub DesyncLogOutputter_expectDesyncLog(ByVal time_ As Long)
    Dim a As Long
    For a = 0 To MAX_GAMEQUEUE_SLOTS - 1
        PlayerDesyncLogOutput_insert playerDesyncLogOutput(a), time_     ' setExpectDesyncLog
    Next
End Sub

Private Sub DesyncLogOutputter_willLogLocalDesyncLog(ByVal time_ As Long)
    If ULT(selectedPlayer, MAX_GAMEQUEUE_SLOTS) Then
        PlayerDesyncLogOutput_insert playerDesyncLogOutput(selectedPlayer), time_
    End If
End Sub

Private Function DesyncLogOutputter_canWriteLogFor(ByVal time_ As Long, ByVal player As Long) As Boolean
    If Not ULT(player, MAX_GAMEQUEUE_SLOTS) Then ASSERT_FAILURE "player < players.size()", "lib/netplay/sync_debug.cpp:392", "canWriteLogFor", "Invalid player idx: %u", player: DesyncLogOutputter_canWriteLogFor = False: Exit Function
    ' PlayerDesyncLogOutput::canWriteLogFor
    DesyncLogOutputter_canWriteLogFor = PlayerDesyncLogOutput_find(playerDesyncLogOutput(player), time_) >= 0 Or playerDesyncLogOutput(player).numOutputs < MaxPlayerSyncDebugDumps
End Function

Private Function DesyncLogOutputter_alreadyWroteLogFor(ByVal time_ As Long, ByVal player As Long) As Boolean
    Dim i As Long
    If Not ULT(player, MAX_GAMEQUEUE_SLOTS) Then ASSERT_FAILURE "player < players.size()", "lib/netplay/sync_debug.cpp:398", "alreadyWroteLogFor", "Invalid player idx: %u", player: DesyncLogOutputter_alreadyWroteLogFor = False: Exit Function
    ' PlayerDesyncLogOutput::alreadyWroteLogFor
    i = PlayerDesyncLogOutput_find(playerDesyncLogOutput(player), time_)
    If i < 0 Then
        DesyncLogOutputter_alreadyWroteLogFor = False
        Exit Function
    End If
    DesyncLogOutputter_alreadyWroteLogFor = playerDesyncLogOutput(player).outputs(i).noFurtherWrites
End Function

' write(time, player, buf, bufLen): buf(0 .. bufLen - 1)
Private Function DesyncLogOutputter_write(ByVal time_ As Long, ByVal player As Long, buf() As Byte, ByVal bufLen As Long) As Boolean
    Dim i As Long
    If Not ULT(player, MAX_GAMEQUEUE_SLOTS) Then ASSERT_FAILURE "player < players.size()", "lib/netplay/sync_debug.cpp:404", "write", "Invalid player idx: %u", player: DesyncLogOutputter_write = False: Exit Function
    ' PlayerDesyncLogOutput::write
    i = PlayerDesyncLogOutput_find(playerDesyncLogOutput(player), time_)
    If i < 0 Then
        If playerDesyncLogOutput(player).numOutputs >= MaxPlayerSyncDebugDumps Then
            DesyncLogOutputter_write = False
            Exit Function
        End If
        i = PlayerDesyncLogOutput_insert(playerDesyncLogOutput(player), time_)
    End If
    DesyncLogOutputter_write = LogOutputter_write(playerDesyncLogOutput(player).outputs(i), time_, player, buf, bufLen)
End Function

Private Function DesyncLogOutputter_closeEOF(ByVal time_ As Long, ByVal player As Long) As Boolean
    Dim i As Long
    ' PlayerDesyncLogOutput::closeEOF
    i = PlayerDesyncLogOutput_find(playerDesyncLogOutput(player), time_)
    If i < 0 Then
        DesyncLogOutputter_closeEOF = False
        Exit Function
    End If
    DesyncLogOutputter_closeEOF = LogOutputter_closeEOF(playerDesyncLogOutput(player).outputs(i))
End Function

Private Sub DesyncLogOutputter_clear()
    Dim a As Long, i As Long
    For a = 0 To MAX_GAMEQUEUE_SLOTS - 1
        ' PlayerDesyncLogOutput::clear: outputs.clear() (each LogOutputter's destructor closes its file)
        For i = 0 To playerDesyncLogOutput(a).numOutputs - 1
            LogOutputter_closeEOF playerDesyncLogOutput(a).outputs(i)
        Next
        playerDesyncLogOutput(a).numOutputs = 0
    Next
End Sub

Private Function LogOutputter_closeEOF(o As LogOutputter_) As Boolean
    If o.fp = 0 Then
        LogOutputter_closeEOF = False
        Exit Function
    End If
    PHYSFS_close o.fp
    o.fp = 0
    o.noFurtherWrites = True
    LogOutputter_closeEOF = True
End Function

Private Function LogOutputter_write(o As LogOutputter_, ByVal time_ As Long, ByVal player As Long, buf() As Byte, ByVal bufLen As Long) As Boolean
    Dim retVal As Boolean, bytesToWrite As Long, fname As String
    retVal = True

    If o.noFurtherWrites Then
        LogOutputter_write = False
        Exit Function
    End If
    bytesToWrite = bufLen
    If bytesToWrite > MaxDesyncLogSize - o.writtenSize Then
        ' truncate - and further attempts to write data will fail
        bytesToWrite = MaxDesyncLogSize - o.writtenSize
        retVal = False
    End If
    If bytesToWrite > 0 Then
        ' Open file if not already open
        If o.fp = 0 Then
            fname = DesyncLogOutputter_getFilename(time_, player)
            o.fp = openSaveFile(fname)
            If o.fp = 0 Then
                debug_ 511, LOG_ERROR, "write", "Failed to open file for writing: %s", fname
                o.noFurtherWrites = True        ' don't repeatedly try to open the file if the first attempt fails
                LogOutputter_write = False
                Exit Function
            End If
            ' WZ_PHYSFS_SETBUFFER(fp, 1024 * 1024)
        End If

        ' Write to file
        ' (bytesToWrite <= PHYSFS_uint32's max: a Long always is)
        PHYSFS_writeBytes o.fp, buf, 0, bytesToWrite

        o.writtenSize = o.writtenSize + bytesToWrite
    End If
    LogOutputter_write = retVal
End Function

'--------------------------------------------------------------------------

Private Sub dbgOutputDumpedSyncLog(ByVal time_ As Long, ByVal player As Long, Optional ByVal partial As Boolean = False, Optional ByVal syncError As Boolean = True)
    Dim fname As String, isSpectator As Boolean, typeDescription As String, playerDescription As String, partialText As String
    fname = DesyncLogOutputter_getFilename(time_, player)
    If ULT(player, MAX_CONNECTED_PLAYERS) Then isSpectator = NetPlay_players_isSpectator(player)
    If syncError Then typeDescription = "sync error" Else typeDescription = "sync log"
    If isSpectator Then playerDescription = "spectator" Else playerDescription = "player"
    If partial Then partialText = "<partial> " Else partialText = ""
    debug_ 640, LOG_ERROR, "dbgOutputDumpedSyncLog", "Dumped %s %u's %s%s at gameTime %u to file: %s%s", playerDescription, player, partialText, typeDescription, time_, PHYSFS_getRealDir(fname), fname
End Sub

' dumpLocalPlayerDebugSync(buf, bufLen, time, syncError): buf is the text dumpLocalDebugSyncLog gave
Private Sub dumpLocalPlayerDebugSync(ByVal buf As String, ByVal time_ As Long, Optional ByVal syncError As Boolean = True)
    Dim b() As Byte, n As Long
    n = utf8_encode(buf, b)
    DesyncLogOutputter_write time_, selectedPlayer, b, n
    DesyncLogOutputter_closeEOF time_, selectedPlayer
    dbgOutputDumpedSyncLog time_, selectedPlayer, False, syncError
End Sub

' sendDebugSync(buf, bufLen, time): buf is the text dumpLocalDebugSyncLog gave
Private Sub sendDebugSync(ByVal buf As String, ByVal time_ As Long)
    Dim b() As Byte, bufLen As Long, bytesWritten As Long, bytesToWrite As Long, numMessages As Long, w As MessageWriter
    ' Save our own, before sending, so that if we have 2 clients running on the same computer, to guarantee that it is done saving before the other client saves on top.
    dumpLocalPlayerDebugSync buf, time_

    ' Send multiple messages, of max MAX_DESYNC_LOG_TRANSFER_PACKET bytes of log data, until the full desync log is sent
    bufLen = utf8_encode(buf, b)
    bytesWritten = 0
    bytesToWrite = 0
    numMessages = 0
    Do While bytesWritten < bufLen
        bytesToWrite = bufLen - bytesWritten
        If bytesToWrite > MAX_DESYNC_LOG_TRANSFER_PACKET Then bytesToWrite = MAX_DESYNC_LOG_TRANSFER_PACKET
        Set w = NETbeginEncode(NETbroadcastQueue(), NET_DEBUG_SYNC)
        NETuint32_t_w w, time_
        NETuint32_t_w w, bytesToWrite
        w.bytes b, bytesWritten, bytesToWrite   ' NETbin(w, buf + bytesWritten, bytesToWrite)
        NETend_w w
        bytesWritten = bytesWritten + bytesToWrite
        numMessages = numMessages + 1
    Loop
    ' plus one 0-length buffer packet to signify EOF
    bytesToWrite = 0
    Set w = NETbeginEncode(NETbroadcastQueue(), NET_DEBUG_SYNC)
    NETuint32_t_w w, time_
    NETuint32_t_w w, bytesToWrite
    NETend_w w
    numMessages = numMessages + 1

    debug_ 678, LOG_INFO, "sendDebugSync", "Broadcast sync log (time: %u, size: %u) in %zu messages", time_, bufLen, numMessages
End Sub

' checkDebugSync(checkGameTime, checkCrc): checkCrc is a GameCrcType (uint16_t)
Public Function checkDebugSync(ByVal checkGameTime As Long, ByVal checkCrc As Long) As Boolean
    Dim logIndex As Long, buf As String, problemGameTime As Long
    checkCrc = checkCrc And &HFFFF&
    If checkGameTime = syncDebugLog(syncDebugNext).time_ Then  ' Can't happen - and syncDebugGameTime[] == 0, until just before sending the CRC, anyway.
        debug_ 828, LOG_ERROR, "checkDebugSync", "Huh? We aren't done yet..."
        checkDebugSync = True
        Exit Function
    End If

    For logIndex = 0 To MAX_SYNC_HISTORY - 1
        If syncDebugLog(logIndex).time_ = checkGameTime Then
            If (SyncDebugLog_getCrc(syncDebugLog(logIndex)) And &HFFFF&) = checkCrc Then
                checkDebugSync = True           ' Check passed. (So far... There might still be more players to compare CRCs with.)
                Exit Function
            End If

            debug_ 842, LOG_NET, "checkDebugSync", "Non-matching crc for gameTime: %u", checkGameTime
            Exit For                            ' Check failed!
        End If
    Next

    If logIndex >= MAX_SYNC_HISTORY And syncDebugExtraGameTime = checkGameTime Then
        If (syncDebugExtraCrc And &HFFFF&) = checkCrc Then
            checkDebugSync = True
            Exit Function
        End If
    End If

    If logIndex >= MAX_SYNC_HISTORY Then
        debug_ 857, LOG_NET, "checkDebugSync", "Couldn't check gameTime: %u", checkGameTime
        checkDebugSync = False                  ' Couldn't check. May have dumped already, or MAX_SYNC_HISTORY isn't big enough compared to the maximum latency.
        Exit Function
    End If

    If ULT(syncDebugNumDumps, MaxPlayerSyncDebugDumps) Then
        buf = dumpLocalDebugSyncLog(logIndex)
        syncDebugNumDumps = syncDebugNumDumps + 1
        problemGameTime = syncDebugLog(logIndex).time_
        DesyncLogOutputter_expectDesyncLog problemGameTime
        sendDebugSync buf, problemGameTime
    End If

    ' Backup correct CRC for checking against remaining players, even though we erased the logs (which were dumped already).
    syncDebugExtraGameTime = syncDebugLog(logIndex).time_
    syncDebugExtraCrc = SyncDebugLog_getCrc(syncDebugLog(logIndex))

    ' Finish erasing our version.
    SyncDebugLog_clear syncDebugLog(logIndex)

    checkDebugSync = False                      ' Ouch.
End Function

Public Sub recvDebugSync(ByVal queue As Long)
    Dim player As Long, time_ As Long, bufLen As Long, recvBuff() As Byte, r As MessageReader, buf As String
    If Not ULT(NETQUEUE_index(queue), MAX_CONNECTED_PLAYERS) Then ASSERT_FAILURE "queue.index < MAX_CONNECTED_PLAYERS", "lib/netplay/sync_debug.cpp:732", "recvDebugSync", "Invalid queue.index: %u", NETQUEUE_index(queue): Exit Sub
    player = NETQUEUE_index(queue)

    time_ = 0
    bufLen = 0

    ReDim recvBuff(0 To MAX_DESYNC_LOG_TRANSFER_PACKET - 1)

    ' Receive a message of desync data (the next chunk of a desync log)
    Set r = NETbeginDecode(queue, NET_DEBUG_SYNC)
    NETuint32_t_r r, time_
    NETuint32_t_r r, bufLen
    If Not DesyncLogOutputter_canWriteLogFor(time_, player) Then
        NETend_r r
        Exit Sub
    End If
    If ULT(MAX_DESYNC_LOG_TRANSFER_PACKET, bufLen) Then
        ' invalid packet - terminate
        If DesyncLogOutputter_closeEOF(time_, player) Then
            dbgOutputDumpedSyncLog time_, player, True
        End If
        NETend_r r
        Exit Sub
    End If
    NETbin_r r, recvBuff, bufLen
    NETend_r r

    If bufLen > 0 Then
        DesyncLogOutputter_write time_, player, recvBuff, bufLen
    Else
        ' 0-length buffer packet considered EOF
        If DesyncLogOutputter_closeEOF(time_, player) Then
            dbgOutputDumpedSyncLog time_, player, False
        End If
    End If

    ' Also dump the debug sync log for this local client (if possible & needed)
    If Not DesyncLogOutputter_alreadyWroteLogFor(time_, selectedPlayer) Then
        debug_ 778, LOG_INFO, "recvDebugSync", "Dumping local debug sync log in response to received for gameTime: %u", time_
        buf = dumpLocalDebugSyncLogByTime(time_)
        If Len(buf) > 0 Then
            DesyncLogOutputter_willLogLocalDesyncLog time_
            dumpLocalPlayerDebugSync buf, time_, False
        End If
    End If
End Sub
