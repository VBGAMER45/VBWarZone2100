Attribute VB_Name = "ziptests_main"
' (VB) ZipTests.exe: tools/zip_test.py's driver of the ported PhysFS (physfs_win, physfs_zip_c,
' physfs_miniz_c, compiled in without WzFramework). ZipTests.exe <commands file> <results file>: each line
' of the commands file (UTF-8, tab-separated) is a command, each writes result lines (tab-separated, UTF-8):
'   archive <path> <mount point or empty> <chunked 0/1>  mounts, walks the tree (enumerate, stat), reads
'                       every file whole and (chunked 1) in odd-sized chunks with seeks; E lines and a T line
'   mount <path> <mount point or empty> <append>   unmount <path>   exists <name>   isdir <name>
'   stat <name>   realdir <name>   enum <dir>   read <name>   seekread <name> <pos> <len>
'   keepopen <name>   closekept   searchpath   lookups <rounds> <name>|<name>|...   writedir <path>
'   mkdir <name>   write <name> <text>   delete <name>   lasterror   permitsymlinks <0/1>   init   deinit
' Every result line ends with the PhysFS error message the command left (WZ_PHYSFS_getLastError).
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function CryptAcquireContextW Lib "advapi32" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, ByVal pbData As Long, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As Long, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long

Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA1 As Long = &H8004&
Private Const HP_HASHVAL As Long = 2

Private outLines() As String
Private outCount As Long
Private hProv As Long
Private freq As Currency
Private keptHandle As Long
Private chunkSizes(0 To 6) As Long

Private Sub emit(ByVal s As String)
    If outCount > UBound(outLines) Then ReDim Preserve outLines(0 To 2 * outCount)
    outLines(outCount) = s
    outCount = outCount + 1
End Sub

Private Function nowMs() As Double
    Dim c As Currency
    QueryPerformanceCounter c
    nowMs = CDbl(c) * 1000# / CDbl(freq)
End Function

Private Function sha1(buf() As Byte, ByVal n As Long) As String
    Dim hHash As Long, h(0 To 19) As Byte, hl As Long, i As Long, s As String
    CryptCreateHash hProv, CALG_SHA1, 0, 0, hHash
    If n > 0 Then CryptHashData hHash, VarPtr(buf(0)), n, 0
    hl = 20
    CryptGetHashParam hHash, HP_HASHVAL, VarPtr(h(0)), hl, 0
    CryptDestroyHash hHash
    For i = 0 To 19
        s = s & Right$("0" & LCase$(Hex$(h(i))), 2)
    Next
    sha1 = s
End Function

Private Function errText() As String
    errText = WZ_PHYSFS_getLastError()
End Function

Private Function nz(ByVal s As String) As String
    If StrPtr(s) = 0 Then nz = "(null)" Else nz = s
End Function

Private Function optStr(ByVal s As String) As String
    If Len(s) = 0 Then optStr = vbNullString Else optStr = s
End Function

Private Function joinPath(ByVal d As String, ByVal n As String) As String
    If Len(d) = 0 Then joinPath = n Else joinPath = d & "/" & n
End Function

' reads the whole of name: its count, sha1, eof and tell after it, and what one more read gives
Private Sub readWhole(ByVal fname As String, ByVal size As Double, sha As String, count As Long, eofAfter As Long, tellAfter As Double, extra As Long, ok As Boolean)
    Dim h As Long, buf() As Byte, one(0 To 0) As Byte
    h = PHYSFS_openRead(fname)
    If h = 0 Then Exit Sub
    If size < 0 Then size = PHYSFS_fileLength(h)
    If size > 0 Then
        ReDim buf(0 To size - 1)
        count = PHYSFS_readBytes(h, buf, 0, size)
    Else
        ReDim buf(0 To 0)
        count = 0
    End If
    eofAfter = PHYSFS_eof(h)
    tellAfter = PHYSFS_tell(h)
    extra = PHYSFS_readBytes(h, one, 0, 1)
    PHYSFS_close h
    If count > 0 Then sha = sha1(buf, count) Else sha = sha1(buf, 0)
    ok = True
End Sub

' reads [from, upto) of h into buf in chunks of the cycling sizes; False when a read gives less or tell disagrees
Private Function readRange(ByVal h As Long, buf() As Byte, ByVal from As Long, ByVal upto As Long, ci As Long) As Boolean
    Dim pos As Long, n As Long, got As Long
    pos = from
    Do While pos < upto
        n = chunkSizes(ci Mod 7)
        ci = ci + 1
        If n > upto - pos Then n = upto - pos
        got = PHYSFS_readBytes(h, buf, pos, n)
        If got <> n Then Exit Function
        pos = pos + n
        If PHYSFS_tell(h) <> pos Then Exit Function
    Loop
    readRange = True
End Function

' the file in odd-sized chunks: the first half, a seek forward to 3/4 and the rest (eof then), a seek back to
' the half and up to 3/4; then a seek past the end. Its sha1, whether every step agreed, the past-end seek's result.
Private Sub readChunked(ByVal fname As String, ByVal size As Long, sha As String, stepsOk As Boolean, pastSeek As Long, pastErr As String)
    Dim h As Long, buf() As Byte, a As Long, b As Long, ci As Long
    h = PHYSFS_openRead(fname)
    If h = 0 Then Exit Sub
    If size > 0 Then ReDim buf(0 To size - 1) Else ReDim buf(0 To 0)
    a = size \ 2
    b = (3& * size) \ 4
    stepsOk = True
    If Not readRange(h, buf, 0, a, ci) Then stepsOk = False
    If PHYSFS_seek(h, b) = 0 Then stepsOk = False
    If PHYSFS_tell(h) <> b Then stepsOk = False
    If Not readRange(h, buf, b, size, ci) Then stepsOk = False
    If PHYSFS_eof(h) = 0 Then stepsOk = False
    If PHYSFS_seek(h, a) = 0 Then stepsOk = False
    If Not readRange(h, buf, a, b, ci) Then stepsOk = False
    sha = sha1(buf, size)
    errText                                             ' (clears it)
    pastSeek = PHYSFS_seek(h, CDbl(size) + 1)
    pastErr = errText()
    PHYSFS_close h
End Sub

Private Sub walk(ByVal d As String, ByVal rel As String, ByVal chunked As Boolean, nEntries As Long, bytes As Double, readMs As Double, chunkMs As Double, statMs As Double)
    Dim names() As String, n As Long, i As Long, full As String, r As String
    Dim filesize As Double, modtime As Double, filetype As Long, ro As Long, rc As Long, t As Double
    Dim sha As String, count As Long, eofAfter As Long, tellAfter As Double, extra As Long, ok As Boolean
    Dim shac As String, stepsOk As Boolean, pastSeek As Long, pastErr As String
    t = nowMs()
    n = PHYSFS_enumerateFiles(d, names)
    statMs = statMs + (nowMs() - t)
    If n < 0 Then
        emit "EERR" & vbTab & rel & vbTab & errText()
        Exit Sub
    End If
    For i = 0 To n - 1
        full = joinPath(d, names(i))
        r = joinPath(rel, names(i))
        t = nowMs()
        rc = PHYSFS_stat(full, filesize, modtime, filetype, ro)
        statMs = statMs + (nowMs() - t)
        nEntries = nEntries + 1
        If rc = 0 Then
            emit "E" & vbTab & r & vbTab & "stat failed" & vbTab & errText()
        ElseIf filetype = PHYSFS_FILETYPE_DIRECTORY Then
            emit "E" & vbTab & r & vbTab & "D" & vbTab & filesize & vbTab & modtime & vbTab & ro
            walk full, r, chunked, nEntries, bytes, readMs, chunkMs, statMs
        Else
            sha = "": count = -2: eofAfter = -2: tellAfter = -2: extra = -2: ok = False
            t = nowMs()
            readWhole full, filesize, sha, count, eofAfter, tellAfter, extra, ok
            readMs = readMs + (nowMs() - t)
            bytes = bytes + filesize
            shac = "-": stepsOk = False: pastSeek = -1: pastErr = "-"
            If chunked Then
                t = nowMs()
                readChunked full, filesize, shac, stepsOk, pastSeek, pastErr
                chunkMs = chunkMs + (nowMs() - t)
            End If
            emit "E" & vbTab & r & vbTab & filetype & vbTab & filesize & vbTab & modtime & vbTab & ro & vbTab & _
                 IIf(ok, "1", "0") & vbTab & sha & vbTab & count & vbTab & eofAfter & vbTab & tellAfter & vbTab & extra & vbTab & _
                 shac & vbTab & IIf(stepsOk, "1", "0") & vbTab & pastSeek & vbTab & pastErr
        End If
    Next
End Sub

Private Sub cmdArchive(f() As String)
    Dim rc As Long, t0 As Double, mountMs As Double, nEntries As Long, bytes As Double, readMs As Double, chunkMs As Double
    Dim statMs As Double, t1 As Double, mp As String
    mp = optStr(f(2))
    errText
    t0 = nowMs()
    rc = PHYSFS_mount(f(1), mp, 1)
    mountMs = nowMs() - t0
    emit "A" & vbTab & f(1) & vbTab & rc & vbTab & errText() & vbTab & Format$(mountMs, "0.0")
    If rc = 0 Then Exit Sub
    t1 = nowMs()
    walk f(2), "", (f(3) = "1"), nEntries, bytes, readMs, chunkMs, statMs
    emit "T" & vbTab & f(1) & vbTab & nEntries & vbTab & bytes & vbTab & Format$(mountMs, "0.0") & vbTab & Format$(statMs, "0.0") & vbTab & _
         Format$(readMs, "0.0") & vbTab & Format$(chunkMs, "0.0") & vbTab & Format$(nowMs() - t1, "0.0")
    rc = PHYSFS_unmount(f(1))
    emit "AU" & vbTab & f(1) & vbTab & rc & vbTab & errText()
End Sub

Private Sub cmdLookups(f() As String)
    Dim names() As String, rounds As Long, r As Long, i As Long, t0 As Double, nx As Long, h As Long, n As Long, ms As Double
    rounds = CLng(f(1))
    names = Split(f(2), "|")
    t0 = nowMs()
    For r = 1 To rounds
        For i = 0 To UBound(names)
            If PHYSFS_exists(names(i)) <> 0 Then nx = nx + 1
        Next
    Next
    ms = nowMs() - t0
    n = rounds * (UBound(names) + 1)
    emit "L" & vbTab & "exists" & vbTab & n & vbTab & nx & vbTab & Format$(ms * 1000# / n, "0.0")
    nx = 0
    t0 = nowMs()
    For r = 1 To rounds
        For i = 0 To UBound(names)
            h = PHYSFS_openRead(names(i))
            If h <> 0 Then
                nx = nx + 1
                PHYSFS_close h
            End If
        Next
    Next
    ms = nowMs() - t0
    emit "L" & vbTab & "openRead+close" & vbTab & n & vbTab & nx & vbTab & Format$(ms * 1000# / n, "0.0")
End Sub

' rawlookups <rounds> <names> <folders>: each name looked up as the port's PhysFS did before (one
' GetFileAttributesW per folder until one has it): "rawlookups <count> <found> <us each>"
Private Sub cmdRawLookups(f() As String)
    Dim names() As String, dirs() As String, rounds As Long, r As Long, i As Long, k As Long, t0 As Double, nx As Long, n As Long
    Dim p As String
    rounds = CLng(f(1))
    names = Split(f(2), "|")
    dirs = Split(f(3), "|")
    t0 = nowMs()
    For r = 1 To rounds
        For i = 0 To UBound(names)
            For k = 0 To UBound(dirs)
                p = dirs(k) & "\" & Replace(names(i), "/", "\")
                If GetFileAttributesW(StrPtr(p)) <> -1 Then
                    nx = nx + 1
                    Exit For
                End If
            Next
        Next
    Next
    n = rounds * (UBound(names) + 1)
    emit "rawlookups" & vbTab & n & vbTab & nx & vbTab & Format$((nowMs() - t0) * 1000# / n, "0.0")
End Sub

Private Sub runCommand(ByVal cmdLine As String)
    Dim f() As String, rc As Long, filesize As Double, modtime As Double, filetype As Long, ro As Long
    Dim names() As String, n As Long, i As Long, s As String, h As Long, buf() As Byte, got As Long, sha As String
    Dim count As Long, eofAfter As Long, tellAfter As Double, extra As Long, ok As Boolean
    f = Split(cmdLine & vbTab & vbTab & vbTab & vbTab, vbTab)
    Select Case f(0)
    Case "init"
        rc = PHYSFS_init("ZipTests")
        PHYSFS_permitSymbolicLinks 0                    ' as main.cpp does
        emit "init" & vbTab & rc & vbTab & errText()
    Case "deinit"
        rc = PHYSFS_deinit()
        emit "deinit" & vbTab & rc & vbTab & errText()
    Case "archive"
        cmdArchive f
    Case "mount"
        rc = PHYSFS_mount(f(1), optStr(f(2)), CLng(f(3)))
        emit "mount" & vbTab & f(1) & vbTab & f(2) & vbTab & f(3) & vbTab & rc & vbTab & errText()
    Case "unmount"
        rc = PHYSFS_unmount(f(1))
        emit "unmount" & vbTab & f(1) & vbTab & rc & vbTab & errText()
    Case "exists"
        rc = PHYSFS_exists(f(1))
        emit "exists" & vbTab & f(1) & vbTab & rc & vbTab & errText()
    Case "isdir"
        rc = WZ_PHYSFS_isDirectory(f(1))
        emit "isdir" & vbTab & f(1) & vbTab & rc & vbTab & errText()
    Case "stat"
        rc = PHYSFS_stat(f(1), filesize, modtime, filetype, ro)
        emit "stat" & vbTab & f(1) & vbTab & rc & vbTab & filetype & vbTab & filesize & vbTab & modtime & vbTab & ro & vbTab & errText()
    Case "realdir"
        s = PHYSFS_getRealDir(f(1))
        emit "realdir" & vbTab & f(1) & vbTab & nz(s) & vbTab & errText()
    Case "enum"
        n = PHYSFS_enumerateFiles(f(1), names)
        s = ""
        For i = 0 To n - 1
            If i > 0 Then s = s & "|"
            s = s & names(i)
        Next
        emit "enum" & vbTab & f(1) & vbTab & n & vbTab & s & vbTab & errText()
    Case "read"
        sha = "": count = -2: eofAfter = -2: tellAfter = -2: extra = -2: ok = False
        readWhole f(1), -1, sha, count, eofAfter, tellAfter, extra, ok
        emit "read" & vbTab & f(1) & vbTab & IIf(ok, "1", "0") & vbTab & count & vbTab & sha & vbTab & eofAfter & vbTab & tellAfter & vbTab & extra & vbTab & errText()
    Case "seekread"
        h = PHYSFS_openRead(f(1))
        If h = 0 Then
            emit "seekread" & vbTab & f(1) & vbTab & f(2) & vbTab & f(3) & vbTab & "open failed" & vbTab & errText()
        Else
            rc = PHYSFS_seek(h, CDbl(f(2)))
            s = errText()
            ReDim buf(0 To CLng(f(3)))
            got = PHYSFS_readBytes(h, buf, 0, CLng(f(3)))
            If got > 0 Then sha = sha1(buf, got) Else sha = "-"
            emit "seekread" & vbTab & f(1) & vbTab & f(2) & vbTab & f(3) & vbTab & rc & vbTab & s & vbTab & got & vbTab & sha & vbTab & PHYSFS_tell(h) & vbTab & PHYSFS_eof(h) & vbTab & PHYSFS_fileLength(h)
            PHYSFS_close h
        End If
    Case "keepopen"
        keptHandle = PHYSFS_openRead(f(1))
        emit "keepopen" & vbTab & f(1) & vbTab & IIf(keptHandle <> 0, "1", "0") & vbTab & errText()
    Case "closekept"
        rc = PHYSFS_close(keptHandle)
        keptHandle = 0
        emit "closekept" & vbTab & rc & vbTab & errText()
    Case "searchpath"
        n = PHYSFS_getSearchPath(names)
        s = ""
        For i = 0 To n - 1
            If i > 0 Then s = s & "|"
            s = s & names(i)
        Next
        emit "searchpath" & vbTab & n & vbTab & s
    Case "lookups"
        cmdLookups f
    Case "rawlookups"
        cmdRawLookups f
    Case "writedir"
        rc = PHYSFS_setWriteDir(optStr(f(1)))
        emit "writedir" & vbTab & f(1) & vbTab & rc & vbTab & nz(PHYSFS_getWriteDir()) & vbTab & errText()
    Case "mkdir"
        rc = PHYSFS_mkdir(f(1))
        emit "mkdir" & vbTab & f(1) & vbTab & rc & vbTab & errText()
    Case "write"
        h = PHYSFS_openWrite(f(1))
        If h = 0 Then
            emit "write" & vbTab & f(1) & vbTab & f(2) & vbTab & "open failed" & vbTab & errText()
        Else
            n = utf8_encode(f(2), buf)
            got = PHYSFS_writeBytes(h, buf, 0, n)
            rc = PHYSFS_close(h)
            emit "write" & vbTab & f(1) & vbTab & f(2) & vbTab & got & vbTab & rc & vbTab & errText()
        End If
    Case "delete"
        rc = PHYSFS_delete(f(1))
        emit "delete" & vbTab & f(1) & vbTab & rc & vbTab & errText()
    Case "lasterror"
        emit "lasterror" & vbTab & errText()
    Case "permitsymlinks"
        PHYSFS_permitSymbolicLinks CLng(f(1))
        emit "permitsymlinks" & vbTab & f(1)
    Case ""
    Case Else
        emit "unknown" & vbTab & f(0)
    End Select
End Sub

Private Function readText(ByVal path As String) As String
    Dim ff As Integer, b() As Byte, n As Long
    ff = FreeFile
    Open path For Binary Access Read As #ff
    n = LOF(ff)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #ff, 1, b
    End If
    Close #ff
    If n > 0 Then readText = utf8_decode(b, 0, n)
End Function

Private Sub writeText(ByVal path As String, ByVal s As String)
    Dim ff As Integer, b() As Byte, n As Long
    n = utf8_encode(s, b)
    On Error Resume Next
    Kill path
    On Error GoTo 0
    ff = FreeFile
    Open path For Binary Access Write As #ff
    If n > 0 Then Put #ff, 1, b
    Close #ff
End Sub

Sub Main()
    Dim args() As String, lines() As String, i As Long, outPath As String, ln As String
    On Error GoTo failed
    args = Split(Command$, " ")
    If UBound(args) < 1 Then ExitProcess 2
    outPath = Replace(args(1), """", "")
    ReDim outLines(0 To 255)
    QueryPerformanceFrequency freq
    CryptAcquireContextW hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT
    chunkSizes(0) = 1: chunkSizes(1) = 7: chunkSizes(2) = 333: chunkSizes(3) = 4093
    chunkSizes(4) = 65537: chunkSizes(5) = 17: chunkSizes(6) = 1000003
    lines = Split(readText(Replace(args(0), """", "")), vbLf)
    For i = 0 To UBound(lines)
        ln = Replace(lines(i), vbCr, "")
        If Len(ln) > 0 Then
            If Left$(ln, 1) <> "#" Then runCommand ln
        End If
    Next
    emit "done"
    ReDim Preserve outLines(0 To outCount - 1)
    writeText outPath, Join(outLines, vbLf)
    ExitProcess 0
    Exit Sub
failed:
    emit "VBERROR" & vbTab & Err.Number & vbTab & Err.Description & vbTab & "after line " & i
    writeText outPath, Join(outLines, vbLf)
    ExitProcess 1
End Sub
