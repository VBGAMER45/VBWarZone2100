Attribute VB_Name = "debug_c"
' lib/framework/debug.cpp, with the macros of debug.h. Function pointers
' (debug_callback_fn, _init, _exit) are the CB_ numbers of FrameworkApi,
' dispatched by DebugCallbackCall/Init/Exit; the callback list is an array
' in registration order. Single-threaded, so the mutexes and thread_local
' buffers are plain module variables. Not ported: debug_callback_win32debug
' and the MSVC memory checks (DEBUG builds only), the Emscripten log,
' _debugBacktrace (Linux only), wz_assert (assertEnabled is false in release).
' The C++ macros carry __LINE__ and __FUNCTION__: here the callers pass them.
Option Explicit

Private Const MAX_LEN_LOG_LINE As Long = 1024
Private Const MAX_EVENT_NAME_LEN As Long = 100

Private Type debug_callback
    callback As Long
    init As Long
    exit_ As Long
    data As Variant
End Type

Public last_called_script_event As String     ' char[MAX_EVENT_NAME_LEN]
Public traceID As Long                         ' UDWORD; -1, set by debug_init (C initialises it statically)

Private callbackRegistry() As debug_callback
Private callbackCount As Long
Public enabled_debug(0 To LOG_LAST - 1) As Boolean
Public assertEnabled As Boolean                ' release: false

Private code_part_names(0 To LOG_LAST) As String
Private namesReady As Boolean

Private useInputBuffer1 As Boolean
Private debug_flush_stderr As Boolean
Private warning_list As Object                 ' std::map<std::string, int>: Scripting.Dictionary

Private WZ_DBGFile As String
Private WZDebugfilename As String

Private captureBuffer As String               ' (VB) CB_debug_callback_capture's lines

' thread_local in C (a single thread here)
Private errorStore As String
Private errorWaiting As Boolean
Private inputBuffer(0 To 1) As String
Private repeated As Long
Private next_ As Long
Private prev As Long

' _debugFromGfxCallback's DebugGfxCallbackPersistentData
Private lastGfxCallbackMsg As String
Private gfxRepeated As Long
Private gfxPrev As Long

Private Sub InitNames()
    Dim names As Variant, i As Long
    ' This list _must_ match the enum code_part (FrameworkApi)!
    names = Array("all", "main", "sound", "video", "wz", "3d", "texture", "net", "memory", "warning", _
        "error", "never", "script", "movement", "attack", "fog", "sensor", "gui", "map", "save", _
        "sync", "death", "life", "gateway", "message", "info", "terrain", "feature", "fatal", "input", _
        "popup", "console", "activity", "research", "savegame", "repairs", "lobby", "last")
    For i = 0 To LOG_LAST
        code_part_names(i) = names(i)
    Next
    next_ = 2
    namesReady = True
End Sub

' Convert code_part names to enum. Case insensitive. Codepart number or LOG_LAST if can't match.
Private Function code_part_from_str(ByVal str_ As String) As Long
    Dim i As Long
    If Not namesReady Then InitNames
    For i = 0 To LOG_LAST - 1
        If StrComp(code_part_names(i), str_, vbTextCompare) = 0 Then
            code_part_from_str = i
            Exit Function
        End If
    Next
    code_part_from_str = LOG_LAST
End Function

' Callback for outputing to stderr
Public Sub debug_callback_stderr(data As Variant, ByVal outputBuffer As String, ByVal part As Long)
    If Right$(outputBuffer, 1) <> vbLf Then
        fprintf STDERR_FILE, "%s\n", outputBuffer
    Else
        fprintf STDERR_FILE, "%s", outputBuffer
    End If

    ' Make sure that all output is flushed to stderr when requested by the user
    If debug_flush_stderr Then
        fflush STDERR_FILE
    End If
End Sub

' Callback for outputing to a file. data: the file (stdio_win FILE number).
Public Sub debug_callback_file(data As Variant, ByVal outputBuffer As String, ByVal part As Long)
    Dim logfile As Long
    logfile = data

    If InStr(outputBuffer, vbLf) = 0 Then
        fprintf logfile, "%s\n", outputBuffer
    Else
        fprintf logfile, "%s", outputBuffer
    End If
End Sub

' One prefix of replaceUsernameInPath's patterns: ^<lead>(<a>|<b>)<sep>[^<sep>]+
Private Function ReplaceUserFolder(ByVal path As String, ByVal leadLen As Long, ByVal a As String, ByVal b As String, ByVal sep As String, ByVal replacement As String) As String
    Dim p As Long, n As Long
    ReplaceUserFolder = path
    p = leadLen + 1
    If Mid$(path, p, Len(a) + 1) = a & sep Then
        p = p + Len(a) + 1
    ElseIf Mid$(path, p, Len(b) + 1) = b & sep Then
        p = p + Len(b) + 1
    Else
        Exit Function
    End If
    n = p
    Do While n <= Len(path)
        If Mid$(path, n, 1) = sep Then Exit Do
        n = n + 1
    Loop
    If n = p Then Exit Function                 ' [^sep]+ needs one character
    ReplaceUserFolder = replacement & Mid$(path, n)
End Function

Private Function replaceUsernameInPath(ByVal path As String) As String
    Dim result As String
    ' Some common forms:
    ' C:\Users\<USERNAME>\... (Windows Vista+)
    ' C:\Documents and Settings\<USERNAME>\... (Windows XP)
    ' /Users/<USERNAME>/... (macOS)
    ' /home/<USERNAME>/... (Linux)

    ' First pass, do a simple string regex replacement
    result = path
    If Len(result) >= 3 Then
        If Left$(result, 1) Like "[A-Za-z]" And Mid$(result, 2, 2) = ":\" Then
            result = ReplaceUserFolder(result, 3, "Users", "Documents and Settings", "\", "[WIN_USER_FOLDER]")
        End If
    End If
    If Left$(result, 1) = "/" Then
        result = ReplaceUserFolder(result, 1, "Users", "home", "/", "[USER_HOME]")
    End If
    replaceUsernameInPath = result
End Function

' Setup the file callback. data: in, the filename; out, the file.
Public Function debug_callback_file_init(data As Variant) As Boolean
    Dim logfile As Long, fileNameStripped As String
    If IsEmpty(data) Then
        ASSERT_FAILURE "pFileName", "lib/framework/debug.cpp:253", "debug_callback_file_init", "Debug filename required"
    End If
    WZDebugfilename = data & ""

    ' (VB) Warzone converts the UTF-8 name for _wfopen; a VB String already is UTF-16.
    logfile = fopen(WZDebugfilename, "w")
    If logfile = 0 Then
        fprintf STDERR_FILE, "Could not open %s for appending!\n", WZDebugfilename
        Exit Function
    End If
    WZ_DBGFile = Left$(WZDebugfilename, 260 - 1)  ' char WZ_DBGFile[PATH_MAX]
    ' setbuf(logfile, nullptr): stdio_win writes straight through
    fileNameStripped = replaceUsernameInPath(WZDebugfilename)
    fprintf logfile, "--- Starting log [%s]---\n", fileNameStripped
    data = logfile

    debug_callback_file_init = True
End Function

' Shutdown the file callback: closes the logfile.
Public Sub debug_callback_file_exit(data As Variant)
    Dim logfile As Long
    logfile = data
    fclose logfile
    data = 0
End Sub

' (VB) Callback for the test runs: keeps the lines (as debug_callback_file writes them) for debug_capture_take
Private Sub debug_callback_capture(data As Variant, ByVal outputBuffer As String, ByVal part As Long)
    If InStr(outputBuffer, vbLf) = 0 Then
        captureBuffer = captureBuffer & outputBuffer & vbLf
    Else
        captureBuffer = captureBuffer & outputBuffer
    End If
End Sub

' (VB) the captured lines since the last call
Public Function debug_capture_take() As String
    debug_capture_take = captureBuffer
    captureBuffer = ""
End Function

Private Sub DebugCallbackCall(ByVal callback As Long, data As Variant, ByVal outputBuffer As String, ByVal part As Long)
    Select Case callback
    Case CB_debug_callback_stderr: debug_callback_stderr data, outputBuffer, part
    Case CB_debug_callback_file: debug_callback_file data, outputBuffer, part
    Case CB_debug_callback_capture: debug_callback_capture data, outputBuffer, part
    End Select
End Sub

Private Function DebugCallbackInit(ByVal init As Long, data As Variant) As Boolean
    Select Case init
    Case CB_debug_callback_file_init: DebugCallbackInit = debug_callback_file_init(data)
    End Select
End Function

Private Sub DebugCallbackExit(ByVal exit_ As Long, data As Variant)
    Select Case exit_
    Case CB_debug_callback_file_exit: debug_callback_file_exit data
    End Select
End Sub

Public Sub debugFlushStderr()
    debug_flush_stderr = True
End Sub

Public Sub debug_init()
    Dim i As Long
    '*** Initialize the debug subsystem ***
    If Not namesReady Then InitNames
    traceID = -1

    For i = 0 To LOG_LAST - 1
        enabled_debug(i) = False
    Next
    enabled_debug(LOG_ERROR) = True
    enabled_debug(LOG_INFO) = True
    enabled_debug(LOG_FATAL) = True
    enabled_debug(LOG_POPUP) = True
End Sub

Public Sub debug_exit()
    Dim i As Long
    For i = 0 To callbackCount - 1
        If callbackRegistry(i).exit_ <> CB_NULL Then
            DebugCallbackExit callbackRegistry(i).exit_, callbackRegistry(i).data
        End If
    Next
    If Not warning_list Is Nothing Then warning_list.RemoveAll
    callbackCount = 0
    Erase callbackRegistry
End Sub

Public Sub debug_register_callback(ByVal callback As Long, ByVal init As Long, ByVal exit_ As Long, ByVal data As Variant)
    Dim tmpCallback As debug_callback

    tmpCallback.callback = callback
    tmpCallback.init = init
    tmpCallback.exit_ = exit_
    tmpCallback.data = data

    If tmpCallback.init <> CB_NULL Then
        If Not DebugCallbackInit(tmpCallback.init, tmpCallback.data) Then
            fprintf STDERR_FILE, "Failed to initialise debug callback, debugfile set to [%s]!\n", WZDebugfilename
            Exit Sub
        End If
    End If

    ReDim Preserve callbackRegistry(0 To callbackCount)
    callbackRegistry(callbackCount) = tmpCallback
    callbackCount = callbackCount + 1
End Sub

Public Function debug_enable_switch(ByVal str_ As String) As Boolean
    Dim part As Long, i As Long
    part = code_part_from_str(str_)

    If part <> LOG_LAST Then
        enabled_debug(part) = Not enabled_debug(part)
    End If
    If part = LOG_ALL Then
        For i = 0 To LOG_LAST - 1
            enabled_debug(i) = True
        Next
    End If
    debug_enable_switch = (part <> LOG_LAST)
End Function

' Send the given string to all debug callbacks.
Private Sub printToDebugCallbacks(ByVal str_ As String, ByVal part As Long)
    Dim i As Long
    ' Loop over all callbacks, invoking them with the given data string
    For i = 0 To callbackCount - 1
        DebugCallbackCall callbackRegistry(i).callback, callbackRegistry(i).data, str_, part
    Next
End Sub

' C: _realObjTrace (id, function, str, ...)
Public Sub realObjTrace_(ByVal id As Long, ByVal function_ As String, ByVal str_ As String, args As Variant)
    Dim vaBuffer As String, outputBuffer As String

    vaBuffer = Left$(vsprintf(str_, args), MAX_LEN_LOG_LINE - 1)

    outputBuffer = Left$(va("[%6d]: [%s] %s", id, function_, vaBuffer), MAX_LEN_LOG_LINE - 1)
    printToDebugCallbacks outputBuffer, LOG_INFO
End Sub

' Returns vbNullString (C: nullptr) when there is no error waiting.
Public Function debugLastError() As String
    If errorWaiting Then
        errorWaiting = False
        debugLastError = errorStore
    Else
        debugLastError = vbNullString
    End If
End Function

Private Sub debugDisplayFatalErrorMsgBox(ByVal outputLogLine As String)
    Dim wbuf As String
    If wzIsFullscreen() Then
        wzChangeWindowMode WINDOW_MODE_windowed, True
    End If
    ' The C literal continues with a backslash-newline: the next line's two tabs belong to the string.
    wbuf = Left$(va("%s\n\nPlease check the file (%s) in your configuration directory for more details. " & vbTab & vbTab & _
        "\nDo not forget to upload the %s file, WZdebuginfo.txt and the warzone2100.rpt files in your bug reports at https://github.com/Warzone2100/warzone2100/issues/new!", _
        outputLogLine, WZ_DBGFile, WZ_DBGFile), MAX_LEN_LOG_LINE + 512 - 1)
    wzDisplayDialog Dialog_Error, "Warzone has terminated unexpectedly", wbuf
End Sub

' Special version of _debug that doesn't use any of the normal global debug-logging state (or special sauce)
' *except* the callbackRegistry (via printToDebugCallbacks). C: _debugFromGfxCallback.
Public Sub debugFromGfxCallback_(ByVal line_ As Long, ByVal part As Long, ByVal function_ As String, ByVal str_ As String, args As Variant)
    Dim outputBuffer As String, tmpRepeatBuffer As String, prefix As String, contents As String
    Dim ourtime As String, curr_repeated As Long
    Dim isRepeatedPowOf2 As Boolean

    ourtime = strftime("%H:%M:%S", getLocalTime(time_(), True))

    prefix = Left$(va("%-8s|%s: [%s:%d] ", code_part_names(part), ourtime, function_, line_), MAX_LEN_LOG_LINE - 1)

    ' Then calculate the remaining length manually
    contents = Left$(vsprintf(str_, args), MAX_LEN_LOG_LINE - Len(prefix) - 1)

    If lastGfxCallbackMsg = contents Then
        ' last message, duplicated
        gfxRepeated = gfxRepeated + 1
        isRepeatedPowOf2 = gfxRepeated >= 2 And ((gfxRepeated And (gfxRepeated - 1)) = 0)
        If isRepeatedPowOf2 Then
            If gfxRepeated > 2 Then
                contents = Left$(va("last message repeated %u times (total %u repeats)", gfxRepeated - gfxPrev, gfxRepeated), MAX_LEN_LOG_LINE - Len(prefix) - 1)
            Else
                contents = Left$(va("last message repeated %u times", gfxRepeated - gfxPrev), MAX_LEN_LOG_LINE - Len(prefix) - 1)
            End If
            gfxPrev = gfxRepeated
        Else
            Exit Sub ' do nothing - no printing
        End If
    Else
        ' Received another line, cleanup the old
        If gfxRepeated > 0 And gfxRepeated <> gfxPrev And gfxRepeated <> 1 Then
            ' just repeat the previous message when only one repeat occurred
            If gfxRepeated > 2 Then
                tmpRepeatBuffer = Left$(va("last message repeated %u times (total %u repeats)", gfxRepeated - gfxPrev, gfxRepeated), 127)
            Else
                tmpRepeatBuffer = Left$(va("last message repeated %u times", gfxRepeated - gfxPrev), 127)
            End If
        End If
        gfxRepeated = 0
        gfxPrev = 0
        lastGfxCallbackMsg = contents
    End If
    curr_repeated = gfxRepeated
    outputBuffer = prefix & contents

    If Len(tmpRepeatBuffer) > 0 Then
        printToDebugCallbacks tmpRepeatBuffer, part
    End If

    printToDebugCallbacks outputBuffer, part

    If curr_repeated = 0 Then
        If part = LOG_FATAL Then
            debugDisplayFatalErrorMsgBox contents
        End If
    End If
End Sub

' C: _debug (line, part, function, str, ...), with the arguments as a Variant array.
Public Sub debugV(ByVal line_ As Long, ByVal part As Long, ByVal function_ As String, ByVal str_ As String, args As Variant)
    Dim outputBuffer As String, cur As Long, key As String
    Dim ourtime As String

    If Not namesReady Then InitNames
    outputBuffer = Left$(vsprintf(str_, args), MAX_LEN_LOG_LINE - 1)

    If useInputBuffer1 Then cur = 1 Else cur = 0
    If part = LOG_WARNING Then
        If warning_list Is Nothing Then Set warning_list = CreateObject("Scripting.Dictionary")
        key = function_ & "-" & outputBuffer
        If Not warning_list.Exists(key) Then
            warning_list.Add key, line_
            inputBuffer(cur) = Left$(va("[%s:%d] %s (**Further warnings of this type are suppressed.)", function_, line_, outputBuffer), MAX_LEN_LOG_LINE - 1)
        Else
            Exit Sub ' don't bother adding any more
        End If
    Else
        inputBuffer(cur) = Left$(va("[%s:%d] %s", function_, line_, outputBuffer), MAX_LEN_LOG_LINE - 1)
    End If

    If inputBuffer(0) = inputBuffer(1) Then
        ' Received again the same line
        repeated = repeated + 1
        If repeated = next_ Then
            If repeated > 2 Then
                outputBuffer = va("last message repeated %u times (total %u repeats)", repeated - prev, repeated)
            Else
                outputBuffer = va("last message repeated %u times", repeated - prev)
            End If
            printToDebugCallbacks outputBuffer, part
            prev = repeated
            next_ = next_ * 2
        End If
    Else
        ' Received another line, cleanup the old
        If repeated > 0 And repeated <> prev And repeated <> 1 Then
            ' just repeat the previous message when only one repeat occurred
            If repeated > 2 Then
                outputBuffer = va("last message repeated %u times (total %u repeats)", repeated - prev, repeated)
            Else
                outputBuffer = va("last message repeated %u times", repeated - prev)
            End If
            printToDebugCallbacks outputBuffer, part
        End If
        repeated = 0
        next_ = 2
        prev = 0
    End If

    If repeated = 0 Then
        ourtime = strftime("%H:%M:%S", getLocalTime(time_(), True))

        ' Assemble the outputBuffer:
        outputBuffer = Left$(va("%-8s|%s: %s", code_part_names(part), ourtime, inputBuffer(cur)), MAX_LEN_LOG_LINE - 1)

        printToDebugCallbacks outputBuffer, part

        If part = LOG_ERROR Then
            ' used to signal user that there was a error condition, and to check the logs.
            errorStore = Left$(inputBuffer(cur), 511)
            errorWaiting = True
        End If

        ' Throw up a dialog box for users since most don't have a clue to check the dump file for information.
        ' Use for (duh) Fatal errors, that force us to terminate the game.
        If part = LOG_FATAL Then
            debugDisplayFatalErrorMsgBox inputBuffer(cur)
        End If

        ' Throw up a dialog box for windows users since most don't have a clue to check the stderr.txt file for information
        ' This is a popup dialog used for times when the error isn't fatal, but we still need to notify user what is going on.
        If part = LOG_POPUP Then
            wzDisplayDialog Dialog_Information, "Warzone has detected a problem.", inputBuffer(cur)
        End If
    End If
    useInputBuffer1 = Not useInputBuffer1 ' Swap buffers
End Sub

' C: the debug(part, ...) macro: logs when the part is enabled.
Public Sub debug_(ByVal line_ As Long, ByVal part As Long, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    If enabled_debug(part) Then
        a = args
        debugV line_, part, function_, str_, a
    End If
End Sub

' C: _debug (line, part, function, str, ...): logs whether or not the part is enabled.
Public Sub debug__(ByVal line_ As Long, ByVal part As Long, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    debugV line_, part, function_, str_, a
End Sub

' C: the wz_info(...) macro
Public Sub wz_info(ByVal line_ As Long, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    debugV line_, LOG_INFO, function_, str_, a
End Sub

' C: _debug_multiline (line, part, function, multilineString): one log line per line of the string.
Public Sub debug_multiline_(ByVal line_ As Long, ByVal part As Long, ByVal function_ As String, ByVal multilineString As String)
    Dim p As Long, q As Long
    ' std::getline: a last line without '\n' counts, an empty one after the last '\n' does not
    p = 1
    Do While p <= Len(multilineString)
        q = InStr(p, multilineString, vbLf)
        If q = 0 Then q = Len(multilineString) + 1
        debugV line_, part, function_, "%s", Array(Mid$(multilineString, p, q - p))
        p = q + 1
    Loop
End Sub

' C: the debug_multiline(part, string) macro
Public Sub debug_multiline(ByVal line_ As Long, ByVal part As Long, ByVal function_ As String, ByVal multilineString As String)
    If enabled_debug(part) Then debug_multiline_ line_, part, function_, multilineString
End Sub

' C: ASSERT_FAILURE (expr, expr_string, location_description, function, ...): location is "file:line"
' (AT_MACRO), and its line is the __LINE__ of the two log lines.
Public Sub ASSERT_FAILUREV(ByVal expr_string As String, ByVal location_description As String, ByVal function_ As String, ByVal str_ As String, args As Variant)
    ASSERT_HELPER_FAILUREV Val(Mid$(location_description, InStrRev(location_description, ":") + 1)), expr_string, location_description, function_, str_, args
End Sub

' C: ASSERT_HELPER (expr, location_description, function, ...) failing where the location and
' function are passed on (checkDroid's CHECK_DROID callers): the log lines carry the __LINE__ of the
' ASSERT_HELPER call, line_, not the line in location_description.
Public Sub ASSERT_HELPER_FAILUREV(ByVal line_ As Long, ByVal expr_string As String, ByVal location_description As String, ByVal function_ As String, ByVal str_ As String, args As Variant)
    debugV line_, LOG_INFO, function_, str_, args
    debugV line_, LOG_INFO, function_, "Assert in Warzone: %s (%s), last script event: '%s'", _
        Array(location_description, expr_string, last_called_script_event)
    ' wz_assert(expr) when assertEnabled: never in a release build
End Sub

Public Sub ASSERT_FAILURE(ByVal expr_string As String, ByVal location_description As String, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    ASSERT_FAILUREV expr_string, location_description, function_, str_, a
End Sub

Public Function debugPartEnabled(ByVal codePart As Long) As Boolean
    debugPartEnabled = enabled_debug(codePart)
End Function

Public Sub debugDisableAssert()
    assertEnabled = False
End Sub
