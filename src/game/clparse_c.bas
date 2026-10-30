Attribute VB_Name = "clparse_c"
' src/clparse.cpp: the command line, with clparse.cpp's own replacement for
' popt (a context UDT stands for its static struct; a poptContext is the
' struct, passed ByRef). qFatal prints to stderr and ends the process with
' code 1, like the macro. Translations (_(), N_(), gettext) are not ported:
' the English text is used.
' VB-only options for scripted test runs, in the table so that
' ParseCommandLineEarly accepts them: --frames=<n> (quit after n frames,
' and log dialogs instead of showing them) and --nodialogs.
Option Explicit

Private Type COORD_
    X As Integer
    Y As Integer
End Type

Private Type SMALL_RECT
    Left_ As Integer
    Top As Integer
    Right_ As Integer
    Bottom As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD_
    dwCursorPosition As COORD_
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD_
End Type

Private Declare Function AttachConsole Lib "kernel32" (ByVal dwProcessId As Long) As Long
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function SetConsoleTitleW Lib "kernel32" (ByVal lpConsoleTitle As Long) As Long
Private Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function SetConsoleScreenBufferSize Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal dwSize As Long) As Long

Private Const ATTACH_PARENT_PROCESS As Long = -1
Private Const STD_OUTPUT_HANDLE As Long = -11
Private Const CP_UTF8 As Long = 65001

'////
' Our fine replacement for the popt abomination follows

Private Const POPT_ARG_STRING As Boolean = True
Private Const POPT_ARG_NONE As Boolean = False
Private Const POPT_ERROR_BADOPT As Long = -1
Private Const POPT_SKIP_MAC_PSN As Long = 666

Private Const PATH_MAX As Long = 260            ' MAX_PATH on Windows

Private Type poptOption
    strval As String                ' C: string; vbNullString ends the table
    argument As Boolean
    enumeration As Long
    descrip As String               ' vbNullString = nullptr
    argDescrip As String
End Type

' C: struct _poptContext
Private Type poptContext_
    argc As Long
    current As Long
    size As Long
    argv() As String
    parameter As String             ' vbNullString = nullptr
    bad As String
    table() As poptOption
End Type

Public Enum ParseCLIEarlyResult
    ParseCLIEarlyResult_OK_CONTINUE
    ParseCLIEarlyResult_HANDLED_QUIT_EARLY_COMMAND
End Enum

' Enable automatic test games
Private wz_autogame As Boolean
Private wz_saveandquit As String
Private wz_test As String
Private wz_cli_headless As Boolean
Private wz_streamer_spectator_mode As Boolean
Private wz_lobby_slashcommands As Boolean
Private wz_min_autostart_players As Long
Private wz_lobby_game_to_connect_str As String
Private staticsReady As Boolean

Private has_setup_console As Boolean        ' SetStdOutToConsole_Win's static

Private Enum CLI_OPTIONS
    ' We don't want to use zero, so start at one (1)
    CLI_CONFIGDIR = 1
    CLI_DATADIR
    CLI_DEBUG
    CLI_DEBUGFILE
    CLI_FLUSHDEBUGSTDERR
    CLI_FULLSCREEN
    CLI_GAME
    CLI_HELP
    CLI_MOD_GLOB
    CLI_MOD_CA
    CLI_MOD_MP
    CLI_LOADSKIRMISH
    CLI_LOADCAMPAIGN
    CLI_LOADREPLAY
    CLI_WINDOW
    CLI_VERSION
    CLI_RESOLUTION
    CLI_SHADOWS
    CLI_NOSHADOWS
    CLI_SOUND
    CLI_NOSOUND
    CLI_CONNECTTOIP
    CLI_CONNECTTOIP_SPECTATE
    CLI_CONNECTTOLOBBYGAME
    CLI_CONNECTTOLOBBYGAME_SPECTATE
    CLI_HOSTLAUNCH
    CLI_NOASSERT
    CLI_CRASH
    CLI_TEXTURECOMPRESSION
    CLI_NOTEXTURECOMPRESSION
    CLI_GFXBACKEND
    CLI_GFXDEBUG
    CLI_JSBACKEND
    CLI_AUTOGAME
    CLI_SAVEANDQUIT
    CLI_SKIRMISH
    CLI_CONTINUE
    CLI_AUTOHOST
    CLI_AUTOHEADLESS
    CLI_WIN_ENABLE_CONSOLE
    CLI_GAMEPORT
    CLI_NET_PORTMAPPING
    CLI_WZ_CRASH_RPT
    CLI_WZ_DEBUG_CRASH_HANDLER
    CLI_STREAMER_SPECTATOR
    CLI_LOBBY_SLASHCOMMANDS
    CLI_ADD_LOBBY_ADMINHASH
    CLI_ADD_LOBBY_ADMINPUBLICKEY
    CLI_COMMAND_INTERFACE
    CLI_STARTPLAYERS
    CLI_GAMELOG_OUTPUTMODES
    CLI_GAMELOG_OUTPUTKEY
    CLI_GAMELOG_OUTPUTNAMING
    CLI_GAMELOG_FRAMEINTERVAL
    CLI_GAMETIMELIMITMINUTES
    CLI_CONVERT_SPECULAR_MAP
    CLI_DEBUG_VERBOSE_SYNCLOG_OUTPUT
    CLI_ALLOW_VULKAN_IMPLICIT_LAYERS
    CLI_HOST_CHAT_CONFIG
    CLI_HOST_ASYNC_JOIN_APPROVAL
    CLI_AUTOHOST_START_NOT_READY
    CLI_HOST_CONNECTION_PROVIDER
    CLI_VB_FRAMES                   ' (VB)
    CLI_VB_NODIALOGS                ' (VB)
    CLI_VB_TEST                     ' (VB)
    CLI_VB_DRAWTEST                 ' (VB)
    CLI_VB_WATCHSHOT                ' (VB)
    CLI_VB_SCRIPT                   ' (VB)
    CLI_VB_SNDCAPTURE               ' (VB)
End Enum

Private debugOptionsTable() As poptOption
Private optionsTable() As poptOption
Private tablesReady As Boolean

' C: #define qFatal(...) { fprintf(stderr, __VA_ARGS__); fprintf(stderr, "\n"); exit(1); }
Private Sub qFatal(ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    fputs vsprintf(fmt, a), STDERR_FILE
    fprintf STDERR_FILE, "\n"
    exit_ 1
End Sub

' the C++ initialisers of the statics
Private Sub InitStatics()
    If staticsReady Then Exit Sub
    staticsReady = True
    wz_min_autostart_players = -1
End Sub

Private Sub SetStdOutToConsole_Win()
    Const MAX_CONSOLE_LINES As Long = 9999
    Dim consoleHandle As Long, csbi As CONSOLE_SCREEN_BUFFER_INFO

    If has_setup_console Then Exit Sub

    If AttachConsole(ATTACH_PARENT_PROCESS) = 0 Then
        ' failed to attach to parent process console
        ' allocate a console for this app
        If AllocConsole() = 0 Then
            ' failed to allocate a console
            Exit Sub
        End If

        ' give the new console window a nicer title
        SetConsoleTitleW StrPtr("Warzone 2100")
    End If

    SetConsoleOutputCP CP_UTF8

    consoleHandle = GetStdHandle(STD_OUTPUT_HANDLE)

    ' give the console window a bigger buffer size / scroll-back
    If GetConsoleScreenBufferInfo(consoleHandle, csbi) <> 0 Then
        csbi.dwSize.Y = MAX_CONSOLE_LINES
        SetConsoleScreenBufferSize consoleHandle, (csbi.dwSize.X And &HFFFF&) Or (CLng(csbi.dwSize.Y) * &H10000)
    End If

    freopen "CONOUT$", "w", STDOUT_FILE
    freopen "CONOUT$", "w", STDERR_FILE

    has_setup_console = True
End Sub

Private Sub poptPrintHelp(ctx As poptContext_, ByVal output As Long)
    Dim i As Long, txt As String, txtSize As Long, txtOffset As Long, b() As Byte

    ' TRANSLATORS: Summary of commandline option syntax
    fprintf output, "Usage: %s [OPTION...]\n", ctx.argv(0)
    For i = 0 To ctx.size - 1
        txt = Left$(va("  --%s", ctx.table(i).strval), 127)    ' char txt[128]

        If ctx.table(i).argument Then
            txt = Left$(txt & "=", 127)
            If StrPtr(ctx.table(i).argDescrip) <> 0 Then
                txt = Left$(txt & ctx.table(i).argDescrip, 127)
            End If
        End If

        ' calculate number of terminal columns required to print
        ' for languages with multibyte characters
        txtSize = utf8_encode(txt, b) + 1
        txtOffset = Len(txt) + 1 - txtSize                  ' mbsrtowcs: the characters
        ' CJK characters take up two columns (the language: no translations here, so never zh or ko)

        fprintf output, "%-*s", 40 - txtOffset, txt

        If StrPtr(ctx.table(i).descrip) <> 0 Then
            fprintf output, "%s", ctx.table(i).descrip
        End If
        fprintf output, "\n"
    Next
End Sub

Private Function poptBadOption(ctx As poptContext_, ByVal unused As Long) As String
    poptBadOption = ctx.bad
End Function

Private Function poptGetOptArg(ctx As poptContext_) As String
    poptGetOptArg = ctx.parameter
End Function

Private Function poptGetNextOpt(ctx As poptContext_) As Long
    Dim match As String, parameter As String, pparam As String, hasParam As Boolean
    Dim i As Long, p As Long, slong As String

    ctx.bad = vbNullString
    ctx.parameter = vbNullString
    parameter = ""
    match = ""

    If ctx.current >= ctx.argc Then ' counts from 1
        poptGetNextOpt = 0
        Exit Function
    End If

    If InStr(ctx.argv(ctx.current), "-psn_") > 0 Then
        ctx.current = ctx.current + 1 ' skip mac -psn_*  Yum!
        poptGetNextOpt = POPT_SKIP_MAC_PSN
        Exit Function
    End If

    match = Left$(ctx.argv(ctx.current), PATH_MAX - 1)
    ctx.current = ctx.current + 1
    p = InStr(match, "=")
    hasParam = (p > 0)
    If hasParam Then                                    ' option's got a parameter
        pparam = Mid$(match, p + 1)
        match = Left$(match, p - 1)                     ' split option from parameter and increment past '='
        If Left$(pparam, 1) = """" Then                 ' found scary quotes
            pparam = Mid$(pparam, 2)                    ' skip start quote
            parameter = Left$(pparam, PATH_MAX - 1)     ' copy first parameter
            If InStrRev(pparam, """") = 0 Then          ' if no end quote, then find it
                Do While InStrRev(parameter, """") = 0 And ctx.current < ctx.argc
                    parameter = Left$(parameter & " ", PATH_MAX - 1)    ' insert space
                    parameter = Left$(parameter & ctx.argv(ctx.current), PATH_MAX - 1)
                    ctx.current = ctx.current + 1       ' next part, please!
                Loop
            End If
            p = InStrRev(parameter, """")
            If p > 0 Then                               ' its not an else for above!
                parameter = Left$(parameter, p - 1)     ' remove end qoute
            End If
        Else
            parameter = Left$(pparam, PATH_MAX - 1)     ' copy parameter
        End If
    End If

    For i = 0 To ctx.size - 1
        slong = Left$("--" & ctx.table(i).strval, 63)
        If slong = match Then
            If ctx.table(i).argument And hasParam Then
                ctx.parameter = parameter
                If StrPtr(ctx.parameter) = 0 Then ctx.parameter = ""   ' (VB) an empty parameter is not nullptr
            End If
            poptGetNextOpt = ctx.table(i).enumeration
            Exit Function
        End If
    Next
    ctx.bad = match
    poptGetNextOpt = POPT_ERROR_BADOPT
End Function

Private Sub poptGetContext(ctx As poptContext_, ByVal argc As Long, argv() As String, table() As poptOption)
    ctx.argc = argc
    ctx.argv = argv
    ctx.table = table
    ctx.current = 1
    ctx.parameter = vbNullString

    ctx.size = 0
    Do While StrPtr(table(ctx.size).strval) <> 0       ' count table size
        ctx.size = ctx.size + 1
    Loop
End Sub

Private Sub AddOption(t() As poptOption, n As Long, ByVal strval As String, ByVal argument As Boolean, ByVal enumeration As CLI_OPTIONS, ByVal descrip As String, ByVal argDescrip As String)
    If n > UBound(t) Then ReDim Preserve t(0 To 2 * n)
    t(n).strval = strval
    t(n).argument = argument
    t(n).enumeration = enumeration
    t(n).descrip = descrip
    t(n).argDescrip = argDescrip
    n = n + 1
End Sub

' The two tables, each with its terminating entry (strval = vbNullString)
Private Sub InitTables()
    Dim n As Long
    If tablesReady Then Exit Sub
    tablesReady = True

    ' Separate table that avoids *any* translated strings, to avoid any risk of gettext / libintl function calls
    ReDim debugOptionsTable(0 To 3)
    n = 0
    AddOption debugOptionsTable, n, "debug", POPT_ARG_STRING, CLI_DEBUG, vbNullString, vbNullString
    AddOption debugOptionsTable, n, "debugfile", POPT_ARG_STRING, CLI_DEBUGFILE, vbNullString, vbNullString
    AddOption debugOptionsTable, n, "flush-debug-stderr", POPT_ARG_NONE, CLI_FLUSHDEBUGSTDERR, vbNullString, vbNullString
    ' Terminating entry
    AddOption debugOptionsTable, n, vbNullString, False, 0, vbNullString, vbNullString

    ReDim optionsTable(0 To 63)
    n = 0
    AddOption optionsTable, n, "configdir", POPT_ARG_STRING, CLI_CONFIGDIR, "Set configuration directory", "configuration directory"
    AddOption optionsTable, n, "datadir", POPT_ARG_STRING, CLI_DATADIR, "Add data directory", "data directory"
    AddOption optionsTable, n, "debug", POPT_ARG_STRING, CLI_DEBUG, "Show debug for given level", "debug level"
    AddOption optionsTable, n, "debugfile", POPT_ARG_STRING, CLI_DEBUGFILE, "Log debug output to file", "file"
    AddOption optionsTable, n, "flush-debug-stderr", POPT_ARG_NONE, CLI_FLUSHDEBUGSTDERR, "Flush all debug output written to stderr", vbNullString
    AddOption optionsTable, n, "fullscreen", POPT_ARG_NONE, CLI_FULLSCREEN, "Play in fullscreen mode", vbNullString
    AddOption optionsTable, n, "game", POPT_ARG_STRING, CLI_GAME, "Load a specific game mode", "level name"
    AddOption optionsTable, n, "help", POPT_ARG_NONE, CLI_HELP, "Show options and exit", vbNullString
    AddOption optionsTable, n, "mod", POPT_ARG_STRING, CLI_MOD_GLOB, "Enable a global mod", "mod"
    AddOption optionsTable, n, "mod_ca", POPT_ARG_STRING, CLI_MOD_CA, "Enable a campaign only mod", "mod"
    AddOption optionsTable, n, "mod_mp", POPT_ARG_STRING, CLI_MOD_MP, "Enable a multiplay only mod", "mod"
    AddOption optionsTable, n, "noassert", POPT_ARG_NONE, CLI_NOASSERT, "Disable asserts", vbNullString
    AddOption optionsTable, n, "crash", POPT_ARG_NONE, CLI_CRASH, "Causes a crash to test the crash handler", vbNullString
    AddOption optionsTable, n, "loadskirmish", POPT_ARG_STRING, CLI_LOADSKIRMISH, "Load a saved skirmish game", "savegame"
    AddOption optionsTable, n, "loadcampaign", POPT_ARG_STRING, CLI_LOADCAMPAIGN, "Load a saved campaign game", "savegame"
    AddOption optionsTable, n, "loadreplay", POPT_ARG_STRING, CLI_LOADREPLAY, "Load a replay", "replay file"
    AddOption optionsTable, n, "window", POPT_ARG_NONE, CLI_WINDOW, "Play in windowed mode", vbNullString
    AddOption optionsTable, n, "version", POPT_ARG_NONE, CLI_VERSION, "Show version information and exit", vbNullString
    AddOption optionsTable, n, "resolution", POPT_ARG_STRING, CLI_RESOLUTION, "Set the resolution to use", "WIDTHxHEIGHT"
    AddOption optionsTable, n, "shadows", POPT_ARG_NONE, CLI_SHADOWS, "Enable shadows", vbNullString
    AddOption optionsTable, n, "noshadows", POPT_ARG_NONE, CLI_NOSHADOWS, "Disable shadows", vbNullString
    AddOption optionsTable, n, "sound", POPT_ARG_NONE, CLI_SOUND, "Enable sound", vbNullString
    AddOption optionsTable, n, "nosound", POPT_ARG_NONE, CLI_NOSOUND, "Disable sound", vbNullString
    AddOption optionsTable, n, "join", POPT_ARG_STRING, CLI_CONNECTTOIP, "Connect directly to IP/hostname", "host"
    AddOption optionsTable, n, "spectate", POPT_ARG_STRING, CLI_CONNECTTOIP_SPECTATE, "Connect directly to IP/hostname as a spectator", "host"
    AddOption optionsTable, n, "lobbyjoin", POPT_ARG_STRING, CLI_CONNECTTOLOBBYGAME, "Connect to lobby game", "lobbyGameId"
    AddOption optionsTable, n, "lobbyspectate", POPT_ARG_STRING, CLI_CONNECTTOLOBBYGAME_SPECTATE, "Connect to lobby game as a spectator", "lobbyGameId"
    AddOption optionsTable, n, "host", POPT_ARG_NONE, CLI_HOSTLAUNCH, "Go directly to host screen", vbNullString
    AddOption optionsTable, n, "texturecompression", POPT_ARG_NONE, CLI_TEXTURECOMPRESSION, "Enable texture compression", vbNullString
    AddOption optionsTable, n, "notexturecompression", POPT_ARG_NONE, CLI_NOTEXTURECOMPRESSION, "Disable texture compression", vbNullString
    AddOption optionsTable, n, "gfxbackend", POPT_ARG_STRING, CLI_GFXBACKEND, "Set gfx backend", _
        "(opengl, opengles" & _
        ", vulkan" & _
        ")"                                             ' WZ_VULKAN_ENABLED, as in the Windows build
    AddOption optionsTable, n, "gfxdebug", POPT_ARG_NONE, CLI_GFXDEBUG, "Use gfx backend debug", vbNullString
    AddOption optionsTable, n, "jsbackend", POPT_ARG_STRING, CLI_JSBACKEND, "Set JS backend", _
        "(" & _
        "quickjs" & _
        ")"
    AddOption optionsTable, n, "autogame", POPT_ARG_NONE, CLI_AUTOGAME, "Run games automatically for testing", vbNullString
    AddOption optionsTable, n, "headless", POPT_ARG_NONE, CLI_AUTOHEADLESS, "Headless mode (only supported when also specifying --autogame, --autohost, --skirmish)", vbNullString
    AddOption optionsTable, n, "saveandquit", POPT_ARG_STRING, CLI_SAVEANDQUIT, "Immediately save game and quit", "save name"
    AddOption optionsTable, n, "skirmish", POPT_ARG_STRING, CLI_SKIRMISH, "Start skirmish game with given settings file", "test"
    AddOption optionsTable, n, "continue", POPT_ARG_NONE, CLI_CONTINUE, "Continue the last saved game", vbNullString
    AddOption optionsTable, n, "autohost", POPT_ARG_STRING, CLI_AUTOHOST, "Start host game with given settings file", "autohost"
    AddOption optionsTable, n, "enableconsole", POPT_ARG_NONE, CLI_WIN_ENABLE_CONSOLE, "Attach or create a console window and display console output (Windows only)", vbNullString
    AddOption optionsTable, n, "gameport", POPT_ARG_STRING, CLI_GAMEPORT, "Set game server port", "port"
    AddOption optionsTable, n, "portmapping", POPT_ARG_STRING, CLI_NET_PORTMAPPING, "Enable / disable port mapping when hosting", "[1, true, 0, false]"
    AddOption optionsTable, n, "wz-crash-rpt", POPT_ARG_NONE, CLI_WZ_CRASH_RPT, vbNullString, vbNullString
    AddOption optionsTable, n, "wz-debug-crash-handler", POPT_ARG_NONE, CLI_WZ_DEBUG_CRASH_HANDLER, vbNullString, vbNullString
    AddOption optionsTable, n, "spectator-min-ui", POPT_ARG_NONE, CLI_STREAMER_SPECTATOR, vbNullString, vbNullString
    AddOption optionsTable, n, "enablelobbyslashcmd", POPT_ARG_NONE, CLI_LOBBY_SLASHCOMMANDS, "Enable lobby slash commands (for connecting clients)", vbNullString
    AddOption optionsTable, n, "addlobbyadminhash", POPT_ARG_STRING, CLI_ADD_LOBBY_ADMINHASH, "Add a lobby admin identity hash (for slash commands)", "hash string"
    AddOption optionsTable, n, "addlobbyadminpublickey", POPT_ARG_STRING, CLI_ADD_LOBBY_ADMINPUBLICKEY, "Add a lobby admin public key (for slash commands)", "b64-pub-key"
    AddOption optionsTable, n, "enablecmdinterface", POPT_ARG_STRING, CLI_COMMAND_INTERFACE, "Enable command interface", "(stdin, unixsocket:path)"
    AddOption optionsTable, n, "startplayers", POPT_ARG_STRING, CLI_STARTPLAYERS, "Minimum required players to auto-start game", "startplayers"
    AddOption optionsTable, n, "gamelog-output", POPT_ARG_STRING, CLI_GAMELOG_OUTPUTMODES, "Game history log output mode(s)", "(log,cmdinterface)"
    AddOption optionsTable, n, "gamelog-outputkey", POPT_ARG_STRING, CLI_GAMELOG_OUTPUTKEY, "Game history log output key", "[playerindex, playerposition]"
    AddOption optionsTable, n, "gamelog-outputnaming", POPT_ARG_STRING, CLI_GAMELOG_OUTPUTNAMING, "Game history log output naming", "[default, autohosterclassic]"
    AddOption optionsTable, n, "gamelog-frameinterval", POPT_ARG_STRING, CLI_GAMELOG_FRAMEINTERVAL, "Game history log frame interval", "interval in seconds"
    AddOption optionsTable, n, "gametimelimit", POPT_ARG_STRING, CLI_GAMETIMELIMITMINUTES, "Multiplayer game time limit (in minutes)", "number of minutes"
    AddOption optionsTable, n, "convert-specular-map", POPT_ARG_STRING, CLI_CONVERT_SPECULAR_MAP, "Convert a specular-map .png to a luma, single-channel, grayscale .png (and exit)", "inputpath/filename.png:outputpath/filename.png"
    AddOption optionsTable, n, "debug-verbose-sync-logs-until", POPT_ARG_STRING, CLI_DEBUG_VERBOSE_SYNCLOG_OUTPUT, vbNullString, vbNullString
    AddOption optionsTable, n, "allow-vulkan-implicit-layers", POPT_ARG_NONE, CLI_ALLOW_VULKAN_IMPLICIT_LAYERS, "Allow Vulkan implicit layers (that may be default-disabled due to potential crashes or bugs)", vbNullString
    AddOption optionsTable, n, "host-chat-config", POPT_ARG_STRING, CLI_HOST_CHAT_CONFIG, "Set the default hosting chat configuration / permissions", "[allow,quickchat]"
    AddOption optionsTable, n, "async-join-approve", POPT_ARG_NONE, CLI_HOST_ASYNC_JOIN_APPROVAL, "Enable async join approval (for connecting clients)", vbNullString
    AddOption optionsTable, n, "autohost-not-ready", POPT_ARG_NONE, CLI_AUTOHOST_START_NOT_READY, "Starts the host (autohost) as not ready, even if it's a spectator host", vbNullString
    AddOption optionsTable, n, "host-connection-provider", POPT_ARG_STRING, CLI_HOST_CONNECTION_PROVIDER, "Specify connection provider type to use when hosting game sessions", "[tcp]"
    AddOption optionsTable, n, "frames", POPT_ARG_STRING, CLI_VB_FRAMES, "(VB) Quit after this many frames, logging dialogs instead of showing them", "frames"
    AddOption optionsTable, n, "nodialogs", POPT_ARG_NONE, CLI_VB_NODIALOGS, "(VB) Log dialogs instead of showing them", vbNullString
    AddOption optionsTable, n, "vbtest", POPT_ARG_STRING, CLI_VB_TEST, "(VB) Run the test commands of this file (tools\run_tests.py) and quit", "file"
    AddOption optionsTable, n, "vbdrawtest", POPT_ARG_STRING, CLI_VB_DRAWTEST, "(VB) Draw the renderer test over the title screen, save a screenshot named so and quit (tools\draw_test.py)", "name"
    AddOption optionsTable, n, "vbwatchshot", POPT_ARG_STRING, CLI_VB_WATCHSHOT, "(VB) In a game, log the frame rate, save a screenshot named so at this game time and quit (tools\watch_test.py)", "ms,name"
    AddOption optionsTable, n, "vbscript", POPT_ARG_STRING, CLI_VB_SCRIPT, "(VB) Run this test script of mouse clicks, keys, waits and screenshots on the game's frames (tools\ui_test.py)", "file"
    AddOption optionsTable, n, "vbsndcapture", POPT_ARG_STRING, CLI_VB_SNDCAPTURE, "(VB) Write the mix of the sound to this WAV file instead of playing it (no sound device opens)", "file"

    ' Terminating entry
    AddOption optionsTable, n, vbNullString, False, 0, vbNullString, vbNullString
End Sub

' C: the options table, its descriptions translated with gettext (English here)
Private Function getOptionsTable() As poptOption()
    InitTables
    getOptionsTable = optionsTable
End Function

' Must not trigger or call any gettext / libintl routines!
Public Function ParseCommandLineDebugFlags(ByVal argc As Long, argv() As String) As Boolean
    Dim poptCon As poptContext_, iOption As Long, token As String, debug_filename As String

    InitTables
    poptGetContext poptCon, argc, argv, debugOptionsTable

    ' loop through command line
    iOption = poptGetNextOpt(poptCon)
    Do While iOption > 0 Or iOption = POPT_ERROR_BADOPT
        If iOption <> POPT_ERROR_BADOPT Then
            Select Case iOption
            Case CLI_DEBUG
                ' retrieve the debug section name
                token = poptGetOptArg(poptCon)
                If StrPtr(token) = 0 Then
                    qFatal "Usage: --debug=<flag>"
                End If

                ' Attempt to enable the given debug section
                If Not debug_enable_switch(token) Then
                    qFatal "Debug flag ""%s"" not found!", token
                End If

            Case CLI_DEBUGFILE
                ' find the file name
                token = poptGetOptArg(poptCon)
                If StrPtr(token) = 0 Then
                    qFatal "Missing debugfile filename?"
                End If
                debug_filename = token
                debug_register_callback CB_debug_callback_file, CB_debug_callback_file_init, CB_debug_callback_file_exit, debug_filename ' note: by the time this function returns, all use of debug_filename has completed
                customDebugfile = True

            Case CLI_FLUSHDEBUGSTDERR
                ' Tell the debug stderr output callback to always flush its output
                debugFlushStderr
            End Select
        End If
        iOption = poptGetNextOpt(poptCon)
    Loop

    ParseCommandLineDebugFlags = True
End Function

Private Function specialGetBaseDir(ByVal platformSpecificPath As String, output_filename As String) As String
    Dim result As String, dirSeparator As String, lastSlash As Long, pBaseDir As String
    result = platformSpecificPath
    dirSeparator = PHYSFS_getDirSeparator()
    Do While Len(result) > 0
        If Right$(result, Len(dirSeparator)) <> dirSeparator Then Exit Do
        result = Left$(result, Len(result) - Len(dirSeparator)) ' Remove trailing path separators
    Loop
    lastSlash = InStrRev(result, dirSeparator)
    If lastSlash > 0 Then
        output_filename = Mid$(result, lastSlash + 1)
        specialGetBaseDir = Left$(result, lastSlash - 1) ' Trim off the last path component
    Else
        ' no dir ahead of path
        output_filename = platformSpecificPath
        ' use PhysFS_BaseDir
        pBaseDir = PHYSFS_getBaseDir()
        If StrPtr(pBaseDir) <> 0 Then specialGetBaseDir = pBaseDir Else specialGetBaseDir = ""
    End If
End Function

'! Early parsing of the commandline
' First half of the command line parsing. Also see ParseCommandLine()
' below. The parameters here are needed early in the boot process,
' while the ones in ParseCommandLine can benefit from debugging being
' set up first.
Public Function ParseCommandLineEarly(ByVal argc As Long, argv() As String) As ParseCLIEarlyResult
    Dim poptCon As poptContext_, iOption As Long, token As String, b() As Byte
    Dim fullArg As String, firstDelimiter As Long, expectedInputPathExample As String, expectedOutputPathExample As String
    Dim inputPath As String, outputPath As String, inputFilename As String, outputFilename As String
    Dim inputDir As String, outputDir As String

    poptGetContext poptCon, argc, argv, getOptionsTable()

    ' loop through command line
    iOption = poptGetNextOpt(poptCon)
    Do While iOption > 0 Or iOption = POPT_ERROR_BADOPT
        If iOption = POPT_ERROR_BADOPT Then
            qFatal "Unrecognized option: %s", poptBadOption(poptCon, 0)
        End If

        Select Case iOption
        Case CLI_CONFIGDIR
            ' retrieve the configuration directory
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Unrecognised configuration directory"
            End If
            If utf8_encode(token, b) >= PATH_MAX Then      ' strlen(token) >= sizeof(configdir)
                qFatal "Configuration directory exceeds maximum supported length on this platform"
            End If
            configdir = token

        Case CLI_HELP
            poptPrintHelp poptCon, STDOUT_FILE
            ParseCommandLineEarly = ParseCLIEarlyResult_HANDLED_QUIT_EARLY_COMMAND
            Exit Function

        Case CLI_VERSION
            fprintf STDOUT_FILE, "Warzone 2100 - %s\n", version_getFormattedVersionString()
            ParseCommandLineEarly = ParseCLIEarlyResult_HANDLED_QUIT_EARLY_COMMAND
            Exit Function

        Case CLI_WIN_ENABLE_CONSOLE
            SetStdOutToConsole_Win

        Case CLI_WZ_CRASH_RPT, CLI_WZ_DEBUG_CRASH_HANDLER
            ' this is currently a no-op because it must be parsed even earlier than ParseCommandLineEarly

        Case CLI_CONVERT_SPECULAR_MAP
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Missing convert-specular-map value"
            ElseIf Len(token) = 0 Then
                qFatal "Missing convert-specular-map value"
            End If
            ' Should be a string with input and output filenames in the format:
            ' inputpath/filename.png:outputpath/filename.png
            ' (Where "/" is the platform path separator - on Windows, this would be "\")
            fullArg = token
            firstDelimiter = InStr(fullArg, ":") - 1           ' C: index, or npos (-1 here)
            If firstDelimiter = -1 Or Not (firstDelimiter + 1 < Len(fullArg)) Then
                expectedInputPathExample = "inputpath" & PHYSFS_getDirSeparator() & "filename.png"
                expectedOutputPathExample = "outputpath" & PHYSFS_getDirSeparator() & "filename.png"
                qFatal "Invalid convert-specular-map value - expecting format: %s:%s", expectedInputPathExample, expectedOutputPathExample
            End If

            inputPath = Left$(fullArg, firstDelimiter)
            outputPath = Mid$(fullArg, firstDelimiter + 2)

            inputDir = specialGetBaseDir(inputPath, inputFilename)
            outputDir = specialGetBaseDir(outputPath, outputFilename)

            ' This is a bit of a hack, but set up PhysFS to use:
            ' - the path containing the inputFilename as the read path
            ' - the outputFilename path as the write path

            If Len(inputDir) = 0 Then
                qFatal "convert-specular-map value does not seem to include the path to a file (including its directory)"
            End If
            If Len(outputDir) = 0 Then
                qFatal "convert-specular-map value does not seem to include an output path"
            End If

            If PHYSFS_setWriteDir(outputDir) = 0 Then
                qFatal "convert-specular-map - unable to configure output directory to: %s", outputDir
            End If

            ' Output dir first so we can see what we write
            PHYSFS_mount PHYSFS_getWriteDir(), "", PHYSFS_PREPEND

            PHYSFS_mount inputDir, "input", PHYSFS_APPEND

            If Not iV_LoadAndSavePNG_AsLumaSingleChannel("input/" & inputFilename, outputFilename, True) Then
                qFatal "convert-specular-map - failed to convert image: %s", inputDir
            End If

            PHYSFS_deinit
            exit_ 0
        End Select

        iOption = poptGetNextOpt(poptCon)
    Loop

    ParseCommandLineEarly = ParseCLIEarlyResult_OK_CONTINUE
End Function

' C: optional<bool> cliOptValueToBool(const char *token): has_value, and the value ByRef
Private Function cliOptValueToBool(ByVal token As String, value As Boolean) As Boolean
    If StrPtr(token) = 0 Then
        Exit Function
    End If

    If token = "1" Or token = "true" Then
        value = True
        cliOptValueToBool = True
        Exit Function
    End If
    If token = "0" Or token = "false" Then
        value = False
        cliOptValueToBool = True
        Exit Function
    End If
End Function

' C: sscanf(token, "%ux%u", &width, &height): the number of fields read (-1 for empty input)
Private Function ScanResolution(ByVal token As String, width As Long, height As Long) As Long
    Dim endIdx As Long, v As Long, i As Long
    i = 1
    Do While Mid$(token, i, 1) = " " Or Mid$(token, i, 1) = vbTab
        i = i + 1
    Loop
    If i > Len(token) Then
        ScanResolution = -1                             ' EOF before the first conversion
        Exit Function
    End If
    v = strtoul(Mid$(token, i), endIdx, 10)
    If endIdx = 0 Then Exit Function
    width = v
    ScanResolution = 1
    i = i + endIdx
    If Mid$(token, i, 1) <> "x" Then Exit Function
    v = strtoul(Mid$(token, i + 1), endIdx, 10)
    If endIdx = 0 Then Exit Function
    height = v
    ScanResolution = 2
End Function

'! second half of parsing the commandline
' Second half of command line parsing. See ParseCommandLineEarly() for
' the first half. Note that render mode must come before resolution flag.
Public Function ParseCommandLine(ByVal argc As Long, argv() As String) As Boolean
    Dim poptCon As poptContext_, iOption As Long, token As String, i As Long
    Dim width As Long, height As Long, extension As String, foundReplayFile As Boolean
    Dim replaySubdirs As Variant, gfxBackend As Long, jsBackend As Long, optValue As Boolean, fields() As String
    Dim mode As Long, value As String, delimeterIdx As Long, modes As GameStoryLogger_OutputModes
    Dim params() As String, token_intval As Long, pt As Long

    InitStatics
    poptGetContext poptCon, argc, argv, getOptionsTable()

    ' loop through command line
    iOption = poptGetNextOpt(poptCon)
    Do While iOption > 0
        Select Case iOption
        Case CLI_DEBUG, CLI_DEBUGFILE, CLI_FLUSHDEBUGSTDERR, CLI_CONFIGDIR, CLI_HELP, CLI_VERSION, _
             CLI_WIN_ENABLE_CONSOLE, CLI_WZ_CRASH_RPT, CLI_WZ_DEBUG_CRASH_HANDLER, CLI_CONVERT_SPECULAR_MAP
            ' These options are parsed in ParseCommandLineEarly() already, so ignore them

        Case CLI_NOASSERT
            kf_NoAssert

        ' NOTE: The sole purpose of this is to test the crash handler.
        Case CLI_CRASH
            CauseCrash = True
            NetPlay_bComms = False
            SPinit LEVEL_TYPE_CAMPAIGN
            aLevelName = "CAM_3A"
            SetGameMode GS_NORMAL

        Case CLI_DATADIR
            ' retrieve the quoted path name
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Unrecognised datadir"
            End If
            datadir = Left$(token, PATH_MAX - 1)

        Case CLI_FULLSCREEN
            war_setWindowMode WINDOW_MODE_fullscreen

        Case CLI_CONNECTTOIP, CLI_CONNECTTOIP_SPECTATE
            'get the ip we want to connect with, and go directly to join screen.
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "No IP/hostname given"
            End If
            iptoconnect = Left$(token, PATH_MAX - 1)
            ' also set spectate flag
            cliConnectAsSpectator = (iOption = CLI_CONNECTTOIP_SPECTATE)

        Case CLI_CONNECTTOLOBBYGAME, CLI_CONNECTTOLOBBYGAME_SPECTATE
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "No lobbyGameId given"
            End If
            wz_lobby_game_to_connect_str = token
            ' also set spectate flag
            cliConnectAsSpectator = (iOption = CLI_CONNECTTOLOBBYGAME_SPECTATE)

        Case CLI_HOSTLAUNCH
            ' go directly to host screen, bypass all others.
            setHostLaunch HostLaunch_Host

        Case CLI_GAME
            ' retrieve the game name
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "The game parameter requires one of the following keywords:" & _
                       "CAM_1A, CAM_2A, CAM_3A, TUTORIAL3, or FASTPLAY."
            ElseIf token <> "CAM_1A" And token <> "CAM_2A" And token <> "CAM_3A" _
                   And token <> "TUTORIAL3" And token <> "FASTPLAY" Then
                qFatal "The game parameter requires one of the following keywords:" & _
                       "CAM_1A, CAM_2A, CAM_3A, TUTORIAL3, or FASTPLAY."
            End If
            NetPlay_bComms = False
            bMultiPlayer = False
            bMultiMessages = False
            For i = 0 To MAX_PLAYERS - 1
                NET_InitPlayer i, True
            Next

            'NET_InitPlayer deallocates Player 0, who must be allocated so that a later invocation of processDebugMappings does not trigger DEBUG mode
            NetPlay_players_allocated(0) = True

            If token = "CAM_1A" Or token = "CAM_2A" Or token = "CAM_3A" Then
                game.typ = LEVEL_TYPE_CAMPAIGN
            Else
                game.typ = LEVEL_TYPE_SKIRMISH ' tutorial is skirmish for some reason
            End If
            aLevelName = Left$(token, MAX_LEVEL_NAME_SIZE)
            SetGameMode GS_NORMAL

        Case CLI_MOD_GLOB
            ' retrieve the file name
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Missing mod name?"
            End If

            StringVector_push_back global_mods, token

        Case CLI_MOD_CA
            ' retrieve the file name
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Missing mod name?"
            End If

            StringVector_push_back campaign_mods, token

        Case CLI_MOD_MP
            ' retrieve the file name
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Missing mod name?"
            End If

            StringVector_push_back multiplay_mods, token

        Case CLI_RESOLUTION
            token = poptGetOptArg(poptCon)
            ' (C: sscanf on nullptr crashes; here it reads nothing)
            If ScanResolution(token, width, height) <> 2 Then
                qFatal "Invalid parameter specified (format is WIDTHxHEIGHT, e.g. 800x600)"
            End If
            If ULT(width, 640) Then
                debug_ 898, LOG_ERROR, "ParseCommandLine", "Screen width < 640 unsupported, using 640"
                width = 640
            End If
            If ULT(height, 480) Then
                debug_ 903, LOG_ERROR, "ParseCommandLine", "Screen height < 480 unsupported, using 480"
                height = 480
            End If
            ' tell the display system of the desired resolution
            pie_SetVideoBufferWidth width
            pie_SetVideoBufferHeight height
            ' and update the configuration
            war_SetWidth width
            war_SetHeight height

        Case CLI_LOADSKIRMISH
            ' retrieve the game name
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Unrecognised skirmish savegame name"
            End If
            saveGameName = Left$(va("%s/skirmish/%s.gam", SaveGamePath, token), 255)
            sRequestResult = Left$(saveGameName, PATH_MAX - 1) ' hack to avoid crashes
            SPinit LEVEL_TYPE_SKIRMISH
            bMultiPlayer = True
            SetGameMode GS_SAVEGAMELOAD

        Case CLI_LOADCAMPAIGN
            ' retrieve the game name
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Unrecognised campaign savegame name"
            End If
            saveGameName = Left$(va("%s/campaign/%s.gam", SaveGamePath, token), 255)
            SPinit LEVEL_TYPE_CAMPAIGN
            SetGameMode GS_SAVEGAMELOAD

        Case CLI_LOADREPLAY
            ' retrieve the replay name
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Unrecognised replay name"
            End If
            extension = ""
            If Not strEndsWith(token, ".wzrp") Then
                extension = ".wzrp"
            End If
            ' check if we have a full path (relative to the replay dir)
            saveGameName = Left$(va("%s/%s%s", ReplayPath, token, extension), 255)
            foundReplayFile = PHYSFS_exists(saveGameName) <> 0
            If Not foundReplayFile Then
                ' look in all possible replay subdirs (maybe we just have a filename)
                replaySubdirs = Array("skirmish", "multiplay")
                For i = 0 To UBound(replaySubdirs)
                    saveGameName = Left$(va("%s/%s/%s%s", ReplayPath, replaySubdirs(i), token, extension), 255)
                    If PHYSFS_exists(saveGameName) <> 0 Then
                        foundReplayFile = True
                        Exit For
                    End If
                Next
            End If
            If Not foundReplayFile Then
                qFatal "Unable to find specified replay"
            End If
            setHostLaunch HostLaunch_LoadReplay
            sRequestResult = Left$(saveGameName, PATH_MAX - 1) ' hack to avoid crashes
            SPinit LEVEL_TYPE_SKIRMISH
            bMultiPlayer = True
            game.maxPlayers = 4 'DEFAULTSKIRMISHMAPMAXPLAYERS;
            SetGameMode GS_SAVEGAMELOAD

        Case CLI_CONTINUE
            If findLastSave() Then
                runContinue
                SetGameMode GS_SAVEGAMELOAD
            End If

        Case CLI_WINDOW
            war_setWindowMode WINDOW_MODE_windowed

        Case CLI_SHADOWS
            setDrawShadows True

        Case CLI_NOSHADOWS
            setDrawShadows False

        Case CLI_SOUND
            war_setSoundEnabled True

        Case CLI_NOSOUND
            war_setSoundEnabled False

        Case CLI_TEXTURECOMPRESSION
            wz_texture_compression = True

        Case CLI_NOTEXTURECOMPRESSION
            wz_texture_compression = False

        Case CLI_GFXBACKEND
            ' retrieve the backend
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Unrecognised backend"
            End If
            If video_backend_from_str(token, gfxBackend) Then
                war_setGfxBackend gfxBackend
            Else
                qFatal "Unsupported / invalid gfxbackend value"
            End If

        Case CLI_GFXDEBUG
            uses_gfx_debug = True

        Case CLI_JSBACKEND
            ' retrieve the backend
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Unrecognised backend"
            End If
            If js_backend_from_str(token, jsBackend) Then
                war_setJSBackend jsBackend
            Else
                qFatal "Unsupported / invalid jsbackend value"
            End If

        Case CLI_AUTOGAME
            wz_autogame = True
            ' need to cause wrappers.cpp to update calculated effective headless mode
            setHeadlessGameMode wz_cli_headless

        Case CLI_SAVEANDQUIT
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad savegame name (needs to be a full path)"
            ElseIf InStr(token, "/") = 0 Then
                qFatal "Bad savegame name (needs to be a full path)"
            End If
            wz_saveandquit = token

        Case CLI_SKIRMISH
            setHostLaunch HostLaunch_Skirmish
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad test key"
            End If
            wz_test = token

        Case CLI_AUTOHOST
            setHostLaunch HostLaunch_Autohost
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad test key"
            End If
            wz_test = token

        Case CLI_AUTOHEADLESS
            wz_cli_headless = True
            setHeadlessGameMode True

        Case CLI_GAMEPORT
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad game server port"
            End If
            NETsetGameserverPort atoi(token)
            netGameserverPortOverride = True
            debug_ 1108, LOG_INFO, "ParseCommandLine", "Games will be hosted on port [%d]", NETgetGameserverPort()

        Case CLI_NET_PORTMAPPING
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "No value for portmapping option"
            End If
            If cliOptValueToBool(token, optValue) Then
                NetPlay_isPortMappingEnabled = optValue
            Else
                qFatal "Invalid value for portmapping option"
            End If

        Case CLI_STREAMER_SPECTATOR
            wz_streamer_spectator_mode = True

        Case CLI_LOBBY_SLASHCOMMANDS
            wz_lobby_slashcommands = True

        Case CLI_ADD_LOBBY_ADMINHASH
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad admin hash"
            ElseIf Len(token) = 0 Then
                qFatal "Bad admin hash"
            End If
            addLobbyAdminIdentityHash token

        Case CLI_ADD_LOBBY_ADMINPUBLICKEY
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad admin public key"
            ElseIf Len(token) = 0 Then
                qFatal "Bad admin public key"
            End If
            addLobbyAdminPublicKey token

        Case CLI_COMMAND_INTERFACE
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                ' use default, which is currently "stdin"
                token = "stdin"
            ElseIf Len(token) = 0 Then
                token = "stdin"
            End If
            mode = WZ_Command_Interface_None
            value = ""
            If token = "stdin" Then
                mode = WZ_Command_Interface_StdIn_Interface
            ElseIf Left$(token, Len("unixsocket")) = "unixsocket" Then
                mode = WZ_Command_Interface_Unix_Socket
                ' expected form is "unixsocket:path" - parse for the path
                If Len(token) > Len("unixsocket") Then
                    delimeterIdx = Len("unixsocket")
                    If Mid$(token, delimeterIdx + 1, 1) = ":" And Len(token) > delimeterIdx + 1 Then
                        ' grab the rest of the string as the path value
                        value = Mid$(token, delimeterIdx + 2)
                    Else
                        qFatal "Invalid enablecmdinterface unixsocket value (expecting unixsocket:path)"
                    End If
                End If
            Else
                qFatal "Unsupported / invalid enablecmdinterface value"
            End If
            configSetCmdInterface mode, value

        Case CLI_STARTPLAYERS
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad start players count"
            End If
            wz_min_autostart_players = atoi(token)
            If wz_min_autostart_players < 0 Then
                qFatal "Invalid start players count"
            End If
            debug_ 1206, LOG_INFO, "ParseCommandLine", "Games will automatically start with [%d] players (when ready)", wz_min_autostart_players

        Case CLI_GAMELOG_OUTPUTMODES
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                ' use default, which is currently "log,cmdinterface"
                token = "log,cmdinterface"
            ElseIf Len(token) = 0 Then
                token = "log,cmdinterface"
            End If
            modes.cmdInterface = False
            modes.logFile = False
            params = Split(token, ",")                      ' WzString::split
            For i = 0 To UBound(params)
                If params(i) = "log" Then
                    modes.logFile = True
                ElseIf params(i) = "cmdinterface" Then
                    modes.cmdInterface = True
                Else
                    qFatal "Unsupported / invalid gamelog-output value"
                End If
            Next
            GameStoryLogger_setOutputModes modes

        Case CLI_GAMELOG_OUTPUTKEY
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Missing gamelog-outputkey value"
            ElseIf Len(token) = 0 Then
                qFatal "Missing gamelog-outputkey value"
            End If
            If token = "playerindex" Then
                GameStoryLogger_setOutputKey GameStoryLogger_OutputKey_PlayerIndex
            ElseIf token = "playerposition" Then
                GameStoryLogger_setOutputKey GameStoryLogger_OutputKey_PlayerPosition
            Else
                qFatal "Unsupported / invalid gamelog-outputkey value"
            End If

        Case CLI_GAMELOG_OUTPUTNAMING
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Missing gamelog-outputnaming value"
            ElseIf Len(token) = 0 Then
                qFatal "Missing gamelog-outputnaming value"
            End If
            If token = "default" Then
                GameStoryLogger_setOutputNaming GameStoryLogger_OutputNaming_Default
            ElseIf token = "autohosterclassic" Then
                GameStoryLogger_setOutputNaming GameStoryLogger_OutputNaming_AutohosterClassic
            Else
                qFatal "Unsupported / invalid gamelog-outputnaming value"
            End If

        Case CLI_GAMELOG_FRAMEINTERVAL
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad gamelog-frameinterval count"
            End If
            token_intval = atoi(token)
            If token_intval < 0 Then
                qFatal "Invalid gamelog-frameinterval count"
            End If
            GameStoryLogger_setFrameLoggingInterval token_intval

        Case CLI_GAMETIMELIMITMINUTES
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad game time limit minutes count"
            End If
            token_intval = atoi(token)
            If token_intval < 0 Then
                qFatal "Invalid game time limit minutes count"
            End If
            war_setMPGameTimeLimitMinutes token_intval
            game.gameTimeLimitMinutes = war_getMPGameTimeLimitMinutes()

        Case CLI_DEBUG_VERBOSE_SYNCLOG_OUTPUT
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Bad debug verbose synclog output gametime limit"
            End If
            NET_setDebuggingModeVerboseOutputAllSyncLogs atoi(token)

        Case CLI_ALLOW_VULKAN_IMPLICIT_LAYERS
            war_runtimeOnlySetAllowVulkanImplicitLayers True

        Case CLI_HOST_CHAT_CONFIG
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Missing host-chat-config value"
            ElseIf Len(token) = 0 Then
                qFatal "Missing host-chat-config value"
            End If
            If token = "allow" Then
                NETsetDefaultMPHostFreeChatPreference True
            ElseIf token = "quickchat" Or token = "qc" Then
                NETsetDefaultMPHostFreeChatPreference False
            Else
                qFatal "Unsupported / invalid host-chat-config value"
            End If

        Case CLI_HOST_ASYNC_JOIN_APPROVAL
            NETsetAsyncJoinApprovalRequired True

        Case CLI_AUTOHOST_START_NOT_READY
            setHostLaunchStartNotReady True

        Case CLI_HOST_CONNECTION_PROVIDER
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "Missing value for the host connection provider argument"
            ElseIf Len(token) = 0 Then
                qFatal "Missing value for the host connection provider argument"
            End If
            If Not net_backend_from_str(token, pt) Then
                qFatal "Unsupported / invalid network backend"
            End If
            war_setHostConnectionProvider pt

        Case CLI_VB_FRAMES                              ' (VB) scripted test runs
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "(VB) Usage: --frames=<count>"
            End If
            wzSetQuitAfterFrames atoi(token)
            wzSetDialogsSuppressed True

        Case CLI_VB_NODIALOGS                           ' (VB)
            wzSetDialogsSuppressed True

        Case CLI_VB_TEST                                ' (VB) tools\run_tests.py
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "(VB) Usage: --vbtest=<file>"
            End If
            vbTestCommandFile = token
            wzSetDialogsSuppressed True

        Case CLI_VB_DRAWTEST                            ' (VB) tools\draw_test.py
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "(VB) Usage: --vbdrawtest=<name>"
            End If
            vbDrawTest = token
            wzSetDialogsSuppressed True

        Case CLI_VB_SCRIPT                              ' (VB) tools\ui_test.py
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "(VB) Usage: --vbscript=<file>"
            End If
            vbScriptFile = token
            wzSetDialogsSuppressed True

        Case CLI_VB_SNDCAPTURE                          ' (VB) the sound's tests: WzSound captures its mix
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "(VB) Usage: --vbsndcapture=<file.wav>"
            End If
            sound_vbSetCapture token             ' before systemInitialise's audio_Init

        Case CLI_VB_WATCHSHOT                           ' (VB) tools\watch_test.py
            token = poptGetOptArg(poptCon)
            If StrPtr(token) = 0 Then
                qFatal "(VB) Usage: --vbwatchshot=<gameTime ms>,<name>"
            End If
            If InStr(token, ",") = 0 Then
                qFatal "(VB) Usage: --vbwatchshot=<gameTime ms>,<name>"
            End If
            fields = Split(token, ",")
            vbWatchShotTime = atoi(fields(0))
            vbWatchShotName = fields(1)
            vbWatchCameraPlayer = -1
            For i = 2 To UBound(fields)
                If fields(i) = "rt" Then
                    vbWatchRealtime = True
                ElseIf Left$(fields(i), 1) = "p" Then
                    vbWatchCameraPlayer = atoi(Mid$(fields(i), 2))
                End If
            Next
            wzSetDialogsSuppressed True
        End Select ' switch (option)

        iOption = poptGetNextOpt(poptCon)
    Loop ' while

    ParseCommandLine = True
End Function

Public Function autogame_enabled() As Boolean
    autogame_enabled = wz_autogame
End Function

Public Function saveandquit_enabled() As String
    saveandquit_enabled = wz_saveandquit
End Function

Public Function wz_skirmish_test() As String
    wz_skirmish_test = wz_test
End Function

Public Function streamer_spectator_mode() As Boolean
    streamer_spectator_mode = wz_streamer_spectator_mode
End Function

Public Function lobby_slashcommands_enabled() As Boolean
    lobby_slashcommands_enabled = wz_lobby_slashcommands
End Function

Public Function min_autostart_player_count() As Long
    InitStatics
    min_autostart_player_count = wz_min_autostart_players
End Function

Public Function cli_lobby_game_to_connect_str() As String
    cli_lobby_game_to_connect_str = wz_lobby_game_to_connect_str
End Function
