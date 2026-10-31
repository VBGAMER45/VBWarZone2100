Attribute VB_Name = "exemain_vb"
' The process entry point (VB-only). Warzone's main() (lib/sdl/main_sdl.cpp)
' is "return realmain(argc, argv)": here the command line is split into argv
' the way the C runtime splits it, and a VB run-time error is logged and ends
' the process with code 1 instead of opening VB's modal error box.
' A non-zero exit ends the process with TerminateProcess, not ExitProcess: ExitProcess runs the DLLs'
' detach code in the reverse order of their loading, so a DLL loaded late (AppXDeploymentClient, which
' the NVIDIA GL driver loads) is torn down before MSVBVM60, whose detach then calls CoFreeUnusedLibraries,
' and combase calls the torn-down DLL's DllCanUnloadNow: an access violation (exit code 0xC0000005 after
' the error was logged). A return from Sub Main frees the libraries while all of them are alive.
Option Explicit

Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Sub GetArgs(argv() As String, argc As Long)
    Dim p As Long, i As Long, sp As Long, n As Long
    p = CommandLineToArgvW(GetCommandLineW(), argc)
    ReDim argv(0 To argc)
    For i = 0 To argc - 1
        CopyMemory sp, ByVal p + i * 4, 4
        n = lstrlenW(sp)
        argv(i) = String$(n, 0)
        If n > 0 Then CopyMemory ByVal StrPtr(argv(i)), ByVal sp, n * 2
    Next
    argv(argc) = vbNullString             ' argv[argc] = NULL
    LocalFree p
End Sub

Sub Main()
    Dim argv() As String, argc As Long, exitCode As Long
    Dim errNumber As Long, errDescription As String
    On Error GoTo fail
    stubs_game_init                       ' the C++ initialisers of the stand-ins' globals
    structure_c_init                      ' and of the ported modules' globals
    multiplay_c_init
    difficulty_c_init
    feature_c_init
    research_c_init
    template_c_init
    random_c_init
    map_c_init
    gtime_c_init
    loop_c_init
    init_c_init
    faction_c_init
    paged_entity_container_h_init
    display3d_c_init
    display3d_overlay_c_init
    display_c_init
    terrain_c_init
    texture_c_init
    advvis_c_init
    shadowcascades_c_init
    radar_c_init
    warcam_c_init
    widget_c_init
    hci_c_init
    edit3d_c_init
    groups_c_init
    ingameop_c_init
    intorder_c_init
    input_manager_c_init
    keybind_c_init
    multimenu_c_init
    campaigninfo_c_init
    scores_c_init
    seqdisp_c_init
    loadsave_c_init
    GetArgs argv, argc
    exitCode = realmain(argc, argv)
    debug_exit                            ' (VB) closes the log file; init.cpp's systemShutdown does this in C++
    If exitCode <> 0 Then TerminateProcess GetCurrentProcess(), exitCode
    Exit Sub
fail:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    debug__ 0, LOG_ERROR, "main", "(VB) run-time error %d: %s", errNumber, errDescription
    debug_exit
    TerminateProcess GetCurrentProcess(), 1
End Sub
