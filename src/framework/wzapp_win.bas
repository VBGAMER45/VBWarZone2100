Attribute VB_Name = "wzapp_win"
' The window-free part of lib/framework/wzapp.h, which Warzone implements in
' lib/sdl/main_sdl.cpp: wzGetTicks, wzDelay, wzDisplayDialog (VB-only, on
' Win32). The window state these need (headless mode, the main window,
' fullscreen) comes from WzPlatform through IFrameworkPlatform once it has
' called frameSetPlatform; before that, dialogs have no owner window, as in
' C++ before the window exists.
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (c As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (f As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long

Private Const MB_ICONERROR As Long = &H10
Private Const MB_ICONWARNING As Long = &H30
Private Const MB_ICONINFORMATION As Long = &H40

Private platform As IFrameworkPlatform
Private ticksStart As Currency
Private ticksFreq As Currency
Private ticksReady As Boolean
Private dialogsSuppressed As Boolean
Private testTicksSet As Boolean
Private testTicks As Long

' (VB) WzPlatform registers itself here; Nothing at shutdown.
Public Sub frameSetPlatform(ByVal p As IFrameworkPlatform)
    Set platform = p
End Sub

' (VB) Scripted test runs: dialogs are logged instead of shown, as in headless mode.
Public Sub wzSetDialogsSuppressed(ByVal suppressed As Boolean)
    dialogsSuppressed = suppressed
End Sub

' (VB) Scripted test runs: a clock that only moves when the test sets it
' (statsref's wzGetTicks gives the same values); False goes back to the real one.
Public Sub wzSetTestTicks(ByVal enabled As Boolean, ByVal ticks As Long)
    testTicksSet = enabled
    testTicks = ticks
End Sub

' Milliseconds since the first call (SDL: since SDL was initialised), as a uint32 bit pattern.
Public Function wzGetTicks() As Long
    Dim c As Currency, ms As Double
    If testTicksSet Then
        wzGetTicks = testTicks
        Exit Function
    End If
    If Not ticksReady Then
        QueryPerformanceCounter ticksStart
        QueryPerformanceFrequency ticksFreq
        ticksReady = True
    End If
    QueryPerformanceCounter c
    ms = Int(CDbl(c - ticksStart) * 1000# / CDbl(ticksFreq))
    ms = ms - Int(ms / 4294967296#) * 4294967296#
    If ms >= 2147483648# Then ms = ms - 4294967296#
    wzGetTicks = ms
End Function

Public Sub wzDelay(ByVal delay As Long)
    Sleep delay
End Sub

Public Sub wzDisplayDialog(ByVal typ As Long, ByVal title As String, ByVal message As String)
    Static processingDialog As Boolean
    Dim headless As Boolean, owner As Long, flags As Long

    If Not platform Is Nothing Then
        headless = Not platform.hasBackend()
        owner = platform.windowHandle()
    End If
    If headless Or dialogsSuppressed Then
        ' while this could be thread_local, thread_local may not yet be supported on all platforms properly
        ' and wzDisplayDialog should probably be called **only** on the main thread anyway
        If Not processingDialog Then
            ' in headless mode, do not display a messagebox (which might block the main thread)
            ' but just log the details
            processingDialog = True
            debug_ 590, LOG_INFO, "wzDisplayDialog", "Suppressed dialog (headless):\n\tTitle: %s\n\tMessage: %s", title, message
            processingDialog = False
        End If
        Exit Sub
    End If

    Select Case typ
    Case Dialog_Error
        flags = MB_ICONERROR
    Case Dialog_Warning
        flags = MB_ICONWARNING
    Case Dialog_Information
        flags = MB_ICONINFORMATION
    End Select
    MessageBoxW owner, StrPtr(message), StrPtr(title), flags
End Sub

' The window functions of wzapp.h that the framework itself calls (debug.cpp, frame.cpp):
' WzPlatform answers them.
Public Function wzIsFullscreen() As Boolean
    If Not platform Is Nothing Then wzIsFullscreen = platform.wzIsFullscreen()
End Function

Public Function wzChangeWindowMode(ByVal mode As Long, Optional ByVal silent As Boolean = False) As Boolean
    If Not platform Is Nothing Then wzChangeWindowMode = platform.wzChangeWindowMode(mode, silent)
End Function

Public Sub inputInitialise()
    If Not platform Is Nothing Then platform.inputInitialise
End Sub
