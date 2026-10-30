Attribute VB_Name = "main_win"
' lib/sdl/main_sdl.cpp, on Win32 instead of SDL (ported in part). SDL's event
' queue is a ring of WzEvent records: the window procedure only queues (no
' engine code runs inside a Windows callback) and wzEventLoopOneFrame empties
' the ring as SDL_PollEvent would. Keys are Windows scan codes (bit 8 = the
' extended flag) mapped to KEY_CODEs, as initKeycodes maps SDL scancodes.
' Ported: the key and mouse state machines and the input buffer (initKeycodes
' .. inputHandleMouseMotionEvent), wzMain, wzQuit, wzGetQuitExitCode,
' wzShowMouse, wzWindowHasFocus, wzIsFullscreen, wzGetCurrentWindowMode,
' handleGameScreenSizeChange, handleWindowSizeChange,
' wzGetMinimumWindowSizeForDisplayScaleFactor,
' processScreenSizeChangeNotificationIfNeeded, setDisplayScale,
' wzGetCurrentDisplayScale, wzMainScreenSetup_CreateVideoWindow (windowed,
' primary display), wzAttemptInitializeBackend and wzMainScreenSetup (OpenGL,
' no fallback), handleActiveEvent, wzEventLoopOneFrame, wzMainEventLoop,
' wzShutdown, StartTextInput, StopTextInput, isInTextInputMode.
' Missing: fullscreen and display modes, automatic display scaling, cursors
' (cursors_sdl.cpp), the clipboard, the window icon, Vulkan, the backend
' fallback and its prompts, wzAsyncExecOnMainThread, Emscripten.
' Key names (keyScanToString) come from Windows, not SDL, so some differ.
' Also lib/ivis_opengl/gfx_api.h's inline to_swap_mode.
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSG_t
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time_ As Long
    pt As POINTAPI
End Type

Private Type WNDCLASSEXW
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As Long
    hIconSm As Long
End Type

Private Type TRACKMOUSEEVENT_t
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Declare Function RegisterClassExW Lib "user32" (lpwcx As WNDCLASSEXW) As Integer
Private Declare Function UnregisterClassW Lib "user32" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PeekMessageW Lib "user32" (lpMsg As MSG_t, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG_t) As Long
Private Declare Function DispatchMessageW Lib "user32" (lpMsg As MSG_t) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long, ByVal dwExStyle As Long) As Long
Private Declare Function SystemParametersInfoW Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_t) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetKeyNameTextW Lib "user32" (ByVal lParam As Long, ByVal lpString As Long, ByVal nSize As Long) As Long
Private Declare Function MapVirtualKeyW Lib "user32" (ByVal uCode As Long, ByVal uMapType As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function LoadCursorW Lib "user32" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const WM_MOVE As Long = &H3
Private Const WM_SIZE As Long = &H5
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_CLOSE As Long = &H10
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_SHOWWINDOW As Long = &H18
Private Const WM_GETMINMAXINFO As Long = &H24
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_XBUTTONDOWN As Long = &H20B
Private Const WM_XBUTTONUP As Long = &H20C
Private Const WM_MOUSEHWHEEL As Long = &H20E
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const SC_KEYMENU As Long = &HF100&
Private Const SIZE_RESTORED As Long = 0
Private Const SIZE_MINIMIZED As Long = 1
Private Const SIZE_MAXIMIZED As Long = 2
Private Const PM_REMOVE As Long = 1
Private Const CS_OWNDC As Long = &H20
Private Const WS_OVERLAPPEDWINDOW As Long = &HCF0000
Private Const SW_SHOW As Long = 5
Private Const SPI_GETWORKAREA As Long = &H30
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const TME_LEAVE As Long = 2
Private Const IDC_ARROW As Long = 32512
Private Const MAPVK_VK_TO_CHAR As Long = 2
Private Const VK_LEFT As Long = &H25
Private Const VK_UP As Long = &H26
Private Const VK_RIGHT As Long = &H27
Private Const VK_DOWN As Long = &H28
Private Const VK_HOME As Long = &H24
Private Const VK_END As Long = &H23
Private Const VK_INSERT As Long = &H2D
Private Const VK_DELETE As Long = &H2E
Private Const VK_PRIOR As Long = &H21
Private Const VK_NEXT As Long = &H22
Private Const VK_BACK As Long = &H8
Private Const VK_TAB As Long = &H9
Private Const VK_RETURN As Long = &HD
Private Const VK_ESCAPE As Long = &H1B

' The possible states for keys
Private Const KEY_UP As Long = 0
Private Const KEY_PRESSED As Long = 1
Private Const KEY_DOWN As Long = 2
Private Const KEY_RELEASED As Long = 3
Private Const KEY_PRESSRELEASE As Long = 4    ' When a key goes up and down in a frame
Private Const KEY_DOUBLECLICK As Long = 5     ' Only used by mouse keys
Private Const KEY_DRAG As Long = 6            ' Only used by mouse keys

Private Type INPUT_STATE
    state As Long           ' Last key/mouse state
    lastdown As Long        ' last key/mouse button down timestamp
    pressPosX As Long       ' Location of last mouse press event.
    pressPosY As Long
    releasePosX As Long     ' Location of last mouse release event.
    releasePosY As Long
End Type

' constant for the interval between 2 singleclicks for doubleclick event in ms
Private Const DOUBLE_CLICK_INTERVAL As Long = 250
' How far the mouse has to move to start a drag
Private Const DRAG_THRESHOLD As Long = 5
' The size of the input buffer
Private Const INPUT_MAXSTR As Long = 256
Private Const MIN_WZ_GAMESCREEN_WIDTH As Long = 640
Private Const MIN_WZ_GAMESCREEN_HEIGHT As Long = 480
Private Const PACKAGE_NAME As String = "Warzone 2100"

' SDL's events, as the window procedure queues them
Private Const EV_QUIT As Long = 1
Private Const EV_KEY_DOWN As Long = 2
Private Const EV_KEY_UP As Long = 3
Private Const EV_TEXT_INPUT As Long = 4
Private Const EV_MOUSE_BUTTON_DOWN As Long = 5
Private Const EV_MOUSE_BUTTON_UP As Long = 6
Private Const EV_MOUSE_MOTION As Long = 7
Private Const EV_MOUSE_WHEEL As Long = 8
Private Const EV_WINDOW_SHOWN As Long = 20
Private Const EV_WINDOW_HIDDEN As Long = 21
Private Const EV_WINDOW_MOVED As Long = 22
Private Const EV_WINDOW_RESIZED As Long = 23
Private Const EV_WINDOW_MINIMIZED As Long = 24
Private Const EV_WINDOW_MAXIMIZED As Long = 25
Private Const EV_WINDOW_RESTORED As Long = 26
Private Const EV_WINDOW_MOUSE_ENTER As Long = 27
Private Const EV_WINDOW_MOUSE_LEAVE As Long = 28
Private Const EV_WINDOW_FOCUS_GAINED As Long = 29
Private Const EV_WINDOW_FOCUS_LOST As Long = 30
Private Const EV_WINDOW_CLOSE_REQUESTED As Long = 31

Private Type WzEvent
    typ As Long
    a As Long       ' key: scan code; text: UTF-32 character; button; wheel x; window: data1
    b As Long       ' key: virtual key; mouse: x; wheel y; window: data2
    c As Long       ' mouse: y
End Type

Private Const EVQ_SIZE As Long = 1024

' At this time, we only have 1 window.
Public WZwindow As Long
Public WZbackendHasValue As Boolean       ' C: optional<video_backend> WZbackend
Public WZbackend As Long

' The screen that the game window is on.
Public screenIndex As Long
' The logical resolution of the game in the game's coordinate system (points).
Public screenWidth As Long
Public screenHeight As Long
' The logical resolution of the window in the window's coordinate system (points) -
' i.e. not accounting for the Game Display Scale setting.
Public windowWidth As Long
Public windowHeight As Long
' The current display scale factor.
Private current_displayScale As Long
Private current_displayScaleFactor As Single

Private game As IPlatformGame
Private frameworkPlatform As PlatformFramework

' The current state of the keyboard
Private aKeyState(0 To KEY_MAXSCAN - 1) As INPUT_STATE        ' the logical key state (impacted by attempts to clear input)
Private actualKeyState(0 To KEY_MAXSCAN - 1) As INPUT_STATE   ' the underlying key state (ignoring any attempts to clear input)
Private KEY_CODE_to_scancode(0 To KEY_MAXSCAN - 1) As Long
Private scancode_to_KEY_CODE(0 To 511) As Long

' The current location of the mouse (C: Uint16)
Private mouseXPos As Long
Private mouseYPos As Long
Private mouseWheelSpeedX As Long
Private mouseWheelSpeedY As Long
Private mouseInWindow As Boolean
Private windowHasFocus As Boolean

' Which button is being used for a drag
Private dragKey As Long
' The start of a possible drag by the mouse
Private dragX As Long
Private dragY As Long

' The current mouse button state
Private aMouseState(0 To MOUSE_END - 1) As INPUT_STATE
' MousePresses mousePresses
Private mpAction() As Long
Private mpKey() As Long
Private mpX() As Long
Private mpY() As Long
Private mpCount As Long

' The current screen resizing state for this iteration through the game loop, in the game coordinate system
Private resizePending As Boolean     ' C: currentScreenResizingStatus != nullptr
Private resizeOldWidth As Long
Private resizeOldHeight As Long
Private resizeNewWidth As Long
Private resizeNewHeight As Long

' The input string buffer
Private pInputBufferKey(0 To INPUT_MAXSTR - 1) As Long
Private pInputBufferUnicode(0 To INPUT_MAXSTR - 1) As Long
Private pStartBuffer As Long
Private pEndBuffer As Long
Private GetTextEventsOwner As Long

Private wzQuitExitCodeHasValue As Boolean
Private wzQuitExitCode As Long

' (VB) the Win32 side: window class, event ring, window state seen by the window procedure
Private className As String
Private classRegistered As Boolean
Private evq(0 To EVQ_SIZE - 1) As WzEvent
Private evHead As Long
Private evTail As Long
Private queueEnabled As Boolean
Private trackingMouse As Boolean
Private buttonsDown As Long
Private lastClientWidth As Long
Private lastClientHeight As Long
Private lastSizeType As Long
Private minTrackWidth As Long
Private minTrackHeight As Long
Private highSurrogate As Long
Private quitAfterFrames As Long
Private framesDone As Long

'--------------------------------------------------------------------------
' (VB) the event ring

Private Sub QueueEvent(ByVal typ As Long, Optional ByVal a As Long = 0, Optional ByVal b As Long = 0, Optional ByVal c As Long = 0)
    Dim nxt As Long
    If Not queueEnabled Then Exit Sub
    nxt = (evTail + 1) Mod EVQ_SIZE
    If nxt = evHead Then Exit Sub        ' full: drop, as SDL does
    evq(evTail).typ = typ
    evq(evTail).a = a
    evq(evTail).b = b
    evq(evTail).c = c
    evTail = nxt
End Sub

Private Function PollEvent(ev As WzEvent) As Boolean
    If evHead = evTail Then Exit Function
    ev = evq(evHead)
    evHead = (evHead + 1) Mod EVQ_SIZE
    PollEvent = True
End Function

Private Function LoWordSigned(ByVal l As Long) As Long
    LoWordSigned = l And &HFFFF&
    If LoWordSigned >= 32768 Then LoWordSigned = LoWordSigned - 65536
End Function

Private Function HiWordSigned(ByVal l As Long) As Long
    HiWordSigned = (l And &HFFFF0000) \ &H10000
End Function

' The window procedure: queues events, runs no engine code (a VB error cannot unwind through it).
Public Function WzWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim scan As Long, tme As TRACKMOUSEEVENT_t, rc As RECT, w As Long, h As Long, ch As Long
    On Error Resume Next
    Select Case uMsg
    Case WM_CLOSE
        QueueEvent EV_WINDOW_CLOSE_REQUESTED
        QueueEvent EV_QUIT          ' SDL quits when its last window is closed
        Exit Function
    Case WM_SYSCOMMAND
        If (wParam And &HFFF0&) = SC_KEYMENU Then Exit Function    ' Alt alone must not open the system menu
    Case WM_KEYDOWN, WM_SYSKEYDOWN, WM_KEYUP, WM_SYSKEYUP
        scan = (lParam And &HFF0000) \ &H10000
        If (lParam And &H1000000) <> 0 Then scan = scan Or &H100
        If uMsg = WM_KEYDOWN Or uMsg = WM_SYSKEYDOWN Then
            QueueEvent EV_KEY_DOWN, scan, wParam
        Else
            QueueEvent EV_KEY_UP, scan, wParam
        End If
        If uMsg = WM_KEYDOWN Or uMsg = WM_KEYUP Then Exit Function
        ' WM_SYSKEY*: on to DefWindowProc, for Alt+F4
    Case WM_CHAR
        ch = wParam And &HFFFF&
        If ch >= &HD800& And ch <= &HDBFF& Then
            highSurrogate = ch
        ElseIf ch >= &HDC00& And ch <= &HDFFF& Then
            If highSurrogate <> 0 Then QueueEvent EV_TEXT_INPUT, &H10000 + (highSurrogate - &HD800&) * &H400& + (ch - &HDC00&)
            highSurrogate = 0
        ElseIf ch >= 32 And ch <> 127 Then
            QueueEvent EV_TEXT_INPUT, ch     ' SDL sends no text for control characters
        End If
        Exit Function
    Case WM_MOUSEMOVE
        If Not trackingMouse Then
            tme.cbSize = LenB(tme)
            tme.dwFlags = TME_LEAVE
            tme.hwndTrack = hwnd
            TrackMouseEvent tme
            trackingMouse = True
            QueueEvent EV_WINDOW_MOUSE_ENTER
        End If
        QueueEvent EV_MOUSE_MOTION, 0, LoWordSigned(lParam), HiWordSigned(lParam)
        Exit Function
    Case WM_MOUSELEAVE
        trackingMouse = False
        QueueEvent EV_WINDOW_MOUSE_LEAVE
        Exit Function
    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_XBUTTONDOWN
        Select Case uMsg
        Case WM_LBUTTONDOWN: QueueEvent EV_MOUSE_BUTTON_DOWN, MOUSE_LMB, LoWordSigned(lParam), HiWordSigned(lParam)
        Case WM_RBUTTONDOWN: QueueEvent EV_MOUSE_BUTTON_DOWN, MOUSE_RMB, LoWordSigned(lParam), HiWordSigned(lParam)
        Case WM_MBUTTONDOWN: QueueEvent EV_MOUSE_BUTTON_DOWN, MOUSE_MMB, LoWordSigned(lParam), HiWordSigned(lParam)
        Case Else
            If HiWordSigned(wParam) = 1 Then
                QueueEvent EV_MOUSE_BUTTON_DOWN, MOUSE_X1, LoWordSigned(lParam), HiWordSigned(lParam)
            Else
                QueueEvent EV_MOUSE_BUTTON_DOWN, MOUSE_X2, LoWordSigned(lParam), HiWordSigned(lParam)
            End If
        End Select
        If buttonsDown = 0 Then SetCapture hwnd     ' SDL keeps the mouse while a button is down
        buttonsDown = buttonsDown + 1
        If uMsg = WM_XBUTTONDOWN Then WzWndProc = 1
        Exit Function
    Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP, WM_XBUTTONUP
        Select Case uMsg
        Case WM_LBUTTONUP: QueueEvent EV_MOUSE_BUTTON_UP, MOUSE_LMB, LoWordSigned(lParam), HiWordSigned(lParam)
        Case WM_RBUTTONUP: QueueEvent EV_MOUSE_BUTTON_UP, MOUSE_RMB, LoWordSigned(lParam), HiWordSigned(lParam)
        Case WM_MBUTTONUP: QueueEvent EV_MOUSE_BUTTON_UP, MOUSE_MMB, LoWordSigned(lParam), HiWordSigned(lParam)
        Case Else
            If HiWordSigned(wParam) = 1 Then
                QueueEvent EV_MOUSE_BUTTON_UP, MOUSE_X1, LoWordSigned(lParam), HiWordSigned(lParam)
            Else
                QueueEvent EV_MOUSE_BUTTON_UP, MOUSE_X2, LoWordSigned(lParam), HiWordSigned(lParam)
            End If
        End Select
        If buttonsDown > 0 Then buttonsDown = buttonsDown - 1
        If buttonsDown = 0 Then ReleaseCapture
        If uMsg = WM_XBUTTONUP Then WzWndProc = 1
        Exit Function
    Case WM_MOUSEWHEEL
        QueueEvent EV_MOUSE_WHEEL, 0, HiWordSigned(wParam) \ 120
        Exit Function
    Case WM_MOUSEHWHEEL
        QueueEvent EV_MOUSE_WHEEL, HiWordSigned(wParam) \ 120, 0
        Exit Function
    Case WM_SIZE
        w = LoWordSigned(lParam)
        h = HiWordSigned(lParam)
        Select Case wParam
        Case SIZE_MINIMIZED
            QueueEvent EV_WINDOW_MINIMIZED
        Case SIZE_MAXIMIZED
            QueueEvent EV_WINDOW_MAXIMIZED
        Case SIZE_RESTORED
            If lastSizeType <> SIZE_RESTORED Then QueueEvent EV_WINDOW_RESTORED
        End Select
        If wParam <> SIZE_MINIMIZED And (w <> lastClientWidth Or h <> lastClientHeight) Then
            lastClientWidth = w
            lastClientHeight = h
            QueueEvent EV_WINDOW_RESIZED, w, h
        End If
        lastSizeType = wParam
    Case WM_MOVE
        QueueEvent EV_WINDOW_MOVED, LoWordSigned(lParam), HiWordSigned(lParam)
    Case WM_SHOWWINDOW
        If wParam <> 0 Then QueueEvent EV_WINDOW_SHOWN Else QueueEvent EV_WINDOW_HIDDEN
    Case WM_SETFOCUS
        QueueEvent EV_WINDOW_FOCUS_GAINED
    Case WM_KILLFOCUS
        QueueEvent EV_WINDOW_FOCUS_LOST
    Case WM_GETMINMAXINFO
        If minTrackWidth > 0 Then
            rc.Right = minTrackWidth
            rc.Bottom = minTrackHeight
            AdjustWindowRectEx rc, WS_OVERLAPPEDWINDOW, 0, 0
            w = rc.Right - rc.Left
            h = rc.Bottom - rc.Top
            CopyMemory ByVal lParam + 24, w, 4         ' MINMAXINFO.ptMinTrackSize
            CopyMemory ByVal lParam + 28, h, 4
        End If
        Exit Function
    Case WM_ERASEBKGND
        WzWndProc = 1
        Exit Function
    End Select
    WzWndProc = DefWindowProcW(hwnd, uMsg, wParam, lParam)
End Function

Private Function FnPtr(ByVal p As Long) As Long
    FnPtr = p
End Function

'--------------------------------------------------------------------------
' Misc support

Public Function wzGetPlatform() As String
    wzGetPlatform = "Windows"
End Function

' (VB) The game's side of lib/sdl; Nothing at shutdown.
Public Sub PlatformSetGame(ByVal g As IPlatformGame)
    Set game = g
End Sub

' (VB) Scripted test runs: wzQuit(0) after this many frames (0 = never).
Public Sub wzSetQuitAfterFrames(ByVal frames As Long)
    quitAfterFrames = frames
End Sub

Public Function wzGetFramesDone() As Long
    wzGetFramesDone = framesDone
End Function

Private Sub setDisplayScale(ByVal displayScale As Long)
    current_displayScale = displayScale
    current_displayScaleFactor = CSng(displayScale) / 100!
End Sub

Public Function wzGetCurrentDisplayScale() As Long
    wzGetCurrentDisplayScale = current_displayScale
End Function

' The window's drawable size in pixels over its logical size (SDL_GL_GetDrawableSize over
' SDL_GetWindowSize): on Windows both are the client area in pixels
Public Sub wzGetWindowToRendererScaleFactor(horizScaleFactor As Single, vertScaleFactor As Single)
    Dim rc As RECT, drawableWidth As Long, drawableHeight As Long, logicalWindowWidth As Long, logicalWindowHeight As Long
    If Not WZbackendHasValue Then
        horizScaleFactor = 1!
        vertScaleFactor = 1!
        Exit Sub
    End If
    If Not (WZwindow <> 0) Then
        ASSERT_FAILURE "WZwindow != nullptr", "lib/sdl/main_sdl.cpp:3427", "wzGetWindowToRendererScaleFactor", "wzGetWindowToRendererScaleFactor called when window is not available"
        Exit Sub
    End If

    ' Obtain the window context's drawable size in pixels
    GetClientRect WZwindow, rc
    drawableWidth = rc.Right - rc.Left
    drawableHeight = rc.Bottom - rc.Top

    ' Obtain the logical window size (in points)
    logicalWindowWidth = drawableWidth
    logicalWindowHeight = drawableHeight

    debug_ 3437, LOG_WZ, "wzGetWindowToRendererScaleFactor", "Window Logical Size (%d, %d) vs Drawable Size in Pixels (%d, %d)", logicalWindowWidth, logicalWindowHeight, drawableWidth, drawableHeight

    horizScaleFactor = CSng(drawableWidth) / CSng(logicalWindowWidth) ' Do **NOT** multiply by current_displayScaleFactor
    vertScaleFactor = CSng(drawableHeight) / CSng(logicalWindowHeight) ' Do **NOT** multiply by current_displayScaleFactor
End Sub

Public Sub wzGetGameToRendererScaleFactorInt(horizScalePercentage As Long, vertScalePercentage As Long)
    Dim horizWindowScaleFactor As Single, vertWindowScaleFactor As Single
    wzGetWindowToRendererScaleFactor horizWindowScaleFactor, vertWindowScaleFactor

    horizScalePercentage = -Int(-(horizWindowScaleFactor * current_displayScale))  ' ceil
    vertScalePercentage = -Int(-(vertWindowScaleFactor * current_displayScale))
End Sub

Public Sub wzShowMouse(ByVal visible As Boolean)
    If visible Then
        ShowCursor 1
    Else
        ShowCursor 0
    End If
End Sub

Public Function wzGetCurrentWindowMode() As WINDOW_MODE
    ' (VB) windowed only so far
    wzGetCurrentWindowMode = WINDOW_MODE_windowed
End Function

Public Function wzChangeWindowMode(ByVal mode As WINDOW_MODE, Optional ByVal silent As Boolean = False) As Boolean
    ' (VB) windowed only so far
    wzChangeWindowMode = (mode = WINDOW_MODE_windowed)
End Function

Public Function wzIsFullscreen() As Boolean
    wzIsFullscreen = False
End Function

Public Function wzWindowHasFocus() As Boolean
    wzWindowHasFocus = windowHasFocus
End Function

Public Sub wzQuit(ByVal exitCode As Long)
    If Not wzQuitExitCodeHasValue Then
        wzQuitExitCode = exitCode
        wzQuitExitCodeHasValue = True
    End If
    ' Create a quit event to halt game loop.
    QueueEvent EV_QUIT
End Sub

Public Function wzGetQuitExitCode() As Long
    If wzQuitExitCodeHasValue Then wzGetQuitExitCode = wzQuitExitCode
End Function

'--------------------------------------------------------------------------
' Keys

Private Sub AddKeycode(ByVal code As Long, ByVal scancode As Long)
    KEY_CODE_to_scancode(code) = scancode
End Sub

' The keycodes we care about (Windows scan codes; &H100 = extended)
Private Sub initKeycodes()
    Dim code As Long, scan As Long
    AddKeycode KEY_ESC, &H1
    AddKeycode KEY_1, &H2
    AddKeycode KEY_2, &H3
    AddKeycode KEY_3, &H4
    AddKeycode KEY_4, &H5
    AddKeycode KEY_5, &H6
    AddKeycode KEY_6, &H7
    AddKeycode KEY_7, &H8
    AddKeycode KEY_8, &H9
    AddKeycode KEY_9, &HA
    AddKeycode KEY_0, &HB
    AddKeycode KEY_MINUS, &HC
    AddKeycode KEY_EQUALS, &HD
    AddKeycode KEY_BACKSPACE, &HE
    AddKeycode KEY_TAB, &HF
    AddKeycode KEY_Q, &H10
    AddKeycode KEY_W, &H11
    AddKeycode KEY_E, &H12
    AddKeycode KEY_R, &H13
    AddKeycode KEY_T, &H14
    AddKeycode KEY_Y, &H15
    AddKeycode KEY_U, &H16
    AddKeycode KEY_I, &H17
    AddKeycode KEY_O, &H18
    AddKeycode KEY_P, &H19
    AddKeycode KEY_LBRACE, &H1A
    AddKeycode KEY_RBRACE, &H1B
    AddKeycode KEY_RETURN, &H1C
    AddKeycode KEY_LCTRL, &H1D
    AddKeycode KEY_A, &H1E
    AddKeycode KEY_S, &H1F
    AddKeycode KEY_D, &H20
    AddKeycode KEY_F, &H21
    AddKeycode KEY_G, &H22
    AddKeycode KEY_H, &H23
    AddKeycode KEY_J, &H24
    AddKeycode KEY_K, &H25
    AddKeycode KEY_L, &H26
    AddKeycode KEY_SEMICOLON, &H27
    AddKeycode KEY_QUOTE, &H28
    AddKeycode KEY_BACKQUOTE, &H29
    AddKeycode KEY_LSHIFT, &H2A
    AddKeycode KEY_LMETA, &H15B
    AddKeycode KEY_LSUPER, &H15B
    AddKeycode KEY_BACKSLASH, &H2B
    AddKeycode KEY_Z, &H2C
    AddKeycode KEY_X, &H2D
    AddKeycode KEY_C, &H2E
    AddKeycode KEY_V, &H2F
    AddKeycode KEY_B, &H30
    AddKeycode KEY_N, &H31
    AddKeycode KEY_M, &H32
    AddKeycode KEY_COMMA, &H33
    AddKeycode KEY_FULLSTOP, &H34
    AddKeycode KEY_FORWARDSLASH, &H35
    AddKeycode KEY_RSHIFT, &H36
    AddKeycode KEY_RMETA, &H15C
    AddKeycode KEY_RSUPER, &H15C
    AddKeycode KEY_KP_STAR, &H37
    AddKeycode KEY_LALT, &H38
    AddKeycode KEY_SPACE, &H39
    AddKeycode KEY_CAPSLOCK, &H3A
    AddKeycode KEY_F1, &H3B
    AddKeycode KEY_F2, &H3C
    AddKeycode KEY_F3, &H3D
    AddKeycode KEY_F4, &H3E
    AddKeycode KEY_F5, &H3F
    AddKeycode KEY_F6, &H40
    AddKeycode KEY_F7, &H41
    AddKeycode KEY_F8, &H42
    AddKeycode KEY_F9, &H43
    AddKeycode KEY_F10, &H44
    AddKeycode KEY_NUMLOCK, &H145
    AddKeycode KEY_SCROLLLOCK, &H46
    AddKeycode KEY_KP_7, &H47
    AddKeycode KEY_KP_8, &H48
    AddKeycode KEY_KP_9, &H49
    AddKeycode KEY_KP_MINUS, &H4A
    AddKeycode KEY_KP_4, &H4B
    AddKeycode KEY_KP_5, &H4C
    AddKeycode KEY_KP_6, &H4D
    AddKeycode KEY_KP_PLUS, &H4E
    AddKeycode KEY_KP_1, &H4F
    AddKeycode KEY_KP_2, &H50
    AddKeycode KEY_KP_3, &H51
    AddKeycode KEY_KP_0, &H52
    AddKeycode KEY_KP_FULLSTOP, &H53
    AddKeycode KEY_F11, &H57
    AddKeycode KEY_F12, &H58
    AddKeycode KEY_RCTRL, &H11D
    AddKeycode KEY_KP_BACKSLASH, &H135
    AddKeycode KEY_RALT, &H138
    AddKeycode KEY_HOME, &H147
    AddKeycode KEY_UPARROW, &H148
    AddKeycode KEY_PAGEUP, &H149
    AddKeycode KEY_LEFTARROW, &H14B
    AddKeycode KEY_RIGHTARROW, &H14D
    AddKeycode KEY_END, &H14F
    AddKeycode KEY_DOWNARROW, &H150
    AddKeycode KEY_PAGEDOWN, &H151
    AddKeycode KEY_INSERT, &H152
    AddKeycode KEY_DELETE, &H153
    AddKeycode KEY_KPENTER, &H11C

    ' The reverse map, built in KEY_CODE order with the first one kept (std::map::insert),
    ' so the shared GUI keys map back to KEY_LMETA / KEY_RMETA as in C++
    For scan = 0 To 511
        scancode_to_KEY_CODE(scan) = KEY_MAXSCAN
    Next
    For code = 0 To KEY_MAXSCAN - 1
        scan = KEY_CODE_to_scancode(code)
        If scan <> 0 Then
            If scancode_to_KEY_CODE(scan) = KEY_MAXSCAN Then scancode_to_KEY_CODE(scan) = code
        End If
    Next
End Sub

Private Function scancodeToKeyCode(ByVal key As Long) As Long
    If key < 0 Or key > 511 Then
        scancodeToKeyCode = KEY_MAXSCAN
    Else
        scancodeToKeyCode = scancode_to_KEY_CODE(key)
    End If
End Function

Private Function keyCodeToScancode(ByVal code As Long) As Long
    If code >= 0 And code < KEY_MAXSCAN Then
        If KEY_CODE_to_scancode(code) <> 0 Then
            keyCodeToScancode = KEY_CODE_to_scancode(code)
            Exit Function
        End If
    End If
    keyCodeToScancode = code
End Function

' The key's name in the current keyboard layout (SDL_GetScancodeName / SDL_GetKeyName).
Private Function ScancodeName(ByVal scan As Long) As String
    Dim buf As String, n As Long, lp As Long
    lp = (scan And &HFF&) * &H10000
    If (scan And &H100&) <> 0 Then lp = lp Or &H1000000
    buf = String$(64, 0)
    n = GetKeyNameTextW(lp, StrPtr(buf), 64)
    ScancodeName = Left$(buf, n)
End Function

' Cyclic increment.
Private Function inputPointerNext(ByVal p As Long) As Long
    If p + 1 = INPUT_MAXSTR Then inputPointerNext = 0 Else inputPointerNext = p + 1
End Function

' add count copies of the characater code to the input buffer
Private Sub inputAddBuffer(ByVal key As Long, ByVal unicode As Long)
    Dim pNext As Long
    If GetTextEventsOwner = 0 Then
        ' Text input events aren't enabled
        Exit Sub
    End If

    ' Calculate what pEndBuffer will be set to next
    pNext = inputPointerNext(pEndBuffer)

    If pNext = pStartBuffer Then
        Exit Sub ' Buffer full.
    End If

    ' Add key to buffer.
    pInputBufferKey(pEndBuffer) = key
    pInputBufferUnicode(pEndBuffer) = unicode
    pEndBuffer = pNext
End Sub

' Returns the human-readable name for the key *in the current keyboard layout*
Public Function keyScanToString(ByVal code As Long) As String
    Dim ascii As String
    If code = KEY_LCTRL Then
        ' shortcuts with modifier keys work with either key.
        keyScanToString = "Ctrl"
        Exit Function
    ElseIf code = KEY_LSHIFT Then
        keyScanToString = "Shift"
        Exit Function
    ElseIf code = KEY_LALT Then
        keyScanToString = "Alt"
        Exit Function
    ElseIf code = KEY_LMETA Then
        keyScanToString = "Meta"
        Exit Function
    End If

    If code < KEY_MAXSCAN Then
        ascii = ScancodeName(keyCodeToScancode(code))
        If Len(ascii) > 1 Then
            If Left$(ascii, 1) Like "[a-z]" Then
                ' capitalize
                ascii = UCase$(Left$(ascii, 1)) & Mid$(ascii, 2)
            End If
        End If
    Else
        ascii = "???"
    End If
    keyScanToString = ascii
End Function

Public Function mouseKeyCodeToString(ByVal code As Long) As String
    Select Case code
    Case MOUSE_LMB: mouseKeyCodeToString = "Mouse Left"
    Case MOUSE_MMB: mouseKeyCodeToString = "Mouse Middle"
    Case MOUSE_RMB: mouseKeyCodeToString = "Mouse Right"
    Case MOUSE_X1: mouseKeyCodeToString = "Mouse 4"
    Case MOUSE_X2: mouseKeyCodeToString = "Mouse 5"
    Case MOUSE_WUP: mouseKeyCodeToString = "Mouse Wheel Up"
    Case MOUSE_WDN: mouseKeyCodeToString = "Mouse Wheel Down"
    Case Else: mouseKeyCodeToString = "Mouse ???"
    End Select
End Function

' Initialise the input module
Public Sub inputInitialise()
    Dim i As Long
    For i = 0 To KEY_MAXSCAN - 1
        aKeyState(i).state = KEY_UP
        actualKeyState(i).state = KEY_UP
    Next

    For i = 0 To MOUSE_END - 1
        aMouseState(i).state = KEY_UP
    Next

    pStartBuffer = 0
    pEndBuffer = 0

    mouseXPos = screenWidth \ 2
    dragX = mouseXPos
    mouseYPos = screenHeight \ 2
    dragY = mouseYPos
    dragKey = MOUSE_LMB
End Sub

' Clear the input buffer
Public Sub inputClearBuffer()
    pStartBuffer = 0
    pEndBuffer = 0
End Sub

' Return the next key press or 0 if no key in the buffer; unicode gets its character.
' All key presses are buffered up (including windows auto repeat).
Public Function inputGetKey(unicode As Long) As Long
    Dim retVal As Long
    If pStartBuffer = pEndBuffer Then
        Exit Function ' Buffer empty.
    End If

    retVal = pInputBufferKey(pStartBuffer)
    unicode = pInputBufferUnicode(pStartBuffer)
    If retVal = 0 Then
        retVal = 32 ' Don't return 0 if we got a virtual key, since that's interpreted as no input.
    End If
    pStartBuffer = inputPointerNext(pStartBuffer)

    inputGetKey = retVal
End Function

' MousePresses const &inputGetClicks(): the count, then each with inputGetClick.
Public Function inputGetClickCount() As Long
    inputGetClickCount = mpCount
End Function

Public Sub inputGetClick(ByVal i As Long, action As Long, key As Long, x As Long, y As Long)
    action = mpAction(i)
    key = mpKey(i)
    x = mpX(i)
    y = mpY(i)
End Sub

' This is called once a frame so that the system can tell
' whether a key was pressed this turn or held down from the last frame.
Public Sub inputNewFrame()
    Dim i As Long
    ' handle the keyboard
    For i = 0 To KEY_MAXSCAN - 1
        If aKeyState(i).state = KEY_PRESSED Then
            aKeyState(i).state = KEY_DOWN
            actualKeyState(i).state = KEY_DOWN
            If debugPartEnabled(LOG_NEVER) Then debug_ 1379, LOG_NEVER, "inputNewFrame", "This key is DOWN! %x, %d [%s]", i, i, ScancodeName(keyCodeToScancode(i))
        ElseIf aKeyState(i).state = KEY_RELEASED Or aKeyState(i).state = KEY_PRESSRELEASE Then
            aKeyState(i).state = KEY_UP
            actualKeyState(i).state = KEY_UP
            If debugPartEnabled(LOG_NEVER) Then debug_ 1386, LOG_NEVER, "inputNewFrame", "This key is UP! %x, %d [%s]", i, i, ScancodeName(keyCodeToScancode(i))
        End If
    Next

    ' handle the mouse
    For i = 0 To MOUSE_END - 1
        If aMouseState(i).state = KEY_PRESSED Then
            aMouseState(i).state = KEY_DOWN
        ElseIf aMouseState(i).state = KEY_RELEASED Or aMouseState(i).state = KEY_DOUBLECLICK Or aMouseState(i).state = KEY_PRESSRELEASE Then
            aMouseState(i).state = KEY_UP
        End If
    Next
    mpCount = 0
    mouseWheelSpeedX = 0
    mouseWheelSpeedY = 0
End Sub

' Release all keys (and buttons) when we lose focus
Public Sub inputLoseFocus()
    Dim i As Long
    ' Lost the window focus, have to take this as a global key up
    For i = 0 To KEY_MAXSCAN - 1
        aKeyState(i).state = KEY_UP
        ' Do *NOT* clear actualKeyState here!
    Next
    For i = 0 To MOUSE_END - 1
        aMouseState(i).state = KEY_UP
    Next
End Sub

Private Sub restoreKeyDownState(ByVal code As Long)
    If actualKeyState(code).state <> KEY_UP Then
        aKeyState(code) = actualKeyState(code)
    End If
End Sub

Public Sub inputRestoreMetaKeyState()
    restoreKeyDownState KEY_RALT
    restoreKeyDownState KEY_LALT

    restoreKeyDownState KEY_RCTRL
    restoreKeyDownState KEY_LCTRL

    restoreKeyDownState KEY_RSHIFT
    restoreKeyDownState KEY_LSHIFT

    restoreKeyDownState KEY_RMETA
    restoreKeyDownState KEY_LMETA
End Sub

' This returns true if the key is currently depressed
Public Function keyDown(ByVal code As Long) As Boolean
    If Not (code < KEY_MAXSCAN) Then
        ASSERT_FAILURE "code < KEY_MAXSCAN", "lib/sdl/main_sdl.cpp:1451", "keyDown", "Invalid keycode of %d!", code
        Exit Function
    End If
    keyDown = (aKeyState(code).state <> KEY_UP)
End Function

' This returns true if the key went from being up to being down this frame
Public Function keyPressed(ByVal code As Long) As Boolean
    If Not (code < KEY_MAXSCAN) Then
        ASSERT_FAILURE "code < KEY_MAXSCAN", "lib/sdl/main_sdl.cpp:1458", "keyPressed", "Invalid keycode of %d!", code
        Exit Function
    End If
    keyPressed = (aKeyState(code).state = KEY_PRESSED) Or (aKeyState(code).state = KEY_PRESSRELEASE)
End Function

' This returns true if the key went from being down to being up this frame
Public Function keyReleased(ByVal code As Long) As Boolean
    If Not (code < KEY_MAXSCAN) Then
        ASSERT_FAILURE "code < KEY_MAXSCAN", "lib/sdl/main_sdl.cpp:1465", "keyReleased", "Invalid keycode of %d!", code
        Exit Function
    End If
    keyReleased = (aKeyState(code).state = KEY_RELEASED) Or (aKeyState(code).state = KEY_PRESSRELEASE)
End Function

' Return the X coordinate of the mouse
Public Function mouseX() As Long
    mouseX = mouseXPos
End Function

' Return the Y coordinate of the mouse
Public Function mouseY() As Long
    mouseY = mouseYPos
End Function

Public Function wzMouseInWindow() As Boolean
    wzMouseInWindow = mouseInWindow
End Function

' Vector2i const& getMouseWheelSpeed()
Public Sub getMouseWheelSpeed(x As Long, y As Long)
    x = mouseWheelSpeedX
    y = mouseWheelSpeedY
End Sub

Public Sub mousePressPos_DEPRECATED(ByVal code As Long, x As Long, y As Long)
    x = aMouseState(code).pressPosX
    y = aMouseState(code).pressPosY
End Sub

Public Sub mouseReleasePos_DEPRECATED(ByVal code As Long, x As Long, y As Long)
    x = aMouseState(code).releasePosX
    y = aMouseState(code).releasePosY
End Sub

' This returns true if the mouse key is currently depressed
Public Function mouseDown(ByVal code As Long) As Boolean
    If aMouseState(code).state <> KEY_UP Then
        mouseDown = True
    ElseIf code = MOUSE_MMB Then
        ' holding down LMB and RMB counts as holding down MMB
        mouseDown = aMouseState(MOUSE_LMB).state <> KEY_UP And aMouseState(MOUSE_RMB).state <> KEY_UP
    End If
End Function

' This returns true if the mouse key was double clicked
Public Function mouseDClicked(ByVal code As Long) As Boolean
    mouseDClicked = (aMouseState(code).state = KEY_DOUBLECLICK)
End Function

' This returns true if the mouse key went from being up to being down this frame
Public Function mousePressed(ByVal code As Long) As Boolean
    mousePressed = (aMouseState(code).state = KEY_PRESSED) Or _
                   (aMouseState(code).state = KEY_DOUBLECLICK) Or _
                   (aMouseState(code).state = KEY_PRESSRELEASE)
End Function

' This returns true if the mouse key went from being down to being up this frame
Public Function mouseReleased(ByVal code As Long) As Boolean
    mouseReleased = (aMouseState(code).state = KEY_RELEASED) Or _
                    (aMouseState(code).state = KEY_DOUBLECLICK) Or _
                    (aMouseState(code).state = KEY_PRESSRELEASE)
End Function

' Check for a mouse drag, return the drag start coords if dragging
Public Function mouseDrag(ByVal code As Long, px As Long, py As Long) As Boolean
    Dim dragging As Boolean
    If aMouseState(code).state = KEY_DRAG Then
        dragging = True
    ElseIf code = MOUSE_MMB Then
        ' dragging LMB and RMB counts as dragging MMB
        If aMouseState(MOUSE_LMB).state = KEY_DRAG And aMouseState(MOUSE_RMB).state <> KEY_UP Then
            dragging = True
        ElseIf aMouseState(MOUSE_LMB).state <> KEY_UP And aMouseState(MOUSE_RMB).state = KEY_DRAG Then
            dragging = True
        End If
    End If
    If dragging Then
        px = dragX
        py = dragY
        mouseDrag = True
    End If
End Function

' Handle keyboard events. scancode: the Windows scan code; vk: the virtual key (SDL's keycode).
Private Sub inputHandleKeyEvent(ByVal typ As Long, ByVal currentKey As Long, ByVal vkey As Long)
    Dim vk As Long, code As Long, sdlKey As Long
    Select Case typ
    Case EV_KEY_DOWN
        vk = 0
        Select Case vkey
        ' our "editing" keys for text
        Case VK_LEFT: vk = INPBUF_LEFT
        Case VK_RIGHT: vk = INPBUF_RIGHT
        Case VK_UP: vk = INPBUF_UP
        Case VK_DOWN: vk = INPBUF_DOWN
        Case VK_HOME: vk = INPBUF_HOME
        Case VK_END: vk = INPBUF_END
        Case VK_INSERT: vk = INPBUF_INS
        Case VK_DELETE: vk = INPBUF_DEL
        Case VK_PRIOR: vk = INPBUF_PGUP
        Case VK_NEXT: vk = INPBUF_PGDN
        Case VK_BACK: vk = INPBUF_BKSPACE
        Case VK_TAB: vk = INPBUF_TAB
        Case VK_RETURN: vk = INPBUF_CR
        Case VK_ESCAPE: vk = INPBUF_ESC
        End Select

        If vk <> 0 Then
            ' Take care of adding 'editing' keys that were pressed to the input buffer (for text editing control handling)
            inputAddBuffer vk, 0
            debug_ 1611, LOG_INPUT, "inputHandleKeyEvent", "Editing key: 0x%x, %d SDLkey=[%s] pressed", vk, vk, ScancodeName(currentKey)
        Else
            ' add everything else; SDL's keycode is the key's character (lower case), or the scancode with bit 30
            sdlKey = MapVirtualKeyW(vkey, MAPVK_VK_TO_CHAR) And &H7FFFFFFF
            If sdlKey >= 65 And sdlKey <= 90 Then sdlKey = sdlKey + 32
            If sdlKey = 0 Then sdlKey = &H40000000 Or currentKey
            inputAddBuffer sdlKey, 0
        End If

        debug_ 1621, LOG_INPUT, "inputHandleKeyEvent", "Key Code (pressed): 0x%x, %d, SDLscancode=[%s]", currentKey, currentKey, ScancodeName(currentKey)

        code = scancodeToKeyCode(currentKey)
        If code >= KEY_MAXSCAN Then
            Exit Sub
        End If
        If aKeyState(code).state = KEY_UP Or _
           aKeyState(code).state = KEY_RELEASED Or _
           aKeyState(code).state = KEY_PRESSRELEASE Then
            ' whether double key press or not
            aKeyState(code).state = KEY_PRESSED
            aKeyState(code).lastdown = 0
            actualKeyState(code).state = KEY_PRESSED
            actualKeyState(code).lastdown = 0
        End If

    Case EV_KEY_UP
        debug_ 1644, LOG_INPUT, "inputHandleKeyEvent", "Key Code (*Depressed*): 0x%x, %d, SDLscancode=[%s]", currentKey, currentKey, ScancodeName(currentKey)
        code = scancodeToKeyCode(currentKey)
        If code >= KEY_MAXSCAN Then
            Exit Sub
        End If
        actualKeyState(code).state = KEY_UP
        If aKeyState(code).state = KEY_PRESSED Then
            aKeyState(code).state = KEY_PRESSRELEASE
        ElseIf aKeyState(code).state = KEY_DOWN Then
            aKeyState(code).state = KEY_RELEASED
        End If
    End Select
End Sub

' Handle text events: one UTF-32 character
Private Sub inputhandleText(ByVal ch As Long)
    Dim text As String
    If ch >= &H10000 Then
        text = ChrW$(&HD800& + ((ch - &H10000) \ &H400&)) & ChrW$(&HDC00& + ((ch - &H10000) And &H3FF&))
    Else
        text = ChrW$(ch)
    End If
    debug_ 1682, LOG_INPUT, "inputhandleText", "Keyboard: text input ""%s""", text
    inputAddBuffer 0, ch
End Sub

' Handle mouse wheel events
Private Sub inputHandleMouseWheelEvent(ByVal wheelX As Long, ByVal wheelY As Long)
    mouseWheelSpeedX = mouseWheelSpeedX + wheelX
    mouseWheelSpeedY = mouseWheelSpeedY + wheelY

    If wheelX > 0 Or wheelY > 0 Then
        aMouseState(MOUSE_WUP).state = KEY_PRESSED
        aMouseState(MOUSE_WUP).lastdown = 0
    ElseIf wheelX < 0 Or wheelY < 0 Then
        aMouseState(MOUSE_WDN).state = KEY_PRESSED
        aMouseState(MOUSE_WDN).lastdown = 0
    End If
End Sub

Private Sub AddMousePress(ByVal action As Long, ByVal key As Long, ByVal x As Long, ByVal y As Long)
    ReDim Preserve mpAction(0 To mpCount)
    ReDim Preserve mpKey(0 To mpCount)
    ReDim Preserve mpX(0 To mpCount)
    ReDim Preserve mpY(0 To mpCount)
    mpAction(mpCount) = action
    mpKey(mpCount) = key
    mpX(mpCount) = x
    mpY(mpCount) = y
    mpCount = mpCount + 1
End Sub

' Handle mouse button events (We can handle up to 5)
Private Sub inputHandleMouseButtonEvent(ByVal typ As Long, ByVal mouseKeyCode As Long, ByVal x As Long, ByVal y As Long)
    Dim rt As Long
    mouseXPos = Fix(CSng(x) / current_displayScaleFactor) And &HFFFF&
    mouseYPos = Fix(CSng(y) / current_displayScaleFactor) And &HFFFF&

    Select Case typ
    Case EV_MOUSE_BUTTON_DOWN
        AddMousePress MOUSEPRESS_Press, mouseKeyCode, mouseXPos, mouseYPos

        aMouseState(mouseKeyCode).pressPosX = mouseXPos
        aMouseState(mouseKeyCode).pressPosY = mouseYPos
        If aMouseState(mouseKeyCode).state = KEY_UP Or _
           aMouseState(mouseKeyCode).state = KEY_RELEASED Or _
           aMouseState(mouseKeyCode).state = KEY_PRESSRELEASE Then
            ' whether double click or not
            If Not game Is Nothing Then rt = game.realTime()
            If rt - aMouseState(mouseKeyCode).lastdown < DOUBLE_CLICK_INTERVAL Then
                aMouseState(mouseKeyCode).state = KEY_DOUBLECLICK
                aMouseState(mouseKeyCode).lastdown = 0
            Else
                aMouseState(mouseKeyCode).state = KEY_PRESSED
                aMouseState(mouseKeyCode).lastdown = rt
            End If

            If mouseKeyCode < MOUSE_X1 Then ' Assume they are draggin' with either LMB|RMB|MMB
                dragKey = mouseKeyCode
                dragX = mouseXPos
                dragY = mouseYPos
            End If
        End If
    Case EV_MOUSE_BUTTON_UP
        AddMousePress MOUSEPRESS_Release, mouseKeyCode, mouseXPos, mouseYPos

        aMouseState(mouseKeyCode).releasePosX = mouseXPos
        aMouseState(mouseKeyCode).releasePosY = mouseYPos
        If aMouseState(mouseKeyCode).state = KEY_PRESSED Then
            aMouseState(mouseKeyCode).state = KEY_PRESSRELEASE
        ElseIf aMouseState(mouseKeyCode).state = KEY_DOWN Or _
               aMouseState(mouseKeyCode).state = KEY_DRAG Or _
               aMouseState(mouseKeyCode).state = KEY_DOUBLECLICK Then
            aMouseState(mouseKeyCode).state = KEY_RELEASED
        End If
    End Select
End Sub

' Handle mousemotion events
Private Sub inputHandleMouseMotionEvent(ByVal x As Long, ByVal y As Long)
    ' store the current mouse position
    mouseXPos = Fix(CSng(x) / current_displayScaleFactor) And &HFFFF&
    mouseYPos = Fix(CSng(y) / current_displayScaleFactor) And &HFFFF&

    ' now see if a drag has started
    If aMouseState(dragKey).state = KEY_PRESSED Or aMouseState(dragKey).state = KEY_DOWN Then
        If Abs(dragX - mouseXPos) > DRAG_THRESHOLD Or Abs(dragY - mouseYPos) > DRAG_THRESHOLD Then
            aMouseState(dragKey).state = KEY_DRAG
        End If
    End If
End Sub

Public Sub StartTextInput(ByVal pTextInputRequester As Long)
    GetTextEventsOwner = pTextInputRequester
End Sub

Public Sub StopTextInput(ByVal pTextInputResigner As Long)
    If GetTextEventsOwner = pTextInputResigner Then GetTextEventsOwner = 0
End Sub

Public Function isInTextInputMode() As Boolean
    isInTextInputMode = (GetTextEventsOwner <> 0)
End Function

' This stage, we only setup keycodes.
Public Sub wzMain()
    initKeycodes
    mouseInWindow = True
    windowHasFocus = True
    current_displayScale = 100
    current_displayScaleFactor = 1!
    ' (VB) the window functions lib/framework calls
    Set frameworkPlatform = New PlatformFramework
    frameSetPlatform frameworkPlatform
End Sub

'--------------------------------------------------------------------------
' Window size

Public Sub handleGameScreenSizeChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    screenWidth = newWidth
    screenHeight = newHeight

    pie_SetVideoBufferWidth screenWidth
    pie_SetVideoBufferHeight screenHeight
    pie_UpdateSurfaceGeometry

    If Not resizePending Then
        ' The screen size change details are stored in scaled, logical units (points)
        ' i.e. the values expect by the game engine.
        resizePending = True
        resizeOldWidth = oldWidth
        resizeOldHeight = oldHeight
    End If
    ' update the new screen width / height, in case more than one resize message is processed this event loop
    resizeNewWidth = screenWidth
    resizeNewHeight = screenHeight
End Sub

Public Sub handleWindowSizeChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim oldScreenWidth As Long, oldScreenHeight As Long, newScreenWidth As Long, newScreenHeight As Long
    windowWidth = newWidth
    windowHeight = newHeight

    ' This function receives the window size in the window's logical units, not accounting for the interface
    ' scale factor: divide by it to get the new *game* screen logical width / height.
    oldScreenWidth = Fix(oldWidth / current_displayScaleFactor)
    oldScreenHeight = Fix(oldHeight / current_displayScaleFactor)
    newScreenWidth = Fix(newWidth / current_displayScaleFactor)
    newScreenHeight = Fix(newHeight / current_displayScaleFactor)

    handleGameScreenSizeChange oldScreenWidth, oldScreenHeight, newScreenWidth, newScreenHeight

    gfx_api_context_handleWindowSizeChange oldWidth, oldHeight, newWidth, newHeight
End Sub

Private Sub wzGetMinimumWindowSizeForDisplayScaleFactor(minWindowWidth As Long, minWindowHeight As Long, ByVal displayScaleFactor As Single)
    minWindowWidth = -Int(-(MIN_WZ_GAMESCREEN_WIDTH * displayScaleFactor))     ' ceil
    minWindowHeight = -Int(-(MIN_WZ_GAMESCREEN_HEIGHT * displayScaleFactor))
End Sub

Private Sub processScreenSizeChangeNotificationIfNeeded()
    If resizePending Then
        ' WZ must process the screen size change
        screen_updateGeometry ' must come after gfx_api::context::handleWindowSizeChange
        If Not game Is Nothing Then game.gameScreenSizeDidChange resizeOldWidth, resizeOldHeight, resizeNewWidth, resizeNewHeight
        resizePending = False
    End If
End Sub

'--------------------------------------------------------------------------
' Screen setup

' Registers the window class (C: wzSDLOneTimeInit).
Private Function wzSDLOneTimeInit() As Boolean
    Dim wc As WNDCLASSEXW
    If classRegistered Then
        wzSDLOneTimeInit = True
        Exit Function
    End If
    className = "WarzoneVB"
    wc.cbSize = LenB(wc)
    wc.style = CS_OWNDC
    wc.lpfnWndProc = FnPtr(AddressOf WzWndProc)
    wc.hInstance = GetModuleHandleW(0)
    wc.hCursor = LoadCursorW(0, IDC_ARROW)
    wc.lpszClassName = StrPtr(className)
    If RegisterClassExW(wc) = 0 Then
        debug_ 0, LOG_ERROR, "wzSDLOneTimeInit", "(VB) RegisterClassExW failed: %d", Err.LastDllError
        Exit Function
    End If
    classRegistered = True
    wzSDLOneTimeInit = True
End Function

' C: wzMainScreenSetup_CreateVideoWindow: windowed, on the primary display. Returns False on failure.
Private Function wzMainScreenSetup_CreateVideoWindow(ByVal backend As Long, ByVal antialiasing As Long, ByVal fullscreen As Long, ByVal highDPI As Boolean) As Boolean
    Dim desiredWindowWidth As Long, desiredWindowHeight As Long
    Dim minWindowWidth As Long, minWindowHeight As Long
    Dim displayUsableBounds As RECT, rc As RECT, x As Long, y As Long, w As Long, h As Long
    Dim resultingWidth As Long, resultingHeight As Long

    ' populate with the saved configuration values (if we had any) for windowed mode
    desiredWindowWidth = game.warGetWidth()
    desiredWindowHeight = game.warGetHeight()

    If desiredWindowWidth = 0 Or desiredWindowHeight = 0 Then
        windowWidth = 1024
        windowHeight = 768
    Else
        windowWidth = desiredWindowWidth
        windowHeight = desiredWindowHeight
    End If

    setDisplayScale game.warGetDisplayScale()

    debug_ 2777, LOG_WZ, "wzMainScreenSetup_CreateVideoWindow", "Monitor %d: pos %d x %d : res %d x %d", 0, 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN)

    ' Check desired screen index values versus current system (one display here)
    If game.warGetScreen() > 1 Then
        debug_ 2790, LOG_WARNING, "wzMainScreenSetup_CreateVideoWindow", "Invalid screen [%d] defined in configuration; there are only %d displays; falling back to display 0", game.warGetScreen(), 1
        game.warSetScreen 0
    End If

    screenIndex = game.warGetScreen()

    ' Calculate the minimum window size (in *logical* points) given the current display scale
    wzGetMinimumWindowSizeForDisplayScaleFactor minWindowWidth, minWindowHeight, current_displayScaleFactor

    If fullscreen <> WINDOW_MODE_fullscreen Then
        If windowWidth < minWindowWidth Then windowWidth = minWindowWidth
        If windowHeight < minWindowHeight Then windowHeight = minWindowHeight

        If fullscreen = WINDOW_MODE_windowed Then
            ' Determine the maximum usable windowed size for this display/screen
            If SystemParametersInfoW(SPI_GETWORKAREA, 0, displayUsableBounds, 0) <> 0 Then
                w = displayUsableBounds.Right - displayUsableBounds.Left
                h = displayUsableBounds.Bottom - displayUsableBounds.Top
                If w > 0 And h > 0 Then
                    If windowWidth > w Then windowWidth = w
                    If windowHeight > h Then windowHeight = h
                End If
            End If
        End If
    End If

    ' The client area gets windowWidth x windowHeight, centred in the usable area
    rc.Right = windowWidth
    rc.Bottom = windowHeight
    AdjustWindowRectEx rc, WS_OVERLAPPEDWINDOW, 0, 0
    w = rc.Right - rc.Left
    h = rc.Bottom - rc.Top
    x = displayUsableBounds.Left + ((displayUsableBounds.Right - displayUsableBounds.Left) - w) \ 2
    y = displayUsableBounds.Top + ((displayUsableBounds.Bottom - displayUsableBounds.Top) - h) \ 2
    If x < displayUsableBounds.Left Then x = displayUsableBounds.Left
    If y < displayUsableBounds.Top Then y = displayUsableBounds.Top

    minTrackWidth = minWindowWidth
    minTrackHeight = minWindowHeight
    lastClientWidth = windowWidth
    lastClientHeight = windowHeight
    lastSizeType = SIZE_RESTORED
    WZwindow = CreateWindowExW(0, StrPtr(className), StrPtr(PACKAGE_NAME), WS_OVERLAPPEDWINDOW, x, y, w, h, 0, 0, GetModuleHandleW(0), 0)
    If WZwindow = 0 Then
        debug_ 2879, LOG_INFO, "wzMainScreenSetup_CreateVideoWindow", "Failed to create a window, because: %s", "CreateWindowExW failed (" & Err.LastDllError & ")"
        Exit Function
    End If
    queueEnabled = True
    ShowWindow WZwindow, SW_SHOW
    SetForegroundWindow WZwindow

    ' Get resulting logical window size
    GetClientRect WZwindow, rc
    resultingWidth = rc.Right - rc.Left
    resultingHeight = rc.Bottom - rc.Top
    lastClientWidth = resultingWidth
    lastClientHeight = resultingHeight
    If Not (resultingWidth > 0 And resultingHeight > 0) Then
        ASSERT_FAILURE "resultingWidth > 0 && resultingHeight > 0", "lib/sdl/main_sdl.cpp:2951", "wzMainScreenSetup_CreateVideoWindow", _
            "Invalid - SDL returned window width: %d, window height: %d", resultingWidth, resultingHeight
    End If
    windowWidth = resultingWidth
    windowHeight = resultingHeight

    ' Calculate the game screen's logical dimensions
    screenWidth = windowWidth
    screenHeight = windowHeight
    If current_displayScaleFactor > 1! Then
        screenWidth = Fix(windowWidth / current_displayScaleFactor)
        screenHeight = Fix(windowHeight / current_displayScaleFactor)
    End If
    pie_SetVideoBufferWidth screenWidth
    pie_SetVideoBufferHeight screenHeight

    wzMainScreenSetup_CreateVideoWindow = True
End Function

Private Function to_display_string(ByVal backend As Long) As String
    Select Case backend
    Case video_backend_opengl: to_display_string = "OpenGL"
    Case video_backend_opengles: to_display_string = "OpenGL ES"
    Case video_backend_vulkan: to_display_string = "Vulkan"
    End Select
End Function

Private Function wzAttemptInitializeBackend(ByVal hasBackend As Boolean, ByVal backend As Long, ByVal antialiasing As Long, ByVal fullscreen As Long, ByVal swapMode As Long, ByVal lodDistanceBias As Single, ByVal depthMapResolution As Long) As Boolean
    Dim impl As win_OpenGL_Impl, gfxapi_backend As backend_type
    If hasBackend Then
        debug_ 3124, LOG_INFO, "wzAttemptInitializeBackend", "Attempting to initialize backend: %s", to_display_string(backend)
    Else
        debug_ 3128, LOG_INFO, "wzAttemptInitializeBackend", "Attempting to initialize backend: (headless)"
    End If

    If hasBackend Then
        If Not wzMainScreenSetup_CreateVideoWindow(backend, antialiasing, fullscreen, True) Then
            Exit Function
        End If
    End If

    gfxapi_backend = backend_type_null_backend
    If hasBackend Then
        gfxapi_backend = backend_type_opengl_backend ' (VB) OpenGL only
    End If

    ' (VB) SDL_gfx_api_Impl_Factory: the implementation for the backend type
    If gfxapi_backend = backend_type_null_backend Then
        queueEnabled = True       ' (VB) SDL's event queue works without a window (wzQuit's quit event in headless mode)
        If Not gfx_api_context_initialize(New win_Null_Impl, antialiasing, swapMode, lodDistanceBias, depthMapResolution, gfxapi_backend) Then
            ' headless mode failed in gfx_api::context::initialize??
            debug_ 3184, LOG_FATAL, "wzAttemptInitializeBackend", "gfx_api::context::get().initialize failed for headless-mode backend?"
            exit_ 1                   ' EXIT_FAILURE (SDL_Quit: nothing to quit)
        End If
        wzAttemptInitializeBackend = True
        Exit Function
    End If
    Set impl = New win_OpenGL_Impl
    impl.Init WZwindow, backend = video_backend_opengles, False
    If Not gfx_api_context_initialize(impl, antialiasing, swapMode, lodDistanceBias, depthMapResolution, gfxapi_backend) Then
        ' Failed to initialize desired backend / renderer settings: destroy the window
        If WZwindow <> 0 Then
            queueEnabled = False
            DestroyWindow WZwindow
            WZwindow = 0
        End If
        Exit Function
    End If

    wzAttemptInitializeBackend = True
End Function

' lib/ivis_opengl/gfx_api.h's inline to_swap_mode (value: the config's vsync)
Private Function to_swap_mode(ByVal value As Long) As Long
    Select Case value
    Case swap_interval_mode_immediate
        to_swap_mode = swap_interval_mode_immediate
        Exit Function
    Case swap_interval_mode_vsync
        to_swap_mode = swap_interval_mode_vsync
        Exit Function
    Case swap_interval_mode_adaptive_vsync
        to_swap_mode = swap_interval_mode_adaptive_vsync
        Exit Function
    Case Else
        debug_ 1352, LOG_WARNING, "to_swap_mode", "Invalid vsync value (%d); defaulting to vsync ON", value
    End Select
    to_swap_mode = swap_interval_mode_vsync
End Function

' backend: hasBackend = False for headless (C: nullopt). OpenGL only; Vulkan and OpenGL ES requests use OpenGL.
' C: the first of wzAvailableGfxBackends() (on Windows Vulkan, where it
' works). (VB) The port has only the OpenGL backend.
Public Function wzGetDefaultGfxBackendForCurrentSystem() As video_backend
    wzGetDefaultGfxBackendForCurrentSystem = video_backend_opengl
End Function

Public Function wzMainScreenSetup(ByVal hasBackend As Boolean, ByVal backend As video_backend, ByVal antialiasing As Long, ByVal fullscreen As WINDOW_MODE, ByVal vsync As Long, ByVal lodDistanceBiasPercentage As Long, ByVal depthMapResolution As Long, Optional ByVal highDPI As Boolean = True) As Boolean
    Dim swapMode As Long, lodDistanceBias As Single

    If Not wzSDLOneTimeInit() Then
        Exit Function
    End If

    swapMode = to_swap_mode(vsync)

    If lodDistanceBiasPercentage <> 0 Then
        lodDistanceBias = CSng(lodDistanceBiasPercentage) / 100!
    End If

    If hasBackend And backend <> video_backend_opengl Then backend = video_backend_opengl
    WZbackendHasValue = hasBackend ' various other functions might need this before the call to wzAttemptInitializeBackend returns
    WZbackend = backend
    If Not wzAttemptInitializeBackend(hasBackend, backend, antialiasing, fullscreen, swapMode, lodDistanceBias, depthMapResolution) Then
        WZbackendHasValue = False
        Exit Function
    End If
    game.warSetDisplayScale wzGetCurrentDisplayScale() ' persist any auto-adjustments to the display scale
    wzMainScreenSetup = True
End Function

'--------------------------------------------------------------------------
' Events

Private Sub handleActiveEvent(ev As WzEvent)
    Dim oldWindowWidth As Long, oldWindowHeight As Long, newWindowWidth As Long, newWindowHeight As Long
    Dim rc As RECT, i As Long
    Select Case ev.typ
    Case EV_WINDOW_SHOWN
        debug_ 3540, LOG_WZ, "handleActiveEvent", "Window %d shown", 1
    Case EV_WINDOW_HIDDEN
        debug_ 3543, LOG_WZ, "handleActiveEvent", "Window %d hidden", 1
    Case EV_WINDOW_MOVED
        debug_ 3549, LOG_WZ, "handleActiveEvent", "Window %d moved to %d,%d", 1, ev.a, ev.b
    Case EV_WINDOW_RESIZED
        debug_ 3554, LOG_WZ, "handleActiveEvent", "Window %d resized to %dx%d", 1, ev.a, ev.b
        oldWindowWidth = windowWidth
        oldWindowHeight = windowHeight

        GetClientRect WZwindow, rc
        newWindowWidth = rc.Right - rc.Left
        newWindowHeight = rc.Bottom - rc.Top
        handleWindowSizeChange oldWindowWidth, oldWindowHeight, newWindowWidth, newWindowHeight

        ' Store the new values (in case the user manually resized the window bounds)
        If wzGetCurrentWindowMode() = WINDOW_MODE_windowed Then
            game.warSetWidth newWindowWidth
            game.warSetHeight newWindowHeight
        End If
    Case EV_WINDOW_MINIMIZED
        debug_ 3594, LOG_INFO, "handleActiveEvent", "Window %d minimized", 1
    Case EV_WINDOW_MAXIMIZED
        debug_ 3597, LOG_WZ, "handleActiveEvent", "Window %d maximized", 1
    Case EV_WINDOW_RESTORED
        debug_ 3600, LOG_INFO, "handleActiveEvent", "Window %d restored", 1
        oldWindowWidth = windowWidth
        oldWindowHeight = windowHeight
        GetClientRect WZwindow, rc
        newWindowWidth = rc.Right - rc.Left
        newWindowHeight = rc.Bottom - rc.Top
        If oldWindowWidth <> newWindowWidth Or oldWindowHeight <> newWindowHeight Then
            debug_ 3615, LOG_WZ, "handleActiveEvent", "Triggering handleWindowSizeChange from window restore event"
            handleWindowSizeChange oldWindowWidth, oldWindowHeight, newWindowWidth, newWindowHeight
        End If
    Case EV_WINDOW_MOUSE_ENTER
        mouseInWindow = True
        debug_ 3622, LOG_WZ, "handleActiveEvent", "Mouse entered window %d", 1
    Case EV_WINDOW_MOUSE_LEAVE
        mouseInWindow = False
        debug_ 3627, LOG_WZ, "handleActiveEvent", "Mouse left window %d", 1
    Case EV_WINDOW_FOCUS_GAINED
        windowHasFocus = True
        debug_ 3632, LOG_WZ, "handleActiveEvent", "Window %d gained keyboard focus", 1
    Case EV_WINDOW_FOCUS_LOST
        windowHasFocus = False
        debug_ 3636, LOG_WZ, "handleActiveEvent", "Window %d lost keyboard focus", 1
        For i = 0 To KEY_MAXSCAN - 1
            aKeyState(i).state = KEY_UP
            actualKeyState(i).state = KEY_UP
        Next
    Case EV_WINDOW_CLOSE_REQUESTED
        debug_ 3644, LOG_WZ, "handleActiveEvent", "Window %d closed", 1
    End Select
End Sub

' Returns False when the loop must stop (C: *bContinue = false).
Private Function wzEventLoopOneFrame() As Boolean
    Dim m As MSG_t, ev As WzEvent

    ' Deal with any windows messages
    Do While PeekMessageW(m, 0, 0, 0, PM_REMOVE) <> 0
        TranslateMessage m
        DispatchMessageW m
    Loop
    Do While PollEvent(ev)
        Select Case ev.typ
        Case EV_KEY_UP, EV_KEY_DOWN
            inputHandleKeyEvent ev.typ, ev.a, ev.b
        Case EV_MOUSE_BUTTON_UP, EV_MOUSE_BUTTON_DOWN
            inputHandleMouseButtonEvent ev.typ, ev.a, ev.b, ev.c
        Case EV_MOUSE_MOTION
            inputHandleMouseMotionEvent ev.b, ev.c
        Case EV_MOUSE_WHEEL
            inputHandleMouseWheelEvent ev.a, ev.b
        Case EV_TEXT_INPUT
            inputhandleText ev.a
        Case EV_QUIT
            Exit Function
        Case Else
            handleActiveEvent ev
        End Select
    Loop

    processScreenSizeChangeNotificationIfNeeded
    game.mainLoop       ' WZ does its thing
    inputNewFrame       ' reset input states

    ' (VB) scripted test runs
    framesDone = framesDone + 1
    If quitAfterFrames > 0 And framesDone >= quitAfterFrames Then wzQuit 0
    wzEventLoopOneFrame = True
End Function

' Actual mainloop. callOnShutdown: whether to call the game's mainShutdown afterwards (C: onShutdown).
Public Sub wzMainEventLoop(ByVal callOnShutdown As Boolean)
    Dim bContinue As Boolean
    bContinue = True
    Do While bContinue
        bContinue = wzEventLoopOneFrame()
    Loop

    If callOnShutdown Then
        game.mainShutdown
    End If
End Sub

Public Sub wzPumpEventsWhileLoading()
    Dim m As MSG_t
    Do While PeekMessageW(m, 0, 0, 0, PM_REMOVE) <> 0
        TranslateMessage m
        DispatchMessageW m
    Loop
End Sub

Public Sub wzShutdown()
    ' order is important!
    queueEnabled = False
    If WZwindow <> 0 Then
        DestroyWindow WZwindow
        WZwindow = 0
    End If
    If classRegistered Then
        UnregisterClassW StrPtr(className), GetModuleHandleW(0)
        classRegistered = False
    End If
    ' (VB) the cross-project references
    frameSetPlatform Nothing
    Set frameworkPlatform = Nothing
End Sub
