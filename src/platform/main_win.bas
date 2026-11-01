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
' wzGetCurrentDisplayScale, wzMainScreenSetup_CreateVideoWindow,
' wzAttemptInitializeBackend and wzMainScreenSetup (OpenGL, no fallback),
' handleActiveEvent, wzEventLoopOneFrame, wzMainEventLoop, wzShutdown,
' StartTextInput, StopTextInput, isInTextInputMode; the window modes
' (windowed, fullscreen: borderless or with a display mode set through
' ChangeDisplaySettingsEx, as SDL's Windows driver does), the display list
' and modes, the display and cursor scales, the clipboard, the mouse grab,
' minimize on focus loss, the system's RAM and CPUs (cluster S, fork S6;
' cursors_sdl.cpp is cursors_win). SDL's own parts these need are win_*
' functions at the end. The displays are SDL's: ids from 1, the primary
' first. The process is not DPI-aware: the display content scale is 1.
' TEST-ONLY: WZ_TEST_NO_DESKTOP_EFFECTS logs what would change the user's
' desktop (the display mode, a fullscreen window, the cursor clip, the
' clipboard, opening links) instead of doing it; set to "borderless", it
' makes a borderless ("Desktop Full") fullscreen window and logs the rest.
' Missing: the window icon, Vulkan, the backend fallback and its prompts,
' wzAsyncExecOnMainThread (but the fullscreen transition at start), touch
' input, Emscripten.
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

' (VB) the display, window, clipboard and system functions (cluster S, fork S6)
Private Type DEVMODEW
    dmDeviceName(0 To 63) As Byte
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmPositionX As Long
    dmPositionY As Long
    dmDisplayOrientation As Long
    dmDisplayFixedOutput As Long
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName(0 To 63) As Byte
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

Private Type DISPLAY_DEVICEW
    cb As Long
    DeviceName(0 To 63) As Byte
    DeviceString(0 To 255) As Byte
    StateFlags As Long
    DeviceID(0 To 255) As Byte
    DeviceKey(0 To 255) As Byte
End Type

Private Type MONITORINFOEXW
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
    szDevice(0 To 63) As Byte
End Type

Private Type SYSTEM_INFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Type MEMORYSTATUSEX_
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As Currency
    ullAvailPhys As Currency
    ullTotalPageFile As Currency
    ullAvailPageFile As Currency
    ullTotalVirtual As Currency
    ullAvailVirtual As Currency
    ullAvailExtendedVirtual As Currency
End Type

Private Declare Function EnumDisplayDevicesW Lib "user32" (ByVal lpDevice As Long, ByVal iDevNum As Long, lpDisplayDevice As DISPLAY_DEVICEW, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettingsExW Lib "user32" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As DEVMODEW, ByVal dwFlags As Long) As Long
Private Declare Function ChangeDisplaySettingsExW Lib "user32" (ByVal lpszDeviceName As Long, lpDevMode As Any, ByVal hwnd As Long, ByVal dwflags As Long, ByVal lParam As Long) As Long
Private Declare Function MonitorFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfoW Lib "user32" (ByVal hMonitor As Long, lpmi As MONITORINFOEXW) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal format As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX_) As Long
Private Declare Function ShellExecuteW Lib "shell32" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const WM_SETCURSOR As Long = &H20
Private Const HTCLIENT As Long = 1
Private Const GWL_STYLE As Long = -16
Private Const WS_POPUP As Long = &H80000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_THICKFRAME As Long = &H40000
Private Const STYLE_NORMAL As Long = &HCA0000                 ' WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX
Private Const STYLE_RESIZABLE As Long = WS_THICKFRAME Or WS_MAXIMIZEBOX
Private Const STYLE_FULLSCREEN As Long = WS_POPUP Or WS_MINIMIZEBOX
Private Const STYLE_MASK As Long = STYLE_NORMAL Or STYLE_RESIZABLE Or STYLE_FULLSCREEN
Private Const HWND_TOP As Long = 0
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOCOPYBITS As Long = &H100
Private Const SW_MINIMIZE As Long = 6
Private Const SW_SHOWNORMAL As Long = 1
Private Const ENUM_CURRENT_SETTINGS As Long = -1
Private Const ENUM_REGISTRY_SETTINGS As Long = -2
Private Const CDS_FULLSCREEN As Long = &H4
Private Const DISP_CHANGE_SUCCESSFUL As Long = 0
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const DM_DISPLAYFLAGS As Long = &H200000
Private Const DM_DISPLAYFREQUENCY As Long = &H400000
Private Const DISPLAY_DEVICE_ATTACHED_TO_DESKTOP As Long = &H1
Private Const DISPLAY_DEVICE_PRIMARY_DEVICE As Long = &H4
Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = &H2
Private Const LOGPIXELSX As Long = 88
' enum class MinimizeOnFocusLossBehavior
Private Const MinimizeOnFocusLossBehavior_Auto As Long = -1
Private Const MinimizeOnFocusLossBehavior_Off As Long = 0
Private Const MinimizeOnFocusLossBehavior_On_Fullscreen As Long = 1

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

' main_sdl.cpp's displaylist: std::vector<optional<screeninfo>> (dlHas False: nullopt, "Desktop Full")
Private displaylistCount As Long
Private dlHas() As Boolean
Private dlScreen() As Long
Private dlWidth() As Long
Private dlHeight() As Long
Private dlPixelDensity() As Single
Private dlRefreshRate() As Single
' (VB) SDL's displays: ids 1 .. displayCount (the primary first), their device names
Private displayCount As Long
Private displayDeviceName() As String
' (VB) SDL's fullscreen display modes of each display, as SDL_AddFullscreenDisplayMode keeps and sorts them
Private fsModeCount As Long
Private fsDisplay() As Long
Private fsW() As Long
Private fsH() As Long
Private fsBpp() As Long
Private fsRefresh() As Single
Private fsFrequency() As Long                   ' DEVMODE's dmDisplayFrequency
Private fsDisplayFlags() As Long
' (VB) SDL's window state: its fullscreen mode (SDL_SetWindowFullscreenMode; False: borderless "Desktop Full"),
' SDL_WINDOW_FULLSCREEN, the display mode changed for it (to restore), its style, the mouse grab
Private hasFullscreenMode As Boolean
Private fullscreenModeIndex As Long
Private windowIsFullscreen As Boolean
Private exclusiveModeApplied As Boolean
Private exclusiveModeDisplay As Long
Private windowStyle As Long
Private windowResizable As Boolean               ' SDL_WINDOW_RESIZABLE
Private windowedRestoreRect As RECT
Private mouseGrabbed As Boolean
Private cursorClipped As Boolean
' SDL_HINT_VIDEO_MINIMIZE_ON_FOCUS_LOSS ("": not set)
Private minimizeOnFocusLossHint As String
' main_sdl.cpp's deferredDimensionReset (optional<QueuedWindowDimensions>)
Private deferredDimensionResetHasValue As Boolean
Private deferredDimensionResetWidth As Long
Private deferredDimensionResetHeight As Long
Private deferredDimensionResetRecenter As Boolean
' (VB) wzAsyncExecOnMainThread([]() { wzChangeWindowMode(WINDOW_MODE::fullscreen); }): the next frame
Private pendingFullscreenTransition As Boolean
' SDL's SDL_cursor (the cursor shown in the client area; 0: the class's arrow)
Private SDL_cursor As Long
' TEST-ONLY (WZ_TEST_NO_DESKTOP_EFFECTS set): what would change the user's desktop (the display mode, a
' fullscreen window, the cursor clip, the clipboard, opening a link) is logged instead, and the clipboard is
' this process's own
Private testNoDesktopEffects As Boolean
' TEST-ONLY (WZ_TEST_NO_DESKTOP_EFFECTS=borderless): a borderless fullscreen window is made all the same (it
' changes no display mode); a fullscreen display mode is still only logged
Private testAllowBorderless As Boolean
Private testClipboardHas As Boolean
Private testClipboardText As String

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
    Case WM_SETCURSOR
        If (lParam And &HFFFF&) = HTCLIENT And SDL_cursor <> 0 Then
            SetCursor SDL_cursor
            WzWndProc = 1
            Exit Function
        End If
    Case WM_GETMINMAXINFO
        If minTrackWidth > 0 And Not windowIsFullscreen Then
            rc.Right = minTrackWidth
            rc.Bottom = minTrackHeight
            AdjustWindowRectEx rc, windowStyle, 0, 0
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
    If Not WZbackendHasValue Then
        ' return a dummy value
        wzGetCurrentWindowMode = WINDOW_MODE_windowed
        Exit Function
    End If
    If WZwindow = 0 Then
        debug_ 694, LOG_WARNING, "wzGetCurrentWindowMode", "wzGetCurrentWindowMode called when window is not available"
        wzGetCurrentWindowMode = WINDOW_MODE_windowed
        Exit Function
    End If

    If windowIsFullscreen Then          ' SDL_GetWindowFlags(WZwindow) & SDL_WINDOW_FULLSCREEN
        wzGetCurrentWindowMode = WINDOW_MODE_fullscreen
    Else
        wzGetCurrentWindowMode = WINDOW_MODE_windowed
    End If
End Function

Public Function wzChangeWindowMode(ByVal mode As WINDOW_MODE, Optional ByVal silent As Boolean = False) As Boolean
    Dim previousMode As Long, currDisplayId As Long, desiredWidth As Long, desiredHeight As Long, displayUsableBounds As RECT
    previousMode = wzGetCurrentWindowMode()
    If previousMode = mode Then
        ' already in this mode
        wzChangeWindowMode = True
        Exit Function
    End If

    If Not wzIsSupportedWindowMode(mode) Then
        ' not a supported mode on this system
        Exit Function
    End If

    If Not silent Then
        debug_ 785, LOG_INFO, "wzChangeWindowMode", "Changing window mode: %s -> %s", to_display_string_WINDOW_MODE(previousMode), to_display_string_WINDOW_MODE(mode)
    End If

    Select Case mode
    Case WINDOW_MODE_windowed
        currDisplayId = win_getDisplayForWindow()
        If currDisplayId = 0 Then
            currDisplayId = screenIndex
        End If
        ' disable fullscreen mode
        If Not win_setWindowFullscreen(False) Then Exit Function
        wzSetWindowIsResizable True
        ' Determine the maximum usable windowed size for this display/screen, and cap the desired window size at that
        desiredWidth = game.warGetWidth()
        desiredHeight = game.warGetHeight()
        If win_getDisplayUsableBounds(currDisplayId, displayUsableBounds) Then
            If displayUsableBounds.Right - displayUsableBounds.Left > 0 And displayUsableBounds.Bottom - displayUsableBounds.Top > 0 Then
                If displayUsableBounds.Right - displayUsableBounds.Left < desiredWidth Then desiredWidth = displayUsableBounds.Right - displayUsableBounds.Left
                If displayUsableBounds.Bottom - displayUsableBounds.Top < desiredHeight Then desiredHeight = displayUsableBounds.Bottom - displayUsableBounds.Top
            End If
        End If
        ' restore the old windowed size
        win_setWindowSize desiredWidth, desiredHeight
        ' Position the window (centered) on the screen (for its new size)
        win_centerWindowOnDisplay currDisplayId
        ' Workaround for issues properly restoring window dimensions when changing from fullscreen -> windowed (on some platforms)
        deferredDimensionResetHasValue = True
        deferredDimensionResetWidth = desiredWidth
        deferredDimensionResetHeight = desiredHeight
        deferredDimensionResetRecenter = True
    Case WINDOW_MODE_fullscreen
        If Not win_setWindowFullscreen(True) Then
            Exit Function
        End If
        wzSetWindowIsResizable False
    End Select

    wzChangeWindowMode = True
End Function

Public Function wzIsFullscreen() As Boolean
    If WZwindow = 0 Then
        ' NOTE: Can't use debug(...) to log here or it will risk overwriting fatal error messages,
        ' as `_debug(...)` calls this function
        Exit Function
    End If
    wzIsFullscreen = windowIsFullscreen     ' SDL_GetWindowFlags(WZwindow) & SDL_WINDOW_FULLSCREEN
End Function

Public Function wzWindowHasFocus() As Boolean
    wzWindowHasFocus = windowHasFocus
End Function

Public Function wzIsMaximized() As Boolean
    If WZwindow = 0 Then
        debug_ 868, LOG_WARNING, "wzIsMaximized", "wzIsMaximized called when window is not available"
        Exit Function
    End If
    wzIsMaximized = (IsZoomed(WZwindow) <> 0)   ' SDL_GetWindowFlags(WZwindow) & SDL_WINDOW_MAXIMIZED
End Function

Public Sub wzGrabMouse()
    If Not WZbackendHasValue Then
        Exit Sub
    End If
    If WZwindow = 0 Then
        debug_ 909, LOG_WARNING, "wzGrabMouse", "wzGrabMouse called when window is not available - ignoring"
        Exit Sub
    End If
    mouseGrabbed = True                 ' SDL_SetWindowMouseGrab(WZwindow, true)
    win_updateClipCursor
    If testNoDesktopEffects Then debug_ 0, LOG_INFO, "wzGrabMouse", "(VB) TEST-ONLY: SDL_SetWindowMouseGrab(true): the clip follows the focus (focus: %d)", CLng(windowHasFocus)
End Sub

Public Sub wzReleaseMouse()
    If Not WZbackendHasValue Then
        Exit Sub
    End If
    If WZwindow = 0 Then
        debug_ 923, LOG_WARNING, "wzReleaseMouse", "wzReleaseMouse called when window is not available - ignoring"
        Exit Sub
    End If
    mouseGrabbed = False                ' SDL_SetWindowMouseGrab(WZwindow, false)
    win_updateClipCursor
    If testNoDesktopEffects Then debug_ 0, LOG_INFO, "wzReleaseMouse", "(VB) TEST-ONLY: SDL_SetWindowMouseGrab(false): the clip follows the focus (focus: %d)", CLng(windowHasFocus)
End Sub

'--------------------------------------------------------------------------
' The clipboard (SDL's on Win32: CF_UNICODETEXT, LF written as CRLF and read back as LF)

' See if we have TEXT in the clipboard
Private Function has_scrap() As Boolean
    If testNoDesktopEffects Then
        has_scrap = testClipboardHas
        Exit Function
    End If
    has_scrap = (IsClipboardFormatAvailable(CF_UNICODETEXT) <> 0)   ' SDL_HasClipboardText()
End Function

' Set the clipboard text
Public Function wzSetClipboardText(ByVal src As String) As Boolean
    If Not win_setClipboardText(src) Then
        debug_ 247, LOG_ERROR, "wzSetClipboardText", "Could not put clipboard text because : %s", "(VB) the clipboard could not be opened or set (" & Err.LastDllError & ")"
        Exit Function
    End If
    wzSetClipboardText = True
End Function

' Get text from the clipboard
Private Function get_scrap(dst As String) As Boolean
    If has_scrap() Then
        If Not win_getClipboardText(dst) Then
            debug_ 261, LOG_ERROR, "get_scrap", "Could not get clipboard text because : %s", "(VB) the clipboard could not be opened or read (" & Err.LastDllError & ")"
            Exit Function
        End If
        get_scrap = True
    Else
        ' wasn't text or no text in the clipboard
    End If
End Function

' SDL_SetClipboardText on Win32
Private Function win_setClipboardText(ByVal src As String) As Boolean
    Dim s As String, h As Long, p As Long, n As Long
    If testNoDesktopEffects Then
        testClipboardText = src
        testClipboardHas = True
        debug_ 0, LOG_INFO, "win_setClipboardText", "(VB) TEST-ONLY: clipboard text set (this process's, not the user's): %s", src
        win_setClipboardText = True
        Exit Function
    End If
    ' LF to CRLF
    s = Replace(Replace(src, vbCrLf, vbLf), vbLf, vbCrLf)
    If OpenClipboard(WZwindow) = 0 Then Exit Function
    EmptyClipboard
    n = (Len(s) + 1) * 2
    h = GlobalAlloc(GMEM_MOVEABLE, n)
    If h <> 0 Then
        p = GlobalLock(h)
        CopyMemory ByVal p, ByVal StrPtr(s), n
        GlobalUnlock h
        If SetClipboardData(CF_UNICODETEXT, h) = 0 Then
            GlobalFree h
            h = 0
        End If
    End If
    CloseClipboard
    win_setClipboardText = (h <> 0)
End Function

' SDL_GetClipboardText on Win32
Private Function win_getClipboardText(dst As String) As Boolean
    Dim h As Long, p As Long, n As Long, s As String
    If testNoDesktopEffects Then
        dst = testClipboardText
        win_getClipboardText = True
        Exit Function
    End If
    If OpenClipboard(WZwindow) = 0 Then Exit Function
    h = GetClipboardData(CF_UNICODETEXT)
    If h <> 0 Then
        p = GlobalLock(h)
        If p <> 0 Then
            n = lstrlenW(p)
            s = String$(n, vbNullChar)
            If n > 0 Then CopyMemory ByVal StrPtr(s), ByVal p, n * 2
            GlobalUnlock h
            ' CRLF to LF
            dst = Replace(s, vbCrLf, vbLf)
            win_getClipboardText = True
        End If
    End If
    CloseClipboard
End Function

' Put a character into a text buffer overwriting any text under the cursor
Public Function wzGetSelection() As String
    Dim scrap As String
    If get_scrap(scrap) Then
        wzGetSelection = scrap
    End If
End Function

' SDL_GetTouchDevices: SDL lists the touch devices it has seen events from; the port handles no touch events
Private Function wzHasTouchInputDevices() As Boolean
    Dim numDevices As Long
    numDevices = 0
    wzHasTouchInputDevices = numDevices > 0
End Function

Public Function wzSeemsLikeNonTouchPlatform() As Boolean
    Dim hasScreenKeyboardSupport As Boolean     ' SDL_HasScreenKeyboardSupport(): the port drives no screen keyboard
    hasScreenKeyboardSupport = False
    wzSeemsLikeNonTouchPlatform = Not wzHasTouchInputDevices() Or (hasScreenKeyboardSupport = False)
End Function

'--------------------------------------------------------------------------
' Displays and display modes

' std::vector<optional<screeninfo>> wzAvailableResolutions(): (VB) the count, each entry's fields in the arrays
' (has_(i) False: nullopt, "Desktop Full")
Public Function wzAvailableResolutions(screens() As Long, widths() As Long, heights() As Long, pixelDensities() As Single, refreshRates() As Single, has_() As Boolean) As Long
    Dim i As Long
    If displaylistCount = 0 Then Exit Function
    ReDim screens(0 To displaylistCount - 1)
    ReDim widths(0 To displaylistCount - 1)
    ReDim heights(0 To displaylistCount - 1)
    ReDim pixelDensities(0 To displaylistCount - 1)
    ReDim refreshRates(0 To displaylistCount - 1)
    ReDim has_(0 To displaylistCount - 1)
    For i = 0 To displaylistCount - 1
        has_(i) = dlHas(i)
        screens(i) = dlScreen(i)
        widths(i) = dlWidth(i)
        heights(i) = dlHeight(i)
        pixelDensities(i) = dlPixelDensity(i)
        refreshRates(i) = dlRefreshRate(i)
    Next
    wzAvailableResolutions = displaylistCount
End Function

' optional<screeninfo> wzGetCurrentFullscreenDisplayMode(): has_value, the value in the ByRef fields
Public Function wzGetCurrentFullscreenDisplayMode(screen As Long, width As Long, height As Long, pixel_density As Single, refresh_rate As Single) As Boolean
    Dim displayId As Long
    If WZwindow = 0 Then
        debug_ 384, LOG_WARNING, "wzGetCurrentFullscreenDisplayMode", "wzGetCurrentFullscreenDisplayMode called when window is not available"
        Exit Function
    End If
    displayId = win_getDisplayForWindow()
    If displayId = 0 Then
        debug_ 390, LOG_WZ, "wzGetCurrentFullscreenDisplayMode", "Failed to get current screen index: %s", "(VB) the window's monitor is none of the displays"
        displayId = 1
    End If
    If Not hasFullscreenMode Then       ' SDL_GetWindowFullscreenMode(WZwindow) == nullptr
        ' This actually means the current setting is borderless fullscreen
        Exit Function
    End If
    screen = displayId
    width = fsW(fullscreenModeIndex)
    height = fsH(fullscreenModeIndex)
    pixel_density = 1!
    refresh_rate = fsRefresh(fullscreenModeIndex)
    wzGetCurrentFullscreenDisplayMode = True
End Function

' std::vector<unsigned int> wzAvailableDisplayScales(): (VB) into scales(0 .. n - 1), n returned
Public Function wzAvailableDisplayScales(scales() As Long) As Long
    ReDim scales(0 To 7)
    scales(0) = 100
    scales(1) = 125
    scales(2) = 150
    scales(3) = 200
    scales(4) = 250
    scales(5) = 300
    scales(6) = 400
    scales(7) = 500
    wzAvailableDisplayScales = 8
End Function

' std::vector<video_backend> wzAvailableGfxBackends(): (VB) into result(0 .. n - 1), n returned. Windows on x86
' offers Vulkan, OpenGL, DirectX (ANGLE) and OpenGL ES; the port has only the OpenGL backend
' (wzMainScreenSetup takes OpenGL for any other), so it offers OpenGL alone.
Public Function wzAvailableGfxBackends(result() As Long) As Long
    ReDim result(0 To 0)
    result(0) = video_backend_opengl
    wzAvailableGfxBackends = 1
End Function

' std::vector<WINDOW_MODE> wzSupportedWindowModes(): (VB) into modes(0 .. n - 1), n returned
Private Function wzSupportedWindowModes(modes() As Long) As Long
    ReDim modes(0 To 1)
    modes(0) = WINDOW_MODE_fullscreen
    modes(1) = WINDOW_MODE_windowed
    wzSupportedWindowModes = 2
End Function

Public Function wzIsSupportedWindowMode(ByVal mode As WINDOW_MODE) As Boolean
    Dim supportedModes() As Long, n As Long, i As Long
    n = wzSupportedWindowModes(supportedModes)
    For i = 0 To n - 1
        If supportedModes(i) = mode Then
            wzIsSupportedWindowMode = True
            Exit Function
        End If
    Next
End Function

Public Function wzAltEnterToggleFullscreen() As WINDOW_MODE
    Dim mode As Long
    mode = wzGetCurrentWindowMode()
    Select Case mode
    Case WINDOW_MODE_fullscreen
        If wzChangeWindowMode(WINDOW_MODE_windowed) Then
            mode = WINDOW_MODE_windowed
        End If
    Case WINDOW_MODE_windowed
        If wzChangeWindowMode(WINDOW_MODE_fullscreen) Then
            mode = WINDOW_MODE_fullscreen
        End If
    End Select
    wzAltEnterToggleFullscreen = mode
End Function

Public Sub wzPostChangedSwapInterval()
    ' currently, no-op
End Sub

'--------------------------------------------------------------------------
' Display scale and cursor scale

Private Sub wzGetMaximumDisplayScaleFactorsForWindowSize(ByVal width As Long, ByVal height As Long, horizScaleFactor As Single, vertScaleFactor As Single)
    horizScaleFactor = CSng(width) / CSng(MIN_WZ_GAMESCREEN_WIDTH)
    vertScaleFactor = CSng(height) / CSng(MIN_WZ_GAMESCREEN_HEIGHT)
End Sub

Private Function wzGetMaximumDisplayScaleFactorForWindowSize(ByVal width As Long, ByVal height As Long) As Single
    Dim maxHorizScaleFactor As Single, maxVertScaleFactor As Single
    wzGetMaximumDisplayScaleFactorsForWindowSize width, height, maxHorizScaleFactor, maxVertScaleFactor
    If maxHorizScaleFactor < maxVertScaleFactor Then
        wzGetMaximumDisplayScaleFactorForWindowSize = maxHorizScaleFactor
    Else
        wzGetMaximumDisplayScaleFactorForWindowSize = maxVertScaleFactor
    End If
End Function

' returns: the maximum display scale percentage (sourced from wzAvailableDisplayScales), or 0 if window is below the minimum required size for the minimum supported display scale
Private Function wzGetMaximumDisplayScaleForWindowSize(ByVal width As Long, ByVal height As Long) As Long
    Dim maxDisplayScaleFactor As Single, maxDisplayScalePercentage As Long
    Dim availableDisplayScales() As Long, n As Long, maxDisplayScale As Long
    maxDisplayScaleFactor = wzGetMaximumDisplayScaleFactorForWindowSize(width, height)
    maxDisplayScalePercentage = Int(maxDisplayScaleFactor * 100!)

    n = wzAvailableDisplayScales(availableDisplayScales)
    ' (std::sort: the list is sorted already)

    ' std::lower_bound
    maxDisplayScale = 0
    Do While maxDisplayScale < n
        If availableDisplayScales(maxDisplayScale) >= maxDisplayScalePercentage Then Exit Do
        maxDisplayScale = maxDisplayScale + 1
    Loop
    If maxDisplayScale = n Then
        ' return the largest available display scale
        wzGetMaximumDisplayScaleForWindowSize = availableDisplayScales(n - 1)
        Exit Function
    End If
    If availableDisplayScales(maxDisplayScale) <> maxDisplayScalePercentage Then
        If maxDisplayScale = 0 Then
            ' no lower display scale to return
            wzGetMaximumDisplayScaleForWindowSize = 0
            Exit Function
        End If
        maxDisplayScale = maxDisplayScale - 1
    End If
    wzGetMaximumDisplayScaleForWindowSize = availableDisplayScales(maxDisplayScale)
End Function

Public Function wzGetMaximumDisplayScaleForCurrentWindowSize() As Long
    wzGetMaximumDisplayScaleForCurrentWindowSize = wzGetMaximumDisplayScaleForWindowSize(windowWidth, windowHeight)
End Function

Public Function wzGetSuggestedDisplayScaleForCurrentWindowSize(ByVal desiredMaxScreenDimension As Long) As Long
    Dim maxDisplayScale As Long, availableDisplayScales() As Long, n As Long, it As Long
    Dim resultingLogicalScreenWidth As Long, resultingLogicalScreenHeight As Long
    maxDisplayScale = wzGetMaximumDisplayScaleForCurrentWindowSize()

    n = wzAvailableDisplayScales(availableDisplayScales)
    ' (std::sort: the list is sorted already)

    it = 0
    Do While it < n
        If Not (availableDisplayScales(it) <= maxDisplayScale) Then Exit Do
        resultingLogicalScreenWidth = (windowWidth * 100) \ availableDisplayScales(it)
        resultingLogicalScreenHeight = (windowHeight * 100) \ availableDisplayScales(it)
        If resultingLogicalScreenWidth <= desiredMaxScreenDimension And resultingLogicalScreenHeight <= desiredMaxScreenDimension Then
            wzGetSuggestedDisplayScaleForCurrentWindowSize = availableDisplayScales(it)
            Exit Function
        End If
        it = it + 1
    Loop

    wzGetSuggestedDisplayScaleForCurrentWindowSize = maxDisplayScale
End Function

Private Function wzWindowSizeIsSmallerThanMinimumRequired(ByVal width As Long, ByVal height As Long, ByVal displayScaleFactor As Single) As Boolean
    Dim minWindowWidth As Long, minWindowHeight As Long
    wzGetMinimumWindowSizeForDisplayScaleFactor minWindowWidth, minWindowHeight, displayScaleFactor
    wzWindowSizeIsSmallerThanMinimumRequired = ((width < minWindowWidth) Or (height < minWindowHeight))
End Function

' SDL_GetDisplayContentScale: the system's scale over 96 DPI as this process sees it (not DPI-aware: 1)
Public Function wzGetDisplayContentScale() As Single
    Dim defaultContentScale As Single
    defaultContentScale = win_contentScale()
    If defaultContentScale < 1! Then
        wzGetDisplayContentScale = 1!
        Exit Function
    End If
    wzGetDisplayContentScale = defaultContentScale
End Function

Private Function wzGetDefaultBaseDisplayScale(ByVal displayIndex As Long) As Long
    Dim defaultContentScale As Single, approxActualDisplayScale As Long, availableDisplayScales() As Long, n As Long, displayScale As Long
    defaultContentScale = win_contentScale()
    debug_ 1991, LOG_WZ, "wzGetDefaultBaseDisplayScale", "GetDisplayContentScale(%d)=%f", displayIndex, defaultContentScale
    If defaultContentScale <= 0! Then
        wzGetDefaultBaseDisplayScale = 100
        Exit Function
    End If
    ' This is an system-configured value that _could_ be applied *IN ADDITION TO* the user-configured display scaling value for WZ
    ' But for now, simply apply it as the default display scale
    approxActualDisplayScale = -Int(-(defaultContentScale * 100!))     ' ceilf
    If approxActualDisplayScale < 100 Then approxActualDisplayScale = 100
    n = wzAvailableDisplayScales(availableDisplayScales)
    displayScale = 0
    Do While displayScale < n
        If availableDisplayScales(displayScale) >= approxActualDisplayScale Then Exit Do
        displayScale = displayScale + 1
    Loop
    If displayScale = n Then
        ' return the largest available display scale
        wzGetDefaultBaseDisplayScale = availableDisplayScales(n - 1)
        Exit Function
    End If
    If availableDisplayScales(displayScale) <> approxActualDisplayScale Then
        If displayScale = 0 Then
            ' no lower display scale to return
            wzGetDefaultBaseDisplayScale = 100
            Exit Function
        End If
        displayScale = displayScale - 1
    End If
    wzGetDefaultBaseDisplayScale = availableDisplayScales(displayScale)
End Function

Public Function wzChangeDisplayScale(ByVal displayScale As Long) As Boolean
    Dim newDisplayScaleFactor As Single, minWindowWidth As Long, minWindowHeight As Long
    Dim oldScreenWidth As Long, oldScreenHeight As Long, newScreenWidth As Long, newScreenHeight As Long, pt As POINTAPI
    If WZwindow = 0 Then
        debug_ 2026, LOG_WARNING, "wzChangeDisplayScale", "wzChangeDisplayScale called when window is not available"
        Exit Function
    End If

    newDisplayScaleFactor = CSng(displayScale) / 100!

    If wzWindowSizeIsSmallerThanMinimumRequired(windowWidth, windowHeight, newDisplayScaleFactor) Then
        ' The current window width and/or height are below the required minimum window size
        ' for this display scale factor.
        Exit Function
    End If

    ' Store the new display scale factor
    setDisplayScale displayScale

    ' Set the new minimum window size
    If wzGetCurrentWindowMode() = WINDOW_MODE_windowed Then
        wzGetMinimumWindowSizeForDisplayScaleFactor minWindowWidth, minWindowHeight, newDisplayScaleFactor
        minTrackWidth = minWindowWidth          ' SDL_SetWindowMinimumSize
        minTrackHeight = minWindowHeight
    End If

    ' Update the game's logical screen size
    oldScreenWidth = screenWidth
    oldScreenHeight = screenHeight
    newScreenWidth = windowWidth
    newScreenHeight = windowHeight
    If newDisplayScaleFactor > 1! Then
        newScreenWidth = Fix(windowWidth / newDisplayScaleFactor)
        newScreenHeight = Fix(windowHeight / newDisplayScaleFactor)
    End If
    handleGameScreenSizeChange oldScreenWidth, oldScreenHeight, newScreenWidth, newScreenHeight
    game.gameDisplayScaleFactorDidChange newDisplayScaleFactor

    ' Update the current mouse coordinates
    ' (The prior stored mouseXPos / mouseYPos apply to the old coordinate system, and must be translated to the
    ' new game coordinate system. Since the mouse hasn't moved - or it would generate events that override this -
    ' the current position with respect to the window (which hasn't changed size) can be queried and used to
    ' calculate the new game coordinate system mouse position.)
    GetCursorPos pt                     ' SDL_GetMouseState: relative to the window
    ScreenToClient WZwindow, pt
    debug_ 2069, LOG_WZ, "wzChangeDisplayScale", "Old mouse position: %d, %d", mouseXPos, mouseYPos
    mouseXPos = Fix(CSng(pt.x) / current_displayScaleFactor)
    mouseYPos = Fix(CSng(pt.y) / current_displayScaleFactor)
    debug_ 2072, LOG_WZ, "wzChangeDisplayScale", "New mouse position: %d, %d", mouseXPos, mouseYPos

    processScreenSizeChangeNotificationIfNeeded

    wzChangeDisplayScale = True
End Function

Public Function wzChangeCursorScale(ByVal cursorScale As Long) As Boolean
    If WZwindow = 0 Then
        debug_ 2084, LOG_WARNING, "wzChangeCursorScale", "wzChangeCursorScale called when window is not available"
        Exit Function
    End If

    If cursorScale = game.warGetCursorScale() Then
        wzChangeCursorScale = True
        Exit Function
    End If

    game.warSetCursorScale cursorScale  ' must be set before cursors are reinit
    wzSDLReinitCursors
    wzChangeCursorScale = True
End Function

Private Function wzReduceDisplayScalingIfNeeded(ByVal currWidth As Long, ByVal currHeight As Long) As Boolean
    Dim maxDisplayScale As Long
    ' Check whether the desired window size is smaller than the minimum required for the current Display Scale
    If wzWindowSizeIsSmallerThanMinimumRequired(currWidth, currHeight, current_displayScaleFactor) Then
        ' The new window size is smaller than the minimum required size for the current display scale level.

        maxDisplayScale = wzGetMaximumDisplayScaleForWindowSize(currWidth, currHeight)
        If maxDisplayScale < 100 Then
            ' Cannot adjust display scale factor below 1. Desired window size is below the minimum supported.
            debug_ 2109, LOG_WZ, "wzReduceDisplayScalingIfNeeded", "Size (%d x %d) is smaller than the minimum supported at a 100%% display scale", currWidth, currHeight
            Exit Function
        End If

        ' Adjust the current display scale level to the nearest supported level.
        debug_ 2114, LOG_WZ, "wzReduceDisplayScalingIfNeeded", "The current Display Scale (%d%%) is too high for the desired window size. Reducing the current Display Scale to the maximum possible for the desired window size: %d%%.", current_displayScale, maxDisplayScale
        wzChangeDisplayScale maxDisplayScale

        ' Store the new display scale
        game.warSetDisplayScale maxDisplayScale
    End If

    wzReduceDisplayScalingIfNeeded = True
End Function

Public Function wzChangeFullscreenDisplayMode(ByVal has_config As Boolean, ByVal req_screen As Long, ByVal req_width As Long, ByVal req_height As Long, ByVal req_pixel_density As Single, ByVal req_refresh_rate As Single) As Boolean
    Dim closest As Long, dm As DEVMODEW, desktopDisplay As Long
    If WZwindow = 0 Then
        debug_ 2128, LOG_WARNING, "wzChangeFullscreenDisplayMode", "wzChangeFullscreenDisplayMode called when window is not available"
        Exit Function
    End If

    If Not has_config Then
        debug_ 2134, LOG_INFO, "wzChangeFullscreenDisplayMode", "Changing fullscreen mode to Desktop Full"
        If Not win_setWindowFullscreenMode(False, 0) Then
            debug_ 2137, LOG_INFO, "wzChangeFullscreenDisplayMode", "Unable to set desktop fullscreen mode?: %s", "(VB) the window could not take it"
        End If

        game.warSetFullscreenModeScreen 0
        game.warSetFullscreenModeWidth 0
        game.warSetFullscreenModeHeight 0
        game.warSetFullscreenModePixelDensity 1!
        game.warSetFullscreenModeRefreshRate 0!
        wzChangeFullscreenDisplayMode = True
        Exit Function
    End If

    debug_ 2149, LOG_INFO, "wzChangeFullscreenDisplayMode", "Changing fullscreen mode to [%u] %dx%d", req_screen, req_width, req_height

    If req_width = 0 Or req_height = 0 Then
        debug_ 2154, LOG_INFO, "wzChangeFullscreenDisplayMode", "Getting desktop display mode"
        If req_screen > 0 Then desktopDisplay = req_screen Else desktopDisplay = win_getDisplayForWindow()
        If Not win_getDesktopDisplayMode(desktopDisplay, dm) Then
            debug_ 2158, LOG_INFO, "wzChangeFullscreenDisplayMode", "Unable to get desktop DisplayMode?: %s", "(VB) no such display"
            Exit Function
        End If
        req_screen = desktopDisplay
        req_width = dm.dmPelsWidth
        req_height = dm.dmPelsHeight
        req_pixel_density = 1!
        req_refresh_rate = win_refreshRate(dm.dmDisplayFrequency)
    End If

    closest = win_getClosestFullscreenDisplayMode(req_screen, req_width, req_height, req_refresh_rate, (req_pixel_density > 1!))
    If closest < 0 Then
        ' no match was found
        debug_ 2171, LOG_INFO, "wzChangeFullscreenDisplayMode", "No closest DisplayMode found ([%u] [%u x %u]); error: %s", req_screen, req_width, req_height, "Couldn't find any matching video modes"
        Exit Function
    End If

    If (fsW(closest) < MIN_WZ_GAMESCREEN_WIDTH) Or (fsH(closest) < MIN_WZ_GAMESCREEN_HEIGHT) Then
        debug_ 2177, LOG_INFO, "wzChangeFullscreenDisplayMode", "Closest fullscreen mode minimum logical resolution < %d x %d", MIN_WZ_GAMESCREEN_WIDTH, MIN_WZ_GAMESCREEN_HEIGHT
        Exit Function
    End If

    If Not win_setWindowFullscreenMode(True, closest) Then
        ' SDL_SetWindowFullscreenMode failed
        debug_ 2184, LOG_INFO, "wzChangeFullscreenDisplayMode", "SDL_SetWindowFullscreenMode ([%u] [%u x %u]) failed: %s", req_screen, req_width, req_height, "(VB) the window could not take it"
        Exit Function
    End If

    ' Note: Reducing display scaling if needed is handled in the window resize event

    game.warSetFullscreenModeScreen fsDisplay(closest)
    game.warSetFullscreenModeWidth fsW(closest)
    game.warSetFullscreenModeHeight fsH(closest)
    game.warSetFullscreenModePixelDensity 1!
    game.warSetFullscreenModeRefreshRate fsRefresh(closest)

    wzChangeFullscreenDisplayMode = True
End Function

Private Function wzSetDefaultFullscreenDisplayMode() As Boolean
    Dim info_screen As Long, info_width As Long, info_height As Long, info_pixel_density As Single, info_refresh_rate As Single
    info_screen = game.warGetFullscreenModeScreen()
    info_width = game.warGetFullscreenModeWidth()
    info_height = game.warGetFullscreenModeHeight()
    info_pixel_density = game.warGetFullscreenModePixelDensity()
    info_refresh_rate = game.warGetFullscreenModeRefreshRate()

    If info_width <= 0 Or info_height <= 0 Then
        ' Desktop Full mode
        wzSetDefaultFullscreenDisplayMode = wzChangeFullscreenDisplayMode(False, 0, 0, 0, 0!, 0!)
        Exit Function
    End If

    wzSetDefaultFullscreenDisplayMode = wzChangeFullscreenDisplayMode(True, info_screen, info_width, info_height, info_pixel_density, info_refresh_rate)
End Function

Public Function wzGetCurrentMinimizeOnFocusLossBehavior() As Long
    Dim hint As String
    hint = minimizeOnFocusLossHint              ' SDL_GetHint(SDL_HINT_VIDEO_MINIMIZE_ON_FOCUS_LOSS)
    If Len(hint) = 0 Or StrComp(hint, "auto", vbTextCompare) = 0 Then
        wzGetCurrentMinimizeOnFocusLossBehavior = MinimizeOnFocusLossBehavior_Auto
        Exit Function
    End If
    If win_hintBoolean(hint) Then
        wzGetCurrentMinimizeOnFocusLossBehavior = MinimizeOnFocusLossBehavior_On_Fullscreen
    Else
        wzGetCurrentMinimizeOnFocusLossBehavior = MinimizeOnFocusLossBehavior_Off
    End If
End Function

Public Sub wzSetMinimizeOnFocusLoss(ByVal behavior As Long)
    Dim value As String
    value = "auto"
    Select Case behavior
    Case MinimizeOnFocusLossBehavior_Off
        value = "0"
    Case MinimizeOnFocusLossBehavior_On_Fullscreen
        value = "1"
    Case MinimizeOnFocusLossBehavior_Auto
        value = "auto"
    Case Else
        Exit Sub
    End Select
    minimizeOnFocusLossHint = value             ' SDL_SetHintWithPriority(..., SDL_HINT_OVERRIDE)
End Sub

' Returns the current window screen, width, and height
Public Sub wzGetWindowResolution(screen As Long, width As Long, height As Long)
    Dim rc As RECT
    If WZwindow = 0 Then
        ASSERT_FAILURE "WZwindow != nullptr", "lib/sdl/main_sdl.cpp:2375", "wzGetWindowResolution", "wzGetWindowResolution called when window is not available"
        Exit Sub
    End If

    screen = screenIndex

    GetClientRect WZwindow, rc          ' SDL_GetWindowSize
    width = rc.Right - rc.Left
    height = rc.Bottom - rc.Top
End Sub

Public Sub wzSetWindowIsResizable(ByVal resizable As Boolean)
    Dim minWindowWidth As Long, minWindowHeight As Long
    If WZwindow = 0 Then
        debug_ 3501, LOG_WARNING, "wzSetWindowIsResizable", "wzSetWindowIsResizable called when window is not available"
        Exit Sub
    End If
    win_setWindowResizable resizable    ' SDL_SetWindowResizable(WZwindow, resizable)

    If resizable Then
        ' Set the minimum window size
        wzGetMinimumWindowSizeForDisplayScaleFactor minWindowWidth, minWindowHeight, current_displayScaleFactor
        minTrackWidth = minWindowWidth          ' SDL_SetWindowMinimumSize
        minTrackHeight = minWindowHeight
    End If
End Sub

Public Function wzIsWindowResizable() As Boolean
    If WZwindow = 0 Then
        debug_ 3519, LOG_WARNING, "wzIsWindowResizable", "wzIsWindowResizable called when window is not available"
        Exit Function
    End If
    wzIsWindowResizable = windowResizable       ' SDL_GetWindowFlags(WZwindow) & SDL_WINDOW_RESIZABLE
End Function

' NOTE: wzBackendAttemptOpenURL should *not* be called directly - instead, call openURLInBrowser() from urlhelpers.h
Public Function wzBackendAttemptOpenURL(ByVal url As String) As Boolean
    Dim rc As Long
    ' SDL_OpenURL on Win32: ShellExecuteW "open"
    If testNoDesktopEffects Then
        debug_ 0, LOG_INFO, "wzBackendAttemptOpenURL", "(VB) TEST-ONLY: not opened: %s", url
        wzBackendAttemptOpenURL = True
        Exit Function
    End If
    rc = ShellExecuteW(0, StrPtr("open"), StrPtr(url), 0, 0, SW_SHOWNORMAL)
    wzBackendAttemptOpenURL = (rc > 32)
End Function

' Gets the system RAM in MiB (a uint64_t: a Double)
Public Function wzGetCurrentSystemRAM() As Double
    Dim value As Double, m As MEMORYSTATUSEX_
    ' SDL_GetSystemRAM on Win32: GlobalMemoryStatusEx's ullTotalPhys / (1024 * 1024)
    m.dwLength = LenB(m)
    If GlobalMemoryStatusEx(m) <> 0 Then value = Fix(CDbl(m.ullTotalPhys) * 10000# / 1048576#)
    If value > 0 Then wzGetCurrentSystemRAM = value
End Function

Public Function wzGetLogicalCPUCount() As Long
    Dim result As Long, si As SYSTEM_INFO
    GetSystemInfo si                    ' SDL_GetNumLogicalCPUCores on Win32
    result = si.dwNumberOfProcessors
    If result <= 0 Then
        debug_ 3830, LOG_ERROR, "wzGetLogicalCPUCount", "Failed to get logical CPU count - defaulting to 1"
        result = 1
    End If
    wzGetLogicalCPUCount = result
End Function

'--------------------------------------------------------------------------
' (VB) SDL's display and window primitives on Win32

' (VB) the game's side, for cursors_win
Public Function platformGame() As IPlatformGame
    Set platformGame = game
End Function

' TEST-ONLY: True when WZ_TEST_NO_DESKTOP_EFFECTS is set (see the state's comment)
Public Function wzTestNoDesktopEffects() As Boolean
    wzTestNoDesktopEffects = testNoDesktopEffects
End Function

' a WCHAR[] field's text
Private Function wideName(b() As Byte) As String
    Dim s As String, i As Long
    s = b
    i = InStr(s, vbNullChar)
    If i > 0 Then s = Left$(s, i - 1)
    wideName = s
End Function

' SDL_GetHintBoolean: "0" and "false" are false, any other value true
Private Function win_hintBoolean(ByVal hint As String) As Boolean
    win_hintBoolean = Not (hint = "0" Or StrComp(hint, "false", vbTextCompare) = 0)
End Function

' the system's scale over 96 DPI, as this process sees it
Private Function win_contentScale() As Single
    Dim hdc As Long
    hdc = GetDC(0)
    win_contentScale = CSng(GetDeviceCaps(hdc, LOGPIXELSX)) / 96!
    ReleaseDC 0, hdc
End Function

' (VB) SDL_Init's displays (WIN_InitModes: the attached adapters, the primary first, ids from 1) and each
' display's fullscreen modes (WIN_GetDisplayModes and SDL_AddFullscreenDisplayMode: palettized and unknown
' formats skipped, duplicates dropped, sorted as cmpmodes sorts)
Private Sub win_enumerateDisplays()
    Dim pass As Long, i As Long, dd As DISPLAY_DEVICEW, isPrimary As Boolean
    displayCount = 0
    ReDim displayDeviceName(0 To 0)
    fsModeCount = 0
    For pass = 0 To 1
        i = 0
        Do
            dd.cb = LenB(dd)
            If EnumDisplayDevicesW(0, i, dd, 0) = 0 Then Exit Do
            If (dd.StateFlags And DISPLAY_DEVICE_ATTACHED_TO_DESKTOP) <> 0 Then
                isPrimary = ((dd.StateFlags And DISPLAY_DEVICE_PRIMARY_DEVICE) <> 0)
                If isPrimary = (pass = 0) Then
                    displayCount = displayCount + 1
                    ReDim Preserve displayDeviceName(0 To displayCount)
                    displayDeviceName(displayCount) = wideName(dd.DeviceName)
                    win_addDisplayModes displayCount
                End If
            End If
            i = i + 1
        Loop
    Next
End Sub

' WIN_GetRefreshRate: NTSC timings for 119, 59 and 29 Hz
Private Function win_refreshRate(ByVal frequency As Long) As Single
    Select Case frequency
    Case 119, 59, 29
        win_refreshRate = ((100 * (frequency + 1) * 1000) \ 1001) / 100!
    Case Else
        win_refreshRate = CSng(frequency)
    End Select
End Function

' cmpmodes: < 0 when mode a sorts before mode b (wider, then taller, more bits per pixel, faster refresh first)
Private Function win_cmpmodes(ByVal aw As Long, ByVal ah As Long, ByVal abpp As Long, ByVal arefresh As Single, ByVal bw As Long, ByVal bh As Long, ByVal bbpp As Long, ByVal brefresh As Single) As Long
    If aw <> bw Then
        win_cmpmodes = bw - aw
    ElseIf ah <> bh Then
        win_cmpmodes = bh - ah
    ElseIf abpp <> bbpp Then
        win_cmpmodes = bbpp - abpp
    ElseIf Fix(arefresh * 100) <> Fix(brefresh * 100) Then
        win_cmpmodes = Fix(brefresh * 100) - Fix(arefresh * 100)
    End If
End Function

Private Sub win_addDisplayModes(ByVal id As Long)
    Dim i As Long, j As Long, k As Long, dm As DEVMODEW, w As Long, h As Long, bpp As Long, refresh As Single
    i = 0
    Do
        dm.dmSize = LenB(dm)
        If EnumDisplaySettingsExW(StrPtr(displayDeviceName(id)), i, dm, 0) = 0 Then Exit Do
        i = i + 1
        bpp = dm.dmBitsPerPel
        ' We don't support palettized modes now (8, 4); other bit counts have no SDL format
        If bpp <> 32 And bpp <> 24 And bpp <> 16 And bpp <> 15 Then GoTo nextMode
        w = dm.dmPelsWidth
        h = dm.dmPelsHeight
        refresh = win_refreshRate(dm.dmDisplayFrequency)
        ' Make sure we don't already have the mode in the list
        For j = 0 To fsModeCount - 1
            If fsDisplay(j) = id Then
                If win_cmpmodes(w, h, bpp, refresh, fsW(j), fsH(j), fsBpp(j), fsRefresh(j)) = 0 Then GoTo nextMode
            End If
        Next
        ' sorted in: before the first of this display's modes it sorts before
        k = fsModeCount
        For j = 0 To fsModeCount - 1
            If fsDisplay(j) = id Then
                If win_cmpmodes(w, h, bpp, refresh, fsW(j), fsH(j), fsBpp(j), fsRefresh(j)) < 0 Then
                    k = j
                    Exit For
                End If
            End If
        Next
        ReDim Preserve fsDisplay(0 To fsModeCount)
        ReDim Preserve fsW(0 To fsModeCount)
        ReDim Preserve fsH(0 To fsModeCount)
        ReDim Preserve fsBpp(0 To fsModeCount)
        ReDim Preserve fsRefresh(0 To fsModeCount)
        ReDim Preserve fsFrequency(0 To fsModeCount)
        ReDim Preserve fsDisplayFlags(0 To fsModeCount)
        For j = fsModeCount To k + 1 Step -1
            fsDisplay(j) = fsDisplay(j - 1)
            fsW(j) = fsW(j - 1)
            fsH(j) = fsH(j - 1)
            fsBpp(j) = fsBpp(j - 1)
            fsRefresh(j) = fsRefresh(j - 1)
            fsFrequency(j) = fsFrequency(j - 1)
            fsDisplayFlags(j) = fsDisplayFlags(j - 1)
        Next
        fsDisplay(k) = id
        fsW(k) = w
        fsH(k) = h
        fsBpp(k) = bpp
        fsRefresh(k) = refresh
        fsFrequency(k) = dm.dmDisplayFrequency
        fsDisplayFlags(k) = dm.dmDisplayFlags
        fsModeCount = fsModeCount + 1
nextMode:
    Loop
End Sub

' SDL_GetPixelFormatName of the mode's format (WIN_GetDisplayMode's format for its bits per pixel)
Private Function win_pixelFormatName(ByVal bpp As Long) As String
    Select Case bpp
    Case 32: win_pixelFormatName = "SDL_PIXELFORMAT_XRGB8888"
    Case 24: win_pixelFormatName = "SDL_PIXELFORMAT_RGB24"
    Case 16: win_pixelFormatName = "SDL_PIXELFORMAT_RGB565"
    Case 15: win_pixelFormatName = "SDL_PIXELFORMAT_XRGB1555"
    Case Else: win_pixelFormatName = "SDL_PIXELFORMAT_UNKNOWN"
    End Select
End Function

' main_sdl.cpp's displaylist.push_back
Private Sub win_displaylistAdd(ByVal has As Boolean, ByVal screen As Long, ByVal width As Long, ByVal height As Long, ByVal pixel_density As Single, ByVal refresh_rate As Single)
    ReDim Preserve dlHas(0 To displaylistCount)
    ReDim Preserve dlScreen(0 To displaylistCount)
    ReDim Preserve dlWidth(0 To displaylistCount)
    ReDim Preserve dlHeight(0 To displaylistCount)
    ReDim Preserve dlPixelDensity(0 To displaylistCount)
    ReDim Preserve dlRefreshRate(0 To displaylistCount)
    dlHas(displaylistCount) = has
    dlScreen(displaylistCount) = screen
    dlWidth(displaylistCount) = width
    dlHeight(displaylistCount) = height
    dlPixelDensity(displaylistCount) = pixel_density
    dlRefreshRate(displaylistCount) = refresh_rate
    displaylistCount = displaylistCount + 1
End Sub

' SDL_GetDisplayBounds: the display's current mode's position and size
Private Function win_getDisplayBounds(ByVal id As Long, rc As RECT) As Boolean
    Dim dm As DEVMODEW
    If id < 1 Or id > displayCount Then Exit Function
    dm.dmSize = LenB(dm)
    If EnumDisplaySettingsExW(StrPtr(displayDeviceName(id)), ENUM_CURRENT_SETTINGS, dm, 0) = 0 Then Exit Function
    rc.Left = dm.dmPositionX
    rc.Top = dm.dmPositionY
    rc.Right = rc.Left + dm.dmPelsWidth
    rc.Bottom = rc.Top + dm.dmPelsHeight
    win_getDisplayBounds = True
End Function

' SDL_GetDisplayUsableBounds: the work area of the display's monitor
Private Function win_getDisplayUsableBounds(ByVal id As Long, rc As RECT) As Boolean
    Dim b As RECT, mi As MONITORINFOEXW, hmon As Long
    If Not win_getDisplayBounds(id, b) Then Exit Function
    hmon = MonitorFromPoint((b.Left + b.Right) \ 2, (b.Top + b.Bottom) \ 2, MONITOR_DEFAULTTONEAREST)
    mi.cbSize = LenB(mi)
    If GetMonitorInfoW(hmon, mi) = 0 Then Exit Function
    rc = mi.rcWork
    win_getDisplayUsableBounds = True
End Function

' SDL_GetDesktopDisplayMode: the display's mode in the registry (the desktop's, while a fullscreen mode is set)
Private Function win_getDesktopDisplayMode(ByVal id As Long, dm As DEVMODEW) As Boolean
    If id < 1 Or id > displayCount Then Exit Function
    dm.dmSize = LenB(dm)
    win_getDesktopDisplayMode = (EnumDisplaySettingsExW(StrPtr(displayDeviceName(id)), ENUM_REGISTRY_SETTINGS, dm, 0) <> 0)
End Function

' SDL_GetDisplayForWindow (0: none)
Private Function win_getDisplayForWindow() As Long
    Dim mi As MONITORINFOEXW, hmon As Long, nm As String, i As Long
    If WZwindow = 0 Then Exit Function
    hmon = MonitorFromWindow(WZwindow, MONITOR_DEFAULTTONEAREST)
    mi.cbSize = LenB(mi)
    If GetMonitorInfoW(hmon, mi) = 0 Then Exit Function
    nm = wideName(mi.szDevice)
    For i = 1 To displayCount
        If StrComp(displayDeviceName(i), nm, vbTextCompare) = 0 Then
            win_getDisplayForWindow = i
            Exit Function
        End If
    Next
End Function

' a display id, or the primary for 0 or an unknown id (SDL_WINDOWPOS_CENTERED_DISPLAY(0))
Private Function win_displayOrPrimary(ByVal id As Long) As Long
    If id < 1 Or id > displayCount Then id = 1
    win_displayOrPrimary = id
End Function

' SDL_GetClosestFullscreenDisplayMode: the index of the closest of the display's fullscreen modes (-1: none)
Private Function win_getClosestFullscreenDisplayMode(ByVal displayID As Long, ByVal w As Long, ByVal h As Long, ByVal refresh_rate As Single, ByVal include_high_density_modes As Boolean) As Long
    Dim i As Long, closest As Long, aspect_ratio As Single, current_aspect_ratio As Single, closest_aspect_ratio As Single, dm As DEVMODEW
    closest = -1
    If displayID < 1 Or displayID > displayCount Then
        win_getClosestFullscreenDisplayMode = -1
        Exit Function
    End If
    If h > 0 Then
        aspect_ratio = CSng(w) / CSng(h)
    Else
        aspect_ratio = 1!
    End If
    If refresh_rate <= 0! Then
        If win_getDesktopDisplayMode(displayID, dm) Then refresh_rate = win_refreshRate(dm.dmDisplayFrequency)
    End If
    ' (the modes' pixel density is 1 here: include_high_density_modes changes nothing)
    For i = 0 To fsModeCount - 1
        If fsDisplay(i) <> displayID Then GoTo nextMode
        If fsW(i) < w Then
            ' Out of sorted modes large enough here
            Exit For
        End If
        If fsH(i) < h Then
            ' Wider, but not tall enough, due to a different aspect ratio.
            GoTo nextMode
        End If
        If closest >= 0 Then
            current_aspect_ratio = CSng(fsW(i)) / CSng(fsH(i))
            closest_aspect_ratio = CSng(fsW(closest)) / CSng(fsH(closest))
            If Abs(aspect_ratio - closest_aspect_ratio) < Abs(aspect_ratio - current_aspect_ratio) Then
                ' The mode we already found has a better aspect ratio match
                GoTo nextMode
            End If
            If fsW(i) = fsW(closest) And fsH(i) = fsH(closest) And Abs(fsRefresh(closest) - refresh_rate) < Abs(fsRefresh(i) - refresh_rate) Then
                ' We already found a mode and the new mode is further from our refresh rate target
                GoTo nextMode
            End If
        End If
        closest = i
nextMode:
    Next
    win_getClosestFullscreenDisplayMode = closest
End Function

' SDL_SetWindowFullscreenMode: the window's fullscreen mode (hasMode False: borderless, "Desktop Full"),
' applied at once when the window is fullscreen
Private Function win_setWindowFullscreenMode(ByVal hasMode As Boolean, ByVal modeIndex As Long) As Boolean
    hasFullscreenMode = hasMode
    fullscreenModeIndex = modeIndex
    If windowIsFullscreen Then
        win_setWindowFullscreenMode = win_setWindowFullscreen(True)
        Exit Function
    End If
    win_setWindowFullscreenMode = True
End Function

' (SDL_SetDisplayModeForDisplay, WIN_SetDisplayMode) the display changed to the window's fullscreen mode
Private Function win_applyFullscreenDisplayMode() As Boolean
    Dim dm As DEVMODEW, cur As DEVMODEW, id As Long, r As Long
    id = fsDisplay(fullscreenModeIndex)
    ' the display shows this mode already: nothing to change
    cur.dmSize = LenB(cur)
    If EnumDisplaySettingsExW(StrPtr(displayDeviceName(id)), ENUM_CURRENT_SETTINGS, cur, 0) <> 0 Then
        If cur.dmPelsWidth = fsW(fullscreenModeIndex) And cur.dmPelsHeight = fsH(fullscreenModeIndex) And cur.dmBitsPerPel = fsBpp(fullscreenModeIndex) _
            And cur.dmDisplayFrequency = fsFrequency(fullscreenModeIndex) Then
            win_applyFullscreenDisplayMode = True
            Exit Function
        End If
    End If
    If testNoDesktopEffects Then
        debug_ 0, LOG_INFO, "win_applyFullscreenDisplayMode", "(VB) TEST-ONLY: ChangeDisplaySettingsEx not called: display %d to %dx%d %d bpp @%d Hz", id, fsW(fullscreenModeIndex), fsH(fullscreenModeIndex), fsBpp(fullscreenModeIndex), fsFrequency(fullscreenModeIndex)
        Exit Function
    End If
    dm.dmSize = LenB(dm)
    dm.dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_DISPLAYFREQUENCY Or DM_DISPLAYFLAGS
    dm.dmBitsPerPel = fsBpp(fullscreenModeIndex)
    dm.dmPelsWidth = fsW(fullscreenModeIndex)
    dm.dmPelsHeight = fsH(fullscreenModeIndex)
    dm.dmDisplayFrequency = fsFrequency(fullscreenModeIndex)
    dm.dmDisplayFlags = fsDisplayFlags(fullscreenModeIndex)
    r = ChangeDisplaySettingsExW(StrPtr(displayDeviceName(id)), dm, 0, CDS_FULLSCREEN, 0)
    If r <> DISP_CHANGE_SUCCESSFUL Then
        debug_ 0, LOG_ERROR, "win_applyFullscreenDisplayMode", "(VB) ChangeDisplaySettingsEx failed: %d", r
        Exit Function
    End If
    exclusiveModeApplied = True
    exclusiveModeDisplay = id
    win_applyFullscreenDisplayMode = True
End Function

' (VB) the display's desktop mode back
Private Sub win_restoreDesktopDisplayMode()
    If Not exclusiveModeApplied Then Exit Sub
    ChangeDisplaySettingsExW StrPtr(displayDeviceName(exclusiveModeDisplay)), ByVal 0&, 0, CDS_FULLSCREEN, 0
    exclusiveModeApplied = False
End Sub

' (VB) GetWindowStyle and WIN_SetWindowResizable: the window's style for its state
Private Sub win_applyStyle(ByVal newStyle As Long)
    Dim style As Long
    style = GetWindowLongW(WZwindow, GWL_STYLE)
    style = style And Not STYLE_MASK
    style = style Or newStyle
    windowStyle = newStyle
    SetWindowLongW WZwindow, GWL_STYLE, style
End Sub

' SDL_SetWindowResizable
Private Sub win_setWindowResizable(ByVal resizable As Boolean)
    windowResizable = resizable
    If windowIsFullscreen Then Exit Sub         ' (a fullscreen window's style has no frame)
    If resizable Then win_applyStyle STYLE_NORMAL Or STYLE_RESIZABLE Else win_applyStyle STYLE_NORMAL
End Sub

' SDL_SetWindowFullscreen on Win32. Into fullscreen: the window's fullscreen mode set on its display (as SDL's
' WIN_SetDisplayMode does) or, with none, the display left as it is (borderless); a popup window over the
' display's bounds. Out of it: the desktop mode back, the windowed style and rect.
Private Function win_setWindowFullscreen(ByVal fullscreen As Boolean) As Boolean
    Dim id As Long, b As RECT
    If fullscreen Then
        If hasFullscreenMode Then
            id = fsDisplay(fullscreenModeIndex)
        Else
            id = win_displayOrPrimary(win_getDisplayForWindow())
        End If
        If testNoDesktopEffects And (hasFullscreenMode Or Not testAllowBorderless) Then
            If hasFullscreenMode Then
                debug_ 0, LOG_INFO, "win_setWindowFullscreen", "(VB) TEST-ONLY: not entering fullscreen: display %d at %dx%d @%d Hz", id, fsW(fullscreenModeIndex), fsH(fullscreenModeIndex), fsFrequency(fullscreenModeIndex)
            Else
                debug_ 0, LOG_INFO, "win_setWindowFullscreen", "(VB) TEST-ONLY: not entering fullscreen: display %d, Desktop Full", id
            End If
            Exit Function
        End If
        If Not windowIsFullscreen Then GetWindowRect WZwindow, windowedRestoreRect
        If hasFullscreenMode Then
            If Not win_applyFullscreenDisplayMode() Then Exit Function
        ElseIf exclusiveModeApplied Then
            win_restoreDesktopDisplayMode
        End If
        If Not win_getDisplayBounds(id, b) Then Exit Function
        windowIsFullscreen = True
        win_applyStyle STYLE_FULLSCREEN
        SetWindowPos WZwindow, HWND_TOP, b.Left, b.Top, b.Right - b.Left, b.Bottom - b.Top, SWP_FRAMECHANGED Or SWP_NOCOPYBITS Or SWP_NOACTIVATE
    Else
        If Not windowIsFullscreen Then
            win_setWindowFullscreen = True
            Exit Function
        End If
        win_restoreDesktopDisplayMode
        windowIsFullscreen = False
        If windowResizable Then win_applyStyle STYLE_NORMAL Or STYLE_RESIZABLE Else win_applyStyle STYLE_NORMAL
        SetWindowPos WZwindow, HWND_TOP, windowedRestoreRect.Left, windowedRestoreRect.Top, windowedRestoreRect.Right - windowedRestoreRect.Left, _
            windowedRestoreRect.Bottom - windowedRestoreRect.Top, SWP_FRAMECHANGED Or SWP_NOCOPYBITS Or SWP_NOACTIVATE
    End If
    win_updateClipCursor
    win_setWindowFullscreen = True
End Function

' SDL_SetWindowSize: the client area's size
Private Sub win_setWindowSize(ByVal w As Long, ByVal h As Long)
    Dim rc As RECT
    rc.Right = w
    rc.Bottom = h
    AdjustWindowRectEx rc, windowStyle, 0, 0
    SetWindowPos WZwindow, 0, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE
End Sub

' SDL_SetWindowPosition(SDL_WINDOWPOS_CENTERED_DISPLAY(id), ...): centred in the display's usable area
Private Sub win_centerWindowOnDisplay(ByVal id As Long)
    Dim wr As RECT, b As RECT, x As Long, y As Long
    If Not win_getDisplayUsableBounds(win_displayOrPrimary(id), b) Then Exit Sub
    GetWindowRect WZwindow, wr
    x = b.Left + ((b.Right - b.Left) - (wr.Right - wr.Left)) \ 2
    y = b.Top + ((b.Bottom - b.Top) - (wr.Bottom - wr.Top)) \ 2
    If x < b.Left Then x = b.Left
    If y < b.Top Then y = b.Top
    SetWindowPos WZwindow, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
End Sub

' SDL_ShouldMinimizeOnFocusLoss: a fullscreen window with a fullscreen mode (exclusive) minimizes, unless the hint
' says otherwise ("Real fullscreen windows should minimize on focus loss so the desktop video mode is restored")
Private Function win_shouldMinimizeOnFocusLoss() As Boolean
    Dim hint As String
    If Not windowIsFullscreen Then Exit Function
    hint = minimizeOnFocusLossHint
    If Len(hint) = 0 Or StrComp(hint, "auto", vbTextCompare) = 0 Then
        win_shouldMinimizeOnFocusLoss = hasFullscreenMode
        Exit Function
    End If
    win_shouldMinimizeOnFocusLoss = win_hintBoolean(hint)
End Function

' WIN_UpdateClipCursor: the cursor kept in the client area while the mouse is grabbed and the window has the
' focus and is not minimized
Private Sub win_updateClipCursor()
    Dim rc As RECT, pt As POINTAPI, wantClip As Boolean
    wantClip = mouseGrabbed And windowHasFocus And WZwindow <> 0
    If wantClip Then
        If IsIconic(WZwindow) <> 0 Then wantClip = False
    End If
    If wantClip Then
        GetClientRect WZwindow, rc
        pt.x = 0
        pt.y = 0
        ClientToScreen WZwindow, pt
        rc.Left = pt.x
        rc.Top = pt.y
        rc.Right = rc.Right + pt.x
        rc.Bottom = rc.Bottom + pt.y
        If testNoDesktopEffects Then
            If Not cursorClipped Then debug_ 0, LOG_INFO, "win_updateClipCursor", "(VB) TEST-ONLY: ClipCursor not called: (%d, %d) - (%d, %d)", rc.Left, rc.Top, rc.Right, rc.Bottom
        Else
            ClipCursor rc
        End If
        cursorClipped = True
    ElseIf cursorClipped Then
        If testNoDesktopEffects Then
            debug_ 0, LOG_INFO, "win_updateClipCursor", "(VB) TEST-ONLY: ClipCursor(NULL) not called"
        Else
            ClipCursor ByVal 0&
        End If
        cursorClipped = False
    End If
End Sub

' SDL_SetCursor (WIN_ShowCursor): the cursor the client area shows (0: the default arrow)
Public Sub win_setCursor(ByVal hCursor As Long)
    SDL_cursor = hCursor
    If WZwindow <> 0 And mouseInWindow Then
        If hCursor <> 0 Then SetCursor hCursor Else SetCursor LoadCursorW(0, IDC_ARROW)
    End If
End Sub

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
    ' (VB) SDL reads a hint's default from the environment variable of its name
    minimizeOnFocusLossHint = Environ$("SDL_VIDEO_MINIMIZE_ON_FOCUS_LOSS")
    ' TEST-ONLY
    testNoDesktopEffects = (Len(Environ$("WZ_TEST_NO_DESKTOP_EFFECTS")) > 0)
    testAllowBorderless = (Environ$("WZ_TEST_NO_DESKTOP_EFFECTS") = "borderless")
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

' C: wzMainScreenSetup_CreateVideoWindow. Returns False on failure.
Private Function wzMainScreenSetup_CreateVideoWindow(ByVal backend As Long, ByVal antialiasing As Long, ByVal fullscreen As Long, ByVal highDPI As Boolean) As Boolean
    Dim desiredWindowWidth As Long, desiredWindowHeight As Long
    Dim minWindowWidth As Long, minWindowHeight As Long
    Dim displayUsableBounds As RECT, bounds As RECT, rc As RECT, x As Long, y As Long, w As Long, h As Long
    Dim resultingWidth As Long, resultingHeight As Long, i As Long
    Dim screenBaseDisplayScale As Long, priorDisplayScale As Long, displayScaleDiff As Single, maxDisplayScale As Long

    ' ensure "fullscreen" is in the supported fullscreen modes for the current system
    If Not wzIsSupportedWindowMode(fullscreen) Then
        debug_ 2668, LOG_ERROR, "wzMainScreenSetup_CreateVideoWindow", "Unsupported fullscreen mode specified: %d; using default", fullscreen
        fullscreen = WINDOW_MODE_windowed
        game.warSetWindowMode fullscreen ' persist the change
    End If

    ' (VB) SDL_Init's displays and their fullscreen modes
    win_enumerateDisplays

    displaylistCount = 0

    ' Add the "Desktop Full" (borderless fullscreen) mode to the displaylist
    win_displaylistAdd False, 0, 0, 0, 0!, 0!

    ' Populate our resolution list (does all displays now)
    For i = 0 To fsModeCount - 1
        debug_ 2713, LOG_WZ, "wzMainScreenSetup_CreateVideoWindow", "Monitor [%d] %dx%d %f @%f %s", fsDisplay(i), fsW(i), fsH(i), fsRefresh(i), 1!, win_pixelFormatName(fsBpp(i))

        If (fsW(i) < MIN_WZ_GAMESCREEN_WIDTH) Or (fsH(i) < MIN_WZ_GAMESCREEN_HEIGHT) Then
            debug_ 2719, LOG_WZ, "wzMainScreenSetup_CreateVideoWindow", "Monitor mode logical resolution < %d x %d -- discarding entry", MIN_WZ_GAMESCREEN_WIDTH, MIN_WZ_GAMESCREEN_HEIGHT
        ElseIf fsRefresh(i) < 59! Then
            debug_ 2723, LOG_WZ, "wzMainScreenSetup_CreateVideoWindow", "Monitor mode refresh rate < 59 -- discarding entry"
            ' only store 60Hz & higher modes, some display report 59 on Linux
        ElseIf fsBpp(i) < 24 Then           ' the 16 and 15 bits per pixel formats: R, G or B bits < 8
            debug_ 2734, LOG_INFO, "wzMainScreenSetup_CreateVideoWindow", "Monitor mode pixel format has R, G, or B bits < 8 -- discarding entry"
        Else
            win_displaylistAdd True, fsDisplay(i), fsW(i), fsH(i), 1!, fsRefresh(i)    ' which display this belongs to

            debug_ 2745, LOG_WZ, "wzMainScreenSetup_CreateVideoWindow", "Monitor [%d] %dx%d %f @%f %s", fsDisplay(i), fsW(i), fsH(i), fsRefresh(i), 1!, win_pixelFormatName(fsBpp(i))
        End If
    Next

    ' populate with the saved configuration values (if we had any) for windowed mode
    desiredWindowWidth = game.warGetWidth()
    desiredWindowHeight = game.warGetHeight()

    ' NOTE: Prior to wzMainScreenSetup being run, the display system is populated with the window width + height
    ' (i.e. not taking into account the game display scale). This function later sets the display system
    ' to the *game screen* width and height (taking into account the display scale).

    If desiredWindowWidth = 0 Or desiredWindowHeight = 0 Then
        windowWidth = 1024
        windowHeight = 768
    Else
        windowWidth = desiredWindowWidth
        windowHeight = desiredWindowHeight
    End If

    setDisplayScale game.warGetDisplayScale()

    For i = 1 To displayCount
        win_getDisplayBounds i, bounds
        debug_ 2777, LOG_WZ, "wzMainScreenSetup_CreateVideoWindow", "Monitor %d: pos %d x %d : res %d x %d", i, bounds.Left, bounds.Top, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top
    Next
    If displayCount < 1 Then
        debug_ 2781, LOG_FATAL, "wzMainScreenSetup_CreateVideoWindow", "SDL_GetDisplays returned: %d, with error: %s", displayCount, "(VB) EnumDisplayDevices found no display attached to the desktop"
        exit_ 1                         ' EXIT_FAILURE
    End If

    ' Check desired screen index values versus current system
    If game.warGetScreen() > displayCount Then
        debug_ 2790, LOG_WARNING, "wzMainScreenSetup_CreateVideoWindow", "Invalid screen [%d] defined in configuration; there are only %d displays; falling back to display 0", game.warGetScreen(), displayCount
        game.warSetScreen 0
    End If
    If game.warGetFullscreenModeScreen() > displayCount Or game.warGetFullscreenModeScreen() < 0 Then
        debug_ 2795, LOG_WARNING, "wzMainScreenSetup_CreateVideoWindow", "Invalid fullscreen screen [%d] defined in configuration; there are only %d displays; falling back to display 0", game.warGetFullscreenModeScreen(), displayCount
        game.warSetFullscreenModeScreen 0
    End If

    screenIndex = game.warGetScreen()

    ' Calculate the minimum window size (in *logical* points) given the current display scale
    wzGetMinimumWindowSizeForDisplayScaleFactor minWindowWidth, minWindowHeight, current_displayScaleFactor

    If fullscreen <> WINDOW_MODE_fullscreen Then
        If game.warGetAutoAdjustDisplayScale() Then
            ' Since SDL (at least <= 2.0.14) does not provide built-in support for high-DPI displays on *some* platforms
            ' (example: Windows), do our best to bump up the game's Display Scale setting to be a better match if the screen
            ' on which the game starts has a higher effective Display Scale than the game's starting Display Scale setting.
            If screenIndex > 0 Then
                screenBaseDisplayScale = wzGetDefaultBaseDisplayScale(screenIndex)
            Else
                screenBaseDisplayScale = wzGetDefaultBaseDisplayScale(1)    ' SDL_GetPrimaryDisplay()
            End If

            If screenBaseDisplayScale > wzGetCurrentDisplayScale() Then
                ' When bumping up the display scale, also increase the target window size proportionally
                debug_ 2817, LOG_WZ, "wzMainScreenSetup_CreateVideoWindow", "Increasing game Display Scale to better match calculated default base display scale: %u%% -> %u%%", wzGetCurrentDisplayScale(), screenBaseDisplayScale
                priorDisplayScale = wzGetCurrentDisplayScale()
                setDisplayScale screenBaseDisplayScale
                displayScaleDiff = CSng(screenBaseDisplayScale) / CSng(priorDisplayScale)
                windowWidth = -Int(-(CSng(windowWidth) * displayScaleDiff))     ' ceil
                windowHeight = -Int(-(CSng(windowHeight) * displayScaleDiff))
            End If
        End If

        If (windowWidth < minWindowWidth) Or (windowHeight < minWindowHeight) Then
            ' The desired window width and/or height is lower than the required minimum for the current display scale.
            ' Reduce the display scale to the maximum supported (for the desired window size), and recalculate the required minimum window size.
            maxDisplayScale = wzGetMaximumDisplayScaleForWindowSize(windowWidth, windowHeight)
            If maxDisplayScale < 100 Then maxDisplayScale = 100     ' if wzGetMaximumDisplayScaleForWindowSize fails, it returns < 100
            setDisplayScale maxDisplayScale
            wzGetMinimumWindowSizeForDisplayScaleFactor minWindowWidth, minWindowHeight, current_displayScaleFactor
        End If

        If windowWidth < minWindowWidth Then windowWidth = minWindowWidth
        If windowHeight < minWindowHeight Then windowHeight = minWindowHeight

        If fullscreen = WINDOW_MODE_windowed Then
            ' Determine the maximum usable windowed size for this display/screen
            ' (4.7.0 tests SDL3's bool result against 0: the cap runs only when the call fails, with empty bounds,
            ' so it never caps; the port does the same)
            If win_getDisplayUsableBounds(screenIndex, displayUsableBounds) = False Then
                w = displayUsableBounds.Right - displayUsableBounds.Left
                h = displayUsableBounds.Bottom - displayUsableBounds.Top
                If w > 0 And h > 0 Then
                    If windowWidth > w Then windowWidth = w
                    If windowHeight > h Then windowHeight = h
                End If
            End If
        End If
    End If

    ' SDL_CreateWindowWithProperties: the client area windowWidth x windowHeight, centred on the display
    ' Allow the window to be manually resized, if not fullscreen
    windowResizable = (fullscreen = WINDOW_MODE_windowed)
    If windowResizable Then windowStyle = STYLE_NORMAL Or STYLE_RESIZABLE Else windowStyle = STYLE_NORMAL
    rc.Right = windowWidth
    rc.Bottom = windowHeight
    AdjustWindowRectEx rc, windowStyle, 0, 0
    w = rc.Right - rc.Left
    h = rc.Bottom - rc.Top
    win_getDisplayUsableBounds win_displayOrPrimary(screenIndex), displayUsableBounds
    x = displayUsableBounds.Left + ((displayUsableBounds.Right - displayUsableBounds.Left) - w) \ 2
    y = displayUsableBounds.Top + ((displayUsableBounds.Bottom - displayUsableBounds.Top) - h) \ 2
    If x < displayUsableBounds.Left Then x = displayUsableBounds.Left
    If y < displayUsableBounds.Top Then y = displayUsableBounds.Top

    minTrackWidth = minWindowWidth
    minTrackHeight = minWindowHeight
    lastClientWidth = windowWidth
    lastClientHeight = windowHeight
    lastSizeType = SIZE_RESTORED
    WZwindow = CreateWindowExW(0, StrPtr(className), StrPtr(PACKAGE_NAME), windowStyle, x, y, w, h, 0, 0, GetModuleHandleW(0), 0)
    If WZwindow = 0 Then
        debug_ 2879, LOG_INFO, "wzMainScreenSetup_CreateVideoWindow", "Failed to create a window, because: %s", "CreateWindowExW failed (" & Err.LastDllError & ")"
        Exit Function
    End If
    queueEnabled = True
    ShowWindow WZwindow, SW_SHOW
    SetForegroundWindow WZwindow
    If testNoDesktopEffects Then
        debug_ 0, LOG_INFO, "wzMainScreenSetup_CreateVideoWindow", "(VB) TEST-ONLY: WZ_TEST_NO_DESKTOP_EFFECTS is set: display mode changes, fullscreen, the cursor clip, the clipboard and links are logged, not made"
        If testAllowBorderless Then debug_ 0, LOG_INFO, "wzMainScreenSetup_CreateVideoWindow", "(VB) TEST-ONLY: borderless: a borderless fullscreen window is made all the same"
    End If

    ' Always set the fullscreen mode (so switching works to the desired mode, even if we don't start in fullscreen mode)
    If wzSetDefaultFullscreenDisplayMode() Then
        If fullscreen = WINDOW_MODE_fullscreen Then
            ' queue a later transition to fullscreen mode
            pendingFullscreenTransition = True
        End If
    Else
        If fullscreen = WINDOW_MODE_fullscreen Then
            debug_ 2899, LOG_ERROR, "wzMainScreenSetup_CreateVideoWindow", "Failed to initialize at fullscreen mode ([%d] [%u x %u]); reverting to windowed mode", screenIndex, game.warGetFullscreenModeWidth(), game.warGetFullscreenModeHeight()
            wzChangeWindowMode WINDOW_MODE_windowed
            fullscreen = WINDOW_MODE_windowed
        End If
    End If

    ' Get resulting logical window size
    GetClientRect WZwindow, rc
    resultingWidth = rc.Right - rc.Left
    resultingHeight = rc.Bottom - rc.Top

    ' Check that the actual window size matches the desired window size (but not for classic fullscreen mode)
    If (wzGetCurrentWindowMode() <> WINDOW_MODE_fullscreen) And (resultingWidth < minWindowWidth Or resultingHeight < minWindowHeight) Then
        ' The created window size (that the system returned) is less than the minimum required window size (for this display scale)
        debug_ 2913, LOG_WARNING, "wzMainScreenSetup_CreateVideoWindow", "Failed to create window at desired resolution: [%d] %d x %d; instead, received window of resolution: [%d] %d x %d; which is below the required minimum size of %d x %d for the current display scale level", game.warGetScreen(), windowWidth, windowHeight, game.warGetScreen(), resultingWidth, resultingHeight, minWindowWidth, minWindowHeight

        ' Adjust the display scale (if possible)
        maxDisplayScale = wzGetMaximumDisplayScaleForWindowSize(resultingWidth, resultingHeight)
        If maxDisplayScale >= 100 Then
            ' Reduce the display scale
            debug_ 2920, LOG_WARNING, "wzMainScreenSetup_CreateVideoWindow", "Reducing the display scale level to the maximum supported for this window size: %u", maxDisplayScale
            setDisplayScale maxDisplayScale
            wzGetMinimumWindowSizeForDisplayScaleFactor minWindowWidth, minWindowHeight, current_displayScaleFactor
        Else
            ' Problem: The resulting window size that the system provided is below the minimum required
            '          for the lowest possible display scale - i.e. the window is just too small
            debug_ 2928, LOG_ERROR, "wzMainScreenSetup_CreateVideoWindow", "The window size created by the system is too small!"

            ' Attempt to default to base resolution at 100%
            setDisplayScale 100
            wzGetMinimumWindowSizeForDisplayScaleFactor minWindowWidth, minWindowHeight, current_displayScaleFactor

            win_setWindowSize minWindowWidth, minWindowHeight
            windowWidth = minWindowWidth
            windowHeight = minWindowHeight

            ' Center window on screen
            win_centerWindowOnDisplay screenIndex

            ' Re-request resulting window size
            GetClientRect WZwindow, rc
            resultingWidth = rc.Right - rc.Left
            resultingHeight = rc.Bottom - rc.Top
        End If
    End If
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

    ' Set the minimum window size
    minTrackWidth = minWindowWidth
    minTrackHeight = minWindowHeight

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
    Dim swapMode As Long, lodDistanceBias As Single, minOnFocusLossSettingVal As Long, hint As String

    If Not wzSDLOneTimeInit() Then
        Exit Function
    End If

    minOnFocusLossSettingVal = game.warGetMinimizeOnFocusLoss()
    If minOnFocusLossSettingVal < -1 Or minOnFocusLossSettingVal > 1 Then
        minOnFocusLossSettingVal = -1
    End If
    hint = minimizeOnFocusLossHint                  ' SDL_GetHint(SDL_HINT_VIDEO_MINIMIZE_ON_FOCUS_LOSS)
    If Len(hint) = 0 Or StrComp(hint, "auto", vbTextCompare) = 0 Then
        wzSetMinimizeOnFocusLoss minOnFocusLossSettingVal
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

    If hasBackend Then
        ' initialise all cursors
        If game.warGetColouredCursor() Then
            sdlInitColoredCursors
        Else
            sdlInitCursors
        End If
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
        win_updateClipCursor                ' (SDL: WIN_UpdateClipCursor on WM_WINDOWPOSCHANGED)
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

        ' Handle deferred size reset
        If deferredDimensionResetHasValue Then
            If WZwindow <> 0 Then
                win_setWindowSize deferredDimensionResetWidth, deferredDimensionResetHeight
                If deferredDimensionResetRecenter Then
                    win_centerWindowOnDisplay screenIndex
                End If
                wzReduceDisplayScalingIfNeeded deferredDimensionResetWidth, deferredDimensionResetHeight
            End If
            deferredDimensionResetHasValue = False
        Else
            wzReduceDisplayScalingIfNeeded newWindowWidth, newWindowHeight
        End If
        win_updateClipCursor
    Case EV_WINDOW_MINIMIZED
        debug_ 3594, LOG_INFO, "handleActiveEvent", "Window %d minimized", 1
        ' (SDL_OnWindowMinimized: a fullscreen mode gives the display back its desktop mode)
        If windowIsFullscreen Then win_restoreDesktopDisplayMode
        win_updateClipCursor
    Case EV_WINDOW_MAXIMIZED
        debug_ 3597, LOG_WZ, "handleActiveEvent", "Window %d maximized", 1
    Case EV_WINDOW_RESTORED
        debug_ 3600, LOG_INFO, "handleActiveEvent", "Window %d restored", 1
        ' (SDL_OnWindowRestored: the fullscreen mode set again)
        If windowIsFullscreen And hasFullscreenMode And Not exclusiveModeApplied Then win_setWindowFullscreen True
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
        win_updateClipCursor
    Case EV_WINDOW_FOCUS_LOST
        windowHasFocus = False
        debug_ 3636, LOG_WZ, "handleActiveEvent", "Window %d lost keyboard focus", 1
        For i = 0 To KEY_MAXSCAN - 1
            aKeyState(i).state = KEY_UP
            actualKeyState(i).state = KEY_UP
        Next
        win_updateClipCursor
        ' (SDL_OnWindowFocusLost: SDL_ShouldMinimizeOnFocusLoss)
        If win_shouldMinimizeOnFocusLoss() Then ShowWindow WZwindow, SW_MINIMIZE
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
    ' (VB) wzAsyncExecOnMainThread's task queued by wzMainScreenSetup_CreateVideoWindow
    If pendingFullscreenTransition Then
        pendingFullscreenTransition = False
        ' transition to fullscreen mode
        wzChangeWindowMode WINDOW_MODE_fullscreen
    End If
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
    sdlFreeCursors
    ' (VB) SDL_Quit: the cursor clip and the display mode given back
    mouseGrabbed = False
    win_updateClipCursor
    win_restoreDesktopDisplayMode
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
