VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlatformApi"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' WzPlatform.dll's public interface: lib/sdl (main_win), with the input API of
' lib/framework/input.h and the window part of wzapp.h. GlobalMultiUse: the
' exe calls these procedures as if they were its own. Calls back into the
' game go through IPlatformGame.
Option Explicit

' input.h
Public Enum KEY_CODE
    KEY_ESC = 27
    KEY_1 = 49
    KEY_2 = 50
    KEY_3 = 51
    KEY_4 = 52
    KEY_5 = 53
    KEY_6 = 54
    KEY_7 = 55
    KEY_8 = 56
    KEY_9 = 57
    KEY_0 = 48
    KEY_MINUS = 45
    KEY_EQUALS = 61
    KEY_BACKSPACE = 8
    KEY_TAB = 9
    KEY_Q = 113
    KEY_W = 119
    KEY_E = 101
    KEY_R = 114
    KEY_T = 116
    KEY_Y = 121
    KEY_U = 117
    KEY_I = 105
    KEY_O = 111
    KEY_P = 112
    KEY_LBRACE = 91
    KEY_RBRACE = 93
    KEY_RETURN = 13         ' Comes from SDL, don't know why not '\n'.
    KEY_LCTRL = 306
    KEY_A = 97
    KEY_S = 115
    KEY_D = 100
    KEY_F = 102
    KEY_G = 103
    KEY_H = 104
    KEY_J = 106
    KEY_K = 107
    KEY_L = 108
    KEY_SEMICOLON = 59
    KEY_QUOTE = 39
    KEY_BACKQUOTE = 96
    KEY_LSHIFT = 304
    KEY_LMETA = 310
    KEY_LSUPER = 311
    KEY_BACKSLASH = 92
    KEY_Z = 122
    KEY_X = 120
    KEY_C = 99
    KEY_V = 118
    KEY_B = 98
    KEY_N = 110
    KEY_M = 109
    KEY_COMMA = 44
    KEY_FULLSTOP = 46
    KEY_FORWARDSLASH = 47
    KEY_RSHIFT = 303
    KEY_RMETA = 309
    KEY_RSUPER = 312
    KEY_KP_STAR = 268
    KEY_LALT = 308
    KEY_SPACE = 32
    KEY_CAPSLOCK = 301
    KEY_F1 = 282
    KEY_F2 = 283
    KEY_F3 = 284
    KEY_F4 = 285
    KEY_F5 = 286
    KEY_F6 = 287
    KEY_F7 = 288
    KEY_F8 = 289
    KEY_F9 = 290
    KEY_F10 = 291
    KEY_NUMLOCK = 300
    KEY_SCROLLLOCK = 302
    KEY_KP_7 = 263
    KEY_KP_8 = 264
    KEY_KP_9 = 265
    KEY_KP_MINUS = 269
    KEY_KP_4 = 260
    KEY_KP_5 = 261
    KEY_KP_6 = 262
    KEY_KP_PLUS = 270
    KEY_KP_1 = 257
    KEY_KP_2 = 258
    KEY_KP_3 = 259
    KEY_KP_0 = 256
    KEY_KP_FULLSTOP = 266
    KEY_F11 = 292
    KEY_F12 = 293
    KEY_RCTRL = 305
    KEY_KP_BACKSLASH = 267  ' Most keypads just have a forward slash.
    KEY_RALT = 307
    KEY_HOME = 278
    KEY_UPARROW = 273
    KEY_PAGEUP = 280
    KEY_LEFTARROW = 276
    KEY_RIGHTARROW = 275
    KEY_END = 279
    KEY_DOWNARROW = 274
    KEY_PAGEDOWN = 281
    KEY_INSERT = 277
    KEY_DELETE = 127
    KEY_KPENTER = 271
    KEY_MAXSCAN = 323       ' The largest possible scan code.
    KEY_IGNORE = 5190
End Enum

Public Enum MOUSE_KEY_CODE
    MOUSE_LMB = 1
    MOUSE_MMB
    MOUSE_RMB
    MOUSE_X1                ' mouse button 4
    MOUSE_X2                ' mouse button 5
    MOUSE_WUP
    MOUSE_WDN
    MOUSE_END               ' end of our enum
End Enum

' MousePress::Action
Public Enum MousePress_Action
    MousePress_None
    MousePress_Press
    MousePress_Release
End Enum

' input.h's INPBUF_ macros
Public Enum INPBUF_KEYS
    INPBUF_LEFT = 276
    INPBUF_RIGHT = 275
    INPBUF_UP = 273
    INPBUF_DOWN = 274
    INPBUF_HOME = 278
    INPBUF_END = 279
    INPBUF_INS = 277
    INPBUF_DEL = 127
    INPBUF_PGUP = 280
    INPBUF_PGDN = 281
    INPBUF_BKSPACE = 8
    INPBUF_TAB = 9
    INPBUF_CR = 13
    INPBUF_ESC = 27
End Enum

' Name and version of this DLL, for the log header.
Public Function PlatformBuildInfo() As String
    PlatformBuildInfo = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision & " (uses " & FrameworkBuildInfo() & ")"
End Function

'--------------------------------------------------------------------------
' Startup, screen setup, the event loop

Public Sub wzMain()
    main_win.wzMain
End Sub

' (VB) The game's IPlatformGame; Nothing at shutdown.
Public Sub PlatformSetGame(ByVal g As IPlatformGame)
    main_win.PlatformSetGame g
End Sub

' (VB) Scripted test runs: quit after this many frames (0 = never).
Public Sub wzSetQuitAfterFrames(ByVal frames As Long)
    main_win.wzSetQuitAfterFrames frames
End Sub

Public Function wzGetFramesDone() As Long
    wzGetFramesDone = main_win.wzGetFramesDone()
End Function

Public Function wzGetDefaultGfxBackendForCurrentSystem() As video_backend
    wzGetDefaultGfxBackendForCurrentSystem = main_win.wzGetDefaultGfxBackendForCurrentSystem()
End Function

' backend: hasBackend = False for headless (C: nullopt)
Public Function wzMainScreenSetup(ByVal hasBackend As Boolean, ByVal backend As video_backend, ByVal antialiasing As Long, ByVal fullscreen As WINDOW_MODE, ByVal vsync As Long, ByVal lodDistanceBiasPercentage As Long, ByVal depthMapResolution As Long, Optional ByVal highDPI As Boolean = True) As Boolean
    wzMainScreenSetup = main_win.wzMainScreenSetup(hasBackend, backend, antialiasing, fullscreen, vsync, lodDistanceBiasPercentage, depthMapResolution, highDPI)
End Function

' callOnShutdown: call IPlatformGame.mainShutdown when the loop ends (C: onShutdown)
Public Sub wzMainEventLoop(ByVal callOnShutdown As Boolean)
    main_win.wzMainEventLoop callOnShutdown
End Sub

Public Sub wzPumpEventsWhileLoading()
    main_win.wzPumpEventsWhileLoading
End Sub

Public Sub wzQuit(ByVal exitCode As Long)
    main_win.wzQuit exitCode
End Sub

Public Function wzGetQuitExitCode() As Long
    wzGetQuitExitCode = main_win.wzGetQuitExitCode()
End Function

Public Sub wzShutdown()
    main_win.wzShutdown
End Sub

Public Function wzGetPlatform() As String
    wzGetPlatform = main_win.wzGetPlatform()
End Function

'--------------------------------------------------------------------------
' Window

Public Property Get screenWidth() As Long
    screenWidth = main_win.screenWidth
End Property

Public Property Get screenHeight() As Long
    screenHeight = main_win.screenHeight
End Property

Public Function wzGetCurrentWindowMode() As WINDOW_MODE
    wzGetCurrentWindowMode = main_win.wzGetCurrentWindowMode()
End Function

Public Function wzChangeWindowMode(ByVal mode As WINDOW_MODE, Optional ByVal silent As Boolean = False) As Boolean
    wzChangeWindowMode = main_win.wzChangeWindowMode(mode, silent)
End Function

Public Function wzIsFullscreen() As Boolean
    wzIsFullscreen = main_win.wzIsFullscreen()
End Function

Public Function wzWindowHasFocus() As Boolean
    wzWindowHasFocus = main_win.wzWindowHasFocus()
End Function

Public Function wzGetCurrentDisplayScale() As Long
    wzGetCurrentDisplayScale = main_win.wzGetCurrentDisplayScale()
End Function

' wzapp.h: wzGetGameToRendererScaleFactorInt (unsigned int percentages)
Public Sub wzGetGameToRendererScaleFactorInt(horizScalePercentage As Long, vertScalePercentage As Long)
    main_win.wzGetGameToRendererScaleFactorInt horizScalePercentage, vertScalePercentage
End Sub

Public Sub wzShowMouse(ByVal visible As Boolean)
    main_win.wzShowMouse visible
End Sub

'--------------------------------------------------------------------------
' Input (input.h)

Public Sub inputInitialise()
    main_win.inputInitialise
End Sub

Public Sub inputNewFrame()
    main_win.inputNewFrame
End Sub

Public Sub inputLoseFocus()
    main_win.inputLoseFocus
End Sub

Public Sub inputRestoreMetaKeyState()
    main_win.inputRestoreMetaKeyState
End Sub

Public Sub inputClearBuffer()
    main_win.inputClearBuffer
End Sub

Public Function inputGetKey(unicode As Long) As Long
    inputGetKey = main_win.inputGetKey(unicode)
End Function

Public Function inputGetClickCount() As Long
    inputGetClickCount = main_win.inputGetClickCount()
End Function

Public Sub inputGetClick(ByVal i As Long, action As Long, key As Long, x As Long, y As Long)
    main_win.inputGetClick i, action, key, x, y
End Sub

Public Function keyDown(ByVal code As KEY_CODE) As Boolean
    keyDown = main_win.keyDown(code)
End Function

Public Function keyPressed(ByVal code As KEY_CODE) As Boolean
    keyPressed = main_win.keyPressed(code)
End Function

Public Function keyReleased(ByVal code As KEY_CODE) As Boolean
    keyReleased = main_win.keyReleased(code)
End Function

Public Function keyScanToString(ByVal code As KEY_CODE) As String
    keyScanToString = main_win.keyScanToString(code)
End Function

Public Function mouseKeyCodeToString(ByVal code As MOUSE_KEY_CODE) As String
    mouseKeyCodeToString = main_win.mouseKeyCodeToString(code)
End Function

Public Function mouseX() As Long
    mouseX = main_win.mouseX()
End Function

Public Function mouseY() As Long
    mouseY = main_win.mouseY()
End Function

Public Function wzMouseInWindow() As Boolean
    wzMouseInWindow = main_win.wzMouseInWindow()
End Function

Public Sub getMouseWheelSpeed(x As Long, y As Long)
    main_win.getMouseWheelSpeed x, y
End Sub

Public Function mouseDown(ByVal code As MOUSE_KEY_CODE) As Boolean
    mouseDown = main_win.mouseDown(code)
End Function

Public Function mouseDClicked(ByVal code As MOUSE_KEY_CODE) As Boolean
    mouseDClicked = main_win.mouseDClicked(code)
End Function

Public Function mousePressed(ByVal code As MOUSE_KEY_CODE) As Boolean
    mousePressed = main_win.mousePressed(code)
End Function

Public Function mouseReleased(ByVal code As MOUSE_KEY_CODE) As Boolean
    mouseReleased = main_win.mouseReleased(code)
End Function

Public Function mouseDrag(ByVal code As MOUSE_KEY_CODE, px As Long, py As Long) As Boolean
    mouseDrag = main_win.mouseDrag(code, px, py)
End Function

Public Sub StartTextInput(ByVal pTextInputRequester As Long)
    main_win.StartTextInput pTextInputRequester
End Sub

Public Sub StopTextInput(ByVal pTextInputResigner As Long)
    main_win.StopTextInput pTextInputResigner
End Sub

Public Function isInTextInputMode() As Boolean
    isInTextInputMode = main_win.isInTextInputMode()
End Function
