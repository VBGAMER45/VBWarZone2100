VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "win_OpenGL_Impl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sdl/gfx_api_gl_sdl.cpp: sdl_OpenGL_Impl, on WGL instead of SDL. The
' pixel format carries what wzSDLPreWindowCreate_InitOpenGLAttributes asks
' SDL for (RGB at least 5 bits each, double buffer, 24-bit depth, 8-bit
' stencil, no sRGB, no multisampling); contexts are requested through
' wglCreateContextAttribsARB, down the same list (3.3, 3.2, 3.1, 3.0 core,
' forward compatible). OpenGL ES (OpenGLES30) is not available through WGL:
' that request fails, as when the ES library is missing. Calls through GL
' function pointers use DispCallFunc.
Option Explicit

Implements IGLBackend

Private Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare Function SetPixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal iPixelFormat As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DescribePixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal iPixelFormat As Long, ByVal nBytes As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare Function SwapBuffers Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function wglCreateContext Lib "opengl32" (ByVal hdc As Long) As Long
Private Declare Function wglMakeCurrent Lib "opengl32" (ByVal hdc As Long, ByVal hglrc As Long) As Long
Private Declare Function wglDeleteContext Lib "opengl32" (ByVal hglrc As Long) As Long
Private Declare Function wglGetCurrentContext Lib "opengl32" () As Long
Private Declare Function wglGetProcAddress Lib "opengl32" (ByVal lpszProc As String) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, prgvt As Integer, prgpvarg As Long, pvargResult As Variant) As Long

Private Const PFD_DOUBLEBUFFER As Long = 1
Private Const PFD_DRAW_TO_WINDOW As Long = 4
Private Const PFD_SUPPORT_OPENGL As Long = &H20
Private Const PFD_TYPE_RGBA As Byte = 0
Private Const CC_STDCALL As Long = 4
Private Const WGL_CONTEXT_MAJOR_VERSION_ARB As Long = &H2091&
Private Const WGL_CONTEXT_MINOR_VERSION_ARB As Long = &H2092&
Private Const WGL_CONTEXT_FLAGS_ARB As Long = &H2094&
Private Const WGL_CONTEXT_PROFILE_MASK_ARB As Long = &H9126&
Private Const WGL_CONTEXT_FORWARD_COMPATIBLE_BIT_ARB As Long = 2
Private Const WGL_CONTEXT_CORE_PROFILE_BIT_ARB As Long = 1

' sdl_OpenGL_Impl::GLContextRequests
Private Const OpenGLCore_3_3 As Long = 0
Private Const OpenGLCore_3_2 As Long = 1
Private Const OpenGLCore_3_1 As Long = 2
Private Const OpenGLCore_3_0 As Long = 3
Private Const OpenGLES30 As Long = 4
Private Const MAX_CONTEXT_REQUESTS As Long = 5
Private Const OpenGLCore_HighestAvailable As Long = OpenGLCore_3_3

Private window As Long
Private hdc As Long
Private glrc As Long
Private hOpenGL As Long
Private useOpenglES As Boolean
Private useOpenGLESLibrary As Boolean
Private contextRequest As Long
Private requestMajor As Long
Private requestMinor As Long
Private requestFlags As Long
Private requestProfile As Long
Private lastError As String

' C: sdl_OpenGL_Impl(SDL_Window* window, bool useOpenGLES, bool useOpenGLESLibrary)
Friend Sub Init(ByVal hwnd As Long, ByVal useOpenGLES_ As Boolean, ByVal useOpenGLESLibrary_ As Boolean)
    If hwnd = 0 Then ASSERT_FAILURE "_window != nullptr", "lib/sdl/gfx_api_gl_sdl.cpp:32", "sdl_OpenGL_Impl", "Invalid SDL_Window*"
    window = hwnd
    useOpenglES = useOpenGLES_
    contextRequest = getInitialContextRequest(useOpenglES)
    useOpenGLESLibrary = useOpenGLESLibrary_
    configureOpenGLContextRequest contextRequest, useOpenGLESLibrary
End Sub

Private Function IGLBackend_getGLGetProcAddress(ByVal procName As String) As Long
    Dim p As Long
    p = wglGetProcAddress(procName)
    ' wglGetProcAddress knows only extensions and post-1.1 functions; the 1.1 ones are opengl32.dll exports
    If p = 0 Or p = 1 Or p = 2 Or p = 3 Or p = -1 Then
        If hOpenGL = 0 Then hOpenGL = LoadLibraryA("opengl32.dll")
        p = GetProcAddress(hOpenGL, procName)
    End If
    IGLBackend_getGLGetProcAddress = p
End Function

Private Function configureNextOpenGLContextRequest() As Boolean
    contextRequest = contextRequest + 1
    configureNextOpenGLContextRequest = configureOpenGLContextRequest(contextRequest, useOpenGLESLibrary)
End Function

Private Function getInitialContextRequest(ByVal useOpenglES_ As Boolean) As Long
    If useOpenglES_ Then
        getInitialContextRequest = OpenGLES30
    Else
        getInitialContextRequest = OpenGLCore_HighestAvailable
    End If
End Function

' The SDL_GL_SetAttribute calls of the request, kept for createGLContext.
Private Function configureOpenGLContextRequest(ByVal request As Long, ByVal useOpenGLESLibrary_ As Boolean) As Boolean
    Select Case request
    Case OpenGLCore_3_3, OpenGLCore_3_2, OpenGLCore_3_1, OpenGLCore_3_0
        requestFlags = WGL_CONTEXT_FORWARD_COMPATIBLE_BIT_ARB
        requestProfile = WGL_CONTEXT_CORE_PROFILE_BIT_ARB
        requestMajor = 3
        requestMinor = 3 - request
        configureOpenGLContextRequest = True
    Case OpenGLES30
        ' (VB) not through WGL
        configureOpenGLContextRequest = False
    Case Else
        configureOpenGLContextRequest = False
    End Select
End Function

Private Function to_string(ByVal request As Long) As String
    ' (C++ reads the member contextRequest here, not the argument)
    Select Case contextRequest
    Case OpenGLCore_3_3: to_string = "OpenGL 3.3"
    Case OpenGLCore_3_2: to_string = "OpenGL 3.2"
    Case OpenGLCore_3_1: to_string = "OpenGL 3.1"
    Case OpenGLCore_3_0: to_string = "OpenGL 3.0"
    Case OpenGLES30: to_string = "OpenGL ES 3.0"
    End Select
End Function

Private Function IGLBackend_isOpenGLES() As Boolean
    IGLBackend_isOpenGLES = contextRequest >= OpenGLES30
End Function

' Calls a stdcall function pointer with Long arguments.
Private Function CallPtr(ByVal fn As Long, ParamArray args() As Variant) As Long
    Dim n As Long, i As Long, dummyT As Integer, dummyP As Long
    Dim vt(0 To 7) As Integer, pv(0 To 7) As Long, a(0 To 7) As Variant, r As Variant
    n = UBound(args) - LBound(args) + 1
    For i = 0 To n - 1
        a(i) = CLng(args(i))
        vt(i) = vbLong
        pv(i) = VarPtr(a(i))
    Next
    If n = 0 Then
        DispCallFunc 0, fn, CC_STDCALL, vbLong, 0, dummyT, dummyP, r
    Else
        DispCallFunc 0, fn, CC_STDCALL, vbLong, n, vt(0), pv(0), r
    End If
    CallPtr = r
End Function

' The pixel format SDL sets on a window created for OpenGL.
Private Function SetWindowPixelFormat() As Boolean
    Dim pfd As PIXELFORMATDESCRIPTOR, pf As Long
    If GetPixelFormat(hdc) <> 0 Then
        SetWindowPixelFormat = True
        Exit Function
    End If
    pfd.nSize = LenB(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cRedBits = 5
    pfd.cGreenBits = 5
    pfd.cBlueBits = 5
    pfd.cColorBits = 15
    pfd.cDepthBits = 24
    pfd.cStencilBits = 8
    pf = ChoosePixelFormat(hdc, pfd)
    If pf = 0 Then
        lastError = "ChoosePixelFormat failed"
        Exit Function
    End If
    If SetPixelFormat(hdc, pf, pfd) = 0 Then
        lastError = "SetPixelFormat failed"
        Exit Function
    End If
    SetWindowPixelFormat = True
End Function

' C: SDL_GL_CreateContext with the configured request: 0 on failure (lastError says why).
Private Function CreateContextForRequest() As Long
    Dim legacy As Long, pCreate As Long, ctx As Long, attribs(0 To 8) As Long
    If Not configureOpenGLContextRequest(contextRequest, useOpenGLESLibrary) Then
        lastError = "OpenGL ES is not available through WGL"
        Exit Function
    End If
    legacy = wglCreateContext(hdc)
    If legacy = 0 Then
        lastError = "wglCreateContext failed"
        Exit Function
    End If
    wglMakeCurrent hdc, legacy
    pCreate = wglGetProcAddress("wglCreateContextAttribsARB")
    If pCreate = 0 Then
        lastError = "WGL_ARB_create_context is not supported"
    Else
        attribs(0) = WGL_CONTEXT_MAJOR_VERSION_ARB: attribs(1) = requestMajor
        attribs(2) = WGL_CONTEXT_MINOR_VERSION_ARB: attribs(3) = requestMinor
        attribs(4) = WGL_CONTEXT_FLAGS_ARB: attribs(5) = requestFlags
        attribs(6) = WGL_CONTEXT_PROFILE_MASK_ARB: attribs(7) = requestProfile
        ctx = CallPtr(pCreate, hdc, 0, VarPtr(attribs(0)))
        If ctx = 0 Then lastError = "wglCreateContextAttribsARB failed"
    End If
    wglMakeCurrent 0, 0
    wglDeleteContext legacy
    If ctx <> 0 Then wglMakeCurrent hdc, ctx
    CreateContextForRequest = ctx
End Function

Private Function IGLBackend_createGLContext() As Boolean
    Dim WZglcontext As Long, glContextErrors As String
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim windowWidth As Long, windowHeight As Long, rc As RECT

    If hdc = 0 Then hdc = GetDC(window)
    If Not SetWindowPixelFormat() Then
        debug_ 0, LOG_ERROR, "createGLContext", "(VB) Failed to set the window's pixel format: %s", lastError
        Exit Function
    End If

    WZglcontext = CreateContextForRequest()
    Do While WZglcontext = 0
        glContextErrors = glContextErrors & "Failed to create an " & to_string(contextRequest) & " context! [" & lastError & "]" & vbLf
        If Not configureNextOpenGLContextRequest() Then
            debug_multiline 184, LOG_ERROR, "createGLContext", glContextErrors
            Exit Function
        End If
        WZglcontext = CreateContextForRequest()
    Loop
    glrc = WZglcontext

    If Len(glContextErrors) > 0 Then
        debug_multiline 192, LOG_3D, "createGLContext", glContextErrors
    End If

    debug_ 194, LOG_INFO, "createGLContext", "Requested %s context", to_string(contextRequest)

    DescribePixelFormat hdc, GetPixelFormat(hdc), LenB(pfd), pfd
    If (pfd.dwFlags And PFD_DOUBLEBUFFER) = 0 Then
        debug_ 201, LOG_ERROR, "createGLContext", "OpenGL initialization did not give double buffering! (%d)", 0
        debug_ 202, LOG_ERROR, "createGLContext", "Double buffering is required for this game!"
        wglMakeCurrent 0, 0
        wglDeleteContext WZglcontext
        glrc = 0
        Exit Function
    End If

    debug_ 234, LOG_3D, "createGLContext", "Current values for: SDL_GL_RED_SIZE (%d), SDL_GL_GREEN_SIZE (%d), SDL_GL_BLUE_SIZE (%d), SDL_GL_ALPHA_SIZE (%d)", _
        CLng(pfd.cRedBits), CLng(pfd.cGreenBits), CLng(pfd.cBlueBits), CLng(pfd.cAlphaBits)
    debug_ 240, LOG_3D, "createGLContext", "Current value for SDL_GL_DEPTH_SIZE: (%d)", CLng(pfd.cDepthBits)
    debug_ 246, LOG_3D, "createGLContext", "Current value for SDL_GL_STENCIL_SIZE: (%d)", CLng(pfd.cStencilBits)
    debug_ 254, LOG_3D, "createGLContext", "Current value for SDL_GL_FRAMEBUFFER_SRGB_CAPABLE: (%d)", 0

    GetClientRect window, rc
    windowWidth = rc.Right - rc.Left
    windowHeight = rc.Bottom - rc.Top
    debug_ 259, LOG_WZ, "createGLContext", "Logical Window Size: %d x %d", windowWidth, windowHeight
    IGLBackend_createGLContext = True
End Function

Private Function IGLBackend_destroyGLContext() As Boolean
    Dim WZglcontext As Long
    WZglcontext = wglGetCurrentContext()
    If WZglcontext = 0 Then Exit Function
    wglMakeCurrent 0, 0
    wglDeleteContext WZglcontext
    glrc = 0
    If hdc <> 0 Then
        ReleaseDC window, hdc
        hdc = 0
    End If
    IGLBackend_destroyGLContext = True
End Function

Private Sub IGLBackend_swapWindow()
    SwapBuffers hdc
End Sub

Private Sub IGLBackend_getDrawableSize(w As Long, h As Long)
    Dim rc As RECT
    GetClientRect window, rc
    w = rc.Right - rc.Left
    h = rc.Bottom - rc.Top
End Sub

Private Function IGLBackend_setSwapInterval(ByVal mode As swap_interval_mode) As Boolean
    Dim interval As Long, pSwap As Long
    Select Case mode
    Case swap_interval_mode_immediate: interval = 0
    Case swap_interval_mode_vsync: interval = 1
    Case swap_interval_mode_adaptive_vsync: interval = -1
    Case Else
        debug_ 360, LOG_3D, "setSwapInterval", "Unsupported swap_interval_mode for SDL OpenGL backend: %d", CLng(mode)
        Exit Function
    End Select
    pSwap = wglGetProcAddress("wglSwapIntervalEXT")
    If pSwap = 0 Then
        debug_ 365, LOG_WARNING, "setSwapInterval", "SDL_GL_SetSwapInterval(%d) failed with error (%s)", interval, "WGL_EXT_swap_control is not supported"
        Exit Function
    End If
    If CallPtr(pSwap, interval) = 0 Then
        debug_ 365, LOG_WARNING, "setSwapInterval", "SDL_GL_SetSwapInterval(%d) failed with error (%s)", interval, "wglSwapIntervalEXT failed"
        Exit Function
    End If
    IGLBackend_setSwapInterval = True
End Function

Private Function IGLBackend_getSwapInterval() As swap_interval_mode
    Dim pGet As Long, interval As Long
    pGet = wglGetProcAddress("wglGetSwapIntervalEXT")
    If pGet <> 0 Then interval = CallPtr(pGet)
    Select Case interval
    Case 0: IGLBackend_getSwapInterval = swap_interval_mode_immediate
    Case -1: IGLBackend_getSwapInterval = swap_interval_mode_adaptive_vsync
    Case Else: IGLBackend_getSwapInterval = swap_interval_mode_vsync
    End Select
End Function
