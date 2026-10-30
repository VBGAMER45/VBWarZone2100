Attribute VB_Name = "glprobe"
' Feasibility probe for a VB6 Warzone 2100 port: modern OpenGL from VB6.
'  1. a hidden window and a legacy context (opengl32.dll, declared directly)
'  2. an OpenGL 3.3 core context through wglCreateContextAttribsARB, a
'     function pointer called with DispCallFunc (oleaut32)
'  3. Warzone's own rect.vert / rect.frag (with rect_common.glsl), and
'     tcmask.vert / tcmask.frag (the model shader), compiled and linked as
'     gfx_api_gl.cpp does: "#version 330 core" + the source, includes inlined
'  4. the cost of a GL call: direct Declare, DispCallFunc, CallWindowProcW
' The shader folder is the command line. Writes glprobe.log next to the exe.
Option Explicit

Private Type WNDCLASSEXW
    cbSize As Long
    clsStyle As Long
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

Private Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    b(0 To 19) As Byte      ' iPixelType .. bReserved
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Private Declare Function RegisterClassExW Lib "user32" (wc As WNDCLASSEXW) As Integer
Private Declare Function CreateWindowExW Lib "user32" (ByVal exStyle As Long, ByVal cls As Long, ByVal title As Long, ByVal wstyle As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal parent As Long, ByVal menu As Long, ByVal inst As Long, ByVal param As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal fn As Long, ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal mname As String) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal mname As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal pname As String) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal p As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (c As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (f As Currency) As Long
Private Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare Function SetPixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal fmt As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare Function wglCreateContext Lib "opengl32" (ByVal hdc As Long) As Long
Private Declare Function wglMakeCurrent Lib "opengl32" (ByVal hdc As Long, ByVal hglrc As Long) As Long
Private Declare Function wglDeleteContext Lib "opengl32" (ByVal hglrc As Long) As Long
Private Declare Function wglGetProcAddress Lib "opengl32" (ByVal pname As String) As Long
Private Declare Function glGetString Lib "opengl32" (ByVal pname As Long) As Long
Private Declare Function glGetError Lib "opengl32" () As Long
Private Declare Sub glViewport Lib "opengl32" (ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
Private Declare Sub glGetIntegerv Lib "opengl32" (ByVal pname As Long, params As Long)
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, prgvt As Integer, prgpvarg As Long, pvargResult As Variant) As Long

Private Const CC_STDCALL As Long = 4
Private Const GL_VENDOR As Long = &H1F00&
Private Const GL_RENDERER As Long = &H1F01&
Private Const GL_VERSION As Long = &H1F02&
Private Const GL_SHADING_LANGUAGE_VERSION As Long = &H8B8C&
Private Const GL_FRAGMENT_SHADER As Long = &H8B30&
Private Const GL_VERTEX_SHADER As Long = &H8B31&
Private Const GL_COMPILE_STATUS As Long = &H8B81&
Private Const GL_LINK_STATUS As Long = &H8B82&

Private logf As Integer
Private className As String
Private shaderDir As String
Private hOpenGL As Long

Private Sub LogLine(ByVal s As String)
    Print #logf, s
End Sub

Private Function Now_() As Double
    Dim c As Currency, f As Currency
    QueryPerformanceCounter c
    QueryPerformanceFrequency f
    Now_ = CDbl(c) / CDbl(f)
End Function

Private Function CStrA(ByVal p As Long) As String
    Dim n As Long, b() As Byte
    If p = 0 Then
        CStrA = "(null)"
        Exit Function
    End If
    n = lstrlenA(p)
    If n = 0 Then Exit Function
    ReDim b(0 To n - 1)
    CopyMemory b(0), ByVal p, n
    CStrA = StrConv(b, vbUnicode)
End Function

' Calls a stdcall function pointer; the arguments are Longs (ints, enums, pointers).
Private Function CallPtr(ByVal fn As Long, ParamArray args() As Variant) As Long
    Dim n As Long, i As Long, hr As Long, dummyT As Integer, dummyP As Long
    Dim vt(0 To 7) As Integer, pv(0 To 7) As Long, a(0 To 7) As Variant, r As Variant
    If fn = 0 Then Err.Raise vbObjectError + 3, , "CallPtr: null function pointer"
    n = UBound(args) - LBound(args) + 1
    For i = 0 To n - 1
        a(i) = CLng(args(i))
        vt(i) = vbLong
        pv(i) = VarPtr(a(i))
    Next
    If n = 0 Then
        hr = DispCallFunc(0, fn, CC_STDCALL, vbLong, 0, dummyT, dummyP, r)
    Else
        hr = DispCallFunc(0, fn, CC_STDCALL, vbLong, n, vt(0), pv(0), r)
    End If
    If hr <> 0 Then Err.Raise vbObjectError + 2, , "DispCallFunc failed: 0x" & Hex$(hr)
    CallPtr = r
End Function

Private Function GLProc(ByVal pname As String) As Long
    GLProc = wglGetProcAddress(pname)
    If GLProc = 0 Then GLProc = GetProcAddress(hOpenGL, pname)    ' GL 1.1 functions
    If GLProc = 0 Then Err.Raise vbObjectError + 4, , "no GL function " & pname
End Function

Private Function ReadText(ByVal path As String) As String
    Dim f As Integer, b() As Byte, n As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    If n > 0 Then ReadText = StrConv(b, vbUnicode)
End Function

' gfx_api_gl.cpp's #include handling: a line #include "x" is replaced by the file x.
Private Function LoadShader(ByVal fname As String) As String
    Dim lines() As String, i As Long, t As String, q1 As Long, q2 As Long, out As String
    lines = Split(Replace(ReadText(shaderDir & "\" & fname), vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        t = Trim$(Replace(lines(i), vbTab, " "))
        If Left$(t, 8) = "#include" Then
            q1 = InStr(t, """")
            q2 = InStr(q1 + 1, t, """")
            out = out & LoadShader(Mid$(t, q1 + 1, q2 - q1 - 1)) & vbLf
        Else
            out = out & lines(i) & vbLf
        End If
    Next
    LoadShader = out
End Function

Private Function CompileShader(ByVal kind As Long, ByVal fname As String) As Long
    Dim src As String, b() As Byte, ptrs(0) As Long, sh As Long, status As Long
    Dim logLen As Long, lg() As Byte
    src = "#version 330 core" & vbLf & LoadShader(fname)
    b = StrConv(src & vbNullChar, vbFromUnicode)
    ptrs(0) = VarPtr(b(0))
    sh = CallPtr(GLProc("glCreateShader"), kind)
    CallPtr GLProc("glShaderSource"), sh, 1, VarPtr(ptrs(0)), 0
    CallPtr GLProc("glCompileShader"), sh
    CallPtr GLProc("glGetShaderiv"), sh, GL_COMPILE_STATUS, VarPtr(status)
    If status = 0 Then
        ReDim lg(0 To 4095)
        CallPtr GLProc("glGetShaderInfoLog"), sh, 4096, VarPtr(logLen), VarPtr(lg(0))
        LogLine "  " & fname & " FAILED: " & CStrA(VarPtr(lg(0)))
        CompileShader = 0
    Else
        LogLine "  " & fname & " compiled (" & Len(src) & " chars)"
        CompileShader = sh
    End If
End Function

Private Sub LinkProgram(ByVal vertName As String, ByVal fragName As String)
    Dim vs As Long, fs As Long, prog As Long, status As Long
    vs = CompileShader(GL_VERTEX_SHADER, vertName)
    fs = CompileShader(GL_FRAGMENT_SHADER, fragName)
    If vs = 0 Or fs = 0 Then Exit Sub
    prog = CallPtr(GLProc("glCreateProgram"))
    CallPtr GLProc("glAttachShader"), prog, vs
    CallPtr GLProc("glAttachShader"), prog, fs
    CallPtr GLProc("glLinkProgram"), prog
    CallPtr GLProc("glGetProgramiv"), prog, GL_LINK_STATUS, VarPtr(status)
    LogLine "  program " & vertName & " + " & fragName & ": " & IIf(status <> 0, "linked", "LINK FAILED")
End Sub

Sub Main()
    Dim wc As WNDCLASSEXW, pfd As PIXELFORMATDESCRIPTOR, hwnd As Long, hdc As Long
    Dim legacy As Long, core As Long, pf As Long, pCreate As Long, attribs(0 To 6) As Long
    Dim pViewport As Long, i As Long, t0 As Double, t1 As Double, n As Long
    Dim vt(0 To 3) As Integer, pv(0 To 3) As Long, a(0 To 3) As Variant, r As Variant
    On Error GoTo fail
    logf = FreeFile
    Open App.Path & "\glprobe.log" For Output As #logf
    shaderDir = Trim$(Command$)
    hOpenGL = LoadLibraryA("opengl32.dll")

    ' 1. window and legacy context
    className = "WzVbGlProbe"
    wc.cbSize = Len(wc)
    wc.clsStyle = &H20                ' CS_OWNDC
    wc.lpfnWndProc = GetProcAddress(GetModuleHandleA("user32"), "DefWindowProcW")
    wc.hInstance = App.hInstance
    wc.lpszClassName = StrPtr(className)
    If RegisterClassExW(wc) = 0 Then Err.Raise vbObjectError + 5, , "RegisterClassExW failed"
    hwnd = CreateWindowExW(0, StrPtr(className), StrPtr("probe"), &H80000000, 0, 0, 64, 64, 0, 0, App.hInstance, 0)
    hdc = GetDC(hwnd)
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = &H25                ' PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pfd.b(1) = 32                     ' cColorBits
    pfd.b(15) = 24                    ' cDepthBits
    pfd.b(16) = 8                     ' cStencilBits
    pf = ChoosePixelFormat(hdc, pfd)
    If SetPixelFormat(hdc, pf, pfd) = 0 Then Err.Raise vbObjectError + 6, , "SetPixelFormat failed"
    legacy = wglCreateContext(hdc)
    wglMakeCurrent hdc, legacy
    LogLine "legacy context: " & CStrA(glGetString(GL_VENDOR)) & " / " & CStrA(glGetString(GL_RENDERER)) & " / " & CStrA(glGetString(GL_VERSION))

    ' 2. OpenGL 3.3 core context through a function pointer
    pCreate = wglGetProcAddress("wglCreateContextAttribsARB")
    If pCreate = 0 Then Err.Raise vbObjectError + 7, , "no wglCreateContextAttribsARB"
    attribs(0) = &H2091: attribs(1) = 3        ' WGL_CONTEXT_MAJOR_VERSION_ARB
    attribs(2) = &H2092: attribs(3) = 3        ' WGL_CONTEXT_MINOR_VERSION_ARB
    attribs(4) = &H9126&: attribs(5) = 1       ' WGL_CONTEXT_PROFILE_MASK_ARB = CORE
    core = CallPtr(pCreate, hdc, 0, VarPtr(attribs(0)))
    If core = 0 Then Err.Raise vbObjectError + 8, , "wglCreateContextAttribsARB(3.3 core) failed"
    wglMakeCurrent hdc, core
    wglDeleteContext legacy
    LogLine "3.3 core context: GL " & CStrA(glGetString(GL_VERSION)) & ", GLSL " & CStrA(glGetString(GL_SHADING_LANGUAGE_VERSION))

    ' 3. Warzone's shaders
    If Len(shaderDir) > 0 Then
        LogLine "Warzone shaders from " & shaderDir & ":"
        LinkProgram "rect.vert", "rect.frag"
        LinkProgram "tcmask.vert", "tcmask.frag"
    End If

    ' 4. cost of a call
    n = 1000000
    t0 = Now_()
    For i = 1 To n
        glViewport 0, 0, 64, 64
    Next
    t1 = Now_()
    LogLine "glViewport, direct Declare: " & Format$((t1 - t0) * 1000000# / n, "0.000") & " us per call"

    pViewport = GLProc("glViewport")
    For i = 0 To 3
        vt(i) = vbLong
        pv(i) = VarPtr(a(i))
    Next
    a(0) = 0&: a(1) = 0&: a(2) = 64&: a(3) = 64&
    n = 200000
    t0 = Now_()
    For i = 1 To n
        DispCallFunc 0, pViewport, CC_STDCALL, vbEmpty, 4, vt(0), pv(0), r
    Next
    t1 = Now_()
    a(2) = 33&: a(3) = 17&
    i = DispCallFunc(0, pViewport, CC_STDCALL, vbEmpty, 4, vt(0), pv(0), r)
    glGetIntegerv &HBA2&, attribs(0)
    LogLine "  check: DispCallFunc hr=0x" & Hex$(i) & ", viewport now " & attribs(2) & "x" & attribs(3) & " (expect 33x17)"
    LogLine "glViewport, pointer via DispCallFunc (arrays prepared): " & Format$((t1 - t0) * 1000000# / n, "0.000") & " us per call"

    n = 1000000
    t0 = Now_()
    For i = 1 To n
        CallWindowProcW pViewport, 0, 0, 64, 64
    Next
    t1 = Now_()
    LogLine "glViewport, pointer via CallWindowProcW (4 arguments only): " & Format$((t1 - t0) * 1000000# / n, "0.000") & " us per call"
    LogLine "glGetError after the calls: " & glGetError()

    wglMakeCurrent 0, 0
    wglDeleteContext core
    ReleaseDC hwnd, hdc
    DestroyWindow hwnd
    LogLine "done"
    Close #logf
    Exit Sub
fail:
    LogLine "VB error " & Err.Number & ": " & Err.Description
    Close #logf
End Sub
