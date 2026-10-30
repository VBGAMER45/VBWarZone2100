Attribute VB_Name = "chakra"
' Feasibility probe for a VB6 Warzone 2100 port: Windows' built-in Chakra
' JavaScript engine (SysWOW64\chakra.dll, the JsRT API) driven from VB6.
'  1. runtime + context, an ES2015 script (class, getter, let/const, arrow
'     functions, template literal, includes, **), result read back
'  2. ES2018 object spread / rest (legacy Chakra may lack them)
'  3. a VB function (AddressOf) called from JS: result and cost per call
'  4. JsParseScript over every Warzone .js file named in the list file given
'     on the command line: which ones legacy Chakra cannot parse
' Writes chakra.log next to the exe.
Option Explicit

Private Declare Function JsCreateRuntime Lib "chakra.dll" (ByVal attributes As Long, ByVal threadService As Long, runtime As Long) As Long
Private Declare Function JsCreateContext Lib "chakra.dll" (ByVal runtime As Long, context As Long) As Long
Private Declare Function JsSetCurrentContext Lib "chakra.dll" (ByVal context As Long) As Long
Private Declare Function JsDisposeRuntime Lib "chakra.dll" (ByVal runtime As Long) As Long
Private Declare Function JsGetGlobalObject Lib "chakra.dll" (globalObject As Long) As Long
Private Declare Function JsCreateFunction Lib "chakra.dll" (ByVal nativeFunction As Long, ByVal callbackState As Long, func As Long) As Long
Private Declare Function JsGetPropertyIdFromName Lib "chakra.dll" (ByVal pname As Long, propertyId As Long) As Long
Private Declare Function JsSetProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, ByVal value As Long, ByVal useStrictRules As Long) As Long
Private Declare Function JsGetProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, value As Long) As Long
Private Declare Function JsRunScript Lib "chakra.dll" (ByVal script As Long, ByVal sourceContext As Long, ByVal sourceUrl As Long, result As Long) As Long
Private Declare Function JsParseScript Lib "chakra.dll" (ByVal script As Long, ByVal sourceContext As Long, ByVal sourceUrl As Long, result As Long) As Long
Private Declare Function JsNumberToDouble Lib "chakra.dll" (ByVal value As Long, d As Double) As Long
Private Declare Function JsDoubleToNumber Lib "chakra.dll" (ByVal d As Double, value As Long) As Long
Private Declare Function JsConvertValueToString Lib "chakra.dll" (ByVal value As Long, s As Long) As Long
Private Declare Function JsStringToPointer Lib "chakra.dll" (ByVal value As Long, p As Long, length As Long) As Long
Private Declare Function JsGetAndClearException Lib "chakra.dll" (exception As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal cp As Long, ByVal flags As Long, ByVal src As Long, ByVal srcLen As Long, ByVal dst As Long, ByVal dstLen As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (c As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (f As Currency) As Long

Private logf As Integer
Private cookie As Long

Private Sub LogLine(ByVal s As String)
    Print #logf, s
End Sub

Private Function Now_() As Double
    Dim c As Currency, f As Currency
    QueryPerformanceCounter c
    QueryPerformanceFrequency f
    Now_ = CDbl(c) / CDbl(f)
End Function

Private Sub Check(ByVal r As Long, ByVal what As String)
    If r <> 0 Then Err.Raise vbObjectError + 1, , what & " failed: 0x" & Hex$(r)
End Sub

' The JS function add(a, b), written in VB (JsNativeFunction, stdcall).
' isConstructCall is a C++ bool and argumentCount an unsigned short, each in
' a 4-byte stack slot: only their low bytes are meaningful.
Public Function NativeAdd(ByVal callee As Long, ByVal isConstructCall As Long, ByVal args As Long, ByVal argCount As Long, ByVal state As Long) As Long
    Dim a(0 To 2) As Long, x As Double, y As Double, r As Long
    If (argCount And &HFFFF&) >= 3 Then
        CopyMemory a(0), ByVal args, 12     ' args[0] is this
        JsNumberToDouble a(1), x
        JsNumberToDouble a(2), y
    End If
    JsDoubleToNumber x + y, r
    NativeAdd = r
End Function

Private Function ValueToString(ByVal v As Long) As String
    Dim sref As Long, p As Long, n As Long, s As String
    If JsConvertValueToString(v, sref) <> 0 Then
        ValueToString = "(not convertible)"
        Exit Function
    End If
    If JsStringToPointer(sref, p, n) <> 0 Then Exit Function
    If n > 0 Then
        s = String$(n, 0)
        CopyMemory ByVal StrPtr(s), ByVal p, n * 2
    End If
    ValueToString = s
End Function

Private Function PropText(ByVal obj As Long, ByVal pname As String) As String
    Dim pid As Long, v As Long
    If JsGetPropertyIdFromName(StrPtr(pname), pid) <> 0 Then Exit Function
    If JsGetProperty(obj, pid, v) <> 0 Then Exit Function
    PropText = " " & pname & "=" & ValueToString(v)
End Function

Private Function ExceptionText(ByVal r As Long) As String
    Dim ex As Long, t As String
    t = "0x" & Hex$(r)
    If JsGetAndClearException(ex) = 0 Then
        t = t & " " & ValueToString(ex) & PropText(ex, "line") & PropText(ex, "column")
    End If
    ExceptionText = t
End Function

Private Function RunToString(ByVal script As String) As String
    Dim res As Long, r As Long
    cookie = cookie + 1
    r = JsRunScript(StrPtr(script), cookie, StrPtr("probe"), res)
    If r <> 0 Then
        RunToString = "ERROR " & ExceptionText(r)
    Else
        RunToString = ValueToString(res)
    End If
End Function

Private Function ReadUtf8(ByVal path As String) As String
    Dim f As Integer, n As Long, b() As Byte, st As Long, w As Long, s As String
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n = 0 Then
        Close #f
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #f, , b
    Close #f
    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then st = 3
    End If
    If n - st <= 0 Then Exit Function
    w = MultiByteToWideChar(65001, 0, VarPtr(b(st)), n - st, 0, 0)
    s = String$(w, 0)
    MultiByteToWideChar 65001, 0, VarPtr(b(st)), n - st, StrPtr(s), w
    ReadUtf8 = s
End Function

Sub Main()
    Dim rt As Long, ctx As Long, glob As Long, fn As Long, pid As Long, res As Long, r As Long
    Dim s As String, t0 As Double, t1 As Double, t2 As Double
    Dim listPath As String, path As String, lf As Integer
    Dim total As Long, ok As Long, bytes As Double
    On Error GoTo fail
    logf = FreeFile
    Open App.Path & "\chakra.log" For Output As #logf

    r = JsCreateRuntime(0, 0, rt)
    LogLine "JsCreateRuntime: 0x" & Hex$(r)
    If r <> 0 Then GoTo done
    Check JsCreateContext(rt, ctx), "JsCreateContext"
    Check JsSetCurrentContext(ctx), "JsSetCurrentContext"
    Check JsGetGlobalObject(glob), "JsGetGlobalObject"
    Check JsCreateFunction(AddressOf NativeAdd, 0, fn), "JsCreateFunction"
    Check JsGetPropertyIdFromName(StrPtr("add"), pid), "JsGetPropertyIdFromName"
    Check JsSetProperty(glob, pid, fn, 1), "JsSetProperty"

    ' 1. ES2015 and a VB function, expected 1,2,4,5,7,8|42|true|1024
    s = "(function () {" & vbLf & _
        "  'use strict';" & vbLf & _
        "  class Unit { constructor(id, hp) { this.id = id; this.hp = hp; } get alive() { return this.hp > 0; } }" & vbLf & _
        "  const units = [];" & vbLf & _
        "  for (let i = 0; i < 10; i++) units.push(new Unit(i, i % 3));" & vbLf & _
        "  const alive = units.filter(u => u.alive).map(u => u.id);" & vbLf & _
        "  return `${alive.join(',')}|${add(40, 2)}|${[1, 2, 3].includes(2)}|${2 ** 10}`;" & vbLf & _
        "})()"
    LogLine "ES2015 script (expect 1,2,4,5,7,8|42|true|1024): " & RunToString(s)

    ' 2. ES2018 object spread and rest, expected x,y|2
    s = "(function () { const o = {...{x: 1}, y: 2}; const {x, ...rest} = o; return Object.keys(o).join(',') + '|' + rest.y; })()"
    LogLine "ES2018 object spread/rest (expect x,y|2): " & RunToString(s)

    ' 3. cost of a VB function called from JS
    t0 = Now_()
    s = RunToString("(function () { let s = 0; for (let k = 0; k < 1000000; k++) s = add(s, 1); return s; })()")
    t1 = Now_()
    RunToString "(function () { function addj(a, b) { return a + b; } let s = 0; for (let k = 0; k < 1000000; k++) s = addj(s, 1); return s; })()"
    t2 = Now_()
    LogLine "1,000,000 calls of VB add() from JS: result " & s & ", " & Format$((t1 - t0) * 1000, "0") & " ms (" & _
        Format$(((t1 - t0) - (t2 - t1)) * 1000000# / 1000000#, "0.00") & " us per call more than a JS function)"

    ' 4. parse every Warzone script
    listPath = Trim$(Command$)
    If Len(listPath) > 0 Then
        lf = FreeFile
        Open listPath For Input As #lf
        t0 = Now_()
        Do While Not EOF(lf)
            Line Input #lf, path
            If Len(path) > 0 Then
                total = total + 1
                s = ReadUtf8(path)
                bytes = bytes + Len(s)
                cookie = cookie + 1
                r = JsParseScript(StrPtr(s), cookie, StrPtr(path), res)
                If r = 0 Then
                    ok = ok + 1
                Else
                    LogLine "  FAIL " & path & ": " & ExceptionText(r)
                End If
            End If
        Loop
        t1 = Now_()
        Close #lf
        LogLine "Warzone scripts parsed: " & ok & " of " & total & " (" & Format$(bytes / 1048576#, "0.0") & " M chars in " & Format$((t1 - t0) * 1000, "0") & " ms)"
    End If

done:
    JsSetCurrentContext 0
    If rt <> 0 Then JsDisposeRuntime rt
    LogLine "done"
    Close #logf
    Exit Sub
fail:
    LogLine "VB error " & Err.Number & ": " & Err.Description
    Close #logf
End Sub
