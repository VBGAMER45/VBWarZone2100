Attribute VB_Name = "map_script_c"
' lib/wzmaplib/src/map_script.cpp, map_script.h: script-generated maps (a map
' folder's game.js, run by runMapScript with a seed). QuickJS's place is taken
' by Windows' chakra.dll (JsRT, the legacy Edge mode, as the exe's scripts:
' docs/CONVENTIONS.md, "The scripts"), with its own runtime for each run, as
' runMapScript makes a JSRuntime for each run.
' * The limited context (JS_NewLimitedContext: base objects and Map / Set only):
'   a JsRT context has every intrinsic, so the ones QuickJS leaves out (Date,
'   RegExp, JSON, Proxy, the typed arrays, Promise, WeakRef, ...) are deleted
'   from the global object first (LIMITED_CONTEXT_SOURCE), and globalThis (which
'   legacy Chakra lacks) defined. Without its eval intrinsic QuickJS throws a
'   TypeError "eval is not supported" from eval and the Function constructor:
'   here both are replaced by functions that throw it, and the runtime has eval
'   off (JsRuntimeAttributeDisableEval: the generator function constructor
'   would end the script uncaught). Legacy Chakra has no BigInt (QuickJS's base
'   objects include it).
' * The runtime keeps Chakra's background work (its JIT and GC threads; they run
'   no VB code: the native functions are called on this thread): without it
'   WaterLoop and Entropy took 5.5 s each in place of 0.4 s. (The exe's scripts
'   run with JsRuntimeAttributeDisableBackgroundWork.)
' * The runtime's limits: the 100 MiB memory limit is JsSetRuntimeMemoryLimit's;
'   QuickJS's 512 KiB stack limit has no JsRT setting (Chakra checks its own
'   stack); the interrupt handler (a script stopped after 30 s) is not ported:
'   JsRT stops a script only from another thread (JsDisableRuntimeExecution).
' * ScriptMapData (the context's opaque data) is this module's state for the
'   run in progress (one at a time): the Map, the logger, gameRand's
'   std::mt19937 (mt19937_seed / mt19937_next).
' * The four native functions have one AddressOf entry (runMap_nativeFunction,
'   the function's number as its callback state), which traps every VB error
'   and turns it into a JS exception and clears the x87's flags (jsrt_win's
'   rules). QuickJS pads argv to the function's length with undefined.
' * SCRIPT_ASSERT logs and returns null (Bellard's QuickJS raises only a
'   returned JS_EXCEPTION: the ReferenceError it makes is dropped);
'   SCRIPT_ASSERT_AND_RETURNERROR logs and raises the ReferenceError
'   "<expr> failed in <function> at line <line>" (scriptError).
' * JS_ToInt32 / JS_ToUint32 are ToNumber then modulo 2^32 (a failed
'   conversion gives 0, its exception dropped as QuickJS drops it when the C
'   function returns normally); uint32_t values the noise generator computes
'   are Doubles holding them exactly, wrapped modulo 2^32 where C wraps.
' * (VB) The x87's precision is set to 53 bits while the script runs (and
'   restored after). A VB process runs the x87 at 64 bits, and Chakra's 32-bit
'   code computes on it, so Math.sqrt double-rounded (1 ulp off in near-halfway
'   cases); QuickJS in the official x64 build computes in SSE2, correctly
'   rounded, as the x87 does at 53 bits (tools/map_script_test.py checks every
'   Math.sqrt / Math.pow the maps make against ucrtbase.dll). VB's runtime
'   reloads its own control word (its conversions: CLng, Fix, ...), so the
'   precision is set again wherever control goes back to the script: a native
'   function's return and a callback's call (jsEnterScript). msvcrt's
'   _controlfp is cdecl: a Declare would leave its arguments on the stack (a
'   small procedure calling it crashed), so it is called through DispCallFunc.
' * C++ would crash (an integer division by zero) where
'   generateFractalValueNoise's scale is 1, a layer's scale squared wraps to 0
'   (a scale of 2^17 or more) or the normalized noise is flat; here VB's error
'   11 becomes a JS exception.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function clearfp_ Lib "msvcrt" Alias "_clearfp" () As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, prgvt As Integer, prgpvarg As Long, pvargResult As Variant) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Const CC_CDECL As Long = 1
Private Const MCW_PC As Long = &H30000           ' float.h: the x87's precision control
Private Const PC_53 As Long = &H10000

' jsrt.h (the functions and values used here)
Private Const JsNoError As Long = 0
Private Const JsRuntimeAttributeDisableEval As Long = &H10
Private Const JsNumber As Long = 2
Private Const JsBoolean As Long = 4
Private Const JsObject As Long = 5
Private Const JsFunction As Long = 6
Private Const JsError As Long = 7
Private Const JsArray As Long = 8
Private Declare Function JsCreateRuntime Lib "chakra.dll" (ByVal attributes As Long, ByVal threadService As Long, runtime As Long) As Long
Private Declare Function JsDisposeRuntime Lib "chakra.dll" (ByVal runtime As Long) As Long
Private Declare Function JsSetRuntimeMemoryLimit Lib "chakra.dll" (ByVal runtime As Long, ByVal memoryLimit As Long) As Long
Private Declare Function JsCreateContext Lib "chakra.dll" (ByVal runtime As Long, newContext As Long) As Long
Private Declare Function JsGetCurrentContext Lib "chakra.dll" (currentContext As Long) As Long
Private Declare Function JsSetCurrentContext Lib "chakra.dll" (ByVal context As Long) As Long
Private Declare Function JsGetGlobalObject Lib "chakra.dll" (globalObject As Long) As Long
Private Declare Function JsGetUndefinedValue Lib "chakra.dll" (undefinedValue As Long) As Long
Private Declare Function JsGetNullValue Lib "chakra.dll" (nullValue As Long) As Long
Private Declare Function JsBoolToBoolean Lib "chakra.dll" (ByVal value As Long, booleanValue As Long) As Long
Private Declare Function JsBooleanToBool Lib "chakra.dll" (ByVal value As Long, boolValue As Long) As Long
Private Declare Function JsGetValueType Lib "chakra.dll" (ByVal value As Long, typ As Long) As Long
Private Declare Function JsDoubleToNumber Lib "chakra.dll" (ByVal dbl As Double, value As Long) As Long
Private Declare Function JsNumberToDouble Lib "chakra.dll" (ByVal value As Long, dbl As Double) As Long
Private Declare Function JsConvertValueToNumber Lib "chakra.dll" (ByVal value As Long, numberValue As Long) As Long
Private Declare Function JsPointerToString Lib "chakra.dll" (ByVal stringValue As Long, ByVal stringLength As Long, value As Long) As Long
Private Declare Function JsStringToPointer Lib "chakra.dll" (ByVal value As Long, stringValue As Long, stringLength As Long) As Long
Private Declare Function JsConvertValueToString Lib "chakra.dll" (ByVal value As Long, stringValue As Long) As Long
Private Declare Function JsConvertValueToObject Lib "chakra.dll" (ByVal value As Long, object As Long) As Long
Private Declare Function JsGetPropertyIdFromName Lib "chakra.dll" (ByVal name_ As Long, propertyId As Long) As Long
Private Declare Function JsGetProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, value As Long) As Long
Private Declare Function JsSetProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, ByVal value As Long, ByVal useStrictRules As Long) As Long
Private Declare Function JsDefineProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, ByVal propertyDescriptor As Long, result As Long) As Long
Private Declare Function JsCreateObject Lib "chakra.dll" (obj As Long) As Long
Private Declare Function JsCreateArray Lib "chakra.dll" (ByVal length As Long, result As Long) As Long
Private Declare Function JsGetIndexedProperty Lib "chakra.dll" (ByVal obj As Long, ByVal index As Long, result As Long) As Long
Private Declare Function JsSetIndexedProperty Lib "chakra.dll" (ByVal obj As Long, ByVal index As Long, ByVal value As Long) As Long
Private Declare Function JsCreateFunction Lib "chakra.dll" (ByVal nativeFunction As Long, ByVal callbackState As Long, func As Long) As Long
Private Declare Function JsCallFunction Lib "chakra.dll" (ByVal func As Long, ByVal arguments As Long, ByVal argumentCount As Long, result As Long) As Long
Private Declare Function JsCreateError Lib "chakra.dll" (ByVal message As Long, errorValue As Long) As Long
Private Declare Function JsCreateTypeError Lib "chakra.dll" (ByVal message As Long, errorValue As Long) As Long
Private Declare Function JsCreateReferenceError Lib "chakra.dll" (ByVal message As Long, errorValue As Long) As Long
Private Declare Function JsHasException Lib "chakra.dll" (hasException As Long) As Long
Private Declare Function JsGetAndClearException Lib "chakra.dll" (exception As Long) As Long
Private Declare Function JsSetException Lib "chakra.dll" (ByVal exception As Long) As Long
Private Declare Function JsParseScript Lib "chakra.dll" (ByVal script As Long, ByVal sourceContext As Long, ByVal sourceUrl As Long, result As Long) As Long
Private Declare Function JsAddRef Lib "chakra.dll" (ByVal ref As Long, count As Long) As Long
Private Declare Function JsRelease Lib "chakra.dll" (ByVal ref As Long, count As Long) As Long

Private Const MAX_PLAYERS As Long = 11          ' Maximum number of players in the game.
Private Const MAX_MAPSCRIPT_MEMORY As Long = 104857600     ' JS_SetMemoryLimit(rt, 100*1024*1024)
Private Const TWO_32_ As Double = 4294967296#

' (VB) the native functions' numbers (their callback state)
Private Const MAPFN_GAMERAND As Long = 19400
Private Const MAPFN_LOG As Long = 19401
Private Const MAPFN_SETMAPDATA As Long = 19402
Private Const MAPFN_NOISE As Long = 19403

' (VB) what JS_NewLimitedContext leaves out, deleted from a full JsRT context: the options runMapScript turns off
' (dateObject, eval, stringNormalize, regExp, json, proxy, typedArrays, promise, weakRef) and what QuickJS lacks
' (Intl, WebAssembly, JScript's own globals)
Private Const LIMITED_CONTEXT_SOURCE As String = _
    "(function (g) {" & vbLf & _
    "  ['Date', 'RegExp', 'JSON', 'Proxy', 'ArrayBuffer', 'SharedArrayBuffer', 'DataView', 'Int8Array', 'Uint8Array'," & vbLf & _
    "   'Uint8ClampedArray', 'Int16Array', 'Uint16Array', 'Int32Array', 'Uint32Array', 'Float32Array', 'Float64Array'," & vbLf & _
    "   'BigInt64Array', 'BigUint64Array', 'Atomics', 'Promise', 'WeakRef', 'FinalizationRegistry', 'Intl', 'WebAssembly'," & vbLf & _
    "   'CollectGarbage', 'ScriptEngine', 'ScriptEngineMajorVersion', 'ScriptEngineMinorVersion', 'ScriptEngineBuildVersion'," & vbLf & _
    "   'Debug'].forEach(function (n) { delete g[n]; });" & vbLf & _
    "  delete String.prototype.normalize;" & vbLf & _
    "  function def(o, name, v) { Object.defineProperty(o, name, { value: v, writable: true, configurable: true, enumerable: false }); }" & vbLf & _
    "  if (typeof globalThis === 'undefined') def(g, 'globalThis', g);" & vbLf & _
    "  def(g, 'eval', function eval(x) { throw new TypeError('eval is not supported'); });" & vbLf & _
    "  var NF = function Function() { throw new TypeError('eval is not supported'); };" & vbLf & _
    "  NF.prototype = Function.prototype;" & vbLf & _
    "  def(Function.prototype, 'constructor', NF);" & vbLf & _
    "  def(g, 'Function', NF);" & vbLf & _
    "})(this);"

' ScriptMapData (for the run in progress)
Private dataMap As Map                          ' std::shared_ptr<WzMap::Map> map
Private dataLogger As LoggingProtocol           ' WzMap::LoggingProtocol* pCustomLogger
Private dataActive As Boolean                   ' (VB) JS_GetContextOpaque(ctx) != nullptr

' std::mt19937
Private mtState(0 To 623) As Long
Private mtIndex As Long

Private sourceCookie As Long                    ' JsParseScript's source contexts
Private controlfpAddr As Long                   ' msvcrt's _controlfp

'--------------------------------------------------------------------------
' std::mt19937 (the 32-bit Mersenne Twister, default seeding)

Private Sub mt19937_seed(ByVal seed As Long)
    Dim i As Long, prev As Long
    mtState(0) = seed
    For i = 1 To 623
        prev = mtState(i - 1)
        mtState(i) = FromU32(U32(MUL32(1812433253, prev Xor SHRU(prev, 30))) + i)
    Next
    mtIndex = 624
End Sub

Private Function mt19937_next() As Long
    Dim i As Long, y As Long
    If mtIndex >= 624 Then
        For i = 0 To 623
            y = (mtState(i) And &H80000000) Or (mtState((i + 1) Mod 624) And &H7FFFFFFF)
            mtState(i) = mtState((i + 397) Mod 624) Xor SHRU(y, 1)
            If (y And 1) <> 0 Then mtState(i) = mtState(i) Xor &H9908B0DF
        Next
        mtIndex = 0
    End If
    y = mtState(mtIndex)
    mtIndex = mtIndex + 1
    y = y Xor SHRU(y, 11)
    y = y Xor (SHL(y, 7) And &H9D2C5680)
    y = y Xor (SHL(y, 15) And &HEFC60000)
    y = y Xor SHRU(y, 18)
    mt19937_next = y
End Function

'--------------------------------------------------------------------------
' (VB) JsRT helpers (QuickJS's API as map_script.cpp uses it; jsrt_win's in the exe)

Private Sub jsClearFPU()
    clearfp_
End Sub

' (VB) msvcrt's _controlfp: a cdecl function (a Declare would leave its arguments on the stack), called
' through DispCallFunc
Private Function controlfp_(ByVal newControl As Long, ByVal mask As Long) As Long
    Dim a(0 To 1) As Variant, t(0 To 1) As Integer, p(0 To 1) As Long, r As Variant
    If controlfpAddr = 0 Then controlfpAddr = GetProcAddress(LoadLibraryA("msvcrt"), "_controlfp")
    a(0) = newControl
    a(1) = mask
    t(0) = vbLong
    t(1) = vbLong
    p(0) = VarPtr(a(0))
    p(1) = VarPtr(a(1))
    DispCallFunc 0, controlfpAddr, CC_CDECL, vbLong, 2, t(0), p(0), r
    controlfp_ = r
End Function

' (VB) the x87 at 53 bits before control goes (back) to the script (see the header)
Private Sub jsEnterScript()
    controlfp_ PC_53, MCW_PC
End Sub

Private Function jsUndefined() As Long
    JsGetUndefinedValue jsUndefined
End Function

Private Function jsNull() As Long
    JsGetNullValue jsNull
End Function

Private Function jsBool(ByVal b As Boolean) As Long
    If b Then JsBoolToBoolean 1, jsBool Else JsBoolToBoolean 0, jsBool
End Function

' JS_NewUint32 / JS_NewInt32 (a Double holding the value)
Private Function jsNewNumber(ByVal d As Double) As Long
    JsDoubleToNumber d, jsNewNumber
End Function

Private Function jsString(ByVal s As String) As Long
    If Len(s) = 0 Then
        s = String$(1, 0)                           ' a non-null buffer (JsPointerToString refuses a null pointer)
        JsPointerToString StrPtr(s), 0, jsString
    Else
        JsPointerToString StrPtr(s), Len(s), jsString
    End If
End Function

Private Function jsType(ByVal v As Long) As Long
    jsType = -1
    If v = 0 Then Exit Function
    If JsGetValueType(v, jsType) <> JsNoError Then jsType = -1
End Function

Private Function jsPropId(ByVal name_ As String) As Long
    JsGetPropertyIdFromName StrPtr(name_), jsPropId
End Function

' Drops a pending exception
Private Sub jsClearException()
    Dim e As Long, has As Long
    JsHasException has
    If (has And &HFF&) <> 0 Then JsGetAndClearException e
End Sub

' JS_GetPropertyStr: a primitive is boxed (QuickJS reads its prototype's properties); undefined where it throws
Private Function jsGetProp(ByVal obj As Long, ByVal name_ As String) As Long
    Dim o As Long, v As Long
    Select Case jsType(obj)
    Case JsObject, JsFunction, JsError, JsArray
        o = obj
    Case Else
        If JsConvertValueToObject(obj, o) <> JsNoError Then
            jsClearException
            jsGetProp = jsUndefined()
            Exit Function
        End If
    End Select
    If JsGetProperty(o, jsPropId(name_), v) <> JsNoError Then
        jsClearException
        v = jsUndefined()
    End If
    jsGetProp = v
End Function

' JS_GetPropertyUint32 (undefined where it throws)
Private Function jsGetIndex(ByVal arr As Long, ByVal i As Double) As Long
    Dim v As Long
    If JsGetIndexedProperty(arr, jsNewNumber(i), v) <> JsNoError Then
        jsClearException
        v = jsUndefined()
    End If
    jsGetIndex = v
End Function

' JS ToInt32 of a number: NaN and the infinities 0, else truncated, modulo 2^32 (the uint32_t's bits)
Private Function dblToInt32(ByVal d As Double) As Long
    Dim hi As Long, m As Double
    CopyMemory hi, ByVal VarPtr(d) + 4, 4
    If (hi And &H7FF00000) = &H7FF00000 Then Exit Function
    m = Fix(d)
    If m >= -2147483648# And m <= 2147483647# Then
        dblToInt32 = CLng(m)
        Exit Function
    End If
    m = m - Int(m / TWO_32_) * TWO_32_
    If m >= 2147483648# Then m = m - TWO_32_
    dblToInt32 = CLng(m)
End Function

' ToNumber; ok False (and NaN's 0) where it throws (the exception dropped)
Private Function jsToNumber(ByVal v As Long, ok As Boolean) As Double
    Dim n As Long, d As Double
    If JsConvertValueToNumber(v, n) <> JsNoError Then
        jsClearException
        ok = False
        Exit Function
    End If
    JsNumberToDouble n, d
    jsClearFPU
    ok = True
    jsToNumber = d
End Function

' JSValueToInt32 / JSValueToUint32 (JS_ToInt32 / JS_ToUint32: the same bits; 0 where the conversion fails)
Private Function JSValueToInt32(ByVal v As Long) As Long
    Dim ok As Boolean, d As Double
    d = jsToNumber(v, ok)
    If ok Then JSValueToInt32 = dblToInt32(d)
End Function

' JS_ToInt64 for normalizeToRange's check: True with the value where the int64_t result lies in [0, 2^32)
' (truncation; past 2^63 the remainder modulo 2^64, its sign kept, as QuickJS computes it)
Private Function jsToInt64InUint32Range(ByVal v As Long, result As Double) As Boolean
    Dim ok As Boolean, d As Double, hi As Long, e As Long, m As Double
    d = jsToNumber(v, ok)
    If Not ok Then Exit Function                    ' the C++ checks JS_ToInt64's result first
    CopyMemory hi, ByVal VarPtr(d) + 4, 4
    e = (hi And &H7FF00000) \ &H100000
    If e <= 1023 + 62 Then
        m = Fix(d)
    ElseIf e <= 1023 + 62 + 53 Then
        m = Abs(d)
        m = m - Fix(m / (TWO_32_ * TWO_32_)) * (TWO_32_ * TWO_32_)      ' |d| modulo 2^64 (exact)
        If d < 0 And m <> 0 Then m = TWO_32_ * TWO_32_ - m              ' -ret as a uint64_t
        If m >= TWO_32_ * TWO_32_ / 2 Then m = -1                      ' an int64_t below 0
    Else
        m = 0                                       ' also NaN and the infinities
    End If
    If m >= 0 And m <= 4294967295# Then
        result = m
        jsToInt64InUint32Range = True
    End If
End Function

' JSValueToStdString (JS_ToCString: "" where the conversion fails)
Private Function JSValueToStdString(ByVal v As Long) As String
    Dim sref As Long, p As Long, n As Long, s As String
    If JsConvertValueToString(v, sref) <> JsNoError Then
        jsClearException
        Exit Function
    End If
    If JsStringToPointer(sref, p, n) <> JsNoError Then Exit Function
    If n > 0 Then
        s = String$(n, 0)
        CopyMemory ByVal StrPtr(s), ByVal p, n * 2
    End If
    JSValueToStdString = s
End Function

' QuickJS_GetArrayLength: False where arr is no array
Private Function QuickJS_GetArrayLength(ByVal arr As Long, length As Double) As Boolean
    Dim ok As Boolean
    length = 0
    If jsType(arr) <> JsArray Then Exit Function
    length = jsToNumber(jsGetProp(arr, "length"), ok)
    If Not ok Then length = 0: Exit Function
    QuickJS_GetArrayLength = True
End Function

' QuickJS_DumpObject / QuickJS_DumpError: the pending exception's text (and its stack), cleared
Private Function QuickJS_DumpError() As String
    Dim e As Long, result As String, stack As Long
    JsGetAndClearException e
    result = JSValueToStdString(e)
    If jsType(e) = JsError Then
        stack = jsGetProp(e, "stack")
        If jsType(stack) <> 0 Then result = result & vbLf & JSValueToStdString(stack)     ' JsUndefined = 0
    End If
    QuickJS_DumpError = result
End Function

' JS_ThrowReferenceError(context, "%s failed in %s at line %d", #expr, __FUNCTION__, __LINE__): the exception set,
' 0 returned (JS_EXCEPTION)
Private Function scriptError(ByVal expr As String, ByVal function_ As String, ByVal line_ As Long) As Long
    Dim e As Long
    jsClearException
    JsCreateReferenceError jsString(expr & " failed in " & function_ & " at line " & CStr(line_)), e
    JsSetException e
    scriptError = 0
End Function

' JS_ThrowInternalError / JS_ThrowTypeError (JsRT has no InternalError: an Error)
Private Sub throwError(ByVal message As String, ByVal typeError As Boolean)
    Dim e As Long
    jsClearException
    If typeError Then JsCreateTypeError jsString(message), e Else JsCreateError jsString(message), e
    JsSetException e
End Sub

' JS_DefinePropertyValueStr / a JSCFunctionListEntry: a data property with the given attributes
Private Sub defineValue(ByVal obj As Long, ByVal name_ As String, ByVal value As Long, ByVal writableConfigurable As Boolean)
    Dim desc As Long, r As Long
    JsCreateObject desc
    JsSetProperty desc, jsPropId("value"), value, 0
    JsSetProperty desc, jsPropId("writable"), jsBool(writableConfigurable), 0
    JsSetProperty desc, jsPropId("enumerable"), jsBool(False), 0
    JsSetProperty desc, jsPropId("configurable"), jsBool(writableConfigurable), 0
    JsDefineProperty obj, jsPropId(name_), desc, r
End Sub

'--------------------------------------------------------------------------
' The native functions

' (VB) the native functions' entry (JsNativeFunction): arguments(0) is this, then the arguments; callbackState the
' function's number. QuickJS's C functions are no constructors.
Public Function runMap_nativeFunction(ByVal callee As Long, ByVal isConstructCall As Long, ByVal arguments As Long, ByVal argumentCount As Long, ByVal callbackState As Long) As Long
    Dim argc As Long, n As Long, argv() As Long, thisVal As Long, r As Long, i As Long, length As Long
    Dim errNumber As Long, errDescription As String, has As Long
    On Error GoTo fail
    jsClearFPU
    argumentCount = argumentCount And &HFFFF&
    argc = argumentCount - 1
    If argc < 0 Then argc = 0
    If (isConstructCall And &HFF&) <> 0 Then
        Select Case callbackState
        Case MAPFN_GAMERAND: throwError "gameRand is not a constructor", True
        Case MAPFN_LOG: throwError "log is not a constructor", True
        Case MAPFN_SETMAPDATA: throwError "setMapData is not a constructor", True
        Case Else: throwError "generateFractalValueNoise is not a constructor", True
        End Select
        runMap_nativeFunction = 0
        jsEnterScript
        Exit Function
    End If
    Select Case callbackState
    Case MAPFN_LOG: length = 1
    Case MAPFN_SETMAPDATA: length = 7
    Case MAPFN_NOISE: length = 6
    End Select
    n = argc
    If n < length Then n = length
    ReDim argv(0 To n)
    If argumentCount > 0 Then CopyMemory thisVal, ByVal arguments, 4
    If argc > 0 Then CopyMemory argv(0), ByVal arguments + 4, argc * 4
    For i = argc To n
        argv(i) = jsUndefined()
    Next
    ' MAPSCRIPT_GET_CONTEXT_DATA()
    If Not dataActive Then
        throwError "Unable to acquire context information", False
        runMap_nativeFunction = 0
        jsEnterScript
        Exit Function
    End If

    Select Case callbackState
    Case MAPFN_GAMERAND: r = runMap_gameRand(argc, argv)
    Case MAPFN_LOG: r = runMap_log(argc, argv)
    Case MAPFN_SETMAPDATA: r = runMap_setMapData(argc, argv)
    Case MAPFN_NOISE: r = runMap_generateFractalValueNoise(thisVal, argc, argv)
    End Select

    If r <> 0 Then
        jsClearException                            ' a normal return: a leftover exception is dropped, as QuickJS never raises it
    Else
        JsHasException has                          ' a throw (JS_EXCEPTION): the exception set is raised; none set, undefined
        If (has And &HFF&) = 0 Then r = jsUndefined()
    End If
    runMap_nativeFunction = r
    jsEnterScript
    Exit Function
fail:
    errNumber = Err.Number
    errDescription = Err.Description
    printLog_ dataLogger, 0, LogLevel_Error, "runMap_nativeFunction", "VB error %d in native function %d: %s", errNumber, callbackState, errDescription
    throwError "VB error " & errNumber & ": " & errDescription, False
    runMap_nativeFunction = 0
    jsEnterScript
End Function

'MAP-- ## gameRand([mod])
'MAP--
'MAP-- Generates a random number in a "safe" way, that will ensure that randomly-generated
'MAP-- maps are the same for all connected players.
'MAP--
Private Function runMap_gameRand(ByVal argc As Long, argv() As Long) As Long
    Dim num As Double, mod_ As Double, result As Double
    num = U32(mt19937_next())
    If argc >= 1 Then mod_ = U32(JSValueToInt32(argv(0)))
    If mod_ <> 0 Then result = num - Fix(num / mod_) * mod_ Else result = num
    runMap_gameRand = jsNewNumber(result)
End Function

'MAP-- ## log(string)
'MAP--
'MAP-- Outputs to the game's debug log a string (or object that can be converted to string)
'MAP--
Private Function runMap_log(ByVal argc As Long, argv() As Long) As Long
    Dim pCustomLogger As LoggingProtocol, str_ As String
    Set pCustomLogger = dataLogger
    If Not (argc = 1) Then printLog_ pCustomLogger, 246, LogLevel_Error, "runMap_log", "Must have one parameter": runMap_log = jsNull(): Exit Function
    str_ = JSValueToStdString(argv(0))
    printLog_ pCustomLogger, 248, LogLevel_Info, "runMap_log", "game.js: ""%s""", str_
    runMap_log = jsUndefined()
End Function

' (VB) an object's name, position and direction (the part of setMapData's loops the three kinds share): False
' with the exception set (the C++ returns it) where the position is no pair; x and y the uint32_t bits
Private Function readObjectBase(ByVal obj As Long, name_ As String, x As Long, y As Long, ByVal kind As Long) As Boolean
    Dim position As Long, arrayLen As Double, bGotArrayLength As Boolean, pCustomLogger As LoggingProtocol
    Set pCustomLogger = dataLogger
    x = 0                                           ' WorldPos's default
    y = 0
    name_ = JSValueToStdString(jsGetProp(obj, "name"))
    position = jsGetProp(obj, "position")
    If jsType(position) = JsArray Then
        ' [x, y]
        bGotArrayLength = QuickJS_GetArrayLength(position, arrayLen)
        If Not (bGotArrayLength And (arrayLen = 2)) Then
            Select Case kind
            Case 0: printLog_ pCustomLogger, 333, LogLevel_Error, "runMap_setMapData", "position array length must equal 2; actual length is: %u", CDec(arrayLen): scriptError "bGotArrayLength && (arrayLen == 2)", "runMap_setMapData", 333
            Case 1: printLog_ pCustomLogger, 368, LogLevel_Error, "runMap_setMapData", "position array length must equal 2; actual length is: %u", CDec(arrayLen): scriptError "bGotArrayLength && (arrayLen == 2)", "runMap_setMapData", 368
            Case 2: printLog_ pCustomLogger, 400, LogLevel_Error, "runMap_setMapData", "position array length must equal 2; actual length is: %u", CDec(arrayLen): scriptError "bGotArrayLength && (arrayLen == 2)", "runMap_setMapData", 400
            End Select
            Exit Function
        End If
        x = JSValueToInt32(jsGetIndex(position, 0))
        y = JSValueToInt32(jsGetIndex(position, 1))
    End If
    readObjectBase = True
End Function

'MAP-- ## setMapData(mapWidth, mapHeight, texture, height, structures, droids, features)
'MAP--
'MAP-- Sets map data from that generated by the script.
'MAP--
'MAP-- (NOTE: Documentation incomplete. Look at the example mapScript maps: 6c-Entropy/game.js, 10c-WaterLoop/game.js)
'MAP--
'MAP-- `structures` is an array of structure data objects. Example:
'MAP--    {name: "A0CommandCentre", position: [x, y], direction: 0x4000*gameRand(4), modules: 0, player: 1}
'MAP-- `droids` is an array of droid data objects. Example:
'MAP--    {name: "ConstructionDroid", position: [x, y], direction: gameRand(0x10000), player: 1}
'MAP-- `features` is an array of feature data objects. Example:
'MAP--    {name: "Tree1", position: [x, y], direction: gameRand(0x10000)}
'MAP--
Private Function runMap_setMapData(ByVal argc As Long, argv() As Long) As Long
    Dim pCustomLogger As LoggingProtocol, jsVal_mapWidth As Long, jsVal_mapHeight As Long
    Dim texture As Long, height_ As Long, structures As Long, droids As Long, features As Long
    Dim mapData_ As MapData, N As Long, arrayLen As Double, bGotArrayLength As Boolean
    Dim textureUint32 As Double, tileHeightUint32 As Double, n_ As Long, textureVal As Long, heightVal As Long
    Dim structureCount As Long, droidCount As Long, featureCount As Long, i As Long, obj As Long
    Dim mapStructures As ObjectVector, mapDroids As ObjectVector, mapFeatures As ObjectVector
    Dim sd As Structure, dd As Droid, fd As Feature, direction As Long, modules As Double, player As Long
    Dim name_ As String, posX As Long, posY As Long
    Set pCustomLogger = dataLogger
    If Not (argc = 7) Then printLog_ pCustomLogger, 269, LogLevel_Error, "runMap_setMapData", "Must have 7 parameters": runMap_setMapData = scriptError("argc == 7", "runMap_setMapData", 269): Exit Function
    jsVal_mapWidth = argv(0)
    jsVal_mapHeight = argv(1)
    texture = argv(2)
    height_ = argv(3)
    structures = argv(4)
    droids = argv(5)
    features = argv(6)
    If Not (jsType(jsVal_mapWidth) = JsNumber) Then printLog_ pCustomLogger, 277, LogLevel_Error, "runMap_setMapData", "mapWidth must be number": runMap_setMapData = scriptError("JS_IsNumber(jsVal_mapWidth)", "runMap_setMapData", 277): Exit Function
    If Not (jsType(jsVal_mapHeight) = JsNumber) Then printLog_ pCustomLogger, 278, LogLevel_Error, "runMap_setMapData", "mapHeight must be number": runMap_setMapData = scriptError("JS_IsNumber(jsVal_mapHeight)", "runMap_setMapData", 278): Exit Function
    If Not (jsType(texture) = JsArray) Then printLog_ pCustomLogger, 279, LogLevel_Error, "runMap_setMapData", "texture must be array": runMap_setMapData = scriptError("WZ_QJS_IsArray(ctx, texture)", "runMap_setMapData", 279): Exit Function
    If Not (jsType(height_) = JsArray) Then printLog_ pCustomLogger, 280, LogLevel_Error, "runMap_setMapData", "height must be array": runMap_setMapData = scriptError("WZ_QJS_IsArray(ctx, height)", "runMap_setMapData", 280): Exit Function
    If Not (jsType(structures) = JsArray) Then printLog_ pCustomLogger, 281, LogLevel_Error, "runMap_setMapData", "structures must be array": runMap_setMapData = scriptError("WZ_QJS_IsArray(ctx, structures)", "runMap_setMapData", 281): Exit Function
    If Not (jsType(droids) = JsArray) Then printLog_ pCustomLogger, 282, LogLevel_Error, "runMap_setMapData", "droids must be array": runMap_setMapData = scriptError("WZ_QJS_IsArray(ctx, droids)", "runMap_setMapData", 282): Exit Function
    If Not (jsType(features) = JsArray) Then printLog_ pCustomLogger, 283, LogLevel_Error, "runMap_setMapData", "features must be array": runMap_setMapData = scriptError("WZ_QJS_IsArray(ctx, features)", "runMap_setMapData", 283): Exit Function
    Set mapData_ = dataMap.mapData()
    mapData_.width = JSValueToInt32(jsVal_mapWidth)
    mapData_.height = JSValueToInt32(jsVal_mapHeight)
    If Not (U32(mapData_.width) > 1 And U32(mapData_.height) > 1 And U32(mapData_.width) * U32(mapData_.height) <= MAP_MAXAREA) Then printLog_ pCustomLogger, 287, LogLevel_Error, "runMap_setMapData", "Map size out of bounds": runMap_setMapData = scriptError("mapData->width > 1 && mapData->height > 1 && (uint64_t)mapData->width*mapData->height <= MAP_MAXAREA", "runMap_setMapData", 287): Exit Function
    If Not (U32(mapData_.width) <= MAP_MAXWIDTH) Then printLog_ pCustomLogger, 288, LogLevel_Error, "runMap_setMapData", "Map width too large": runMap_setMapData = scriptError("mapData->width <= MAP_MAXWIDTH", "runMap_setMapData", 288): Exit Function
    If Not (U32(mapData_.height) <= MAP_MAXHEIGHT) Then printLog_ pCustomLogger, 289, LogLevel_Error, "runMap_setMapData", "Map height too large": runMap_setMapData = scriptError("mapData->height <= MAP_MAXHEIGHT", "runMap_setMapData", 289): Exit Function
    N = mapData_.width * mapData_.height
    mapData_.mMapTiles_resize N
    bGotArrayLength = QuickJS_GetArrayLength(texture, arrayLen)
    If Not (bGotArrayLength And (arrayLen = N)) Then printLog_ pCustomLogger, 294, LogLevel_Error, "runMap_setMapData", "texture array length must equal (mapWidth * mapHeight); actual length is: %u", CDec(arrayLen): runMap_setMapData = scriptError("bGotArrayLength && (arrayLen == ((uint64_t)mapData->width * (uint64_t)mapData->height))", "runMap_setMapData", 294): Exit Function
    bGotArrayLength = QuickJS_GetArrayLength(height_, arrayLen)
    If Not (bGotArrayLength And (arrayLen = N)) Then printLog_ pCustomLogger, 296, LogLevel_Error, "runMap_setMapData", "height array length must equal (mapWidth * mapHeight); actual length is: %u", CDec(arrayLen): runMap_setMapData = scriptError("bGotArrayLength && (arrayLen == ((uint64_t)mapData->width * (uint64_t)mapData->height))", "runMap_setMapData", 296): Exit Function
    For n_ = 0 To N - 1
        textureVal = jsGetIndex(texture, n_)
        heightVal = jsGetIndex(height_, n_)
        textureUint32 = U32(JSValueToInt32(textureVal))
        If Not (textureUint32 <= 65535) Then printLog_ pCustomLogger, 306, LogLevel_Error, "runMap_setMapData", "texture value exceeds uint16::max: %u", CDec(textureUint32): runMap_setMapData = scriptError("textureUint32 <= (uint32_t)std::numeric_limits<uint16_t>::max()", "runMap_setMapData", 306): Exit Function
        mapData_.mMapTiles_texture(n_) = CLng(textureUint32)
        tileHeightUint32 = U32(JSValueToInt32(heightVal))
        If tileHeightUint32 > TILE_MAX_HEIGHT Then
            ' treat as non-fatal error (to support older script maps) - log and cap at TILE_MAX_HEIGHT
            printLog_ pCustomLogger, 312, LogLevel_Error, "runMap_setMapData", "tile height (%u) exceeds TILE_MAX_HEIGHT (%u)", CDec(tileHeightUint32), CLng(TILE_MAX_HEIGHT)
            tileHeightUint32 = TILE_MAX_HEIGHT
        End If
        mapData_.mMapTiles_height(n_) = CLng(tileHeightUint32)
    Next
    bGotArrayLength = QuickJS_GetArrayLength(structures, arrayLen)
    If Not (bGotArrayLength And (arrayLen <= 65535)) Then printLog_ pCustomLogger, 318, LogLevel_Error, "runMap_setMapData", "structures array length must be <= uint16::max; actual length is: %u", CDec(arrayLen): runMap_setMapData = scriptError("bGotArrayLength && (arrayLen <= (uint64_t)std::numeric_limits<uint16_t>::max())", "runMap_setMapData", 318): Exit Function
    structureCount = CLng(arrayLen)
    Set mapStructures = dataMap.mapStructures()
    For i = 0 To structureCount - 1
        obj = jsGetIndex(structures, i)
        Set sd = New Structure
        If Not readObjectBase(obj, name_, posX, posY, 0) Then runMap_setMapData = 0: Exit Function
        sd.name_ = name_
        sd.position_x = posX
        sd.position_y = posY
        direction = JSValueToInt32(jsGetProp(obj, "direction"))
        If Not (direction >= 0 And direction <= 65535) Then printLog_ pCustomLogger, 342, LogLevel_Error, "runMap_setMapData", "Invalid direction (%d) for structure", direction: runMap_setMapData = scriptError("direction >= 0 && direction <= static_cast<int32_t>(std::numeric_limits<uint16_t>::max())", "runMap_setMapData", 342): Exit Function
        sd.direction = direction
        modules = U32(JSValueToInt32(jsGetProp(obj, "modules")))
        If Not (modules < 255) Then printLog_ pCustomLogger, 345, LogLevel_Error, "runMap_setMapData", "Invalid modules (%u) for structure", CDec(modules): runMap_setMapData = scriptError("modules < static_cast<uint32_t>(std::numeric_limits<uint8_t>::max())", "runMap_setMapData", 345): Exit Function
        sd.modules = CLng(modules)
        player = JSValueToInt32(jsGetProp(obj, "player"))
        ' (the C++ checks sd.player, still its default 0, before assigning it: always true)
        If Not (sd.player >= -1 And sd.player < MAX_PLAYERS) Then printLog_ pCustomLogger, 348, LogLevel_Error, "runMap_setMapData", "Invalid player (%d) for structure", sd.player: runMap_setMapData = scriptError("sd.player >= -1 && sd.player < MAX_PLAYERS", "runMap_setMapData", 348): Exit Function
        sd.player = ToInt8(player)
        mapStructures.push_back sd
    Next
    bGotArrayLength = QuickJS_GetArrayLength(droids, arrayLen)
    If Not (bGotArrayLength And (arrayLen <= 65535)) Then printLog_ pCustomLogger, 353, LogLevel_Error, "runMap_setMapData", "droids array length be <= uint16::max; actual length is: %u", CDec(arrayLen): runMap_setMapData = scriptError("bGotArrayLength && (arrayLen <= (uint64_t)std::numeric_limits<uint16_t>::max())", "runMap_setMapData", 353): Exit Function
    droidCount = CLng(arrayLen)
    Set mapDroids = dataMap.mapDroids()
    For i = 0 To droidCount - 1
        obj = jsGetIndex(droids, i)
        Set dd = New Droid
        If Not readObjectBase(obj, name_, posX, posY, 1) Then runMap_setMapData = 0: Exit Function
        dd.name_ = name_
        dd.position_x = posX
        dd.position_y = posY
        direction = JSValueToInt32(jsGetProp(obj, "direction"))
        If Not (direction >= 0 And direction <= 65535) Then printLog_ pCustomLogger, 377, LogLevel_Error, "runMap_setMapData", "Invalid direction (%d) for droid", direction: runMap_setMapData = scriptError("direction >= 0 && direction <= static_cast<int32_t>(std::numeric_limits<uint16_t>::max())", "runMap_setMapData", 377): Exit Function
        dd.direction = direction
        player = JSValueToInt32(jsGetProp(obj, "player"))
        If Not (dd.player >= -1 And dd.player < MAX_PLAYERS) Then printLog_ pCustomLogger, 380, LogLevel_Error, "runMap_setMapData", "Invalid player (%d) for droid", dd.player: runMap_setMapData = scriptError("sd.player >= -1 && sd.player < MAX_PLAYERS", "runMap_setMapData", 380): Exit Function
        dd.player = ToInt8(player)
        mapDroids.push_back dd
    Next
    bGotArrayLength = QuickJS_GetArrayLength(features, arrayLen)
    If Not (bGotArrayLength And (arrayLen <= 65535)) Then printLog_ pCustomLogger, 385, LogLevel_Error, "runMap_setMapData", "features array length be <= uint16::max; actual length is: %u", CDec(arrayLen): runMap_setMapData = scriptError("bGotArrayLength && (arrayLen <= (uint64_t)std::numeric_limits<uint16_t>::max())", "runMap_setMapData", 385): Exit Function
    featureCount = CLng(arrayLen)
    Set mapFeatures = dataMap.mapFeatures()
    For i = 0 To featureCount - 1
        obj = jsGetIndex(features, i)
        Set fd = New Feature
        If Not readObjectBase(obj, name_, posX, posY, 2) Then runMap_setMapData = 0: Exit Function
        fd.name_ = name_
        fd.position_x = posX
        fd.position_y = posY
        direction = JSValueToInt32(jsGetProp(obj, "direction"))
        If Not (direction >= 0 And direction <= 65535) Then printLog_ pCustomLogger, 409, LogLevel_Error, "runMap_setMapData", "Invalid direction (%d) for feature", direction: runMap_setMapData = scriptError("direction >= 0 && direction <= static_cast<int32_t>(std::numeric_limits<uint16_t>::max())", "runMap_setMapData", 409): Exit Function
        fd.direction = direction
        mapFeatures.push_back fd
    Next
    runMap_setMapData = jsBool(True)
End Function

'MAP-- ## generateFractalValueNoise(width, height, range, crispness, scale, normalizeToRange = 0,
'MAP--                              [[x1, y1, x2, y2, (x, y, layerIdx, layerRange) => value], ...] = [],
'MAP--                              rowMajorOrder = false)
'MAP--
'MAP-- A fast (native) fractal value noise generator, a cheaper version of Perlin noise.
'MAP-- (map_script.cpp's documentation in full: the C++ file.) The noise is returned in column-major order by
'MAP-- default; `range` is the first layer's range, `crispness` its decay in tenths, `scale` the first layer's
'MAP-- resolution, `normalizeToRange` scales the output to [0, normalizeToRange), the rigged regions' callbacks
'MAP-- give the layers' values inside them, `rowMajorOrder` returns the rows first.
Private Function runMap_generateFractalValueNoise(ByVal thisVal As Long, ByVal argc As Long, argv() As Long) As Long
    Dim pCustomLogger As LoggingProtocol
    Dim jsVal_width As Long, jsVal_height As Long, jsVal_range As Long, jsVal_crispness As Long, jsVal_scale As Long
    Dim width As Double, height As Double, range_ As Double, scale_ As Double, crispness As Double
    Dim normalizeToRange As Double, jsVal_riggedRegions As Long, regionsArrayLength As Double, numRiggedRegions As Long
    Dim jsVal_riggedRegion As Long, regionArrayLength As Double, i As Long, c As Long
    Dim jsVal_x1 As Long, jsVal_y1 As Long, jsVal_x2 As Long, jsVal_y2 As Long, jsVal_callback As Long
    Dim rX1() As Double, rY1() As Double, rX2() As Double, rY2() As Double, rCallback() As Long
    Dim returnRowMajorOrder As Boolean, b As Long, size As Long, maxLayerSize As Long
    Dim noiseData() As Double, layerData() As Double, noiseDataRowOrder() As Double
    Dim layerScale As Double, layerRange As Double, layerIdx As Double, layerWidth As Long, layerHeight As Long
    Dim layerScaleArea As Double, x As Long, y As Long, mapX As Double, mapY As Double, rigged As Boolean, r As Long
    Dim callArgv(0 To 4) As Long, jsVal_ret As Long, ret As Double, num As Double, a1 As Long, a2 As Long, a3 As Long, a4 As Long
    Dim layerDataMax As Double
    Dim tl As Double, tr As Double, bl As Double, br As Double, innerXEnd As Long, innerYEnd As Long
    Dim innerX As Long, innerY As Long, ls1 As Double, t As Double, k As Long, exact As Boolean
    Dim min_ As Double, max_ As Double, retVal As Long, ok As Boolean
    Set pCustomLogger = dataLogger

    If Not (argc >= 5 And argc <= 8) Then printLog_ pCustomLogger, 473, LogLevel_Error, "runMap_generateFractalValueNoise", "Must have 5 to 8 parameters": runMap_generateFractalValueNoise = scriptError("argc >= 5 && argc <= 8", "runMap_generateFractalValueNoise", 473): Exit Function
    jsVal_width = argv(0)
    jsVal_height = argv(1)
    jsVal_range = argv(2)
    jsVal_crispness = argv(3)
    jsVal_scale = argv(4)

    If Not (jsType(jsVal_width) = JsNumber) Then printLog_ pCustomLogger, 480, LogLevel_Error, "runMap_generateFractalValueNoise", "width must be number": runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_width)", "runMap_generateFractalValueNoise", 480): Exit Function
    If Not (jsType(jsVal_height) = JsNumber) Then printLog_ pCustomLogger, 481, LogLevel_Error, "runMap_generateFractalValueNoise", "height must be number": runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_height)", "runMap_generateFractalValueNoise", 481): Exit Function
    If Not (jsType(jsVal_range) = JsNumber) Then printLog_ pCustomLogger, 482, LogLevel_Error, "runMap_generateFractalValueNoise", "range must be number": runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_range)", "runMap_generateFractalValueNoise", 482): Exit Function
    If Not (jsType(jsVal_crispness) = JsNumber) Then printLog_ pCustomLogger, 483, LogLevel_Error, "runMap_generateFractalValueNoise", "crispness must be number": runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_crispness)", "runMap_generateFractalValueNoise", 483): Exit Function
    If Not (jsType(jsVal_scale) = JsNumber) Then printLog_ pCustomLogger, 484, LogLevel_Error, "runMap_generateFractalValueNoise", "scale must be number": runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_scale)", "runMap_generateFractalValueNoise", 484): Exit Function

    width = U32(JSValueToInt32(jsVal_width))
    height = U32(JSValueToInt32(jsVal_height))
    range_ = U32(JSValueToInt32(jsVal_range))
    scale_ = U32(JSValueToInt32(jsVal_scale))
    crispness = U32(JSValueToInt32(jsVal_crispness))

    If Not (width > 0) Then printLog_ pCustomLogger, 492, LogLevel_Error, "runMap_generateFractalValueNoise", "width must be > 0": runMap_generateFractalValueNoise = scriptError("width > 0", "runMap_generateFractalValueNoise", 492): Exit Function
    If Not (height > 0) Then printLog_ pCustomLogger, 493, LogLevel_Error, "runMap_generateFractalValueNoise", "height must be > 0": runMap_generateFractalValueNoise = scriptError("height > 0", "runMap_generateFractalValueNoise", 493): Exit Function
    If Not (range_ > 0) Then printLog_ pCustomLogger, 494, LogLevel_Error, "runMap_generateFractalValueNoise", "range must be > 0": runMap_generateFractalValueNoise = scriptError("range > 0", "runMap_generateFractalValueNoise", 494): Exit Function
    If Not (scale_ > 0) Then printLog_ pCustomLogger, 495, LogLevel_Error, "runMap_generateFractalValueNoise", "scale must be > 0": runMap_generateFractalValueNoise = scriptError("scale > 0", "runMap_generateFractalValueNoise", 495): Exit Function
    If Not (crispness > 0) Then printLog_ pCustomLogger, 496, LogLevel_Error, "runMap_generateFractalValueNoise", "crispness must be > 0": runMap_generateFractalValueNoise = scriptError("crispness > 0", "runMap_generateFractalValueNoise", 496): Exit Function

    If Not (width <= 256) Then printLog_ pCustomLogger, 498, LogLevel_Error, "runMap_generateFractalValueNoise", "width must be <= 256": runMap_generateFractalValueNoise = scriptError("width <= 256", "runMap_generateFractalValueNoise", 498): Exit Function
    If Not (height <= 256) Then printLog_ pCustomLogger, 499, LogLevel_Error, "runMap_generateFractalValueNoise", "height must be <= 256": runMap_generateFractalValueNoise = scriptError("height <= 256", "runMap_generateFractalValueNoise", 499): Exit Function

    normalizeToRange = 0
    If argc >= 6 Then
        If Not (jsType(argv(5)) = JsNumber) Then printLog_ pCustomLogger, 504, LogLevel_Error, "runMap_generateFractalValueNoise", "normalizeToRange must be number": runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_normalizeToRange)", "runMap_generateFractalValueNoise", 504): Exit Function
        ' (JS_ToInt64 of a number cannot fail: result == 0 always holds)
        If Not jsToInt64InUint32Range(argv(5), normalizeToRange) Then printLog_ pCustomLogger, 510, LogLevel_Error, "runMap_generateFractalValueNoise", "normalizeToRange must be positive or zero": runMap_generateFractalValueNoise = scriptError("normalizeToRange_int64 >= 0 && normalizeToRange_int64 <= UINT32_MAX", "runMap_generateFractalValueNoise", 510): Exit Function
    End If

    numRiggedRegions = 0
    If argc >= 7 Then
        ' Unwrap JS regions for faster lookup.
        jsVal_riggedRegions = argv(6)
        If Not (jsType(jsVal_riggedRegions) = JsArray) Then printLog_ pCustomLogger, 526, LogLevel_Error, "runMap_generateFractalValueNoise", "width must be array": runMap_generateFractalValueNoise = scriptError("WZ_QJS_IsArray(ctx, jsVal_riggedRegions)", "runMap_generateFractalValueNoise", 526): Exit Function

        If Not QuickJS_GetArrayLength(jsVal_riggedRegions, regionsArrayLength) Then printLog_ pCustomLogger, 530, LogLevel_Error, "runMap_generateFractalValueNoise", "Unable to retrieve riggedRegions array length": runMap_generateFractalValueNoise = scriptError("bGotRegionsArrayLength", "runMap_generateFractalValueNoise", 530): Exit Function
        If Not (regionsArrayLength <= 65535) Then printLog_ pCustomLogger, 532, LogLevel_Error, "runMap_generateFractalValueNoise", "Too many riggedRegions (%u)", CDec(regionsArrayLength): runMap_generateFractalValueNoise = scriptError("regionsArrayLength <= UINT16_MAX", "runMap_generateFractalValueNoise", 532): Exit Function

        If regionsArrayLength > 0 Then
            ReDim rX1(0 To regionsArrayLength - 1)
            ReDim rY1(0 To regionsArrayLength - 1)
            ReDim rX2(0 To regionsArrayLength - 1)
            ReDim rY2(0 To regionsArrayLength - 1)
            ReDim rCallback(0 To regionsArrayLength - 1)
        End If
        For i = 0 To regionsArrayLength - 1
            jsVal_riggedRegion = jsGetIndex(jsVal_riggedRegions, i)
            If Not (jsType(jsVal_riggedRegion) = JsArray) Then printLog_ pCustomLogger, 549, LogLevel_Error, "runMap_generateFractalValueNoise", "riggedRegion %u must be array", i: runMap_generateFractalValueNoise = scriptError("WZ_QJS_IsArray(ctx, jsVal_riggedRegion)", "runMap_generateFractalValueNoise", 549): GoTo freeRegions
            ok = QuickJS_GetArrayLength(jsVal_riggedRegion, regionArrayLength)
            If Not (ok And regionArrayLength = 5) Then printLog_ pCustomLogger, 553, LogLevel_Error, "runMap_generateFractalValueNoise", "riggedRegion[%u] length must be 5; actual length %u", i, CDec(regionArrayLength): runMap_generateFractalValueNoise = scriptError("bGotRegionArrayLength && regionArrayLength == 5", "runMap_generateFractalValueNoise", 553): GoTo freeRegions
            jsVal_x1 = jsGetIndex(jsVal_riggedRegion, 0)
            jsVal_y1 = jsGetIndex(jsVal_riggedRegion, 1)
            jsVal_x2 = jsGetIndex(jsVal_riggedRegion, 2)
            jsVal_y2 = jsGetIndex(jsVal_riggedRegion, 3)
            jsVal_callback = jsGetIndex(jsVal_riggedRegion, 4)

            If Not (jsType(jsVal_x1) = JsNumber) Then printLog_ pCustomLogger, 571, LogLevel_Error, "runMap_generateFractalValueNoise", "riggedRegion[%u][0] must be number", i: runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_x1)", "runMap_generateFractalValueNoise", 571): GoTo freeRegions
            If Not (jsType(jsVal_y1) = JsNumber) Then printLog_ pCustomLogger, 572, LogLevel_Error, "runMap_generateFractalValueNoise", "riggedRegion[%u][1] must be number", i: runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_y1)", "runMap_generateFractalValueNoise", 572): GoTo freeRegions
            If Not (jsType(jsVal_x2) = JsNumber) Then printLog_ pCustomLogger, 573, LogLevel_Error, "runMap_generateFractalValueNoise", "riggedRegion[%u][2] must be number", i: runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_x2)", "runMap_generateFractalValueNoise", 573): GoTo freeRegions
            If Not (jsType(jsVal_y2) = JsNumber) Then printLog_ pCustomLogger, 574, LogLevel_Error, "runMap_generateFractalValueNoise", "riggedRegion[%u][3] must be number", i: runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_y2)", "runMap_generateFractalValueNoise", 574): GoTo freeRegions
            If Not (jsType(jsVal_callback) = JsFunction) Then printLog_ pCustomLogger, 575, LogLevel_Error, "runMap_generateFractalValueNoise", "riggedRegion[%u][4] must be function", i: runMap_generateFractalValueNoise = scriptError("JS_IsFunction(ctx, jsVal_callback)", "runMap_generateFractalValueNoise", 575): GoTo freeRegions

            rX1(i) = U32(JSValueToInt32(jsVal_x1))
            rY1(i) = U32(JSValueToInt32(jsVal_y1))
            rX2(i) = U32(JSValueToInt32(jsVal_x2))
            rY2(i) = U32(JSValueToInt32(jsVal_y2))
            rCallback(i) = jsVal_callback                   ' JS_DupValue: JsAddRef'd while kept
            JsAddRef jsVal_callback, c
            numRiggedRegions = numRiggedRegions + 1
        Next
    End If

    returnRowMajorOrder = False
    If argc >= 8 Then
        If Not (jsType(argv(7)) = JsBoolean) Then printLog_ pCustomLogger, 592, LogLevel_Error, "runMap_generateFractalValueNoise", "rowMajorOrder must be bool": runMap_generateFractalValueNoise = scriptError("JS_IsBool(jsVal_rowMajorOrder)", "runMap_generateFractalValueNoise", 592): GoTo freeRegions
        JsBooleanToBool argv(7), b
        returnRowMajorOrder = ((b And &HFF&) <> 0)
    End If

    If Not (width * height <= 65535) Then printLog_ pCustomLogger, 599, LogLevel_Error, "runMap_generateFractalValueNoise", "Requested data too large to fit into Array": runMap_generateFractalValueNoise = scriptError("size <= UINT16_MAX", "runMap_generateFractalValueNoise", 599): GoTo freeRegions
    size = CLng(width * height)

    maxLayerSize = CLng((width + 1) * (height + 1))
    ' (the integer overflow check at line 605 holds for any width and height <= 256)

    ReDim noiseData(0 To size - 1)

    ' Allocate layer data once and reuse it for performance.
    ReDim layerData(0 To maxLayerSize - 1)

    layerScale = scale_
    layerRange = range_
    layerIdx = 4294967295#                              ' uint32_t layerIdx = -1

    Do
        layerScale = Fix(layerScale / 2)
        layerRange = Fix(U32(MUL32(FromU32(layerRange), FromU32(crispness))) / 10)
        If layerRange = 0 Then
            Exit Do
        End If
        layerIdx = layerIdx + 1
        If layerIdx = TWO_32_ Then layerIdx = 0

        layerWidth = CLng(Fix(width / layerScale)) + 1
        layerHeight = CLng(Fix(height / layerScale)) + 1

        layerScaleArea = U32(MUL32(FromU32(layerScale), FromU32(layerScale)))

        For x = 0 To layerWidth - 1
            For y = 0 To layerHeight - 1
                mapX = U32(MUL32(x, FromU32(layerScale)))
                mapY = U32(MUL32(y, FromU32(layerScale)))

                rigged = False
                For r = 0 To numRiggedRegions - 1
                    If mapX >= rX1(r) And mapY >= rY1(r) And mapX <= rX2(r) And mapY <= rY2(r) Then
                        rigged = True
                        a1 = jsNewNumber(mapX)
                        a2 = jsNewNumber(mapY)
                        a3 = jsNewNumber(layerIdx)
                        a4 = jsNewNumber(layerRange)
                        callArgv(0) = thisVal
                        callArgv(1) = a1
                        callArgv(2) = a2
                        callArgv(3) = a3
                        callArgv(4) = a4
                        jsVal_ret = 0
                        jsEnterScript
                        If JsCallFunction(rCallback(r), VarPtr(callArgv(0)), 5, jsVal_ret) <> JsNoError Then jsVal_ret = 0
                        jsClearFPU
                        If Not (jsType(jsVal_ret) = JsNumber) Then printLog_ pCustomLogger, 652, LogLevel_Error, "runMap_generateFractalValueNoise", "riggedRegion callback must return number": runMap_generateFractalValueNoise = scriptError("JS_IsNumber(jsVal_ret)", "runMap_generateFractalValueNoise", 652): GoTo freeRegions
                        ret = U32(JSValueToInt32(jsVal_ret))
                        If Not (ret < layerRange) Then printLog_ pCustomLogger, 655, LogLevel_Error, "runMap_generateFractalValueNoise", "riggedRegion callback must return non-negative number within range": runMap_generateFractalValueNoise = scriptError("ret < layerRange", "runMap_generateFractalValueNoise", 655): GoTo freeRegions
                        layerData(x * layerHeight + y) = ret
                        If ret > layerDataMax Then layerDataMax = ret
                    End If
                Next

                If Not rigged Then
                    ' This is equivalent to gameRand(layerRange).
                    num = U32(mt19937_next())
                    num = num - Fix(num / layerRange) * layerRange
                    layerData(x * layerHeight + y) = num
                    If num > layerDataMax Then layerDataMax = num
                End If
            Next
        Next

        ' (VB) the interpolation's sum is at most (the largest value in layerData, this layer's or an earlier's) *
        ' (layerScale - 1)^2: in Doubles when that stays exact, else each product modulo 2^32 as uint32_t computes it
        ls1 = layerScale - 1
        exact = (layerDataMax * ls1 * ls1 < 9007199254740992#)
        For x = 0 To layerWidth - 1
            mapX = U32(MUL32(x, FromU32(layerScale)))
            If mapX >= width Then
                GoTo nextX
            End If
            For y = 0 To layerHeight - 1
                mapY = U32(MUL32(y, FromU32(layerScale)))
                If mapY >= height Then
                    GoTo nextY
                End If

                tl = layerData(x * layerHeight + y)
                tr = layerData((x + 1) * layerHeight + y)
                bl = layerData(x * layerHeight + (y + 1))
                br = layerData((x + 1) * layerHeight + (y + 1))

                If layerScale < width - mapX Then innerXEnd = CLng(layerScale) Else innerXEnd = CLng(width - mapX)
                If layerScale < height - mapY Then innerYEnd = CLng(layerScale) Else innerYEnd = CLng(height - mapY)
                For innerX = 0 To innerXEnd - 1
                    For innerY = 0 To innerYEnd - 1
                        ' Interpolation.
                        If exact Then
                            t = br * innerX * innerY + bl * (ls1 - innerX) * innerY + tr * innerX * (ls1 - innerY) + _
                                tl * (ls1 - innerX) * (ls1 - innerY)
                            t = t - Fix(t / TWO_32_) * TWO_32_
                        Else
                            t = U32(MUL32(MUL32(FromU32(br), innerX), innerY)) + _
                                U32(MUL32(MUL32(FromU32(bl), FromU32(ls1 - innerX)), innerY)) + _
                                U32(MUL32(MUL32(FromU32(tr), innerX), FromU32(ls1 - innerY))) + _
                                U32(MUL32(MUL32(FromU32(tl), FromU32(ls1 - innerX)), FromU32(ls1 - innerY)))
                            t = t - Fix(t / TWO_32_) * TWO_32_
                        End If
                        k = CLng((mapX + innerX) * height + (mapY + innerY))
                        noiseData(k) = noiseData(k) + Fix(t / layerScaleArea)
                        If noiseData(k) >= TWO_32_ Then noiseData(k) = noiseData(k) - TWO_32_
                    Next
                Next
nextY:
            Next
nextX:
        Next
    Loop While layerScale > 1 And layerRange > 1

    ' Normalization.
    If normalizeToRange > 0 Then
        min_ = noiseData(0)
        max_ = noiseData(0)
        For k = 1 To size - 1
            If noiseData(k) < min_ Then min_ = noiseData(k)
            If noiseData(k) > max_ Then max_ = noiseData(k)
        Next
        For k = 0 To size - 1
            ' (val - min) * normalizeToRange / (max - min), in uint32_t
            noiseData(k) = Fix(U32(MUL32(FromU32(noiseData(k) - min_), FromU32(normalizeToRange))) / (max_ - min_))
        Next
    End If

    If returnRowMajorOrder Then
        ' convert noiseData from column-major order to row-major order
        ReDim noiseDataRowOrder(0 To size - 1)
        For x = 0 To width - 1
            For y = 0 To height - 1
                noiseDataRowOrder(y * width + x) = noiseData(x * height + y)
            Next
        Next
        noiseData = noiseDataRowOrder
        Erase noiseDataRowOrder
    End If

    JsCreateArray 0, retVal
    For k = 0 To size - 1
        JsSetIndexedProperty retVal, jsNewNumber(k), jsNewNumber(noiseData(k))
    Next
    runMap_generateFractalValueNoise = retVal
freeRegions:
    For r = 0 To numRiggedRegions - 1
        JsRelease rCallback(r), c
    Next
End Function

'--------------------------------------------------------------------------
' runMapScript

' std::shared_ptr<Map> runMapScript(const std::vector<char>& fileBuffer, const std::string &path, uint32_t seed, bool preview, LoggingProtocol* pCustomLogger)
' (fileBuffer(0 .. fileSize - 1), its last byte the null terminator mapIO's loadFullFile appended)
Public Function runMapScript(fileBuffer() As Byte, ByVal fileSize As Long, ByVal path As String, ByVal seed As Long, ByVal preview As Boolean, ByVal logger As LoggingProtocol) As Map
    Dim pCustomLogger As LoggingProtocol, rt As Long, ctx As Long, prevCtx As Long, global_obj As Long
    Dim limitSource As String, limitUrl As String, limitScript As Long, source As String, compiledScriptObj As Long
    Dim errorAsString As String, result As Long, args(0 To 0) As Long, fn As Long, ok As Boolean, oldControl As Long
    Set pCustomLogger = logger
    ' ScriptMapData data
    Set dataMap = New Map
    dataMap.construct_empty
    mt19937_seed seed
    Set dataLogger = logger

    If JsCreateRuntime(JsRuntimeAttributeDisableEval, 0, rt) <> JsNoError Then
        printLog_ pCustomLogger, 796, LogLevel_Error, "runMapScript", "JS_NewRuntime failed"
        GoTo done
    End If
    JsGetCurrentContext prevCtx
    If JsCreateContext(rt, ctx) <> JsNoError Then
        printLog_ pCustomLogger, 819, LogLevel_Error, "runMapScript", "JS_NewLimitedContext failed"
        GoTo freeRuntime
    End If
    JsSetCurrentContext ctx
    ' (VB) double precision on the x87 while the script runs (see the header)
    oldControl = controlfp_(0, 0)
    jsEnterScript
    ' (VB) the limited context: what QuickJS's leaves out deleted
    limitSource = LIMITED_CONTEXT_SOURCE
    limitUrl = "<limited context>"
    sourceCookie = sourceCookie + 1
    If JsParseScript(StrPtr(limitSource), sourceCookie, StrPtr(limitUrl), limitScript) = JsNoError Then
        args(0) = jsUndefined()
        jsEnterScript
        JsCallFunction limitScript, VarPtr(args(0)), 1, result
        jsClearFPU
    End If
    jsClearException
    JsGetGlobalObject global_obj

    If fileSize <= 0 Then
        printLog_ pCustomLogger, 828, LogLevel_Error, "runMapScript", "fileBuffer is empty"
        GoTo freeRuntime
    End If
    ' ensure fileBuffer is null-terminated, but ignore the null terminator when calculating the size
    If fileBuffer(fileSize - 1) <> 0 Then
        printLog_ pCustomLogger, 834, LogLevel_Error, "runMapScript", "fileBuffer must be null-terminated"
        GoTo freeRuntime
    End If
    source = utf8_decode(fileBuffer, 0, fileSize - 1)
    sourceCookie = sourceCookie + 1
    If JsParseScript(StrPtr(source), sourceCookie, StrPtr(path), compiledScriptObj) <> JsNoError Then
        ' compilation error / syntax error
        errorAsString = QuickJS_DumpError()
        printLog_ pCustomLogger, 843, LogLevel_Error, "runMapScript", "Syntax / compilation error in %s: %s", path, errorAsString
        GoTo freeRuntime
    End If

    defineValue global_obj, "preview", jsBool(preview), False
    defineValue global_obj, "XFLIP", jsNewNumber(TILE_XFLIP), False
    defineValue global_obj, "YFLIP", jsNewNumber(TILE_YFLIP), False
    defineValue global_obj, "ROTMASK", jsNewNumber(TILE_ROTMASK), False
    defineValue global_obj, "ROTSHIFT", jsNewNumber(TILE_ROTSHIFT), False
    defineValue global_obj, "TRIFLIP", jsNewNumber(TILE_TRIFLIP), False
    dataActive = True                                  ' JS_SetContextOpaque(ctx, &data)

    ' Special mapScript functions
    JsCreateFunction AddressOf runMap_nativeFunction, MAPFN_GAMERAND, fn
    defineValue global_obj, "gameRand", fn, True
    JsCreateFunction AddressOf runMap_nativeFunction, MAPFN_LOG, fn
    defineValue global_obj, "log", fn, True
    JsCreateFunction AddressOf runMap_nativeFunction, MAPFN_SETMAPDATA, fn
    defineValue global_obj, "setMapData", fn, True
    JsCreateFunction AddressOf runMap_nativeFunction, MAPFN_NOISE, fn
    defineValue global_obj, "generateFractalValueNoise", fn, True

    ' configure limitations on runtime
    JsSetRuntimeMemoryLimit rt, MAX_MAPSCRIPT_MEMORY       ' 100MiB better be enough...

    ' run the map script
    args(0) = jsUndefined()
    jsEnterScript
    ok = (JsCallFunction(compiledScriptObj, VarPtr(args(0)), 1, result) = JsNoError)
    jsClearFPU
    dataActive = False
    If Not ok Then
        ' compilation error / syntax error / runtime error
        errorAsString = QuickJS_DumpError()
        printLog_ pCustomLogger, 889, LogLevel_Error, "runMapScript", "Uncaught exception in %s: %s", path, errorAsString
        GoTo freeRuntime
    End If

    Set runMapScript = dataMap
freeRuntime:
    dataActive = False
    If ctx <> 0 Then controlfp_ oldControl, MCW_PC
    JsSetCurrentContext prevCtx
    JsDisposeRuntime rt
done:
    Set dataMap = Nothing
    Set dataLogger = Nothing
End Function
