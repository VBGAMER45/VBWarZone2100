Attribute VB_Name = "jsrt_win"
' Windows' JavaScript Runtime API (JsRT, the legacy Edge mode of
' SysWOW64\chakra.dll, jsrt.h): the declarations the script backend uses
' (chakra_backend_c, ChakraScriptingInstance) and thin helpers over them.
' It stands where quickjs_backend.cpp uses QuickJS's C API. VB-only.
' * A JsValueRef, JsContextRef, JsRuntimeHandle or JsPropertyIdRef is a Long.
'   Values on the stack (VB locals) are found by Chakra's conservative stack
'   scan; a value kept anywhere else (module or class variables, arrays) is
'   JsAddRef'd while it is kept (jsValueAddRef / jsValueRelease).
' * bool parameters and results: a C++ bool in a 4-byte stack slot (pass a
'   Long 0 or 1); bool * results are read into a Long cleared first.
' * The helpers work on the current context (JsSetCurrentContext): property
'   ids come from one cache per runtime (jsPropId), cleared when the runtime
'   is disposed (jsPropCacheClear).
' * QuickJS's property flags (JS_PROP_CONFIGURABLE, WRITABLE, ENUMERABLE)
'   are kept as numbers; jsDefineProp defines a property with them through a
'   property descriptor (one per flag combination and context, kept by the
'   instance: jsDescriptor), as JS_DefinePropertyValue does.
' * A failed conversion (a toString or valueOf that throws) leaves the
'   runtime in its exception state; the helpers clear it (jsClearException)
'   and return QuickJS's failure value, as QuickJS's pending exception is
'   dropped when the C function returns normally.
Option Explicit

Private Declare Function clearfp_ Lib "msvcrt" Alias "_clearfp" () As Long

' jsrt.h: JsErrorCode (the ones the port tests)
Public Const JsNoError As Long = 0
Public Const JsErrorInExceptionState As Long = &H10007
Public Const JsErrorScriptException As Long = &H30001
Public Const JsErrorScriptCompile As Long = &H30002

' jsrt.h: JsRuntimeAttributes
Public Const JsRuntimeAttributeNone As Long = 0
Public Const JsRuntimeAttributeDisableBackgroundWork As Long = &H1
Public Const JsRuntimeAttributeDisableEval As Long = &H10

' jsrt.h: JsValueType
Public Const JsUndefined As Long = 0
Public Const JsNull As Long = 1
Public Const JsNumber As Long = 2
Public Const JsString As Long = 3
Public Const JsBoolean As Long = 4
Public Const JsObject As Long = 5
Public Const JsFunction As Long = 6
Public Const JsError As Long = 7
Public Const JsArray As Long = 8
Public Const JsSymbol As Long = 9
Public Const JsArrayBuffer As Long = 10
Public Const JsTypedArray As Long = 11
Public Const JsDataView As Long = 12

' quickjs.h: the property flags (JS_DefinePropertyValue)
Public Const JS_PROP_CONFIGURABLE As Long = 1
Public Const JS_PROP_WRITABLE As Long = 2
Public Const JS_PROP_ENUMERABLE As Long = 4
Public Const JS_PROP_C_W_E As Long = 7
Public Const JS_PROP_THROW As Long = &H4000&

Public Declare Function JsCreateRuntime Lib "chakra.dll" (ByVal attributes As Long, ByVal threadService As Long, runtime As Long) As Long
Public Declare Function JsDisposeRuntime Lib "chakra.dll" (ByVal runtime As Long) As Long
Public Declare Function JsCollectGarbage Lib "chakra.dll" (ByVal runtime As Long) As Long
Public Declare Function JsAddRef Lib "chakra.dll" (ByVal ref As Long, count As Long) As Long
Public Declare Function JsRelease Lib "chakra.dll" (ByVal ref As Long, count As Long) As Long
Public Declare Function JsCreateContext Lib "chakra.dll" (ByVal runtime As Long, newContext As Long) As Long
Public Declare Function JsGetCurrentContext Lib "chakra.dll" (currentContext As Long) As Long
Public Declare Function JsSetCurrentContext Lib "chakra.dll" (ByVal context As Long) As Long
Public Declare Function JsGetGlobalObject Lib "chakra.dll" (globalObject As Long) As Long
Public Declare Function JsGetUndefinedValue Lib "chakra.dll" (undefinedValue As Long) As Long
Public Declare Function JsGetNullValue Lib "chakra.dll" (nullValue As Long) As Long
Public Declare Function JsGetTrueValue Lib "chakra.dll" (trueValue As Long) As Long
Public Declare Function JsGetFalseValue Lib "chakra.dll" (falseValue As Long) As Long
Public Declare Function JsBoolToBoolean Lib "chakra.dll" (ByVal value As Long, booleanValue As Long) As Long
Public Declare Function JsBooleanToBool Lib "chakra.dll" (ByVal value As Long, boolValue As Long) As Long
Public Declare Function JsConvertValueToBoolean Lib "chakra.dll" (ByVal value As Long, booleanValue As Long) As Long
Public Declare Function JsGetValueType Lib "chakra.dll" (ByVal value As Long, typ As Long) As Long
Public Declare Function JsDoubleToNumber Lib "chakra.dll" (ByVal dbl As Double, value As Long) As Long
Public Declare Function JsIntToNumber Lib "chakra.dll" (ByVal intValue As Long, value As Long) As Long
Public Declare Function JsNumberToDouble Lib "chakra.dll" (ByVal value As Long, dbl As Double) As Long
Public Declare Function JsConvertValueToNumber Lib "chakra.dll" (ByVal value As Long, numberValue As Long) As Long
Public Declare Function JsPointerToString Lib "chakra.dll" (ByVal stringValue As Long, ByVal stringLength As Long, value As Long) As Long
Public Declare Function JsStringToPointer Lib "chakra.dll" (ByVal value As Long, stringValue As Long, stringLength As Long) As Long
Public Declare Function JsConvertValueToString Lib "chakra.dll" (ByVal value As Long, stringValue As Long) As Long
Public Declare Function JsGetPropertyIdFromName Lib "chakra.dll" (ByVal name_ As Long, propertyId As Long) As Long
Public Declare Function JsConvertValueToObject Lib "chakra.dll" (ByVal value As Long, object As Long) As Long
Public Declare Function JsGetProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, value As Long) As Long
Public Declare Function JsSetProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, ByVal value As Long, ByVal useStrictRules As Long) As Long
Public Declare Function JsHasProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, hasProperty As Long) As Long
Public Declare Function JsDeleteProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, ByVal useStrictRules As Long, result As Long) As Long
Public Declare Function JsDefineProperty Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, ByVal propertyDescriptor As Long, result As Long) As Long
Public Declare Function JsGetOwnPropertyNames Lib "chakra.dll" (ByVal obj As Long, propertyNames As Long) As Long
Public Declare Function JsGetOwnPropertyDescriptor Lib "chakra.dll" (ByVal obj As Long, ByVal propertyId As Long, propertyDescriptor As Long) As Long
Public Declare Function JsCreateObject Lib "chakra.dll" (obj As Long) As Long
Public Declare Function JsCreateArray Lib "chakra.dll" (ByVal length As Long, result As Long) As Long
Public Declare Function JsGetIndexedProperty Lib "chakra.dll" (ByVal obj As Long, ByVal index As Long, result As Long) As Long
Public Declare Function JsSetIndexedProperty Lib "chakra.dll" (ByVal obj As Long, ByVal index As Long, ByVal value As Long) As Long
Public Declare Function JsStrictEquals Lib "chakra.dll" (ByVal object1 As Long, ByVal object2 As Long, result As Long) As Long
Public Declare Function JsCreateFunction Lib "chakra.dll" (ByVal nativeFunction As Long, ByVal callbackState As Long, func As Long) As Long
Public Declare Function JsCallFunction Lib "chakra.dll" (ByVal func As Long, ByVal arguments As Long, ByVal argumentCount As Long, result As Long) As Long
Public Declare Function JsCreateError Lib "chakra.dll" (ByVal message As Long, errorValue As Long) As Long
Public Declare Function JsCreateReferenceError Lib "chakra.dll" (ByVal message As Long, errorValue As Long) As Long
Public Declare Function JsHasException Lib "chakra.dll" (hasException As Long) As Long
Public Declare Function JsGetAndClearException Lib "chakra.dll" (exception As Long) As Long
Public Declare Function JsSetException Lib "chakra.dll" (ByVal exception As Long) As Long
Public Declare Function JsParseScript Lib "chakra.dll" (ByVal script As Long, ByVal sourceContext As Long, ByVal sourceUrl As Long, result As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private propCache As Object             ' Scripting.Dictionary: property name_ -> JsPropertyIdRef (JsAddRef'd), for the runtime
Private emptyText As String             ' a non-null buffer for the empty string (JsPointerToString refuses a null pointer)

' The property ids are per runtime: cleared (released) before the runtime is disposed.
Public Sub jsPropCacheClear()
    Dim k As Variant, c As Long
    If propCache Is Nothing Then Exit Sub
    For Each k In propCache.Keys
        JsRelease propCache(k), c
    Next
    Set propCache = Nothing
End Sub

' JS_NewAtom / a property's id by name_ (cached)
Public Function jsPropId(ByVal name_ As String) As Long
    Dim id As Long, c As Long
    If propCache Is Nothing Then
        Set propCache = CreateObject("Scripting.Dictionary")
        propCache.CompareMode = 0       ' BinaryCompare
    End If
    If propCache.Exists(name_) Then
        jsPropId = propCache(name_)
        Exit Function
    End If
    If Len(name_) = 0 Then
        If StrPtr(emptyText) = 0 Then emptyText = String$(1, 0)
        JsGetPropertyIdFromName StrPtr(emptyText), id       ' a buffer holding one null character: ""
    Else
        JsGetPropertyIdFromName StrPtr(name_), id
    End If
    If id <> 0 Then
        JsAddRef id, c
        propCache.Add name_, id
    End If
    jsPropId = id
End Function

Public Sub jsValueAddRef(ByVal ref As Long)
    Dim c As Long
    If ref <> 0 Then JsAddRef ref, c
End Sub

Public Sub jsValueRelease(ByVal ref As Long)
    Dim c As Long
    If ref <> 0 Then JsRelease ref, c
End Sub

' JS_UNDEFINED, JS_NULL, JS_TRUE / JS_FALSE (JS_NewBool)
Public Function jsUndefinedValue() As Long
    JsGetUndefinedValue jsUndefinedValue
End Function

Public Function jsNullValue() As Long
    JsGetNullValue jsNullValue
End Function

Public Function jsBool(ByVal b As Boolean) As Long
    If b Then JsGetTrueValue jsBool Else JsGetFalseValue jsBool
End Function

' JS_NewInt32
Public Function jsInt(ByVal i As Long) As Long
    JsIntToNumber i, jsInt
End Function

' JS_NewUint32: the uint32_t's bits as an unsigned number
Public Function jsUint(ByVal u As Long) As Long
    If u >= 0 Then
        JsIntToNumber u, jsUint
    Else
        JsDoubleToNumber CDbl(u) + 4294967296#, jsUint
    End If
End Function

' JS_NewFloat64
Public Function jsDouble(ByVal d As Double) As Long
    JsDoubleToNumber d, jsDouble
End Function

' JS_NewString / JS_NewStringLen
Public Function jsStringValue(ByVal s As String) As Long
    If Len(s) = 0 Then
        If StrPtr(emptyText) = 0 Then emptyText = String$(1, 0)
        JsPointerToString StrPtr(emptyText), 0, jsStringValue
    Else
        JsPointerToString StrPtr(s), Len(s), jsStringValue
    End If
End Function

' JS_NewObject, JS_NewArray
Public Function jsNewObject() As Long
    JsCreateObject jsNewObject
End Function

Public Function jsNewArray() As Long
    JsCreateArray 0, jsNewArray
End Function

Public Function jsValueType(ByVal v As Long) As Long
    If JsGetValueType(v, jsValueType) <> JsNoError Then jsValueType = -1
End Function

Public Function jsIsUndefined(ByVal v As Long) As Boolean
    jsIsUndefined = (jsValueType(v) = JsUndefined)
End Function

Public Function jsIsNull(ByVal v As Long) As Boolean
    jsIsNull = (jsValueType(v) = JsNull)
End Function

Public Function jsIsNumber(ByVal v As Long) As Boolean
    jsIsNumber = (jsValueType(v) = JsNumber)
End Function

Public Function jsIsString(ByVal v As Long) As Boolean
    jsIsString = (jsValueType(v) = JsString)
End Function

Public Function jsIsBool(ByVal v As Long) As Boolean
    jsIsBool = (jsValueType(v) = JsBoolean)
End Function

Public Function jsIsFunction(ByVal v As Long) As Boolean
    jsIsFunction = (jsValueType(v) = JsFunction)
End Function

Public Function jsIsArray(ByVal v As Long) As Boolean
    jsIsArray = (jsValueType(v) = JsArray)
End Function

Public Function jsIsError(ByVal v As Long) As Boolean
    jsIsError = (jsValueType(v) = JsError)
End Function

' JS_IsObject: any object (functions, arrays, errors... too)
Public Function jsIsObject(ByVal v As Long) As Boolean
    Select Case jsValueType(v)
    Case JsObject, JsFunction, JsError, JsArray, JsArrayBuffer, JsTypedArray, JsDataView
        jsIsObject = True
    End Select
End Function

' JS_IsConstructor: QuickJS knows a function's constructor bit; JsRT does not show it. A function
' with an own "prototype" property (a function declaration or expression, a class) is taken for a
' constructor; arrow functions, methods and the native functions have none.
Public Function jsIsConstructor(ByVal v As Long) As Boolean
    Dim has As Long
    If jsValueType(v) <> JsFunction Then Exit Function
    If JsGetOwnPropertyDescriptor(v, jsPropId("prototype"), has) <> JsNoError Then
        jsClearException
        Exit Function
    End If
    jsIsConstructor = Not jsIsUndefined(has)
End Function

' JS_SameValueZero for objects (the circular reference check): identity
Public Function jsSameObject(ByVal a As Long, ByVal b As Long) As Boolean
    Dim r As Long
    If JsStrictEquals(a, b, r) <> JsNoError Then
        jsClearException
        Exit Function
    End If
    jsSameObject = ((r And &HFF&) <> 0)
End Function

' Drops a pending exception (see the header)
Public Sub jsClearException()
    Dim has As Long, e As Long
    If JsHasException(has) <> JsNoError Then Exit Sub
    If (has And &HFF&) <> 0 Then JsGetAndClearException e
End Sub

Public Function jsExceptionPending() As Boolean
    Dim has As Long
    If JsHasException(has) <> JsNoError Then Exit Function
    jsExceptionPending = ((has And &HFF&) <> 0)
End Function

' JS_GetException: the pending exception, cleared (undefined if none)
Public Function jsTakeException() As Long
    Dim e As Long
    If JsGetAndClearException(e) <> JsNoError Then e = jsUndefinedValue()
    jsTakeException = e
End Function

' JS_ThrowReferenceError(ctx, "%s", message) (the message already formatted)
Public Sub jsThrowReferenceError(ByVal message As String)
    Dim e As Long
    jsClearException
    JsCreateReferenceError jsStringValue(message), e
    JsSetException e
End Sub

' A JS string's text (no conversion)
Private Function jsStringText(ByVal sref As Long) As String
    Dim p As Long, n As Long, s As String
    If JsStringToPointer(sref, p, n) <> JsNoError Then Exit Function
    If n > 0 Then
        s = String$(n, 0)
        CopyMemory ByVal StrPtr(s), ByVal p, n * 2
    Else
        s = ""
    End If
    jsStringText = s
End Function

' JS_ToCString: ToString of any value; ok False where QuickJS returns NULL (the conversion threw)
Public Function jsToStringOk(ByVal v As Long, ok As Boolean) As String
    Dim sref As Long
    If JsConvertValueToString(v, sref) <> JsNoError Then
        jsClearException
        ok = False
        Exit Function
    End If
    ok = True
    jsToStringOk = jsStringText(sref)
End Function

' JSValueToStdString: "" where the conversion fails
Public Function jsToString(ByVal v As Long) As String
    Dim ok As Boolean
    jsToString = jsToStringOk(v, ok)
End Function

' JS_ToFloat64: ok False (and NaN) where ToNumber throws
Public Function jsToDoubleOk(ByVal v As Long, ok As Boolean) As Double
    Dim n As Long, d As Double
    If JsConvertValueToNumber(v, n) <> JsNoError Then
        jsClearException
        ok = False
        jsToDoubleOk = NaN_()
        Exit Function
    End If
    JsNumberToDouble n, d
    ok = True
    jsToDoubleOk = d
End Function

' JS ToInt32 of a number (NaN and infinities are 0; else truncated, modulo 2^32)
Public Function jsDoubleToInt32(ByVal d As Double) As Long
    Dim m As Double
    If dblNotFinite(d) Then Exit Function               ' NaN, +-Infinity (tested by bits: see float_vb's IsNaNF)
    m = Fix(d)
    If m >= -2147483648# And m <= 2147483647# Then
        jsDoubleToInt32 = CLng(m)
        Exit Function
    End If
    m = m - Int(m / 4294967296#) * 4294967296#          ' 0 <= m < 2^32
    If m >= 2147483648# Then m = m - 4294967296#
    jsDoubleToInt32 = CLng(m)
End Function

' JS_ToInt32 (JS_ToUint32 is the same bits): ok False (and 0) where ToNumber throws
Public Function jsToInt32Ok(ByVal v As Long, ok As Boolean) As Long
    Dim d As Double
    d = jsToDoubleOk(v, ok)
    If Not ok Then Exit Function
    jsToInt32Ok = jsDoubleToInt32(d)
End Function

' JS_ToBool
Public Function jsToBool(ByVal v As Long) As Boolean
    Dim b As Long, r As Long
    If JsConvertValueToBoolean(v, b) <> JsNoError Then
        jsClearException
        Exit Function
    End If
    JsBooleanToBool b, r
    jsToBool = ((r And &HFF&) <> 0)
End Function

' A number's value (no conversion)
Public Function jsNumberValue(ByVal v As Long) As Double
    JsNumberToDouble v, jsNumberValue
End Function

' (VB) a number, string or boolean as its wrapper object, as JS reads a primitive's property (QuickJS's
' JS_GetPropertyStr boxes it; JsRT's JsGetProperty refuses a non-object); anything else as it is
Private Function jsBoxPrimitive(ByVal v As Long) As Long
    Dim o As Long
    Select Case jsValueType(v)
    Case JsNumber, JsString, JsBoolean
        If JsConvertValueToObject(v, o) = JsNoError Then jsBoxPrimitive = o Else jsBoxPrimitive = v
    Case Else
        jsBoxPrimitive = v
    End Select
End Function

' JS_GetPropertyStr (undefined where the getter throws: the exception dropped)
Public Function jsGetProp(ByVal obj As Long, ByVal name_ As String) As Long
    Dim v As Long
    obj = jsBoxPrimitive(obj)
    If JsGetProperty(obj, jsPropId(name_), v) <> JsNoError Then
        jsClearException
        v = jsUndefinedValue()
    End If
    jsGetProp = v
End Function

' JS_GetProperty by id; False where it threw
Public Function jsGetPropOk(ByVal obj As Long, ByVal name_ As String, v As Long) As Boolean
    obj = jsBoxPrimitive(obj)
    If JsGetProperty(obj, jsPropId(name_), v) <> JsNoError Then
        jsClearException
        v = jsUndefinedValue()
        Exit Function
    End If
    jsGetPropOk = True
End Function

' JS_SetPropertyStr (non-strict: True if no exception)
Public Function jsSetProp(ByVal obj As Long, ByVal name_ As String, ByVal value As Long) As Boolean
    If JsSetProperty(obj, jsPropId(name_), value, 0) <> JsNoError Then
        jsClearException
        Exit Function
    End If
    jsSetProp = True
End Function

' JS_DefinePropertyValue(Str): 1 defined, 0 refused, -1 threw (JS_PROP_THROW: the exception is
' left pending for the caller, as QuickJS does)
Public Function jsDefineProp(ByVal obj As Long, ByVal name_ As String, ByVal value As Long, ByVal flags As Long) As Long
    Dim desc As Long, r As Long, err_ As Long
    desc = jsDescriptor(flags And JS_PROP_C_W_E)
    If JsSetProperty(desc, jsPropId("value"), value, 0) <> JsNoError Then
        jsClearException
        jsDefineProp = -1
        Exit Function
    End If
    err_ = JsDefineProperty(obj, jsPropId(name_), desc, r)
    If err_ <> JsNoError Then
        If (flags And JS_PROP_THROW) = 0 Then jsClearException
        jsDefineProp = -1
        Exit Function
    End If
    If (r And &HFF&) <> 0 Then
        jsDefineProp = 1
    ElseIf (flags And JS_PROP_THROW) <> 0 Then
        JsSetException jsTypeErrorValue("Cannot redefine property: " & name_)
        jsDefineProp = -1
    End If
End Function

' The property descriptor for a combination of QuickJS's flags, in the current context (made once
' per context by its instance)
Private Function jsDescriptor(ByVal flags As Long) As Long
    jsDescriptor = chakraDescriptor(flags)
End Function

Private Function jsTypeErrorValue(ByVal message As String) As Long
    Dim e As Long
    JsCreateError jsStringValue(message), e
    jsTypeErrorValue = e
End Function

' JS_DefinePropertyValueUint32(ctx, arr, i, v, JS_PROP_C_W_E) on a new array: an indexed set
Public Function jsSetIndex(ByVal arr As Long, ByVal i As Long, ByVal value As Long) As Boolean
    If JsSetIndexedProperty(arr, jsInt(i), value) <> JsNoError Then
        jsClearException
        Exit Function
    End If
    jsSetIndex = True
End Function

' JS_GetPropertyUint32
Public Function jsGetIndex(ByVal arr As Long, ByVal i As Long) As Long
    Dim v As Long
    If JsGetIndexedProperty(arr, jsInt(i), v) <> JsNoError Then
        jsClearException
        v = jsUndefinedValue()
    End If
    jsGetIndex = v
End Function

' JS_DeletePropertyStr: 1 deleted, 0 not, -1 threw
Public Function jsDeleteProp(ByVal obj As Long, ByVal name_ As String) As Long
    Dim r As Long
    If JsDeleteProperty(obj, jsPropId(name_), 0, r) <> JsNoError Then
        jsClearException
        jsDeleteProp = -1
        Exit Function
    End If
    If jsToBool(r) Then jsDeleteProp = 1
End Function

' The own property names of an object (JsGetOwnPropertyNames: strings only, in [[OwnPropertyKeys]]
' order), with enumerableOnly the enumerable ones; False where it failed
Public Function jsOwnPropertyNames(ByVal obj As Long, ByVal enumerableOnly As Boolean, names As StringVector) As Boolean
    Dim arr As Long, n As Long, i As Long, name_ As String, desc As Long, ok As Boolean
    StringVector_clear names
    If JsGetOwnPropertyNames(obj, arr) <> JsNoError Then
        jsClearException
        Exit Function
    End If
    n = jsToInt32Ok(jsGetProp(arr, "length"), ok)
    For i = 0 To n - 1
        name_ = jsToString(jsGetIndex(arr, i))
        If enumerableOnly Then
            If JsGetOwnPropertyDescriptor(obj, jsPropId(name_), desc) = JsNoError Then
                If Not jsIsUndefined(desc) Then
                    If jsToBool(jsGetProp(desc, "enumerable")) Then StringVector_push_back names, name_
                End If
            Else
                jsClearException
            End If
        Else
            StringVector_push_back names, name_
        End If
    Next
    jsOwnPropertyNames = True
End Function

' QuickJS_GetArrayLength: False if not an array
Public Function jsArrayLength(ByVal arr As Long, length As Long) As Boolean
    Dim ok As Boolean
    length = 0
    If Not jsIsArray(arr) Then Exit Function
    length = jsToInt32Ok(jsGetProp(arr, "length"), ok)
    jsArrayLength = ok
End Function

' JS_Call(ctx, fn, JS_UNDEFINED, n, args): the error code; args(0 .. n - 1)
Public Function jsCall(ByVal fn As Long, args() As Long, ByVal n As Long, result As Long) As Long
    Dim callArgs() As Long, i As Long
    ReDim callArgs(0 To n)
    callArgs(0) = jsUndefinedValue()             ' this
    For i = 0 To n - 1
        callArgs(i + 1) = args(i)
    Next
    jsCall = JsCallFunction(fn, VarPtr(callArgs(0)), n + 1, result)
    jsClearFPU                                          ' the script's NaNs must not reach VB's next conversion
End Function

' (VB) whether d is a NaN or an infinity, from its exponent bits (no comparison with a NaN)
Public Function dblNotFinite(ByVal d As Double) As Boolean
    Dim hi As Long
    CopyMemory hi, ByVal VarPtr(d) + 4, 4
    dblNotFinite = ((hi And &H7FF00000) = &H7FF00000)
End Function

' (VB) clears the x87's sticky exception flags. Chakra runs the scripts' floating point on this thread: a
' NaN a script makes (0 / 0, Math.sqrt(-1), ...) leaves the "invalid" flag set, which VB's next float to
' integer conversion anywhere reports as an overflow (error 6). Cleared where control passes from the
' scripts to VB: after every call into Chakra and on entry to every native function.
Public Sub jsClearFPU()
    clearfp_
End Sub
