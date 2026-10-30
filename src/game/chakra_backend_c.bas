Attribute VB_Name = "chakra_backend_c"
' src/quickjs_backend.cpp on Windows' Chakra (JsRT, jsrt_win) instead of
' QuickJS: everything but the API functions' wrappers and registerFunctions
' (lines 3140-3694: chakra_functions_c) and the class quickjs_scripting_instance
' (ChakraScriptingInstance). quickjs_execution_context is the class
' wzapi_execution_context, the timer lambdas ChakraTimerFunc, the
' quickjs_timer_additionaldata ChakraTimerData.
'
' Runtime and contexts: one JsRT runtime shared by every script instance,
' each instance with its own context (QuickJS: a runtime and a context per
' instance). Events cross instances while a script runs (a native function
' of one instance triggers another instance's event), and JsRT switches
' contexts of one runtime in the middle of a script, but not runtimes. The
' runtime is made with the first instance and disposed with the last; it
' disables eval (and new Function) unless the game is a campaign, as QuickJS's
' limited context does per instance (ctxOptions.eval). The other limits of
' that context: no Promise (deleted from each new context's global object);
' BigInt and WeakRef do not exist in legacy Chakra.
'
' The map engineToInstanceMap (JSContext * -> instance) holds weak
' references (ObjPtr), so that an instance terminates when the engine lets it
' go, as the C++ deletes it.
'
' Exceptions, as QuickJS behaves (quickjs-wz, Bellard's engine):
' * A native function that returns JS_EXCEPTION throws. Here: it calls
'   jsThrowReferenceError (or leaves a script's exception pending) and returns
'   0; jsNativeFunction returns that 0 and Chakra throws.
' * JS_ThrowReferenceError followed by a normal return (QuickJS's own
'   SCRIPT_ASSERT, UNBOX_SCRIPT_ASSERT, wzapi's throwError, js_include's
'   failure) leaves a pending exception that QuickJS never raises: OP_call
'   only checks the returned value. So those are logged and dropped, and
'   jsNativeFunction clears any exception left over by a normal return.
' * A VB error inside a native function is trapped by jsNativeFunction and
'   becomes a JS exception (a VB error never unwinds through Chakra).
'
' The lambdas log under their enclosing function's name (as netreplay_c's); the
' members of the anonymous namespace's structs (unbox<T>::operator()) and the
' box overloads log as GCC's __FUNCTION__ names them ("operator()", "box").
'
' Values: see jsrt_win. A JsValueRef kept in a VB array (the arguments of an
' event or a timer) is JsAddRef'd while it is there (jsArgsPush,
' jsArgsRelease), as the C++ std::vector<JSValue> owns its values.
'
' Semantics that may differ from QuickJS (where the game's output does not
' depend on them so far): the int / float distinction of numbers saved to
' JSON (QuickJS's tags; here an integral number in int32 range is an int);
' JS_IsConstructor (a function with an own "prototype"); functions'
' "length" and "name" properties (JsCreateFunction makes neither);
' symbol-keyed properties are not enumerated (JsGetOwnPropertyNames);
' error texts and stacks of script exceptions; Array.prototype.sort's order
' of equal elements (legacy Chakra's sort vs QuickJS's); the QuickJS debugger
' builtins (debugGetCallerFuncObject, debugGetCallerFuncName,
' debugGetBacktrace) are JS functions using Function.prototype.caller, which
' strict-mode callers do not have.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' quickjs_scripting_instance::LoadFileSearchOptions (wzapi.h)
Public Const LoadFileSearchOptions_ScriptPath_FileNameOnlyBackwardsCompat As Long = &H1
Public Const LoadFileSearchOptions_ScriptPath As Long = &H2
Public Const LoadFileSearchOptions_DataDir As Long = &H4
Public Const LoadFileSearchOptions_ConfigScriptDir As Long = &H8
Public Const LoadFileSearchOptions_All As Long = LoadFileSearchOptions_ScriptPath Or LoadFileSearchOptions_DataDir Or LoadFileSearchOptions_ConfigScriptDir
Public Const LoadFileSearchOptions_All_BackwardsCompat As Long = LoadFileSearchOptions_ScriptPath_FileNameOnlyBackwardsCompat Or LoadFileSearchOptions_All

' wzapi::scripting_instance::DebugSpecialStringType
Public Const DebugSpecialStringType_TYPE_DESCRIPTION As Long = 0

' transporter.h

' (VB) the native functions' numbers (jsNativeFunction's callbackState): the builtins of
' js_builtin_funcs here; from JSFN_API_FIRST on, chakra_functions_c's (round 2)
Public Const JSFN_SETTIMER As Long = 1
Public Const JSFN_QUEUE As Long = 2
Public Const JSFN_REMOVETIMER As Long = 3
Public Const JSFN_PROFILE As Long = 4
Public Const JSFN_INCLUDE As Long = 5
Public Const JSFN_INCLUDEJSON As Long = 6
Public Const JSFN_NAMESPACE As Long = 7
Public Const JSFN_API_FIRST As Long = 100

' struct JSToJsonContext
Private Type JSToJsonContext
    stack() As Long                     ' std::vector<JSValue>
    stack_size As Long
    skip_constructors As Boolean
End Type

Private engineToInstanceMap As Object   ' Scripting.Dictionary: JsContextRef -> ObjPtr(instance) (weak)
Private chakraRuntime As Long           ' the JsRuntimeHandle (0: none)
Private chakraRuntimeUsers As Long      ' the instances using it
Private ctxDescriptors As Object        ' Scripting.Dictionary: JsContextRef -> its 8 property descriptors (by flags)
Private descCtx As Long                 ' (cache) the context of descArr
Private descArr(0 To 7) As Long
Private fnLengths() As Long             ' the native functions' "length" (QuickJS pads argv to it with undefined)
Private fnLengths_size As Long
Private sourceCookie As Long            ' JsParseScript's source contexts

'--------------------------------------------------------------------------
' (VB) the runtime, the contexts and the instances

' debugGetCallerFuncObject, debugGetCallerFuncName, debugGetBacktrace (quickjs-debugger's
' js_debugger_get_caller_funcObject, _get_caller_name, _build_backtrace): JsRT cannot walk the script
' stack from a native function, so they are JS functions (non-strict, for Function.prototype.caller),
' defined as js_builtin_funcs defines its entries (writable, configurable, not enumerable). The caller
' is the function that called the function calling them. (VB) It also gives Object.hasOwn (ES2022, which
' QuickJS has and legacy Chakra lacks: libcampaign's artifact.js calls it).
Private Const DEBUG_BUILTINS_SOURCE As String = _
    "(function (g) {" & vbLf & _
    "  function def(name, fn) { Object.defineProperty(g, name, { value: fn, writable: true, configurable: true, enumerable: false }); }" & vbLf & _
    "  def('debugGetCallerFuncObject', function debugGetCallerFuncObject() { var f = debugGetCallerFuncObject.caller; return f ? f.caller : undefined; });" & vbLf & _
    "  def('debugGetCallerFuncName', function debugGetCallerFuncName() { var f = debugGetCallerFuncName.caller; f = f ? f.caller : null; return (f && f.name) ? f.name : '<anonymous>'; });" & vbLf & _
    "  def('debugGetBacktrace', function debugGetBacktrace() { try { throw new Error(); } catch (e) { return String(e.stack); } });" & vbLf & _
    "  if (!Object.hasOwn) Object.defineProperty(Object, 'hasOwn', { value: function hasOwn(o, p) { return Object.prototype.hasOwnProperty.call(Object(o), p); }, writable: true, configurable: true, enumerable: false });" & vbLf & _
    "})(this);"

' The shared runtime, made by the first instance (JS_NewRuntime)
Public Function chakraRuntimeAcquire() As Long
    Dim attributes As Long
    If chakraRuntime = 0 Then
        attributes = JsRuntimeAttributeDisableBackgroundWork
        If game.typ <> LEVEL_TYPE_CAMPAIGN Then
            attributes = attributes Or JsRuntimeAttributeDisableEval   ' allow "eval" only for campaign (which currently has lots of implicit eval usage)
        End If
        If JsCreateRuntime(attributes, 0, chakraRuntime) <> JsNoError Then chakraRuntime = 0
    End If
    If chakraRuntime <> 0 Then chakraRuntimeUsers = chakraRuntimeUsers + 1
    chakraRuntimeAcquire = chakraRuntime
End Function

' Disposed with its last instance (JS_FreeRuntime)
Public Sub chakraRuntimeRelease()
    chakraRuntimeUsers = chakraRuntimeUsers - 1
    If chakraRuntimeUsers > 0 Then Exit Sub
    chakraRuntimeUsers = 0
    JsSetCurrentContext 0
    jsPropCacheClear
    descCtx = 0
    If chakraRuntime <> 0 Then JsDisposeRuntime chakraRuntime
    chakraRuntime = 0
End Sub

' Makes ctx current; returns the context that was (for jsLeaveContext)
Public Function jsEnterContext(ByVal ctx As Long) As Long
    Dim prev As Long
    JsGetCurrentContext prev
    If prev <> ctx Then JsSetCurrentContext ctx
    jsEnterContext = prev
End Function

Public Sub jsLeaveContext(ByVal prev As Long)
    Dim cur As Long
    JsGetCurrentContext cur
    If cur <> prev Then JsSetCurrentContext prev
End Sub

Public Sub engineToInstanceMap_insert(ByVal ctx As Long, ByVal instance As ChakraScriptingInstance)
    If engineToInstanceMap Is Nothing Then Set engineToInstanceMap = CreateObject("Scripting.Dictionary")
    engineToInstanceMap(ctx) = ObjPtr(instance)
End Sub

Public Sub engineToInstanceMap_erase(ByVal ctx As Long)
    If engineToInstanceMap Is Nothing Then Exit Sub
    If engineToInstanceMap.Exists(ctx) Then engineToInstanceMap.Remove ctx
End Sub

' engineToInstanceMap.at(ctx) (std::out_of_range: a VB error)
Public Function engineToInstanceMap_at(ByVal ctx As Long) As ChakraScriptingInstance
    Dim o As ChakraScriptingInstance, p As Long
    If Not engineToInstanceMap Is Nothing Then
        If engineToInstanceMap.Exists(ctx) Then p = engineToInstanceMap(ctx)
    End If
    If p = 0 Then Err.Raise vbObjectError + 1, "std::out_of_range", "map::at"
    CopyMemory o, p, 4                  ' a strong reference made from the weak one
    Set engineToInstanceMap_at = o
    CopyMemory o, 0&, 4
End Function

' The instance of the current context
Public Function jsCurrentInstance() As ChakraScriptingInstance
    Dim ctx As Long
    JsGetCurrentContext ctx
    Set jsCurrentInstance = engineToInstanceMap_at(ctx)
End Function

' A new context's property descriptors, one per combination of QuickJS's flags (ctx current)
Public Sub chakraCreateDescriptors(ByVal ctx As Long)
    Dim f As Long, d As Long, arr(0 To 7) As Long
    If ctxDescriptors Is Nothing Then Set ctxDescriptors = CreateObject("Scripting.Dictionary")
    For f = 0 To 7
        d = jsNewObject()
        jsSetProp d, "writable", jsBool((f And JS_PROP_WRITABLE) <> 0)
        jsSetProp d, "enumerable", jsBool((f And JS_PROP_ENUMERABLE) <> 0)
        jsSetProp d, "configurable", jsBool((f And JS_PROP_CONFIGURABLE) <> 0)
        jsValueAddRef d
        arr(f) = d
    Next
    ctxDescriptors(ctx) = arr
End Sub

Public Sub chakraReleaseDescriptors(ByVal ctx As Long)
    Dim arr As Variant, f As Long
    If ctxDescriptors Is Nothing Then Exit Sub
    If Not ctxDescriptors.Exists(ctx) Then Exit Sub
    arr = ctxDescriptors(ctx)
    For f = 0 To 7
        jsValueRelease arr(f)
    Next
    ctxDescriptors.Remove ctx
    If descCtx = ctx Then descCtx = 0
End Sub

' The current context's descriptor for flags (jsrt_win's jsDefineProp)
Public Function chakraDescriptor(ByVal flags As Long) As Long
    Dim c As Long, arr As Variant, f As Long
    JsGetCurrentContext c
    If c <> descCtx Then
        arr = ctxDescriptors(c)
        For f = 0 To 7
            descArr(f) = arr(f)
        Next
        descCtx = c
    End If
    chakraDescriptor = descArr(flags)
End Function

' A script's source for JsParseScript: a fresh source context number
Public Function jsNextSourceCookie() As Long
    sourceCookie = sourceCookie + 1
    jsNextSourceCookie = sourceCookie
End Function

' The values of an event's or a timer's arguments (std::vector<JSValue>): each JsAddRef'd while kept
Public Sub jsArgsPush(args() As Long, n As Long, ByVal v As Long)
    If n = 0 Then
        ReDim args(0 To 3)
    ElseIf n > UBound(args) Then
        ReDim Preserve args(0 To n * 2)
    End If
    jsValueAddRef v
    args(n) = v
    n = n + 1
End Sub

Public Sub jsArgsRelease(args() As Long, ByVal n As Long)
    Dim i As Long
    For i = 0 To n - 1
        jsValueRelease args(i)
    Next
End Sub

' lib/framework/wzpaths.cpp's WzPathInfo (the path given as fromPlatformIndependentPath's)
Private Function WzPathInfo_fileName(ByVal file As String) As String
    Dim lastSlashPos As Long
    lastSlashPos = InStrRev(file, "/")
    If lastSlashPos = 0 Then
        ' no slash - treat the entire string as the filename
        WzPathInfo_fileName = file
        Exit Function
    End If
    WzPathInfo_fileName = Mid$(file, lastSlashPos + 1)
End Function

Private Function WzPathInfo_baseName(ByVal file As String) As String
    Dim result As String, firstPeriodPos As Long
    result = WzPathInfo_fileName(file)
    firstPeriodPos = InStr(result, ".")
    If firstPeriodPos = 0 Then
        ' no period
        WzPathInfo_baseName = result
        Exit Function
    End If
    WzPathInfo_baseName = Left$(result, firstPeriodPos - 1)
End Function

Private Function WzPathInfo_path(ByVal file As String) As String
    Dim lastSlashPos As Long
    lastSlashPos = InStrRev(file, "/")
    If lastSlashPos = 0 Then
        ' no slash - return "." as the path
        WzPathInfo_path = "."
        Exit Function
    End If
    If lastSlashPos - 1 < 1 Then
        WzPathInfo_path = Left$(file, 1)  ' std::max<size_t>(lastSlashPos, 1)
    Else
        WzPathInfo_path = Left$(file, lastSlashPos - 1)
    End If
End Function

Private Sub QJSRuntimeFree_LeakHandler_Error(ByVal msg As String)
    debug_ 175, LOG_ERROR, "QJSRuntimeFree_LeakHandler_Error", "QuickJS FreeRuntime leak: %s", msg
End Sub

'--------------------------------------------------------------------------
' Utility functions -- not called directly from scripts

' (VB) JS_DefinePropertyValue(Str) on a new object: C_W_E is a plain set (no property exists yet;
' "__proto__" would set the prototype), else jsDefineProp
Private Function defineOnNew(ByVal obj As Long, ByVal prop As String, ByVal val As Long, ByVal flags As Long) As Long
    If (flags And JS_PROP_C_W_E) = JS_PROP_C_W_E And prop <> "__proto__" Then
        If jsSetProp(obj, prop, val) Then defineOnNew = 1 Else defineOnNew = -1
        Exit Function
    End If
    defineOnNew = jsDefineProp(obj, prop, val, flags)
End Function

' JS_DefinePropertyValueUint32 on a new array
Private Function defineIndexOnNew(ByVal arr As Long, ByVal i As Long, ByVal val As Long, ByVal flags As Long) As Long
    If (flags And JS_PROP_C_W_E) = JS_PROP_C_W_E Then
        If jsSetIndex(arr, i, val) Then defineIndexOnNew = 1 Else defineIndexOnNew = -1
        Exit Function
    End If
    defineIndexOnNew = jsDefineProp(arr, CStr(i), val, flags)
End Function

Private Function mapJsonObjectToQuickJSValue(ByVal obj As nlohmann_json, ByVal prop_flags As Long) As Long
    Dim value As Long, prop_val As Long, it As Long
    value = jsNewObject()
    For it = 0 To obj.size() - 1
        prop_val = mapJsonToQuickJSValue(obj.value_at(it), prop_flags)
        defineOnNew value, obj.key_at(it), prop_val, prop_flags
    Next
    mapJsonObjectToQuickJSValue = value
End Function

Private Function mapJsonArrayToQuickJSValue(ByVal array_ As nlohmann_json, ByVal prop_flags As Long) As Long
    Dim value As Long, i As Long
    value = jsNewArray()
    For i = 0 To array_.size() - 1
        defineIndexOnNew value, i, mapJsonToQuickJSValue(array_.at_idx(i), prop_flags), prop_flags
    Next
    mapJsonArrayToQuickJSValue = value
End Function

Public Function mapJsonToQuickJSValue(ByVal instance As nlohmann_json, ByVal prop_flags As Long) As Long
    Select Case instance.type_
        ' IMPORTANT: To match the prior behavior of loading a QVariant from the JSON value and using engine->toScriptValue(QVariant)
        '			  to convert to a QScriptValue, "null" JSON values *MUST* map to QScriptValue::UndefinedValue.
        '
        '			  If they are set to QScriptValue::NullValue, it causes issues for libcampaign.js. (As the values become "defined".)
        '
        Case value_t_null: mapJsonToQuickJSValue = jsUndefinedValue()
        Case value_t_boolean: mapJsonToQuickJSValue = jsBool(instance.get_bool())
        Case value_t_number_integer: mapJsonToQuickJSValue = jsInt(instance.get_int())
        Case value_t_number_unsigned: mapJsonToQuickJSValue = jsUint(instance.get_uint())
        Case value_t_number_float: mapJsonToQuickJSValue = jsDouble(instance.get_double())
        Case value_t_string: mapJsonToQuickJSValue = jsStringValue(instance.get_string())
        Case value_t_array: mapJsonToQuickJSValue = mapJsonArrayToQuickJSValue(instance, prop_flags)
        Case value_t_object: mapJsonToQuickJSValue = mapJsonObjectToQuickJSValue(instance, prop_flags)
        Case value_t_binary
            debug_ 767, LOG_ERROR, "mapJsonToQuickJSValue", "Unexpected binary value type"
            mapJsonToQuickJSValue = jsUndefinedValue()
        Case value_t_discarded: mapJsonToQuickJSValue = jsUndefinedValue()
        Case Else: mapJsonToQuickJSValue = jsUndefinedValue() ' should never be reached
    End Select
End Function

Public Function JS_DeletePropertyStr(ByVal this_obj As Long, ByVal prop As String) As Long
    JS_DeletePropertyStr = jsDeleteProp(this_obj, prop)
End Function

Private Function QuickJS_DefinePropertyValue(ByVal this_obj As Long, ByVal prop As String, ByVal val As Long, ByVal flags As Long) As Long
    QuickJS_DefinePropertyValue = defineOnNew(this_obj, prop, val, flags)
End Function

' (VB) the selected field of an object (BASE_OBJECT::selected)
Private Function BASE_OBJECT_selected_(ByVal psObj As Long) As Long
    Select Case objHandleType(psObj)
    Case OBJ_DROID: BASE_OBJECT_selected_ = aDroid(psObj).selected
    Case OBJ_STRUCTURE: BASE_OBJECT_selected_ = aStruct(psObj).selected
    Case OBJ_FEATURE: BASE_OBJECT_selected_ = aFeature(psObj).selected
    End Select
End Function

';; ## Research (see quickjs_backend.cpp for the scripts' documentation of the objects)
' psResearch: a research index (-1 = nullptr)
Public Function convResearch(ByVal psResearch As Long, ByVal player As Long) As Long
    Dim value As Long, started As Boolean, i As Long, bits As Long
    If psResearch < 0 Then
        convResearch = jsNullValue()
        Exit Function
    End If
    value = jsNewObject()
    QuickJS_DefinePropertyValue value, "power", jsInt(asResearch(psResearch).researchPower), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "points", jsInt(asResearch(psResearch).researchPoints), JS_PROP_ENUMERABLE
    started = False
    For i = 0 To game.maxPlayers - 1
        If aiCheckAlliances(player, i) Or player = i Then
            bits = asPlayerResList(i, asResearch(psResearch).index).ResearchStatus
            started = started Or ((bits And STARTED_RESEARCH) <> 0) Or ((bits And STARTED_RESEARCH_PENDING) <> 0) Or ((bits And RESBITS_PENDING_ONLY) <> 0)
        End If
    Next
    QuickJS_DefinePropertyValue value, "started", jsBool(started), JS_PROP_ENUMERABLE ' including whether an ally has started it
    QuickJS_DefinePropertyValue value, "done", jsBool(IsResearchCompleted(asPlayerResList(player, asResearch(psResearch).index))), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "fullname", jsStringValue(asResearch(psResearch).name_), JS_PROP_ENUMERABLE ' temporary
    QuickJS_DefinePropertyValue value, "name", jsStringValue(asResearch(psResearch).id), JS_PROP_ENUMERABLE ' will be changed to contain fullname
    QuickJS_DefinePropertyValue value, "id", jsStringValue(asResearch(psResearch).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "type", jsInt(SCRIPT_RESEARCH), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "results", mapJsonToQuickJSValue(asResearch(psResearch).results, JS_PROP_ENUMERABLE), JS_PROP_ENUMERABLE
    convResearch = value
End Function

';; ## Structure
Public Function convStructure(ByVal psStruct As Long) As Long
    Dim aa As Boolean, ga As Boolean, indirect As Boolean, range As Long, i As Long, psWeap As Long
    Dim value As Long, stattype_ As Long, typ As Long, weaponlist As Long, j As Long, weapon_ As Long, psStats As Long
    aa = False
    ga = False
    indirect = False
    range = -1
    For i = 0 To aStruct(psStruct).numWeaps - 1
        If aStruct(psStruct).asWeaps(i).nStat <> 0 Then
            psWeap = BASE_OBJECT_getWeaponStats(psStruct, i)
            aa = aa Or ((asWeaponStats(psWeap).surfaceToAir And SHOOT_IN_AIR) <> 0)
            ga = ga Or ((asWeaponStats(psWeap).surfaceToAir And SHOOT_ON_GROUND) <> 0)
            indirect = indirect Or asWeaponStats(psWeap).movementModel = MM_INDIRECT Or asWeaponStats(psWeap).movementModel = MM_HOMINGINDIRECT
            range = MAX_(proj_GetLongRange(asWeaponStats(psWeap), aStruct(psStruct).player), range)
        End If
    Next
    value = convObj(psStruct)
    QuickJS_DefinePropertyValue value, "isCB", jsBool(structCBSensor(psStruct)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "isSensor", jsBool(structStandardSensor(psStruct)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "canHitAir", jsBool(aa), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "canHitGround", jsBool(ga), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "hasIndirect", jsBool(indirect), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "isRadarDetector", jsBool(objRadarDetector(psStruct)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "range", jsInt(range), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "status", jsInt(aStruct(psStruct).status), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "health", jsInt(FromU32(Fix(U32(MUL32(100, aStruct(psStruct).body)) / U32(MAXU_(1, STRUCTURE_structureBody(psStruct)))))), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "cost", jsInt(asStructureStats(aStruct(psStruct).pStructureType).powerToBuild), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "direction", jsInt(ToInt(UNDEG(aStruct(psStruct).rot.direction))), JS_PROP_ENUMERABLE
    stattype_ = 0
    typ = asStructureStats(aStruct(psStruct).pStructureType).typ
    Select Case typ ' don't bleed our source insanities into the scripting world
    Case REF_WALL, REF_WALLCORNER, REF_GATE
        stattype_ = REF_WALL
    Case REF_FORTRESS, REF_DEFENSE
        stattype_ = REF_DEFENSE
    Case Else
        stattype_ = typ
    End Select
    QuickJS_DefinePropertyValue value, "stattype", jsInt(stattype_), JS_PROP_ENUMERABLE
    If typ = REF_FACTORY Or typ = REF_CYBORG_FACTORY Or typ = REF_VTOL_FACTORY Or typ = REF_RESEARCH Or typ = REF_POWER_GEN Then
        QuickJS_DefinePropertyValue value, "modules", jsUint(aStruct(psStruct).capacity), JS_PROP_ENUMERABLE
    Else
        QuickJS_DefinePropertyValue value, "modules", jsNullValue(), JS_PROP_ENUMERABLE
    End If
    weaponlist = jsNewArray()
    For j = 0 To aStruct(psStruct).numWeaps - 1
        weapon_ = jsNewObject()
        psStats = BASE_OBJECT_getWeaponStats(psStruct, j)
        QuickJS_DefinePropertyValue weapon_, "fullname", jsStringValue(asWeaponStats(psStats).name_), JS_PROP_ENUMERABLE
        QuickJS_DefinePropertyValue weapon_, "name", jsStringValue(asWeaponStats(psStats).id), JS_PROP_ENUMERABLE ' will be changed to contain full name
        QuickJS_DefinePropertyValue weapon_, "id", jsStringValue(asWeaponStats(psStats).id), JS_PROP_ENUMERABLE
        QuickJS_DefinePropertyValue weapon_, "lastFired", jsUint(aStruct(psStruct).asWeaps(j).lastFired), JS_PROP_ENUMERABLE
        defineIndexOnNew weaponlist, j, weapon_, JS_PROP_ENUMERABLE
    Next
    QuickJS_DefinePropertyValue value, "weapons", weaponlist, JS_PROP_ENUMERABLE
    convStructure = value
End Function

';; ## Feature
Public Function convFeature(ByVal psFeature As Long) As Long
    Dim value As Long, psStats As Long
    value = convObj(psFeature)
    psStats = aFeature(psFeature).psStats
    QuickJS_DefinePropertyValue value, "health", jsUint(FromU32(Fix(U32(MUL32(100, asFeatureStats(psStats).body)) / U32(MAXU_(1, aFeature(psFeature).body))))), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "damageable", jsBool(asFeatureStats(psStats).damageable), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "stattype", jsInt(asFeatureStats(psStats).subType), JS_PROP_ENUMERABLE
    convFeature = value
End Function

';; ## Droid
Public Function convDroid(ByVal psDroid As Long) As Long
    Dim aa As Boolean, ga As Boolean, indirect As Boolean, range As Long, psBodyStats As Long, i As Long
    Dim psWeap As Long, typ As Long, value As Long, weaponlist As Long, j As Long, armed As Long, weapon_ As Long
    Dim psStats As Long, health As Double, ob As Long, cargoCount As Long
    aa = False
    ga = False
    indirect = False
    range = -1
    psBodyStats = DROID_getBodyStats(psDroid)

    For i = 0 To aDroid(psDroid).numWeaps - 1
        If aDroid(psDroid).asWeaps(i).nStat <> 0 Then
            If Not (aDroid(psDroid).asWeaps(i).nStat < asWeaponStats_size) Then ASSERT_FAILURE "psDroid->asWeaps[i].nStat < asWeaponStats.size()", "src/quickjs_backend.cpp:1032", "convDroid", "Invalid nStat (%d) referenced for asWeaps[%d]; numWeaponStats (%zu); droid: ""%s"" (numWeaps: %u)", aDroid(psDroid).asWeaps(i).nStat, i, asWeaponStats_size, aDroid(psDroid).aName, aDroid(psDroid).numWeaps
            psWeap = DROID_getWeaponStats(psDroid, i)
            aa = aa Or ((asWeaponStats(psWeap).surfaceToAir And SHOOT_IN_AIR) <> 0)
            ga = ga Or ((asWeaponStats(psWeap).surfaceToAir And SHOOT_ON_GROUND) <> 0)
            indirect = indirect Or asWeaponStats(psWeap).movementModel = MM_INDIRECT Or asWeaponStats(psWeap).movementModel = MM_HOMINGINDIRECT
            range = MAX_(proj_GetLongRange(asWeaponStats(psWeap), aDroid(psDroid).player), range)
        End If
    Next
    typ = aDroid(psDroid).droidType
    value = convObj(psDroid)
    QuickJS_DefinePropertyValue value, "action", jsInt(aDroid(psDroid).action), JS_PROP_ENUMERABLE
    If range >= 0 Then
        QuickJS_DefinePropertyValue value, "range", jsInt(range), JS_PROP_ENUMERABLE
    Else
        QuickJS_DefinePropertyValue value, "range", jsNullValue(), JS_PROP_ENUMERABLE
    End If
    QuickJS_DefinePropertyValue value, "order", jsInt(aDroid(psDroid).order.typ), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "cost", jsUint(calcDroidPower(psDroid)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "hasIndirect", jsBool(indirect), JS_PROP_ENUMERABLE
    Select Case aDroid(psDroid).droidType ' hide some engine craziness
    Case DROID_CYBORG_CONSTRUCT
        typ = DROID_CONSTRUCT
    Case DROID_CYBORG_SUPER
        typ = DROID_CYBORG
    Case DROID_DEFAULT
        typ = DROID_WEAPON
    Case DROID_CYBORG_REPAIR
        typ = DROID_REPAIR
    End Select
    QuickJS_DefinePropertyValue value, "bodySize", jsInt(asBodyStats(psBodyStats).size), JS_PROP_ENUMERABLE
    If DROID_isTransporter(psDroid) Then
        QuickJS_DefinePropertyValue value, "cargoCapacity", jsInt(TRANSPORTER_CAPACITY), JS_PROP_ENUMERABLE
        QuickJS_DefinePropertyValue value, "cargoLeft", jsInt(calcRemainingCapacity(psDroid)), JS_PROP_ENUMERABLE
        If aDroid(psDroid).psGroup <> 0 Then cargoCount = DROID_GROUP_getNumMembers(aDroid(psDroid).psGroup) Else cargoCount = 0
        QuickJS_DefinePropertyValue value, "cargoCount", jsUint(cargoCount), JS_PROP_ENUMERABLE
    End If
    QuickJS_DefinePropertyValue value, "isRadarDetector", jsBool(objRadarDetector(psDroid)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "isCB", jsBool(cbSensorDroid(psDroid)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "isSensor", jsBool(standardSensorDroid(psDroid)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "canHitAir", jsBool(aa), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "canHitGround", jsBool(ga), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "isVTOL", jsBool(DROID_isVtol(psDroid)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "isFlying", jsBool(DROID_isFlying(psDroid)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "droidType", jsInt(typ), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "experience", jsDouble(U32(aDroid(psDroid).experience) / 65536#), JS_PROP_ENUMERABLE
    ' 100.0 / (double)originalBody * (double)body (a 0 divisor: C's infinity, or NaN for 0 * infinity)
    ob = aDroid(psDroid).originalBody
    If ob <> 0 Then
        health = 100# / U32(ob) * U32(aDroid(psDroid).body)
    ElseIf aDroid(psDroid).body = 0 Then
        health = NaN_()
    Else
        health = DivByZero(1#)
    End If
    QuickJS_DefinePropertyValue value, "health", jsDouble(health), JS_PROP_ENUMERABLE

    QuickJS_DefinePropertyValue value, "body", jsStringValue(asBodyStats(DROID_getBodyStats(psDroid)).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "propulsion", jsStringValue(asPropulsionStats(DROID_getPropulsionStats(psDroid)).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "armed", jsDouble(0#), JS_PROP_ENUMERABLE ' deprecated!

    weaponlist = jsNewArray()
    For j = 0 To aDroid(psDroid).numWeaps - 1
        armed = droidReloadBar(psDroid, aDroid(psDroid).asWeaps(j), j)
        weapon_ = jsNewObject()
        psStats = DROID_getWeaponStats(psDroid, j)
        QuickJS_DefinePropertyValue weapon_, "fullname", jsStringValue(asWeaponStats(psStats).name_), JS_PROP_ENUMERABLE
        QuickJS_DefinePropertyValue weapon_, "name", jsStringValue(asWeaponStats(psStats).id), JS_PROP_ENUMERABLE ' will be changed to contain full name
        QuickJS_DefinePropertyValue weapon_, "id", jsStringValue(asWeaponStats(psStats).id), JS_PROP_ENUMERABLE
        QuickJS_DefinePropertyValue weapon_, "lastFired", jsUint(aDroid(psDroid).asWeaps(j).lastFired), JS_PROP_ENUMERABLE
        QuickJS_DefinePropertyValue weapon_, "armed", jsInt(armed), JS_PROP_ENUMERABLE
        defineIndexOnNew weaponlist, j, weapon_, JS_PROP_ENUMERABLE
    Next
    QuickJS_DefinePropertyValue value, "weapons", weaponlist, JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "cargoSize", jsInt(transporterSpaceRequired(psDroid)), JS_PROP_ENUMERABLE
    convDroid = value
End Function

Public Function convObj(ByVal psObj As Long) As Long
    Dim value As Long, psMap As GROUPMAP, group As Long, pos As Vector3i, objToGroup As Object
    value = jsNewObject()
    If Not (psObj <> 0) Then ASSERT_FAILURE "psObj", "src/quickjs_backend.cpp:1130", "convObj", "No object for conversion": convObj = value: Exit Function
    pos = SIMPLE_OBJECT_pos(psObj)
    QuickJS_DefinePropertyValue value, "id", jsUint(SIMPLE_OBJECT_id(psObj)), 0
    QuickJS_DefinePropertyValue value, "x", jsInt(map_coord(pos.x)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "y", jsInt(map_coord(pos.y)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "z", jsInt(map_coord(pos.z)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "player", jsUint(SIMPLE_OBJECT_player(psObj)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "armour", jsInt(objArmour(psObj, WC_KINETIC)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "thermal", jsInt(objArmour(psObj, WC_HEAT)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "type", jsInt(SIMPLE_OBJECT_type(psObj)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "selected", jsUint(BASE_OBJECT_selected_(psObj)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "name", jsStringValue(objInfo(psObj)), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "born", jsUint(SIMPLE_OBJECT_born(psObj)), JS_PROP_ENUMERABLE
    Set psMap = scripting_engine_getGroupMap(jsCurrentInstance())
    If Not psMap Is Nothing Then Set objToGroup = psMap.map_()
    If Not psMap Is Nothing Then
        If objToGroup.Exists(psObj) Then ' FIXME: psMap->map().count(psObj) > 0
            group = objToGroup(psObj) ' FIXME: psMap->map().at(psObj)
            QuickJS_DefinePropertyValue value, "group", jsInt(group), JS_PROP_ENUMERABLE
        Else
            QuickJS_DefinePropertyValue value, "group", jsNullValue(), JS_PROP_ENUMERABLE
        End If
    Else
        QuickJS_DefinePropertyValue value, "group", jsNullValue(), JS_PROP_ENUMERABLE
    End If
    convObj = value
End Function

' psTempl: a template handle (apsTemplates, 0 = nullptr)
Public Function convTemplate(ByVal psTempl As Long) As Long
    Dim value As Long, weaponlist As Long, j As Long
    value = jsNewObject()
    If Not (psTempl <> 0) Then ASSERT_FAILURE "psTempl", "src/quickjs_backend.cpp:1174", "convTemplate", "No object for conversion": convTemplate = value: Exit Function
    QuickJS_DefinePropertyValue value, "fullname", jsStringValue(apsTemplates(psTempl).name_), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "name", jsStringValue(apsTemplates(psTempl).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "id", jsStringValue(apsTemplates(psTempl).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "points", jsUint(calcTemplateBuild(apsTemplates(psTempl))), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "power", jsUint(calcTemplatePower(apsTemplates(psTempl))), JS_PROP_ENUMERABLE ' deprecated, use cost below
    QuickJS_DefinePropertyValue value, "cost", jsUint(calcTemplatePower(apsTemplates(psTempl))), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "droidType", jsInt(apsTemplates(psTempl).droidType), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "body", jsStringValue(asBodyStats(DROID_TEMPLATE_getBodyStats(apsTemplates(psTempl))).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "propulsion", jsStringValue(asPropulsionStats(DROID_TEMPLATE_getPropulsionStats(apsTemplates(psTempl))).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "brain", jsStringValue(asBrainStats(DROID_TEMPLATE_getBrainStats(apsTemplates(psTempl))).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "repair", jsStringValue(asRepairStats(DROID_TEMPLATE_getRepairStats(apsTemplates(psTempl))).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "ecm", jsStringValue(asECMStats(DROID_TEMPLATE_getECMStats(apsTemplates(psTempl))).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "sensor", jsStringValue(asSensorStats(DROID_TEMPLATE_getSensorStats(apsTemplates(psTempl))).id), JS_PROP_ENUMERABLE
    QuickJS_DefinePropertyValue value, "construct", jsStringValue(asConstructStats(DROID_TEMPLATE_getConstructStats(apsTemplates(psTempl))).id), JS_PROP_ENUMERABLE
    weaponlist = jsNewArray()
    For j = 0 To apsTemplates(psTempl).numWeaps - 1
        defineIndexOnNew weaponlist, j, jsStringValue(asWeaponStats(DROID_TEMPLATE_getWeaponStats(apsTemplates(psTempl), j)).id), JS_PROP_ENUMERABLE
    Next
    QuickJS_DefinePropertyValue value, "weapons", weaponlist, JS_PROP_ENUMERABLE
    convTemplate = value
End Function

Public Function convMax(ByVal psObj As Long) As Long
    If psObj = 0 Then
        convMax = jsNullValue()
        Exit Function
    End If
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID: convMax = convDroid(psObj)
    Case OBJ_STRUCTURE: convMax = convStructure(psObj)
    Case OBJ_FEATURE: convMax = convFeature(psObj)
    Case Else
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1209", "convMax", "No such supported object type: %d", SIMPLE_OBJECT_type(psObj)
        convMax = convObj(psObj)
    End Select
End Function

' (VB) C's MAX of 1 and an unsigned (compared unsigned)
Private Function MAXU_(ByVal a As Long, ByVal b As Long) As Long
    If ULT(b, a) Then MAXU_ = a Else MAXU_ = b
End Function

' Call a function by name (args(0 .. nargs - 1)); the result is undefined when the function threw
Public Function callFunction(ByVal ctx As Long, ByVal function_ As String, args() As Long, ByVal nargs As Long, Optional ByVal event_ As Boolean = True) As Long
    Dim instance As ChakraScriptingInstance, global_obj As Long, i As Long, funcName As String, value As Long
    Dim level As Long, result As Long, err_ As Long, time_begin As Currency
    Dim errVal As Long, isError As Boolean, result_str As String, ok As Boolean, nameStr As String, messageStr As String, stack As Long
    Dim prev As Long
    Set instance = engineToInstanceMap_at(ctx)
    prev = jsEnterContext(ctx)
    global_obj = instance.Get_Global_Obj()
    If event_ Then
        ' recurse into variants, if any
        For i = 0 To instance.eventNamespaces_size() - 1
            funcName = instance.eventNamespaces_item(i) & function_
            value = jsGetProp(global_obj, funcName)
            If jsIsFunction(value) Then
                callFunction ctx, funcName, args, nargs, event_
            End If
        Next
    End If
    If event_ Then level = LOG_SCRIPT Else level = LOG_ERROR
    value = jsGetProp(global_obj, function_)
    If Not jsIsFunction(value) Then
        ' not necessarily an error, may just be a trigger that is not defined (ie not needed)
        ' or it could be a typo in the function name or ...
        debug_ 1239, level, "callFunction", "called function (%s) not defined", function_
        callFunction = jsBool(False) ' ?? Shouldn't this be "undefined?"
        jsLeaveContext prev
        Exit Function
    End If

    ' scripting_engine::instance().executeWithPerformanceMonitoring(instance, function, [...](){ result = JS_Call(...); })
    time_begin = scripting_engine_executeWithPerformanceMonitoring_begin()
    err_ = jsCall(value, args, nargs, result)
    scripting_engine_executeWithPerformanceMonitoring_end instance, function_, time_begin

    If err_ <> JsNoError Then
        errVal = jsTakeException()
        isError = jsIsError(errVal)
        result_str = ""
        If isError Then
            nameStr = jsToStringOk(jsGetProp(errVal, "name"), ok)
            If ok Then result_str = nameStr Else result_str = "<unknown error name>"
            result_str = result_str & " - "
            messageStr = jsToStringOk(jsGetProp(errVal, "message"), ok)
            If ok Then result_str = result_str & """" & messageStr & """" Else result_str = result_str & "<no message>"
            result_str = result_str & ":: " & vbLf
            stack = jsGetProp(errVal, "stack")
            If Not jsIsUndefined(stack) Then
                messageStr = jsToStringOk(stack, ok)
                If ok Then result_str = result_str & messageStr
            End If
        End If
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1278", "callFunction", "Uncaught exception calling function ""%s"": %s", function_, result_str
        callFunction = jsUndefinedValue()
        jsLeaveContext prev
        Exit Function
    End If
    callFunction = result
    jsLeaveContext prev
End Function

'--------------------------------------------------------------------------
' The unbox / box helpers of the wrappers (one per C++ template instantiation the wrappers
' use). unbox_X(idx, argc, argv, function_) reads argv(idx), advances idx, as the C++
' unbox<X>()(idx, ctx, argc, argv, function) does; argv is padded with undefined up to the
' function's length (as QuickJS pads it), so a read at argc is undefined. An optional<X>
' parameter is `has = unbox_has(idx, argc): If has Then opt = unbox_X(...)`.
' UNBOX_SCRIPT_ASSERT logs and goes on (its JS_ThrowReferenceError is never raised: see the
' module header).

' (UNBOX_SCRIPT_ASSERT's JS_ThrowReferenceError: dropped, see the header)
Private Sub unboxScriptAssertFailed(ByVal expr As String, ByVal idx As Long, ByVal function_ As String)
End Sub

' unbox<optional<T>>: whether there is an argument at idx (argc <= idx: nullopt)
Public Function unbox_has(ByVal idx As Long, ByVal argc As Long) As Boolean
    unbox_has = (argc > idx)
End Function

' unbox<int>
Public Function unbox_int(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As Long
    Dim value As Long, ok As Boolean
    If argc <= idx Then Exit Function
    value = jsToInt32Ok(argv(idx), ok)
    idx = idx + 1
    If Not ok Then
        ' failed
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1349", "operator()", "Failed" ' TODO:
    End If
    unbox_int = value
End Function

' unbox<unsigned int> (the uint32_t's bits)
Public Function unbox_uint(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As Long
    Dim value As Long, ok As Boolean
    If argc <= idx Then Exit Function
    value = jsToInt32Ok(argv(idx), ok)
    idx = idx + 1
    If Not ok Then
        ' failed
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1366", "operator()", "Failed" ' TODO:
    End If
    unbox_uint = value
End Function

' unbox<bool>
Public Function unbox_bool(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As Boolean
    If argc <= idx Then Exit Function
    unbox_bool = jsToBool(argv(idx))
    idx = idx + 1
End Function

' unbox<float>
Public Function unbox_float(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As Single
    Dim value As Double, ok As Boolean
    If argc <= idx Then Exit Function
    value = jsToDoubleOk(argv(idx), ok)
    idx = idx + 1
    If Not ok Then
        ' failed
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1396", "operator()", "Failed" ' TODO:
    End If
    unbox_float = F32(value)
End Function

' unbox<double>
Public Function unbox_double(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As Double
    Dim value As Double, ok As Boolean
    If argc <= idx Then Exit Function
    value = jsToDoubleOk(argv(idx), ok)
    idx = idx + 1
    If Not ok Then
        ' failed
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1413", "operator()", "Failed" ' TODO:
    End If
    unbox_double = value
End Function

Public Function JSValueToInt32(ByVal value As Long) As Long
    Dim intVal As Long, ok As Boolean
    intVal = jsToInt32Ok(value, ok)
    If Not ok Then
        ' failed
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1425", "JSValueToInt32", "Failed" ' TODO:
    End If
    JSValueToInt32 = intVal
End Function

' (a failed property read is a failed conversion, as JS_ToInt32 of QuickJS's JS_EXCEPTION)
Public Function QuickJS_GetInt32(ByVal this_obj As Long, ByVal prop As String) As Long
    Dim val As Long
    If Not jsGetPropOk(this_obj, prop, val) Then
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1425", "JSValueToInt32", "Failed" ' TODO:
        Exit Function
    End If
    QuickJS_GetInt32 = JSValueToInt32(val)
End Function

Public Function JSValueToUint32(ByVal value As Long) As Long
    Dim uintVal As Long, ok As Boolean
    uintVal = jsToInt32Ok(value, ok)
    If Not ok Then
        ' failed
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1444", "JSValueToUint32", "Failed" ' TODO:
    End If
    JSValueToUint32 = uintVal
End Function

Public Function QuickJS_GetUint32(ByVal this_obj As Long, ByVal prop As String) As Long
    Dim val As Long
    If Not jsGetPropOk(this_obj, prop, val) Then
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1444", "JSValueToUint32", "Failed" ' TODO:
        Exit Function
    End If
    QuickJS_GetUint32 = JSValueToUint32(val)
End Function

Public Function JSValueToStdString(ByVal value As Long) As String
    JSValueToStdString = jsToString(value)
End Function

Public Function QuickJS_GetStdString(ByVal this_obj As Long, ByVal prop As String) As String
    QuickJS_GetStdString = JSValueToStdString(jsGetProp(this_obj, prop))
End Function

' The length of an array; False if arr is not one
Public Function QuickJS_GetArrayLength(ByVal arr As Long, len_ As Long) As Boolean
    QuickJS_GetArrayLength = jsArrayLength(arr, len_)
End Function

' unbox<DROID*> (and const DROID*)
Public Function unbox_DROID(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As Long
    Dim droidVal As Long, id As Long, player As Long, psDroid As Long
    If argc <= idx Then Exit Function
    droidVal = argv(idx)
    idx = idx + 1
    id = QuickJS_GetInt32(droidVal, "id")
    player = QuickJS_GetInt32(droidVal, "player")
    psDroid = IdToDroid(id, player)
    If psDroid = 0 Then
        debug_ 1505, LOG_ERROR, "operator()", "No such droid id %d belonging to player %d", id, player
        unboxScriptAssertFailed "psDroid", idx, function_
    End If
    unbox_DROID = psDroid
End Function

' unbox<STRUCTURE*> (and const STRUCTURE*)
Public Function unbox_STRUCTURE(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As Long
    Dim structVal As Long, id As Long, player As Long, psStruct As Long
    If argc <= idx Then Exit Function
    structVal = argv(idx)
    idx = idx + 1
    id = QuickJS_GetInt32(structVal, "id")
    player = QuickJS_GetInt32(structVal, "player")
    psStruct = IdToStruct(id, player)
    If psStruct = 0 Then
        debug_ 1530, LOG_ERROR, "operator()", "No such structure id %d belonging to player %d", id, player
        unboxScriptAssertFailed "psStruct", idx, function_
    End If
    unbox_STRUCTURE = psStruct
End Function

' unbox<BASE_OBJECT*> (and const BASE_OBJECT*)
Public Function unbox_BASE_OBJECT(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As Long
    Dim objVal As Long, oid As Long, oplayer As Long, otype As Long, psObj As Long
    If argc <= idx Then Exit Function
    objVal = argv(idx)
    idx = idx + 1
    oid = QuickJS_GetInt32(objVal, "id")
    oplayer = QuickJS_GetInt32(objVal, "player")
    otype = QuickJS_GetInt32(objVal, "type")
    psObj = IdToObject(otype, oid, oplayer)
    If psObj = 0 Then
        debug_ 1556, LOG_ERROR, "operator()", "No such object id %d belonging to player %d", oid, oplayer
        unboxScriptAssertFailed "psObj", idx, function_
    End If
    unbox_BASE_OBJECT = psObj
End Function

' unbox<std::string>
Public Function unbox_string(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As String
    If argc <= idx Then Exit Function
    unbox_string = JSValueToStdString(argv(idx))
    idx = idx + 1
End Function

' unbox<wzapi::STRUCTURE_TYPE_or_statsName_string>
Public Function unbox_STRUCTURE_TYPE_or_statsName_string(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As STRUCTURE_TYPE_or_statsName_string
    Dim result As STRUCTURE_TYPE_or_statsName_string, val As Long
    result.typ = NUM_DIFF_BUILDINGS
    If argc <= idx Then
        unbox_STRUCTURE_TYPE_or_statsName_string = result
        Exit Function
    End If
    val = argv(idx)
    idx = idx + 1
    If jsIsNumber(val) Then
        result.typ = JSValueToInt32(val)
    Else
        result.statsName = JSValueToStdString(val)
    End If
    unbox_STRUCTURE_TYPE_or_statsName_string = result
End Function

' unbox<wzapi::reservedParam>
Public Function unbox_reservedParam(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As Long
    If argc <= idx Then Exit Function
    ' just ignore parameter value, and increment idx
    idx = idx + 1
End Function

' unbox<wzapi::game_object_identifier> (the C++ tests argc < idx: at idx = argc it reads the padding)
Public Function unbox_game_object_identifier(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As game_object_identifier
    Dim objVal As Long, result As game_object_identifier
    result.typ = OBJ_NUM_TYPES
    result.id = -1
    result.player = -1
    If argc < idx Then
        unbox_game_object_identifier = result
        Exit Function
    End If
    objVal = argv(idx)
    idx = idx + 1
    result.id = QuickJS_GetInt32(objVal, "id")
    result.player = QuickJS_GetInt32(objVal, "player")
    result.typ = QuickJS_GetInt32(objVal, "type")
    unbox_game_object_identifier = result
End Function

' unbox<wzapi::string_or_string_list>
Public Function unbox_string_or_string_list(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As StringVector
    Dim strings As StringVector, list_or_string As Long, length As Long, k As Long
    If argc <= idx Then
        unbox_string_or_string_list = strings
        Exit Function
    End If
    list_or_string = argv(idx)
    idx = idx + 1
    If jsIsArray(list_or_string) Then
        If QuickJS_GetArrayLength(list_or_string, length) Then
            For k = 0 To length - 1 ' TODO: uint64_t isn't correct here, as we call GetUint32...
                StringVector_push_back strings, JSValueToStdString(jsGetIndex(list_or_string, k))
            Next
        End If
    Else
        StringVector_push_back strings, JSValueToStdString(list_or_string)
    End If
    unbox_string_or_string_list = strings
End Function

' unbox<wzapi::va_list_treat_as_strings>
Public Function unbox_va_list_treat_as_strings(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As StringVector
    Dim strings As StringVector, s As String, ok As Boolean
    If argc <= idx Then
        unbox_va_list_treat_as_strings = strings
        Exit Function
    End If
    Do While idx < argc
        s = jsToStringOk(argv(idx), ok)
        If Not ok Then ' (context->state() == QScriptContext::ExceptionState)
            Exit Do
        End If
        StringVector_push_back strings, s
        idx = idx + 1
    Loop
    unbox_va_list_treat_as_strings = strings
End Function

' unbox<wzapi::va_list<wzapi::string_or_string_list>>: the lists in result(0 .. count - 1)
Public Function unbox_va_list_string_or_string_list(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String, result() As StringVector) As Long
    Dim n As Long, before_idx As Long
    If argc <= idx Then Exit Function
    before_idx = idx
    Do While idx < argc
        before_idx = idx
        If n = 0 Then
            ReDim result(0 To 3)
        ElseIf n > UBound(result) Then
            ReDim Preserve result(0 To n * 2)
        End If
        result(n) = unbox_string_or_string_list(idx, argc, argv, function_)
        n = n + 1
        If before_idx = idx Then
            idx = idx + 1
        End If
    Loop
    unbox_va_list_string_or_string_list = n
End Function

' unbox<scripting_engine::area_by_values_or_area_label_lookup>
Public Function unbox_area_by_values_or_area_label_lookup(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As area_by_values_or_area_label_lookup
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    If argc <= idx Then
        unbox_area_by_values_or_area_label_lookup = area_by_values_or_area_label_lookup_new()
        Exit Function
    End If

    If jsIsString(argv(idx)) Then
        unbox_area_by_values_or_area_label_lookup = area_by_values_or_area_label_lookup_newLabel(JSValueToStdString(argv(idx)))
        idx = idx + 1
        Exit Function
    ElseIf (argc - idx) >= 4 Then
        x1 = JSValueToInt32(argv(idx))
        y1 = JSValueToInt32(argv(idx + 1))
        x2 = JSValueToInt32(argv(idx + 2))
        y2 = JSValueToInt32(argv(idx + 3))
        idx = idx + 4
        unbox_area_by_values_or_area_label_lookup = area_by_values_or_area_label_lookup_newArea(x1, y1, x2, y2)
        Exit Function
    Else
        ' could log an error here
    End If
    unbox_area_by_values_or_area_label_lookup = area_by_values_or_area_label_lookup_new()
End Function

' unbox<generic_script_object>
Public Function unbox_generic_script_object(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As generic_script_object
    Dim qval As Long, typ As Long, triggered As Long, id As Long, player As Long, psObj As Long
    If argc <= idx Then
        unbox_generic_script_object = generic_script_object_new()
        Exit Function
    End If

    qval = argv(idx)
    idx = idx + 1
    typ = QuickJS_GetInt32(qval, "type")

    triggered = jsGetProp(qval, "triggered")
    If jsIsNumber(triggered) Then
'        UNBOX_SCRIPT_ASSERT(context, type != SCRIPT_POSITION, "Cannot assign a trigger to a position");
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:1763", "operator()", "Not currently handling triggered property - does anything use this?"
    End If

    If typ = SCRIPT_RADIUS Then
        unbox_generic_script_object = generic_script_object_fromRadius(QuickJS_GetInt32(qval, "x"), QuickJS_GetInt32(qval, "y"), QuickJS_GetInt32(qval, "radius"))
        Exit Function
    ElseIf typ = SCRIPT_AREA Then
        unbox_generic_script_object = generic_script_object_fromArea(QuickJS_GetInt32(qval, "x"), QuickJS_GetInt32(qval, "y"), QuickJS_GetInt32(qval, "x2"), QuickJS_GetInt32(qval, "y2"))
        Exit Function
    ElseIf typ = SCRIPT_POSITION Then
        unbox_generic_script_object = generic_script_object_fromPosition(QuickJS_GetInt32(qval, "x"), QuickJS_GetInt32(qval, "y"))
        Exit Function
    ElseIf typ = SCRIPT_GROUP Then
        unbox_generic_script_object = generic_script_object_fromGroup(QuickJS_GetInt32(qval, "id"))
        Exit Function
    ElseIf typ = OBJ_DROID Or typ = OBJ_STRUCTURE Or typ = OBJ_FEATURE Then
        id = QuickJS_GetInt32(qval, "id")
        player = QuickJS_GetInt32(qval, "player")
        psObj = IdToObject(typ, id, player)
        If psObj = 0 Then
            debug_ 1802, LOG_ERROR, "operator()", "Object id %d not found belonging to player %d", id, player ' TODO: fail out
            unboxScriptAssertFailed "psObj", idx, function_
        End If
        unbox_generic_script_object = generic_script_object_fromObject(psObj)
        Exit Function
    End If
    unbox_generic_script_object = generic_script_object_new()
End Function

' unbox<wzapi::object_request> (the object_request constructors' fields)
Public Function unbox_object_request(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As object_request
    Dim r As object_request
    r.requestType = RequestType_INVALID_REQUEST
    r.val1 = -1
    r.val2 = -1
    r.val3 = -1
    If argc <= idx Then
        unbox_object_request = r
        Exit Function
    End If

    If (argc - idx) >= 3 Then ' get by ID case (3 parameters)
        r.requestType = RequestType_OBJECTID_REQUEST
        r.val1 = JSValueToInt32(argv(idx))      ' type
        r.val2 = JSValueToInt32(argv(idx + 1))  ' player
        r.val3 = JSValueToInt32(argv(idx + 2))  ' id
        idx = idx + 3
    ElseIf (argc - idx) >= 2 Then ' get at position case (2 parameters)
        r.requestType = RequestType_MAPPOS_REQUEST
        r.val1 = JSValueToInt32(argv(idx))      ' x
        r.val2 = JSValueToInt32(argv(idx + 1))  ' y
        idx = idx + 2
    Else
        ' get by label case (1 parameter)
        r.requestType = RequestType_LABEL_REQUEST
        r.str_ = JSValueToStdString(argv(idx))
        idx = idx + 1
    End If
    unbox_object_request = r
End Function

' unbox<wzapi::label_or_position_values> (its constructors' fields)
Public Function unbox_label_or_position_values(idx As Long, ByVal argc As Long, argv() As Long, ByVal function_ As String) As label_or_position_values
    Dim r As label_or_position_values
    r.typ = Invalid_Request
    r.x1 = -2                               ' VERY_LOW_INVALID_POS_VALUE
    r.y1 = -2
    If (argc - idx) >= 4 Then ' square area
        r.typ = Position_Values_Request
        r.x1 = JSValueToInt32(argv(idx))
        r.y1 = JSValueToInt32(argv(idx + 1))
        r.has_x2 = True
        r.x2 = JSValueToInt32(argv(idx + 2))
        r.has_y2 = True
        r.y2 = JSValueToInt32(argv(idx + 3))
        idx = idx + 4
    ElseIf (argc - idx) >= 2 Then ' single tile
        r.typ = Position_Values_Request
        r.x1 = JSValueToInt32(argv(idx))
        r.y1 = JSValueToInt32(argv(idx + 1))
        idx = idx + 2
    ElseIf (argc - idx) >= 1 Then ' label
        r.typ = Label_Request
        r.label = JSValueToStdString(argv(idx))
        idx = idx + 1
    End If
    unbox_label_or_position_values = r
End Function

' box(unsigned char), box(bool), box(int), box(unsigned int)
Public Function box_uchar(ByVal val As Long) As Long
    box_uchar = jsUint(val)
End Function

Public Function box_bool(ByVal val As Boolean) As Long
    box_bool = jsBool(val)
End Function

Public Function box_int(ByVal val As Long) As Long
    box_int = jsInt(val)
End Function

Public Function box_uint(ByVal val As Long) As Long
    box_uint = jsUint(val)
End Function

' box(const char *), box(std::string)
Public Function box_cstring(ByVal str_ As String) As Long
    box_cstring = jsStringValue(str_)
End Function

Public Function box_string(ByVal str_ As String) As Long
    box_string = jsStringValue(str_)
End Function

' box(wzapi::no_return_value)
Public Function box_no_return_value() As Long
    box_no_return_value = jsUndefinedValue()
End Function

' box(const BASE_OBJECT *), box(const STRUCTURE *), box(const DROID *), box(const FEATURE *),
' box(const DROID_TEMPLATE *): null for nullptr (also event_nullable_ptr<> and returned_nullable_ptr<>)
Public Function box_BASE_OBJECT(ByVal psObj As Long) As Long
    If psObj = 0 Then
        box_BASE_OBJECT = jsNullValue()
        Exit Function
    End If
    box_BASE_OBJECT = convMax(psObj)
End Function

Public Function box_STRUCTURE(ByVal psStruct As Long) As Long
    If psStruct = 0 Then
        box_STRUCTURE = jsNullValue()
        Exit Function
    End If
    box_STRUCTURE = convStructure(psStruct)
End Function

Public Function box_DROID(ByVal psDroid As Long) As Long
    If psDroid = 0 Then
        box_DROID = jsNullValue()
        Exit Function
    End If
    box_DROID = convDroid(psDroid)
End Function

Public Function box_FEATURE(ByVal psFeat As Long) As Long
    If psFeat = 0 Then
        box_FEATURE = jsNullValue()
        Exit Function
    End If
    box_FEATURE = convFeature(psFeat)
End Function

Public Function box_DROID_TEMPLATE(ByVal psTemplate As Long) As Long
    If psTemplate = 0 Then
        box_DROID_TEMPLATE = jsNullValue()
        Exit Function
    End If
    box_DROID_TEMPLATE = convTemplate(psTemplate)
End Function

Public Function box_scr_radius(r As scr_radius) As Long
    Dim ret As Long
    ret = jsNewObject()
    QuickJS_DefinePropertyValue ret, "type", jsInt(SCRIPT_RADIUS), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "x", jsInt(r.x), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "y", jsInt(r.y), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "radius", jsInt(r.radius), JS_PROP_C_W_E
    box_scr_radius = ret
End Function

Public Function box_scr_area(r As scr_area) As Long
    Dim ret As Long
    ret = jsNewObject()
    QuickJS_DefinePropertyValue ret, "type", jsInt(SCRIPT_AREA), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "x", jsInt(r.x1), JS_PROP_C_W_E ' TODO: Rename scr_area x1 to x
    QuickJS_DefinePropertyValue ret, "y", jsInt(r.y1), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "x2", jsInt(r.x2), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "y2", jsInt(r.y2), JS_PROP_C_W_E
    box_scr_area = ret
End Function

Public Function box_scr_position(p As scr_position) As Long
    Dim ret As Long
    ret = jsNewObject()
    QuickJS_DefinePropertyValue ret, "type", jsInt(SCRIPT_POSITION), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "x", jsInt(p.x), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "y", jsInt(p.y), JS_PROP_C_W_E
    box_scr_position = ret
End Function

Public Function box_generic_script_object(p As generic_script_object) As Long
    Dim typ As Long, ret As Long, r As scr_radius, a As scr_area, pos As scr_position
    typ = generic_script_object_getType(p)
    Select Case typ
    Case SCRIPT_RADIUS
        r = generic_script_object_getRadius(p)
        box_generic_script_object = box_scr_radius(r)
        Exit Function
    Case SCRIPT_AREA
        a = generic_script_object_getArea(p)
        box_generic_script_object = box_scr_area(a)
        Exit Function
    Case SCRIPT_POSITION
        pos = generic_script_object_getPosition(p)
        box_generic_script_object = box_scr_position(pos)
        Exit Function
    Case SCRIPT_GROUP
        ret = jsNewObject()
        QuickJS_DefinePropertyValue ret, "type", jsInt(typ), JS_PROP_C_W_E
        QuickJS_DefinePropertyValue ret, "id", jsInt(generic_script_object_getGroupId(p)), JS_PROP_C_W_E
        box_generic_script_object = ret
        Exit Function
    Case OBJ_DROID, OBJ_FEATURE, OBJ_STRUCTURE
        box_generic_script_object = convMax(generic_script_object_getObject(p))
        Exit Function
    Case Else
        If generic_script_object_isNull(p) Then
            box_generic_script_object = jsNullValue()
            Exit Function
        End If
        debug_ 2025, LOG_SCRIPT, "box", "Unsupported object label type: %d", typ
    End Select

    box_generic_script_object = jsNullValue()
End Function

Public Function box_GATEWAY(psGateway As GATEWAY) As Long
    Dim ret As Long
    ret = jsNewObject()
    QuickJS_DefinePropertyValue ret, "x1", jsInt(psGateway.x1), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "y1", jsInt(psGateway.y1), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "x2", jsInt(psGateway.x2), JS_PROP_C_W_E
    QuickJS_DefinePropertyValue ret, "y2", jsInt(psGateway.y2), JS_PROP_C_W_E
    box_GATEWAY = ret
End Function

' (VB) box(const std::vector<VectorType>&)'s element definition and its failure message
Private Sub boxVectorSet(ByVal result As Long, ByVal i As Long, ByVal v As Long)
    If Not jsSetIndex(result, i, v) Then
        ' Failed to define property value??
        debug_ 2052, LOG_ERROR, "box", "Failed to define property value vector[%u]", i
    End If
End Sub

' box(const std::vector<const BASE_OBJECT *>&) and the vectors of the other object types
Public Function box_vector_BASE_OBJECT(value() As Long, ByVal n As Long) As Long
    Dim result As Long, i As Long
    result = jsNewArray()
    For i = 0 To n - 1
        boxVectorSet result, i, box_BASE_OBJECT(value(i))
    Next
    box_vector_BASE_OBJECT = result
End Function

Public Function box_vector_STRUCTURE(value() As Long, ByVal n As Long) As Long
    Dim result As Long, i As Long
    result = jsNewArray()
    For i = 0 To n - 1
        boxVectorSet result, i, box_STRUCTURE(value(i))
    Next
    box_vector_STRUCTURE = result
End Function

Public Function box_vector_DROID(value() As Long, ByVal n As Long) As Long
    Dim result As Long, i As Long
    result = jsNewArray()
    For i = 0 To n - 1
        boxVectorSet result, i, box_DROID(value(i))
    Next
    box_vector_DROID = result
End Function

Public Function box_vector_FEATURE(value() As Long, ByVal n As Long) As Long
    Dim result As Long, i As Long
    result = jsNewArray()
    For i = 0 To n - 1
        boxVectorSet result, i, box_FEATURE(value(i))
    Next
    box_vector_FEATURE = result
End Function

' box(const std::vector<std::string>&)
Public Function box_vector_string(value As StringVector) As Long
    Dim result As Long, i As Long
    result = jsNewArray()
    For i = 0 To value.size - 1
        boxVectorSet result, i, box_string(value.items(i))
    Next
    box_vector_string = result
End Function

' box(const std::vector<int>&)
Public Function box_vector_int(value() As Long, ByVal n As Long) As Long
    Dim result As Long, i As Long
    result = jsNewArray()
    For i = 0 To n - 1
        boxVectorSet result, i, box_int(value(i))
    Next
    box_vector_int = result
End Function

' box(const std::vector<scr_position>&)
Public Function box_vector_scr_position(value() As scr_position, ByVal n As Long) As Long
    Dim result As Long, i As Long
    result = jsNewArray()
    For i = 0 To n - 1
        boxVectorSet result, i, box_scr_position(value(i))
    Next
    box_vector_scr_position = result
End Function

' box(const std::list<GATEWAY *>&) (GATEWAY_LIST)
Public Function box_list_GATEWAY(value() As GATEWAY, ByVal n As Long) As Long
    Dim result As Long, i As Long
    result = jsNewArray()
    For i = 0 To n - 1
        If Not jsSetIndex(result, i, box_GATEWAY(value(i))) Then
            ' Failed to define property value??
            debug_ 2069, LOG_ERROR, "box", "Failed to define property value list[%u]", i
        End If
    Next
    box_list_GATEWAY = result
End Function

Public Function box_researchResult(result As wzapi_researchResult) As Long
    If result.psResearch < 0 Then
        box_researchResult = jsNullValue()
        Exit Function
    End If
    box_researchResult = convResearch(result.psResearch, result.player)
End Function

Public Function box_researchResults(results As wzapi_researchResults) As Long
    Dim result As Long, i As Long, psResearch As Long
    result = jsNewArray()
    For i = 0 To results.resList_size - 1
        psResearch = results.resList(i)
        jsSetIndex result, i, convResearch(psResearch, results.player) ' TODO: Check return value?
    Next
    box_researchResults = result
End Function

' box(nlohmann::json)
Public Function box_json(ByVal results As nlohmann_json) As Long
    box_json = mapJsonToQuickJSValue(results, JS_PROP_C_W_E)
End Function

'--------------------------------------------------------------------------
' Wraps a QuickJS instance

' (QuickJS's SCRIPT_ASSERT in the builtins: debug(LOG_ERROR, ...), a JS_ThrowReferenceError that is
' never raised (see the header) and JS_NULL)

'-- ## profile(functionName[, arguments])
Private Function js_profile(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim functionName As String, args() As Long, n As Long, i As Long
    If Not (argc >= 1) Then debug_ 2349, LOG_ERROR, "js_profile", "Must have at least one parameter": js_profile = jsNullValue(): Exit Function
    If Not (jsIsString(argv(0))) Then debug_ 2350, LOG_ERROR, "js_profile", "Profiled functions must be quoted": js_profile = jsNullValue(): Exit Function
    functionName = JSValueToStdString(argv(0))
    For i = 1 To argc - 1
        jsArgsPush args, n, argv(i)
    Next
    js_profile = callFunction(ctx, functionName, args, n)
    jsArgsRelease args, n
End Function

Private Function QuickJS_DumpObject(ByVal obj As Long) As String
    Dim result As String, ok As Boolean, s As String
    s = jsToStringOk(obj, ok)
    If ok Then
        result = s
    Else
        result = "[failed to convert object to string]"
    End If
    QuickJS_DumpObject = result
End Function

Public Function QuickJS_DumpError() As String
    Dim result As String, exception_val As Long, isError As Boolean, stack As Long
    exception_val = jsTakeException()
    isError = jsIsError(exception_val)
    result = QuickJS_DumpObject(exception_val)
    If isError Then
        stack = jsGetProp(exception_val, "stack")
        If Not jsIsUndefined(stack) Then
            result = result & vbLf & QuickJS_DumpObject(stack)
        End If
    End If
    QuickJS_DumpError = result
End Function

' A script's text for JsParseScript (UTF-8 to UTF-16)
Public Function scriptSourceText(bytes() As Byte, ByVal size As Long) As String
    If size <= 0 Then
        scriptSourceText = ""
        Exit Function
    End If
    scriptSourceText = utf8_decode(bytes, 0, size)
End Function

'-- ## include(filePath)
Private Function js_include(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim filePath As String, instance As ChakraScriptingInstance, size As Long, bytes() As Byte, loadedFilePath As String
    Dim source As String, compiledFuncObj As Long, result As Long, errorAsString As String, args() As Long, k As Long
    If Not (argc = 1) Then debug_ 2403, LOG_ERROR, "js_include", "Must specify a file to include": js_include = jsNullValue(): Exit Function
    filePath = JSValueToStdString(argv(0))
    If Not (strEndsWith(filePath, ".js")) Then debug_ 2405, LOG_ERROR, "js_include", "Include file must end in .js": js_include = jsNullValue(): Exit Function
    Set instance = engineToInstanceMap_at(ctx)
    If Not instance.loadFileForInclude(filePath, loadedFilePath, bytes, size, LoadFileSearchOptions_All_BackwardsCompat) Then
        debug_ 2412, LOG_ERROR, "js_include", "Failed to read include file ""%s""", filePath
        ' (JS_ThrowReferenceError(ctx, "Failed to read include file \"%s\"", ...) and a normal return: never raised)
        js_include = jsBool(False)
        Exit Function
    End If
    source = scriptSourceText(bytes, size)
    k = instance.keepSource(source, loadedFilePath)
    If JsParseScript(instance.sourcePtr(k), jsNextSourceCookie(), instance.urlPtr(k), compiledFuncObj) <> JsNoError Then
        ' compilation error / syntax error
        errorAsString = QuickJS_DumpError()
        debug_ 2422, LOG_ERROR, "js_include", "Syntax error in include file %s: %s", loadedFilePath, errorAsString
        js_include = jsBool(False)
        Exit Function
    End If
    If jsCall(compiledFuncObj, args, 0, result) <> JsNoError Then
        errorAsString = QuickJS_DumpError()
        debug_ 2432, LOG_ERROR, "js_include", "Uncaught exception in include file %s: %s", loadedFilePath, errorAsString
        js_include = jsBool(False)
        Exit Function
    End If
    debug_ 2437, LOG_SCRIPT, "js_include", "Included new script file %s", loadedFilePath
    js_include = jsBool(True)
End Function

'-- ## includeJSON(filePath[, quiet])
' (JS_ParseJSON is the global JSON.parse; its exception is raised, as QuickJS's JS_EXCEPTION is)
Private Function js_includeJSON(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim filePath As String, quiet As Boolean, instance As ChakraScriptingInstance, size As Long, bytes() As Byte
    Dim loadedFilePath As String, level As Long, r As Long, jsonObj As Long, parseFn As Long, args() As Long, n As Long
    Dim e As Long
    If Not (argc >= 1) Then debug_ 2451, LOG_ERROR, "js_includeJSON", "Must specify a file to include": js_includeJSON = jsNullValue(): Exit Function
    filePath = JSValueToStdString(argv(0))
    If Not (strEndsWith(filePath, ".json")) Then debug_ 2453, LOG_ERROR, "js_includeJSON", "Include file must end in .json": js_includeJSON = jsNullValue(): Exit Function
    quiet = False
    If argc >= 2 Then
        If jsIsBool(argv(1)) Then
            quiet = jsToBool(argv(1))
        End If
    End If
    Set instance = engineToInstanceMap_at(ctx)
    If Not instance.loadFileForInclude(filePath, loadedFilePath, bytes, size, LoadFileSearchOptions_ScriptPath) Then
        If quiet Then level = LOG_SCRIPT Else level = LOG_ERROR
        debug_ 2469, level, "js_includeJSON", "Failed to read include file ""%s""", filePath
        ' (if (!quiet) JS_ThrowReferenceError(...) and a normal return: never raised)
        js_includeJSON = jsBool(False)
        Exit Function
    End If
    jsonObj = jsGetProp(instance.Get_Global_Obj(), "JSON")
    parseFn = jsGetProp(jsonObj, "parse")
    jsArgsPush args, n, jsStringValue(scriptSourceText(bytes, size))
    If jsCall(parseFn, args, n, r) <> JsNoError Then
        ' the exception stays pending: raised when the function returns (JS_EXCEPTION)
        e = jsTakeException()
        JsSetException e
        r = 0
    End If
    jsArgsRelease args, n
    debug_ 2478, LOG_SCRIPT, "js_includeJSON", "Included new JSON file %s", loadedFilePath
    js_includeJSON = r
End Function

' SetQuickJSTimer
Private Function SetQuickJSTimer(ByVal ctx As Long, ByVal player As Long, ByVal funcName As String, ByVal ms As Long, ByVal stringArg As String, ByVal psObj As Long, ByVal typ As timerType) As Double
    Dim timerFunc As ChakraTimerFunc, additionalParams As ChakraTimerData
    ' timerFunc
    Set timerFunc = New ChakraTimerFunc
    timerFunc.init ctx, funcName
    ' additionalParams
    Set additionalParams = New ChakraTimerData
    additionalParams.stringArg = stringArg
    SetQuickJSTimer = scripting_engine_setTimer(engineToInstanceMap_at(ctx), timerFunc, player, ms, funcName, psObj, typ, additionalParams)
End Function

' The function SetQuickJSTimer and restoreTimerFunction give a timer (their lambda): the
' object or the string argument, then the call
Public Sub runQuickJSTimerFunction(ByVal ctx As Long, ByVal funcName As String, ByVal baseObject As Long, ByVal additionalParams As Object)
    Dim pData As ChakraTimerData, args() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    If Not additionalParams Is Nothing Then
        If TypeOf additionalParams Is ChakraTimerData Then Set pData = additionalParams
    End If
    If baseObject <> 0 Then
        jsArgsPush args, n, convMax(baseObject)
    ElseIf Not pData Is Nothing Then
        If Len(pData.stringArg) > 0 Then
            jsArgsPush args, n, jsStringValue(pData.stringArg)
        End If
    End If
    callFunction ctx, funcName, args, n, True
    jsArgsRelease args, n
    jsLeaveContext prev
End Sub

'-- ## setTimer(functionName, milliseconds[, object])
Private Function js_setTimer(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim functionName As String, ms As Long, global_obj As Long, player As Long, funcObj As Long
    Dim stringArg As String, psObj As Long, obj As Long, baseobj As Long, baseobjtype As Long
    If Not (argc >= 2) Then debug_ 2534, LOG_ERROR, "js_setTimer", "Must have at least two parameters": js_setTimer = jsNullValue(): Exit Function
    If Not (jsIsString(argv(0))) Then debug_ 2535, LOG_ERROR, "js_setTimer", "Timer functions must be quoted": js_setTimer = jsNullValue(): Exit Function
    functionName = JSValueToStdString(argv(0))
    ms = JSValueToInt32(argv(1))

    JsGetGlobalObject global_obj
    player = QuickJS_GetInt32(global_obj, "me")

    funcObj = jsGetProp(global_obj, functionName) ' check existence
    If Not (jsIsFunction(funcObj)) Then debug_ 2545, LOG_ERROR, "js_setTimer", "No such function: %s", functionName: js_setTimer = jsNullValue(): Exit Function

    stringArg = ""
    psObj = 0
    If argc = 3 Then
        obj = argv(2)
        If jsIsString(obj) Then
            stringArg = JSValueToStdString(obj)
        Else ' is game object
            baseobj = QuickJS_GetInt32(obj, "id")
            baseobjtype = QuickJS_GetInt32(obj, "type")
            psObj = IdToObject(baseobjtype, baseobj, player)
        End If
    End If

    SetQuickJSTimer ctx, player, functionName, ms, stringArg, psObj, TIMER_REPEAT

    js_setTimer = jsBool(True)
End Function

'-- ## removeTimer(functionName)
Private Function js_removeTimer(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim functionName As String, global_obj As Long, player As Long, instance As ChakraScriptingInstance
    Dim removedTimerIDs() As Double, removedTimerIDs_size As Long, warnName As String
    If Not (argc = 1) Then debug_ 2576, LOG_ERROR, "js_removeTimer", "Must have one parameter": js_removeTimer = jsNullValue(): Exit Function
    If Not (jsIsString(argv(0))) Then debug_ 2577, LOG_ERROR, "js_removeTimer", "Timer functions must be quoted": js_removeTimer = jsNullValue(): Exit Function
    functionName = JSValueToStdString(argv(0))

    JsGetGlobalObject global_obj
    player = QuickJS_GetInt32(global_obj, "me")

    Set instance = engineToInstanceMap_at(ctx)
    ' removeTimersIf([instance, functionName, player](const timerNode& node) { return node.instance == instance && node.timerName == functionName && node.player == player; })
    removedTimerIDs_size = scripting_engine_removeTimersIf(TIMERPRED_INSTANCE_NAME_PLAYER, instance, functionName, player, 0, removedTimerIDs)
    If removedTimerIDs_size = 0 Then
        ' Friendly warning
        warnName = functionName
        debug_ 2594, LOG_ERROR, "js_removeTimer", "Did not find timer %s to remove", warnName
        js_removeTimer = jsBool(False)
        Exit Function
    End If
    js_removeTimer = jsBool(True)
End Function

'-- ## queue(functionName[, milliseconds[, object]])
' TODO, check if an identical call is already queued up - and in this case,
' do not add anything.
Private Function js_queue(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim functionName As String, global_obj As Long, funcObj As Long, ms As Long, player As Long
    Dim stringArg As String, psObj As Long, obj As Long, baseobj As Long, baseobjtype As Long
    If Not (argc >= 1) Then debug_ 2615, LOG_ERROR, "js_queue", "Must have at least one parameter": js_queue = jsNullValue(): Exit Function
    If Not (jsIsString(argv(0))) Then debug_ 2616, LOG_ERROR, "js_queue", "Queued functions must be quoted": js_queue = jsNullValue(): Exit Function
    functionName = JSValueToStdString(argv(0))

    JsGetGlobalObject global_obj

    funcObj = jsGetProp(global_obj, functionName) ' check existence
    If Not (jsIsFunction(funcObj)) Then debug_ 2624, LOG_ERROR, "js_queue", "No such function: %s", functionName: js_queue = jsNullValue(): Exit Function

    ms = 0
    If argc > 1 Then
        ms = JSValueToInt32(argv(1))
    End If
    player = QuickJS_GetInt32(global_obj, "me")

    stringArg = ""
    psObj = 0
    If argc = 3 Then
        obj = argv(2)
        If jsIsString(obj) Then
            stringArg = JSValueToStdString(obj)
        Else ' is game object
            baseobj = QuickJS_GetInt32(obj, "id")
            baseobjtype = QuickJS_GetInt32(obj, "type")
            psObj = IdToObject(baseobjtype, baseobj, player)
        End If
    End If

    SetQuickJSTimer ctx, player, functionName, ms, stringArg, psObj, TIMER_ONESHOT_READY

    js_queue = jsBool(True)
End Function

'-- ## namespace(prefix)
Private Function js_namespace(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim prefix As String, instance As ChakraScriptingInstance
    If Not (argc = 1) Then debug_ 2664, LOG_ERROR, "js_namespace", "Must have one parameter": js_namespace = jsNullValue(): Exit Function
    If Not (jsIsString(argv(0))) Then debug_ 2665, LOG_ERROR, "js_namespace", "Must provide a string namespace prefix": js_namespace = jsNullValue(): Exit Function
    prefix = JSValueToStdString(argv(0))
    Set instance = engineToInstanceMap_at(ctx)
    instance.eventNamespaces_push_back prefix
    js_namespace = jsBool(True)
End Function

' js_builtin_funcs: JS_SetPropertyFunctionList(ctx, global_obj, js_builtin_funcs, ...) (QJS_CFUNC_DEF:
' writable and configurable, not enumerable); ctx current
Public Function setBuiltinFunctions(ByVal instance As ChakraScriptingInstance) As Boolean
    Dim global_obj As Long, script As Long, result As Long, args() As Long, k As Long
    global_obj = instance.Get_Global_Obj()
    defineBuiltin global_obj, "setTimer", JSFN_SETTIMER, 2 ' JS-specific implementation
    defineBuiltin global_obj, "queue", JSFN_QUEUE, 1 ' JS-specific implementation
    defineBuiltin global_obj, "removeTimer", JSFN_REMOVETIMER, 1 ' JS-specific implementation
    defineBuiltin global_obj, "profile", JSFN_PROFILE, 1 ' JS-specific implementation
    defineBuiltin global_obj, "include", JSFN_INCLUDE, 1 ' backend-specific (a scripting_instance can't directly include a different type of script)
    defineBuiltin global_obj, "includeJSON", JSFN_INCLUDEJSON, 1 ' JS-specific JSON loading
    defineBuiltin global_obj, "namespace", JSFN_NAMESPACE, 1 ' JS-specific implementation
    ' debugGetCallerFuncObject, debugGetCallerFuncName, debugGetBacktrace: backend-specific
    k = instance.keepSource(DEBUG_BUILTINS_SOURCE, "<builtins>")
    If JsParseScript(instance.sourcePtr(k), jsNextSourceCookie(), instance.urlPtr(k), script) = JsNoError Then
        If jsCall(script, args, 0, result) = JsNoError Then setBuiltinFunctions = True
    End If
    jsClearException
End Function

Private Sub defineBuiltin(ByVal global_obj As Long, ByVal name_ As String, ByVal fnNumber As Long, ByVal length As Long)
    jsDefineProp global_obj, name_, jsNewCFunction(fnNumber, length), JS_PROP_WRITABLE Or JS_PROP_CONFIGURABLE
End Sub

' JS_NewCFunction(ctx, impl, name, length): a native function whose calls come to jsNativeFunction
' with fnNumber; QuickJS pads its arguments with undefined up to length
Public Function jsNewCFunction(ByVal fnNumber As Long, ByVal length As Long) As Long
    Dim fn As Long, i As Long
    If fnNumber >= 0 And fnNumber < 65536 Then
        If fnNumber >= fnLengths_size Then
            ReDim Preserve fnLengths(0 To fnNumber + 256)
            For i = fnLengths_size To fnNumber + 256
                fnLengths(i) = 0
            Next
            fnLengths_size = fnNumber + 257
        End If
        fnLengths(fnNumber) = length
    End If
    If JsCreateFunction(AddressOf jsNativeFunction, fnNumber, fn) <> JsNoError Then fn = 0
    jsNewCFunction = fn
End Function

' The native functions' entry (JsNativeFunction): arguments(0) is this, then the arguments;
' callbackState the function's number. QuickJS's C functions are no constructors.
Public Function jsNativeFunction(ByVal callee As Long, ByVal isConstructCall As Long, ByVal arguments As Long, ByVal argumentCount As Long, ByVal callbackState As Long) As Long
    Dim argc As Long, n As Long, argv() As Long, thisVal As Long, ctx As Long, r As Long, i As Long
    Dim errNumber As Long, errDescription As String
    On Error GoTo fail
    jsClearFPU                                          ' (VB) the script's floating point flags (see jsrt_win)
    argumentCount = argumentCount And &HFFFF&
    argc = argumentCount - 1
    If argc < 0 Then argc = 0
    If (isConstructCall And &HFF&) <> 0 Then
        jsThrowTypeError "not a constructor"
        jsNativeFunction = 0
        Exit Function
    End If
    ' QuickJS pads argv to the function's length with undefined (argc stays the number given)
    n = argc
    If callbackState >= 0 And callbackState < fnLengths_size Then
        If n < fnLengths(callbackState) Then n = fnLengths(callbackState)
    End If
    ReDim argv(0 To n)
    If argumentCount > 0 Then CopyMemory thisVal, ByVal arguments, 4
    If argc > 0 Then CopyMemory argv(0), ByVal arguments + 4, argc * 4
    For i = argc To n
        argv(i) = jsUndefinedValue()
    Next
    JsGetCurrentContext ctx

    Select Case callbackState
    Case JSFN_SETTIMER: r = js_setTimer(ctx, argc, argv)
    Case JSFN_QUEUE: r = js_queue(ctx, argc, argv)
    Case JSFN_REMOVETIMER: r = js_removeTimer(ctx, argc, argv)
    Case JSFN_PROFILE: r = js_profile(ctx, argc, argv)
    Case JSFN_INCLUDE: r = js_include(ctx, argc, argv)
    Case JSFN_INCLUDEJSON: r = js_includeJSON(ctx, argc, argv)
    Case JSFN_NAMESPACE: r = js_namespace(ctx, argc, argv)
    Case Else: r = jsCallApiFunction(callbackState, ctx, thisVal, argc, argv)
    End Select

    If r = 0 Then
        ' a throw (JS_EXCEPTION): the exception set is raised; none set, undefined
        If Not jsExceptionPending() Then r = jsUndefinedValue()
    Else
        ' a normal return: a leftover exception is dropped, as QuickJS never raises it
        jsClearException
    End If
    jsNativeFunction = r
    Exit Function
fail:
    errNumber = Err.Number
    errDescription = Err.Description
    debug_ 0, LOG_ERROR, "jsNativeFunction", "VB error %d in native function %d: %s", errNumber, callbackState, errDescription
    jsThrowReferenceError "VB error " & errNumber & ": " & errDescription
    jsNativeFunction = 0
End Function

' JS_ThrowTypeError
Private Sub jsThrowTypeError(ByVal message As String)
    Dim e As Long
    jsClearException
    JsCreateError jsStringValue(message), e
    JsSetException e
End Sub

' quickjs_execution_context execution_context(ctx): made for each call of a native function
Public Function wzapi_execution_context_new(ByVal ctx As Long) As wzapi_execution_context
    Dim c As wzapi_execution_context
    Set c = New wzapi_execution_context
    c.construct ctx
    Set wzapi_execution_context_new = c
End Function

' createQuickJSScriptInstance
Public Function createChakraScriptInstance(ByVal path As String, ByVal player As Long, ByVal difficulty As Long) As ChakraScriptingInstance
    Dim pNewInstance As ChakraScriptingInstance
    Set pNewInstance = New ChakraScriptingInstance
    pNewInstance.construct player, WzPathInfo_baseName(path), WzPathInfo_path(path)
    If Not pNewInstance.loadScript(path, player, difficulty) Then
        Set pNewInstance = Nothing      ' delete pNewInstance
        Set createChakraScriptInstance = Nothing
        Exit Function
    End If
    Set createChakraScriptInstance = pNewInstance
End Function

' QuickJS_EnumerateObjectProperties: the property names (strings) for the caller's loop; False where
' it failed
Public Function QuickJS_EnumerateObjectProperties(ByVal obj As Long, names As StringVector, Optional ByVal enumerableOnly As Boolean = True) As Boolean
    If Not (jsIsObject(obj)) Then ASSERT_FAILURE "JS_IsObject(obj)", "src/quickjs_backend.cpp:2718", "QuickJS_EnumerateObjectProperties", "obj is not a JS object": QuickJS_EnumerateObjectProperties = False: Exit Function
    If Not jsOwnPropertyNames(obj, enumerableOnly, names) Then
        ' JS_GetOwnPropertyNames failed?
        debug_ 2729, LOG_ERROR, "QuickJS_EnumerateObjectProperties", "JS_GetOwnPropertyNames failed"
        QuickJS_EnumerateObjectProperties = False
        Exit Function
    End If
    QuickJS_EnumerateObjectProperties = True
End Function

' wz_qjs_to_json with a fresh JSToJsonContext; False (and the std::runtime_error's what()) where it
' threw (a circular reference). isResetAssertLine, isResetAssertFunction: the line and function of the
' caller's ASSERT(toJsonContext.isReset(), ...)
Public Function wz_qjs_to_json_try(ByVal value As Long, ByVal skip_constructors As Boolean, result As nlohmann_json, what As String, ByVal isResetAssertLine As Long, ByVal isResetAssertFunction As String) As Boolean
    Dim c As JSToJsonContext
    c.skip_constructors = skip_constructors
    On Error GoTo threw
    Set result = wz_qjs_to_json(c, value)
    If Not (c.stack_size = 0) Then ASSERT_FAILURE "toJsonContext.isReset()", "src/quickjs_backend.cpp:" & isResetAssertLine, isResetAssertFunction, "JSToJsonContext has non-empty stack!"
    wz_qjs_to_json_try = True
    Exit Function
threw:
    If Err.Source <> "std::runtime_error" Then
        Dim n As Long, d As String
        n = Err.Number: d = Err.Description
        On Error GoTo 0
        Err.Raise n, , d
    End If
    what = Err.Description
    wz_qjs_to_json_try = False
End Function

' NOTE: May throw if there is a circular reference! (Err.Raise with Source "std::runtime_error")
Private Function wz_qjs_to_json(c As JSToJsonContext, ByVal value As Long) As nlohmann_json
    ' IMPORTANT: This largely follows the Qt documentation on QJsonValue::fromVariant
    ' See: http://doc.qt.io/qt-5/qjsonvalue.html#fromVariant
    '
    ' The main change is that numeric types are independently handled (instead of
    ' converting everything to `double`), because nlohmann::json handles them a bit
    ' differently.
    Dim j As nlohmann_json, i As Long, length As Long, k As Long, jsVal As Long, names As StringVector
    Dim nameStr As String, d As Double, s As String, ok As Boolean, typ As Long

    ' Note: Older versions of Qt 5.x (5.6?) do not define QMetaType::Nullptr,
    '		 so check value.isNull() instead.
    If jsIsNull(value) Then
        Set wz_qjs_to_json = json_value(Empty) ' null value
        Exit Function
    End If

    If jsIsObject(value) Then
        If jsIsConstructor(value) Then
            If Not c.skip_constructors Then Set wz_qjs_to_json = json_value("<constructor>") Else Set wz_qjs_to_json = json_value(Empty) ' null value
            Exit Function
        End If

        For i = 0 To c.stack_size - 1
            If jsSameObject(c.stack(i), value) Then
                ' circular reference
                Err.Raise vbObjectError + 2, "std::runtime_error", "Circular reference detected!"
            End If
        Next

        If c.stack_size = 0 Then
            ReDim c.stack(0 To 7)
        ElseIf c.stack_size > UBound(c.stack) Then
            ReDim Preserve c.stack(0 To c.stack_size * 2)
        End If
        c.stack(c.stack_size) = value
        c.stack_size = c.stack_size + 1

        If jsIsArray(value) Then
            ' Handle array
            Set j = json_array()
            length = 0
            If QuickJS_GetArrayLength(value, length) Then
                For k = 0 To length - 1 ' TODO: uint64_t isn't correct here, as we call GetUint32...
                    jsVal = jsGetIndex(value, k)
                    j.push_back_ref wz_qjs_to_json(c, jsVal)
                Next
            End If
        Else
            ' Handle actual objects
            Set j = json_object()
            QuickJS_EnumerateObjectProperties value, names, False
            For i = 0 To names.size - 1
                nameStr = names.items(i)
                If jsGetPropOk(value, nameStr, jsVal) Then
                    If Not jsIsConstructor(jsVal) Then
                        j.set_item_ref nameStr, wz_qjs_to_json(c, jsVal)
                    ElseIf Not c.skip_constructors Then
                        j.set_item_ref nameStr, json_value("<constructor>")
                    End If
                Else
                    debug_ 3763, LOG_INFO, "wz_qjs_to_json", "Got an exception trying to get the value of ""%s""?", nameStr
                End If
            Next
        End If

        c.stack_size = c.stack_size - 1
        Set wz_qjs_to_json = j
        Exit Function
    End If

    typ = jsValueType(value)
    Select Case typ
    Case JsBoolean
        Set wz_qjs_to_json = json_value(jsToBool(value))
    Case JsNumber
        ' (QuickJS's JS_TAG_INT / JS_TAG_FLOAT64: an integral number in int32's range, not -0, is an int)
        d = jsNumberValue(value)
        If isInt32Number(d) Then
            Set wz_qjs_to_json = json_value(jsDoubleToInt32(d))
        Else
            Set wz_qjs_to_json = json_value(d)
        End If
    Case JsUndefined
        Set wz_qjs_to_json = json_value(Empty) ' null value // ???
    Case JsString
        s = jsToStringOk(value, ok)
        If ok Then Set wz_qjs_to_json = json_value(s) Else Set wz_qjs_to_json = json_value("")
    Case Else
        ' In every other case, a conversion to a string will be attempted
        s = jsToStringOk(value, ok)
        ' If the returned string is empty, a Null QJsonValue will be stored, otherwise a String value using the returned QString.
        If ok Then
            If Len(s) = 0 Then
                Set wz_qjs_to_json = json_value(Empty) ' null value
            Else
                Set wz_qjs_to_json = json_value(s)
            End If
        Else
            Set wz_qjs_to_json = json_value(Empty) ' null value
        End If
    End Select
End Function

' (VB) whether QuickJS would tag the number an int: integral, in int32's range, and not -0
Private Function isInt32Number(ByVal d As Double) As Boolean
    Dim hi As Long
    If dblNotFinite(d) Then Exit Function
    If d <> Fix(d) Then Exit Function
    If d < -2147483648# Or d > 2147483647# Then Exit Function
    If d = 0 Then
        CopyMemory hi, ByVal VarPtr(d) + 4, 4
        If hi < 0 Then Exit Function    ' -0
    End If
    isInt32Number = True
End Function
