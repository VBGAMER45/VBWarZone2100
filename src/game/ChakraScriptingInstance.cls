VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChakraScriptingInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/quickjs_backend.cpp: class quickjs_scripting_instance, on Chakra, with
' its base class wzapi::scripting_instance (src/wzapi.h, src/wzapi.cpp: the
' constructor, isHostAI, loadFileForInclude, dumpScriptLog,
' debugGetScriptGlobalSpecialStringValues): one script (an AI, rules.js...)
' in its own JsRT context of the shared runtime (see chakra_backend_c's
' header). Made by createChakraScriptInstance (construct is the
' constructor); Class_Terminate is the destructor (when qtscript_c lets the
' last reference go: the context map holds a weak one).
' The lambdas log under their enclosing function's name (as netreplay_c's).
' Each method that runs JavaScript makes the instance's context current
' first and restores the previous one (jsEnterContext / jsLeaveContext).
' registerFunctions is chakra_functions_c's chakra_registerFunctions (round
' 2), js_builtin_funcs chakra_backend_c's setBuiltinFunctions.
' The handlers' parameters (event_nullable_ptr<> and optional<> of pointers:
' a handle, 0 = nullptr; optional<const char *>: has_ + the string) follow
' the C1 rules; a method with a UDT parameter is Friend.
Option Explicit

' wzapi::scripting_instance
Private m_player As Long
Private m_scriptName As String
Private m_scriptPath As String
Private m_isReceivingAllEvents As Boolean

Private rt As Long                      ' JSRuntime * (the shared runtime)
Private ctx As Long                     ' JSContext *
Private global_obj As Long              ' JsAddRef'd

Private compiledScriptObj As Long       ' JSValue compiledScriptObj = JS_UNINITIALIZED (0); JsAddRef'd
Private m_path As String
' Remember what names are used internally in the scripting engine, we don't want to save these to the savegame
Private internalNamespace As Object     ' std::unordered_set<std::string>: Scripting.Dictionary
' Separate event namespaces for libraries
Private eventNamespaces As StringVector
Private sourceTexts() As String         ' (VB) the texts and URLs given to JsParseScript, kept alive (unmoved) with the context
Private sourceUrls() As String
Private sources_size As Long
Private constructed As Boolean

Friend Sub construct(ByVal player As Long, ByVal scriptName As String, ByVal scriptPath As String)
    Dim prev As Long, c As Long
    m_player = player
    m_scriptName = scriptName
    m_scriptPath = scriptPath
    Set internalNamespace = CreateObject("Scripting.Dictionary")
    internalNamespace.CompareMode = 0       ' BinaryCompare

    rt = chakraRuntimeAcquire()
    If Not (rt <> 0) Then ASSERT_FAILURE "rt != nullptr", "src/quickjs_backend.cpp:191", "quickjs_scripting_instance", "JS_NewRuntime failed?"

    ' JSLimitedContextOptions: eval only for campaign (the runtime's JsRuntimeAttributeDisableEval), no
    ' promise (deleted below), no bigInt or weakRef (legacy Chakra has neither); the rest is Chakra's
    If rt <> 0 Then
        If JsCreateContext(rt, ctx) <> JsNoError Then ctx = 0
    End If
    If Not (ctx <> 0) Then ASSERT_FAILURE "ctx != nullptr", "src/quickjs_backend.cpp:212", "quickjs_scripting_instance", "JS_NewContext failed?"

    If ctx <> 0 Then
        JsAddRef ctx, c
        prev = jsEnterContext(ctx)
        JsGetGlobalObject global_obj
        jsValueAddRef global_obj
        jsDeleteProp global_obj, "Promise"  ' ctxOptions.promise = false: disable promise, async, await
        chakraCreateDescriptors ctx
        jsLeaveContext prev
        engineToInstanceMap_insert ctx, Me
    End If
    constructed = True
End Sub

Private Sub Class_Terminate()
    Dim prev As Long, c As Long
    If Not constructed Then Exit Sub
    engineToInstanceMap_erase ctx

    debug_ 222, LOG_INFO, "~quickjs_scripting_instance", "Destroying [%d]:%s/%s", player(), scriptPath(), scriptName()

    If ctx <> 0 Then
        prev = jsEnterContext(ctx)
        If compiledScriptObj <> 0 Then
            jsValueRelease compiledScriptObj
            compiledScriptObj = 0
        End If
        jsValueRelease global_obj
        global_obj = 0
        chakraReleaseDescriptors ctx
        If prev = ctx Then prev = 0
        jsLeaveContext prev
    End If
    If Not (ctx <> 0) Then ASSERT_FAILURE "ctx != nullptr", "src/quickjs_backend.cpp:231", "~quickjs_scripting_instance", "context is null??"
    If ctx <> 0 Then
        JsRelease ctx, c
        ctx = 0
    End If
    Erase sourceTexts
    Erase sourceUrls
    sources_size = 0
    If Not (rt <> 0) Then ASSERT_FAILURE "rt != nullptr", "src/quickjs_backend.cpp:237", "~quickjs_scripting_instance", "runtime is null??"
    If rt <> 0 Then
        chakraRuntimeRelease
        rt = 0
    End If
End Sub

'--------------------------------------------------------------------------
' wzapi::scripting_instance

Public Function scriptName() As String
    scriptName = m_scriptName
End Function

Public Function scriptPath() As String
    scriptPath = m_scriptPath
End Function

Public Function player() As Long
    player = m_player
End Function

Public Function isHostAI() As Boolean
    If Not (m_player >= 0 And m_player < NetPlay_players_size()) Then ASSERT_FAILURE "m_player >= 0 && m_player < NetPlay.players.size()", "src/wzapi.cpp:131", "isHostAI", "Invalid player: %d", m_player: isHostAI = False: Exit Function
    If Not NetPlay_isHost Then Exit Function
    If Not myResponsibility(m_player) Then Exit Function
    If isHumanPlayer(m_player) Then Exit Function
    isHostAI = (NetPlay_players_ai(m_player) >= 0 Or m_player = scavengerPlayer())
End Function

Public Sub setReceiveAllEvents(ByVal value As Boolean)
    m_isReceivingAllEvents = value
End Sub

Public Function isReceivingAllEvents() As Boolean
    isReceivingAllEvents = m_isReceivingAllEvents
End Function

' Loads a file.
' (Intended for use from implementations of things like "include" functions.)
'
' Lookup order is as follows (based on the value of `searchFlags`):
' - 1.) The filePath is checked relative to the read-only data dir search paths (LoadFileSearchOptions::DataDir)
' - 2.) The filePath is checked relative to "<user's config dir>/script/" (LoadFileSearchOptions::ConfigScriptDir)
' - 3.) The filename *only* is checked relative to the main scriptPath (LoadFileSearchOptions::ScriptPath_FileNameOnlyBackwardsCompat) - for backwards-compat only
' - 4.) The filePath is checked relative to the main scriptPath (LoadFileSearchOptions::ScriptPath)
' (ppFileData(0 .. pFileSize - 1))
Public Function loadFileForInclude(ByVal filePath As String, loadedFilePath As String, ppFileData() As Byte, pFileSize As Long, Optional ByVal searchFlags As Long = LoadFileSearchOptions_All) As Boolean
    Dim path As String, fileName As String, lastSlashPos As Long
    ' WzPathInfo filePathInfo = WzPathInfo::fromPlatformIndependentPath(filePath): filePath() is filePath
    lastSlashPos = InStrRev(filePath, "/")
    If lastSlashPos = 0 Then fileName = filePath Else fileName = Mid$(filePath, lastSlashPos + 1)
    path = ""

    If Len(path) = 0 And (searchFlags And LoadFileSearchOptions_DataDir) <> 0 Then
        If PHYSFS_exists(filePath) <> 0 Then
            path = filePath ' use this path instead (from read-only data dir)
        End If
    End If
    If Len(path) = 0 And (searchFlags And LoadFileSearchOptions_ConfigScriptDir) <> 0 Then
        If PHYSFS_exists("scripts/" & filePath) <> 0 Then
            path = "scripts/" & filePath ' use this path instead (in user write dir)
        End If
    End If
    If Len(path) = 0 And (searchFlags And LoadFileSearchOptions_ScriptPath_FileNameOnlyBackwardsCompat) <> 0 Then
        ' to provide backwards-compat, start by checking the scriptPath for the passed-in filename *only*
        path = scriptPath() & "/" & fileName
        If PHYSFS_exists(path) = 0 Then
            path = ""
        End If
    End If
    If Len(path) = 0 And (searchFlags And LoadFileSearchOptions_ScriptPath) <> 0 Then
        path = scriptPath() & "/" & filePath
        If PHYSFS_exists(path) = 0 Then
            path = ""
        End If
    End If
    If Len(path) = 0 Then
        debug_ 181, LOG_ERROR, "loadFileForInclude", "Failed to find file: %s", filePath
        Erase ppFileData
        pFileSize = 0
        loadFileForInclude = False
        Exit Function
    End If
    If Not loadFile(path, ppFileData, pFileSize) Then
        debug_ 188, LOG_ERROR, "loadFileForInclude", "Failed to read file ""%s"" (name=""%s"")", path, filePath
        Erase ppFileData
        pFileSize = 0
        loadFileForInclude = False
        Exit Function
    End If
    If Not isHostAI() Then
        calcDataHash ppFileData, pFileSize, DATA_SCRIPT
    End If

    loadedFilePath = path
    loadFileForInclude = True
End Function

' output to debug log file (dumpScriptLog(info) and dumpScriptLog(info, me))
Public Sub dumpScriptLog(ByVal info As String, Optional ByVal me_ As Variant)
    Dim path As String, fp As Long, m As Long
    If IsMissing(me_) Then m = player() Else m = CLng(me_)
    path = PHYSFS_getWriteDir()
    path = path & "/logs/" & scriptName() & "." & CStr(m) & ".log"
    fp = fopen(path, "a") ' TODO: This will fail for unicode paths on Windows. Should use PHYSFS to open / write files
    If fp <> 0 Then
        fputs info, fp
        fclose fp
    End If
End Sub

' std::unordered_map<std::string, DebugSpecialStringType> (quickjs_scripting_instance's override): a
' Scripting.Dictionary of name -> DebugSpecialStringType_*
Public Function debugGetScriptGlobalSpecialStringValues() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("<constructor>") = DebugSpecialStringType_TYPE_DESCRIPTION
    Set debugGetScriptGlobalSpecialStringValues = result
End Function

'--------------------------------------------------------------------------
' quickjs_scripting_instance

' (VB) for chakra_backend_c
Public Function Get_Global_Obj() As Long
    Get_Global_Obj = global_obj
End Function

Public Function eventNamespaces_size() As Long
    eventNamespaces_size = eventNamespaces.size
End Function

Public Function eventNamespaces_item(ByVal i As Long) As String
    eventNamespaces_item = eventNamespaces.items(i)
End Function

Public Sub eventNamespaces_push_back(ByVal prefix As String)
    StringVector_push_back eventNamespaces, prefix
End Sub

' (VB) keeps a script's text and URL alive (their buffers unmoved) as long as the context; returns
' the index for sourcePtr / urlPtr (the pointers JsParseScript takes)
Public Function keepSource(ByVal source As String, ByVal url As String) As Long
    If sources_size = 0 Then
        ReDim sourceTexts(0 To 7)
        ReDim sourceUrls(0 To 7)
    ElseIf sources_size > UBound(sourceTexts) Then
        ReDim Preserve sourceTexts(0 To sources_size * 2)
        ReDim Preserve sourceUrls(0 To sources_size * 2)
    End If
    sourceTexts(sources_size) = source
    sourceUrls(sources_size) = url
    keepSource = sources_size
    sources_size = sources_size + 1
End Function

Public Function sourcePtr(ByVal k As Long) As Long
    If Len(sourceTexts(k)) = 0 Then sourceTexts(k) = " "      ' (a null BSTR is no script)
    sourcePtr = StrPtr(sourceTexts(k))
End Function

Public Function urlPtr(ByVal k As Long) As Long
    If Len(sourceUrls(k)) = 0 Then sourceUrls(k) = " "
    urlPtr = StrPtr(sourceUrls(k))
End Function

' (VB) the context (JSContext *)
Public Function context() As Long
    context = ctx
End Function

Public Function loadScript(ByVal path As String, ByVal player As Long, ByVal difficulty As Long) As Boolean
    Dim size As Long, bytes() As Byte, source As String, errorAsString As String, prev As Long
    Dim names As StringVector, i As Long, k As Long
    If Not loadFile(path, bytes, size) Then
        debug_ 2776, LOG_ERROR, "loadScript", "Failed to read script file ""%s""", path
        loadScript = False
        Exit Function
    End If
    If Not isHostAI() Then
        calcDataHash bytes, size, DATA_SCRIPT
    End If
    m_path = path
    prev = jsEnterContext(ctx)
    source = scriptSourceText(bytes, size)
    k = keepSource(source, path)
    If JsParseScript(sourcePtr(k), jsNextSourceCookie(), urlPtr(k), compiledScriptObj) <> JsNoError Then
        ' compilation error / syntax error
        errorAsString = QuickJS_DumpError()
        debug_ 2790, LOG_ERROR, "loadScript", "Syntax / compilation error in %s: %s", path, errorAsString
        compiledScriptObj = 0
        jsLeaveContext prev
        loadScript = False
        Exit Function
    End If
    jsValueAddRef compiledScriptObj

    ' Special functions
    setBuiltinFunctions Me

    ' Regular functions
    If Not registerFunctions(WzPathInfo_baseName(path)) Then
        ' Failure to register functions - presumably OOM?
        debug_ 2805, LOG_ERROR, "loadScript", "Failed to register script functions for: %s", path
        jsValueRelease compiledScriptObj
        compiledScriptObj = 0
        jsLeaveContext prev
        loadScript = False
        Exit Function
    End If
    ' Remember internal, reserved names
    QuickJS_EnumerateObjectProperties global_obj, names, False
    For i = 0 To names.size - 1
        internalNamespace(names.items(i)) = True
    Next
    jsLeaveContext prev

    loadScript = True
End Function

Private Function registerFunctions(ByVal scriptName_ As String) As Boolean
    registerFunctions = chakra_registerFunctions(Me, scriptName_)
End Function

' lib/framework/wzpaths.cpp's WzPathInfo::baseName (fromPlatformIndependentPath(file))
Private Function WzPathInfo_baseName(ByVal file As String) As String
    Dim result As String, firstPeriodPos As Long, lastSlashPos As Long
    lastSlashPos = InStrRev(file, "/")
    If lastSlashPos = 0 Then result = file Else result = Mid$(file, lastSlashPos + 1)
    firstPeriodPos = InStr(result, ".")
    If firstPeriodPos = 0 Then
        WzPathInfo_baseName = result
        Exit Function
    End If
    WzPathInfo_baseName = Left$(result, firstPeriodPos - 1)
End Function

Public Function readyInstanceForExecution() As Boolean
    Dim result As Long, errorAsString As String, prev As Long, args() As Long
    If Not (compiledScriptObj <> 0) Then ASSERT_FAILURE "!JS_IsUninitialized(compiledScriptObj)", "src/quickjs_backend.cpp:2822", "readyInstanceForExecution", "compiledScriptObj is uninitialized": readyInstanceForExecution = False: Exit Function
    prev = jsEnterContext(ctx)
    If jsCall(compiledScriptObj, args, 0, result) <> JsNoError Then
        jsValueRelease compiledScriptObj
        compiledScriptObj = 0
        ' compilation error / syntax error
        errorAsString = QuickJS_DumpError()
        debug_ 2829, LOG_ERROR, "readyInstanceForExecution", "Syntax / compilation error in %s: %s", m_path, errorAsString
        jsLeaveContext prev
        readyInstanceForExecution = False
        Exit Function
    End If
    jsValueRelease compiledScriptObj
    compiledScriptObj = 0
    jsLeaveContext prev
    readyInstanceForExecution = True
End Function

' save / restore state
Public Function saveScriptGlobals(ByVal result As nlohmann_json) As Boolean
    Dim names As StringVector, i As Long, nameStr As String, jsVal As Long, j As nlohmann_json, what As String, prev As Long
    prev = jsEnterContext(ctx)
    ' we save 'scriptName' and 'me' implicitly
    QuickJS_EnumerateObjectProperties global_obj, names
    For i = 0 To names.size - 1
        nameStr = names.items(i)
        If jsGetPropOk(global_obj, nameStr, jsVal) Then
            If Not internalNamespace.Exists(nameStr) Then
                If Not jsIsFunction(jsVal) Then
                    If Not jsIsConstructor(jsVal) Then
                        '&& !it.value().equals(engine->globalObject()))
                        If wz_qjs_to_json_try(jsVal, True, j, what, 2855, "saveScriptGlobals") Then
                            result.set_item_ref nameStr, j
                        Else
                            debug_ 2857, LOG_ERROR, "saveScriptGlobals", "%s: Failed to convert global ""%s"" to json with error: %s", m_path, nameStr, what
                        End If
                    End If
                End If
            End If
        Else
            debug_ 2864, LOG_INFO, "saveScriptGlobals", "Got an exception trying to get the value of ""%s""?", nameStr
        End If
    Next
    jsLeaveContext prev
    saveScriptGlobals = True
End Function

Public Function loadScriptGlobals(ByVal result As nlohmann_json) As Boolean
    Dim it As Long, ret As Long, prev As Long, key As String
    If Not (result.is_object()) Then ASSERT_FAILURE "result.is_object()", "src/quickjs_backend.cpp:2873", "loadScriptGlobals", "Can't load script globals from non-json-object": loadScriptGlobals = False: Exit Function
    prev = jsEnterContext(ctx)
    For it = 0 To result.size() - 1
        ' IMPORTANT: "null" JSON values *MUST* map to JS_UNDEFINED.
        '			  If they are set to JS_NULL, it causes issues for libcampaign.js. (As the values become "defined".)
        '			  (mapJsonToQuickJSValue handles this properly.)
        ' NOTE: Properties created on the JS global object (as "global variables") should be non-configurable.
        '		 However *their* properties should probably be C_W_E.
        key = result.key_at(it)
        ret = jsDefineProp(global_obj, key, mapJsonToQuickJSValue(result.value_at(it), JS_PROP_C_W_E), JS_PROP_WRITABLE Or JS_PROP_ENUMERABLE Or JS_PROP_THROW)
        If ret <> 1 Then
            ' Failed to load script global
            If ret = -1 Then
                ' threw exception
                debug_ 2888, LOG_SCRIPT, "loadScriptGlobals", "[script: %s] Failed to load script global: ""%s""; %s", scriptName(), key, QuickJS_DumpError()
            Else
                ' other failure
                debug_ 2893, LOG_SCRIPT, "loadScriptGlobals", "[script: %s] Failed to load script global: ""%s""", scriptName(), key
            End If
        End If
    Next
    jsLeaveContext prev
    loadScriptGlobals = True
End Function

Public Function saveTimerFunction(ByVal timerID As Double, ByVal timerName As String, ByVal additionalParam As Object) As nlohmann_json
    Dim result As nlohmann_json, pData As ChakraTimerData
    Set result = json_object()
    result.set_item "function", json_value(timerName)
    If Not additionalParam Is Nothing Then
        If TypeOf additionalParam Is ChakraTimerData Then Set pData = additionalParam
    End If
    If Not pData Is Nothing Then
        result.set_item "stringArg", json_value(pData.stringArg)
    End If
    Set saveTimerFunction = result
End Function

' recreates timer functions (and additional userdata) based on the information saved by the saveTimer() method
' (the tuple: the TimerFunc returned, the timerAdditionalData in additionalParam; std::runtime_error: Err.Raise)
Public Function restoreTimerFunction(ByVal savedTimerFuncData As nlohmann_json, additionalParam As Object) As ITimerFunc
    Dim funcName As String, stringArg As String, timerFunc As ChakraTimerFunc, pData As ChakraTimerData
    funcName = json_getValue(savedTimerFuncData, "function").toWzString()
    If Len(funcName) = 0 Then
        Err.Raise vbObjectError + 2, "std::runtime_error", "Invalid timer restore data"
    End If
    stringArg = json_getValue(savedTimerFuncData, "stringArg").toWzString()

    ' timerFunc
    Set timerFunc = New ChakraTimerFunc
    timerFunc.init ctx, funcName
    ' additionalParams
    Set pData = New ChakraTimerData
    pData.stringArg = stringArg
    Set additionalParam = pData
    Set restoreTimerFunction = timerFunc
End Function

' get state for debugging
Public Function debugGetAllScriptGlobals() As nlohmann_json
    Dim globals As nlohmann_json, names As StringVector, i As Long, nameStr As String, jsVal As Long
    Dim j As nlohmann_json, what As String, prev As Long, take As Boolean
    Set globals = json_object()
    prev = jsEnterContext(ctx)
    QuickJS_EnumerateObjectProperties global_obj, names
    For i = 0 To names.size - 1
        nameStr = names.items(i)
        jsGetPropOk global_obj, nameStr, jsVal
        take = False
        If Not internalNamespace.Exists(nameStr) Then
            If Not jsIsFunction(jsVal) Then take = True
            '/*&& !it.value().equals(engine->globalObject())*/)
        End If
        If nameStr = "Upgrades" Or nameStr = "Stats" Then take = True
        If take Then
            If wz_qjs_to_json_try(jsVal, False, j, what, 2960, "debugGetAllScriptGlobals") Then
                globals.set_item_ref nameStr, j
            Else
                debug_ 2962, LOG_ERROR, "debugGetAllScriptGlobals", "[D] %s: Failed to convert global ""%s"" to json with error: %s", m_path, nameStr, what
                globals.set_item nameStr, json_value("<ERROR: OBJECT WITH CIRCULAR REFERENCES>")
            End If
        End If
    Next
    jsLeaveContext prev
    Set debugGetAllScriptGlobals = globals
End Function

Public Function debugEvaluateCommand(ByVal text As String) As Boolean
    Dim compiledFuncObj As Long, result As Long, errorAsString As String, resultStr As String, prev As Long
    Dim args() As Long, k As Long
    prev = jsEnterContext(ctx)
    k = keepSource(text, "<debug_evaluate_command>")
    If JsParseScript(sourcePtr(k), jsNextSourceCookie(), urlPtr(k), compiledFuncObj) <> JsNoError Then
        ' compilation error / syntax error
        errorAsString = QuickJS_DumpError()
        debug_ 2986, LOG_ERROR, "debugEvaluateCommand", "Syntax error in %s: %s", text, errorAsString
        jsLeaveContext prev
        debugEvaluateCommand = False
        Exit Function
    End If
    If jsCall(compiledFuncObj, args, 0, result) <> JsNoError Then
        ' compilation error / syntax error
        errorAsString = QuickJS_DumpError()
        debug_ 2998, LOG_ERROR, "debugEvaluateCommand", "Uncaught exception in %s: %s", text, errorAsString
        jsLeaveContext prev
        debugEvaluateCommand = False
        Exit Function
    End If
    resultStr = JSValueToStdString(result)
    console "%s", resultStr
    jsLeaveContext prev
    debugEvaluateCommand = True
End Function

Public Sub updateGameTime(ByVal newGameTime As Long)
    Dim ret As Long, prev As Long
    prev = jsEnterContext(ctx)
    ret = jsDefineProp(global_obj, "gameTime", jsUint(newGameTime), JS_PROP_WRITABLE Or JS_PROP_ENUMERABLE)
    If Not (ret >= 1) Then ASSERT_FAILURE "ret >= 1", "src/quickjs_backend.cpp:3012", "updateGameTime", "Failed to update gameTime"
    jsLeaveContext prev
End Sub

Public Sub updateGroupSizes(ByVal groupId As Long, ByVal size As Long)
    Dim groupMembersObj As Long, prev As Long
    prev = jsEnterContext(ctx)
    groupMembersObj = jsGetProp(global_obj, "groupSizes")
    jsDefineProp groupMembersObj, CStr(groupId), jsInt(size), JS_PROP_C_W_E
    jsLeaveContext prev
End Sub

Private Function toQuickJSPropertyFlags(ByVal flags As Long) As Long
    Dim propertyFlags As Long
    propertyFlags = 0
    If ((flags And GlobalVariableFlags_ReadOnly) <> GlobalVariableFlags_ReadOnly) _
        Or ((flags And GlobalVariableFlags_ReadOnlyUpdatedFromApp) = GlobalVariableFlags_ReadOnlyUpdatedFromApp) Then
        ' NOTE: QuickJS prevents even the app - not just scripts - from updating unwritable / read-only variables
        ' So if ReadOnlyUpdatedFromApp is set, we must set the variable to writable to support updates from WZ
        propertyFlags = propertyFlags Or JS_PROP_WRITABLE
    End If
    toQuickJSPropertyFlags = propertyFlags
End Function

' set "global" variables (flags: GlobalVariableFlags_*; default ReadOnly | DoNotSave)
Public Sub setSpecifiedGlobalVariables(ByVal variables As nlohmann_json, Optional ByVal flags As Long = -1)
    Dim propertyFlags As Long, markGlobalAsInternal As Boolean, it As Long, key As String, prev As Long
    If flags = -1 Then flags = GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave
    If Not variables.is_object() Then
        ASSERT_FAILURE "false", "src/quickjs_backend.cpp:3026", "setSpecifiedGlobalVariables", "setSpecifiedGlobalVariables expects a JSON object"
        Exit Sub
    End If
    propertyFlags = toQuickJSPropertyFlags(flags) Or JS_PROP_ENUMERABLE
    markGlobalAsInternal = (flags And GlobalVariableFlags_DoNotSave) = GlobalVariableFlags_DoNotSave
    prev = jsEnterContext(ctx)
    For it = 0 To variables.size() - 1
        key = variables.key_at(it)
        If Not (Len(key) <> 0) Then ASSERT_FAILURE "!it.key().empty()", "src/quickjs_backend.cpp:3033", "setSpecifiedGlobalVariables", "Empty key"
        jsDefineProp global_obj, key, mapJsonToQuickJSValue(variables.value_at(it), propertyFlags), propertyFlags
        If markGlobalAsInternal Then
            internalNamespace(key) = True
        End If
    Next
    jsLeaveContext prev
End Sub

Public Sub setSpecifiedGlobalVariable(ByVal name_ As String, ByVal value As nlohmann_json, Optional ByVal flags As Long = -1)
    Dim propertyFlags As Long, markGlobalAsInternal As Boolean, prev As Long
    If flags = -1 Then flags = GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave
    If Not (Len(name_) <> 0) Then ASSERT_FAILURE "!name.empty()", "src/quickjs_backend.cpp:3044", "setSpecifiedGlobalVariable", "Empty key"
    propertyFlags = toQuickJSPropertyFlags(flags) Or JS_PROP_ENUMERABLE
    markGlobalAsInternal = (flags And GlobalVariableFlags_DoNotSave) = GlobalVariableFlags_DoNotSave
    prev = jsEnterContext(ctx)
    jsDefineProp global_obj, name_, mapJsonToQuickJSValue(value, propertyFlags), propertyFlags
    jsLeaveContext prev
    If markGlobalAsInternal Then
        internalNamespace(name_) = True
    End If
End Sub

Public Sub doNotSaveGlobal(ByVal global_ As String)
    internalNamespace(global_) = True
End Sub

'--------------------------------------------------------------------------
' The event handlers (IMPL_EVENT_HANDLER, IMPL_EVENT_HANDLER_NO_PARAMS): each argument boxed, then
' callFunction(ctx, "<event>", args) (wrap_event_handler__). The caller's context is restored.

Private Function wrap_event_handler__(ByVal functionName As String, args() As Long, ByVal n As Long) As Boolean
    '/*JSValue result =*/
    callFunction ctx, functionName, args, n
    jsArgsRelease args, n
    wrap_event_handler__ = True '//nlohmann::json(result.toVariant());
End Function

' (VB) an event without arguments
Private Function eventNoParams(ByVal functionName As String) As Boolean
    Dim a() As Long, prev As Long
    prev = jsEnterContext(ctx)
    eventNoParams = wrap_event_handler__(functionName, a, 0)
    jsLeaveContext prev
End Function

' (VB) an event with one object argument (box(const BASE_OBJECT *))
Private Function eventObject(ByVal functionName As String, ByVal psObj As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_BASE_OBJECT(psObj)
    eventObject = wrap_event_handler__(functionName, a, n)
    jsLeaveContext prev
End Function

' MARK: General events
Public Function handle_eventGameInit() As Boolean
    handle_eventGameInit = eventNoParams("eventGameInit")
End Function

Public Function handle_eventStartLevel() As Boolean
    handle_eventStartLevel = eventNoParams("eventStartLevel")
End Function

Public Function handle_eventMissionTimeout() As Boolean
    handle_eventMissionTimeout = eventNoParams("eventMissionTimeout")
End Function

Public Function handle_eventVideoDone() As Boolean
    handle_eventVideoDone = eventNoParams("eventVideoDone")
End Function

Public Function handle_eventGameLoaded() As Boolean
    handle_eventGameLoaded = eventNoParams("eventGameLoaded")
End Function

Public Function handle_eventGameSaving() As Boolean
    handle_eventGameSaving = eventNoParams("eventGameSaving")
End Function

Public Function handle_eventGameSaved() As Boolean
    handle_eventGameSaved = eventNoParams("eventGameSaved")
End Function

' MARK: Transporter events
Public Function handle_eventLaunchTransporter() As Boolean ' DEPRECATED!
    handle_eventLaunchTransporter = eventNoParams("eventLaunchTransporter")
End Function

Public Function handle_eventTransporterLaunch(ByVal psTransport As Long) As Boolean
    handle_eventTransporterLaunch = eventObject("eventTransporterLaunch", psTransport)
End Function

Public Function handle_eventReinforcementsArrived() As Boolean ' DEPRECATED!
    handle_eventReinforcementsArrived = eventNoParams("eventReinforcementsArrived")
End Function

Public Function handle_eventTransporterArrived(ByVal psTransport As Long) As Boolean
    handle_eventTransporterArrived = eventObject("eventTransporterArrived", psTransport)
End Function

Public Function handle_eventTransporterExit(ByVal psObj As Long) As Boolean
    handle_eventTransporterExit = eventObject("eventTransporterExit", psObj)
End Function

Public Function handle_eventTransporterDone(ByVal psTransport As Long) As Boolean
    handle_eventTransporterDone = eventObject("eventTransporterDone", psTransport)
End Function

Public Function handle_eventTransporterLanded(ByVal psTransport As Long) As Boolean
    handle_eventTransporterLanded = eventObject("eventTransporterLanded", psTransport)
End Function

Public Function handle_eventTransporterEmbarked(ByVal psTransport As Long) As Boolean
    handle_eventTransporterEmbarked = eventObject("eventTransporterEmbarked", psTransport)
End Function

Public Function handle_eventTransporterDisembarked(ByVal psTransport As Long) As Boolean
    handle_eventTransporterDisembarked = eventObject("eventTransporterDisembarked", psTransport)
End Function

' MARK: UI-related events (intended for the tutorial)
Public Function handle_eventDeliveryPointMoving(ByVal psStruct As Long) As Boolean
    handle_eventDeliveryPointMoving = eventObject("eventDeliveryPointMoving", psStruct)
End Function

Public Function handle_eventDeliveryPointMoved(ByVal psStruct As Long) As Boolean
    handle_eventDeliveryPointMoved = eventObject("eventDeliveryPointMoved", psStruct)
End Function

Public Function handle_eventDesignBody() As Boolean
    handle_eventDesignBody = eventNoParams("eventDesignBody")
End Function

Public Function handle_eventDesignPropulsion() As Boolean
    handle_eventDesignPropulsion = eventNoParams("eventDesignPropulsion")
End Function

Public Function handle_eventDesignWeapon() As Boolean
    handle_eventDesignWeapon = eventNoParams("eventDesignWeapon")
End Function

Public Function handle_eventDesignCommand() As Boolean
    handle_eventDesignCommand = eventNoParams("eventDesignCommand")
End Function

Public Function handle_eventDesignSystem() As Boolean
    handle_eventDesignSystem = eventNoParams("eventDesignSystem")
End Function

Public Function handle_eventDesignQuit() As Boolean
    handle_eventDesignQuit = eventNoParams("eventDesignQuit")
End Function

Public Function handle_eventMenuBuildSelected() As Boolean
    handle_eventMenuBuildSelected = eventNoParams("eventMenuBuildSelected")
End Function

Public Function handle_eventMenuResearchSelected() As Boolean
    handle_eventMenuResearchSelected = eventNoParams("eventMenuResearchSelected")
End Function

Public Function handle_eventMenuBuild() As Boolean
    handle_eventMenuBuild = eventNoParams("eventMenuBuild")
End Function

Public Function handle_eventMenuResearch() As Boolean
    handle_eventMenuResearch = eventNoParams("eventMenuResearch")
End Function

Public Function handle_eventMenuDesign() As Boolean
    handle_eventMenuDesign = eventNoParams("eventMenuDesign")
End Function

Public Function handle_eventMenuManufacture() As Boolean
    handle_eventMenuManufacture = eventNoParams("eventMenuManufacture")
End Function

' (const std::vector<const BASE_OBJECT *>& objects: objects(0 .. objects_size - 1))
Public Function handle_eventSelectionChanged(objects() As Long, ByVal objects_size As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_vector_BASE_OBJECT(objects, objects_size)
    handle_eventSelectionChanged = wrap_event_handler__("eventSelectionChanged", a, n)
    jsLeaveContext prev
End Function

' MARK: Game state-change events
Public Function handle_eventObjectRecycled(ByVal psObj As Long) As Boolean
    handle_eventObjectRecycled = eventObject("eventObjectRecycled", psObj)
End Function

Public Function handle_eventPlayerLeft(ByVal player As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_int(player)
    handle_eventPlayerLeft = wrap_event_handler__("eventPlayerLeft", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventCheatMode(ByVal entered As Boolean) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_bool(entered)
    handle_eventCheatMode = wrap_event_handler__("eventCheatMode", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventDroidIdle(ByVal psDroid As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_DROID(psDroid)
    handle_eventDroidIdle = wrap_event_handler__("eventDroidIdle", a, n)
    jsLeaveContext prev
End Function

' (optional<const STRUCTURE *> psFactory: has_psFactory, psFactory; nullopt is undefined)
Public Function handle_eventDroidBuilt(ByVal psDroid As Long, ByVal has_psFactory As Boolean, ByVal psFactory As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_DROID(psDroid)
    If has_psFactory Then jsArgsPush a, n, box_STRUCTURE(psFactory) Else jsArgsPush a, n, jsUndefinedValue()
    handle_eventDroidBuilt = wrap_event_handler__("eventDroidBuilt", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventStructureBuilt(ByVal psStruct As Long, ByVal has_psDroid As Boolean, ByVal psDroid As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_STRUCTURE(psStruct)
    If has_psDroid Then jsArgsPush a, n, box_DROID(psDroid) Else jsArgsPush a, n, jsUndefinedValue()
    handle_eventStructureBuilt = wrap_event_handler__("eventStructureBuilt", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventStructureDemolish(ByVal psStruct As Long, ByVal has_psDroid As Boolean, ByVal psDroid As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_STRUCTURE(psStruct)
    If has_psDroid Then jsArgsPush a, n, box_DROID(psDroid) Else jsArgsPush a, n, jsUndefinedValue()
    handle_eventStructureDemolish = wrap_event_handler__("eventStructureDemolish", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventStructureReady(ByVal psStruct As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_STRUCTURE(psStruct)
    handle_eventStructureReady = wrap_event_handler__("eventStructureReady", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventStructureUpgradeStarted(ByVal psStruct As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_STRUCTURE(psStruct)
    handle_eventStructureUpgradeStarted = wrap_event_handler__("eventStructureUpgradeStarted", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventDroidRankGained(ByVal psDroid As Long, ByVal rankNum As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_DROID(psDroid)
    jsArgsPush a, n, box_int(rankNum)
    handle_eventDroidRankGained = wrap_event_handler__("eventDroidRankGained", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventAttacked(ByVal psVictim As Long, ByVal psAttacker As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_BASE_OBJECT(psVictim)
    jsArgsPush a, n, box_BASE_OBJECT(psAttacker)
    handle_eventAttacked = wrap_event_handler__("eventAttacked", a, n)
    jsLeaveContext prev
End Function

' (wzapi::event_nullable_ptr<const STRUCTURE> psStruct: a handle, 0 = null)
Friend Function handle_eventResearched(research_ As wzapi_researchResult, ByVal psStruct As Long, ByVal player As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_researchResult(research_)
    jsArgsPush a, n, box_STRUCTURE(psStruct)
    jsArgsPush a, n, box_int(player)
    handle_eventResearched = wrap_event_handler__("eventResearched", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventDestroyed(ByVal psVictim As Long) As Boolean
    handle_eventDestroyed = eventObject("eventDestroyed", psVictim)
End Function

Public Function handle_eventPickup(ByVal psFeat As Long, ByVal psDroid As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_FEATURE(psFeat)
    jsArgsPush a, n, box_DROID(psDroid)
    handle_eventPickup = wrap_event_handler__("eventPickup", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventObjectSeen(ByVal psViewer As Long, ByVal psSeen As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_BASE_OBJECT(psViewer)
    jsArgsPush a, n, box_BASE_OBJECT(psSeen)
    handle_eventObjectSeen = wrap_event_handler__("eventObjectSeen", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventGroupSeen(ByVal psViewer As Long, ByVal groupId As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_BASE_OBJECT(psViewer)
    jsArgsPush a, n, box_int(groupId)
    handle_eventGroupSeen = wrap_event_handler__("eventGroupSeen", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventObjectTransfer(ByVal psObj As Long, ByVal from As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_BASE_OBJECT(psObj)
    jsArgsPush a, n, box_int(from)
    handle_eventObjectTransfer = wrap_event_handler__("eventObjectTransfer", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventChat(ByVal from As Long, ByVal to_ As Long, ByVal message As String) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_int(from)
    jsArgsPush a, n, box_int(to_)
    jsArgsPush a, n, box_cstring(message)
    handle_eventChat = wrap_event_handler__("eventChat", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventQuickChat(ByVal from As Long, ByVal to_ As Long, ByVal messageEnum As Long, ByVal teamSpecific As Boolean) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_int(from)
    jsArgsPush a, n, box_int(to_)
    jsArgsPush a, n, box_int(messageEnum)
    jsArgsPush a, n, box_bool(teamSpecific)
    handle_eventQuickChat = wrap_event_handler__("eventQuickChat", a, n)
    jsLeaveContext prev
End Function

' (optional<const char *> message: has_message, message; nullopt is undefined)
Public Function handle_eventBeacon(ByVal x As Long, ByVal y As Long, ByVal from As Long, ByVal to_ As Long, ByVal has_message As Boolean, ByVal message As String) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_int(x)
    jsArgsPush a, n, box_int(y)
    jsArgsPush a, n, box_int(from)
    jsArgsPush a, n, box_int(to_)
    If has_message Then jsArgsPush a, n, box_cstring(message) Else jsArgsPush a, n, jsUndefinedValue()
    handle_eventBeacon = wrap_event_handler__("eventBeacon", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventBeaconRemoved(ByVal from As Long, ByVal to_ As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_int(from)
    jsArgsPush a, n, box_int(to_)
    handle_eventBeaconRemoved = wrap_event_handler__("eventBeaconRemoved", a, n)
    jsLeaveContext prev
End Function

' // Since groups are entities local to one context, we do not iterate over them here.
Public Function handle_eventGroupLoss(ByVal psObj As Long, ByVal group As Long, ByVal size As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_BASE_OBJECT(psObj)
    jsArgsPush a, n, box_int(group)
    jsArgsPush a, n, box_int(size)
    handle_eventGroupLoss = wrap_event_handler__("eventGroupLoss", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventArea(ByVal label As String, ByVal psDroid As Long) As Boolean
    Dim args() As Long, n As Long, funcname As String, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush args, n, convDroid(psDroid)
    funcname = "eventArea" & label
    debug_ 3124, LOG_SCRIPT, "handle_eventArea", "Triggering %s for %s", funcname, scriptName()
    callFunction ctx, funcname, args, n
    jsArgsRelease args, n
    jsLeaveContext prev
    handle_eventArea = True
End Function

Public Function handle_eventDesignCreated(ByVal psTemplate As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_DROID_TEMPLATE(psTemplate)
    handle_eventDesignCreated = wrap_event_handler__("eventDesignCreated", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventAllianceOffer(ByVal from As Long, ByVal to_ As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_uchar(from)
    jsArgsPush a, n, box_uchar(to_)
    handle_eventAllianceOffer = wrap_event_handler__("eventAllianceOffer", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventAllianceAccepted(ByVal from As Long, ByVal to_ As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_uchar(from)
    jsArgsPush a, n, box_uchar(to_)
    handle_eventAllianceAccepted = wrap_event_handler__("eventAllianceAccepted", a, n)
    jsLeaveContext prev
End Function

Public Function handle_eventAllianceBroken(ByVal from As Long, ByVal to_ As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_uchar(from)
    jsArgsPush a, n, box_uchar(to_)
    handle_eventAllianceBroken = wrap_event_handler__("eventAllianceBroken", a, n)
    jsLeaveContext prev
End Function

' MARK: Special input events
Public Function handle_eventSyncRequest(ByVal from As Long, ByVal req_id As Long, ByVal x As Long, ByVal y As Long, ByVal psObj As Long, ByVal psObj2 As Long) As Boolean
    Dim a() As Long, n As Long, prev As Long
    prev = jsEnterContext(ctx)
    jsArgsPush a, n, box_int(from)
    jsArgsPush a, n, box_int(req_id)
    jsArgsPush a, n, box_int(x)
    jsArgsPush a, n, box_int(y)
    jsArgsPush a, n, box_BASE_OBJECT(psObj)
    jsArgsPush a, n, box_BASE_OBJECT(psObj2)
    handle_eventSyncRequest = wrap_event_handler__("eventSyncRequest", a, n)
    jsLeaveContext prev
End Function
