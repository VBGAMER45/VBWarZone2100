Attribute VB_Name = "input_mapping_c"
' src/input/mapping.cpp, src/input/mapping.h: the key mappings (a key combination bound to a key function)
' and their list, gInputManager's KeyMappings.
' (VB) A KeyMapping& is a handle (1-based, 0 = nullopt) into the pool aKeyMapping, valid until the mapping is
' removed (as the std::list element's address); its fields are KeyMapping_<field>(h) (keys.input:
' KeyMapping_inputSource / KeyMapping_inputValue). The one KeyMappings is display.cpp's gInputManager's
' (InputManager's mappings()): its methods are gInputManager_mappings_<method>; the std::list is
' keyMappings(), the handles in the list's order (gInputManager_mappings_count / _at iterate it). A
' std::vector<std::reference_wrapper<KeyMapping>> result is its count, the handles in result(0 .. count - 1).
' The methods taking a KeyCombination take its meta, input (source, value) and action: the callers'
' braced { meta, input, action } is the constructor KeyCombination_ the method runs.
Option Explicit

' struct KeyMapping
Private Type KeyMapping
    info As Long                ' const KeyFunctionInfo& (its handle)
    lastCalled As Long          ' UDWORD
    keys As KeyCombination
    slot As Long                ' KeyMappingSlot
    inUse As Boolean            ' (VB) the pool's
End Type

Private aKeyMapping() As KeyMapping
Private aKeyMapping_size As Long
Private kmFree() As Long
Private kmFree_size As Long

' class KeyMappings: std::list<KeyMapping> keyMappings (the handles, in order); bool bDirty = true
Private keyMappings() As Long
Private keyMappings_size As Long
Private bDirty As Boolean

' (VB) the member initialiser bool bDirty = true (InputManager_construct runs it, as C++ constructs
' gInputManager's keyMappings)
Public Sub KeyMappings_construct()
    bDirty = True
End Sub

Private Function isCombination(ByVal mapping As Long) As Boolean
    isCombination = aKeyMapping(mapping).keys.meta <> KEY_IGNORE
End Function

Private Function isActiveSingleKey(ByVal mapping As Long) As Boolean
    With aKeyMapping(mapping).keys
        Select Case .action
        Case KeyAction_PRESSED
            isActiveSingleKey = KeyMappingInput_isPressed(.input_source, .input_value)
        Case KeyAction_DOWN
            isActiveSingleKey = KeyMappingInput_isDown(.input_source, .input_value)
        Case KeyAction_RELEASED
            isActiveSingleKey = KeyMappingInput_isReleased(.input_source, .input_value)
        Case Else
            debug_ 45, LOG_WARNING, "isActiveSingleKey", "Unknown key action (action code %u) while processing keymap.", .action
            isActiveSingleKey = False
        End Select
    End With
End Function

Private Function getAlternativeForMetaKey(ByVal meta As Long) As Long
    Dim altMeta As Long
    altMeta = KEY_IGNORE
    If meta = KEY_LCTRL Then
        altMeta = KEY_RCTRL
    ElseIf meta = KEY_LALT Then
        altMeta = KEY_RALT
    ElseIf meta = KEY_LSHIFT Then
        altMeta = KEY_RSHIFT
    ElseIf meta = KEY_LMETA Then
        altMeta = KEY_RMETA
    End If

    getAlternativeForMetaKey = altMeta
End Function

Private Function isActiveCombination(ByVal mapping As Long) As Boolean
    Dim bSubKeyIsPressed As Boolean, bMetaIsDown As Boolean, altMeta As Long, bHasAlt As Boolean, bAltMetaIsDown As Boolean
    If Not (KeyMapping_hasMeta(mapping)) Then ASSERT_FAILURE "mapping.hasMeta()", "src/input/mapping.cpp:75", "isActiveCombination", "isActiveCombination called for non-meta key mapping!"

    With aKeyMapping(mapping).keys
        bSubKeyIsPressed = KeyMappingInput_isPressed(.input_source, .input_value)
        bMetaIsDown = keyDown(.meta)

        altMeta = getAlternativeForMetaKey(.meta)
    End With
    bHasAlt = altMeta <> KEY_IGNORE
    bAltMetaIsDown = False
    If bHasAlt Then bAltMetaIsDown = keyDown(altMeta)

    isActiveCombination = bSubKeyIsPressed And (bMetaIsDown Or bAltMetaIsDown)
End Function

' bool KeyMapping::isInvalid() const
Public Function KeyMapping_isInvalid(ByVal mapping As Long) As Boolean
    KeyMapping_isInvalid = KeyMappingInput_is(aKeyMapping(mapping).keys.input_source, aKeyMapping(mapping).keys.input_value, KEY_MAXSCAN)
End Function

' bool KeyMapping::isActivated() const
Public Function KeyMapping_isActivated(ByVal mapping As Long) As Boolean
    If isCombination(mapping) Then
        KeyMapping_isActivated = isActiveCombination(mapping)
    Else
        KeyMapping_isActivated = isActiveSingleKey(mapping)
    End If
End Function

' bool KeyMapping::hasMeta() const
Public Function KeyMapping_hasMeta(ByVal mapping As Long) As Boolean
    KeyMapping_hasMeta = aKeyMapping(mapping).keys.meta <> KEY_IGNORE
End Function

' std::string KeyMapping::toString() const (keyScanToString / mouseKeyCodeToString return the name, the
' 20-character buffers' 19 characters)
Public Function KeyMapping_toString(ByVal mapping As Long) As String
    Dim asciiSub As String, asciiMeta As String
    ' Figure out if the keycode is for mouse or keyboard and print the name of
    ' the respective key/mouse button to `asciiSub`
    asciiSub = ""
    With aKeyMapping(mapping).keys
        Select Case .input_source
        Case KeyMappingInputSource_KEY_CODE
            asciiSub = Left$(keyScanToString(.input_value), 19)
        Case KeyMappingInputSource_MOUSE_KEY_CODE
            asciiSub = Left$(mouseKeyCodeToString(.input_value), 19)
        Case Else
            debug_ 118, LOG_WZ, "toString", "Encountered invalid key mapping source %u while converting mapping to string!", .input_source
            KeyMapping_toString = "NOT VALID"
            Exit Function
        End Select

        If KeyMapping_hasMeta(mapping) Then
            asciiMeta = Left$(keyScanToString(.meta), 19)

            KeyMapping_toString = asciiMeta & " " & asciiSub
        Else
            KeyMapping_toString = asciiSub
        End If
    End With
End Function

' bool operator==(const KeyMapping& lhs, const KeyMapping& rhs)
Public Function KeyMapping_equals(ByVal lhs As Long, ByVal rhs As Long) As Boolean
    KeyMapping_equals = KeyCombination_equals(aKeyMapping(lhs).keys, aKeyMapping(rhs).keys) And aKeyMapping(lhs).slot = aKeyMapping(rhs).slot And aKeyMapping(lhs).info = aKeyMapping(rhs).info ' Assume infos are immutable with only one copy existing at a time.
End Function

' the fields
Public Function KeyMapping_info(ByVal mapping As Long) As Long
    KeyMapping_info = aKeyMapping(mapping).info
End Function

Public Function KeyMapping_lastCalled(ByVal mapping As Long) As Long
    KeyMapping_lastCalled = aKeyMapping(mapping).lastCalled
End Function

Public Function KeyMapping_keys(ByVal mapping As Long) As KeyCombination
    KeyMapping_keys = aKeyMapping(mapping).keys
End Function

Public Function KeyMapping_meta(ByVal mapping As Long) As Long
    KeyMapping_meta = aKeyMapping(mapping).keys.meta
End Function

Public Function KeyMapping_inputSource(ByVal mapping As Long) As Long
    KeyMapping_inputSource = aKeyMapping(mapping).keys.input_source
End Function

Public Function KeyMapping_inputValue(ByVal mapping As Long) As Long
    KeyMapping_inputValue = aKeyMapping(mapping).keys.input_value
End Function

Public Function KeyMapping_action(ByVal mapping As Long) As Long
    KeyMapping_action = aKeyMapping(mapping).keys.action
End Function

Public Function KeyMapping_slot(ByVal mapping As Long) As Long
    KeyMapping_slot = aKeyMapping(mapping).slot
End Function

' mapping.keys.input.isCleared()
Public Function KeyMapping_inputIsCleared(ByVal mapping As Long) As Boolean
    KeyMapping_inputIsCleared = KeyMappingInput_isCleared(aKeyMapping(mapping).keys.input_source, aKeyMapping(mapping).keys.input_value)
End Function

' (VB) the list's iteration (InputManager's range-based fors): its size, the handle at position i
Public Function gInputManager_mappings_count() As Long
    gInputManager_mappings_count = keyMappings_size
End Function

Public Function gInputManager_mappings_at(ByVal i As Long) As Long
    gInputManager_mappings_at = keyMappings(i)
End Function

' (VB) a new list element's pool slot
Private Function KeyMapping_alloc() As Long
    Dim h As Long
    If kmFree_size > 0 Then
        kmFree_size = kmFree_size - 1
        h = kmFree(kmFree_size)
    Else
        aKeyMapping_size = aKeyMapping_size + 1
        h = aKeyMapping_size
        ReDim Preserve aKeyMapping(1 To aKeyMapping_size)
    End If
    aKeyMapping(h).inUse = True
    KeyMapping_alloc = h
End Function

' (VB) erases the list element at position pos (its handle dies)
Private Sub keyMappings_erase(ByVal pos As Long)
    Dim h As Long, i As Long, blank As KeyMapping
    h = keyMappings(pos)
    For i = pos To keyMappings_size - 2
        keyMappings(i) = keyMappings(i + 1)
    Next
    keyMappings_size = keyMappings_size - 1
    aKeyMapping(h) = blank
    ReDim Preserve kmFree(0 To kmFree_size)
    kmFree(kmFree_size) = h
    kmFree_size = kmFree_size + 1
End Sub

' KeyMapping& KeyMappings::add(const KeyCombination keys, const KeyFunctionInfo& info, const KeyMappingSlot slot): keys
' is KeyCombination{ meta, input, action }
Public Function gInputManager_mappings_add(ByVal meta As Long, ByVal inputSource As Long, ByVal inputValue As Long, ByVal action As Long, ByVal info As Long, ByVal slot As Long) As Long
    Dim keys As KeyCombination, leftMeta As Long, keysWithLeftMeta As KeyCombination, h As Long
    keys = KeyCombination_(meta, inputSource, inputValue, action)

    ' Make sure the meta key is the left variant
    leftMeta = keys.meta
    If keys.meta = KEY_RCTRL Then
        leftMeta = KEY_LCTRL
    ElseIf keys.meta = KEY_RALT Then
        leftMeta = KEY_LALT
    ElseIf keys.meta = KEY_RSHIFT Then
        leftMeta = KEY_LSHIFT
    ElseIf keys.meta = KEY_RMETA Then
        leftMeta = KEY_LMETA
    End If

    ' Create the mapping as the last element in the list
    keysWithLeftMeta = KeyCombination_(leftMeta, keys.input_source, keys.input_value, keys.action)
    h = KeyMapping_alloc()
    aKeyMapping(h).info = info
    aKeyMapping(h).lastCalled = gameTime
    aKeyMapping(h).keys = keysWithLeftMeta
    aKeyMapping(h).slot = slot
    ReDim Preserve keyMappings(0 To keyMappings_size)
    keyMappings(keyMappings_size) = h
    keyMappings_size = keyMappings_size + 1

    ' Invalidate the sorting order and return the newly created mapping
    bDirty = True
    gInputManager_mappings_add = h
End Function

' nonstd::optional<std::reference_wrapper<KeyMapping>> KeyMappings::get(const std::string& name, const KeyMappingSlot slot)
Public Function gInputManager_mappings_get_name(ByVal name_ As String, ByVal slot As Long) As Long
    Dim i As Long, mapping As Long
    For i = 0 To keyMappings_size - 1
        mapping = keyMappings(i)
        If KeyFunctionInfo_name(aKeyMapping(mapping).info) = name_ And aKeyMapping(mapping).slot = slot Then
            gInputManager_mappings_get_name = mapping
            Exit Function
        End If
    Next

    gInputManager_mappings_get_name = 0
End Function

' nonstd::optional<std::reference_wrapper<KeyMapping>> KeyMappings::get(const KeyFunctionInfo& info, const KeyMappingSlot slot)
Public Function gInputManager_mappings_get(ByVal info As Long, ByVal slot As Long) As Long
    gInputManager_mappings_get = gInputManager_mappings_get_name(KeyFunctionInfo_name(info), slot)
End Function

' std::vector<std::reference_wrapper<KeyMapping>> KeyMappings::find(const KEY_CODE meta, const KeyMappingInput input)
Public Function gInputManager_mappings_find(ByVal meta As Long, ByVal inputSource As Long, ByVal inputValue As Long, matches() As Long) As Long
    Dim i As Long, mapping As Long, n As Long
    Erase matches
    n = 0
    For i = 0 To keyMappings_size - 1
        mapping = keyMappings(i)
        If aKeyMapping(mapping).keys.meta = meta And KeyMappingInput_equals(aKeyMapping(mapping).keys.input_source, aKeyMapping(mapping).keys.input_value, inputSource, inputValue) Then
            ReDim Preserve matches(0 To n)
            matches(n) = mapping
            n = n + 1
        End If
    Next

    gInputManager_mappings_find = n
End Function

' bool KeyMappings::remove(const KeyMapping& mappingToRemove): the first mapping equal to it
Public Function gInputManager_mappings_remove(ByVal mappingToRemove As Long) As Boolean
    Dim i As Long
    For i = 0 To keyMappings_size - 1
        If KeyMapping_equals(keyMappings(i), mappingToRemove) Then
            keyMappings_erase i
            bDirty = True
            gInputManager_mappings_remove = True
            Exit Function
        End If
    Next
    gInputManager_mappings_remove = False
End Function

' std::vector<std::reference_wrapper<KeyMapping>> KeyMappings::findConflicting(const KEY_CODE meta, const KeyMappingInput input,
' const ContextId contextId, const ContextManager& contexts)
Public Function gInputManager_mappings_findConflicting(ByVal meta As Long, ByVal inputSource As Long, ByVal inputValue As Long, ByVal contextId As String, conflicts() As Long) As Long
    Dim matches() As Long, nMatches As Long, i As Long, mapping As Long, mappingContext As Long, bConflicts As Boolean, n As Long
    ' Find any mapping with same keys
    nMatches = gInputManager_mappings_find(meta, inputSource, inputValue, matches)
    Erase conflicts
    n = 0
    For i = 0 To nMatches - 1
        mapping = matches(i)
        mappingContext = gInputManager_contexts_get(KeyFunctionInfo_context(aKeyMapping(mapping).info))

        ' Keys conflict if they are for the same context. Always active contexts are special (always have highest priority), so those keys will always conflict.
        bConflicts = InputContext_isAlwaysActive(mappingContext) Or KeyFunctionInfo_context(aKeyMapping(mapping).info) = contextId
        If bConflicts Then
            ReDim Preserve conflicts(0 To n)
            conflicts(n) = mapping
            n = n + 1
        End If
    Next

    gInputManager_mappings_findConflicting = n
End Function

' std::vector<KeyMapping> KeyMappings::removeConflicting(const KEY_CODE meta, const KeyMappingInput input, const ContextId& contextId,
' const ContextManager& contexts): the number removed (the callers drop the copies)
Public Function gInputManager_mappings_removeConflicting(ByVal meta As Long, ByVal inputSource As Long, ByVal inputValue As Long, ByVal contextId As String) As Long
    Dim conflicting() As Long, nConflicting As Long, i As Long, mapping As Long, removed As Long
    ' Find any mapping with same keys
    nConflicting = gInputManager_mappings_findConflicting(meta, inputSource, inputValue, contextId, conflicting)
    removed = 0
    For i = 0 To nConflicting - 1
        mapping = conflicting(i)
        ' Clear only if the mapping is for an assignable binding
        If aKeyMapping(mapping).inUse Then
            If KeyFunctionInfo_type(aKeyMapping(mapping).info) = KeyMappingType_ASSIGNABLE Then
                removed = removed + 1
                gInputManager_mappings_remove mapping
            End If
        End If
    Next

    gInputManager_mappings_removeConflicting = removed
End Function

' bool KeyMappings::isDirty() const
Public Function gInputManager_mappings_isDirty() As Boolean
    gInputManager_mappings_isDirty = bDirty
End Function

' void KeyMappings::clear(nonstd::optional<KeyMappingType> filter = nonstd::nullopt): filter -1 is nullopt
Public Sub gInputManager_mappings_clear(Optional ByVal filter As Long = -1)
    Dim i As Long
    If filter = -1 Then
        For i = keyMappings_size - 1 To 0 Step -1
            keyMappings_erase i
        Next
    Else
        ' keyMappings.remove_if([filter](const KeyMapping& mapping) { return mapping.info.type == filter.value(); });
        For i = keyMappings_size - 1 To 0 Step -1
            If KeyFunctionInfo_type(aKeyMapping(keyMappings(i)).info) = filter Then
                keyMappings_erase i
            End If
        Next
    End If

    bDirty = True
End Sub

' void KeyMappings::sort(const ContextManager& contexts): std::list::sort, stable, by the lambda: mappings with meta first,
' then by priority, highest first. (VB) an insertion sort, the keys (hasMeta, the context's priority) read once
Public Sub gInputManager_mappings_sort()
    Dim keyMeta() As Boolean, keyPriority() As Long, i As Long, j As Long
    Dim xMapping As Long, xMeta As Boolean, xPriority As Long, less As Boolean
    If keyMappings_size < 2 Then
        bDirty = False
        Exit Sub
    End If
    ReDim keyMeta(0 To keyMappings_size - 1)
    ReDim keyPriority(0 To keyMappings_size - 1)
    For i = 0 To keyMappings_size - 1
        keyMeta(i) = KeyMapping_hasMeta(keyMappings(i))
        keyPriority(i) = gInputManager_contexts_getPriority(KeyFunctionInfo_context(aKeyMapping(keyMappings(i)).info))
    Next
    For i = 1 To keyMappings_size - 1
        xMapping = keyMappings(i)
        xMeta = keyMeta(i)
        xPriority = keyPriority(i)
        j = i - 1
        Do While j >= 0
            ' Sort by meta. This causes all mappings with meta to be checked before non-meta mappings,
            ' avoiding having to check for meta-conflicts in the processing loop. (e.g. if we should execute
            ' a mapping with right arrow key, depending on if another binding on shift+right-arrow is executed
            ' or not). In other words, if any mapping with meta is executed, it will consume the respective input,
            ' preventing any non-meta mappings with the same input from being executed.
            If xMeta <> keyMeta(j) Then
                less = xMeta And Not keyMeta(j)
            Else
                ' Sort by priority
                less = ULT(keyPriority(j), xPriority)
            End If
            If Not less Then Exit Do
            keyMappings(j + 1) = keyMappings(j)
            keyMeta(j + 1) = keyMeta(j)
            keyPriority(j + 1) = keyPriority(j)
            j = j - 1
        Loop
        keyMappings(j + 1) = xMapping
        keyMeta(j + 1) = xMeta
        keyPriority(j + 1) = xPriority
    Next
    bDirty = False
End Sub

' static KeyMappingInput createInputForSource(const KeyMappingInputSource source, const unsigned int keyCode): the input in
' inputSource, inputValue
Private Sub createInputForSource(ByVal source As Long, ByVal keyCode As Long, inputSource As Long, inputValue As Long)
    Select Case source
    Case KeyMappingInputSource_KEY_CODE
        inputSource = KeyMappingInputSource_KEY_CODE
        inputValue = keyCode
    Case KeyMappingInputSource_MOUSE_KEY_CODE
        inputSource = KeyMappingInputSource_MOUSE_KEY_CODE
        inputValue = keyCode
    Case Else
        debug_ 319, LOG_WZ, "createInputForSource", "Encountered invalid key mapping source %u while loading keymap!", source
        inputSource = KeyMappingInputSource_KEY_CODE
        inputValue = KEY_MAXSCAN
    End Select
End Sub

' bool KeyMappings::load(const char* path, const KeyFunctionConfiguration& keyFuncConfig) (gKeyFuncConfig's)
Public Function gInputManager_mappings_load(ByVal path As String) As Boolean
    Dim ini As WzConfig, meta As Long, sub_ As Long, action As Long, functionName As String, info As Long
    Dim sourceName As String, source As Long, inputSource As Long, inputValue As Long, slotName As String, slot As Long
    ' Clear all assignable mappings
    gInputManager_mappings_clear KeyMappingType_ASSIGNABLE

    Set ini = WzConfig_new(path, WzConfig_ReadOnly)
    If Not ini.status() Then
        debug_ 332, LOG_WZ, "load", "%s not found", path
        gInputManager_mappings_load = False
        Exit Function
    End If

    ini.beginArray "mappings"
    Do While ini.remainingArrayItems() <> 0
        meta = ini.value("meta", json_variant_new(0)).toInt()
        sub_ = ini.value("sub", json_variant_new(0)).toInt()
        action = ini.value("action", json_variant_new(0)).toInt()
        functionName = ini.value("function", json_variant_new("")).toWzString()
        info = gKeyFuncConfig_keyFunctionInfoByName(functionName)
        If info = 0 Then
            debug_ 345, LOG_WARNING, "load", "Skipping unknown keymap function ""%s"".", functionName
            GoTo continue_
        ElseIf KeyFunctionInfo_type(info) <> KeyMappingType_ASSIGNABLE Then
            ' No need to load non-assignable mappings
            debug_ 351, LOG_WARNING, "load", "Skipping non-assignable keymap function ""%s"".", functionName
            GoTo continue_
        End If

        sourceName = ini.value("source", json_variant_new("default")).toWzString()
        source = keyMappingSourceByName(sourceName)
        createInputForSource source, sub_, inputSource, inputValue

        slotName = ini.value("slot", json_variant_new("primary")).toWzString()
        slot = keyMappingSlotByName(slotName)

        gInputManager_mappings_add meta, inputSource, inputValue, action, info, slot
continue_:
        ini.nextArrayItem
    Loop
    ini.endArray
    gInputManager_mappings_load = True
End Function

' bool KeyMappings::save(const char* path) const
Public Function gInputManager_mappings_save(ByVal path As String) As Boolean
    Dim ini As WzConfig, i As Long, mapping As Long
    Set ini = WzConfig_new(path, WzConfig_ReadAndWrite)
    If Not ini.status() Or Not ini.isWritable() Then
        ' NOTE: Changed to LOG_FATAL, since we want to inform user via pop-up (windows only)
        debug_ 374, LOG_FATAL, "save", "Could not open %s", ini.fileName()
        gInputManager_mappings_save = False
        Exit Function
    End If

    ini.setValue "version", json_value(2)

    ini.beginArray "mappings"
    For i = 0 To keyMappings_size - 1
        mapping = keyMappings(i)
        ' No need to save non-assignable mappings
        If KeyFunctionInfo_type(aKeyMapping(mapping).info) <> KeyMappingType_ASSIGNABLE Then
            GoTo continue_
        End If

        ini.setValue "name", json_value(KeyFunctionInfo_name(aKeyMapping(mapping).info))
        ini.setValue "meta", json_value(aKeyMapping(mapping).keys.meta)

        Select Case aKeyMapping(mapping).keys.input_source
        Case KeyMappingInputSource_KEY_CODE
            ini.setValue "source", json_value("default")
            ini.setValue "sub", json_value(aKeyMapping(mapping).keys.input_value)
        Case KeyMappingInputSource_MOUSE_KEY_CODE
            ini.setValue "source", json_value("mouse_key")
            ini.setValue "sub", json_value(aKeyMapping(mapping).keys.input_value)
        Case Else
            debug_ 402, LOG_WZ, "save", "Encountered invalid key mapping source %u while saving keymap!", aKeyMapping(mapping).keys.input_source
        End Select
        Select Case aKeyMapping(mapping).slot
        Case KeyMappingSlot_PRIMARY
            ini.setValue "slot", json_value("primary")
        Case KeyMappingSlot_SECONDARY
            ini.setValue "slot", json_value("secondary")
        Case Else
            debug_ 414, LOG_WZ, "save", "Encountered invalid key mapping slot %u while saving keymap!", aKeyMapping(mapping).slot
        End Select

        ini.setValue "action", json_value(aKeyMapping(mapping).keys.action)
        ini.setValue "function", json_value(KeyFunctionInfo_name(aKeyMapping(mapping).info))

        ini.nextArrayItem
continue_:
    Next
    ini.endArray

    debug_ 425, LOG_WZ, "save", "Keymap written ok to %s.", path
    gInputManager_mappings_save = True
End Function
