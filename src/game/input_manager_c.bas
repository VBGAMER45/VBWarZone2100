Attribute VB_Name = "input_manager_c"
' src/input/manager.cpp, src/input/manager.h: the input manager, which runs the key mappings' functions
' each frame (processMappings) and keeps the map markers.
' (VB) The one InputManager is display.cpp's gInputManager: its methods are gInputManager_<method>; its
' members are modules: contexts() input_context_c (gInputManager_contexts_<method>), mappings()
' input_mapping_c (gInputManager_mappings_<method>), debugManager() input_debugmappings_c
' (DebugInputManager_<method>). resetMappings's KeyFunctionConfiguration is gKeyFuncConfig
' (input_keyconfig_c). input_manager_c_init is display.cpp's two globals' constructors (InputManager
' gInputManager; KeyFunctionConfiguration gKeyFuncConfig;): Sub Main runs it.
Option Explicit

' std::unordered_map<KEY_CODE, KeyFunctionInfo, KeyCodeHash> markerKeyFunctions: the keys and their
' KeyFunctionInfo handles (only found, erased and emplaced: never iterated)
Private markerKeyFunctions_key() As Long
Private markerKeyFunctions_info() As Long
Private markerKeyFunctions_size As Long
Private bMappingsSortOrderDirty As Boolean      ' = true
Private currentKeyMapJsonPath As String

' static const KEY_CODE qwertyCodes[26]
Private qwertyCodes(0 To 25) As Long

' display.cpp's InputManager gInputManager; KeyFunctionConfiguration gKeyFuncConfig; (their constructors, in
' that order)
Public Sub input_manager_c_init()
    InputManager_construct
    KeyFunctionConfiguration_construct
End Sub

' InputManager::InputManager() (with the members' initialisers)
Public Sub InputManager_construct()
    KeyMappings_construct
    markerKeyFunctions_size = 0
    bMappingsSortOrderDirty = True
    currentKeyMapJsonPath = ""
    DebugInputManager_construct MAX_PLAYERS

    '  +---+   +---+   +---+   +---+   +---+   +---+   +---+   +---+   +---+   +---+
    qwertyCodes(0) = KEY_Q: qwertyCodes(1) = KEY_W: qwertyCodes(2) = KEY_E: qwertyCodes(3) = KEY_R: qwertyCodes(4) = KEY_T
    qwertyCodes(5) = KEY_Y: qwertyCodes(6) = KEY_U: qwertyCodes(7) = KEY_I: qwertyCodes(8) = KEY_O: qwertyCodes(9) = KEY_P
    '    +---+   +---+   +---+   +---+   +---+   +---+   +---+   +---+   +---+
    qwertyCodes(10) = KEY_A: qwertyCodes(11) = KEY_S: qwertyCodes(12) = KEY_D: qwertyCodes(13) = KEY_F: qwertyCodes(14) = KEY_G
    qwertyCodes(15) = KEY_H: qwertyCodes(16) = KEY_J: qwertyCodes(17) = KEY_K: qwertyCodes(18) = KEY_L
    '        +---+   +---+   +---+   +---+   +---+   +---+   +---+
    qwertyCodes(19) = KEY_Z: qwertyCodes(20) = KEY_X: qwertyCodes(21) = KEY_C: qwertyCodes(22) = KEY_V: qwertyCodes(23) = KEY_B
    qwertyCodes(24) = KEY_N: qwertyCodes(25) = KEY_M

    registerDefaultContexts
End Sub

' void InputManager::shutdown()
Public Sub gInputManager_shutdown()
    gInputManager_mappings_clear
End Sub

' void InputManager::setKeyMapJsonPath(const std::string& path)
Public Sub gInputManager_setKeyMapJsonPath(ByVal path As String)
    If path = currentKeyMapJsonPath Then
        Exit Sub
    End If
    currentKeyMapJsonPath = path
End Sub

' bool InputManager::mappingsSortRequired() const
Private Function mappingsSortRequired() As Boolean
    mappingsSortRequired = bMappingsSortOrderDirty Or gInputManager_contexts_isDirty() Or gInputManager_mappings_isDirty()
End Function

' (VB) markerKeyFunctions.clear()
Private Sub markerKeyFunctions_clear()
    Dim i As Long
    For i = 0 To markerKeyFunctions_size - 1
        KeyFunctionInfo_free markerKeyFunctions_info(i)
    Next
    markerKeyFunctions_size = 0
End Sub

' (VB) markerKeyFunctions.find(keyCode): the entry's position, -1 = end()
Private Function markerKeyFunctions_find(ByVal keyCode As Long) As Long
    Dim i As Long
    For i = 0 To markerKeyFunctions_size - 1
        If markerKeyFunctions_key(i) = keyCode Then
            markerKeyFunctions_find = i
            Exit Function
        End If
    Next
    markerKeyFunctions_find = -1
End Function

' void InputManager::resetMappings(bool bForceDefaults, const KeyFunctionConfiguration& keyFuncConfig) (gKeyFuncConfig)
Public Sub gInputManager_resetMappings(ByVal bForceDefaults As Boolean)
    Dim n As Long, allEntries() As Long, nEntries As Long, e As Long, info As Long, m As Long
    Dim slot As Long, keys As KeyCombination
    gInputManager_contexts_resetStates
    gInputManager_mappings_clear
    markerKeyFunctions_clear

    bMappingsSortOrderDirty = True
    For n = 0 To MAX_PLAYERS - 1
        DebugInputManager_setPlayerWantsDebugMappings n, False
    Next

    ' load the mappings.
    If Not bForceDefaults Then
        If gInputManager_mappings_load(currentKeyMapJsonPath) Then
            debug_ 98, LOG_WZ, "resetMappings", "Loaded key map successfully"
        Else
            debug_ 102, LOG_WARNING, "resetMappings", "Loading key map failed, forcing defaults"
            bForceDefaults = True
        End If
    End If

    ' Add in the default mappings if we are forcing defaults (e.g. "reset to defaults" button was pressed from the UI) or loading key map failed.
    nEntries = gKeyFuncConfig_allKeyFunctionEntries(allEntries)
    For e = 0 To nEntries - 1
        info = allEntries(e)
        For m = 0 To KeyFunctionInfo_defaultMappingsCount(info) - 1
            slot = KeyFunctionInfo_defaultMappingSlot(info, m)
            keys = KeyFunctionInfo_defaultMappingKeys(info, m)
            ' Always add non-assignable mappings as they are not saved.
            If bForceDefaults Or KeyFunctionInfo_type(info) <> KeyMappingType_ASSIGNABLE Then
                addDefaultMapping keys.meta, keys.input_source, keys.input_value, keys.action, info, slot
            End If
        Next
    Next

    gInputManager_mappings_save currentKeyMapJsonPath
End Sub

' void InputManager::saveMappings()
Public Sub gInputManager_saveMappings()
    debug_ 127, LOG_WZ, "saveMappings", "Saving: %s", currentKeyMapJsonPath
    gInputManager_mappings_save currentKeyMapJsonPath
End Sub

' bool InputManager::addDefaultMapping(const KEY_CODE metaCode, const KeyMappingInput input, const KeyAction action,
' const KeyFunctionInfo& info, const KeyMappingSlot slot = KeyMappingSlot::PRIMARY)
Private Function addDefaultMapping(ByVal metaCode As Long, ByVal inputSource As Long, ByVal inputValue As Long, ByVal action As Long, ByVal info As Long, Optional ByVal slot As Long = KeyMappingSlot_PRIMARY) As Boolean
    Dim psMapping As Long
    psMapping = gInputManager_mappings_get(info, slot)
    If psMapping <> 0 Then
        ' Remove any existing mapping for this function
        gInputManager_mappings_remove psMapping
    End If

    ' Clear the keys from any other mappings
    gInputManager_mappings_removeConflicting metaCode, inputSource, inputValue, KeyFunctionInfo_context(info)

    ' Set default key mapping
    gInputManager_mappings_add metaCode, inputSource, inputValue, action, info, slot
    addDefaultMapping = True
End Function

' ----------------------------------------------------------------------------------
' Returns the key code of the first ascii key that its finds has been PRESSED
Private Function getQwertyKey() As Long
    Dim i As Long
    For i = 0 To 25
        If keyPressed(qwertyCodes(i)) Then
            getQwertyKey = qwertyCodes(i) ' Top-, middle- or bottom-row key pressed.
            Exit Function
        End If
    Next

    getQwertyKey = 0 ' no ascii key pressed
End Function

' void InputManager::updateMapMarkers()
Public Sub gInputManager_updateMapMarkers()
    Dim qKey As Long, existing() As Long, nExisting As Long, i As Long, anyOther As Boolean, entry As Long, info As Long
    ' Are we trying to make a new map marker?
    If Not keyDown(KEY_LALT) Then
        Exit Sub
    End If

    ' Did we press a key
    qKey = getQwertyKey()
    If qKey = 0 Then
        Exit Sub
    End If

    nExisting = gInputManager_mappings_find(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, qKey, existing)
    If nExisting > 0 Then
        ' std::any_of(existing, [](const KeyMapping& mapping) { return mapping.info.name != "JumpToMapMarker"; })
        anyOther = False
        For i = 0 To nExisting - 1
            If KeyFunctionInfo_name(KeyMapping_info(existing(i))) <> "JumpToMapMarker" Then
                anyOther = True
                Exit For
            End If
        Next
        If anyOther Then
            Exit Sub
        End If
    End If

    For i = 0 To nExisting - 1
        If KeyFunctionInfo_name(KeyMapping_info(existing(i))) = "JumpToMapMarker" Then
            gInputManager_mappings_remove existing(i)
        End If
    Next

    ' Destroy any existing keymap entries for the key
    entry = markerKeyFunctions_find(qKey)
    If entry <> -1 Then
        KeyFunctionInfo_free markerKeyFunctions_info(entry)
        markerKeyFunctions_key(entry) = markerKeyFunctions_key(markerKeyFunctions_size - 1)
        markerKeyFunctions_info(entry) = markerKeyFunctions_info(markerKeyFunctions_size - 1)
        markerKeyFunctions_size = markerKeyFunctions_size - 1
    End If

    ' Create a new keymap entry. x/z/yaw are captured within the lambda in kf_JumpToMapMarker
    info = KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_HIDDEN, kf_JumpToMapMarker(playerPos.p.x, playerPos.p.z, playerPos.r.y), "JumpToMapMarker")
    ReDim Preserve markerKeyFunctions_key(0 To markerKeyFunctions_size)
    ReDim Preserve markerKeyFunctions_info(0 To markerKeyFunctions_size)
    markerKeyFunctions_key(markerKeyFunctions_size) = qKey
    markerKeyFunctions_info(markerKeyFunctions_size) = info
    markerKeyFunctions_size = markerKeyFunctions_size + 1
    entry = markerKeyFunctions_find(qKey)
    If entry <> -1 Then
        info = markerKeyFunctions_info(entry)
        gInputManager_mappings_add KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, qKey, KeyAction_PRESSED, info, KeyMappingSlot_PRIMARY
    End If
End Sub

' ----------------------------------------------------------------------------------
' allows checking if mapping should currently be ignored in processMappings
Private Function isIgnoredMapping(ByVal bAllowMouseWheelEvents As Boolean, ByVal mapping As Long, Optional ByVal bypassActiveContextCheck As Boolean = False) As Boolean
    Dim info As Long
    info = KeyMapping_info(mapping)
    If Not bypassActiveContextCheck Then
        If Not gInputManager_contexts_isActive(KeyFunctionInfo_context(info)) Then
            isIgnoredMapping = True
            Exit Function
        End If
    End If

    If KeyMapping_isInvalid(mapping) Then
        isIgnoredMapping = True
        Exit Function
    End If

    If Not bAllowMouseWheelEvents Then
        If KeyMappingInput_isMouse(KeyMapping_inputSource(mapping), KeyMapping_inputValue(mapping), MOUSE_WUP) Or KeyMappingInput_isMouse(KeyMapping_inputSource(mapping), KeyMapping_inputValue(mapping), MOUSE_WDN) Then
            isIgnoredMapping = True
            Exit Function
        End If
    End If

    If KeyFunctionInfo_function(info) Is Nothing Then
        isIgnoredMapping = True
        Exit Function
    End If

    If KeyFunctionInfo_bIsDebugOnly(info) And Not DebugInputManager_debugMappingsAllowed() And Not NETisReplay() Then
        isIgnoredMapping = True
        Exit Function
    End If

    isIgnoredMapping = False
End Function

' ----------------------------------------------------------------------------------
' Manages update of all the active function mappings
Public Sub gInputManager_processMappings(ByVal bAllowMouseWheelEvents As Boolean)
    Dim consumedSource() As Long, consumedValue() As Long, nConsumed As Long
    Dim i As Long, j As Long, keyToProcess As Long, bIsAlreadyConsumed As Boolean, inputSource As Long, inputValue As Long

    ' Check if player has made new camera markers
    gInputManager_updateMapMarkers

    ' If mappings have been updated or context priorities have changed, sort the mappings by priority and whether or not they have meta keys
    If mappingsSortRequired() Then
        gInputManager_mappings_sort
        gInputManager_contexts_clearDirty
        bMappingsSortOrderDirty = False
    End If
    ' std::unordered_set<KeyMappingInput, KeyMappingInput::Hash> consumedInputs;
    nConsumed = 0

    ' Run through all sorted mappings (VB: the list's size read again each time)
    i = 0
    Do While i < gInputManager_mappings_count()
        keyToProcess = gInputManager_mappings_at(i)
        ' Skip inappropriate ones when necessary
        If isIgnoredMapping(bAllowMouseWheelEvents, keyToProcess) Then
            GoTo continue_
        End If

        ' Skip if the input is already consumed. Handles skips for context/priority and meta-conflicts
        inputSource = KeyMapping_inputSource(keyToProcess)
        inputValue = KeyMapping_inputValue(keyToProcess)
        bIsAlreadyConsumed = False
        For j = 0 To nConsumed - 1
            If KeyMappingInput_equals(consumedSource(j), consumedValue(j), inputSource, inputValue) Then
                bIsAlreadyConsumed = True
                Exit For
            End If
        Next
        If bIsAlreadyConsumed Then
            GoTo continue_
        End If

        ' Execute the action if mapping was hit
        If KeyMapping_isActivated(keyToProcess) Then
            WzFunc_call KeyFunctionInfo_function(KeyMapping_info(keyToProcess))
            ReDim Preserve consumedSource(0 To nConsumed)
            ReDim Preserve consumedValue(0 To nConsumed)
            consumedSource(nConsumed) = inputSource
            consumedValue(nConsumed) = inputValue
            nConsumed = nConsumed + 1
        End If
continue_:
        i = i + 1
    Loop
End Sub

' nonstd::optional<std::reference_wrapper<const KeyMapping>> InputManager::findCurrentMapping(const bool bAllowMouseWheelEvents,
' bool filterHidden): the mapping's handle, 0 = nullopt
Public Function gInputManager_findCurrentMapping(ByVal bAllowMouseWheelEvents As Boolean, ByVal filterHidden As Boolean) As Long
    Dim i As Long, keyToProcess As Long
    ' If mappings have been updated or context priorities have changed, sort the mappings by priority and whether or not they have meta keys
    If mappingsSortRequired() Then
        gInputManager_mappings_sort
        gInputManager_contexts_clearDirty
        bMappingsSortOrderDirty = False
    End If

    ' Run through all sorted mappings
    For i = 0 To gInputManager_mappings_count() - 1
        keyToProcess = gInputManager_mappings_at(i)
        ' Skip inappropriate ones when necessary
        If isIgnoredMapping(bAllowMouseWheelEvents, keyToProcess, True) Then
            GoTo continue_
        End If
        If filterHidden And KeyFunctionInfo_type(KeyMapping_info(keyToProcess)) = KeyMappingType_HIDDEN Then
            GoTo continue_
        End If

        ' Execute the action if mapping was hit
        If KeyMapping_isActivated(keyToProcess) Then
            gInputManager_findCurrentMapping = keyToProcess
            Exit Function
        End If
continue_:
    Next

    gInputManager_findCurrentMapping = 0
End Function
