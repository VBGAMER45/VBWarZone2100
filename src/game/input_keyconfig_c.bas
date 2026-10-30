Attribute VB_Name = "input_keyconfig_c"
' src/input/keyconfig.cpp, src/input/keyconfig.h: the mappable key functions (KeyFunctionInfo: the table of
' all of them, gKeyFuncConfig's), the key inputs and combinations.
' (VB) A const KeyFunctionInfo& is a handle (1-based, 0 = none) into the pool aKeyFunctionInfo: the
' table's entries first, then the map markers' (InputManager's markerKeyFunctions, freed when erased).
' Its fields are KeyFunctionInfo_<field>(h); defaultMappings is KeyFunctionInfo_defaultMappingsCount(h)
' and KeyFunctionInfo_defaultMapping<Slot|Meta|InputSource|InputValue|Action|InputIsCleared>(h, i).
' A KeyMappingInput is its source and value (two Longs: KeyMappingInput_<method>(source, value)); a
' KeyCombination a UDT. A MappableFunction (std::function<void()>) is a WzFunc: the plain kf_ functions
' are keybind_c's FN_kf_<name>, the factories (kf_Zoom(...)) return one. The one KeyFunctionConfiguration
' is display.cpp's gKeyFuncConfig: KeyFunctionConfiguration_construct is its constructor (Sub Main runs it,
' as C++ runs the global's), its methods gKeyFuncConfig_<method>.
Option Explicit

' enum class KeyMappingInputSource
Public Const KeyMappingInputSource_KEY_CODE As Long = 1
Public Const KeyMappingInputSource_MOUSE_KEY_CODE As Long = 2

' enum class KeyAction
Public Const KeyAction_DOWN As Long = 0
Public Const KeyAction_PRESSED As Long = 1
Public Const KeyAction_RELEASED As Long = 2

' enum class KeyMappingType
' The mapping value can be freely re-assigned.
Public Const KeyMappingType_ASSIGNABLE As Long = 0
' The mapping value is visible, but cannot be re-assigned.
Public Const KeyMappingType_FIXED As Long = 1
' The mapping does not show up in the key map options
Public Const KeyMappingType_HIDDEN As Long = 2

' enum class KeyMappingSlot
Public Const KeyMappingSlot_PRIMARY As Long = 0
Public Const KeyMappingSlot_SECONDARY As Long = 1
Public Const KeyMappingSlot_LAST As Long = 2

' struct KeyCombination (its KeyMappingInput input: input_source, input_value)
Public Type KeyCombination
    meta As Long            ' KEY_CODE
    input_source As Long    ' KeyMappingInputSource
    input_value As Long     ' KeyMappingInputValue: KEY_CODE or MOUSE_KEY_CODE
    action As Long          ' KeyAction
End Type

' (VB) the most default mappings a key function has (the table's)
Private Const KEYFUNCTION_MAX_DEFAULTS As Long = 2

' std::pair<KeyMappingSlot, KeyCombination>
Private Type KeyFunctionInfo_defaultMapping
    first As Long
    second As KeyCombination
End Type

' struct KeyFunctionInfo
Private Type KeyFunctionInfo
    context As String                   ' const ContextId&
    typ As Long                         ' const KeyMappingType
    function_ As WzFunc                 ' const MappableFunction
    name_ As String
    displayName As String
    defaultMappings(0 To KEYFUNCTION_MAX_DEFAULTS - 1) As KeyFunctionInfo_defaultMapping
    defaultMappings_size As Long
    bIsDebugOnly As Boolean
    inUse As Boolean                    ' (VB) the pool's
End Type

' (VB) the KeyFunctionInfos: the table's and the map markers', a free list for the markers'
Private aKeyFunctionInfo() As KeyFunctionInfo
Private aKeyFunctionInfo_size As Long
Private kfiFree() As Long
Private kfiFree_size As Long

' class KeyFunctionConfiguration (gKeyFuncConfig): std::vector<KeyFunctionInfo> entries (their handles);
' nameToIndexMap is keyFunctionInfoByName's search
Private entries() As Long
Private entries_size As Long

' MARK: - KeyFunctionConfiguration

' (VB) initializeKeyFunctions's entries.emplace_back(KeyFunctionInfo(...)): adds the info's handle, returns it
Private Function entries_emplace_back(ByVal info As Long) As Long
    ReDim Preserve entries(0 To entries_size)
    entries(entries_size) = info
    entries_size = entries_size + 1
    entries_emplace_back = info
End Function

' Definitions/Configuration for all mappable Key Functions
'
' NOTE: The initialization is done as a function with bunch of emplace_backs instead of an initializer list for two reasons:
'        1.) KeyFunctionInfo is marked as non-copy to avoid unnecessarily copying them around. Using an initializer list requires
'            types to be copyable, so we cannot use initializer lists, at all (we use move-semantics)
'        2.) The initializer list itself would require >20kb of stack memory due to sheer size of this thing. Inserting all
'            entries one-by-one requires only one entry on the stack at a time, mitigating the risk of a stack overflow.
' (VB) the entries are in four procedures (VB's procedure size); KeyFunctionInfo_new and its
' defaultMappings_push_back are the constructor's arguments
Private Sub initializeKeyFunctions()
    initializeKeyFunctions_1
    initializeKeyFunctions_2
    initializeKeyFunctions_3
    initializeKeyFunctions_4
End Sub

' (VB) initializeKeyFunctions's entries, part 1
Private Sub initializeKeyFunctions_1()
    Dim e As Long
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_ALWAYS_ACTIVE, KeyMappingType_FIXED, WzFunc_new(FN_kf_ChooseManufacture), "ChooseManufacture", "Manufacture"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F1, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_ALWAYS_ACTIVE, KeyMappingType_FIXED, WzFunc_new(FN_kf_ChooseResearch), "ChooseResearch", "Research"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F2, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_ALWAYS_ACTIVE, KeyMappingType_FIXED, WzFunc_new(FN_kf_ChooseBuild), "ChooseBuild", "Build"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F3, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_ALWAYS_ACTIVE, KeyMappingType_FIXED, WzFunc_new(FN_kf_ChooseDesign), "ChooseDesign", "Design"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F4, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_ALWAYS_ACTIVE, KeyMappingType_FIXED, WzFunc_new(FN_kf_ChooseIntelligence), "ChooseIntelligence", "Intelligence Display"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F5, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_ALWAYS_ACTIVE, KeyMappingType_FIXED, WzFunc_new(FN_kf_ChooseCommand), "ChooseCommand", "Commanders"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F6, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_QuickSave), "QuickSave", "QuickSave"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F7, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleRadar), "ToggleRadar", "Toggle Radar"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_F7, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_QuickLoad), "QuickLoad", "QuickLoad"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F8, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleConsole), "ToggleConsole", "Toggle Console Display"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_F8, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleEnergyBars), "ToggleEnergyBars", "Toggle Damage Bars On/Off"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F9, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_BACKGROUND, KeyMappingType_FIXED, WzFunc_new(FN_kf_ScreenDump), "ScreenDump", "Take Screen Shot"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F10, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleFormationSpeedLimiting), "ToggleFormationSpeedLimiting", "Toggle Formation Speed Limiting"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F11, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_MoveToLastMessagePos), "MoveToLastMessagePos", "View Location of Previous Message"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F12, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleSensorDisplay), "ToggleSensorDisplay", "Toggle Sensor display"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_F12, KeyAction_PRESSED)
    ' ASSIGN GROUPS
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(0), "AssignGrouping_0", "Assign Group 0"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_0, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(1), "AssignGrouping_1", "Assign Group 1"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_1, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(2), "AssignGrouping_2", "Assign Group 2"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_2, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(3), "AssignGrouping_3", "Assign Group 3"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_3, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(4), "AssignGrouping_4", "Assign Group 4"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_4, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(5), "AssignGrouping_5", "Assign Group 5"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_5, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(6), "AssignGrouping_6", "Assign Group 6"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_6, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(7), "AssignGrouping_7", "Assign Group 7"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_7, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(8), "AssignGrouping_8", "Assign Group 8"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_8, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AssignGrouping_N(9), "AssignGrouping_9", "Assign Group 9"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_9, KeyAction_PRESSED)
    ' ADD TO GROUP
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(0), "AddGrouping_0", "Add to Group 0"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_0, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(1), "AddGrouping_1", "Add to Group 1"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_1, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(2), "AddGrouping_2", "Add to Group 2"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_2, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(3), "AddGrouping_3", "Add to Group 3"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_3, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(4), "AddGrouping_4", "Add to Group 4"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_4, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(5), "AddGrouping_5", "Add to Group 5"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_5, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(6), "AddGrouping_6", "Add to Group 6"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_6, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(7), "AddGrouping_7", "Add to Group 7"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_7, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(8), "AddGrouping_8", "Add to Group 8"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_8, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_AddGrouping_N(9), "AddGrouping_9", "Add to Group 9"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_9, KeyAction_PRESSED)
    ' REMOVE FROM GROUP
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_RemoveFromGrouping(), "RemoveFromGrouping", "Remove from current Group"))

    ' SELECT GROUPS - Will jump to the group as well as select if group is ALREADY selected
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(0), "SelectGrouping_0", "Select Group 0"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_0, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(1), "SelectGrouping_1", "Select Group 1"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_1, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(2), "SelectGrouping_2", "Select Group 2"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_2, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(3), "SelectGrouping_3", "Select Group 3"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_3, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(4), "SelectGrouping_4", "Select Group 4"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_4, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(5), "SelectGrouping_5", "Select Group 5"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_5, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(6), "SelectGrouping_6", "Select Group 6"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_6, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(7), "SelectGrouping_7", "Select Group 7"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_7, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(8), "SelectGrouping_8", "Select Group 8"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_8, KeyAction_PRESSED)
End Sub

' (VB) initializeKeyFunctions's entries, part 2
Private Sub initializeKeyFunctions_2()
    Dim e As Long
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectGrouping_N(9), "SelectGrouping_9", "Select Group 9"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_9, KeyAction_PRESSED)
    ' SELECT COMMANDER - Will jump to the group as well as select if group is ALREADY selected
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(0), "SelectCommander_0", "Select Commander 0"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_0, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(1), "SelectCommander_1", "Select Commander 1"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_1, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(2), "SelectCommander_2", "Select Commander 2"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_2, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(3), "SelectCommander_3", "Select Commander 3"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_3, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(4), "SelectCommander_4", "Select Commander 4"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_4, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(5), "SelectCommander_5", "Select Commander 5"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_5, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(6), "SelectCommander_6", "Select Commander 6"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_6, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(7), "SelectCommander_7", "Select Commander 7"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_7, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(8), "SelectCommander_8", "Select Commander 8"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_8, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectCommander_N(9), "SelectCommander_9", "Select Commander 9"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_9, KeyAction_PRESSED)
    ' MULTIPLAYER
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_BACKGROUND, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_addMultiMenu), "addMultiMenu", "Multiplayer Options / Alliance dialog"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KPENTER, KeyAction_PRESSED)
    ' GAME CONTROLS - Moving around, zooming in, rotating etc
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_ScrollCamera(0, 1), "CameraUp", "Move Camera Up"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_UPARROW, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_ScrollCamera(0, -1), "CameraDown", "Move Camera Down"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_DOWNARROW, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_ScrollCamera(1, 0), "CameraRight", "Move Camera Right"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_RIGHTARROW, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_ScrollCamera(-1, 0), "CameraLeft", "Move Camera Left"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_LEFTARROW, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_SeekNorth), "SeekNorth", "Snap View to North"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_BACKSPACE, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleCamera), "ToggleCamera", "Toggle Tracking Camera"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_SPACE, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_BACKGROUND, KeyMappingType_FIXED, WzFunc_new(FN_kf_addInGameOptions), "addInGameOptions", "Display In-Game Options"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_ESC, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_RADAR, KeyMappingType_ASSIGNABLE, kf_RadarZoom(-1), "RadarZoomOut", "Zoom Radar Out"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_MINUS, KeyAction_PRESSED)
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_SECONDARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_MOUSE_KEY_CODE, MOUSE_WDN, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_RADAR, KeyMappingType_ASSIGNABLE, kf_RadarZoom(1), "RadarZoomIn", "Zoom Radar In"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_EQUALS, KeyAction_PRESSED)
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_SECONDARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_MOUSE_KEY_CODE, MOUSE_WUP, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_Zoom(-1), "ZoomIn", "Zoom In"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KP_PLUS, KeyAction_DOWN)
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_SECONDARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_MOUSE_KEY_CODE, MOUSE_WUP, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_Zoom(1), "ZoomOut", "Zoom Out"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KP_MINUS, KeyAction_DOWN)
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_SECONDARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_MOUSE_KEY_CODE, MOUSE_WDN, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_PitchForward), "PitchForward", "Pitch Forward"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KP_2, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_RotateLeft), "RotateLeft", "Rotate Left"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KP_4, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ResetPitch), "ResetPitch", "Reset Pitch"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KP_5, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_RotateRight), "RotateRight", "Rotate Right"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KP_6, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_PitchBack), "PitchBack", "Pitch Back"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KP_8, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_RightOrderMenu), "RightOrderMenu", "Orders Menu"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KP_0, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_SlowDown), "SlowDown", "Decrease Game Speed"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_MINUS, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_SpeedUp), "SpeedUp", "Increase Game Speed"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_EQUALS, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_NormalSpeed), "NormalSpeed", "Reset Game Speed"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_BACKSPACE, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_FaceNorth), "FaceNorth", "View North"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_UPARROW, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_FaceSouth), "FaceSouth", "View South"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_DOWNARROW, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_FaceEast), "FaceEast", "View East"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_LEFTARROW, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_FaceWest), "FaceWest", "View West"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_RIGHTARROW, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_JumpToResourceExtractor), "JumpToResourceExtractor", "View next Oil Derrick"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_KP_STAR, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_JumpToUnits(DROID_REPAIR), "JumpToRepairUnits", "View next Repair Unit"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_JumpToUnits(DROID_CONSTRUCT), "JumpToConstructorUnits", "View next Truck"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_JumpToUnits(DROID_SENSOR), "JumpToSensorUnits", "View next Sensor Unit"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_JumpToUnits(DROID_COMMAND), "JumpToCommandUnits", "View next Commander"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleOverlays), "ToggleOverlays", "Toggle Overlays"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_TAB, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleConsoleDrop), "ToggleConsoleDrop", "Toggle Console History "))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_BACKQUOTE, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleTeamChat), "ToggleTeamChat", "Toggle Team Chat History"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_BACKQUOTE, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_RotateBuildingCW), "RotateBuildingClockwise", "Rotate Building Clockwise"))
End Sub

' (VB) initializeKeyFunctions's entries, part 3
Private Sub initializeKeyFunctions_3()
    Dim e As Long
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_RotateBuildingACW), "RotateBuildingAnticlockwise", "Rotate Building Anticlockwise"))
    ' IN GAME MAPPINGS - Single key presses - ALL __DEBUG keymappings will be removed for master
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_CentreOnBase), "CentreOnBase", "Center View on HQ"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_B, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_ATTACK_LEVEL, DSS_ALEV_NEVER), "SetDroidAttackCease", "Hold Fire"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_C, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_JumpToUnassignedUnits), "JumpToUnassignedUnits", "View Unassigned Units"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_D, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_ATTACK_LEVEL, DSS_ALEV_ATTACKED), "SetDroidAttackReturn", "Return Fire"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_E, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_ATTACK_LEVEL, DSS_ALEV_ALWAYS), "SetDroidAttackAtWill", "Fire at Will"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_F, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_HALTTYPE, DSS_HALT_GUARD), "SetDroidMoveGuard", "Guard Position"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_G, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_RETURN_TO_LOC, DSS_RTL_BASE), "SetDroidReturnToBase", "Return to HQ"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_H, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_OrderDroid(DORDER_HOLD), "SetDroidOrderHold", "Hold Position"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_H, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_ATTACK_RANGE, DSS_ARANGE_OPTIMUM), "SetDroidRangeOptimum", "Optimum Range"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_I, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_ATTACK_RANGE, DSS_ARANGE_SHORT), "SetDroidRangeShort", "Short Range"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_O, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_HALTTYPE, DSS_HALT_PURSUE), "SetDroidMovePursue", "Pursue"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_P, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_PATROL, DSS_PATROL_SET), "SetDroidMovePatrol", "Patrol"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_Q, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_RETURN_TO_LOC, DSS_RTL_REPAIR), "SetDroidGoForRepair", "Return For Repair"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_R, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_OrderDroid(DORDER_STOP), "SetDroidOrderStop", "Stop Droid"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_S, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_RETURN_TO_LOC, DSS_RTL_TRANSPORT), "SetDroidGoToTransport", "Go to Transport"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_T, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_ATTACK_RANGE, DSS_ARANGE_LONG), "SetDroidRangeLong", "Long Range"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_U, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_SendGlobalMessage), "SendGlobalMessage", "Send Global Text Message"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_RETURN, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_SendTeamMessage), "SendTeamMessage", "Send Team Text Message"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_RETURN, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_SendGlobalQuickChat), "SendGlobalQuickChat", "Send Global Quick Chat"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_SendTeamQuickChat), "SendTeamQuickChat", "Send Team Quick Chat"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_RETURN, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_AddHelpBlip), "AddHelpBlip", "Drop a beacon"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_H, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleShadows), "ToggleShadows", "Toggles shadows"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_S, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_toggleTrapCursor), "toggleTrapCursor", "Trap cursor"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_T, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_RADAR, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleRadarTerrain), "ToggleRadarTerrain", "Toggle radar terrain"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_TAB, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_RADAR, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ToggleRadarAllyEnemy), "ToggleRadarAllyEnemy", "Toggle ally-enemy radar view"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_TAB, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, WzFunc_new(FN_kf_ShowMappings), "ShowMappings", "Show all keyboard mappings"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_M, KeyAction_PRESSED)
    ' Some extra non QWERTY mappings but functioning in same way
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_REPAIR_LEVEL, DSS_REPLEV_LOW), "SetDroidRetreatMedium", "Retreat at Medium Damage"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_COMMA, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_REPAIR_LEVEL, DSS_REPLEV_HIGH), "SetDroidRetreatHeavy", "Retreat at Heavy Damage"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_FULLSTOP, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_REPAIR_LEVEL, DSS_REPLEV_NEVER), "SetDroidRetreatNever", "Do or Die!"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_FORWARDSLASH, KeyAction_PRESSED)
    ' In game mappings - COMBO (CTRL + LETTER) presses
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_ALL_COMBAT), "SelectAllCombatUnits", "Select all Combat Units"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_A, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_CYBORG), "SelectAllCyborgs", "Select all Cyborgs"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_C, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_ALL_DAMAGED), "SelectAllDamaged", "Select all Heavily Damaged Units"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_D, KeyAction_PRESSED)
    ' xgettext:no-c-format
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_ALL_LAND_MILDLY_OR_NOT_DAMAGED), "SelectAllLandMildNoDamage", "Select all Land Combat Units with >50% HP and no group"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_HALF_TRACKED), "SelectAllHalfTracked", "Select all Half-tracks"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_F, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNoGroupUnits(DST_HALF_TRACKED), "SelectAllHalfTrackedNG", "Select all Half-tracks without group"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_F, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_HOVER), "SelectAllHovers", "Select all Hovers"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_H, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SetDroid(DSO_RECYCLE, DSS_RECYCLE_SET), "SetDroidRecycle", "Return for Recycling"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_R, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_UNUSED, DS_ALL_UNITS, True), "SelectAllOnScreenUnits", "Select all Units on Screen"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_S, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_TRACKED), "SelectAllTracked", "Select all Tracks"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_T, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNoGroupUnits(DST_TRACKED), "SelectAllTrackedNG", "Select all Tracks without group"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_D, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_UNUSED, DS_ALL_UNITS), "SelectAllUnits", "Select EVERY unit"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_U, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_VTOL), "SelectAllVTOLs", "Select all VTOLs"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_V, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_VTOL_ARMED), "SelectAllArmedVTOLs", "Select all fully-armed VTOLs"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_V, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNoGroupUnits(DST_VTOL_ARMED), "SelectAllArmedVTOLsNG", "Select all fully-armed VTOLs without group"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_V, KeyAction_PRESSED)
End Sub

' (VB) initializeKeyFunctions's entries, part 4
Private Sub initializeKeyFunctions_4()
    Dim e As Long
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_WHEELED), "SelectAllWheeled", "Select all Wheels"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_W, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_FrameRate), "FrameRate", "Show frame rate"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_Y, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_ALL_SAME), "SelectAllSameType", "Select all units with the same components"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_Z, KeyAction_PRESSED)
    ' In game mappings - COMBO (SHIFT + LETTER) presses
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_ALL_COMBAT_CYBORG), "SelectAllCombatCyborgs", "Select all Combat Cyborgs"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_C, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNoGroupUnits(DST_ALL_COMBAT_CYBORG), "SelectAllCombatCyborgsNG", "Select all Combat Cyborgs without group"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_C, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_ENGINEER), "SelectAllEngineers", "Select all Engineers"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_E, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_ALL_COMBAT_LAND), "SelectAllLandCombatUnits", "Select all Land Combat Units"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_G, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNoGroupUnits(DST_ALL_COMBAT_LAND), "SelectAllLandCombatUnitsNG", "Select all Land Combat Units without group"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_G, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_MECHANIC), "SelectAllMechanics", "Select all Mechanics"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_M, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_TRANSPORTER), "SelectAllTransporters", "Select all Transporters"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_P, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_REPAIR_TANK), "SelectAllRepairTanks", "Select all Repair Tanks"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_R, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_SENSOR), "SelectAllSensorUnits", "Select all Sensor Units"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_S, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNoGroupUnits(DST_SENSOR), "SelectAllSensorUnitsNG", "Select all Sensor Units without group"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_J, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectUnits(DST_TRUCK), "SelectAllTrucks", "Select all Trucks"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_T, KeyAction_PRESSED)
    ' SELECT PLAYERS - DEBUG ONLY
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextFactory(REF_FACTORY), "SelectNextFactory", "Select next Factory"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextResearch(), "SelectNextResearch", "Select next Research Facility"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextPowerStation(), "SelectNextPowerStation", "Select next Power Generator"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextFactory(REF_CYBORG_FACTORY), "SelectNextCyborgFactory", "Select next Cyborg Factory"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextFactory(REF_VTOL_FACTORY), "SelectNextVtolFactory", "Select next VTOL Factory"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextFactory(REF_FACTORY, True), "JumpNextFactory", "Jump to next Factory"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextResearch(True), "JumpNextResearch", "Jump to next Research Facility"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextPowerStation(True), "JumpNextPowerStation", "Jump to next Power Generator"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextFactory(REF_CYBORG_FACTORY, True), "JumpNextCyborgFactory", "Jump to next Cyborg Factory"))
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_GAMEPLAY, KeyMappingType_ASSIGNABLE, kf_SelectNextFactory(REF_VTOL_FACTORY, True), "JumpNextVtolFactory", "Jump to next VTOL Factory"))
    ' Debug options
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_BACKGROUND, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ToggleDebugMappings), "ToggleDebugMappings", "Toggle Debug Mappings"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_BACKSPACE, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_PrioritizeDebugMappings), "PrioritizeDebugMappings", "Prioritize Debug Mappings", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_BACKSPACE, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ToggleLevelEditor), "ToggleLevelEditor", "Toggle Level Editor", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_L, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ToggleShowPath), "ToggleShowPath", "Toggle display of droid path", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_M, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ToggleShowGateways), "ToggleShowGateways", "Toggle display of gateways", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_E, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ToggleVisibility), "ToggleVisibility", "Toggle visibility", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_V, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_LEVEL_EDITOR, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_RaiseTile), "RaiseTile", "Raise tile height", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_W, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_LEVEL_EDITOR, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_LowerTile), "LowerTile", "Lower tile height", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_A, KeyAction_DOWN)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ToggleFog), "ToggleFog", "Toggles All fog", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_J, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ToggleWeather), "ToggleWeather", "Trigger some weather", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_Q, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_LEVEL_EDITOR, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_TriFlip), "TriFlip", "Flip terrain triangle", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_IGNORE, KeyMappingInputSource_KEY_CODE, KEY_K, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_PerformanceSample), "PerformanceSample", "Make a performance measurement sample", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_K, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_AllAvailable), "AllAvailable", "Make all items available", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_A, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_HAS_SELECTION, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_KillSelected), "KillSelected", "Kill Selected Unit(s)", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_K, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ToggleGodMode), "ToggleGodMode", "Toggle god Mode Status", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_G, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ChooseOptions), "ChooseOptions", "Display Options Screen", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_O, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_FinishResearch), "FinishResearch", "Complete current research", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_X, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_MISC, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_RevealMapAtPos), "RevealMapAtPos", "Reveal map at mouse position", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LSHIFT, KeyMappingInputSource_KEY_CODE, KEY_W, KeyAction_PRESSED)
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_DEBUG_HAS_SELECTION, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_TraceObject), "TraceObject", "Trace a game object", True))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LCTRL, KeyMappingInputSource_KEY_CODE, KEY_L, KeyAction_PRESSED)

    ' Hidden/"Hardcoded" mappings
    e = entries_emplace_back(KeyFunctionInfo_new(InputContext_ALWAYS_ACTIVE, KeyMappingType_HIDDEN, WzFunc_new(FN_kf_ToggleFullscreen), "ToggleFullscreen", "Toggle fullscreen"))
    KeyFunctionInfo_defaultMappings_push_back e, KeyMappingSlot_PRIMARY, KeyCombination_(KEY_LALT, KeyMappingInputSource_KEY_CODE, KEY_RETURN, KeyAction_PRESSED)
End Sub
' KeyFunctionConfiguration::KeyFunctionConfiguration() (gKeyFuncConfig's: Sub Main calls it before
' anything reads gKeyFuncConfig, as C++ constructs the global)
Public Sub KeyFunctionConfiguration_construct()
    Erase entries
    entries_size = 0
    initializeKeyFunctions
    ' for (size_t i = 0; i < entries.size(); ++i) nameToIndexMap[entries[i].name] = i; (keyFunctionInfoByName
    ' searches entries)
End Sub

' nonstd::optional<std::reference_wrapper<const KeyFunctionInfo>> KeyFunctionConfiguration::keyFunctionInfoByName(std::string const& name) const:
' the handle, 0 = nullopt (nameToIndexMap keeps the last entry of a name: the search runs from the end)
Public Function gKeyFuncConfig_keyFunctionInfoByName(ByVal name_ As String) As Long
    Dim i As Long
    For i = entries_size - 1 To 0 Step -1
        If StrComp(aKeyFunctionInfo(entries(i)).name_, name_, vbBinaryCompare) = 0 Then
            gKeyFuncConfig_keyFunctionInfoByName = entries(i)
            Exit Function
        End If
    Next
    gKeyFuncConfig_keyFunctionInfoByName = 0
End Function

' const KeyFunctionEntries KeyFunctionConfiguration::allKeyFunctionEntries() const: the count, the handles in
' allEntries(0 .. count - 1)
Public Function gKeyFuncConfig_allKeyFunctionEntries(allEntries() As Long) As Long
    Dim i As Long
    If entries_size = 0 Then
        Erase allEntries
        gKeyFuncConfig_allKeyFunctionEntries = 0
        Exit Function
    End If
    ReDim allEntries(0 To entries_size - 1)
    For i = 0 To entries_size - 1
        allEntries(i) = entries(i)
    Next
    gKeyFuncConfig_allKeyFunctionEntries = entries_size
End Function

' MARK: - KeyFunctionInfo

' KeyFunctionInfo::KeyFunctionInfo(const ContextId& context, const KeyMappingType type, const MappableFunction function,
' const std::string name, const std::string displayName, const std::vector<...> defaultMappings, const bool bIsDebugOnly)
' and the four-argument one (displayName "", no default mappings): the new info's handle. The defaultMappings
' are KeyFunctionInfo_defaultMappings_push_back's.
Public Function KeyFunctionInfo_new(ByVal context As String, ByVal typ As Long, ByVal function_ As WzFunc, ByVal name_ As String, Optional ByVal displayName As String = "", Optional ByVal bIsDebugOnly As Boolean = False) As Long
    Dim h As Long, blank As KeyFunctionInfo
    If kfiFree_size > 0 Then
        kfiFree_size = kfiFree_size - 1
        h = kfiFree(kfiFree_size)
    Else
        aKeyFunctionInfo_size = aKeyFunctionInfo_size + 1
        h = aKeyFunctionInfo_size
        ReDim Preserve aKeyFunctionInfo(1 To aKeyFunctionInfo_size)
    End If
    aKeyFunctionInfo(h) = blank
    aKeyFunctionInfo(h).context = context
    aKeyFunctionInfo(h).typ = typ
    Set aKeyFunctionInfo(h).function_ = function_
    aKeyFunctionInfo(h).name_ = name_
    aKeyFunctionInfo(h).displayName = displayName
    aKeyFunctionInfo(h).defaultMappings_size = 0
    aKeyFunctionInfo(h).bIsDebugOnly = bIsDebugOnly
    aKeyFunctionInfo(h).inUse = True
    KeyFunctionInfo_new = h
End Function

' (VB) KeyFunctionInfo's constructor argument defaultMappings: one more { slot, keys }
Public Sub KeyFunctionInfo_defaultMappings_push_back(ByVal info As Long, ByVal slot As Long, keys As KeyCombination)
    Dim n As Long
    n = aKeyFunctionInfo(info).defaultMappings_size
    If n >= KEYFUNCTION_MAX_DEFAULTS Then debug_ 0, LOG_ERROR, "KeyFunctionInfo", "(VB) %s: more than KEYFUNCTION_MAX_DEFAULTS default mappings", aKeyFunctionInfo(info).name_: Exit Sub
    aKeyFunctionInfo(info).defaultMappings(n).first = slot
    aKeyFunctionInfo(info).defaultMappings(n).second = keys
    aKeyFunctionInfo(info).defaultMappings_size = n + 1
End Sub

' (VB) the end of a map marker's KeyFunctionInfo (markerKeyFunctions' erase and clear)
Public Sub KeyFunctionInfo_free(ByVal info As Long)
    Dim blank As KeyFunctionInfo
    If info = 0 Then Exit Sub
    If Not aKeyFunctionInfo(info).inUse Then Exit Sub
    aKeyFunctionInfo(info) = blank
    ReDim Preserve kfiFree(0 To kfiFree_size)
    kfiFree(kfiFree_size) = info
    kfiFree_size = kfiFree_size + 1
End Sub

Public Function KeyFunctionInfo_context(ByVal info As Long) As String
    KeyFunctionInfo_context = aKeyFunctionInfo(info).context
End Function

Public Function KeyFunctionInfo_type(ByVal info As Long) As Long
    KeyFunctionInfo_type = aKeyFunctionInfo(info).typ
End Function

Public Function KeyFunctionInfo_function(ByVal info As Long) As WzFunc
    Set KeyFunctionInfo_function = aKeyFunctionInfo(info).function_
End Function

Public Function KeyFunctionInfo_name(ByVal info As Long) As String
    KeyFunctionInfo_name = aKeyFunctionInfo(info).name_
End Function

Public Function KeyFunctionInfo_displayName(ByVal info As Long) As String
    KeyFunctionInfo_displayName = aKeyFunctionInfo(info).displayName
End Function

Public Function KeyFunctionInfo_bIsDebugOnly(ByVal info As Long) As Boolean
    KeyFunctionInfo_bIsDebugOnly = aKeyFunctionInfo(info).bIsDebugOnly
End Function

' info.defaultMappings.size()
Public Function KeyFunctionInfo_defaultMappingsCount(ByVal info As Long) As Long
    KeyFunctionInfo_defaultMappingsCount = aKeyFunctionInfo(info).defaultMappings_size
End Function

' info.defaultMappings[i].first
Public Function KeyFunctionInfo_defaultMappingSlot(ByVal info As Long, ByVal i As Long) As Long
    KeyFunctionInfo_defaultMappingSlot = aKeyFunctionInfo(info).defaultMappings(i).first
End Function

' info.defaultMappings[i].second (the KeyCombination)
Public Function KeyFunctionInfo_defaultMappingKeys(ByVal info As Long, ByVal i As Long) As KeyCombination
    KeyFunctionInfo_defaultMappingKeys = aKeyFunctionInfo(info).defaultMappings(i).second
End Function

' info.defaultMappings[i].second.meta
Public Function KeyFunctionInfo_defaultMappingMeta(ByVal info As Long, ByVal i As Long) As Long
    KeyFunctionInfo_defaultMappingMeta = aKeyFunctionInfo(info).defaultMappings(i).second.meta
End Function

' info.defaultMappings[i].second.input.source
Public Function KeyFunctionInfo_defaultMappingInputSource(ByVal info As Long, ByVal i As Long) As Long
    KeyFunctionInfo_defaultMappingInputSource = aKeyFunctionInfo(info).defaultMappings(i).second.input_source
End Function

' info.defaultMappings[i].second.input.value
Public Function KeyFunctionInfo_defaultMappingInputValue(ByVal info As Long, ByVal i As Long) As Long
    KeyFunctionInfo_defaultMappingInputValue = aKeyFunctionInfo(info).defaultMappings(i).second.input_value
End Function

' info.defaultMappings[i].second.action
Public Function KeyFunctionInfo_defaultMappingAction(ByVal info As Long, ByVal i As Long) As Long
    KeyFunctionInfo_defaultMappingAction = aKeyFunctionInfo(info).defaultMappings(i).second.action
End Function

' info.defaultMappings[i].second.input.isCleared()
Public Function KeyFunctionInfo_defaultMappingInputIsCleared(ByVal info As Long, ByVal i As Long) As Boolean
    With aKeyFunctionInfo(info).defaultMappings(i).second
        KeyFunctionInfo_defaultMappingInputIsCleared = KeyMappingInput_isCleared(.input_source, .input_value)
    End With
End Function

' MARK: - KeyMappingInput

' bool KeyMappingInput::isPressed() const
Public Function KeyMappingInput_isPressed(ByVal source As Long, ByVal value As Long) As Boolean
    Select Case source
    Case KeyMappingInputSource_KEY_CODE
        KeyMappingInput_isPressed = keyPressed(value)
    Case KeyMappingInputSource_MOUSE_KEY_CODE
        KeyMappingInput_isPressed = mousePressed(value)
    Case Else
        KeyMappingInput_isPressed = False
    End Select
End Function

' bool KeyMappingInput::isDown() const
Public Function KeyMappingInput_isDown(ByVal source As Long, ByVal value As Long) As Boolean
    Select Case source
    Case KeyMappingInputSource_KEY_CODE
        KeyMappingInput_isDown = keyDown(value)
    Case KeyMappingInputSource_MOUSE_KEY_CODE
        KeyMappingInput_isDown = mouseDown(value)
    Case Else
        KeyMappingInput_isDown = False
    End Select
End Function

' bool KeyMappingInput::isReleased() const
Public Function KeyMappingInput_isReleased(ByVal source As Long, ByVal value As Long) As Boolean
    Select Case source
    Case KeyMappingInputSource_KEY_CODE
        KeyMappingInput_isReleased = keyReleased(value)
    Case KeyMappingInputSource_MOUSE_KEY_CODE
        KeyMappingInput_isReleased = mouseReleased(value)
    Case Else
        KeyMappingInput_isReleased = False
    End Select
End Function

' bool KeyMappingInput::isCleared() const
Public Function KeyMappingInput_isCleared(ByVal source As Long, ByVal value As Long) As Boolean
    KeyMappingInput_isCleared = source = KeyMappingInputSource_KEY_CODE And value = KEY_MAXSCAN
End Function

' bool KeyMappingInput::is(const KEY_CODE keyCode) const
Public Function KeyMappingInput_is(ByVal source As Long, ByVal value As Long, ByVal keyCode As Long) As Boolean
    KeyMappingInput_is = source = KeyMappingInputSource_KEY_CODE And value = keyCode
End Function

' bool KeyMappingInput::is(const MOUSE_KEY_CODE mouseKeyCode) const
Public Function KeyMappingInput_isMouse(ByVal source As Long, ByVal value As Long, ByVal mouseKeyCode As Long) As Boolean
    KeyMappingInput_isMouse = source = KeyMappingInputSource_MOUSE_KEY_CODE And value = mouseKeyCode
End Function

' nonstd::optional<KEY_CODE> KeyMappingInput::asKeyCode() const: has_value, the value in keyCode
Public Function KeyMappingInput_asKeyCode(ByVal source As Long, ByVal value As Long, keyCode As Long) As Boolean
    If source = KeyMappingInputSource_KEY_CODE Then
        keyCode = value
        KeyMappingInput_asKeyCode = True
    Else
        KeyMappingInput_asKeyCode = False
    End If
End Function

' nonstd::optional<MOUSE_KEY_CODE> KeyMappingInput::asMouseKeyCode() const: has_value, the value in mouseKeyCode
Public Function KeyMappingInput_asMouseKeyCode(ByVal source As Long, ByVal value As Long, mouseKeyCode As Long) As Boolean
    If source = KeyMappingInputSource_MOUSE_KEY_CODE Then
        mouseKeyCode = value
        KeyMappingInput_asMouseKeyCode = True
    Else
        KeyMappingInput_asMouseKeyCode = False
    End If
End Function

' bool operator==(const KeyMappingInput& lhs, const KeyMappingInput& rhs)
Public Function KeyMappingInput_equals(ByVal lhsSource As Long, ByVal lhsValue As Long, ByVal rhsSource As Long, ByVal rhsValue As Long) As Boolean
    If lhsSource <> rhsSource Then
        KeyMappingInput_equals = False
        Exit Function
    End If

    Select Case lhsSource
    Case KeyMappingInputSource_KEY_CODE
        KeyMappingInput_equals = lhsValue = rhsValue
    Case KeyMappingInputSource_MOUSE_KEY_CODE
        KeyMappingInput_equals = lhsValue = rhsValue
    Case Else
        KeyMappingInput_equals = False
    End Select
End Function

' KeyMappingInputSource keyMappingSourceByName(std::string const& name)
Public Function keyMappingSourceByName(ByVal name_ As String) As Long
    If name_ = "default" Then
        keyMappingSourceByName = KeyMappingInputSource_KEY_CODE
    ElseIf name_ = "mouse_key" Then
        keyMappingSourceByName = KeyMappingInputSource_MOUSE_KEY_CODE
    Else
        debug_ 433, LOG_WZ, "keyMappingSourceByName", "Encountered invalid key mapping source name '%s', falling back to using 'default'", name_
        keyMappingSourceByName = KeyMappingInputSource_KEY_CODE
    End If
End Function

' KeyMappingSlot keyMappingSlotByName(std::string const& name)
Public Function keyMappingSlotByName(ByVal name_ As String) As Long
    If name_ = "primary" Then
        keyMappingSlotByName = KeyMappingSlot_PRIMARY
    ElseIf name_ = "secondary" Then
        keyMappingSlotByName = KeyMappingSlot_SECONDARY
    Else
        debug_ 450, LOG_WZ, "keyMappingSlotByName", "Encountered invalid key mapping slot name '%s', falling back to using 'primary'", name_
        keyMappingSlotByName = KeyMappingSlot_PRIMARY
    End If
End Function

' MARK: - KeyCombination

' KeyCombination::KeyCombination(const KEY_CODE meta, const KeyMappingInput input, const KeyAction action) (the
' other constructors: meta KEY_IGNORE, action KeyAction::PRESSED)
Public Function KeyCombination_(ByVal meta As Long, ByVal inputSource As Long, ByVal inputValue As Long, ByVal action As Long) As KeyCombination
    KeyCombination_.meta = meta
    KeyCombination_.input_source = inputSource
    KeyCombination_.input_value = inputValue
    If KeyMappingInput_isMouse(inputSource, inputValue, MOUSE_WUP) Or KeyMappingInput_isMouse(inputSource, inputValue, MOUSE_WDN) Then
        KeyCombination_.action = KeyAction_PRESSED
    Else
        KeyCombination_.action = action
    End If
End Function

' bool operator==(const KeyCombination& lhs, const KeyCombination& rhs)
Public Function KeyCombination_equals(lhs As KeyCombination, rhs As KeyCombination) As Boolean
    KeyCombination_equals = KeyMappingInput_equals(lhs.input_source, lhs.input_value, rhs.input_source, rhs.input_value) And lhs.meta = rhs.meta And lhs.action = rhs.action
End Function
