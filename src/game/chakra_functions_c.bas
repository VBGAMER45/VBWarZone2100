Attribute VB_Name = "chakra_functions_c"
' src/quickjs_backend.cpp, lines 3140-3694: the script API's wrappers
' (IMPL_JS_FUNC, js_enumTemplates and the other custom js_ functions),
' js_stats_get / js_stats_set, setStatsFunc, constructUpgradesQuickJSValue
' and quickjs_scripting_instance::registerFunctions (round 2 of cluster C,
' step 1).
'
' How a wrapper is written (IMPL_JS_FUNC(name, wzapi::name) and wrap__):
' * Give the function a number: Private Const JSFN_<name> As Long =
'   JSFN_API_FIRST + k, k the registration's place in registerFunctions
'   (chakra_backend_c's jsNativeFunction passes every number from
'   JSFN_API_FIRST on to jsCallApiFunction). A JS name registered for
'   another name's implementation (JS_REGISTER_FUNC_NAME: "_", "label",
'   "resetArea", "addReticuleButton", "setScrollParams") has its own
'   number, so that its length is its own.
' * registerFunctions registers it: JS_REGISTER_FUNC(name, n) is
'   JS_REGISTER_FUNC_(global_obj, "name", JSFN_<name>, n, <its C++ line>)
'   (jsNewCFunction: QuickJS pads argv to n with undefined; then
'   JS_SetPropertyStr, with the macro's ASSERT_OR_RETURNs at the line of
'   the macro's use, as GCC's __LINE__). JS_REGISTER_FUNC2(name, min,
'   max): length min.
' * jsCallApiFunction's Select Case calls js_<name>(ctx, argc, argv), which
'   does what wrap__ does: idx = 0; context = wzapi_execution_context_new(ctx);
'   each parameter unboxed in order with chakra_backend_c's unbox_<type>
'   (optional<T>: has_x = unbox_has(idx, argc): If has_x Then opt_x =
'   unbox_<T>(...)); the wzapi function called; its result boxed (box_<type>;
'   no_return_value: box_no_return_value; optional<T>: box_<T> or
'   jsUndefinedValue; a vector: box_vector_<T>(result, count)).
'   Example:
'       ' IMPL_JS_FUNC(setPower, wzapi::setPower)
'       Private Function js_setPower(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
'           Dim idx As Long, context As wzapi_execution_context, p_power As Long, has_player As Boolean, opt_player As Long
'           Set context = wzapi_execution_context_new(ctx)
'           p_power = unbox_int(idx, argc, argv, "wzapi::setPower")
'           has_player = unbox_has(idx, argc)
'           If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::setPower")
'           wzapi_setPower context, p_power, has_player, opt_player
'           js_setPower = box_no_return_value()
'       End Function
'   (the function name passed is wrap_'s #wzapi_function, the text of the macro's argument;
'   WZAPI_NO_PARAMS_NO_CONTEXT (enumSelected): no context)
' * A wrapper returns the boxed value (never 0, unless it throws: then
'   jsThrowReferenceError and 0, as a QuickJS function returning JS_EXCEPTION).
'   wzapi's SCRIPT_ASSERTs (context.throwError) and the UNBOX_SCRIPT_ASSERTs do
'   not throw (see chakra_backend_c's header). Every VB error is trapped by
'   jsNativeFunction.
' * std::unique_ptr results (makeTemplate) are boxed, then freed, as the
'   unique_ptr's destructor does.
'
' The Upgrades object: each property of Upgrades[player][class][entity] is a
' getter and a setter (setStatsFunc), as QuickJS defines them. QuickJS's
' getter and setter read their player, type, index and name from their own
' function object (js_debugger_get_current_funcObject); JsRT gives a native
' function no access to itself, so they are kept in statsFuncs instead, and
' the getter's and setter's function numbers say which entry is theirs:
' JSFN_STATS_FIRST + 2 * entry (the getter) and + 1 (the setter), above the
' numbers that have a length (jsNewCFunction records lengths below 65536).
' The entries are shared by every instance (one per player, type, index and
' name); the function objects' own "player", "type", "index" and "name"
' properties, which nothing but js_stats_get / js_stats_set reads, are not
' defined. The properties are defined in GameEntityRules' order (its
' NameToTypeMap, a std::map: by name, ascending): wzapi_c's
' GameEntityRules_properties.
Option Explicit

Private Declare Function JsCreateTypeError Lib "chakra.dll" (ByVal message As Long, errorValue As Long) As Long

' (VB) the getters and setters of the Upgrades object: JSFN_STATS_FIRST + 2 * entry (+ 1 for the setter)
Private Const JSFN_STATS_FIRST As Long = &H100000

' (VB) the API functions' numbers: one per JS_REGISTER_FUNC* of registerFunctions, in its order
Private Const JSFN_UNDERSCORE As Long = JSFN_API_FIRST + 0
Private Const JSFN_dump As Long = JSFN_API_FIRST + 1
Private Const JSFN_syncRandom As Long = JSFN_API_FIRST + 2
Private Const JSFN_label As Long = JSFN_API_FIRST + 3
Private Const JSFN_getObject As Long = JSFN_API_FIRST + 4
Private Const JSFN_addLabel As Long = JSFN_API_FIRST + 5
Private Const JSFN_removeLabel As Long = JSFN_API_FIRST + 6
Private Const JSFN_getLabel As Long = JSFN_API_FIRST + 7
Private Const JSFN_enumLabels As Long = JSFN_API_FIRST + 8
Private Const JSFN_enumGateways As Long = JSFN_API_FIRST + 9
Private Const JSFN_enumTemplates As Long = JSFN_API_FIRST + 10
Private Const JSFN_makeTemplate As Long = JSFN_API_FIRST + 11
Private Const JSFN_setAlliance As Long = JSFN_API_FIRST + 12
Private Const JSFN_sendAllianceRequest As Long = JSFN_API_FIRST + 13
Private Const JSFN_setAssemblyPoint As Long = JSFN_API_FIRST + 14
Private Const JSFN_setSunPosition As Long = JSFN_API_FIRST + 15
Private Const JSFN_setSunIntensity As Long = JSFN_API_FIRST + 16
Private Const JSFN_setFogColour As Long = JSFN_API_FIRST + 17
Private Const JSFN_setWeather As Long = JSFN_API_FIRST + 18
Private Const JSFN_setSky As Long = JSFN_API_FIRST + 19
Private Const JSFN_cameraSlide As Long = JSFN_API_FIRST + 20
Private Const JSFN_cameraTrack As Long = JSFN_API_FIRST + 21
Private Const JSFN_cameraZoom As Long = JSFN_API_FIRST + 22
Private Const JSFN_resetArea As Long = JSFN_API_FIRST + 23
Private Const JSFN_resetLabel As Long = JSFN_API_FIRST + 24
Private Const JSFN_addSpotter As Long = JSFN_API_FIRST + 25
Private Const JSFN_removeSpotter As Long = JSFN_API_FIRST + 26
Private Const JSFN_syncRequest As Long = JSFN_API_FIRST + 27
Private Const JSFN_replaceTexture As Long = JSFN_API_FIRST + 28
Private Const JSFN_changePlayerColour As Long = JSFN_API_FIRST + 29
Private Const JSFN_setHealth As Long = JSFN_API_FIRST + 30
Private Const JSFN_useSafetyTransport As Long = JSFN_API_FIRST + 31
Private Const JSFN_restoreLimboMissionData As Long = JSFN_API_FIRST + 32
Private Const JSFN_getMultiTechLevel As Long = JSFN_API_FIRST + 33
Private Const JSFN_setCampaignNumber As Long = JSFN_API_FIRST + 34
Private Const JSFN_getMissionType As Long = JSFN_API_FIRST + 35
Private Const JSFN_getRevealStatus As Long = JSFN_API_FIRST + 36
Private Const JSFN_setRevealStatus As Long = JSFN_API_FIRST + 37
Private Const JSFN_autoSave As Long = JSFN_API_FIRST + 38
Private Const JSFN_hackNetOff As Long = JSFN_API_FIRST + 39
Private Const JSFN_hackNetOn As Long = JSFN_API_FIRST + 40
Private Const JSFN_hackAddMessage As Long = JSFN_API_FIRST + 41
Private Const JSFN_hackRemoveMessage As Long = JSFN_API_FIRST + 42
Private Const JSFN_hackGetObj As Long = JSFN_API_FIRST + 43
Private Const JSFN_hackAssert As Long = JSFN_API_FIRST + 44
Private Const JSFN_hackMarkTiles As Long = JSFN_API_FIRST + 45
Private Const JSFN_receiveAllEvents As Long = JSFN_API_FIRST + 46
Private Const JSFN_hackDoNotSave As Long = JSFN_API_FIRST + 47
Private Const JSFN_hackPlayIngameAudio As Long = JSFN_API_FIRST + 48
Private Const JSFN_hackStopIngameAudio As Long = JSFN_API_FIRST + 49
Private Const JSFN_debug As Long = JSFN_API_FIRST + 50
Private Const JSFN_console As Long = JSFN_API_FIRST + 51
Private Const JSFN_clearConsole As Long = JSFN_API_FIRST + 52
Private Const JSFN_structureIdle As Long = JSFN_API_FIRST + 53
Private Const JSFN_enumStruct As Long = JSFN_API_FIRST + 54
Private Const JSFN_enumStructOffWorld As Long = JSFN_API_FIRST + 55
Private Const JSFN_enumDroid As Long = JSFN_API_FIRST + 56
Private Const JSFN_enumGroup As Long = JSFN_API_FIRST + 57
Private Const JSFN_enumFeature As Long = JSFN_API_FIRST + 58
Private Const JSFN_enumBlips As Long = JSFN_API_FIRST + 59
Private Const JSFN_enumSelected As Long = JSFN_API_FIRST + 60
Private Const JSFN_enumResearch As Long = JSFN_API_FIRST + 61
Private Const JSFN_enumRange As Long = JSFN_API_FIRST + 62
Private Const JSFN_enumArea As Long = JSFN_API_FIRST + 63
Private Const JSFN_getResearch As Long = JSFN_API_FIRST + 64
Private Const JSFN_pursueResearch As Long = JSFN_API_FIRST + 65
Private Const JSFN_findResearch As Long = JSFN_API_FIRST + 66
Private Const JSFN_distBetweenTwoPoints As Long = JSFN_API_FIRST + 67
Private Const JSFN_newGroup As Long = JSFN_API_FIRST + 68
Private Const JSFN_groupAddArea As Long = JSFN_API_FIRST + 69
Private Const JSFN_groupAddDroid As Long = JSFN_API_FIRST + 70
Private Const JSFN_groupAdd As Long = JSFN_API_FIRST + 71
Private Const JSFN_groupSize As Long = JSFN_API_FIRST + 72
Private Const JSFN_orderDroidLoc As Long = JSFN_API_FIRST + 73
Private Const JSFN_playerPower As Long = JSFN_API_FIRST + 74
Private Const JSFN_queuedPower As Long = JSFN_API_FIRST + 75
Private Const JSFN_isStructureAvailable As Long = JSFN_API_FIRST + 76
Private Const JSFN_pickStructLocation As Long = JSFN_API_FIRST + 77
Private Const JSFN_structureCanFit As Long = JSFN_API_FIRST + 78
Private Const JSFN_droidCanReach As Long = JSFN_API_FIRST + 79
Private Const JSFN_propulsionCanReach As Long = JSFN_API_FIRST + 80
Private Const JSFN_terrainType As Long = JSFN_API_FIRST + 81
Private Const JSFN_tileIsBurning As Long = JSFN_API_FIRST + 82
Private Const JSFN_orderDroidBuild As Long = JSFN_API_FIRST + 83
Private Const JSFN_orderDroidObj As Long = JSFN_API_FIRST + 84
Private Const JSFN_orderDroid As Long = JSFN_API_FIRST + 85
Private Const JSFN_buildDroid As Long = JSFN_API_FIRST + 86
Private Const JSFN_addDroid As Long = JSFN_API_FIRST + 87
Private Const JSFN_addDroidToTransporter As Long = JSFN_API_FIRST + 88
Private Const JSFN_addFeature As Long = JSFN_API_FIRST + 89
Private Const JSFN_componentAvailable As Long = JSFN_API_FIRST + 90
Private Const JSFN_isVTOL As Long = JSFN_API_FIRST + 91
Private Const JSFN_safeDest As Long = JSFN_API_FIRST + 92
Private Const JSFN_activateStructure As Long = JSFN_API_FIRST + 93
Private Const JSFN_chat As Long = JSFN_API_FIRST + 94
Private Const JSFN_quickChat As Long = JSFN_API_FIRST + 95
Private Const JSFN_getDroidPath As Long = JSFN_API_FIRST + 96
Private Const JSFN_addBeacon As Long = JSFN_API_FIRST + 97
Private Const JSFN_removeBeacon As Long = JSFN_API_FIRST + 98
Private Const JSFN_getDroidProduction As Long = JSFN_API_FIRST + 99
Private Const JSFN_getDroidLimit As Long = JSFN_API_FIRST + 100
Private Const JSFN_getExperienceModifier As Long = JSFN_API_FIRST + 101
Private Const JSFN_setDroidLimit As Long = JSFN_API_FIRST + 102
Private Const JSFN_setCommanderLimit As Long = JSFN_API_FIRST + 103
Private Const JSFN_setConstructorLimit As Long = JSFN_API_FIRST + 104
Private Const JSFN_setExperienceModifier As Long = JSFN_API_FIRST + 105
Private Const JSFN_getWeaponInfo As Long = JSFN_API_FIRST + 106
Private Const JSFN_enumCargo As Long = JSFN_API_FIRST + 107
Private Const JSFN_centreView As Long = JSFN_API_FIRST + 108
Private Const JSFN_playSound As Long = JSFN_API_FIRST + 109
Private Const JSFN_gameOverMessage As Long = JSFN_API_FIRST + 110
Private Const JSFN_addGuideTopic As Long = JSFN_API_FIRST + 111
Private Const JSFN_setStructureLimits As Long = JSFN_API_FIRST + 112
Private Const JSFN_applyLimitSet As Long = JSFN_API_FIRST + 113
Private Const JSFN_emitSound As Long = JSFN_API_FIRST + 114
Private Const JSFN_setMissionTime As Long = JSFN_API_FIRST + 115
Private Const JSFN_getMissionTime As Long = JSFN_API_FIRST + 116
Private Const JSFN_setReinforcementTime As Long = JSFN_API_FIRST + 117
Private Const JSFN_completeResearch As Long = JSFN_API_FIRST + 118
Private Const JSFN_completeAllResearch As Long = JSFN_API_FIRST + 119
Private Const JSFN_enableResearch As Long = JSFN_API_FIRST + 120
Private Const JSFN_setPower As Long = JSFN_API_FIRST + 121
Private Const JSFN_setPowerModifier As Long = JSFN_API_FIRST + 122
Private Const JSFN_setPowerStorageMaximum As Long = JSFN_API_FIRST + 123
Private Const JSFN_extraPowerTime As Long = JSFN_API_FIRST + 124
Private Const JSFN_setTutorialMode As Long = JSFN_API_FIRST + 125
Private Const JSFN_setDesign As Long = JSFN_API_FIRST + 126
Private Const JSFN_enableTemplate As Long = JSFN_API_FIRST + 127
Private Const JSFN_removeTemplate As Long = JSFN_API_FIRST + 128
Private Const JSFN_setMiniMap As Long = JSFN_API_FIRST + 129
Private Const JSFN_setReticuleButton As Long = JSFN_API_FIRST + 130
Private Const JSFN_setReticuleFlash As Long = JSFN_API_FIRST + 131
Private Const JSFN_showReticuleWidget As Long = JSFN_API_FIRST + 132
Private Const JSFN_showInterface As Long = JSFN_API_FIRST + 133
Private Const JSFN_hideInterface As Long = JSFN_API_FIRST + 134
Private Const JSFN_addReticuleButton As Long = JSFN_API_FIRST + 135
Private Const JSFN_removeReticuleButton As Long = JSFN_API_FIRST + 136
Private Const JSFN_enableStructure As Long = JSFN_API_FIRST + 137
Private Const JSFN_makeComponentAvailable As Long = JSFN_API_FIRST + 138
Private Const JSFN_enableComponent As Long = JSFN_API_FIRST + 139
Private Const JSFN_allianceExistsBetween As Long = JSFN_API_FIRST + 140
Private Const JSFN_removeStruct As Long = JSFN_API_FIRST + 141
Private Const JSFN_removeObject As Long = JSFN_API_FIRST + 142
Private Const JSFN_setScrollParams As Long = JSFN_API_FIRST + 143
Private Const JSFN_setScrollLimits As Long = JSFN_API_FIRST + 144
Private Const JSFN_getScrollLimits As Long = JSFN_API_FIRST + 145
Private Const JSFN_addStructure As Long = JSFN_API_FIRST + 146
Private Const JSFN_getStructureLimit As Long = JSFN_API_FIRST + 147
Private Const JSFN_countStruct As Long = JSFN_API_FIRST + 148
Private Const JSFN_countDroid As Long = JSFN_API_FIRST + 149
Private Const JSFN_loadLevel As Long = JSFN_API_FIRST + 150
Private Const JSFN_setDroidExperience As Long = JSFN_API_FIRST + 151
Private Const JSFN_donateObject As Long = JSFN_API_FIRST + 152
Private Const JSFN_donatePower As Long = JSFN_API_FIRST + 153
Private Const JSFN_setNoGoArea As Long = JSFN_API_FIRST + 154
Private Const JSFN_startTransporterEntry As Long = JSFN_API_FIRST + 155
Private Const JSFN_setTransporterExit As Long = JSFN_API_FIRST + 156
Private Const JSFN_setObjectFlag As Long = JSFN_API_FIRST + 157
Private Const JSFN_fireWeaponAtLoc As Long = JSFN_API_FIRST + 158
Private Const JSFN_fireWeaponAtObj As Long = JSFN_API_FIRST + 159
Private Const JSFN_transformPlayerToSpectator As Long = JSFN_API_FIRST + 160
Private Const JSFN_isSpectator As Long = JSFN_API_FIRST + 161
Private Const JSFN_setGameStoryLogPlayerDataValue As Long = JSFN_API_FIRST + 162

'--------------------------------------------------------------------------
' (VB) the getters' and setters' player, type, index and name (js_debugger_get_current_funcObject's
' properties of the function object), by entry

Private statsFunc_player() As Long
Private statsFunc_type() As Long
Private statsFunc_index() As Long
Private statsFunc_name() As String
Private statsFunc_size As Long
Private statsFuncKeys As Object         ' Scripting.Dictionary: "player,type,index,name" -> entry

' (VB) a GameEntityRules' properties by name (its NameToTypeMap's order), per map: the maps are fixed
Private Type UPGRADE_PROPERTIES
    built As Boolean
    names() As String
    types() As Long
    size As Long
End Type

Private upgradeProperties(0 To 15) As UPGRADE_PROPERTIES

'--------------------------------------------------------------------------
' Script functions
'
' All script functions should be prefixed with "js_" then followed by same name as in script.

' IMPL_JS_FUNC(getWeaponInfo, wzapi::getWeaponInfo)
Private Function js_getWeaponInfo(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_weaponName As String
    Set context = wzapi_execution_context_new(ctx)
    p_weaponName = unbox_string(idx, argc, argv, "wzapi::getWeaponInfo")
    js_getWeaponInfo = box_json(wzapi_getWeaponInfo(context, p_weaponName))
End Function

' IMPL_JS_FUNC(resetLabel, scripting_engine::resetLabel)
Private Function js_resetLabel(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_labelName As String, has_playerFilter As Boolean, opt_playerFilter As Long
    Set context = wzapi_execution_context_new(ctx)
    p_labelName = unbox_string(idx, argc, argv, "scripting_engine::resetLabel")
    has_playerFilter = unbox_has(idx, argc)
    If has_playerFilter Then opt_playerFilter = unbox_int(idx, argc, argv, "scripting_engine::resetLabel")
    scripting_engine_resetLabel context, p_labelName, has_playerFilter, opt_playerFilter
    js_resetLabel = box_no_return_value()
End Function

' IMPL_JS_FUNC(enumLabels, scripting_engine::enumLabels)
Private Function js_enumLabels(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, has_filterLabelType As Boolean, opt_filterLabelType As Long, result As StringVector
    Set context = wzapi_execution_context_new(ctx)
    has_filterLabelType = unbox_has(idx, argc)
    If has_filterLabelType Then opt_filterLabelType = unbox_int(idx, argc, argv, "scripting_engine::enumLabels")
    result = scripting_engine_enumLabels(context, has_filterLabelType, opt_filterLabelType)
    js_enumLabels = box_vector_string(result)
End Function

' IMPL_JS_FUNC(addLabel, scripting_engine::addLabel)
Private Function js_addLabel(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_object As generic_script_object, p_label As String, has_triggered As Boolean
    Dim opt_triggered As Long
    Set context = wzapi_execution_context_new(ctx)
    p_object = unbox_generic_script_object(idx, argc, argv, "scripting_engine::addLabel")
    p_label = unbox_string(idx, argc, argv, "scripting_engine::addLabel")
    has_triggered = unbox_has(idx, argc)
    If has_triggered Then opt_triggered = unbox_int(idx, argc, argv, "scripting_engine::addLabel")
    scripting_engine_addLabel context, p_object, p_label, has_triggered, opt_triggered
    js_addLabel = box_no_return_value()
End Function

' IMPL_JS_FUNC(removeLabel, scripting_engine::removeLabel)
Private Function js_removeLabel(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_label As String
    Set context = wzapi_execution_context_new(ctx)
    p_label = unbox_string(idx, argc, argv, "scripting_engine::removeLabel")
    js_removeLabel = box_int(scripting_engine_removeLabel(context, p_label))
End Function

' IMPL_JS_FUNC(getLabel, scripting_engine::getLabelJS)
Private Function js_getLabel(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_obj_id As game_object_identifier, result As String
    Set context = wzapi_execution_context_new(ctx)
    p_obj_id = unbox_game_object_identifier(idx, argc, argv, "scripting_engine::getLabelJS")
    If scripting_engine_getLabelJS(context, p_obj_id, result) Then
        js_getLabel = box_string(result)
    Else
        js_getLabel = jsUndefinedValue()
    End If
End Function

' IMPL_JS_FUNC(getObject, scripting_engine::getObject)
Private Function js_getObject(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_request As object_request, result As generic_script_object
    Set context = wzapi_execution_context_new(ctx)
    p_request = unbox_object_request(idx, argc, argv, "scripting_engine::getObject")
    result = scripting_engine_getObject(context, p_request)
    js_getObject = box_generic_script_object(result)
End Function

' IMPL_JS_FUNC(enumBlips, wzapi::enumBlips)
Private Function js_enumBlips(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_player As Long, result() As scr_position, n As Long
    Set context = wzapi_execution_context_new(ctx)
    p_player = unbox_int(idx, argc, argv, "wzapi::enumBlips")
    n = wzapi_enumBlips(context, p_player, result)
    js_enumBlips = box_vector_scr_position(result, n)
End Function

' IMPL_JS_FUNC(enumSelected, wzapi::enumSelected)
Private Function js_enumSelected(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim result() As Long, n As Long
    n = wzapi_enumSelected(result)
    js_enumSelected = box_vector_BASE_OBJECT(result, n)
End Function

' IMPL_JS_FUNC(enumGateways, wzapi::enumGateways)
Private Function js_enumGateways(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim context As wzapi_execution_context, result() As GATEWAY, n As Long
    Set context = wzapi_execution_context_new(ctx)
    n = wzapi_enumGateways(context, result)
    js_enumGateways = box_list_GATEWAY(result, n)
End Function

'-- ## enumTemplates(player)
'--
'-- Return an array containing all the buildable templates for the given player. (3.2+ only)
'--
Private Function js_enumTemplates(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim player As Long, result As Long, count As Long, i As Long, psTemplate As Long
    If Not (argc = 1) Then debug_ 3160, LOG_ERROR, "js_enumTemplates", "Must have one parameter": js_enumTemplates = jsNullValue(): Exit Function
    If Not (jsIsNumber(argv(0))) Then debug_ 3161, LOG_ERROR, "js_enumTemplates", "Supplied parameter must be a player number": js_enumTemplates = jsNullValue(): Exit Function
    player = JSValueToInt32(argv(0))

    result = jsNewArray() 'engine->newArray(droidTemplates[player].size());
    count = 0
    ' enumerateTemplates(player, [ctx, &result, &count](DROID_TEMPLATE* psTemplate) { ...; return true; })
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:659", "enumerateTemplates", "Invalid player: %u", player: js_enumTemplates = result: Exit Function
    For i = 0 To templateCount(player) - 1
        psTemplate = droidTemplates_at(player, i)
        jsDefineProp result, CStr(count), convTemplate(psTemplate), 0 ' TODO: Check return value?
        count = count + 1
    Next
    js_enumTemplates = result
End Function

' IMPL_JS_FUNC(enumGroup, scripting_engine::enumGroup)
Private Function js_enumGroup(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_groupId As Long, result() As Long, n As Long
    Set context = wzapi_execution_context_new(ctx)
    p_groupId = unbox_int(idx, argc, argv, "scripting_engine::enumGroup")
    n = scripting_engine_enumGroup(context, p_groupId, result)
    js_enumGroup = box_vector_BASE_OBJECT(result, n)
End Function

' IMPL_JS_FUNC(newGroup, scripting_engine::newGroup)
Private Function js_newGroup(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    js_newGroup = box_int(scripting_engine_newGroup(context))
End Function

' IMPL_JS_FUNC(groupAddArea, scripting_engine::groupAddArea)
Private Function js_groupAddArea(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_groupId As Long, p_x1 As Long, p_y1 As Long
    Dim p_x2 As Long, p_y2 As Long
    Set context = wzapi_execution_context_new(ctx)
    p_groupId = unbox_int(idx, argc, argv, "scripting_engine::groupAddArea")
    p_x1 = unbox_int(idx, argc, argv, "scripting_engine::groupAddArea")
    p_y1 = unbox_int(idx, argc, argv, "scripting_engine::groupAddArea")
    p_x2 = unbox_int(idx, argc, argv, "scripting_engine::groupAddArea")
    p_y2 = unbox_int(idx, argc, argv, "scripting_engine::groupAddArea")
    scripting_engine_groupAddArea context, p_groupId, p_x1, p_y1, p_x2, p_y2
    js_groupAddArea = box_no_return_value()
End Function

' IMPL_JS_FUNC(groupAddDroid, scripting_engine::groupAddDroid)
Private Function js_groupAddDroid(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_groupId As Long, p_psDroid As Long
    Set context = wzapi_execution_context_new(ctx)
    p_groupId = unbox_int(idx, argc, argv, "scripting_engine::groupAddDroid")
    p_psDroid = unbox_DROID(idx, argc, argv, "scripting_engine::groupAddDroid")
    scripting_engine_groupAddDroid context, p_groupId, p_psDroid
    js_groupAddDroid = box_no_return_value()
End Function

' IMPL_JS_FUNC(groupAdd, scripting_engine::groupAdd)
Private Function js_groupAdd(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_groupId As Long, p_psObj As Long
    Set context = wzapi_execution_context_new(ctx)
    p_groupId = unbox_int(idx, argc, argv, "scripting_engine::groupAdd")
    p_psObj = unbox_BASE_OBJECT(idx, argc, argv, "scripting_engine::groupAdd")
    scripting_engine_groupAdd context, p_groupId, p_psObj
    js_groupAdd = box_no_return_value()
End Function

' IMPL_JS_FUNC(groupSize, scripting_engine::groupSize)
Private Function js_groupSize(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_groupId As Long
    Set context = wzapi_execution_context_new(ctx)
    p_groupId = unbox_int(idx, argc, argv, "scripting_engine::groupSize")
    js_groupSize = box_int(scripting_engine_groupSize(context, p_groupId))
End Function

' IMPL_JS_FUNC(activateStructure, wzapi::activateStructure)
Private Function js_activateStructure(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psStruct As Long, has_psTarget As Boolean, opt_psTarget As Long
    Set context = wzapi_execution_context_new(ctx)
    p_psStruct = unbox_STRUCTURE(idx, argc, argv, "wzapi::activateStructure")
    has_psTarget = unbox_has(idx, argc)
    If has_psTarget Then opt_psTarget = unbox_BASE_OBJECT(idx, argc, argv, "wzapi::activateStructure")
    js_activateStructure = box_bool(wzapi_activateStructure(context, p_psStruct, has_psTarget, opt_psTarget))
End Function

' IMPL_JS_FUNC(findResearch, wzapi::findResearch)
Private Function js_findResearch(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_researchName As String, has_player As Boolean, opt_player As Long
    Dim result As wzapi_researchResults
    Set context = wzapi_execution_context_new(ctx)
    p_researchName = unbox_string(idx, argc, argv, "wzapi::findResearch")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::findResearch")
    result = wzapi_findResearch(context, p_researchName, has_player, opt_player)
    js_findResearch = box_researchResults(result)
End Function

' IMPL_JS_FUNC(pursueResearch, wzapi::pursueResearch)
Private Function js_pursueResearch(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psStruct As Long, p_research As StringVector
    Set context = wzapi_execution_context_new(ctx)
    p_psStruct = unbox_STRUCTURE(idx, argc, argv, "wzapi::pursueResearch")
    p_research = unbox_string_or_string_list(idx, argc, argv, "wzapi::pursueResearch")
    js_pursueResearch = box_bool(wzapi_pursueResearch(context, p_psStruct, p_research))
End Function

' IMPL_JS_FUNC(getResearch, wzapi::getResearch)
Private Function js_getResearch(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_researchName As String, has_player As Boolean, opt_player As Long
    Dim result As wzapi_researchResult
    Set context = wzapi_execution_context_new(ctx)
    p_researchName = unbox_string(idx, argc, argv, "wzapi::getResearch")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::getResearch")
    result = wzapi_getResearch(context, p_researchName, has_player, opt_player)
    js_getResearch = box_researchResult(result)
End Function

' IMPL_JS_FUNC(enumResearch, wzapi::enumResearch)
Private Function js_enumResearch(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim context As wzapi_execution_context, result As wzapi_researchResults
    Set context = wzapi_execution_context_new(ctx)
    result = wzapi_enumResearch(context)
    js_enumResearch = box_researchResults(result)
End Function

' IMPL_JS_FUNC(componentAvailable, wzapi::componentAvailable)
Private Function js_componentAvailable(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_componentType As String, has_componentName As Boolean, opt_componentName As String
    Set context = wzapi_execution_context_new(ctx)
    p_componentType = unbox_string(idx, argc, argv, "wzapi::componentAvailable")
    has_componentName = unbox_has(idx, argc)
    If has_componentName Then opt_componentName = unbox_string(idx, argc, argv, "wzapi::componentAvailable")
    js_componentAvailable = box_bool(wzapi_componentAvailable(context, p_componentType, has_componentName, opt_componentName))
End Function

' IMPL_JS_FUNC(addFeature, wzapi::addFeature)
Private Function js_addFeature(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_featureName As String, p_x As Long, p_y As Long
    Set context = wzapi_execution_context_new(ctx)
    p_featureName = unbox_string(idx, argc, argv, "wzapi::addFeature")
    p_x = unbox_int(idx, argc, argv, "wzapi::addFeature")
    p_y = unbox_int(idx, argc, argv, "wzapi::addFeature")
    js_addFeature = box_FEATURE(wzapi_addFeature(context, p_featureName, p_x, p_y))
End Function

' IMPL_JS_FUNC(addDroid, wzapi::addDroid)
Private Function js_addDroid(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_player As Long, p_x As Long, p_y As Long
    Dim p_templateName As String, p_body As StringVector, p_propulsion As StringVector, p_reserved1 As Long, p_reserved2 As Long
    Dim p_turrets() As StringVector, n_turrets As Long
    Set context = wzapi_execution_context_new(ctx)
    p_player = unbox_int(idx, argc, argv, "wzapi::addDroid")
    p_x = unbox_int(idx, argc, argv, "wzapi::addDroid")
    p_y = unbox_int(idx, argc, argv, "wzapi::addDroid")
    p_templateName = unbox_string(idx, argc, argv, "wzapi::addDroid")
    p_body = unbox_string_or_string_list(idx, argc, argv, "wzapi::addDroid")
    p_propulsion = unbox_string_or_string_list(idx, argc, argv, "wzapi::addDroid")
    p_reserved1 = unbox_reservedParam(idx, argc, argv, "wzapi::addDroid")
    p_reserved2 = unbox_reservedParam(idx, argc, argv, "wzapi::addDroid")
    n_turrets = unbox_va_list_string_or_string_list(idx, argc, argv, "wzapi::addDroid", p_turrets)
    js_addDroid = box_DROID(wzapi_addDroid(context, p_player, p_x, p_y, p_templateName, p_body, p_propulsion, p_reserved1, p_reserved2, p_turrets, n_turrets))
End Function

' IMPL_JS_FUNC(addDroidToTransporter, wzapi::addDroidToTransporter)
Private Function js_addDroidToTransporter(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_transporter As game_object_identifier, p_droid As game_object_identifier
    Set context = wzapi_execution_context_new(ctx)
    p_transporter = unbox_game_object_identifier(idx, argc, argv, "wzapi::addDroidToTransporter")
    p_droid = unbox_game_object_identifier(idx, argc, argv, "wzapi::addDroidToTransporter")
    js_addDroidToTransporter = box_bool(wzapi_addDroidToTransporter(context, p_transporter, p_droid))
End Function

' IMPL_JS_FUNC(makeTemplate, wzapi::makeTemplate)
Private Function js_makeTemplate(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_player As Long, p_templateName As String, p_body As StringVector
    Dim p_propulsion As StringVector, p_reserved1 As Long, p_turrets() As StringVector, n_turrets As Long, psTemplate As Long
    Set context = wzapi_execution_context_new(ctx)
    p_player = unbox_int(idx, argc, argv, "wzapi::makeTemplate")
    p_templateName = unbox_string(idx, argc, argv, "wzapi::makeTemplate")
    p_body = unbox_string_or_string_list(idx, argc, argv, "wzapi::makeTemplate")
    p_propulsion = unbox_string_or_string_list(idx, argc, argv, "wzapi::makeTemplate")
    p_reserved1 = unbox_reservedParam(idx, argc, argv, "wzapi::makeTemplate")
    n_turrets = unbox_va_list_string_or_string_list(idx, argc, argv, "wzapi::makeTemplate", p_turrets)
    psTemplate = wzapi_makeTemplate(context, p_player, p_templateName, p_body, p_propulsion, p_reserved1, p_turrets, n_turrets)
    js_makeTemplate = box_DROID_TEMPLATE(psTemplate)
    ' (the std::unique_ptr<const DROID_TEMPLATE> goes out of scope)
    If psTemplate <> 0 Then deleteTemplate psTemplate
End Function

' IMPL_JS_FUNC(buildDroid, wzapi::buildDroid)
Private Function js_buildDroid(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psFactory As Long, p_templateName As String, p_body As StringVector
    Dim p_propulsion As StringVector, p_reserved1 As Long, p_reserved2 As Long, p_turrets() As StringVector, n_turrets As Long
    Set context = wzapi_execution_context_new(ctx)
    p_psFactory = unbox_STRUCTURE(idx, argc, argv, "wzapi::buildDroid")
    p_templateName = unbox_string(idx, argc, argv, "wzapi::buildDroid")
    p_body = unbox_string_or_string_list(idx, argc, argv, "wzapi::buildDroid")
    p_propulsion = unbox_string_or_string_list(idx, argc, argv, "wzapi::buildDroid")
    p_reserved1 = unbox_reservedParam(idx, argc, argv, "wzapi::buildDroid")
    p_reserved2 = unbox_reservedParam(idx, argc, argv, "wzapi::buildDroid")
    n_turrets = unbox_va_list_string_or_string_list(idx, argc, argv, "wzapi::buildDroid", p_turrets)
    js_buildDroid = box_bool(wzapi_buildDroid(context, p_psFactory, p_templateName, p_body, p_propulsion, p_reserved1, p_reserved2, p_turrets, n_turrets))
End Function

' IMPL_JS_FUNC(enumStruct, wzapi::enumStruct)
Private Function js_enumStruct(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, has_player As Boolean, opt_player As Long, has_structureType As Boolean
    Dim opt_structureType As STRUCTURE_TYPE_or_statsName_string, has_playerFilter As Boolean, opt_playerFilter As Long, result() As Long, n As Long
    Set context = wzapi_execution_context_new(ctx)
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::enumStruct")
    has_structureType = unbox_has(idx, argc)
    If has_structureType Then opt_structureType = unbox_STRUCTURE_TYPE_or_statsName_string(idx, argc, argv, "wzapi::enumStruct")
    has_playerFilter = unbox_has(idx, argc)
    If has_playerFilter Then opt_playerFilter = unbox_int(idx, argc, argv, "wzapi::enumStruct")
    n = wzapi_enumStruct(context, has_player, opt_player, has_structureType, opt_structureType, has_playerFilter, opt_playerFilter, result)
    js_enumStruct = box_vector_STRUCTURE(result, n)
End Function

' IMPL_JS_FUNC(enumStructOffWorld, wzapi::enumStructOffWorld)
Private Function js_enumStructOffWorld(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, has_player As Boolean, opt_player As Long, has_structureType As Boolean
    Dim opt_structureType As STRUCTURE_TYPE_or_statsName_string, has_playerFilter As Boolean, opt_playerFilter As Long, result() As Long, n As Long
    Set context = wzapi_execution_context_new(ctx)
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::enumStructOffWorld")
    has_structureType = unbox_has(idx, argc)
    If has_structureType Then opt_structureType = unbox_STRUCTURE_TYPE_or_statsName_string(idx, argc, argv, "wzapi::enumStructOffWorld")
    has_playerFilter = unbox_has(idx, argc)
    If has_playerFilter Then opt_playerFilter = unbox_int(idx, argc, argv, "wzapi::enumStructOffWorld")
    n = wzapi_enumStructOffWorld(context, has_player, opt_player, has_structureType, opt_structureType, has_playerFilter, opt_playerFilter, result)
    js_enumStructOffWorld = box_vector_STRUCTURE(result, n)
End Function

' IMPL_JS_FUNC(enumFeature, wzapi::enumFeature)
Private Function js_enumFeature(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_playerFilter As Long, has_featureName As Boolean, opt_featureName As String
    Dim result() As Long, n As Long
    Set context = wzapi_execution_context_new(ctx)
    p_playerFilter = unbox_int(idx, argc, argv, "wzapi::enumFeature")
    has_featureName = unbox_has(idx, argc)
    If has_featureName Then opt_featureName = unbox_string(idx, argc, argv, "wzapi::enumFeature")
    n = wzapi_enumFeature(context, p_playerFilter, has_featureName, opt_featureName, result)
    js_enumFeature = box_vector_FEATURE(result, n)
End Function

' IMPL_JS_FUNC(enumCargo, wzapi::enumCargo)
Private Function js_enumCargo(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psDroid As Long, result() As Long, n As Long
    Set context = wzapi_execution_context_new(ctx)
    p_psDroid = unbox_DROID(idx, argc, argv, "wzapi::enumCargo")
    n = wzapi_enumCargo(context, p_psDroid, result)
    js_enumCargo = box_vector_DROID(result, n)
End Function

' IMPL_JS_FUNC(enumDroid, wzapi::enumDroid)
Private Function js_enumDroid(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, has_player As Boolean, opt_player As Long, has_droidType As Boolean
    Dim opt_droidType As Long, has_playerFilter As Boolean, opt_playerFilter As Long, result() As Long, n As Long
    Set context = wzapi_execution_context_new(ctx)
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::enumDroid")
    has_droidType = unbox_has(idx, argc)
    If has_droidType Then opt_droidType = unbox_int(idx, argc, argv, "wzapi::enumDroid")
    has_playerFilter = unbox_has(idx, argc)
    If has_playerFilter Then opt_playerFilter = unbox_int(idx, argc, argv, "wzapi::enumDroid")
    n = wzapi_enumDroid(context, has_player, opt_player, has_droidType, opt_droidType, has_playerFilter, opt_playerFilter, result)
    js_enumDroid = box_vector_DROID(result, n)
End Function

' IMPL_JS_FUNC(dump, wzapi::dump)
Private Function js_dump(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_strings As StringVector
    Set context = wzapi_execution_context_new(ctx)
    p_strings = unbox_va_list_treat_as_strings(idx, argc, argv, "wzapi::dump")
    wzapi_dump context, p_strings
    js_dump = box_no_return_value()
End Function

' IMPL_JS_FUNC(debug, wzapi::debugOutputStrings)
Private Function js_debug(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_strings As StringVector
    Set context = wzapi_execution_context_new(ctx)
    p_strings = unbox_va_list_treat_as_strings(idx, argc, argv, "wzapi::debugOutputStrings")
    wzapi_debugOutputStrings context, p_strings
    js_debug = box_no_return_value()
End Function

' IMPL_JS_FUNC(pickStructLocation, wzapi::pickStructLocation)
Private Function js_pickStructLocation(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psDroid As Long, p_structureName As String, p_startX As Long
    Dim p_startY As Long, has_maxBlockingTiles As Boolean, opt_maxBlockingTiles As Long, result As scr_position
    Set context = wzapi_execution_context_new(ctx)
    p_psDroid = unbox_DROID(idx, argc, argv, "wzapi::pickStructLocation")
    p_structureName = unbox_string(idx, argc, argv, "wzapi::pickStructLocation")
    p_startX = unbox_int(idx, argc, argv, "wzapi::pickStructLocation")
    p_startY = unbox_int(idx, argc, argv, "wzapi::pickStructLocation")
    has_maxBlockingTiles = unbox_has(idx, argc)
    If has_maxBlockingTiles Then opt_maxBlockingTiles = unbox_int(idx, argc, argv, "wzapi::pickStructLocation")
    If wzapi_pickStructLocation(context, p_psDroid, p_structureName, p_startX, p_startY, has_maxBlockingTiles, opt_maxBlockingTiles, result) Then
        js_pickStructLocation = box_scr_position(result)
    Else
        js_pickStructLocation = jsUndefinedValue()
    End If
End Function

' IMPL_JS_FUNC(structureCanFit, wzapi::structureCanFit)
Private Function js_structureCanFit(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_structureName As String, p_x As Long, p_y As Long
    Dim has_direction As Boolean, opt_direction As Single
    Set context = wzapi_execution_context_new(ctx)
    p_structureName = unbox_string(idx, argc, argv, "wzapi::structureCanFit")
    p_x = unbox_int(idx, argc, argv, "wzapi::structureCanFit")
    p_y = unbox_int(idx, argc, argv, "wzapi::structureCanFit")
    has_direction = unbox_has(idx, argc)
    If has_direction Then opt_direction = unbox_float(idx, argc, argv, "wzapi::structureCanFit")
    js_structureCanFit = box_bool(wzapi_structureCanFit(context, p_structureName, p_x, p_y, has_direction, opt_direction))
End Function

' IMPL_JS_FUNC(structureIdle, wzapi::structureIdle)
Private Function js_structureIdle(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psStruct As Long
    Set context = wzapi_execution_context_new(ctx)
    p_psStruct = unbox_STRUCTURE(idx, argc, argv, "wzapi::structureIdle")
    js_structureIdle = box_bool(wzapi_structureIdle(context, p_psStruct))
End Function

' IMPL_JS_FUNC(removeStruct, wzapi::removeStruct)
Private Function js_removeStruct(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psStruct As Long
    Set context = wzapi_execution_context_new(ctx)
    p_psStruct = unbox_STRUCTURE(idx, argc, argv, "wzapi::removeStruct")
    js_removeStruct = box_bool(wzapi_removeStruct(context, p_psStruct))
End Function

' IMPL_JS_FUNC(removeObject, wzapi::removeObject)
Private Function js_removeObject(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psObj As Long, has_sfx As Boolean, opt_sfx As Boolean
    Set context = wzapi_execution_context_new(ctx)
    p_psObj = unbox_BASE_OBJECT(idx, argc, argv, "wzapi::removeObject")
    has_sfx = unbox_has(idx, argc)
    If has_sfx Then opt_sfx = unbox_bool(idx, argc, argv, "wzapi::removeObject")
    js_removeObject = box_bool(wzapi_removeObject(context, p_psObj, has_sfx, opt_sfx))
End Function

' IMPL_JS_FUNC(clearConsole, wzapi::clearConsole)
Private Function js_clearConsole(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    js_clearConsole = box_bool(wzapi_clearConsole(context))
End Function

' IMPL_JS_FUNC(console, wzapi::console)
Private Function js_console(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_strings As StringVector
    Set context = wzapi_execution_context_new(ctx)
    p_strings = unbox_va_list_treat_as_strings(idx, argc, argv, "wzapi::console")
    js_console = box_bool(wzapi_console(context, p_strings))
End Function

' IMPL_JS_FUNC(distBetweenTwoPoints, wzapi::distBetweenTwoPoints)
Private Function js_distBetweenTwoPoints(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_x1 As Long, p_y1 As Long, p_x2 As Long
    Dim p_y2 As Long
    Set context = wzapi_execution_context_new(ctx)
    p_x1 = unbox_int(idx, argc, argv, "wzapi::distBetweenTwoPoints")
    p_y1 = unbox_int(idx, argc, argv, "wzapi::distBetweenTwoPoints")
    p_x2 = unbox_int(idx, argc, argv, "wzapi::distBetweenTwoPoints")
    p_y2 = unbox_int(idx, argc, argv, "wzapi::distBetweenTwoPoints")
    js_distBetweenTwoPoints = box_int(wzapi_distBetweenTwoPoints(context, p_x1, p_y1, p_x2, p_y2))
End Function

' IMPL_JS_FUNC(droidCanReach, wzapi::droidCanReach)
Private Function js_droidCanReach(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psDroid As Long, p_x As Long, p_y As Long
    Set context = wzapi_execution_context_new(ctx)
    p_psDroid = unbox_DROID(idx, argc, argv, "wzapi::droidCanReach")
    p_x = unbox_int(idx, argc, argv, "wzapi::droidCanReach")
    p_y = unbox_int(idx, argc, argv, "wzapi::droidCanReach")
    js_droidCanReach = box_bool(wzapi_droidCanReach(context, p_psDroid, p_x, p_y))
End Function

' IMPL_JS_FUNC(propulsionCanReach, wzapi::propulsionCanReach)
Private Function js_propulsionCanReach(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_propulsionName As String, p_x1 As Long, p_y1 As Long
    Dim p_x2 As Long, p_y2 As Long
    Set context = wzapi_execution_context_new(ctx)
    p_propulsionName = unbox_string(idx, argc, argv, "wzapi::propulsionCanReach")
    p_x1 = unbox_int(idx, argc, argv, "wzapi::propulsionCanReach")
    p_y1 = unbox_int(idx, argc, argv, "wzapi::propulsionCanReach")
    p_x2 = unbox_int(idx, argc, argv, "wzapi::propulsionCanReach")
    p_y2 = unbox_int(idx, argc, argv, "wzapi::propulsionCanReach")
    js_propulsionCanReach = box_bool(wzapi_propulsionCanReach(context, p_propulsionName, p_x1, p_y1, p_x2, p_y2))
End Function

' IMPL_JS_FUNC(terrainType, wzapi::terrainType)
Private Function js_terrainType(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_x As Long, p_y As Long
    Set context = wzapi_execution_context_new(ctx)
    p_x = unbox_int(idx, argc, argv, "wzapi::terrainType")
    p_y = unbox_int(idx, argc, argv, "wzapi::terrainType")
    js_terrainType = box_int(wzapi_terrainType(context, p_x, p_y))
End Function

' IMPL_JS_FUNC(tileIsBurning, wzapi::tileIsBurning)
Private Function js_tileIsBurning(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_x As Long, p_y As Long
    Set context = wzapi_execution_context_new(ctx)
    p_x = unbox_int(idx, argc, argv, "wzapi::tileIsBurning")
    p_y = unbox_int(idx, argc, argv, "wzapi::tileIsBurning")
    js_tileIsBurning = box_bool(wzapi_tileIsBurning(context, p_x, p_y))
End Function

' IMPL_JS_FUNC(orderDroid, wzapi::orderDroid)
Private Function js_orderDroid(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psDroid As Long, p_order As Long
    Set context = wzapi_execution_context_new(ctx)
    p_psDroid = unbox_DROID(idx, argc, argv, "wzapi::orderDroid")
    p_order = unbox_int(idx, argc, argv, "wzapi::orderDroid")
    js_orderDroid = box_bool(wzapi_orderDroid(context, p_psDroid, p_order))
End Function

' IMPL_JS_FUNC(orderDroidObj, wzapi::orderDroidObj)
Private Function js_orderDroidObj(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psDroid As Long, p_order As Long, p_psObj As Long
    Set context = wzapi_execution_context_new(ctx)
    p_psDroid = unbox_DROID(idx, argc, argv, "wzapi::orderDroidObj")
    p_order = unbox_int(idx, argc, argv, "wzapi::orderDroidObj")
    p_psObj = unbox_BASE_OBJECT(idx, argc, argv, "wzapi::orderDroidObj")
    js_orderDroidObj = box_bool(wzapi_orderDroidObj(context, p_psDroid, p_order, p_psObj))
End Function

' IMPL_JS_FUNC(orderDroidBuild, wzapi::orderDroidBuild)
Private Function js_orderDroidBuild(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psDroid As Long, p_order As Long, p_structureName As String
    Dim p_x As Long, p_y As Long, has_direction As Boolean, opt_direction As Single
    Set context = wzapi_execution_context_new(ctx)
    p_psDroid = unbox_DROID(idx, argc, argv, "wzapi::orderDroidBuild")
    p_order = unbox_int(idx, argc, argv, "wzapi::orderDroidBuild")
    p_structureName = unbox_string(idx, argc, argv, "wzapi::orderDroidBuild")
    p_x = unbox_int(idx, argc, argv, "wzapi::orderDroidBuild")
    p_y = unbox_int(idx, argc, argv, "wzapi::orderDroidBuild")
    has_direction = unbox_has(idx, argc)
    If has_direction Then opt_direction = unbox_float(idx, argc, argv, "wzapi::orderDroidBuild")
    js_orderDroidBuild = box_bool(wzapi_orderDroidBuild(context, p_psDroid, p_order, p_structureName, p_x, p_y, has_direction, opt_direction))
End Function

' IMPL_JS_FUNC(orderDroidLoc, wzapi::orderDroidLoc)
Private Function js_orderDroidLoc(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_psDroid As Long, p_order As Long, p_x As Long
    Dim p_y As Long
    Set context = wzapi_execution_context_new(ctx)
    p_psDroid = unbox_DROID(idx, argc, argv, "wzapi::orderDroidLoc")
    p_order = unbox_int(idx, argc, argv, "wzapi::orderDroidLoc")
    p_x = unbox_int(idx, argc, argv, "wzapi::orderDroidLoc")
    p_y = unbox_int(idx, argc, argv, "wzapi::orderDroidLoc")
    js_orderDroidLoc = box_bool(wzapi_orderDroidLoc(context, p_psDroid, p_order, p_x, p_y))
End Function

' IMPL_JS_FUNC(setMissionTime, wzapi::setMissionTime)
Private Function js_setMissionTime(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_time As Long
    Set context = wzapi_execution_context_new(ctx)
    p_time = unbox_int(idx, argc, argv, "wzapi::setMissionTime")
    wzapi_setMissionTime context, p_time
    js_setMissionTime = box_no_return_value()
End Function

' IMPL_JS_FUNC(getMissionTime, wzapi::getMissionTime)
Private Function js_getMissionTime(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    js_getMissionTime = box_int(wzapi_getMissionTime(context))
End Function

' IMPL_JS_FUNC(setTransporterExit, wzapi::setTransporterExit)
Private Function js_setTransporterExit(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_x As Long, p_y As Long, p_player As Long
    Set context = wzapi_execution_context_new(ctx)
    p_x = unbox_int(idx, argc, argv, "wzapi::setTransporterExit")
    p_y = unbox_int(idx, argc, argv, "wzapi::setTransporterExit")
    p_player = unbox_int(idx, argc, argv, "wzapi::setTransporterExit")
    wzapi_setTransporterExit context, p_x, p_y, p_player
    js_setTransporterExit = box_no_return_value()
End Function

' IMPL_JS_FUNC(startTransporterEntry, wzapi::startTransporterEntry)
Private Function js_startTransporterEntry(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_x As Long, p_y As Long, p_player As Long
    Set context = wzapi_execution_context_new(ctx)
    p_x = unbox_int(idx, argc, argv, "wzapi::startTransporterEntry")
    p_y = unbox_int(idx, argc, argv, "wzapi::startTransporterEntry")
    p_player = unbox_int(idx, argc, argv, "wzapi::startTransporterEntry")
    wzapi_startTransporterEntry context, p_x, p_y, p_player
    js_startTransporterEntry = box_no_return_value()
End Function

' IMPL_JS_FUNC(useSafetyTransport, wzapi::useSafetyTransport)
Private Function js_useSafetyTransport(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_flag As Boolean
    Set context = wzapi_execution_context_new(ctx)
    p_flag = unbox_bool(idx, argc, argv, "wzapi::useSafetyTransport")
    js_useSafetyTransport = box_bool(wzapi_useSafetyTransport(context, p_flag))
End Function

' IMPL_JS_FUNC(restoreLimboMissionData, wzapi::restoreLimboMissionData)
Private Function js_restoreLimboMissionData(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    js_restoreLimboMissionData = box_bool(wzapi_restoreLimboMissionData(context))
End Function

' IMPL_JS_FUNC(setReinforcementTime, wzapi::setReinforcementTime)
Private Function js_setReinforcementTime(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_time As Long, has_removeLaunch As Boolean, opt_removeLaunch As Boolean
    Set context = wzapi_execution_context_new(ctx)
    p_time = unbox_int(idx, argc, argv, "wzapi::setReinforcementTime")
    has_removeLaunch = unbox_has(idx, argc)
    If has_removeLaunch Then opt_removeLaunch = unbox_bool(idx, argc, argv, "wzapi::setReinforcementTime")
    wzapi_setReinforcementTime context, p_time, has_removeLaunch, opt_removeLaunch
    js_setReinforcementTime = box_no_return_value()
End Function

' IMPL_JS_FUNC(setStructureLimits, wzapi::setStructureLimits)
Private Function js_setStructureLimits(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_structureName As String, p_limit As Long, has_player As Boolean
    Dim opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    p_structureName = unbox_string(idx, argc, argv, "wzapi::setStructureLimits")
    p_limit = unbox_int(idx, argc, argv, "wzapi::setStructureLimits")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::setStructureLimits")
    js_setStructureLimits = box_bool(wzapi_setStructureLimits(context, p_structureName, p_limit, has_player, opt_player))
End Function

' IMPL_JS_FUNC(centreView, wzapi::centreView)
Private Function js_centreView(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_x As Long, p_y As Long
    Set context = wzapi_execution_context_new(ctx)
    p_x = unbox_int(idx, argc, argv, "wzapi::centreView")
    p_y = unbox_int(idx, argc, argv, "wzapi::centreView")
    js_centreView = box_bool(wzapi_centreView(context, p_x, p_y))
End Function

' IMPL_JS_FUNC(hackPlayIngameAudio, wzapi::hackPlayIngameAudio)
Private Function js_hackPlayIngameAudio(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    wzapi_hackPlayIngameAudio context
    js_hackPlayIngameAudio = box_no_return_value()
End Function

' IMPL_JS_FUNC(hackStopIngameAudio, wzapi::hackStopIngameAudio)
Private Function js_hackStopIngameAudio(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    wzapi_hackStopIngameAudio context
    js_hackStopIngameAudio = box_no_return_value()
End Function

' IMPL_JS_FUNC(playSound, wzapi::playSound)
Private Function js_playSound(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_sound As String, has_x As Boolean, opt_x As Long
    Dim has_y As Boolean, opt_y As Long, has_z As Boolean, opt_z As Long
    Set context = wzapi_execution_context_new(ctx)
    p_sound = unbox_string(idx, argc, argv, "wzapi::playSound")
    has_x = unbox_has(idx, argc)
    If has_x Then opt_x = unbox_int(idx, argc, argv, "wzapi::playSound")
    has_y = unbox_has(idx, argc)
    If has_y Then opt_y = unbox_int(idx, argc, argv, "wzapi::playSound")
    has_z = unbox_has(idx, argc)
    If has_z Then opt_z = unbox_int(idx, argc, argv, "wzapi::playSound")
    js_playSound = box_bool(wzapi_playSound(context, p_sound, has_x, opt_x, has_y, opt_y, has_z, opt_z))
End Function

' IMPL_JS_FUNC(emitSound, wzapi::emitSound)
Private Function js_emitSound(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_sound As String, p_x As Long, p_y As Long
    Set context = wzapi_execution_context_new(ctx)
    p_sound = unbox_string(idx, argc, argv, "wzapi::emitSound")
    p_x = unbox_int(idx, argc, argv, "wzapi::emitSound")
    p_y = unbox_int(idx, argc, argv, "wzapi::emitSound")
    js_emitSound = box_bool(wzapi_emitSound(context, p_sound, p_x, p_y))
End Function

' IMPL_JS_FUNC_DEBUGMSGUPDATE(gameOverMessage, wzapi::gameOverMessage)
Private Function js_gameOverMessage(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_gameWon As Boolean, has_showBackDrop As Boolean, opt_showBackDrop As Boolean
    Dim has_showOutro As Boolean, opt_showOutro As Boolean
    Set context = wzapi_execution_context_new(ctx)
    p_gameWon = unbox_bool(idx, argc, argv, "wzapi::gameOverMessage")
    has_showBackDrop = unbox_has(idx, argc)
    If has_showBackDrop Then opt_showBackDrop = unbox_bool(idx, argc, argv, "wzapi::gameOverMessage")
    has_showOutro = unbox_has(idx, argc)
    If has_showOutro Then opt_showOutro = unbox_bool(idx, argc, argv, "wzapi::gameOverMessage")
    js_gameOverMessage = box_bool(wzapi_gameOverMessage(context, p_gameWon, has_showBackDrop, opt_showBackDrop, has_showOutro, opt_showOutro))
    jsDebugMessageUpdate
End Function

' IMPL_JS_FUNC(completeResearch, wzapi::completeResearch)
Private Function js_completeResearch(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_researchName As String, has_player As Boolean, opt_player As Long
    Dim has_forceResearch As Boolean, opt_forceResearch As Boolean
    Set context = wzapi_execution_context_new(ctx)
    p_researchName = unbox_string(idx, argc, argv, "wzapi::completeResearch")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::completeResearch")
    has_forceResearch = unbox_has(idx, argc)
    If has_forceResearch Then opt_forceResearch = unbox_bool(idx, argc, argv, "wzapi::completeResearch")
    wzapi_completeResearch context, p_researchName, has_player, opt_player, has_forceResearch, opt_forceResearch
    js_completeResearch = box_no_return_value()
End Function

' IMPL_JS_FUNC(completeAllResearch, wzapi::completeAllResearch)
Private Function js_completeAllResearch(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::completeAllResearch")
    wzapi_completeAllResearch context, has_player, opt_player
    js_completeAllResearch = box_no_return_value()
End Function

' IMPL_JS_FUNC(enableResearch, wzapi::enableResearch)
Private Function js_enableResearch(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_researchName As String, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    p_researchName = unbox_string(idx, argc, argv, "wzapi::enableResearch")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::enableResearch")
    js_enableResearch = box_bool(wzapi_enableResearch(context, p_researchName, has_player, opt_player))
End Function

' IMPL_JS_FUNC(extraPowerTime, wzapi::extraPowerTime)
Private Function js_extraPowerTime(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_time As Long, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    p_time = unbox_int(idx, argc, argv, "wzapi::extraPowerTime")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::extraPowerTime")
    wzapi_extraPowerTime context, p_time, has_player, opt_player
    js_extraPowerTime = box_no_return_value()
End Function

' IMPL_JS_FUNC(setPower, wzapi::setPower)
Private Function js_setPower(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_power As Long, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    p_power = unbox_int(idx, argc, argv, "wzapi::setPower")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::setPower")
    wzapi_setPower context, p_power, has_player, opt_player
    js_setPower = box_no_return_value()
End Function

' IMPL_JS_FUNC(setPowerModifier, wzapi::setPowerModifier)
Private Function js_setPowerModifier(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_powerModifier As Long, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    p_powerModifier = unbox_int(idx, argc, argv, "wzapi::setPowerModifier")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::setPowerModifier")
    wzapi_setPowerModifier context, p_powerModifier, has_player, opt_player
    js_setPowerModifier = box_no_return_value()
End Function

' IMPL_JS_FUNC(setPowerStorageMaximum, wzapi::setPowerStorageMaximum)
Private Function js_setPowerStorageMaximum(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_powerMaximum As Long, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    p_powerMaximum = unbox_int(idx, argc, argv, "wzapi::setPowerStorageMaximum")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::setPowerStorageMaximum")
    wzapi_setPowerStorageMaximum context, p_powerMaximum, has_player, opt_player
    js_setPowerStorageMaximum = box_no_return_value()
End Function

' IMPL_JS_FUNC(enableStructure, wzapi::enableStructure)
Private Function js_enableStructure(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_structureName As String, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    p_structureName = unbox_string(idx, argc, argv, "wzapi::enableStructure")
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, "wzapi::enableStructure")
    wzapi_enableStructure context, p_structureName, has_player, opt_player
    js_enableStructure = box_no_return_value()
End Function

' IMPL_JS_FUNC(setTutorialMode, wzapi::setTutorialMode)
Private Function js_setTutorialMode(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_enableTutorialMode As Boolean
    Set context = wzapi_execution_context_new(ctx)
    p_enableTutorialMode = unbox_bool(idx, argc, argv, "wzapi::setTutorialMode")
    wzapi_setTutorialMode context, p_enableTutorialMode
    js_setTutorialMode = box_no_return_value()
End Function

' IMPL_JS_FUNC(setMiniMap, wzapi::setMiniMap)
Private Function js_setMiniMap(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_visible As Boolean
    Set context = wzapi_execution_context_new(ctx)
    p_visible = unbox_bool(idx, argc, argv, "wzapi::setMiniMap")
    wzapi_setMiniMap context, p_visible
    js_setMiniMap = box_no_return_value()
End Function

' IMPL_JS_FUNC(setDesign, wzapi::setDesign)
Private Function js_setDesign(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_allowDesignValue As Boolean
    Set context = wzapi_execution_context_new(ctx)
    p_allowDesignValue = unbox_bool(idx, argc, argv, "wzapi::setDesign")
    wzapi_setDesign context, p_allowDesignValue
    js_setDesign = box_no_return_value()
End Function

' IMPL_JS_FUNC(enableTemplate, wzapi::enableTemplate)
Private Function js_enableTemplate(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_templateName As String
    Set context = wzapi_execution_context_new(ctx)
    p_templateName = unbox_string(idx, argc, argv, "wzapi::enableTemplate")
    js_enableTemplate = box_bool(wzapi_enableTemplate(context, p_templateName))
End Function

' IMPL_JS_FUNC(removeTemplate, wzapi::removeTemplate)
Private Function js_removeTemplate(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_templateName As String
    Set context = wzapi_execution_context_new(ctx)
    p_templateName = unbox_string(idx, argc, argv, "wzapi::removeTemplate")
    js_removeTemplate = box_bool(wzapi_removeTemplate(context, p_templateName))
End Function

' IMPL_JS_FUNC(setReticuleButton, wzapi::setReticuleButton)
Private Function js_setReticuleButton(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_buttonId As Long, p_tooltip As String, p_filename As String
    Dim p_filenameDown As String, has_callbackFuncName As Boolean, opt_callbackFuncName As String
    Set context = wzapi_execution_context_new(ctx)
    p_buttonId = unbox_int(idx, argc, argv, "wzapi::setReticuleButton")
    p_tooltip = unbox_string(idx, argc, argv, "wzapi::setReticuleButton")
    p_filename = unbox_string(idx, argc, argv, "wzapi::setReticuleButton")
    p_filenameDown = unbox_string(idx, argc, argv, "wzapi::setReticuleButton")
    has_callbackFuncName = unbox_has(idx, argc)
    If has_callbackFuncName Then opt_callbackFuncName = unbox_string(idx, argc, argv, "wzapi::setReticuleButton")
    wzapi_setReticuleButton context, p_buttonId, p_tooltip, p_filename, p_filenameDown, has_callbackFuncName, opt_callbackFuncName
    js_setReticuleButton = box_no_return_value()
End Function

' IMPL_JS_FUNC(showReticuleWidget, wzapi::showReticuleWidget)
Private Function js_showReticuleWidget(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_buttonId As Long
    Set context = wzapi_execution_context_new(ctx)
    p_buttonId = unbox_int(idx, argc, argv, "wzapi::showReticuleWidget")
    wzapi_showReticuleWidget context, p_buttonId
    js_showReticuleWidget = box_no_return_value()
End Function

' quickjs_backend.cpp:3244: IMPL_JS_FUNC(setReticuleFlash, wzapi::setReticuleFlash)
Private Function js_setReticuleFlash(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setReticuleFlash"
    Dim idx As Long, context As wzapi_execution_context, buttonId As Long, flash As Boolean
    Set context = wzapi_execution_context_new(ctx)
    buttonId = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    flash = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_setReticuleFlash context, buttonId, flash
    js_setReticuleFlash = box_no_return_value()
End Function

' quickjs_backend.cpp:3245: IMPL_JS_FUNC(showInterface, wzapi::showInterface)
Private Function js_showInterface(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::showInterface"
    Dim idx As Long, context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    wzapi_showInterface context
    js_showInterface = box_no_return_value()
End Function

' quickjs_backend.cpp:3246: IMPL_JS_FUNC(hideInterface, wzapi::hideInterface)
Private Function js_hideInterface(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::hideInterface"
    Dim idx As Long, context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    wzapi_hideInterface context
    js_hideInterface = box_no_return_value()
End Function

' quickjs_backend.cpp:3247: IMPL_JS_FUNC(addGuideTopic, wzapi::addGuideTopic)
Private Function js_addGuideTopic(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::addGuideTopic"
    Dim idx As Long, context As wzapi_execution_context, guideTopicID As String, has_showFlags As Boolean, opt_showFlags As Long, has_excludedTopicIDs As Boolean
    Dim opt_excludedTopicIDs As StringVector
    Set context = wzapi_execution_context_new(ctx)
    guideTopicID = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    has_showFlags = unbox_has(idx, argc)
    If has_showFlags Then opt_showFlags = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_excludedTopicIDs = unbox_has(idx, argc)
    If has_excludedTopicIDs Then opt_excludedTopicIDs = unbox_string_or_string_list(idx, argc, argv, WRAPPED_FUNCTION)
    js_addGuideTopic = box_bool(wzapi_addGuideTopic(context, guideTopicID, has_showFlags, opt_showFlags, has_excludedTopicIDs, opt_excludedTopicIDs))
End Function

' quickjs_backend.cpp:3258: IMPL_JS_FUNC(applyLimitSet, wzapi::applyLimitSet)
Private Function js_applyLimitSet(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::applyLimitSet"
    Dim idx As Long, context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    js_applyLimitSet = box_bool(wzapi_applyLimitSet(context))
End Function

' quickjs_backend.cpp:3259: IMPL_JS_FUNC(enableComponent, wzapi::enableComponent)
Private Function js_enableComponent(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::enableComponent"
    Dim idx As Long, context As wzapi_execution_context, componentName As String, player As Long
    Set context = wzapi_execution_context_new(ctx)
    componentName = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_enableComponent context, componentName, player
    js_enableComponent = box_no_return_value()
End Function

' quickjs_backend.cpp:3260: IMPL_JS_FUNC(makeComponentAvailable, wzapi::makeComponentAvailable)
Private Function js_makeComponentAvailable(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::makeComponentAvailable"
    Dim idx As Long, context As wzapi_execution_context, componentName As String, player As Long
    Set context = wzapi_execution_context_new(ctx)
    componentName = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_makeComponentAvailable context, componentName, player
    js_makeComponentAvailable = box_no_return_value()
End Function

' quickjs_backend.cpp:3261: IMPL_JS_FUNC(allianceExistsBetween, wzapi::allianceExistsBetween)
Private Function js_allianceExistsBetween(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::allianceExistsBetween"
    Dim idx As Long, context As wzapi_execution_context, player1 As Long, player2 As Long
    Set context = wzapi_execution_context_new(ctx)
    player1 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    player2 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_allianceExistsBetween = box_bool(wzapi_allianceExistsBetween(context, player1, player2))
End Function

' quickjs_backend.cpp:3262: IMPL_JS_FUNC(translate, wzapi::translate)
Private Function js_translate(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::translate"
    Dim idx As Long, context As wzapi_execution_context, str_ As String
    Set context = wzapi_execution_context_new(ctx)
    str_ = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    js_translate = box_string(wzapi_translate(context, str_))
End Function

' quickjs_backend.cpp:3263: IMPL_JS_FUNC(playerPower, wzapi::playerPower)
Private Function js_playerPower(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::playerPower"
    Dim idx As Long, context As wzapi_execution_context, player As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_playerPower = box_int(wzapi_playerPower(context, player))
End Function

' quickjs_backend.cpp:3264: IMPL_JS_FUNC(queuedPower, wzapi::queuedPower)
Private Function js_queuedPower(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::queuedPower"
    Dim idx As Long, context As wzapi_execution_context, player As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_queuedPower = box_int(wzapi_queuedPower(context, player))
End Function

' quickjs_backend.cpp:3265: IMPL_JS_FUNC(isStructureAvailable, wzapi::isStructureAvailable)
Private Function js_isStructureAvailable(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::isStructureAvailable"
    Dim idx As Long, context As wzapi_execution_context, structureName As String, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    structureName = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_isStructureAvailable = box_bool(wzapi_isStructureAvailable(context, structureName, has_player, opt_player))
End Function

' quickjs_backend.cpp:3266: IMPL_JS_FUNC(isVTOL, wzapi::isVTOL)
Private Function js_isVTOL(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::isVTOL"
    Dim idx As Long, context As wzapi_execution_context, psDroid As Long
    Set context = wzapi_execution_context_new(ctx)
    psDroid = unbox_DROID(idx, argc, argv, WRAPPED_FUNCTION)
    js_isVTOL = box_bool(wzapi_isVTOL(context, psDroid))
End Function

' quickjs_backend.cpp:3267: IMPL_JS_FUNC(hackGetObj, wzapi::hackGetObj)
Private Function js_hackGetObj(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::hackGetObj"
    Dim idx As Long, context As wzapi_execution_context, objectType_ As Long, player As Long, id As Long
    Set context = wzapi_execution_context_new(ctx)
    objectType_ = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    id = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_hackGetObj = box_BASE_OBJECT(wzapi_hackGetObj(context, objectType_, player, id))
End Function

' quickjs_backend.cpp:3268: IMPL_JS_FUNC(receiveAllEvents, wzapi::receiveAllEvents)
Private Function js_receiveAllEvents(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::receiveAllEvents"
    Dim idx As Long, context As wzapi_execution_context, has_enabled As Boolean, opt_enabled As Boolean
    Set context = wzapi_execution_context_new(ctx)
    has_enabled = unbox_has(idx, argc)
    If has_enabled Then opt_enabled = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    js_receiveAllEvents = box_bool(wzapi_receiveAllEvents(context, has_enabled, opt_enabled))
End Function

' quickjs_backend.cpp:3269: IMPL_JS_FUNC(hackAssert, wzapi::hackAssert)
Private Function js_hackAssert(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::hackAssert"
    Dim idx As Long, context As wzapi_execution_context, condition As Boolean, message As StringVector
    Set context = wzapi_execution_context_new(ctx)
    condition = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    message = unbox_va_list_treat_as_strings(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_hackAssert context, condition, message
    js_hackAssert = box_no_return_value()
End Function

' quickjs_backend.cpp:3270: IMPL_JS_FUNC(setDroidExperience, wzapi::setDroidExperience)
Private Function js_setDroidExperience(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setDroidExperience"
    Dim idx As Long, context As wzapi_execution_context, psDroid As Long, experience As Double
    Set context = wzapi_execution_context_new(ctx)
    psDroid = unbox_DROID(idx, argc, argv, WRAPPED_FUNCTION)
    experience = unbox_double(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_setDroidExperience context, psDroid, experience
    js_setDroidExperience = box_no_return_value()
End Function

' quickjs_backend.cpp:3271: IMPL_JS_FUNC(donateObject, wzapi::donateObject)
Private Function js_donateObject(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::donateObject"
    Dim idx As Long, context As wzapi_execution_context, psObject As Long, player As Long
    Set context = wzapi_execution_context_new(ctx)
    psObject = unbox_BASE_OBJECT(idx, argc, argv, WRAPPED_FUNCTION)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_donateObject = box_bool(wzapi_donateObject(context, psObject, player))
End Function

' quickjs_backend.cpp:3272: IMPL_JS_FUNC(donatePower, wzapi::donatePower)
Private Function js_donatePower(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::donatePower"
    Dim idx As Long, context As wzapi_execution_context, amount As Long, player As Long
    Set context = wzapi_execution_context_new(ctx)
    amount = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_donatePower = box_bool(wzapi_donatePower(context, amount, player))
End Function

' quickjs_backend.cpp:3273: IMPL_JS_FUNC(safeDest, wzapi::safeDest)
Private Function js_safeDest(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::safeDest"
    Dim idx As Long, context As wzapi_execution_context, player As Long, x As Long, y As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    x = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_safeDest = box_bool(wzapi_safeDest(context, player, x, y))
End Function

' quickjs_backend.cpp:3274: IMPL_JS_FUNC(addStructure, wzapi::addStructure)
Private Function js_addStructure(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::addStructure"
    Dim idx As Long, context As wzapi_execution_context, structureName As String, player As Long, x As Long, y As Long
    Dim has_direction As Boolean, opt_direction As Long
    Set context = wzapi_execution_context_new(ctx)
    structureName = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    x = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_direction = unbox_has(idx, argc)
    If has_direction Then opt_direction = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_addStructure = box_STRUCTURE(wzapi_addStructure(context, structureName, player, x, y, has_direction, opt_direction))
End Function

' quickjs_backend.cpp:3275: IMPL_JS_FUNC(getStructureLimit, wzapi::getStructureLimit)
Private Function js_getStructureLimit(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::getStructureLimit"
    Dim idx As Long, context As wzapi_execution_context, structureName As String, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    structureName = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_getStructureLimit = box_uint(wzapi_getStructureLimit(context, structureName, has_player, opt_player))
End Function

' quickjs_backend.cpp:3276: IMPL_JS_FUNC(countStruct, wzapi::countStruct)
Private Function js_countStruct(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::countStruct"
    Dim idx As Long, context As wzapi_execution_context, structureName As String, has_playerFilter As Boolean, opt_playerFilter As Long
    Set context = wzapi_execution_context_new(ctx)
    structureName = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    has_playerFilter = unbox_has(idx, argc)
    If has_playerFilter Then opt_playerFilter = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_countStruct = box_int(wzapi_countStruct(context, structureName, has_playerFilter, opt_playerFilter))
End Function

' quickjs_backend.cpp:3277: IMPL_JS_FUNC(countDroid, wzapi::countDroid)
Private Function js_countDroid(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::countDroid"
    Dim idx As Long, context As wzapi_execution_context, has_droidType As Boolean, opt_droidType As Long, has_playerFilter As Boolean, opt_playerFilter As Long
    Set context = wzapi_execution_context_new(ctx)
    has_droidType = unbox_has(idx, argc)
    If has_droidType Then opt_droidType = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_playerFilter = unbox_has(idx, argc)
    If has_playerFilter Then opt_playerFilter = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_countDroid = box_int(wzapi_countDroid(context, has_droidType, opt_droidType, has_playerFilter, opt_playerFilter))
End Function

' quickjs_backend.cpp:3278: IMPL_JS_FUNC(setNoGoArea, wzapi::setNoGoArea)
Private Function js_setNoGoArea(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setNoGoArea"
    Dim idx As Long, context As wzapi_execution_context, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim playerFilter As Long
    Set context = wzapi_execution_context_new(ctx)
    x1 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y1 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    x2 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y2 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    playerFilter = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_setNoGoArea context, x1, y1, x2, y2, playerFilter
    js_setNoGoArea = box_no_return_value()
End Function

' quickjs_backend.cpp:3279: IMPL_JS_FUNC(setScrollLimits, wzapi::setScrollLimits)
Private Function js_setScrollLimits(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setScrollLimits"
    Dim idx As Long, context As wzapi_execution_context, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Set context = wzapi_execution_context_new(ctx)
    x1 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y1 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    x2 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y2 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_setScrollLimits context, x1, y1, x2, y2
    js_setScrollLimits = box_no_return_value()
End Function

' quickjs_backend.cpp:3280: IMPL_JS_FUNC(getScrollLimits, wzapi::getScrollLimits)
Private Function js_getScrollLimits(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::getScrollLimits"
    Dim idx As Long, context As wzapi_execution_context, result As scr_area
    Set context = wzapi_execution_context_new(ctx)
    result = wzapi_getScrollLimits(context)
    js_getScrollLimits = box_scr_area(result)
End Function

' quickjs_backend.cpp:3281: IMPL_JS_FUNC(loadLevel, wzapi::loadLevel)
Private Function js_loadLevel(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::loadLevel"
    Dim idx As Long, context As wzapi_execution_context, levelName As String
    Set context = wzapi_execution_context_new(ctx)
    levelName = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_loadLevel context, levelName
    js_loadLevel = box_no_return_value()
End Function

' quickjs_backend.cpp:3282: IMPL_JS_FUNC(autoSave, wzapi::autoSave)
Private Function js_autoSave(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::autoSave"
    Dim idx As Long, context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    js_autoSave = box_bool(wzapi_autoSave(context))
End Function

' quickjs_backend.cpp:3283: IMPL_JS_FUNC(enumRange, wzapi::enumRange)
Private Function js_enumRange(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::enumRange"
    Dim idx As Long, context As wzapi_execution_context, x As Long, y As Long, range_ As Long, has_playerFilter As Boolean
    Dim opt_playerFilter As Long, has_seen As Boolean, opt_seen As Boolean, result() As Long, n As Long
    Set context = wzapi_execution_context_new(ctx)
    x = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    range_ = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_playerFilter = unbox_has(idx, argc)
    If has_playerFilter Then opt_playerFilter = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_seen = unbox_has(idx, argc)
    If has_seen Then opt_seen = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    n = wzapi_enumRange(context, x, y, range_, has_playerFilter, opt_playerFilter, has_seen, opt_seen, result)
    js_enumRange = box_vector_BASE_OBJECT(result, n)
End Function

' quickjs_backend.cpp:3284: IMPL_JS_FUNC(enumArea, scripting_engine::enumAreaJS)
Private Function js_enumArea(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "scripting_engine::enumAreaJS"
    Dim idx As Long, context As wzapi_execution_context, area_lookup As area_by_values_or_area_label_lookup, has_playerFilter As Boolean, opt_playerFilter As Long, has_seen As Boolean
    Dim opt_seen As Boolean, result() As Long, n As Long
    Set context = wzapi_execution_context_new(ctx)
    area_lookup = unbox_area_by_values_or_area_label_lookup(idx, argc, argv, WRAPPED_FUNCTION)
    has_playerFilter = unbox_has(idx, argc)
    If has_playerFilter Then opt_playerFilter = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_seen = unbox_has(idx, argc)
    If has_seen Then opt_seen = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    n = scripting_engine_enumAreaJS(context, area_lookup, has_playerFilter, opt_playerFilter, has_seen, opt_seen, result)
    js_enumArea = box_vector_BASE_OBJECT(result, n)
End Function

' quickjs_backend.cpp:3285: IMPL_JS_FUNC(addBeacon, wzapi::addBeacon)
Private Function js_addBeacon(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::addBeacon"
    Dim idx As Long, context As wzapi_execution_context, x As Long, y As Long, playerFilter As Long, has_message As Boolean
    Dim opt_message As String
    Set context = wzapi_execution_context_new(ctx)
    x = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    playerFilter = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_message = unbox_has(idx, argc)
    If has_message Then opt_message = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    js_addBeacon = box_bool(wzapi_addBeacon(context, x, y, playerFilter, has_message, opt_message))
End Function

' quickjs_backend.cpp:3297: IMPL_JS_FUNC(chat, wzapi::chat)
Private Function js_chat(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::chat"
    Dim idx As Long, context As wzapi_execution_context, playerFilter As Long, message As String
    Set context = wzapi_execution_context_new(ctx)
    playerFilter = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    message = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    js_chat = box_bool(wzapi_chat(context, playerFilter, message))
End Function

' quickjs_backend.cpp:3298: IMPL_JS_FUNC(quickChat, wzapi::quickChat)
Private Function js_quickChat(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::quickChat"
    Dim idx As Long, context As wzapi_execution_context, playerFilter As Long, messageEnum As Long
    Set context = wzapi_execution_context_new(ctx)
    playerFilter = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    messageEnum = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_quickChat = box_bool(wzapi_quickChat(context, playerFilter, messageEnum))
End Function

' quickjs_backend.cpp:3299: IMPL_JS_FUNC(getDroidPath, wzapi::getDroidPath)
Private Function js_getDroidPath(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::getDroidPath"
    Dim idx As Long, context As wzapi_execution_context, psDroid As Long, result() As scr_position, n As Long
    Set context = wzapi_execution_context_new(ctx)
    psDroid = unbox_DROID(idx, argc, argv, WRAPPED_FUNCTION)
    n = wzapi_getDroidPath(context, psDroid, result)
    js_getDroidPath = box_vector_scr_position(result, n)
End Function

' quickjs_backend.cpp:3300: IMPL_JS_FUNC(setAlliance, wzapi::setAlliance)
Private Function js_setAlliance(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setAlliance"
    Dim idx As Long, context As wzapi_execution_context, player1 As Long, player2 As Long, areAllies As Boolean
    Set context = wzapi_execution_context_new(ctx)
    player1 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    player2 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    areAllies = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    js_setAlliance = box_bool(wzapi_setAlliance(context, player1, player2, areAllies))
End Function

' quickjs_backend.cpp:3301: IMPL_JS_FUNC(sendAllianceRequest, wzapi::sendAllianceRequest)
Private Function js_sendAllianceRequest(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::sendAllianceRequest"
    Dim idx As Long, context As wzapi_execution_context, player2 As Long
    Set context = wzapi_execution_context_new(ctx)
    player2 = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_sendAllianceRequest context, player2
    js_sendAllianceRequest = box_no_return_value()
End Function

' quickjs_backend.cpp:3302: IMPL_JS_FUNC(setAssemblyPoint, wzapi::setAssemblyPoint)
Private Function js_setAssemblyPoint(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setAssemblyPoint"
    Dim idx As Long, context As wzapi_execution_context, psStruct As Long, x As Long, y As Long
    Set context = wzapi_execution_context_new(ctx)
    psStruct = unbox_STRUCTURE(idx, argc, argv, WRAPPED_FUNCTION)
    x = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_setAssemblyPoint = box_bool(wzapi_setAssemblyPoint(context, psStruct, x, y))
End Function

' quickjs_backend.cpp:3303: IMPL_JS_FUNC(hackNetOff, wzapi::hackNetOff)
Private Function js_hackNetOff(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::hackNetOff"
    Dim idx As Long, context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    wzapi_hackNetOff context
    js_hackNetOff = box_no_return_value()
End Function

' quickjs_backend.cpp:3304: IMPL_JS_FUNC(hackNetOn, wzapi::hackNetOn)
Private Function js_hackNetOn(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::hackNetOn"
    Dim idx As Long, context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    wzapi_hackNetOn context
    js_hackNetOn = box_no_return_value()
End Function

' quickjs_backend.cpp:3305: IMPL_JS_FUNC(getDroidProduction, wzapi::getDroidProduction)
Private Function js_getDroidProduction(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::getDroidProduction"
    Dim idx As Long, context As wzapi_execution_context, psFactory As Long, result As Long
    Set context = wzapi_execution_context_new(ctx)
    psFactory = unbox_STRUCTURE(idx, argc, argv, WRAPPED_FUNCTION)
    result = wzapi_getDroidProduction(context, psFactory)
    js_getDroidProduction = box_DROID(result)  ' box(std::unique_ptr<const DROID>): box(result.get())
    If result <> 0 Then DroidContainer_erase result ' the unique_ptr's delete (~DROID)
End Function

' quickjs_backend.cpp:3306: IMPL_JS_FUNC(getDroidLimit, wzapi::getDroidLimit)
Private Function js_getDroidLimit(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::getDroidLimit"
    Dim idx As Long, context As wzapi_execution_context, has_player As Boolean, opt_player As Long, has_droidType As Boolean, opt_droidType As Long
    Set context = wzapi_execution_context_new(ctx)
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_droidType = unbox_has(idx, argc)
    If has_droidType Then opt_droidType = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_getDroidLimit = box_int(wzapi_getDroidLimit(context, has_player, opt_player, has_droidType, opt_droidType))
End Function

' quickjs_backend.cpp:3307: IMPL_JS_FUNC(getExperienceModifier, wzapi::getExperienceModifier)
Private Function js_getExperienceModifier(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::getExperienceModifier"
    Dim idx As Long, context As wzapi_execution_context, player As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_getExperienceModifier = box_int(wzapi_getExperienceModifier(context, player))
End Function

' quickjs_backend.cpp:3308: IMPL_JS_FUNC(setExperienceModifier, wzapi::setExperienceModifier)
Private Function js_setExperienceModifier(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setExperienceModifier"
    Dim idx As Long, context As wzapi_execution_context, player As Long, percent As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    percent = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_setExperienceModifier = box_bool(wzapi_setExperienceModifier(context, player, percent))
End Function

' quickjs_backend.cpp:3309: IMPL_JS_FUNC(setDroidLimit, wzapi::setDroidLimit)
Private Function js_setDroidLimit(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setDroidLimit"
    Dim idx As Long, context As wzapi_execution_context, player As Long, maxNumber As Long, has_droidType As Boolean, opt_droidType As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    maxNumber = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_droidType = unbox_has(idx, argc)
    If has_droidType Then opt_droidType = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_setDroidLimit = box_bool(wzapi_setDroidLimit(context, player, maxNumber, has_droidType, opt_droidType))
End Function

' quickjs_backend.cpp:3310: IMPL_JS_FUNC(setCommanderLimit, wzapi::setCommanderLimit)
Private Function js_setCommanderLimit(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setCommanderLimit"
    Dim idx As Long, context As wzapi_execution_context, player As Long, maxNumber As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    maxNumber = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_setCommanderLimit = box_bool(wzapi_setCommanderLimit(context, player, maxNumber))
End Function

' quickjs_backend.cpp:3311: IMPL_JS_FUNC(setConstructorLimit, wzapi::setConstructorLimit)
Private Function js_setConstructorLimit(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setConstructorLimit"
    Dim idx As Long, context As wzapi_execution_context, player As Long, maxNumber As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    maxNumber = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_setConstructorLimit = box_bool(wzapi_setConstructorLimit(context, player, maxNumber))
End Function

' quickjs_backend.cpp:3312: IMPL_JS_FUNC_DEBUGMSGUPDATE(hackAddMessage, wzapi::hackAddMessage)
Private Function js_hackAddMessage(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::hackAddMessage"
    Dim idx As Long, context As wzapi_execution_context, message As String, messageType As Long, player As Long, immediate As Boolean
    Set context = wzapi_execution_context_new(ctx)
    message = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    messageType = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    immediate = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_hackAddMessage context, message, messageType, player, immediate
    js_hackAddMessage = box_no_return_value()
    jsDebugMessageUpdate
End Function

' quickjs_backend.cpp:3313: IMPL_JS_FUNC_DEBUGMSGUPDATE(hackRemoveMessage, wzapi::hackRemoveMessage)
Private Function js_hackRemoveMessage(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::hackRemoveMessage"
    Dim idx As Long, context As wzapi_execution_context, message As String, messageType As Long, player As Long
    Set context = wzapi_execution_context_new(ctx)
    message = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    messageType = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_hackRemoveMessage context, message, messageType, player
    js_hackRemoveMessage = box_no_return_value()
    jsDebugMessageUpdate
End Function

' quickjs_backend.cpp:3314: IMPL_JS_FUNC(setSunPosition, wzapi::setSunPosition)
Private Function js_setSunPosition(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setSunPosition"
    Dim idx As Long, context As wzapi_execution_context, x As Single, y As Single, z As Single
    Set context = wzapi_execution_context_new(ctx)
    x = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    z = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    js_setSunPosition = box_bool(wzapi_setSunPosition(context, x, y, z))
End Function

' quickjs_backend.cpp:3315: IMPL_JS_FUNC(setSunIntensity, wzapi::setSunIntensity)
Private Function js_setSunIntensity(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setSunIntensity"
    Dim idx As Long, context As wzapi_execution_context, ambient_r As Single, ambient_g As Single, ambient_b As Single, diffuse_r As Single
    Dim diffuse_g As Single, diffuse_b As Single, specular_r As Single, specular_g As Single, specular_b As Single
    Set context = wzapi_execution_context_new(ctx)
    ambient_r = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    ambient_g = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    ambient_b = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    diffuse_r = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    diffuse_g = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    diffuse_b = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    specular_r = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    specular_g = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    specular_b = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    js_setSunIntensity = box_bool(wzapi_setSunIntensity(context, ambient_r, ambient_g, ambient_b, diffuse_r, diffuse_g, diffuse_b, specular_r, specular_g, specular_b))
End Function

' quickjs_backend.cpp:3316: IMPL_JS_FUNC(setFogColour, wzapi::setFogColour)
Private Function js_setFogColour(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setFogColour"
    Dim idx As Long, context As wzapi_execution_context, r As Long, g As Long, b As Long
    Set context = wzapi_execution_context_new(ctx)
    r = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    g = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    b = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_setFogColour = box_bool(wzapi_setFogColour(context, r, g, b))
End Function

' quickjs_backend.cpp:3317: IMPL_JS_FUNC(setWeather, wzapi::setWeather)
Private Function js_setWeather(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setWeather"
    Dim idx As Long, context As wzapi_execution_context, weatherType As Long
    Set context = wzapi_execution_context_new(ctx)
    weatherType = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_setWeather = box_bool(wzapi_setWeather(context, weatherType))
End Function

' quickjs_backend.cpp:3318: IMPL_JS_FUNC(setSky, wzapi::setSky)
Private Function js_setSky(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setSky"
    Dim idx As Long, context As wzapi_execution_context, textureFilename As String, windSpeed As Single, scale_ As Single
    Set context = wzapi_execution_context_new(ctx)
    textureFilename = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    windSpeed = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    scale_ = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    js_setSky = box_bool(wzapi_setSky(context, textureFilename, windSpeed, scale_))
End Function

' quickjs_backend.cpp:3319: IMPL_JS_FUNC(hackDoNotSave, wzapi::hackDoNotSave)
Private Function js_hackDoNotSave(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::hackDoNotSave"
    Dim idx As Long, context As wzapi_execution_context, name_ As String
    Set context = wzapi_execution_context_new(ctx)
    name_ = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_hackDoNotSave context, name_
    js_hackDoNotSave = box_no_return_value()
End Function

' quickjs_backend.cpp:3320: IMPL_JS_FUNC(hackMarkTiles, wzapi::hackMarkTiles)
Private Function js_hackMarkTiles(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::hackMarkTiles"
    Dim idx As Long, context As wzapi_execution_context, has_tilePosOrArea As Boolean, opt_tilePosOrArea As label_or_position_values
    Set context = wzapi_execution_context_new(ctx)
    has_tilePosOrArea = unbox_has(idx, argc)
    If has_tilePosOrArea Then opt_tilePosOrArea = unbox_label_or_position_values(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_hackMarkTiles context, has_tilePosOrArea, opt_tilePosOrArea
    js_hackMarkTiles = box_no_return_value()
End Function

' quickjs_backend.cpp:3321: IMPL_JS_FUNC(cameraSlide, wzapi::cameraSlide)
Private Function js_cameraSlide(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::cameraSlide"
    Dim idx As Long, context As wzapi_execution_context, x As Single, y As Single
    Set context = wzapi_execution_context_new(ctx)
    x = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    js_cameraSlide = box_bool(wzapi_cameraSlide(context, x, y))
End Function

' quickjs_backend.cpp:3322: IMPL_JS_FUNC(cameraZoom, wzapi::cameraZoom)
Private Function js_cameraZoom(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::cameraZoom"
    Dim idx As Long, context As wzapi_execution_context, viewDistance As Single, speed As Single
    Set context = wzapi_execution_context_new(ctx)
    viewDistance = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    speed = unbox_float(idx, argc, argv, WRAPPED_FUNCTION)
    js_cameraZoom = box_bool(wzapi_cameraZoom(context, viewDistance, speed))
End Function

' quickjs_backend.cpp:3323: IMPL_JS_FUNC(cameraTrack, wzapi::cameraTrack)
Private Function js_cameraTrack(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::cameraTrack"
    Dim idx As Long, context As wzapi_execution_context, has_droid As Boolean, opt_droid As Long
    Set context = wzapi_execution_context_new(ctx)
    has_droid = unbox_has(idx, argc)
    If has_droid Then opt_droid = unbox_DROID(idx, argc, argv, WRAPPED_FUNCTION)
    js_cameraTrack = box_bool(wzapi_cameraTrack(context, has_droid, opt_droid))
End Function

' quickjs_backend.cpp:3324: IMPL_JS_FUNC(setHealth, wzapi::setHealth)
Private Function js_setHealth(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setHealth"
    Dim idx As Long, context As wzapi_execution_context, psObject As Long, health As Long
    Set context = wzapi_execution_context_new(ctx)
    psObject = unbox_BASE_OBJECT(idx, argc, argv, WRAPPED_FUNCTION)
    health = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_setHealth = box_bool(wzapi_setHealth(context, psObject, health))
End Function

' quickjs_backend.cpp:3325: IMPL_JS_FUNC(setObjectFlag, wzapi::setObjectFlag)
Private Function js_setObjectFlag(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setObjectFlag"
    Dim idx As Long, context As wzapi_execution_context, psObj As Long, flag_ As Long, flagValue As Boolean
    Set context = wzapi_execution_context_new(ctx)
    psObj = unbox_BASE_OBJECT(idx, argc, argv, WRAPPED_FUNCTION)
    flag_ = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    flagValue = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_setObjectFlag context, psObj, flag_, flagValue
    js_setObjectFlag = box_no_return_value()
End Function

' quickjs_backend.cpp:3326: IMPL_JS_FUNC(addSpotter, wzapi::addSpotter)
Private Function js_addSpotter(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::addSpotter"
    Dim idx As Long, context As wzapi_execution_context, x As Long, y As Long, player As Long, range_ As Long
    Dim radar As Boolean, expiry As Long
    Set context = wzapi_execution_context_new(ctx)
    x = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    range_ = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    radar = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    expiry = unbox_uint(idx, argc, argv, WRAPPED_FUNCTION)
    js_addSpotter = box_uint(wzapi_addSpotter(context, x, y, player, range_, radar, expiry))
End Function

' quickjs_backend.cpp:3327: IMPL_JS_FUNC(removeSpotter, wzapi::removeSpotter)
Private Function js_removeSpotter(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::removeSpotter"
    Dim idx As Long, context As wzapi_execution_context, spotterId As Long
    Set context = wzapi_execution_context_new(ctx)
    spotterId = unbox_uint(idx, argc, argv, WRAPPED_FUNCTION)
    js_removeSpotter = box_bool(wzapi_removeSpotter(context, spotterId))
End Function

' quickjs_backend.cpp:3328: IMPL_JS_FUNC(syncRandom, wzapi::syncRandom)
Private Function js_syncRandom(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::syncRandom"
    Dim idx As Long, context As wzapi_execution_context, limit As Long
    Set context = wzapi_execution_context_new(ctx)
    limit = unbox_uint(idx, argc, argv, WRAPPED_FUNCTION)
    js_syncRandom = box_uint(wzapi_syncRandom(context, limit))
End Function

' quickjs_backend.cpp:3329: IMPL_JS_FUNC(syncRequest, wzapi::syncRequest)
Private Function js_syncRequest(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::syncRequest"
    Dim idx As Long, context As wzapi_execution_context, req_id As Long, x As Long, y As Long, has_psObj As Boolean
    Dim opt_psObj As Long, has_psObj2 As Boolean, opt_psObj2 As Long
    Set context = wzapi_execution_context_new(ctx)
    req_id = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    x = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_psObj = unbox_has(idx, argc)
    If has_psObj Then opt_psObj = unbox_BASE_OBJECT(idx, argc, argv, WRAPPED_FUNCTION)
    has_psObj2 = unbox_has(idx, argc)
    If has_psObj2 Then opt_psObj2 = unbox_BASE_OBJECT(idx, argc, argv, WRAPPED_FUNCTION)
    js_syncRequest = box_bool(wzapi_syncRequest(context, req_id, x, y, has_psObj, opt_psObj, has_psObj2, opt_psObj2))
End Function

' quickjs_backend.cpp:3330: IMPL_JS_FUNC(replaceTexture, wzapi::replaceTexture)
Private Function js_replaceTexture(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::replaceTexture"
    Dim idx As Long, context As wzapi_execution_context, oldFilename As String, newFilename As String
    Set context = wzapi_execution_context_new(ctx)
    oldFilename = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    newFilename = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    js_replaceTexture = box_bool(wzapi_replaceTexture(context, oldFilename, newFilename))
End Function

' quickjs_backend.cpp:3331: IMPL_JS_FUNC(fireWeaponAtLoc, wzapi::fireWeaponAtLoc)
Private Function js_fireWeaponAtLoc(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::fireWeaponAtLoc"
    Dim idx As Long, context As wzapi_execution_context, weaponName As String, x As Long, y As Long, has_player As Boolean
    Dim opt_player As Long, has_center As Boolean, opt_center As Boolean
    Set context = wzapi_execution_context_new(ctx)
    weaponName = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    x = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    y = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    has_center = unbox_has(idx, argc)
    If has_center Then opt_center = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_fireWeaponAtLoc context, weaponName, x, y, has_player, opt_player, has_center, opt_center
    js_fireWeaponAtLoc = box_no_return_value()
End Function

' quickjs_backend.cpp:3332: IMPL_JS_FUNC(fireWeaponAtObj, wzapi::fireWeaponAtObj)
Private Function js_fireWeaponAtObj(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::fireWeaponAtObj"
    Dim idx As Long, context As wzapi_execution_context, weaponName As String, psObj As Long, has_player As Boolean, opt_player As Long
    Set context = wzapi_execution_context_new(ctx)
    weaponName = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    psObj = unbox_BASE_OBJECT(idx, argc, argv, WRAPPED_FUNCTION)
    has_player = unbox_has(idx, argc)
    If has_player Then opt_player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    wzapi_fireWeaponAtObj context, weaponName, psObj, has_player, opt_player
    js_fireWeaponAtObj = box_no_return_value()
End Function

' quickjs_backend.cpp:3333: IMPL_JS_FUNC(transformPlayerToSpectator, wzapi::transformPlayerToSpectator)
Private Function js_transformPlayerToSpectator(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::transformPlayerToSpectator"
    Dim idx As Long, context As wzapi_execution_context, player As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_transformPlayerToSpectator = box_bool(wzapi_transformPlayerToSpectator(context, player))
End Function

' quickjs_backend.cpp:3334: IMPL_JS_FUNC(isSpectator, wzapi::isSpectator)
Private Function js_isSpectator(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::isSpectator"
    Dim idx As Long, context As wzapi_execution_context, player As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_isSpectator = box_bool(wzapi_isSpectator(context, player))
End Function

' quickjs_backend.cpp:3335: IMPL_JS_FUNC(changePlayerColour, wzapi::changePlayerColour)
Private Function js_changePlayerColour(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::changePlayerColour"
    Dim idx As Long, context As wzapi_execution_context, player As Long, colour As Long
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    colour = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_changePlayerColour = box_bool(wzapi_changePlayerColour(context, player, colour))
End Function

' quickjs_backend.cpp:3336: IMPL_JS_FUNC(getMultiTechLevel, wzapi::getMultiTechLevel)
Private Function js_getMultiTechLevel(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::getMultiTechLevel"
    Dim idx As Long, context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    js_getMultiTechLevel = box_uint(wzapi_getMultiTechLevel(context))
End Function

' quickjs_backend.cpp:3337: IMPL_JS_FUNC(setCampaignNumber, wzapi::setCampaignNumber)
Private Function js_setCampaignNumber(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setCampaignNumber"
    Dim idx As Long, context As wzapi_execution_context, campaignNumber As Long
    Set context = wzapi_execution_context_new(ctx)
    campaignNumber = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    js_setCampaignNumber = box_bool(wzapi_setCampaignNumber(context, campaignNumber))
End Function

' quickjs_backend.cpp:3338: IMPL_JS_FUNC(getMissionType, wzapi::getMissionType)
Private Function js_getMissionType(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::getMissionType"
    Dim idx As Long, context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    js_getMissionType = box_int(wzapi_getMissionType(context))
End Function

' quickjs_backend.cpp:3339: IMPL_JS_FUNC(getRevealStatus, wzapi::getRevealStatus)
Private Function js_getRevealStatus(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::getRevealStatus"
    Dim idx As Long, context As wzapi_execution_context
    Set context = wzapi_execution_context_new(ctx)
    js_getRevealStatus = box_bool(wzapi_getRevealStatus(context))
End Function

' quickjs_backend.cpp:3340: IMPL_JS_FUNC(setRevealStatus, wzapi::setRevealStatus)
Private Function js_setRevealStatus(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setRevealStatus"
    Dim idx As Long, context As wzapi_execution_context, status As Boolean
    Set context = wzapi_execution_context_new(ctx)
    status = unbox_bool(idx, argc, argv, WRAPPED_FUNCTION)
    js_setRevealStatus = box_bool(wzapi_setRevealStatus(context, status))
End Function

' quickjs_backend.cpp:3341: IMPL_JS_FUNC(setGameStoryLogPlayerDataValue, wzapi::setGameStoryLogPlayerDataValue)
Private Function js_setGameStoryLogPlayerDataValue(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Const WRAPPED_FUNCTION As String = "wzapi::setGameStoryLogPlayerDataValue"
    Dim idx As Long, context As wzapi_execution_context, player As Long, key_str As String, value_str As String
    Set context = wzapi_execution_context_new(ctx)
    player = unbox_int(idx, argc, argv, WRAPPED_FUNCTION)
    key_str = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    value_str = unbox_string(idx, argc, argv, WRAPPED_FUNCTION)
    js_setGameStoryLogPlayerDataValue = box_bool(wzapi_setGameStoryLogPlayerDataValue(context, player, key_str, value_str))
End Function

'-- ## removeReticuleButton(buttonId)
'--
'-- Remove reticule button. DO NOT USE FOR ANYTHING.
'--
Private Function js_removeReticuleButton(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    js_removeReticuleButton = jsUndefinedValue()
End Function

' js_removeBeacon: wrap_(wzapi::removeBeacon, ...), then jsDebugMessageUpdate if it returned true
Private Function js_removeBeacon(ByVal ctx As Long, ByVal argc As Long, argv() As Long) As Long
    Dim idx As Long, context As wzapi_execution_context, p_playerFilter As Long, retVal As Long
    Set context = wzapi_execution_context_new(ctx)
    p_playerFilter = unbox_int(idx, argc, argv, "wzapi::removeBeacon")
    retVal = box_bool(wzapi_removeBeacon(context, p_playerFilter))
    If jsIsBool(retVal) Then
        If jsToBool(retVal) Then
            jsDebugMessageUpdate
        End If
    End If
    js_removeBeacon = retVal
End Function

'--------------------------------------------------------------------------
' The Upgrades object

' (VB) the entry of statsFuncs for a getter / setter of (name, player, type, index), made the first time
Private Function statsFuncEntry(ByVal name_ As String, ByVal player As Long, ByVal typ As Long, ByVal index As Long) As Long
    Dim key As String, k As Long
    If statsFuncKeys Is Nothing Then
        Set statsFuncKeys = CreateObject("Scripting.Dictionary")
        statsFuncKeys.CompareMode = vbBinaryCompare
    End If
    key = player & "," & typ & "," & index & "," & name_
    If statsFuncKeys.Exists(key) Then
        statsFuncEntry = statsFuncKeys(key)
        Exit Function
    End If
    k = statsFunc_size
    If k = 0 Then
        ReDim statsFunc_player(0 To 4095)
        ReDim statsFunc_type(0 To 4095)
        ReDim statsFunc_index(0 To 4095)
        ReDim statsFunc_name(0 To 4095)
    ElseIf k > UBound(statsFunc_player) Then
        ReDim Preserve statsFunc_player(0 To 2 * k - 1)
        ReDim Preserve statsFunc_type(0 To 2 * k - 1)
        ReDim Preserve statsFunc_index(0 To 2 * k - 1)
        ReDim Preserve statsFunc_name(0 To 2 * k - 1)
    End If
    statsFunc_player(k) = player
    statsFunc_type(k) = typ
    statsFunc_index(k) = index
    statsFunc_name(k) = name_
    statsFunc_size = k + 1
    statsFuncKeys.Add key, k
    statsFuncEntry = k
End Function

' static JSValue js_stats_get(JSContext *ctx, JSValueConst this_val) (entry: the function's statsFuncs entry)
Private Function js_stats_get(ByVal ctx As Long, ByVal this_val As Long, ByVal entry As Long) As Long
    Dim typ As Long, player As Long, index As Long, name_ As String
    ' (js_debugger_get_current_funcObject's "type", "player", "index", "name")
    typ = statsFunc_type(entry)
    player = statsFunc_player(entry)
    index = statsFunc_index(entry)
    name_ = statsFunc_name(entry)
    ' (quickjs_execution_context execution_context(ctx): wzapi_getUpgradeStats takes none)
    js_stats_get = mapJsonToQuickJSValue(wzapi_getUpgradeStats(player, name_, typ, index), JS_PROP_C_W_E)
End Function

' static JSValue js_stats_set(JSContext *ctx, JSValueConst this_val, JSValueConst val)
Private Function js_stats_set(ByVal ctx As Long, ByVal this_val As Long, ByVal val As Long, ByVal entry As Long) As Long
    Dim typ As Long, player As Long, index As Long, name_ As String, newValue As nlohmann_json, what As String, e As Long
    typ = statsFunc_type(entry)
    player = statsFunc_player(entry)
    index = statsFunc_index(entry)
    name_ = statsFunc_name(entry)
    ' (quickjs_execution_context execution_context(ctx): wzapi_setUpgradeStats takes none)
    ' auto toJsonContext = JSToJsonContext(ctx, true);
    If Not wz_qjs_to_json_try(val, True, newValue, what, 3366, "js_stats_set") Then
        ' catch (const std::runtime_error& e)
        debug_ 3368, LOG_ERROR, "js_stats_set", "Failed to convert input value to json with error: %s", what
        ' JS_ThrowTypeError
        jsClearException
        JsCreateTypeError jsStringValue("Failed to convert input value to json with error: " & what), e
        JsSetException e
        js_stats_set = 0
        Exit Function
    End If
    wzapi_setUpgradeStats player, name_, typ, index, newValue
    ' Now read value and return it
    js_stats_set = mapJsonToQuickJSValue(wzapi_getUpgradeStats(player, name_, typ, index), JS_PROP_C_W_E)
End Function

' static bool setStatsFunc(JSValue &base, JSContext *ctx, const std::string& name, int player, int type, unsigned index)
' (desc: the accessor descriptor constructUpgradesQuickJSValue made: JS_DefinePropertyGetSet's
' JS_PROP_CONFIGURABLE (QJS_CGETSET_DEF's prop_flags) | JS_PROP_WRITABLE | JS_PROP_ENUMERABLE, of which an
' accessor keeps configurable and enumerable)
Private Function setStatsFunc(ByVal base As Long, ByVal name_ As String, ByVal player As Long, ByVal typ As Long, ByVal index As Long, ByVal desc As Long) As Boolean
    Dim getter As Long, setter As Long, atom As Long, entry As Long, r As Long

    atom = jsPropId(name_)
    If Not (atom <> 0) Then ASSERT_FAILURE "atom != JS_ATOM_NULL", "src/quickjs_backend.cpp:3388", "setStatsFunc", "Failed to create atom: %s", name_: setStatsFunc = False: Exit Function

    entry = statsFuncEntry(name_, player, typ, index)
    getter = jsNewCFunction(JSFN_STATS_FIRST + 2 * entry, 0)
    If getter = 0 Then
        setStatsFunc = False
        Exit Function
    End If

    setter = jsNewCFunction(JSFN_STATS_FIRST + 2 * entry + 1, 1)
    If setter = 0 Then
        setStatsFunc = False
        Exit Function
    End If

    ' (the funcObjects' "player", "type", "index", "name": statsFuncs' entry)

    ' JS_DefinePropertyGetSet(ctx, base, atom, getter, setter, js_stats_getter_setter_func[0].prop_flags | JS_PROP_WRITABLE | JS_PROP_ENUMERABLE);
    jsSetProp desc, "get", getter
    jsSetProp desc, "set", setter
    If JsDefineProperty(base, atom, desc, r) <> JsNoError Then jsClearException

    setStatsFunc = True
End Function

' (VB) the GameEntityRules' properties (name, type) by name, from the cache of its map
Private Function upgradePropertiesOf(rules As GameEntityRules, names() As String, types() As Long) As Long
    Dim m As Long, n As Long, nm() As String, ty() As Long
    m = rules.propertyNameToTypeMap
    If m < 0 Or m > UBound(upgradeProperties) Then
        upgradePropertiesOf = GameEntityRules_properties(rules, names, types)
        Exit Function
    End If
    If Not upgradeProperties(m).built Then
        n = GameEntityRules_properties(rules, nm, ty)
        upgradeProperties(m).size = n
        If n > 0 Then
            upgradeProperties(m).names = nm
            upgradeProperties(m).types = ty
        End If
        upgradeProperties(m).built = True
    End If
    n = upgradeProperties(m).size
    If n > 0 Then
        names = upgradeProperties(m).names
        types = upgradeProperties(m).types
    End If
    upgradePropertiesOf = n
End Function

' JSValue constructUpgradesQuickJSValue(JSContext *ctx) (the current context; 0 = JS_EXCEPTION)
Public Function constructUpgradesQuickJSValue() As Long
    Dim upgradesObj() As PerPlayerUpgrades, nPlayers As Long, i As Long, c As Long, e As Long, p As Long
    Dim upgrades As Long, node As Long, entityBase As Long, v As Long, desc As Long
    Dim gameEntityClassName As String, gameEntityName As String, entityRules As GameEntityRules
    Dim propNames() As String, propTypes() As Long, nProps As Long
    nPlayers = wzapi_getUpgradesObject(upgradesObj)

    upgrades = jsNewArray()
    ' (VB) the accessor descriptor of setStatsFunc's JS_DefinePropertyGetSet: configurable, enumerable
    desc = jsNewObject()
    jsSetProp desc, "enumerable", jsBool(True)
    jsSetProp desc, "configurable", jsBool(True)
    For i = 0 To nPlayers - 1
        node = jsNewObject()

        For c = 0 To upgradesObj(i).upgrades_size - 1
            gameEntityClassName = upgradesObj(i).upgrades_keys(c)
            entityBase = jsNewObject()
            For e = 0 To upgradesObj(i).upgrades_values(c).rules_size - 1
                gameEntityName = upgradesObj(i).upgrades_values(c).rules(e).first
                entityRules = upgradesObj(i).upgrades_values(c).rules(e).second ' gameEntityRules
                v = jsNewObject()
                nProps = upgradePropertiesOf(entityRules, propNames, propTypes)
                For p = 0 To nProps - 1
                    If Not setStatsFunc(v, propNames(p), entityRules.player, propTypes(p), entityRules.index, desc) Then
                        ' Presumably, OOM
                        constructUpgradesQuickJSValue = 0
                        Exit Function
                    End If
                Next
                jsDefineProp entityBase, gameEntityName, v, JS_PROP_ENUMERABLE
            Next
            jsDefineProp node, gameEntityClassName, entityBase, JS_PROP_ENUMERABLE
        Next

        ' Finally
        If Not (upgradesObj(i).player >= 0 And upgradesObj(i).player < MAX_PLAYERS) Then ASSERT_FAILURE "playerUpgrades.getPlayer() >= 0 && playerUpgrades.getPlayer() < 11", "src/quickjs_backend.cpp:3462", "constructUpgradesQuickJSValue", "Invalid player index"
'        QuickJS_DefinePropertyValue(ctx, upgrades, std::to_string(playerUpgrades.getPlayer()).c_str(), node, 0);
        jsDefineProp upgrades, CStr(upgradesObj(i).player), node, JS_PROP_ENUMERABLE ' JS_DefinePropertyValueUint32
    Next

    constructUpgradesQuickJSValue = upgrades
End Function

' JS_REGISTER_FUNC, JS_REGISTER_FUNC2, JS_REGISTER_FUNC_NAME, JS_REGISTER_FUNC_NAME2 (line_: the macro's use,
' GCC's __LINE__ in its ASSERT_OR_RETURNs): False where they return false
Private Function JS_REGISTER_FUNC_(ByVal global_obj As Long, ByVal js_func_name As String, ByVal fnNumber As Long, ByVal num_parameters As Long, ByVal line_ As Long) As Boolean
    Dim newFunc As Long
    newFunc = jsNewCFunction(fnNumber, num_parameters)
    If Not (newFunc <> 0) Then ASSERT_FAILURE "!JS_IsException(newFunc)", "src/quickjs_backend.cpp:" & line_, "registerFunctions", "Failure to create function: %s", js_func_name: JS_REGISTER_FUNC_ = False: Exit Function
    If Not jsSetProp(global_obj, js_func_name, newFunc) Then ASSERT_FAILURE "setRes == 1", "src/quickjs_backend.cpp:" & line_, "registerFunctions", "Failure to register function property: %s", js_func_name: JS_REGISTER_FUNC_ = False: Exit Function
    JS_REGISTER_FUNC_ = True
End Function

' bool quickjs_scripting_instance::registerFunctions(const std::string& scriptName) (the instance's context current)
Public Function chakra_registerFunctions(ByVal instance As ChakraScriptingInstance, ByVal scriptName As String) As Boolean
    Dim global_obj As Long, upgrades As Long
    debug_ 3506, LOG_WZ, "registerFunctions", "Loading functions for context %p, script %s", instance.context(), scriptName
    global_obj = instance.Get_Global_Obj()

    '== * ```Upgrades``` A special array containing per-player rules information for game entity types,
    '== which can be written to in order to implement upgrades and other dynamic rules changes. Each item in the
    '== array contains a subset of the sparse array of rules information in the ```Stats``` global.
    '== These values are defined:
    upgrades = constructUpgradesQuickJSValue()
    If Not (upgrades <> 0) Then ASSERT_FAILURE "!JS_IsException(upgrades)", "src/quickjs_backend.cpp:3513", "registerFunctions", "Failure to create upgrades object": chakra_registerFunctions = False: Exit Function
    jsDefineProp global_obj, "Upgrades", upgrades, JS_PROP_WRITABLE Or JS_PROP_ENUMERABLE

    ' Register functions to the script engine here
    If Not JS_REGISTER_FUNC_(global_obj, "_", JSFN_UNDERSCORE, 1, 3517) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "dump", JSFN_dump, 1, 3518) Then Exit Function
    If Not JS_REGISTER_FUNC_(global_obj, "syncRandom", JSFN_syncRandom, 1, 3519) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "label", JSFN_label, 1, 3520) Then Exit Function ' deprecated // scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "getObject", JSFN_getObject, 1, 3521) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "addLabel", JSFN_addLabel, 2, 3522) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "removeLabel", JSFN_removeLabel, 1, 3523) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "getLabel", JSFN_getLabel, 1, 3524) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "enumLabels", JSFN_enumLabels, 1, 3525) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "enumGateways", JSFN_enumGateways, 0, 3526) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumTemplates", JSFN_enumTemplates, 1, 3527) Then Exit Function
    If Not JS_REGISTER_FUNC_(global_obj, "makeTemplate", JSFN_makeTemplate, 6, 3528) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setAlliance", JSFN_setAlliance, 3, 3529) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "sendAllianceRequest", JSFN_sendAllianceRequest, 1, 3530) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setAssemblyPoint", JSFN_setAssemblyPoint, 3, 3531) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setSunPosition", JSFN_setSunPosition, 3, 3532) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setSunIntensity", JSFN_setSunIntensity, 9, 3533) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setFogColour", JSFN_setFogColour, 3, 3534) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setWeather", JSFN_setWeather, 1, 3535) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setSky", JSFN_setSky, 3, 3536) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "cameraSlide", JSFN_cameraSlide, 2, 3537) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "cameraTrack", JSFN_cameraTrack, 1, 3538) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "cameraZoom", JSFN_cameraZoom, 2, 3539) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "resetArea", JSFN_resetArea, 1, 3540) Then Exit Function ' deprecated // scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "resetLabel", JSFN_resetLabel, 1, 3541) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "addSpotter", JSFN_addSpotter, 6, 3542) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "removeSpotter", JSFN_removeSpotter, 1, 3543) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "syncRequest", JSFN_syncRequest, 3, 3544) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "replaceTexture", JSFN_replaceTexture, 2, 3545) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "changePlayerColour", JSFN_changePlayerColour, 2, 3546) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setHealth", JSFN_setHealth, 2, 3547) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "useSafetyTransport", JSFN_useSafetyTransport, 1, 3548) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "restoreLimboMissionData", JSFN_restoreLimboMissionData, 0, 3549) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getMultiTechLevel", JSFN_getMultiTechLevel, 0, 3550) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setCampaignNumber", JSFN_setCampaignNumber, 1, 3551) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getMissionType", JSFN_getMissionType, 0, 3552) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getRevealStatus", JSFN_getRevealStatus, 0, 3553) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setRevealStatus", JSFN_setRevealStatus, 1, 3554) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "autoSave", JSFN_autoSave, 0, 3555) Then Exit Function ' WZAPI

    ' horrible hacks follow -- do not rely on these being present!
    If Not JS_REGISTER_FUNC_(global_obj, "hackNetOff", JSFN_hackNetOff, 0, 3558) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hackNetOn", JSFN_hackNetOn, 0, 3559) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hackAddMessage", JSFN_hackAddMessage, 4, 3560) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hackRemoveMessage", JSFN_hackRemoveMessage, 3, 3561) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hackGetObj", JSFN_hackGetObj, 3, 3562) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hackAssert", JSFN_hackAssert, 2, 3563) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hackMarkTiles", JSFN_hackMarkTiles, 1, 3564) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "receiveAllEvents", JSFN_receiveAllEvents, 0, 3565) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hackDoNotSave", JSFN_hackDoNotSave, 1, 3566) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hackPlayIngameAudio", JSFN_hackPlayIngameAudio, 0, 3567) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hackStopIngameAudio", JSFN_hackStopIngameAudio, 0, 3568) Then Exit Function ' WZAPI

    ' General functions -- geared for use in AI scripts
    If Not JS_REGISTER_FUNC_(global_obj, "debug", JSFN_debug, 1, 3571) Then Exit Function
    If Not JS_REGISTER_FUNC_(global_obj, "console", JSFN_console, 1, 3572) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "clearConsole", JSFN_clearConsole, 0, 3573) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "structureIdle", JSFN_structureIdle, 1, 3574) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumStruct", JSFN_enumStruct, 0, 3575) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumStructOffWorld", JSFN_enumStructOffWorld, 0, 3576) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumDroid", JSFN_enumDroid, 0, 3577) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumGroup", JSFN_enumGroup, 1, 3578) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "enumFeature", JSFN_enumFeature, 1, 3579) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumBlips", JSFN_enumBlips, 1, 3580) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumSelected", JSFN_enumSelected, 0, 3581) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumResearch", JSFN_enumResearch, 0, 3582) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumRange", JSFN_enumRange, 3, 3583) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enumArea", JSFN_enumArea, 1, 3584) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "getResearch", JSFN_getResearch, 1, 3585) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "pursueResearch", JSFN_pursueResearch, 2, 3586) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "findResearch", JSFN_findResearch, 1, 3587) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "distBetweenTwoPoints", JSFN_distBetweenTwoPoints, 4, 3588) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "newGroup", JSFN_newGroup, 0, 3589) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "groupAddArea", JSFN_groupAddArea, 5, 3590) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "groupAddDroid", JSFN_groupAddDroid, 2, 3591) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "groupAdd", JSFN_groupAdd, 2, 3592) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "groupSize", JSFN_groupSize, 1, 3593) Then Exit Function ' scripting_engine
    If Not JS_REGISTER_FUNC_(global_obj, "orderDroidLoc", JSFN_orderDroidLoc, 4, 3594) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "playerPower", JSFN_playerPower, 1, 3595) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "queuedPower", JSFN_queuedPower, 1, 3596) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "isStructureAvailable", JSFN_isStructureAvailable, 1, 3597) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "pickStructLocation", JSFN_pickStructLocation, 4, 3598) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "structureCanFit", JSFN_structureCanFit, 3, 3599) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "droidCanReach", JSFN_droidCanReach, 3, 3600) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "propulsionCanReach", JSFN_propulsionCanReach, 5, 3601) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "terrainType", JSFN_terrainType, 2, 3602) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "tileIsBurning", JSFN_tileIsBurning, 2, 3603) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "orderDroidBuild", JSFN_orderDroidBuild, 5, 3604) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "orderDroidObj", JSFN_orderDroidObj, 3, 3605) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "orderDroid", JSFN_orderDroid, 2, 3606) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "buildDroid", JSFN_buildDroid, 7, 3607) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "addDroid", JSFN_addDroid, 9, 3608) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "addDroidToTransporter", JSFN_addDroidToTransporter, 2, 3609) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "addFeature", JSFN_addFeature, 3, 3610) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "componentAvailable", JSFN_componentAvailable, 1, 3611) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "isVTOL", JSFN_isVTOL, 1, 3612) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "safeDest", JSFN_safeDest, 3, 3613) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "activateStructure", JSFN_activateStructure, 1, 3614) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "chat", JSFN_chat, 2, 3615) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "quickChat", JSFN_quickChat, 2, 3616) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getDroidPath", JSFN_getDroidPath, 1, 3617) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "addBeacon", JSFN_addBeacon, 3, 3618) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "removeBeacon", JSFN_removeBeacon, 1, 3619) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getDroidProduction", JSFN_getDroidProduction, 1, 3620) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getDroidLimit", JSFN_getDroidLimit, 0, 3621) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getExperienceModifier", JSFN_getExperienceModifier, 1, 3622) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setDroidLimit", JSFN_setDroidLimit, 2, 3623) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setCommanderLimit", JSFN_setCommanderLimit, 2, 3624) Then Exit Function ' deprecated!!
    If Not JS_REGISTER_FUNC_(global_obj, "setConstructorLimit", JSFN_setConstructorLimit, 2, 3625) Then Exit Function ' deprecated!!
    If Not JS_REGISTER_FUNC_(global_obj, "setExperienceModifier", JSFN_setExperienceModifier, 2, 3626) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getWeaponInfo", JSFN_getWeaponInfo, 1, 3627) Then Exit Function ' WZAPI // deprecated!!
    If Not JS_REGISTER_FUNC_(global_obj, "enumCargo", JSFN_enumCargo, 1, 3628) Then Exit Function ' WZAPI

    ' Functions that operate on the current player only
    If Not JS_REGISTER_FUNC_(global_obj, "centreView", JSFN_centreView, 2, 3631) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "playSound", JSFN_playSound, 1, 3632) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "gameOverMessage", JSFN_gameOverMessage, 1, 3633) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "addGuideTopic", JSFN_addGuideTopic, 1, 3634) Then Exit Function ' WZAPI

    ' Global state manipulation -- not for use with skirmish AI (unless you want it to cheat, obviously)
    If Not JS_REGISTER_FUNC_(global_obj, "setStructureLimits", JSFN_setStructureLimits, 2, 3637) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "applyLimitSet", JSFN_applyLimitSet, 0, 3638) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "emitSound", JSFN_emitSound, 3, 3639) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setMissionTime", JSFN_setMissionTime, 1, 3640) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getMissionTime", JSFN_getMissionTime, 0, 3641) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setReinforcementTime", JSFN_setReinforcementTime, 1, 3642) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "completeResearch", JSFN_completeResearch, 1, 3643) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "completeAllResearch", JSFN_completeAllResearch, 0, 3644) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enableResearch", JSFN_enableResearch, 1, 3645) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setPower", JSFN_setPower, 1, 3646) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setPowerModifier", JSFN_setPowerModifier, 1, 3647) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setPowerStorageMaximum", JSFN_setPowerStorageMaximum, 1, 3648) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "extraPowerTime", JSFN_extraPowerTime, 1, 3649) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setTutorialMode", JSFN_setTutorialMode, 1, 3650) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setDesign", JSFN_setDesign, 1, 3651) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enableTemplate", JSFN_enableTemplate, 1, 3652) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "removeTemplate", JSFN_removeTemplate, 1, 3653) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setMiniMap", JSFN_setMiniMap, 1, 3654) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setReticuleButton", JSFN_setReticuleButton, 4, 3655) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setReticuleFlash", JSFN_setReticuleFlash, 2, 3656) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "showReticuleWidget", JSFN_showReticuleWidget, 1, 3657) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "showInterface", JSFN_showInterface, 0, 3658) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "hideInterface", JSFN_hideInterface, 0, 3659) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "addReticuleButton", JSFN_addReticuleButton, 1, 3660) Then Exit Function ' deprecated!!
    If Not JS_REGISTER_FUNC_(global_obj, "removeReticuleButton", JSFN_removeReticuleButton, 1, 3661) Then Exit Function ' deprecated!!
    If Not JS_REGISTER_FUNC_(global_obj, "enableStructure", JSFN_enableStructure, 1, 3662) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "makeComponentAvailable", JSFN_makeComponentAvailable, 2, 3663) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "enableComponent", JSFN_enableComponent, 2, 3664) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "allianceExistsBetween", JSFN_allianceExistsBetween, 2, 3665) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "removeStruct", JSFN_removeStruct, 1, 3666) Then Exit Function ' WZAPI // deprecated!!
    If Not JS_REGISTER_FUNC_(global_obj, "removeObject", JSFN_removeObject, 1, 3667) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setScrollParams", JSFN_setScrollParams, 4, 3668) Then Exit Function ' deprecated!!
    If Not JS_REGISTER_FUNC_(global_obj, "setScrollLimits", JSFN_setScrollLimits, 4, 3669) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getScrollLimits", JSFN_getScrollLimits, 0, 3670) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "addStructure", JSFN_addStructure, 4, 3671) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "getStructureLimit", JSFN_getStructureLimit, 1, 3672) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "countStruct", JSFN_countStruct, 1, 3673) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "countDroid", JSFN_countDroid, 0, 3674) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "loadLevel", JSFN_loadLevel, 1, 3675) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setDroidExperience", JSFN_setDroidExperience, 2, 3676) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "donateObject", JSFN_donateObject, 2, 3677) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "donatePower", JSFN_donatePower, 2, 3678) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setNoGoArea", JSFN_setNoGoArea, 5, 3679) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "startTransporterEntry", JSFN_startTransporterEntry, 3, 3680) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setTransporterExit", JSFN_setTransporterExit, 3, 3681) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setObjectFlag", JSFN_setObjectFlag, 3, 3682) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "fireWeaponAtLoc", JSFN_fireWeaponAtLoc, 3, 3683) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "fireWeaponAtObj", JSFN_fireWeaponAtObj, 2, 3684) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "transformPlayerToSpectator", JSFN_transformPlayerToSpectator, 1, 3685) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "isSpectator", JSFN_isSpectator, 1, 3686) Then Exit Function ' WZAPI
    If Not JS_REGISTER_FUNC_(global_obj, "setGameStoryLogPlayerDataValue", JSFN_setGameStoryLogPlayerDataValue, 3, 3687) Then Exit Function ' WZAPI

    chakra_registerFunctions = True
End Function

' The API functions' numbers (JSFN_API_FIRST on): the wrappers (thisVal: the call's this)
Public Function jsCallApiFunction(ByVal fnNumber As Long, ByVal ctx As Long, ByVal thisVal As Long, ByVal argc As Long, argv() As Long) As Long
    If fnNumber >= JSFN_STATS_FIRST Then
        ' (VB) a getter or setter of the Upgrades object (setStatsFunc)
        If ((fnNumber - JSFN_STATS_FIRST) And 1) = 0 Then
            jsCallApiFunction = js_stats_get(ctx, thisVal, (fnNumber - JSFN_STATS_FIRST) \ 2)
        Else
            jsCallApiFunction = js_stats_set(ctx, thisVal, argv(0), (fnNumber - JSFN_STATS_FIRST) \ 2)
        End If
        Exit Function
    End If
    Select Case fnNumber
    Case JSFN_UNDERSCORE: jsCallApiFunction = js_translate(ctx, argc, argv)
    Case JSFN_dump: jsCallApiFunction = js_dump(ctx, argc, argv)
    Case JSFN_syncRandom: jsCallApiFunction = js_syncRandom(ctx, argc, argv)
    Case JSFN_label: jsCallApiFunction = js_getObject(ctx, argc, argv)
    Case JSFN_getObject: jsCallApiFunction = js_getObject(ctx, argc, argv)
    Case JSFN_addLabel: jsCallApiFunction = js_addLabel(ctx, argc, argv)
    Case JSFN_removeLabel: jsCallApiFunction = js_removeLabel(ctx, argc, argv)
    Case JSFN_getLabel: jsCallApiFunction = js_getLabel(ctx, argc, argv)
    Case JSFN_enumLabels: jsCallApiFunction = js_enumLabels(ctx, argc, argv)
    Case JSFN_enumGateways: jsCallApiFunction = js_enumGateways(ctx, argc, argv)
    Case JSFN_enumTemplates: jsCallApiFunction = js_enumTemplates(ctx, argc, argv)
    Case JSFN_makeTemplate: jsCallApiFunction = js_makeTemplate(ctx, argc, argv)
    Case JSFN_setAlliance: jsCallApiFunction = js_setAlliance(ctx, argc, argv)
    Case JSFN_sendAllianceRequest: jsCallApiFunction = js_sendAllianceRequest(ctx, argc, argv)
    Case JSFN_setAssemblyPoint: jsCallApiFunction = js_setAssemblyPoint(ctx, argc, argv)
    Case JSFN_setSunPosition: jsCallApiFunction = js_setSunPosition(ctx, argc, argv)
    Case JSFN_setSunIntensity: jsCallApiFunction = js_setSunIntensity(ctx, argc, argv)
    Case JSFN_setFogColour: jsCallApiFunction = js_setFogColour(ctx, argc, argv)
    Case JSFN_setWeather: jsCallApiFunction = js_setWeather(ctx, argc, argv)
    Case JSFN_setSky: jsCallApiFunction = js_setSky(ctx, argc, argv)
    Case JSFN_cameraSlide: jsCallApiFunction = js_cameraSlide(ctx, argc, argv)
    Case JSFN_cameraTrack: jsCallApiFunction = js_cameraTrack(ctx, argc, argv)
    Case JSFN_cameraZoom: jsCallApiFunction = js_cameraZoom(ctx, argc, argv)
    Case JSFN_resetArea: jsCallApiFunction = js_resetLabel(ctx, argc, argv)
    Case JSFN_resetLabel: jsCallApiFunction = js_resetLabel(ctx, argc, argv)
    Case JSFN_addSpotter: jsCallApiFunction = js_addSpotter(ctx, argc, argv)
    Case JSFN_removeSpotter: jsCallApiFunction = js_removeSpotter(ctx, argc, argv)
    Case JSFN_syncRequest: jsCallApiFunction = js_syncRequest(ctx, argc, argv)
    Case JSFN_replaceTexture: jsCallApiFunction = js_replaceTexture(ctx, argc, argv)
    Case JSFN_changePlayerColour: jsCallApiFunction = js_changePlayerColour(ctx, argc, argv)
    Case JSFN_setHealth: jsCallApiFunction = js_setHealth(ctx, argc, argv)
    Case JSFN_useSafetyTransport: jsCallApiFunction = js_useSafetyTransport(ctx, argc, argv)
    Case JSFN_restoreLimboMissionData: jsCallApiFunction = js_restoreLimboMissionData(ctx, argc, argv)
    Case JSFN_getMultiTechLevel: jsCallApiFunction = js_getMultiTechLevel(ctx, argc, argv)
    Case JSFN_setCampaignNumber: jsCallApiFunction = js_setCampaignNumber(ctx, argc, argv)
    Case JSFN_getMissionType: jsCallApiFunction = js_getMissionType(ctx, argc, argv)
    Case JSFN_getRevealStatus: jsCallApiFunction = js_getRevealStatus(ctx, argc, argv)
    Case JSFN_setRevealStatus: jsCallApiFunction = js_setRevealStatus(ctx, argc, argv)
    Case JSFN_autoSave: jsCallApiFunction = js_autoSave(ctx, argc, argv)
    Case JSFN_hackNetOff: jsCallApiFunction = js_hackNetOff(ctx, argc, argv)
    Case JSFN_hackNetOn: jsCallApiFunction = js_hackNetOn(ctx, argc, argv)
    Case JSFN_hackAddMessage: jsCallApiFunction = js_hackAddMessage(ctx, argc, argv)
    Case JSFN_hackRemoveMessage: jsCallApiFunction = js_hackRemoveMessage(ctx, argc, argv)
    Case JSFN_hackGetObj: jsCallApiFunction = js_hackGetObj(ctx, argc, argv)
    Case JSFN_hackAssert: jsCallApiFunction = js_hackAssert(ctx, argc, argv)
    Case JSFN_hackMarkTiles: jsCallApiFunction = js_hackMarkTiles(ctx, argc, argv)
    Case JSFN_receiveAllEvents: jsCallApiFunction = js_receiveAllEvents(ctx, argc, argv)
    Case JSFN_hackDoNotSave: jsCallApiFunction = js_hackDoNotSave(ctx, argc, argv)
    Case JSFN_hackPlayIngameAudio: jsCallApiFunction = js_hackPlayIngameAudio(ctx, argc, argv)
    Case JSFN_hackStopIngameAudio: jsCallApiFunction = js_hackStopIngameAudio(ctx, argc, argv)
    Case JSFN_debug: jsCallApiFunction = js_debug(ctx, argc, argv)
    Case JSFN_console: jsCallApiFunction = js_console(ctx, argc, argv)
    Case JSFN_clearConsole: jsCallApiFunction = js_clearConsole(ctx, argc, argv)
    Case JSFN_structureIdle: jsCallApiFunction = js_structureIdle(ctx, argc, argv)
    Case JSFN_enumStruct: jsCallApiFunction = js_enumStruct(ctx, argc, argv)
    Case JSFN_enumStructOffWorld: jsCallApiFunction = js_enumStructOffWorld(ctx, argc, argv)
    Case JSFN_enumDroid: jsCallApiFunction = js_enumDroid(ctx, argc, argv)
    Case JSFN_enumGroup: jsCallApiFunction = js_enumGroup(ctx, argc, argv)
    Case JSFN_enumFeature: jsCallApiFunction = js_enumFeature(ctx, argc, argv)
    Case JSFN_enumBlips: jsCallApiFunction = js_enumBlips(ctx, argc, argv)
    Case JSFN_enumSelected: jsCallApiFunction = js_enumSelected(ctx, argc, argv)
    Case JSFN_enumResearch: jsCallApiFunction = js_enumResearch(ctx, argc, argv)
    Case JSFN_enumRange: jsCallApiFunction = js_enumRange(ctx, argc, argv)
    Case JSFN_enumArea: jsCallApiFunction = js_enumArea(ctx, argc, argv)
    Case JSFN_getResearch: jsCallApiFunction = js_getResearch(ctx, argc, argv)
    Case JSFN_pursueResearch: jsCallApiFunction = js_pursueResearch(ctx, argc, argv)
    Case JSFN_findResearch: jsCallApiFunction = js_findResearch(ctx, argc, argv)
    Case JSFN_distBetweenTwoPoints: jsCallApiFunction = js_distBetweenTwoPoints(ctx, argc, argv)
    Case JSFN_newGroup: jsCallApiFunction = js_newGroup(ctx, argc, argv)
    Case JSFN_groupAddArea: jsCallApiFunction = js_groupAddArea(ctx, argc, argv)
    Case JSFN_groupAddDroid: jsCallApiFunction = js_groupAddDroid(ctx, argc, argv)
    Case JSFN_groupAdd: jsCallApiFunction = js_groupAdd(ctx, argc, argv)
    Case JSFN_groupSize: jsCallApiFunction = js_groupSize(ctx, argc, argv)
    Case JSFN_orderDroidLoc: jsCallApiFunction = js_orderDroidLoc(ctx, argc, argv)
    Case JSFN_playerPower: jsCallApiFunction = js_playerPower(ctx, argc, argv)
    Case JSFN_queuedPower: jsCallApiFunction = js_queuedPower(ctx, argc, argv)
    Case JSFN_isStructureAvailable: jsCallApiFunction = js_isStructureAvailable(ctx, argc, argv)
    Case JSFN_pickStructLocation: jsCallApiFunction = js_pickStructLocation(ctx, argc, argv)
    Case JSFN_structureCanFit: jsCallApiFunction = js_structureCanFit(ctx, argc, argv)
    Case JSFN_droidCanReach: jsCallApiFunction = js_droidCanReach(ctx, argc, argv)
    Case JSFN_propulsionCanReach: jsCallApiFunction = js_propulsionCanReach(ctx, argc, argv)
    Case JSFN_terrainType: jsCallApiFunction = js_terrainType(ctx, argc, argv)
    Case JSFN_tileIsBurning: jsCallApiFunction = js_tileIsBurning(ctx, argc, argv)
    Case JSFN_orderDroidBuild: jsCallApiFunction = js_orderDroidBuild(ctx, argc, argv)
    Case JSFN_orderDroidObj: jsCallApiFunction = js_orderDroidObj(ctx, argc, argv)
    Case JSFN_orderDroid: jsCallApiFunction = js_orderDroid(ctx, argc, argv)
    Case JSFN_buildDroid: jsCallApiFunction = js_buildDroid(ctx, argc, argv)
    Case JSFN_addDroid: jsCallApiFunction = js_addDroid(ctx, argc, argv)
    Case JSFN_addDroidToTransporter: jsCallApiFunction = js_addDroidToTransporter(ctx, argc, argv)
    Case JSFN_addFeature: jsCallApiFunction = js_addFeature(ctx, argc, argv)
    Case JSFN_componentAvailable: jsCallApiFunction = js_componentAvailable(ctx, argc, argv)
    Case JSFN_isVTOL: jsCallApiFunction = js_isVTOL(ctx, argc, argv)
    Case JSFN_safeDest: jsCallApiFunction = js_safeDest(ctx, argc, argv)
    Case JSFN_activateStructure: jsCallApiFunction = js_activateStructure(ctx, argc, argv)
    Case JSFN_chat: jsCallApiFunction = js_chat(ctx, argc, argv)
    Case JSFN_quickChat: jsCallApiFunction = js_quickChat(ctx, argc, argv)
    Case JSFN_getDroidPath: jsCallApiFunction = js_getDroidPath(ctx, argc, argv)
    Case JSFN_addBeacon: jsCallApiFunction = js_addBeacon(ctx, argc, argv)
    Case JSFN_removeBeacon: jsCallApiFunction = js_removeBeacon(ctx, argc, argv)
    Case JSFN_getDroidProduction: jsCallApiFunction = js_getDroidProduction(ctx, argc, argv)
    Case JSFN_getDroidLimit: jsCallApiFunction = js_getDroidLimit(ctx, argc, argv)
    Case JSFN_getExperienceModifier: jsCallApiFunction = js_getExperienceModifier(ctx, argc, argv)
    Case JSFN_setDroidLimit: jsCallApiFunction = js_setDroidLimit(ctx, argc, argv)
    Case JSFN_setCommanderLimit: jsCallApiFunction = js_setCommanderLimit(ctx, argc, argv)
    Case JSFN_setConstructorLimit: jsCallApiFunction = js_setConstructorLimit(ctx, argc, argv)
    Case JSFN_setExperienceModifier: jsCallApiFunction = js_setExperienceModifier(ctx, argc, argv)
    Case JSFN_getWeaponInfo: jsCallApiFunction = js_getWeaponInfo(ctx, argc, argv)
    Case JSFN_enumCargo: jsCallApiFunction = js_enumCargo(ctx, argc, argv)
    Case JSFN_centreView: jsCallApiFunction = js_centreView(ctx, argc, argv)
    Case JSFN_playSound: jsCallApiFunction = js_playSound(ctx, argc, argv)
    Case JSFN_gameOverMessage: jsCallApiFunction = js_gameOverMessage(ctx, argc, argv)
    Case JSFN_addGuideTopic: jsCallApiFunction = js_addGuideTopic(ctx, argc, argv)
    Case JSFN_setStructureLimits: jsCallApiFunction = js_setStructureLimits(ctx, argc, argv)
    Case JSFN_applyLimitSet: jsCallApiFunction = js_applyLimitSet(ctx, argc, argv)
    Case JSFN_emitSound: jsCallApiFunction = js_emitSound(ctx, argc, argv)
    Case JSFN_setMissionTime: jsCallApiFunction = js_setMissionTime(ctx, argc, argv)
    Case JSFN_getMissionTime: jsCallApiFunction = js_getMissionTime(ctx, argc, argv)
    Case JSFN_setReinforcementTime: jsCallApiFunction = js_setReinforcementTime(ctx, argc, argv)
    Case JSFN_completeResearch: jsCallApiFunction = js_completeResearch(ctx, argc, argv)
    Case JSFN_completeAllResearch: jsCallApiFunction = js_completeAllResearch(ctx, argc, argv)
    Case JSFN_enableResearch: jsCallApiFunction = js_enableResearch(ctx, argc, argv)
    Case JSFN_setPower: jsCallApiFunction = js_setPower(ctx, argc, argv)
    Case JSFN_setPowerModifier: jsCallApiFunction = js_setPowerModifier(ctx, argc, argv)
    Case JSFN_setPowerStorageMaximum: jsCallApiFunction = js_setPowerStorageMaximum(ctx, argc, argv)
    Case JSFN_extraPowerTime: jsCallApiFunction = js_extraPowerTime(ctx, argc, argv)
    Case JSFN_setTutorialMode: jsCallApiFunction = js_setTutorialMode(ctx, argc, argv)
    Case JSFN_setDesign: jsCallApiFunction = js_setDesign(ctx, argc, argv)
    Case JSFN_enableTemplate: jsCallApiFunction = js_enableTemplate(ctx, argc, argv)
    Case JSFN_removeTemplate: jsCallApiFunction = js_removeTemplate(ctx, argc, argv)
    Case JSFN_setMiniMap: jsCallApiFunction = js_setMiniMap(ctx, argc, argv)
    Case JSFN_setReticuleButton: jsCallApiFunction = js_setReticuleButton(ctx, argc, argv)
    Case JSFN_setReticuleFlash: jsCallApiFunction = js_setReticuleFlash(ctx, argc, argv)
    Case JSFN_showReticuleWidget: jsCallApiFunction = js_showReticuleWidget(ctx, argc, argv)
    Case JSFN_showInterface: jsCallApiFunction = js_showInterface(ctx, argc, argv)
    Case JSFN_hideInterface: jsCallApiFunction = js_hideInterface(ctx, argc, argv)
    Case JSFN_addReticuleButton: jsCallApiFunction = js_removeReticuleButton(ctx, argc, argv)
    Case JSFN_removeReticuleButton: jsCallApiFunction = js_removeReticuleButton(ctx, argc, argv)
    Case JSFN_enableStructure: jsCallApiFunction = js_enableStructure(ctx, argc, argv)
    Case JSFN_makeComponentAvailable: jsCallApiFunction = js_makeComponentAvailable(ctx, argc, argv)
    Case JSFN_enableComponent: jsCallApiFunction = js_enableComponent(ctx, argc, argv)
    Case JSFN_allianceExistsBetween: jsCallApiFunction = js_allianceExistsBetween(ctx, argc, argv)
    Case JSFN_removeStruct: jsCallApiFunction = js_removeStruct(ctx, argc, argv)
    Case JSFN_removeObject: jsCallApiFunction = js_removeObject(ctx, argc, argv)
    Case JSFN_setScrollParams: jsCallApiFunction = js_setScrollLimits(ctx, argc, argv)
    Case JSFN_setScrollLimits: jsCallApiFunction = js_setScrollLimits(ctx, argc, argv)
    Case JSFN_getScrollLimits: jsCallApiFunction = js_getScrollLimits(ctx, argc, argv)
    Case JSFN_addStructure: jsCallApiFunction = js_addStructure(ctx, argc, argv)
    Case JSFN_getStructureLimit: jsCallApiFunction = js_getStructureLimit(ctx, argc, argv)
    Case JSFN_countStruct: jsCallApiFunction = js_countStruct(ctx, argc, argv)
    Case JSFN_countDroid: jsCallApiFunction = js_countDroid(ctx, argc, argv)
    Case JSFN_loadLevel: jsCallApiFunction = js_loadLevel(ctx, argc, argv)
    Case JSFN_setDroidExperience: jsCallApiFunction = js_setDroidExperience(ctx, argc, argv)
    Case JSFN_donateObject: jsCallApiFunction = js_donateObject(ctx, argc, argv)
    Case JSFN_donatePower: jsCallApiFunction = js_donatePower(ctx, argc, argv)
    Case JSFN_setNoGoArea: jsCallApiFunction = js_setNoGoArea(ctx, argc, argv)
    Case JSFN_startTransporterEntry: jsCallApiFunction = js_startTransporterEntry(ctx, argc, argv)
    Case JSFN_setTransporterExit: jsCallApiFunction = js_setTransporterExit(ctx, argc, argv)
    Case JSFN_setObjectFlag: jsCallApiFunction = js_setObjectFlag(ctx, argc, argv)
    Case JSFN_fireWeaponAtLoc: jsCallApiFunction = js_fireWeaponAtLoc(ctx, argc, argv)
    Case JSFN_fireWeaponAtObj: jsCallApiFunction = js_fireWeaponAtObj(ctx, argc, argv)
    Case JSFN_transformPlayerToSpectator: jsCallApiFunction = js_transformPlayerToSpectator(ctx, argc, argv)
    Case JSFN_isSpectator: jsCallApiFunction = js_isSpectator(ctx, argc, argv)
    Case JSFN_setGameStoryLogPlayerDataValue: jsCallApiFunction = js_setGameStoryLogPlayerDataValue(ctx, argc, argv)
    Case Else
        jsCallApiFunction = jsUndefinedValue()
    End Select
End Function
