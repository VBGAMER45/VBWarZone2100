Attribute VB_Name = "wzapi_c"
' src/wzapi.cpp, wzapi.h: the script API, wzapi.cpp whole (cluster C), with
' the wzapi.h types the script engine shares and wzapi.h's inline
' constructors and getters of game_object_identifier, object_request and
' label_or_position_values; IdToObject; the per-player Upgrades object
' (GameEntityRules, GameEntityRuleContainer, PerPlayerUpgrades,
' getUpgradesObject), its getter and setter (getUpgradeStats,
' setUpgradeStats), the read-only Stats object (register_common,
' constructStatsObject), getUsefulConstants, constructStaticPlayerData,
' constructMapTilesArray and the queue of objects removeObject removes
' (scriptQueuedObjectRemovals: the arrays queuedObjRemovals_obj and _sfx,
' std::vector<std::pair<BASE_OBJECT*, bool>>; callers that only test it use
' wzapi_scriptQueuedObjectRemovals_empty).
' wzapi::scripting_instance's and execution_context's members are the classes
' ChakraScriptingInstance and wzapi_execution_context.
' A wzapi:: function is wzapi_<name>, with the signature rule of CONVENTIONS
' ("The scripts"): the execution context first (none for
' WZAPI_NO_PARAMS_NO_CONTEXT), an optional<T> name as has_<name> and
' opt_<name>, a std::vector<T> result as the count returned plus a last
' result() array filled from 0, an optional<T> result as has_value returned
' plus a last ByRef result, string_or_string_list and
' va_list_treat_as_strings as StringVector, va_list<string_or_string_list> as
' an array of StringVector and its size, reservedParam as an ignored Long.
' The static helpers keep their names (_enumStruct_fromList is
' enumStruct_fromList_, the static makeTemplate is makeTemplate_, the LOC_OK
' macro a function).
' SCRIPT_ASSERT(retval, context, expr, ...) is the log line,
' context.throwError (which records the error and does not throw, as with the
' official build's QuickJS) and the return; SCRIPT_ASSERT_PLAYER its expansion.
' A std::unique_ptr<DROID_TEMPLATE> is a template handle of template_c's pool
' (deleteTemplate frees it, as the unique_ptr would); getDroidProduction's
' virtual DROID is a droid slot on no list (DroidContainer_erase frees it).
' gettext is not ported (translate returns its string). The lambdas of
' enumerateTemplates are loops over droidTemplates_at(player, i), after
' enumerateTemplates' player check.
' The classes are UDTs and their methods <Class>_<method>(this, ...).
' GameEntityRules' NameToTypeMap (std::map<std::string, int>) is the index
' of its map in nameToTypeMaps: the rules of a game entity class have equal
' maps, built once (find is a linear search; GameEntityRules_properties
' gives them in std::map's order).
' PerPlayerUpgrades' std::map<GameEntityClass, GameEntityRuleContainer> is
' two arrays in the keys' order; GameEntityRuleContainer's lookup_table
' (std::unordered_map) is a Scripting.Dictionary.
' wzapi_setUpgradeStats and wzapi_getUpgradeStats take no execution context:
' the only caller, research.cpp, passes an internal_execution_context_base
' whose throwError does nothing, so SCRIPT_ASSERT is the log line and the
' return.
' Not ported: the operator[]s of GameEntityRuleContainer and
' PerPlayerUpgrades and the iterators (the scripts use them).
' A nlohmann::json is WzFramework's nlohmann_json; a field's C++ type picks
' the JSON type it becomes (an unsigned one json_value_unsigned, an int or
' short json_value of a Long, a bool or string json_value).
Option Explicit


' (wzapi.cpp's enum Scrcb)
Public Enum Scrcb
    SCRCB_FIRST = COMP_NUMCOMPONENTS
    SCRCB_RES = SCRCB_FIRST     ' Research upgrade
    SCRCB_MODULE_RES            ' Research module upgrade
    SCRCB_REP                   ' Repair upgrade
    SCRCB_POW                   ' Power upgrade
    SCRCB_MODULE_POW            ' And so on...
    SCRCB_CON
    SCRCB_MODULE_CON
    SCRCB_REA
    SCRCB_ARM
    SCRCB_HEA
    SCRCB_ELW
    SCRCB_HIT
    SCRCB_LIMIT
    SCRCB_LAST = SCRCB_LIMIT
End Enum

' class GameEntityRules
Public Type GameEntityRules
    player As Long                      ' int
    index As Long                       ' unsigned
    propertyNameToTypeMap As Long       ' (VB) NameToTypeMap: its index in nameToTypeMaps
End Type

' GameEntityRuleContainer::GameEntityRulesPair (std::pair<GameEntityName, GameEntityRules>)
Public Type GameEntityRulesPair
    first As String
    second As GameEntityRules
End Type

' class GameEntityRuleContainer
Public Type GameEntityRuleContainer
    rules() As GameEntityRulesPair      ' std::vector<GameEntityRulesPair>
    rules_size As Long
    lookup_table As Object              ' std::unordered_map<GameEntityName, size_t>: Scripting.Dictionary
End Type

' class PerPlayerUpgrades
Public Type PerPlayerUpgrades
    upgrades_keys() As String           ' std::map<GameEntityClass, GameEntityRuleContainer>: the keys, ascending
    upgrades_values() As GameEntityRuleContainer
    upgrades_size As Long
    player As Long
End Type

' wzapi.h: enum timerType
Public Enum timerType
    TIMER_REPEAT
    TIMER_ONESHOT_READY
    TIMER_ONESHOT_DONE
    TIMER_REMOVED
End Enum

' wzapi.h: struct scr_radius, scr_area, scr_position
Public Type scr_radius
    x As Long
    y As Long
    radius As Long
End Type

Public Type scr_area
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Public Type scr_position
    x As Long
    y As Long
End Type

' wzapi.h: struct game_object_identifier (game_object_identifier_new(psObj) is its constructor)
Public Type game_object_identifier
    typ As Long                         ' int type = OBJ_NUM_TYPES
    id As Long                          ' -1
    player As Long                      ' -1
End Type

' wzapi.h: struct droid_id_player
Public Type droid_id_player
    id As Long                          ' -1
    player As Long                      ' -1
End Type

' wzapi.h: struct optional_position
Public Type optional_position
    valid As Boolean
    x As Long
    y As Long
End Type

' wzapi.h: struct specified_player
Public Type specified_player
    player As Long                      ' -1
End Type

' wzapi.h: struct STRUCTURE_TYPE_or_statsName_string
Public Type STRUCTURE_TYPE_or_statsName_string
    typ As Long                         ' STRUCTURE_TYPE type = NUM_DIFF_BUILDINGS
    statsName As String
End Type

' wzapi.h: struct object_request (its constructors: object_request_new*)
Public Enum object_request_RequestType
    RequestType_INVALID_REQUEST
    RequestType_LABEL_REQUEST
    RequestType_MAPPOS_REQUEST
    RequestType_OBJECTID_REQUEST
End Enum

Public Type object_request
    requestType As object_request_RequestType
    str_ As String
    val1 As Long                        ' -1
    val2 As Long                        ' -1
    val3 As Long                        ' -1
End Type

' wzapi.h: struct label_or_position_values (optional<int32_t> x2, y2: has_x2, has_y2)
Public Enum label_or_position_values_Type
    Invalid_Request
    Label_Request
    Position_Values_Request
End Enum

Public Type label_or_position_values
    typ As label_or_position_values_Type
    x1 As Long                          ' VERY_LOW_INVALID_POS_VALUE = -2
    y1 As Long
    has_x2 As Boolean
    x2 As Long
    has_y2 As Boolean
    y2 As Long
    label As String
End Type

' wzapi.h: struct researchResult, wzapi_researchResult (research_c has the Sub researchResult) (const RESEARCH *: a research index, -1 = nullptr)
Public Type wzapi_researchResult
    psResearch As Long
    player As Long
End Type

' wzapi.h: struct researchResults, wzapi_researchResults
Public Type wzapi_researchResults
    resList() As Long                   ' std::vector<const RESEARCH *>: resList(0 .. resList_size - 1)
    resList_size As Long
    player As Long
End Type

' wzapi.h: enum class GlobalVariableFlags (bit flags)
Public Const GlobalVariableFlags_None As Long = 0
Public Const GlobalVariableFlags_ReadOnly As Long = 1       ' binary 0001
Public Const GlobalVariableFlags_ReadOnlyUpdatedFromApp As Long = 2 ' binary 0010
Public Const GlobalVariableFlags_DoNotSave As Long = 4      ' binary 0100

' (VB) a GameEntityRules::NameToTypeMap
Private Type NAME_TO_TYPE_MAP
    keys() As String
    values() As Long
    size As Long
End Type

' (VB) the maps of getUpgradesObject's initializer lists, built once
Private Const NTT_BODY As Long = 0
Private Const NTT_SENSOR As Long = 1
Private Const NTT_PROPULSION As Long = 2
Private Const NTT_ECM As Long = 3
Private Const NTT_REPAIR As Long = 4
Private Const NTT_CONSTRUCT As Long = 5
Private Const NTT_BRAIN As Long = 6
Private Const NTT_WEAPON As Long = 7
Private Const NTT_BUILDING As Long = 8
Private nameToTypeMaps(0 To 8) As NAME_TO_TYPE_MAP
Private nameToTypeMaps_built As Boolean

' wzapi::scriptQueuedObjectRemovals()'s static vector: each pair is <GameObject, NeedToApplyEffectsOnDestruction>
Private queuedObjRemovals_obj() As Long     ' BASE_OBJECT *
Private queuedObjRemovals_sfx() As Boolean
Private queuedObjRemovals_size As Long
Private queuedObjRemovals_capacity As Long  ' (VB) the arrays' length (reserve(initialCapacity) on first use)


Public Const ALL_PLAYERS As Long = -1
Public Const ALLIES As Long = -2
Public Const ENEMIES As Long = -3

' (VB) the RESEARCH * nullptr of this module: a research index of -1
Private Const NO_RESEARCH As Long = -1

' (VB) push_back on a vector of handles: result(0 .. n - 1), grown by doubling
Private Sub pushHandle(result() As Long, n As Long, ByVal h As Long)
    If n = 0 Then
        ReDim result(0 To 15)
    ElseIf n > UBound(result) Then
        ReDim Preserve result(0 To 2 * n - 1)
    End If
    result(n) = h
    n = n + 1
End Sub

' (VB) push_back on a std::vector<scr_position>
Private Sub pushPosition(result() As scr_position, n As Long, ByVal x As Long, ByVal y As Long)
    If n = 0 Then
        ReDim result(0 To 15)
    ElseIf n > UBound(result) Then
        ReDim Preserve result(0 To 2 * n - 1)
    End If
    result(n).x = x
    result(n).y = y
    n = n + 1
End Sub

' (VB) researchResults::resList.push_back
Private Sub pushResearch(r As wzapi_researchResults, ByVal psResearch As Long)
    If r.resList_size = 0 Then
        ReDim r.resList(0 To 15)
    ElseIf r.resList_size > UBound(r.resList) Then
        ReDim Preserve r.resList(0 To 2 * r.resList_size - 1)
    End If
    r.resList(r.resList_size) = psResearch
    r.resList_size = r.resList_size + 1
End Sub

' (VB) std::list<RESEARCH *>::push_back of pursueResearch's and findResearch's stack
Private Sub pushResList(reslist() As Long, reslist_tail As Long, ByVal psResearch As Long)
    If reslist_tail = 0 Then
        ReDim reslist(0 To 15)
    ElseIf reslist_tail > UBound(reslist) Then
        ReDim Preserve reslist(0 To 2 * reslist_tail - 1)
    End If
    reslist(reslist_tail) = psResearch
    reslist_tail = reslist_tail + 1
End Sub

Public Function IdToObject(ByVal typ As OBJECT_TYPE, ByVal id As Long, ByVal player As Long) As Long
    Select Case typ
    Case OBJ_DROID: IdToObject = IdToDroid(id, player)
    Case OBJ_FEATURE: IdToObject = IdToFeature(id, player)
    Case OBJ_STRUCTURE: IdToObject = IdToStruct(id, player)
    Case Else: IdToObject = 0
    End Select
End Function

' wzapi.h: game_object_identifier() (type OBJ_NUM_TYPES, id -1, player -1)
Public Function game_object_identifier_default() As game_object_identifier
    game_object_identifier_default.typ = OBJ_NUM_TYPES
    game_object_identifier_default.id = -1
    game_object_identifier_default.player = -1
End Function

' wzapi.h: game_object_identifier(const BASE_OBJECT* psObj)
Public Function game_object_identifier_new(ByVal psObj As Long) As game_object_identifier
    game_object_identifier_new.typ = SIMPLE_OBJECT_type(psObj)
    game_object_identifier_new.id = SIMPLE_OBJECT_id(psObj)
    game_object_identifier_new.player = SIMPLE_OBJECT_player(psObj)
End Function

Public Function object_request_new() As object_request
    object_request_new.requestType = RequestType_INVALID_REQUEST
    object_request_new.val1 = -1
    object_request_new.val2 = -1
    object_request_new.val3 = -1
End Function

Public Function object_request_new_label(ByVal label As String) As object_request
    object_request_new_label.requestType = RequestType_LABEL_REQUEST
    object_request_new_label.str_ = label
    object_request_new_label.val1 = -1
    object_request_new_label.val2 = -1
    object_request_new_label.val3 = -1
End Function

Public Function object_request_new_mappos(ByVal x As Long, ByVal y As Long) As object_request
    object_request_new_mappos.requestType = RequestType_MAPPOS_REQUEST
    object_request_new_mappos.val1 = x
    object_request_new_mappos.val2 = y
    object_request_new_mappos.val3 = -1
End Function

Public Function object_request_new_objectid(ByVal typ As OBJECT_TYPE, ByVal player As Long, ByVal id As Long) As object_request
    object_request_new_objectid.requestType = RequestType_OBJECTID_REQUEST
    object_request_new_objectid.val1 = typ
    object_request_new_objectid.val2 = player
    object_request_new_objectid.val3 = id
End Function

Public Function object_request_getLabel(r As object_request) As String
    If Not (r.requestType = RequestType_LABEL_REQUEST) Then ASSERT_FAILURE "requestType == wzapi::object_request::RequestType::LABEL_REQUEST", "src/wzapi.cpp:260", "getLabel", "Not a label request"
    object_request_getLabel = r.str_
End Function

Public Function object_request_getMapPosition(r As object_request) As scr_position
    If Not (r.requestType = RequestType_MAPPOS_REQUEST) Then ASSERT_FAILURE "requestType == wzapi::object_request::RequestType::MAPPOS_REQUEST", "src/wzapi.cpp:265", "getMapPosition", "Not a map position request"
    object_request_getMapPosition.x = r.val1
    object_request_getMapPosition.y = r.val2
End Function

' std::tuple<OBJECT_TYPE, int, int>: its three values ByRef
Public Sub object_request_getObjectIDRequest(r As object_request, typ As OBJECT_TYPE, ByRef player As Long, ByRef id As Long)
    If Not (r.requestType = RequestType_OBJECTID_REQUEST) Then ASSERT_FAILURE "requestType == wzapi::object_request::RequestType::OBJECTID_REQUEST", "src/wzapi.cpp:270", "getObjectIDRequest", "Not an object ID request"
    typ = r.val1
    player = r.val2
    id = r.val3
End Sub

' wzapi.h: label_or_position_values() (Invalid_Request, x1 = y1 = VERY_LOW_INVALID_POS_VALUE)
Public Function label_or_position_values_new() As label_or_position_values
    label_or_position_values_new.typ = Invalid_Request
    label_or_position_values_new.x1 = -2
    label_or_position_values_new.y1 = -2
End Function

Public Function label_or_position_values_new_label(ByVal label As String) As label_or_position_values
    label_or_position_values_new_label.typ = Label_Request
    label_or_position_values_new_label.x1 = -2
    label_or_position_values_new_label.y1 = -2
    label_or_position_values_new_label.label = label
End Function

' label_or_position_values(int32_t x1, int32_t y1, optional<int32_t> x2 = nullopt, optional<int32_t> y2 = nullopt)
Public Function label_or_position_values_new_values(ByVal x1 As Long, ByVal y1 As Long, ByVal has_x2 As Boolean, ByVal x2 As Long, ByVal has_y2 As Boolean, ByVal y2 As Long) As label_or_position_values
    label_or_position_values_new_values.typ = Position_Values_Request
    label_or_position_values_new_values.x1 = x1
    label_or_position_values_new_values.y1 = y1
    label_or_position_values_new_values.has_x2 = has_x2
    If has_x2 Then label_or_position_values_new_values.x2 = x2
    label_or_position_values_new_values.has_y2 = has_y2
    If has_y2 Then label_or_position_values_new_values.y2 = y2
End Function

Public Function label_or_position_values_isValid(v As label_or_position_values) As Boolean
    label_or_position_values_isValid = (v.typ <> Invalid_Request)
End Function

Public Function label_or_position_values_isLabel(v As label_or_position_values) As Boolean
    label_or_position_values_isLabel = (v.typ = Label_Request)
End Function

Public Function label_or_position_values_isPositionValues(v As label_or_position_values) As Boolean
    label_or_position_values_isPositionValues = (v.typ = Position_Values_Request)
End Function

' MARK: - Labels



' MARK: -

'-- ## _(string)
'--
'-- Mark string for translation.
'--
Public Function wzapi_translate(ByVal context As wzapi_execution_context, ByVal str_ As String) As String
    If Len(str_) = 0 Then wzapi_translate = str_: Exit Function
    wzapi_translate = str_              ' std::string(gettext(str.c_str())): gettext(s) is s
End Function

'-- ## syncRandom(limit)
'--
'-- Generate a synchronized random number in range 0...(limit - 1) that will be the same if this function is
'-- run on all network peers in the same game frame. If it is called on just one peer (such as would be
'-- the case for AIs, for instance), then game sync will break. (3.2+ only)
'--
Public Function wzapi_syncRandom(ByVal context As wzapi_execution_context, ByVal limit As Long) As Long
    If limit = 0 Then wzapi_syncRandom = 0: Exit Function
    wzapi_syncRandom = gameRand(limit)
End Function

'-- ## setAlliance(player1, player2, areAllies)
'--
'-- Set alliance status between two players to either true or false. (3.2+ only)
'--
Public Function wzapi_setAlliance(ByVal context As wzapi_execution_context, ByVal player1 As Long, ByVal player2 As Long, ByVal areAllies As Boolean) As Boolean
    If areAllies Then
        formAlliance player1, player2, True, False, True
    Else
        breakAlliance player1, player2, True, True
    End If
    wzapi_setAlliance = True
End Function

'-- ## sendAllianceRequest(player)
'--
'-- Send an alliance request to a player. (3.3+ only)
'--
Public Sub wzapi_sendAllianceRequest(ByVal context As wzapi_execution_context, ByVal player As Long)
    If Not alliancesFixed(game.alliance) Then
        requestAlliance context.player(), player, True, True
    End If
End Sub

'-- ## orderDroid(droid, order)
'--
'-- Give a droid an order to do something. (3.2+ only)
'--
Public Function wzapi_orderDroid(ByVal context As wzapi_execution_context, ByVal psDroid As Long, ByVal order As Long) As Boolean
    Dim psStruct As Long
    If Not (psDroid <> 0) Then debug_ 338, LOG_ERROR, "orderDroid", "No valid droid provided": context.throwError "psDroid", 338, "orderDroid": wzapi_orderDroid = False: Exit Function
    If Not (order = DORDER_STOP Or order = DORDER_RTB Or order = DORDER_RTR Or order = DORDER_RECYCLE Or order = DORDER_REARM Or order = DORDER_HOLD) Then
        debug_ 339, LOG_ERROR, "orderDroid", "Invalid order: %s", getDroidOrderName(order)
        context.throwError "order == DORDER_STOP || order == DORDER_RTB || order == DORDER_RTR || order == DORDER_RECYCLE || order == DORDER_REARM || order == DORDER_HOLD", 339, "orderDroid"
        wzapi_orderDroid = False
        Exit Function
    End If

    If aDroid(psDroid).order.typ = order Then
        wzapi_orderDroid = True
        Exit Function
    End If
    If order = DORDER_REARM Then
        psStruct = findNearestReArmPad(psDroid, aDroid(psDroid).psBaseStruct, False)
        If psStruct <> 0 Then
            orderDroidObj psDroid, order, psStruct, ModeQueue
        Else
            orderDroid psDroid, DORDER_RTB, ModeQueue
        End If
    Else
        orderDroid psDroid, order, ModeQueue
    End If
    wzapi_orderDroid = True
End Function

'-- ## orderDroidBuild(droid, order, structureName, x, y[, direction])
'--
'-- Give a droid an order to build something at the given position. Returns true if allowed.
'--
Public Function wzapi_orderDroidBuild(ByVal context As wzapi_execution_context, ByVal psDroid As Long, ByVal order As Long, ByVal structureName As String, ByVal x As Long, ByVal y As Long, ByVal has_direction As Boolean, ByVal opt_direction As Single) As Boolean
    Dim structureIndex As Long, psStats As Long, direction As Long, directionValue As Single
    If Not (psDroid <> 0) Then debug_ 372, LOG_ERROR, "orderDroidBuild", "No valid droid provided": context.throwError "psDroid", 372, "orderDroidBuild": wzapi_orderDroidBuild = False: Exit Function

    structureIndex = getStructStatFromName(structureName)
    If Not (structureIndex >= 0 And structureIndex < numStructureStats) Then debug_ 375, LOG_ERROR, "orderDroidBuild", "Structure %s not found", structureName: context.throwError "structureIndex >= 0 && structureIndex < numStructureStats", 375, "orderDroidBuild": wzapi_orderDroidBuild = False: Exit Function
    psStats = structureIndex

    If Not (order = DORDER_BUILD) Then debug_ 378, LOG_ERROR, "orderDroidBuild", "Invalid order": context.throwError "order == DORDER_BUILD", 378, "orderDroidBuild": wzapi_orderDroidBuild = False: Exit Function
    If Not (asStructureStats(psStats).id <> "A0ADemolishStructure") Then debug_ 379, LOG_ERROR, "orderDroidBuild", "Cannot build demolition": context.throwError "psStats->id.compare(""A0ADemolishStructure"") != 0", 379, "orderDroidBuild": wzapi_orderDroidBuild = False: Exit Function

    If has_direction Then
        If IsNaNF(opt_direction) Then
            opt_direction = 0! ' avoid undefined behavior (nan is outside the range of representable values of type 'unsigned short')
        End If
    End If
    If has_direction Then directionValue = opt_direction Else directionValue = 0!
    direction = ToInt(fdiv(fmul(directionValue, 8192!), 45!)) And &HFFFF&    ' static_cast<uint16_t>(DEG(_direction.value_or(0)))

    If aDroid(psDroid).order.typ = order And aDroid(psDroid).actionPos.x = world_coord(x) And aDroid(psDroid).actionPos.y = world_coord(y) Then
        wzapi_orderDroidBuild = True
        Exit Function
    End If
    orderDroidStatsLocDir psDroid, order, psStats, world_coord(x) + TILE_UNITS \ 2, world_coord(y) + TILE_UNITS \ 2, direction, ModeQueue
    wzapi_orderDroidBuild = True
End Function

'-- ## setAssemblyPoint(structure, x, y)
'--
'-- Set the assembly point droids go to when built for the specified structure. (3.2+ only)
'--
Public Function wzapi_setAssemblyPoint(ByVal context As wzapi_execution_context, ByVal psStruct As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim t As STRUCTURE_TYPE
    If Not (psStruct <> 0) Then debug_ 402, LOG_ERROR, "setAssemblyPoint", "No valid structure provided": context.throwError "psStruct", 402, "setAssemblyPoint": wzapi_setAssemblyPoint = False: Exit Function
    t = asStructureStats(aStruct(psStruct).pStructureType).typ
    If Not (t = REF_FACTORY Or t = REF_CYBORG_FACTORY Or t = REF_VTOL_FACTORY) Then
        debug_ 403, LOG_ERROR, "setAssemblyPoint", "Structure not a factory"
        context.throwError "psStruct->pStructureType->type == REF_FACTORY || psStruct->pStructureType->type == REF_CYBORG_FACTORY || psStruct->pStructureType->type == REF_VTOL_FACTORY", 403, "setAssemblyPoint"
        wzapi_setAssemblyPoint = False
        Exit Function
    End If
    setAssemblyPoint aStruct(psStruct).functionality.factory.psAssemblyPoint, x, y, aStruct(psStruct).player, True
    wzapi_setAssemblyPoint = True
End Function

'-- ## setSunPosition(x, y, z)
'--
'-- Move the position of the Sun, which in turn moves where shadows are cast. (3.2+ only)
'--
Public Function wzapi_setSunPosition(ByVal context As wzapi_execution_context, ByVal x As Single, ByVal y As Single, ByVal z As Single) As Boolean
    Dim newSun As Vector3f
    If Not (Not IsNaNF(x) And Not IsNaNF(y) And Not IsNaNF(z)) Then debug_ 416, LOG_ERROR, "setSunPosition", "Inputs must not be nan": context.throwError "!std::isnan(x) && !std::isnan(y) && !std::isnan(z)", 416, "setSunPosition": wzapi_setSunPosition = False: Exit Function
    newSun.x = x
    newSun.y = y
    newSun.z = z
    setTheSun newSun
    wzapi_setSunPosition = True
End Function

'-- ## setSunIntensity(ambient_r, ambient_g, ambient_b, diffuse_r, diffuse_g, diffuse_b, specular_r, specular_g, specular_b)
'--
'-- Set the ambient, diffuse and specular colour intensities of the Sun lighting source. (3.2+ only)
'--
Public Function wzapi_setSunIntensity(ByVal context As wzapi_execution_context, ByVal ambient_r As Single, ByVal ambient_g As Single, ByVal ambient_b As Single, ByVal diffuse_r As Single, ByVal diffuse_g As Single, ByVal diffuse_b As Single, ByVal specular_r As Single, ByVal specular_g As Single, ByVal specular_b As Single) As Boolean
    Dim ambient(0 To 3) As Single, diffuse(0 To 3) As Single, specular(0 To 3) As Single
    If Not (Not IsNaNF(ambient_r) And Not IsNaNF(ambient_g) And Not IsNaNF(ambient_b)) Then debug_ 427, LOG_ERROR, "setSunIntensity", "ambient inputs must not be nan": context.throwError "!std::isnan(ambient_r) && !std::isnan(ambient_g) && !std::isnan(ambient_b)", 427, "setSunIntensity": wzapi_setSunIntensity = False: Exit Function
    If Not (Not IsNaNF(diffuse_r) And Not IsNaNF(diffuse_g) And Not IsNaNF(diffuse_b)) Then debug_ 428, LOG_ERROR, "setSunIntensity", "diffuse inputs must not be nan": context.throwError "!std::isnan(diffuse_r) && !std::isnan(diffuse_g) && !std::isnan(diffuse_b)", 428, "setSunIntensity": wzapi_setSunIntensity = False: Exit Function
    If Not (Not IsNaNF(specular_r) And Not IsNaNF(specular_g) And Not IsNaNF(specular_b)) Then debug_ 429, LOG_ERROR, "setSunIntensity", "specular inputs must not be nan": context.throwError "!std::isnan(specular_r) && !std::isnan(specular_g) && !std::isnan(specular_b)", 429, "setSunIntensity": wzapi_setSunIntensity = False: Exit Function

    ambient(0) = ambient_r
    ambient(1) = ambient_g
    ambient(2) = ambient_b
    ambient(3) = 1!
    diffuse(0) = diffuse_r
    diffuse(1) = diffuse_g
    diffuse(2) = diffuse_b
    diffuse(3) = 1!
    specular(0) = specular_r
    specular(1) = specular_g
    specular(2) = specular_b
    specular(3) = 1!
    pie_Lighting0 LIGHT_AMBIENT, ambient
    pie_Lighting0 LIGHT_DIFFUSE, diffuse
    pie_Lighting0 LIGHT_SPECULAR, specular
    wzapi_setSunIntensity = True
End Function

'-- ## setFogColour(r, g, b)
'--
'-- Set the colour of the fog (4.2.5+ only)
'--
Public Function wzapi_setFogColour(ByVal context As wzapi_execution_context, ByVal r As Long, ByVal g As Long, ByVal b As Long) As Boolean
    Dim newFogColour As Long
    newFogColour = pal_RGBA(r, g, b, 255)
    pie_SetFogColour newFogColour
    wzapi_setFogColour = True
End Function

'-- ## setWeather(weatherType)
'--
'-- Set the current weather. This should be one of ```WEATHER_RAIN```, ```WEATHER_SNOW``` or ```WEATHER_CLEAR```. (3.2+ only)
'--
Public Function wzapi_setWeather(ByVal context As wzapi_execution_context, ByVal weatherType As Long) As Boolean
    If Not (weatherType >= 0 And weatherType <= WT_NONE) Then debug_ 469, LOG_ERROR, "setWeather", "Bad weather type": context.throwError "weatherType >= 0 && weatherType <= WT_NONE", 469, "setWeather": wzapi_setWeather = False: Exit Function
    atmosSetWeatherType weatherType
    wzapi_setWeather = True
End Function

'-- ## setSky(textureFilename, windSpeed, scale)
'--
'-- Change the skybox. (3.2+ only)
'--
Public Function wzapi_setSky(ByVal context As wzapi_execution_context, ByVal textureFilename As String, ByVal windSpeed As Single, ByVal scale_ As Single) As Boolean
    If Not (Not IsNaNF(windSpeed)) Then debug_ 480, LOG_ERROR, "setSky", "windSpeed must not be nan": context.throwError "!std::isnan(windSpeed)", 480, "setSky": wzapi_setSky = False: Exit Function
    If Not (Not IsNaNF(scale_)) Then debug_ 481, LOG_ERROR, "setSky", "scale must not be nan": context.throwError "!std::isnan(scale)", 481, "setSky": wzapi_setSky = False: Exit Function

    setSkyBox textureFilename, windSpeed, scale_
    wzapi_setSky = True ' TODO: modify setSkyBox to return bool, success / failure
End Function

'-- ## cameraSlide(x, y)
'--
'-- Slide the camera over to the given position on the map. (3.2+ only)
'--
Public Function wzapi_cameraSlide(ByVal context As wzapi_execution_context, ByVal x As Single, ByVal y As Single) As Boolean
    If Not (Not IsNaNF(x)) Then debug_ 493, LOG_ERROR, "cameraSlide", "x must not be nan": context.throwError "!std::isnan(x)", 493, "cameraSlide": wzapi_cameraSlide = False: Exit Function
    If Not (Not IsNaNF(y)) Then debug_ 494, LOG_ERROR, "cameraSlide", "y must not be nan": context.throwError "!std::isnan(y)", 494, "cameraSlide": wzapi_cameraSlide = False: Exit Function

    requestRadarTrack ToInt(x), ToInt(y)
    wzapi_cameraSlide = True
End Function

'-- ## cameraZoom(viewDistance, speed)
'--
'-- Slide the camera to the given zoom distance. Normal camera zoom ranges between 500 and 5000. (3.2+ only)
'--
Public Function wzapi_cameraZoom(ByVal context As wzapi_execution_context, ByVal viewDistance As Single, ByVal speed As Single) As Boolean
    If Not (Not IsNaNF(viewDistance)) Then debug_ 506, LOG_ERROR, "cameraZoom", "viewDistance must not be nan": context.throwError "!std::isnan(viewDistance)", 506, "cameraZoom": wzapi_cameraZoom = False: Exit Function
    If Not (Not IsNaNF(speed)) Then debug_ 507, LOG_ERROR, "cameraZoom", "speed must not be nan": context.throwError "!std::isnan(speed)", 507, "cameraZoom": wzapi_cameraZoom = False: Exit Function

    animateToViewDistance viewDistance, speed
    wzapi_cameraZoom = True
End Function

'-- ## cameraTrack([droid])
'--
'-- Make the camera follow the given droid object around. Pass in a null object to stop. (3.2+ only)
'--
Public Function wzapi_cameraTrack(ByVal context As wzapi_execution_context, ByVal has_droid As Boolean, ByVal opt_droid As Long) As Boolean
    Dim droid_ As Long, it As Long, psDroid As Long
    If has_droid Then
        droid_ = opt_droid
        If Not (droid_ <> 0) Then debug_ 522, LOG_ERROR, "cameraTrack", "No valid droid provided": context.throwError "droid", 522, "cameraTrack": wzapi_cameraTrack = False: Exit Function
        If Not (selectedPlayer < MAX_PLAYERS) Then debug_ 523, LOG_ERROR, "cameraTrack", "Invalid selectedPlayer for current client: %u", selectedPlayer: context.throwError "selectedPlayer < MAX_PLAYERS", 523, "cameraTrack": wzapi_cameraTrack = False: Exit Function
        it = apsDroidLists(selectedPlayer).head
        Do While it <> 0
            psDroid = lnObj(it)
            aDroid(psDroid).selected = -CLng(psDroid = droid_) ' select only the target droid
            it = lnNext(it)
        Loop
        setWarCamActive True
    Else
        setWarCamActive False
    End If
    wzapi_cameraTrack = True
End Function

'-- ## addSpotter(x, y, player, range, radar, expiry)
'--
'-- Add an invisible viewer at a given position for given player that shows map in given range. ```radar```
'-- is false for vision reveal, or true for radar reveal. The difference is that a radar reveal can be obstructed
'-- by ECM jammers. ```expiry```, if non-zero, is the game time at which the spotter shall automatically be
'-- removed. The function returns a unique ID that can be used to remove the spotter with ```removeSpotter```. (3.2+ only)
'--
Public Function wzapi_addSpotter(ByVal context As wzapi_execution_context, ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal range As Long, ByVal radar As Boolean, ByVal expiry As Long) As Long
    wzapi_addSpotter = addSpotter(x, y, player, range, radar, expiry)
End Function

'-- ## removeSpotter(spotterId)
'--
'-- Remove a spotter given its unique ID. (3.2+ only)
'--
Public Function wzapi_removeSpotter(ByVal context As wzapi_execution_context, ByVal spotterId As Long) As Boolean
    wzapi_removeSpotter = removeSpotter(spotterId)
End Function

'-- ## syncRequest(req_id, x, y[, object[, object2]])
'--
'-- Generate a synchronized event request that is sent over the network to all clients and executed simultaneously.
'-- Must be caught in an eventSyncRequest() function. All sync requests must be validated when received, and always
'-- take care only to define sync requests that can be validated against cheating. (3.2+ only)
'--
Public Function wzapi_syncRequest(ByVal context As wzapi_execution_context, ByVal req_id As Long, ByVal x_ As Long, ByVal y_ As Long, ByVal has_psObj As Boolean, ByVal opt_psObj As Long, ByVal has_psObj2 As Boolean, ByVal opt_psObj2 As Long) As Boolean
    Dim x As Long, y As Long, psObj As Long, psObj2 As Long
    x = world_coord(x_)
    y = world_coord(y_)
    psObj = 0
    psObj2 = 0
    If has_psObj Then
        psObj = opt_psObj
        If Not (psObj <> 0) Then debug_ 572, LOG_ERROR, "syncRequest", "No valid object (obj1) provided": context.throwError "psObj", 572, "syncRequest": wzapi_syncRequest = False: Exit Function
    End If
    If has_psObj2 Then
        psObj2 = opt_psObj2
        If Not (psObj2 <> 0) Then debug_ 577, LOG_ERROR, "syncRequest", "No valid object (obj2) provided": context.throwError "psObj2", 577, "syncRequest": wzapi_syncRequest = False: Exit Function
    End If
    sendSyncRequest req_id, x, y, psObj, psObj2
    wzapi_syncRequest = True
End Function

'-- ## replaceTexture(oldFilename, newFilename)
'--
'-- Replace one texture with another. This can be used to for example give buildings on a specific tileset different
'-- looks, or to add variety to the looks of droids in campaign missions. (3.2+ only)
'--
Public Function wzapi_replaceTexture(ByVal context As wzapi_execution_context, ByVal oldFilename As String, ByVal newFilename As String) As Boolean
    wzapi_replaceTexture = replaceTexture(oldFilename, newFilename)
End Function

'-- ## changePlayerColour(player, colour)
'--
'-- Change a player's colour slot. The current player colour can be read from the ```playerData``` array. Available colours
'-- are green, orange, gray, black, red, blue, pink, cyan, yellow, purple, white, bright blue, neon green, infrared,
'-- ultraviolet, and brown, represented by the integers 0 - 15 respectively.
'--
Public Function wzapi_changePlayerColour(ByVal context As wzapi_execution_context, ByVal player As Long, ByVal colour As Long) As Boolean
    ' (VB) component.cpp's setPlayerColour (its stand-in until cluster D) is a Sub: C++ returns its bool
    setPlayerColour player, colour
    wzapi_changePlayerColour = True
End Function

'-- ## setHealth(object, health)
'--
'-- Change the health of the given game object, in percentage. Does not take care of network sync, so for multiplayer games,
'-- needs wrapping in a syncRequest. (3.2.3+ only.)
'--
Public Function wzapi_setHealth(ByVal context As wzapi_execution_context, ByVal psObject As Long, ByVal health As Long) As Boolean
    Dim id As Long, player As Long, objectType As OBJECT_TYPE, psDroid As Long, psStruct As Long, psFeat As Long, m As Long
    If Not (psObject <> 0) Then debug_ 611, LOG_ERROR, "setHealth", "No valid object provided": context.throwError "psObject", 611, "setHealth": wzapi_setHealth = False: Exit Function
    If Not (health >= 1) Then debug_ 612, LOG_ERROR, "setHealth", "Bad health value %d", health: context.throwError "health >= 1", 612, "setHealth": wzapi_setHealth = False: Exit Function
    id = SIMPLE_OBJECT_id(psObject)
    player = SIMPLE_OBJECT_player(psObject)
    objectType = SIMPLE_OBJECT_type(psObject)
    If Not (objectType = OBJ_DROID Or objectType = OBJ_STRUCTURE Or objectType = OBJ_FEATURE) Then debug_ 616, LOG_ERROR, "setHealth", "Bad object type": context.throwError "objectType == OBJ_DROID || objectType == OBJ_STRUCTURE || objectType == OBJ_FEATURE", 616, "setHealth": wzapi_setHealth = False: Exit Function
    If objectType = OBJ_DROID Then
        psDroid = psObject
        If Not (psDroid <> 0) Then debug_ 620, LOG_ERROR, "setHealth", "No such droid id %d belonging to player %d", id, player: context.throwError "psDroid", 620, "setHealth": wzapi_setHealth = False: Exit Function
        aDroid(psDroid).body = ToInt(CDbl(health) * CDbl(aDroid(psDroid).originalBody) / 100#)
    ElseIf objectType = OBJ_STRUCTURE Then
        psStruct = psObject
        If Not (psStruct <> 0) Then debug_ 626, LOG_ERROR, "setHealth", "No such structure id %d belonging to player %d", id, player: context.throwError "psStruct", 626, "setHealth": wzapi_setHealth = False: Exit Function
        m = STRUCTURE_structureBody(psStruct)
        If m < 1 Then m = 1                 ' MAX(1, psStruct->structureBody())
        aStruct(psStruct).body = Fix(CDbl(health) * CDbl(m) / 100#)
    Else
        psFeat = psObject
        If Not (psFeat <> 0) Then debug_ 632, LOG_ERROR, "setHealth", "No such feature id %d belonging to player %d", id, player: context.throwError "psFeat", 632, "setHealth": wzapi_setHealth = False: Exit Function
        aFeature(psFeat).body = Fix(CDbl(health) * CDbl(asFeatureStats(aFeature(psFeat).psStats).body) / 100#)
    End If
    wzapi_setHealth = True
End Function

'-- ## useSafetyTransport(flag)
'--
'-- Change if the mission transporter will fetch droids in non offworld missions
'-- setReinforcementTime() is be used to hide it before coming back after the set time
'-- which is handled by the campaign library in the victory data section (3.3+ only).
'--
Public Function wzapi_useSafetyTransport(ByVal context As wzapi_execution_context, ByVal flag As Boolean) As Boolean
    setDroidsToSafetyFlag flag
    wzapi_useSafetyTransport = True
End Function

'-- ## restoreLimboMissionData()
'--
'-- Swap mission type and bring back units previously stored at the start
'-- of the mission (see cam3-c mission). (3.3+ only).
'--
Public Function wzapi_restoreLimboMissionData(ByVal context As wzapi_execution_context) As Boolean
    resetLimboMission
    wzapi_restoreLimboMissionData = True
End Function

'-- ## getMultiTechLevel()
'--
'-- Returns the current multiplayer tech level. (3.3+ only)
'--
Public Function wzapi_getMultiTechLevel(ByVal context As wzapi_execution_context) As Long
    wzapi_getMultiTechLevel = game.techLevel
End Function

'-- ## setCampaignNumber(campaignNumber)
'--
'-- Set the campaign number. (3.3+ only)
'--
Public Function wzapi_setCampaignNumber(ByVal context As wzapi_execution_context, ByVal campaignNumber As Long) As Boolean
    setCampaignNumber campaignNumber
    wzapi_setCampaignNumber = True
End Function

'-- ## getMissionType()
'--
'-- Return the current mission type. (3.3+ only)
'--
Public Function wzapi_getMissionType(ByVal context As wzapi_execution_context) As Long
    wzapi_getMissionType = mission_type
End Function

'-- ## getRevealStatus()
'--
'-- Return the current fog reveal status. (3.3+ only)
'--
Public Function wzapi_getRevealStatus(ByVal context As wzapi_execution_context) As Boolean
    wzapi_getRevealStatus = getRevealStatus()
End Function

'-- ## setRevealStatus(status)
'--
'-- Set the fog reveal status. (3.3+ only)
'--
Public Function wzapi_setRevealStatus(ByVal context As wzapi_execution_context, ByVal status As Boolean) As Boolean
    setRevealStatus status
    preProcessVisibility
    wzapi_setRevealStatus = True
End Function

'-- ## autoSave()
'--
'-- Perform automatic save
'--
Public Function wzapi_autoSave(ByVal context As wzapi_execution_context) As Boolean
    wzapi_autoSave = autoSave(False)
End Function

' MARK: - horrible hacks follow -- do not rely on these being present!

'-- ## hackNetOff()
'--
'-- Turn off network transmissions. FIXME - find a better way.
'--
Public Sub wzapi_hackNetOff(ByVal context As wzapi_execution_context)
    bMultiPlayer = False
    bMultiMessages = False
End Sub

'-- ## hackNetOn()
'--
'-- Turn on network transmissions. FIXME - find a better way.
'--
Public Sub wzapi_hackNetOn(ByVal context As wzapi_execution_context)
    bMultiPlayer = True
    bMultiMessages = True
End Sub

'-- ## hackAddMessage(message, messageType, player, immediate)
'--
'-- See wzscript docs for info, to the extent any exist. (3.2+ only)
'--
Public Sub wzapi_hackAddMessage(ByVal context As wzapi_execution_context, ByVal message As String, ByVal messageType As Long, ByVal player As Long, ByVal immediate As Boolean)
    Dim msgType As MESSAGE_TYPE, psMessage As Long, psViewData As Long, height As Long
    msgType = messageType
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 749, LOG_ERROR, "hackAddMessage", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 749, "hackAddMessage": Exit Sub
    psMessage = addMessage(msgType, False, player)
    If psMessage <> 0 Then
        psViewData = getViewData(message)
        If Not (psViewData <> 0) Then debug_ 754, LOG_ERROR, "hackAddMessage", "Viewdata not found": context.throwError "psViewData", 754, "hackAddMessage": Exit Sub
        MESSAGE_pViewData(psMessage) = psViewData
        debug_ 756, LOG_MSG, "hackAddMessage", "Adding %s pViewData=%p", VIEWDATA_name(psViewData), MESSAGE_pViewData(psMessage)
        If msgType = MSG_PROXIMITY Then
            ' check the z value is at least the height of the terrain
            height = map_Height(VIEW_PROXIMITY_x(psViewData), VIEW_PROXIMITY_y(psViewData))
            If VIEW_PROXIMITY_z(psViewData) < height Then
                VIEW_PROXIMITY_z(psViewData) = height
            End If
        End If
        If immediate Then
            displayImmediateMessage psMessage
        End If
    End If
'	jsDebugMessageUpdate();
End Sub

'-- ## hackRemoveMessage(message, messageType, player)
'--
'-- See wzscript docs for info, to the extent any exist. (3.2+ only)
'--
Public Sub wzapi_hackRemoveMessage(ByVal context As wzapi_execution_context, ByVal message As String, ByVal messageType As Long, ByVal player As Long)
    Dim msgType As MESSAGE_TYPE, psViewData As Long, psMessage As Long
    msgType = messageType
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 783, LOG_ERROR, "hackRemoveMessage", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 783, "hackRemoveMessage": Exit Sub
    psViewData = getViewData(message)
    If Not (psViewData <> 0) Then debug_ 785, LOG_ERROR, "hackRemoveMessage", "Viewdata not found": context.throwError "psViewData", 785, "hackRemoveMessage": Exit Sub
    psMessage = findMessage_viewData(psViewData, msgType, player)
    If psMessage <> 0 Then
        debug_ 789, LOG_MSG, "hackRemoveMessage", "Removing %s", VIEWDATA_name(psViewData)
        removeMessage psMessage, player
    Else
        debug_ 794, LOG_ERROR, "hackRemoveMessage", "cannot find message - %s", VIEWDATA_name(psViewData)
    End If
'	jsDebugMessageUpdate();
End Sub

'-- ## hackGetObj(objectType, player, id)
'--
'-- Function to find and return a game object of ```DROID```, ```FEATURE``` or ```STRUCTURE``` types, if it exists.
'-- Otherwise, it will return null. This function is DEPRECATED by getObject(). (3.2+ only)
'--
Public Function wzapi_hackGetObj(ByVal context As wzapi_execution_context, ByVal objectType_ As Long, ByVal player As Long, ByVal id As Long) As Long
    Dim objectType As OBJECT_TYPE
    objectType = objectType_
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 808, LOG_ERROR, "hackGetObj", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 808, "hackGetObj": wzapi_hackGetObj = 0: Exit Function
    wzapi_hackGetObj = IdToObject(objectType, id, player)
End Function

'-- ## hackAssert(condition, message...)
'--
'-- Function to perform unit testing. It will throw a script error and a game assert. (3.2+ only)
'--
Public Sub wzapi_hackAssert(ByVal context As wzapi_execution_context, ByVal condition As Boolean, message As StringVector)
    Dim result As String, i As Long
    If condition Then
        Exit Sub ' pass
    End If
    ' fail
    For i = 0 To message.size - 1
        If Len(result) <> 0 Then
            result = result & " "
        End If
        result = result & message.items(i)
    Next
    context.throwError result, 832, "hackAssert"
End Sub

'-- ## receiveAllEvents([enabled])
'--
'-- Make the current script receive all events, even those not meant for 'me'. (3.2+ only)
'--
Public Function wzapi_receiveAllEvents(ByVal context As wzapi_execution_context, ByVal has_enabled As Boolean, ByVal opt_enabled As Boolean) As Boolean
    If has_enabled Then
        context.set_isReceivingAllEvents opt_enabled
    End If
    wzapi_receiveAllEvents = context.get_isReceivingAllEvents()
End Function

'-- ## hackDoNotSave(name)
'--
'-- Do not save the given global given by name to savegames. Must be
'-- done again each time game is loaded, since this too is not saved.
'--
Public Sub wzapi_hackDoNotSave(ByVal context As wzapi_execution_context, ByVal name_ As String)
    context.doNotSaveGlobal name_
End Sub

'-- ## hackPlayIngameAudio()
'--
'-- (3.3+ only)
'--
Public Sub wzapi_hackPlayIngameAudio(ByVal context As wzapi_execution_context)
    debug_ 866, LOG_SOUND, "hackPlayIngameAudio", "Script wanted music to start"
    cdAudio_PlayTrack SONG_INGAME
End Sub

'-- ## hackStopIngameAudio()
'--
'-- Stop the in-game music. (3.3+ only)
'-- This should be called from the eventStartLevel() event (or later).
'-- Currently only used from the tutorial.
'--
Public Sub wzapi_hackStopIngameAudio(ByVal context As wzapi_execution_context)
    debug_ 879, LOG_SOUND, "hackStopIngameAudio", "Script wanted music to stop"
    cdAudio_Stop
End Sub

'-- ## hackMarkTiles([label | x, y[, x2, y2]])
'--
'-- Mark the given tile(s) on the map. Either give a ```POSITION``` or ```AREA``` label,
'-- or a tile x, y position, or four positions for a square area. If no parameter
'-- is given, all marked tiles are cleared. (3.2+ only)
'--
Public Sub wzapi_hackMarkTiles(ByVal context As wzapi_execution_context, ByVal has_tilePosOrArea As Boolean, opt_tilePosOrArea As label_or_position_values)
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, x As Long, y As Long, psTile As Long
    If has_tilePosOrArea Then
        If label_or_position_values_isLabel(opt_tilePosOrArea) Then ' label
            scripting_engine_hackMarkTiles_ByLabel context, opt_tilePosOrArea.label
            Exit Sub
        ElseIf label_or_position_values_isPositionValues(opt_tilePosOrArea) Then
            If opt_tilePosOrArea.has_x2 Or opt_tilePosOrArea.has_y2 Then
                If Not (opt_tilePosOrArea.has_x2 And opt_tilePosOrArea.has_y2) Then debug_ 904, LOG_ERROR, "hackMarkTiles", "If x2 or y2 are provided, *both* must be provided": context.throwError "tilePosOrArea.x2.has_value() && tilePosOrArea.y2.has_value()", 904, "hackMarkTiles": Exit Sub
                x1 = opt_tilePosOrArea.x1
                y1 = opt_tilePosOrArea.y1
                x2 = opt_tilePosOrArea.x2
                y2 = opt_tilePosOrArea.y2
                For x = x1 To x2 - 1
                    For y = y1 To y2 - 1
                        psTile = mapTile(x, y)
                        psMapTiles(psTile).tileInfoBits = psMapTiles(psTile).tileInfoBits Or BITS_MARKED
                    Next
                Next
            Else ' single tile
                x = opt_tilePosOrArea.x1
                y = opt_tilePosOrArea.y1
                psTile = mapTile(x, y)
                psMapTiles(psTile).tileInfoBits = psMapTiles(psTile).tileInfoBits Or BITS_MARKED
            End If
        End If
    Else ' clear all marks
        clearMarks
    End If
End Sub

' MARK: - General functions -- geared for use in AI scripts

'-- ## dump(string...)
'--
'-- Output text to a debug file. (3.2+ only)
'--
Public Sub wzapi_dump(ByVal context As wzapi_execution_context, strings As StringVector)
    Dim result As String, idx As Long, me_ As Long
    For idx = 0 To strings.size - 1
        If idx <> 0 Then
            result = result & " "
        End If
        result = result & strings.items(idx)
    Next
    result = result & vbLf

    me_ = context.player()
    context.currentInstance().dumpScriptLog result, me_
End Sub

'-- ## debug(string...)
'--
'-- Output text to the command line.
'--
Public Sub wzapi_debugOutputStrings(ByVal context As wzapi_execution_context, strings As StringVector)
    Dim idx As Long
    For idx = 0 To strings.size - 1
        If idx = 0 Then
            fprintf STDERR_FILE, "%s", strings.items(idx)
        Else
            fprintf STDERR_FILE, " %s", strings.items(idx)
        End If
    Next
    fprintf STDERR_FILE, "\n"
End Sub

'-- ## console(strings...)
'--
'-- Print text to the player console.
'--
' TODO, should cover scrShowConsoleText, scrAddConsoleText, scrTagConsoleText and scrConsole
Public Function wzapi_console(ByVal context As wzapi_execution_context, strings As StringVector) As Boolean
    Dim player As Long, result As String, i As Long
    player = context.player()
    If player = selectedPlayer Then
        For i = 0 To strings.size - 1
            If Len(result) <> 0 Then
                result = result & " "
            End If
            result = result & strings.items(i)
        Next
        'permitNewConsoleMessages(true);
        'setConsolePermanence(true,true);
        addConsoleMessage result, CENTRE_JUSTIFY, SYSTEM_MESSAGE
        'permitNewConsoleMessages(false);
    End If
    wzapi_console = True
End Function

'-- ## clearConsole()
'--
'-- Clear the console. (3.3+ only)
'--
Public Function wzapi_clearConsole(ByVal context As wzapi_execution_context) As Boolean
    flushConsoleMessages
    wzapi_clearConsole = True
End Function

'-- ## structureIdle(structure)
'--
'-- Is given structure idle?
'--
Public Function wzapi_structureIdle(ByVal context As wzapi_execution_context, ByVal psStruct As Long) As Boolean
    If Not (psStruct <> 0) Then debug_ 1022, LOG_ERROR, "structureIdle", "No valid structure provided": context.throwError "psStruct", 1022, "structureIdle": wzapi_structureIdle = False: Exit Function
    wzapi_structureIdle = STRUCTURE_isIdle(psStruct)
End Function

' std::vector<const STRUCTURE *> _enumStruct_fromList(WZAPI_PARAMS(optional<int> _player, optional<wzapi::STRUCTURE_TYPE_or_statsName_string> _structureType, optional<int> _playerFilter), const PerPlayerStructureLists& psStructLists)
Private Function enumStruct_fromList_(ByVal context As wzapi_execution_context, ByVal has_player As Boolean, ByVal opt_player As Long, ByVal has_structureType As Boolean, opt_structureType As STRUCTURE_TYPE_or_statsName_string, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long, psStructLists() As OBJ_LIST, result() As Long) As Long
    Dim matches As Long, statsName As String, typ As STRUCTURE_TYPE, player As Long, playerFilter As Long
    Dim it As Long, psStruct As Long, ok As Boolean
    typ = NUM_DIFF_BUILDINGS

    If has_player Then player = opt_player Else player = context.player()
    If has_playerFilter Then playerFilter = opt_playerFilter Else playerFilter = ALL_PLAYERS

    If has_structureType Then
        typ = opt_structureType.typ
        statsName = opt_structureType.statsName
    End If

    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1041, LOG_ERROR, "_enumStruct_fromList", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1041, "_enumStruct_fromList": enumStruct_fromList_ = 0: Exit Function
    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALL_PLAYERS) Then debug_ 1042, LOG_ERROR, "_enumStruct_fromList", "Player filter index out of range: %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALL_PLAYERS", 1042, "_enumStruct_fromList": enumStruct_fromList_ = 0: Exit Function
    it = psStructLists(player).head
    Do While it <> 0
        psStruct = lnObj(it)
        If playerFilter = ALL_PLAYERS Then ok = True Else ok = (aStruct(psStruct).visible(playerFilter) <> 0)
        If ok Then ok = (aStruct(psStruct).died = 0)
        If ok Then ok = (typ = NUM_DIFF_BUILDINGS Or typ = asStructureStats(aStruct(psStruct).pStructureType).typ)
        If ok Then ok = (Len(statsName) = 0 Or statsName = asStructureStats(aStruct(psStruct).pStructureType).id)
        If ok Then
            pushHandle result, matches, psStruct
        End If
        it = lnNext(it)
    Loop

    enumStruct_fromList_ = matches
End Function

'-- ## enumStruct([player[, structureType[, playerFilter]]])
'--
'-- Returns an array of structure objects. If no parameters given, it will
'-- return all of the structures for the current player. The second parameter
'-- can be either a string with the name of the structure type as defined in
'-- "structures.json", or a stattype as defined in ```Structure```. The
'-- third parameter can be used to filter by visibility, the default is not
'-- to filter.
'--
Public Function wzapi_enumStruct(ByVal context As wzapi_execution_context, ByVal has_player As Boolean, ByVal opt_player As Long, ByVal has_structureType As Boolean, opt_structureType As STRUCTURE_TYPE_or_statsName_string, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long, result() As Long) As Long
    wzapi_enumStruct = enumStruct_fromList_(context, has_player, opt_player, has_structureType, opt_structureType, has_playerFilter, opt_playerFilter, apsStructLists, result)
End Function

'-- ## enumStructOffWorld([player[, structureType[, playerFilter]]])
'--
'-- Returns an array of structure objects in your base when on an off-world mission, NULL otherwise.
'-- If no parameters given, it will return all of the structures for the current player.
'-- The second parameter can be either a string with the name of the structure type as defined
'-- in "structures.json", or a stattype as defined in ```Structure```.
'-- The third parameter can be used to filter by visibility, the default is not
'-- to filter.
'--
Public Function wzapi_enumStructOffWorld(ByVal context As wzapi_execution_context, ByVal has_player As Boolean, ByVal opt_player As Long, ByVal has_structureType As Boolean, opt_structureType As STRUCTURE_TYPE_or_statsName_string, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long, result() As Long) As Long
    wzapi_enumStructOffWorld = enumStruct_fromList_(context, has_player, opt_player, has_structureType, opt_structureType, has_playerFilter, opt_playerFilter, mission_apsStructLists, result)
End Function

'-- ## enumDroid([player[, droidType[, playerFilter]]])
'--
'-- Returns an array of droid objects. If no parameters given, it will
'-- return all of the droids for the current player. The second, optional parameter
'-- is the name of the droid type. The third parameter can be used to filter by
'-- visibility - the default is not to filter.
'--
Public Function wzapi_enumDroid(ByVal context As wzapi_execution_context, ByVal has_player As Boolean, ByVal opt_player As Long, ByVal has_droidType As Boolean, ByVal opt_droidType As Long, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long, result() As Long) As Long
    Dim matches As Long, droidType2 As DROID_TYPE, droidType As DROID_TYPE, player As Long, playerFilter As Long
    Dim it As Long, psDroid As Long, ok As Boolean

    If has_player Then player = opt_player Else player = context.player()
    If has_playerFilter Then playerFilter = opt_playerFilter Else playerFilter = ALL_PLAYERS

    If has_droidType Then droidType = opt_droidType Else droidType = DROID_ANY

    Select Case droidType ' hide some engine craziness
    Case DROID_CONSTRUCT
        droidType2 = DROID_CYBORG_CONSTRUCT
    Case DROID_WEAPON
        droidType2 = DROID_CYBORG_SUPER
    Case DROID_REPAIR
        droidType2 = DROID_CYBORG_REPAIR
    Case DROID_CYBORG
        droidType2 = DROID_CYBORG_SUPER
    Case Else
        droidType2 = droidType
    End Select
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1116, LOG_ERROR, "enumDroid", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1116, "enumDroid": wzapi_enumDroid = 0: Exit Function
    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALL_PLAYERS) Then debug_ 1117, LOG_ERROR, "enumDroid", "Player filter index out of range: %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALL_PLAYERS", 1117, "enumDroid": wzapi_enumDroid = 0: Exit Function
    it = apsDroidLists(player).head
    Do While it <> 0
        psDroid = lnObj(it)
        If playerFilter = ALL_PLAYERS Then ok = True Else ok = (aDroid(psDroid).visible(playerFilter) <> 0)
        If ok Then ok = (aDroid(psDroid).died = 0)
        If ok Then ok = (droidType = DROID_ANY Or droidType = aDroid(psDroid).droidType Or droidType2 = aDroid(psDroid).droidType)
        If ok Then
            pushHandle result, matches, psDroid
        End If
        it = lnNext(it)
    Loop
    wzapi_enumDroid = matches
End Function

'-- ## enumFeature(playerFilter[, featureName])
'--
'-- Returns an array of all features seen by player of given name, as defined in "features.json".
'-- If player is ```ALL_PLAYERS```, it will return all features irrespective of visibility to any player. If
'-- name is empty, it will return any feature.
'--
Public Function wzapi_enumFeature(ByVal context As wzapi_execution_context, ByVal playerFilter As Long, ByVal has_featureName As Boolean, ByVal opt_featureName As String, result() As Long) As Long
    Dim featureName As String, matches As Long, it As Long, psFeat As Long, ok As Boolean
    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALL_PLAYERS) Then debug_ 1138, LOG_ERROR, "enumFeature", "Player filter index out of range: %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALL_PLAYERS", 1138, "enumFeature": wzapi_enumFeature = 0: Exit Function
    If has_featureName Then
        featureName = opt_featureName
    End If

    it = apsFeatureList(0).head
    Do While it <> 0
        psFeat = lnObj(it)
        If playerFilter = ALL_PLAYERS Then ok = True Else ok = (aFeature(psFeat).visible(playerFilter) <> 0)
        If ok Then ok = (aFeature(psFeat).died = 0)
        If ok Then ok = (Len(featureName) = 0 Or featureName = asFeatureStats(aFeature(psFeat).psStats).id)
        If ok Then
            pushHandle result, matches, psFeat
        End If
        it = lnNext(it)
    Loop
    wzapi_enumFeature = matches
End Function

'-- ## enumBlips(player)
'--
'-- Return an array containing all the non-transient radar blips that the given player
'-- can see. This includes sensors revealed by radar detectors, as well as ECM jammers.
'-- It does not include units going out of view.
'--
Public Function wzapi_enumBlips(ByVal context As wzapi_execution_context, ByVal player As Long, result() As scr_position) As Long
    Dim matches As Long, it As Long, psSensor As Long, v As Long, pos As Vector3i
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1166, LOG_ERROR, "enumBlips", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1166, "enumBlips": wzapi_enumBlips = 0: Exit Function
    it = apsSensorList(0).head
    Do While it <> 0
        psSensor = lnObj(it)
        v = BASE_OBJECT_visible(psSensor, player)
        If v > 0 And v < UBYTE_MAX Then
            pos = SIMPLE_OBJECT_pos(psSensor)
            pushPosition result, matches, map_coord(pos.x), map_coord(pos.y)
        End If
        it = lnNext(it)
    Loop
    wzapi_enumBlips = matches
End Function

'-- ## enumSelected()
'--
'-- Return an array containing all game objects currently selected by the host player. (3.2+ only)
'--
Public Function wzapi_enumSelected(result() As Long) As Long
    Dim matches As Long, it As Long
    If selectedPlayer >= MAX_PLAYERS Then
        wzapi_enumSelected = matches
        Exit Function
    End If
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        If aDroid(lnObj(it)).selected <> 0 Then
            pushHandle result, matches, lnObj(it)
        End If
        it = lnNext(it)
    Loop
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        If aStruct(lnObj(it)).selected <> 0 Then
            pushHandle result, matches, lnObj(it)
        End If
        it = lnNext(it)
    Loop
    ' TODO - also add selected delivery points
    wzapi_enumSelected = matches
End Function

'-- ## enumGateways()
'--
'-- Return an array containing all the gateways on the current map. The array contains object with the properties
'-- x1, y1, x2 and y2. (3.2+ only)
'--
' (VB) GATEWAY_LIST (std::list<GATEWAY *>): the gateways copied, as gateway_c's gwGetGateway gives them
Public Function wzapi_enumGateways(ByVal context As wzapi_execution_context, result() As GATEWAY) As Long
    Dim n As Long, i As Long
    n = gwNumGateways()
    If n > 0 Then
        ReDim result(0 To n - 1)
        For i = 0 To n - 1
            gwGetGateway i, result(i)
        Next
    End If
    wzapi_enumGateways = n
End Function

'-- ## getResearch(researchName[, player])
'--
'-- Fetch information about a given technology item, given by a string that matches
'-- its definition in "research.json". If not found, returns null.
'--
Public Function wzapi_getResearch(ByVal context As wzapi_execution_context, ByVal researchName As String, ByVal has_player As Boolean, ByVal opt_player As Long) As wzapi_researchResult
    Dim result As wzapi_researchResult
    result.psResearch = getResearch(researchName)
    If has_player Then result.player = opt_player Else result.player = context.player()
    wzapi_getResearch = result
End Function

'-- ## enumResearch()
'--
'-- Returns an array of all research objects that are currently and immediately available for research.
'--
Public Function wzapi_enumResearch(ByVal context As wzapi_execution_context) As wzapi_researchResults
    Dim result As wzapi_researchResults, player As Long, i As Long, psResearch As Long
    player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1238, LOG_ERROR, "enumResearch", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1238, "enumResearch": Exit Function
    For i = 0 To asResearch_size - 1
        psResearch = i
        If Not IsResearchCompleted(asPlayerResList(player, i)) Then
            If researchAvailable(i, player, ModeQueue) Then
                pushResearch result, psResearch
            End If
        End If
    Next
    result.player = player
    wzapi_enumResearch = result
End Function

'-- ## enumRange(x, y, range[, playerFilter[, seen]])
'--
'-- Returns an array of game objects seen within range of given position that passes the optional playerFilter
'-- which can be one of a player index, ```ALL_PLAYERS```, ```ALLIES``` or ```ENEMIES```. By default, playerFilter is
'-- ```ALL_PLAYERS```. Finally an optional parameter can specify whether only visible objects should be
'-- returned; by default only visible objects are returned. Calling this function is much faster than
'-- iterating over all game objects using other enum functions. (3.2+ only)
'--
Public Function wzapi_enumRange(ByVal context As wzapi_execution_context, ByVal x_ As Long, ByVal y_ As Long, ByVal range_ As Long, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long, ByVal has_seen As Boolean, ByVal opt_seen As Boolean, result() As Long) As Long
    Dim player As Long, x As Long, y As Long, range As Long, playerFilter As Long, seen As Boolean
    Dim gridList() As Long, n As Long, gi As Long, psObj As Long, list As Long, ok As Boolean, match As Boolean
    Dim objPlayer As Long, objType As OBJECT_TYPE
    x_ = MAX_(x_, 0)
    y_ = MAX_(y_, 0)
    x_ = MIN_(x_, mapWidth)
    y_ = MIN_(y_, mapHeight)

    player = context.player()
    x = world_coord(x_)
    y = world_coord(y_)
    range = world_coord(range_)
    If has_playerFilter Then playerFilter = opt_playerFilter Else playerFilter = ALL_PLAYERS
    If has_seen Then seen = opt_seen Else seen = True

    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALL_PLAYERS Or playerFilter = ALLIES Or playerFilter = ENEMIES) Then debug_ 1273, LOG_ERROR, "enumRange", "Filter player index out of range: %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALL_PLAYERS || playerFilter == ALLIES || playerFilter == ENEMIES", 1273, "enumRange": wzapi_enumRange = 0: Exit Function

    ' static GridList gridList; (static to avoid allocations)
    n = gridStartIterate(x, y, range, gridList)
    For gi = 0 To n - 1
        psObj = gridList(gi)
        If Not seen Then
            ok = True
        Else
            ok = (player < MAX_PLAYERS)
            If ok Then ok = (BASE_OBJECT_visible(psObj, player) <> 0)
        End If
        If ok Then ok = (SIMPLE_OBJECT_died(psObj) = 0)
        If ok Then
            objPlayer = SIMPLE_OBJECT_player(psObj)
            objType = SIMPLE_OBJECT_type(psObj)
            match = False
            If playerFilter >= 0 Then match = (objPlayer = playerFilter)
            If Not match Then match = (playerFilter = ALL_PLAYERS)
            If Not match Then
                If playerFilter = ALLIES And objType <> OBJ_FEATURE Then match = aiCheckAlliances(objPlayer, player)
            End If
            If Not match Then
                If playerFilter = ENEMIES And objType <> OBJ_FEATURE Then match = Not aiCheckAlliances(objPlayer, player)
            End If
            If match Then
                pushHandle result, list, psObj
            End If
        End If
    Next
    wzapi_enumRange = list
End Function

'-- ## pursueResearch(labStructure, research)
'--
'-- Start researching the first available technology on the way to the given technology.
'-- First parameter is the structure to research in, which must be a research lab. The
'-- second parameter is the technology to pursue, as a text string as defined in "research.json".
'-- The second parameter may also be an array of such strings. The first technology that has
'-- not yet been researched in that list will be pursued.
'--
Public Function wzapi_pursueResearch(ByVal context As wzapi_execution_context, ByVal psStruct As Long, research_ As StringVector) As Boolean
    Dim player As Long, psResearch As Long, k As Long, psCurrResearch As Long, reslist() As Long, reslist_head As Long, reslist_tail As Long
    Dim curResearch As Long, iterations As Long, started As Boolean, i As Long, bits As Long, prevResearch As Long
    If Not (psStruct <> 0) Then debug_ 1304, LOG_ERROR, "pursueResearch", "No valid structure provided": context.throwError "psStruct", 1304, "pursueResearch": wzapi_pursueResearch = False: Exit Function
    player = aStruct(psStruct).player

    psResearch = NO_RESEARCH  ' Dummy initialisation.
    For k = 0 To research_.size - 1
        psCurrResearch = getResearch(research_.items(k))
        If Not (psCurrResearch <> NO_RESEARCH) Then debug_ 1311, LOG_ERROR, "pursueResearch", "No such research: %s", research_.items(k): context.throwError "psCurrResearch", 1311, "pursueResearch": wzapi_pursueResearch = False: Exit Function
        If Not IsResearchStartedPending(asPlayerResList(player, asResearch(psCurrResearch).index)) Then
            If Not IsResearchCompleted(asPlayerResList(player, asResearch(psCurrResearch).index)) Then
                ' use this one
                psResearch = psCurrResearch
                Exit For
            End If
        End If
    Next
    If psResearch = NO_RESEARCH Then
        If research_.size = 1 Then
            debug_ 1324, LOG_SCRIPT, "pursueResearch", "%s has already been researched!", research_.items(0)
        Else
            debug_ 1328, LOG_SCRIPT, "pursueResearch", "Exhausted research list -- doing nothing"
        End If
        wzapi_pursueResearch = False
        Exit Function
    End If

    If Not (asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESEARCH) Then debug_ 1333, LOG_ERROR, "pursueResearch", "Not a research lab: %s", objInfo(psStruct): context.throwError "psStruct->pStructureType->type == REF_RESEARCH", 1333, "pursueResearch": wzapi_pursueResearch = False: Exit Function
    ' RESEARCH_FACILITY *psResLab = (RESEARCH_FACILITY *)psStruct->pFunctionality;
    If Not (aStruct(psStruct).functionality.researchFacility.psSubject = NO_RESEARCH) Then debug_ 1335, LOG_ERROR, "pursueResearch", "Research lab not ready": context.throwError "psResLab->psSubject == nullptr", 1335, "pursueResearch": wzapi_pursueResearch = False: Exit Function
    ' Go down the requirements list for the desired tech
    curResearch = psResearch
    iterations = 0  ' Only used to assert we're not stuck in the loop.
    Do While curResearch <> NO_RESEARCH
        If researchAvailable(asResearch(curResearch).index, player, ModeQueue) Then
            started = False
            For i = 0 To game.maxPlayers - 1
                If i = player Then
                    bits = asPlayerResList(i, asResearch(curResearch).index).ResearchStatus
                    started = started Or (bits And STARTED_RESEARCH) <> 0 Or (bits And STARTED_RESEARCH_PENDING) <> 0 _
                              Or (bits And RESBITS_PENDING_ONLY) <> 0 Or (bits And RESEARCHED) <> 0
                ElseIf aiCheckAlliances(player, i) And alliancesSharedResearch(game.alliance) Then
                    bits = asPlayerResList(i, asResearch(curResearch).index).ResearchStatus
                    started = started Or (bits And STARTED_RESEARCH) <> 0 Or (bits And STARTED_RESEARCH_PENDING) <> 0 _
                              Or (bits And RESBITS_PENDING_ONLY) <> 0 Or (bits And RESEARCHED) <> 0
                End If
            Next
            If Not started Then ' found relevant item on the path?
                sendResearchStatus psStruct, asResearch(curResearch).index, player, True
                ' (#if defined (DEBUG): NETlogEntry of the topic started)
                debug_ 1362, LOG_SCRIPT, "pursueResearch", "Started research in %d's %s(%d) of %s", player, _
                      objInfo(psStruct), aStruct(psStruct).id, asResearch(curResearch).name_
                wzapi_pursueResearch = True
                Exit Function
            End If
        End If
        prevResearch = curResearch
        curResearch = NO_RESEARCH
        If asResearch(prevResearch).pPRList_size <> 0 Then
            curResearch = asResearch(prevResearch).pPRList(0) ' get first pre-req
        End If
        For i = 1 To asResearch(prevResearch).pPRList_size - 1
            ' push any other pre-reqs on the stack
            pushResList reslist, reslist_tail, asResearch(prevResearch).pPRList(i)
        Next
        If curResearch = NO_RESEARCH And reslist_head < reslist_tail Then
            curResearch = reslist(reslist_head) ' retrieve options from the stack
            reslist_head = reslist_head + 1
        End If
        iterations = iterations + 1
        If Not (iterations < asResearch_size * 100 Or curResearch = NO_RESEARCH) Then ASSERT_FAILURE "++iterations < asResearch.size() * 100 || !curResearch", "src/wzapi.cpp:1383", "pursueResearch", "Possible cyclic dependencies in prerequisites, possibly of research ""%s"".", asResearch(curResearch).name_: wzapi_pursueResearch = False: Exit Function
    Loop
    debug_ 1385, LOG_SCRIPT, "pursueResearch", "No research topic found for %s(%d)", objInfo(psStruct), aStruct(psStruct).id
    wzapi_pursueResearch = False ' none found
End Function

'-- ## findResearch(researchName[, player])
'--
'-- Return list of research items remaining to be researched for the given research item. (3.2+ only)
'-- (Optional second argument 3.2.3+ only)
'--
Public Function wzapi_findResearch(ByVal context As wzapi_execution_context, ByVal researchName As String, ByVal has_player As Boolean, ByVal opt_player As Long) As wzapi_researchResults
    Dim player As Long, result As wzapi_researchResults, psTarget As Long, curResearch As Long, prevResearch As Long
    Dim reslist() As Long, reslist_head As Long, reslist_tail As Long, i As Long
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1397, LOG_ERROR, "findResearch", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1397, "findResearch": Exit Function

    result.player = player

    psTarget = getResearch(researchName)
    If Not (psTarget <> NO_RESEARCH) Then debug_ 1403, LOG_ERROR, "findResearch", "No such research: %s", researchName: context.throwError "psTarget", 1403, "findResearch": Exit Function
    If IsResearchStartedPending(asPlayerResList(player, asResearch(psTarget).index)) Or IsResearchCompleted(asPlayerResList(player, asResearch(psTarget).index)) Then
        debug_ 1407, LOG_SCRIPT, "findResearch", "Find reqs for %s for player %d - research pending or completed", researchName, player
        wzapi_findResearch = result ' return empty array
        Exit Function
    End If
    debug_ 1410, LOG_SCRIPT, "findResearch", "Find reqs for %s for player %d", researchName, player
    ' Go down the requirements list for the desired tech
    curResearch = psTarget
    Do While curResearch <> NO_RESEARCH
        If (asPlayerResList(player, asResearch(curResearch).index).ResearchStatus And RESEARCHED) = 0 Then
            debug_ 1418, LOG_SCRIPT, "findResearch", "Added research in %d's %s for %s", player, asResearch(curResearch).id, asResearch(psTarget).id
            pushResearch result, curResearch
        End If
        prevResearch = curResearch
        curResearch = NO_RESEARCH
        If asResearch(prevResearch).pPRList_size <> 0 Then
            curResearch = asResearch(prevResearch).pPRList(0) ' get first pre-req
        End If
        For i = 1 To asResearch(prevResearch).pPRList_size - 1
            ' push any other pre-reqs on the stack
            pushResList reslist, reslist_tail, asResearch(prevResearch).pPRList(i)
        Next
        If curResearch = NO_RESEARCH And reslist_head < reslist_tail Then
            ' retrieve options from the stack
            curResearch = reslist(reslist_head)
            reslist_head = reslist_head + 1
        End If
    Loop
    wzapi_findResearch = result
End Function

'-- ## distBetweenTwoPoints(x1, y1, x2, y2)
'--
'-- Return distance between two points.
'--
Public Function wzapi_distBetweenTwoPoints(ByVal context As wzapi_execution_context, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    wzapi_distBetweenTwoPoints = iHypot(x1 - x2, y1 - y2)
End Function

'-- ## orderDroidLoc(droid, order, x, y)
'--
'-- Give a droid an order to do something at the given location.
'--
Public Function wzapi_orderDroidLoc(ByVal context As wzapi_execution_context, ByVal psDroid As Long, ByVal order_ As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim order As DroidOrderType
    If Not (psDroid <> 0) Then debug_ 1457, LOG_ERROR, "orderDroidLoc", "No valid droid provided": context.throwError "psDroid", 1457, "orderDroidLoc": wzapi_orderDroidLoc = False: Exit Function
    order = order_
    If Not (validOrderForLoc(order)) Then debug_ 1459, LOG_ERROR, "orderDroidLoc", "Invalid location based order: %s", getDroidOrderName(order): context.throwError "validOrderForLoc(order)", 1459, "orderDroidLoc": wzapi_orderDroidLoc = False: Exit Function
    If Not (tileOnMap(x, y)) Then debug_ 1460, LOG_ERROR, "orderDroidLoc", "Outside map bounds (%d, %d)", x, y: context.throwError "tileOnMap(x, y)", 1460, "orderDroidLoc": wzapi_orderDroidLoc = False: Exit Function
    If aDroid(psDroid).order.typ = order And aDroid(psDroid).actionPos.x = world_coord(x) And aDroid(psDroid).actionPos.y = world_coord(y) Then
        wzapi_orderDroidLoc = True
        Exit Function
    End If
    orderDroidLoc psDroid, order, world_coord(x), world_coord(y), ModeQueue
    wzapi_orderDroidLoc = True
End Function

'-- ## playerPower(player)
'--
'-- Return amount of power held by the given player.
'--
Public Function wzapi_playerPower(ByVal context As wzapi_execution_context, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1476, LOG_ERROR, "playerPower", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1476, "playerPower": wzapi_playerPower = -1: Exit Function
    wzapi_playerPower = getPower(player)
End Function

'-- ## queuedPower(player)
'--
'-- Return amount of power queued up for production by the given player. (3.2+ only)
'--
Public Function wzapi_queuedPower(ByVal context As wzapi_execution_context, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1486, LOG_ERROR, "queuedPower", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1486, "queuedPower": wzapi_queuedPower = -1: Exit Function
    wzapi_queuedPower = getQueuedPower(player)
End Function

'-- ## isStructureAvailable(structureName[, player])
'--
'-- Returns true if given structure can be built. It checks both research and unit limits.
'--
Public Function wzapi_isStructureAvailable(ByVal context As wzapi_execution_context, ByVal structureName As String, ByVal has_player As Boolean, ByVal opt_player As Long) As Boolean
    Dim structureIndex As Long, player As Long, status As Long
    structureIndex = getStructStatFromName(structureName)
    If Not (structureIndex >= 0 And structureIndex < numStructureStats) Then debug_ 1497, LOG_ERROR, "isStructureAvailable", "Structure %s not found", structureName: context.throwError "structureIndex >= 0 && structureIndex < numStructureStats", 1497, "isStructureAvailable": wzapi_isStructureAvailable = False: Exit Function
    If has_player Then player = opt_player Else player = context.player()

    status = apStructTypeLists(player, structureIndex)
    wzapi_isStructureAvailable = (status = AVAILABLE Or status = REDUNDANT) _
        And asStructureStats(structureIndex).curCount(player) < asStructureStats(structureIndex).upgrade(player).limit
End Function

' additional structure check
' (VB) psStat: the STRUCTURE_STATS * (C: the BASE_STATS * cast back)
Private Function structDoubleCheck(ByVal psStat As Long, ByVal xx As Long, ByVal yy As Long, ByVal maxBlockingTiles As Long, ByVal propType As PROPULSION_TYPE) As Boolean
    Dim x As Long, y As Long, xTL As Long, yTL As Long, xBR As Long, yBR As Long
    Dim count As Long, psBuilding As Long, g As Long, psGate As GATEWAY
    count = 0
    psBuilding = castStructureStats(psStat)

    xTL = xx - 1
    yTL = yy - 1
    xBR = (xx + asStructureStats(psBuilding).baseWidth)
    yBR = (yy + asStructureStats(psBuilding).baseBreadth)

    ' check against building in a gateway, as this can seriously block AI passages
    For g = 0 To gwNumGateways() - 1
        gwGetGateway g, psGate
        For x = xx To xBR
            For y = yy To yBR
                If x >= psGate.x1 And x <= psGate.x2 And y >= psGate.y1 And y <= psGate.y2 Then
                    structDoubleCheck = False
                    Exit Function
                End If
            Next
        Next
    Next

    ' can you get past it?
    y = yTL ' top
    x = xTL
    Do While x <> xBR + 1
        If fpathBlockingTile(x, y, propType) Then
            count = count + 1
            Exit Do
        End If
        x = x + 1
    Loop

    y = yBR ' bottom
    x = xTL
    Do While x <> xBR + 1
        If fpathBlockingTile(x, y, propType) Then
            count = count + 1
            Exit Do
        End If
        x = x + 1
    Loop

    x = xTL ' left
    y = yTL + 1
    Do While y <> yBR
        If fpathBlockingTile(x, y, propType) Then
            count = count + 1
            Exit Do
        End If
        y = y + 1
    Loop

    x = xBR ' right
    y = yTL + 1
    Do While y <> yBR
        If fpathBlockingTile(x, y, propType) Then
            count = count + 1
            Exit Do
        End If
        y = y + 1
    Loop

    'make sure this location is not blocked from too many sides
    If count <= maxBlockingTiles Or maxBlockingTiles = -1 Then
        structDoubleCheck = True
        Exit Function
    End If
    structDoubleCheck = False
End Function

' save a lot of typing... checks whether a position is valid
' (VB) pickStructLocation's macro LOC_OK, with what it captures; its && in order
Private Function LOC_OK(ByVal x As Long, ByVal y As Long, ByVal psDroid As Long, ByVal psStat As Long, offset As Vector2i, ByVal player As Long, ByVal maxBlockingTiles As Long, ByVal propType As PROPULSION_TYPE) As Boolean
    Dim dest As Vector3i, pos As Vector2i
    If Not tileOnMap(x, y) Then Exit Function
    If psDroid <> 0 Then
        dest.x = world_coord(x)
        dest.y = world_coord(y)
        dest.z = 0
        If Not fpathCheck(aDroid(psDroid).pos, dest, propType) Then Exit Function
    End If
    pos.x = world_coord(x) + offset.x
    pos.y = world_coord(y) + offset.y
    ' (VB) psStat is the macro's STRUCTURE_STATS * (an index); validLocation and structDoubleCheck take the
    ' BASE_STATS * it converts to: the handle STAT_STRUCTURE + index
    If Not validLocation(STAT_STRUCTURE + psStat, pos, 0, player, False) Then Exit Function
    LOC_OK = structDoubleCheck(STAT_STRUCTURE + psStat, x, y, maxBlockingTiles, propType)
End Function

'-- ## pickStructLocation(droid, structureName, x, y[, maxBlockingTiles])
'--
'-- Pick a location for constructing a certain type of building near some given position.
'-- Returns an object containing "type" ```POSITION```, and "x" and "y" values, if successful.
'--
Public Function wzapi_pickStructLocation(ByVal context As wzapi_execution_context, ByVal psDroid As Long, ByVal structureName As String, ByVal startX As Long, ByVal startY As Long, ByVal has_maxBlockingTiles As Boolean, ByVal opt_maxBlockingTiles As Long, result As scr_position) As Boolean
    Dim player As Long, structureIndex As Long, psStat As Long, numIterations As Long, found_ As Boolean
    Dim incX As Long, incY As Long, x As Long, y As Long, maxBlockingTiles As Long, offset As Vector2i, propType As PROPULSION_TYPE
    If Not (psDroid <> 0) Then debug_ 1588, LOG_ERROR, "pickStructLocation", "No valid droid provided": context.throwError "psDroid", 1588, "pickStructLocation": wzapi_pickStructLocation = False: Exit Function
    player = aDroid(psDroid).player
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1590, LOG_ERROR, "pickStructLocation", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1590, "pickStructLocation": wzapi_pickStructLocation = False: Exit Function
    structureIndex = getStructStatFromName(structureName)
    If Not (structureIndex >= 0 And structureIndex < numStructureStats) Then debug_ 1592, LOG_ERROR, "pickStructLocation", "Structure %s not found", structureName: context.throwError "structureIndex >= 0 && structureIndex < numStructureStats", 1592, "pickStructLocation": wzapi_pickStructLocation = False: Exit Function
    psStat = structureIndex
    If Not (psStat <> -1) Then debug_ 1594, LOG_ERROR, "pickStructLocation", "No such stat found: %s", structureName: context.throwError "psStat", 1594, "pickStructLocation": wzapi_pickStructLocation = False: Exit Function

    numIterations = 30
    found_ = False
    If has_maxBlockingTiles Then maxBlockingTiles = opt_maxBlockingTiles Else maxBlockingTiles = 0

    If Not (startX >= 0 And startX < mapWidth And startY >= 0 And startY < mapHeight) Then debug_ 1601, LOG_ERROR, "pickStructLocation", "Bad position (%d, %d)", startX, startY: context.throwError "startX >= 0 && startX < mapWidth && startY >= 0 && startY < mapHeight", 1601, "pickStructLocation": wzapi_pickStructLocation = False: Exit Function

    x = startX
    y = startY

    offset.x = asStructureStats(psStat).baseWidth * (TILE_UNITS \ 2)
    offset.y = asStructureStats(psStat).baseBreadth * (TILE_UNITS \ 2)

    If psDroid <> 0 Then propType = asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType Else propType = PROPULSION_TYPE_WHEELED

    ' first try the original location
    If LOC_OK(startX, startY, psDroid, psStat, offset, player, maxBlockingTiles, propType) Then
        found_ = True
    End If

    ' try some locations nearby
    incX = 1
    incY = 1
    Do While incX < numIterations And Not found_
        y = startY - incY ' top
        For x = startX - incX To startX + incX - 1
            If LOC_OK(x, y, psDroid, psStat, offset, player, maxBlockingTiles, propType) Then
                found_ = True
                GoTo endstructloc
            End If
        Next
        x = startX + incX ' right
        For y = startY - incY To startY + incY - 1
            If LOC_OK(x, y, psDroid, psStat, offset, player, maxBlockingTiles, propType) Then
                found_ = True
                GoTo endstructloc
            End If
        Next
        y = startY + incY ' bottom
        For x = startX + incX To startX - incX + 1 Step -1
            If LOC_OK(x, y, psDroid, psStat, offset, player, maxBlockingTiles, propType) Then
                found_ = True
                GoTo endstructloc
            End If
        Next
        x = startX - incX ' left
        For y = startY + incY To startY - incY + 1 Step -1
            If LOC_OK(x, y, psDroid, psStat, offset, player, maxBlockingTiles, propType) Then
                found_ = True
                GoTo endstructloc
            End If
        Next
        incX = incX + 1
        incY = incY + 1
    Loop

endstructloc:
    If found_ Then
        result.x = x + map_coord(offset.x)
        result.y = y + map_coord(offset.y)
        wzapi_pickStructLocation = True
        Exit Function
    Else
        debug_ 1669, LOG_SCRIPT, "pickStructLocation", "Did not find valid positioning for %s", asStructureStats(psStat).name_
    End If
    wzapi_pickStructLocation = False
End Function

'-- ## structureCanFit(structureName, x, y[, direction])
'--
'-- Return true if given building can be built at the position. (4.6+ only)
'--
Public Function wzapi_structureCanFit(ByVal context As wzapi_execution_context, ByVal structureName As String, ByVal x As Long, ByVal y As Long, ByVal has_direction As Boolean, ByVal opt_direction As Single) As Boolean
    Dim player As Long, structureIndex As Long, psStat As Long, direction As Long, directionValue As Single, pos As Vector2i
    player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1681, LOG_ERROR, "structureCanFit", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1681, "structureCanFit": wzapi_structureCanFit = False: Exit Function
    structureIndex = getStructStatFromName(structureName)
    If Not (structureIndex >= 0 And structureIndex < numStructureStats) Then debug_ 1683, LOG_ERROR, "structureCanFit", "Structure %s not found", structureName: context.throwError "structureIndex >= 0 && structureIndex < numStructureStats", 1683, "structureCanFit": wzapi_structureCanFit = False: Exit Function
    psStat = structureIndex
    If Not (psStat <> -1) Then debug_ 1685, LOG_ERROR, "structureCanFit", "No such stat found: %s", structureName: context.throwError "psStat", 1685, "structureCanFit": wzapi_structureCanFit = False: Exit Function

    If has_direction Then
        If IsNaNF(opt_direction) Then
            opt_direction = 0! ' avoid undefined behavior (nan is outside the range of representable values of type 'unsigned short')
        End If
    End If
    If has_direction Then directionValue = opt_direction Else directionValue = 0!
    direction = ToInt(fdiv(fmul(directionValue, 8192!), 45!)) And &HFFFF&    ' static_cast<uint16_t>(DEG(_direction.value_or(0)))

    If Not tileOnMap(x, y) Then
        wzapi_structureCanFit = False
        Exit Function
    End If
    pos.x = world_coord(x)
    pos.y = world_coord(y)
    wzapi_structureCanFit = validLocation(STAT_STRUCTURE + psStat, pos, direction, player, False) ' (VB) the BASE_STATS * handle
End Function

'-- ## droidCanReach(droid, x, y)
'--
'-- Return whether or not the given droid could possibly drive to the given position. Does
'-- not take player built blockades into account.
'--
Public Function wzapi_droidCanReach(ByVal context As wzapi_execution_context, ByVal psDroid As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim psPropStats As Long, dest As Vector3i
    If Not (psDroid <> 0) Then debug_ 1704, LOG_ERROR, "droidCanReach", "No valid droid provided": context.throwError "psDroid", 1704, "droidCanReach": wzapi_droidCanReach = False: Exit Function
    psPropStats = DROID_getPropulsionStats(psDroid)
    dest.x = world_coord(x)
    dest.y = world_coord(y)
    dest.z = 0
    wzapi_droidCanReach = fpathCheck(aDroid(psDroid).pos, dest, asPropulsionStats(psPropStats).propulsionType)
End Function

'-- ## propulsionCanReach(propulsionName, x1, y1, x2, y2)
'--
'-- Return true if a droid with a given propulsion is able to travel from (x1, y1) to (x2, y2).
'-- Does not take player built blockades into account. (3.2+ only)
'--
Public Function wzapi_propulsionCanReach(ByVal context As wzapi_execution_context, ByVal propulsionName As String, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    Dim propulsionIndex As Long, psPropStats As Long, orig As Vector3i, dest As Vector3i
    propulsionIndex = getCompFromName(COMP_PROPULSION, propulsionName)
    If Not (propulsionIndex > 0) Then debug_ 1717, LOG_ERROR, "propulsionCanReach", "No such propulsion: %s", propulsionName: context.throwError "propulsionIndex > 0", 1717, "propulsionCanReach": wzapi_propulsionCanReach = False: Exit Function
    psPropStats = propulsionIndex
    orig.x = world_coord(x1)
    orig.y = world_coord(y1)
    orig.z = 0
    dest.x = world_coord(x2)
    dest.y = world_coord(y2)
    dest.z = 0
    wzapi_propulsionCanReach = fpathCheck(orig, dest, asPropulsionStats(psPropStats).propulsionType)
End Function

'-- ## terrainType(x, y)
'--
'-- Returns tile type of a given map tile, such as ```TER_WATER``` for water tiles or ```TER_CLIFFFACE``` for cliffs.
'-- Tile types regulate which units may pass through this tile. (3.2+ only)
'--
Public Function wzapi_terrainType(ByVal context As wzapi_execution_context, ByVal x As Long, ByVal y As Long) As Long
    wzapi_terrainType = terrainType(mapTile(x, y))
End Function

'-- ## tileIsBurning(x, y)
'--
'-- Returns whether the given map tile is burning. (3.5+ only)
'--
Public Function wzapi_tileIsBurning(ByVal context As wzapi_execution_context, ByVal x As Long, ByVal y As Long) As Boolean
    Dim psTile As Long
    psTile = mapTile(x, y)
    ' SCRIPT_ASSERT(false, context, psTile, "Checking fire on tile outside the map (%d, %d)", x, y): mapTile
    ' clamps its coordinates and never gives nullptr (here a tile index, 0 a valid one)
    wzapi_tileIsBurning = TileIsBurning(psTile)
End Function

'-- ## orderDroidObj(droid, order, object)
'--
'-- Give a droid an order to do something to something.
'--
Public Function wzapi_orderDroidObj(ByVal context As wzapi_execution_context, ByVal psDroid As Long, ByVal order_ As Long, ByVal psObj As Long) As Boolean
    Dim order As DroidOrderType
    If Not (psDroid <> 0) Then debug_ 1749, LOG_ERROR, "orderDroidObj", "No valid droid provided": context.throwError "psDroid", 1749, "orderDroidObj": wzapi_orderDroidObj = False: Exit Function
    order = order_
    If Not (psObj <> 0) Then debug_ 1751, LOG_ERROR, "orderDroidObj", "No valid object provided": context.throwError "psObj", 1751, "orderDroidObj": wzapi_orderDroidObj = False: Exit Function
    If Not (validOrderForObj(order)) Then debug_ 1752, LOG_ERROR, "orderDroidObj", "Invalid order: %s", getDroidOrderName(order): context.throwError "validOrderForObj(order)", 1752, "orderDroidObj": wzapi_orderDroidObj = False: Exit Function
    If aDroid(psDroid).order.typ = order And aDroid(psDroid).order.psObj = psObj Then
        wzapi_orderDroidObj = True
        Exit Function
    End If
    orderDroidObj psDroid, order, psObj, ModeQueue
    wzapi_orderDroidObj = True
End Function

Private Function get_first_available_component(ByVal player As Long, ByVal capacity As Long, list As StringVector, ByVal componentType As COMPONENT_TYPE, ByVal strict As Boolean) As Long
    Dim k As Long, componentIndex As Long, status As Long, ok As Boolean
    For k = 0 To list.size - 1
        componentIndex = getCompFromName(componentType, list.items(k))
        If componentIndex >= 0 Then
            status = apCompLists(player, componentType, componentIndex)
            ok = (status = AVAILABLE Or status = REDUNDANT Or Not strict)
            If ok Then
                If componentType = COMP_BODY Then ok = (asBodyStats(componentIndex).size <= capacity)
            End If
            If ok Then
                get_first_available_component = componentIndex ' found one!
                Exit Function
            End If
        End If
        If componentIndex < 0 Then
            debug_ 1778, LOG_ERROR, "get_first_available_component", "No such component: %s", list.items(k)
        End If
    Next
    get_first_available_component = -1 ' no available component found in list
End Function

' static std::unique_ptr<DROID_TEMPLATE> makeTemplate(...): a template handle of template_c's pool, 0 = nullptr
' (the unique_ptr's delete is deleteTemplate)
Private Function makeTemplate_(ByVal player As Long, ByVal templateName As String, body_ As StringVector, propulsion_ As StringVector, turrets_() As StringVector, ByVal turrets_size As Long, ByVal capacity As Long, ByVal strict As Boolean) As Long
    Dim psTemplate As Long, numTurrets As Long, result As Long, body As Long, prop As Long, i As Long
    Dim componentName As String, psComp As Long, valid As Boolean
    psTemplate = newTemplate()
    numTurrets = turrets_size

    For i = 0 To DROID_MAXCOMP - 1  ' reset to defaults
        apsTemplates(psTemplate).asParts(i) = 0
    Next
    For i = 0 To MAX_WEAPONS - 1
        apsTemplates(psTemplate).asWeaps(i) = 0
    Next
    body = get_first_available_component(player, capacity, body_, COMP_BODY, strict)
    If body < 0 Then
        debug_ 1795, LOG_SCRIPT, "makeTemplate", "Wanted to build %s but body types all unavailable", _
              templateName
        deleteTemplate psTemplate
        makeTemplate_ = 0 ' no component available
        Exit Function
    End If
    prop = get_first_available_component(player, capacity, propulsion_, COMP_PROPULSION, strict)
    If prop < 0 Then
        debug_ 1802, LOG_SCRIPT, "makeTemplate", "Wanted to build %s but propulsion types all unavailable", _
              templateName
        deleteTemplate psTemplate
        makeTemplate_ = 0 ' no component available
        Exit Function
    End If
    apsTemplates(psTemplate).asParts(COMP_BODY) = body
    apsTemplates(psTemplate).asParts(COMP_PROPULSION) = prop

    apsTemplates(psTemplate).numWeaps = 0
    numTurrets = MIN_(numTurrets, asBodyStats(body).weaponSlots) ' Restrict max no. turrets
    If asBodyStats(body).droidTypeOverride <> DROID_ANY Then
        apsTemplates(psTemplate).droidType = asBodyStats(body).droidTypeOverride ' set droidType based on body
    End If
    ' Find first turret component type (assume every component in list is same type)
    If turrets_size = 0 Then valid = False Else valid = (turrets_(0).size <> 0)
    If Not valid Then
        debug_ 1818, LOG_SCRIPT, "makeTemplate", "Wanted to build %s but no turrets provided", templateName
        deleteTemplate psTemplate
        makeTemplate_ = 0
        Exit Function
    End If
    componentName = turrets_(0).items(0)
    psComp = getCompStatsFromName(componentName)
    If psComp = 0 Then
        debug_ 1825, LOG_ERROR, "makeTemplate", "Wanted to build %s but %s does not exist", templateName, componentName
        deleteTemplate psTemplate
        makeTemplate_ = 0
        Exit Function
    End If
    If COMPONENT_STATS_droidTypeOverride(psComp) <> DROID_ANY Then
        apsTemplates(psTemplate).droidType = COMPONENT_STATS_droidTypeOverride(psComp) ' set droidType based on component
    End If
    If COMPONENT_STATS_compType(psComp) = COMP_WEAPON Then
        For i = 0 To MIN_(numTurrets, MAX_WEAPONS) - 1 ' may be multi-weapon
            result = get_first_available_component(player, SIZE_NUM, turrets_(i), COMP_WEAPON, strict)
            If result < 0 Then
                debug_ 1839, LOG_SCRIPT, "makeTemplate", "Wanted to build %s but no weapon available", templateName
                deleteTemplate psTemplate
                makeTemplate_ = 0
                Exit Function
            End If
            apsTemplates(psTemplate).asWeaps(i) = result
            apsTemplates(psTemplate).numWeaps = apsTemplates(psTemplate).numWeaps + 1
        Next
    Else
        If COMPONENT_STATS_compType(psComp) = COMP_BRAIN Then
            apsTemplates(psTemplate).numWeaps = 1 ' hack, necessary to pass intValidTemplate
        End If
        result = get_first_available_component(player, SIZE_NUM, turrets_(0), COMPONENT_STATS_compType(psComp), strict)
        If result < 0 Then
            debug_ 1855, LOG_SCRIPT, "makeTemplate", "Wanted to build %s but turret unavailable", templateName
            deleteTemplate psTemplate
            makeTemplate_ = 0
            Exit Function
        End If
        apsTemplates(psTemplate).asParts(COMPONENT_STATS_compType(psComp)) = result
    End If
    valid = intValidTemplate(apsTemplates(psTemplate), templateName, True, player)
    If valid Then
        makeTemplate_ = psTemplate
    Else
        debug_ 1867, LOG_ERROR, "makeTemplate", "Invalid template %s", templateName
        deleteTemplate psTemplate
        makeTemplate_ = 0
    End If
End Function

'-- ## buildDroid(factory, templateName, body, propulsion, reserved, reserved, turrets...)
'--
'-- Start factory production of new droid with the given name, body, propulsion and turrets.
'-- The reserved parameter should be passed **null** for now. The components can be
'-- passed as ordinary strings, or as a list of strings. If passed as a list, the first available
'-- component in the list will be used. The second reserved parameter used to be a droid type.
'-- It is now unused and in 3.2+ should be passed "", while in 3.1 it should be the
'-- droid type to be built. Returns a boolean that is true if production was started.
'--
Public Function wzapi_buildDroid(ByVal context As wzapi_execution_context, ByVal psFactory As Long, ByVal templateName As String, body As StringVector, propulsion As StringVector, ByVal reserved1 As Long, ByVal reserved2 As Long, turrets() As StringVector, ByVal turrets_size As Long) As Boolean
    Dim psStruct As Long, t As STRUCTURE_TYPE, player As Long, capacity As Long, psTemplate As Long, k As Long, templ As Long
    Dim psAddedTemplate As Long, ok As Boolean
    psStruct = psFactory
    If Not (psStruct <> 0) Then debug_ 1884, LOG_ERROR, "buildDroid", "No valid structure provided": context.throwError "psStruct", 1884, "buildDroid": wzapi_buildDroid = False: Exit Function
    t = asStructureStats(aStruct(psStruct).pStructureType).typ
    If Not (t = REF_FACTORY Or t = REF_CYBORG_FACTORY Or t = REF_VTOL_FACTORY) Then
        debug_ 1885, LOG_ERROR, "buildDroid", "Structure %s is not a factory", objInfo(psStruct)
        context.throwError "(psStruct->pStructureType->type == REF_FACTORY || psStruct->pStructureType->type == REF_CYBORG_FACTORY || psStruct->pStructureType->type == REF_VTOL_FACTORY)", 1885, "buildDroid"
        wzapi_buildDroid = False
        Exit Function
    End If
    player = aStruct(psStruct).player
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1888, LOG_ERROR, "buildDroid", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1888, "buildDroid": wzapi_buildDroid = False: Exit Function
    capacity = aStruct(psStruct).capacity ' body size limit
    If turrets_size <> 0 Then ok = (turrets(0).size <> 0) Else ok = False
    If Not ok Then debug_ 1890, LOG_ERROR, "buildDroid", "No turrets provided": context.throwError "!turrets.va_list.empty() && !turrets.va_list[0].strings.empty()", 1890, "buildDroid": wzapi_buildDroid = False: Exit Function
    psTemplate = makeTemplate_(player, templateName, body, propulsion, turrets, turrets_size, capacity, True)
    If psTemplate <> 0 Then
        If Not (validTemplateForFactory(psTemplate, psStruct, True)) Then
            debug_ 1894, LOG_ERROR, "buildDroid", "Invalid template %s for factory %s", _
                  apsTemplates(psTemplate).name_, asStructureStats(aStruct(psStruct).pStructureType).name_
            context.throwError "validTemplateForFactory(psTemplate.get(), psStruct, true)", 1894, "buildDroid"
            deleteTemplate psTemplate      ' (the unique_ptr's delete)
            wzapi_buildDroid = False
            Exit Function
        End If
        ' Delete similar template from existing list before adding this one
        For k = 0 To apsTemplateList_size - 1
            templ = apsTemplateList(k)
            If apsTemplates(templ).name_ = apsTemplates(psTemplate).name_ Then
                debug_ 1902, LOG_SCRIPT, "buildDroid", "deleting %s for player %d", apsTemplates(templ).name_, player
                deleteTemplateFromProduction templ, player, ModeQueue ' duplicate? done below?
                Exit For
            End If
        Next
        ' Add to list
        debug_ 1908, LOG_SCRIPT, "buildDroid", "adding template %s for player %d", apsTemplates(psTemplate).name_, player
        apsTemplates(psTemplate).multiPlayerID = generateNewObjectId()
        psAddedTemplate = addTemplate(player, psTemplate)
        If Not structSetManufacture(psStruct, psAddedTemplate, ModeQueue) Then
            debug_ 1913, LOG_ERROR, "buildDroid", "Could not produce template %s in %s", apsTemplates(psAddedTemplate).name_, objInfo(psStruct)
            wzapi_buildDroid = False
            Exit Function
        End If
        wzapi_buildDroid = True
        Exit Function
    End If
    wzapi_buildDroid = False
End Function

'-- ## addDroid(player, x, y, templateName, body, propulsion, reserved, reserved, turrets...)
'--
'-- Create and place a droid at the given x, y position as belonging to the given player, built with
'-- the given components. Currently does not support placing droids in multiplayer, doing so will
'-- cause a desync. Returns the created droid on success, otherwise returns null. Passing "" for
'-- reserved parameters is recommended. In 3.2+ only, to create droids in off-world (campaign mission list),
'-- pass -1 as both x and y.
'--
Public Function wzapi_addDroid(ByVal context As wzapi_execution_context, ByVal player As Long, ByVal x As Long, ByVal y As Long, ByVal templateName As String, body As StringVector, propulsion As StringVector, ByVal reserved1 As Long, ByVal reserved2 As Long, turrets() As StringVector, ByVal turrets_size As Long) As Long
    Dim onMission As Boolean, psTemplate As Long, psDroid As Long, oldMulti As Boolean, ok As Boolean
    Dim initialOrders As INITIAL_DROID_ORDERS, rot As Rotation
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1931, LOG_ERROR, "addDroid", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1931, "addDroid": wzapi_addDroid = 0: Exit Function
    onMission = (x = -1) And (y = -1)
    If Not ((onMission Or (x >= 0 And y >= 0))) Then debug_ 1933, LOG_ERROR, "addDroid", "Invalid coordinates (%d, %d) for droid", x, y: context.throwError "(onMission || (x >= 0 && y >= 0))", 1933, "addDroid": wzapi_addDroid = 0: Exit Function
    If turrets_size <> 0 Then ok = (turrets(0).size <> 0) Else ok = False
    If Not ok Then debug_ 1934, LOG_ERROR, "addDroid", "No turrets provided": context.throwError "!turrets.va_list.empty() && !turrets.va_list[0].strings.empty()", 1934, "addDroid": wzapi_addDroid = 0: Exit Function
    psTemplate = makeTemplate_(player, templateName, body, propulsion, turrets, turrets_size, SIZE_NUM, False)
    If psTemplate <> 0 Then
        psDroid = 0
        oldMulti = bMultiMessages
        bMultiMessages = False ' ugh, fixme
        If onMission Then
            psDroid = buildMissionDroid(psTemplate, 128, 128, player)
            If psDroid <> 0 Then
                debug_ 1946, LOG_LIFE, "addDroid", "Created mission-list droid %s by script for player %d: %u", objInfo(psDroid), player, aDroid(psDroid).id
            Else
                debug_ 1950, LOG_ERROR, "addDroid", "Invalid droid %s", templateName
            End If
        Else
            psDroid = buildDroid(psTemplate, world_coord(x) + TILE_UNITS \ 2, world_coord(y) + TILE_UNITS \ 2, player, onMission, False, initialOrders, rot)
            If psDroid <> 0 Then
                addDroid psDroid, apsDroidLists
                debug_ 1959, LOG_LIFE, "addDroid", "Created droid %s by script for player %d: %u", objInfo(psDroid), player, aDroid(psDroid).id
            Else
                debug_ 1963, LOG_ERROR, "addDroid", "Invalid droid %s", templateName
            End If
        End If
        bMultiMessages = oldMulti ' ugh
        deleteTemplate psTemplate          ' (the unique_ptr's delete)
        wzapi_addDroid = psDroid
        Exit Function
    End If
    wzapi_addDroid = 0
End Function

'-- ## makeTemplate(player, templateName, body, propulsion, reserved, turrets...)
'--
'-- Create a template (virtual droid) with the given components. Can be useful for calculating the cost
'-- of droids before putting them into production, for instance. Will fail and return null if template
'-- could not possibly be built using current research. (3.2+ only)
'--
' (VB) std::unique_ptr<const DROID_TEMPLATE>: a template handle (0 = nullptr) the caller frees with deleteTemplate
Public Function wzapi_makeTemplate(ByVal context As wzapi_execution_context, ByVal player As Long, ByVal templateName As String, body As StringVector, propulsion As StringVector, ByVal reserved1 As Long, turrets() As StringVector, ByVal turrets_size As Long) As Long
    Dim ok As Boolean
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 1980, LOG_ERROR, "makeTemplate", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 1980, "makeTemplate": wzapi_makeTemplate = 0: Exit Function
    If turrets_size <> 0 Then ok = (turrets(0).size <> 0) Else ok = False
    If Not ok Then debug_ 1981, LOG_ERROR, "makeTemplate", "No turrets provided": context.throwError "!turrets.va_list.empty() && !turrets.va_list[0].strings.empty()", 1981, "makeTemplate": wzapi_makeTemplate = 0: Exit Function
    wzapi_makeTemplate = makeTemplate_(player, templateName, body, propulsion, turrets, turrets_size, SIZE_NUM, True)
End Function

'-- ## addDroidToTransporter(transporter, droid)
'--
'-- Load a droid, which is currently located on the campaign off-world mission list,
'-- into a transporter, which is also currently on the campaign off-world mission list.
'-- (3.2+ only)
'--
Public Function wzapi_addDroidToTransporter(ByVal context As wzapi_execution_context, transporter As game_object_identifier, droid_ As game_object_identifier) As Boolean
    Dim transporterId As Long, transporterPlayer As Long, psTransporter As Long, droidId As Long, droidPlayer As Long
    Dim psDroid As Long, removeSuccessful As Boolean
    transporterId = transporter.id
    transporterPlayer = transporter.player
    psTransporter = IdToMissionDroid(transporterId, transporterPlayer)
    If Not (psTransporter <> 0) Then debug_ 1997, LOG_ERROR, "addDroidToTransporter", "No such transporter id %d belonging to player %d", transporterId, transporterPlayer: context.throwError "psTransporter", 1997, "addDroidToTransporter": wzapi_addDroidToTransporter = False: Exit Function
    If Not (DROID_isTransporter(psTransporter)) Then debug_ 1998, LOG_ERROR, "addDroidToTransporter", "Droid id %d belonging to player %d is not a transporter", transporterId, transporterPlayer: context.throwError "psTransporter->isTransporter()", 1998, "addDroidToTransporter": wzapi_addDroidToTransporter = False: Exit Function
    droidId = droid_.id
    droidPlayer = droid_.player
    psDroid = IdToMissionDroid(droidId, droidPlayer)
    If Not (psDroid <> 0) Then debug_ 2002, LOG_ERROR, "addDroidToTransporter", "No such droid id %d belonging to player %d", droidId, droidPlayer: context.throwError "psDroid", 2002, "addDroidToTransporter": wzapi_addDroidToTransporter = False: Exit Function
    If Not (checkTransporterSpace(psTransporter, psDroid)) Then debug_ 2003, LOG_ERROR, "addDroidToTransporter", "Not enough room in transporter %d for droid %d", transporterId, droidId: context.throwError "checkTransporterSpace(psTransporter, psDroid)", 2003, "addDroidToTransporter": wzapi_addDroidToTransporter = False: Exit Function
    removeSuccessful = droidRemove(psDroid, mission_apsDroidLists)
    If Not (removeSuccessful) Then debug_ 2005, LOG_ERROR, "addDroidToTransporter", "Could not remove droid id %d from mission list", droidId: context.throwError "removeSuccessful", 2005, "addDroidToTransporter": wzapi_addDroidToTransporter = False: Exit Function
    DROID_GROUP_add aDroid(psTransporter).psGroup, psDroid
    wzapi_addDroidToTransporter = True
End Function

'-- ## addFeature(featureName, x, y)
'--
'-- Create and place a feature at the given x, y position. Will cause a desync in multiplayer.
'-- Returns the created game object on success, null otherwise. (3.2+ only)
'--
Public Function wzapi_addFeature(ByVal context As wzapi_execution_context, ByVal featureName As String, ByVal x As Long, ByVal y As Long) As Long
    Dim feature_ As Long, psStats As Long, it As Long, psFeat As Long, psFeature As Long
    feature_ = getFeatureStatFromName(featureName)
    If Not (feature_ >= 0 And feature_ < asFeatureStats_size) Then debug_ 2018, LOG_ERROR, "addFeature", "Unknown feature name: %s", featureName: context.throwError "feature >= 0 && feature < asFeatureStats.size()", 2018, "addFeature": wzapi_addFeature = 0: Exit Function
    psStats = feature_
    it = apsFeatureList(0).head
    Do While it <> 0
        psFeat = lnObj(it)
        If Not (map_coord(aFeature(psFeat).pos.x) <> x Or map_coord(aFeature(psFeat).pos.y) <> y) Then
            debug_ 2022, LOG_ERROR, "addFeature", "Building feature on tile already occupied"
            context.throwError "map_coord(psFeat->pos.x) != x || map_coord(psFeat->pos.y) != y", 2022, "addFeature"
            wzapi_addFeature = 0
            Exit Function
        End If
        it = lnNext(it)
    Loop
    psFeature = buildFeature(psStats, world_coord(x), world_coord(y), False)
    wzapi_addFeature = psFeature
End Function

'-- ## componentAvailable([componentType, ]componentName)
'--
'-- Checks whether a given component is available to the current player. The first argument is
'-- optional and deprecated.
'--
Public Function wzapi_componentAvailable(ByVal context As wzapi_execution_context, ByVal componentType As String, ByVal has_componentName As Boolean, ByVal opt_componentName As String) As Boolean
    Dim player As Long, componentName As String, psComp As Long, status As Long
    player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2037, LOG_ERROR, "componentAvailable", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2037, "componentAvailable": wzapi_componentAvailable = False: Exit Function
    If has_componentName Then componentName = opt_componentName Else componentName = componentType
    psComp = getCompStatsFromName(componentName)
    If Not (psComp <> 0) Then debug_ 2040, LOG_ERROR, "componentAvailable", "No such component: %s", componentName: context.throwError "psComp", 2040, "componentAvailable": wzapi_componentAvailable = False: Exit Function
    status = apCompLists(player, COMPONENT_STATS_compType(psComp), BASE_STATS_index(psComp))
    wzapi_componentAvailable = (status = AVAILABLE Or status = REDUNDANT)
End Function

'-- ## isVTOL(droid)
'--
'-- Returns true if given droid is a VTOL (not including transports).
'--
Public Function wzapi_isVTOL(ByVal context As wzapi_execution_context, ByVal psDroid As Long) As Boolean
    If Not (psDroid <> 0) Then debug_ 2051, LOG_ERROR, "isVTOL", "No valid droid provided": context.throwError "psDroid", 2051, "isVTOL": wzapi_isVTOL = False: Exit Function
    wzapi_isVTOL = DROID_isVtol(psDroid)
End Function

'-- ## safeDest(player, x, y)
'--
'-- Returns true if given player is safe from hostile fire at the given location, to
'-- the best of that player's map knowledge. Does not work in campaign at the moment.
'--
Public Function wzapi_safeDest(ByVal context As wzapi_execution_context, ByVal player As Long, ByVal x As Long, ByVal y As Long) As Boolean
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2062, LOG_ERROR, "safeDest", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2062, "safeDest": wzapi_safeDest = False: Exit Function
    If Not (tileOnMap(x, y)) Then debug_ 2063, LOG_ERROR, "safeDest", "Out of bounds coordinates(%d, %d)", x, y: context.throwError "tileOnMap(x, y)", 2063, "safeDest": wzapi_safeDest = False: Exit Function
    wzapi_safeDest = ((auxTile(x, y, player) And AUXBITS_DANGER) = 0)
End Function

'-- ## activateStructure(structure[, target])
'--
'-- Activate a special ability on a structure. Currently only works on the lassat.
'-- The lassat needs a target.
'--
Public Function wzapi_activateStructure(ByVal context As wzapi_execution_context, ByVal psStruct As Long, ByVal has_psTarget As Boolean, ByVal opt_psTarget As Long) As Boolean
    Dim player As Long, psTarget As Long
    If Not (psStruct <> 0) Then debug_ 2074, LOG_ERROR, "activateStructure", "No valid structure provided": context.throwError "psStruct", 2074, "activateStructure": wzapi_activateStructure = False: Exit Function
    player = aStruct(psStruct).player
    ' ... and then do nothing with psStruct yet
    If has_psTarget Then psTarget = opt_psTarget Else psTarget = 0
    If Not (psTarget <> 0) Then debug_ 2078, LOG_ERROR, "activateStructure", "No valid target provided": context.throwError "psTarget", 2078, "activateStructure": wzapi_activateStructure = False: Exit Function
    orderStructureObj player, psTarget
    wzapi_activateStructure = True
End Function

'-- ## chat(playerFilter, message)
'--
'-- Send a message to playerFilter. playerFilter may also be ```ALL_PLAYERS``` or ```ALLIES```.
'-- Returns a boolean that is true on success. (3.2+ only)
'--
Public Function wzapi_chat(ByVal context As wzapi_execution_context, ByVal playerFilter As Long, ByVal message As String) As Boolean
    Dim player As Long, chatMessage As InGameChatMessage
    player = context.player()
    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALL_PLAYERS Or playerFilter = ALLIES) Then debug_ 2091, LOG_ERROR, "chat", "Message to invalid player %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALL_PLAYERS || playerFilter == ALLIES", 2091, "chat": wzapi_chat = False: Exit Function
    chatMessage = InGameChatMessage_new(player, message)
    If playerFilter = ALLIES Then ' allies
        chatMessage.toAllies = True
    ElseIf playerFilter <> ALL_PLAYERS Then ' specific player
        InGameChatMessage_addReceiverByIndex chatMessage, playerFilter
    End If

    InGameChatMessage_send chatMessage
    wzapi_chat = True
End Function

'-- ## quickChat(playerFilter, messageEnum)
'--
'-- Send a message to playerFilter. playerFilter may also be ```ALL_PLAYERS``` or ```ALLIES```.
'-- ```messageEnum``` is the WzQuickChatMessage value. (See the WzQuickChatMessages global
'-- object for constants to use. Pass in these constants directly, do not hard-code values!)
'-- Returns a boolean that is true on success. (4.4+ only)
'--
Public Function wzapi_quickChat(ByVal context As wzapi_execution_context, ByVal playerFilter As Long, ByVal messageEnum As Long) As Boolean
    Dim player As Long, targeting As WzQuickChatTargeting, message As Long
    player = context.player()
    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALL_PLAYERS Or playerFilter = ALLIES) Then debug_ 2116, LOG_ERROR, "quickChat", "Message to invalid player %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALL_PLAYERS || playerFilter == ALLIES", 2116, "quickChat": wzapi_quickChat = False: Exit Function

    If playerFilter = ALLIES Then ' allies
        targeting.humanTeammates = True
        targeting.aiTeammates = True
    ElseIf playerFilter = ALL_PLAYERS Then
        targeting.all_ = True
    Else    ' specific player
        WzQuickChatTargeting_insertSpecificPlayer targeting, playerFilter
    End If

    If Not (to_WzQuickChatMessage(messageEnum, message)) Then debug_ 2134, LOG_ERROR, "quickChat", "Invalid messageEnum value: %d", messageEnum: context.throwError "to_WzQuickChatMessage(static_cast<uint32_t>(messageEnum), message)", 2134, "quickChat": wzapi_quickChat = False: Exit Function

    sendQuickChat message, player, targeting
    wzapi_quickChat = True
End Function

'-- ## getDroidPath(droid)
'--
'-- Get path of a droid.
'-- Returns an array of positions.
'--
Public Function wzapi_getDroidPath(ByVal context As wzapi_execution_context, ByVal psDroid As Long, result() As scr_position) As Long
    Dim startPos As Long, length As Long, i As Long, n As Long, pathCoords As Vector2i
    If Not (psDroid <> 0) Then debug_ 2147, LOG_ERROR, "getDroidPath", "No valid droid provided": context.throwError "psDroid", 2147, "getDroidPath": wzapi_getDroidPath = 0: Exit Function

    startPos = MAX_(aDroid(psDroid).sMove.pathIndex - 1, 0)
    length = aDroid(psDroid).sMove.asPath_size
    For i = startPos To length - 1
        pathCoords = aDroid(psDroid).sMove.asPath(i)
        If Not (worldOnMap(pathCoords.x, pathCoords.y)) Then ASSERT_FAILURE "worldOnMap(pathCoords.x, pathCoords.y)", "src/wzapi.cpp:2155", "getDroidPath", "Path off map!"
        pushPosition result, n, map_coord(pathCoords.x), map_coord(pathCoords.y)
    Next

    wzapi_getDroidPath = n
End Function

'-- ## addBeacon(x, y, playerFilter[, message])
'--
'-- Send a beacon message to target player. Target may also be ```ALLIES```.
'-- Message is currently unused. Returns a boolean that is true on success. (3.2+ only)
'--
Public Function wzapi_addBeacon(ByVal context As wzapi_execution_context, ByVal x_ As Long, ByVal y_ As Long, ByVal playerFilter As Long, ByVal has_message As Boolean, ByVal opt_message As String) As Boolean
    Dim x As Long, y As Long, message As String, me_ As Long, i As Long, send As Boolean
    If Not (x_ >= 0) Then debug_ 2169, LOG_ERROR, "addBeacon", "Beacon x value %d is less than zero", x_: context.throwError "_x >= 0", 2169, "addBeacon": wzapi_addBeacon = False: Exit Function
    If Not (y_ >= 0) Then debug_ 2170, LOG_ERROR, "addBeacon", "Beacon y value %d is less than zero", y_: context.throwError "_y >= 0", 2170, "addBeacon": wzapi_addBeacon = False: Exit Function
    If Not (x_ <= mapWidth) Then debug_ 2171, LOG_ERROR, "addBeacon", "Beacon x value %d is greater than mapWidth %d", x_, mapWidth: context.throwError "_x <= mapWidth", 2171, "addBeacon": wzapi_addBeacon = False: Exit Function
    If Not (y_ <= mapHeight) Then debug_ 2172, LOG_ERROR, "addBeacon", "Beacon y value %d is greater than mapHeight %d", y_, mapHeight: context.throwError "_y <= mapHeight", 2172, "addBeacon": wzapi_addBeacon = False: Exit Function

    x = world_coord(x_) + (TILE_UNITS \ 2)
    y = world_coord(y_) + (TILE_UNITS \ 2)

    If has_message Then
        message = opt_message
    End If
    me_ = context.player()
    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALLIES) Then debug_ 2183, LOG_ERROR, "addBeacon", "Message to invalid player filter %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALLIES", 2183, "addBeacon": wzapi_addBeacon = False: Exit Function
    For i = 0 To MAX_PLAYERS - 1
        send = False
        If i <> me_ Then
            If i = playerFilter Then
                send = True
            ElseIf playerFilter = ALLIES Then
                send = aiCheckAlliances(i, me_)
            End If
        End If
        If send Then
            debug_ 2188, LOG_MSG, "addBeacon", "adding script beacon to %d from %d", i, me_
            sendBeaconToPlayer x, y, i, me_, message
        End If
    Next
    wzapi_addBeacon = True
End Function

'-- ## removeBeacon(playerFilter)
'--
'-- Remove a beacon message sent to playerFilter. Target may also be ```ALLIES```.
'-- Returns a boolean that is true on success. (3.2+ only)
'--
Public Function wzapi_removeBeacon(ByVal context As wzapi_execution_context, ByVal playerFilter As Long) As Boolean
    Dim me_ As Long, i As Long, psMessage As Long, match As Boolean
    me_ = context.player()

    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALLIES) Then debug_ 2204, LOG_ERROR, "removeBeacon", "Message to invalid player filter %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALLIES", 2204, "removeBeacon": wzapi_removeBeacon = False: Exit Function
    For i = 0 To MAX_PLAYERS - 1
        match = (i = playerFilter)
        If Not match And playerFilter = ALLIES Then match = aiCheckAlliances(i, me_)
        If match Then
            psMessage = findBeaconMsg(i, me_)
            If psMessage <> 0 Then
                removeMessage psMessage, i
                triggerEventBeaconRemoved me_, i
            End If
        End If
    Next
'	jsDebugMessageUpdate();
    wzapi_removeBeacon = True
End Function

'-- ## getDroidProduction(factory)
'--
'-- Return droid in production in given factory. Note that this droid is fully
'-- virtual, and should never be passed anywhere. (3.2+ only)
'--
' (VB) std::unique_ptr<const DROID>: a droid slot on no list (0 = nullptr), which the caller frees with
' DroidContainer_erase (the unique_ptr's delete: ~DROID)
Public Function wzapi_getDroidProduction(ByVal context As wzapi_execution_context, ByVal psFactory_ As Long) As Long
    Dim psStruct As Long, player As Long, t As STRUCTURE_TYPE, psTemp As Long, psDroid As Long
    psStruct = psFactory_
    If Not (psStruct <> 0) Then debug_ 2229, LOG_ERROR, "getDroidProduction", "No valid structure provided": context.throwError "psStruct", 2229, "getDroidProduction": wzapi_getDroidProduction = 0: Exit Function
    player = aStruct(psStruct).player
    t = asStructureStats(aStruct(psStruct).pStructureType).typ
    If Not (t = REF_FACTORY Or t = REF_CYBORG_FACTORY Or t = REF_VTOL_FACTORY) Then
        debug_ 2231, LOG_ERROR, "getDroidProduction", "Structure not a factory"
        context.throwError "psStruct->pStructureType->type == REF_FACTORY || psStruct->pStructureType->type == REF_CYBORG_FACTORY || psStruct->pStructureType->type == REF_VTOL_FACTORY", 2231, "getDroidProduction"
        wzapi_getDroidProduction = 0
        Exit Function
    End If
    ' const FACTORY *psFactory = &psStruct->pFunctionality->factory;
    psTemp = aStruct(psStruct).functionality.factory.psSubject
    If psTemp = 0 Then
        wzapi_getDroidProduction = 0
        Exit Function
    End If
    ' Since it's not intended to be used anywhere, don't put it in the global droid storage.
    psDroid = DroidContainer_alloc()   ' (VB) new DROID(0, player): a slot of the pool, on no list
    DROID_construct_ psDroid, 0, player
    aDroid(psDroid).pos = aStruct(psStruct).pos
    aDroid(psDroid).rot = aStruct(psStruct).rot
    aDroid(psDroid).experience = 0
    droidSetName psDroid, apsTemplates(psTemp).name_
    droidSetBits psTemp, psDroid
    aDroid(psDroid).weight = calcDroidWeight(apsTemplates(psTemp))
    aDroid(psDroid).baseSpeed = calcDroidBaseSpeed(apsTemplates(psTemp), aDroid(psDroid).weight, player)
    wzapi_getDroidProduction = psDroid
End Function

'-- ## getDroidLimit([player[, droidType]])
'--
'-- Return maximum number of droids that this player can produce. This limit is usually
'-- fixed throughout a game and the same for all players. If no arguments are passed,
'-- returns general droid limit for the current player. If a second, droid type argument
'-- is passed, the limit for this droid type is returned, which may be different from
'-- the general droid limit (eg for commanders and construction droids). (3.2+ only)
'--
Public Function wzapi_getDroidLimit(ByVal context As wzapi_execution_context, ByVal has_player As Boolean, ByVal opt_player As Long, ByVal has_droidType As Boolean, ByVal opt_droidType As Long) As Long
    Dim player As Long, droidType As DROID_TYPE
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2263, LOG_ERROR, "getDroidLimit", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2263, "getDroidLimit": wzapi_getDroidLimit = 0: Exit Function
    If has_droidType Then
        droidType = opt_droidType
        If droidType = DROID_COMMAND Then
            wzapi_getDroidLimit = getMaxCommanders(player)
            Exit Function
        ElseIf droidType = DROID_CONSTRUCT Then
            wzapi_getDroidLimit = getMaxConstructors(player)
            Exit Function
        End If
        ' else return general unit limit
    End If
    wzapi_getDroidLimit = getMaxDroids(player)
End Function

'-- ## getExperienceModifier(player)
'--
'-- Get the percentage of experience this player droids are going to gain. (3.2+ only)
'--
Public Function wzapi_getExperienceModifier(ByVal context As wzapi_execution_context, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2286, LOG_ERROR, "getExperienceModifier", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2286, "getExperienceModifier": wzapi_getExperienceModifier = 0: Exit Function
    wzapi_getExperienceModifier = getExpGain(player)
End Function

'-- ## setDroidLimit(player, maxNumber[, droidType])
'--
'-- Set the maximum number of droids that this player can produce. If a third
'-- parameter is added, this is the droid type to limit. It can be ```DROID_ANY```
'-- for droids in general, ```DROID_CONSTRUCT``` for constructors, or ```DROID_COMMAND```
'-- for commanders. (3.2+ only)
'--
Public Function wzapi_setDroidLimit(ByVal context As wzapi_execution_context, ByVal player As Long, ByVal maxNumber As Long, ByVal has_droidType As Boolean, ByVal opt_droidType As Long) As Boolean
    Dim droidType As DROID_TYPE
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2299, LOG_ERROR, "setDroidLimit", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2299, "setDroidLimit": wzapi_setDroidLimit = False: Exit Function
    If has_droidType Then droidType = opt_droidType Else droidType = DROID_ANY

    Select Case droidType
    Case DROID_CONSTRUCT
        setMaxConstructors player, maxNumber
    Case DROID_COMMAND
        setMaxCommanders player, maxNumber
    Case Else  ' default:, DROID_ANY
        setMaxDroids player, maxNumber
    End Select
    wzapi_setDroidLimit = True
End Function

'-- ## setCommanderLimit(player, maxNumber)
'--
'-- Set the maximum number of commanders that this player can produce.
'-- THIS FUNCTION IS DEPRECATED AND WILL BE REMOVED! (3.2+ only)
'--
Public Function wzapi_setCommanderLimit(ByVal context As wzapi_execution_context, ByVal player As Long, ByVal maxNumber As Long) As Boolean
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2325, LOG_ERROR, "setCommanderLimit", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2325, "setCommanderLimit": wzapi_setCommanderLimit = False: Exit Function
    setMaxCommanders player, maxNumber
    wzapi_setCommanderLimit = True
End Function

'-- ## setConstructorLimit(player, maxNumber)
'--
'-- Set the maximum number of constructors that this player can produce.
'-- THIS FUNCTION IS DEPRECATED AND WILL BE REMOVED! (3.2+ only)
'--
Public Function wzapi_setConstructorLimit(ByVal context As wzapi_execution_context, ByVal player As Long, ByVal maxNumber As Long) As Boolean
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2337, LOG_ERROR, "setConstructorLimit", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2337, "setConstructorLimit": wzapi_setConstructorLimit = False: Exit Function
    setMaxConstructors player, maxNumber
    wzapi_setConstructorLimit = True
End Function

'-- ## setExperienceModifier(player, percent)
'--
'-- Set the percentage of experience this player droids are going to gain. (3.2+ only)
'--
Public Function wzapi_setExperienceModifier(ByVal context As wzapi_execution_context, ByVal player As Long, ByVal percent As Long) As Boolean
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2348, LOG_ERROR, "setExperienceModifier", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2348, "setExperienceModifier": wzapi_setExperienceModifier = False: Exit Function
    setExpGain player, percent
    wzapi_setExperienceModifier = True
End Function

'-- ## enumCargo(transporterDroid)
'--
'-- Returns an array of droid objects inside given transport. (3.2+ only)
'--
Public Function wzapi_enumCargo(ByVal context As wzapi_execution_context, ByVal psDroid As Long, result() As Long) As Long
    Dim n As Long, it As Long, psCurr As Long, members As Long
    If Not (psDroid <> 0) Then debug_ 2359, LOG_ERROR, "enumCargo", "No valid droid provided": context.throwError "psDroid", 2359, "enumCargo": wzapi_enumCargo = 0: Exit Function
    If Not (DROID_isTransporter(psDroid)) Then debug_ 2360, LOG_ERROR, "enumCargo", "Wrong droid type (expecting: transporter)": context.throwError "psDroid->isTransporter()", 2360, "enumCargo": wzapi_enumCargo = 0: Exit Function
    n = 0
    If Not (aDroid(psDroid).psGroup <> 0) Then debug_ 2362, LOG_ERROR, "enumCargo", "Droid is not assigned to a group": context.throwError "psDroid->psGroup", 2362, "enumCargo": wzapi_enumCargo = 0: Exit Function
    members = DROID_GROUP_getNumMembers(aDroid(psDroid).psGroup)  ' result.reserve(...)
    If members > 0 Then ReDim result(0 To members - 1)
    it = aGroup(aDroid(psDroid).psGroup).psList.head
    Do While it <> 0
        psCurr = lnObj(it)
        If psDroid <> psCurr Then
            If n >= members Then members = n + 1: ReDim Preserve result(0 To n)
            result(n) = psCurr
            n = n + 1
        End If
        it = lnNext(it)
    Loop
    wzapi_enumCargo = n
End Function

'-- ## isSpectator(player)
'--
'-- Returns whether a particular player is a spectator. (4.2+ only)
'-- Can pass -1 as player to get the spectator status of the client running the script. (Useful for the "rules" scripts.)
'--
Public Function wzapi_isSpectator(ByVal context As wzapi_execution_context, ByVal player As Long) As Boolean
    If Not (player = -1 Or (player >= 0 And (player < NetPlay_players_size() Or player = selectedPlayer))) Then debug_ 2381, LOG_ERROR, "isSpectator", "Invalid player index %d", player: context.throwError "player == -1 || (player >= 0 && (player < NetPlay.players.size() || player == selectedPlayer))", 2381, "isSpectator": wzapi_isSpectator = False: Exit Function
    If player = -1 Or player = selectedPlayer Then
        ' TODO: Offers the ability to store this outside of NetPlayer.players later
        ' For now, it's stored in NetPlay.players[selectedPlayer]
        wzapi_isSpectator = NetPlay_players_isSpectator(selectedPlayer)
        Exit Function
    ElseIf player >= 0 And player < NetPlay_players_size() Then
        wzapi_isSpectator = NetPlay_players_isSpectator(player)
        Exit Function
    End If
    wzapi_isSpectator = True
End Function

'-- ## getWeaponInfo(weaponName)
'--
'-- Return information about a particular weapon type. DEPRECATED - query the Stats object instead. (3.2+ only)
'--
Public Function wzapi_getWeaponInfo(ByVal context As wzapi_execution_context, ByVal weaponName As String) As nlohmann_json
    Dim weaponIndex As Long, result As nlohmann_json, impactClass As String
    weaponIndex = getCompFromName(COMP_WEAPON, weaponName)
    If Not (weaponIndex >= 0 And weaponIndex < asWeaponStats_size) Then debug_ 2402, LOG_ERROR, "getWeaponInfo", "No such weapon: %s", weaponName: context.throwError "weaponIndex >= 0 && weaponIndex < asWeaponStats.size()", 2402, "getWeaponInfo": Set wzapi_getWeaponInfo = jsonNull(): Exit Function
    ' WEAPON_STATS *psStats = &asWeaponStats[weaponIndex]
    Set result = json_object()
    result.set_item_ref "id", jsonString(weaponName)
    result.set_item_ref "name", jsonString(asWeaponStats(weaponIndex).name_)
    If asWeaponStats(weaponIndex).weaponClass = WC_KINETIC Then impactClass = "KINETIC" Else impactClass = "HEAT"
    result.set_item_ref "impactClass", jsonString(impactClass)
    result.set_item_ref "damage", jsonUnsigned(asWeaponStats(weaponIndex).base.damage)
    result.set_item_ref "firePause", jsonUnsigned(asWeaponStats(weaponIndex).base.firePause)
    result.set_item_ref "fireOnMove", jsonBool(asWeaponStats(weaponIndex).fireOnMove)
    Set wzapi_getWeaponInfo = result
End Function

' MARK: - Functions that operate on the current player only

'-- ## centreView(x, y)
'--
'-- Center the player's camera at the given position.
'--
Public Function wzapi_centreView(ByVal context As wzapi_execution_context, ByVal x As Long, ByVal y As Long) As Boolean
    If Not (x >= 0) Then debug_ 2422, LOG_ERROR, "centreView", "x value %d is less than zero", x: context.throwError "x >= 0", 2422, "centreView": wzapi_centreView = False: Exit Function
    If Not (y >= 0) Then debug_ 2423, LOG_ERROR, "centreView", "y value %d is less than zero", y: context.throwError "y >= 0", 2423, "centreView": wzapi_centreView = False: Exit Function
    If Not (x <= mapWidth) Then debug_ 2424, LOG_ERROR, "centreView", "x value %d is greater than mapWidth %d", x, mapWidth: context.throwError "x <= mapWidth", 2424, "centreView": wzapi_centreView = False: Exit Function
    If Not (y <= mapHeight) Then debug_ 2425, LOG_ERROR, "centreView", "y value %d is greater than mapHeight %d", y, mapHeight: context.throwError "y <= mapHeight", 2425, "centreView": wzapi_centreView = False: Exit Function
    setViewPos x, y, False
    wzapi_centreView = True
End Function

'-- ## playSound(sound[, x, y, z])
'--
'-- Play a sound, optionally at a location.
'--
Public Function wzapi_playSound(ByVal context As wzapi_execution_context, ByVal sound As String, ByVal has_x As Boolean, ByVal opt_x As Long, ByVal has_y As Boolean, ByVal opt_y As Long, ByVal has_z As Boolean, ByVal opt_z As Long) As Boolean
    Dim player As Long, soundID As Long, x As Long, y As Long, z As Long
    player = context.player()
    If player <> selectedPlayer Then
        wzapi_playSound = False
        Exit Function
    End If
    soundID = audio_GetTrackID(sound)
    If soundID = SAMPLE_NOT_FOUND Then
        soundID = audio_SetTrackVals(sound, False, 100, 1800)
    End If
    If has_x Then
        x = world_coord(opt_x)
        If has_y Then y = world_coord(opt_y) Else y = world_coord(0)
        If has_z Then z = world_coord(opt_z) Else z = world_coord(0)
        audio_QueueTrackPos soundID, x, y, z
    Else
        audio_QueueTrack soundID
    End If
    wzapi_playSound = True
End Function

'-- ## emitSound(sound, x, y)
'--
'-- Emit a sound from an in-game location.
'--
Public Function wzapi_emitSound(ByVal context As wzapi_execution_context, ByVal sound As String, ByVal x As Long, ByVal y As Long) As Boolean
    Dim soundID As Long
    If Not (x >= 0) Then debug_ 2466, LOG_ERROR, "emitSound", "x value %d is less than zero", x: context.throwError "x >= 0", 2466, "emitSound": wzapi_emitSound = False: Exit Function
    If Not (y >= 0) Then debug_ 2467, LOG_ERROR, "emitSound", "y value %d is less than zero", y: context.throwError "y >= 0", 2467, "emitSound": wzapi_emitSound = False: Exit Function
    If Not (x <= mapWidth) Then debug_ 2468, LOG_ERROR, "emitSound", "x value %d is greater than mapWidth %d", x, mapWidth: context.throwError "x <= mapWidth", 2468, "emitSound": wzapi_emitSound = False: Exit Function
    If Not (y <= mapHeight) Then debug_ 2469, LOG_ERROR, "emitSound", "y value %d is greater than mapHeight %d", y, mapHeight: context.throwError "y <= mapHeight", 2469, "emitSound": wzapi_emitSound = False: Exit Function

    soundID = audio_GetTrackID(sound)
    If soundID = SAMPLE_NOT_FOUND Then
        soundID = audio_SetTrackVals(sound, False, 100, 1800)
    End If

    wzapi_emitSound = audio_PlayStaticTrack(x * TILE_UNITS, y * TILE_UNITS, soundID)
End Function

'-- ## addGuideTopic(guideTopicID[, showFlags[, excludedTopicIDs]])
'--
'-- Add a guide topic to the in-game guide.
'--
'-- guideTopicID is expected to be a "::"-delimited guide topic id (which corresponds to the .json file containing the guide topic information).
'-- > For example, ```"wz2100::structures::hq"``` will attempt to load ```"guidetopics/wz2100/structures/hq.json"```, the guide topic file about the hq / command center.
'--
'-- guideTopicID also has limited support for trailing wildcards. For example:
'-- - ```"wz2100::units::*"``` will load all guide topic .json files in the folder ```"guidetopics/wz2100/units/"``` (but not any subfolders)
'-- - ```"wz2100::**"``` will load all guide topic .json files within the folder ```"guidetopics/wz2100/"``` **and** all subfolders
'--
'-- (The wildcard is only supported in the last position.)
'--
'-- showFlags can be used to configure automatic display of the guide, and can be set to one or more of:
'-- - ```SHOWTOPIC_FIRSTADD```: open guide only if this topic is newly-added (this playthrough) - if topic was already added, the guide won't be automatically displayed
'-- - ```SHOWTOPIC_NEVERVIEWED```: open guide only if this topic has never been viewed by the player before (in any playthrough)
'-- You can also specify multiple flags (ex. ```SHOWTOPIC_FIRSTADD | SHOWTOPIC_NEVERVIEWED```).
'-- The default behavior (where showFlags is omitted) merely adds the topic to the guide, but does not automatically display / open the guide.
'--
'-- excludedTopicIDs can be a string or a list of string guide topic IDs (non-wildcard) to be excluded, when supplying a wildcard guideTopicID.
'--
Public Function wzapi_addGuideTopic(ByVal context As wzapi_execution_context, ByVal guideTopicID As String, ByVal has_showFlags As Boolean, ByVal opt_showFlags As Long, ByVal has_excludedTopicIDs As Boolean, opt_excludedTopicIDs As StringVector) As Boolean
    Dim showFlags As Long
    If has_showFlags Then showFlags = opt_showFlags Else showFlags = 0
    If has_excludedTopicIDs Then
        ' (std::unordered_set<std::string> excludedTopicIDsSet: the list, its duplicates harmless)
        addGuideTopic_excluding guideTopicID, showFlags, opt_excludedTopicIDs
    Else
        addGuideTopic guideTopicID, showFlags
    End If
    wzapi_addGuideTopic = True
End Function

'-- ## gameOverMessage(gameWon[, showBackDrop[, showOutro]])
'--
'-- End game in victory or defeat.
'--
Public Function wzapi_gameOverMessage(ByVal context As wzapi_execution_context, ByVal gameWon As Boolean, ByVal has_showBackDrop As Boolean, ByVal opt_showBackDrop As Boolean, ByVal has_showOutro As Boolean, ByVal opt_showOutro As Boolean) As Boolean
    Dim player As Long, msgType As MESSAGE_TYPE, showBackDrop As Boolean, showOutro As Boolean
    Dim psViewData As Long, psMessage As Long
    player = context.player()
    msgType = MSG_MISSION   ' always
    If has_showBackDrop Then showBackDrop = opt_showBackDrop Else showBackDrop = True
    If has_showOutro Then showOutro = opt_showOutro Else showOutro = False
    If gameWon Then
        'Quick hack to stop assert when trying to play outro in campaign.
        If (Not bMultiPlayer) And showOutro Then psViewData = getViewData("END") Else psViewData = getViewData("WIN")
        addConsoleMessage "YOU ARE VICTORIOUS!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    Else
        psViewData = getViewData("END")    ' FIXME: rename to FAILED|LOST ?
        If Not testPlayerHasLost() Then     ' check for whether the player started as a spectator or already lost (in either case the player is already marked as having lost)
            addConsoleMessage "YOU WERE DEFEATED!", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        End If
    End If
    If Not (psViewData <> 0) Then ASSERT_FAILURE "psViewData", "src/wzapi.cpp:2540", "gameOverMessage", "Viewdata not found"
    psMessage = 0
    If player < MAX_PLAYERS Then
        psMessage = addMessage(msgType, False, player)
    End If
    If (Not bMultiPlayer) And psMessage <> 0 Then
        'we need to set this here so the VIDEO_QUIT callback is not called
        If gameWon Then setScriptWinLoseVideo PLAY_WIN Else setScriptWinLoseVideo PLAY_LOSE
        seq_ClearSeqList
        If gameWon And showOutro Then
            showBackDrop = False
            seq_AddSeqToList "outro.ogg", vbNullString, "outro.txa", False
            seq_StartNextFullScreenVideo
        Else
            'set the data
            MESSAGE_pViewData(psMessage) = psViewData
            displayImmediateMessage psMessage
            stopReticuleButtonFlash IDRET_INTEL_MAP
        End If
    End If
    '	jsDebugMessageUpdate();
    displayGameOver gameWon, showBackDrop
    If challengeActive Then
        updateChallenge gameWon
    End If
    GameStoryLogger_logGameOver
    If autogame_enabled() Then
        debug_ 2574, LOG_WARNING, "gameOverMessage", "Autogame completed successfully!"
        If headlessGameMode() Then
            stdOutGameSummary 0
        End If
        wzQuit 0 ' Trigger a *graceful* shutdown
    ElseIf headlessGameMode() Then
        debug_ 2583, LOG_WARNING, "gameOverMessage", "Headless game completed successfully!"
        wzQuit 0 ' Trigger a *graceful* shutdown
    End If
    wzapi_gameOverMessage = True
End Function

' MARK: - Global state manipulation -- not for use with skirmish AI (unless you want it to cheat, obviously)

'-- ## setStructureLimits(structureName, limit[, player])
'--
'-- Set build limits for a structure.
'--
Public Function wzapi_setStructureLimits(ByVal context As wzapi_execution_context, ByVal structureName As String, ByVal limit As Long, ByVal has_player As Boolean, ByVal opt_player As Long) As Boolean
    Dim structureIndex As Long, player As Long
    structureIndex = getStructStatFromName(structureName)
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2599, LOG_ERROR, "setStructureLimits", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2599, "setStructureLimits": wzapi_setStructureLimits = False: Exit Function
    ' (limit < LOTS_OF compares unsigned: LOTS_OF is 0xFFFFFFFF)
    If Not (ULT(limit, LOTS_OF) And limit >= 0) Then debug_ 2600, LOG_ERROR, "setStructureLimits", "Invalid limit": context.throwError "limit < LOTS_OF && limit >= 0", 2600, "setStructureLimits": wzapi_setStructureLimits = False: Exit Function
    If Not (structureIndex >= 0 And structureIndex < numStructureStats) Then debug_ 2601, LOG_ERROR, "setStructureLimits", "Structure %s not found", structureName: context.throwError "structureIndex >= 0 && structureIndex < numStructureStats", 2601, "setStructureLimits": wzapi_setStructureLimits = False: Exit Function

    asStructureStats(structureIndex).upgrade(player).limit = limit

    wzapi_setStructureLimits = True
End Function

'-- ## applyLimitSet()
'--
'-- Mix user set limits with script set limits and defaults.
'--
Public Function wzapi_applyLimitSet(ByVal context As wzapi_execution_context) As Boolean
    wzapi_applyLimitSet = applyLimitSet()
End Function

'-- ## setMissionTime(time)
'--
'-- Set mission countdown in seconds.
'--
Public Sub wzapi_setMissionTime(ByVal context As wzapi_execution_context, ByVal time_ As Long)
    Dim missionTime As Long
    missionTime = time_ * GAME_TICKS_PER_SEC   ' int time
    mission_startTime = gameTime
    mission_time = missionTime
    setMissionCountDown
    If mission_time >= 0 Then
        mission_startTime = gameTime
        addMissionTimerInterface
    Else
        intRemoveMissionTimer
        mission_cheatTime = False
    End If
End Sub

'-- ## getMissionTime()
'--
'-- Get time remaining on mission countdown in seconds. (3.2+ only)
'--
Public Function wzapi_getMissionTime(ByVal context As wzapi_execution_context) As Long
    If mission_time < 0 Then
        wzapi_getMissionTime = -1
        Exit Function
    End If
    ' (gameTime and mission.startTime are UDWORDs: the difference and the division are unsigned)
    wzapi_getMissionTime = FromU32(Int(U32(mission_time - (gameTime - mission_startTime)) / GAME_TICKS_PER_SEC))
End Function

'-- ## setReinforcementTime(time[, removeLaunch])
'--
'-- Set time for reinforcements to arrive. If time is negative, the reinforcement GUI
'-- is removed and the timer stopped. Time is in seconds.
'-- If time equals to the magic ```LZ_COMPROMISED_TIME``` constant, reinforcement GUI ticker
'-- is set to "--:--" and reinforcements are suppressed until this function is called
'-- again with a regular time value.
'--
Public Sub wzapi_setReinforcementTime(ByVal context As wzapi_execution_context, ByVal time_ As Long, ByVal has_removeLaunch As Boolean, ByVal opt_removeLaunch As Boolean)
    Dim reinforcementTime As Long, removeLaunch As Boolean, it As Long, psDroid As Long
    reinforcementTime = time_ * GAME_TICKS_PER_SEC     ' int time
    If has_removeLaunch Then removeLaunch = opt_removeLaunch Else removeLaunch = True
    If Not (reinforcementTime = LZ_COMPROMISED_TIME Or reinforcementTime < 60& * 60 * GAME_TICKS_PER_SEC) Then debug_ 2665, LOG_ERROR, "setReinforcementTime", "The transport timer cannot be set to more than 1 hour!": context.throwError "time == LZ_COMPROMISED_TIME || time < 60 * 60 * GAME_TICKS_PER_SEC", 2665, "setReinforcementTime": Exit Sub
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then debug_ 2666, LOG_ERROR, "setReinforcementTime", "Invalid selectedPlayer for current client: %u", selectedPlayer: context.throwError "selectedPlayer < MAX_PLAYERS", 2666, "setReinforcementTime": Exit Sub
    mission_ETA = reinforcementTime
    If reinforcementTime < 0 Then
        intRemoveTransporterTimer
    End If
    If removeLaunch Then
        ' Search for a transport that is idle; if we can't find any, remove the launch button
        ' since there's no transport to launch
        psDroid = 0
        it = apsDroidLists(selectedPlayer).head
        Do While it <> 0
            If DROID_isTransporter(lnObj(it)) Then
                If Not transporterFlying(lnObj(it)) Then
                    psDroid = lnObj(it)
                    Exit Do
                End If
            End If
            it = lnNext(it)
        Loop

        ' Didn't find an idle transporter, we can remove the launch button
        If psDroid = 0 Then
            intRemoveTransporterLaunch
        End If
    End If
    resetMissionWidgets
    addTransporterTimerInterface
End Sub

'-- ## completeResearch(researchName[, player[, forceResearch]])
'--
'-- Finish a research for the given player.
'-- forceResearch will allow a research topic to be researched again. 3.3+
'--
Public Sub wzapi_completeResearch(ByVal context As wzapi_execution_context, ByVal researchName As String, ByVal has_player As Boolean, ByVal opt_player As Long, ByVal has_forceResearch As Boolean, ByVal opt_forceResearch As Boolean)
    Dim player As Long, forceIt As Boolean, psResearch As Long
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2701, LOG_ERROR, "completeResearch", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2701, "completeResearch": Exit Sub
    If has_forceResearch Then forceIt = opt_forceResearch Else forceIt = False
    psResearch = getResearch(researchName)
    If Not (psResearch <> -1) Then debug_ 2704, LOG_ERROR, "completeResearch", "No such research %s for player %d", researchName, player: context.throwError "psResearch", 2704, "completeResearch": Exit Sub
    If Not (ULT(asResearch(psResearch).index, asResearch_size)) Then debug_ 2705, LOG_ERROR, "completeResearch", "Research index out of bounds": context.throwError "psResearch->index < asResearch.size()", 2705, "completeResearch": Exit Sub
    ' PLAYER_RESEARCH *plrRes = &asPlayerResList[player][psResearch->index]
    If Not forceIt Then
        If IsResearchCompleted(asPlayerResList(player, asResearch(psResearch).index)) Then
            Exit Sub
        End If
    End If
    If bMultiMessages And (gameTime > 2) Then ' ??? "gameTime > 2" ??
        SendResearch player, asResearch(psResearch).index, False
        ' Wait for our message before doing anything.
    Else
        researchResult asResearch(psResearch).index, player, False, 0, False
    End If
End Sub

'-- ## completeAllResearch([player])
'--
'-- Finish all researches for the given player.
'--
Public Sub wzapi_completeAllResearch(ByVal context As wzapi_execution_context, ByVal has_player As Boolean, ByVal opt_player As Long)
    Dim player As Long, i As Long
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2730, LOG_ERROR, "completeAllResearch", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2730, "completeAllResearch": Exit Sub
    For i = 0 To asResearch_size - 1
        ' RESEARCH *psResearch = &asResearch[i]; PLAYER_RESEARCH *plrRes = &asPlayerResList[player][psResearch->index]
        If Not IsResearchCompleted(asPlayerResList(player, asResearch(i).index)) Then
            If bMultiMessages And (gameTime > 2) Then
                SendResearch player, asResearch(i).index, False
                ' Wait for our message before doing anything.
            Else
                researchResult asResearch(i).index, player, False, 0, False
            End If
        End If
    Next
End Sub

'-- ## enableResearch(researchName[, player])
'--
'-- Enable a research for the given player, allowing it to be researched.
'--
Public Function wzapi_enableResearch(ByVal context As wzapi_execution_context, ByVal researchName As String, ByVal has_player As Boolean, ByVal opt_player As Long) As Boolean
    Dim player As Long, psResearch As Long
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2758, LOG_ERROR, "enableResearch", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2758, "enableResearch": wzapi_enableResearch = False: Exit Function
    psResearch = getResearch(researchName)
    If Not (psResearch <> -1) Then debug_ 2760, LOG_ERROR, "enableResearch", "No such research %s for player %d", researchName, player: context.throwError "psResearch", 2760, "enableResearch": wzapi_enableResearch = False: Exit Function
    If Not enableResearch(psResearch, player) Then
        debug_ 2763, LOG_ERROR, "enableResearch", "Unable to enable research %s for player %d", researchName, player
        wzapi_enableResearch = False
        Exit Function
    End If
    wzapi_enableResearch = True
End Function

'-- ## setPower(power[, player])
'--
'-- Set a player's power directly. (Do not use this in an AI script.)
'--
Public Sub wzapi_setPower(ByVal context As wzapi_execution_context, ByVal power As Long, ByVal has_player As Boolean, ByVal opt_player As Long)
    Dim player As Long
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2776, LOG_ERROR, "setPower", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2776, "setPower": Exit Sub
    setPower player, power
End Sub

'-- ## setPowerModifier(powerModifier[, player])
'--
'-- Set a player's power modifier percentage. (Do not use this in an AI script.) (3.2+ only)
'--
Public Sub wzapi_setPowerModifier(ByVal context As wzapi_execution_context, ByVal powerModifier As Long, ByVal has_player As Boolean, ByVal opt_player As Long)
    Dim player As Long
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2788, LOG_ERROR, "setPowerModifier", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2788, "setPowerModifier": Exit Sub
    setPowerModifier player, powerModifier
End Sub

'-- ## setPowerStorageMaximum(powerMaximum[, player])
'--
'-- Set a player's power storage maximum. (Do not use this in an AI script.) (3.2+ only)
'--
Public Sub wzapi_setPowerStorageMaximum(ByVal context As wzapi_execution_context, ByVal powerMaximum As Long, ByVal has_player As Boolean, ByVal opt_player As Long)
    Dim player As Long
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2800, LOG_ERROR, "setPowerStorageMaximum", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2800, "setPowerStorageMaximum": Exit Sub
    setPowerMaxStorage player, powerMaximum
End Sub

'-- ## extraPowerTime(time[, player])
'--
'-- Increase a player's power as if that player had power income equal to current income
'-- over the given amount of extra time. (3.2+ only)
'--
Public Sub wzapi_extraPowerTime(ByVal context As wzapi_execution_context, ByVal time_ As Long, ByVal has_player As Boolean, ByVal opt_player As Long)
    Dim ticks As Long, player As Long
    ticks = time_ * GAME_UPDATES_PER_SEC
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 2814, LOG_ERROR, "extraPowerTime", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 2814, "extraPowerTime": Exit Sub
    updatePlayerPower player, ticks
End Sub

'-- ## setTutorialMode(enableTutorialMode)
'--
'-- Sets a number of restrictions appropriate for tutorial if set to true.
'--
Public Sub wzapi_setTutorialMode(ByVal context As wzapi_execution_context, ByVal enableTutorialMode As Boolean)
    bInTutorial = enableTutorialMode
End Sub

' (VB) template.cpp's enumerateTemplates(player, func): its ASSERT_PLAYER_OR_RETURN; the caller loops over
' droidTemplates_at(player, i), i below templateCount(player), when it is true
Private Function enumerateTemplates_player(ByVal player As Long) As Boolean
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:659", "enumerateTemplates", "Invalid player: %u", player: Exit Function
    enumerateTemplates_player = True
End Function

'-- ## setDesign(allowDesignValue)
'--
'-- Whether to allow player to design stuff.
'--
Public Sub wzapi_setDesign(ByVal context As wzapi_execution_context, ByVal allowDesignValue As Boolean)
    Dim psCurr As Long, i As Long, psTempl As Long, isResearched As Boolean
    If Not ULT(selectedPlayer, MAX_PLAYERS) Then Exit Sub
    ' Switch on or off future templates
    ' FIXME: This dual data structure for templates is just plain insane.
    If enumerateTemplates_player(selectedPlayer) Then
        For i = 0 To templateCount(selectedPlayer) - 1
            psTempl = droidTemplates_at(selectedPlayer, i)
            isResearched = researchedTemplate(psTempl, selectedPlayer, True)
            apsTemplates(psTempl).enabled = (isResearched Or allowDesignValue)
        Next
    End If
    For i = 0 To localTemplates_size - 1
        psCurr = localTemplates(i)
        isResearched = researchedTemplate(psCurr, selectedPlayer, True)
        apsTemplates(psCurr).enabled = (isResearched Or allowDesignValue)
    Next
End Sub

'-- ## enableTemplate(templateName)
'--
'-- Enable a specific template (even if design is disabled).
'--
Public Function wzapi_enableTemplate(ByVal context As wzapi_execution_context, ByVal templateName_ As String) As Boolean
    Dim psCurr As Long, templateName As String, bFound As Boolean, i As Long, psTempl As Long
    templateName = templateName_
    bFound = False
    ' FIXME: This dual data structure for templates is just plain insane.
    If enumerateTemplates_player(selectedPlayer) Then
        For i = 0 To templateCount(selectedPlayer) - 1
            psTempl = droidTemplates_at(selectedPlayer, i)
            If templateName = apsTemplates(psTempl).id Then
                apsTemplates(psTempl).enabled = True
                bFound = True
                Exit For ' break;
            End If
        Next
    End If
    If Not bFound Then
        debug_ 2874, LOG_ERROR, "enableTemplate", "Template %s was not found!", templateName
        wzapi_enableTemplate = False
        Exit Function
    End If
    For i = 0 To localTemplates_size - 1
        psCurr = localTemplates(i)
        If templateName = apsTemplates(psCurr).id Then
            apsTemplates(psCurr).enabled = True
            Exit For
        End If
    Next
    wzapi_enableTemplate = True
End Function

' (VB) hci.cpp's localTemplates.erase(i): the copy freed, the later handles moved up
Private Sub localTemplates_erase(ByVal k As Long)
    Dim j As Long
    deleteTemplate localTemplates(k)
    For j = k To localTemplates_size - 2
        localTemplates(j) = localTemplates(j + 1)
    Next
    localTemplates_size = localTemplates_size - 1
End Sub

'-- ## removeTemplate(templateName)
'--
'-- Remove a template.
'--
Public Function wzapi_removeTemplate(ByVal context As wzapi_execution_context, ByVal templateName_ As String) As Boolean
    Dim psCurr As Long, templateName As String, bFound As Boolean, i As Long, psTempl As Long
    templateName = templateName_
    bFound = False
    ' FIXME: This dual data structure for templates is just plain insane.
    If enumerateTemplates_player(selectedPlayer) Then
        For i = 0 To templateCount(selectedPlayer) - 1
            psTempl = droidTemplates_at(selectedPlayer, i)
            If templateName = apsTemplates(psTempl).id Then
                apsTemplates(psTempl).enabled = False
                bFound = True
                Exit For ' break;
            End If
        Next
    End If
    If Not bFound Then
        debug_ 2910, LOG_ERROR, "removeTemplate", "Template %s was not found!", templateName
        wzapi_removeTemplate = False
        Exit Function
    End If
    For i = 0 To localTemplates_size - 1
        psCurr = localTemplates(i)
        If templateName = apsTemplates(psCurr).id Then
            localTemplates_erase i
            Exit For
        End If
    Next
    wzapi_removeTemplate = True
End Function

'-- ## setMiniMap(visible)
'--
'-- Turns visible minimap on or off in the GUI.
'--
Public Sub wzapi_setMiniMap(ByVal context As wzapi_execution_context, ByVal visible As Boolean)
    radarPermitted = visible
End Sub

'-- ## setReticuleButton(buttonId, tooltip, filename, filenameDown[, callback])
'--
'-- Add reticule button. buttonId is which button to change, where zero is zero is the middle button, then going clockwise from
'-- the uppermost button. filename is button graphics and filenameDown is for highlighting. The tooltip is the text you see when
'-- you mouse over the button. Finally, the callback is which scripting function to call. Hide and show the user interface
'-- for such changes to take effect. (3.2+ only)
'--
Public Sub wzapi_setReticuleButton(ByVal context As wzapi_execution_context, ByVal buttonId As Long, ByVal tooltip As String, ByVal filename As String, ByVal filenameDown As String, ByVal has_callbackFuncName As Boolean, ByVal opt_callbackFuncName As String)
    Dim func As String, callback As Object
    If Not (buttonId >= 0 And buttonId <= 6) Then debug_ 2944, LOG_ERROR, "setReticuleButton", "Invalid button %d", buttonId: context.throwError "buttonId >= 0 && buttonId <= 6", 2944, "setReticuleButton": Exit Sub

    If has_callbackFuncName Then
        func = opt_callbackFuncName
    End If
    If MissionResUp Then
        Exit Sub ' no-op
    End If
    If Len(func) <> 0 Then Set callback = context.getNamedScriptCallback(func)   ' func.isEmpty() ? nullptr : ...
    setReticuleStats buttonId, tooltip, filename, filenameDown, callback
End Sub

'-- ## setReticuleFlash(buttonId, flash)
'--
'-- Set reticule flash on or off. (3.2.3+ only)
'--
Public Sub wzapi_setReticuleFlash(ByVal context As wzapi_execution_context, ByVal buttonId As Long, ByVal flash As Boolean)
    If Not (buttonId >= 0 And buttonId <= 6) Then debug_ 2965, LOG_ERROR, "setReticuleFlash", "Invalid button %d", buttonId: context.throwError "buttonId >= 0 && buttonId <= 6", 2965, "setReticuleFlash": Exit Sub
    setReticuleFlash buttonId, flash
End Sub

'-- ## showReticuleWidget(buttonId)
'--
'-- Open the reticule menu widget. (3.3+ only)
'--
Public Sub wzapi_showReticuleWidget(ByVal context As wzapi_execution_context, ByVal buttonId As Long)
    If Not (buttonId >= 0 And buttonId <= 6) Then debug_ 2976, LOG_ERROR, "showReticuleWidget", "Invalid button %d", buttonId: context.throwError "buttonId >= 0 && buttonId <= 6", 2976, "showReticuleWidget": Exit Sub
    intShowWidget buttonId
End Sub

'-- ## showInterface()
'--
'-- Show user interface. (3.2+ only)
'--
Public Sub wzapi_showInterface(ByVal context As wzapi_execution_context)
    intShowInterface
End Sub

'-- ## hideInterface()
'--
'-- Hide user interface. (3.2+ only)
'--
Public Sub wzapi_hideInterface(ByVal context As wzapi_execution_context)
    intHideInterface
End Sub

'-- ## enableStructure(structureName[, player])
'--
'-- The given structure type is made available to the given player. It will appear in the
'-- player's build list.
'--
Public Sub wzapi_enableStructure(ByVal context As wzapi_execution_context, ByVal structureName As String, ByVal has_player As Boolean, ByVal opt_player As Long)
    Dim structureIndex As Long, player As Long
    structureIndex = getStructStatFromName(structureName)
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3010, LOG_ERROR, "enableStructure", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3010, "enableStructure": Exit Sub
    If Not (structureIndex >= 0 And structureIndex < numStructureStats) Then debug_ 3011, LOG_ERROR, "enableStructure", "Structure %s not found", structureName: context.throwError "structureIndex >= 0 && structureIndex < numStructureStats", 3011, "enableStructure": Exit Sub
    ' enable the appropriate structure
    apStructTypeLists(player, structureIndex) = AVAILABLE
End Sub

Private Sub setComponent(ByVal componentName As String, ByVal player As Long, ByVal availability As Long)
    Dim psComp As Long
    psComp = getCompStatsFromName(componentName)
    If Not (psComp <> 0) Then ASSERT_FAILURE "psComp", "src/wzapi.cpp:3020", "setComponent", "Bad component %s", componentName: Exit Sub
    apCompLists(player, COMPONENT_STATS_compType(psComp), BASE_STATS_index(psComp)) = availability
End Sub

'-- ## enableComponent(componentName, player)
'--
'-- The given component is made available for research for the given player.
'--
Public Sub wzapi_enableComponent(ByVal context As wzapi_execution_context, ByVal componentName As String, ByVal player As Long)
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3030, LOG_ERROR, "enableComponent", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3030, "enableComponent": Exit Sub
    setComponent componentName, player, FOUND
End Sub

'-- ## makeComponentAvailable(componentName, player)
'--
'-- The given component is made available to the given player. This means the player can
'-- actually build designs with it.
'--
Public Sub wzapi_makeComponentAvailable(ByVal context As wzapi_execution_context, ByVal componentName As String, ByVal player As Long)
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3042, LOG_ERROR, "makeComponentAvailable", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3042, "makeComponentAvailable": Exit Sub
    setComponent componentName, player, AVAILABLE
End Sub

'-- ## allianceExistsBetween(player1, player2)
'--
'-- Returns true if an alliance exists between the two players, or they are the same player.
'--
Public Function wzapi_allianceExistsBetween(ByVal context As wzapi_execution_context, ByVal player1 As Long, ByVal player2 As Long) As Boolean
    If Not (player1 >= 0 And player1 < MAX_PLAYERS) Then debug_ 3053, LOG_ERROR, "allianceExistsBetween", "Invalid player index %d", player1: context.throwError "player1 >= 0 && player1 < MAX_PLAYERS", 3053, "allianceExistsBetween": wzapi_allianceExistsBetween = False: Exit Function
    If Not (player2 >= 0 And player2 < MAX_PLAYERS) Then debug_ 3054, LOG_ERROR, "allianceExistsBetween", "Invalid player index %d", player2: context.throwError "player2 >= 0 && player2 < MAX_PLAYERS", 3054, "allianceExistsBetween": wzapi_allianceExistsBetween = False: Exit Function
    wzapi_allianceExistsBetween = alliances(player1, player2) = ALLIANCE_FORMED
End Function

'-- ## removeStruct(structure)
'--
'-- Immediately remove the given structure from the map. Returns a boolean that is true on success.
'-- No special effects are applied. DEPRECATED since 3.2. Use `removeObject` instead.
'--
Public Function wzapi_removeStruct(ByVal context As wzapi_execution_context, ByVal psStruct As Long) As Boolean
    If Not (psStruct <> 0) Then debug_ 3065, LOG_ERROR, "removeStruct", "No valid structure provided": context.throwError "psStruct", 3065, "removeStruct": wzapi_removeStruct = False: Exit Function
    wzapi_removeStruct = removeStruct(psStruct, True)
End Function

'-- ## removeObject(gameObject[, sfx])
'--
'-- Queue the given game object for removal with or without special effects. Returns a boolean that is true on success.
'-- A second, optional boolean parameter specifies whether special effects are to be applied. (3.2+ only)
'--
'-- BREAKING CHANGE (4.5+): the effect of this function is not immediate anymore, the object will be
'-- queued for later removal instead of destroying it right away.
'-- User scripts should not rely on this function having immediate side-effects.
'--
Public Function wzapi_removeObject(ByVal context As wzapi_execution_context, ByVal psObj As Long, ByVal has_sfx As Boolean, ByVal opt_sfx As Boolean) As Boolean
    Dim typ As OBJECT_TYPE, sfx As Boolean
    If Not (psObj <> 0) Then debug_ 3080, LOG_ERROR, "removeObject", "No valid object provided": context.throwError "psObj", 3080, "removeObject": wzapi_removeObject = False: Exit Function
    typ = SIMPLE_OBJECT_type(psObj)
    If Not (typ = OBJ_STRUCTURE Or typ = OBJ_DROID Or typ = OBJ_FEATURE) Then
        debug_ 3081, LOG_ERROR, "removeObject", "Wrong game object type"
        context.throwError "psObj->type == OBJ_STRUCTURE || psObj->type == OBJ_DROID || psObj->type == OBJ_FEATURE", 3081, "removeObject"
        wzapi_removeObject = False
        Exit Function
    End If

    If has_sfx Then sfx = opt_sfx Else sfx = False
    queuedObjRemovals_emplace_back psObj, sfx
    wzapi_removeObject = True
End Function

'-- ## setScrollLimits(x1, y1, x2, y2)
'--
'-- Limit the scrollable area of the map to the given rectangle. (3.2+ only)
'--
Public Sub wzapi_setScrollLimits(ByVal context As wzapi_execution_context, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    Dim minX As Long, minY As Long, maxX As Long, maxY As Long
    Dim prevMinX As Long, prevMinY As Long, prevMaxX As Long, prevMaxY As Long
    Dim lx1 As Long, ly1 As Long, lx2 As Long, ly2 As Long
    minX = x1
    minY = y1
    maxX = x2
    maxY = y2

    If Not (minX >= 0) Then debug_ 3100, LOG_ERROR, "setScrollLimits", "Minimum scroll x value %d is less than zero - ", minX: context.throwError "minX >= 0", 3100, "setScrollLimits": Exit Sub
    If Not (minY >= 0) Then debug_ 3101, LOG_ERROR, "setScrollLimits", "Minimum scroll y value %d is less than zero - ", minY: context.throwError "minY >= 0", 3101, "setScrollLimits": Exit Sub
    If Not (maxX <= mapWidth) Then debug_ 3102, LOG_ERROR, "setScrollLimits", "Maximum scroll x value %d is greater than mapWidth %d", maxX, mapWidth: context.throwError "maxX <= mapWidth", 3102, "setScrollLimits": Exit Sub
    If Not (maxY <= mapHeight) Then debug_ 3103, LOG_ERROR, "setScrollLimits", "Maximum scroll y value %d is greater than mapHeight %d", maxY, mapHeight: context.throwError "maxY <= mapHeight", 3103, "setScrollLimits": Exit Sub

    prevMinX = scrollMinX
    prevMinY = scrollMinY
    prevMaxX = scrollMaxX
    prevMaxY = scrollMaxY

    scrollMinX = minX
    scrollMaxX = maxX
    scrollMinY = minY
    scrollMaxY = maxY

    ' When the scroll limits change midgame - need to redo the lighting
    If prevMinX < scrollMinX Then lx1 = prevMinX Else lx1 = scrollMinX
    If prevMinY < scrollMinY Then ly1 = prevMinY Else ly1 = scrollMinY
    If prevMaxX < scrollMaxX Then lx2 = prevMaxX Else lx2 = scrollMaxX
    If prevMaxY < scrollMaxY Then ly2 = prevMaxY Else ly2 = scrollMaxY
    initLighting lx1, ly1, lx2, ly2

    ' need to reset radar to take into account of new size
    resizeRadar
End Sub

'-- ## getScrollLimits()
'--
'-- Get the limits of the scrollable area of the map as an area object. (3.2+ only)
'--
Public Function wzapi_getScrollLimits(ByVal context As wzapi_execution_context) As scr_area
    Dim limits As scr_area
    limits.x1 = scrollMinX
    limits.y1 = scrollMinY
    limits.x2 = scrollMaxX
    limits.y2 = scrollMaxY
    wzapi_getScrollLimits = limits
End Function

'-- ## addStructure(structureName, player, x, y[, direction])
'--
'-- Create a structure on the given position. Returns the structure on success, null otherwise.
'-- Position uses world coordinates, if you want use position based on Map Tiles, then
'-- use as addStructure(structureName, players, x*128, y*128)
'-- Direction is optional, and is specified in degrees.
'--
Public Function wzapi_addStructure(ByVal context As wzapi_execution_context, ByVal structureName As String, ByVal player As Long, ByVal x As Long, ByVal y As Long, ByVal has_direction As Boolean, ByVal opt_direction As Long) As Long
    Dim structureIndex As Long, direction As Long, psStat As Long, psStruct As Long
    structureIndex = getStructStatFromName(structureName)
    If Not (structureIndex >= 0 And structureIndex < numStructureStats) Then debug_ 3150, LOG_ERROR, "addStructure", "Structure %s not found", structureName: context.throwError "structureIndex >= 0 && structureIndex < numStructureStats", 3150, "addStructure": wzapi_addStructure = 0: Exit Function
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3151, LOG_ERROR, "addStructure", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3151, "addStructure": wzapi_addStructure = 0: Exit Function

    ' uint16_t direction = static_cast<uint16_t>(DEG(_direction.value_or(0)))
    If has_direction Then direction = DEG(opt_direction) And &HFFFF& Else direction = DEG(0) And &HFFFF&

    psStat = structureIndex     ' STRUCTURE_STATS *psStat = &asStructureStats[structureIndex]
    psStruct = buildStructureDir(psStat, x, y, direction, player, False)
    If psStruct <> 0 Then
        aStruct(psStruct).status = SS_BUILT
        buildingComplete psStruct
        wzapi_addStructure = psStruct
        Exit Function
    End If
    wzapi_addStructure = 0
End Function

'-- ## getStructureLimit(structureName[, player])
'--
'-- Returns build limits for a structure.
'--
' (unsigned int: the Long's bits)
Public Function wzapi_getStructureLimit(ByVal context As wzapi_execution_context, ByVal structureName As String, ByVal has_player As Boolean, ByVal opt_player As Long) As Long
    Dim structureIndex As Long, player As Long
    structureIndex = getStructStatFromName(structureName)
    If Not (structureIndex >= 0 And structureIndex < numStructureStats) Then debug_ 3173, LOG_ERROR, "getStructureLimit", "Structure %s not found", structureName: context.throwError "structureIndex >= 0 && structureIndex < numStructureStats", 3173, "getStructureLimit": wzapi_getStructureLimit = 0: Exit Function
    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3175, LOG_ERROR, "getStructureLimit", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3175, "getStructureLimit": wzapi_getStructureLimit = 0: Exit Function
    wzapi_getStructureLimit = asStructureStats(structureIndex).upgrade(player).limit
End Function

' (VB) countStruct's and countDroid's player filter: playerFilter == i || playerFilter == ALL_PLAYERS
' || (playerFilter == ALLIES && aiCheckAlliances(i, me)) || (playerFilter == ENEMIES && !aiCheckAlliances(i, me))
Private Function countFilterMatches(ByVal playerFilter As Long, ByVal i As Long, ByVal me_ As Long) As Boolean
    If playerFilter = i Or playerFilter = ALL_PLAYERS Then
        countFilterMatches = True
    ElseIf playerFilter = ALLIES Then
        countFilterMatches = aiCheckAlliances(i, me_)
    ElseIf playerFilter = ENEMIES Then
        countFilterMatches = Not aiCheckAlliances(i, me_)
    End If
End Function

'-- ## countStruct(structureName[, playerFilter])
'--
'-- Count the number of structures of a given type.
'-- The playerFilter parameter can be a specific player, ```ALL_PLAYERS```, ```ALLIES``` or ```ENEMIES```.
'--
Public Function wzapi_countStruct(ByVal context As wzapi_execution_context, ByVal structureName As String, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long) As Long
    Dim structureIndex As Long, me_ As Long, playerFilter As Long, quantity As Long, i As Long
    structureIndex = getStructStatFromName(structureName)
    If Not (structureIndex >= 0 And structureIndex < numStructureStats) Then debug_ 3187, LOG_ERROR, "countStruct", "Structure %s not found", structureName: context.throwError "structureIndex >= 0 && structureIndex < numStructureStats", 3187, "countStruct": wzapi_countStruct = -1: Exit Function
    me_ = context.player()
    If has_playerFilter Then playerFilter = opt_playerFilter Else playerFilter = me_
    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALL_PLAYERS Or playerFilter = ALLIES Or playerFilter = ENEMIES) Then debug_ 3190, LOG_ERROR, "countStruct", "Player filter index out of range: %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALL_PLAYERS || playerFilter == ALLIES || playerFilter == ENEMIES", 3190, "countStruct": wzapi_countStruct = -1: Exit Function

    quantity = 0
    For i = 0 To MAX_PLAYERS - 1
        If countFilterMatches(playerFilter, i, me_) Then
            quantity = quantity + asStructureStats(structureIndex).curCount(i)
        End If
    Next
    wzapi_countStruct = quantity
End Function

'-- ## countDroid([droidType[, playerFilter]])
'--
'-- Count the number of droids that a given player has. Droid type must be either
'-- ```DROID_ANY```, ```DROID_COMMAND``` or ```DROID_CONSTRUCT```.
'-- The playerFilter parameter can be a specific player, ```ALL_PLAYERS```, ```ALLIES``` or ```ENEMIES```.
'--
Public Function wzapi_countDroid(ByVal context As wzapi_execution_context, ByVal has_droidType As Boolean, ByVal opt_droidType As Long, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long) As Long
    Dim droidType As Long, me_ As Long, playerFilter As Long, quantity As Long, i As Long
    If has_droidType Then droidType = opt_droidType Else droidType = DROID_ANY
    If Not (droidType <= DROID_ANY) Then debug_ 3214, LOG_ERROR, "countDroid", "Bad droid type parameter": context.throwError "droidType <= DROID_ANY", 3214, "countDroid": wzapi_countDroid = -1: Exit Function
    me_ = context.player()
    If has_playerFilter Then playerFilter = opt_playerFilter Else playerFilter = me_
    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALL_PLAYERS Or playerFilter = ALLIES Or playerFilter = ENEMIES) Then debug_ 3217, LOG_ERROR, "countDroid", "Player index out of range: %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALL_PLAYERS || playerFilter == ALLIES || playerFilter == ENEMIES", 3217, "countDroid": wzapi_countDroid = -1: Exit Function

    quantity = 0
    For i = 0 To MAX_PLAYERS - 1
        If countFilterMatches(playerFilter, i, me_) Then
            If droidType = DROID_ANY Then
                quantity = quantity + getNumDroids(i)
            ElseIf droidType = DROID_CONSTRUCT Then
                quantity = quantity + getNumConstructorDroids(i)
            ElseIf droidType = DROID_COMMAND Then
                quantity = quantity + getNumCommandDroids(i)
            End If
        End If
    Next
    wzapi_countDroid = quantity
End Function

'-- ## loadLevel(levelName)
'--
'-- Load the level with the given name.
'--
Public Sub wzapi_loadLevel(ByVal context As wzapi_execution_context, ByVal levelName As String)
    Dim psNewLevel As LEVEL_DATASET
    ' Find the level dataset
    Set psNewLevel = levFindDataSet(levelName)
    If psNewLevel Is Nothing Then debug_ 3251, LOG_ERROR, "loadLevel", "Could not find level data for %s", levelName: context.throwError "psNewLevel", 3251, "loadLevel": Exit Sub

    ' Get the mission rolling...
    aLevelName = Left$(levelName, MAX_LEVEL_NAME_SIZE)  ' sstrcpy(aLevelName, levelName.c_str()) (char[MAX_LEVEL_NAME_SIZE + 1])
    prevMissionType = mission_type
    nextMissionType = psNewLevel.typ
    loopMissionState = LMS_CLEAROBJECTS
End Sub

'-- ## setDroidExperience(droid, experience)
'--
'-- Set the amount of experience a droid has. Experience is read using floating point precision.
'--
Public Sub wzapi_setDroidExperience(ByVal context As wzapi_execution_context, ByVal psDroid As Long, ByVal experience As Double)
    If Not (psDroid <> 0) Then debug_ 3267, LOG_ERROR, "setDroidExperience", "No valid droid provided": context.throwError "psDroid", 3267, "setDroidExperience": Exit Sub
    aDroid(psDroid).experience = FromU32(Fix(experience * 65536#))   ' static_cast<uint32_t>(experience * 65536)
End Sub

'-- ## donateObject(object, player)
'--
'-- Donate a game object (restricted to droids before 3.2.3) to another player. Returns true if
'-- donation was successful. May return false if this donation would push the receiving player
'-- over unit limits. (3.2+ only)
'--
Public Function wzapi_donateObject(ByVal context As wzapi_execution_context, ByVal psObject As Long, ByVal player As Long) As Boolean
    Dim object_id As Long, from As Long, to_ As Long, giftType As Long, psDroid As Long, psStruct As Long
    Dim statidx As Long, overLimit As Boolean, w As MessageWriter
    If Not (psObject <> 0) Then debug_ 3280, LOG_ERROR, "donateObject", "No valid object provided": context.throwError "psObject", 3280, "donateObject": wzapi_donateObject = False: Exit Function
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3281, LOG_ERROR, "donateObject", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3281, "donateObject": wzapi_donateObject = False: Exit Function

    object_id = SIMPLE_OBJECT_id(psObject)
    from = SIMPLE_OBJECT_player(psObject) And &HFF&     ' uint8_t
    to_ = player And &HFF&                              ' static_cast<uint8_t>(player)
    giftType = 0
    If SIMPLE_OBJECT_type(psObject) = OBJ_DROID Then
        ' Check unit limits.
        psDroid = psObject
        overLimit = False
        If aDroid(psDroid).droidType = DROID_COMMAND Then
            If getNumCommandDroids(to_) + 1 > getMaxCommanders(to_) Then overLimit = True
        End If
        If Not overLimit Then
            If aDroid(psDroid).droidType = DROID_CONSTRUCT Then
                If getNumConstructorDroids(to_) + 1 > getMaxConstructors(to_) Then overLimit = True
            End If
        End If
        If Not overLimit Then
            If getNumDroids(to_) + 1 > getMaxDroids(to_) Then overLimit = True
        End If
        If overLimit Then
            wzapi_donateObject = False
            Exit Function
        End If
        giftType = DROID_GIFT
    ElseIf SIMPLE_OBJECT_type(psObject) = OBJ_STRUCTURE Then
        psStruct = psObject
        statidx = aStruct(psStruct).pStructureType      ' psStruct->pStructureType - asStructureStats
        ' (curCount + 1 > limit: unsigned)
        If ULT(asStructureStats(statidx).upgrade(to_).limit, asStructureStats(statidx).curCount(to_) + 1) Then
            wzapi_donateObject = False
            Exit Function
        End If
        giftType = STRUCTURE_GIFT
    Else
        wzapi_donateObject = False
        Exit Function
    End If
    Set w = NETbeginEncode(NETgameQueue(realSelectedPlayer), GAME_GIFT)
    NETuint8_t w, giftType
    NETuint8_t w, from
    NETuint8_t w, to_
    NETuint32_t w, object_id
    NETend w
    wzapi_donateObject = True
End Function

'-- ## donatePower(amount, player)
'--
'-- Donate power to another player. Returns true. (3.2+ only)
'--
Public Function wzapi_donatePower(ByVal context As wzapi_execution_context, ByVal amount As Long, ByVal player As Long) As Boolean
    Dim from As Long
    from = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3329, LOG_ERROR, "donatePower", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3329, "donatePower": wzapi_donatePower = False: Exit Function
    ' giftPower(uint8_t from, uint8_t to, uint32_t amount, bool send)
    giftPower from And &HFF&, player And &HFF&, amount, True
    wzapi_donatePower = True
End Function

'-- ## setNoGoArea(x1, y1, x2, y2, playerFilter)
'--
'-- Creates an area on the map on which nothing can be built. If playerFilter is zero,
'-- then landing lights are placed. If playerFilter is ```ALL_PLAYERS```, then a limbo landing zone
'-- is created and limbo droids placed.
'--
Public Sub wzapi_setNoGoArea(ByVal context As wzapi_execution_context, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal playerFilter As Long)
    If Not (x1 >= 0) Then debug_ 3342, LOG_ERROR, "setNoGoArea", "Minimum scroll x value %d is less than zero - ", x1: context.throwError "x1 >= 0", 3342, "setNoGoArea": Exit Sub
    If Not (y1 >= 0) Then debug_ 3343, LOG_ERROR, "setNoGoArea", "Minimum scroll y value %d is less than zero - ", y1: context.throwError "y1 >= 0", 3343, "setNoGoArea": Exit Sub
    If Not (x2 <= mapWidth) Then debug_ 3344, LOG_ERROR, "setNoGoArea", "Maximum scroll x value %d is greater than mapWidth %d", x2, mapWidth: context.throwError "x2 <= mapWidth", 3344, "setNoGoArea": Exit Sub
    If Not (y2 <= mapHeight) Then debug_ 3345, LOG_ERROR, "setNoGoArea", "Maximum scroll y value %d is greater than mapHeight %d", y2, mapHeight: context.throwError "y2 <= mapHeight", 3345, "setNoGoArea": Exit Sub
    If Not ((playerFilter >= 0 And playerFilter < MAX_PLAYERS) Or playerFilter = ALL_PLAYERS) Then debug_ 3346, LOG_ERROR, "setNoGoArea", "Bad player filter value %d", playerFilter: context.throwError "(playerFilter >= 0 && playerFilter < MAX_PLAYERS) || playerFilter == ALL_PLAYERS", 3346, "setNoGoArea": Exit Sub

    ' (::setNoGoArea takes UBYTEs)
    If playerFilter = ALL_PLAYERS Then
        setNoGoArea x1 And &HFF&, y1 And &HFF&, x2 And &HFF&, y2 And &HFF&, LIMBO_LANDING
        placeLimboDroids   ' this calls the Droids from the Limbo list onto the map
    Else
        setNoGoArea x1 And &HFF&, y1 And &HFF&, x2 And &HFF&, y2 And &HFF&, playerFilter And &HFF&
    End If
End Sub

'-- ## startTransporterEntry(x, y, player)
'--
'-- Set the entry position for the mission transporter, and make it start flying in
'-- reinforcements. If you want the camera to follow it in, use cameraTrack() on it.
'-- The transport needs to be set up with the mission droids, and the first transport
'-- found will be used. (3.2+ only)
'--
Public Sub wzapi_startTransporterEntry(ByVal context As wzapi_execution_context, ByVal x As Long, ByVal y As Long, ByVal player As Long)
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3369, LOG_ERROR, "startTransporterEntry", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3369, "startTransporterEntry": Exit Sub
    missionSetTransporterEntry player, x, y
    missionFlyTransportersIn player, False
End Sub

'-- ## setTransporterExit(x, y, player)
'--
'-- Set the exit position for the mission transporter. (3.2+ only)
'--
Public Sub wzapi_setTransporterExit(ByVal context As wzapi_execution_context, ByVal x As Long, ByVal y As Long, ByVal player As Long)
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3381, LOG_ERROR, "setTransporterExit", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3381, "setTransporterExit": Exit Sub
    missionSetTransporterExit player, x, y
End Sub

'-- ## setObjectFlag(object, flag, flagValue)
'--
'-- Set or unset an object flag on a given game object. Does not take care of network sync, so for multiplayer games,
'-- needs wrapping in a syncRequest. (3.3+ only.)
'-- Recognized object flags: ```OBJECT_FLAG_UNSELECTABLE``` - makes object unavailable for selection from player UI.
'--
Public Sub wzapi_setObjectFlag(ByVal context As wzapi_execution_context, ByVal psObj As Long, ByVal flag_ As Long, ByVal flagValue As Boolean)
    Dim typ As OBJECT_TYPE, flag As Long
    If Not (psObj <> 0) Then debug_ 3394, LOG_ERROR, "setObjectFlag", "No valid object provided": context.throwError "psObj", 3394, "setObjectFlag": Exit Sub
    typ = SIMPLE_OBJECT_type(psObj)
    If Not (typ = OBJ_DROID Or typ = OBJ_STRUCTURE Or typ = OBJ_FEATURE) Then debug_ 3395, LOG_ERROR, "setObjectFlag", "Bad object type": context.throwError "psObj->type == OBJ_DROID || psObj->type == OBJ_STRUCTURE || psObj->type == OBJ_FEATURE", 3395, "setObjectFlag": Exit Sub
    flag = flag_    ' (OBJECT_FLAG)_flag
    If Not (flag >= 0 And flag < OBJECT_FLAG_COUNT) Then debug_ 3397, LOG_ERROR, "setObjectFlag", "Bad flag value %d", flag: context.throwError "flag >= 0 && flag < OBJECT_FLAG_COUNT", 3397, "setObjectFlag": Exit Sub
    ' psObj->flags.set(flag, flagValue)
    If flagValue Then
        BASE_OBJECT_flags(psObj) = BASE_OBJECT_flags(psObj) Or BIT_(flag)
    Else
        BASE_OBJECT_flags(psObj) = BASE_OBJECT_flags(psObj) And Not BIT_(flag)
    End If
End Sub

'-- ## fireWeaponAtLoc(weaponName, x, y[, player[, center]])
'--
'-- Fires a weapon at the given coordinates (3.3+ only). The player is who owns the projectile.
'--
'-- Please use fireWeaponAtObj() to damage objects as multiplayer and campaign
'-- may have different friendly fire logic for a few weapons (like the lassat).
'--
'-- The optional ```center``` parameter (4.6.2+ only) can be set to ```true```
'-- to target the center of the tile. (The default is ```false```.)
'--
Public Sub wzapi_fireWeaponAtLoc(ByVal context As wzapi_execution_context, ByVal weaponName As String, ByVal x As Long, ByVal y As Long, ByVal has_player As Boolean, ByVal opt_player As Long, ByVal has_center As Boolean, ByVal opt_center As Boolean)
    Dim weaponIndex As Long, player As Long, target As Vector3i, sWeapon As WEAPON, center As Boolean
    weaponIndex = getCompFromName(COMP_WEAPON, weaponName)
    If Not (weaponIndex > 0) Then debug_ 3415, LOG_ERROR, "fireWeaponAtLoc", "No such weapon: %s", weaponName: context.throwError "weaponIndex > 0", 3415, "fireWeaponAtLoc": Exit Sub

    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3418, LOG_ERROR, "fireWeaponAtLoc", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3418, "fireWeaponAtLoc": Exit Sub

    target.x = world_coord(x)
    target.y = world_coord(y)
    If has_center Then center = opt_center Else center = False
    If center Then
        target.x = target.x + (TILE_UNITS \ 2)
        target.y = target.y + (TILE_UNITS \ 2)
    End If
    target.z = psMapTiles(mapTile(x, y)).height

    ' WEAPON sWeapon (its constructor zeroes the rest)
    sWeapon.nStat = weaponIndex

    proj_SendProjectile sWeapon, 0, player, target, 0, True, 0
End Sub

'-- ## fireWeaponAtObj(weaponName, gameObject[, player])
'--
'-- Fires a weapon at a game object (3.3+ only). The player is who owns the projectile.
'--
Public Sub wzapi_fireWeaponAtObj(ByVal context As wzapi_execution_context, ByVal weaponName As String, ByVal psObj As Long, ByVal has_player As Boolean, ByVal opt_player As Long)
    Dim weaponIndex As Long, player As Long, target As Vector3i, sWeapon As WEAPON
    weaponIndex = getCompFromName(COMP_WEAPON, weaponName)
    If Not (weaponIndex > 0) Then debug_ 3444, LOG_ERROR, "fireWeaponAtObj", "No such weapon: %s", weaponName: context.throwError "weaponIndex > 0", 3444, "fireWeaponAtObj": Exit Sub
    If Not (psObj <> 0) Then debug_ 3445, LOG_ERROR, "fireWeaponAtObj", "No valid object provided": context.throwError "psObj", 3445, "fireWeaponAtObj": Exit Sub

    If has_player Then player = opt_player Else player = context.player()
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3448, LOG_ERROR, "fireWeaponAtObj", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3448, "fireWeaponAtObj": Exit Sub

    target = SIMPLE_OBJECT_pos(psObj)

    ' WEAPON sWeapon (its constructor zeroes the rest)
    sWeapon.nStat = weaponIndex

    proj_SendProjectile sWeapon, 0, player, target, psObj, True, 0
End Sub

'-- ## setGameStoryLogPlayerDataValue(player, key_str, value_str)
'--
'-- Set a key + value string associated with a player, to be output to the playerData in the game story log (4.4+ only)
'-- (Intended to be called from endconditions.js to pass a useful string that describes the player state for display-only purposes.)
'-- (Not currently intended to be called from anywhere else.)
'--
Public Function wzapi_setGameStoryLogPlayerDataValue(ByVal context As wzapi_execution_context, ByVal player As Long, ByVal key_str As String, ByVal value_str As String) As Boolean
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3467, LOG_ERROR, "setGameStoryLogPlayerDataValue", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3467, "setGameStoryLogPlayerDataValue": wzapi_setGameStoryLogPlayerDataValue = False: Exit Function
    wzapi_setGameStoryLogPlayerDataValue = setGameStoryLogPlayerDataValue(player, key_str, value_str)
End Function

'-- ## transformPlayerToSpectator(player)
'--
'-- Transform a player to a spectator. (4.2+ only)
'-- This is a one-time transformation, destroys the player's HQ and all of their remaining units, and must occur deterministically on all clients.
'--
Public Function wzapi_transformPlayerToSpectator(ByVal context As wzapi_execution_context, ByVal player As Long) As Boolean
    If Not (player >= 0 And player < MAX_PLAYERS) Then debug_ 3478, LOG_ERROR, "transformPlayerToSpectator", "Invalid player index %d", player: context.throwError "player >= 0 && player < MAX_PLAYERS", 3478, "transformPlayerToSpectator": wzapi_transformPlayerToSpectator = False: Exit Function
    wzapi_transformPlayerToSpectator = makePlayerSpectator(player, False, False)
End Function

' flag all droids as requiring update on next frame
Private Sub dirtyAllDroids(ByVal player As Long)
    Dim it As Long
    it = apsDroidLists(player).head
    Do While it <> 0
        aDroid(lnObj(it)).flags = aDroid(lnObj(it)).flags Or BIT_(OBJECT_FLAG_DIRTY)
        it = lnNext(it)
    Loop
    it = mission_apsDroidLists(player).head
    Do While it <> 0
        aDroid(lnObj(it)).flags = aDroid(lnObj(it)).flags Or BIT_(OBJECT_FLAG_DIRTY)
        it = lnNext(it)
    Loop
    it = apsLimboDroids(player).head
    Do While it <> 0
        aDroid(lnObj(it)).flags = aDroid(lnObj(it)).flags Or BIT_(OBJECT_FLAG_DIRTY)
        it = lnNext(it)
    Loop
End Sub

Private Sub dirtyAllStructures(ByVal player As Long)
    Dim it As Long
    it = apsStructLists(player).head
    Do While it <> 0
        aStruct(lnObj(it)).flags = aStruct(lnObj(it)).flags Or BIT_(OBJECT_FLAG_DIRTY)
        it = lnNext(it)
    Loop
    it = mission_apsStructLists(player).head
    Do While it <> 0
        aStruct(lnObj(it)).flags = aStruct(lnObj(it)).flags Or BIT_(OBJECT_FLAG_DIRTY)
        it = lnNext(it)
    Loop
End Sub

Public Function wzapi_setUpgradeStats(ByVal player As Long, ByVal name_ As String, ByVal typ As Long, ByVal index As Long, ByVal newValue As nlohmann_json) As Boolean
    Dim value As Long, psStats As Long, length As Long, i As Long, it As Long, psCurr As Long
    value = json_variant_new(newValue).toInt()
    syncDebug "setUpgradeStats", "stats[p%d,t%d,%s,i%d] = %d", player, typ, name_, index, value
    If typ = COMP_BODY Then
        If Not (ULT(index, asBodyStats_size)) Then debug_ 3535, LOG_ERROR, "setUpgradeStats", "Bad index": wzapi_setUpgradeStats = False: Exit Function
        psStats = index
        If name_ = "HitPoints" Then
            asBodyStats(psStats).upgrade(player).hitpoints = value
            dirtyAllDroids player
        ElseIf name_ = "HitPointPct" Then
            asBodyStats(psStats).upgrade(player).hitpointPct = value
            dirtyAllDroids player
        ElseIf name_ = "Armour" Then
            asBodyStats(psStats).upgrade(player).armour = value
        ElseIf name_ = "Thermal" Then
            asBodyStats(psStats).upgrade(player).thermal = value
        ElseIf name_ = "Power" Then
            asBodyStats(psStats).upgrade(player).power = value
            dirtyAllDroids player
        ElseIf name_ = "Resistance" Then
            ' TBD FIXME - not updating resistance points in droids...
            asBodyStats(psStats).upgrade(player).resistance = value
        Else
            debug_ 3567, LOG_ERROR, "setUpgradeStats", "Upgrade component %s not found", name_: wzapi_setUpgradeStats = False: Exit Function
        End If
    ElseIf typ = COMP_BRAIN Then
        If Not (ULT(index, asBrainStats_size)) Then debug_ 3572, LOG_ERROR, "setUpgradeStats", "Bad index": wzapi_setUpgradeStats = False: Exit Function
        psStats = index
        If name_ = "BaseCommandLimit" Then
            asBrainStats(psStats).upgrade(player).maxDroids = value
        ElseIf name_ = "CommandLimitByLevel" Then
            asBrainStats(psStats).upgrade(player).maxDroidsMult = value
        ElseIf name_ = "RankThresholds" Then
            If Not (newValue.is_array()) Then debug_ 3584, LOG_ERROR, "setUpgradeStats", "Level thresholds not an array!": wzapi_setUpgradeStats = False: Exit Function
            length = newValue.size()
            If Not (length <= asBrainStats(psStats).upgrade(player).rankThresholds_size) Then debug_ 3586, LOG_ERROR, "setUpgradeStats", "Invalid thresholds array length": wzapi_setUpgradeStats = False: Exit Function
            For i = 0 To length - 1
                ' Use json_variant to support conversion from other value types to an int
                asBrainStats(psStats).upgrade(player).rankThresholds(i) = json_variant_new(newValue.at_idx(i)).toInt()
            Next
        ElseIf name_ = "HitPoints" Then
            asBrainStats(psStats).upgrade(player).hitpoints = value
            dirtyAllDroids player
        ElseIf name_ = "HitPointPct" Then
            asBrainStats(psStats).upgrade(player).hitpointPct = value
            dirtyAllDroids player
        Else
            debug_ 3605, LOG_ERROR, "setUpgradeStats", "Upgrade component %s not found", name_: wzapi_setUpgradeStats = False: Exit Function
        End If
    ElseIf typ = COMP_SENSOR Then
        If Not (ULT(index, asSensorStats_size)) Then debug_ 3610, LOG_ERROR, "setUpgradeStats", "Bad index": wzapi_setUpgradeStats = False: Exit Function
        psStats = index
        If name_ = "Range" Then
            asSensorStats(psStats).upgrade(player).range = value
            dirtyAllDroids player
            dirtyAllStructures player
        ElseIf name_ = "HitPoints" Then
            asSensorStats(psStats).upgrade(player).hitpoints = value
            dirtyAllDroids player
        ElseIf name_ = "HitPointPct" Then
            asSensorStats(psStats).upgrade(player).hitpointPct = value
            dirtyAllDroids player
        Else
            debug_ 3630, LOG_ERROR, "setUpgradeStats", "Upgrade component %s not found", name_: wzapi_setUpgradeStats = False: Exit Function
        End If
    ElseIf typ = COMP_ECM Then
        If Not (ULT(index, asECMStats_size)) Then debug_ 3635, LOG_ERROR, "setUpgradeStats", "Bad index": wzapi_setUpgradeStats = False: Exit Function
        psStats = index
        If name_ = "Range" Then
            asECMStats(psStats).upgrade(player).range = value
            dirtyAllDroids player
            dirtyAllStructures player
        ElseIf name_ = "HitPoints" Then
            asECMStats(psStats).upgrade(player).hitpoints = value
            dirtyAllDroids player
        ElseIf name_ = "HitPointPct" Then
            asECMStats(psStats).upgrade(player).hitpointPct = value
            dirtyAllDroids player
        Else
            debug_ 3655, LOG_ERROR, "setUpgradeStats", "Upgrade component %s not found", name_: wzapi_setUpgradeStats = False: Exit Function
        End If
    ElseIf typ = COMP_PROPULSION Then
        If Not (ULT(index, asPropulsionStats_size)) Then debug_ 3660, LOG_ERROR, "setUpgradeStats", "Bad index": wzapi_setUpgradeStats = False: Exit Function
        psStats = index
        If name_ = "HitPoints" Then
            asPropulsionStats(psStats).upgrade(player).hitpoints = value
            dirtyAllDroids player
        ElseIf name_ = "HitPointPct" Then
            asPropulsionStats(psStats).upgrade(player).hitpointPct = value
            dirtyAllDroids player
        ElseIf name_ = "HitPointPctOfBody" Then
            asPropulsionStats(psStats).upgrade(player).hitpointPctOfBody = value
            dirtyAllDroids player
        Else
            debug_ 3679, LOG_ERROR, "setUpgradeStats", "Upgrade component %s not found", name_: wzapi_setUpgradeStats = False: Exit Function
        End If
    ElseIf typ = COMP_CONSTRUCT Then
        If Not (ULT(index, asConstructStats_size)) Then debug_ 3684, LOG_ERROR, "setUpgradeStats", "Bad index": wzapi_setUpgradeStats = False: Exit Function
        psStats = index
        If name_ = "ConstructorPoints" Then
            asConstructStats(psStats).upgrade(player).constructPoints = value
        ElseIf name_ = "HitPoints" Then
            asConstructStats(psStats).upgrade(player).hitpoints = value
            dirtyAllDroids player
        ElseIf name_ = "HitPointPct" Then
            asConstructStats(psStats).upgrade(player).hitpointPct = value
            dirtyAllDroids player
        Else
            debug_ 3702, LOG_ERROR, "setUpgradeStats", "Upgrade component %s not found", name_: wzapi_setUpgradeStats = False: Exit Function
        End If
    ElseIf typ = COMP_REPAIRUNIT Then
        If Not (ULT(index, asRepairStats_size)) Then debug_ 3707, LOG_ERROR, "setUpgradeStats", "Bad index": wzapi_setUpgradeStats = False: Exit Function
        psStats = index
        If name_ = "RepairPoints" Then
            asRepairStats(psStats).upgrade(player).repairPoints = value
        ElseIf name_ = "HitPoints" Then
            asRepairStats(psStats).upgrade(player).hitpoints = value
            dirtyAllDroids player
        ElseIf name_ = "HitPointPct" Then
            asRepairStats(psStats).upgrade(player).hitpointPct = value
            dirtyAllDroids player
        Else
            debug_ 3725, LOG_ERROR, "setUpgradeStats", "Upgrade component %s not found", name_: wzapi_setUpgradeStats = False: Exit Function
        End If
    ElseIf typ = COMP_WEAPON Then
        If Not (ULT(index, asWeaponStats_size)) Then debug_ 3730, LOG_ERROR, "setUpgradeStats", "Bad index": wzapi_setUpgradeStats = False: Exit Function
        psStats = index
        If name_ = "MaxRange" Then
            asWeaponStats(psStats).upgrade(player).maxRange = value
        ElseIf name_ = "ShortRange" Then
            asWeaponStats(psStats).upgrade(player).shortRange = value
        ElseIf name_ = "MinRange" Then
            asWeaponStats(psStats).upgrade(player).minRange = value
        ElseIf name_ = "HitChance" Then
            asWeaponStats(psStats).upgrade(player).hitChance = value
        ElseIf name_ = "ShortHitChance" Then
            asWeaponStats(psStats).upgrade(player).shortHitChance = value
        ElseIf name_ = "FirePause" Then
            asWeaponStats(psStats).upgrade(player).firePause = value
        ElseIf name_ = "Rounds" Then
            asWeaponStats(psStats).upgrade(player).numRounds = value And &HFF&      ' uint8_t
        ElseIf name_ = "ReloadTime" Then
            asWeaponStats(psStats).upgrade(player).reloadTime = value
        ElseIf name_ = "Damage" Then
            asWeaponStats(psStats).upgrade(player).damage = value
        ElseIf name_ = "MinimumDamage" Then
            asWeaponStats(psStats).upgrade(player).minimumDamage = value
        ElseIf name_ = "EmpRadius" Then
            asWeaponStats(psStats).upgrade(player).empRadius = value
        ElseIf name_ = "Radius" Then
            asWeaponStats(psStats).upgrade(player).radius = value
        ElseIf name_ = "RadiusDamage" Then
            asWeaponStats(psStats).upgrade(player).radiusDamage = value
        ElseIf name_ = "RepeatDamage" Then
            asWeaponStats(psStats).upgrade(player).periodicalDamage = value
        ElseIf name_ = "RepeatTime" Then
            asWeaponStats(psStats).upgrade(player).periodicalDamageTime = value
        ElseIf name_ = "RepeatRadius" Then
            asWeaponStats(psStats).upgrade(player).periodicalDamageRadius = value
        ElseIf name_ = "HitPoints" Then
            asWeaponStats(psStats).upgrade(player).hitpoints = value
            dirtyAllDroids player
        ElseIf name_ = "HitPointPct" Then
            asWeaponStats(psStats).upgrade(player).hitpointPct = value
            dirtyAllDroids player
        Else
            debug_ 3808, LOG_ERROR, "setUpgradeStats", "Invalid weapon method": wzapi_setUpgradeStats = False: Exit Function
        End If
    ElseIf typ >= SCRCB_FIRST And typ <= SCRCB_LAST Then
        If Not (ULT(index, numStructureStats)) Then debug_ 3813, LOG_ERROR, "setUpgradeStats", "Bad index": wzapi_setUpgradeStats = False: Exit Function
        psStats = index
        Select Case typ
        Case SCRCB_RES: asStructureStats(psStats).upgrade(player).research = value
        Case SCRCB_MODULE_RES: asStructureStats(psStats).upgrade(player).moduleResearch = value
        Case SCRCB_REP: asStructureStats(psStats).upgrade(player).repair = value
        Case SCRCB_POW: asStructureStats(psStats).upgrade(player).power = value
        Case SCRCB_MODULE_POW: asStructureStats(psStats).upgrade(player).modulePower = value
        Case SCRCB_CON: asStructureStats(psStats).upgrade(player).production = value
        Case SCRCB_MODULE_CON: asStructureStats(psStats).upgrade(player).moduleProduction = value
        Case SCRCB_REA: asStructureStats(psStats).upgrade(player).rearm = value
        Case SCRCB_HEA: asStructureStats(psStats).upgrade(player).thermal = value
        Case SCRCB_ARM: asStructureStats(psStats).upgrade(player).armour = value
        Case SCRCB_ELW
            ' Update resistance points for all structures, to avoid making them damaged
            ' FIXME - this is _really_ slow! we could be doing this for dozens of buildings one at a time!
            it = apsStructLists(player).head
            Do While it <> 0
                psCurr = lnObj(it)
                If psStats = aStruct(psCurr).pStructureType And ULT(asStructureStats(psStats).upgrade(player).resistance, value) Then
                    aStruct(psCurr).resistance = value
                End If
                it = lnNext(it)
            Loop
            it = mission_apsStructLists(player).head
            Do While it <> 0
                psCurr = lnObj(it)
                If psStats = aStruct(psCurr).pStructureType And ULT(asStructureStats(psStats).upgrade(player).resistance, value) Then
                    aStruct(psCurr).resistance = value
                End If
                it = lnNext(it)
            Loop
            asStructureStats(psStats).upgrade(player).resistance = value
        Case SCRCB_HIT
            ' Update body points for all structures, to avoid making them damaged
            ' FIXME - this is _really_ slow! we could be doing this for
            ' dozens of buildings one at a time!
            it = apsStructLists(player).head
            Do While it <> 0
                psCurr = lnObj(it)
                If psStats = aStruct(psCurr).pStructureType And (Not bMultiPlayer Or (bMultiPlayer And ULT(asStructureStats(psStats).upgrade(player).hitpoints, value))) Then
                    aStruct(psCurr).body = FromU32(Fix(U32(MUL32(aStruct(psCurr).body, value)) / U32(asStructureStats(psStats).upgrade(player).hitpoints)))
                End If
                it = lnNext(it)
            Loop
            it = mission_apsStructLists(player).head
            Do While it <> 0
                psCurr = lnObj(it)
                If psStats = aStruct(psCurr).pStructureType And (Not bMultiPlayer Or (bMultiPlayer And ULT(asStructureStats(psStats).upgrade(player).hitpoints, value))) Then
                    aStruct(psCurr).body = FromU32(Fix(U32(MUL32(aStruct(psCurr).body, value)) / U32(asStructureStats(psStats).upgrade(player).hitpoints)))
                End If
                it = lnNext(it)
            Loop
            asStructureStats(psStats).upgrade(player).hitpoints = value
        Case SCRCB_LIMIT
            asStructureStats(psStats).upgrade(player).limit = value
        End Select
    Else
        debug_ 3872, LOG_ERROR, "setUpgradeStats", "Component type not found for upgrade": wzapi_setUpgradeStats = False: Exit Function
    End If

    wzapi_setUpgradeStats = True
End Function

' (VB) json(int) and json(unsigned), for getUpgradeStats and constructStatsObject
Private Function jsonInt(ByVal v As Long) As nlohmann_json
    Set jsonInt = json_value(v)
End Function

Private Function jsonUnsigned(ByVal v As Long) As nlohmann_json
    Set jsonUnsigned = json_value_unsigned(v)
End Function

' (VB) nlohmann::json(): null
Private Function jsonNull() As nlohmann_json
    Set jsonNull = json_value(Empty)
End Function

Public Function wzapi_getUpgradeStats(ByVal player As Long, ByVal name_ As String, ByVal typ As Long, ByVal index As Long) As nlohmann_json
    Dim psStats As Long, length As Long, i As Long, value As nlohmann_json
    If typ = COMP_BODY Then
        If Not (ULT(index, asBodyStats_size)) Then debug_ 3882, LOG_ERROR, "getUpgradeStats", "Bad index": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        psStats = index
        If name_ = "HitPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asBodyStats(psStats).upgrade(player).hitpoints)
            Exit Function
        ElseIf name_ = "HitPointPct" Then
            Set wzapi_getUpgradeStats = jsonInt(asBodyStats(psStats).upgrade(player).hitpointPct)
            Exit Function
        ElseIf name_ = "Armour" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asBodyStats(psStats).upgrade(player).armour)
            Exit Function
        ElseIf name_ = "Thermal" Then
            Set wzapi_getUpgradeStats = jsonInt(asBodyStats(psStats).upgrade(player).thermal)
            Exit Function
        ElseIf name_ = "Power" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asBodyStats(psStats).upgrade(player).power)
            Exit Function
        ElseIf name_ = "Resistance" Then
            Set wzapi_getUpgradeStats = jsonInt(asBodyStats(psStats).upgrade(player).resistance)
            Exit Function
        Else
            debug_ 3910, LOG_ERROR, "getUpgradeStats", "Upgrade component %s not found", name_: Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        End If
    ElseIf typ = COMP_BRAIN Then
        If Not (ULT(index, asBrainStats_size)) Then debug_ 3915, LOG_ERROR, "getUpgradeStats", "Bad index": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        psStats = index
        If name_ = "BaseCommandLimit" Then
            Set wzapi_getUpgradeStats = jsonInt(asBrainStats(psStats).upgrade(player).maxDroids)
            Exit Function
        ElseIf name_ = "CommandLimitByLevel" Then
            Set wzapi_getUpgradeStats = jsonInt(asBrainStats(psStats).upgrade(player).maxDroidsMult)
            Exit Function
        ElseIf name_ = "RankThresholds" Then
            length = asBrainStats(psStats).upgrade(player).rankThresholds_size
            Set value = json_array()
            For i = 0 To length - 1
                value.push_back_ref jsonInt(asBrainStats(psStats).upgrade(player).rankThresholds(i))
            Next
            Set wzapi_getUpgradeStats = value
            Exit Function
        ElseIf name_ = "HitPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asBrainStats(psStats).upgrade(player).hitpoints)
            Exit Function
        ElseIf name_ = "HitPointPct" Then
            Set wzapi_getUpgradeStats = jsonInt(asBrainStats(psStats).upgrade(player).hitpointPct)
            Exit Function
        Else
            debug_ 3945, LOG_ERROR, "getUpgradeStats", "Upgrade component %s not found", name_: Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        End If
    ElseIf typ = COMP_SENSOR Then
        If Not (ULT(index, asSensorStats_size)) Then debug_ 3950, LOG_ERROR, "getUpgradeStats", "Bad index": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        psStats = index
        If name_ = "Range" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asSensorStats(psStats).upgrade(player).range)
            Exit Function
        ElseIf name_ = "HitPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asSensorStats(psStats).upgrade(player).hitpoints)
            Exit Function
        ElseIf name_ = "HitPointPct" Then
            Set wzapi_getUpgradeStats = jsonInt(asSensorStats(psStats).upgrade(player).hitpointPct)
            Exit Function
        Else
            debug_ 3966, LOG_ERROR, "getUpgradeStats", "Upgrade component %s not found", name_: Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        End If
    ElseIf typ = COMP_ECM Then
        If Not (ULT(index, asECMStats_size)) Then debug_ 3971, LOG_ERROR, "getUpgradeStats", "Bad index": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        psStats = index
        If name_ = "Range" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asECMStats(psStats).upgrade(player).range)
            Exit Function
        ElseIf name_ = "HitPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asECMStats(psStats).upgrade(player).hitpoints)
            Exit Function
        ElseIf name_ = "HitPointPct" Then
            Set wzapi_getUpgradeStats = jsonInt(asECMStats(psStats).upgrade(player).hitpointPct)
            Exit Function
        Else
            debug_ 3987, LOG_ERROR, "getUpgradeStats", "Upgrade component %s not found", name_: Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        End If
    ElseIf typ = COMP_PROPULSION Then
        If Not (ULT(index, asPropulsionStats_size)) Then debug_ 3992, LOG_ERROR, "getUpgradeStats", "Bad index": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        psStats = index
        If name_ = "HitPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asPropulsionStats(psStats).upgrade(player).hitpoints)
            Exit Function
        ElseIf name_ = "HitPointPct" Then
            Set wzapi_getUpgradeStats = jsonInt(asPropulsionStats(psStats).upgrade(player).hitpointPct)
            Exit Function
        ElseIf name_ = "HitPointPctOfBody" Then
            Set wzapi_getUpgradeStats = jsonInt(asPropulsionStats(psStats).upgrade(player).hitpointPctOfBody)
            Exit Function
        Else
            debug_ 4008, LOG_ERROR, "getUpgradeStats", "Upgrade component %s not found", name_: Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        End If
    ElseIf typ = COMP_CONSTRUCT Then
        If Not (ULT(index, asConstructStats_size)) Then debug_ 4013, LOG_ERROR, "getUpgradeStats", "Bad index": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        psStats = index
        If name_ = "ConstructorPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asConstructStats(psStats).upgrade(player).constructPoints)
            Exit Function
        ElseIf name_ = "HitPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asConstructStats(psStats).upgrade(player).hitpoints)
            Exit Function
        ElseIf name_ = "HitPointPct" Then
            Set wzapi_getUpgradeStats = jsonInt(asConstructStats(psStats).upgrade(player).hitpointPct)
            Exit Function
        Else
            debug_ 4029, LOG_ERROR, "getUpgradeStats", "Upgrade component %s not found", name_: Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        End If
    ElseIf typ = COMP_REPAIRUNIT Then
        If Not (ULT(index, asRepairStats_size)) Then debug_ 4034, LOG_ERROR, "getUpgradeStats", "Bad index": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        psStats = index
        If name_ = "RepairPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asRepairStats(psStats).upgrade(player).repairPoints)
            Exit Function
        ElseIf name_ = "HitPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asRepairStats(psStats).upgrade(player).hitpoints)
            Exit Function
        ElseIf name_ = "HitPointPct" Then
            Set wzapi_getUpgradeStats = jsonInt(asRepairStats(psStats).upgrade(player).hitpointPct)
            Exit Function
        Else
            debug_ 4050, LOG_ERROR, "getUpgradeStats", "Upgrade component %s not found", name_: Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        End If
    ElseIf typ = COMP_WEAPON Then
        If Not (ULT(index, asWeaponStats_size)) Then debug_ 4055, LOG_ERROR, "getUpgradeStats", "Bad index": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        psStats = index
        If name_ = "MaxRange" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).maxRange)
            Exit Function
        ElseIf name_ = "ShortRange" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).shortRange)
            Exit Function
        ElseIf name_ = "MinRange" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).minRange)
            Exit Function
        ElseIf name_ = "HitChance" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).hitChance)
            Exit Function
        ElseIf name_ = "ShortHitChance" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).shortHitChance)
            Exit Function
        ElseIf name_ = "FirePause" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).firePause)
            Exit Function
        ElseIf name_ = "Rounds" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).numRounds)
            Exit Function
        ElseIf name_ = "ReloadTime" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).reloadTime)
            Exit Function
        ElseIf name_ = "Damage" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).damage)
            Exit Function
        ElseIf name_ = "MinimumDamage" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).minimumDamage)
            Exit Function
        ElseIf name_ = "EmpRadius" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).empRadius)
            Exit Function
        ElseIf name_ = "Radius" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).radius)
            Exit Function
        ElseIf name_ = "RadiusDamage" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).radiusDamage)
            Exit Function
        ElseIf name_ = "RepeatDamage" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).periodicalDamage)
            Exit Function
        ElseIf name_ = "RepeatTime" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).periodicalDamageTime)
            Exit Function
        ElseIf name_ = "RepeatRadius" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).periodicalDamageRadius)
            Exit Function
        ElseIf name_ = "HitPoints" Then
            Set wzapi_getUpgradeStats = jsonUnsigned(asWeaponStats(psStats).upgrade(player).hitpoints)
            Exit Function
        ElseIf name_ = "HitPointPct" Then
            Set wzapi_getUpgradeStats = jsonInt(asWeaponStats(psStats).upgrade(player).hitpointPct)
            Exit Function
        Else
            debug_ 4131, LOG_ERROR, "getUpgradeStats", "Invalid weapon method": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        End If
    ElseIf typ >= SCRCB_FIRST And typ <= SCRCB_LAST Then
        If Not (ULT(index, numStructureStats)) Then debug_ 4136, LOG_ERROR, "getUpgradeStats", "Bad index": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        psStats = index
        Select Case typ
        Case SCRCB_RES: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).research): Exit Function
        Case SCRCB_MODULE_RES: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).moduleResearch): Exit Function
        Case SCRCB_REP: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).repair): Exit Function
        Case SCRCB_POW: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).power): Exit Function
        Case SCRCB_MODULE_POW: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).modulePower): Exit Function
        Case SCRCB_CON: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).production): Exit Function
        Case SCRCB_MODULE_CON: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).moduleProduction): Exit Function
        Case SCRCB_REA: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).rearm): Exit Function
        Case SCRCB_ELW: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).resistance): Exit Function
        Case SCRCB_HEA: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).thermal): Exit Function
        Case SCRCB_ARM: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).armour): Exit Function
        Case SCRCB_HIT: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).hitpoints): Exit Function
        Case SCRCB_LIMIT: Set wzapi_getUpgradeStats = jsonUnsigned(asStructureStats(psStats).upgrade(player).limit): Exit Function
        Case Else: debug_ 4153, LOG_ERROR, "getUpgradeStats", "Component type not found for upgrade": Set wzapi_getUpgradeStats = jsonNull(): Exit Function
        End Select
    End If
    Set wzapi_getUpgradeStats = jsonNull()
End Function

' GameEntityRules(int player, unsigned index, const NameToTypeMap& nameToTypeMap)
Private Function GameEntityRules_construct(ByVal player As Long, ByVal index As Long, ByVal nameToTypeMap As Long) As GameEntityRules
    GameEntityRules_construct.player = player
    GameEntityRules_construct.index = index
    GameEntityRules_construct.propertyNameToTypeMap = nameToTypeMap
End Function

' (VB) propertyNameToTypeMap.find(name): the type, or False (end())
Private Function NameToTypeMap_find(ByVal m As Long, ByVal name_ As String, typ As Long) As Boolean
    Dim i As Long
    For i = 0 To nameToTypeMaps(m).size - 1
        If nameToTypeMaps(m).keys(i) = name_ Then
            typ = nameToTypeMaps(m).values(i)
            NameToTypeMap_find = True
            Exit Function
        End If
    Next
    NameToTypeMap_find = False
End Function

Public Function GameEntityRules_getPropertyValue(this As GameEntityRules, ByVal name_ As String) As nlohmann_json
    Dim typ As Long
    If Not NameToTypeMap_find(this.propertyNameToTypeMap, name_, typ) Then
        ' Failed to find `name`
        Set GameEntityRules_getPropertyValue = jsonNull()
        Exit Function
    End If
    Set GameEntityRules_getPropertyValue = wzapi_getUpgradeStats(this.player, name_, typ, this.index)
End Function

Public Function GameEntityRules_setPropertyValue(this As GameEntityRules, ByVal name_ As String, ByVal newValue As nlohmann_json) As nlohmann_json
    Dim typ As Long
    If Not NameToTypeMap_find(this.propertyNameToTypeMap, name_, typ) Then
        ' Failed to find `name`
        Set GameEntityRules_setPropertyValue = jsonNull()
        Exit Function
    End If
    Set GameEntityRules_setPropertyValue = json_value(wzapi_setUpgradeStats(this.player, name_, typ, this.index, newValue))
End Function

' (VB) GameEntityRules' begin() / end(): the properties of its NameToTypeMap (a std::map<std::string, int>:
' by name, ascending, as std::string's < compares the characters' codes) in names(0 .. n - 1) and types(), and
' their number n
Public Function GameEntityRules_properties(this As GameEntityRules, names() As String, types() As Long) As Long
    Dim m As Long, n As Long, i As Long, j As Long, s As String, t As Long
    BuildNameToTypeMaps
    m = this.propertyNameToTypeMap
    n = nameToTypeMaps(m).size
    If n = 0 Then Exit Function
    ReDim names(0 To n - 1)
    ReDim types(0 To n - 1)
    For i = 0 To n - 1
        names(i) = nameToTypeMaps(m).keys(i)
        types(i) = nameToTypeMaps(m).values(i)
    Next
    ' (the initializer list's order sorted: an insertion sort)
    For i = 1 To n - 1
        s = names(i)
        t = types(i)
        j = i - 1
        Do While j >= 0
            If StrComp(names(j), s, vbBinaryCompare) <= 0 Then Exit Do
            names(j + 1) = names(j)
            types(j + 1) = types(j)
            j = j - 1
        Loop
        names(j + 1) = s
        types(j + 1) = t
    Next
    GameEntityRules_properties = n
End Function

Private Sub GameEntityRuleContainer_addRules(this As GameEntityRuleContainer, ByVal statsName As String, entityRules As GameEntityRules)
    If this.lookup_table Is Nothing Then
        Set this.lookup_table = CreateObject("Scripting.Dictionary")
        this.lookup_table.CompareMode = vbBinaryCompare
    End If
    ' (VB) the vector's capacity doubles
    If this.rules_size = 0 Then
        ReDim this.rules(0 To 15)
    ElseIf this.rules_size > UBound(this.rules) Then
        ReDim Preserve this.rules(0 To 2 * this.rules_size - 1)
    End If
    this.rules(this.rules_size).first = statsName
    this.rules(this.rules_size).second = entityRules
    this.rules_size = this.rules_size + 1
    this.lookup_table.Item(statsName) = this.rules_size - 1
End Sub

' PerPlayerUpgrades(int player)
Private Sub PerPlayerUpgrades_construct(this As PerPlayerUpgrades, ByVal player As Long)
    this.player = player
    this.upgrades_size = 0
    Erase this.upgrades_keys
    Erase this.upgrades_values
End Sub

' addGameEntity: upgrades.emplace(entityClass, rulesContainer), which keeps an existing entry
Private Sub PerPlayerUpgrades_addGameEntity(this As PerPlayerUpgrades, ByVal entityClass As String, rulesContainer As GameEntityRuleContainer)
    Dim i As Long, k As Long
    i = 0
    Do While i < this.upgrades_size
        If this.upgrades_keys(i) >= entityClass Then Exit Do
        i = i + 1
    Loop
    If i < this.upgrades_size Then
        If this.upgrades_keys(i) = entityClass Then Exit Sub
    End If
    ReDim Preserve this.upgrades_keys(0 To this.upgrades_size)
    ReDim Preserve this.upgrades_values(0 To this.upgrades_size)
    For k = this.upgrades_size To i + 1 Step -1
        this.upgrades_keys(k) = this.upgrades_keys(k - 1)
        this.upgrades_values(k) = this.upgrades_values(k - 1)
    Next
    this.upgrades_keys(i) = entityClass
    this.upgrades_values(i) = rulesContainer
    this.upgrades_size = this.upgrades_size + 1
End Sub

' find(entityClass): the container's index in upgrades_values, -1 = nullptr
Public Function PerPlayerUpgrades_find(this As PerPlayerUpgrades, ByVal entityClass As String) As Long
    Dim i As Long
    For i = 0 To this.upgrades_size - 1
        If this.upgrades_keys(i) = entityClass Then
            PerPlayerUpgrades_find = i
            Exit Function
        End If
    Next
    PerPlayerUpgrades_find = -1
End Function

' (VB) a NameToTypeMap from getUpgradesObject's initializer list (name, type, name, type, ...)
Private Sub BuildNameToTypeMap(ByVal m As Long, ParamArray pairs() As Variant)
    Dim i As Long, n As Long
    n = (UBound(pairs) + 1) \ 2
    ReDim nameToTypeMaps(m).keys(0 To n - 1)
    ReDim nameToTypeMaps(m).values(0 To n - 1)
    For i = 0 To n - 1
        nameToTypeMaps(m).keys(i) = pairs(2 * i)
        nameToTypeMaps(m).values(i) = pairs(2 * i + 1)
    Next
    nameToTypeMaps(m).size = n
End Sub

Private Sub BuildNameToTypeMaps()
    If nameToTypeMaps_built Then Exit Sub
    BuildNameToTypeMap NTT_BODY, _
        "HitPoints", COMP_BODY, _
        "HitPointPct", COMP_BODY, _
        "Power", COMP_BODY, _
        "Armour", COMP_BODY, _
        "Thermal", COMP_BODY, _
        "Resistance", COMP_BODY
    BuildNameToTypeMap NTT_SENSOR, _
        "HitPoints", COMP_SENSOR, _
        "HitPointPct", COMP_SENSOR, _
        "Range", COMP_SENSOR
    BuildNameToTypeMap NTT_PROPULSION, _
        "HitPoints", COMP_PROPULSION, _
        "HitPointPct", COMP_PROPULSION, _
        "HitPointPctOfBody", COMP_PROPULSION
    BuildNameToTypeMap NTT_ECM, _
        "Range", COMP_ECM, _
        "HitPoints", COMP_ECM, _
        "HitPointPct", COMP_ECM
    BuildNameToTypeMap NTT_REPAIR, _
        "RepairPoints", COMP_REPAIRUNIT, _
        "HitPoints", COMP_REPAIRUNIT, _
        "HitPointPct", COMP_REPAIRUNIT
    BuildNameToTypeMap NTT_CONSTRUCT, _
        "ConstructorPoints", COMP_CONSTRUCT, _
        "HitPoints", COMP_CONSTRUCT, _
        "HitPointPct", COMP_CONSTRUCT
    BuildNameToTypeMap NTT_BRAIN, _
        "BaseCommandLimit", COMP_BRAIN, _
        "CommandLimitByLevel", COMP_BRAIN, _
        "RankThresholds", COMP_BRAIN, _
        "HitPoints", COMP_BRAIN, _
        "HitPointPct", COMP_BRAIN
    BuildNameToTypeMap NTT_WEAPON, _
        "MaxRange", COMP_WEAPON, "ShortRange", COMP_WEAPON, "MinRange", COMP_WEAPON, _
        "HitChance", COMP_WEAPON, "ShortHitChance", COMP_WEAPON, "FirePause", COMP_WEAPON, _
        "ReloadTime", COMP_WEAPON, "Rounds", COMP_WEAPON, "Radius", COMP_WEAPON, _
        "Damage", COMP_WEAPON, "MinimumDamage", COMP_WEAPON, "RadiusDamage", COMP_WEAPON, _
        "RepeatDamage", COMP_WEAPON, "RepeatTime", COMP_WEAPON, "RepeatRadius", COMP_WEAPON, _
        "HitPoints", COMP_WEAPON, "HitPointPct", COMP_WEAPON, "EmpRadius", COMP_WEAPON
    BuildNameToTypeMap NTT_BUILDING, _
        "ResearchPoints", SCRCB_RES, "ModuleResearchPoints", SCRCB_MODULE_RES, _
        "RepairPoints", SCRCB_REP, "PowerPoints", SCRCB_POW, _
        "ModulePowerPoints", SCRCB_MODULE_POW, "ProductionPoints", SCRCB_CON, _
        "ModuleProductionPoints", SCRCB_MODULE_CON, "RearmPoints", SCRCB_REA, _
        "Armour", SCRCB_ARM, "Resistance", SCRCB_ELW, _
        "Thermal", SCRCB_HEA, "HitPoints", SCRCB_HIT, _
        "Limit", SCRCB_LIMIT
    nameToTypeMaps_built = True
End Sub

' std::vector<wzapi::PerPlayerUpgrades> getUpgradesObject(): the players' upgrades in upgrades(0 ..), and their count
Public Function wzapi_getUpgradesObject(upgrades() As PerPlayerUpgrades) As Long
    Dim i As Long, j As Long, n As Long
    Dim bodybase As GameEntityRuleContainer, sensorbase As GameEntityRuleContainer, propbase As GameEntityRuleContainer
    Dim ecmbase As GameEntityRuleContainer, repairbase As GameEntityRuleContainer, conbase As GameEntityRuleContainer
    Dim brainbase As GameEntityRuleContainer, wbase As GameEntityRuleContainer, structbase As GameEntityRuleContainer
    Dim empty_ As GameEntityRuleContainer, rules As GameEntityRules
    '== * ```Upgrades``` A special array containing per-player rules information for game entity types,
    '== which can be written to in order to implement upgrades and other dynamic rules changes. Each item in the
    '== array contains a subset of the sparse array of rules information in the ```Stats``` global.
    '== These values are defined:
    BuildNameToTypeMaps
    ReDim upgrades(0 To MAX_PLAYERS - 1)
    For i = 0 To MAX_PLAYERS - 1
        PerPlayerUpgrades_construct upgrades(i), i
        n = n + 1

        '==   * ```Body``` Droid bodies
        bodybase = empty_
        For j = 0 To asBodyStats_size - 1
            rules = GameEntityRules_construct(i, j, NTT_BODY)
            GameEntityRuleContainer_addRules bodybase, asBodyStats(j).name_, rules
        Next
        PerPlayerUpgrades_addGameEntity upgrades(i), "Body", bodybase

        '==   * ```Sensor``` Sensor turrets
        sensorbase = empty_
        For j = 0 To asSensorStats_size - 1
            rules = GameEntityRules_construct(i, j, NTT_SENSOR)
            GameEntityRuleContainer_addRules sensorbase, asSensorStats(j).name_, rules
        Next
        PerPlayerUpgrades_addGameEntity upgrades(i), "Sensor", sensorbase

        '==   * ```Propulsion``` Propulsions
        propbase = empty_
        For j = 0 To asPropulsionStats_size - 1
            rules = GameEntityRules_construct(i, j, NTT_PROPULSION)
            GameEntityRuleContainer_addRules propbase, asPropulsionStats(j).name_, rules
        Next
        PerPlayerUpgrades_addGameEntity upgrades(i), "Propulsion", propbase

        '==   * ```ECM``` ECM (Electronic Counter-Measure) turrets
        ecmbase = empty_
        For j = 0 To asECMStats_size - 1
            rules = GameEntityRules_construct(i, j, NTT_ECM)
            GameEntityRuleContainer_addRules ecmbase, asECMStats(j).name_, rules
        Next
        PerPlayerUpgrades_addGameEntity upgrades(i), "ECM", ecmbase

        '==   * ```Repair``` Repair turrets (not used, incidentally, for repair centers)
        repairbase = empty_
        For j = 0 To asRepairStats_size - 1
            rules = GameEntityRules_construct(i, j, NTT_REPAIR)
            GameEntityRuleContainer_addRules repairbase, asRepairStats(j).name_, rules
        Next
        PerPlayerUpgrades_addGameEntity upgrades(i), "Repair", repairbase

        '==   * ```Construct``` Constructor turrets (eg for trucks)
        conbase = empty_
        For j = 0 To asConstructStats_size - 1
            rules = GameEntityRules_construct(i, j, NTT_CONSTRUCT)
            GameEntityRuleContainer_addRules conbase, asConstructStats(j).name_, rules
        Next
        PerPlayerUpgrades_addGameEntity upgrades(i), "Construct", conbase

        '==   * ```Brain``` Brains
        '== BaseCommandLimit: How many droids a commander can command. CommandLimitByLevel: How many extra droids
        '== a commander can command for each of its rank levels. RankThresholds: An array describing how many
        '== kills are required for this brain to level up to the next rank. To alter this from scripts, you must
        '== set the entire array at once. Setting each item in the array will not work at the moment.
        brainbase = empty_
        For j = 0 To asBrainStats_size - 1
            rules = GameEntityRules_construct(i, j, NTT_BRAIN)
            GameEntityRuleContainer_addRules brainbase, asBrainStats(j).name_, rules
        Next
        PerPlayerUpgrades_addGameEntity upgrades(i), "Brain", brainbase

        '==   * ```Weapon``` Weapon turrets
        wbase = empty_
        For j = 0 To asWeaponStats_size - 1
            rules = GameEntityRules_construct(i, j, NTT_WEAPON)
            GameEntityRuleContainer_addRules wbase, asWeaponStats(j).name_, rules
        Next
        PerPlayerUpgrades_addGameEntity upgrades(i), "Weapon", wbase

        '==   * ```Building``` Buildings
        structbase = empty_
        For j = 0 To numStructureStats - 1
            rules = GameEntityRules_construct(i, j, NTT_BUILDING)
            GameEntityRuleContainer_addRules structbase, asStructureStats(j).name_, rules
        Next
        PerPlayerUpgrades_addGameEntity upgrades(i), "Building", structbase
    Next
    wzapi_getUpgradesObject = n
End Function

Private Function register_common(ByVal psStats As Long) As nlohmann_json
    Dim v As nlohmann_json
    Set v = json_object()
    v.set_item_ref "Id", json_value(BASE_STATS_id(psStats))
    v.set_item_ref "Weight", jsonUnsigned(COMPONENT_STATS_weight(psStats))
    v.set_item_ref "BuildPower", jsonUnsigned(COMPONENT_STATS_buildPower(psStats))
    v.set_item_ref "BuildTime", jsonUnsigned(COMPONENT_STATS_buildPoints(psStats))
    v.set_item_ref "HitPoints", jsonUnsigned(COMPONENT_STATS_getBase_hitpoints(psStats))
    v.set_item_ref "HitPointPct", jsonInt(COMPONENT_STATS_getBase_hitpointPct(psStats))
    Set register_common = v
End Function

' (VB) json(bool)
Private Function jsonBool(ByVal b As Boolean) As nlohmann_json
    Set jsonBool = json_value(b)
End Function

' (VB) json(const char *) and json(WzString)
Private Function jsonString(ByVal s As String) As nlohmann_json
    Set jsonString = json_value(s)
End Function

Public Function wzapi_constructStatsObject() As nlohmann_json
    Dim stats As nlohmann_json, bodybase As nlohmann_json, body As nlohmann_json, sensorbase As nlohmann_json
    Dim sensor As nlohmann_json, ecmbase As nlohmann_json, ecm As nlohmann_json, propbase As nlohmann_json
    Dim v As nlohmann_json, repairbase As nlohmann_json, repair As nlohmann_json, conbase As nlohmann_json
    Dim con As nlohmann_json, brainbase As nlohmann_json, br As nlohmann_json, thresholds As nlohmann_json
    Dim ranks As nlohmann_json, wbase As nlohmann_json, weap As nlohmann_json, weaponTypes As nlohmann_json
    Dim structbase As nlohmann_json, strct As nlohmann_json, weaps As nlohmann_json, researchbase As nlohmann_json
    Dim res As nlohmann_json, reqs As nlohmann_json, resstrcts As nlohmann_json, rescomps As nlohmann_json
    Dim j As Long, x As Long, k As Long, t As STRUCTURE_TYPE
    ' Register 'Stats' object. It is a read-only representation of basic game component states.
    '== * ```Stats``` A sparse, read-only array containing rules information for game entity types.
    '== (For now only the highest level member attributes are documented here. Use the 'jsdebug' cheat
    '== to see them all.)
    '== These values are defined:
    Set stats = json_object()
    '==   * ```Body``` Droid bodies
    Set bodybase = json_object()
    For j = 0 To asBodyStats_size - 1
        Set body = register_common(STAT_BODY + j)
        body.set_item_ref "Power", jsonUnsigned(asBodyStats(j).base.power)
        body.set_item_ref "Armour", jsonUnsigned(asBodyStats(j).base.armour)
        body.set_item_ref "Thermal", jsonInt(asBodyStats(j).base.thermal)
        body.set_item_ref "Resistance", jsonInt(asBodyStats(j).base.resistance)
        body.set_item_ref "Size", jsonUnsigned(asBodyStats(j).size)     ' (an enum with no negative value: GCC's underlying type is unsigned)
        body.set_item_ref "WeaponSlots", jsonUnsigned(asBodyStats(j).weaponSlots)
        body.set_item_ref "BodyClass", jsonString(asBodyStats(j).bodyClass)
        bodybase.set_item_ref asBodyStats(j).name_, body
    Next
    stats.set_item_ref "Body", bodybase

    '==   * ```Sensor``` Sensor turrets
    Set sensorbase = json_object()
    For j = 0 To asSensorStats_size - 1
        Set sensor = register_common(STAT_SENSOR + j)
        sensor.set_item_ref "Range", jsonUnsigned(asSensorStats(j).base.range)
        sensorbase.set_item_ref asSensorStats(j).name_, sensor
    Next
    stats.set_item_ref "Sensor", sensorbase

    '==   * ```ECM``` ECM (Electronic Counter-Measure) turrets
    Set ecmbase = json_object()
    For j = 0 To asECMStats_size - 1
        Set ecm = register_common(STAT_ECM + j)
        ecm.set_item_ref "Range", jsonUnsigned(asECMStats(j).base.range)
        ecmbase.set_item_ref asECMStats(j).name_, ecm
    Next
    stats.set_item_ref "ECM", ecmbase

    '==   * ```Propulsion``` Propulsions
    Set propbase = json_object()
    For j = 0 To asPropulsionStats_size - 1
        Set v = register_common(STAT_PROPULSION + j)
        v.set_item_ref "HitpointPctOfBody", jsonInt(asPropulsionStats(j).base.hitpointPctOfBody)
        v.set_item_ref "MaxSpeed", jsonUnsigned(asPropulsionStats(j).maxSpeed)
        v.set_item_ref "TurnSpeed", jsonUnsigned(asPropulsionStats(j).turnSpeed)
        v.set_item_ref "SpinSpeed", jsonUnsigned(asPropulsionStats(j).spinSpeed)
        v.set_item_ref "SpinAngle", jsonUnsigned(asPropulsionStats(j).spinAngle)
        v.set_item_ref "SkidDeceleration", jsonUnsigned(asPropulsionStats(j).skidDeceleration)
        v.set_item_ref "Acceleration", jsonUnsigned(asPropulsionStats(j).acceleration)
        v.set_item_ref "Deceleration", jsonUnsigned(asPropulsionStats(j).deceleration)
        propbase.set_item_ref asPropulsionStats(j).name_, v
    Next
    stats.set_item_ref "Propulsion", propbase

    '==   * ```Repair``` Repair turrets (not used, incidentally, for repair centers)
    Set repairbase = json_object()
    For j = 0 To asRepairStats_size - 1
        Set repair = register_common(STAT_REPAIR + j)
        repair.set_item_ref "RepairPoints", jsonUnsigned(asRepairStats(j).base.repairPoints)
        repairbase.set_item_ref asRepairStats(j).name_, repair
    Next
    stats.set_item_ref "Repair", repairbase

    '==   * ```Construct``` Constructor turrets (eg for trucks)
    Set conbase = json_object()
    For j = 0 To asConstructStats_size - 1
        Set con = register_common(STAT_CONSTRUCT + j)
        con.set_item_ref "ConstructorPoints", jsonUnsigned(asConstructStats(j).base.constructPoints)
        conbase.set_item_ref asConstructStats(j).name_, con
    Next
    stats.set_item_ref "Construct", conbase

    '==   * ```Brain``` Brains
    Set brainbase = json_object()
    For j = 0 To asBrainStats_size - 1
        Set br = register_common(STAT_BRAIN + j)
        br.set_item_ref "BaseCommandLimit", jsonInt(asBrainStats(j).base.maxDroids)
        br.set_item_ref "CommandLimitByLevel", jsonInt(asBrainStats(j).base.maxDroidsMult)
        Set thresholds = json_array() 'engine->newArray(psStats->base.rankThresholds.size());
        For x = 0 To asBrainStats(j).base.rankThresholds_size - 1
            thresholds.push_back_ref jsonInt(asBrainStats(j).base.rankThresholds(x))
        Next
        br.set_item "RankThresholds", thresholds
        Set ranks = json_array() 'engine->newArray(psStats->rankNames.size());
        For x = 0 To asBrainStats(j).rankNames.size - 1
            ranks.push_back_ref jsonString(asBrainStats(j).rankNames.items(x))
        Next
        br.set_item "RankNames", ranks
        brainbase.set_item_ref asBrainStats(j).name_, br
    Next
    stats.set_item_ref "Brain", brainbase

    '==   * ```Weapon``` Weapon turrets
    Set wbase = json_object()
    For j = 0 To asWeaponStats_size - 1
        Set weap = register_common(STAT_WEAPON + j)
        weap.set_item_ref "MaxRange", jsonUnsigned(asWeaponStats(j).base.maxRange)
        weap.set_item_ref "ShortRange", jsonUnsigned(asWeaponStats(j).base.shortRange)
        weap.set_item_ref "MinRange", jsonUnsigned(asWeaponStats(j).base.minRange)
        weap.set_item_ref "HitChance", jsonUnsigned(asWeaponStats(j).base.hitChance)
        weap.set_item_ref "ShortHitChance", jsonUnsigned(asWeaponStats(j).base.shortHitChance)
        weap.set_item_ref "FirePause", jsonUnsigned(asWeaponStats(j).base.firePause)
        weap.set_item_ref "ReloadTime", jsonUnsigned(asWeaponStats(j).base.reloadTime)
        weap.set_item_ref "Rounds", jsonUnsigned(asWeaponStats(j).base.numRounds)
        weap.set_item_ref "Damage", jsonUnsigned(asWeaponStats(j).base.damage)
        weap.set_item_ref "MinimumDamage", jsonUnsigned(asWeaponStats(j).base.minimumDamage)
        weap.set_item_ref "RadiusDamage", jsonUnsigned(asWeaponStats(j).base.radiusDamage)
        weap.set_item_ref "RepeatDamage", jsonUnsigned(asWeaponStats(j).base.periodicalDamage)
        weap.set_item_ref "RepeatRadius", jsonUnsigned(asWeaponStats(j).base.periodicalDamageRadius)
        weap.set_item_ref "RepeatTime", jsonUnsigned(asWeaponStats(j).base.periodicalDamageTime)
        weap.set_item_ref "Radius", jsonUnsigned(asWeaponStats(j).base.radius)
        weap.set_item_ref "EmpRadius", jsonUnsigned(asWeaponStats(j).base.empRadius)
        If asWeaponStats(j).weaponClass = WC_KINETIC Then
            weap.set_item_ref "ImpactType", jsonString("KINETIC")
        Else
            weap.set_item_ref "ImpactType", jsonString("HEAT")
        End If
        If asWeaponStats(j).periodicalDamageWeaponClass = WC_KINETIC Then
            weap.set_item_ref "RepeatType", jsonString("KINETIC")
        Else
            weap.set_item_ref "RepeatType", jsonString("HEAT")
        End If
        weap.set_item_ref "ImpactClass", jsonString(getWeaponSubClass_name(asWeaponStats(j).weaponSubClass))
        weap.set_item_ref "RepeatClass", jsonString(getWeaponSubClass_name(asWeaponStats(j).periodicalDamageWeaponSubClass))
        weap.set_item_ref "FireOnMove", jsonBool(asWeaponStats(j).fireOnMove)
        weap.set_item_ref "Effect", jsonString(getWeaponEffect_name(asWeaponStats(j).weaponEffect))
        weap.set_item_ref "ShootInAir", jsonBool((asWeaponStats(j).surfaceToAir And SHOOT_IN_AIR) <> 0)
        weap.set_item_ref "ShootOnGround", jsonBool((asWeaponStats(j).surfaceToAir And SHOOT_ON_GROUND) <> 0)
        weap.set_item_ref "NoFriendlyFire", jsonBool((asWeaponStats(j).flags And BIT_(WEAPON_FLAG_NO_FRIENDLY_FIRE)) <> 0)
        weap.set_item_ref "AllowedOnTransporter", jsonBool((asWeaponStats(j).flags And BIT_(WEAPON_FLAG_ALLOWED_ON_TRANSPORTER)) <> 0)
        weap.set_item_ref "TeleportCapture", jsonBool((asWeaponStats(j).flags And BIT_(WEAPON_FLAG_TELEPORT_CAPTURE)) <> 0)
        weap.set_item_ref "FlightSpeed", jsonUnsigned(asWeaponStats(j).flightSpeed)
        weap.set_item_ref "Rotate", jsonInt(asWeaponStats(j).rotate)
        weap.set_item_ref "MinElevation", jsonInt(asWeaponStats(j).minElevation)
        weap.set_item_ref "MaxElevation", jsonInt(asWeaponStats(j).maxElevation)
        weap.set_item_ref "Recoil", jsonUnsigned(asWeaponStats(j).recoilValue)
        weap.set_item_ref "Penetrate", jsonBool(asWeaponStats(j).penetrate)
        wbase.set_item_ref asWeaponStats(j).name_, weap
    Next
    stats.set_item_ref "Weapon", wbase

    '==   * ```WeaponClass``` Defined weapon classes
    Set weaponTypes = json_array() 'engine->newArray(WSC_NUM_WEAPON_SUBCLASSES);
    For j = 0 To WSC_NUM_WEAPON_SUBCLASSES - 1
        weaponTypes.push_back_ref jsonString(getWeaponSubClass_name(j))
    Next
    stats.set_item_ref "WeaponClass", weaponTypes

    '==   * ```Building``` Buildings
    Set structbase = json_object()
    For j = 0 To numStructureStats - 1
        Set strct = json_object()
        strct.set_item_ref "Id", jsonString(asStructureStats(j).id)
        t = asStructureStats(j).typ
        If t = REF_DEFENSE Or t = REF_WALL Or t = REF_WALLCORNER Or t = REF_GATE Or t = REF_FORTRESS Then
            strct.set_item_ref "Type", jsonString("Wall")
        ElseIf t = REF_GENERIC Then
            strct.set_item_ref "Type", jsonString("Generic")
        ElseIf t <> REF_DEMOLISH Then
            strct.set_item_ref "Type", jsonString("Structure")
        Else
            strct.set_item_ref "Type", jsonString("Demolish")
        End If
        strct.set_item_ref "ResearchPoints", jsonUnsigned(asStructureStats(j).base.research)
        strct.set_item_ref "RepairPoints", jsonUnsigned(asStructureStats(j).base.repair)
        strct.set_item_ref "PowerPoints", jsonUnsigned(asStructureStats(j).base.power)
        strct.set_item_ref "ProductionPoints", jsonUnsigned(asStructureStats(j).base.production)
        strct.set_item_ref "RearmPoints", jsonUnsigned(asStructureStats(j).base.rearm)
        strct.set_item_ref "Armour", jsonUnsigned(asStructureStats(j).base.armour)
        strct.set_item_ref "Thermal", jsonUnsigned(asStructureStats(j).base.thermal)
        strct.set_item_ref "HitPoints", jsonUnsigned(asStructureStats(j).base.hitpoints)
        strct.set_item_ref "Resistance", jsonUnsigned(asStructureStats(j).base.resistance)
        strct.set_item_ref "BuildPower", jsonUnsigned(asStructureStats(j).powerToBuild)
        Set weaps = json_array()
        For k = 0 To MAX_WEAPONS - 1
            If asStructureStats(j).psWeapStat(k) <> -1 Then
                weaps.push_back_ref jsonString(asWeaponStats(asStructureStats(j).psWeapStat(k)).id)
            End If
        Next
        strct.set_item "Weapons", weaps
        structbase.set_item_ref asStructureStats(j).name_, strct
    Next
    stats.set_item_ref "Building", structbase

    '==   * ```Research``` Researches
    Set researchbase = json_object()
    For j = 0 To asResearch_size - 1
        Set res = json_object()
        res.set_item_ref "Id", jsonString(asResearch(j).id)
        Set reqs = json_array()
        For k = 0 To asResearch(j).pPRList_size - 1
            reqs.push_back_ref jsonString(asResearch(asResearch(j).pPRList(k)).id)
        Next
        res.set_item "Requires", reqs
        Set resstrcts = json_array()
        For k = 0 To asResearch(j).pStructureResults_size - 1
            resstrcts.push_back_ref jsonString(asStructureStats(asResearch(j).pStructureResults(k)).id)
        Next
        res.set_item "ResultStructures", resstrcts
        Set rescomps = json_array()
        For k = 0 To asResearch(j).componentResults_size - 1
            rescomps.push_back_ref jsonString(BASE_STATS_id(asResearch(j).componentResults(k)))
        Next
        res.set_item "ResultComponents", rescomps
        researchbase.set_item_ref asResearch(j).name_, res
    Next
    stats.set_item_ref "Research", researchbase
    Set wzapi_constructStatsObject = stats
End Function

Public Function wzapi_getUsefulConstants() As nlohmann_json
    Dim constants As nlohmann_json
    Set constants = json_object()

    constants.set_item_ref "TER_WATER", jsonInt(TER_WATER)
    constants.set_item_ref "TER_CLIFFFACE", jsonInt(TER_CLIFFFACE)
    constants.set_item_ref "WEATHER_RAIN", jsonInt(WT_RAINING)
    constants.set_item_ref "WEATHER_SNOW", jsonInt(WT_SNOWING)
    constants.set_item_ref "WEATHER_CLEAR", jsonInt(WT_NONE)
    constants.set_item_ref "DORDER_STOP", jsonInt(DORDER_STOP)
    constants.set_item_ref "DORDER_MOVE", jsonInt(DORDER_MOVE)
    constants.set_item_ref "DORDER_ATTACK", jsonInt(DORDER_ATTACK)
    constants.set_item_ref "DORDER_BUILD", jsonInt(DORDER_BUILD)
    constants.set_item_ref "DORDER_HELPBUILD", jsonInt(DORDER_HELPBUILD)
    constants.set_item_ref "DORDER_LINEBUILD", jsonInt(DORDER_LINEBUILD)
    constants.set_item_ref "DORDER_DEMOLISH", jsonInt(DORDER_DEMOLISH)
    constants.set_item_ref "DORDER_REPAIR", jsonInt(DORDER_REPAIR)
    constants.set_item_ref "DORDER_OBSERVE", jsonInt(DORDER_OBSERVE)
    constants.set_item_ref "DORDER_FIRESUPPORT", jsonInt(DORDER_FIRESUPPORT)
    constants.set_item_ref "DORDER_RTB", jsonInt(DORDER_RTB)
    constants.set_item_ref "DORDER_RTR", jsonInt(DORDER_RTR)
    constants.set_item_ref "DORDER_EMBARK", jsonInt(DORDER_EMBARK)
    constants.set_item_ref "DORDER_DISEMBARK", jsonInt(DORDER_DISEMBARK)
    constants.set_item_ref "DORDER_COMMANDERSUPPORT", jsonInt(DORDER_COMMANDERSUPPORT)
    constants.set_item_ref "DORDER_RECYCLE", jsonInt(DORDER_RECYCLE)
    constants.set_item_ref "DORDER_SCOUT", jsonInt(DORDER_SCOUT)
    constants.set_item_ref "DORDER_PATROL", jsonInt(DORDER_PATROL)
    constants.set_item_ref "DORDER_REARM", jsonInt(DORDER_REARM)
    constants.set_item_ref "DORDER_RECOVER", jsonInt(DORDER_RECOVER)
    constants.set_item_ref "DORDER_HOLD", jsonInt(DORDER_HOLD)
    constants.set_item_ref "BUILD", jsonInt(IDRET_BUILD) ' deprecated
    constants.set_item_ref "MANUFACTURE", jsonInt(IDRET_MANUFACTURE) ' deprecated
    constants.set_item_ref "RESEARCH", jsonInt(IDRET_RESEARCH) ' deprecated
    constants.set_item_ref "INTELMAP", jsonInt(IDRET_INTEL_MAP) ' deprecated
    constants.set_item_ref "DESIGN", jsonInt(IDRET_DESIGN) ' deprecated
    constants.set_item_ref "CANCEL", jsonInt(IDRET_CANCEL) ' deprecated
    constants.set_item_ref "COMMAND", jsonInt(IDRET_COMMAND) ' deprecated
    constants.set_item_ref "CAMP_CLEAN", jsonInt(CAMP_CLEAN)
    constants.set_item_ref "CAMP_BASE", jsonInt(CAMP_BASE)
    constants.set_item_ref "CAMP_WALLS", jsonInt(CAMP_WALLS)
    constants.set_item_ref "NO_ALLIANCES", jsonInt(NO_ALLIANCES)
    constants.set_item_ref "ALLIANCES", jsonInt(ALLIANCES_)
    constants.set_item_ref "ALLIANCES_TEAMS", jsonInt(ALLIANCES_TEAMS)
    constants.set_item_ref "ALLIANCES_UNSHARED", jsonInt(ALLIANCES_UNSHARED)
    constants.set_item_ref "NO_SCAVENGERS", jsonInt(NO_SCAVENGERS)
    constants.set_item_ref "SCAVENGERS", jsonInt(SCAVENGERS)
    constants.set_item_ref "ULTIMATE_SCAVENGERS", jsonInt(ULTIMATE_SCAVENGERS)
    constants.set_item_ref "BEING_BUILT", jsonInt(SS_BEING_BUILT)
    constants.set_item_ref "BUILT", jsonInt(SS_BUILT)
    constants.set_item_ref "DROID_WEAPON", jsonInt(DROID_WEAPON)
    constants.set_item_ref "DROID_SENSOR", jsonInt(DROID_SENSOR)
    constants.set_item_ref "DROID_ECM", jsonInt(DROID_ECM)
    constants.set_item_ref "DROID_CONSTRUCT", jsonInt(DROID_CONSTRUCT)
    constants.set_item_ref "DROID_PERSON", jsonInt(DROID_PERSON)
    constants.set_item_ref "DROID_CYBORG", jsonInt(DROID_CYBORG)
    constants.set_item_ref "DROID_TRANSPORTER", jsonInt(DROID_TRANSPORTER)
    constants.set_item_ref "DROID_COMMAND", jsonInt(DROID_COMMAND)
    constants.set_item_ref "DROID_REPAIR", jsonInt(DROID_REPAIR)
    constants.set_item_ref "DROID_SUPERTRANSPORTER", jsonInt(DROID_SUPERTRANSPORTER)
    constants.set_item_ref "DROID_ANY", jsonInt(DROID_ANY)
    constants.set_item_ref "ARTIFACT", jsonInt(FEAT_GEN_ARTE)
    constants.set_item_ref "OIL_RESOURCE", jsonInt(FEAT_OIL_RESOURCE)
    constants.set_item_ref "BUILDING", jsonInt(FEAT_BUILDING)
    constants.set_item_ref "OIL_DRUM", jsonInt(FEAT_OIL_DRUM)
    constants.set_item_ref "HQ", jsonInt(REF_HQ)
    constants.set_item_ref "FACTORY", jsonInt(REF_FACTORY)
    constants.set_item_ref "POWER_GEN", jsonInt(REF_POWER_GEN)
    constants.set_item_ref "RESOURCE_EXTRACTOR", jsonInt(REF_RESOURCE_EXTRACTOR)
    constants.set_item_ref "DEFENSE", jsonInt(REF_DEFENSE)
    constants.set_item_ref "WALL", jsonInt(REF_WALL)
    constants.set_item_ref "RESEARCH_LAB", jsonInt(REF_RESEARCH)
    constants.set_item_ref "REPAIR_FACILITY", jsonInt(REF_REPAIR_FACILITY)
    constants.set_item_ref "COMMAND_CONTROL", jsonInt(REF_COMMAND_CONTROL)
    constants.set_item_ref "CYBORG_FACTORY", jsonInt(REF_CYBORG_FACTORY)
    constants.set_item_ref "VTOL_FACTORY", jsonInt(REF_VTOL_FACTORY)
    constants.set_item_ref "REARM_PAD", jsonInt(REF_REARM_PAD)
    constants.set_item_ref "SAT_UPLINK", jsonInt(REF_SAT_UPLINK)
    constants.set_item_ref "GATE", jsonInt(REF_GATE)
    constants.set_item_ref "LASSAT", jsonInt(REF_LASSAT)
    constants.set_item_ref "STRUCT_GENERIC", jsonInt(REF_GENERIC)
    constants.set_item_ref "SUPEREASY", jsonInt(AIDifficulty_SUPEREASY)   ' static_cast<int8_t>(AIDifficulty::SUPEREASY)
    constants.set_item_ref "EASY", jsonInt(AIDifficulty_EASY)
    constants.set_item_ref "MEDIUM", jsonInt(AIDifficulty_MEDIUM)
    constants.set_item_ref "HARD", jsonInt(AIDifficulty_HARD)
    constants.set_item_ref "INSANE", jsonInt(AIDifficulty_INSANE)
    constants.set_item_ref "ALL_PLAYERS", jsonInt(ALL_PLAYERS)
    constants.set_item_ref "ALLIES", jsonInt(ALLIES)
    constants.set_item_ref "ENEMIES", jsonInt(ENEMIES)
    constants.set_item_ref "DROID", jsonInt(OBJ_DROID)
    constants.set_item_ref "STRUCTURE", jsonInt(OBJ_STRUCTURE)
    constants.set_item_ref "FEATURE", jsonInt(OBJ_FEATURE)
    constants.set_item_ref "POSITION", jsonInt(SCRIPT_POSITION)
    constants.set_item_ref "AREA", jsonInt(SCRIPT_AREA)
    constants.set_item_ref "PLAYER_DATA", jsonInt(SCRIPT_PLAYER)
    constants.set_item_ref "RESEARCH_DATA", jsonInt(SCRIPT_RESEARCH)
    constants.set_item_ref "GROUP", jsonInt(SCRIPT_GROUP)
    constants.set_item_ref "RADIUS", jsonInt(SCRIPT_RADIUS)
    constants.set_item_ref "LZ_COMPROMISED_TIME", jsonInt(JS_LZ_COMPROMISED_TIME)
    constants.set_item_ref "OBJECT_FLAG_UNSELECTABLE", jsonInt(OBJECT_FLAG_UNSELECTABLE)
    ' ShowTopicFlags constants
    constants.set_item_ref "SHOWTOPIC_FIRSTADD", jsonInt(ShowTopicFlags_FirstAdd)
    constants.set_item_ref "SHOWTOPIC_NEVERVIEWED", jsonInt(ShowTopicFlags_NeverViewed)
    ' the constants below are subject to change without notice...
    constants.set_item_ref "RES_MSG", jsonInt(MSG_RESEARCH)
    constants.set_item_ref "CAMP_MSG", jsonInt(MSG_CAMPAIGN)
    constants.set_item_ref "MISS_MSG", jsonInt(MSG_MISSION)
    constants.set_item_ref "PROX_MSG", jsonInt(MSG_PROXIMITY)
    constants.set_item_ref "LDS_EXPAND_LIMBO", jsonInt(LEVEL_TYPE_LDS_EXPAND_LIMBO)  ' static_cast<int8_t>(LEVEL_TYPE::LDS_EXPAND_LIMBO)

    constants.set_item_ref "WzQuickChatMessages", getWzQuickChatMessageValueMap()

    Set wzapi_getUsefulConstants = constants
End Function

Public Function wzapi_constructStaticPlayerData() As nlohmann_json
    Dim playerData As nlohmann_json, vector As nlohmann_json, i As Long, aiPlayer As Boolean, js As String
    Dim aidata_size As Long
    ' Static knowledge about players
    '== * ```playerData``` An array of information about the players in a game. Each item in the array is an object
    '== containing the following variables:
    '==   * ```difficulty``` (see ```difficulty``` global constant)
    '==   * ```faction``` chosen faction in MP. 0 - normal, 1 - NEXUS, 2 - Collective (4.0.0+ only)
    '==   * ```colour``` number describing the colour of the player
    '==   * ```position``` number describing the position of the player in the game's setup screen
    '==   * ```isAI``` whether the player is an AI (3.2+ only)
    '==   * ```isHuman``` whether the player is human (3.2+ only)
    '==   * ```name``` the name of the player (3.2+ only)
    '==   * ```team``` the number of the team the player is part of
    '==   * ```scriptName``` File name of the AI's JS file. (4.6.2+ only)
    Set playerData = json_array()  ' engine->newArray(game.maxPlayers);
    aidata_size = getAIData_size()  ' const auto& aidata = getAIData();
    For i = 0 To game.maxPlayers - 1
        Set vector = json_object()
        aiPlayer = (Not NetPlay_players_allocated(i)) And NetPlay_players_ai(i) >= 0
        If game.blindMode >= BLIND_MODE_BLIND_GAME Then
            ' to ensure the "name" field exposed to api is consistent across blind games _and_ replays of those games, always set it to the generic name if in blind mode
            vector.set_item_ref "name", jsonString(getPlayerGenericName(i))
        Else
            vector.set_item_ref "name", jsonString(NetPlay_players_name(i))
        End If
        vector.set_item_ref "difficulty", jsonInt(NetPlay_players_difficulty(i))  ' static_cast<int8_t>
        vector.set_item_ref "faction", jsonUnsigned(NetPlay_players_faction(i))   ' FactionID : uint8_t
        vector.set_item_ref "colour", jsonInt(NetPlay_players_colour(i))
        vector.set_item_ref "position", jsonInt(NetPlay_players_position(i))
        vector.set_item_ref "team", jsonInt(NetPlay_players_team(i))
        vector.set_item_ref "isAI", jsonBool(aiPlayer)
        vector.set_item_ref "isHuman", jsonBool(NetPlay_players_allocated(i))
        vector.set_item_ref "type", jsonInt(SCRIPT_PLAYER)
        If aiPlayer Then
            If NetPlay_players_ai(i) < aidata_size Then js = getAIData_js(NetPlay_players_ai(i)) Else js = "AI"
            vector.set_item_ref "scriptName", jsonString(js)
        End If
        playerData.push_back_ref vector
    Next
    Set wzapi_constructStaticPlayerData = playerData
End Function

Public Function wzapi_constructMapTilesArray() As nlohmann_json
    Dim mapTileArray As nlohmann_json, mapRow As nlohmann_json, tileObj As nlohmann_json, x As Long, y As Long, psTile As Long
    ' Static knowledge about map tiles
    '== * ```MapTiles``` A two-dimensional array of static information about the map tiles in a game. Each item in MapTiles[y][x] is an object
    '== containing the following variables:
    '==   * ```terrainType``` (see ```terrainType(x, y)``` function)
    '==   * ```height``` the height at the top left of the tile
    '==   * ```hoverContinent``` (For hover type propulsions)
    '==   * ```limitedContinent``` (For land or sea limited propulsion types)
    Set mapTileArray = json_array()
    For y = 0 To mapHeight - 1
        Set mapRow = json_array()
        For x = 0 To mapWidth - 1
            psTile = mapTile(x, y)
            Set tileObj = json_object()     ' nlohmann::json mapTile
            tileObj.set_item_ref "terrainType", jsonUnsigned(terrainType(psTile))     ' unsigned char
            tileObj.set_item_ref "height", jsonInt(psMapTiles(psTile).height)
            tileObj.set_item_ref "hoverContinent", jsonUnsigned(psMapTiles(psTile).hoverContinent)     ' uint16_t
            tileObj.set_item_ref "limitedContinent", jsonUnsigned(psMapTiles(psTile).limitedContinent) ' uint16_t
            mapRow.push_back_ref tileObj
        Next
        mapTileArray.push_back_ref mapRow
    Next
    Set wzapi_constructMapTilesArray = mapTileArray
End Function

' (VB) scriptQueuedObjectRemovals().emplace_back(psObj, sfx); the first use reserves initialCapacity (32)
Private Sub queuedObjRemovals_emplace_back(ByVal psObj As Long, ByVal sfx As Boolean)
    Const initialCapacity As Long = 32
    If queuedObjRemovals_capacity = 0 Then
        queuedObjRemovals_capacity = initialCapacity
        ReDim queuedObjRemovals_obj(0 To queuedObjRemovals_capacity - 1)
        ReDim queuedObjRemovals_sfx(0 To queuedObjRemovals_capacity - 1)
    ElseIf queuedObjRemovals_size = queuedObjRemovals_capacity Then
        queuedObjRemovals_capacity = queuedObjRemovals_capacity * 2
        ReDim Preserve queuedObjRemovals_obj(0 To queuedObjRemovals_capacity - 1)
        ReDim Preserve queuedObjRemovals_sfx(0 To queuedObjRemovals_capacity - 1)
    End If
    queuedObjRemovals_obj(queuedObjRemovals_size) = psObj
    queuedObjRemovals_sfx(queuedObjRemovals_size) = sfx
    queuedObjRemovals_size = queuedObjRemovals_size + 1
End Sub

' (VB) wzapi::scriptQueuedObjectRemovals().empty() (executeFnAndProcessScriptQueuedRemovals' ASSERT) and size()
Public Function wzapi_scriptQueuedObjectRemovals_empty() As Boolean
    wzapi_scriptQueuedObjectRemovals_empty = (queuedObjRemovals_size = 0)
End Function

Public Function wzapi_scriptQueuedObjectRemovals_size() As Long
    wzapi_scriptQueuedObjectRemovals_size = queuedObjRemovals_size
End Function

Public Function wzapi_scriptIsObjectQueuedForRemoval(ByVal psObj As Long) As Boolean
    Dim i As Long
    ' std::any_of over the queue
    For i = 0 To queuedObjRemovals_size - 1
        If psObj = queuedObjRemovals_obj(i) Then
            wzapi_scriptIsObjectQueuedForRemoval = True
            Exit Function
        End If
    Next
    wzapi_scriptIsObjectQueuedForRemoval = False
End Function

' (the range-for's end is taken once: removals queued while it runs are cleared with the rest, unprocessed)
Public Sub wzapi_processScriptQueuedObjectRemovals()
    Dim i As Long, n As Long, psObj As Long
    n = queuedObjRemovals_size
    For i = 0 To n - 1
        psObj = queuedObjRemovals_obj(i)
        If SIMPLE_OBJECT_died(psObj) <> 0 Then
            debug_ 4831, LOG_MSG, "processScriptQueuedObjectRemovals", "Object %p is already dead, not processing", psObj
            GoTo continue_
        End If
        If queuedObjRemovals_sfx(i) Then
            Select Case SIMPLE_OBJECT_type(psObj)
            Case OBJ_STRUCTURE: destroyStruct psObj, gameTime
            Case OBJ_DROID: destroyDroid psObj, gameTime
            Case OBJ_FEATURE: destroyFeature psObj, gameTime
            Case Else: ASSERT_FAILURE "false", "src/wzapi.cpp:4841", "processScriptQueuedObjectRemovals", "Wrong game object type"
            End Select
        Else
            Select Case SIMPLE_OBJECT_type(psObj)
            Case OBJ_STRUCTURE: removeStruct psObj, True
            Case OBJ_DROID: removeDroidBase psObj
            Case OBJ_FEATURE: removeFeature psObj
            Case Else: ASSERT_FAILURE "false", "src/wzapi.cpp:4851", "processScriptQueuedObjectRemovals", "Wrong game object type"
            End Select
        End If
continue_:
    Next
    queuedObjRemovals_size = 0     ' queuedObjRemovals.clear()
End Sub
