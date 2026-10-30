Attribute VB_Name = "game_objects_c"
' src/game.cpp, its object files (lines 5243-8208 and the ini helpers at 2424-2497): the droids, the
' structures, the features, the templates, the terrain types, the component and structure type lists, the
' research, the messages, the structure limits, the fire support designators, the script state and the
' guide topics of a save game, read and written in the C++'s formats (WzConfig / JSON, and the old binary
' structure and feature files). game_c has the rest of game.cpp (loadGame, saveGame, the main file; the
' map's loadWzMap* functions); it calls these by their C++ names. A PerPlayerDroidLists& is the list
' array passed ByRef (its first element's address tells which one: mission.apsDroidLists'); a BASE_OBJECT *
' is a handle, a STRUCTURE_STATS * an asStructureStats index (-1 = nullptr), a DROID_TEMPLATE * a
' template_c pool handle (0 = nullptr: a local DROID_TEMPLATE is a pool slot of its own, given back after).
' The JSON values keep the C++ types' signedness (json_value / json_value_unsigned), as their dumps would
' differ past 2^31 otherwise.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' (VB) the binary layouts (game.cpp: STRUCT_SAVEHEADER and SAVE_STRUCTURE_V2 / OBJECT_SAVE_V19, natural alignment):
' the header is aFileType[4], version, quantity; a structure is name[40] @0, id @40, x @44, y @48, z @52,
' direction @56, player @60, inFire @64, periodicalDamageStart @68, periodicalDamage @72, status @76 (UBYTE),
' currentBuildPts @80, body @84, armour @88, resistance @92, dummy1 @96, subjectInc @100, timeStarted @104,
' output @108, capacity @112, quantity @116: 120 bytes
Private Const STRUCT_HEADER_SIZE As Long = 12
Private Const FEATURE_HEADER_SIZE As Long = 12
Private Const SIZEOF_SAVE_STRUCTURE_V2 As Long = 120
Private Const SIZEOF_SAVE_FEATURE_V2 As Long = 76     ' FEATURE_SAVE_V2 is OBJECT_SAVE_V19
Private Const SIZEOF_SAVE_FEATURE_V14 As Long = ((76 + MAX_PLAYERS + 3) \ 4) * 4    ' and UBYTE visible[MAX_PLAYERS] @76, to 4 bytes
Private Const MAX_SAVE_NAME_SIZE_V19 As Long = 40
Private Const VERSION_7 As Long = 7
Private Const VERSION_8 As Long = 8
Private Const VERSION_14 As Long = 14
Private Const VERSION_19 As Long = 19
Private Const NULL_ID As Long = UDWORD_MAX  ' objmem.h: #define NULL_ID UDWORD_MAX

' (VB) aigroupToDroidGroupMapping's entries (std::unordered_map<int, DROID_GROUP *>)
Private Type GROUP_MAPPING
    id As Long
    psGroup As Long
End Type

' ---- (VB) helpers: json values of the C++ types, the BASE_OBJECT fields basedef_h has no accessor for

Private Function jI(ByVal v As Long) As nlohmann_json
    Set jI = json_value(v)
End Function

Private Function jU(ByVal v As Long) As nlohmann_json
    Set jU = json_value_unsigned(v)
End Function

Private Function jB(ByVal v As Boolean) As nlohmann_json
    Set jB = json_value(v)
End Function

Private Function jS(ByVal v As String) As nlohmann_json
    Set jS = json_value(v)
End Function

' to_json(ivec2): [x, y]
Private Function jV2(ByVal x As Long, ByVal y As Long) As nlohmann_json
    Dim a As nlohmann_json
    Set a = json_array()
    a.push_back json_value(x)
    a.push_back json_value(y)
    Set jV2 = a
End Function

' to_json(ivec3): [x, y, z]
Private Function jV3(ByVal x As Long, ByVal y As Long, ByVal z As Long) As nlohmann_json
    Dim a As nlohmann_json
    Set a = json_array()
    a.push_back json_value(x)
    a.push_back json_value(y)
    a.push_back json_value(z)
    Set jV3 = a
End Function

' json_variant(int)
Private Function jvi(ByVal v As Long) As json_variant
    Set jvi = json_variant_new(v)
End Function

' C: (a <= b), (a > b), (a >= b) for uint32_t
Private Function ULE(ByVal a As Long, ByVal b As Long) As Boolean
    ULE = Not ULT(b, a)
End Function

Private Function UGT(ByVal a As Long, ByVal b As Long) As Boolean
    UGT = ULT(b, a)
End Function

Private Function UGE(ByVal a As Long, ByVal b As Long) As Boolean
    UGE = Not ULT(a, b)
End Function

' WzString::number(i).leftPadToMinimumLength('0', 10)
Private Function zeroPad10(ByVal n As Long) As String
    Dim s As String
    s = CStr(n)
    If Len(s) < 10 Then s = String$(10 - Len(s), "0") & s
    zeroPad10 = s
End Function

Private Sub objSetBorn(ByVal psObj As Long, ByVal v As Long)
    Select Case objHandleType(psObj)
    Case OBJ_DROID: aDroid(psObj).born = v
    Case OBJ_STRUCTURE: aStruct(psObj).born = v
    Case OBJ_FEATURE: aFeature(psObj).born = v
    End Select
End Sub

Private Function objSelected(ByVal psObj As Long) As Long
    Select Case objHandleType(psObj)
    Case OBJ_DROID: objSelected = aDroid(psObj).selected
    Case OBJ_STRUCTURE: objSelected = aStruct(psObj).selected
    Case OBJ_FEATURE: objSelected = aFeature(psObj).selected
    End Select
End Function

Private Sub objSetSelected(ByVal psObj As Long, ByVal v As Long)
    Select Case objHandleType(psObj)
    Case OBJ_DROID: aDroid(psObj).selected = v
    Case OBJ_STRUCTURE: aStruct(psObj).selected = v
    Case OBJ_FEATURE: aFeature(psObj).selected = v
    End Select
End Sub

Private Function objLastEmission(ByVal psObj As Long) As Long
    Select Case objHandleType(psObj)
    Case OBJ_DROID: objLastEmission = aDroid(psObj).lastEmission
    Case OBJ_STRUCTURE: objLastEmission = aStruct(psObj).lastEmission
    Case OBJ_FEATURE: objLastEmission = aFeature(psObj).lastEmission
    End Select
End Function

Private Sub objSetLastEmission(ByVal psObj As Long, ByVal v As Long)
    Select Case objHandleType(psObj)
    Case OBJ_DROID: aDroid(psObj).lastEmission = v
    Case OBJ_STRUCTURE: aStruct(psObj).lastEmission = v
    Case OBJ_FEATURE: aFeature(psObj).lastEmission = v
    End Select
End Sub

Private Function objPeriodicalDamageStart(ByVal psObj As Long) As Long
    Select Case objHandleType(psObj)
    Case OBJ_DROID: objPeriodicalDamageStart = aDroid(psObj).periodicalDamageStart
    Case OBJ_STRUCTURE: objPeriodicalDamageStart = aStruct(psObj).periodicalDamageStart
    Case OBJ_FEATURE: objPeriodicalDamageStart = aFeature(psObj).periodicalDamageStart
    End Select
End Function

Private Sub objSetPeriodicalDamageStart(ByVal psObj As Long, ByVal v As Long)
    Select Case objHandleType(psObj)
    Case OBJ_DROID: aDroid(psObj).periodicalDamageStart = v
    Case OBJ_STRUCTURE: aStruct(psObj).periodicalDamageStart = v
    Case OBJ_FEATURE: aFeature(psObj).periodicalDamageStart = v
    End Select
End Sub

' (VB) order.cpp's orderListResize (Private there): psDroid->asOrderList.resize(n)
Private Sub droidOrderListResize(ByVal psDroid As Long, ByVal n As Long)
    Dim i As Long, oldSize As Long
    oldSize = aDroid(psDroid).asOrderList_size
    If n > oldSize Then
        If oldSize = 0 Then
            ReDim aDroid(psDroid).asOrderList(0 To n - 1)
        ElseIf UBound(aDroid(psDroid).asOrderList) < n - 1 Then
            ReDim Preserve aDroid(psDroid).asOrderList(0 To n - 1)
        End If
        For i = oldSize To n - 1
            DroidOrder_construct aDroid(psDroid).asOrderList(i)
        Next
    End If
    aDroid(psDroid).asOrderList_size = n
End Sub

' ---- game.cpp:2424-2497: the ini helpers

Private Sub getIniBaseObject(ByVal ini As WzConfig, ByVal key As String, object As Long)
    Dim tid As Long, tplayer As Long, ttype As OBJECT_TYPE
    object = 0
    If ini.contains(key & "/id") Then
        tid = ini.value(key & "/id", jvi(-1)).toInt()
        tplayer = ini.value(key & "/player", jvi(-1)).toInt()
        ttype = ini.value(key & "/type", jvi(0)).toInt()
        If Not (tid >= 0 And tplayer >= 0) Then ASSERT_FAILURE "tid >= 0 && tplayer >= 0", "src/game.cpp:2432", "getIniBaseObject", "Bad ID": Exit Sub
        object = getBaseObjFromData(tid, tplayer, ttype)
        If Not (object <> 0) Then ASSERT_FAILURE "object != nullptr", "src/game.cpp:2434", "getIniBaseObject", "Failed to find target"
    End If
End Sub

Private Sub getIniStructureStats(ByVal ini As WzConfig, ByVal key As String, stats As Long)
    Dim statName As String, tid As Long
    stats = -1
    If ini.contains(key) Then
        statName = ini.value(key).toWzString()
        tid = getStructStatFromName(statName)
        If Not (tid >= 0) Then ASSERT_FAILURE "tid >= 0", "src/game.cpp:2445", "getIniStructureStats", "Target stats not found %s", statName: Exit Sub
        stats = tid
    End If
End Sub

Private Sub getIniDroidOrder(ByVal ini As WzConfig, ByVal key As String, order As DROID_ORDER_DATA)
    order.typ = ini.value(key & "/type", jvi(DORDER_NONE)).toInt()
    ini.vector2i key & "/pos", order.pos.x, order.pos.y
    ini.vector2i key & "/pos2", order.pos2.x, order.pos2.y
    order.direction = ini.value(key & "/direction").toInt() And &HFFFF&
    order.index = ini.value(key & "/index").toUInt()
    order.rtrType = ini.value(key & "/rtrType", jvi(RTR_TYPE_NO_RESULT)).toInt()
    getIniBaseObject ini, key & "/obj", order.psObj
    getIniStructureStats ini, key & "/stats", order.psStats
End Sub

Private Sub setIniBaseObject(ByVal json As nlohmann_json, ByVal key As String, ByVal object As Long)
    If object <> 0 Then
        If ULE(SIMPLE_OBJECT_died(object), NOT_CURRENT_LIST) Then
            json.set_item key & "/id", jU(SIMPLE_OBJECT_id(object))
            json.set_item key & "/player", jU(SIMPLE_OBJECT_player(object))
            json.set_item key & "/type", jI(SIMPLE_OBJECT_type(object))
            ' (#ifdef DEBUG: the debugfunc and debugline, commented out)
        End If
    End If
End Sub

Private Sub setIniStructureStats(ByVal jsonObj As nlohmann_json, ByVal key As String, ByVal stats As Long)
    If stats <> -1 Then
        jsonObj.set_item key, jS(asStructureStats(stats).id)
    End If
End Sub

Private Sub setIniDroidOrder(ByVal jsonObj As nlohmann_json, ByVal key As String, order As DROID_ORDER_DATA)
    jsonObj.set_item key & "/type", jI(order.typ)
    jsonObj.set_item key & "/pos", jV2(order.pos.x, order.pos.y)
    jsonObj.set_item key & "/pos2", jV2(order.pos2.x, order.pos2.y)
    jsonObj.set_item key & "/direction", jU(order.direction)
    jsonObj.set_item key & "/index", jU(order.index)
    jsonObj.set_item key & "/rtrType", jI(order.rtrType)
    setIniBaseObject jsonObj, key & "/obj", order.psObj
    setIniStructureStats jsonObj, key & "/stats", order.psStats
End Sub

' ---- game.cpp:5332: the players

Private Function RemapPlayerNumber(ByVal OldNumber As Long) As Long
    Dim i As Long

    If game.typ = LEVEL_TYPE_CAMPAIGN Then      ' don't remap for SP games
        RemapPlayerNumber = OldNumber
        Exit Function
    End If

    For i = 0 To MAX_PLAYERS - 1
        If OldNumber = NetPlay_players_position(i) Then
            game.mapHasScavengers = game.mapHasScavengers Or i = scavengerSlot()
            RemapPlayerNumber = i
            Exit Function
        End If
    Next
    ASSERT_FAILURE "false", "src/game.cpp:5349", "RemapPlayerNumber", "Found no player position for player %d", OldNumber
    RemapPlayerNumber = 0
End Function

Private Function getPlayer(ByVal ini As WzConfig) As Long
    Dim result As json_variant, position As Long, i As Long
    If ini.contains("player") Then
        Set result = ini.value("player")
        If Left$(result.toWzString(), 9) = "scavenger" Then
            game.mapHasScavengers = True
            getPlayer = scavengerSlot()
            Exit Function
        End If
        getPlayer = result.toInt()
        Exit Function
    ElseIf ini.contains("startpos") Then
        position = ini.value("startpos").toInt()
        For i = 0 To game.maxPlayers - 1
            If NetPlay_players_position(i) = position Then
                getPlayer = i
                Exit Function
            End If
        Next
    End If
    ASSERT_FAILURE "false", "src/game.cpp:5376", "getPlayer", "No player info found!"
    getPlayer = 0
End Function

Private Sub setPlayer(ByVal ini As WzConfig, ByVal player As Long)
    If scavengerSlot() = player Then
        ini.setValue "player", jS("scavenger")
    Else
        ini.setValue "player", jI(player)
    End If
End Sub

Private Sub setPlayerJSON(ByVal jsonObj As nlohmann_json, ByVal player As Long)
    If scavengerSlot() = player Then
        jsonObj.set_item "player", jS("scavenger")
    Else
        jsonObj.set_item "player", jI(player)
    End If
End Sub

' ---- the droids

Public Function loadSaveDroidPointers(ByVal pFileName As String, ppsCurrentDroidLists() As OBJ_LIST) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, psDroid As Long, id As Long, player As Long
    Dim it As Long, d As Long, it2 As Long, psTrDroid As Long, droidIdx As Long, j As Long
    Dim tid As Long, tplayer As Long, ttype As OBJECT_TYPE, psObj As Long, psCommander As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)
    n = ini.childGroups(list)

    For i = 0 To n - 1
        ini.beginGroup list(i)
        psDroid = 0
        id = ini.value("id", jvi(-1)).toInt()
        player = getPlayer(ini)

        If id <= 0 Then
            ini.endGroup
            GoTo continue_      ' special hack for campaign missions, cannot have targets
        End If

        it = ppsCurrentDroidLists(player).head
        Do While it <> 0
            d = lnObj(it)
            If aDroid(d).id = id Then
                psDroid = d
                Exit Do
            End If
            If DROID_isTransporter(d) And aDroid(d).psGroup <> 0 Then   ' Check for droids in the transporter.
                it2 = aGroup(aDroid(d).psGroup).psList.head
                Do While it2 <> 0
                    psTrDroid = lnObj(it2)
                    If aDroid(psTrDroid).id = id Then
                        psDroid = psTrDroid
                        GoTo foundDroid
                    End If
                    it2 = lnNext(it2)
                Loop
            End If
            it = lnNext(it)
        Loop
foundDroid:
        If psDroid = 0 Then
            d = getBaseObjFromId_list(mission_apsDroidLists(player), id)
            ' FIXME
            If d <> 0 Then
                debug_ 5448, LOG_ERROR, "loadSaveDroidPointers", "Droid %s (%d) was in wrong file/list (was in %s)...", objInfo(d), id, pFileName
            End If
        End If
        If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid", "src/game.cpp:5451", "loadSaveDroidPointers", "Droid %d not found", id: loadSaveDroidPointers = False: Exit Function
        aDroid(psDroid).listSize = clip_(ini.value("orderList/size", jvi(0)).toInt(), 0, 10000)
        droidOrderListResize psDroid, aDroid(psDroid).listSize    ' Must resize before setting any orders, and must set in-place, since pointers are updated later.
        For droidIdx = 0 To aDroid(psDroid).listSize - 1
            getIniDroidOrder ini, "orderList/" & CStr(droidIdx), aDroid(psDroid).asOrderList(droidIdx)
        Next
        aDroid(psDroid).listPendingBegin = 0
        For j = 0 To MAX_WEAPONS - 1
            objTrace aDroid(psDroid).id, "loadSaveDroidPointers", "weapon %d, nStat %d", j, aDroid(psDroid).asWeaps(j).nStat
            getIniBaseObject ini, "actionTarget/" & CStr(j), aDroid(psDroid).psActionTarget(j)
        Next
        If ini.contains("baseStruct/id") Then
            tid = ini.value("baseStruct/id", jvi(-1)).toInt()
            tplayer = ini.value("baseStruct/player", jvi(-1)).toInt()
            ttype = ini.value("baseStruct/type", jvi(0)).toInt()
            If Not (tid >= 0 And tplayer >= 0) Then ASSERT_FAILURE "tid >= 0 && tplayer >= 0", "src/game.cpp:5469", "loadSaveDroidPointers", "Bad ID"
            psObj = getBaseObjFromData(tid, tplayer, ttype)
            If Not (psObj <> 0) Then ASSERT_FAILURE "psObj", "src/game.cpp:5471", "loadSaveDroidPointers", "Failed to find droid base structure"
            If psObj <> 0 Then
                If Not (SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE) Then ASSERT_FAILURE "!psObj || psObj->type == OBJ_STRUCTURE", "src/game.cpp:5472", "loadSaveDroidPointers", "Droid base structure not a structure"
            End If
            aDroid(psDroid).psBaseStruct = psObj    ' setSaveDroidBase(psDroid, (STRUCTURE *)psObj) (#ifdef DEBUG: baseLine, baseFunc)
        End If
        If ini.contains("commander") Then
            tid = ini.value("commander", jvi(-1)).toInt()
            psCommander = getBaseObjFromData(tid, aDroid(psDroid).player, OBJ_DROID)
            If Not (psCommander <> 0) Then ASSERT_FAILURE "psCommander", "src/game.cpp:5479", "loadSaveDroidPointers", "Failed to find droid commander"
            cmdDroidAddDroid psCommander, psDroid
        End If
        getIniDroidOrder ini, "order", aDroid(psDroid).order
        ini.endGroup
continue_:
    Next
    loadSaveDroidPointers = True
End Function

Private Function healthValue(ByVal ini As WzConfig, ByVal defaultValue As Long) As Long
    Dim health As String, perc As Long
    health = ini.value("health").toWzString()
    If Len(health) = 0 Or defaultValue = 0 Then
        healthValue = defaultValue
    ElseIf InStr(1, health, "%", vbBinaryCompare) <> 0 Then
        perc = atoi(Replace(health, "%", ""))
        healthValue = MAX_(defaultValue * perc \ 100, 1)   ' hp not supposed to be 0
    Else
        healthValue = MIN_(atoi(health), defaultValue)
    End If
End Function

Private Sub loadSaveObject(ByVal ini As WzConfig, ByVal psObj As Long)
    Dim j As Long
    SIMPLE_OBJECT_died(psObj) = ini.value("died", jvi(0)).toInt()
    For j = 0 To MAX_PLAYERS - 1        ' memset(psObj->visible, 0, sizeof(psObj->visible))
        BASE_OBJECT_visible(psObj, j) = 0
    Next
    For j = 0 To game.maxPlayers - 1
        BASE_OBJECT_visible(psObj, j) = ini.value("visible/" & CStr(j), jvi(0)).toInt() And &HFF&
    Next
    BASE_OBJECT_periodicalDamage(psObj) = ini.value("periodicalDamage", jvi(0)).toInt()
    objSetPeriodicalDamageStart psObj, ini.value("periodicalDamageStart", jvi(0)).toInt()
    BASE_OBJECT_timeAnimationStarted(psObj) = ini.value("timeAnimationStarted", jvi(0)).toInt()
    BASE_OBJECT_animationEvent(psObj) = ini.value("animationEvent", jvi(0)).toInt() And &HFF&
    BASE_OBJECT_timeLastHit(psObj) = ini.value("timeLastHit", jvi(UDWORD_MAX)).toInt()
    objSetLastEmission psObj, ini.value("lastEmission", jvi(0)).toInt()
    objSetSelected psObj, -CLng(ini.value("selected", json_variant_new(False)).toBool())
    objSetBorn psObj, ini.value("born", jvi(2)).toInt()
End Sub

Private Sub writeSaveObject(ByVal ini As WzConfig, ByVal psObj As Long)
    Dim i As Long, p As Vector3i, r As Rotation
    ini.setValue "id", jU(SIMPLE_OBJECT_id(psObj))
    setPlayer ini, SIMPLE_OBJECT_player(psObj)
    ini.setValue "health", jU(BASE_OBJECT_body(psObj))
    p = SIMPLE_OBJECT_pos(psObj)
    ini.setVector3i "position", p.x, p.y, p.z
    r = SIMPLE_OBJECT_rot(psObj)
    ini.setVector3i "rotation", r.direction, r.pitch, r.roll     ' toVector(psObj->rot)
    If BASE_OBJECT_timeAnimationStarted(psObj) <> 0 Then
        ini.setValue "timeAnimationStarted", jU(BASE_OBJECT_timeAnimationStarted(psObj))
    End If
    If BASE_OBJECT_animationEvent(psObj) <> 0 Then
        ini.setValue "animationEvent", jU(BASE_OBJECT_animationEvent(psObj))
    End If
    ini.setValue "selected", jU(objSelected(psObj))    ' third kind of group
    If objLastEmission(psObj) <> 0 Then
        ini.setValue "lastEmission", jU(objLastEmission(psObj))
    End If
    If UGT(objPeriodicalDamageStart(psObj), 0) Then
        ini.setValue "periodicalDamageStart", jU(objPeriodicalDamageStart(psObj))
    End If
    If UGT(BASE_OBJECT_periodicalDamage(psObj), 0) Then
        ini.setValue "periodicalDamage", jU(BASE_OBJECT_periodicalDamage(psObj))
    End If
    ini.setValue "born", jU(SIMPLE_OBJECT_born(psObj))
    If UGE(SIMPLE_OBJECT_died(psObj), NOT_CURRENT_LIST) Then
        ini.setValue "died", jU(SIMPLE_OBJECT_died(psObj))
    End If
    If BASE_OBJECT_timeLastHit(psObj) <> UDWORD_MAX Then
        ini.setValue "timeLastHit", jU(BASE_OBJECT_timeLastHit(psObj))
    End If
    If objSelected(psObj) <> 0 Then
        ini.setValue "selected", jU(objSelected(psObj))
    End If
    For i = 0 To game.maxPlayers - 1
        If BASE_OBJECT_visible(psObj, i) <> 0 Then
            ini.setValue "visible/" & CStr(i), jU(BASE_OBJECT_visible(psObj, i))
        End If
    Next
End Sub

Private Sub writeSaveObjectJSON(ByVal jsonObj As nlohmann_json, ByVal psObj As Long)
    Dim i As Long, p As Vector3i, r As Rotation
    jsonObj.set_item "id", jU(SIMPLE_OBJECT_id(psObj))
    setPlayerJSON jsonObj, SIMPLE_OBJECT_player(psObj)
    jsonObj.set_item "health", jU(BASE_OBJECT_body(psObj))
    p = SIMPLE_OBJECT_pos(psObj)
    jsonObj.set_item "position", jV3(p.x, p.y, p.z)
    r = SIMPLE_OBJECT_rot(psObj)
    jsonObj.set_item "rotation", jV3(r.direction, r.pitch, r.roll)     ' toVector(psObj->rot)
    If BASE_OBJECT_timeAnimationStarted(psObj) <> 0 Then
        jsonObj.set_item "timeAnimationStarted", jU(BASE_OBJECT_timeAnimationStarted(psObj))
    End If
    If BASE_OBJECT_animationEvent(psObj) <> 0 Then
        jsonObj.set_item "animationEvent", jU(BASE_OBJECT_animationEvent(psObj))
    End If
    jsonObj.set_item "selected", jU(objSelected(psObj))    ' third kind of group
    If objLastEmission(psObj) <> 0 Then
        jsonObj.set_item "lastEmission", jU(objLastEmission(psObj))
    End If
    If UGT(objPeriodicalDamageStart(psObj), 0) Then
        jsonObj.set_item "periodicalDamageStart", jU(objPeriodicalDamageStart(psObj))
    End If
    If UGT(BASE_OBJECT_periodicalDamage(psObj), 0) Then
        jsonObj.set_item "periodicalDamage", jU(BASE_OBJECT_periodicalDamage(psObj))
    End If
    jsonObj.set_item "born", jU(SIMPLE_OBJECT_born(psObj))
    If UGE(SIMPLE_OBJECT_died(psObj), NOT_CURRENT_LIST) Then
        jsonObj.set_item "died", jU(SIMPLE_OBJECT_died(psObj))
    End If
    If BASE_OBJECT_timeLastHit(psObj) <> UDWORD_MAX Then
        jsonObj.set_item "timeLastHit", jU(BASE_OBJECT_timeLastHit(psObj))
    End If
    If objSelected(psObj) <> 0 Then
        jsonObj.set_item "selected", jU(objSelected(psObj))
    End If
    For i = 0 To game.maxPlayers - 1
        If BASE_OBJECT_visible(psObj, i) <> 0 Then
            jsonObj.set_item "visible/" & CStr(i), jU(BASE_OBJECT_visible(psObj, i))
        End If
    Next
End Sub

' template<typename T> T getCompFromName_NullCompOnFail(COMPONENT_TYPE compType, const WzString &name)
Private Function getCompFromName_NullCompOnFail(ByVal compType As COMPONENT_TYPE, ByVal name_ As String) As Long
    Dim index As Long
    index = getCompFromName(compType, name_)
    If index >= 0 Then getCompFromName_NullCompOnFail = index Else getCompFromName_NullCompOnFail = 0 ' 0 to reference the null weapon / body / etc
End Function

Public Function loadSaveDroid(ByVal pFileName As String, ppsCurrentDroidLists() As OBJ_LIST) As Boolean
    Dim fName As String, ini As WzConfig, list() As String, n As Long, i As Long, j As Long, k As Long
    Dim sortedPri() As Long, sortedName() As String, missionList As Boolean, droidType As DROID_TYPE, priority As Long
    Dim mapping() As GROUP_MAPPING, nMapping As Long, sKey As String, sPri As Long
    Dim psDroid As Long, player As Long, id As Long, pos As Vector3i, rot As Rotation, onMission As Boolean
    Dim templ As Long, psTemplate As Long, templName As String, w As Long
    Dim aigroup As Long, psGroup As Long, bFound As Boolean, numPoints As Long, tmpX As Long, tmpY As Long
    Dim tmpUnderRepair As Double, formationInfo As nlohmann_json, it_direction As nlohmann_json, it_x As nlohmann_json, it_y As nlohmann_json
    Dim formationX As Long, formationY As Long, formationDirection As Long, dr As FPATH_RETVAL, startpos As Vector2i
    Dim wr As Rotation

    If PHYSFS_exists(pFileName) = 0 Then
        debug_ 5635, LOG_SAVE, "loadSaveDroid", "No %s found -- use fallback method", pFileName
        loadSaveDroid = False   ' try to use fallback method
        Exit Function
    End If
    fName = pFileName
    Set ini = WzConfig_new(fName, WzConfig_ReadOnly)
    n = ini.childGroups(list)
    ' Sort list so transports are loaded first, since they must be loaded before the droids they contain.
    ' (VB) std::vector<std::pair<int, WzString>> sortedList: sortedPri, sortedName
    missionList = StrComp(fName, "mdroid", vbBinaryCompare) <> 0     ' (fName.compare("mdroid"): nonzero unless equal)

    ' (VB) DroidGroupIdMapping aigroupToDroidGroupMapping: mapping(0 .. nMapping - 1)
    nMapping = 0

    If n > 0 Then
        ReDim sortedPri(0 To n - 1)
        ReDim sortedName(0 To n - 1)
    End If
    For i = 0 To n - 1
        ini.beginGroup list(i)
        droidType = ini.value("droidType").toInt()
        priority = 0
        Select Case droidType
        Case DROID_TRANSPORTER, DROID_SUPERTRANSPORTER, DROID_COMMAND
            ' (the C++'s fallthrough: DROID_TRANSPORTER adds 1, then DROID_SUPERTRANSPORTER's 1, then DROID_COMMAND's)
            If droidType = DROID_TRANSPORTER Then priority = priority + 1   ' fallthrough
            If droidType = DROID_TRANSPORTER Or droidType = DROID_SUPERTRANSPORTER Then priority = priority + 1   ' fallthrough
            'Don't care about sorting commanders in the mission list for safety missions. They
            'don't have a group to command and it messes up the order of the list sorting them
            'which causes problems getting the first transporter group for Gamma-1.
            If Not missionList Or (missionList And Not getDroidsToSafetyFlag()) Then
                priority = priority + 1
            End If
        Case Else
        End Select
        sortedPri(i) = -priority
        sortedName(i) = list(i)
        ini.endGroup
    Next
    ' std::sort(sortedList.begin(), sortedList.end()): by the priority, then the name (an insertion sort)
    For i = 1 To n - 1
        sPri = sortedPri(i)
        sKey = sortedName(i)
        j = i - 1
        Do While j >= 0
            If sortedPri(j) < sPri Then Exit Do
            If sortedPri(j) = sPri Then
                If StrComp(sortedName(j), sKey, vbBinaryCompare) <= 0 Then Exit Do
            End If
            sortedPri(j + 1) = sortedPri(j)
            sortedName(j + 1) = sortedName(j)
            j = j - 1
        Loop
        sortedPri(j + 1) = sPri
        sortedName(j + 1) = sKey
    Next

    For i = 0 To n - 1
        ini.beginGroup sortedName(i)
        player = getPlayer(ini)
        id = ini.value("id", jvi(-1)).toInt()
        ini.vector3i "position", pos.x, pos.y, pos.z
        ini.vector3i "rotation", rot.direction, rot.pitch, rot.roll     ' Rotation rot = ini.vector3i("rotation")
        rot.direction = rot.direction And &HFFFF&
        rot.pitch = rot.pitch And &HFFFF&
        rot.roll = rot.roll And &HFFFF&
        onMission = ini.value("onMission", json_variant_new(False)).toBool()
        templ = 0
        psTemplate = 0

        If ini.contains("template") Then
            ' Use real template (for maps)
            templName = ini.value("template").toWzString()
            psTemplate = getTemplateFromTranslatedNameNoPlayer(templName)
            If psTemplate = 0 Then
                debug_ 5694, LOG_ERROR, "loadSaveDroid", "Unable to find template for %s for player %d -- unit skipped", templName, player
                ini.endGroup
                GoTo continue_
            End If
        Else
            ' Create fake template
            templ = newTemplate()       ' (VB) DROID_TEMPLATE templ: a pool slot, given back below
            apsTemplates(templ).name_ = ini.string_("name", "UNKNOWN")
            apsTemplates(templ).droidType = ini.value("droidType").toInt()
            apsTemplates(templ).numWeaps = ini.value("weapons", jvi(0)).toInt()
            If apsTemplates(templ).numWeaps > MAX_WEAPONS Then
                debug_ 5707, LOG_ERROR, "loadSaveDroid", "Invalid numWeaps (%d)", apsTemplates(templ).numWeaps
                apsTemplates(templ).numWeaps = MAX_WEAPONS
            End If
            ini.beginGroup "parts"      ' the following is copy-pasted from loadSaveTemplate() -- fixme somehow
            apsTemplates(templ).asParts(COMP_BODY) = getCompFromName_NullCompOnFail(COMP_BODY, ini.value("body", json_variant_new("ZNULLBODY")).toWzString()) And &HFF&
            apsTemplates(templ).asParts(COMP_BRAIN) = getCompFromName_NullCompOnFail(COMP_BRAIN, ini.value("brain", json_variant_new("ZNULLBRAIN")).toWzString()) And &HFF&
            apsTemplates(templ).asParts(COMP_PROPULSION) = getCompFromName_NullCompOnFail(COMP_PROPULSION, ini.value("propulsion", json_variant_new("ZNULLPROP")).toWzString()) And &HFF&
            apsTemplates(templ).asParts(COMP_REPAIRUNIT) = getCompFromName_NullCompOnFail(COMP_REPAIRUNIT, ini.value("repair", json_variant_new("ZNULLREPAIR")).toWzString()) And &HFF&
            apsTemplates(templ).asParts(COMP_ECM) = getCompFromName_NullCompOnFail(COMP_ECM, ini.value("ecm", json_variant_new("ZNULLECM")).toWzString()) And &HFF&
            apsTemplates(templ).asParts(COMP_SENSOR) = getCompFromName_NullCompOnFail(COMP_SENSOR, ini.value("sensor", json_variant_new("ZNULLSENSOR")).toWzString()) And &HFF&
            apsTemplates(templ).asParts(COMP_CONSTRUCT) = getCompFromName_NullCompOnFail(COMP_CONSTRUCT, ini.value("construct", json_variant_new("ZNULLCONSTRUCT")).toWzString()) And &HFF&
            apsTemplates(templ).asWeaps(0) = getCompFromName_NullCompOnFail(COMP_WEAPON, ini.value("weapon/1", json_variant_new("ZNULLWEAPON")).toWzString())
            apsTemplates(templ).asWeaps(1) = getCompFromName_NullCompOnFail(COMP_WEAPON, ini.value("weapon/2", json_variant_new("ZNULLWEAPON")).toWzString())
            apsTemplates(templ).asWeaps(2) = getCompFromName_NullCompOnFail(COMP_WEAPON, ini.value("weapon/3", json_variant_new("ZNULLWEAPON")).toWzString())
            ini.endGroup
            psTemplate = templ
        End If

        ' If droid is on a mission, calling with the saved position might cause an assertion. Or something like that.
        If Not onMission Then
            pos.x = clip_(pos.x, world_coord(1), world_coord(mapWidth - 1))
            pos.y = clip_(pos.y, world_coord(1), world_coord(mapHeight - 1))
        End If

        ' Create the Droid
        turnOffMultiMsg True
        If id > 0 Then
            psDroid = reallyBuildDroid(psTemplate, pos, player, onMission, rot, id)
        Else
            ' will generate a new id
            psDroid = reallyBuildDroid(psTemplate, pos, player, onMission, rot)
        End If
        If templ <> 0 Then deleteTemplate templ     ' (VB) the local template's end
        If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/game.cpp:5742", "loadSaveDroid", "Failed to build unit %s", sortedName(i): loadSaveDroid = False: Exit Function
        turnOffMultiMsg False
        If Not (id <> 0) Then ASSERT_FAILURE "id != 0", "src/game.cpp:5744", "loadSaveDroid", "Droid ID should never be zero here"
        ' conditional check so that existing saved games don't break
        If ini.contains("originalBody") Then
            ' we need to set "originalBody" before setting "body", otherwise CHECK_DROID throws assertion errors
            ' we cannot use droidUpgradeBody here to calculate "originalBody", because upgrades aren't loaded yet
            ' so it's much simplier just store/retrieve originalBody value
            aDroid(psDroid).originalBody = ini.value("originalBody").toInt()
        End If
        aDroid(psDroid).body = healthValue(ini, aDroid(psDroid).originalBody)
        If Not (aDroid(psDroid).body <> 0) Then ASSERT_FAILURE "psDroid->body != 0", "src/game.cpp:5754", "loadSaveDroid", "%s : %d has zero hp!", pFileName, i
        aDroid(psDroid).experience = ini.value("experience", jvi(0)).toInt()
        aDroid(psDroid).shieldPoints = ini.value("shieldPoints", jvi(-1)).toInt()
        aDroid(psDroid).shieldRegenTime = ini.value("shieldRegenTime", jvi(0)).toUInt()
        aDroid(psDroid).shieldInterruptRegenTime = ini.value("shieldInterruptRegenTime", jvi(0)).toUInt()
        aDroid(psDroid).kills = ini.value("kills", jvi(0)).toInt()
        aDroid(psDroid).secondaryOrder = ini.value("secondaryOrder", jvi(aDroid(psDroid).secondaryOrder)).toInt()
        aDroid(psDroid).secondaryOrderPending = aDroid(psDroid).secondaryOrder
        aDroid(psDroid).action = ini.value("action", jvi(DACTION_NONE)).toInt()
        ini.vector2i "action/pos", aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y
        aDroid(psDroid).actionStarted = ini.value("actionStarted", jvi(0)).toInt()
        aDroid(psDroid).actionPoints = ini.value("actionPoints", jvi(0)).toInt()
        aDroid(psDroid).resistance = ToInt16(ini.value("resistance", jvi(0)).toInt())    ' zero resistance == no electronic damage
        aDroid(psDroid).lastFrustratedTime = ini.value("lastFrustratedTime", jvi(0)).toInt()

        ' common BASE_OBJECT info
        loadSaveObject ini, psDroid

        ' copy the droid's weapon stats
        For j = 0 To aDroid(psDroid).numWeaps - 1
            If aDroid(psDroid).asWeaps(j).nStat > 0 Then
                aDroid(psDroid).asWeaps(j).ammo = ini.value("ammo/" & CStr(j)).toInt()
                aDroid(psDroid).asWeaps(j).lastFired = ini.value("lastFired/" & CStr(j)).toInt()
                aDroid(psDroid).asWeaps(j).shotsFired = ini.value("shotsFired/" & CStr(j)).toInt()
                ini.vector3i "rotation/" & CStr(j), wr.direction, wr.pitch, wr.roll
                wr.direction = wr.direction And &HFFFF&
                wr.pitch = wr.pitch And &HFFFF&
                wr.roll = wr.roll And &HFFFF&
                aDroid(psDroid).asWeaps(j).rot = wr
            End If
        Next

        aDroid(psDroid).group = ini.value("group", jvi(UBYTE_MAX)).toInt() And &HFF&
        aDroid(psDroid).repairGroup = ini.value("repairGroup", jvi(UBYTE_MAX)).toInt() And &HFF&
        aigroup = ini.value("aigroup", jvi(-1)).toInt()
        If aigroup >= 0 Then
            psGroup = 0

            ' Find mapping and add to group
            bFound = False
            For k = 0 To nMapping - 1
                If mapping(k).id = aigroup Then
                    bFound = True
                    Exit For
                End If
            Next
            If bFound Then
                psGroup = mapping(k).psGroup
                DROID_GROUP_add psGroup, psDroid
            ElseIf DROID_isTransporter(psDroid) Or aDroid(psDroid).droidType = DROID_COMMAND Then
                If Not (aDroid(psDroid).psGroup <> 0) Then ASSERT_FAILURE "psDroid->psGroup != nullptr", "src/game.cpp:5800", "loadSaveDroid", "Droid did not get created with group!: %s", sortedName(i)
                If aDroid(psDroid).psGroup <> 0 Then
                    ' Store mapping
                    ReDim Preserve mapping(0 To nMapping)
                    mapping(nMapping).id = aigroup
                    mapping(nMapping).psGroup = aDroid(psDroid).psGroup
                    nMapping = nMapping + 1
                    ' (ASSERT(result.second, "Already a mapping entry for group??"): the search above found none)
                    psGroup = aDroid(psDroid).psGroup
                End If
            Else
                debug_ 5811, LOG_ERROR, "loadSaveDroid", "Failed to find & map group (%d) for droid: %s", aigroup, sortedName(i)
            End If

            If psGroup <> 0 Then
                If aGroup(psGroup).typ = GT_TRANSPORTER Then
                    aDroid(psDroid).selected = 0    ' Droid should be visible in the transporter interface.
                    If Not DROID_isTransporter(psDroid) Then
                        visRemoveVisibility psDroid ' should not have visibility data when in a transporter
                    End If
                End If
            End If
        Else
            If DROID_isTransporter(psDroid) Or aDroid(psDroid).droidType = DROID_COMMAND Then
                If Not (aDroid(psDroid).psGroup <> 0) Then ASSERT_FAILURE "psDroid->psGroup != nullptr", "src/game.cpp:5827", "loadSaveDroid", "Droid did not get created with group!: %s", sortedName(i)
                If aDroid(psDroid).psGroup <> 0 Then
                    bFound = False
                    For k = 0 To nMapping - 1
                        If mapping(k).id = aGroup(aDroid(psDroid).psGroup).id Then bFound = True
                    Next
                    If Not (Not bFound) Then ASSERT_FAILURE "aigroupToDroidGroupMapping.count(psDroid->psGroup->id) == 0", "src/game.cpp:5830", "loadSaveDroid", "Droid created with conflicting group id (%d)!: %s", aGroup(aDroid(psDroid).psGroup).id, sortedName(i)
                End If
            Else
                aDroid(psDroid).psGroup = 0
            End If
        End If

        aDroid(psDroid).sMove.Status = ini.value("moveStatus", jvi(0)).toInt()
        aDroid(psDroid).sMove.pathIndex = ini.value("pathIndex", jvi(0)).toInt()
        numPoints = ini.value("pathLength", jvi(0)).toInt()
        ' psDroid->sMove.asPath.resize(numPoints)
        If numPoints > 0 Then
            ReDim aDroid(psDroid).sMove.asPath(0 To numPoints - 1)
        Else
            Erase aDroid(psDroid).sMove.asPath
        End If
        aDroid(psDroid).sMove.asPath_size = numPoints
        For j = 0 To numPoints - 1
            ini.vector2i "pathNode/" & CStr(j), aDroid(psDroid).sMove.asPath(j).x, aDroid(psDroid).sMove.asPath(j).y
        Next
        ini.vector2i "moveDestination", aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y
        ini.vector2i "moveSource", aDroid(psDroid).sMove.src.x, aDroid(psDroid).sMove.src.y
        ini.vector2i "moveTarget", aDroid(psDroid).sMove.target.x, aDroid(psDroid).sMove.target.y
        aDroid(psDroid).sMove.speed = ini.value("moveSpeed").toInt()
        aDroid(psDroid).sMove.moveDir = ini.value("moveDirection").toInt() And &HFFFF&
        aDroid(psDroid).sMove.bumpDir = ini.value("bumpDir").toInt() And &HFFFF&
        aDroid(psDroid).sMove.iVertSpeed = ini.value("vertSpeed").toInt()
        aDroid(psDroid).sMove.bumpTime = ini.value("bumpTime").toInt()
        aDroid(psDroid).sMove.shuffleStart = ini.value("shuffleStart").toInt()
        For j = 0 To MAX_WEAPONS - 1
            aDroid(psDroid).asWeaps(j).usedAmmo = ini.value("attackRun/" & CStr(j)).toInt()
        Next
        aDroid(psDroid).sMove.lastBump = ini.value("lastBump").toInt() And &HFFFF&
        aDroid(psDroid).sMove.pauseTime = ini.value("pauseTime").toInt() And &HFFFF&
        ini.vector2i "bumpPosition", tmpX, tmpY
        aDroid(psDroid).sMove.bumpPos.x = tmpX
        aDroid(psDroid).sMove.bumpPos.y = tmpY
        aDroid(psDroid).sMove.bumpPos.z = 0

        If DROID_isVtol(psDroid) And aDroid(psDroid).sMove.Status <> MOVEINACTIVE Then
            aDroid(psDroid).rot.pitch = 0
        End If

        tmpUnderRepair = U32(ini.value("underRepair").toUInt())
        If tmpUnderRepair > 65535# Then     ' std::numeric_limits<decltype(psDroid->underRepair)>::max() (uint16_t)
            debug_ 5873, LOG_INFO, "loadSaveDroid", "Out of bounds underRepair value: %u", tmpUnderRepair
            tmpUnderRepair = 65535#
        End If
        aDroid(psDroid).underRepair = CLng(tmpUnderRepair)

        ' recreate formation if present
        Set formationInfo = ini.value("formation").jsonValue()
        If formationInfo.is_object() Then
            Set it_direction = formationInfo.find("direction")
            Set it_x = formationInfo.find("x")
            Set it_y = formationInfo.find("y")
            If Not it_direction Is Nothing And Not it_x Is Nothing And Not it_y Is Nothing Then
                formationX = json_variant_new(it_x).toInt()
                formationY = json_variant_new(it_y).toInt()

                aDroid(psDroid).sMove.psFormation = formationFind(aDroid(psDroid).player, formationX, formationY)
                ' join a formation if it exists at the destination
                If aDroid(psDroid).sMove.psFormation <> 0 Then
                    formationJoin aDroid(psDroid).sMove.psFormation, psDroid
                Else
                    ' no formation so create a new one
                    formationDirection = json_variant_new(it_direction).toUInt()
                    If formationNew(aDroid(psDroid).sMove.psFormation, aDroid(psDroid).player, FT_LINE, formationX, formationY, formationDirection And &HFFFF&) Then
                        formationJoin aDroid(psDroid).sMove.psFormation, psDroid
                    End If
                End If
            Else
                ASSERT_FAILURE "false", "src/game.cpp:5911", "loadSaveDroid", "Missing one or more expected formation details"
            End If
        End If

        ' Recreate path-finding jobs
        If aDroid(psDroid).sMove.Status = MOVEWAITROUTE Then
            aDroid(psDroid).sMove.Status = MOVEINACTIVE
            fpathDroidRoute psDroid, aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y, FMT_MOVE
            aDroid(psDroid).sMove.Status = MOVEWAITROUTE

            ' Droid might be on a mission, so finish pathfinding now, in case pointers swap and map size changes.
            dr = fpathDroidRoute(psDroid, aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y, FMT_MOVE)
            If dr = FPR_OK Then
                aDroid(psDroid).sMove.Status = MOVENAVIGATE
                aDroid(psDroid).sMove.pathIndex = 0
            Else    ' if (retVal == FPR_FAILED)
                aDroid(psDroid).sMove.Status = MOVEINACTIVE
                actionDroid psDroid, DACTION_SULK
            End If
            If Not (dr <> FPR_WAIT) Then ASSERT_FAILURE "dr != FPR_WAIT", "src/game.cpp:5934", "loadSaveDroid", " "
        End If

        ' HACK!!
        startpos = getPlayerStartPosition(player)
        If aDroid(psDroid).droidType = DROID_CONSTRUCT And startpos.x = 0 And startpos.y = 0 Then
            scriptSetStartPos aDroid(psDroid).player, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y  ' set map start position, FIXME - save properly elsewhere!
        End If

        ' do not add to list if on a transport, then the group list is used instead
        If aDroid(psDroid).psGroup = 0 Then
            addDroid psDroid, ppsCurrentDroidLists
        ElseIf aGroup(aDroid(psDroid).psGroup).typ <> GT_TRANSPORTER Or DROID_isTransporter(psDroid) Then
            addDroid psDroid, ppsCurrentDroidLists
        End If

        ini.endGroup
continue_:
    Next
    loadSaveDroid = True
End Function

' -----------------------------------------------------------------------------------------
'
'Writes the linked list of droids for each player to a file
'
Private Function writeDroid(ByVal psCurr As Long, ByVal onMission As Boolean, counter As Long) As nlohmann_json
    Dim droidObj As nlohmann_json, partsObj As nlohmann_json, formationObj As nlohmann_json, i As Long, j As Long, numStr As String
    Dim wr As Rotation
    Set droidObj = json_object()
    droidObj.set_item "name", jS(aDroid(psCurr).aName)
    droidObj.set_item "originalBody", jU(aDroid(psCurr).originalBody)
    ' write common BASE_OBJECT info
    writeSaveObjectJSON droidObj, psCurr

    For i = 0 To aDroid(psCurr).numWeaps - 1
        If aDroid(psCurr).asWeaps(i).nStat > 0 Then
            numStr = CStr(i)
            droidObj.set_item "ammo/" & numStr, jU(aDroid(psCurr).asWeaps(i).ammo)
            droidObj.set_item "lastFired/" & numStr, jU(aDroid(psCurr).asWeaps(i).lastFired)
            droidObj.set_item "shotsFired/" & numStr, jU(aDroid(psCurr).asWeaps(i).shotsFired)
            wr = aDroid(psCurr).asWeaps(i).rot
            droidObj.set_item "rotation/" & numStr, jV3(wr.direction, wr.pitch, wr.roll)  ' toVector(psCurr->asWeaps[i].rot)
        End If
    Next
    For i = 0 To MAX_WEAPONS - 1
        setIniBaseObject droidObj, "actionTarget/" & CStr(i), aDroid(psCurr).psActionTarget(i)
    Next
    If UGT(aDroid(psCurr).lastFrustratedTime, 0) Then
        droidObj.set_item "lastFrustratedTime", jU(aDroid(psCurr).lastFrustratedTime)
    End If
    If UGT(aDroid(psCurr).experience, 0) Then
        droidObj.set_item "experience", jU(aDroid(psCurr).experience)
    End If
    If UGT(aDroid(psCurr).kills, 0) Then
        droidObj.set_item "kills", jU(aDroid(psCurr).kills)
    End If

    If aDroid(psCurr).shieldPoints > -1 Then    ' -1 is the default
        droidObj.set_item "shieldPoints", jI(aDroid(psCurr).shieldPoints)
    End If
    If UGT(aDroid(psCurr).shieldRegenTime, 0) Then
        droidObj.set_item "shieldRegenTime", jU(aDroid(psCurr).shieldRegenTime)
    End If
    If UGT(aDroid(psCurr).shieldInterruptRegenTime, 0) Then
        droidObj.set_item "shieldInterruptRegenTime", jU(aDroid(psCurr).shieldInterruptRegenTime)
    End If

    setIniDroidOrder droidObj, "order", aDroid(psCurr).order
    droidObj.set_item "orderList/size", jI(aDroid(psCurr).listSize)
    For i = 0 To aDroid(psCurr).listSize - 1
        setIniDroidOrder droidObj, "orderList/" & CStr(i), aDroid(psCurr).asOrderList(i)
    Next
    If aDroid(psCurr).timeLastHit <> UDWORD_MAX Then
        droidObj.set_item "timeLastHit", jU(aDroid(psCurr).timeLastHit)
    End If
    droidObj.set_item "secondaryOrder", jU(aDroid(psCurr).secondaryOrder)
    droidObj.set_item "action", jI(aDroid(psCurr).action)
    droidObj.set_item "actionString", jS(getDroidActionName(aDroid(psCurr).action))  ' future-proofing
    droidObj.set_item "action/pos", jV2(aDroid(psCurr).actionPos.x, aDroid(psCurr).actionPos.y)
    droidObj.set_item "actionStarted", jU(aDroid(psCurr).actionStarted)
    droidObj.set_item "actionPoints", jU(aDroid(psCurr).actionPoints)
    If aDroid(psCurr).psBaseStruct <> 0 Then
        droidObj.set_item "baseStruct/id", jU(aStruct(aDroid(psCurr).psBaseStruct).id)
        droidObj.set_item "baseStruct/player", jU(aStruct(aDroid(psCurr).psBaseStruct).player)  ' always ours, but for completeness
        droidObj.set_item "baseStruct/type", jI(aStruct(aDroid(psCurr).psBaseStruct).typ)     ' always a building, but for completeness
    End If
    If aDroid(psCurr).psGroup <> 0 Then
        droidObj.set_item "aigroup", jI(aGroup(aDroid(psCurr).psGroup).id)     ' AI and commander/transport group
        droidObj.set_item "aigroup/type", jI(aGroup(aDroid(psCurr).psGroup).typ)
    End If
    droidObj.set_item "group", jU(aDroid(psCurr).group)    ' different kind of group. of course.
    droidObj.set_item "repairGroup", jU(aDroid(psCurr).repairGroup)
    If hasCommander(psCurr) Then
        If ULE(aDroid(aGroup(aDroid(psCurr).psGroup).psCommander).died, NOT_CURRENT_LIST) Then
            droidObj.set_item "commander", jU(aDroid(aGroup(aDroid(psCurr).psGroup).psCommander).id)
        End If
    End If
    If aDroid(psCurr).resistance > 0 Then
        droidObj.set_item "resistance", jI(aDroid(psCurr).resistance)
    End If
    droidObj.set_item "droidType", jI(aDroid(psCurr).droidType)
    droidObj.set_item "weapons", jU(aDroid(psCurr).numWeaps)
    Set partsObj = json_object()
    partsObj.set_item "body", jS(asBodyStats(DROID_getBodyStats(psCurr)).id)
    partsObj.set_item "propulsion", jS(asPropulsionStats(DROID_getPropulsionStats(psCurr)).id)
    partsObj.set_item "brain", jS(asBrainStats(DROID_getBrainStats(psCurr)).id)
    partsObj.set_item "repair", jS(asRepairStats(DROID_getRepairStats(psCurr)).id)
    partsObj.set_item "ecm", jS(asECMStats(DROID_getECMStats(psCurr)).id)
    partsObj.set_item "sensor", jS(asSensorStats(DROID_getSensorStats(psCurr)).id)
    partsObj.set_item "construct", jS(asConstructStats(DROID_getConstructStats(psCurr)).id)
    For j = 0 To aDroid(psCurr).numWeaps - 1
        partsObj.set_item "weapon/" & CStr(j + 1), jS(asWeaponStats(DROID_getWeaponStats(psCurr, j)).id)
    Next
    droidObj.set_item "parts", partsObj
    droidObj.set_item "moveStatus", jI(aDroid(psCurr).sMove.Status)
    droidObj.set_item "pathIndex", jI(aDroid(psCurr).sMove.pathIndex)
    droidObj.set_item "pathLength", jU(aDroid(psCurr).sMove.asPath_size)
    For i = 0 To aDroid(psCurr).sMove.asPath_size - 1
        droidObj.set_item "pathNode/" & CStr(i), jV2(aDroid(psCurr).sMove.asPath(i).x, aDroid(psCurr).sMove.asPath(i).y)
    Next
    droidObj.set_item "moveDestination", jV2(aDroid(psCurr).sMove.destination.x, aDroid(psCurr).sMove.destination.y)
    droidObj.set_item "moveSource", jV2(aDroid(psCurr).sMove.src.x, aDroid(psCurr).sMove.src.y)
    droidObj.set_item "moveTarget", jV2(aDroid(psCurr).sMove.target.x, aDroid(psCurr).sMove.target.y)
    droidObj.set_item "moveSpeed", jI(aDroid(psCurr).sMove.speed)
    droidObj.set_item "moveDirection", jU(aDroid(psCurr).sMove.moveDir)
    droidObj.set_item "bumpDir", jU(aDroid(psCurr).sMove.bumpDir)
    droidObj.set_item "vertSpeed", jI(aDroid(psCurr).sMove.iVertSpeed)
    droidObj.set_item "bumpTime", jU(aDroid(psCurr).sMove.bumpTime)
    droidObj.set_item "shuffleStart", jU(aDroid(psCurr).sMove.shuffleStart)
    For i = 0 To MAX_WEAPONS - 1
        droidObj.set_item "attackRun/" & CStr(i), jU(aDroid(psCurr).asWeaps(i).usedAmmo)
    Next
    droidObj.set_item "lastBump", jU(aDroid(psCurr).sMove.lastBump)
    droidObj.set_item "pauseTime", jU(aDroid(psCurr).sMove.pauseTime)
    droidObj.set_item "bumpPosition", jV2(aDroid(psCurr).sMove.bumpPos.x, aDroid(psCurr).sMove.bumpPos.y)

    ' formation info
    If aDroid(psCurr).sMove.psFormation <> 0 Then
        Set formationObj = json_object()
        formationObj.set_item "direction", jU(aFormation(aDroid(psCurr).sMove.psFormation).direction)
        formationObj.set_item "x", jI(aFormation(aDroid(psCurr).sMove.psFormation).x)
        formationObj.set_item "y", jI(aFormation(aDroid(psCurr).sMove.psFormation).y)
        droidObj.set_item "formation", formationObj
    End If

    droidObj.set_item "underRepair", jU(aDroid(psCurr).underRepair)

    droidObj.set_item "onMission", jB(onMission)
    Set writeDroid = droidObj
End Function

Public Function writeDroidFile(ByVal pFileName As String, ppsCurrentDroidLists() As OBJ_LIST) As Boolean
    Dim mRoot As nlohmann_json, counter As Long, onMission As Boolean, player As Long, it As Long, psCurr As Long
    Dim it2 As Long, psTrans As Long, droidKey As String, d As nlohmann_json
    Set mRoot = json_object()
    counter = 0
    ' (&ppsCurrentDroidLists == &mission.apsDroidLists: the array passed ByRef is mission.apsDroidLists itself)
    onMission = (VarPtr(ppsCurrentDroidLists(0)) = VarPtr(mission_apsDroidLists(0)))

    For player = 0 To MAX_PLAYERS - 1
        it = ppsCurrentDroidLists(player).head
        Do While it <> 0
            psCurr = lnObj(it)
            droidKey = "droid_" & zeroPad10(counter)   ' Zero padded so that alphabetical sort works.
            counter = counter + 1
            mRoot.set_item droidKey, writeDroid(psCurr, onMission, counter)
            If DROID_isTransporter(psCurr) Then    ' if transporter save any droids in the grp
                If aDroid(psCurr).psGroup <> 0 Then
                    it2 = aGroup(aDroid(psCurr).psGroup).psList.head
                    Do While it2 <> 0
                        psTrans = lnObj(it2)
                        If psTrans <> psCurr Then
                            droidKey = "droid_" & zeroPad10(counter)   ' Zero padded so that alphabetical sort works.
                            counter = counter + 1
                            mRoot.set_item droidKey, writeDroid(psTrans, onMission, counter)
                        End If
                        it2 = lnNext(it2)
                    Loop
                End If
                'always save transporter droids that are in the mission list with an invalid value
                If VarPtr(ppsCurrentDroidLists(player)) = VarPtr(mission_apsDroidLists(player)) Then
                    Set d = mRoot.item(droidKey)
                    d.set_item "position", jV3(INVALID_XY, INVALID_XY, -1) ' Must be INVALID_XY or else unit placement could get messed up in missionResetDroids().
                End If
            End If
            it = lnNext(it)
        Loop
    Next

    writeDroidFile = saveJSONToFile(mRoot, pFileName)
End Function

' ---- the structures

' (VB) a UDWORD / SDWORD of the old binary files at byte offset off (little-endian: endian_udword does nothing)
Private Function rdU32(b() As Byte, ByVal off As Long) As Long
    CopyMemory rdU32, b(off), 4
End Function

' (VB) char name[MAX_SAVE_NAME_SIZE_V19] at byte offset off (to its NUL)
Private Function rdName(b() As Byte, ByVal off As Long, ByVal size As Long) As String
    Dim i As Long, s As String
    For i = 0 To size - 1
        If b(off + i) = 0 Then Exit For
        s = s & Chr$(b(off + i))
    Next
    rdName = s
End Function

' (VB) copies of structure.cpp's production run helpers (Private there)
Private Sub productionRunPushBack(prodRun As ProductionRun, e As ProductionRunEntry)
    If prodRun.size = 0 Then
        ReDim prodRun.items(0 To 0)
    ElseIf prodRun.size > UBound(prodRun.items) Then
        ReDim Preserve prodRun.items(0 To prodRun.size)
    End If
    prodRun.items(prodRun.size) = e
    prodRun.size = prodRun.size + 1
End Sub

Private Sub productionRunListResize(list As ProductionRunList, ByVal n As Long)
    ReDim Preserve list.runs(0 To n - 1)
    list.size = n
End Sub


Public Function loadSaveStructure(pFileData() As Byte, ByVal filesize As Long) As Boolean
    Dim base_ As Long, version As Long, quantity As Long, rec As Long, count As Long, statInc As Long
    Dim psStats As Long, psStructure As Long, bFound As Boolean, NumberOfSkippedStructures As Long
    Dim periodicalDamageTime As Long, name_ As String, id As Long, x As Long, y As Long, direction As Long, player As Long
    Dim periodicalDamageStart As Long, periodicalDamage As Long, status As Long

    ' Check the file type
    base_ = LBound(pFileData)       ' psHeader = (STRUCT_SAVEHEADER *)pFileData
    If pFileData(base_) <> Asc("s") Or pFileData(base_ + 1) <> Asc("t") Or _
       pFileData(base_ + 2) <> Asc("r") Or pFileData(base_ + 3) <> Asc("u") Then
        debug_ 6155, LOG_ERROR, "loadSaveStructure", "loadSaveStructure: Incorrect file type"

        loadSaveStructure = False
        Exit Function
    End If

    ' STRUCT_SAVEHEADER
    version = rdU32(pFileData, base_ + 4)
    quantity = rdU32(pFileData, base_ + 8)

    'increment to the start of the data
    rec = base_ + STRUCT_HEADER_SIZE

    debug_ 6167, LOG_SAVE, "loadSaveStructure", "file version is %u ", version

    ' Check the file version
    If ULT(version, VERSION_7) Or UGT(version, VERSION_8) Then
        debug_ 6172, LOG_ERROR, "loadSaveStructure", "StructLoad: unsupported save format version %d", version

        loadSaveStructure = False
        Exit Function
    End If

    If CDbl(SIZEOF_SAVE_STRUCTURE_V2) * U32(quantity) + STRUCT_HEADER_SIZE > U32(filesize) Then
        debug_ 6181, LOG_ERROR, "loadSaveStructure", "structureLoad: unexpected end of file"
        loadSaveStructure = False
        Exit Function
    End If

    ' Load in the structure data
    For count = 0 To quantity - 1
        ' memcpy(psSaveStructure, pFileData, sizeof(SAVE_STRUCTURE_V2)) (and the endian_* swaps: none on little-endian)
        name_ = rdName(pFileData, rec, MAX_SAVE_NAME_SIZE_V19)
        id = rdU32(pFileData, rec + 40)
        x = rdU32(pFileData, rec + 44)
        y = rdU32(pFileData, rec + 48)
        direction = rdU32(pFileData, rec + 56)
        player = rdU32(pFileData, rec + 60)
        periodicalDamageStart = rdU32(pFileData, rec + 68)
        periodicalDamage = rdU32(pFileData, rec + 72)
        status = pFileData(rec + 76)

        player = RemapPlayerNumber(player)

        If UGE(player, MAX_PLAYERS) Then
            player = MAX_PLAYERS - 1
            NumberOfSkippedStructures = NumberOfSkippedStructures + 1
        End If
        'get the stats for this structure
        bFound = False

        For statInc = 0 To numStructureStats - 1
            psStats = statInc
            'loop until find the same name

            If asStructureStats(psStats).id = name_ Then
                bFound = True
                Exit For
            End If
        Next
        'if haven't found the structure - ignore this record!
        If Not bFound Then
            debug_ 6235, LOG_ERROR, "loadSaveStructure", "This structure no longer exists - %s", name_   ' getSaveStructNameV19(...)
            'ignore this
            GoTo continue_
        End If

        'for modules - need to check the base structure exists
        If IsStatExpansionModule(psStats) Then
            psStructure = getTileStructure(map_coord(x), map_coord(y))
            If psStructure = 0 Then
                debug_ 6246, LOG_ERROR, "loadSaveStructure", "No owning structure for module - %s for player - %d", name_, player
                'ignore this module
                GoTo continue_
            End If
        End If

        'check not trying to build too near the edge
        If map_coord(x) < TOO_NEAR_EDGE Or map_coord(x) > mapWidth - TOO_NEAR_EDGE Then
            debug_ 6255, LOG_ERROR, "loadSaveStructure", "Structure %s, x coord too near the edge of the map. id - %d", name_, id
            'ignore this
            GoTo continue_
        End If
        If map_coord(y) < TOO_NEAR_EDGE Or map_coord(y) > mapHeight - TOO_NEAR_EDGE Then
            debug_ 6261, LOG_ERROR, "loadSaveStructure", "Structure %s, y coord too near the edge of the map. id - %d", name_, id
            'ignore this
            GoTo continue_
        End If

        psStructure = buildStructureDir(psStats, x, y, DEG(direction), player, True)
        If Not (psStructure <> 0) Then ASSERT_FAILURE "psStructure", "src/game.cpp:6267", "loadSaveStructure", "Unable to create structure"
        If psStructure = 0 Then
            GoTo continue_
        End If

        aStruct(psStructure).periodicalDamage = periodicalDamage
        periodicalDamageTime = periodicalDamageStart
        aStruct(psStructure).periodicalDamageStart = periodicalDamageTime
        aStruct(psStructure).status = status
        If aStruct(psStructure).status = SS_BUILT Then
            buildingComplete psStructure
        End If
        If asStructureStats(aStruct(psStructure).pStructureType).typ = REF_HQ Then
            scriptSetStartPos player, aStruct(psStructure).pos.x, aStruct(psStructure).pos.y
        ElseIf asStructureStats(aStruct(psStructure).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then
            scriptSetDerrickPos aStruct(psStructure).pos.x, aStruct(psStructure).pos.y
        End If
continue_:
        rec = rec + SIZEOF_SAVE_STRUCTURE_V2
    Next

    If NumberOfSkippedStructures > 0 Then
        debug_ 6293, LOG_ERROR, "loadSaveStructure", "structureLoad: invalid player number in %d structures ... assigned to the last player!" & vbLf & vbLf, NumberOfSkippedStructures
        loadSaveStructure = False
        Exit Function
    End If

    loadSaveStructure = True
End Function

' -----------------------------------------------------------------------------------------
'return id of a research topic based on the name
Private Function getResearchIdFromName(ByVal name_ As String) As Long
    Dim inc As Long
    For inc = 0 To asResearch_size - 1
        If asResearch(inc).id = name_ Then
            getResearchIdFromName = inc
            Exit Function
        End If
    Next
    debug_ 6311, LOG_ERROR, "getResearchIdFromName", "Unknown research - %s", name_
    getResearchIdFromName = NULL_ID
End Function

Public Function loadSaveStructure2(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, k As Long, j As Long
    Dim psModule As Long, capacity As Long, researchId As Long, psStructure As Long, psTileStructure As Long
    Dim player As Long, id As Long, pos As Vector3i, rot As Rotation, name_ As String, psStats As Long
    Dim moduleIdx As Long, templId As Long, pnt As Vector3i, runNum As Long, currentProd As ProductionRunEntry
    Dim tid As Long, psTempl As Long, fType As Long, fInc As Long, wr As Rotation, imdIdx As Long, blankProd As ProductionRunEntry

    If PHYSFS_exists(pFileName) = 0 Then
        debug_ 6432, LOG_SAVE, "loadSaveStructure2", "No %s found -- use fallback method", pFileName
        loadSaveStructure2 = False  ' try to use fallback method
        Exit Function
    End If
    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)

    freeAllFlagPositions        'clear any flags put in during level loads

    n = ini.childGroups(list)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        player = getPlayer(ini)
        id = ini.value("id", jvi(-1)).toInt()
        ini.vector3i "position", pos.x, pos.y, pos.z
        ini.vector3i "rotation", rot.direction, rot.pitch, rot.roll  ' Rotation rot = ini.vector3i("rotation")
        rot.direction = rot.direction And &HFFFF&
        rot.pitch = rot.pitch And &HFFFF&
        rot.roll = rot.roll And &HFFFF&
        name_ = ini.string_("name")

        'get the stats for this structure
        psStats = numStructureStats         ' std::find_if: asStructureStats + numStructureStats when not found
        For k = 0 To numStructureStats - 1
            If asStructureStats(k).id = name_ Then
                psStats = k
                Exit For
            End If
        Next
        'if haven't found the structure - ignore this record!
        If Not (psStats <> numStructureStats) Then ASSERT_FAILURE "psStats != asStructureStats + numStructureStats", "src/game.cpp:6460", "loadSaveStructure2", "This structure no longer exists - %s", name_
        If psStats = numStructureStats Then
            ini.endGroup
            GoTo continue_      ' ignore this
        End If
        'create the Structure
        'for modules - need to check the base structure exists
        If IsStatExpansionModule(psStats) Then
            psTileStructure = getTileStructure(map_coord(pos.x), map_coord(pos.y))
            If psTileStructure = 0 Then
                debug_ 6473, LOG_ERROR, "loadSaveStructure2", "No owning structure for module - %s for player - %d", name_, player
                ini.endGroup
                GoTo continue_  ' ignore this module
            End If
        End If
        'check not trying to build too near the edge
        If map_coord(pos.x) < TOO_NEAR_EDGE Or map_coord(pos.x) > mapWidth - TOO_NEAR_EDGE _
           Or map_coord(pos.y) < TOO_NEAR_EDGE Or map_coord(pos.y) > mapHeight - TOO_NEAR_EDGE Then
            debug_ 6482, LOG_ERROR, "loadSaveStructure2", "Structure %s (%s), coord too near the edge of the map", name_, list(i)
            ini.endGroup
            GoTo continue_      ' skip it
        End If
        If id <= 0 Then
            id = generateSynchronisedObjectId()
        End If
        debug_ 6490, LOG_NEVER, "loadSaveStructure2", "trying to build structure %i;%i;%s;%i;%i", id, player, asStructureStats(psStats).name_, map_coord(pos.y), map_coord(pos.y)
        psStructure = buildStructureDir(psStats, pos.x, pos.y, rot.direction, player, True, id)
        If Not (psStructure <> 0) Then ASSERT_FAILURE "psStructure", "src/game.cpp:6492", "loadSaveStructure2", "Unable to create structure"
        If psStructure = 0 Then
            ini.endGroup
            GoTo continue_
        End If
        ' common BASE_OBJECT info
        loadSaveObject ini, psStructure

        If asStructureStats(aStruct(psStructure).pStructureType).typ = REF_HQ Then
            scriptSetStartPos player, aStruct(psStructure).pos.x, aStruct(psStructure).pos.y
        End If
        aStruct(psStructure).resistance = ini.value("resistance", jvi(aStruct(psStructure).resistance)).toInt()
        capacity = ini.value("modules", jvi(0)).toInt()
        aStruct(psStructure).capacity = 0   ' increased when modules are built
        Select Case asStructureStats(aStruct(psStructure).pStructureType).typ
        Case REF_FACTORY, REF_VTOL_FACTORY, REF_CYBORG_FACTORY
            'if factory save the current build info
            With aStruct(psStructure).functionality.factory    ' (VB) psFactory: only ini's calls inside
                .productionLoops = ini.value("Factory/productionLoops", jvi(.productionLoops)).toUInt() And &HFF&
                .timeStarted = ini.value("Factory/timeStarted", jvi(.timeStarted)).toInt()
                .buildPointsRemaining = ini.value("Factory/buildPointsRemaining", jvi(.buildPointsRemaining)).toInt()
                .timeStartHold = ini.value("Factory/timeStartHold", jvi(.timeStartHold)).toInt()
                .loopsPerformed = ini.value("Factory/loopsPerformed", jvi(.loopsPerformed)).toInt() And &HFF&
                ' statusPending and pendingCount belong to the GUI, not the game state.
                .secondaryOrder = ini.value("Factory/secondaryOrder", jvi(.secondaryOrder)).toInt()
            End With
            'adjust the module structures IMD
            If capacity <> 0 Then
                psModule = getModuleStat(psStructure)
                'build the appropriate number of modules
                For moduleIdx = 0 To capacity - 1
                    buildStructure psModule, aStruct(psStructure).pos.x, aStruct(psStructure).pos.y, aStruct(psStructure).player, True
                Next
            End If
            If ini.contains("Factory/template") Then
                templId = ini.value("Factory/template").toInt()
                aStruct(psStructure).functionality.factory.psSubject = getTemplateFromMultiPlayerID(templId)
            End If
            If ini.contains("Factory/assemblyPoint/pos") Then
                ini.vector3i "Factory/assemblyPoint/pos", pnt.x, pnt.y, pnt.z
                setAssemblyPoint aStruct(psStructure).functionality.factory.psAssemblyPoint, pnt.x, pnt.y, player, False
                aFlagPos(aStruct(psStructure).functionality.factory.psAssemblyPoint).selected = ini.value("Factory/assemblyPoint/selected", json_variant_new(False)).toBool()
            End If
            If ini.contains("Factory/assemblyPoint/number") Then
                aFlagPos(aStruct(psStructure).functionality.factory.psAssemblyPoint).factoryInc = ini.value("Factory/assemblyPoint/number", jvi(42)).toInt() And &HFF&
            End If
            If player = productionPlayer Then
                fType = aFlagPos(aStruct(psStructure).functionality.factory.psAssemblyPoint).factoryType
                fInc = aFlagPos(aStruct(psStructure).functionality.factory.psAssemblyPoint).factoryInc
                If fInc >= asProductionRun(fType).size Then
                    productionRunListResize asProductionRun(fType), fInc + 1
                End If

                runNum = 0
                Do While runNum < ini.value("Factory/productionRuns", jvi(0)).toInt()
                    currentProd = blankProd     ' ProductionRunEntry currentProd (psTemplate nullptr)
                    currentProd.quantity = ini.value("Factory/Run/" & CStr(runNum) & "/quantity").toInt()
                    currentProd.built = ini.value("Factory/Run/" & CStr(runNum) & "/built").toInt()
                    If ini.contains("Factory/Run/" & CStr(runNum) & "/template") Then
                        tid = ini.value("Factory/Run/" & CStr(runNum) & "/template").toInt()
                        psTempl = getTemplateFromMultiPlayerID(tid)
                        If psTempl = 0 Then
                            ASSERT_FAILURE "psTempl", "src/game.cpp:6566", "loadSaveStructure2", "No template found for template ID %d for %s (%d)", tid, objInfo(psStructure), id
                            GoTo nextRun
                        End If
                        currentProd.psTemplate = psTempl
                    End If
                    productionRunPushBack asProductionRun(fType).runs(fInc), currentProd
nextRun:
                    runNum = runNum + 1
                Loop
            End If
            aStruct(psStructure).productToGroup = ini.value("productToGroup", jvi(UBYTE_MAX)).toInt() And &HFF&
        Case REF_RESEARCH
            'adjust the module structures IMD
            If capacity <> 0 Then
                psModule = getModuleStat(psStructure)
                buildStructure psModule, aStruct(psStructure).pos.x, aStruct(psStructure).pos.y, aStruct(psStructure).player, True
            End If
            'clear subject
            aStruct(psStructure).functionality.researchFacility.psSubject = -1
            aStruct(psStructure).functionality.researchFacility.timeStartHold = 0
            'set the subject
            If ini.contains("Research/target") Then
                researchId = getResearchIdFromName(ini.value("Research/target").toWzString())
                If researchId <> NULL_ID Then
                    aStruct(psStructure).functionality.researchFacility.psSubject = researchId
                    aStruct(psStructure).functionality.researchFacility.timeStartHold = ini.value("Research/timeStartHold").toInt()
                Else
                    debug_ 6598, LOG_ERROR, "loadSaveStructure2", "Failed to look up research target %s", ini.value("Research/target").toWzString()
                End If
            End If
        Case REF_POWER_GEN
            ' adjust the module structures IMD
            If capacity <> 0 Then
                psModule = getModuleStat(psStructure)
                buildStructure psModule, aStruct(psStructure).pos.x, aStruct(psStructure).pos.y, aStruct(psStructure).player, True
            End If
        Case REF_RESOURCE_EXTRACTOR
        Case REF_REPAIR_FACILITY
            If ini.contains("Repair/deliveryPoint/pos") Then
                ini.vector3i "Repair/deliveryPoint/pos", pnt.x, pnt.y, pnt.z
                setAssemblyPoint aStruct(psStructure).functionality.repairFacility.psDeliveryPoint, pnt.x, pnt.y, player, False
                aFlagPos(aStruct(psStructure).functionality.repairFacility.psDeliveryPoint).selected = ini.value("Repair/deliveryPoint/selected", json_variant_new(False)).toBool()
            End If
        Case REF_REARM_PAD
            aStruct(psStructure).functionality.rearmPad.timeStarted = ini.value("Rearm/timeStarted", jvi(aStruct(psStructure).functionality.rearmPad.timeStarted)).toInt()
            aStruct(psStructure).functionality.rearmPad.timeLastUpdated = ini.value("Rearm/timeLastUpdated", jvi(aStruct(psStructure).functionality.rearmPad.timeLastUpdated)).toInt()
        Case REF_WALL, REF_GATE
            aStruct(psStructure).functionality.wall.typ = ini.value("Wall/type").toInt()
            ' pIMD[std::min<unsigned>(wall.type, pIMD.size() - 1)]
            imdIdx = asStructureStats(aStruct(psStructure).pStructureType).pIMD_size - 1
            If ULT(aStruct(psStructure).functionality.wall.typ, imdIdx) Then imdIdx = aStruct(psStructure).functionality.wall.typ
            aStruct(psStructure).sDisplay.imd = asStructureStats(aStruct(psStructure).pStructureType).pIMD(imdIdx)
        Case Else
        End Select
        aStruct(psStructure).body = healthValue(ini, STRUCTURE_structureBody(psStructure))
        aStruct(psStructure).currentBuildPts = ini.value("currentBuildPts", jvi(structureBuildPointsToCompletion(psStructure))).toInt()
        If aStruct(psStructure).status = SS_BUILT Then
            Select Case asStructureStats(aStruct(psStructure).pStructureType).typ
            Case REF_POWER_GEN
                checkForResExtractors psStructure
            Case REF_RESOURCE_EXTRACTOR
                checkForPowerGen psStructure
            Case Else
                'do nothing for factories etc
            End Select
        End If
        ' weapons
        For j = 0 To asStructureStats(aStruct(psStructure).pStructureType).numWeaps - 1
            If aStruct(psStructure).asWeaps(j).nStat > 0 Then
                aStruct(psStructure).asWeaps(j).ammo = ini.value("ammo/" & CStr(j)).toInt()
                aStruct(psStructure).asWeaps(j).lastFired = ini.value("lastFired/" & CStr(j)).toInt()
                aStruct(psStructure).asWeaps(j).shotsFired = ini.value("shotsFired/" & CStr(j)).toInt()
                ini.vector3i "rotation/" & CStr(j), wr.direction, wr.pitch, wr.roll
                wr.direction = wr.direction And &HFFFF&
                wr.pitch = wr.pitch And &HFFFF&
                wr.roll = wr.roll And &HFFFF&
                aStruct(psStructure).asWeaps(j).rot = wr
            End If
        Next
        aStruct(psStructure).status = ini.value("status", jvi(SS_BUILT)).toInt()
        If aStruct(psStructure).status = SS_BUILT Then
            buildingComplete psStructure
        End If
        ini.endGroup
continue_:
    Next
    resetFactoryNumFlag     'reset flags into the masks

    loadSaveStructure2 = True
End Function

' -----------------------------------------------------------------------------------------
'
'Writes the linked list of structure for each player to a file
'
Public Function writeStructFile(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, counter As Long, player As Long, it As Long, psCurr As Long, j As Long, i As Long
    Dim t As Long, psFlag As Long, haveRun As Boolean, fType As Long, fInc As Long, runNum As Long, runSize As Long
    Dim wr As Rotation, psRepairObj As Long, psRearmObj As Long

    Set ini = WzConfig_new(pFileName, WzConfig_ReadAndWrite)
    counter = 0

    For player = 0 To MAX_PLAYERS - 1
        it = apsStructLists(player).head
        Do While it <> 0
            psCurr = lnObj(it)
            If aStruct(psCurr).pStructureType < 0 Then
                ASSERT_FAILURE "psCurr->pStructureType", "src/game.cpp:6724", "writeStructFile", "Structure has null pStructureType??"
                GoTo continue_
            End If
            ini.beginGroup "structure_" & zeroPad10(counter)   ' Zero padded so that alphabetical sort works.
            counter = counter + 1
            t = asStructureStats(aStruct(psCurr).pStructureType).typ
            ini.setValue "name", jS(asStructureStats(aStruct(psCurr).pStructureType).id)

            writeSaveObject ini, psCurr

            If aStruct(psCurr).resistance > 0 Then
                ini.setValue "resistance", jI(aStruct(psCurr).resistance)
            End If
            If aStruct(psCurr).status <> SS_BUILT Then
                ini.setValue "status", jI(aStruct(psCurr).status)
            End If
            ini.setValue "weapons", jU(aStruct(psCurr).numWeaps)
            For j = 0 To aStruct(psCurr).numWeaps - 1
                ini.setValue "parts/weapon/" & CStr(j + 1), jS(asWeaponStats(BASE_OBJECT_getWeaponStats(psCurr, j)).id)
                If aStruct(psCurr).asWeaps(j).nStat > 0 Then
                    ini.setValue "ammo/" & CStr(j), jU(aStruct(psCurr).asWeaps(j).ammo)
                    ini.setValue "lastFired/" & CStr(j), jU(aStruct(psCurr).asWeaps(j).lastFired)
                    ini.setValue "shotsFired/" & CStr(j), jU(aStruct(psCurr).asWeaps(j).shotsFired)
                    wr = aStruct(psCurr).asWeaps(j).rot
                    ini.setVector3i "rotation/" & CStr(j), wr.direction, wr.pitch, wr.roll     ' toVector(psCurr->asWeaps[j].rot)
                End If
            Next
            For i = 0 To aStruct(psCurr).numWeaps - 1
                If aStruct(psCurr).psTarget(i) <> 0 Then
                    If SIMPLE_OBJECT_died(aStruct(psCurr).psTarget(i)) = 0 Then
                        ini.setValue "target/" & CStr(i) & "/id", jU(SIMPLE_OBJECT_id(aStruct(psCurr).psTarget(i)))
                        ini.setValue "target/" & CStr(i) & "/player", jU(SIMPLE_OBJECT_player(aStruct(psCurr).psTarget(i)))
                        ini.setValue "target/" & CStr(i) & "/type", jI(SIMPLE_OBJECT_type(aStruct(psCurr).psTarget(i)))
                        ' (#ifdef DEBUG: target/<i>/debugfunc and debugline)
                    End If
                End If
            Next
            ini.setValue "currentBuildPts", jU(aStruct(psCurr).currentBuildPts)
            If aStruct(psCurr).pFunctionality <> 0 Then
                If t = REF_FACTORY Or t = REF_CYBORG_FACTORY Or t = REF_VTOL_FACTORY Then
                    ini.setValue "modules", jU(aStruct(psCurr).capacity)
                    ini.setValue "Factory/productionLoops", jU(aStruct(psCurr).functionality.factory.productionLoops)
                    ini.setValue "Factory/timeStarted", jU(aStruct(psCurr).functionality.factory.timeStarted)
                    ini.setValue "Factory/buildPointsRemaining", jI(aStruct(psCurr).functionality.factory.buildPointsRemaining)
                    ini.setValue "Factory/timeStartHold", jU(aStruct(psCurr).functionality.factory.timeStartHold)
                    ini.setValue "Factory/loopsPerformed", jU(aStruct(psCurr).functionality.factory.loopsPerformed)
                    ' statusPending and pendingCount belong to the GUI, not the game state.
                    ini.setValue "Factory/secondaryOrder", jU(aStruct(psCurr).functionality.factory.secondaryOrder)

                    If aStruct(psCurr).functionality.factory.psSubject <> 0 Then
                        ini.setValue "Factory/template", jU(apsTemplates(aStruct(psCurr).functionality.factory.psSubject).multiPlayerID)
                    End If
                    psFlag = aStruct(psCurr).functionality.factory.psAssemblyPoint
                    If psFlag <> 0 Then
                        ini.setVector3i "Factory/assemblyPoint/pos", aFlagPos(psFlag).coords.x, aFlagPos(psFlag).coords.y, aFlagPos(psFlag).coords.z
                        If aFlagPos(psFlag).selected Then
                            ini.setValue "Factory/assemblyPoint/selected", jB(aFlagPos(psFlag).selected)
                        End If
                        ini.setValue "Factory/assemblyPoint/number", jU(aFlagPos(psFlag).factoryInc)
                    End If
                    If aStruct(psCurr).functionality.factory.psCommander <> 0 Then
                        ini.setValue "Factory/commander/id", jU(aDroid(aStruct(psCurr).functionality.factory.psCommander).id)
                        ini.setValue "Factory/commander/player", jU(aDroid(aStruct(psCurr).functionality.factory.psCommander).player)
                    End If
                    ini.setValue "Factory/secondaryOrder", jU(aStruct(psCurr).functionality.factory.secondaryOrder)
                    If player = productionPlayer Then
                        ' ProductionRun emptyRun; productionRun: asProductionRun[type][inc] when it exists, else emptyRun
                        fType = aFlagPos(aStruct(psCurr).functionality.factory.psAssemblyPoint).factoryType
                        fInc = aFlagPos(aStruct(psCurr).functionality.factory.psAssemblyPoint).factoryInc
                        haveRun = fInc < asProductionRun(fType).size
                        If haveRun Then runSize = asProductionRun(fType).runs(fInc).size Else runSize = 0
                        ini.setValue "Factory/productionRuns", jI(runSize)
                        For runNum = 0 To runSize - 1
                            ini.setValue "Factory/Run/" & CStr(runNum) & "/quantity", jI(asProductionRun(fType).runs(fInc).items(runNum).quantity)
                            ini.setValue "Factory/Run/" & CStr(runNum) & "/built", jI(asProductionRun(fType).runs(fInc).items(runNum).built)
                            If asProductionRun(fType).runs(fInc).items(runNum).psTemplate <> 0 Then ini.setValue "Factory/Run/" & CStr(runNum) & "/template", _
                                        jU(apsTemplates(asProductionRun(fType).runs(fInc).items(runNum).psTemplate).multiPlayerID)
                        Next
                    Else
                        ini.setValue "Factory/productionRuns", jI(0)
                    End If
                    ini.setValue "productToGroup", jU(aStruct(psCurr).productToGroup)
                ElseIf t = REF_RESEARCH Then
                    ini.setValue "modules", jU(aStruct(psCurr).capacity)
                    ini.setValue "Research/timeStartHold", jU(aStruct(psCurr).functionality.researchFacility.timeStartHold)
                    If aStruct(psCurr).functionality.researchFacility.psSubject <> -1 Then
                        ini.setValue "Research/target", jS(asResearch(aStruct(psCurr).functionality.researchFacility.psSubject).id)
                    End If
                ElseIf t = REF_POWER_GEN Then
                    ini.setValue "modules", jU(aStruct(psCurr).capacity)
                ElseIf t = REF_REPAIR_FACILITY Then
                    psRepairObj = aStruct(psCurr).functionality.repairFacility.psObj
                    If psRepairObj <> 0 Then
                        ini.setValue "Repair/target/id", jU(SIMPLE_OBJECT_id(psRepairObj))
                        ini.setValue "Repair/target/player", jU(SIMPLE_OBJECT_player(psRepairObj))
                        ini.setValue "Repair/target/type", jI(SIMPLE_OBJECT_type(psRepairObj))
                    End If
                    psFlag = aStruct(psCurr).functionality.repairFacility.psDeliveryPoint
                    If psFlag <> 0 Then
                        ini.setVector3i "Repair/deliveryPoint/pos", aFlagPos(psFlag).coords.x, aFlagPos(psFlag).coords.y, aFlagPos(psFlag).coords.z
                        If aFlagPos(psFlag).selected Then
                            ini.setValue "Repair/deliveryPoint/selected", jB(aFlagPos(psFlag).selected)
                        End If
                    End If
                ElseIf t = REF_REARM_PAD Then
                    ini.setValue "Rearm/timeStarted", jU(aStruct(psCurr).functionality.rearmPad.timeStarted)
                    ini.setValue "Rearm/timeLastUpdated", jU(aStruct(psCurr).functionality.rearmPad.timeLastUpdated)
                    psRearmObj = aStruct(psCurr).functionality.rearmPad.psObj
                    If psRearmObj <> 0 Then
                        ini.setValue "Rearm/target/id", jU(SIMPLE_OBJECT_id(psRearmObj))
                        ini.setValue "Rearm/target/player", jU(SIMPLE_OBJECT_player(psRearmObj))
                        ini.setValue "Rearm/target/type", jI(SIMPLE_OBJECT_type(psRearmObj))
                    End If
                ElseIf t = REF_WALL Or t = REF_GATE Then
                    ini.setValue "Wall/type", jU(aStruct(psCurr).functionality.wall.typ)
                End If
            End If
            ini.endGroup
continue_:
            it = lnNext(it)
        Loop
    Next
    writeStructFile = True
End Function

' -----------------------------------------------------------------------------------------
Public Function loadSaveStructurePointers(ByVal filename As String, ppList() As OBJ_LIST) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, j As Long, psStruct As Long, player As Long, id As Long
    Dim tid As Long, tplayer As Long, ttype As OBJECT_TYPE, psCommander As Long, t As Long, key As String
    Set ini = WzConfig_new(filename, WzConfig_ReadOnly)
    n = ini.childGroups(list)

    For i = 0 To n - 1
        ini.beginGroup list(i)
        psStruct = 0
        player = getPlayer(ini)
        id = ini.value("id", jvi(-1)).toInt()
        psStruct = getBaseObjFromId_list(ppList(player), id)
        If psStruct = 0 Then
            ini.endGroup
            GoTo continue_  ' it is not unusual for a structure to 'disappear' like this; it can happen eg because of module upgrades
        End If
        For j = 0 To MAX_WEAPONS - 1
            objTrace aStruct(psStruct).id, "loadSaveStructurePointers", "weapon %d, nStat %d", j, aStruct(psStruct).asWeaps(j).nStat
            key = "target/" & CStr(j)
            If ini.contains(key & "/id") Then
                tid = ini.value(key & "/id", jvi(-1)).toInt()
                tplayer = ini.value(key & "/player", jvi(-1)).toInt()
                ttype = ini.value(key & "/type", jvi(0)).toInt()
                If tid >= 0 And tplayer >= 0 Then
                    setStructureTarget psStruct, getBaseObjFromData(tid, tplayer, ttype), j, ORIGIN_UNKNOWN
                    If Not (aStruct(psStruct).psTarget(j) <> 0) Then ASSERT_FAILURE "psStruct->psTarget[j]", "src/game.cpp:6906", "loadSaveStructurePointers", "Failed to find target"
                Else
                    If Not (tid >= 0 And tplayer >= 0) Then ASSERT_FAILURE "tid >= 0 && tplayer >= 0", "src/game.cpp:6910", "loadSaveStructurePointers", "Bad ID"
                End If
            End If
        Next
        t = asStructureStats(aStruct(psStruct).pStructureType).typ
        If ini.contains("Factory/commander/id") Then
            If Not (t = REF_FACTORY Or t = REF_CYBORG_FACTORY Or t = REF_VTOL_FACTORY) Then ASSERT_FAILURE "psStruct->pStructureType->type == REF_FACTORY || psStruct->pStructureType->type == REF_CYBORG_FACTORY || psStruct->pStructureType->type == REF_VTOL_FACTORY", "src/game.cpp:6915", "loadSaveStructurePointers", "Bad type"
            ttype = OBJ_DROID
            tid = ini.value("Factory/commander/id", jvi(-1)).toInt()
            tplayer = ini.value("Factory/commander/player", jvi(-1)).toInt()
            If tid >= 0 And tplayer >= 0 Then
                psCommander = getBaseObjFromData(tid, tplayer, ttype)
                If Not (psCommander <> 0) Then ASSERT_FAILURE "psCommander", "src/game.cpp:6924", "loadSaveStructurePointers", "Commander %d not found for building %d", tid, id
                If VarPtr(ppList(0)) = VarPtr(mission_apsStructLists(0)) Then   ' ppList == &mission.apsStructLists
                    aStruct(psStruct).functionality.factory.psCommander = psCommander
                Else
                    assignFactoryCommandDroid psStruct, psCommander
                End If
            Else
                If Not (tid >= 0 And tplayer >= 0) Then ASSERT_FAILURE "tid >= 0 && tplayer >= 0", "src/game.cpp:6936", "loadSaveStructurePointers", "Bad commander ID %d for player %d for building %d", tid, tplayer, id
            End If
        End If
        If ini.contains("Repair/target/id") Then
            If Not (t = REF_REPAIR_FACILITY) Then ASSERT_FAILURE "psStruct->pStructureType->type == REF_REPAIR_FACILITY", "src/game.cpp:6940", "loadSaveStructurePointers", "Bad type"
            ttype = ini.value("Repair/target/type", jvi(OBJ_DROID)).toInt()
            tid = ini.value("Repair/target/id", jvi(-1)).toInt()
            tplayer = ini.value("Repair/target/player", jvi(-1)).toInt()
            If tid >= 0 And tplayer >= 0 Then
                aStruct(psStruct).functionality.repairFacility.psObj = getBaseObjFromData(tid, tplayer, ttype)
                If Not (aStruct(psStruct).functionality.repairFacility.psObj <> 0) Then ASSERT_FAILURE "psRepair->psObj", "src/game.cpp:6948", "loadSaveStructurePointers", "Repair target %d not found for building %d", tid, id
            Else
                If Not (tid >= 0 And tplayer >= 0) Then ASSERT_FAILURE "tid >= 0 && tplayer >= 0", "src/game.cpp:6952", "loadSaveStructurePointers", "Bad repair ID %d for player %d for building %d", tid, tplayer, id
                aStruct(psStruct).functionality.repairFacility.psObj = 0
            End If
        End If
        If ini.contains("Rearm/target/id") Then
            If Not (t = REF_REARM_PAD) Then ASSERT_FAILURE "psStruct->pStructureType->type == REF_REARM_PAD", "src/game.cpp:6957", "loadSaveStructurePointers", "Bad type"
            ttype = OBJ_DROID   ' always, for now
            tid = ini.value("Rearm/target/id", jvi(-1)).toInt()
            tplayer = ini.value("Rearm/target/player", jvi(-1)).toInt()
            If tid >= 0 And tplayer >= 0 Then
                aStruct(psStruct).functionality.rearmPad.psObj = getBaseObjFromData(tid, tplayer, ttype)
                If Not (aStruct(psStruct).functionality.rearmPad.psObj <> 0) Then ASSERT_FAILURE "psReArmPad->psObj", "src/game.cpp:6965", "loadSaveStructurePointers", "Rearm target %d not found for building %d", tid, id
            Else
                If Not (tid >= 0 And tplayer >= 0) Then ASSERT_FAILURE "tid >= 0 && tplayer >= 0", "src/game.cpp:6969", "loadSaveStructurePointers", "Bad rearm ID %d for player %d for building %d", tid, tplayer, id
                aStruct(psStruct).functionality.rearmPad.psObj = 0
            End If
        End If

        ini.endGroup
continue_:
    Next
    loadSaveStructurePointers = True
End Function

' ---- the features

Public Function loadSaveFeature(pFileData() As Byte, ByVal filesize As Long) As Boolean
    Dim base_ As Long, version As Long, quantity As Long, rec As Long, count As Long, i As Long
    Dim pFeature As Long, psStats As Long, bFound As Boolean, sizeOfSaveFeature As Long, k As Long
    Dim name_ As String, id As Long, x As Long, y As Long, direction As Long, periodicalDamage As Long

    ' Check the file type
    base_ = LBound(pFileData)       ' psHeader = (FEATURE_SAVEHEADER *)pFileData
    If pFileData(base_) <> Asc("f") Or pFileData(base_ + 1) <> Asc("e") Or _
       pFileData(base_ + 2) <> Asc("a") Or pFileData(base_ + 3) <> Asc("t") Then
        debug_ 6995, LOG_ERROR, "loadSaveFeature", "loadSaveFeature: Incorrect file type"
        loadSaveFeature = False
        Exit Function
    End If

    ' FEATURE_SAVEHEADER
    version = rdU32(pFileData, base_ + 4)
    quantity = rdU32(pFileData, base_ + 8)

    debug_ 7003, LOG_SAVE, "loadSaveFeature", "Feature file version is %u ", version

    'increment to the start of the data
    rec = base_ + FEATURE_HEADER_SIZE

    ' Check the file version
    If ULT(version, VERSION_7) Or UGT(version, VERSION_19) Then
        debug_ 7011, LOG_ERROR, "loadSaveFeature", "Unsupported save format version %u", version
        loadSaveFeature = False
        Exit Function
    End If
    If ULT(version, VERSION_14) Then
        sizeOfSaveFeature = SIZEOF_SAVE_FEATURE_V2
    Else
        sizeOfSaveFeature = SIZEOF_SAVE_FEATURE_V14
    End If
    If CDbl(sizeOfSaveFeature) * U32(quantity) + FEATURE_HEADER_SIZE > U32(filesize) Then
        debug_ 7024, LOG_ERROR, "loadSaveFeature", "featureLoad: unexpected end of file"
        loadSaveFeature = False
        Exit Function
    End If

    ' Load in the feature data
    For count = 0 To quantity - 1
        ' psSaveFeature = (SAVE_FEATURE_V14 *) pFileData (FEATURE_SAVE_V14 is FEATURE_SAVE_V2 is OBJECT_SAVE_V19;
        ' the endian_udword swaps: none on little-endian)
        name_ = rdName(pFileData, rec, MAX_SAVE_NAME_SIZE_V19)
        id = rdU32(pFileData, rec + 40)
        x = rdU32(pFileData, rec + 44)
        y = rdU32(pFileData, rec + 48)
        direction = rdU32(pFileData, rec + 56)
        periodicalDamage = rdU32(pFileData, rec + 72)

        'get the stats for this feature
        bFound = False

        For k = 0 To asFeatureStats_size - 1
            'loop until find the same name
            If asFeatureStats(k).id = name_ Then
                psStats = k
                bFound = True
                Exit For
            End If
        Next
        'if haven't found the feature - ignore this record!
        If Not bFound Then
            debug_ 7061, LOG_ERROR, "loadSaveFeature", "This feature no longer exists - %s;%i", name_, id
            'ignore this
            GoTo continue_
        End If
        'create the Feature
        pFeature = buildFeature(psStats, x, y, True, id)
        If pFeature = 0 Then
            debug_ 7069, LOG_ERROR, "loadSaveFeature", "Unable to create feature %s", name_
            GoTo continue_
        End If
        If asFeatureStats(aFeature(pFeature).psStats).subType = FEAT_OIL_RESOURCE Then
            scriptSetDerrickPos aFeature(pFeature).pos.x, aFeature(pFeature).pos.y
        End If
        'restore values
        aFeature(pFeature).rot.direction = DEG(direction)
        aFeature(pFeature).periodicalDamage = periodicalDamage
        If UGE(version, VERSION_14) Then
            For i = 0 To MAX_PLAYERS - 1
                aFeature(pFeature).visible(i) = pFileData(rec + 76 + i)     ' psSaveFeature->visible[i]
            Next
        End If
continue_:
        rec = rec + sizeOfSaveFeature
    Next

    loadSaveFeature = True
End Function

Public Function loadSaveFeature2(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, k As Long, pFeature As Long
    Dim name_ As String, pos As Vector3i, bFound As Boolean, psStats As Long, id As Long, rot As Rotation

    If PHYSFS_exists(pFileName) = 0 Then
        debug_ 7145, LOG_SAVE, "loadSaveFeature2", "No %s found -- use fallback method", pFileName
        loadSaveFeature2 = False
        Exit Function
    End If
    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)
    n = ini.childGroups(list)
    debug_ 7150, LOG_SAVE, "loadSaveFeature2", "Loading new style features (%zu found)", n

    For i = 0 To n - 1
        ini.beginGroup list(i)
        name_ = ini.string_("name")
        ini.vector3i "position", pos.x, pos.y, pos.z
        bFound = False
        psStats = -1

        'get the stats for this feature
        For k = 0 To asFeatureStats_size - 1
            'loop until find the same name
            If asFeatureStats(k).id = name_ Then
                psStats = k
                bFound = True
                Exit For
            End If
        Next
        'if haven't found the feature - ignore this record!
        If Not bFound Then
            debug_ 7175, LOG_ERROR, "loadSaveFeature2", "This feature no longer exists - %s", name_
            'ignore this
            GoTo continue_      ' (the C++ continues without ini.endGroup())
        End If
        'create the Feature
        id = ini.value("id", jvi(-1)).toInt()
        If id > 0 Then
            pFeature = buildFeature(psStats, pos.x, pos.y, True, id)
        Else
            pFeature = buildFeature(psStats, pos.x, pos.y, True)
        End If
        If pFeature = 0 Then
            debug_ 7191, LOG_ERROR, "loadSaveFeature2", "Unable to create feature %s", name_
            GoTo continue_      ' (the C++ continues without ini.endGroup())
        End If
        If asFeatureStats(aFeature(pFeature).psStats).subType = FEAT_OIL_RESOURCE Then
            scriptSetDerrickPos aFeature(pFeature).pos.x, aFeature(pFeature).pos.y
        End If
        'restore values
        ini.vector3i "rotation", rot.direction, rot.pitch, rot.roll     ' pFeature->rot = ini.vector3i("rotation")
        rot.direction = rot.direction And &HFFFF&
        rot.pitch = rot.pitch And &HFFFF&
        rot.roll = rot.roll And &HFFFF&
        aFeature(pFeature).rot = rot
        aFeature(pFeature).player = getPlayer(ini) And &HFF&

        ' common BASE_OBJECT info
        loadSaveObject ini, pFeature

        aFeature(pFeature).body = healthValue(ini, asFeatureStats(aFeature(pFeature).psStats).body)

        ini.endGroup
continue_:
    Next
    loadSaveFeature2 = True
End Function

' -----------------------------------------------------------------------------------------
'
'Writes the linked list of features to a file
'
Public Function writeFeatureFile(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, counter As Long, it As Long, psCurr As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadAndWrite)
    counter = 0

    it = apsFeatureList(0).head
    Do While it <> 0
        psCurr = lnObj(it)
        ini.beginGroup "feature_" & zeroPad10(counter)     ' Zero padded so that alphabetical sort works.
        counter = counter + 1
        ini.setValue "name", jS(asFeatureStats(aFeature(psCurr).psStats).id)
        writeSaveObject ini, psCurr
        ini.endGroup
        it = lnNext(it)
    Loop
    writeFeatureFile = True
End Function

' ---- the templates

' loadSaveTemplate's lambda loadTemplate
Private Function loadTemplate(ByVal ini As WzConfig) As DROID_TEMPLATE
    Dim t As DROID_TEMPLATE
    DROID_TEMPLATE_construct t
    If Not loadTemplateCommon_ini(ini, t) Then
        debug_ 7241, LOG_ERROR, "loadSaveTemplate", "Stored template ""%s"" contains an unknown component.", ini.string_("name")
    End If
    t.name_ = ini.string_("name")
    t.multiPlayerID = ini.value("multiPlayerID", jvi(generateNewObjectId())).toInt()
    t.enabled = ini.value("enabled", json_variant_new(False)).toBool()
    t.stored = ini.value("stored", json_variant_new(False)).toBool()
    t.prefab = ini.value("prefab", json_variant_new(False)).toBool()
    ini.nextArrayItem
    loadTemplate = t
End Function

Public Function loadSaveTemplate(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, version As Long, player As Long
    Dim t As DROID_TEMPLATE, h As Long, k As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)
    n = ini.childGroups(list)

    version = ini.value("version", jvi(0)).toInt()
    If version = 0 Then
        loadSaveTemplate = False
        Exit Function
    End If
    For i = 0 To n - 1
        ini.beginGroup list(i)
        player = getPlayer(ini)
        ini.beginArray "templates"
        Do While ini.remainingArrayItems() > 0
            t = loadTemplate(ini)       ' (VB) std::make_unique<DROID_TEMPLATE>(loadTemplate()): a pool slot of its own
            h = newTemplate()
            apsTemplates(h) = t
            addTemplate player, h
        Loop
        ini.endArray
        ini.endGroup
    Next

    If ini.contains("localTemplates") Then
        ini.beginArray "localTemplates"
        Do While ini.remainingArrayItems() > 0
            t = loadTemplate(ini)
            localTemplates_push_back t  ' localTemplates.emplace_back(loadTemplate())
        Loop
        ini.endArray
    Else
        ' Old savegame compatibility, should remove this branch sometime.
        ' enumerateTemplates(selectedPlayer, [](DROID_TEMPLATE * psTempl) { localTemplates.push_back(*psTempl); return true; })
        For k = 0 To templateCount(selectedPlayer) - 1
            t = apsTemplates(droidTemplates_at(selectedPlayer, k))
            localTemplates_push_back t
        Next
    End If

    loadSaveTemplate = True
End Function

Private Function convGameTemplateToJSON(ByVal psCurr As Long) As nlohmann_json
    Dim templateObj As nlohmann_json
    Set templateObj = saveTemplateCommon(psCurr)
    templateObj.set_item "ref", jU(apsTemplates(psCurr).ref)
    templateObj.set_item "multiPlayerID", jU(apsTemplates(psCurr).multiPlayerID)
    templateObj.set_item "enabled", jB(apsTemplates(psCurr).enabled)
    templateObj.set_item "stored", jB(apsTemplates(psCurr).stored)
    templateObj.set_item "prefab", jB(apsTemplates(psCurr).prefab)
    Set convGameTemplateToJSON = templateObj
End Function

Public Function writeTemplateFile(ByVal pFileName As String) As Boolean
    Dim mRoot As nlohmann_json, player As Long, playerTemplatesObj As nlohmann_json, templates_array As nlohmann_json
    Dim localtemplates_array As nlohmann_json, k As Long
    Set mRoot = json_object()

    mRoot.set_item "version", jI(1)
    For player = 0 To MAX_PLAYERS - 1
        If apsDroidLists(player).size = 0 And apsStructLists(player).size = 0 Then  ' only write out templates of players that are still 'alive'
            GoTo continue_
        End If
        Set playerTemplatesObj = json_object()
        setPlayerJSON playerTemplatesObj, player
        Set templates_array = json_array()
        ' enumerateTemplates(player, [&templates_array](DROID_TEMPLATE* psTemplate) { ... push_back ...; return true; })
        For k = 0 To templateCount(player) - 1
            templates_array.push_back convGameTemplateToJSON(droidTemplates_at(player, k))
        Next
        playerTemplatesObj.set_item "templates", templates_array
        mRoot.set_item "player_" & CStr(player), playerTemplatesObj
continue_:
    Next
    Set localtemplates_array = json_array()
    For k = 0 To localTemplates_size - 1
        localtemplates_array.push_back convGameTemplateToJSON(localTemplates(k))
    Next
    mRoot.set_item "localTemplates", localtemplates_array

    writeTemplateFile = saveJSONToFile(mRoot, pFileName)
End Function

' ---- the terrain types

Public Function writeTerrainTypeMapFile(ByVal pFileName As String) As Boolean
    Dim ttypeData As TerrainTypeData, i As Long, tType As Long, mapIO As WzMapPhysFSIO, debugLoggerInstance As WzMapDebugLogger
    If Not (Len(pFileName) <> 0) Then ASSERT_FAILURE "pFileName != nullptr", "src/game.cpp:7421", "writeTerrainTypeMapFile", "pFileName is null": writeTerrainTypeMapFile = False: Exit Function

    Set ttypeData = TerrainTypeData_new()   ' (VB) WzMap::TerrainTypeData ttypeData (MapApi's factory: the class is PublicNotCreatable)
    ' (ttypeData.terrainTypes.reserve(MAX_TILE_TEXTURES))
    For i = 0 To MAX_TILE_TEXTURES - 1
        tType = terrainTypes(i)
        If tType > TER_MAX Then
            debug_ 7430, LOG_ERROR, "writeTerrainTypeMapFile", "Terrain type exceeds TER_MAX: %u", tType     ' (PRIu8)
        End If
        ttypeData.terrainTypes_push_back tType
    Next

    ' Write out the map data
    Set mapIO = New WzMapPhysFSIO
    Set debugLoggerInstance = New WzMapDebugLogger
    writeTerrainTypeMapFile = writeTerrainTypes(ttypeData, pFileName, mapIO, OutputFormat_VER3, debugLoggerInstance)  ' WzMap::LatestOutputFormat
End Function

' ---- the component and structure type lists

Public Function isComponentStateValid(ByVal state As Long) As Boolean
    isComponentStateValid = state = UNAVAILABLE _
        Or state = AVAILABLE _
        Or state = FOUND _
        Or state = REDUNDANT _
        Or state = REDUNDANT_FOUND _
        Or state = REDUNDANT_UNAVAILABLE
End Function

' -----------------------------------------------------------------------------------------
Public Function loadSaveCompList(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, player As Long, list() As String, n As Long, i As Long, name_ As String, state As Long, psComp As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)

    For player = 0 To MAX_PLAYERS - 1
        ini.beginGroup "player_" & CStr(player)
        n = ini.childKeys(list)
        For i = 0 To n - 1
            name_ = list(i)
            state = ini.value(name_, jvi(UNAVAILABLE)).toInt()
            psComp = getCompStatsFromName(name_)
            If Not (psComp <> 0) Then ASSERT_FAILURE "psComp", "src/game.cpp:7465", "loadSaveCompList", "Bad component %s", name_: loadSaveCompList = False: Exit Function
            If Not (COMPONENT_STATS_compType(psComp) >= 0 And COMPONENT_STATS_compType(psComp) <> COMP_NUMCOMPONENTS) Then ASSERT_FAILURE "psComp->compType >= 0 && psComp->compType != COMP_NUMCOMPONENTS", "src/game.cpp:7466", "loadSaveCompList", "Bad type %d", COMPONENT_STATS_compType(psComp): loadSaveCompList = False: Exit Function
            If Not (isComponentStateValid(state)) Then ASSERT_FAILURE "isComponentStateValid(state)", "src/game.cpp:7467", "loadSaveCompList", "Bad state %d for %s", state, name_: loadSaveCompList = False: Exit Function
            apCompLists(player, COMPONENT_STATS_compType(psComp), BASE_STATS_index(psComp)) = state
        Next
        ini.endGroup
    Next
    loadSaveCompList = True
End Function

' (VB) writeCompListFile's loop over one type's stats (COMPONENT_STATS *psStats = &asXStats[i]: the stats' ids)
Private Sub writeCompListType(ByVal ini As WzConfig, ByVal player As Long, ByVal compType As COMPONENT_TYPE, ids() As String, ByVal n As Long)
    Dim i As Long, state As Long
    For i = 0 To n - 1
        state = apCompLists(player, compType, i)
        If state <> UNAVAILABLE Then
            ini.setValue ids(i), jI(state)
        End If
    Next
End Sub

' -----------------------------------------------------------------------------------------
' Write out the current state of the Comp lists per player
Public Function writeCompListFile(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, player As Long, i As Long, ids() As String
    Set ini = WzConfig_new(pFileName, WzConfig_ReadAndWrite)

    ' Save each type of struct type
    For player = 0 To MAX_PLAYERS - 1
        ini.beginGroup "player_" & CStr(player)
        If asBodyStats_size > 0 Then ReDim ids(0 To asBodyStats_size - 1)
        For i = 0 To asBodyStats_size - 1: ids(i) = asBodyStats(i).id: Next
        writeCompListType ini, player, COMP_BODY, ids, asBodyStats_size
        If asWeaponStats_size > 0 Then ReDim ids(0 To asWeaponStats_size - 1)
        For i = 0 To asWeaponStats_size - 1: ids(i) = asWeaponStats(i).id: Next
        writeCompListType ini, player, COMP_WEAPON, ids, asWeaponStats_size
        If asConstructStats_size > 0 Then ReDim ids(0 To asConstructStats_size - 1)
        For i = 0 To asConstructStats_size - 1: ids(i) = asConstructStats(i).id: Next
        writeCompListType ini, player, COMP_CONSTRUCT, ids, asConstructStats_size
        If asECMStats_size > 0 Then ReDim ids(0 To asECMStats_size - 1)
        For i = 0 To asECMStats_size - 1: ids(i) = asECMStats(i).id: Next
        writeCompListType ini, player, COMP_ECM, ids, asECMStats_size
        If asPropulsionStats_size > 0 Then ReDim ids(0 To asPropulsionStats_size - 1)
        For i = 0 To asPropulsionStats_size - 1: ids(i) = asPropulsionStats(i).id: Next
        writeCompListType ini, player, COMP_PROPULSION, ids, asPropulsionStats_size
        If asSensorStats_size > 0 Then ReDim ids(0 To asSensorStats_size - 1)
        For i = 0 To asSensorStats_size - 1: ids(i) = asSensorStats(i).id: Next
        writeCompListType ini, player, COMP_SENSOR, ids, asSensorStats_size
        If asRepairStats_size > 0 Then ReDim ids(0 To asRepairStats_size - 1)
        For i = 0 To asRepairStats_size - 1: ids(i) = asRepairStats(i).id: Next
        writeCompListType ini, player, COMP_REPAIRUNIT, ids, asRepairStats_size
        If asBrainStats_size > 0 Then ReDim ids(0 To asBrainStats_size - 1)
        For i = 0 To asBrainStats_size - 1: ids(i) = asBrainStats(i).id: Next
        writeCompListType ini, player, COMP_BRAIN, ids, asBrainStats_size
        ini.endGroup
    Next
    writeCompListFile = True
End Function

' -----------------------------------------------------------------------------------------
' load up structure type list file
Public Function loadSaveStructTypeList(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, player As Long, list() As String, n As Long, i As Long, name_ As String, state As Long, statInc As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)

    For player = 0 To MAX_PLAYERS - 1
        ini.beginGroup "player_" & CStr(player)
        n = ini.childKeys(list)
        For i = 0 To n - 1
            name_ = list(i)
            state = ini.value(name_, jvi(UNAVAILABLE)).toInt()

            If Not (isComponentStateValid(state)) Then ASSERT_FAILURE "isComponentStateValid(state)", "src/game.cpp:7578", "loadSaveStructTypeList", "Bad state %d for %s", state, name_: loadSaveStructTypeList = False: Exit Function
            For statInc = 0 To numStructureStats - 1    ' loop until find the same name
                If name_ = asStructureStats(statInc).id Then
                    apStructTypeLists(player, statInc) = state
                    Exit For
                End If
            Next
            If Not (statInc <> numStructureStats) Then ASSERT_FAILURE "statInc != numStructureStats", "src/game.cpp:7589", "loadSaveStructTypeList", "Did not find structure %s", name_: loadSaveStructTypeList = False: Exit Function
        Next
        ini.endGroup
    Next
    loadSaveStructTypeList = True
End Function

' -----------------------------------------------------------------------------------------
' Write out the current state of the Struct Type List per player
Public Function writeStructTypeListFile(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, player As Long, i As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadAndWrite)

    ' Save each type of struct type
    For player = 0 To MAX_PLAYERS - 1
        ini.beginGroup "player_" & CStr(player)
        For i = 0 To numStructureStats - 1
            If apStructTypeLists(player, i) <> UNAVAILABLE Then
                ini.setValue asStructureStats(i).id, jI(apStructTypeLists(player, i))
            End If
        Next
        ini.endGroup
    Next
    writeStructTypeListFile = True
End Function

' ---- the research

Public Sub loadFixupResearchPendingStates()
    Dim players As Long, statInc As Long, plr As Long, psLab As Long
    players = game.maxPlayers
    For statInc = 0 To asResearch_size - 1
        For plr = 0 To players - 1
            ' (VB) PLAYER_RESEARCH *pPlayerRes = &asPlayerResList[plr][statInc]: asPlayerResList(plr, statInc)

            ' Handle "pending" states
            ' - i.e. starting/cancelling research was queued, but the game tick was not processed before the save occurred
            ' - (Ideally the game would prevent this from happening by ensuring the save is queued to happen after the next tick...)
            If (asPlayerResList(plr, statInc).ResearchStatus And CANCELLED_RESEARCH_PENDING) <> 0 Then
                psLab = findResearchingFacilityByResearchIndex_lists(apsStructLists, plr, statInc)
                If psLab = 0 Then
                    ' check the mission list
                    psLab = findResearchingFacilityByResearchIndex_lists(mission_apsStructLists, plr, statInc)
                End If

                If psLab <> 0 Then
                    ' Process the pending cancellation with immediate effect
                    debug_ 7643, LOG_INFO, "loadFixupResearchPendingStates", "Processing CANCELLED_RESEARCH_PENDING for: %s", asResearch(statInc).id
                    cancelResearch psLab, ModeImmediate
                Else
                    ' did not find in *either* list - log and convert the CANCELLED_RESEARCH_PENDING status appropriately
                    If asPlayerResList(plr, statInc).currentPoints = 0 Then
                        ' Reset this topic as not having been researched
                        debug_ 7652, LOG_INFO, "loadFixupResearchPendingStates", "Resetting CANCELLED_RESEARCH_PENDING to 0 for: %s", asResearch(statInc).id
                        ResetResearchStatus asPlayerResList(plr, statInc)
                    Else
                        ' Set the cancelled flag
                        debug_ 7658, LOG_INFO, "loadFixupResearchPendingStates", "Resetting CANCELLED_RESEARCH_PENDING to CANCELLED_RESEARCH for: %s", asResearch(statInc).id
                        MakeResearchCancelled asPlayerResList(plr, statInc)
                    End If
                End If
            ElseIf (asPlayerResList(plr, statInc).ResearchStatus And STARTED_RESEARCH_PENDING) <> 0 Then
                ' Possible Future TODO: Could try to "recover", and queue this research item in an available research facility (however, the save doesn't contain which facility the user queued it in)
                ' For now: Just clear the STARTED_RESEARCH_PENDING bit, so that the user can re-start the research after loading the save
                debug_ 7667, LOG_INFO, "loadFixupResearchPendingStates", "Resetting STARTED_RESEARCH_PENDING to 0 for: %s", asResearch(statInc).id
                asPlayerResList(plr, statInc).ResearchStatus = asPlayerResList(plr, statInc).ResearchStatus And Not STARTED_RESEARCH_PENDING
            End If
        Next
    Next
End Sub

' -----------------------------------------------------------------------------------------
' load up saved research file
Public Function loadSaveResearch(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, players As Long, list() As String, n As Long, i As Long, bFound As Boolean, name_ As String, statInc As Long
    Dim researchedList As nlohmann_json, possiblesList As nlohmann_json, pointsList As nlohmann_json
    Dim plr As Long, researchedVal As Long, possible As Long, points As Long, resAlreadyCompleted As Boolean
    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)
    players = game.maxPlayers
    n = ini.childGroups(list)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        bFound = False
        name_ = ini.value("name").toWzString()
        For statInc = 0 To asResearch_size - 1
            'loop until find the same name
            If asResearch(statInc).id = name_ Then
                bFound = True
                Exit For
            End If
        Next
        If Not bFound Then
            'ignore this record
            debug_ 7701, LOG_SAVE, "loadSaveResearch", "Skipping unknown research named '%s'", name_
            ini.endGroup
            GoTo continue_
        End If
        Set researchedList = ini.value("researched").jsonValue()
        Set possiblesList = ini.value("possible").jsonValue()
        Set pointsList = ini.value("currentPoints").jsonValue()
        If Not (researchedList.is_array()) Then ASSERT_FAILURE "researchedList.is_array()", "src/game.cpp:7708", "loadSaveResearch", "Bad (non-array) researched list for %s", name_
        If Not (possiblesList.is_array()) Then ASSERT_FAILURE "possiblesList.is_array()", "src/game.cpp:7709", "loadSaveResearch", "Bad (non-array) possible list for %s", name_
        If Not (pointsList.is_array()) Then ASSERT_FAILURE "pointsList.is_array()", "src/game.cpp:7710", "loadSaveResearch", "Bad (non-array) points list for %s", name_
        If Not (researchedList.size() = players) Then ASSERT_FAILURE "researchedList.size() == players", "src/game.cpp:7711", "loadSaveResearch", "Bad researched list for %s", name_
        If Not (possiblesList.size() = players) Then ASSERT_FAILURE "possiblesList.size() == players", "src/game.cpp:7712", "loadSaveResearch", "Bad possible list for %s", name_
        If Not (pointsList.size() = players) Then ASSERT_FAILURE "pointsList.size() == players", "src/game.cpp:7713", "loadSaveResearch", "Bad points list for %s", name_
        For plr = 0 To players - 1
            researchedVal = json_getValue_idx(researchedList, plr).toInt()
            possible = json_getValue_idx(possiblesList, plr).toInt()
            points = json_getValue_idx(pointsList, plr).toInt()

            ' (VB) psPlRes = &asPlayerResList[plr][statInc]: asPlayerResList(plr, statInc)
            resAlreadyCompleted = IsResearchCompleted(asPlayerResList(plr, statInc))
            ' Copy the research status
            asPlayerResList(plr, statInc).ResearchStatus = (researchedVal And RESBITS_ALL)
            SetResearchPossible asPlayerResList(plr, statInc), possible
            asPlayerResList(plr, statInc).currentPoints = points
            'for any research that has been completed - perform so that upgrade values are set up
            If researchedVal = RESEARCHED Then
                If Not resAlreadyCompleted Then
                    researchResult statInc, plr, False, 0, False
                Else
                    debug_ 7736, LOG_INFO, "loadSaveResearch", "Player %d already completed research: %s - skipping re-apply", plr, name_
                End If
            End If
        Next
        ini.endGroup
continue_:
    Next
    loadSaveResearch = True
End Function

' -----------------------------------------------------------------------------------------
' Write out the current state of the Research per player
Public Function writeResearchFile(ByVal pFileName As String) As Boolean
    Dim mRoot As nlohmann_json, possibles As nlohmann_json, researchedArr As nlohmann_json, points As nlohmann_json
    Dim i As Long, player As Long, valid As Boolean, resObj As nlohmann_json
    Set mRoot = json_object()
    ' (VB) std::vector<int> possibles, std::vector<UBYTE> researched, std::vector<UDWORD> points: the json arrays they become
    For i = 0 To asResearch_size - 1
        valid = False
        Set possibles = json_array()
        Set researchedArr = json_array()
        Set points = json_array()
        For player = 0 To game.maxPlayers - 1
            possibles.push_back jI(GetResearchPossible(asPlayerResList(player, i)))
            researchedArr.push_back jU(asPlayerResList(player, i).ResearchStatus And RESBITS_ALL)
            points.push_back jU(asPlayerResList(player, i).currentPoints)
            If IsResearchPossible(asPlayerResList(player, i)) Or (asPlayerResList(player, i).ResearchStatus And RESBITS_ALL) <> 0 Or asPlayerResList(player, i).currentPoints <> 0 Then
                valid = True    ' write this entry
            End If
        Next
        If valid Then
            Set resObj = json_object()
            resObj.set_item "name", jS(asResearch(i).id)
            resObj.set_item "possible", possibles
            resObj.set_item "researched", researchedArr
            resObj.set_item "currentPoints", points

            mRoot.set_item "research_" & CStr(i), resObj
        End If
    Next
    writeResearchFile = saveJSONToFile(mRoot, pFileName)
End Function

' ---- the messages

' -----------------------------------------------------------------------------------------
' load up saved message file
Public Function loadSaveMessage(ByVal pFileName As String, ByVal levelType As Long) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, typ As MESSAGE_TYPE, bObj As Boolean
    Dim player As Long, id As Long, dataType As Long, objId As Long, objPlayer As Long, objType As OBJECT_TYPE
    Dim psMessage As Long, psViewData As Long, pos As Vector2i, sender As Long, terrainHeight As Long

    ' Only clear the messages if its a mid save game
    If getSaveGameType() = GTYPE_SAVE_MIDMISSION Then      ' (VB) gameType: game_c's
        freeMessages
    ElseIf getSaveGameType() = GTYPE_SAVE_START Then
        ' If we are loading in a CamStart or a CamChange then we are not interested in any saved messages
        If levelType = LEVEL_TYPE_LDS_CAMSTART Or levelType = LEVEL_TYPE_LDS_CAMCHANGE Then
            loadSaveMessage = True
            Exit Function
        End If
    End If

    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)
    n = ini.childGroups(list)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        typ = ini.value("type").toInt()
        bObj = ini.contains("obj/id")
        player = ini.value("player").toInt()
        id = ini.value("id").toInt()
        dataType = ini.value("dataType").toInt()

        If typ = MSG_PROXIMITY Then
            'only load proximity if a mid-mission save game
            If getSaveGameType() = GTYPE_SAVE_MIDMISSION Then
                If bObj Then
                    ' Proximity object so create get the obj from saved idy
                    objId = ini.value("obj/id").toInt()
                    objPlayer = ini.value("obj/player").toInt()
                    objType = ini.value("obj/type").toInt()
                    psMessage = addMessage(typ, True, player)
                    If psMessage <> 0 Then
                        aMessage(psMessage).psObj = getBaseObjFromData(objId, objPlayer, objType)
                        If Not (aMessage(psMessage).psObj <> 0) Then ASSERT_FAILURE "psMessage->psObj", "src/game.cpp:7830", "loadSaveMessage", "Viewdata object id %d not found for message %d", objId, id
                    Else
                        debug_ 7834, LOG_ERROR, "loadSaveMessage", "Proximity object could not be created (type=%d, player=%d, message=%d)", typ, player, id
                    End If
                Else
                    psViewData = 0

                    ' Proximity position so get viewdata pointer from the name
                    psMessage = addMessage(typ, False, player)

                    If psMessage <> 0 Then
                        If dataType = MSG_DATA_BEACON Then
                            'See addBeaconMessage(). psMessage->dataType is wrong here because
                            'addMessage() calls createMessage() which defaults dataType to MSG_DATA_DEFAULT.
                            'Later when findBeaconMsg() attempts to find a placed beacon it can't because
                            'the dataType is wrong.
                            aMessage(psMessage).dataType = MSG_DATA_BEACON
                            ini.vector2i "position", pos.x, pos.y
                            sender = ini.value("sender").toInt()
                            psViewData = CreateBeaconViewData(sender, pos.x, pos.y)
                            If Not (psViewData <> 0) Then ASSERT_FAILURE "psViewData", "src/game.cpp:7856", "loadSaveMessage", "Could not create view data for message %d", id

                        ElseIf ini.contains("name") Then
                            psViewData = getViewData(ini.value("name").toWzString())
                            If Not (psViewData <> 0) Then ASSERT_FAILURE "psViewData", "src/game.cpp:7862", "loadSaveMessage", "Failed to find view data for proximity position - skipping message %d", id
                        Else
                            debug_ 7866, LOG_ERROR, "loadSaveMessage", "Proximity position with empty name skipped (message %d)", id
                        End If

                        If psViewData <> 0 Then
                            aMessage(psMessage).pViewData = psViewData
                            ' Check the z value is at least the height of the terrain
                            ' ((VIEW_PROXIMITY*)psViewData->pData): the view data's prox_ fields
                            terrainHeight = map_Height(aViewData(psViewData).prox_x, aViewData(psViewData).prox_y)
                            If aViewData(psViewData).prox_z < terrainHeight Then
                                aViewData(psViewData).prox_z = terrainHeight
                            End If
                        Else
                            ' Skip this message
                            removeMessage psMessage, player
                        End If
                    Else
                        debug_ 7887, LOG_ERROR, "loadSaveMessage", "Proximity position could not be created (type=%d, player=%d, message=%d)", typ, player, id
                    End If
                End If
            End If
        Else
            ' Only load Campaign/Mission messages if a mid-mission save game; always load research messages
            If typ = MSG_RESEARCH Or getSaveGameType() = GTYPE_SAVE_MIDMISSION Then
                psMessage = addMessage(typ, False, player)
                If Not (psMessage <> 0) Then ASSERT_FAILURE "psMessage", "src/game.cpp:7898", "loadSaveMessage", "Could not create message %d", id
                If psMessage <> 0 Then
                    aMessage(psMessage).pViewData = getViewData(ini.value("name").toWzString())
                    If Not (aMessage(psMessage).pViewData <> 0) Then ASSERT_FAILURE "psMessage->pViewData", "src/game.cpp:7902", "loadSaveMessage", "Failed to find view data for message %d", id
                    If aMessage(psMessage).pViewData = 0 Then
                        ' Skip this message
                        removeMessage psMessage, player
                    End If
                End If
            End If
        End If
        ini.endGroup
    Next
    jsDebugMessageUpdate
    loadSaveMessage = True
End Function

' -----------------------------------------------------------------------------------------
' Write out the current messages per player
Public Function writeMessageFile(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, numMessages As Long, player As Long, it As Long, psMessage As Long, itProx As Long, psProx As Long
    Dim pViewData As Long, psObj As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadAndWrite)
    numMessages = 0

    ' save each type of research
    For player = 0 To game.maxPlayers - 1
        If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < 11", "src/game.cpp:7927", "writeMessageFile", "player out of bounds: %d", player
        it = apsMessages(player).head
        Do While it <> 0
            psMessage = lnObj(it)
            ini.beginGroup "message_" & CStr(numMessages)
            numMessages = numMessages + 1
            ini.setValue "id", jI(numMessages - 1)     ' for future use
            ini.setValue "player", jI(player)
            ini.setValue "type", jI(aMessage(psMessage).typ)
            ini.setValue "dataType", jI(aMessage(psMessage).dataType)
            If aMessage(psMessage).typ = MSG_PROXIMITY Then
                'get the matching proximity message
                itProx = apsProxDisp(player).head
                Do While itProx <> 0
                    If aProxDisp(lnObj(itProx)).psMessage = psMessage Then Exit Do  ' compare the pointers
                    itProx = lnNext(itProx)
                Loop
                If itProx <> 0 Then
                    psProx = lnObj(itProx)
                    If psProx <> 0 And aProxDisp(psProx).typ = POS_PROXDATA Then
                        'message has viewdata so store the name
                        pViewData = aMessage(psMessage).pViewData
                        If pViewData <> 0 Then ini.setValue "name", jS(aViewData(pViewData).name) Else ini.setValue "name", jS("NULL")

                        ' save beacon data
                        If aMessage(psMessage).dataType = MSG_DATA_BEACON Then
                            ' const VIEW_PROXIMITY *viewData: the view data's prox_ fields, when it has them
                            If aMessage(psMessage).pViewData <> 0 Then
                                If aViewData(aMessage(psMessage).pViewData).hasProxData Then
                                    ini.setVector2i "position", aViewData(aMessage(psMessage).pViewData).prox_x, aViewData(aMessage(psMessage).pViewData).prox_y
                                    ini.setValue "sender", jI(aViewData(aMessage(psMessage).pViewData).prox_sender)
                                End If
                            End If
                        End If
                    Else
                        ' message has object so store Object Id
                        psObj = aMessage(psMessage).psObj
                        If psObj <> 0 Then
                            ini.setValue "obj/id", jU(SIMPLE_OBJECT_id(psObj))
                            ini.setValue "obj/player", jU(SIMPLE_OBJECT_player(psObj))
                            ini.setValue "obj/type", jI(SIMPLE_OBJECT_type(psObj))
                        Else
                            ASSERT_FAILURE "false", "src/game.cpp:7974", "writeMessageFile", "Message type has no object data to save ?"
                        End If
                    End If
                Else
                    ASSERT_FAILURE "psProxIt != apsProxDisp[player].end()", "src/game.cpp:7980", "writeMessageFile", "Save message; proximity display not found for message"
                End If
            Else
                pViewData = aMessage(psMessage).pViewData
                If pViewData <> 0 Then ini.setValue "name", jS(aViewData(pViewData).name) Else ini.setValue "name", jS("NULL")
            End If
            ini.setValue "read", jB(aMessage(psMessage).read)    ' flag to indicate whether message has been read; not that this is/was _not_ read by loading code!?
            If Not (player = aMessage(psMessage).player) Then ASSERT_FAILURE "player == psMessage->player", "src/game.cpp:7989", "writeMessageFile", "Bad player number (%d == %d)", player, aMessage(psMessage).player
            ini.endGroup
            it = lnNext(it)
        Loop
    Next
    writeMessageFile = True
End Function

' ---- the structure limits

Public Function loadSaveLimits(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, player As Long, list() As String, n As Long, i As Long, name_ As String, limit As Long, statInc As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)

    For player = 0 To game.maxPlayers - 1
        ini.beginGroup "player_" & CStr(player)
        n = ini.childKeys(list)
        For i = 0 To n - 1
            name_ = list(i)
            limit = ini.value(name_, jvi(0)).toInt()

            If name_ = "@Droid" Then
                setMaxDroids player, limit
            ElseIf name_ = "@Commander" Then
                setMaxCommanders player, limit
            ElseIf name_ = "@Constructor" Then
                setMaxConstructors player, limit
            Else
                For statInc = 0 To numStructureStats - 1
                    If name_ = asStructureStats(statInc).id Then
                        If limit <> 255 Then asStructureStats(statInc).upgrade(player).limit = limit Else asStructureStats(statInc).upgrade(player).limit = LOTS_OF
                        Exit For
                    End If
                Next
                If Not (statInc <> numStructureStats) Then ASSERT_FAILURE "statInc != numStructureStats", "src/game.cpp:8034", "loadSaveLimits", "Did not find structure %s", name_: loadSaveLimits = False: Exit Function
            End If
        Next
        ini.endGroup
    Next
    loadSaveLimits = True
End Function

' -----------------------------------------------------------------------------------------
'
'Writes the list of structure limits to a file
'
Public Function writeStructLimitsFile(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, player As Long, i As Long, limit As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadAndWrite)

    ' Save each type of struct type
    For player = 0 To game.maxPlayers - 1
        ini.beginGroup "player_" & CStr(player)

        ini.setValue "@Droid", jI(getMaxDroids(player))
        ini.setValue "@Commander", jI(getMaxCommanders(player))
        ini.setValue "@Constructor", jI(getMaxConstructors(player))

        For i = 0 To numStructureStats - 1
            ' MIN(asStructureStats[i].upgrade[player].limit, 255): the unsigned limit (LOTS_OF is UINT32_MAX)
            If ULT(asStructureStats(i).upgrade(player).limit, 255) Then limit = asStructureStats(i).upgrade(player).limit Else limit = 255
            If limit <> 255 Then
                ini.setValue asStructureStats(i).id, jI(limit)
            End If
        Next
        ini.endGroup
    Next
    writeStructLimitsFile = True
End Function

' ---- the fire support designators

'
' Load the current fire-support designated commanders (the one who has fire-support enabled)
'
Public Function readFiresupportDesignators(ByVal pFileName As String) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, id As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnly)
    n = ini.childGroups(list)

    For i = 0 To n - 1
        id = ini.value("Player_" & CStr(i) & "/id", jvi(NULL_ID)).toInt()
        If id <> NULL_ID Then
            cmdDroidSetDesignator getBaseObjFromId(id)
        End If
    Next
    readFiresupportDesignators = True
End Function

'
' Save the current fire-support designated commanders (the one who has fire-support enabled)
'
Public Function writeFiresupportDesignators(ByVal pFileName As String) As Boolean
    Dim player As Long, ini As WzConfig, psDroid As Long
    Set ini = WzConfig_new(pFileName, WzConfig_ReadAndWrite)

    For player = 0 To MAX_PLAYERS - 1
        psDroid = cmdDroidGetDesignator(player)
        If psDroid <> 0 Then
            ini.setValue "Player_" & CStr(player) & "/id", jU(aDroid(psDroid).id)
        End If
    Next
    writeFiresupportDesignators = True
End Function

' ---- the script state

' -----------------------------------------------------------------------------------------
' write the event state to a file on disk
Public Function writeScriptState(ByVal pFileName As String) As Boolean
    Dim jsFilename As String, ext As Long

    ' The below belongs to the new javascript stuff
    jsFilename = pFileName
    ext = InStrRev(jsFilename, "/")         ' ext = strrchr(jsFilename, '/'); *ext = '\0'
    jsFilename = Left$(jsFilename, ext - 1)
    jsFilename = jsFilename & "/scriptstate.json"
    saveScriptStates jsFilename

    writeScriptState = True
End Function

' -----------------------------------------------------------------------------------------
' load the script state given a .gam name
' (VB) the C++ changes the caller's pFileName to <name>/scriptstate.es; the stand-in's (and levLoadData's) ByVal
' keeps it: levLoadData only logs it, when this fails, which it never does
Public Function loadScriptState(ByVal pFileName As String) As Boolean
    Dim jsFilename As String

    pFileName = Left$(pFileName, Len(pFileName) - 4)  ' pFileName[strlen(pFileName) - 4] = '\0'

    ' The below belongs to the new javascript stuff
    jsFilename = pFileName
    jsFilename = jsFilename & "/scriptstate.json"
    loadScriptStates jsFilename

    ' change the file extension
    pFileName = pFileName & "/scriptstate.es"

    loadScriptState = True
End Function

' ---- the game guide's topics

' -----------------------------------------------------------------------------------------
' load / save the game guide topics
Public Function writeSaveGuideTopics(ByVal pFileName As String) As Boolean
    Dim loadedTopics() As String, nTopics As Long, mRoot As nlohmann_json, topics As nlohmann_json, options As nlohmann_json, k As Long
    nTopics = saveLoadedGuideTopics(loadedTopics)
    Set mRoot = json_object()
    Set topics = json_array()       ' std::vector<std::string> as a json array
    For k = 0 To nTopics - 1
        topics.push_back jS(loadedTopics(k))
    Next
    mRoot.set_item "loaded_topics", topics

    Set options = json_object()
    options.set_item "disable_topic_popups", jB(getGameGuideDisableTopicPopups())
    mRoot.set_item "options", options

    writeSaveGuideTopics = saveJSONToFile(mRoot, pFileName)
End Function

Public Function loadSaveGuideTopics(ByVal pFileName As String) As Boolean
    Dim saveGuideTopicsOpt As nlohmann_json, mRoot As nlohmann_json, it As nlohmann_json, options As nlohmann_json, opt As nlohmann_json
    Dim loadedTopics() As String, nTopics As Long, k As Long, errDescription As String

    If PHYSFS_exists(pFileName) = 0 Then
        loadSaveGuideTopics = True  ' older saves will have this file - expected
        Exit Function
    End If

    If Not parseJsonFile(pFileName, saveGuideTopicsOpt) Then     ' nonstd::optional<nlohmann::json>: false is nullopt
        ASSERT_FAILURE "saveGuideTopicsOpt.has_value()", "src/game.cpp:8173", "loadSaveGuideTopics", "guidetopics.json looks broken"
        loadSaveGuideTopics = False
        Exit Function
    End If

    Set mRoot = saveGuideTopicsOpt
    If Not (mRoot.is_object()) Then ASSERT_FAILURE "mRoot.is_object()", "src/game.cpp:8178", "loadSaveGuideTopics", "guidetopics.json looks broken - expected object": loadSaveGuideTopics = False: Exit Function

    Set it = mRoot.find("options")
    If Not it Is Nothing Then
        If it.is_object() Then
            Set options = it
            Set opt = options.find("disable_topic_popups")
            If Not opt Is Nothing Then
                If opt.is_boolean() Then
                    setGameGuideDisableTopicPopups opt.get_bool()
                End If
            End If
        End If
    End If

    Set it = mRoot.find("loaded_topics")
    If Not (Not it Is Nothing) Then ASSERT_FAILURE "it != mRoot.end()", "src/game.cpp:8192", "loadSaveGuideTopics", "guidetopics.json looks broken - missing expected ""loaded_topics"" key": loadSaveGuideTopics = False: Exit Function

    ' loadedTopics = it.value().get<std::vector<std::string>>(), its exceptions caught
    On Error GoTo invalid
    If Not it.is_array() Then json_hpp_throw_type_error it
    nTopics = it.size()
    If nTopics > 0 Then ReDim loadedTopics(0 To nTopics - 1)
    For k = 0 To nTopics - 1
        loadedTopics(k) = it.at_idx(k).get_string()
    Next
    On Error GoTo 0

    loadSaveGuideTopics = restoreLoadedGuideTopics(loadedTopics, nTopics)
    Exit Function
invalid:
    errDescription = Err.Description
    On Error GoTo 0
    ' Failed to parse the JSON
    debug_ 8200, LOG_ERROR, "loadSaveGuideTopics", "JSON document from %s is invalid: %s", pFileName, errDescription
    loadSaveGuideTopics = False
End Function

' (VB) get<std::vector<std::string>>() of a value that is no array: nlohmann's type_error 302, as an error
Private Sub json_hpp_throw_type_error(ByVal j As nlohmann_json)
    Err.Raise vbObjectError + 302, "json", "[json.exception.type_error.302] type must be array, but is " & j.type_name()
End Sub
