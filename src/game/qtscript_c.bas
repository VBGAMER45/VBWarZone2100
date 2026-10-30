Attribute VB_Name = "qtscript_c"
' src/qtscript.cpp, qtscript.h: the script engine (class scripting_engine):
' the script instances, timers, labels, groups and the events the game
' triggers. The singleton's members are this module's variables; its
' methods are scripting_engine_<name> (the static ones, which become JS
' functions, take the execution context first, as C1's signature rules say).
' A wzapi::scripting_instance * is a ChakraScriptingInstance (the backend:
' chakra_backend_c); createQuickJSScriptInstance is createChakraScriptInstance.
' * std::map<std::string, LABEL> labels: sorted arrays (labels_keys,
'   labels_values; binary compare of the keys, as std::string's). A loop that
'   calls out (areaLabelCheck) goes on from the next key after the current
'   one, as a std::map iterator does when scripts add labels meanwhile.
' * std::map<scripting_instance *, GROUPMAP *> groups (ENGINEMAP): arrays in
'   the order the instances were made, which stands for the order of their
'   addresses. GROUPMAP is a class (GROUPMAP.cls).
' * std::list<std::shared_ptr<timerNode>> timers: an array of timerNode
'   objects (timerNode.cls) in list order; timerIDMap a Scripting.Dictionary
'   of timer id -> node. removeTimersIf's predicate (a lambda) is a number:
'   TIMERPRED_*.
' * executeWithPerformanceMonitoring (a template taking a lambda) is
'   scripting_engine_executeWithPerformanceMonitoring_begin before the call
'   and _end after it.
' * std::unordered_set<uint16_t> derricks: a Boolean array over the 65536
'   values plus a count; constructDerrickPositions lists them in ascending
'   order (no script of the game data reads derrickPositions, so MSVC's hash
'   order is not reproduced).
' * WzString::normalized(NormalizationForm_KD) is Windows' NormalizeString;
'   WzPathInfo's baseName and path (wzpaths.cpp) are private helpers here.
' * The debug interfaces (scripting_engine::DebugInterface, wzscriptdebug's
'   dialog) wait for cluster G: jsShowDebug hands the dialog stand-in no
'   interface; the snapshot functions are here for it.
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function NormalizeString Lib "normaliz.dll" (ByVal NormForm As Long, ByVal lpSrcString As Long, ByVal cwSrcLength As Long, ByVal lpDstString As Long, ByVal cwDstLength As Long) As Long

Private Const NormalizationKD As Long = 6

' qtscript.h: enum SCRIPT_TRIGGER_TYPE
Public Enum SCRIPT_TRIGGER_TYPE
    TRIGGER_GAME_INIT
    TRIGGER_START_LEVEL
    TRIGGER_TRANSPORTER_ARRIVED
    TRIGGER_TRANSPORTER_EMBARKED
    TRIGGER_TRANSPORTER_DISEMBARKED
    TRIGGER_TRANSPORTER_LANDED
    TRIGGER_TRANSPORTER_LAUNCH
    TRIGGER_TRANSPORTER_EXIT
    TRIGGER_TRANSPORTER_DONE
    TRIGGER_DELIVERY_POINT_MOVING
    TRIGGER_DELIVERY_POINT_MOVED
    TRIGGER_VIDEO_QUIT
    TRIGGER_MISSION_TIMEOUT
    TRIGGER_GAME_LOADED
    TRIGGER_GAME_SAVING
    TRIGGER_GAME_SAVED
    TRIGGER_DESIGN_BODY
    TRIGGER_DESIGN_WEAPON
    TRIGGER_DESIGN_COMMAND
    TRIGGER_DESIGN_SYSTEM
    TRIGGER_DESIGN_PROPULSION
    TRIGGER_DESIGN_QUIT
    TRIGGER_MENU_DESIGN_UP
    TRIGGER_MENU_BUILD_UP
    TRIGGER_MENU_BUILD_SELECTED
    TRIGGER_MENU_MANUFACTURE_UP
    TRIGGER_MENU_RESEARCH_UP
    TRIGGER_MENU_RESEARCH_SELECTED
    TRIGGER_OBJECT_RECYCLED
End Enum

' qtscript.h: enum SCRIPT_TYPE
Public Enum SCRIPT_TYPE
    SCRIPT_POSITION = OBJ_NUM_TYPES
    SCRIPT_AREA
    SCRIPT_PLAYER
    SCRIPT_RESEARCH
    SCRIPT_GROUP
    SCRIPT_RADIUS
    SCRIPT_COUNT
End Enum

' qtscript.h: struct LABEL (SCRIPT_LABEL: VB has a Label control class)
Public Type SCRIPT_LABEL
    p1 As Vector2i                      ' world coordinates
    p2 As Vector2i                      ' world coordinates
    id As Long
    typ As Long                         ' int type
    player As Long
    subscriber As Long
    idlist() As Long                    ' std::vector<int>: idlist(0 .. idlist_size - 1)
    idlist_size As Long
    triggered As Long
End Type

' qtscript.h: struct generic_script_object (all coordinates are in *map* coordinates)
Public Type generic_script_object
    p1 As Vector2i
    p2 As Vector2i
    id As Long
    player As Long
    typ As Long                         ' int type
End Type

' qtscript.h: struct scripting_engine::area_by_values_or_area_label_lookup
Public Type area_by_values_or_area_label_lookup
    m_isLabel As Boolean
    m_label As String
    x1 As Long                          ' -1
    y1 As Long
    x2 As Long
    y2 As Long
End Type

' qtscript.h: struct scripting_engine::timerNodeSnapshot (for the debug dialog)
Public Type timerNodeSnapshot
    timerID As Double
    timerName As String
    instance As ChakraScriptingInstance
    baseobj As Long
    baseobjtype As Long                 ' OBJECT_TYPE
    frameTime As Long
    ms As Long
    player As Long
    calls As Long
    typ As timerType
    instanceTimerRestoreData As nlohmann_json
End Type

' qtscript.h: struct scripting_engine::LabelInfo (for the debug dialog)
Public Type LabelInfo
    label As String
    typ As String                       ' WzString type
    trigger As String
    owner As String
    subscriber As String
End Type

' (VB) removeTimersIf's predicates
Public Const TIMERPRED_ONESHOT_DONE As Long = 0         ' node.type == TIMER_ONESHOT_DONE (updateScripts)
Public Const TIMERPRED_BASEOBJ As Long = 1              ' node.baseobj == baseobj (scriptRemoveObject)
Public Const TIMERPRED_INSTANCE_NAME_PLAYER As Long = 2 ' node.instance == instance && node.timerName == timerName && node.player == player (js_removeTimer)

' (VB) jsDebugCreate's shutdown handler (jsDebugShutdownHandlerFunction): its number
Public Const JSDEBUG_SHUTDOWN_jsHandleDebugClosed As Long = 1

Private Const ATTACK_THROTTLE As Long = 1000

Private Const MAX_US As Long = 20000
Private Const HALF_MAX_US As Long = 10000

' (ALL_PLAYERS, ALLIES, ENEMIES: wzapi_c's, the same #defines)

' typedef struct monitor_bin
Private Type MONITOR_BIN
    worst As Long
    worstGameTime As Long               ' uint32_t
    calls As Long
    overMaxTimeCalls As Long
    overHalfMaxTimeCalls As Long
    time_ As Double                     ' uint64_t time
End Type

' Structure for research events put on hold
Private Type researchEvent
    research As Long                    ' RESEARCH *
    structure As Long                   ' STRUCTURE *
    player As Long
End Type

' scripting_engine's members
Private labels_keys() As String         ' LABELMAP labels
Private labels_values() As SCRIPT_LABEL
Private labels_size As Long

Private groups_keys() As ChakraScriptingInstance ' ENGINEMAP groups
Private groups_values() As GROUPMAP
Private groups_size As Long

' List of timer events for scripts. Before running them, we sort the list then run as many as we have time for.
' In this way, we implement load balancing of events and keep frame rates tidy for users. Since scripts run on the
' host, we do not need to worry about each peer simulating the world differently.
Private timers() As timerNode           ' std::list<std::shared_ptr<timerNode>>
Private timers_size As Long
Private lastTimerID As Double           ' uniqueTimerID
Private timerIDMap_ As Object           ' std::unordered_map<uniqueTimerID, iterator>: timer id -> timerNode

' selection changes are too often and too erratic to trigger immediately,
' so until we have a queue system for events, delay triggering this way.
Private selectionChanged As Boolean

' Scripting engine (what others call the scripting context, but QtScript's nomenclature is different).
Private scripts() As ChakraScriptingInstance
Private scripts_size As Long

' Whether the scripts have been set up or not
Private scriptsReady As Boolean

' Research events that are put on hold until the scripts are ready
Private eventQueue() As researchEvent   ' std::queue<struct researchEvent>
Private eventQueue_size As Long

' static std::unordered_map<wzapi::scripting_instance *, MONITOR *> monitors: per instance a Dictionary of
' function name -> its MONITOR_BIN's index in monitorBins
Private monitors_keys() As ChakraScriptingInstance
Private monitors_values() As Object
Private monitors_size As Long
Private monitorBins() As MONITOR_BIN
Private monitorBins_size As Long

Private globalDialog As Boolean

Public bInTutorial As Boolean

' ----------------------------------------------------------

Public positions(0 To MAX_PLAYERS - 1) As Vector2i
Private derricks(0 To 65535) As Boolean
Private derricks_size As Long

'--------------------------------------------------------------------------
' (VB) the containers

' timerIDMap, made on first use
Private Function timerIDMap() As Object
    If timerIDMap_ Is Nothing Then Set timerIDMap_ = CreateObject("Scripting.Dictionary")
    Set timerIDMap = timerIDMap_
End Function

' labels: the index of the first key not less than key
Private Function labels_lower_bound(ByVal key As String) As Long
    Dim lo As Long, hi As Long, m As Long
    lo = 0
    hi = labels_size
    Do While lo < hi
        m = (lo + hi) \ 2
        If StrComp(labels_keys(m), key, vbBinaryCompare) < 0 Then lo = m + 1 Else hi = m
    Loop
    labels_lower_bound = lo
End Function

' labels: the index of the first key greater than key
Private Function labels_upper_bound(ByVal key As String) As Long
    Dim i As Long
    i = labels_lower_bound(key)
    If i < labels_size Then
        If labels_keys(i) = key Then i = i + 1
    End If
    labels_upper_bound = i
End Function

' labels.find: the index, or -1 (labels.count(key) > 0 is labels_find(key) >= 0)
Private Function labels_find(ByVal key As String) As Long
    Dim i As Long
    i = labels_lower_bound(key)
    labels_find = -1
    If i < labels_size Then
        If labels_keys(i) = key Then labels_find = i
    End If
End Function

' labels[key]: the index, a default LABEL inserted when the key is new
Private Function labels_index(ByVal key As String) As Long
    Dim i As Long, j As Long, blank As SCRIPT_LABEL
    i = labels_lower_bound(key)
    If i < labels_size Then
        If labels_keys(i) = key Then
            labels_index = i
            Exit Function
        End If
    End If
    If labels_size = 0 Then
        ReDim labels_keys(0 To 7)
        ReDim labels_values(0 To 7)
    ElseIf labels_size > UBound(labels_keys) Then
        ReDim Preserve labels_keys(0 To 2 * labels_size - 1)
        ReDim Preserve labels_values(0 To 2 * labels_size - 1)
    End If
    For j = labels_size To i + 1 Step -1
        labels_keys(j) = labels_keys(j - 1)
        labels_values(j) = labels_values(j - 1)
    Next
    labels_keys(i) = key
    labels_values(i) = blank            ' LABEL(): p1, p2 (0, 0)
    labels_size = labels_size + 1
    labels_index = i
End Function

' labels.erase(key): the number erased
Private Function labels_erase(ByVal key As String) As Long
    Dim i As Long, j As Long, blank As SCRIPT_LABEL
    i = labels_find(key)
    If i < 0 Then Exit Function
    For j = i To labels_size - 2
        labels_keys(j) = labels_keys(j + 1)
        labels_values(j) = labels_values(j + 1)
    Next
    labels_size = labels_size - 1
    labels_keys(labels_size) = vbNullString
    labels_values(labels_size) = blank
    labels_erase = 1
End Function

Private Sub labels_clear()
    Erase labels_keys
    Erase labels_values
    labels_size = 0
End Sub

' LABEL::operator==
Private Function SCRIPT_LABEL_equals(a As SCRIPT_LABEL, b As SCRIPT_LABEL) As Boolean
    SCRIPT_LABEL_equals = a.id = b.id And a.typ = b.typ And a.player = b.player
End Function

Private Sub SCRIPT_LABEL_idlist_push_back(l As SCRIPT_LABEL, ByVal id As Long)
    If l.idlist_size = 0 Then
        ReDim l.idlist(0 To 7)
    ElseIf l.idlist_size > UBound(l.idlist) Then
        ReDim Preserve l.idlist(0 To 2 * l.idlist_size - 1)
    End If
    l.idlist(l.idlist_size) = id
    l.idlist_size = l.idlist_size + 1
End Sub

' groups.find: the index, or -1
Private Function groups_find(ByVal instance As ChakraScriptingInstance) As Long
    Dim i As Long
    For i = 0 To groups_size - 1
        If groups_keys(i) Is instance Then
            groups_find = i
            Exit Function
        End If
    Next
    groups_find = -1
End Function

' monitors.at(instance): the instance's Dictionary
Private Function monitors_at(ByVal instance As ChakraScriptingInstance) As Object
    Dim i As Long
    For i = 0 To monitors_size - 1
        If monitors_keys(i) Is instance Then
            Set monitors_at = monitors_values(i)
            Exit Function
        End If
    Next
    Set monitors_at = Nothing
End Function

Private Sub timers_push_back(ByVal node As timerNode)
    If timers_size = 0 Then
        ReDim timers(0 To 15)
    ElseIf timers_size > UBound(timers) Then
        ReDim Preserve timers(0 To 2 * timers_size - 1)
    End If
    Set timers(timers_size) = node
    timers_size = timers_size + 1
End Sub

' timers.erase(iterator): node's entry
Private Sub timers_erase(ByVal node As timerNode)
    Dim i As Long, j As Long
    For i = 0 To timers_size - 1
        If timers(i) Is node Then
            For j = i To timers_size - 2
                Set timers(j) = timers(j + 1)
            Next
            timers_size = timers_size - 1
            Set timers(timers_size) = Nothing
            Exit Sub
        End If
    Next
End Sub

' nlohmann::json of a std::vector<std::string>
Private Function stringVectorToJson(v As StringVector) As nlohmann_json
    Dim j As nlohmann_json, i As Long
    Set j = json_array()
    For i = 0 To v.size - 1
        j.push_back json_value(v.items(i))
    Next
    Set stringVectorToJson = j
End Function

' std::right << std::setw(w) << s
Private Function setwRight(ByVal s As String, ByVal w As Long) As String
    If Len(s) < w Then s = Space$(w - Len(s)) & s
    setwRight = s
End Function

' lib/framework/wzpaths.cpp: WzPathInfo::fileName, baseName and path
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
    ' file.substr(0, std::max<size_t>(lastSlashPos, (size_t)1)): the 0-based position is lastSlashPos - 1
    If lastSlashPos - 1 > 1 Then WzPathInfo_path = Left$(file, lastSlashPos - 1) Else WzPathInfo_path = Left$(file, 1)
End Function

' WzString::normalized(WzString::NormalizationForm_KD)
Private Function WzString_normalized_KD(ByVal s As String) As String
    Dim n As Long, buf As String
    WzString_normalized_KD = s
    If Len(s) = 0 Then Exit Function
    n = NormalizeString(NormalizationKD, StrPtr(s), Len(s), 0, 0)
    If n <= 0 Then Exit Function
    buf = String$(n, vbNullChar)
    n = NormalizeString(NormalizationKD, StrPtr(s), Len(s), StrPtr(buf), n)
    If n <= 0 Then Exit Function
    WzString_normalized_KD = Left$(buf, n)
End Function

'--------------------------------------------------------------------------
' struct scripting_engine::timerNode

' timerNode(caller, func, timerName, plr, frame, additionalParam)
Public Function timerNode_new(ByVal caller As ChakraScriptingInstance, ByVal func As ITimerFunc, ByVal timerName As String, ByVal plr As Long, ByVal frame As Long, ByVal additionalParam As Object) As timerNode
    Dim node As timerNode
    Set node = New timerNode
    Set node.function_ = func
    node.timerName = timerName
    Set node.instance = caller
    node.baseobj = -1
    node.baseobjtype = OBJ_NUM_TYPES
    Set node.additionalTimerFuncParam = additionalParam
    node.frameTime = frame + gameTime
    node.ms = frame
    node.player = plr
    node.calls = 0
    node.typ = TIMER_REPEAT
    Set timerNode_new = node
End Function

'--------------------------------------------------------------------------
' struct scripting_engine::area_by_values_or_area_label_lookup

Public Function area_by_values_or_area_label_lookup_new() As area_by_values_or_area_label_lookup
    Dim r As area_by_values_or_area_label_lookup
    r.m_isLabel = False
    r.x1 = -1
    r.y1 = -1
    r.x2 = -1
    r.y2 = -1
    area_by_values_or_area_label_lookup_new = r
End Function

Public Function area_by_values_or_area_label_lookup_newLabel(ByVal label_ As String) As area_by_values_or_area_label_lookup
    Dim r As area_by_values_or_area_label_lookup
    r = area_by_values_or_area_label_lookup_new()
    r.m_isLabel = True
    r.m_label = label_
    area_by_values_or_area_label_lookup_newLabel = r
End Function

Public Function area_by_values_or_area_label_lookup_newArea(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As area_by_values_or_area_label_lookup
    Dim r As area_by_values_or_area_label_lookup
    r = area_by_values_or_area_label_lookup_new()
    r.m_isLabel = False
    r.x1 = x1
    r.y1 = y1
    r.x2 = x2
    r.y2 = y2
    area_by_values_or_area_label_lookup_newArea = r
End Function

Public Function area_by_values_or_area_label_lookup_isLabel(o As area_by_values_or_area_label_lookup) As Boolean
    area_by_values_or_area_label_lookup_isLabel = o.m_isLabel
End Function

' optional<std::string> label(): has_value, the label in result
Public Function area_by_values_or_area_label_lookup_label(o As area_by_values_or_area_label_lookup, result As String) As Boolean
    If o.m_isLabel Then
        result = o.m_label
        area_by_values_or_area_label_lookup_label = True
    End If
End Function

' optional<scr_area> area(): has_value, the area in result
Public Function area_by_values_or_area_label_lookup_area(o As area_by_values_or_area_label_lookup, result As scr_area) As Boolean
    If Not o.m_isLabel Then
        result.x1 = o.x1
        result.y1 = o.y1
        result.x2 = o.x2
        result.y2 = o.y2
        area_by_values_or_area_label_lookup_area = True
    End If
End Function

'--------------------------------------------------------------------------

Private Function scripting_engine_getNextAvailableTimerID() As Double
    Do
        lastTimerID = lastTimerID + 1
    Loop While timerIDMap().Exists(lastTimerID)
    scripting_engine_getNextAvailableTimerID = lastTimerID
End Function

' setTimer(caller, timerFunc, player, milliseconds, timerName = "", obj = nullptr, type = TIMER_REPEAT, additionalParam = nullptr)
Public Function scripting_engine_setTimer(ByVal caller As ChakraScriptingInstance, ByVal timerFunc As ITimerFunc, ByVal player As Long, ByVal milliseconds As Long, ByVal timerName As String, ByVal obj As Long, ByVal typ As timerType, ByVal additionalParam As Object) As Double
    Dim newTimerID As Double, node As timerNode
    newTimerID = scripting_engine_getNextAvailableTimerID()
    Set node = timerNode_new(caller, timerFunc, timerName, player, milliseconds, additionalParam)
    If obj <> 0 Then
        node.baseobj = SIMPLE_OBJECT_id(obj)
        node.baseobjtype = SIMPLE_OBJECT_type(obj)
    End If
    node.typ = typ
    node.timerID = newTimerID
    timers_push_back node
    timerIDMap().Add newTimerID, node
    scripting_engine_setTimer = newTimerID
End Function

' internal-only function that adds a Timer node (used for restoring saved games)
Private Sub scripting_engine_addTimerNode(ByVal node As timerNode)
    If Not (Not timerIDMap().Exists(node.timerID)) Then ASSERT_FAILURE "timerIDMap.count(node->timerID) == 0", "src/qtscript.cpp:243", "addTimerNode", "Duplicate timerID found: %s", Format$(node.timerID, "0")
    timers_push_back node
    Set timerIDMap().Item(node.timerID) = node
End Sub

' removes any timer(s) that satisfy _pred (a TIMERPRED_ number with its captured values); the ids of
' the removed timers in removedTimerIDs(0 .. count - 1), their count returned
Public Function scripting_engine_removeTimersIf(ByVal pred As Long, ByVal instance As ChakraScriptingInstance, ByVal timerName As String, ByVal player As Long, ByVal baseobj As Long, removedTimerIDs() As Double) As Long
    Dim i As Long, n As Long, cnt As Long, node As timerNode, match As Boolean
    For i = 0 To timers_size - 1
        Set node = timers(i)
        Select Case pred
        Case TIMERPRED_ONESHOT_DONE
            match = (node.typ = TIMER_ONESHOT_DONE)
        Case TIMERPRED_BASEOBJ
            match = (node.baseobj = baseobj)
        Case TIMERPRED_INSTANCE_NAME_PLAYER
            match = False
            If node.instance Is instance Then
                match = node.timerName = timerName And node.player = player
            End If
        End Select
        If match Then
            node.typ = TIMER_REMOVED ' in case a timer is removed while running timers
            ReDim Preserve removedTimerIDs(0 To cnt)
            removedTimerIDs(cnt) = node.timerID
            cnt = cnt + 1
        Else
            Set timers(n) = node
            n = n + 1
        End If
    Next
    For i = n To timers_size - 1
        Set timers(i) = Nothing
    Next
    timers_size = n
    For i = 0 To cnt - 1
        If timerIDMap().Exists(removedTimerIDs(i)) Then timerIDMap().Remove removedTimerIDs(i)
    Next
    scripting_engine_removeTimersIf = cnt
End Function

' Monitoring performance of function calls: executeWithPerformanceMonitoring(instance, function, f) is
' t = scripting_engine_executeWithPerformanceMonitoring_begin(), f(), then
' scripting_engine_executeWithPerformanceMonitoring_end instance, function, t
Public Function scripting_engine_executeWithPerformanceMonitoring_begin() As Currency
    Dim c As Currency
    QueryPerformanceCounter c
    scripting_engine_executeWithPerformanceMonitoring_begin = c
End Function

Public Sub scripting_engine_executeWithPerformanceMonitoring_end(ByVal instance As ChakraScriptingInstance, ByVal function_ As String, ByVal time_begin As Currency)
    Dim c As Currency, f As Currency, ticks As Long
    QueryPerformanceCounter c
    QueryPerformanceFrequency f
    ticks = Fix(CDbl(c - time_begin) * 1000000# / CDbl(f))   ' duration_cast<microseconds> truncates
    scripting_engine_logFunctionPerformance instance, function_, ticks
End Sub

' ----------------------------------------------------------

Public Sub scriptSetStartPos(ByVal position As Long, ByVal x As Long, ByVal y As Long)
    positions(position).x = x
    positions(position).y = y
    debug_ 292, LOG_SCRIPT, "scriptSetStartPos", "Setting start position %d to (%d, %d)", position, x, y
End Sub

Public Sub scriptSetDerrickPos(ByVal x As Long, ByVal y As Long)
    Dim mx As Long, my As Long, out As Long
    mx = map_coord(x)
    my = map_coord(y)
    ' MAX_TILE_TEXTURES is 255, so 2 bytes are enough
    ' to describe a map position
    out = (mx * 256 Or my) And &HFFFF&
    If Not derricks(out) Then
        derricks(out) = True
        derricks_size = derricks_size + 1
    End If
End Sub

Public Function scriptInit() As Boolean
    Dim i As Long

    For i = 0 To MAX_PLAYERS - 1
        scriptSetStartPos i, 0, 0
    Next
    Erase derricks
    derricks_size = 0
    scriptInit = True
End Function

Public Function getPlayerStartPosition(ByVal player As Long) As Vector2i
    getPlayerStartPosition = positions(player)
End Function

' (VB) for the tests: whether the set holds pos, and its size
Public Function derricks_contains(ByVal pos As Long) As Boolean
    derricks_contains = derricks(pos)
End Function

Public Function derricks_count() As Long
    derricks_count = derricks_size
End Function

Public Function scripting_engine_constructDerrickPositions() As nlohmann_json
    Dim derrickPositions As nlohmann_json, vector As nlohmann_json, pos As Long, mx As Long, my As Long
    ' Static map knowledge about start positions
    '== * ```derrickPositions``` A set of derrick starting positions on the current map. Each item in the set is an
    '== object containing the x and y variables for a derrick.
    Set derrickPositions = json_array() 'engine->newArray(derricks.size());
    For pos = 0 To 65535
        If derricks(pos) Then
            Set vector = json_object()
            mx = (pos And &HFF00&) \ 256
            my = pos And &HFF&
            vector.set_item "x", json_value(mx)
            vector.set_item "y", json_value(my)
            vector.set_item "type", json_value(CLng(SCRIPT_POSITION))
            derrickPositions.push_back vector
        End If
    Next
    Set scripting_engine_constructDerrickPositions = derrickPositions
End Function

Public Function scripting_engine_constructStartPositions() As nlohmann_json
    Dim startPositions As nlohmann_json, vector As nlohmann_json, i As Long
    ' Static map knowledge about start positions
    '== * ```startPositions``` An array of player start positions on the current map. Each item in the array is an
    '== object containing the x and y variables for a player start position.
    Set startPositions = json_array() 'engine->newArray(game.maxPlayers);
    For i = 0 To game.maxPlayers - 1
        Set vector = json_object()
        vector.set_item "x", json_value(map_coord(positions(i).x))
        vector.set_item "y", json_value(map_coord(positions(i).y))
        vector.set_item "type", json_value(CLng(SCRIPT_POSITION))
        startPositions.push_back vector
    Next
    Set scripting_engine_constructStartPositions = startPositions
End Function

' ----------------------------------------------------------

Public Function scripting_engine_removeTimer(ByVal timerID As Double) As Boolean
    Dim node As timerNode
    If timerIDMap().Exists(timerID) Then
        Set node = timerIDMap().Item(timerID)
        node.typ = TIMER_REMOVED ' in case a timer is removed while running timers
        timers_erase node
        timerIDMap().Remove timerID
        scripting_engine_removeTimer = True
        Exit Function
    End If
    scripting_engine_removeTimer = False
End Function

Public Sub scriptRemoveObject(ByVal psObj As Long)
    Dim removed() As Double
    ' Weed out timers with dead objects
    scripting_engine_removeTimersIf TIMERPRED_BASEOBJ, Nothing, vbNullString, 0, SIMPLE_OBJECT_id(psObj), removed
    scripting_engine_groupRemoveObject psObj
End Sub

' do not want to call this 'init', since scripts are often loaded before we get here
Public Function prepareScripts(ByVal loadGame As Boolean) As Boolean
    Dim resEvent As researchEvent, i As Long
    debug_ 388, LOG_WZ, "prepareScripts", "Scripts prepared"
    If Not loadGame Then ' labels saved in savegame
        scripting_engine_prepareLabels
    End If
    ' Assume that by this point all scripts are loaded
    scriptsReady = True
    Do While eventQueue_size > 0
        resEvent = eventQueue(0)
        triggerEventResearched resEvent.research, resEvent.structure, resEvent.player
        ' eventQueue.pop()
        For i = 0 To eventQueue_size - 2
            eventQueue(i) = eventQueue(i + 1)
        Next
        eventQueue_size = eventQueue_size - 1
    Loop
    prepareScripts = True
End Function

Public Function initScripts() As Boolean
    initScripts = True
End Function

Public Function shutdownScripts() As Boolean
    shutdownScripts = scripting_engine_shutdownScripts()
End Function

Public Function scripting_engine_shutdownScripts() As Boolean
    Dim i As Long, instance As ChakraScriptingInstance, monitor As Object, scriptName As String
    Dim k As Variant, function_ As String, m As MONITOR_BIN, info As String
    scriptsReady = False
    jsDebugShutdown
    globalDialog = False
    For i = 0 To scripts_size - 1
        Set instance = scripts(i)
        Set monitor = monitors_at(instance)
        scriptName = instance.scriptName()
        instance.dumpScriptLog "=== PERFORMANCE DATA ===" & vbLf
        instance.dumpScriptLog "    calls | avg (usec) | worst (usec) | worst call at | >=limit | >=limit/2 | function" & vbLf
        For Each k In monitor.Keys
            function_ = k
            m = monitorBins(monitor.Item(k))
            info = setwRight(CStr(m.calls), 9) & " | "
            info = info & setwRight(Format$(Fix(m.time_ / m.calls), "0"), 10) & " | "
            info = info & setwRight(CStr(m.worst), 12) & " | "
            info = info & setwRight(Format$(U32(m.worstGameTime), "0"), 13) & " | "
            info = info & setwRight(CStr(m.overMaxTimeCalls), 7) & " | "
            info = info & setwRight(CStr(m.overHalfMaxTimeCalls), 9) & " | "
            info = info & function_ & vbLf
            instance.dumpScriptLog info
        Next
        monitor.RemoveAll
        scripting_engine_unregisterFunctions instance
    Next
    Erase timers
    timers_size = 0
    lastTimerID = 0
    timerIDMap().RemoveAll
    Erase monitors_keys
    Erase monitors_values
    monitors_size = 0
    Erase monitorBins
    monitorBins_size = 0
    Erase scripts
    scripts_size = 0
    scripting_engine_shutdownScripts = True
End Function

Public Function updateScripts() As Boolean
    updateScripts = scripting_engine_updateScripts()
End Function

Public Function scripting_engine_updateScripts() As Boolean
    Dim i As Long, sel() As Long, nSel As Long, removed() As Double
    Dim runlist() As timerNode, runlist_size As Long, node As timerNode
    ' Call delayed triggers here
    If selectionChanged Then
        For i = 0 To scripts_size - 1
            nSel = wzapi_enumSelected(sel)
            scripts(i).handle_eventSelectionChanged sel, nSel
        Next
        selectionChanged = False
    End If

    ' Update gameTime
    For i = 0 To scripts_size - 1
        scripts(i).updateGameTime gameTime
    Next
    ' Weed out dead timers
    scripting_engine_removeTimersIf TIMERPRED_ONESHOT_DONE, Nothing, vbNullString, 0, 0, removed
    ' Check for timers, and run them if applicable.
    ' TODO - load balancing
    ' make a new list here, since we might trample all over the timer list during execution
    If timers_size > 0 Then ReDim runlist(0 To timers_size - 1)
    For i = 0 To timers_size - 1
        Set node = timers(i)
        If Not ULT(gameTime, node.frameTime) Then ' frameTime <= gameTime (unsigned)
            node.frameTime = node.ms + gameTime ' update for next invokation
            If node.typ = TIMER_ONESHOT_READY Then
                node.typ = TIMER_ONESHOT_DONE ' unless there is none
            End If
            node.calls = node.calls + 1
            Set runlist(runlist_size) = node
            runlist_size = runlist_size + 1
        End If
    Next

    For i = 0 To runlist_size - 1
        Set node = runlist(i)
        ' IMPORTANT: A queued function can delete a timer that is in the runlist!
        ' So we must verify that the node is not one of the deleted ones.
        If node.typ = TIMER_REMOVED Then
            GoTo continue_ ' skip
        End If
        node.function_.invoke node.timerID, IdToObject(node.baseobjtype, node.baseobj, node.player), node.additionalTimerFuncParam
continue_:
    Next

    scripting_engine_updateScripts = True
End Function

Public Function loadPlayerScript(ByVal path As String, ByVal player As Long, ByVal difficulty As Long) As ChakraScriptingInstance
    Set loadPlayerScript = scripting_engine_loadPlayerScript(path, player, difficulty)
End Function

Private Function loadPlayerScriptByBackend(ByVal path As String, ByVal player As Long, ByVal realDifficulty As Long) As ChakraScriptingInstance
    ' JS scripts:
    Select Case war_getJSBackend()
    Case JS_BACKEND_quickjs
        ' (VB) the Chakra backend stands for QuickJS
        Set loadPlayerScriptByBackend = createChakraScriptInstance(path, player, realDifficulty)
        Exit Function
    Case JS_BACKEND_num_backends
        debug_ 526, LOG_ERROR, "loadPlayerScriptByBackend", "Invalid js backend value" ' should not happen
    End Select
    Set loadPlayerScriptByBackend = Nothing
End Function

Public Function scripting_engine_loadPlayerScript(ByVal path As String, ByVal player As Long, ByVal difficulty As Long) As ChakraScriptingInstance
    Dim realDifficulty As Long, pNewInstance As ChakraScriptingInstance, psMap As GROUPMAP, inserted As Boolean
    Dim globalVars As nlohmann_json, tilesetType As String, globalVarsToSave As nlohmann_json

    If Not (player >= 0 And (player < MAX_PLAYERS Or player = selectedPlayer)) Then ASSERT_FAILURE "player >= 0 && (player < MAX_PLAYERS || player == selectedPlayer)", "src/qtscript.cpp:533", "loadPlayerScript", "Player index %d out of bounds", player: Set scripting_engine_loadPlayerScript = Nothing: Exit Function

    debug_ 535, LOG_SCRIPT, "loadPlayerScript", "loadPlayerScript[%d]: %s", player, path

    If game.typ = LEVEL_TYPE_SKIRMISH Then
        realDifficulty = difficulty     ' static_cast<int8_t>(difficulty)
    Else ' campaign
        realDifficulty = getDifficultyLevel()
    End If

    Set pNewInstance = loadPlayerScriptByBackend(path, player, realDifficulty)
    If pNewInstance Is Nothing Then
        ' failed to create new scripting instance
        debug_ 551, LOG_ERROR, "loadPlayerScript", "Failed to create new scripting instance: path:""%s"", player: %d, difficulty: %d", path, player, difficulty
        Set scripting_engine_loadPlayerScript = Nothing
        Exit Function
    End If

    ' Create group map
    Set psMap = New GROUPMAP
    inserted = groups_find(pNewInstance) < 0
    If inserted Then
        If groups_size = 0 Then
            ReDim groups_keys(0 To 7)
            ReDim groups_values(0 To 7)
        ElseIf groups_size > UBound(groups_keys) Then
            ReDim Preserve groups_keys(0 To 2 * groups_size - 1)
            ReDim Preserve groups_values(0 To 2 * groups_size - 1)
        End If
        Set groups_keys(groups_size) = pNewInstance
        Set groups_values(groups_size) = psMap
        groups_size = groups_size + 1
    End If
    If Not (inserted) Then ASSERT_FAILURE "insert_result.second", "src/qtscript.cpp:558", "loadPlayerScript", "Entry for this engine %p already exists in ENGINEMAP!", ObjPtr(pNewInstance)

    Set globalVars = json_object()
    ' Special global variables
    '== * ```version``` Current version of the game, set in *major.minor* format.
    globalVars.set_item "version", json_value(version_getVersionString())
    '== * ```selectedPlayer``` The player controlled by the client on which the script runs.
    globalVars.set_item "selectedPlayer", json_value_unsigned(selectedPlayer)
    '== * ```gameTime``` The current game time. Updated before every invokation of a script.
    pNewInstance.setSpecifiedGlobalVariable "gameTime", json_value_unsigned(gameTime), GlobalVariableFlags_ReadOnlyUpdatedFromApp Or GlobalVariableFlags_DoNotSave
    '== * ```modList``` An array of the current loaded mods (mod names).
    globalVars.set_item "modList", stringVectorToJson(getModNamesList())


    '== * ```difficulty``` The currently set campaign difficulty, or the current AI's difficulty setting. It will be one of
    '== ```SUPEREASY``` (campaign only), ```EASY```, ```MEDIUM```, ```HARD``` or ```INSANE```.
    If game.typ = LEVEL_TYPE_SKIRMISH Then
        globalVars.set_item "difficulty", json_value(CLng(difficulty))
    Else ' campaign
        globalVars.set_item "difficulty", json_value(CLng(getDifficultyLevel()))
    End If
    '== * ```mapName``` The name of the current map.
    globalVars.set_item "mapName", json_value(game.map)
    '== * ```tilesetType``` The area name of the map.
    tilesetType = "CUSTOM"
    Select Case currentMapTileset
    Case MAP_TILESET_ARIZONA
        tilesetType = "ARIZONA"
    Case MAP_TILESET_URBAN
        tilesetType = "URBAN"
    Case MAP_TILESET_ROCKIES
        tilesetType = "ROCKIES"
    End Select
    globalVars.set_item "tilesetType", json_value(tilesetType)
    '== * ```baseType``` The type of base that the game starts with. It will be one of ```CAMP_CLEAN```, ```CAMP_BASE``` or ```CAMP_WALLS```.
    globalVars.set_item "baseType", json_value_unsigned(game.base)
    '== * ```alliancesType``` The type of alliances permitted in this game. It will be one of ```NO_ALLIANCES```, ```ALLIANCES```, ```ALLIANCES_UNSHARED``` or ```ALLIANCES_TEAMS```.
    globalVars.set_item "alliancesType", json_value_unsigned(game.alliance)
    '== * ```powerType``` The power level set for this game.
    globalVars.set_item "powerType", json_value_unsigned(game.power)
    '== * ```maxPlayers``` The number of active players in this game.
    globalVars.set_item "maxPlayers", json_value_unsigned(game.maxPlayers)
    '== * ```scavengers``` Whether or not scavengers are activated in this game, and, if so, which type.
    globalVars.set_item "scavengers", json_value_unsigned(game.scavengers)
    '== * ```mapWidth``` Width of map in tiles.
    globalVars.set_item "mapWidth", json_value(mapWidth)
    '== * ```mapHeight``` Height of map in tiles.
    globalVars.set_item "mapHeight", json_value(mapHeight)
    '== * ```scavengerPlayer``` Index of scavenger player. (3.2+ only)
    globalVars.set_item "scavengerPlayer", json_value(scavengerSlot())
    '== * ```isMultiplayer``` If the current game is a online multiplayer game or not. (3.2+ only)
    globalVars.set_item "isMultiplayer", json_value(NetPlay_bComms)
    '== * ```challenge``` If the current game is a challenge. (4.1.4+ only)
    globalVars.set_item "challenge", json_value(challengeActive)
    '== * ```idleTime``` The amount of game time without active play before a player should be considered "inactive". (0 = disable activity alerts / AFK check) (4.2.0+ only)
    globalVars.set_item "idleTime", json_value_unsigned(game.inactivityMinutes * 60 * 1000)
    '== * ```gameTimeLimit``` The game time limit (match will automatically end if it reaches this duration). (0 = disable limit) (4.4.0+ only)
    globalVars.set_item "gameTimeLimit", json_value_unsigned(game.gameTimeLimitMinutes * 60 * 1000)
    '== * ```playerLeaveMode``` The mode used to handle human players leaving a multiplayer game in progress. (0 = destroy resources, 1 = split resources with team) (4.4.0+ only)
    globalVars.set_item "playerLeaveMode", json_value_unsigned(game.playerLeaveMode And &HFF&)
    '== * ```tweakOptions``` The tweakOptions offered by the current mod / mode, as configured by the user. (4.5.0+ only)
    globalVars.set_item "tweakOptions", getCamTweakOptions()

    pNewInstance.setSpecifiedGlobalVariables globalVars, GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave

    ' Register 'Stats' object. It is a read-only representation of basic game component states.
    pNewInstance.setSpecifiedGlobalVariable "Stats", wzapi_constructStatsObject(), GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave

    ' Register 'MapTiles' two-dimensional array. It is a read-only representation of static map tile states.
    pNewInstance.setSpecifiedGlobalVariable "MapTiles", wzapi_constructMapTilesArray(), GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave

    ' Set some useful constants
    pNewInstance.setSpecifiedGlobalVariables wzapi_getUsefulConstants(), GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave

    '/ Place to store group sizes
    '== * ```groupSizes``` A sparse array of group sizes. If a group has never been used, the entry in this array will
    '== be undefined.
    pNewInstance.setSpecifiedGlobalVariable "groupSizes", json_object(), GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave

    ' Static knowledge about players
    pNewInstance.setSpecifiedGlobalVariable "playerData", wzapi_constructStaticPlayerData(), GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave

    ' Static map knowledge about start positions
    pNewInstance.setSpecifiedGlobalVariable "derrickPositions", scripting_engine_constructDerrickPositions(), GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave
    pNewInstance.setSpecifiedGlobalVariable "startPositions", scripting_engine_constructStartPositions(), GlobalVariableFlags_ReadOnly Or GlobalVariableFlags_DoNotSave

    ' WzPathInfo basename = WzPathInfo::fromPlatformIndependentPath(path.toUtf8());
    Set globalVarsToSave = json_object()
    ' We need to always save the 'me' special variable.
    '== * ```me``` The player the script is currently running as.
    globalVarsToSave.set_item "me", json_value(player)

    ' We also need to save the special 'scriptName' variable.
    '== * ```scriptName``` Base name of the script that is running.
    globalVarsToSave.set_item "scriptName", json_value(WzPathInfo_baseName(path))

    ' We also need to save the special 'scriptPath' variable.
    '== * ```scriptPath``` Base path of the script that is running.
    globalVarsToSave.set_item "scriptPath", json_value(WzPathInfo_path(path))

    pNewInstance.setSpecifiedGlobalVariables globalVarsToSave, GlobalVariableFlags_ReadOnly ' ensure these are saved

    ' Clear previous log file
    PHYSFS_delete "logs/" & pNewInstance.scriptName() & ".log"

    ' Attempt to ready instance for execution
    If Not pNewInstance.readyInstanceForExecution() Then
        Set pNewInstance = Nothing
        debug_ 674, LOG_ERROR, "loadPlayerScript", "Unable to ready instance for execution: %s", path
        Set scripting_engine_loadPlayerScript = Nothing
        Exit Function
    End If

    ' Register script
    If scripts_size = 0 Then
        ReDim scripts(0 To 7)
    ElseIf scripts_size > UBound(scripts) Then
        ReDim Preserve scripts(0 To 2 * scripts_size - 1)
    End If
    Set scripts(scripts_size) = pNewInstance
    scripts_size = scripts_size + 1

    ' MONITOR *monitor = new MONITOR; monitors[pNewInstance] = monitor;
    If monitors_size = 0 Then
        ReDim monitors_keys(0 To 7)
        ReDim monitors_values(0 To 7)
    ElseIf monitors_size > UBound(monitors_keys) Then
        ReDim Preserve monitors_keys(0 To 2 * monitors_size - 1)
        ReDim Preserve monitors_values(0 To 2 * monitors_size - 1)
    End If
    Set monitors_keys(monitors_size) = pNewInstance
    Set monitors_values(monitors_size) = CreateObject("Scripting.Dictionary")
    monitors_size = monitors_size + 1

    debug_ 684, LOG_SAVE, "loadPlayerScript", "Created script engine %zu for player %d from %s", scripts_size - 1, player, path
    Set scripting_engine_loadPlayerScript = pNewInstance
End Function

Public Function loadGlobalScript(ByVal path As String) As Boolean
    loadGlobalScript = Not (loadPlayerScript(path, selectedPlayer, AIDifficulty_DISABLED) Is Nothing)
End Function

Public Function saveScriptStates(ByVal filename As String) As Boolean
    saveScriptStates = scripting_engine_saveScriptStates(filename)
End Function

Public Function scripting_engine_saveScriptStates(ByVal filename As String) As Boolean
    Dim ini As WzConfig, i As Long, instance As ChakraScriptingInstance, globalsResult As nlohmann_json
    Dim groupsResult As nlohmann_json, timerIdx As Long, node As timerNode, nodeInfo As nlohmann_json
    Set ini = WzConfig_new(filename, WzConfig_ReadAndWrite)
    For i = 0 To scripts_size - 1
        Set instance = scripts(i)

        Set globalsResult = json_object()
        instance.saveScriptGlobals globalsResult
        ' 'scriptName' and 'me' should be saved implicitly by the backend's saveScriptGlobals
        If Not (globalsResult.contains("me")) Then ASSERT_FAILURE "globalsResult.contains(""me"")", "src/qtscript.cpp:708", "saveScriptStates", "Missing required global ""me"""
        If Not (globalsResult.contains("scriptName")) Then ASSERT_FAILURE "globalsResult.contains(""scriptName"")", "src/qtscript.cpp:709", "saveScriptStates", "Missing required global ""scriptName"""
        ini.setValue "globals_" & CStr(i), globalsResult

        ' we have to save 'scriptName' and 'me' explicitly
        Set groupsResult = json_object()
        scripting_engine_saveGroups groupsResult, instance
        groupsResult.set_item "me", json_value(instance.player())
        groupsResult.set_item "scriptName", json_value(instance.scriptName())
        ini.setValue "groups_" & CStr(i), groupsResult
    Next
    timerIdx = 0
    For i = 0 To timers_size - 1
        Set node = timers(i)
        Set nodeInfo = json_object()
        nodeInfo.set_item "timerID", json_value_unsigned(node.timerID)
        nodeInfo.set_item "timerName", json_value(node.timerName)
        ' we have to save 'scriptName' and 'me' explicitly
        nodeInfo.set_item "me", json_value(node.player)
        nodeInfo.set_item "scriptName", json_value(node.instance.scriptName())
        nodeInfo.set_item "functionRestoreInfo", node.instance.saveTimerFunction(node.timerID, node.timerName, node.additionalTimerFuncParam)
        If node.baseobj >= 0 Then
            nodeInfo.set_item "object", json_value(node.baseobj)
            nodeInfo.set_item "objectType", json_value(CLng(node.baseobjtype))
        End If
        nodeInfo.set_item "frame", json_value(node.frameTime)
        nodeInfo.set_item "ms", json_value(node.ms)
        nodeInfo.set_item "calls", json_value(node.calls)
        nodeInfo.set_item "type", json_value(CLng(node.typ))

        ini.setValue "triggers_" & CStr(timerIdx), nodeInfo
        timerIdx = timerIdx + 1
    Next
    scripting_engine_saveScriptStates = True
End Function

Private Function scripting_engine_findInstanceForPlayer(ByVal match As Long, ByVal scriptName_ As String) As ChakraScriptingInstance
    Dim scriptName As String, i As Long, player As Long, matchName As String
    scriptName = WzString_normalized_KD(scriptName_)
    For i = 0 To scripts_size - 1
        player = scripts(i).player()
        matchName = WzString_normalized_KD(scripts(i).scriptName())
        If match = player And (matchName = scriptName Or Len(scriptName) = 0) Then
            Set scripting_engine_findInstanceForPlayer = scripts(i)
            Exit Function
        End If
    Next
    ASSERT_FAILURE "false", "src/qtscript.cpp:757", "findInstanceForPlayer", "Script context for player %d and script name %s not found", match, scriptName
    Set scripting_engine_findInstanceForPlayer = Nothing
End Function

Public Function loadScriptStates(ByVal filename As String) As Boolean
    loadScriptStates = scripting_engine_loadScriptStates(filename)
End Function

' (VB) loadScriptStates' try block around restoreTimerFunction: False where the C++ continues with the next
' group (the catch, or a trigger without its function; ini.endGroup() is skipped then, as in C++)
Private Function restoreTimerFunctionInfo(ByVal ini As WzConfig, ByVal instance As ChakraScriptingInstance, ByVal groupName As String, timerFunc As ITimerFunc, additionalParam As Object) As Boolean
    Dim backwardsCompatJSFunctionRestoreInfo As nlohmann_json, errSource As String, errNumber As Long, errDescription As String
    On Error GoTo catch_
    If ini.contains("functionRestoreInfo") Then
        Set timerFunc = instance.restoreTimerFunction(ini.value("functionRestoreInfo").jsonValue(), additionalParam)
    Else
        ' backwards-compat with old saves
        ' construct a JS-compatible functionRestoreInfo using the "function" key, which should be set in an older save
        Set backwardsCompatJSFunctionRestoreInfo = json_object()
        If Not ini.contains("function") Then
            ASSERT_FAILURE "false", "src/qtscript.cpp:825", "loadScriptStates", "Invalid trigger in save (%s) - missing new functionRestoreInfo block, and old function parameter", groupName
            Exit Function
        End If
        backwardsCompatJSFunctionRestoreInfo.set_item "function", ini.value("function").jsonValue()
        Set timerFunc = instance.restoreTimerFunction(backwardsCompatJSFunctionRestoreInfo, additionalParam)
    End If
    restoreTimerFunctionInfo = True
    Exit Function
catch_:
    ' catch (const std::exception& e): the JSON errors; anything else goes on up
    errSource = Err.Source
    errNumber = Err.Number
    errDescription = Err.Description
    If errSource <> "nlohmann::json" Then
        On Error GoTo 0
        Err.Raise errNumber, errSource, errDescription
    End If
    ASSERT_FAILURE "false", "src/qtscript.cpp:834", "loadScriptStates", "Failed to restore saved timer function info: %s", errDescription
End Function

Public Function scripting_engine_loadScriptStates(ByVal filename As String) As Boolean
    Dim maxRestoredTimerID As Double, ini As WzConfig, list() As String, n As Long, i As Long, player As Long
    Dim scriptName As String, instance As ChakraScriptingInstance, node As timerNode, timerFunc As ITimerFunc, additional As Object
    Dim result As nlohmann_json, keys() As String, nKeys As Long, j As Long, values() As String, nValues As Long
    Dim ok As Boolean, droidId As Long, errText As String, k As Long, groupId As Long
    Dim bHasLastNewGroupId As Variant, lastNewGroupId As Long, psMap As GROUPMAP
    maxRestoredTimerID = 0
    Set ini = WzConfig_new(filename, WzConfig_ReadOnly)
    n = ini.childGroups(list)
    debug_ 772, LOG_SAVE, "loadScriptStates", "Loading script states for %zu script contexts", scripts_size
    For i = 0 To n - 1
        ini.beginGroup list(i)
        player = ini.value("me").toInt()
        scriptName = ini.value("scriptName").toWzString()
        Set instance = scripting_engine_findInstanceForPlayer(player, scriptName)
        If Not instance Is Nothing And Left$(list(i), 9) = "triggers_" Then
            Set node = New timerNode
            If ini.contains("timerID") Then
                node.timerID = ini.value("timerID").toInt()
            Else
                ' backwards-compat with old saves
                node.timerID = scripting_engine_getNextAvailableTimerID()
            End If
            If ini.contains("timerName") Then
                node.timerName = ini.value("timerName").toWzString()
            Else
                ' backwards-compat with old saves
                node.timerName = ini.value("function").toWzString()
            End If
            Set node.instance = instance
            debug_ 801, LOG_SAVE, "loadScriptStates", "Registering trigger %zu for player %d, script %s", i, player, scriptName
            node.baseobj = ini.value("baseobj", json_variant_new(-1)).toInt()
            node.baseobjtype = ini.value("objectType", json_variant_new(CLng(OBJ_NUM_TYPES))).toInt()
            node.frameTime = ini.value("frame").toInt()
            node.ms = ini.value("ms").toInt()
            node.player = player
            node.calls = ini.value("calls").toInt()
            node.typ = ini.value("type", json_variant_new(CLng(TIMER_REPEAT))).toInt()

            Set timerFunc = Nothing
            Set additional = Nothing
            If Not restoreTimerFunctionInfo(ini, instance, list(i), timerFunc, additional) Then
                GoTo continue_
            End If

            Set node.function_ = timerFunc
            Set node.additionalTimerFuncParam = additional

            If node.timerID > maxRestoredTimerID Then maxRestoredTimerID = node.timerID

            scripting_engine_addTimerNode node
        ElseIf Not instance Is Nothing And Left$(list(i), 8) = "globals_" Then
            Set result = ini.currentJsonValue().clone()
            debug_ 848, LOG_SAVE, "loadScriptStates", "Loading script globals for player %d, script %s -- found %zu values", instance.player(), instance.scriptName(), result.size()
            ' filter out "scriptName" and "me" variables
            result.erase_ "me"
            result.erase_ "scriptName"
            instance.loadScriptGlobals result
        ElseIf Not instance Is Nothing And Left$(list(i), 7) = "groups_" Then
            nKeys = ini.childKeys(keys)
            For j = 0 To nKeys - 1
                nValues = ini.value(keys(j)).toWzStringList(values)
                ok = std_stoi(keys(j), droidId, errText) ' check if number
                If Not ok Then droidId = 0
                For k = 0 To nValues - 1
                    If Not ok Then Exit For
                    If Not std_stoi(values(k), groupId, errText) Then groupId = 0
                    scripting_engine_loadGroup instance, groupId, droidId
                Next
            Next
            bHasLastNewGroupId = False
            lastNewGroupId = ini.value("lastNewGroupId").toInt(bHasLastNewGroupId)
            If bHasLastNewGroupId Then
                Set psMap = scripting_engine_getGroupMap(instance)
                If Not psMap Is Nothing Then
                    psMap.saveLoadSetLastNewGroupId lastNewGroupId
                End If
            End If
        Else
            If Not instance Is Nothing Then
                debug_ 884, LOG_WARNING, "loadScriptStates", "Encountered unexpected group '%s' in loadScriptStates", list(i)
            End If
        End If
        ini.endGroup
continue_:
    Next
    lastTimerID = maxRestoredTimerID
    scripting_engine_loadScriptStates = True
End Function

' std::unordered_map<wzapi::scripting_instance *, nlohmann::json> (MODELMAP): the count, the instances and their
' globals in instances(0 .. count - 1) and globals_(0 .. count - 1)
Public Function scripting_engine_debug_GetGlobalsSnapshot(instances() As ChakraScriptingInstance, globals_() As nlohmann_json) As Long
    Dim i As Long
    If scripts_size > 0 Then
        ReDim instances(0 To scripts_size - 1)
        ReDim globals_(0 To scripts_size - 1)
    End If
    For i = 0 To scripts_size - 1
        Set instances(i) = scripts(i)
        Set globals_(i) = scripts(i).debugGetAllScriptGlobals()
    Next
    scripting_engine_debug_GetGlobalsSnapshot = scripts_size
End Function

' timerNodeSnapshot(const std::shared_ptr<timerNode>& node)
Private Function timerNodeSnapshot_new(ByVal node As timerNode) As timerNodeSnapshot
    Dim s As timerNodeSnapshot
    s.timerID = 0
    s.baseobj = -1
    s.baseobjtype = OBJ_NUM_TYPES
    s.frameTime = -1
    s.ms = -1
    s.player = -1
    s.calls = 0
    s.typ = TIMER_REMOVED
    If node Is Nothing Then ASSERT_FAILURE "node != nullptr", "src/qtscript.h:461", "timerNodeSnapshot", "null timerNode": timerNodeSnapshot_new = s: Exit Function
    s.timerID = node.timerID
    s.timerName = node.timerName
    Set s.instance = node.instance
    s.baseobj = node.baseobj
    s.baseobjtype = node.baseobjtype
    s.frameTime = node.frameTime
    s.ms = node.ms
    s.player = node.player
    s.calls = node.calls
    s.typ = node.typ
    Set s.instanceTimerRestoreData = node.instance.saveTimerFunction(node.timerID, node.timerName, node.additionalTimerFuncParam)
    timerNodeSnapshot_new = s
End Function

' std::vector<timerNodeSnapshot>: the count, the snapshots in result(0 .. count - 1)
Public Function scripting_engine_debug_GetTimersSnapshot(result() As timerNodeSnapshot) As Long
    Dim i As Long
    If timers_size > 0 Then ReDim result(0 To timers_size - 1)
    For i = 0 To timers_size - 1
        result(i) = timerNodeSnapshot_new(timers(i))
    Next
    scripting_engine_debug_GetTimersSnapshot = timers_size
End Function

Public Sub jsAutogameSpecific(ByVal name_ As String, ByVal player As Long, ByVal difficulty As Long)
    Dim instance As ChakraScriptingInstance
    Set instance = loadPlayerScript(name_, player, difficulty)
    If instance Is Nothing Then
        console "Failed to load selected AI! Check your logs to see why."
        Exit Sub
    End If
    console "Loaded the %s AI script for current player!", name_
    instance.handle_eventGameInit
    instance.handle_eventStartLevel
End Sub

Public Sub jsHandleDebugClosed()
    globalDialog = False
End Sub

' (VB) the dialog's DebugInterface waits for cluster G (Nothing); its shutdown handler is a number
Public Sub jsShowDebug()
    Dim isSpectator As Boolean
    globalDialog = True
    isSpectator = NetPlay_players_isSpectator(realSelectedPlayer)
    jsDebugCreate Nothing, JSDEBUG_SHUTDOWN_jsHandleDebugClosed, isSpectator
End Sub

' ----------------------------------------------------------------------------------------
' Events

'/ For generic, parameter-less triggers
Public Function triggerEvent(ByVal trigger As SCRIPT_TRIGGER_TYPE, Optional ByVal psObj As Long = 0) As Boolean
    Dim i As Long, instance As ChakraScriptingInstance, player As Long, receiveAll As Boolean
    ' HACK: TRIGGER_VIDEO_QUIT is called before scripts for initial campaign video
    If Not (scriptsReady Or trigger = TRIGGER_VIDEO_QUIT) Then ASSERT_FAILURE "scriptsReady || trigger == TRIGGER_VIDEO_QUIT", "src/qtscript.cpp:1044", "triggerEvent", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        Set instance = scripts(i)
        If psObj <> 0 Then
            player = instance.player()
            receiveAll = instance.isReceivingAllEvents()
            If player <> SIMPLE_OBJECT_player(psObj) And Not receiveAll Then
                GoTo continue_
            End If
        End If

        Select Case trigger
        Case TRIGGER_GAME_INIT
            instance.handle_eventGameInit
        Case TRIGGER_START_LEVEL
            processVisibility ' make sure we initialize visibility first
            instance.handle_eventStartLevel
        Case TRIGGER_TRANSPORTER_LAUNCH
            instance.handle_eventLaunchTransporter ' deprecated!
            instance.handle_eventTransporterLaunch psObj
        Case TRIGGER_TRANSPORTER_ARRIVED
            instance.handle_eventReinforcementsArrived ' deprecated!
            instance.handle_eventTransporterArrived psObj
        Case TRIGGER_DELIVERY_POINT_MOVING
            instance.handle_eventDeliveryPointMoving psObj
        Case TRIGGER_DELIVERY_POINT_MOVED
            instance.handle_eventDeliveryPointMoved psObj
        Case TRIGGER_OBJECT_RECYCLED
            instance.handle_eventObjectRecycled psObj
        Case TRIGGER_TRANSPORTER_EXIT
            instance.handle_eventTransporterExit psObj
        Case TRIGGER_TRANSPORTER_DONE
            instance.handle_eventTransporterDone psObj
        Case TRIGGER_TRANSPORTER_LANDED
            instance.handle_eventTransporterLanded psObj
        Case TRIGGER_TRANSPORTER_EMBARKED
            instance.handle_eventTransporterEmbarked psObj
        Case TRIGGER_TRANSPORTER_DISEMBARKED
            instance.handle_eventTransporterDisembarked psObj
        Case TRIGGER_MISSION_TIMEOUT
            instance.handle_eventMissionTimeout
        Case TRIGGER_VIDEO_QUIT
            instance.handle_eventVideoDone
        Case TRIGGER_GAME_LOADED
            instance.handle_eventGameLoaded
        Case TRIGGER_GAME_SAVING
            instance.handle_eventGameSaving
        Case TRIGGER_GAME_SAVED
            instance.handle_eventGameSaved
        Case TRIGGER_DESIGN_BODY
            instance.handle_eventDesignBody
        Case TRIGGER_DESIGN_PROPULSION
            instance.handle_eventDesignPropulsion
        Case TRIGGER_DESIGN_WEAPON
            instance.handle_eventDesignWeapon
        Case TRIGGER_DESIGN_COMMAND
            instance.handle_eventDesignCommand
        Case TRIGGER_DESIGN_SYSTEM
            instance.handle_eventDesignSystem
        Case TRIGGER_DESIGN_QUIT
            instance.handle_eventDesignQuit
        Case TRIGGER_MENU_BUILD_SELECTED
            instance.handle_eventMenuBuildSelected
        Case TRIGGER_MENU_RESEARCH_SELECTED
            instance.handle_eventMenuResearchSelected
        Case TRIGGER_MENU_BUILD_UP
            instance.handle_eventMenuBuild
        Case TRIGGER_MENU_RESEARCH_UP
            instance.handle_eventMenuResearch
        Case TRIGGER_MENU_DESIGN_UP
            instance.handle_eventMenuDesign
        Case TRIGGER_MENU_MANUFACTURE_UP
            instance.handle_eventMenuManufacture
        End Select
continue_:
    Next

    If (trigger = TRIGGER_START_LEVEL Or trigger = TRIGGER_GAME_LOADED) And Len(saveandquit_enabled()) <> 0 Then
        saveGame saveandquit_enabled(), GTYPE_SAVE_START
        wzQuit 0
    End If

    triggerEvent = True
End Function

Public Function triggerEventPlayerLeft(ByVal player As Long) As Boolean
    Dim i As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1167", "triggerEventPlayerLeft", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventPlayerLeft player
    Next
    triggerEventPlayerLeft = True
End Function

Public Function triggerEventCheatMode(ByVal entered As Boolean) As Boolean
    Dim i As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1182", "triggerEventCheatMode", "Scripts not initialized yet"
    GameStoryLogger_logDebugModeChange entered
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventCheatMode entered
    Next
    triggerEventCheatMode = True
End Function

Public Function triggerEventDroidIdle(ByVal psDroid As Long) As Boolean
    Dim i As Long, player As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1197", "triggerEventDroidIdle", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        player = scripts(i).player()
        If player = aDroid(psDroid).player Then
            scripts(i).handle_eventDroidIdle psDroid
        End If
    Next
    triggerEventDroidIdle = True
End Function

Public Function triggerEventDroidBuilt(ByVal psDroid As Long, ByVal psFactory As Long) As Boolean
    Dim i As Long, player As Long, receiveAll As Boolean, has_factory As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1217", "triggerEventDroidBuilt", "Scripts not initialized yet"

    If psFactory <> 0 Then ' Ignore droids that were gifted or cheated (debug add)
        ' Increment built count for droids built in a factory
        updateMultiStatsBuilt psDroid
    End If

    has_factory = (psFactory <> 0) ' optional<const STRUCTURE *> opt_factory
    For i = 0 To scripts_size - 1
        player = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If player = aDroid(psDroid).player Or receiveAll Then
            scripts(i).handle_eventDroidBuilt psDroid, has_factory, psFactory
        End If
    Next
    triggerEventDroidBuilt = True
End Function

Public Function triggerEventStructBuilt(ByVal psStruct As Long, ByVal psDroid As Long) As Boolean
    Dim i As Long, player As Long, receiveAll As Boolean, has_droid As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1246", "triggerEventStructBuilt", "Scripts not initialized yet"

    If psDroid <> 0 Then ' Ignore structures that were gifted or cheated (debug add)
        ' Increment structure count for structures built by a droid
        updateMultiStatsBuilt psStruct
    End If

    has_droid = (psDroid <> 0) ' optional<const DROID *> opt_droid
    For i = 0 To scripts_size - 1
        player = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If player = aStruct(psStruct).player Or receiveAll Then
            scripts(i).handle_eventStructureBuilt psStruct, has_droid, psDroid
        End If
    Next
    triggerEventStructBuilt = True
End Function

Public Function triggerEventStructDemolish(ByVal psStruct As Long, ByVal psDroid As Long) As Boolean
    Dim i As Long, player As Long, receiveAll As Boolean, has_droid As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1274", "triggerEventStructDemolish", "Scripts not initialized yet"
    has_droid = (psDroid <> 0) ' optional<const DROID *> opt_droid
    For i = 0 To scripts_size - 1
        player = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If player = aStruct(psStruct).player Or receiveAll Then
            scripts(i).handle_eventStructureDemolish psStruct, has_droid, psDroid
        End If
    Next
    triggerEventStructDemolish = True
End Function

Public Function triggerEventStructureReady(ByVal psStruct As Long) As Boolean
    Dim i As Long, player As Long, receiveAll As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1296", "triggerEventStructureReady", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        player = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If player = aStruct(psStruct).player Or receiveAll Then
            scripts(i).handle_eventStructureReady psStruct
        End If
    Next
    triggerEventStructureReady = True
End Function

Public Function triggerEventStructureUpgradeStarted(ByVal psStruct As Long) As Boolean
    Dim i As Long, player As Long, receiveAll As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1315", "triggerEventStructureUpgradeStarted", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        player = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If player = aStruct(psStruct).player Or receiveAll Then
            scripts(i).handle_eventStructureUpgradeStarted psStruct
        End If
    Next
    triggerEventStructureUpgradeStarted = True
End Function

Public Function triggerEventDroidRankGained(ByVal psDroid As Long, ByVal rankNum As Long) As Boolean
    Dim i As Long, player As Long, receiveAll As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1334", "triggerEventDroidRankGained", "Scripts not initialized yet"
    If psDroid = 0 Then triggerEventDroidRankGained = True: Exit Function
    For i = 0 To scripts_size - 1
        player = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If player = aDroid(psDroid).player Or receiveAll Then
            scripts(i).handle_eventDroidRankGained psDroid, rankNum
        End If
    Next
    triggerEventDroidRankGained = True
End Function

Public Function triggerEventAttacked(ByVal psVictim As Long, ByVal psAttacker As Long, ByVal lastHit As Long) As Boolean
    Dim i As Long, player As Long, receiveAll As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1355", "triggerEventAttacked", "Scripts not initialized yet"
    If psAttacker = 0 Then
        ' do not fire off this event if there is no attacker -- nothing do respond to
        ' (FIXME -- consider this carefully)
        triggerEventAttacked = False
        Exit Function
    End If
    ' throttle the event for performance
    If ULT(gameTime - lastHit, ATTACK_THROTTLE) Then
        triggerEventAttacked = False
        Exit Function
    End If
    For i = 0 To scripts_size - 1
        player = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If player = SIMPLE_OBJECT_player(psVictim) Or receiveAll Then
            scripts(i).handle_eventAttacked psVictim, psAttacker
        End If
    Next
    triggerEventAttacked = True
End Function

Public Function triggerEventResearched(ByVal psResearch As Long, ByVal psStruct As Long, ByVal player As Long) As Boolean
    Dim i As Long, me_ As Long, receiveAll As Boolean, rr As wzapi_researchResult
    'HACK: This event can be triggered when loading savegames, before the script engines are initialized.
    ' if this is the case, we need to store these events and replay them later
    If Not scriptsReady Then
        ' eventQueue.emplace(psResearch, psStruct, player)
        If eventQueue_size = 0 Then
            ReDim eventQueue(0 To 7)
        ElseIf eventQueue_size > UBound(eventQueue) Then
            ReDim Preserve eventQueue(0 To 2 * eventQueue_size - 1)
        End If
        eventQueue(eventQueue_size).research = psResearch
        eventQueue(eventQueue_size).structure = psStruct
        eventQueue(eventQueue_size).player = player
        eventQueue_size = eventQueue_size + 1
        triggerEventResearched = True
        Exit Function
    End If

    updateMultiStatsResearchComplete psResearch, player
    GameStoryLogger_logResearchCompleted psResearch, psStruct, player

    For i = 0 To scripts_size - 1
        me_ = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If me_ = player Or receiveAll Then
            rr.psResearch = psResearch  ' wzapi::researchResult(psResearch, player)
            rr.player = player
            scripts(i).handle_eventResearched rr, psStruct, player
        End If
    Next
    triggerEventResearched = True
End Function

Public Function triggerEventDestroyed(ByVal psVictim As Long) As Boolean
    Dim i As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1418", "triggerEventDestroyed", "Scripts not initialized yet"
    If psVictim = 0 Then triggerEventDestroyed = True: Exit Function
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventDestroyed psVictim
    Next
    triggerEventDestroyed = True
End Function

Public Function triggerEventPickup(ByVal psFeat As Long, ByVal psDroid As Long) As Boolean
    Dim i As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1435", "triggerEventPickup", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventPickup psFeat, psDroid
    Next
    triggerEventPickup = True
End Function

Public Function triggerEventSeen(ByVal psViewer As Long, ByVal psSeen As Long) As Boolean
    triggerEventSeen = scripting_engine_triggerEventSeen(psViewer, psSeen)
End Function

Public Function scripting_engine_triggerEventSeen(ByVal psViewer As Long, ByVal psSeen As Long) As Boolean
    Dim i As Long, callbacksFirst As Boolean, groupId As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1464", "triggerEventSeen", "Scripts not initialized yet"
    If psSeen = 0 Or psViewer = 0 Then scripting_engine_triggerEventSeen = True: Exit Function
    For i = 0 To scripts_size - 1
        callbacksFirst = scripting_engine_seenLabelCheck(scripts(i), psSeen, psViewer, groupId)
        If callbacksFirst Then
            scripts(i).handle_eventObjectSeen psViewer, psSeen
        End If
        If groupId <> 0 Then
            scripts(i).handle_eventGroupSeen psViewer, groupId
        End If
    Next
    scripting_engine_triggerEventSeen = True
End Function

Public Function triggerEventObjectTransfer(ByVal psObj As Long, ByVal from As Long) As Boolean
    Dim i As Long, me_ As Long, receiveAll As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1491", "triggerEventObjectTransfer", "Scripts not initialized yet"
    If psObj = 0 Then triggerEventObjectTransfer = True: Exit Function
    For i = 0 To scripts_size - 1
        me_ = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If me_ = SIMPLE_OBJECT_player(psObj) Or me_ = from Or receiveAll Then
            scripts(i).handle_eventObjectTransfer psObj, from
        End If
    Next
    triggerEventObjectTransfer = True
End Function

' message: vbNullString for nullptr
Public Function triggerEventChat(ByVal from As Long, ByVal to_ As Long, ByVal message As String) As Boolean
    Dim i As Long, me_ As Long, receiveAll As Boolean
    If Not scriptsReady Then
        ' just ignore chat messages before scripts are ready / initialized
        triggerEventChat = False
        Exit Function
    End If
    If Not (StrPtr(message) <> 0) Then ASSERT_FAILURE "message", "src/qtscript.cpp:1518", "triggerEventChat", "No message provided": triggerEventChat = False: Exit Function
    For i = 0 To scripts_size - 1
        me_ = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If me_ = to_ Or (receiveAll And to_ = from) Then
            scripts(i).handle_eventChat from, to_, message
        End If
    Next
    triggerEventChat = True
End Function

' message: WzQuickChatMessage
Public Function triggerEventQuickChatMessage(ByVal from As Long, ByVal to_ As Long, ByVal message As Long, ByVal teamSpecific As Boolean) As Boolean
    Dim i As Long, me_ As Long, receiveAll As Boolean
    If Not scriptsReady Then
        ' just ignore chat messages before scripts are ready / initialized
        triggerEventQuickChatMessage = False
        Exit Function
    End If
    For i = 0 To scripts_size - 1
        me_ = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If me_ = to_ Or (receiveAll And to_ = from) Then
            scripts(i).handle_eventQuickChat from, to_, message, teamSpecific
        End If
    Next
    triggerEventQuickChatMessage = True
End Function

' message: vbNullString for nullptr
Public Function triggerEventBeacon(ByVal from As Long, ByVal to_ As Long, ByVal message As String, ByVal x As Long, ByVal y As Long) As Boolean
    Dim i As Long, me_ As Long, receiveAll As Boolean, has_message As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1566", "triggerEventBeacon", "Scripts not initialized yet"
    has_message = (StrPtr(message) <> 0) ' optional<const char *> opt_message
    For i = 0 To scripts_size - 1
        me_ = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If me_ = to_ Or receiveAll Then
            scripts(i).handle_eventBeacon map_coord(x), map_coord(y), from, to_, has_message, message
        End If
    Next
    triggerEventBeacon = True
End Function

Public Function triggerEventBeaconRemoved(ByVal from As Long, ByVal to_ As Long) As Boolean
    Dim i As Long, me_ As Long, receiveAll As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1587", "triggerEventBeaconRemoved", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        me_ = scripts(i).player()
        receiveAll = scripts(i).isReceivingAllEvents()
        If me_ = to_ Or receiveAll Then
            scripts(i).handle_eventBeaconRemoved from, to_
        End If
    Next
    triggerEventBeaconRemoved = True
End Function

Public Function triggerEventSelected() As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1612", "triggerEventSelected", "Scripts not initialized yet"
    selectionChanged = True
    triggerEventSelected = True
End Function

' Since groups are entities local to one context, we do not iterate over them here.
Public Function triggerEventGroupLoss(ByVal psObj As Long, ByVal group As Long, ByVal size As Long, ByVal instance As ChakraScriptingInstance) As Boolean
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1625", "triggerEventGroupLoss", "Scripts not initialized yet"
    triggerEventGroupLoss = instance.handle_eventGroupLoss(psObj, group, size)
End Function

' This is not a trigger yet.
Public Function triggerEventDroidMoved(ByVal psDroid As Long, ByVal oldx As Long, ByVal oldy As Long) As Boolean
    triggerEventDroidMoved = scripting_engine_areaLabelCheck(psDroid)
End Function

Public Function triggerEventArea(ByVal label_ As String, ByVal psDroid As Long) As Boolean
    Dim i As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1643", "triggerEventArea", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventArea label_, psDroid
    Next
    triggerEventArea = True
End Function

Public Function triggerEventDesignCreated(ByVal psTemplate As Long) As Boolean
    Dim i As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1658", "triggerEventDesignCreated", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventDesignCreated psTemplate
    Next
    triggerEventDesignCreated = True
End Function

Public Function triggerEventAllianceOffer(ByVal from As Long, ByVal to_ As Long) As Boolean
    Dim i As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1672", "triggerEventAllianceOffer", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventAllianceOffer from, to_
    Next
    triggerEventAllianceOffer = True
End Function

Public Function triggerEventAllianceAccepted(ByVal from As Long, ByVal to_ As Long) As Boolean
    Dim i As Long
    If Not scriptsReady Then
        triggerEventAllianceAccepted = False 'silently ignore
        Exit Function
    End If
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventAllianceAccepted from, to_
    Next
    triggerEventAllianceAccepted = True
End Function

Public Function triggerEventAllianceBroken(ByVal from As Long, ByVal to_ As Long) As Boolean
    Dim i As Long
    If Not scriptsReady Then
        triggerEventAllianceBroken = False 'silently ignore
        Exit Function
    End If
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventAllianceBroken from, to_
    Next
    triggerEventAllianceBroken = True
End Function

Public Function triggerEventSyncRequest(ByVal from As Long, ByVal req_id As Long, ByVal x As Long, ByVal y As Long, ByVal psObj As Long, ByVal psObj2 As Long) As Boolean
    Dim i As Long
    If Not (scriptsReady) Then ASSERT_FAILURE "scriptsReady", "src/qtscript.cpp:1722", "triggerEventSyncRequest", "Scripts not initialized yet"
    For i = 0 To scripts_size - 1
        scripts(i).handle_eventSyncRequest from, req_id, x, y, psObj, psObj2
    Next
    triggerEventSyncRequest = True
End Function

' ----

' std::vector<scripting_engine::LabelInfo>: the count, the labels' information in results(0 .. count - 1)
Public Function scripting_engine_debug_GetLabelInfo(results() As LabelInfo) As Long
    Dim i As Long, c As String
    If labels_size > 0 Then ReDim results(0 To labels_size - 1)
    For i = 0 To labels_size - 1
        results(i).label = labels_keys(i)
        c = "?"
        Select Case labels_values(i).typ
        Case OBJ_DROID: c = "DROID"
        Case OBJ_FEATURE: c = "FEATURE"
        Case OBJ_STRUCTURE: c = "STRUCTURE"
        Case SCRIPT_POSITION: c = "POSITION"
        Case SCRIPT_AREA: c = "AREA"
        Case SCRIPT_RADIUS: c = "RADIUS"
        Case SCRIPT_GROUP: c = "GROUP"
        Case SCRIPT_PLAYER, SCRIPT_RESEARCH, OBJ_PROJECTILE, SCRIPT_COUNT: c = "ERROR"
        End Select
        results(i).typ = c
        Select Case labels_values(i).triggered
        Case -1: results(i).trigger = "N/A"
        Case 0: results(i).trigger = "Active"
        Case Else: results(i).trigger = "Done"
        End Select
        If labels_values(i).player = ALL_PLAYERS Then
            results(i).owner = "ALL"
        Else
            results(i).owner = CStr(labels_values(i).player)
        End If
        If labels_values(i).subscriber = ALL_PLAYERS Then
            results(i).subscriber = "ALL"
        Else
            results(i).subscriber = CStr(labels_values(i).subscriber)
        End If
    Next
    scripting_engine_debug_GetLabelInfo = labels_size
End Function

Public Sub clearMarks()
    Dim x As Long, y As Long, psTile As Long
    For y = 0 To mapHeight - 1
        For x = 0 To mapWidth - 1 ' clear old marks
            psTile = mapTile(x, y)
            psMapTiles(psTile).tileInfoBits = psMapTiles(psTile).tileInfoBits And Not BITS_MARKED
        Next
    Next
End Sub

Public Sub scripting_engine_markAllLabels(ByVal only_active As Boolean)
    Dim i As Long
    For i = 0 To labels_size - 1
        If Not only_active Or labels_values(i).triggered <= 0 Then
            scripting_engine_showLabel labels_keys(i), False, False
        End If
    Next
End Sub

Public Sub scripting_engine_showLabel(ByVal key As String, Optional ByVal clear_old As Boolean = True, Optional ByVal jump_to As Boolean = True)
    Dim li As Long, l As SCRIPT_LABEL, maxx As Long, maxy As Long, x As Long, y As Long, psTile As Long, psObj As Long
    Dim cameraMoved As Boolean, i As Long, k As Variant, m As Object
    If labels_find(key) < 0 Then
        debug_ 1824, LOG_ERROR, "showLabel", "label %s not found", key
        Exit Sub
    End If
    li = labels_index(key)
    l = labels_values(li)
    If clear_old Then
        clearMarks
    End If
    If l.typ = SCRIPT_AREA Or l.typ = SCRIPT_POSITION Then
        If jump_to Then
            setViewPos map_coord(l.p1.x), map_coord(l.p1.y), False ' move camera position
        End If
        maxx = map_coord(l.p2.x)
        maxy = map_coord(l.p2.y)
        If l.typ = SCRIPT_POSITION Then
            maxx = MIN_(mapWidth, maxx + 1)
            maxy = MIN_(mapHeight, maxy + 1)
        End If
        For x = map_coord(l.p1.x) To maxx - 1 ' make new ones
            For y = map_coord(l.p1.y) To maxy - 1
                psTile = mapTile(x, y)
                psMapTiles(psTile).tileInfoBits = psMapTiles(psTile).tileInfoBits Or BITS_MARKED
            Next
        Next
    ElseIf l.typ = SCRIPT_RADIUS Then
        If jump_to Then
            setViewPos map_coord(l.p1.x), map_coord(l.p1.y), False ' move camera position
        End If
        For x = MAX_(map_coord(l.p1.x - l.p2.x), 0) To MIN_(map_coord(l.p1.x + l.p2.x), mapWidth) - 1
            For y = MAX_(map_coord(l.p1.y - l.p2.x), 0) To MIN_(map_coord(l.p1.y + l.p2.x), mapHeight) - 1 ' l.p2.x is radius, not a bug
                If iHypot(map_coord(l.p1.x) - x, map_coord(l.p1.y) - y) < map_coord(l.p2.x) Then
                    psTile = mapTile(x, y)
                    psMapTiles(psTile).tileInfoBits = psMapTiles(psTile).tileInfoBits Or BITS_MARKED
                End If
            Next
        Next
    ElseIf l.typ = OBJ_DROID Or l.typ = OBJ_FEATURE Or l.typ = OBJ_STRUCTURE Then
        psObj = IdToObject(l.typ, l.id, l.player)
        If psObj <> 0 Then
            If jump_to Then
                setViewPos map_coord(SIMPLE_OBJECT_pos(psObj).x), map_coord(SIMPLE_OBJECT_pos(psObj).y), False ' move camera position
            End If
            psTile = mapTile(map_coord(SIMPLE_OBJECT_pos(psObj).x), map_coord(SIMPLE_OBJECT_pos(psObj).y))
            psMapTiles(psTile).tileInfoBits = psMapTiles(psTile).tileInfoBits Or BITS_MARKED
        End If
    ElseIf l.typ = SCRIPT_GROUP Then
        cameraMoved = False
        For i = 0 To groups_size - 1
            Set m = groups_values(i).map_()
            For Each k In m.Keys
                If m.Item(k) = l.id Then
                    psObj = k
                    If Not cameraMoved And jump_to Then
                        setViewPos map_coord(SIMPLE_OBJECT_pos(psObj).x), map_coord(SIMPLE_OBJECT_pos(psObj).y), False ' move camera position
                        cameraMoved = True
                    End If
                    psTile = mapTile(map_coord(SIMPLE_OBJECT_pos(psObj).x), map_coord(SIMPLE_OBJECT_pos(psObj).y))
                    psMapTiles(psTile).tileInfoBits = psMapTiles(psTile).tileInfoBits Or BITS_MARKED
                End If
            Next
        Next
    End If
End Sub

' The bool return value is true when an object callback needs to be called.
' The int return value holds group id when a group callback needs to be called, 0 otherwise.
' (VB) std::pair<bool, int>: the bool returned, the int in groupOut
Private Function scripting_engine_seenLabelCheck(ByVal instance As ChakraScriptingInstance, ByVal seen As Long, ByVal viewer As Long, groupOut As Long) As Boolean
    Dim psMap As GROUPMAP, groupId As Long, foundObj As Boolean, foundGroup As Boolean, i As Long
    Dim m As Object, viewerPlayer As Long, viewerId As Long, seenId As Long
    groupOut = 0
    Set psMap = scripting_engine_getGroupMap(instance)
    If Not (Not psMap Is Nothing) Then ASSERT_FAILURE "psMap != nullptr", "src/qtscript.cpp:1914", "seenLabelCheck", "Non-existent groupmap for engine": scripting_engine_seenLabelCheck = False: Exit Function
    Set m = psMap.map_()
    If m.Exists(seen) Then groupId = m.Item(seen) Else groupId = 0
    viewerPlayer = SIMPLE_OBJECT_player(viewer)
    viewerId = SIMPLE_OBJECT_id(viewer)
    seenId = SIMPLE_OBJECT_id(seen)
    For i = 0 To labels_size - 1
        If labels_values(i).triggered <> 0 Or Not (labels_values(i).subscriber = ALL_PLAYERS Or labels_values(i).subscriber = viewerPlayer) Then
            GoTo continue_
        End If

        ' Don't let a seen game object ID which matches a group label ID to prematurely
        ' trigger a group label.
        If labels_values(i).typ <> SCRIPT_GROUP And labels_values(i).id = seenId Then
            labels_values(i).triggered = viewerId ' record who made the discovery
            foundObj = True
        ElseIf labels_values(i).typ = SCRIPT_GROUP And labels_values(i).id = groupId Then
            labels_values(i).triggered = viewerId ' record who made the discovery
            foundGroup = True
        End If
continue_:
    Next
    If foundObj Or foundGroup Then
        jsDebugUpdateLabels
    End If
    If foundGroup Then groupOut = groupId
    scripting_engine_seenLabelCheck = foundObj
End Function

Public Function scripting_engine_areaLabelCheck(ByVal psDroid As Long) As Boolean
    Dim x As Long, y As Long, activated As Boolean, i As Long, key As String, inside As Boolean
    x = aDroid(psDroid).pos.x
    y = aDroid(psDroid).pos.y
    activated = False
    i = 0
    Do While i < labels_size
        key = labels_keys(i)
        inside = False
        With labels_values(i)
            If .triggered = 0 And (.subscriber = ALL_PLAYERS Or .subscriber = aDroid(psDroid).player) Then
                If .typ = SCRIPT_AREA Then
                    inside = .p1.x < x And .p1.y < y And .p2.x > x And .p2.y > y
                ElseIf .typ = SCRIPT_RADIUS Then
                    inside = iHypot(.p1.x - x, .p1.y - y) < .p2.x
                End If
            End If
        End With
        If inside Then
            ' We're inside an untriggered area
            activated = True
            labels_values(i).triggered = aDroid(psDroid).id
            triggerEventArea key, psDroid
        End If
        i = labels_upper_bound(key) ' i++ (the scripts may have changed the labels)
    Loop
    If activated Then
        jsDebugUpdateLabels
    End If
    scripting_engine_areaLabelCheck = activated
End Function

' ----------------------------------------------------------------------------------------
' Group system
'

Public Function scripting_engine_getGroupMap(ByVal instance As ChakraScriptingInstance) As GROUPMAP
    Dim i As Long
    i = groups_find(instance)
    If i >= 0 Then Set scripting_engine_getGroupMap = groups_values(i) Else Set scripting_engine_getGroupMap = Nothing
End Function

Private Sub scripting_engine_removeFromGroup(ByVal instance As ChakraScriptingInstance, ByVal psMap As GROUPMAP, ByVal psObj As Long)
    Dim groupId As Long, newGroupSize As Long
    If psMap.removeObjectFromGroup(psObj, groupId) Then
        newGroupSize = psMap.groupSize(groupId)
        instance.updateGroupSizes groupId, newGroupSize
        triggerEventGroupLoss psObj, groupId, newGroupSize, instance
    End If
End Sub

Public Sub scripting_engine_groupRemoveObject(ByVal psObj As Long)
    Dim i As Long
    For i = 0 To groups_size - 1
        scripting_engine_removeFromGroup groups_keys(i), groups_values(i), psObj
    Next
End Sub

Private Function scripting_engine_groupAddObject(ByVal psObj As Long, ByVal groupId As Long, ByVal instance As ChakraScriptingInstance) As Boolean
    Dim psMap As GROUPMAP, ok As Boolean
    ok = psObj <> 0
    If ok Then ok = Not instance Is Nothing
    If Not (ok) Then ASSERT_FAILURE "psObj && instance", "src/qtscript.cpp:2004", "groupAddObject", "Bad parameter": scripting_engine_groupAddObject = False: Exit Function
    Set psMap = scripting_engine_getGroupMap(instance)
    scripting_engine_removeFromGroup instance, psMap, psObj
    psMap.insertObjectIntoGroup psObj, groupId
    instance.updateGroupSizes groupId, psMap.groupSize(groupId)
    scripting_engine_groupAddObject = True ' inserted
End Function

Public Function scripting_engine_loadGroup(ByVal instance As ChakraScriptingInstance, ByVal groupId As Long, ByVal objId As Long) As Boolean
    Dim psObj As Long
    psObj = IdToPointer(objId, ANYPLAYER)
    If Not (psObj <> 0) Then ASSERT_FAILURE "psObj", "src/qtscript.cpp:2015", "loadGroup", "Non-existent object %d in group %d in savegame", objId, groupId: scripting_engine_loadGroup = False: Exit Function
    scripting_engine_loadGroup = scripting_engine_groupAddObject(psObj, groupId, instance)
End Function

Public Function scripting_engine_saveGroups(result As nlohmann_json, ByVal instance As ChakraScriptingInstance) As Boolean
    Dim psMap As GROUPMAP, m As Object, k As Variant, psObj As Long, value() As String, n As Long, arr As nlohmann_json, j As Long
    ' Save group info as a list of group memberships for each droid
    Set psMap = scripting_engine_getGroupMap(instance)
    If Not (Not psMap Is Nothing) Then ASSERT_FAILURE "psMap", "src/qtscript.cpp:2023", "saveGroups", "Non-existent groupmap for engine": scripting_engine_saveGroups = False: Exit Function
    result.set_item "lastNewGroupId", json_value(psMap.getLastNewGroupId())
    Set m = psMap.map_()
    For Each k In m.Keys
        psObj = k
        If Not (Not isDead(psObj)) Then ASSERT_FAILURE "!isDead(psObj)", "src/qtscript.cpp:2028", "saveGroups", "Wanted to save dead %s to savegame!", objInfo(psObj)
        n = json_getValue(result, CStr(SIMPLE_OBJECT_id(psObj))).toWzStringList(value)
        Set arr = json_array()
        For j = 0 To n - 1
            arr.push_back json_value(value(j))
        Next
        arr.push_back json_value(CStr(m.Item(k)))
        result.set_item CStr(SIMPLE_OBJECT_id(psObj)), arr
    Next
    scripting_engine_saveGroups = True
End Function

' ----------------------------------------------------------------------------------------
' Label system (function defined in qtscript.h header)
'

' fixedMapIdToGeneratedId: a Scripting.Dictionary of Longs; moduleToBuilding: one Dictionary of Longs per player slot
Public Function loadLabels(ByVal filename As String, ByVal fixedMapIdToGeneratedId As Object, moduleToBuilding() As Object, ByVal UserSaveGame As Boolean) As Boolean
    loadLabels = scripting_engine_loadLabels(filename, fixedMapIdToGeneratedId, moduleToBuilding, UserSaveGame)
End Function

' Load labels
Public Function scripting_engine_loadLabels(ByVal filename As String, ByVal fixedMapIdToGeneratedId As Object, moduleToBuilding() As Object, ByVal UserSaveGame As Boolean) As Boolean
    Dim groupidx As Long, ini As WzConfig, list() As String, n As Long, i As Long, p As SCRIPT_LABEL, blank As SCRIPT_LABEL
    Dim label_ As String, id As Long, player As Long, checkFoundObject As Long, memberList() As String
    Dim nMembers As Long, j As Long, errText As String, psObj As Long, li As Long
    groupidx = -1

    If PHYSFS_exists(filename) = 0 Then
        debug_ 2052, LOG_SAVE, "loadLabels", "No %s found -- not adding any labels", filename
        scripting_engine_loadLabels = False
        Exit Function
    End If
    Set ini = WzConfig_new(filename, WzConfig_ReadOnly)
    labels_clear
    n = ini.childGroups(list)
    debug_ 2058, LOG_SAVE, "loadLabels", "Loading %zu labels... (fixedMapToGeneratedId.count = %zu)", n, fixedMapIdToGeneratedId.Count
    For i = 0 To n - 1
        ini.beginGroup list(i)
        p = blank
        label_ = ini.value("label").toWzString()
        If labels_find(label_) >= 0 Then
            debug_ 2066, LOG_ERROR, "loadLabels", "Duplicate label found"
        ElseIf Left$(list(i), 8) = "position" Then
            ini.vector2i "pos", p.p1.x, p.p1.y
            p.p2 = p.p1
            p.typ = SCRIPT_POSITION
            p.player = ALL_PLAYERS
            p.id = -1
            p.triggered = ini.value("triggered", json_variant_new(-1)).toInt() ' deactivated by default
            p.subscriber = ALL_PLAYERS
            li = labels_index(label_)
            labels_values(li) = p
        ElseIf Left$(list(i), 4) = "area" Then
            ini.vector2i "pos1", p.p1.x, p.p1.y
            ini.vector2i "pos2", p.p2.x, p.p2.y
            p.typ = SCRIPT_AREA
            p.player = ini.value("player", json_variant_new(ALL_PLAYERS)).toInt()
            p.triggered = ini.value("triggered", json_variant_new(0)).toInt() ' activated by default
            p.id = -1
            p.subscriber = ini.value("subscriber", json_variant_new(ALL_PLAYERS)).toInt()
            li = labels_index(label_)
            labels_values(li) = p
        ElseIf Left$(list(i), 6) = "radius" Then
            ini.vector2i "pos", p.p1.x, p.p1.y
            p.p2.x = ini.value("radius").toInt()
            p.p2.y = 0 ' unused
            p.typ = SCRIPT_RADIUS
            p.player = ini.value("player", json_variant_new(ALL_PLAYERS)).toInt()
            p.triggered = ini.value("triggered", json_variant_new(0)).toInt() ' activated by default
            p.subscriber = ini.value("subscriber", json_variant_new(ALL_PLAYERS)).toInt()
            p.id = -1
            li = labels_index(label_)
            labels_values(li) = p
        ElseIf Left$(list(i), 6) = "object" Then
            id = ini.value("id").toInt()
            If Not (id > 0) Then ASSERT_FAILURE "id > 0", "src/qtscript.cpp:2105", "loadLabels", "Unexpected id %d for object label", id
            If fixedMapIdToGeneratedId.Exists(id) Then
                ' replace fixed hard-coded map-load id with its new generated (synchronized) id
                ' note: must come *before* the moduleToBuilding call below
                debug_ 2111, LOG_MAP, "loadLabels", "replaced fixed map id %d with %d", id, fixedMapIdToGeneratedId.Item(id)
                id = fixedMapIdToGeneratedId.Item(id)
            End If
            player = ini.value("player").toInt()
            If moduleToBuilding(player).Exists(id) Then
                ' replace moduleId with its building id
                debug_ 2119, LOG_MAP, "loadLabels", "replaced with %i;%i", id, moduleToBuilding(player).Item(id)
                id = moduleToBuilding(player).Item(id)
            End If
            p.id = id
            p.typ = ini.value("type").toInt()
            p.player = player
            p.triggered = ini.value("triggered", json_variant_new(-1)).toInt() ' deactivated by default
            p.subscriber = ini.value("subscriber", json_variant_new(ALL_PLAYERS)).toInt()
            checkFoundObject = IdToObject(p.typ, p.id, p.player)
            If Not UserSaveGame Then
                If Not (checkFoundObject <> 0) Then ASSERT_FAILURE "checkFoundObject != nullptr", "src/qtscript.cpp:2130", "loadLabels", "Failed to find object that label references: %s", label_
            ElseIf checkFoundObject = 0 Then
                debug_ 2134, LOG_SAVEGAME, "loadLabels", "Failed to find object that label references (probably destroyed before save): %s", label_
            End If
            li = labels_index(label_)
            labels_values(li) = p
        ElseIf Left$(list(i), 5) = "group" Then
            p.id = groupidx
            groupidx = groupidx - 1
            p.typ = SCRIPT_GROUP
            p.player = ini.value("player").toInt()
            nMembers = ini.value("members").toWzStringList(memberList)
            For j = 0 To nMembers - 1
                If Not std_stoi(memberList(j), id, errText) Then id = 0 ' WzString::toInt
                If Not (id > 0) Then ASSERT_FAILURE "id > 0", "src/qtscript.cpp:2147", "loadLabels", "Unexpected id %d for group member", id
                If fixedMapIdToGeneratedId.Exists(id) Then
                    ' replace fixed hard-coded map-load id with its new generated (synchronized) id
                    debug_ 2152, LOG_NEVER, "loadLabels", "replaced fixed map id %d with %d", id, fixedMapIdToGeneratedId.Item(id)
                    id = fixedMapIdToGeneratedId.Item(id)
                End If
                psObj = IdToPointer(id, p.player)
                If Not (psObj <> 0) Then ASSERT_FAILURE "psObj", "src/qtscript.cpp:2156", "loadLabels", "Unit %d belonging to player %d not found from label %s", id, p.player, list(i)
                SCRIPT_LABEL_idlist_push_back p, id
            Next
            p.triggered = ini.value("triggered", json_variant_new(-1)).toInt() ' deactivated by default
            p.subscriber = ini.value("subscriber", json_variant_new(ALL_PLAYERS)).toInt()
            li = labels_index(label_)
            labels_values(li) = p
        Else
            debug_ 2166, LOG_ERROR, "loadLabels", "Misnamed group in %s", filename
        End If
        ini.endGroup
    Next
    scripting_engine_loadLabels = True
End Function

Public Function writeLabels(ByVal filename As String) As Boolean
    writeLabels = scripting_engine_writeLabels(filename)
End Function

Public Function scripting_engine_writeLabels(ByVal filename As String) As Boolean
    Dim c(0 To 4) As Long, ini As WzConfig, i As Long, key As String, l As SCRIPT_LABEL, list As nlohmann_json, j As Long
    ' make unique, incremental section names
    Set ini = WzConfig_new(filename, WzConfig_ReadAndWrite)
    For i = 0 To labels_size - 1
        key = labels_keys(i)
        l = labels_values(i)
        If l.typ = SCRIPT_POSITION Then
            ini.beginGroup "position_" & CStr(c(0))
            c(0) = c(0) + 1
            ini.setVector2i "pos", l.p1.x, l.p1.y
            ini.setValue "label", json_value(key)
            ini.setValue "triggered", json_value(l.triggered)
            ini.endGroup
        ElseIf l.typ = SCRIPT_AREA Then
            ini.beginGroup "area_" & CStr(c(1))
            c(1) = c(1) + 1
            ini.setVector2i "pos1", l.p1.x, l.p1.y
            ini.setVector2i "pos2", l.p2.x, l.p2.y
            ini.setValue "label", json_value(key)
            ini.setValue "player", json_value(l.player)
            ini.setValue "triggered", json_value(l.triggered)
            ini.setValue "subscriber", json_value(l.subscriber)
            ini.endGroup
        ElseIf l.typ = SCRIPT_RADIUS Then
            ini.beginGroup "radius_" & CStr(c(2))
            c(2) = c(2) + 1
            ini.setVector2i "pos", l.p1.x, l.p1.y
            ini.setValue "radius", json_value(l.p2.x)
            ini.setValue "label", json_value(key)
            ini.setValue "player", json_value(l.player)
            ini.setValue "triggered", json_value(l.triggered)
            ini.setValue "subscriber", json_value(l.subscriber)
            ini.endGroup
        ElseIf l.typ = SCRIPT_GROUP Then
            ini.beginGroup "group_" & CStr(c(3))
            c(3) = c(3) + 1
            ini.setValue "player", json_value(l.player)
            ini.setValue "triggered", json_value(l.triggered)
            Set list = json_array()
            For j = 0 To l.idlist_size - 1
                list.push_back json_value(CStr(l.idlist(j)))
            Next
            ini.setValue "members", list
            ini.setValue "label", json_value(key)
            ini.setValue "subscriber", json_value(l.subscriber)
            ini.endGroup
        Else
            ini.beginGroup "object_" & CStr(c(4))
            c(4) = c(4) + 1
            ini.setValue "id", json_value(l.id)
            ini.setValue "player", json_value(l.player)
            ini.setValue "type", json_value(l.typ)
            ini.setValue "label", json_value(key)
            ini.setValue "triggered", json_value(l.triggered)
            ini.endGroup
        End If
    Next
    scripting_engine_writeLabels = True
End Function

'--------------------------------------------------------------------------
' The label and group functions of the script API (static members of scripting_engine)

Public Sub scripting_engine_resetLabel(ByVal context As wzapi_execution_context, ByVal labelName As String, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long)
    Dim li As Long
    If Not (labels_find(labelName) >= 0) Then
        debug_ 2284, LOG_ERROR, "resetLabel", "Label %s not found", labelName
        context.throwError "labels.count(labelName) > 0", 2284, "resetLabel"
        Exit Sub
    End If
    li = labels_index(labelName)
    labels_values(li).triggered = 0 ' make active again
    If has_playerFilter Then labels_values(li).subscriber = opt_playerFilter Else labels_values(li).subscriber = ALL_PLAYERS
End Sub

Public Function scripting_engine_enumLabels(ByVal context As wzapi_execution_context, ByVal has_filterLabelType As Boolean, ByVal opt_filterLabelType As Long) As StringVector
    Dim matches As StringVector, typ As Long, i As Long
    If has_filterLabelType Then ' filter
        typ = opt_filterLabelType
        For i = 0 To labels_size - 1
            If labels_values(i).typ = typ Then
                StringVector_push_back matches, labels_keys(i)
            End If
        Next
    Else ' fast path, give all
        For i = 0 To labels_size - 1
            StringVector_push_back matches, labels_keys(i)
        Next
    End If
    scripting_engine_enumLabels = matches
End Function

'--------------------------------------------------------------------------
' struct generic_script_object

Public Function generic_script_object_new() As generic_script_object
    Dim r As generic_script_object
    r.p1.x = 0
    r.p1.y = 0
    r.p2.x = 0
    r.p2.y = 0
    r.id = -1
    r.player = -1
    r.typ = -1
    generic_script_object_new = r
End Function

Public Function generic_script_object_Null() As generic_script_object
    generic_script_object_Null = generic_script_object_new()
End Function

Public Function generic_script_object_fromRadius(ByVal x As Long, ByVal y As Long, ByVal radius As Long) As generic_script_object
    Dim result As generic_script_object
    result = generic_script_object_new()
    result.typ = SCRIPT_RADIUS
    result.p1.x = x
    result.p1.y = y
    result.p2.x = radius
    generic_script_object_fromRadius = result
End Function

Public Function generic_script_object_isNull(o As generic_script_object) As Boolean
    generic_script_object_isNull = o.typ < 0
End Function

Public Function generic_script_object_isRadius(o As generic_script_object) As Boolean
    generic_script_object_isRadius = o.typ = SCRIPT_RADIUS
End Function

Public Function generic_script_object_isArea(o As generic_script_object) As Boolean
    generic_script_object_isArea = o.typ = SCRIPT_AREA
End Function

Public Function generic_script_object_isPosition(o As generic_script_object) As Boolean
    generic_script_object_isPosition = o.typ = SCRIPT_POSITION
End Function

Public Function generic_script_object_isGroup(o As generic_script_object) As Boolean
    generic_script_object_isGroup = o.typ = SCRIPT_GROUP
End Function

Public Function generic_script_object_isObject(o As generic_script_object) As Boolean
    generic_script_object_isObject = o.typ = OBJ_DROID Or o.typ = OBJ_FEATURE Or o.typ = OBJ_STRUCTURE
End Function

Public Function generic_script_object_getType(o As generic_script_object) As Long
    generic_script_object_getType = o.typ
End Function

Public Function generic_script_object_getRadius(o As generic_script_object) As scr_radius ' if type == SCRIPT_RADIUS, returns the radius
    Dim r As scr_radius
    If Not (generic_script_object_isRadius(o)) Then ASSERT_FAILURE "isRadius()", "src/qtscript.cpp:2348", "getRadius", "generic_script_object is not a radius; type: %d", o.typ
    r.x = o.p1.x
    r.y = o.p1.y
    r.radius = o.p2.x
    generic_script_object_getRadius = r
End Function

Public Function generic_script_object_fromArea(ByVal x As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long) As generic_script_object
    Dim result As generic_script_object
    result = generic_script_object_new()
    result.typ = SCRIPT_AREA
    result.p1.x = x
    result.p1.y = y
    result.p2.x = x2
    result.p2.y = y2
    generic_script_object_fromArea = result
End Function

Public Function generic_script_object_getArea(o As generic_script_object) As scr_area ' if type == SCRIPT_AREA, returns the area
    Dim r As scr_area
    If Not (generic_script_object_isArea(o)) Then ASSERT_FAILURE "isArea()", "src/qtscript.cpp:2365", "getArea", "generic_script_object is not an area; type: %d", o.typ
    r.x1 = o.p1.x
    r.y1 = o.p1.y
    r.x2 = o.p2.x
    r.y2 = o.p2.y
    generic_script_object_getArea = r
End Function

Public Function generic_script_object_fromPosition(ByVal x As Long, ByVal y As Long) As generic_script_object
    Dim result As generic_script_object
    result = generic_script_object_new()
    result.typ = SCRIPT_POSITION
    result.p1.x = x
    result.p1.y = y
    generic_script_object_fromPosition = result
End Function

Public Function generic_script_object_getPosition(o As generic_script_object) As scr_position ' if type == SCRIPT_POSITION, returns the position
    Dim r As scr_position
    If Not (generic_script_object_isPosition(o)) Then ASSERT_FAILURE "isPosition()", "src/qtscript.cpp:2380", "getPosition", "generic_script_object is not a position; type: %d", o.typ
    r.x = o.p1.x
    r.y = o.p1.y
    generic_script_object_getPosition = r
End Function

Public Function generic_script_object_fromGroup(ByVal groupId As Long) As generic_script_object
    Dim result As generic_script_object
    result = generic_script_object_new()
    result.typ = SCRIPT_GROUP
    result.id = groupId
    generic_script_object_fromGroup = result
End Function

Public Function generic_script_object_getGroupId(o As generic_script_object) As Long ' if type == SCRIPT_GROUP, returns the groupId
    If Not (generic_script_object_isGroup(o)) Then ASSERT_FAILURE "isGroup()", "src/qtscript.cpp:2394", "getGroupId", "generic_script_object is not a group; type: %d", o.typ
    generic_script_object_getGroupId = o.id
End Function

Public Function generic_script_object_fromObject(ByVal psObj As Long) As generic_script_object
    Dim result As generic_script_object
    result = generic_script_object_new()
    If psObj = 0 Then
        generic_script_object_fromObject = generic_script_object_Null()
        Exit Function
    End If
    result.typ = SIMPLE_OBJECT_type(psObj)
    result.id = SIMPLE_OBJECT_id(psObj)
    result.player = SIMPLE_OBJECT_player(psObj)
    generic_script_object_fromObject = result
End Function

Public Function generic_script_object_getObject(o As generic_script_object) As Long ' if type == OBJ_DROID, OBJ_FEATURE, OBJ_STRUCTURE, returns the game object
    If Not (generic_script_object_isObject(o)) Then ASSERT_FAILURE "isObject()", "src/qtscript.cpp:2413", "getObject", "generic_script_object is not an object; type: %d", o.typ
    generic_script_object_getObject = IdToObject(o.typ, o.id, o.player)
End Function

Public Function generic_script_object_toNewLabel(o As generic_script_object) As SCRIPT_LABEL
    Dim value As SCRIPT_LABEL
    value.typ = o.typ
    value.id = o.id
    value.player = o.player
    value.p1.x = world_coord(o.p1.x)
    value.p1.y = world_coord(o.p1.y)
    If o.typ = SCRIPT_AREA Then
        value.p2.x = world_coord(o.p2.x)
        value.p2.y = world_coord(o.p2.y)
    ElseIf o.typ = SCRIPT_RADIUS Then
        value.p2.x = world_coord(o.p2.x)
    End If

    value.triggered = -1 ' default off

    generic_script_object_toNewLabel = value
End Function

'--------------------------------------------------------------------------

Public Sub scripting_engine_addLabel(ByVal context As wzapi_execution_context, object_ As generic_script_object, ByVal label_ As String, ByVal has_triggered As Boolean, ByVal opt_triggered As Long)
    Dim value As SCRIPT_LABEL, psObj As Long, li As Long
    value = generic_script_object_toNewLabel(object_)

    If value.typ = OBJ_DROID Or value.typ = OBJ_STRUCTURE Or value.typ = OBJ_FEATURE Then
        psObj = IdToObject(value.typ, value.id, value.player)
        If Not (psObj <> 0) Then
            debug_ 2454, LOG_ERROR, "addLabel", "Object id %d not found belonging to player %d", value.id, value.player
            context.throwError "psObj", 2454, "addLabel"
            Exit Sub
        End If
    End If

    value.triggered = -1 ' default off
    If has_triggered Then
        If Not (value.typ <> SCRIPT_POSITION) Then
            debug_ 2460, LOG_ERROR, "addLabel", "Cannot assign a trigger to a position"
            context.throwError "value.type != SCRIPT_POSITION", 2460, "addLabel"
            Exit Sub
        End If
        value.triggered = opt_triggered
    End If

    li = labels_index(label_)
    labels_values(li) = value
    jsDebugUpdateLabels
End Sub

Public Function scripting_engine_removeLabel(ByVal context As wzapi_execution_context, ByVal label_ As String) As Long
    Dim result As Long
    result = labels_erase(label_)
    jsDebugUpdateLabels
    scripting_engine_removeLabel = result
End Function

' optional<std::string>: has_value, the label in result
Public Function scripting_engine_getLabel(ByVal context As wzapi_execution_context, ByVal psObj As Long, result As String) As Boolean
    Dim tmp As game_object_identifier
    If Not (psObj <> 0) Then ASSERT_FAILURE "psObj", "src/qtscript.cpp:2490", "getLabel", "No valid object provided": scripting_engine_getLabel = False: Exit Function
    tmp.id = SIMPLE_OBJECT_id(psObj)
    tmp.player = SIMPLE_OBJECT_player(psObj)
    tmp.typ = SIMPLE_OBJECT_type(psObj)
    scripting_engine_getLabel = scripting_engine__findMatchingLabel(tmp, result)
End Function

' optional<std::string>: has_value, the label in result
Public Function scripting_engine_getLabelJS(ByVal context As wzapi_execution_context, obj_id As game_object_identifier, result As String) As Boolean
    scripting_engine_getLabelJS = scripting_engine__findMatchingLabel(obj_id, result)
End Function

Private Function scripting_engine__findMatchingLabel(obj_id As game_object_identifier, result As String) As Boolean
    Dim value As SCRIPT_LABEL, label_ As String, i As Long
    value.id = obj_id.id
    value.player = obj_id.player
    value.typ = obj_id.typ
    debug_ 2508, LOG_NEVER, "_findMatchingLabel", "looking for label %i;%i;%i", obj_id.id, obj_id.player, obj_id.typ
    For i = 0 To labels_size - 1
        If SCRIPT_LABEL_equals(labels_values(i), value) Then
            label_ = labels_keys(i)
            Exit For
        End If
    Next

    If Len(label_) <> 0 Then
        result = label_
        scripting_engine__findMatchingLabel = True
    End If
End Function

Public Function scripting_engine_getObject(ByVal context As wzapi_execution_context, request As object_request) As generic_script_object
    Dim x As Long, y As Long, psTile As Long, typ As Long, player As Long, id As Long
    If request.requestType = RequestType_MAPPOS_REQUEST Then
        ' request.getMapPosition(): {val1, val2}
        x = request.val1
        y = request.val2
        If Not (tileOnMap(x, y)) Then
            debug_ 2553, LOG_ERROR, "getObject", "Map position (%d, %d) not on the map!", x, y
            context.throwError "tileOnMap(x, y)", 2553, "getObject"
            scripting_engine_getObject = generic_script_object_new()
            Exit Function
        End If
        psTile = mapTile(x, y)
        scripting_engine_getObject = generic_script_object_fromObject(psMapTiles(psTile).psObject)
        Exit Function
    ElseIf request.requestType = RequestType_OBJECTID_REQUEST Then
        ' request.getObjectIDRequest(): {(OBJECT_TYPE)val1, val2, val3}
        typ = request.val1
        player = request.val2
        id = request.val3
        If Not (player >= 0 And player < MAX_PLAYERS) Then
            debug_ 2563, LOG_ERROR, "getObject", "Invalid player index %d", player
            context.throwError "player >= 0 && player < MAX_PLAYERS", 2563, "getObject"
            scripting_engine_getObject = generic_script_object_new()
            Exit Function
        End If
        scripting_engine_getObject = generic_script_object_fromObject(IdToObject(typ, id, player))
        Exit Function
    ElseIf request.requestType = RequestType_LABEL_REQUEST Then
        scripting_engine_getObject = scripting_engine_getObjectFromLabel(context, request.str_)
        Exit Function
    End If
    scripting_engine_getObject = generic_script_object_new()
End Function

Public Function scripting_engine_getObjectFromLabel(ByVal context As wzapi_execution_context, ByVal label_ As String) As generic_script_object
    Dim psObj As Long, li As Long, p As SCRIPT_LABEL
    ' get by label case
    psObj = 0
    If labels_find(label_) >= 0 Then
        li = labels_index(label_)
        p = labels_values(li)
        Select Case p.typ
        Case SCRIPT_RADIUS
            scripting_engine_getObjectFromLabel = generic_script_object_fromRadius(map_coord(p.p1.x), map_coord(p.p1.y), map_coord(p.p2.x))
            Exit Function
        Case SCRIPT_AREA
            scripting_engine_getObjectFromLabel = generic_script_object_fromArea(map_coord(p.p1.x), map_coord(p.p1.y), map_coord(p.p2.x), map_coord(p.p2.y))
            Exit Function
        Case SCRIPT_POSITION
            scripting_engine_getObjectFromLabel = generic_script_object_fromPosition(map_coord(p.p1.x), map_coord(p.p1.y))
            Exit Function
        Case SCRIPT_GROUP
            scripting_engine_getObjectFromLabel = generic_script_object_fromGroup(p.id)
            Exit Function
        Case OBJ_DROID, OBJ_FEATURE, OBJ_STRUCTURE
            psObj = IdToObject(p.typ, p.id, p.player)
            If psObj = 0 Then
                scripting_engine_getObjectFromLabel = generic_script_object_Null()
                Exit Function
            End If
            scripting_engine_getObjectFromLabel = generic_script_object_fromObject(psObj)
            Exit Function
        Case Else
            debug_ 2604, LOG_ERROR, "getObjectFromLabel", "Bad object label type found for label %s!", label_
            context.throwError "false", 2604, "getObjectFromLabel"
            scripting_engine_getObjectFromLabel = generic_script_object_new()
            Exit Function
        End Select
    End If
    scripting_engine_getObjectFromLabel = generic_script_object_Null()
End Function

Public Function scripting_engine_enumAreaByLabel(ByVal context As wzapi_execution_context, ByVal label_ As String, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long, ByVal has_seen As Boolean, ByVal opt_seen As Boolean, result() As Long) As Long
    Dim li As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    If Not (labels_find(label_) >= 0) Then
        debug_ 2622, LOG_ERROR, "enumAreaByLabel", "Label %s not found", label_
        context.throwError "instance().labels.count(label) > 0", 2622, "enumAreaByLabel"
        Exit Function
    End If
    li = labels_index(label_)
    If Not (labels_values(li).typ = SCRIPT_AREA) Then
        debug_ 2624, LOG_ERROR, "enumAreaByLabel", "Wrong label type for %s", label_
        context.throwError "p.type == SCRIPT_AREA", 2624, "enumAreaByLabel"
        Exit Function
    End If
    x1 = labels_values(li).p1.x
    y1 = labels_values(li).p1.y
    x2 = labels_values(li).p2.x
    y2 = labels_values(li).p2.y
    scripting_engine_enumAreaByLabel = scripting_engine__enumAreaWorldCoords(context, x1, y1, x2, y2, has_playerFilter, opt_playerFilter, has_seen, opt_seen, result)
End Function

Public Function scripting_engine_enumArea(ByVal context As wzapi_execution_context, area As scr_area, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long, ByVal has_seen As Boolean, ByVal opt_seen As Boolean, result() As Long) As Long
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    x1 = world_coord(area.x1)
    y1 = world_coord(area.y1)
    x2 = world_coord(area.x2)
    y2 = world_coord(area.y2)
    scripting_engine_enumArea = scripting_engine__enumAreaWorldCoords(context, x1, y1, x2, y2, has_playerFilter, opt_playerFilter, has_seen, opt_seen, result)
End Function

Private Function scripting_engine__enumAreaWorldCoords(ByVal context As wzapi_execution_context, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long, ByVal has_seen As Boolean, ByVal opt_seen As Boolean, result() As Long) As Long
    Static gridList() As Long           ' static to avoid allocations. // not thread-safe
    Dim player As Long, playerFilter As Long, seen As Boolean, n As Long, gi As Long, psObj As Long, cnt As Long
    Dim matches As Boolean, psObjPlayer As Long
    player = context.player()
    If has_playerFilter Then playerFilter = opt_playerFilter Else playerFilter = ALL_PLAYERS
    If has_seen Then seen = opt_seen Else seen = True

    n = gridStartIterateArea(x1, y1, x2, y2, gridList)
    For gi = 0 To n - 1
        psObj = gridList(gi)
        If (BASE_OBJECT_visible(psObj, player) <> 0 Or Not seen) And SIMPLE_OBJECT_died(psObj) = 0 Then
            psObjPlayer = SIMPLE_OBJECT_player(psObj)
            matches = (playerFilter >= 0 And psObjPlayer = playerFilter) Or playerFilter = ALL_PLAYERS
            If Not matches And SIMPLE_OBJECT_type(psObj) <> OBJ_FEATURE Then
                If playerFilter = ALLIES Then
                    matches = aiCheckAlliances(psObjPlayer, player)
                ElseIf playerFilter = ENEMIES Then
                    matches = Not aiCheckAlliances(psObjPlayer, player)
                End If
            End If
            If matches Then
                ReDim Preserve result(0 To cnt)
                result(cnt) = psObj
                cnt = cnt + 1
            End If
        End If
    Next
    scripting_engine__enumAreaWorldCoords = cnt
End Function

Public Function scripting_engine_enumAreaJS(ByVal context As wzapi_execution_context, area_lookup As area_by_values_or_area_label_lookup, ByVal has_playerFilter As Boolean, ByVal opt_playerFilter As Long, ByVal has_seen As Boolean, ByVal opt_seen As Boolean, result() As Long) As Long
    Dim optLabel As String, hasLabel As Boolean, optArea As scr_area, hasArea As Boolean
    If area_by_values_or_area_label_lookup_isLabel(area_lookup) Then
        hasLabel = area_by_values_or_area_label_lookup_label(area_lookup, optLabel)
        If Not (hasLabel) Then ASSERT_FAILURE "optLabel.has_value()", "src/qtscript.cpp:2674", "enumAreaJS", "optional label is null"
        scripting_engine_enumAreaJS = scripting_engine_enumAreaByLabel(context, optLabel, has_playerFilter, opt_playerFilter, has_seen, opt_seen, result)
    Else
        hasArea = area_by_values_or_area_label_lookup_area(area_lookup, optArea)
        If Not (hasArea) Then ASSERT_FAILURE "optArea.has_value()", "src/qtscript.cpp:2680", "enumAreaJS", "optional area is null"
        scripting_engine_enumAreaJS = scripting_engine_enumArea(context, optArea, has_playerFilter, opt_playerFilter, has_seen, opt_seen, result)
    End If
End Function

Public Function scripting_engine_enumGroup(ByVal context As wzapi_execution_context, ByVal groupId As Long, result() As Long) As Long
    Dim psMap As GROUPMAP
    Set psMap = scripting_engine_getGroupMap(context.currentInstance())

    If Not psMap Is Nothing Then
        scripting_engine_enumGroup = psMap.getGroupObjects(groupId, result)
    End If
End Function

Public Function scripting_engine_newGroup(ByVal context As wzapi_execution_context) As Long
    Dim psMap As GROUPMAP, i As Long
    Set psMap = scripting_engine_getGroupMap(context.currentInstance())
    i = psMap.newGroupID()
    ' NOTE: group zero is reserved
    If Not (i <> 0) Then
        debug_ 2711, LOG_ERROR, "newGroup", "Group 0 is reserved - error in WZ code"
        context.throwError "i != 0", 2711, "newGroup"
        scripting_engine_newGroup = 1
        Exit Function
    End If
    scripting_engine_newGroup = i
End Function

Public Sub scripting_engine_groupAddArea(ByVal context As wzapi_execution_context, ByVal groupId As Long, ByVal x1_ As Long, ByVal y1_ As Long, ByVal x2_ As Long, ByVal y2_ As Long)
    Dim player As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, it As Long, psDroid As Long
    player = context.player()
    x1 = world_coord(x1_)
    y1 = world_coord(y1_)
    x2 = world_coord(x2_)
    y2 = world_coord(y2_)

    it = apsDroidLists(player).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).pos.x >= x1 And aDroid(psDroid).pos.x <= x2 And aDroid(psDroid).pos.y >= y1 And aDroid(psDroid).pos.y <= y2 Then
            scripting_engine_groupAddObject psDroid, groupId, context.currentInstance()
        End If
        it = lnNext(it)
    Loop
End Sub

Public Sub scripting_engine_groupAddDroid(ByVal context As wzapi_execution_context, ByVal groupId As Long, ByVal psDroid As Long)
    If Not (psDroid <> 0) Then
        debug_ 2743, LOG_ERROR, "groupAddDroid", "No valid droid provided"
        context.throwError "psDroid", 2743, "groupAddDroid"
        Exit Sub
    End If
    scripting_engine_groupAddObject psDroid, groupId, context.currentInstance()
End Sub

Public Sub scripting_engine_groupAdd(ByVal context As wzapi_execution_context, ByVal groupId As Long, ByVal psObj As Long)
    If Not (psObj <> 0) Then
        debug_ 2754, LOG_ERROR, "groupAdd", "No valid object provided"
        context.throwError "psObj", 2754, "groupAdd"
        Exit Sub
    End If
    scripting_engine_groupAddObject psObj, groupId, context.currentInstance()
End Sub

Public Function scripting_engine_groupSize(ByVal context As wzapi_execution_context, ByVal groupId As Long) As Long
    Dim groupMap As GROUPMAP
    Set groupMap = scripting_engine_getGroupMap(context.currentInstance())
    If Not (Not groupMap Is Nothing) Then ASSERT_FAILURE "groupMap", "src/qtscript.cpp:2766", "groupSize", "Failed to obtain groupMap": scripting_engine_groupSize = 0: Exit Function
    scripting_engine_groupSize = groupMap.groupSize(groupId)
End Function

' ----------------------------------------------------------------------------------------
' Register functions with scripting system

Public Function scripting_engine_unregisterFunctions(ByVal instance As ChakraScriptingInstance) As Boolean
    Dim num As Long, i As Long, j As Long
    num = 0
    i = groups_find(instance)
    If i >= 0 Then
        ' groups.erase(it); delete psMap;
        For j = i To groups_size - 2
            Set groups_keys(j) = groups_keys(j + 1)
            Set groups_values(j) = groups_values(j + 1)
        Next
        groups_size = groups_size - 1
        Set groups_keys(groups_size) = Nothing
        Set groups_values(groups_size) = Nothing
        num = 1
    End If
    If Not (num = 1) Then ASSERT_FAILURE "num == 1", "src/qtscript.cpp:2784", "unregisterFunctions", "Number of engines removed from group map is %d!", num
    labels_clear
    scripting_engine_unregisterFunctions = True
End Function

' Call this just before launching game; we can't do everything in registerFunctions()
' since all game state may not be fully loaded by then
Public Sub scripting_engine_prepareLabels()
    Dim g As Long, instance As ChakraScriptingInstance, i As Long, key As String, j As Long, id As Long, psObj As Long
    Dim l As SCRIPT_LABEL
    ' load the label group data into every scripting context, with the same negative group id
    For g = 0 To groups_size - 1
        Set instance = groups_keys(g)
        i = 0
        Do While i < labels_size
            key = labels_keys(i)
            l = labels_values(i)
            If l.typ = SCRIPT_GROUP Then
                For j = 0 To l.idlist_size - 1
                    id = l.idlist(j)
                    psObj = IdToPointer(id, l.player)
                    If Not (psObj <> 0) Then ASSERT_FAILURE "psObj", "src/qtscript.cpp:2806", "prepareLabels", "Unit %d belonging to player %d not found", id, l.player
                    If psObj <> 0 Then
                        scripting_engine_groupAddObject psObj, l.id, instance
                    End If
                Next
            End If
            i = labels_upper_bound(key) ' ++i
        Loop
    Next
    jsDebugUpdateLabels
End Sub

Public Sub scripting_engine_hackMarkTiles_ByLabel(ByVal context As wzapi_execution_context, ByVal label_ As String)
    Dim l As SCRIPT_LABEL, x As Long, y As Long, psTile As Long, li As Long
    If Not (labels_find(label_) >= 0) Then
        debug_ 2820, LOG_ERROR, "hackMarkTiles_ByLabel", "Label %s not found", label_
        context.throwError "labels.count(label) > 0", 2820, "hackMarkTiles_ByLabel"
        Exit Sub
    End If

    li = labels_index(label_)
    l = labels_values(li)
    If l.typ = SCRIPT_AREA Then
        For x = map_coord(l.p1.x) To map_coord(l.p2.x) - 1
            For y = map_coord(l.p1.y) To map_coord(l.p2.y) - 1
                psTile = mapTile(x, y)
                psMapTiles(psTile).tileInfoBits = psMapTiles(psTile).tileInfoBits Or BITS_MARKED
            Next
        Next
    ElseIf l.typ = SCRIPT_RADIUS Then
        For x = map_coord(l.p1.x - l.p2.x) To map_coord(l.p1.x + l.p2.x) - 1
            For y = map_coord(l.p1.y - l.p2.x) To map_coord(l.p1.y + l.p2.x) - 1
                If x >= -1 And x < mapWidth + 1 And y >= -1 And y < mapWidth + 1 Then
                    If iHypot(map_coord(l.p1.x) - x, map_coord(l.p1.y) - y) < map_coord(l.p2.x) Then
                        psTile = mapTile(x, y)
                        psMapTiles(psTile).tileInfoBits = psMapTiles(psTile).tileInfoBits Or BITS_MARKED
                    End If
                End If
            Next
        Next
    End If
End Sub

' (private in C++; Public here: executeWithPerformanceMonitoring is written out at the backend's callers)
Public Sub scripting_engine_logFunctionPerformance(ByVal instance As ChakraScriptingInstance, ByVal function_ As String, ByVal ticks As Long)
    Dim monitor As Object, m As MONITOR_BIN, idx As Long
    Set monitor = monitors_at(instance) ' pick right one for this instance
    If monitor.Exists(function_) Then
        idx = monitor.Item(function_)
        m = monitorBins(idx)
    Else
        idx = -1
    End If
    If ticks > MAX_US Then
        debug_ 2863, LOG_SCRIPT, "logFunctionPerformance", "%s took %dus at time %d", function_, ticks, wzGetTicks()
        m.overMaxTimeCalls = m.overMaxTimeCalls + 1
    ElseIf ticks > HALF_MAX_US Then
        m.overHalfMaxTimeCalls = m.overHalfMaxTimeCalls + 1
    End If
    m.calls = m.calls + 1
    If ticks > m.worst Then
        m.worst = ticks
        m.worstGameTime = gameTime
    End If
    m.time_ = m.time_ + ticks
    ' (*monitor)[function] = m
    If idx < 0 Then
        If monitorBins_size = 0 Then
            ReDim monitorBins(0 To 63)
        ElseIf monitorBins_size > UBound(monitorBins) Then
            ReDim Preserve monitorBins(0 To 2 * monitorBins_size - 1)
        End If
        idx = monitorBins_size
        monitorBins_size = monitorBins_size + 1
        monitor.Add function_, idx
    End If
    monitorBins(idx) = m
End Sub
