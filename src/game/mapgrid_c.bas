Attribute VB_Name = "mapgrid_c"
' src/mapgrid.cpp, mapgrid.h: the objects in a quad-tree-like object
' (pointtree_c's PointTree) over the map. The GridList const & that the
' gridStartIterate functions return (a static vector of each template
' instance) is copied into the caller's array, gridList(0 .. n - 1), and n
' is returned. The template's Condition is a GRID_CONDITION.
Option Explicit

' (VB) gridStartIterateFiltered's Condition
Private Enum GRID_CONDITION
    ConditionTrue
    ConditionDroidsByPlayer
    ConditionDroidCandidateForRepair
    ConditionUnseen
End Enum

Private gridPointTree As Boolean                ' (VB) gridPointTree != nullptr
Private gridFiltersUnseen(0 To MAX_PLAYERS - 1) As PointTreeFilter
Private gridFiltersDroidsByPlayer(0 To MAX_PLAYERS - 1) As PointTreeFilter
Private gridFiltersDroidsRepairCandidates(0 To MAX_PLAYERS - 1) As PointTreeFilter

' initialise the grid system
Public Function gridInitialise() As Boolean
    Dim i As Long
    If Not (Not gridPointTree) Then ASSERT_FAILURE "gridPointTree == nullptr", "src/mapgrid.cpp:43", "gridInitialise", "gridInitialise already called, without calling gridShutDown."
    gridPointTree = True
    PointTree_clear
    For i = 0 To MAX_PLAYERS - 1
        PointTreeFilter_construct gridFiltersUnseen(i)
        PointTreeFilter_construct gridFiltersDroidsByPlayer(i)
        PointTreeFilter_construct gridFiltersDroidsRepairCandidates(i)
    Next

    gridInitialise = True                       ' Yay, nothing failed!
End Function

' (VB) gridReset's loop body: insert psObj unless it died, and clear its seenThisTick
Private Sub gridResetInsert(ByVal psObj As Long)
    Dim viewer As Long
    If SIMPLE_OBJECT_died(psObj) = 0 Then
        PointTree_insert psObj, SIMPLE_OBJECT_pos(psObj).x, SIMPLE_OBJECT_pos(psObj).y
        Select Case psObj \ OBJ_HANDLE_UNIT
        Case OBJ_DROID
            For viewer = 0 To MAX_PLAYERS - 1
                aDroid(psObj).seenThisTick(viewer) = 0
            Next
        Case OBJ_STRUCTURE
            For viewer = 0 To MAX_PLAYERS - 1
                aStruct(psObj).seenThisTick(viewer) = 0
            Next
        Case OBJ_FEATURE
            For viewer = 0 To MAX_PLAYERS - 1
                aFeature(psObj).seenThisTick(viewer) = 0
            Next
        End Select
    End If
End Sub

' reset the grid system
' (Called once per update. Resets seenThisTick[] to false.)
Public Sub gridReset()
    Dim player As Long, it As Long
    PointTree_clear

    ' Put all existing objects into the point tree.
    For player = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(player).head
        Do While it <> 0
            gridResetInsert lnObj(it)
            it = lnNext(it)
        Loop
        it = apsStructLists(player).head
        Do While it <> 0
            gridResetInsert lnObj(it)
            it = lnNext(it)
        Loop
    Next
    it = apsFeatureList(0).head
    Do While it <> 0
        gridResetInsert lnObj(it)
        it = lnNext(it)
    Loop

    PointTree_sort

    For player = 0 To MAX_PLAYERS - 1
        PointTreeFilter_reset gridFiltersUnseen(player)
        PointTreeFilter_reset gridFiltersDroidsByPlayer(player)
        PointTreeFilter_reset gridFiltersDroidsRepairCandidates(player)
    Next
End Sub

' shutdown the grid system
Public Sub gridShutDown()
    PointTree_clear
    gridPointTree = False
    ' (the filters are released with the tree in C++; gridFiltersDroidsRepairCandidates is not deleted there)
End Sub

' cast to int64 to avoid integer overflow: (VB) Doubles, exact while |x|, |y| and radius are below 2^26
Private Function isInRadius(ByVal x As Long, ByVal y As Long, ByVal radius As Long) As Boolean
    isInRadius = (CDbl(x) * x + CDbl(y) * y) <= (U32(radius) * U32(radius))
End Function

' (VB) Condition::test(obj)
Private Function conditionTest(ByVal condition As GRID_CONDITION, ByVal player As Long, ByVal obj As Long) As Boolean
    Dim isOwnOrAlly As Boolean, isVTOL As Boolean, isOnGround As Boolean
    Select Case condition
    Case ConditionTrue
        conditionTest = True
    Case ConditionDroidsByPlayer
        conditionTest = (SIMPLE_OBJECT_type(obj) = OBJ_DROID And SIMPLE_OBJECT_player(obj) = player)
    Case ConditionDroidCandidateForRepair
        If SIMPLE_OBJECT_type(obj) <> OBJ_DROID Then Exit Function
        isOwnOrAlly = (aDroid(obj).player = player) Or aiCheckAlliances(aDroid(obj).player, player)
        isVTOL = (asPropulsionStats(DROID_getPropulsionStats(obj)).propulsionType = PROPULSION_TYPE_LIFT)
        ' either it's a ground unit, or it's a VTOL on ground
        isOnGround = (Not isVTOL) Or (isVTOL And (aDroid(obj).sMove.Status = MOVEINACTIVE And aDroid(obj).sMove.iVertSpeed = 0))
        ' Note: no check for droidIsDamaged(psDroid) this is intentional
        conditionTest = aDroid(obj).died = 0 And isOwnOrAlly And isOnGround
    Case ConditionUnseen
        Select Case obj \ OBJ_HANDLE_UNIT
        Case OBJ_DROID: conditionTest = aDroid(obj).seenThisTick(player) < UINT8_MAX
        Case OBJ_STRUCTURE: conditionTest = aStruct(obj).seenThisTick(player) < UINT8_MAX
        Case OBJ_FEATURE: conditionTest = aFeature(obj).seenThisTick(player) < UINT8_MAX
        End Select
    End Select
End Function

' initialise the grid system to start iterating through units that
' could affect a location (x,y in world coords)
' (VB) filter: the player's filter of condition's kind, or none for ConditionTrue
Private Function gridStartIterateFiltered(ByVal x As Long, ByVal y As Long, ByVal radius As Long, ByVal condition As GRID_CONDITION, _
                                          ByVal player As Long, gridList() As Long) As Long
    Dim w As Long, i As Long, obj As Long, n As Long
    Select Case condition
    Case ConditionTrue
        PointTree_query x, y, radius
    Case ConditionDroidsByPlayer
        PointTree_queryFiltered gridFiltersDroidsByPlayer(player), x, y, radius
    Case ConditionDroidCandidateForRepair
        PointTree_queryFiltered gridFiltersDroidsRepairCandidates(player), x, y, radius
    Case ConditionUnseen
        PointTree_queryFiltered gridFiltersUnseen(player), x, y, radius
    End Select
    w = 0
    For i = w To lastQueryResults_size - 1
        obj = lastQueryResults(i)
        If Not conditionTest(condition, player, obj) Then        ' Check if we should skip this object.
            ' Stop the object from appearing in future searches.
            Select Case condition
            Case ConditionDroidsByPlayer: PointTreeFilter_erase gridFiltersDroidsByPlayer(player), lastFilteredQueryIndices(i)
            Case ConditionDroidCandidateForRepair: PointTreeFilter_erase gridFiltersDroidsRepairCandidates(player), lastFilteredQueryIndices(i)
            Case ConditionUnseen: PointTreeFilter_erase gridFiltersUnseen(player), lastFilteredQueryIndices(i)
            End Select
        ElseIf isInRadius(SIMPLE_OBJECT_pos(obj).x - x, SIMPLE_OBJECT_pos(obj).y - y, radius) Then  ' Check that search result is less than radius (since they can be up to a factor of sqrt(2) more).
            lastQueryResults(w) = lastQueryResults(i)
            w = w + 1
        End If
    Next
    lastQueryResults_size = w                   ' Erase all points that were a bit too far.

    ' In case you are curious.
    'debug(LOG_WARNING, "gridStartIterateFiltered(%d, %d, %u) found %u objects", x, y, radius, (unsigned)gridPointTree->lastQueryResults.size());

    n = lastQueryResults_size
    If n > 0 Then
        ReDim gridList(0 To n - 1)
        For i = 0 To n - 1
            gridList(i) = lastQueryResults(i)
        Next
    End If
    gridStartIterateFiltered = n
End Function

Private Function gridStartIterateFilteredArea(ByVal x As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long, gridList() As Long) As Long
    Dim i As Long, n As Long
    PointTree_queryArea x, y, x2, y2

    n = lastQueryResults_size
    If n > 0 Then
        ReDim gridList(0 To n - 1)
        For i = 0 To n - 1
            gridList(i) = lastQueryResults(i)
        Next
    End If
    gridStartIterateFilteredArea = n
End Function

' Find all objects within radius.
Public Function gridStartIterate(ByVal x As Long, ByVal y As Long, ByVal radius As Long, gridList() As Long) As Long
    gridStartIterate = gridStartIterateFiltered(x, y, radius, ConditionTrue, 0, gridList)
End Function

' Find all objects within radius.
Public Function gridStartIterateArea(ByVal x As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long, gridList() As Long) As Long
    gridStartIterateArea = gridStartIterateFilteredArea(x, y, x2, y2, gridList)
End Function

' Find all objects within radius where object->type == OBJ_DROID && object->player == player.
Public Function gridStartIterateDroidsByPlayer(ByVal x As Long, ByVal y As Long, ByVal radius As Long, ByVal player As Long, gridList() As Long) As Long
    gridStartIterateDroidsByPlayer = gridStartIterateFiltered(x, y, radius, ConditionDroidsByPlayer, player, gridList)
End Function

' Find all objects within radius where (object->type == OBJ_DROID && !object->died)
Public Function gridStartIterateRepairCandidates(ByVal x As Long, ByVal y As Long, ByVal radius As Long, ByVal player As Long, gridList() As Long) As Long
    gridStartIterateRepairCandidates = gridStartIterateFiltered(x, y, radius, ConditionDroidCandidateForRepair, player, gridList)
End Function

' Used for visibility.
' Find all objects within radius where object->seenThisTick[player] != 255.
Public Function gridStartIterateUnseen(ByVal x As Long, ByVal y As Long, ByVal radius As Long, ByVal player As Long, gridList() As Long) As Long
    gridStartIterateUnseen = gridStartIterateFiltered(x, y, radius, ConditionUnseen, player, gridList)
End Function

' BASE_OBJECT **gridIterateDup(): the last query's results, copied (VB: into ret, their number returned)
Public Function gridIterateDup(ret() As Long) As Long
    Dim i As Long
    If lastQueryResults_size > 0 Then
        ReDim ret(0 To lastQueryResults_size - 1)
        For i = 0 To lastQueryResults_size - 1
            ret(i) = lastQueryResults(i)
        Next
    End If
    gridIterateDup = lastQueryResults_size
End Function
