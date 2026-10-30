Attribute VB_Name = "mechanics_c"
' src/mechanics.cpp, mechanics.h: so far the players' component and
' structure lists (stats_c's apCompLists and apStructTypeLists). The
' component lists are one array apCompLists(player, type, comp) whose last
' dimension is the longest list allocated; apCompLists_count(type) is a
' type's length (0: nullptr). Not ported yet: makeAllAvailable
' (enableResearch).
Option Explicit

Private apCompListsCap As Long          ' (VB) the last dimension of apCompLists

' Shutdown the mechanics system
Public Function mechanicsShutdown() As Boolean
    Dim it As Long
    If psDestroyedObj.size <> 0 Then
        debug_ 42, LOG_INFO, "mechanicsShutdown", "%zu destroyed objects", psDestroyedObj.size
    End If
    it = psDestroyedObj.head
    Do While it <> 0
        objmemDestroy lnObj(it), True
        it = lnNext(it)
    Loop
    objlist_clear psDestroyedObj

    mechanicsShutdown = True
End Function

' Allocate the list for a component
Public Function allocComponentList(ByVal typ As COMPONENT_TYPE, ByVal number As Long) As Boolean
    Dim inc As Long, comp As Long

    ' (VB) room for the list (C: each player's list freed and allocated)
    If number > apCompListsCap Then
        If apCompListsCap = 0 Then
            ReDim apCompLists(0 To MAX_PLAYERS - 1, 0 To COMP_NUMCOMPONENTS - 1, 0 To number - 1)
        Else
            ReDim Preserve apCompLists(0 To MAX_PLAYERS - 1, 0 To COMP_NUMCOMPONENTS - 1, 0 To number - 1)
        End If
        apCompListsCap = number
    End If
    apCompLists_count(typ) = number

    'allocate the space for the Players' component lists
    For inc = 0 To MAX_PLAYERS - 1
        'initialise the players' lists
        For comp = 0 To number - 1
            apCompLists(inc, typ, comp) = UNAVAILABLE
        Next
    Next

    allocComponentList = True
End Function

' release all the component lists
Public Sub freeComponentLists()
    Dim typ As Long
    'free the component lists
    For typ = 0 To COMP_NUMCOMPONENTS - 1
        apCompLists_count(typ) = 0
    Next
    Erase apCompLists
    apCompListsCap = 0
End Sub

'allocate the space for the Players' structure lists
Public Function allocStructLists() As Boolean
    Dim inc As Long, stat As Long

    If numStructureStats <> 0 Then
        ReDim apStructTypeLists(0 To MAX_PLAYERS - 1, 0 To numStructureStats - 1)
        apStructTypeLists_allocated = True
        For inc = 0 To MAX_PLAYERS - 1
            For stat = 0 To numStructureStats - 1
                apStructTypeLists(inc, stat) = UNAVAILABLE
            Next
        Next
    Else
        Erase apStructTypeLists
        apStructTypeLists_allocated = False
    End If

    allocStructLists = True
End Function


' release the structure lists
Public Sub freeStructureLists()
    'free the structure lists
    Erase apStructTypeLists
    apStructTypeLists_allocated = False
End Sub
