Attribute VB_Name = "multilimit_c"
' src/multilimit.cpp, multilimit.h: so far the structure limits of the
' game's options: createLimitSet (the host's, from the structure stats),
' applyLimitSet (applied when the game starts; rules.js calls it),
' freeLimitSet and resetLimits. Not ported yet: the limits screen
' (WzMultiLimitTitleUI, displayStructureBar): cluster F.
Option Explicit

' ////////////////////////////////////////////////////////////////////////////
Private Sub freeLimitSet()
    ' Free the old set if required
    Erase ingame.structureLimits
    ingame.structureLimitsCount = 0
End Sub

' ////////////////////////////////////////////////////////////////////////////
Public Sub createLimitSet()
    Dim numchanges As Long, i As Long
    numchanges = 0                      ' UDWORD

    debug_ 331, LOG_NET, "createLimitSet", "LimitSet created"
    ' free old limiter structure
    freeLimitSet

    ' don't bother creating if a challenge mode is active
    If challengeActive Then
        Exit Sub
    End If

    ' Count the number of changes
    For i = 0 To numStructureStats - 1
        ' If the limit differs from the default
        If asStructureStats(i).upgrade(0).limit <> LOTS_OF Then
            numchanges = numchanges + 1
        End If
    Next

    If numchanges > 0 Then
        ' Allocate some memory for the changes
        ReDim ingame.structureLimits(0 To numchanges - 1)  ' ingame.structureLimits.reserve(numchanges)

        ' Prepare chunk
        For i = 0 To numStructureStats - 1
            If asStructureStats(i).upgrade(0).limit <> LOTS_OF Then
                If Not (ingame.structureLimitsCount < numchanges) Then ASSERT_FAILURE "ingame.structureLimits.size() < numchanges", "src/multilimit.cpp:361", "createLimitSet", "Bad number of changed limits": Exit Sub
                ingame.structureLimits(ingame.structureLimitsCount).id = i
                ingame.structureLimits(ingame.structureLimitsCount).limit = asStructureStats(i).upgrade(0).limit
                ingame.structureLimitsCount = ingame.structureLimitsCount + 1
            End If
        Next
    End If

    If NetPlay_isHost Then
        updateStructureDisabledFlags
        sendOptions

        NETsetLobbyLimitFlags ingame.flags
    End If
End Sub

' ////////////////////////////////////////////////////////////////////////////
Public Function applyLimitSet() As Boolean
    Dim flags As Long, k As Long, id As Long, player As Long, it As Long, nxt As Long, psStruct As Long, i As Long

    If ingame.structureLimitsCount = 0 Then
        applyLimitSet = False
        Exit Function
    End If

    ' Get the limits and decode
    flags = ingame.flags And MPFLAGS_FORCELIMITS
    For k = 0 To ingame.structureLimitsCount - 1
        id = ingame.structureLimits(k).id

        ' So long as the ID is valid
        If ULT(id, numStructureStats) Then
            For player = 0 To MAX_PLAYERS - 1
                asStructureStats(id).upgrade(player).limit = ingame.structureLimits(k).limit

                If (ingame.flags And MPFLAGS_FORCELIMITS) <> 0 Then
                    Do While ULT(asStructureStats(id).upgrade(player).limit, asStructureStats(id).curCount(player))
                        ' mutating_list_iterate
                        it = apsStructLists(player).head
                        Do While it <> 0
                            nxt = lnNext(it)
                            psStruct = lnObj(it)
                            If asStructureStats(aStruct(psStruct).pStructureType).typ = asStructureStats(id).typ Then
                                removeStruct psStruct, True
                                Exit Do                 ' IterationResult::BREAK_ITERATION
                            End If
                            it = nxt                    ' IterationResult::CONTINUE_ITERATION
                        Loop
                    Loop
                End If
            Next
            If asStructureStats(id).typ = REF_VTOL_FACTORY And ingame.structureLimits(k).limit = 0 Then
                flags = flags Or MPFLAGS_NO_VTOLS
            End If
            If asStructureStats(id).typ = REF_CYBORG_FACTORY And ingame.structureLimits(k).limit = 0 Then
                flags = flags Or MPFLAGS_NO_CYBORGS
            End If
            If asStructureStats(id).typ = REF_LASSAT And ingame.structureLimits(k).limit = 0 Then
                flags = flags Or MPFLAGS_NO_LASSAT
            End If
            If asStructureStats(id).typ = REF_SAT_UPLINK And ingame.structureLimits(k).limit = 0 Then
                flags = flags Or MPFLAGS_NO_UPLINK
            End If
        End If
    Next

    ' If structure limits are enforced and VTOLs are disabled, disable also AA structures
    If (flags And (MPFLAGS_FORCELIMITS Or MPFLAGS_NO_VTOLS)) = (MPFLAGS_FORCELIMITS Or MPFLAGS_NO_VTOLS) Then
        For i = 0 To numStructureStats - 1
            If asStructureStats(i).numWeaps > 0 Then
                If asWeaponStats(asStructureStats(i).psWeapStat(0)).surfaceToAir = SHOOT_IN_AIR Then
                    For player = 0 To MAX_PLAYERS - 1
                        asStructureStats(i).upgrade(player).limit = 0
                    Next
                End If
            End If
        Next
    End If
    ' Disable related research according to disabledWhen in research.json
    If (flags And MPFLAGS_FORCELIMITS) <> 0 Then
        RecursivelyDisableResearchByFlags flags And Not (MPFLAGS_FORCELIMITS Or MPFLAGS_NO_UPLINK)
    End If
    '
    '  Special rules for LasSat and Uplink which do not produce units and have no upgrades:
    '  - research can be disabled even if limits are not enforced
    '  - because of the research path, only disable Uplink if also LasSat is disabled
    '
    If (flags And MPFLAGS_NO_LASSAT) <> 0 Then
        RecursivelyDisableResearchByFlags MPFLAGS_NO_LASSAT
    End If
    If (flags And (MPFLAGS_NO_LASSAT Or MPFLAGS_NO_UPLINK)) = (MPFLAGS_NO_LASSAT Or MPFLAGS_NO_UPLINK) Then
        RecursivelyDisableResearchByFlags MPFLAGS_NO_UPLINK
    End If

    ' A bit of a hack:
    ' - We have to retain the `structureLimits` for saving in the replay, but the process of actually saving the replay init data
    '   happens *after* the scripts initialize and are sent eventGameInit, which calls applyLimitSet() - this function - which then
    '   calls `freeLimitSet` below (which clears ingame.structureLimits)
    ' - While this should be refactored at some point... for now, just stash `structureLimits` in `lastAppliedStructureLimits`
    '   before clearing it, and then the replay init data can save `ingame.lastAppliedStructureLimits`
    ReDim ingame.lastAppliedStructureLimits(0 To ingame.structureLimitsCount - 1)
    For k = 0 To ingame.structureLimitsCount - 1
        ingame.lastAppliedStructureLimits(k) = ingame.structureLimits(k)
    Next
    ingame.lastAppliedStructureLimitsCount = ingame.structureLimitsCount

    freeLimitSet
    applyLimitSet = True
End Function

Public Sub resetLimits()
    Dim i As Long
    For i = 0 To numStructureStats - 1
        asStructureStats(i).upgrade(0).limit = asStructureStats(i).base.limit
    Next
End Sub
