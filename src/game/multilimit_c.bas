Attribute VB_Name = "multilimit_c"
' src/multilimit.cpp, multilimit.h: the structure limits of the game's
' options: createLimitSet (the host's, from the structure stats),
' applyLimitSet (applied when the game starts; rules.js calls it),
' freeLimitSet and resetLimits; and the limits screen, class
' WzMultiLimitTitleUI (WzMultiLimitTitleUI.cls, made by
' WzMultiLimitTitleUI_new) with displayStructureBar and the screen's ids.
Option Explicit

' ////////////////////////////////////////////////////////////////////////////
' defines
Public Const IDLIMITS As Long = 22000
Public Const IDLIMITS_RETURN As Long = IDLIMITS + 1
Public Const IDLIMITS_OK As Long = IDLIMITS + 2
Public Const IDLIMITS_FORCE As Long = IDLIMITS + 3
Public Const IDLIMITS_FORCEOFF As Long = IDLIMITS + 4
Public Const IDLIMITS_ENTRIES_START As Long = IDLIMITS + 5
Public Const IDLIMITS_ENTRIES_END As Long = IDLIMITS + 99

Public Const LIMITSX As Long = 25
Public Const LIMITSY As Long = 30
Private Const LIMITSW As Long = 580
Private Const LIMITSH As Long = 430

Public Const LIMITS_OKX As Long = LIMITSW - 90
Public Const LIMITS_OKY As Long = LIMITSH - 42

Public Const LIMITS_FORCEOFFX As Long = LIMITSX + 25
Public Const LIMITS_FORCEX As Long = LIMITSX + 65

Private Const BARWIDTH As Long = 480
Private Const BARHEIGHT As Long = 40

' the FN ids of this module's lambdas and function pointers (WzFunc; cluster S, fork S1: 18020-18039)
Public Const FN_displayStructureBar As Long = 18020
Public Const FN_limitsForm_calcLayout As Long = 18021          ' LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_limitsList_calcLayout As Long = 18022          ' LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_limitsButton_onDelete As Long = 18023          ' [](WIDGET *psWidget) { delete the DisplayStructureBarCache }


' ////////////////////////////////////////////////////////////////////////////
' (VB) Public: WzMultiLimitTitleUI.cls's run calls it (static in multilimit.cpp)
Public Sub freeLimitSet()
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

' ////////////////////////////////////////////////////////////////////////////
' the limits screen (WzMultiLimitTitleUI.cls)

' std::make_shared<WzMultiLimitTitleUI>(parent)
Public Function WzMultiLimitTitleUI_new(ByVal parent As WzMultiplayerOptionsTitleUI) As WzTitleUI
    Dim t As WzMultiLimitTitleUI
    Set t = New WzMultiLimitTitleUI
    t.init parent
    Set WzMultiLimitTitleUI_new = t
End Function

' ////////////////////////////////////////////////////////////////////////////

' (a DisplayStructureBarCache {WzText wzNameText; WzText wzLimitText;} is a Variant array of the two WzTexts)
Private Sub displayStructureBar(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim cache As Variant, wzNameText As WzText, wzLimitText As WzText, x As Long, y As Long, w As Long, h As Long
    Dim stat As Long, position As Vector3i, rotationV As Vector3i, scale_ As Long, Radius As Long
    ' Any widget using displayStructureBar must have its pUserData initialized to a (DisplayStructureBarCache*)
    ' (assert(psWidget->pUserData != nullptr))
    cache = psWidget.pUserData
    Set wzNameText = cache(0)
    Set wzLimitText = cache(1)

    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    w = psWidget.width()
    h = psWidget.height()
    stat = psWidget.UserData            ' STRUCTURE_STATS *stat = asStructureStats + psWidget->UserData;

    drawBlueBox x, y, w, h

    ' draw image
    pie_SetGeometricOffset x + 35, y + psWidget.height() \ 2 + 9
    rotationV.x = -15
    rotationV.y = ((realTime \ 45) Mod 360)  '45
    rotationV.z = 0
    position.x = 0
    position.y = 0
    position.z = BUTTON_DEPTH * 2       'getStructureStatSize(stat)  * 38 * OBJECT_RADIUS;

    Radius = getStructureStatSizeMax(stat)
    If Radius <= 128 Then
        scale_ = SMALL_STRUCT_SCALE
    ElseIf Radius <= 256 Then
        scale_ = MED_STRUCT_SCALE
    Else
        scale_ = LARGE_STRUCT_SCALE
    End If

    displayStructureStatButton stat, rotationV, position, scale_

    ' draw name
    wzNameText.setText asStructureStats(stat).name_, font_regular     ' getLocalizedStatsName(stat)
    wzNameText.render x + 80, y + psWidget.height() \ 2 + 3, psPalette(WZCOL_TEXT_BRIGHT)

    ' draw limit
    wzLimitText.setText CStr(W_SLIDER_(widgGetFromID(psWScreen, psWidget.id + 1)).pos), font_regular   ' "%u"
    wzLimitText.render x + 270, y + psWidget.height() \ 2 + 3, psPalette(WZCOL_TEXT_BRIGHT)
End Sub

'--------------------------------------------------------------------------
' (VB) the lambdas (docs/CONVENTIONS.md, "Widgets")
Public Function multilimit_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET
    Select Case f.id
    Case FN_displayStructureBar
        displayStructureBar args(0), args(1), args(2)

    Case FN_limitsForm_calcLayout                   ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        psWidget.setGeometry LIMITSX, LIMITSY, LIMITSW, LIMITSH

    Case FN_limitsList_calcLayout                   ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)          ' IntListTabWidget *limitsList = static_cast<IntListTabWidget *>(psWidget);
        ListTabWidget_setChildSize psWidget, BARWIDTH, BARHEIGHT
        ListTabWidget_setChildSpacing psWidget, 5, 5
        psWidget.setGeometry 50, 10, BARWIDTH, 370

    Case FN_limitsButton_onDelete                   ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        ' (assert(psWidget->pUserData != nullptr)); delete static_cast<DisplayStructureBarCache *>(psWidget->pUserData)
        psWidget.pUserData = Empty
    End Select
End Function
