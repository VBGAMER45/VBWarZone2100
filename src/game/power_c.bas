Attribute VB_Name = "power_c"
' src/power.cpp, power.h: each player's power, the queue of power requests,
' and the extraction by the derricks linked to power generators. The 32.32
' fixed-point int64_t values are Doubles: they stay below 2^53 (MAX_POWER
' is 1000000 * 2^32 < 2^43), so sums and products of them are exact, and a
' division is Fix(a / b), which truncates as C does for |a| < 2^53. A value
' printed with PRId64 is passed as value (vsprintf prints a Decimal
' in full).
Option Explicit

' Free power on collection of oildrum.
Public Const OILDRUM_POWER As Long = 100

Private Const EXTRACT_POINTS As Long = 1
Private Const MAX_POWER As Long = 1000000

Private Const FP_ONE As Double = 4294967296#    ' ((int64_t)1 << 32)

'flag used to check for power calculations to be done or not
Public powerCalculated As Boolean

Private Type PlayerPower
    ' All fields are 32.32 fixed point.
    currentPower As Double                      ' The current amount of power available to the player.
    powerQueue_amount() As Double               ' std::vector<PowerRequest> powerQueue: PowerRequest::amount, Amount of power being requested.
    powerQueue_id() As Long                     ' PowerRequest::id (unsigned), Structure which is requesting power.
    powerQueue_size As Long
    powerModifier As Long                       ' Percentage modifier on power from each derrick.
    maxStorage As Double                        ' Maximum storage of power, in total.
    extractedPower As Double                    ' Total amount of extracted power in this game.
    wastedPower As Double                       ' Total amount of wasted power in this game.
    powerGeneratedLastUpdate As Double          ' The power generated the last time updatePlayerPower was called for this player
End Type

Private asPower(0 To MAX_PLAYERS - 1) As PlayerPower

Public Sub setPowerModifier(ByVal player As Long, ByVal modifier As Long)
    asPower(player).powerModifier = modifier
End Sub

Public Sub setPowerMaxStorage(ByVal player As Long, ByVal max As Long)
    asPower(player).maxStorage = max * FP_ONE
    If asPower(player).currentPower > asPower(player).maxStorage Then asPower(player).currentPower = asPower(player).maxStorage   ' std::min<int64_t>
End Sub

'allocate the space for the playerPower
Public Function allocPlayerPower() As Boolean
    clearPlayerPower
    powerCalculated = True
    allocPlayerPower = True
End Function

'clear the playerPower
Public Sub clearPlayerPower()
    Dim player As Long
    For player = 0 To MAX_PLAYERS - 1
        asPower(player).currentPower = 0
        asPower(player).extractedPower = 0
        asPower(player).wastedPower = 0
        asPower(player).powerModifier = 100
        asPower(player).powerQueue_size = 0
        asPower(player).maxStorage = MAX_POWER * FP_ONE
        asPower(player).powerGeneratedLastUpdate = 0
    Next
End Sub

' Returns true iff the power is available. New requests replace old ones (without losing the position in the queue).
Private Function addPowerRequest(ByVal player As Long, ByVal id As Long, ByVal amount As Double) As Boolean
    Dim requiredPower As Double, n As Long

    requiredPower = amount
    n = 0
    Do While n < asPower(player).powerQueue_size
        If asPower(player).powerQueue_id(n) = id Then Exit Do
        requiredPower = requiredPower + asPower(player).powerQueue_amount(n)
        n = n + 1
    Loop
    If n = asPower(player).powerQueue_size Then
        ' p->powerQueue.resize(n + 1)
        If n = 0 Then
            ReDim Preserve asPower(player).powerQueue_amount(0 To 7)
            ReDim Preserve asPower(player).powerQueue_id(0 To 7)
        ElseIf n > UBound(asPower(player).powerQueue_id) Then
            ReDim Preserve asPower(player).powerQueue_amount(0 To 2 * n - 1)
            ReDim Preserve asPower(player).powerQueue_id(0 To 2 * n - 1)
        End If
        asPower(player).powerQueue_size = n + 1
        asPower(player).powerQueue_id(n) = id
    End If
    asPower(player).powerQueue_amount(n) = amount
    addPowerRequest = (requiredPower <= asPower(player).currentPower)
End Function

' Removes any pending power request from this structure.
Public Sub delPowerRequest(ByVal psStruct As Long)
    Dim player As Long, n As Long, k As Long
    If Not (psStruct <> 0) Then ASSERT_FAILURE "(psStruct) != nullptr", "src/power.cpp:134", "delPowerRequest", "%s should not be nullptr", "psStruct": Exit Sub
    player = aStruct(psStruct).player

    For n = 0 To asPower(player).powerQueue_size - 1
        If asPower(player).powerQueue_id(n) = aStruct(psStruct).id Then
            ' p->powerQueue.erase(p->powerQueue.begin() + n)
            For k = n To asPower(player).powerQueue_size - 2
                asPower(player).powerQueue_amount(k) = asPower(player).powerQueue_amount(k + 1)
                asPower(player).powerQueue_id(k) = asPower(player).powerQueue_id(k + 1)
            Next
            asPower(player).powerQueue_size = asPower(player).powerQueue_size - 1
            Exit Sub
        End If
    Next
End Sub

Private Function checkPrecisePowerRequest(ByVal psStruct As Long) As Double
    Dim player As Long, requiredPower As Double, n As Long
    If Not (psStruct <> 0) Then ASSERT_FAILURE "(psStruct) != nullptr", "src/power.cpp:149", "checkPrecisePowerRequest", "%s should not be nullptr", "psStruct": checkPrecisePowerRequest = -1: Exit Function
    player = aStruct(psStruct).player

    requiredPower = 0
    For n = 0 To asPower(player).powerQueue_size - 1
        requiredPower = requiredPower + asPower(player).powerQueue_amount(n)
        If asPower(player).powerQueue_id(n) = aStruct(psStruct).id Then
            If requiredPower <= asPower(player).currentPower Then
                checkPrecisePowerRequest = -1     ' Have enough power.
                Exit Function
            End If
            checkPrecisePowerRequest = requiredPower - asPower(player).currentPower
            Exit Function
        End If
    Next

    checkPrecisePowerRequest = -1
End Function

' Checks how much power must be accumulated, before the power request from this structure can be satisfied.
' Returns -1 if there is no power request or if there is enough power already.
Public Function checkPowerRequest(ByVal psStruct As Long) As Long
    Dim power As Double
    power = checkPrecisePowerRequest(psStruct)
    If power <> -1 Then checkPowerRequest = Fix(power / FP_ONE) Else checkPowerRequest = -1
End Function

Private Function getPreciseQueuedPower(ByVal player As Long) As Double
    Dim requiredPower As Double, n As Long

    requiredPower = 0
    For n = 0 To asPower(player).powerQueue_size - 1
        requiredPower = requiredPower + asPower(player).powerQueue_amount(n)
    Next
    getPreciseQueuedPower = requiredPower
End Function

Public Function getQueuedPower(ByVal player As Long) As Long
    Dim requiredPower As Double, n As Long

    requiredPower = 0
    For n = 0 To asPower(player).powerQueue_size - 1
        requiredPower = requiredPower + asPower(player).powerQueue_amount(n)
    Next
    getQueuedPower = Fix(requiredPower / FP_ONE)
End Function

Private Sub syncDebugEconomy(ByVal player As Long, ByVal ch As Long)
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:201", "syncDebugEconomy", "Bad player (%d)", player: Exit Sub

    syncDebug "syncDebugEconomy", "%c economy%u = %lld", ch, player, asPower(player).currentPower
End Sub

Public Sub usePower(ByVal player As Long, ByVal quantity As Long)
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:208", "usePower", "Invalid player (%d)", player: Exit Sub
    syncDebug "usePower", "usePower%d %lld-=%u", player, asPower(player).currentPower, quantity
    asPower(player).currentPower = asPower(player).currentPower - U32(quantity) * FP_ONE
    If asPower(player).currentPower < 0 Then asPower(player).currentPower = 0     ' MAX(0, ...)
End Sub

Public Sub addPower(ByVal player As Long, ByVal quantity As Long)
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:215", "addPower", "Bad player (%d)", player: Exit Sub
    syncDebug "addPower", "addPower%d %lld+=%d", player, asPower(player).currentPower, quantity
    asPower(player).currentPower = asPower(player).currentPower + quantity * FP_ONE
    If asPower(player).currentPower > asPower(player).maxStorage Then
        asPower(player).wastedPower = asPower(player).wastedPower + asPower(player).currentPower - asPower(player).maxStorage
        asPower(player).currentPower = asPower(player).maxStorage
    End If
End Sub

'resets the power calc flag for all players
Public Sub powerCalc(ByVal on_ As Boolean)
    powerCalculated = on_
End Sub

' Each Resource Extractor yields EXTRACT_POINTS per second FOREVER
Private Function updateExtractedPower(ByVal psBuilding As Long) As Double
    Dim pResExtractor As Long, extractedPoints As Double

    pResExtractor = aStruct(psBuilding).pFunctionality
    extractedPoints = 0

    'only extracts points whilst its active ie associated with a power gen
    'and has got some power to extract
    If aStruct(pResExtractor).functionality.resourceExtractor.psPowerGen <> 0 Then
        ' include modifier as a %
        extractedPoints = Fix(CDbl(asPower(aStruct(psBuilding).player).powerModifier * EXTRACT_POINTS) * FP_ONE / (100 * GAME_UPDATES_PER_SEC))
        syncDebug "updateExtractedPower", "updateExtractedPower%d = %lld", aStruct(psBuilding).player, extractedPoints
    End If
    If Not (extractedPoints >= 0) Then ASSERT_FAILURE "extractedPoints >= 0", "src/power.cpp:248", "updateExtractedPower", "extracted negative amount of power"
    updateExtractedPower = extractedPoints
End Function

' returns the relevant list based on OffWorld or OnWorld: (VB) the list's
' first node (0 for an empty list), or -1 where C++ returns nullptr
Private Function powerStructList(ByVal player As Long) As Long
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:255", "powerStructList", "Invalid player %d", player: powerStructList = -1: Exit Function
    If offWorldKeepLists Then
        powerStructList = mission_apsStructLists(player).head
    Else
        powerStructList = apsStructLists(player).head
    End If
End Function

' Update current power based on what Power Generators exist
' (ticks: the C++ default is 1)
Public Sub updatePlayerPower(ByVal player As Long, Optional ByVal ticks As Long = 1)
    Dim powerBefore As Double, it As Long, psStruct As Long
    powerBefore = asPower(player).currentPower

    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:271", "updatePlayerPower", "Invalid player %d", player: Exit Sub

    syncDebugEconomy player, Asc("<")

    it = powerStructList(player)
    If it <> -1 Then
        Do While it <> 0
            psStruct = lnObj(it)
            If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_POWER_GEN And aStruct(psStruct).status = SS_BUILT Then
                updateCurrentPower psStruct, player, ticks
            End If
            it = lnNext(it)
        Loop
    End If
    syncDebug "updatePlayerPower", "updatePlayerPower%u %lld->%lld", player, powerBefore, asPower(player).currentPower
    asPower(player).powerGeneratedLastUpdate = asPower(player).currentPower - powerBefore

    syncDebugEconomy player, Asc("<")
End Sub

' Updates the current power based on the extracted power and a Power Generator
Private Sub updateCurrentPower(ByVal psStruct As Long, ByVal player As Long, ByVal ticks As Long)
    Dim psPowerGen As Long, extractedPower As Double, i As Long, extractor As Long, multiplier As Long, added As Double
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:295", "updateCurrentPower", "Invalid player %u", player: Exit Sub

    psPowerGen = aStruct(psStruct).pFunctionality
    If Not (psPowerGen <> 0) Then ASSERT_FAILURE "psPowerGen != nullptr", "src/power.cpp:298", "updateCurrentPower", "Null pFunctionality?": Exit Sub

    'each power gen can cope with its associated resource extractors
    extractedPower = 0
    For i = 0 To NUM_POWER_MODULES - 1
        extractor = aStruct(psPowerGen).functionality.powerGenerator.apResExtractors(i)
        If extractor <> 0 Then
            If aStruct(extractor).died <> 0 Then
                syncDebugStructure_ "updateCurrentPower", extractor, Asc("-")
                aStruct(psPowerGen).functionality.powerGenerator.apResExtractors(i) = 0  ' Clear pointer.
                extractor = 0
            End If
        End If
        If extractor <> 0 Then
            extractedPower = extractedPower + updateExtractedPower(extractor)
        End If
    Next

    multiplier = getBuildingPowerPoints(psStruct)
    syncDebug "updateCurrentPower", "updateCurrentPower%d = %lld,%u", player, extractedPower, multiplier

    added = Fix(extractedPower * multiplier / 100) * ticks
    asPower(player).currentPower = asPower(player).currentPower + added
    asPower(player).extractedPower = asPower(player).extractedPower + added
    If Not (asPower(player).currentPower >= 0) Then ASSERT_FAILURE "asPower[player].currentPower >= 0", "src/power.cpp:320", "updateCurrentPower", "negative power"
    If asPower(player).currentPower > asPower(player).maxStorage Then
        asPower(player).wastedPower = asPower(player).wastedPower + asPower(player).currentPower - asPower(player).maxStorage
        asPower(player).currentPower = asPower(player).maxStorage
    End If
End Sub

' Used in multiplayer to force power levels.
Public Sub setPower(ByVal player As Long, ByVal power As Long)
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:330", "setPower", "Invalid player (%u)", player: Exit Sub

    syncDebug "setPower", "setPower%d %lld->%d", player, asPower(player).currentPower, power
    asPower(player).currentPower = power * FP_ONE
    If Not (asPower(player).currentPower >= 0) Then ASSERT_FAILURE "asPower[player].currentPower >= 0", "src/power.cpp:334", "setPower", "negative power"
End Sub

' Get the amount of power current held by the given player.
Public Function getPower(ByVal player As Long) As Long
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:339", "getPower", "Invalid player (%u)", player: getPower = 0: Exit Function

    getPower = Fix(asPower(player).currentPower / FP_ONE)
End Function

Public Function getPrecisePower(ByVal player As Long) As Double
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:346", "getPrecisePower", "Invalid player (%u)", player: getPrecisePower = 0: Exit Function

    getPrecisePower = asPower(player).currentPower
End Function

' Get amount of power extracted during the whole game
Public Function getExtractedPower(ByVal player As Long) As Double
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:353", "getExtractedPower", "Invalid player (%u)", player: getExtractedPower = 0: Exit Function

    getExtractedPower = Fix(asPower(player).extractedPower / FP_ONE)
End Function

' Get amount of power wasted during the whole game
Public Function getWastedPower(ByVal player As Long) As Double
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/power.cpp:360", "getWastedPower", "Invalid player (%u)", player: getWastedPower = 0: Exit Function

    getWastedPower = Fix(asPower(player).wastedPower / FP_ONE)
End Function

Public Function getPowerMinusQueued(ByVal player As Long) As Long
    If Not ULT(player, MAX_PLAYERS) Then
        getPowerMinusQueued = 0
        Exit Function
    End If

    getPowerMinusQueued = Fix((asPower(player).currentPower - getPreciseQueuedPower(player)) / FP_ONE)
End Function

' Get the approximate power generated per second for the specified player (for display purposes - not to be used for calculations)
Public Function getApproxPowerGeneratedPerSecForDisplay(ByVal player As Long) As String
    Dim floatingValue As Double
    If Not ULT(player, MAX_PLAYERS) Then
        getApproxPowerGeneratedPerSecForDisplay = "-"
        Exit Function
    End If

    floatingValue = (asPower(player).powerGeneratedLastUpdate / FP_ONE) * CDbl(GAME_UPDATES_PER_SEC)
    getApproxPowerGeneratedPerSecForDisplay = va("%+.0f", floatingValue)     ' fmt::format("{:+.0f}", ...)
End Function

Public Function requestPowerFor(ByVal psStruct As Long, ByVal amount As Long) As Boolean
    requestPowerFor = requestPrecisePowerFor(psStruct, amount * FP_ONE)
End Function

Public Function requestPrecisePowerFor(ByVal psStruct As Long, ByVal amount As Double) As Boolean
    Dim haveEnoughPower As Boolean
    If amount <= 0 Or Not powerCalculated Then
        requestPrecisePowerFor = True
        Exit Function
    End If

    haveEnoughPower = addPowerRequest(aStruct(psStruct).player, aStruct(psStruct).id, amount)
    If haveEnoughPower Then
        ' you can have it
        asPower(aStruct(psStruct).player).currentPower = asPower(aStruct(psStruct).player).currentPower - amount
        delPowerRequest psStruct
        syncDebug "requestPrecisePowerFor", "requestPrecisePowerFor%d,%u amount%lld", aStruct(psStruct).player, aStruct(psStruct).id, amount
        requestPrecisePowerFor = True
        Exit Function
    End If
    syncDebug "requestPrecisePowerFor", "requestPrecisePowerFor%d,%u wait,amount%lld", aStruct(psStruct).player, aStruct(psStruct).id, amount
    requestPrecisePowerFor = False                  ' Not enough power in the queue.
End Function
