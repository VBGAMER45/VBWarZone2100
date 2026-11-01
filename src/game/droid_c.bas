Attribute VB_Name = "droid_c"
' src/droid.cpp, droid.h: what the template loader uses
' (getDroidResourceName, droidTemplateType), the droid container
' (GlobalDroidContainer), DROID's constructor and destructor, its names,
' type tests and component getters, droid.h's target setters,
' nextModuleToBuild, the sums of a design's stats (weight, body, speed,
' build points, power), recycled experience, reallyBuildDroid,
' droidSetBits, and what movement uses: removeDroidBase, destroyDroid,
' vanishDroid, the droid sync log line (syncDebugDroid_), the commander
' range, calcDroidSpeed, the levels and experience, isRetreatingForRepair,
' isReturningToBase, droidOnMap and checkDroid (CHECK_DROID). For the
' actions and orders: cancelBuild, recycleDroid, droidRemove (with
' droidCancelRepairers), droidUpdate (with droidBodyUpgrade and
' templateSetParts), building, demolishing, restoring and repairing
' (droidStartBuild, droidUpdateBuild, droidUpdateDemolishing,
' droidUpdateRestore, droidUpdateRepair, droidUpdateDroidRepair, the repair
' counts and droidWasFullyRepaired), calcDroidMuzzleBaseLocation,
' the tile pickers (noDroid, zonedPAT, pickATileGen, pickATileGenThreat,
' pickHalfATile), checkDroidsBuilding, checkDroidsDemolishing,
' setUpBuildModule, electronicDroid, the VTOL ammunition tests, the sensor
' tests, giftSingleDroid, droidResistance, the selection (SelectDroid, SelectGroupDroid,
' DeSelectDroid), droidSetPosition and droidSqDist, addDroidDeathAnimationEffect. For the structures'
' update: buildDroid (initialOrders != nullptr is hasInitialOrders),
' vtolReadyToRearm and fillVtolDroid. The AUDIO_CALLBACKs
' droidCheckBuildStillInProgress and droidBuildStartAudioCallback are
' numbers (CB_...), as move_c's. pickATileGen's function pointer is a
' number too (PAT_zonedPAT, PAT_canFitDroid); its Vector2i overload is
' pickATileGen_pos.
' calcSum's and calcUpgradeSum's lambdas are a selector
' (SUM_WEIGHT, ...); std::priority_queue<int> recycled_experience[] is a
' list per player kept in descending order. The rest of the droids comes
' later in cluster B.
' The groups (cluster G, at the end): assignObjectToGroup, removeObjectFromGroup,
' activateGroupAndMove, activateNoGroup, activateGroup and the
' groupConsoleInformOf* messages (ngettext is a Private copy: no translations).
Option Explicit

Private Const DROID_EXPLOSION_VELOCITY_FACTOR As Single = 0.9

Public Const OFF_SCREEN As Long = 9999      ' world->screen check - alex
Public Const REPAIRLEV_LOW As Long = 50     ' percentage of body points remaining at which to repair droid automatically.
Public Const REPAIRLEV_HIGH As Long = 75    ' ditto, but this will repair much sooner..
Public Const DROID_RESISTANCE_FACTOR As Long = 30
Public Const MAX_RECYCLED_DROIDS As Long = 450
Public Const TOO_NEAR_EDGE As Long = 3
' droid.h
Public Const MIN_WEAPON_DAMAGE As Long = 1

Public Const EXP_REDUCE_DAMAGE As Long = 6  ' damage of a droid is reduced by this value per experience level, in %
Private Const UNIT_LOST_DELAY As Long = 5 * GAME_TICKS_PER_SEC

Public Const EXP_ACCURACY_BONUS As Long = 5 ' accuracy of a droid is increased by this value per experience level, in %
Public Const EXP_SPEED_BONUS As Long = 5    ' speed of a droid is increased by this value per experience level, in %

' Height the transporter hovers at above the terrain.
Private Const TRANSPORTER_HOVER_HEIGHT As Long = 10

' store the experience of recently recycled droids
Private recycled_experience() As Long   ' (player, i), descending
Private recycled_experience_size(0 To MAX_PLAYERS - 1) As Long
Private recycled_experience_cap As Long

' the structure that was last hit
Public psLastDroidHit As Long

' droid.h
Public Type INITIAL_DROID_ORDERS
    secondaryOrder As Long              ' uint32_t
    moveToX As Long                     ' int32_t
    moveToY As Long
    factoryId As Long                   ' uint32_t
End Type

' (VB) DroidStartBuild_: the function droidStartBuild keeps the name
Public Enum DroidStartBuild_
    DroidStartBuildFailed
    DroidStartBuildSuccess
    DroidStartBuildPending
End Enum

' (VB) the AUDIO_CALLBACKs of this module (numbers; SoundGame.cls calls them by number)
Public Const CB_droidAudioTrackStopped As Long = 1     ' the one systemInitialise passes to audio_Init
Public Const CB_droidCheckBuildStillInProgress As Long = 4
Public Const CB_droidBuildStartAudioCallback As Long = 5

' (VB) calcSum's and calcUpgradeSum's lambdas: what they sum
Private Const SUM_WEIGHT As Long = 0
Private Const SUM_BUILDPOINTS As Long = 1
Private Const SUM_BUILDPOWER As Long = 2
Private Const SUM_HITPOINTS As Long = 3
Private Const SUM_HITPOINTPCT As Long = 4

' droid.h
Public Enum PICKTILE
    NO_FREE_TILE
    FREE_TILE
End Enum

' (VB) pickATileGen's function pointers bool (*function)(UDWORD x, UDWORD y): numbers
Public Const PAT_zonedPAT As Long = 0
Private Const PAT_canFitDroid As Long = 1

' the string resource of the name, vbNullString when there is none (C: nullptr)
Public Function getDroidResourceName(ByVal pName As String) As String
    ' See if the name has a string resource associated with it by trying
    ' to get the string resource.
    getDroidResourceName = strresGetString(psStringRes, pName)
End Function

' (VB) GlobalDroidContainer().emplace(id, player)
Public Function GlobalDroidContainer_emplace(ByVal id As Long, ByVal player As Long) As Long
    Dim psDroid As Long
    psDroid = DroidContainer_alloc()
    DROID_construct_ psDroid, id, player
    GlobalDroidContainer_emplace = psDroid
End Function

' DROID::DROID(uint32_t id, unsigned player), with BASE_OBJECT's and SIMPLE_OBJECT's constructors
Public Sub DROID_construct_(ByVal this As Long, ByVal id As Long, ByVal player As Long)
    Dim vPlayer As Long, i As Long
    ' SIMPLE_OBJECT(type, id, player)
    aDroid(this).typ = OBJ_DROID
    aDroid(this).id = id
    aDroid(this).player = player And &HFF&
    aDroid(this).born = gameTime
    aDroid(this).died = 0
    aDroid(this).time_ = 0
    ' BASE_OBJECT(type, id, player)
    aDroid(this).repairGroup = UBYTE_MAX
    aDroid(this).selected = 0
    aDroid(this).lastEmission = 0
    aDroid(this).lastHitWeapon = WSC_NUM_WEAPON_SUBCLASSES  ' No such weapon.
    aDroid(this).timeLastHit = UDWORD_MAX
    aDroid(this).body = 0
    aDroid(this).periodicalDamageStart = 0
    aDroid(this).periodicalDamage = 0
    aDroid(this).timeAnimationStarted = 0
    aDroid(this).animationEvent = ANIM_EVENT_NONE
    For i = 0 To MAX_PLAYERS - 1
        aDroid(this).visible(i) = 0
    Next
    aDroid(this).sDisplay.imd = 0
    aDroid(this).sDisplay.frameNumber = 0
    aDroid(this).sDisplay.screenX = 0
    aDroid(this).sDisplay.screenY = 0
    aDroid(this).sDisplay.screenR = 0
    ' DROID
    aDroid(this).droidType = DROID_ANY
    aDroid(this).psGroup = 0
    aDroid(this).secondaryOrder = DSS_ARANGE_LONG Or DSS_REPLEV_NEVER Or DSS_ALEV_ALWAYS Or DSS_HALT_GUARD
    aDroid(this).secondaryOrderPending = DSS_ARANGE_LONG Or DSS_REPLEV_NEVER Or DSS_ALEV_ALWAYS Or DSS_HALT_GUARD
    aDroid(this).secondaryOrderPendingCount = 0
    aDroid(this).action = DACTION_NONE
    aDroid(this).actionPos.x = 0
    aDroid(this).actionPos.y = 0
    aDroid(this).heightAboveMap = 0
    aDroid(this).underRepair = 0
    DroidOrder_construct aDroid(this).order     ' (the member's default constructor)

    aDroid(this).aName = ""
    For i = 0 To DROID_MAXCOMP - 1
        aDroid(this).asBits(i) = 0
    Next
    aDroid(this).pos = Vector3i_(0, 0, 0)
    aDroid(this).rot = Rotation_(0, 0, 0)
    aDroid(this).order.typ = DORDER_NONE
    aDroid(this).order.pos = Vector2i_(0, 0)
    aDroid(this).order.pos2 = Vector2i_(0, 0)
    aDroid(this).order.direction = 0
    aDroid(this).order.psObj = 0
    aDroid(this).order.psStats = -1
    aDroid(this).sMove.Status = MOVEINACTIVE
    aDroid(this).listSize = 0
    aDroid(this).listPendingBegin = 0
    aDroid(this).iAudioID = NO_SOUND
    aDroid(this).group = UBYTE_MAX
    aDroid(this).psBaseStruct = 0
    aDroid(this).sDisplay.frameNumber = 0   ' it was never drawn before
    For vPlayer = 0 To MAX_PLAYERS - 1
        If hasSharedVision(vPlayer, player) Then aDroid(this).visible(vPlayer) = UINT8_MAX Else aDroid(this).visible(vPlayer) = 0
    Next
    For i = 0 To MAX_PLAYERS - 1
        aDroid(this).seenThisTick(i) = 0
    Next
    aDroid(this).periodicalDamageStart = 0
    aDroid(this).periodicalDamage = 0
    aDroid(this).sDisplay.screenX = OFF_SCREEN
    aDroid(this).sDisplay.screenY = OFF_SCREEN
    aDroid(this).sDisplay.screenR = 0
    aDroid(this).sDisplay.imd = 0
    aDroid(this).illumination = UBYTE_MAX
    aDroid(this).resistance = ACTION_START_TIME ' init the resistance to indicate no EW performed on this droid
    aDroid(this).lastFrustratedTime = 0         ' make sure we do not start the game frustrated
End Sub

' DROID::~DROID: release all resources associated with a droid -
' should only be called by objmem - use vanishDroid preferably
' (with ~BASE_OBJECT and ~SIMPLE_OBJECT after it)
Public Sub DROID_destruct(ByVal this As Long)
    Dim psDroid As Long, it As Long, nx As Long, psCurr As Long
    psDroid = this

    If DROID_isTransporter(psDroid) Then
        If aDroid(psDroid).psGroup <> 0 Then
            'free all droids associated with this Transporter
            ' (mutating_list_iterate: each erased droid removes itself from psList)
            it = aGroup(aDroid(psDroid).psGroup).psList.head
            Do While it <> 0
                nx = lnNext(it)
                psCurr = lnObj(it)
                If psCurr = psDroid Then
                    Exit Do
                End If
                ' This will cause each droid to self-remove from `psGroup->psList`.
                If Not (DroidContainer_find(psCurr)) Then ASSERT_FAILURE "it != droidContainer.end()", "src/droid.cpp:463", "~DROID", "Droid not found in the global container!"
                DroidContainer_erase psCurr
                it = nx
            Loop
        End If
    End If

    fpathRemoveDroidData aDroid(psDroid).id

    ' leave the current formation if any
    If aDroid(psDroid).sMove.psFormation <> 0 Then
        formationLeave aDroid(psDroid).sMove.psFormation, psDroid
        aDroid(psDroid).sMove.psFormation = 0
    End If

    ' leave the current group if any
    If aDroid(psDroid).psGroup <> 0 Then
        DROID_GROUP_remove aDroid(psDroid).psGroup, psDroid
    End If

    ' ~BASE_OBJECT
    visRemoveVisibility this
    ' (the members: OrderList, asPath, watchedTiles)
    Erase aDroid(this).asOrderList
    aDroid(this).asOrderList_size = 0
    Erase aDroid(this).sMove.asPath
    aDroid(this).sMove.asPath_size = 0
    Erase aDroid(this).watchedTiles
    aDroid(this).watchedTiles_size = 0
    ' ~SIMPLE_OBJECT
    aDroid(this).typ = aDroid(this).typ + 1000000000    ' Hopefully this will trigger an assert              if someone uses the freed object.
    aDroid(this).player = (aDroid(this).player + 100) And &HFF&  ' Hopefully this will trigger an assert and/or crash if someone uses the freed object.
End Sub

Public Function droidGetName(ByVal psDroid As Long) As String
    If Not (psDroid <> 0) Then ASSERT_FAILURE "(psDroid) != nullptr", "src/droid.cpp:2624", "droidGetName", "%s should not be nullptr", "psDroid": droidGetName = "": Exit Function
    droidGetName = aDroid(psDroid).aName
End Function

'
' Set the name of a droid in it's DROID structure.
'
' - only possible on the PC where you can adjust the names,
'
Public Sub droidSetName(ByVal psDroid As Long, ByVal pName As String)
    aDroid(psDroid).aName = Left$(pName, MAX_STR_LENGTH - 1)
End Sub

Private Function isTransporter_type(ByVal typ As DROID_TYPE) As Boolean
    isTransporter_type = typ = DROID_TRANSPORTER Or typ = DROID_SUPERTRANSPORTER
End Function

Public Function DROID_isTransporter(ByVal this As Long) As Boolean
    DROID_isTransporter = isTransporter_type(aDroid(this).droidType)
End Function

' bool isTransporter(DROID_TEMPLATE const *psTemplate)
Public Function isTransporter_template(psTemplate As DROID_TEMPLATE) As Boolean
    isTransporter_template = isTransporter_type(psTemplate.droidType)
End Function

Public Function DROID_isCyborg(ByVal this As Long) As Boolean
    Select Case aDroid(this).droidType
    Case DROID_CYBORG, DROID_CYBORG_CONSTRUCT, DROID_CYBORG_REPAIR, DROID_CYBORG_SUPER
        DROID_isCyborg = True
    End Select
End Function

Public Function DROID_isConstructionDroid(ByVal this As Long) As Boolean
    DROID_isConstructionDroid = aDroid(this).droidType = DROID_CONSTRUCT Or aDroid(this).droidType = DROID_CYBORG_CONSTRUCT
End Function

' droid.h's setDroidTarget (_setDroidTarget)
Public Sub setDroidTarget(ByVal psDroid As Long, ByVal psNewTarget As Long)
    aDroid(psDroid).order.psObj = psNewTarget
    If psNewTarget <> 0 Then
        If Not (SIMPLE_OBJECT_died(psNewTarget) = 0) Then ASSERT_FAILURE "psNewTarget == nullptr || !psNewTarget->died", "src/droid.h:332", "_setDroidTarget", "setDroidTarget: Set dead target"
        If Not (SIMPLE_OBJECT_died(psNewTarget) = 0 Or (SIMPLE_OBJECT_died(psNewTarget) = NOT_CURRENT_LIST And aDroid(psDroid).died = NOT_CURRENT_LIST)) Then
            ASSERT_FAILURE "psNewTarget == nullptr || !psNewTarget->died || (psNewTarget->died == 1 && psDroid->died == 1)", "src/droid.h:333", "_setDroidTarget", "setDroidTarget: Set dead target"
        End If
    End If
End Sub

' droid.h's setDroidActionTarget (_setDroidActionTarget)
Public Sub setDroidActionTarget(ByVal psDroid As Long, ByVal psNewTarget As Long, ByVal idx As Long)
    aDroid(psDroid).psActionTarget(idx) = psNewTarget
    If psNewTarget <> 0 Then
        If Not (SIMPLE_OBJECT_died(psNewTarget) = 0 Or (SIMPLE_OBJECT_died(psNewTarget) = NOT_CURRENT_LIST And aDroid(psDroid).died = NOT_CURRENT_LIST)) Then
            ASSERT_FAILURE "psNewTarget == nullptr || !psNewTarget->died || (psNewTarget->died == 1 && psDroid->died == 1)", "src/droid.h:349", "_setDroidActionTarget", "setDroidActionTarget: Set dead target"
        End If
    End If
End Sub

' droid.h's setDroidBase (_setDroidBase)
Public Sub setDroidBase(ByVal psDroid As Long, ByVal psNewBase As Long)
    aDroid(psDroid).psBaseStruct = psNewBase
    If psNewBase <> 0 Then
        If Not (aStruct(psNewBase).died = 0) Then ASSERT_FAILURE "psNewBase == nullptr || !psNewBase->died", "src/droid.h:365", "_setDroidBase", "setDroidBase: Set dead target"
    End If
End Sub

Public Function droidTemplateType(psTemplate As DROID_TEMPLATE) As DROID_TYPE
    Dim typ As DROID_TYPE
    typ = DROID_DEFAULT

    If psTemplate.droidType = DROID_PERSON Or _
       psTemplate.droidType = DROID_CYBORG Or _
       psTemplate.droidType = DROID_CYBORG_SUPER Or _
       psTemplate.droidType = DROID_CYBORG_CONSTRUCT Or _
       psTemplate.droidType = DROID_CYBORG_REPAIR Or _
       psTemplate.droidType = DROID_TRANSPORTER Or _
       psTemplate.droidType = DROID_SUPERTRANSPORTER Then
        typ = psTemplate.droidType
    ElseIf psTemplate.asParts(COMP_BRAIN) <> 0 Then
        typ = DROID_COMMAND
    ElseIf asSensorStats(DROID_TEMPLATE_getSensorStats(psTemplate)).location = LOC_TURRET Then
        typ = DROID_SENSOR
    ElseIf asECMStats(DROID_TEMPLATE_getECMStats(psTemplate)).location = LOC_TURRET Then
        typ = DROID_ECM
    ElseIf psTemplate.asParts(COMP_CONSTRUCT) <> 0 Then
        typ = DROID_CONSTRUCT
    ElseIf asRepairStats(DROID_TEMPLATE_getRepairStats(psTemplate)).location = LOC_TURRET Then
        typ = DROID_REPAIR
    ElseIf psTemplate.asWeaps(0) <> 0 Then
        typ = DROID_WEAPON
    ' with more than weapon is still a DROID_WEAPON
    ElseIf psTemplate.numWeaps > 1 Then
        typ = DROID_WEAPON
    End If

    droidTemplateType = typ
End Function

' droid.h: True iff object is a droid.
Public Function isDroid(ByVal psObject As Long) As Boolean
    If psObject <> 0 Then isDroid = SIMPLE_OBJECT_type(psObject) = OBJ_DROID
End Function

' droid.h: Returns DROID * if droid or NULL if not.
Public Function castDroid(ByVal psObject As Long) As Long
    If isDroid(psObject) Then castDroid = psObject
End Function

Public Function nextModuleToBuild(ByVal psStruct As Long, ByVal lastOrderedModule As Long) As Long
    Dim order As Long, i As Long, nxt As Long, maxModules As Long

    If psStruct = 0 Then ASSERT_FAILURE "psStruct != nullptr && psStruct->pStructureType != nullptr", "src/droid.cpp:2903", "nextModuleToBuild", "Invalid structure pointer": Exit Function
    If Not (aStruct(psStruct).pStructureType <> -1) Then ASSERT_FAILURE "psStruct != nullptr && psStruct->pStructureType != nullptr", "src/droid.cpp:2903", "nextModuleToBuild", "Invalid structure pointer": Exit Function

    If aStruct(psStruct).status = SS_BUILT Then nxt = 1 Else nxt = 0  ' If complete, next is one after the current number of modules, otherwise next is the one we're working on.
    Select Case asStructureStats(aStruct(psStruct).pStructureType).typ
    Case REF_POWER_GEN
        'check room for one more!
        maxModules = MAX_(aStruct(psStruct).capacity + nxt, lastOrderedModule + 1)
        If maxModules <= 1 Then
            i = powerModuleStat
            order = maxModules
        End If
    Case REF_FACTORY, REF_VTOL_FACTORY
        'check room for one more!
        maxModules = MAX_(aStruct(psStruct).capacity + nxt, lastOrderedModule + 1)
        If maxModules <= NUM_FACTORY_MODULES Then
            i = factoryModuleStat
            order = maxModules
        End If
    Case REF_RESEARCH
        'check room for one more!
        maxModules = MAX_(aStruct(psStruct).capacity + nxt, lastOrderedModule + 1)
        If maxModules <= 1 Then
            i = researchModuleStat
            order = maxModules  ' Research modules are weird. Build one, get three free.
        End If
    Case Else
        'no other structures can have modules attached
    End Select

    If order <> 0 Then
        ' Check availability of Module
        If Not ULT(i, numStructureStats) Then
            order = 0
        ElseIf apStructTypeLists(aStruct(psStruct).player, i) <> AVAILABLE Then
            order = 0
        End If
    End If

    nextModuleToBuild = order
End Function

' ---- the droids' stats

' DROID::getBodyStats() .. getConstructStats(): the indices of its components' stats
Public Function DROID_getBodyStats(ByVal this As Long) As Long
    DROID_getBodyStats = aDroid(this).asBits(COMP_BODY)
End Function

Public Function DROID_getBrainStats(ByVal this As Long) As Long
    DROID_getBrainStats = aDroid(this).asBits(COMP_BRAIN)
End Function

Public Function DROID_getPropulsionStats(ByVal this As Long) As Long
    DROID_getPropulsionStats = aDroid(this).asBits(COMP_PROPULSION)
End Function

Public Function DROID_getSensorStats(ByVal this As Long) As Long
    DROID_getSensorStats = aDroid(this).asBits(COMP_SENSOR)
End Function

Public Function DROID_getECMStats(ByVal this As Long) As Long
    DROID_getECMStats = aDroid(this).asBits(COMP_ECM)
End Function

Public Function DROID_getRepairStats(ByVal this As Long) As Long
    DROID_getRepairStats = aDroid(this).asBits(COMP_REPAIRUNIT)
End Function

Public Function DROID_getConstructStats(ByVal this As Long) As Long
    DROID_getConstructStats = aDroid(this).asBits(COMP_CONSTRUCT)
End Function

Public Function DROID_isFlightBasedTransporter(ByVal this As Long) As Boolean
    If asPropulsionStats(DROID_getPropulsionStats(this)).propulsionType = PROPULSION_TYPE_LIFT Then
        DROID_isFlightBasedTransporter = DROID_isTransporter(this)
    End If
End Function

'access functions for vtols
Public Function DROID_isVtol(ByVal this As Long) As Boolean
    If asPropulsionStats(DROID_getPropulsionStats(this)).propulsionType = PROPULSION_TYPE_LIFT Then
        DROID_isVtol = Not DROID_isTransporter(this)
    End If
End Function

' returns true if the droid has lift propulsion and is moving
Public Function DROID_isFlying(ByVal this As Long) As Boolean
    If asPropulsionStats(DROID_getPropulsionStats(this)).propulsionType = PROPULSION_TYPE_LIFT Then
        DROID_isFlying = aDroid(this).sMove.Status <> MOVEINACTIVE
    End If
End Function

' (VB) func(stat): the COMPONENT_STATS field of a component (type comp, index i)
Private Function componentField(ByVal comp As COMPONENT_TYPE, ByVal i As Long, ByVal what As Long) As Long
    Dim h As Long
    Select Case comp
    Case COMP_BRAIN: h = STAT_BRAIN + i
    Case COMP_SENSOR: h = STAT_SENSOR + i
    Case COMP_ECM: h = STAT_ECM + i
    Case COMP_REPAIRUNIT: h = STAT_REPAIR + i
    Case COMP_CONSTRUCT: h = STAT_CONSTRUCT + i
    Case COMP_WEAPON: h = STAT_WEAPON + i
    End Select
    Select Case what
    Case SUM_WEIGHT: componentField = COMPONENT_STATS_weight(h)
    Case SUM_BUILDPOINTS: componentField = COMPONENT_STATS_buildPoints(h)
    Case SUM_BUILDPOWER: componentField = COMPONENT_STATS_buildPower(h)
    End Select
End Function

' (VB) func(upgrade): the COMPONENT_STATS::UPGRADE field of a component (type comp, index i) for player
Private Function componentUpgradeField(ByVal comp As COMPONENT_TYPE, ByVal i As Long, ByVal player As Long, ByVal what As Long) As Long
    Dim hp As Long, pct As Long
    Select Case comp
    Case COMP_BRAIN: hp = asBrainStats(i).upgrade(player).hitpoints: pct = asBrainStats(i).upgrade(player).hitpointPct
    Case COMP_SENSOR: hp = asSensorStats(i).upgrade(player).hitpoints: pct = asSensorStats(i).upgrade(player).hitpointPct
    Case COMP_ECM: hp = asECMStats(i).upgrade(player).hitpoints: pct = asECMStats(i).upgrade(player).hitpointPct
    Case COMP_REPAIRUNIT: hp = asRepairStats(i).upgrade(player).hitpoints: pct = asRepairStats(i).upgrade(player).hitpointPct
    Case COMP_CONSTRUCT: hp = asConstructStats(i).upgrade(player).hitpoints: pct = asConstructStats(i).upgrade(player).hitpointPct
    Case COMP_WEAPON: hp = asWeaponStats(i).upgrade(player).hitpoints: pct = asWeaponStats(i).upgrade(player).hitpointPct
    End Select
    If what = SUM_HITPOINTS Then componentUpgradeField = hp Else componentUpgradeField = pct - 100
End Function

' calcSum(asParts, numWeaps, asWeaps, func, propulsionFunc) (unsigned)
Private Function calcSum(asParts() As Long, ByVal numWeaps As Long, asWeaps() As Long, ByVal what As Long) As Long
    Dim sum As Long, i As Long, body As Long, prop As Long, propPct As Long
    body = asParts(COMP_BODY)
    prop = asParts(COMP_PROPULSION)
    Select Case what
    Case SUM_WEIGHT
        ' Propulsion weight is a percentage of the body weight.
        propPct = FromU32(Int(U32(MUL32(asBodyStats(body).weight, 100 + asPropulsionStats(prop).weight)) / 100))
    Case SUM_BUILDPOINTS
        ' Propulsion power points are a percentage of the body's build points.
        propPct = FromU32(Int(U32(MUL32(asBodyStats(body).buildPoints, 100 + asPropulsionStats(prop).buildPoints)) / 100))
    Case SUM_BUILDPOWER
        ' Propulsion power points are a percentage of the body's power points.
        propPct = FromU32(Int(U32(MUL32(asBodyStats(body).buildPower, 100 + asPropulsionStats(prop).buildPower)) / 100))
    End Select
    sum = componentField(COMP_BRAIN, asParts(COMP_BRAIN), what) + _
          componentField(COMP_SENSOR, asParts(COMP_SENSOR), what) + _
          componentField(COMP_ECM, asParts(COMP_ECM), what) + _
          componentField(COMP_REPAIRUNIT, asParts(COMP_REPAIRUNIT), what) + _
          componentField(COMP_CONSTRUCT, asParts(COMP_CONSTRUCT), what) + _
          propPct
    For i = 0 To numWeaps - 1
        sum = sum + componentField(COMP_WEAPON, asWeaps(i), what)
    Next
    calcSum = sum
End Function

' calcUpgradeSum(asParts, numWeaps, asWeaps, player, func, propulsionFunc) (unsigned)
Private Function calcUpgradeSum(asParts() As Long, ByVal numWeaps As Long, asWeaps() As Long, ByVal player As Long, ByVal what As Long) As Long
    Dim sum As Long, i As Long, body As Long, prop As Long, propPart As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/droid.cpp:1656", "calcUpgradeSum", "Invalid player: %u", player: Exit Function
    body = asParts(COMP_BODY)
    prop = asParts(COMP_PROPULSION)
    If what = SUM_HITPOINTS Then
        ' propulsion hitpoints can be a percentage of the body's hitpoints
        propPart = FromU32(Int(U32(MUL32(asBodyStats(body).upgrade(player).hitpoints, 100 + asPropulsionStats(prop).upgrade(player).hitpointPctOfBody)) / 100)) + asPropulsionStats(prop).upgrade(player).hitpoints
    Else
        propPart = asBodyStats(body).upgrade(player).hitpointPct - 100 + asPropulsionStats(prop).upgrade(player).hitpointPct - 100
    End If
    sum = componentUpgradeField(COMP_BRAIN, asParts(COMP_BRAIN), player, what) + _
          componentUpgradeField(COMP_SENSOR, asParts(COMP_SENSOR), player, what) + _
          componentUpgradeField(COMP_ECM, asParts(COMP_ECM), player, what) + _
          componentUpgradeField(COMP_REPAIRUNIT, asParts(COMP_REPAIRUNIT), player, what) + _
          componentUpgradeField(COMP_CONSTRUCT, asParts(COMP_CONSTRUCT), player, what) + _
          propPart
    For i = 0 To numWeaps - 1
        ' asWeaps[i] > 0 check only needed for droids, not templates.
        If asWeaps(i) <> 0 Then
            If Not (ULT(asWeaps(i), asWeaponStats_size)) Then ASSERT_FAILURE "asWeaps[i] < asWeaponStats.size()", "src/droid.cpp:1669", "calcUpgradeSum", "Invalid weapon stat index (%u, numWeaponStats: %zu) (player: %d)", asWeaps(i), asWeaponStats_size, player
            sum = sum + componentUpgradeField(COMP_WEAPON, asWeaps(i), player, what)
        End If
    Next
    calcUpgradeSum = sum
End Function

' (VB) a template's asParts and asWeaps as Long arrays
Private Sub templateParts(psTemplate As DROID_TEMPLATE, asParts() As Long, asWeaps() As Long)
    Dim i As Long
    ReDim asParts(0 To DROID_MAXCOMP - 1)
    ReDim asWeaps(0 To MAX_WEAPONS - 1)
    For i = 0 To DROID_MAXCOMP - 1
        asParts(i) = psTemplate.asParts(i)
    Next
    For i = 0 To MAX_WEAPONS - 1
        asWeaps(i) = psTemplate.asWeaps(i)
    Next
End Sub

' (VB) a droid's asBits, and FilterDroidWeaps: its weapons' stats without the zeros
Private Function droidParts(ByVal psDroid As Long, asParts() As Long, asWeaps() As Long) As Long
    Dim i As Long, n As Long
    ReDim asParts(0 To DROID_MAXCOMP - 1)
    ReDim asWeaps(0 To MAX_WEAPONS - 1)
    For i = 0 To DROID_MAXCOMP - 1
        asParts(i) = aDroid(psDroid).asBits(i)
    Next
    For i = 0 To aDroid(psDroid).numWeaps - 1
        If aDroid(psDroid).asWeaps(i).nStat <> 0 Then
            asWeaps(n) = aDroid(psDroid).asWeaps(i).nStat
            n = n + 1
        End If
    Next
    droidParts = n
End Function

' Calculate the weight of a droid from it's template
Public Function calcDroidWeight(psTemplate As DROID_TEMPLATE) As Long
    Dim asParts() As Long, asWeaps() As Long
    templateParts psTemplate, asParts, asWeaps
    calcDroidWeight = calcSum(asParts, psTemplate.numWeaps, asWeaps, SUM_WEIGHT)
End Function

' calcBody<T>(obj, player)
Private Function calcBody(asParts() As Long, ByVal numWeaps As Long, asWeaps() As Long, ByVal player As Long) As Long
    Dim hitpoints As Long, hitpointPct As Long
    hitpoints = calcUpgradeSum(asParts, numWeaps, asWeaps, player, SUM_HITPOINTS)
    hitpointPct = calcUpgradeSum(asParts, numWeaps, asWeaps, player, SUM_HITPOINTPCT)

    ' Final adjustment based on the hitpoint modifier
    calcBody = hitpoints * (100 + hitpointPct) \ 100
End Function

' Calculate the body points of a droid from its template (psTemplate: a DROID_TEMPLATE * handle, 0 = nullptr)
Public Function calcTemplateBody(ByVal psTemplate As Long, ByVal player As Long) As Long
    Dim asParts() As Long, asWeaps() As Long
    If psTemplate = 0 Then
        ASSERT_FAILURE "false", "src/droid.cpp:1754", "calcTemplateBody", "null template"
        calcTemplateBody = 0
        Exit Function
    End If

    templateParts apsTemplates(psTemplate), asParts, asWeaps
    calcTemplateBody = calcBody(asParts, apsTemplates(psTemplate).numWeaps, asWeaps, player And &HFF&)
End Function

' Calculate the base body points of a droid with upgrades
Private Function calcDroidBaseBody(ByVal psDroid As Long) As Long
    Dim asParts() As Long, asWeaps() As Long, n As Long
    n = droidParts(psDroid, asParts, asWeaps)
    calcDroidBaseBody = calcBody(asParts, n, asWeaps, aDroid(psDroid).player)
End Function

' Calculate the base speed of a droid from it's template
Public Function calcDroidBaseSpeed(psTemplate As DROID_TEMPLATE, ByVal weight As Long, ByVal player As Long) As Long
    Dim speed As Double, prop As Long, body As Long, w As Long
    prop = DROID_TEMPLATE_getPropulsionStats(psTemplate)
    body = DROID_TEMPLATE_getBodyStats(psTemplate)
    w = weight
    If ULT(w, 1) Then w = 1
    speed = Int(U32(asPropulsionTypes(asPropulsionStats(prop).propulsionType).powerRatioMult * _
                bodyPower(asBodyStats(body), player And &HFF&)) / U32(w))

    ' reduce the speed of medium/heavy VTOLs
    If asPropulsionStats(prop).propulsionType = PROPULSION_TYPE_LIFT Then
        If asBodyStats(body).size = SIZE_HEAVY Then
            speed = Int(speed / 4)
        ElseIf asBodyStats(body).size = SIZE_MEDIUM Then
            speed = Int(U32(FromU32(speed * 3)) / 4)
        End If
    End If

    ' applies the engine output bonus if output > weight
    If ULT(weight, asBodyStats(body).base.power) Then
        speed = Int(U32(FromU32(speed * 3)) / 2)
    End If

    calcDroidBaseSpeed = FromU32(speed)
End Function

' calcBuild<T>
Private Function calcBuild(asParts() As Long, ByVal numWeaps As Long, asWeaps() As Long) As Long
    calcBuild = calcSum(asParts, numWeaps, asWeaps, SUM_BUILDPOINTS)
End Function

' Calculate the points required to build the template - used to calculate time
Public Function calcTemplateBuild(psTemplate As DROID_TEMPLATE) As Long
    Dim asParts() As Long, asWeaps() As Long
    templateParts psTemplate, asParts, asWeaps
    calcTemplateBuild = calcBuild(asParts, psTemplate.numWeaps, asWeaps)
End Function

Public Function calcDroidPoints(ByVal psDroid As Long) As Long
    Dim asParts() As Long, asWeaps() As Long, n As Long
    n = droidParts(psDroid, asParts, asWeaps)
    calcDroidPoints = calcBuild(asParts, n, asWeaps)
End Function

' Calculate the power points required to build/maintain a template
Public Function calcTemplatePower(psTemplate As DROID_TEMPLATE) As Long
    Dim asParts() As Long, asWeaps() As Long
    templateParts psTemplate, asParts, asWeaps
    calcTemplatePower = calcSum(asParts, psTemplate.numWeaps, asWeaps, SUM_BUILDPOWER)
End Function

' Calculate the power points required to build/maintain a droid
Public Function calcDroidPower(ByVal psDroid As Long) As Long
    Dim asParts() As Long, asWeaps() As Long, n As Long
    If Not (psDroid <> 0) Then ASSERT_FAILURE "(obj) != nullptr", "src/droid.cpp:1840", "calcPower", "%s should not be nullptr", "obj": Exit Function
    n = droidParts(psDroid, asParts, asWeaps)
    calcDroidPower = calcSum(asParts, n, asWeaps, SUM_BUILDPOWER)
End Function

' ---- recycled_experience (std::priority_queue<int> per player): a list kept in descending order

' initialise droid module
Public Function droidInit() As Boolean
    Dim i As Long
    For i = 0 To MAX_PLAYERS - 1
        recycled_experience_size(i) = 0  ' clear it
    Next
    psLastDroidHit = 0

    moveInit

    droidInit = True
End Function


' (psWeap: the object's asWeaps(weapon_slot), passed ByRef)
Public Function droidReloadBar(ByVal psObj As Long, psWeap As WEAPON, ByVal weapon_slot As Long) As Long
    Dim psStats As Long, bSalvo As Boolean, firingStage As Long, interval As Long, player As Long, isVtolDroid As Boolean

    If psWeap.nStat = 0 Then ' no weapon
        droidReloadBar = -1
        Exit Function
    End If
    psStats = psWeap.nStat
    player = SIMPLE_OBJECT_player(psObj)

    ' Justifiable only when greater than a one second reload or intra salvo time
    bSalvo = (asWeaponStats(psStats).upgrade(player).numRounds > 1)
    isVtolDroid = False
    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then isVtolDroid = DROID_isVtol(psObj)
    If (bSalvo And asWeaponStats(psStats).upgrade(player).reloadTime > GAME_TICKS_PER_SEC) _
        Or asWeaponStats(psStats).upgrade(player).firePause > GAME_TICKS_PER_SEC _
        Or isVtolDroid Then
        If isVtolDroid Then
            'deal with VTOLs
            firingStage = getNumAttackRuns(psObj, weapon_slot) - aDroid(psObj).asWeaps(weapon_slot).usedAmmo

            'compare with max value
            interval = getNumAttackRuns(psObj, weapon_slot)
        Else
            firingStage = gameTime - psWeap.lastFired
            If bSalvo Then
                interval = weaponReloadTime(asWeaponStats(psStats), player)
            Else
                interval = weaponFirePause(asWeaponStats(psStats), player)
            End If
        End If
        If firingStage < interval And interval > 0 Then
            droidReloadBar = (firingStage * 100) \ interval  ' PERCENT(firingStage, interval)
            Exit Function
        End If
        droidReloadBar = 100
        Exit Function
    End If
    droidReloadBar = -1
End Function

' std::priority_queue<int> copy_experience_queue(int player): (VB) the queue's values in values(0 .. n - 1),
' the top first (the order the copy's pops give); returns n
Public Function copy_experience_queue(ByVal player As Long, values() As Long) As Long
    Dim k As Long, n As Long
    n = recycled_experience_size(player)
    If n > 0 Then
        ReDim values(0 To n - 1)
        For k = 0 To n - 1
            values(k) = recycled_experience(player, k)     ' kept sorted, the largest first
        Next
    End If
    copy_experience_queue = n
End Function

Public Function getTopExperience(ByVal player As Long) As Long
    If recycled_experience_size(player) = 0 Then
        getTopExperience = 0
        Exit Function
    End If
    getTopExperience = recycled_experience(player, 0)
End Function

Public Sub add_to_experience_queue(ByVal player As Long, ByVal value As Long)
    Dim n As Long, k As Long
    n = recycled_experience_size(player)
    If n >= recycled_experience_cap Then
        If recycled_experience_cap = 0 Then recycled_experience_cap = 64 Else recycled_experience_cap = recycled_experience_cap * 2
        ReDim Preserve recycled_experience(0 To MAX_PLAYERS - 1, 0 To recycled_experience_cap - 1)
    End If
    k = n
    Do While k > 0
        If recycled_experience(player, k - 1) >= value Then Exit Do
        recycled_experience(player, k) = recycled_experience(player, k - 1)
        k = k - 1
    Loop
    recycled_experience(player, k) = value
    recycled_experience_size(player) = n + 1
End Sub

' (VB) recycled_experience[player].pop()
Private Sub recycled_experience_pop(ByVal player As Long)
    Dim k As Long
    For k = 1 To recycled_experience_size(player) - 1
        recycled_experience(player, k - 1) = recycled_experience(player, k)
    Next
    recycled_experience_size(player) = recycled_experience_size(player) - 1
End Sub

' ---- building a droid

'Builds an instance of a Droid - the x/y passed in are in world coords.
' Sends a GAME_DROID message if bMultiMessages is true, or actually creates it if false. Only uses initialOrders if sending a GAME_DROID message.
Public Function buildDroid(ByVal pTemplate As Long, ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal onMission As Boolean, ByVal hasInitialOrders As Boolean, initialOrders As INITIAL_DROID_ORDERS, rot As Rotation) As Long
    Dim pos As Vector3i
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/droid.cpp:1979", "buildDroid", "invalid player?: %u", player: buildDroid = 0: Exit Function
    ' ajl. droid will be created, so inform others
    If bMultiMessages Then
        ' Only sends if it's ours, otherwise the owner should send the message.
        SendDroid pTemplate, x, y, player, generateNewObjectId(), hasInitialOrders, initialOrders
        buildDroid = 0
    Else
        pos.x = x
        pos.y = y
        pos.z = 0
        buildDroid = reallyBuildDroid(pTemplate, pos, player, onMission, rot)
    End If
End Function

' (the two overloads: the id missing, a new synchronised one)
Public Function reallyBuildDroid(ByVal pTemplate As Long, pos As Vector3i, ByVal player As Long, ByVal onMission As Boolean, rot As Rotation, Optional ByVal id As Variant) As Long
    Dim psDroid As Long, psGrp As Long

    If IsMissing(id) Then id = generateSynchronisedObjectId()

    ' Don't use this assertion in single player, since droids can finish building while on an away mission
    If bMultiPlayer Then
        If Not worldOnMap(pos.x, pos.y) Then ASSERT_FAILURE "!bMultiPlayer || worldOnMap(pos.x, pos.y)", "src/droid.cpp:1866", "reallyBuildDroid", "the build locations are not on the map"
    End If

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/droid.cpp:1868", "reallyBuildDroid", "Invalid player: %u", player: Exit Function

    psDroid = GlobalDroidContainer_emplace(CLng(id), player)
    droidSetName psDroid, apsTemplates(pTemplate).name_

    ' Set the droids type
    aDroid(psDroid).droidType = droidTemplateType(apsTemplates(pTemplate))  ' Is set again later to the same thing, in droidSetBits.
    aDroid(psDroid).pos = pos
    aDroid(psDroid).rot = rot

    'don't worry if not on homebase cos not being drawn yet
    If Not onMission Then
        'set droid height
        aDroid(psDroid).pos.z = map_Height(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)
    End If

    If DROID_isTransporter(psDroid) Or aDroid(psDroid).droidType = DROID_COMMAND Then
        psGrp = grpCreate()
        DROID_GROUP_add psGrp, psDroid
    End If

    ' find the highest stored experience
    ' Unless game time is stopped, then we're hopefully loading a game and
    ' don't want to use up recycled experience for the droids we just loaded.
    If Not gameTimeIsStopped() And _
       (aDroid(psDroid).droidType <> DROID_CONSTRUCT) And _
       (aDroid(psDroid).droidType <> DROID_CYBORG_CONSTRUCT) And _
       (aDroid(psDroid).droidType <> DROID_REPAIR) And _
       (aDroid(psDroid).droidType <> DROID_CYBORG_REPAIR) And _
       Not DROID_isTransporter(psDroid) And _
       recycled_experience_size(aDroid(psDroid).player) <> 0 Then
        aDroid(psDroid).experience = recycled_experience(aDroid(psDroid).player, 0)
        recycled_experience_pop aDroid(psDroid).player
    Else
        aDroid(psDroid).experience = 0
    End If
    aDroid(psDroid).kills = 0
    aDroid(psDroid).shieldPoints = -1

    droidSetBits pTemplate, psDroid

    'calculate the droids total weight
    aDroid(psDroid).weight = calcDroidWeight(apsTemplates(pTemplate))

    ' Initialise the movement stuff
    aDroid(psDroid).baseSpeed = calcDroidBaseSpeed(apsTemplates(pTemplate), aDroid(psDroid).weight, player And &HFF&)

    initDroidMovement psDroid

    'allocate 'easy-access' data!
    aDroid(psDroid).body = calcDroidBaseBody(psDroid) ' includes upgrades
    If Not (aDroid(psDroid).body <> 0) Then ASSERT_FAILURE "droid.body > 0", "src/droid.cpp:1924", "reallyBuildDroid", "Invalid number of hitpoints"
    aDroid(psDroid).originalBody = aDroid(psDroid).body

    ' Set droid's initial illumination
    aDroid(psDroid).sDisplay.imd = asBodyStats(DROID_getBodyStats(psDroid)).pIMD   ' BODY_IMD(&droid, droid.player)

    'don't worry if not on homebase cos not being drawn yet
    If Not onMission Then
        ' People always stand upright
        If aDroid(psDroid).droidType <> DROID_PERSON Then
            updateDroidOrientation psDroid
        End If
        visTilesUpdate psDroid
    End If

    ' transporter-specific stuff
    If DROID_isTransporter(psDroid) Then
        'add transporter launch button if selected player and not a reinforcable situation
        If player = selectedPlayer Then
            If Not missionCanReEnforce() Then
                intAddTransporterLaunch psDroid
            End If
        End If

        'set droid height to be above the terrain
        aDroid(psDroid).pos.z = aDroid(psDroid).pos.z + TRANSPORTER_HOVER_HEIGHT

        ' reset halt secondary order from guard to hold
        secondarySetState psDroid, DSO_HALTTYPE, DSS_HALT_HOLD
    End If

    If player = selectedPlayer Then
        scoreUpdateVar WD_UNITS_BUILT
    End If

    ' Avoid droid appearing to jump or turn on spawn.
    aDroid(psDroid).prevSpacetime.pos = aDroid(psDroid).pos
    aDroid(psDroid).prevSpacetime.rot = aDroid(psDroid).rot

    debug_ 1966, LOG_LIFE, "reallyBuildDroid", "created droid for player %d, droid = %p, id=%d (%s): position: x(%d)y(%d)z(%d)", player, psDroid, aDroid(psDroid).id, aDroid(psDroid).aName, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, aDroid(psDroid).pos.z

    reallyBuildDroid = psDroid
End Function

'initialises the droid movement model
Public Sub initDroidMovement(ByVal psDroid As Long)
    Erase aDroid(psDroid).sMove.asPath
    aDroid(psDroid).sMove.asPath_size = 0
    aDroid(psDroid).sMove.pathIndex = 0
End Sub

' Set the asBits in a DROID structure given it's template.
Public Sub droidSetBits(ByVal pTemplate As Long, ByVal psDroid As Long)
    Dim inc As Long
    aDroid(psDroid).droidType = droidTemplateType(apsTemplates(pTemplate))
    aDroid(psDroid).numWeaps = apsTemplates(pTemplate).numWeaps
    aDroid(psDroid).body = calcTemplateBody(pTemplate, aDroid(psDroid).player)
    aDroid(psDroid).originalBody = aDroid(psDroid).body
    aDroid(psDroid).expectedDamageDirect = 0  ' Begin life optimistically.
    aDroid(psDroid).expectedDamageIndirect = 0  ' Begin life optimistically.
    aDroid(psDroid).time_ = gameTime - deltaGameTime + 1         ' Start at beginning of tick.
    aDroid(psDroid).prevSpacetime.time_ = aDroid(psDroid).time_ - 1  ' -1 for interpolation.

    'create the droids weapons
    For inc = 0 To MAX_WEAPONS - 1
        aDroid(psDroid).psActionTarget(inc) = 0
        aDroid(psDroid).asWeaps(inc).lastFired = 0
        aDroid(psDroid).asWeaps(inc).shotsFired = 0
        ' no weapon (could be a construction droid for example)
        ' this is also used to check if a droid has a weapon, so zero it
        aDroid(psDroid).asWeaps(inc).nStat = 0
        aDroid(psDroid).asWeaps(inc).ammo = 0
        aDroid(psDroid).asWeaps(inc).rot.direction = 0
        aDroid(psDroid).asWeaps(inc).rot.pitch = 0
        aDroid(psDroid).asWeaps(inc).rot.roll = 0
        aDroid(psDroid).asWeaps(inc).prevRot = aDroid(psDroid).asWeaps(inc).rot
        aDroid(psDroid).asWeaps(inc).origin = ORIGIN_UNKNOWN
        If inc < apsTemplates(pTemplate).numWeaps Then
            aDroid(psDroid).asWeaps(inc).nStat = apsTemplates(pTemplate).asWeaps(inc)
            aDroid(psDroid).asWeaps(inc).ammo = asWeaponStats(aDroid(psDroid).asWeaps(inc).nStat).upgrade(aDroid(psDroid).player).numRounds
        End If
        aDroid(psDroid).asWeaps(inc).usedAmmo = 0
    Next
    For inc = 0 To DROID_MAXCOMP - 1
        aDroid(psDroid).asBits(inc) = apsTemplates(pTemplate).asParts(inc)
    Next

    Select Case asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType  ' getPropulsionStats(psDroid) only defined after psDroid->asBits[COMP_PROPULSION] is set.
    Case PROPULSION_TYPE_LIFT
        aDroid(psDroid).blockedBits = AIR_BLOCKED
    Case PROPULSION_TYPE_HOVER
        aDroid(psDroid).blockedBits = FEATURE_BLOCKED
    Case PROPULSION_TYPE_PROPELLOR
        aDroid(psDroid).blockedBits = FEATURE_BLOCKED Or LAND_BLOCKED
    Case Else
        aDroid(psDroid).blockedBits = FEATURE_BLOCKED Or WATER_BLOCKED
    End Select
End Sub

' ---- removing and destroying droids

Public Function removeDroidBase(ByVal psDel As Long) As Boolean
    Dim it As Long, nx As Long, psCurr As Long, numSelectedConstructors As Long, psDroid As Long

    checkDroid psDel, "src/droid.cpp:536", "removeDroidBase", max_check_object_recursion

    If isDead(psDel) Then
        ' droid has already been killed, quit
        syncDebug "removeDroidBase", "droid already dead"
        removeDroidBase = True
        Exit Function
    End If

    syncDebugDroid_ "removeDroidBase", psDel, Asc("#")

    ' leave the current formation if any
    If aDroid(psDel).sMove.psFormation <> 0 Then
        formationLeave aDroid(psDel).sMove.psFormation, psDel
        aDroid(psDel).sMove.psFormation = 0
    End If

    If aDroid(psDel).action = DACTION_DROIDREPAIR Then
        droidRepairStopped castDroid(aDroid(psDel).psActionTarget(0)), psDel
    End If

    'kill all the droids inside the transporter
    If DROID_isTransporter(psDel) Then
        If aDroid(psDel).psGroup <> 0 Then
            'free all droids associated with this Transporter
            ' (mutating_list_iterate)
            it = aGroup(aDroid(psDel).psGroup).psList.head
            Do While it <> 0
                nx = lnNext(it)
                psCurr = lnObj(it)
                If psCurr = psDel Then
                    Exit Do
                End If
                ' add droid to droid list then vanish it - hope this works! - GJ
                addDroid psCurr, apsDroidLists
                vanishDroid psCurr
                it = nx
            Loop
        End If
    End If

    ' leave the current group if any
    If aDroid(psDel).psGroup <> 0 Then
        DROID_GROUP_remove aDroid(psDel).psGroup, psDel
        aDroid(psDel).psGroup = 0
    End If

    ' Put Deliv. Pts back into world when a command droid dies
    If aDroid(psDel).droidType = DROID_COMMAND Then
        it = apsStructLists(aDroid(psDel).player).head
        Do While it <> 0
            psCurr = lnObj(it)
            ' alexl's stab at a right answer.
            If psCurr <> 0 Then
                If STRUCTURE_isFactory(psCurr) Then
                    If aStruct(psCurr).functionality.factory.psCommander = psDel Then
                        assignFactoryCommandDroid psCurr, 0
                    End If
                End If
            End If
            it = lnNext(it)
        Loop
    End If

    ' Check to see if constructor droid currently trying to find a location to build
    If aDroid(psDel).player = selectedPlayer And aDroid(psDel).selected <> 0 And DROID_isConstructionDroid(psDel) Then
        ' If currently trying to build, kill off the placement
        If tryingToGetLocation() Then
            numSelectedConstructors = 0
            it = apsDroidLists(aDroid(psDel).player).head
            Do While it <> 0
                psDroid = lnObj(it)
                numSelectedConstructors = numSelectedConstructors - CLng(aDroid(psDroid).selected <> 0 And DROID_isConstructionDroid(psDroid))
                it = lnNext(it)
            Loop
            If numSelectedConstructors <= 1 Then  ' If we were the last selected construction droid.
                kill3DBuilding
            End If
        End If
    End If

    If aDroid(psDel).group <> UBYTE_MAX Then
        intGroupsChanged
    End If

    If aDroid(psDel).player = selectedPlayer Then
        intInformInterfaceObjectRemoved psDel
        intRefreshScreen
    End If

    killDroid psDel
    removeDroidBase = True
End Function

Private Sub removeDroidFX(ByVal psDel As Long, ByVal impactTime As Long)
    Dim pos As Vector3i, adiff As Long, move As Long, velocity As Vector3f

    checkDroid psDel, "src/droid.cpp:638", "removeDroidFX", max_check_object_recursion

    ' only display anything if the droid is visible
    If BASE_OBJECT_visibleForLocalDisplay(psDel) = 0 Then
        Exit Sub
    End If

    adiff = aDroid(psDel).rot.direction - aDroid(psDel).sMove.moveDir
    move = iCosR(adiff, aDroid(psDel).sMove.speed)
    velocity.x = iSinR(aDroid(psDel).rot.direction, move)
    velocity.y = 0!
    velocity.z = iCosR(aDroid(psDel).rot.direction, move)
    velocity.x = fmul(velocity.x, DROID_EXPLOSION_VELOCITY_FACTOR)
    velocity.y = fmul(velocity.y, DROID_EXPLOSION_VELOCITY_FACTOR)
    velocity.z = fmul(velocity.z, DROID_EXPLOSION_VELOCITY_FACTOR)

    If aDroid(psDel).animationEvent <> ANIM_EVENT_DYING Then
        compPersonToBits psDel, velocity
    End If

    ' if baba then squish
    If aDroid(psDel).droidType = DROID_PERSON Then
        ' The barbarian has been run over ...
        audio_PlayStaticTrack aDroid(psDel).pos.x, aDroid(psDel).pos.y, ID_SOUND_BARB_SQUISH
    Else
        destroyFXDroid psDel, impactTime, velocity
        pos.x = aDroid(psDel).pos.x
        pos.z = aDroid(psDel).pos.y
        pos.y = aDroid(psDel).pos.z
        If aDroid(psDel).droidType = DROID_SUPERTRANSPORTER Then
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_LARGE, False, 0, 0, impactTime
        Else
            addEffect pos, EFFECT_DESTRUCTION, DESTRUCTION_TYPE_DROID, False, 0, 0, impactTime
        End If
        audio_PlayStaticTrack aDroid(psDel).pos.x, aDroid(psDel).pos.y, ID_SOUND_EXPLOSION
    End If
End Sub

Public Function destroyDroid(ByVal psDel As Long, ByVal impactTime As Long) As Boolean
    Dim width As Long, breadth As Long, mapX As Long, mapY As Long, psTile As Long

    If Not (Not ULT(impactTime, gameTime - deltaGameTime)) Then ASSERT_FAILURE "gameTime - deltaGameTime <= impactTime", "src/droid.cpp:683", "destroyDroid", "Expected %u <= %u, gameTime = %u, bad impactTime", gameTime - deltaGameTime, impactTime, gameTime

    If aDroid(psDel).lastHitWeapon = WSC_LAS_SAT Then     ' darken tile if lassat.
        mapX = map_coord(aDroid(psDel).pos.x)
        mapY = map_coord(aDroid(psDel).pos.y)
        width = mapX - 1
        Do While Not ULT(mapX + 1, width)
            breadth = mapY - 1
            Do While Not ULT(mapY + 1, breadth)
                psTile = mapTile(width, breadth)
                If TEST_TILE_VISIBLE_TO_SELECTEDPLAYER(psTile) Then
                    psMapTiles(psTile).illumination = psMapTiles(psTile).illumination \ 2
                    psMapTiles(psTile).ambientOcclusion = psMapTiles(psTile).ambientOcclusion \ 2
                End If
                breadth = breadth + 1
            Loop
            width = width + 1
        Loop
    End If

    removeDroidFX psDel, impactTime
    removeDroidBase psDel
    aDroid(psDel).died = impactTime
    destroyDroid = True
End Function

Public Sub vanishDroid(ByVal psDel As Long)
    removeDroidBase psDel
End Sub

' ---- the droid sync log

' syncDebugDroid(psDroid, ch) is syncDebugDroid_ function, psDroid, ch
Public Sub syncDebugDroid_(ByVal function_ As String, ByVal psDroid As Long, ByVal ch As Long)
    Dim list(0 To 34) As Long
    If aDroid(psDroid).typ <> OBJ_DROID Then
        ASSERT_FAILURE "false", "src/droid.cpp:835", "_syncDebugDroid", "%c Broken psDroid->type %u!", ch, aDroid(psDroid).typ
        syncDebug "_syncDebugDroid", "Broken psDroid->type %u!", aDroid(psDroid).typ
    End If
    list(0) = ch

    list(1) = aDroid(psDroid).id

    list(2) = aDroid(psDroid).player
    list(3) = aDroid(psDroid).pos.x
    list(4) = aDroid(psDroid).pos.y
    list(5) = aDroid(psDroid).pos.z
    list(6) = aDroid(psDroid).rot.direction
    list(7) = aDroid(psDroid).rot.pitch
    list(8) = aDroid(psDroid).rot.roll
    list(9) = aDroid(psDroid).order.typ
    list(10) = aDroid(psDroid).order.pos.x
    list(11) = aDroid(psDroid).order.pos.y
    list(12) = aDroid(psDroid).listSize
    list(13) = aDroid(psDroid).action
    list(14) = aDroid(psDroid).secondaryOrder
    list(15) = aDroid(psDroid).body
    list(16) = aDroid(psDroid).sMove.Status
    list(17) = aDroid(psDroid).sMove.speed
    list(18) = aDroid(psDroid).sMove.moveDir
    list(19) = aDroid(psDroid).sMove.pathIndex
    list(20) = aDroid(psDroid).sMove.asPath_size
    list(21) = aDroid(psDroid).sMove.src.x
    list(22) = aDroid(psDroid).sMove.src.y
    list(23) = aDroid(psDroid).sMove.target.x
    list(24) = aDroid(psDroid).sMove.target.y
    list(25) = aDroid(psDroid).sMove.destination.x
    list(26) = aDroid(psDroid).sMove.destination.y
    list(27) = aDroid(psDroid).sMove.bumpDir
    list(28) = aDroid(psDroid).sMove.bumpTime
    list(29) = aDroid(psDroid).sMove.lastBump
    list(30) = aDroid(psDroid).sMove.pauseTime
    list(31) = aDroid(psDroid).sMove.bumpPos.x
    list(32) = aDroid(psDroid).sMove.bumpPos.y
    list(33) = aDroid(psDroid).sMove.shuffleStart
    list(34) = aDroid(psDroid).experience
    syncDebugIntList_ function_, "%c droid%d = p%d;pos(%d,%d,%d),rot(%d,%d,%d),order%d(%d,%d)^%d,action%d,secondaryOrder%X,body%d,sMove(status%d,speed%d,moveDir%d,path%d/%d,src(%d,%d),target(%d,%d),destination(%d,%d),bump(%d,%d,%d,%d,(%d,%d),%d)),exp%u", list, 35
End Sub

' ---- commanders, speed and experience

Public Function droidWithinCommanderRange(ByVal psDroid As Long) As Boolean
    Dim ok As Boolean, psCommander As Long, rangeArraySize As Long, level As Long, sqDist As Long, maxSqDist As Long, brain As Long

    If aDroid(psDroid).droidType = DROID_COMMAND Then
        droidWithinCommanderRange = True
        Exit Function
    End If

    ok = aDroid(psDroid).psGroup <> 0
    If ok Then ok = aGroup(aDroid(psDroid).psGroup).psCommander <> 0
    If Not ok Then ASSERT_FAILURE "psDroid->psGroup && psDroid->psGroup->psCommander", "src/droid.cpp:1058", "droidWithinCommanderRange", "Droid group or commander is NULL": droidWithinCommanderRange = False: Exit Function
    psCommander = aGroup(aDroid(psDroid).psGroup).psCommander

    brain = DROID_getBrainStats(psDroid)
    rangeArraySize = asBrainStats(brain).cmdExpRange_size
    If rangeArraySize = 0 Then
        droidWithinCommanderRange = True ' default to true (matches old behavior, which didn't have limits for cmdExpRange, etc)
        Exit Function
    End If

    level = getDroidLevel(psCommander)
    If level >= rangeArraySize Then
        level = rangeArraySize - 1 ' use the last listed value, for the last listed level in the array
    End If

    sqDist = objPosDiffSq_obj(psDroid, psCommander)
    maxSqDist = asBrainStats(brain).cmdExpRange(level) * asBrainStats(brain).cmdExpRange(level)

    droidWithinCommanderRange = sqDist <= maxSqDist
End Function

' UDWORD calcDroidSpeed(UDWORD baseSpeed, UDWORD terrainType, UDWORD propIndex, UDWORD level)
Public Function calcDroidSpeed(ByVal baseSpeed As Long, ByVal terrainType_ As Long, ByVal propIndex As Long, ByVal level As Long) As Long
    Dim speed As Long

    ' Factor in terrain
    speed = baseSpeed * getSpeedFactor(terrainType_, asPropulsionStats(propIndex).propulsionType) \ 100

    ' Need to ensure doesn't go over the max speed possible for this propulsion
    If ULT(asPropulsionStats(propIndex).maxSpeed, speed) Then speed = asPropulsionStats(propIndex).maxSpeed

    ' Factor in experience
    speed = speed * (100 + EXP_SPEED_BONUS * level)
    speed = speed \ 100

    calcDroidSpeed = speed
End Function

' unsigned int getDroidLevel(unsigned int experience, uint8_t player, uint8_t brainComponent)
Public Function getDroidLevel_exp(ByVal experience As Long, ByVal player As Long, ByVal brainComponent As Long) As Long
    Dim numKills As Long, i As Long, n As Long
    numKills = SHRU(experience, 16)         ' experience / 65536
    If Not (brainComponent < asBrainStats_size) Then ASSERT_FAILURE "brainComponent < asBrainStats.size()", "src/droid.cpp:2491", "getDroidLevel", "Invalid brainComponent: %u", brainComponent: getDroidLevel_exp = 0: Exit Function
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/droid.cpp:2493", "getDroidLevel", "Invalid player: %u", player: getDroidLevel_exp = 0: Exit Function
    n = asBrainStats(brainComponent).upgrade(player).rankThresholds_size
    If Not (n > 0) Then ASSERT_FAILURE "vec.size() > 0", "src/droid.cpp:2495", "getDroidLevel", "rankThreshold was empty?": getDroidLevel_exp = 0: Exit Function
    For i = 1 To n - 1
        If ULT(numKills, asBrainStats(brainComponent).upgrade(player).rankThresholds(i)) Then
            getDroidLevel_exp = i - 1
            Exit Function
        End If
    Next

    ' If the criteria of the last rank are met, then select the last one
    getDroidLevel_exp = n - 1
End Function

Public Function getDroidLevel(ByVal psDroid As Long) As Long
    Dim brain As Long
    ' Sensors will use the first non-null brain component for ranks
    If aDroid(psDroid).droidType <> DROID_SENSOR Then brain = aDroid(psDroid).asBits(COMP_BRAIN) Else brain = 1
    getDroidLevel = getDroidLevel_exp(aDroid(psDroid).experience, aDroid(psDroid).player, brain)
End Function

' (VB) PE_("rank", ...): the rank's name untranslated
Public Function getDroidLevelName(ByVal psDroid As Long) As String
    Dim psStats As Long
    psStats = DROID_getBrainStats(psDroid)
    getDroidLevelName = asBrainStats(psStats).rankNames.items(getDroidLevel(psDroid))
End Function

Public Function getDroidEffectiveLevel(ByVal psDroid As Long, ByVal commanderDistanceCheck As Boolean) As Long
    Dim level As Long, cmdLevel As Long, inRange As Boolean
    level = getDroidLevel(psDroid)
    cmdLevel = 0

    ' get commander level
    If hasCommander(psDroid) Then
        inRange = Not commanderDistanceCheck
        If Not inRange Then inRange = droidWithinCommanderRange(psDroid)
        If inRange Then
            cmdLevel = cmdGetCommanderLevel(psDroid)

            ' Commanders boost units' effectiveness just by being assigned to it
            level = level + 1
        End If
    End If

    getDroidEffectiveLevel = MAX_(level, cmdLevel)
End Function

Public Sub droidIncreaseExperience(ByVal psDroid As Long, ByVal experienceInc As Long)
    Dim startingDroidRank As Long, finalDroidRank As Long
    If Not bMultiPlayer Then
        If getCamTweakOption_FastExp() Then
            experienceInc = FromU32(U32(experienceInc) * 2)
        End If
    End If

    If Not (ULT(experienceInc, 137625)) Then ASSERT_FAILURE "experienceInc < (int)(2.1 * 65536)", "src/droid.cpp:2595", "droidIncreaseExperience", "Experience increase out of range": Exit Sub

    startingDroidRank = getDroidLevel(psDroid)

    aDroid(psDroid).experience = FromU32(U32(aDroid(psDroid).experience) + U32(experienceInc))

    finalDroidRank = getDroidLevel(psDroid)
    If startingDroidRank <> finalDroidRank Then
        ' Trigger new event - unit rank increased
        triggerEventDroidRankGained psDroid, finalDroidRank
    End If
End Sub

' Possibly increase experience when squishing a scavenger.
Public Sub giveExperienceForSquish(ByVal psDroid As Long)
    Dim expGain As Long
    If aDroid(psDroid).droidType = DROID_WEAPON Or aDroid(psDroid).droidType = DROID_SENSOR Or aDroid(psDroid).droidType = DROID_COMMAND Then
        expGain = MAX_(65536 \ 2, 65536 * getExpGain(aDroid(psDroid).player) \ 100)
        droidIncreaseExperience psDroid, expGain
        cmdDroidUpdateExperience psDroid, expGain
    End If
End Sub

' true if a droid is retreating for repair
Public Function DROID_isRetreatingForRepair(ByVal this As Long) As Boolean
    DROID_isRetreatingForRepair = aDroid(this).order.typ = DORDER_RTR Or aDroid(this).order.typ = DORDER_RTR_SPECIFIED
End Function

' true if a droid is returning to base
Public Function DROID_isReturningToBase(ByVal this As Long) As Boolean
    DROID_isReturningToBase = aDroid(this).order.typ = DORDER_RTB
End Function

Public Function droidOnMap(ByVal psDroid As Long) As Boolean
    If aDroid(psDroid).died = NOT_CURRENT_LIST Or DROID_isTransporter(psDroid) _
       Or aDroid(psDroid).pos.x = INVALID_XY Or aDroid(psDroid).pos.y = INVALID_XY Or missionIsOffworld() _
       Or mapHeight = 0 Then
        ' Off world or on a transport or is a transport or in mission list, or on a mission, or no map - ignore
        droidOnMap = True
        Exit Function
    End If
    droidOnMap = worldOnMap(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)
End Function

' Check validity of a droid. Crash hard if it fails.
' droid.h: CHECK_DROID(droid) is checkDroid droid, "<file>:<line>", "<function>", max_check_object_recursion.
' (VB) Its ASSERT_HELPERs log under the caller's location, as in C++; the [function:line] of a failure's log
' line is the caller's here (C++: checkDroid's own line).
Public Sub checkDroid(ByVal droid_ As Long, ByVal location As String, ByVal function_ As String, ByVal recurse As Long)
    Dim i As Long
    If recurse < 0 Then
        Exit Sub
    End If

    If Not (droid_ <> 0) Then ASSERT_HELPER_FAILURE 3899, "droid != nullptr", location, function_, "CHECK_DROID: NULL pointer": Exit Sub
    If Not (aDroid(droid_).typ = OBJ_DROID) Then ASSERT_HELPER_FAILURE 3900, "droid->type == OBJ_DROID", location, function_, "CHECK_DROID: Not droid (type %d)", aDroid(droid_).typ
    If Not (aDroid(droid_).numWeaps <= MAX_WEAPONS) Then ASSERT_HELPER_FAILURE 3901, "droid->numWeaps <= MAX_WEAPONS", location, function_, "CHECK_DROID: Bad number of droid weapons %d", aDroid(droid_).numWeaps
    If Not (Not ULT(aDroid(droid_).asOrderList_size, aDroid(droid_).listSize) And Not ULT(aDroid(droid_).asOrderList_size, aDroid(droid_).listPendingBegin)) Then ASSERT_HELPER_FAILURE 3902, "(unsigned)droid->listSize <= droid->asOrderList.size() && (unsigned)droid->listPendingBegin <= droid->asOrderList.size()", location, function_, "CHECK_DROID: Bad number of droid orders %d %d %d", aDroid(droid_).listSize, aDroid(droid_).listPendingBegin, aDroid(droid_).asOrderList_size
    If Not (aDroid(droid_).player < MAX_PLAYERS) Then ASSERT_HELPER_FAILURE 3903, "droid->player < MAX_PLAYERS", location, function_, "CHECK_DROID: Bad droid owner %d", aDroid(droid_).player
    If Not (droidOnMap(droid_)) Then ASSERT_HELPER_FAILURE 3904, "droidOnMap(droid)", location, function_, "CHECK_DROID: Droid off map"
    If Not (Not ULT(aDroid(droid_).originalBody, aDroid(droid_).body)) Then ASSERT_HELPER_FAILURE 3905, "droid->body <= droid->originalBody", location, function_, "CHECK_DROID: More body points (%u) than original body points (%u).", aDroid(droid_).body, aDroid(droid_).originalBody

    For i = 0 To MAX_WEAPONS - 1
        If Not (Not ULT(gameTime, aDroid(droid_).asWeaps(i).lastFired)) Then ASSERT_HELPER_FAILURE 3909, "droid->asWeaps[i].lastFired <= gameTime", location, function_, "CHECK_DROID: Bad last fired time for turret %u", i
    Next
End Sub
' ---- B7: what the actions and orders call

Public Sub cancelBuild(ByVal psDroid As Long)
    Dim t As DroidOrderType
    t = aDroid(psDroid).order.typ
    If t = DORDER_NONE Or t = DORDER_PATROL Or t = DORDER_HOLD Or t = DORDER_SCOUT Or t = DORDER_GUARD Then
        objTrace aDroid(psDroid).id, "cancelBuild", "Droid build action cancelled"
        aDroid(psDroid).order.psObj = 0
        aDroid(psDroid).action = DACTION_NONE
        setDroidActionTarget psDroid, 0, 0
        Exit Sub  ' Don't cancel orders.
    End If

    If orderDroidList(psDroid) Then
        objTrace aDroid(psDroid).id, "cancelBuild", "Droid build order cancelled - changing to next order"
    Else
        objTrace aDroid(psDroid).id, "cancelBuild", "Droid build order cancelled"
        aDroid(psDroid).action = DACTION_NONE
        aDroid(psDroid).order = DroidOrder_(DORDER_NONE)
        setDroidActionTarget psDroid, 0, 0

        ' The droid has no more build orders, so halt in place rather than clumping around the build objective
        moveStopDroid psDroid

        triggerEventDroidIdle psDroid
    End If
End Sub

Private Sub droidBodyUpgrade(ByVal psDroid As Long)
    Const factor As Long = 10000 ' use big numbers to scare away rounding errors
    Dim prev As Long, increase As Long, sTemplate As DROID_TEMPLATE, it As Long, psCurr As Long, v As Long
    prev = aDroid(psDroid).originalBody
    aDroid(psDroid).originalBody = calcDroidBaseBody(psDroid)
    increase = FromU32(Fix(U32(MUL32(aDroid(psDroid).originalBody, factor)) / U32(prev)))
    ' MIN(originalBody, (body * increase) / factor + 1), unsigned
    v = FromU32(Fix(U32(MUL32(aDroid(psDroid).body, increase)) / factor) + 1)
    If ULT(aDroid(psDroid).originalBody, v) Then v = aDroid(psDroid).originalBody
    aDroid(psDroid).body = v
    templateSetParts psDroid, sTemplate
    ' update engine too
    aDroid(psDroid).baseSpeed = calcDroidBaseSpeed(sTemplate, aDroid(psDroid).weight, aDroid(psDroid).player)
    If DROID_isTransporter(psDroid) Then
        it = aGroup(aDroid(psDroid).psGroup).psList.head
        Do While it <> 0
            psCurr = lnObj(it)
            If psCurr <> psDroid Then
                droidBodyUpgrade psCurr
            End If
            it = lnNext(it)
        Loop
    End If
End Sub

' recycle a droid (retain it's experience and some of it's cost)
Public Sub recycleDroid(ByVal psDroid As Long)
    Dim cost As Long, position As Vector3i, mapCoordX As Long, mapCoordY As Long, psTile As Long, i As Long

    checkDroid psDroid, "src/droid.cpp:499", "recycleDroid", max_check_object_recursion

    ' store the droids kills
    If aDroid(psDroid).experience <> 0 Then
        add_to_experience_queue aDroid(psDroid).player, aDroid(psDroid).experience
    End If

    ' return part of the cost of the droid
    cost = calcDroidPower(psDroid)
    cost = FromU32(Fix(U32(MUL32(cost \ 2, aDroid(psDroid).body)) / U32(aDroid(psDroid).originalBody)))
    addPower aDroid(psDroid).player, cost

    ' hide the droid
    For i = 0 To MAX_PLAYERS - 1
        aDroid(psDroid).visible(i) = 0
    Next

    If aDroid(psDroid).psGroup <> 0 Then
        DROID_GROUP_remove aDroid(psDroid).psGroup, psDroid
    End If

    triggerEvent TRIGGER_OBJECT_RECYCLED, psDroid
    vanishDroid psDroid

    position.x = aDroid(psDroid).pos.x
    position.y = aDroid(psDroid).pos.z
    position.z = aDroid(psDroid).pos.y
    mapCoordX = map_coord(aDroid(psDroid).pos.x)
    mapCoordY = map_coord(aDroid(psDroid).pos.y)
    psTile = mapTile(mapCoordX, mapCoordY)
    If tileIsClearlyVisible(psTile) Then
        addEffect position, EFFECT_EXPLOSION, EXPLOSION_TYPE_DISCOVERY, False, 0, 0, gameTime - deltaGameTime + 1
    End If

    checkDroid psDroid, "src/droid.cpp:531", "recycleDroid", max_check_object_recursion
End Sub

Private Function droidCancelRepairers(ByVal psDroid As Long, pList() As OBJ_LIST) As Long
    Dim numRepairersRemoved As Long, player As Long, isUsOrAlly As Boolean, it As Long, psCurr As Long
    Dim wasRepairingDroid As Boolean, i As Long

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/droid.cpp:719", "droidCancelRepairers", "Null droid": droidCancelRepairers = 0: Exit Function

    numRepairersRemoved = 0

    For player = 0 To MAX_PLAYERS - 1
        isUsOrAlly = aDroid(psDroid).player = player
        If Not isUsOrAlly Then isUsOrAlly = aiCheckAlliances(aDroid(psDroid).player, player)
        If Not isUsOrAlly Then
            GoTo continue_ ' skip non-self and non-allied players
        End If

        ' search the list it's being removed from for fellow droids that are repairing it, and instruct them that they are done repairing
        it = pList(player).head
        Do While it <> 0
            psCurr = lnObj(it)
            If aDroid(psCurr).action = DACTION_DROIDREPAIR Then
                wasRepairingDroid = False
                For i = 0 To MAX_WEAPONS - 1
                    If aDroid(psCurr).psActionTarget(i) = psDroid Then
                        ' found a droid repairing this droid
                        wasRepairingDroid = True
                        setDroidActionTarget psDroid, 0, i
                        ' actionUpdateDroid will handle when DACTION_DROIDREPAIR and the action target is null

                        If i <> 0 Then
                            debug_ 748, LOG_INFO, "droidCancelRepairers", "Found actionTarget[%d] that matches droid", i
                        End If
                    End If
                Next
                If wasRepairingDroid Then
                    droidRepairStopped psDroid, psCurr
                    numRepairersRemoved = numRepairersRemoved + 1
                End If
            End If
            it = lnNext(it)
        Loop

        ' search the main (current) list of active structures to see if any point to this
        it = apsStructLists(player).head
        Do While it <> 0
            psCurr = lnObj(it)
            If asStructureStats(aStruct(psCurr).pStructureType).typ = REF_REPAIR_FACILITY And aStruct(psCurr).pFunctionality <> 0 Then
                If aStruct(psCurr).functionality.repairFacility.state = RepairState_Repairing And aStruct(psCurr).functionality.repairFacility.psObj = psDroid Then
                    ' reproduce logic from aiUpdateRepair_handleEvents - RepairEvents::UnitMovedAway
                    droidRepairStopped psDroid, psCurr
                    aStruct(psCurr).functionality.repairFacility.psObj = 0
                    aStruct(psCurr).functionality.repairFacility.state = RepairState_Idle
                    numRepairersRemoved = numRepairersRemoved + 1
                End If
            End If
            it = lnNext(it)
        Loop
continue_:
    Next

    syncDebugObject_ "droidCancelRepairers", psDroid, Asc("r")

    If Not (aDroid(psDroid).underRepair = 0) Then ASSERT_FAILURE "psDroid->underRepair == 0", "src/droid.cpp:781", "droidCancelRepairers", "DROID.underRepair = %u", aDroid(psDroid).underRepair
    droidCancelRepairers = numRepairersRemoved
End Function

' Remove a droid from the List so doesn't update or get drawn etc
' TAKE CARE with removeDroid() - usually want droidRemove since it deal with grid code
' returns false if the droid wasn't removed - because it died!
Public Function droidRemove(ByVal psDroid As Long, pList() As OBJ_LIST) As Boolean
    checkDroid psDroid, "src/droid.cpp:790", "droidRemove", max_check_object_recursion

    If isDead(psDroid) Then
        ' droid has already been killed, quit
        droidRemove = False
        Exit Function
    End If

    ' leave the current formation if any
    If aDroid(psDroid).sMove.psFormation <> 0 Then
        formationLeave aDroid(psDroid).sMove.psFormation, psDroid
        aDroid(psDroid).sMove.psFormation = 0
    End If

    ' leave the current group if any - not if its a Transporter droid
    If Not DROID_isTransporter(psDroid) And aDroid(psDroid).psGroup <> 0 Then
        DROID_GROUP_remove aDroid(psDroid).psGroup, psDroid
        aDroid(psDroid).psGroup = 0
    End If

    ' instruct other droids / structures that are repairing this droid to stop
    If aDroid(psDroid).underRepair > 0 Then
        droidCancelRepairers psDroid, pList
    End If

    ' reset the baseStruct
    setDroidBase psDroid, 0

    removeDroid psDroid, pList

    If aDroid(psDroid).player = selectedPlayer Then
        intInformInterfaceObjectRemoved psDroid
        intRefreshScreen
    End If

    droidRemove = True
End Function

' The main update routine for all droids
Public Sub droidUpdate(ByVal psDroid As Long)
    Dim dv As Vector3i, percentDamage As Long, emissionInterval As Long, psBeingTargetted As Long, i As Long
    Dim baseImd As Long, imd As Long, effectTime As Long, cond As Boolean, a As Double, b As Double

    checkDroid psDroid, "src/droid.cpp:869", "droidUpdate", max_check_object_recursion

    ' (#ifdef DEBUG: the sensor list check)

    syncDebugDroid_ "droidUpdate", psDroid, Asc("<")

    If (aDroid(psDroid).flags And BIT_(OBJECT_FLAG_DIRTY)) <> 0 Then
        visTilesUpdate psDroid
        droidBodyUpgrade psDroid
        aDroid(psDroid).flags = aDroid(psDroid).flags And Not BIT_(OBJECT_FLAG_DIRTY)
    End If

    ' Save old droid position, update time.
    aDroid(psDroid).prevSpacetime = getSpacetime(psDroid)
    aDroid(psDroid).time_ = gameTime
    For i = 0 To MAX_(1, aDroid(psDroid).numWeaps) - 1
        aDroid(psDroid).asWeaps(i).prevRot = aDroid(psDroid).asWeaps(i).rot
    Next

    If aDroid(psDroid).animationEvent <> ANIM_EVENT_NONE Then
        baseImd = 0
        If aDroid(psDroid).sDisplay.imd <> 0 Then baseImd = iIMDShape_objanimpie(iIMDBaseShape_displayModel(aDroid(psDroid).sDisplay.imd), aDroid(psDroid).animationEvent)
        imd = 0
        If baseImd <> 0 Then imd = iIMDBaseShape_displayModel(baseImd)
        cond = False
        If imd <> 0 Then
            If iIMDShape_objanimcycles(imd) > 0 Then
                cond = ULT(FromU32(CDbl(aDroid(psDroid).timeAnimationStarted) + CDbl(iIMDShape_objanimtime(imd)) * iIMDShape_objanimcycles(imd)), gameTime)
            End If
        End If
        If cond Then
            ' Done animating (animation is defined by body - other components should follow suit)
            If aDroid(psDroid).animationEvent = ANIM_EVENT_DYING Then
                ' case theoretically should not happen anymore, as death animations are handled separately as effects now?
                debug_ 908, LOG_DEATH, "droidUpdate", "%s (%d) died to burn anim (died=%d)", objInfo(psDroid), aDroid(psDroid).id, aDroid(psDroid).died
                Exit Sub
            End If
            aDroid(psDroid).animationEvent = ANIM_EVENT_NONE
        End If
    ElseIf aDroid(psDroid).animationEvent = ANIM_EVENT_DYING Then
        Exit Sub ' rest below is irrelevant if dead
    End If

    ' Restore group from repairGroup if the droid gets interrupted while retreating.
    If aDroid(psDroid).repairGroup <> UBYTE_MAX And _
       aDroid(psDroid).order.typ <> DORDER_RTR And _
       aDroid(psDroid).order.typ <> DORDER_RTR_SPECIFIED And _
       aDroid(psDroid).order.typ <> DORDER_RTB Then
        If secondaryGetState(psDroid, DSO_REPAIR_LEVEL) = DSS_REPLEV_NEVER Then
            droidWasFullyRepairedBase psDroid
        End If
    End If

    ' ai update droid
    aiUpdateDroid psDroid

    ' Update the droids order.
    If Not orderUpdateDroid(psDroid) Then
        Exit Sub ' skip further processing - droid was moved to a different list!
    End If

    ' update the action of the droid
    actionUpdateDroid psDroid

    syncDebugDroid_ "droidUpdate", psDroid, Asc("M")

    ' update the move system
    moveUpdateDroid psDroid

    ' Only add smoke if they're visible
    If BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 And aDroid(psDroid).droidType <> DROID_PERSON Then
        ' need to clip this value to prevent overflow condition
        percentDamage = FromU32(Fix(U32(MUL32(aDroid(psDroid).body, 100)) / U32(aDroid(psDroid).originalBody)))  ' PERCENT(body, originalBody)
        If ULT(100, percentDamage) Then percentDamage = 100
        percentDamage = 100 - percentDamage

        ' Is there any damage?
        If percentDamage >= 25 Then
            If percentDamage >= 100 Then
                percentDamage = 99
            End If

            emissionInterval = (((100 - percentDamage) + 10) \ 10) * DROID_DAMAGE_SCALING   ' CALC_DROID_SMOKE_INTERVAL(percentDamage)

            ' std::max of two uint32_t
            a = U32(FromU32(CDbl(gameTime) - deltaGameTime + 1))
            b = U32(FromU32(CDbl(aDroid(psDroid).lastEmission) + emissionInterval))
            If a < b Then effectTime = FromU32(b) Else effectTime = FromU32(a)
            If Not ULT(gameTime, effectTime) Then
                dv.x = aDroid(psDroid).pos.x + (16 - rand() Mod 32)   ' DROID_DAMAGE_SPREAD
                dv.z = aDroid(psDroid).pos.y + (16 - rand() Mod 32)
                dv.y = aDroid(psDroid).pos.z

                dv.y = dv.y + (imdMaxY(aDroid(psDroid).sDisplay.imd) * 2)
                addEffect dv, EFFECT_SMOKE, SMOKE_TYPE_DRIFTING_SMALL, False, 0, 0, effectTime
                aDroid(psDroid).lastEmission = effectTime
            End If
        End If
    End If

    ' -----------------
    ' Are we a sensor droid or a command droid? Show where we target for selectedPlayer.
    If aDroid(psDroid).player = selectedPlayer And (aDroid(psDroid).droidType = DROID_SENSOR Or aDroid(psDroid).droidType = DROID_COMMAND) Then
        ' If we're attacking or sensing (observing), then...
        psBeingTargetted = orderStateObj(psDroid, DORDER_ATTACK)
        If psBeingTargetted = 0 Then psBeingTargetted = orderStateObj(psDroid, DORDER_OBSERVE)
        If psBeingTargetted <> 0 Then
            BASE_OBJECT_flags(psBeingTargetted) = BASE_OBJECT_flags(psBeingTargetted) Or BIT_(OBJECT_FLAG_TARGETED)
        Else
            cond = secondaryGetState(psDroid, DSO_HALTTYPE) <> DSS_HALT_PURSUE And aDroid(psDroid).psActionTarget(0) <> 0
            If cond Then cond = validTarget(psDroid, aDroid(psDroid).psActionTarget(0), 0)
            If cond Then cond = aDroid(psDroid).action = DACTION_ATTACK Or aDroid(psDroid).action = DACTION_OBSERVE Or orderState(psDroid, DORDER_HOLD)
            If cond Then
                psBeingTargetted = aDroid(psDroid).psActionTarget(0)
                BASE_OBJECT_flags(psBeingTargetted) = BASE_OBJECT_flags(psBeingTargetted) Or BIT_(OBJECT_FLAG_TARGETED)
            End If
        End If
    End If
    ' ------------------------
    ' See if we can and need to self repair.
    If Not DROID_isVtol(psDroid) And DROID_isDamaged(psDroid) And aDroid(psDroid).asBits(COMP_REPAIRUNIT) <> 0 Then
        If selfRepairEnabled(aDroid(psDroid).player) Then
            droidUpdateDroidSelfRepair psDroid
        End If
    End If

    If bMultiPlayer Then
        droidUpdateShields psDroid
    End If

    ' Update the fire damage data
    If aDroid(psDroid).periodicalDamageStart <> 0 And aDroid(psDroid).periodicalDamageStart <> FromU32(CDbl(gameTime) - deltaGameTime) Then  ' -deltaGameTime, since projectiles are updated after droids.
        ' The periodicalDamageStart has been set, but is not from the previous tick, so we must be out of the fire.
        aDroid(psDroid).periodicalDamage = 0  ' Reset periodical damage done this tick.
        If ULT(FromU32(CDbl(aDroid(psDroid).periodicalDamageStart) + BURN_TIME), gameTime) Then
            ' Finished periodical damaging.
            aDroid(psDroid).periodicalDamageStart = 0
        Else
            ' do hardcoded burn damage (this damage automatically applied after periodical damage finished)
            droidDamage psDroid, 0, BURN_DAMAGE, WC_HEAT, WSC_FLAME, FromU32(CDbl(gameTime) - deltaGameTime \ 2 + 1), True, BURN_MIN_DAMAGE, False
        End If
    End If

    ' At this point, the droid may be dead due to periodical damage or hardcoded burn damage.
    If isDead(psDroid) Then
        Exit Sub
    End If

    calcDroidIllumination psDroid

    ' Check the resistance level of the droid
    a = U32(FromU32(CDbl(aDroid(psDroid).id) + U32(gameTime)))
    b = U32(FromU32(CDbl(aDroid(psDroid).id) + U32(gameTime) - U32(deltaGameTime)))
    If Int(a / 833) <> Int(b / 833) Then
        ' Zero resistance means not currently been attacked - ignore these
        If aDroid(psDroid).resistance <> 0 Then
            If aDroid(psDroid).resistance < DROID_droidResistance(psDroid) Then
                ' Increase over time if low
                aDroid(psDroid).resistance = aDroid(psDroid).resistance + 1
            End If
        End If
    End If

    syncDebugDroid_ "droidUpdate", psDroid, Asc(">")

    checkDroid psDroid, "src/droid.cpp:1047", "droidUpdate", max_check_object_recursion
End Sub

Public Sub droidUpdateShields(ByVal psDroid As Long)
    ' shields are down (left for experimentation with different implementations)
    aDroid(psDroid).shieldPoints = -1
End Sub

' See if a droid is next to a structure
Private Function droidNextToStruct(ByVal psDroid As Long, ByVal psStruct As Long) As Boolean
    Dim posX As Long, posY As Long, minX As Long, minY As Long, maxX As Long, maxY As Long, x As Long, y As Long

    checkDroid psDroid, "src/droid.cpp:1093", "droidNextToStruct", max_check_object_recursion

    posX = map_coord(aDroid(psDroid).pos.x)
    posY = map_coord(aDroid(psDroid).pos.y)
    minX = MAX_(posX - 1, 0)
    minY = MAX_(posY - 1, 0)
    maxX = MIN_(posX + 1, mapWidth)
    maxY = MIN_(posY + 1, mapHeight)
    For y = minY To maxY
        For x = minX To maxX
            If TileHasStructure(mapTile(x, y)) Then
                If getTileStructure(x, y) = psStruct Then
                    droidNextToStruct = True
                    Exit Function
                End If
            End If
        Next
    Next

    droidNextToStruct = False
End Function

' (VB) the AUDIO_CALLBACK droidCheckBuildStillInProgress
Public Function droidCheckBuildStillInProgress(ByVal psObj As Long) As Boolean
    Dim psDroid As Long
    If psObj = 0 Then
        droidCheckBuildStillInProgress = False
        Exit Function
    End If

    psDroid = psObj
    checkDroid psDroid, "src/droid.cpp:1123", "droidCheckBuildStillInProgress", max_check_object_recursion

    droidCheckBuildStillInProgress = aDroid(psDroid).died = 0 And aDroid(psDroid).action = DACTION_BUILD
End Function

' (VB) the AUDIO_CALLBACK droidBuildStartAudioCallback
Public Function droidBuildStartAudioCallback(ByVal psObj As Long) As Boolean
    Dim psDroid As Long
    psDroid = psObj

    If psDroid <> 0 Then
        If BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 Then
            audio_PlayObjDynamicTrack psDroid, ID_SOUND_CONSTRUCTION_LOOP, CB_droidCheckBuildStillInProgress
        End If
    End If

    droidBuildStartAudioCallback = True
End Function

' Callback function for stopped audio tracks (the AUDIO_CALLBACK droidAudioTrackStopped)
'  Sets the droid's current track id to NO_SOUND
'  \return true on success, false on failure
Public Function droidAudioTrackStopped(ByVal psObj As Long) As Boolean
    Dim psDroid As Long

    psDroid = psObj
    If psDroid = 0 Then
        debug_ 3838, LOG_ERROR, "droidAudioTrackStopped", "droid pointer invalid"
        droidAudioTrackStopped = False
        Exit Function
    End If

    ' (VB) psDroid->type != OBJ_DROID || psDroid->died, in that order
    If SIMPLE_OBJECT_type(psDroid) <> OBJ_DROID Then
        droidAudioTrackStopped = False
        Exit Function
    End If
    If aDroid(psDroid).died <> 0 Then
        droidAudioTrackStopped = False
        Exit Function
    End If

    aDroid(psDroid).iAudioID = NO_SOUND
    droidAudioTrackStopped = True
End Function

' Set up a droid to build a structure - returns true if successful
Public Function droidStartBuild(ByVal psDroid As Long) As DroidStartBuild_
    Dim psStruct As Long, psStructStat As Long, ia As Long, id As Long, cond As Boolean

    psStruct = 0
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/droid.cpp:1145", "droidStartBuild", "Bad Droid": droidStartBuild = DroidStartBuildFailed: Exit Function
    checkDroid psDroid, "src/droid.cpp:1146", "droidStartBuild", max_check_object_recursion

    ' See if we are starting a new structure
    If aDroid(psDroid).order.psObj = 0 And _
       (aDroid(psDroid).order.typ = DORDER_BUILD Or _
        aDroid(psDroid).order.typ = DORDER_LINEBUILD) Then
        psStructStat = aDroid(psDroid).order.psStats

        ia = apStructTypeLists(aDroid(psDroid).player, psStructStat)
        If ia <> AVAILABLE And ia <> REDUNDANT Then
            ASSERT_FAILURE "false", "src/droid.cpp:1158", "droidStartBuild", "Cannot build ""%s"" for player %d.", asStructureStats(psStructStat).name_, aDroid(psDroid).player
            cancelBuild psDroid
            objTrace aDroid(psDroid).id, "droidStartBuild", "DroidStartBuildFailed: not researched"
            droidStartBuild = DroidStartBuildFailed
            Exit Function
        End If

        'need to check structLimits have not been exceeded
        If Not ULT(asStructureStats(psStructStat).curCount(aDroid(psDroid).player), asStructureStats(psStructStat).upgrade(aDroid(psDroid).player).limit) Then
            cancelBuild psDroid
            objTrace aDroid(psDroid).id, "droidStartBuild", "DroidStartBuildFailed: structure limits"
            droidStartBuild = DroidStartBuildFailed
            Exit Function
        End If
        ' Can't build on burning oil derricks.
        cond = asStructureStats(psStructStat).typ = REF_RESOURCE_EXTRACTOR
        If cond Then cond = fireOnLocation(aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y)
        If cond Then
            ' Don't cancel build, since we can wait for it to stop burning.
            objTrace aDroid(psDroid).id, "droidStartBuild", "DroidStartBuildPending: burning"
            droidStartBuild = DroidStartBuildPending
            Exit Function
        End If
        'ok to build
        id = generateSynchronisedObjectId()
        psStruct = buildStructureDir(psStructStat, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y, aDroid(psDroid).order.direction, aDroid(psDroid).player, False, id)
        If psStruct = 0 Then
            cancelBuild psDroid
            objTrace aDroid(psDroid).id, "droidStartBuild", "DroidStartBuildFailed: buildStructureDir failed"
            droidStartBuild = DroidStartBuildFailed
            Exit Function
        End If
        aStruct(psStruct).body = (aStruct(psStruct).body + 9) \ 10  ' Structures start at 10% health. Round up.
    Else
        ' Check the structure is still there to build (joining a partially built struct)
        psStruct = castStructure(aDroid(psDroid).order.psObj)
        If psStruct = 0 Then
            psStruct = castStructure(psMapTiles(worldTile(aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y)).psObject)
        End If
        If psStruct <> 0 Then
            If Not droidNextToStruct(psDroid, psStruct) Then
                ' Nope - stop building
                debug_ 1200, LOG_NEVER, "droidStartBuild", "not next to structure"
                objTrace aDroid(psDroid).id, "droidStartBuild", "DroidStartBuildSuccess: not next to structure"
            End If
        End If
    End If

    'check structure not already built, and we still 'own' it
    If psStruct <> 0 Then
        If aStruct(psStruct).status <> SS_BUILT And aiCheckAlliances(aStruct(psStruct).player, aDroid(psDroid).player) Then
            aDroid(psDroid).actionStarted = gameTime
            aDroid(psDroid).actionPoints = 0
            setDroidTarget psDroid, psStruct
            setDroidActionTarget psDroid, psStruct, 0
            objTrace aDroid(psDroid).id, "droidStartBuild", "DroidStartBuild: set target"
        End If

        If BASE_OBJECT_visibleForLocalDisplay(psStruct) <> 0 Then
            audio_PlayObjStaticTrackCallback psDroid, ID_SOUND_CONSTRUCTION_START, CB_droidBuildStartAudioCallback
        End If
    End If
    checkDroid psDroid, "src/droid.cpp:1222", "droidStartBuild", max_check_object_recursion

    objTrace aDroid(psDroid).id, "droidStartBuild", "DroidStartBuildSuccess"
    droidStartBuild = DroidStartBuildSuccess
End Function

Private Sub droidAddWeldSound(iVecEffect As Vector3i)
    Dim iAudioID As Long
    iAudioID = ID_SOUND_CONSTRUCTION_1 + (rand() Mod 4)

    audio_PlayStaticTrack iVecEffect.x, iVecEffect.z, iAudioID
End Sub

Private Sub addConstructorEffect(ByVal psStruct As Long)
    Dim size As Vector2i, temp As Vector3i
    If (rand() Mod 10 = 0) Then   ' ONEINTEN
        If BASE_OBJECT_visibleForLocalDisplay(psStruct) <> 0 Then
            ' This needs fixing - it's an arse effect!
            size = STRUCTURE_size(psStruct)
            size.x = size.x * TILE_UNITS \ 4
            size.y = size.y * TILE_UNITS \ 4
            If Not (size.x > 0 And size.y > 0) Then ASSERT_FAILURE "size.x > 0 && size.y > 0", "src/droid.cpp:1241", "addConstructorEffect", "Zero-size building?: %s", asStructureStats(aStruct(psStruct).pStructureType).id: Exit Sub
            temp.x = aStruct(psStruct).pos.x + ((rand() Mod (2 * size.x)) - size.x)
            temp.y = map_TileHeight(map_coord(aStruct(psStruct).pos.x), map_coord(aStruct(psStruct).pos.y)) + (imdMaxY(aStruct(psStruct).sDisplay.imd) \ 6)
            temp.z = aStruct(psStruct).pos.y + ((rand() Mod (2 * size.y)) - size.y)
            If (rand() Mod 2) <> 0 Then
                droidAddWeldSound temp
            End If
        End If
    End If
End Sub

' Update a construction droid while it is building
' returns true while building continues
Public Function droidUpdateBuild(ByVal psDroid As Long) As Boolean
    Dim psStruct As Long, distanceSq As Long, constructPoints As Long, pointsToAdd As Long, buildPointsToAdd As Long

    checkDroid psDroid, "src/droid.cpp:1257", "droidUpdateBuild", max_check_object_recursion
    If Not (aDroid(psDroid).action = DACTION_BUILD) Then ASSERT_FAILURE "psDroid->action == DACTION_BUILD", "src/droid.cpp:1258", "droidUpdateBuild", "%s (order %s) has wrong action for construction: %s", _
                     droidGetName(psDroid), getDroidOrderName(aDroid(psDroid).order.typ), getDroidActionName(aDroid(psDroid).action): droidUpdateBuild = False: Exit Function

    psStruct = castStructure(aDroid(psDroid).order.psObj)
    If psStruct = 0 Then
        ' Target missing, stop trying to build it.
        aDroid(psDroid).action = DACTION_NONE
        droidUpdateBuild = False
        Exit Function
    End If

    If Not (aStruct(psStruct).typ = OBJ_STRUCTURE) Then ASSERT_FAILURE "psStruct->type == OBJ_STRUCTURE", "src/droid.cpp:1269", "droidUpdateBuild", "target is not a structure": droidUpdateBuild = False: Exit Function
    If Not (aDroid(psDroid).asBits(COMP_CONSTRUCT) < asConstructStats_size) Then ASSERT_FAILURE "psDroid->asBits[COMP_CONSTRUCT] < asConstructStats.size()", "src/droid.cpp:1270", "droidUpdateBuild", "Invalid construct pointer for unit": droidUpdateBuild = False: Exit Function

    ' droid distance sanity check
    distanceSq = droidSqDist(psDroid, psStruct)
    If ULT(REPAIR_MAXDIST * REPAIR_MAXDIST, distanceSq) Then
        aDroid(psDroid).action = DACTION_NONE
        droidUpdateBuild = False
        Exit Function
    End If

    ' First check the structure hasn't been completed by another droid
    If aStruct(psStruct).status = SS_BUILT Then
        ' Check if line order build is completed, or we are not carrying out a line order build
        If aDroid(psDroid).order.typ <> DORDER_LINEBUILD Or _
           (map_coord(aDroid(psDroid).order.pos.x) = map_coord(aDroid(psDroid).order.pos2.x) And map_coord(aDroid(psDroid).order.pos.y) = map_coord(aDroid(psDroid).order.pos2.y)) Then
            cancelBuild psDroid
        Else
            aDroid(psDroid).action = DACTION_NONE ' make us continue line build
            setDroidTarget psDroid, 0
            setDroidActionTarget psDroid, 0, 0
        End If
        droidUpdateBuild = False
        Exit Function
    End If

    ' make sure we still 'own' the building in question
    If Not aiCheckAlliances(aStruct(psStruct).player, aDroid(psDroid).player) Then
        cancelBuild psDroid ' stop what you are doing fool it isn't ours anymore!
        droidUpdateBuild = False
        Exit Function
    End If

    constructPoints = constructorPoints(asConstructStats(DROID_getConstructStats(psDroid)), aDroid(psDroid).player)

    pointsToAdd = FromU32(Fix(U32(MUL32(constructPoints, FromU32(CDbl(gameTime) - aDroid(psDroid).actionStarted))) / GAME_TICKS_PER_SEC))
    buildPointsToAdd = FromU32(CDbl(pointsToAdd) - aDroid(psDroid).actionPoints)

    structureBuild psStruct, psDroid, MAX_(1, buildPointsToAdd), constructPoints

    'store the amount just added
    aDroid(psDroid).actionPoints = pointsToAdd

    addConstructorEffect psStruct

    droidUpdateBuild = True
End Function

Public Function droidUpdateDemolishing(ByVal psDroid As Long) As Boolean
    Dim psStruct As Long, distanceSq As Long, constructRate As Long, pointsToAdd As Long

    checkDroid psDroid, "src/droid.cpp:1322", "droidUpdateDemolishing", max_check_object_recursion

    If Not (aDroid(psDroid).action = DACTION_DEMOLISH) Then ASSERT_FAILURE "psDroid->action == DACTION_DEMOLISH", "src/droid.cpp:1324", "droidUpdateDemolishing", "unit is not demolishing": droidUpdateDemolishing = False: Exit Function
    psStruct = aDroid(psDroid).order.psObj
    If Not (SIMPLE_OBJECT_type(psStruct) = OBJ_STRUCTURE) Then ASSERT_FAILURE "psStruct->type == OBJ_STRUCTURE", "src/droid.cpp:1326", "droidUpdateDemolishing", "target is not a structure": droidUpdateDemolishing = False: Exit Function

    ' droid distance sanity check
    distanceSq = droidSqDist(psDroid, psStruct)
    If ULT(REPAIR_MAXDIST * REPAIR_MAXDIST, distanceSq) Then
        aDroid(psDroid).action = DACTION_NONE
        droidUpdateDemolishing = False
        Exit Function
    End If

    constructRate = 5 * constructorPoints(asConstructStats(DROID_getConstructStats(psDroid)), aDroid(psDroid).player)
    pointsToAdd = gameTimeAdjustedAverage(constructRate)

    structureDemolish psStruct, psDroid, pointsToAdd

    addConstructorEffect psStruct

    checkDroid psDroid, "src/droid.cpp:1343", "droidUpdateDemolishing", max_check_object_recursion

    droidUpdateDemolishing = True
End Function

Public Sub droidStartAction(ByVal psDroid As Long)
    aDroid(psDroid).actionStarted = gameTime
    aDroid(psDroid).actionPoints = 0
End Sub

' continue restoring a structure
Public Function droidUpdateRestore(ByVal psDroid As Long) As Boolean
    Dim psStruct As Long, compIndex As Long, psStats As Long, distanceSq As Long, restorePoints As Long, pointsToAdd As Long

    checkDroid psDroid, "src/droid.cpp:1357", "droidUpdateRestore", max_check_object_recursion

    If Not (aDroid(psDroid).action = DACTION_RESTORE) Then ASSERT_FAILURE "psDroid->action == DACTION_RESTORE", "src/droid.cpp:1359", "droidUpdateRestore", "Unit is not restoring": droidUpdateRestore = False: Exit Function
    psStruct = aDroid(psDroid).order.psObj
    If Not (SIMPLE_OBJECT_type(psStruct) = OBJ_STRUCTURE) Then ASSERT_FAILURE "psStruct->type == OBJ_STRUCTURE", "src/droid.cpp:1361", "droidUpdateRestore", "Target is not a structure": droidUpdateRestore = False: Exit Function
    If Not (aDroid(psDroid).asWeaps(0).nStat <> 0) Then ASSERT_FAILURE "psDroid->asWeaps[0].nStat > 0", "src/droid.cpp:1362", "droidUpdateRestore", "Droid doesn't have any weapons": droidUpdateRestore = False: Exit Function

    compIndex = aDroid(psDroid).asWeaps(0).nStat
    If Not (ULT(compIndex, asWeaponStats_size)) Then ASSERT_FAILURE "compIndex < asWeaponStats.size()", "src/droid.cpp:1365", "droidUpdateRestore", "Invalid range referenced for numWeaponStats, %u > %zu", compIndex, asWeaponStats_size: droidUpdateRestore = False: Exit Function
    psStats = compIndex

    If Not (asWeaponStats(psStats).weaponSubClass = WSC_ELECTRONIC) Then ASSERT_FAILURE "psStats->weaponSubClass == WSC_ELECTRONIC", "src/droid.cpp:1368", "droidUpdateRestore", "unit's weapon is not EW": droidUpdateRestore = False: Exit Function

    ' droid distance sanity check
    distanceSq = droidSqDist(psDroid, psStruct)
    If ULT(REPAIR_MAXDIST * REPAIR_MAXDIST, distanceSq) Then
        aDroid(psDroid).action = DACTION_NONE
        droidUpdateRestore = False
        Exit Function
    End If

    restorePoints = calcDamage(weaponDamage(asWeaponStats(psStats), aDroid(psDroid).player), _
                               asWeaponStats(psStats).weaponEffect, psStruct)

    pointsToAdd = FromU32(Fix(U32(MUL32(restorePoints, FromU32(CDbl(gameTime) - aDroid(psDroid).actionStarted))) / _
                  GAME_TICKS_PER_SEC))

    aStruct(psStruct).resistance = ToInt16(FromU32(CDbl(aStruct(psStruct).resistance) + U32(FromU32(U32(pointsToAdd) - U32(aDroid(psDroid).actionPoints)))))

    'store the amount just added
    aDroid(psDroid).actionPoints = pointsToAdd

    checkDroid psDroid, "src/droid.cpp:1389", "droidUpdateRestore", max_check_object_recursion

    ' check if structure is restored
    If aStruct(psStruct).resistance < structureResistance(aStruct(psStruct).pStructureType, aStruct(psStruct).player) Then
        droidUpdateRestore = True
    Else
        addConsoleMessage "Structure Restored", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        aStruct(psStruct).resistance = structureResistance(aStruct(psStruct).pStructureType, _
                                       aStruct(psStruct).player) And &HFFFF&
        droidUpdateRestore = False
    End If
End Function

Public Sub droidRepairStarted(ByVal psDroid As Long, ByVal psRepairer As Long)
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/droid.cpp:1426", "droidRepairStarted", "Null droid": Exit Sub
    If Not (aDroid(psDroid).underRepair < UWORD_MAX) Then ASSERT_FAILURE "psDroid->underRepair < std::numeric_limits<decltype(psDroid->underRepair)>::max()", "src/droid.cpp:1427", "droidRepairStarted", "underRepair hit max for droid: [%d] %s", aDroid(psDroid).id, objInfo(psDroid): Exit Sub
    ' (#if defined(DEBUG): the LOG_NEVER line)
    aDroid(psDroid).underRepair = aDroid(psDroid).underRepair + 1
End Sub

Public Sub droidRepairStopped(ByVal psDroid As Long, ByVal psFormerRepairer As Long)
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/droid.cpp:1436", "droidRepairStopped", "Null droid": Exit Sub
    ' (#if defined(DEBUG): the LOG_NEVER line)
    If aDroid(psDroid).underRepair = 0 Then
        ' This is expected after initially loading pre-4.6.0 saves (WZ did not calculate underRepair until version 4.6)
        ' But should not happen when playing new games
        debug_ 1444, LOG_WARNING, "droidRepairStopped", "underRepair count already 0 for droid: [%d] %s", aDroid(psDroid).id, objInfo(psDroid)
        Exit Sub
    End If
    aDroid(psDroid).underRepair = aDroid(psDroid).underRepair - 1
End Sub

' Droid was completely repaired by another droid, auto-repair, or repair facility
Public Sub droidWasFullyRepairedBase(ByVal psDroid As Long)
    If aDroid(psDroid).repairGroup <> UBYTE_MAX Then
        aDroid(psDroid).group = aDroid(psDroid).repairGroup
        aDroid(psDroid).repairGroup = UBYTE_MAX
        SelectGroupDroid psDroid
        intGroupsChanged aDroid(psDroid).group ' update groups UI
    End If
End Sub

' psRepairFac: a const REPAIR_FACILITY *, the repair facility structure's pFunctionality (0: nullptr)
Public Sub droidWasFullyRepaired(ByVal psDroid As Long, ByVal psRepairFac As Long)
    Dim prevWasRTR As Boolean, dp As Long
    prevWasRTR = aDroid(psDroid).order.typ = DORDER_RTR Or aDroid(psDroid).order.typ = DORDER_RTR_SPECIFIED
    dp = 0
    If psRepairFac <> 0 Then dp = aStruct(psRepairFac).functionality.repairFacility.psDeliveryPoint
    If prevWasRTR And hasCommander(psDroid) Then
        objTrace aDroid(psDroid).id, "droidWasFullyRepaired", "Repair complete - move to commander"
        orderDroidObj psDroid, DORDER_GUARD, aGroup(aDroid(psDroid).psGroup).psCommander, ModeImmediate
    ElseIf prevWasRTR And psRepairFac <> 0 And dp <> 0 Then
        objTrace aDroid(psDroid).id, "droidWasFullyRepaired", "Repair complete - move to delivery point"
        ' ModeQueue because delivery points are not yet synchronised!
        orderDroidLoc psDroid, DORDER_MOVE, aFlagPos(dp).coords.x, aFlagPos(dp).coords.y, ModeQueue
    ElseIf prevWasRTR Then
        ' nothing to do, no commander, no repair point to go to. Stop, and guard this place.
        aDroid(psDroid).order.psObj = 0
        objTrace aDroid(psDroid).id, "droidWasFullyRepaired", "Repair complete - guarding the place at x=%i y=%i", aDroid(psDroid).pos.x, aDroid(psDroid).pos.y
        orderDroidLoc psDroid, DORDER_GUARD, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, ModeImmediate
    End If

    droidWasFullyRepairedBase psDroid
End Sub

Public Function droidUpdateRepair(ByVal psDroid As Long) As Boolean
    Dim psStruct As Long, iRepairRate As Long

    checkDroid psDroid, "src/droid.cpp:1489", "droidUpdateRepair", max_check_object_recursion

    If Not (aDroid(psDroid).action = DACTION_REPAIR) Then ASSERT_FAILURE "psDroid->action == DACTION_REPAIR", "src/droid.cpp:1491", "droidUpdateRepair", "unit does not have repair order": droidUpdateRepair = False: Exit Function
    psStruct = aDroid(psDroid).psActionTarget(0)

    If Not (SIMPLE_OBJECT_type(psStruct) = OBJ_STRUCTURE) Then ASSERT_FAILURE "psStruct->type == OBJ_STRUCTURE", "src/droid.cpp:1494", "droidUpdateRepair", "target is not a structure": droidUpdateRepair = False: Exit Function
    iRepairRate = constructorPoints(asConstructStats(DROID_getConstructStats(psDroid)), aDroid(psDroid).player)

    ' add points to structure
    structureRepair psStruct, psDroid, iRepairRate

    ' if not finished repair return true else complete repair and return false
    If ULT(aStruct(psStruct).body, STRUCTURE_structureBody(psStruct)) Then
        droidUpdateRepair = True
    Else
        objTrace aDroid(psDroid).id, "droidUpdateRepair", "Repaired of %s all done with %u", objInfo(psStruct), iRepairRate
        droidUpdateRepair = False
    End If
End Function

' Updates a Repair Droid working on a damaged droid
Private Function droidUpdateDroidRepairBase(ByVal psRepairDroid As Long, ByVal psDroidToRepair As Long) As Boolean
    Dim repairStats As Long, iRepairRateNumerator As Long, iRepairRateDenominator As Long, iPointsToAdd As Long
    Dim iVecEffect As Vector3i, needMoreRepair As Boolean, v As Long, sx As Long, sy As Long, sz As Long

    checkDroid psRepairDroid, "src/droid.cpp:1515", "droidUpdateDroidRepairBase", max_check_object_recursion

    repairStats = DROID_getRepairStats(psRepairDroid)
    iRepairRateNumerator = repairPoints(asRepairStats(repairStats), aDroid(psRepairDroid).player)
    iRepairRateDenominator = 1

    'if self repair then add repair points depending on the time delay for the stat
    If psRepairDroid = psDroidToRepair Then
        iRepairRateNumerator = iRepairRateNumerator * GAME_TICKS_PER_SEC
        iRepairRateDenominator = iRepairRateDenominator * asRepairStats(repairStats).time_
    End If

    iPointsToAdd = gameTimeAdjustedAverage(iRepairRateNumerator, iRepairRateDenominator)

    ' clip<UDWORD>(body + iPointsToAdd, 0, originalBody)
    v = FromU32(CDbl(aDroid(psDroidToRepair).body) + iPointsToAdd)
    If ULT(aDroid(psDroidToRepair).originalBody, v) Then v = aDroid(psDroidToRepair).originalBody
    aDroid(psDroidToRepair).body = v

    ' add plasma repair effect whilst being repaired
    If (rand() Mod 5 = 0) Then   ' ONEINFIVE
        If BASE_OBJECT_visibleForLocalDisplay(psDroidToRepair) <> 0 Then
            sx = 20 - rand() Mod 40   ' DROID_REPAIR_SPREAD (the arguments taken left to right)
            sy = 20 - rand() Mod 40
            sz = rand() Mod 8
            iVecEffect.x = aDroid(psDroidToRepair).pos.x + sx
            iVecEffect.y = aDroid(psDroidToRepair).pos.z + sz
            iVecEffect.z = aDroid(psDroidToRepair).pos.y + sy
            effectGiveAuxVar 90 + (rand() Mod 20)
            addEffect iVecEffect, EFFECT_EXPLOSION, EXPLOSION_TYPE_LASER, False, 0, 0, gameTime - deltaGameTime + 1 + (rand() Mod deltaGameTime)
            droidAddWeldSound iVecEffect
        End If
    End If

    checkDroid psRepairDroid, "src/droid.cpp:1541", "droidUpdateDroidRepairBase", max_check_object_recursion
    ' if not finished repair return true else complete repair and return false
    needMoreRepair = ULT(aDroid(psDroidToRepair).body, aDroid(psDroidToRepair).originalBody)
    If Not needMoreRepair Then
        droidWasFullyRepairedBase psDroidToRepair
    End If
    droidUpdateDroidRepairBase = needMoreRepair
End Function

Public Function droidUpdateDroidRepair(ByVal psRepairDroid As Long) As Boolean
    Dim psDroidToRepair As Long, needMoreRepair As Boolean

    If Not (aDroid(psRepairDroid).action = DACTION_DROIDREPAIR) Then ASSERT_FAILURE "psRepairDroid->action == DACTION_DROIDREPAIR", "src/droid.cpp:1553", "droidUpdateDroidRepair", "Unit does not have unit repair order": droidUpdateDroidRepair = False: Exit Function
    If Not (aDroid(psRepairDroid).asBits(COMP_REPAIRUNIT) <> 0) Then ASSERT_FAILURE "psRepairDroid->asBits[COMP_REPAIRUNIT] != 0", "src/droid.cpp:1554", "droidUpdateDroidRepair", "Unit does not have a repair turret": droidUpdateDroidRepair = False: Exit Function

    psDroidToRepair = aDroid(psRepairDroid).psActionTarget(0)
    If Not (SIMPLE_OBJECT_type(psDroidToRepair) = OBJ_DROID) Then ASSERT_FAILURE "psDroidToRepair->type == OBJ_DROID", "src/droid.cpp:1557", "droidUpdateDroidRepair", "Target is not a unit": droidUpdateDroidRepair = False: Exit Function
    needMoreRepair = droidUpdateDroidRepairBase(psRepairDroid, psDroidToRepair)
    If needMoreRepair And aDroid(psDroidToRepair).order.typ = DORDER_RTR And aDroid(psDroidToRepair).order.rtrType = RTR_TYPE_DROID And aDroid(psDroidToRepair).action = DACTION_NONE Then
        aDroid(psDroidToRepair).action = DACTION_WAITDURINGREPAIR
    End If
    If Not needMoreRepair And aDroid(psDroidToRepair).order.typ = DORDER_RTR And aDroid(psDroidToRepair).order.rtrType = RTR_TYPE_DROID Then
        ' if psDroidToRepair has a commander, commander will call him back anyway
        ' if no commanders, just DORDER_GUARD the repair turret
        orderDroidObj psDroidToRepair, DORDER_GUARD, psRepairDroid, ModeImmediate
        secondarySetState psDroidToRepair, DSO_RETURN_TO_LOC, DSS_NONE
        aDroid(psDroidToRepair).order.psObj = 0
    End If
    droidUpdateDroidRepair = needMoreRepair
End Function

Private Sub droidUpdateDroidSelfRepair(ByVal psRepairDroid As Long)
    droidUpdateDroidRepairBase psRepairDroid, psRepairDroid
End Sub

' Sets the parts array in a template given a droid.
Public Sub templateSetParts(ByVal psDroid As Long, psTemplate As DROID_TEMPLATE)
    Dim inc As Long
    psTemplate.numWeaps = 0
    psTemplate.droidType = aDroid(psDroid).droidType
    For inc = 0 To MAX_WEAPONS - 1
        'this should fix the NULL weapon stats for empty weaponslots
        psTemplate.asWeaps(inc) = 0
        If aDroid(psDroid).asWeaps(inc).nStat <> 0 Then
            psTemplate.numWeaps = psTemplate.numWeaps + 1
            psTemplate.asWeaps(inc) = aDroid(psDroid).asWeaps(inc).nStat
        End If
    Next
    For inc = 0 To DROID_MAXCOMP - 1
        psTemplate.asParts(inc) = aDroid(psDroid).asBits(inc)
    Next
End Sub

' calculate muzzle base location in 3d world
Public Function calcDroidMuzzleBaseLocation(ByVal psDroid As Long, muzzle As Vector3i, ByVal weapon_slot As Long) As Boolean
    Dim psBodyImd As Long, barrel As Vector3i, af As Affine3F, cx As Long, cy As Long, cz As Long, r As Vector3i, ok As Boolean

    psBodyImd = asBodyStats(DROID_getBodyStats(psDroid)).pIMD    ' BODY_IMD(psDroid, psDroid->player)

    checkDroid psDroid, "src/droid.cpp:2370", "calcDroidMuzzleBaseLocation", max_check_object_recursion

    ok = psBodyImd <> 0
    If ok Then ok = ULT(weapon_slot, iIMDBaseShape_connectors_size(psBodyImd))
    If ok Then
        barrel.x = 0: barrel.y = 0: barrel.z = 0

        Affine3F_construct af

        Affine3F_Trans af, aDroid(psDroid).pos.x, -aDroid(psDroid).pos.z, aDroid(psDroid).pos.y

        'matrix = the center of droid
        Affine3F_RotY af, aDroid(psDroid).rot.direction
        Affine3F_RotX af, aDroid(psDroid).rot.pitch
        Affine3F_RotZ af, -aDroid(psDroid).rot.roll
        iIMDBaseShape_connector psBodyImd, weapon_slot, cx, cy, cz
        Affine3F_Trans af, cx, -cz, -cy 'note y and z flipped

        r = Affine3F_mul_vec(af, barrel)
        muzzle.x = r.x  ' xzy()
        muzzle.y = r.z
        muzzle.z = r.y
        muzzle.z = -muzzle.z
    Else
        muzzle.x = aDroid(psDroid).pos.x
        muzzle.y = aDroid(psDroid).pos.y
        muzzle.z = aDroid(psDroid).pos.z + imdMaxY(aDroid(psDroid).sDisplay.imd)
    End If

    checkDroid psDroid, "src/droid.cpp:2395", "calcDroidMuzzleBaseLocation", max_check_object_recursion

    calcDroidMuzzleBaseLocation = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' returns true when no droid on x,y square.
Public Function noDroid(ByVal x As Long, ByVal y As Long) As Boolean
    Dim i As Long, it As Long, psDroid As Long

    ' check each droid list
    For i = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            If map_coord(aDroid(psDroid).pos.x) = x _
               And map_coord(aDroid(psDroid).pos.y) = y Then
                noDroid = False
                Exit Function
            End If
            it = lnNext(it)
        Loop
    Next
    noDroid = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' returns true when at most one droid on x,y square.
Private Function oneDroidMax(ByVal x As Long, ByVal y As Long) As Boolean
    Dim i As Long, bFound As Boolean, it As Long, pD As Long
    bFound = False
    ' check each droid list
    For i = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(i).head
        Do While it <> 0
            pD = lnObj(it)
            If map_coord(aDroid(pD).pos.x) = x _
               And map_coord(aDroid(pD).pos.y) = y Then
                If bFound Then
                    oneDroidMax = False
                    Exit Function
                End If

                bFound = True 'first droid on this square so continue
            End If
            it = lnNext(it)
        Loop
    Next
    oneDroidMax = True
End Function

' ////////////////////////////////////////////////////////////////////////////
' returns true if it's a sensible place to put that droid.
Private Function sensiblePlace(ByVal x As Long, ByVal y As Long, ByVal propulsion As Long) As Boolean
    ' not too near the edges.
    If (x < TOO_NEAR_EDGE) Or (x > (mapWidth - TOO_NEAR_EDGE)) Then
        sensiblePlace = False
        Exit Function
    End If
    If (y < TOO_NEAR_EDGE) Or (y > (mapHeight - TOO_NEAR_EDGE)) Then
        sensiblePlace = False
        Exit Function
    End If

    ' not on a blocking tile.
    If fpathBlockingTile(x, y, propulsion) Then
        sensiblePlace = False
        Exit Function
    End If

    sensiblePlace = True
End Function

' ------------------------------------------------------------------------------------
' Should stop things being placed in inaccessible areas? Assume wheeled propulsion.
Public Function zonedPAT(ByVal x As Long, ByVal y As Long) As Boolean
    ' (&&: noDroid only for a sensible place)
    If sensiblePlace(x, y, PROPULSION_TYPE_WHEELED) Then zonedPAT = noDroid(x, y)
End Function

Private Function canFitDroid(ByVal x As Long, ByVal y As Long) As Boolean
    ' (&&: oneDroidMax only for a sensible place)
    If sensiblePlace(x, y, PROPULSION_TYPE_WHEELED) Then canFitDroid = oneDroidMax(x, y)
End Function

' (VB) calls pickATileGen's function pointer
Private Function callPickATileFunction(ByVal function_ As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Select Case function_
    Case PAT_zonedPAT: callPickATileFunction = zonedPAT(x, y)
    Case PAT_canFitDroid: callPickATileFunction = canFitDroid(x, y)
    End Select
End Function

' find a tile for which the function will return true
Public Function pickATileGen(x As Long, y As Long, ByVal numIterations As Long, ByVal function_ As Long) As Boolean
    pickATileGen = pickATileGenThreat(x, y, numIterations, -1, -1, function_)
End Function

' (VB) the Vector2i overload
Public Function pickATileGen_pos(pos As Vector2i, ByVal numIterations As Long, ByVal function_ As Long) As Boolean
    Dim x As Long, y As Long, ret As Boolean
    x = pos.x
    y = pos.y
    ret = pickATileGenThreat(x, y, numIterations And &HFF&, -1, -1, function_)
    pos = Vector2i_(x, y)
    pickATileGen_pos = ret
End Function

' (VB) hypotf of two tile differences, truncated: the square root in Double
' truncates to the same integer (the sum of squares is exact)
Private Function ThreatInRange(ByVal player As Long, ByVal range As Long, ByVal rangeX As Long, ByVal rangeY As Long, ByVal bVTOLs As Boolean) As Boolean
    Dim i As Long, structType As Long, tx As Long, ty As Long, it As Long, psStruct As Long, psDroid As Long, dx As Long, dy As Long

    tx = map_coord(rangeX)
    ty = map_coord(rangeY)

    For i = 0 To MAX_PLAYERS - 1
        If (alliances(player, i) = ALLIANCE_FORMED) Or (i = player) Then
            GoTo continue_
        End If

        'check structures
        it = apsStructLists(i).head
        Do While it <> 0
            psStruct = lnObj(it)
            If aStruct(psStruct).visible(player) <> 0 Or aStruct(psStruct).born = 2 Then    ' if can see it or started there
                If aStruct(psStruct).status = SS_BUILT Then
                    structType = asStructureStats(aStruct(psStruct).pStructureType).typ

                    Select Case structType      'dangerous to get near these structures
                    Case REF_DEFENSE, REF_CYBORG_FACTORY, REF_FACTORY, REF_VTOL_FACTORY, REF_REARM_PAD, REF_FORTRESS
                        dx = tx - map_coord(aStruct(psStruct).pos.x)
                        dy = ty - map_coord(aStruct(psStruct).pos.y)
                        If range < 0 _
                           Or world_coord(Fix(Sqr(CDbl(dx) * dx + CDbl(dy) * dy))) < range Then  'enemy in range
                            ThreatInRange = True
                            Exit Function
                        End If
                    End Select
                End If
            End If
            it = lnNext(it)
        Loop

        'check droids
        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            If aDroid(psDroid).visible(player) <> 0 Then    'can see this droid?
                If Not objHasWeapon(psDroid) Then
                    GoTo nextDroid
                End If

                'if VTOLs are excluded, skip them
                If Not bVTOLs And ((asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType = PROPULSION_TYPE_LIFT) Or DROID_isTransporter(psDroid)) Then
                    GoTo nextDroid
                End If

                dx = tx - map_coord(aDroid(psDroid).pos.x)
                dy = ty - map_coord(aDroid(psDroid).pos.y)
                If range < 0 _
                   Or world_coord(Fix(Sqr(CDbl(dx) * dx + CDbl(dy) * dy))) < range Then  'enemy in range
                    ThreatInRange = True
                    Exit Function
                End If
            End If
nextDroid:
            it = lnNext(it)
        Loop
continue_:
    Next

    ThreatInRange = False
End Function

' find a tile for which the passed function will return true without any threat in the specified range
Public Function pickATileGenThreat(x As Long, y As Long, ByVal numIterations As Long, ByVal threatRange As Long, _
                                   ByVal player As Long, ByVal function_ As Long) As Boolean
    Dim i As Long, j As Long
    Dim startX As Long, endX As Long, startY As Long, endY As Long
    Dim passes As Long
    Dim origin As Vector3i, newPos As Vector3i, good As Boolean
    origin = Vector3i_(world_coord(x), world_coord(y), 0)

    If Not (ULT(x, mapWidth)) Then ASSERT_FAILURE "*x < mapWidth", "src/droid.cpp:2815", "pickATileGenThreat", "x coordinate is off-map for pickATileGen": pickATileGenThreat = False: Exit Function
    If Not (ULT(y, mapHeight)) Then ASSERT_FAILURE "*y < mapHeight", "src/droid.cpp:2816", "pickATileGenThreat", "y coordinate is off-map for pickATileGen": pickATileGenThreat = False: Exit Function

    ' (&& and ||: ThreatInRange only with a threat range, for a tile the function accepts)
    If callPickATileFunction(function_, x, y) Then
        If threatRange <= 0 Then
            pickATileGenThreat = True
            Exit Function
        ElseIf Not ThreatInRange(player, threatRange, x, y, False) Then   'TODO: vtol check really not needed?
            pickATileGenThreat = True
            Exit Function
        End If
    End If

    ' Initial box dimensions and set iteration count to zero
    startX = x: endX = x: startY = y: endY = y: passes = 0

    ' Keep going until we get a tile or we exceed distance
    Do While passes < numIterations
        ' Process whole box
        For i = startX To endX
            For j = startY To endY
                ' Test only perimeter as internal tested previous iteration
                If i = startX Or i = endX Or j = startY Or j = endY Then
                    newPos = Vector3i_(world_coord(i), world_coord(j), 0)

                    ' Good enough?
                    ' (&&: fpathCheck and ThreatInRange only for a tile the function accepts)
                    If callPickATileFunction(function_, i, j) Then
                        If fpathCheck(origin, newPos, PROPULSION_TYPE_WHEELED) Then
                            If threatRange <= 0 Then
                                good = True
                            Else
                                good = Not ThreatInRange(player, threatRange, world_coord(i), world_coord(j), False)
                            End If
                            If good Then
                                ' Set exit conditions and get out NOW
                                x = i: y = j
                                pickATileGenThreat = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        Next
        ' Expand the box out in all directions - off map handled by tileAcceptable
        startX = startX - 1: startY = startY - 1: endX = endX + 1: endY = endY + 1: passes = passes + 1
    Loop
    ' If we got this far, then we failed - passed in values will be unchanged
    pickATileGenThreat = False

End Function

' find a tile for a wheeled droid with only one other droid present
Public Function pickHalfATile(x As Long, y As Long, ByVal numIterations As Long) As PICKTILE
    If pickATileGen(x, y, numIterations, PAT_canFitDroid) Then
        pickHalfATile = FREE_TILE
    Else
        pickHalfATile = NO_FREE_TILE
    End If
End Function

' Looks through the players list of droids to see if any of them are
' building the specified structure - returns true if finds one
Public Function checkDroidsBuilding(ByVal psStructure As Long) As Boolean
    Dim it As Long, psDroid As Long, psStruct As Long
    it = apsDroidLists(aStruct(psStructure).player).head
    Do While it <> 0
        psDroid = lnObj(it)
        'check DORDER_BUILD, HELP_BUILD is handled the same
        psStruct = orderStateObj(psDroid, DORDER_BUILD)
        If psStruct = psStructure Then
            checkDroidsBuilding = True
            Exit Function
        End If
        it = lnNext(it)
    Loop
    checkDroidsBuilding = False
End Function

' Looks through the players list of droids to see if any of them are
' demolishing the specified structure - returns true if finds one
Public Function checkDroidsDemolishing(ByVal psStructure As Long) As Boolean
    Dim it As Long, psDroid As Long, psStruct As Long
    it = apsDroidLists(aStruct(psStructure).player).head
    Do While it <> 0
        psDroid = lnObj(it)
        'check DORDER_DEMOLISH
        psStruct = orderStateObj(psDroid, DORDER_DEMOLISH)
        If psStruct = psStructure Then
            checkDroidsDemolishing = True
            Exit Function
        End If
        it = lnNext(it)
    Loop
    checkDroidsDemolishing = False
End Function

' Deals with building a module - checking if any droid is currently doing this
'  - if so, helping to build the current one
Public Sub setUpBuildModule(ByVal psDroid As Long)
    Dim tileX As Long, tileY As Long, psStruct As Long
    tileX = map_coord(aDroid(psDroid).order.pos.x)
    tileY = map_coord(aDroid(psDroid).order.pos.y)

    'check not another Truck started
    psStruct = getTileStructure(tileX, tileY)
    If psStruct <> 0 Then
        ' if a droid is currently building, or building is in progress of being built/upgraded the droid's order should be DORDER_HELPBUILD
        If checkDroidsBuilding(psStruct) Or aStruct(psStruct).status = 0 Then
            'set up the help build scenario
            aDroid(psDroid).order.typ = DORDER_HELPBUILD
            setDroidTarget psDroid, psStruct
            If droidStartBuild(psDroid) <> DroidStartBuildFailed Then
                aDroid(psDroid).action = DACTION_BUILD
                Exit Sub
            End If
        Else
            If nextModuleToBuild(psStruct, -1) > 0 Then
                'no other droids building so just start it off
                If droidStartBuild(psDroid) <> DroidStartBuildFailed Then
                    aDroid(psDroid).action = DACTION_BUILD
                    Exit Sub
                End If
            End If
        End If
    End If
    cancelBuild psDroid
End Sub

Public Function DROID_isDamaged(ByVal this As Long) As Boolean
    DROID_isDamaged = ULT(aDroid(this).body, aDroid(this).originalBody)
End Function

' checks to see if an electronic warfare weapon is attached to the droid
Public Function electronicDroid(ByVal psDroid As Long) As Boolean
    Dim it As Long, psCurr As Long

    checkDroid psDroid, "src/droid.cpp:3012", "electronicDroid", max_check_object_recursion

    'use slot 0 for now
    If aDroid(psDroid).numWeaps > 0 Then
        If asWeaponStats(DROID_getWeaponStats(psDroid, 0)).weaponSubClass = WSC_ELECTRONIC Then
            electronicDroid = True
            Exit Function
        End If
    End If

    If aDroid(psDroid).droidType = DROID_COMMAND And aDroid(psDroid).psGroup <> 0 Then
        If aGroup(aDroid(psDroid).psGroup).psCommander = psDroid Then
            ' if a commander has EW units attached it is electronic
            it = aGroup(aDroid(psDroid).psGroup).psList.head
            Do While it <> 0
                psCurr = lnObj(it)
                If psDroid <> psCurr Then
                    If electronicDroid(psCurr) Then
                        electronicDroid = True
                        Exit Function
                    End If
                End If
                it = lnNext(it)
            Loop
        End If
    End If

    electronicDroid = False
End Function

' checks to see if the droid is currently being repaired by another droid or structure
Public Function droidUnderRepair(ByVal psDroid As Long) As Boolean
    checkDroid psDroid, "src/droid.cpp:3038", "droidUnderRepair", max_check_object_recursion

    'droid must be damaged
    If DROID_isDamaged(psDroid) Then
        droidUnderRepair = aDroid(psDroid).underRepair <> 0
        Exit Function
    End If
    droidUnderRepair = False
End Function

' returns true if it's a VTOL weapon droid which has completed all runs
Public Function vtolEmpty(ByVal psDroid As Long) As Boolean
    Dim i As Long

    checkDroid psDroid, "src/droid.cpp:3096", "vtolEmpty", max_check_object_recursion

    If Not DROID_isVtol(psDroid) Then
        vtolEmpty = False
        Exit Function
    End If
    If aDroid(psDroid).droidType <> DROID_WEAPON Then
        vtolEmpty = False
        Exit Function
    End If

    For i = 0 To aDroid(psDroid).numWeaps - 1
        If asWeaponStats(DROID_getWeaponStats(psDroid, i)).vtolAttackRuns > 0 Then
            If ULT(aDroid(psDroid).asWeaps(i).usedAmmo, getNumAttackRuns(psDroid, i)) Then
                vtolEmpty = False
                Exit Function
            End If
        End If
    Next

    vtolEmpty = True
End Function

' returns true if it's a VTOL weapon droid which still has full ammo
Public Function vtolFull(ByVal psDroid As Long) As Boolean
    Dim i As Long

    checkDroid psDroid, "src/droid.cpp:3122", "vtolFull", max_check_object_recursion

    If Not DROID_isVtol(psDroid) Then
        vtolFull = False
        Exit Function
    End If
    If aDroid(psDroid).droidType <> DROID_WEAPON Then
        vtolFull = False
        Exit Function
    End If

    For i = 0 To aDroid(psDroid).numWeaps - 1
        If asWeaponStats(DROID_getWeaponStats(psDroid, i)).vtolAttackRuns > 0 And _
           aDroid(psDroid).asWeaps(i).usedAmmo <> 0 Then
            vtolFull = False
            Exit Function
        End If
    Next

    vtolFull = True
End Function

' true if a vtol droid currently returning to be rearmed
Public Function DROID_isVtolRearming(ByVal this As Long) As Boolean
    Dim a As DROID_ACTION

    checkDroid this, "src/droid.cpp:3203", "isVtolRearming", max_check_object_recursion

    If Not DROID_isVtol(this) Then
        DROID_isVtolRearming = False
        Exit Function
    End If
    If aDroid(this).droidType <> DROID_WEAPON Then
        DROID_isVtolRearming = False
        Exit Function
    End If

    a = aDroid(this).action
    If a = DACTION_MOVETOREARM Or _
       a = DACTION_WAITFORREARM Or _
       a = DACTION_MOVETOREARMPOINT Or _
       a = DACTION_WAITDURINGREARM Then
        DROID_isVtolRearming = True
        Exit Function
    End If

    DROID_isVtolRearming = False
End Function

' true if a droid is currently attacking
Public Function DROID_isAttacking(ByVal this As Long) As Boolean
    Dim a As DROID_ACTION, t As DROID_TYPE

    checkDroid this, "src/droid.cpp:3228", "isAttacking", max_check_object_recursion

    'what about cyborgs?
    t = aDroid(this).droidType
    If Not (t = DROID_WEAPON Or t = DROID_CYBORG Or _
            t = DROID_CYBORG_SUPER) Then
        DROID_isAttacking = False
        Exit Function
    End If

    a = aDroid(this).action
    If a = DACTION_ATTACK Or _
       a = DACTION_MOVETOATTACK Or _
       a = DACTION_ROTATETOATTACK Or _
       a = DACTION_VTOLATTACK Or _
       a = DACTION_MOVEFIRE Then
        DROID_isAttacking = True
        Exit Function
    End If

    DROID_isAttacking = False
End Function

' see if there are any other vtols attacking the same target
' but still rearming
Public Function allVtolsRearmed(ByVal psDroid As Long) As Boolean
    Dim stillRearming As Boolean, it As Long, psCurr As Long

    checkDroid psDroid, "src/droid.cpp:3253", "allVtolsRearmed", max_check_object_recursion

    ' ignore all non vtols
    If Not DROID_isVtol(psDroid) Then
        allVtolsRearmed = True
        Exit Function
    End If

    stillRearming = False
    it = apsDroidLists(aDroid(psDroid).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If DROID_isVtolRearming(psCurr) Then
            If aDroid(psCurr).order.typ = aDroid(psDroid).order.typ And _
               aDroid(psCurr).order.psObj = aDroid(psDroid).order.psObj Then
                stillRearming = True
                Exit Do
            End If
        End If
        it = lnNext(it)
    Loop

    allVtolsRearmed = Not stillRearming
End Function

' returns a count of the base number of attack runs for the weapon attached to the droid
Public Function getNumAttackRuns(ByVal psDroid As Long, ByVal weapon_slot As Long) As Long
    Dim weaponStats As Long
    If Not (DROID_isVtol(psDroid)) Then ASSERT_FAILURE "psDroid->isVtol()", "src/droid.cpp:3280", "getNumAttackRuns", "not a VTOL Droid": getNumAttackRuns = 0: Exit Function
    weaponStats = DROID_getWeaponStats(psDroid, weapon_slot)
    ' if weapon is a salvo weapon, then number of shots that can be fired = vtolAttackRuns * numRounds
    If asWeaponStats(weaponStats).upgrade(aDroid(psDroid).player).reloadTime <> 0 Then
        getNumAttackRuns = (asWeaponStats(weaponStats).upgrade(aDroid(psDroid).player).numRounds _
                            * asWeaponStats(weaponStats).vtolAttackRuns) And &HFFFF&
        Exit Function
    End If
    getNumAttackRuns = asWeaponStats(weaponStats).vtolAttackRuns And &HFFFF&
End Function

' true if a vtol is waiting to be rearmed by a particular rearm pad
Public Function vtolReadyToRearm(ByVal psDroid As Long, ByVal psStruct As Long) As Boolean
    Dim psRearmPad As Long, psTarget As Long

    checkDroid psDroid, "src/droid.cpp:3148", "vtolReadyToRearm", max_check_object_recursion

    If Not DROID_isVtol(psDroid) Or aDroid(psDroid).action <> DACTION_WAITFORREARM Then
        vtolReadyToRearm = False
        Exit Function
    End If

    ' If a unit has been ordered to rearm make sure it goes to the correct base
    psRearmPad = castStructure(orderStateObj(psDroid, DORDER_REARM))
    If psRearmPad <> 0 Then
        If psRearmPad <> psStruct Then
            If Not vtolOnRearmPad(psRearmPad, psDroid) Then
                ' target rearm pad is clear - let it go there
                objTrace aDroid(psDroid).id, "vtolReadyToRearm", "rearm pad at %d,%d won't snatch us - we already have another available at %d,%d", aStruct(psStruct).pos.x \ TILE_UNITS, aStruct(psStruct).pos.y \ TILE_UNITS, aStruct(psRearmPad).pos.x \ TILE_UNITS, aStruct(psRearmPad).pos.y \ TILE_UNITS
                vtolReadyToRearm = False
                Exit Function
            End If
        End If
    End If

    If vtolHappy(psDroid) Then
        If vtolOnRearmPad(psStruct, psDroid) Then
            ' there is a vtol on the pad and this vtol is already rearmed
            ' don't bother shifting the other vtol off
            objTrace aDroid(psDroid).id, "vtolReadyToRearm", "rearm pad at %d,%d won't snatch us - we're rearmed and pad is busy", aStruct(psStruct).pos.x \ TILE_UNITS, aStruct(psStruct).pos.y \ TILE_UNITS
            vtolReadyToRearm = False
            Exit Function
        End If
    End If

    psTarget = castStructure(aDroid(psDroid).psActionTarget(0))
    If psTarget <> 0 Then
        If aStruct(psTarget).pFunctionality <> 0 Then
            If aStruct(psTarget).functionality.rearmPad.psObj = psDroid Then
                ' vtol is rearming at a different base, leave it alone
                objTrace aDroid(psDroid).id, "vtolReadyToRearm", "rearm pad at %d,%d won't snatch us - we already are snatched by %d,%d", aStruct(psStruct).pos.x \ TILE_UNITS, aStruct(psStruct).pos.y \ TILE_UNITS, aStruct(psTarget).pos.x \ TILE_UNITS, aStruct(psTarget).pos.y \ TILE_UNITS
                vtolReadyToRearm = False
                Exit Function
            End If
        End If
    End If

    vtolReadyToRearm = True
End Function

' Fills all the weapons on a VTOL droid.
Public Sub fillVtolDroid(ByVal psDroid As Long)
    Dim i As Long
    checkDroid psDroid, "src/droid.cpp:3186", "fillVtolDroid", max_check_object_recursion
    If Not DROID_isVtol(psDroid) Then
        Exit Sub
    End If
    For i = 0 To aDroid(psDroid).numWeaps - 1
        ' Set rearm value to no runs made.
        aDroid(psDroid).asWeaps(i).usedAmmo = 0
        aDroid(psDroid).asWeaps(i).ammo = asWeaponStats(DROID_getWeaponStats(psDroid, i)).upgrade(aDroid(psDroid).player).numRounds
        aDroid(psDroid).asWeaps(i).lastFired = 0
    Next
End Sub

' Checks a vtol for being fully armed and fully repaired to see if ready to
' leave reArm pad
Public Function vtolHappy(ByVal psDroid As Long) As Boolean
    Dim i As Long

    checkDroid psDroid, "src/droid.cpp:3295", "vtolHappy", max_check_object_recursion

    If Not (DROID_isVtol(psDroid)) Then ASSERT_FAILURE "psDroid->isVtol()", "src/droid.cpp:3297", "vtolHappy", "not a VTOL droid": vtolHappy = False: Exit Function

    If DROID_isDamaged(psDroid) Then
        ' VTOLs with less health than their original aren't happy
        vtolHappy = False
        Exit Function
    End If

    If aDroid(psDroid).droidType <> DROID_WEAPON Then
        ' Not an armed droid, so don't check the (non-existent) weapons
        vtolHappy = True
        Exit Function
    End If

    ' NOTE: Previous code (r5410) returned false if a droid had no weapon,
    '       which IMO isn't correct, but might be expected behaviour. I'm
    '       also not sure if weapon droids (see the above droidType check)
    '       can even have zero weapons. -- Giel
    If Not (aDroid(psDroid).numWeaps > 0) Then ASSERT_FAILURE "psDroid->numWeaps > 0", "src/droid.cpp:3316", "vtolHappy", "VTOL weapon droid without weapons found!": vtolHappy = False: Exit Function

    'check full complement of ammo
    For i = 0 To aDroid(psDroid).numWeaps - 1
        If asWeaponStats(DROID_getWeaponStats(psDroid, i)).vtolAttackRuns > 0 _
           And aDroid(psDroid).asWeaps(i).usedAmmo <> 0 Then
            vtolHappy = False
            Exit Function
        End If
    Next

    vtolHappy = True
End Function

' checks if the droid is a VTOL droid and updates the attack runs as required
Public Sub updateVtolAttackRun(ByVal psDroid As Long, ByVal weapon_slot As Long)
    If DROID_isVtol(psDroid) Then
        If aDroid(psDroid).numWeaps > 0 Then
            If asWeaponStats(DROID_getWeaponStats(psDroid, weapon_slot)).vtolAttackRuns > 0 Then
                aDroid(psDroid).asWeaps(weapon_slot).usedAmmo = aDroid(psDroid).asWeaps(weapon_slot).usedAmmo + 1
                If aDroid(psDroid).asWeaps(weapon_slot).usedAmmo = getNumAttackRuns(psDroid, weapon_slot) Then
                    aDroid(psDroid).asWeaps(weapon_slot).ammo = 0
                End If
                'quick check doesn't go over limit
                If Not (aDroid(psDroid).asWeaps(weapon_slot).usedAmmo < UWORD_MAX) Then ASSERT_FAILURE "psDroid->asWeaps[weapon_slot].usedAmmo < 65535", "src/droid.cpp:3346", "updateVtolAttackRun", "too many attack runs"
            End If
        End If
    End If
End Sub

' assign rearmPad to the VTOL
Public Sub assignVTOLPad(ByVal psNewDroid As Long, ByVal psReArmPad As Long)
    Dim ok As Boolean
    If Not (DROID_isVtol(psNewDroid)) Then ASSERT_FAILURE "psNewDroid->isVtol()", "src/droid.cpp:3355", "assignVTOLPad", "%s is not a VTOL droid", objInfo(psNewDroid): Exit Sub
    ok = SIMPLE_OBJECT_type(psReArmPad) = OBJ_STRUCTURE
    If ok Then ok = asStructureStats(aStruct(psReArmPad).pStructureType).typ = REF_REARM_PAD
    If Not ok Then ASSERT_FAILURE "psReArmPad->type == OBJ_STRUCTURE && psReArmPad->pStructureType->type == REF_REARM_PAD", "src/droid.cpp:3356", "assignVTOLPad", "%s cannot rearm", objInfo(psReArmPad): Exit Sub

    setDroidBase psNewDroid, psReArmPad
End Sub

' compares the droid sensor type with the droid weapon type to see if the
' FIRE_SUPPORT order can be assigned
Public Function droidSensorDroidWeapon(ByVal psObj As Long, ByVal psDroid As Long) As Boolean
    Dim psStats As Long, compIndex As Long

    psStats = -1

    checkDroid psDroid, "src/droid.cpp:3369", "droidSensorDroidWeapon", max_check_object_recursion

    If psObj = 0 Or psDroid = 0 Then
        droidSensorDroidWeapon = False
        Exit Function
    End If

    'first check if the object is a droid or a structure
    If (SIMPLE_OBJECT_type(psObj) <> OBJ_DROID) And _
       (SIMPLE_OBJECT_type(psObj) <> OBJ_STRUCTURE) Then
        droidSensorDroidWeapon = False
        Exit Function
    End If
    'check same player
    If SIMPLE_OBJECT_player(psObj) <> aDroid(psDroid).player Then
        droidSensorDroidWeapon = False
        Exit Function
    End If
    'check obj is a sensor droid/structure
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        If aDroid(psObj).droidType <> DROID_SENSOR And _
           aDroid(psObj).droidType <> DROID_COMMAND Then
            droidSensorDroidWeapon = False
            Exit Function
        End If
        compIndex = aDroid(psObj).asBits(COMP_SENSOR)
        If Not (compIndex < asSensorStats_size) Then ASSERT_FAILURE "compIndex < asSensorStats.size()", "src/droid.cpp:3397", "droidSensorDroidWeapon", "Invalid range referenced for numSensorStats, %d > %zu", compIndex, asSensorStats_size: droidSensorDroidWeapon = False: Exit Function
        psStats = compIndex
    Case OBJ_STRUCTURE
        psStats = asStructureStats(aStruct(psObj).pStructureType).pSensor
        If psStats = -1 Then
            droidSensorDroidWeapon = False
            Exit Function
        End If
        If asSensorStats(psStats).location <> LOC_TURRET Then
            droidSensorDroidWeapon = False
            Exit Function
        End If
    Case Else
    End Select

    'check droid is a weapon droid - or Cyborg!!
    If Not (aDroid(psDroid).droidType = DROID_WEAPON Or aDroid(psDroid).droidType = _
            DROID_CYBORG Or aDroid(psDroid).droidType = DROID_CYBORG_SUPER) Then
        droidSensorDroidWeapon = False
        Exit Function
    End If

    'finally check the right droid/sensor combination
    ' check vtol droid with commander
    If DROID_isVtol(psDroid) And SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        If aDroid(psObj).droidType = DROID_COMMAND Then
            droidSensorDroidWeapon = True
            Exit Function
        End If
    End If

    'check vtol droid with vtol sensor
    If DROID_isVtol(psDroid) And aDroid(psDroid).asWeaps(0).nStat <> 0 Then
        If asSensorStats(psStats).typ = VTOL_INTERCEPT_SENSOR Or asSensorStats(psStats).typ = VTOL_CB_SENSOR Or asSensorStats(psStats).typ = SUPER_SENSOR Then ' || psStats->type == RADAR_DETECTOR_SENSOR
            droidSensorDroidWeapon = True
            Exit Function
        End If
        droidSensorDroidWeapon = False
        Exit Function
    End If

    ' Check indirect weapon droid with standard/CB/radar detector sensor
    If Not proj_Direct(DROID_getWeaponStats(psDroid, 0)) Then
        If (asSensorStats(psStats).typ = STANDARD_SENSOR Or asSensorStats(psStats).typ = INDIRECT_CB_SENSOR Or asSensorStats(psStats).typ = SUPER_SENSOR) Then ' || psStats->type == RADAR_DETECTOR_SENSOR
            If Not (SIMPLE_OBJECT_type(psObj) = OBJ_DROID) Then
                droidSensorDroidWeapon = True
                Exit Function
            ElseIf aDroid(psObj).droidType <> DROID_COMMAND Then
                droidSensorDroidWeapon = True
                Exit Function
            End If
        End If
        droidSensorDroidWeapon = False
        Exit Function
    End If
    droidSensorDroidWeapon = False
End Function

' return whether a droid has a CB sensor on it
Public Function cbSensorDroid(ByVal psDroid As Long) As Boolean
    Dim sensorType As SENSOR_TYPE
    If aDroid(psDroid).droidType <> DROID_SENSOR Then
        cbSensorDroid = False
        Exit Function
    End If
    sensorType = asSensorStats(DROID_getSensorStats(psDroid)).typ
    If sensorType = VTOL_CB_SENSOR _
       Or sensorType = INDIRECT_CB_SENSOR Then
        cbSensorDroid = True
        Exit Function
    End If

    cbSensorDroid = False
End Function

' return whether a droid has a standard sensor on it (standard, VTOL strike, or wide spectrum)
Public Function standardSensorDroid(ByVal psDroid As Long) As Boolean
    Dim sensorType As SENSOR_TYPE
    If aDroid(psDroid).droidType <> DROID_SENSOR Then
        standardSensorDroid = False
        Exit Function
    End If
    sensorType = asSensorStats(DROID_getSensorStats(psDroid)).typ
    If sensorType = VTOL_INTERCEPT_SENSOR _
       Or sensorType = STANDARD_SENSOR _
       Or sensorType = SUPER_SENSOR Then
        standardSensorDroid = True
        Exit Function
    End If

    standardSensorDroid = False
End Function

' ////////////////////////////////////////////////////////////////////////////
' Give a droid from one player to another - used in Electronic Warfare and multiplayer.
' Got to destroy the droid and build another since there are too many complications otherwise.
' Returns the droid created.
' (VB) the callers pass the defaults electronic = false, pos = {0, 0}
Public Function giftSingleDroid(ByVal psD As Long, ByVal toPlayer As Long, ByVal electronic As Boolean, pos As Vector2i) As Long
    Dim pT As Long, psNewDroid As Long, newPos As Vector3i, rot As Rotation, pickX As Long, pickY As Long, m As Long, v As Long
    Dim oldPlayer As Long, it As Long, psCurr As Long, psLeader As Long, i As Long, iWeap As Long, psStruct As Long

    checkDroid psD, "src/droid.cpp:3491", "giftSingleDroid", max_check_object_recursion
    If Not (Not isDead(psD)) Then ASSERT_FAILURE "!isDead(psD)", "src/droid.cpp:3492", "giftSingleDroid", "Cannot gift dead unit": giftSingleDroid = 0: Exit Function
    If Not (aDroid(psD).player <> toPlayer) Then ASSERT_FAILURE "psD->player != to", "src/droid.cpp:3493", "giftSingleDroid", "Cannot gift to self": giftSingleDroid = psD: Exit Function
    If Not (ULT(toPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "to < MAX_PLAYERS", "src/droid.cpp:3494", "giftSingleDroid", "Cannot gift to = %u", toPlayer: giftSingleDroid = 0: Exit Function

    ' Check unit limits (multiplayer only)
    syncDebug "giftSingleDroid", "Limits: %u/%d %u/%d %u/%d", getNumDroids(toPlayer), getMaxDroids(toPlayer), getNumConstructorDroids(toPlayer), getMaxConstructors(toPlayer), getNumCommandDroids(toPlayer), getMaxCommanders(toPlayer)
    If bMultiPlayer _
       And (getNumDroids(toPlayer) >= getMaxDroids(toPlayer) _
            Or ((aDroid(psD).droidType = DROID_CYBORG_CONSTRUCT Or aDroid(psD).droidType = DROID_CONSTRUCT) _
                And getNumConstructorDroids(toPlayer) >= getMaxConstructors(toPlayer)) _
            Or (aDroid(psD).droidType = DROID_COMMAND And getNumCommandDroids(toPlayer) >= getMaxCommanders(toPlayer))) Then
        If toPlayer = selectedPlayer Or aDroid(psD).player = selectedPlayer Then
            CONPRINTF "%s", "Unit transfer failed -- unit limits exceeded"
        End If
        giftSingleDroid = 0
        Exit Function
    End If

    ' electronic or campaign will destroy and recreate the droid.
    If electronic Or Not bMultiPlayer Then
        pT = newTemplate()                  ' DROID_TEMPLATE sTemplate (deleted at every return)

        templateSetParts psD, apsTemplates(pT)  ' create a template based on the droid
        apsTemplates(pT).name_ = aDroid(psD).aName  ' copy the name across
        ' update score
        If aDroid(psD).player = selectedPlayer And toPlayer <> selectedPlayer And Not bMultiPlayer Then
            scoreUpdateVar WD_UNITS_LOST
        End If
        ' make the old droid vanish (but is not deleted until next tick)
        adjustDroidCount psD, -1
        vanishDroid psD
        ' Pick coordinates of the new droid if damaged electronically
        newPos = Vector3i_(aDroid(psD).pos.x, aDroid(psD).pos.y, 0)
        If electronic Then
            pickX = map_coord(pos.x)
            pickY = map_coord(pos.y)
            If pickATileGen(pickX, pickY, LOOK_FOR_EMPTY_TILE, PAT_zonedPAT) Then   ' != NO_FREE_TILE
                newPos = Vector3i_(world_coord(pickX), world_coord(pickY), 0)
            Else
                newPos = Vector3i_(pos.x, pos.y, 0) 'Meld with the attacker if we must
            End If
        End If
        ' create a new droid
        rot = aDroid(psD).rot
        psNewDroid = reallyBuildDroid(pT, newPos, toPlayer, False, rot)
        If Not (psNewDroid <> 0) Then ASSERT_FAILURE "psNewDroid", "src/droid.cpp:3544", "giftSingleDroid", "Unable to build unit": deleteTemplate pT: giftSingleDroid = 0: Exit Function

        addDroid psNewDroid, apsDroidLists
        adjustDroidCount psNewDroid, 1

        ' clip((psD->body*psNewDroid->originalBody + psD->originalBody/2)/std::max(psD->originalBody, 1u), 1u, psNewDroid->originalBody), unsigned
        m = aDroid(psD).originalBody
        If ULT(m, 1) Then m = 1
        v = FromU32(U32(MUL32(aDroid(psD).body, aDroid(psNewDroid).originalBody)) + Fix(U32(aDroid(psD).originalBody) / 2))
        v = FromU32(Fix(U32(v) / U32(m)))
        If ULT(v, 1) Then
            v = 1
        ElseIf ULT(aDroid(psNewDroid).originalBody, v) Then
            v = aDroid(psNewDroid).originalBody
        End If
        aDroid(psNewDroid).body = v
        aDroid(psNewDroid).experience = aDroid(psD).experience
        aDroid(psNewDroid).kills = aDroid(psD).kills

        If Not (aDroid(psNewDroid).droidType = DROID_PERSON Or DROID_isCyborg(psNewDroid) Or DROID_isTransporter(psNewDroid)) Then
            updateDroidOrientation psNewDroid
        End If

        triggerEventObjectTransfer psNewDroid, aDroid(psD).player
        deleteTemplate pT
        giftSingleDroid = psNewDroid
        Exit Function
    End If

    ' (&&: transporterIsEmpty, whose ASSERT may log, only for a transporter)
    If DROID_isTransporter(psD) Then
        If Not transporterIsEmpty(psD) Then
            ' Cannot gift non-empty transporters
            giftSingleDroid = 0
            Exit Function
        End If
    End If

    oldPlayer = aDroid(psD).player

    ' reset the assigned state of units attached to a leader
    it = apsDroidLists(oldPlayer).head
    Do While it <> 0
        psCurr = lnObj(it)

        If hasCommander(psCurr) Then
            psLeader = aGroup(aDroid(psCurr).psGroup).psCommander
        Else
            'psLeader can be either a droid or a structure
            psLeader = orderStateObj(psCurr, DORDER_FIRESUPPORT)
        End If

        If psLeader <> 0 Then
            If SIMPLE_OBJECT_id(psLeader) = aDroid(psD).id Then
                aDroid(psCurr).selected = 0
                orderDroid psCurr, DORDER_STOP, ModeQueue
            End If
        End If
        it = lnNext(it)
    Loop

    visRemoveVisibility psD
    aDroid(psD).selected = 0

    adjustDroidCount psD, -1
    scriptRemoveObject psD 'Remove droid from any script groups

    If droidRemove(psD, apsDroidLists) Then
        aDroid(psD).player = toPlayer

        addDroid psD, apsDroidLists
        adjustDroidCount psD, 1

        ' the new player may have different default sensor/ecm/repair components
        If asSensorStats(DROID_getSensorStats(psD)).location = LOC_DEFAULT Then
            If aDroid(psD).asBits(COMP_SENSOR) <> aDefaultSensor(aDroid(psD).player) Then
                aDroid(psD).asBits(COMP_SENSOR) = aDefaultSensor(aDroid(psD).player)
            End If
        End If
        If asECMStats(DROID_getECMStats(psD)).location = LOC_DEFAULT Then
            If aDroid(psD).asBits(COMP_ECM) <> aDefaultECM(aDroid(psD).player) Then
                aDroid(psD).asBits(COMP_ECM) = aDefaultECM(aDroid(psD).player)
            End If
        End If
        If asRepairStats(DROID_getRepairStats(psD)).location = LOC_DEFAULT Then
            If aDroid(psD).asBits(COMP_REPAIRUNIT) <> aDefaultRepair(aDroid(psD).player) Then
                aDroid(psD).asBits(COMP_REPAIRUNIT) = aDefaultRepair(aDroid(psD).player)
            End If
        End If
    Else
        ' if we couldn't remove it, then get rid of it.
        giftSingleDroid = 0
        Exit Function
    End If

    ' Update visibility
    visTilesUpdate psD

    ' check through the players, and our allies, list of droids to see if any are targetting it
    For i = 0 To MAX_PLAYERS - 1
        If Not aiCheckAlliances(i, toPlayer) Then
            GoTo continue1_
        End If

        it = apsDroidLists(i).head
        Do While it <> 0
            psCurr = lnObj(it)
            If aDroid(psCurr).order.psObj = psD Or aDroid(psCurr).psActionTarget(0) = psD Then
                orderDroid psCurr, DORDER_STOP, ModeQueue
                Exit Do
            End If
            For iWeap = 0 To aDroid(psCurr).numWeaps - 1
                If aDroid(psCurr).psActionTarget(iWeap) = psD Then
                    orderDroid psCurr, DORDER_STOP, ModeImmediate
                    Exit For
                End If
            Next
            ' check through order list
            orderClearTargetFromDroidList psCurr, psD
            it = lnNext(it)
        Loop
continue1_:
    Next

    For i = 0 To MAX_PLAYERS - 1
        If Not aiCheckAlliances(i, toPlayer) Then
            GoTo continue2_
        End If

        ' check through the players list, and our allies, of structures to see if any are targetting it
        it = apsStructLists(i).head
        Do While it <> 0
            psStruct = lnObj(it)
            If aStruct(psStruct).psTarget(0) = psD Then
                setStructureTarget psStruct, 0, 0, ORIGIN_UNKNOWN
            End If
            it = lnNext(it)
        Loop
continue2_:
    Next

    triggerEventObjectTransfer psD, oldPlayer
    giftSingleDroid = psD
End Function

' int16_t DROID::droidResistance() const
Public Function DROID_droidResistance(ByVal this As Long) As Long
    Dim psStats As Long, res As Long, r As Long
    checkDroid this, "src/droid.cpp:3688", "droidResistance", max_check_object_recursion
    psStats = DROID_getBodyStats(this)
    r = asBodyStats(psStats).upgrade(aDroid(this).player).resistance
    res = FromU32(Fix(U32(aDroid(this).experience) / U32(65536 \ MAX_(1, r))))
    ' ensure resistance is a base minimum
    res = MAX_(res, r)
    DROID_droidResistance = ToInt16(MIN_(res, INT16_MAX))
End Function

Public Function isSelectable(ByVal psDroid As Long) As Boolean
    If (aDroid(psDroid).flags And BIT_(OBJECT_FLAG_UNSELECTABLE)) <> 0 Then
        isSelectable = False
        Exit Function
    End If

    ' we shouldn't ever control the transporter in SP games
    If DROID_isTransporter(psDroid) And Not bMultiPlayer Then
        isSelectable = False
        Exit Function
    End If

    isSelectable = True
End Function

' Select a droid and do any necessary housekeeping.
Public Sub SelectDroid(ByVal psDroid As Long, Optional ByVal programmaticSelection As Boolean = False)
    If Not isSelectable(psDroid) Then
        Exit Sub
    End If

    aDroid(psDroid).selected = 1
    intRefreshScreen
    If Not programmaticSelection Then
        triggerEventSelected
        jsDebugSelected psDroid
    End If
End Sub

' If all other droids with psGroupDroid's group are selected, add psGroupDroid to the selection after production/repair/etc.
Public Sub SelectGroupDroid(ByVal psGroupDroid As Long)
    Dim groupDroids() As Long, n As Long, it As Long, psDroid As Long, bDoSelection As Boolean, i As Long
    n = 0
    it = apsDroidLists(aDroid(psGroupDroid).player).head
    Do While it <> 0
        psDroid = lnObj(it)
        ' skip itself because psGroupDroid may already exist in apsDroidLists
        If psDroid <> psGroupDroid Then
            If aDroid(psDroid).group = aDroid(psGroupDroid).group Then
                If n = 0 Then
                    ReDim groupDroids(0 To 15)
                ElseIf n > UBound(groupDroids) Then
                    ReDim Preserve groupDroids(0 To n * 2 - 1)
                End If
                groupDroids(n) = psDroid
                n = n + 1
            End If
        End If
        it = lnNext(it)
    Loop

    If n > 0 Then
        bDoSelection = True
        For i = 0 To n - 1
            If aDroid(groupDroids(i)).selected = 0 Then
                bDoSelection = False
                Exit For
            End If
        Next

        If bDoSelection Then
            SelectDroid psGroupDroid, True
        End If
    End If
End Sub

' De-select a droid and do any necessary housekeeping.
Public Sub DeSelectDroid(ByVal psDroid As Long)
    aDroid(psDroid).selected = 0
    intRefreshScreen
    triggerEventSelected
End Sub

' Teleport a droid to a new position on the map
Public Sub droidSetPosition(ByVal psDroid As Long, ByVal x As Long, ByVal y As Long)
    aDroid(psDroid).pos.x = x
    aDroid(psDroid).pos.y = y
    aDroid(psDroid).pos.z = map_Height(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)
    initDroidMovement psDroid
    visTilesUpdate psDroid
End Sub

Public Function droidSqDist(ByVal psDroid As Long, ByVal psObj As Long) As Long
    Dim psPropStats As Long, objPos As Vector3i

    psPropStats = DROID_getPropulsionStats(psDroid)
    objPos = SIMPLE_OBJECT_pos(psObj)

    If Not fpathCheck(aDroid(psDroid).pos, objPos, asPropulsionStats(psPropStats).propulsionType) Then
        droidSqDist = -1
        Exit Function
    End If
    droidSqDist = objPosDiffSq(aDroid(psDroid).pos, objPos)
End Function

' WEAPON_STATS* BASE_OBJECT::getWeaponStats(int weaponSlot) const, called on a DROID (the asWeaponStats index)
Public Function DROID_getWeaponStats(ByVal this As Long, ByVal weaponSlot As Long) As Long
    DROID_getWeaponStats = aDroid(this).asWeaps(weaponSlot).nStat
End Function

' droid.h: static inline int droidSensorRange(const DROID *psDroid)
Public Function droidSensorRange(ByVal psDroid As Long) As Long
    droidSensorRange = objSensorRange(psDroid)
End Function

Public Function droidGetMaxShieldPoints(ByVal psDroid As Long) As Long
    droidGetMaxShieldPoints = 0 ' currently disabled, but left for experimentation with different implementations
End Function

' DROID_ANIMEVENT_DYING's effect: the body's dying animation drawn by the effects (droid.cpp)
Public Sub addDroidDeathAnimationEffect(ByVal psDroid As Long)
    Dim iDroidDir As Long, adiff As Long, move As Long, velocity As Vector3f, position As Vector3i, rotationV As Vector3i
    Dim psBaseShapeBody As Long, psShapeBody As Long, strImd As Long, psPropStats As Long, typ As EFFECT_TYPE

    ' DERIVED from moveCalcNormalSpeed and moveGetDroidPosDiffs:

    iDroidDir = aDroid(psDroid).rot.direction
    adiff = (iDroidDir - aDroid(psDroid).sMove.moveDir) And &HFFFF&   ' Cast wrapping intended.
    move = iCosR(adiff, aDroid(psDroid).sMove.speed)
    velocity.x = iSinR(iDroidDir, move)
    velocity.y = iCosR(iDroidDir, move)
    velocity.z = 0!

    ' DERIVED FROM displayComponentObject:

    ' Get the real position
    position.x = aDroid(psDroid).pos.x      ' psDroid->pos.xzy()
    position.y = aDroid(psDroid).pos.z
    position.z = aDroid(psDroid).pos.y

    If DROID_isTransporter(psDroid) Then
        position.y = position.y - bobTransporterHeight()
    End If

    ' Get all the pitch,roll,yaw info
    rotationV.y = -aDroid(psDroid).rot.direction
    rotationV.x = aDroid(psDroid).rot.pitch
    rotationV.z = aDroid(psDroid).rot.roll

    ' ignore electronic damage shimmy

    If BASE_OBJECT_visibleForLocalDisplay(psDroid) = UBYTE_MAX Then
        ' Get the body graphic now
        psBaseShapeBody = asBodyStats(DROID_getBodyStats(psDroid)).pIMD     ' BODY_IMD(psDroid, psDroid->player)
        If psBaseShapeBody = 0 Then
            Exit Sub
        End If
        psShapeBody = iIMDBaseShape_displayModel(psBaseShapeBody)
        If iIMDShape_objanimpie(psShapeBody, ANIM_EVENT_DYING) <> 0 Then
            strImd = iIMDBaseShape_displayModel(iIMDShape_objanimpie(psShapeBody, ANIM_EVENT_DYING))
            ' get propulsion stats (psPropStats: an index, never nullptr)
            psPropStats = DROID_getPropulsionStats(psDroid)
            If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_PROPELLOR Then
                ' FIXME: change when adding submarines to the game
                ' modifiedModelMatrix *= glm::translate(glm::vec3(0.f, -world_coord(1) / 2.3f, 0.f));
                position.y = position.y + Fix(CSng(world_coord(1)) / 2.3!)
            End If

            typ = DROID_ANIMEVENT_DYING_NORMAL
            If aDroid(psDroid).droidType = DROID_PERSON Then
                typ = DROID_ANIMEVENT_DYING_NORMAL_ST
                ' Applies:
                ' modifiedModelMatrix *= glm::scale(glm::vec3(.75f)); // FIXME - hideous....!!!!
            End If

            SetEffectForPlayer aDroid(psDroid).player
            effectSetSize 100
            addEffect_rotVelocity position, EFFECT_DROID_ANIMEVENT_DYING, typ, True, strImd, 0, graphicsTime, rotationV, velocity
        End If
    End If
End Sub

' Deals damage to a droid
' \param psDroid droid to deal damage to
' \param psProjectile projectile which hit the object (may be nullptr)
' \param damage amount of damage to deal
' \param weaponClass the class of the weapon that deals the damage
' \param weaponSubClass the subclass of the weapon that deals the damage
' \param angle angle of impact (from the damage dealing projectile in relation to this droid)
' \return > 0 when the dealt damage destroys the droid, < 0 when the droid survives
Public Function droidDamage(ByVal psDroid As Long, ByVal psProjectile As Long, ByVal damage As Long, ByVal weaponClass As WEAPON_CLASS, ByVal weaponSubClass As WEAPON_SUBCLASS, ByVal impactTime As Long, ByVal isDamagePerSecond As Boolean, ByVal minDamage As Long, ByVal empRadiusHit As Boolean) As Long
    Dim relativeDamage As Long, c As Boolean, useDeathAnimation As Boolean, additional As Double

    checkDroid psDroid, "src/droid.cpp:297", "droidDamage", max_check_object_recursion

    ' VTOLs (and transporters in MP) on the ground take triple damage
    c = DROID_isVtol(psDroid)
    If Not c Then c = DROID_isTransporter(psDroid) And bMultiPlayer
    If c And (aDroid(psDroid).sMove.Status = MOVEINACTIVE) Then
        damage = MUL32(damage, 3)
    End If

    relativeDamage = objDamage(psDroid, psProjectile, damage, aDroid(psDroid).originalBody, weaponClass, weaponSubClass, isDamagePerSecond, minDamage, empRadiusHit)

    If relativeDamage <> 0 And aDroid(psDroid).player = selectedPlayer And aDroid(psDroid).timeLastHit = gameTime Then
        If relativeDamage > 0 Then additional = relativeDamage Else additional = 0
        If aDroid(psDroid).group <> UBYTE_MAX Then
            intGroupDamaged aDroid(psDroid).group, additional, (relativeDamage < 0) ' update UI information
        ElseIf aDroid(psDroid).psGroup <> 0 Then
            If aGroup(aDroid(psDroid).psGroup).psCommander <> 0 Then
                intCommanderGroupDamaged aGroup(aDroid(psDroid).psGroup).psCommander, additional, (relativeDamage < 0) ' update UI information
            End If
        End If
    End If

    If relativeDamage > 0 Then
        ' reset the attack level
        If secondaryGetState(psDroid, DSO_ATTACK_LEVEL) = DSS_ALEV_ATTACKED Then
            secondarySetState psDroid, DSO_ATTACK_LEVEL, DSS_ALEV_ALWAYS
        End If
        ' Now check for auto return on droid's secondary orders (i.e. return on medium/heavy damage)
        secondaryCheckDamageLevel psDroid

        checkDroid psDroid, "src/droid.cpp:329", "droidDamage", max_check_object_recursion
    ElseIf relativeDamage < 0 Then
        ' Droid destroyed
        debug_ 334, LOG_ATTACK, "droidDamage", "droid (%d): DESTROYED", aDroid(psDroid).id

        ' Deal with score increase/decrease and messages to the player
        If aDroid(psDroid).player = selectedPlayer Then
            ' TRANSLATORS:	Refers to the loss of a single unit, known by its name
            CONPRINTF "%s Lost!", objInfo(psDroid)
            scoreUpdateVar WD_UNITS_LOST
            audio_QueueTrackMinDelayPos ID_SOUND_UNIT_DESTROYED, UNIT_LOST_DELAY, _
                                        aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, aDroid(psDroid).pos.z
        ' only counts as a kill if it's not our ally
        ElseIf ULT(selectedPlayer, MAX_PLAYERS) Then
            If Not aiCheckAlliances(aDroid(psDroid).player, selectedPlayer) Then
                scoreUpdateVar WD_UNITS_KILLED
            End If
        End If

        ' Do we have a dying animation?
        c = iIMDShape_objanimpie(iIMDBaseShape_displayModel(aDroid(psDroid).sDisplay.imd), ANIM_EVENT_DYING) <> 0
        If c And aDroid(psDroid).animationEvent <> ANIM_EVENT_DYING Then
            useDeathAnimation = True
            'Babas should not burst into flames from non-heat weapons
            If aDroid(psDroid).droidType = DROID_PERSON Then
                If weaponClass = WC_HEAT Then
                    ' NOTE: 3 types of screams are available ID_SOUND_BARB_SCREAM - ID_SOUND_BARB_SCREAM3
                    audio_PlayObjDynamicTrack psDroid, ID_SOUND_BARB_SCREAM + (rand() Mod 3), 0
                Else
                    useDeathAnimation = False
                End If
            End If
            If useDeathAnimation Then
                If DrawnInLastFrame(aDroid(psDroid).sDisplay.frameNumber) Then
                    addDroidDeathAnimationEffect psDroid ' spawn an effect to handle drawing the death animation
                End If
                debug_ 374, LOG_DEATH, "droidDamage", "%s droid %d (%p) is starting death animation", objInfo(psDroid), aDroid(psDroid).id, psDroid
                aDroid(psDroid).animationEvent = ANIM_EVENT_DYING ' set this so that removeDroidFX (called by destroyDroid) knows not to do certain effects
            End If
        End If
        ' Destroy the droid
        debug_ 379, LOG_DEATH, "droidDamage", "%s droid %d (%p) is toast", objInfo(psDroid), aDroid(psDroid).id, psDroid
        ' This should be sent even if multi messages are turned off, as the group message that was
        ' sent won't contain the destroyed droid
        If bMultiPlayer And Not bMultiMessages Then
            bMultiMessages = True
            destroyDroid psDroid, impactTime
            bMultiMessages = False
        Else
            destroyDroid psDroid, impactTime
        End If
    End If

    droidDamage = relativeDamage
End Function

' calculate muzzle tip location in 3d world
Public Function calcDroidMuzzleLocation(ByVal psDroid As Long, muzzle As Vector3i, ByVal weapon_slot As Long) As Boolean
    Dim psBodyImd As Long, debugStr As String, barrel As Vector3i, psWeaponImd As Long, psMountImd As Long
    Dim af As Affine3F, cx As Long, cy As Long, cz As Long, r As Vector3i, ok As Boolean, connector_num As Long, nConnectors As Long

    psBodyImd = asBodyStats(DROID_getBodyStats(psDroid)).pIMD    ' BODY_IMD(psDroid, psDroid->player)

    checkDroid psDroid, "src/droid.cpp:2407", "calcDroidMuzzleLocation", max_check_object_recursion

    ok = psBodyImd <> 0
    If ok Then ok = ULT(weapon_slot, iIMDBaseShape_connectors_size(psBodyImd))
    If ok Then
        ' (VB) debugStr: char debugStr[250], debugLen (sprintf at debugStr + debugLen) is a string appended to

        barrel.x = 0: barrel.y = 0: barrel.z = 0
        psWeaponImd = 0
        psMountImd = 0

        If aDroid(psDroid).asWeaps(weapon_slot).nStat <> 0 Then
            psMountImd = asWeaponStats(DROID_getWeaponStats(psDroid, weapon_slot)).pMountGraphic  ' WEAPON_MOUNT_IMD
            psWeaponImd = asWeaponStats(DROID_getWeaponStats(psDroid, weapon_slot)).pIMD          ' WEAPON_IMD
        End If

        Affine3F_construct af

        Affine3F_Trans af, aDroid(psDroid).pos.x, -aDroid(psDroid).pos.z, aDroid(psDroid).pos.y

        'matrix = the center of droid
        Affine3F_RotY af, aDroid(psDroid).rot.direction
        Affine3F_RotX af, aDroid(psDroid).rot.pitch
        Affine3F_RotZ af, -aDroid(psDroid).rot.roll
        iIMDBaseShape_connector psBodyImd, weapon_slot, cx, cy, cz
        Affine3F_Trans af, cx, -cz, -cy 'note y and z flipped
        debugStr = debugStr & va("connect:body[%d]=(%d,%d,%d)", weapon_slot, cx, -cz, -cy)

        'matrix = the weapon[slot] mount on the body
        Affine3F_RotY af, aDroid(psDroid).asWeaps(weapon_slot).rot.direction  ' +ve anticlockwise

        ' process turret mount
        If psMountImd <> 0 Then
            If iIMDBaseShape_connectors_size(psMountImd) <> 0 Then
                iIMDBaseShape_connector psMountImd, 0, cx, cy, cz
                Affine3F_Trans af, cx, -cz, -cy
                debugStr = debugStr & va(",turret=(%d,%d,%d)", cx, -cz, -cy)
            End If
        End If

        'matrix = the turret connector for the gun
        Affine3F_RotX af, aDroid(psDroid).asWeaps(weapon_slot).rot.pitch  ' +ve up

        'process the gun
        If psWeaponImd <> 0 Then
            nConnectors = iIMDBaseShape_connectors_size(psWeaponImd)
            If nConnectors <> 0 Then
                connector_num = 0

                ' which barrel is firing if model have multiple muzzle connectors?
                If aDroid(psDroid).asWeaps(weapon_slot).shotsFired <> 0 And (nConnectors > 1) Then
                    ' shoot first, draw later - substract one shot to get correct results
                    connector_num = FromU32(U32(FromU32(U32(aDroid(psDroid).asWeaps(weapon_slot).shotsFired) - 1)) - Int((U32(FromU32(U32(aDroid(psDroid).asWeaps(weapon_slot).shotsFired) - 1))) / nConnectors) * nConnectors)
                End If

                iIMDBaseShape_connector psWeaponImd, connector_num, cx, cy, cz
                barrel = Vector3i_(cx, -cz, -cy)
                debugStr = debugStr & va(",barrel[%u]=(%d,%d,%d)", connector_num, cx, -cy, -cz)
            End If
        End If

        r = Affine3F_mul_vec(af, barrel)
        muzzle.x = r.x  ' xzy()
        muzzle.y = r.z
        muzzle.z = r.y
        muzzle.z = -muzzle.z
        debugStr = debugStr & va(",muzzle=(%d,%d,%d)", muzzle.x, muzzle.y, muzzle.z)

        syncDebug "calcDroidMuzzleLocation", "%s", debugStr
    Else
        muzzle.x = aDroid(psDroid).pos.x
        muzzle.y = aDroid(psDroid).pos.y
        muzzle.z = aDroid(psDroid).pos.z + imdMaxY(aDroid(psDroid).sDisplay.imd)
    End If

    checkDroid psDroid, "src/droid.cpp:2476", "calcDroidMuzzleLocation", max_check_object_recursion

    calcDroidMuzzleLocation = True
End Function

' (VB) lib/framework/gettext.h's ngettext without translations: the singular for 1, else the plural
Private Function ngettext(ByVal msgid1 As String, ByVal msgid2 As String, ByVal n As Long) As String
    If n = 1 Then ngettext = msgid1 Else ngettext = msgid2
End Function

' Make all the droids for a certain player a member of a specific group
' If no droids are selected: If a structure is selected, set its group to which droids will be automatically assigned
Public Sub assignObjectToGroup(ByVal playerNumber As Long, ByVal groupNumber As Long, ByVal clearGroup As Boolean)
    Dim bAtLeastOne As Boolean, numCleared As Long, it As Long, psDroid As Long, psStruct As Long, newSelectedGroup As Long
    bAtLeastOne = False
    numCleared = 0

    If Not (ULT(playerNumber, MAX_PLAYERS)) Then ASSERT_FAILURE "playerNumber < MAX_PLAYERS", "src/droid.cpp:2079", "assignObjectToGroup", "Invalid player: %u", playerNumber: Exit Sub

    If ULT(groupNumber, UBYTE_MAX) Then
        ' Run through all the droids
        it = apsDroidLists(playerNumber).head
        Do While it <> 0
            psDroid = lnObj(it)
            ' Clear out the old ones
            If clearGroup And aDroid(psDroid).group = groupNumber Then
                aDroid(psDroid).group = UBYTE_MAX
                numCleared = numCleared + 1
            End If

            ' Only assign the currently selected ones
            If aDroid(psDroid).selected Then
                ' Set them to the right group - they can only be a member of one group
                aDroid(psDroid).group = groupNumber And &HFF&
                aDroid(psDroid).repairGroup = UBYTE_MAX
                bAtLeastOne = True
            End If
            it = lnNext(it)
        Loop
    End If
    If bAtLeastOne Then
        'clear the Deliv Point if one
        If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/droid.cpp:2106", "assignObjectToGroup", "Unsupported selectedPlayer: %u", selectedPlayer: Exit Sub
        it = apsFlagPosLists(selectedPlayer).head
        Do While it <> 0
            aFlagPos(lnObj(it)).selected = False
            it = lnNext(it)
        Loop
        groupConsoleInformOfCreation groupNumber
        secondarySetAverageGroupState selectedPlayer, groupNumber
    End If
    If bAtLeastOne Or numCleared > 0 Then
        ' optional<UBYTE> newSelectedGroup; (-1: nullopt)
        newSelectedGroup = -1
        If bAtLeastOne Then
            newSelectedGroup = groupNumber And &HFF&
        End If
        intGroupsChanged newSelectedGroup
    End If
    If Not bAtLeastOne And ULT(groupNumber, UBYTE_MAX) Then
        ' If no droids were selected to be added to the group, check if a factory is selected
        ' Run through all the structures
        it = apsStructLists(playerNumber).head
        Do While it <> 0
            psStruct = lnObj(it)
            If aStruct(psStruct).selected Then
                If STRUCTURE_isFactory(psStruct) Then
                    If aStruct(psStruct).productToGroup <> (groupNumber And &HFF&) Then
                        aStruct(psStruct).productToGroup = groupNumber And &HFF&
                    Else
                        ' To make it possible to clear factory group assignment via "toggle" behavior of assigning a factory to the same group it's already assigned
                        aStruct(psStruct).productToGroup = UBYTE_MAX
                    End If
                    Exit Sub
                End If
            End If
            it = lnNext(it)
        Loop
    End If
End Sub

Public Sub removeObjectFromGroup(ByVal playerNumber As Long)
    Dim removedCount As Long, it As Long, psStruct As Long, psDroid As Long
    removedCount = 0

    If Not (ULT(playerNumber, MAX_PLAYERS)) Then ASSERT_FAILURE "playerNumber < MAX_PLAYERS", "src/droid.cpp:2151", "removeObjectFromGroup", "Invalid player: %u", playerNumber: Exit Sub

    it = apsStructLists(playerNumber).head
    Do While it <> 0
        psStruct = lnObj(it)
        If aStruct(psStruct).selected Then
            If STRUCTURE_isFactory(psStruct) Then
                aStruct(psStruct).productToGroup = UBYTE_MAX
                Exit Sub
            End If
        End If
        it = lnNext(it)
    Loop

    it = apsDroidLists(playerNumber).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected Then
            aDroid(psDroid).group = UBYTE_MAX
            aDroid(psDroid).repairGroup = UBYTE_MAX
            removedCount = removedCount + 1
        End If
        it = lnNext(it)
    Loop
    If removedCount <> 0 Then
        groupConsoleInformOfRemoval
        intGroupsChanged
    End If
End Sub

Public Function activateGroupAndMove(ByVal playerNumber As Long, ByVal groupNumber As Long) As Boolean
    Dim psCentreDroid As Long, selected As Boolean, numDeselected As Long, it As Long, psDroid As Long
    psCentreDroid = 0
    selected = False
    numDeselected = 0

    If Not (ULT(playerNumber, MAX_PLAYERS)) Then ASSERT_FAILURE "playerNumber < MAX_PLAYERS", "src/droid.cpp:2184", "activateGroupAndMove", "Invalid player: %u", playerNumber: activateGroupAndMove = False: Exit Function

    If ULT(groupNumber, UBYTE_MAX) Then
        it = apsDroidLists(playerNumber).head
        Do While it <> 0
            psDroid = lnObj(it)
            ' Wipe out the ones in the wrong group
            If aDroid(psDroid).selected And aDroid(psDroid).group <> groupNumber Then
                DeSelectDroid psDroid
                numDeselected = numDeselected + 1
            End If
            ' Get the right ones
            If aDroid(psDroid).group = groupNumber Then
                SelectDroid psDroid
                psCentreDroid = psDroid
            End If
            it = lnNext(it)
        Loop

        ' There was at least one in the group
        If psCentreDroid <> 0 Then
            'clear the Deliv Point if one
            If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/droid.cpp:2208", "activateGroupAndMove", "Unsupported selectedPlayer: %u", selectedPlayer
            If ULT(selectedPlayer, MAX_PLAYERS) Then
                it = apsFlagPosLists(selectedPlayer).head
                Do While it <> 0
                    aFlagPos(lnObj(it)).selected = False
                    it = lnNext(it)
                Loop
            End If

            selected = True
            If getWarCamStatus() Then
                camToggleStatus          ' messy - fix this
                processWarCam 'odd, but necessary
                camToggleStatus             ' messy - FIXME
            Else
                ' Centre display on him if warcam isn't active
                setViewPos map_coord(aDroid(psCentreDroid).pos.x), map_coord(aDroid(psCentreDroid).pos.y), True
            End If
        End If
    End If

    If selected Then
        groupConsoleInformOfCentering groupNumber
    End If

    If selected Or numDeselected > 0 Then
        If selected Then
            intGroupsChanged groupNumber And &HFF&
        Else
            intGroupsChanged UBYTE_MAX
        End If
    End If

    activateGroupAndMove = selected
End Function

Public Function activateNoGroup(ByVal playerNumber As Long, ByVal selectionType_ As SELECTIONTYPE, ByVal selectionClass As SELECTION_CLASS, ByVal bOnScreen As Boolean) As Boolean
    Dim selected As Boolean, dselectionType As SELECTIONTYPE, dselectionClass As SELECTION_CLASS, selectionCount As Long, dbOnScreen As Boolean
    Dim it As Long, psDroid As Long
    selected = False
    dselectionType = selectionType_
    dselectionClass = selectionClass
    selectionCount = 0
    dbOnScreen = bOnScreen
    If Not (ULT(playerNumber, MAX_PLAYERS)) Then ASSERT_FAILURE "playerNumber < MAX_PLAYERS", "src/droid.cpp:2252", "activateNoGroup", "Invalid player: %u", playerNumber: activateNoGroup = False: Exit Function

    selectionCount = selDroidSelection(selectedPlayer, dselectionClass, dselectionType, dbOnScreen)
    it = apsDroidLists(playerNumber).head
    Do While it <> 0
        psDroid = lnObj(it)
        ' Wipe out the ones in the wrong group
        If aDroid(psDroid).selected And aDroid(psDroid).group <> UBYTE_MAX Then
            DeSelectDroid psDroid
            selectionCount = selectionCount - 1
        End If
        it = lnNext(it)
    Loop
    If selected Then
        'clear the Deliv Point if one
        If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/droid.cpp:2267", "activateNoGroup", "Unsupported selectedPlayer: %u", selectedPlayer: activateNoGroup = False: Exit Function
        it = apsFlagPosLists(selectedPlayer).head
        Do While it <> 0
            aFlagPos(lnObj(it)).selected = False
            it = lnNext(it)
        Loop
    End If
    intGroupsChanged UBYTE_MAX
    CONPRINTF ngettext("%u unit selected", "%u units selected", selectionCount), selectionCount
    activateNoGroup = selected
End Function

Public Function activateGroup(ByVal playerNumber As Long, ByVal groupNumber As Long) As Boolean
    Dim selected As Boolean, numDeselected As Long, it As Long, psDroid As Long
    selected = False
    numDeselected = 0

    If Not (ULT(playerNumber, MAX_PLAYERS)) Then ASSERT_FAILURE "playerNumber < MAX_PLAYERS", "src/droid.cpp:2283", "activateGroup", "Invalid player: %u", playerNumber: activateGroup = False: Exit Function

    If ULT(groupNumber, UBYTE_MAX) Then
        it = apsDroidLists(playerNumber).head
        Do While it <> 0
            psDroid = lnObj(it)
            ' Wipe out the ones in the wrong group
            If aDroid(psDroid).selected And aDroid(psDroid).group <> groupNumber Then
                DeSelectDroid psDroid
                numDeselected = numDeselected + 1
            End If
            ' Get the right ones
            If aDroid(psDroid).group = groupNumber Then
                SelectDroid psDroid
                selected = True
            End If
            it = lnNext(it)
        Loop
    End If

    If selected Then
        'clear the Deliv Point if one
        If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/droid.cpp:2307", "activateGroup", "Unsupported selectedPlayer: %u", selectedPlayer: activateGroup = False: Exit Function
        it = apsFlagPosLists(selectedPlayer).head
        Do While it <> 0
            aFlagPos(lnObj(it)).selected = False
            it = lnNext(it)
        Loop
        groupConsoleInformOfSelection groupNumber
    End If

    If selected Or numDeselected > 0 Then
        If selected Then
            intGroupsChanged groupNumber And &HFF&
        Else
            intGroupsChanged UBYTE_MAX
        End If
    End If

    activateGroup = selected
End Function

Public Sub groupConsoleInformOfSelection(ByVal groupNumber As Long)
    Dim num_selected As Long
    num_selected = selNumSelected(selectedPlayer)
    CONPRINTF ngettext("Group %u selected - %u Unit", "Group %u selected - %u Units", num_selected), groupNumber, num_selected
End Sub

Public Sub groupConsoleInformOfCreation(ByVal groupNumber As Long)
    Dim num_selected As Long
    If Not getWarCamStatus() Then
        num_selected = selNumSelected(selectedPlayer)
        CONPRINTF ngettext("%u unit assigned to Group %u", "%u units assigned to Group %u", num_selected), num_selected, groupNumber
    End If
End Sub

Public Sub groupConsoleInformOfRemoval()
    Dim num_selected As Long
    If Not getWarCamStatus() Then
        num_selected = selNumSelected(selectedPlayer)
        CONPRINTF ngettext("%u units removed from their Group", "%u units removed from their Group", num_selected), num_selected
    End If
End Sub

Public Sub groupConsoleInformOfCentering(ByVal groupNumber As Long)
    Dim num_selected As Long
    num_selected = selNumSelected(selectedPlayer)
    If Not getWarCamStatus() Then
        CONPRINTF ngettext("Centered on Group %u - %u Unit", "Centered on Group %u - %u Units", num_selected), groupNumber, num_selected
    Else
        CONPRINTF ngettext("Aligning with Group %u - %u Unit", "Aligning with Group %u - %u Units", num_selected), groupNumber, num_selected
    End If
End Sub
