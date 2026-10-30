Attribute VB_Name = "droiddef_h"
' src/droiddef.h: the droid templates and struct DROID. DROID_TEMPLATE's
' base class BASE_STATS gives its first fields (see statsdef_h). A
' DROID_TEMPLATE * is a 1-based index into template_c's pool apsTemplates
' (0 = nullptr); as a BASE_STATS * it is STAT_TEMPLATE plus that index.
' The DROID_TEMPLATE methods are template_c's DROID_TEMPLATE_<method>. A
' DROID * is a handle into aDroid (see basedef_h); OrderList is the array
' asOrderList plus asOrderList_size.
Option Explicit

Public Const DROID_MAXCOMP As Long = COMP_NUMCOMPONENTS - 1

' The maximum number of droid weapons
Public Const DROID_DAMAGE_SCALING As Long = 400

'defines how many times to perform the iteration on looking for a blank location
Public Const LOOK_FOR_EMPTY_TILE As Long = 20

Public Type DROID_TEMPLATE
    ' BASE_STATS
    id As String                        ' Text id (i.e. short language-independent name)
    name_ As String                     ' Full / real name of the item
    ref As Long                         ' Unique ID of the item
    index As Long                       ' Index into containing array

    ' The droid components.
    '
    ' This array is indexed by COMPONENT_TYPE so the ECM would be accessed
    ' using asParts[COMP_ECM].
    '
    ' Weapons are stored in asWeaps, _not_ here at index COMP_WEAPON! (Which is the reason we do not have a COMP_NUMCOMPONENTS sized array here.)
    asParts(0 To DROID_MAXCOMP - 1) As Long     ' uint8_t
    ' The weapon systems
    numWeaps As Long                    ' int8_t: Number of weapons
    asWeaps(0 To MAX_WEAPONS - 1) As Long       ' uint32_t: weapon indices
    droidType As DROID_TYPE             ' The type of droid
    multiPlayerID As Long               ' multiplayer unique descriptor(cant use id's for templates). Used for save games as well now - AB 29/10/98
    prefab As Boolean                   ' Not player designed, not saved, never delete or change
    stored As Boolean                   ' Stored template
    enabled As Boolean                  ' Has been enabled
End Type

Public Type DROID
    ' SIMPLE_OBJECT
    typ As OBJECT_TYPE                  ' The type of object
    id As Long                          ' uint32_t: ID number of the object
    pos As Vector3i                     ' Position of the object
    rot As Rotation                     ' Object's yaw +ve rotation around up-axis
    player As Long                      ' uint8_t: Which player the object belongs to
    born As Long                        ' uint32_t: Time the game object was born
    died As Long                        ' uint32_t: When an object was destroyed, if 0 still alive
    time_ As Long                       ' uint32_t: Game time of given space-time position.
    ' BASE_OBJECT
    sDisplay As SCREEN_DISP_DATA        ' screen coordinate details
    group As Long                       ' UBYTE: Which group selection is the droid currently in?
    repairGroup As Long                 ' UBYTE: Group to be set when the unit has been repaired
    selected As Long                    ' UBYTE: Whether the object is selected (might want this elsewhere)
    visible(0 To MAX_PLAYERS - 1) As Long       ' UBYTE: Whether object is visible to specific player
    seenThisTick(0 To MAX_PLAYERS - 1) As Long  ' UBYTE: Whether object has been seen this tick by the specific player.
    lastEmission As Long                ' When did it last puff out smoke?
    lastHitWeapon As WEAPON_SUBCLASS    ' The weapon that last hit it
    timeLastHit As Long                 ' The time the structure was last attacked
    body As Long                        ' Hit points with lame name
    periodicalDamageStart As Long       ' When the object entered the fire
    periodicalDamage As Long            ' How much damage has been done since the object entered the fire
    watchedTiles() As TILEPOS           ' Variable size array of watched tiles, empty for features
    watchedTiles_size As Long
    timeAnimationStarted As Long        ' DISPLAY-ONLY: Animation start time, zero for do not animate
    animationEvent As Long              ' DISPLAY-ONLY: UBYTE: If animation start time > 0, this points to which animation to run
    numWeaps As Long                    ' unsigned
    asWeaps(0 To MAX_WEAPONS - 1) As WEAPON
    flags As Long                       ' std::bitset<OBJECT_FLAG_COUNT>
    hasExtraFunction As Boolean         ' Does this object include some extra functionality?
    ' DROID
    aName As String                     ' char[MAX_STR_LENGTH]: UTF-8 name of the droid. This is generated from the droid template
    droidType As DROID_TYPE             ' The type of droid
    asBits(0 To DROID_MAXCOMP - 1) As Long  ' uint8_t: Holds the specifics for the component parts. Indexed by COMPONENT_TYPE.
    weight As Long
    baseSpeed As Long                   ' the base speed dependent on propulsion type
    originalBody As Long                ' the original body points
    experience As Long                  ' uint32_t
    shieldPoints As Long                ' Shield points of droid, which will be drained instead of health (default: -1 if no shields)
    shieldRegenTime As Long             ' How long should it be before the next regeneration step
    shieldInterruptRegenTime As Long    ' Standby time in case the shield was destroyed to begin regenerating again
    kills As Long                       ' uint32_t
    lastFrustratedTime As Long          ' Set when eg being stuck; used for eg firing indiscriminately at map features to clear the way
    resistance As Long                  ' SWORD: used in Electronic Warfare
    psGroup As Long                     ' DROID_GROUP *: The group the droid belongs to
    psBaseStruct As Long                ' STRUCTURE *: a structure that this droid might be associated with. For VTOLs this is the rearming pad
    listSize As Long                    ' Gives the number of synchronised orders.
    asOrderList() As DROID_ORDER_DATA   ' OrderList (std::vector<DROID_ORDER_DATA>)
    asOrderList_size As Long
    listPendingBegin As Long            ' unsigned: Index of first order which will not be erased by a pending order.
    order As DROID_ORDER_DATA           ' Order data
    secondaryOrder As Long              ' secondary order data
    secondaryOrderPending As Long       ' uint32_t: What the secondary order will be, after synchronisation.
    secondaryOrderPendingCount As Long  ' Number of pending secondary order changes.
    action As DROID_ACTION              ' Action data
    actionPos As Vector2i
    psActionTarget(0 To MAX_WEAPONS - 1) As Long    ' BASE_OBJECT *: Action target object
    lastCheckNearestTargetFailed(0 To MAX_WEAPONS - 1) As Long  ' Set to the last gameTime that aiBestNearestTarget was called on for each weapon slot for this droid (and failed)
    actionStarted As Long               ' Game time action started
    actionPoints As Long                ' number of points done by action since start
    expectedDamageDirect As Long        ' Expected damage to be caused by all currently incoming direct projectiles.
    expectedDamageIndirect As Long      ' Expected damage to be caused by all currently incoming indirect projectiles.
    illumination As Long                ' UBYTE
    sMove As MOVE_CONTROL               ' Movement control data
    prevSpacetime As Spacetime          ' Location of droid in previous tick.
    blockedBits As Long                 ' uint8_t: Bit set telling which tiles block this type of droid (TODO)
    iAudioID As Long                    ' anim data
    heightAboveMap As Long              ' Current calculated height above the terrain (set for VTOL-propulsion units)
    underRepair As Long                 ' uint16_t: Number of other droids / structures currently repairing this unit
End Type

' stats: a BASE_STATS * handle; the DROID_TEMPLATE * (0: nullptr)
Public Function castDroidTemplate(ByVal stats As Long) As Long
    If stats <> 0 Then
        If BASE_STATS_hasType(stats, STAT_TEMPLATE) Then
            castDroidTemplate = stats And &HFFFF&
            Exit Function
        End If
    End If
    castDroidTemplate = 0
End Function
