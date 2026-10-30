Attribute VB_Name = "statsdef_h"
' src/statsdef.h: the stats' enums and structs. A struct that derives from
' BASE_STATS and COMPONENT_STATS is one UDT with the base classes' fields
' first (so psStats->buildPower is asWeaponStats(i).buildPower); generic
' access through a BASE_STATS * or COMPONENT_STATS * is stats_c's
' BASE_STATS_<field> and COMPONENT_STATS_<field>. The upgrade structs are
' named after their class (WEAPON_STATS::UPGRADE is WEAPON_STATS_UPGRADE,
' with COMPONENT_STATS::UPGRADE's fields first). A pointer to another stat
' (WEAPON_STATS *psWeaponStat) is its index in its array, -1 = nullptr; an
' iIMDBaseShape * is a model handle, 0 = nullptr. The C++ default member
' initialisers are the <TYPE>_construct procedures, which the arrays'
' allocations call. Not ported: StringToEnum and StringToEnumMap (their
' users come later), the macros getStatsName etc. (written out where used).
Option Explicit

' The different types of droid
' NOTE, if you add to, or change this list then you'll need
' to update the DroidSelectionWeights lookup table in Display.c
Public Enum DROID_TYPE
    DROID_WEAPON            ' Weapon droid
    DROID_SENSOR            ' Sensor droid
    DROID_ECM               ' ECM droid
    DROID_CONSTRUCT         ' Constructor droid
    DROID_PERSON            ' person
    DROID_CYBORG            ' cyborg-type thang
    DROID_TRANSPORTER       ' guess what this is!
    DROID_COMMAND           ' Command droid
    DROID_REPAIR            ' Repair droid
    DROID_DEFAULT           ' Default droid
    DROID_CYBORG_CONSTRUCT  ' cyborg constructor droid - new for update 28/5/99
    DROID_CYBORG_REPAIR     ' cyborg repair droid - new for update 28/5/99
    DROID_CYBORG_SUPER      ' cyborg repair droid - new for update 7/6/99
    DROID_SUPERTRANSPORTER  ' SuperTransport (MP)
    DROID_ANY               ' Any droid. Used as a parameter for various stuff.
End Enum

Public Enum COMPONENT_TYPE
    COMP_BODY
    COMP_BRAIN
    COMP_PROPULSION
    COMP_REPAIRUNIT
    COMP_ECM
    COMP_SENSOR
    COMP_CONSTRUCT
    COMP_WEAPON
    COMP_NUMCOMPONENTS      ' The number of enumerators in this enum.
End Enum

Public Enum WEAPON_FLAGS
    WEAPON_FLAG_NO_FRIENDLY_FIRE
    WEAPON_FLAG_ALLOWED_ON_TRANSPORTER
    WEAPON_FLAG_TELEPORT_CAPTURE
    WEAPON_FLAG_COUNT
End Enum

' LOC used for holding locations for Sensors and ECM's
Public Enum LOC
    LOC_DEFAULT
    LOC_TURRET
End Enum

' SIZE used for specifying body size
Public Enum BODY_SIZE
    SIZE_LIGHT
    SIZE_MEDIUM
    SIZE_HEAVY
    SIZE_SUPER_HEAVY
    SIZE_NUM
End Enum

' SIZE used for specifying weapon size
Public Enum WEAPON_SIZE
    WEAPON_SIZE_LIGHT
    WEAPON_SIZE_HEAVY
    WEAPON_SIZE_ANY
    WEAPON_SIZE_NUM
End Enum

' Basic weapon type
Public Enum WEAPON_CLASS
    WC_KINETIC              ' bullets etc
    WC_HEAT                 ' laser etc
    WC_NUM_WEAPON_CLASSES   ' The number of enumerators in this enum.
End Enum

' weapon subclasses used to define which weapons are affected by weapon upgrade
' functions
Public Enum WEAPON_SUBCLASS
    WSC_MGUN
    WSC_CANNON
    WSC_MORTARS
    WSC_MISSILE
    WSC_ROCKET
    WSC_ENERGY
    WSC_GAUSS
    WSC_FLAME
    'WSC_CLOSECOMBAT,
    WSC_HOWITZERS
    WSC_ELECTRONIC
    WSC_AAGUN
    WSC_SLOWMISSILE
    WSC_SLOWROCKET
    WSC_LAS_SAT
    WSC_BOMB
    WSC_COMMAND
    WSC_EMP
    WSC_NUM_WEAPON_SUBCLASSES   ' The number of enumerators in this enum.
End Enum

' Used to define which projectile model to use for the weapon.
Public Enum MOVEMENT_MODEL
    MM_DIRECT
    MM_INDIRECT
    MM_HOMINGDIRECT
    MM_HOMINGINDIRECT
End Enum

' Used to modify the damage to a propuslion type (or structure) based on
' weapon.
Public Enum WEAPON_EFFECT
    WE_ANTI_PERSONNEL
    WE_ANTI_TANK
    WE_BUNKER_BUSTER
    WE_ARTILLERY_ROUND
    WE_FLAMER
    WE_ANTI_AIRCRAFT
    WE_NUMEFFECTS           ' The number of enumerators in this enum.
End Enum

' Defines the left and right sides for propulsion IMDs
Public Enum PROP_SIDE
    LEFT_PROP
    RIGHT_PROP
    NUM_PROP_SIDES          ' The number of enumerators in this enum.
End Enum

Public Enum PROPULSION_TYPE
    PROPULSION_TYPE_WHEELED
    PROPULSION_TYPE_TRACKED
    PROPULSION_TYPE_LEGGED
    PROPULSION_TYPE_HOVER
    PROPULSION_TYPE_LIFT
    PROPULSION_TYPE_PROPELLOR
    PROPULSION_TYPE_HALF_TRACKED
    PROPULSION_TYPE_NUM     ' The number of enumerators in this enum.
End Enum

Public Enum SENSOR_TYPE
    STANDARD_SENSOR
    INDIRECT_CB_SENSOR
    VTOL_CB_SENSOR
    VTOL_INTERCEPT_SENSOR
    SUPER_SENSOR            ' works as all of the above together! - new for updates
    RADAR_DETECTOR_SENSOR
End Enum

Public Enum TRAVEL_MEDIUM
    GROUND
    AIR
End Enum

' Stats structures type definitions

' Elements common to all stats structures

' What number the ref numbers start at for each type of stat
Public Enum StatType
    STAT_BODY = &H10000
    STAT_BRAIN = &H20000
    STAT_PROPULSION = &H40000
    STAT_SENSOR = &H50000
    STAT_ECM = &H60000
    STAT_REPAIR = &H80000
    STAT_WEAPON = &HA0000
    STAT_RESEARCH = &HB0000
    STAT_TEMPLATE = &HC0000
    STAT_STRUCTURE = &HD0000
    STAT_FUNCTION = &HE0000
    STAT_CONSTRUCT = &HF0000
    STAT_FEATURE = &H100000

    ' The maximum number of refs for a type of stat
    STAT_MASK = &HFFFF0000
End Enum

' (VB) the BASE_STATS * of a local stats object whose ref is 0 (structure.cpp's
' findAssemblyPointPosition: STRUCTURE_STATS sStats; sStats.ref = 0): no StatType
Public Const DUMMY_STATS_REF0 As Long = &H7FFF0000

' Classify some components for more detailed template checking in normal multiplayer
' (enum class UsageClass)
Public Enum UsageClass
    UsageClass_Other
    UsageClass_Cyborg
    UsageClass_SuperCyborg
End Enum

' Stats structures type definitions
Public Const SHOOT_ON_GROUND As Long = &H1
Public Const SHOOT_IN_AIR As Long = &H2

' COMPONENT_STATS::UPGRADE
Public Type COMPONENT_STATS_UPGRADE
    hitpoints As Long                   ' Number of upgradeable hitpoints
    hitpointPct As Long                 ' Adjust final droid hitpoints by this percentage amount
End Type

Public Type PROPULSION_STATS_UPGRADE
    hitpoints As Long
    hitpointPct As Long
    hitpointPctOfBody As Long           ' Increase hitpoints by this percentage of the body's hitpoints
End Type

Public Type PROPULSION_STATS
    ' BASE_STATS
    id As String                        ' Text id (i.e. short language-independent name)
    name_ As String                     ' Full / real name of the item
    ref As Long                         ' Unique ID of the item
    index As Long                       ' Index into containing array
    ' COMPONENT_STATS
    pIMD As Long                        ' The IMD to draw for this component
    buildPower As Long                  ' Power required to build the component
    buildPoints As Long                 ' Time required to build the component
    weight As Long                      ' Component's weight
    compType As COMPONENT_TYPE
    droidTypeOverride As DROID_TYPE
    designable As Boolean               ' Flag to indicate whether this component can be used in the design screen
    usageClass As UsageClass            ' Used to check if certain template designs are legit for multiplayer

    propulsionType As PROPULSION_TYPE
    maxSpeed As Long                    ' Max speed for the droid
    turnSpeed As Long
    spinSpeed As Long
    spinAngle As Long
    skidDeceleration As Long
    deceleration As Long
    acceleration As Long

    upgrade(0 To MAX_PLAYERS - 1) As PROPULSION_STATS_UPGRADE
    base As PROPULSION_STATS_UPGRADE
End Type

Public Type SENSOR_STATS_UPGRADE
    hitpoints As Long
    hitpointPct As Long
    range As Long
End Type

Public Type SENSOR_STATS
    ' BASE_STATS
    id As String
    name_ As String
    ref As Long
    index As Long
    ' COMPONENT_STATS
    pIMD As Long
    buildPower As Long
    buildPoints As Long
    weight As Long
    compType As COMPONENT_TYPE
    droidTypeOverride As DROID_TYPE
    designable As Boolean
    usageClass As UsageClass

    pMountGraphic As Long               ' The turret mount to use
    location As Long                    ' specifies whether the Sensor is default or for the Turret
    typ As SENSOR_TYPE                  ' used for combat

    upgrade(0 To MAX_PLAYERS - 1) As SENSOR_STATS_UPGRADE
    base As SENSOR_STATS_UPGRADE
End Type

Public Type ECM_STATS_UPGRADE
    hitpoints As Long
    hitpointPct As Long
    range As Long
End Type

Public Type ECM_STATS
    ' BASE_STATS
    id As String
    name_ As String
    ref As Long
    index As Long
    ' COMPONENT_STATS
    pIMD As Long
    buildPower As Long
    buildPoints As Long
    weight As Long
    compType As COMPONENT_TYPE
    droidTypeOverride As DROID_TYPE
    designable As Boolean
    usageClass As UsageClass

    pMountGraphic As Long               ' The turret mount to use
    location As Long                    ' Specifies whether the ECM is default or for the Turret

    upgrade(0 To MAX_PLAYERS - 1) As ECM_STATS_UPGRADE
    base As ECM_STATS_UPGRADE
End Type

Public Type REPAIR_STATS_UPGRADE
    hitpoints As Long
    hitpointPct As Long
    repairPoints As Long                ' The number of points contributed each cycle
End Type

Public Type REPAIR_STATS
    ' BASE_STATS
    id As String
    name_ As String
    ref As Long
    index As Long
    ' COMPONENT_STATS
    pIMD As Long
    buildPower As Long
    buildPoints As Long
    weight As Long
    compType As COMPONENT_TYPE
    droidTypeOverride As DROID_TYPE
    designable As Boolean
    usageClass As UsageClass

    pMountGraphic As Long               ' The turret mount to use
    location As Long                    ' Specifies whether the Repair is default or for the Turret
    time_ As Long                       ' Time delay for repair cycle

    upgrade(0 To MAX_PLAYERS - 1) As REPAIR_STATS_UPGRADE
    base As REPAIR_STATS_UPGRADE
End Type

Public Type WEAPON_STATS_UPGRADE
    hitpoints As Long
    hitpointPct As Long
    shortRange As Long
    maxRange As Long                    ' Max distance to target for long range shot
    minRange As Long                    ' Min distance to target for shot
    hitChance As Long                   ' Chance to hit at
    shortHitChance As Long
    firePause As Long                   ' Pause between each shot
    numRounds As Long                   ' uint8_t: The number of rounds per salvo
    reloadTime As Long                  ' Time to reload the round of ammo
    damage As Long
    empRadius As Long                   ' EMP blast radius -- does no damage
    radius As Long                      ' Basic blast radius of weapon
    radiusDamage As Long                ' "Splash damage"
    periodicalDamage As Long            ' Repeat damage each second after hit
    periodicalDamageRadius As Long      ' Repeat damage radius
    periodicalDamageTime As Long        ' How long the round keeps damaging
    minimumDamage As Long               ' Minimum amount of damage done, in percentage of damage
End Type

Public Type WEAPON_STATS
    ' BASE_STATS
    id As String
    name_ As String
    ref As Long
    index As Long
    ' COMPONENT_STATS
    pIMD As Long
    buildPower As Long
    buildPoints As Long
    weight As Long
    compType As COMPONENT_TYPE
    droidTypeOverride As DROID_TYPE
    designable As Boolean
    usageClass As UsageClass

    base As WEAPON_STATS_UPGRADE
    upgrade(0 To MAX_PLAYERS - 1) As WEAPON_STATS_UPGRADE

    periodicalDamageWeaponClass As WEAPON_CLASS         ' Periodical damage weapon class by damage type (KINETIC, HEAT)
    periodicalDamageWeaponSubClass As WEAPON_SUBCLASS   ' Periodical damage weapon subclass (research class)
    periodicalDamageWeaponEffect As WEAPON_EFFECT       ' Periodical damage weapon effect (propulsion/body damage modifier)

    weaponClass As WEAPON_CLASS         ' the class of weapon  (KINETIC, HEAT)
    weaponSubClass As WEAPON_SUBCLASS   ' the subclass to which the weapon belongs (research class)
    movementModel As MOVEMENT_MODEL     ' which projectile model to use for the bullet
    weaponEffect As WEAPON_EFFECT       ' which type of warhead is associated with the weapon (propulsion/body damage modifier)
    weaponSize As WEAPON_SIZE           ' eg light weapons can be put on light bodies or as sidearms
    flightSpeed As Long                 ' speed ammo travels at
    recoilValue As Long                 ' used to compare with weight to see if recoils or not
    distanceExtensionFactor As Long     ' max extra distance a projectile can travel if misses target
    rotate As Long                      ' short: amount the weapon(turret) can rotate 0 = none
    maxElevation As Long                ' short: max amount the turret can be elevated up
    minElevation As Long                ' short: min amount the turret can be elevated down
    effectSize As Long                  ' uint16_t: size of the effect 100 = normal, 50 = half etc
    vtolAttackRuns As Long              ' short: number of attack runs a VTOL droid can do with this weapon
    facePlayer As Long                  ' UBYTE: flag to make the (explosion) effect face the player when drawn
    faceInFlight As Long                ' UBYTE: flag to make the inflight effect face the player when drawn
    surfaceToAir As Long                ' UBYTE: indicates how good in the air - SHOOT_ON_GROUND, SHOOT_IN_AIR or both
    lightWorld As Boolean               ' flag to indicate whether the effect lights up the world
    penetrate As Boolean                ' flag to indicate whether pentrate droid or not
    fireOnMove As Boolean               ' indicates whether the droid has to stop before firing

    flags As Long                       ' std::bitset<WEAPON_FLAG_COUNT>: bit WEAPON_FLAG_x

    ' Graphics control stats
    radiusLife As Long                  ' How long a blast radius is visible
    numExplosions As Long               ' The number of explosions per shot

    ' Graphics used for the weapon
    pMountGraphic As Long               ' The turret mount to use
    pMuzzleGraphic As Long              ' The muzzle flash
    pInFlightGraphic As Long            ' The ammo in flight
    pTargetHitGraphic As Long           ' The ammo hitting a target
    pTargetMissGraphic As Long          ' The ammo missing a target
    pWaterHitGraphic As Long            ' The ammo hitting water
    pTrailGraphic As Long               ' The trail used for in flight

    ' Audio
    iAudioFireID As Long
    iAudioImpactID As Long
End Type

Public Type CONSTRUCT_STATS_UPGRADE
    hitpoints As Long
    hitpointPct As Long
    constructPoints As Long             ' The number of points contributed each cycle
End Type

Public Type CONSTRUCT_STATS
    ' BASE_STATS
    id As String
    name_ As String
    ref As Long
    index As Long
    ' COMPONENT_STATS
    pIMD As Long
    buildPower As Long
    buildPoints As Long
    weight As Long
    compType As COMPONENT_TYPE
    droidTypeOverride As DROID_TYPE
    designable As Boolean
    usageClass As UsageClass

    pMountGraphic As Long               ' The turret mount to use

    upgrade(0 To MAX_PLAYERS - 1) As CONSTRUCT_STATS_UPGRADE
    base As CONSTRUCT_STATS_UPGRADE
End Type

Public Type BRAIN_STATS_UPGRADE
    hitpoints As Long
    hitpointPct As Long
    rankThresholds() As Long            ' std::vector<int>
    rankThresholds_size As Long
    maxDroids As Long                   ' base maximum number of droids that the commander can control
    maxDroidsMult As Long               ' maximum number of controlled droids multiplied by level
End Type

Public Type BRAIN_STATS
    ' BASE_STATS
    id As String
    name_ As String
    ref As Long
    index As Long
    ' COMPONENT_STATS
    pIMD As Long
    buildPower As Long
    buildPoints As Long
    weight As Long
    compType As COMPONENT_TYPE
    droidTypeOverride As DROID_TYPE
    designable As Boolean
    usageClass As UsageClass

    psWeaponStat As Long                ' WEAPON_STATS *: weapon stats associated with this brain - for Command Droids

    upgrade(0 To MAX_PLAYERS - 1) As BRAIN_STATS_UPGRADE
    base As BRAIN_STATS_UPGRADE
    rankNames As StringVector           ' std::vector<std::string>
    cmdExpRange() As Long               ' std::vector<int>
    cmdExpRange_size As Long
End Type

Public Type BODY_STATS_UPGRADE
    hitpoints As Long
    hitpointPct As Long
    power As Long                       ' this is the engine output of the body
    armour As Long                      ' A measure of how much protection the armour provides
    thermal As Long
    resistance As Long
End Type

Public Type BODY_STATS
    ' BASE_STATS
    id As String
    name_ As String
    ref As Long
    index As Long
    ' COMPONENT_STATS
    pIMD As Long
    buildPower As Long
    buildPoints As Long
    weight As Long
    compType As COMPONENT_TYPE
    droidTypeOverride As DROID_TYPE
    designable As Boolean
    usageClass As UsageClass

    size As BODY_SIZE                   ' How big the body is - affects how hit
    weaponSlots As Long                 ' The number of weapon slots on the body

    ppIMDList() As Long                 ' list of IMDs to use for propulsion unit - up to numPropulsionStats
    ppIMDList_size As Long
    ppMoveIMDList() As Long             ' list of IMDs to use when droid is moving - up to numPropulsionStats
    ppMoveIMDList_size As Long
    ppStillIMDList() As Long            ' list of IMDs to use when droid is still - up to numPropulsionStats
    ppStillIMDList_size As Long
    bodyClass As String                 ' rules hint to script about its classification

    upgrade(0 To MAX_PLAYERS - 1) As BODY_STATS_UPGRADE
    base As BODY_STATS_UPGRADE
End Type

'***********************************************************************************
'* Additional stats tables
'***********************************************************************************
Public Type PROPULSION_TYPES
    travel As TRAVEL_MEDIUM             ' Which medium the propulsion travels in
    powerRatioMult As Long              ' uint16_t: Multiplier for the calculated power ratio of the droid
    startID As Long                     ' int16_t: sound to play when this prop type starts
    idleID As Long                      ' int16_t: sound to play when this prop type is idle
    moveOffID As Long                   ' int16_t: sound to link moveID and idleID
    moveID As Long                      ' int16_t: sound to play when this prop type is moving
    hissID As Long                      ' int16_t: sound to link moveID and idleID
    shutDownID As Long                  ' int16_t: sound to play when this prop type shuts down
End Type

' typedef uint16_t WEAPON_MODIFIER (a Long)

' (VB) the C++ default member initialisers (and COMPONENT_STATS's), for a blank element

Public Sub PROPULSION_STATS_construct(s As PROPULSION_STATS)
    Dim j As Long
    s.compType = COMP_NUMCOMPONENTS
    s.droidTypeOverride = DROID_ANY
    s.usageClass = UsageClass_Other
    s.propulsionType = PROPULSION_TYPE_NUM
    For j = 0 To MAX_PLAYERS - 1
        s.upgrade(j).hitpointPct = 100
    Next
    s.base.hitpointPct = 100
End Sub

Public Sub SENSOR_STATS_construct(s As SENSOR_STATS)
    Dim j As Long
    s.compType = COMP_NUMCOMPONENTS
    s.droidTypeOverride = DROID_ANY
    s.usageClass = UsageClass_Other
    s.typ = STANDARD_SENSOR
    For j = 0 To MAX_PLAYERS - 1
        s.upgrade(j).hitpointPct = 100
    Next
    s.base.hitpointPct = 100
End Sub

Public Sub ECM_STATS_construct(s As ECM_STATS)
    Dim j As Long
    s.compType = COMP_NUMCOMPONENTS
    s.droidTypeOverride = DROID_ANY
    s.usageClass = UsageClass_Other
    For j = 0 To MAX_PLAYERS - 1
        s.upgrade(j).hitpointPct = 100
    Next
    s.base.hitpointPct = 100
End Sub

Public Sub REPAIR_STATS_construct(s As REPAIR_STATS)
    Dim j As Long
    s.compType = COMP_NUMCOMPONENTS
    s.droidTypeOverride = DROID_ANY
    s.usageClass = UsageClass_Other
    For j = 0 To MAX_PLAYERS - 1
        s.upgrade(j).hitpointPct = 100
    Next
    s.base.hitpointPct = 100
End Sub

Public Sub WEAPON_STATS_construct(s As WEAPON_STATS)
    Dim j As Long
    s.compType = COMP_NUMCOMPONENTS
    s.droidTypeOverride = DROID_ANY
    s.usageClass = UsageClass_Other
    s.base.hitpointPct = 100
    For j = 0 To MAX_PLAYERS - 1
        s.upgrade(j).hitpointPct = 100
    Next
    s.periodicalDamageWeaponClass = WC_NUM_WEAPON_CLASSES
    s.periodicalDamageWeaponSubClass = WSC_NUM_WEAPON_SUBCLASSES
    s.periodicalDamageWeaponEffect = WE_NUMEFFECTS
    s.weaponClass = WC_NUM_WEAPON_CLASSES
    s.weaponSubClass = WSC_NUM_WEAPON_SUBCLASSES
    s.movementModel = MM_DIRECT
    s.weaponEffect = WE_NUMEFFECTS
    s.weaponSize = WEAPON_SIZE_NUM
End Sub

Public Sub CONSTRUCT_STATS_construct(s As CONSTRUCT_STATS)
    Dim j As Long
    s.compType = COMP_NUMCOMPONENTS
    s.droidTypeOverride = DROID_ANY
    s.usageClass = UsageClass_Other
    For j = 0 To MAX_PLAYERS - 1
        s.upgrade(j).hitpointPct = 100
    Next
    s.base.hitpointPct = 100
End Sub

Public Sub BRAIN_STATS_construct(s As BRAIN_STATS)
    Dim j As Long
    s.compType = COMP_NUMCOMPONENTS
    s.droidTypeOverride = DROID_ANY
    s.usageClass = UsageClass_Other
    s.psWeaponStat = -1
    For j = 0 To MAX_PLAYERS - 1
        s.upgrade(j).hitpointPct = 100
    Next
    s.base.hitpointPct = 100
End Sub

Public Sub BODY_STATS_construct(s As BODY_STATS)
    Dim j As Long
    s.compType = COMP_NUMCOMPONENTS
    s.droidTypeOverride = DROID_ANY
    s.usageClass = UsageClass_Other
    s.size = SIZE_NUM
    For j = 0 To MAX_PLAYERS - 1
        s.upgrade(j).hitpointPct = 100
    Next
    s.base.hitpointPct = 100
End Sub
