Attribute VB_Name = "structuredef_h"
' src/structuredef.h: so far the structure types and the structure stats
' (the structures themselves and their functionality come with cluster B).
' STRUCTURE_STATS's base class BASE_STATS gives its first fields (see
' statsdef_h); a STRUCTURE_STATS * is its index in asStructureStats, -1 =
' nullptr, and so are its ECM_STATS *, SENSOR_STATS * and WEAPON_STATS *.
' uint64_t flags is a Long (only STRUCTURE_CONNECTED is defined). Not
' ported yet: STRUCTURE_STATS::size (snapDirection).
' The production runs: ProductionRunEntry (its methods are
' ProductionRunEntry_<method>; operator== is structure_c's
' ProductionRunEntry_equals), ProductionRun (std::vector<ProductionRunEntry>:
' items with a size) and ProductionRunList (std::vector<ProductionRun>, one
' element of asProductionRun[NUM_FACTORY_TYPES]).
Option Explicit

Public Const NUM_FACTORY_MODULES As Long = 2
Public Const NUM_POWER_MODULES As Long = 4
Public Const REF_ANY As Long = 255         ' Used to indicate any kind of building when calling intGotoNextStructureType()

' Defines for indexing an appropriate IMD object given a buildings purpose.
Public Enum STRUCTURE_TYPE
    REF_HQ
    REF_FACTORY
    REF_FACTORY_MODULE          'draw as factory 2
    REF_POWER_GEN
    REF_POWER_MODULE
    REF_RESOURCE_EXTRACTOR
    REF_DEFENSE
    REF_WALL
    REF_WALLCORNER              'corner wall - no gun
    REF_GENERIC
    REF_RESEARCH
    REF_RESEARCH_MODULE
    REF_REPAIR_FACILITY
    REF_COMMAND_CONTROL         'control centre for command droids
    REF_BRIDGE                  'NOT USED, but removing it would change savegames
    REF_DEMOLISH                'the demolish structure type - should only be one stat with this type
    REF_CYBORG_FACTORY
    REF_VTOL_FACTORY
    REF_LAB
    REF_REARM_PAD
    REF_MISSILE_SILO
    REF_SAT_UPLINK              'added for updates - AB 8/6/99
    REF_GATE
    REF_LASSAT
    REF_FORTRESS                'added in WZ 4.5.0
    NUM_DIFF_BUILDINGS          'need to keep a count of how many types for IMD loading
End Enum

Public Enum STRUCT_STRENGTH
    STRENGTH_SOFT
    STRENGTH_MEDIUM
    STRENGTH_HARD
    STRENGTH_BUNKER

    NUM_STRUCT_STRENGTH
End Enum

' typedef UWORD STRUCTSTRENGTH_MODIFIER (a Long)

Public Const SAS_OPEN_SPEED As Long = GAME_TICKS_PER_SEC
Public Const SAS_STAY_OPEN_TIME As Long = GAME_TICKS_PER_SEC * 6

Public Enum STRUCT_ANIM_STATES
    SAS_NORMAL
    SAS_OPEN
    SAS_OPENING
    SAS_CLOSING
End Enum

Public Const STRUCTURE_CONNECTED As Long = &H1     ' This structure must be built side by side with another of the same player

' STRUCTURE_STATS's upgrade struct
Public Type STRUCTURE_STATS_UPGRADE
    research As Long
    moduleResearch As Long
    repair As Long
    power As Long
    modulePower As Long
    production As Long
    moduleProduction As Long
    rearm As Long
    armour As Long
    thermal As Long
    hitpoints As Long
    resistance As Long          ' resist enemy takeover; 0 = immune
    limit As Long               ' current max limit for this type, LOTS_OF = no limit
End Type

'this structure is used to hold the permanent stats for each type of building
Public Type STRUCTURE_STATS
    ' BASE_STATS
    id As String                ' Text id (i.e. short language-independent name)
    name_ As String             ' Full / real name of the item
    ref As Long                 ' Unique ID of the item
    index As Long               ' Index into containing array

    typ As STRUCTURE_TYPE       ' the type of structure
    strength As STRUCT_STRENGTH ' strength against the weapon effects
    baseWidth As Long           'The width of the base in tiles
    baseBreadth As Long         'The breadth of the base in tiles
    buildPoints As Long         'The number of build points required to build the structure
    height As Long              'The height above/below the terrain - negative values denote below the terrain
    powerToBuild As Long        'How much power the structure requires to build
    pIMD() As Long              ' The IMDs to draw for this structure, for each possible number of modules.
    pIMD_size As Long
    pBaseIMD As Long            'The base IMD to draw for this structure
    pECM As Long                'Which ECM is standard for the structure -if any
    pSensor As Long             'Which Sensor is standard for the structure -if any
    weaponSlots As Long         'Number of weapons that can be attached to the building
    numWeaps As Long            'Number of weapons for default
    psWeapStat(0 To MAX_WEAPONS - 1) As Long
    flags As Long               ' uint64_t
    combinesWithWall As Boolean 'If the structure will trigger nearby walls to try combining with it

    minLimit As Long            ' lowest value user can set limit to (currently unused)
    maxLimit As Long            ' highest value user can set limit to, LOTS_OF = no limit
    curCount(0 To MAX_PLAYERS - 1) As Long   ' current number of instances of this type

    upgrade(0 To MAX_PLAYERS - 1) As STRUCTURE_STATS_UPGRADE
    base As STRUCTURE_STATS_UPGRADE
    isFavorite As Boolean       ' on Favorites list
End Type

Public Enum STRUCT_STATES
    SS_BEING_BUILT
    SS_BUILT
    SS_BLUEPRINT_VALID
    SS_BLUEPRINT_INVALID
    SS_BLUEPRINT_PLANNED
    SS_BLUEPRINT_PLANNED_BY_ALLY
End Enum

Public Enum StatusPending
    FACTORY_NOTHING_PENDING = 0
    FACTORY_START_PENDING
    FACTORY_HOLD_PENDING
    FACTORY_CANCEL_PENDING
End Enum

Public Const LOTS_OF As Long = &HFFFFFFFF   ' highest number the limit can be set to

' positiondef.h's OBJECT_POSITION fields first
Public Type FLAG_POSITION
    typ As POSITION_TYPE                ' the type of position obj - FlagPos or ProxDisp
    frameNumber As Long                 ' when the Position was last drawn
    screenX As Long                     ' screen coords and radius of Position imd
    screenY As Long
    screenR As Long
    player As Long                      ' which player the Position belongs to
    selected As Boolean                 ' flag to indicate whether the Position is to be highlighted
    ' FLAG_POSITION
    coords As Vector3i                  ' the world coords of the Position
    factoryInc As Long                  ' UBYTE: indicates whether the first, second etc factory
    factoryType As Long                 ' UBYTE: indicates whether standard, cyborg or vtol factory
End Type

Public Type RESEARCH_FACILITY
    psSubject As Long                   ' RESEARCH *: The subject the structure is working on.
    psSubjectPending As Long            ' RESEARCH *: The subject the structure is going to work on when the GAME_RESEARCHSTATUS message is received.
    statusPending As StatusPending      ' Pending = not yet synchronised.
    pendingCount As Long                ' unsigned: Number of messages sent but not yet processed.
    psBestTopic As Long                 ' RESEARCH *: The topic with the most research points that was last performed
    timeStartHold As Long               ' The time the research facility was put on hold
End Type

Public Type FACTORY
    productionLoops As Long             ' uint8_t: Number of loops to perform. Not synchronised, and only meaningful for selectedPlayer.
    loopsPerformed As Long              ' UBYTE: how many times the loop has been performed
    psSubject As Long                   ' DROID_TEMPLATE *: The subject the structure is working on.
    psSubjectPending As Long            ' DROID_TEMPLATE *: The subject the structure is going to working on. (Pending = not yet synchronised.)
    statusPending As StatusPending      ' Pending = not yet synchronised.
    pendingCount As Long                ' unsigned: Number of messages sent but not yet processed.
    timeStarted As Long                 ' The time the building started on the subject
    buildPointsRemaining As Long        ' Build points required to finish building the droid.
    timeStartHold As Long               ' The time the factory was put on hold
    psAssemblyPoint As Long             ' FLAG_POSITION *: Place for the new droids to assemble at
    psCommander As Long                 ' DROID *: command droid to produce droids for (if any)
    secondaryOrder As Long              ' uint32_t: Secondary order state for all units coming out of the factory.
End Type

Public Type RES_EXTRACTOR
    psPowerGen As Long                  ' STRUCTURE *: owning power generator
End Type

Public Type POWER_GEN
    apResExtractors(0 To NUM_POWER_MODULES - 1) As Long ' STRUCTURE *: Pointers to associated oil derricks
End Type

' enum class RepairEvents
Public Enum RepairEvents
    RepairEvents_NoEvents
    RepairEvents_RepairTargetFound
    RepairEvents_UnitReachedMaxHP
    RepairEvents_UnitDied
    RepairEvents_UnitMovedAway
End Enum

' enum class RepairState
Public Enum RepairState
    RepairState_Invalid = -1
    RepairState_Idle
    RepairState_Repairing
End Enum

Public Type REPAIR_FACILITY
    psObj As Long                       ' BASE_OBJECT *: Object being repaired
    psDeliveryPoint As Long             ' FLAG_POSITION *: Place for the repaired droids to assemble at
    state As RepairState
End Type

Public Type REARM_PAD
    timeStarted As Long                 ' Time reArm started on current object
    psObj As Long                       ' BASE_OBJECT *: Object being rearmed
    timeLastUpdated As Long             ' Time rearm was last updated
End Type

Public Type WALL
    typ As Long                         ' unsigned: Type of wall, 0 = ─, 1 = ┼, 2 = ┴, 3 = ┘.
End Type

' union FUNCTIONALITY: the members do not overlap in VB. A FUNCTIONALITY * is
' the handle of the STRUCTURE that holds it (STRUCTURE::functionality), 0 =
' nullptr; STRUCTURE_calloc_functionality gives calloc's zeroes (nullptr is
' -1 for a RESEARCH *).
Public Type FUNCTIONALITY
    researchFacility As RESEARCH_FACILITY
    factory As FACTORY
    resourceExtractor As RES_EXTRACTOR
    powerGenerator As POWER_GEN
    repairFacility As REPAIR_FACILITY
    rearmPad As REARM_PAD
    wall As WALL
End Type

'this structure is used whenever an instance of a building is required in game
Public Type STRUCTURE
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
    ' STRUCTURE
    pStructureType As Long              ' STRUCTURE_STATS *: pointer to the structure stats for this type of building
    status As STRUCT_STATES             ' defines whether the structure is being built, doing nothing or performing a function
    currentBuildPts As Long             ' uint32_t: the build points currently assigned to this structure
    resistance As Long                  ' current resistance points, 0 = cannot be attacked electrically
    lastResistance As Long              ' time the resistance was last increased
    pFunctionality As Long              ' FUNCTIONALITY *: this structure's handle when it has functionality, else 0
    functionality As FUNCTIONALITY      ' (VB) what pFunctionality points to
    buildRate As Long                   ' Rate that this structure is being built, calculated each tick. Only meaningful if status == SS_BEING_BUILT. If construction hasn't started and build rate is 0, remove the structure.
    lastBuildRate As Long               ' Needed if wanting the buildRate between buildRate being reset to 0 each tick and the trucks calculating it.
    psTarget(0 To MAX_WEAPONS - 1) As Long  ' BASE_OBJECT *
    expectedDamage As Long              ' Expected damage to be caused by all currently incoming projectiles.
    prevTime As Long                    ' uint32_t: Time of structure's previous tick.
    foundationDepth As Single           ' Depth of structure's foundation		// DISPLAY-ONLY
    capacity As Long                    ' uint8_t: Lame name: current number of module upgrades (*not* maximum nb of upgrades)
    state As STRUCT_ANIM_STATES
    lastStateTime As Long
    prebuiltImd As Long                 ' iIMDBaseShape *
    productToGroup As Long              ' UBYTE
End Type

'the three different types of factory (currently) - FACTORY, CYBORG_FACTORY, VTOL_FACTORY
' added repair facilities as they need an assembly point as well
Public Enum FLAG_TYPE
    FACTORY_FLAG
    CYBORG_FLAG
    VTOL_FLAG
    REPAIR_FLAG
    'separate the numfactory from numflag
    NUM_FLAG_TYPES
    NUM_FACTORY_TYPES = REPAIR_FLAG
End Enum

'this is used for module graphics - factory and vtol factory
Public Const NUM_FACMOD_TYPES As Long = 2

Public Type ProductionRunEntry
    quantity As Long                    ' number to build
    built As Long                       ' number built on current run
    psTemplate As Long                  ' DROID_TEMPLATE *: template to build
End Type

' typedef std::vector<ProductionRunEntry> ProductionRun
Public Type ProductionRun
    items() As ProductionRunEntry
    size As Long
End Type

' (VB) std::vector<ProductionRun>
Public Type ProductionRunList
    runs() As ProductionRun
    size As Long
End Type

' Vector2i STRUCTURE_STATS::size(uint16_t direction) const
Public Function STRUCTURE_STATS_size(ByVal psStats As Long, ByVal direction As Long) As Vector2i
    Dim t As Long
    STRUCTURE_STATS_size.x = asStructureStats(psStats).baseWidth
    STRUCTURE_STATS_size.y = asStructureStats(psStats).baseBreadth
    If (snapDirection(direction) And &H4000&) <> 0 Then ' if building is rotated left or right by 90°, swap width and height
        t = STRUCTURE_STATS_size.x
        STRUCTURE_STATS_size.x = STRUCTURE_STATS_size.y
        STRUCTURE_STATS_size.y = t
    End If
End Function

' Vector2i STRUCTURE::size() const
Public Function STRUCTURE_size(ByVal psStruct As Long) As Vector2i
    STRUCTURE_size = STRUCTURE_STATS_size(aStruct(psStruct).pStructureType, aStruct(psStruct).rot.direction)
End Function

' (VB) the constructor: STRUCTURE_STATS() : pBaseIMD(nullptr), pECM(nullptr), pSensor(nullptr), curCount and upgrade zeroed
Public Sub STRUCTURE_STATS_construct(s As STRUCTURE_STATS)
    s.pBaseIMD = 0
    s.pECM = -1
    s.pSensor = -1
End Sub

' (VB) a STRUCTURE_STATS * (its index, -1 = nullptr) as a BASE_STATS * handle (0 = nullptr), as C++ converts it implicitly
Public Function structureStatsBase(ByVal psStats As Long) As Long
    If psStats >= 0 Then
        structureStatsBase = STAT_STRUCTURE + psStats
    Else
        structureStatsBase = 0
    End If
End Function

' stats: a BASE_STATS * handle; the STRUCTURE_STATS's index, or -1
Public Function castStructureStats(ByVal stats As Long) As Long
    If stats <> 0 Then
        If BASE_STATS_hasType(stats, STAT_STRUCTURE) Then
            castStructureStats = stats And &HFFFF&
            Exit Function
        End If
    End If
    castStructureStats = -1
End Function

' ProductionRunEntry's methods (ProductionRunEntry() : quantity(0), built(0), psTemplate(nullptr) is a blank UDT)
Public Sub ProductionRunEntry_restart(e As ProductionRunEntry)
    e.built = 0
End Sub

Public Sub ProductionRunEntry_removeComplete(e As ProductionRunEntry)
    e.quantity = e.quantity - e.built
    e.built = 0
End Sub

Public Function ProductionRunEntry_numRemaining(e As ProductionRunEntry) As Long
    ProductionRunEntry_numRemaining = e.quantity - e.built
End Function

Public Function ProductionRunEntry_isComplete(e As ProductionRunEntry) As Boolean
    ProductionRunEntry_isComplete = ProductionRunEntry_numRemaining(e) <= 0
End Function

Public Function ProductionRunEntry_isValid(e As ProductionRunEntry) As Boolean
    ProductionRunEntry_isValid = e.psTemplate <> 0 And e.quantity > 0 And e.built <= e.quantity
End Function
