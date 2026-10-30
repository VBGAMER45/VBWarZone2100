Attribute VB_Name = "featuredef_h"
' src/featuredef.h: the feature types, stats and struct FEATURE.
' FEATURE_STATS's base class BASE_STATS gives its first fields (see
' statsdef_h); a FEATURE_STATS * is its index in asFeatureStats, -1 =
' nullptr. A FEATURE * is a handle into aFeature (see basedef_h).
Option Explicit

Public Enum FEATURE_TYPE
    FEAT_TANK = 2       ' hack to keep enums the same value
    FEAT_GEN_ARTE
    FEAT_OIL_RESOURCE
    FEAT_BOULDER
    FEAT_VEHICLE
    FEAT_BUILDING
    FEAT_UNUSED
    FEAT_LOS_OBJ
    FEAT_OIL_DRUM
    FEAT_TREE
    FEAT_SKYSCRAPER
    FEAT_COUNT
End Enum

' Stats for a feature
Public Type FEATURE_STATS
    ' BASE_STATS
    id As String                ' Text id (i.e. short language-independent name)
    name_ As String             ' Full / real name of the item
    ref As Long                 ' Unique ID of the item
    index As Long               ' Index into containing array

    subType As FEATURE_TYPE     ' type of feature

    psImd As Long               ' Graphic for the feature
    baseWidth As Long           ' UWORD: The width of the base in tiles
    baseBreadth As Long         ' UWORD: The breadth of the base in tiles

    tileDraw As Boolean         ' Whether the tile needs to be drawn
    allowLOS As Boolean         ' Whether the feature allows the LOS. true = can see through the feature
    visibleAtStart As Boolean   ' Whether the feature is visible at the start of the mission
    damageable As Boolean       ' Whether the feature can be destroyed
    body As Long                ' Number of body points
    armourValue As Long         ' Feature armour
End Type

Public Type FEATURE
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
    ' FEATURE
    psStats As Long                     ' FEATURE_STATS const *
    foundationDepth As Single           ' Depth of features's foundation		// DISPLAY-ONLY
End Type

' (VB) the constructor FEATURE_STATS(int idx = 0) : BASE_STATS(idx) and the default member initialisers
Public Sub FEATURE_STATS_construct(s As FEATURE_STATS, ByVal idx As Long)
    s.ref = idx
    s.subType = FEAT_COUNT
End Sub

' Vector2i FEATURE_STATS::size() const
Public Function FEATURE_STATS_size(ByVal psStats As Long) As Vector2i
    FEATURE_STATS_size.x = asFeatureStats(psStats).baseWidth
    FEATURE_STATS_size.y = asFeatureStats(psStats).baseBreadth
End Function

' Vector2i FEATURE::size() const
Public Function FEATURE_size(ByVal psFeature As Long) As Vector2i
    FEATURE_size = FEATURE_STATS_size(aFeature(psFeature).psStats)
End Function
