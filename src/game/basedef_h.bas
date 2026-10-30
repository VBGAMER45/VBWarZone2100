Attribute VB_Name = "basedef_h"
' src/basedef.h: the base object types. With displaydef.h's
' SCREEN_DISP_DATA, weapondef.h's WEAPON and positiondef.h.
'
' The objects (docs/CONVENTIONS.md, "Objects"): DROID, STRUCTURE and FEATURE
' (and later PROJECTILE) are UDTs, each in its own pool
' (paged_entity_container_h's aDroid, aStruct, aFeature), with the
' SIMPLE_OBJECT and BASE_OBJECT fields first, in the C++ order. A pointer to
' an object is a Long handle, 0 = nullptr. Each pool's array starts at
' type * OBJ_HANDLE_UNIT + 1, so one handle is both the DROID * and the
' BASE_OBJECT * (aDroid(h).pos is psDroid->pos), and its pool is
' h \ OBJ_HANDLE_UNIT. A field read through a BASE_OBJECT * or
' SIMPLE_OBJECT * goes through the Property procedures below
' (SIMPLE_OBJECT_<field>, BASE_OBJECT_<field>), which pick the pool.
' SIMPLE_OBJECT::type is a field (typ), as in C++: the destructor changes
' it.
Option Explicit

' the died flag for a droid is set to this when it gets added to the non-current list
Public Const NOT_CURRENT_LIST As Long = 1

' lib/framework/types.h and <cstdint>'s limits
Public Const UBYTE_MAX As Long = 255
Public Const UINT8_MAX As Long = 255
Public Const UWORD_MAX As Long = 65535
Public Const UDWORD_MAX As Long = &HFFFFFFFF
Public Const INT32_MAX As Long = &H7FFFFFFF
Public Const INT32_MIN As Long = &H80000000
Public Const INT16_MAX As Long = 32767
Public Const SDWORD_MAX As Long = &H7FFFFFFF

' frame.h
Public Enum QUEUE_MODE
    ModeQueue       ' Sends a message on the game queue, which will get synchronised, by sending a GAME_ message.
    ModeImmediate   ' Performs the action immediately. Must already have been synchronised, for example by sending a GAME_ message.
End Enum

' frame.h
Public Const PLAYER_FEATURE As Long = MAX_PLAYERS + 1
Public Const MAX_PLAYERS_IN_GUI As Long = MAX_PLAYERS - 1

Public Enum OBJECT_TYPE
    OBJ_DROID           ' Droids
    OBJ_STRUCTURE       ' All Buildings
    OBJ_FEATURE         ' Things like roads, trees, bridges, fires
    OBJ_PROJECTILE      ' Comes out of guns, stupid :-)
    OBJ_NUM_TYPES       ' number of object types - MUST BE LAST
End Enum

' (VB) a handle is type * OBJ_HANDLE_UNIT + 1 + its slot in the type's pool
Public Const OBJ_HANDLE_UNIT As Long = &H10000000

Public Type TILEPOS
    x As Byte
    y As Byte
    typ As Byte
End Type

Public Const MAX_WEAPONS As Long = 3

Public Enum OBJECT_FLAG
    OBJECT_FLAG_JAMMED_TILES
    OBJECT_FLAG_TARGETED
    OBJECT_FLAG_DIRTY
    OBJECT_FLAG_UNSELECTABLE
    OBJECT_FLAG_COUNT
End Enum

' lib/ivis_opengl/ivisdef.h's ANIMATION_EVENTS (WzGfx keeps its own copy)
Public Enum ANIMATION_EVENTS
    ANIM_EVENT_NONE
    ANIM_EVENT_ACTIVE
    ANIM_EVENT_FIRING       ' should not be combined with fire-on-move, as this will look weird
    ANIM_EVENT_DYING
    ANIM_EVENT_COUNT
End Enum

' displaydef.h
Public Type SCREEN_DISP_DATA
    imd As Long                         ' iIMDBaseShape *
    frameNumber As Long                 ' last frame it was drawn
    screenX As Long
    screenY As Long
    screenR As Long
End Type

' weapondef.h: who specified the target?
Public Enum TARGET_ORIGIN
    ORIGIN_UNKNOWN          ' Default value if unknown
    ORIGIN_PLAYER           ' Came directly from player
    ORIGIN_VISUAL           ' Visual targeting
    ORIGIN_ALLY             ' Came from allied unit/structure
    ORIGIN_COMMANDER        ' Came from commander
    ORIGIN_SENSOR           ' Came from standard sensor
    ORIGIN_CB_SENSOR        ' Came from counter-battery sensor
    ORIGIN_AIRDEF_SENSOR    ' Came from Air Defense sensor
    ORIGIN_RADAR_DETECTOR   ' Came from Radar Detector sensor
End Enum

' weapondef.h (WEAPON() zeroes every field, as a fresh UDT does)
Public Type WEAPON
    nStat As Long                       ' uint32_t: Index into the asWeaponStats global array
    ammo As Long                        ' uint32_t
    lastFired As Long                   ' uint32_t: The gametime when this weapon last fired
    shotsFired As Long                  ' uint32_t
    rot As Rotation
    prevRot As Rotation
    usedAmmo As Long                    ' unsigned: Amount of ammunition used up by a VTOL
    origin As TARGET_ORIGIN
End Type

' positiondef.h
Public Enum POSITION_TYPE
    POS_DELIVERY            ' Delivery Points NOT wayPoints
    POS_PROXDATA            ' proximity messages that are data generated
    POS_PROXOBJ             ' proximity messages that are in game generated
    POS_TEMPDELIVERY        ' SAVE ONLY delivery point for factory currently assigned to commander
End Enum

' Space-time coordinate, including orientation.
Public Type Spacetime
    time_ As Long                       ' uint32_t: Game time
    pos As Vector3i                     ' Position of the object
    rot As Rotation                     ' Rotation of the object
End Type

' (VB) imd->max.y of an iIMDBaseShape * (sDisplay.imd)
Public Function imdMaxY(ByVal imd As Long) As Long
    Dim x As Long, z As Long
    iIMDBaseShape_max imd, x, imdMaxY, z
End Function

' (VB) the pool a handle belongs to (C++: the class of *psObj)
Public Function objHandleType(ByVal psObj As Long) As OBJECT_TYPE
    objHandleType = psObj \ OBJ_HANDLE_UNIT
End Function

' ---- SIMPLE_OBJECT's fields through a SIMPLE_OBJECT *, BASE_OBJECT * or PROJECTILE *

Public Property Get SIMPLE_OBJECT_type(ByVal psObj As Long) As OBJECT_TYPE
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: SIMPLE_OBJECT_type = aDroid(psObj).typ
    Case OBJ_STRUCTURE: SIMPLE_OBJECT_type = aStruct(psObj).typ
    Case OBJ_FEATURE: SIMPLE_OBJECT_type = aFeature(psObj).typ
    Case OBJ_PROJECTILE: SIMPLE_OBJECT_type = aProj(psObj).typ
    End Select
End Property

Public Property Get SIMPLE_OBJECT_id(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: SIMPLE_OBJECT_id = aDroid(psObj).id
    Case OBJ_STRUCTURE: SIMPLE_OBJECT_id = aStruct(psObj).id
    Case OBJ_FEATURE: SIMPLE_OBJECT_id = aFeature(psObj).id
    Case OBJ_PROJECTILE: SIMPLE_OBJECT_id = aProj(psObj).id
    End Select
End Property

Public Property Get SIMPLE_OBJECT_pos(ByVal psObj As Long) As Vector3i
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: SIMPLE_OBJECT_pos = aDroid(psObj).pos
    Case OBJ_STRUCTURE: SIMPLE_OBJECT_pos = aStruct(psObj).pos
    Case OBJ_FEATURE: SIMPLE_OBJECT_pos = aFeature(psObj).pos
    Case OBJ_PROJECTILE: SIMPLE_OBJECT_pos = aProj(psObj).pos
    End Select
End Property

Public Property Get SIMPLE_OBJECT_rot(ByVal psObj As Long) As Rotation
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: SIMPLE_OBJECT_rot = aDroid(psObj).rot
    Case OBJ_STRUCTURE: SIMPLE_OBJECT_rot = aStruct(psObj).rot
    Case OBJ_FEATURE: SIMPLE_OBJECT_rot = aFeature(psObj).rot
    Case OBJ_PROJECTILE: SIMPLE_OBJECT_rot = aProj(psObj).rot
    End Select
End Property

Public Property Get SIMPLE_OBJECT_player(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: SIMPLE_OBJECT_player = aDroid(psObj).player
    Case OBJ_STRUCTURE: SIMPLE_OBJECT_player = aStruct(psObj).player
    Case OBJ_FEATURE: SIMPLE_OBJECT_player = aFeature(psObj).player
    Case OBJ_PROJECTILE: SIMPLE_OBJECT_player = aProj(psObj).player
    End Select
End Property

Public Property Let SIMPLE_OBJECT_player(ByVal psObj As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).player = v And &HFF&
    Case OBJ_STRUCTURE: aStruct(psObj).player = v And &HFF&
    Case OBJ_FEATURE: aFeature(psObj).player = v And &HFF&
    Case OBJ_PROJECTILE: aProj(psObj).player = v And &HFF&
    End Select
End Property

Public Property Get SIMPLE_OBJECT_born(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: SIMPLE_OBJECT_born = aDroid(psObj).born
    Case OBJ_STRUCTURE: SIMPLE_OBJECT_born = aStruct(psObj).born
    Case OBJ_FEATURE: SIMPLE_OBJECT_born = aFeature(psObj).born
    Case OBJ_PROJECTILE: SIMPLE_OBJECT_born = aProj(psObj).born
    End Select
End Property

Public Property Get SIMPLE_OBJECT_died(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: SIMPLE_OBJECT_died = aDroid(psObj).died
    Case OBJ_STRUCTURE: SIMPLE_OBJECT_died = aStruct(psObj).died
    Case OBJ_FEATURE: SIMPLE_OBJECT_died = aFeature(psObj).died
    Case OBJ_PROJECTILE: SIMPLE_OBJECT_died = aProj(psObj).died
    End Select
End Property

Public Property Let SIMPLE_OBJECT_died(ByVal psObj As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).died = v
    Case OBJ_STRUCTURE: aStruct(psObj).died = v
    Case OBJ_FEATURE: aFeature(psObj).died = v
    Case OBJ_PROJECTILE: aProj(psObj).died = v
    End Select
End Property

Public Property Get SIMPLE_OBJECT_time(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: SIMPLE_OBJECT_time = aDroid(psObj).time_
    Case OBJ_STRUCTURE: SIMPLE_OBJECT_time = aStruct(psObj).time_
    Case OBJ_FEATURE: SIMPLE_OBJECT_time = aFeature(psObj).time_
    Case OBJ_PROJECTILE: SIMPLE_OBJECT_time = aProj(psObj).time_
    End Select
End Property

' ---- BASE_OBJECT's fields through a BASE_OBJECT *

Public Property Get BASE_OBJECT_visible(ByVal psObj As Long, ByVal player As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_visible = aDroid(psObj).visible(player)
    Case OBJ_STRUCTURE: BASE_OBJECT_visible = aStruct(psObj).visible(player)
    Case OBJ_FEATURE: BASE_OBJECT_visible = aFeature(psObj).visible(player)
    End Select
End Property

Public Property Let BASE_OBJECT_visible(ByVal psObj As Long, ByVal player As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).visible(player) = v And &HFF&
    Case OBJ_STRUCTURE: aStruct(psObj).visible(player) = v And &HFF&
    Case OBJ_FEATURE: aFeature(psObj).visible(player) = v And &HFF&
    End Select
End Property

' std::bitset<OBJECT_FLAG_COUNT> flags, as bits (BIT_(OBJECT_FLAG_...))
Public Property Get BASE_OBJECT_flags(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_flags = aDroid(psObj).flags
    Case OBJ_STRUCTURE: BASE_OBJECT_flags = aStruct(psObj).flags
    Case OBJ_FEATURE: BASE_OBJECT_flags = aFeature(psObj).flags
    End Select
End Property

Public Property Let BASE_OBJECT_flags(ByVal psObj As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).flags = v
    Case OBJ_STRUCTURE: aStruct(psObj).flags = v
    Case OBJ_FEATURE: aFeature(psObj).flags = v
    End Select
End Property

' sDisplay.imd (an iIMDBaseShape * model handle)
Public Property Get BASE_OBJECT_sDisplay_imd(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_sDisplay_imd = aDroid(psObj).sDisplay.imd
    Case OBJ_STRUCTURE: BASE_OBJECT_sDisplay_imd = aStruct(psObj).sDisplay.imd
    Case OBJ_FEATURE: BASE_OBJECT_sDisplay_imd = aFeature(psObj).sDisplay.imd
    End Select
End Property

Public Property Get BASE_OBJECT_timeLastHit(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_timeLastHit = aDroid(psObj).timeLastHit
    Case OBJ_STRUCTURE: BASE_OBJECT_timeLastHit = aStruct(psObj).timeLastHit
    Case OBJ_FEATURE: BASE_OBJECT_timeLastHit = aFeature(psObj).timeLastHit
    End Select
End Property

Public Property Let BASE_OBJECT_timeLastHit(ByVal psObj As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).timeLastHit = v
    Case OBJ_STRUCTURE: aStruct(psObj).timeLastHit = v
    Case OBJ_FEATURE: aFeature(psObj).timeLastHit = v
    End Select
End Property

Public Property Get BASE_OBJECT_lastHitWeapon(ByVal psObj As Long) As WEAPON_SUBCLASS
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_lastHitWeapon = aDroid(psObj).lastHitWeapon
    Case OBJ_STRUCTURE: BASE_OBJECT_lastHitWeapon = aStruct(psObj).lastHitWeapon
    Case OBJ_FEATURE: BASE_OBJECT_lastHitWeapon = aFeature(psObj).lastHitWeapon
    End Select
End Property

Public Property Let BASE_OBJECT_lastHitWeapon(ByVal psObj As Long, ByVal v As WEAPON_SUBCLASS)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).lastHitWeapon = v
    Case OBJ_STRUCTURE: aStruct(psObj).lastHitWeapon = v
    Case OBJ_FEATURE: aFeature(psObj).lastHitWeapon = v
    End Select
End Property

Public Property Get BASE_OBJECT_periodicalDamage(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_periodicalDamage = aDroid(psObj).periodicalDamage
    Case OBJ_STRUCTURE: BASE_OBJECT_periodicalDamage = aStruct(psObj).periodicalDamage
    Case OBJ_FEATURE: BASE_OBJECT_periodicalDamage = aFeature(psObj).periodicalDamage
    End Select
End Property

Public Property Let BASE_OBJECT_periodicalDamage(ByVal psObj As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).periodicalDamage = v
    Case OBJ_STRUCTURE: aStruct(psObj).periodicalDamage = v
    Case OBJ_FEATURE: aFeature(psObj).periodicalDamage = v
    End Select
End Property

Public Property Let BASE_OBJECT_body(ByVal psObj As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).body = v
    Case OBJ_STRUCTURE: aStruct(psObj).body = v
    Case OBJ_FEATURE: aFeature(psObj).body = v
    End Select
End Property

Public Property Get BASE_OBJECT_body(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_body = aDroid(psObj).body
    Case OBJ_STRUCTURE: BASE_OBJECT_body = aStruct(psObj).body
    Case OBJ_FEATURE: BASE_OBJECT_body = aFeature(psObj).body
    End Select
End Property

Public Property Get BASE_OBJECT_hasExtraFunction(ByVal psObj As Long) As Boolean
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_hasExtraFunction = aDroid(psObj).hasExtraFunction
    Case OBJ_STRUCTURE: BASE_OBJECT_hasExtraFunction = aStruct(psObj).hasExtraFunction
    Case OBJ_FEATURE: BASE_OBJECT_hasExtraFunction = aFeature(psObj).hasExtraFunction
    End Select
End Property

Public Property Let BASE_OBJECT_hasExtraFunction(ByVal psObj As Long, ByVal v As Boolean)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).hasExtraFunction = v
    Case OBJ_STRUCTURE: aStruct(psObj).hasExtraFunction = v
    Case OBJ_FEATURE: aFeature(psObj).hasExtraFunction = v
    End Select
End Property

Public Property Get BASE_OBJECT_timeAnimationStarted(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_timeAnimationStarted = aDroid(psObj).timeAnimationStarted
    Case OBJ_STRUCTURE: BASE_OBJECT_timeAnimationStarted = aStruct(psObj).timeAnimationStarted
    Case OBJ_FEATURE: BASE_OBJECT_timeAnimationStarted = aFeature(psObj).timeAnimationStarted
    End Select
End Property

Public Property Let BASE_OBJECT_timeAnimationStarted(ByVal psObj As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).timeAnimationStarted = v
    Case OBJ_STRUCTURE: aStruct(psObj).timeAnimationStarted = v
    Case OBJ_FEATURE: aFeature(psObj).timeAnimationStarted = v
    End Select
End Property

Public Property Get BASE_OBJECT_animationEvent(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_animationEvent = aDroid(psObj).animationEvent
    Case OBJ_STRUCTURE: BASE_OBJECT_animationEvent = aStruct(psObj).animationEvent
    Case OBJ_FEATURE: BASE_OBJECT_animationEvent = aFeature(psObj).animationEvent
    End Select
End Property

Public Property Let BASE_OBJECT_animationEvent(ByVal psObj As Long, ByVal v As Long)
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: aDroid(psObj).animationEvent = v And &HFF&
    Case OBJ_STRUCTURE: aStruct(psObj).animationEvent = v And &HFF&
    Case OBJ_FEATURE: aFeature(psObj).animationEvent = v And &HFF&
    End Select
End Property

' ---- basedef.h's inline functions

' Spacetime getSpacetime(SIMPLE_OBJECT const *psObj)
Public Function getSpacetime(ByVal psObj As Long) As Spacetime
    getSpacetime.pos = SIMPLE_OBJECT_pos(psObj)
    getSpacetime.rot = SIMPLE_OBJECT_rot(psObj)
    getSpacetime.time_ = SIMPLE_OBJECT_time(psObj)
End Function

Public Function isDead(ByVal psObj As Long) As Boolean
    ' See objmem.c for comments on the NOT_CURRENT_LIST hack
    isDead = ULT(NOT_CURRENT_LIST, SIMPLE_OBJECT_died(psObj))
End Function

' int objPosDiffSq(Position pos1, Position pos2)
Public Function objPosDiffSq(pos1 As Vector3i, pos2 As Vector3i) As Long
    Dim diff As Vector2i
    diff.x = pos1.x - pos2.x
    diff.y = pos1.y - pos2.y
    objPosDiffSq = dot2i(diff, diff)
End Function

' int objPosDiffSq(SIMPLE_OBJECT const *pos1, SIMPLE_OBJECT const *pos2)
Public Function objPosDiffSq_obj(ByVal pos1 As Long, ByVal pos2 As Long) As Long
    Dim p1 As Vector3i, p2 As Vector3i
    p1 = SIMPLE_OBJECT_pos(pos1)
    p2 = SIMPLE_OBJECT_pos(pos2)
    objPosDiffSq_obj = objPosDiffSq(p1, p2)
End Function

' True iff object is a droid, structure or feature (not a projectile). Will incorrectly return true if passed a nonsense object of type OBJ_NUM_TYPES.
Public Function isBaseObject(ByVal psObject As Long) As Boolean
    If psObject = 0 Then Exit Function
    isBaseObject = SIMPLE_OBJECT_type(psObject) <> OBJ_PROJECTILE
End Function

' Returns BASE_OBJECT * if base_object or NULL if not.
Public Function castBaseObject(ByVal psObject As Long) As Long
    If isBaseObject(psObject) Then castBaseObject = psObject
End Function
