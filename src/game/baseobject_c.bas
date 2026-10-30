Attribute VB_Name = "baseobject_c"
' src/baseobject.cpp, baseobject.h: functions for the base object type.
' The constructors and destructors of SIMPLE_OBJECT and BASE_OBJECT are
' part of each type's (DROID_construct, STRUCTURE_construct,
' FEATURE_construct and their _destruct). checkObject's and
' _syncDebugObject's projectile cases call projectile_c's checkProjectile
' and syncDebugProjectile_.
' CHECK_OBJECT(object) is checkObject object, "<file>:<line>", "<function>",
' max_check_object_recursion; syncDebugObject(psObject, ch) is
' syncDebugObject_ "<function>", psObject, ch.
Option Explicit

Public Type StructureBounds
    map As Vector2i                     ' Map coordinates of upper left corner of structure.
    size As Vector2i                    ' Size (in map coordinates) of the structure.
End Type

Public Const max_check_object_recursion As Long = 4

' StructureBounds(Vector2i const &map, Vector2i const &size)
Public Function StructureBounds_(map As Vector2i, size As Vector2i) As StructureBounds
    StructureBounds_.map = map
    StructureBounds_.size = size
End Function

Public Function StructureBounds_valid(b As StructureBounds) As Boolean
    StructureBounds_valid = b.size.x >= 0
End Function

Private Function interpolateAngle(ByVal v1 As Long, ByVal v2 As Long, ByVal t1 As Long, ByVal t2 As Long, ByVal t As Long) As Long
    Dim numer As Long, denom As Long
    numer = t - t1
    denom = t2 - t1
    interpolateAngle = (v1 + angleDelta(v2 - v1) * numer \ denom) And &HFFFF&
End Function

Private Function interpolatePos(p1 As Vector3i, p2 As Vector3i, ByVal t1 As Long, ByVal t2 As Long, ByVal t As Long) As Vector3i
    Dim numer As Long, denom As Long
    numer = t - t1
    denom = t2 - t1
    interpolatePos.x = p1.x + (p2.x - p1.x) * numer \ denom
    interpolatePos.y = p1.y + (p2.y - p1.y) * numer \ denom
    interpolatePos.z = p1.z + (p2.z - p1.z) * numer \ denom
End Function

Public Function interpolateRot(v1 As Rotation, v2 As Rotation, ByVal t1 As Long, ByVal t2 As Long, ByVal t As Long) As Rotation
    If t1 = t2 Then
        ' spacetime overlap
        ' just return v2 (the current rotation)
        interpolateRot = v2
        Exit Function
    End If
    'return v1 + (v2 - v1) * (t - t1) / (t2 - t1);
    interpolateRot = Rotation_(interpolateAngle(v1.direction, v2.direction, t1, t2, t), _
                               interpolateAngle(v1.pitch, v2.pitch, t1, t2, t), _
                               interpolateAngle(v1.roll, v2.roll, t1, t2, t))
End Function

Private Function interpolateSpacetime(st1 As Spacetime, st2 As Spacetime, ByVal t As Long) As Spacetime
    ' Cyp says this should never happen, #3037 and #3238 say it does though.
    If Not (st1.time_ <> st2.time_) Then ASSERT_FAILURE "st1.time != st2.time", "src/baseobject.cpp:67", "interpolateSpacetime", "Spacetime overlap!": interpolateSpacetime = st2: Exit Function
    interpolateSpacetime.pos = interpolatePos(st1.pos, st2.pos, st1.time_, st2.time_, t)
    interpolateSpacetime.rot = interpolateRot(st1.rot, st2.rot, st1.time_, st2.time_, t)
    interpolateSpacetime.time_ = t
End Function

Public Function interpolateObjectSpacetime(ByVal obj As Long, ByVal t As Long) As Spacetime
    Dim st As Spacetime
    Select Case SIMPLE_OBJECT_type(obj)
    Case OBJ_DROID
        st = getSpacetime(obj)
        interpolateObjectSpacetime = interpolateSpacetime(aDroid(obj).prevSpacetime, st, t)
    Case OBJ_PROJECTILE
        st = getSpacetime(obj)
        interpolateObjectSpacetime = interpolateSpacetime(aProj(castProjectile(obj)).prevSpacetime, st, t)
    Case Else
        interpolateObjectSpacetime = getSpacetime(obj)
    End Select
End Function

' pType: a BASE_STATS * handle
Public Function getStatsSize(ByVal pType As Long, ByVal direction As Long) As Vector2i
    If StatIsStructure(pType) Then
        getStatsSize = STRUCTURE_STATS_size(pType And &HFFFF&, direction)
    ElseIf StatIsFeature(pType) Then
        getStatsSize = FEATURE_STATS_size(pType And &HFFFF&)
    Else
        getStatsSize = Vector2i_(1, 1)
    End If
End Function

' StructureBounds getStructureBounds(BASE_OBJECT const *object)
Public Function getStructureBounds(ByVal object As Long) As StructureBounds
    Dim psStructure As Long, psFeature As Long
    psStructure = castStructure(object)
    psFeature = castFeature(object)

    If psStructure <> 0 Then
        getStructureBounds = getStructureBounds_struct(psStructure)
        Exit Function
    ElseIf psFeature <> 0 Then
        getStructureBounds = getStructureBounds_feature(psFeature)
        Exit Function
    End If

    getStructureBounds = StructureBounds_(Vector2i_(32767, 32767), Vector2i_(-65535, -65535))  ' Default to an invalid area.
End Function

' StructureBounds getStructureBounds(BASE_STATS const *stats, Vector2i pos, uint16_t direction)
Public Function getStructureBounds_stats(ByVal stats As Long, pos As Vector2i, ByVal direction As Long) As StructureBounds
    If StatIsStructure(stats) Then
        getStructureBounds_stats = getStructureBounds_structStats(stats And &HFFFF&, pos, direction)
    ElseIf StatIsFeature(stats) Then
        getStructureBounds_stats = getStructureBounds_featureStats(stats And &HFFFF&, pos)
    Else
        getStructureBounds_stats = StructureBounds_(Vector2i_(map_coord(pos.x), map_coord(pos.y)), Vector2i_(1, 1))  ' Default to a 1x1 tile.
    End If
End Function

' structure.cpp's StructureBounds getStructureBounds(const STRUCTURE *object)
Public Function getStructureBounds_struct(ByVal object As Long) As StructureBounds
    Dim size As Vector2i
    size = STRUCTURE_size(object)
    getStructureBounds_struct.map.x = map_coord(aStruct(object).pos.x) - size.x \ 2
    getStructureBounds_struct.map.y = map_coord(aStruct(object).pos.y) - size.y \ 2
    getStructureBounds_struct.size = size
End Function

' structure.cpp's StructureBounds getStructureBounds(const STRUCTURE_STATS *stats, Vector2i pos, uint16_t direction)
Public Function getStructureBounds_structStats(ByVal stats As Long, pos As Vector2i, ByVal direction As Long) As StructureBounds
    Dim size As Vector2i
    size = STRUCTURE_STATS_size(stats, direction)
    getStructureBounds_structStats.map.x = map_coord(pos.x) - size.x \ 2
    getStructureBounds_structStats.map.y = map_coord(pos.y) - size.y \ 2
    getStructureBounds_structStats.size = size
End Function

' feature.cpp's StructureBounds getStructureBounds(FEATURE const *object)
Public Function getStructureBounds_feature(ByVal object As Long) As StructureBounds
    getStructureBounds_feature = getStructureBounds_featureStats(aFeature(object).psStats, Vector2i_(aFeature(object).pos.x, aFeature(object).pos.y))
End Function

' feature.cpp's StructureBounds getStructureBounds(FEATURE_STATS const *stats, Vector2i pos)
Public Function getStructureBounds_featureStats(ByVal stats As Long, pos As Vector2i) As StructureBounds
    Dim size As Vector2i
    size = FEATURE_STATS_size(stats)
    getStructureBounds_featureStats.map.x = map_coord(pos.x) - size.x \ 2
    getStructureBounds_featureStats.map.y = map_coord(pos.y) - size.y \ 2
    getStructureBounds_featureStats.size = size
End Function

' Query visibility for display purposes (i.e. for `selectedPlayer`)
' *DO NOT USE TO QUERY VISIBILITY FOR CALCULATIONS INVOLVING GAME / SIMULATION STATE*
Public Function BASE_OBJECT_visibleForLocalDisplay(ByVal this As Long) As Long
    If godMode Then
        ' is visible to selectedPlayer
        BASE_OBJECT_visibleForLocalDisplay = UBYTE_MAX
        Exit Function
    End If
    If Not ULT(selectedPlayer, MAX_PLAYERS) Then
        BASE_OBJECT_visibleForLocalDisplay = 0
        Exit Function
    End If
    BASE_OBJECT_visibleForLocalDisplay = BASE_OBJECT_visible(this, selectedPlayer)
End Function

' WEAPON_STATS* BASE_OBJECT::getWeaponStats(int weaponSlot) const
Public Function BASE_OBJECT_getWeaponStats(ByVal this As Long, ByVal weaponSlot As Long) As Long
    Select Case this \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_getWeaponStats = aDroid(this).asWeaps(weaponSlot).nStat
    Case OBJ_STRUCTURE: BASE_OBJECT_getWeaponStats = aStruct(this).asWeaps(weaponSlot).nStat
    Case OBJ_FEATURE: BASE_OBJECT_getWeaponStats = aFeature(this).asWeaps(weaponSlot).nStat
    End Select
End Function

' (VB) as checkDroid, its ASSERT_HELPER logs under the caller's location
Public Sub checkObject(ByVal psObject As Long, ByVal location_description As String, ByVal function_ As String, ByVal recurse As Long)
    If recurse < 0 Then
        Exit Sub
    End If

    If Not (psObject <> 0) Then ASSERT_FAILURE "psObject != nullptr", "src/baseobject.cpp:155", "checkObject", "NULL pointer"

    Select Case SIMPLE_OBJECT_type(psObject)
    Case OBJ_DROID
        checkDroid psObject, location_description, function_, recurse - 1

    Case OBJ_STRUCTURE
        checkStructure psObject, location_description, function_, recurse - 1

    Case OBJ_PROJECTILE
        checkProjectile psObject, location_description, function_, recurse - 1

    Case OBJ_FEATURE

    Case Else
        ASSERT_HELPER_FAILURE 175, "!""invalid object type""", location_description, function_, "CHECK_OBJECT: Invalid object type (type num %u)", SIMPLE_OBJECT_type(psObject)
    End Select
End Sub

Public Sub syncDebugObject_(ByVal function_ As String, ByVal psObject As Long, ByVal ch As Long)
    Select Case SIMPLE_OBJECT_type(psObject)
    Case OBJ_DROID: syncDebugDroid_ function_, psObject, ch
    Case OBJ_STRUCTURE: syncDebugStructure_ function_, psObject, ch
    Case OBJ_FEATURE: syncDebugFeature_ function_, psObject, ch
    Case OBJ_PROJECTILE: syncDebugProjectile_ function_, psObject, ch
    Case Else
        syncDebug function_, "%c unidentified_object%d = p%d;objectType%d", ch, SIMPLE_OBJECT_id(psObject), SIMPLE_OBJECT_player(psObject), SIMPLE_OBJECT_type(psObject)
        ASSERT_HELPER_FAILURE 189, "!""invalid object type""", "_syncDebugObject", function_, "syncDebug: Invalid object type (type num %u)", SIMPLE_OBJECT_type(psObject)
    End Select
End Sub

' Reset animation state so some units don't walk in place in the menu sometimes.
Public Sub resetObjectAnimationState(ByVal psObj As Long)
    If psObj = 0 Then
        Exit Sub
    End If
    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID, OBJ_STRUCTURE, OBJ_FEATURE
        BASE_OBJECT_timeAnimationStarted(psObj) = 0
        BASE_OBJECT_animationEvent(psObj) = ANIM_EVENT_NONE
    Case Else
        debug_ 252, LOG_WARNING, "resetObjectAnimationState", "Tried resetting animation state on unknown object type"
    End Select
End Sub
