Attribute VB_Name = "projectiledef_h"
' src/projectiledef.h: Definitions for projectiles. A PROJECTILE * is a
' handle into the pool aProj (paged_entity_container_h; OBJ_PROJECTILE's
' handles, see basedef_h), SIMPLE_OBJECT's fields first, as DROID's. Its
' psDamaged (std::vector<BASE_OBJECT *>) is psDamaged() plus
' psDamaged_size; its constructor and destructor are projectile_c's
' PROJECTILE_construct and PROJECTILE_destruct, update() is PROJECTILE_update.
Option Explicit

Public Enum PROJ_STATE
    PROJ_INFLIGHT
    PROJ_IMPACT
    PROJ_POSTIMPACT
    PROJ_INACTIVE
End Enum

Public Type PROJECTILE
    ' SIMPLE_OBJECT
    typ As OBJECT_TYPE                  ' The type of object
    id As Long                          ' uint32_t: ID number of the object
    pos As Vector3i                     ' Position of the object
    rot As Rotation                     ' Object's yaw +ve rotation around up-axis
    player As Long                      ' uint8_t: Which player the object belongs to
    born As Long                        ' uint32_t: Time the game object was born
    died As Long                        ' uint32_t: When an object was destroyed, if 0 still alive
    time_ As Long                       ' uint32_t: Game time of given space-time position.
    ' PROJECTILE
    state As Long                       ' UBYTE: current projectile state
    bVisible As Long                    ' UBYTE: whether the selected player should see the projectile
    psWStats As Long                    ' WEAPON_STATS *: firing weapon stats (its index, -1 = nullptr)
    psSource As Long                    ' BASE_OBJECT *: what fired the projectile
    psDest As Long                      ' BASE_OBJECT *: target of this projectile
    psDamaged() As Long                 ' std::vector<BASE_OBJECT *>: the targets that have already been dealt damage to (don't damage the same target twice)
    psDamaged_size As Long
    src As Vector3i                     ' Where projectile started
    dst As Vector3i                     ' The target coordinates
    vXY As Long                         ' SDWORD: axis velocities
    vZ As Long
    prevSpacetime As Spacetime          ' Location of projectile in previous tick.
    expectedDamageCaused As Long        ' UDWORD: Expected damage that this projectile will cause to the target.
    partVisible As Long                 ' how much of target was visible on shooting (important for homing)
End Type

' True iff object is a projectile.
Public Function isProjectile(ByVal psObject As Long) As Boolean
    If psObject <> 0 Then
        isProjectile = SIMPLE_OBJECT_type(psObject) = OBJ_PROJECTILE
    End If
End Function

' Returns PROJECTILE * if projectile or NULL if not.
Public Function castProjectile(ByVal psObject As Long) As Long
    If isProjectile(psObject) Then
        castProjectile = psObject
    Else
        castProjectile = 0
    End If
End Function
