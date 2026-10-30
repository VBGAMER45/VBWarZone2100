Attribute VB_Name = "steering_c"
' src/steering/steering.cpp, steering.h, utils.h: Core steering behaviors
' interface and types; the steering manager.
' Provides a modular and extensible steering behavior system for unit
' movement. All calculations use integer-only arithmetic for deterministic
' multiplayer synchronization.
' (VB) An ISteeringBehavior is a number (STEERING_BEHAVIOR_...) dispatched
' by a Select Case. The game has one SteeringManager (move.cpp's
' moveSteeringManager): this module's behaviour list and force cache. The
' int64_t sums are Doubles, exact below 2^53.
Option Explicit

' Fixed-point precision constant (65536 = 1.0)
' All steering calculations use this as the base unit for fractional values.
Public Const PRECISION As Long = 65536

' Steering force vector with weight for blending.
' The force vector uses fixed-point precision where 65536 represents
' a magnitude of 1.0.
' The weight is used when combining multiple steering forces together.
Public Type SteeringForce
    force As Vector2i                   ' Force vector in fixed-point precision
    weight As Long                      ' Weight for blending (in PRECISION units, i.e. weight == 65536 acts as a 1.0 multiplier)
End Type

' Context provided to each steering behavior for calculations.
' Contains all the information a steering behavior might need
' to calculate its force contribution.
Public Type SteeringContext
    droid As Long                       ' const DROID *: The droid currently being steered
    currentPos As Vector2i              ' Current position in world coordinates
    targetPos As Vector2i               ' Target/destination position
    velocity As Vector2i                ' Current velocity vector
    moveDir As Long                     ' uint16_t: Current movement direction (0-65535 = 0-360 degrees)
    speed As Long                       ' Current speed
    maxSpeed As Long                    ' Max speed from droid's propulsion stats
    radius As Long                      ' Droid collision radius
End Type

' (VB) the ISteeringBehavior implementations
Public Enum STEERING_BEHAVIOR
    STEERING_BEHAVIOR_COLLISION_AVOIDANCE = 1   ' CollisionAvoidanceBehavior
End Enum

' Cached storage for calculated steering forces, used to reduce
' the number of allocations on the hot path
Private forces_() As SteeringForce
Private forces_size As Long
Private behaviors_() As Long            ' std::vector<std::unique_ptr<ISteeringBehavior>>
Private behaviors_size As Long

' Check if this force has any effect
Public Function SteeringForce_isZero(f As SteeringForce) As Boolean
    SteeringForce_isZero = f.force.x = 0 And f.force.y = 0 And f.weight = 0
End Function

' (VB) behavior->calculate(ctx)
Private Function behaviorCalculate(ByVal behavior As STEERING_BEHAVIOR, ctx As SteeringContext) As SteeringForce
    Select Case behavior
    Case STEERING_BEHAVIOR_COLLISION_AVOIDANCE
        behaviorCalculate = CollisionAvoidanceBehavior_calculate(ctx)
    End Select
End Function

' (VB) behavior->isEnabled(ctx)
Private Function behaviorIsEnabled(ByVal behavior As STEERING_BEHAVIOR, ctx As SteeringContext) As Boolean
    Select Case behavior
    Case STEERING_BEHAVIOR_COLLISION_AVOIDANCE
        behaviorIsEnabled = CollisionAvoidanceBehavior_isEnabled(ctx)
    Case Else
        behaviorIsEnabled = True
    End Select
End Function

' Add a behavior to the manager.
' Behaviors are evaluated in the order they are added.
Public Sub SteeringManager_addBehavior(ByVal behavior As STEERING_BEHAVIOR)
    ReDim Preserve behaviors_(0 To behaviors_size)
    behaviors_(behaviors_size) = behavior
    behaviors_size = behaviors_size + 1
    If forces_size = 0 Then ReDim forces_(0 To behaviors_size - 1) Else ReDim Preserve forces_(0 To behaviors_size - 1)
End Sub

' (VB) the manager has behaviours (move.cpp's static instance is made)
Public Function SteeringManager_ready() As Boolean
    SteeringManager_ready = behaviors_size > 0
End Function

' Calculate combined steering vector.
' Evaluates all enabled behaviors and combines their forces
' using weighted summation.
Public Function SteeringManager_calculateSteering(ctx As SteeringContext) As Vector2i
    Dim i As Long, force As SteeringForce
    forces_size = 0

    For i = 0 To behaviors_size - 1
        If behaviorIsEnabled(behaviors_(i), ctx) Then
            force = behaviorCalculate(behaviors_(i), ctx)
            If Not SteeringForce_isZero(force) Then
                forces_(forces_size) = force
                forces_size = forces_size + 1
            End If
        End If
    Next

    SteeringManager_calculateSteering = combineForcesImpl()
End Function

' Calculate combined steering as a direction angle.
' Same as `calculateSteering()` but returns the result as an
' angle suitable for use with the current movement system.
Public Function SteeringManager_calculateSteeringDirection(ctx As SteeringContext) As Long
    Dim steering As Vector2i, toTarget As Vector2i
    steering = SteeringManager_calculateSteering(ctx)

    ' If no steering force, use direction to target
    If steering.x = 0 And steering.y = 0 Then
        toTarget.x = ctx.targetPos.x - ctx.currentPos.x
        toTarget.y = ctx.targetPos.y - ctx.currentPos.y
        SteeringManager_calculateSteeringDirection = iAtan2(toTarget.x, toTarget.y)
        Exit Function
    End If

    SteeringManager_calculateSteeringDirection = iAtan2(steering.x, steering.y)
End Function

' Combine multiple forces using weighted average.
' (VB) the forces are forces_(0 .. forces_size - 1)
Private Function combineForcesImpl() As Vector2i
    Dim totalX As Double, totalY As Double, totalWeight As Double
    Dim weightedX As Double, weightedY As Double, i As Long

    If forces_size = 0 Then
        Exit Function                       ' Vector2i(0, 0)
    End If

    ' Calculate weighted sum of forces
    For i = 0 To forces_size - 1
        ' Weight the force by its weight value
        ' Force is already in fixed-point, weight is also fixed-point
        ' Result: force * weight / PRECISION
        weightedX = Fix(CDbl(forces_(i).force.x) * forces_(i).weight / PRECISION)
        weightedY = Fix(CDbl(forces_(i).force.y) * forces_(i).weight / PRECISION)

        totalX = totalX + weightedX
        totalY = totalY + weightedY
        totalWeight = totalWeight + forces_(i).weight
    Next

    If totalWeight = 0 Then
        Exit Function                       ' Vector2i(0, 0)
    End If

    ' Normalize by total weight to get average direction
    ' Scale back to PRECISION units
    combineForcesImpl.x = FromU32(Fix(totalX * PRECISION / totalWeight))
    combineForcesImpl.y = FromU32(Fix(totalY * PRECISION / totalWeight))
End Function

' utils.h: Calculate orthogonal vector (90 degrees rotation).
' Returns the vector rotated 90 degrees clockwise.
Public Function orthogonalCW(v As Vector2i) As Vector2i
    orthogonalCW.x = v.y
    orthogonalCW.y = -v.x
End Function
