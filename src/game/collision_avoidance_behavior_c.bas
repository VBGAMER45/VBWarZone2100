Attribute VB_Name = "collision_avoidance_behavior_c"
' src/steering/collision_avoidance_behavior.cpp, .h: Collision avoidance
' steering behavior.
' This behavior detects nearby obstacles and calculates an avoidance force
' to prevent collisions. Uses integer-only arithmetic.
' (VB) The class's members are CollisionAvoidanceBehavior_<member>.
Option Explicit

' Radius to scan for obstacles.
Public Const OBSTACLE_SCAN_RADIUS As Long = TILE_UNITS * 2

Private gridList() As Long              ' static GridList gridList: static to avoid allocations

' Calculate collision avoidance force for the given context.
Public Function CollisionAvoidanceBehavior_calculate(ctx As SteeringContext) As SteeringForce
    Dim numObstacles As Long, distTotal As Long, totalDir As Vector2i
    Dim toTarget As Vector2i, n As Long, gi As Long, obj As Long, obstacle As Long
    Dim obstaclePropStats As Long, obstacleRadius As Long, combinedRadius As Long
    Dim obstacleVel As Vector2i, obstacleSpeedGuess As Long, obstDirectionGuess As Long
    Dim obstaclePos As Vector2i, diff As Vector2i, distToObstacle As Long, predictionFactor As Long
    Dim deltaDiff As Vector2i, centreDist As Long, effectiveDist As Long
    Dim perp As Vector2i, avoid As Vector2i, destNorm As Vector2i, avoidNorm As Vector2i
    Dim ratio As Long, result As Vector2i, weight As Long, h As Long

    numObstacles = 0
    distTotal = 0

    ' Can't avoid anything if we can't move
    If ctx.maxSpeed = 0 Then
        Exit Function                       ' SteeringForce(Vector2i(0, 0), 0)
    End If

    ' Vector to target (for blending)
    toTarget.x = ctx.targetPos.x - ctx.currentPos.x
    toTarget.y = ctx.targetPos.y - ctx.currentPos.y

    ' Scan nearby objects for obstacles
    n = gridStartIterate(ctx.currentPos.x, ctx.currentPos.y, OBSTACLE_SCAN_RADIUS, gridList)
    For gi = 0 To n - 1
        obj = gridList(gi)

        ' Skip invalid obstacles
        If Not isValidObstacle(obj, ctx.droid) Then
            GoTo continue_
        End If

        obstacle = castDroid(obj)
        If obstacle = 0 Then
            GoTo continue_
        End If

        ' Get obstacle properties
        obstaclePropStats = DROID_getPropulsionStats(obstacle)
        obstacleRadius = moveObjRadius(obstacle)
        combinedRadius = ctx.radius + obstacleRadius

        ' Estimate obstacle velocity
        obstacleVel = estimateObstacleVelocity(obstacle)
        ' Find the guessed obstacle speed and direction, clamped to half our speed.
        obstacleSpeedGuess = MIN_(iHypot(obstacleVel.x, obstacleVel.y), ctx.maxSpeed \ 2)
        obstDirectionGuess = iAtan2(obstacleVel.x, obstacleVel.y)

        ' Position of obstacle relative to us
        obstaclePos.x = aDroid(obstacle).pos.x
        obstaclePos.y = aDroid(obstacle).pos.y
        diff.x = obstaclePos.x - ctx.currentPos.x
        diff.y = obstaclePos.y - ctx.currentPos.y

        ' Predict where the obstacle will be when we get close
        distToObstacle = iHypot(diff.x, diff.y)
        ' Find very approximate position of obstacle relative to us when we get close, based on our guesses.
        predictionFactor = MAX_(distToObstacle - combinedRadius * 2 \ 3, 0) * obstacleSpeedGuess \ ctx.maxSpeed
        deltaDiff.x = iSinR(obstDirectionGuess, predictionFactor)
        deltaDiff.y = iCosR(obstDirectionGuess, predictionFactor)

        ' Don't assume obstacle can go through blocking tiles
        If Not fpathBlockingTile(map_coord(aDroid(obstacle).pos.x + deltaDiff.x), _
                                 map_coord(aDroid(obstacle).pos.y + deltaDiff.y), _
                                 asPropulsionStats(obstaclePropStats).propulsionType) Then
            diff.x = diff.x + deltaDiff.x
            diff.y = diff.y + deltaDiff.y
        End If

        ' Check if obstacle is ahead of us (in direction of movement)
        If dot2i(diff, toTarget) < 0 Then
            ' Object is behind us relative to target
            GoTo continue_
        End If

        ' Calculate distance metrics
        centreDist = MAX_(iHypot(diff.x, diff.y), 1)
        effectiveDist = MAX_(centreDist - combinedRadius, 1)

        ' Accumulate avoidance direction
        ' The formula: diff * PRECISION / (centreDist * effectiveDist)
        ' This gives more weight to closer obstacles
        totalDir.x = totalDir.x + diff.x * PRECISION \ (centreDist * effectiveDist)
        totalDir.y = totalDir.y + diff.y * PRECISION \ (centreDist * effectiveDist)
        distTotal = distTotal + PRECISION \ effectiveDist
        numObstacles = numObstacles + 1
continue_:
    Next

    ' No obstacles to avoid
    If (totalDir.x = 0 And totalDir.y = 0) Or numObstacles = 0 Then
        Exit Function                       ' SteeringForce(Vector2i(0, 0), 0)
    End If

    ' Average the accumulated direction
    totalDir.x = totalDir.x \ numObstacles
    totalDir.y = totalDir.y \ numObstacles
    distTotal = distTotal \ numObstacles

    ' Create perpendicular avoid vector (choose side toward target)
    perp = orthogonalCW(totalDir)
    If dot2i(toTarget, perp) < 0 Then
        avoid.x = -perp.x
        avoid.y = -perp.y
    Else
        avoid = perp
    End If

    ' Normalize both destination and avoid vectors
    h = iHypot(toTarget.x, toTarget.y) + 1
    destNorm.x = toTarget.x * 32767 \ h
    destNorm.y = toTarget.y * 32767 \ h
    ' avoid.x and avoid.y are up to 65536, so we can multiply by at most 32767 here without potential overflow.
    h = iHypot(avoid.x, avoid.y) + 1
    avoidNorm.x = avoid.x * 32767 \ h
    avoidNorm.y = avoid.y * 32767 \ h

    ' Calculate blend ratio based on proximity
    ratio = MIN_(distTotal * ctx.radius \ 2, PRECISION)

    ' Blend destination and avoid vectors
    result.x = destNorm.x * (PRECISION - ratio) + avoidNorm.x * ratio
    result.y = destNorm.y * (PRECISION - ratio) + avoidNorm.y * ratio

    weight = PRECISION ' currently 1.0 magnitude to keep existing behavior
    ' Maybe calculate weight based on threat level?
    ' // Maximum avoidance weight (cap to prevent over-reaction)
    ' static constexpr int32_t MAX_AVOIDANCE_WEIGHT = 49152; // 0.75 in fixed-point
    ' int32_t weight = std::min(ratio, MAX_AVOIDANCE_WEIGHT);

    CollisionAvoidanceBehavior_calculate.force = result
    CollisionAvoidanceBehavior_calculate.weight = weight
End Function

' Disabled for transporters (they have their own movement logic).
Public Function CollisionAvoidanceBehavior_isEnabled(ctx As SteeringContext) As Boolean
    If ctx.droid = 0 Then
        CollisionAvoidanceBehavior_isEnabled = False
        Exit Function
    End If
    ' Transporters have their own movement logic
    CollisionAvoidanceBehavior_isEnabled = Not DROID_isTransporter(ctx.droid)
End Function

' Estimate obstacle's velocity, average of its actual and its intended one.
Private Function estimateObstacleVelocity(ByVal obstacle As Long) As Vector2i
    Dim velocityGuess1 As Vector2i, velocityGuess2 As Vector2i, obstaclePos As Vector2i, targetDiff As Vector2i
    Dim targetDist As Long, propStats As Long, maxSpeed As Long, intendedSpeed As Long, dir_ As Long

    ' Velocity guess 1: Guess the velocity the droid is actually moving at.
    velocityGuess1.x = iSinR(aDroid(obstacle).sMove.moveDir, aDroid(obstacle).sMove.speed)
    velocityGuess1.y = iCosR(aDroid(obstacle).sMove.moveDir, aDroid(obstacle).sMove.speed)

    ' Velocity guess 2: Guess the velocity the droid wants to move at.
    obstaclePos.x = aDroid(obstacle).pos.x
    obstaclePos.y = aDroid(obstacle).pos.y
    targetDiff.x = aDroid(obstacle).sMove.target.x - obstaclePos.x
    targetDiff.y = aDroid(obstacle).sMove.target.y - obstaclePos.y
    targetDist = iHypot(targetDiff.x, targetDiff.y)

    propStats = DROID_getPropulsionStats(obstacle)
    maxSpeed = asPropulsionStats(propStats).maxSpeed

    ' Scale intended speed by distance (slower when close to target)
    intendedSpeed = maxSpeed * MIN_(targetDist, OBSTACLE_SCAN_RADIUS) \ OBSTACLE_SCAN_RADIUS
    dir_ = iAtan2(targetDiff.x, targetDiff.y)
    velocityGuess2.x = iSinR(dir_, intendedSpeed)
    velocityGuess2.y = iCosR(dir_, intendedSpeed)

    ' If blocked, assume no intended movement
    If moveBlocked(obstacle) Then
        velocityGuess2.x = 0
        velocityGuess2.y = 0
    End If

    ' Average the two guesses
    estimateObstacleVelocity.x = (velocityGuess1.x + velocityGuess2.x) \ 2
    estimateObstacleVelocity.y = (velocityGuess1.y + velocityGuess2.y) \ 2
End Function

' Check if an object should be considered as an obstacle.
Private Function isValidObstacle(ByVal obj As Long, ByVal ourDroid As Long) As Boolean
    Dim obstacle As Long

    ' Skip ourselves
    If obj = ourDroid Then
        isValidObstacle = False
        Exit Function
    End If

    ' Only consider droids
    obstacle = castDroid(obj)
    If obstacle = 0 Then
        isValidObstacle = False
        Exit Function
    End If

    ' VTOL droids only avoid each other and don't affect ground droids
    If DROID_isVtol(ourDroid) <> DROID_isVtol(obstacle) Then
        isValidObstacle = False
        Exit Function
    End If

    ' Don't avoid transporters
    If DROID_isTransporter(obstacle) Then
        isValidObstacle = False
        Exit Function
    End If

    ' Don't avoid people on the other side - run over them
    If aDroid(obstacle).droidType = DROID_PERSON And aDroid(obstacle).player <> aDroid(ourDroid).player Then
        isValidObstacle = False
        Exit Function
    End If

    isValidObstacle = True
End Function
