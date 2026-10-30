Attribute VB_Name = "warcam_c"
' src/warcam.cpp, warcam.h: handles tracking/following of in game objects
' (the camera that tracks the selected droids) and the radar jump's
' animation. The globals' C++ initialisers are warcam_c_init's.
Option Explicit

' Alex McLean, Pumpkin Studios, EIDOS Interactive, 1998
'    23rd September, 1998 - This code is now so hideously complex
'    and unreadable that's it's inadvisable to attempt changing
'    how the camera works, since I'm not sure that I'll be able to even
'    get it working the way it used to, should anything get broken.
'    I really hope that no further changes are needed here...:-(
'    Alex M.

Public Const X_UPDATE As Long = &H1
Public Const Y_UPDATE As Long = &H2
Public Const Z_UPDATE As Long = &H4
Public Const CAM_X_AND_Y As Long = (X_UPDATE + Y_UPDATE)
Public Const CAM_ALL As Long = (X_UPDATE + Y_UPDATE + Z_UPDATE)

Public Const ACCEL_CONSTANT As Single = 12!
Public Const VELOCITY_CONSTANT As Single = 4!
Public Const ROT_ACCEL_CONSTANT As Single = 4!
Public Const ROT_VELOCITY_CONSTANT As Single = 2.5!

' The different tracking states
Public Enum WARSTATUS
    CAM_INACTIVE
    CAM_REQUEST
    CAM_TRACK_DROID
    CAM_RADAR_JUMP
    CAM_RESET
End Enum

' Storage for old viewnagles etc
Private Type WARCAM
    status As WARSTATUS
    trackClass As Long                  ' UDWORD
    lastUpdate As Long                  ' UDWORD
    oldView As iView

    acceleration As Vector3f            ' = Vector3f(0.f, 0.f, 0.f)
    velocity As Vector3f
    position As Vector3f

    rotation As Vector3f
    rotVel As Vector3f
    rotAccel As Vector3f

    oldDistance As Single
    target As Long                      ' DROID *
End Type

Private Const MIN_TRACK_HEIGHT As Long = 16

' Holds all the details of our camera
Private trackingCamera As WARCAM

' Present rotation for the 3d camera logo
Private warCamLogoRotation As Long

' Used to animate radar jump
Private Type RADAR_JUMP
    position As Animation_Vector3f
    rotation As Animation_Vector3f      ' RotationAnimation
End Type

Private radarJumpAnimation As RADAR_JUMP

Private presAvAngle As Long             ' = 0

'    These are the DEFAULT offsets that make us track _behind_ a droid and allow
'    it to be pretty far _down_ the screen, so we can see more

' Offset from droid's world coords
' How far we track relative to the droids location - direction matters
Private Const CAM_DEFAULT_OFFSET As Long = -400

' How much info do you want when tracking a droid - this toggles full stat info
Private bFullInfo As Boolean            ' = false

Private Const LEADER_LEFT As Long = 1
Private Const LEADER_RIGHT As Long = 2
Private Const LEADER_UP As Long = 3
Private Const LEADER_DOWN As Long = 4
Private Const LEADER_STATIC As Long = 5

Private Const GROUP_SELECTED As Long = &HFFFFFFFF

' (VB) the C++ initialisers: static RADAR_JUMP radarJumpAnimation{Animation<Vector3f>(graphicsTime, EASE_IN_OUT),
' RotationAnimation(graphicsTime, EASE_IN_OUT)}
Public Sub warcam_c_init()
    radarJumpAnimation.position = Animation_Vector3f_new(ANIMATION_TIME_GRAPHICS, EASE_IN_OUT)
    radarJumpAnimation.rotation = RotationAnimation_new(ANIMATION_TIME_GRAPHICS, EASE_IN_OUT)
    presAvAngle = 0
    bFullInfo = False
End Sub

'-----------------------------------------------------------------------------------
' Sets the camera to inactive to begin with
Public Sub initWarCam()
    ' We're not initially following anything
    trackingCamera.status = CAM_INACTIVE

    ' Logo setup
    warCamLogoRotation = 0
End Sub

Private Sub processLeaderSelection()
    Dim psNew As Long, leaderClass As Long, bSuccess As Boolean, dif As Double, bestSoFar As Double
    Dim it As Long, psDroid As Long, target As Long
    psNew = 0

    If Not getWarCamStatus() Then
        Exit Sub
    End If

    If keyPressed(KEY_LEFTARROW) Then
        leaderClass = LEADER_LEFT

    ElseIf keyPressed(KEY_RIGHTARROW) Then
        leaderClass = LEADER_RIGHT

    ElseIf keyPressed(KEY_UPARROW) Then
        leaderClass = LEADER_UP

    ElseIf keyPressed(KEY_DOWNARROW) Then
        leaderClass = LEADER_DOWN
    Else
        leaderClass = LEADER_STATIC
    End If

    bSuccess = False
    bestSoFar = 4294967295#             ' UDWORD_MAX

    If selectedPlayer >= MAX_PLAYERS Then
        ' Not currently supported when selectedPlayer >= MAX_PLAYERS
        Exit Sub
    End If

    target = trackingCamera.target
    If leaderClass <> LEADER_STATIC Then
        it = apsDroidLists(selectedPlayer).head
        Do While it <> 0
            psDroid = lnObj(it)
            ' Is it even on the sscreen?
            If DrawnInLastFrame(aDroid(psDroid).sDisplay.frameNumber) And aDroid(psDroid).selected And psDroid <> target Then
                Select Case leaderClass
                Case LEADER_LEFT
                    If aDroid(psDroid).sDisplay.screenX < aDroid(target).sDisplay.screenX Then
                        dif = U32(aDroid(target).sDisplay.screenX - aDroid(psDroid).sDisplay.screenX)
                        If dif < bestSoFar Then
                            bestSoFar = dif
                            bSuccess = True
                            psNew = psDroid
                        End If
                    End If
                Case LEADER_RIGHT
                    If aDroid(psDroid).sDisplay.screenX > aDroid(target).sDisplay.screenX Then
                        dif = U32(aDroid(psDroid).sDisplay.screenX - aDroid(target).sDisplay.screenX)
                        If dif < bestSoFar Then
                            bestSoFar = dif
                            bSuccess = True
                            psNew = psDroid
                        End If
                    End If
                Case LEADER_UP
                    If aDroid(psDroid).sDisplay.screenY < aDroid(target).sDisplay.screenY Then
                        dif = U32(aDroid(target).sDisplay.screenY - aDroid(psDroid).sDisplay.screenY)
                        If dif < bestSoFar Then
                            bestSoFar = dif
                            bSuccess = True
                            psNew = psDroid
                        End If
                    End If
                Case LEADER_DOWN
                    If aDroid(psDroid).sDisplay.screenY > aDroid(target).sDisplay.screenY Then
                        dif = U32(aDroid(psDroid).sDisplay.screenY - aDroid(target).sDisplay.screenY)
                        If dif < bestSoFar Then
                            bestSoFar = dif
                            bSuccess = True
                            psNew = psDroid
                        End If
                    End If
                End Select
            End If
            it = lnNext(it)
        Loop
    End If
    If bSuccess Then
        camAlignWithTarget psNew
    End If
End Sub

' Updates the camera position/angle along with the object movement
Public Sub processWarCam()
    Dim foundTarget As Long

    ' Ensure that the camera only ever flips state within this routine!
    Select Case trackingCamera.status
    Case CAM_INACTIVE

    Case CAM_REQUEST

        ' See if we can find the target to follow
        foundTarget = camFindDroidTarget()

        If foundTarget <> 0 Then
            If aDroid(foundTarget).died = 0 Then
                ' We've got one, so store away info
                camAlignWithTarget foundTarget
                ' We're now into tracking status
                trackingCamera.status = CAM_TRACK_DROID
                ' Inform via console
                If getWarCamStatus() Then
                    CONPRINTF "WZ/CAM  - %s", droidGetName(foundTarget)
                End If
            Else
                ' We've requested a track with no droid selected
                trackingCamera.status = CAM_INACTIVE
            End If
        Else
            ' We've requested a track with no droid selected
            trackingCamera.status = CAM_INACTIVE
        End If

    Case CAM_TRACK_DROID
        ' Track the droid unless routine comes back false
        If Not camTrackCamera() Then
            '
            '    Camera track came back false, either because droid died or is
            '    no longer selected, so reset to old values
            '
            foundTarget = camFindDroidTarget()
            trackingCamera.status = CAM_RESET
            If foundTarget <> 0 Then
                If aDroid(foundTarget).died = 0 Then
                    trackingCamera.status = CAM_REQUEST
                End If
            End If
        End If

        processLeaderSelection

    Case CAM_RADAR_JUMP
        camRadarJump

    Case CAM_RESET
        ' Switch to inactive mode
        trackingCamera.status = CAM_INACTIVE
    End Select
End Sub

'-----------------------------------------------------------------------------------

' Flips states for camera active
Public Sub setWarCamActive(ByVal status As Boolean)
    debug_ 344, LOG_NEVER, "setWarCamActive", "setWarCamActive(%d)" & vbLf, status

    ' We're trying to switch it on
    If status = True Then
        ' If it's not inactive then it's already in use - so return
        ' We're tracking a droid
        If trackingCamera.status <> CAM_INACTIVE Then
            Exit Sub
        Else
            ' Otherwise request the camera to track
            trackingCamera.status = CAM_REQUEST
        End If
    Else
        ' We trying to switch off
        ' Is it already off?
        If trackingCamera.status = CAM_INACTIVE Then
            Exit Sub
        Else
            ' Attempt to set to normal
            trackingCamera.status = CAM_RESET
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------

Public Function camFindDroidTarget() As Long
    Dim it As Long
    If selectedPlayer >= MAX_PLAYERS Then
        ' Not currently supported when selectedPlayer >= MAX_PLAYERS
        camFindDroidTarget = 0
        Exit Function
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        If aDroid(lnObj(it)).selected Then
            ' Return the first one found
            camFindDroidTarget = lnObj(it)
            Exit Function
        End If
        it = lnNext(it)
    Loop

    ' We didn't find one
    camFindDroidTarget = 0
End Function

'-----------------------------------------------------------------------------------

' Stores away old viewangle info and sets up new distance and angles
Public Sub camAlignWithTarget(ByVal psDroid As Long)
    ' Store away the target
    trackingCamera.target = psDroid

    ' Save away all the view angles
    trackingCamera.rotation.x = CSng(playerPos.r.x)
    trackingCamera.rotation.y = CSng(playerPos.r.y)
    trackingCamera.rotation.z = CSng(playerPos.r.z)
    trackingCamera.oldView.r.x = playerPos.r.x
    trackingCamera.oldView.r.y = playerPos.r.y
    trackingCamera.oldView.r.z = playerPos.r.z

    ' Store away the old positions and set the start position too
    trackingCamera.position.x = CSng(playerPos.p.x)
    trackingCamera.position.y = CSng(playerPos.p.y)
    trackingCamera.position.z = CSng(playerPos.p.z)
    trackingCamera.oldView.p.x = playerPos.p.x
    trackingCamera.oldView.p.y = playerPos.p.y
    trackingCamera.oldView.p.z = playerPos.p.z

    '    trackingCamera.rotation.x = player.r.x = DEG(-90);
    ' No initial velocity for moving
    trackingCamera.velocity.x = 0!: trackingCamera.velocity.y = 0!: trackingCamera.velocity.z = 0!
    ' Nor for rotation
    trackingCamera.rotVel.x = 0!: trackingCamera.rotVel.y = 0!: trackingCamera.rotVel.z = 0!
    ' No initial acceleration for moving
    trackingCamera.acceleration.x = 0!: trackingCamera.acceleration.y = 0!: trackingCamera.acceleration.z = 0!
    ' Nor for rotation
    trackingCamera.rotAccel.x = 0!: trackingCamera.rotAccel.y = 0!: trackingCamera.rotAccel.z = 0!

    ' Sote the old distance
    trackingCamera.oldDistance = getViewDistance()  'distance;

    ' Store away when we started
    trackingCamera.lastUpdate = realTime
End Sub

' (VB) groupNumber == GROUP_SELECTED ? psDroid->selected : psDroid->group == groupNumber
Private Function inTrackGroup(ByVal psDroid As Long, ByVal groupNumber As Long) As Boolean
    If groupNumber = GROUP_SELECTED Then
        inTrackGroup = aDroid(psDroid).selected
    Else
        inTrackGroup = (aDroid(psDroid).group = groupNumber)
    End If
End Function

'-----------------------------------------------------------------------------------
Private Function getAverageTrackAngle(ByVal groupNumber As Long, ByVal bCheckOnScreen As Boolean) As Long
    Dim xTotal As Long, yTotal As Long, it As Long, psDroid As Long, onScreen As Boolean
    xTotal = 0: yTotal = 0

    If selectedPlayer >= MAX_PLAYERS Then
        ' Not currently supported when selectedPlayer >= MAX_PLAYERS
        getAverageTrackAngle = 0
        Exit Function
    End If

    ' Got thru' all droids
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        ' Is he worth selecting?
        If inTrackGroup(psDroid, groupNumber) Then
            If bCheckOnScreen Then
                onScreen = objectOnScreen(psDroid, pie_GetVideoBufferWidth() \ 6)
            Else
                onScreen = True
            End If
            If onScreen Then
                xTotal = xTotal + iSin(aDroid(psDroid).rot.direction)
                yTotal = yTotal + iCos(aDroid(psDroid).rot.direction)
            End If
        End If
        it = lnNext(it)
    Loop
    presAvAngle = iAtan2(xTotal, yTotal)
    getAverageTrackAngle = presAvAngle And &HFFFF&
End Function

'-----------------------------------------------------------------------------------
Private Sub getTrackingConcerns(x As Long, y As Long, z As Long, ByVal groupNumber As Long, ByVal bOnScreen As Boolean)
    Dim xTotals As Long, yTotals As Long, zTotals As Long, count As Long, it As Long, psDroid As Long, take As Boolean
    xTotals = 0: yTotals = 0: zTotals = 0
    count = 0

    If selectedPlayer >= MAX_PLAYERS Then
        ' Not currently supported when selectedPlayer >= MAX_PLAYERS
        Exit Sub
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If inTrackGroup(psDroid, groupNumber) Then
            take = Not bOnScreen
            If Not take Then take = objectOnScreen(psDroid, pie_GetVideoBufferWidth() \ 4)
            If take Then
                count = count + 1
                xTotals = xTotals + aDroid(psDroid).pos.x
                yTotals = yTotals + aDroid(psDroid).pos.z   ' note the flip
                zTotals = zTotals + aDroid(psDroid).pos.y
            End If
        End If
        it = lnNext(it)
    Loop

    If count <> 0 Then ' necessary!!!!!!!
        ' (SDWORD / UDWORD: the totals are converted to unsigned)
        x = FromU32(Fix(U32(xTotals) / count))
        y = FromU32(Fix(U32(yTotals) / count))
        z = FromU32(Fix(U32(zTotals) / count))
    End If
End Sub

'-----------------------------------------------------------------------------------
' How this all works
'
'Each frame we calculate the new acceleration, velocity and positions for the location
'and rotation of the camera. The velocity is obviously based on the acceleration and this
'in turn is based on the separation between the two objects. This separation is distance
'in the case of location and degrees of arc in the case of rotation.
'
'  Each frame:-
'
'  ACCELERATION    -   A
'  VELOCITY        -   V
'  POSITION        -   P
'  Location of camera  (x1,y1)
'  Location of droid       (x2,y2)
'  Separation(distance) = D. This is the distance between (x1,y1) and (x2,y2)
'
'  A = c1D - c2V       Where c1 and c2 are two constants to be found (by experiment)
'  V = V + A(frameTime/GAME_TICKS_PER_SEC)
'  P = P + V(frameTime/GAME_TICKS_PER_SEC)
'
'  Things are the same for the rotation except that D is then the difference in angles
'  between the way the camera and droid being tracked are facing. AND.... the two
'  constants c1 and c2 will be different as we're dealing with entirely different scales
'  and units. Separation in terms of distance could be in the thousands whereas degrees
'  cannot exceed 180.
'
'  This all works because acceleration is based on how far apart they are minus some factor
'  times the camera's present velocity. This minus factor is what slows it down when the
'  separation gets very small. Without this, it would continually oscillate about it's target
'  point. The four constants (two each for rotation and position) need to be found
'  by trial and error since the metrics of time,space and rotation are entirely warzone
'  specific.
'
'  And that's all folks.

'-----------------------------------------------------------------------------------

Private Sub updateCameraAcceleration(ByVal update As Long)
    Dim concern As Vector3i, behind As Vector2i, bFlying As Boolean, angle As Long, psPropStats As Long
    Dim group As Long, multiAngle As Long, realPos As Vector3i, separation As Vector3f, acceleration As Vector3f
    Dim target As Long
    target = trackingCamera.target
    concern.x = aDroid(target).pos.x    ' pos.xzy()
    concern.y = aDroid(target).pos.z
    concern.z = aDroid(target).pos.y
    behind.x = 0: behind.y = 0          ' Irrelevant for normal radar tracking
    bFlying = False

    '
    '    This is where we check what it is we're tracking.
    '    Were we to track a building or location - this is
    '    where it'd be set up
    '
    '
    '    If we're tracking a droid, then we need to track slightly in front
    '    of it in order that the droid appears down the screen a bit. This means
    '    that we need to find an offset point from it relative to it's present
    '    direction
    '
    angle = 90 - Abs((playerPos.r.x \ 182) Mod 90)

    psPropStats = DROID_getPropulsionStats(target)

    If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
        bFlying = True
    End If

    ' Present direction is important
    If getNumDroidsSelected() > 2 Then
        If aDroid(target).selected Then group = GROUP_SELECTED Else group = aDroid(target).group

        multiAngle = getAverageTrackAngle(group, True)
        getTrackingConcerns concern.x, concern.y, concern.z, group, True

        behind.x = iSinR(multiAngle, CAM_DEFAULT_OFFSET) ' iSinCosR(multiAngle, CAM_DEFAULT_OFFSET)
        behind.y = iCosR(multiAngle, CAM_DEFAULT_OFFSET)
    Else
        behind.x = iSinR(aDroid(target).rot.direction, CAM_DEFAULT_OFFSET)
        behind.y = iCosR(aDroid(target).rot.direction, CAM_DEFAULT_OFFSET)
    End If

    concern.y = concern.y + angle * 5

    ' realPos = concern - Vector3i(-behind.x, 0, -behind.y)
    realPos.x = concern.x + behind.x
    realPos.y = concern.y
    realPos.z = concern.z + behind.y
    separation.x = CSng(realPos.x) - trackingCamera.position.x
    separation.y = CSng(realPos.y) - trackingCamera.position.y
    separation.z = CSng(realPos.z) - trackingCamera.position.z
    If Not bFlying Then
        acceleration.x = separation.x * ACCEL_CONSTANT - trackingCamera.velocity.x * VELOCITY_CONSTANT
        acceleration.y = separation.y * ACCEL_CONSTANT - trackingCamera.velocity.y * VELOCITY_CONSTANT
        acceleration.z = separation.z * ACCEL_CONSTANT - trackingCamera.velocity.z * VELOCITY_CONSTANT
    Else
        separation.y = separation.y / 2!
        acceleration.x = separation.x * (ACCEL_CONSTANT * 4) - trackingCamera.velocity.x * (VELOCITY_CONSTANT * 2)
        acceleration.y = separation.y * (ACCEL_CONSTANT * 4) - trackingCamera.velocity.y * (VELOCITY_CONSTANT * 2)
        acceleration.z = separation.z * (ACCEL_CONSTANT * 4) - trackingCamera.velocity.z * (VELOCITY_CONSTANT * 2)
    End If

    If (update And X_UPDATE) <> 0 Then
        ' Need to update acceleration along x axis
        trackingCamera.acceleration.x = acceleration.x
    End If

    If (update And Y_UPDATE) <> 0 Then
        ' Need to update acceleration along y axis
        trackingCamera.acceleration.y = acceleration.y
    End If

    If (update And Z_UPDATE) <> 0 Then
        ' Need to update acceleration along z axis
        trackingCamera.acceleration.z = acceleration.z
    End If
End Sub

'-----------------------------------------------------------------------------------

Private Sub updateCameraVelocity(ByVal update As Long)
    If (update And X_UPDATE) <> 0 Then
        trackingCamera.velocity.x = trackingCamera.velocity.x + realTimeAdjustedIncrement(trackingCamera.acceleration.x)
    End If

    If (update And Y_UPDATE) <> 0 Then
        trackingCamera.velocity.y = trackingCamera.velocity.y + realTimeAdjustedIncrement(trackingCamera.acceleration.y)
    End If

    If (update And Z_UPDATE) <> 0 Then
        trackingCamera.velocity.z = trackingCamera.velocity.z + realTimeAdjustedIncrement(trackingCamera.acceleration.z)
    End If
End Sub

'-----------------------------------------------------------------------------------

Private Sub updateCameraPosition(ByVal update As Long)
    If (update And X_UPDATE) <> 0 Then
        ' Need to update position along x axis
        trackingCamera.position.x = trackingCamera.position.x + realTimeAdjustedIncrement(trackingCamera.velocity.x)
    End If

    If (update And Y_UPDATE) <> 0 Then
        ' Need to update position along y axis
        trackingCamera.position.y = trackingCamera.position.y + realTimeAdjustedIncrement(trackingCamera.velocity.y)
    End If

    If (update And Z_UPDATE) <> 0 Then
        ' Need to update position along z axis
        trackingCamera.position.z = trackingCamera.position.z + realTimeAdjustedIncrement(trackingCamera.velocity.z)
    End If
End Sub

'-----------------------------------------------------------------------------------
' Calculate the acceleration that the camera spins around at
Private Sub updateCameraRotationAcceleration(ByVal update As Long)
    Dim worldAngle As Long, separation As Single, xConcern As Long, yConcern As Long, zConcern As Long
    Dim bTooLow As Boolean, psPropStats As Long, bGotFlying As Boolean, xPos As Long, yPos As Long, zPos As Long
    Dim droidHeight As Long, difHeight As Long, droidMapHeight As Long, pitch As Long, group As Long, target As Long
    bGotFlying = False
    xPos = 0: yPos = 0: zPos = 0
    target = trackingCamera.target

    bTooLow = False
    psPropStats = DROID_getPropulsionStats(target)
    If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
        bGotFlying = True
        droidHeight = aDroid(target).pos.z
        droidMapHeight = map_Height(aDroid(target).pos.x, aDroid(target).pos.y)
        difHeight = Abs(droidHeight - droidMapHeight)
        If difHeight < MIN_TRACK_HEIGHT Then
            bTooLow = True
        End If
    End If

    If (update And Y_UPDATE) <> 0 Then
        ' Presently only y rotation being calculated - but same idea for other axes
        ' Check what we're tracking
        If getNumDroidsSelected() > 2 Then
            If aDroid(target).selected Then group = GROUP_SELECTED Else group = aDroid(target).group
            yConcern = getAverageTrackAngle(group, False)
        Else
            yConcern = aDroid(target).rot.direction
        End If
        yConcern = yConcern + DEG(180)

        Do While trackingCamera.rotation.y < 0
            trackingCamera.rotation.y = trackingCamera.rotation.y + DEG(360)
        Loop

        ' Which way are we facing?
        worldAngle = Fix(trackingCamera.rotation.y)
        separation = angleDelta(yConcern - worldAngle)

        ' Make new acceleration
        trackingCamera.rotAccel.y = ROT_ACCEL_CONSTANT * separation - ROT_VELOCITY_CONSTANT * trackingCamera.rotVel.y
    End If

    If (update And X_UPDATE) <> 0 Then
        If Not bGotFlying Then
            If aDroid(target).selected Then group = GROUP_SELECTED Else group = aDroid(target).group
            getTrackingConcerns xPos, yPos, zPos, GROUP_SELECTED, True  ' FIXME Should this be group instead of GROUP_SELECTED?
            getBestPitchToEdgeOfGrid xPos, zPos, (DEG(180) - getAverageTrackAngle(group, True)) And &HFFFF&, pitch
            ' pitch = MAX(angleDelta(pitch), DEG(14)): a uint16_t
            pitch = angleDelta(pitch)
            If pitch < DEG(14) Then pitch = DEG(14)
            pitch = pitch And &HFFFF&
            xConcern = -pitch
        Else
            xConcern = aDroid(target).rot.pitch
            xConcern = xConcern + DEG(-16)
        End If
        Do While trackingCamera.rotation.x < 0
            trackingCamera.rotation.x = trackingCamera.rotation.x + DEG(360)
        Loop
        worldAngle = Fix(trackingCamera.rotation.x)
        separation = angleDelta(xConcern - worldAngle)

        ' Make new acceleration
        ' Make this really slow
        trackingCamera.rotAccel.x = _
            ((ROT_ACCEL_CONSTANT) * separation - ROT_VELOCITY_CONSTANT * CSng(trackingCamera.rotVel.x))
    End If

    ' This looks a bit arse - looks like a flight sim
    If (update And Z_UPDATE) <> 0 Then
        If bTooLow Then
            zConcern = 0
        Else
            zConcern = aDroid(target).rot.roll
        End If
        Do While trackingCamera.rotation.z < 0
            trackingCamera.rotation.z = trackingCamera.rotation.z + DEG(360)
        Loop
        worldAngle = Fix(trackingCamera.rotation.z)
        separation = CSng((zConcern - worldAngle))
        If separation < DEG(-180) Then
            separation = separation + DEG(360)
        ElseIf separation > DEG(180) Then
            separation = separation - DEG(360)
        End If

        ' Make new acceleration
        ' Make this really slow
        trackingCamera.rotAccel.z = _
            ((ROT_ACCEL_CONSTANT / 1) * separation - ROT_VELOCITY_CONSTANT * CSng(trackingCamera.rotVel.z))
    End If
End Sub

'-----------------------------------------------------------------------------------
'    Calculate the velocity that the camera spins around at - just add previously
'    calculated acceleration
Private Sub updateCameraRotationVelocity(ByVal update As Long)
    If (update And Y_UPDATE) <> 0 Then
        trackingCamera.rotVel.y = trackingCamera.rotVel.y + realTimeAdjustedIncrement(trackingCamera.rotAccel.y)
    End If
    If (update And X_UPDATE) <> 0 Then
        trackingCamera.rotVel.x = trackingCamera.rotVel.x + realTimeAdjustedIncrement(trackingCamera.rotAccel.x)
    End If
    If (update And Z_UPDATE) <> 0 Then
        trackingCamera.rotVel.z = trackingCamera.rotVel.z + realTimeAdjustedIncrement(trackingCamera.rotAccel.z)
    End If
End Sub

'-----------------------------------------------------------------------------------
' Move the camera around by adding the velocity
Private Sub updateCameraRotationPosition(ByVal update As Long)
    If (update And Y_UPDATE) <> 0 Then
        trackingCamera.rotation.y = trackingCamera.rotation.y + realTimeAdjustedIncrement(trackingCamera.rotVel.y)
    End If
    If (update And X_UPDATE) <> 0 Then
        trackingCamera.rotation.x = trackingCamera.rotation.x + realTimeAdjustedIncrement(trackingCamera.rotVel.x)
    End If
    If (update And Z_UPDATE) <> 0 Then
        trackingCamera.rotation.z = trackingCamera.rotation.z + realTimeAdjustedIncrement(trackingCamera.rotVel.z)
    End If
End Sub

' Returns how far away we are from our goal in rotation
' Updates the viewpoint according to the object being tracked
Private Function camTrackCamera() As Boolean
    Dim psPropStats As Long, bFlying As Boolean
    bFlying = False

    ' Most importantly - see if the target we're tracking is dead!
    If aDroid(trackingCamera.target).died <> 0 Then
        camTrackCamera = False
        Exit Function
    End If

    updateCameraAcceleration CAM_ALL
    updateCameraVelocity CAM_ALL
    updateCameraPosition CAM_ALL

    ' Update the acceleration,velocity and rotation of the camera for rotation
    '    You can track roll as well (z axis) but it makes you ill and looks
    '    like a flight sim, so for now just pitch and orientation
    psPropStats = DROID_getPropulsionStats(trackingCamera.target)
    If asPropulsionStats(psPropStats).propulsionType = PROPULSION_TYPE_LIFT Then
        bFlying = True
    End If

    If bFlying Then
        updateCameraRotationAcceleration CAM_ALL
    Else
        updateCameraRotationAcceleration CAM_X_AND_Y
    End If

    If bFlying Then
        updateCameraRotationVelocity CAM_ALL
        updateCameraRotationPosition CAM_ALL
    Else
        updateCameraRotationVelocity CAM_X_AND_Y
        updateCameraRotationPosition CAM_X_AND_Y
    End If

    ' Update the position that's now stored in trackingCamera.position
    playerPos.p.x = Fix(trackingCamera.position.x)
    playerPos.p.y = Fix(trackingCamera.position.y)
    playerPos.p.z = Fix(trackingCamera.position.z)

    ' Update the rotations that're now stored in trackingCamera.rotation
    playerPos.r.x = Fix(trackingCamera.rotation.x)
    playerPos.r.y = Fix(trackingCamera.rotation.y)
    playerPos.r.z = Fix(trackingCamera.rotation.z)

    ' There's a minimum for this - especially when John's VTOL code lets them land vertically on cliffs
    If playerPos.r.x > DEG(360 + MAX_PLAYER_X_ANGLE) Then
        playerPos.r.x = DEG(360 + MAX_PLAYER_X_ANGLE)
    End If

    ' Clip the position to the edge of the map
    CheckScrollLimits

    ' Store away our last update as acceleration and velocity are all fn()/dt
    trackingCamera.lastUpdate = realTime
    If bFullInfo Then
        flushConsoleMessages
        printDroidInfo trackingCamera.target
    End If

    camTrackCamera = True
End Function

' Updates the viewpoint animation to jump to the location pointed at the radar
Private Sub camRadarJump()
    Dim v As Vector3f
    Animation_Vector3f_update radarJumpAnimation.position
    Animation_Vector3f_update radarJumpAnimation.rotation
    v = Animation_Vector3f_getCurrent(radarJumpAnimation.position)
    playerPos.p.x = Fix(v.x): playerPos.p.y = Fix(v.y): playerPos.p.z = Fix(v.z)
    v = Animation_Vector3f_getCurrent(radarJumpAnimation.rotation)
    playerPos.r.x = Fix(v.x): playerPos.r.y = Fix(v.y): playerPos.r.z = Fix(v.z)

    If Not Animation_Vector3f_isActive(radarJumpAnimation.position) And Not Animation_Vector3f_isActive(radarJumpAnimation.rotation) Then
        trackingCamera.status = CAM_INACTIVE
    End If
End Sub

'-----------------------------------------------------------------------------------
Public Function getTrackingDroid() As Long
    If Not getWarCamStatus() Then
        getTrackingDroid = 0
        Exit Function
    End If
    If trackingCamera.status <> CAM_TRACK_DROID Then
        getTrackingDroid = 0
        Exit Function
    End If
    getTrackingDroid = trackingCamera.target
End Function

'-----------------------------------------------------------------------------------
Public Function getNumDroidsSelected() As Long
    getNumDroidsSelected = selNumSelected(selectedPlayer)
End Function

'-----------------------------------------------------------------------------------

' Returns whether or not the tracking camera is active
Public Function getWarCamStatus() As Boolean
    ' Is it switched off?
    If trackingCamera.status = CAM_INACTIVE Then
        getWarCamStatus = False
    Else
        ' Tracking is ON
        getWarCamStatus = True
    End If
End Function

'-----------------------------------------------------------------------------------

' Flips the status of tracking to the opposite of what it presently is
Public Sub camToggleStatus()
    ' If it's off
    If trackingCamera.status = CAM_INACTIVE Then
        ' Switch it on
        setWarCamActive True
    Else
        ' Otherwise, switch it off
        setWarCamActive False
    End If
End Sub

'    Flips on/off whether we print out full info about the droid being tracked.
'    If ON then this info is permanent on screen and realtime updating
Public Sub camToggleInfo()
    bFullInfo = Not bFullInfo
End Sub

' Informs the tracking camera that we want to start tracking to a new radar target
Public Sub requestRadarTrack(ByVal x As Long, ByVal y As Long)
    Dim initialPosition As Vector3f, targetPosition As Vector3f, animationDuration As Long, finalRotation As Vector3i
    Dim lenTarget As Double, r As Vector3f, f As Vector3f
    initialPosition.x = playerPos.p.x: initialPosition.y = playerPos.p.y: initialPosition.z = playerPos.p.z
    targetPosition.x = x
    targetPosition.y = calculateCameraHeightAt(map_coord(x), map_coord(y))
    targetPosition.z = y
    lenTarget = Sqr(CDbl(targetPosition.x - initialPosition.x) ^ 2 + CDbl(targetPosition.y - initialPosition.y) ^ 2 + CDbl(targetPosition.z - initialPosition.z) ^ 2)
    ' static_cast<uint32_t>(glm::log(glm::length(targetPosition - initialPosition)) * 100) (VB: Log of 0 is an error: no animation)
    If lenTarget > 1# Then animationDuration = Fix(Log(lenTarget) * 100) Else animationDuration = 0
    If trackingCamera.status = CAM_TRACK_DROID Then finalRotation = trackingCamera.oldView.r Else finalRotation = playerPos.r
    finalRotation.z = 0

    Animation_Vector3f_setInitialData radarJumpAnimation.position, initialPosition
    Animation_Vector3f_setFinalData radarJumpAnimation.position, targetPosition
    Animation_Vector3f_setDuration radarJumpAnimation.position, animationDuration
    Animation_Vector3f_start radarJumpAnimation.position

    r.x = playerPos.r.x: r.y = playerPos.r.y: r.z = playerPos.r.z
    f.x = finalRotation.x: f.y = finalRotation.y: f.z = finalRotation.z
    Animation_Vector3f_setInitialData radarJumpAnimation.rotation, r
    Animation_Vector3f_setFinalData radarJumpAnimation.rotation, f
    Animation_Vector3f_setDuration radarJumpAnimation.rotation, animationDuration
    Animation_Vector3f_start radarJumpAnimation.rotation

    trackingCamera.status = CAM_RADAR_JUMP
End Sub

' Returns whether we're presently tracking to a new _location_
Public Function getRadarTrackingStatus() As Boolean
    getRadarTrackingStatus = (trackingCamera.status = CAM_RADAR_JUMP)
End Function

Public Sub camInformOfRotation(rot_ As Vector3i)
    trackingCamera.rotation.x = rot_.x
    trackingCamera.rotation.y = rot_.y
    trackingCamera.rotation.z = rot_.z
End Sub
