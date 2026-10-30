Attribute VB_Name = "animation_c"
' src/animation.cpp, animation.h: ValueTracker, Animation<AnimatableData> and
' RotationAnimation. A class is a UDT and its methods procedures that take it
' first (ValueTracker_update(t)); the methods that return this for chaining
' are Subs (chains become statements). Animation<T> has one copy per type used:
' Animation_float (Animation<float>) and Animation_Vector3f (Animation<Vector3f>,
' with RotationAnimation's virtual start chosen by isRotation). The time an
' Animation reads (C++'s uint32_t const &time) is one of the game clocks:
' ANIMATION_TIME_REAL (realTime) or ANIMATION_TIME_GRAPHICS (graphicsTime).
' ValueTracker's speed starts at 10: a tracker comes from ValueTracker_new
' (std::make_unique<ValueTracker>()), an Animation from Animation_float_new.
Option Explicit

Public Type ValueTracker
    startTime As Long                   ' UDWORD
    initial As Long
    target As Long
    targetDelta As Long
    reachedTarget_ As Boolean           ' _reachedTarget
    current As Double
    speed As Long                       ' = 10
End Type

Public Enum EasingType
    LINEAR
    EASE_IN_OUT
    EASE_IN
    EASE_OUT
End Enum

' (VB) the clock an Animation follows (its constructor's uint32_t const &time)
Public Const ANIMATION_TIME_REAL As Long = 0
Public Const ANIMATION_TIME_GRAPHICS As Long = 1

Public Type Animation_float
    time_ As Long                       ' (VB) ANIMATION_TIME_*: uint32_t const &time
    easingType As EasingType
    duration As Long                    ' uint32_t
    startTime As Long                   ' uint32_t = 0
    progress As Long                    ' uint16_t = UINT16_MAX
    initialData As Single
    finalData As Single
    currentData As Single
End Type

Public Type Animation_Vector3f
    time_ As Long                       ' (VB) ANIMATION_TIME_*: uint32_t const &time
    easingType As EasingType
    duration As Long                    ' uint32_t
    startTime As Long                   ' uint32_t = 0
    progress As Long                    ' uint16_t = UINT16_MAX
    initialData As Vector3f
    finalData As Vector3f
    currentData As Vector3f
    isRotation As Boolean               ' (VB) a RotationAnimation (its start overrides Animation's)
End Type

Private Const UINT16_MAX As Long = 65535

' (VB) std::make_unique<ValueTracker>(): int speed = 10;
Public Function ValueTracker_new() As ValueTracker
    ValueTracker_new.speed = 10
End Function

Public Sub ValueTracker_startTracking(t As ValueTracker, ByVal value As Long)
    t.initial = value
    t.target = value
    t.targetDelta = 0
    t.current = value
    t.startTime = graphicsTime
    t.reachedTarget_ = False
End Sub

Public Sub ValueTracker_stopTracking(t As ValueTracker)
    t.initial = 0
    t.current = 0
    t.startTime = 0
    t.reachedTarget_ = False
End Sub

Public Function ValueTracker_isTracking(t As ValueTracker) As Boolean
    ValueTracker_isTracking = (t.startTime <> 0)
End Function

Public Sub ValueTracker_setSpeed(t As ValueTracker, ByVal value As Long)
    t.speed = value
End Sub

Public Sub ValueTracker_setTargetDelta(t As ValueTracker, ByVal value As Long)
    t.targetDelta = value
    t.target = t.initial + value
    t.reachedTarget_ = False
End Sub

Public Sub ValueTracker_setTarget(t As ValueTracker, ByVal value As Long)
    t.targetDelta = value - t.initial
    t.target = value
    t.reachedTarget_ = False
End Sub

Public Sub ValueTracker_update(t As ValueTracker)
    Dim deltaRemaining As Double, adjustedChange As Double
    If t.reachedTarget_ Then
        Exit Sub
    End If

    If Abs(t.target - t.current) < 1 Then
        t.reachedTarget_ = True
        Exit Sub
    End If

    deltaRemaining = (t.initial + t.targetDelta - t.current)
    adjustedChange = deltaRemaining * CDbl(realTimeAdjustedIncrement(t.speed))

    ' prevent "over-shooting" / rubber-banding
    If deltaRemaining >= 0# And adjustedChange > deltaRemaining Then
        adjustedChange = deltaRemaining
        t.reachedTarget_ = True
    ElseIf deltaRemaining < 0# And adjustedChange < deltaRemaining Then
        adjustedChange = deltaRemaining
        t.reachedTarget_ = True
    End If

    t.current = adjustedChange + t.current
End Sub

Public Function ValueTracker_getCurrent(t As ValueTracker) As Long
    If t.reachedTarget_ Then
        ValueTracker_getCurrent = t.target
        Exit Function
    End If
    ValueTracker_getCurrent = Fix(t.current)
End Function

Public Function ValueTracker_getCurrentDelta(t As ValueTracker) As Long
    If t.reachedTarget_ Then
        ValueTracker_getCurrentDelta = t.targetDelta
        Exit Function
    End If
    ValueTracker_getCurrentDelta = Fix(t.current - t.initial)
End Function

Public Function ValueTracker_getInitial(t As ValueTracker) As Long
    ValueTracker_getInitial = t.initial
End Function

Public Function ValueTracker_getTarget(t As ValueTracker) As Long
    ValueTracker_getTarget = t.target
End Function

Public Function ValueTracker_getTargetDelta(t As ValueTracker) As Long
    ValueTracker_getTargetDelta = t.targetDelta
End Function

Public Function ValueTracker_reachedTarget(t As ValueTracker) As Boolean
    ValueTracker_reachedTarget = t.reachedTarget_
End Function

Private Function calculateEasing(ByVal easingType_ As EasingType, ByVal progress As Long) As Long
    Dim v As Long
    Select Case easingType_
    Case LINEAR
        calculateEasing = progress
        Exit Function
    Case EASE_IN_OUT
        v = iCos(UINT16_MAX \ 2 + progress \ 2) \ 2 + 32768
        If v > UINT16_MAX Then v = UINT16_MAX
        If v < 0 Then v = 0
        calculateEasing = v
        Exit Function
    Case EASE_IN
        calculateEasing = Fix((CDbl(progress) * progress) / UINT16_MAX)
        Exit Function
    Case EASE_OUT
        calculateEasing = (2 * progress - Fix((CDbl(progress) * progress) / (UINT16_MAX - 1))) And &HFFFF&
        Exit Function
    End Select

    calculateEasing = UINT16_MAX
End Function

' (VB) the clock an Animation reads
Private Function animationTime(ByVal timeSource As Long) As Long
    If timeSource = ANIMATION_TIME_GRAPHICS Then
        animationTime = graphicsTime
    Else
        animationTime = realTime
    End If
End Function

' (VB) Animation(uint32_t const &time, EasingType easingType = LINEAR, uint16_t duration = 0)
Public Function Animation_float_new(ByVal timeSource As Long, Optional ByVal easingType_ As EasingType = LINEAR, Optional ByVal duration As Long = 0) As Animation_float
    Animation_float_new.time_ = timeSource
    Animation_float_new.easingType = easingType_
    Animation_float_new.duration = duration
    Animation_float_new.startTime = 0
    Animation_float_new.progress = UINT16_MAX
End Function

Public Function Animation_Vector3f_new(ByVal timeSource As Long, Optional ByVal easingType_ As EasingType = LINEAR, Optional ByVal duration As Long = 0) As Animation_Vector3f
    Animation_Vector3f_new.time_ = timeSource
    Animation_Vector3f_new.easingType = easingType_
    Animation_Vector3f_new.duration = duration
    Animation_Vector3f_new.startTime = 0
    Animation_Vector3f_new.progress = UINT16_MAX
End Function

' RotationAnimation(uint32_t const &time, EasingType easingType = LINEAR, uint16_t duration = 0): Animation(time, easingType, duration)
Public Function RotationAnimation_new(ByVal timeSource As Long, Optional ByVal easingType_ As EasingType = LINEAR, Optional ByVal duration As Long = 0) As Animation_Vector3f
    RotationAnimation_new = Animation_Vector3f_new(timeSource, easingType_, duration)
    RotationAnimation_new.isRotation = True
End Function

' (VB) progress = MAX(0, MIN(UINT16_MAX, UINT16_MAX * deltaTime / duration)), deltaTime = time - (int64_t)startTime
Private Function animationProgress(ByVal timeSource As Long, ByVal startTime As Long, ByVal duration As Long) As Long
    Dim deltaTime As Double, v As Double
    If duration <> 0 Then
        deltaTime = U32(animationTime(timeSource)) - U32(startTime)
        v = Fix((65535# * deltaTime) / U32(duration))
        If v > UINT16_MAX Then v = UINT16_MAX
        If v < 0 Then v = 0
        animationProgress = v
    Else
        animationProgress = UINT16_MAX
    End If
End Function

Public Sub Animation_float_start(a As Animation_float)
    a.startTime = animationTime(a.time_)
    a.progress = 0
End Sub

Public Sub Animation_float_update(a As Animation_float)
    a.progress = animationProgress(a.time_, a.startTime, a.duration)

    a.currentData = a.initialData + (a.finalData - a.initialData) * (Animation_float_getEasedProgress(a) / 65535!)
End Sub

Public Function Animation_float_isActive(a As Animation_float) As Boolean
    Animation_float_isActive = (a.progress < UINT16_MAX)
End Function

Public Function Animation_float_getCurrent(a As Animation_float) As Single
    Animation_float_getCurrent = a.currentData
End Function

Public Function Animation_float_getFinalData(a As Animation_float) As Single
    Animation_float_getFinalData = a.finalData
End Function

Private Function Animation_float_getEasedProgress(a As Animation_float) As Long
    Animation_float_getEasedProgress = calculateEasing(a.easingType, a.progress)
End Function

Public Sub Animation_float_setInitialData(a As Animation_float, ByVal initial As Single)
    a.initialData = initial
    a.currentData = initial
End Sub

Public Sub Animation_float_setFinalData(a As Animation_float, ByVal final_ As Single)
    a.finalData = final_
End Sub

Public Sub Animation_float_setEasing(a As Animation_float, ByVal easing As EasingType)
    a.easingType = easing
End Sub

Public Sub Animation_float_setDuration(a As Animation_float, ByVal durationMilliseconds As Long)
    a.duration = Fix(durationMilliseconds * 0.001 * GAME_TICKS_PER_SEC)
End Sub

' virtual void start(): RotationAnimation overrides it
Public Sub Animation_Vector3f_start(a As Animation_Vector3f)
    If a.isRotation Then
        RotationAnimation_start a
        Exit Sub
    End If
    Animation_Vector3f_start_base a
End Sub

' Animation<Vector3f>::start
Private Sub Animation_Vector3f_start_base(a As Animation_Vector3f)
    a.startTime = animationTime(a.time_)
    a.progress = 0
End Sub

Public Sub Animation_Vector3f_update(a As Animation_Vector3f)
    Dim f As Single
    a.progress = animationProgress(a.time_, a.startTime, a.duration)

    f = Animation_Vector3f_getEasedProgress(a) / 65535!
    a.currentData.x = a.initialData.x + (a.finalData.x - a.initialData.x) * f
    a.currentData.y = a.initialData.y + (a.finalData.y - a.initialData.y) * f
    a.currentData.z = a.initialData.z + (a.finalData.z - a.initialData.z) * f
End Sub

Public Function Animation_Vector3f_isActive(a As Animation_Vector3f) As Boolean
    Animation_Vector3f_isActive = (a.progress < UINT16_MAX)
End Function

Public Function Animation_Vector3f_getCurrent(a As Animation_Vector3f) As Vector3f
    Animation_Vector3f_getCurrent = a.currentData
End Function

Public Function Animation_Vector3f_getFinalData(a As Animation_Vector3f) As Vector3f
    Animation_Vector3f_getFinalData = a.finalData
End Function

Private Function Animation_Vector3f_getEasedProgress(a As Animation_Vector3f) As Long
    Animation_Vector3f_getEasedProgress = calculateEasing(a.easingType, a.progress)
End Function

Public Sub Animation_Vector3f_setInitialData(a As Animation_Vector3f, initial As Vector3f)
    a.initialData = initial
    a.currentData = initial
End Sub

Public Sub Animation_Vector3f_setFinalData(a As Animation_Vector3f, final_ As Vector3f)
    a.finalData = final_
End Sub

Public Sub Animation_Vector3f_setEasing(a As Animation_Vector3f, ByVal easing As EasingType)
    a.easingType = easing
End Sub

Public Sub Animation_Vector3f_setDuration(a As Animation_Vector3f, ByVal durationMilliseconds As Long)
    a.duration = Fix(durationMilliseconds * 0.001 * GAME_TICKS_PER_SEC)
End Sub

'*
' * Find the angle equivalent to `from` in the interval between `to - 180 degrees` and to `to + 180 degrees`.
' *
' * For example:
' * - if `from` is `10 degrees` and `to` is `350 degrees`, it will return `370 degrees`.
' * - if `from` is `350 degrees` and `to` is `0 degrees`, it will return `-10 degrees`.
' *
' * Useful while animating a rotation, to always animate the shortest angle delta.
Public Function calculateRelativeAngle(ByVal from As Long, ByVal to_ As Long) As Long
    calculateRelativeAngle = (to_ And &HFFFF&) + ToInt16((from And &HFFFF&) - (to_ And &HFFFF&))
End Function

' (VB) a float's conversion to uint16_t (in range: truncation)
Private Function floatToU16(ByVal f As Single) As Long
    floatToU16 = Fix(f) And &HFFFF&
End Function

Private Sub RotationAnimation_start(a As Animation_Vector3f)
    Dim initialData As Vector3f
    a.finalData.x = floatToU16(a.finalData.x)
    a.finalData.y = floatToU16(a.finalData.y)
    a.finalData.z = floatToU16(a.finalData.z)
    initialData.x = calculateRelativeAngle(floatToU16(a.initialData.x), floatToU16(a.finalData.x))
    initialData.y = calculateRelativeAngle(floatToU16(a.initialData.y), floatToU16(a.finalData.y))
    initialData.z = calculateRelativeAngle(floatToU16(a.initialData.z), floatToU16(a.finalData.z))
    a.initialData = initialData
    Animation_Vector3f_start_base a
End Sub
