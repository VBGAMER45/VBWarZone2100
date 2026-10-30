Attribute VB_Name = "display_c"
' src/display.cpp, display.h, displaydef.h: display routines - the camera's
' scrolling, panning and rotation, the screen shake, the mouse's selection,
' orders and cursor (processInput, processMouseClickInput, dealWithLMB/RMB,
' itemUnderMouse), delivery point repositioning, the start-of-game fade and
' displayWorld, which draws the 3D scene (display3d_c's draw3DScene).
' What needs the key mappings (lib/framework/input's InputManager,
' gInputManager, the kf_* key functions), the widgets and hci (cluster G) is
' called by its C++ name and stood in for (gInputManager.contexts().X() is
' gInputManager_contexts_X). The cursor table (src/cursorselection, included
' into arnMPointers) is init_arnMPointers. std::chrono::steady_clock is the
' performance counter. The globals' C++ initialisers are display_c_init's.
' A DragBox3D / WallDrag type is DragBox3D_ / WallDrag_ (VB would take the
' variables dragBox3D / wallDrag for the types).
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' display.h
Public Const MAX_PLAYER_X_ANGLE As Long = -1
Public Const MIN_PLAYER_X_ANGLE As Long = -88 ' -90 breaks some things on Vulkan...
Public Const MAXDISTANCE_REPLAY As Long = 7000
Public Const MAXDISTANCE As Long = 5000
Public Const MINDISTANCE As Long = 0
Public Const MINDISTANCE_CONFIG As Long = 1600
Public Const MAP_ZOOM_CONFIG_STEP As Long = 200
Public Const STARTDISTANCE As Long = 2600
Public Const OLD_START_HEIGHT As Long = 1500 ' only used in savegames <= 10
Public Const CAMERA_PIVOT_HEIGHT As Long = 500
Public Const INITIAL_STARTING_PITCH As Long = -40
Public Const OLD_INITIAL_ROTATION As Long = -45 ' only used in savegames <= 10
Public Const HIDDEN_FRONTEND_WIDTH As Long = 640
Public Const HIDDEN_FRONTEND_HEIGHT As Long = 480
Public Const DEFAULT_VIEW_DISTANCE_ANIMATION_SPEED As Long = 5000

' displaydef.h
Public Const BOUNDARY_X As Long = 2
Public Const BOUNDARY_Y As Long = 2

Public Enum DragState
    DRAG_INACTIVE
    DRAG_DRAGGING
    DRAG_RELEASED
    DRAG_PLACING
End Enum

Public Type DragBox3D_
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    status As DragState
    pulse As Single                     ' = 0
End Type

Public Type WallDrag_
    pos As Vector2i
    pos2 As Vector2i
    status As DragState
End Type

Public Enum MOUSE_POINTER
    MP_ATTACH = 99
    MP_ATTACK
    MP_BRIDGE
    MP_BUILD
    MP_EMBARK
    MP_FIX
    MP_GUARD
    MP_JAM
    MP_MOVE
    MP_PICKUP
    MP_REPAIR
    MP_SELECT
    MP_LOCKON
    MP_MENSELECT
    MP_BOMB
End Enum

Public Enum SELECTION_TYPE
    SC_DROID_CONSTRUCT
    SC_DROID_DIRECT
    SC_DROID_INDIRECT
    SC_DROID_CLOSE
    SC_DROID_SENSOR
    SC_DROID_ECM
    SC_DROID_BRIDGE
    SC_DROID_RECOVERY
    SC_DROID_COMMAND
    SC_DROID_BOMBER
    SC_DROID_TRANSPORTER
    SC_DROID_SUPERTRANSPORTER
    SC_DROID_DEMOLISH
    SC_DROID_REPAIR
    SC_INVALID
End Enum

Public Enum MOUSE_TARGET
    MT_TERRAIN
    MT_RESOURCE
    MT_BLOCKING
    MT_RIVER
    MT_TRENCH
    MT_OWNSTRDAM
    MT_OWNSTROK
    MT_OWNSTRINCOMP
    MT_REPAIR
    MT_REPAIRDAM
    MT_ENEMYSTR
    MT_TRANDROID
    MT_OWNDROID
    MT_OWNDROIDDAM
    MT_ENEMYDROID
    MT_COMMAND
    MT_ARTIFACT
    MT_DAMFEATURE
    MT_SENSOR
    MT_UNUSED
    MT_CONSTRUCT
    MT_SENSORSTRUCT
    MT_SENSORSTRUCTDAM
    MT_NOTARGET                         'leave as last one
End Enum

Public dragBox3D As DragBox3D_
Public wallDrag As WallDrag_

Private Const POSSIBLE_SELECTIONS As Long = 14
Private Const POSSIBLE_TARGETS As Long = 23

' NOTE: the external file "cursorselection" is used, so you can import that into a spreadsheet, and edit it there, much easier.
Private arnMPointers(0 To POSSIBLE_TARGETS - 1, 0 To POSSIBLE_SELECTIONS - 1) As Long ' const CURSOR

Public scrollDirLeftRight As Long       ' = 0
Public scrollDirUpDown As Long          ' = 0

Private bInvertMouse As Boolean         ' = true
Private bRightClickOrders As Boolean    ' = false
Public rotateMouseKey As Long           ' optional<MOUSE_KEY_CODE> = MOUSE_RMB (0 = nullopt)
Public panMouseKey As Long              ' optional<MOUSE_KEY_CODE> = nullopt
Private bDrawShadows As Boolean         ' = true
Private bEdgeScrollOutsideWindowBounds As Boolean ' = DEFAULT_EDGE_SCROLL_OUTSIDE_WINDOW (true)

Private currentFrame As Long            ' UDWORD
Private StartOfLastFrame As Long        ' UDWORD
Private rotX As Long
Private rotY As Long
Public rotationHorizontalTracker As ValueTracker ' std::unique_ptr<ValueTracker> = std::make_unique<ValueTracker>()
Public rotationVerticalTracker As ValueTracker
Private scrollRefTime As Currency       ' std::chrono::steady_clock::time_point (the performance counter)
Private scrollSpeedLeftRight As Double  'use two directions and add them because its simple
Private scrollStepLeftRight As Double
Private scrollSpeedUpDown As Double
Private scrollStepUpDown As Double
Private mouseOverConsole As Boolean     ' = false
Private ignoreOrder As Boolean          ' = false
Private ignoreRMBC As Boolean           ' = true
Private psSelectedVtol As Long          ' DROID *
Private psDominantSelected As Long      ' DROID *
Private cameraAccel As Boolean          ' = true

Public rotActive As Boolean             ' = false
Public gameStats As Boolean             ' = false
Public lockCameraScrollWhileRotating As Boolean ' = false

' Hackety hack hack hack
Private screenShakeTable(0 To 99) As Long

Private bScreenShakeActive As Boolean   ' = false
Private screenShakeStarted As Long      ' UDWORD
Private screenShakeLength As Long       ' UDWORD

Private Const FADE_START_OF_GAME_TIME As Long = 1000
Private fadeEndTime As Long             ' UDWORD = 0

'used to determine is a weapon droid is assigned to a sensor tower or sensor droid
Private bSensorAssigned As Boolean
'used to determine if the player has selected a Las Sat structure
Private bLasSatStruct As Boolean
' If shaking is allowed
Private bShakingPermitted As Boolean    ' = true

Private viewDistanceAnimation As Animation_float ' = Animation<float>(realTime)
Private viewDistanceIncrementCooldownTime As Long ' uint32_t = 0

' Mouse X coordinate at start of panning.
Public panMouseX As Long                ' UDWORD
' Mouse Y coordinate at start of panning.
Public panMouseY As Long                ' UDWORD
Public panXTracker As ValueTracker      ' std::unique_ptr<ValueTracker> = std::make_unique<ValueTracker>()
Public panZTracker As ValueTracker
Public panActive As Boolean

' Start repositioning a delivery point.
Private flagPos As FLAG_POSITION
Private flagStructId As Long
Private flagReposVarsValid As Boolean
Private flagReposFinished As Boolean

' Indicates the priority given to any given droid
' type in a multiple droid selection, the larger the
' number, the lower the priority. The order of entries
' corresponds to the order of droid types in the DROID_TYPE
' enum in DroidDef.h
'
Private Const NUM_DROID_WEIGHTS As Long = 14
Private DroidSelectionWeights(0 To NUM_DROID_WEIGHTS - 1) As Long ' UBYTE

Private Const M_PI_ As Double = 3.14159265358979    ' M_PI

' (VB) the vbtest runs set it (vbtest_vb): the display's entry points that statsref stands in for
' (statsref_render.cpp: displayRenderLoop, processInput, displayWorld, ...) then do nothing
Public vbDisplayOff As Boolean

' (VB) the C++ initialisers of display.cpp's globals and statics (Sub Main runs it before anything else)
Public Sub display_c_init()
    Dim i As Long, shake As Variant, weights As Variant
    bInvertMouse = True
    bRightClickOrders = False
    rotateMouseKey = MOUSE_RMB
    panMouseKey = 0
    bDrawShadows = True
    bEdgeScrollOutsideWindowBounds = True ' DEFAULT_EDGE_SCROLL_OUTSIDE_WINDOW
    ignoreRMBC = True
    cameraAccel = True
    bShakingPermitted = True
    rotationHorizontalTracker = ValueTracker_new()
    rotationVerticalTracker = ValueTracker_new()
    panXTracker = ValueTracker_new()
    panZTracker = ValueTracker_new()
    viewDistanceAnimation = Animation_float_new(ANIMATION_TIME_REAL)
    shake = Array(-2, -2, -3, -4, -3, -3, -5, -4, -4, -4, _
                  -4, -5, -5, -5, -5, -7, -5, -6, -8, -6, _
                  -7, -8, -6, -4, -8, -7, -7, -7, -6, -5, _
                  -6, -5, -2, -5, -6, -3, -5, -3, -2, -4, _
                  -5, -3, -2, -0, 1, 2, 2, 1, 0, 0, _
                  0, 1, 1, 3, 2, 1, 0, 2, 3, 4, _
                  4, 2, 6, 4, 5, 3, 7, 7, 3, 6, _
                  4, 7, 9, 10, 9, 8, 6, 4, 7, 5, _
                  5, 4, 6, 2, 4, 5, 3, 3, 2, 1, _
                  1, 0, -1, -1, -2, -1, 1, 0, 1, 0)
    For i = 0 To 99
        screenShakeTable(i) = shake(i)
    Next
    ' DROID_WEAPON, DROID_SENSOR, DROID_ECM, DROID_CONSTRUCT, DROID_PERSON, DROID_CYBORG, DROID_TRANSPORTER,
    ' DROID_COMMAND, DROID_REPAIR, DROID_DEFAULT, DROID_CYBORG_CONSTRUCT, DROID_CYBORG_REPAIR,
    ' DROID_CYBORG_SUPER, DROID_SUPERTRANSPORTER
    weights = Array(3, 1, 2, 4, 3, 3, 9, 0, 4, 5, 4, 4, 3, 10)
    For i = 0 To NUM_DROID_WEIGHTS - 1
        DroidSelectionWeights(i) = weights(i)
    Next
    init_arnMPointers
End Sub

' (VB) src/cursorselection, the table display.cpp includes: arnMPointers[target][selection]
Private Sub init_arnMPointers()
    ' empty terrain tile [MT_TERRAIN]
    arnMPointers(0, 0) = CURSOR_MOVE: arnMPointers(0, 1) = CURSOR_MOVE: arnMPointers(0, 2) = CURSOR_MOVE: arnMPointers(0, 3) = CURSOR_MOVE: arnMPointers(0, 4) = CURSOR_MOVE: arnMPointers(0, 5) = CURSOR_MOVE: arnMPointers(0, 6) = CURSOR_MOVE
    arnMPointers(0, 7) = CURSOR_MOVE: arnMPointers(0, 8) = CURSOR_MOVE: arnMPointers(0, 9) = CURSOR_BOMB: arnMPointers(0, 10) = CURSOR_MOVE: arnMPointers(0, 11) = CURSOR_MOVE: arnMPointers(0, 12) = CURSOR_MOVE: arnMPointers(0, 13) = CURSOR_MOVE
    ' resource tile [MT_RESOURCE]
    arnMPointers(1, 0) = CURSOR_BUILD: arnMPointers(1, 1) = CURSOR_MOVE: arnMPointers(1, 2) = CURSOR_MOVE: arnMPointers(1, 3) = CURSOR_MOVE: arnMPointers(1, 4) = CURSOR_MOVE: arnMPointers(1, 5) = CURSOR_MOVE: arnMPointers(1, 6) = CURSOR_MOVE
    arnMPointers(1, 7) = CURSOR_MOVE: arnMPointers(1, 8) = CURSOR_MOVE: arnMPointers(1, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(1, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(1, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(1, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(1, 13) = CURSOR_NOTPOSSIBLE
    ' impassible tile [MT_BLOCKING]
    arnMPointers(2, 0) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 1) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 2) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 3) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 4) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 5) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 6) = CURSOR_NOTPOSSIBLE
    arnMPointers(2, 7) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 8) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(2, 13) = CURSOR_NOTPOSSIBLE
    ' river tile [MT_RIVER]
    arnMPointers(3, 0) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 1) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 2) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 3) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 4) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 5) = CURSOR_BRIDGE: arnMPointers(3, 6) = CURSOR_NOTPOSSIBLE
    arnMPointers(3, 7) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 8) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(3, 13) = CURSOR_NOTPOSSIBLE
    ' trench tile [MT_TRENCH]
    arnMPointers(4, 0) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 1) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 2) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 3) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 4) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 5) = CURSOR_BRIDGE: arnMPointers(4, 6) = CURSOR_NOTPOSSIBLE
    arnMPointers(4, 7) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 8) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(4, 13) = CURSOR_NOTPOSSIBLE
    ' damaged friendly structure [MT_OWNSTRDAM]
    arnMPointers(5, 0) = CURSOR_FIX: arnMPointers(5, 1) = CURSOR_GUARD: arnMPointers(5, 2) = CURSOR_GUARD: arnMPointers(5, 3) = CURSOR_GUARD: arnMPointers(5, 4) = CURSOR_GUARD: arnMPointers(5, 5) = CURSOR_GUARD: arnMPointers(5, 6) = CURSOR_MOVE
    arnMPointers(5, 7) = CURSOR_MOVE: arnMPointers(5, 8) = CURSOR_GUARD: arnMPointers(5, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(5, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(5, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(5, 12) = CURSOR_DEST: arnMPointers(5, 13) = CURSOR_MOVE
    ' undamaged friendly structure [MT_OWNSTROK]
    arnMPointers(6, 0) = CURSOR_NOTPOSSIBLE: arnMPointers(6, 1) = CURSOR_GUARD: arnMPointers(6, 2) = CURSOR_GUARD: arnMPointers(6, 3) = CURSOR_GUARD: arnMPointers(6, 4) = CURSOR_GUARD: arnMPointers(6, 5) = CURSOR_GUARD: arnMPointers(6, 6) = CURSOR_MOVE
    arnMPointers(6, 7) = CURSOR_MOVE: arnMPointers(6, 8) = CURSOR_GUARD: arnMPointers(6, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(6, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(6, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(6, 12) = CURSOR_DEST: arnMPointers(6, 13) = CURSOR_MOVE
    ' partially built friendly structure [MT_OWNSTRINCOMP]
    arnMPointers(7, 0) = CURSOR_BUILD: arnMPointers(7, 1) = CURSOR_GUARD: arnMPointers(7, 2) = CURSOR_GUARD: arnMPointers(7, 3) = CURSOR_GUARD: arnMPointers(7, 4) = CURSOR_GUARD: arnMPointers(7, 5) = CURSOR_GUARD: arnMPointers(7, 6) = CURSOR_MOVE
    arnMPointers(7, 7) = CURSOR_MOVE: arnMPointers(7, 8) = CURSOR_GUARD: arnMPointers(7, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(7, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(7, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(7, 12) = CURSOR_DEST: arnMPointers(7, 13) = CURSOR_MOVE
    ' friendly repair facility [MT_REPAIR]
    arnMPointers(8, 0) = CURSOR_SEEKREPAIR: arnMPointers(8, 1) = CURSOR_SEEKREPAIR: arnMPointers(8, 2) = CURSOR_SEEKREPAIR: arnMPointers(8, 3) = CURSOR_SEEKREPAIR: arnMPointers(8, 4) = CURSOR_SEEKREPAIR: arnMPointers(8, 5) = CURSOR_SEEKREPAIR: arnMPointers(8, 6) = CURSOR_SEEKREPAIR
    arnMPointers(8, 7) = CURSOR_SEEKREPAIR: arnMPointers(8, 8) = CURSOR_SEEKREPAIR: arnMPointers(8, 9) = CURSOR_SEEKREPAIR: arnMPointers(8, 10) = CURSOR_SEEKREPAIR: arnMPointers(8, 11) = CURSOR_SEEKREPAIR: arnMPointers(8, 12) = CURSOR_DEST: arnMPointers(8, 13) = CURSOR_SEEKREPAIR
    ' friendly damagedrepair facility [MT_REPAIRDAM]
    arnMPointers(9, 0) = CURSOR_FIX: arnMPointers(9, 1) = CURSOR_SEEKREPAIR: arnMPointers(9, 2) = CURSOR_SEEKREPAIR: arnMPointers(9, 3) = CURSOR_SEEKREPAIR: arnMPointers(9, 4) = CURSOR_SEEKREPAIR: arnMPointers(9, 5) = CURSOR_SEEKREPAIR: arnMPointers(9, 6) = CURSOR_SEEKREPAIR
    arnMPointers(9, 7) = CURSOR_SEEKREPAIR: arnMPointers(9, 8) = CURSOR_SEEKREPAIR: arnMPointers(9, 9) = CURSOR_SEEKREPAIR: arnMPointers(9, 10) = CURSOR_SEEKREPAIR: arnMPointers(9, 11) = CURSOR_SEEKREPAIR: arnMPointers(9, 12) = CURSOR_DEST: arnMPointers(9, 13) = CURSOR_SEEKREPAIR
    ' enemy structure [MT_ENEMYSTR]
    arnMPointers(10, 0) = CURSOR_MOVE: arnMPointers(10, 1) = CURSOR_ATTACK: arnMPointers(10, 2) = CURSOR_ATTACK: arnMPointers(10, 3) = CURSOR_ATTACK: arnMPointers(10, 4) = CURSOR_LOCKON: arnMPointers(10, 5) = CURSOR_JAM: arnMPointers(10, 6) = CURSOR_ATTACK
    arnMPointers(10, 7) = CURSOR_ATTACK: arnMPointers(10, 8) = CURSOR_ATTACK: arnMPointers(10, 9) = CURSOR_BOMB: arnMPointers(10, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(10, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(10, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(10, 13) = CURSOR_MOVE
    ' transport [MT_TRANDROID]
    arnMPointers(11, 0) = CURSOR_EMBARK: arnMPointers(11, 1) = CURSOR_EMBARK: arnMPointers(11, 2) = CURSOR_EMBARK: arnMPointers(11, 3) = CURSOR_EMBARK: arnMPointers(11, 4) = CURSOR_EMBARK: arnMPointers(11, 5) = CURSOR_EMBARK: arnMPointers(11, 6) = CURSOR_EMBARK
    arnMPointers(11, 7) = CURSOR_EMBARK: arnMPointers(11, 8) = CURSOR_EMBARK: arnMPointers(11, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(11, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(11, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(11, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(11, 13) = CURSOR_EMBARK
    ' friendly droid [MT_OWNDROID]
    arnMPointers(12, 0) = CURSOR_SELECT: arnMPointers(12, 1) = CURSOR_SELECT: arnMPointers(12, 2) = CURSOR_SELECT: arnMPointers(12, 3) = CURSOR_SELECT: arnMPointers(12, 4) = CURSOR_SELECT: arnMPointers(12, 5) = CURSOR_SELECT: arnMPointers(12, 6) = CURSOR_SELECT
    arnMPointers(12, 7) = CURSOR_SELECT: arnMPointers(12, 8) = CURSOR_SELECT: arnMPointers(12, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(12, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(12, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(12, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(12, 13) = CURSOR_SELECT
    ' damaged friendly droid [MT_OWNDROIDDAM]
    arnMPointers(13, 0) = CURSOR_SELECT: arnMPointers(13, 1) = CURSOR_SELECT: arnMPointers(13, 2) = CURSOR_SELECT: arnMPointers(13, 3) = CURSOR_SELECT: arnMPointers(13, 4) = CURSOR_SELECT: arnMPointers(13, 5) = CURSOR_SELECT: arnMPointers(13, 6) = CURSOR_SELECT
    arnMPointers(13, 7) = CURSOR_PICKUP: arnMPointers(13, 8) = CURSOR_SELECT: arnMPointers(13, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(13, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(13, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(13, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(13, 13) = CURSOR_FIX
    ' enemy droid [MT_ENEMYDROID]
    arnMPointers(14, 0) = CURSOR_MOVE: arnMPointers(14, 1) = CURSOR_ATTACK: arnMPointers(14, 2) = CURSOR_ATTACK: arnMPointers(14, 3) = CURSOR_ATTACK: arnMPointers(14, 4) = CURSOR_LOCKON: arnMPointers(14, 5) = CURSOR_JAM: arnMPointers(14, 6) = CURSOR_MOVE
    arnMPointers(14, 7) = CURSOR_MOVE: arnMPointers(14, 8) = CURSOR_ATTACK: arnMPointers(14, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(14, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(14, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(14, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(14, 13) = CURSOR_MOVE
    ' command droid [MT_COMMAND]
    arnMPointers(15, 0) = CURSOR_SELECT: arnMPointers(15, 1) = CURSOR_ATTACH: arnMPointers(15, 2) = CURSOR_ATTACH: arnMPointers(15, 3) = CURSOR_ATTACH: arnMPointers(15, 4) = CURSOR_ATTACH: arnMPointers(15, 5) = CURSOR_ATTACH: arnMPointers(15, 6) = CURSOR_ATTACH
    arnMPointers(15, 7) = CURSOR_ATTACH: arnMPointers(15, 8) = CURSOR_SELECT: arnMPointers(15, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(15, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(15, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(15, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(15, 13) = CURSOR_ATTACH
    ' artefact [MT_ARTIFACT]
    arnMPointers(16, 0) = CURSOR_PICKUP: arnMPointers(16, 1) = CURSOR_PICKUP: arnMPointers(16, 2) = CURSOR_PICKUP: arnMPointers(16, 3) = CURSOR_PICKUP: arnMPointers(16, 4) = CURSOR_PICKUP: arnMPointers(16, 5) = CURSOR_PICKUP: arnMPointers(16, 6) = CURSOR_PICKUP
    arnMPointers(16, 7) = CURSOR_PICKUP: arnMPointers(16, 8) = CURSOR_PICKUP: arnMPointers(16, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(16, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(16, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(16, 12) = CURSOR_PICKUP: arnMPointers(16, 13) = CURSOR_PICKUP
    ' feature that can be damaged ie attacked [MT_DAMFEATURE]
    arnMPointers(17, 0) = CURSOR_NOTPOSSIBLE: arnMPointers(17, 1) = CURSOR_ATTACK: arnMPointers(17, 2) = CURSOR_ATTACK: arnMPointers(17, 3) = CURSOR_ATTACK: arnMPointers(17, 4) = CURSOR_LOCKON: arnMPointers(17, 5) = CURSOR_NOTPOSSIBLE: arnMPointers(17, 6) = CURSOR_NOTPOSSIBLE
    arnMPointers(17, 7) = CURSOR_ATTACK: arnMPointers(17, 8) = CURSOR_ATTACK: arnMPointers(17, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(17, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(17, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(17, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(17, 13) = CURSOR_NOTPOSSIBLE
    ' sensor droid [MT_SENSOR]
    arnMPointers(18, 0) = CURSOR_SELECT: arnMPointers(18, 1) = CURSOR_GUARD: arnMPointers(18, 2) = CURSOR_LOCKON: arnMPointers(18, 3) = CURSOR_SELECT: arnMPointers(18, 4) = CURSOR_SELECT: arnMPointers(18, 5) = CURSOR_SELECT: arnMPointers(18, 6) = CURSOR_SELECT
    arnMPointers(18, 7) = CURSOR_SELECT: arnMPointers(18, 8) = CURSOR_SELECT: arnMPointers(18, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(18, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(18, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(18, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(18, 13) = CURSOR_FIX
    ' wrecked building feature [MT_UNUSED] FIXME: ?
    arnMPointers(19, 0) = CURSOR_DEST: arnMPointers(19, 1) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 2) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 3) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 4) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 5) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 6) = CURSOR_NOTPOSSIBLE
    arnMPointers(19, 7) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 8) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(19, 13) = CURSOR_NOTPOSSIBLE
    ' construction droid [MT_CONSTRUCT]
    arnMPointers(20, 0) = CURSOR_SELECT: arnMPointers(20, 1) = CURSOR_GUARD: arnMPointers(20, 2) = CURSOR_SELECT: arnMPointers(20, 3) = CURSOR_SELECT: arnMPointers(20, 4) = CURSOR_SELECT: arnMPointers(20, 5) = CURSOR_SELECT: arnMPointers(20, 6) = CURSOR_SELECT
    arnMPointers(20, 7) = CURSOR_SELECT: arnMPointers(20, 8) = CURSOR_SELECT: arnMPointers(20, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(20, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(20, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(20, 12) = CURSOR_NOTPOSSIBLE: arnMPointers(20, 13) = CURSOR_FIX
    ' sensor structure [MT_SENSORSTRUCT]
    arnMPointers(21, 0) = CURSOR_NOTPOSSIBLE: arnMPointers(21, 1) = CURSOR_GUARD: arnMPointers(21, 2) = CURSOR_LOCKON: arnMPointers(21, 3) = CURSOR_GUARD: arnMPointers(21, 4) = CURSOR_GUARD: arnMPointers(21, 5) = CURSOR_GUARD: arnMPointers(21, 6) = CURSOR_MOVE
    arnMPointers(21, 7) = CURSOR_MOVE: arnMPointers(21, 8) = CURSOR_GUARD: arnMPointers(21, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(21, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(21, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(21, 12) = CURSOR_DEST: arnMPointers(21, 13) = CURSOR_MOVE
    ' damaged sensor structure. [MT_SENSORSTRUCTDAM]
    arnMPointers(22, 0) = CURSOR_FIX: arnMPointers(22, 1) = CURSOR_GUARD: arnMPointers(22, 2) = CURSOR_LOCKON: arnMPointers(22, 3) = CURSOR_GUARD: arnMPointers(22, 4) = CURSOR_GUARD: arnMPointers(22, 5) = CURSOR_GUARD: arnMPointers(22, 6) = CURSOR_MOVE
    arnMPointers(22, 7) = CURSOR_MOVE: arnMPointers(22, 8) = CURSOR_GUARD: arnMPointers(22, 9) = CURSOR_NOTPOSSIBLE: arnMPointers(22, 10) = CURSOR_NOTPOSSIBLE: arnMPointers(22, 11) = CURSOR_NOTPOSSIBLE: arnMPointers(22, 12) = CURSOR_DEST: arnMPointers(22, 13) = CURSOR_MOVE
End Sub

Public Sub animateToViewDistance(ByVal target As Single, Optional ByVal speed As Single = DEFAULT_VIEW_DISTANCE_ANIMATION_SPEED)
    Animation_float_setInitialData viewDistanceAnimation, getViewDistance()
    Animation_float_setFinalData viewDistanceAnimation, target
    If Animation_float_isActive(viewDistanceAnimation) Then
        Animation_float_setEasing viewDistanceAnimation, EASE_OUT
    Else
        Animation_float_setEasing viewDistanceAnimation, EASE_IN_OUT
    End If
    If speed <= 0 Then
        Animation_float_setDuration viewDistanceAnimation, 0
    Else
        Animation_float_setDuration viewDistanceAnimation, Fix(Log(Abs(target - getViewDistance())) * 100 * DEFAULT_VIEW_DISTANCE_ANIMATION_SPEED / speed)
    End If
    Animation_float_start viewDistanceAnimation
End Sub

Public Sub incrementViewDistance(ByVal amount As Single)
    Dim target As Single
    If InGameOpUp Or bDisplayMultiJoiningStatus Or isInGamePopupUp Or ULT(realTime, viewDistanceIncrementCooldownTime) Then
        Exit Sub
    End If

    viewDistanceIncrementCooldownTime = realTime + GAME_TICKS_PER_SEC \ 50
    If Animation_float_isActive(viewDistanceAnimation) Then
        target = Animation_float_getFinalData(viewDistanceAnimation) + amount
    Else
        target = getViewDistance() + amount
    End If
    If Not gInputManager_debugManager_debugMappingsAllowed() Then
        If target < MINDISTANCE Then target = MINDISTANCE
        If Not NETisReplay() Then
            If target > MAXDISTANCE Then target = MAXDISTANCE
        Else
            If target > MAXDISTANCE_REPLAY Then target = MAXDISTANCE_REPLAY
        End If
    End If

    animateToViewDistance target
End Sub

Private Sub updateViewDistanceAnimation()
    If Animation_float_isActive(viewDistanceAnimation) Then
        Animation_float_update viewDistanceAnimation
        setViewDistance Animation_float_getCurrent(viewDistanceAnimation)
    End If
End Sub

Public Function getShakeStatus() As Boolean
    getShakeStatus = bShakingPermitted
End Function

Public Sub setShakeStatus(ByVal val_ As Boolean)
    bShakingPermitted = val_
End Sub

Public Sub shakeStart(ByVal length As Long)
    If bShakingPermitted Then
        If Not bScreenShakeActive Then
            bScreenShakeActive = True
            screenShakeStarted = gameTime
            screenShakeLength = length
        End If
    End If
End Sub

Public Sub shakeStop()
    bScreenShakeActive = False
    playerPos.r.z = 0
End Sub

Private Sub shakeUpdate()
    Dim screenShakePercentage As Long

    ' Check if we're shaking the screen or not
    If bScreenShakeActive Then
        ' PERCENT(a, b): (a) * 100 / (b), in UDWORD
        screenShakePercentage = Fix((U32(gameTime - screenShakeStarted) * 100) / U32(screenShakeLength))
        If screenShakePercentage < 100 Then
            playerPos.r.z = 0 + DEG(screenShakeTable(screenShakePercentage))
        End If
        If U32(gameTime) > U32(screenShakeStarted) + U32(screenShakeLength) Then
            bScreenShakeActive = False
            playerPos.r.z = 0
        End If
    Else
        If Not getWarCamStatus() Then
            playerPos.r.z = 0
        End If
    End If
End Sub

Public Function getCameraAccel() As Boolean
    getCameraAccel = cameraAccel
End Function

Public Sub setCameraAccel(ByVal val_ As Boolean)
    cameraAccel = val_
End Sub

Public Function getInvertMouseStatus() As Boolean
    getInvertMouseStatus = bInvertMouse
End Function

Public Sub setInvertMouseStatus(ByVal val_ As Boolean)
    bInvertMouse = val_
End Sub

Public Function getRightClickOrders() As Boolean
    getRightClickOrders = bRightClickOrders
End Function

Public Sub setRightClickOrders(ByVal val_ As Boolean)
    If val_ = bRightClickOrders Then
        Exit Sub
    End If

    bRightClickOrders = val_
    If bRightClickOrders Then
        ' Check for conflicting rotate / pan mouse button settings, and reset to a saner default
        ' (VB) rotateMouseKey.value_or(MOUSE_END): 0 is nullopt
        If rotateMouseKey = MOUSE_RMB Then
            setRotateMouseKey MOUSE_MMB
        End If
        If panMouseKey = MOUSE_RMB Then
            panMouseKey = 0
        End If
    End If
End Sub

' optional<MOUSE_KEY_CODE>: 0 = nullopt (MOUSE_KEY_CODE has no 0 member)
Public Function getRotateMouseKey() As Long
    getRotateMouseKey = rotateMouseKey
End Function

Public Function setRotateMouseKey(ByVal key As Long) As Boolean
    If key = 0 Then
        rotateMouseKey = key
        setRotateMouseKey = True
        Exit Function
    End If

    ' otherwise, check for conflicts
    If key = MOUSE_LMB Then Exit Function
    If bRightClickOrders And key = MOUSE_RMB Then Exit Function
    If key = panMouseKey Then
        panMouseKey = 0
    End If

    rotateMouseKey = key
    setRotateMouseKey = True
End Function

Public Function getPanMouseKey() As Long
    getPanMouseKey = panMouseKey
End Function

Public Function setPanMouseKey(ByVal key As Long) As Boolean
    If key = 0 Then
        panMouseKey = key
        setPanMouseKey = True
        Exit Function
    End If

    ' otherwise, check for conflicts
    If key = MOUSE_LMB Then Exit Function
    If bRightClickOrders And key = MOUSE_RMB Then Exit Function
    If key = rotateMouseKey Then
        rotateMouseKey = 0
    End If

    panMouseKey = key
    setPanMouseKey = True
End Function

Public Function getDrawShadows() As Boolean
    getDrawShadows = bDrawShadows
End Function

Public Sub setDrawShadows(ByVal val_ As Boolean)
    bDrawShadows = val_
    pie_setShadows val_
End Sub

Public Sub setEdgeScrollOutsideWindowBounds(ByVal val_ As Boolean)
    bEdgeScrollOutsideWindowBounds = val_
End Sub

Public Function getEdgeScrollOutsideWindowBounds() As Boolean
    getEdgeScrollOutsideWindowBounds = bEdgeScrollOutsideWindowBounds
End Function

' reset the input state
Public Sub resetInput()
    rotActive = False
    dragBox3D.status = DRAG_INACTIVE
    wallDrag.status = DRAG_INACTIVE
    gInputManager_contexts_resetStates
End Sub

Private Function localPlayerHasSelection() As Boolean
    Dim it As Long
    If selectedPlayer >= MAX_PLAYERS Then
        Exit Function
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        If aDroid(lnObj(it)).selected Then
            localPlayerHasSelection = True
            Exit Function
        End If
        it = lnNext(it)
    Loop

    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        If aStruct(lnObj(it)).selected Then
            localPlayerHasSelection = True
            Exit Function
        End If
        it = lnNext(it)
    Loop
End Function

' Process the user input. This just processes the key input and jumping around the radar
Public Sub processInput()
    If vbDisplayOff Then Exit Sub        ' (VB) the vbtest runs: statsref's stand-in does nothing
    Dim mOverConstruction As Boolean, mouseIsOverScreenOverlayChild As Boolean, allowMouseWheelEvents As Boolean
    If InGameOpUp Or isInGamePopupUp Then
        dragBox3D.status = DRAG_INACTIVE   ' disengage the dragging since it stops menu input
    End If

    StartOfLastFrame = currentFrame
    currentFrame = frameGetFrameNumber()

    ignoreRMBC = False

    mOverConstruction = CoordInBuild(mouseX(), mouseY())
    mouseIsOverScreenOverlayChild = isMouseOverScreenOverlayChild(mouseX(), mouseY())

    If Not mouseIsOverScreenOverlayChild Then
        mouseOverConsole = mouseOverHistoryConsoleBox()

        ' Process all of our key mappings
        If mOverConstruction Then
            If mousePressed(MOUSE_WUP) Then
                kf_BuildPrevPage
            End If

            If mousePressed(MOUSE_WDN) Then
                kf_BuildNextPage
            End If
        End If
    End If

    If localPlayerHasSelection() Then
        gInputManager_contexts_set InputContext_DEBUG_HAS_SELECTION, InputContext_State_ACTIVE
    Else
        gInputManager_contexts_set InputContext_DEBUG_HAS_SELECTION, InputContext_State_INACTIVE
    End If
    gInputManager_contexts_updatePriorityStatus

    If Not isInTextInputMode() Then
        allowMouseWheelEvents = Not mouseIsOverScreenOverlayChild And Not mouseOverConsole And Not mOverConstruction
        gInputManager_processMappings allowMouseWheelEvents
    End If
    ' Allow the user to clear the (Active) console if need be
    If Not mouseIsOverScreenOverlayChild Then
        If mouseOverConsoleBox() Then
            If mousePressed(MOUSE_LMB) Then
                clearActiveConsole
            End If
        End If
    End If
End Sub

Private Function OverRadarAndNotDragging() As Boolean
    OverRadarAndNotDragging = isMouseOverRadar() And dragBox3D.status <> DRAG_DRAGGING And wallDrag.status <> DRAG_DRAGGING
End Function

Private Sub CheckFinishedDrag(ByVal selection As SELECTION_TYPE)
    If mouseReleased(MOUSE_LMB) Or mouseDown(MOUSE_RMB) Then
        selectAttempt = False
        If dragBox3D.status = DRAG_DRAGGING Then
            If wallDrag.status = DRAG_DRAGGING Then
                'if invalid location keep looking for a valid one
                If (buildState = BUILD3D_VALID Or buildState = BUILD3D_FINISHED) Then
                    If BASE_STATS_hasType(sBuildDetails.psStats, STAT_STRUCTURE) Then
                        If canLineBuild() Then
                            wallDrag.pos2 = mousePos
                            wallDrag.status = DRAG_RELEASED
                        End If
                    End If
                End If
            End If

            ' Only clear if shift isn't down - this is for the drag selection box for units
            If Not ctrlShiftDown() And wallDrag.status = DRAG_INACTIVE Then
                clearSelection
            End If
            dragBox3D.status = DRAG_RELEASED
            dragBox3D.x2 = mouseX()
            dragBox3D.y2 = mouseY()
            If selection = SC_DROID_DEMOLISH Then
                If ctrlShiftDown() Then
                    handleAreaDemolition
                End If
            End If
        Else
            dragBox3D.status = DRAG_INACTIVE
            wallDrag.status = DRAG_INACTIVE
        End If
    End If
End Sub

'*
' * Demolish all structures in given area
' * Note: Does not attempt to optimize movement paths,
' * so demolishing can be a little out of order
Private Sub handleAreaDemolition()
    Dim pt1 As Vector2i, pt2 As Vector2i, worldCoord1 As Vector2i, worldCoord2 As Vector2i, tmp As Vector2i
    Dim notused1 As Long, notused2 As Long, gridList() As Long, n As Long, gi As Long, psObj As Long
    pt1.x = dragBox3D.x1: pt1.y = dragBox3D.y1
    pt2.x = dragBox3D.x2: pt2.y = dragBox3D.y2
    screenCoordToWorld pt1, worldCoord1, notused1, notused2
    screenCoordToWorld pt2, worldCoord2, notused1, notused2
    ' swap coordinates to be in increasing order.. otherwise gridIterate doesn't work
    tmp = worldCoord1
    If worldCoord1.x < worldCoord2.x Then worldCoord1.x = worldCoord1.x Else worldCoord1.x = worldCoord2.x
    If worldCoord1.y < worldCoord2.y Then worldCoord1.y = worldCoord1.y Else worldCoord1.y = worldCoord2.y
    If worldCoord2.x > tmp.x Then worldCoord2.x = worldCoord2.x Else worldCoord2.x = tmp.x
    If worldCoord2.y > tmp.y Then worldCoord2.y = worldCoord2.y Else worldCoord2.y = tmp.y

    debug_ 556, LOG_INFO, "handleAreaDemolition", "demolish everything in the area (%i %i) -> (%i %i)", worldCoord1.x, worldCoord1.y, worldCoord2.x, worldCoord2.y
    n = gridStartIterateArea(worldCoord1.x, worldCoord1.y, worldCoord2.x, worldCoord2.y, gridList)
    For gi = 0 To n - 1
        psObj = gridList(gi)
        If SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE And SIMPLE_OBJECT_player(psObj) = selectedPlayer Then
            ' add demolish order to queue for every selected unit
            orderSelectedObjAdd selectedPlayer, psObj, True
        End If
    Next
End Sub

Private Sub CheckStartWallDrag()
    If mousePressed(MOUSE_LMB) Then
        ' Store away the details if we're building
        ' You can start dragging walls from invalid locations so check for
        ' BUILD3D_POS or BUILD3D_VALID, used tojust check for BUILD3D_VALID.
        Dim isStructure As Boolean
        If buildState = BUILD3D_POS Or buildState = BUILD3D_VALID Then
            isStructure = BASE_STATS_hasType(sBuildDetails.psStats, STAT_STRUCTURE)
        End If
        If isStructure Then
            If canLineBuild() Then
                wallDrag.pos = mousePos
                wallDrag.pos2 = mousePos
                wallDrag.status = DRAG_PLACING
                debug_ 583, LOG_NEVER, "CheckStartWallDrag", "Start Wall Drag" & vbLf
            End If
        ElseIf intBuildSelectMode() Then 'if we were in build select mode
            'uhoh no place to build here
            audio_PlayBuildFailedOnce
        End If
    End If
End Sub

'this function is called when a location has been chosen to place a structure or a DP
Private Function CheckFinishedFindPosition() As Boolean
    Dim OverRadar As Boolean
    OverRadar = OverRadarAndNotDragging()

    ' Do not let the player position buildings 'under' the radar
    If mouseReleased(MOUSE_LMB) And Not OverRadar Then
        If deliveryReposValid() Then
            finishDeliveryPosition
            CheckFinishedFindPosition = True
            Exit Function
        ElseIf buildState = BUILD3D_VALID Then
            If BASE_STATS_hasType(sBuildDetails.psStats, STAT_STRUCTURE) Then
                If canLineBuild() Then
                    wallDrag.pos2 = mousePos
                    wallDrag.status = DRAG_RELEASED
                End If
            End If
            debug_ 615, LOG_NEVER, "CheckFinishedFindPosition", "BUILD3D_FINISHED" & vbLf
            buildState = BUILD3D_FINISHED
            CheckFinishedFindPosition = True
            Exit Function
        End If
    End If

    CheckFinishedFindPosition = False
End Function

Private Sub HandleDrag()
    Dim dragX As Long, dragY As Long
    dragX = 0: dragY = 0

    If mouseDrag(MOUSE_LMB, dragX, dragY) Then
        If Not isMouseOverRadar() And Not mouseDown(MOUSE_RMB) Then
            dragBox3D.x1 = dragX
            dragBox3D.x2 = mouseX()
            dragBox3D.y1 = dragY
            dragBox3D.y2 = mouseY()
            dragBox3D.status = DRAG_DRAGGING

            If buildState = BUILD3D_VALID Then
                If canLineBuild() Then
                    wallDrag.pos2 = mousePos
                    wallDrag.status = DRAG_DRAGGING
                End If
            End If
        End If
    End If
End Sub

'don't want to do any of these whilst in the Intelligence Screen
Public Sub processMouseClickInput()
    If vbDisplayOff Then Exit Sub        ' (VB) the vbtest runs: statsref's stand-in does nothing
    Dim i As Long, selection As SELECTION_TYPE, item As MOUSE_TARGET, OverRadar As Boolean
    Dim ObjUnderMouse As Long, ObjAllied As Boolean, ObjDepartedAlly As Boolean, it As Long, psCurr As Long
    item = MT_NOTARGET
    OverRadar = OverRadarAndNotDragging()
    selection = establishSelection(selectedPlayer)
    If Not (selection <= POSSIBLE_SELECTIONS) Then ASSERT_FAILURE "selection <= 14", "src/display.cpp:660", "processMouseClickInput", "Weirdy selection!"

    ignoreOrder = CheckFinishedFindPosition()

    CheckStartWallDrag

    HandleDrag

    CheckFinishedDrag selection

    If isMouseOverScreenOverlayChild(mouseX(), mouseY()) Then
        ' ignore clicks
        Exit Sub
    End If

    If mouseReleased(MOUSE_LMB) And Not OverRadar And dragBox3D.status <> DRAG_RELEASED And Not ignoreOrder And Not mouseOverConsole And Not bDisplayMultiJoiningStatus Then
        If bRightClickOrders Then
            dealWithRMB
        Else
            If Not bMultiPlayer And establishSelection(selectedPlayer) = SC_DROID_SUPERTRANSPORTER Then
                ' Never, *ever* let user control the transport in SP games--it breaks the scripts!
                If Not (game.typ = LEVEL_TYPE_CAMPAIGN) Then ASSERT_FAILURE "game.type == LEVEL_TYPE::CAMPAIGN", "src/display.cpp:687", "processMouseClickInput", "Game type was set incorrectly!"
            Else
                dealWithLMB
            End If
        End If
    End If

    If mouseDClicked(MOUSE_LMB) Then
        dealWithLMBDClick
    End If

    If mouseReleased(MOUSE_RMB) And Not rotActive And Not panActive And Not ignoreRMBC Then
        dragBox3D.status = DRAG_INACTIVE
        ' Pretty sure we wan't set walldrag status here aswell.
        wallDrag.status = DRAG_INACTIVE
        If bRightClickOrders Then
            dealWithLMB
        Else
            dealWithRMB
        End If
        ' Why?
        If getWarCamStatus() Then
            camToggleStatus
        End If
    End If

    ' Right mouse click kills a building placement
    If Not rotActive And mouseReleased(MOUSE_RMB) And (buildState = BUILD3D_POS Or buildState = BUILD3D_VALID) Then
        ' Stop the placement
        kill3DBuilding
    End If
    If mouseReleased(MOUSE_RMB) Then
        cancelDeliveryRepos
    End If
    If rotateMouseKey <> 0 Then
        If mouseDrag(rotateMouseKey, rotX, rotY) Then
            If Not rotActive And Not isRadarDragging() And Not getRadarTrackingStatus() Then
                ValueTracker_startTracking rotationVerticalTracker, playerPos.r.x And &HFFFF&
                ValueTracker_startTracking rotationHorizontalTracker, playerPos.r.y And &HFFFF& ' negative values caused problems with float conversion
                rotActive = True
            End If
        End If
    End If
    If panMouseKey <> 0 Then
        If mouseDrag(panMouseKey, panMouseX, panMouseY) Then
            If Not rotActive And Not panActive And Not isRadarDragging() And Not getRadarTrackingStatus() Then
                ValueTracker_startTracking panXTracker, playerPos.p.x
                ValueTracker_startTracking panZTracker, playerPos.p.z
                panActive = True
            End If
        End If
    End If

    If gamePaused() Then
        wzSetCursor CURSOR_DEFAULT
    End If
    If buildState = BUILD3D_VALID Then
        ' special casing for building
        wzSetCursor CURSOR_BUILD
    ElseIf buildState = BUILD3D_POS Then
        ' special casing for building - can't build here
        wzSetCursor CURSOR_NOTPOSSIBLE
    ElseIf selection <> SC_INVALID Then
        item = itemUnderMouse(ObjUnderMouse)
        If Not (item < POSSIBLE_TARGETS) Then ASSERT_FAILURE "item < 23", "src/display.cpp:766", "processMouseClickInput", "Weirdy target!"

        If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < 11", "src/display.cpp:768", "processMouseClickInput", "selectedPlayer is too high: %u", selectedPlayer
        ObjAllied = False
        If ObjUnderMouse <> 0 Then
            If selectedPlayer <> SIMPLE_OBJECT_player(ObjUnderMouse) Then ObjAllied = aiCheckAlliances(selectedPlayer, SIMPLE_OBJECT_player(ObjUnderMouse))
        End If
        ObjDepartedAlly = False
        If ObjUnderMouse <> 0 And bMultiPlayer Then
            ObjDepartedAlly = (NetPlay_players_difficulty(SIMPLE_OBJECT_player(ObjUnderMouse)) = AIDifficulty_HUMAN And Not NetPlay_players_allocated(SIMPLE_OBJECT_player(ObjUnderMouse)))
        End If

        If item <> MT_NOTARGET Then
            ' exceptions to the lookup table.
            If ctrlShiftDown() And isOwnObjectOfType(ObjUnderMouse, OBJ_DROID) Then
                item = MT_OWNDROID
            ElseIf specialOrderKeyDown() And isOwnObject(ObjUnderMouse) Then
                If selection = SC_DROID_REPAIR Then
                    item = MT_OWNDROIDDAM
                Else
                    ' attacking own unit
                    item = MT_ENEMYDROID
                End If
            ElseIf selection = SC_DROID_REPAIR Then
                ' We can't repair ourselves, so change it to a blocking cursor
                it = apsDroidLists(selectedPlayer).head
                Do While it <> 0
                    psCurr = lnObj(it)
                    If aDroid(psCurr).selected Then
                        If isOwnObject(ObjUnderMouse) Then
                            If aDroid(psCurr).id = SIMPLE_OBJECT_id(ObjUnderMouse) Then
                                item = MT_BLOCKING
                            End If
                        End If
                        Exit Do
                    End If
                    it = lnNext(it)
                Loop
            ElseIf selection = SC_DROID_DEMOLISH Then
                ' Can't demolish allied objects, or something that isn't built yet
                If ObjAllied And Not ObjDepartedAlly Then
                    item = MT_BLOCKING
                ElseIf ObjUnderMouse <> 0 Then
                    If SIMPLE_OBJECT_type(ObjUnderMouse) <> OBJ_STRUCTURE Then
                        item = MT_BLOCKING
                    ElseIf aStruct(ObjUnderMouse).status = SS_BLUEPRINT_PLANNED Then
                        item = MT_BLOCKING
                    End If
                End If
            ' in multiPlayer check for what kind of unit can use it (TODO)
            ElseIf bMultiPlayer And item = MT_TRANDROID Then
                If SIMPLE_OBJECT_player(ObjUnderMouse) <> selectedPlayer Then
                    item = MT_OWNDROID
                End If
            ElseIf selection = SC_DROID_CONSTRUCT Then
                ' We don't allow the build cursor under certain circumstances ....
                ' can't build if res extractors arent available.
                If item = MT_RESOURCE Then
                    i = 0
                    Do While i < numStructureStats
                        If asStructureStats(i).typ = REF_RESOURCE_EXTRACTOR Then Exit Do
                        i = i + 1
                    Loop ' find resource stat
                    If i < numStructureStats Then
                        If apStructTypeLists(selectedPlayer, i) <> AVAILABLE Then  ' check if you can build it!
                            item = MT_BLOCKING                ' don't allow build pointer.
                        End If
                    End If

                ' repair instead of sensor/guard with cons. droids.
                ElseIf item = MT_SENSOR Then
                    If ObjUnderMouse <> 0 Then                  ' something valid
                        If SIMPLE_OBJECT_type(ObjUnderMouse) = OBJ_STRUCTURE Then ' check if struct
                            If buildingDamaged(ObjUnderMouse) Then
                                item = MT_OWNSTRDAM           ' replace guard/sense with usual icons.
                            Else
                                item = MT_OWNSTROK
                            End If
                        End If
                    End If
                End If
            ElseIf item = MT_SENSOR And selection = SC_DROID_INDIRECT And (keyDown(KEY_LSHIFT) Or keyDown(KEY_RSHIFT)) Then
                selection = SC_DROID_SENSOR

            ' check the type of sensor for indirect weapons
            ElseIf (item = MT_SENSOR Or item = MT_SENSORSTRUCT Or item = MT_SENSORSTRUCTDAM) And selection = SC_DROID_INDIRECT Then
                If ObjUnderMouse <> 0 Then
                    If Not droidSensorDroidWeapon(ObjUnderMouse, psDominantSelected) Then
                        item = MT_BLOCKING
                    End If
                End If

            'check for VTOL droids being assigned to a sensor droid/structure
            ' (VB) VB's And evaluates vtolDroidSelected even where C++'s && would not: it only sets
            ' psSelectedVtol, which only this branch reads, and logs nothing
            ElseIf (item = MT_SENSOR Or item = MT_SENSORSTRUCT Or item = MT_SENSORSTRUCTDAM) And selection = SC_DROID_DIRECT And vtolDroidSelected(selectedPlayer And &HFF&) Then
                ' NB. psSelectedVtol was set by vtolDroidSelected - yes I know its horrible, but it
                ' only smells as much as the rest of display.c so I don't feel so bad
                If ObjUnderMouse <> 0 Then
                    If droidSensorDroidWeapon(ObjUnderMouse, psSelectedVtol) Then
                        selection = SC_DROID_INDIRECT
                    Else
                        item = MT_BLOCKING
                    End If
                Else
                    item = MT_BLOCKING
                End If

            'vtols cannot pick up artifacts
            ElseIf item = MT_ARTIFACT And selection = SC_DROID_DIRECT And vtolDroidSelected(selectedPlayer And &HFF&) Then
                item = MT_BLOCKING
            End If

            If item = MT_TERRAIN Then
                If terrainType(mapTile(mouseTileX, mouseTileY)) = TER_CLIFFFACE Then
                    item = MT_BLOCKING
                End If
            End If
            ' special droid at full health
            If arnMPointers(item, selection) = CURSOR_FIX And ObjUnderMouse <> 0 Then
                If SIMPLE_OBJECT_type(ObjUnderMouse) = OBJ_DROID Then
                    If Not DROID_isDamaged(ObjUnderMouse) Then
                        item = MT_OWNDROID
                    End If
                End If
            End If
            If (arnMPointers(item, selection) = CURSOR_SELECT Or _
                arnMPointers(item, selection) = CURSOR_EMBARK Or _
                arnMPointers(item, selection) = CURSOR_ATTACH Or _
                arnMPointers(item, selection) = CURSOR_LOCKON Or _
                arnMPointers(item, selection) = CURSOR_DEST) And ObjAllied _
                And (selection <> SC_DROID_DEMOLISH Or Not ObjDepartedAlly) Then
                ' If you want to do these things, just gift your unit to your ally.
                item = MT_BLOCKING
            End If

            If specialOrderKeyDown() And (selection = SC_DROID_TRANSPORTER Or selection = SC_DROID_SUPERTRANSPORTER) And _
               arnMPointers(item, selection) = CURSOR_MOVE And bMultiPlayer Then
                ' Alt+move = disembark transporter
                wzSetCursor CURSOR_DISEMBARK
            ElseIf specialOrderKeyDown() And selection = SC_DROID_DIRECT And _
                   arnMPointers(item, selection) = CURSOR_MOVE Then
                ' Alt+move = scout
                wzSetCursor CURSOR_SCOUT
            ElseIf arnMPointers(item, selection) = CURSOR_NOTPOSSIBLE And isOwnLasSat(ObjUnderMouse) Then
                wzSetCursor CURSOR_SELECT ' Special casing for LasSat
            Else
                wzSetCursor arnMPointers(item, selection)
            End If
        Else
            wzSetCursor CURSOR_DEFAULT
        End If
    Else
        item = itemUnderMouse(ObjUnderMouse)

        'exceptions, exceptions...AB 10/06/99
        If bMultiPlayer And bLasSatStruct Then
            If Not (item < POSSIBLE_TARGETS) Then ASSERT_FAILURE "item < 23", "src/display.cpp:958", "processMouseClickInput", "Weirdy target!"
            If item = MT_ENEMYDROID Or item = MT_ENEMYSTR Or item = MT_DAMFEATURE Then
                'display attack cursor
                wzSetCursor CURSOR_ATTACK
            ElseIf isOwnObjectOfType(ObjUnderMouse, OBJ_DROID) Then
                ' Special casing for selectables
                wzSetCursor CURSOR_SELECT
            ElseIf isOwnObjectOfType(ObjUnderMouse, OBJ_STRUCTURE) Then
                If lasSatStructSelected(ObjUnderMouse) Then
                    ' Special casing for selectables
                    wzSetCursor CURSOR_SELECT
                Else
                    wzSetCursor CURSOR_DEFAULT
                End If
            Else
                'display block cursor
                wzSetCursor CURSOR_NOTPOSSIBLE
            End If
        ElseIf isOwnLasSat(ObjUnderMouse) Or isOwnObjectOfType(ObjUnderMouse, OBJ_DROID) Then
            wzSetCursor CURSOR_SELECT ' Special casing for LasSat or own unit
        End If
    End If
End Sub

' (VB) ObjUnderMouse && ObjUnderMouse->player == selectedPlayer (&& ->type == type), without evaluating a null object
Private Function isOwnObject(ByVal psObj As Long) As Boolean
    If psObj = 0 Then Exit Function
    isOwnObject = (SIMPLE_OBJECT_player(psObj) = selectedPlayer)
End Function

Private Function isOwnObjectOfType(ByVal psObj As Long, ByVal typ As OBJECT_TYPE) As Boolean
    If Not isOwnObject(psObj) Then Exit Function
    isOwnObjectOfType = (SIMPLE_OBJECT_type(psObj) = typ)
End Function

' (VB) an own structure whose first weapon is a laser satellite:
' ObjUnderMouse->type == OBJ_STRUCTURE && asWeaps[0].nStat && getWeaponStats(0)->weaponSubClass == WSC_LAS_SAT
Private Function isOwnLasSat(ByVal psObj As Long) As Boolean
    If Not isOwnObjectOfType(psObj, OBJ_STRUCTURE) Then Exit Function
    If aStruct(psObj).asWeaps(0).nStat = 0 Then Exit Function
    isOwnLasSat = (asWeaponStats(BASE_OBJECT_getWeaponStats(psObj, 0)).weaponSubClass = WSC_LAS_SAT)
End Function


Private Sub calcScroll(y As Double, dydt As Double, ByVal accel As Double, ByVal decel As Double, ByVal targetVelocity As Double, ByVal dt As Double)
    Dim tMid As Double

    ' Stop instantly, if trying to change direction.
    If targetVelocity * dydt < -0.00000001 Then
        dydt = 0#
    End If

    If targetVelocity < dydt Then
        accel = -accel
        decel = -decel
    End If

    ' Decelerate if needed.
    tMid = (0# - dydt) / decel
    If tMid < 0# Then tMid = 0#
    If tMid > dt Then tMid = dt
    y = y + (dydt * tMid + decel / 2# * tMid * tMid)
    If cameraAccel Then
        dydt = dydt + (decel * tMid)
    End If
    dt = dt - (tMid)

    ' Accelerate if needed.
    tMid = (targetVelocity - dydt) / accel
    If tMid < 0# Then tMid = 0#
    If tMid > dt Then tMid = dt
    y = y + (dydt * tMid + accel / 2# * tMid * tMid)
    If cameraAccel Then
        dydt = dydt + accel * tMid
    Else
        dydt = targetVelocity
    End If
    dt = dt - tMid

    ' Continue at target velocity.
    y = y + dydt * dt
End Sub

Private Function shouldProcessEdgeScroll() As Boolean
    shouldProcessEdgeScroll = wzMouseInWindow() Or (bEdgeScrollOutsideWindowBounds And wzWindowHasFocus())
End Function

Private Sub handleCameraScrolling()
    Dim xDif As Long, yDif As Long, scroll_zoom_factor As Double, scaled_max_scroll_speed As Double, scaled_accel As Double
    Dim now As Currency, timeDiff As Double, timeDiff_double As Double, freq As Currency, ang As Double
    scroll_zoom_factor = 1# + 2# * ((getViewDistance() - MINDISTANCE) / CSng(MAXDISTANCE - MINDISTANCE))

    If cameraAccel Then
        scaled_max_scroll_speed = scroll_zoom_factor * war_GetCameraSpeed()
    Else
        scaled_max_scroll_speed = scroll_zoom_factor * (war_GetCameraSpeed() / 2#)
    End If
    scaled_accel = scaled_max_scroll_speed / 2#

    If InGameOpUp Or bDisplayMultiJoiningStatus Or isInGamePopupUp Then      ' cant scroll when menu up. or when over radar
        Exit Sub
    End If

    If lockCameraScrollWhileRotating And rotActive And (scrollDirUpDown = 0 And scrollDirLeftRight = 0) Then
        resetScroll
        Exit Sub
    End If

    If shouldProcessEdgeScroll() Then
        If mouseY() < BOUNDARY_Y Then
            scrollDirUpDown = scrollDirUpDown + 1
            wzSetCursor CURSOR_UARROW
        End If
        If mouseY() >= (pie_GetVideoBufferHeight() - BOUNDARY_Y) Then
            scrollDirUpDown = scrollDirUpDown - 1
            wzSetCursor CURSOR_DARROW
        End If
        If mouseX() < BOUNDARY_X Then
            wzSetCursor CURSOR_LARROW
            scrollDirLeftRight = scrollDirLeftRight - 1
        End If
        If mouseX() >= (pie_GetVideoBufferWidth() - BOUNDARY_X) Then
            wzSetCursor CURSOR_RARROW
            scrollDirLeftRight = scrollDirLeftRight + 1
        End If
    End If
    If scrollDirLeftRight < -1 Then scrollDirLeftRight = -1
    If scrollDirLeftRight > 1 Then scrollDirLeftRight = 1
    If scrollDirUpDown < -1 Then scrollDirUpDown = -1
    If scrollDirUpDown > 1 Then scrollDirUpDown = 1

    If scrollDirLeftRight <> 0 Or scrollDirUpDown <> 0 Then
        setWarCamActive False  ' Don't let this thing override the user trying to scroll.
    End If

    ' Apparently there's stutter if using deltaRealTime, so we have our very own delta time here, just for us.
    ' (VB) std::chrono::steady_clock: the performance counter
    QueryPerformanceCounter now
    QueryPerformanceFrequency freq
    timeDiff = CDbl(now - scrollRefTime) / CDbl(freq)
    If timeDiff >= 0.004 Then ' throttle max update rate to avoid really jerky scrolling
        scrollRefTime = now

        timeDiff_double = timeDiff
        If timeDiff_double > 0.5 Then timeDiff_double = 0.5  ' Since we're using our own time variable, which isn't updated when dragging a box, clamp the time here so releasing the box doesn't scroll to the edge of the map suddenly.

        scrollStepLeftRight = 0#
        scrollStepUpDown = 0#
        calcScroll scrollStepLeftRight, scrollSpeedLeftRight, scaled_accel, 2# * scaled_accel, scrollDirLeftRight * scaled_max_scroll_speed, timeDiff_double
        calcScroll scrollStepUpDown, scrollSpeedUpDown, scaled_accel, 2# * scaled_accel, scrollDirUpDown * scaled_max_scroll_speed, timeDiff_double

        ang = CDbl(-playerPos.r.y) * (M_PI_ / 32768#)
        ' Get x component of movement
        xDif = Fix(Cos(ang) * scrollStepLeftRight + Sin(ang) * scrollStepUpDown)
        ' Get y component of movement
        yDif = Fix(Sin(ang) * scrollStepLeftRight - Cos(ang) * scrollStepUpDown)

        ' Adjust player's position by these components
        playerPos.p.x = playerPos.p.x + xDif
        playerPos.p.z = playerPos.p.z + yDif

        CheckScrollLimits
    End If

    ' Reset scroll directions
    scrollDirLeftRight = 0
    scrollDirUpDown = 0
End Sub

Public Sub displayRenderLoop()
    If vbDisplayOff Then Exit Sub        ' (VB) the vbtest runs: statsref's stand-in does nothing
    handleCameraScrolling
    updateViewDistanceAnimation
End Sub

'
' * Reset scrolling, so we don't jump around after unpausing.
Public Sub resetScroll()
    QueryPerformanceCounter scrollRefTime
    scrollSpeedUpDown = 0#
    scrollSpeedLeftRight = 0#
    scrollDirLeftRight = 0
    scrollDirUpDown = 0
End Sub

' Checks if coordinate is inside scroll limits, returns false if not.
Public Function CheckInScrollLimits(ByVal xPos As Long, ByVal yPos As Long) As Boolean
    Dim minX As Long, maxX As Long, minY As Long, maxY As Long
    minX = world_coord(scrollMinX)
    maxX = world_coord(scrollMaxX - 1)
    minY = world_coord(scrollMinY)
    maxY = world_coord(scrollMaxY - 1)

    If (xPos < minX) Or (xPos >= maxX) Or (yPos < minY) Or (yPos >= maxY) Then
        CheckInScrollLimits = False
        Exit Function
    End If

    CheckInScrollLimits = True
End Function

' Check a coordinate is within the scroll limits, SDWORD version.
' Returns true if edge hit.
'
Public Function CheckInScrollLimitsCamera(xPos As Long, zPos As Long) As Boolean
    Dim EdgeHit As Boolean, minX As Long, minY As Long, maxX As Long, maxY As Long
    EdgeHit = False

    minX = world_coord(scrollMinX)
    maxX = world_coord(scrollMaxX - 1)
    minY = world_coord(scrollMinY)
    maxY = world_coord(scrollMaxY - 1)

    'scroll is limited to what can be seen for current campaign
    If xPos < minX Then
        xPos = minX
        EdgeHit = True
    ElseIf xPos >= maxX Then
        xPos = maxX
        EdgeHit = True
    End If

    If zPos < minY Then
        zPos = minY
        EdgeHit = True
    ElseIf zPos >= maxY Then
        zPos = maxY
        EdgeHit = True
    End If

    CheckInScrollLimitsCamera = EdgeHit
End Function

' Check the view is within the scroll limits,
' Returns true if edge hit.
'
Public Function CheckScrollLimits() As Boolean
    Dim xp As Long, zp As Long, ret As Boolean
    xp = playerPos.p.x
    zp = playerPos.p.z
    ret = CheckInScrollLimitsCamera(xp, zp)

    playerPos.p.x = xp
    playerPos.p.z = zp

    CheckScrollLimits = ret
End Function

' Do the 3D display
Public Sub displayWorld()
    If vbDisplayOff Then Exit Sub        ' (VB) the vbtest runs: statsref's stand-in does nothing
    Dim pos As Vector3i, mouseDeltaX As Long, mouseDeltaY As Long, panningSpeed As Long
    Dim horizontalMovement As Single, verticalMovement As Single, ang As Double
    Dim fMouseDeltaX As Single, fMouseDeltaY As Single
    If headlessGameMode() Then
        Exit Sub
    End If

    shakeUpdate

    If panMouseKey <> 0 And panActive Then
        If Not mouseDown(panMouseKey) Then
            panActive = False
        Else
            mouseDeltaX = mouseX() - panMouseX
            mouseDeltaY = mouseY() - panMouseY

            panningSpeed = war_GetCameraSpeed() \ 200
            If panningSpeed < 1 Then panningSpeed = 1

            ValueTracker_setTargetDelta panXTracker, mouseDeltaX * panningSpeed
            ValueTracker_update panXTracker
            horizontalMovement = ValueTracker_getCurrentDelta(panXTracker)
            ValueTracker_setTargetDelta panZTracker, mouseDeltaY * panningSpeed
            ValueTracker_update panZTracker
            verticalMovement = -1 * ValueTracker_getCurrentDelta(panZTracker)

            ang = -playerPos.r.y * (M_PI_ / 32768)
            playerPos.p.x = Fix(ValueTracker_getInitial(panXTracker) _
                + Cos(ang) * horizontalMovement _
                + Sin(ang) * verticalMovement)
            playerPos.p.z = Fix(ValueTracker_getInitial(panZTracker) _
                + Sin(ang) * horizontalMovement _
                - Cos(ang) * verticalMovement)
            CheckScrollLimits
        End If
    End If

    If rotateMouseKey <> 0 And rotActive Then
        If mouseDown(rotateMouseKey) Then
            fMouseDeltaX = mouseX() - rotX
            fMouseDeltaY = mouseY() - rotY

            ValueTracker_setTargetDelta rotationHorizontalTracker, Fix(DEG(-fMouseDeltaX) / 4)
            ValueTracker_update rotationHorizontalTracker
            playerPos.r.y = ValueTracker_getCurrent(rotationHorizontalTracker)

            If bInvertMouse Then
                fMouseDeltaY = fMouseDeltaY * -1
            End If

            ValueTracker_setTargetDelta rotationVerticalTracker, Fix(DEG(fMouseDeltaY) / 4)
            ValueTracker_update rotationVerticalTracker
            playerPos.r.x = ValueTracker_getCurrent(rotationVerticalTracker)
            If playerPos.r.x < DEG(360 + MIN_PLAYER_X_ANGLE) Then playerPos.r.x = DEG(360 + MIN_PLAYER_X_ANGLE)
            If playerPos.r.x > DEG(360 + MAX_PLAYER_X_ANGLE) Then playerPos.r.x = DEG(360 + MAX_PLAYER_X_ANGLE)
        Else
            rotActive = False
            ignoreRMBC = True
            pos.x = playerPos.r.x
            pos.y = playerPos.r.y
            pos.z = playerPos.r.z
            camInformOfRotation pos
        End If
    End If

    draw3DScene

    If fadeEndTime <> 0 Then
        If ULT(graphicsTime, fadeEndTime) Then
            fadeStartOfGame
        Else
            ' ensure the fade only happens once (per call to transitionInit() & graphicsTime init) - i.e. at game start - regardless of graphicsTime wrap-around
            fadeEndTime = 0
        End If
    End If
End Sub

Public Function transitionInit() As Boolean
    fadeEndTime = FADE_START_OF_GAME_TIME
    transitionInit = True
End Function

Private Sub fadeStartOfGame()
    Dim color As Long, delta As Single, a As Double
    color = psPalette(WZCOL_BLACK)
    delta = (CSng(U32(graphicsTime)) / CSng(U32(fadeEndTime)) - 1!)
    a = -Int(-(255! * (1! - (delta * delta * delta + 1!)))) ' std::ceil; cubic easing
    If a > 255 Then a = 255
    If a < 0 Then a = 0 ' (static_cast<uint32_t> of a value in 0..255)
    color = (color And &HFFFFFF) Or SHL(CLng(a), 24)
    pie_UniTransBoxFill 0, 0, pie_GetVideoBufferWidth(), pie_GetVideoBufferHeight(), color
End Sub

Private Function mouseInBox(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long) As Boolean
    mouseInBox = mouseX() > x0 And mouseX() < x1 And mouseY() > y0 And mouseY() < y1
End Function

Public Function DrawnInLastFrame(ByVal frame As Long) As Boolean
    DrawnInLastFrame = (frame >= StartOfLastFrame)
End Function

'
'    Returns what the mouse was clicked on. Only called if there was a mouse pressed message
'    on MOUSE_LMB. We aren't concerned here with setting selection flags - just what it
'    actually was
Public Function mouseTarget() As Long
    Dim psReturn As Long, dispX As Long, dispY As Long, dispR As Long, i As Long, it As Long, psDroid As Long
    psReturn = 0

    If mouseTileX < 0 Or mouseTileY < 0 Or mouseTileX > mapWidth - 1 Or mouseTileY > mapHeight - 1 Then
        mouseTarget = 0
        Exit Function
    End If

    ' First have a look through the droid lists
    For i = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            dispX = aDroid(psDroid).sDisplay.screenX
            dispY = aDroid(psDroid).sDisplay.screenY
            dispR = aDroid(psDroid).sDisplay.screenR

            ' Has the droid been drawn since the start of the last frame
            If BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 Then
                If DrawnInLastFrame(aDroid(psDroid).sDisplay.frameNumber) Then
                    If mouseInBox(dispX - dispR, dispY - dispR, dispX + dispR, dispY + dispR) Then
                        ' We HAVE clicked on droid!
                        psReturn = psDroid
                        ' There's no point in checking other object types
                        mouseTarget = psReturn
                        Exit Function
                    End If
                End If
            End If
            it = lnNext(it)
        Loop
    Next ' end of checking for droids

    '    Not a droid, so maybe a structure or feature?
    '    If still NULL after this then nothing
    psReturn = getTileOccupier(mouseTileX, mouseTileY)

    If psReturn = 0 Then
        psReturn = getTileBlueprintStructure(mouseTileX, mouseTileY)
    End If

    ' Send the result back - if it's null then we clicked on an area of terrain
    mouseTarget = psReturn
End Function

' Start repositioning a delivery point.
'
Public Sub startDeliveryPosition(ByVal psFlag As Long)
    Dim it As Long, psStruct As Long
    If tryingToGetLocation() Then ' if we're placing a building don't place
        Exit Sub
    End If

    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/display.cpp:1379", "startDeliveryPosition", "Invalid player (selectedPlayer: %u)", selectedPlayer: Exit Sub

    'clear the selected delivery point
    it = apsFlagPosLists(selectedPlayer).head
    Do While it <> 0
        aFlagPos(lnObj(it)).selected = False
        it = lnNext(it)
    Loop

    'set this object position to be highlighted
    aFlagPos(psFlag).selected = True
    flagPos = aFlagPos(psFlag)

    psStruct = findDeliveryFactory(psFlag)
    If psStruct = 0 Then
        flagStructId = 0 ' not a struct, just a flag.
    Else
        flagStructId = aStruct(psStruct).id
    End If
    flagReposVarsValid = True
    flagReposFinished = False

    triggerEvent TRIGGER_DELIVERY_POINT_MOVING, psStruct
End Sub

' Finished repositioning a delivery point.
'
Public Sub finishDeliveryPosition()
    Dim psStruct As Long, it As Long
    psStruct = 0

    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/display.cpp:1412", "finishDeliveryPosition", "Invalid player (selectedPlayer: %u)", selectedPlayer: Exit Sub

    If flagStructId <> 0 Then
        flagReposVarsValid = False
        psStruct = IdToStruct(flagStructId, selectedPlayer)
        If psStruct <> 0 Then
            Dim isFactoryAP As Boolean
            If STRUCTURE_isFactory(psStruct) And aStruct(psStruct).pFunctionality <> 0 Then
                isFactoryAP = (aStruct(psStruct).functionality.factory.psAssemblyPoint <> 0)
            End If
            If isFactoryAP Then
                setAssemblyPoint aStruct(psStruct).functionality.factory.psAssemblyPoint, _
                                 flagPos.coords.x, flagPos.coords.y, selectedPlayer, True
            ElseIf aStruct(psStruct).pStructureType <> -1 Then
                If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REPAIR_FACILITY And aStruct(psStruct).pFunctionality <> 0 Then
                    setAssemblyPoint aStruct(psStruct).functionality.repairFacility.psDeliveryPoint, _
                                     flagPos.coords.x, flagPos.coords.y, selectedPlayer, True
                End If
            End If
        End If
        'deselect once moved
        it = apsFlagPosLists(selectedPlayer).head
        Do While it <> 0
            aFlagPos(lnObj(it)).selected = False
            it = lnNext(it)
        Loop
    End If
    triggerEvent TRIGGER_DELIVERY_POINT_MOVED, psStruct
    flagReposFinished = True
End Sub

' Is there a valid delivery point repositioning going on.
Public Function deliveryReposValid() As Boolean
    Dim map_ As Vector2i, flagTile As Vector2i, it As Long
    If Not flagReposVarsValid Then
        deliveryReposValid = False
        Exit Function
    End If

    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/display.cpp:1450", "deliveryReposValid", "Invalid player (selectedPlayer: %u)", selectedPlayer: deliveryReposValid = False: Exit Function

    map_.x = map_coord(flagPos.coords.x)
    map_.y = map_coord(flagPos.coords.y)

    'make sure we are not too near map edge
    If map_.x < scrollMinX + TOO_NEAR_EDGE Or map_.x + 1 > scrollMaxX - TOO_NEAR_EDGE Or _
       map_.y < scrollMinY + TOO_NEAR_EDGE Or map_.y + 1 > scrollMaxY - TOO_NEAR_EDGE Then
        deliveryReposValid = False
        Exit Function
    End If

    ' cant place on top of a delivery point...
    it = apsFlagPosLists(selectedPlayer).head
    Do While it <> 0
        flagTile.x = map_coord(aFlagPos(lnObj(it)).coords.x)
        flagTile.y = map_coord(aFlagPos(lnObj(it)).coords.y)
        If flagTile.x = map_.x And flagTile.y = map_.y Then
            deliveryReposValid = False
            Exit Function
        End If
        it = lnNext(it)
    Loop

    If fpathBlockingTile(map_.x, map_.y, PROPULSION_TYPE_WHEELED) Then
        deliveryReposValid = False
        Exit Function
    End If

    deliveryReposValid = True
End Function

' psFlag: an address to copy the flag into (C++'s FLAG_POSITION *psFlag = nullptr): (VB) a FLAG_POSITION
' UDT, with hasFlag False for nullptr
Public Function deliveryReposFinished(Optional ByVal hasFlag As Boolean = False) As Boolean
    If Not flagReposVarsValid Then
        deliveryReposFinished = False
        Exit Function
    End If

    deliveryReposFinished = flagReposFinished
End Function

' (VB) deliveryReposFinished(psFlag) with psFlag set: *psFlag = flagPos
Public Function deliveryReposFinished_flag(psFlag As FLAG_POSITION) As Boolean
    If Not flagReposVarsValid Then
        deliveryReposFinished_flag = False
        Exit Function
    End If

    psFlag = flagPos
    deliveryReposFinished_flag = flagReposFinished
End Function

Public Sub processDeliveryRepos()
    If vbDisplayOff Then Exit Sub        ' (VB) the vbtest runs: statsref's stand-in does nothing
    Dim bX As Long, bY As Long
    If Not flagReposVarsValid Then
        Exit Sub
    End If

    bX = mouseTileX
    If bX < 2 Then bX = 2
    If bX > mapWidth - 3 Then bX = mapWidth - 3
    bY = mouseTileY
    If bY < 2 Then bY = 2
    If bY > mapHeight - 3 Then bY = mapHeight - 3

    flagPos.coords.x = world_coord(bX) + TILE_UNITS \ 2
    flagPos.coords.y = world_coord(bY) + TILE_UNITS \ 2
    flagPos.coords.z = map_TileHeight(bX, bY) + 2 * ASSEMBLY_POINT_Z_PADDING
End Sub

' Cancel repositioning of the delivery point without moving it.
'
Public Sub cancelDeliveryRepos()
    flagReposVarsValid = False
End Sub

' viewMatrix, perspectiveViewMatrix: mat4s (C++'s const glm::mat4 &)
Public Sub renderDeliveryRepos(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    If flagReposVarsValid Then
        renderDeliveryPoint_flag flagPos, True, viewMatrix, perspectiveViewMatrix
    End If
End Sub

' check whether a clicked on droid is in a command group or assigned to a sensor
Private Function droidHasLeader(ByVal psDroid As Long) As Boolean
    Dim psLeader As Long

    If aDroid(psDroid).droidType = DROID_COMMAND Or _
       aDroid(psDroid).droidType = DROID_SENSOR Then
        droidHasLeader = False
        Exit Function
    End If

    If hasCommander(psDroid) Then
        psLeader = aGroup(aDroid(psDroid).psGroup).psCommander
    Else
        'psLeader can be either a droid or a structure
        psLeader = orderStateObj(psDroid, DORDER_FIRESUPPORT)
    End If

    If psLeader <> 0 Then
        If SIMPLE_OBJECT_type(psLeader) = OBJ_DROID Then
            SelectDroid psLeader
        End If
        assignSensorTarget psLeader
        droidHasLeader = True
        Exit Function
    End If

    droidHasLeader = False
End Function

' deal with selecting a droid
Public Sub dealWithDroidSelect(ByVal psDroid As Long, ByVal bDragBox As Boolean)
    '    Toggle selection on and off - allows you drag around a big
    '    area of droids and then exclude certain individuals
    Dim doSelect As Boolean
    If Not bDragBox And aDroid(psDroid).selected Then
        DeSelectDroid psDroid
    Else
        ' (VB) ctrlShiftDown() || !droidHasLeader(psDroid): droidHasLeader (which selects) only when no queue key is down
        doSelect = ctrlShiftDown()
        If Not doSelect Then doSelect = Not droidHasLeader(psDroid)
        If doSelect Then
            If specialOrderKeyDown() Then
                ' We only want to select weapon units if ALT is down on a drag
                If aDroid(psDroid).asWeaps(0).nStat > 0 Then
                    SelectDroid psDroid
                End If
            Else
                SelectDroid psDroid
            End If
        End If
    End If
End Sub

Private Sub FeedbackOrderGiven()
    Static LastFrame As Long
    Dim ThisFrame As Long
    ThisFrame = frameGetFrameNumber()

    ' Ensure only played once per game cycle.
    If ThisFrame <> LastFrame Then
        audio_PlayTrack ID_SOUND_SELECT
        LastFrame = ThisFrame
    End If
End Sub

' check whether the queue order keys are pressed
Public Function ctrlShiftDown() As Boolean
    ctrlShiftDown = keyDown(KEY_LCTRL) Or keyDown(KEY_RCTRL) Or keyDown(KEY_LSHIFT) Or keyDown(KEY_RSHIFT)
End Function

Public Sub AddDerrickBurningMessage()
    If addConsoleMessageDebounced("Cannot Build. Oil Resource Burning.", DEFAULT_JUSTIFY, SYSTEM_MESSAGE, CANNOT_BUILD_BURNING) Then
        audio_PlayBuildFailedOnce
    End If
End Sub

Private Sub printDroidClickInfo(ByVal psDroid As Long)
    If gInputManager_debugManager_debugMappingsAllowed() Then ' cheating on, so output debug info
        console "%s - Hitpoints %d/%d - ID %d - experience %f, %s - order %s - action %s - sensor range %d - ECM %u - pitch %.0f - frust %u - kills %u", _
                droidGetName(psDroid), aDroid(psDroid).body, aDroid(psDroid).originalBody, aDroid(psDroid).id, _
                CDbl(CSng(aDroid(psDroid).experience / 65536!)), getDroidLevelName(psDroid), getDroidOrderName(aDroid(psDroid).order.typ), getDroidActionName(aDroid(psDroid).action), _
                droidSensorRange(psDroid), objJammerPower(psDroid), CDbl(UNDEG(aDroid(psDroid).rot.pitch)), aDroid(psDroid).lastFrustratedTime, aDroid(psDroid).kills
        FeedbackOrderGiven
    ElseIf Not aDroid(psDroid).selected Then
        console "%s - Hitpoints %d/%d - Experience %.1f, %s, Kills %d", droidGetName(psDroid), aDroid(psDroid).body, aDroid(psDroid).originalBody, _
                CDbl(CSng(aDroid(psDroid).experience / 65536!)), getDroidLevelName(psDroid), aDroid(psDroid).kills
        FeedbackOrderGiven
    End If
    clearSelection
    dealWithDroidSelect psDroid, False
End Sub

Private Sub dealWithLMBDroid(ByVal psDroid As Long, ByVal selection As SELECTION_TYPE)
    Dim ownDroid As Boolean, isSpectator As Boolean, it As Long, psCurr As Long, fits As Boolean ' ownDroid: Not an allied droid
    isSpectator = False
    If bMultiPlayer Then isSpectator = NetPlay_players_isSpectator(selectedPlayer)
    If isSpectator Then
        console "%s - Hitpoints %d/%d - ID %d - %s", droidGetName(psDroid), aDroid(psDroid).body, aDroid(psDroid).originalBody, aDroid(psDroid).id, getDroidLevelName(psDroid)
        Exit Sub
    End If

    If selectedPlayer >= MAX_PLAYERS Then
        Exit Sub ' no-op
    End If

    If Not aiCheckAlliances(selectedPlayer, aDroid(psDroid).player) Then
        DROIDDOING = "" ' take over the other players droid by debug menu.
        ' We've clicked on enemy droid
        If gInputManager_debugManager_debugMappingsAllowed() Then
            console "(Enemy!) %s - Hitpoints %d/%d - ID %d - experience %f, %s - order %s - action %s - sensor range %d - ECM %d - pitch %.0f", _
                    droidGetName(psDroid), aDroid(psDroid).body, aDroid(psDroid).originalBody, aDroid(psDroid).id, _
                    CDbl(CSng(aDroid(psDroid).experience / 65536!)), getDroidLevelName(psDroid), getDroidOrderName(aDroid(psDroid).order.typ), _
                    getDroidActionName(aDroid(psDroid).action), droidSensorRange(psDroid), objJammerPower(psDroid), CDbl(UNDEG(aDroid(psDroid).rot.pitch))
            FeedbackOrderGiven
        End If
        orderSelectedObjAdd selectedPlayer, psDroid, ctrlShiftDown()

        'lasSat structure can select a target - in multiPlayer only
        If bMultiPlayer And bLasSatStruct Then
            orderStructureObj selectedPlayer, psDroid
        End If

        FeedbackOrderGiven
        Exit Sub
    End If

    ownDroid = (selectedPlayer = aDroid(psDroid).player)
    ' Hack to detect if sensor was assigned
    bSensorAssigned = True
    If Not bRightClickOrders And ctrlShiftDown() And ownDroid Then
        ' select/deselect etc. the droid
        dealWithDroidSelect psDroid, False
    ElseIf specialOrderKeyDown() And ownDroid Then
        ' try to attack your own unit
        orderSelectedObjAdd selectedPlayer, psDroid, ctrlShiftDown()
        FeedbackOrderGiven
    ElseIf DROID_isTransporter(psDroid) Then
        If selection = SC_INVALID Then
            'in multiPlayer mode we RMB to get the interface up
            If bMultiPlayer And Not bRightClickOrders Then
                aDroid(psDroid).selected = True
                triggerEventSelected
            Else
                intResetScreen False
                If Not getWidgetsStatus() Then
                    setWidgetsStatus True
                End If
                addTransporterInterface psDroid, False
            End If
        Else
            ' We can order all units to use the transport now
            If cyborgDroidSelected(selectedPlayer) Then
                ' TODO add special processing for cyborgDroids
            End If
            orderSelectedObj selectedPlayer, psDroid
            FeedbackOrderGiven
        End If
    ' Clicked on a commander? Will link to it.
    ElseIf aDroid(psDroid).droidType = DROID_COMMAND And selection <> SC_INVALID And _
           selection <> SC_DROID_COMMAND And _
           selection <> SC_DROID_CONSTRUCT And _
           Not ctrlShiftDown() And ownDroid Then
        turnOffMultiMsg True
        orderSelectedObj selectedPlayer, psDroid
        FeedbackOrderGiven
        clearSelection
        assignSensorTarget psDroid
        dealWithDroidSelect psDroid, False
        turnOffMultiMsg False
    ' Clicked on a sensor? Will assign to it.
    ElseIf aDroid(psDroid).droidType = DROID_SENSOR Then
        bSensorAssigned = False
        it = apsDroidLists(selectedPlayer).head
        Do While it <> 0
            psCurr = lnObj(it)
            'must be indirect weapon droid or VTOL weapon droid
            fits = False
            If aDroid(psCurr).droidType = DROID_WEAPON And aDroid(psCurr).selected And aDroid(psCurr).asWeaps(0).nStat > 0 Then
                If Not proj_Direct(DROID_getWeaponStats(psCurr, 0)) Then
                    fits = True
                ElseIf DROID_isVtol(psCurr) Then
                    fits = True
                End If
                If fits Then fits = droidSensorDroidWeapon(psDroid, psCurr)
            End If
            If fits Then
                bSensorAssigned = True
                orderDroidObj psCurr, DORDER_FIRESUPPORT, psDroid, ModeQueue
                FeedbackOrderGiven
            End If
            it = lnNext(it)
        Loop
        If bSensorAssigned Then
            clearSelection
            assignSensorTarget psDroid
        End If
    ' Hack to detect if anything was done with sensor
    Else
        bSensorAssigned = False
    End If
    If bSensorAssigned Then
        Exit Sub
    ' Clicked on a construction unit? Will guard it.
    ElseIf (aDroid(psDroid).droidType = DROID_CONSTRUCT Or aDroid(psDroid).droidType = DROID_SENSOR Or _
            aDroid(psDroid).droidType = DROID_COMMAND) _
           And selection = SC_DROID_DIRECT Then
        orderSelectedObj selectedPlayer, psDroid
        FeedbackOrderGiven
    ' Clicked on a damaged unit? Will repair it.
    ElseIf DROID_isDamaged(psDroid) And repairDroidSelected(selectedPlayer) Then
        assignDestTarget
        orderSelectedObjAdd selectedPlayer, psDroid, ctrlShiftDown()
        FeedbackOrderGiven
    ElseIf bRightClickOrders And ownDroid Then
        If Not aDroid(psDroid).selected Then
            clearSelection
            SelectDroid psDroid
        End If
        intObjectSelected psDroid
    ' Just plain clicked on?
    ElseIf ownDroid Then
        printDroidClickInfo psDroid
    Else ' Clicked on allied unit with no other possible actions
        console "%s - Allied - Hitpoints %d/%d - Experience %d, %s", droidGetName(psDroid), aDroid(psDroid).body, aDroid(psDroid).originalBody, _
                aDroid(psDroid).experience \ 65536, getDroidLevelName(psDroid)
        FeedbackOrderGiven
    End If
End Sub

Private Sub dealWithLMBStructure(ByVal psStructure As Long, ByVal selection As SELECTION_TYPE)
    Dim ownStruct As Boolean, isSpectator As Boolean, shouldDisplayInterface As Boolean, it As Long, selectable As Boolean
    ownStruct = (aStruct(psStructure).player = selectedPlayer)
    isSpectator = False
    If bMultiPlayer Then isSpectator = NetPlay_players_isSpectator(selectedPlayer)

    If isSpectator Then
        printStructureInfo psStructure
        Exit Sub
    End If

    If selectedPlayer < MAX_PLAYERS Then
        If Not aiCheckAlliances(aStruct(psStructure).player, selectedPlayer) Then
            ' We've clicked on an enemy building
            If gInputManager_debugManager_debugMappingsAllowed() Then
                ' TRANSLATORS: "ref" is an internal unique id of the item (can leave untranslated as a technical term)
                console "(Enemy!) %s, ref: %d, ID: %d Hitpoints: %d/%d", BASE_STATS_id(STAT_STRUCTURE + aStruct(psStructure).pStructureType), asStructureStats(aStruct(psStructure).pStructureType).ref, _
                        aStruct(psStructure).id, aStruct(psStructure).body, asStructureStats(aStruct(psStructure).pStructureType).upgrade(aStruct(psStructure).player).hitpoints
            End If
            orderSelectedObjAdd selectedPlayer, psStructure, ctrlShiftDown()
            'lasSat structure can select a target - in multiPlayer only
            If bMultiPlayer And bLasSatStruct Then
                orderStructureObj selectedPlayer, psStructure
            End If
            FeedbackOrderGiven
            Exit Sub
        End If
    End If

    ' We've clicked on allied or own building

    'print some info at the top of the screen for the specific structure
    If Not bRightClickOrders Then
        printStructureInfo psStructure
    End If

    If selectedPlayer >= MAX_PLAYERS Then
        Exit Sub ' do not proceed
    End If

    ' Got to be built. Also, you can't 'select' derricks
    selectable = (aStruct(psStructure).status = SS_BUILT) And (aStruct(psStructure).flags And BIT_(OBJECT_FLAG_UNSELECTABLE)) = 0
    If Not specialOrderKeyDown() And selectable And _
       (asStructureStats(aStruct(psStructure).pStructureType).typ <> REF_RESOURCE_EXTRACTOR) And ownStruct Then
        If bRightClickOrders Then
            If STRUCTURE_isFactory(psStructure) And selection <> SC_DROID_CONSTRUCT Then
                intAddFactoryOrder psStructure
            End If
        Else
            shouldDisplayInterface = Not anyDroidSelected(selectedPlayer)
            If selection = SC_INVALID Then
                ' Clear old building selection(s) - should only be one
                it = apsStructLists(selectedPlayer).head
                Do While it <> 0
                    aStruct(lnObj(it)).selected = False
                    it = lnNext(it)
                Loop
                ' Establish new one
                aStruct(psStructure).selected = True
                triggerEventSelected
                jsDebugSelected psStructure
            End If
            'determine if LasSat structure has been selected
            bLasSatStruct = lasSatStructSelected(psStructure)

            If shouldDisplayInterface Then
                intObjectSelected psStructure
                FeedbackOrderGiven
            End If
        End If

    ElseIf selectable And _
           (asStructureStats(aStruct(psStructure).pStructureType).typ = REF_RESOURCE_EXTRACTOR) And _
           selection = SC_INVALID And ownStruct Then
        ' Clear old building selection(s) - should only be one
        it = apsStructLists(selectedPlayer).head
        Do While it <> 0
            aStruct(lnObj(it)).selected = False
            it = lnNext(it)
        Loop
        ' Establish new one
        aStruct(psStructure).selected = True
        triggerEventSelected
        jsDebugSelected psStructure
    End If
    bSensorAssigned = False
    orderSelectedObjAdd selectedPlayer, psStructure, ctrlShiftDown()
    FeedbackOrderGiven
    If bSensorAssigned Then
        clearSelection
        assignSensorTarget psStructure
    End If
    If intDemolishSelectMode() Then
        ' we were demolishing something - now we're done
        If ctrlShiftDown() Then
            quickQueueMode = True
        Else
            intDemolishCancel
        End If
    End If
End Sub

Private Sub dealWithLMBFeature(ByVal psFeature As Long)
    Dim i As Long, it As Long, psCurr As Long, psNearestUnit As Long, fpos As Vector2i
    If selectedPlayer >= MAX_PLAYERS Then
        GoTo debugOuput
    End If

    'go on to check for
    If asFeatureStats(aFeature(psFeature).psStats).damageable Then
        orderSelectedObjAdd selectedPlayer, psFeature, ctrlShiftDown()
        'lasSat structure can select a target - in multiPlayer only
        If bMultiPlayer And bLasSatStruct Then
            orderStructureObj selectedPlayer, psFeature
        End If
        FeedbackOrderGiven
    End If

    'clicking an oil field should start a build..
    If asFeatureStats(aFeature(psFeature).psStats).subType = FEAT_OIL_RESOURCE Then
        ' find any construction droids. and order them to build an oil resource.

        ' first find the derrick.
        i = 0
        Do While i < numStructureStats
            If asStructureStats(i).typ = REF_RESOURCE_EXTRACTOR Then Exit Do
            i = i + 1
        Loop

        If i < numStructureStats Then
            If apStructTypeLists(selectedPlayer, i) = AVAILABLE Then ' don't go any further if no derrick stat found.
                ' for each droid
                it = apsDroidLists(selectedPlayer).head
                Do While it <> 0
                    psCurr = lnObj(it)
                    If (aDroid(psCurr).droidType = DROID_CONSTRUCT Or _
                        aDroid(psCurr).droidType = DROID_CYBORG_CONSTRUCT) And aDroid(psCurr).selected Then
                        If fireOnLocation(aFeature(psFeature).pos.x, aFeature(psFeature).pos.y) Then
                            ' Can't build because it's burning
                            AddDerrickBurningMessage
                        End If

                        fpos.x = aFeature(psFeature).pos.x
                        fpos.y = aFeature(psFeature).pos.y
                        sendDroidInfo psCurr, DroidOrder_stats(DORDER_BUILD, i, fpos, playerPos.r.y), ctrlShiftDown()
                        FeedbackOrderGiven
                    End If
                    it = lnNext(it)
                Loop
            End If
        End If
    Else
        Select Case asFeatureStats(aFeature(psFeature).psStats).subType
        Case FEAT_GEN_ARTE, FEAT_OIL_DRUM
            psNearestUnit = getNearestDroid(mouseTileX * TILE_UNITS + TILE_UNITS \ 2, _
                                            mouseTileY * TILE_UNITS + TILE_UNITS \ 2, True)
            ' If so then find the nearest unit!
            If psNearestUnit <> 0 Then ' bloody well should be!!!
                sendDroidInfo psNearestUnit, DroidOrder_obj(DORDER_RECOVER, psFeature), ctrlShiftDown()
                FeedbackOrderGiven
            End If
        Case Else ' FEAT_BOULDER, FEAT_OIL_RESOURCE, FEAT_VEHICLE
        End Select
    End If

debugOuput:
    If gInputManager_debugManager_debugMappingsAllowed() Then
        console "(Feature) %s ID: %d ref: %d Hitpoints: %d/%d", BASE_STATS_id(STAT_FEATURE + aFeature(psFeature).psStats), aFeature(psFeature).id, asFeatureStats(aFeature(psFeature).psStats).ref, aFeature(psFeature).body, asFeatureStats(aFeature(psFeature).psStats).body
    End If
End Sub

Private Sub dealWithLMBObject(ByVal psClickedOn As Long)
    Dim selection As SELECTION_TYPE, typ As OBJECT_TYPE
    selection = establishSelection(selectedPlayer)
    typ = SIMPLE_OBJECT_type(psClickedOn)

    Select Case typ
    Case OBJ_DROID
        dealWithLMBDroid psClickedOn, selection

    Case OBJ_STRUCTURE
        dealWithLMBStructure psClickedOn, selection

    Case OBJ_FEATURE
        dealWithLMBFeature psClickedOn

    Case Else
        ' assert only when the value is outside of the valid range
        If Not (typ < OBJ_NUM_TYPES) Then ASSERT_FAILURE "type < OBJ_NUM_TYPES", "src/display.cpp:2020", "dealWithLMBObject", "Weird selection from LMB - type of clicked object is %d", typ
    End Select
End Sub

Public Sub dealWithLMB()
    Dim psClickedOn As Long, psStructure As Long, deliveryPoint As Long
    Dim psTile As Long, aux As Long, flipVal As Long, explored As String, danger As String, threat As String, decal As String

    ' Don't process if in game options are on screen
    If isMouseOverRadar() Or InGameOpUp = True Then
        Exit Sub
    End If
    If Not widgGetFromID(psWScreen, INTINGAMEOP) Is Nothing Then
        Exit Sub
    End If

    ' What have we clicked on?
    psClickedOn = mouseTarget()
    If psClickedOn <> 0 Then
        dealWithLMBObject psClickedOn

        Exit Sub
    End If

    If selectedPlayer >= MAX_PLAYERS Then
        Exit Sub ' shortcut the rest (for now)
    End If

    deliveryPoint = findMouseDeliveryPoint()
    If deliveryPoint <> 0 Then
        If selNumSelected(selectedPlayer) = 0 Then
            If bRightClickOrders Then
                'centre the view on the owning Factory
                psStructure = findDeliveryFactory(deliveryPoint)
                If psStructure <> 0 Then
                    setViewPos map_coord(aStruct(psStructure).pos.x), map_coord(aStruct(psStructure).pos.y), True
                End If
            Else
                startDeliveryPosition deliveryPoint
            End If
            Exit Sub
        End If
    End If

    ' now changed to use the multiple order stuff
    ' clicked on a destination.
    orderSelectedLoc selectedPlayer, mousePos.x, mousePos.y, ctrlShiftDown()  ' ctrlShiftDown() = ctrl clicked a destination, add an order
    ' Otherwise send them all
    If getNumDroidsSelected() <> 0 Then
        assignDestTarget
        audio_PlayTrack ID_SOUND_SELECT
    End If

    If gInputManager_debugManager_debugMappingsAllowed() Then
        If tileOnMap(mouseTileX, mouseTileY) Then
            psTile = mapTile(mouseTileX, mouseTileY)
            aux = auxTile(mouseTileX, mouseTileY, selectedPlayer)

            flipVal = 0
            If (TileNumber_texture(psMapTiles(psTile).texture) And TILE_XFLIP) <> 0 Then
                flipVal = flipVal + 2
            End If
            If (TileNumber_texture(psMapTiles(psTile).texture) And TILE_YFLIP) <> 0 Then
                flipVal = flipVal + 1
            End If

            If tileIsExplored(psTile) Then explored = "Explored" Else explored = "Unexplored"
            If (aux And AUXBITS_DANGER) <> 0 Then danger = "danger" Else danger = ""
            If (aux And AUXBITS_THREAT) <> 0 Then threat = "threat" Else threat = ""
            If TILE_HAS_DECAL(psTile) Then decal = "y" Else decal = "n"
            console "%s tile %d, %d [%d, %d] continent(l%d, h%d) level %g illum %d ao %d col %x %s %s w=%d s=%d j=%d tile#%d (decal=%s, ground [#%d, size=%.3f], f%d r%d)", _
                    explored, _
                    mouseTileX, mouseTileY, world_coord(mouseTileX), world_coord(mouseTileY), _
                    psMapTiles(psTile).limitedContinent, psMapTiles(psTile).hoverContinent, CDbl(psMapTiles(psTile).level), Fix(psMapTiles(psTile).illumination), _
                    Fix(psMapTiles(psTile).ambientOcclusion), LightMap_get_const(mouseTileX, mouseTileY), _
                    danger, threat, _
                    psMapTiles(psTile).watchers(selectedPlayer), psMapTiles(psTile).sensors(selectedPlayer), psMapTiles(psTile).jammers(selectedPlayer), _
                    TileNumber_tile(psMapTiles(psTile).texture), decal, _
                    psMapTiles(psTile).ground, CDbl(getGroundType(psMapTiles(psTile).ground).textureSize), _
                    flipVal, SHRU(TileNumber_texture(psMapTiles(psTile).texture) And TILE_ROTMASK, TILE_ROTSHIFT)
        End If
    End If
End Sub

Public Function getRotActive() As Boolean
    getRotActive = rotActive
End Function

' process LMB double clicks
Private Sub dealWithLMBDClick()
    Dim psClickedOn As Long, psDroid As Long, psStructure As Long

    ' What have we clicked on?
    psClickedOn = mouseTarget()
    ' If not NULL, then it's a droid or a structure
    If psClickedOn <> 0 Then
        ' We've got a droid or a structure
        If SIMPLE_OBJECT_type(psClickedOn) = OBJ_DROID Then
            ' We clicked on droid
            psDroid = psClickedOn
            If aDroid(psDroid).player = selectedPlayer Then
                ' Now selects all of same type on screen
                selDroidSelection selectedPlayer, DS_BY_TYPE, DST_ALL_SAME, True
            End If
        ElseIf SIMPLE_OBJECT_type(psClickedOn) = OBJ_STRUCTURE Then
            ' We clicked on structure
            psStructure = psClickedOn
            If aStruct(psStructure).player = selectedPlayer And Not STRUCTURE_isBlueprint(psStructure) Then
                If STRUCTURE_isFactory(psStructure) Then
                    setViewPos map_coord(aFlagPos(aStruct(psStructure).functionality.factory.psAssemblyPoint).coords.x), _
                               map_coord(aFlagPos(aStruct(psStructure).functionality.factory.psAssemblyPoint).coords.y), _
                               True
                ElseIf asStructureStats(aStruct(psStructure).pStructureType).typ = REF_REPAIR_FACILITY Then
                    setViewPos map_coord(aFlagPos(aStruct(psStructure).functionality.repairFacility.psDeliveryPoint).coords.x), _
                               map_coord(aFlagPos(aStruct(psStructure).functionality.repairFacility.psDeliveryPoint).coords.y), _
                               True
                End If
            End If
        End If
    End If
End Sub

'*
' * Find a delivery point, owned by `selectedPlayer`, pointed by the mouse.
Private Function findMouseDeliveryPoint() As Long
    Dim it As Long, psPoint As Long, dispX As Long, dispY As Long, dispR As Long
    If selectedPlayer >= MAX_PLAYERS Then
        findMouseDeliveryPoint = 0
        Exit Function
    End If

    it = apsFlagPosLists(selectedPlayer).head
    Do While it <> 0
        psPoint = lnObj(it)
        If aFlagPos(psPoint).typ <> POS_DELIVERY Then
            GoTo continue_
        End If

        dispX = aFlagPos(psPoint).screenX
        dispY = aFlagPos(psPoint).screenY
        dispR = aFlagPos(psPoint).screenR
        If DrawnInLastFrame(aFlagPos(psPoint).frameNumber) = True Then ' Only check DP's that are on screen
            If mouseInBox(dispX - dispR, dispY - dispR, dispX + dispR, dispY + dispR) Then
                ' We HAVE clicked on DP!
                findMouseDeliveryPoint = psPoint
                Exit Function
            End If
        End If
continue_:
        it = lnNext(it)
    Loop

    findMouseDeliveryPoint = 0
End Function

Private Sub dealWithRMB()
    Dim psClickedOn As Long, psDroid As Long, psStructure As Long, deliveryPoint As Long

    If isMouseOverRadar() Or InGameOpUp = True Then
        Exit Sub
    End If
    If Not widgGetFromID(psWScreen, INTINGAMEOP) Is Nothing Then
        Exit Sub
    End If

    ' What have we clicked on?
    psClickedOn = mouseTarget()
    ' If not NULL, then it's a droid or a structure
    If psClickedOn <> 0 Then
        ' We've got a droid or a structure
        If SIMPLE_OBJECT_type(psClickedOn) = OBJ_DROID Then
            ' We clicked on droid
            psDroid = psClickedOn
            If aDroid(psDroid).player = selectedPlayer Then
                If bRightClickOrders And ctrlShiftDown() Then
                    dealWithDroidSelect psDroid, False
                ' Not a transporter
                ElseIf Not DROID_isTransporter(psDroid) Then
                    If bRightClickOrders Then
                        ' We've clicked on one of our own droids
                        printDroidClickInfo psDroid
                    Else
                        If Not aDroid(psDroid).selected Then
                            clearSelection
                            SelectDroid psDroid
                        End If
                        intObjectSelected psDroid
                    End If
                ' Transporter
                Else
                    If bMultiPlayer Then
                        If bRightClickOrders And Not aDroid(psDroid).selected Then
                            clearSelection
                            SelectDroid psDroid
                        Else
                            intResetScreen False
                            If Not getWidgetsStatus() Then
                                setWidgetsStatus True
                            End If
                            addTransporterInterface psDroid, False
                        End If
                    End If
                End If
            ElseIf bMultiPlayer And isHumanPlayer(aDroid(psDroid).player) Then
                console "%s", droidGetName(psDroid)
                FeedbackOrderGiven
            End If
        ' end if its a droid
        ElseIf SIMPLE_OBJECT_type(psClickedOn) = OBJ_STRUCTURE Then
            ' We clicked on structure
            psStructure = psClickedOn
            If aStruct(psStructure).player = selectedPlayer Then
                ' We've clicked on our own building
                If bRightClickOrders And intDemolishSelectMode() Then
                    orderSelectedObjAdd selectedPlayer, psClickedOn, ctrlShiftDown()
                    FeedbackOrderGiven
                    ' we were demolishing something - now we're done
                    If ctrlShiftDown() Then
                        quickQueueMode = True
                    Else
                        intDemolishCancel
                    End If
                ElseIf aStruct(psStructure).selected Then
                    aStruct(psStructure).selected = False
                    intObjectSelected 0
                    triggerEventSelected
                    jsDebugSelected psStructure
                ElseIf Not STRUCTURE_isBlueprint(psStructure) Then
                    clearSelection

                    If bRightClickOrders Then
                        If (aStruct(psStructure).status = SS_BUILT) And _
                           (asStructureStats(aStruct(psStructure).pStructureType).typ <> REF_RESOURCE_EXTRACTOR) Then
                            printStructureInfo psStructure

                            aStruct(psStructure).selected = True
                            jsDebugSelected psStructure

                            ' Open structure menu
                            intObjectSelected psStructure
                            FeedbackOrderGiven

                            bLasSatStruct = lasSatStructSelected(psStructure)
                            triggerEventSelected
                        End If
                    ElseIf STRUCTURE_isFactory(psStructure) Then
                        'pop up the order interface for the factory
                        intAddFactoryOrder psStructure
                    Else
                        intObjectSelected psStructure
                    End If
                End If
            End If
        ' end if its a structure
        Else
            ' And if it's not a feature, then we're in trouble!
            If Not (SIMPLE_OBJECT_type(psClickedOn) = OBJ_FEATURE) Then ASSERT_FAILURE "psClickedOn->type == OBJ_FEATURE", "src/display.cpp:2331", "dealWithRMB", "Weird selection from RMB - type of clicked object is %d", SIMPLE_OBJECT_type(psClickedOn)
        End If
    Else
        deliveryPoint = findMouseDeliveryPoint()
        If deliveryPoint <> 0 Then
            If bRightClickOrders Then
                startDeliveryPosition deliveryPoint
            Else
                'centre the view on the owning Factory
                psStructure = findDeliveryFactory(deliveryPoint)
                If psStructure <> 0 Then
                    setViewPos map_coord(aStruct(psStructure).pos.x), map_coord(aStruct(psStructure).pos.y), True
                End If
            End If
        Else
            handleDeselectionClick
        End If
    End If
End Sub

' if there is a valid object under the mouse this routine returns not only the type of the object in the
'return code, but also a pointer to the BASE_OBJECT) ... well if your going to be "object orientated" you might as well do it right
'- it sets it to null if we don't find anything
Private Function itemUnderMouse(ppObjectUnderMouse As Long) As MOUSE_TARGET
    Dim i As Long, retVal As MOUSE_TARGET, psNotDroid As Long, dispX As Long, dispY As Long, dispR As Long
    Dim psStructure As Long, it As Long, psDroid As Long, allied As Boolean, sensorTurret As Boolean

    ppObjectUnderMouse = 0

    If mouseTileX < 0 Or mouseTileY < 0 Or mouseTileX > (mapWidth - 1) Or mouseTileY > (mapHeight - 1) Then
        retVal = MT_BLOCKING
        itemUnderMouse = retVal
        Exit Function
    End If

    ' We haven't found anything yet
    retVal = MT_NOTARGET

    ' First have a look through the droid lists
    For i = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            dispX = aDroid(psDroid).sDisplay.screenX
            dispY = aDroid(psDroid).sDisplay.screenY
            dispR = aDroid(psDroid).sDisplay.screenR
            ' Only check droids that're on screen
            If aDroid(psDroid).sDisplay.frameNumber + 1 = currentFrame Then
                If BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 Then
                    If mouseInBox(dispX - dispR, dispY - dispR, dispX + dispR, dispY + dispR) Then
                        ' We HAVE clicked on droid!
                        allied = False
                        If selectedPlayer < MAX_PLAYERS Then allied = aiCheckAlliances(aDroid(psDroid).player, selectedPlayer)
                        If allied Then
                            ppObjectUnderMouse = psDroid
                            ' need to check for command droids here as well
                            If aDroid(psDroid).droidType = DROID_SENSOR Then
                                If selectedPlayer <> aDroid(psDroid).player Then
                                    retVal = MT_CONSTRUCT ' Can't assign to allied units
                                Else
                                    retVal = MT_SENSOR
                                End If
                            ElseIf DROID_isTransporter(psDroid) And _
                                   selectedPlayer = aDroid(psDroid).player Then
                                'check the transporter is not full
                                If calcRemainingCapacity(psDroid) <> 0 Then
                                    retVal = MT_TRANDROID
                                Else
                                    retVal = MT_BLOCKING
                                End If
                            ElseIf aDroid(psDroid).droidType = DROID_CONSTRUCT Or _
                                   aDroid(psDroid).droidType = DROID_CYBORG_CONSTRUCT Then
                                itemUnderMouse = MT_CONSTRUCT
                                Exit Function
                            ElseIf aDroid(psDroid).droidType = DROID_COMMAND Then
                                If selectedPlayer <> aDroid(psDroid).player Then
                                    retVal = MT_CONSTRUCT ' Can't assign to allied units
                                Else
                                    retVal = MT_COMMAND
                                End If
                            Else
                                If DROID_isDamaged(psDroid) Then
                                    retVal = MT_OWNDROIDDAM
                                Else
                                    retVal = MT_OWNDROID
                                End If
                            End If
                        Else
                            ppObjectUnderMouse = psDroid
                            retVal = MT_ENEMYDROID
                        End If
                        ' There's no point in checking other object types
                        itemUnderMouse = retVal
                        Exit Function
                    End If
                End If
            End If
            it = lnNext(it)
        Loop
    Next ' end of checking for droids

    '    Not a droid, so maybe a structure or feature?
    '    If still NULL after this then nothing
    psNotDroid = getTileOccupier(mouseTileX, mouseTileY)
    If psNotDroid = 0 Then
        psNotDroid = getTileBlueprintStructure(mouseTileX, mouseTileY)
    End If

    If psNotDroid <> 0 Then
        ppObjectUnderMouse = psNotDroid

        If SIMPLE_OBJECT_type(psNotDroid) = OBJ_FEATURE Then
            If (asFeatureStats(aFeature(psNotDroid).psStats).subType = FEAT_GEN_ARTE) _
               Or (asFeatureStats(aFeature(psNotDroid).psStats).subType = FEAT_OIL_DRUM) Then
                retVal = MT_ARTIFACT
            ElseIf asFeatureStats(aFeature(psNotDroid).psStats).damageable Then 'make damageable features return 'target' mouse pointer
                retVal = MT_DAMFEATURE
            ElseIf asFeatureStats(aFeature(psNotDroid).psStats).subType = FEAT_OIL_RESOURCE Then
                retVal = MT_RESOURCE
            Else
                retVal = MT_BLOCKING
            End If
        ElseIf SIMPLE_OBJECT_type(psNotDroid) = OBJ_STRUCTURE Then
            psStructure = psNotDroid

            allied = False
            If selectedPlayer < MAX_PLAYERS Then allied = aiCheckAlliances(SIMPLE_OBJECT_player(psNotDroid), selectedPlayer)
            If allied Then
                sensorTurret = False
                If asStructureStats(aStruct(psStructure).pStructureType).pSensor <> -1 Then
                    sensorTurret = (asSensorStats(asStructureStats(aStruct(psStructure).pStructureType).pSensor).location = LOC_TURRET)
                End If
                If aStruct(psStructure).status = SS_BEING_BUILT Or isBlueprint(psStructure) Then
                    retVal = MT_OWNSTRINCOMP
                ' repair center.
                ElseIf asStructureStats(aStruct(psStructure).pStructureType).typ = REF_REPAIR_FACILITY Then
                    If buildingDamaged(psStructure) Then
                        retVal = MT_REPAIRDAM
                    Else
                        retVal = MT_REPAIR
                    End If
                'sensor tower
                ElseIf sensorTurret Then
                    If buildingDamaged(psStructure) Then
                        retVal = MT_SENSORSTRUCTDAM
                    Else
                        retVal = MT_SENSORSTRUCT
                    End If

                ' standard buildings. - check for buildingDamaged BEFORE upgrades
                ElseIf buildingDamaged(psStructure) Then
                    retVal = MT_OWNSTRDAM
                ' If this building is a factory/power generator/research facility
                ' which isn't upgraded. Make the build icon available.
                ElseIf nextModuleToBuild(psStructure, -1) > 0 Then
                    retVal = MT_OWNSTRINCOMP
                Else
                    ' All the different stages of construction
                    retVal = MT_OWNSTROK
                End If
            Else
                retVal = MT_ENEMYSTR    ' enemy structure
            End If
        End If
    End If

    ' Send the result back - if it's null then we clicked on an area of terrain
    ' make unseen objects just look like terrain.
    If retVal = MT_NOTARGET Then
        retVal = MT_TERRAIN
    ElseIf psNotDroid <> 0 Then
        If BASE_OBJECT_visibleForLocalDisplay(psNotDroid) = 0 Then retVal = MT_TERRAIN
    End If
    itemUnderMouse = retVal
End Function

' Only deals with one type of droid being selected!!!!
'    We'll have to make it assess which selection is to be dominant in the case
'    of multiple selections
Private Function establishSelection(ByVal selectedPlayer_ As Long) As SELECTION_TYPE
    Dim psDominant As Long, CurrWeight As Long, selectionClass As SELECTION_TYPE, it As Long, psDroid As Long
    psDominant = 0
    CurrWeight = UBYTE_MAX
    selectionClass = SC_INVALID

    If selectedPlayer_ >= MAX_PLAYERS Then
        establishSelection = SC_INVALID
        Exit Function
    End If

    it = apsDroidLists(selectedPlayer_).head
    Do While it <> 0
        psDroid = lnObj(it)
        ' This works, uses the DroidSelectionWeights[] table to priorities the different
        ' droid types and find the dominant selection.
        If aDroid(psDroid).selected Then
            If Not (aDroid(psDroid).droidType < NUM_DROID_WEIGHTS) Then ASSERT_FAILURE "psDroid->droidType < NUM_DROID_WEIGHTS", "src/display.cpp:2610", "establishSelection", "droidType exceeds NUM_DROID_WEIGHTS": establishSelection = SC_INVALID: Exit Function
            If DroidSelectionWeights(aDroid(psDroid).droidType) < CurrWeight Then
                CurrWeight = DroidSelectionWeights(aDroid(psDroid).droidType)
                psDominant = psDroid
            End If
        End If
        it = lnNext(it)
    Loop

    If psDominant <> 0 Then
        psDominantSelected = psDominant
        Select Case aDroid(psDominant).droidType
        Case DROID_WEAPON
            If proj_Direct(DROID_getWeaponStats(psDominant, 0)) Then
                selectionClass = SC_DROID_DIRECT
            Else
                selectionClass = SC_DROID_INDIRECT
            End If

        Case DROID_PERSON
            selectionClass = SC_DROID_DIRECT
        Case DROID_CYBORG, DROID_CYBORG_SUPER
            selectionClass = SC_DROID_DIRECT
        Case DROID_TRANSPORTER, DROID_SUPERTRANSPORTER
            'can remove this is NEVER going to select the Transporter to move
            'Never say Never!! cos here we go in multiPlayer!!
            selectionClass = SC_DROID_TRANSPORTER
        Case DROID_SENSOR
            selectionClass = SC_DROID_SENSOR

        Case DROID_ECM
            selectionClass = SC_DROID_ECM

        Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
            If intDemolishSelectMode() Then
                establishSelection = SC_DROID_DEMOLISH
                Exit Function
            End If
            selectionClass = SC_DROID_CONSTRUCT

        Case DROID_COMMAND
            selectionClass = SC_DROID_COMMAND

        Case DROID_REPAIR, DROID_CYBORG_REPAIR
            selectionClass = SC_DROID_REPAIR

        Case Else
            ASSERT_FAILURE "!""unknown droid type""", "src/display.cpp:2675", "establishSelection", "Weirdy droid type on what you've clicked on!!!"
        End Select
    End If
    establishSelection = selectionClass
End Function

' Just returns true if the building's present body points aren't 100 percent
Private Function buildingDamaged(ByVal psStructure As Long) As Boolean
    buildingDamaged = U32(aStruct(psStructure).body) < U32(STRUCTURE_structureBody(psStructure))
End Function

'Looks through the list of selected players droids to see if one is a repair droid
Public Function repairDroidSelected(ByVal player As Long) As Boolean
    Dim it As Long, psCurr As Long
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/display.cpp:2692", "repairDroidSelected", "Invalid player (%u)", player: Exit Function

    it = apsDroidLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).selected And ( _
           aDroid(psCurr).droidType = DROID_REPAIR Or _
           aDroid(psCurr).droidType = DROID_CYBORG_REPAIR) Then
            repairDroidSelected = True
            Exit Function
        End If
        it = lnNext(it)
    Loop

    'didn't find one...
    repairDroidSelected = False
End Function

'Looks through the list of selected players droids to see if one is a VTOL droid
Public Function vtolDroidSelected(ByVal player As Long) As Boolean
    Dim it As Long, psCurr As Long
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/display.cpp:2711", "vtolDroidSelected", "player: %u", player: Exit Function

    it = apsDroidLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).selected Then
            If DROID_isVtol(psCurr) Then
                ' horrible hack to note one of the selected vtols
                psSelectedVtol = psCurr
                vtolDroidSelected = True
                Exit Function
            End If
        End If
        it = lnNext(it)
    Loop

    'didn't find one...
    vtolDroidSelected = False
End Function

'Looks through the list of selected players droids to see if any is selected
Public Function anyDroidSelected(ByVal player As Long) As Boolean
    Dim it As Long
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/display.cpp:2730", "anyDroidSelected", "Invalid player (%u)", player: Exit Function

    it = apsDroidLists(player).head
    Do While it <> 0
        If aDroid(lnObj(it)).selected Then
            anyDroidSelected = True
            Exit Function
        End If
        it = lnNext(it)
    Loop

    'didn't find one...
    anyDroidSelected = False
End Function

'Looks through the list of selected players droids to see if one is a cyborg droid
Public Function cyborgDroidSelected(ByVal player As Long) As Boolean
    Dim it As Long, psCurr As Long
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/display.cpp:2747", "cyborgDroidSelected", "Invalid player (%u)", player: Exit Function

    it = apsDroidLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).selected Then
            If DROID_isCyborg(psCurr) Then
                cyborgDroidSelected = True
                Exit Function
            End If
        End If
        it = lnNext(it)
    Loop

    'didn't find one...
    cyborgDroidSelected = False
End Function

' Clear the selection flag for a player
Public Sub clearSelection()
    Dim it As Long
    DROIDDOING = ""   ' clear string when deselected

    If selectedPlayer >= MAX_PLAYERS Then
        Exit Sub
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        aDroid(lnObj(it)).selected = False
        it = lnNext(it)
    Loop
    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        aStruct(lnObj(it)).selected = False
        it = lnNext(it)
    Loop
    bLasSatStruct = False
    'clear the Deliv Point if one
    it = apsFlagPosLists(selectedPlayer).head
    Do While it <> 0
        aFlagPos(lnObj(it)).selected = False
        it = lnNext(it)
    Loop

    intRefreshScreen

    triggerEventSelected
End Sub

Private Sub handleDeselectionClick()
    clearSelection
    intObjectSelected 0
End Sub

'access function for bSensorAssigned variable
Public Sub setSensorAssigned()
    bSensorAssigned = True
End Sub

' Initialise the display system
Public Function dispInitialise() As Boolean
    flagReposVarsValid = False
    gInputManager_contexts_resetStates
    dispInitialise = True
End Function
