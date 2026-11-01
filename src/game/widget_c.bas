Attribute VB_Name = "widget_c"
' lib/widget/widget.cpp, widget.h and widgbase.h's rest: the widget screen's
' run and display, the widg* API, the overlay screens, the init structs, the
' WIDGET class's virtual methods' defaults (WIDGET_<method>, the C++
' WIDGET::<method> bodies) and WidgetGraphicsContext (a UDT and its
' procedures). The object model: docs/CONVENTIONS.md, "Widgets" (WIDGET.cls,
' W_SCREEN.cls, WzFunc.cls; widget_classes_vb is generated). A default that
' needs WIDGET's private members calls its Friend def_<method>.
' (VB) A W_CONTEXT argument of a WzFunc call (WIDGET_CALLBACK) is its four
' fields: xOffset, yOffset, mx, my (a module's UDT cannot be a Variant).
' std::chrono::steady_clock is the performance counter (widgSteadyNowMs).
' widgRunScreen returns the number of triggers (psScreen.retWidgetAt(i)).
'
' (VB) widget class WIDGET
' (VB) virtual WIDGET.widgetLost
' (VB) virtual WIDGET.focusLost
' (VB) virtual WIDGET.clicked
' (VB) virtual WIDGET.getTip
' (VB) virtual WIDGET.getHelp
' (VB) virtual WIDGET.capturesMouseDrag
' (VB) virtual WIDGET.released
' (VB) virtual WIDGET.highlight
' (VB) virtual WIDGET.highlightLost
' (VB) virtual WIDGET.run
' (VB) virtual WIDGET.display
' (VB) virtual WIDGET.geometryChanged
' (VB) virtual WIDGET.hitTest
' (VB) virtual WIDGET.mouseDragged
' (VB) virtual WIDGET.getState
' (VB) virtual WIDGET.setState
' (VB) virtual WIDGET.setFlash
' (VB) virtual WIDGET.getString
' (VB) virtual WIDGET.setString
' (VB) virtual WIDGET.setTip
' (VB) virtual WIDGET.setHelp
' (VB) virtual WIDGET.screenSizeDidChange
' (VB) virtual WIDGET.parentRelativeXOffset
' (VB) virtual WIDGET.parentRelativeYOffset
' (VB) virtual WIDGET.setGeometryFromScreenRect
' (VB) virtual WIDGET.idealWidth
' (VB) virtual WIDGET.idealHeight
' (VB) virtual WIDGET.getScrollSnapOffsets
' (VB) virtual WIDGET.findMouseTargetRecursive
' (VB) virtual WIDGET.runRecursive
' (VB) virtual WIDGET.displayRecursive
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'***********************************************************************************
'
' Widget style definitions - these control how the basic widget appears on screen

'************ Form styles ****************

Public Const WFORM_PLAIN As Long = 0        ' Plain form

' Invisible (i.e. see through) form - can be used in conjunction with WFORM_PLAIN.
Public Const WFORM_INVISIBLE As Long = 2

Public Const WFORM_CLICKABLE As Long = 4    ' Clickable form - return form id when the form is clicked
Public Const WFORM_NOCLICKMOVE As Long = 8  ' Disable movement on a clickable form

' Control whether the primary or secondary buttons work on a clickable form.
' Primary works by default - this turns it off.
Public Const WFORM_NOPRIMARY As Long = &H10
Public Const WFORM_SECONDARY As Long = &H20 ' Enable secondary buttons

'************ Label styles ***************

Public Const WLAB_PLAIN As Long = 0         ' Plain text only label
Public Enum WzTextAlignment
    WLAB_ALIGNLEFT = 0                      ' Align the text at the left of the box
    WLAB_ALIGNCENTRE = 1                    ' Centre the text
    WLAB_ALIGNRIGHT = 2                     ' Align the text at the right of the box
    WLAB_ALIGNTOPLEFT = 4                   ' Align the text at the top left of the box
    WLAB_ALIGNTOP = 4 + 1                   ' Align the text at the top of the box
    WLAB_ALIGNTOPRIGHT = 4 + 2              ' Align the text at the top right of the box
    WLAB_ALIGNBOTTOMLEFT = 8                ' Align the text at the bottom left of the box
    WLAB_ALIGNBOTTOM = 8 + 1                ' Align the text at the bottom of the box
    WLAB_ALIGNBOTTOMRIGHT = 8 + 2           ' Align the text at the bottom rightof the box
End Enum

'************ Button styles **************

Public Const WBUT_PLAIN As Long = 0         ' Plain button (text with a box around it)

' Control whether the primary or secondary buttons work on a button. Primary works by default -
' this turns it off.
Public Const WBUT_NOPRIMARY As Long = &H10
Public Const WBUT_SECONDARY As Long = &H20  ' Enable secondary buttons.
Public Const WBUT_TXTCENTRE As Long = &H40  ' Text only buttons. centre the text?

'*********** Edit Box styles *************

Public Const WEDB_PLAIN As Long = 0         ' Plain edit box (text with a box around it)

'*********** Bar Graph styles ************

Public Const WBAR_PLAIN As Long = 0         ' Plain bar graph
Public Const WBAR_TROUGH As Long = 1        ' Bar graph with a trough showing empty percentage
Public Const WBAR_DOUBLE As Long = 2        ' Double bar graph, one on top of other

'*********** Slider styles ***************

Public Const WSLD_PLAIN As Long = 0         ' Plain slider

' The maximum value for bar graph size
Public Const WBAR_SCALE As Long = 100

Public Type Padding
    top As Long                             ' uint32_t = 0
    right_ As Long
    bottom As Long
    left_ As Long
End Type

' The basic initialisation structure (W_INIT::W_INIT(): all zero, every function nullptr: a VB UDT
' starts so; W_INIT_construct for uniformity)
Public Type W_INIT
    formID As Long                          ' UDWORD: ID number of form to put widget on. ID == 0 specifies the default form for the screen
    majorID As Long                         ' UWORD: Which tab to put the widget on for a tabbed form
    id As Long                              ' UDWORD: Unique id number (chosen by user)
    style As Long                           ' UDWORD: widget style
    x As Long                               ' SWORD: screen location
    y As Long
    width As Long                           ' UWORD: widget size
    height As Long
    pDisplay As WzFunc                      ' WIDGET_DISPLAY: Optional display function
    pCallback As WzFunc                     ' WIDGET_CALLBACK: Optional callback function
    pUserData As Variant                    ' void *: Optional user data pointer
    UserData As Long                        ' UDWORD: User data (if any)
    calcLayout As WzFunc                    ' WIDGET_CALCLAYOUT_FUNC: Optional calculate layout callback function
    onDelete As WzFunc                      ' WIDGET_ONDELETE_FUNC: Optional callback called when the Widget is about to be deleted
    customHitTest As WzFunc                 ' WIDGET_HITTEST_FUNC: Optional custom hit-testing function
    initPUserDataFunc As WzFunc             ' WIDGET_INITIALIZE_PUSERDATA_FUNC: (Optional) Used to initialize the pUserData pointer per widget instance
End Type

' Form initialisation structure (W_FORMINIT : W_INIT: W_INIT's fields first, then its own)
Public Type W_FORMINIT
    formID As Long
    majorID As Long
    id As Long
    style As Long
    x As Long
    y As Long
    width As Long
    height As Long
    pDisplay As WzFunc
    pCallback As WzFunc
    pUserData As Variant
    UserData As Long
    calcLayout As WzFunc
    onDelete As WzFunc
    customHitTest As WzFunc
    initPUserDataFunc As WzFunc
    disableChildren As Boolean
    pTip As String                          ' Tool tip for the form itself
End Type

' Label initialisation structure
Public Type W_LABINIT
    formID As Long
    majorID As Long
    id As Long
    style As Long
    x As Long
    y As Long
    width As Long
    height As Long
    pDisplay As WzFunc
    pCallback As WzFunc
    pUserData As Variant
    UserData As Long
    calcLayout As WzFunc
    onDelete As WzFunc
    customHitTest As WzFunc
    initPUserDataFunc As WzFunc
    pText As String                         ' label text
    pTip As String                          ' Tool tip for the label.
    FontID As iV_fonts                      ' ID of the IVIS font to use for this widget.
End Type

' Button initialisation structure
Public Type W_BUTINIT
    formID As Long
    majorID As Long
    id As Long
    style As Long
    x As Long
    y As Long
    width As Long
    height As Long
    pDisplay As WzFunc
    pCallback As WzFunc
    pUserData As Variant
    UserData As Long
    calcLayout As WzFunc
    onDelete As WzFunc
    customHitTest As WzFunc
    initPUserDataFunc As WzFunc
    pText As String                         ' const char *: Button text (vbNullString = nullptr)
    pTip As String                          ' Tool tip text
    FontID As iV_fonts                      ' ID of the IVIS font to use for this widget.
End Type

' Edit box initialisation structure
Public Type W_EDBINIT
    formID As Long
    majorID As Long
    id As Long
    style As Long
    x As Long
    y As Long
    width As Long
    height As Long
    pDisplay As WzFunc
    pCallback As WzFunc
    pUserData As Variant
    UserData As Long
    calcLayout As WzFunc
    onDelete As WzFunc
    customHitTest As WzFunc
    initPUserDataFunc As WzFunc
    pText As String                         ' const char *: initial contents of the edit box
    FontID As iV_fonts                      ' ID of the IVIS font to use for this widget.
    pBoxDisplay As WzFunc                   ' WIDGET_DISPLAY: Optional callback to display the form.
End Type

' Orientation flags for the bar graph
Public Enum WBAR_ORIENTATION
    WBAR_LEFT = 1                           ' Bar graph fills from left to right
    WBAR_RIGHT                              ' Bar graph fills from right to left
    WBAR_TOP                                ' Bar graph fills from top to bottom
    WBAR_BOTTOM                             ' Bar graph fills from bottom to top
End Enum

' Bar Graph initialisation structure
Public Type W_BARINIT
    formID As Long
    majorID As Long
    id As Long
    style As Long
    x As Long
    y As Long
    width As Long
    height As Long
    pDisplay As WzFunc
    pCallback As WzFunc
    pUserData As Variant
    UserData As Long
    calcLayout As WzFunc
    onDelete As WzFunc
    customHitTest As WzFunc
    initPUserDataFunc As WzFunc
    orientation As WBAR_ORIENTATION         ' Orientation of the bar on the widget
    size As Long                            ' UWORD: Initial percentage of the graph that is filled
    minorSize As Long                       ' UWORD: Percentage of second bar graph if there is one
    iRange As Long                          ' UWORD: Maximum range
    denominator As Long                     ' Denominator, 1 by default.
    precision As Long                       ' Number of places after the decimal point to display, 0 by default.
    sCol As Long                            ' PIELIGHT: Bar colour
    sMinorCol As Long                       ' PIELIGHT: Minor bar colour
    pTip As String                          ' Tool tip text
End Type

' Orientation of the slider
Public Enum WSLD_ORIENTATION
    WSLD_LEFT = 1                           ' Slider is horizontal and starts at left
    WSLD_RIGHT                              ' Slider is horizontal and starts at the right
    WSLD_TOP                                ' Slider is vertical and starts at the top
    WSLD_BOTTOM                             ' Slider is vertical and starts at the bottom
End Enum

' Slider initialisation structure
Public Type W_SLDINIT
    formID As Long
    majorID As Long
    id As Long
    style As Long
    x As Long
    y As Long
    width As Long
    height As Long
    pDisplay As WzFunc
    pCallback As WzFunc
    pUserData As Variant
    UserData As Long
    calcLayout As WzFunc
    onDelete As WzFunc
    customHitTest As WzFunc
    initPUserDataFunc As WzFunc
    orientation As WSLD_ORIENTATION         ' Orientation of the slider
    numStops As Long                        ' UWORD: Number of stops on the slider
    barSize As Long                         ' UWORD: Size of the bar
    pos As Long                             ' UWORD: Initial position of the slider bar
    pTip As String                          ' Tip string
End Type

' Colour numbers (enum _w_colour)
Public Enum w_colour_
    WCOL_BKGRND                             ' Background colours
    WCOL_TEXT                               ' Text colour
    WCOL_LIGHT                              ' Light colour for 3D effects
    WCOL_DARK                               ' Dark colour for 3D effects
    WCOL_HILITE                             ' Hilite colour
    WCOL_CURSOR                             ' Edit Box cursor colour
    WCOL_TIPBKGRND                          ' Background for the tool tip window
    WCOL_DISABLE                            ' Text colour on a disabled button

    WCOL_MAX                                ' All colour numbers are less than this
End Enum

' widgbase.h: WidgetHelp::InteractionTriggers (a WidgetInteractionTriggerFlagSet is a mask of
' 2 ^ trigger: WidgetHelp's InteractionTriggers_flag)
Public Enum InteractionTriggers
    InteractionTriggers_PrimaryClick
    InteractionTriggers_SecondaryClick
    InteractionTriggers_ClickAndHold
    InteractionTriggers_Misc
End Enum
Public Const NUM_WIDGET_INTERACTION_TRIGGERS As Long = 4

' widgbase.h: struct WidgetTrigger (WidgetTriggers: an array of them; widgRunScreen's)
Public Type WidgetTrigger
    widget As WIDGET                        ' std::shared_ptr<WIDGET>
End Type

' widgbase.h: class WidgetGraphicsContext (its methods: WidgetGraphicsContext_<method>(c, ...))
Public Type WidgetGraphicsContext
    offset As Vector2i                      ' = {0, 0}
    clipRect As WzRect                      ' = {0, 0, 0, 0}
    clipped As Boolean                      ' = false
    allowChildDisplayIfSelfClipped As Boolean ' = false
End Type

Private Type WIDGET_KEYSTATE
    pressed As Boolean                      ' = false
    dragStartPos As W_CONTEXT               ' = W_CONTEXT::ZeroContext()
    dragLastPos As W_CONTEXT
    capturedDragWidget As WIDGET
End Type

Private bWidgetsInitialized As Boolean      ' = false
Private bWidgetsActive As Boolean           ' = true (widget_c_init)

' The widget the mouse is over this update
Public psMouseOverWidgetPtr As Long         ' (VB) std::weak_ptr<WIDGET> psMouseOverWidget: its address (WIDGET's Class_Terminate clears it)
Private psClickDownWidgetScreen As W_SCREEN
Private psMouseOverWidgetScreen As W_SCREEN

Private widgetKeyCurrentState(0 To WKEY_SECONDARY) As WIDGET_KEYSTATE

Private AudioCallback As WzFunc             ' WIDGET_AUDIOCALLBACK = nullptr
Private HilightAudioID As Long              ' SWORD = -1 (widget_c_init)
Private ClickedAudioID As Long
Private ErrorAudioID As Long

Private lastReleasedKey_DEPRECATED As WIDGET_KEY ' = WKEY_NONE

Private widgetDeletionQueue() As WIDGET     ' std::deque<std::shared_ptr<WIDGET>>
Private widgetDeletionQueueCount As Long
Private widgetScheduledTasks() As WzFunc    ' std::vector<std::function<void()>>
Private widgetScheduledTasksCount As Long

Private debugBoundingBoxesOnly As Boolean   ' = false

' struct OverlayScreen {psScreen, zOrder}: the overlays, in decreasing z-order
Private overlayScreens() As W_SCREEN
Private overlayZOrders() As Long            ' uint16_t
Private overlayCount As Long
Private overlaySet As Object                ' std::unordered_set<std::shared_ptr<W_SCREEN>> (for quick lookup): a Scripting.Dictionary, key ObjPtr
Private overlaysToDelete As Object
Private initDone As Boolean

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub widget_c_init()
    bWidgetsActive = True
    HilightAudioID = -1
    ClickedAudioID = -1
    ErrorAudioID = -1
    lastReleasedKey_DEPRECATED = WKEY_NONE
    ensureSets
End Sub

Private Sub ensureSets()
    If initDone Then Exit Sub
    initDone = True
    Set overlaySet = CreateObject("Scripting.Dictionary")
    Set overlaysToDelete = CreateObject("Scripting.Dictionary")
End Sub

' (VB) std::chrono::steady_clock::now(), in milliseconds (the performance counter)
Public Function widgSteadyNowMs() As Double
    Dim c As Currency, f As Currency
    QueryPerformanceCounter c
    QueryPerformanceFrequency f
    If f = 0 Then Exit Function
    widgSteadyNowMs = CDbl(c) * 1000# / CDbl(f)
End Function

' (VB) W_INIT::W_INIT() and the derived init structs' constructors
Public Sub W_INIT_construct(i As W_INIT)
    Dim z As W_INIT
    i = z
End Sub

Private Sub runScheduledTasks()
    Dim tasksCopy() As WzFunc, n As Long, i As Long
    n = widgetScheduledTasksCount
    If n = 0 Then Exit Sub
    tasksCopy = widgetScheduledTasks    ' auto tasksCopy = std::move(widgetScheduledTasks);
    Erase widgetScheduledTasks
    widgetScheduledTasksCount = 0
    For i = 0 To n - 1
        WzFunc_call tasksCopy(i)
    Next
End Sub

Private Sub deleteOldWidgets()
    Dim guiltyWidget As WIDGET, i As Long
    Do While widgetDeletionQueueCount > 0
        Set guiltyWidget = widgetDeletionQueue(0)
        widgDelete_widget guiltyWidget
        ' widgetDeletionQueue.pop_front()
        For i = 1 To widgetDeletionQueueCount - 1
            Set widgetDeletionQueue(i - 1) = widgetDeletionQueue(i)
        Next
        widgetDeletionQueueCount = widgetDeletionQueueCount - 1
        Set widgetDeletionQueue(widgetDeletionQueueCount) = Nothing
        Set guiltyWidget = Nothing
    Loop
End Sub

' (VB) WIDGET::deleteLater's queue: whether the widget is queued (std::find), and the queuing
Public Function widgetDeletionQueue_contains(ByVal w As WIDGET) As Boolean
    Dim i As Long
    For i = 0 To widgetDeletionQueueCount - 1
        If widgetDeletionQueue(i) Is w Then
            widgetDeletionQueue_contains = True
            Exit Function
        End If
    Next
End Function

Public Sub widgetDeletionQueue_push_back(ByVal w As WIDGET)
    If widgetDeletionQueueCount = 0 Then
        ReDim widgetDeletionQueue(0 To 7)
    ElseIf widgetDeletionQueueCount > UBound(widgetDeletionQueue) Then
        ReDim Preserve widgetDeletionQueue(0 To 2 * widgetDeletionQueueCount - 1)
    End If
    Set widgetDeletionQueue(widgetDeletionQueueCount) = w
    widgetDeletionQueueCount = widgetDeletionQueueCount + 1
End Sub

' (_widgDebugAssertIfRunningScreen: DEBUG only, nothing in a release build)

' Initialise the widget module
Public Function widgInitialise() As Boolean
    bWidgetsInitialized = True
    tipInitialise
    widgInitialise = True
End Function

' Reset the widgets.
'
Public Sub widgReset()
    tipInitialise
End Sub

' Shut down the widget module
Public Sub widgShutDown()
    Dim k As Long
    ensureSets
    Set psClickDownWidgetScreen = Nothing
    Set psMouseOverWidgetScreen = Nothing
    For k = 0 To WKEY_SECONDARY
        WIDGET_KEYSTATE_reset widgetKeyCurrentState(k)
    Next
    tipShutdown
    Erase overlayScreens
    Erase overlayZOrders
    overlayCount = 0
    overlaySet.RemoveAll
    overlaysToDelete.RemoveAll
    deleteOldWidgets
    Erase widgetScheduledTasks
    widgetScheduledTasksCount = 0
    ' (#ifdef DEBUG: the widgets still alive)
    bWidgetsInitialized = False
End Sub

Private Sub WIDGET_KEYSTATE_reset(s As WIDGET_KEYSTATE)
    Dim z As W_CONTEXT
    s.pressed = False
    s.dragStartPos = z
    s.dragLastPos = z
    Set s.capturedDragWidget = Nothing
End Sub

Public Sub widgScheduleTask(ByVal task As WzFunc)
    If widgetScheduledTasksCount = 0 Then
        ReDim widgetScheduledTasks(0 To 7)
    ElseIf widgetScheduledTasksCount > UBound(widgetScheduledTasks) Then
        ReDim Preserve widgetScheduledTasks(0 To 2 * widgetScheduledTasksCount - 1)
    End If
    Set widgetScheduledTasks(widgetScheduledTasksCount) = task
    widgetScheduledTasksCount = widgetScheduledTasksCount + 1
End Sub

' (VB) the index of a screen in the overlays, -1 if none
Private Function overlayIndexOf(ByVal psScreen As W_SCREEN) As Long
    Dim i As Long
    For i = 0 To overlayCount - 1
        If overlayScreens(i) Is psScreen Then
            overlayIndexOf = i
            Exit Function
        End If
    Next
    overlayIndexOf = -1
End Function

Private Sub overlayInsertAt(ByVal pos As Long, ByVal psScreen As W_SCREEN, ByVal zOrder As Long)
    Dim i As Long
    If overlayCount = 0 Then
        ReDim overlayScreens(0 To 7)
        ReDim overlayZOrders(0 To 7)
    ElseIf overlayCount > UBound(overlayScreens) Then
        ReDim Preserve overlayScreens(0 To 2 * overlayCount - 1)
        ReDim Preserve overlayZOrders(0 To 2 * overlayCount - 1)
    End If
    For i = overlayCount To pos + 1 Step -1
        Set overlayScreens(i) = overlayScreens(i - 1)
        overlayZOrders(i) = overlayZOrders(i - 1)
    Next
    Set overlayScreens(pos) = psScreen
    overlayZOrders(pos) = zOrder
    overlayCount = overlayCount + 1
End Sub

Private Sub overlayEraseAt(ByVal pos As Long)
    Dim i As Long
    For i = pos To overlayCount - 2
        Set overlayScreens(i) = overlayScreens(i + 1)
        overlayZOrders(i) = overlayZOrders(i + 1)
    Next
    overlayCount = overlayCount - 1
    Set overlayScreens(overlayCount) = Nothing
End Sub

Private Function screenKey(ByVal psScreen As W_SCREEN) As String
    screenKey = CStr(ObjPtr(psScreen))
End Function

Public Sub widgRegisterOverlayScreen(ByVal psScreen As W_SCREEN, ByVal zOrder As Long)
    Dim it As Long, pos As Long
    ensureSets
    ' OverlayScreenAssert(psScreen, zOrder): a null screen returns; the reserved z-orders' check (the
    ' Bits macro: NetPlay.players[selectedPlayer].isSpectator) only guards an order no caller here uses
    If psScreen Is Nothing Then Exit Sub
    it = overlayIndexOf(psScreen)
    If it <> -1 Then
        ' screen already exists in overlay stack

        ' check if it's queued for removal, and if so clear that status
        If overlaysToDelete.Exists(screenKey(psScreen)) Then
            debug_ 218, LOG_WZ, "widgRegisterOverlayScreen", "Overlay was queued for deletion, but is being re-added - clear the queued deletion"
            overlaysToDelete.Remove screenKey(psScreen)
        End If

        If zOrder = overlayZOrders(it) Then
            ' no need to update - duplicate call (same zOrder)
            Exit Sub
        Else
            ' remove the existing entry
            overlayEraseAt it
            ' fall-through to inserting it again in the new zOrder
        End If
    End If
    ' the screens are stored in decreasing z-order
    ' (std::upper_bound with a.zOrder > b.zOrder: the first overlay whose zOrder is less than zOrder)
    pos = 0
    Do While pos < overlayCount
        If zOrder > overlayZOrders(pos) Then Exit Do
        pos = pos + 1
    Loop
    overlayInsertAt pos, psScreen, zOrder
    If Not overlaySet.Exists(screenKey(psScreen)) Then overlaySet.Add screenKey(psScreen), psScreen
End Sub

' Note: If priorScreen is not the "regular" screen (i.e. if priorScreen is an overlay screen) it should already be registered
Public Sub widgRegisterOverlayScreenOnTopOfScreen(ByVal psScreen As W_SCREEN, ByVal priorScreen As W_SCREEN)
    Dim it As Long, i As Long
    ensureSets
    it = overlayIndexOf(priorScreen)
    If it <> -1 Then
        ' use the same z-order as priorScreen, but insert *before* it in the list (i.e. "above" it, since overlays are stored in decreasing z-order)
        overlayInsertAt it, psScreen, overlayZOrders(it)
        If Not overlaySet.Exists(screenKey(psScreen)) Then overlaySet.Add screenKey(psScreen), psScreen
    Else
        ' priorScreen does not exist in the overlays list, so it is probably the "regular" screen
        ' so use z-order 0
        it = -1
        For i = 0 To overlayCount - 1
            If overlayZOrders(i) = 0 Then
                it = i
                Exit For
            End If
        Next
        If it <> -1 Then
            ' found existing screen with z-order 0
            ' insert *before* it in the list (i.e. "above" it, since overlays are stored in decreasing z-order)
            overlayInsertAt it, psScreen, 0
            If Not overlaySet.Exists(screenKey(psScreen)) Then overlaySet.Add screenKey(psScreen), psScreen
        Else
            ' just insert this overlay at the bottom of the overlay list
            overlayInsertAt overlayCount, psScreen, 0
            If Not overlaySet.Exists(screenKey(psScreen)) Then overlaySet.Add screenKey(psScreen), psScreen
        End If
    End If
End Sub

Public Sub widgRemoveOverlayScreen(ByVal psScreen As W_SCREEN)
    ensureSets
    If overlayIndexOf(psScreen) <> -1 Then
        If Not overlaysToDelete.Exists(screenKey(psScreen)) Then overlaysToDelete.Add screenKey(psScreen), psScreen
    End If
End Sub

Public Function isRegisteredOverlayScreen(ByVal psScreen As W_SCREEN) As Boolean
    isRegisteredOverlayScreen = overlayIndexOf(psScreen) <> -1
End Function

Private Sub cleanupDeletedOverlays()
    Dim i As Long, k As Variant
    ensureSets
    If overlaysToDelete.Count = 0 Then Exit Sub
    i = 0
    Do While i < overlayCount
        If overlaysToDelete.Exists(screenKey(overlayScreens(i))) Then
            overlayEraseAt i
        Else
            i = i + 1
        End If
    Loop
    For Each k In overlaysToDelete.Keys
        If overlaySet.Exists(k) Then overlaySet.Remove k
    Next
    overlaysToDelete.RemoveAll
End Sub

Public Function WIDGET_idealWidth(ByVal self As WIDGET) As Long
    If Not self.has_defaultIdealWidth Then
        self.has_defaultIdealWidth = True
        self.defaultIdealWidth = self.width()
    End If

    WIDGET_idealWidth = self.defaultIdealWidth
End Function

Public Function WIDGET_idealHeight(ByVal self As WIDGET) As Long
    If Not self.has_defaultIdealHeight Then
        self.has_defaultIdealHeight = True
        self.defaultIdealHeight = self.height()
    End If

    WIDGET_idealHeight = self.defaultIdealHeight
End Function

' (VB) forEachOverlayScreen / forEachOverlayScreenBottomUp are loops at their uses (an index over a copy of
' the overlays, then cleanupDeletedOverlays, as iterateOverlayScreens does)
Private Function overlaysCopy(out() As W_SCREEN, zOut() As Long) As Long
    Dim i As Long
    If overlayCount = 0 Then Exit Function
    ReDim out(0 To overlayCount - 1)
    ReDim zOut(0 To overlayCount - 1)
    For i = 0 To overlayCount - 1
        Set out(i) = overlayScreens(i)
        zOut(i) = overlayZOrders(i)
    Next
    overlaysCopy = overlayCount
End Function

' void widgForEachOverlayScreen(func): func(psScreen, zOrder) returns false to stop enumerating
Public Sub widgForEachOverlayScreen(ByVal func As WzFunc)
    Dim scr() As W_SCREEN, z() As Long, n As Long, i As Long
    If func Is Nothing Then Exit Sub
    n = overlaysCopy(scr, z)
    For i = 0 To n - 1
        If Not CBool(WzFunc_call(func, scr(i), z(i))) Then
            Exit For ' stop enumerating
        End If
    Next
    ' now that we aren't in the middle of enumerating overlays, handling removing any that were queued for deletion
    cleanupDeletedOverlays
End Sub

' (VB) widgForEachOverlayScreen with multiplay.cpp sendDataCheck2's lambda (it counts each overlay's zOrder): the
' overlays' zOrders, in the overlays' order, in zOrders(0 .. count - 1); the count returned
Public Function widgGetOverlayScreenZOrders(zOrders() As Long) As Long
    Dim scr() As W_SCREEN, n As Long
    n = overlaysCopy(scr, zOrders)
    ' now that we aren't in the middle of enumerating overlays, handling removing any that were queued for deletion
    cleanupDeletedOverlays
    widgGetOverlayScreenZOrders = n
End Function

Public Sub widgOverlaysScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim scr() As W_SCREEN, z() As Long, n As Long, i As Long
    n = overlaysCopy(scr, z)
    For i = n - 1 To 0 Step -1          ' bottom-up
        scr(i).screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
    Next
    cleanupDeletedOverlays
End Sub

Private Function isScreenARegisteredOverlay(ByVal psScreen As W_SCREEN) As Boolean
    ensureSets
    If psScreen Is Nothing Then Exit Function
    isScreenARegisteredOverlay = overlaySet.Exists(screenKey(psScreen))
End Function

Public Function isMouseOverScreen(ByVal psScreen As W_SCREEN) As Boolean
    isMouseOverScreen = (psMouseOverWidgetScreen Is psScreen)
End Function

' global mouse coordinates - i.e. those returned from mouseX()/mouseY()
Public Function isMouseOverScreenOverlayChild(ByVal mx As Long, ByVal my As Long) As Boolean
    Dim bMouseIsOverOverlayChild As Boolean, mouseOverWidget As WIDGET, scr() As W_SCREEN, z() As Long
    Dim n As Long, i As Long, j As Long, psRoot As WIDGET, psCurr As WIDGET
    If Not psMouseOverWidgetScreen Is Nothing Then
        isMouseOverScreenOverlayChild = isScreenARegisteredOverlay(psMouseOverWidgetScreen)
        Exit Function
    End If
    bMouseIsOverOverlayChild = False
    Set mouseOverWidget = WIDGET_fromPtr(psMouseOverWidgetPtr)
    If Not mouseOverWidget Is Nothing Then
        ' Fallback - the mouse is over a widget, but the widget does not have a screen pointer
        ' Hit-test each overlay screen to identify if the mouse is over one
        ' (Note: This can currently occur with DropdownWidget members.)
        n = overlaysCopy(scr, z)
        For i = 0 To n - 1
            Set psRoot = scr(i).psForm

            ' Hit-test all children of root form
            For j = 0 To psRoot.childCount() - 1
                Set psCurr = psRoot.childAt(j)
                If psCurr.visible() Then
                    If psCurr.hitTest(mx, my) Then
                        ' hit test succeeded for child widget
                        bMouseIsOverOverlayChild = True
                        Exit For
                    End If
                End If
            Next
            If bMouseIsOverOverlayChild Then Exit For ' stop enumerating
        Next
        cleanupDeletedOverlays
    End If
    isMouseOverScreenOverlayChild = bMouseIsOverOverlayChild
End Function

Public Function isMouseClickDownOnScreenOverlayChild() As Boolean
    If vbDisplayOff Then isMouseClickDownOnScreenOverlayChild = False: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    If psClickDownWidgetScreen Is Nothing Then Exit Function
    isMouseClickDownOnScreenOverlayChild = isScreenARegisteredOverlay(psClickDownWidgetScreen)
End Function

'--------------------------------------------------------------------------
' WIDGET's constructors and destructor

' WIDGET::WIDGET(WIDGET_TYPE type)
Public Sub WIDGET_construct(ByVal self As WIDGET, Optional ByVal typ As WIDGET_TYPE = WIDG_UNSPECIFIED_TYPE)
    self.construct typ
End Sub

' WIDGET::WIDGET(W_INIT const *init, WIDGET_TYPE type)
Public Sub WIDGET_construct_init(ByVal self As WIDGET, init As W_INIT, Optional ByVal typ As WIDGET_TYPE = WIDG_UNSPECIFIED_TYPE)
    self.constructInit init, typ
End Sub

' std::make_shared<WIDGET>(type): a plain widget (a container)
Public Function WIDGET_new(Optional ByVal typ As WIDGET_TYPE = WIDG_UNSPECIFIED_TYPE) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_WIDGET)
    WIDGET_construct w, typ
    Set WIDGET_new = w
End Function

'--------------------------------------------------------------------------
' WIDGET's virtual methods: the defaults (WIDGET::<method>)

Public Sub WIDGET_setGeometryFromScreenRect(ByVal self As WIDGET, r As WzRect)
    Dim strongParent As WIDGET, xOffset As Long, yOffset As Long, parentRelativeRect As WzRect
    Set strongParent = self.parent()
    If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "lib/widget/widget.cpp:545", "setGeometryFromScreenRect", "No parent - failed": Exit Sub
    xOffset = strongParent.screenPosX()
    yOffset = strongParent.screenPosY()
    parentRelativeRect = WzRect_(r.x - xOffset, r.y - yOffset, r.w, r.h)
    self.setGeometryRect parentRelativeRect
End Sub

Public Sub WIDGET_screenSizeDidChange(ByVal self As WIDGET, ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    self.def_screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
End Sub

Public Sub WIDGET_widgetLost(ByVal self As WIDGET, ByVal w As WIDGET)
    Dim lockedParent As WIDGET
    Set lockedParent = self.parent()
    If Not lockedParent Is Nothing Then
        lockedParent.widgetLost w   ' We don't care about the lost widget, maybe the parent does. (Special code for W_TABFORM.)
    End If
End Sub

Public Sub WIDGET_focusLost(ByVal self As WIDGET)
End Sub

Public Sub WIDGET_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
End Sub

Public Function WIDGET_getTip(ByVal self As WIDGET) As String
    WIDGET_getTip = ""
End Function

Public Function WIDGET_getHelp(ByVal self As WIDGET) As WidgetHelp
    Set WIDGET_getHelp = Nothing
End Function

Public Function WIDGET_capturesMouseDrag(ByVal self As WIDGET, ByVal key As WIDGET_KEY) As Boolean
    WIDGET_capturesMouseDrag = False
End Function

Public Sub WIDGET_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
End Sub

Public Sub WIDGET_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
End Sub

Public Sub WIDGET_highlightLost(ByVal self As WIDGET)
End Sub

Public Sub WIDGET_run(ByVal self As WIDGET, psContext As W_CONTEXT)
End Sub

Public Sub WIDGET_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
End Sub

Public Sub WIDGET_geometryChanged(ByVal self As WIDGET)
End Sub

Public Sub WIDGET_mouseDragged(ByVal self As WIDGET, ByVal key As WIDGET_KEY, start As W_CONTEXT, current As W_CONTEXT)
End Sub

Public Function WIDGET_parentRelativeXOffset(ByVal self As WIDGET, ByVal coord As Long) As Long
    WIDGET_parentRelativeXOffset = self.x() + coord
End Function

Public Function WIDGET_parentRelativeYOffset(ByVal self As WIDGET, ByVal coord As Long) As Long
    WIDGET_parentRelativeYOffset = self.y() + coord
End Function

Public Function WIDGET_getScrollSnapOffsets(ByVal self As WIDGET, offsets() As Long) As Boolean
    WIDGET_getScrollSnapOffsets = False     ' nonstd::nullopt
End Function

Public Sub WIDGET_setTip(ByVal self As WIDGET, ByVal string_ As String)
    ASSERT_FAILURE "false", "lib/widget/widget.cpp:994", "setTip", "Can't set widget type %u's tip.", self.typ
End Sub

Public Sub WIDGET_setHelp(ByVal self As WIDGET, ByVal help As WidgetHelp)
    ASSERT_FAILURE "false", "lib/widget/widget.cpp:999", "setHelp", "Can't set widget type %u's help.", self.typ
End Sub

Public Function WIDGET_getState(ByVal self As WIDGET) As Long
    ASSERT_FAILURE "false", "lib/widget/widget.cpp:1029", "getState", "Can't get widget type %u's state.", self.typ
    WIDGET_getState = 0
End Function

Public Sub WIDGET_setFlash(ByVal self As WIDGET, ByVal enable As Boolean)
    ASSERT_FAILURE "false", "lib/widget/widget.cpp:1046", "setFlash", "Can't set widget type %u's flash.", self.typ
End Sub

Public Sub WIDGET_setState(ByVal self As WIDGET, ByVal state As Long)
    ASSERT_FAILURE "false", "lib/widget/widget.cpp:1072", "setState", "Can't set widget type %u's state.", self.typ
End Sub

Public Function WIDGET_getString(ByVal self As WIDGET) As String
    ASSERT_FAILURE "false", "lib/widget/widget.cpp:1088", "getString", "Can't get widget type %u's string.", self.typ
    WIDGET_getString = ""
End Function

Public Sub WIDGET_setString(ByVal self As WIDGET, ByVal string_ As String)
    ASSERT_FAILURE "false", "lib/widget/widget.cpp:1120", "setString", "Can't set widget type %u's string.", self.typ
End Sub

' Process all the widgets on a form.
' mx and my are the coords of the mouse relative to the form origin.
Public Sub WIDGET_runRecursive(ByVal self As WIDGET, psContext As W_CONTEXT)
    self.def_runRecursive psContext
End Sub

Public Function WIDGET_hitTest(ByVal self As WIDGET, ByVal x As Long, ByVal y As Long) As Boolean
    WIDGET_hitTest = self.def_hitTest(x, y)
End Function

Public Function WIDGET_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Set WIDGET_findMouseTargetRecursive = self.def_findMouseTargetRecursive(psContext, key, wasPressed)
End Function

Public Sub WIDGET_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    self.def_displayRecursive context
End Sub

' (VB) debugBoundingBoxesOnly, for WIDGET's displayRecursive
Public Function widgDebugBoundingBoxesOnly() As Boolean
    widgDebugBoundingBoxesOnly = debugBoundingBoxesOnly
End Function

' (VB) WIDGET::isMouseOverWidget's psMouseOverWidget.lock().get() == this, and its reset
Public Sub psMouseOverWidget_reset()
    psMouseOverWidgetPtr = 0
End Sub

'--------------------------------------------------------------------------
' W_SCREEN::make(customRootForm)
Public Function W_SCREEN_make(Optional ByVal customRootForm As WIDGET = Nothing) As W_SCREEN
    Dim screen As W_SCREEN
    Set screen = New W_SCREEN
    screen.initialize customRootForm
    Set W_SCREEN_make = screen
End Function

' (VB) W_SCREEN::~W_SCREEN's check: bWidgetsInitialized
Public Function widgIsInitialized() As Boolean
    widgIsInitialized = bWidgetsInitialized
End Function

' Check whether an ID has been used on a form
Private Function widgCheckIDForm(ByVal psForm As WIDGET, ByVal id As Long) As Boolean
    Dim i As Long
    If psForm.id = id Then
        widgCheckIDForm = True
        Exit Function
    End If
    For i = 0 To psForm.childCount() - 1
        If widgCheckIDForm(psForm.childAt(i), id) Then
            widgCheckIDForm = True
            Exit Function
        End If
    Next
    widgCheckIDForm = False
End Function

Private Function widgAddWidget(ByVal psScreen As W_SCREEN, ByVal formID As Long, ByVal initId As Long, ByVal widget As WIDGET) As Boolean
    Dim psParent As WIDGET
    If Not (Not widget Is Nothing) Then ASSERT_FAILURE "widget != nullptr", "lib/widget/widget.cpp:740", "widgAddWidget", "Invalid widget": Exit Function
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:741", "widgAddWidget", "Invalid screen pointer": Exit Function
    If Not (Not widgCheckIDForm(psScreen.psForm, initId)) Then ASSERT_FAILURE "!widgCheckIDForm(psScreen->psForm.get(), psInit->id)", "lib/widget/widget.cpp:742", "widgAddWidget", "ID number has already been used (%d)", initId: Exit Function
    ' Find the form to add the widget to.
    If formID = 0 Then
        ' Add to the base form
        Set psParent = psScreen.psForm
    Else
        Set psParent = widgGetFromID(psScreen, formID)
    End If
    If psParent Is Nothing Then
        ASSERT_FAILURE "psParent != nullptr && psParent->type == WIDG_FORM", "lib/widget/widget.cpp:754", "widgAddWidget", "Could not find parent form from formID"
        Exit Function
    End If
    If Not (psParent.typ = WIDG_FORM) Then ASSERT_FAILURE "psParent != nullptr && psParent->type == WIDG_FORM", "lib/widget/widget.cpp:754", "widgAddWidget", "Could not find parent form from formID": Exit Function

    psParent.attach widget
    widgAddWidget = True
End Function

' Add a form to the widget screen
Public Function widgAddForm(ByVal psScreen As W_SCREEN, psInit As W_FORMINIT) As WIDGET
    Dim psForm As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:763", "widgAddForm", "Invalid screen pointer": Exit Function
    If psInit.style And WFORM_CLICKABLE Then
        Set psForm = W_CLICKFORM_new_init(psInit)
    Else
        Set psForm = W_FORM_new_init(psInit)
    End If

    If widgAddWidget(psScreen, psInit.formID, psInit.id, psForm) Then Set widgAddForm = psForm
End Function

' Add a label to the widget screen
Public Function widgAddLabel(ByVal psScreen As W_SCREEN, psInit As W_LABINIT) As WIDGET
    Dim psLabel As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:780", "widgAddLabel", "Invalid screen pointer": Exit Function
    Set psLabel = W_LABEL_new_init(psInit)
    If widgAddWidget(psScreen, psInit.formID, psInit.id, psLabel) Then Set widgAddLabel = psLabel
End Function

' Add a button to the widget screen
Public Function widgAddButton(ByVal psScreen As W_SCREEN, psInit As W_BUTINIT) As WIDGET
    Dim psButton As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:788", "widgAddButton", "Invalid screen pointer": Exit Function
    Set psButton = W_BUTTON_new_init(psInit)
    If widgAddWidget(psScreen, psInit.formID, psInit.id, psButton) Then Set widgAddButton = psButton
End Function

' Add an edit box to the widget screen
Public Function widgAddEditBox(ByVal psScreen As W_SCREEN, psInit As W_EDBINIT) As WIDGET
    Dim psEdBox As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:796", "widgAddEditBox", "Invalid screen pointer": Exit Function
    Set psEdBox = W_EDITBOX_new_init(psInit)
    If widgAddWidget(psScreen, psInit.formID, psInit.id, psEdBox) Then Set widgAddEditBox = psEdBox
End Function

' Add a bar graph to the widget screen
Public Function widgAddBarGraph(ByVal psScreen As W_SCREEN, psInit As W_BARINIT) As WIDGET
    Dim psBarGraph As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:804", "widgAddBarGraph", "Invalid screen pointer": Exit Function
    Set psBarGraph = W_BARGRAPH_new_init(psInit)
    If widgAddWidget(psScreen, psInit.formID, psInit.id, psBarGraph) Then Set widgAddBarGraph = psBarGraph
End Function

' Add a slider to a form
Public Function widgAddSlider(ByVal psScreen As W_SCREEN, psInit As W_SLDINIT) As WIDGET
    Dim psSlider As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:812", "widgAddSlider", "Invalid screen pointer": Exit Function
    Set psSlider = W_SLIDER_new_init(psInit)
    If widgAddWidget(psScreen, psInit.formID, psInit.id, psSlider) Then Set widgAddSlider = psSlider
End Function

' Delete a widget from the screen (void widgDelete(WIDGET *widget): (VB) widgDelete_widget)
Public Sub widgDelete_widget(ByVal widget As WIDGET)
    Dim lockedScreen As W_SCREEN, widgetWithFocus As WIDGET, parent As WIDGET
    If widget Is Nothing Then
        Exit Sub
    End If
    ' (widgDebugAssertIfRunningScreen: DEBUG only)

    Set lockedScreen = widget.screenPointer
    If Not lockedScreen Is Nothing Then
        Set widgetWithFocus = lockedScreen.getWidgetWithFocus()
        If Not widgetWithFocus Is Nothing Then
            If widgetWithFocus.hasAncestor(widget) Then
                ' must trigger a resignation of focus
                lockedScreen.setFocus Nothing
            End If
        End If
        If lockedScreen.psForm Is widget Then
            Set lockedScreen.psForm = Nothing
        End If
    End If

    Set parent = widget.parent()
    If Not parent Is Nothing Then
        parent.detach widget
    End If
End Sub

' Delete a widget from the screen (void widgDelete(const std::shared_ptr<W_SCREEN> &psScreen, UDWORD id))
Public Sub widgDelete(ByVal psScreen As W_SCREEN, ByVal id As Long)
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:851", "widgDelete", "Invalid screen pointer": Exit Sub
    widgDelete_widget widgGetFromID(psScreen, id)
End Sub

Public Sub widgDeleteLater(ByVal psScreen As W_SCREEN, ByVal id As Long)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:856", "widgDeleteLater", "Invalid screen pointer": Exit Sub
    Set psWidget = widgGetFromID(psScreen, id)
    If Not psWidget Is Nothing Then
        psWidget.deleteLater
    End If
End Sub

' Find a widget on a form from its id number
Public Function widgFormGetFromID(ByVal widget As WIDGET, ByVal id As Long) As WIDGET
    Dim i As Long, w As WIDGET
    If Not (Not widget Is Nothing) Then ASSERT_FAILURE "widget != nullptr", "lib/widget/widget.cpp:867", "widgFormGetFromID", "Invalid widget pointer": Exit Function
    If widget.id = id Then
        Set widgFormGetFromID = widget
        Exit Function
    End If
    For i = 0 To widget.childCount() - 1
        Set w = widgFormGetFromID(widget.childAt(i), id)
        If Not w Is Nothing Then
            Set widgFormGetFromID = w
            Exit Function
        End If
    Next
    Set widgFormGetFromID = Nothing
End Function

' Find a widget in a screen from its ID number
Public Function widgGetFromID(ByVal psScreen As W_SCREEN, ByVal id As Long) As WIDGET
    If vbDisplayOff Then Set widgGetFromID = Nothing: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:886", "widgGetFromID", "Invalid screen pointer": Exit Function
    Set widgGetFromID = widgFormGetFromID(psScreen.psForm, id)
End Function

Public Sub widgHide(ByVal psScreen As W_SCREEN, ByVal id As Long)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:892", "widgHide", "Invalid screen pointer": Exit Sub
    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget != nullptr", "lib/widget/widget.cpp:894", "widgHide", "Couldn't find widget from id.": Exit Sub

    psWidget.hide
End Sub

Public Sub widgReveal(ByVal psScreen As W_SCREEN, ByVal id As Long)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:901", "widgReveal", "Invalid screen pointer": Exit Sub
    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget != nullptr", "lib/widget/widget.cpp:903", "widgReveal", "Couldn't find widget from id.": Exit Sub

    psWidget.show
End Sub

' Return the ID of the widget the mouse was over this frame
Public Function widgGetMouseOver(ByVal psScreen As W_SCREEN) As Long
    Dim lockedMouserOverWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:911", "widgGetMouseOver", "Invalid screen pointer": widgGetMouseOver = 0: Exit Function
    ' Don't actually need the screen parameter at the moment - but it might be
    '   handy if psMouseOverWidget needs to stop being a static and moves into
    '   the screen structure

    Set lockedMouserOverWidget = WIDGET_fromPtr(psMouseOverWidgetPtr)
    If Not lockedMouserOverWidget Is Nothing Then
        widgGetMouseOver = lockedMouserOverWidget.id
        Exit Function
    End If

    widgGetMouseOver = 0
End Function

Public Function isMouseOverSomeWidget(ByVal psScreen As W_SCREEN) As Boolean
    Dim lockedMouserOverWidget As WIDGET
    Set lockedMouserOverWidget = WIDGET_fromPtr(psMouseOverWidgetPtr)
    If Not lockedMouserOverWidget Is Nothing Then
        isMouseOverSomeWidget = Not (lockedMouserOverWidget Is psScreen.psForm)
        Exit Function
    End If

    isMouseOverSomeWidget = False
End Function

' Return the user data for a widget (void *: an object, a handle or Empty)
Public Function widgGetUserData(ByVal psScreen As W_SCREEN, ByVal id As Long) As Variant
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:937", "widgGetUserData", "Invalid screen pointer": Exit Function

    Set psWidget = widgGetFromID(psScreen, id)
    If Not psWidget Is Nothing Then
        WzFunc_assign widgGetUserData, psWidget.pUserData
        Exit Function
    End If

    widgGetUserData = Empty
End Function

' Return the user data for a widget
Public Function widgGetUserData2(ByVal psScreen As W_SCREEN, ByVal id As Long) As Long
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:953", "widgGetUserData2", "Invalid screen pointer": widgGetUserData2 = 0: Exit Function

    Set psWidget = widgGetFromID(psScreen, id)
    If Not psWidget Is Nothing Then
        widgGetUserData2 = psWidget.UserData
        Exit Function
    End If

    widgGetUserData2 = 0
End Function

' Set user data for a widget
Public Sub widgSetUserData(ByVal psScreen As W_SCREEN, ByVal id As Long, ByVal UserData As Variant)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:969", "widgSetUserData", "Invalid screen pointer": Exit Sub

    Set psWidget = widgGetFromID(psScreen, id)
    If Not psWidget Is Nothing Then
        If IsObject(UserData) Then Set psWidget.pUserData = UserData Else psWidget.pUserData = UserData
    End If
End Sub

' Set user data for a widget
Public Sub widgSetUserData2(ByVal psScreen As W_SCREEN, ByVal id As Long, ByVal UserData As Long)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:982", "widgSetUserData2", "Invalid screen pointer": Exit Sub

    Set psWidget = widgGetFromID(psScreen, id)
    If Not psWidget Is Nothing Then
        psWidget.UserData = UserData
    End If
End Sub

' Set tip string for a widget
Public Sub widgSetTip(ByVal psScreen As W_SCREEN, ByVal id As Long, ByVal pTip As String)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1005", "widgSetTip", "Invalid screen pointer": Exit Sub
    Set psWidget = widgGetFromID(psScreen, id)

    If psWidget Is Nothing Then
        Exit Sub
    End If

    psWidget.setTip pTip
End Sub

' Return which key was used to press the last returned widget
Public Function widgGetButtonKey_DEPRECATED(ByVal psScreen As W_SCREEN) As Long
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1019", "widgGetButtonKey_DEPRECATED", "Invalid screen pointer": widgGetButtonKey_DEPRECATED = lastReleasedKey_DEPRECATED: Exit Function
    ' Don't actually need the screen parameter at the moment - but it might be
    '   handy if released needs to stop being a static and moves into
    '   the screen structure

    widgGetButtonKey_DEPRECATED = lastReleasedKey_DEPRECATED
End Function

' Get a button or clickable form's state
Public Function widgGetButtonState(ByVal psScreen As W_SCREEN, ByVal id As Long) As Long
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1036", "widgGetButtonState", "Invalid screen pointer": widgGetButtonState = 0: Exit Function
    ' Get the button
    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget", "lib/widget/widget.cpp:1039", "widgGetButtonState", "Couldn't find widget by ID %u", id: widgGetButtonState = 0: Exit Function

    widgGetButtonState = psWidget.getState()
End Function

Public Sub widgSetButtonFlash(ByVal psScreen As W_SCREEN, ByVal id As Long)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1051", "widgSetButtonFlash", "Invalid screen pointer": Exit Sub
    ' Get the button
    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget", "lib/widget/widget.cpp:1054", "widgSetButtonFlash", "Couldn't find widget by ID %u", id: Exit Sub

    psWidget.setFlash True
End Sub

Public Sub widgClearButtonFlash(ByVal psScreen As W_SCREEN, ByVal id As Long)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1061", "widgClearButtonFlash", "Invalid screen pointer": Exit Sub
    ' Get the button
    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget", "lib/widget/widget.cpp:1064", "widgClearButtonFlash", "Couldn't find widget by ID %u", id: Exit Sub

    psWidget.setFlash False
End Sub

' Set a button or clickable form's state
Public Sub widgSetButtonState(ByVal psScreen As W_SCREEN, ByVal id As Long, ByVal state As Long)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1078", "widgSetButtonState", "Invalid screen pointer": Exit Sub
    ' Get the button
    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget", "lib/widget/widget.cpp:1081", "widgSetButtonState", "Couldn't find widget by ID %u", id: Exit Sub

    psWidget.setState state
End Sub

' Return a pointer to a buffer containing the current string of a widget.
' NOTE: The string must be copied out of the buffer (VB: the String itself)
Public Function widgGetString(ByVal psScreen As W_SCREEN, ByVal id As Long) As String
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1097", "widgGetString", "Invalid screen pointer": widgGetString = "": Exit Function
    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget", "lib/widget/widget.cpp:1099", "widgGetString", "Couldn't find widget by ID %u", id: widgGetString = "": Exit Function

    widgGetString = psWidget.getString()
End Function

Public Function widgGetWzString(ByVal psScreen As W_SCREEN, ByVal id As Long) As String
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1109", "widgGetWzString", "Invalid screen pointer": widgGetWzString = "": Exit Function
    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget", "lib/widget/widget.cpp:1111", "widgGetWzString", "Couldn't find widget by ID %u", id: widgGetWzString = "": Exit Function

    widgGetWzString = psWidget.getString()
End Function

' Set the text in a widget
Public Sub widgSetString(ByVal psScreen As W_SCREEN, ByVal id As Long, ByVal pText As String)
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1126", "widgSetString", "Invalid screen pointer": Exit Sub
    ' Get the widget
    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget", "lib/widget/widget.cpp:1129", "widgSetString", "Couldn't find widget by ID %u", id: Exit Sub

    If psWidget.typ = WIDG_EDITBOX Then
        If psScreen.hasFocus(psWidget) Then
            psScreen.setFocus Nothing
        End If
    End If

    psWidget.setString pText
End Sub

Private Sub widgProcessMouseOver(ByVal widget As WIDGET, psContext As W_CONTEXT)
    Dim lockedScreen As W_SCREEN, lockedLastHighlight As WIDGET
    If Not widget.visible() Then
        ' special case for root form
        ' since the findMouseTargetRecursive of children is only called if they are visible
        ' this should only trigger for the root form of a screen that's been set to invisible
        Exit Sub
    End If

    If widget.transparentToMouse() Then
        Exit Sub
    End If

    If psMouseOverWidgetPtr = 0 Then
        psMouseOverWidgetPtr = ObjPtr(widget) ' Mark that the mouse is over a widget (if we haven't already).
    End If

    If widget.isMouseOverWidget() Then
        Set lockedScreen = widget.screenPointer
        If Not lockedScreen Is Nothing Then
            If Not lockedScreen.isLastHighlight(widget) Then
                Set lockedLastHighlight = lockedScreen.lastHighlight
                If Not lockedLastHighlight Is Nothing Then
                    lockedLastHighlight.highlightLost
                End If
                lockedScreen.setLastHighlight widget ' Mark that the mouse is over a widget (if we haven't already).
                widget.highlight psContext
            End If
            If Not (psMouseOverWidgetScreen Is lockedScreen) Then
                If Not psMouseOverWidgetScreen Is Nothing Then
                    ' ensure the last mouseOverWidgetScreen receives highlightLost event
                    Set lockedLastHighlight = psMouseOverWidgetScreen.lastHighlight
                    If Not lockedLastHighlight Is Nothing Then
                        lockedLastHighlight.highlightLost
                    End If
                    psMouseOverWidgetScreen.lastHighlightReset
                End If
                ' update psMouseOverWidgetScreen
                Set psMouseOverWidgetScreen = lockedScreen
            End If
        Else
            ' Note: There are rare exceptions where this can happen, if a WIDGET is doing something funky like drawing widgets it hasn't attached.
            Set psMouseOverWidgetScreen = Nothing
        End If
    End If
End Sub

Private Sub widgProcessMouseClick(ByVal widget As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean)
    Dim psClickedWidget As WIDGET
    widgProcessMouseOver widget, psContext

    If widget.isMouseOverWidget() Then
        Set psClickedWidget = widget.processClick(psContext, key, wasPressed)
        If psClickedWidget Is Nothing Then
            Exit Sub
        End If
        If wasPressed Then
            Set psClickDownWidgetScreen = psClickedWidget.screenPointer
            If psClickedWidget.capturesMouseDrag(key) Then
                Set widgetKeyCurrentState(key).capturedDragWidget = psClickedWidget
                widgetKeyCurrentState(key).dragStartPos = W_CONTEXT_convertToScreenContext(psContext)
                widgetKeyCurrentState(key).dragLastPos = widgetKeyCurrentState(key).dragStartPos
            End If
        Else
            Set psClickDownWidgetScreen = Nothing
        End If
    End If
End Sub

' static void WIDGET::processMouseDragEvent(const W_CONTEXT &sContext, WIDGET_KEY wkey, WIDGET_KEYSTATE* pState, bool alsoTriggerReleased)
Private Sub WIDGET_processMouseDragEvent(sContext As W_CONTEXT, ByVal wkey As WIDGET_KEY, pState As WIDGET_KEYSTATE, ByVal alsoTriggerReleased As Boolean)
    Dim dragPosChanged As Boolean, shiftedStartPos As W_CONTEXT, shiftedContext As W_CONTEXT
    Dim capturedWidgetParent As WIDGET, parentScreenX As Long, parentScreenY As Long
    dragPosChanged = Not (pState.dragLastPos.xOffset = sContext.xOffset And pState.dragLastPos.yOffset = sContext.yOffset And pState.dragLastPos.mx = sContext.mx And pState.dragLastPos.my = sContext.my)
    If Not dragPosChanged And Not alsoTriggerReleased Then
        ' no change since last event
        Exit Sub
    End If

    shiftedStartPos = pState.dragStartPos
    shiftedContext = sContext
    Set capturedWidgetParent = pState.capturedDragWidget.parent()
    If Not capturedWidgetParent Is Nothing Then
        parentScreenX = capturedWidgetParent.screenPosX()
        parentScreenY = capturedWidgetParent.screenPosY()
        shiftedStartPos.xOffset = parentScreenX
        shiftedStartPos.yOffset = parentScreenY
        shiftedStartPos.mx = shiftedStartPos.mx - parentScreenX
        shiftedStartPos.my = shiftedStartPos.my - parentScreenY
        shiftedContext.xOffset = parentScreenX
        shiftedContext.yOffset = parentScreenY
        shiftedContext.mx = shiftedContext.mx - parentScreenX
        shiftedContext.my = shiftedContext.my - parentScreenY
    End If
    If dragPosChanged Then
        ' Trigger a mouseDragged event for the widget
        pState.capturedDragWidget.mouseDragged wkey, shiftedStartPos, shiftedContext
        pState.dragLastPos = sContext
    End If

    If alsoTriggerReleased Then
        ' Trigger a released() call on the widget that captured the drag (to ensure it always gets one, even if the mouse was released off-widget)
        pState.capturedDragWidget.released shiftedContext, wkey
        Set pState.capturedDragWidget = Nothing
    End If
End Sub

' Execute a set of widgets for one cycle.
' Returns a list of activated widgets: (VB) WidgetTriggers const &: triggers(0 .. n - 1) (psScreen's retWidgets),
' n returned (triggers erased when 0).
Public Function widgRunScreen(ByVal psScreen As W_SCREEN, triggers() As WidgetTrigger) As Long
    If vbDisplayOff Then widgRunScreen = 0: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    Dim sContext As W_CONTEXT, zero As W_CONTEXT, n As Long, c As Long, action As Long, mkey As Long, cx As Long, cy As Long
    Dim wkey As WIDGET_KEY, pressed As Boolean, mouseOverWidget As WIDGET, key As Long
    Dim scr() As W_SCREEN, z() As Long, ns As Long, i As Long
    Erase triggers
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1421", "widgRunScreen", "Invalid screen pointer": widgRunScreen = 0: Exit Function
    psScreen.clearRetWidgets
    cleanupDeletedOverlays

    ' Initialise the context
    sContext = zero
    psMouseOverWidgetPtr = 0            ' psMouseOverWidget.reset();

    ' (#ifdef DEBUG: psCurrentlyRunningScreen)

    ' Note which keys have been pressed
    lastReleasedKey_DEPRECATED = WKEY_NONE
    If getWidgetsStatus() Then
        n = inputGetClickCount()
        For c = 0 To n - 1
            inputGetClick c, action, mkey, cx, cy
            Select Case mkey
            Case MOUSE_LMB: wkey = WKEY_PRIMARY
            Case MOUSE_RMB: wkey = WKEY_SECONDARY
            Case Else: GoTo continue_  ' Who cares about other mouse buttons?
            End Select
            Select Case action
            Case MousePress_Press: pressed = True
            Case MousePress_Release: pressed = False
            Case Else: GoTo continue_
            End Select

            sContext = zero
            sContext.mx = cx
            sContext.my = cy
            Set mouseOverWidget = Nothing

            If Not pressed Or (pressed And Not widgetKeyCurrentState(wkey).capturedDragWidget Is Nothing) Then
                If Not widgetKeyCurrentState(wkey).capturedDragWidget Is Nothing Then
                    ' Deliver a last mouseDragged event for the widget, with the final position sourced from this mousepress event, as well as a released event
                    WIDGET_processMouseDragEvent sContext, wkey, widgetKeyCurrentState(wkey), True

                    ' consume the release event - prevent it from going to another widget on top of which the mouse may have released
                    Set psClickDownWidgetScreen = Nothing

                    widgetKeyCurrentState(wkey).pressed = pressed
                    lastReleasedKey_DEPRECATED = wkey
                    GoTo continue_
                End If
            End If

            widgetKeyCurrentState(wkey).pressed = pressed

            If mouseOverWidget Is Nothing Then
                ns = overlaysCopy(scr, z)
                For i = 0 To ns - 1
                    Set mouseOverWidget = scr(i).psForm.findMouseTargetRecursive(sContext, wkey, pressed)
                    If Not mouseOverWidget Is Nothing Then Exit For
                Next
                cleanupDeletedOverlays
            End If
            If mouseOverWidget Is Nothing Then
                Set mouseOverWidget = psScreen.psForm.findMouseTargetRecursive(sContext, wkey, pressed)
            End If

            If Not mouseOverWidget Is Nothing Then
                widgProcessMouseClick mouseOverWidget, sContext, wkey, pressed
            End If

            lastReleasedKey_DEPRECATED = wkey
continue_:
        Next
    End If

    sContext = zero
    sContext.mx = mouseX()
    sContext.my = mouseY()

    ' For each current key state
    For key = WKEY_PRIMARY To WKEY_SECONDARY
        ' If widget_key is captured, and is down at the end of this batch of mouse events
        If widgetKeyCurrentState(key).pressed Then
            If Not widgetKeyCurrentState(key).capturedDragWidget Is Nothing Then
                ' Trigger a mouseDragged event for the capture widget, with the final position (sourced from mouseX / mouseY)
                WIDGET_processMouseDragEvent sContext, key, widgetKeyCurrentState(key), False
            End If
        End If
    Next

    Set mouseOverWidget = Nothing
    ns = overlaysCopy(scr, z)
    For i = 0 To ns - 1
        Set mouseOverWidget = scr(i).psForm.findMouseTargetRecursive(sContext, WKEY_NONE, True)
        If Not mouseOverWidget Is Nothing Then Exit For
    Next
    cleanupDeletedOverlays
    If mouseOverWidget Is Nothing Then
        Set mouseOverWidget = psScreen.psForm.findMouseTargetRecursive(sContext, WKEY_NONE, True)
    End If
    If Not mouseOverWidget Is Nothing Then
        widgProcessMouseOver mouseOverWidget, sContext
    Else
        Set psMouseOverWidgetScreen = Nothing
    End If

    ' reset context to screen context
    sContext = zero
    sContext.mx = mouseX()
    sContext.my = mouseY()

    ' Process the screen's widgets
    ns = overlaysCopy(scr, z)
    For i = 0 To ns - 1
        scr(i).psForm.runRecursive sContext
        scr(i).psForm.run sContext  ' ensure run() is called on root form
    Next
    cleanupDeletedOverlays
    psScreen.psForm.runRecursive sContext

    runScheduledTasks
    deleteOldWidgets   ' Delete any widgets that called deleteLater() while being run.
    cleanupDeletedOverlays

    ' Return the ID of a pressed button or finished edit box if any
    n = psScreen.retWidgetCount()
    If n > 0 Then
        ReDim triggers(0 To n - 1)
        For i = 0 To n - 1
            Set triggers(i).widget = psScreen.retWidgetAt(i)
        Next
    End If
    widgRunScreen = n
End Function

' Display the screen's widgets in their current state
' (Call after calling widgRunScreen, this allows the input
'  processing to be separated from the display of the widgets).
Public Sub widgDisplayScreen(ByVal psScreen As W_SCREEN)
    ' To toggle debug bounding boxes: Press: Left Shift   --  --  --------------
    '                                        Left Ctrl  ------------  --  --  ----
    Static debugSequence(0 To 16) As Long, debugLoc As Long, debugBoundingBoxes As Boolean, seqReady As Boolean
    Dim debugCode As Long, skipDrawing As Boolean, sContext As W_CONTEXT, scr() As W_SCREEN, z() As Long, ns As Long, i As Long
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/widget.cpp:1633", "widgDisplayScreen", "Invalid screen pointer": Exit Sub
    If Not seqReady Then
        ' static const int debugSequence[] = { -1, 0, 1, 3, 1, 3, 1, 3, 2, 3, 2, 3, 2, 3, 1, 0, -1}; (debugLoc: an index)
        seqReady = True
        debugSequence(0) = -1: debugSequence(1) = 0: debugSequence(2) = 1: debugSequence(3) = 3: debugSequence(4) = 1
        debugSequence(5) = 3: debugSequence(6) = 1: debugSequence(7) = 3: debugSequence(8) = 2: debugSequence(9) = 3
        debugSequence(10) = 2: debugSequence(11) = 3: debugSequence(12) = 2: debugSequence(13) = 3: debugSequence(14) = 1
        debugSequence(15) = 0: debugSequence(16) = -1
        debugLoc = 0
    End If
    debugCode = -keyDown(KEY_LCTRL) + 2 * -keyDown(KEY_LSHIFT)
    If debugSequence(debugLoc + 1) = -1 Then
        debugLoc = 0
    ElseIf debugSequence(debugLoc) = debugCode Then
        ' debugLoc stays
    ElseIf debugSequence(debugLoc + 1) = debugCode Then
        debugLoc = debugLoc + 1
    Else
        debugLoc = 0
    End If
    debugBoundingBoxes = debugBoundingBoxes Xor (debugSequence(debugLoc + 1) = -1)

    skipDrawing = Not gfx_api_context_shouldDraw()

    cleanupDeletedOverlays

    ' Process any user callback functions
    sContext.mx = mouseX()
    sContext.my = mouseY()
    psScreen.psForm.processCallbacksRecursive sContext

    If Not skipDrawing Then
        ' Display the widgets.
        psScreen.psForm.displayRecursiveRoot
    End If

    ' Always overlays on-top (i.e. draw them last)
    ns = overlaysCopy(scr, z)
    For i = ns - 1 To 0 Step -1         ' <- enumerate in *increasing* z-order for drawing
        scr(i).psForm.processCallbacksRecursive sContext
        If Not skipDrawing Then
            scr(i).psForm.displayRecursiveRoot
        End If
    Next
    cleanupDeletedOverlays

    deleteOldWidgets   ' Delete any widgets that called deleteLater() while being displayed.

    If Not skipDrawing Then
        ' Display the tool tip if there is one
        tipDisplay
    End If

    If debugBoundingBoxes And Not skipDrawing Then
        debugBoundingBoxesOnly = True
        psScreen.psForm.displayRecursiveRoot
        debugBoundingBoxesOnly = False
    End If
End Sub

Public Sub WidgSetAudio(ByVal Callback As WzFunc, ByVal HilightID As Long, ByVal ClickedID As Long, ByVal ErrorID As Long)
    Set AudioCallback = Callback
    HilightAudioID = HilightID
    ClickedAudioID = ClickedID
    ErrorAudioID = ErrorID
End Sub

Public Function WidgGetAudioCallback() As WzFunc
    Set WidgGetAudioCallback = AudioCallback
End Function

Public Function WidgGetHilightAudioID() As Long
    WidgGetHilightAudioID = HilightAudioID
End Function

Public Function WidgGetClickedAudioID() As Long
    WidgGetClickedAudioID = ClickedAudioID
End Function

Public Function WidgGetErrorAudioID() As Long
    WidgGetErrorAudioID = ErrorAudioID
End Function

Public Sub setWidgetsStatus(ByVal var As Boolean)
    bWidgetsActive = var
End Sub

Public Function getWidgetsStatus() As Boolean
    getWidgetsStatus = bWidgetsActive
End Function

'--------------------------------------------------------------------------
' WidgetGraphicsContext

' (VB) the constructor: every field zero (a VB UDT starts so)
Public Sub WidgetGraphicsContext_construct(c As WidgetGraphicsContext)
    Dim z As WidgetGraphicsContext
    c = z
End Sub

Public Function WidgetGraphicsContext_getXOffset(c As WidgetGraphicsContext) As Long
    WidgetGraphicsContext_getXOffset = c.offset.x
End Function

Public Function WidgetGraphicsContext_getYOffset(c As WidgetGraphicsContext) As Long
    WidgetGraphicsContext_getYOffset = c.offset.y
End Function

Public Function WidgetGraphicsContext_allowChildDisplayRecursiveIfSelfClipped(c As WidgetGraphicsContext) As Boolean
    WidgetGraphicsContext_allowChildDisplayRecursiveIfSelfClipped = c.allowChildDisplayIfSelfClipped
End Function

Public Function WidgetGraphicsContext_clipContains(c As WidgetGraphicsContext, rect As WzRect) As Boolean
    Dim r As WzRect
    If Not c.clipped Then
        WidgetGraphicsContext_clipContains = True
        Exit Function
    End If
    r = WzRect_(c.offset.x + rect.x, c.offset.y + rect.y, rect.w, rect.h)
    WidgetGraphicsContext_clipContains = WzRect_containsRect(c.clipRect, r)
End Function

' bool clipIntersects(WzRect const& rect, WzRect* output_intersection) const (VB: output_intersection always written; a
' caller passing nullptr passes a scratch rect)
Public Function WidgetGraphicsContext_clipIntersects(c As WidgetGraphicsContext, rect As WzRect, output_intersection As WzRect) As Boolean
    Dim comparisonRect As WzRect
    If Not c.clipped Then
        output_intersection = WzRect_(c.offset.x + rect.x, c.offset.y + rect.y, rect.w, rect.h)
        WidgetGraphicsContext_clipIntersects = True
        Exit Function
    End If
    comparisonRect = WzRect_(c.offset.x + rect.x, c.offset.y + rect.y, rect.w, rect.h)
    If WzRect_intersects(c.clipRect, comparisonRect) Then
        output_intersection = WzRect_intersectionWith(c.clipRect, comparisonRect)
        WidgetGraphicsContext_clipIntersects = True
        Exit Function
    End If
    WidgetGraphicsContext_clipIntersects = False
End Function

Public Function WidgetGraphicsContext_translatedBy(c As WidgetGraphicsContext, ByVal x As Long, ByVal y As Long) As WidgetGraphicsContext
    Dim newContext As WidgetGraphicsContext
    newContext = c
    newContext.offset.x = newContext.offset.x + x
    newContext.offset.y = newContext.offset.y + y
    WidgetGraphicsContext_translatedBy = newContext
End Function

Public Function WidgetGraphicsContext_clippedBy(c As WidgetGraphicsContext, newRect As WzRect) As WidgetGraphicsContext
    Dim newContext As WidgetGraphicsContext
    newContext = c
    newContext.clipRect = WzRect_(c.offset.x + newRect.x, c.offset.y + newRect.y, newRect.w, newRect.h)

    If c.clipped Then
        newContext.clipRect = WzRect_intersectionWith(newContext.clipRect, c.clipRect)
    End If

    newContext.clipped = True

    WidgetGraphicsContext_clippedBy = newContext
End Function

Public Function WidgetGraphicsContext_setAllowChildDisplayRecursiveIfSelfClipped(c As WidgetGraphicsContext, ByVal val_ As Boolean) As WidgetGraphicsContext
    Dim newContext As WidgetGraphicsContext
    newContext = c

    newContext.allowChildDisplayIfSelfClipped = val_
    WidgetGraphicsContext_setAllowChildDisplayRecursiveIfSelfClipped = newContext
End Function

' std::weak_ptr<WIDGET> getMouseOverWidget() (its lock: Nothing if none)
Public Function getMouseOverWidget() As WIDGET
    Set getMouseOverWidget = WIDGET_fromPtr(psMouseOverWidgetPtr)
End Function

' std::chrono::milliseconds widgGetClickHoldMS()
Public Function widgGetClickHoldMS() As Long
    widgGetClickHoldMS = 2000 ' TODO: Make this configurable?
End Function

' (VB) a WidgetHelp::WidgetInteractionTriggerFlagSet (std::set<InteractionTriggers>{t, ...}): the mask of 2 ^ t
Public Function WidgetHelp_triggers(ParamArray t() As Variant) As Long
    Dim i As Long, m As Long
    For i = 0 To UBound(t)
        m = m Or CLng(2 ^ CLng(t(i)))
    Next
    WidgetHelp_triggers = m
End Function
