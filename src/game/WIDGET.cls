VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WIDGET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/widget/widgbase.h: class WIDGET, the object of every widget (docs/CONVENTIONS.md,
' "Widgets"): kind says which C++ class it is (widget_classes_vb's WIDGET_KIND), ext(DEPTH_NAME)
' holds the fields of each class of its chain that has some (NAME_(w)). A std::shared_ptr<X> or X *
' of any widget class is a WIDGET (Nothing = nullptr); std::make_shared<X>(...) is X_new(...).
' This class has WIDGET's own fields and methods: a virtual method calls the generated vt_WIDGET_<name>
' (the most derived class's NAME_<name>, WIDGET's own being widget_c's WIDGET_<name>); a method that
' takes a UDT is Friend (a class's public method cannot take one). The parent and screen pointers
' are weak (std::weak_ptr): their addresses, cleared when the parent or screen dies.
Option Explicit

' (VB) the object model
Public kind As WIDGET_KIND
Private ext_(0 To 15) As Object

' WIDGET's public fields
Public id As Long                       ' UDWORD: The user set ID number for the widget. This is returned when e.g. a button is pressed.
Public typ As WIDGET_TYPE               ' The widget type
Public style As Long                    ' UDWORD: The style of the widget
Public displayFunction As WzFunc        ' WIDGET_DISPLAY: Override function to display the widget (Nothing = nullptr)
Public callback As WzFunc               ' WIDGET_CALLBACK: User callback (if any)
Private pUserData_ As Variant           ' void *: Pointer to a user data block (if any): an object, a handle or Empty
Public UserData As Long                 ' UDWORD: User data (if any)
Private screenPointer_ As Long          ' std::weak_ptr<W_SCREEN>: Pointer to screen the widget is on (if attached).
Public dirty As Boolean                 ' Whether widget is changed and needs to be redrawn

' protected
Public has_defaultIdealWidth As Boolean ' nonstd::optional<int32_t> defaultIdealWidth;
Public defaultIdealWidth As Long
Public has_defaultIdealHeight As Boolean
Public defaultIdealHeight As Long

' private
Private calcLayout As WzFunc            ' WIDGET_CALCLAYOUT_FUNC: Optional calc layout callback
Private onDelete As WzFunc              ' WIDGET_ONDELETE_FUNC: Optional callback called when the Widget is about to be deleted
Private customHitTest As WzFunc         ' WIDGET_HITTEST_FUNC: Optional hit-testing custom function
Private parentWidget As Long            ' std::weak_ptr<WIDGET>
Private childWidgets() As WIDGET        ' std::vector<std::shared_ptr<WIDGET>>
Private childCount_ As Long
Private dim_ As WzRect
Private isTransparentToClicks As Boolean ' = false
Private isTransparentToMouse As Boolean ' = false

'--------------------------------------------------------------------------
' (VB) the object model

Public Property Get ext(ByVal i As Long) As Object
    Set ext = ext_(i)
End Property

Public Property Set ext(ByVal i As Long, ByVal v As Object)
    Set ext_(i) = v
End Property

Public Property Get pUserData() As Variant
    If IsObject(pUserData_) Then Set pUserData = pUserData_ Else pUserData = pUserData_
End Property

Public Property Let pUserData(ByVal v As Variant)
    pUserData_ = v
End Property

Public Property Set pUserData(ByVal v As Variant)
    Set pUserData_ = v
End Property

Public Property Get screenPointer() As W_SCREEN
    Set screenPointer = W_SCREEN_fromPtr(screenPointer_)
End Property

'--------------------------------------------------------------------------
' virtual methods (C++'s virtual: the most derived class's implementation)

Public Sub widgetLost(ByVal w As WIDGET)
    vt_WIDGET_widgetLost Me, w
End Sub

Public Sub focusLost()
    vt_WIDGET_focusLost Me
End Sub

Friend Sub clicked(psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    vt_WIDGET_clicked Me, psContext, key
End Sub

Public Function getTip() As String
    getTip = vt_WIDGET_getTip(Me)
End Function

' WidgetHelp const *: Nothing = nullptr
Public Function getHelp() As WidgetHelp
    Set getHelp = vt_WIDGET_getHelp(Me)
End Function

Public Function capturesMouseDrag(ByVal key As WIDGET_KEY) As Boolean
    capturesMouseDrag = vt_WIDGET_capturesMouseDrag(Me, key)
End Function

Friend Sub released(psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    vt_WIDGET_released Me, psContext, key
End Sub

Friend Sub highlight(psContext As W_CONTEXT)
    vt_WIDGET_highlight Me, psContext
End Sub

Public Sub highlightLost()
    vt_WIDGET_highlightLost Me
End Sub

Friend Sub run(psContext As W_CONTEXT)
    vt_WIDGET_run Me, psContext
End Sub

Public Sub display(ByVal xOffset As Long, ByVal yOffset As Long)
    vt_WIDGET_display Me, xOffset, yOffset
End Sub

Public Sub geometryChanged()
    vt_WIDGET_geometryChanged Me
End Sub

Public Function hitTest(ByVal x As Long, ByVal y As Long) As Boolean
    hitTest = vt_WIDGET_hitTest(Me, x, y)
End Function

Friend Sub mouseDragged(ByVal key As WIDGET_KEY, start As W_CONTEXT, current As W_CONTEXT)
    vt_WIDGET_mouseDragged Me, key, start, current
End Sub

Public Function getState() As Long
    getState = vt_WIDGET_getState(Me)
End Function

Public Sub setState(ByVal state As Long)
    vt_WIDGET_setState Me, state
End Sub

Public Sub setFlash(ByVal enable As Boolean)
    vt_WIDGET_setFlash Me, enable
End Sub

Public Function getString() As String
    getString = vt_WIDGET_getString(Me)
End Function

Public Sub setString(ByVal string_ As String)
    vt_WIDGET_setString Me, string_
End Sub

' tooltip (C++'s setTip(std::string) and setTip(char const *): vbNullString is nullptr's empty string)
Public Sub setTip(ByVal string_ As String)
    vt_WIDGET_setTip Me, string_
End Sub

' formatted help information for this UI element (optional<WidgetHelp>: Nothing = nullopt)
Public Sub setHelp(ByVal help As WidgetHelp)
    vt_WIDGET_setHelp Me, help
End Sub

' used to handle screen resizing
Public Sub screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    vt_WIDGET_screenSizeDidChange Me, oldWidth, oldHeight, newWidth, newHeight
End Sub

Public Function parentRelativeXOffset(ByVal coord As Long) As Long
    parentRelativeXOffset = vt_WIDGET_parentRelativeXOffset(Me, coord)
End Function

Public Function parentRelativeYOffset(ByVal coord As Long) As Long
    parentRelativeYOffset = vt_WIDGET_parentRelativeYOffset(Me, coord)
End Function

Friend Sub setGeometryFromScreenRect(r As WzRect)
    vt_WIDGET_setGeometryFromScreenRect Me, r
End Sub

Public Function idealWidth() As Long
    idealWidth = vt_WIDGET_idealWidth(Me)
End Function

Public Function idealHeight() As Long
    idealHeight = vt_WIDGET_idealHeight(Me)
End Function

' optional<std::vector<uint32_t>>: has_value returned, the vector in offsets (0-based)
Public Function getScrollSnapOffsets(offsets() As Long) As Boolean
    getScrollSnapOffsets = vt_WIDGET_getScrollSnapOffsets(Me, offsets)
End Function

Friend Function findMouseTargetRecursive(psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Set findMouseTargetRecursive = vt_WIDGET_findMouseTargetRecursive(Me, psContext, key, wasPressed)
End Function

Friend Sub runRecursive(psContext As W_CONTEXT)
    vt_WIDGET_runRecursive Me, psContext
End Sub

' Display this widget, and all visible children.
Friend Sub displayRecursive(context As WidgetGraphicsContext)
    vt_WIDGET_displayRecursive Me, context
End Sub

' void displayRecursive() (a default WidgetGraphicsContext)
Public Sub displayRecursiveRoot()
    Dim context As WidgetGraphicsContext
    WidgetGraphicsContext_construct context
    displayRecursive context
End Sub

'--------------------------------------------------------------------------
' non-virtual methods

Public Sub show(Optional ByVal doShow As Boolean = True)
    If doShow Then
        style = style And Not WIDG_HIDDEN
    Else
        style = style Or WIDG_HIDDEN
    End If
End Sub

Public Sub hide()
    show False
End Sub

Public Function visible() As Boolean
    visible = (style And WIDG_HIDDEN) = 0
End Function

' std::shared_ptr<WIDGET> parent() const: Nothing if none (or gone)
Public Function parent() As WIDGET
    Set parent = WIDGET_fromPtr(parentWidget)
End Function

' Children const &children() const: childCount() of them, childAt(i) (0-based)
Public Function childCount() As Long
    childCount = childCount_
End Function

Public Function childAt(ByVal i As Long) As WIDGET
    Set childAt = childWidgets(i)
End Function

Friend Function screenGeometry() As WzRect
    screenGeometry = WzRect_(screenPosX(), screenPosY(), width(), height())
End Function

Friend Function geometry() As WzRect
    geometry = dim_
End Function

Public Function x() As Long
    x = dim_.x
End Function

Public Function y() As Long
    y = dim_.y
End Function

Public Function width() As Long
    width = dim_.w
End Function

Public Function height() As Long
    height = dim_.h
End Function

Public Function screenPosX() As Long
    Dim screenX As Long, psParent As WIDGET
    screenX = x()
    Set psParent = parent()
    Do While Not psParent Is Nothing
        screenX = psParent.parentRelativeXOffset(screenX)
        Set psParent = psParent.parent()
    Loop
    screenPosX = screenX
End Function

Public Function screenPosY() As Long
    Dim screenY As Long, psParent As WIDGET
    screenY = y()
    Set psParent = parent()
    Do While Not psParent Is Nothing
        screenY = psParent.parentRelativeYOffset(screenY)
        Set psParent = psParent.parent()
    Loop
    screenPosY = screenY
End Function

Public Function hasAncestor(ByVal psWidget As WIDGET) As Boolean
    Dim psParent As WIDGET
    Set psParent = parent()
    Do While Not psParent Is Nothing
        If psParent Is psWidget Then
            hasAncestor = True
            Exit Function
        End If
        Set psParent = psParent.parent()
    Loop
End Function

Public Sub move(ByVal x As Long, ByVal y As Long)
    setGeometry x, y, width(), height()
End Sub

Public Sub setGeometry(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim r As WzRect
    r = WzRect_(x, y, w, h)
    setGeometryRect r
End Sub

' void setGeometry(WzRect const &r) (widget.cpp)
Friend Sub setGeometryRect(r As WzRect)
    If WzRect_equal(dim_, r) Then
        Exit Sub  ' Nothing to do.
    End If
    dim_ = r
    geometryChanged
    dirty = True
End Sub

Public Function transparentToMouse() As Boolean
    transparentToMouse = isTransparentToMouse
End Function

'--------------------------------------------------------------------------
' widget.cpp's WIDGET methods

' WIDGET::WIDGET(WIDGET_TYPE type) (widget_c's WIDGET_construct)
Friend Sub construct(ByVal type_ As WIDGET_TYPE)
    id = &HFFFFEEEE
    typ = type_
    style = 0
    Set displayFunction = Nothing
    Set callback = Nothing
    pUserData_ = Empty
    UserData = 0
    Set calcLayout = Nothing
    Set onDelete = Nothing
    Set customHitTest = Nothing
    dim_ = WzRect_(0, 0, 1, 1)
    dirty = True
End Sub

' WIDGET::WIDGET(W_INIT const *init, WIDGET_TYPE type) (widget_c's WIDGET_construct_init)
Friend Sub constructInit(init As W_INIT, ByVal type_ As WIDGET_TYPE)
    id = init.id
    typ = type_
    style = init.style
    Set displayFunction = init.pDisplay
    Set callback = init.pCallback
    If IsObject(init.pUserData) Then Set pUserData_ = init.pUserData Else pUserData_ = init.pUserData
    UserData = init.UserData
    Set calcLayout = init.calcLayout
    Set onDelete = init.onDelete
    Set customHitTest = init.customHitTest
    dim_ = WzRect_(init.x, init.y, init.width, init.height)
    dirty = True

    ' Initialize and set the pUserData if necessary
    If Not init.initPUserDataFunc Is Nothing Then
        ' assert(pUserData == nullptr); // if the initPUserDataFunc is set, pUserData should not be already set (NDEBUG in a release build)
        WzFunc_assign pUserData_, WzFunc_call(init.initPUserDataFunc)
    End If

    ' if calclayout is not null, call it
    callCalcLayout
End Sub

' WIDGET::~WIDGET(), after the derived classes' destructors (C++'s virtual destructors); then (VB) the
' weak references to this widget that it knows of expire: the mouse-over widget, its screen's focus and
' last highlight, its children's parent pointers
Private Sub Class_Terminate()
    Dim scr As W_SCREEN, i As Long, self As WIDGET
    On Error GoTo fail                  ' (VB) VB can't raise an error out of Class_Terminate (it would show its
                                        ' message box and end the process): logged, as C++ would std::terminate
    WIDGET_destructAll Me

    If Not onDelete Is Nothing Then
        Set self = Me                ' (VB) WzFunc_call takes a copy of Me (CONVENTIONS, "Widgets")
        WzFunc_call onDelete, self    ' Call the onDelete function to handle any extra logic
        Set self = Nothing
    End If

    If psMouseOverWidgetPtr = ObjPtr(Me) Then psMouseOverWidget_reset
    If screenPointer_ <> 0 Then
        Set scr = W_SCREEN_fromPtr(screenPointer_)
        scr.forgetWidget ObjPtr(Me)
        Set scr = Nothing
    End If
    For i = 0 To childCount_ - 1
        childWidgets(i).clearParentPtr ObjPtr(Me)
    Next
    Erase childWidgets
    childCount_ = 0
    Exit Sub
fail:
    debug__ 0, LOG_ERROR, "~WIDGET", "(VB) run-time error %d in a widget's destructor (kind %d): %s", Err.Number, kind, Err.Description
End Sub

' Like "delete this;", but safe to call from display/run callbacks.
Public Sub deleteLater()
    If Not (Not widgetDeletionQueue_contains(Me)) Then ASSERT_FAILURE "std::find(widgetDeletionQueue.begin(), widgetDeletionQueue.end(), shared_widget_ptr) == widgetDeletionQueue.end()", "lib/widget/widget.cpp:527", "deleteLater", "Called deleteLater() twice on the same widget.": Exit Sub
    widgetDeletionQueue_push_back Me
End Sub

' widgbase.h
Public Sub removeAllChildren()
    Dim i As Long, child As WIDGET
    For i = 0 To childCount_ - 1
        Set child = childWidgets(i)
        If child.parent() Is Me Then
            child.clearParentPtr ObjPtr(Me)
            child.setScreenPointer Nothing
        End If
    Next
    Erase childWidgets
    childCount_ = 0
End Sub

Public Sub setCalcLayout(ByVal calcLayoutFunc As WzFunc)
    Set calcLayout = calcLayoutFunc
    callCalcLayout
End Sub

Public Sub callCalcLayout()
    Dim self As WIDGET
    If Not calcLayout Is Nothing Then
        Set self = Me
        WzFunc_call calcLayout, self
    End If
    ' (#ifdef DEBUG: an else branch, commented out)
End Sub

Public Sub setOnDelete(ByVal onDeleteFunc As WzFunc)
    Set onDelete = onDeleteFunc
End Sub

Public Sub setCustomHitTest(ByVal newCustomHitTestFunc As WzFunc)
    Set customHitTest = newCustomHitTestFunc
End Sub

' void attach(const std::shared_ptr<WIDGET> &widget, ChildZPos zPos = ChildZPos::Front) (and the deprecated
' attach(WIDGET *))
Public Sub attach(ByVal widget As WIDGET, Optional ByVal zPos As ChildZPos = ChildZPos_Front)
    Dim i As Long
    If widget Is Nothing Then ASSERT_FAILURE "widget != nullptr && widget->parentWidget.expired()", "lib/widget/widget.cpp:600", "attach", "Bad attach.": Exit Sub
    If Not (widget.parent() Is Nothing) Then ASSERT_FAILURE "widget != nullptr && widget->parentWidget.expired()", "lib/widget/widget.cpp:600", "attach", "Bad attach.": Exit Sub
    widget.setParentPtr ObjPtr(Me)
    widget.setScreenPointer screenPointer
    If childCount_ = 0 Then
        ReDim childWidgets(0 To 3)
    ElseIf childCount_ > UBound(childWidgets) Then
        ReDim Preserve childWidgets(0 To 2 * childCount_ - 1)
    End If
    Select Case zPos
    Case ChildZPos_Front
        ' insert at the end of the list of children, and thus the top of the z-order (childWidgets order is bottom -> top)
        Set childWidgets(childCount_) = widget
    Case ChildZPos_Back
        ' insert at the very beginning of the list of children
        For i = childCount_ To 1 Step -1
            Set childWidgets(i) = childWidgets(i - 1)
        Next
        Set childWidgets(0) = widget
    End Select
    childCount_ = childCount_ + 1
End Sub

' void detach(const std::shared_ptr<WIDGET> &widget) (and the deprecated detach(WIDGET *))
Public Sub detach(ByVal widget As WIDGET)
    Dim it As Long, i As Long
    If widget Is Nothing Then ASSERT_FAILURE "widget != nullptr && !widget->parentWidget.expired()", "lib/widget/widget.cpp:618", "detach", "Bad detach.": Exit Sub
    If widget.parent() Is Nothing Then ASSERT_FAILURE "widget != nullptr && !widget->parentWidget.expired()", "lib/widget/widget.cpp:618", "detach", "Bad detach.": Exit Sub

    widget.setParentPtr 0
    widget.setScreenPointer Nothing

    it = -1
    For i = 0 To childCount_ - 1
        If childWidgets(i) Is widget Then
            it = i
            Exit For
        End If
    Next
    If it <> -1 Then
        For i = it To childCount_ - 2
            Set childWidgets(i) = childWidgets(i + 1)
        Next
        childCount_ = childCount_ - 1
        Set childWidgets(childCount_) = Nothing
    End If

    widgetLost widget
End Sub

' Set screen pointer for us and all children.
Friend Sub setScreenPointer(ByVal screen As W_SCREEN)
    Dim lockedScreen As W_SCREEN, i As Long
    If screenPointer_ = ObjPtr(screen) Then
        Exit Sub
    End If

    If isMouseOverWidget() Then
        psMouseOverWidget_reset
    End If

    Set lockedScreen = screenPointer
    If Not lockedScreen Is Nothing Then
        If lockedScreen.hasFocus(Me) Then
            lockedScreen.psFocusReset
            Me.focusLost
        End If
        If lockedScreen.isLastHighlight(Me) Then
            lockedScreen.lastHighlightReset
            Me.highlightLost
        End If
    End If

    screenPointer_ = ObjPtr(screen)
    For i = 0 To childCount_ - 1
        childWidgets(i).setScreenPointer screen
    Next
End Sub

Public Function isMouseOverWidget() As Boolean
    isMouseOverWidget = (psMouseOverWidgetPtr = ObjPtr(Me))
End Function

Public Sub setTransparentToClicks(ByVal hasClickTransparency As Boolean)
    isTransparentToClicks = hasClickTransparency
End Sub

Public Sub setTransparentToMouse(ByVal hasMouseTransparency As Boolean)
    isTransparentToMouse = hasMouseTransparency
End Sub

Public Function transparentToClicks() As Boolean
    transparentToClicks = isTransparentToClicks
End Function

Friend Sub processCallbacksRecursive(psContext As W_CONTEXT)
    Dim kids() As WIDGET, n As Long, i As Long, psCurr As WIDGET, sFormContext As W_CONTEXT
    ' Go through all the widgets on the form
    n = childrenCopy(kids)
    For i = 0 To n - 1
        Set psCurr = kids(i)
        ' Call the callback
        If Not psCurr.callback Is Nothing Then
            WzFunc_call psCurr.callback, psCurr, psContext.xOffset, psContext.yOffset, psContext.mx, psContext.my
        End If

        ' and then recurse
        sFormContext = psContext
        sFormContext.mx = psContext.mx - psCurr.x()
        sFormContext.my = psContext.my - psCurr.y()
        sFormContext.xOffset = psContext.xOffset + psCurr.x()
        sFormContext.yOffset = psContext.yOffset + psCurr.y()
        psCurr.processCallbacksRecursive sFormContext
    Next
End Sub

Friend Sub manuallyCallRun(psContext As W_CONTEXT)
    run psContext
End Sub

Friend Function processClick(psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim psClickedWidget As WIDGET, clickedContext As W_CONTEXT, shiftX As Long, shiftY As Long
    Set psClickedWidget = Me
    clickedContext = psContext
    If isTransparentToClicks Then
        Do
            Set psClickedWidget = psClickedWidget.parent()
            If psClickedWidget Is Nothing Then Exit Do  ' (VB) C++ reads the missing parent's x() here
            shiftX = psClickedWidget.x()
            shiftY = psClickedWidget.y()
            clickedContext.mx = clickedContext.mx + shiftX
            clickedContext.my = clickedContext.my + shiftY
            clickedContext.xOffset = clickedContext.xOffset - shiftX
            clickedContext.yOffset = clickedContext.yOffset - shiftY
        Loop While psClickedWidget.transparentToClicks()
        If psClickedWidget Is Nothing Then
            Exit Function
        End If
    End If
    If wasPressed Then
        psClickedWidget.clicked clickedContext, key
    Else
        psClickedWidget.released clickedContext, key
    End If
    psContext = clickedContext

    Set processClick = psClickedWidget
End Function

'--------------------------------------------------------------------------
' (VB) the virtual methods' defaults that need the private members (widget_c's WIDGET_<method> calls them)

Friend Sub def_screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim kids() As WIDGET, n As Long, i As Long
    ' Default implementation of screenSizeDidChange calls its own calcLayout callback function (if present)
    callCalcLayout

    ' Then propagates the event to all children
    n = childrenCopy(kids)
    For i = 0 To n - 1
        kids(i).screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
    Next
End Sub

' Process all the widgets on a form.
' mx and my are the coords of the mouse relative to the form origin.
Friend Sub def_runRecursive(psContext As W_CONTEXT)
    Dim kids() As WIDGET, n As Long, i As Long, psCurr As WIDGET, sFormContext As W_CONTEXT
    ' Process the form's widgets
    n = childrenCopy(kids)
    For i = 0 To n - 1
        Set psCurr = kids(i)
        ' Skip any hidden widgets
        If psCurr.visible() Then
            ' Found a sub form, so set up the context
            sFormContext = psContext
            sFormContext.mx = psContext.mx - psCurr.x()
            sFormContext.my = psContext.my - psCurr.y()
            sFormContext.xOffset = psContext.xOffset + psCurr.x()
            sFormContext.yOffset = psContext.yOffset + psCurr.y()

            ' Process it
            psCurr.runRecursive sFormContext

            ' Run the widget
            psCurr.run psContext
        End If
    Next
End Sub

Friend Function def_hitTest(ByVal x As Long, ByVal y As Long) As Boolean
    Dim hitTestResult As Boolean, self As WIDGET
    ' default hit-testing bounding rect (based on the widget's x, y, width, height)
    hitTestResult = WzRect_contains(dim_, x, y)

    If Not customHitTest Is Nothing Then
        ' if the default bounding-rect hit-test succeeded, use the custom hit-testing func
        Set self = Me
        If hitTestResult Then hitTestResult = CBool(WzFunc_call(customHitTest, self, x, y))
    End If

    def_hitTest = hitTestResult
End Function

Friend Function def_findMouseTargetRecursive(psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim widgProcessedClick As WIDGET, shiftedContext As W_CONTEXT, skipProcessingChildren As Boolean, i As Long, psCurr As WIDGET

    shiftedContext = psContext
    shiftedContext.mx = psContext.mx - x()
    shiftedContext.my = psContext.my - y()
    shiftedContext.xOffset = psContext.xOffset + x()
    shiftedContext.yOffset = psContext.yOffset + y()

    skipProcessingChildren = False
    If typ = WIDG_FORM Then
        If formDisablesChildren() Then
            skipProcessingChildren = True
        End If
    End If

    If Not skipProcessingChildren Then
        ' Process subwidgets.
        ' (Enumerate the child widgets in decreasing z-order (i.e. "top-down"))
        i = childCount_
        Do While i > 0
            i = i - 1
            If i < childCount_ Then
                Set psCurr = childWidgets(i)
                If psCurr.visible() Then
                    If psCurr.hitTest(shiftedContext.mx, shiftedContext.my) Then
                        ' Process it (recursively).
                        Set widgProcessedClick = psCurr.findMouseTargetRecursive(shiftedContext, key, wasPressed)
                        If Not widgProcessedClick Is Nothing Then
                            psContext = shiftedContext ' bubble-up the context for the hit
                            Set def_findMouseTargetRecursive = widgProcessedClick
                            Exit Function
                        End If
                    End If
                End If  ' (else: Skip any hidden widgets, or widgets the click missed.)
            End If
        Loop
    End If

    If Not visible() Then
        Exit Function
    End If

    If isTransparentToMouse Then
        Exit Function
    End If

    Set def_findMouseTargetRecursive = Me
End Function

Friend Sub def_displayRecursive(context As WidgetGraphicsContext)
    Dim widgetIsClipped As Boolean, col As Long, g As WzRect, childrenContext As WidgetGraphicsContext, translated As WidgetGraphicsContext
    Dim kids() As WIDGET, n As Long, i As Long, self As WIDGET
    g = dim_
    widgetIsClipped = Not WidgetGraphicsContext_clipContains(context, g)

    If Not widgetIsClipped Then
        If widgDebugBoundingBoxesOnly() Then
            ' Display bounding boxes.
            col = pal_RGBA(128 + iSinSR(realTime, 2000, 127), 128 + iSinSR(realTime + 667, 2000, 127), 128 + iSinSR(realTime + 1333, 2000, 127), 128)
            iV_Box WidgetGraphicsContext_getXOffset(context) + x(), WidgetGraphicsContext_getYOffset(context) + y(), WidgetGraphicsContext_getXOffset(context) + x() + width() - 1, WidgetGraphicsContext_getYOffset(context) + y() + height() - 1, col
        ElseIf Not displayFunction Is Nothing Then
            Set self = Me
            WzFunc_call displayFunction, self, WidgetGraphicsContext_getXOffset(context), WidgetGraphicsContext_getYOffset(context)
        Else
            ' Display widget.
            display WidgetGraphicsContext_getXOffset(context), WidgetGraphicsContext_getYOffset(context)
        End If
    End If

    If widgetIsClipped And Not WidgetGraphicsContext_allowChildDisplayRecursiveIfSelfClipped(context) Then
        Exit Sub
    End If

    If typ = WIDG_FORM Then
        If formDisablesChildren() Then
            Exit Sub
        End If
    End If

    translated = WidgetGraphicsContext_translatedBy(context, x(), y())
    childrenContext = WidgetGraphicsContext_setAllowChildDisplayRecursiveIfSelfClipped(translated, False)

    ' If this is a clickable form, the widgets on it have to move when it's down.
    If typ = WIDG_FORM And (style And WFORM_NOCLICKMOVE) = 0 Then
        If (style And WFORM_CLICKABLE) <> 0 And WIDGET_isKindOf(Me, K_W_CLICKFORM) Then
            If (W_CLICKFORM_(Me).state And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0 Then
                translated = childrenContext
                childrenContext = WidgetGraphicsContext_translatedBy(translated, 1, 1)
            End If
        End If
    End If

    ' Display the widgets on this widget.
    ' NOTE: Draw them in the opposite order that clicks are processed, so behavior matches the visual.
    ' i.e. Since findMouseTargetRecursive handles processing children in decreasing z-order (i.e. "top-down")
    '      we want to draw things in list order (bottom-up) so the "top-most" is drawn last
    n = childrenCopy(kids)
    For i = 0 To n - 1
        If kids(i).visible() Then
            kids(i).displayRecursive childrenContext
        End If
    Next
End Sub

'--------------------------------------------------------------------------
' (VB) the object model's helpers

' type == WIDG_FORM && ((W_FORM *)this)->disableChildren (the cast checked: a WIDG_FORM that is no W_FORM has none)
Private Function formDisablesChildren() As Boolean
    If WIDGET_isKindOf(Me, K_W_FORM) Then formDisablesChildren = W_FORM_(Me).disableChildren
End Function

' a copy of the children (C++'s range-for over childWidgets), returns their number
Private Function childrenCopy(out() As WIDGET) As Long
    Dim i As Long
    If childCount_ = 0 Then Exit Function
    ReDim out(0 To childCount_ - 1)
    For i = 0 To childCount_ - 1
        Set out(i) = childWidgets(i)
    Next
    childrenCopy = childCount_
End Function

' widget->parentWidget = shared_from_this() / parentWidget.reset() (0)
Friend Sub setParentPtr(ByVal p As Long)
    parentWidget = p
End Sub

' the parent p dies or lets this child go
Friend Sub clearParentPtr(ByVal p As Long)
    If parentWidget = p Then parentWidget = 0
End Sub

' psForm->screenPointer = shared_from_this() (W_SCREEN::initialize: the root form, without setScreenPointer)
Friend Sub assignScreenPointer(ByVal screen As W_SCREEN)
    screenPointer_ = ObjPtr(screen)
End Sub

' the screen dies: its widgets' weak screen pointers expire (without locking the dying screen)
Friend Sub clearScreenPointerRecursive()
    Dim i As Long
    screenPointer_ = 0
    For i = 0 To childCount_ - 1
        childWidgets(i).clearScreenPointerRecursive
    Next
End Sub
