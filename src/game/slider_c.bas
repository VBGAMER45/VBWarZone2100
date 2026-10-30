Attribute VB_Name = "slider_c"
' lib/widget/slider.cpp, slider.h: the widget class W_SLIDER (docs/CONVENTIONS.md,
' "Widgets"; its fields are W_SLIDER.cls), W_SLDINIT's constructor (the struct is
' widget_c's) and widgGetSliderPos. A SliderOnChangeFunc is a WzFunc called with
' the slider; optional<WidgetHelp> is a WidgetHelp (Nothing = no value).
' (VB) widget class W_SLIDER : WIDGET
Option Explicit

Private Enum SliderState
    ' Slider is being dragged
    SLD_DRAG = 1                        ' 1 << 0

    ' Slider is hilited
    SLD_HILITE = 2                      ' 1 << 1

    ' Slider is disabled
    SLD_DISABLED = 4                    ' 1 << 2
End Enum

' W_SLDINIT::W_SLDINIT()
Public Sub W_SLDINIT_construct(init As W_SLDINIT)
    Dim z As W_SLDINIT
    init = z
    init.orientation = WSLD_LEFT
    init.numStops = 0
    init.barSize = 0
    init.pos = 0
End Sub

' (VB) the W_INIT part of a W_SLDINIT
Public Function W_SLDINIT_toW_INIT(init As W_SLDINIT) As W_INIT
    Dim r As W_INIT
    r.formID = init.formID
    r.majorID = init.majorID
    r.id = init.id
    r.style = init.style
    r.x = init.x
    r.y = init.y
    r.width = init.width
    r.height = init.height
    Set r.pDisplay = init.pDisplay
    Set r.pCallback = init.pCallback
    WzFunc_assign r.pUserData, init.pUserData
    r.UserData = init.UserData
    Set r.calcLayout = init.calcLayout
    Set r.onDelete = init.onDelete
    Set r.customHitTest = init.customHitTest
    Set r.initPUserDataFunc = init.initPUserDataFunc
    W_SLDINIT_toW_INIT = r
End Function

' std::make_shared<W_SLIDER>(init)
Public Function W_SLIDER_new_init(init As W_SLDINIT) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_SLIDER)
    W_SLIDER_construct_init w, init
    Set W_SLIDER_new_init = w
End Function

Public Sub W_SLIDER_construct_init(ByVal self As WIDGET, init As W_SLDINIT)
    Dim d As W_SLIDER, baseInit As W_INIT, horizontal As Boolean, vertical As Boolean
    baseInit = W_SLDINIT_toW_INIT(init)
    WIDGET_construct_init self, baseInit, WIDG_SLIDER
    self.kind = K_W_SLIDER ' (VB) virtual calls from here on reach W_SLIDER's overrides
    Set d = W_SLIDER_(self)
    d.orientation = init.orientation
    d.numStops = init.numStops
    d.barSize = init.barSize
    d.pos = init.pos
    d.state = 0
    d.pTip = init.pTip
    d.isHandlingDrag = False

    If Not ((init.style And Not (WBAR_PLAIN Or WIDG_HIDDEN)) = 0) Then ASSERT_FAILURE "(init->style & ~(0 | WIDG_HIDDEN)) == 0", "lib/widget/slider.cpp:57", "W_SLIDER", "Unknown style"
    If Not (init.orientation >= WSLD_LEFT Or init.orientation <= WSLD_BOTTOM) Then ASSERT_FAILURE "init->orientation >= WSLD_LEFT || init->orientation <= WSLD_BOTTOM", "lib/widget/slider.cpp:62", "W_SLIDER", "Unknown orientation"
    horizontal = init.orientation = WSLD_LEFT Or init.orientation = WSLD_RIGHT
    vertical = init.orientation = WSLD_TOP Or init.orientation = WSLD_BOTTOM
    If Not ((Not horizontal Or init.numStops <= init.width - init.barSize) And (Not vertical Or init.numStops <= init.height - init.barSize)) Then ASSERT_FAILURE "(!horizontal || init->numStops <= init->width  - init->barSize) && (!vertical   || init->numStops <= init->height - init->barSize)", "lib/widget/slider.cpp:68", "W_SLIDER", "Too many stops for slider length"
    If Not (init.pos <= init.numStops) Then ASSERT_FAILURE "init->pos <= init->numStops", "lib/widget/slider.cpp:70", "W_SLIDER", "Slider position greater than stops (%d/%d)", init.pos, init.numStops
    If Not ((Not horizontal Or init.barSize <= init.width) And (Not vertical Or init.barSize <= init.height)) Then ASSERT_FAILURE "(!horizontal || init->barSize <= init->width) && (!vertical   || init->barSize <= init->height)", "lib/widget/slider.cpp:71", "W_SLIDER", "Slider bar is larger than slider width"
End Sub

' Get the current position of a slider bar
Public Function widgGetSliderPos(ByVal psScreen As W_SCREEN, ByVal id As Long) As Long
    Dim psWidget As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/slider.cpp:78", "widgGetSliderPos", "Invalid screen pointer": Exit Function

    Set psWidget = widgGetFromID(psScreen, id)
    If Not (Not psWidget Is Nothing) Then ASSERT_FAILURE "psWidget != nullptr", "lib/widget/slider.cpp:82", "widgGetSliderPos", "Could not find widget from id %u", id
    If Not psWidget Is Nothing Then
        widgGetSliderPos = W_SLIDER_(psWidget).pos
        Exit Function
    End If
    widgGetSliderPos = 0
End Function

' Run a slider widget
Public Sub W_SLIDER_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_SLIDER, lockedScreen As W_SCREEN
    Set d = W_SLIDER_(self)
    If (d.state And SLD_DRAG) <> 0 And Not d.isHandlingDrag Then
        d.state = d.state And Not SLD_DRAG
        Set lockedScreen = self.screenPointer
        If Not lockedScreen Is Nothing Then
            lockedScreen.setReturn self
        End If
        self.dirty = True
    End If
End Sub

Public Sub W_SLIDER_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_SLIDER
    If W_SLIDER_isEnabled(self) Then
        Set d = W_SLIDER_(self)
        self.dirty = True
        d.state = d.state Or SLD_DRAG
        d.isHandlingDrag = True
        W_SLIDER_updateSliderFromMousePosition self, psContext
    End If
End Sub

Public Sub W_SLIDER_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    W_SLIDER_(self).isHandlingDrag = False
End Sub

Public Sub W_SLIDER_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    W_SLIDER_(self).state = W_SLIDER_(self).state Or SLD_HILITE
    self.dirty = True
End Sub

' Respond to the mouse moving off a slider
Public Sub W_SLIDER_highlightLost(ByVal self As WIDGET)
    W_SLIDER_(self).state = W_SLIDER_(self).state And Not SLD_HILITE
    self.dirty = True
End Sub

Public Sub W_SLIDER_setTip(ByVal self As WIDGET, ByVal string_ As String)
    W_SLIDER_(self).pTip = string_
End Sub

Public Sub W_SLIDER_setHelp(ByVal self As WIDGET, ByVal help As WidgetHelp)
    Set W_SLIDER_(self).help = help
End Sub

Public Sub W_SLIDER_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ASSERT_FAILURE "false", "lib/widget/slider.cpp:146", "display", "No default implementation exists for sliders"
End Sub

Public Sub W_SLIDER_enable(ByVal self As WIDGET)
    W_SLIDER_(self).state = W_SLIDER_(self).state And Not SLD_DISABLED
End Sub

Public Sub W_SLIDER_disable(ByVal self As WIDGET)
    W_SLIDER_(self).state = W_SLIDER_(self).state Or SLD_DISABLED
End Sub

Public Function W_SLIDER_isHighlighted(ByVal self As WIDGET) As Boolean
    W_SLIDER_isHighlighted = (W_SLIDER_(self).state And SLD_HILITE) <> 0
End Function

Public Function W_SLIDER_isEnabled(ByVal self As WIDGET) As Boolean
    W_SLIDER_isEnabled = (W_SLIDER_(self).state And SLD_DISABLED) = 0
End Function

Public Function W_SLIDER_getTip(ByVal self As WIDGET) As String
    W_SLIDER_getTip = W_SLIDER_(self).pTip
End Function

Public Function W_SLIDER_getHelp(ByVal self As WIDGET) As WidgetHelp
    If W_SLIDER_(self).help Is Nothing Then Exit Function
    Set W_SLIDER_getHelp = W_SLIDER_(self).help
End Function

Public Function W_SLIDER_capturesMouseDrag(ByVal self As WIDGET, ByVal key As WIDGET_KEY) As Boolean
    W_SLIDER_capturesMouseDrag = True
End Function

Public Sub W_SLIDER_mouseDragged(ByVal self As WIDGET, ByVal key As WIDGET_KEY, start As W_CONTEXT, current As W_CONTEXT)
    If W_SLIDER_isEnabled(self) Then
        W_SLIDER_updateSliderFromMousePosition self, current
    End If
End Sub

Private Sub W_SLIDER_updateSliderFromMousePosition(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_SLIDER, mx As Long, my As Long, stopSize As Long, prevPos As Long
    Dim barSize As Long, numStops As Long
    Set d = W_SLIDER_(self)
    ' Figure out where the drag box should be
    mx = psContext.mx - self.x()
    my = psContext.my - self.y()
    prevPos = d.pos
    barSize = d.barSize
    numStops = d.numStops
    Select Case d.orientation
    Case WSLD_LEFT
        If mx <= barSize \ 2 Then
            d.pos = 0
        ElseIf mx >= self.width() - barSize \ 2 Then
            d.pos = numStops
        Else
            ' Mouse is in the middle of the slider, calculate which stop
            stopSize = (self.width() - barSize) \ MAX_(numStops, 1)
            d.pos = ((mx + stopSize \ 2 - barSize \ 2) * numStops \ MAX_((self.width() - barSize), 1)) And &HFFFF&
        End If
    Case WSLD_RIGHT
        If mx <= barSize \ 2 Then
            d.pos = numStops
        ElseIf mx >= self.width() - barSize \ 2 Then
            d.pos = 0
        Else
            ' Mouse is in the middle of the slider, calculate which stop
            stopSize = (self.width() - barSize) \ MAX_(numStops, 1)
            d.pos = (numStops - (mx + stopSize \ 2 - barSize \ 2) * numStops \ MAX_((self.width() - barSize), 1)) And &HFFFF&
        End If
    Case WSLD_TOP
        If my <= barSize \ 2 Then
            d.pos = 0
        ElseIf my >= self.height() - barSize \ 2 Then
            d.pos = numStops
        Else
            ' Mouse is in the middle of the slider, calculate which stop
            stopSize = (self.height() - barSize) \ MAX_(numStops, 1)
            d.pos = ((my + stopSize \ 2 - barSize \ 2) * numStops \ MAX_((self.height() - barSize), 1)) And &HFFFF&
        End If
    Case WSLD_BOTTOM
        If my <= barSize \ 2 Then
            d.pos = numStops
        ElseIf my >= self.height() - barSize \ 2 Then
            d.pos = 0
        Else
            ' Mouse is in the middle of the slider, calculate which stop
            stopSize = (self.height() - barSize) \ MAX_(numStops, 1)
            d.pos = (numStops - (my + stopSize \ 2 - barSize \ 2) * numStops \ MAX_((self.height() - barSize), 1)) And &HFFFF&
        End If
    End Select
    If prevPos <> d.pos Then
        W_SLIDER_callOnChangeFuncs self
    End If
End Sub

Public Sub W_SLIDER_addOnChange(ByVal self As WIDGET, ByVal func As WzFunc)
    If Not (Not func Is Nothing) Then ASSERT_FAILURE "func != nullptr", "lib/widget/slider.cpp:285", "addOnChange", "Null func": Exit Sub
    W_SLIDER_(self).onChangeFuncs.Add func
End Sub

Private Sub W_SLIDER_callOnChangeFuncs(ByVal self As WIDGET)
    Dim f As Variant
    For Each f In W_SLIDER_(self).onChangeFuncs
        WzFunc_call f, self
    Next
End Sub
