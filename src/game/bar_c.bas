Attribute VB_Name = "bar_c"
' lib/widget/bar.cpp, bar.h: the bar graph widget W_BARGRAPH (its fields:
' W_BARGRAPH.cls), W_BARINIT's constructor, widgSetBarSize,
' widgSetMinorBarSize, widgSetBarRange and the display functions
' (barGraphDisplayTrough is also a WIDGET_DISPLAY: FN_barGraphDisplayTrough).
' The UWORD fields are masked to 16 bits where C++ narrows.
' (VB) widget class W_BARGRAPH : WIDGET
Option Explicit

' WzFunc ids (fork W1: 1300-1399 here)
Public Const FN_barGraphDisplayTrough As Long = 1300     ' barGraphDisplayTrough as a WIDGET_DISPLAY

' the function pointers of this module
Public Function bar_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_barGraphDisplayTrough
        barGraphDisplayTrough args(0), args(1), args(2)
    End Select
End Function

' W_BARINIT::W_BARINIT()
Public Sub W_BARINIT_construct(i As W_BARINIT)
    Dim z As W_BARINIT
    i = z
    i.orientation = WBAR_LEFT
    i.size = 0
    i.minorSize = 0
    i.iRange = 100
    i.denominator = 1
    i.precision = 0
    ' sCol
    ' sMinorCol
    i.sCol = 0                              ' sCol.clear();
    i.sMinorCol = 0                         ' sMinorCol.clear();
End Sub

' (VB) the W_INIT part of a W_BARINIT (W_BARINIT : W_INIT)
Public Function W_BARINIT_toW_INIT(init As W_BARINIT) As W_INIT
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
    If IsObject(init.pUserData) Then Set r.pUserData = init.pUserData Else r.pUserData = init.pUserData
    r.UserData = init.UserData
    Set r.calcLayout = init.calcLayout
    Set r.onDelete = init.onDelete
    Set r.customHitTest = init.customHitTest
    Set r.initPUserDataFunc = init.initPUserDataFunc
    W_BARINIT_toW_INIT = r
End Function

' W_BARGRAPH::W_BARGRAPH(W_BARINIT const *init)
Public Sub W_BARGRAPH_construct_init(ByVal self As WIDGET, init As W_BARINIT)
    Dim baseInit As W_INIT, d As W_BARGRAPH
    baseInit = W_BARINIT_toW_INIT(init)
    WIDGET_construct_init self, baseInit, WIDG_BARGRAPH
    self.kind = K_W_BARGRAPH ' (VB) virtual calls from here on reach W_BARGRAPH's overrides
    Set d = W_BARGRAPH_(self)
    d.barPos = init.orientation
    d.majorSize = init.size And &HFFFF&
    d.minorSize = init.minorSize And &HFFFF&
    d.iRange = init.iRange And &HFFFF&
    d.majorValue = (d.majorSize * d.iRange \ WBAR_SCALE) And &HFFFF&
    d.minorValue = (d.minorSize * d.iRange \ WBAR_SCALE) And &HFFFF&
    d.denominator = MAX_(init.denominator, 1)
    d.precision_ = init.precision
    d.majorCol = init.sCol
    d.minorCol = init.sMinorCol
    d.textCol = psPalette(WZCOL_BLACK)
    d.pTip = init.pTip
    d.backgroundColour = psPalette(WZCOL_FORM_BACKGROUND)
    d.sizesDirty = False

    ' Set the minor colour if necessary
    ' Actually, this sets the major colour to the minor colour. The minor colour used to be left completely uninitialised... Wonder what it was for..?
    If (self.style And WBAR_DOUBLE) <> 0 Then
        d.majorCol = d.minorCol
    End If

    If Not ((init.style And Not (WBAR_PLAIN Or WBAR_TROUGH Or WBAR_DOUBLE Or WIDG_HIDDEN)) = 0) Then ASSERT_FAILURE "(init->style & ~(0 | 1 | 2 | WIDG_HIDDEN)) == 0", "lib/widget/bar.cpp:69", "W_BARGRAPH", "Unknown bar graph style"
    ' (#pragma GCC diagnostic push)
    If Not (init.orientation >= WBAR_LEFT Or init.orientation <= WBAR_BOTTOM) Then ASSERT_FAILURE "init->orientation >= WBAR_LEFT || init->orientation <= WBAR_BOTTOM", "lib/widget/bar.cpp:74", "W_BARGRAPH", "Unknown orientation"
    ' (#pragma GCC diagnostic pop)
    If Not (init.size <= WBAR_SCALE) Then ASSERT_FAILURE "init->size <= 100", "lib/widget/bar.cpp:78", "W_BARGRAPH", "Bar size out of range"
    If Not ((init.style And WBAR_DOUBLE) = 0 Or init.minorSize <= WBAR_SCALE) Then ASSERT_FAILURE "(init->style & 2) == 0 || init->minorSize <= 100", "lib/widget/bar.cpp:79", "W_BARGRAPH", "Minor bar size out of range"
End Sub

' std::make_shared<W_BARGRAPH>(init)
Public Function W_BARGRAPH_new_init(init As W_BARINIT) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_BARGRAPH)
    W_BARGRAPH_construct_init w, init
    Set W_BARGRAPH_new_init = w
End Function

Private Function getBarGraph(ByVal psScreen As W_SCREEN, ByVal id As Long) As WIDGET
    Dim psBGraph As WIDGET
    If Not (Not psScreen Is Nothing) Then ASSERT_FAILURE "psScreen != nullptr", "lib/widget/bar.cpp:84", "getBarGraph", "Invalid screen pointer": Exit Function
    Set psBGraph = widgGetFromID(psScreen, id)
    If Not (Not psBGraph Is Nothing) Then ASSERT_FAILURE "psBGraph != nullptr", "lib/widget/bar.cpp:86", "getBarGraph", "Could not find widget from ID": Exit Function
    If Not (psBGraph.typ = WIDG_BARGRAPH) Then ASSERT_FAILURE "psBGraph->type == WIDG_BARGRAPH", "lib/widget/bar.cpp:87", "getBarGraph", "Wrong widget type": Exit Function

    Set getBarGraph = psBGraph
End Function

' Set the current size of a bar graph
Public Sub widgSetBarSize(ByVal psScreen As W_SCREEN, ByVal id As Long, ByVal iValue As Long)
    Dim psBGraph As WIDGET
    Set psBGraph = getBarGraph(psScreen, id)
    If Not psBGraph Is Nothing Then
        W_BARGRAPH_(psBGraph).majorValue = iValue And &HFFFF&
        W_BARGRAPH_(psBGraph).sizesDirty = True
    End If
End Sub

' Set the current size of a minor bar on a double graph
Public Sub widgSetMinorBarSize(ByVal psScreen As W_SCREEN, ByVal id As Long, ByVal iValue As Long)
    Dim psBGraph As WIDGET
    Set psBGraph = getBarGraph(psScreen, id)
    If Not psBGraph Is Nothing Then
        W_BARGRAPH_(psBGraph).minorValue = iValue And &HFFFF&
        W_BARGRAPH_(psBGraph).sizesDirty = True
    End If
End Sub

' Set the range on a double graph
Public Sub widgSetBarRange(ByVal psScreen As W_SCREEN, ByVal id As Long, ByVal iValue As Long)
    Dim psBGraph As WIDGET
    Set psBGraph = getBarGraph(psScreen, id)
    If Not psBGraph Is Nothing Then
        W_BARGRAPH_(psBGraph).iRange = iValue And &HFFFF&
        W_BARGRAPH_(psBGraph).sizesDirty = True
    End If
End Sub

Public Sub W_BARGRAPH_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_BARGRAPH
    Set d = W_BARGRAPH_(self)
    If d.sizesDirty Then
        d.majorSize = (WBAR_SCALE * MIN_(d.majorValue, d.iRange) \ MAX_(d.iRange, 1)) And &HFFFF&
        d.minorSize = MIN_(WBAR_SCALE * d.minorValue \ MAX_(d.iRange, 1), WBAR_SCALE) And &HFFFF&
        d.sizesDirty = False
    End If
End Sub

Private Sub barGraphDisplayText(ByVal barGraph As WIDGET, ByVal x0 As Long, ByVal x1 As Long, ByVal y1 As Long)
    Dim d As W_BARGRAPH, textWidth As Long, pos As Vector2i, dx As Long, dy As Long
    Set d = W_BARGRAPH_(barGraph)
    If Len(d.text) <> 0 Then
        d.wzCachedText.setText d.text, font_bar
        textWidth = d.wzCachedText.width()
        pos = Vector2i_((x0 + x1 - textWidth) \ 2, y1)
        ' Add a shadow, to make it visible against any background.
        For dx = -2 To 2
            For dy = -2 To 2
                If dx * dx + dy * dy <= 4 Then
                    d.wzCachedText.render pos.x + dx, pos.y + dy, psPalette(WZCOL_BLACK)
                End If
            Next
        Next
        d.wzCachedText.render pos.x, pos.y, d.textCol
    End If
End Sub

' The simple bar graph display function
Private Sub barGraphDisplay(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, psBGraph As W_BARGRAPH

    Set psBGraph = W_BARGRAPH_(psWidget)

    ' figure out which way the bar graph fills
    Select Case psBGraph.barPos
    Case WBAR_LEFT
        x0 = xOffset + psWidget.x()
        y0 = yOffset + psWidget.y()
        x1 = x0 + psWidget.width() * psBGraph.majorSize \ WBAR_SCALE
        y1 = y0 + psWidget.height()
    Case WBAR_RIGHT
        y0 = yOffset + psWidget.y()
        x1 = xOffset + psWidget.x() + psWidget.width()
        x0 = x1 - psWidget.width() * psBGraph.majorSize \ WBAR_SCALE
        y1 = y0 + psWidget.height()
    Case WBAR_TOP
        x0 = xOffset + psWidget.x()
        y0 = yOffset + psWidget.y()
        x1 = x0 + psWidget.width()
        y1 = y0 + psWidget.height() * psBGraph.majorSize \ WBAR_SCALE
    Case WBAR_BOTTOM
        x0 = xOffset + psWidget.x()
        x1 = x0 + psWidget.width()
        y1 = yOffset + psWidget.y() + psWidget.height()
        y0 = y1 - psWidget.height() * psBGraph.majorSize \ WBAR_SCALE
    End Select

    ' Now draw the graph
    iV_ShadowBox x0, y0, x1, y1, 0, psPalette(WZCOL_FORM_LIGHT), psPalette(WZCOL_FORM_DARK), psBGraph.majorCol

    barGraphDisplayText psWidget, x0, x1, y1
End Sub

' The double bar graph display function
Private Sub barGraphDisplayDouble(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long
    Dim psBGraph As W_BARGRAPH
    Set psBGraph = W_BARGRAPH_(psWidget)

    ' figure out which way the bar graph fills
    Select Case psBGraph.barPos
    Case WBAR_LEFT
        ' Calculate the major bar
        x0 = xOffset + psWidget.x()
        y0 = yOffset + psWidget.y()
        x1 = x0 + psWidget.width() * psBGraph.majorSize \ WBAR_SCALE
        y1 = y0 + 2 * psWidget.height() \ 3

        ' Calculate the minor bar
        x2 = x0
        y2 = y0 + psWidget.height() \ 3
        x3 = x2 + psWidget.width() * psBGraph.minorSize \ WBAR_SCALE
        y3 = y0 + psWidget.height()
    Case WBAR_RIGHT
        ' Calculate the major bar
        y0 = yOffset + psWidget.y()
        x1 = xOffset + psWidget.x() + psWidget.width()
        x0 = x1 - psWidget.width() * psBGraph.majorSize \ WBAR_SCALE
        y1 = y0 + 2 * psWidget.height() \ 3

        ' Calculate the minor bar
        x3 = x1
        y2 = y0 + psWidget.height() \ 3
        x2 = x3 - psWidget.width() * psBGraph.minorSize \ WBAR_SCALE
        y3 = y0 + psWidget.height()
    Case WBAR_TOP
        ' Calculate the major bar
        x0 = xOffset + psWidget.x()
        y0 = yOffset + psWidget.y()
        x1 = x0 + 2 * psWidget.width() \ 3
        y1 = y0 + psWidget.height() * psBGraph.majorSize \ WBAR_SCALE

        ' Calculate the minor bar
        x2 = x0 + psWidget.width() \ 3
        y2 = y0
        x3 = x0 + psWidget.width()
        y3 = y2 + psWidget.height() * psBGraph.minorSize \ WBAR_SCALE
    Case WBAR_BOTTOM
        ' Calculate the major bar
        x0 = xOffset + psWidget.x()
        x1 = x0 + 2 * psWidget.width() \ 3
        y1 = yOffset + psWidget.y() + psWidget.height()
        y0 = y1 - psWidget.height() * psBGraph.majorSize \ WBAR_SCALE

        ' Calculate the minor bar
        x2 = x0 + psWidget.width() \ 3
        x3 = x0 + psWidget.width()
        y3 = y1
        y2 = y3 - psWidget.height() * psBGraph.minorSize \ WBAR_SCALE
    End Select

    ' Draw the minor bar graph
    If psBGraph.minorSize > 0 Then
        iV_ShadowBox x2, y2, x3, y3, 0, psPalette(WZCOL_FORM_LIGHT), psPalette(WZCOL_FORM_DARK), psBGraph.minorCol
    End If

    ' Draw the major bar graph
    iV_ShadowBox x0, y0, x1, y1, 0, psPalette(WZCOL_FORM_LIGHT), psPalette(WZCOL_FORM_DARK), psBGraph.majorCol

    barGraphDisplayText psWidget, x0, x1, y1
End Sub

' The trough bar graph display function
Public Sub barGraphDisplayTrough(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long         ' Position of the bar
    Dim tx0 As Long, ty0 As Long, tx1 As Long, ty1 As Long     ' Position of the trough
    Dim showBar As Boolean, showTrough As Boolean, psBGraph As W_BARGRAPH
    showBar = True
    showTrough = True
    Set psBGraph = W_BARGRAPH_(psWidget)

    ' figure out which way the bar graph fills
    Select Case psBGraph.barPos
    Case WBAR_LEFT
        x0 = xOffset + psWidget.x()
        y0 = yOffset + psWidget.y()
        x1 = x0 + psWidget.width() * psBGraph.majorSize \ WBAR_SCALE
        y1 = y0 + psWidget.height()
        If x0 = x1 Then
            showBar = False
        End If
        tx0 = x1 + 1
        ty0 = y0
        tx1 = x0 + psWidget.width()
        ty1 = y1
        If tx0 >= tx1 Then
            showTrough = False
        End If
    Case WBAR_RIGHT
        y0 = yOffset + psWidget.y()
        x1 = xOffset + psWidget.x() + psWidget.width()
        x0 = x1 - psWidget.width() * psBGraph.majorSize \ WBAR_SCALE
        y1 = y0 + psWidget.height()
        If x0 = x1 Then
            showBar = False
        End If
        tx0 = xOffset + psWidget.x()
        ty0 = y0
        tx1 = x0 - 1
        ty1 = y1
        If tx0 >= tx1 Then
            showTrough = False
        End If
    Case WBAR_TOP
        x0 = xOffset + psWidget.x()
        y0 = yOffset + psWidget.y()
        x1 = x0 + psWidget.width()
        y1 = y0 + psWidget.height() * psBGraph.majorSize \ WBAR_SCALE
        If y0 = y1 Then
            showBar = False
        End If
        tx0 = x0
        ty0 = y1 + 1
        tx1 = x1
        ty1 = y0 + psWidget.height()
        If ty0 >= ty1 Then
            showTrough = False
        End If
    Case WBAR_BOTTOM
        x0 = xOffset + psWidget.x()
        x1 = x0 + psWidget.width()
        y1 = yOffset + psWidget.y() + psWidget.height()
        y0 = y1 - psWidget.height() * psBGraph.majorSize \ WBAR_SCALE
        If y0 = y1 Then
            showBar = False
        End If
        tx0 = x0
        ty0 = yOffset + psWidget.y()
        tx1 = x1
        ty1 = y0 - 1
        If ty0 >= ty1 Then
            showTrough = False
        End If
    End Select

    ' Now draw the graph
    If showBar Then
        pie_BoxFill x0, y0, x1, y1, psBGraph.majorCol
    End If
    If showTrough Then
        iV_ShadowBox tx0, ty0, tx1, ty1, 0, psPalette(WZCOL_FORM_DARK), psPalette(WZCOL_FORM_LIGHT), psBGraph.backgroundColour
    End If

    barGraphDisplayText psWidget, x0, tx1, ty1
End Sub

Public Sub W_BARGRAPH_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    If (self.style And WBAR_TROUGH) <> 0 Then
        barGraphDisplayTrough self, xOffset, yOffset
    ElseIf (self.style And WBAR_DOUBLE) <> 0 Then
        barGraphDisplayDouble self, xOffset, yOffset
    Else
        barGraphDisplay self, xOffset, yOffset
    End If
End Sub

Public Sub W_BARGRAPH_setTip(ByVal self As WIDGET, ByVal string_ As String)
    W_BARGRAPH_(self).pTip = string_
End Sub

Public Sub W_BARGRAPH_setHelp(ByVal self As WIDGET, ByVal help As WidgetHelp)
    Set W_BARGRAPH_(self).help = help
End Sub

Public Function W_BARGRAPH_getTip(ByVal self As WIDGET) As String
    W_BARGRAPH_getTip = W_BARGRAPH_(self).pTip
End Function

' WidgetHelp const * getHelp() const (Nothing if !help.has_value())
Public Function W_BARGRAPH_getHelp(ByVal self As WIDGET) As WidgetHelp
    Set W_BARGRAPH_getHelp = W_BARGRAPH_(self).help
End Function

Public Sub W_BARGRAPH_setBackgroundColour(ByVal self As WIDGET, ByVal colour As Long)
    W_BARGRAPH_(self).backgroundColour = colour
End Sub
