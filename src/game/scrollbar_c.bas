Attribute VB_Name = "scrollbar_c"
' lib/widget/scrollbar.cpp, scrollbar.h: the widget class ScrollBarWidget
' (docs/CONVENTIONS.md, "Widgets"; its fields are ScrollBarWidget.cls).
' ScrollBarWidget::make() is ScrollBarWidget_make. displayScrollBar is the
' WzFunc FN_displayScrollBar (a WIDGET_DISPLAY: widget, xOffset, yOffset).
' (VB) widget class ScrollBarWidget : WIDGET
' (VB) virtual ScrollBarWidget.initialize
Option Explicit

Public Const FN_displayScrollBar As Long = 2000

Public Function scrollbar_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_displayScrollBar
        displayScrollBar args(0), args(1), args(2)
    End Select
End Function

Private Sub displayScrollBar(ByVal widget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim slider As WIDGET, x0 As Long, y0 As Long, sliderY As Long, sliderDrawY0 As Long, sliderDrawY1 As Long
    Dim m As mat4, d As W_SLIDER
    Set slider = widget                 ' (W_SLIDER *)widget
    Set d = W_SLIDER_(slider)

    x0 = slider.x() + xOffset
    y0 = slider.y() + yOffset

    defaultProjectionMatrix VarPtr(m)
    RenderWindowFrame FRAME_NORMAL, x0, y0, slider.width(), slider.height(), VarPtr(m)
    ' pie_UniTransBoxFill(x0, y0, x0 + slider->width(), y0 + slider->height(), WZCOL_DBLUE);

    If d.numStops > 0 Then
        sliderY = (slider.height() - d.barSize) * d.pos \ d.numStops
    Else
        sliderY = 0
    End If
    sliderDrawY0 = y0 + sliderY + 2
    sliderDrawY1 = y0 + sliderY + MAX_(d.barSize, 5) - 2
    If sliderDrawY1 - sliderDrawY0 <= 2 Then
        sliderDrawY0 = sliderDrawY0 - 1
        sliderDrawY1 = sliderDrawY1 + 1
    End If
    If W_SLIDER_isEnabled(slider) Then
        pie_UniTransBoxFill x0 + 2, sliderDrawY0, x0 + slider.width() - 2, sliderDrawY1, psPalette(WZCOL_LBLUE)
    Else
        pie_UniTransBoxFill x0 + 2, sliderDrawY0, x0 + slider.width() - 2, sliderDrawY1, psPalette(WZCOL_FORM_DISABLE)
    End If
End Sub

' static std::shared_ptr<ScrollBarWidget> make()
Public Function ScrollBarWidget_make() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ScrollBarWidget)
    ScrollBarWidget_construct w
    vt_ScrollBarWidget_initialize w
    Set ScrollBarWidget_make = w
End Function

' ScrollBarWidget(): WIDGET() {}
Public Sub ScrollBarWidget_construct(ByVal self As WIDGET)
    Dim d As ScrollBarWidget
    WIDGET_construct self
    self.kind = K_ScrollBarWidget ' (VB) virtual calls from here on reach ScrollBarWidget's overrides
    Set d = ScrollBarWidget_(self)
    d.scrollableSize = 1
    d.viewSize = 1
    d.stickToBottom = False
End Sub

Public Sub ScrollBarWidget_initialize(ByVal self As WIDGET)
    Dim sliderInit As W_SLDINIT, slider As WIDGET
    W_SLDINIT_construct sliderInit
    Set slider = W_SLIDER_new_init(sliderInit)
    Set ScrollBarWidget_(self).slider = slider
    self.attach slider
    W_SLIDER_(slider).numStops = 0
    W_SLIDER_(slider).barSize = 0
    W_SLIDER_(slider).orientation = WSLD_TOP
    Set slider.displayFunction = WzFunc_new(FN_displayScrollBar)
End Sub

Public Sub ScrollBarWidget_geometryChanged(ByVal self As WIDGET)
    ScrollBarWidget_(self).slider.setGeometry 0, 0, self.width(), self.height()
End Sub

Public Function ScrollBarWidget_position(ByVal self As WIDGET) As Long
    ScrollBarWidget_position = W_SLIDER_(ScrollBarWidget_(self).slider).pos
End Function

Public Sub ScrollBarWidget_setPosition(ByVal self As WIDGET, ByVal newPosition As Long)
    Dim s As W_SLIDER
    Set s = W_SLIDER_(ScrollBarWidget_(self).slider)
    s.pos = MIN_(newPosition And &HFFFF&, s.numStops)
End Sub

Public Sub ScrollBarWidget_incrementPosition(ByVal self As WIDGET, ByVal amount As Long)
    Dim s As W_SLIDER, pos As Long
    If ScrollBarWidget_isEnabled(self) Then
        Set s = W_SLIDER_(ScrollBarWidget_(self).slider)
        pos = amount + s.pos
        CLIP pos, 0, s.numStops
        s.pos = pos
    End If
End Sub

Public Sub ScrollBarWidget_setScrollableSize(ByVal self As WIDGET, ByVal value As Long)
    Dim d As ScrollBarWidget, s As W_SLIDER, moveToBottom As Boolean
    Set d = ScrollBarWidget_(self)
    Set s = W_SLIDER_(d.slider)
    moveToBottom = d.stickToBottom And s.pos = s.numStops
    d.scrollableSize = value And &HFFFF&
    ScrollBarWidget_updateSlider self
    If moveToBottom Then
        s.pos = s.numStops
    End If
End Sub

Public Sub ScrollBarWidget_setViewSize(ByVal self As WIDGET, ByVal value As Long)
    ScrollBarWidget_(self).viewSize = value And &HFFFF&
    ScrollBarWidget_updateSlider self
End Sub

Public Sub ScrollBarWidget_setStickToBottom(ByVal self As WIDGET, ByVal value As Boolean)
    ScrollBarWidget_(self).stickToBottom = value
End Sub

Private Sub ScrollBarWidget_updateSlider(ByVal self As WIDGET)
    Dim d As ScrollBarWidget, s As W_SLIDER
    Set d = ScrollBarWidget_(self)
    Set s = W_SLIDER_(d.slider)
    If d.scrollableSize > 0 Then
        s.barSize = (d.viewSize * self.height() \ d.scrollableSize) And &HFFFF&
    Else
        s.barSize = 0
    End If
    s.numStops = MAX_(0, d.scrollableSize - d.viewSize)
    s.pos = MIN_(s.pos, s.numStops)
End Sub

Public Sub ScrollBarWidget_enable(ByVal self As WIDGET)
    W_SLIDER_enable ScrollBarWidget_(self).slider
End Sub

Public Sub ScrollBarWidget_disable(ByVal self As WIDGET)
    W_SLIDER_disable ScrollBarWidget_(self).slider
End Sub

Public Function ScrollBarWidget_isEnabled(ByVal self As WIDGET) As Boolean
    ScrollBarWidget_isEnabled = W_SLIDER_isEnabled(ScrollBarWidget_(self).slider)
End Function
