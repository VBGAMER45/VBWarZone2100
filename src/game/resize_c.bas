Attribute VB_Name = "resize_c"
' lib/widget/resize.cpp, resize.h: struct Resize (a UDT: Resize::width(min, max) is Resize_width, height
' Resize_height, fixed Resize_fixed, wrap Resize_wrap) and the widget class ResizeWidget (its fields:
' ResizeWidget.cls; docs/CONVENTIONS.md, "Widgets"). A nonstd::optional<int32_t> is has_<name> + <name>; the
' factories take a Variant each, Missing or Empty for nullopt.
' (VB) widget class ResizeWidget : WIDGET
Option Explicit

Public Type Resize
    has_minWidth As Boolean
    minWidth As Long
    has_maxWidth As Boolean
    maxWidth As Long
    has_minHeight As Boolean
    minHeight As Long
    has_maxHeight As Boolean
    maxHeight As Long
End Type

' static Resize width(nonstd::optional<int32_t> min, nonstd::optional<int32_t> max = nonstd::nullopt)
Public Function Resize_width(ByVal minValue As Variant, Optional ByVal maxValue As Variant) As Resize
    If Not IsMissing(minValue) Then If Not IsEmpty(minValue) Then Resize_width.has_minWidth = True: Resize_width.minWidth = minValue
    If Not IsMissing(maxValue) Then If Not IsEmpty(maxValue) Then Resize_width.has_maxWidth = True: Resize_width.maxWidth = maxValue
End Function

' static Resize height(nonstd::optional<int32_t> min, nonstd::optional<int32_t> max = nonstd::nullopt)
Public Function Resize_height(ByVal minValue As Variant, Optional ByVal maxValue As Variant) As Resize
    If Not IsMissing(minValue) Then If Not IsEmpty(minValue) Then Resize_height.has_minHeight = True: Resize_height.minHeight = minValue
    If Not IsMissing(maxValue) Then If Not IsEmpty(maxValue) Then Resize_height.has_maxHeight = True: Resize_height.maxHeight = maxValue
End Function

' static Resize fixed(int32_t width, int32_t height): Resize{width, width, height, height}
Public Function Resize_fixed(ByVal width As Long, ByVal height As Long) As Resize
    Resize_fixed.has_minWidth = True: Resize_fixed.minWidth = width
    Resize_fixed.has_maxWidth = True: Resize_fixed.maxWidth = width
    Resize_fixed.has_minHeight = True: Resize_fixed.minHeight = height
    Resize_fixed.has_maxHeight = True: Resize_fixed.maxHeight = height
End Function

' std::shared_ptr<ResizeWidget> Resize::wrap(std::shared_ptr<WIDGET> widget)
Public Function Resize_wrap(r As Resize, ByVal widget As WIDGET) As WIDGET
    Dim wrap As WIDGET
    Set wrap = WIDGET_alloc(K_ResizeWidget)
    ResizeWidget_construct wrap, r
    wrap.attach widget
    Set Resize_wrap = wrap
End Function

' explicit ResizeWidget(Resize resize): resize(resize)
Public Sub ResizeWidget_construct(ByVal self As WIDGET, r As Resize)
    Dim d As ResizeWidget
    WIDGET_construct self
    self.kind = K_ResizeWidget ' (VB) virtual calls from here on reach ResizeWidget's overrides
    Set d = ResizeWidget_(self)
    d.has_minWidth = r.has_minWidth: d.minWidth = r.minWidth
    d.has_maxWidth = r.has_maxWidth: d.maxWidth = r.maxWidth
    d.has_minHeight = r.has_minHeight: d.minHeight = r.minHeight
    d.has_maxHeight = r.has_maxHeight: d.maxHeight = r.maxHeight
    self.setTransparentToMouse True
End Sub

Public Sub ResizeWidget_geometryChanged(ByVal self As WIDGET)
    If self.childCount() <> 0 Then
        self.childAt(0).setGeometry 0, 0, self.width(), self.height()
    End If
End Sub

Public Function ResizeWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As ResizeWidget, ideal As Long
    Set d = ResizeWidget_(self)
    If d.has_minWidth Then ideal = d.minWidth Else ideal = 0     ' resize.minWidth.value_or(0)

    If self.childCount() <> 0 Then
        ideal = MAX_(ideal, self.childAt(0).idealWidth())
        If d.has_maxWidth Then
            ideal = MIN_(ideal, d.maxWidth)
        End If
    End If

    ResizeWidget_idealWidth = ideal
End Function

Public Function ResizeWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As ResizeWidget, ideal As Long
    Set d = ResizeWidget_(self)
    If d.has_minHeight Then ideal = d.minHeight Else ideal = 0   ' resize.minHeight.value_or(0)

    If self.childCount() <> 0 Then
        ideal = MAX_(ideal, self.childAt(0).idealHeight())
        If d.has_maxHeight Then
            ideal = MIN_(ideal, d.maxHeight)
        End If
    End If

    ResizeWidget_idealHeight = ideal
End Function
