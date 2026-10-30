Attribute VB_Name = "margin_c"
' lib/widget/margin.cpp, margin.h: struct Margin and class MarginWidget (its
' fields: MarginWidget.cls). Margin(top, right, bottom, left).wrap(widget) is
' Margin_wrap(top, right, bottom, left, widget); Margin(vertical, horizontal)
' and Margin(value) pass their values out in full.
' (VB) widget class MarginWidget : WIDGET
Option Explicit

' explicit MarginWidget(Margin margin)
Public Sub MarginWidget_construct(ByVal self As WIDGET, ByVal top As Long, ByVal right_ As Long, ByVal bottom As Long, ByVal left_ As Long)
    Dim d As MarginWidget
    WIDGET_construct self
    self.kind = K_MarginWidget ' (VB) virtual calls from here on reach MarginWidget's overrides
    Set d = MarginWidget_(self)
    d.margin_top = top
    d.margin_right = right_
    d.margin_bottom = bottom
    d.margin_left = left_
    self.setTransparentToMouse True
End Sub

' std::shared_ptr<MarginWidget> Margin::wrap(std::shared_ptr<WIDGET> widget)
Public Function Margin_wrap(ByVal top As Long, ByVal right_ As Long, ByVal bottom As Long, ByVal left_ As Long, ByVal widget As WIDGET) As WIDGET
    Dim wrap As WIDGET
    Set wrap = WIDGET_alloc(K_MarginWidget)
    MarginWidget_construct wrap, top, right_, bottom, left_
    wrap.attach widget
    Set Margin_wrap = wrap
End Function

Public Sub MarginWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As MarginWidget, child As WIDGET, w As Long, h As Long
    Set d = MarginWidget_(self)
    If self.childCount() <> 0 Then
        Set child = self.childAt(0)
        w = self.width() - d.margin_left - d.margin_right
        If w < 0 Then w = 0
        h = self.height() - d.margin_top - d.margin_bottom
        If h < 0 Then h = 0
        child.setGeometry d.margin_left, d.margin_top, w, h
    End If
End Sub

Public Function MarginWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As MarginWidget, innerWidth As Long
    Set d = MarginWidget_(self)
    If self.childCount() = 0 Then innerWidth = 0 Else innerWidth = self.childAt(0).idealWidth()
    MarginWidget_idealWidth = innerWidth + d.margin_left + d.margin_right
End Function

Public Function MarginWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As MarginWidget, innerHeight As Long
    Set d = MarginWidget_(self)
    If self.childCount() = 0 Then innerHeight = 0 Else innerHeight = self.childAt(0).idealHeight()
    MarginWidget_idealHeight = innerHeight + d.margin_top + d.margin_bottom
End Function
