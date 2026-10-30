Attribute VB_Name = "cliprect_c"
' lib/widget/cliprect.cpp, cliprect.h: the widget class ClipRectWidget
' (docs/CONVENTIONS.md, "Widgets"; its field, glm::ivec2 offset, is
' ClipRectWidget.cls's offsetX and offsetY).
'
' This widget will render only the childrens that are visible, and whose rectangle is contained in the clip rect's rectangle.
' (VB) widget class ClipRectWidget : WIDGET
Option Explicit

' std::make_shared<ClipRectWidget>()
Public Function ClipRectWidget_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ClipRectWidget)
    ClipRectWidget_construct w
    Set ClipRectWidget_new = w
End Function

' ClipRectWidget() : WIDGET() {}
Public Sub ClipRectWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_ClipRectWidget ' (VB) virtual calls from here on reach ClipRectWidget's overrides
    ClipRectWidget_(self).offsetX = 0
    ClipRectWidget_(self).offsetY = 0
End Sub

Public Sub ClipRectWidget_runRecursive(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim newContext As W_CONTEXT
    newContext = psContext
    newContext.xOffset = psContext.xOffset + self.x()
    newContext.yOffset = psContext.yOffset + self.y()
    newContext.mx = psContext.mx - self.x()
    newContext.my = psContext.my - self.y() + ClipRectWidget_(self).offsetY

    WIDGET_runRecursive self, newContext
End Sub

Public Function ClipRectWidget_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim newContext As W_CONTEXT, result As WIDGET, d As ClipRectWidget
    Set d = ClipRectWidget_(self)
    newContext = psContext
    newContext.xOffset = psContext.xOffset - d.offsetX
    newContext.yOffset = psContext.yOffset - d.offsetY
    newContext.mx = psContext.mx + d.offsetX
    newContext.my = psContext.my + d.offsetY

    Set result = WIDGET_findMouseTargetRecursive(self, newContext, key, wasPressed)
    If Not result Is Nothing Then
        psContext = newContext  ' bubble-up the matching hit's context
    End If
    Set ClipRectWidget_findMouseTargetRecursive = result
End Function

Public Sub ClipRectWidget_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim childrenContext As WidgetGraphicsContext, translated As WidgetGraphicsContext, d As ClipRectWidget
    Dim i As Long, child As WIDGET
    Set d = ClipRectWidget_(self)
    If WidgetGraphicsContext_clipContains(context, self.geometry()) Then
        self.display WidgetGraphicsContext_getXOffset(context), WidgetGraphicsContext_getYOffset(context)
    End If

    translated = WidgetGraphicsContext_translatedBy(context, self.x() - d.offsetX, +self.y() - d.offsetY)
    childrenContext = WidgetGraphicsContext_clippedBy(translated, WzRect_(d.offsetX, d.offsetY, self.width(), self.height()))

    For i = 0 To self.childCount() - 1
        Set child = self.childAt(i)
        If child.visible() Then
            child.displayRecursive childrenContext
        End If
    Next
End Sub

Public Function ClipRectWidget_isChildVisible(ByVal self As WIDGET, ByVal child As WIDGET) As Boolean
    Dim childrenContext As WidgetGraphicsContext, base As WidgetGraphicsContext, translated As WidgetGraphicsContext, d As ClipRectWidget
    If Not (child.parent() Is self) Then ASSERT_FAILURE "child->parent() == shared_from_this()", "lib/widget/cliprect.cpp:78", "isChildVisible", "Not a child of this widget?": Exit Function
    Set d = ClipRectWidget_(self)
    WidgetGraphicsContext_construct base
    translated = WidgetGraphicsContext_translatedBy(base, self.screenPosX(), self.screenPosY())
    childrenContext = WidgetGraphicsContext_clippedBy(translated, WzRect_(d.offsetX, d.offsetY, self.width(), self.height()))
    ClipRectWidget_isChildVisible = WidgetGraphicsContext_clipContains(childrenContext, child.geometry())
End Function

Public Function ClipRectWidget_setTopOffset(ByVal self As WIDGET, ByVal value As Long) As Boolean
    value = value And &HFFFF&           ' uint16_t
    If value = ClipRectWidget_(self).offsetY Then
        ClipRectWidget_setTopOffset = False
        Exit Function
    End If
    ClipRectWidget_(self).offsetY = value
    ClipRectWidget_setTopOffset = True
End Function

Public Function ClipRectWidget_getTopOffset(ByVal self As WIDGET) As Long
    ClipRectWidget_getTopOffset = ClipRectWidget_(self).offsetY And &HFFFF&
End Function

Public Function ClipRectWidget_setLeftOffset(ByVal self As WIDGET, ByVal value As Long) As Boolean
    value = value And &HFFFF&           ' uint16_t
    If value = ClipRectWidget_(self).offsetX Then
        ClipRectWidget_setLeftOffset = False
        Exit Function
    End If
    ClipRectWidget_(self).offsetX = value
    ClipRectWidget_setLeftOffset = True
End Function

Public Function ClipRectWidget_parentRelativeXOffset(ByVal self As WIDGET, ByVal coord As Long) As Long
    ClipRectWidget_parentRelativeXOffset = self.x() - ClipRectWidget_(self).offsetX + coord
End Function

Public Function ClipRectWidget_parentRelativeYOffset(ByVal self As WIDGET, ByVal coord As Long) As Long
    ClipRectWidget_parentRelativeYOffset = self.y() - ClipRectWidget_(self).offsetY + coord
End Function
