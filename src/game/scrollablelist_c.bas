Attribute VB_Name = "scrollablelist_c"
' lib/widget/scrollablelist.cpp, scrollablelist.h: the widget classes
' ScrollableListWidget and ClickableScrollableList (docs/CONVENTIONS.md,
' "Widgets"; their fields are ScrollableListWidget.cls and
' ClickableScrollableList.cls). make() is <Class>_make. A PIELIGHT is a Long;
' an optional<size_t> return is a Boolean (has_value) and a ByRef value.
' getItems() returns a new Collection of the items. lineDraws, a member that
' display clears before each use, is a local of displayRecursive. The
' ClickableScrollableList_OnClick_Func is a WzFunc called with the list, the
' OnHighlight_Func with the list and isHighlighted.
' (VB) widget class ScrollableListWidget : WIDGET
' (VB) virtual ScrollableListWidget.initialize
' (VB) widget class ClickableScrollableList : ScrollableListWidget
Option Explicit

' pietypes.h's PIERECT_DrawRequest (WzGfx's pie_DrawMultiRect reads an array of them)
Private Type PIERECT_DrawRequest
    x0 As Long
    y0 As Long
    x1 As Long
    y1 As Long
    color As Long                       ' PIELIGHT
End Type

Private Const SCROLLBAR_WIDTH As Long = 15

' (VB) PIELIGHT::isTransparent(): its alpha is 0
Private Function PIELIGHT_isTransparent(ByVal c As Long) As Boolean
    PIELIGHT_isTransparent = (c And &HFF000000) = 0
End Function

' (VB) the size of a std::vector<uint32_t> that getScrollSnapOffsets filled (an erased array is empty)
Private Function snapOffsetsCount(a() As Long) As Long
    On Error Resume Next
    snapOffsetsCount = UBound(a) - LBound(a) + 1
    If Err.Number <> 0 Then snapOffsetsCount = 0
    Err.Clear
End Function

' child->getScrollSnapOffsets().value_or(std::vector<uint32_t>{0})
Private Function snapOffsetsOrZero(ByVal child As WIDGET, offsets() As Long) As Long
    If child.getScrollSnapOffsets(offsets) Then
        snapOffsetsOrZero = snapOffsetsCount(offsets)
    Else
        ReDim offsets(0 To 0)
        offsets(0) = 0
        snapOffsetsOrZero = 1
    End If
End Function

' static std::shared_ptr<ScrollableListWidget> make()
Public Function ScrollableListWidget_make() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ScrollableListWidget)
    ScrollableListWidget_construct w
    vt_ScrollableListWidget_initialize w
    Set ScrollableListWidget_make = w
End Function

' ScrollableListWidget(): WIDGET() {}
Public Sub ScrollableListWidget_construct(ByVal self As WIDGET)
    Dim d As ScrollableListWidget, p As Padding
    WIDGET_construct self
    self.kind = K_ScrollableListWidget ' (VB) virtual calls from here on reach ScrollableListWidget's overrides
    Set d = ScrollableListWidget_(self)
    d.scrollableHeight = 0
    d.snapOffset = True
    d.layoutDirty = False
    d.padding_ = p                       ' {0, 0, 0, 0}
    d.itemSpacing = 0
    d.scrollbarWidth = 0
    d.expandWidthWhenScrollbarInvisible = True
    d.drawRowLines = False
    d.topVisibleItemIdx = 0
End Sub

Public Sub ScrollableListWidget_initialize(ByVal self As WIDGET)
    Dim d As ScrollableListWidget
    Set d = ScrollableListWidget_(self)
    Set d.scrollBar = ScrollBarWidget_make()
    self.attach d.scrollBar
    Set d.listView = ClipRectWidget_new()
    self.attach d.listView
    d.scrollBar.show False
    d.scrollbarWidth = SCROLLBAR_WIDTH
    d.backgroundColor = 0               ' clear()
    d.borderColor = 0
    d.rowLinesColor = pal_RGBA(0, 0, 0, 255)
End Sub

Public Sub ScrollableListWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As ScrollableListWidget
    Set d = ScrollableListWidget_(self)
    d.scrollBar.setGeometry self.width() - d.scrollbarWidth, 0, d.scrollbarWidth, self.height()
    ScrollBarWidget_setViewSize d.scrollBar, self.height()
    d.layoutDirty = True
End Sub

' (VB) listView->setTopOffset(snapOffset ? snappedOffset() : scrollBar->position())
Private Sub ScrollableListWidget_setTopOffsetFromScroll(ByVal self As WIDGET)
    Dim d As ScrollableListWidget
    Set d = ScrollableListWidget_(self)
    If d.snapOffset Then
        ClipRectWidget_setTopOffset d.listView, ScrollableListWidget_snappedOffset(self)
    Else
        ClipRectWidget_setTopOffset d.listView, ScrollBarWidget_position(d.scrollBar)
    End If
End Sub

Public Sub ScrollableListWidget_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    ScrollableListWidget_updateLayout self
    ScrollableListWidget_setTopOffsetFromScroll self
End Sub

'
' Snap offset to first visible child.
'
' This wouldn't be necessary if it were possible to clip the rendering.
'
Private Function ScrollableListWidget_snappedOffset(ByVal self As WIDGET) As Long
    Dim d As ScrollableListWidget, items As WIDGET, idx As Long, child As WIDGET, childOffsets() As Long
    Dim n As Long, i As Long, y As Long
    Set d = ScrollableListWidget_(self)
    Set items = d.listView
    For idx = 0 To items.childCount() - 1
        Set child = items.childAt(idx)
        If WzRect_bottom(child.geometry()) < ScrollBarWidget_position(d.scrollBar) Then
            GoTo continue_
        End If

        n = snapOffsetsOrZero(child, childOffsets)
        For i = 0 To n - 1
            y = child.y() + childOffsets(i)
            If y >= ScrollBarWidget_position(d.scrollBar) Then
                d.topVisibleItemIdx = idx
                ScrollableListWidget_snappedOffset = y
                Exit Function
            End If
        Next
continue_:
    Next

    d.topVisibleItemIdx = 0
    ScrollableListWidget_snappedOffset = 0
End Function

Private Function ScrollableListWidget_getScrollPositionForItem(ByVal self As WIDGET, ByVal itemNum As Long) As Long
    Dim d As ScrollableListWidget, child As WIDGET, childOffsets() As Long, n As Long
    Set d = ScrollableListWidget_(self)
    If Not (itemNum < d.listView.childCount()) Then ASSERT_FAILURE "itemNum < listView->children().size()", "lib/widget/scrollablelist.cpp:88", "getScrollPositionForItem", "Invalid itemNum: %zu", itemNum: Exit Function
    Set child = d.listView.childAt(itemNum)
    n = snapOffsetsOrZero(child, childOffsets)
    If n > 0 Then
        ScrollableListWidget_getScrollPositionForItem = child.y() + childOffsets(0)
        Exit Function
    End If
    ScrollableListWidget_getScrollPositionForItem = child.y()
End Function

Public Function ScrollableListWidget_getCurrentYPosOfItem(ByVal self As WIDGET, ByVal itemNum As Long) As Long
    Dim d As ScrollableListWidget, currTopOffset As Long, itemYPos As Long
    Set d = ScrollableListWidget_(self)
    currTopOffset = ClipRectWidget_getTopOffset(d.listView)
    itemYPos = ScrollableListWidget_getScrollPositionForItem(self, itemNum)
    ScrollableListWidget_getCurrentYPosOfItem = (itemYPos - currTopOffset) + d.listView.y()
End Function

Public Function ScrollableListWidget_addItem(ByVal self As WIDGET, ByVal item As WIDGET) As Long
    Dim d As ScrollableListWidget, newItemIdx As Long
    Set d = ScrollableListWidget_(self)
    newItemIdx = d.listView.childCount()
    d.listView.attach item
    d.layoutDirty = True
    ScrollableListWidget_addItem = newItemIdx
End Function

Public Sub ScrollableListWidget_clear(ByVal self As WIDGET)
    Dim d As ScrollableListWidget
    Set d = ScrollableListWidget_(self)
    d.listView.removeAllChildren
    d.layoutDirty = True
    ScrollableListWidget_updateLayout self
    ClipRectWidget_setTopOffset d.listView, 0
    d.topVisibleItemIdx = 0
End Sub

Public Function ScrollableListWidget_numItems(ByVal self As WIDGET) As Long
    ScrollableListWidget_numItems = ScrollableListWidget_(self).listView.childCount()
End Function

Public Function ScrollableListWidget_getItemAtIdx(ByVal self As WIDGET, ByVal itemNum As Long) As WIDGET
    Dim d As ScrollableListWidget
    Set d = ScrollableListWidget_(self)
    If Not (itemNum < d.listView.childCount()) Then ASSERT_FAILURE "itemNum < listView->children().size()", "lib/widget/scrollablelist.cpp:130", "getItemAtIdx", "itemNum out of bounds - number of items: %zu", d.listView.childCount(): Exit Function
    Set ScrollableListWidget_getItemAtIdx = d.listView.childAt(itemNum)
End Function

Public Function ScrollableListWidget_getItems(ByVal self As WIDGET) As Collection
    Dim d As ScrollableListWidget, c As Collection, i As Long
    Set d = ScrollableListWidget_(self)
    Set c = New Collection
    For i = 0 To d.listView.childCount() - 1
        c.Add d.listView.childAt(i)
    Next
    Set ScrollableListWidget_getItems = c
End Function

Private Sub ScrollableListWidget_updateLayout(ByVal self As WIDGET)
    Dim d As ScrollableListWidget, p As Padding, listViewWidthWithoutScrollBar As Long, widthOfScrollbar As Long
    Dim listViewWidthWithScrollBar As Long, listViewHeight As Long
    Set d = ScrollableListWidget_(self)
    If Not d.layoutDirty Then
        Exit Sub
    End If
    d.layoutDirty = False

    p = d.padding_
    listViewWidthWithoutScrollBar = ScrollableListWidget_calculateListViewWidth(self)
    widthOfScrollbar = d.scrollBar.width()
    listViewWidthWithScrollBar = listViewWidthWithoutScrollBar
    If ULT(p.right_, widthOfScrollbar) Then
        If widthOfScrollbar >= 0 And listViewWidthWithoutScrollBar > widthOfScrollbar Then
            listViewWidthWithScrollBar = listViewWidthWithoutScrollBar - widthOfScrollbar
        Else
            listViewWidthWithScrollBar = 0
        End If
    End If
    listViewHeight = ScrollableListWidget_calculateListViewHeight(self)

    ScrollableListWidget_resizeChildren self, listViewWidthWithScrollBar

    d.scrollBar.show d.scrollableHeight > listViewHeight

    If d.scrollBar.visible() Or Not d.expandWidthWhenScrollbarInvisible Then
        d.listView.setGeometry p.left_, p.top, listViewWidthWithScrollBar, listViewHeight
    Else
        If listViewWidthWithScrollBar <> listViewWidthWithoutScrollBar Then
            ScrollableListWidget_resizeChildren self, listViewWidthWithoutScrollBar
        End If
        d.listView.setGeometry p.left_, p.top, listViewWidthWithoutScrollBar, listViewHeight
    End If

    ScrollBarWidget_setScrollableSize d.scrollBar, d.scrollableHeight + p.top + p.bottom
End Sub

Private Sub ScrollableListWidget_resizeChildren(ByVal self As WIDGET, ByVal width As Long)
    Dim d As ScrollableListWidget, nextOffset As Long, i As Long, child As WIDGET
    Set d = ScrollableListWidget_(self)
    d.scrollableHeight = 0
    nextOffset = 0
    For i = 0 To d.listView.childCount() - 1
        Set child = d.listView.childAt(i)
        If Not child.visible() Then
            GoTo continue_
        End If
        child.setGeometry 0, nextOffset, width, child.height()
        d.scrollableHeight = nextOffset + child.height()
        nextOffset = d.scrollableHeight + d.itemSpacing
continue_:
    Next
End Sub

Public Function ScrollableListWidget_calculateListViewHeight(ByVal self As WIDGET) As Long
    Dim p As Padding, result As Long
    p = ScrollableListWidget_(self).padding_
    result = self.height() - p.top - p.bottom
    If result > 0 Then
        ScrollableListWidget_calculateListViewHeight = result
    Else
        ScrollableListWidget_calculateListViewHeight = 0
    End If
End Function

Public Function ScrollableListWidget_calculateListViewWidth(ByVal self As WIDGET) As Long
    Dim p As Padding, result As Long
    p = ScrollableListWidget_(self).padding_
    result = self.width() - p.left_ - p.right_
    If result > 0 Then
        ScrollableListWidget_calculateListViewWidth = result
    Else
        ScrollableListWidget_calculateListViewWidth = 0
    End If
End Function

Public Function ScrollableListWidget_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim wheelX As Long, wheelY As Long
    getMouseWheelSpeed wheelX, wheelY
    ScrollBarWidget_incrementPosition ScrollableListWidget_(self).scrollBar, -wheelY * 20
    Set ScrollableListWidget_findMouseTargetRecursive = WIDGET_findMouseTargetRecursive(self, psContext, key, wasPressed)
End Function

Public Sub ScrollableListWidget_enableScroll(ByVal self As WIDGET)
    ScrollBarWidget_enable ScrollableListWidget_(self).scrollBar
End Sub

Public Sub ScrollableListWidget_disableScroll(ByVal self As WIDGET)
    ScrollBarWidget_disable ScrollableListWidget_(self).scrollBar
End Sub

Public Sub ScrollableListWidget_setStickToBottom(ByVal self As WIDGET, ByVal value As Boolean)
    ScrollBarWidget_setStickToBottom ScrollableListWidget_(self).scrollBar, value
End Sub

Public Sub ScrollableListWidget_setPadding(ByVal self As WIDGET, rect As Padding)
    ScrollableListWidget_(self).padding_ = rect
    ScrollableListWidget_(self).layoutDirty = True
End Sub

Public Function ScrollableListWidget_getPadding(ByVal self As WIDGET) As Padding
    ScrollableListWidget_getPadding = ScrollableListWidget_(self).padding_
End Function

Public Sub ScrollableListWidget_setDrawRowLines(ByVal self As WIDGET, ByVal bEnabled As Boolean)
    ScrollableListWidget_(self).drawRowLines = bEnabled
End Sub

Public Sub ScrollableListWidget_setBackgroundColor(ByVal self As WIDGET, ByVal color As Long)
    ScrollableListWidget_(self).backgroundColor = color
End Sub

Public Sub ScrollableListWidget_setBorderColor(ByVal self As WIDGET, ByVal color As Long)
    ScrollableListWidget_(self).borderColor = color
End Sub

Public Sub ScrollableListWidget_setSnapOffset(ByVal self As WIDGET, ByVal value As Boolean)
    ScrollableListWidget_(self).snapOffset = value
End Sub

Public Sub ScrollableListWidget_setItemSpacing(ByVal self As WIDGET, ByVal value As Long)
    If value = ScrollableListWidget_(self).itemSpacing Then
        Exit Sub
    End If
    ScrollableListWidget_(self).itemSpacing = value
    ScrollableListWidget_(self).layoutDirty = True
End Sub

Public Function ScrollableListWidget_getItemSpacing(ByVal self As WIDGET) As Long
    ScrollableListWidget_getItemSpacing = ScrollableListWidget_(self).itemSpacing
End Function

Public Sub ScrollableListWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As ScrollableListWidget, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Set d = ScrollableListWidget_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    If Not PIELIGHT_isTransparent(d.backgroundColor) Then
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), d.backgroundColor
    End If

    If Not PIELIGHT_isTransparent(d.borderColor) Then
        x1 = x0 + self.width()
        y1 = y0 + self.height()
        iV_Box x0, y0, x1, y1, d.borderColor
    End If
End Sub

Public Sub ScrollableListWidget_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim d As ScrollableListWidget, lineDraws() As PIERECT_DrawRequest, lineDrawsCount As Long
    Dim xOffset As Long, yOffset As Long, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Dim listViewTopOffset As Long, i As Long, child As WIDGET, childY As Long, childDisplayY0 As Long, rowLineY0 As Long
    Set d = ScrollableListWidget_(self)
    ScrollableListWidget_updateLayout self
    WIDGET_displayRecursive self, context

    ' *after* displaying all children, draw the row lines (if desired)
    If Not d.drawRowLines Or d.itemSpacing = 0 Then
        Exit Sub
    End If

    If Not WidgetGraphicsContext_clipContains(context, self.geometry()) Then
        Exit Sub
    End If

    lineDrawsCount = 0                  ' lineDraws.clear()

    xOffset = WidgetGraphicsContext_getXOffset(context)
    yOffset = WidgetGraphicsContext_getYOffset(context)

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset + d.listView.y()
    x1 = x0 + d.listView.width()
    y1 = y0 + d.listView.height()

    listViewTopOffset = ClipRectWidget_getTopOffset(d.listView)
    For i = 0 To d.listView.childCount() - 1
        Set child = d.listView.childAt(i)
        childY = child.y()
        If childY < listViewTopOffset Then
            GoTo continue_
        End If
        childDisplayY0 = childY - listViewTopOffset
        rowLineY0 = y0 + childDisplayY0 - d.itemSpacing
        If rowLineY0 > y1 Or rowLineY0 < y0 Then
            GoTo continue_
        End If
        ReDim Preserve lineDraws(0 To lineDrawsCount)
        lineDraws(lineDrawsCount).x0 = x0
        lineDraws(lineDrawsCount).y0 = rowLineY0
        lineDraws(lineDrawsCount).x1 = x1
        lineDraws(lineDrawsCount).y1 = rowLineY0 + d.itemSpacing
        lineDraws(lineDrawsCount).color = d.rowLinesColor
        lineDrawsCount = lineDrawsCount + 1
continue_:
    Next

    If lineDrawsCount > 0 Then
        pie_DrawMultiRect VarPtr(lineDraws(0)), lineDrawsCount
    End If
End Sub

Public Function ScrollableListWidget_getScrollbarWidth(ByVal self As WIDGET) As Long
    ScrollableListWidget_getScrollbarWidth = ScrollableListWidget_(self).scrollbarWidth
End Function

Public Sub ScrollableListWidget_setScrollbarWidth(ByVal self As WIDGET, ByVal newWidth As Long)
    ScrollableListWidget_(self).scrollbarWidth = newWidth
    self.geometryChanged
    ScrollableListWidget_(self).layoutDirty = True
    ScrollableListWidget_updateLayout self
End Sub

Public Sub ScrollableListWidget_setExpandWhenScrollbarInvisible(ByVal self As WIDGET, ByVal expandWidth As Boolean)
    ScrollableListWidget_(self).expandWidthWhenScrollbarInvisible = expandWidth
End Sub

Public Function ScrollableListWidget_getScrollPosition(ByVal self As WIDGET) As Long
    ScrollableListWidget_getScrollPosition = ScrollBarWidget_position(ScrollableListWidget_(self).scrollBar)
End Function

Public Sub ScrollableListWidget_setScrollPosition(ByVal self As WIDGET, ByVal newPosition As Long)
    ScrollableListWidget_updateLayout self
    ScrollBarWidget_setPosition ScrollableListWidget_(self).scrollBar, newPosition
    ScrollableListWidget_setTopOffsetFromScroll self
End Sub

Public Function ScrollableListWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As ScrollableListWidget, maxItemIdealWidth As Long, i As Long, p As Padding
    Set d = ScrollableListWidget_(self)
    maxItemIdealWidth = 0
    For i = 0 To d.listView.childCount() - 1
        maxItemIdealWidth = MAX_(maxItemIdealWidth, d.listView.childAt(i).idealWidth())
    Next

    p = d.padding_
    ScrollableListWidget_idealWidth = maxItemIdealWidth + p.left_ + p.right_ + d.scrollbarWidth
End Function

Public Function ScrollableListWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim p As Padding
    ScrollableListWidget_updateLayout self
    p = ScrollableListWidget_(self).padding_
    ScrollableListWidget_idealHeight = ScrollableListWidget_(self).scrollableHeight + p.top + p.bottom
End Function

Public Sub ScrollableListWidget_scrollToItem(ByVal self As WIDGET, ByVal itemNum As Long)
    ScrollableListWidget_updateLayout self
    ScrollBarWidget_setPosition ScrollableListWidget_(self).scrollBar, ScrollableListWidget_getScrollPositionForItem(self, itemNum)
    ScrollableListWidget_setTopOffsetFromScroll self
End Sub

Public Function ScrollableListWidget_isItemVisible(ByVal self As WIDGET, ByVal itemNum As Long) As Boolean
    Dim d As ScrollableListWidget
    Set d = ScrollableListWidget_(self)
    If Not (itemNum < d.listView.childCount()) Then ASSERT_FAILURE "itemNum < listView->children().size()", "lib/widget/scrollablelist.cpp:382", "isItemVisible", "Invalid itemNum: %zu", itemNum: Exit Function
    ScrollableListWidget_updateLayout self
    ScrollableListWidget_isItemVisible = ClipRectWidget_isChildVisible(d.listView, d.listView.childAt(itemNum))
End Function

Public Sub ScrollableListWidget_scrollEnsureItemVisible(ByVal self As WIDGET, ByVal itemNum As Long)
    Dim d As ScrollableListWidget, scrollPosOfItem As Long, desiredScrollPos As Long
    Set d = ScrollableListWidget_(self)
    If Not self.visible() Or self.width() <= 0 Or self.height() <= 0 Then
        ' no-op if this isn't visible
        Exit Sub
    End If
    If Not ScrollableListWidget_isItemVisible(self, itemNum) Then
        scrollPosOfItem = ScrollableListWidget_getScrollPositionForItem(self, itemNum)
        If scrollPosOfItem < ScrollBarWidget_position(d.scrollBar) Then
            ' scroll up the minimum amount
            ScrollBarWidget_setPosition d.scrollBar, scrollPosOfItem
        Else
            ' scroll down the minimum amount
            desiredScrollPos = MAX_(scrollPosOfItem - d.listView.height() + d.listView.childAt(itemNum).height(), 0)
            ScrollBarWidget_setPosition d.scrollBar, desiredScrollPos
        End If
        ScrollableListWidget_setTopOffsetFromScroll self
    End If
End Sub

Public Sub ScrollableListWidget_setListTransparentToMouse(ByVal self As WIDGET, ByVal hasMouseTransparency As Boolean)
    ScrollableListWidget_(self).listView.setTransparentToMouse hasMouseTransparency
End Sub

Public Sub ScrollableListWidget_setLayoutDirty(ByVal self As WIDGET)
    ScrollableListWidget_(self).layoutDirty = True
End Sub

' yPos is widget coordinates
Public Function ScrollableListWidget_getItemAtYPos(ByVal self As WIDGET, ByVal yPos As Long) As WIDGET
    Dim itemIdx As Long
    If Not ScrollableListWidget_getItemIdxAtYPos(self, yPos, itemIdx) Then
        Exit Function
    End If
    Set ScrollableListWidget_getItemAtYPos = ScrollableListWidget_(self).listView.childAt(itemIdx)
End Function

' yPos is widget coordinates
Public Function ScrollableListWidget_getItemIdxAtYPos(ByVal self As WIDGET, ByVal yPos As Long, value As Long) As Boolean
    Dim d As ScrollableListWidget, listViewY0 As Long, listViewHeight As Long, listViewY1 As Long
    Dim idx As Long, itemCurrentY0 As Long, itemCurrentY1 As Long, itemWidget As WIDGET
    Set d = ScrollableListWidget_(self)
    listViewY0 = d.listView.y()
    listViewHeight = d.listView.height()
    listViewY1 = listViewY0 + listViewHeight
    If yPos < listViewY0 Or yPos > listViewY1 Then
        Exit Function                   ' nullopt
    End If
    For idx = d.topVisibleItemIdx To d.listView.childCount() - 1
        itemCurrentY0 = ScrollableListWidget_getCurrentYPosOfItem(self, idx)
        If itemCurrentY0 >= listViewHeight Then
            Exit For
        End If
        Set itemWidget = d.listView.childAt(idx)
        itemCurrentY1 = itemCurrentY0 + itemWidget.height()
        If yPos >= itemCurrentY0 And yPos < itemCurrentY1 Then
            value = idx
            ScrollableListWidget_getItemIdxAtYPos = True
            Exit Function
        End If
    Next
End Function

'--------------------------------------------------------------------------
' MARK: - ClickableScrollableList

Public Function ClickableScrollableList_make() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ClickableScrollableList)
    ClickableScrollableList_construct w
    vt_ScrollableListWidget_initialize w    ' calls ScrollableListWidget::initialize
    ScrollableListWidget_setListTransparentToMouse w, True
    Set ClickableScrollableList_make = w
End Function

Public Sub ClickableScrollableList_construct(ByVal self As WIDGET)
    ScrollableListWidget_construct self
    self.kind = K_ClickableScrollableList ' (VB) virtual calls from here on reach ClickableScrollableList's overrides
    ClickableScrollableList_(self).mouseDownOnList = False
End Sub

Public Sub ClickableScrollableList_setOnClickHandler(ByVal self As WIDGET, ByVal onClickFunc As WzFunc)
    Set ClickableScrollableList_(self).onClickFunc = onClickFunc
End Sub

Public Sub ClickableScrollableList_setOnHighlightHandler(ByVal self As WIDGET, ByVal onHighlightFunc As WzFunc)
    Set ClickableScrollableList_(self).onHighlightFunc = onHighlightFunc
End Sub

Public Sub ClickableScrollableList_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    ClickableScrollableList_(self).mouseDownOnList = True
End Sub

Public Sub ClickableScrollableList_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As ClickableScrollableList, wasFullClick As Boolean
    Set d = ClickableScrollableList_(self)
    wasFullClick = d.mouseDownOnList
    d.mouseDownOnList = False
    If wasFullClick Then
        If Not d.onClickFunc Is Nothing Then
            WzFunc_call d.onClickFunc, self
        End If
    End If
End Sub

Public Sub ClickableScrollableList_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As ClickableScrollableList
    Set d = ClickableScrollableList_(self)
    If Not d.onHighlightFunc Is Nothing Then
        WzFunc_call d.onHighlightFunc, self, True
    End If
End Sub

Public Sub ClickableScrollableList_highlightLost(ByVal self As WIDGET)
    Dim d As ClickableScrollableList
    Set d = ClickableScrollableList_(self)
    d.mouseDownOnList = False
    If Not d.onHighlightFunc Is Nothing Then
        WzFunc_call d.onHighlightFunc, self, False
    End If
End Sub
