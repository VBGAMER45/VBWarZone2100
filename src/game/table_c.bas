Attribute VB_Name = "table_c"
' lib/widget/table.cpp, table.h: the widget classes TableRow, TableHeader (of
' table.cpp) and ScrollableTableWidget (docs/CONVENTIONS.md, "Widgets"; their
' fields are TableRow.cls, TableHeader.cls and ScrollableTableWidget.cls), and
' struct TableColumn (a UDT; enum class TableColumn::ResizeBehavior is
' TableColumn_ResizeBehavior).
' (VB) A std::vector<std::shared_ptr<WIDGET>> argument (TableRow::make's columns)
' is a Collection; a std::vector<TableColumn> or std::vector<size_t> argument is
' an array and its count; a returned const std::vector<size_t>& is copied into the
' caller's array, its size returned (getColumnWidths, getMinimumColumnWidths);
' getRows returns the rows' Collection. optional<T> fields are has_<name> and
' <name>; an optional<size_t> return is a Boolean (has_value) and a ByRef value.
' A Vector2i padding is glm_h's Vector2i. TableHeader::parentTable (a weak_ptr)
' is the table's serial (widgweak_vb). size_t arithmetic is Long (the same while
' the table is wide enough for its padding, as the callers make it); std::sort of
' a few column indexes is a stable insertion sort (MSVC's for short ranges).
' (VB) widget class TableRow : W_BUTTON
' (VB) virtual TableRow.resizeColumns
' (VB) widget class TableHeader : W_FORM
' (VB) widget class ScrollableTableWidget : WIDGET
Option Explicit

Private Const TABLE_ROW_PADDING As Long = 5
Private Const TABLE_COL_PADDING As Long = 10
Private Const TABLE_COL_HEADER_HEIGHT As Long = (20 + (TABLE_ROW_PADDING * 2))

' enum class ResizeBehavior
Public Enum TableColumn_ResizeBehavior
    TableColumn_ResizeBehavior_RESIZABLE            ' user-resizable column
    TableColumn_ResizeBehavior_RESIZABLE_AUTOEXPAND ' user-resizable column, that also automatically expands to fill extra space
    TableColumn_ResizeBehavior_FIXED_WIDTH
End Enum

Public Type TableColumn
    columnWidget As WIDGET              ' std::shared_ptr<WIDGET>
    resizeBehavior As TableColumn_ResizeBehavior
End Type

' glm::ivec4 (iV_Lines reads an array of them)
Private Type ivec4
    x As Long
    y As Long
    z As Long
    w As Long
End Type

Public Const FN_TableHeader_calcLayout As Long = 2180
Public Const FN_ScrollableTableWidget_listCalcLayout As Long = 2181

Public Function table_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, psParent As WIDGET, d As ScrollableTableWidget, y0 As Long
    Select Case f.id
    Case FN_TableHeader_calcLayout
        ' LAMBDA_CALCLAYOUT_SIMPLE: (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not psParent Is Nothing Then
            ' set width to parent width
            psWidget.setGeometry 0, 0, psParent.width(), psWidget.height()
        End If
    Case FN_ScrollableTableWidget_listCalcLayout
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If WIDGET_isKindOf(psParent, K_ScrollableTableWidget) Then
            Set d = ScrollableTableWidget_(psParent)
            y0 = 0
            If d.header.visible() Then
                y0 = d.header.y() + d.header.height() + ScrollableListWidget_getItemSpacing(d.scrollableList)
            End If
            psWidget.setGeometry 0, y0, psParent.width(), psParent.height() - y0
        End If
    End Select
End Function

' TableColumn{columnWidget, resizeBehavior}
Public Function TableColumn_(ByVal columnWidget As WIDGET, ByVal resizeBehavior As TableColumn_ResizeBehavior) As TableColumn
    Set TableColumn_.columnWidget = columnWidget
    TableColumn_.resizeBehavior = resizeBehavior
End Function

'--------------------------------------------------------------------------
' MARK: - TableRow

' TableRow::TableRow()
Public Sub TableRow_construct(ByVal self As WIDGET)
    Dim d As TableRow
    W_BUTTON_construct self
    self.kind = K_TableRow ' (VB) virtual calls from here on reach TableRow's overrides
    Set d = TableRow_(self)
    d.highlightsOnMouseOver = False
    d.disabledRow = False
    d.has_lastFrameMouseIsOverRowOrChildren = False
    d.has_borderColor = False
    d.has_backgroundColor = False
    d.highlightColor = psPalette(WZCOL_TRANSPARENT_BOX)
    d.disabledColor = psPalette(WZCOL_FORM_DISABLE)
    d.maxDisplayedColumnX1 = 0
    Set W_BUTTON_(self).AudioCallback = Nothing
End Sub

' Make a new TableRow, given a vector of widgets (for the columns)
' If rowHeight <= 0, expand the TableRow height to match maximum height of any contained (column) widget
' If rowHeight > 0, use it as a fixed row height, and set all contained (column) widgets to that height
'
' NOTE: A column widget should expect its width to be adjustable, unless it is associated with a TableColumn
'       (in its parent ScrollableTableWidget) that has ResizeBehavior::FIXED_WIDTH.
Public Function TableRow_make(ByVal columnWidgets As Collection, Optional ByVal rowHeight As Long = 0) As WIDGET
    Dim result As WIDGET, d As TableRow, i As Long, widget As WIDGET
    Set result = WIDGET_alloc(K_TableRow)
    TableRow_construct result
    Set d = TableRow_(result)

    For i = 1 To columnWidgets.Count
        d.columnWidgets.Add columnWidgets(i)
    Next

    If rowHeight <= 0 Then
        ' expand row height to match maximum height of any contained widget
        rowHeight = 0
        For i = 1 To d.columnWidgets.Count
            Set widget = d.columnWidgets(i)
            rowHeight = MAX_(rowHeight, widget.height())
        Next
        ' vertically center any column widgets that are less than the rowHeight
        For i = 1 To d.columnWidgets.Count
            Set widget = d.columnWidgets(i)
            If widget.height() < rowHeight Then
                widget.setGeometry widget.x(), (rowHeight - widget.height()) \ 2, widget.width(), widget.height()
            End If
        Next
    Else
        ' use fixed row height
        For i = 1 To d.columnWidgets.Count
            Set widget = d.columnWidgets(i)
            widget.setGeometry widget.x(), 0, widget.width(), rowHeight
        Next
    End If
    result.setGeometry result.x(), result.y(), result.width(), rowHeight

    For i = 1 To d.columnWidgets.Count
        result.attach d.columnWidgets(i)
    Next

    Set TableRow_make = result
End Function

' Set whether the row background highlights on mouse-over
Public Sub TableRow_setHighlightsOnMouseOver(ByVal self As WIDGET, ByVal value As Boolean)
    TableRow_(self).highlightsOnMouseOver = value
End Sub

Public Sub TableRow_setHighlightColor(ByVal self As WIDGET, ByVal newHighlightColor As Long)
    TableRow_(self).highlightColor = newHighlightColor
End Sub

' Set whether the row draws an outer border (optional<PIELIGHT> borderColor)
Public Sub TableRow_setDrawBorder(ByVal self As WIDGET, ByVal has_borderColor As Boolean, Optional ByVal newBorderColor As Long = 0)
    TableRow_(self).has_borderColor = has_borderColor
    TableRow_(self).borderColor = newBorderColor
End Sub

' Set a background color for the row (optional<PIELIGHT> backgroundColor)
Public Sub TableRow_setBackgroundColor(ByVal self As WIDGET, ByVal has_backgroundColor As Boolean, Optional ByVal newBackgroundColor As Long = 0)
    TableRow_(self).has_backgroundColor = has_backgroundColor
    TableRow_(self).backgroundColor = newBackgroundColor
End Sub

' Set whether row is "disabled"
Public Sub TableRow_setDisabled(ByVal self As WIDGET, ByVal disabled As Boolean)
    TableRow_(self).disabledRow = disabled
    If disabled Then
        W_BUTTON_setState self, WBUT_DISABLE
    Else
        W_BUTTON_setState self, 0
    End If
End Sub

' Set row disable overlay color
Public Sub TableRow_setDisabledColor(ByVal self As WIDGET, ByVal newDisabledColor As Long)
    TableRow_(self).disabledColor = newDisabledColor
End Sub

' Get the total idealWidth() returned by all column widgets in this row (does not include padding)
Public Function TableRow_getColumnTotalContentIdealWidth(ByVal self As WIDGET) As Long
    Dim d As TableRow, totalContentIdealWidth As Long, i As Long, column As WIDGET
    Set d = TableRow_(self)
    totalContentIdealWidth = 0
    For i = 1 To d.columnWidgets.Count
        Set column = d.columnWidgets(i)
        totalContentIdealWidth = totalContentIdealWidth + column.idealWidth()
    Next
    TableRow_getColumnTotalContentIdealWidth = totalContentIdealWidth
End Function

Public Sub TableRow_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As TableRow, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Set d = TableRow_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    If d.highlightsOnMouseOver And TableRow_isMouseOverRowOrChildren(self) And Not d.disabledRow Then
        pie_UniTransBoxFill x0, y0, x1, y1, d.highlightColor
    ElseIf d.has_backgroundColor Then
        pie_UniTransBoxFill x0, y0, x1, y1, d.backgroundColor
    End If

    If d.has_borderColor Then
        iV_Box x0, y0, x1, y1, d.borderColor
    End If
End Sub

Public Sub TableRow_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim d As TableRow, xOffset As Long, yOffset As Long, x0 As Long, y0 As Long
    Set d = TableRow_(self)
    WIDGET_displayRecursive self, context

    ' *after* displaying all children, draw the disabled overlay (if needed)
    If Not d.disabledRow Then
        Exit Sub
    End If

    If Not WidgetGraphicsContext_clipContains(context, self.geometry()) Then
        Exit Sub
    End If

    xOffset = WidgetGraphicsContext_getXOffset(context)
    yOffset = WidgetGraphicsContext_getYOffset(context)

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), d.disabledColor
End Sub

Public Function TableRow_hitTest(ByVal self As WIDGET, ByVal x As Long, ByVal y As Long) As Boolean
    If x >= TableRow_(self).maxDisplayedColumnX1 Then
        TableRow_hitTest = False
        Exit Function
    End If
    TableRow_hitTest = WIDGET_hitTest(self, x, y)  ' W_BUTTON::hitTest (W_BUTTON has WIDGET's)
End Function

Public Function TableRow_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim d As TableRow
    Set d = TableRow_(self)
    ' if findMouseTargetRecursive was called for this row, it means the mouse is over it
    ' (see WIDGET::findMouseTargetRecursive)
    d.has_lastFrameMouseIsOverRowOrChildren = True
    d.lastFrameMouseIsOverRowOrChildren = frameGetFrameNumber()

    If d.disabledRow Then
        Set TableRow_findMouseTargetRecursive = self
        Exit Function
    End If
    Set TableRow_findMouseTargetRecursive = WIDGET_findMouseTargetRecursive(self, psContext, key, wasPressed)
End Function

Private Function TableRow_isMouseOverRowOrChildren(ByVal self As WIDGET) As Boolean
    Dim d As TableRow
    Set d = TableRow_(self)
    If Not d.has_lastFrameMouseIsOverRowOrChildren Then
        TableRow_isMouseOverRowOrChildren = False
        Exit Function
    End If
    TableRow_isMouseOverRowOrChildren = d.lastFrameMouseIsOverRowOrChildren = frameGetFrameNumber()
End Function

Public Sub TableRow_resizeColumns(ByVal self As WIDGET, newColumnWidths() As Long, ByVal newColumnWidthsCount As Long, ByVal columnPadding As Long)
    Dim d As TableRow, lastColumnEndX As Long, colIdx As Long, w As WIDGET
    Set d = TableRow_(self)
    columnPadding = MAX_(columnPadding, 0)
    lastColumnEndX = -columnPadding
    colIdx = 0
    Do While colIdx < MIN_(newColumnWidthsCount, d.columnWidgets.Count)
        Set w = d.columnWidgets(colIdx + 1)
        w.setGeometry lastColumnEndX + columnPadding + columnPadding, w.y(), newColumnWidths(colIdx), w.height()
        w.show
        lastColumnEndX = (w.x() + w.width())
        colIdx = colIdx + 1
    Loop
    d.maxDisplayedColumnX1 = lastColumnEndX
    ' hide any additional column widgets
    Do While colIdx < d.columnWidgets.Count
        Set w = d.columnWidgets(colIdx + 1)
        w.hide
        colIdx = colIdx + 1
    Loop
End Sub

Public Function TableRow_numColumns(ByVal self As WIDGET) As Long
    TableRow_numColumns = TableRow_(self).columnWidgets.Count
End Function

Public Function TableRow_getWidgetAtColumn(ByVal self As WIDGET, ByVal col As Long) As WIDGET
    Dim d As TableRow
    Set d = TableRow_(self)
    If Not (col >= 0 And col < d.columnWidgets.Count) Then ASSERT_FAILURE "col < columnWidgets.size()", "lib/widget/table.cpp:227", "getWidgetAtColumn", "Invalid column index: %zu", col: Exit Function
    Set TableRow_getWidgetAtColumn = d.columnWidgets(col + 1)
End Function

'--------------------------------------------------------------------------
' MARK: - TableHeader

' std::make_shared<TableHeader>(parentTable)
Private Function TableHeader_new(ByVal parentTable As WIDGET) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_TableHeader)
    TableHeader_construct w, parentTable
    Set TableHeader_new = w
End Function

' TableHeader::TableHeader(const std::shared_ptr<ScrollableTableWidget>& _parentTable)
Public Sub TableHeader_construct(ByVal self As WIDGET, ByVal parentTable As WIDGET)
    Dim d As TableHeader
    W_FORM_construct self
    self.kind = K_TableHeader ' (VB) virtual calls from here on reach TableHeader's overrides
    Set d = TableHeader_(self)
    d.userResizableHeaders = True
    d.has_colBeingResized = False
    d.has_dragStart = False
    If Not (Not parentTable Is Nothing) Then ASSERT_FAILURE "_parentTable != nullptr", "lib/widget/table.cpp:269", "TableHeader", "Null parent table"
    d.parentTable = ScrollableTableWidget_(parentTable).weakSerial
    d.columnPadding = ScrollableTableWidget_getColumnPadding(parentTable)
End Sub

Public Sub TableHeader_addColumn(ByVal self As WIDGET, column As TableColumn)
    Dim d As TableHeader, columnX As Long, colHeaderY As Long, back As TableColumn
    Set d = TableHeader_(self)
    self.attach column.columnWidget
    ' Position it to the right of the prior column (child)
    columnX = d.columnPadding.x
    colHeaderY = 0
    If d.columnsCount > 0 Then
        back = d.columns(d.columnsCount - 1)
        columnX = columnX + back.columnWidget.x() + back.columnWidget.width() + d.columnPadding.x
    End If
    column.columnWidget.setGeometry columnX, colHeaderY, column.columnWidget.width(), TableHeader_columnWidgetHeight(self)

    d.columns_push_back column
End Sub

Public Sub TableHeader_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As TableHeader, x0 As Long, y0 As Long, x1 As Long, y1 As Long, i As Long, child As TableColumn, xPos As Long
    Set d = TableHeader_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height() - 1
    iV_TransBoxFill x0, y0, x1, y1
    iV_Line x0, y1, x1, y1, psPalette(WZCOL_MENU_SEPARATOR)

    ' draw a line between every column widget
    For i = 0 To d.columnsCount - 1
        child = d.columns(i)
        ' column lines
        xPos = x0 + child.columnWidget.x() + child.columnWidget.width() + d.columnPadding.x
        iV_Line xPos, y0 + 5, xPos, y1 - 5, psPalette(WZCOL_MENU_SEPARATOR)
    Next
End Sub

Public Function TableHeader_capturesMouseDrag(ByVal self As WIDGET, ByVal key As WIDGET_KEY) As Boolean
    TableHeader_capturesMouseDrag = TableHeader_(self).userResizableHeaders
End Function

Public Sub TableHeader_mouseDragged(ByVal self As WIDGET, ByVal key As WIDGET_KEY, start As W_CONTEXT, current As W_CONTEXT)
    Dim d As TableHeader, mx As Long, colIdx As Long, columnWidget As WIDGET, col_x1 As Long, col As TableColumn
    Dim currentMousePos As Vector2i, dragDelta As Vector2i, priorColumnWidth As Long, newProposedColumnWidth As Long
    Dim table As WIDGET, resultValue As Long, currentDragLogicalPos As Vector2i
    Set d = TableHeader_(self)
    If Not d.userResizableHeaders Or key <> WKEY_PRIMARY Or d.columnsCount = 0 Then
        Exit Sub
    End If

    If Not d.has_dragStart Then
        d.has_dragStart = True
        d.dragStart = Vector2i_(start.mx, start.my)

        ' determine the column that is being resized
        mx = start.mx - self.x()
        d.has_colBeingResized = False
        colIdx = d.columnsCount - 1
        Do                              ' termination handled in loop body
            col = d.columns(colIdx)
            Set columnWidget = col.columnWidget
            col_x1 = columnWidget.x() + columnWidget.width()
            If mx > col_x1 Then
                If col.resizeBehavior <> TableColumn_ResizeBehavior_FIXED_WIDTH Then
                    d.has_colBeingResized = True
                    d.colBeingResized = colIdx
                End If
                Exit Do
            End If
            If colIdx = 0 Then
                Exit Do
            End If
            colIdx = colIdx - 1
        Loop
        If Not d.has_colBeingResized Then Exit Sub
    End If

    If d.has_dragStart And d.has_colBeingResized Then
        ' handle column resize while dragging
        currentMousePos = Vector2i_(current.mx, current.my)
        dragDelta = Vector2i_(currentMousePos.x - d.dragStart.x, currentMousePos.y - d.dragStart.y)

        col = d.columns(d.colBeingResized)
        priorColumnWidth = col.columnWidget.width()
        newProposedColumnWidth = MAX_(priorColumnWidth + dragDelta.x, 0)
        Set table = WeakWidget_lock(d.parentTable)
        If Not table Is Nothing Then
            If ScrollableTableWidget_header_changeColumnWidth(table, d.colBeingResized, newProposedColumnWidth, resultValue) Then
                currentDragLogicalPos = d.dragStart
                currentDragLogicalPos.x = currentDragLogicalPos.x + (resultValue - priorColumnWidth)
                d.dragStart = currentDragLogicalPos
            End If
        End If
    End If
End Sub

Public Sub TableHeader_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    ' currently, no-op
End Sub

Public Sub TableHeader_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    TableHeader_(self).has_colBeingResized = False
    TableHeader_(self).has_dragStart = False
End Sub

Public Function TableHeader_isUserDraggingColumnHeader(ByVal self As WIDGET) As Boolean
    TableHeader_isUserDraggingColumnHeader = TableHeader_(self).userResizableHeaders And TableHeader_(self).has_dragStart
End Function

Public Sub TableHeader_setColumnPadding(ByVal self As WIDGET, padding_ As Vector2i)
    Dim d As TableHeader, lastColumnX As Long, columnIdx As Long, columnX As Long, colHeaderY As Long, prior As TableColumn, column As TableColumn
    Set d = TableHeader_(self)
    d.columnPadding = padding_

    ' Reposition all column widgets
    lastColumnX = d.columnPadding.x
    For columnIdx = 0 To d.columnsCount - 1
        ' Position it to the right of the prior column (child)
        columnX = lastColumnX
        colHeaderY = 0
        If columnIdx > 0 Then
            prior = d.columns(columnIdx - 1)
            columnX = columnX + prior.columnWidget.x() + prior.columnWidget.width() + d.columnPadding.x
        End If
        column = d.columns(columnIdx)
        column.columnWidget.setGeometry columnX, colHeaderY, column.columnWidget.width(), TableHeader_columnWidgetHeight(self)
    Next
End Sub

Public Sub TableHeader_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    ' currently, no-op
End Sub

Public Sub TableHeader_geometryChanged(ByVal self As WIDGET)
    Dim d As TableHeader, i As Long, column As TableColumn, columnWidget As WIDGET
    Set d = TableHeader_(self)
    ' resize column header widgets height based on parent TableHeader height
    For i = 0 To d.columnsCount - 1
        column = d.columns(i)
        Set columnWidget = column.columnWidget
        columnWidget.setGeometry columnWidget.x(), columnWidget.y(), columnWidget.width(), TableHeader_columnWidgetHeight(self)
    Next
End Sub

Private Function TableHeader_columnWidgetHeight(ByVal self As WIDGET) As Long
    TableHeader_columnWidgetHeight = self.height() - 2
End Function

Public Sub TableHeader_changeColumnWidths(ByVal self As WIDGET, newColumnWidths() As Long, ByVal newColumnWidthsCount As Long)
    Dim d As TableHeader, columnX As Long, i As Long, column As TableColumn, columnWidget As WIDGET
    Set d = TableHeader_(self)
    columnX = d.columnPadding.x
    i = 0
    Do While i < MIN_(newColumnWidthsCount, d.columnsCount)
        column = d.columns(i)
        Set columnWidget = column.columnWidget
        columnWidget.setGeometry columnX, columnWidget.y(), newColumnWidths(i), columnWidget.height()
        columnX = columnX + d.columnPadding.x + columnWidget.width() + d.columnPadding.x
        i = i + 1
    Loop
    Do While i < d.columnsCount
        column = d.columns(i)
        Set columnWidget = column.columnWidget
        columnWidget.setGeometry columnX, columnWidget.y(), columnWidget.width(), columnWidget.height()
        columnX = columnX + d.columnPadding.x + columnWidget.width() + d.columnPadding.x
        i = i + 1
    Loop
End Sub

'--------------------------------------------------------------------------
' MARK: - ScrollableTableWidget

' ScrollableTableWidget::ScrollableTableWidget()
Public Sub ScrollableTableWidget_construct(ByVal self As WIDGET)
    Dim d As ScrollableTableWidget
    WIDGET_construct self
    self.kind = K_ScrollableTableWidget ' (VB) virtual calls from here on reach ScrollableTableWidget's overrides
    Set d = ScrollableTableWidget_(self)
    d.weakSerial = WeakWidget_register(self)
    d.columnPadding = Vector2i_(TABLE_COL_PADDING, 0)
    d.drawColumnLines = False
    d.userDidResizeColumnWidths = False
End Sub

Public Sub ScrollableTableWidget_destruct(ByVal self As WIDGET)
    WeakWidget_unregister ScrollableTableWidget_(self).weakSerial
End Sub

' Make a new ScrollableTableWidget
'
' Note: The TableColumn.columnWidget widgets are expected to have pre-set widths, which are used to
'       initialize the ScrollableTableWidget's columnWidths.
Public Function ScrollableTableWidget_make(columns() As TableColumn, ByVal columnsCount As Long, Optional ByVal headerHeight As Long = -1) As WIDGET
    Dim result As WIDGET, d As ScrollableTableWidget, i As Long, column As TableColumn
    Set result = WIDGET_alloc(K_ScrollableTableWidget)
    ScrollableTableWidget_construct result
    Set d = ScrollableTableWidget_(result)

    ' Add table header row
    If headerHeight <= 0 Then
        headerHeight = TABLE_COL_HEADER_HEIGHT
    End If
    Set d.header = TableHeader_new(result)
    result.attach d.header
    d.header.setGeometry 0, 0, result.width(), headerHeight ' initialize height
    d.header.setCalcLayout WzFunc_new(FN_TableHeader_calcLayout)

    ' Add columns
    For i = 0 To columnsCount - 1
        d.tableColumns_push_back columns(i)
    Next
    For i = 0 To d.tableColumnsCount - 1
        column = d.tableColumns(i)
        TableHeader_addColumn d.header, column
        d.columnWidths_push_back column.columnWidget.width()
    Next
    d.minColumnWidths_resize d.tableColumnsCount     ' (filled with 0)

    ' Create and add scrollableList
    Set d.scrollableList = ScrollableListWidget_make()
    result.attach d.scrollableList
    ScrollableListWidget_setBackgroundColor d.scrollableList, psPalette(WZCOL_TRANSPARENT_BOX)
    ' Position scrollableList below column headers
    d.scrollableList.setCalcLayout WzFunc_new(FN_ScrollableTableWidget_listCalcLayout)

    Set ScrollableTableWidget_make = result
End Function

Public Sub ScrollableTableWidget_geometryChanged(ByVal self As WIDGET)
    ' Trigger calcLayout on children
    ScrollableTableWidget_(self).header.callCalcLayout
    ScrollableTableWidget_(self).scrollableList.callCalcLayout
End Sub

Public Function ScrollableTableWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As ScrollableTableWidget
    Set d = ScrollableTableWidget_(self)
    If Not (Not d.scrollableList Is Nothing) Then ASSERT_FAILURE "scrollableList != nullptr", "lib/widget/table.cpp:506", "idealHeight", "scrollableList not yet initialized?": ScrollableTableWidget_idealHeight = 0: Exit Function
    ScrollableTableWidget_idealHeight = d.scrollableList.y() + d.scrollableList.idealHeight()
End Function

' Change the table background color
Public Sub ScrollableTableWidget_setBackgroundColor(ByVal self As WIDGET, ByVal color As Long)
    ScrollableListWidget_setBackgroundColor ScrollableTableWidget_(self).scrollableList, color
End Sub

Public Function ScrollableTableWidget_getScrollPosition(ByVal self As WIDGET) As Long
    ScrollableTableWidget_getScrollPosition = ScrollableListWidget_getScrollPosition(ScrollableTableWidget_(self).scrollableList)
End Function

Public Sub ScrollableTableWidget_setScrollPosition(ByVal self As WIDGET, ByVal newPosition As Long)
    ScrollableListWidget_setScrollPosition ScrollableTableWidget_(self).scrollableList, newPosition
End Sub

' Add a new table row
' See: ``TableRow``
Public Sub ScrollableTableWidget_addRow(ByVal self As WIDGET, ByVal row As WIDGET)
    Dim d As ScrollableTableWidget, widths() As Long, n As Long
    Set d = ScrollableTableWidget_(self)
    If Not (Not row Is Nothing) Then ASSERT_FAILURE "row != nullptr", "lib/widget/table.cpp:527", "addRow", "row is null": Exit Sub
    If Not (TableRow_numColumns(row) = d.columnWidthsCount) Then ASSERT_FAILURE "row->numColumns() == columnWidths.size()", "lib/widget/table.cpp:528", "addRow", "Unexpected number of columns in row: %zu", TableRow_numColumns(row): Exit Sub
    n = d.columnWidths_copyTo(widths)
    vt_TableRow_resizeColumns row, widths, n, d.columnPadding.x
    ' Add row to internal ScrollableListWidget, and rows list
    ScrollableListWidget_addItem d.scrollableList, row
    d.rows.Add row
End Sub

Public Sub ScrollableTableWidget_clearRows(ByVal self As WIDGET)
    ScrollableListWidget_clear ScrollableTableWidget_(self).scrollableList
    Set ScrollableTableWidget_(self).rows = New Collection
End Sub

Public Function ScrollableTableWidget_getRows(ByVal self As WIDGET) As Collection
    Set ScrollableTableWidget_getRows = ScrollableTableWidget_(self).rows
End Function

' Show / hide header
Public Sub ScrollableTableWidget_setHeaderVisible(ByVal self As WIDGET, ByVal visible As Boolean)
    Dim d As ScrollableTableWidget
    Set d = ScrollableTableWidget_(self)
    If visible = d.header.visible() Then
        Exit Sub
    End If
    d.header.show visible
    self.geometryChanged
End Sub

' Disable / enable a row
Public Sub ScrollableTableWidget_setRowDisabled(ByVal self As WIDGET, ByVal row As Long, ByVal disabled As Boolean)
    Dim d As ScrollableTableWidget
    Set d = ScrollableTableWidget_(self)
    If Not (row >= 0 And row < d.rows.Count) Then ASSERT_FAILURE "row < rows.size()", "lib/widget/table.cpp:553", "setRowDisabled", "Invalid row idx: %zu", row: Exit Sub
    TableRow_setDisabled d.rows(row + 1), disabled
End Sub

' Get the current column widths (copied into widths; returns their number)
Public Function ScrollableTableWidget_getColumnWidths(ByVal self As WIDGET, widths() As Long) As Long
    ScrollableTableWidget_getColumnWidths = ScrollableTableWidget_(self).columnWidths_copyTo(widths)
End Function

' Propose a new set of column widths
' The ScrollableTableWidget will automatically adjust the submitted widths for various constraints
' (like the minimum column widths, total maximum width, which columns can be resized / expanded-to-fill-space, fixed width sizes, etc)
' It is expected that newColumnWidths.size() will be equal to the number of columns in the table.
Public Function ScrollableTableWidget_changeColumnWidths(ByVal self As WIDGET, newColumnWidths() As Long, ByVal newColumnWidthsCount As Long, Optional ByVal overrideUserColumnResizing As Boolean = False) As Boolean
    Dim d As ScrollableTableWidget, noPriority() As Long
    Set d = ScrollableTableWidget_(self)
    If Not (newColumnWidthsCount = d.columnWidthsCount) Then ASSERT_FAILURE "newColumnWidths.size() == columnWidths.size()", "lib/widget/table.cpp:559", "changeColumnWidths", "newColumnWidths.size (%zu) does not match existing number of columns (%zu)", newColumnWidthsCount, d.columnWidthsCount: ScrollableTableWidget_changeColumnWidths = False: Exit Function
    If d.userDidResizeColumnWidths And Not overrideUserColumnResizing Then
        ScrollableTableWidget_changeColumnWidths = False
        Exit Function
    End If
    If Not ScrollableTableWidget_relayoutColumns(self, newColumnWidths, newColumnWidthsCount, noPriority, 0) Then
        debug_ 566, LOG_WZ, "changeColumnWidths", "The proposed column widths are not possible to achieve given the table width and layout constraints."
        ScrollableTableWidget_changeColumnWidths = False
        Exit Function
    End If
    ScrollableTableWidget_changeColumnWidths = True
End Function

' Change a single column width
' The ScrollableTableWidget will preferentially re-layout / flow the other column widths according to the table's column constraints.
' However, it will fallback to adjusting the newColumnWidth for col if necessary to satisfy all constraints.
' Returns:
' - `nullopt` on failure (if there was no way to change the specified column to the newColumnWidth and layout the other columns)
' - the actual newColumnWidth, on success, which may differ from the input newColumnWidth if required to satisfy layout constraints
Public Function ScrollableTableWidget_changeColumnWidth(ByVal self As WIDGET, ByVal col As Long, ByVal newColumnWidth As Long, value As Long, Optional ByVal overrideUserColumnResizing As Boolean = False) As Boolean
    Dim d As ScrollableTableWidget, newProposedColumnWidths() As Long, n As Long, priority(0 To 0) As Long
    Set d = ScrollableTableWidget_(self)
    If Not (col >= 0 And col < d.columnWidthsCount) Then ASSERT_FAILURE "col < columnWidths.size()", "lib/widget/table.cpp:574", "changeColumnWidth", "Invalid column index: %zu", col: Exit Function
    If d.userDidResizeColumnWidths And Not overrideUserColumnResizing Then
        Exit Function                   ' nullopt
    End If
    n = d.columnWidths_copyTo(newProposedColumnWidths)
    newProposedColumnWidths(col) = newColumnWidth
    priority(0) = col
    If Not ScrollableTableWidget_relayoutColumns(self, newProposedColumnWidths, n, priority, 1) Then
        Exit Function                   ' nullopt
    End If
    value = d.columnWidth(col)
    ScrollableTableWidget_changeColumnWidth = True
End Function

' Called specifically from the header to inform the Table that the user has resized the columns
Public Function ScrollableTableWidget_header_changeColumnWidth(ByVal self As WIDGET, ByVal col As Long, ByVal newColumnWidth As Long, value As Long) As Boolean
    Dim result As Boolean
    result = ScrollableTableWidget_changeColumnWidth(self, col, newColumnWidth, value, True)
    ScrollableTableWidget_(self).userDidResizeColumnWidths = True
    ScrollableTableWidget_header_changeColumnWidth = result
End Function

' Get the current minimum column width layout constraints (copied into widths; returns their number)
Public Function ScrollableTableWidget_getMinimumColumnWidths(ByVal self As WIDGET, widths() As Long) As Long
    ScrollableTableWidget_getMinimumColumnWidths = ScrollableTableWidget_(self).minColumnWidths_copyTo(widths)
End Function

' Add minimum column width layout constraints for all columns
Public Sub ScrollableTableWidget_setMinimumColumnWidths(ByVal self As WIDGET, newMinColumnWidths() As Long, ByVal newMinColumnWidthsCount As Long)
    Dim d As ScrollableTableWidget, widths() As Long, n As Long, noPriority() As Long
    Set d = ScrollableTableWidget_(self)
    If Not (newMinColumnWidthsCount = d.columnWidthsCount) Then ASSERT_FAILURE "newMinColumnWidths.size() == columnWidths.size()", "lib/widget/table.cpp:598", "setMinimumColumnWidths", "newMinColumnWidths.size (%zu) does not match existing number of columns (%zu)", newMinColumnWidthsCount, d.columnWidthsCount: Exit Sub
    d.minColumnWidths_assign newMinColumnWidths, newMinColumnWidthsCount
    n = d.columnWidths_copyTo(widths)
    ScrollableTableWidget_relayoutColumns self, widths, n, noPriority, 0
End Sub

' Add or update a minimum column width layout constraint for a single column
' Note: This triggers a relayout, so if you are setting constraints on multiple columns you should use:
' setMinimumColumnWidths(const std::vector<size_t>& newMinColumnWidths)
Public Sub ScrollableTableWidget_setMinimumColumnWidth(ByVal self As WIDGET, ByVal col As Long, ByVal newMinColumnWidth As Long)
    Dim d As ScrollableTableWidget, widths() As Long, n As Long, noPriority() As Long
    Set d = ScrollableTableWidget_(self)
    If Not (col >= 0 And col < d.minColumnWidthsCount) Then ASSERT_FAILURE "col < minColumnWidths.size()", "lib/widget/table.cpp:605", "setMinimumColumnWidth", "Invalid column index: %zu", col: Exit Sub
    d.minColumnWidth(col) = newMinColumnWidth
    n = d.columnWidths_copyTo(widths)
    ScrollableTableWidget_relayoutColumns self, widths, n, noPriority, 0
End Sub

' Get the maximum idealWidth() returned by any of the row widgets in the specified column
Public Function ScrollableTableWidget_getColumnMaxContentIdealWidth(ByVal self As WIDGET, ByVal col As Long) As Long
    Dim d As ScrollableTableWidget, maxIdealWidth As Long, i As Long, pColWidget As WIDGET
    Set d = ScrollableTableWidget_(self)
    maxIdealWidth = 0
    For i = 1 To d.rows.Count
        Set pColWidget = TableRow_getWidgetAtColumn(d.rows(i), col)
        If pColWidget Is Nothing Then GoTo continue_
        maxIdealWidth = MAX_(maxIdealWidth, pColWidget.idealWidth())
continue_:
    Next
    ScrollableTableWidget_getColumnMaxContentIdealWidth = maxIdealWidth
End Function

Public Function ScrollableTableWidget_getNumRows(ByVal self As WIDGET) As Long
    ScrollableTableWidget_getNumRows = ScrollableTableWidget_(self).rows.Count
End Function

Public Function ScrollableTableWidget_getNumColumns(ByVal self As WIDGET) As Long
    ScrollableTableWidget_getNumColumns = ScrollableTableWidget_(self).tableColumnsCount
End Function

' (VB) getColumnIndexes with getShrinkableColumnIndexes's lambda; returns the indexes' number
Private Function ScrollableTableWidget_getShrinkableColumnIndexes(ByVal self As WIDGET, currentColumnWidths() As Long, ByVal currentColumnWidthsCount As Long, columnIndexes() As Long) As Long
    Dim d As ScrollableTableWidget, i As Long, n As Long, col As TableColumn
    Set d = ScrollableTableWidget_(self)
    n = 0
    For i = 0 To d.tableColumnsCount - 1
        col = d.tableColumns(i)
        If col.resizeBehavior = TableColumn_ResizeBehavior_FIXED_WIDTH Then
            GoTo continue_
        End If
        If MIN_(currentColumnWidthsCount, d.minColumnWidthsCount) > i Then
            If currentColumnWidths(i) <= d.minColumnWidth(i) Then
                GoTo continue_
            End If
        End If
        ReDim Preserve columnIndexes(0 To n)
        columnIndexes(n) = i
        n = n + 1
continue_:
    Next
    ScrollableTableWidget_getShrinkableColumnIndexes = n
End Function

' (VB) getColumnIndexes with getExpandToFillColumnIndexes's lambda; returns the indexes' number
Private Function ScrollableTableWidget_getExpandToFillColumnIndexes(ByVal self As WIDGET, columnIndexes() As Long) As Long
    Dim d As ScrollableTableWidget, i As Long, n As Long, col As TableColumn
    Set d = ScrollableTableWidget_(self)
    n = 0
    For i = 0 To d.tableColumnsCount - 1
        col = d.tableColumns(i)
        If col.resizeBehavior = TableColumn_ResizeBehavior_RESIZABLE_AUTOEXPAND Then
            ReDim Preserve columnIndexes(0 To n)
            columnIndexes(n) = i
            n = n + 1
        End If
    Next
    ScrollableTableWidget_getExpandToFillColumnIndexes = n
End Function

Private Sub ScrollableTableWidget_updateColumnWidths(ByVal self As WIDGET)
    Dim d As ScrollableTableWidget, widths() As Long, n As Long, i As Long
    Set d = ScrollableTableWidget_(self)
    n = d.columnWidths_copyTo(widths)
    ' actually resize header and row columns
    TableHeader_changeColumnWidths d.header, widths, n
    For i = 1 To d.rows.Count
        vt_TableRow_resizeColumns d.rows(i), widths, n, d.columnPadding.x
    Next
End Sub

' (VB) maxColumnWidthReduction for a column (shrinkColumnWidths's sort and loop)
Private Function ScrollableTableWidget_maxColumnWidthReduction(ByVal d As ScrollableTableWidget, proposedColumnWidths() As Long, ByVal colIdx As Long) As Long
    ScrollableTableWidget_maxColumnWidthReduction = proposedColumnWidths(colIdx)
    If d.minColumnWidthsCount > colIdx Then
        If proposedColumnWidths(colIdx) > d.minColumnWidth(colIdx) Then
            ScrollableTableWidget_maxColumnWidthReduction = proposedColumnWidths(colIdx) - d.minColumnWidth(colIdx)
        Else
            ScrollableTableWidget_maxColumnWidthReduction = 0
        End If
    End If
End Function

' (VB) relayoutColumns's shrinkColumnWidths lambda ([this, &proposedColumnWidths]); shrinkableColumnIndexes by value
Private Function ScrollableTableWidget_shrinkColumnWidths(ByVal self As WIDGET, proposedColumnWidths() As Long, shrinkableIn() As Long, ByVal shrinkableCount As Long, ByVal extraWidth As Long) As Long
    Dim d As ScrollableTableWidget, shrinkableColumnIndexes() As Long, currentTotalColumnWidthReduction As Long
    Dim i As Long, j As Long, tmp As Long, colIdx As Long, remainingExtraWidth As Long, widthReductionAmount As Long, maxColumnWidthReduction As Long
    Set d = ScrollableTableWidget_(self)
    currentTotalColumnWidthReduction = 0
    If shrinkableCount = 0 Then Exit Function
    ReDim shrinkableColumnIndexes(0 To shrinkableCount - 1)
    For i = 0 To shrinkableCount - 1
        shrinkableColumnIndexes(i) = shrinkableIn(i)
    Next
    ' order shrinkableColumnIndexes by maxColumnWidthReduction for associated column
    For i = 1 To shrinkableCount - 1
        tmp = shrinkableColumnIndexes(i)
        j = i - 1
        Do While j >= 0
            If Not (ScrollableTableWidget_maxColumnWidthReduction(d, proposedColumnWidths, tmp) < ScrollableTableWidget_maxColumnWidthReduction(d, proposedColumnWidths, shrinkableColumnIndexes(j))) Then Exit Do
            shrinkableColumnIndexes(j + 1) = shrinkableColumnIndexes(j)
            j = j - 1
        Loop
        shrinkableColumnIndexes(j + 1) = tmp
    Next
    ' try to shrink all shrinkable columns proportionally
    For i = 0 To shrinkableCount - 1
        colIdx = shrinkableColumnIndexes(i)
        remainingExtraWidth = extraWidth - currentTotalColumnWidthReduction
        If i < shrinkableCount - 1 Then
            widthReductionAmount = remainingExtraWidth \ (shrinkableCount - i)
        Else
            widthReductionAmount = remainingExtraWidth
        End If
        maxColumnWidthReduction = ScrollableTableWidget_maxColumnWidthReduction(d, proposedColumnWidths, colIdx)
        widthReductionAmount = MIN_(widthReductionAmount, maxColumnWidthReduction)
        proposedColumnWidths(colIdx) = proposedColumnWidths(colIdx) - widthReductionAmount
        currentTotalColumnWidthReduction = currentTotalColumnWidthReduction + widthReductionAmount
    Next
    ScrollableTableWidget_shrinkColumnWidths = currentTotalColumnWidthReduction
End Function

Private Function sumOf(values() As Long, ByVal n As Long) As Long
    Dim i As Long
    For i = 0 To n - 1
        sumOf = sumOf + values(i)
    Next
End Function

Private Function ScrollableTableWidget_relayoutColumns(ByVal self As WIDGET, proposedIn() As Long, ByVal proposedCount As Long, priorityIndexes() As Long, ByVal priorityCount As Long) As Boolean
    Dim d As ScrollableTableWidget, proposedColumnWidths() As Long, colIndexesIncreasedToMinimumSize() As Long, increasedCount As Long
    Dim i As Long, j As Long, maxColumnWidthAvailable As Long, totalColWidth As Long, extraWidth As Long
    Dim allShrinkableColumnIndexes() As Long, allShrinkableCount As Long, shrinkableColumnIndexes() As Long, shrinkableCount As Long
    Dim ignoringPriorityIndexes As Boolean, currentTotalColumnWidthReduction As Long, isPriority As Boolean
    Dim expandToFillColumnIndexes() As Long, expandCount As Long, currentTotalColumnWidthIncrease As Long, colIdx As Long, widthIncreaseAmount As Long
    Set d = ScrollableTableWidget_(self)
    ' (std::vector<size_t> proposedColumnWidths: by value)
    If proposedCount > 0 Then
        ReDim proposedColumnWidths(0 To proposedCount - 1)
        For i = 0 To proposedCount - 1
            proposedColumnWidths(i) = proposedIn(i)
        Next
    End If

    ' respect any minimum column widths (first pass)
    increasedCount = 0
    For i = 0 To MIN_(proposedCount, d.minColumnWidthsCount) - 1
        If proposedColumnWidths(i) < d.minColumnWidth(i) Then
            proposedColumnWidths(i) = d.minColumnWidth(i)
            ReDim Preserve colIndexesIncreasedToMinimumSize(0 To increasedCount)
            colIndexesIncreasedToMinimumSize(increasedCount) = i
            increasedCount = increasedCount + 1
        End If
    Next

    maxColumnWidthAvailable = ScrollableTableWidget_getMaxColumnTotalWidth(self, proposedCount)
    totalColWidth = sumOf(proposedColumnWidths, proposedCount)
    If totalColWidth > maxColumnWidthAvailable Then
        ' need to shrink all shrinkable columns to fit
        extraWidth = totalColWidth - maxColumnWidthAvailable
        allShrinkableCount = ScrollableTableWidget_getShrinkableColumnIndexes(self, proposedColumnWidths, proposedCount, allShrinkableColumnIndexes)
        ' shrinkableColumnIndexes = allShrinkableColumnIndexes without the priorityIndexes
        shrinkableCount = 0
        For i = 0 To allShrinkableCount - 1
            isPriority = False
            For j = 0 To priorityCount - 1
                If priorityIndexes(j) = allShrinkableColumnIndexes(i) Then isPriority = True
            Next
            If Not isPriority Then
                ReDim Preserve shrinkableColumnIndexes(0 To shrinkableCount)
                shrinkableColumnIndexes(shrinkableCount) = allShrinkableColumnIndexes(i)
                shrinkableCount = shrinkableCount + 1
            End If
        Next
        ignoringPriorityIndexes = False
        If shrinkableCount = 0 Then
            ' ignore "priorityIndexes"
            shrinkableCount = allShrinkableCount
            If allShrinkableCount > 0 Then shrinkableColumnIndexes = allShrinkableColumnIndexes
            ignoringPriorityIndexes = True
        End If
        If shrinkableCount = 0 Then
            If Not (shrinkableCount <> 0) Then ASSERT_FAILURE "!shrinkableColumnIndexes.empty()", "lib/widget/table.cpp:689", "relayoutColumns", "All columns are fixed width (or non-shrinkable) but desired sizes exceed maxColumnWidthAvailable: %zu", maxColumnWidthAvailable
            ScrollableTableWidget_relayoutColumns = False
            Exit Function
        End If
        currentTotalColumnWidthReduction = ScrollableTableWidget_shrinkColumnWidths(self, proposedColumnWidths, shrinkableColumnIndexes, shrinkableCount, extraWidth)
        If currentTotalColumnWidthReduction < extraWidth Then
            If Not ignoringPriorityIndexes And priorityCount > 0 Then
                ' One more thing we can do - try to shrink the priority column indexes
                currentTotalColumnWidthReduction = currentTotalColumnWidthReduction + ScrollableTableWidget_shrinkColumnWidths(self, proposedColumnWidths, priorityIndexes, priorityCount, (extraWidth - currentTotalColumnWidthReduction))
            End If
            If currentTotalColumnWidthReduction < extraWidth Then
                ' Unable to shrink other columns - reject this proposed set of column widths
                ScrollableTableWidget_relayoutColumns = False
                Exit Function
            End If
        End If
        If Not (currentTotalColumnWidthReduction = extraWidth) Then ASSERT_FAILURE "currentTotalColumnWidthReduction == extraWidth", "lib/widget/table.cpp:742", "relayoutColumns", "Logic error"
        totalColWidth = sumOf(proposedColumnWidths, proposedCount)
    End If
    If totalColWidth < maxColumnWidthAvailable Then
        ' expand any "expand to fill" columns to proportionally take up the extra space
        extraWidth = maxColumnWidthAvailable - totalColWidth
        expandCount = ScrollableTableWidget_getExpandToFillColumnIndexes(self, expandToFillColumnIndexes)
        currentTotalColumnWidthIncrease = 0
        For i = 0 To expandCount - 1
            colIdx = expandToFillColumnIndexes(i)
            If i < expandCount - 1 Then
                widthIncreaseAmount = extraWidth \ expandCount
            Else
                widthIncreaseAmount = extraWidth - currentTotalColumnWidthIncrease
            End If
            proposedColumnWidths(colIdx) = proposedColumnWidths(colIdx) + widthIncreaseAmount
            currentTotalColumnWidthIncrease = currentTotalColumnWidthIncrease + widthIncreaseAmount
        Next
    End If
    totalColWidth = sumOf(proposedColumnWidths, proposedCount)

    ' store proposed column widths
    d.columnWidths_assign proposedColumnWidths, proposedCount

    ' actually resize header and row columns
    ScrollableTableWidget_updateColumnWidths self

    ScrollableTableWidget_relayoutColumns = True
End Function

Private Function ScrollableTableWidget_totalPaddingWidthFor(ByVal self As WIDGET, ByVal numColumns As Long) As Long
    Dim d As ScrollableTableWidget, totalPaddingWidth As Long
    Set d = ScrollableTableWidget_(self)
    ' TABLE_COL_PADDING is on all sides of all columns
    ' also factor in scrollbar width
    totalPaddingWidth = d.columnPadding.x + ((numColumns - 1) * 2 * d.columnPadding.x) + d.columnPadding.x + ScrollableListWidget_getScrollbarWidth(d.scrollableList)
    ScrollableTableWidget_totalPaddingWidthFor = totalPaddingWidth
End Function

' Get the maximum width that can be used by the column widths passed to changeColumnWidths, based on the current widget size (minus padding)
Public Function ScrollableTableWidget_getMaxColumnTotalWidth(ByVal self As WIDGET, ByVal numColumns As Long) As Long
    ScrollableTableWidget_getMaxColumnTotalWidth = self.width() - ScrollableTableWidget_totalPaddingWidthFor(self, numColumns)
End Function

' Get the table width needed to display columns with the specified total minimum column width
Public Function ScrollableTableWidget_getTableWidthNeededForTotalColumnWidth(ByVal self As WIDGET, ByVal numColumns As Long, ByVal totalMinimumColumnWidth As Long) As Long
    ScrollableTableWidget_getTableWidthNeededForTotalColumnWidth = ScrollableTableWidget_totalPaddingWidthFor(self, numColumns) + totalMinimumColumnWidth
End Function

Public Function ScrollableTableWidget_isUserDraggingColumnHeader(ByVal self As WIDGET) As Boolean
    ScrollableTableWidget_isUserDraggingColumnHeader = TableHeader_isUserDraggingColumnHeader(ScrollableTableWidget_(self).header)
End Function

Public Sub ScrollableTableWidget_setColumnPadding(ByVal self As WIDGET, newPadding As Vector2i)
    Dim d As ScrollableTableWidget, cur As Vector2i
    Set d = ScrollableTableWidget_(self)
    cur = d.columnPadding
    If newPadding.x = cur.x And newPadding.y = cur.y Then
        Exit Sub
    End If
    d.columnPadding = newPadding

    ' Need to inform header + resize columns and also resize columns in list
    TableHeader_setColumnPadding d.header, newPadding
    ScrollableTableWidget_updateColumnWidths self
End Sub

Public Function ScrollableTableWidget_getColumnPadding(ByVal self As WIDGET) As Vector2i
    ScrollableTableWidget_getColumnPadding = ScrollableTableWidget_(self).columnPadding
End Function

' Configure whether to draw column lines for list rows
Public Sub ScrollableTableWidget_setDrawColumnLines(ByVal self As WIDGET, ByVal bEnabled As Boolean)
    ScrollableTableWidget_(self).drawColumnLines = bEnabled
End Sub

Public Sub ScrollableTableWidget_setItemSpacing(ByVal self As WIDGET, ByVal value As Long)
    ScrollableListWidget_setItemSpacing ScrollableTableWidget_(self).scrollableList, value
End Sub

Public Sub ScrollableTableWidget_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim d As ScrollableTableWidget, xOffset As Long, yOffset As Long, x0 As Long, y0 As Long, y1 As Long
    Dim lines() As ivec4, linesCount As Long, lastColumnEndX As Long, colIdx As Long, columnXPos As Long, columRightLineXPos As Long
    Set d = ScrollableTableWidget_(self)
    WIDGET_displayRecursive self, context

    ' *after* displaying all children, draw the column lines (if desired)
    If Not d.drawColumnLines Then
        Exit Sub
    End If

    If Not WidgetGraphicsContext_clipContains(context, self.geometry()) Then
        Exit Sub
    End If

    xOffset = WidgetGraphicsContext_getXOffset(context)
    yOffset = WidgetGraphicsContext_getYOffset(context)

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset + d.scrollableList.y()
    y1 = y0 + d.scrollableList.height()

    linesCount = 0                      ' lines.clear(); lines.reserve(columnWidths.size());
    If d.columnWidthsCount > 0 Then ReDim lines(0 To d.columnWidthsCount - 1)

    ' draw a line between every column widget
    lastColumnEndX = -d.columnPadding.x
    For colIdx = 0 To d.columnWidthsCount - 1
        ' column lines
        columnXPos = lastColumnEndX + d.columnPadding.x + d.columnPadding.x
        columRightLineXPos = columnXPos + d.columnWidth(colIdx) + d.columnPadding.x

        lines(linesCount).x = x0 + columRightLineXPos
        lines(linesCount).y = y0
        lines(linesCount).z = x0 + columRightLineXPos
        lines(linesCount).w = y1
        linesCount = linesCount + 1
        lastColumnEndX = columnXPos + d.columnWidth(colIdx)
    Next

    If linesCount > 0 Then
        iV_Lines VarPtr(lines(0)), linesCount, psPalette(WZCOL_MENU_SEPARATOR)
    End If
End Sub
