Attribute VB_Name = "gridlayout_c"
' lib/widget/gridlayout.cpp, gridlayout.h: the widget class GridLayout
' (docs/CONVENTIONS.md, "Widgets"; its fields are GridLayout.cls). struct
' GridLayout::Placement is the UDT GridLayout_Placement. A
' nonstd::optional<grid_allocation::allocator> is a grid_allocation_allocator
' (Nothing = no value), a std::map<uint32_t, int32_t> a grid_offsets
' (gridallocation_c). The PlacementTransformer that getAllocationItems takes
' (placementColumn or placementRow) is a Boolean, True for placementColumn.
' (VB) widget class GridLayout : WIDGET
Option Explicit

Public Type GridLayout_Placement
    column As grid_allocation_slot
    row As grid_allocation_slot
    widget As WIDGET                    ' std::shared_ptr<WIDGET>
End Type

' Placement(grid_allocation::slot column, grid_allocation::slot row, std::shared_ptr<WIDGET> widget)
Public Function GridLayout_Placement_(column As grid_allocation_slot, row As grid_allocation_slot, ByVal widget As WIDGET) As GridLayout_Placement
    GridLayout_Placement_.column = column
    GridLayout_Placement_.row = row
    Set GridLayout_Placement_.widget = widget
End Function

' std::make_shared<GridLayout>()
Public Function GridLayout_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_GridLayout)
    GridLayout_construct w
    Set GridLayout_new = w
End Function

Public Sub GridLayout_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_GridLayout ' (VB) virtual calls from here on reach GridLayout's overrides
    GridLayout_(self).layoutDirty = False
End Sub

Public Sub GridLayout_place(ByVal self As WIDGET, column As grid_allocation_slot, row As grid_allocation_slot, ByVal widget As WIDGET)
    Dim p As GridLayout_Placement
    self.attach widget
    p = GridLayout_Placement_(column, row, widget)
    GridLayout_(self).placements_push_back p
    GridLayout_invalidateAllocation self
End Sub

Public Function GridLayout_idealWidth(ByVal self As WIDGET) As Long
    GridLayout_idealWidth = GridLayout_getColumnAllocator(self).get_minimum_size()
End Function

Public Function GridLayout_idealHeight(ByVal self As WIDGET) As Long
    GridLayout_idealHeight = GridLayout_getRowAllocator(self).get_minimum_size()
End Function

Public Function GridLayout_getScrollSnapOffsets(ByVal self As WIDGET, offsets() As Long) As Boolean
    GridLayout_updateLayout self
    GridLayout_(self).scrollSnapOffsets_get offsets
    GridLayout_getScrollSnapOffsets = True
End Function

Public Sub GridLayout_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim overrideContext As WidgetGraphicsContext
    GridLayout_updateLayout self
    overrideContext = WidgetGraphicsContext_setAllowChildDisplayRecursiveIfSelfClipped(context, True)
    WIDGET_displayRecursive self, overrideContext
End Sub

Public Sub GridLayout_geometryChanged(ByVal self As WIDGET)
    GridLayout_invalidateLayout self
End Sub

Private Sub GridLayout_updateLayout(ByVal self As WIDGET)
    Dim d As GridLayout, alloc As grid_allocation_allocator, columnOffsets As grid_offsets, rowOffsets As grid_offsets
    Dim i As Long, placement As GridLayout_Placement, left_ As Long, right_ As Long, top As Long, bottom As Long
    Set d = GridLayout_(self)
    If Not d.layoutDirty Then
        Exit Sub
    End If

    d.layoutDirty = False

    Set alloc = GridLayout_getColumnAllocator(self)
    alloc.calculate_offsets self.width(), columnOffsets
    Set alloc = GridLayout_getRowAllocator(self)
    alloc.calculate_offsets self.height(), rowOffsets
    d.columnOffsets = columnOffsets
    d.rowOffsets = rowOffsets
    d.scrollSnapOffsets_resize rowOffsets.size
    For i = 0 To rowOffsets.size - 1
        d.scrollSnapOffsets(i) = rowOffsets.values(i)
    Next

    For i = 0 To d.placementsCount - 1
        placement = d.placements(i)
        left_ = grid_offsets_get(columnOffsets, placement.column.start)
        right_ = grid_offsets_get(columnOffsets, placement.column.start + placement.column.span)
        top = grid_offsets_get(rowOffsets, placement.row.start)
        bottom = grid_offsets_get(rowOffsets, placement.row.start + placement.row.span)
        d.columnOffsets = columnOffsets         ' (VB) operator[] inserts the keys it does not find
        d.rowOffsets = rowOffsets
        placement.widget.setGeometry left_, top, right_ - left_, bottom - top
    Next
End Sub

Public Function GridLayout_getColumnOffsets(ByVal self As WIDGET) As grid_offsets
    GridLayout_updateLayout self
    GridLayout_getColumnOffsets = GridLayout_(self).columnOffsets
End Function

Public Function GridLayout_getRowOffsets(ByVal self As WIDGET) As grid_offsets
    GridLayout_updateLayout self
    GridLayout_getRowOffsets = GridLayout_(self).rowOffsets
End Function

Private Sub GridLayout_invalidateLayout(ByVal self As WIDGET)
    GridLayout_(self).layoutDirty = True
End Sub

Private Sub GridLayout_invalidateAllocation(ByVal self As WIDGET)
    Set GridLayout_(self).columnAllocation = Nothing
    Set GridLayout_(self).rowAllocation = Nothing
    GridLayout_invalidateLayout self
End Sub

Private Sub GridLayout_getAllocationItems(ByVal self As WIDGET, ByVal placementTransformerIsColumn As Boolean, items() As grid_allocation_item, n As Long)
    Dim d As GridLayout, i As Long, placement As GridLayout_Placement
    Set d = GridLayout_(self)
    n = 0
    For i = 0 To d.placementsCount - 1
        placement = d.placements(i)
        ReDim Preserve items(0 To n)
        If placementTransformerIsColumn Then
            items(n) = GridLayout_placementColumn(placement)
        Else
            items(n) = GridLayout_placementRow(placement)
        End If
        n = n + 1
    Next
End Sub

Private Function GridLayout_placementColumn(placement As GridLayout_Placement) As grid_allocation_item
    GridLayout_placementColumn = grid_allocation_item_(placement.column, placement.widget.idealWidth())
End Function

Private Function GridLayout_placementRow(placement As GridLayout_Placement) As grid_allocation_item
    GridLayout_placementRow = grid_allocation_item_(placement.row, placement.widget.idealHeight())
End Function

Private Function GridLayout_getColumnAllocator(ByVal self As WIDGET) As grid_allocation_allocator
    Dim d As GridLayout, items() As grid_allocation_item, n As Long
    Set d = GridLayout_(self)
    If d.columnAllocation Is Nothing Then
        GridLayout_getAllocationItems self, True, items, n
        Set d.columnAllocation = New grid_allocation_allocator
        d.columnAllocation.construct items, n
    End If

    Set GridLayout_getColumnAllocator = d.columnAllocation
End Function

Private Function GridLayout_getRowAllocator(ByVal self As WIDGET) As grid_allocation_allocator
    Dim d As GridLayout, items() As grid_allocation_item, n As Long
    Set d = GridLayout_(self)
    If d.rowAllocation Is Nothing Then
        GridLayout_getAllocationItems self, False, items, n
        Set d.rowAllocation = New grid_allocation_allocator
        d.rowAllocation.construct items, n
    End If

    Set GridLayout_getRowAllocator = d.rowAllocation
End Function
