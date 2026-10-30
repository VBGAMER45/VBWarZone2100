Attribute VB_Name = "gridallocation_c"
' lib/widget/gridallocation.cpp, gridallocation.h: namespace grid_allocation's
' structs slot and item (UDTs grid_allocation_slot, grid_allocation_item; their
' methods grid_allocation_item_<name>) and the std::map<uint32_t, int32_t> of
' offsets the allocator returns (grid_offsets: keys ascending, as std::map
' iterates). The class grid_allocation::allocator is grid_allocation_allocator.cls.
Option Explicit

Public Type grid_allocation_slot
    start As Long                       ' uint32_t
    span As Long                        ' uint32_t
    expandable As Boolean
End Type

Public Type grid_allocation_item
    slot As grid_allocation_slot
    ideal_size As Long                  ' int32_t
End Type

' (VB) std::map<uint32_t, int32_t>
Public Type grid_offsets
    keys() As Long
    values() As Long
    size As Long
End Type

' slot(uint32_t start = 0, uint32_t span = 1, bool expandable = true)
Public Function grid_allocation_slot_(Optional ByVal start As Long = 0, Optional ByVal span As Long = 1, Optional ByVal expandable As Boolean = True) As grid_allocation_slot
    grid_allocation_slot_.start = start
    grid_allocation_slot_.span = span
    grid_allocation_slot_.expandable = expandable
End Function

' item{slot, ideal_size}
Public Function grid_allocation_item_(s As grid_allocation_slot, ByVal ideal_size As Long) As grid_allocation_item
    grid_allocation_item_.slot = s
    grid_allocation_item_.ideal_size = ideal_size
End Function

Public Function grid_allocation_item_start(it As grid_allocation_item) As Long
    grid_allocation_item_start = it.slot.start
End Function

Public Function grid_allocation_item_end(it As grid_allocation_item) As Long
    grid_allocation_item_end = it.slot.start + it.slot.span
End Function

Public Function grid_allocation_item_span(it As grid_allocation_item) As Long
    grid_allocation_item_span = it.slot.span
End Function

Public Function grid_allocation_item_expandable(it As grid_allocation_item) As Boolean
    grid_allocation_item_expandable = it.slot.expandable
End Function

'--------------------------------------------------------------------------
' (VB) std::map<uint32_t, int32_t> operations

Public Sub grid_offsets_clear(m As grid_offsets)
    Erase m.keys
    Erase m.values
    m.size = 0
End Sub

' m[key] = value
Public Sub grid_offsets_set(m As grid_offsets, ByVal key As Long, ByVal value As Long)
    Dim i As Long, j As Long
    For i = 0 To m.size - 1
        If m.keys(i) = key Then
            m.values(i) = value
            Exit Sub
        End If
        If ULT(key, m.keys(i)) Then Exit For
    Next
    ReDim Preserve m.keys(0 To m.size)
    ReDim Preserve m.values(0 To m.size)
    For j = m.size To i + 1 Step -1
        m.keys(j) = m.keys(j - 1)
        m.values(j) = m.values(j - 1)
    Next
    m.keys(i) = key
    m.values(i) = value
    m.size = m.size + 1
End Sub

' m[key] (operator[]: a missing key is inserted with 0, as std::map does)
Public Function grid_offsets_get(m As grid_offsets, ByVal key As Long) As Long
    Dim i As Long
    For i = 0 To m.size - 1
        If m.keys(i) = key Then
            grid_offsets_get = m.values(i)
            Exit Function
        End If
    Next
    grid_offsets_set m, key, 0
    grid_offsets_get = 0
End Function
