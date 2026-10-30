VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "grid_allocation_allocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/widget/gridallocation.cpp: class grid_allocation::allocator.
' allocator(items) is New + construct(items, count); calculate_offsets returns
' its std::map in a grid_offsets (gridallocation_c). steps_index (the index of a
' step in the sorted steps) is stepIndex. std::sort's comparison orders items by
' (span, start), which normalize_items makes unique, so any sort gives C++'s order.
Option Explicit

Private slots_size As Long              ' size_t
Private steps() As Long                 ' std::vector<uint32_t> (ascending, unique)
Private steps_size As Long
Private base_sizes() As Long            ' std::vector<int32_t>
Private expandable_map() As Boolean     ' std::vector<bool>

Friend Sub construct(items() As grid_allocation_item, ByVal itemsCount As Long)
    Dim normalized_items() As grid_allocation_item, n As Long
    If itemsCount <> 0 Then
        initialize_steps items, itemsCount
        normalize_items items, itemsCount, normalized_items, n
        initialize_expandable_map normalized_items, n
        initialize_base_sizes normalized_items, n
    End If
End Sub

' std::map<uint32_t, int32_t> calculate_offsets(int32_t available_space)
Friend Sub calculate_offsets(ByVal space As Long, offsets As grid_offsets)
    Dim sizes() As Long, i As Long, current_offset As Long
    grid_offsets_clear offsets
    If steps_size = 0 Then
        Exit Sub
    End If

    sizes = base_sizes
    expand sizes, space, 0, slots_size
    shrink sizes, space, 0, slots_size

    grid_offsets_set offsets, 0, 0
    current_offset = 0
    For i = 0 To slots_size - 1
        current_offset = current_offset + sizes(i)
        grid_offsets_set offsets, steps(i + 1), current_offset
    Next
End Sub

Private Sub expand(sizes() As Long, ByVal space As Long, ByVal start As Long, ByVal end_ As Long)
    Dim current_size As Long, space_to_expand As Long, slots_to_expand As Long, i As Long
    Dim has_expandable As Boolean, span As Long, increment As Long
    current_size = 0
    For i = start To end_ - 1
        current_size = current_size + sizes(i)
    Next
    If current_size >= space Then
        Exit Sub
    End If

    space_to_expand = space - current_size

    slots_to_expand = 0
    For i = start To end_ - 1
        If expandable_map(i) Then
            slots_to_expand = slots_to_expand + steps(i + 1) - steps(i)
        End If
    Next

    has_expandable = slots_to_expand > 0
    If Not has_expandable Then
        slots_to_expand = steps(end_) - steps(start)
    End If

    For i = start To end_ - 1
        If Not has_expandable Or expandable_map(i) Then
            span = steps(i + 1) - steps(i)
            increment = space_to_expand * span \ slots_to_expand
            sizes(i) = sizes(i) + increment
            space_to_expand = space_to_expand - increment
            slots_to_expand = slots_to_expand - span
        End If
    Next
End Sub

Private Sub shrink(sizes() As Long, ByVal space As Long, ByVal start As Long, ByVal end_ As Long)
    Dim current_size As Long, space_to_reduce As Long, i As Long, decrement As Long
    current_size = 0
    For i = start To end_ - 1
        current_size = current_size + sizes(i)
    Next
    If current_size <= space Then
        Exit Sub
    End If

    space_to_reduce = current_size - space
    i = start
    Do While i < end_ And current_size > 0
        decrement = space_to_reduce * sizes(i) \ current_size
        current_size = current_size - sizes(i)
        sizes(i) = sizes(i) - decrement
        space_to_reduce = space_to_reduce - decrement
        i = i + 1
    Loop
End Sub

Public Function get_minimum_size() As Long
    Dim i As Long
    For i = 0 To slots_size - 1
        get_minimum_size = get_minimum_size + base_sizes(i)
    Next
End Function

' (VB) steps_index[value]: the index of value in steps
Private Function stepIndex(ByVal value As Long) As Long
    Dim i As Long
    For i = 0 To steps_size - 1
        If steps(i) = value Then
            stepIndex = i
            Exit Function
        End If
    Next
End Function

' (VB) std::set<uint32_t> insert, steps kept ascending and unique
Private Sub insertStep(ByVal value As Long)
    Dim i As Long, j As Long
    For i = 0 To steps_size - 1
        If steps(i) = value Then Exit Sub
        If ULT(value, steps(i)) Then Exit For
    Next
    ReDim Preserve steps(0 To steps_size)
    For j = steps_size To i + 1 Step -1
        steps(j) = steps(j - 1)
    Next
    steps(i) = value
    steps_size = steps_size + 1
End Sub

Private Sub initialize_steps(items() As grid_allocation_item, ByVal itemsCount As Long)
    Dim i As Long
    steps_size = 0
    Erase steps
    For i = 0 To itemsCount - 1
        insertStep grid_allocation_item_start(items(i))
        insertStep grid_allocation_item_end(items(i))
    Next

    slots_size = steps_size - 1
End Sub

Private Sub normalize_items(items() As grid_allocation_item, ByVal itemsCount As Long, normalized_items() As grid_allocation_item, n As Long)
    Dim i As Long, j As Long, it As grid_allocation_item, existing As Long, tmp As grid_allocation_item
    n = 0
    For i = 0 To itemsCount - 1
        it = items(i)
        it.slot.span = stepIndex(grid_allocation_item_end(it)) - stepIndex(grid_allocation_item_start(it))
        it.slot.start = stepIndex(grid_allocation_item_start(it))

        ' items_map_key key{item.start(), item.span()}
        existing = -1
        For j = 0 To n - 1
            If normalized_items(j).slot.start = it.slot.start And normalized_items(j).slot.span = it.slot.span Then
                existing = j
                Exit For
            End If
        Next
        If existing = -1 Then
            ReDim Preserve normalized_items(0 To n)
            normalized_items(n) = it
            n = n + 1
        Else
            normalized_items(existing).slot.expandable = normalized_items(existing).slot.expandable And it.slot.expandable
            If it.ideal_size > normalized_items(existing).ideal_size Then normalized_items(existing).ideal_size = it.ideal_size
        End If
    Next

    ' std::sort by (span, start)
    For i = 1 To n - 1
        tmp = normalized_items(i)
        j = i - 1
        Do While j >= 0
            If Not itemLess(tmp, normalized_items(j)) Then Exit Do
            normalized_items(j + 1) = normalized_items(j)
            j = j - 1
        Loop
        normalized_items(j + 1) = tmp
    Next
End Sub

Private Function itemLess(a As grid_allocation_item, b As grid_allocation_item) As Boolean
    If a.slot.span = b.slot.span Then
        itemLess = ULT(a.slot.start, b.slot.start)
        Exit Function
    End If

    itemLess = ULT(a.slot.span, b.slot.span)
End Function

Private Sub initialize_expandable_map(items() As grid_allocation_item, ByVal n As Long)
    Dim i As Long, slot As Long
    If slots_size > 0 Then
        ReDim expandable_map(0 To slots_size - 1)
        For i = 0 To slots_size - 1
            expandable_map(i) = True
        Next
    End If
    For i = 0 To n - 1
        If Not grid_allocation_item_expandable(items(i)) Then
            For slot = grid_allocation_item_start(items(i)) To grid_allocation_item_end(items(i)) - 1
                expandable_map(slot) = False
            Next
        End If
    Next

    ' (C++'s second loop sets item.slot.expandable on a copy of each item: it changes nothing)
End Sub

Private Sub initialize_base_sizes(items() As grid_allocation_item, ByVal n As Long)
    ' std::vector<std::set<const item*>> slot_items(slots_size): slot_items(slot, item index)
    Dim slot_items() As Boolean, i As Long, slot As Long, borrow_slot As Long, k As Long
    Dim allocated As Long, allocatable As Long, unitSize As Long, same As Boolean
    If slots_size <= 0 Or n <= 0 Then Exit Sub
    ReDim slot_items(0 To slots_size - 1, 0 To n - 1)
    ReDim base_sizes(0 To slots_size - 1)
    For i = 0 To n - 1
        allocated = 0
        For slot = grid_allocation_item_start(items(i)) To grid_allocation_item_end(items(i)) - 1
            allocated = allocated + base_sizes(slot)
        Next
        If allocated >= items(i).ideal_size Then
            GoTo continue_
        End If

        For slot = grid_allocation_item_start(items(i)) To grid_allocation_item_end(items(i)) - 1
            slot_items(slot, i) = True
        Next

        slot = grid_allocation_item_start(items(i))
        Do While slot < grid_allocation_item_end(items(i)) And allocated < items(i).ideal_size
            ' auto& current_unit = slot_items[slot];
            unitSize = 0
            For k = 0 To n - 1
                If slot_items(slot, k) Then unitSize = unitSize + 1
            Next

            If unitSize = 1 Then
                allocatable = items(i).ideal_size - allocated
                allocated = allocated + allocatable
                base_sizes(slot) = base_sizes(slot) + allocatable
                Exit Do
            End If

            ' auto borrow_items = slot_items[slot]; borrow_items.erase(&item);
            For borrow_slot = 0 To grid_allocation_item_start(items(i)) - 1
                ' slot_items[borrow_slot] != borrow_items
                same = True
                For k = 0 To n - 1
                    If slot_items(borrow_slot, k) <> (slot_items(slot, k) And k <> i) Then
                        same = False
                        Exit For
                    End If
                Next
                If base_sizes(borrow_slot) = 0 Or Not same Then
                    GoTo continue_borrow
                End If

                allocatable = items(i).ideal_size - allocated
                If base_sizes(borrow_slot) < allocatable Then allocatable = base_sizes(borrow_slot)
                base_sizes(borrow_slot) = base_sizes(borrow_slot) - allocatable
                allocated = allocated + allocatable
                base_sizes(slot) = base_sizes(slot) + allocatable
continue_borrow:
            Next
            slot = slot + 1
        Loop

        expand base_sizes, items(i).ideal_size, grid_allocation_item_start(items(i)), grid_allocation_item_end(items(i))
continue_:
    Next
End Sub
