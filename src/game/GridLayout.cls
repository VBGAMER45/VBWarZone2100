VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class GridLayout
' lib/widget/gridlayout.h: GridLayout's fields (gridlayout_c). A nonstd::optional<grid_allocation::allocator>
' is Nothing when it has no value; a std::map<uint32_t, int32_t> is a grid_offsets.
Option Explicit

Public columnAllocation As grid_allocation_allocator
Public rowAllocation As grid_allocation_allocator
Private placements_() As GridLayout_Placement
Public placementsCount As Long
Public layoutDirty As Boolean           ' = false
Private scrollSnapOffsets_() As Long
Public scrollSnapOffsetsCount As Long
Private m_columnOffsets As grid_offsets
Private m_rowOffsets As grid_offsets

Friend Property Get placements(ByVal i As Long) As GridLayout_Placement
    placements = placements_(i)
End Property

Friend Property Let placements(ByVal i As Long, v As GridLayout_Placement)
    placements_(i) = v
End Property

Friend Sub placements_push_back(v As GridLayout_Placement)
    ReDim Preserve placements_(0 To placementsCount)
    placements_(placementsCount) = v
    placementsCount = placementsCount + 1
End Sub

Friend Sub placements_clear()
    Erase placements_
    placementsCount = 0
End Sub

' (std::vector::resize: new elements are blank)
Friend Sub placements_resize(ByVal n As Long)
    If n <= 0 Then
        placements_clear
        Exit Sub
    End If
    ReDim Preserve placements_(0 To n - 1)
    placementsCount = n
End Sub

Friend Sub placements_assign(src() As GridLayout_Placement, ByVal n As Long)
    Dim i As Long
    placements_resize n
    For i = 0 To n - 1
        placements_(i) = src(i)
    Next
End Sub

' (the elements copied into dst; returns their number)
Friend Function placements_copyTo(dst() As GridLayout_Placement) As Long
    Dim i As Long
    If placementsCount = 0 Then
        Erase dst
        Exit Function
    End If
    ReDim dst(0 To placementsCount - 1)
    For i = 0 To placementsCount - 1
        dst(i) = placements_(i)
    Next
    placements_copyTo = placementsCount
End Function

Friend Sub placements_get(dst() As GridLayout_Placement)
    placements_copyTo dst
End Sub

Friend Property Get scrollSnapOffsets(ByVal i As Long) As Long
    scrollSnapOffsets = scrollSnapOffsets_(i)
End Property

Friend Property Let scrollSnapOffsets(ByVal i As Long, ByVal v As Long)
    scrollSnapOffsets_(i) = v
End Property

Friend Sub scrollSnapOffsets_push_back(ByVal v As Long)
    ReDim Preserve scrollSnapOffsets_(0 To scrollSnapOffsetsCount)
    scrollSnapOffsets_(scrollSnapOffsetsCount) = v
    scrollSnapOffsetsCount = scrollSnapOffsetsCount + 1
End Sub

Friend Sub scrollSnapOffsets_clear()
    Erase scrollSnapOffsets_
    scrollSnapOffsetsCount = 0
End Sub

' (std::vector::resize: new elements are blank)
Friend Sub scrollSnapOffsets_resize(ByVal n As Long)
    If n <= 0 Then
        scrollSnapOffsets_clear
        Exit Sub
    End If
    ReDim Preserve scrollSnapOffsets_(0 To n - 1)
    scrollSnapOffsetsCount = n
End Sub

Friend Sub scrollSnapOffsets_assign(src() As Long, ByVal n As Long)
    Dim i As Long
    scrollSnapOffsets_resize n
    For i = 0 To n - 1
        scrollSnapOffsets_(i) = src(i)
    Next
End Sub

' (the elements copied into dst; returns their number)
Friend Function scrollSnapOffsets_copyTo(dst() As Long) As Long
    Dim i As Long
    If scrollSnapOffsetsCount = 0 Then
        Erase dst
        Exit Function
    End If
    ReDim dst(0 To scrollSnapOffsetsCount - 1)
    For i = 0 To scrollSnapOffsetsCount - 1
        dst(i) = scrollSnapOffsets_(i)
    Next
    scrollSnapOffsets_copyTo = scrollSnapOffsetsCount
End Function

Friend Sub scrollSnapOffsets_get(dst() As Long)
    scrollSnapOffsets_copyTo dst
End Sub

Friend Property Get columnOffsets() As grid_offsets
    columnOffsets = m_columnOffsets
End Property

Friend Property Let columnOffsets(v As grid_offsets)
    m_columnOffsets = v
End Property

Friend Property Get rowOffsets() As grid_offsets
    rowOffsets = m_rowOffsets
End Property

Friend Property Let rowOffsets(v As grid_offsets)
    m_rowOffsets = v
End Property
