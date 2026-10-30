VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TableHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class TableHeader
' lib/widget/table.cpp: TableHeader's fields (table_c). An optional<T> is has_<name> and <name>;
' std::weak_ptr<ScrollableTableWidget> parentTable is the table's serial (widgweak_vb).
Option Explicit

Public userResizableHeaders As Boolean  ' = true
Private m_columnPadding As Vector2i
Private columns_() As TableColumn
Public columnsCount As Long
Public has_colBeingResized As Boolean
Public colBeingResized As Long          ' size_t
Public has_dragStart As Boolean
Private m_dragStart As Vector2i
Public parentTable As Long

Friend Property Get columnPadding() As Vector2i
    columnPadding = m_columnPadding
End Property

Friend Property Let columnPadding(v As Vector2i)
    m_columnPadding = v
End Property

Friend Property Get columns(ByVal i As Long) As TableColumn
    columns = columns_(i)
End Property

Friend Property Let columns(ByVal i As Long, v As TableColumn)
    columns_(i) = v
End Property

Friend Sub columns_push_back(v As TableColumn)
    ReDim Preserve columns_(0 To columnsCount)
    columns_(columnsCount) = v
    columnsCount = columnsCount + 1
End Sub

Friend Sub columns_clear()
    Erase columns_
    columnsCount = 0
End Sub

' (std::vector::resize: new elements are blank)
Friend Sub columns_resize(ByVal n As Long)
    If n <= 0 Then
        columns_clear
        Exit Sub
    End If
    ReDim Preserve columns_(0 To n - 1)
    columnsCount = n
End Sub

Friend Sub columns_assign(src() As TableColumn, ByVal n As Long)
    Dim i As Long
    columns_resize n
    For i = 0 To n - 1
        columns_(i) = src(i)
    Next
End Sub

' (the elements copied into dst; returns their number)
Friend Function columns_copyTo(dst() As TableColumn) As Long
    Dim i As Long
    If columnsCount = 0 Then
        Erase dst
        Exit Function
    End If
    ReDim dst(0 To columnsCount - 1)
    For i = 0 To columnsCount - 1
        dst(i) = columns_(i)
    Next
    columns_copyTo = columnsCount
End Function

Friend Sub columns_get(dst() As TableColumn)
    columns_copyTo dst
End Sub

Friend Property Get dragStart() As Vector2i
    dragStart = m_dragStart
End Property

Friend Property Let dragStart(v As Vector2i)
    m_dragStart = v
End Property
