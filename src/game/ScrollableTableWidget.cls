VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScrollableTableWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ScrollableTableWidget
' lib/widget/table.h: ScrollableTableWidget's fields (table_c). rows is a Collection (element n is item
' n + 1); std::vector<glm::ivec4> lines is displayRecursive's local.
Option Explicit

Public weakSerial As Long
Private tableColumns_() As TableColumn
Public tableColumnsCount As Long
Private columnWidths_() As Long
Public columnWidthsCount As Long
Private minColumnWidths_() As Long
Public minColumnWidthsCount As Long
Private m_columnPadding As Vector2i
Public header As WIDGET                 ' std::shared_ptr<TableHeader>
Public scrollableList As WIDGET         ' std::shared_ptr<ScrollableListWidget>
Public rows As Collection               ' std::vector<std::shared_ptr<TableRow>>
Public drawColumnLines As Boolean       ' = false
Public userDidResizeColumnWidths As Boolean ' = false

Private Sub Class_Initialize()
    Set rows = New Collection
End Sub

Friend Property Get tableColumns(ByVal i As Long) As TableColumn
    tableColumns = tableColumns_(i)
End Property

Friend Property Let tableColumns(ByVal i As Long, v As TableColumn)
    tableColumns_(i) = v
End Property

Friend Sub tableColumns_push_back(v As TableColumn)
    ReDim Preserve tableColumns_(0 To tableColumnsCount)
    tableColumns_(tableColumnsCount) = v
    tableColumnsCount = tableColumnsCount + 1
End Sub

Friend Sub tableColumns_clear()
    Erase tableColumns_
    tableColumnsCount = 0
End Sub

' (std::vector::resize: new elements are blank)
Friend Sub tableColumns_resize(ByVal n As Long)
    If n <= 0 Then
        tableColumns_clear
        Exit Sub
    End If
    ReDim Preserve tableColumns_(0 To n - 1)
    tableColumnsCount = n
End Sub

Friend Sub tableColumns_assign(src() As TableColumn, ByVal n As Long)
    Dim i As Long
    tableColumns_resize n
    For i = 0 To n - 1
        tableColumns_(i) = src(i)
    Next
End Sub

' (the elements copied into dst; returns their number)
Friend Function tableColumns_copyTo(dst() As TableColumn) As Long
    Dim i As Long
    If tableColumnsCount = 0 Then
        Erase dst
        Exit Function
    End If
    ReDim dst(0 To tableColumnsCount - 1)
    For i = 0 To tableColumnsCount - 1
        dst(i) = tableColumns_(i)
    Next
    tableColumns_copyTo = tableColumnsCount
End Function

Friend Sub tableColumns_get(dst() As TableColumn)
    tableColumns_copyTo dst
End Sub

Friend Property Get columnWidth(ByVal i As Long) As Long
    columnWidth = columnWidths_(i)
End Property

Friend Property Let columnWidth(ByVal i As Long, ByVal v As Long)
    columnWidths_(i) = v
End Property

Friend Sub columnWidths_push_back(ByVal v As Long)
    ReDim Preserve columnWidths_(0 To columnWidthsCount)
    columnWidths_(columnWidthsCount) = v
    columnWidthsCount = columnWidthsCount + 1
End Sub

Friend Sub columnWidths_clear()
    Erase columnWidths_
    columnWidthsCount = 0
End Sub

' (std::vector::resize: new elements are blank)
Friend Sub columnWidths_resize(ByVal n As Long)
    If n <= 0 Then
        columnWidths_clear
        Exit Sub
    End If
    ReDim Preserve columnWidths_(0 To n - 1)
    columnWidthsCount = n
End Sub

Friend Sub columnWidths_assign(src() As Long, ByVal n As Long)
    Dim i As Long
    columnWidths_resize n
    For i = 0 To n - 1
        columnWidths_(i) = src(i)
    Next
End Sub

' (the elements copied into dst; returns their number)
Friend Function columnWidths_copyTo(dst() As Long) As Long
    Dim i As Long
    If columnWidthsCount = 0 Then
        Erase dst
        Exit Function
    End If
    ReDim dst(0 To columnWidthsCount - 1)
    For i = 0 To columnWidthsCount - 1
        dst(i) = columnWidths_(i)
    Next
    columnWidths_copyTo = columnWidthsCount
End Function

Friend Sub columnWidths_get(dst() As Long)
    columnWidths_copyTo dst
End Sub

Friend Property Get minColumnWidth(ByVal i As Long) As Long
    minColumnWidth = minColumnWidths_(i)
End Property

Friend Property Let minColumnWidth(ByVal i As Long, ByVal v As Long)
    minColumnWidths_(i) = v
End Property

Friend Sub minColumnWidths_push_back(ByVal v As Long)
    ReDim Preserve minColumnWidths_(0 To minColumnWidthsCount)
    minColumnWidths_(minColumnWidthsCount) = v
    minColumnWidthsCount = minColumnWidthsCount + 1
End Sub

Friend Sub minColumnWidths_clear()
    Erase minColumnWidths_
    minColumnWidthsCount = 0
End Sub

' (std::vector::resize: new elements are blank)
Friend Sub minColumnWidths_resize(ByVal n As Long)
    If n <= 0 Then
        minColumnWidths_clear
        Exit Sub
    End If
    ReDim Preserve minColumnWidths_(0 To n - 1)
    minColumnWidthsCount = n
End Sub

Friend Sub minColumnWidths_assign(src() As Long, ByVal n As Long)
    Dim i As Long
    minColumnWidths_resize n
    For i = 0 To n - 1
        minColumnWidths_(i) = src(i)
    Next
End Sub

' (the elements copied into dst; returns their number)
Friend Function minColumnWidths_copyTo(dst() As Long) As Long
    Dim i As Long
    If minColumnWidthsCount = 0 Then
        Erase dst
        Exit Function
    End If
    ReDim dst(0 To minColumnWidthsCount - 1)
    For i = 0 To minColumnWidthsCount - 1
        dst(i) = minColumnWidths_(i)
    Next
    minColumnWidths_copyTo = minColumnWidthsCount
End Function

Friend Sub minColumnWidths_get(dst() As Long)
    minColumnWidths_copyTo dst
End Sub

Friend Property Get columnPadding() As Vector2i
    columnPadding = m_columnPadding
End Property

Friend Property Let columnPadding(v As Vector2i)
    m_columnPadding = v
End Property
