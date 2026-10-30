VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TableRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class TableRow
' lib/widget/table.h: TableRow's fields (table_c). columnWidgets is a Collection (element n is item n + 1);
' an optional<T> is has_<name> and <name>.
Option Explicit

Public highlightsOnMouseOver As Boolean ' = false
Public disabledRow As Boolean           ' = false
Public columnWidgets As Collection
Public has_lastFrameMouseIsOverRowOrChildren As Boolean
Public lastFrameMouseIsOverRowOrChildren As Long ' UDWORD
Public has_borderColor As Boolean
Public borderColor As Long              ' PIELIGHT
Public has_backgroundColor As Boolean
Public backgroundColor As Long          ' PIELIGHT
Public highlightColor As Long           ' PIELIGHT
Public disabledColor As Long            ' PIELIGHT
Public maxDisplayedColumnX1 As Long     ' int32_t = 0

Private Sub Class_Initialize()
    Set columnWidgets = New Collection
End Sub
