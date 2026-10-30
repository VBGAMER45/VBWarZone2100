VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ListWidget
' lib/widget/listwidget.h: ListWidget's fields (listwidget_c). myChildren and the handlers are Collections
' (element n is item n + 1).
Option Explicit

Public numberOfPages As Long            ' size_t = 1
Private m_childSize As WzSize
Private m_spacing As WzSize
Public currentPage_ As Long             ' size_t
Public myChildren As Collection
Public order As ListWidget_Order
Public onCurrentPageChangedHandlers As Collection
Public onNumberOfPagesChangedHandlers As Collection

Private Sub Class_Initialize()
    Set myChildren = New Collection
    Set onCurrentPageChangedHandlers = New Collection
    Set onNumberOfPagesChangedHandlers = New Collection
End Sub

Friend Property Get childSize() As WzSize
    childSize = m_childSize
End Property

Friend Property Let childSize(v As WzSize)
    m_childSize = v
End Property

Friend Property Get spacing() As WzSize
    spacing = m_spacing
End Property

Friend Property Let spacing(v As WzSize)
    m_spacing = v
End Property
