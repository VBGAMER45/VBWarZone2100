VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScrollableListWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ScrollableListWidget
' lib/widget/scrollablelist.h: ScrollableListWidget's fields (scrollablelist_c). A PIELIGHT is a Long.
' Padding padding is padding_.
Option Explicit

Public scrollBar As WIDGET              ' std::shared_ptr<ScrollBarWidget>
Public listView As WIDGET               ' std::shared_ptr<ClipRectWidget>
Public scrollableHeight As Long         ' uint32_t = 0
Public snapOffset As Boolean            ' = true
Public layoutDirty As Boolean           ' = false
Private m_padding_ As Padding   ' Padding padding = {0, 0, 0, 0}
Public backgroundColor As Long
Public borderColor As Long
Public itemSpacing As Long              ' uint32_t = 0
Public scrollbarWidth As Long           ' = 0
Public expandWidthWhenScrollbarInvisible As Boolean ' = true
Public drawRowLines As Boolean          ' = false
Public rowLinesColor As Long
Public topVisibleItemIdx As Long        ' size_t = 0

Friend Property Get padding_() As Padding
    padding_ = m_padding_
End Property

Friend Property Let padding_(v As Padding)
    m_padding_ = v
End Property
