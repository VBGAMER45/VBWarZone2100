VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScrollBarWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ScrollBarWidget
' lib/widget/scrollbar.h: ScrollBarWidget's fields (scrollbar_c).
Option Explicit

Public slider As WIDGET                 ' std::shared_ptr<W_SLIDER>
Public scrollableSize As Long           ' uint16_t = 1
Public viewSize As Long                 ' uint16_t = 1
Public stickToBottom As Boolean         ' = false
