VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarginWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/widget/margin.h: the fields of class MarginWidget
' (VB) data of widget class MarginWidget
Option Explicit

' Margin margin; (struct Margin's top, right, bottom, left)
Public margin_top As Long
Public margin_right As Long
Public margin_bottom As Long
Public margin_left As Long
