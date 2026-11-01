VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResizeWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/widget/resize.h: the fields of class ResizeWidget (its Resize resize: has_ + value per optional)
' (VB) data of widget class ResizeWidget
Option Explicit

Public has_minWidth As Boolean
Public minWidth As Long
Public has_maxWidth As Boolean
Public maxWidth As Long
Public has_minHeight As Boolean
Public minHeight As Long
Public has_maxHeight As Boolean
Public maxHeight As Long
