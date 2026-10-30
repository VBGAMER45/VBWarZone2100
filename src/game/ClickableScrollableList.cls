VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClickableScrollableList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ClickableScrollableList
' lib/widget/scrollablelist.h: ClickableScrollableList's fields (scrollablelist_c).
Option Explicit

Public onClickFunc As WzFunc            ' ClickableScrollableList_OnClick_Func
Public onHighlightFunc As WzFunc        ' ClickableScrollableList_OnHighlight_Func
Public mouseDownOnList As Boolean       ' = false
