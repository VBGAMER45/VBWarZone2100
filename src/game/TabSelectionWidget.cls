VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabSelectionWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class TabSelectionWidget
' lib/widget/listwidget.h: TabSelectionWidget's fields (listwidget_c). std::vector<std::shared_ptr<W_BUTTON>>
' tabButtons and the handlers are Collections (element n is item n + 1).
Option Explicit

Private styles_() As TabSelectionStyle
Public stylesCount As Long
Public currentTab As Long               ' size_t
Public tabsAtOnce As Long               ' size_t
Public tabButtons As Collection
Public prevTabPageButton As WIDGET
Public nextTabPageButton As WIDGET
Public onTabChangedHandlers As Collection

Private Sub Class_Initialize()
    Set tabButtons = New Collection
    Set onTabChangedHandlers = New Collection
End Sub

Friend Property Get styles(ByVal i As Long) As TabSelectionStyle
    styles = styles_(i)
End Property

Friend Property Let styles(ByVal i As Long, v As TabSelectionStyle)
    styles_(i) = v
End Property

Friend Sub styles_push_back(v As TabSelectionStyle)
    ReDim Preserve styles_(0 To stylesCount)
    styles_(stylesCount) = v
    stylesCount = stylesCount + 1
End Sub

Friend Sub styles_clear()
    Erase styles_
    stylesCount = 0
End Sub

' (std::vector::resize: new elements are blank)
Friend Sub styles_resize(ByVal n As Long)
    If n <= 0 Then
        styles_clear
        Exit Sub
    End If
    ReDim Preserve styles_(0 To n - 1)
    stylesCount = n
End Sub

Friend Sub styles_assign(src() As TabSelectionStyle, ByVal n As Long)
    Dim i As Long
    styles_resize n
    For i = 0 To n - 1
        styles_(i) = src(i)
    Next
End Sub

' (the elements copied into dst; returns their number)
Friend Function styles_copyTo(dst() As TabSelectionStyle) As Long
    Dim i As Long
    If stylesCount = 0 Then
        Erase dst
        Exit Function
    End If
    ReDim dst(0 To stylesCount - 1)
    For i = 0 To stylesCount - 1
        dst(i) = styles_(i)
    Next
    styles_copyTo = stylesCount
End Function

Friend Sub styles_get(dst() As TabSelectionStyle)
    styles_copyTo dst
End Sub
