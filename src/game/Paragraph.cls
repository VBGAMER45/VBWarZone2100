VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Paragraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class Paragraph
' lib/widget/paragraph.h: Paragraph's fields (paragraph_c). std::vector<std::unique_ptr<ParagraphElement>>
' elements is a Collection of ParagraphElement objects (element n is item n + 1).
Option Explicit

Public elements As Collection
Public layoutDirty As Boolean           ' = true
Public layoutWidth As Long              ' uint32_t = 0
Private m_textStyle As ParagraphTextStyle
Public lineSpacing As Long              ' uint32_t = 0
Private scrollSnapOffsets_() As Long
Public scrollSnapOffsetsCount As Long
Public isMouseDown As Boolean           ' = false
Public onClickHandlers As Collection    ' std::vector<W_PARAGRAPH_ONCLICK_FUNC>

Private Sub Class_Initialize()
    Set elements = New Collection
    Set onClickHandlers = New Collection
End Sub

Friend Property Get textStyle() As ParagraphTextStyle
    textStyle = m_textStyle
End Property

Friend Property Let textStyle(v As ParagraphTextStyle)
    m_textStyle = v
End Property

Friend Property Get scrollSnapOffsets(ByVal i As Long) As Long
    scrollSnapOffsets = scrollSnapOffsets_(i)
End Property

Friend Property Let scrollSnapOffsets(ByVal i As Long, ByVal v As Long)
    scrollSnapOffsets_(i) = v
End Property

Friend Sub scrollSnapOffsets_push_back(ByVal v As Long)
    ReDim Preserve scrollSnapOffsets_(0 To scrollSnapOffsetsCount)
    scrollSnapOffsets_(scrollSnapOffsetsCount) = v
    scrollSnapOffsetsCount = scrollSnapOffsetsCount + 1
End Sub

Friend Sub scrollSnapOffsets_clear()
    Erase scrollSnapOffsets_
    scrollSnapOffsetsCount = 0
End Sub

' (std::vector::resize: new elements are blank)
Friend Sub scrollSnapOffsets_resize(ByVal n As Long)
    If n <= 0 Then
        scrollSnapOffsets_clear
        Exit Sub
    End If
    ReDim Preserve scrollSnapOffsets_(0 To n - 1)
    scrollSnapOffsetsCount = n
End Sub

Friend Sub scrollSnapOffsets_assign(src() As Long, ByVal n As Long)
    Dim i As Long
    scrollSnapOffsets_resize n
    For i = 0 To n - 1
        scrollSnapOffsets_(i) = src(i)
    Next
End Sub

' (the elements copied into dst; returns their number)
Friend Function scrollSnapOffsets_copyTo(dst() As Long) As Long
    Dim i As Long
    If scrollSnapOffsetsCount = 0 Then
        Erase dst
        Exit Function
    End If
    ReDim dst(0 To scrollSnapOffsetsCount - 1)
    For i = 0 To scrollSnapOffsetsCount - 1
        dst(i) = scrollSnapOffsets_(i)
    Next
    scrollSnapOffsets_copyTo = scrollSnapOffsetsCount
End Function

Friend Sub scrollSnapOffsets_get(dst() As Long)
    scrollSnapOffsets_copyTo dst
End Sub
