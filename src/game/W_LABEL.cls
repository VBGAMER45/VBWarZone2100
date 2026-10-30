VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_LABEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_LABEL
' lib/widget/label.h: W_LABEL's fields (label_c); the initialisers are W_LABEL_construct's.
' aTextLines (std::vector<TextLine>: text, dimensions, offset) is aTextLinesCount lines read with
' lineText(i), lineDimX(i), lineDimY(i), lineOffX(i), lineOffY(i); displayCache (LabelDisplayCache:
' std::vector<WzCachedText> wzText) is wzTextCount objects read with wzTextAt(i). 0-based.
Option Explicit

Public aTextLinesCount As Long          ' text lines on the label
Private lineText_() As String
Private lineDimX_() As Long
Private lineDimY_() As Long
Private lineOffX_() As Long
Private lineOffY_() As Long
Public maxLineWidth As Long             ' = -1
Public FontID As iV_fonts
Public pTip As String                   ' std::string: The tool tip for the button
Public fontColour As Long               ' PIELIGHT
Private wzText_() As WzCachedText       ' LabelDisplayCache displayCache
Public wzTextCount As Long
Public lineSpacing As Long              ' = 0
Public cacheNeverExpires As Boolean     ' = false
Public canTruncate As Boolean           ' = false
Public isTruncated As Boolean           ' = false

' aTextLines = iV_FormatText(...) (the GfxApi arrays, n lines)
Public Sub setTextLines(ByVal n As Long, txt() As String, dimX() As Long, dimY() As Long, offX() As Long, offY() As Long)
    Dim i As Long
    clearTextLines
    For i = 0 To n - 1
        pushTextLine txt(i), dimX(i), dimY(i), offX(i), offY(i)
    Next
End Sub

' aTextLines.clear()
Public Sub clearTextLines()
    Erase lineText_
    Erase lineDimX_
    Erase lineDimY_
    Erase lineOffX_
    Erase lineOffY_
    aTextLinesCount = 0
End Sub

' aTextLines.push_back({text, Vector2i(dimX, dimY), Vector2i(offX, offY)})
Public Sub pushTextLine(ByVal text As String, ByVal dimX As Long, ByVal dimY As Long, ByVal offX As Long, ByVal offY As Long)
    If aTextLinesCount = 0 Then
        ReDim lineText_(0 To 3)
        ReDim lineDimX_(0 To 3)
        ReDim lineDimY_(0 To 3)
        ReDim lineOffX_(0 To 3)
        ReDim lineOffY_(0 To 3)
    ElseIf aTextLinesCount > UBound(lineText_) Then
        ReDim Preserve lineText_(0 To 2 * aTextLinesCount - 1)
        ReDim Preserve lineDimX_(0 To 2 * aTextLinesCount - 1)
        ReDim Preserve lineDimY_(0 To 2 * aTextLinesCount - 1)
        ReDim Preserve lineOffX_(0 To 2 * aTextLinesCount - 1)
        ReDim Preserve lineOffY_(0 To 2 * aTextLinesCount - 1)
    End If
    lineText_(aTextLinesCount) = text
    lineDimX_(aTextLinesCount) = dimX
    lineDimY_(aTextLinesCount) = dimY
    lineOffX_(aTextLinesCount) = offX
    lineOffY_(aTextLinesCount) = offY
    aTextLinesCount = aTextLinesCount + 1
End Sub

Public Function lineText(ByVal i As Long) As String
    lineText = lineText_(i)
End Function

Public Function lineDimX(ByVal i As Long) As Long
    lineDimX = lineDimX_(i)
End Function

Public Function lineDimY(ByVal i As Long) As Long
    lineDimY = lineDimY_(i)
End Function

Public Function lineOffX(ByVal i As Long) As Long
    lineOffX = lineOffX_(i)
End Function

Public Function lineOffY(ByVal i As Long) As Long
    lineOffY = lineOffY_(i)
End Function

Public Function wzTextAt(ByVal i As Long) As WzCachedText
    Set wzTextAt = wzText_(i)
End Function

' displayCache.wzText.clear()
Public Sub wzTextClear()
    Erase wzText_
    wzTextCount = 0
End Sub

' displayCache.wzText.push_back(t)
Public Sub wzTextPush(ByVal t As WzCachedText)
    If wzTextCount = 0 Then
        ReDim wzText_(0 To 3)
    ElseIf wzTextCount > UBound(wzText_) Then
        ReDim Preserve wzText_(0 To 2 * wzTextCount - 1)
    End If
    Set wzText_(wzTextCount) = t
    wzTextCount = wzTextCount + 1
End Sub
