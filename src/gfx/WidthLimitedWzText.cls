VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WidthLimitedWzText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/textdraw.h, textdraw.cpp: class WidthLimitedWzText: public
' WzText. The base is contained (base) and its methods forwarded; asText gives
' the WzText itself. A font is an iV_fonts as a Long (see WzText.cls).
Option Explicit

Private base As WzText
Private mFullText As String
Private mLimitWidthPts As Long                  ' size_t = 0

Private Sub Class_Initialize()
    Set base = New WzText
End Sub

' (VB) the WzText base
Public Function asText() As WzText
    Set asText = base
End Function

' Sets the text, truncating to a desired width limit (in *points*) if needed
' returns: the length of the string that will be drawn (may be less than the input text.length() if truncated)
Public Function setTruncatableText(ByVal text As String, ByVal fontID As Long, ByVal limitWidthInPoints As Long) As Long
    Dim truncatedText As String
    If (mFullText = text) And (mLimitWidthPts = limitWidthInPoints) And (base.getFontID() = fontID) Then
        setTruncatableText = WzString_length(base.getText()) ' skip; no change
        Exit Function
    End If

    mFullText = text
    mLimitWidthPts = limitWidthInPoints

    truncatedText = text
    Do While WzString_length(truncatedText) > 0
        If Not ULT(limitWidthInPoints, iV_GetTextWidth(truncatedText, fontID)) Then Exit Do
        If Not WzString_pop_back(truncatedText) Then
            ASSERT_FAILURE "false", "lib/ivis_opengl/textdraw.cpp:2014", "setTruncatableText", "WzString::pop_back() failed??"
            Exit Do
        End If
    Loop

    base.setText truncatedText, fontID
    setTruncatableText = WzString_length(truncatedText)
End Function

' (WzText)
Public Function width() As Long
    width = base.width()
End Function

Public Function height() As Long
    height = base.height()
End Function

Public Function aboveBase() As Long
    aboveBase = base.aboveBase()
End Function

Public Function belowBase() As Long
    belowBase = base.belowBase()
End Function

Public Function lineSize() As Long
    lineSize = base.lineSize()
End Function

Public Sub setText(ByVal text As String, ByVal fontID As Long)
    base.setText text, fontID
End Sub

Public Function getText() As String
    getText = base.getText()
End Function

Public Function getFontID() As Long
    getFontID = base.getFontID()
End Function

Public Sub render(ByVal x As Single, ByVal y As Single, ByVal colour As Long, Optional ByVal rotation_ As Single = 0!, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1)
    base.render x, y, colour, rotation_, maxWidth, maxHeight
End Sub

Public Sub renderOutlined(ByVal x As Long, ByVal y As Long, ByVal colour As Long, ByVal outlineColour As Long)
    base.renderOutlined x, y, colour, outlineColour
End Sub

Public Sub renderClipped(ByVal x As Single, ByVal y As Single, ByVal colour As Long, ByVal clipLeft As Long, ByVal clipTop As Long, ByVal clipWidth As Long, ByVal clipHeight As Long, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1)
    base.renderClipped x, y, colour, clipLeft, clipTop, clipWidth, clipHeight, maxWidth, maxHeight
End Sub
