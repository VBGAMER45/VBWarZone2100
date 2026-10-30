Attribute VB_Name = "label_c"
' lib/widget/label.cpp, label.h: the label widget W_LABEL (its fields:
' W_LABEL.cls) and W_LABINIT's constructor. The display cache's WzCachedText is
' paragraph.h's (WzCachedText.cls, paragraph_h).
' (VB) widget class W_LABEL : WIDGET
Option Explicit

Private Const LABEL_DEFAULT_CACHE_EXPIRY As Long = 250

' W_LABINIT::W_LABINIT(): FontID(font_regular) (a VB UDT starts so: for uniformity)
Public Sub W_LABINIT_construct(i As W_LABINIT)
    Dim z As W_LABINIT
    i = z
    i.FontID = font_regular
End Sub

' (VB) the W_INIT part of a W_LABINIT (W_LABINIT : W_INIT)
Public Function W_LABINIT_toW_INIT(init As W_LABINIT) As W_INIT
    Dim r As W_INIT
    r.formID = init.formID
    r.majorID = init.majorID
    r.id = init.id
    r.style = init.style
    r.x = init.x
    r.y = init.y
    r.width = init.width
    r.height = init.height
    Set r.pDisplay = init.pDisplay
    Set r.pCallback = init.pCallback
    If IsObject(init.pUserData) Then Set r.pUserData = init.pUserData Else r.pUserData = init.pUserData
    r.UserData = init.UserData
    Set r.calcLayout = init.calcLayout
    Set r.onDelete = init.onDelete
    Set r.customHitTest = init.customHitTest
    Set r.initPUserDataFunc = init.initPUserDataFunc
    W_LABINIT_toW_INIT = r
End Function

' W_LABEL::W_LABEL(W_LABINIT const *init)
Public Sub W_LABEL_construct_init(ByVal self As WIDGET, init As W_LABINIT)
    Dim baseInit As W_INIT, d As W_LABEL
    baseInit = W_LABINIT_toW_INIT(init)
    WIDGET_construct_init self, baseInit, WIDG_LABEL
    self.kind = K_W_LABEL ' (VB) virtual calls from here on reach W_LABEL's overrides
    W_LABEL_initMembers self
    Set d = W_LABEL_(self)
    d.FontID = init.FontID
    d.pTip = init.pTip
    d.fontColour = psPalette(WZCOL_FORM_TEXT)
    If Not ((init.style And Not (WLAB_PLAIN Or WLAB_ALIGNLEFT Or WLAB_ALIGNRIGHT Or WLAB_ALIGNCENTRE Or WLAB_ALIGNTOP Or WLAB_ALIGNBOTTOM Or WIDG_HIDDEN)) = 0) Then ASSERT_FAILURE "(init->style & ~(0 | WLAB_ALIGNLEFT | WLAB_ALIGNRIGHT | WLAB_ALIGNCENTRE | WLAB_ALIGNTOP | WLAB_ALIGNBOTTOM | WIDG_HIDDEN)) == 0", "lib/widget/label.cpp:46", "W_LABEL", "Unknown button style"
    W_LABEL_setString self, init.pText
End Sub

' W_LABEL::W_LABEL()
Public Sub W_LABEL_construct(ByVal self As WIDGET)
    Dim d As W_LABEL
    WIDGET_construct self, WIDG_LABEL
    self.kind = K_W_LABEL ' (VB) virtual calls from here on reach W_LABEL's overrides
    W_LABEL_initMembers self
    Set d = W_LABEL_(self)
    d.FontID = font_regular
    d.fontColour = psPalette(WZCOL_FORM_TEXT)
End Sub

' the members' initialisers (label.h)
Private Sub W_LABEL_initMembers(ByVal self As WIDGET)
    Dim d As W_LABEL
    Set d = W_LABEL_(self)
    d.maxLineWidth = -1
    d.lineSpacing = 0
    d.cacheNeverExpires = False
    d.canTruncate = False
    d.isTruncated = False
End Sub

' std::make_shared<W_LABEL>(init)
Public Function W_LABEL_new_init(init As W_LABINIT) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_LABEL)
    W_LABEL_construct_init w, init
    Set W_LABEL_new_init = w
End Function

' std::make_shared<W_LABEL>()
Public Function W_LABEL_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_LABEL)
    W_LABEL_construct w
    Set W_LABEL_new = w
End Function

' Sets a string for the label
' - line-wraps at max width
' Returns the height of the formatted text
Public Function W_LABEL_setFormattedString(ByVal self As WIDGET, ByVal string_ As String, ByVal MaxWidth As Long, ByVal fontID As iV_fonts, Optional ByVal lineSpacing_ As Long = 0, Optional ByVal ignoreNewlines As Boolean = False) As Long
    Dim d As W_LABEL, n As Long, idx As Long
    Dim lineText() As String, lineWidth() As Long, lineHeight() As Long, lineOffsetX() As Long, lineOffsetY() As Long
    Set d = W_LABEL_(self)
    d.lineSpacing = lineSpacing_
    d.FontID = fontID
    n = iV_FormatText(string_, MaxWidth, FTEXT_LEFTJUSTIFY, fontID, ignoreNewlines, lineText, lineWidth, lineHeight, lineOffsetX, lineOffsetY)
    d.setTextLines n, lineText, lineWidth, lineHeight, lineOffsetX, lineOffsetY

    d.maxLineWidth = 0
    For idx = 0 To d.aTextLinesCount - 1
        d.maxLineWidth = MAX_(d.maxLineWidth, d.lineDimX(idx))
    Next

    d.wzTextClear
    For idx = 0 To d.aTextLinesCount - 1
        d.wzTextPush WzCachedText_new_text(d.lineText(idx), d.FontID, LABEL_DEFAULT_CACHE_EXPIRY)
    Next

    W_LABEL_setFormattedString = W_LABEL_requiredHeight(self)
End Function

Public Function W_LABEL_requiredHeight(ByVal self As WIDGET) As Long
    Dim d As W_LABEL, gapHeight As Long, textHeight As Long
    Set d = W_LABEL_(self)
    If d.aTextLinesCount = 0 Then
        W_LABEL_requiredHeight = 0
        Exit Function
    End If

    gapHeight = (d.aTextLinesCount - 1) * d.lineSpacing
    textHeight = d.lineOffY(d.aTextLinesCount - 1) + iV_GetTextLineSize(d.FontID)
    W_LABEL_requiredHeight = textHeight + gapHeight
End Function

' (#ifdef DEBUG_BOUNDING_BOXES: not defined)

Public Sub W_LABEL_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_LABEL, maxWidth As Long, textTotalHeight As Long, idx As Long, textBoundingBoxOffset As Vector2i
    Dim jy As Long, wzTextLine As WzCachedText, fx As Long, fw As Long, fy As Single, lineWidthLimit As Long
    Set d = W_LABEL_(self)
    maxWidth = 0
    textTotalHeight = 0
    For idx = 0 To d.wzTextCount - 1
        maxWidth = MAX_(maxWidth, d.wzTextAt(idx).deref().width())
        textTotalHeight = textTotalHeight + d.wzTextAt(idx).deref().lineSize()
        If idx < (d.wzTextCount - 1) Then
            textTotalHeight = textTotalHeight + d.lineSpacing
        End If
    Next

    If d.wzTextCount = 0 Then Exit Sub

    textBoundingBoxOffset = Vector2i_(0, 0)
    If (self.style And WLAB_ALIGNTOPLEFT) <> 0 Then ' Align top
        textBoundingBoxOffset.y = yOffset + self.y()
    ElseIf (self.style And WLAB_ALIGNBOTTOMLEFT) <> 0 Then ' Align bottom
        textBoundingBoxOffset.y = yOffset + self.y() + (self.height() - textTotalHeight)
    Else
        textBoundingBoxOffset.y = yOffset + self.y() + (self.height() - textTotalHeight) \ 2
    End If

    jy = 0
    d.isTruncated = False
    For idx = 0 To d.wzTextCount - 1
        Set wzTextLine = d.wzTextAt(idx)
        fx = 0
        If (self.style And WLAB_ALIGNCENTRE) <> 0 Then
            fw = wzTextLine.deref().width()
            fx = xOffset + self.x() + (self.width() - fw) \ 2
        ElseIf (self.style And WLAB_ALIGNRIGHT) <> 0 Then
            fw = wzTextLine.deref().width()
            fx = xOffset + self.x() + self.width() - fw
        Else
            fx = xOffset + self.x()
        End If

        fy = CSng(textBoundingBoxOffset.y) + CSng(jy) - CSng(wzTextLine.deref().aboveBase())

        lineWidthLimit = -1
        If d.canTruncate And (wzTextLine.deref().width() > self.width()) Then
            ' text would render outside the width of the label, so figure out a maxLineWidth that can be displayed (leaving room for ellipsis)
            lineWidthLimit = self.width() - iV_GetEllipsisWidth(d.FontID) - 2
        End If
        wzTextLine.deref().render CSng(textBoundingBoxOffset.x + fx), fy, d.fontColour, 0!, lineWidthLimit
        If lineWidthLimit > -1 Then
            ' Render ellipsis
            iV_DrawEllipsis d.FontID, CSng(textBoundingBoxOffset.x + fx + lineWidthLimit + 2), fy, d.fontColour
            d.isTruncated = True
        End If
        jy = jy + wzTextLine.deref().lineSize() + d.lineSpacing
    Next
End Sub

Public Function W_LABEL_getTip(ByVal self As WIDGET) As String
    Dim d As W_LABEL, labelText As String, i As Long
    Set d = W_LABEL_(self)
    If Len(d.pTip) = 0 And Not d.isTruncated Then
        W_LABEL_getTip = ""
        Exit Function
    End If

    If Not d.isTruncated Then
        W_LABEL_getTip = d.pTip
        Exit Function
    End If

    ' std::stringstream labelText;
    For i = 0 To d.aTextLinesCount - 1
        labelText = labelText & d.lineText(i) & vbLf
    Next

    If Len(d.pTip) <> 0 Then
        labelText = labelText & vbLf & "(" & d.pTip & ")"
    End If

    W_LABEL_getTip = labelText
End Function

Public Function W_LABEL_getString(ByVal self As WIDGET) As String
    Dim d As W_LABEL
    Set d = W_LABEL_(self)
    If d.aTextLinesCount = 0 Then W_LABEL_getString = "": Exit Function
    W_LABEL_getString = d.lineText(0)
End Function

Public Sub W_LABEL_setString(ByVal self As WIDGET, ByVal string_ As String)
    Dim d As W_LABEL
    Set d = W_LABEL_(self)
    If d.aTextLinesCount = 1 Then
        If d.lineText(0) = string_ Then
            ' no change - skip
            Exit Sub
        End If
    End If
    d.clearTextLines
    d.pushTextLine string_, 0, 0, 0, 0
    d.wzTextClear
    d.wzTextPush WzCachedText_new_text(string_, d.FontID, LABEL_DEFAULT_CACHE_EXPIRY)
    d.maxLineWidth = -1 ' delay calculating line width until it's requested
    self.dirty = True
End Sub

Public Sub W_LABEL_setTip(ByVal self As WIDGET, ByVal string_ As String)
    W_LABEL_(self).pTip = string_
End Sub

Public Sub W_LABEL_setTextAlignment(ByVal self As WIDGET, ByVal align As WzTextAlignment)
    self.style = self.style And Not (WLAB_ALIGNLEFT Or WLAB_ALIGNCENTRE Or WLAB_ALIGNRIGHT)
    self.style = self.style Or align
    self.dirty = True
End Sub

Public Sub W_LABEL_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_LABEL, i As Long
    Set d = W_LABEL_(self)
    If Not d.cacheNeverExpires Then
        For i = 0 To d.wzTextCount - 1
            d.wzTextAt(i).tick
        Next
    End If
End Sub

Public Function W_LABEL_getMaxLineWidth(ByVal self As WIDGET) As Long
    Dim d As W_LABEL
    Set d = W_LABEL_(self)
    If d.maxLineWidth > -1 Then W_LABEL_getMaxLineWidth = d.maxLineWidth: Exit Function
    If d.aTextLinesCount = 0 Then W_LABEL_getMaxLineWidth = 0: Exit Function
    ' delayed calculation of maxLineWidth until first requested
    d.maxLineWidth = iV_GetTextWidth(d.lineText(0), d.FontID)
    W_LABEL_getMaxLineWidth = d.maxLineWidth
End Function

' label.h's inline methods

' void setFont(iV_fonts font) and void setFont(iV_fonts font, PIELIGHT colour) (colour missing: the first)
Public Sub W_LABEL_setFont(ByVal self As WIDGET, ByVal font As iV_fonts, Optional ByVal colour As Variant)
    W_LABEL_(self).FontID = font
    If Not IsMissing(colour) Then
        W_LABEL_setFontColour self, CLng(colour)
    End If
End Sub

Public Sub W_LABEL_setFontColour(ByVal self As WIDGET, ByVal colour As Long)
    W_LABEL_(self).fontColour = colour
End Sub

Public Sub W_LABEL_setCanTruncate(ByVal self As WIDGET, ByVal canTruncate_ As Boolean)
    W_LABEL_(self).canTruncate = canTruncate_
End Sub

Public Sub W_LABEL_setCacheNeverExpires(ByVal self As WIDGET, ByVal value As Boolean)
    W_LABEL_(self).cacheNeverExpires = value
End Sub

Public Function W_LABEL_idealWidth(ByVal self As WIDGET) As Long
    W_LABEL_idealWidth = W_LABEL_getMaxLineWidth(self)
End Function

Public Function W_LABEL_idealHeight(ByVal self As WIDGET) As Long
    W_LABEL_idealHeight = W_LABEL_requiredHeight(self)
End Function
