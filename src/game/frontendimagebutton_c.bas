Attribute VB_Name = "frontendimagebutton_c"
' src/titleui/widgets/frontendimagebutton.cpp, frontendimagebutton.h: class
' WzFrontendImageButton, a button with a frontend.img image and text (its fields:
' WzFrontendImageButton.cls; an optional<UWORD> / optional<PIELIGHT> is a has_
' flag and the value). WzFrontendImageButton::make is
' WzFrontendImageButton_make.
' (VB) widget class WzFrontendImageButton : W_BUTTON
Option Explicit

' WzFrontendImageButton::BorderDrawMode
Public Enum BorderDrawMode
    BorderDrawMode_Never
    BorderDrawMode_Highlighted
    BorderDrawMode_Always
End Enum

' (VB) a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' WzFrontendImageButton::WzFrontendImageButton()
Public Sub WzFrontendImageButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzFrontendImageButton ' (VB) virtual calls from here on reach WzFrontendImageButton's overrides
    With WzFrontendImageButton_(self)
        .has_frontendImgID = False       ' = nullopt
        .has_customTextColor = False
        .has_customTextHighlightColor = False
        .has_customImgColor = False
        .has_customBackgroundColor = False
        .missingImage = False
        .imageDimensions = 16
        .cachedIdealWidth = 0
        .lastWidgetWidth = 0
        .horizontalPadding = 5
        .verticalPadding = 3
        .imageHorizontalOffset = 0
        .borderDrawMode_ = BorderDrawMode_Always
    End With
End Sub

' static std::shared_ptr<WzFrontendImageButton> make(optional<UWORD> frontendImgID)
Public Function WzFrontendImageButton_make(ByVal has_frontendImgID As Boolean, ByVal frontendImgID As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzFrontendImageButton)
    WzFrontendImageButton_construct widget
    WzFrontendImageButton_setImage widget, has_frontendImgID, frontendImgID
    Set WzFrontendImageButton_make = widget
End Function

Public Sub WzFrontendImageButton_setImage(ByVal self As WIDGET, ByVal has_newFrontendImgID As Boolean, ByVal newFrontendImgID As Long)
    Dim d As WzFrontendImageButton
    Set d = WzFrontendImageButton_(self)
    d.missingImage = False
    ' validate newFrontendImgID value
    If has_newFrontendImgID Then
        If newFrontendImgID >= IMAGEFILE_numImages(FrontImages) Then
            ' out of bounds value (likely a mod has an incomplete / older frontend.img file)
            debug_ 49, LOG_ERROR, "setImage", "frontend.img missing expected image id ""%u"" (num loaded: %zu) - remove any mods", newFrontendImgID, IMAGEFILE_numImages(FrontImages)
            has_newFrontendImgID = False
            d.missingImage = True
        End If
    End If
    d.has_frontendImgID = has_newFrontendImgID
    d.frontendImgID = newFrontendImgID
    recalcIdealWidth self
End Sub

Public Sub WzFrontendImageButton_setImageDimensions(ByVal self As WIDGET, ByVal imageSize As Long)
    If imageSize = WzFrontendImageButton_(self).imageDimensions Then
        Exit Sub
    End If
    WzFrontendImageButton_(self).imageDimensions = imageSize
    recalcIdealWidth self
End Sub

Public Sub WzFrontendImageButton_setPadding(ByVal self As WIDGET, ByVal newHorizPadding As Long, ByVal newVertPadding As Long)
    WzFrontendImageButton_(self).horizontalPadding = newHorizPadding
    WzFrontendImageButton_(self).verticalPadding = newVertPadding
    recalcIdealWidth self
End Sub

Public Sub WzFrontendImageButton_setImageHorizontalOffset(ByVal self As WIDGET, ByVal xOffset As Long)
    WzFrontendImageButton_(self).imageHorizontalOffset = xOffset
End Sub

Public Sub WzFrontendImageButton_setCustomTextColors(ByVal self As WIDGET, ByVal has_textColor As Boolean, ByVal textColor As Long, ByVal has_highlightedTextColor As Boolean, ByVal highlightedTextColor As Long)
    With WzFrontendImageButton_(self)
        .has_customTextColor = has_textColor
        .customTextColor = textColor
        .has_customTextHighlightColor = has_highlightedTextColor
        .customTextHighlightColor = highlightedTextColor
    End With
End Sub

Public Sub WzFrontendImageButton_setCustomImageColor(ByVal self As WIDGET, ByVal has_color As Boolean, ByVal color As Long)
    WzFrontendImageButton_(self).has_customImgColor = has_color
    WzFrontendImageButton_(self).customImgColor = color
End Sub

Public Sub WzFrontendImageButton_setBackgroundColor(ByVal self As WIDGET, ByVal has_color As Boolean, ByVal color As Long)
    WzFrontendImageButton_(self).has_customBackgroundColor = has_color
    WzFrontendImageButton_(self).customBackgroundColor = color
End Sub

' (override of WIDGET::setString)
Public Sub WzFrontendImageButton_setString(ByVal self As WIDGET, ByVal string_ As String)
    W_BUTTON_setString self, string_
    recalcIdealWidth self
End Sub

Private Function shouldDisplayImage(ByVal self As WIDGET) As Boolean
    With WzFrontendImageButton_(self)
        shouldDisplayImage = (.has_frontendImgID Or .missingImage) And .imageDimensions > 0
    End With
End Function

Private Sub recalcIdealWidth(ByVal self As WIDGET)
    Dim d As WzFrontendImageButton, textWidth As Long
    Set d = WzFrontendImageButton_(self)
    d.cachedIdealWidth = d.horizontalPadding
    If shouldDisplayImage(self) Then
        d.cachedIdealWidth = d.cachedIdealWidth + d.imageDimensions + d.horizontalPadding
    End If
    textWidth = iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID)
    If textWidth > 0 Then
        d.cachedIdealWidth = d.cachedIdealWidth + textWidth + d.horizontalPadding
    End If
End Sub

Public Function WzFrontendImageButton_idealWidth(ByVal self As WIDGET) As Long
    WzFrontendImageButton_idealWidth = WzFrontendImageButton_(self).cachedIdealWidth
End Function

Public Function WzFrontendImageButton_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzFrontendImageButton, lineSize As Long
    Set d = WzFrontendImageButton_(self)
    lineSize = iV_GetTextLineSize(W_BUTTON_(self).FontID)
    If d.imageDimensions > lineSize Then lineSize = d.imageDimensions
    WzFrontendImageButton_idealHeight = (d.verticalPadding * 2) + lineSize
End Function

Public Sub WzFrontendImageButton_setBorderDrawMode(ByVal self As WIDGET, ByVal mode As BorderDrawMode)
    WzFrontendImageButton_(self).borderDrawMode_ = mode
End Sub

Public Sub WzFrontendImageButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzFrontendImageButton, x0 As Long, y0 As Long, w As Long
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, hasImage As Boolean
    Dim textColor As Long, boxX0 As Long, boxY0 As Long, boxX1 As Long, boxY1 As Long
    Dim highlightRectDimensions As Long, borderColor As Long, bHasText As Boolean, fontID As Long, result As iV_fonts
    Dim availableButtonTextWidth As Long, textX0 As Long, textY0 As Long, maxTextDisplayableWidth As Long
    Dim maxDisplayableMainTextWidth As Long, isTruncated As Boolean, imgPosX0 As Long, imgPosY0 As Long, imgColor As Long
    Dim pText As String
    Set d = WzFrontendImageButton_(self)
    pText = W_BUTTON_(self).pText
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)
    hasImage = shouldDisplayImage(self)
    If isHighlight Then
        If d.has_customTextHighlightColor Then textColor = d.customTextHighlightColor Else textColor = psPalette(WZCOL_TEXT_BRIGHT)
    Else
        If d.has_customTextColor Then textColor = d.customTextColor Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
    End If
    If isDisabled Then
        textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
    End If

    If d.has_customBackgroundColor Then
        boxX0 = x0
        boxY0 = y0
        boxX1 = boxX0 + self.width()
        boxY1 = boxY0 + self.height()
        pie_UniTransBoxFill boxX0, boxY0, boxX1, boxY1, d.customBackgroundColor
    End If

    Select Case d.borderDrawMode_
    Case BorderDrawMode_Never
    Case BorderDrawMode_Highlighted
        highlightRectDimensions = self.width()
        If self.height() < highlightRectDimensions Then highlightRectDimensions = self.height()
        If d.imageDimensions + 2 > highlightRectDimensions Then highlightRectDimensions = d.imageDimensions + 2
        boxX0 = x0 + (self.width() - highlightRectDimensions) \ 2
        boxY0 = y0 + (self.height() - highlightRectDimensions) \ 2
        If isDown Then
            pie_UniTransBoxFill boxX0 + 1, boxY0 + 1, boxX0 + highlightRectDimensions, boxY0 + highlightRectDimensions, psPalette(WZCOL_FORM_DARK)
        End If
        If (isHighlight Or isDown) And Not isDisabled Then
            ' Display highlight rect
            If isDown Then borderColor = psPalette(WZCOL_TEXT_BRIGHT) Else borderColor = psPalette(WZCOL_TEXT_MEDIUM)
            iV_Box boxX0 + 1, boxY0 + 1, boxX0 + highlightRectDimensions, boxY0 + highlightRectDimensions, borderColor
        End If
    Case BorderDrawMode_Always
        ' Display outer border
        If isDown Then
            borderColor = psPalette(WZCOL_TEXT_DARK)
        ElseIf isHighlight Then
            borderColor = psPalette(WZCOL_TEXT_BRIGHT)
        Else
            borderColor = psPalette(WZCOL_TEXT_MEDIUM)
        End If
        If isDisabled Then
            borderColor = pielightWithAlpha(borderColor, pielightAlpha(borderColor) \ 2)
        End If
        boxX0 = x0
        boxY0 = y0
        boxX1 = boxX0 + self.width()
        boxY1 = boxY0 + self.height()
        iV_Box boxX0 + 1, boxY0 + 1, boxX1, boxY1, borderColor
    End Select

    bHasText = Len(pText) > 0
    If bHasText Then
        ' Display text
        fontID = d.wzText.getFontID()
        If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzText.getText() <> pText Then
            fontID = W_BUTTON_(self).FontID
        End If
        d.wzText.setText pText, fontID
        availableButtonTextWidth = w - (d.horizontalPadding * 2)
        If hasImage Then availableButtonTextWidth = availableButtonTextWidth - (d.imageDimensions + d.horizontalPadding + d.imageHorizontalOffset)
        If d.wzText.width() > availableButtonTextWidth Then
            ' text would exceed the bounds of the button area
            ' try to shrink font so it fits
            Do
                If fontID = font_small Or fontID = font_bar Then
                    Exit Do
                End If
                If Not iV_ShrinkFont(fontID, result) Then
                    Exit Do
                End If
                fontID = result
                d.wzText.setText pText, fontID
            Loop While d.wzText.width() > availableButtonTextWidth
        End If
        d.lastWidgetWidth = self.width()

        ' Draw the main text
        textX0 = x0 + d.horizontalPadding
        If hasImage Then textX0 = textX0 + d.imageDimensions + d.horizontalPadding + d.imageHorizontalOffset
        textY0 = Fix(y0 + (self.height() - d.wzText.lineSize()) \ 2 - CSng(d.wzText.aboveBase()))
        maxTextDisplayableWidth = w - (d.horizontalPadding * 2)
        maxDisplayableMainTextWidth = maxTextDisplayableWidth
        isTruncated = maxDisplayableMainTextWidth < d.wzText.width()
        If isTruncated Then
            maxDisplayableMainTextWidth = maxDisplayableMainTextWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
        End If
        d.wzText.render textX0, textY0, textColor, 0!, maxDisplayableMainTextWidth
        If isTruncated Then
            ' Render ellipsis
            iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxDisplayableMainTextWidth + 2, textY0, textColor
        End If
    End If

    If bHasText Then
        ' Display the image to the left
        imgPosX0 = x0 + d.horizontalPadding + d.imageHorizontalOffset
    Else
        ' Display the image horizontally centered
        imgPosX0 = x0 + (self.width() - d.imageDimensions) \ 2
    End If
    imgPosY0 = y0 + (self.height() - d.imageDimensions) \ 2
    If d.has_frontendImgID Then
        If d.has_customImgColor Then imgColor = d.customImgColor Else imgColor = textColor
        iV_DrawImageFileAnisotropicTint FrontImages, d.frontendImgID, imgPosX0, imgPosY0, d.imageDimensions, d.imageDimensions, imgColor
    ElseIf d.missingImage Then
        pie_UniTransBoxFill imgPosX0, imgPosY0, imgPosX0 + d.imageDimensions, imgPosY0 + d.imageDimensions, psPalette(WZCOL_RED)
    End If
End Sub
