Attribute VB_Name = "advcheckbox_c"
' src/titleui/widgets/advcheckbox.cpp, advcheckbox.h: class WzAdvCheckbox, the
' "Advanced Fancy Checkbox": a check image, a line of text and a description
' paragraph (its fields: WzAdvCheckbox.cls). WzAdvCheckbox::make is
' WzAdvCheckbox_make.
' (VB) widget class WzAdvCheckbox : W_BUTTON
Option Explicit

' (VB) a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' WzAdvCheckbox::WzAdvCheckbox() (and the members' initialisers)
Public Sub WzAdvCheckbox_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzAdvCheckbox ' (VB) virtual calls from here on reach WzAdvCheckbox's overrides
    With WzAdvCheckbox_(self)
        .imageDimensions = 16
        .cachedIdealWidth = 0
        .lastWidgetWidth = 0
        .outerHorizontalPadding = 5
        .innerHorizontalPadding = 10
        .outerVerticalPadding = 5
        .innerVerticalPadding = 5
        .checked = False
    End With
End Sub

Public Function WzAdvCheckbox_make(ByVal displayName As String, ByVal description As String) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzAdvCheckbox)
    WzAdvCheckbox_construct widget
    WzAdvCheckbox_initialize widget, displayName, description
    Set WzAdvCheckbox_make = widget
End Function

Private Sub recalcIdealWidth(ByVal self As WIDGET)
    With WzAdvCheckbox_(self)
        .cachedIdealWidth = .outerHorizontalPadding + .imageDimensions + .innerHorizontalPadding + iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID) + .outerHorizontalPadding
    End With
End Sub

Private Sub WzAdvCheckbox_initialize(ByVal self As WIDGET, ByVal displayName As String, ByVal description As String)
    WzAdvCheckbox_setDescription self, description
    self.setString displayName
End Sub

Public Function WzAdvCheckbox_idealWidth(ByVal self As WIDGET) As Long
    WzAdvCheckbox_idealWidth = WzAdvCheckbox_(self).cachedIdealWidth
End Function

Public Function WzAdvCheckbox_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzAdvCheckbox, result As Long, a As Long, descriptionHeight As Long
    Set d = WzAdvCheckbox_(self)
    a = d.wzText.lineSize()
    If d.imageDimensions > a Then a = d.imageDimensions
    result = d.outerVerticalPadding + a
    descriptionHeight = d.descriptionWidget.height()
    If descriptionHeight > 0 Then
        result = result + d.innerVerticalPadding + d.descriptionWidget.height()
    End If
    result = result + d.outerVerticalPadding
    WzAdvCheckbox_idealHeight = result
End Function

Public Sub WzAdvCheckbox_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim descriptionHighlightColor As Long
    W_BUTTON_highlight self, psContext
    descriptionHighlightColor = psPalette(WZCOL_TEXT_BRIGHT)
    descriptionHighlightColor = pielightWithAlpha(descriptionHighlightColor, ((pielightAlpha(descriptionHighlightColor) * 3) \ 4) And &HFF&)
    Paragraph_forceSetAllFontColor WzAdvCheckbox_(self).descriptionWidget, descriptionHighlightColor
End Sub

Public Sub WzAdvCheckbox_highlightLost(ByVal self As WIDGET)
    W_BUTTON_highlightLost self
    Paragraph_forceSetAllFontColor WzAdvCheckbox_(self).descriptionWidget, psPalette(WZCOL_TEXT_MEDIUM)
End Sub

Public Sub WzAdvCheckbox_setImageDimensions(ByVal self As WIDGET, ByVal imageSize As Long)
    WzAdvCheckbox_(self).imageDimensions = imageSize
    recalcIdealWidth self
End Sub

Public Function WzAdvCheckbox_isChecked(ByVal self As WIDGET) As Boolean
    WzAdvCheckbox_isChecked = WzAdvCheckbox_(self).checked
End Function

Public Sub WzAdvCheckbox_setIsChecked(ByVal self As WIDGET, ByVal val_ As Boolean)
    WzAdvCheckbox_(self).checked = val_
End Sub

Public Sub WzAdvCheckbox_setOuterHorizontalPadding(ByVal self As WIDGET, ByVal padding_ As Long)
    WzAdvCheckbox_(self).outerHorizontalPadding = padding_
    recalcIdealWidth self
End Sub

Public Sub WzAdvCheckbox_setInnerHorizontalPadding(ByVal self As WIDGET, ByVal padding_ As Long)
    WzAdvCheckbox_(self).innerHorizontalPadding = padding_
    recalcIdealWidth self
End Sub

Public Sub WzAdvCheckbox_setOuterVerticalPadding(ByVal self As WIDGET, ByVal padding_ As Long)
    WzAdvCheckbox_(self).outerVerticalPadding = padding_
    recalcIdealWidth self
End Sub

Public Sub WzAdvCheckbox_setInnerVerticalPadding(ByVal self As WIDGET, ByVal padding_ As Long)
    WzAdvCheckbox_(self).innerVerticalPadding = padding_
    recalcIdealWidth self
End Sub

' Respond to a mouse button up
Public Sub WzAdvCheckbox_released(ByVal self As WIDGET, context As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim clickAndReleaseOnThisButton As Boolean
    clickAndReleaseOnThisButton = ((W_BUTTON_(self).state And WBUT_DOWN) <> 0) ' relies on W_BUTTON handling to properly set WBUT_DOWN
    If clickAndReleaseOnThisButton Then
        ' toggle click-lock ("checked") state
        WzAdvCheckbox_setIsChecked self, Not WzAdvCheckbox_isChecked(self)
    End If
    W_BUTTON_released self, context, key
End Sub

Public Sub WzAdvCheckbox_setString(ByVal self As WIDGET, ByVal string_ As String)
    W_BUTTON_setString self, string_
    recalcIdealWidth self
End Sub

Public Sub WzAdvCheckbox_setDescription(ByVal self As WIDGET, ByVal description As String)
    Dim d As WzAdvCheckbox
    Set d = WzAdvCheckbox_(self)
    If Not d.descriptionWidget Is Nothing Then
        self.detach d.descriptionWidget
    End If
    Set d.descriptionWidget = Paragraph_new()
    Paragraph_setFont d.descriptionWidget, font_small
    Paragraph_setFontColour d.descriptionWidget, psPalette(WZCOL_TEXT_MEDIUM)
    d.descriptionWidget.setGeometry 0, 0, 400, 40
    Paragraph_addText d.descriptionWidget, description
    d.descriptionWidget.setTransparentToMouse True
    self.attach d.descriptionWidget
End Sub

Public Sub WzAdvCheckbox_geometryChanged(ByVal self As WIDGET)
    Dim d As WzAdvCheckbox, descX0 As Long, descY0 As Long, a As Long
    Set d = WzAdvCheckbox_(self)
    ' reposition description paragraph beneath the image and line of text
    descX0 = d.outerHorizontalPadding + d.imageDimensions + d.innerHorizontalPadding
    a = d.wzText.lineSize()
    If d.imageDimensions > a Then a = d.imageDimensions
    descY0 = d.outerVerticalPadding + a + d.innerVerticalPadding
    d.descriptionWidget.setGeometry descX0, descY0, self.width() - descX0 - d.outerHorizontalPadding, self.height() - descY0 - d.outerVerticalPadding ' while we specify a height, the paragraph widget will calculate and set its to its needed full height, given the supplied width
End Sub

Public Sub WzAdvCheckbox_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzAdvCheckbox, x0 As Long, y0 As Long, w As Long, pText As String
    Dim isDown As Boolean, isChecked As Boolean, isDisabled As Boolean, isHighlight As Boolean
    Dim fontID As Long, result As iV_fonts, availableButtonTextWidth As Long, textColor As Long, textXOffset As Long
    Dim textX0 As Long, textY0 As Long, maxTextDisplayableWidth As Long, maxDisplayableMainTextWidth As Long
    Dim isTruncated As Boolean, imgPosX0 As Long, imgPosY0 As Long, imgColor As Long, imageID As Long
    Set d = WzAdvCheckbox_(self)
    pText = W_BUTTON_(self).pText
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isChecked = d.checked
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And self.isMouseOverWidget() '!isDisabled && ((getState() & WBUT_HIGHLIGHT) != 0);

    fontID = d.wzText.getFontID()
    If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzText.getText() <> pText Then
        fontID = W_BUTTON_(self).FontID
    End If
    d.wzText.setText pText, fontID

    availableButtonTextWidth = w - (d.outerHorizontalPadding * 2) - (d.imageDimensions + d.innerHorizontalPadding)
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

    ' TODO: Draw background, if configured

    ' Draw the main text
    If isHighlight Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isDisabled Then
        textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
    End If

    textXOffset = d.outerHorizontalPadding + d.imageDimensions + d.innerHorizontalPadding
    textX0 = x0 + textXOffset
    textY0 = y0 + d.outerVerticalPadding - d.wzText.aboveBase()

    maxTextDisplayableWidth = w - textXOffset - d.outerHorizontalPadding
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

    ' Draw the image to the left
    imgPosX0 = x0 + d.outerHorizontalPadding
    imgPosY0 = y0 + d.outerVerticalPadding + (d.wzText.lineSize() - d.imageDimensions) \ 2
    If isHighlight Then imgColor = psPalette(WZCOL_TEXT_BRIGHT) Else imgColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isDisabled Then
        imgColor = pielightWithAlpha(imgColor, pielightAlpha(imgColor) \ 2)
    ElseIf isDown Then
        imgColor = psPalette(WZCOL_TEXT_DARK)
    End If
    If isChecked Then imageID = IMAGE_CHECK_SQUARE_FILL Else imageID = IMAGE_CHECK_SQUARE_EMPTY
    iV_DrawImageFileAnisotropicTint FrontImages, imageID, imgPosX0, imgPosY0, d.imageDimensions, d.imageDimensions, imgColor
End Sub
