Attribute VB_Name = "titleformwrapper_c"
' src/titleui/widgets/titleformwrapper.cpp, titleformwrapper.h: the title
' screens' form with its sideways title (makeTitleFormWrapper), from the classes
' IntFormAnimatedWrapper, TitleFormSideText and TitleFormWrapper (their fields:
' the classes of the same names). GetTitleStringFunc (std::function<WzString(void)>)
' is a WzFunc returning the String.
' (VB) widget class IntFormAnimatedWrapper : IntFormAnimated
' (VB) widget class TitleFormSideText : WIDGET
' (VB) widget class TitleFormWrapper : WIDGET
Option Explicit

' MARK: - IntFormAnimatedWrapper

' IntFormAnimatedWrapper() : IntFormAnimated(true)
Public Sub IntFormAnimatedWrapper_construct(ByVal self As WIDGET)
    IntFormAnimated_construct self, True
    self.kind = K_IntFormAnimatedWrapper ' (VB) virtual calls from here on reach IntFormAnimatedWrapper's overrides
End Sub

Private Sub IntFormAnimatedWrapper_initialize(ByVal self As WIDGET, ByVal displayWidget_ As WIDGET)
    Set IntFormAnimatedWrapper_(self).displayWidget = displayWidget_
    self.attach IntFormAnimatedWrapper_(self).displayWidget
End Sub

Public Function IntFormAnimatedWrapper_make(ByVal displayWidget As WIDGET) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_IntFormAnimatedWrapper)
    IntFormAnimatedWrapper_construct result
    IntFormAnimatedWrapper_initialize result, displayWidget
    Set IntFormAnimatedWrapper_make = result
End Function

Public Function IntFormAnimatedWrapper_idealWidth(ByVal self As WIDGET) As Long
    IntFormAnimatedWrapper_idealWidth = 2 + IntFormAnimatedWrapper_(self).displayWidget.idealWidth()
End Function

Public Function IntFormAnimatedWrapper_idealHeight(ByVal self As WIDGET) As Long
    IntFormAnimatedWrapper_idealHeight = 2 + IntFormAnimatedWrapper_(self).displayWidget.idealHeight()
End Function

Public Sub IntFormAnimatedWrapper_geometryChanged(ByVal self As WIDGET)
    Dim w As Long, h As Long
    w = self.width()
    h = self.height()
    IntFormAnimatedWrapper_(self).displayWidget.setGeometry 1, 1, w - 2, h - 2
End Sub

' MARK: - TitleFormSideText

' TitleFormSideText() { } (and the members' initialisers)
Public Sub TitleFormSideText_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_TitleFormSideText ' (VB) virtual calls from here on reach TitleFormSideText's overrides
    With TitleFormSideText_(self)
        .horizontalPadding = 0
        .verticalPadding = 0
        .lastWidgetWidth = 0
        .isTruncated = False
    End With
End Sub

Public Function TitleFormSideText_make(ByVal font As Long, ByVal text As String) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_TitleFormSideText)
    TitleFormSideText_construct result
    TitleFormSideText_initialize result, font, text
    Set TitleFormSideText_make = result
End Function

Private Sub TitleFormSideText_initialize(ByVal self As WIDGET, ByVal font As Long, ByVal str_ As String)
    TitleFormSideText_(self).fontID = font
    self.setString str_
End Sub

Public Sub TitleFormSideText_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As TitleFormSideText, x0 As Long, y0 As Long, w As Long, h As Long, currFontID As Long, result As iV_fonts
    Dim availableButtonTextHeight As Long, maxTextDisplayableHeight As Long, textX0 As Long, textY0 As Long
    Set d = TitleFormSideText_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()

    currFontID = d.wzText.getFontID()
    If d.fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzText.getText() <> d.text Then
        currFontID = d.fontID
    End If
    d.wzText.setText d.text, currFontID

    availableButtonTextHeight = h - (d.verticalPadding * 2)
    If d.wzText.width() > availableButtonTextHeight Then
        ' text would exceed the bounds of the button area
        ' try to shrink font so it fits
        Do
            If currFontID = font_small Or currFontID = font_bar Then
                Exit Do
            End If
            If Not iV_ShrinkFont(currFontID, result) Then
                Exit Do
            End If
            currFontID = result
            d.wzText.setText d.text, currFontID
        Loop While d.wzText.width() > availableButtonTextHeight
    End If
    d.lastWidgetWidth = self.width()

    maxTextDisplayableHeight = availableButtonTextHeight
    d.isTruncated = maxTextDisplayableHeight < d.wzText.width()
    If d.isTruncated Then
        maxTextDisplayableHeight = maxTextDisplayableHeight - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
    End If

    textX0 = Fix(x0 + (w - d.wzText.lineSize()) \ 2 - CSng(d.wzText.aboveBase()))
    If d.wzText.width() < availableButtonTextHeight Then textY0 = y0 + d.wzText.width() Else textY0 = y0 + availableButtonTextHeight
    d.wzText.render textX0 + 2, textY0 + 2, psPalette(WZCOL_GREY), 270!, maxTextDisplayableHeight
    d.wzText.render textX0, textY0, psPalette(WZCOL_TEXT_BRIGHT), 270!, maxTextDisplayableHeight

    If d.isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzText.getFontID(), textX0, textY0 + maxTextDisplayableHeight + 2, psPalette(WZCOL_TEXT_BRIGHT), 270!
    End If
End Sub

Public Function TitleFormSideText_getString(ByVal self As WIDGET) As String
    TitleFormSideText_getString = TitleFormSideText_(self).text
End Function

Public Sub TitleFormSideText_setString(ByVal self As WIDGET, ByVal string_ As String)
    TitleFormSideText_(self).text = string_
    TitleFormSideText_(self).wzText.setText TitleFormSideText_(self).text, TitleFormSideText_(self).fontID
End Sub

Public Function TitleFormSideText_idealWidth(ByVal self As WIDGET) As Long
    ' is the height of the text
    TitleFormSideText_idealWidth = (TitleFormSideText_(self).horizontalPadding * 2) + TitleFormSideText_(self).wzText.lineSize()
End Function

Public Function TitleFormSideText_idealHeight(ByVal self As WIDGET) As Long
    ' is the width of the text
    TitleFormSideText_idealHeight = (TitleFormSideText_(self).verticalPadding * 2) + TitleFormSideText_(self).wzText.width()
End Function

' MARK: - TitleFormWrapper

' TitleFormWrapper() { } (and the members' initialisers)
Public Sub TitleFormWrapper_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_TitleFormWrapper ' (VB) virtual calls from here on reach TitleFormWrapper's overrides
    With TitleFormWrapper_(self)
        .outerPaddingX = 10
        .titleFormSpacing = 10
        .outerPaddingY = 20
    End With
End Sub

Private Sub TitleFormWrapper_initialize(ByVal self As WIDGET, ByVal getTitleStringFunc_ As WzFunc, ByVal displayWidget_ As WIDGET)
    Dim d As TitleFormWrapper
    Set d = TitleFormWrapper_(self)
    Set d.getTitleStringFunc = getTitleStringFunc_
    Set d.sideText = TitleFormSideText_make(font_large, WzFunc_call(d.getTitleStringFunc))
    self.attach d.sideText
    Set d.formWrapper = IntFormAnimatedWrapper_make(displayWidget_)
    self.attach d.formWrapper
End Sub

Public Function TitleFormWrapper_make(ByVal getTitleStringFunc As WzFunc, ByVal displayWidget As WIDGET) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_TitleFormWrapper)
    TitleFormWrapper_construct result
    TitleFormWrapper_initialize result, getTitleStringFunc, displayWidget
    Set TitleFormWrapper_make = result
End Function

Public Function TitleFormWrapper_idealWidth(ByVal self As WIDGET) As Long
    Dim d As TitleFormWrapper, outerTitleWidth As Long
    Set d = TitleFormWrapper_(self)
    outerTitleWidth = d.outerPaddingX + d.sideText.idealWidth() + d.titleFormSpacing
    TitleFormWrapper_idealWidth = (outerTitleWidth * 2) + d.formWrapper.idealWidth()
End Function

Public Function TitleFormWrapper_idealHeight(ByVal self As WIDGET) As Long
    Dim d As TitleFormWrapper, a As Long, b As Long
    Set d = TitleFormWrapper_(self)
    a = d.sideText.idealHeight()
    b = d.formWrapper.idealHeight()
    If b > a Then a = b
    TitleFormWrapper_idealHeight = (d.outerPaddingY * 2) + a
End Function

Public Sub TitleFormWrapper_geometryChanged(ByVal self As WIDGET)
    Dim d As TitleFormWrapper, w As Long, h As Long
    Dim sideTextWidth As Long, sideTextHeight As Long, formUnavailableWidth As Long, maxAvailableFormWidth As Long
    Dim formWidth As Long, maxAvailableHeight As Long, formHeight As Long, childY0 As Long, formX0 As Long, sideTextX0 As Long
    Set d = TitleFormWrapper_(self)
    w = self.width()
    h = self.height()
    If w = 0 Or h = 0 Then
        Exit Sub
    End If

    sideTextWidth = d.sideText.idealWidth()
    sideTextHeight = d.sideText.idealHeight()

    formUnavailableWidth = (d.outerPaddingX + sideTextWidth + d.titleFormSpacing) * 2
    maxAvailableFormWidth = w - formUnavailableWidth
    formWidth = d.formWrapper.idealWidth()
    If maxAvailableFormWidth < formWidth Then formWidth = maxAvailableFormWidth
    maxAvailableHeight = h - (d.outerPaddingY * 2)
    formHeight = d.formWrapper.idealHeight()
    If maxAvailableHeight < formHeight Then formHeight = maxAvailableHeight

    childY0 = (h - formHeight) \ 2

    formX0 = (w - formWidth) \ 2
    d.formWrapper.setGeometry formX0, childY0, formWidth, formHeight

    sideTextX0 = formX0 - (sideTextWidth + d.titleFormSpacing)
    d.sideText.setGeometry sideTextX0, childY0, sideTextWidth, sideTextHeight
End Sub

Public Sub TitleFormWrapper_screenSizeDidChange(ByVal self As WIDGET, ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim d As TitleFormWrapper
    Set d = TitleFormWrapper_(self)
    ' refresh the side text (in case the current language or display scaling changed)
    d.sideText.setString WzFunc_call(d.getTitleStringFunc)
    WIDGET_screenSizeDidChange self, oldWidth, oldHeight, newWidth, newHeight
    If oldWidth = newWidth And oldHeight = newHeight Then
        self.geometryChanged ' may need to relayout title text anyway if it changed
    End If
End Sub

' MARK: -

Public Function makeTitleFormWrapper(ByVal getTitleStringFunc As WzFunc, ByVal displayWidget As WIDGET) As WIDGET
    Set makeTitleFormWrapper = TitleFormWrapper_make(getTitleStringFunc, displayWidget)
End Function
