Attribute VB_Name = "optionsforms_c"
' src/titleui/options/optionsforms.cpp and optionsforms.h: the options browser (OptionsBrowserForm,
' createOptionsBrowser) with its switchers (HorizontalButtonsWidget, HorizontalScrollButton,
' WzClippableButton), whole. The widget classes follow docs/CONVENTIONS.md, "Widgets" (their
' fields: the data classes <Class>.cls). (VB) The weak pointers that scheduled tasks lock
' (HorizontalButtonsWidget's, OptionsBrowserForm's) are widgweak_vb serials.
Option Explicit

' (VB) widget class WzClippableButton : W_BUTTON
' (VB) widget class HorizontalScrollButton : W_BUTTON
' (VB) widget class HorizontalButtonsWidget : WIDGET
' (VB) widget class OptionsBrowserForm : WIDGET

Public Const FN_HBW_scrollLeftClick As Long = 4100          ' HorizontalButtonsWidget::initialize, scrollLeftButton->addOnClickHandler([weakSelf](W_BUTTON&) {...})
Public Const FN_HBW_scrollRightClick As Long = 4101         ' HorizontalButtonsWidget::initialize, scrollRightButton->addOnClickHandler([weakSelf](W_BUTTON&) {...})
Public Const FN_HBW_buttonClick As Long = 4102              ' HorizontalButtonsWidget::addButton, [weakSelf, newButtonIdx](W_BUTTON& but)
Public Const FN_HBW_buttonClickTask As Long = 4103          ' HorizontalButtonsWidget::addButton, widgScheduleTask([weakSelf, newButtonIdx]() {...})
Public Const FN_OBF_formSwitcherSelect As Long = 4104       ' OptionsBrowserForm::initialize, [weakSelf](const WzString& untranslatedTitle)
Public Const FN_OBF_sectionSwitcherSelect As Long = 4105    ' OptionsBrowserForm::initialize, [weakSelf](const WzString& sectionId)
Public Const FN_OBF_configDirLinkCalcLayout As Long = 4106  ' OptionsBrowserForm::initialize, LAMBDA_CALCLAYOUT_SIMPLE({...})
Public Const FN_OBF_configDirLinkClick As Long = 4107       ' OptionsBrowserForm::initialize, addOnClickHandler([](W_BUTTON&) {...})
Public Const FN_OBF_configDirLinkTask As Long = 4108        ' OptionsBrowserForm::initialize, widgScheduleTask([]() {...})
Public Const FN_createOptionsBrowser_informLanguage As Long = 4109 ' createOptionsBrowser, informOnLanguageChangeHandler: [weakOptionsBrowserForm]()
Public Const FN_createOptionsBrowser_interface As Long = 4110 ' createOptionsBrowser, [inGame, informOnLanguageChangeHandler]() { return makeInterfaceOptionsForm(...); }
Public Const FN_makeDefaultsOptionsForm As Long = 4111      ' std::shared_ptr<OptionsForm> makeDefaultsOptionsForm()
Public Const FN_makeGraphicsOptionsForm As Long = 4112      ' std::shared_ptr<OptionsForm> makeGraphicsOptionsForm()
Public Const FN_makeAudioOptionsForm As Long = 4113         ' std::shared_ptr<OptionsForm> makeAudioOptionsForm()
Public Const FN_makeControlsOptionsForm As Long = 4114      ' std::shared_ptr<OptionsForm> makeControlsOptionsForm()
Public Const FN_makeWindowOptionsForm As Long = 4115        ' std::shared_ptr<OptionsForm> makeWindowOptionsForm()

' enum class OptionsBrowserForm::Modes
Public Enum OptionsBrowserForm_Modes
    OptionsBrowserForm_Modes_Interface
    OptionsBrowserForm_Modes_Defaults
    OptionsBrowserForm_Modes_Graphics
    OptionsBrowserForm_Modes_Audio
    OptionsBrowserForm_Modes_Controls
    OptionsBrowserForm_Modes_Window
End Enum

' OptionsBrowserForm::OptionsFormDetails
Public Type OptionsBrowserForm_OptionsFormDetails
    mode As OptionsBrowserForm_Modes
    untranslatedTitle As String
    createFormFunc As WzFunc                ' CreateOptionsFormFunc: std::shared_ptr<OptionsForm>(void)
    optionsForm As WIDGET
End Type

' How far the mouse has to move to start a drag
Private Const HorizontalButtonsWidgetDragThreshold As Long = 5

' OptionsBrowserForm's const members
Private Const OptionsBrowserForm_paddingBeforeOptionsForm As Long = 10
Private Const OptionsBrowserForm_topPadding As Long = 0
Private Const OptionsBrowserForm_leftOpenConfigDirLinkPadding As Long = 10
Private Const OptionsBrowserForm_bottomPadding As Long = 20

Private Const MIN_IDEAL_OPTIONS_BROWSER_FORM_WIDTH As Long = 640
Private Const MIN_IDEAL_OPTIONS_BROWSER_FORM_HEIGHT As Long = 640

' (VB) PIELIGHT's byte.a, read and set
Private Function pielightA(ByVal c As Long) As Long
    pielightA = SHRU(c, 24) And &HFF&
End Function

Private Function pielightWithA(ByVal c As Long, ByVal a As Long) As Long
    pielightWithA = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' MARK: -

' Note: Currently does not support W_BUTTON's:
' - setImages functions
' - progressBorder functionality

' (VB) std::make_shared<WzClippableButton>()
Public Function WzClippableButton_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_WzClippableButton)
    WzClippableButton_construct w
    Set WzClippableButton_new = w
End Function

Public Sub WzClippableButton_construct(ByVal self As WIDGET)
    Dim d As WzClippableButton
    W_BUTTON_construct self
    self.kind = K_WzClippableButton
    Set d = WzClippableButton_(self)
    Set d.wzText = WzCachedText_new()
    d.horizontalPadding = 15
    d.verticalPadding = 10
End Sub

Public Sub WzClippableButton_setFont(ByVal self As WIDGET, ByVal font As iV_fonts)
    Dim b As W_BUTTON
    Set b = W_BUTTON_(self)
    b.FontID = font
    WzClippableButton_(self).wzText.setText b.pText, b.FontID
End Sub

Public Sub WzClippableButton_setPadding(ByVal self As WIDGET, ByVal horizontalPadding_ As Long, ByVal verticalPadding_ As Long)
    Dim d As WzClippableButton
    Set d = WzClippableButton_(self)
    d.horizontalPadding = horizontalPadding_
    d.verticalPadding = verticalPadding_
End Sub

Public Sub WzClippableButton_setString(ByVal self As WIDGET, ByVal string_ As String)
    W_BUTTON_setString self, string_
    WzClippableButton_(self).wzText.setText string_, W_BUTTON_(self).FontID
End Sub

Public Sub WzClippableButton_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzClippableButton_(self).wzText.tick
End Sub

Public Sub WzClippableButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim noClip As WzRect
    WzClippableButton_displayImpl self, xOffset, yOffset, False, noClip
End Sub

' void displayImpl(int xOffset, int yOffset, optional<WzRect> clipRectScreenCoords = nullopt)
Private Sub WzClippableButton_displayImpl(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long, ByVal has_clipRectScreenCoords As Boolean, clipRectScreenCoords As WzRect)
    Dim d As WzClippableButton, state As Long, x0 As Long, y0 As Long, w As Long, h As Long
    Dim isDown As Boolean, isLocked As Boolean, isDisabled As Boolean, isHighlight As Boolean, backgroundColor As Long
    Dim backX0 As Long, backY0 As Long, backX1 As Long, backY1 As Long, a As Long, textColor As Long
    Dim textWidth As Long, textX0 As Long, textY0 As Long
    Set d = WzClippableButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()

    state = W_BUTTON_(self).state
    isDown = (state And WBUT_DOWN) <> 0
    isLocked = (state And (WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (state And WBUT_DISABLE) <> 0
    isHighlight = (state And WBUT_HIGHLIGHT) <> 0

    ' Draw the background (clipped if needed)
    If isDown Or isLocked Or isHighlight Then
        If isLocked Then
            backgroundColor = psPalette(WZCOL_MENU_SCORE_BUILT)
        ElseIf isDown Then
            backgroundColor = psPalette(WZCOL_FORM_DARK)
        Else
            backgroundColor = psPalette(WZCOL_MENU_SCORE_BUILT)
            backgroundColor = pielightWithA(backgroundColor, pielightA(backgroundColor) \ 2)
        End If
        If has_clipRectScreenCoords Then a = clipRectScreenCoords.x Else a = 0
        backX0 = MAX_(x0, a)
        If has_clipRectScreenCoords Then a = clipRectScreenCoords.y Else a = 0
        backY0 = MAX_(y0, a)
        If has_clipRectScreenCoords Then a = WzRect_right(clipRectScreenCoords) Else a = x0 + w
        backX1 = MIN_(x0 + w, a)
        If has_clipRectScreenCoords Then a = WzRect_bottom(clipRectScreenCoords) Else a = y0 + h
        backY1 = MIN_(y0 + h, a)
        pie_UniTransBoxFill backX0, backY0, backX1, backY1, backgroundColor
    End If

    ' Draw the main text
    If Not isDisabled And (isLocked Or isHighlight) Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isDisabled Then
        textColor = pielightWithA(textColor, pielightA(textColor) \ 2)
    End If

    textWidth = d.wzText.deref().width()
    textX0 = x0 + (w - textWidth) \ 2
    textY0 = y0 + (h - d.wzText.deref().lineSize()) \ 2 - d.wzText.deref().aboveBase()

    If has_clipRectScreenCoords Then
        d.wzText.deref().renderClipped textX0, textY0, textColor, clipRectScreenCoords.x, clipRectScreenCoords.y, clipRectScreenCoords.w, clipRectScreenCoords.h
    Else
        d.wzText.deref().render textX0, textY0, textColor
    End If
End Sub

Public Sub WzClippableButton_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim clipRectScreenCoords As WzRect, geom As WzRect, translated As WidgetGraphicsContext, childrenContext As WidgetGraphicsContext
    Dim i As Long, child As WIDGET
    geom = self.geometry()
    If WidgetGraphicsContext_clipIntersects(context, geom, clipRectScreenCoords) Then
        WzClippableButton_displayImpl self, WidgetGraphicsContext_getXOffset(context), WidgetGraphicsContext_getYOffset(context), True, clipRectScreenCoords
    Else
        If Not WidgetGraphicsContext_allowChildDisplayRecursiveIfSelfClipped(context) Then
            Exit Sub
        End If
    End If

    translated = WidgetGraphicsContext_translatedBy(context, self.x(), self.y())
    childrenContext = WidgetGraphicsContext_setAllowChildDisplayRecursiveIfSelfClipped(translated, False)

    For i = 0 To self.childCount() - 1
        Set child = self.childAt(i)
        If child.visible() Then
            child.displayRecursive childrenContext
        End If
    Next
End Sub

Public Function WzClippableButton_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzClippableButton
    Set d = WzClippableButton_(self)
    WzClippableButton_idealWidth = (d.horizontalPadding * 2) + d.wzText.getTextWidth()
End Function

Public Function WzClippableButton_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzClippableButton
    Set d = WzClippableButton_(self)
    WzClippableButton_idealHeight = (d.verticalPadding * 2) + d.wzText.getTextLineSize()
End Function

' MARK: - HorizontalScrollButton

Public Sub HorizontalScrollButton_construct(ByVal self As WIDGET)
    Dim d As HorizontalScrollButton
    W_BUTTON_construct self
    self.kind = K_HorizontalScrollButton
    Set d = HorizontalScrollButton_(self)
    Set d.wzText = New WzText
    d.horizontalPadding = 5
    d.verticalPadding = 5
End Sub

Public Function HorizontalScrollButton_make(ByVal c As String, ByVal buttonFontID As iV_fonts) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_HorizontalScrollButton)
    HorizontalScrollButton_construct widget
    HorizontalScrollButton_initialize widget, c, buttonFontID
    Set HorizontalScrollButton_make = widget
End Function

Private Sub HorizontalScrollButton_initialize(ByVal self As WIDGET, ByVal c As String, ByVal buttonFontID As iV_fonts)
    HorizontalScrollButton_(self).wzText.setText c, buttonFontID
End Sub

Public Function HorizontalScrollButton_idealWidth(ByVal self As WIDGET) As Long
    Dim d As HorizontalScrollButton
    Set d = HorizontalScrollButton_(self)
    HorizontalScrollButton_idealWidth = (d.horizontalPadding * 2) + d.wzText.width()
End Function

Public Function HorizontalScrollButton_idealHeight(ByVal self As WIDGET) As Long
    Dim d As HorizontalScrollButton
    Set d = HorizontalScrollButton_(self)
    HorizontalScrollButton_idealHeight = (d.verticalPadding * 2) + d.wzText.lineSize()
End Function

Public Sub HorizontalScrollButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As HorizontalScrollButton, state As Long, x0 As Long, y0 As Long, w As Long, h As Long
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean
    Dim backX0 As Long, backY0 As Long, backX1 As Long, backY1 As Long, textColor As Long
    Dim textWidth As Long, textX0 As Long, textY0 As Long
    Set d = HorizontalScrollButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()

    state = W_BUTTON_(self).state
    isDown = (state And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (state And WBUT_DISABLE) <> 0
    isHighlight = (state And WBUT_HIGHLIGHT) <> 0

    backX0 = x0
    backY0 = y0
    backX1 = x0 + w
    backY1 = y0 + h

    ' Draw the background
    If isDisabled Then
        ' no background
    ElseIf isDown Then
        pie_UniTransBoxFill backX0, backY0, backX1, backY1, pal_RGBA(255, 255, 255, 200)
        pie_UniTransBoxFill backX0 + 1, backY0 + 1, backX1 - 1, backY1 - 1, psPalette(WZCOL_FORM_DARK)
    ElseIf isHighlight Then
        pie_UniTransBoxFill backX0, backY0, backX1, backY1, psPalette(WZCOL_MENU_SCORE_BUILT)
    Else
        pie_UniTransBoxFill backX0, backY0, backX1, backY1, psPalette(WZCOL_TRANSPARENT_BOX)
        pie_UniTransBoxFill backX0, backY0, backX1, backY1, psPalette(WZCOL_TRANSPARENT_BOX)
    End If

    ' Draw the main text
    If Not isDisabled And isHighlight Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isDisabled Then
        textColor = pielightWithA(textColor, pielightA(textColor) \ 3)
    End If

    textWidth = d.wzText.width()
    textX0 = x0 + (w - textWidth) \ 2
    textY0 = y0 + (h - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()
    d.wzText.render textX0, textY0, textColor
End Sub

' MARK: - HorizontalButtonsWidget

Public Sub HorizontalButtonsWidget_construct(ByVal self As WIDGET)
    Dim d As HorizontalButtonsWidget
    WIDGET_construct self
    self.kind = K_HorizontalButtonsWidget
    Set d = HorizontalButtonsWidget_(self)
    d.weakSerial = WeakWidget_register(self)
    d.clickLock = False
    d.horizontalButtonPadding = 15
    d.verticalButtonPadding = 10
    d.verticalOuterPadding = 0
    d.has_backgroundColor = False       ' = nullopt
    Set d.optionsFormUntranslatedTitles = New Collection
    Set d.optionsTitles = New Collection
    d.has_currentButtonIdx = False      ' = nullopt
    d.scrollableWidth = 0
    d.maxItemHeight = 0
    d.itemSpacing = 0
    d.clipContainerInnerPaddingX = 28
    d.currentXScrollPos = 0
    d.has_dragStartXScrollPos = False   ' = nullopt
    d.isHandlingDrag = False
    d.layoutDirty = True
End Sub

Public Sub HorizontalButtonsWidget_destruct(ByVal self As WIDGET)
    WeakWidget_unregister HorizontalButtonsWidget_(self).weakSerial
End Sub

' static std::shared_ptr<HorizontalButtonsWidget> make(const OnSelectForm& onSelectForm, iV_fonts buttonFontID, bool clickLock = true, int32_t horizontalButtonPadding = 15, int32_t verticalButtonPadding = 10)
Public Function HorizontalButtonsWidget_make(ByVal onSelectForm As WzFunc, ByVal buttonFontID As iV_fonts, Optional ByVal clickLock As Boolean = True, Optional ByVal horizontalButtonPadding As Long = 15, Optional ByVal verticalButtonPadding As Long = 10) As WIDGET
    Dim widget As WIDGET
    If Not (Not onSelectForm Is Nothing) Then ASSERT_FAILURE "onSelectForm != nullptr", "src/titleui/options/optionsforms.cpp:339", "make", "Missing expected onSelectForm function": Exit Function
    Set widget = WIDGET_alloc(K_HorizontalButtonsWidget)
    HorizontalButtonsWidget_construct widget
    HorizontalButtonsWidget_initialize widget, onSelectForm, buttonFontID, clickLock, horizontalButtonPadding, verticalButtonPadding
    Set HorizontalButtonsWidget_make = widget
End Function

Private Sub HorizontalButtonsWidget_initialize(ByVal self As WIDGET, ByVal onSelectForm_ As WzFunc, ByVal buttonFontID_ As iV_fonts, ByVal clickLock_ As Boolean, ByVal horizontalButtonPadding_ As Long, ByVal verticalButtonPadding_ As Long)
    Dim d As HorizontalButtonsWidget, weakSelf As Long
    Set d = HorizontalButtonsWidget_(self)
    Set d.onSelectForm = onSelectForm_
    d.buttonFontID = buttonFontID_
    d.clickLock = clickLock_
    d.horizontalButtonPadding = horizontalButtonPadding_
    d.verticalButtonPadding = verticalButtonPadding_

    Set d.clipContainer = ClipRectWidget_new()
    self.attach d.clipContainer

    Set d.scrollLeftButton = HorizontalScrollButton_make(ChrW$(&HAB), font_medium_bold)    ' "<<" (U+00AB)
    self.attach d.scrollLeftButton, ChildZPos_Front
    d.scrollLeftButton.hide

    Set d.scrollRightButton = HorizontalScrollButton_make(ChrW$(&HBB), font_medium_bold)   ' ">>" (U+00BB)
    self.attach d.scrollRightButton, ChildZPos_Front
    d.scrollRightButton.hide

    weakSelf = d.weakSerial
    W_BUTTON_addOnClickHandler d.scrollLeftButton, WzFunc_new(FN_HBW_scrollLeftClick, weakSelf)
    W_BUTTON_addOnClickHandler d.scrollRightButton, WzFunc_new(FN_HBW_scrollRightClick, weakSelf)
End Sub

Private Function HorizontalButtonsWidget_currentLeftMostFullyVisibleButtonIdx(ByVal self As WIDGET) As Long
    Dim d As HorizontalButtonsWidget, i As Long, buttonXPos As Long, minRequiredVisibleXPos As Long
    Set d = HorizontalButtonsWidget_(self)
    For i = 0 To d.optionsTitles.Count - 1
        buttonXPos = d.optionsTitles(i + 1).x()
        minRequiredVisibleXPos = buttonXPos - d.clipContainerInnerPaddingX
        If d.currentXScrollPos <= minRequiredVisibleXPos Then
            HorizontalButtonsWidget_currentLeftMostFullyVisibleButtonIdx = i
            Exit Function
        End If
    Next
    HorizontalButtonsWidget_currentLeftMostFullyVisibleButtonIdx = d.optionsTitles.Count
End Function

Private Function HorizontalButtonsWidget_currentRightMostFullyVisibleButtonIdx(ByVal self As WIDGET) As Long
    Dim d As HorizontalButtonsWidget, i As Long, buttonXPos As Long, buttonWidth As Long, maxXPosVisible As Long, maxDesiredXPosVisible As Long
    Set d = HorizontalButtonsWidget_(self)
    If d.optionsTitles.Count = 0 Then
        HorizontalButtonsWidget_currentRightMostFullyVisibleButtonIdx = 0
        Exit Function
    End If
    For i = d.optionsTitles.Count - 1 To 1 Step -1
        buttonXPos = d.optionsTitles(i + 1).x()
        buttonWidth = d.optionsTitles(i + 1).width()
        maxXPosVisible = d.currentXScrollPos + self.width()
        maxDesiredXPosVisible = (buttonXPos + buttonWidth + d.clipContainerInnerPaddingX)
        If maxXPosVisible >= maxDesiredXPosVisible Then
            HorizontalButtonsWidget_currentRightMostFullyVisibleButtonIdx = i
            Exit Function
        End If
    Next
    HorizontalButtonsWidget_currentRightMostFullyVisibleButtonIdx = 0
End Function

Public Sub HorizontalButtonsWidget_addButton(ByVal self As WIDGET, ByVal untranslatedTitle As String)
    Dim d As HorizontalButtonsWidget, newButtonIdx As Long, but As WIDGET, weakSelf As Long
    Set d = HorizontalButtonsWidget_(self)
    newButtonIdx = d.optionsFormUntranslatedTitles.Count

    d.optionsFormUntranslatedTitles.Add untranslatedTitle

    Set but = WzClippableButton_new()
    WzClippableButton_setPadding but, d.horizontalButtonPadding, d.verticalButtonPadding
    WzClippableButton_setFont but, d.buttonFontID
    but.setString untranslatedTitle             ' gettext (translations not ported)

    weakSelf = d.weakSerial
    W_BUTTON_addOnClickHandler but, WzFunc_new(FN_HBW_buttonClick, weakSelf, newButtonIdx)

    d.clipContainer.attach but
    d.optionsTitles.Add but
    d.layoutDirty = True
End Sub

Public Sub HorizontalButtonsWidget_informLanguageDidChange(ByVal self As WIDGET)
    Dim d As HorizontalButtonsWidget, i As Long
    Set d = HorizontalButtonsWidget_(self)
    ' Must re-set all of the button text strings, and recalculate layout
    For i = 0 To d.optionsFormUntranslatedTitles.Count - 1
        d.optionsTitles(i + 1).setString d.optionsFormUntranslatedTitles(i + 1)
    Next
    d.layoutDirty = True
End Sub

Public Sub HorizontalButtonsWidget_selectItem(ByVal self As WIDGET, ByVal buttonIdx As Long)
    Dim d As HorizontalButtonsWidget, i As Long
    Set d = HorizontalButtonsWidget_(self)
    If Not (buttonIdx < d.optionsFormUntranslatedTitles.Count) Then ASSERT_FAILURE "buttonIdx < optionsFormUntranslatedTitles.size()", "src/titleui/options/optionsforms.cpp:457", "selectItem", "Invalid buttonIdx: %zu", buttonIdx: Exit Sub
    d.has_currentButtonIdx = True
    d.currentButtonIdx = buttonIdx

    WzFunc_call d.onSelectForm, CStr(d.optionsFormUntranslatedTitles(buttonIdx + 1))

    If d.clickLock Then
        For i = 0 To d.optionsTitles.Count - 1
            If i = buttonIdx Then d.optionsTitles(i + 1).setState WBUT_CLICKLOCK Else d.optionsTitles(i + 1).setState 0
        Next
    End If

    HorizontalButtonsWidget_scrollEnsureButtonVisible self, buttonIdx
End Sub

Public Sub HorizontalButtonsWidget_clear(ByVal self As WIDGET)
    Dim d As HorizontalButtonsWidget
    Set d = HorizontalButtonsWidget_(self)
    Set d.optionsFormUntranslatedTitles = New Collection
    d.clipContainer.removeAllChildren
    Set d.optionsTitles = New Collection
    d.has_currentButtonIdx = False
    Set d.lastMouseOverInternalWidget = Nothing
    d.currentXScrollPos = 0
    d.layoutDirty = True
End Sub

Public Sub HorizontalButtonsWidget_setItemSpacing(ByVal self As WIDGET, ByVal spacing As Long)
    Dim d As HorizontalButtonsWidget
    Set d = HorizontalButtonsWidget_(self)
    d.itemSpacing = spacing
    d.layoutDirty = True
End Sub

Public Sub HorizontalButtonsWidget_setVerticalOuterPadding(ByVal self As WIDGET, ByVal padding_ As Long)
    Dim d As HorizontalButtonsWidget
    Set d = HorizontalButtonsWidget_(self)
    d.verticalOuterPadding = padding_
    d.layoutDirty = True
End Sub

' void setBackgroundColor(optional<PIELIGHT> bckColor)
Public Sub HorizontalButtonsWidget_setBackgroundColor(ByVal self As WIDGET, ByVal has_color As Boolean, ByVal color As Long)
    Dim d As HorizontalButtonsWidget
    Set d = HorizontalButtonsWidget_(self)
    d.has_backgroundColor = has_color
    d.backgroundColor = color
End Sub

Private Sub HorizontalButtonsWidget_scrollEnsureButtonVisible(ByVal self As WIDGET, ByVal buttonIdx As Long)
    Dim d As HorizontalButtonsWidget, buttonXPos As Long, buttonWidth As Long, minRequiredVisibleXPos As Long
    Dim w As Long, maxXPosVisible As Long, maxDesiredXPosVisible As Long
    Set d = HorizontalButtonsWidget_(self)
    ' scroll if needed to ensure that the full button is visible
    buttonXPos = d.optionsTitles(buttonIdx + 1).x()
    buttonWidth = d.optionsTitles(buttonIdx + 1).width()
    minRequiredVisibleXPos = buttonXPos - d.clipContainerInnerPaddingX
    If d.currentXScrollPos > minRequiredVisibleXPos Then
        HorizontalButtonsWidget_setCurrentXScrollPos self, minRequiredVisibleXPos
    Else
        w = self.width()
        maxXPosVisible = d.currentXScrollPos + w - d.clipContainerInnerPaddingX
        maxDesiredXPosVisible = (buttonXPos + buttonWidth + d.clipContainerInnerPaddingX)
        If maxXPosVisible < maxDesiredXPosVisible Then
            HorizontalButtonsWidget_setCurrentXScrollPos self, d.currentXScrollPos + (maxDesiredXPosVisible - maxXPosVisible - d.clipContainerInnerPaddingX)
        End If
    End If
End Sub

Private Sub HorizontalButtonsWidget_resizeChildren(ByVal self As WIDGET)
    Dim d As HorizontalButtonsWidget, nextXOffset As Long, itemHeight As Long, i As Long, child As WIDGET
    Set d = HorizontalButtonsWidget_(self)
    d.scrollableWidth = 0
    d.maxItemHeight = 0
    nextXOffset = d.clipContainerInnerPaddingX
    itemHeight = self.height() - (d.verticalOuterPadding * 2)
    For i = 0 To d.clipContainer.childCount() - 1
        Set child = d.clipContainer.childAt(i)
        If Not d.clipContainer.visible() Then
            GoTo continue_
        End If
        child.setGeometry nextXOffset, 0, child.idealWidth(), itemHeight
        d.scrollableWidth = nextXOffset + child.idealWidth()
        d.maxItemHeight = MAX_(d.maxItemHeight, child.idealHeight())
        nextXOffset = d.scrollableWidth + d.itemSpacing
continue_:
    Next
    d.scrollableWidth = d.scrollableWidth + d.clipContainerInnerPaddingX
End Sub

Public Sub HorizontalButtonsWidget_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    HorizontalButtonsWidget_updateLayout self
    WIDGET_displayRecursive self, context
End Sub

Private Sub HorizontalButtonsWidget_updateLayout(ByVal self As WIDGET)
    Dim d As HorizontalButtonsWidget, h As Long, w As Long, clipContainerWidth As Long, clipContainerX0 As Long
    Dim clipContainerX1 As Long, clipContainerHeight As Long
    Set d = HorizontalButtonsWidget_(self)
    If Not d.layoutDirty Then
        Exit Sub
    End If
    d.layoutDirty = False
    HorizontalButtonsWidget_resizeChildren self

    h = self.height()
    w = self.width()

    clipContainerWidth = MIN_(w, d.scrollableWidth)
    clipContainerX0 = (w - clipContainerWidth) \ 2
    clipContainerX1 = clipContainerX0 + clipContainerWidth
    clipContainerHeight = h - (d.verticalOuterPadding * 2)
    d.clipContainer.setGeometry clipContainerX0, d.verticalOuterPadding, clipContainerWidth, clipContainerHeight

    d.scrollLeftButton.setGeometry clipContainerX0, 0, d.clipContainerInnerPaddingX, h
    d.scrollRightButton.setGeometry clipContainerX1 - d.clipContainerInnerPaddingX, 0, d.clipContainerInnerPaddingX, h

    If d.has_currentButtonIdx Then
        HorizontalButtonsWidget_scrollEnsureButtonVisible self, d.currentButtonIdx
    End If

    HorizontalButtonsWidget_updateScrollButtonsState self
End Sub

Public Sub HorizontalButtonsWidget_geometryChanged(ByVal self As WIDGET)
    HorizontalButtonsWidget_(self).layoutDirty = True
End Sub

Public Sub HorizontalButtonsWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As HorizontalButtonsWidget, x0 As Long, y0 As Long
    Set d = HorizontalButtonsWidget_(self)
    If d.has_backgroundColor Then
        x0 = self.x() + xOffset
        y0 = self.y() + yOffset

        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), d.backgroundColor
    End If
End Sub

Public Function HorizontalButtonsWidget_capturesMouseDrag(ByVal self As WIDGET, ByVal key As WIDGET_KEY) As Boolean
    HorizontalButtonsWidget_capturesMouseDrag = True
End Function

Public Sub HorizontalButtonsWidget_mouseDragged(ByVal self As WIDGET, ByVal wkey As WIDGET_KEY, psStartContext As W_CONTEXT, psContext As W_CONTEXT)
    Dim d As HorizontalButtonsWidget, deltaDragX As Long
    Set d = HorizontalButtonsWidget_(self)
    If wkey <> WKEY_PRIMARY Then
        Exit Sub
    End If

    If Not d.isHandlingDrag Then
        If Abs(psStartContext.mx - psContext.mx) > HorizontalButtonsWidgetDragThreshold Or Abs(psStartContext.my - psContext.my) > HorizontalButtonsWidgetDragThreshold Then
            ' start processing as drag
            d.isHandlingDrag = True
            d.has_dragStartXScrollPos = True
            d.dragStartXScrollPos = d.currentXScrollPos
        End If
    End If

    If d.isHandlingDrag Then
        deltaDragX = psContext.mx - psStartContext.mx
        HorizontalButtonsWidget_setCurrentXScrollPos self, d.dragStartXScrollPos - deltaDragX
    End If
End Sub

Public Sub HorizontalButtonsWidget_released(ByVal self As WIDGET, context As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As HorizontalButtonsWidget
    Set d = HorizontalButtonsWidget_(self)
    If Not d.isHandlingDrag Then
        ' Must forward this event to child buttons, as we aren't processing a drag (yet)
        ' but widgRunScreen has custom handling when a widget captures mouse drag
        ' (which this does) and won't ultimately call findMouseTargetRecursive, but instead
        ' just mouseDragged() and released() on the capture widget.
        Set d.lastMouseOverInternalWidget = HorizontalButtonsWidget_processInternalButtonOver(self, context, key, False)
    End If
    d.isHandlingDrag = False
    d.has_dragStartXScrollPos = False
End Sub

Private Sub HorizontalButtonsWidget_setCurrentXScrollPos(ByVal self As WIDGET, ByVal val_ As Long)
    Dim d As HorizontalButtonsWidget, maxScrollXPos As Long
    Set d = HorizontalButtonsWidget_(self)
    If d.scrollableWidth > self.width() Then maxScrollXPos = d.scrollableWidth - self.width() Else maxScrollXPos = 0
    If maxScrollXPos > 0 Then
        d.currentXScrollPos = clip_(val_, 0, maxScrollXPos)
    Else
        d.currentXScrollPos = 0
    End If
    ClipRectWidget_setLeftOffset d.clipContainer, d.currentXScrollPos

    ' Update scroll side buttons state
    HorizontalButtonsWidget_updateScrollButtonsState self
End Sub

Private Sub HorizontalButtonsWidget_updateScrollButtonsState(ByVal self As WIDGET)
    Dim d As HorizontalButtonsWidget, maxScrollXPos As Long, leftScrollButtonDisabled As Boolean, rightScrollButtonDisabled As Boolean
    Dim buttonsVisible As Boolean
    Set d = HorizontalButtonsWidget_(self)
    If d.scrollableWidth > self.width() Then maxScrollXPos = d.scrollableWidth - self.width() Else maxScrollXPos = 0

    leftScrollButtonDisabled = (d.currentXScrollPos <= 0)
    rightScrollButtonDisabled = (d.currentXScrollPos >= maxScrollXPos)

    If leftScrollButtonDisabled Then
        d.scrollLeftButton.setState WBUT_DISABLE
    ElseIf (d.scrollLeftButton.getState() And WBUT_DISABLE) <> 0 Then
        d.scrollLeftButton.setState 0
    End If

    If rightScrollButtonDisabled Then
        d.scrollRightButton.setState WBUT_DISABLE
    ElseIf (d.scrollRightButton.getState() And WBUT_DISABLE) <> 0 Then
        d.scrollRightButton.setState 0
    End If

    buttonsVisible = (Not leftScrollButtonDisabled Or Not rightScrollButtonDisabled)
    d.scrollLeftButton.show buttonsVisible
    d.scrollRightButton.show buttonsVisible
End Sub

Public Function HorizontalButtonsWidget_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim d As HorizontalButtonsWidget, wheelX As Long, wheelY As Long
    Set d = HorizontalButtonsWidget_(self)
    getMouseWheelSpeed wheelX, wheelY
    HorizontalButtonsWidget_setCurrentXScrollPos self, d.currentXScrollPos + (wheelX * 10)

    ' skip processing of children - just process clicks for ourself

    If Not self.visible() Or self.transparentToMouse() Then
        Set HorizontalButtonsWidget_findMouseTargetRecursive = Nothing
        Exit Function
    End If

    ' however, forward certain events (highlight, click) to buttons if not dragging
    If Not d.isHandlingDrag Then
        Set d.lastMouseOverInternalWidget = HorizontalButtonsWidget_processInternalButtonOver(self, psContext, key, wasPressed)
    Else
        If Not d.lastMouseOverInternalWidget Is Nothing Then
            d.lastMouseOverInternalWidget.highlightLost
        End If
        Set d.lastMouseOverInternalWidget = Nothing
    End If

    Set HorizontalButtonsWidget_findMouseTargetRecursive = self
End Function

Public Sub HorizontalButtonsWidget_highlightLost(ByVal self As WIDGET)
    Dim d As HorizontalButtonsWidget
    Set d = HorizontalButtonsWidget_(self)
    If Not d.lastMouseOverInternalWidget Is Nothing Then
        If Not d.lastMouseOverInternalWidget Is self Then
            d.lastMouseOverInternalWidget.highlightLost
        End If
        Set d.lastMouseOverInternalWidget = Nothing
    End If
End Sub

' (C++: psContext is never nullptr here; the assert is kept as a comment)
Private Function HorizontalButtonsWidget_processInternalButtonOver(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim d As HorizontalButtonsWidget, mouseOverWidgetCtx As W_CONTEXT, mouseOverWidget As WIDGET
    Set d = HorizontalButtonsWidget_(self)
    ' ASSERT_OR_RETURN(nullptr, psContext != nullptr, "Null context?");

    mouseOverWidgetCtx = psContext
    Set mouseOverWidget = WIDGET_findMouseTargetRecursive(self, mouseOverWidgetCtx, key, wasPressed)

    If mouseOverWidget Is Nothing Then
        If Not d.lastMouseOverInternalWidget Is Nothing Then
            d.lastMouseOverInternalWidget.highlightLost
        End If
        Exit Function                       ' nullptr
    End If
    If mouseOverWidget Is self Then
        If Not d.lastMouseOverInternalWidget Is Nothing And Not d.lastMouseOverInternalWidget Is mouseOverWidget Then
            d.lastMouseOverInternalWidget.highlightLost
        End If
        Exit Function
    End If

    If Not mouseOverWidget Is d.lastMouseOverInternalWidget Then
        If Not d.lastMouseOverInternalWidget Is Nothing Then
            d.lastMouseOverInternalWidget.highlightLost
        End If
        mouseOverWidget.highlight mouseOverWidgetCtx
    End If

    If key <> WKEY_NONE Then
        If wasPressed Then
            mouseOverWidget.clicked mouseOverWidgetCtx, key
        Else
            mouseOverWidget.released mouseOverWidgetCtx, key
        End If
    End If
    Set HorizontalButtonsWidget_processInternalButtonOver = mouseOverWidget
End Function

Public Function HorizontalButtonsWidget_idealWidth(ByVal self As WIDGET) As Long
    HorizontalButtonsWidget_updateLayout self
    HorizontalButtonsWidget_idealWidth = HorizontalButtonsWidget_(self).scrollableWidth
End Function

Public Function HorizontalButtonsWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As HorizontalButtonsWidget
    HorizontalButtonsWidget_updateLayout self
    Set d = HorizontalButtonsWidget_(self)
    HorizontalButtonsWidget_idealHeight = d.maxItemHeight + (d.verticalOuterPadding * 2)
End Function

' MARK: - OptionsBrowserForm

Public Sub OptionsBrowserForm_construct(ByVal self As WIDGET)
    Dim d As OptionsBrowserForm
    WIDGET_construct self
    self.kind = K_OptionsBrowserForm
    Set d = OptionsBrowserForm_(self)
    d.weakSerial = WeakWidget_register(self)
    d.optionsFormsCount = 0
    d.has_currentOptionsForm = False    ' = nullopt
    d.maxOptionsFormWidth = 450
End Sub

Public Sub OptionsBrowserForm_destruct(ByVal self As WIDGET)
    WeakWidget_unregister OptionsBrowserForm_(self).weakSerial
End Sub

Private Sub OptionsBrowserForm_initialize(ByVal self As WIDGET, ByVal optionalBackButton_ As WIDGET)
    Dim d As OptionsBrowserForm, weakSelf As Long, formSwitcher As WIDGET, sectionSwitcher As WIDGET
    Dim openConfigurationDirectoryLinkTmp As WIDGET
    Set d = OptionsBrowserForm_(self)
    Set d.optionalBackButton = optionalBackButton_
    If Not d.optionalBackButton Is Nothing Then
        self.attach d.optionalBackButton
    End If

    weakSelf = d.weakSerial
    Set formSwitcher = HorizontalButtonsWidget_make(WzFunc_new(FN_OBF_formSwitcherSelect, weakSelf), font_medium_bold, True, 12, 7)
    HorizontalButtonsWidget_setItemSpacing formSwitcher, 3
    HorizontalButtonsWidget_setVerticalOuterPadding formSwitcher, 7
    self.attach formSwitcher
    Set d.optionsFormSwitcher = formSwitcher

    Set sectionSwitcher = HorizontalButtonsWidget_make(WzFunc_new(FN_OBF_sectionSwitcherSelect, weakSelf), font_regular, False, 9, 5)
    HorizontalButtonsWidget_setItemSpacing sectionSwitcher, 1
    HorizontalButtonsWidget_setVerticalOuterPadding sectionSwitcher, 6
    HorizontalButtonsWidget_setBackgroundColor sectionSwitcher, True, pal_RGBA(0, 0, 0, 65)
    self.attach sectionSwitcher
    Set d.optionsFormSectionSwitcher = sectionSwitcher

    Set openConfigurationDirectoryLinkTmp = WzClippableButton_new()
    WzClippableButton_setFont openConfigurationDirectoryLinkTmp, font_small
    openConfigurationDirectoryLinkTmp.setString "Open Configuration Directory"
    WzClippableButton_setPadding openConfigurationDirectoryLinkTmp, 6, 6
    self.attach openConfigurationDirectoryLinkTmp
    openConfigurationDirectoryLinkTmp.setGeometry 0, 0, openConfigurationDirectoryLinkTmp.width(), iV_GetTextLineSize(W_BUTTON_(openConfigurationDirectoryLinkTmp).FontID) + 8
    openConfigurationDirectoryLinkTmp.setCalcLayout WzFunc_new(FN_OBF_configDirLinkCalcLayout)
    W_BUTTON_addOnClickHandler openConfigurationDirectoryLinkTmp, WzFunc_new(FN_OBF_configDirLinkClick)
    Set d.openConfigurationDirectoryLink = openConfigurationDirectoryLinkTmp
End Sub

Public Sub OptionsBrowserForm_showOpenConfigDirLink(ByVal self As WIDGET, ByVal show As Boolean)
    Dim d As OptionsBrowserForm
    Set d = OptionsBrowserForm_(self)
    If Not d.openConfigurationDirectoryLink Is Nothing Then
        d.openConfigurationDirectoryLink.show show
    End If
End Sub

' static std::shared_ptr<OptionsBrowserForm> make(const std::shared_ptr<WIDGET>& optionalBackButton = nullptr)
Public Function OptionsBrowserForm_make(Optional ByVal optionalBackButton As WIDGET = Nothing) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_OptionsBrowserForm)
    OptionsBrowserForm_construct widget
    OptionsBrowserForm_initialize widget, optionalBackButton
    Set OptionsBrowserForm_make = widget
End Function

Public Sub OptionsBrowserForm_addOptionsForm(ByVal self As WIDGET, ByVal mode As OptionsBrowserForm_Modes, ByVal createFormFunc As WzFunc, ByVal untranslatedTitle As String)
    Dim d As OptionsBrowserForm, switcher As WIDGET, details As OptionsBrowserForm_OptionsFormDetails
    Set d = OptionsBrowserForm_(self)
    Set switcher = d.optionsFormSwitcher
    HorizontalButtonsWidget_addButton switcher, untranslatedTitle
    details.mode = mode
    details.untranslatedTitle = untranslatedTitle
    Set details.createFormFunc = createFormFunc
    Set details.optionsForm = Nothing
    d.optionsForms_push_back details
    If Not d.has_currentOptionsForm Then
        HorizontalButtonsWidget_selectItem switcher, 0
    End If
End Sub

Private Sub OptionsBrowserForm_displayOptionsForm(ByVal self As WIDGET, ByVal idx As Long)
    Dim d As OptionsBrowserForm, oldForm As WIDGET, alreadyHadForm As Boolean, details As OptionsBrowserForm_OptionsFormDetails
    Set d = OptionsBrowserForm_(self)
    If Not (idx < d.optionsFormsCount) Then ASSERT_FAILURE "idx < optionsForms.size()", "src/titleui/options/optionsforms.cpp:876", "displayOptionsForm", "Invalid index: %zu", idx: Exit Sub
    If d.has_currentOptionsForm Then
        Set oldForm = d.optionsForms(d.currentOptionsForm).optionsForm
        If Not oldForm Is Nothing Then
            self.detach oldForm
        End If
    End If

    d.has_currentOptionsForm = True
    d.currentOptionsForm = idx
    alreadyHadForm = True
    details = d.optionsForms(idx)
    If details.optionsForm Is Nothing Then
        alreadyHadForm = False
        Set details.optionsForm = WzFunc_call(details.createFormFunc)
        d.optionsForms(idx) = details
        OptionsForm_setMaxListWidth details.optionsForm, d.maxOptionsFormWidth
    End If
    self.attach details.optionsForm
    If alreadyHadForm Then
        ' since the display scale may have changed (due to edits on the one option screen *or* various other conditions)
        ' and non-attached options forms would not have received the screenSizeDidChange event, trigger it manually now
        details.optionsForm.screenSizeDidChange screenWidth, screenHeight, screenWidth, screenHeight

        ' trigger a refresh + row layout recalc
        OptionsForm_refreshOptions details.optionsForm, True
    End If

    OptionsBrowserForm_updateSectionSwitcher self

    OptionsBrowserForm_positionCurrentOptionsForm self
End Sub

Private Sub OptionsBrowserForm_updateSectionSwitcher(ByVal self As WIDGET)
    Dim d As OptionsBrowserForm, sectionSwitcher As WIDGET, optionsFormSections() As OptionsSection, n As Long, i As Long
    Set d = OptionsBrowserForm_(self)
    Set sectionSwitcher = d.optionsFormSectionSwitcher

    If Not d.has_currentOptionsForm Then
        HorizontalButtonsWidget_clear sectionSwitcher
        sectionSwitcher.hide
        Exit Sub
    End If

    HorizontalButtonsWidget_clear sectionSwitcher
    n = OptionsForm_getSectionInfos(d.optionsForms(d.currentOptionsForm).optionsForm, optionsFormSections)
    If n > 0 Then
        For i = 0 To n - 1
            HorizontalButtonsWidget_addButton sectionSwitcher, OptionsSection_sectionId(optionsFormSections(i))
        Next
        sectionSwitcher.show
    Else
        sectionSwitcher.hide
    End If
End Sub

Private Function OptionsBrowserForm_displayOptionsForm_title(ByVal self As WIDGET, ByVal untranslatedTitle As String) As Boolean
    Dim d As OptionsBrowserForm, i As Long
    Set d = OptionsBrowserForm_(self)
    For i = 0 To d.optionsFormsCount - 1
        If StrComp(d.optionsForms(i).untranslatedTitle, untranslatedTitle, vbBinaryCompare) = 0 Then
            OptionsBrowserForm_displayOptionsForm self, i
            OptionsBrowserForm_displayOptionsForm_title = True
            Exit Function
        End If
    Next
    OptionsBrowserForm_displayOptionsForm_title = False
End Function

Private Function OptionsBrowserForm_displayOptionsFormSection(ByVal self As WIDGET, ByVal sectionId As String) As Boolean
    Dim d As OptionsBrowserForm
    Set d = OptionsBrowserForm_(self)
    If Not d.has_currentOptionsForm Then
        OptionsBrowserForm_displayOptionsFormSection = False
        Exit Function
    End If
    OptionsBrowserForm_displayOptionsFormSection = OptionsForm_jumpToSectionId(d.optionsForms(d.currentOptionsForm).optionsForm, sectionId)
End Function

Private Sub OptionsBrowserForm_positionCurrentOptionsForm(ByVal self As WIDGET)
    Dim d As OptionsBrowserForm, w As Long, h As Long, bottomOfLastWidget As Long, switcherWidth As Long, switcherX0 As Long
    Dim formY0 As Long, formAvailableHeight As Long, current As WIDGET, formWidth As Long, formX0 As Long
    Set d = OptionsBrowserForm_(self)
    w = self.width()
    h = self.height()

    bottomOfLastWidget = d.optionsFormSwitcher.y() + d.optionsFormSwitcher.height()

    If d.optionsFormSectionSwitcher.visible() Then
        switcherWidth = w
        switcherX0 = 0
        d.optionsFormSectionSwitcher.setGeometry switcherX0, bottomOfLastWidget, switcherWidth, d.optionsFormSectionSwitcher.idealHeight()
        bottomOfLastWidget = d.optionsFormSectionSwitcher.y() + d.optionsFormSectionSwitcher.height()
    End If

    formY0 = bottomOfLastWidget + OptionsBrowserForm_paddingBeforeOptionsForm
    formAvailableHeight = h
    If Not d.openConfigurationDirectoryLink Is Nothing Then
        If d.openConfigurationDirectoryLink.visible() Then
            formAvailableHeight = d.openConfigurationDirectoryLink.y()
        End If
    End If
    formAvailableHeight = formAvailableHeight - (formY0 + OptionsBrowserForm_bottomPadding)

    If d.has_currentOptionsForm Then
        Set current = d.optionsForms(d.currentOptionsForm).optionsForm
        If Not (Not current Is Nothing) Then ASSERT_FAILURE "current != nullptr", "src/titleui/options/optionsforms.cpp:986", "positionCurrentOptionsForm", "Options form not instantiated?": Exit Sub
        formWidth = w - 10
        formX0 = (w - formWidth) \ 2
        current.setGeometry formX0, formY0, formWidth, formAvailableHeight
    End If
End Sub

Public Sub OptionsBrowserForm_geometryChanged(ByVal self As WIDGET)
    Dim d As OptionsBrowserForm, w As Long, h As Long, availableFormSwitcherWidth As Long, minSwitcherX0 As Long
    Dim switcherWidth As Long, switcherX0 As Long, configDirLinkHeight As Long, configDirLinkY0 As Long
    Set d = OptionsBrowserForm_(self)
    w = self.width()
    h = self.height()

    availableFormSwitcherWidth = w
    minSwitcherX0 = 0
    If Not d.optionalBackButton Is Nothing Then
        ' position to left of the form switcher
        d.optionalBackButton.setGeometry 0, 0, d.optionalBackButton.idealWidth(), d.optionalBackButton.idealHeight()
        minSwitcherX0 = d.optionalBackButton.x() + d.optionalBackButton.width()
        availableFormSwitcherWidth = availableFormSwitcherWidth - minSwitcherX0
    End If

    switcherWidth = MIN_(d.optionsFormSwitcher.idealWidth(), availableFormSwitcherWidth)
    switcherX0 = MAX_((availableFormSwitcherWidth - switcherWidth) \ 2, minSwitcherX0)
    d.optionsFormSwitcher.setGeometry switcherX0, OptionsBrowserForm_topPadding, switcherWidth, d.optionsFormSwitcher.idealHeight()

    If Not d.openConfigurationDirectoryLink Is Nothing Then
        configDirLinkHeight = d.openConfigurationDirectoryLink.idealHeight()
        configDirLinkY0 = h - configDirLinkHeight - 1
        d.openConfigurationDirectoryLink.setGeometry OptionsBrowserForm_leftOpenConfigDirLinkPadding, configDirLinkY0, d.openConfigurationDirectoryLink.idealWidth(), configDirLinkHeight
    End If

    OptionsBrowserForm_positionCurrentOptionsForm self
End Sub

Public Sub OptionsBrowserForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As OptionsBrowserForm, x0 As Long, y0 As Long, h As Long, switcherY0 As Long, switcherY1 As Long
    Dim configDirLinkY0 As Long, configDirLinkY1 As Long
    Set d = OptionsBrowserForm_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    h = self.height()

    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + h, pal_RGBA(150, 150, 150, 50)
    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + h, pal_RGBA(0, 0, 0, 70)
    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + h, psPalette(WZCOL_TRANSPARENT_BOX)

    ' Draw a darker border underneath where the switcher is
    switcherY0 = d.optionsFormSwitcher.y()
    switcherY1 = switcherY0 + d.optionsFormSwitcher.height()
    pie_UniTransBoxFill x0, y0 + switcherY0, x0 + self.width(), y0 + switcherY1, psPalette(WZCOL_TRANSPARENT_BOX)
    pie_UniTransBoxFill x0, y0 + switcherY0, x0 + self.width(), y0 + switcherY1, pal_RGBA(0, 0, 0, 128)

    If Not d.openConfigurationDirectoryLink Is Nothing Then
        If d.openConfigurationDirectoryLink.visible() Then
            configDirLinkY0 = d.openConfigurationDirectoryLink.y()
            configDirLinkY1 = h
            pie_UniTransBoxFill x0, y0 + configDirLinkY0, x0 + self.width(), y0 + configDirLinkY1, psPalette(WZCOL_TRANSPARENT_BOX)
            pie_UniTransBoxFill x0, y0 + configDirLinkY0, x0 + self.width(), y0 + configDirLinkY1, pal_RGBA(0, 0, 0, 128)
        End If
    End If
End Sub

Public Function OptionsBrowserForm_idealWidth(ByVal self As WIDGET) As Long
    OptionsBrowserForm_idealWidth = MAX_(OptionsBrowserForm_(self).maxOptionsFormWidth, MIN_IDEAL_OPTIONS_BROWSER_FORM_WIDTH)
End Function

Public Function OptionsBrowserForm_idealHeight(ByVal self As WIDGET) As Long
    Dim minCalculatedIdealHeight As Long
    minCalculatedIdealHeight = (OptionsBrowserForm_topPadding + OptionsBrowserForm_bottomPadding) + OptionsBrowserForm_(self).optionsFormSwitcher.idealHeight() + OptionsBrowserForm_paddingBeforeOptionsForm + 300
    OptionsBrowserForm_idealHeight = MAX_(minCalculatedIdealHeight, MIN_IDEAL_OPTIONS_BROWSER_FORM_HEIGHT)
End Function

Public Function OptionsBrowserForm_switchToOptionsForm(ByVal self As WIDGET, ByVal mode As OptionsBrowserForm_Modes) As Boolean
    Dim d As OptionsBrowserForm, i As Long
    Set d = OptionsBrowserForm_(self)
    For i = 0 To d.optionsFormsCount - 1
        If d.optionsForms(i).mode = mode Then
            HorizontalButtonsWidget_selectItem d.optionsFormSwitcher, i
            OptionsBrowserForm_switchToOptionsForm = True
            Exit Function
        End If
    Next
    OptionsBrowserForm_switchToOptionsForm = False
End Function

Public Sub OptionsBrowserForm_informLanguageDidChange(ByVal self As WIDGET)
    Dim d As OptionsBrowserForm
    Set d = OptionsBrowserForm_(self)
    ' Refresh switchers
    HorizontalButtonsWidget_informLanguageDidChange d.optionsFormSwitcher
    HorizontalButtonsWidget_informLanguageDidChange d.optionsFormSectionSwitcher

    ' Refresh open configuration directory link text
    d.openConfigurationDirectoryLink.setString "Open Configuration Directory"

    ' Trigger layout recalc
    self.geometryChanged
End Sub

' MARK: -

' std::shared_ptr<OptionsBrowserForm> createOptionsBrowser(bool inGame, const std::shared_ptr<WIDGET>& optionalBackButton = nullptr)
Public Function createOptionsBrowser(ByVal inGame_ As Boolean, Optional ByVal optionalBackButton As WIDGET = Nothing) As WIDGET
    Dim result As WIDGET, weakOptionsBrowserForm As Long, informOnLanguageChangeHandler As WzFunc
    Set result = OptionsBrowserForm_make(optionalBackButton)
    OptionsBrowserForm_showOpenConfigDirLink result, Not inGame_

    weakOptionsBrowserForm = OptionsBrowserForm_(result).weakSerial
    Set informOnLanguageChangeHandler = WzFunc_new(FN_createOptionsBrowser_informLanguage, weakOptionsBrowserForm)

    OptionsBrowserForm_addOptionsForm result, OptionsBrowserForm_Modes_Interface, WzFunc_new(FN_createOptionsBrowser_interface, inGame_, informOnLanguageChangeHandler), "Interface"
    If Not inGame_ Then
        OptionsBrowserForm_addOptionsForm result, OptionsBrowserForm_Modes_Defaults, WzFunc_new(FN_makeDefaultsOptionsForm), "Defaults"
    End If
    OptionsBrowserForm_addOptionsForm result, OptionsBrowserForm_Modes_Graphics, WzFunc_new(FN_makeGraphicsOptionsForm), "Graphics"
    OptionsBrowserForm_addOptionsForm result, OptionsBrowserForm_Modes_Audio, WzFunc_new(FN_makeAudioOptionsForm), "Audio"
    OptionsBrowserForm_addOptionsForm result, OptionsBrowserForm_Modes_Controls, WzFunc_new(FN_makeControlsOptionsForm), "Controls"

    ' (#if !defined(__EMSCRIPTEN__))
    OptionsBrowserForm_addOptionsForm result, OptionsBrowserForm_Modes_Window, WzFunc_new(FN_makeWindowOptionsForm), "Window"

    Set createOptionsBrowser = result
End Function

' MARK: - the lambdas

Public Function optionsforms_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongSelf As WIDGET, currLeftMostFullButtonIdx As Long, currRightMostFullButtonIdx As Long, psWidget As WIDGET
    Dim psParent As WIDGET, h As Long, y0 As Long, configErrorMessage As String, strongOptionsBrowserForm As WIDGET
    Select Case f.id
    Case FN_HBW_scrollLeftClick                 ' [weakSelf](W_BUTTON&)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/options/optionsforms.cpp:368", "initialize", "Null pointer": Exit Function
        currLeftMostFullButtonIdx = HorizontalButtonsWidget_currentLeftMostFullyVisibleButtonIdx(strongSelf)
        If currLeftMostFullButtonIdx > 0 Then
            HorizontalButtonsWidget_scrollEnsureButtonVisible strongSelf, currLeftMostFullButtonIdx - 1
        End If

    Case FN_HBW_scrollRightClick                ' [weakSelf](W_BUTTON&)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/options/optionsforms.cpp:377", "initialize", "Null pointer": Exit Function
        currRightMostFullButtonIdx = HorizontalButtonsWidget_currentRightMostFullyVisibleButtonIdx(strongSelf)
        If currRightMostFullButtonIdx + 1 < HorizontalButtonsWidget_(strongSelf).optionsTitles.Count Then
            HorizontalButtonsWidget_scrollEnsureButtonVisible strongSelf, currRightMostFullButtonIdx + 1
        End If

    Case FN_HBW_buttonClick                     ' [weakSelf, newButtonIdx](W_BUTTON& but)
        widgScheduleTask WzFunc_new(FN_HBW_buttonClickTask, f.cap(0), f.cap(1))

    Case FN_HBW_buttonClickTask                 ' [weakSelf, newButtonIdx]()
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/options/optionsforms.cpp:435", "addButton", "Null pointer?": Exit Function
        HorizontalButtonsWidget_selectItem strongSelf, CLng(f.cap(1))

    Case FN_OBF_formSwitcherSelect              ' [weakSelf](const WzString& untranslatedTitle)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/options/optionsforms.cpp:794", "initialize", "No self?": Exit Function
        OptionsBrowserForm_displayOptionsForm_title strongSelf, CStr(args(0))

    Case FN_OBF_sectionSwitcherSelect           ' [weakSelf](const WzString& sectionId)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/options/optionsforms.cpp:804", "initialize", "No self?": Exit Function
        OptionsBrowserForm_displayOptionsFormSection strongSelf, CStr(args(0))

    Case FN_OBF_configDirLinkCalcLayout         ' LAMBDA_CALCLAYOUT_SIMPLE: (WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/options/optionsforms.cpp:821", "initialize", "No parent?": Exit Function
        h = psParent.height()
        y0 = h - psWidget.height()
        psWidget.setGeometry 10, y0, psWidget.width(), psWidget.height()

    Case FN_OBF_configDirLinkClick              ' [](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_OBF_configDirLinkTask)

    Case FN_OBF_configDirLinkTask               ' []()
        If Not openFolderInDefaultFileManager(PHYSFS_getWriteDir()) Then
            ' Failed to open write dir in default filesystem browser
            configErrorMessage = "Failed to open configuration directory in system default file browser."
            configErrorMessage = configErrorMessage & vbLf & vbLf
            configErrorMessage = configErrorMessage & "Configuration directory is reported as:"
            configErrorMessage = configErrorMessage & vbLf
            configErrorMessage = configErrorMessage & PHYSFS_getWriteDir()
            configErrorMessage = configErrorMessage & vbLf & vbLf
            configErrorMessage = configErrorMessage & "If running inside a container / isolated environment, this may differ from the actual path on disk."
            configErrorMessage = configErrorMessage & vbLf
            configErrorMessage = configErrorMessage & "Please see the documentation for more information on how to locate it manually."
            wzDisplayDialog Dialog_Warning, "Failed to open configuration directory", configErrorMessage
        End If

    Case FN_createOptionsBrowser_informLanguage ' [weakOptionsBrowserForm]()
        Set strongOptionsBrowserForm = WeakWidget_lock(f.cap(0))
        If Not (Not strongOptionsBrowserForm Is Nothing) Then ASSERT_FAILURE "strongOptionsBrowserForm != nullptr", "src/titleui/options/optionsforms.cpp:1099", "createOptionsBrowser", "No parent?": Exit Function
        OptionsBrowserForm_informLanguageDidChange strongOptionsBrowserForm

    Case FN_createOptionsBrowser_interface      ' [inGame, informOnLanguageChangeHandler]()
        Set optionsforms_c_fn = makeInterfaceOptionsForm(CBool(f.cap(0)), f.cap(1))

    Case FN_makeDefaultsOptionsForm
        Set optionsforms_c_fn = makeDefaultsOptionsForm()

    Case FN_makeGraphicsOptionsForm
        Set optionsforms_c_fn = makeGraphicsOptionsForm()

    Case FN_makeAudioOptionsForm
        Set optionsforms_c_fn = makeAudioOptionsForm()

    Case FN_makeControlsOptionsForm
        Set optionsforms_c_fn = makeControlsOptionsForm()

    Case FN_makeWindowOptionsForm
        Set optionsforms_c_fn = makeWindowOptionsForm()
    End Select
End Function
