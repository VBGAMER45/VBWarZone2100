Attribute VB_Name = "optionsform_c"
' src/titleui/widgets/optionsform.cpp and optionsform.h: the widgets that display a list of
' changeable options (OptionsForm), their value changers (OptionsSlider, OptionsDropdown<T>),
' OptionInfo (OptionInfo.cls), OptionsSection, and the rows, whole. The widget classes follow
' docs/CONVENTIONS.md, "Widgets" (their fields: the data classes <Class>.cls).
' (VB) OptionValueChangerInterface, the second base class of a value changer, is the WIDGET
' virtuals ovc_<method> (declared below, defaults here); ovc_is is dynamic_cast's test.
' OptionsDropdown<T> is one kind for every T: the values are Variants (OptionChoices.cls).
' OptionsForm's unordered_maps are a Scripting.Dictionary (id -> index) over arrays in insertion
' order. WzOptionsLabelWidget (a typedef of WzOptionsChoiceWidget) is that kind.
Option Explicit

' (VB) widget class OptionsSlider : WIDGET
' (VB) widget class WzOptionsChoiceWidget : WIDGET
' (VB) widget class WzClickableOptionsChoiceWidget : WzOptionsChoiceWidget
' (VB) widget class WzOptionsDropdownWidget : DropdownWidget
' (VB) widget class OptionsDropdown : WzOptionsDropdownWidget
' (VB) widget class OptionsValueChangerWidgetWrapper : W_FORM
' (VB) widget class OptionsRowWidgetBase : WIDGET
' (VB) widget class OptionsPaddingRow : OptionsRowWidgetBase
' (VB) widget class OptionsSectionRow : OptionsRowWidgetBase
' (VB) widget class OptionRow : OptionsRowWidgetBase
' (VB) widget class OptionsForm : W_FORM
' (VB) widget class WzHelpPopoverWidget : WIDGET

' class OptionValueChangerInterface's virtuals (on WIDGET)
' (VB) virtual WIDGET.ovc_is
' (VB) virtual WIDGET.ovc_update
' (VB) virtual WIDGET.ovc_informAvailable
' (VB) virtual WIDGET.ovc_addOnChangeHandler
' (VB) virtual WIDGET.ovc_addStartingCapturedEditingHandler
' (VB) virtual WIDGET.ovc_addEndingCapturedEditingHandler
' (VB) virtual WIDGET.ovc_getHelpDescriptionForCurrentValue
' (VB) virtual WIDGET.ovc_getHelpDescriptions
' (VB) virtual OptionsRowWidgetBase.setIsHighlighted

Public Const FN_optionsform_displayBigSlider As Long = 4000  ' static void displayBigSlider(WIDGET *psWidget, UDWORD xOffset, UDWORD yOffset)
Public Const FN_IsNotInGame As Long = 4001                  ' OptionInfo::AvailabilityResult IsNotInGame(const OptionInfo&)
Public Const FN_IsNotMultiplayer As Long = 4002             ' OptionInfo::AvailabilityResult IsNotMultiplayer(const OptionInfo&)
Public Const FN_OptionsSlider_onSliderChange As Long = 4003 ' OptionsSlider::make, [weakParent](W_SLIDER& s)
Public Const FN_OptionsForm_onOptionChange As Long = 4004   ' OptionsForm::addOptionInternal, [pWeakOptionsForm, optionIdCopy](WIDGET&)
Public Const FN_OptionsForm_onStartedEditing As Long = 4005 ' OptionsForm::addOptionInternal, [pWeakOptionsForm, optionIdCopy](WIDGET&)
Public Const FN_OptionsForm_onEndedEditing As Long = 4006   ' OptionsForm::addOptionInternal, [pWeakOptionsForm, optionIdCopy](WIDGET&)
Public Const FN_OptionsDropdown_canChange As Long = 4007    ' OptionsDropdown<T>::make, setCanChange([](DropdownWidget &widget, size_t newIndex, std::shared_ptr<WIDGET> newSelectedWidget) -> bool)
Public Const FN_OptionsDropdown_updateTask As Long = 4008   ' OptionsDropdown<T>::make, widgScheduleTask([strongSelf]() {...})

Public Type OptionChoiceHelpDescription
    displayName As String
    helpDescription As String
End Type

' template<typename T> class OptionChoice: a unique option choice (T: a Variant)
Public Type OptionChoice
    displayName As String
    helpDescription As String
    value As Variant
    disabled As Boolean                     ' = false
    icon As Long                            ' AtlasImageDef *: a handle (0 = nullptr)
End Type

' class OptionsSection: a section header for a subsequent series of options (OptionsSection_(displayName, helpDescription))
Public Type OptionsSection
    displayName As String                   ' private
    helpDescription As String
End Type

' OptionsForm::OptionDetails
Public Type OptionsForm_OptionDetails
    info As OptionInfo
    optionLabelWidget As WIDGET
    optionValueChangerWidgetWrapper As WIDGET
    idxInOptionsList As Long
End Type

' OptionsForm::SectionDetails
Public Type OptionsForm_SectionDetails
    info As OptionsSection
    optionSectionWidget As WIDGET
    idxInOptionsList As Long
End Type

' the const members of the rows
Private Const OptionsSectionRow_outerPaddingX As Long = 1
Private Const OptionsSectionRow_outerPaddingY As Long = 8
Private Const OptionsSectionRow_innerPaddingX As Long = 4
Private Const OptionRow_outerPaddingX As Long = 1
Private Const OptionRow_outerPaddingY As Long = 0
Private Const OptionRow_innerPaddingX As Long = 6
Private Const OptionRow_indentPaddingPerLevel As Long = 10
Private Const OptionRow_minimumLabelWidth As Long = 100
Private Const WzHelpPopoverWidget_outerPadding As Long = 10

' MARK: - (VB) helpers

' (VB) PIELIGHT's byte.a, read and set
Private Function pielightA(ByVal c As Long) As Long
    pielightA = SHRU(c, 24) And &HFF&
End Function

Private Function pielightWithA(ByVal c As Long, ByVal a As Long) As Long
    pielightWithA = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' (VB) FrontImages == nullptr (init_c's Variant: Empty, or a loader's Nothing)
Private Function isFrontImagesNull() As Boolean
    If IsObject(FrontImages) Then
        isFrontImagesNull = (FrontImages Is Nothing)
    Else
        isFrontImagesNull = IsEmpty(FrontImages)
    End If
End Function

' (VB) OptionChoice::operator== on the value: T's operator== (Empty is nullopt; a string equals only a string)
Public Function OptionValue_equal(a As Variant, b As Variant) As Boolean
    If IsEmpty(a) Or IsEmpty(b) Then
        OptionValue_equal = (IsEmpty(a) And IsEmpty(b))
    ElseIf (VarType(a) = vbString) <> (VarType(b) = vbString) Then
        OptionValue_equal = False
    Else
        OptionValue_equal = (a = b)
    End If
End Function

' MARK: OptionsSlider

Private Sub displayBigSlider(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim Slider As W_SLIDER, x As Long, y As Long, sx As Long
    Set Slider = W_SLIDER_(psWidget)
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    iV_DrawImage IntImages, IMAGE_SLIDER_BIG, x, y                ' draw bdrop

    sx = ((psWidget.width() - 3 - Slider.barSize) * Slider.pos) \ MAX_(Slider.numStops, 1) ' determine pos.
    iV_DrawImage IntImages, IMAGE_SLIDER_BIGBUT, x + 3 + sx, y + 3 ' draw amount
End Sub

' OptionsSlider::OptionsSlider() (with the member initialisers)
Public Sub OptionsSlider_construct(ByVal self As WIDGET)
    Dim d As OptionsSlider
    WIDGET_construct self
    self.kind = K_OptionsSlider
    Set d = OptionsSlider_(self)
    d.cachedIdealSliderWidth = 0
    Set d.wzText = WzCachedText_new()
    d.FontID = font_regular
    d.cachedIdealTextWidth = 0
    d.cachedIdealTextLineSize = 0
    d.lastWidgetWidth = 0
    d.outerPaddingY = 8
    d.horizontalPadding = 10
    d.isDisabled = False
    d.isTruncated = False
    d.minValue = 0
    d.maxValue = 0
    d.totalRange = 0
    d.stepValue = 1
End Sub

' static std::shared_ptr<OptionsSlider> make(int32_t minValue, int32_t maxValue, int32_t stepValue, CurrentValueFunc currentValFunc, OnChangeValueFunc onChange, bool needsStateUpdates)
Public Function OptionsSlider_make(ByVal minValue As Long, ByVal maxValue As Long, ByVal stepValue As Long, ByVal currentValFunc As WzFunc, ByVal onChange As WzFunc, ByVal needsStateUpdates As Boolean) As WIDGET
    Dim totalRange As Long, numStops As Long, currPos As Long, sSldInit As W_SLDINIT
    Dim sliderWidget As WIDGET, result As WIDGET, d As OptionsSlider, weakParent As Long
    W_SLDINIT_construct sSldInit                ' W_SLDINIT sSldInit;
    If Not (minValue < maxValue) Then ASSERT_FAILURE "minValue < maxValue", "src/titleui/widgets/optionsform.cpp:61", "make", "minValue is not < maxValue": Exit Function
    If Not (Not currentValFunc Is Nothing) Then ASSERT_FAILURE "currentValFunc != nullptr", "src/titleui/widgets/optionsform.cpp:62", "make", "Invalid currentVal func": Exit Function
    If Not (IntImages <> 0) Then ASSERT_FAILURE "IntImages != nullptr", "src/titleui/widgets/optionsform.cpp:63", "make", "IntImages not loaded?": Exit Function

    totalRange = maxValue - minValue                        ' uint32_t
    numStops = ((maxValue - minValue) \ stepValue) And &HFFFF&    ' UWORD
    currPos = ((CLng(WzFunc_call(currentValFunc)) - minValue) \ stepValue) And &HFFFF&

    sSldInit.formID = 0
    sSldInit.id = 0
    sSldInit.width = iV_GetImageWidth(IntImages, IMAGE_SLIDER_BIG)
    sSldInit.height = iV_GetImageHeight(IntImages, IMAGE_SLIDER_BIG)
    sSldInit.numStops = numStops
    sSldInit.barSize = iV_GetImageHeight(IntImages, IMAGE_SLIDER_BIG)
    sSldInit.pos = currPos
    Set sSldInit.pDisplay = WzFunc_new(FN_optionsform_displayBigSlider)

    Set sliderWidget = W_SLIDER_new_init(sSldInit)

    Set result = WIDGET_alloc(K_OptionsSlider)
    OptionsSlider_construct result
    Set d = OptionsSlider_(result)
    Set d.sliderWidget = sliderWidget
    d.minValue = minValue
    d.maxValue = maxValue
    d.totalRange = totalRange
    d.stepValue = stepValue
    Set d.currentValFunc = currentValFunc
    Set d.onChangeValueFunc = onChange
    d.needsStateUpdates = needsStateUpdates

    weakParent = ObjPtr(result)
    W_SLIDER_addOnChange d.sliderWidget, WzFunc_new(FN_OptionsSlider_onSliderChange, weakParent)
    result.attach d.sliderWidget

    d.cachedIdealSliderWidth = iV_GetImageWidth(IntImages, IMAGE_SLIDER_BIG)
    OptionsSlider_updateSliderValueText result

    Set OptionsSlider_make = result
End Function

Public Function OptionsSlider_currentValue(ByVal self As WIDGET) As Long
    Dim d As OptionsSlider
    Set d = OptionsSlider_(self)
    OptionsSlider_currentValue = (W_SLIDER_(d.sliderWidget).pos * d.stepValue) + d.minValue
End Function

Private Sub OptionsSlider_handleSliderPosChanged(ByVal self As WIDGET)
    Dim d As OptionsSlider, newValue As Long
    Set d = OptionsSlider_(self)
    newValue = OptionsSlider_currentValue(self)
    If Not d.onChangeValueFunc Is Nothing Then
        WzFunc_call d.onChangeValueFunc, newValue
    End If
    OptionsSlider_updateSliderValueText self
End Sub

Private Sub OptionsSlider_updateSliderValueText(ByVal self As WIDGET)
    OptionsSlider_(self).text = CStr(OptionsSlider_currentValue(self))
    OptionsSlider_cacheIdealTextSize self
End Sub

Private Function OptionsSlider_maxNumValueCharacters(ByVal self As WIDGET) As Long
    Dim d As OptionsSlider, minValueStr As String, maxValueStr As String
    Set d = OptionsSlider_(self)
    minValueStr = CStr(d.minValue)
    maxValueStr = CStr(d.maxValue)
    OptionsSlider_maxNumValueCharacters = MAX_(Len(minValueStr), Len(maxValueStr))
End Function

Private Sub OptionsSlider_cacheIdealTextSize(ByVal self As WIDGET)
    Dim d As OptionsSlider, tmpStr As String
    Set d = OptionsSlider_(self)
    tmpStr = "444444"
    tmpStr = Left$(tmpStr, OptionsSlider_maxNumValueCharacters(self))  ' truncate
    d.cachedIdealTextWidth = iV_GetTextWidth(tmpStr, d.FontID) + 2
    d.cachedIdealTextLineSize = iV_GetTextLineSize(d.FontID)
End Sub

Public Sub OptionsSlider_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As OptionsSlider, x0 As Long, y0 As Long, w As Long, h As Long, isHighlight As Boolean
    Dim fontID As iV_fonts, shrunk As iV_fonts, maxTextDisplayableWidth As Long, textColor As Long
    Dim maxDisplayableMainTextWidth As Long, textX0 As Long, textY0 As Long
    Set d = OptionsSlider_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()

    isHighlight = False

    fontID = d.wzText.getFontID()
    If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzText.getText() <> d.text Then
        fontID = d.FontID
        d.wzText.setText d.text, fontID
    End If

    maxTextDisplayableWidth = w - d.cachedIdealSliderWidth - (d.horizontalPadding * 3)
    If d.wzText.deref().width() > maxTextDisplayableWidth Then
        ' text would exceed the bounds of the button area
        ' try to shrink font so it fits
        Do
            If fontID = font_small Or fontID = font_bar Then
                Exit Do
            End If
            If Not iV_ShrinkFont(fontID, shrunk) Then
                Exit Do
            End If
            fontID = shrunk
            d.wzText.setText d.text, fontID
        Loop While d.wzText.deref().width() > maxTextDisplayableWidth
    End If
    d.lastWidgetWidth = self.width()

    ' Draw the main text
    If isHighlight Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
    If d.isDisabled Then
        textColor = pielightWithA(textColor, pielightA(textColor) \ 2)
    End If

    maxDisplayableMainTextWidth = maxTextDisplayableWidth

    textX0 = x0 + w - d.horizontalPadding - MIN_(maxDisplayableMainTextWidth, d.wzText.deref().width())
    textY0 = y0 + (h - d.wzText.deref().lineSize()) \ 2 - d.wzText.deref().aboveBase()

    d.isTruncated = maxDisplayableMainTextWidth < d.wzText.deref().width()
    If d.isTruncated Then
        maxDisplayableMainTextWidth = maxDisplayableMainTextWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
    End If
    d.wzText.deref().render textX0, textY0, textColor, 0!, maxDisplayableMainTextWidth
    If d.isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxDisplayableMainTextWidth + 2, textY0, textColor
    End If
End Sub

Public Sub OptionsSlider_geometryChanged(ByVal self As WIDGET)
    Dim d As OptionsSlider, w As Long, h As Long, sliderX0 As Long, sliderY0 As Long, maxSliderWidth As Long, sliderWidth As Long
    Set d = OptionsSlider_(self)
    w = self.width()
    h = self.height()

    If w = 0 Or h = 0 Then
        Exit Sub
    End If

    sliderX0 = d.horizontalPadding
    sliderY0 = (h - d.sliderWidget.height()) \ 2
    maxSliderWidth = w - (d.horizontalPadding * 2)
    sliderWidth = MIN_(maxSliderWidth, d.cachedIdealSliderWidth)
    d.sliderWidget.setGeometry sliderX0, sliderY0, sliderWidth, d.sliderWidget.height()
End Sub

Public Function OptionsSlider_idealWidth(ByVal self As WIDGET) As Long
    Dim d As OptionsSlider
    Set d = OptionsSlider_(self)
    OptionsSlider_idealWidth = d.horizontalPadding + d.cachedIdealSliderWidth + d.horizontalPadding + d.cachedIdealTextWidth + d.horizontalPadding
End Function

Public Function OptionsSlider_idealHeight(ByVal self As WIDGET) As Long
    OptionsSlider_idealHeight = (OptionsSlider_(self).outerPaddingY * 2) + iV_GetTextLineSize(font_regular)
End Function

Public Function OptionsSlider_ovc_is(ByVal self As WIDGET) As Boolean
    OptionsSlider_ovc_is = True
End Function

Public Sub OptionsSlider_ovc_update(ByVal self As WIDGET, Optional ByVal force As Boolean = False)
    Dim d As OptionsSlider
    Set d = OptionsSlider_(self)
    If Not force And Not d.needsStateUpdates Then
        Exit Sub
    End If

    W_SLIDER_(d.sliderWidget).pos = ((CLng(WzFunc_call(d.currentValFunc)) - d.minValue) \ d.stepValue) And &HFFFF&
End Sub

Public Sub OptionsSlider_ovc_informAvailable(ByVal self As WIDGET, ByVal isAvailable As Boolean)
    ' TODO: Set disabled flag so display can disable the slider appearance?
End Sub

Public Sub OptionsSlider_ovc_addOnChangeHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    W_SLIDER_addOnChange OptionsSlider_(self).sliderWidget, handler
End Sub

' MARK: -

' WzOptionsDropdownWidget::WzOptionsDropdownWidget() (with the member initialisers)
Public Sub WzOptionsDropdownWidget_construct(ByVal self As WIDGET)
    Dim d As WzOptionsDropdownWidget, caretPadding As Padding
    DropdownWidget_construct self
    self.kind = K_WzOptionsDropdownWidget
    Set d = WzOptionsDropdownWidget_(self)
    d.paddingLeft = 4
    d.isHighlighted = False
    If Not isFrontImagesNull() Then
        caretPadding.top = 0
        caretPadding.right_ = 8
        caretPadding.bottom = 0
        caretPadding.left_ = 0
        DropdownWidget_setDropdownCaretImage self, True, AtlasImage_(FrontImages, IMAGE_CARET_DOWN_FILL), WzSize_(12, 12), caretPadding
    Else
        ' FUTURE TODO: Different image for in-game options? (For now, no image)
    End If
    DropdownWidget_setStyle self, DropdownWidget_DropdownMenuStyle_Separate
    DropdownWidget_setListBackgroundColor self, pal_RGBA(20, 20, 20, 255)
End Sub

Public Sub WzOptionsDropdownWidget_setTextAlignment(ByVal self As WIDGET, ByVal align As WzTextAlignment)
    self.style = self.style And Not (WLAB_ALIGNLEFT Or WLAB_ALIGNCENTRE Or WLAB_ALIGNRIGHT)
    self.style = self.style Or align
End Sub

Public Function WzOptionsDropdownWidget_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    ' deliberately skip the default DropdownWidget::findMouseTargetRecursive behavior (which forwards highlighting events to the selected item)
    Set WzOptionsDropdownWidget_findMouseTargetRecursive = WIDGET_findMouseTargetRecursive(self, psContext, key, wasPressed)
End Function

Public Sub WzOptionsDropdownWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    If DropdownWidget_isOpen(self) Or WzOptionsDropdownWidget_(self).isHighlighted Then
        vt_DropdownWidget_drawOpenedHighlight self, xOffset, yOffset
    End If

    If Not DropdownWidget_getIsDisabled(self) Then
        DropdownWidget_drawDropdownCaretImage self, xOffset, yOffset, psPalette(WZCOL_TEXT_MEDIUM)
    End If
End Sub

Public Sub WzOptionsDropdownWidget_geometryChanged(ByVal self As WIDGET)
    WzOptionsDropdownWidget_updateCurrentChoiceDisplayWidgetGeometry self
    DropdownWidget_geometryChanged self
End Sub

Public Sub WzOptionsDropdownWidget_drawOpenedHighlight(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzOptionsDropdownWidget, x0 As Long, y0 As Long, w As Long, h As Long, backgroundColor As Long
    Dim backX0 As Long, backX1 As Long, backY0 As Long, backY1 As Long, highlightWidth As Long, itemWidth As Long
    Set d = WzOptionsDropdownWidget_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    w = self.width()
    h = self.height()

    If Not d.currentOptionChoiceDisplayWidg Is Nothing Then
        If d.isHighlighted Then
            backgroundColor = psPalette(WZCOL_MENU_SCORE_BUILT)
        Else
            backgroundColor = psPalette(WZCOL_FORM_DARK)
            backgroundColor = pielightWithA(backgroundColor, MIN_(pielightA(backgroundColor), 60))
        End If
        backX0 = x0
        backX1 = backX0 + w
        If (self.style And WLAB_ALIGNRIGHT) <> 0 Then
            highlightWidth = d.currentOptionChoiceDisplayWidg.width() + d.paddingLeft
            itemWidth = DropdownWidget_getItemDisplayWidth(self)
            backX0 = (x0 + itemWidth) - highlightWidth
        End If
        backY0 = y0
        backY1 = backY0 + h
        pie_UniTransBoxFill backX0, backY0, backX1, backY1, backgroundColor
    End If
End Sub

Public Sub WzOptionsDropdownWidget_drawSelectedItem(ByVal self As WIDGET, ByVal item As WIDGET, displayArea As WzRect)
    ' no-op because we handle this with the currentOptionChoiceDisplayWidg
End Sub

Public Sub WzOptionsDropdownWidget_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzOptionsDropdownWidget
    Set d = WzOptionsDropdownWidget_(self)
    d.isHighlighted = True
    If Not d.currentOptionChoiceDisplayWidg Is Nothing Then
        WzOptionsChoiceWidget_setDisplayHighlighted d.currentOptionChoiceDisplayWidg, True
    End If
End Sub

Public Sub WzOptionsDropdownWidget_highlightLost(ByVal self As WIDGET)
    Dim d As WzOptionsDropdownWidget
    Set d = WzOptionsDropdownWidget_(self)
    d.isHighlighted = False
    If Not d.currentOptionChoiceDisplayWidg Is Nothing Then
        WzOptionsChoiceWidget_setDisplayHighlighted d.currentOptionChoiceDisplayWidg, False
    End If
End Sub

Public Function WzOptionsDropdownWidget_calculateDropdownListScreenPosX(ByVal self As WIDGET) As Long
    Dim xPos As Long, displayWidth As Long
    ' align with left side of current displayed text label (if possible)
    xPos = WzOptionsDropdownWidget_(self).currentOptionChoiceDisplayWidg.screenPosX()
    displayWidth = vt_DropdownWidget_calculateDropdownListDisplayWidth(self)
    If xPos + displayWidth > screenWidth Then
        xPos = MAX_(screenWidth - displayWidth, 0)
    End If
    WzOptionsDropdownWidget_calculateDropdownListScreenPosX = xPos
End Function

Public Function WzOptionsDropdownWidget_calculateDropdownListDisplayWidth(ByVal self As WIDGET) As Long
    WzOptionsDropdownWidget_calculateDropdownListDisplayWidth = MAX_(WzOptionsDropdownWidget_calculateCurrentUsedWidth(self), DropdownWidget_calculateDropdownListDisplayWidth(self))
End Function

Private Function WzOptionsDropdownWidget_calculateCurrentUsedWidth(ByVal self As WIDGET) As Long
    Dim d As WzOptionsDropdownWidget, result As Long
    Set d = WzOptionsDropdownWidget_(self)
    result = DropdownWidget_getCaretImageUsedWidth(self)
    If Not d.currentOptionChoiceDisplayWidg Is Nothing Then
        result = result + d.currentOptionChoiceDisplayWidg.idealWidth()
    End If
    WzOptionsDropdownWidget_calculateCurrentUsedWidth = result
End Function

Public Sub WzOptionsDropdownWidget_onSelectedItemChanged(ByVal self As WIDGET)
    Dim d As WzOptionsDropdownWidget, item As WIDGET
    Set d = WzOptionsDropdownWidget_(self)
    Set item = DropdownWidget_getSelectedItem(self)
    If d.currentOptionChoiceDisplayWidg Is Nothing Then
        Set d.currentOptionChoiceDisplayWidg = WzOptionsChoiceWidget_make(font_regular_bold)
        WzOptionsChoiceWidget_setPadding d.currentOptionChoiceDisplayWidg, 10, 8
        d.currentOptionChoiceDisplayWidg.setTransparentToMouse True
        WzOptionsChoiceWidget_setHighlightBackgroundColor d.currentOptionChoiceDisplayWidg, pal_RGBA(0, 0, 0, 0)
        self.attach d.currentOptionChoiceDisplayWidg
    End If
    d.currentOptionChoiceDisplayWidg.setString item.getString()
    WzOptionsDropdownWidget_updateCurrentChoiceDisplayWidgetGeometry self

    DropdownWidget_onSelectedItemChanged self
End Sub

Private Sub WzOptionsDropdownWidget_updateCurrentChoiceDisplayWidgetGeometry(ByVal self As WIDGET)
    Dim d As WzOptionsDropdownWidget, maxItemWidth As Long, itemWidth As Long, itemHeight As Long, itemX0 As Long, itemY0 As Long
    Set d = WzOptionsDropdownWidget_(self)
    If Not d.currentOptionChoiceDisplayWidg Is Nothing Then
        maxItemWidth = DropdownWidget_getItemDisplayWidth(self)
        itemWidth = MIN_(maxItemWidth, d.currentOptionChoiceDisplayWidg.idealWidth())
        itemHeight = d.currentOptionChoiceDisplayWidg.idealHeight()
        itemX0 = 0
        If (self.style And WLAB_ALIGNRIGHT) <> 0 Then
            itemX0 = maxItemWidth - itemWidth
        Else
            itemX0 = 0
        End If
        itemY0 = (self.height() - itemHeight) \ 2
        d.currentOptionChoiceDisplayWidg.setGeometry itemX0, itemY0, itemWidth, itemHeight
    End If
End Sub

Public Function WzOptionsDropdownWidget_idealWidth(ByVal self As WIDGET) As Long
    WzOptionsDropdownWidget_idealWidth = WzOptionsDropdownWidget_calculateCurrentUsedWidth(self) + WzOptionsDropdownWidget_(self).paddingLeft
End Function

Public Function WzOptionsDropdownWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzOptionsDropdownWidget, a As Long
    Set d = WzOptionsDropdownWidget_(self)
    If Not d.currentOptionChoiceDisplayWidg Is Nothing Then a = d.currentOptionChoiceDisplayWidg.idealHeight() Else a = 0
    WzOptionsDropdownWidget_idealHeight = MAX_(a, DropdownWidget_idealHeight(self))
End Function

' MARK: - WzOptionsChoiceWidget

' WzOptionsChoiceWidget::WzOptionsChoiceWidget(iV_fonts desiredFontID) (with the member initialisers)
Public Sub WzOptionsChoiceWidget_construct(ByVal self As WIDGET, ByVal desiredFontID As iV_fonts)
    Dim d As WzOptionsChoiceWidget
    WIDGET_construct self
    self.kind = K_WzOptionsChoiceWidget
    Set d = WzOptionsChoiceWidget_(self)
    Set d.wzText = WzCachedText_new()
    d.FontID = desiredFontID
    d.icon = 0
    d.iconDrawWidth = 0
    d.iconDrawHeight = 0
    d.lastWidgetWidth = 0
    d.horizontalPadding = 10
    d.iconPaddingAfter = 5
    d.verticalPadding = 5
    d.isHighlighted = False
    d.forceDisplayHighlight = False
    d.isDisabled = False
    d.isTruncated = False
    d.highlightBackgroundColor = psPalette(WZCOL_MENU_SCORE_BUILT)
End Sub

Public Function WzOptionsChoiceWidget_make(ByVal desiredFontID As iV_fonts) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzOptionsChoiceWidget)
    WzOptionsChoiceWidget_construct widget, desiredFontID
    Set WzOptionsChoiceWidget_make = widget
End Function

Private Sub WzOptionsChoiceWidget_recalcIconDrawSize(ByVal self As WIDGET)
    Dim d As WzOptionsChoiceWidget, scaledIconSize As WzSize
    Set d = WzOptionsChoiceWidget_(self)
    scaledIconSize = WzOptionsChoiceWidget_scaleIconSizeFor(self, True, self.width(), True, self.height())
    d.iconDrawWidth = scaledIconSize.width_
    d.iconDrawHeight = scaledIconSize.height_
End Sub

' WzSize scaleIconSizeFor(optional<int> width, optional<int> height)
Private Function WzOptionsChoiceWidget_scaleIconSizeFor(ByVal self As WIDGET, ByVal has_width As Boolean, ByVal width As Long, ByVal has_height As Boolean, ByVal height As Long) As WzSize
    Dim d As WzOptionsChoiceWidget, result As WzSize, iconMaxWidth As Long, iconMaxHeight As Long
    Dim widthScale As Single, heightScale As Single, scale_ As Single
    Set d = WzOptionsChoiceWidget_(self)
    If d.icon = 0 Then
        WzOptionsChoiceWidget_scaleIconSizeFor = result
        Exit Function
    End If

    If has_width Then iconMaxWidth = width - (d.horizontalPadding * 2) Else iconMaxWidth = INT32_MAX - (d.horizontalPadding * 2)
    If has_height Then iconMaxHeight = height - (d.verticalPadding * 2) Else iconMaxHeight = INT32_MAX - (d.verticalPadding * 2)

    result.width_ = AtlasImageDef_Width(d.icon)
    result.height_ = AtlasImageDef_Height(d.icon)
    If result.width_ <= iconMaxWidth And result.height_ <= iconMaxHeight Then
        WzOptionsChoiceWidget_scaleIconSizeFor = result
        Exit Function
    End If

    widthScale = 1!
    heightScale = 1!
    If result.width_ > iconMaxWidth Then
        widthScale = CSng(iconMaxWidth) / result.width_
    End If
    If result.height_ > iconMaxHeight Then
        heightScale = CSng(iconMaxHeight) / result.height_
    End If
    If widthScale < heightScale Then scale_ = widthScale Else scale_ = heightScale
    result.width_ = Fix(CSng(result.width_ * scale_))
    result.height_ = Fix(CSng(result.height_ * scale_))
    WzOptionsChoiceWidget_scaleIconSizeFor = result
End Function

Public Function WzOptionsChoiceWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzOptionsChoiceWidget, result As Long, iconSize As WzSize
    Set d = WzOptionsChoiceWidget_(self)
    result = (d.horizontalPadding * 2) + d.wzText.getTextWidth()
    If d.icon <> 0 Then
        iconSize = WzOptionsChoiceWidget_scaleIconSizeFor(self, False, 0, True, self.idealHeight())
        result = result + (iconSize.width_ + d.iconPaddingAfter)
    End If
    WzOptionsChoiceWidget_idealWidth = result
End Function

Public Function WzOptionsChoiceWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzOptionsChoiceWidget
    Set d = WzOptionsChoiceWidget_(self)
    WzOptionsChoiceWidget_idealHeight = (d.verticalPadding * 2) + d.wzText.getTextLineSize()
End Function

Public Sub WzOptionsChoiceWidget_setString(ByVal self As WIDGET, ByVal string_ As String)
    Dim d As WzOptionsChoiceWidget
    Set d = WzOptionsChoiceWidget_(self)
    d.text = string_
    d.wzText.setText string_, d.FontID      ' any time the string is changed, set wzText to the desired string & FontID, so we can calculate the ideal width + height
End Sub

' void setIcon(const AtlasImageDef *icon) (an AtlasImageDef handle, 0 = nullptr)
Public Sub WzOptionsChoiceWidget_setIcon(ByVal self As WIDGET, ByVal icon_ As Long)
    WzOptionsChoiceWidget_(self).icon = icon_
End Sub

Public Sub WzOptionsChoiceWidget_setDisabled(ByVal self As WIDGET, ByVal value As Boolean)
    WzOptionsChoiceWidget_(self).isDisabled = value
End Sub

Public Function WzOptionsChoiceWidget_getIsDisabled(ByVal self As WIDGET) As Boolean
    WzOptionsChoiceWidget_getIsDisabled = WzOptionsChoiceWidget_(self).isDisabled
End Function

Public Sub WzOptionsChoiceWidget_setDisplayHighlighted(ByVal self As WIDGET, ByVal highlighted As Boolean)
    WzOptionsChoiceWidget_(self).forceDisplayHighlight = highlighted
End Sub

Public Sub WzOptionsChoiceWidget_setPadding(ByVal self As WIDGET, ByVal horizontal As Long, ByVal vertical As Long)
    Dim d As WzOptionsChoiceWidget
    Set d = WzOptionsChoiceWidget_(self)
    d.horizontalPadding = horizontal
    d.verticalPadding = vertical
End Sub

Public Sub WzOptionsChoiceWidget_setTextAlignment(ByVal self As WIDGET, ByVal align As WzTextAlignment)
    self.style = self.style And Not (WLAB_ALIGNLEFT Or WLAB_ALIGNCENTRE Or WLAB_ALIGNRIGHT)
    self.style = self.style Or align
End Sub

Public Sub WzOptionsChoiceWidget_setHighlightBackgroundColor(ByVal self As WIDGET, ByVal color As Long)
    WzOptionsChoiceWidget_(self).highlightBackgroundColor = color
End Sub

Public Function WzOptionsChoiceWidget_getTip(ByVal self As WIDGET) As String
    Dim d As WzOptionsChoiceWidget
    Set d = WzOptionsChoiceWidget_(self)
    If Not d.isTruncated Then
        WzOptionsChoiceWidget_getTip = ""
        Exit Function
    End If
    WzOptionsChoiceWidget_getTip = d.text
End Function

Public Function WzOptionsChoiceWidget_getString(ByVal self As WIDGET) As String
    WzOptionsChoiceWidget_getString = WzOptionsChoiceWidget_(self).text
End Function

Public Sub WzOptionsChoiceWidget_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzOptionsChoiceWidget_(self).isHighlighted = True
End Sub

Public Sub WzOptionsChoiceWidget_highlightLost(ByVal self As WIDGET)
    WzOptionsChoiceWidget_(self).isHighlighted = False
End Sub

Public Sub WzOptionsChoiceWidget_geometryChanged(ByVal self As WIDGET)
    WzOptionsChoiceWidget_recalcIconDrawSize self
End Sub

Public Sub WzOptionsChoiceWidget_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzOptionsChoiceWidget_(self).wzText.tick
End Sub

Public Sub WzOptionsChoiceWidget_screenSizeDidChange(ByVal self As WIDGET, ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim d As WzOptionsChoiceWidget
    Set d = WzOptionsChoiceWidget_(self)
    ' use as a signal that we should re-check any font scaling or layout calculations (as display scaling setting may have changed)
    If d.wzText.getFontID() <> d.FontID Then
        d.wzText.setText d.text, d.FontID
    End If
    d.wzText.resetCachedDimensions
End Sub

Public Sub WzOptionsChoiceWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzOptionsChoiceWidget, x0 As Long, y0 As Long, w As Long, h As Long, isHighlight As Boolean
    Dim fontID As iV_fonts, shrunk As iV_fonts, availableButtonTextWidth As Long, textColor As Long
    Dim maxDisplayableMainTextWidth As Long, textX0 As Long, textY0 As Long, iconX0 As Long, iconY0 As Long
    Set d = WzOptionsChoiceWidget_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()

    isHighlight = Not d.isDisabled And (d.isHighlighted Or d.forceDisplayHighlight)

    fontID = d.wzText.getFontID()
    If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzText.getText() <> d.text Then
        fontID = d.FontID
        d.wzText.setText d.text, fontID
    End If

    availableButtonTextWidth = w - (d.horizontalPadding * 2)
    If d.iconDrawWidth > 0 Then
        availableButtonTextWidth = availableButtonTextWidth - (d.iconDrawWidth + d.iconPaddingAfter)
    End If
    If d.wzText.deref().width() > availableButtonTextWidth Then
        ' text would exceed the bounds of the button area
        ' try to shrink font so it fits
        Do
            If fontID = font_small Or fontID = font_bar Then
                Exit Do
            End If
            If Not iV_ShrinkFont(fontID, shrunk) Then
                Exit Do
            End If
            fontID = shrunk
            d.wzText.setText d.text, fontID
        Loop While d.wzText.deref().width() > availableButtonTextWidth
    End If
    d.lastWidgetWidth = self.width()

    ' Draw background (if highlighted)
    If isHighlight And pielightA(d.highlightBackgroundColor) <> 0 Then   ' !isTransparent()
        pie_UniTransBoxFill x0, y0, x0 + w, y0 + h, d.highlightBackgroundColor
    End If

    ' Draw the main text
    textColor = psPalette(WZCOL_TEXT_BRIGHT)
    If d.isDisabled Then
        textColor = pielightWithA(textColor, pielightA(textColor) \ 2)
    End If

    maxDisplayableMainTextWidth = availableButtonTextWidth

    textX0 = 0
    If (self.style And WLAB_ALIGNRIGHT) <> 0 Then
        textX0 = x0 + w - d.horizontalPadding - MIN_(maxDisplayableMainTextWidth, d.wzText.deref().width())
    Else
        textX0 = x0 + d.horizontalPadding
        If d.iconDrawWidth > 0 Then
            textX0 = textX0 + (d.iconDrawWidth + d.iconPaddingAfter)
        End If
    End If
    textY0 = y0 + (h - d.wzText.deref().lineSize()) \ 2 - d.wzText.deref().aboveBase()

    d.isTruncated = maxDisplayableMainTextWidth < d.wzText.deref().width()
    If d.isTruncated Then
        maxDisplayableMainTextWidth = maxDisplayableMainTextWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
    End If
    d.wzText.deref().render textX0, textY0, textColor, 0!, maxDisplayableMainTextWidth
    If d.isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxDisplayableMainTextWidth + 2, textY0, textColor
    End If

    If d.icon <> 0 Then
        iconX0 = x0 + d.horizontalPadding
        If (self.style And WLAB_ALIGNRIGHT) <> 0 Then
            iconX0 = x0 + w - d.horizontalPadding - MIN_(availableButtonTextWidth, d.wzText.deref().width()) - (d.iconDrawWidth + d.iconPaddingAfter)
        End If
        iconY0 = y0 + (h - d.iconDrawHeight) \ 2
        iV_DrawImage2_image d.icon, iconX0, iconY0, d.iconDrawWidth, d.iconDrawHeight
    End If
End Sub

' typedef WzOptionsChoiceWidget WzOptionsLabelWidget; (the same kind)

' MARK: - WzClickableOptionsChoiceWidget

Public Sub WzClickableOptionsChoiceWidget_construct(ByVal self As WIDGET, ByVal desiredFontID As iV_fonts)
    WzOptionsChoiceWidget_construct self, desiredFontID
    self.kind = K_WzClickableOptionsChoiceWidget
    WzClickableOptionsChoiceWidget_(self).mouseDownOnWidget = False
    Set WzClickableOptionsChoiceWidget_(self).onClickHandlers = New Collection
End Sub

Public Function WzClickableOptionsChoiceWidget_make(ByVal desiredFontID As iV_fonts) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzClickableOptionsChoiceWidget)
    WzClickableOptionsChoiceWidget_construct widget, desiredFontID
    Set WzClickableOptionsChoiceWidget_make = widget
End Function

' OnClickHandler: void (WzClickableOptionsChoiceWidget& button, WIDGET_KEY key)
Public Sub WzClickableOptionsChoiceWidget_addOnClickHandler(ByVal self As WIDGET, ByVal onClickFunc As WzFunc)
    WzClickableOptionsChoiceWidget_(self).onClickHandlers.Add onClickFunc
End Sub

Public Sub WzClickableOptionsChoiceWidget_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    If WzOptionsChoiceWidget_getIsDisabled(self) Then
        Exit Sub
    End If

    If key = WKEY_PRIMARY Or key = WKEY_SECONDARY Then
        WzClickableOptionsChoiceWidget_(self).mouseDownOnWidget = True
    End If
End Sub

' Respond to a mouse button up
Public Sub WzClickableOptionsChoiceWidget_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As WzClickableOptionsChoiceWidget, onClickHandler As WzFunc, i As Long
    Set d = WzClickableOptionsChoiceWidget_(self)
    If WzOptionsChoiceWidget_getIsDisabled(self) Then
        d.mouseDownOnWidget = False
        Exit Sub
    End If

    If key = WKEY_PRIMARY Or key = WKEY_SECONDARY Then
        If d.mouseDownOnWidget Then
            ' Call all onClick event handlers
            For i = 1 To d.onClickHandlers.Count
                Set onClickHandler = d.onClickHandlers(i)
                If Not onClickHandler Is Nothing Then
                    WzFunc_call onClickHandler, self, key
                End If
            Next
            d.mouseDownOnWidget = False
        End If
    End If
End Sub

' Respond to a mouse moving over a button
Public Sub WzClickableOptionsChoiceWidget_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzOptionsChoiceWidget_highlight self, psContext
End Sub

' Respond to the mouse moving off a button
Public Sub WzClickableOptionsChoiceWidget_highlightLost(ByVal self As WIDGET)
    WzClickableOptionsChoiceWidget_(self).mouseDownOnWidget = False
    WzOptionsChoiceWidget_highlightLost self
End Sub

' MARK: - OptionValueChangerInterface

' (VB) the interface's virtuals, declared on WIDGET (a widget that implements the interface overrides
' them; ovc_is says that it does: dynamic_cast<OptionValueChangerInterface*>). The pure virtuals
' (update, informAvailable, addOnChangeHandler) assert here.

Public Function WIDGET_ovc_is(ByVal self As WIDGET) As Boolean
    WIDGET_ovc_is = False
End Function

Public Sub WIDGET_ovc_update(ByVal self As WIDGET, Optional ByVal force As Boolean = False)
    debug_ 0, LOG_ERROR, "update", "(VB) pure virtual OptionValueChangerInterface::update called"
End Sub

Public Sub WIDGET_ovc_informAvailable(ByVal self As WIDGET, ByVal isAvailable As Boolean)
    debug_ 0, LOG_ERROR, "informAvailable", "(VB) pure virtual OptionValueChangerInterface::informAvailable called"
End Sub

Public Sub WIDGET_ovc_addOnChangeHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    debug_ 0, LOG_ERROR, "addOnChangeHandler", "(VB) pure virtual OptionValueChangerInterface::addOnChangeHandler called"
End Sub

Public Sub WIDGET_ovc_addStartingCapturedEditingHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
End Sub

Public Sub WIDGET_ovc_addEndingCapturedEditingHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
End Sub

' optional<OptionChoiceHelpDescription> getHelpDescriptionForCurrentValue(): has_value, the value in result
Public Function WIDGET_ovc_getHelpDescriptionForCurrentValue(ByVal self As WIDGET, result As OptionChoiceHelpDescription) As Boolean
    WIDGET_ovc_getHelpDescriptionForCurrentValue = False    ' nullopt
End Function

' optional<std::vector<OptionChoiceHelpDescription>> getHelpDescriptions(): has_value, the vector in
' result(0 .. count - 1)
Public Function WIDGET_ovc_getHelpDescriptions(ByVal self As WIDGET, result() As OptionChoiceHelpDescription, count As Long) As Boolean
    WIDGET_ovc_getHelpDescriptions = False                  ' nullopt
End Function

' MARK: - OptionInfo

' OptionInfo(const WzString& optionId, const WzString& displayName, const WzString& helpDescription)
' (OptionInfo.cls)
Public Function OptionInfo_new(ByVal optionId As String, ByVal displayName As String, ByVal helpDescription As String) As OptionInfo
    Dim result As OptionInfo
    Set result = New OptionInfo
    result.construct optionId, displayName, helpDescription
    Set OptionInfo_new = result
End Function

' OptionInfo::AvailabilityResult{available, localizedUnavailabilityReason} (OptionInfo_AvailabilityResult.cls)
Public Function OptionInfo_AvailabilityResult_new(ByVal available_ As Boolean, ByVal localizedUnavailabilityReason As String) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = available_
    result.localizedUnavailabilityReason = localizedUnavailabilityReason
    Set OptionInfo_AvailabilityResult_new = result
End Function

' MARK: - OptionAvailabilityConditions (as an OptionAvailabilityCondition: WzFunc_new(FN_IsNotInGame), WzFunc_new(FN_IsNotMultiplayer))

Public Function IsNotInGame(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = GetGameMode() <> GS_NORMAL
    result.localizedUnavailabilityReason = "Not editable while in a game."
    Set IsNotInGame = result
End Function

Public Function IsNotMultiplayer(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult, bIsTrueMultiplayerGame As Boolean
    Set result = New OptionInfo_AvailabilityResult
    bIsTrueMultiplayerGame = bMultiPlayer And NetPlay_bComms
    result.available_ = Not bIsTrueMultiplayerGame
    result.localizedUnavailabilityReason = "Not editable while in a multiplayer game."
    Set IsNotMultiplayer = result
End Function

' MARK: - OptionsSection

Public Function OptionsSection_(ByVal displayName As String, ByVal helpDescription As String) As OptionsSection
    OptionsSection_.displayName = displayName
    OptionsSection_.helpDescription = helpDescription
End Function

Public Function OptionsSection_sectionId(s As OptionsSection) As String
    ' just use the original (base language) displayName as the sectionId
    OptionsSection_sectionId = s.displayName
End Function

Public Function OptionsSection_getTranslatedDisplayName(s As OptionsSection) As String
    If Len(s.displayName) = 0 Then
        OptionsSection_getTranslatedDisplayName = ""
        Exit Function
    End If
    OptionsSection_getTranslatedDisplayName = s.displayName     ' gettext (translations not ported)
End Function

Public Function OptionsSection_getTranslatedHelpDescription(s As OptionsSection) As String
    If Len(s.helpDescription) = 0 Then
        OptionsSection_getTranslatedHelpDescription = ""
        Exit Function
    End If
    OptionsSection_getTranslatedHelpDescription = s.helpDescription
End Function

' MARK: - OptionsValueChangerWidgetWrapper

Public Sub OptionsValueChangerWidgetWrapper_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_OptionsValueChangerWidgetWrapper
End Sub

Public Function OptionsValueChangerWidgetWrapper_wrap(ByVal optionValueChangerWidget As WIDGET) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_OptionsValueChangerWidgetWrapper)
    OptionsValueChangerWidgetWrapper_construct result
    Set OptionsValueChangerWidgetWrapper_(result).optionValueChangerWidget = optionValueChangerWidget
    result.attach optionValueChangerWidget
    Set OptionsValueChangerWidgetWrapper_wrap = result
End Function

Public Function OptionsValueChangerWidgetWrapper_getOptionValueChangerWidget(ByVal self As WIDGET) As WIDGET
    Set OptionsValueChangerWidgetWrapper_getOptionValueChangerWidget = OptionsValueChangerWidgetWrapper_(self).optionValueChangerWidget
End Function

Public Function OptionsValueChangerWidgetWrapper_idealWidth(ByVal self As WIDGET) As Long
    OptionsValueChangerWidgetWrapper_idealWidth = OptionsValueChangerWidgetWrapper_(self).optionValueChangerWidget.idealWidth()
End Function

Public Function OptionsValueChangerWidgetWrapper_idealHeight(ByVal self As WIDGET) As Long
    OptionsValueChangerWidgetWrapper_idealHeight = OptionsValueChangerWidgetWrapper_(self).optionValueChangerWidget.idealHeight()
End Function

Public Sub OptionsValueChangerWidgetWrapper_geometryChanged(ByVal self As WIDGET)
    Dim widg As WIDGET, w As Long, h As Long, idealWidgetWidth As Long
    Dim widgX0 As Long, widgY0 As Long, widgWidth As Long, widgHeight As Long
    Set widg = OptionsValueChangerWidgetWrapper_(self).optionValueChangerWidget
    ' The optionValueChangerWidget should be right-aligned inside the wrapper, and a maximum width of its idealWidth()
    w = self.width()
    h = self.height()
    idealWidgetWidth = widg.idealWidth()

    If idealWidgetWidth < w Then widgX0 = w - idealWidgetWidth Else widgX0 = 0
    widgY0 = 0
    widgWidth = MIN_(idealWidgetWidth, w)
    widgHeight = h
    widg.setGeometry widgX0, widgY0, widgWidth, widgHeight
End Sub

Public Sub OptionsValueChangerWidgetWrapper_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' no-op
End Sub

Public Sub OptionsValueChangerWidgetWrapper_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim childrenContext As WidgetGraphicsContext, translated As WidgetGraphicsContext, child As WIDGET, i As Long
    Dim widgetIsClipped As Boolean, geom As WzRect, widg As WIDGET
    Dim x0 As Long, y0 As Long, widgX0 As Long, widgY0 As Long, widgX1 As Long, widgY1 As Long
    ' call parent displayRecursive
    W_FORM_displayRecursive self, context

    If W_FORM_(self).disableChildren Then
        ' when children are disabled on a W_FORM or subclass, WIDGET::displayRecursive does not actually draw children - but in this case we actually want to
        translated = WidgetGraphicsContext_translatedBy(context, self.x(), self.y())
        childrenContext = WidgetGraphicsContext_setAllowChildDisplayRecursiveIfSelfClipped(translated, False)
        ' Display the widgets on this widget.
        ' NOTE: Draw them in the opposite order that clicks are processed, so behavior matches the visual.
        ' i.e. Since findMouseTargetRecursive handles processing children in decreasing z-order (i.e. "top-down")
        '      we want to draw things in list order (bottom-up) so the "top-most" is drawn last
        For i = 0 To self.childCount() - 1
            Set child = self.childAt(i)
            If child.visible() Then
                child.displayRecursive childrenContext
            End If
        Next

        geom = self.geometry()
        widgetIsClipped = Not WidgetGraphicsContext_clipContains(context, geom)
        If Not widgetIsClipped Then
            ' "over-draw" with the disabled color over the contained changer widget
            Set widg = OptionsValueChangerWidgetWrapper_(self).optionValueChangerWidget
            x0 = WidgetGraphicsContext_getXOffset(context) + self.x()
            y0 = WidgetGraphicsContext_getYOffset(context) + self.y()
            widgX0 = x0 + widg.x()
            widgY0 = y0 + widg.y()
            widgX1 = widgX0 + widg.width()
            widgY1 = widgY0 + widg.height()
            pie_UniTransBoxFill widgX0, widgY0, widgX1, widgY1, psPalette(WZCOL_TRANSPARENT_BOX)
        End If
    End If
End Sub

Public Sub OptionsValueChangerWidgetWrapper_enable(ByVal self As WIDGET, Optional ByVal enableValue As Boolean = True)
    W_FORM_(self).disableChildren = Not enableValue
End Sub

Public Sub OptionsValueChangerWidgetWrapper_disable(ByVal self As WIDGET)
    W_FORM_(self).disableChildren = True    ' skips processing clicks / mouse for all children
End Sub

' MARK: - Options Row Widgets

Public Sub OptionsRowWidgetBase_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_OptionsRowWidgetBase
End Sub

' virtual void setIsHighlighted(bool val) = 0;
Public Sub OptionsRowWidgetBase_setIsHighlighted(ByVal self As WIDGET, ByVal val_ As Boolean)
    debug_ 0, LOG_ERROR, "setIsHighlighted", "(VB) pure virtual OptionsRowWidgetBase::setIsHighlighted called"
End Sub

Public Sub OptionsPaddingRow_construct(ByVal self As WIDGET)
    OptionsRowWidgetBase_construct self
    self.kind = K_OptionsPaddingRow
    OptionsPaddingRow_(self).paddingHeight = 0
End Sub

Public Function OptionsPaddingRow_make(ByVal paddingHeight As Long) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_OptionsPaddingRow)
    OptionsPaddingRow_construct result
    OptionsPaddingRow_(result).paddingHeight = paddingHeight
    result.setGeometry 0, 0, result.idealWidth(), result.idealHeight()
    Set OptionsPaddingRow_make = result
End Function

Public Sub OptionsPaddingRow_setIsHighlighted(ByVal self As WIDGET, ByVal val_ As Boolean)
    ' no-op
End Sub

Public Function OptionsPaddingRow_idealWidth(ByVal self As WIDGET) As Long
    OptionsPaddingRow_idealWidth = 100
End Function

Public Function OptionsPaddingRow_idealHeight(ByVal self As WIDGET) As Long
    OptionsPaddingRow_idealHeight = OptionsPaddingRow_(self).paddingHeight
End Function

Public Sub OptionsPaddingRow_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' no-op
End Sub

' MARK: OptionsSectionRow

Public Sub OptionsSectionRow_construct(ByVal self As WIDGET)
    Dim d As OptionsSectionRow
    OptionsRowWidgetBase_construct self
    self.kind = K_OptionsSectionRow
    Set d = OptionsSectionRow_(self)
    d.showBottomBorder = False
    d.has_flashHighlightStart = False   ' = nullopt
End Sub

Public Function OptionsSectionRow_make(ByVal sectionName As String) As WIDGET
    Dim result As WIDGET, d As OptionsSectionRow
    Set result = WIDGET_alloc(K_OptionsSectionRow)
    OptionsSectionRow_construct result
    Set d = OptionsSectionRow_(result)

    Set d.m_labelWidget = W_LABEL_new()
    W_LABEL_setFont d.m_labelWidget, font_regular_bold, psPalette(WZCOL_TEXT_MEDIUM)
    d.m_labelWidget.setString sectionName
    W_LABEL_setTextAlignment d.m_labelWidget, WLAB_ALIGNLEFT
    d.m_labelWidget.setGeometry 0, 0, d.m_labelWidget.idealWidth(), d.m_labelWidget.idealHeight()
    result.attach d.m_labelWidget

    result.setGeometry 0, 0, result.idealWidth(), result.idealHeight()

    Set OptionsSectionRow_make = result
End Function

Public Sub OptionsSectionRow_setIsHighlighted(ByVal self As WIDGET, ByVal val_ As Boolean)
    ' no-op
End Sub

Public Sub OptionsSectionRow_setShowBottomBorder(ByVal self As WIDGET, ByVal show As Boolean)
    OptionsSectionRow_(self).showBottomBorder = show
End Sub

Public Sub OptionsSectionRow_setSectionName(ByVal self As WIDGET, ByVal sectionName As String)
    OptionsSectionRow_(self).m_labelWidget.setString sectionName
End Sub

Public Function OptionsSectionRow_idealWidth(ByVal self As WIDGET) As Long
    OptionsSectionRow_idealWidth = (OptionsSectionRow_outerPaddingX * 2) + (OptionsSectionRow_innerPaddingX * 2) + OptionsSectionRow_(self).m_labelWidget.idealWidth()
End Function

Public Function OptionsSectionRow_idealHeight(ByVal self As WIDGET) As Long
    OptionsSectionRow_idealHeight = (OptionsSectionRow_outerPaddingY * 2) + OptionsSectionRow_(self).m_labelWidget.idealHeight()
End Function

Public Sub OptionsSectionRow_geometryChanged(ByVal self As WIDGET)
    OptionsSectionRow_(self).m_labelWidget.setGeometry OptionsSectionRow_outerPaddingX + OptionsSectionRow_innerPaddingX, OptionsSectionRow_outerPaddingY, self.width() - (OptionsSectionRow_outerPaddingX * 2) - (OptionsSectionRow_innerPaddingX * 2), self.height() - (OptionsSectionRow_outerPaddingY * 2)
End Sub

' (VB) 180 - static_cast<uint8_t>((float(delta) / 500.f) * 180), as a uint8_t
Private Function flashAlphaValue(ByVal delta As Long) As Long
    flashAlphaValue = (180 - (CLng(Fix(CSng(CSng(delta) / 500!) * 180!)) And &HFF&)) And &HFF&
End Function

Public Sub OptionsSectionRow_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As OptionsSectionRow, x0 As Long, y0 As Long, x1 As Long, y1 As Long, delta As Long, flashColor As Long
    Dim bottomBorderX0 As Long, bottomBorderX1 As Long, bottomBorderY0 As Long, bottomBorderY1 As Long, lineColor As Long, b As Long
    Set d = OptionsSectionRow_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    If d.has_flashHighlightStart Then
        delta = realTime - d.flashHighlightStart
        If ULT(delta, 500) Then
            flashColor = psPalette(WZCOL_MENU_SCORE_BUILT)
            flashColor = pielightWithA(flashColor, flashAlphaValue(delta))
            If d.showBottomBorder Then b = 1 Else b = 0
            pie_UniTransBoxFill x0, y0, x1 - 1, y1 - b, flashColor
        Else
            d.has_flashHighlightStart = False
        End If
    End If

    If d.showBottomBorder Then
        bottomBorderX0 = x0 + OptionsSectionRow_outerPaddingX
        bottomBorderX1 = x1 - OptionsSectionRow_outerPaddingX
        bottomBorderY0 = y1 - 1
        bottomBorderY1 = bottomBorderY0 + 1
        lineColor = psPalette(WZCOL_TEXT_MEDIUM)
        lineColor = pielightWithA(lineColor, pielightA(lineColor) \ 2)
        pie_UniTransBoxFill bottomBorderX0, bottomBorderY0, bottomBorderX1, bottomBorderY1, lineColor
    End If
End Sub

Public Sub OptionsSectionRow_flashHighlight(ByVal self As WIDGET)
    Dim d As OptionsSectionRow
    Set d = OptionsSectionRow_(self)
    d.has_flashHighlightStart = True
    d.flashHighlightStart = realTime
End Sub

' MARK: OptionRow

Public Sub OptionRow_construct(ByVal self As WIDGET)
    Dim d As OptionRow
    OptionsRowWidgetBase_construct self
    self.kind = K_OptionRow
    Set d = OptionRow_(self)
    d.m_indentLevel = 0
    d.has_flashHighlightStart = False   ' = nullopt
    d.showBottomBorder = False
    d.isHighlighted = False
End Sub

' static std::shared_ptr<OptionRow> make(const WzString& optionId, const std::shared_ptr<WzOptionsLabelWidget>& labelWidget, const std::shared_ptr<OptionsValueChangerWidgetWrapper>& wrappedOptionValueChangerWidget, int16_t indentLevel = 0)
Public Function OptionRow_make(ByVal optionId As String, ByVal labelWidget As WIDGET, ByVal wrappedOptionValueChangerWidget As WIDGET, Optional ByVal indentLevel As Long = 0) As WIDGET
    Dim result As WIDGET, d As OptionRow
    Set result = WIDGET_alloc(K_OptionRow)
    OptionRow_construct result
    Set d = OptionRow_(result)
    d.m_optionId = optionId
    d.m_indentLevel = indentLevel

    Set d.m_labelWidget = labelWidget
    result.attach d.m_labelWidget

    Set d.m_wrappedOptionValueChangerWidget = wrappedOptionValueChangerWidget
    result.attach d.m_wrappedOptionValueChangerWidget

    Set OptionRow_make = result
End Function

Public Sub OptionRow_setShowBottomBorder(ByVal self As WIDGET, ByVal show As Boolean)
    OptionRow_(self).showBottomBorder = show
End Sub

Public Function OptionRow_optionId(ByVal self As WIDGET) As String
    OptionRow_optionId = OptionRow_(self).m_optionId
End Function

Public Function OptionRow_getLabelWidget(ByVal self As WIDGET) As WIDGET
    Set OptionRow_getLabelWidget = OptionRow_(self).m_labelWidget
End Function

Public Function OptionRow_getWrappedOptionValueChangerWidget(ByVal self As WIDGET) As WIDGET
    Set OptionRow_getWrappedOptionValueChangerWidget = OptionRow_(self).m_wrappedOptionValueChangerWidget
End Function

Private Function OptionRow_leftIndent(ByVal self As WIDGET) As Long
    OptionRow_leftIndent = OptionRow_(self).m_indentLevel * OptionRow_indentPaddingPerLevel
End Function

Public Function OptionRow_idealWidth(ByVal self As WIDGET) As Long
    Dim d As OptionRow
    Set d = OptionRow_(self)
    OptionRow_idealWidth = (OptionRow_outerPaddingX * 2) + OptionRow_leftIndent(self) + MAX_(d.m_labelWidget.idealWidth(), OptionRow_minimumLabelWidth) + OptionRow_innerPaddingX + d.m_wrappedOptionValueChangerWidget.idealWidth()
End Function

Public Function OptionRow_idealHeight(ByVal self As WIDGET) As Long
    Dim d As OptionRow, b As Long
    Set d = OptionRow_(self)
    If d.showBottomBorder Then b = 1 Else b = 0
    OptionRow_idealHeight = (OptionRow_outerPaddingY * 2) + MAX_(d.m_labelWidget.idealHeight(), d.m_wrappedOptionValueChangerWidget.idealHeight()) + b
End Function

Public Sub OptionRow_geometryChanged(ByVal self As WIDGET)
    Dim d As OptionRow, w As Long, h As Long, b As Long, leftIndentWidth As Long, availableWidthForWidgets As Long, availableWidgetHeight As Long
    Dim labelWidth As Long, valueChangerWidth As Long, shrinkByTotalWidth As Long, shrinkLabelByWidth As Long, a As Long
    Dim labelX0 As Long, labelHeight As Long, labelY0 As Long, valueChangerX0 As Long, valueChangerHeight As Long, valueChangerY0 As Long
    Set d = OptionRow_(self)
    w = self.width()
    h = self.height()

    If w = 0 Or h = 0 Then
        Exit Sub
    End If

    If d.showBottomBorder Then b = 1 Else b = 0
    leftIndentWidth = OptionRow_leftIndent(self)
    availableWidthForWidgets = w - (OptionRow_outerPaddingX * 2) - leftIndentWidth - OptionRow_innerPaddingX
    availableWidgetHeight = h - (OptionRow_outerPaddingY * 2) - b

    labelWidth = d.m_labelWidget.idealWidth()
    valueChangerWidth = d.m_wrappedOptionValueChangerWidget.idealWidth()

    If labelWidth + valueChangerWidth > availableWidthForWidgets Then
        ' Must shrink label widget below its ideal width
        shrinkByTotalWidth = (labelWidth + valueChangerWidth) - availableWidthForWidgets
        If labelWidth > OptionRow_minimumLabelWidth Then a = labelWidth - OptionRow_minimumLabelWidth Else a = 0
        shrinkLabelByWidth = MIN_(shrinkByTotalWidth, a)
        labelWidth = labelWidth - shrinkLabelByWidth
        shrinkByTotalWidth = shrinkByTotalWidth - shrinkLabelByWidth
        If shrinkByTotalWidth > 0 Then
            ' Must also shrink the valueChangerWidget?
            valueChangerWidth = valueChangerWidth - shrinkByTotalWidth
            If valueChangerWidth < OptionRow_minimumLabelWidth Then
                valueChangerWidth = OptionRow_minimumLabelWidth
                debug_ 1234, LOG_WZ, "geometryChanged", "Dimensions too small"
            End If
        End If
    End If

    ' label is left-aligned, with left padding based on indentLevel
    labelX0 = OptionRow_outerPaddingX + leftIndentWidth
    labelHeight = MIN_(d.m_labelWidget.idealHeight(), availableWidgetHeight)
    labelY0 = OptionRow_outerPaddingY + ((availableWidgetHeight - labelHeight) \ 2)
    d.m_labelWidget.setGeometry labelX0, labelY0, labelWidth, labelHeight

    ' valueChanger is right-aligned
    valueChangerX0 = w - OptionRow_outerPaddingX - valueChangerWidth
    valueChangerHeight = MIN_(d.m_wrappedOptionValueChangerWidget.idealHeight(), availableWidgetHeight)
    valueChangerY0 = OptionRow_outerPaddingY + ((availableWidgetHeight - valueChangerHeight) \ 2)
    d.m_wrappedOptionValueChangerWidget.setGeometry valueChangerX0, valueChangerY0, valueChangerWidth, valueChangerHeight
End Sub

Public Sub OptionRow_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As OptionRow, x0 As Long, y0 As Long, x1 As Long, y1 As Long, b As Long, highlightColor As Long, delta As Long, flashColor As Long
    Dim bottomBorderX0 As Long, bottomBorderX1 As Long, bottomBorderY0 As Long, bottomBorderY1 As Long, lineColor As Long
    Set d = OptionRow_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    x1 = x0 + self.width()
    y1 = y0 + self.height()
    If d.showBottomBorder Then b = 1 Else b = 0

    If d.isHighlighted Then
        highlightColor = psPalette(WZCOL_MENU_SCORE_BUILT)
        highlightColor = pielightWithA(highlightColor, pielightA(highlightColor) \ 2)
        pie_UniTransBoxFill x0, y0, x1 - 1, y1 - b, highlightColor
    End If

    If d.has_flashHighlightStart Then
        delta = realTime - d.flashHighlightStart
        If ULT(delta, 500) Then
            flashColor = pal_RGBA(16, 150, 2, 255)
            flashColor = pielightWithA(flashColor, flashAlphaValue(delta))
            pie_UniTransBoxFill x0, y0, x1 - 1, y1 - b, flashColor
        Else
            d.has_flashHighlightStart = False
        End If
    End If

    If d.showBottomBorder Then
        bottomBorderX0 = x0 + OptionRow_outerPaddingX
        bottomBorderX1 = x1 - OptionRow_outerPaddingX
        bottomBorderY0 = y1 - 1
        bottomBorderY1 = bottomBorderY0 + 1
        lineColor = psPalette(WZCOL_TEXT_MEDIUM)
        lineColor = pielightWithA(lineColor, pielightA(lineColor) \ 2)
        pie_UniTransBoxFill bottomBorderX0, bottomBorderY0, bottomBorderX1, bottomBorderY1, lineColor
    End If
End Sub

Public Sub OptionRow_setParentOptionsForm(ByVal self As WIDGET, ByVal parentOptionsForm As WIDGET)
    OptionRow_(self).m_weakParentForm = ObjPtr(parentOptionsForm)
End Sub

Public Sub OptionRow_setIsHighlighted(ByVal self As WIDGET, ByVal newValue As Boolean)
    Dim d As OptionRow
    Set d = OptionRow_(self)
    d.isHighlighted = newValue
    WzOptionsChoiceWidget_setDisplayHighlighted d.m_labelWidget, d.isHighlighted
End Sub

Public Sub OptionRow_flashHighlight(ByVal self As WIDGET)
    Dim d As OptionRow
    Set d = OptionRow_(self)
    d.has_flashHighlightStart = True
    d.flashHighlightStart = realTime
End Sub

' MARK: - OptionsForm

Public Function OptionsForm_make() As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_OptionsForm)
    OptionsForm_construct result
    OptionsForm_initialize result
    Set OptionsForm_make = result
End Function

' OptionsForm::OptionsForm() (with the member initialisers)
Public Sub OptionsForm_construct(ByVal self As WIDGET)
    Dim d As OptionsForm
    W_FORM_construct self
    self.kind = K_OptionsForm
    Set d = OptionsForm_(self)
    Set d.optionDetailsMap = CreateObject("Scripting.Dictionary")
    d.optionDetailsMap.CompareMode = vbBinaryCompare
    Set d.sectionDetailsMap = CreateObject("Scripting.Dictionary")
    d.sectionDetailsMap.CompareMode = vbBinaryCompare
    d.orderedSectionIdsCount = 0
    d.has_mouseOverFormPosThisFrame = False  ' = nullopt
    d.maxListWidth = 0
    d.refreshOptionsOnScreenSizeChange = False
    d.isInRefreshOptionsInternal = False
    d.weakSerial = WeakWidget_register(self)     ' (VB) the std::weak_ptr serial (widgweak_vb)
End Sub

Public Sub OptionsForm_destruct(ByVal self As WIDGET)
    Dim d As OptionsForm
    Set d = OptionsForm_(self)
    If Not d.currentHelpPopoverWidget Is Nothing Then
        PopoverWidget_close d.currentHelpPopoverWidget
        Set d.currentHelpPopoverWidget = Nothing
    End If
    WeakWidget_unregister d.weakSerial
End Sub

Public Sub OptionsForm_initialize(ByVal self As WIDGET)
    Dim d As OptionsForm, listPadding As Padding
    Set d = OptionsForm_(self)
    Set d.optionsList = ScrollableListWidget_make()
    listPadding.top = 1
    listPadding.right_ = 0
    listPadding.bottom = 1
    listPadding.left_ = 0
    ScrollableListWidget_setPadding d.optionsList, listPadding
    ScrollableListWidget_setItemSpacing d.optionsList, 0
    self.attach d.optionsList
End Sub

Public Function OptionsForm_idealWidth(ByVal self As WIDGET) As Long
    OptionsForm_idealWidth = OptionsForm_(self).optionsList.idealWidth()
End Function

Public Function OptionsForm_idealHeight(ByVal self As WIDGET) As Long
    OptionsForm_idealHeight = OptionsForm_(self).optionsList.idealHeight()
End Function

' MARK: WzHelpPopoverWidget

' static std::shared_ptr<WzHelpPopoverWidget> make(const OptionInfo& optionInfo, const std::vector<OptionChoiceHelpDescription>& choiceHelpDescriptions, int32_t paragraphWidth, iV_fonts baseFontId = font_small)
Public Function WzHelpPopoverWidget_make(ByVal optionInfo As OptionInfo, choiceHelpDescriptions() As OptionChoiceHelpDescription, ByVal choiceHelpDescriptionsCount As Long, ByVal paragraphWidth As Long, Optional ByVal baseFontId As iV_fonts = font_small) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WzHelpPopoverWidget)
    WzHelpPopoverWidget_construct result
    If Not WzHelpPopoverWidget_initialize(result, optionInfo, choiceHelpDescriptions, choiceHelpDescriptionsCount, paragraphWidth, baseFontId) Then
        Exit Function                       ' nullptr
    End If
    result.setGeometry 0, 0, result.idealWidth(), result.idealHeight()
    Set WzHelpPopoverWidget_make = result
End Function

Public Sub WzHelpPopoverWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzHelpPopoverWidget
End Sub

Private Function WzHelpPopoverWidget_initialize(ByVal self As WIDGET, ByVal optionInfo As OptionInfo, choiceHelpDescriptions() As OptionChoiceHelpDescription, ByVal choiceHelpDescriptionsCount As Long, ByVal paragraphWidth As Long, ByVal baseFontId As iV_fonts) As Boolean
    Dim d As WzHelpPopoverWidget, baseFont As iV_fonts, baseFontBold As iV_fonts, availabilityResults() As OptionInfo_AvailabilityResult
    Dim availabilityResultsCount As Long, optionIsAvailable As Boolean, wroteALine As Boolean, optionHelpDescription As String
    Dim i As Long, result As OptionInfo_AvailabilityResult
    Set d = WzHelpPopoverWidget_(self)
    If Not iV_FontModifyBold(baseFontId, False, baseFont) Then baseFont = baseFontId
    If Not iV_FontModifyBold(baseFontId, True, baseFontBold) Then baseFontBold = baseFontId

    availabilityResultsCount = optionInfo.getAvailabilityResults(availabilityResults)
    optionIsAvailable = True
    For i = 0 To availabilityResultsCount - 1
        If Not availabilityResults(i).available_ Then
            optionIsAvailable = False
            Exit For
        End If
    Next

    ' Build description paragraph
    Set d.paragraph = Paragraph_new()
    wroteALine = False
    Paragraph_setLineSpacing d.paragraph, 1

    ' Add general option info text description
    Paragraph_setFont d.paragraph, baseFont
    Paragraph_setFontColour d.paragraph, psPalette(WZCOL_TEXT_BRIGHT)
    optionHelpDescription = optionInfo.getTranslatedHelpDescription()
    If Len(optionHelpDescription) <> 0 Then
        Paragraph_addText d.paragraph, optionHelpDescription
        wroteALine = True
    End If

    If optionIsAvailable Then
        For i = 0 To choiceHelpDescriptionsCount - 1
            If Len(choiceHelpDescriptions(i).helpDescription) = 0 Then
                GoTo continue_
            End If

            If wroteALine Then
                Paragraph_addText d.paragraph, vbLf & " " & vbLf
            End If

            ' Add a bold string with the displayName and then append the choice help description
            If Len(choiceHelpDescriptions(i).displayName) <> 0 Then
                Paragraph_setFont d.paragraph, baseFontBold
                Paragraph_addText d.paragraph, choiceHelpDescriptions(i).displayName & ": "
            End If
            Paragraph_setFont d.paragraph, baseFont
            Paragraph_addText d.paragraph, choiceHelpDescriptions(i).helpDescription

            wroteALine = True
continue_:
        Next

        If optionInfo.requiresRestart() Then
            If wroteALine Then
                Paragraph_addText d.paragraph, vbLf & " " & vbLf
            End If

            Paragraph_setFont d.paragraph, baseFontBold
            Paragraph_addText d.paragraph, "* " & "Takes effect on game restart"
            wroteALine = True

            Paragraph_setFont d.paragraph, baseFont
        End If
    Else
        For i = 0 To availabilityResultsCount - 1
            Set result = availabilityResults(i)
            If result.available_ Then
                GoTo continue2_
            End If

            If wroteALine Then
                Paragraph_addText d.paragraph, vbLf & " " & vbLf
            End If

            ' Add a bold string with the
            If Len(result.localizedUnavailabilityReason) <> 0 Then
                Paragraph_setFont d.paragraph, baseFontBold
                Paragraph_addText d.paragraph, result.localizedUnavailabilityReason
                wroteALine = True
            End If
continue2_:
        Next
    End If

    If Not wroteALine Then
        WzHelpPopoverWidget_initialize = False
        Exit Function
    End If

    self.attach d.paragraph
    d.paragraph.setGeometry 0, 0, paragraphWidth, 100      ' paragraph calculates actual height based on width

    WzHelpPopoverWidget_initialize = True
End Function

Public Function WzHelpPopoverWidget_idealWidth(ByVal self As WIDGET) As Long
    WzHelpPopoverWidget_idealWidth = (WzHelpPopoverWidget_outerPadding * 2) + WzHelpPopoverWidget_(self).paragraph.width()
End Function

Public Function WzHelpPopoverWidget_idealHeight(ByVal self As WIDGET) As Long
    WzHelpPopoverWidget_idealHeight = (WzHelpPopoverWidget_outerPadding * 2) + WzHelpPopoverWidget_(self).paragraph.height()
End Function

Public Sub WzHelpPopoverWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    ' display background
    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), pal_RGBA(20, 20, 20, 200)
End Sub

Public Sub WzHelpPopoverWidget_geometryChanged(ByVal self As WIDGET)
    Dim w As Long, h As Long, desiredParagraphWidth As Long, maxParagraphHeight As Long
    w = self.width()
    h = self.height()

    desiredParagraphWidth = w - (WzHelpPopoverWidget_outerPadding * 2)
    maxParagraphHeight = h - (WzHelpPopoverWidget_outerPadding * 2)
    WzHelpPopoverWidget_(self).paragraph.setGeometry WzHelpPopoverWidget_outerPadding, WzHelpPopoverWidget_outerPadding, desiredParagraphWidth, maxParagraphHeight
End Sub

Private Sub OptionsForm_openHelpPopover(ByVal self As WIDGET, ByVal optionInfo As OptionInfo, choiceHelpDescriptions() As OptionChoiceHelpDescription, ByVal choiceHelpDescriptionsCount As Long, ByVal parent As WIDGET)
    Dim paragraphWidth As Long, helpWidget As WIDGET
    OptionsForm_closeHelpPopover self

    paragraphWidth = MIN_(240, MAX_(self.width() - 100, 240))
    Set helpWidget = WzHelpPopoverWidget_make(optionInfo, choiceHelpDescriptions, choiceHelpDescriptionsCount, paragraphWidth)
    If helpWidget Is Nothing Then
        ' nothing to display
        Exit Sub
    End If

    ' Then open a popover with the paragraph as the widget
    Set OptionsForm_(self).currentHelpPopoverWidget = PopoverWidget_makePopover(parent, helpWidget, PopoverWidget_Style_NonInteractive, PopoverWidget_Alignment_RightOfParent, 0, 10)
End Sub

Private Sub OptionsForm_closeHelpPopover(ByVal self As WIDGET)
    Dim d As OptionsForm
    Set d = OptionsForm_(self)
    If d.currentHelpPopoverWidget Is Nothing Then
        Exit Sub
    End If
    PopoverWidget_close d.currentHelpPopoverWidget
    Set d.currentHelpPopoverWidget = Nothing
End Sub

Private Sub OptionsForm_handleStartedRowEditing(ByVal self As WIDGET, ByVal optionId As String)
    Dim d As OptionsForm, row As WIDGET, od As OptionsForm_OptionDetails
    Set d = OptionsForm_(self)
    OptionsForm_closeHelpPopover self

    If Not (d.optionDetailsMap.Exists(optionId)) Then ASSERT_FAILURE "it != optionDetailsMap.end()", "src/titleui/widgets/optionsform.cpp:1526", "handleStartedRowEditing", "Could not find OptionRow for: %s", optionId: Exit Sub
    od = d.optionDetails(d.optionDetailsMap(optionId))
    Set row = ScrollableListWidget_getItemAtIdx(d.optionsList, od.idxInOptionsList)
    If Not WIDGET_isKindOf(row, K_OptionRow) Then Set row = Nothing
    If Not (Not row Is Nothing) Then ASSERT_FAILURE "row != nullptr", "src/titleui/widgets/optionsform.cpp:1528", "handleStartedRowEditing", "Could not find OptionRow for: %s", optionId: Exit Sub

    Set d.capturedEditingRowWidget = row
End Sub

Private Sub OptionsForm_handleEndedRowEditing(ByVal self As WIDGET, ByVal optionId As String)
    Dim d As OptionsForm, row As WIDGET, od As OptionsForm_OptionDetails
    Set d = OptionsForm_(self)
    If Not (d.optionDetailsMap.Exists(optionId)) Then ASSERT_FAILURE "it != optionDetailsMap.end()", "src/titleui/widgets/optionsform.cpp:1536", "handleEndedRowEditing", "Could not find OptionRow for: %s", optionId: Exit Sub
    od = d.optionDetails(d.optionDetailsMap(optionId))
    Set row = ScrollableListWidget_getItemAtIdx(d.optionsList, od.idxInOptionsList)
    If Not WIDGET_isKindOf(row, K_OptionRow) Then Set row = Nothing
    If Not (Not row Is Nothing) Then ASSERT_FAILURE "row != nullptr", "src/titleui/widgets/optionsform.cpp:1538", "handleEndedRowEditing", "Could not find OptionRow for: %s", optionId: Exit Sub

    If Not row Is d.capturedEditingRowWidget Then
        debug_ 1542, LOG_INFO, "handleEndedRowEditing", "Ignorning call for non-current captured editing row widget: %s", optionId
    End If

    Set d.capturedEditingRowWidget = Nothing
End Sub

Public Sub OptionsForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' no-op
End Sub

Public Function OptionsForm_hasOptionsThatRequireRestart(ByVal self As WIDGET) As Boolean
    Dim d As OptionsForm, i As Long, od As OptionsForm_OptionDetails
    Set d = OptionsForm_(self)
    For i = 0 To d.optionDetailsCount - 1
        od = d.optionDetails(i)
        If od.info.requiresRestart() Then
            OptionsForm_hasOptionsThatRequireRestart = True
            Exit Function
        End If
    Next
    OptionsForm_hasOptionsThatRequireRestart = False
End Function

Public Function OptionsForm_jumpToSectionId(ByVal self As WIDGET, ByVal sectionId As String) As Boolean
    Dim d As OptionsForm, idx As Long, sectionRow As WIDGET, sd As OptionsForm_SectionDetails
    Set d = OptionsForm_(self)
    If Not d.sectionDetailsMap.Exists(sectionId) Then
        OptionsForm_jumpToSectionId = False
        Exit Function
    End If
    sd = d.sectionDetails(d.sectionDetailsMap(sectionId))
    idx = sd.idxInOptionsList
    ScrollableListWidget_scrollToItem d.optionsList, idx
    Set sectionRow = ScrollableListWidget_getItemAtIdx(d.optionsList, idx)
    If WIDGET_isKindOf(sectionRow, K_OptionsSectionRow) Then
        OptionsSectionRow_flashHighlight sectionRow
    End If
    OptionsForm_jumpToSectionId = True
End Function

Public Function OptionsForm_jumpToOptionId(ByVal self As WIDGET, ByVal optionId As String) As Boolean
    Dim d As OptionsForm, od As OptionsForm_OptionDetails
    Set d = OptionsForm_(self)
    If Not d.optionDetailsMap.Exists(optionId) Then
        OptionsForm_jumpToOptionId = False
        Exit Function
    End If
    od = d.optionDetails(d.optionDetailsMap(optionId))
    OptionsForm_jumpToItemIdx self, od.idxInOptionsList, True
    OptionsForm_jumpToOptionId = True
End Function

Public Function OptionsForm_numSections(ByVal self As WIDGET) As Long
    OptionsForm_numSections = OptionsForm_(self).optionDetailsMap.Count
End Function

' std::vector<OptionsSection> getSectionInfos() const: the count, the sections in result(0 .. count - 1)
Public Function OptionsForm_getSectionInfos(ByVal self As WIDGET, result() As OptionsSection) As Long
    Dim d As OptionsForm, i As Long, n As Long, sectionId As String, sd As OptionsForm_SectionDetails
    Set d = OptionsForm_(self)
    n = 0
    For i = 0 To d.orderedSectionIdsCount - 1
        sectionId = d.orderedSectionIds(i)
        If Not d.sectionDetailsMap.Exists(sectionId) Then
            GoTo continue_
        End If
        sd = d.sectionDetails(d.sectionDetailsMap(sectionId))
        If n = 0 Then ReDim result(0 To 0) Else ReDim Preserve result(0 To n)
        result(n) = sd.info
        n = n + 1
continue_:
    Next
    OptionsForm_getSectionInfos = n
End Function

Public Sub OptionsForm_setMaxListWidth(ByVal self As WIDGET, ByVal maxListWidth_ As Long)
    OptionsForm_(self).maxListWidth = maxListWidth_
    self.geometryChanged
End Sub

Public Sub OptionsForm_setRefreshOptionsOnScreenSizeChange(ByVal self As WIDGET, ByVal enabled As Boolean)
    OptionsForm_(self).refreshOptionsOnScreenSizeChange = enabled
End Sub

Private Function OptionsForm_jumpToItemIdx(ByVal self As WIDGET, ByVal idx As Long, ByVal flash As Boolean) As Boolean
    Dim d As OptionsForm, row As WIDGET
    Set d = OptionsForm_(self)
    ScrollableListWidget_scrollEnsureItemVisible d.optionsList, idx
    If flash Then
        Set row = ScrollableListWidget_getItemAtIdx(d.optionsList, idx)
        If WIDGET_isKindOf(row, K_OptionRow) Then
            OptionRow_flashHighlight row
        End If
    End If
    OptionsForm_jumpToItemIdx = True
End Function

Public Sub OptionsForm_addSection(ByVal self As WIDGET, sectionInfo As OptionsSection, Optional ByVal bottomBorder As Boolean = False)
    Dim d As OptionsForm, optionsSectionWidget As WIDGET, sd As OptionsForm_SectionDetails, itemListIdx As Long, sectionIdx As Long
    Set d = OptionsForm_(self)
    If ScrollableListWidget_numItems(d.optionsList) > 0 Then
        ' Add "padding" row above the section row
        ScrollableListWidget_addItem d.optionsList, OptionsPaddingRow_make(20)
    End If

    Set optionsSectionWidget = OptionsSectionRow_make(OptionsSection_getTranslatedDisplayName(sectionInfo))
    OptionsSectionRow_setShowBottomBorder optionsSectionWidget, bottomBorder

    ' Store mapping
    If Not (Not d.sectionDetailsMap.Exists(OptionsSection_sectionId(sectionInfo))) Then ASSERT_FAILURE "insertResult.second", "src/titleui/widgets/optionsform.cpp:1649", "addSection", "Option section already added to OptionsForm: %s", OptionsSection_sectionId(sectionInfo): Exit Sub
    sd.info = sectionInfo
    Set sd.optionSectionWidget = optionsSectionWidget
    sd.idxInOptionsList = 0
    sectionIdx = d.sectionDetails_push_back(sd)
    d.sectionDetailsMap.Add OptionsSection_sectionId(sectionInfo), sectionIdx
    d.orderedSectionIds_push_back OptionsSection_sectionId(sectionInfo)

    ' Add options section widget to list
    itemListIdx = ScrollableListWidget_addItem(d.optionsList, optionsSectionWidget)
    sd.idxInOptionsList = itemListIdx
    d.sectionDetails(sectionIdx) = sd
End Sub

Private Function getOptionInfoTranslatedDisplayName(ByVal optionInfo As OptionInfo) As String
    Dim result As String
    result = optionInfo.getTranslatedDisplayName()
    If optionInfo.requiresRestart() Then
        result = result & "*"
    End If
    getOptionInfoTranslatedDisplayName = result
End Function

' template<typename T> addOption(const OptionInfo& optionInfo, const std::shared_ptr<T>& optionValueChangerWidget, bool bottomBorder = false, int16_t indentLevel = 0)
Public Sub OptionsForm_addOption(ByVal self As WIDGET, ByVal optionInfo As OptionInfo, ByVal optionValueChangerWidget As WIDGET, Optional ByVal bottomBorder As Boolean = False, Optional ByVal indentLevel As Long = 0)
    OptionsForm_addOptionInternal self, optionInfo, optionValueChangerWidget, bottomBorder, indentLevel
End Sub

Private Sub OptionsForm_addOptionInternal(ByVal self As WIDGET, ByVal optionInfo As OptionInfo, ByVal optionValueChangerWidget As WIDGET, Optional ByVal bottomBorder As Boolean = False, Optional ByVal indentLevel As Long = 0)
    Dim d As OptionsForm, pWeakOptionsForm As Long, optionLabelWidget As WIDGET, optionIdCopy As String
    Dim wrappedOptionValueChangerWidget As WIDGET, isAvailable As Boolean, od As OptionsForm_OptionDetails
    Dim row As WIDGET, itemListIdx As Long, detailsIdx As Long
    Set d = OptionsForm_(self)
    pWeakOptionsForm = ObjPtr(self)

    ' Create left side widget for displaying the Option label / display name
    Set optionLabelWidget = WzOptionsChoiceWidget_make(font_regular)     ' WzOptionsLabelWidget::make
    optionLabelWidget.setString getOptionInfoTranslatedDisplayName(optionInfo)
    WzOptionsChoiceWidget_setPadding optionLabelWidget, 4, 8
    WzOptionsChoiceWidget_setTextAlignment optionLabelWidget, WLAB_ALIGNLEFT
    optionLabelWidget.setGeometry 0, 0, optionLabelWidget.idealWidth(), optionLabelWidget.idealHeight()
    WzOptionsChoiceWidget_setHighlightBackgroundColor optionLabelWidget, pal_RGBA(0, 0, 0, 0)

    ' Add an onChangeHandler to the optionValueChangerWidget
    If Not (vt_WIDGET_ovc_is(optionValueChangerWidget)) Then ASSERT_FAILURE "pOptionValueChangerInterface != nullptr", "src/titleui/widgets/optionsform.cpp:1681", "addOptionInternal", "optionValueChangerWidget does not implement appropriate interface": Exit Sub
    optionIdCopy = optionInfo.optionId
    vt_WIDGET_ovc_addOnChangeHandler optionValueChangerWidget, WzFunc_new(FN_OptionsForm_onOptionChange, pWeakOptionsForm, optionIdCopy)
    vt_WIDGET_ovc_addStartingCapturedEditingHandler optionValueChangerWidget, WzFunc_new(FN_OptionsForm_onStartedEditing, pWeakOptionsForm, optionIdCopy)
    vt_WIDGET_ovc_addEndingCapturedEditingHandler optionValueChangerWidget, WzFunc_new(FN_OptionsForm_onEndedEditing, pWeakOptionsForm, optionIdCopy)

    ' Wrap the optionValueChangerWidget in an OptionsValueChangerWidgetWrapper
    Set wrappedOptionValueChangerWidget = OptionsValueChangerWidgetWrapper_wrap(optionValueChangerWidget)
    isAvailable = optionInfo.isAvailable()
    If Not isAvailable Then
        OptionsValueChangerWidgetWrapper_disable wrappedOptionValueChangerWidget
    End If
    vt_WIDGET_ovc_informAvailable optionValueChangerWidget, isAvailable

    ' Store mapping
    If Not (Not d.optionDetailsMap.Exists(optionInfo.optionId)) Then ASSERT_FAILURE "insertResult.second", "src/titleui/widgets/optionsform.cpp:1710", "addOptionInternal", "Option id already added to OptionsForm: %s", optionInfo.optionId: Exit Sub
    Set od.info = optionInfo
    Set od.optionLabelWidget = optionLabelWidget
    Set od.optionValueChangerWidgetWrapper = wrappedOptionValueChangerWidget
    od.idxInOptionsList = 0
    detailsIdx = d.optionDetails_push_back(od)
    d.optionDetailsMap.Add optionInfo.optionId, detailsIdx

    ' Add both widgets to an OptionRow, insert into list
    Set row = OptionRow_make(optionInfo.optionId, optionLabelWidget, wrappedOptionValueChangerWidget, indentLevel)
    row.setGeometry 0, 0, row.idealWidth(), row.idealHeight()
    If bottomBorder Then
        OptionRow_setShowBottomBorder row, True
    End If
    itemListIdx = ScrollableListWidget_addItem(d.optionsList, row)
    od.idxInOptionsList = itemListIdx
    d.optionDetails(detailsIdx) = od
End Sub

Public Function OptionsForm_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim d As OptionsForm, shiftedContext As W_CONTEXT
    Set d = OptionsForm_(self)
    If key = WKEY_NONE Then ' called this way for highlight processing by WIDGET::findMouseTargetRecursive, only if the mouse is actually over this widget (i.e. if hit-testing passes)
        ' mouse is over the options form this frame
        shiftedContext = psContext
        shiftedContext.mx = psContext.mx - self.x()
        shiftedContext.my = psContext.my - self.y()
        shiftedContext.xOffset = psContext.xOffset + self.x()
        shiftedContext.yOffset = psContext.yOffset + self.y()
        d.has_mouseOverFormPosThisFrame = True ' consumed by run()
        d.mouseOverFormPosThisFrameX = shiftedContext.mx
        d.mouseOverFormPosThisFrameY = shiftedContext.my
    End If

    ' forward to parent
    Set OptionsForm_findMouseTargetRecursive = W_FORM_findMouseTargetRecursive(self, psContext, key, wasPressed)
End Function

Public Sub OptionsForm_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As OptionsForm, currentHoverWidget As WIDGET, listPadding As Padding
    Set d = OptionsForm_(self)
    Set currentHoverWidget = Nothing
    If Not d.capturedEditingRowWidget Is Nothing Then
        Set currentHoverWidget = d.capturedEditingRowWidget
    ElseIf d.has_mouseOverFormPosThisFrame Then
        listPadding = ScrollableListWidget_getPadding(d.optionsList)
        If (d.mouseOverFormPosThisFrameX >= listPadding.left_) And (d.mouseOverFormPosThisFrameX <= (d.optionsList.width() - listPadding.right_)) Then
            ' determine which row the mouse is currently over
            Set currentHoverWidget = ScrollableListWidget_getItemAtYPos(d.optionsList, d.mouseOverFormPosThisFrameY)
            If Not WIDGET_isKindOf(currentHoverWidget, K_OptionsRowWidgetBase) Then Set currentHoverWidget = Nothing
        Else
            Set currentHoverWidget = Nothing
        End If
        d.has_mouseOverFormPosThisFrame = False ' consume it
    End If

    If Not currentHoverWidget Is Nothing Then
        OptionsForm_handleMouseIsOverRow self, currentHoverWidget
    ElseIf Not d.priorMouseOverRowWidget Is Nothing Then
        OptionsForm_handleMouseIsOverRow self, Nothing
    End If
End Sub

Public Sub OptionsForm_geometryChanged(ByVal self As WIDGET)
    Dim d As OptionsForm, w As Long, h As Long, listHeight As Long, listPaddingX As Long, listPadding As Padding
    Set d = OptionsForm_(self)
    w = self.width()
    h = self.height()

    listHeight = MIN_(d.optionsList.idealHeight(), h)
    If d.maxListWidth > 0 And d.maxListWidth < w Then listPaddingX = (w - d.maxListWidth) \ 2 Else listPaddingX = 0
    listPadding.top = 0
    listPadding.right_ = listPaddingX
    listPadding.bottom = 0
    listPadding.left_ = listPaddingX
    ScrollableListWidget_setPadding d.optionsList, listPadding
    d.optionsList.setGeometry 0, 0, w, listHeight
End Sub

Public Sub OptionsForm_screenSizeDidChange(ByVal self As WIDGET, ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If OptionsForm_(self).refreshOptionsOnScreenSizeChange Then
        OptionsForm_refreshOptions self, True
    End If
    W_FORM_screenSizeDidChange self, oldWidth, oldHeight, newWidth, newHeight
End Sub

Private Sub OptionsForm_handleMouseIsOverRow(ByVal self As WIDGET, ByVal rowWidget As WIDGET)
    Dim d As OptionsForm, psOptionRowWidget As WIDGET, rowIsInCapturedEditMode As Boolean, od As OptionsForm_OptionDetails
    Dim wrappedOptionValueChangerWidget As WIDGET, pOptionValueChangerInterface As WIDGET
    Dim choiceHelpDescriptions() As OptionChoiceHelpDescription, choiceHelpDescriptionsCount As Long
    Set d = OptionsForm_(self)
    If rowWidget Is Nothing Then
        If Not d.priorMouseOverRowWidget Is Nothing Then
            vt_OptionsRowWidgetBase_setIsHighlighted d.priorMouseOverRowWidget, False
            Set d.priorMouseOverRowWidget = Nothing
        End If
        OptionsForm_closeHelpPopover self
        Exit Sub
    End If

    If Not rowWidget Is d.priorMouseOverRowWidget Then
        If Not d.priorMouseOverRowWidget Is Nothing Then
            vt_OptionsRowWidgetBase_setIsHighlighted d.priorMouseOverRowWidget, False
        End If
        vt_OptionsRowWidgetBase_setIsHighlighted rowWidget, True
        If WIDGET_isKindOf(rowWidget, K_OptionRow) Then Set psOptionRowWidget = rowWidget
        rowIsInCapturedEditMode = Not d.capturedEditingRowWidget Is Nothing
        If Not psOptionRowWidget Is Nothing And Not rowIsInCapturedEditMode Then
            ' it's an option row - display the help popover (as long as not in captured editing mode)
            If d.optionDetailsMap.Exists(OptionRow_optionId(psOptionRowWidget)) Then
                od = d.optionDetails(d.optionDetailsMap(OptionRow_optionId(psOptionRowWidget)))
                Set wrappedOptionValueChangerWidget = OptionRow_getWrappedOptionValueChangerWidget(psOptionRowWidget)
                Set pOptionValueChangerInterface = OptionsValueChangerWidgetWrapper_getOptionValueChangerWidget(wrappedOptionValueChangerWidget)
                choiceHelpDescriptionsCount = 0
                If vt_WIDGET_ovc_is(pOptionValueChangerInterface) Then
                    If Not vt_WIDGET_ovc_getHelpDescriptions(pOptionValueChangerInterface, choiceHelpDescriptions, choiceHelpDescriptionsCount) Then
                        choiceHelpDescriptionsCount = 0
                    End If
                End If
                OptionsForm_openHelpPopover self, od.info, choiceHelpDescriptions, choiceHelpDescriptionsCount, wrappedOptionValueChangerWidget
            Else
                OptionsForm_closeHelpPopover self
            End If
        Else
            OptionsForm_closeHelpPopover self
        End If
        Set d.priorMouseOverRowWidget = rowWidget
    End If
End Sub

Private Sub OptionsForm_handleOptionValueChanged(ByVal self As WIDGET, ByVal optionId As String)
    Dim d As OptionsForm, row As WIDGET, od As OptionsForm_OptionDetails
    Set d = OptionsForm_(self)
    ' Call update on all other options
    OptionsForm_refreshOptionsInternal self, optionId

    ' Update row layout (if needed)
    If d.optionDetailsMap.Exists(optionId) Then
        od = d.optionDetails(d.optionDetailsMap(optionId))
        Set row = ScrollableListWidget_getItemAtIdx(d.optionsList, od.idxInOptionsList)
        If WIDGET_isKindOf(row, K_OptionRow) Then
            row.geometryChanged
        End If
    End If
End Sub

' void refreshOptionsInternal(const std::unordered_set<WzString>& except): except is {} (Empty) or
' {optionId} (the option id) (VB: the only sets the C++ passes). The options are visited in the
' order they were added (C++: the unordered_map's order; each update is independent of the others').
Private Sub OptionsForm_refreshOptionsInternal(ByVal self As WIDGET, Optional ByVal except As Variant)
    Dim d As OptionsForm, currentTranslationLocale As String, changedTranslationLocale As Boolean, forceUpdate As Boolean
    Dim i As Long, sd As OptionsForm_SectionDetails, od As OptionsForm_OptionDetails, pSectionRow As WIDGET, pOptionLabel As WIDGET
    Dim isAvailable As Boolean, wrappedOptionValueChangerWidget As WIDGET, pOptionValueChangerInterface As WIDGET, row As WIDGET
    Set d = OptionsForm_(self)
    If d.isInRefreshOptionsInternal Then
        Exit Sub
    End If
    d.isInRefreshOptionsInternal = True
    currentTranslationLocale = getLanguage()
    changedTranslationLocale = StrComp(d.lastConfiguredTranslationLocale, currentTranslationLocale, vbBinaryCompare) <> 0
    If changedTranslationLocale Then
        d.lastConfiguredTranslationLocale = currentTranslationLocale

        For i = 0 To d.sectionDetailsCount - 1
            sd = d.sectionDetails(i)
            ' Update the section row text
            Set pSectionRow = sd.optionSectionWidget
            If WIDGET_isKindOf(pSectionRow, K_OptionsSectionRow) Then
                OptionsSectionRow_setSectionName pSectionRow, OptionsSection_getTranslatedDisplayName(sd.info)
            End If
        Next
    End If
    forceUpdate = changedTranslationLocale
    For i = 0 To d.optionDetailsCount - 1
        od = d.optionDetails(i)
        Set pOptionLabel = od.optionLabelWidget
        If WIDGET_isKindOf(pOptionLabel, K_WzOptionsChoiceWidget) And forceUpdate Then
            pOptionLabel.setString getOptionInfoTranslatedDisplayName(od.info)
        End If
        If Not IsMissing(except) Then
            If Not IsEmpty(except) Then
                If StrComp(CStr(except), od.info.optionId, vbBinaryCompare) = 0 Then
                    GoTo continue_
                End If
            End If
        End If
        isAvailable = od.info.isAvailable()
        Set wrappedOptionValueChangerWidget = od.optionValueChangerWidgetWrapper
        Set pOptionValueChangerInterface = OptionsValueChangerWidgetWrapper_getOptionValueChangerWidget(wrappedOptionValueChangerWidget)
        If vt_WIDGET_ovc_is(pOptionValueChangerInterface) Then
            vt_WIDGET_ovc_update pOptionValueChangerInterface, forceUpdate
            vt_WIDGET_ovc_informAvailable pOptionValueChangerInterface, isAvailable
        End If
        OptionsValueChangerWidgetWrapper_enable wrappedOptionValueChangerWidget, isAvailable
        If changedTranslationLocale Then
            Set row = ScrollableListWidget_getItemAtIdx(d.optionsList, od.idxInOptionsList)
            If WIDGET_isKindOf(row, K_OptionRow) Then
                row.geometryChanged
            End If
        End If
continue_:
    Next
    d.isInRefreshOptionsInternal = False
End Sub

Public Sub OptionsForm_refreshOptions(ByVal self As WIDGET, Optional ByVal forceRowLayoutUpdates As Boolean = False)
    Dim d As OptionsForm, i As Long, od As OptionsForm_OptionDetails, row As WIDGET
    Set d = OptionsForm_(self)
    OptionsForm_refreshOptionsInternal self, Empty
    If forceRowLayoutUpdates Then
        For i = 0 To d.optionDetailsCount - 1
            od = d.optionDetails(i)
            Set row = ScrollableListWidget_getItemAtIdx(d.optionsList, od.idxInOptionsList)
            If WIDGET_isKindOf(row, K_OptionRow) Then
                row.geometryChanged
            End If
        Next
    End If
End Sub

Public Sub OptionsForm_refreshOptionsLayoutCalc(ByVal self As WIDGET)
    Dim d As OptionsForm, i As Long, od As OptionsForm_OptionDetails, sd As OptionsForm_SectionDetails, row As WIDGET
    Set d = OptionsForm_(self)
    ' trigger screenSizeDidChange to ensure all widgets have a chance to respond to display scale changes that occurred while not part of the active widget hierarchy
    WIDGET_screenSizeDidChange self, screenWidth, screenHeight, screenWidth, screenHeight

    ' call geometryChanged() on rows to recalculate layout based on child widget ideal width (etc)
    For i = 0 To d.optionDetailsCount - 1
        od = d.optionDetails(i)
        Set row = ScrollableListWidget_getItemAtIdx(d.optionsList, od.idxInOptionsList)
        If WIDGET_isKindOf(row, K_OptionRow) Then
            row.geometryChanged
        End If
    Next
    For i = 0 To d.sectionDetailsCount - 1
        sd = d.sectionDetails(i)
        Set row = ScrollableListWidget_getItemAtIdx(d.optionsList, sd.idxInOptionsList)
        If WIDGET_isKindOf(row, K_OptionsSectionRow) Then
            row.geometryChanged
        End If
    Next
End Sub

Public Sub OptionsForm_refreshOptionAvailability(ByVal self As WIDGET)
    Dim d As OptionsForm, i As Long, od As OptionsForm_OptionDetails, isAvailable As Boolean
    Dim wrappedOptionValueChangerWidget As WIDGET, pOptionValueChangerInterface As WIDGET
    Set d = OptionsForm_(self)
    For i = 0 To d.optionDetailsCount - 1
        od = d.optionDetails(i)
        isAvailable = od.info.isAvailable()
        Set wrappedOptionValueChangerWidget = od.optionValueChangerWidgetWrapper
        OptionsValueChangerWidgetWrapper_enable wrappedOptionValueChangerWidget, isAvailable
        Set pOptionValueChangerInterface = OptionsValueChangerWidgetWrapper_getOptionValueChangerWidget(wrappedOptionValueChangerWidget)
        If vt_WIDGET_ovc_is(pOptionValueChangerInterface) Then
            vt_WIDGET_ovc_informAvailable pOptionValueChangerInterface, isAvailable
        End If
    Next
End Sub

' MARK: - OptionsDropdown<T> (optionsform.h: one kind for every T; the values are Variants)

Public Sub OptionsDropdown_construct(ByVal self As WIDGET)
    WzOptionsDropdownWidget_construct self
    self.kind = K_OptionsDropdown
End Sub

' static std::shared_ptr<OptionsDropdown<T>> make(PopulateFunc populate, OnSelectItemFunc onSelect, bool needsStateUpdates, int maxListItemsShown = 5)
' (populate returns an OptionChoices; onSelect(newValue) returns True if able to select the item)
Public Function OptionsDropdown_make(ByVal populate As WzFunc, ByVal onSelect As WzFunc, ByVal needsStateUpdates As Boolean, Optional ByVal maxListItemsShown As Long = 5) As WIDGET
    Dim result As WIDGET, d As OptionsDropdown
    If Not (Not populate Is Nothing) Then ASSERT_FAILURE "populate != nullptr", "src/titleui/widgets/optionsform.h:282", "make", "Invalid populate func": Exit Function
    Set result = WIDGET_alloc(K_OptionsDropdown)
    OptionsDropdown_construct result
    Set d = OptionsDropdown_(result)
    Set d.populateFunc = populate
    Set d.onSelectFunc = onSelect
    d.maxListItemsShown = maxListItemsShown
    d.needsStateUpdates = needsStateUpdates
    WzOptionsDropdownWidget_setTextAlignment result, WLAB_ALIGNRIGHT
    Set d.current = WzFunc_call(populate)
    OptionsDropdown_populateWithCurrentItems result
    DropdownWidget_setCanChange result, WzFunc_new(FN_OptionsDropdown_canChange)
    Set OptionsDropdown_make = result
End Function

Public Function OptionsDropdown_ovc_is(ByVal self As WIDGET) As Boolean
    OptionsDropdown_ovc_is = True
End Function

Public Sub OptionsDropdown_ovc_update(ByVal self As WIDGET, Optional ByVal force As Boolean = False)
    Dim d As OptionsDropdown, newItemsList As OptionChoices
    Set d = OptionsDropdown_(self)
    If Not force And Not d.needsStateUpdates Then
        Exit Sub
    End If
    Set newItemsList = WzFunc_call(d.populateFunc)
    If Not newItemsList.choicesEqual(d.current) Then
        Set d.current = newItemsList
        OptionsDropdown_populateWithCurrentItems self
    ElseIf newItemsList.currentIdx <> d.current.currentIdx Then
        d.current.currentIdx = newItemsList.currentIdx
        DropdownWidget_setSelectedIndex self, d.current.currentIdx
    End If
End Sub

Public Sub OptionsDropdown_ovc_informAvailable(ByVal self As WIDGET, ByVal isAvailable As Boolean)
    DropdownWidget_setDisabled self, Not isAvailable
End Sub

Public Sub OptionsDropdown_ovc_addOnChangeHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    ' in the future, call *DropdownWiget::addOnChange*
    DropdownWidget_setOnChange self, handler
End Sub

Public Sub OptionsDropdown_ovc_addStartingCapturedEditingHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    ' Since dropdowns create an overlay screen, add this to the onOpen handler
    DropdownWidget_setOnOpen self, handler
End Sub

Public Sub OptionsDropdown_ovc_addEndingCapturedEditingHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    DropdownWidget_setOnClose self, handler
End Sub

Public Function OptionsDropdown_ovc_getHelpDescriptionForCurrentValue(ByVal self As WIDGET, result As OptionChoiceHelpDescription) As Boolean
    Dim d As OptionsDropdown, currIndex As Long
    Set d = OptionsDropdown_(self)
    If Not DropdownWidget_getSelectedIndex(self, currIndex) Then
        OptionsDropdown_ovc_getHelpDescriptionForCurrentValue = False   ' nullopt
        Exit Function
    End If
    If Not (currIndex < d.current.count) Then ASSERT_FAILURE "currIndex.value() < current.choices.size()", "src/titleui/widgets/optionsform.h:368", "getHelpDescriptionForCurrentValue", "Invalid current index": OptionsDropdown_ovc_getHelpDescriptionForCurrentValue = False: Exit Function
    result.displayName = d.current.displayName(currIndex)
    result.helpDescription = d.current.helpDescription(currIndex)
    OptionsDropdown_ovc_getHelpDescriptionForCurrentValue = True
End Function

Public Function OptionsDropdown_ovc_getHelpDescriptions(ByVal self As WIDGET, result() As OptionChoiceHelpDescription, count As Long) As Boolean
    Dim d As OptionsDropdown, i As Long
    Set d = OptionsDropdown_(self)
    count = 0
    For i = 0 To d.current.count - 1
        If Len(d.current.helpDescription(i)) = 0 Then
            GoTo continue_
        End If
        If count = 0 Then ReDim result(0 To 0) Else ReDim Preserve result(0 To count)
        result(count).displayName = d.current.displayName(i)
        result(count).helpDescription = d.current.helpDescription(i)
        count = count + 1
continue_:
    Next
    OptionsDropdown_ovc_getHelpDescriptions = True
End Function

Private Sub OptionsDropdown_populateWithCurrentItems(ByVal self As WIDGET)
    Dim d As OptionsDropdown, maxItemHeight As Long, i As Long, item As WIDGET, listPadding As Padding, shown As Long
    Set d = OptionsDropdown_(self)
    DropdownWidget_clear self
    maxItemHeight = 0
    For i = 0 To d.current.count - 1
        Set item = WzOptionsChoiceWidget_make(font_regular)
        item.setString d.current.displayName(i)
        WzOptionsChoiceWidget_setIcon item, d.current.icon(i)
        WzOptionsChoiceWidget_setDisabled item, d.current.disabled(i)
        item.setGeometry 0, 0, item.idealWidth(), item.idealHeight()
        maxItemHeight = MAX_(maxItemHeight, item.idealHeight())
        DropdownWidget_addItem self, item
    Next
    listPadding = DropdownWidget_getDropdownMenuOuterPadding(self)
    shown = MIN_(d.maxListItemsShown, d.current.count)
    DropdownWidget_setListHeight self, maxItemHeight * shown + listPadding.top + listPadding.bottom
    DropdownWidget_setSelectedIndex self, d.current.currentIdx
End Sub

' MARK: - the lambdas

Public Function optionsform_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongParent As WIDGET, strongOptionsForm As WIDGET, strongSelf As WIDGET, newIndex As Long, onSelectResult As Boolean
    Dim d As OptionsDropdown, newSelectedWidget As WIDGET
    Select Case f.id
    Case FN_optionsform_displayBigSlider        ' static void displayBigSlider(WIDGET *psWidget, UDWORD xOffset, UDWORD yOffset)
        displayBigSlider args(0), CLng(args(1)), CLng(args(2))

    Case FN_IsNotInGame                         ' OptionInfo::AvailabilityResult IsNotInGame(const OptionInfo&)
        Set optionsform_c_fn = IsNotInGame(args(0))

    Case FN_IsNotMultiplayer                    ' OptionInfo::AvailabilityResult IsNotMultiplayer(const OptionInfo&)
        Set optionsform_c_fn = IsNotMultiplayer(args(0))

    Case FN_OptionsSlider_onSliderChange        ' OptionsSlider::make: [weakParent](W_SLIDER& s)
        Set strongParent = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/titleui/widgets/optionsform.cpp:95", "make", "Null parent": Exit Function
        OptionsSlider_handleSliderPosChanged strongParent

    Case FN_OptionsForm_onOptionChange          ' addOptionInternal: [pWeakOptionsForm, optionIdCopy](WIDGET&)
        Set strongOptionsForm = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongOptionsForm Is Nothing) Then ASSERT_FAILURE "strongOptionsForm != nullptr", "src/titleui/widgets/optionsform.cpp:1685", "addOptionInternal", "Invalid pointer": Exit Function
        OptionsForm_handleOptionValueChanged strongOptionsForm, CStr(f.cap(1))

    Case FN_OptionsForm_onStartedEditing        ' addOptionInternal: [pWeakOptionsForm, optionIdCopy](WIDGET&)
        Set strongOptionsForm = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongOptionsForm Is Nothing) Then ASSERT_FAILURE "strongOptionsForm != nullptr", "src/titleui/widgets/optionsform.cpp:1690", "addOptionInternal", "Invalid pointer": Exit Function
        OptionsForm_handleStartedRowEditing strongOptionsForm, CStr(f.cap(1))

    Case FN_OptionsForm_onEndedEditing          ' addOptionInternal: [pWeakOptionsForm, optionIdCopy](WIDGET&)
        Set strongOptionsForm = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongOptionsForm Is Nothing) Then ASSERT_FAILURE "strongOptionsForm != nullptr", "src/titleui/widgets/optionsform.cpp:1695", "addOptionInternal", "Invalid pointer": Exit Function
        OptionsForm_handleEndedRowEditing strongOptionsForm, CStr(f.cap(1))

    Case FN_OptionsDropdown_canChange           ' OptionsDropdown<T>::make: [](DropdownWidget &widget, size_t newIndex, std::shared_ptr<WIDGET> newSelectedWidget) -> bool
        Set strongSelf = args(0)
        newIndex = CLng(args(1))
        Set newSelectedWidget = args(2)
        Set d = OptionsDropdown_(strongSelf)
        If Not (newIndex < d.current.count) Then ASSERT_FAILURE "newIndex < strongSelf->current.choices.size()", "src/titleui/widgets/optionsform.h:297", "make", "Invalid index": optionsform_c_fn = False: Exit Function
        If d.current.disabled(newIndex) Then
            optionsform_c_fn = False
            Exit Function
        End If
        onSelectResult = CBool(WzFunc_call(d.onSelectFunc, d.current.value(newIndex)))
        If onSelectResult Then
            d.current.currentIdx = newIndex
            If d.needsStateUpdates Then
                widgScheduleTask WzFunc_new(FN_OptionsDropdown_updateTask, strongSelf)
            End If
        Else
            d.current.disabled(newIndex) = True
            WzOptionsChoiceWidget_setDisabled newSelectedWidget, True
        End If
        optionsform_c_fn = onSelectResult

    Case FN_OptionsDropdown_updateTask          ' [strongSelf]() { strongSelf->update(false); }
        vt_WIDGET_ovc_update f.cap(0), False
    End Select
End Function
