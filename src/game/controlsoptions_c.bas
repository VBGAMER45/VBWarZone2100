Attribute VB_Name = "controlsoptions_c"
' src/titleui/options/controlsoptions.cpp: the Controls Options form (makeControlsOptionsForm) with
' its key bindings (KeyOptionsForm, OptionsKeyBindingWidget, OptionsKeyBindingsEdit,
' WzFullscreenKeyBindingEditingOverlay, addKeyBindingsToOptionsForm), whole. The widget classes
' follow docs/CONVENTIONS.md, "Widgets" (their fields: the data classes <Class>.cls).
' (VB) The input manager is cluster G's (lib/framework/input): InputManager& inputManager is the
' global gInputManager (gInputManager_<method> procedures), a const KeyFunctionInfo& is a handle
' (KeyFunctionInfo_<field>), a KeyMapping& / optional<std::reference_wrapper<KeyMapping>> a handle
' (KeyMapping_<field>; 0 = nullopt), a KeyMappingInput its source and value (two Longs).
Option Explicit

' (VB) widget class KeyOptionsForm : OptionsForm
' (VB) widget class OptionsKeyBindingWidget : W_BUTTON
' (VB) widget class OptionsKeyBindingsEdit : WIDGET
' (VB) widget class WzFullscreenKeyBindingEditingOverlay : W_FULLSCREENOVERLAY_CLICKFORM

Public Const FN_MouseDragToRotateIsBound As Long = 4650     ' OptionInfo::AvailabilityResult MouseDragToRotateIsBound(const OptionInfo&)
Public Const FN_KBE_buttonClick As Long = 4651              ' OptionsKeyBindingsEdit::initialize, [weakSelf, slot](W_BUTTON&)
Public Const FN_KBE_overlayCalcLayout As Long = 4652        ' WzFullscreenKeyBindingEditingOverlay::make, LAMBDA_CALCLAYOUT_SIMPLE({...})
Public Const FN_KBE_onPushedKeyCombo As Long = 4653         ' OptionsKeyBindingsEdit::triggerEditModeForSlot, [weakSelf, slot](const KeyMappingInput input, optional<KEY_CODE> metaKey)
Public Const FN_KBE_onCancelPressed As Long = 4654          ' OptionsKeyBindingsEdit::triggerEditModeForSlot, [weakSelf]()
Public Const FN_controls_rightClickOrders_populate As Long = 4655 ' makeControlsOptionsForm, "controls.mouse.rightClickOrders"
Public Const FN_controls_rightClickOrders_select As Long = 4656
Public Const FN_controls_trapCursor_populate As Long = 4657 ' "controls.mouse.trapCursor"
Public Const FN_controls_trapCursor_select As Long = 4658
Public Const FN_controls_coloredCursors_populate As Long = 4659 ' "controls.mouse.coloredCursors"
Public Const FN_controls_coloredCursors_select As Long = 4660
Public Const FN_controls_cursorSize_populate As Long = 4661 ' "controls.mouse.cursorSize"
Public Const FN_controls_cursorSize_select As Long = 4662
Public Const FN_controls_camMoveSpeed_get As Long = 4663    ' "controls.camMoveSpeed"
Public Const FN_controls_camMoveSpeed_set As Long = 4664
Public Const FN_controls_camStartZoom_get As Long = 4665    ' "controls.camStartZoom"
Public Const FN_controls_camStartZoom_set As Long = 4666
Public Const FN_controls_camZomRate_get As Long = 4667      ' "controls.camZomRate"
Public Const FN_controls_camZomRate_set As Long = 4668
Public Const FN_controls_rotateCamera_populate As Long = 4669 ' "controls.mouse.rotateCamera"
Public Const FN_controls_rotateCamera_select As Long = 4670
Public Const FN_controls_reverseRotation_populate As Long = 4671 ' "controls.mouse.reverseRotation"
Public Const FN_controls_reverseRotation_select As Long = 4672
Public Const FN_controls_panCamera_populate As Long = 4673  ' "controls.mouse.panCamera"
Public Const FN_controls_panCamera_select As Long = 4674
Public Const FN_controls_resetKeymappings_populate As Long = 4675 ' "controls.keys.resetKeymappings"
Public Const FN_controls_resetKeymappings_select As Long = 4676

' OptionsKeyBindingWidget's const members
Private Const OptionsKeyBindingWidget_horizontalPadding As Long = 5
Private Const OptionsKeyBindingWidget_verticalPadding As Long = 8

' OptionsKeyBindingsEdit's const members
Private Const OptionsKeyBindingsEdit_minimumWidth As Long = 10
Private Const OptionsKeyBindingsEdit_buttonWidth As Long = 110
Private Const OptionsKeyBindingsEdit_buttonSpacing As Long = 5

' (VB) PIELIGHT's byte.a, read and set
Private Function pielightA(ByVal c As Long) As Long
    pielightA = SHRU(c, 24) And &HFF&
End Function

Private Function pielightWithA(ByVal c As Long, ByVal a As Long) As Long
    pielightWithA = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' static inline WzString keyFunctionInfoToOptionId(const KeyFunctionInfo& info)
Private Function keyFunctionInfoToOptionId(ByVal info As Long) As String
    keyFunctionInfoToOptionId = "keybinding." & KeyFunctionInfo_name(info)
End Function

' MARK: - KeyOptionsForm

Public Sub KeyOptionsForm_construct(ByVal self As WIDGET)
    OptionsForm_construct self
    self.kind = K_KeyOptionsForm
    KeyOptionsForm_(self).didEditKeybinding = False
End Sub

Public Function KeyOptionsForm_make() As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_KeyOptionsForm)
    KeyOptionsForm_construct result
    OptionsForm_initialize result
    Set KeyOptionsForm_make = result
End Function

Public Sub KeyOptionsForm_destruct(ByVal self As WIDGET)
    If KeyOptionsForm_(self).didEditKeybinding Then
        gInputManager_saveMappings
    End If
End Sub

Public Sub KeyOptionsForm_runRecursive(ByVal self As WIDGET, psContext As W_CONTEXT)
    inputRestoreMetaKeyState    ' HACK: to ensure meta keys are set to down if physically down (regardless of prior calls to clear logical input state)
    WIDGET_runRecursive self, psContext     ' OptionsForm::runRecursive (WIDGET's)
End Sub

Public Sub KeyOptionsForm_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim currMapping As Long, info As Long, keyOptionId As String
    currMapping = gInputManager_findCurrentMapping(False, True)
    If currMapping <> 0 Then
        info = KeyMapping_info(currMapping)
        keyOptionId = keyFunctionInfoToOptionId(info)
        OptionsForm_jumpToOptionId self, keyOptionId
    End If

    OptionsForm_run self, psContext
End Sub

Public Sub KeyOptionsForm_informDidEditKeybinding(ByVal self As WIDGET)
    KeyOptionsForm_(self).didEditKeybinding = True
End Sub

' MARK: - OptionsKeyBindingWidget

' OptionsKeyBindingWidget() (with the member initialisers)
Public Sub OptionsKeyBindingWidget_construct(ByVal self As WIDGET)
    Dim d As OptionsKeyBindingWidget
    W_BUTTON_construct self
    self.kind = K_OptionsKeyBindingWidget
    Set d = OptionsKeyBindingWidget_(self)
    Set d.wzText = WzCachedText_new()
    d.FontID = font_regular
    d.bindingColor = psPalette(WZCOL_FORM_TEXT)
    d.cachedIdealTextWidth = 0
    d.bEditable = False
    d.isSelected = False
    d.isEditingMode = False
    d.isTruncated = False
    d.lastWidgetWidth = 0
    d.highlightBackgroundColor = psPalette(WZCOL_MENU_SCORE_BUILT)
End Sub

' static std::shared_ptr<OptionsKeyBindingWidget> make(const std::shared_ptr<OptionsKeyBindingsEdit>& parent, KeyMappingSlot slot)
Public Function OptionsKeyBindingWidget_make(ByVal parent As WIDGET, ByVal slot As Long) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_OptionsKeyBindingWidget)
    OptionsKeyBindingWidget_construct result
    OptionsKeyBindingWidget_update result, parent, slot
    Set OptionsKeyBindingWidget_make = result
End Function

Public Sub OptionsKeyBindingWidget_setIsEditingMode(ByVal self As WIDGET, ByVal val_ As Boolean)
    Dim d As OptionsKeyBindingWidget, newProgressBorder As W_BUTTON_ProgressBorder, has_newProgressBorder As Boolean
    Dim inset As W_BUTTON_ProgressBorder_BorderInset
    Set d = OptionsKeyBindingWidget_(self)
    d.isEditingMode = val_
    has_newProgressBorder = False       ' optional<ProgressBorder> newProgressBorder;
    If d.isEditingMode Then
        inset = W_BUTTON_ProgressBorder_BorderInset_(1, 1, 1, 1)
        newProgressBorder = W_BUTTON_ProgressBorder_indeterminate(inset)
        has_newProgressBorder = True
    End If
    If has_newProgressBorder Then
        W_BUTTON_setProgressBorder self, newProgressBorder, pal_RGBA(255, 255, 255, 200)
    Else
        W_BUTTON_setProgressBorder_nullopt self, pal_RGBA(255, 255, 255, 200)
    End If
End Sub

Public Sub OptionsKeyBindingWidget_update(ByVal self As WIDGET, ByVal parent As WIDGET, ByVal slot As Long)
    Dim d As OptionsKeyBindingWidget, p As OptionsKeyBindingsEdit, newStringValue As String, mapping As Long
    Dim isBoundToNumpad As Boolean
    Set d = OptionsKeyBindingWidget_(self)
    Set p = OptionsKeyBindingsEdit_(parent)
    d.bEditable = (KeyFunctionInfo_type(p.info) = KeyMappingType_ASSIGNABLE)
    If Not d.bEditable Then
        self.setState WBUT_DISABLE
    ElseIf d.bEditable And (self.getState() And WBUT_DISABLE) <> 0 Then
        self.setState 0
    End If
    d.bindingColor = psPalette(WZCOL_TEXT_BRIGHT)

    newStringValue = ""
    mapping = p.mappings(slot)
    If mapping <> 0 Then
        If Not KeyMapping_inputIsCleared(mapping) Then
            ' Check to see if key is on the numpad, if so tell user and change color
            isBoundToNumpad = KeyMapping_inputSource(mapping) = KeyMappingInputSource_KEY_CODE And KeyMapping_inputValue(mapping) >= KEY_KP_0 And KeyMapping_inputValue(mapping) <= KEY_KPENTER
            If isBoundToNumpad Then
                d.bindingColor = psPalette(WZCOL_YELLOW)
            End If
            newStringValue = KeyMapping_toString(mapping)
        End If
    End If
    d.text = newStringValue
    d.cachedIdealTextWidth = iV_GetTextWidth(d.text, d.FontID)
End Sub

Public Sub OptionsKeyBindingWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As OptionsKeyBindingWidget, x0 As Long, y0 As Long, w As Long, h As Long, isHighlight As Boolean
    Dim fontID As iV_fonts, shrunk As iV_fonts, availableButtonTextWidth As Long, backgroundColor As Long, divisor As Long
    Dim backX0 As Long, backY0 As Long, backX1 As Long, backY1 As Long, drawHighlightBorder As Boolean, textColor As Long
    Dim maxDisplayableMainTextWidth As Long, textX0 As Long, textY0 As Long
    Set d = OptionsKeyBindingWidget_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()

    isHighlight = d.bEditable And d.isSelected

    fontID = d.wzText.getFontID()
    If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzText.getText() <> d.text Then
        fontID = d.FontID
        d.wzText.setText d.text, fontID
    End If

    availableButtonTextWidth = w - (OptionsKeyBindingWidget_horizontalPadding * 2)
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

    ' Draw background
    backgroundColor = psPalette(WZCOL_FORM_DARK)
    If d.bEditable Then divisor = 2 Else divisor = 3
    backgroundColor = pielightWithA(backgroundColor, pielightA(backgroundColor) \ divisor)
    backX0 = x0
    backY0 = y0
    backX1 = x0 + w - 1
    backY1 = y0 + h - 1
    drawHighlightBorder = (isHighlight And Not d.isEditingMode)
    If drawHighlightBorder Then
        pie_UniTransBoxFill backX0, backY0, backX1, backY1, pal_RGBA(255, 255, 255, 100)
        backX0 = backX0 + 1
        backY0 = backY0 + 1
        backX1 = backX1 - 1
        backY1 = backY1 - 1
    End If
    If (isHighlight Or d.isEditingMode) And pielightA(d.highlightBackgroundColor) <> 0 Then     ' !isTransparent()
        backgroundColor = d.highlightBackgroundColor
    End If
    pie_UniTransBoxFill backX0, backY0, backX1, backY1, backgroundColor

    ' Draw the main text
    textColor = d.bindingColor
    If Not d.bEditable Then
        textColor = psPalette(WZCOL_TEXT_MEDIUM)
    End If

    maxDisplayableMainTextWidth = availableButtonTextWidth

    textX0 = x0 + OptionsKeyBindingWidget_horizontalPadding + (availableButtonTextWidth - MIN_(maxDisplayableMainTextWidth, d.wzText.deref().width())) \ 2
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

Public Sub OptionsKeyBindingWidget_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    OptionsKeyBindingWidget_(self).wzText.tick
End Sub

Public Sub OptionsKeyBindingWidget_geometryChanged(ByVal self As WIDGET)
    ' currently no-op
End Sub

Public Function OptionsKeyBindingWidget_idealWidth(ByVal self As WIDGET) As Long
    OptionsKeyBindingWidget_idealWidth = (OptionsKeyBindingWidget_horizontalPadding * 2) + OptionsKeyBindingWidget_(self).cachedIdealTextWidth
End Function

Public Function OptionsKeyBindingWidget_idealHeight(ByVal self As WIDGET) As Long
    OptionsKeyBindingWidget_idealHeight = (OptionsKeyBindingWidget_verticalPadding * 2) + iV_GetTextLineSize(OptionsKeyBindingWidget_(self).FontID)
End Function

Public Function OptionsKeyBindingWidget_getTip(ByVal self As WIDGET) As String
    Dim d As OptionsKeyBindingWidget
    Set d = OptionsKeyBindingWidget_(self)
    If Not d.isTruncated Then
        OptionsKeyBindingWidget_getTip = ""
        Exit Function
    End If
    OptionsKeyBindingWidget_getTip = d.text
End Function

Public Sub OptionsKeyBindingWidget_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    OptionsKeyBindingWidget_(self).isSelected = True
End Sub

Public Sub OptionsKeyBindingWidget_highlightLost(ByVal self As WIDGET)
    OptionsKeyBindingWidget_(self).isSelected = False
End Sub

' MARK: - OptionsKeyBindingsEdit

' static std::shared_ptr<OptionsKeyBindingsEdit> make(const std::shared_ptr<KeyOptionsForm>& parentForm, InputManager& inputManager, const KeyFunctionInfo& info)
' (VB: inputManager is gInputManager)
Public Function OptionsKeyBindingsEdit_make(ByVal parentForm As WIDGET, ByVal info As Long) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_OptionsKeyBindingsEdit)
    OptionsKeyBindingsEdit_construct result, parentForm, info
    OptionsKeyBindingsEdit_initialize result
    Set OptionsKeyBindingsEdit_make = result
End Function

' OptionsKeyBindingsEdit(const std::shared_ptr<KeyOptionsForm>& parentForm, InputManager& inputManager, const KeyFunctionInfo& info)
Public Sub OptionsKeyBindingsEdit_construct(ByVal self As WIDGET, ByVal parentForm As WIDGET, ByVal info As Long)
    Dim d As OptionsKeyBindingsEdit
    WIDGET_construct self
    self.kind = K_OptionsKeyBindingsEdit
    Set d = OptionsKeyBindingsEdit_(self)
    d.weakSerial = WeakWidget_register(self)
    d.parentOptionsForm = OptionsForm_(parentForm).weakSerial
    d.info = info
    Set d.onChangeHandlers = New Collection
    d.cachedIdealWidth = 0
    d.cachedIdealHeight = 0
End Sub

Public Sub OptionsKeyBindingsEdit_destruct(ByVal self As WIDGET)
    WeakWidget_unregister OptionsKeyBindingsEdit_(self).weakSerial
End Sub

Private Sub OptionsKeyBindingsEdit_initialize(ByVal self As WIDGET)
    Dim d As OptionsKeyBindingsEdit, weakSelf As Long, numSlots As Long, slotIndex As Long, slot As Long, button As WIDGET
    Set d = OptionsKeyBindingsEdit_(self)
    OptionsKeyBindingsEdit_updateData self

    weakSelf = d.weakSerial

    numSlots = KeyMappingSlot_LAST
    For slotIndex = 0 To numSlots - 1
        slot = slotIndex
        Set button = OptionsKeyBindingWidget_make(self, slot)
        self.attach button
        Set d.keyMappingSlotButtons(slotIndex) = button

        W_BUTTON_addOnClickHandler button, WzFunc_new(FN_KBE_buttonClick, weakSelf, slot)
    Next

    OptionsKeyBindingsEdit_updateLayout self
End Sub

Private Sub OptionsKeyBindingsEdit_updateData(ByVal self As WIDGET)
    Dim d As OptionsKeyBindingsEdit, numSlots As Long, slotIndex As Long
    Set d = OptionsKeyBindingsEdit_(self)
    numSlots = KeyMappingSlot_LAST
    For slotIndex = 0 To numSlots - 1
        d.mappings(slotIndex) = gInputManager_mappings_get(d.info, slotIndex)
    Next
End Sub

Private Sub OptionsKeyBindingsEdit_updateButton(ByVal self As WIDGET, ByVal slot As Long)
    OptionsKeyBindingWidget_update OptionsKeyBindingsEdit_(self).keyMappingSlotButtons(slot), self, slot
End Sub

Private Sub OptionsKeyBindingsEdit_updateButtons(ByVal self As WIDGET)
    Dim numSlots As Long, slotIndex As Long
    numSlots = KeyMappingSlot_LAST
    For slotIndex = 0 To numSlots - 1
        OptionsKeyBindingsEdit_updateButton self, slotIndex
    Next
End Sub

Public Sub OptionsKeyBindingsEdit_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' no-op
End Sub

Public Sub OptionsKeyBindingsEdit_geometryChanged(ByVal self As WIDGET)
    ' reposition the mapping slot buttons
    If self.width() = 0 Or self.height() = 0 Then
        Exit Sub
    End If
    OptionsKeyBindingsEdit_updateLayout self
End Sub

' updateLayout's [&](size_t slot) -> bool hasDefaultMapping
Private Function hasDefaultMapping(ByVal info As Long, ByVal slot As Long) As Boolean
    If slot < KeyFunctionInfo_defaultMappingsCount(info) Then
        hasDefaultMapping = Not KeyFunctionInfo_defaultMappingInputIsCleared(info, slot)
    End If
End Function

Private Sub OptionsKeyBindingsEdit_updateLayout(ByVal self As WIDGET)
    Dim d As OptionsKeyBindingsEdit, w As Long, h As Long, availableWidth As Long, nextPosX0 As Long, numButtonsShown As Long
    Dim i As Long, slotHasMapping As Boolean, slotHasDefaultMapping As Boolean, actualButtonWidth As Long, a As Long
    Dim buttonIdealHeight As Long, buttonY0 As Long, button As WIDGET
    Set d = OptionsKeyBindingsEdit_(self)
    w = self.width()
    h = self.height()

    d.cachedIdealWidth = 0
    d.cachedIdealHeight = 0

    availableWidth = w
    nextPosX0 = availableWidth - OptionsKeyBindingsEdit_buttonWidth

    numButtonsShown = 0
    For i = 0 To KeyMappingSlot_LAST - 1
        slotHasMapping = False
        If d.mappings(i) <> 0 Then slotHasMapping = Not KeyMapping_inputIsCleared(d.mappings(i))
        slotHasDefaultMapping = hasDefaultMapping(d.info, i)
        If slotHasMapping Or slotHasDefaultMapping Or i = 0 Then
            numButtonsShown = numButtonsShown + 1
        End If
    Next

    If numButtonsShown > 0 Then a = availableWidth \ numButtonsShown Else a = 0
    actualButtonWidth = MIN_(a, OptionsKeyBindingsEdit_buttonWidth)
    For i = 0 To KeyMappingSlot_LAST - 1
        slotHasMapping = False
        If d.mappings(i) <> 0 Then slotHasMapping = Not KeyMapping_inputIsCleared(d.mappings(i))
        slotHasDefaultMapping = hasDefaultMapping(d.info, i)
        Set button = d.keyMappingSlotButtons(i)
        buttonIdealHeight = button.idealHeight()
        If slotHasMapping Or slotHasDefaultMapping Or i = 0 Then
            buttonY0 = (h - buttonIdealHeight) \ 2
            button.setGeometry nextPosX0, buttonY0, actualButtonWidth, buttonIdealHeight
            button.show
            d.cachedIdealWidth = d.cachedIdealWidth + OptionsKeyBindingsEdit_buttonWidth + OptionsKeyBindingsEdit_buttonSpacing
            nextPosX0 = nextPosX0 - (actualButtonWidth + OptionsKeyBindingsEdit_buttonSpacing)
        Else
            button.hide
        End If
        d.cachedIdealHeight = MAX_(d.cachedIdealHeight, buttonIdealHeight)
    Next

    d.cachedIdealWidth = MAX_(d.cachedIdealWidth, OptionsKeyBindingsEdit_minimumWidth)
End Sub

Public Function OptionsKeyBindingsEdit_idealWidth(ByVal self As WIDGET) As Long
    OptionsKeyBindingsEdit_idealWidth = OptionsKeyBindingsEdit_(self).cachedIdealWidth
End Function

Public Function OptionsKeyBindingsEdit_idealHeight(ByVal self As WIDGET) As Long
    OptionsKeyBindingsEdit_idealHeight = OptionsKeyBindingsEdit_(self).cachedIdealHeight
End Function

Public Function OptionsKeyBindingsEdit_ovc_is(ByVal self As WIDGET) As Boolean
    OptionsKeyBindingsEdit_ovc_is = True
End Function

Public Sub OptionsKeyBindingsEdit_ovc_update(ByVal self As WIDGET, Optional ByVal force As Boolean = False)
    OptionsKeyBindingsEdit_updateData self
    OptionsKeyBindingsEdit_updateButtons self
    OptionsKeyBindingsEdit_updateLayout self
End Sub

Public Sub OptionsKeyBindingsEdit_ovc_informAvailable(ByVal self As WIDGET, ByVal isAvailable As Boolean)
    ' TODO: Handle disabled?
End Sub

Public Sub OptionsKeyBindingsEdit_ovc_addOnChangeHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    OptionsKeyBindingsEdit_(self).onChangeHandlers.Add handler
End Sub

Public Sub OptionsKeyBindingsEdit_ovc_addStartingCapturedEditingHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    Set OptionsKeyBindingsEdit_(self).onStartEditingHandler = handler
End Sub

Public Sub OptionsKeyBindingsEdit_ovc_addEndingCapturedEditingHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    Set OptionsKeyBindingsEdit_(self).onEndEditingHandler = handler
End Sub

' static nonstd::optional<KEY_CODE> scanKeyBoardForPressedBindableKey(): has_value, the key in result
Private Function scanKeyBoardForPressedBindableKey(result As Long) As Boolean
    Dim i As Long
    For i = 0 To KEY_MAXSCAN - 1
        If keyPressed(i) Then
            If i <> KEY_RALT And i <> KEY_LALT And i <> KEY_RCTRL And i <> KEY_LCTRL And i <> KEY_RSHIFT And i <> KEY_LSHIFT And i <> KEY_LMETA And i <> KEY_RMETA Then   ' exceptions
                result = i              ' top row key pressed
                scanKeyBoardForPressedBindableKey = True
                Exit Function
            End If
        End If
    Next
    scanKeyBoardForPressedBindableKey = False
End Function

' static nonstd::optional<MOUSE_KEY_CODE> scanMouseForPressedBindableKey(): has_value, the key in result
Private Function scanMouseForPressedBindableKey(result As Long) As Boolean
    Dim i As Long, mouseKeyCode As Long
    For i = 0 To MOUSE_END - 1
        mouseKeyCode = i
        If mousePressed(mouseKeyCode) Then
            If mouseKeyCode <> MOUSE_LMB And mouseKeyCode <> MOUSE_RMB And mouseKeyCode <> MOUSE_MMB Then  ' exceptions
                result = mouseKeyCode   ' Bindable mouse key pressed
                scanMouseForPressedBindableKey = True
                Exit Function
            End If
        End If
    Next
    scanMouseForPressedBindableKey = False
End Function

' static optional<KEY_CODE> scanForPressedMetaKey(): has_value, the key in metakey
Private Function scanForPressedMetaKey(metakey As Long) As Boolean
    If keyDown(KEY_RALT) Or keyDown(KEY_LALT) Then
        metakey = KEY_LALT
    ElseIf keyDown(KEY_RCTRL) Or keyDown(KEY_LCTRL) Then
        metakey = KEY_LCTRL
    ElseIf keyDown(KEY_RSHIFT) Or keyDown(KEY_LSHIFT) Then
        metakey = KEY_LSHIFT
    ElseIf keyDown(KEY_RMETA) Or keyDown(KEY_LMETA) Then
        metakey = KEY_LMETA
    Else
        scanForPressedMetaKey = False
        Exit Function
    End If
    scanForPressedMetaKey = True
End Function

' MARK: - WzFullscreenKeyBindingEditingOverlay

' WzFullscreenKeyBindingEditingOverlay(W_FORMINIT const *init) (with the member initialisers)
Public Sub WzFullscreenKeyBindingEditingOverlay_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    Dim d As WzKeyBindingEditOverlayData
    W_FULLSCREENOVERLAY_CLICKFORM_construct_init self, init
    self.kind = K_WzFullscreenKeyBindingEditingOverlay
    Set d = WzFullscreenKeyBindingEditingOverlay_(self)
    d.metakeyDown_RALT = False
    d.metakeyDown_LALT = False
    d.metakeyDown_RCTRL = False
    d.metakeyDown_LCTRL = False
    d.metakeyDown_RSHIFT = False
    d.metakeyDown_LSHIFT = False
    d.metakeyDown_RMETA = False
    d.metakeyDown_LMETA = False
    d.mouseDownOutsideTargetWidget = False
End Sub

' OnPushedKeyComboFunc: void (const KeyMappingInput input, optional<KEY_CODE> metaKey): (VB) called with
' the input's source and value, has_value and the meta key
Public Function WzFullscreenKeyBindingEditingOverlay_make() As WIDGET
    Dim sInit As W_FORMINIT, result As WIDGET
    W_FORMINIT_construct sInit
    sInit.id = 0
    sInit.style = WFORM_PLAIN Or WFORM_CLICKABLE
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1
    Set sInit.calcLayout = WzFunc_new(FN_KBE_overlayCalcLayout)

    Set result = WIDGET_alloc(K_WzFullscreenKeyBindingEditingOverlay)
    WzFullscreenKeyBindingEditingOverlay_construct_init result, sInit
    Set WzFullscreenKeyBindingEditingOverlay_make = result
End Function

Public Sub WzFullscreenKeyBindingEditingOverlay_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim strongCutout As WIDGET, screenContext As W_CONTEXT, g As WzRect
    If key <> WKEY_PRIMARY Then
        Exit Sub
    End If

    Set strongCutout = W_FULLSCREENOVERLAY_CLICKFORM_(self).cutoutWidget
    If Not strongCutout Is Nothing Then
        screenContext = W_CONTEXT_convertToScreenContext(psContext)
        g = strongCutout.screenGeometry()
        If WzRect_contains(g, screenContext.mx, screenContext.my) Then
            ' ignore click inside the cutout widget
            Exit Sub
        End If
    End If

    WzFullscreenKeyBindingEditingOverlay_(self).mouseDownOutsideTargetWidget = True
End Sub

Public Sub WzFullscreenKeyBindingEditingOverlay_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As WzKeyBindingEditOverlayData, o As W_FULLSCREENOVERLAY_CLICKFORM
    Set d = WzFullscreenKeyBindingEditingOverlay_(self)
    If Not d.mouseDownOutsideTargetWidget Then
        Exit Sub
    End If

    d.mouseDownOutsideTargetWidget = False

    Set o = W_FULLSCREENOVERLAY_CLICKFORM_(self)
    If Not o.onClickedFunc Is Nothing Then
        WzFunc_call o.onClickedFunc
    End If
End Sub

Public Sub WzFullscreenKeyBindingEditingOverlay_highlightLost(ByVal self As WIDGET)
    WzFullscreenKeyBindingEditingOverlay_(self).mouseDownOutsideTargetWidget = False
End Sub

Public Sub WzFullscreenKeyBindingEditingOverlay_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzKeyBindingEditOverlayData, has_metaKey As Boolean, metaKey As Long, kc As Long, mkc As Long
    Dim o As W_FULLSCREENOVERLAY_CLICKFORM
    Set d = WzFullscreenKeyBindingEditingOverlay_(self)
    inputRestoreMetaKeyState    ' HACK: to ensure meta keys are set to down if physically down (regardless of prior calls to clear logical input state)

    has_metaKey = scanForPressedMetaKey(metaKey)

    If scanKeyBoardForPressedBindableKey(kc) Then
        If kc = KEY_ESC Then
            ' Special-handling for ESC - unhighlight the selection
            inputLoseFocus  ' clear the input buffer.
            Set o = W_FULLSCREENOVERLAY_CLICKFORM_(self)
            If Not o.onCancelPressed Is Nothing Then WzFunc_call o.onCancelPressed
            Exit Sub
        End If
        WzFunc_call d.onPushedKeyCombo, KeyMappingInputSource_KEY_CODE, kc, has_metaKey, metaKey
    End If

    If scanMouseForPressedBindableKey(mkc) Then
        WzFunc_call d.onPushedKeyCombo, KeyMappingInputSource_MOUSE_KEY_CODE, mkc, has_metaKey, metaKey
    End If

    inputLoseFocus  ' clear the input buffer.
End Sub

Private Sub WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyState(ByVal key As Long, s As Boolean)
    If keyDown(key) Then
        s = True
    ElseIf keyReleased(key) Then
        s = False
    End If
End Sub

Private Sub WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyStates(ByVal self As WIDGET)
    Dim d As WzKeyBindingEditOverlayData, s As Boolean
    Set d = WzFullscreenKeyBindingEditingOverlay_(self)
    s = d.metakeyDown_RALT: WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyState KEY_RALT, s: d.metakeyDown_RALT = s
    s = d.metakeyDown_LALT: WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyState KEY_LALT, s: d.metakeyDown_LALT = s

    s = d.metakeyDown_RCTRL: WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyState KEY_RCTRL, s: d.metakeyDown_RCTRL = s
    s = d.metakeyDown_LCTRL: WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyState KEY_LCTRL, s: d.metakeyDown_LCTRL = s

    s = d.metakeyDown_RSHIFT: WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyState KEY_RSHIFT, s: d.metakeyDown_RSHIFT = s
    s = d.metakeyDown_LSHIFT: WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyState KEY_LSHIFT, s: d.metakeyDown_LSHIFT = s

    s = d.metakeyDown_RMETA: WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyState KEY_RMETA, s: d.metakeyDown_RMETA = s
    s = d.metakeyDown_LMETA: WzFullscreenKeyBindingEditingOverlay_updateSavedMetaKeyState KEY_LMETA, s: d.metakeyDown_LMETA = s
End Sub

' MARK: - OptionsKeyBindingsEdit (the editing overlay)

Private Function OptionsKeyBindingsEdit_triggerEditModeForSlot(ByVal self As WIDGET, ByVal slot As Long) As Boolean
    Dim d As OptionsKeyBindingsEdit, weakSelf As Long, targetSlotButton As WIDGET, lockedScreen As W_SCREEN
    Dim newRootFrm As WIDGET, o As W_FULLSCREENOVERLAY_CLICKFORM
    Set d = OptionsKeyBindingsEdit_(self)
    If KeyFunctionInfo_type(d.info) <> KeyMappingType_ASSIGNABLE Then
        OptionsKeyBindingsEdit_triggerEditModeForSlot = False
        Exit Function
    End If

    weakSelf = d.weakSerial
    Set targetSlotButton = d.keyMappingSlotButtons(slot)

    Set lockedScreen = self.screenPointer
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "src/titleui/options/controlsoptions.cpp:757", "triggerEditModeForSlot", "The CampaignStartOptionsForm does not have an associated screen pointer?": OptionsKeyBindingsEdit_triggerEditModeForSlot = False: Exit Function

    If Not d.onStartEditingHandler Is Nothing Then
        WzFunc_call d.onStartEditingHandler, self
    End If

    ' Initialize the editing overlay screen
    Set d.overlayScreen = W_SCREEN_make()
    Set newRootFrm = WzFullscreenKeyBindingEditingOverlay_make()
    Set WzFullscreenKeyBindingEditingOverlay_(newRootFrm).onPushedKeyCombo = WzFunc_new(FN_KBE_onPushedKeyCombo, weakSelf, slot)
    Set o = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm)
    Set o.onCancelPressed = WzFunc_new(FN_KBE_onCancelPressed, weakSelf)
    Set o.onClickedFunc = o.onCancelPressed
    d.overlayScreen.psForm.attach newRootFrm

    OptionsKeyBindingWidget_setIsEditingMode targetSlotButton, True
    W_FULLSCREENOVERLAY_CLICKFORM_setCutoutWidget newRootFrm, targetSlotButton

    widgRegisterOverlayScreenOnTopOfScreen d.overlayScreen, lockedScreen

    OptionsKeyBindingsEdit_triggerEditModeForSlot = True
End Function

Private Sub OptionsKeyBindingsEdit_closeEditModeOverlay(ByVal self As WIDGET)
    Dim d As OptionsKeyBindingsEdit, i As Long
    Set d = OptionsKeyBindingsEdit_(self)
    If Not d.overlayScreen Is Nothing Then
        If Not d.onEndEditingHandler Is Nothing Then
            WzFunc_call d.onEndEditingHandler, self
        End If

        widgRemoveOverlayScreen d.overlayScreen
        Set d.overlayScreen = Nothing
    End If

    For i = 0 To KeyMappingSlot_LAST - 1
        OptionsKeyBindingWidget_setIsEditingMode d.keyMappingSlotButtons(i), False
    Next
End Sub

' bool onPushedKeyCombo(KeyMappingSlot slot, const KeyMappingInput input, optional<KEY_CODE> metaKeyOpt)
Private Function OptionsKeyBindingsEdit_onPushedKeyCombo(ByVal self As WIDGET, ByVal slot As Long, ByVal inputSource As Long, ByVal inputValue As Long, ByVal has_metaKeyOpt As Boolean, ByVal metaKeyOpt As Long) As Boolean
    Dim d As OptionsKeyBindingsEdit, metakey As Long, selectedInfo As Long, conflicting() As Long, conflictingCount As Long
    Dim i As Long, maybeOld As Long, foundPrimary As Long, foundMatch As Long, action As Long, newMapping As Long
    Dim onChangeFunc As WzFunc, n As Long
    Set d = OptionsKeyBindingsEdit_(self)
    If has_metaKeyOpt Then metakey = metaKeyOpt Else metakey = KEY_IGNORE

    selectedInfo = d.info
    ' Disallow modifying non-assignable mappings. (Null-check the `info` in case assertions are disabled)
    If selectedInfo = 0 Then
        OptionsKeyBindingsEdit_onPushedKeyCombo = False
        Exit Function
    End If
    If KeyFunctionInfo_type(selectedInfo) <> KeyMappingType_ASSIGNABLE Then
        OptionsKeyBindingsEdit_onPushedKeyCombo = False
        Exit Function
    End If

    ' Disallow conflicts with ALWAYS_ACTIVE keybinds, as that context always has max priority, preventing
    '   any conflicting keys from triggering.
    conflictingCount = gInputManager_mappings_findConflicting(metakey, inputSource, inputValue, KeyFunctionInfo_context(selectedInfo), conflicting)
    For i = 0 To conflictingCount - 1
        If gInputManager_contexts_isAlwaysActive(KeyFunctionInfo_context(KeyMapping_info(conflicting(i)))) Then
            audio_PlayBuildFailedOnce
            OptionsKeyBindingsEdit_onPushedKeyCombo = False
            Exit Function
        End If
    Next

    ' Clear conflicting mappings using these keys
    gInputManager_mappings_removeConflicting metakey, inputSource, inputValue, KeyFunctionInfo_context(selectedInfo)

    ' Try and see if the mapping already exists. Remove the old mapping if one does exist
    maybeOld = gInputManager_mappings_get(selectedInfo, slot)
    If maybeOld <> 0 Then
        gInputManager_mappings_remove maybeOld
    End If

    ' Figure out which `KeyAction` the mapping should use
    n = KeyFunctionInfo_defaultMappingsCount(selectedInfo)
    foundPrimary = n
    For i = 0 To n - 1
        If KeyFunctionInfo_defaultMappingSlot(selectedInfo, i) = KeyMappingSlot_PRIMARY Then foundPrimary = i: Exit For
    Next

    foundMatch = n
    For i = 0 To n - 1
        If KeyFunctionInfo_defaultMappingInputSource(selectedInfo, i) = inputSource And KeyFunctionInfo_defaultMappingInputValue(selectedInfo, i) = inputValue Then foundMatch = i: Exit For
    Next

    action = KeyAction_PRESSED
    If foundMatch <> n Then
        action = KeyFunctionInfo_defaultMappingAction(selectedInfo, foundMatch)
    ElseIf foundPrimary <> n Then
        action = KeyFunctionInfo_defaultMappingAction(selectedInfo, foundPrimary)
    End If

    ' Finally, create the new mapping
    newMapping = gInputManager_mappings_add(metakey, inputSource, inputValue, action, selectedInfo, slot)

    d.mappings(slot) = newMapping
    OptionsKeyBindingsEdit_updateButton self, slot

    For i = 1 To d.onChangeHandlers.Count
        Set onChangeFunc = d.onChangeHandlers(i)
        WzFunc_call onChangeFunc, self
    Next

    OptionsKeyBindingsEdit_onPushedKeyCombo = True
End Function

' MARK: -

' KeyFunctionEntries getVisibleKeyFunctionEntries(const KeyFunctionConfiguration& keyFuncConfig): (VB) the
' count, the KeyFunctionInfo handles in visible(0 .. count - 1); keyFuncConfig is gKeyFuncConfig
Public Function getVisibleKeyFunctionEntries(visible() As Long) As Long
    Dim allEntries() As Long, n As Long, i As Long, count As Long
    count = 0
    n = gKeyFuncConfig_allKeyFunctionEntries(allEntries)
    For i = 0 To n - 1
        If KeyFunctionInfo_type(allEntries(i)) <> KeyMappingType_HIDDEN Then
            If count = 0 Then ReDim visible(0 To n - 1)
            visible(count) = allEntries(i)
            count = count + 1
        End If
    Next
    getVisibleKeyFunctionEntries = count
End Function

' addKeyBindingsToOptionsForm's sort comparator: [&](const KeyFunctionInfo& a, const KeyFunctionInfo& b)
Private Function keyFunctionInfoLess(ByVal a As Long, ByVal b As Long) As Boolean
    Dim bContextsAreSame As Boolean
    bContextsAreSame = KeyFunctionInfo_context(a) = KeyFunctionInfo_context(b)
    If bContextsAreSame Then
        keyFunctionInfoLess = StrComp(KeyFunctionInfo_displayName(a), KeyFunctionInfo_displayName(b), vbBinaryCompare) < 0
    Else
        keyFunctionInfoLess = StrComp(gInputManager_contexts_getDisplayName(KeyFunctionInfo_context(a)), gInputManager_contexts_getDisplayName(KeyFunctionInfo_context(b)), vbBinaryCompare) < 0
    End If
End Function

' size_t addKeyBindingsToOptionsForm(const std::shared_ptr<KeyOptionsForm>& result, InputManager& inputManager, const KeyFunctionConfiguration& keyFuncConfig)
' (VB: gInputManager, gKeyFuncConfig)
Public Function addKeyBindingsToOptionsForm(ByVal result As WIDGET) As Long
    Dim infos() As Long, n As Long, i As Long, j As Long, tmp As Long, info As Long, bShouldAddSeparator As Boolean
    Dim optionInfo As OptionInfo, valueChanger As WIDGET
    n = getVisibleKeyFunctionEntries(infos)
    ' std::sort(infos.begin(), infos.end(), ...) (an insertion sort)
    For i = 1 To n - 1
        tmp = infos(i)
        j = i - 1
        Do While j >= 0
            If Not keyFunctionInfoLess(tmp, infos(j)) Then Exit Do
            infos(j + 1) = infos(j)
            j = j - 1
        Loop
        infos(j + 1) = tmp
    Next

    For i = 0 To n - 1
        info = infos(i)

        ' Add separator if changing categories
        If i = 0 Then
            bShouldAddSeparator = True
        Else
            bShouldAddSeparator = KeyFunctionInfo_context(infos(i - 1)) <> KeyFunctionInfo_context(info)
        End If
        If bShouldAddSeparator Then
            OptionsForm_addSection result, OptionsSection_(gInputManager_contexts_getDisplayName(KeyFunctionInfo_context(info)), ""), True
        End If

        Set optionInfo = OptionInfo_new(keyFunctionInfoToOptionId(info), KeyFunctionInfo_displayName(info), "")
        Set valueChanger = OptionsKeyBindingsEdit_make(result, info)
        OptionsForm_addOption result, optionInfo, valueChanger, True
    Next

    addKeyBindingsToOptionsForm = n
End Function

' static WzString mouseKeyCodeToWzString(MOUSE_KEY_CODE code)
Private Function mouseKeyCodeToWzString(ByVal code As Long) As String
    mouseKeyCodeToWzString = mouseKeyCodeToString(code)
End Function

Public Function MouseDragToRotateIsBound(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = getRotateMouseKey() <> 0     ' has_value(): 0 is nullopt
    Set MouseDragToRotateIsBound = result
End Function

' MARK: -

Public Function makeControlsOptionsForm() As WIDGET
    Dim result As WIDGET, optionInfo As OptionInfo, valueChanger As WIDGET
    Set result = KeyOptionsForm_make()

    ' Mouse:
    OptionsForm_addSection result, OptionsSection_("Mouse", ""), True
    Set optionInfo = OptionInfo_new("controls.mouse.rightClickOrders", "Right-Click Orders", "By default, Warzone 2100 uses left-click for both selection and ordering / targeting. If you prefer modern RTS mouse input, set this to On. If you are using a trackpad or touch input, you may want to set this to Off.")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_controls_rightClickOrders_populate), WzFunc_new(FN_controls_rightClickOrders_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("controls.mouse.trapCursor", "Trap Cursor", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_controls_trapCursor_populate), WzFunc_new(FN_controls_trapCursor_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("controls.mouse.coloredCursors", "Colored Cursors", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_controls_coloredCursors_populate), WzFunc_new(FN_controls_coloredCursors_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("controls.mouse.cursorSize", "Cursor Size", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_controls_cursorSize_populate), WzFunc_new(FN_controls_cursorSize_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Camera:
    OptionsForm_addSection result, OptionsSection_("Camera", ""), True
    Set optionInfo = OptionInfo_new("controls.camMoveSpeed", "Camera Speed", "")
    Set valueChanger = OptionsSlider_make(CAMERASPEED_MIN, CAMERASPEED_MAX, CAMERASPEED_STEP, WzFunc_new(FN_controls_camMoveSpeed_get), WzFunc_new(FN_controls_camMoveSpeed_set), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("controls.camStartZoom", "Camera Starting Zoom", "")
    Set valueChanger = OptionsSlider_make(MINDISTANCE_CONFIG, MAXDISTANCE, MAP_ZOOM_CONFIG_STEP, WzFunc_new(FN_controls_camStartZoom_get), WzFunc_new(FN_controls_camStartZoom_set), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("controls.camZomRate", "Camera Zoom Rate", "")
    Set valueChanger = OptionsSlider_make(MAP_ZOOM_RATE_MIN, MAP_ZOOM_RATE_MAX, MAP_ZOOM_RATE_STEP, WzFunc_new(FN_controls_camZomRate_get), WzFunc_new(FN_controls_camZomRate_set), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("controls.mouse.rotateCamera", "Mouse Rotate", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_controls_rotateCamera_populate), WzFunc_new(FN_controls_rotateCamera_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("controls.mouse.reverseRotation", "Reverse Rotation", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_MouseDragToRotateIsBound)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_controls_reverseRotation_populate), WzFunc_new(FN_controls_reverseRotation_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True, 1

    Set optionInfo = OptionInfo_new("controls.mouse.panCamera", "Mouse Pan", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_controls_panCamera_populate), WzFunc_new(FN_controls_panCamera_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    addKeyBindingsToOptionsForm result

    OptionsForm_addSection result, OptionsSection_("Reset", ""), True
    Set optionInfo = OptionInfo_new("controls.keys.resetKeymappings", "Reset Keys to:", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_controls_resetKeymappings_populate), WzFunc_new(FN_controls_resetKeymappings_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set makeControlsOptionsForm = result
End Function

' MARK: - the lambdas

Public Function controlsoptions_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim result As OptionChoices, strongBindingsEdit As WIDGET, strongSelf As WIDGET, psWidget As WIDGET, currValue As Long
    Dim strongParentOptionsForm As WIDGET
    Select Case f.id
    Case FN_MouseDragToRotateIsBound
        Set controlsoptions_c_fn = MouseDragToRotateIsBound(args(0))

    Case FN_KBE_buttonClick                     ' [weakSelf, slot](W_BUTTON&)
        Set strongBindingsEdit = WeakWidget_lock(f.cap(0))
        If Not (Not strongBindingsEdit Is Nothing) Then ASSERT_FAILURE "strongBindingsEdit != nullptr", "src/titleui/options/controlsoptions.cpp:402", "initialize", "Null parent?": Exit Function
        OptionsKeyBindingsEdit_triggerEditModeForSlot strongBindingsEdit, CLng(f.cap(1))

    Case FN_KBE_overlayCalcLayout               ' LAMBDA_CALCLAYOUT_SIMPLE: (WIDGET *psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight

    Case FN_KBE_onPushedKeyCombo                ' [weakSelf, slot](const KeyMappingInput input, optional<KEY_CODE> metaKey)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/options/controlsoptions.cpp:770", "triggerEditModeForSlot", "Widget already gone": Exit Function
        If OptionsKeyBindingsEdit_onPushedKeyCombo(strongSelf, CLng(f.cap(1)), CLng(args(0)), CLng(args(1)), CBool(args(2)), CLng(args(3))) Then
            Set strongParentOptionsForm = WeakWidget_lock(OptionsKeyBindingsEdit_(strongSelf).parentOptionsForm)
            If Not strongParentOptionsForm Is Nothing Then
                KeyOptionsForm_informDidEditKeybinding strongParentOptionsForm
            Else
                ASSERT_FAILURE "false", "src/titleui/options/controlsoptions.cpp:779", "triggerEditModeForSlot", "Unable to inform parent options form?"
            End If
        End If
        OptionsKeyBindingsEdit_closeEditModeOverlay strongSelf

    Case FN_KBE_onCancelPressed                 ' [weakSelf]()
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/options/controlsoptions.cpp:786", "triggerEditModeForSlot", "Widget already gone": Exit Function
        OptionsKeyBindingsEdit_closeEditModeOverlay strongSelf

    Case FN_controls_rightClickOrders_populate  ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "Use left-click for both selection and ordering / targeting. (The classic Warzone 2100 default.)", False
        result.addChoice "On", "Order units to move or target with right-click. Select units with left-click. (Matches other common RTS game defaults.)", True
        result.setCurrentIdxForValue getRightClickOrders()
        Set controlsoptions_c_fn = result
    Case FN_controls_rightClickOrders_select
        setRightClickOrders CBool(args(0))
        controlsoptions_c_fn = True

    Case FN_controls_trapCursor_populate        ' OptionChoices<TrapCursorMode> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", TrapCursorMode_Disabled
        result.addChoice "On", "", TrapCursorMode_Enabled
        result.addChoice "Auto", "", TrapCursorMode_Automatic
        result.setCurrentIdxForValue war_GetTrapCursor()
        Set controlsoptions_c_fn = result
    Case FN_controls_trapCursor_select
        war_SetTrapCursor CLng(args(0))
        controlsoptions_c_fn = True

    Case FN_controls_coloredCursors_populate    ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue war_GetColouredCursor()
        Set controlsoptions_c_fn = result
    Case FN_controls_coloredCursors_select
        war_SetColouredCursor CBool(args(0))
        controlsoptions_c_fn = True

    Case FN_controls_cursorSize_populate        ' OptionChoices<unsigned int> ()
        Set result = New OptionChoices
        result.addChoice "100%", "", 100
        result.addChoice "125%", "", 125
        result.addChoice "150%", "", 150
        result.addChoice "200%", "", 200
        currValue = war_getCursorScale()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add "Custom" item
            result.addChoice va("(%u%%)", currValue), "", currValue
            result.currentIdx = result.count - 1
        End If
        Set controlsoptions_c_fn = result
    Case FN_controls_cursorSize_select
        controlsoptions_c_fn = wzChangeCursorScale(CLng(args(0)))

    Case FN_controls_camMoveSpeed_get           ' []() { return war_GetCameraSpeed(); }
        controlsoptions_c_fn = war_GetCameraSpeed()
    Case FN_controls_camMoveSpeed_set           ' [](int32_t newValue)
        war_SetCameraSpeed CLng(args(0))
    Case FN_controls_camStartZoom_get
        controlsoptions_c_fn = war_GetMapZoom()
    Case FN_controls_camStartZoom_set
        war_SetMapZoom CLng(args(0))
    Case FN_controls_camZomRate_get
        controlsoptions_c_fn = war_GetMapZoomRate()
    Case FN_controls_camZomRate_set
        war_SetMapZoomRate CLng(args(0))

    Case FN_controls_rotateCamera_populate      ' OptionChoices<optional<MOUSE_KEY_CODE>> () (0 is nullopt)
        Set result = New OptionChoices
        result.addChoice "Disabled", "", 0&
        result.addChoice mouseKeyCodeToWzString(MOUSE_MMB), "", CLng(MOUSE_MMB)
        result.addChoice mouseKeyCodeToWzString(MOUSE_RMB), "", CLng(MOUSE_RMB), getRightClickOrders()
        result.setCurrentIdxForValue getRotateMouseKey()
        Set controlsoptions_c_fn = result
    Case FN_controls_rotateCamera_select
        controlsoptions_c_fn = setRotateMouseKey(CLng(args(0)))

    Case FN_controls_reverseRotation_populate   ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue getInvertMouseStatus()
        Set controlsoptions_c_fn = result
    Case FN_controls_reverseRotation_select
        setInvertMouseStatus CBool(args(0))
        controlsoptions_c_fn = True

    Case FN_controls_panCamera_populate         ' OptionChoices<optional<MOUSE_KEY_CODE>> () (0 is nullopt)
        Set result = New OptionChoices
        result.addChoice "Edge Scrolling", "", 0&
        result.addChoice mouseKeyCodeToWzString(MOUSE_MMB), "", CLng(MOUSE_MMB)
        result.addChoice mouseKeyCodeToWzString(MOUSE_RMB), "", CLng(MOUSE_RMB), getRightClickOrders()
        result.setCurrentIdxForValue getPanMouseKey()
        Set controlsoptions_c_fn = result
    Case FN_controls_panCamera_select
        controlsoptions_c_fn = setPanMouseKey(CLng(args(0)))

    Case FN_controls_resetKeymappings_populate  ' OptionChoices<WzString> ()
        Set result = New OptionChoices
        result.addChoice "(Select Profile)", "", ""
        result.addChoice "Reset Keys to: Defaults", "", "BUILTIN::DEFAULTS"
        result.currentIdx = 0
        Set controlsoptions_c_fn = result
    Case FN_controls_resetKeymappings_select    ' [/*, weakOptionForm*/](const auto& newValue) mutable -> bool
        If Len(CStr(args(0))) = 0 Then
            controlsoptions_c_fn = True
            Exit Function
        End If
        ' FUTURE TODO: Prompt before resetting?
        If CStr(args(0)) = "BUILTIN::DEFAULTS" Then
            ' reinitialise default key mappings
            gInputManager_resetMappings True
            controlsoptions_c_fn = True
            Exit Function
        End If
        controlsoptions_c_fn = False
    End Select
End Function
