Attribute VB_Name = "multibutform_c"
' lib/widget/multibutform.cpp, multibutform.h: the widget classes
' MultibuttonWidget and MultichoiceWidget (docs/CONVENTIONS.md, "Widgets"; the
' fields are MultibuttonWidget.cls). enum class MultibuttonWidget::ButtonAlignment
' is MultibuttonWidget_ButtonAlignment. buttons (a vector of pairs) is two
' Collections, buttons (the W_BUTTONs) and buttonValues (their ints), element n
' being item n + 1 of each. setLabel(const std::shared_ptr<WIDGET>&) is
' setLabel_widget. W_CAN_CHOOSE_FUNC (widget, newValue) returns a Boolean,
' W_ON_CHOOSE_FUNC (widget, newValue) nothing.
' Functions for the MultibuttonWidget widget
' (VB) widget class MultibuttonWidget : W_FORM
' (VB) virtual MultibuttonWidget.stateChanged
' (VB) widget class MultichoiceWidget : MultibuttonWidget
Option Explicit

' enum class ButtonAlignment
Public Enum MultibuttonWidget_ButtonAlignment
    MultibuttonWidget_ButtonAlignment_RIGHT_ALIGN
    MultibuttonWidget_ButtonAlignment_CENTER_ALIGN
End Enum

Public Const FN_MultibuttonWidget_buttonClick As Long = 2050

Public Function multibutform_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim button As WIDGET, pParent As WIDGET
    Select Case f.id
    Case FN_MultibuttonWidget_buttonClick
        ' [value](W_BUTTON& button)
        Set button = args(0)
        Set pParent = button.parent()
        ' assert(pParent != nullptr); (NDEBUG in a release build)
        MultibuttonWidget_internalHandleButtonClick pParent, f.cap(0)
    End Select
End Function

' std::make_shared<MultibuttonWidget>(value)
Public Function MultibuttonWidget_new(Optional ByVal value As Long = -1) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_MultibuttonWidget)
    MultibuttonWidget_construct w, value
    Set MultibuttonWidget_new = w
End Function

' MultibuttonWidget(int value = -1)
Public Sub MultibuttonWidget_construct(ByVal self As WIDGET, Optional ByVal value As Long = -1)
    Dim d As MultibuttonWidget
    W_FORM_construct self
    self.kind = K_MultibuttonWidget ' (VB) virtual calls from here on reach MultibuttonWidget's overrides
    Set d = MultibuttonWidget_(self)
    Set d.label = Nothing
    d.currentValue_ = value
    d.disabled = False
    d.gap_ = 3
    d.outerPaddingLeft_ = 0
    d.outerPaddingRight_ = 0
    d.lockCurrent = False
    d.butAlign = MultibuttonWidget_ButtonAlignment_RIGHT_ALIGN
End Sub

Public Sub MultibuttonWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    iV_ShadowBox xOffset + self.x(), yOffset + self.y(), xOffset + self.x() + self.width() - 1, yOffset + self.y() + self.height() - 1, 0, psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BACKGROUND)
End Sub

Public Sub MultibuttonWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As MultibuttonWidget, s As Long, width_of_all_buttons As Long, i As Long, button As WIDGET, labelX0 As Long
    Set d = MultibuttonWidget_(self)
    s = self.width() - d.gap_ - d.outerPaddingRight_
    If d.butAlign = MultibuttonWidget_ButtonAlignment_CENTER_ALIGN Then
        width_of_all_buttons = MultibuttonWidget_widthOfAllButtons(self)
        If width_of_all_buttons < self.width() Then
            s = self.width() - d.gap_ - ((self.width() - width_of_all_buttons) \ 2)
        End If
        ' fall-through
    End If
    If d.butAlign = MultibuttonWidget_ButtonAlignment_CENTER_ALIGN Or d.butAlign = MultibuttonWidget_ButtonAlignment_RIGHT_ALIGN Then
        For i = d.buttons.Count To 1 Step -1
            Set button = d.buttons(i)
            button.move s - button.width(), (self.height() - button.height()) \ 2
            s = s - (button.width() + d.gap_)
        Next
    End If
    If Not d.label Is Nothing Then
        labelX0 = d.outerPaddingLeft_ + d.gap_
        d.label.setGeometry labelX0, 0, s - labelX0, self.height()
    End If
End Sub

Private Function MultibuttonWidget_widthOfAllButtons(ByVal self As WIDGET) As Long
    Dim d As MultibuttonWidget, width_of_all_buttons As Long, i As Long, button As WIDGET
    Set d = MultibuttonWidget_(self)
    width_of_all_buttons = 0
    For i = 0 To d.buttons.Count - 1
        If i > 0 Then
            width_of_all_buttons = width_of_all_buttons + d.gap_
        End If
        Set button = d.buttons(i + 1)
        width_of_all_buttons = width_of_all_buttons + button.width()
    Next
    MultibuttonWidget_widthOfAllButtons = width_of_all_buttons
End Function

Public Function MultibuttonWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As MultibuttonWidget, result As Long
    Set d = MultibuttonWidget_(self)
    result = d.gap_ + MultibuttonWidget_widthOfAllButtons(self) + d.gap_
    If Not d.label Is Nothing Then
        result = result + d.gap_ + d.label.idealWidth()
    End If
    MultibuttonWidget_idealWidth = result
End Function

Public Function MultibuttonWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As MultibuttonWidget, result As Long, i As Long, button As WIDGET
    Set d = MultibuttonWidget_(self)
    result = 0
    For i = 0 To d.buttons.Count - 1
        Set button = d.buttons(i + 1)
        result = MAX_(result, button.idealHeight())
    Next
    If Not d.label Is Nothing Then
        result = MAX_(result, d.label.idealHeight())
    End If
    MultibuttonWidget_idealHeight = result
End Function

Public Sub MultibuttonWidget_setLabel(ByVal self As WIDGET, ByVal text As String)
    Dim d As MultibuttonWidget, newLabel As WIDGET
    Set d = MultibuttonWidget_(self)
    If Not d.label Is Nothing Then
        self.detach d.label
    End If
    Set newLabel = W_LABEL_new()
    self.attach newLabel
    newLabel.setString text
    W_LABEL_setCacheNeverExpires newLabel, True
    W_LABEL_setCanTruncate newLabel, True
    Set d.label = newLabel

    self.geometryChanged
End Sub

Public Sub MultibuttonWidget_setLabel_widget(ByVal self As WIDGET, ByVal widget As WIDGET)
    Dim d As MultibuttonWidget
    Set d = MultibuttonWidget_(self)
    If Not d.label Is Nothing Then
        self.detach d.label
    End If
    If Not widget Is Nothing Then
        self.attach widget
    End If
    Set d.label = widget

    self.geometryChanged
End Sub

' (VB) value == currentValue_ && lockCurrent ? WBUT_LOCK : disabled ? WBUT_DISABLE : 0
Private Function MultibuttonWidget_buttonState(ByVal self As WIDGET, ByVal value As Long) As Long
    Dim d As MultibuttonWidget
    Set d = MultibuttonWidget_(self)
    If value = d.currentValue_ And d.lockCurrent Then
        MultibuttonWidget_buttonState = WBUT_LOCK
    ElseIf d.disabled Then
        MultibuttonWidget_buttonState = WBUT_DISABLE
    Else
        MultibuttonWidget_buttonState = 0
    End If
End Function

Public Sub MultibuttonWidget_addButton(ByVal self As WIDGET, ByVal value As Long, ByVal button As WIDGET)
    Dim d As MultibuttonWidget
    Set d = MultibuttonWidget_(self)
    self.attach button
    button.setState MultibuttonWidget_buttonState(self, value)
    d.buttons.Add button
    d.buttonValues.Add value

    W_BUTTON_addOnClickHandler button, WzFunc_new(FN_MultibuttonWidget_buttonClick, value)

    self.geometryChanged
End Sub

Public Sub MultibuttonWidget_clearButtons(ByVal self As WIDGET)
    Dim d As MultibuttonWidget, i As Long
    Set d = MultibuttonWidget_(self)
    For i = 1 To d.buttons.Count
        self.detach d.buttons(i)
    Next
    Set d.buttons = New Collection
    Set d.buttonValues = New Collection

    self.geometryChanged
End Sub

Public Function MultibuttonWidget_numButtons(ByVal self As WIDGET) As Long
    MultibuttonWidget_numButtons = MultibuttonWidget_(self).buttons.Count
End Function

Public Function MultibuttonWidget_getButtonByValue(ByVal self As WIDGET, ByVal value As Long) As WIDGET
    Dim d As MultibuttonWidget, i As Long
    Set d = MultibuttonWidget_(self)
    For i = 1 To d.buttonValues.Count
        If d.buttonValues(i) = value Then
            Set MultibuttonWidget_getButtonByValue = d.buttons(i)
            Exit Function
        End If
    Next
End Function

Public Sub MultibuttonWidget_setButtonMinClickInterval(ByVal self As WIDGET, ByVal interval As Long)
    Dim d As MultibuttonWidget, i As Long, button As WIDGET
    Set d = MultibuttonWidget_(self)
    For i = 1 To d.buttons.Count
        Set button = d.buttons(i)
        If Not button Is Nothing Then
            W_BUTTON_(button).minClickInterval = interval
        End If
    Next
End Sub

Public Sub MultibuttonWidget_setButtonAlignment(ByVal self As WIDGET, ByVal newAlignment As MultibuttonWidget_ButtonAlignment)
    If newAlignment = MultibuttonWidget_(self).butAlign Then
        Exit Sub
    End If
    MultibuttonWidget_(self).butAlign = newAlignment
    self.geometryChanged
End Sub

Public Sub MultibuttonWidget_enable(ByVal self As WIDGET, Optional ByVal enabled As Boolean = True)
    If (Not enabled) = MultibuttonWidget_(self).disabled Then
        Exit Sub
    End If

    MultibuttonWidget_(self).disabled = Not enabled
    vt_MultibuttonWidget_stateChanged self
End Sub

Public Sub MultibuttonWidget_disable(ByVal self As WIDGET)
    MultibuttonWidget_enable self, False
End Sub

Public Function MultibuttonWidget_isDisabled(ByVal self As WIDGET) As Boolean
    MultibuttonWidget_isDisabled = MultibuttonWidget_(self).disabled
End Function

Public Sub MultibuttonWidget_setGap(ByVal self As WIDGET, ByVal gap As Long)
    If gap = MultibuttonWidget_(self).gap_ Then
        Exit Sub
    End If

    MultibuttonWidget_(self).gap_ = gap
    self.geometryChanged
End Sub

Public Function MultibuttonWidget_getGap(ByVal self As WIDGET) As Long
    MultibuttonWidget_getGap = MultibuttonWidget_(self).gap_
End Function

Public Sub MultibuttonWidget_setOuterPaddingX(ByVal self As WIDGET, ByVal left_ As Long, ByVal right_ As Long)
    Dim d As MultibuttonWidget
    Set d = MultibuttonWidget_(self)
    If left_ = d.outerPaddingLeft_ And right_ = d.outerPaddingRight_ Then
        Exit Sub
    End If

    d.outerPaddingLeft_ = left_
    d.outerPaddingRight_ = right_
    self.geometryChanged
End Sub

Public Function MultibuttonWidget_currentValue(ByVal self As WIDGET) As Long
    MultibuttonWidget_currentValue = MultibuttonWidget_(self).currentValue_
End Function

Public Sub MultibuttonWidget_setCanChooseHandler(ByVal self As WIDGET, ByVal canChooseFunc As WzFunc)
    Set MultibuttonWidget_(self).canChooseHandler = canChooseFunc
End Sub

Public Sub MultibuttonWidget_addOnChooseHandler(ByVal self As WIDGET, ByVal onChooseFunc As WzFunc)
    MultibuttonWidget_(self).onChooseHandlers.Add onChooseFunc
End Sub

Private Sub MultibuttonWidget_internalHandleButtonClick(ByVal self As WIDGET, ByVal value As Long)
    Dim d As MultibuttonWidget
    Set d = MultibuttonWidget_(self)
    If value = d.currentValue_ And d.lockCurrent Then
        Exit Sub
    End If

    If Not d.canChooseHandler Is Nothing Then
        If Not CBool(WzFunc_call(d.canChooseHandler, self, value)) Then
            ' abort change
            Exit Sub
        End If
    End If

    MultibuttonWidget_choose self, value
End Sub

Public Sub MultibuttonWidget_choose(ByVal self As WIDGET, ByVal value As Long)
    Dim d As MultibuttonWidget, i As Long, onChoose As Variant, lockedScreen As W_SCREEN
    Set d = MultibuttonWidget_(self)
    If value = d.currentValue_ And d.lockCurrent Then
        Exit Sub
    End If

    d.currentValue_ = value
    vt_MultibuttonWidget_stateChanged self

    ' Call all onChoose event handlers
    For i = 1 To d.onChooseHandlers.Count
        Set onChoose = d.onChooseHandlers(i)
        If Not onChoose Is Nothing Then
            WzFunc_call onChoose, self, d.currentValue_
        End If
    Next

    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        lockedScreen.setReturn self
    End If
End Sub

Public Sub MultibuttonWidget_stateChanged(ByVal self As WIDGET)
    Dim d As MultibuttonWidget, i As Long, button As WIDGET
    Set d = MultibuttonWidget_(self)
    For i = 1 To d.buttons.Count
        Set button = d.buttons(i)
        button.setState MultibuttonWidget_buttonState(self, d.buttonValues(i))
    Next
End Sub

'--------------------------------------------------------------------------
' MultichoiceWidget

' std::make_shared<MultichoiceWidget>(value)
Public Function MultichoiceWidget_new(Optional ByVal value As Long = -1) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_MultichoiceWidget)
    MultichoiceWidget_construct w, value
    Set MultichoiceWidget_new = w
End Function

' MultichoiceWidget(int value = -1)
Public Sub MultichoiceWidget_construct(ByVal self As WIDGET, Optional ByVal value As Long = -1)
    MultibuttonWidget_construct self, value
    self.kind = K_MultichoiceWidget ' (VB) virtual calls from here on reach MultichoiceWidget's overrides
    MultibuttonWidget_(self).lockCurrent = True
End Sub
