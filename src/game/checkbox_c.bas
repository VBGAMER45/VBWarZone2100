Attribute VB_Name = "checkbox_c"
' lib/widget/checkbox.cpp, checkbox.h: the checkbox button WzCheckboxButton
' (its fields: WzCheckboxButton.cls).
' (VB) widget class WzCheckboxButton : W_BUTTON
Option Explicit

Private Const CB_INNER_INSET As Long = 2

' WzFunc ids (fork W1: 1500-1599 here)
Public Const FN_WzCheckboxButton_onClick As Long = 1500 ' the constructor's addOnClickHandler lambda

' the lambdas of this module
Public Function checkbox_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim button As WIDGET, self As WzCheckboxButton
    Select Case f.id
    Case FN_WzCheckboxButton_onClick
        Set button = args(0)
        Set self = WzCheckboxButton_(button) ' static_cast<WzCheckboxButton&>(button)
        self.isChecked = Not self.isChecked
    End Select
End Function

' (VB) a PIELIGHT with its byte.a replaced
Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' WzCheckboxButton::WzCheckboxButton()
Public Sub WzCheckboxButton_construct(ByVal self As WIDGET)
    Dim d As WzCheckboxButton
    W_BUTTON_construct self
    self.kind = K_WzCheckboxButton ' (VB) virtual calls from here on reach WzCheckboxButton's overrides
    Set d = WzCheckboxButton_(self)
    d.isChecked = False
    d.textColor = psPalette(WZCOL_TEXT_MEDIUM)
    W_BUTTON_addOnClickHandler self, WzFunc_new(FN_WzCheckboxButton_onClick)
End Sub

' std::make_shared<WzCheckboxButton>()
Public Function WzCheckboxButton_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_WzCheckboxButton)
    WzCheckboxButton_construct w
    Set WzCheckboxButton_new = w
End Function

Public Function WzCheckboxButton_calculateDesiredDimensions(ByVal self As WIDGET) As Vector2i
    Dim d As WzCheckboxButton, cbSize As Long, checkboxPos As Vector2i, textLeftPos As Long
    Set d = WzCheckboxButton_(self)
    cbSize = WzCheckboxButton_checkboxSize(self)
    checkboxPos = Vector2i_(self.x(), self.y())
    textLeftPos = checkboxPos.x + cbSize + 7

    ' TODO: Incorporate padding?
    WzCheckboxButton_calculateDesiredDimensions = Vector2i_(textLeftPos + d.wzText.width(), MAX_(d.wzText.lineSize(), cbSize))
End Function

Public Sub WzCheckboxButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzCheckboxButton, x0 As Long, y0 As Long, down As Boolean, grey As Boolean, cbSize As Long
    Dim checkboxOffset As Vector2i, checkboxPos As Vector2i, notifyBoxAddColor As Long, checkBoxInsideColor As Long
    Dim textLeftPos As Long, fx As Long, fw As Long, fy As Long
    Set d = WzCheckboxButton_(self)
    d.wzText.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID

    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    down = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    grey = (self.getState() And WBUT_DISABLE) <> 0

    ' calculate checkbox dimensions
    cbSize = WzCheckboxButton_checkboxSize(self)
    checkboxOffset = Vector2i_(0, (self.height() - cbSize) \ 2) ' left-align, center vertically
    checkboxPos = Vector2i_(x0 + checkboxOffset.x, y0 + checkboxOffset.y)

    ' draw checkbox border
    notifyBoxAddColor = psPalette(WZCOL_NOTIFICATION_BOX)
    notifyBoxAddColor = pielightWithAlpha(notifyBoxAddColor, CLng(Fix(CSng(SHRU(notifyBoxAddColor, 24)) * 0.7!)))
    pie_UniTransBoxFill checkboxPos.x, checkboxPos.y, checkboxPos.x + cbSize, checkboxPos.y + cbSize, notifyBoxAddColor
    iV_Box2 checkboxPos.x, checkboxPos.y, checkboxPos.x + cbSize, checkboxPos.y + cbSize, psPalette(WZCOL_TEXT_MEDIUM), psPalette(WZCOL_TEXT_MEDIUM)

    If down Or d.isChecked Then
        ' draw checkbox "checked" inside
        checkBoxInsideColor = psPalette(WZCOL_TEXT_MEDIUM)
        checkBoxInsideColor = pielightWithAlpha(checkBoxInsideColor, 200)
        pie_UniTransBoxFill checkboxPos.x + CB_INNER_INSET, checkboxPos.y + CB_INNER_INSET, checkboxPos.x + cbSize - (CB_INNER_INSET), checkboxPos.y + cbSize - (CB_INNER_INSET), checkBoxInsideColor
    End If

    ' Display text to the right of the checkbox image
    textLeftPos = checkboxPos.x + cbSize + 7
    fx = textLeftPos
    fw = d.wzText.width()
    fy = yOffset + self.y() + (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()

    If (self.style And WBUT_TXTCENTRE) <> 0 Then                         'check for centering, calculate offset.
        fx = textLeftPos + ((self.width() - fw) \ 2)
    End If

    d.wzText.render fx, fy, d.textColor

    If grey Then
        ' disabled, render something over it!
        iV_TransBoxFill x0, y0, x0 + self.width(), y0 + self.height()
    End If
End Sub

Public Function WzCheckboxButton_getIsChecked(ByVal self As WIDGET) As Boolean
    WzCheckboxButton_getIsChecked = WzCheckboxButton_(self).isChecked
End Function

Public Sub WzCheckboxButton_setIsChecked(ByVal self As WIDGET, ByVal val_ As Boolean)
    WzCheckboxButton_(self).isChecked = val_
End Sub

Public Sub WzCheckboxButton_setTextColor(ByVal self As WIDGET, ByVal color As Long)
    WzCheckboxButton_(self).textColor = color
End Sub

Private Function WzCheckboxButton_checkboxSize(ByVal self As WIDGET) As Long
    Dim d As WzCheckboxButton
    Set d = WzCheckboxButton_(self)
    d.wzText.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID
    WzCheckboxButton_checkboxSize = d.wzText.lineSize() - 2
End Function
