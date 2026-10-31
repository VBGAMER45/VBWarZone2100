Attribute VB_Name = "paneltabbutton_c"
' lib/widget/paneltabbutton.cpp, paneltabbutton.h, whole: class WzPanelTabButton, the tab buttons of the quick chat
' form's panel switcher (and multimenu.cpp's team strategy switcher: cluster J). Its fields: WzPanelTabButton.cls.
' (VB) widget class WzPanelTabButton : W_BUTTON
Option Explicit

Private Const QCFORM_PANEL_TABS_PADDING As Long = 20
Private Const QCFORM_PANEL_TABS_HEIGHT As Long = 20
Private Const WZCOL_PANELTABS_FILL_COLOR As Long = &HDC6E0019        ' pal_RGBA(25, 0, 110, 220)
Private Const WZCOL_PANELTABS_FILL_COLOR_DARK As Long = &HFA46000A   ' pal_RGBA(10, 0, 70, 250)
Private Const WZCOL_PANELTABS_BORDER_LIGHT As Long = &H96FFFFFF      ' pal_RGBA(255, 255, 255, 150)

' (VB) a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

Public Sub WzPanelTabButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzPanelTabButton ' (VB) virtual calls from here on reach WzPanelTabButton's overrides
End Sub

Public Function WzPanelTabButton_make(ByVal text As String) As WIDGET
    Dim button As WIDGET, minButtonWidthForText As Long
    Set button = WIDGET_alloc(K_WzPanelTabButton)
    WzPanelTabButton_construct button
    button.setString text
    W_BUTTON_(button).FontID = font_regular
    minButtonWidthForText = iV_GetTextWidth(text, W_BUTTON_(button).FontID)
    button.setGeometry 0, 0, minButtonWidthForText + QCFORM_PANEL_TABS_PADDING, QCFORM_PANEL_TABS_HEIGHT
    Set WzPanelTabButton_make = button
End Function

Public Sub WzPanelTabButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzPanelTabButton, x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean, isDown As Boolean
    Dim isDisabled As Boolean, isHighlight As Boolean, light_border As Long, fill_color As Long, border As Long
    Dim fw As Long, fx As Long, fy As Long, textColor As Long
    Set d = WzPanelTabButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(W_BUTTON_(self).pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    ' Display the button.
    light_border = WZCOL_PANELTABS_BORDER_LIGHT
    If isDown Or isDisabled Then fill_color = WZCOL_PANELTABS_FILL_COLOR_DARK Else fill_color = WZCOL_PANELTABS_FILL_COLOR
    If isDown Or isHighlight Then border = light_border Else border = psPalette(WZCOL_FORM_DARK)
    iV_ShadowBox x0, y0, x1, y1, 0, border, border, fill_color

    If haveText Then
        d.wzText.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID
        fw = d.wzText.width()
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()
        If isDown Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_FORM_TEXT)
        If isDisabled Then
            textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
        End If
        d.wzText.render fx, fy, textColor
    End If
End Sub
