Attribute VB_Name = "button_c"
' lib/widget/button.cpp, button.h: the button widgets W_BUTTON and
' MultipleChoiceButton (their fields: W_BUTTON.cls, MultipleChoiceButton.cls),
' W_BUTINIT's constructor, the corner button and PopoverMenuButtonDisplayFunc
' (their caches: CornerButtonDisplayCache.cls, PopoverMenuButtonDisplayCache.cls).
' W_BUTTON::Images is the UDT W_BUTTON_Images (W_BUTTON_Images_(...)), W_BUTTON::
' ProgressBorder W_BUTTON_ProgressBorder with its BorderInset (its static
' constructors W_BUTTON_ProgressBorder_indeterminate / _timed / _fixedFactor,
' its methods W_BUTTON_ProgressBorder_<method>(pb, ...)). A W_BUTTON_ONCLICK_FUNC
' is a WzFunc called with the button.
' (VB) widget class W_BUTTON : WIDGET
' (VB) virtual W_BUTTON.clickHeld
' (VB) widget class MultipleChoiceButton : W_BUTTON
Option Explicit

' struct W_BUTTON::Images
Public Type W_BUTTON_Images
    normal As AtlasImage                    ' The image for the button.
    down As AtlasImage                      ' The image for the button, when down. Is overlaid over image.
    highlighted As AtlasImage               ' The image for the button, when highlighted. Is overlaid over image.
    disabled As AtlasImage                  ' The image for the button, when disabled. Is overlaid over image.
End Type

' struct W_BUTTON::ProgressBorder::BorderInset (BorderInset() is all zero)
Public Type W_BUTTON_ProgressBorder_BorderInset
    left_ As Long                           ' = 0
    top As Long                             ' = 0
    right_ As Long                          ' = 0
    bottom As Long                          ' = 0
End Type

' struct W_BUTTON::ProgressBorder
Public Type W_BUTTON_ProgressBorder
    m_inset As W_BUTTON_ProgressBorder_BorderInset
    m_interval As Long                      ' UDWORD = 0
    m_startingTime As Long                  ' UDWORD = 0
    has_m_factor As Boolean                 ' optional<float> m_factor = 0.0f
    m_factor As Single
    m_repeated As Boolean                   ' = false
End Type

' WzFunc ids (fork W1: 1100-1199 here)
Public Const FN_CornerButtonDisplayFunc As Long = 1100          ' static void CornerButtonDisplayFunc (a WIDGET_DISPLAY)
Public Const FN_CornerButtonOnDelete As Long = 1101             ' makeFormTransparentCornerButton's setOnDelete lambda
Public Const FN_PopoverMenuButtonDisplayFunc As Long = 1102     ' void PopoverMenuButtonDisplayFunc (a WIDGET_DISPLAY)

' the lambdas and function pointers of this module
Public Function button_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET
    Select Case f.id
    Case FN_CornerButtonDisplayFunc
        CornerButtonDisplayFunc args(0), args(1), args(2)
    Case FN_CornerButtonOnDelete
        Set psWidget = args(0)
        ' assert(psWidget->pUserData != nullptr); (NDEBUG in a release build)
        ' delete static_cast<CornerButtonDisplayCache *>(psWidget->pUserData);
        psWidget.pUserData = Empty
    Case FN_PopoverMenuButtonDisplayFunc
        PopoverMenuButtonDisplayFunc args(0), args(1), args(2)
    End Select
End Function

' W_BUTINIT::W_BUTINIT(): pText(nullptr), FontID(font_regular) (a VB UDT starts so: for uniformity)
Public Sub W_BUTINIT_construct(i As W_BUTINIT)
    Dim z As W_BUTINIT
    i = z
    i.pText = vbNullString
    i.FontID = font_regular
End Sub

' (VB) the W_INIT part of a W_BUTINIT (W_BUTINIT : W_INIT)
Public Function W_BUTINIT_toW_INIT(init As W_BUTINIT) As W_INIT
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
    W_BUTINIT_toW_INIT = r
End Function

' Images(AtlasImage normal, AtlasImage down, AtlasImage highlighted, AtlasImage disabled = AtlasImage()) (pass
' AtlasImage_() for the default; Images() is a zeroed W_BUTTON_Images)
Public Function W_BUTTON_Images_(normal As AtlasImage, down As AtlasImage, highlighted As AtlasImage, disabled As AtlasImage) As W_BUTTON_Images
    W_BUTTON_Images_.normal = normal
    W_BUTTON_Images_.down = down
    W_BUTTON_Images_.highlighted = highlighted
    W_BUTTON_Images_.disabled = disabled
End Function

'--------------------------------------------------------------------------
' W_BUTTON

' W_BUTTON::W_BUTTON(W_BUTINIT const *init)
Public Sub W_BUTTON_construct_init(ByVal self As WIDGET, init As W_BUTINIT)
    Dim baseInit As W_INIT, d As W_BUTTON
    baseInit = W_BUTINIT_toW_INIT(init)
    WIDGET_construct_init self, baseInit, WIDG_BUTTON
    self.kind = K_W_BUTTON ' (VB) virtual calls from here on reach W_BUTTON's overrides
    Set d = W_BUTTON_(self)
    d.state = WBUT_PLAIN
    d.pText = init.pText                    ' WzString::fromUtf8(init->pText)
    d.pTip = init.pTip
    d.HilightAudioID = WidgGetHilightAudioID()
    d.ClickedAudioID = WidgGetClickedAudioID()
    Set d.AudioCallback = WidgGetAudioCallback()
    d.FontID = init.FontID
    W_BUTTON_initMembers self
    If Not ((init.style And Not (WBUT_PLAIN Or WIDG_HIDDEN Or WBUT_NOPRIMARY Or WBUT_SECONDARY Or WBUT_TXTCENTRE)) = 0) Then ASSERT_FAILURE "(init->style & ~(0 | WIDG_HIDDEN | 0x10 | 0x20 | 0x40)) == 0", "lib/widget/button.cpp:49", "W_BUTTON", "unknown button style"
End Sub

' W_BUTTON::W_BUTTON()
Public Sub W_BUTTON_construct(ByVal self As WIDGET)
    Dim d As W_BUTTON
    WIDGET_construct self, WIDG_BUTTON
    self.kind = K_W_BUTTON ' (VB) virtual calls from here on reach W_BUTTON's overrides
    Set d = W_BUTTON_(self)
    d.state = WBUT_PLAIN
    d.HilightAudioID = WidgGetHilightAudioID()
    d.ClickedAudioID = WidgGetClickedAudioID()
    Set d.AudioCallback = WidgGetAudioCallback()
    d.FontID = font_regular
    W_BUTTON_initMembers self
End Sub

' the members' initialisers (button.h)
Private Sub W_BUTTON_initMembers(ByVal self As WIDGET)
    Dim d As W_BUTTON
    Set d = W_BUTTON_(self)
    d.minClickInterval = 0
    d.lastClickTime = 0
    d.has_progressBorder = False
    d.lastClickButton = WKEY_NONE
End Sub

' std::make_shared<W_BUTTON>(init)
Public Function W_BUTTON_new_init(init As W_BUTINIT) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_BUTTON)
    W_BUTTON_construct_init w, init
    Set W_BUTTON_new_init = w
End Function

' std::make_shared<W_BUTTON>()
Public Function W_BUTTON_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_BUTTON)
    W_BUTTON_construct w
    Set W_BUTTON_new = w
End Function

Public Function W_BUTTON_getState(ByVal self As WIDGET) As Long
    W_BUTTON_getState = W_BUTTON_(self).state And (WBUT_DISABLE Or WBUT_LOCK Or WBUT_CLICKLOCK Or WBUT_FLASH Or WBUT_DOWN Or WBUT_HIGHLIGHT)
End Function

Public Sub W_BUTTON_setFlash(ByVal self As WIDGET, ByVal enable As Boolean)
    Dim d As W_BUTTON
    Set d = W_BUTTON_(self)
    self.dirty = True
    If enable Then
        d.state = d.state Or WBUT_FLASH
    Else
        d.state = d.state And Not WBUT_FLASH
    End If
End Sub

Public Sub W_BUTTON_unlock(ByVal self As WIDGET)
    self.dirty = True
    W_BUTTON_(self).state = W_BUTTON_(self).state And Not (WBUT_LOCK Or WBUT_CLICKLOCK)
End Sub

Public Sub W_BUTTON_setState(ByVal self As WIDGET, ByVal state As Long)
    Dim d As W_BUTTON, mask As Long
    If Not (Not ((state And WBUT_LOCK) <> 0 And (state And WBUT_CLICKLOCK) <> 0)) Then ASSERT_FAILURE "!((newState & WBUT_LOCK) && (newState & WBUT_CLICKLOCK))", "lib/widget/button.cpp:87", "setState", "Cannot have both WBUT_LOCK and WBUT_CLICKLOCK"

    Set d = W_BUTTON_(self)
    mask = WBUT_DISABLE Or WBUT_LOCK Or WBUT_CLICKLOCK
    d.state = (d.state And Not mask) Or (state And mask)
    self.dirty = True
End Sub

Public Function W_BUTTON_getString(ByVal self As WIDGET) As String
    W_BUTTON_getString = W_BUTTON_(self).pText
End Function

Public Sub W_BUTTON_setString(ByVal self As WIDGET, ByVal string_ As String)
    W_BUTTON_(self).pText = string_
    self.dirty = True
End Sub

Public Sub W_BUTTON_setTip(ByVal self As WIDGET, ByVal string_ As String)
    W_BUTTON_(self).pTip = string_
End Sub

Public Sub W_BUTTON_setHelp(ByVal self As WIDGET, ByVal help As WidgetHelp)
    Set W_BUTTON_(self).help = help
End Sub

Public Function W_BUTTON_getTip(ByVal self As WIDGET) As String
    W_BUTTON_getTip = W_BUTTON_(self).pTip
End Function

' WidgetHelp const * getHelp() const (Nothing if !help.has_value())
Public Function W_BUTTON_getHelp(ByVal self As WIDGET) As WidgetHelp
    Set W_BUTTON_getHelp = W_BUTTON_(self).help
End Function

Public Sub W_BUTTON_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_BUTTON, now_ As Double
    Set d = W_BUTTON_(self)
    If d.has_clickDownStart Then
        now_ = widgSteadyNowMs()
        If now_ - d.clickDownStart >= widgGetClickHoldMS() Then
            If d.has_clickDownKey Then
                If vt_W_BUTTON_clickHeld(self, psContext, d.clickDownKey) Then
                    ' clear button down state, as the clickHeld event "consumed" this click
                    d.state = d.state And Not WBUT_DOWN
                End If
            End If
            d.has_clickDownStart = False
        End If
    End If
End Sub

' Returns true if "consumed" held click
Public Function W_BUTTON_clickHeld(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY) As Boolean
    W_BUTTON_clickHeld = False
End Function

Public Sub W_BUTTON_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_BUTTON
    Set d = W_BUTTON_(self)
    If (d.minClickInterval > 0) And (realTime - d.lastClickTime < d.minClickInterval) Then
        Exit Sub
    End If
    d.lastClickTime = realTime

    self.dirty = True

    ' Can't click a button if it is disabled or locked down
    If (d.state And (WBUT_DISABLE Or WBUT_LOCK)) = 0 Then
        ' Check this is the correct key
        If ((self.style And WBUT_NOPRIMARY) = 0 And key = WKEY_PRIMARY) Or _
           ((self.style And WBUT_SECONDARY) <> 0 And key = WKEY_SECONDARY) Then
            If Not d.AudioCallback Is Nothing Then
                WzFunc_call d.AudioCallback, d.ClickedAudioID
            End If
            d.state = d.state And Not WBUT_FLASH ' Stop it flashing
            d.state = d.state Or WBUT_DOWN

            d.clickDownStart = widgSteadyNowMs()
            d.has_clickDownStart = True
            d.clickDownKey = key
            d.has_clickDownKey = True
        End If
    End If
End Sub

' Respond to a mouse button up
Public Sub W_BUTTON_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_BUTTON, i As Long, onClickHandler As WzFunc, lockedScreen As W_SCREEN
    Set d = W_BUTTON_(self)
    If (d.state And WBUT_DOWN) <> 0 Then
        ' Check this is the correct key
        If ((self.style And WBUT_NOPRIMARY) = 0 And key = WKEY_PRIMARY) Or _
           ((self.style And WBUT_SECONDARY) <> 0 And key = WKEY_SECONDARY) Then
            d.lastClickButton = key

            ' Call all onClick event handlers
            i = 1
            Do While i <= d.onClickHandlers.Count
                Set onClickHandler = d.onClickHandlers(i)
                If Not onClickHandler Is Nothing Then
                    WzFunc_call onClickHandler, self
                End If
                i = i + 1
            Loop

            d.lastClickButton = WKEY_NONE

            Set lockedScreen = self.screenPointer
            If Not lockedScreen Is Nothing Then
                lockedScreen.setReturn self
            End If
            d.state = d.state And Not WBUT_DOWN
            self.dirty = True
        End If
    End If

    d.has_clickDownStart = False
    d.has_clickDownKey = False
End Sub

' May be called from within W_BUTTON_ONCLICK_FUNC (onclick handler) to retrieve the button pressed
Public Function W_BUTTON_getOnClickButtonPressed(ByVal self As WIDGET) As WIDGET_KEY
    W_BUTTON_getOnClickButtonPressed = W_BUTTON_(self).lastClickButton
End Function

Public Function W_BUTTON_isHighlighted(ByVal self As WIDGET) As Boolean
    W_BUTTON_isHighlighted = (W_BUTTON_(self).state And WBUT_HIGHLIGHT) <> 0
End Function

' Respond to a mouse moving over a button
Public Sub W_BUTTON_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_BUTTON
    Set d = W_BUTTON_(self)
    If (d.state And WBUT_HIGHLIGHT) = 0 Then
        d.state = d.state Or WBUT_HIGHLIGHT
        self.dirty = True
    End If
    If Not d.AudioCallback Is Nothing Then
        WzFunc_call d.AudioCallback, d.HilightAudioID
    End If
End Sub

' Respond to the mouse moving off a button
Public Sub W_BUTTON_highlightLost(ByVal self As WIDGET)
    Dim d As W_BUTTON
    Set d = W_BUTTON_(self)
    d.state = d.state And Not (WBUT_DOWN Or WBUT_HIGHLIGHT)
    self.dirty = True

    d.has_clickDownStart = False
    d.has_clickDownKey = False
End Sub

Public Sub W_BUTTON_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_BUTTON, x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, images As W_BUTTON_Images
    Dim fw As Long, fx As Long, fy As Long, secondCol As Long
    Set d = W_BUTTON_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(d.pText) <> 0

    isDown = (d.state And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (d.state And WBUT_DISABLE) <> 0
    isHighlight = (d.state And WBUT_HIGHLIGHT) <> 0
    images = d.images

    ' Display the button.
    If Not AtlasImage_isNull(images.normal) Then
        iV_DrawImageImage images.normal.images, images.normal.id, x0, y0
        If isDown And Not AtlasImage_isNull(images.down) Then
            iV_DrawImageImage images.down.images, images.down.id, x0, y0
        End If
        If isDisabled And Not AtlasImage_isNull(images.disabled) Then
            iV_DrawImageImage images.disabled.images, images.disabled.id, x0, y0
        End If
        If isHighlight And Not AtlasImage_isNull(images.highlighted) Then
            iV_DrawImageImage images.highlighted.images, images.highlighted.id, x0, y0
        End If
    Else
        If isDisabled Then secondCol = psPalette(WZCOL_FORM_LIGHT) Else secondCol = psPalette(WZCOL_FORM_DARK)
        iV_ShadowBox x0, y0, x1, y1, 0, psPalette(WZCOL_FORM_LIGHT), secondCol, psPalette(WZCOL_FORM_BACKGROUND)
        If isHighlight Then
            iV_Box x0 + 2, y0 + 2, x1 - 3, y1 - 3, psPalette(WZCOL_FORM_HILITE)
        End If
    End If

    If haveText Then
        fw = iV_GetTextWidth(d.pText, d.FontID)
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - iV_GetTextLineSize(d.FontID)) \ 2 - iV_GetTextAboveBase(d.FontID)
        If isDisabled Then
            iV_SetTextColour psPalette(WZCOL_FORM_LIGHT)
            iV_DrawText d.pText, fx + 1, fy + 1, d.FontID
            iV_SetTextColour psPalette(WZCOL_FORM_DISABLE)
        Else
            iV_SetTextColour psPalette(WZCOL_FORM_TEXT)
        End If
        iV_DrawText d.pText, fx, fy, d.FontID
    End If

    If isDisabled And Not AtlasImage_isNull(images.normal) And AtlasImage_isNull(images.disabled) Then
        ' disabled, render something over it!
        iV_TransBoxFill x0, y0, x0 + self.width(), y0 + self.height()
    End If
End Sub

' for handling progress border overlay
Public Sub W_BUTTON_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    ' call parent displayRecursive
    WIDGET_displayRecursive self, context

    If W_BUTTON_(self).has_progressBorder Then
        ' "over-draw" with the progress border
        W_BUTTON_drawProgressBorder self, WidgetGraphicsContext_getXOffset(context), WidgetGraphicsContext_getYOffset(context)
    End If
End Sub

' (VB) a glm::ivec4 of a line list: lines(4 * i .. 4 * i + 3) = x, y, z, w
Private Sub setLine(lines() As Long, ByVal i As Long, ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal w As Long)
    lines(4 * i) = x
    lines(4 * i + 1) = y
    lines(4 * i + 2) = z
    lines(4 * i + 3) = w
End Sub

Private Sub W_BUTTON_drawProgressBorder(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_BUTTON, config As W_BUTTON_ProgressBorder, g As WzRect, x0 As Long, y0 As Long
    Dim widthReduction As Long, heightReduction As Long, progressBorderRect As WzRect
    Dim borderLines(0 To 23) As Long, lengthOfEachLine(0 To 5) As Long, borderLineCount As Long
    Dim realLineCount As Long, totalLength As Long, indeterminateMaxLineLength As Long, i As Long
    Dim timeElapsed As Long, factor As Single, progressLength As Long, remainingLength As Long
    Dim linesToDraw(0 To 23) As Long, linesToDrawCount As Long, lineIdx As Long, lineLength As Long
    Dim isVerticalLine As Boolean, extraLength As Long, v(0 To 3) As Long, k As Long
    Const FIXED_INDETERMINATE_MODE_INTERVAL As Long = 4 * GAME_TICKS_PER_SEC
    Set d = W_BUTTON_(self)
    If Not d.has_progressBorder Then ASSERT_FAILURE "progressBorder.has_value()", "lib/widget/button.cpp:320", "drawProgressBorder", "progressBorder is nullopt": Exit Sub
    config = d.progressBorder

    ' Determine total rect to be used for the progress border
    g = self.geometry()
    x0 = g.x + xOffset
    y0 = g.y + yOffset
    widthReduction = config.m_inset.left_ + config.m_inset.right_
    heightReduction = config.m_inset.top + config.m_inset.bottom
    progressBorderRect = WzRect_(x0 + config.m_inset.left_, y0 + config.m_inset.top, g.w - widthReduction, g.h - heightReduction)

    ' Determine border lines
    ' THERE. ARE. 5. LINES!!! (because the "top" line is split in two - the starting point is the middle of the top line, and proceeds clockwise)
    setLine borderLines, 0, progressBorderRect.x + (progressBorderRect.w \ 2), progressBorderRect.y, progressBorderRect.x + progressBorderRect.w, progressBorderRect.y
    lengthOfEachLine(0) = borderLines(2) - borderLines(0)
    setLine borderLines, 1, progressBorderRect.x + progressBorderRect.w, progressBorderRect.y, progressBorderRect.x + progressBorderRect.w, progressBorderRect.y + progressBorderRect.h
    lengthOfEachLine(1) = borderLines(7) - borderLines(5)
    setLine borderLines, 2, progressBorderRect.x + progressBorderRect.w, progressBorderRect.y + progressBorderRect.h, progressBorderRect.x, progressBorderRect.y + progressBorderRect.h
    lengthOfEachLine(2) = progressBorderRect.w
    setLine borderLines, 3, progressBorderRect.x, progressBorderRect.y + progressBorderRect.h, progressBorderRect.x, progressBorderRect.y
    lengthOfEachLine(3) = lengthOfEachLine(1)
    setLine borderLines, 4, progressBorderRect.x, progressBorderRect.y, progressBorderRect.x + (progressBorderRect.w \ 2) + 1, progressBorderRect.y
    lengthOfEachLine(4) = borderLines(18) - borderLines(16)
    borderLineCount = 5
    realLineCount = borderLineCount
    totalLength = 0
    For i = 0 To borderLineCount - 1
        totalLength = totalLength + lengthOfEachLine(i)
    Next
    indeterminateMaxLineLength = CLng(Fix(CSng(totalLength) / 4!))
    If W_BUTTON_ProgressBorder_isIndeterminate(config) Then
        ' "Repeat" indeterminate mode (progress goes all the way to 0 and restarts)
        ' Increase totalLength by indeterminateMaxLineLength, and add a "fake" line at the end of the borderLine list that we don't actually draw
        totalLength = totalLength + indeterminateMaxLineLength
        setLine borderLines, 5, borderLines(0), borderLines(1), borderLines(0) + indeterminateMaxLineLength, borderLines(3)
        lengthOfEachLine(5) = indeterminateMaxLineLength
        borderLineCount = 6
    End If
    timeElapsed = realTime - config.m_startingTime

    factor = 0!
    If config.has_m_factor Then
        factor = config.m_factor
    ElseIf config.m_interval <> 0 Then
        factor = CSng(timeElapsed) / CSng(config.m_interval)
    Else
        ' Indeterminate mode
        factor = CSng(timeElapsed) / CSng(FIXED_INDETERMINATE_MODE_INTERVAL)
    End If

    ' Get length of progress
    If totalLength > 0 Then progressLength = CLng(Fix(factor * CSng(totalLength))) Mod totalLength Else progressLength = 0

    If Not config.has_m_factor And Not config.m_repeated And timeElapsed >= config.m_interval Then
        ' Done - draw once at full, and then stop drawing in the future
        progressLength = totalLength
        d.has_progressBorder = False
    End If

    ' Decide what parts of what lines to draw
    remainingLength = progressLength
    linesToDrawCount = 0
    lineIdx = 0
    Do While lineIdx < borderLineCount And remainingLength > 0
        lineLength = lengthOfEachLine(lineIdx)
        If remainingLength >= lineLength Then
            setLine linesToDraw, linesToDrawCount, borderLines(4 * lineIdx), borderLines(4 * lineIdx + 1), borderLines(4 * lineIdx + 2), borderLines(4 * lineIdx + 3)
            linesToDrawCount = linesToDrawCount + 1
            remainingLength = remainingLength - lineLength
        Else
            ' Last line - partial line - reduce end-point
            For k = 0 To 3
                v(k) = borderLines(4 * lineIdx + k)
            Next
            isVerticalLine = (lineIdx Mod 2) <> 0
            extraLength = lineLength - remainingLength
            If Not isVerticalLine Then
                If v(2) > v(0) Then
                    v(2) = v(2) - extraLength
                Else
                    v(2) = v(2) + extraLength
                End If
            Else
                If v(3) > v(1) Then
                    v(3) = v(3) - extraLength
                Else
                    v(3) = v(3) + extraLength
                End If
            End If
            setLine linesToDraw, linesToDrawCount, v(0), v(1), v(2), v(3)
            linesToDrawCount = linesToDrawCount + 1
            lengthOfEachLine(lineIdx) = remainingLength ' update line length (for indeterminate mode case, below)
            remainingLength = 0
        End If
        lineIdx = lineIdx + 1
    Loop

    If W_BUTTON_ProgressBorder_isIndeterminate(config) And linesToDrawCount > 0 Then
        ' In indeterminate mode, start from the end and remove any lines beyond the indeterminateMaxLineLength
        remainingLength = indeterminateMaxLineLength
        lineIdx = linesToDrawCount - 1
        Do While lineIdx >= 0
            lineLength = lengthOfEachLine(lineIdx)
            If remainingLength >= lineLength Then
                remainingLength = remainingLength - lineLength
            Else
                ' Shorten this line by adjusting its *starting point*
                For k = 0 To 3
                    v(k) = linesToDraw(4 * lineIdx + k)
                Next
                isVerticalLine = (lineIdx Mod 2) <> 0
                extraLength = lineLength - remainingLength
                If Not isVerticalLine Then
                    If v(2) > v(0) Then
                        v(0) = v(0) + extraLength
                    Else
                        v(0) = v(0) - extraLength
                    End If
                Else
                    If v(3) > v(1) Then
                        v(1) = v(1) + extraLength
                    Else
                        v(1) = v(1) - extraLength
                    End If
                End If
                setLine linesToDraw, lineIdx, v(0), v(1), v(2), v(3)
                remainingLength = 0
                Exit Do
            End If
            lineIdx = lineIdx - 1
        Loop
        ' "Repeat" indeterminate mode: remove the last line *if* if it's the "fake" line for indeterminate mode calculations)
        If linesToDrawCount > realLineCount Then
            linesToDrawCount = linesToDrawCount - 1
        End If
        ' Remove any lines earlier in the list than lineIdx
        If lineIdx > 0 Then
            For i = lineIdx To linesToDrawCount - 1
                setLine linesToDraw, i - lineIdx, linesToDraw(4 * i), linesToDraw(4 * i + 1), linesToDraw(4 * i + 2), linesToDraw(4 * i + 3)
            Next
            linesToDrawCount = linesToDrawCount - lineIdx
        End If
    End If

    If linesToDrawCount > 0 Then
        ' Draw the lines
        iV_Lines VarPtr(linesToDraw(0)), linesToDrawCount, d.progressBorderColour
    End If
End Sub

' ProgressBorder(UDWORD interval, optional<float> factor, bool repeated, BorderInset inset = BorderInset()) (private)
Private Function W_BUTTON_ProgressBorder_make(ByVal interval As Long, ByVal has_factor As Boolean, ByVal factor As Single, ByVal repeated As Boolean, inset As W_BUTTON_ProgressBorder_BorderInset) As W_BUTTON_ProgressBorder
    Dim pb As W_BUTTON_ProgressBorder
    pb.m_inset = inset
    pb.m_interval = interval
    pb.has_m_factor = has_factor
    pb.m_factor = factor
    pb.m_repeated = repeated
    W_BUTTON_ProgressBorder_resetStartingTime pb
    W_BUTTON_ProgressBorder_make = pb
End Function

' BorderInset(int left, int top, int right, int bottom) (all omitted: BorderInset())
Public Function W_BUTTON_ProgressBorder_BorderInset_(Optional ByVal left_ As Long = 0, Optional ByVal top As Long = 0, Optional ByVal right_ As Long = 0, Optional ByVal bottom As Long = 0) As W_BUTTON_ProgressBorder_BorderInset
    W_BUTTON_ProgressBorder_BorderInset_.left_ = left_
    W_BUTTON_ProgressBorder_BorderInset_.top = top
    W_BUTTON_ProgressBorder_BorderInset_.right_ = right_
    W_BUTTON_ProgressBorder_BorderInset_.bottom = bottom
End Function

' Create an indeterminate ProgressBorder
Public Function W_BUTTON_ProgressBorder_indeterminate(inset As W_BUTTON_ProgressBorder_BorderInset) As W_BUTTON_ProgressBorder
    W_BUTTON_ProgressBorder_indeterminate = W_BUTTON_ProgressBorder_make(0, False, 0!, True, inset)
End Function

' Create a timed ProgressBorder which proceeds from 0 to 100% over the interval
Public Function W_BUTTON_ProgressBorder_timed(ByVal interval As Long, ByVal repeated As Boolean, inset As W_BUTTON_ProgressBorder_BorderInset) As W_BUTTON_ProgressBorder
    W_BUTTON_ProgressBorder_timed = W_BUTTON_ProgressBorder_make(interval, False, 0!, repeated, inset)
End Function

' Create a fixed factor ProgressBorder, where `factor` is expected to be a float from 0.0 to 1.0 (which corresponds to the percentage progress)
Public Function W_BUTTON_ProgressBorder_fixedFactor(ByVal factor As Single, inset As W_BUTTON_ProgressBorder_BorderInset) As W_BUTTON_ProgressBorder
    W_BUTTON_ProgressBorder_fixedFactor = W_BUTTON_ProgressBorder_make(0, True, factor, False, inset)
End Function

Public Sub W_BUTTON_ProgressBorder_setInset(pb As W_BUTTON_ProgressBorder, newInset As W_BUTTON_ProgressBorder_BorderInset)
    pb.m_inset = newInset
End Sub

Public Sub W_BUTTON_ProgressBorder_resetStartingTime(pb As W_BUTTON_ProgressBorder)
    pb.m_startingTime = realTime
End Sub

Public Function W_BUTTON_ProgressBorder_isIndeterminate(pb As W_BUTTON_ProgressBorder) As Boolean
    W_BUTTON_ProgressBorder_isIndeterminate = Not pb.has_m_factor And pb.m_interval = 0
End Function

Public Sub W_BUTTON_setImages(ByVal self As WIDGET, images_ As W_BUTTON_Images)
    Dim d As W_BUTTON, images As W_BUTTON_Images
    Set d = W_BUTTON_(self)
    d.images = images_
    images = d.images
    self.dirty = True
    If Not AtlasImage_isNull(images.normal) Then
        self.setGeometry self.x(), self.y(), AtlasImage_width(images.normal.images, images.normal.id), AtlasImage_height(images.normal.images, images.normal.id)
    End If
End Sub

' void setImages(AtlasImage image, AtlasImage imageDown, AtlasImage imageHighlight, AtlasImage imageDisabled = AtlasImage())
' (pass AtlasImage_() for the default)
Public Sub W_BUTTON_setImages_atlas(ByVal self As WIDGET, image As AtlasImage, imageDown As AtlasImage, imageHighlight As AtlasImage, imageDisabled As AtlasImage)
    Dim images As W_BUTTON_Images
    self.dirty = True
    images = W_BUTTON_Images_(image, imageDown, imageHighlight, imageDisabled)
    W_BUTTON_setImages self, images
End Sub

' void addOnClickHandler(const W_BUTTON_ONCLICK_FUNC& onClickFunc): onClickFunc is called with the button
Public Sub W_BUTTON_addOnClickHandler(ByVal self As WIDGET, ByVal onClickFunc As WzFunc)
    W_BUTTON_(self).onClickHandlers.Add onClickFunc
End Sub

' Set the current ProgressBorder on the button, optionally specifying a color (borderColour missing: nullopt)
' To disable a ProgressBorder, use `setProgressBorder(nullopt)`: W_BUTTON_setProgressBorder_nullopt
Public Sub W_BUTTON_setProgressBorder(ByVal self As WIDGET, progressBorder As W_BUTTON_ProgressBorder, Optional ByVal borderColour As Variant)
    Dim d As W_BUTTON
    Set d = W_BUTTON_(self)
    d.progressBorder = progressBorder
    d.has_progressBorder = True
    If Not IsMissing(borderColour) Then d.progressBorderColour = CLng(borderColour) Else d.progressBorderColour = psPalette(WZCOL_FORM_HILITE)
End Sub

' setProgressBorder(nullopt, borderColour)
Public Sub W_BUTTON_setProgressBorder_nullopt(ByVal self As WIDGET, Optional ByVal borderColour As Variant)
    Dim d As W_BUTTON
    Set d = W_BUTTON_(self)
    d.has_progressBorder = False
    If Not IsMissing(borderColour) Then d.progressBorderColour = CLng(borderColour) Else d.progressBorderColour = psPalette(WZCOL_FORM_HILITE)
End Sub

'--------------------------------------------------------------------------
' MultipleChoiceButton

' MultipleChoiceButton() : W_BUTTON(), choice(0) {}
Public Sub MultipleChoiceButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_MultipleChoiceButton ' (VB) virtual calls from here on reach MultipleChoiceButton's overrides
    MultipleChoiceButton_(self).choice = 0
End Sub

Public Function MultipleChoiceButton_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_MultipleChoiceButton)
    MultipleChoiceButton_construct w
    Set MultipleChoiceButton_new = w
End Function

Public Sub MultipleChoiceButton_setChoice(ByVal self As WIDGET, ByVal newChoice As Long)
    Dim d As MultipleChoiceButton, image As W_BUTTON_Images
    Set d = MultipleChoiceButton_(self)
    If d.choice = newChoice Then
        Exit Sub
    End If
    self.dirty = True
    d.choice = newChoice
    If d.imageSets_find(d.choice, image) Then
        W_BUTTON_setImages self, image
    End If
    If d.tips.Exists(d.choice) Then
        W_BUTTON_setTip self, d.tips(d.choice)
    End If
End Sub

' void setTip(unsigned choiceValue, const std::string& string) (and the char const * overload: vbNullString is
' nullptr's empty string)
Public Sub MultipleChoiceButton_setTip_choice(ByVal self As WIDGET, ByVal choiceValue As Long, ByVal string_ As String)
    Dim d As MultipleChoiceButton
    Set d = MultipleChoiceButton_(self)
    d.tips(choiceValue) = string_
    If d.choice = choiceValue Then
        W_BUTTON_setTip self, string_
    End If
End Sub

' void setImages(unsigned choiceValue, Images const &stateImages)
Public Sub MultipleChoiceButton_setImages(ByVal self As WIDGET, ByVal choiceValue As Long, stateImages As W_BUTTON_Images)
    Dim d As MultipleChoiceButton
    Set d = MultipleChoiceButton_(self)
    d.imageSets_set choiceValue, stateImages
    self.dirty = True
    If d.choice = choiceValue Then
        W_BUTTON_setImages self, stateImages
    End If
End Sub

Public Function MultipleChoiceButton_getChoice(ByVal self As WIDGET) As Long
    MultipleChoiceButton_getChoice = MultipleChoiceButton_(self).choice
End Function

'--------------------------------------------------------------------------
' MARK: - Corner Button

Private Sub CornerButtonDisplayFunc(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim cache As CornerButtonDisplayCache, psButton As WIDGET, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Dim haveText As Boolean, isDisabled As Boolean, isHighlight As Boolean, fw As Long, fx As Long, fy As Long, textColor As Long
    ' Any widget using CornerButtonDisplayFunc must have its pUserData initialized to a (CornerButtonDisplayCache*)
    ' assert(psWidget->pUserData != nullptr); (NDEBUG in a release build)
    Set cache = psWidget.pUserData

    If WIDGET_isKindOf(psWidget, K_W_BUTTON) Then Set psButton = psWidget
    If psButton Is Nothing Then ASSERT_FAILURE "psButton", "lib/widget/button.cpp:591", "CornerButtonDisplayFunc", "psWidget is null": Exit Sub

    x0 = psButton.x() + xOffset
    y0 = psButton.y() + yOffset
    x1 = x0 + psButton.width()
    y1 = y0 + psButton.height()

    haveText = Len(W_BUTTON_(psButton).pText) <> 0

    ' bool isDown = (psButton->getState() & (WBUT_DOWN | WBUT_LOCK | WBUT_CLICKLOCK)) != 0;
    isDisabled = (psButton.getState() And WBUT_DISABLE) <> 0
    isHighlight = (psButton.getState() And WBUT_HIGHLIGHT) <> 0

    ' Display the button.
    If isHighlight Then
        pie_UniTransBoxFill CSng(x0 + 1), CSng(y0 + 1), CSng(x1 - 1), CSng(y1 - 1), cache.buttonBackgroundFill
        iV_Box x0 + 1, y0 + 1, x1 - 1, y1 - 1, pal_RGBA(255, 255, 255, 150)
    End If

    If haveText Then
        cache.text.setText W_BUTTON_(psButton).pText, W_BUTTON_(psButton).FontID
        fw = cache.text.width()
        fx = x0 + (psButton.width() - fw) \ 2
        fy = y0 + (psButton.height() - cache.text.lineSize()) \ 2 - cache.text.aboveBase()
        textColor = psPalette(WZCOL_FORM_TEXT)
        If isDisabled Then
            cache.text.render fx + 1, fy + 1, psPalette(WZCOL_FORM_LIGHT)
            textColor = psPalette(WZCOL_FORM_DISABLE)
        End If
        cache.text.render fx, fy, textColor
    End If

    If isDisabled Then
        ' disabled, render something over it!
        iV_TransBoxFill x0, y0, x0 + psButton.width(), y0 + psButton.height()
    End If
End Sub

' std::shared_ptr<W_BUTTON> makeFormTransparentCornerButton(const char* text, int buttonPadding /* = TAB_BUTTONS_PADDING */, PIELIGHT buttonBackgroundFill /* = WZCOL_DEBUG_FILL_COLOR */)
Public Function makeFormTransparentCornerButton(ByVal text As String, ByVal buttonPadding As Long, ByVal buttonBackgroundFill As Long) As WIDGET
    Dim button As WIDGET, cache As CornerButtonDisplayCache, minButtonWidthForText As Long, buttonSize As Long
    Set button = W_BUTTON_new()
    button.setString text
    W_BUTTON_(button).FontID = font_regular_bold
    Set button.displayFunction = WzFunc_new(FN_CornerButtonDisplayFunc)
    Set cache = New CornerButtonDisplayCache
    cache.buttonBackgroundFill = buttonBackgroundFill
    Set button.pUserData = cache
    button.setOnDelete WzFunc_new(FN_CornerButtonOnDelete)
    minButtonWidthForText = iV_GetTextWidth(text, W_BUTTON_(button).FontID)
    buttonSize = MAX_(MAX_(minButtonWidthForText + buttonPadding, 18), iV_GetTextLineSize(W_BUTTON_(button).FontID))
    button.setGeometry 0, 0, buttonSize, buttonSize
    Set makeFormTransparentCornerButton = button
End Function

' void PopoverMenuButtonDisplayFunc(WIDGET *psWidget, UDWORD xOffset, UDWORD yOffset) (a WIDGET_DISPLAY:
' FN_PopoverMenuButtonDisplayFunc)
Public Sub PopoverMenuButtonDisplayFunc(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim cache As PopoverMenuButtonDisplayCache, psButton As WIDGET, x0 As Long, y0 As Long
    Dim haveText As Boolean, isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean
    Dim backgroundColor As Long, fx As Long, fy As Long, textColor As Long
    ' Any widget using PopoverMenuButtonDisplayFunc must have its pUserData initialized to a (PopoverMenuButtonDisplayCache*)
    ' assert(psWidget->pUserData != nullptr); (NDEBUG in a release build)
    Set cache = psWidget.pUserData

    If WIDGET_isKindOf(psWidget, K_W_BUTTON) Then Set psButton = psWidget
    If psButton Is Nothing Then ASSERT_FAILURE "psButton", "lib/widget/button.cpp:658", "PopoverMenuButtonDisplayFunc", "psWidget is null": Exit Sub

    x0 = psButton.x() + xOffset
    y0 = psButton.y() + yOffset

    haveText = Len(W_BUTTON_(psButton).pText) <> 0

    isDown = (psButton.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (psButton.getState() And WBUT_DISABLE) <> 0
    isHighlight = (psButton.getState() And WBUT_HIGHLIGHT) <> 0

    ' Display the button background
    backgroundColor = 0                         ' PIELIGHT backgroundColor; (rgba 0: transparent)
    If isDown Then
        backgroundColor = pal_RGBA(10, 0, 70, 250) 'WZCOL_FORM_DARK;
    ElseIf isHighlight Then
        backgroundColor = pal_RGBA(25, 0, 110, 220) 'WZCOL_TEXT_MEDIUM;
    End If
    If (backgroundColor And &HFF000000) <> 0 Then ' !backgroundColor.isTransparent()
        ' Draw the background
        pie_UniTransBoxFill x0, y0, x0 + psButton.width(), y0 + psButton.height(), backgroundColor
    End If

    If haveText Then
        cache.text.setText W_BUTTON_(psButton).pText, W_BUTTON_(psButton).FontID
        fx = x0 + (psButton.width() - cache.text.width()) \ 2
        fy = y0 + (psButton.height() - cache.text.lineSize()) \ 2 - cache.text.aboveBase()
        textColor = psPalette(WZCOL_FORM_TEXT)
        If isDisabled Then
            cache.text.render fx + 1, fy + 1, psPalette(WZCOL_FORM_LIGHT)
            textColor = psPalette(WZCOL_FORM_DISABLE)
        End If
        cache.text.render fx, fy, textColor
    End If

    If isDisabled Then
        ' disabled, render something over it!
        iV_TransBoxFill x0, y0, x0 + psButton.width(), y0 + psButton.height()
    End If
End Sub
