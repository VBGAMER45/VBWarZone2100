Attribute VB_Name = "form_c"
' lib/widget/form.cpp, form.h: the form widgets W_FORM, W_CLICKFORM and
' W_FULLSCREENOVERLAY_CLICKFORM (their fields: W_FORM.cls, W_CLICKFORM.cls,
' W_FULLSCREENOVERLAY_CLICKFORM.cls), W_FORMINIT's constructor and
' displayChildDropShadows (also a WIDGET_DISPLAY: FN_displayChildDropShadows).
' W_FULLSCREENOVERLAY_CLICKFORM::make is W_FULLSCREENOVERLAY_CLICKFORM_make.
' (VB) widget class W_FORM : WIDGET
' (VB) widget class W_CLICKFORM : W_FORM
' (VB) virtual W_CLICKFORM.clickHeld
' (VB) widget class W_FULLSCREENOVERLAY_CLICKFORM : W_CLICKFORM
Option Explicit

Private Const WZ_FORM_MINIMIZED_LEFT_PADDING As Long = 5
Private Const WZ_FORM_MINIMIZED_MAXBUTTON_RIGHT_PADDING As Long = WZ_FORM_MINIMIZED_LEFT_PADDING

' W_FORM::FormState (its values FormState_NORMAL, FormState_MINIMIZED)
Public Enum W_FORM_FormState
    FormState_NORMAL
    FormState_MINIMIZED
End Enum

' WzFunc ids (fork W1: 1000-1099 here)
Public Const FN_W_FULLSCREENOVERLAY_CLICKFORM_calcLayout As Long = 1000  ' make's LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_displayChildDropShadows As Long = 1001                  ' displayChildDropShadows as a WIDGET_DISPLAY

' the lambdas and function pointers of this module
Public Function form_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET
    Select Case f.id
    Case FN_W_FULLSCREENOVERLAY_CLICKFORM_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight
    Case FN_displayChildDropShadows
        displayChildDropShadows args(0), args(1), args(2)
    End Select
End Function

' W_FORMINIT::W_FORMINIT(): disableChildren(false) (a VB UDT starts zeroed: for uniformity)
Public Sub W_FORMINIT_construct(i As W_FORMINIT)
    Dim z As W_FORMINIT
    i = z
    i.disableChildren = False
End Sub

' (VB) the W_INIT part of a W_FORMINIT (W_FORMINIT : W_INIT)
Public Function W_FORMINIT_toW_INIT(init As W_FORMINIT) As W_INIT
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
    W_FORMINIT_toW_INIT = r
End Function

'--------------------------------------------------------------------------
' W_FORM

' W_FORM::W_FORM(W_FORMINIT const *init)
Public Sub W_FORM_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    Dim baseInit As W_INIT
    baseInit = W_FORMINIT_toW_INIT(init)
    WIDGET_construct_init self, baseInit, WIDG_FORM
    self.kind = K_W_FORM ' (VB) virtual calls from here on reach W_FORM's overrides
    W_FORM_initMembers self
    W_FORM_(self).disableChildren = init.disableChildren
    If Not ((init.style And Not (WFORM_INVISIBLE Or WFORM_CLICKABLE Or WFORM_NOCLICKMOVE Or WFORM_NOPRIMARY Or WFORM_SECONDARY Or WIDG_HIDDEN)) = 0) Then ASSERT_FAILURE "(init->style & ~(2 | 4 | 8 | 0x10 | 0x20 | WIDG_HIDDEN)) == 0", "lib/widget/form.cpp:49", "W_FORM", "Unknown style bit"
    If Not ((init.style And WFORM_INVISIBLE) = 0 Or (init.style And WFORM_CLICKABLE) = 0) Then ASSERT_FAILURE "(init->style & 2) == 0 || (init->style & 4) == 0", "lib/widget/form.cpp:50", "W_FORM", "Cannot have an invisible clickable form"
    If Not ((init.style And WFORM_CLICKABLE) <> 0 Or (init.style And (WFORM_NOPRIMARY Or WFORM_SECONDARY)) = 0) Then ASSERT_FAILURE "(init->style & 4) != 0 || (init->style & (0x10 | 0x20)) == 0", "lib/widget/form.cpp:51", "W_FORM", "Cannot set keys if the form isn't clickable"
End Sub

' W_FORM::W_FORM()
Public Sub W_FORM_construct(ByVal self As WIDGET)
    WIDGET_construct self, WIDG_FORM
    self.kind = K_W_FORM ' (VB) virtual calls from here on reach W_FORM's overrides
    W_FORM_initMembers self
    W_FORM_(self).disableChildren = False
End Sub

' the members' initialisers (form.h)
Private Sub W_FORM_initMembers(ByVal self As WIDGET)
    Dim d As W_FORM
    Set d = W_FORM_(self)
    d.userMovable = False
    d.minimizedLeftButtonWidth = 0
    d.formState = FormState_NORMAL
    d.has_dragStart = False
    d.minimizable = False
End Sub

' std::make_shared<W_FORM>(init)
Public Function W_FORM_new_init(init As W_FORMINIT) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_FORM)
    W_FORM_construct_init w, init
    Set W_FORM_new_init = w
End Function

' std::make_shared<W_FORM>()
Public Function W_FORM_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_FORM)
    W_FORM_construct w
    Set W_FORM_new = w
End Function

Public Function W_FORM_getFormState(ByVal self As WIDGET) As W_FORM_FormState
    W_FORM_getFormState = W_FORM_(self).formState
End Function

Private Function W_FORM_isUserMovable(ByVal self As WIDGET) As Boolean
    W_FORM_isUserMovable = W_FORM_(self).userMovable Or W_FORM_(self).formState = FormState_MINIMIZED
End Function

Public Sub W_FORM_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_FORM, g As WzRect
    Set d = W_FORM_(self)
    self.dirty = True
    If W_FORM_isUserMovable(self) And key = WKEY_PRIMARY Then
        g = W_FORM_minimizedGeometry(self)
        If d.formState = FormState_MINIMIZED And (psContext.mx <= g.x + d.minimizedLeftButtonWidth) Then
            Exit Sub
        End If
        d.dragStart = Vector2i_(psContext.mx, psContext.my)
        d.has_dragStart = True
    End If
End Sub

Public Sub W_FORM_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_FORM, g As WzRect
    Set d = W_FORM_(self)
    g = W_FORM_minimizedGeometry(self)
    If d.formState = FormState_MINIMIZED _
        And Not d.has_dragStart _
        And (psContext.mx <= g.x + d.minimizedLeftButtonWidth) Then
        ' mousedown was on "re-open" button (drag is only started when mouse is over the rest of the minimized frame)
        ' mouse is still over "re-open" button
        ' clicked + released mouse-button on "re-open" button
        W_FORM_toggleMinimized self
        Exit Sub
    End If
    If Not W_FORM_isUserMovable(self) Or Not d.has_dragStart Then Exit Sub
    d.has_dragStart = False
    self.dirty = True
End Sub

Public Sub W_FORM_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    ' currently, no-op
End Sub

Public Function W_FORM_capturesMouseDrag(ByVal self As WIDGET, ByVal key As WIDGET_KEY) As Boolean
    W_FORM_capturesMouseDrag = W_FORM_isUserMovable(self) And (key = WKEY_PRIMARY) And W_FORM_(self).has_dragStart
End Function

Public Sub W_FORM_mouseDragged(ByVal self As WIDGET, ByVal key As WIDGET_KEY, start As W_CONTEXT, current As W_CONTEXT)
    Dim d As W_FORM, currentMousePos As Vector2i, dragDelta As Vector2i, pRectToMove As WzRect, newPosition As Vector2i
    Dim maxPossibleX As Long, maxPossibleY As Long, ds As Vector2i, mr As WzRect
    If key <> WKEY_PRIMARY Then
        Exit Sub
    End If

    Set d = W_FORM_(self)
    If Not d.userMovable Or Not d.has_dragStart Then Exit Sub

    currentMousePos = Vector2i_(current.mx, current.my)
    ds = d.dragStart
    If currentMousePos.x = ds.x And currentMousePos.y = ds.y Then Exit Sub
    dragDelta = Vector2i_(currentMousePos.x - ds.x, currentMousePos.y - ds.y)

    pRectToMove = self.geometry()
    If d.formState = FormState_MINIMIZED Then
        pRectToMove = W_FORM_minimizedGeometry(self)
    End If
    newPosition = Vector2i_(pRectToMove.x + dragDelta.x, pRectToMove.y + dragDelta.y)

    maxPossibleX = pie_GetVideoBufferWidth() - pRectToMove.w
    maxPossibleY = pie_GetVideoBufferHeight() - pRectToMove.h
    newPosition.x = clip_(newPosition.x, 0, maxPossibleX)
    newPosition.y = clip_(newPosition.y, 0, maxPossibleY)
    If d.formState <> FormState_MINIMIZED Then
        self.move newPosition.x, newPosition.y
    Else
        mr = d.minimizedRect
        d.minimizedRect = WzRect_(newPosition.x, newPosition.y, mr.w, mr.h)
    End If
    self.dirty = True
    d.dragStart = currentMousePos
End Sub

' Respond to the mouse moving off a form
Public Sub W_FORM_highlightLost(ByVal self As WIDGET)
    self.dirty = True
End Sub

Public Sub W_FORM_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long
    If (self.style And WFORM_INVISIBLE) = 0 Then
        x0 = self.x() + xOffset
        y0 = self.y() + yOffset
        x1 = x0 + self.width()
        y1 = y0 + self.height()

        iV_ShadowBox x0, y0, x1, y1, 1, psPalette(WZCOL_FORM_LIGHT), psPalette(WZCOL_FORM_DARK), psPalette(WZCOL_FORM_BACKGROUND)
    End If
End Sub

Public Sub W_FORM_screenSizeDidChange(ByVal self As WIDGET, ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim d As W_FORM, newMinX0 As Long, newMinY0 As Long, mr As WzRect
    Set d = W_FORM_(self)
    If d.formState = FormState_MINIMIZED Then
        ' must ensure that the minimizedRect is fully within the new screen bounds
        mr = d.minimizedRect
        newMinX0 = MAX_(0, MIN_(mr.x, newWidth - mr.w))
        newMinY0 = MAX_(0, MIN_(mr.y, newHeight - mr.h))
        d.minimizedRect = WzRect_(newMinX0, newMinY0, mr.w, mr.h)
    End If
    WIDGET_screenSizeDidChange self, oldWidth, oldHeight, newWidth, newHeight
End Sub

Public Function W_FORM_hitTest(ByVal self As WIDGET, ByVal x As Long, ByVal y As Long) As Boolean
    Dim d As W_FORM, g As WzRect
    Set d = W_FORM_(self)
    If Not d.minimizable Or d.formState <> FormState_MINIMIZED Then
        W_FORM_hitTest = WIDGET_hitTest(self, x, y)
        Exit Function
    End If
    ' handle: minimized
    g = W_FORM_minimizedGeometry(self)
    W_FORM_hitTest = WzRect_contains(g, x, y)
End Function

Public Function W_FORM_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim d As W_FORM
    Set d = W_FORM_(self)
    If (Not d.minimizable Or d.formState <> FormState_MINIMIZED) And Not d.disableChildren Then
        Set W_FORM_findMouseTargetRecursive = WIDGET_findMouseTargetRecursive(self, psContext, key, wasPressed)
        Exit Function
    End If

    If Not self.visible() Then
        Exit Function
    End If

    If self.transparentToMouse() Then
        Exit Function
    End If

    Set W_FORM_findMouseTargetRecursive = self
End Function

Public Sub W_FORM_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim d As W_FORM
    Set d = W_FORM_(self)
    If Not d.minimizable Or d.formState <> FormState_MINIMIZED Then
        WIDGET_displayRecursive self, context
        Exit Sub
    End If

    ' handle: minimized
    W_FORM_displayMinimized self, WidgetGraphicsContext_getXOffset(context), WidgetGraphicsContext_getYOffset(context)
End Sub

Private Sub W_FORM_displayMinimized(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_FORM, minDim As WzRect, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Dim dropShadowX0 As Long, dropShadowY0 As Long, dropShadowX1 As Long, dropShadowY1 As Long, titleY0 As Long
    Set d = W_FORM_(self)
    minDim = W_FORM_minimizedGeometry(self)

    x0 = minDim.x + xOffset
    y0 = minDim.y + yOffset
    x1 = x0 + minDim.w
    y1 = y0 + minDim.h

    ' draw "drop-shadow"
    dropShadowX0 = MAX_(x0 - 5, 0)
    dropShadowY0 = MAX_(y0 - 5, 0)
    dropShadowX1 = MIN_(x1 + 5, pie_GetVideoBufferWidth())
    dropShadowY1 = MIN_(y1 + 5, pie_GetVideoBufferHeight())
    pie_UniTransBoxFill CSng(dropShadowX0), CSng(dropShadowY0), CSng(dropShadowX1), CSng(dropShadowY1), pal_RGBA(0, 0, 0, 40)

    ' draw background + border
    iV_TransBoxFill CSng(x0), CSng(y0), CSng(x1), CSng(y1)
    iV_Box x0, y0, x1, y1, pal_RGBA(255, 255, 255, 150)

    ' draw minimized title
    titleY0 = y0 + (minDim.h - d.minimizedTitle.height()) \ 2
    d.minimizedTitle.display x0 + WZ_FORM_MINIMIZED_LEFT_PADDING, titleY0

    If self.isMouseOverWidget() And (mouseX() <= x0 + d.minimizedLeftButtonWidth) Then
        ' Draw "re-open" button border when mouse is over
        iV_Box x0 + 1, y0 + 1, x0 + d.minimizedLeftButtonWidth, y1 - 1, pal_RGBA(255, 255, 255, 150)
    End If
End Sub

' void enableMinimizing(const std::string& minimizedTitle, PIELIGHT titleColour = WZCOL_FORM_LIGHT)
Public Sub W_FORM_enableMinimizing(ByVal self As WIDGET, ByVal minimizedTitle_ As String, Optional ByVal titleColour As Variant)
    Dim d As W_FORM, titleString As String, col As Long
    Set d = W_FORM_(self)
    If IsMissing(titleColour) Then col = psPalette(WZCOL_FORM_LIGHT) Else col = CLng(titleColour)
    Set d.minimizedTitle = W_LABEL_new()
    W_LABEL_setFont d.minimizedTitle, font_regular_bold, col
    titleString = ChrW$(&H21F2) & "  " ' (U+21F2: south east arrow to corner)
    If Len(minimizedTitle_) <> 0 Then
        titleString = titleString & minimizedTitle_
    Else
        titleString = titleString & "(untitled)"
    End If
    d.minimizedLeftButtonWidth = WZ_FORM_MINIMIZED_LEFT_PADDING + iV_GetTextWidth(ChrW$(&H21F2), font_regular_bold) + WZ_FORM_MINIMIZED_MAXBUTTON_RIGHT_PADDING
    d.minimizedTitle.setString titleString
    d.minimizedTitle.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.minimizedTitle), iV_GetTextLineSize(font_regular_bold)
    d.minimizable = True
End Sub

Public Sub W_FORM_disableMinimizing(ByVal self As WIDGET)
    Dim d As W_FORM
    Set d = W_FORM_(self)
    Set d.minimizedTitle = Nothing
    d.minimizable = False
    If d.formState = FormState_MINIMIZED Then
        W_FORM_setFormState self, FormState_NORMAL
    End If
End Sub

Public Function W_FORM_setFormState(ByVal self As WIDGET, ByVal state As W_FORM_FormState) As Boolean
    Dim d As W_FORM, minimizedSize As Vector2i
    Set d = W_FORM_(self)
    If d.formState = state Then W_FORM_setFormState = True: Exit Function
    If d.formState = FormState_MINIMIZED Then
        d.disableChildren = False
    End If
    Select Case state
    Case FormState_NORMAL
    Case FormState_MINIMIZED
        If Not d.minimizable Then
            W_FORM_setFormState = False
            Exit Function
        End If
        minimizedSize = W_FORM_calcMinimizedSize(self)
        d.minimizedRect = WzRect_(self.x(), self.y(), minimizedSize.x, minimizedSize.y)
        d.disableChildren = True
    End Select
    d.formState = state
    W_FORM_setFormState = True
End Function

Public Function W_FORM_toggleMinimized(ByVal self As WIDGET) As Boolean
    Dim desiredState As W_FORM_FormState
    desiredState = FormState_MINIMIZED
    If W_FORM_(self).formState = FormState_MINIMIZED Then
        desiredState = FormState_NORMAL
    End If
    W_FORM_toggleMinimized = W_FORM_setFormState(self, desiredState)
End Function

Private Function W_FORM_calcMinimizedSize(ByVal self As WIDGET) As Vector2i
    Dim d As W_FORM, minimizedSize As Vector2i
    Set d = W_FORM_(self)
    minimizedSize = Vector2i_(0, 0)
    If Not d.minimizable Then W_FORM_calcMinimizedSize = minimizedSize: Exit Function
    ' width = width of the minimizedTitle (which includes the "re-open button" as a unicode character) + padding
    If Not d.minimizedTitle Is Nothing Then
        minimizedSize.x = W_LABEL_getMaxLineWidth(d.minimizedTitle) + 15
    End If
    ' height = height of the font's line height + padding
    minimizedSize.y = iV_GetTextLineSize(font_regular_bold) + 6
    W_FORM_calcMinimizedSize = minimizedSize
End Function

Private Function W_FORM_minimizedGeometry(ByVal self As WIDGET) As WzRect
    W_FORM_minimizedGeometry = W_FORM_(self).minimizedRect
End Function

'--------------------------------------------------------------------------
' W_CLICKFORM

' W_CLICKFORM::W_CLICKFORM(W_FORMINIT const *init)
Public Sub W_CLICKFORM_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    Dim d As W_CLICKFORM
    W_FORM_construct_init self, init
    self.kind = K_W_CLICKFORM ' (VB) virtual calls from here on reach W_CLICKFORM's overrides
    Set d = W_CLICKFORM_(self)
    d.state = WBUT_PLAIN
    d.pTip = init.pTip
    d.HilightAudioID = WidgGetHilightAudioID()
    d.ClickedAudioID = WidgGetClickedAudioID()
    Set d.AudioCallback = WidgGetAudioCallback()
End Sub

' W_CLICKFORM::W_CLICKFORM()
Public Sub W_CLICKFORM_construct(ByVal self As WIDGET)
    Dim d As W_CLICKFORM
    W_FORM_construct self
    self.kind = K_W_CLICKFORM ' (VB) virtual calls from here on reach W_CLICKFORM's overrides
    Set d = W_CLICKFORM_(self)
    d.state = WBUT_PLAIN
    d.HilightAudioID = WidgGetHilightAudioID()
    d.ClickedAudioID = WidgGetClickedAudioID()
    Set d.AudioCallback = WidgGetAudioCallback()
End Sub

Public Function W_CLICKFORM_new_init(init As W_FORMINIT) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_CLICKFORM)
    W_CLICKFORM_construct_init w, init
    Set W_CLICKFORM_new_init = w
End Function

Public Function W_CLICKFORM_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_CLICKFORM)
    W_CLICKFORM_construct w
    Set W_CLICKFORM_new = w
End Function

Public Function W_CLICKFORM_getState(ByVal self As WIDGET) As Long
    W_CLICKFORM_getState = W_CLICKFORM_(self).state And (WBUT_DISABLE Or WBUT_LOCK Or WBUT_CLICKLOCK Or WBUT_FLASH Or WBUT_DOWN Or WBUT_HIGHLIGHT)
End Function

Public Sub W_CLICKFORM_setState(ByVal self As WIDGET, ByVal state As Long)
    Dim d As W_CLICKFORM, mask As Long
    If Not (Not ((state And WBUT_LOCK) <> 0 And (state And WBUT_CLICKLOCK) <> 0)) Then ASSERT_FAILURE "!((newState & WBUT_LOCK) && (newState & WBUT_CLICKLOCK))", "lib/widget/form.cpp:83", "setState", "Cannot have both WBUT_LOCK and WBUT_CLICKLOCK"

    Set d = W_CLICKFORM_(self)
    mask = WBUT_DISABLE Or WBUT_LOCK Or WBUT_CLICKLOCK
    d.state = (d.state And Not mask) Or (state And mask)
    self.dirty = True
End Sub

Public Sub W_CLICKFORM_setFlash(ByVal self As WIDGET, ByVal enable As Boolean)
    Dim d As W_CLICKFORM
    Set d = W_CLICKFORM_(self)
    If enable Then
        d.state = d.state Or WBUT_FLASH
    Else
        d.state = d.state And Not WBUT_FLASH
    End If
    self.dirty = True
End Sub

Public Sub W_CLICKFORM_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_CLICKFORM, now_ As Double
    W_FORM_run self, psContext

    Set d = W_CLICKFORM_(self)
    If d.has_clickDownStart Then
        now_ = widgSteadyNowMs()
        If now_ - d.clickDownStart >= widgGetClickHoldMS() Then
            If d.has_clickDownKey Then
                If vt_W_CLICKFORM_clickHeld(self, psContext, d.clickDownKey) Then
                    ' clear button down state, as the clickHeld event "consumed" this click
                    d.state = d.state And Not WBUT_DOWN
                End If
            End If
            d.has_clickDownStart = False
        End If
    End If
End Sub

' Returns true if "consumed" held click
Public Function W_CLICKFORM_clickHeld(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY) As Boolean
    W_CLICKFORM_clickHeld = False
End Function

' Respond to a mouse click
Public Sub W_CLICKFORM_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_CLICKFORM
    W_FORM_clicked self, psContext, key

    Set d = W_CLICKFORM_(self)
    ' Can't click a button if it is disabled or locked down.
    If (d.state And (WBUT_DISABLE Or WBUT_LOCK)) = 0 Then
        ' Check this is the correct key
        If ((self.style And WFORM_NOPRIMARY) = 0 And key = WKEY_PRIMARY) Or _
           ((self.style And WFORM_SECONDARY) <> 0 And key = WKEY_SECONDARY) Then
            d.state = d.state And Not WBUT_FLASH  ' Stop it flashing
            d.state = d.state Or WBUT_DOWN
            d.clickDownStart = widgSteadyNowMs()
            d.has_clickDownStart = True
            d.clickDownKey = key
            d.has_clickDownKey = True
            self.dirty = True

            If Not d.AudioCallback Is Nothing Then
                WzFunc_call d.AudioCallback, d.ClickedAudioID
            End If
        End If
    End If
End Sub

Public Sub W_CLICKFORM_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_CLICKFORM, lockedScreen As W_SCREEN
    Set d = W_CLICKFORM_(self)
    If (d.state And WBUT_DOWN) <> 0 Then
        ' Check this is the correct key.
        If ((self.style And WFORM_NOPRIMARY) = 0 And key = WKEY_PRIMARY) Or _
           ((self.style And WFORM_SECONDARY) <> 0 And key = WKEY_SECONDARY) Then
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

' Respond to a mouse moving over a form
Public Sub W_CLICKFORM_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_CLICKFORM
    Set d = W_CLICKFORM_(self)
    d.state = d.state Or WBUT_HIGHLIGHT

    If Not d.AudioCallback Is Nothing Then
        WzFunc_call d.AudioCallback, d.HilightAudioID
    End If
End Sub

Public Sub W_CLICKFORM_highlightLost(ByVal self As WIDGET)
    Dim d As W_CLICKFORM
    W_FORM_highlightLost self

    Set d = W_CLICKFORM_(self)
    d.state = d.state And Not (WBUT_DOWN Or WBUT_HIGHLIGHT)
    d.has_clickDownStart = False
    d.has_clickDownKey = False
    self.dirty = True
End Sub

Public Function W_CLICKFORM_getTip(ByVal self As WIDGET) As String
    W_CLICKFORM_getTip = W_CLICKFORM_(self).pTip
End Function

' WidgetHelp const * getHelp() const (Nothing if !help.has_value())
Public Function W_CLICKFORM_getHelp(ByVal self As WIDGET) As WidgetHelp
    Set W_CLICKFORM_getHelp = W_CLICKFORM_(self).help
End Function

Public Sub W_CLICKFORM_setTip(ByVal self As WIDGET, ByVal string_ As String)
    W_CLICKFORM_(self).pTip = string_
End Sub

Public Sub W_CLICKFORM_setHelp(ByVal self As WIDGET, ByVal help As WidgetHelp)
    Set W_CLICKFORM_(self).help = help
End Sub

Public Function W_CLICKFORM_isDown(ByVal self As WIDGET) As Boolean
    W_CLICKFORM_isDown = (W_CLICKFORM_(self).state And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
End Function

Public Function W_CLICKFORM_isHighlighted(ByVal self As WIDGET) As Boolean
    W_CLICKFORM_isHighlighted = (W_CLICKFORM_(self).state And WBUT_HIGHLIGHT) <> 0
End Function

Public Sub W_CLICKFORM_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ASSERT_FAILURE "false", "lib/widget/form.cpp:491", "display", "No default implementation exists for click forms"
End Sub

'--------------------------------------------------------------------------
' MARK: - W_FULLSCREENOVERLAY_CLICKFORM

' W_FULLSCREENOVERLAY_CLICKFORM(W_FORMINIT const *init) : W_CLICKFORM(init) {} (protected)
Public Sub W_FULLSCREENOVERLAY_CLICKFORM_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    W_CLICKFORM_construct_init self, init
    self.kind = K_W_FULLSCREENOVERLAY_CLICKFORM ' (VB) virtual calls from here on reach W_FULLSCREENOVERLAY_CLICKFORM's overrides
    W_FULLSCREENOVERLAY_CLICKFORM_initMembers self
End Sub

' W_FULLSCREENOVERLAY_CLICKFORM() : W_CLICKFORM() {} (protected)
Public Sub W_FULLSCREENOVERLAY_CLICKFORM_construct(ByVal self As WIDGET)
    W_CLICKFORM_construct self
    self.kind = K_W_FULLSCREENOVERLAY_CLICKFORM ' (VB) virtual calls from here on reach W_FULLSCREENOVERLAY_CLICKFORM's overrides
    W_FULLSCREENOVERLAY_CLICKFORM_initMembers self
End Sub

Private Sub W_FULLSCREENOVERLAY_CLICKFORM_initMembers(ByVal self As WIDGET)
    W_FULLSCREENOVERLAY_CLICKFORM_(self).backgroundColor = pal_RGBA(0, 0, 0, 125)
End Sub

' static std::shared_ptr<W_FULLSCREENOVERLAY_CLICKFORM> make(UDWORD formID = 0)
Public Function W_FULLSCREENOVERLAY_CLICKFORM_make(Optional ByVal formID As Long = 0) As WIDGET
    Dim sInit As W_FORMINIT, w As WIDGET
    sInit.id = formID
    sInit.style = WFORM_PLAIN Or WFORM_CLICKABLE
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1
    Set sInit.calcLayout = WzFunc_new(FN_W_FULLSCREENOVERLAY_CLICKFORM_calcLayout)

    ' (class make_shared_enabler)
    Set w = WIDGET_alloc(K_W_FULLSCREENOVERLAY_CLICKFORM)
    W_FULLSCREENOVERLAY_CLICKFORM_construct_init w, sInit
    Set W_FULLSCREENOVERLAY_CLICKFORM_make = w
End Function

Public Sub W_FULLSCREENOVERLAY_CLICKFORM_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_FULLSCREENOVERLAY_CLICKFORM
    Set d = W_FULLSCREENOVERLAY_CLICKFORM_(self)
    If Not d.onClickedFunc Is Nothing Then
        WzFunc_call d.onClickedFunc
    End If
End Sub

Public Sub W_FULLSCREENOVERLAY_CLICKFORM_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_FULLSCREENOVERLAY_CLICKFORM, x0 As Long, y0 As Long, strongCutoutWidget As WIDGET, screenRect As WzRect
    Dim cutoutRect As WzRect, other As WzRect, surroundingRects(0 To 3) As WzRect, i As Long, rect As WzRect
    Set d = W_FULLSCREENOVERLAY_CLICKFORM_(self)
    If Not self.visible() Then
        ' skip if hidden
        Exit Sub
    End If

    If (d.backgroundColor And &HFF000000) = 0 Then ' backgroundColor.isTransparent()
        Exit Sub
    End If

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    Set strongCutoutWidget = d.cutoutWidget
    screenRect = self.screenGeometry()
    If Not strongCutoutWidget Is Nothing Then
        other = strongCutoutWidget.screenGeometry()
        cutoutRect = WzRect_intersectionWith(screenRect, other)
    Else
        cutoutRect = WzRect_(0, 0, 0, 0)
    End If

    If cutoutRect.w <= 0 Or cutoutRect.h <= 0 Then
        ' simple path - draw background over everything
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), d.backgroundColor
    Else
        ' left column until left edge of cutout rect
        surroundingRects(0) = WzRect_corners(WzRect_left(screenRect), WzRect_top(screenRect), WzRect_left(cutoutRect), WzRect_bottom(screenRect))
        ' top band above cutout rect
        surroundingRects(1) = WzRect_corners(WzRect_left(cutoutRect), WzRect_top(screenRect), WzRect_right(cutoutRect), WzRect_top(cutoutRect))
        ' right column after cutout rect
        surroundingRects(2) = WzRect_corners(WzRect_right(cutoutRect), WzRect_top(screenRect), WzRect_right(screenRect), WzRect_bottom(screenRect))
        ' bottom band below cutout rect
        surroundingRects(3) = WzRect_corners(WzRect_left(cutoutRect), WzRect_bottom(cutoutRect), WzRect_right(cutoutRect), WzRect_bottom(screenRect))

        For i = 0 To 3
            rect = surroundingRects(i)
            If rect.w > 0 And rect.h > 0 Then
                pie_UniTransBoxFill WzRect_left(rect), WzRect_top(rect), WzRect_right(rect), WzRect_bottom(rect), d.backgroundColor
            End If
        Next
    End If
End Sub

Public Sub W_FULLSCREENOVERLAY_CLICKFORM_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_FULLSCREENOVERLAY_CLICKFORM
    Set d = W_FULLSCREENOVERLAY_CLICKFORM_(self)
    If keyPressed(KEY_ESC) Then
        If Not d.onCancelPressed Is Nothing Then
            WzFunc_call d.onCancelPressed
        End If
    End If
    inputLoseFocus   ' clear the input buffer.
End Sub

' void setCutoutWidget(std::shared_ptr<WIDGET> widget)
Public Sub W_FULLSCREENOVERLAY_CLICKFORM_setCutoutWidget(ByVal self As WIDGET, ByVal widget As WIDGET)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(self).cutoutWidget = widget
End Sub

'--------------------------------------------------------------------------

Public Sub displayChildDropShadows(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Const widerPadding As Long = 4
    Const closerPadding As Long = 2
    Dim dropShadowColor As Long, childXOffset As Long, childYOffset As Long, i As Long, child As WIDGET
    Dim childX0 As Long, childY0 As Long
    Dim childDropshadowWiderX0 As Long, childDropshadowWiderX1 As Long, childDropshadowWiderY1 As Long
    Dim childDropshadowCloserX0 As Long, childDropshadowCloserX1 As Long, childDropshadowCloserY1 As Long
    dropShadowColor = pal_RGBA(0, 0, 0, 40)
    childXOffset = psWidget.x() + xOffset
    childYOffset = psWidget.y() + yOffset
    For i = 0 To psWidget.childCount() - 1
        Set child = psWidget.childAt(i)
        If child.visible() Then
            childX0 = child.x() + childXOffset
            childY0 = child.y() + childYOffset
            childDropshadowWiderX0 = MAX_(childX0 - widerPadding, 0)
            childDropshadowWiderX1 = MIN_(childX0 + child.width() + widerPadding, pie_GetVideoBufferWidth())
            childDropshadowWiderY1 = MIN_(childY0 + child.height() + widerPadding, pie_GetVideoBufferHeight())
            childDropshadowCloserX0 = MAX_(childX0 - closerPadding, 0)
            childDropshadowCloserX1 = MIN_(childX0 + child.width() + closerPadding, pie_GetVideoBufferWidth())
            childDropshadowCloserY1 = MIN_(childY0 + child.height() + closerPadding, pie_GetVideoBufferHeight())
            pie_UniTransBoxFill childDropshadowWiderX0, childY0, childDropshadowWiderX1, childDropshadowWiderY1, dropShadowColor
            pie_UniTransBoxFill childDropshadowCloserX0, childY0, childDropshadowCloserX1, childDropshadowCloserY1, dropShadowColor
        End If
    Next
End Sub
