Attribute VB_Name = "ingameopscreen_c"
' src/screens/ingameopscreen.cpp, ingameopscreen.h: the in-game options screen, an overlay screen
' (struct WzInGameOptionsScreen : W_SCREEN) that shows a form (one of the options forms, the music
' manager: the caller's displayForm) under a title bar with a close button, dims the screen and
' pauses the game while it is up: showInGameOptionsScreen. Its widgets
' WzInGameOptionsScreen_CLICKFORM (the screen's root form), InGameOptionsTitleBarButton,
' InGameOptionsTitleBar and InGameOptionsBaseForm have their fields in the classes of the same names.
' (VB) W_SCREEN has no room for a subclass's fields: WzInGameOptionsScreen's are this module's
' (optionsScreenFields), as only one such screen is alive at a time (showInGameOptionsScreen closes
' the one before; it is previousInGameOptionsScreen), and its methods are
' WzInGameOptionsScreen_<method>(screen). The weak pointers to the screen and to its root form are
' their addresses, locked through previousInGameOptionsScreen; the title bar is locked from its
' close button (the button's parent).
' (VB) widget class WzInGameOptionsScreen_CLICKFORM : W_CLICKFORM
' (VB) widget class InGameOptionsTitleBarButton : W_BUTTON
' (VB) widget class InGameOptionsTitleBar : WIDGET
' (VB) widget class InGameOptionsBaseForm : IntFormAnimated
Option Explicit

' WzFunc ids (fork G6: 12100-12199 here)
Public Const FN_ingameopscreen_clickFormCalcLayout As Long = 12100  ' WzInGameOptionsScreen_CLICKFORM::make's LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_ingameopscreen_closeButtonClicked As Long = 12101   ' InGameOptionsTitleBar::initialize's [weakSelf](W_BUTTON&)
Public Const FN_ingameopscreen_baseFormOnClose As Long = 12102      ' WzInGameOptionsScreen_CLICKFORM::initialize's [weakSelf]()
Public Const FN_ingameopscreen_closeButtonTask As Long = 12103      ' its widgScheduleTask([weakSelf](){ ... })
Public Const FN_ingameopscreen_onClickedFunc As Long = 12104        ' WzInGameOptionsScreen::make's onClickedFunc [psWeakOverlayScreen]()
Public Const FN_ingameopscreen_onCancelPressed As Long = 12105      ' WzInGameOptionsScreen::make's onCancelPressed [psWeakOverlayScreen]()

' WzInGameOptionsScreen_CLICKFORM::INTERNAL_PADDING (static constexpr int32_t)
Private Const INTERNAL_PADDING As Long = 15

Private Const GuideNavBarButtonHorizontalPadding As Long = 5   ' constexpr int

' struct WzInGameOptionsScreen's own fields (W_SCREEN's are W_SCREEN.cls's)
Private Type WzInGameOptionsScreen_fields
    onCloseFunc As WzFunc               ' OnCloseFunc (std::function<void ()>)
    openTime As Double                  ' std::chrono::steady_clock::time_point (widgSteadyNowMs)
    didStopGameTime As Boolean          ' = false
    didSetPauseStates As Boolean        ' = false
End Type

' static std::weak_ptr<WzInGameOptionsScreen> (VB: held here; it has expired once it is no longer
' a registered overlay screen, as nothing else holds it then: lockPreviousInGameOptionsScreen)
Private previousInGameOptionsScreen As W_SCREEN

' (VB) the fields of the WzInGameOptionsScreen that is alive (previousInGameOptionsScreen)
Private optionsScreenFields As WzInGameOptionsScreen_fields

' (VB) a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' (VB) previousInGameOptionsScreen.lock()
Private Function lockPreviousInGameOptionsScreen() As W_SCREEN
    If previousInGameOptionsScreen Is Nothing Then Exit Function
    If Not isRegisteredOverlayScreen(previousInGameOptionsScreen) Then
        Set previousInGameOptionsScreen = Nothing
        Exit Function
    End If
    Set lockPreviousInGameOptionsScreen = previousInGameOptionsScreen
End Function

' (VB) std::weak_ptr<WzInGameOptionsScreen>::lock() of the screen at address p
Private Function lockInGameOptionsScreen(ByVal p As Long) As W_SCREEN
    Dim strongCurrent As W_SCREEN
    Set strongCurrent = lockPreviousInGameOptionsScreen()
    If strongCurrent Is Nothing Then Exit Function
    If ObjPtr(strongCurrent) <> p Then Exit Function
    Set lockInGameOptionsScreen = strongCurrent
End Function

' (VB) std::weak_ptr<WzInGameOptionsScreen_CLICKFORM>::lock() of the form at address p: the root
' form of the screen that is alive
Private Function lockInGameOptionsClickForm(ByVal p As Long) As WIDGET
    Dim strongCurrent As W_SCREEN
    Set strongCurrent = lockPreviousInGameOptionsScreen()
    If strongCurrent Is Nothing Then Exit Function
    If ObjPtr(strongCurrent.psForm) <> p Then Exit Function
    Set lockInGameOptionsClickForm = strongCurrent.psForm
End Function

'--------------------------------------------------------------------------
' (VB) this module's lambdas (WzFunc ids FN_ above): args are the C++ call's arguments (calcLayout:
' the widget; a W_BUTTON onClick: the button; a scheduled task and the std::function<void ()>: none)
Public Function ingameopscreen_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, button As WIDGET, strongSelf As WIDGET, psOverlayScreen As W_SCREEN
    Select Case f.id
    Case FN_ingameopscreen_clickFormCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight
    Case FN_ingameopscreen_closeButtonClicked
        ' (VB: weakSelf, the title bar, is the button's parent while both live)
        Set button = args(0)
        Set strongSelf = button.parent()
        If Not strongSelf Is Nothing Then
            If Not WIDGET_isKindOf(strongSelf, K_InGameOptionsTitleBar) Then Set strongSelf = Nothing
        End If
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/screens/ingameopscreen.cpp:326", "initialize", "No parent?": Exit Function
        InGameOptionsTitleBar_triggerCloseHandler strongSelf
    Case FN_ingameopscreen_baseFormOnClose
        widgScheduleTask WzFunc_new(FN_ingameopscreen_closeButtonTask, f.cap(0))
    Case FN_ingameopscreen_closeButtonTask
        Set strongSelf = lockInGameOptionsClickForm(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/screens/ingameopscreen.cpp:422", "initialize", "No parent?": Exit Function
        WzInGameOptionsScreen_CLICKFORM_handleCloseButtonPressed strongSelf
    Case FN_ingameopscreen_onClickedFunc
        Set psOverlayScreen = lockInGameOptionsScreen(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            WzInGameOptionsScreen_closeScreen psOverlayScreen, False
        End If
    Case FN_ingameopscreen_onCancelPressed
        Set psOverlayScreen = lockInGameOptionsScreen(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            WzInGameOptionsScreen_closeScreen psOverlayScreen, True
        End If
    End Select
End Function

'--------------------------------------------------------------------------
' MARK: - WzInGameOptionsScreen implementation

' WzInGameOptionsScreen_CLICKFORM(W_FORMINIT const *init) : W_CLICKFORM(init) {} (protected)
Public Sub WzInGameOptionsScreen_CLICKFORM_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    W_CLICKFORM_construct_init self, init
    self.kind = K_WzInGameOptionsScreen_CLICKFORM
    WzInGameOptionsScreen_CLICKFORM_initMembers self
End Sub

' WzInGameOptionsScreen_CLICKFORM() : W_CLICKFORM() {} (protected)
Public Sub WzInGameOptionsScreen_CLICKFORM_construct(ByVal self As WIDGET)
    W_CLICKFORM_construct self
    self.kind = K_WzInGameOptionsScreen_CLICKFORM
    WzInGameOptionsScreen_CLICKFORM_initMembers self
End Sub

' the members' initialisers (~WzInGameOptionsScreen_CLICKFORM() { } has nothing to do)
Private Sub WzInGameOptionsScreen_CLICKFORM_initMembers(ByVal self As WIDGET)
    Dim d As WzInGameOptionsScreen_CLICKFORM
    Set d = WzInGameOptionsScreen_CLICKFORM_(self)
    d.backgroundColor = pal_RGBA(0, 0, 0, 80)
    d.edgeMargin = 10
End Sub

' static std::shared_ptr<WzInGameOptionsScreen_CLICKFORM> make(const std::shared_ptr<WIDGET>& displayForm, UDWORD formID = 0)
Public Function WzInGameOptionsScreen_CLICKFORM_make(ByVal displayForm As WIDGET, Optional ByVal formID As Long = 0) As WIDGET
    Dim sInit As W_FORMINIT, widget As WIDGET
    W_FORMINIT_construct sInit
    sInit.id = formID
    sInit.style = WFORM_PLAIN Or WFORM_CLICKABLE
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1
    Set sInit.calcLayout = WzFunc_new(FN_ingameopscreen_clickFormCalcLayout)

    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzInGameOptionsScreen_CLICKFORM)
    WzInGameOptionsScreen_CLICKFORM_construct_init widget, sInit
    WzInGameOptionsScreen_CLICKFORM_initialize widget, displayForm
    Set WzInGameOptionsScreen_CLICKFORM_make = widget
End Function

'--------------------------------------------------------------------------
' MARK: - InGameOptionsTitleBarButton

' (VB) std::make_shared<InGameOptionsTitleBarButton>(): W_BUTTON() and the members' initialisers
Public Sub InGameOptionsTitleBarButton_construct(ByVal self As WIDGET)
    Dim d As InGameOptionsTitleBarButton
    W_BUTTON_construct self
    self.kind = K_InGameOptionsTitleBarButton
    Set d = InGameOptionsTitleBarButton_(self)
    d.has_intfacImageID = False         ' optional<UWORD> intfacImageID = nullopt
    d.imageDimensions = 16
    d.lastWidgetWidth = 0
    d.isTruncated = False
End Sub

Public Function InGameOptionsTitleBarButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_InGameOptionsTitleBarButton)
    InGameOptionsTitleBarButton_construct widget
    Set InGameOptionsTitleBarButton_new = widget
End Function

' static std::shared_ptr<InGameOptionsTitleBarButton> make(const WzString& buttonText, iV_fonts FontID = font_regular)
Public Function InGameOptionsTitleBarButton_make(ByVal buttonText As String, Optional ByVal FontID As iV_fonts = font_regular) As WIDGET
    Dim result As WIDGET
    Set result = InGameOptionsTitleBarButton_new()
    W_BUTTON_(result).FontID = FontID
    W_BUTTON_(result).pText = buttonText
    Set InGameOptionsTitleBarButton_make = result
End Function

' static std::shared_ptr<InGameOptionsTitleBarButton> make(UWORD intfacImageID, int imageDimensions = 16)
Public Function InGameOptionsTitleBarButton_make_image(ByVal intfacImageID As Long, Optional ByVal imageDimensions As Long = 16) As WIDGET
    Dim result As WIDGET, d As InGameOptionsTitleBarButton
    If Not assertValidImage(IntImages, intfacImageID) Then
        Exit Function                   ' nullptr
    End If
    Set result = InGameOptionsTitleBarButton_new()
    Set d = InGameOptionsTitleBarButton_(result)
    d.has_intfacImageID = True
    d.intfacImageID = intfacImageID
    d.imageDimensions = imageDimensions
    Set InGameOptionsTitleBarButton_make_image = result
End Function

Public Function InGameOptionsTitleBarButton_idealWidth(ByVal self As WIDGET) As Long
    Dim d As InGameOptionsTitleBarButton
    Set d = InGameOptionsTitleBarButton_(self)
    If d.has_intfacImageID Then
        InGameOptionsTitleBarButton_idealWidth = d.imageDimensions + (GuideNavBarButtonHorizontalPadding * 2)
        Exit Function
    End If
    InGameOptionsTitleBarButton_idealWidth = iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID) + (GuideNavBarButtonHorizontalPadding * 2)
End Function

' Display the text, shrunk to fit, with a faint border around it (that gets bright when highlighted)
Public Sub InGameOptionsTitleBarButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As InGameOptionsTitleBarButton, pParent As WIDGET, xPos As Long, yPos As Long, h As Long, w As Long
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, has_backColor As Boolean, backColor As Long
    Dim msgColour As Long, imgPosX0 As Long, imgPosY0 As Long, imgColor As Long, pText As String
    Dim fontID As Long, result As iV_fonts, availableButtonTextWidth As Long
    Dim textX0 As Long, textY0 As Long, maxTextDisplayableWidth As Long
    Set d = InGameOptionsTitleBarButton_(self)
    Set pParent = self.parent()
    If Not (Not pParent Is Nothing) Then ASSERT_FAILURE "pParent != nullptr", "src/screens/ingameopscreen.cpp:173", "display", "Keymap buttons should have a parent container!": Exit Sub

    xPos = xOffset + self.x()
    yPos = yOffset + self.y()
    h = self.height()
    w = self.width()

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (W_BUTTON_(self).state And WBUT_DISABLE) <> 0
    isHighlight = (W_BUTTON_(self).state And WBUT_HIGHLIGHT) <> 0

    ' Draw background (if needed)
    has_backColor = False               ' optional<PIELIGHT> backColor
    If Not isDisabled Then
        If isDown Then
            backColor = psPalette(WZCOL_MENU_SCORE_BUILT)
            has_backColor = True
        ElseIf isHighlight Then
            backColor = pal_RGBA(255, 255, 255, 50)
            has_backColor = True
        End If
        If has_backColor Then
            pie_UniTransBoxFill xPos, yPos, xPos + w, yPos + h, backColor
        End If
    End If

    ' Select label text and color
    If Not isDown Then
        If isHighlight Then msgColour = psPalette(WZCOL_TEXT_BRIGHT) Else msgColour = psPalette(WZCOL_TEXT_MEDIUM)
    Else
        msgColour = psPalette(WZCOL_FORM_TEXT)
    End If
    If isDisabled Then
        msgColour = psPalette(WZCOL_TEXT_MEDIUM)
        msgColour = pielightWithAlpha(msgColour, pielightAlpha(msgColour) \ 2)
    End If

    If d.has_intfacImageID Then
        ' Display the image centered
        imgPosX0 = xPos + (w - d.imageDimensions) \ 2
        imgPosY0 = yPos + (h - d.imageDimensions) \ 2
        imgColor = msgColour
        iV_DrawImageFileAnisotropicTint IntImages, d.intfacImageID, imgPosX0, imgPosY0, d.imageDimensions, d.imageDimensions, imgColor
        Exit Sub ' shortcut the rest
    End If

    pText = W_BUTTON_(self).pText
    fontID = d.wzMessageText.getFontID()
    If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzMessageText.getText() <> pText Then
        fontID = W_BUTTON_(self).FontID
    End If
    d.wzMessageText.setText pText, fontID

    availableButtonTextWidth = w - (GuideNavBarButtonHorizontalPadding * 2)
    If d.wzMessageText.width() > availableButtonTextWidth Then
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
            d.wzMessageText.setText pText, fontID
        Loop While d.wzMessageText.width() > availableButtonTextWidth
    End If
    d.lastWidgetWidth = self.width()

    ' std::max(xPos + GuideNavBarButtonHorizontalPadding, xPos + ((w - wzMessageText.width()) / 2))
    textX0 = xPos + ((w - d.wzMessageText.width()) \ 2)
    If xPos + GuideNavBarButtonHorizontalPadding > textX0 Then textX0 = xPos + GuideNavBarButtonHorizontalPadding
    textY0 = yPos + (h - d.wzMessageText.lineSize()) \ 2 - d.wzMessageText.aboveBase()

    maxTextDisplayableWidth = w - (GuideNavBarButtonHorizontalPadding * 2)
    d.isTruncated = maxTextDisplayableWidth < d.wzMessageText.width()
    If d.isTruncated Then
        maxTextDisplayableWidth = maxTextDisplayableWidth - (iV_GetEllipsisWidth(d.wzMessageText.getFontID()) + 2)
    End If
    d.wzMessageText.render textX0, textY0, msgColour, 0!, maxTextDisplayableWidth

    If d.isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzMessageText.getFontID(), textX0 + maxTextDisplayableWidth + 2, textY0, msgColour
    End If
End Sub

Public Function InGameOptionsTitleBarButton_getTip(ByVal self As WIDGET) As String
    If Not InGameOptionsTitleBarButton_(self).isTruncated Then
        InGameOptionsTitleBarButton_getTip = W_BUTTON_(self).pTip
        Exit Function
    End If

    InGameOptionsTitleBarButton_getTip = W_BUTTON_(self).pText
End Function

'--------------------------------------------------------------------------
' MARK: - InGameOptionsTitleBar

' InGameOptionsTitleBar(): WIDGET() {} (protected)
Public Sub InGameOptionsTitleBar_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_InGameOptionsTitleBar
End Sub

' static std::shared_ptr<InGameOptionsTitleBar> make(const OnCloseFunc& onCloseHandler)
Public Function InGameOptionsTitleBar_make(ByVal onCloseHandler As WzFunc) As WIDGET
    Dim result As WIDGET
    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_InGameOptionsTitleBar)
    InGameOptionsTitleBar_construct result
    InGameOptionsTitleBar_initialize result, onCloseHandler
    Set InGameOptionsTitleBar_make = result
End Function

Private Sub InGameOptionsTitleBar_initialize(ByVal self As WIDGET, ByVal onCloseHandler As WzFunc)
    Dim d As InGameOptionsTitleBar
    Set d = InGameOptionsTitleBar_(self)
    Set d.m_onCloseHandler = onCloseHandler

    ' (std::weak_ptr<InGameOptionsTitleBar> weakSelf: the close button's parent, see
    ' FN_ingameopscreen_closeButtonClicked)

    ' Close button
    Set d.m_closeButton = InGameOptionsTitleBarButton_make(ChrW$(&H2715), font_regular) ' U+2715 (multiplication x)
    d.m_closeButton.setTip "Close"
    self.attach d.m_closeButton
    W_BUTTON_addOnClickHandler d.m_closeButton, WzFunc_new(FN_ingameopscreen_closeButtonClicked)
End Sub

Public Sub InGameOptionsTitleBar_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, h As Long, w As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    h = self.height()
    w = self.width()
    iV_Line x0, y0 + h, x0 + w, y0 + h, pal_RGBA(0, 0, 0, 150)
End Sub

Public Sub InGameOptionsTitleBar_geometryChanged(ByVal self As WIDGET)
    InGameOptionsTitleBar_updateLayout self
End Sub

Private Sub InGameOptionsTitleBar_updateLayout(ByVal self As WIDGET)
    Dim d As InGameOptionsTitleBar, w As Long, h As Long, widgetHeight As Long, widgetY0 As Long, buttonWidth As Long
    Set d = InGameOptionsTitleBar_(self)
    w = self.width()
    h = self.height()

    widgetHeight = h - (GuideNavBarButtonHorizontalPadding * 2)
    widgetY0 = (h - widgetHeight) \ 2

    buttonWidth = d.m_closeButton.idealWidth()

    ' Right-align the close button
    d.m_closeButton.setGeometry w - (buttonWidth + GuideNavBarButtonHorizontalPadding), widgetY0, buttonWidth, widgetHeight
End Sub

Private Sub InGameOptionsTitleBar_triggerCloseHandler(ByVal self As WIDGET)
    Dim d As InGameOptionsTitleBar
    Set d = InGameOptionsTitleBar_(self)
    If Not (Not d.m_onCloseHandler Is Nothing) Then ASSERT_FAILURE "m_onCloseHandler != nullptr", "src/screens/ingameopscreen.cpp:361", "triggerCloseHandler", "Missing onCloseHandler": Exit Sub
    WzFunc_call d.m_onCloseHandler
End Sub

Public Function InGameOptionsTitleBar_idealHeight(ByVal self As WIDGET) As Long
    InGameOptionsTitleBar_idealHeight = 30
End Function

'--------------------------------------------------------------------------
' MARK: - InGameOptionsBaseForm

' InGameOptionsBaseForm() : IntFormAnimated(true) { } (protected), and the members' initialisers
Public Sub InGameOptionsBaseForm_construct(ByVal self As WIDGET)
    IntFormAnimated_construct self, True
    self.kind = K_InGameOptionsBaseForm
    InGameOptionsBaseForm_(self).internalFormPadding = 1
End Sub

Private Sub InGameOptionsBaseForm_initialize(ByVal self As WIDGET, ByVal displayForm_ As WIDGET, ByVal onCloseFunc As WzFunc)
    Dim d As InGameOptionsBaseForm
    Set d = InGameOptionsBaseForm_(self)
    Set d.titleBar = InGameOptionsTitleBar_make(onCloseFunc)
    self.attach d.titleBar

    Set d.displayForm = displayForm_
    self.attach d.displayForm
End Sub

' static std::shared_ptr<InGameOptionsBaseForm> make(const std::shared_ptr<WIDGET>& displayForm, const std::function<void ()>& onCloseFunc)
Public Function InGameOptionsBaseForm_make(ByVal displayForm As WIDGET, ByVal onCloseFunc As WzFunc) As WIDGET
    Dim result As WIDGET
    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_InGameOptionsBaseForm)
    InGameOptionsBaseForm_construct result
    InGameOptionsBaseForm_initialize result, displayForm, onCloseFunc
    Set InGameOptionsBaseForm_make = result
End Function

Public Sub InGameOptionsBaseForm_geometryChanged(ByVal self As WIDGET)
    Dim d As InGameOptionsBaseForm, w As Long, h As Long, displayFormY0 As Long, displayFormHeight As Long
    Set d = InGameOptionsBaseForm_(self)
    w = self.width()
    h = self.height()
    d.titleBar.setGeometry d.internalFormPadding, d.internalFormPadding, w - (d.internalFormPadding * 2), d.titleBar.idealHeight()
    displayFormY0 = d.titleBar.y() + d.titleBar.height()
    displayFormHeight = h - displayFormY0 - d.internalFormPadding
    d.displayForm.setGeometry d.internalFormPadding, displayFormY0, w - (d.internalFormPadding * 2), displayFormHeight
End Sub

Public Function InGameOptionsBaseForm_idealWidth(ByVal self As WIDGET) As Long
    Dim d As InGameOptionsBaseForm
    Set d = InGameOptionsBaseForm_(self)
    InGameOptionsBaseForm_idealWidth = d.displayForm.idealWidth() + (d.internalFormPadding * 2)
End Function

Public Function InGameOptionsBaseForm_idealHeight(ByVal self As WIDGET) As Long
    Dim d As InGameOptionsBaseForm
    Set d = InGameOptionsBaseForm_(self)
    InGameOptionsBaseForm_idealHeight = d.titleBar.idealHeight() + d.displayForm.idealHeight() + (d.internalFormPadding * 2)
End Function

'--------------------------------------------------------------------------
' MARK: - WzInGameOptionsScreen_CLICKFORM

Private Sub WzInGameOptionsScreen_CLICKFORM_initialize(ByVal self As WIDGET, ByVal displayForm_ As WIDGET)
    Dim d As WzInGameOptionsScreen_CLICKFORM
    Set d = WzInGameOptionsScreen_CLICKFORM_(self)
    ' (std::weak_ptr<WzInGameOptionsScreen_CLICKFORM> weakSelf: the form's address, see
    ' lockInGameOptionsClickForm)
    Set d.wrapperForm = InGameOptionsBaseForm_make(displayForm_, WzFunc_new(FN_ingameopscreen_baseFormOnClose, ObjPtr(self)))
    self.attach d.wrapperForm
    WzInGameOptionsScreen_CLICKFORM_recalcLayout self
End Sub

Public Sub WzInGameOptionsScreen_CLICKFORM_geometryChanged(ByVal self As WIDGET)
    ' (VB) not when WIDGET's constructor runs the calcLayout (C++ calls WIDGET's geometryChanged
    ' then, while widget_c's constructors do not set the kind yet): no wrapperForm before initialize
    If WzInGameOptionsScreen_CLICKFORM_(self).wrapperForm Is Nothing Then Exit Sub
    WzInGameOptionsScreen_CLICKFORM_recalcLayout self
End Sub

Public Sub WzInGameOptionsScreen_CLICKFORM_recalcLayout(ByVal self As WIDGET)
    Dim d As WzInGameOptionsScreen_CLICKFORM, w As Long, h As Long, maxFormWidth As Long, maxFormHeight As Long
    Dim displayFormWidth As Long, displayFormHeight As Long, x0 As Long, y0 As Long
    Set d = WzInGameOptionsScreen_CLICKFORM_(self)
    w = self.width()
    h = self.height()

    maxFormWidth = w - (d.edgeMargin * 2)
    maxFormHeight = h - (d.edgeMargin * 2)

    ' ensure displayForm is centered
    displayFormWidth = d.wrapperForm.idealWidth()       ' std::min(maxFormWidth, wrapperForm->idealWidth())
    If maxFormWidth < displayFormWidth Then displayFormWidth = maxFormWidth
    displayFormHeight = d.wrapperForm.idealHeight()     ' std::min(maxFormHeight, wrapperForm->idealHeight())
    If maxFormHeight < displayFormHeight Then displayFormHeight = maxFormHeight

    x0 = (w - displayFormWidth) \ 2
    y0 = (h - displayFormHeight) \ 2

    d.wrapperForm.setGeometry x0, y0, displayFormWidth, displayFormHeight
End Sub

Private Sub WzInGameOptionsScreen_CLICKFORM_clearFocusAndEditing(ByVal self As WIDGET)
    Dim lockedScreen As W_SCREEN
    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        lockedScreen.setFocus Nothing
    End If
End Sub

Private Sub WzInGameOptionsScreen_CLICKFORM_handleClickOnForm(ByVal self As WIDGET)
    Dim d As WzInGameOptionsScreen_CLICKFORM
    Set d = WzInGameOptionsScreen_CLICKFORM_(self)
    WzInGameOptionsScreen_CLICKFORM_clearFocusAndEditing self
    If Not d.onClickedFunc Is Nothing Then
        WzFunc_call d.onClickedFunc
    End If
End Sub

Private Sub WzInGameOptionsScreen_CLICKFORM_handleCloseButtonPressed(ByVal self As WIDGET)
    Dim d As WzInGameOptionsScreen_CLICKFORM
    Set d = WzInGameOptionsScreen_CLICKFORM_(self)
    WzInGameOptionsScreen_CLICKFORM_clearFocusAndEditing self
    If Not d.onCancelPressed Is Nothing Then
        WzFunc_call d.onCancelPressed
    End If
End Sub

Public Sub WzInGameOptionsScreen_CLICKFORM_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    WzInGameOptionsScreen_CLICKFORM_handleClickOnForm self
End Sub

Public Sub WzInGameOptionsScreen_CLICKFORM_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzInGameOptionsScreen_CLICKFORM, x0 As Long, y0 As Long
    Set d = WzInGameOptionsScreen_CLICKFORM_(self)
    If Not self.visible() Then
        Exit Sub ' skip if hidden
    End If

    If (d.backgroundColor And &HFF000000) = 0 Then ' backgroundColor.isTransparent()
        Exit Sub
    End If

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    ' draw background over everything
    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), d.backgroundColor
End Sub

Public Sub WzInGameOptionsScreen_CLICKFORM_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzInGameOptionsScreen_CLICKFORM
    Set d = WzInGameOptionsScreen_CLICKFORM_(self)
    If keyPressed(KEY_ESC) Then
        WzInGameOptionsScreen_CLICKFORM_clearFocusAndEditing self
        If Not d.onCancelPressed Is Nothing Then
            WzFunc_call d.onCancelPressed
            inputLoseFocus
        End If
    End If

    If self.visible() Then
        ' while this is displayed, clear the input buffer
        ' (ensuring that subsequent screens / the main screen do not get the input to handle)
        inputLoseFocus
    End If
End Sub

'--------------------------------------------------------------------------
' MARK: - WzInGameOptionsScreen
' (~WzInGameOptionsScreen()'s setFocus(nullptr): W_SCREEN's Class_Terminate resigns the focus)

' static std::shared_ptr<WzInGameOptionsScreen> make(const std::shared_ptr<WIDGET>& displayForm, const OnCloseFunc& onCloseFunction)
Private Function WzInGameOptionsScreen_make(ByVal displayForm As WIDGET, ByVal onCloseFunc_ As WzFunc) As W_SCREEN
    Dim newRootFrm As WIDGET, screen As W_SCREEN, blank As WzInGameOptionsScreen_fields
    ' (class make_shared_enabler)
    Set newRootFrm = WzInGameOptionsScreen_CLICKFORM_make(displayForm)
    Set screen = New W_SCREEN
    optionsScreenFields = blank         ' (VB) WzInGameOptionsScreen(): W_SCREEN() and its members' initialisers
    screen.initialize newRootFrm
    Set optionsScreenFields.onCloseFunc = onCloseFunc_
    ' (std::weak_ptr<WzInGameOptionsScreen> psWeakOverlayScreen: the screen's address, see lockInGameOptionsScreen)
    Set WzInGameOptionsScreen_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_ingameopscreen_onClickedFunc, ObjPtr(screen))
    Set WzInGameOptionsScreen_CLICKFORM_(newRootFrm).onCancelPressed = WzFunc_new(FN_ingameopscreen_onCancelPressed, ObjPtr(screen))

    WzInGameOptionsScreen_pauseGameIfNeeded screen
    optionsScreenFields.openTime = widgSteadyNowMs()

    Set WzInGameOptionsScreen_make = screen
End Function

' void reopeningScreen(const OnCloseFunc& onCloseFunction)
Private Sub WzInGameOptionsScreen_reopeningScreen(ByVal self As W_SCREEN, ByVal onCloseFunction As WzFunc)
    Set optionsScreenFields.onCloseFunc = onCloseFunction
    WzInGameOptionsScreen_pauseGameIfNeeded self
    optionsScreenFields.openTime = widgSteadyNowMs()
End Sub

' void closeScreen(bool force = false)
Private Sub WzInGameOptionsScreen_closeScreen(ByVal self As W_SCREEN, Optional ByVal force As Boolean = False)
    Dim onCloseFunc As WzFunc
    If Not force And (widgSteadyNowMs() - optionsScreenFields.openTime < 250#) Then
        ' ignore errant clicks that might otherwise close screen right after it opened
        Exit Sub
    End If
    widgRemoveOverlayScreen self
    ' (VB) the overlay list lets it go, so previousInGameOptionsScreen (a weak pointer in C++) does too
    If previousInGameOptionsScreen Is self Then Set previousInGameOptionsScreen = Nothing
    WzInGameOptionsScreen_unpauseGameIfNeeded self
    Set onCloseFunc = optionsScreenFields.onCloseFunc
    If Not onCloseFunc Is Nothing Then
        ' "consumes" the onCloseFunc ((VB) before the call: the call may open the next screen, whose
        ' fields these become)
        Set optionsScreenFields.onCloseFunc = Nothing
        WzFunc_call onCloseFunc
    End If
End Sub

Private Sub WzInGameOptionsScreen_pauseGameIfNeeded(ByVal self As W_SCREEN)
    If Not runningMultiplayer() And Not gameTimeIsStopped() Then
        ' Stop the clock / simulation
        gameTimeStop
        addConsoleMessage "PAUSED", CENTRE_JUSTIFY, SYSTEM_MESSAGE, False, MAX_CONSOLE_MESSAGE_DURATION

        self.psForm.show ' show the root form, which means it accepts all clicks on the screen and dims the background

        optionsScreenFields.didStopGameTime = True

        If Not gamePaused() Then
            setGamePauseStatus True
            setConsolePause True
            setScriptPause True
            setAudioPause True
            setScrollPause True
            optionsScreenFields.didSetPauseStates = True

            ' If cursor trapping is enabled allow the cursor to leave the window
            wzReleaseMouse
        End If
    Else
        self.psForm.hide ' hide the root form - it won't accept clicks (but its children will be displayed)
    End If
End Sub

Private Sub WzInGameOptionsScreen_unpauseGameIfNeeded(ByVal self As W_SCREEN)
    If optionsScreenFields.didStopGameTime Then
        If Not runningMultiplayer() And gameTimeIsStopped() Then
            clearActiveConsole
            ' Start the clock again
            gameTimeStart

            'put any widgets back on for the missions
            resetMissionWidgets
        End If
        optionsScreenFields.didStopGameTime = False
    End If

    If optionsScreenFields.didSetPauseStates Then
        setGamePauseStatus False
        setConsolePause False
        setScriptPause False
        setAudioPause False
        setScrollPause False
        optionsScreenFields.didSetPauseStates = False

        ' Re-enable cursor trapping if it is enabled
        If shouldTrapCursor() Then
            wzGrabMouse
        End If
    End If
End Sub

'--------------------------------------------------------------------------
' MARK: -

Public Sub showInGameOptionsScreen(ByVal onTopOfScreen As W_SCREEN, ByVal displayForm As WIDGET, ByVal onCloseFunc As WzFunc)
    Dim strongCurrent As W_SCREEN, newScreen As W_SCREEN
    Set strongCurrent = lockPreviousInGameOptionsScreen()
    If Not strongCurrent Is Nothing Then
        WzInGameOptionsScreen_closeScreen strongCurrent, True
        Set previousInGameOptionsScreen = Nothing   ' previousInGameOptionsScreen.reset()
    End If
    Set newScreen = WzInGameOptionsScreen_make(displayForm, onCloseFunc)
    widgRegisterOverlayScreenOnTopOfScreen newScreen, onTopOfScreen
    Set previousInGameOptionsScreen = newScreen
End Sub
