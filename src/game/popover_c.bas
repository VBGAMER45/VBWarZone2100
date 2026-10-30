Attribute VB_Name = "popover_c"
' lib/widget/popover.cpp, popover.h: the widget class PopoverWidget
' (docs/CONVENTIONS.md, "Widgets"; its fields are PopoverWidget.cls). The enum
' classes PopoverWidget::Style and ::Alignment are PopoverWidget_Style and
' PopoverWidget_Alignment. makePopover's Vector2i positionOffset is two Longs
' (positionOffsetX, positionOffsetY); an OnPopoverCloseHandlerFunc is a WzFunc
' without arguments.
' (VB) The weak pointers: std::weak_ptr<W_SCREEN> overlayScreen is a strong
' reference that lock() gives back only while it is a registered overlay screen
' (isRegisteredOverlayScreen, as long as C++'s overlay list keeps it alive), and
' drops otherwise. The lambdas' weak pointers: the popover is its serial
' (widgweak_vb), the overlay screen its address (the click handler belongs to the
' overlay's root form, which runs only while its screen lives), the parent a
' strong reference; closeImpl drops the calcLayout lambda that holds it.
' (VB) widget class PopoverWidget : WIDGET
Option Explicit

' enum class Style
Public Enum PopoverWidget_Style
    PopoverWidget_Style_Interactive
    PopoverWidget_Style_NonInteractive  ' akin to a tooltip (accepts no mouse-over or other events, expects that something else will call close() on it at some point)
End Enum

' enum class Alignment
Public Enum PopoverWidget_Alignment
    PopoverWidget_Alignment_LeftOfParent
    PopoverWidget_Alignment_RightOfParent
End Enum

Public Const FN_PopoverWidget_openTask As Long = 2070
Public Const FN_PopoverWidget_displayChildDropShadows As Long = 2071
Public Const FN_PopoverWidget_onClicked As Long = 2072
Public Const FN_PopoverWidget_noHitTest As Long = 2073
Public Const FN_PopoverWidget_calcLayout As Long = 2074
Public Const FN_PopoverWidget_closeTask As Long = 2075

Public Function popover_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim popoverWidget As WIDGET, psOverlayScreen As W_SCREEN, psParent As WIDGET
    Select Case f.id
    Case FN_PopoverWidget_openTask
        ' [popoverWidget, weakParent]()
        PopoverWidget_openTask f.cap(0), f.cap(1)
    Case FN_PopoverWidget_displayChildDropShadows
        ' newRootFrm->displayFunction = displayChildDropShadows
        displayChildDropShadows args(0), args(1), args(2)
    Case FN_PopoverWidget_onClicked
        ' [psWeakOverlayScreen, pWeakThis]()
        Set popoverWidget = WeakWidget_lock(f.cap(1))
        If Not popoverWidget Is Nothing Then
            PopoverWidget_close popoverWidget
        End If
        Set psOverlayScreen = W_SCREEN_fromPtr(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            widgRemoveOverlayScreen psOverlayScreen
        End If
    Case FN_PopoverWidget_noHitTest
        ' [](const WIDGET *psWidget, int x, int y) -> bool
        popover_c_fn = False    ' effectively: make this (and all of its children) unclickable
    Case FN_PopoverWidget_calcLayout
        ' [weakParent](WIDGET *psWidget)
        Set psParent = f.cap(0)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "lib/widget/popover.cpp:146", "open", "parent is null": Exit Function
        PopoverWidget_repositionOnScreenRelativeToParent args(0), psParent
    Case FN_PopoverWidget_closeTask
        ' [popoverWidget]()
        PopoverWidget_closeImpl f.cap(0)
    End Select
End Function

' PopoverWidget::PopoverWidget()
Public Sub PopoverWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_PopoverWidget ' (VB) virtual calls from here on reach PopoverWidget's overrides
    PopoverWidget_(self).weakSerial = WeakWidget_register(self)
End Sub

Public Sub PopoverWidget_destruct(ByVal self As WIDGET)
    If Not PopoverWidget_lockOverlayScreen(self) Is Nothing Then
        PopoverWidget_close self    ' (VB) never: the registered overlay screen holds this widget
    End If
    WeakWidget_unregister PopoverWidget_(self).weakSerial
End Sub

' (VB) overlayScreen.lock()
Private Function PopoverWidget_lockOverlayScreen(ByVal self As WIDGET) As W_SCREEN
    Dim d As PopoverWidget
    Set d = PopoverWidget_(self)
    If d.overlayScreen Is Nothing Then Exit Function
    If Not isRegisteredOverlayScreen(d.overlayScreen) Then
        Set d.overlayScreen = Nothing   ' expired
        Exit Function
    End If
    Set PopoverWidget_lockOverlayScreen = d.overlayScreen
End Function

' static std::shared_ptr<PopoverWidget> makePopover(const std::shared_ptr<WIDGET>& parent, std::shared_ptr<WIDGET> popoverContents, Style style = Style::NonInteractive, Alignment align = Alignment::LeftOfParent, Vector2i positionOffset = Vector2i(0, 0), const OnPopoverCloseHandlerFunc& onCloseHandler = nullptr)
Public Function PopoverWidget_makePopover(ByVal parent As WIDGET, ByVal popoverContents As WIDGET, Optional ByVal style As PopoverWidget_Style = PopoverWidget_Style_NonInteractive, Optional ByVal align As PopoverWidget_Alignment = PopoverWidget_Alignment_LeftOfParent, Optional ByVal positionOffsetX As Long = 0, Optional ByVal positionOffsetY As Long = 0, Optional ByVal onCloseHandler As WzFunc = Nothing) As WIDGET
    Dim result As WIDGET, d As PopoverWidget
    Set result = WIDGET_alloc(K_PopoverWidget)
    PopoverWidget_construct result
    Set d = PopoverWidget_(result)
    d.style = style
    d.align = align
    d.positionOffsetX = positionOffsetX
    d.positionOffsetY = positionOffsetY
    result.setGeometry 0, 0, popoverContents.width(), popoverContents.height()
    Set d.popoverContents = popoverContents
    Set d.onCloseHandler = onCloseHandler
    result.attach popoverContents
    PopoverWidget_open result, parent
    Set PopoverWidget_makePopover = result
End Function

Private Sub PopoverWidget_repositionOnScreenRelativeToParent(ByVal self As WIDGET, ByVal parent As WIDGET)
    Dim d As PopoverWidget, popOverX0 As Long, popOverY0 As Long, parentX1 As Long
    Set d = PopoverWidget_(self)
    popOverX0 = 0
    Select Case d.align
    Case PopoverWidget_Alignment_LeftOfParent
        ' (Ideally, with its left aligned with the left side of the "parent" widget, but ensure visibility on-screen)
        popOverX0 = parent.screenPosX()
        If ULT(screenWidth, popOverX0 + self.width()) Then
            popOverX0 = screenWidth - self.width()
        End If
    Case PopoverWidget_Alignment_RightOfParent
        ' (Ideally, with its right aligned with the right side of the "parent" widget, but ensure left side visibility on-screen)
        parentX1 = parent.screenPosX() + parent.width()
        popOverX0 = parentX1 - self.width()
        If popOverX0 < 0 Then
            popOverX0 = 0
        End If
    End Select
    ' (Ideally, with its top aligned with the bottom of the "parent" widget, but ensure full visibility on-screen)
    popOverY0 = parent.screenPosY() + parent.height()
    If ULT(screenHeight, popOverY0 + self.height()) Then
        popOverY0 = screenHeight - self.height()
    End If
    self.move popOverX0 + d.positionOffsetX, popOverY0 + d.positionOffsetY
End Sub

Public Sub PopoverWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As PopoverWidget
    Set d = PopoverWidget_(self)
    If Not d.popoverContents Is Nothing Then
        d.popoverContents.setGeometry 0, 0, self.width(), self.height()
    End If
End Sub

Private Sub PopoverWidget_open(ByVal self As WIDGET, ByVal parent As WIDGET)
    If Not PopoverWidget_lockOverlayScreen(self) Is Nothing Then Exit Sub
    widgScheduleTask WzFunc_new(FN_PopoverWidget_openTask, self, parent)
End Sub

' (VB) the body of open()'s scheduled task
Private Sub PopoverWidget_openTask(ByVal popoverWidget As WIDGET, ByVal strongParent As WIDGET)
    Dim lockedScreen As W_SCREEN, psOverlayScreen As W_SCREEN, newRootFrm As WIDGET, ovl As W_FULLSCREENOVERLAY_CLICKFORM
    If Not strongParent Is Nothing Then
        Set lockedScreen = strongParent.screenPointer
    End If
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "lib/widget/popover.cpp:112", "open", "The parent does not have an associated screen pointer?"

    Set psOverlayScreen = W_SCREEN_make()
    psOverlayScreen.psForm.hide

    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    Set newRootFrm.displayFunction = WzFunc_new(FN_PopoverWidget_displayChildDropShadows)
    Set ovl = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm)
    Set ovl.onClickedFunc = WzFunc_new(FN_PopoverWidget_onClicked, ObjPtr(psOverlayScreen), PopoverWidget_(popoverWidget).weakSerial)
    Set ovl.onCancelPressed = ovl.onClickedFunc
    If PopoverWidget_(popoverWidget).style = PopoverWidget_Style_NonInteractive Then
        newRootFrm.setTransparentToMouse True
        newRootFrm.setCustomHitTest WzFunc_new(FN_PopoverWidget_noHitTest)
    End If
    psOverlayScreen.psForm.attach newRootFrm

    Set PopoverWidget_(popoverWidget).overlayScreen = psOverlayScreen
    newRootFrm.attach popoverWidget

    popoverWidget.setCalcLayout WzFunc_new(FN_PopoverWidget_calcLayout, strongParent)

    widgRegisterOverlayScreenOnTopOfScreen psOverlayScreen, lockedScreen
End Sub

Private Sub PopoverWidget_closeImpl(ByVal self As WIDGET)
    Dim d As PopoverWidget, strongOverlay As W_SCREEN
    Set d = PopoverWidget_(self)
    self.setCalcLayout Nothing      ' (VB) the calcLayout lambda holds the parent
    Set strongOverlay = PopoverWidget_lockOverlayScreen(self)
    If strongOverlay Is Nothing Then
        Exit Sub
    End If
    widgRemoveOverlayScreen strongOverlay
    Set d.overlayScreen = Nothing   ' overlayScreen.reset()
    If Not d.onCloseHandler Is Nothing Then
        WzFunc_call d.onCloseHandler
    End If
End Sub

Public Sub PopoverWidget_close(ByVal self As WIDGET)
    ' because open() uses widgScheduleTask to actually create the overlay screen,
    ' widgScheduleTask must also be used to close it (otherwise, overlayScreen might be accessed before the scheduled open task creates it)
    widgScheduleTask WzFunc_new(FN_PopoverWidget_closeTask, self)
End Sub

Public Sub PopoverWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' currently, no-op
End Sub

Public Function PopoverWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As PopoverWidget
    Set d = PopoverWidget_(self)
    If d.popoverContents Is Nothing Then
        PopoverWidget_idealWidth = 0
        Exit Function
    End If
    PopoverWidget_idealWidth = d.popoverContents.idealWidth()
End Function

Public Function PopoverWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As PopoverWidget
    Set d = PopoverWidget_(self)
    If d.popoverContents Is Nothing Then
        PopoverWidget_idealHeight = 0
        Exit Function
    End If
    PopoverWidget_idealHeight = d.popoverContents.idealHeight()
End Function
