Attribute VB_Name = "dropdown_c"
' lib/widget/dropdown.cpp, dropdown.h: the widget classes DropdownOverlay (of
' dropdown.cpp), DropdownItemWrapper and DropdownWidget (docs/CONVENTIONS.md,
' "Widgets"; their fields are DropdownOverlay.cls, DropdownItemWrapper.cls and
' DropdownWidget.cls). enum class DropdownWidget::DropdownMenuStyle is
' DropdownWidget_DropdownMenuStyle. items is a Collection (element n is item
' n + 1); optional<AtlasImage> dropdownCaretImage is has_dropdownCaretImage and
' dropdownCaretImage; setDropdownCaretImage takes has_image, image, displaySize
' and padding (a blank Padding for the C++ default {}); getSelectedIndex returns
' has_value and sets the index ByRef. The handlers are WzFuncs:
' DropdownOnSelectHandler (item), canChange (dropdown, newIndex,
' newSelectedWidget) returning a Boolean, onChange / onOpen / onClose (dropdown).
' The weak pointers (pWeakThis, weakSelf, DropdownItemWrapper::parent) are
' serials (widgweak_vb).
' (VB) widget class DropdownOverlay : WIDGET
' (VB) widget class DropdownItemWrapper : WIDGET
' (VB) widget class DropdownWidget : WIDGET
' (VB) virtual DropdownWidget.drawOpenedHighlight
' (VB) virtual DropdownWidget.drawSelectedItem
' (VB) virtual DropdownWidget.onSelectedItemChanged
' (VB) virtual DropdownWidget.calculateDropdownListScreenPosX
' (VB) virtual DropdownWidget.calculateDropdownListDisplayWidth
Option Explicit

' enum class DropdownMenuStyle
Public Enum DropdownWidget_DropdownMenuStyle
    DropdownWidget_DropdownMenuStyle_InPlace    ' dropdown menu appears *overtop* of the currently-selected item (i.e. in-place)
    DropdownWidget_DropdownMenuStyle_Separate   ' dropdown menu appears below (or above if not enough room) the parent DropdownWidget
End Enum

Public Const FN_DropdownOverlay_calcLayout As Long = 2110
Public Const FN_DropdownItemWrapper_selectTask As Long = 2111
Public Const FN_DropdownWidget_openTask As Long = 2112
Public Const FN_DropdownWidget_overlayClicked As Long = 2113
Public Const FN_DropdownWidget_closeTask As Long = 2114
Public Const FN_DropdownWidget_itemOnSelect As Long = 2115

Public Function dropdown_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, strongSelf As WIDGET, dropdownWidget As WIDGET, psStrongDropdown As WIDGET
    Select Case f.id
    Case FN_DropdownOverlay_calcLayout
        ' LAMBDA_CALCLAYOUT_SIMPLE: (WIDGET *psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth - 1, screenHeight - 1
    Case FN_DropdownItemWrapper_selectTask
        ' [weakSelf]()
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "lib/widget/dropdown.cpp:97", "findMouseTargetRecursive", "DropdownItemWrapper disappeared?": Exit Function
        WzFunc_call DropdownItemWrapper_(strongSelf).onSelect, strongSelf
    Case FN_DropdownWidget_openTask
        ' [pWeakThis]()
        Set dropdownWidget = WeakWidget_lock(f.cap(0))
        If Not dropdownWidget Is Nothing Then
            DropdownWidget_openTask dropdownWidget, f.cap(0)
        End If
    Case FN_DropdownWidget_overlayClicked
        ' [pWeakThis]()
        Set dropdownWidget = WeakWidget_lock(f.cap(0))
        If Not dropdownWidget Is Nothing Then
            DropdownWidget_close dropdownWidget
        End If
    Case FN_DropdownWidget_closeTask
        ' [pWeakThis]()
        Set dropdownWidget = WeakWidget_lock(f.cap(0))
        If Not dropdownWidget Is Nothing Then
            DropdownWidget_closeTask dropdownWidget
        End If
    Case FN_DropdownWidget_itemOnSelect
        ' [newItemIndex, pWeakThis](std::shared_ptr<DropdownItemWrapper> selected)
        Set psStrongDropdown = WeakWidget_lock(f.cap(1))
        If Not (Not psStrongDropdown Is Nothing) Then ASSERT_FAILURE "psStrongDropdown != nullptr", "lib/widget/dropdown.cpp:394", "addItem", "DropdownWidget no longer exists?": Exit Function
        If DropdownWidget_(psStrongDropdown).overlayScreen Is Nothing Then
            DropdownWidget_open psStrongDropdown
            Exit Function
        End If

        If DropdownWidget_select(psStrongDropdown, args(0), f.cap(0)) Then
            DropdownWidget_close psStrongDropdown
        End If
    End Select
End Function

'--------------------------------------------------------------------------
' DropdownOverlay

' std::make_shared<DropdownOverlay>(onClickedFunc)
Public Function DropdownOverlay_new(ByVal onClickedFunc As WzFunc) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_DropdownOverlay)
    DropdownOverlay_construct w, onClickedFunc
    Set DropdownOverlay_new = w
End Function

' DropdownOverlay(std::function<void ()> onClickedFunc)
Public Sub DropdownOverlay_construct(ByVal self As WIDGET, ByVal onClickedFunc As WzFunc)
    WIDGET_construct self
    self.kind = K_DropdownOverlay ' (VB) virtual calls from here on reach DropdownOverlay's overrides
    Set DropdownOverlay_(self).onClickedFunc = onClickedFunc
    self.setCalcLayout WzFunc_new(FN_DropdownOverlay_calcLayout)
End Sub

Public Sub DropdownOverlay_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    WzFunc_call DropdownOverlay_(self).onClickedFunc
End Sub

Public Sub DropdownOverlay_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    displayChildDropShadows self, xOffset, yOffset
End Sub

'--------------------------------------------------------------------------
' DropdownItemWrapper

' DropdownItemWrapper() {}
Public Sub DropdownItemWrapper_construct(ByVal self As WIDGET)
    Dim d As DropdownItemWrapper
    WIDGET_construct self
    self.kind = K_DropdownItemWrapper ' (VB) virtual calls from here on reach DropdownItemWrapper's overrides
    Set d = DropdownItemWrapper_(self)
    d.weakSerial = WeakWidget_register(self)
    d.selected = False
    d.mouseDownOnWrapper = False
End Sub

Public Sub DropdownItemWrapper_destruct(ByVal self As WIDGET)
    WeakWidget_unregister DropdownItemWrapper_(self).weakSerial
End Sub

' static std::shared_ptr<DropdownItemWrapper> make(const std::shared_ptr<DropdownWidget>& parent, const std::shared_ptr<WIDGET> &item, DropdownOnSelectHandler onSelect)
Public Function DropdownItemWrapper_make(ByVal parent As WIDGET, ByVal item As WIDGET, ByVal onSelect As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_DropdownItemWrapper)
    DropdownItemWrapper_construct widget
    DropdownItemWrapper_initialize widget, parent, item, onSelect
    Set DropdownItemWrapper_make = widget
End Function

Private Sub DropdownItemWrapper_initialize(ByVal self As WIDGET, ByVal newParent As WIDGET, ByVal newItem As WIDGET, ByVal newOnSelect As WzFunc)
    Dim d As DropdownItemWrapper
    Set d = DropdownItemWrapper_(self)
    Set d.item = newItem
    self.attach d.item
    Set d.onSelect = newOnSelect
    d.parent = DropdownWidget_(newParent).weakSerial
End Sub

Public Function DropdownItemWrapper_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim d As DropdownItemWrapper, inputContext As W_CONTEXT, result As WIDGET, psStrongParent As WIDGET
    Set d = DropdownItemWrapper_(self)
    inputContext = psContext
    Set result = WIDGET_findMouseTargetRecursive(self, psContext, key, wasPressed)

    If key <> WKEY_NONE Then
        If self.hitTest(inputContext.mx, inputContext.my) Then
            Set psStrongParent = WeakWidget_lock(d.parent)
            If Not psStrongParent Is Nothing Then
                DropdownWidget_setMouseClickOnItem psStrongParent, self, key, wasPressed
            End If

            If wasPressed Then
                d.mouseDownOnWrapper = True
            Else
                ' if mouse down + mouse up on the item
                If d.mouseDownOnWrapper Then
                    d.mouseDownOnWrapper = False

                    ' Instead of calling the onSelect handler immediately, schedule a task.
                    '
                    ' (The onSelect handler may call close(), and we want that to be delayed by an event processing cycle
                    ' so the mouse event gets properly attributed to the overlay screen by other code - such as the game
                    ' event processing loop, which currently usese isMouseOverScreenOverlayChild)
                    widgScheduleTask WzFunc_new(FN_DropdownItemWrapper_selectTask, d.weakSerial)
                End If
            End If
        End If
    End If

    Set DropdownItemWrapper_findMouseTargetRecursive = result
End Function

Public Sub DropdownItemWrapper_clearMouseDownState(ByVal self As WIDGET)
    DropdownItemWrapper_(self).mouseDownOnWrapper = False
End Sub

Public Sub DropdownItemWrapper_geometryChanged(ByVal self As WIDGET)
    DropdownItemWrapper_(self).item.setGeometry 0, 0, self.width(), self.height()
End Sub

Public Sub DropdownItemWrapper_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long
    If DropdownItemWrapper_(self).selected Then
        x0 = xOffset + self.x()
        y0 = yOffset + self.y()
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), psPalette(WZCOL_MENU_SCORE_BUILT)
    End If
End Sub

Public Function DropdownItemWrapper_getItem(ByVal self As WIDGET) As WIDGET
    Set DropdownItemWrapper_getItem = DropdownItemWrapper_(self).item
End Function

Public Sub DropdownItemWrapper_setSelected(ByVal self As WIDGET, ByVal value As Boolean)
    DropdownItemWrapper_(self).selected = value
End Sub

Public Function DropdownItemWrapper_idealWidth(ByVal self As WIDGET) As Long
    DropdownItemWrapper_idealWidth = DropdownItemWrapper_(self).item.idealWidth()
End Function

Public Function DropdownItemWrapper_idealHeight(ByVal self As WIDGET) As Long
    DropdownItemWrapper_idealHeight = DropdownItemWrapper_(self).item.idealHeight()
End Function

'--------------------------------------------------------------------------
' DropdownWidget

' std::make_shared<DropdownWidget>()
Public Function DropdownWidget_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_DropdownWidget)
    DropdownWidget_construct w
    Set DropdownWidget_new = w
End Function

' DropdownWidget::DropdownWidget()
Public Sub DropdownWidget_construct(ByVal self As WIDGET)
    Dim d As DropdownWidget, p As Padding
    WIDGET_construct self
    self.kind = K_DropdownWidget ' (VB) virtual calls from here on reach DropdownWidget's overrides
    Set d = DropdownWidget_(self)
    d.weakSerial = WeakWidget_register(self)
    d.overlayYPosOffset = 0
    d.has_dropdownCaretImage = False
    d.dropdownCaretImagePadding = p     ' = {}
    d.menuStyle = DropdownWidget_DropdownMenuStyle_InPlace
    d.isDisabled = False

    Set d.itemsList = ScrollableListWidget_make()
    ScrollableListWidget_setBackgroundColor d.itemsList, psPalette(WZCOL_MENU_BACKGROUND)
    DropdownWidget_setListHeight self, 100
End Sub

Public Sub DropdownWidget_destruct(ByVal self As WIDGET)
    WeakWidget_unregister DropdownWidget_(self).weakSerial
End Sub

Public Sub DropdownWidget_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As DropdownWidget
    Set d = DropdownWidget_(self)
    If Not d.overlayScreen Is Nothing Then
        d.itemsList.setGeometry vt_DropdownWidget_calculateDropdownListScreenPosX(self), DropdownWidget_calculateDropdownListScreenPosY(self), d.itemsList.width(), d.itemsList.height()

        If keyPressed(KEY_ESC) Then
            inputLoseFocus  ' clear the input buffer.
            DropdownWidget_close self
        End If
    Else
        DropdownWidget_callRunOnItems self
    End If
End Sub

Public Sub DropdownWidget_callRunOnItems(ByVal self As WIDGET)
    Dim d As DropdownWidget, i As Long, wrapper As WIDGET, inner As WIDGET, ctx As W_CONTEXT, zeroCtx As W_CONTEXT
    Set d = DropdownWidget_(self)
    ' Some child widgets might use run() to clear cached textures (etc), so call run regularly to allow them
    For i = 1 To d.items.Count
        Set wrapper = d.items(i)
        Set inner = DropdownItemWrapper_(wrapper).item
        ctx = zeroCtx                   ' W_CONTEXT::ZeroContext()
        inner.runRecursive ctx
        inner.manuallyCallRun ctx
    Next
End Sub

Public Sub DropdownWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As DropdownWidget, w As Long, h As Long, itemWidth As Long, i As Long, item As WIDGET
    Set d = DropdownWidget_(self)
    w = self.width()
    h = self.height()
    If w = 0 Or h = 0 Then
        Exit Sub
    End If
    itemWidth = DropdownWidget_getItemDisplayWidth(self)
    d.itemsList.setGeometry d.itemsList.x(), d.itemsList.y(), vt_DropdownWidget_calculateDropdownListDisplayWidth(self), d.itemsList.height()
    Select Case d.menuStyle
    Case DropdownWidget_DropdownMenuStyle_InPlace
        For i = 1 To d.items.Count
            Set item = d.items(i)
            item.setGeometry 0, 0, itemWidth, h
        Next
    Case DropdownWidget_DropdownMenuStyle_Separate
    End Select
End Sub

Public Sub DropdownWidget_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    DropdownWidget_open self
End Sub

Public Function DropdownWidget_calculateDropdownListScreenPosX(ByVal self As WIDGET) As Long
    DropdownWidget_calculateDropdownListScreenPosX = self.screenPosX()
End Function

Public Function DropdownWidget_calculateDropdownListDisplayWidth(ByVal self As WIDGET) As Long
    Select Case DropdownWidget_(self).menuStyle
    Case DropdownWidget_DropdownMenuStyle_InPlace
        DropdownWidget_calculateDropdownListDisplayWidth = DropdownWidget_getItemDisplayWidth(self)
        Exit Function
    Case DropdownWidget_DropdownMenuStyle_Separate
        DropdownWidget_calculateDropdownListDisplayWidth = DropdownWidget_(self).itemsList.idealWidth()
        Exit Function
    End Select
    DropdownWidget_calculateDropdownListDisplayWidth = 0 ' silence warning
End Function

Public Sub DropdownWidget_setListHeight(ByVal self As WIDGET, ByVal value As Long)
    Dim l As WIDGET
    Set l = DropdownWidget_(self).itemsList
    l.setGeometry l.x(), l.y(), l.width(), value
End Sub

Public Sub DropdownWidget_setListBackgroundColor(ByVal self As WIDGET, ByVal color As Long)
    ScrollableListWidget_setBackgroundColor DropdownWidget_(self).itemsList, color
End Sub

Public Function DropdownWidget_getDropdownMenuOuterPadding(ByVal self As WIDGET) As Padding
    DropdownWidget_getDropdownMenuOuterPadding = ScrollableListWidget_getPadding(DropdownWidget_(self).itemsList)
End Function

Public Function DropdownWidget_setSelectedIndex(ByVal self As WIDGET, ByVal index As Long) As Boolean
    Dim d As DropdownWidget
    Set d = DropdownWidget_(self)
    If Not (index >= 0 And index < d.items.Count) Then ASSERT_FAILURE "index < items.size()", "lib/widget/dropdown.h:113", "setSelectedIndex", "Invalid dropdown item index": DropdownWidget_setSelectedIndex = False: Exit Function
    DropdownWidget_setSelectedIndex = DropdownWidget_select(self, d.items(index + 1), index)
End Function

Public Sub DropdownWidget_setCanChange(ByVal self As WIDGET, ByVal value As WzFunc)
    Set DropdownWidget_(self).canChange = value
End Sub

Public Sub DropdownWidget_setOnChange(ByVal self As WIDGET, ByVal value As WzFunc)
    Set DropdownWidget_(self).onChange = value
End Sub

Public Sub DropdownWidget_setOnOpen(ByVal self As WIDGET, ByVal value As WzFunc)
    Set DropdownWidget_(self).onOpen = value
End Sub

Public Sub DropdownWidget_setOnClose(ByVal self As WIDGET, ByVal value As WzFunc)
    Set DropdownWidget_(self).onClose = value
End Sub

Public Function DropdownWidget_getItem(ByVal self As WIDGET, ByVal idx As Long) As WIDGET
    Dim d As DropdownWidget
    Set d = DropdownWidget_(self)
    If idx < 0 Or idx >= d.items.Count Then
        Exit Function
    End If
    Set DropdownWidget_getItem = DropdownItemWrapper_getItem(d.items(idx + 1))
End Function

Public Function DropdownWidget_getSelectedItem(ByVal self As WIDGET) As WIDGET
    Dim d As DropdownWidget
    Set d = DropdownWidget_(self)
    If d.selectedItem Is Nothing Then
        Exit Function
    End If
    Set DropdownWidget_getSelectedItem = DropdownItemWrapper_getItem(d.selectedItem)
End Function

' nonstd::optional<size_t> getSelectedIndex() const
Public Function DropdownWidget_getSelectedIndex(ByVal self As WIDGET, value As Long) As Boolean
    Dim d As DropdownWidget, i As Long
    Set d = DropdownWidget_(self)
    If d.selectedItem Is Nothing Then
        Exit Function                   ' nonstd::nullopt
    End If

    For i = 1 To d.items.Count
        If d.items(i) Is d.selectedItem Then
            value = i - 1
            DropdownWidget_getSelectedIndex = True
            Exit Function
        End If
    Next
End Function

Public Function DropdownWidget_getScrollbarWidth(ByVal self As WIDGET) As Long
    Dim d As DropdownWidget
    Set d = DropdownWidget_(self)
    If Not (Not d.itemsList Is Nothing) Then ASSERT_FAILURE "itemsList != nullptr", "lib/widget/dropdown.h:167", "getScrollbarWidth", "null itemsList": DropdownWidget_getScrollbarWidth = 0: Exit Function
    DropdownWidget_getScrollbarWidth = ScrollableListWidget_getScrollbarWidth(d.itemsList)
End Function

Private Function DropdownWidget_calculateDropdownListScreenPosY(ByVal self As WIDGET) As Long
    Dim d As DropdownWidget, yPosOffset As Long, dropDownOverlayPosY As Long
    Set d = DropdownWidget_(self)
    yPosOffset = 0
    Select Case d.menuStyle
    Case DropdownWidget_DropdownMenuStyle_InPlace
        yPosOffset = -d.overlayYPosOffset
    Case DropdownWidget_DropdownMenuStyle_Separate
        yPosOffset = self.height()
    End Select

    dropDownOverlayPosY = self.screenPosY() + yPosOffset
    If ULT(screenHeight, dropDownOverlayPosY + d.itemsList.height()) Then
        ' Positioning the dropdown below would cause it to appear partially or fully offscreen
        ' So, instead, position it above
        dropDownOverlayPosY = self.screenPosY() - d.itemsList.height()
        If dropDownOverlayPosY < 0 Then
            ' Well, this would be off-screen too...
            ' For now: Just position it at the top
            dropDownOverlayPosY = 0
        End If
    End If
    DropdownWidget_calculateDropdownListScreenPosY = dropDownOverlayPosY
End Function

Public Sub DropdownWidget_open(ByVal self As WIDGET)
    Dim d As DropdownWidget
    Set d = DropdownWidget_(self)
    If Not d.overlayScreen Is Nothing Then Exit Sub
    Set d.mouseDownItem = Nothing
    widgScheduleTask WzFunc_new(FN_DropdownWidget_openTask, d.weakSerial)
End Sub

' (VB) the body of open()'s scheduled task, dropdownWidget locked; pWeakThis is its serial
Private Sub DropdownWidget_openTask(ByVal dropdownWidget As WIDGET, ByVal pWeakThis As Long)
    Dim d As DropdownWidget, selectedItemIndex As Long, listViewTopIndex As Long, maxOffset As Long, newRootFrm As WIDGET
    Set d = DropdownWidget_(dropdownWidget)
    If Not d.onOpen Is Nothing Then
        WzFunc_call d.onOpen, dropdownWidget
    End If

    Set d.overlayScreen = W_SCREEN_make()

    Select Case d.menuStyle
    Case DropdownWidget_DropdownMenuStyle_InPlace
        ' calculate the ideal position so that the dropdown appears *overtop* of the currently-selected item (in-place)
        If Not DropdownWidget_getSelectedIndex(dropdownWidget, selectedItemIndex) Then selectedItemIndex = 0
        listViewTopIndex = 0
        If d.itemsList.idealHeight() > d.itemsList.height() Then
            maxOffset = MIN_(2, ScrollableListWidget_numItems(d.itemsList))
            If selectedItemIndex > maxOffset Then
                listViewTopIndex = selectedItemIndex - maxOffset
            End If
            ScrollableListWidget_scrollToItem d.itemsList, listViewTopIndex
        End If
        d.overlayYPosOffset = ScrollableListWidget_getCurrentYPosOfItem(d.itemsList, selectedItemIndex)
    Case DropdownWidget_DropdownMenuStyle_Separate
    End Select

    Set newRootFrm = DropdownOverlay_new(WzFunc_new(FN_DropdownWidget_overlayClicked, pWeakThis))
    d.overlayScreen.psForm.attach newRootFrm

    d.itemsList.setGeometry vt_DropdownWidget_calculateDropdownListScreenPosX(dropdownWidget), DropdownWidget_calculateDropdownListScreenPosY(dropdownWidget), vt_DropdownWidget_calculateDropdownListDisplayWidth(dropdownWidget), d.itemsList.height()
    newRootFrm.attach d.itemsList

    widgRegisterOverlayScreenOnTopOfScreen d.overlayScreen, dropdownWidget.screenPointer
End Sub

Public Function DropdownWidget_isOpen(ByVal self As WIDGET) As Boolean
    DropdownWidget_isOpen = Not DropdownWidget_(self).overlayScreen Is Nothing
End Function

Public Sub DropdownWidget_close(ByVal self As WIDGET)
    Dim d As DropdownWidget
    Set d = DropdownWidget_(self)
    If Not (DropdownWidget_isOpen(self)) Then ASSERT_FAILURE "isOpen()", "lib/widget/dropdown.cpp:305", "close", "Dropdown isn't open": Exit Sub
    Set d.mouseDownItem = Nothing
    widgScheduleTask WzFunc_new(FN_DropdownWidget_closeTask, d.weakSerial)
End Sub

' (VB) the body of close()'s scheduled task, dropdownWidget locked
Private Sub DropdownWidget_closeTask(ByVal dropdownWidget As WIDGET)
    Dim d As DropdownWidget
    Set d = DropdownWidget_(dropdownWidget)
    If Not d.overlayScreen Is Nothing Then
        If Not d.onClose Is Nothing Then
            WzFunc_call d.onClose, dropdownWidget
        End If

        widgRemoveOverlayScreen d.overlayScreen
        If Not d.overlayScreen.psForm Is Nothing Then
            d.overlayScreen.psForm.detach d.itemsList
        End If
        Set d.overlayScreen = Nothing
    End If
End Sub

Public Function DropdownWidget_getItemDisplayWidth(ByVal self As WIDGET) As Long
    DropdownWidget_getItemDisplayWidth = MAX_(self.width() - DropdownWidget_getCaretImageUsedWidth(self), 0)
End Function

Public Function DropdownWidget_getCaretImageUsedWidth(ByVal self As WIDGET) As Long
    Dim d As DropdownWidget, p As Padding
    Set d = DropdownWidget_(self)
    p = d.dropdownCaretImagePadding
    DropdownWidget_getCaretImageUsedWidth = d.dropdownCaretImageSize.width_ + p.left_ + p.right_
End Function

Public Sub DropdownWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As DropdownWidget, x0 As Long, y0 As Long, itemWidth As Long, displayArea As WzRect
    Set d = DropdownWidget_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    If DropdownWidget_isOpen(self) Then
        vt_DropdownWidget_drawOpenedHighlight self, xOffset, yOffset
    End If

    itemWidth = DropdownWidget_getItemDisplayWidth(self)
    DropdownWidget_drawDropdownCaretImage self, xOffset, yOffset, psPalette(WZCOL_TEXT_MEDIUM)

    If Not d.selectedItem Is Nothing Then
        displayArea = WzRect_(x0, y0, itemWidth, self.height())
        vt_DropdownWidget_drawSelectedItem self, DropdownItemWrapper_getItem(d.selectedItem), displayArea
    End If
End Sub

Public Sub DropdownWidget_drawDropdownCaretImage(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long, ByVal color As Long)
    Dim d As DropdownWidget, x0 As Long, y0 As Long, caretX0 As Long, caretY0 As Long, img As AtlasImage, sz As WzSize
    Set d = DropdownWidget_(self)
    If Not d.has_dropdownCaretImage Then
        Exit Sub
    End If
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    sz = d.dropdownCaretImageSize
    caretX0 = x0 + DropdownWidget_getItemDisplayWidth(self) + d.dropdownCaretImagePadding.left_
    caretY0 = y0 + ((self.height() - sz.height_) \ 2)
    img = d.dropdownCaretImage
    iV_DrawImageFileAnisotropicTint img.images, img.id, caretX0, caretY0, CSng(sz.width_), CSng(sz.height_), color
End Sub

Public Sub DropdownWidget_drawOpenedHighlight(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, w As Long, h As Long
    x0 = MAX_(0, (xOffset + self.x()) - 8)
    y0 = yOffset + self.y()
    w = self.width()
    h = self.height()
    pie_UniTransBoxFill x0, y0, x0 + w + 16, y0 + h, psPalette(WZCOL_MENU_SCORE_BUILT)
End Sub

Public Sub DropdownWidget_drawSelectedItem(ByVal self As WIDGET, ByVal item As WIDGET, displayArea As WzRect)
    Dim context As WidgetGraphicsContext, base As WidgetGraphicsContext
    WidgetGraphicsContext_construct base
    context = WidgetGraphicsContext_translatedBy(base, displayArea.x, displayArea.y)
    item.displayRecursive context
End Sub

Public Sub DropdownWidget_addItem(ByVal self As WIDGET, ByVal item As WIDGET)
    Dim d As DropdownWidget, newItemIndex As Long, itemOnSelect As WzFunc, wrapper As WIDGET
    Set d = DropdownWidget_(self)
    newItemIndex = d.items.Count
    Set itemOnSelect = WzFunc_new(FN_DropdownWidget_itemOnSelect, newItemIndex, d.weakSerial)

    Set wrapper = DropdownItemWrapper_make(self, item, itemOnSelect)
    Select Case d.menuStyle
    Case DropdownWidget_DropdownMenuStyle_InPlace
        wrapper.setGeometry 0, 0, MAX_(DropdownWidget_getItemDisplayWidth(self), 0), self.height()
    Case DropdownWidget_DropdownMenuStyle_Separate
        wrapper.setGeometry 0, 0, MAX_(item.width(), 0), item.height()
    End Select

    d.items.Add wrapper
    ScrollableListWidget_addItem d.itemsList, wrapper
End Sub

Public Sub DropdownWidget_clear(ByVal self As WIDGET)
    Dim d As DropdownWidget
    Set d = DropdownWidget_(self)
    Set d.items = New Collection
    ScrollableListWidget_clear d.itemsList
    Set d.selectedItem = Nothing
    Set d.mouseOverItem = Nothing
    Set d.mouseDownItem = Nothing
End Sub

Public Function DropdownWidget_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim d As DropdownWidget, shiftedContext As W_CONTEXT, deltaX As Long, deltaY As Long, highlightedSelectedItemChildWidget As WIDGET
    Set d = DropdownWidget_(self)
    If d.overlayScreen Is Nothing And Not d.selectedItem Is Nothing And Not d.isDisabled And key = WKEY_NONE Then ' only forward highlighting events
        shiftedContext = psContext
        deltaX = self.x() - d.selectedItem.x()
        deltaY = self.y() - d.selectedItem.y()
        shiftedContext.mx = psContext.mx - deltaX
        shiftedContext.my = psContext.my - deltaY
        shiftedContext.xOffset = psContext.xOffset + deltaX
        shiftedContext.yOffset = psContext.yOffset + deltaY
        Set highlightedSelectedItemChildWidget = d.selectedItem.findMouseTargetRecursive(shiftedContext, key, wasPressed)
        If Not highlightedSelectedItemChildWidget Is Nothing Then
            psContext = shiftedContext  ' bubble-up the matching hit's context
            Set DropdownWidget_findMouseTargetRecursive = highlightedSelectedItemChildWidget
            Exit Function
        End If
    End If

    Set DropdownWidget_findMouseTargetRecursive = WIDGET_findMouseTargetRecursive(self, psContext, key, wasPressed)
End Function

Private Sub DropdownWidget_setMouseClickOnItem(ByVal self As WIDGET, ByVal item As WIDGET, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean)
    Dim d As DropdownWidget
    Set d = DropdownWidget_(self)
    If item Is d.mouseDownItem Then Exit Sub

    If Not d.mouseDownItem Is Nothing Then
        DropdownItemWrapper_clearMouseDownState d.mouseDownItem
    End If

    If wasPressed Then
        Set d.mouseDownItem = item
    Else
        Set d.mouseDownItem = Nothing
    End If
End Sub

' void setDropdownCaretImage(optional<AtlasImage> image, const WzSize& displaySize, const Padding& padding = {})
Public Sub DropdownWidget_setDropdownCaretImage(ByVal self As WIDGET, ByVal has_image As Boolean, image As AtlasImage, displaySize As WzSize, padding_ As Padding)
    Dim d As DropdownWidget, p As Padding, img As AtlasImage
    Set d = DropdownWidget_(self)
    d.has_dropdownCaretImage = has_image
    d.dropdownCaretImage = image
    d.dropdownCaretImageSize = WzSize_(0, 0)
    d.dropdownCaretImagePadding = p     ' Padding{}

    If d.has_dropdownCaretImage Then
        img = d.dropdownCaretImage
        If img.images = 0 Then
            d.has_dropdownCaretImage = False    ' reset()
            Exit Sub
        End If
        If Not ULT(img.id, IMAGEFILE_numImages(img.images)) Then
            d.has_dropdownCaretImage = False
            Exit Sub
        End If
        d.dropdownCaretImageSize = displaySize
        d.dropdownCaretImagePadding = padding_
    End If
End Sub

Public Sub DropdownWidget_setDisabled(ByVal self As WIDGET, ByVal isDisabled_ As Boolean)
    DropdownWidget_(self).isDisabled = isDisabled_
End Sub

Public Function DropdownWidget_getIsDisabled(ByVal self As WIDGET) As Boolean
    DropdownWidget_getIsDisabled = DropdownWidget_(self).isDisabled
End Function

Public Sub DropdownWidget_setStyle(ByVal self As WIDGET, ByVal menuStyle_ As DropdownWidget_DropdownMenuStyle)
    Dim d As DropdownWidget, p As Padding
    Set d = DropdownWidget_(self)
    d.menuStyle = menuStyle_
    Select Case d.menuStyle
    Case DropdownWidget_DropdownMenuStyle_InPlace
        ScrollableListWidget_setPadding d.itemsList, p     ' Padding{}
    Case DropdownWidget_DropdownMenuStyle_Separate
        p.top = 5                       ' Padding{5,0,5,0}
        p.right_ = 0
        p.bottom = 5
        p.left_ = 0
        ScrollableListWidget_setPadding d.itemsList, p
    End Select
End Sub

Public Function DropdownWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim result As Long
    result = DropdownWidget_(self).itemsList.idealWidth()
    result = result + DropdownWidget_getCaretImageUsedWidth(self)
    DropdownWidget_idealWidth = result
End Function

Public Function DropdownWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As DropdownWidget, maxValue As Long, i As Long, item As WIDGET
    Set d = DropdownWidget_(self)
    maxValue = 0
    For i = 1 To d.items.Count
        Set item = d.items(i)
        maxValue = MAX_(maxValue, item.idealHeight())
    Next

    DropdownWidget_idealHeight = maxValue
End Function

Private Function DropdownWidget_select(ByVal self As WIDGET, ByVal selected As WIDGET, ByVal selectedIndex As Long) As Boolean
    Dim d As DropdownWidget, selectedWidget As WIDGET
    Set d = DropdownWidget_(self)
    If d.selectedItem Is selected Then
        DropdownWidget_select = True
        Exit Function
    End If

    If Not d.canChange Is Nothing Then
        If Not selected Is Nothing Then Set selectedWidget = DropdownItemWrapper_getItem(selected)
        If Not CBool(WzFunc_call(d.canChange, self, selectedIndex, selectedWidget)) Then
            ' abort change
            DropdownWidget_select = False
            Exit Function
        End If
    End If

    If Not d.selectedItem Is Nothing Then
        DropdownItemWrapper_setSelected d.selectedItem, False
    End If
    Set d.selectedItem = selected
    If d.menuStyle = DropdownWidget_DropdownMenuStyle_InPlace Then
        DropdownItemWrapper_setSelected d.selectedItem, True
    End If

    vt_DropdownWidget_onSelectedItemChanged self
    If Not d.onChange Is Nothing Then
        WzFunc_call d.onChange, self
    End If

    DropdownWidget_select = True
End Function

Public Sub DropdownWidget_onSelectedItemChanged(ByVal self As WIDGET)
    ' currently, no-op - intended for subclasses to override
End Sub
