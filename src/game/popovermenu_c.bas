Attribute VB_Name = "popovermenu_c"
' lib/widget/popovermenu.cpp, popovermenu.h: the widget classes MenuItemWrapper
' and PopoverMenuWidget (docs/CONVENTIONS.md, "Widgets"; their fields are
' MenuItemWrapper.cls and PopoverMenuWidget.cls). The weak pointers
' (MenuItemWrapper::parent, PopoverMenuWidget::currentPopover) are serials
' (widgweak_vb). openMenu's Vector2i positionOffset is two Longs.
' (VB) widget class MenuItemWrapper : WIDGET
' (VB) widget class PopoverMenuWidget : WIDGET
Option Explicit

'--------------------------------------------------------------------------
' MARK: - MenuItemWrapper

' MenuItemWrapper::MenuItemWrapper()
Public Sub MenuItemWrapper_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_MenuItemWrapper ' (VB) virtual calls from here on reach MenuItemWrapper's overrides
    MenuItemWrapper_(self).closeMenuOnClick = False
    MenuItemWrapper_(self).mouseDownOnWrapper = False
End Sub

Private Sub MenuItemWrapper_initialize(ByVal self As WIDGET, ByVal psParent As WIDGET, ByVal newItem As WIDGET, ByVal closeMenuOnClickFunc As Boolean)
    Dim d As MenuItemWrapper
    Set d = MenuItemWrapper_(self)
    Set d.item = newItem
    d.parent = PopoverMenuWidget_(psParent).weakSerial
    d.closeMenuOnClick = closeMenuOnClickFunc
    self.attach d.item
    self.setGeometry 0, 0, self.idealWidth(), self.idealHeight()
End Sub

' static std::shared_ptr<MenuItemWrapper> wrap(const std::shared_ptr<PopoverMenuWidget>& parent, const std::shared_ptr<WIDGET> &item, bool closeMenuOnClick)
Public Function MenuItemWrapper_wrap(ByVal parent As WIDGET, ByVal newItem As WIDGET, ByVal closeMenuOnClick As Boolean) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_MenuItemWrapper)
    MenuItemWrapper_construct widget
    MenuItemWrapper_initialize widget, parent, newItem, closeMenuOnClick
    Set MenuItemWrapper_wrap = widget
End Function

Public Function MenuItemWrapper_getItem(ByVal self As WIDGET) As WIDGET
    Set MenuItemWrapper_getItem = MenuItemWrapper_(self).item
End Function

Public Function MenuItemWrapper_getCloseMenuOnClick(ByVal self As WIDGET) As Boolean
    MenuItemWrapper_getCloseMenuOnClick = MenuItemWrapper_(self).closeMenuOnClick
End Function

Public Function MenuItemWrapper_findMouseTargetRecursive(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean) As WIDGET
    Dim d As MenuItemWrapper, inputContext As W_CONTEXT, result As WIDGET, psStrongParent As WIDGET
    Set d = MenuItemWrapper_(self)
    inputContext = psContext
    Set result = WIDGET_findMouseTargetRecursive(self, psContext, key, wasPressed)

    If Not result Is Nothing And key <> WKEY_NONE Then
        If self.hitTest(inputContext.mx, inputContext.my) Then
            Set psStrongParent = WeakWidget_lock(d.parent)
            If Not psStrongParent Is Nothing Then
                PopoverMenuWidget_handleMouseClickDownOnItem psStrongParent, self, key, wasPressed
            End If

            If wasPressed Then
                d.mouseDownOnWrapper = True
            Else
                ' if mouse down + mouse up on the item
                If d.mouseDownOnWrapper Then
                    d.mouseDownOnWrapper = False
                    Set psStrongParent = WeakWidget_lock(d.parent)
                    If Not psStrongParent Is Nothing Then
                        PopoverMenuWidget_handleMouseReleasedOnItem psStrongParent, self, key
                    End If
                End If
            End If
        End If
    End If

    Set MenuItemWrapper_findMouseTargetRecursive = result
End Function

Public Sub MenuItemWrapper_geometryChanged(ByVal self As WIDGET)
    MenuItemWrapper_(self).item.setGeometry 0, 0, self.width(), self.height()
End Sub

Public Sub MenuItemWrapper_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' no-op
End Sub

Public Function MenuItemWrapper_idealWidth(ByVal self As WIDGET) As Long
    MenuItemWrapper_idealWidth = MenuItemWrapper_(self).item.idealWidth()
End Function

Public Function MenuItemWrapper_idealHeight(ByVal self As WIDGET) As Long
    MenuItemWrapper_idealHeight = MenuItemWrapper_(self).item.idealHeight()
End Function

Public Sub MenuItemWrapper_clearMouseDownState(ByVal self As WIDGET)
    MenuItemWrapper_(self).mouseDownOnWrapper = False
End Sub

'--------------------------------------------------------------------------
' MARK: - PopoverMenuWidget

' PopoverMenuWidget::PopoverMenuWidget()
Public Sub PopoverMenuWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_PopoverMenuWidget ' (VB) virtual calls from here on reach PopoverMenuWidget's overrides
    PopoverMenuWidget_(self).weakSerial = WeakWidget_register(self)
    PopoverMenuWidget_(self).maxItemHeight = 0
End Sub

Public Sub PopoverMenuWidget_destruct(ByVal self As WIDGET)
    WeakWidget_unregister PopoverMenuWidget_(self).weakSerial
End Sub

Private Sub PopoverMenuWidget_initialize(ByVal self As WIDGET)
    Dim d As PopoverMenuWidget, p As Padding
    Set d = PopoverMenuWidget_(self)
    Set d.menuList = ScrollableListWidget_make()
    p.top = 5
    p.right_ = 5
    p.bottom = 5
    p.left_ = 5
    ScrollableListWidget_setPadding d.menuList, p
    ScrollableListWidget_setBackgroundColor d.menuList, pal_RGBA(20, 20, 20, 255)
    self.attach d.menuList
End Sub

Public Function PopoverMenuWidget_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_PopoverMenuWidget)
    PopoverMenuWidget_construct widget
    PopoverMenuWidget_initialize widget
    Set PopoverMenuWidget_make = widget
End Function

Public Sub PopoverMenuWidget_addMenuItem(ByVal self As WIDGET, ByVal widget As WIDGET, Optional ByVal closesMenuOnClick As Boolean = False)
    Dim d As PopoverMenuWidget, wrappedItem As WIDGET
    Set d = PopoverMenuWidget_(self)
    Set wrappedItem = MenuItemWrapper_wrap(self, widget, closesMenuOnClick)
    ScrollableListWidget_addItem d.menuList, wrappedItem
    d.maxItemHeight = MAX_(d.maxItemHeight, wrappedItem.idealHeight())
End Sub

Public Function PopoverMenuWidget_numItems(ByVal self As WIDGET) As Long
    PopoverMenuWidget_numItems = ScrollableListWidget_numItems(PopoverMenuWidget_(self).menuList)
End Function

Public Function PopoverMenuWidget_maxItemIdealHeight(ByVal self As WIDGET) As Long
    PopoverMenuWidget_maxItemIdealHeight = PopoverMenuWidget_(self).maxItemHeight
End Function

Public Function PopoverMenuWidget_getPadding(ByVal self As WIDGET) As Padding
    PopoverMenuWidget_getPadding = ScrollableListWidget_getPadding(PopoverMenuWidget_(self).menuList)
End Function

' std::shared_ptr<PopoverWidget> openMenu(const std::shared_ptr<WIDGET>& parent, PopoverWidget::Alignment align = PopoverWidget::Alignment::LeftOfParent, Vector2i positionOffset = Vector2i(0, 0), const PopoverWidget::OnPopoverCloseHandlerFunc& onCloseHandler = nullptr)
Public Function PopoverMenuWidget_openMenu(ByVal self As WIDGET, ByVal parent As WIDGET, Optional ByVal align As PopoverWidget_Alignment = PopoverWidget_Alignment_LeftOfParent, Optional ByVal positionOffsetX As Long = 0, Optional ByVal positionOffsetY As Long = 0, Optional ByVal onCloseHandler As WzFunc = Nothing) As WIDGET
    Dim d As PopoverMenuWidget, strongCurrentPopover As WIDGET, result As WIDGET
    Set d = PopoverMenuWidget_(self)
    Set strongCurrentPopover = WeakWidget_lock(d.currentPopover)
    If Not strongCurrentPopover Is Nothing Then
        PopoverWidget_close strongCurrentPopover
    End If
    ' Create a PopoverWidget, set this as its contents
    Set result = PopoverWidget_makePopover(parent, self, PopoverWidget_Style_Interactive, align, positionOffsetX, positionOffsetY, onCloseHandler)
    d.currentPopover = PopoverWidget_(result).weakSerial
    Set PopoverMenuWidget_openMenu = result
End Function

Public Sub PopoverMenuWidget_closeMenu(ByVal self As WIDGET)
    Dim d As PopoverMenuWidget, strongCurrentPopover As WIDGET
    Set d = PopoverMenuWidget_(self)
    Set strongCurrentPopover = WeakWidget_lock(d.currentPopover)
    If Not strongCurrentPopover Is Nothing Then
        PopoverWidget_close strongCurrentPopover
        d.currentPopover = 0    ' currentPopover.reset()
    End If
End Sub

Public Function PopoverMenuWidget_idealWidth(ByVal self As WIDGET) As Long
    PopoverMenuWidget_idealWidth = PopoverMenuWidget_(self).menuList.idealWidth()
End Function

Public Function PopoverMenuWidget_idealHeight(ByVal self As WIDGET) As Long
    PopoverMenuWidget_idealHeight = PopoverMenuWidget_(self).menuList.idealHeight()
End Function

Public Sub PopoverMenuWidget_geometryChanged(ByVal self As WIDGET)
    PopoverMenuWidget_(self).menuList.setGeometry 0, 0, self.width(), self.height()
End Sub

Private Sub PopoverMenuWidget_handleMouseClickDownOnItem(ByVal self As WIDGET, ByVal item As WIDGET, ByVal key As WIDGET_KEY, ByVal wasPressed As Boolean)
    Dim d As PopoverMenuWidget
    Set d = PopoverMenuWidget_(self)
    If item Is d.mouseDownItem Then Exit Sub

    If Not d.mouseDownItem Is Nothing Then
        MenuItemWrapper_clearMouseDownState d.mouseDownItem
    End If

    If wasPressed Then
        Set d.mouseDownItem = item
    Else
        Set d.mouseDownItem = Nothing
    End If
End Sub

Private Sub PopoverMenuWidget_handleMouseReleasedOnItem(ByVal self As WIDGET, ByVal item As WIDGET, ByVal key As WIDGET_KEY)
    Dim d As PopoverMenuWidget
    Set d = PopoverMenuWidget_(self)
    If Not item Is d.mouseDownItem Then
        If Not d.mouseDownItem Is Nothing Then
            MenuItemWrapper_clearMouseDownState d.mouseDownItem
            Set d.mouseDownItem = Nothing
        End If
        Exit Sub
    End If

    If Not item Is Nothing Then
        If MenuItemWrapper_getCloseMenuOnClick(item) Then
            PopoverMenuWidget_closeMenu self
        End If
    End If
    Set d.mouseDownItem = Nothing
End Sub
