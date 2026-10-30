Attribute VB_Name = "listwidget_c"
' lib/widget/listwidget.cpp, listwidget.h: enum class TabAlignment, struct
' TabSelectionStyle (a UDT; its constructor with arguments is
' TabSelectionStyle_), and the widget classes TabSelectionWidget, ListWidget and
' ListTabWidget (docs/CONVENTIONS.md, "Widgets"; their fields are
' TabSelectionWidget.cls, ListWidget.cls and ListTabWidget.cls). An enum nested
' in a class is named with the class: ListWidget::Order is ListWidget_Order
' (ListWidget_RightThenDown, ...), ListTabWidget::TabPosition is
' ListTabWidget_TabPosition (ListTabWidget_Top, ListTabWidget_Bottom). A
' std::vector of widgets or handlers is a Collection (element n is item n + 1).
' The handlers are WzFuncs: W_TABSELECTION_ON_TAB_CHANGED_FUNC (widget,
' currentTab), W_LISTWIDGET_ON_CURRENTPAGECHANGED_FUNC (widget, currentPage),
' W_LISTWIDGET_ON_NUMBEROFPAGESCHANGED_FUNC (widget, numberOfPages).
' Implementation of a list widget. Useful for displaying lists of buttons and stuff.
' (VB) widget class TabSelectionWidget : WIDGET
' (VB) virtual TabSelectionWidget.initialize
' (VB) widget class ListWidget : WIDGET
' (VB) widget class ListTabWidget : WIDGET
' (VB) virtual ListTabWidget.initialize
Option Explicit

Public Enum TabAlignment
    TabAlignment_LeftAligned
    TabAlignment_RightAligned
End Enum

Public Type TabSelectionStyle
    tabSize As WzSize
    scrollTabSize As WzSize
    tabImage As AtlasImage
    tabImageDown As AtlasImage
    tabImageHighlight As AtlasImage
    prevScrollTabImage As AtlasImage
    prevScrollTabImageDown As AtlasImage
    prevScrollTabImageHighlight As AtlasImage
    nextScrollTabImage As AtlasImage
    nextScrollTabImageDown As AtlasImage
    nextScrollTabImageHighlight As AtlasImage
    tabGap As Long
    tabAlignment_ As TabAlignment       ' tabAlignment = TabAlignment::LeftAligned
End Type

' enum Order {RightThenDown, DownThenRight}
Public Enum ListWidget_Order
    ListWidget_RightThenDown
    ListWidget_DownThenRight
End Enum

' enum TabPosition {Top, Bottom}
Public Enum ListTabWidget_TabPosition
    ListTabWidget_Top
    ListTabWidget_Bottom
End Enum

Private Const LABEL_X_BETWEEN_PADDING As Long = 5
Private Const LABEL_PADDING_Y As Long = 0

Public Const FN_TabSelectionWidget_prevTabPage As Long = 2030
Public Const FN_TabSelectionWidget_nextTabPage As Long = 2031
Public Const FN_TabSelectionWidget_setTabN As Long = 2032
Public Const FN_ListTabWidget_onTabChanged As Long = 2033
Public Const FN_ListTabWidget_onCurrentPageChanged As Long = 2034
Public Const FN_ListTabWidget_onNumberOfPagesChanged As Long = 2035

Public Function listwidget_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim pParent As WIDGET, w As WIDGET
    Set w = args(0)
    Select Case f.id
    Case FN_TabSelectionWidget_prevTabPage
        ' [](W_BUTTON& button)
        Set pParent = w.parent()
        ' assert(pParent != nullptr); (NDEBUG in a release build)
        TabSelectionWidget_prevTabPage pParent
    Case FN_TabSelectionWidget_nextTabPage
        Set pParent = w.parent()
        ' assert(pParent != nullptr); (NDEBUG in a release build)
        TabSelectionWidget_nextTabPage pParent
    Case FN_TabSelectionWidget_setTabN
        ' [n](W_BUTTON& button)
        Set pParent = w.parent()
        ' assert(pParent != nullptr); (NDEBUG in a release build)
        TabSelectionWidget_setTab pParent, f.cap(0)
    Case FN_ListTabWidget_onTabChanged
        ' [](TabSelectionWidget& tabsWidget, size_t currentTab)
        Set pParent = w.parent()
        If Not (Not pParent Is Nothing) Then ASSERT_FAILURE "pParent != nullptr", "lib/widget/listwidget.cpp:376", "initialize", "pParent is nullptr": Exit Function
        ListTabWidget_setCurrentPage pParent, args(1)
    Case FN_ListTabWidget_onCurrentPageChanged
        ' [](ListWidget& listWidget, size_t currentPage)
        Set pParent = w.parent()
        If Not (Not pParent Is Nothing) Then ASSERT_FAILURE "pParent != nullptr", "lib/widget/listwidget.cpp:381", "initialize", "pParent is nullptr": Exit Function
        TabSelectionWidget_setTab ListTabWidget_(pParent).tabs, args(1)
    Case FN_ListTabWidget_onNumberOfPagesChanged
        ' [](ListWidget& listWidget, size_t numberOfPages)
        Set pParent = w.parent()
        If Not (Not pParent Is Nothing) Then ASSERT_FAILURE "pParent != nullptr", "lib/widget/listwidget.cpp:386", "initialize", "pParent is nullptr": Exit Function
        TabSelectionWidget_setNumberOfTabs ListTabWidget_(pParent).tabs, args(1)
    End Select
End Function

' (VB) clip<size_t>(v, 0, hi): a negative v is a wrapped size_t, larger than any hi
Private Function clipSize(ByVal v As Long, ByVal hi As Long) As Long
    If v < 0 Then
        clipSize = hi
    ElseIf v > hi Then
        clipSize = hi
    Else
        clipSize = v
    End If
End Function

'--------------------------------------------------------------------------
' TabSelectionStyle

' TabSelectionStyle(AtlasImage tab, ..., int gap, TabAlignment tabAlignment = TabAlignment::LeftAligned)
Public Function TabSelectionStyle_(tab_ As AtlasImage, tabDown As AtlasImage, tabHighlight As AtlasImage, prev As AtlasImage, prevDown As AtlasImage, prevHighlight As AtlasImage, next_ As AtlasImage, nextDown As AtlasImage, nextHighlight As AtlasImage, ByVal gap As Long, Optional ByVal align_ As TabAlignment = TabAlignment_LeftAligned) As TabSelectionStyle
    Dim s As TabSelectionStyle
    s.tabSize = WzSize_(AtlasImage_width(tab_.images, tab_.id), AtlasImage_height(tab_.images, tab_.id))
    s.scrollTabSize = WzSize_(0, 0)
    s.tabImage = tab_
    s.tabImageDown = tabDown
    s.tabImageHighlight = tabHighlight
    s.prevScrollTabImage = prev
    s.prevScrollTabImageDown = prevDown
    s.prevScrollTabImageHighlight = prevHighlight
    s.nextScrollTabImage = next_
    s.nextScrollTabImageDown = nextDown
    s.nextScrollTabImageHighlight = nextHighlight
    s.tabGap = gap
    s.tabAlignment_ = align_
    If Not AtlasImage_isNull(prev) Then
        s.scrollTabSize = WzSize_(AtlasImage_width(prev.images, prev.id), AtlasImage_height(prev.images, prev.id))
    End If
    TabSelectionStyle_ = s
End Function

'--------------------------------------------------------------------------
' TabSelectionWidget

' static std::shared_ptr<TabSelectionWidget> make()
Public Function TabSelectionWidget_make() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_TabSelectionWidget)
    TabSelectionWidget_construct w
    vt_TabSelectionWidget_initialize w
    Set TabSelectionWidget_make = w
End Function

' TabSelectionWidget(): WIDGET(), currentTab(0), tabsAtOnce(1) {}
Public Sub TabSelectionWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_TabSelectionWidget ' (VB) virtual calls from here on reach TabSelectionWidget's overrides
    TabSelectionWidget_(self).currentTab = 0
    TabSelectionWidget_(self).tabsAtOnce = 1
End Sub

Public Sub TabSelectionWidget_initialize(ByVal self As WIDGET)
    Dim d As TabSelectionWidget
    Set d = TabSelectionWidget_(self)
    Set d.prevTabPageButton = W_BUTTON_new()
    self.attach d.prevTabPageButton
    Set d.nextTabPageButton = W_BUTTON_new()
    self.attach d.nextTabPageButton

    W_BUTTON_addOnClickHandler d.prevTabPageButton, WzFunc_new(FN_TabSelectionWidget_prevTabPage)
    W_BUTTON_addOnClickHandler d.nextTabPageButton, WzFunc_new(FN_TabSelectionWidget_nextTabPage)

    d.prevTabPageButton.setTip "Tab Scroll left"
    d.nextTabPageButton.setTip "Tab Scroll right"

    TabSelectionWidget_setHeight self, 5
    TabSelectionWidget_setNumberOfTabs self, 1
End Sub

Public Sub TabSelectionWidget_geometryChanged(ByVal self As WIDGET)
    TabSelectionWidget_doLayoutAll self
End Sub

Public Sub TabSelectionWidget_setHeight(ByVal self As WIDGET, ByVal height As Long)
    self.setGeometry self.x(), self.y(), self.width(), height
End Sub

Public Sub TabSelectionWidget_addStyle(ByVal self As WIDGET, tabStyle As TabSelectionStyle)
    TabSelectionWidget_(self).styles_push_back tabStyle

    TabSelectionWidget_doLayoutAll self
End Sub

Public Function TabSelectionWidget_tabs(ByVal self As WIDGET) As Long
    TabSelectionWidget_tabs = TabSelectionWidget_(self).tabButtons.Count
End Function

Public Sub TabSelectionWidget_addOnTabChangedHandler(ByVal self As WIDGET, ByVal onTabChangedFunc As WzFunc)
    TabSelectionWidget_(self).onTabChangedHandlers.Add onTabChangedFunc
End Sub

Public Sub TabSelectionWidget_setTab(ByVal self As WIDGET, ByVal tab_ As Long)
    Dim d As TabSelectionWidget, previousTab As Long, onTabChanged As Variant, i As Long
    Set d = TabSelectionWidget_(self)
    previousTab = d.currentTab
    d.currentTab = clipSize(tab_, d.tabButtons.Count - 1)
    If previousTab = d.currentTab Then
        Exit Sub  ' Nothing to do.
    End If
    TabSelectionWidget_doLayoutAll self

    ' Call all onTabChanged event handlers
    For i = 1 To d.onTabChangedHandlers.Count
        Set onTabChanged = d.onTabChangedHandlers(i)
        If Not onTabChanged Is Nothing Then
            WzFunc_call onTabChanged, self, d.currentTab
        End If
    Next
End Sub

Public Sub TabSelectionWidget_setNumberOfTabs(ByVal self As WIDGET, ByVal tabs As Long)
    Dim d As TabSelectionWidget, previousSize As Long, n As Long, button As WIDGET
    Set d = TabSelectionWidget_(self)
    tabs = MAX_(tabs, 1)
    previousSize = d.tabButtons.Count
    For n = tabs To d.tabButtons.Count - 1
        self.detach d.tabButtons(n + 1)
    Next
    ' tabButtons.resize(tabs)
    Do While d.tabButtons.Count > tabs
        d.tabButtons.Remove d.tabButtons.Count
    Loop
    For n = previousSize To tabs - 1
        Set button = W_BUTTON_new()
        d.tabButtons.Add button
        self.attach button
        W_BUTTON_addOnClickHandler button, WzFunc_new(FN_TabSelectionWidget_setTabN, n)
    Next

    TabSelectionWidget_doLayoutAll self
End Sub

Private Sub TabSelectionWidget_prevTabPage(ByVal self As WIDGET)
    Dim d As TabSelectionWidget
    Set d = TabSelectionWidget_(self)
    TabSelectionWidget_setTab self, d.currentTab - d.currentTab Mod d.tabsAtOnce - 1
End Sub

Private Sub TabSelectionWidget_nextTabPage(ByVal self As WIDGET)
    Dim d As TabSelectionWidget
    Set d = TabSelectionWidget_(self)
    TabSelectionWidget_setTab self, d.currentTab + (d.tabsAtOnce - d.currentTab Mod d.tabsAtOnce)
End Sub

Public Sub TabSelectionWidget_doLayoutAll(ByVal self As WIDGET)
    Dim d As TabSelectionWidget, tabStyle As TabSelectionStyle, scrollSpace As Long, n As Long, tabs As Long
    Dim haveScroll_ As Boolean, scrollSpace_ As Long, tabsAtOnce_ As Long, st As TabSelectionStyle
    Dim numPages As Long, numTabsOnLastPage As Long, x0 As Long, tabPage As Long, tabPos As Long, button As WIDGET
    Dim nullImage As AtlasImage
    Set d = TabSelectionWidget_(self)
    scrollSpace = 0
    tabs = TabSelectionWidget_tabs(self)
    d.tabsAtOnce = MAX_(tabs, 1)
    If tabs = 0 Or self.width() = 0 Or self.height() = 0 Then
        Exit Sub
    End If
    If tabs > 0 Then
        tabStyle.tabSize = WzSize_(self.width() \ tabs, self.height())
    Else
        tabStyle.tabSize = WzSize_(0, self.height())
    End If
    tabStyle.scrollTabSize = WzSize_(0, 0)
    tabStyle.tabGap = 0
    tabStyle.tabAlignment_ = TabAlignment_LeftAligned
    For n = 0 To d.stylesCount - 1
        st = d.styles(n)
        haveScroll_ = Not WzSize_isEmpty(st.scrollTabSize)
        If haveScroll_ Then
            scrollSpace_ = st.scrollTabSize.width_ + st.tabGap
        Else
            scrollSpace_ = 0
        End If
        tabsAtOnce_ = (self.width() + st.tabGap - scrollSpace_ * 2) \ (st.tabSize.width_ + st.tabGap)
        If haveScroll_ Or Not ULT(tabsAtOnce_, tabs) Then
            scrollSpace = scrollSpace_
            d.tabsAtOnce = tabsAtOnce_
            tabStyle = st
            Exit For
        End If
    Next
    d.prevTabPageButton.setGeometry 0, 0, tabStyle.scrollTabSize.width_, tabStyle.scrollTabSize.height_
    W_BUTTON_setImages_atlas d.prevTabPageButton, tabStyle.prevScrollTabImage, tabStyle.prevScrollTabImageDown, tabStyle.prevScrollTabImageHighlight, nullImage
    d.prevTabPageButton.show d.currentTab >= d.tabsAtOnce
    d.nextTabPageButton.setGeometry self.width() - tabStyle.scrollTabSize.width_, 0, tabStyle.scrollTabSize.width_, tabStyle.scrollTabSize.height_
    W_BUTTON_setImages_atlas d.nextTabPageButton, tabStyle.nextScrollTabImage, tabStyle.nextScrollTabImageDown, tabStyle.nextScrollTabImageHighlight, nullImage
    If d.tabsAtOnce > 0 Then
        d.nextTabPageButton.show d.currentTab \ d.tabsAtOnce < (tabs - 1) \ d.tabsAtOnce
    Else
        d.nextTabPageButton.show False
    End If
    If d.tabsAtOnce > 0 Then
        numPages = MAX_((tabs + (d.tabsAtOnce - 1)) \ d.tabsAtOnce, 1)
    Else
        numPages = 1
    End If
    If tabs > 0 Then
        numTabsOnLastPage = ((tabs - 1) Mod d.tabsAtOnce) + 1
    Else
        numTabsOnLastPage = 0
    End If
    For n = 0 To d.tabButtons.Count - 1
        x0 = 0
        tabPage = n \ d.tabsAtOnce
        tabPos = 0
        Select Case tabStyle.tabAlignment_
        Case TabAlignment_LeftAligned
            x0 = scrollSpace + (n Mod d.tabsAtOnce) * (tabStyle.tabSize.width_ + tabStyle.tabGap)
        Case TabAlignment_RightAligned
            If tabPage < (numPages - 1) Then
                tabPos = d.tabsAtOnce - (n Mod d.tabsAtOnce)
            Else
                ' on last page
                tabPos = numTabsOnLastPage - (n Mod d.tabsAtOnce)
            End If
            x0 = self.width() - scrollSpace - (tabPos * (tabStyle.tabSize.width_ + tabStyle.tabGap))
        End Select
        Set button = d.tabButtons(n + 1)
        button.setGeometry x0, 0, tabStyle.tabSize.width_, tabStyle.tabSize.height_
        W_BUTTON_setImages_atlas button, tabStyle.tabImage, tabStyle.tabImageDown, tabStyle.tabImageHighlight, nullImage
        button.show d.currentTab \ d.tabsAtOnce = n \ d.tabsAtOnce
        If n = d.currentTab Then
            button.setState WBUT_LOCK
        Else
            button.setState 0
        End If
    Next
    If d.tabButtons.Count = 1 Then
        Set button = d.tabButtons(1)
        button.hide                     ' Don't show single tabs.
    End If
End Sub

'--------------------------------------------------------------------------
' ListWidget

' std::make_shared<ListWidget>()
Public Function ListWidget_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ListWidget)
    ListWidget_construct w
    Set ListWidget_new = w
End Function

Public Sub ListWidget_construct(ByVal self As WIDGET)
    Dim d As ListWidget
    WIDGET_construct self
    self.kind = K_ListWidget ' (VB) virtual calls from here on reach ListWidget's overrides
    Set d = ListWidget_(self)
    d.numberOfPages = 1
    d.childSize = WzSize_(10, 10)
    d.spacing = WzSize_(4, 4)
    d.currentPage_ = 0
    d.order = ListWidget_RightThenDown
End Sub

Public Sub ListWidget_widgetLost(ByVal self As WIDGET, ByVal w As WIDGET)
    Dim d As ListWidget, i As Long
    Set d = ListWidget_(self)
    For i = 1 To d.myChildren.Count
        If d.myChildren(i) Is w Then
            d.myChildren.Remove i
            ListWidget_doLayoutAll self
            ListWidget_updateNumberOfPages self
            Exit For
        End If
    Next
End Sub

' Sets the size of all child widgets (applied by calling addWidgetToLayout).
Public Sub ListWidget_setChildSize(ByVal self As WIDGET, ByVal width As Long, ByVal height As Long)
    ListWidget_(self).childSize = WzSize_(width, height)
    ListWidget_doLayoutAll self
End Sub

' Sets the distance between child widgets (applied by calling addWidgetToLayout).
Public Sub ListWidget_setChildSpacing(ByVal self As WIDGET, ByVal width As Long, ByVal height As Long)
    ListWidget_(self).spacing = WzSize_(width, height)
    ListWidget_doLayoutAll self
End Sub

' Sets whether subsequent child widgets are placed in horizontal or vertical lines (applied by calling addWidgetToLayout).
Public Sub ListWidget_setOrder(ByVal self As WIDGET, ByVal order_ As ListWidget_Order)
    ListWidget_(self).order = order_
    ListWidget_doLayoutAll self
End Sub

' Manages the geometry of widget, and shows/hides it when changing tabs.
Public Sub ListWidget_addWidgetToLayout(ByVal self As WIDGET, ByVal widget As WIDGET)
    ListWidget_(self).myChildren.Add widget
    ListWidget_doLayout self, ListWidget_(self).myChildren.Count - 1
    ListWidget_updateNumberOfPages self
End Sub

Public Function ListWidget_currentPage(ByVal self As WIDGET) As Long
    ListWidget_currentPage = ListWidget_(self).currentPage_
End Function

Public Function ListWidget_pages(ByVal self As WIDGET) As Long
    ListWidget_pages = ListWidget_(self).numberOfPages
End Function

Public Function ListWidget_childrenSize(ByVal self As WIDGET) As Long
    ListWidget_childrenSize = ListWidget_(self).myChildren.Count
End Function

Public Sub ListWidget_setCurrentPage(ByVal self As WIDGET, ByVal page As Long)
    Dim d As ListWidget, previousPage As Long, pp As Long, n As Long, i As Long, onCurrentPageChanged As Variant, child As WIDGET
    Set d = ListWidget_(self)
    previousPage = d.currentPage_
    pp = ListWidget_widgetsPerPage(self)
    d.currentPage_ = clipSize(page, d.numberOfPages - 1)
    If previousPage = d.currentPage_ Then
        Exit Sub  ' Nothing to do.
    End If
    n = pp * previousPage
    Do While n < pp * (previousPage + 1) And n < d.myChildren.Count
        Set child = d.myChildren(n + 1)
        child.hide
        n = n + 1
    Loop
    n = pp * d.currentPage_
    Do While n < pp * (d.currentPage_ + 1) And n < d.myChildren.Count
        Set child = d.myChildren(n + 1)
        child.show
        n = n + 1
    Loop

    ' Call all onCurrentPageChanged event handlers
    For i = 1 To d.onCurrentPageChangedHandlers.Count
        Set onCurrentPageChanged = d.onCurrentPageChangedHandlers(i)
        If Not onCurrentPageChanged Is Nothing Then
            WzFunc_call onCurrentPageChanged, self, d.currentPage_
        End If
    Next
End Sub

Public Function ListWidget_widgetsPerPage(ByVal self As WIDGET) As Long
    ListWidget_widgetsPerPage = ListWidget_widgetsPerRow(self) * ListWidget_widgetsPerColumn(self)
End Function

Private Sub ListWidget_doLayoutAll(ByVal self As WIDGET)
    Dim n As Long
    For n = 0 To ListWidget_(self).myChildren.Count - 1
        ListWidget_doLayout self, n
    Next
End Sub

Private Sub ListWidget_doLayout(ByVal self As WIDGET, ByVal num As Long)
    Dim d As ListWidget, page As Long, withinPage As Long, column As Long, row As Long, child As WIDGET
    Dim childSize As WzSize
    Set d = ListWidget_(self)
    page = num \ ListWidget_widgetsPerPage(self)
    withinPage = num Mod ListWidget_widgetsPerPage(self)
    column = 0
    row = 0
    Select Case d.order
    Case ListWidget_RightThenDown
        column = withinPage Mod ListWidget_widgetsPerRow(self)
        row = withinPage \ ListWidget_widgetsPerRow(self)
    Case ListWidget_DownThenRight
        column = withinPage \ ListWidget_widgetsPerColumn(self)
        row = withinPage Mod ListWidget_widgetsPerColumn(self)
    End Select
    childSize = d.childSize
    Set child = d.myChildren(num + 1)
    child.setGeometry column * ListWidget_widgetSkipX(self), row * ListWidget_widgetSkipY(self), childSize.width_, childSize.height_
    child.show page = d.currentPage_
End Sub

Private Function ListWidget_widgetsPerRow(ByVal self As WIDGET) As Long
    ListWidget_widgetsPerRow = MAX_((self.width() + ListWidget_(self).spacing.width_) \ ListWidget_widgetSkipX(self), 1)
End Function

Private Function ListWidget_widgetsPerColumn(ByVal self As WIDGET) As Long
    ListWidget_widgetsPerColumn = MAX_((self.height() + ListWidget_(self).spacing.height_) \ ListWidget_widgetSkipY(self), 1)
End Function

Private Function ListWidget_widgetSkipX(ByVal self As WIDGET) As Long
    ListWidget_widgetSkipX = ListWidget_(self).childSize.width_ + ListWidget_(self).spacing.width_
End Function

Private Function ListWidget_widgetSkipY(ByVal self As WIDGET) As Long
    ListWidget_widgetSkipY = ListWidget_(self).childSize.height_ + ListWidget_(self).spacing.height_
End Function

Public Function ListWidget_firstWidgetShownIndex(ByVal self As WIDGET) As Long
    ListWidget_firstWidgetShownIndex = ListWidget_(self).currentPage_ * ListWidget_widgetsPerPage(self)
End Function

Public Function ListWidget_lastWidgetShownIndex(ByVal self As WIDGET) As Long
    Dim a As Long, b As Long
    a = ListWidget_firstWidgetShownIndex(self) + ListWidget_widgetsPerPage(self) - 1
    b = ListWidget_(self).myChildren.Count - 1     ' (size_t: -1 is the largest value)
    If ULT(b, a) Then
        ListWidget_lastWidgetShownIndex = b
    Else
        ListWidget_lastWidgetShownIndex = a
    End If
End Function

Public Function ListWidget_getWidgetAtIndex(ByVal self As WIDGET, ByVal index As Long) As WIDGET
    Dim d As ListWidget
    Set d = ListWidget_(self)
    If Not (index >= 0 And index < d.myChildren.Count) Then ASSERT_FAILURE "index < myChildren.size()", "lib/widget/listwidget.cpp:327", "getWidgetAtIndex", "Invalid widget index: %zu", index: Exit Function
    Set ListWidget_getWidgetAtIndex = d.myChildren(index + 1)
End Function

Public Sub ListWidget_addOnCurrentPageChangedHandler(ByVal self As WIDGET, ByVal handlerFunc As WzFunc)
    ListWidget_(self).onCurrentPageChangedHandlers.Add handlerFunc
End Sub

Public Sub ListWidget_addOnNumberOfPagesChangedHandler(ByVal self As WIDGET, ByVal handlerFunc As WzFunc)
    ListWidget_(self).onNumberOfPagesChangedHandlers.Add handlerFunc
End Sub

Private Sub ListWidget_updateNumberOfPages(ByVal self As WIDGET)
    Dim d As ListWidget, newNumberOfPages As Long, i As Long, onNumberOfPagesChanged As Variant
    Set d = ListWidget_(self)
    If d.myChildren.Count = 0 Then
        newNumberOfPages = 1
    Else
        newNumberOfPages = ((d.myChildren.Count - 1) \ ListWidget_widgetsPerPage(self)) + 1
    End If

    If newNumberOfPages <> d.numberOfPages Then
        d.numberOfPages = newNumberOfPages
        ListWidget_setCurrentPage self, d.currentPage_
        ' Call all onNumberOfPagesChanged event handlers
        For i = 1 To d.onNumberOfPagesChangedHandlers.Count
            Set onNumberOfPagesChanged = d.onNumberOfPagesChangedHandlers(i)
            If Not onNumberOfPagesChanged Is Nothing Then
                WzFunc_call onNumberOfPagesChanged, self, d.numberOfPages
            End If
        Next
    End If
End Sub

Public Sub ListWidget_geometryChanged(ByVal self As WIDGET)
    ListWidget_doLayoutAll self
    ListWidget_updateNumberOfPages self
End Sub

'--------------------------------------------------------------------------
' ListTabWidget

' ListTabWidget(): WIDGET(), tabPos(Top) {}
Public Sub ListTabWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_ListTabWidget ' (VB) virtual calls from here on reach ListTabWidget's overrides
    ListTabWidget_(self).tabPos = ListTabWidget_Top
End Sub

Public Sub ListTabWidget_initialize(ByVal self As WIDGET)
    Dim d As ListTabWidget
    Set d = ListTabWidget_(self)
    Set d.label = W_LABEL_new()
    self.attach d.label
    W_LABEL_setFont d.label, font_small, psPalette(WZCOL_TEXT_MEDIUM)
    d.label.setTransparentToMouse True
    Set d.tabs = TabSelectionWidget_make()
    self.attach d.tabs
    Set d.widgets = ListWidget_new()
    self.attach d.widgets
    TabSelectionWidget_addOnTabChangedHandler d.tabs, WzFunc_new(FN_ListTabWidget_onTabChanged)
    ListWidget_addOnCurrentPageChangedHandler d.widgets, WzFunc_new(FN_ListTabWidget_onCurrentPageChanged)
    ListWidget_addOnNumberOfPagesChangedHandler d.widgets, WzFunc_new(FN_ListTabWidget_onNumberOfPagesChanged)
    TabSelectionWidget_setNumberOfTabs d.tabs, ListWidget_pages(d.widgets)
End Sub

Public Sub ListTabWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As ListTabWidget, label_width As Long, label_x0 As Long, label_y0 As Long, tabs_x0 As Long, tabs_width As Long
    Dim widgets_y0 As Long, label_tabs_height As Long
    Set d = ListTabWidget_(self)
    label_width = MAX_(d.label.idealWidth(), 0)
    label_x0 = 0
    label_y0 = LABEL_PADDING_Y
    tabs_x0 = label_x0 + label_width + LABEL_X_BETWEEN_PADDING
    tabs_width = self.width() - tabs_x0
    Select Case d.tabPos
    Case ListTabWidget_Top
        d.label.setGeometry label_x0, label_y0, label_width, d.label.height()
        d.tabs.setGeometry tabs_x0, 0, tabs_width, d.tabs.height()
        widgets_y0 = MAX_(d.label.idealHeight(), d.tabs.height())
        d.widgets.setGeometry 0, widgets_y0, self.width(), self.height() - widgets_y0
    Case ListTabWidget_Bottom
        label_tabs_height = MAX_(d.label.idealHeight() + (LABEL_PADDING_Y * 2), d.tabs.height())
        label_y0 = self.height() - label_tabs_height + LABEL_PADDING_Y
        d.label.setGeometry label_x0, label_y0, label_width, d.label.height()
        d.tabs.setGeometry tabs_x0, self.height() - label_tabs_height, tabs_width, d.tabs.height()
        d.widgets.setGeometry 0, 0, self.width(), self.height() - d.tabs.height()
    End Select
End Sub

Public Function ListTabWidget_heightOfTabsLabel(ByVal self As WIDGET) As Long
    Dim d As ListTabWidget, widgets_y0 As Long
    Set d = ListTabWidget_(self)
    widgets_y0 = MAX_(LABEL_PADDING_Y + d.label.idealHeight(), d.tabs.height())
    ListTabWidget_heightOfTabsLabel = widgets_y0
End Function

Public Sub ListTabWidget_setTitle(ByVal self As WIDGET, ByVal string_ As String)
    Dim d As ListTabWidget
    Set d = ListTabWidget_(self)
    d.label.setString string_
    d.label.setGeometry 0, 0, d.label.idealWidth(), d.label.idealHeight()
    self.geometryChanged
End Sub

' Sets the size of all child widgets (applied by calling addWidgetToLayout).
Public Sub ListTabWidget_setChildSize(ByVal self As WIDGET, ByVal width As Long, ByVal height As Long)
    ListWidget_setChildSize ListTabWidget_(self).widgets, width, height
End Sub

' Sets the distance between child widgets (applied by calling addWidgetToLayout).
Public Sub ListTabWidget_setChildSpacing(ByVal self As WIDGET, ByVal width As Long, ByVal height As Long)
    ListWidget_setChildSpacing ListTabWidget_(self).widgets, width, height
End Sub

' Sets whether subsequent child widgets are placed in horizontal or vertical lines (applied by calling addWidgetToLayout).
Public Sub ListTabWidget_setOrder(ByVal self As WIDGET, ByVal order As ListWidget_Order)
    ListWidget_setOrder ListTabWidget_(self).widgets, order
End Sub

' Manages the geometry of widget, and shows/hides it when changing tabs.
Public Sub ListTabWidget_addWidgetToLayout(ByVal self As WIDGET, ByVal widget As WIDGET)
    Dim d As ListTabWidget
    Set d = ListTabWidget_(self)
    If widget.parent() Is self Then
        self.detach widget
    End If

    d.widgets.attach widget
    ListWidget_addWidgetToLayout d.widgets, widget
End Sub

Public Function ListTabWidget_setCurrentPage(ByVal self As WIDGET, ByVal page As Long) As Boolean
    ListWidget_setCurrentPage ListTabWidget_(self).widgets, page
    ListTabWidget_setCurrentPage = ListWidget_currentPage(ListTabWidget_(self).widgets) = page
End Function

Public Function ListTabWidget_currentPage(ByVal self As WIDGET) As Long
    ListTabWidget_currentPage = ListWidget_currentPage(ListTabWidget_(self).widgets)
End Function

Public Function ListTabWidget_pages(ByVal self As WIDGET) As Long
    ListTabWidget_pages = ListWidget_pages(ListTabWidget_(self).widgets)
End Function

Public Function ListTabWidget_childrenSize(ByVal self As WIDGET) As Long
    ListTabWidget_childrenSize = ListWidget_childrenSize(ListTabWidget_(self).widgets)
End Function

Public Sub ListTabWidget_setTabPosition(ByVal self As WIDGET, ByVal pos As ListTabWidget_TabPosition)
    ListTabWidget_(self).tabPos = pos
    self.geometryChanged
End Sub

' TabSelectionWidget *tabWidget()
Public Function ListTabWidget_tabWidget(ByVal self As WIDGET) As WIDGET
    Set ListTabWidget_tabWidget = ListTabWidget_(self).tabs
End Function

' ListWidget *listWidget()
Public Function ListTabWidget_listWidget(ByVal self As WIDGET) As WIDGET
    Set ListTabWidget_listWidget = ListTabWidget_(self).widgets
End Function

Public Sub ListTabWidget_goToChildPage(ByVal self As WIDGET, ByVal childIndex As Long)
    ListTabWidget_setCurrentPage self, childIndex \ ListWidget_widgetsPerPage(ListTabWidget_(self).widgets)
End Sub
