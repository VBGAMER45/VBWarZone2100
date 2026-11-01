Attribute VB_Name = "guidescreen_c"
' src/screens/guidescreen.cpp, guidescreen.h, whole: the Game Guide, an overlay screen (struct WzGameGuideScreen :
' W_SCREEN) with a movable form (a nav bar, a sidebar of topics, the topic shown), its topics registry, and the
' loading of the topics from data/base/guidetopics (wz2100.guidetopic.v1 JSON). src/wzjsonlocalizedstring.cpp is
' modinfo_c's (WzJsonLocalizedString_from_json: the string picked when the topic loads; translations are not ported).
' (VB) The registry (class WzGuideTopicsRegistry, one shared_ptr guideTopicsRegistry) is this module's
' WzGuideTopicsRegistry_<method> procedures over regTopics / regLookup / regNonCategory (the forms' m_registry is
' always it, so they do not keep it); its unordered_map is a Collection keyed "k" & identifier (VB's keys ignore
' case; the topic identifiers are lowercase). WzGuideTopic and WzWrappedGuideTopic are the classes of those names;
' WzStatVisual and WzGuideVisual are the UDT WzGuideVisual.
' (VB) The activity database (ActivityManager::instance().getRecord(), SQLite) is not ported: getRecord() is nullptr,
' 4.7.0's own path when the database cannot be opened ("store the prefs for this run at least").
' (VB) W_SCREEN has no room for a subclass's fields: WzGameGuideScreen's are this module's guideScreen* variables, as
' only one guide screen lives at a time (psCurrentGameGuideScreen); its methods are WzGameGuideScreen_<method>(screen).
' The std::weak_ptr captures of the guide's widgets are their widgweak_vb serials (each registers in its construct and
' unregisters in its destruct); a weak pointer to the guide screen is its address, locked through
' psCurrentGameGuideScreen, and one to the options overlay screen through the nav bar that holds it.
' (VB) widget class WzGuideTopicsNavBar : WIDGET
' (VB) widget class WzGuideSidebarTopicButton : W_BUTTON
' (VB) widget class WzGuideSidebar : WIDGET
' (VB) widget class WzGuideForm : W_FORM
' (VB) widget class WzGuideLinkedTopicButton : W_BUTTON
' (VB) widget class WzGuideRetButtonView : WIDGET
' (VB) widget class WzGuideStatPIEView : DynamicIntFancyButton
' (VB) widget class W_ATLASIMAGE : WIDGET
' (VB) widget class W_IMAGE : WIDGET
' (VB) widget class WzGuideLinkedTopicsWidget : WIDGET
' (VB) widget class WzGuideTopicDisplayWidget : WIDGET
' (VB) widget class WzGuideTopicsNavButton : W_BUTTON
' (VB) widget class WzGameGuideScreen_CLICKFORM : W_CLICKFORM
Option Explicit

' guidescreen.h: enum class ShowTopicFlags: int
Public Const ShowTopicFlags_None As Long = 0
Public Const ShowTopicFlags_FirstAdd As Long = 1      ' 1 << 0: open guide only if this topic is newly-added (this playthrough)
Public Const ShowTopicFlags_NeverViewed As Long = 2   ' 1 << 1: open guide only if this topic has never been viewed by the player before (in any playthrough)

' WzFunc ids (fork S3: 18400-18599)
Public Const FN_guide_linkedTopicClicked As Long = 18400        ' linkedTopicWidg->addOnClickHandler([topicIdentifierCopy, psWeakLinkedTopicsWidg](W_BUTTON&))
Public Const FN_guide_topicDisplay_onClickLinked As Long = 18401 ' WzGuideLinkedTopicsWidget::make(..., [psWeakGuideTopicWidg](const std::string&))
Public Const FN_guide_topicDisplay_showTask As Long = 18402     ' its widgScheduleTask([linkedTopicIdentifierCopy, psWeakGuideTopicWidg]())
Public Const FN_guide_navbar_close As Long = 18403              ' m_closeButton->addOnClickHandler([psWeakNavBar](W_BUTTON&))
Public Const FN_guide_navbar_options As Long = 18404            ' m_optionsButton->addOnClickHandler([psWeakNavBar](W_BUTTON& button))
Public Const FN_guide_navbar_sidebar As Long = 18405            ' m_sidebarButton->addOnClickHandler([psWeakNavBar](W_BUTTON&))
Public Const FN_guide_navbar_back As Long = 18406               ' m_backButton->addOnClickHandler([psWeakNavBar](W_BUTTON&))
Public Const FN_guide_navbar_next As Long = 18407               ' m_nextButton->addOnClickHandler([psWeakNavBar](W_BUTTON&))
Public Const FN_guide_optionsCheckbox_click As Long = 18408     ' createOptionsCheckbox's pCheckbox->addOnClickHandler([onClickFunc](W_BUTTON& button))
Public Const FN_guide_disablePopups_onClick As Long = 18409     ' [](WzCheckboxButton& button){ disableTopicPopups = ... }
Public Const FN_guide_optionsOverlay_clicked As Long = 18410    ' newRootFrm->onClickedFunc = [psWeakOptionsOverlayScreen, psWeakNavBar]()
Public Const FN_guide_optionsPopOver_calcLayout As Long = 18411 ' optionsPopOver->setCalcLayout([weakParent](WIDGET *psWidget))
Public Const FN_guide_sidebarButton_click As Long = 18412       ' button->addOnClickHandler([weakParent](W_BUTTON& but))
Public Const FN_guide_sidebarButton_task As Long = 18413        ' its widgScheduleTask([weakParent, topicIdentifierCopy]())
Public Const FN_guide_topicsList_calcLayout As Long = 18414     ' m_topicsList->setCalcLayout([](WIDGET* psWidget))
Public Const FN_guide_form_popState As Long = 18415             ' WzGuideForm::initialize's nav pop state handler [psWeakGuideForm](const std::string&)
Public Const FN_guide_form_popStateTask As Long = 18416         ' its widgScheduleTask([psWeakGuideForm, topicIdentifierCopy])
Public Const FN_guide_form_sidebarToggle As Long = 18417        ' on sidebar toggle handler [psWeakGuideForm]()
Public Const FN_guide_form_sidebarToggleTask As Long = 18418    ' its widgScheduleTask([psWeakGuideForm])
Public Const FN_guide_form_close As Long = 18419                ' close handler [psWeakGuideForm]()
Public Const FN_guide_form_closeTask As Long = 18420            ' its widgScheduleTask([psWeakGuideForm])
Public Const FN_guide_topic_onDelete As Long = 18421            ' m_guideTopic->setOnDelete([onTopicClose, topicIdentifierCopy](WIDGET *))
Public Const FN_guide_topic_onDeleteTask As Long = 18422        ' its widgScheduleTask([onTopicClose, topicIdentifierCopy]())
Public Const FN_guide_clickForm_calcLayout As Long = 18423      ' WzGameGuideScreen_CLICKFORM::make's LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_guide_clickForm_guideFormClose As Long = 18424  ' WzGuideForm::make(..., [weakSelf]()): on guide form close
Public Const FN_guide_clickForm_guideFormCloseTask As Long = 18425 ' its widgScheduleTask([weakSelf]())
Public Const FN_guide_screen_onClicked As Long = 18426          ' newRootFrm->onClickedFunc = [psWeakOverlayScreen]()
Public Const FN_guide_screen_onCancelPressed As Long = 18427    ' newRootFrm->onCancelPressed = [psWeakOverlayScreen]()
Public Const FN_guide_noopOnClose As Long = 18428               ' showGuideTopicImmediately's [](){ // no-op on screen close function }

Private Const GuideNavBarButtonHorizontalPadding As Long = 5   ' constexpr int

' WzStatVisual::OnlyIfKnown
Private Const OnlyIfKnown_False As Long = 0
Private Const OnlyIfKnown_CampaignMode As Long = 1
Private Const OnlyIfKnown_True As Long = 2

' struct WzGuideVisual (with struct WzStatVisual statVisual's fields)
Public Type WzGuideVisual
    has_imagePath As Boolean                    ' optional<std::string> imagePath
    imagePath As String
    has_statVisual As Boolean                   ' optional<WzStatVisual> statVisual
    statName As String                          ' WzStatVisual::statName
    onlyIfKnown As Long                         ' WzStatVisual::onlyIfKnown = OnlyIfKnown::False
    has_reticuleButtonId As Boolean             ' optional<int> reticuleButtonId
    reticuleButtonId As Long
    has_secondaryUnitOrderButtons As Boolean    ' optional<SECONDARY_ORDER> secondaryUnitOrderButtons
    secondaryUnitOrderButtons As Long
    droidTemplateJSON As nlohmann_json          ' optional<nlohmann::json> (Nothing: nullopt)
End Type

' WzGuideSidebarTopicButton's static constexpr ints
Private Const leftPadding As Long = 12
Private Const indentPadding As Long = 12
Private Const rightPadding As Long = 10
' WzGuideSidebar::listTopPadding (const uint32_t)
Private Const listTopPadding As Long = 10
' WzGuideForm's const int32_t topicsDisplayWidth, sidebarDisplayWidth
Private Const topicsDisplayWidth As Long = 400
Private Const sidebarDisplayWidth As Long = 220
' WzGuideLinkedTopicsWidget::lineSpacing (const int)
Private Const lineSpacing As Long = 2
' WzGuideTopicDisplayWidget's const ints
Private Const internalPaddingY As Long = 15
Private Const internalPaddingX As Long = 15
Private Const internalMarginTop As Long = 10
Private Const maxVisualHeight As Long = OBJ_BUTHEIGHT      ' for now, limited to the max height of a stat view button image
Private Const visualBetweenPaddingY As Long = 5

Private Const NUMRETBUTS As Long = 7                        ' Number of reticule buttons - from hci.cpp

' MARK: - Globals

Private psCurrentGameGuideScreen As W_SCREEN                ' static std::shared_ptr<WzGameGuideScreen>
Private disableTopicPopups As Boolean                       ' static bool = false

' the registry (static std::shared_ptr<WzGuideTopicsRegistry> guideTopicsRegistry)
Private regTopics As Collection                             ' std::vector<std::shared_ptr<WzWrappedGuideTopic>> topics
Private regLookup As Collection                             ' std::unordered_map<std::string, ...> identifierLookupMap
Private regNonCategory As Collection                        ' std::unordered_set<std::string> nonCategoryTopicIds

' struct WzGameGuideScreen's own fields (the screen psCurrentGameGuideScreen)
Private guideScreenOnCloseFunc As WzFunc                    ' OnCloseFunc onCloseFunc
Private guideScreenOpenTime As Double                       ' std::chrono::steady_clock::time_point openTime (widgSteadyNowMs)
Private guideScreenDidStopGameTime As Boolean               ' = false
Private guideScreenDidSetPauseStates As Boolean             ' = false

'--------------------------------------------------------------------------
' (VB) helpers

' a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' nlohmann's exceptions (json_hpp's json_throw: Err.Number vbObjectError + id, what() as the description)
Private Sub guideJsonThrow(ByVal kind As String, ByVal id As Long, ByVal what_arg As String)
    Err.Raise vbObjectError + id, "nlohmann::json", "[json.exception." & kind & "." & id & "] " & what_arg
End Sub

' std::weak_ptr<WzGameGuideScreen>::lock() of the screen at address p
Private Function lockGuideScreen(ByVal p As Long) As W_SCREEN
    If psCurrentGameGuideScreen Is Nothing Then Exit Function
    If ObjPtr(psCurrentGameGuideScreen) <> p Then Exit Function
    Set lockGuideScreen = psCurrentGameGuideScreen
End Function

'--------------------------------------------------------------------------
' this module's lambdas (WzFunc ids FN_ above): args are the C++ call's arguments (a W_BUTTON onClick: the button;
' a calcLayout or an onDelete: the widget; a std::function<void (const std::string&)>: the string; a task: none)
Public Function guidescreen_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongParent As WIDGET, strongNavBar As WIDGET, strongForm As WIDGET, psWidget As WIDGET, button As WIDGET
    Dim psParent As WIDGET, psOverlayScreen As W_SCREEN, popOverX0 As Long, popOverY0 As Long, topicIdentifierCopy As String
    Dim parentScreen As W_SCREEN, d As WzGuideTopicsNavBar

    Select Case f.id
    Case FN_guide_linkedTopicClicked
        Set strongParent = WeakWidget_lock(f.cap(1))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/screens/guidescreen.cpp:1102", "initialize", "No parent?": Exit Function
        If Not WzGuideLinkedTopicsWidget_(strongParent).m_onClickLinkedTopicHandler Is Nothing Then
            WzFunc_call WzGuideLinkedTopicsWidget_(strongParent).m_onClickLinkedTopicHandler, CStr(f.cap(0))
        End If

    Case FN_guide_topicDisplay_onClickLinked
        topicIdentifierCopy = args(0)
        widgScheduleTask WzFunc_new(FN_guide_topicDisplay_showTask, topicIdentifierCopy, f.cap(0))
    Case FN_guide_topicDisplay_showTask
        Set strongParent = WeakWidget_lock(f.cap(1))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongGuideTopicWidg != nullptr", "src/screens/guidescreen.cpp:1466", "initialize", "No parent?": Exit Function
        ' The parent of the strongGuideTopicWidg should be the WzGuideForm
        Set strongForm = strongParent.parent()
        If Not strongForm Is Nothing Then
            If Not WIDGET_isKindOf(strongForm, K_WzGuideForm) Then Set strongForm = Nothing
        End If
        If Not (Not strongForm Is Nothing) Then ASSERT_FAILURE "strongGuideForm != nullptr", "src/screens/guidescreen.cpp:1469", "initialize", "No guide form?": Exit Function
        WzGuideForm_showGuideTopic strongForm, CStr(f.cap(0))

    Case FN_guide_navbar_close
        Set strongNavBar = WeakWidget_lock(f.cap(0))
        If Not (Not strongNavBar Is Nothing) Then ASSERT_FAILURE "psStrongNavbar != nullptr", "src/screens/guidescreen.cpp:1674", "initialize", "No parent?": Exit Function
        WzGuideTopicsNavBar_triggerCloseHandler strongNavBar
    Case FN_guide_navbar_options
        Set strongNavBar = WeakWidget_lock(f.cap(0))
        If Not (Not strongNavBar Is Nothing) Then ASSERT_FAILURE "psStrongNavbar != nullptr", "src/screens/guidescreen.cpp:1684", "initialize", "No parent?": Exit Function
        ' Display a "pop-over" options menu
        WzGuideTopicsNavBar_displayOptionsOverlay strongNavBar, strongNavBar
        WzGuideTopicsNavBar_(strongNavBar).m_optionsButton.setState WBUT_CLICKLOCK
    Case FN_guide_navbar_sidebar
        Set strongNavBar = WeakWidget_lock(f.cap(0))
        If Not (Not strongNavBar Is Nothing) Then ASSERT_FAILURE "psStrongNavbar != nullptr", "src/screens/guidescreen.cpp:1702", "initialize", "No parent?": Exit Function
        WzGuideTopicsNavBar_triggerSidebarToggleHandler strongNavBar
    Case FN_guide_navbar_back
        Set strongNavBar = WeakWidget_lock(f.cap(0))
        If Not (Not strongNavBar Is Nothing) Then ASSERT_FAILURE "psStrongNavbar != nullptr", "src/screens/guidescreen.cpp:1714", "initialize", "No parent?": Exit Function
        WzGuideTopicsNavBar_navBack strongNavBar
    Case FN_guide_navbar_next
        Set strongNavBar = WeakWidget_lock(f.cap(0))
        If Not (Not strongNavBar Is Nothing) Then ASSERT_FAILURE "psStrongNavbar != nullptr", "src/screens/guidescreen.cpp:1723", "initialize", "No parent?": Exit Function
        WzGuideTopicsNavBar_navForward strongNavBar

    Case FN_guide_optionsCheckbox_click
        Set button = args(0)
        If Not WIDGET_isKindOf(button, K_WzCheckboxButton) Then Set button = Nothing
        If Not (Not button Is Nothing) Then ASSERT_FAILURE "checkBoxButton != nullptr", "src/screens/guidescreen.cpp:1870", "createOptionsPopoverForm", "checkBoxButton is null": Exit Function
        WzFunc_call f.cap(0), button
    Case FN_guide_disablePopups_onClick
        Set button = args(0)
        disableTopicPopups = WzCheckboxButton_getIsChecked(button)

    Case FN_guide_optionsOverlay_clicked
        ' (VB) psWeakOptionsOverlayScreen.lock(): the overlay screen the nav bar holds (its destructor removes it)
        Set strongNavBar = WeakWidget_lock(f.cap(1))
        If Not strongNavBar Is Nothing Then
            Set psOverlayScreen = WzGuideTopicsNavBar_(strongNavBar).optionsOverlayScreen
            If Not psOverlayScreen Is Nothing Then
                If ObjPtr(psOverlayScreen) = f.cap(0) Then widgRemoveOverlayScreen psOverlayScreen
            End If
        End If
        ' Destroy Options overlay / overlay screen, reset options button state
        If Not strongNavBar Is Nothing Then
            Set d = WzGuideTopicsNavBar_(strongNavBar)
            Set d.optionsOverlayScreen = Nothing
            d.m_optionsButton.setState 0
        End If
    Case FN_guide_optionsPopOver_calcLayout
        Set psWidget = args(0)
        Set psParent = WeakWidget_lock(f.cap(0))
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/screens/guidescreen.cpp:1940", "displayOptionsOverlay", "parent is null": Exit Function
        ' (Ideally, with its right aligned with the right side of the "parent" widget, but ensure full visibility on-screen)
        popOverX0 = (psParent.screenPosX() + psParent.width()) - psWidget.width() - 4
        If popOverX0 + psWidget.width() > screenWidth Then
            popOverX0 = screenWidth - psWidget.width()
        End If
        ' (Ideally, with its top aligned with the bottom of the "parent" widget, but ensure full visibility on-screen)
        popOverY0 = psParent.screenPosY() + psParent.height() - 4
        If popOverY0 + psWidget.height() > screenHeight Then
            popOverY0 = screenHeight - psWidget.height()
        End If
        psWidget.move popOverX0, popOverY0

    Case FN_guide_sidebarButton_click
        Set button = args(0)
        If Not WIDGET_isKindOf(button, K_WzGuideSidebarTopicButton) Then Set button = Nothing
        If Not (Not button Is Nothing) Then ASSERT_FAILURE "psSelf != nullptr", "src/screens/guidescreen.cpp:2142", "makeTopicWidgets", "Unexpected type?": Exit Function
        If Not WzGuideSidebarTopicButton_(button).isSelected Then
            topicIdentifierCopy = WzGuideSidebarTopicButton_(button).topicIdentifier
            widgScheduleTask WzFunc_new(FN_guide_sidebarButton_task, f.cap(0), topicIdentifierCopy)
        End If
    Case FN_guide_sidebarButton_task
        Set strongParent = WeakWidget_lock(f.cap(0))
        ' (VB) C++ dereferences the locks unchecked: gone, the task does nothing here
        If strongParent Is Nothing Then Exit Function
        Set strongForm = WeakWidget_lock(WzGuideSidebar_(strongParent).m_parentGuideForm)
        If strongForm Is Nothing Then Exit Function
        WzGuideForm_showGuideTopic strongForm, CStr(f.cap(1))
    Case FN_guide_topicsList_calcLayout
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not psParent Is Nothing Then
            If Not WIDGET_isKindOf(psParent, K_WzGuideSidebar) Then Set psParent = Nothing
        End If
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/screens/guidescreen.cpp:2207", "initialize", "No parent?": Exit Function
        psWidget.setGeometry 0, 0, MAX_(psParent.width() - 1, 0), MAX_(psParent.height() - 1, 0)

    Case FN_guide_form_popState
        ' nav pop state handler
        topicIdentifierCopy = args(0)
        widgScheduleTask WzFunc_new(FN_guide_form_popStateTask, f.cap(0), topicIdentifierCopy)
    Case FN_guide_form_popStateTask
        Set strongForm = WeakWidget_lock(f.cap(0))
        If Not (Not strongForm Is Nothing) Then ASSERT_FAILURE "psStrongGuideForm != nullptr", "src/screens/guidescreen.cpp:2272", "initialize", "No parent??": Exit Function
        WzGuideForm_displayGuideTopic strongForm, CStr(f.cap(1)), Nothing
    Case FN_guide_form_sidebarToggle
        ' on sidebar toggle handler
        widgScheduleTask WzFunc_new(FN_guide_form_sidebarToggleTask, f.cap(0))
    Case FN_guide_form_sidebarToggleTask
        Set strongForm = WeakWidget_lock(f.cap(0))
        If Not (Not strongForm Is Nothing) Then ASSERT_FAILURE "psStrongGuideForm != nullptr", "src/screens/guidescreen.cpp:2279", "initialize", "No parent??": Exit Function
        WzGuideForm_toggleSidebar strongForm
    Case FN_guide_form_close
        ' close handler
        widgScheduleTask WzFunc_new(FN_guide_form_closeTask, f.cap(0))
    Case FN_guide_form_closeTask
        Set strongForm = WeakWidget_lock(f.cap(0))
        If Not (Not strongForm Is Nothing) Then ASSERT_FAILURE "psStrongGuideForm != nullptr", "src/screens/guidescreen.cpp:2286", "initialize", "No parent??": Exit Function
        WzFunc_call WzGuideForm_(strongForm).m_closeHandler

    Case FN_guide_topic_onDelete
        widgScheduleTask WzFunc_new(FN_guide_topic_onDeleteTask, f.cap(0), f.cap(1))
    Case FN_guide_topic_onDeleteTask
        WzFunc_call f.cap(0), CStr(f.cap(1))

    Case FN_guide_clickForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight
    Case FN_guide_clickForm_guideFormClose
        ' on guide form close
        widgScheduleTask WzFunc_new(FN_guide_clickForm_guideFormCloseTask, f.cap(0))
    Case FN_guide_clickForm_guideFormCloseTask
        ' click on the form so we ultimately close the chat screen
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParentForm != nullptr", "src/screens/guidescreen.cpp:2595", "initialize", "No parent form": Exit Function
        Set parentScreen = strongParent.screenPointer
        If Not parentScreen Is Nothing Then Set parentScreen = lockGuideScreen(ObjPtr(parentScreen)) ' dynamic_pointer_cast<WzGameGuideScreen>
        If Not (Not parentScreen Is Nothing) Then ASSERT_FAILURE "parentScreen != nullptr", "src/screens/guidescreen.cpp:2597", "initialize", "No screen?": Exit Function
        WzGameGuideScreen_closeScreen parentScreen, True

    Case FN_guide_screen_onClicked
        Set psOverlayScreen = lockGuideScreen(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            WzGameGuideScreen_closeScreen psOverlayScreen, False
        End If
    Case FN_guide_screen_onCancelPressed
        Set psOverlayScreen = lockGuideScreen(f.cap(0))
        If Not psOverlayScreen Is Nothing Then
            WzGameGuideScreen_closeScreen psOverlayScreen, True
        End If
    Case FN_guide_noopOnClose
        ' no-op on screen close function
    End Select
End Function

'--------------------------------------------------------------------------
' MARK: - WzGuideTopicsRegistry

Private Sub regInit()
    If regTopics Is Nothing Then
        Set regTopics = New Collection
        Set regLookup = New Collection
        Set regNonCategory = New Collection
    End If
End Sub

' identifierLookupMap.find(identifier)'s value, Nothing for end()
Private Function WzGuideTopicsRegistry_getTopic(ByVal identifier As String) As WzWrappedGuideTopic
    regInit
    On Error Resume Next
    Set WzGuideTopicsRegistry_getTopic = regLookup.Item("k" & identifier)
    Err.Clear
End Function

' optional<std::string> getDisplayNameForTopic(const std::string& identifier) const: has_value, the value in result
Private Function WzGuideTopicsRegistry_getDisplayNameForTopic(ByVal identifier As String, result As String) As Boolean
    Dim it As WzWrappedGuideTopic
    Set it = WzGuideTopicsRegistry_getTopic(identifier)
    If it Is Nothing Then
        Exit Function                               ' nullopt
    End If
    result = it.topic.displayName                   ' getLocalizedString(WZ2100_GUIDE_MESSAGE_CATALOG_DOMAIN)
    WzGuideTopicsRegistry_getDisplayNameForTopic = True
End Function

Private Function WzGuideTopicsRegistry_getCategoryTopic(ByVal topicIdentifier As String) As WzWrappedGuideTopic
    Const topicDelimeter As String = "::"
    Dim lastComponentPos As Long, p As Long, possibleParent As String, parentIt As WzWrappedGuideTopic
    Dim categoryTopic As WzGuideTopic
    If Len(topicIdentifier) = 0 Then
        Exit Function
    End If
    regInit
    lastComponentPos = -1                           ' std::string::npos
    Do While lastComponentPos <> 0                  ' while (lastComponentPos > 0): npos is the largest size_t
        ' lastComponentPos = topicIdentifier.rfind(topicDelimeter, lastComponentPos - 1)
        If lastComponentPos = -1 Then
            p = InStrRev(topicIdentifier, topicDelimeter)
        Else
            p = InStrRev(topicIdentifier, topicDelimeter, MIN_(lastComponentPos - 1 + Len(topicDelimeter), Len(topicIdentifier)))
        End If
        lastComponentPos = p - 1
        If lastComponentPos = -1 Then
            Exit Do
        End If
        possibleParent = Left$(topicIdentifier, lastComponentPos)
        Set parentIt = WzGuideTopicsRegistry_getTopic(possibleParent)
        If Not parentIt Is Nothing Then
            debug_ 207, LOG_WZ, "getCategoryTopic", "Found parent: %s", possibleParent
            Set WzGuideTopicsRegistry_getCategoryTopic = parentIt
            Exit Function
        Else
            ' parent guide category topic not found
            If collectionHasKey(regNonCategory, possibleParent) Then
                ' already checked this - doesn't exist - continue
            Else
                ' did not already check for an _index.json underneath this category - try to load
                Set categoryTopic = loadWZGuideTopicByTopicId(possibleParent)
                If categoryTopic Is Nothing Then
                    debug_ 224, LOG_WZ, "getCategoryTopic", "No category: %s", possibleParent
                    regNonCategory.Add True, "k" & possibleParent
                Else
                    ' add the category topic to the topics list, and return it
                    debug_ 229, LOG_WZ, "getCategoryTopic", "Loaded category: %s", possibleParent
                    Set WzGuideTopicsRegistry_getCategoryTopic = WzGuideTopicsRegistry_addTopicImpl(categoryTopic)
                    Exit Function
                End If
            End If
        End If
    Loop
End Function

' (VB) whether the Collection has the key "k" & key
Private Function collectionHasKey(ByVal c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c.Item("k" & key)
    collectionHasKey = (Err.Number = 0)
    Err.Clear
End Function

Private Function WzGuideTopicsRegistry_addTopicImpl(ByVal topic As WzGuideTopic, Optional ByVal restoringLoadedTopics As Boolean = False) As WzWrappedGuideTopic
    Dim wrappedTopic As WzWrappedGuideTopic, topicIdentifier As String, parentTopic As WzWrappedGuideTopic
    If Not (Not topic Is Nothing) Then ASSERT_FAILURE "topic != nullptr", "src/screens/guidescreen.cpp:240", "addTopicImpl", "Null topic": Exit Function
    regInit

    If Not WzGuideTopicsRegistry_getTopic(topic.identifier) Is Nothing Then
        debug_ 244, LOG_WZ, "addTopicImpl", "Guide topic with identifier ""%s"" was already added", topic.identifier
        Exit Function
    End If

    If topic.has_applicableModes And topic.applicableModesCount > 0 And Not restoringLoadedTopics Then
        If (topic.applicableModes And guideGameModeBit(ActivityManager_getCurrentGameMode())) = 0 Then
            ' current game mode is not in applicableModes - skipping add of topic
            debug_ 253, LOG_WZ, "addTopicImpl", "Skipping topic (not applicable to current game mode): %s", topic.identifier
            Exit Function
        End If
    End If

    Set wrappedTopic = New WzWrappedGuideTopic
    Set wrappedTopic.topic = topic

    ' Determine if it should be inserted underneath a category
    topicIdentifier = topic.identifier
    Set parentTopic = WzGuideTopicsRegistry_getCategoryTopic(topicIdentifier)
    If Not parentTopic Is Nothing Then
        parentTopic.children.Add wrappedTopic
        wrappedTopic.parentId = parentTopic.topic.identifier
    Else
        ' root level
        regTopics.Add wrappedTopic
    End If

    regLookup.Add wrappedTopic, "k" & topic.identifier
    Set WzGuideTopicsRegistry_addTopicImpl = wrappedTopic
End Function

' (VB) a game mode's bit in WzGuideTopic::applicableModes
Private Function guideGameModeBit(ByVal mode As Long) As Long
    guideGameModeBit = 2 ^ mode
End Function

Private Sub WzGuideTopicsRegistry_clear()
    Set regTopics = New Collection
    Set regLookup = New Collection
    If regNonCategory Is Nothing Then Set regNonCategory = New Collection
End Sub

' enumerateTopicsImpl: the children first, then the topic (enumerateLoadedTopics' order); (VB) enumFunc is
' saveLoadedTopics' results.push_back
Private Sub WzGuideTopicsRegistry_enumerateTopicsImpl(ByVal root As Collection, results As StringVector)
    Dim wrappedTopic As WzWrappedGuideTopic
    For Each wrappedTopic In root
        If wrappedTopic.children.Count > 0 Then
            WzGuideTopicsRegistry_enumerateTopicsImpl wrappedTopic.children, results
        End If
        StringVector_push_back results, wrappedTopic.topic.identifier
    Next
End Sub

Private Function WzGuideTopicsRegistry_restoreLoadedTopics(topicIds() As String, ByVal n As Long) As Boolean
    Dim i As Long, loadedTopic As WzGuideTopic
    WzGuideTopicsRegistry_clear
    For i = 0 To n - 1
        Set loadedTopic = loadWZGuideTopicByTopicId(topicIds(i))
        If loadedTopic Is Nothing Then
            debug_ 323, LOG_INFO, "restoreLoadedTopics", "Failed to open / load guide topic: %s", topicIds(i)
        Else
            WzGuideTopicsRegistry_addTopicImpl loadedTopic, True
        End If
    Next
    WzGuideTopicsRegistry_restoreLoadedTopics = True
End Function

' optional<ActivityDBProtocol::GuideTopicPrefs> getTopicPrefs(topic): (VB) has_value; the prefs are the topic's fields
Private Function WzGuideTopicsRegistry_getTopicPrefs(ByVal topic As WzWrappedGuideTopic) As Boolean
    ' find cached data first
    If Not (Not topic Is Nothing) Then ASSERT_FAILURE "topic != nullptr", "src/screens/guidescreen.cpp:334", "getTopicPrefs", "Null topic": Exit Function
    If topic.has_prefs Then
        WzGuideTopicsRegistry_getTopicPrefs = True
        Exit Function
    End If
    ' (activityDB = ActivityManager::instance().getRecord(): nullptr, the activity database is not ported)
    ' database access isn't available, but store the prefs for this run at least
    topic.has_prefs = True                          ' topic->prefs = ActivityDBProtocol::GuideTopicPrefs()
    topic.prefs_has_lastReadVersion = False
    topic.prefs_bookmarked = False
    WzGuideTopicsRegistry_getTopicPrefs = True
End Function

Private Sub WzGuideTopicsRegistry_setTopicSeen(ByVal topic As WzWrappedGuideTopic)
    If Not (Not topic Is Nothing) Then ASSERT_FAILURE "topic != nullptr", "src/screens/guidescreen.cpp:360", "setTopicSeen", "Null topic": Exit Sub
    If Not topic.has_prefs Then
        ' populate it from db
        WzGuideTopicsRegistry_getTopicPrefs topic
        If Not (topic.has_prefs) Then ASSERT_FAILURE "topic->prefs.has_value()", "src/screens/guidescreen.cpp:365", "setTopicSeen", "Failed to initialize topic prefs?": Exit Sub
    End If

    If Not (topic.prefs_has_lastReadVersion And topic.prefs_lastReadVersion = topic.topic.version) Then
        topic.prefs_has_lastReadVersion = True
        topic.prefs_lastReadVersion = topic.topic.version
        ' (activityDB = getRecord(): nullptr, so nothing is stored: setGuideTopicPrefs is not reached)
    End If
End Sub

'--------------------------------------------------------------------------
' MARK: - Linked Topic button

' WzGuideLinkedTopicButton(): font(font_regular), cachedText("", font)
Private Sub WzGuideLinkedTopicButton_construct(ByVal self As WIDGET)
    Dim d As WzGuideLinkedTopicButton
    W_BUTTON_construct self
    self.kind = K_WzGuideLinkedTopicButton
    Set d = WzGuideLinkedTopicButton_(self)
    d.font = font_regular
    d.cachedText.init "", d.font, 100
    d.horizontalPadding = 3
    d.leftMargin = 0
    d.ptsAboveBase = 0
End Sub

Private Sub WzGuideLinkedTopicButton_initialize(ByVal self As WIDGET, ByVal localizedDisplayName As String)
    Dim d As WzGuideLinkedTopicButton
    Set d = WzGuideLinkedTopicButton_(self)
    d.displayName = localizedDisplayName
    d.cachedText.init d.displayName, d.font, 100    ' cachedText = WzCachedText(displayName, font)
    d.ptsAboveBase = d.cachedText.deref().aboveBase()
    self.setGeometry self.x(), self.y(), self.idealWidth(), self.idealHeight()
End Sub

Private Function WzGuideLinkedTopicButton_make(ByVal localizedDisplayName As String) As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideLinkedTopicButton)
    WzGuideLinkedTopicButton_construct widget
    WzGuideLinkedTopicButton_initialize widget, localizedDisplayName
    Set WzGuideLinkedTopicButton_make = widget
End Function

Public Sub WzGuideLinkedTopicButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzGuideLinkedTopicButton, left_ As Long, top As Long, marginLeft As Long, textX As Long, textY As Long
    Dim isDown As Boolean, isHighlight As Boolean, textColor As Long, backColor As Long, borderColor As Long
    Set d = WzGuideLinkedTopicButton_(self)
    left_ = xOffset + self.x()
    top = yOffset + self.y()
    marginLeft = left_ + d.leftMargin
    textX = marginLeft + d.horizontalPadding
    textY = top - d.cachedText.deref().aboveBase()

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isHighlight = (W_BUTTON_(self).state And WBUT_HIGHLIGHT) <> 0

    textColor = psPalette(WZCOL_FORM_TEXT)

    If isHighlight Or isDown Then
        If isDown Then backColor = pal_RGBA(50, 50, 50, 65) Else backColor = pal_RGBA(255, 255, 255, 65)
        pie_UniTransBoxFill left_, top, left_ + self.width() - 1, top + self.height() - 1, backColor

        If isHighlight Then borderColor = pal_RGBA(255, 255, 255, 180) Else borderColor = pal_RGBA(255, 255, 255, 75)
        iV_Box left_, top, left_ + self.width(), top + self.height(), borderColor

        textColor = psPalette(WZCOL_TEXT_BRIGHT)
    End If

    d.cachedText.deref().renderOutlined textX, textY, textColor, pal_RGBA(0, 0, 0, 128)
End Sub

Public Sub WzGuideLinkedTopicButton_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzGuideLinkedTopicButton_(self).cachedText.tick
End Sub

Public Function WzGuideLinkedTopicButton_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzGuideLinkedTopicButton
    Set d = WzGuideLinkedTopicButton_(self)
    WzGuideLinkedTopicButton_idealWidth = d.cachedText.deref().width() + d.leftMargin + 2 * d.horizontalPadding
End Function

Public Function WzGuideLinkedTopicButton_idealHeight(ByVal self As WIDGET) As Long
    WzGuideLinkedTopicButton_idealHeight = WzGuideLinkedTopicButton_(self).cachedText.deref().lineSize()
End Function

'--------------------------------------------------------------------------
' MARK: - Reticule Button image view

Private Function WzGuideRetButtonView_make(ByVal retButId As Long) As WIDGET
    Dim retButtonDisplayFilename As String, widget As WIDGET, d As WzGuideRetButtonView, image As Long
    If Not getReticuleButtonDisplayFilename(retButId, retButtonDisplayFilename) Then
        Exit Function
    End If
    If Len(retButtonDisplayFilename) = 0 Then
        Exit Function
    End If
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideRetButtonView)
    WIDGET_construct widget
    widget.kind = K_WzGuideRetButtonView
    Set d = WzGuideRetButtonView_(widget)
    d.retButtonDisplayFilename = retButtonDisplayFilename

    image = iV_GetImage(d.retButtonDisplayFilename)
    If image <> 0 Then
        ' set the button width/height based on the "normal" image dimensions
        d.imageWidth = AtlasImageDef_Width(image)
        d.imageHeight = AtlasImageDef_Height(image)
        widget.setGeometry widget.x(), widget.y(), d.imageWidth, d.imageHeight
    End If

    Set WzGuideRetButtonView_make = widget
End Function

Public Sub WzGuideRetButtonView_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzGuideRetButtonView, x0 As Long, y0 As Long, w As Single, h As Single
    Set d = WzGuideRetButtonView_(self)
    If Len(d.retButtonDisplayFilename) = 0 Then
        Exit Sub
    End If
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()
    iV_DrawImage2 d.retButtonDisplayFilename, x0, y0, w, h
End Sub

Public Function WzGuideRetButtonView_idealWidth(ByVal self As WIDGET) As Long
    WzGuideRetButtonView_idealWidth = WzGuideRetButtonView_(self).imageWidth
End Function

Public Function WzGuideRetButtonView_idealHeight(ByVal self As WIDGET) As Long
    WzGuideRetButtonView_idealHeight = WzGuideRetButtonView_(self).imageHeight
End Function

'--------------------------------------------------------------------------
' MARK: - Stat image view

' WzGuideStatPIEView() { setTransparentToClicks(true); }
Private Sub WzGuideStatPIEView_construct(ByVal self As WIDGET)
    DynamicIntFancyButton_construct self
    self.kind = K_WzGuideStatPIEView
    self.setTransparentToClicks True
End Sub

' (VB) ~WzGuideStatPIEView(): the DROID_TEMPLATE design it owns goes back to the pool
Public Sub WzGuideStatPIEView_destruct(ByVal self As WIDGET)
    Dim d As WzGuideStatPIEView
    Set d = WzGuideStatPIEView_(self)
    If d.design <> 0 Then
        deleteTemplate d.design
        d.design = 0
    End If
End Sub

Private Function getComponentsNumByType(ByVal typ As COMPONENT_TYPE) As Long
    Select Case typ
    Case COMP_BODY: getComponentsNumByType = asBodyStats_size
    Case COMP_BRAIN: getComponentsNumByType = asBrainStats_size
    Case COMP_PROPULSION: getComponentsNumByType = asPropulsionStats_size
    Case COMP_REPAIRUNIT: getComponentsNumByType = asRepairStats_size
    Case COMP_ECM: getComponentsNumByType = asECMStats_size
    Case COMP_SENSOR: getComponentsNumByType = asSensorStats_size
    Case COMP_CONSTRUCT: getComponentsNumByType = asConstructStats_size
    Case COMP_WEAPON: getComponentsNumByType = asWeaponStats_size
    Case COMP_NUMCOMPONENTS: getComponentsNumByType = 0
    End Select
End Function

' static optional<ItemAvailability> getStatPlayerAvailability(int player, BASE_STATS *Stat): has_value, the value in result
Private Function getStatPlayerAvailability(ByVal player As Long, ByVal Stat As Long, result As Long) As Boolean
    Dim partIndex As COMPONENT_TYPE, numOfComponentsOfType As Long, idx As Long
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/screens/guidescreen.cpp:748", "getStatPlayerAvailability", "Invalid player: %d", player: Exit Function
    If Not (Stat <> 0) Then ASSERT_FAILURE "Stat != nullptr", "src/screens/guidescreen.cpp:749", "getStatPlayerAvailability", "Stat is null": Exit Function

    idx = BASE_STATS_index(Stat)
    If StatIsStructure(Stat) Then
        If Not (idx < numStructureStats) Then ASSERT_FAILURE "Stat->index < numStructureStats", "src/screens/guidescreen.cpp:753", "getStatPlayerAvailability", "Invalid struct stat index: %zu (exceeds numStructureStats: %u)", idx, numStructureStats: Exit Function
        If Not (Stat = STAT_STRUCTURE + idx) Then ASSERT_FAILURE "Stat == &asStructureStats[Stat->index]", "src/screens/guidescreen.cpp:754", "getStatPlayerAvailability", "Sanity check failed": Exit Function
        result = apStructTypeLists(player, idx)
        getStatPlayerAvailability = True
    ElseIf StatIsTemplate(Stat) Then
        ' not currently supported
        Exit Function
    ElseIf StatIsFeature(Stat) Then
        ' treat features as always available
        result = AVAILABLE
        getStatPlayerAvailability = True
    Else
        partIndex = StatIsComponent(Stat) ' Old comment: "This fails for viper body." ... but does it?
        If partIndex <> COMP_NUMCOMPONENTS Then
            numOfComponentsOfType = getComponentsNumByType(partIndex)
            If Not (idx < numOfComponentsOfType) Then ASSERT_FAILURE "Stat->index < numOfComponentsOfType", "src/screens/guidescreen.cpp:773", "getStatPlayerAvailability", "Invalid Stat index: %zu (exceeds component type num: %zu)", idx, numOfComponentsOfType: Exit Function
            result = apCompLists(player, partIndex, idx)
            getStatPlayerAvailability = True
        ElseIf StatIsResearch(Stat) Then
            If Not (idx < asResearch_size) Then ASSERT_FAILURE "Stat->index < asResearch.size()", "src/screens/guidescreen.cpp:778", "getStatPlayerAvailability", "Invalid Stat index: %zu (exceeds research num: %zu)", idx, asResearch_size: Exit Function
            If (asPlayerResList(player, idx).ResearchStatus And RESBITS_ALL) <> 0 Then
                result = AVAILABLE
            ElseIf IsResearchPossible(asPlayerResList(player, idx)) Then
                result = FOUND
            Else
                result = UNAVAILABLE
            End If
            getStatPlayerAvailability = True
        End If
    End If
End Function

' static std::shared_ptr<WzGuideStatPIEView> make(const WzStatVisual& statVisual)
Private Function WzGuideStatPIEView_make(statVisual As WzGuideVisual) As WIDGET
    Dim Stat As Long, featureStatId As Long, availability As Long, widget As WIDGET
    Stat = getBaseStatsFromName(statVisual.statName)
    If Stat = 0 Then
        ' could be a feature stat id (which getBaseStatsFromName doesn't currently handle)
        featureStatId = getFeatureStatFromName(statVisual.statName)
        If featureStatId >= 0 And featureStatId < asFeatureStats_size Then
            Stat = STAT_FEATURE + featureStatId
        Else
            ' still didn't find this stat
            Exit Function
        End If
    End If

    If statVisual.onlyIfKnown = OnlyIfKnown_True Or (statVisual.onlyIfKnown = OnlyIfKnown_CampaignMode And ActivityManager_getCurrentGameMode() = ActivitySink_GameMode_CAMPAIGN) Then
        If getStatPlayerAvailability(selectedPlayer, Stat, availability) Then
            Select Case availability
            Case UNAVAILABLE
                ' skip displaying this item
                Exit Function
            End Select
        End If
    End If

    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideStatPIEView)
    WzGuideStatPIEView_construct widget
    IntFancyButton_(widget).buttonType = TOPBUTTON
    WzGuideStatPIEView_(widget).Stat = Stat
    Set WzGuideStatPIEView_make = widget
End Function

Private Function WzGuideStatPIEView_makeFromDroidTemplateJSON(ByVal droidTemplateJSON As nlohmann_json) As WIDGET
    Dim widget As WIDGET, d As WzGuideStatPIEView
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideStatPIEView)
    WzGuideStatPIEView_construct widget
    Set d = WzGuideStatPIEView_(widget)
    d.design = newTemplate()                        ' DROID_TEMPLATE design (the widget's own)
    If Not loadTemplateCommon(droidTemplateJSON, apsTemplates(d.design)) Then
        debug_ 845, LOG_ERROR, "makeFromDroidTemplateJSON", "Unable to parse template"
        Exit Function
    End If
    IntFancyButton_(widget).buttonType = TOPBUTTON
    d.Stat = STAT_TEMPLATE + d.design               ' point to the DROID_TEMPLATE that the widget owns
    Set WzGuideStatPIEView_makeFromDroidTemplateJSON = widget
End Function

Public Function WzGuideStatPIEView_isHighlighted(ByVal self As WIDGET) As Boolean
    WzGuideStatPIEView_isHighlighted = False
End Function

Public Function WzGuideStatPIEView_idealWidth(ByVal self As WIDGET) As Long
    WzGuideStatPIEView_idealWidth = OBJ_BUTWIDTH    ' because DynamicIntFancyButton largely assumes this, and doesn't currently support resizing
End Function

Public Function WzGuideStatPIEView_idealHeight(ByVal self As WIDGET) As Long
    WzGuideStatPIEView_idealHeight = OBJ_BUTHEIGHT  ' because DynamicIntFancyButton largely assumes this, and doesn't currently support resizing
End Function

Public Sub WzGuideStatPIEView_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim object_ As ImdObject, Stat As Long, compID As Long
    Stat = WzGuideStatPIEView_(self).Stat
    If Stat <> 0 Then
        If StatIsStructure(Stat) Then
            object_ = ImdObject_StructureStat(Stat)
        ElseIf StatIsTemplate(Stat) Then
            object_ = ImdObject_DroidTemplate(Stat)
        ElseIf StatIsFeature(Stat) Then
            object_ = ImdObject_Feature(Stat)
        Else
            compID = StatIsComponent(Stat) ' Old comment: "This fails for viper body." ... but does it?
            If compID <> COMP_NUMCOMPONENTS Then
                object_ = ImdObject_Component(Stat)
            ElseIf StatIsResearch(Stat) Then
                object_ = getResearchObjectImage(Stat And &HFFFF&)
            End If
        End If
    Else
        'BLANK button for now - AB 9/1/98
        object_ = ImdObject_Component(0)
    End If

    IntFancyButton_displayIMD self, AtlasImage_(), object_, xOffset, yOffset
End Sub

'--------------------------------------------------------------------------
' MARK: - Atlas IMAGEFILE image view

' static std::shared_ptr<W_ATLASIMAGE> make(IMAGEFILE *ImageFile, std::vector<UWORD> ImageIDs): ImageIDs(0 .. count - 1)
Private Function W_ATLASIMAGE_make(ByVal ImageFile As Long, ImageIDs() As Long, ByVal count As Long) As WIDGET
    Dim widget As WIDGET, d As W_ATLASIMAGE, maxImageHeight As Long, totalIdealWidth As Long, i As Long, w As Long, h As Long
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_W_ATLASIMAGE)
    WIDGET_construct widget
    widget.kind = K_W_ATLASIMAGE
    If Not (ImageFile <> 0) Then ASSERT_FAILURE "ImageFile != nullptr", "src/screens/guidescreen.cpp:915", "make", "ImageFile is null": Exit Function
    Set d = W_ATLASIMAGE_(widget)
    d.ImageFile = ImageFile
    For i = 0 To count - 1
        d.ImageIDs.Add ImageIDs(i)
    Next

    maxImageHeight = 0
    totalIdealWidth = 0
    For i = 0 To count - 1
        w = iV_GetImageWidth(ImageFile, ImageIDs(i))
        h = iV_GetImageHeight(ImageFile, ImageIDs(i))
        d.imageWidths.Add w
        d.imageHeights.Add h
        maxImageHeight = MAX_(maxImageHeight, h)
        totalIdealWidth = totalIdealWidth + w + 2
    Next
    d.maxImageHeight = maxImageHeight
    d.totalIdealWidth = totalIdealWidth

    Set W_ATLASIMAGE_make = widget
End Function

Public Sub W_ATLASIMAGE_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_ATLASIMAGE, x0 As Long, y0 As Long, h As Long, currImageX0 As Long, idx As Long
    Dim imageWidth As Long, imageHeight As Long, imageID As Long, imageY0 As Long
    Set d = W_ATLASIMAGE_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    h = self.height()

    currImageX0 = x0
    For idx = 1 To d.ImageIDs.Count
        imageWidth = d.imageWidths(idx)
        imageHeight = d.imageHeights(idx)
        imageID = d.ImageIDs(idx)
        imageY0 = y0 + (h - imageHeight) \ 2
        iV_DrawImageTint d.ImageFile, imageID, currImageX0, imageY0, pal_RGBA(255, 255, 255, 255)
        currImageX0 = currImageX0 + imageWidth + 2
    Next
End Sub

Public Function W_ATLASIMAGE_idealWidth(ByVal self As WIDGET) As Long
    W_ATLASIMAGE_idealWidth = W_ATLASIMAGE_(self).totalIdealWidth
End Function

Public Function W_ATLASIMAGE_idealHeight(ByVal self As WIDGET) As Long
    W_ATLASIMAGE_idealHeight = W_ATLASIMAGE_(self).maxImageHeight
End Function

'--------------------------------------------------------------------------
' MARK: - Image view

' (~W_IMAGE() deletes pOwnedBackgroundTexture: the data object's reference goes with the widget)
Private Function W_IMAGE_make(ByVal filename As String) As WIDGET
    Dim widget As WIDGET, d As W_IMAGE, pOwnedBackgroundTexture As gfx_api_texture, tw As Long, th As Long
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_W_IMAGE)
    WIDGET_construct widget
    widget.kind = K_W_IMAGE
    Set d = W_IMAGE_(widget)
    Set pOwnedBackgroundTexture = gfx_api_context_loadTextureFromFile(filename, texture_type_user_interface, -1, -1, True)
    If Not (Not pOwnedBackgroundTexture Is Nothing) Then ASSERT_FAILURE "pOwnedBackgroundTexture != nullptr", "src/screens/guidescreen.cpp:984", "make", "Failed to load texture from file: %s", filename: Exit Function
    Set d.pOwnedBackgroundTexture = pOwnedBackgroundTexture
    pOwnedBackgroundTexture.getDimensions tw, th
    If Not (tw <= 4096 And th <= 4096) Then ASSERT_FAILURE "textureDimensions.width <= 4096 && textureDimensions.height <= 4096", "src/screens/guidescreen.cpp:987", "make", "Unexpectedly large texture dimensions?: %zu x %zu", tw, th: Exit Function
    d.textureWidth = tw
    d.textureHeight = th
    Set W_IMAGE_make = widget
End Function

Public Function W_IMAGE_idealWidth(ByVal self As WIDGET) As Long
    W_IMAGE_idealWidth = W_IMAGE_(self).textureWidth
End Function

Public Function W_IMAGE_idealHeight(ByVal self As WIDGET) As Long
    W_IMAGE_idealHeight = W_IMAGE_(self).textureHeight
End Function

Public Sub W_IMAGE_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_IMAGE, x0 As Long, y0 As Long
    Set d = W_IMAGE_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    ' Draw the background
    If Not d.pOwnedBackgroundTexture Is Nothing Then
        iV_DrawImageAnisotropic d.pOwnedBackgroundTexture, x0, y0, 0!, 0!, self.width(), self.height(), 0!, pal_RGBA(255, 255, 255, 255)
    End If
End Sub

'--------------------------------------------------------------------------
' MARK: - Linked Topics widget

Private Sub WzGuideLinkedTopicsWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzGuideLinkedTopicsWidget
    WzGuideLinkedTopicsWidget_(self).weakSerial = WeakWidget_register(self)
End Sub

Public Sub WzGuideLinkedTopicsWidget_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzGuideLinkedTopicsWidget_(self).weakSerial
End Sub

Public Function WzGuideLinkedTopicsWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzGuideLinkedTopicsWidget
    Set d = WzGuideLinkedTopicsWidget_(self)
    If d.m_scrollableLinkedTopicsContainer Is Nothing Or d.m_seeAlsoLabel Is Nothing Then
        WzGuideLinkedTopicsWidget_idealHeight = 0
        Exit Function
    End If
    WzGuideLinkedTopicsWidget_idealHeight = MAX_(d.m_scrollableLinkedTopicsContainer.idealHeight(), d.m_seeAlsoLabel.idealHeight())
End Function

Private Function WzGuideLinkedTopicsWidget_rowHeight(ByVal self As WIDGET) As Long
    WzGuideLinkedTopicsWidget_rowHeight = iV_GetTextLineSize(font_regular) + lineSpacing
End Function

' make(const std::vector<std::string>& linkedTopicIdentifiers, registry, onClickLinkedTopicHandler): a Collection
' of the identifiers
Private Function WzGuideLinkedTopicsWidget_make(ByVal linkedTopicIdentifiers As Collection, ByVal onClickLinkedTopicHandler As WzFunc) As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideLinkedTopicsWidget)
    WzGuideLinkedTopicsWidget_construct widget
    WzGuideLinkedTopicsWidget_initialize widget, linkedTopicIdentifiers, onClickLinkedTopicHandler
    Set WzGuideLinkedTopicsWidget_make = widget
End Function

Private Sub WzGuideLinkedTopicsWidget_initialize(ByVal self As WIDGET, ByVal linkedTopicIdentifiers As Collection, ByVal onClickLinkedTopicHandler As WzFunc)
    Dim d As WzGuideLinkedTopicsWidget, topicID As Variant, topicDisplayName As String, linkedTopicWidg As WIDGET
    Dim topicIdentifierCopy As String, marginWrappedWidget As WIDGET, p As Padding, topicsContainerX0 As Long
    Set d = WzGuideLinkedTopicsWidget_(self)
    If linkedTopicIdentifiers.Count = 0 Then
        Exit Sub
    End If

    Set d.m_onClickLinkedTopicHandler = onClickLinkedTopicHandler

    Set d.m_seeAlsoLabel = W_LABEL_new()
    self.attach d.m_seeAlsoLabel
    W_LABEL_setFont d.m_seeAlsoLabel, font_regular, psPalette(WZCOL_FORM_TEXT)
    d.m_seeAlsoLabel.setString "See Also:"
    W_LABEL_setCanTruncate d.m_seeAlsoLabel, True
    d.m_seeAlsoLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.m_seeAlsoLabel), iV_GetTextLineSize(font_regular)

    Set d.m_linkedTopics = Paragraph_new()
    Paragraph_setFont d.m_linkedTopics, font_regular
    Paragraph_setFontColour d.m_linkedTopics, psPalette(WZCOL_TEXT_BRIGHT)
    Paragraph_setLineSpacing d.m_linkedTopics, lineSpacing
    d.m_linkedTopics.setGeometry 0, 0, 400, 40

    ' (std::weak_ptr<WzGuideLinkedTopicsWidget> psWeakLinkedTopicsWidg: d.weakSerial)

    For Each topicID In linkedTopicIdentifiers
        If WzGuideTopicsRegistry_getDisplayNameForTopic(CStr(topicID), topicDisplayName) Then
            Set linkedTopicWidg = WzGuideLinkedTopicButton_make(topicDisplayName)
            topicIdentifierCopy = topicID
            W_BUTTON_addOnClickHandler linkedTopicWidg, WzFunc_new(FN_guide_linkedTopicClicked, topicIdentifierCopy, d.weakSerial)
            Set marginWrappedWidget = Margin_wrap(0, 0, 0, 4, linkedTopicWidg)
            marginWrappedWidget.setGeometry 0, 0, marginWrappedWidget.idealWidth(), marginWrappedWidget.idealHeight()
            Paragraph_addWidget d.m_linkedTopics, marginWrappedWidget, WzGuideLinkedTopicButton_(linkedTopicWidg).ptsAboveBase ' linkedTopicWidg->aboveBase()
        End If
    Next

    Set d.m_scrollableLinkedTopicsContainer = ScrollableListWidget_make()
    ScrollableListWidget_setPadding d.m_scrollableLinkedTopicsContainer, p   ' {0, 0, 0, 0}
    ScrollableListWidget_addItem d.m_scrollableLinkedTopicsContainer, d.m_linkedTopics
    topicsContainerX0 = d.m_seeAlsoLabel.x() + d.m_seeAlsoLabel.width() + 10
    d.m_scrollableLinkedTopicsContainer.setGeometry topicsContainerX0, 0, 400, 45
    self.attach d.m_scrollableLinkedTopicsContainer
End Sub

Public Sub WzGuideLinkedTopicsWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As WzGuideLinkedTopicsWidget, w As Long, h As Long, topicsContainerX0 As Long, topicsContainerWidth As Long
    Set d = WzGuideLinkedTopicsWidget_(self)
    w = self.width()
    h = self.height()
    If w <= 1 And h <= 1 Then
        Exit Sub
    End If

    If d.m_seeAlsoLabel Is Nothing Then
        Exit Sub
    End If

    topicsContainerX0 = d.m_seeAlsoLabel.x() + d.m_seeAlsoLabel.width() + 10
    topicsContainerWidth = MAX_(w - topicsContainerX0, 0)
    If topicsContainerWidth <= 1 Then
        Exit Sub
    End If
    d.m_scrollableLinkedTopicsContainer.setGeometry topicsContainerX0, 0, topicsContainerWidth, h
End Sub

'--------------------------------------------------------------------------
' MARK: - WzGuideTopicDisplayWidget

Private Sub WzGuideTopicDisplayWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzGuideTopicDisplayWidget
    WzGuideTopicDisplayWidget_(self).weakSerial = WeakWidget_register(self)
End Sub

Public Sub WzGuideTopicDisplayWidget_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzGuideTopicDisplayWidget_(self).weakSerial
End Sub

Private Function WzGuideTopicDisplayWidget_make(ByVal topic As WzWrappedGuideTopic) As WIDGET
    Dim widget As WIDGET
    If Not (Not topic.topic Is Nothing) Then ASSERT_FAILURE "psTopic != nullptr", "src/screens/guidescreen.cpp:1177", "make", "Null topic?": Exit Function
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideTopicDisplayWidget)
    WzGuideTopicDisplayWidget_construct widget
    WzGuideTopicDisplayWidget_initialize widget, topic
    Set WzGuideTopicDisplayWidget_make = widget
End Function

' returns the width for a given height, maintaining the original widget's aspect ratio (according to idealWidth/idealHeight)
Private Function widgScaledWidthForHeight(ByVal widg As WIDGET, ByVal height As Long) As Long
    Dim idealHeight As Long, idealWidth As Long, heightScaling As Double
    idealHeight = widg.idealHeight()
    idealWidth = widg.idealWidth()
    If height = widg.idealHeight() Then
        widgScaledWidthForHeight = idealWidth
        Exit Function
    End If
    If idealHeight <= 0 Then
        widgScaledWidthForHeight = 0
        Exit Function
    End If
    heightScaling = CDbl(height) / CDbl(idealHeight)
    widgScaledWidthForHeight = CLng(Fix(idealWidth * heightScaling))
End Function

Public Sub WzGuideTopicDisplayWidget_geometryChanged(ByVal self As WIDGET)
    Dim d As WzGuideTopicDisplayWidget, remainingHeight As Long, w As Long, usableWidth As Long, maxUsableX1 As Long
    Dim nextLineY0 As Long, visualX0 As Long, maxIdealVisualHeight As Long, visual As WIDGET, visualIdealHeight As Long
    Dim visualHeight As Long, visualWidth As Long, visualY0 As Long, visualY1 As Long, linkedTopicsHeight As Long
    Set d = WzGuideTopicDisplayWidget_(self)
    ' (VB) not before initialize made the display name (WIDGET's constructor reaches this override)
    If d.m_displayName Is Nothing Then Exit Sub
    remainingHeight = self.height()
    w = self.width()
    usableWidth = MAX_(w - (internalPaddingX * 2), 0)
    maxUsableX1 = MAX_(w - internalPaddingX, 0)

    If usableWidth <= 1 Then
        Exit Sub
    End If

    nextLineY0 = internalMarginTop + internalPaddingY
    remainingHeight = remainingHeight - nextLineY0

    If Not d.m_parentNavPath Is Nothing Then
        ' display nav path bar at top
        d.m_parentNavPath.setGeometry internalPaddingX, nextLineY0, MIN_(W_LABEL_getMaxLineWidth(d.m_parentNavPath), usableWidth), d.m_parentNavPath.height()
        remainingHeight = remainingHeight - (d.m_parentNavPath.height() + internalPaddingY)
        nextLineY0 = nextLineY0 + d.m_parentNavPath.height() + internalPaddingY
    End If

    ' display name is first
    d.m_displayName.setGeometry internalPaddingX, nextLineY0, MIN_(W_LABEL_getMaxLineWidth(d.m_displayName), usableWidth), d.m_displayName.height()
    remainingHeight = remainingHeight - (d.m_displayName.height() + internalPaddingY)
    nextLineY0 = nextLineY0 + d.m_displayName.height() + internalPaddingY

    ' followed by the visual(s) (if present)
    If d.m_visuals.Count > 0 Then
        visualX0 = internalPaddingX
        maxIdealVisualHeight = 0
        For Each visual In d.m_visuals
            maxIdealVisualHeight = MAX_(maxIdealVisualHeight, visual.idealHeight())
        Next
        maxIdealVisualHeight = MIN_(maxVisualHeight, maxIdealVisualHeight)
        For Each visual In d.m_visuals
            visualIdealHeight = visual.idealHeight()
            visualHeight = MIN_(maxVisualHeight, visualIdealHeight)
            visualWidth = widgScaledWidthForHeight(visual, visualHeight)
            visualY0 = nextLineY0 + (maxIdealVisualHeight - visualHeight) \ 2
            visualY1 = visualX0 + visualWidth
            If visualY1 > maxUsableX1 Then
                visual.hide
            Else
                visual.show
                visual.setGeometry visualX0, visualY0, visualWidth, visualHeight
                visualX0 = visualX0 + visualWidth + visualBetweenPaddingY
            End If
        Next

        remainingHeight = remainingHeight - (maxIdealVisualHeight + internalPaddingY)
        nextLineY0 = nextLineY0 + maxIdealVisualHeight + internalPaddingY
    End If

    ' the linked topics container takes up space at the bottom
    If Not d.m_linkedTopicsWidget Is Nothing Then
        If usableWidth > 1 And remainingHeight > 1 Then
            d.m_linkedTopicsWidget.setGeometry internalPaddingX, nextLineY0, usableWidth, d.maxLinkedTopicsHeight ' set once to set width
            ' set again to set height based on idealHeight for that width (if smaller)
            linkedTopicsHeight = MIN_(d.m_linkedTopicsWidget.idealHeight(), d.maxLinkedTopicsHeight)
            d.m_linkedTopicsWidget.setGeometry internalPaddingX, self.height() - (linkedTopicsHeight + internalPaddingY), usableWidth, linkedTopicsHeight
            remainingHeight = remainingHeight - (d.m_linkedTopicsWidget.height() + internalPaddingY)
        End If
    End If

    ' the contents take up the rest of the space
    If Not d.m_scrollableContentsContainer Is Nothing Then
        remainingHeight = remainingHeight - internalPaddingY
        If usableWidth > 1 And remainingHeight > 1 Then
            d.m_scrollableContentsContainer.setGeometry internalPaddingX, nextLineY0, usableWidth, remainingHeight
        End If
    End If
End Sub

Private Function WzGuideTopicDisplayWidget_constructParentNavPathWidget(ByVal topic As WzWrappedGuideTopic) As WIDGET
    Dim psParentTopic As WzWrappedGuideTopic, pathComponents As Collection, pathStr As String, i As Long, result As WIDGET
    If Len(topic.parentId) > 0 Then Set psParentTopic = WzGuideTopicsRegistry_getTopic(topic.parentId)  ' topic.parent.lock()
    If psParentTopic Is Nothing Then
        Exit Function
    End If

    Set pathComponents = New Collection
    Do While Not psParentTopic Is Nothing
        pathComponents.Add psParentTopic.topic.displayName
        If Len(psParentTopic.parentId) > 0 Then
            Set psParentTopic = WzGuideTopicsRegistry_getTopic(psParentTopic.parentId)
        Else
            Set psParentTopic = Nothing
        End If
    Loop

    If pathComponents.Count <= 1 Then ' don't display for first-level entries
        Exit Function
    End If

    For i = pathComponents.Count To 1 Step -1
        If Len(pathStr) > 0 Then
            pathStr = pathStr & " " & ChrW$(&HBB) & " "    ' " » "
        End If
        pathStr = pathStr & pathComponents(i)
    Next

    Set result = W_LABEL_new()
    W_LABEL_setFont result, font_bar, psPalette(WZCOL_TEXT_MEDIUM)
    result.setString pathStr
    W_LABEL_setCanTruncate result, True
    result.setGeometry 0, 0, W_LABEL_getMaxLineWidth(result), iV_GetTextLineSize(font_bar)

    Set WzGuideTopicDisplayWidget_constructParentNavPathWidget = result
End Function

' void initialize(const WzWrappedGuideTopic& topic, const std::shared_ptr<WzGuideTopicsRegistry>& registry)
Private Sub WzGuideTopicDisplayWidget_initialize(ByVal self As WIDGET, ByVal psWrappedTopic As WzWrappedGuideTopic)
    Dim d As WzGuideTopicDisplayWidget, topic As WzGuideTopic, i As Long, visual As WzGuideVisual, retButtonView As WIDGET
    Dim statPieView As WIDGET, imageHeight As Long, orderButtonImages() As Long, nImages As Long, orderButtonImagesWidg As WIDGET
    Dim image As WIDGET, idx As Long, displayString As String, foundLinkedTopicIdentifiers As Collection, topicID As Variant
    Dim p As Padding
    Set d = WzGuideTopicDisplayWidget_(self)
    Set topic = psWrappedTopic.topic

    d.m_topicID = topic.identifier

    ' determine if a the parent navpath needs to be shown
    Set d.m_parentNavPath = WzGuideTopicDisplayWidget_constructParentNavPathWidget(psWrappedTopic)
    If Not d.m_parentNavPath Is Nothing Then
        self.attach d.m_parentNavPath
    End If

    Set d.m_displayName = W_LABEL_new()
    self.attach d.m_displayName
    W_LABEL_setFont d.m_displayName, font_medium_bold, psPalette(WZCOL_TEXT_BRIGHT)
    d.m_displayName.setString topic.displayName
    W_LABEL_setCanTruncate d.m_displayName, True
    d.m_displayName.setGeometry 0, 0, W_LABEL_getMaxLineWidth(d.m_displayName), iV_GetTextLineSize(font_medium)

    ' the visuals (reticuleButtonId, statId, imagePath, etc)
    For i = 0 To topic.visualsCount - 1
        visual = topic.visual(i)
        If visual.has_reticuleButtonId Then
            Set retButtonView = WzGuideRetButtonView_make(visual.reticuleButtonId)
            If Not retButtonView Is Nothing Then
                self.attach retButtonView, ChildZPos_Back
                retButtonView.setGeometry 0, 0, retButtonView.idealWidth(), retButtonView.idealHeight()
                d.m_visuals.Add retButtonView
            End If
        ElseIf visual.has_statVisual Then
            Set statPieView = WzGuideStatPIEView_make(visual)
            If Not statPieView Is Nothing Then
                self.attach statPieView, ChildZPos_Back
                imageHeight = MIN_(maxVisualHeight, statPieView.idealHeight())
                statPieView.setGeometry 0, 0, statPieView.idealWidth(), imageHeight
                d.m_visuals.Add statPieView
            End If
        ElseIf Not visual.droidTemplateJSON Is Nothing Then
            Set statPieView = WzGuideStatPIEView_makeFromDroidTemplateJSON(visual.droidTemplateJSON)
            If Not statPieView Is Nothing Then
                self.attach statPieView, ChildZPos_Back
                imageHeight = MIN_(maxVisualHeight, statPieView.idealHeight())
                statPieView.setGeometry 0, 0, statPieView.idealWidth(), imageHeight
                d.m_visuals.Add statPieView
            End If
        ElseIf visual.has_secondaryUnitOrderButtons Then
            nImages = intOrderGetButtonImageIDs(visual.secondaryUnitOrderButtons, orderButtonImages)
            If nImages > 0 Then
                Set orderButtonImagesWidg = W_ATLASIMAGE_make(IntImages, orderButtonImages, nImages)
                If Not orderButtonImagesWidg Is Nothing Then
                    self.attach orderButtonImagesWidg, ChildZPos_Back
                    imageHeight = MIN_(maxVisualHeight, orderButtonImagesWidg.idealHeight())
                    orderButtonImagesWidg.setGeometry 0, 0, orderButtonImagesWidg.idealWidth(), imageHeight
                    d.m_visuals.Add orderButtonImagesWidg
                End If
            End If
        ElseIf visual.has_imagePath Then
            Set image = W_IMAGE_make(visual.imagePath)
            If Not image Is Nothing Then
                self.attach image, ChildZPos_Back
                imageHeight = MIN_(maxVisualHeight, image.idealHeight())
                image.setGeometry 0, 0, image.width(), imageHeight
                d.m_visuals.Add image
            End If
        End If
    Next

    Set d.m_contents = Paragraph_new()
    Paragraph_setFont d.m_contents, font_regular
    Paragraph_setFontColour d.m_contents, psPalette(WZCOL_TEXT_BRIGHT)
    d.m_contents.setGeometry 0, 0, 400, 400
    Paragraph_setLineSpacing d.m_contents, 2

    For idx = 1 To topic.contents.Count
        If idx > 1 Then
            Paragraph_addText d.m_contents, vbLf
        End If
        displayString = topic.contents(idx)         ' localizedStr.getLocalizedString(WZ2100_GUIDE_MESSAGE_CATALOG_DOMAIN)
        WzGuideTopicDisplayWidget_addFormattedTextLine d.m_contents, displayString
    Next

    Set d.m_scrollableContentsContainer = ScrollableListWidget_make()
    ScrollableListWidget_addItem d.m_scrollableContentsContainer, d.m_contents
    p.top = 0: p.right_ = 4: p.bottom = 0: p.left_ = 0
    ScrollableListWidget_setPadding d.m_scrollableContentsContainer, p
    d.m_scrollableContentsContainer.setGeometry 0, 0, 400, 405
    self.attach d.m_scrollableContentsContainer

    ' display linked topics that have already been added to the guide topics registry
    Set foundLinkedTopicIdentifiers = New Collection
    For Each topicID In topic.linkedTopicIdentifiers
        If Not WzGuideTopicsRegistry_getTopic(CStr(topicID)) Is Nothing Then
            foundLinkedTopicIdentifiers.Add topicID
        End If
    Next
    If foundLinkedTopicIdentifiers.Count > 0 Then
        ' (std::weak_ptr<WzGuideTopicDisplayWidget> psWeakGuideTopicWidg: d.weakSerial)
        Set d.m_linkedTopicsWidget = WzGuideLinkedTopicsWidget_make(foundLinkedTopicIdentifiers, WzFunc_new(FN_guide_topicDisplay_onClickLinked, d.weakSerial))
        d.maxLinkedTopicsHeight = (WzGuideLinkedTopicsWidget_rowHeight(d.m_linkedTopicsWidget) * 2)
        d.m_linkedTopicsWidget.setGeometry 0, 0, 400, d.maxLinkedTopicsHeight
        self.attach d.m_linkedTopicsWidget
    End If
End Sub

' initialize's addFormattedTextLine lambda ([&](std::string& line))
Private Sub WzGuideTopicDisplayWidget_addFormattedTextLine(ByVal contents As WIDGET, ByVal textLine As String)
    Dim has_oldFont As Boolean, oldFont As Long, newFont As Long
    If Len(textLine) > 0 Then
        ' If textLine starts with "#", treat as heading or subheading (i.e. changing font size)
        If Left$(textLine, 1) = "#" Then
            oldFont = Paragraph_getTextStyle(contents).font
            has_oldFont = True
            newFont = font_medium_bold
            If Len(textLine) > 1 And Mid$(textLine, 2, 1) = "#" Then
                newFont = font_regular_bold
            End If
            Paragraph_setFont contents, newFont
        End If
    Else
        textLine = " "
    End If
    Paragraph_addText contents, textLine
    If has_oldFont Then
        Paragraph_setFont contents, oldFont
    End If
End Sub

'--------------------------------------------------------------------------
' MARK: - WzGuideTopicsNavBar implementation

' (VB) std::make_shared<WzGuideTopicsNavButton>(): W_BUTTON() and the members' initialisers (Class_Initialize)
Private Function WzGuideTopicsNavButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzGuideTopicsNavButton)
    W_BUTTON_construct widget
    widget.kind = K_WzGuideTopicsNavButton
    Set WzGuideTopicsNavButton_new = widget
End Function

' static std::shared_ptr<WzGuideTopicsNavButton> make(const WzString& buttonText, iV_fonts FontID = font_regular)
Private Function WzGuideTopicsNavButton_make(ByVal buttonText As String, Optional ByVal FontID As iV_fonts = font_regular) As WIDGET
    Dim result As WIDGET
    Set result = WzGuideTopicsNavButton_new()
    W_BUTTON_(result).FontID = FontID
    W_BUTTON_(result).pText = buttonText
    Set WzGuideTopicsNavButton_make = result
End Function

' static std::shared_ptr<WzGuideTopicsNavButton> make(UWORD intfacImageID, int imageDimensions = 16)
Private Function WzGuideTopicsNavButton_make_image(ByVal intfacImageID As Long, Optional ByVal imageDimensions As Long = 16) As WIDGET
    Dim result As WIDGET, d As WzGuideTopicsNavButton
    If Not assertValidImage(IntImages, intfacImageID) Then
        Exit Function
    End If
    Set result = WzGuideTopicsNavButton_new()
    Set d = WzGuideTopicsNavButton_(result)
    d.has_intfacImageID = True
    d.intfacImageID = intfacImageID
    d.imageDimensions = imageDimensions
    Set WzGuideTopicsNavButton_make_image = result
End Function

Public Function WzGuideTopicsNavButton_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WzGuideTopicsNavButton
    Set d = WzGuideTopicsNavButton_(self)
    If d.has_intfacImageID Then
        WzGuideTopicsNavButton_idealWidth = d.imageDimensions + (GuideNavBarButtonHorizontalPadding * 2)
        Exit Function
    End If
    WzGuideTopicsNavButton_idealWidth = iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID) + (GuideNavBarButtonHorizontalPadding * 2)
End Function

' Display the text, shrunk to fit, with a faint border around it (that gets bright when highlighted)
Public Sub WzGuideTopicsNavButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzGuideTopicsNavButton, pParent As WIDGET, xPos As Long, yPos As Long, h As Long, w As Long
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, has_backColor As Boolean, backColor As Long
    Dim msgColour As Long, imgPosX0 As Long, imgPosY0 As Long, imgColor As Long, pText As String
    Dim fontID As Long, result As iV_fonts, availableButtonTextWidth As Long
    Dim textX0 As Long, textY0 As Long, maxTextDisplayableWidth As Long
    Set d = WzGuideTopicsNavButton_(self)
    Set pParent = self.parent()
    If Not (Not pParent Is Nothing) Then ASSERT_FAILURE "pParent != nullptr", "src/screens/guidescreen.cpp:1521", "display", "Keymap buttons should have a parent container!": Exit Sub

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
    If Not isDown Then msgColour = psPalette(WZCOL_TEXT_BRIGHT) Else msgColour = psPalette(WZCOL_FORM_TEXT)
    If isDisabled Then
        msgColour = psPalette(WZCOL_TEXT_BRIGHT)
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
    textX0 = MAX_(xPos + GuideNavBarButtonHorizontalPadding, xPos + ((w - d.wzMessageText.width()) \ 2))
    textY0 = CLng(Fix(yPos + (h - d.wzMessageText.lineSize()) \ 2 - CSng(d.wzMessageText.aboveBase())))

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

Public Function WzGuideTopicsNavButton_getTip(ByVal self As WIDGET) As String
    If Not WzGuideTopicsNavButton_(self).isTruncated Then
        WzGuideTopicsNavButton_getTip = W_BUTTON_(self).pTip
        Exit Function
    End If

    WzGuideTopicsNavButton_getTip = W_BUTTON_(self).pText
End Function

' WzGuideTopicsNavBar(): WIDGET() {}
Private Sub WzGuideTopicsNavBar_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzGuideTopicsNavBar
    WzGuideTopicsNavBar_(self).weakSerial = WeakWidget_register(self)
End Sub

Public Sub WzGuideTopicsNavBar_destruct(ByVal self As WIDGET)
    Dim d As WzGuideTopicsNavBar
    Set d = WzGuideTopicsNavBar_(self)
    If Not d.optionsOverlayScreen Is Nothing Then
        widgRemoveOverlayScreen d.optionsOverlayScreen
        Set d.optionsOverlayScreen = Nothing
    End If
    WeakWidget_unregister d.weakSerial
End Sub

Public Function WzGuideTopicsNavBar_idealHeight(ByVal self As WIDGET) As Long
    WzGuideTopicsNavBar_idealHeight = 30 ' TODO: ??
End Function

Private Function WzGuideTopicsNavBar_make(ByVal topicDisplayWidth As Long, ByVal popStateHandler As WzFunc, ByVal onSidebarToggleHandler As WzFunc, ByVal onCloseHandler As WzFunc) As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideTopicsNavBar)
    WzGuideTopicsNavBar_construct widget
    WzGuideTopicsNavBar_initialize widget, topicDisplayWidth, popStateHandler, onSidebarToggleHandler, onCloseHandler
    Set WzGuideTopicsNavBar_make = widget
End Function

Private Sub WzGuideTopicsNavBar_initialize(ByVal self As WIDGET, ByVal topicDisplayWidth As Long, ByVal popStateHandler As WzFunc, ByVal onSidebarToggleHandler As WzFunc, ByVal onCloseHandler As WzFunc)
    Dim d As WzGuideTopicsNavBar
    Set d = WzGuideTopicsNavBar_(self)
    d.m_topicDisplayWidth = topicDisplayWidth
    Set d.m_popStateHandler = popStateHandler
    Set d.m_onSidebarToggleHandler = onSidebarToggleHandler
    Set d.m_onCloseHandler = onCloseHandler

    ' Title
    Set d.m_titleLabel = W_LABEL_new()
    W_LABEL_setFont d.m_titleLabel, font_regular
    d.m_titleLabel.setString "Game Guide"
    d.m_titleLabel.setTransparentToMouse True
    self.attach d.m_titleLabel

    ' (std::weak_ptr<WzGuideTopicsNavBar> psWeakNavBar: d.weakSerial)

    ' Close button
    Set d.m_closeButton = WzGuideTopicsNavButton_make(ChrW$(&H2715), font_regular) ' U+2715
    d.m_closeButton.setTip "Close"
    self.attach d.m_closeButton
    W_BUTTON_addOnClickHandler d.m_closeButton, WzFunc_new(FN_guide_navbar_close, d.weakSerial)

    ' Options button ' U+2699
    Set d.m_optionsButton = WzGuideTopicsNavButton_make(ChrW$(&H2699), font_regular)
    d.m_optionsButton.setTip "Options"
    self.attach d.m_optionsButton
    W_BUTTON_addOnClickHandler d.m_optionsButton, WzFunc_new(FN_guide_navbar_options, d.weakSerial)

    ' Sidebar button
    If Not d.m_onSidebarToggleHandler Is Nothing Then
        Set d.m_sidebarButton = WzGuideTopicsNavButton_make_image(IMAGE_INTFAC_SIDEBAR_LIST, 12)
        If d.m_sidebarButton Is Nothing Then
            ' failed to get the image - use an "s"
            Set d.m_sidebarButton = WzGuideTopicsNavButton_make("s", font_regular)
        End If
        self.attach d.m_sidebarButton
        W_BUTTON_addOnClickHandler d.m_sidebarButton, WzFunc_new(FN_guide_navbar_sidebar, d.weakSerial)
        d.m_sidebarButton.setTip "Toggle Sidebar"
    End If

    ' Back / Forward buttons
    Set d.m_backButton = WzGuideTopicsNavButton_make(ChrW$(&H2039), font_medium) ' U+2039
    self.attach d.m_backButton
    d.m_backButton.setState WBUT_DISABLE
    W_BUTTON_addOnClickHandler d.m_backButton, WzFunc_new(FN_guide_navbar_back, d.weakSerial)

    Set d.m_nextButton = WzGuideTopicsNavButton_make(ChrW$(&H203A), font_medium) ' U+203A
    self.attach d.m_nextButton
    d.m_nextButton.setState WBUT_DISABLE
    W_BUTTON_addOnClickHandler d.m_nextButton, WzFunc_new(FN_guide_navbar_next, d.weakSerial)

    WzGuideTopicsNavBar_updateLayout self
End Sub

Public Sub WzGuideTopicsNavBar_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, h As Long, w As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    h = self.height()
    w = self.width()
    iV_Line x0, y0 + h, x0 + w, y0 + h, pal_RGBA(255, 255, 255, 255)
End Sub

Public Sub WzGuideTopicsNavBar_geometryChanged(ByVal self As WIDGET)
    ' (VB) not before initialize made the buttons (WIDGET's constructor reaches this override)
    If WzGuideTopicsNavBar_(self).m_nextButton Is Nothing Then Exit Sub
    WzGuideTopicsNavBar_updateLayout self
End Sub

Private Sub WzGuideTopicsNavBar_updateLayout(ByVal self As WIDGET)
    Dim d As WzGuideTopicsNavBar, x0 As Long, y0 As Long, w As Long, h As Long, widgetHeight As Long, widgetY0 As Long
    Dim buttonWidth As Long, sidebarButtonX1 As Long, navButtonsX0 As Long, navButtonsX1 As Long, titleX0 As Long
    Set d = WzGuideTopicsNavBar_(self)
    x0 = self.x()
    y0 = self.y()
    w = self.width()
    h = self.height()

    ' The "m_topicDisplayWidth" area is a right-aligned area used for:
    ' - the title (centered within it)
    ' - the close button (right side of it)
    ' - the action buttons (back / forward - left side of it*)

    ' The sidebar button is always aligned with the overall left-side of the navbar.
    ' If this is within the m_topicDisplayWidth area, the back/forward buttons are aligned next to it
    ' If not, they are aligned with the left side of the m_topicDisplayWidth area

    widgetHeight = h - (GuideNavBarButtonHorizontalPadding * 2)
    widgetY0 = y0 + ((h - widgetHeight) \ 2)

    ' std::max<int>({back, next, close, options or 0, sidebar or 0}'s idealWidth())
    buttonWidth = MAX_(d.m_backButton.idealWidth(), d.m_nextButton.idealWidth())
    buttonWidth = MAX_(buttonWidth, d.m_closeButton.idealWidth())
    If Not d.m_optionsButton Is Nothing Then buttonWidth = MAX_(buttonWidth, d.m_optionsButton.idealWidth()) Else buttonWidth = MAX_(buttonWidth, 0)
    If Not d.m_sidebarButton Is Nothing Then buttonWidth = MAX_(buttonWidth, d.m_sidebarButton.idealWidth()) Else buttonWidth = MAX_(buttonWidth, 0)

    ' Start by left-aligning the sidebar button with the overall widget
    sidebarButtonX1 = 0
    If Not d.m_sidebarButton Is Nothing Then
        d.m_sidebarButton.setGeometry x0 + GuideNavBarButtonHorizontalPadding, widgetY0, buttonWidth, widgetHeight
        sidebarButtonX1 = d.m_sidebarButton.x() + d.m_sidebarButton.width()
    End If

    ' Align the back & forward buttons
    If w > d.m_topicDisplayWidth Then navButtonsX0 = w - d.m_topicDisplayWidth Else navButtonsX0 = 0
    navButtonsX0 = MAX_(sidebarButtonX1 + GuideNavBarButtonHorizontalPadding, navButtonsX0 + GuideNavBarButtonHorizontalPadding)
    d.m_backButton.setGeometry navButtonsX0, widgetY0, buttonWidth, widgetHeight
    navButtonsX0 = navButtonsX0 + d.m_backButton.width() + GuideNavBarButtonHorizontalPadding
    d.m_nextButton.setGeometry navButtonsX0, widgetY0, buttonWidth, widgetHeight
    navButtonsX1 = d.m_nextButton.x() + d.m_nextButton.width()

    ' Center the title within the right-aligned topicDisplayWidth area
    titleX0 = (MIN_(d.m_topicDisplayWidth, w) - d.m_titleLabel.idealWidth()) \ 2
    If w > d.m_topicDisplayWidth Then titleX0 = titleX0 + (w - d.m_topicDisplayWidth)
    titleX0 = MAX_(titleX0, navButtonsX1 + GuideNavBarButtonHorizontalPadding)
    d.m_titleLabel.setGeometry titleX0, widgetY0, d.m_titleLabel.idealWidth(), widgetHeight

    ' Right-align the options and close buttons
    d.m_closeButton.setGeometry w - (buttonWidth + GuideNavBarButtonHorizontalPadding), widgetY0, buttonWidth, widgetHeight
    d.m_optionsButton.setGeometry d.m_closeButton.x() - (buttonWidth + GuideNavBarButtonHorizontalPadding), widgetY0, buttonWidth, widgetHeight
End Sub

Private Sub WzGuideTopicsNavBar_triggerCloseHandler(ByVal self As WIDGET)
    Dim d As WzGuideTopicsNavBar
    Set d = WzGuideTopicsNavBar_(self)
    If Not (Not d.m_onCloseHandler Is Nothing) Then ASSERT_FAILURE "m_onCloseHandler != nullptr", "src/screens/guidescreen.cpp:1793", "triggerCloseHandler", "Missing onCloseHandler": Exit Sub
    WzFunc_call d.m_onCloseHandler
End Sub

Private Sub WzGuideTopicsNavBar_triggerSidebarToggleHandler(ByVal self As WIDGET)
    Dim d As WzGuideTopicsNavBar
    Set d = WzGuideTopicsNavBar_(self)
    If Not (Not d.m_onSidebarToggleHandler Is Nothing) Then ASSERT_FAILURE "m_onSidebarToggleHandler != nullptr", "src/screens/guidescreen.cpp:1799", "triggerSidebarToggleHandler", "Missing onSidebarToggleHandler": Exit Sub
    WzFunc_call d.m_onSidebarToggleHandler
End Sub

Private Sub WzGuideTopicsNavBar_updateButtonStates(ByVal self As WIDGET)
    Dim d As WzGuideTopicsNavBar
    Set d = WzGuideTopicsNavBar_(self)
    If d.m_currentStateIndex > 0 Then d.m_backButton.setState 0 Else d.m_backButton.setState WBUT_DISABLE
    If d.m_stateHistory.Count > 0 And (d.m_currentStateIndex < (d.m_stateHistory.Count - 1)) Then d.m_nextButton.setState 0 Else d.m_nextButton.setState WBUT_DISABLE
End Sub

Private Sub WzGuideTopicsNavBar_pushState(ByVal self As WIDGET, ByVal topicIdentifier As String)
    Dim d As WzGuideTopicsNavBar
    Set d = WzGuideTopicsNavBar_(self)
    If d.m_stateHistory.Count > 0 And d.m_currentStateIndex <> (d.m_stateHistory.Count - 1) Then
        ' truncate state history
        Do While d.m_stateHistory.Count > d.m_currentStateIndex + 1     ' m_stateHistory.resize(m_currentStateIndex + 1)
            d.m_stateHistory.Remove d.m_stateHistory.Count
        Loop
    End If
    d.m_stateHistory.Add topicIdentifier
    d.m_currentStateIndex = d.m_stateHistory.Count - 1
    WzGuideTopicsNavBar_updateButtonStates self
End Sub

Private Function WzGuideTopicsNavBar_navBack(ByVal self As WIDGET) As Boolean
    Dim d As WzGuideTopicsNavBar
    Set d = WzGuideTopicsNavBar_(self)
    If d.m_currentStateIndex = 0 Then
        Exit Function
    End If
    d.m_currentStateIndex = d.m_currentStateIndex - 1
    If Not d.m_popStateHandler Is Nothing Then
        WzFunc_call d.m_popStateHandler, CStr(d.m_stateHistory(d.m_currentStateIndex + 1))
    End If
    WzGuideTopicsNavBar_updateButtonStates self
    WzGuideTopicsNavBar_navBack = True
End Function

Private Function WzGuideTopicsNavBar_navForward(ByVal self As WIDGET) As Boolean
    Dim d As WzGuideTopicsNavBar
    Set d = WzGuideTopicsNavBar_(self)
    If d.m_currentStateIndex = (d.m_stateHistory.Count - 1) Then
        Exit Function
    End If
    d.m_currentStateIndex = d.m_currentStateIndex + 1
    If Not d.m_popStateHandler Is Nothing Then
        WzFunc_call d.m_popStateHandler, CStr(d.m_stateHistory(d.m_currentStateIndex + 1))
    End If
    WzGuideTopicsNavBar_updateButtonStates self
    WzGuideTopicsNavBar_navForward = True
End Function

' createOptionsPopoverForm's createOptionsCheckbox lambda
Private Function createOptionsCheckbox(ByVal text As String, ByVal isChecked As Boolean, ByVal isDisabled As Boolean, ByVal onClickFunc As WzFunc) As WIDGET
    Dim pCheckbox As WIDGET, minimumDimensions As Vector2i
    Set pCheckbox = WzCheckboxButton_new()
    W_BUTTON_(pCheckbox).pText = text
    W_BUTTON_(pCheckbox).FontID = font_regular
    WzCheckboxButton_setIsChecked pCheckbox, isChecked
    WzCheckboxButton_setTextColor pCheckbox, psPalette(WZCOL_TEXT_BRIGHT)
    If isDisabled Then
        pCheckbox.setState WBUT_DISABLE
    End If
    minimumDimensions = WzCheckboxButton_calculateDesiredDimensions(pCheckbox)
    pCheckbox.setGeometry 0, 0, minimumDimensions.x, minimumDimensions.y
    If Not onClickFunc Is Nothing Then
        W_BUTTON_addOnClickHandler pCheckbox, WzFunc_new(FN_guide_optionsCheckbox_click, onClickFunc)
    End If
    Set createOptionsCheckbox = pCheckbox
End Function

Private Function WzGuideTopicsNavBar_createOptionsPopoverForm(ByVal self As WIDGET) As WIDGET
    Dim buttons As Collection, button As WIDGET, totalButtonHeight As Long, maxButtonWidth As Long, itemsList As WIDGET
    Dim pad As Padding
    Const itemSpacing As Long = 4
    ' create all the buttons / option rows
    Set buttons = New Collection
    buttons.Add createOptionsCheckbox("Disable Topic Pop-ups", disableTopicPopups, False, WzFunc_new(FN_guide_disablePopups_onClick))

    ' determine required height for all buttons
    totalButtonHeight = 0
    For Each button In buttons
        totalButtonHeight = totalButtonHeight + button.height()
    Next
    maxButtonWidth = buttons(1).width()              ' (*max_element(buttons, by width))->width()
    For Each button In buttons
        If maxButtonWidth < button.width() Then maxButtonWidth = button.width()
    Next

    Set itemsList = ScrollableListWidget_make()
    ScrollableListWidget_setBackgroundColor itemsList, psPalette(WZCOL_MENU_BACKGROUND)
    ScrollableListWidget_setBorderColor itemsList, pal_RGBA(152, 151, 154, 255)
    pad.top = 8: pad.right_ = 8: pad.bottom = 8: pad.left_ = 8
    ScrollableListWidget_setPadding itemsList, pad
    ScrollableListWidget_setItemSpacing itemsList, itemSpacing
    itemsList.setGeometry itemsList.x(), itemsList.y(), maxButtonWidth + pad.left_ + pad.right_, totalButtonHeight + ((buttons.Count - 1) * itemSpacing) + pad.top + pad.bottom
    For Each button In buttons
        ScrollableListWidget_addItem itemsList, button
    Next

    Set WzGuideTopicsNavBar_createOptionsPopoverForm = itemsList
End Function

' (VB) psParent is the nav bar itself (its only call): the pop-over's layout locks it through the nav bar's serial
Private Sub WzGuideTopicsNavBar_displayOptionsOverlay(ByVal self As WIDGET, ByVal psParent As WIDGET)
    Dim d As WzGuideTopicsNavBar, lockedScreen As W_SCREEN, newRootFrm As WIDGET, optionsPopOver As WIDGET
    Set d = WzGuideTopicsNavBar_(self)
    Set lockedScreen = self.screenPointer
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "src/screens/guidescreen.cpp:1909", "displayOptionsOverlay", "The WzPlayerBoxTabs does not have an associated screen pointer?"

    ' Initialize the options overlay screen
    Set d.optionsOverlayScreen = W_SCREEN_make()
    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    Set newRootFrm.displayFunction = WzFunc_new(FN_displayChildDropShadows)
    ' (std::weak_ptr<W_SCREEN> psWeakOptionsOverlayScreen: its address; psWeakNavBar: d.weakSerial)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_guide_optionsOverlay_clicked, ObjPtr(d.optionsOverlayScreen), d.weakSerial)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
    d.optionsOverlayScreen.psForm.attach newRootFrm

    ' Create the pop-over form
    Set optionsPopOver = WzGuideTopicsNavBar_createOptionsPopoverForm(self)
    newRootFrm.attach optionsPopOver

    ' Position the pop-over form
    optionsPopOver.setCalcLayout WzFunc_new(FN_guide_optionsPopOver_calcLayout, d.weakSerial)

    widgRegisterOverlayScreenOnTopOfScreen d.optionsOverlayScreen, lockedScreen
End Sub

'--------------------------------------------------------------------------
' MARK: - WzGuideSidebar implementation

' WzGuideSidebarTopicButton() {}
Private Sub WzGuideSidebarTopicButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzGuideSidebarTopicButton
End Sub

Private Sub WzGuideSidebarTopicButton_initialize(ByVal self As WIDGET, ByVal topicIdentifier As String, ByVal text As String, ByVal indentLevel As Long)
    Dim d As WzGuideSidebarTopicButton, b As W_BUTTON, minButtonWidthForText As Long
    Set d = WzGuideSidebarTopicButton_(self)
    Set b = W_BUTTON_(self)
    d.topicIdentifier = topicIdentifier
    d.indentLevel = indentLevel
    If d.indentLevel > 0 Then b.FontID = font_regular Else b.FontID = font_regular_bold
    self.setString text
    minButtonWidthForText = iV_GetTextWidth(b.pText, b.FontID)
    self.setGeometry 0, 0, minButtonWidthForText + (leftPadding + rightPadding), iV_GetTextLineSize(b.FontID) + 12
End Sub

' make(const std::shared_ptr<WzGuideTopic>& topic, size_t indentLevel, optional<ActivityDBProtocol::GuideTopicPrefs> prefs):
' (VB) the wrapped topic, whose prefs fields hold prefs when hasPrefs
Private Function WzGuideSidebarTopicButton_make(ByVal wrappedTopic As WzWrappedGuideTopic, ByVal indentLevel As Long, ByVal hasPrefs As Boolean) As WIDGET
    Dim widget As WIDGET, topic As WzGuideTopic
    Set topic = wrappedTopic.topic
    If Not (Not topic Is Nothing) Then ASSERT_FAILURE "topic != nullptr", "src/screens/guidescreen.cpp:456", "make", "Null topic?": Exit Function
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideSidebarTopicButton)
    WzGuideSidebarTopicButton_construct widget
    WzGuideSidebarTopicButton_initialize widget, topic.identifier, topic.displayName, indentLevel
    If hasPrefs Then
        ' setIsNew(prefs.value().lastReadVersion != topic->version)
        WzGuideSidebarTopicButton_(widget).isNew = Not (wrappedTopic.prefs_has_lastReadVersion And wrappedTopic.prefs_lastReadVersion = topic.version)
    End If
    Set WzGuideSidebarTopicButton_make = widget
End Function

Public Function WzGuideSidebarTopicButton_getTip(ByVal self As WIDGET) As String
    If WzGuideSidebarTopicButton_(self).isTruncated Then WzGuideSidebarTopicButton_getTip = W_BUTTON_(self).pText
End Function

Public Sub WzGuideSidebarTopicButton_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzGuideSidebarTopicButton_(self).wzText.tick
End Sub

Public Sub WzGuideSidebarTopicButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Const indicatorSize As Long = 10
    Dim d As WzGuideSidebarTopicButton, b As W_BUTTON, x0 As Long, y0 As Long, w As Long, h As Long, x1 As Long, y1 As Long
    Dim maxTextDisplayableWidth As Long, haveText As Boolean, isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean
    Dim has_backColor As Boolean, backColor As Long, indicatorColor As Long, newIndicatorX0 As Long, newIndicatorY0 As Long
    Dim textColor As Long, indentOffsetX As Long, textX0 As Long, textY0 As Long
    Set d = WzGuideSidebarTopicButton_(self)
    Set b = W_BUTTON_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()
    x1 = x0 + w
    y1 = y0 + h

    maxTextDisplayableWidth = w - (leftPadding + rightPadding)

    haveText = Len(b.pText) > 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    has_backColor = False                           ' optional<PIELIGHT> backColor = nullopt
    If d.isSelected Then
        backColor = psPalette(WZCOL_MENU_SCORE_BUILT)
        has_backColor = True
    ElseIf isHighlight Then
        backColor = pal_RGBA(255, 255, 255, 50)
        has_backColor = True
    End If
    If has_backColor Then
        pie_UniTransBoxFill x0, y0, x1, y1, backColor
    End If

    ' Handle isNew (display a "new" blue indicator on the right side)
    If d.isNew Then
        indicatorColor = pal_RGBA(35, 82, 252, 255) 'pal_RGBA(35,252,115,255);
        newIndicatorX0 = x0 + (w - indicatorSize - rightPadding)
        newIndicatorY0 = y0 + ((h - indicatorSize) \ 2)
        If IntImages <> 0 Then
            iV_DrawImageTint IntImages, IMAGE_INDICATOR_DOT, newIndicatorX0, newIndicatorY0, indicatorColor, True, indicatorSize, indicatorSize
        Else
            ' Interface Images are not loaded - just use simple rectangles
            pie_UniTransBoxFill newIndicatorX0, newIndicatorY0, newIndicatorX0 + indicatorSize, newIndicatorY0 + indicatorSize, indicatorColor
        End If

        maxTextDisplayableWidth = maxTextDisplayableWidth - (indicatorSize + rightPadding)
    End If

    If haveText Then
        d.wzText.setText b.pText, b.FontID

        If isDown Or d.isSelected Or isHighlight Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_FORM_TEXT)
        If isDisabled Then
            textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
        End If

        indentOffsetX = (indentPadding * d.indentLevel)
        textX0 = x0 + leftPadding + indentOffsetX
        textY0 = y0 + (h - d.wzText.deref().lineSize()) \ 2 - d.wzText.deref().aboveBase()
        maxTextDisplayableWidth = maxTextDisplayableWidth - indentOffsetX

        d.isTruncated = maxTextDisplayableWidth < d.wzText.deref().width()
        If d.isTruncated Then
            maxTextDisplayableWidth = maxTextDisplayableWidth - (iV_GetEllipsisWidth(b.FontID) + 2)
        End If
        d.wzText.deref().render textX0, textY0, textColor, 0!, maxTextDisplayableWidth
        If d.isTruncated Then
            ' Render ellipsis
            iV_DrawEllipsis b.FontID, textX0 + maxTextDisplayableWidth + 2, textY0, textColor
        End If
    End If
End Sub

' (VB) m_topicIdentifierToListIdMap.find(topicIdentifier): the index, -1 for end()
Private Function sidebarListId(ByVal d As WzGuideSidebar, ByVal topicIdentifier As String) As Long
    sidebarListId = -1
    On Error Resume Next
    sidebarListId = d.m_topicIdentifierToListIdMap.Item("k" & topicIdentifier)
    Err.Clear
End Function

Private Sub WzGuideSidebar_changeSelectedItem(ByVal self As WIDGET, ByVal topicIdentifier As String)
    Dim d As WzGuideSidebar, it As Long, but As WIDGET
    Set d = WzGuideSidebar_(self)
    it = sidebarListId(d, topicIdentifier)
    If it = -1 Then
        d.m_currentSelectedTopicIdentifier = ""
        Exit Sub
    End If
    If Not (it < d.m_topicButtons.Count) Then ASSERT_FAILURE "it->second < m_topicButtons.size()", "src/screens/guidescreen.cpp:2054", "changeSelectedItem", "Invalid topic identifier mapping: %zu", it: Exit Sub
    For Each but In d.m_topicButtons
        WzGuideSidebarTopicButton_(but).isSelected = False          ' but->setIsSelected(false)
    Next
    WzGuideSidebarTopicButton_(d.m_topicButtons(it + 1)).isSelected = True
    WzGuideSidebarTopicButton_(d.m_topicButtons(it + 1)).isNew = False
    d.m_currentSelectedTopicIdentifier = topicIdentifier
    If self.width() > 0 And self.height() > 0 Then
        ScrollableListWidget_scrollEnsureItemVisible d.m_topicsList, it
    End If
End Sub

' optional<std::string> getItemTopicIdentifier(size_t idx): has_value, the value in result
Private Function WzGuideSidebar_getItemTopicIdentifier(ByVal self As WIDGET, ByVal idx As Long, result As String) As Boolean
    Dim d As WzGuideSidebar
    Set d = WzGuideSidebar_(self)
    If idx >= d.m_topicButtons.Count Then
        Exit Function
    End If
    result = WzGuideSidebarTopicButton_(d.m_topicButtons(idx + 1)).topicIdentifier
    WzGuideSidebar_getItemTopicIdentifier = True
End Function

' (VB) makeSortedNestedTopics' comparator: a < b
Private Function guideTopicLess(ByVal a As WzWrappedGuideTopic, ByVal b As WzWrappedGuideTopic) As Boolean
    Dim aHasSortKey As Boolean, bHasSortKey As Boolean
    aHasSortKey = a.topic.has_sortKey
    bHasSortKey = b.topic.has_sortKey
    If aHasSortKey <> bHasSortKey Then
        guideTopicLess = aHasSortKey
        Exit Function
    End If
    If aHasSortKey Then
        ' std::locale::classic()'s collate: the characters' order
        guideTopicLess = StrComp(a.topic.sortKey, b.topic.sortKey, vbBinaryCompare) < 0
    Else
        ' std::locale()'s collate: the global C++ locale is the classic one (4.7.0 never sets another)
        guideTopicLess = StrComp(a.topic.displayName, b.topic.displayName, vbBinaryCompare) < 0
    End If
End Function

Private Function makeSortedNestedTopics(ByVal topics As Collection) As Collection
    Dim sortedTopics As Collection, wrappedTopic As WzWrappedGuideTopic, childSortedWrappedTopicCopy As WzWrappedGuideTopic
    Dim items() As WzWrappedGuideTopic, n As Long, i As Long, j As Long, t As WzWrappedGuideTopic
    n = topics.Count
    If n > 0 Then ReDim items(1 To n)
    For Each wrappedTopic In topics
        i = i + 1
        If wrappedTopic.children.Count = 0 Then
            Set items(i) = wrappedTopic
        Else
            ' std::make_shared<WzWrappedGuideTopic>(*wrappedTopic): a copy, its children sorted
            Set childSortedWrappedTopicCopy = New WzWrappedGuideTopic
            Set childSortedWrappedTopicCopy.topic = wrappedTopic.topic
            childSortedWrappedTopicCopy.parentId = wrappedTopic.parentId
            childSortedWrappedTopicCopy.has_prefs = wrappedTopic.has_prefs
            childSortedWrappedTopicCopy.prefs_has_lastReadVersion = wrappedTopic.prefs_has_lastReadVersion
            childSortedWrappedTopicCopy.prefs_lastReadVersion = wrappedTopic.prefs_lastReadVersion
            childSortedWrappedTopicCopy.prefs_bookmarked = wrappedTopic.prefs_bookmarked
            Set childSortedWrappedTopicCopy.children = makeSortedNestedTopics(wrappedTopic.children)
            Set items(i) = childSortedWrappedTopicCopy
        End If
    Next

    ' sort the topics list at this level (std::sort; (VB) an insertion sort: the keys differ)
    For i = 2 To n
        Set t = items(i)
        j = i - 1
        Do While j >= 1
            If Not guideTopicLess(t, items(j)) Then Exit Do
            Set items(j + 1) = items(j)
            j = j - 1
        Loop
        Set items(j + 1) = t
    Next

    Set sortedTopics = New Collection
    For i = 1 To n
        sortedTopics.Add items(i)
    Next
    Set makeSortedNestedTopics = sortedTopics
End Function

Private Function WzGuideSidebar_makeTopicWidgets(ByVal self As WIDGET, ByVal topics As Collection, Optional ByVal level As Long = 0) As Collection
    Dim result As Collection, wrappedTopic As WzWrappedGuideTopic, button As WIDGET, childButtons As Collection, child As WIDGET
    Dim hasPrefs As Boolean
    Set result = New Collection
    For Each wrappedTopic In topics
        If Not wrappedTopic.topic Is Nothing Then
            ' create button for this topic
            hasPrefs = WzGuideTopicsRegistry_getTopicPrefs(wrappedTopic)
            Set button = WzGuideSidebarTopicButton_make(wrappedTopic, level, hasPrefs)
            If button Is Nothing Then
                ASSERT_FAILURE "false", "src/screens/guidescreen.cpp:2136", "makeTopicWidgets", "Failed to create sidebar topic button?"
            Else
                ' (std::weak_ptr<WzGuideSidebar> weakParent: the sidebar's serial)
                W_BUTTON_addOnClickHandler button, WzFunc_new(FN_guide_sidebarButton_click, WzGuideSidebar_(self).weakSerial)
                result.Add button

                ' add buttons for children
                If wrappedTopic.children.Count > 0 Then
                    Set childButtons = WzGuideSidebar_makeTopicWidgets(self, wrappedTopic.children, level + 1)
                    For Each child In childButtons
                        result.Add child
                    Next
                End If
            End If
        End If
    Next

    Set WzGuideSidebar_makeTopicWidgets = result
End Function

Private Sub WzGuideSidebar_updateList(ByVal self As WIDGET)
    Dim d As WzGuideSidebar, sortedTopics As Collection, listIdx As Long, topicButton As WIDGET, previouslySelectedTopicId As String
    Set d = WzGuideSidebar_(self)
    regInit
    ScrollableListWidget_clear d.m_topicsList
    Set d.m_topicIdentifierToListIdMap = New Collection
    Set d.m_topicButtons = New Collection

    ' sort the topics at each level
    Set sortedTopics = makeSortedNestedTopics(regTopics)

    Set d.m_topicButtons = WzGuideSidebar_makeTopicWidgets(self, sortedTopics)

    listIdx = 0
    For Each topicButton In d.m_topicButtons
        ' insert({identifier, listIdx}): a key already there keeps its index
        If sidebarListId(d, WzGuideSidebarTopicButton_(topicButton).topicIdentifier) = -1 Then
            d.m_topicIdentifierToListIdMap.Add listIdx, "k" & WzGuideSidebarTopicButton_(topicButton).topicIdentifier
        End If
        ScrollableListWidget_addItem d.m_topicsList, topicButton
        listIdx = listIdx + 1
    Next

    ' re-select the previously-selected item
    If Len(d.m_currentSelectedTopicIdentifier) > 0 Then
        previouslySelectedTopicId = d.m_currentSelectedTopicIdentifier
        WzGuideSidebar_changeSelectedItem self, previouslySelectedTopicId
    End If

    WzGuideSidebar_geometryChanged self
End Sub

Public Sub WzGuideSidebar_geometryChanged(ByVal self As WIDGET)
    ' (VB) not before initialize made the list (WIDGET's constructor reaches this override)
    If WzGuideSidebar_(self).m_topicsList Is Nothing Then Exit Sub
    WzGuideSidebar_(self).m_topicsList.callCalcLayout
End Sub

Private Sub WzGuideSidebar_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzGuideSidebar
    WzGuideSidebar_(self).weakSerial = WeakWidget_register(self)
End Sub

Public Sub WzGuideSidebar_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzGuideSidebar_(self).weakSerial
End Sub

Private Sub WzGuideSidebar_initialize(ByVal self As WIDGET, ByVal parentGuideForm As WIDGET)
    Dim d As WzGuideSidebar, p As Padding
    Set d = WzGuideSidebar_(self)
    Set d.m_topicsList = ScrollableListWidget_make()
    p.top = listTopPadding: p.right_ = 0: p.bottom = 0: p.left_ = 1
    ScrollableListWidget_setPadding d.m_topicsList, p
    self.attach d.m_topicsList
    d.m_topicsList.setCalcLayout WzFunc_new(FN_guide_topicsList_calcLayout)
    d.m_parentGuideForm = WzGuideForm_(parentGuideForm).weakSerial
End Sub

Public Sub WzGuideSidebar_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    iV_Line x1, y0, x1, y1, pal_RGBA(255, 255, 255, 100)
End Sub

Private Function WzGuideSidebar_make(ByVal parentGuideForm As WIDGET) As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideSidebar)
    WzGuideSidebar_construct widget
    WzGuideSidebar_initialize widget, parentGuideForm
    WzGuideSidebar_updateList widget
    Set WzGuideSidebar_make = widget
End Function

'--------------------------------------------------------------------------
' MARK: - WzGuideForm implementation

Public Function WzGuideForm_idealWidth(ByVal self As WIDGET) As Long
    Dim w As Long
    w = topicsDisplayWidth
    If WzGuideForm_(self).m_sidebarVisible Then
        w = w + sidebarDisplayWidth
    End If
    WzGuideForm_idealWidth = w
End Function

Public Function WzGuideForm_idealHeight(ByVal self As WIDGET) As Long
    WzGuideForm_idealHeight = MAX_(480, 200) ' TODO: Implement
End Function

Private Sub WzGuideForm_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_WzGuideForm
    WzGuideForm_(self).weakSerial = WeakWidget_register(self)
End Sub

Public Sub WzGuideForm_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzGuideForm_(self).weakSerial
End Sub

Private Function WzGuideForm_make(ByVal onCloseHandler As WzFunc) As WIDGET
    Dim widget As WIDGET
    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGuideForm)
    WzGuideForm_construct widget
    WzGuideForm_initialize widget, onCloseHandler
    Set WzGuideForm_make = widget
End Function

Private Sub WzGuideForm_initialize(ByVal self As WIDGET, ByVal onCloseHandler As WzFunc)
    Dim d As WzGuideForm
    Set d = WzGuideForm_(self)
    Set d.m_closeHandler = onCloseHandler

    W_FORM_enableMinimizing self, "Game Guide"
    W_FORM_(self).userMovable = True

    ' (std::weak_ptr<WzGuideForm> psWeakGuideForm: d.weakSerial)
    Set d.m_navBar = WzGuideTopicsNavBar_make(topicsDisplayWidth, WzFunc_new(FN_guide_form_popState, d.weakSerial), WzFunc_new(FN_guide_form_sidebarToggle, d.weakSerial), WzFunc_new(FN_guide_form_close, d.weakSerial))
    d.m_navBar.setTransparentToClicks True
    self.attach d.m_navBar

    Set d.m_sideBar = WzGuideSidebar_make(self)
    self.attach d.m_sideBar
End Sub

Private Sub WzGuideForm_toggleSidebar(ByVal self As WIDGET)
    WzGuideForm_(self).m_sidebarVisible = Not WzGuideForm_(self).m_sidebarVisible
    WzGuideForm_resizeToIdeal self
End Sub

Private Sub WzGuideForm_setSidebarVisible(ByVal self As WIDGET, ByVal visible As Boolean)
    WzGuideForm_(self).m_sidebarVisible = visible
    WzGuideForm_resizeToIdeal self
End Sub

Private Sub WzGuideForm_resizeToIdeal(ByVal self As WIDGET)
    Dim guideFormHeight As Long, guideFormWidth As Long, x0 As Long, y0 As Long, maxX1 As Long, maxY1 As Long, psParent As WIDGET
    guideFormHeight = self.idealHeight()
    guideFormWidth = self.idealWidth()

    x0 = self.x() + (self.width() - guideFormWidth)

    y0 = self.y()
    maxX1 = x0 + guideFormWidth
    maxY1 = y0 + guideFormHeight
    Set psParent = self.parent()
    If Not psParent Is Nothing Then
        maxX1 = psParent.width() - 10
        maxY1 = psParent.height()
    End If
    If x0 + guideFormWidth > maxX1 Then
        x0 = MAX_(maxX1 - guideFormWidth, 0)
    End If
    If y0 + guideFormHeight > maxY1 Then
        y0 = MAX_(maxY1 - guideFormHeight, 0)
    End If
    x0 = MAX_(x0, 0)
    y0 = MAX_(y0, 0)

    ' set new width + height, adjusting x,y if needed to keep in bounds
    self.setGeometry x0, y0, guideFormWidth, guideFormHeight
End Sub

Private Sub WzGuideForm_queueRefreshTopicsList(ByVal self As WIDGET)
    WzGuideForm_(self).m_queuedTopicsListRefresh = True
End Sub

Public Sub WzGuideForm_geometryChanged(ByVal self As WIDGET)
    Dim d As WzGuideForm, w As Long, h As Long, navbarHeight As Long, heightMinusNavbar As Long, guideTopicX0 As Long
    Dim sidebarWidth As Long
    Set d = WzGuideForm_(self)
    ' (VB) not before initialize made the nav bar (WIDGET's constructor reaches this override)
    If d.m_navBar Is Nothing Then Exit Sub
    w = self.width()
    h = self.height()

    navbarHeight = d.m_navBar.idealHeight()
    d.m_navBar.setGeometry 0, 0, w, navbarHeight

    heightMinusNavbar = MAX_(h - navbarHeight, 0)
    If heightMinusNavbar = 0 Then
        Exit Sub
    End If

    If Not d.m_guideTopic Is Nothing Then
        guideTopicX0 = MAX_(w - topicsDisplayWidth, 0)
        d.m_guideTopic.setGeometry guideTopicX0, navbarHeight, MIN_(topicsDisplayWidth, w), heightMinusNavbar
    End If

    If d.m_sidebarVisible Then
        d.m_sideBar.show
        sidebarWidth = MAX_(w - topicsDisplayWidth, 0)
        d.m_sideBar.setGeometry 0, navbarHeight, sidebarWidth, heightMinusNavbar
    Else
        d.m_sideBar.hide
    End If
End Sub

Public Sub WzGuideForm_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    WzGuideForm_updateSidebarIfNeeded self
End Sub

Private Sub WzGuideForm_updateSidebarIfNeeded(ByVal self As WIDGET)
    Dim d As WzGuideForm
    Set d = WzGuideForm_(self)
    If d.m_queuedTopicsListRefresh Then
        If Not d.m_sideBar Is Nothing Then
            WzGuideSidebar_updateList d.m_sideBar
        End If
        d.m_queuedTopicsListRefresh = False
    End If
End Sub

Public Sub WzGuideForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, dropShadowX0 As Long, dropShadowY0 As Long, dropShadowX1 As Long
    Dim dropShadowY1 As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    ' draw "drop-shadow"
    dropShadowX0 = MAX_(x0 - 6, 0)
    dropShadowY0 = MAX_(y0 - 6, 0)
    dropShadowX1 = MIN_(x1 + 6, pie_GetVideoBufferWidth())
    dropShadowY1 = MIN_(y1 + 6, pie_GetVideoBufferHeight())
    pie_UniTransBoxFill CSng(dropShadowX0), CSng(dropShadowY0), CSng(dropShadowX1), CSng(dropShadowY1), pal_RGBA(0, 0, 0, 40)

    ' draw background + border
    pie_UniTransBoxFill CSng(x0), CSng(y0), CSng(x1), CSng(y1), pal_RGBA(5, 0, 15, 170)
    iV_Box x0, y0, x1, y1, pal_RGBA(255, 255, 255, 150)
End Sub

Private Function WzGuideForm_displayGuideTopic(ByVal self As WIDGET, ByVal topicIdentifier As String, ByVal onTopicClose As WzFunc) As Boolean
    Dim d As WzGuideForm, psTopic As WzWrappedGuideTopic
    Set d = WzGuideForm_(self)
    Set psTopic = WzGuideTopicsRegistry_getTopic(topicIdentifier)
    If Not (Not psTopic Is Nothing) Then ASSERT_FAILURE "psTopic != nullptr", "src/screens/guidescreen.cpp:2417", "displayGuideTopic", "Topic not found: %s", topicIdentifier: Exit Function
    If Not (Not psTopic.topic Is Nothing) Then ASSERT_FAILURE "psTopic->topic != nullptr", "src/screens/guidescreen.cpp:2418", "displayGuideTopic", "Topic is null?: %s", topicIdentifier: Exit Function

    WzGuideTopicsRegistry_setTopicSeen psTopic

    If Not d.m_guideTopic Is Nothing Then
        If WzGuideTopicDisplayWidget_(d.m_guideTopic).m_topicID = topicIdentifier Then
            ' already displaying this topic
            WzGuideForm_displayGuideTopic = True
            Exit Function
        End If
        self.detach d.m_guideTopic
        Set d.m_guideTopic = Nothing
    End If

    Set d.m_guideTopic = WzGuideTopicDisplayWidget_make(psTopic)
    If Not (Not d.m_guideTopic Is Nothing) Then ASSERT_FAILURE "m_guideTopic != nullptr", "src/screens/guidescreen.cpp:2434", "displayGuideTopic", "Failed to create guide topic display for: %s", topicIdentifier: Exit Function
    If Not onTopicClose Is Nothing Then
        d.m_guideTopic.setOnDelete WzFunc_new(FN_guide_topic_onDelete, onTopicClose, topicIdentifier)
    End If
    self.attach d.m_guideTopic
    WzGuideForm_updateSidebarIfNeeded self
    WzGuideSidebar_changeSelectedItem d.m_sideBar, topicIdentifier
    WzGuideForm_geometryChanged self
    WzGuideForm_displayGuideTopic = True
End Function

Private Sub WzGuideForm_showGuideTopic(ByVal self As WIDGET, ByVal topicIdentifier As String, Optional ByVal onTopicClose As WzFunc = Nothing)
    Dim d As WzGuideForm
    Set d = WzGuideForm_(self)
    If Not d.m_guideTopic Is Nothing Then
        If WzGuideTopicDisplayWidget_(d.m_guideTopic).m_topicID = topicIdentifier Then
            ' already displaying this topic
            Exit Sub
        End If
    End If

    If WzGuideForm_displayGuideTopic(self, topicIdentifier, onTopicClose) Then
        WzGuideTopicsNavBar_pushState d.m_navBar, topicIdentifier
    End If
End Sub

' shows the first guide topic if not topic is currently displayed
Private Sub WzGuideForm_showFirstGuideTopicIfNeeded(ByVal self As WIDGET)
    Dim d As WzGuideForm, firstSortedTopicId As String
    Set d = WzGuideForm_(self)
    If Not d.m_guideTopic Is Nothing Then
        ' already displaying a guide topic
        Exit Sub
    End If

    ' otherwise, fetch the first guide topic in the sorted sidebar list
    WzGuideForm_updateSidebarIfNeeded self
    If WzGuideSidebar_getItemTopicIdentifier(d.m_sideBar, 0, firstSortedTopicId) Then
        If WzGuideForm_displayGuideTopic(self, firstSortedTopicId, Nothing) Then
            WzGuideTopicsNavBar_pushState d.m_navBar, firstSortedTopicId
        End If
    End If
End Sub

'--------------------------------------------------------------------------
' MARK: - WzGameGuideScreen implementation

' WzGameGuideScreen_CLICKFORM(W_FORMINIT const *init) : W_CLICKFORM(init) {}
Private Sub WzGameGuideScreen_CLICKFORM_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    Dim d As WzGameGuideScreen_CLICKFORM
    W_CLICKFORM_construct_init self, init
    self.kind = K_WzGameGuideScreen_CLICKFORM
    Set d = WzGameGuideScreen_CLICKFORM_(self)
    d.backgroundColor = pal_RGBA(0, 0, 0, 80)
    d.weakSerial = WeakWidget_register(self)
End Sub

' (~WzGameGuideScreen_CLICKFORM() { } has nothing to do but its weak serial)
Public Sub WzGameGuideScreen_CLICKFORM_destruct(ByVal self As WIDGET)
    WeakWidget_unregister WzGameGuideScreen_CLICKFORM_(self).weakSerial
End Sub

' static std::shared_ptr<WzGameGuideScreen_CLICKFORM> make(UDWORD formID = 0)
Private Function WzGameGuideScreen_CLICKFORM_make(Optional ByVal formID As Long = 0) As WIDGET
    Dim sInit As W_FORMINIT, widget As WIDGET
    W_FORMINIT_construct sInit
    sInit.id = formID
    sInit.style = WFORM_PLAIN Or WFORM_CLICKABLE
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1
    Set sInit.calcLayout = WzFunc_new(FN_guide_clickForm_calcLayout)

    ' (class make_shared_enabler)
    Set widget = WIDGET_alloc(K_WzGameGuideScreen_CLICKFORM)
    WzGameGuideScreen_CLICKFORM_construct_init widget, sInit
    WzGameGuideScreen_CLICKFORM_initialize widget
    Set WzGameGuideScreen_CLICKFORM_make = widget
End Function

Private Sub WzGameGuideScreen_CLICKFORM_initialize(ByVal self As WIDGET)
    Dim d As WzGameGuideScreen_CLICKFORM, guideFormHeight As Long, guideFormWidth As Long, x0 As Long, y0 As Long
    Set d = WzGameGuideScreen_CLICKFORM_(self)
    ' (std::weak_ptr<WzGameGuideScreen_CLICKFORM> weakSelf: d.weakSerial)

    Set d.guideForm = WzGuideForm_make(WzFunc_new(FN_guide_clickForm_guideFormClose, d.weakSerial))
    self.attach d.guideForm
    ' initially position right-aligned in parent (which takes up the whole screen)
    guideFormHeight = d.guideForm.idealHeight()
    guideFormWidth = d.guideForm.idealWidth()
    x0 = (screenWidth - guideFormWidth - d.edgeMargin)
    y0 = (screenHeight - guideFormHeight) \ 2
    d.guideForm.setGeometry x0, y0, guideFormWidth, guideFormHeight
End Sub

Public Sub WzGameGuideScreen_CLICKFORM_geometryChanged(ByVal self As WIDGET)
    ' (VB) not before initialize made the guide form (WIDGET's constructor runs the calcLayout)
    If WzGameGuideScreen_CLICKFORM_(self).guideForm Is Nothing Then Exit Sub
    WzGameGuideScreen_CLICKFORM_recalcLayout self
End Sub

Private Sub WzGameGuideScreen_CLICKFORM_recalcLayout(ByVal self As WIDGET)
    Dim d As WzGameGuideScreen_CLICKFORM, w As Long, h As Long, guideFormHeight As Long, guideFormWidth As Long
    Dim x0 As Long, y0 As Long
    Set d = WzGameGuideScreen_CLICKFORM_(self)
    w = self.width()
    h = self.height()

    ' ensure guideForm is within window bounds
    guideFormHeight = d.guideForm.idealHeight()
    guideFormWidth = d.guideForm.idealWidth()

    x0 = clip_(d.guideForm.x(), d.edgeMargin, w - guideFormWidth - d.edgeMargin)
    y0 = MIN_(d.guideForm.y(), h - guideFormHeight)

    d.guideForm.setGeometry x0, y0, guideFormWidth, guideFormHeight
End Sub

Private Sub WzGameGuideScreen_CLICKFORM_clearFocusAndEditing(ByVal self As WIDGET)
    Dim lockedScreen As W_SCREEN
    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        lockedScreen.setFocus Nothing
    End If
End Sub

Private Sub WzGameGuideScreen_CLICKFORM_handleClickOnForm(ByVal self As WIDGET)
    Dim d As WzGameGuideScreen_CLICKFORM
    Set d = WzGameGuideScreen_CLICKFORM_(self)
    WzGameGuideScreen_CLICKFORM_clearFocusAndEditing self
    If Not d.onClickedFunc Is Nothing Then
        WzFunc_call d.onClickedFunc
    End If
End Sub

Public Sub WzGameGuideScreen_CLICKFORM_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    WzGameGuideScreen_CLICKFORM_handleClickOnForm self
End Sub

Public Sub WzGameGuideScreen_CLICKFORM_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzGameGuideScreen_CLICKFORM, x0 As Long, y0 As Long
    Set d = WzGameGuideScreen_CLICKFORM_(self)
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

Public Sub WzGameGuideScreen_CLICKFORM_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As WzGameGuideScreen_CLICKFORM
    Set d = WzGameGuideScreen_CLICKFORM_(self)
    If keyPressed(KEY_ESC) Then
        WzGameGuideScreen_CLICKFORM_clearFocusAndEditing self
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

' (~WzGameGuideScreen()'s setFocus(nullptr): W_SCREEN's Class_Terminate resigns the focus)
' static std::shared_ptr<WzGameGuideScreen> make(const OnCloseFunc& _onCloseFunc, bool expandedSidebarByDefault)
Private Function WzGameGuideScreen_make(ByVal onCloseFunc_ As WzFunc, ByVal expandedSidebarByDefault As Boolean) As W_SCREEN
    Dim newRootFrm As WIDGET, screen As W_SCREEN
    ' (class make_shared_enabler)
    Set newRootFrm = WzGameGuideScreen_CLICKFORM_make()
    Set screen = New W_SCREEN
    ' (VB) WzGameGuideScreen(): W_SCREEN() and its members' initialisers
    Set guideScreenOnCloseFunc = Nothing
    guideScreenOpenTime = 0
    guideScreenDidStopGameTime = False
    guideScreenDidSetPauseStates = False
    screen.initialize newRootFrm
    Set guideScreenOnCloseFunc = onCloseFunc_
    ' (std::weak_ptr<WzGameGuideScreen> psWeakOverlayScreen: the screen's address, see lockGuideScreen)
    Set WzGameGuideScreen_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_guide_screen_onClicked, ObjPtr(screen))
    Set WzGameGuideScreen_CLICKFORM_(newRootFrm).onCancelPressed = WzFunc_new(FN_guide_screen_onCancelPressed, ObjPtr(screen))

    If expandedSidebarByDefault Then
        WzGuideForm_setSidebarVisible WzGameGuideScreen_getGuideForm(screen), True
    End If

    WzGameGuideScreen_pauseGameIfNeeded screen
    guideScreenOpenTime = widgSteadyNowMs()

    Set WzGameGuideScreen_make = screen
End Function

Private Sub WzGameGuideScreen_reopeningScreen(ByVal self As W_SCREEN, ByVal onCloseFunction As WzFunc, ByVal expandedSidebarByDefault As Boolean)
    Set guideScreenOnCloseFunc = onCloseFunction
    WzGameGuideScreen_pauseGameIfNeeded self
    guideScreenOpenTime = widgSteadyNowMs()
    If expandedSidebarByDefault Then
        WzGuideForm_setSidebarVisible WzGameGuideScreen_getGuideForm(self), True
    End If
End Sub

Private Sub WzGameGuideScreen_closeScreen(ByVal self As W_SCREEN, Optional ByVal force As Boolean = False)
    Dim onCloseFunc As WzFunc
    If Not force And (widgSteadyNowMs() - guideScreenOpenTime < 2000#) Then
        ' ignore errant clicks that might otherwise close screen right after it opened
        Exit Sub
    End If
    widgRemoveOverlayScreen self
    WzGameGuideScreen_unpauseGameIfNeeded self
    Set onCloseFunc = guideScreenOnCloseFunc
    If Not onCloseFunc Is Nothing Then
        ' "consumes" the onCloseFunc ((VB) before the call, which may reopen the screen with another)
        Set guideScreenOnCloseFunc = Nothing
        WzFunc_call onCloseFunc
    End If
End Sub

Private Sub WzGameGuideScreen_pauseGameIfNeeded(ByVal self As W_SCREEN)
    If Not runningMultiplayer() And Not gameTimeIsStopped() Then
        ' Stop the clock / simulation
        gameTimeStop
        addConsoleMessage "PAUSED", CENTRE_JUSTIFY, SYSTEM_MESSAGE, False, MAX_CONSOLE_MESSAGE_DURATION

        self.psForm.show ' show the root form, which means it accepts all clicks on the screen and dims the background

        guideScreenDidStopGameTime = True

        If Not gamePaused() Then
            setGamePauseStatus True
            setConsolePause True
            setScriptPause True
            setAudioPause True
            setScrollPause True
            guideScreenDidSetPauseStates = True
        End If
    Else
        self.psForm.hide ' hide the root form - it won't accept clicks (but its children will be displayed)
    End If
End Sub

Private Sub WzGameGuideScreen_unpauseGameIfNeeded(ByVal self As W_SCREEN)
    If guideScreenDidStopGameTime Then
        If Not runningMultiplayer() And gameTimeIsStopped() Then
            clearActiveConsole
            ' Start the clock again
            gameTimeStart

            'put any widgets back on for the missions
            resetMissionWidgets
        End If
        guideScreenDidStopGameTime = False
    End If

    If guideScreenDidSetPauseStates Then
        setGamePauseStatus False
        setConsolePause False
        setScriptPause False
        setAudioPause False
        setScrollPause False
        guideScreenDidSetPauseStates = False
    End If
End Sub

Private Function WzGameGuideScreen_getGuideForm(ByVal self As W_SCREEN) As WIDGET
    Set WzGameGuideScreen_getGuideForm = WzGameGuideScreen_CLICKFORM_(self.psForm).guideForm
End Function

'--------------------------------------------------------------------------
' MARK: - Loading topic from JSON

' void from_json(const nlohmann::json& j, WzStatVisual::OnlyIfKnown& v)
Private Function OnlyIfKnown_from_json(ByVal j As nlohmann_json) As Long
    Dim valStr As String
    If j.is_boolean() Then
        If j.get_bool() Then OnlyIfKnown_from_json = OnlyIfKnown_True Else OnlyIfKnown_from_json = OnlyIfKnown_False
    ElseIf j.is_string() Then
        valStr = j.get_string()
        If valStr = "false" Then
            OnlyIfKnown_from_json = OnlyIfKnown_False
        ElseIf valStr = "true" Then
            OnlyIfKnown_from_json = OnlyIfKnown_True
        ElseIf valStr = "campaign" Then
            OnlyIfKnown_from_json = OnlyIfKnown_CampaignMode
        Else
            guideJsonThrow "other_error", 600, "Unknown value (""" & valStr & """)"
        End If
    Else
        guideJsonThrow "type_error", 302, "invalid type of value: " & j.type_name()
    End If
End Function

' void from_json(const nlohmann::json& j, WzStatVisual& v): the visual's statVisual fields
Private Sub WzStatVisual_from_json(ByVal j As nlohmann_json, v As WzGuideVisual)
    Dim it As nlohmann_json
    If j.is_object() Then
        v.statName = j.at("id").get_string()
        Set it = j.find("onlyIfKnown")
        If Not it Is Nothing Then
            v.onlyIfKnown = OnlyIfKnown_from_json(it)
        End If
    ElseIf j.is_string() Then
        ' special-case: treat string value as just the id
        v.statName = j.get_string()
    Else
        guideJsonThrow "type_error", 302, "invalid type of value for ""stat"": " & j.type_name()
    End If
End Sub

' void from_json(const nlohmann::json& j, ActivitySink::GameMode& v) (strToGameModeMap)
Private Function GameMode_from_json(ByVal j As nlohmann_json) As Long
    Dim strVal As String
    If Not j.is_string() Then
        guideJsonThrow "type_error", 302, "invalid type of value - expecting string, got: " & j.type_name()
    End If

    strVal = j.get_string()
    Select Case strVal
    Case "campaign": GameMode_from_json = ActivitySink_GameMode_CAMPAIGN
    Case "challenge": GameMode_from_json = ActivitySink_GameMode_CHALLENGE
    Case "skirmish": GameMode_from_json = ActivitySink_GameMode_SKIRMISH
    Case "multiplayer": GameMode_from_json = ActivitySink_GameMode_MULTIPLAYER
    Case Else
        guideJsonThrow "other_error", 600, "Unknown value (""" & strVal & """)"
    End Select
End Function

' void from_json(const nlohmann::json& j, SECONDARY_ORDER& v) (strToSecondaryOrderMap)
Private Function SECONDARY_ORDER_from_json(ByVal j As nlohmann_json) As Long
    Dim strVal As String
    If Not j.is_string() Then
        guideJsonThrow "type_error", 302, "invalid type of value - expecting string, got: " & j.type_name()
    End If

    strVal = j.get_string()
    Select Case strVal
    Case "ATTACK_RANGE": SECONDARY_ORDER_from_json = DSO_ATTACK_RANGE
    Case "REPAIR_LEVEL": SECONDARY_ORDER_from_json = DSO_REPAIR_LEVEL
    Case "ATTACK_LEVEL": SECONDARY_ORDER_from_json = DSO_ATTACK_LEVEL
    Case "ASSIGN_PRODUCTION": SECONDARY_ORDER_from_json = DSO_ASSIGN_PRODUCTION
    Case "ASSIGN_CYBORG_PRODUCTION": SECONDARY_ORDER_from_json = DSO_ASSIGN_CYBORG_PRODUCTION
    Case "CLEAR_PRODUCTION": SECONDARY_ORDER_from_json = DSO_CLEAR_PRODUCTION
    Case "RECYCLE": SECONDARY_ORDER_from_json = DSO_RECYCLE
    Case "PATROL": SECONDARY_ORDER_from_json = DSO_PATROL
    Case "HALTTYPE": SECONDARY_ORDER_from_json = DSO_HALTTYPE
    Case "RETURN_TO_LOC": SECONDARY_ORDER_from_json = DSO_RETURN_TO_LOC
    Case "FIRE_DESIGNATOR": SECONDARY_ORDER_from_json = DSO_FIRE_DESIGNATOR
    Case "ASSIGN_VTOL_PRODUCTION": SECONDARY_ORDER_from_json = DSO_ASSIGN_VTOL_PRODUCTION
    Case "CIRCLE": SECONDARY_ORDER_from_json = DSO_CIRCLE
    Case "ACCEPT_RETREP": SECONDARY_ORDER_from_json = DSO_ACCEPT_RETREP
    Case Else
        guideJsonThrow "other_error", 600, "Unknown value (""" & strVal & """)"
    End Select
End Function

' from_json(WzGuideVisual)'s findVisualKey lambda
Private Function findVisualKey(ByVal j As nlohmann_json, ByVal key As String, haveVisualType As Boolean) As nlohmann_json
    Dim it As nlohmann_json
    Set it = j.find(key)
    If Not it Is Nothing Then
        If haveVisualType Then
            guideJsonThrow "other_error", 600, "Visual must be a single type, but found additional key: (""" & key & """)"
        End If
        haveVisualType = True
    End If
    Set findVisualKey = it
End Function

' void from_json(const nlohmann::json& j, WzGuideVisual& v)
Private Sub WzGuideVisual_from_json(ByVal j As nlohmann_json, v As WzGuideVisual)
    Dim haveVisualType As Boolean, it As nlohmann_json, buttonIdStr As String, buttonId As Long
    ' A visual must be a single type, so only allow one of the top-level keys to be specified
    haveVisualType = False

    Set it = findVisualKey(j, "stat", haveVisualType)
    If Not it Is Nothing Then
        WzStatVisual_from_json it, v
        v.has_statVisual = True
    End If
    Set it = findVisualKey(j, "droidTemplate", haveVisualType)
    If Not it Is Nothing Then
        Set v.droidTemplateJSON = it
    End If
    Set it = findVisualKey(j, "retButId", haveVisualType)
    If Not it Is Nothing Then
        If it.is_string() Then
            buttonIdStr = it.get_string()
            Select Case buttonIdStr
            Case "CANCEL": v.reticuleButtonId = RETBUT_CANCEL
            Case "FACTORY": v.reticuleButtonId = RETBUT_FACTORY
            Case "RESEARCH": v.reticuleButtonId = RETBUT_RESEARCH
            Case "BUILD": v.reticuleButtonId = RETBUT_BUILD
            Case "DESIGN": v.reticuleButtonId = RETBUT_DESIGN
            Case "INTELMAP": v.reticuleButtonId = RETBUT_INTELMAP
            Case "COMMAND": v.reticuleButtonId = RETBUT_COMMAND
            Case Else
                guideJsonThrow "other_error", 600, "Unknown value (""" & buttonIdStr & """)"
            End Select
            v.has_reticuleButtonId = True
        Else
            buttonId = it.get_int()
            If buttonId >= 0 And buttonId < NUMRETBUTS Then
                v.reticuleButtonId = buttonId
                v.has_reticuleButtonId = True
            Else
                guideJsonThrow "other_error", 600, "Unknown value (""" & CStr(buttonId) & """)"
            End If
        End If
    End If
    Set it = findVisualKey(j, "imagePath", haveVisualType)
    If Not it Is Nothing Then
        v.imagePath = it.get_string()
        v.has_imagePath = True
    End If
    Set it = findVisualKey(j, "secondaryUnitOrderButtons", haveVisualType)
    If Not it Is Nothing Then
        v.secondaryUnitOrderButtons = SECONDARY_ORDER_from_json(it)
        v.has_secondaryUnitOrderButtons = True
    End If
End Sub

Private Sub fixupVisualPaths(visual As WzGuideVisual, ByVal baseDir As String, ByVal j As nlohmann_json)
    Const parentDir As String = "../"
    Dim relativePath As String, fullImagePath As String, relStart As Long, basePos As Long
    If visual.has_imagePath And Len(visual.imagePath) > 0 Then
        relativePath = visual.imagePath

        fullImagePath = baseDir
        If strEndsWith(fullImagePath, "/") Then
            fullImagePath = Left$(fullImagePath, Len(fullImagePath) - 1)
        End If

        ' supporting "../"s at the beginning of the relativePath
        relStart = 0
        Do While Mid$(relativePath, relStart + 1, Len(parentDir)) = parentDir
            ' remove the last path component from the baseDir
            If Len(fullImagePath) = 0 Then
                ' no parent directory to traverse
                guideJsonThrow "other_error", 600, "Invalid imagePath (""" & relativePath & """), has too many ""../"""
            End If
            basePos = InStrRev(fullImagePath, "/") - 1
            If basePos = -1 Or basePos = 0 Then
                fullImagePath = ""
            Else
                fullImagePath = Left$(fullImagePath, basePos)
            End If
            relStart = relStart + 3
        Loop

        If Len(fullImagePath) > 0 And Not strEndsWith(fullImagePath, "/") Then
            fullImagePath = fullImagePath & "/"
        End If
        fullImagePath = fullImagePath & Mid$(relativePath, relStart + 1)
        visual.imagePath = fullImagePath
    End If
End Sub

' std::vector<WzGuideVisual> loadWZGuideVisualsFromJSON(const nlohmann::json& j, const std::string& baseDir): (VB) added to
' the topic's visuals
Private Sub loadWZGuideVisualsFromJSON(ByVal j As nlohmann_json, ByVal baseDir As String, ByVal topic As WzGuideTopic)
    Dim i As Long, v As WzGuideVisual, blank As WzGuideVisual
    If j.is_array() Then
        For i = 0 To j.size() - 1
            v = blank
            WzGuideVisual_from_json j.at_idx(i), v
            fixupVisualPaths v, baseDir, j.at_idx(i)
            topic.addVisual v
        Next
    ElseIf j.is_object() Then
        ' a single visual
        v = blank
        WzGuideVisual_from_json j, v
        fixupVisualPaths v, baseDir, j
        topic.addVisual v
    Else
        guideJsonThrow "type_error", 302, "type must be an object or array, but is " & j.type_name()
    End If
End Sub

' WzGuideTopic loadWZGuideTopicFromJSON(const std::string& expectedTopicId, const nlohmann::json& j, const std::string& baseDir)
' (throws as the C++ does: a VB error)
Private Function loadWZGuideTopicFromJSON(ByVal expectedTopicId As String, ByVal j As nlohmann_json, ByVal baseDir As String) As WzGuideTopic
    Dim v As WzGuideTopic, it As nlohmann_json, contents As nlohmann_json, i As Long, mode As Long
    Set v = New WzGuideTopic

    If Not j.is_object() Then
        guideJsonThrow "type_error", 302, "type must be an object, but is " & j.type_name()
    End If

    v.identifier = j.at("id").get_string()
    If v.identifier <> expectedTopicId Then
        guideJsonThrow "other_error", 600, """id"" must match expected (""" & expectedTopicId & """), but is """ & v.identifier & """"
    End If
    Set it = j.find("version")
    If Not it Is Nothing Then
        v.version = it.get_string()
    Else
        v.version = "0"
    End If
    Set it = j.find("sort")
    If Not it Is Nothing Then
        v.sortKey = it.get_string()
        v.has_sortKey = True
    End If
    v.displayName = WzJsonLocalizedString_from_json(j.at("title"))
    Set it = j.find("visual")
    If Not it Is Nothing Then
        loadWZGuideVisualsFromJSON it, baseDir, v
    End If
    ' v.contents = j.at("contents").get<std::vector<WzJsonLocalizedString>>()
    Set contents = j.at("contents")
    If Not contents.is_array() Then
        guideJsonThrow "type_error", 302, "type must be array, but is " & contents.type_name()
    End If
    For i = 0 To contents.size() - 1
        v.contents.Add WzJsonLocalizedString_from_json(contents.at_idx(i))
    Next
    Set it = j.find("links")
    If Not it Is Nothing Then
        ' it.value().get<std::vector<std::string>>()
        If Not it.is_array() Then
            guideJsonThrow "type_error", 302, "type must be array, but is " & it.type_name()
        End If
        For i = 0 To it.size() - 1
            v.linkedTopicIdentifiers.Add it.at_idx(i).get_string()
        Next
    End If
    Set it = j.find("applicableModes")
    If Not it Is Nothing Then
        ' it.value().get<std::unordered_set<ActivitySink::GameMode>>()
        If Not it.is_array() Then
            guideJsonThrow "type_error", 302, "type must be array, but is " & it.type_name()
        End If
        v.has_applicableModes = True
        For i = 0 To it.size() - 1
            mode = GameMode_from_json(it.at_idx(i))
            If (v.applicableModes And guideGameModeBit(mode)) = 0 Then
                v.applicableModes = v.applicableModes Or guideGameModeBit(mode)
                v.applicableModesCount = v.applicableModesCount + 1
            End If
        Next
    End If

    Set loadWZGuideTopicFromJSON = v
End Function

' (VB) try { result = loadWZGuideTopicFromJSON(...); } catch (const std::exception &e): false with e.what() in what
Private Function loadWZGuideTopicFromJSON_try(ByVal expectedTopicId As String, ByVal j As nlohmann_json, ByVal baseDir As String, result As WzGuideTopic, what As String) As Boolean
    On Error GoTo threw
    Set result = loadWZGuideTopicFromJSON(expectedTopicId, j, baseDir)
    loadWZGuideTopicFromJSON_try = True
    Exit Function
threw:
    what = Err.Description
    loadWZGuideTopicFromJSON_try = False
End Function

' optional<std::string> convertGuideTopicIdToExpectedJSONFilePath(const std::string& guideTopicId, bool appendJSONFileExtension = true):
' has_value, the value in result
Private Function convertGuideTopicIdToExpectedJSONFilePath(ByVal guideTopicId As String, result As String, Optional ByVal appendJSONFileExtension As Boolean = True) As Boolean
    Const delim As String = "::"
    Dim start As Long, end_ As Long
    If Not (Len(guideTopicId) > 0) Then ASSERT_FAILURE "!guideTopicId.empty()", "src/screens/guidescreen.cpp:3140", "convertGuideTopicIdToExpectedJSONFilePath", "Empty guide topic id": Exit Function
    result = "guidetopics"
    start = 0                                       ' (0-based, as the C++)
    end_ = InStr(1, guideTopicId, delim) - 1
    Do While end_ <> -1
        result = result & "/" & Mid$(guideTopicId, start + 1, end_ - start)
        start = end_ + Len(delim)
        end_ = InStr(start + 1, guideTopicId, delim) - 1
    Loop
    If start < Len(guideTopicId) Then
        result = result & "/" & Mid$(guideTopicId, start + 1)
    End If
    If appendJSONFileExtension Then
        result = result & ".json"
    End If
    convertGuideTopicIdToExpectedJSONFilePath = True
End Function

' (VB) WzPathInfo::fromPlatformIndependentPath(file).path(): the file's path without the file name
Private Function pathInfoPath(ByVal file As String) As String
    Dim lastSlashPos As Long
    lastSlashPos = InStrRev(file, "/") - 1
    If lastSlashPos = -1 Then
        ' no slash - return "." as the path
        pathInfoPath = "."
        Exit Function
    End If
    pathInfoPath = Left$(file, MAX_(lastSlashPos, 1))
End Function

Private Function loadWZGuideTopicByTopicId(ByVal guideTopicId As String) As WzGuideTopic
    Dim filePath As String, pathInfo As String, jsonObj As nlohmann_json, result As WzGuideTopic, what As String
    ' Convert the guide topic id to an expected virtual filesystem path
    ' Example: "wz2100::hq" should map to "guidetopics/wz2100/hq.json"
    If Not convertGuideTopicIdToExpectedJSONFilePath(guideTopicId, filePath) Then
        Exit Function
    End If
    pathInfo = pathInfoPath(filePath)               ' WzPathInfo pathInfo = WzPathInfo::fromPlatformIndependentPath(filePath.value())

    Set jsonObj = wzLoadJsonObjectFromFile(filePath)
    If jsonObj Is Nothing Then
        ' try to load "_index.json" for this topic id (in case it's a category topic)
        convertGuideTopicIdToExpectedJSONFilePath guideTopicId & "::_index", filePath
        Set jsonObj = wzLoadJsonObjectFromFile(filePath)
        If jsonObj Is Nothing Then
            Exit Function
        End If
    End If

    ' from_json
    If Not loadWZGuideTopicFromJSON_try(guideTopicId, jsonObj, pathInfo, result, what) Then
        ' Failed to parse the JSON
        debug_ 3192, LOG_ERROR, "loadWZGuideTopicByTopicId", "JSON document from %s is invalid: %s", filePath, what
        Exit Function
    End If
    Set loadWZGuideTopicByTopicId = result
End Function

'--------------------------------------------------------------------------
' MARK: - Private API

Private Function createGuideScreen(ByVal onCloseFunc As WzFunc, Optional ByVal expandedSidebarByDefault As Boolean = False) As W_SCREEN
    Dim screen As W_SCREEN
    Set screen = WzGameGuideScreen_make(onCloseFunc, expandedSidebarByDefault)
    If Not (Not screen Is Nothing) Then ASSERT_FAILURE "screen != nullptr", "src/screens/guidescreen.cpp:3203", "createGuideScreen", "Failed to create screen": Exit Function
    widgRegisterOverlayScreenOnTopOfScreen screen, psWScreen
    Set psCurrentGameGuideScreen = screen
    Set createGuideScreen = screen
End Function

Private Function createOrGetGuideScreen(ByVal onCloseFunc As WzFunc, Optional ByVal expandedSidebarByDefault As Boolean = False) As W_SCREEN
    If psCurrentGameGuideScreen Is Nothing Then
        Set psCurrentGameGuideScreen = createGuideScreen(onCloseFunc, expandedSidebarByDefault)
        If Not (Not psCurrentGameGuideScreen Is Nothing) Then ASSERT_FAILURE "psCurrentGameGuideScreen != nullptr", "src/screens/guidescreen.cpp:3214", "createOrGetGuideScreen", "Failed to create game guide screen?": Exit Function
    Else
        If Not isRegisteredOverlayScreen(psCurrentGameGuideScreen) Then
            widgRegisterOverlayScreenOnTopOfScreen psCurrentGameGuideScreen, psWScreen
            ' manually trigger a screenSizeDidChange event to ensure the existing screen's form resizes if needed
            psCurrentGameGuideScreen.screenSizeDidChange screenWidth - 1, screenHeight - 1, screenWidth, screenHeight
        End If
        WzGameGuideScreen_reopeningScreen psCurrentGameGuideScreen, onCloseFunc, expandedSidebarByDefault
    End If
    Set createOrGetGuideScreen = psCurrentGameGuideScreen
End Function

Private Function showGuideTopicImmediately(ByVal guideTopicIdentifier As String, Optional ByVal onTopicClose As WzFunc = Nothing) As Boolean
    Dim strongGameGuideScreen As W_SCREEN, psGuideForm As WIDGET
    Set strongGameGuideScreen = createOrGetGuideScreen(WzFunc_new(FN_guide_noopOnClose))
    If Not (Not strongGameGuideScreen Is Nothing) Then ASSERT_FAILURE "strongGameGuideScreen != nullptr", "src/screens/guidescreen.cpp:3234", "showGuideTopicImmediately", "Failed to open game guide screen?": Exit Function
    Set psGuideForm = WzGameGuideScreen_getGuideForm(strongGameGuideScreen)
    If Not (Not psGuideForm Is Nothing) Then ASSERT_FAILURE "psGuideForm != nullptr", "src/screens/guidescreen.cpp:3236", "showGuideTopicImmediately", "No guide form?": Exit Function
    WzGuideForm_setSidebarVisible psGuideForm, False
    WzGuideForm_showGuideTopic psGuideForm, guideTopicIdentifier, onTopicClose
    showGuideTopicImmediately = True
End Function

Private Function handleShowFlags(ByVal addedTopic As WzWrappedGuideTopic, ByVal showFlags As Long, ByVal newlyAddedTopic As Boolean, ByVal onTopicClose As WzFunc) As Boolean
    Dim shouldShow As Boolean
    If showFlags = ShowTopicFlags_None Then
        handleShowFlags = True
        Exit Function
    End If

    shouldShow = True
    If (showFlags And ShowTopicFlags_FirstAdd) <> 0 Then
        If Not newlyAddedTopic Then
            shouldShow = False
        End If
    End If
    If (showFlags And ShowTopicFlags_NeverViewed) <> 0 Then
        If WzGuideTopicsRegistry_getTopicPrefs(addedTopic) Then
            If addedTopic.prefs_has_lastReadVersion And addedTopic.prefs_lastReadVersion = addedTopic.topic.version Then
                shouldShow = False
            End If
        End If
    End If

    If shouldShow Then
        ' (VB) no topic (one its game mode skipped): C++ dereferences nullptr here
        If addedTopic Is Nothing Then Exit Function
        If Not disableTopicPopups Then
            showGuideTopicImmediately addedTopic.topic.identifier, onTopicClose
        End If
        handleShowFlags = True
    Else
        handleShowFlags = False
    End If
End Function

' bool addGuideTopic(const std::shared_ptr<WzGuideTopic>& topic, ShowTopicFlags showFlags, OnTopicCloseFunc onTopicClose = nullptr)
Private Function addGuideTopic_topic(ByVal topic As WzGuideTopic, ByVal showFlags As Long, ByVal onTopicClose As WzFunc) As Boolean
    Dim newlyAddedWrappedTopic As WzWrappedGuideTopic
    Set newlyAddedWrappedTopic = WzGuideTopicsRegistry_addTopicImpl(topic)

    If Not newlyAddedWrappedTopic Is Nothing And Not psCurrentGameGuideScreen Is Nothing Then
        WzGuideForm_queueRefreshTopicsList WzGameGuideScreen_getGuideForm(psCurrentGameGuideScreen)
    End If

    If showFlags <> ShowTopicFlags_None Then
        If Not newlyAddedWrappedTopic Is Nothing Then
            handleShowFlags newlyAddedWrappedTopic, showFlags, True, onTopicClose
        Else
            handleShowFlags WzGuideTopicsRegistry_getTopic(topic.identifier), showFlags, False, onTopicClose
        End If
    End If

    addGuideTopic_topic = Not newlyAddedWrappedTopic Is Nothing
End Function

' std::string guideTopicFilePathToTopicId(const WzPathInfo& pathInfo): (VB) the path
Private Function guideTopicFilePathToTopicId(ByVal filePath As String) As String
    Dim components() As String, result As String, idx As Long
    components = Split(filePath, "/")               ' pathInfo.pathComponents()
    If Not (UBound(components) >= 0) Then ASSERT_FAILURE "!components.empty()", "src/screens/guidescreen.cpp:3307", "guideTopicFilePathToTopicId", "Empty path?": Exit Function
    idx = 0
    If components(0) = "guidetopics" Then
        idx = idx + 1
    End If
    For idx = idx To UBound(components)
        If Len(result) > 0 Then
            result = result & "::"
        End If
        result = result & components(idx)
    Next
    If strEndsWith(result, ".json") Then
        result = Left$(result, Len(result) - 5)
    End If
    guideTopicFilePathToTopicId = result
End Function

' lib/framework/physfs_ext.cpp's WZ_PHYSFS_enumerateFilesEx_Impl, (VB) its enumFunc collecting the relative paths in
' order into files (loadWildcardTopicIds' lambda never stops the enumeration, and what it does changes no listing)
Private Sub enumerateGuideTopicFiles(ByVal baseDir As String, ByVal dirPath As String, ByVal recurse As Boolean, files As StringVector)
    Dim fullEnumPath As String, names() As String, n As Long, i As Long, fullPath As String
    fullEnumPath = baseDir
    If Len(dirPath) > 0 Then
        fullEnumPath = fullEnumPath & "/" & dirPath
    End If
    If Not WZ_PHYSFS_enumerateFiles(fullEnumPath, names, n) Then
        Exit Sub                                    ' (the error is logged there)
    End If
    If Not strEndsWith(fullEnumPath, "/") Then
        fullEnumPath = fullEnumPath & "/"
    End If
    If Len(dirPath) > 0 And Not strEndsWith(dirPath, "/") Then
        dirPath = dirPath & "/"
    End If
    For i = 0 To n - 1
        fullPath = fullEnumPath & names(i)
        If WZ_PHYSFS_isDirectory(fullPath) <> 0 Then
            If recurse Then
                enumerateGuideTopicFiles baseDir, dirPath & names(i), recurse, files
            End If
        Else
            StringVector_push_back files, dirPath & names(i)
        End If
    Next
End Sub

Private Function loadWildcardTopicIds(ByVal guideTopicId As String, ByVal showFlags As Long, ByVal onTopicClose As WzFunc, excludedTopicIDs As StringVector) As Boolean
    Const jsonExtension As String = ".json"
    Dim numAsterisks As Long, recurse As Boolean, guideTopicBaseId As String, baseDir As String, files As StringVector
    Dim i As Long, fullPath As String, expectedGuideTopicId As String, jsonObj As nlohmann_json, loadedTopic As WzGuideTopic
    Dim what As String, newlyAddedWrappedTopic As WzWrappedGuideTopic, newlyAddedWrappedTopics As Collection
    Dim topic As WzWrappedGuideTopic
    If Len(guideTopicId) = 0 Then
        Exit Function
    End If
    If Right$(guideTopicId, 1) <> "*" Then
        Exit Function
    End If

    numAsterisks = 1
    If Len(guideTopicId) >= 2 Then
        If Mid$(guideTopicId, Len(guideTopicId) - 1, 1) = "*" Then
            numAsterisks = 2
        End If
    End If
    recurse = numAsterisks > 1

    ' otherwise, see if this is a proper wildcard guideTopicId, and try to enumerate
    ' this currently only supports the asterisk in the last position, as in: "wz2100:*"
    guideTopicBaseId = Left$(guideTopicId, Len(guideTopicId) - numAsterisks)
    If Not (Len(guideTopicBaseId) > 2 And strEndsWith(guideTopicBaseId, "::")) Then ASSERT_FAILURE "guideTopicBaseId.size() > 2 && strEndsWith(guideTopicBaseId, ""::"")", "src/screens/guidescreen.cpp:3346", "loadWildcardTopicIds", "Unexpected wildcard guide topic id: %s", guideTopicId: Exit Function
    If Not convertGuideTopicIdToExpectedJSONFilePath(guideTopicBaseId, baseDir, False) Then
        Exit Function
    End If
    If Not strEndsWith(baseDir, "/") Then
        baseDir = baseDir & "/"
    End If
    Set newlyAddedWrappedTopics = New Collection
    enumerateGuideTopicFiles baseDir, "", recurse, files
    For i = 0 To files.size - 1
        fullPath = baseDir & files.items(i)
        ' check if ends with .json
        If strEndsWith(fullPath, jsonExtension) Then
            ' ignore category "_index.json" (pathInfo.fileName())
            If Mid$(fullPath, InStrRev(fullPath, "/") + 1) <> "_index.json" Then
                expectedGuideTopicId = guideTopicFilePathToTopicId(fullPath)
                If StringVector_contains(excludedTopicIDs, expectedGuideTopicId) Then
                    ' skip excluded guide topic
                Else
                    ' try to load it
                    Set jsonObj = wzLoadJsonObjectFromFile(fullPath)
                    If Not jsonObj Is Nothing Then
                        ' from_json
                        If Not loadWZGuideTopicFromJSON_try(expectedGuideTopicId, jsonObj, pathInfoPath(fullPath), loadedTopic, what) Then
                            ' Failed to parse the JSON
                            debug_ 3392, LOG_ERROR, "loadWildcardTopicIds", "JSON document from %s is invalid: %s", fullPath, what
                        Else
                            ' add to guide registry
                            Set newlyAddedWrappedTopic = WzGuideTopicsRegistry_addTopicImpl(loadedTopic)
                            If Not newlyAddedWrappedTopic Is Nothing Then
                                If Not psCurrentGameGuideScreen Is Nothing Then
                                    WzGuideForm_queueRefreshTopicsList WzGameGuideScreen_getGuideForm(psCurrentGameGuideScreen)
                                End If
                                newlyAddedWrappedTopics.Add newlyAddedWrappedTopic
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next

    If showFlags <> ShowTopicFlags_None And newlyAddedWrappedTopics.Count > 0 And Not disableTopicPopups Then
        ' Just highlight the first new entry added (that matches the showFlags), and show it
        For Each topic In newlyAddedWrappedTopics
            If handleShowFlags(topic, showFlags, True, onTopicClose) Then
                Exit For
            End If
        Next
    End If
    loadWildcardTopicIds = True
End Function

Private Function addGuideTopicImpl(ByVal guideTopicId As String, ByVal showFlags As Long, ByVal onTopicClose As WzFunc, excludedTopicIDs As StringVector) As Boolean
    Dim psExistingAddedTopic As WzWrappedGuideTopic, loadedTopic As WzGuideTopic
    Set psExistingAddedTopic = WzGuideTopicsRegistry_getTopic(guideTopicId)
    If Not psExistingAddedTopic Is Nothing Then
        debug_ 3429, LOG_WZ, "addGuideTopicImpl", "Guide topic with identifier ""%s"" was already added", guideTopicId
        handleShowFlags psExistingAddedTopic, showFlags, False, onTopicClose
        Exit Function
    End If
    If loadWildcardTopicIds(guideTopicId, showFlags, onTopicClose, excludedTopicIDs) Then
        ' already handled
        addGuideTopicImpl = True
        Exit Function
    End If
    If Not (excludedTopicIDs.size = 0) Then ASSERT_FAILURE "excludedTopicIDs.empty()", "src/screens/guidescreen.cpp:3438", "addGuideTopicImpl", "excludedTopicIDs provided, but did not specify a wildcard topic ID to add"
    Set loadedTopic = loadWZGuideTopicByTopicId(guideTopicId)
    If loadedTopic Is Nothing Then
        debug_ 3442, LOG_ERROR, "addGuideTopicImpl", "Failed to open / load guide topic: %s", guideTopicId
        Exit Function
    End If
    addGuideTopicImpl = addGuideTopic_topic(loadedTopic, showFlags, onTopicClose)
End Function

'--------------------------------------------------------------------------
' MARK: - Public API

Public Function hasGameGuideTopics() As Boolean
    regInit
    hasGameGuideTopics = regTopics.Count > 0       ' guideTopicsRegistry->hasTopics()
End Function

Public Function showGuideScreen(ByVal onCloseFunc As WzFunc, ByVal expandedSidebarByDefault As Boolean) As Boolean
    Dim screen As W_SCREEN
    Set screen = createOrGetGuideScreen(onCloseFunc, expandedSidebarByDefault)
    WzGuideForm_showFirstGuideTopicIfNeeded WzGameGuideScreen_getGuideForm(screen)
    showGuideScreen = Not screen Is Nothing
End Function

Public Sub closeGuideScreen()
    If Not psCurrentGameGuideScreen Is Nothing Then
        WzGameGuideScreen_closeScreen psCurrentGameGuideScreen, True
    End If
End Sub

' void addGuideTopic(const std::string& guideTopicId, ShowTopicFlags showFlags = ShowTopicFlags::None, ...): the overload
' with excluded topics is addGuideTopic_excluding (the set as a StringVector)
Public Sub addGuideTopic(ByVal guideTopicId As String, Optional ByVal showFlags As Long = 0)
    Dim excludedTopicIDs As StringVector
    addGuideTopicImpl guideTopicId, showFlags, Nothing, excludedTopicIDs
End Sub

Public Sub addGuideTopic_excluding(ByVal guideTopicId As String, ByVal showFlags As Long, excludedTopicIDs As StringVector)
    addGuideTopicImpl guideTopicId, showFlags, Nothing, excludedTopicIDs
End Sub

Public Sub shutdownGameGuide()
    WzGuideTopicsRegistry_clear
    closeGuideScreen
    Set psCurrentGameGuideScreen = Nothing
    disableTopicPopups = False
End Sub

' std::vector<std::string> saveLoadedGuideTopics(): (VB) the topics in topicIds(0 .. n - 1); returns n
Public Function saveLoadedGuideTopics(topicIds() As String) As Long
    Dim results As StringVector, i As Long
    regInit
    WzGuideTopicsRegistry_enumerateTopicsImpl regTopics, results
    If results.size > 0 Then
        ReDim topicIds(0 To results.size - 1)
        For i = 0 To results.size - 1
            topicIds(i) = results.items(i)
        Next
    End If
    saveLoadedGuideTopics = results.size
End Function

' bool restoreLoadedGuideTopics(const std::vector<std::string>& topicIds): (VB) topicIds(0 .. n - 1)
Public Function restoreLoadedGuideTopics(topicIds() As String, ByVal n As Long) As Boolean
    restoreLoadedGuideTopics = WzGuideTopicsRegistry_restoreLoadedTopics(topicIds, n)
End Function

Public Function getGameGuideDisableTopicPopups() As Boolean
    getGameGuideDisableTopicPopups = disableTopicPopups
End Function

Public Sub setGameGuideDisableTopicPopups(ByVal disablePopups As Boolean)
    disableTopicPopups = disablePopups
End Sub
