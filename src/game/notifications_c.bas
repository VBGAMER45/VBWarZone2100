Attribute VB_Name = "notifications_c"
' src/notifications.cpp, notifications.h, whole: the in-game notifications (a queue; one shown at a time, on an
' overlay screen of the highest z-order), their "do not show again" / "one time" preferences (notifications.json in
' the write directory) and the notification widget W_NOTIFICATION (its fields: W_NOTIFICATION.cls).
' (VB) class WZ_Notification is a UDT with its members flattened: stubs_ui's layout first (its callers set those),
' then isModal, onDisplay, onIgnored and the image's memory buffer. WZ_Notification_Image is largeIcon (its file
' path) or largeIcon_memoryBuffer (PNG bytes, largeIcon_memoryBufferSize of them, the array exactly that long);
' WZ_Notification_Action is action_title / action_onAction; WZ_Notification_Display_Options is
' displayOptions_identifier ("" = the default options), displayOptions_numTimesSeenBeforeDoNotShowAgainOption
' (makeIgnorable) and displayOptions_oneTime (makeOneTime). A WZ_Notification_Trigger is its timeInterval, a Long
' (0 = Immediate()). The callbacks are WzFuncs: onAction(notification), onDisplay(notification),
' onIgnored(notification), onDismissed(notification, reason); the notification argument is the queued request's
' handle (no caller reads it). class WZ_Queued_Notification (with its WZ_Notification_Status status and trigger) is
' a record of this module's pool, a request its handle (1-based; 0 = nullptr); the std::list queue is an array of
' handles in order, the unique_ptrs the handles (freeing a record is the unique_ptr's delete). class
' WZ_Notification_Preferences is this module's prefs_ variables (its one instance: notificationPrefs). The action
' buttons' DisplayNotificationButtonCache (its one field a WzText) is the WzText itself.
' Translations (_()) are not ported: the English text is used.
' (VB) widget class W_NOTIFICATION : W_FORM
Option Explicit

' notifications.h: enum class WZ_Notification_Dismissal_Reason
Public Enum WZ_Notification_Dismissal_Reason
    ' the user pressed the "Dismiss" button (or dragged the notification upwards)
    WZ_Notification_Dismissal_Reason_USER_DISMISSED
    ' the notification was dismissed because the "Action" button was pressed
    WZ_Notification_Dismissal_Reason_ACTION_BUTTON_CLICK
    ' the notification was cancelled or dismissed by one of the `cancelOrDismissNotification*()` functions (or similar)
    WZ_Notification_Dismissal_Reason_PROGRAMMATIC
End Enum

' notifications.h: enum class NotificationScope
Public Enum NotificationScope
    NotificationScope_DISPLAYED_ONLY
    NotificationScope_QUEUED_ONLY
    NotificationScope_DISPLAYED_AND_QUEUED
End Enum

' notifications.h: class WZ_Notification (flattened: see the header)
Public Type WZ_Notification
    duration As Long                    ' uint32_t: in game ticks; 0 = "until dismissed by user"
    contentTitle As String              ' a title for the top of the notification - REQUIRED
    contentText As String               ' content text for the notification - REQUIRED
    largeIcon As String                 ' WZ_Notification_Image largeIcon: its imagePath ("" = none, but a memory buffer)
    action_title As String              ' WZ_Notification_Action action: its title ("" = no action button)
    action_onAction As WzFunc           ' and its onAction(notification) (Nothing = none)
    onDismissed As WzFunc               ' (notification, WZ_Notification_Dismissal_Reason reason)
    tag As String                       ' to filter on to cancel / dismiss existing notifications
    displayOptions_identifier As String ' WZ_Notification_Display_Options: _uniqueNotificationIdentifier
    displayOptions_numTimesSeenBeforeDoNotShowAgainOption As Long ' uint8_t = 0
    displayOptions_oneTime As Boolean   ' _isOneTimeNotification = false
    isModal As Boolean                  ' = false: whether the notification should be displayed as a modal overlay
    onDisplay As WzFunc                 ' (notification): called when the notification is initially (fully) displayed
    onIgnored As WzFunc                 ' (notification): called if an ignorable notification is ignored / not displayed
    largeIcon_memoryBuffer() As Byte    ' WZ_Notification_Image's memoryBuffer (a PNG)
    largeIcon_memoryBufferSize As Long
End Type

' the lambdas (WzFunc: notifications_c_fn)
Public Const FN_notifActionButton_onClick As Long = 18600     ' psActionButton->addOnClickHandler([](W_BUTTON& button) {...})
Public Const FN_notifDismissButton_onClick As Long = 18601    ' psDismissButton->addOnClickHandler([](W_BUTTON& button) {...})
Public Const FN_notifButton_initPUserData As Long = 18602     ' sButInit.initPUserDataFunc = []() -> void * {...}
Public Const FN_notifButton_onDelete As Long = 18603          ' sButInit.onDelete = [](WIDGET *psWidget) {...}
Public Const FN_displayNotificationAction As Long = 18604     ' displayNotificationAction (sButInit.pDisplay)
Public Const FN_notifTagEquals As Long = 18605                ' cancelOrDismissNotificationsWithTag's [desiredTag](const std::string& tag)
Public Const FN_notifTagHasPrefix As Long = 18606             ' (VB) cancelOrDismissNotificationIfTagHasPrefix's matchTagFunc
Public Const FN_notifIdentifierHasPrefix As Long = 18607      ' (VB) removeNotificationPreferencesIfHasPrefix's matchIdentifierFunc

' MARK: - Notification Ignore List

' class WZ_Notification_Preferences (its one instance, notificationPrefs)
Private notificationPrefs As Boolean    ' static WZ_Notification_Preferences* notificationPrefs = nullptr: (VB) whether it exists
Private prefs_mRoot As nlohmann_json    ' json mRoot
Private prefs_mFilename As String       ' std::string mFilename

' MARK: - In-Game Notification System Types

' class WZ_Notification_Status: enum NotificationState
Private Const NotificationState_waiting As Long = 0
Private Const NotificationState_opening As Long = 1
Private Const NotificationState_shown As Long = 2
Private Const NotificationState_closing As Long = 3
Private Const NotificationState_closed As Long = 4

' class WZ_Queued_Notification, with its WZ_Notification_Status status and WZ_Notification_Trigger trigger
Private Type WZ_Queued_Notification
    inUse As Boolean                    ' (VB) a live record of the pool
    notification As WZ_Notification
    status_state As Long                ' NotificationState state = waiting
    status_stateStartTime As Long       ' uint32_t = 0
    status_queuedTime As Long           ' uint32_t
    status_animationSpeed As Single     ' float = 1.0f
    trigger_timeInterval As Long        ' uint32_t
    bWasInitiallyShown As Boolean       ' = false
    bWasFullyShown As Boolean           ' = false
    dismissalCause As WZ_Notification_Dismissal_Reason ' = USER_DISMISSED
End Type

Private queuedPool() As WZ_Queued_Notification
Private queuedPoolCount As Long

' MARK: - In-Game Notification Widgets

Private Const WZ_NOTIFICATION_WIDTH As Long = 500
Private Const WZ_NOTIFICATION_PADDING As Long = 15
Private Const WZ_NOTIFICATION_IMAGE_SIZE As Long = 36
Private Const WZ_NOTIFICATION_CONTENTS_LINE_SPACING As Long = 0
Private Const WZ_NOTIFICATION_CONTENTS_TOP_PADDING As Long = 5
Private Const WZ_NOTIFICATION_BUTTON_HEIGHT As Long = 20
Private Const WZ_NOTIFICATION_MODAL_BUTTON_HEIGHT As Long = 25
Private Const WZ_NOTIFICATION_BETWEEN_BUTTON_PADDING As Long = 10

Private Const WZ_NOTIFY_DONOTSHOWAGAINCB_ID As Long = 5

Private Const M_PI_2 As Double = 1.5707963267949        ' pi/2

Private Const WZ_NOTIFICATION_OPEN_DURATION As Long = GAME_TICKS_PER_SEC * 1  ' Time duration for notification open animation, in game ticks
Private Const WZ_NOTIFICATION_CLOSE_DURATION As Long = WZ_NOTIFICATION_OPEN_DURATION \ 2
Private Const WZ_NOTIFICATION_TOP_PADDING As Long = 5
Private Const WZ_NOTIFICATION_TOP_MODAL_PADDING As Long = 20
Private Const WZ_NOTIFICATION_MIN_DELAY_BETWEEN As Long = GAME_TICKS_PER_SEC * 1 ' Minimum delay between display of notifications, in game ticks

Private Const WZ_NOTIFICATION_DOWN_DRAG_DISCARD_CLICK_THRESHOLD As Long = 5

' MARK: - In-game Notification Display

Private notificationQueue() As Long     ' std::list<std::unique_ptr<WZ_Queued_Notification>>: the handles in order
Private notificationQueueCount As Long
Private psNotificationOverlayScreen As W_SCREEN ' std::shared_ptr<W_SCREEN> = nullptr
Private currentNotification As Long     ' std::unique_ptr<WZ_Queued_Notification> (0 = nullptr)
Private currentInGameNotification As WIDGET ' std::shared_ptr<W_NOTIFICATION> = nullptr
Private lastNotificationClosed As Long  ' uint32_t = 0
Private lastDragOnNotificationStartPos As Vector2i ' = (-1,-1)
Private staticInitDone As Boolean       ' (VB) lastDragOnNotificationStartPos's initialiser has run

'--------------------------------------------------------------------------
' (VB) helpers

' the C++ initialiser of lastDragOnNotificationStartPos
Private Sub notificationsStaticInit()
    If staticInitDone Then Exit Sub
    staticInitDone = True
    lastDragOnNotificationStartPos.x = -1
    lastDragOnNotificationStartPos.y = -1
End Sub

' std::make_unique<WZ_Queued_Notification>(notification, WZ_Notification_Status(queuedTime), trigger): its handle
Private Function newQueuedNotification(notification As WZ_Notification, ByVal queuedTime As Long, ByVal timeInterval As Long) As Long
    Dim i As Long
    For i = 0 To queuedPoolCount - 1
        If Not queuedPool(i).inUse Then Exit For
    Next
    If i = queuedPoolCount Then
        If queuedPoolCount = 0 Then
            ReDim queuedPool(0 To 7)
        ElseIf queuedPoolCount > UBound(queuedPool) Then
            ReDim Preserve queuedPool(0 To 2 * queuedPoolCount - 1)
        End If
        queuedPoolCount = queuedPoolCount + 1
    End If
    With queuedPool(i)
        .inUse = True
        .notification = notification
        .status_state = NotificationState_waiting
        .status_stateStartTime = 0
        .status_queuedTime = queuedTime
        .status_animationSpeed = 1!
        .trigger_timeInterval = timeInterval
        .bWasInitiallyShown = False
        .bWasFullyShown = False
        .dismissalCause = WZ_Notification_Dismissal_Reason_USER_DISMISSED
    End With
    newQueuedNotification = i + 1
End Function

' a std::unique_ptr<WZ_Queued_Notification>'s delete: the record freed (its callbacks released)
Private Sub freeQueuedNotification(ByVal h As Long)
    Dim empty_ As WZ_Queued_Notification
    If h = 0 Then Exit Sub
    queuedPool(h - 1) = empty_
End Sub

Private Sub notificationQueue_eraseAt(ByVal it As Long)
    Dim i As Long
    For i = it To notificationQueueCount - 2
        notificationQueue(i) = notificationQueue(i + 1)
    Next
    notificationQueueCount = notificationQueueCount - 1
End Sub

Private Sub notificationQueue_push_back(ByVal h As Long)
    If notificationQueueCount = 0 Then
        ReDim notificationQueue(0 To 7)
    ElseIf notificationQueueCount > UBound(notificationQueue) Then
        ReDim Preserve notificationQueue(0 To 2 * notificationQueueCount - 1)
    End If
    notificationQueue(notificationQueueCount) = h
    notificationQueueCount = notificationQueueCount + 1
End Sub

' WZ_Notification::isIgnorable()
Private Function WZ_Notification_isIgnorable(notification As WZ_Notification) As Boolean
    WZ_Notification_isIgnorable = Len(notification.displayOptions_identifier) <> 0
End Function

' WZ_Notification_Image::empty()
Private Function WZ_Notification_Image_empty(notification As WZ_Notification) As Boolean
    WZ_Notification_Image_empty = Len(notification.largeIcon) = 0 And notification.largeIcon_memoryBufferSize = 0
End Function

'--------------------------------------------------------------------------
' MARK: - Notification Ignore List

' WZ_Notification_Preferences::WZ_Notification_Preferences(const std::string &fileName)
Private Sub WZ_Notification_Preferences_construct(ByVal fileName As String)
    Dim size As Long, data() As Byte, notificationsObject As nlohmann_json, text As String
    prefs_mFilename = fileName
    Set prefs_mRoot = Nothing
    If PHYSFS_exists(fileName) <> 0 Then
        If loadFile(fileName, data, size) Then
            On Error GoTo parse_failed
            Set prefs_mRoot = json_parse_bytes(data, 0, size)
            On Error GoTo 0
parsed:
            If prefs_mRoot Is Nothing Then Set prefs_mRoot = json_parse("null")   ' (an exception left mRoot as it was: null)
            If prefs_mRoot.is_null() Then ASSERT_FAILURE "!mRoot.is_null()", "src/notifications.cpp:111", "WZ_Notification_Preferences", "JSON document from %s is null", fileName
            If Not prefs_mRoot.is_object() Then
                text = utf8_decode(data, 0, size)
                ASSERT_FAILURE "mRoot.is_object()", "src/notifications.cpp:112", "WZ_Notification_Preferences", "JSON document from %s is not an object. Read: \n%s", fileName, text
            End If
            Erase data
        Else
            debug_ 117, LOG_ERROR, "WZ_Notification_Preferences", "Could not open ""%s""", fileName
            ' treat as if no preferences exist yet
            Set prefs_mRoot = json_object()
        End If
    Else
        ' no preferences exist yet
        Set prefs_mRoot = json_object()
    End If

    ' always ensure there's a "notifications" dictionary in the root object
    Set notificationsObject = Nothing
    If prefs_mRoot.is_object() Then Set notificationsObject = prefs_mRoot.find("notifications")
    If notificationsObject Is Nothing Then
        ' create a dictionary object
        prefs_mRoot.set_item "notifications", json_object()
    ElseIf Not notificationsObject.is_object() Then
        ' create a dictionary object
        prefs_mRoot.set_item "notifications", json_object()
    End If
    Exit Sub

parse_failed:
    ' catch (const std::exception &e)
    ASSERT_FAILURE "false", "src/notifications.cpp:106", "WZ_Notification_Preferences", "JSON document from %s is invalid: %s", fileName, Err.Description
    Resume parsed
End Sub

Private Sub WZ_Notification_Preferences_storeValueForNotification(ByVal uniqueNotificationIdentifier As String, ByVal key As String, ByVal value As nlohmann_json)
    Dim notificationsObj As nlohmann_json, notificationData As nlohmann_json
    Set notificationsObj = prefs_mRoot.item("notifications")
    Set notificationData = notificationsObj.find(uniqueNotificationIdentifier)
    If notificationData Is Nothing Then
        notificationsObj.set_item uniqueNotificationIdentifier, json_object()
        Set notificationData = notificationsObj.find(uniqueNotificationIdentifier)
    ElseIf Not notificationData.is_object() Then
        notificationsObj.set_item uniqueNotificationIdentifier, json_object()
        Set notificationData = notificationsObj.find(uniqueNotificationIdentifier)
    End If
    notificationData.set_item key, value
End Sub

' json getJSONValueForNotification(const std::string& uniqueNotificationIdentifier, const std::string& key,
' const json& defaultValue = json()) const
Private Function WZ_Notification_Preferences_getJSONValueForNotification(ByVal uniqueNotificationIdentifier As String, ByVal key As String, ByVal defaultValue As nlohmann_json) As nlohmann_json
    Dim n As Long
    On Error GoTo failed
    Set WZ_Notification_Preferences_getJSONValueForNotification = prefs_mRoot.at("notifications").at(uniqueNotificationIdentifier).at(key)
    Exit Function
failed:
    n = Err.Number - vbObjectError
    If n >= 400 And n < 500 Then
        ' catch (json::out_of_range&): some part of the path doesn't exist yet - return the default value
        Set WZ_Notification_Preferences_getJSONValueForNotification = defaultValue
        Exit Function
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' template<typename T> T getValueForNotification(..., T defaultValue) const: T = int ("seen")
Private Function WZ_Notification_Preferences_getValueForNotification_int(ByVal uniqueNotificationIdentifier As String, ByVal key As String, ByVal defaultValue As Long) As Long
    Dim result As Long
    result = defaultValue
    On Error GoTo failed
    result = WZ_Notification_Preferences_getJSONValueForNotification(uniqueNotificationIdentifier, key, json_value(defaultValue)).get_int()
    WZ_Notification_Preferences_getValueForNotification_int = result
    Exit Function
failed:
    debug_ 79, LOG_WARNING, "getValueForNotification", "Failed to convert json_variant to %s because of error: %s", "int", Err.Description
    WZ_Notification_Preferences_getValueForNotification_int = defaultValue
End Function

' template<typename T> T getValueForNotification(..., T defaultValue) const: T = bool ("skip")
Private Function WZ_Notification_Preferences_getValueForNotification_bool(ByVal uniqueNotificationIdentifier As String, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    Dim result As Boolean
    result = defaultValue
    On Error GoTo failed
    result = WZ_Notification_Preferences_getJSONValueForNotification(uniqueNotificationIdentifier, key, json_value(defaultValue)).get_bool()
    WZ_Notification_Preferences_getValueForNotification_bool = result
    Exit Function
failed:
    debug_ 79, LOG_WARNING, "getValueForNotification", "Failed to convert json_variant to %s because of error: %s", "bool", Err.Description
    WZ_Notification_Preferences_getValueForNotification_bool = defaultValue
End Function

Private Sub WZ_Notification_Preferences_incrementNotificationRuns(ByVal uniqueNotificationIdentifier As String)
    Dim seenCount As Long
    seenCount = WZ_Notification_Preferences_getNotificationRuns(uniqueNotificationIdentifier)
    seenCount = seenCount + 1
    WZ_Notification_Preferences_storeValueForNotification uniqueNotificationIdentifier, "seen", json_value_unsigned(seenCount)
End Sub

' uint32_t getNotificationRuns(const std::string& uniqueNotificationIdentifier) const
Private Function WZ_Notification_Preferences_getNotificationRuns(ByVal uniqueNotificationIdentifier As String) As Long
    WZ_Notification_Preferences_getNotificationRuns = WZ_Notification_Preferences_getValueForNotification_int(uniqueNotificationIdentifier, "seen", 0)
End Function

Private Sub WZ_Notification_Preferences_doNotShowNotificationAgain(ByVal uniqueNotificationIdentifier As String)
    WZ_Notification_Preferences_storeValueForNotification uniqueNotificationIdentifier, "skip", json_value(True)
End Sub

Private Function WZ_Notification_Preferences_getDoNotShowNotificationValue(ByVal uniqueNotificationIdentifier As String) As Boolean
    WZ_Notification_Preferences_getDoNotShowNotificationValue = WZ_Notification_Preferences_getValueForNotification_bool(uniqueNotificationIdentifier, "skip", False)
End Function

' (no caller in 4.7.0)
Private Sub WZ_Notification_Preferences_clearAllNotificationPreferences()
    prefs_mRoot.set_item "notifications", json_object()
End Sub

' bool removeNotificationPreferencesIf(const std::function<bool (const std::string& uniqueNotificationIdentifier)>& matchIdentifierFunc)
Private Function WZ_Notification_Preferences_removeNotificationPreferencesIf(ByVal matchIdentifierFunc As WzFunc) As Boolean
    Dim notificationsObjCopy As nlohmann_json, identifiersToRemove() As String, n As Long, i As Long
    Dim uniqueNotificationIdentifier As String
    If Not prefs_mRoot.contains("notifications") Then ASSERT_FAILURE "mRoot.contains(""notifications"")", "src/notifications.cpp:189", "removeNotificationPreferencesIf", "root missing notifications object": Exit Function
    Set notificationsObjCopy = prefs_mRoot.at("notifications").clone()
    n = 0
    For i = 0 To notificationsObjCopy.size() - 1
        uniqueNotificationIdentifier = notificationsObjCopy.key_at(i)
        If CBool(WzFunc_call(matchIdentifierFunc, uniqueNotificationIdentifier)) Then
            ReDim Preserve identifiersToRemove(0 To n)
            identifiersToRemove(n) = uniqueNotificationIdentifier
            n = n + 1
        End If
    Next
    If n = 0 Then
        WZ_Notification_Preferences_removeNotificationPreferencesIf = False
        Exit Function
    End If
    For i = 0 To n - 1
        notificationsObjCopy.erase_ identifiersToRemove(i)
    Next
    prefs_mRoot.set_item_ref "notifications", notificationsObjCopy
    WZ_Notification_Preferences_removeNotificationPreferencesIf = True
End Function

Private Function WZ_Notification_Preferences_savePreferences() As Boolean
    Dim jsonString As String, bytes() As Byte, n As Long
    On Error GoTo dump_failed
    jsonString = prefs_mRoot.dump(4)
    On Error GoTo 0
    n = utf8_encode(jsonString, bytes)
    saveFile prefs_mFilename, bytes, n
    WZ_Notification_Preferences_savePreferences = True
    Exit Function
dump_failed:
    ASSERT_FAILURE "false", "src/notifications.cpp:219", "savePreferences", "Failed to save JSON to %s with error: %s", prefs_mFilename, Err.Description
    WZ_Notification_Preferences_savePreferences = False
End Function

Private Function WZ_Notification_Preferences_canShowNotification(notification As WZ_Notification) As Boolean
    Dim suppressNotification As Boolean
    If Len(notification.displayOptions_identifier) = 0 Then
        WZ_Notification_Preferences_canShowNotification = True
        Exit Function
    End If

    suppressNotification = False
    If notification.displayOptions_oneTime Then
        suppressNotification = (U32(WZ_Notification_Preferences_getNotificationRuns(notification.displayOptions_identifier)) > 0)
    Else
        suppressNotification = WZ_Notification_Preferences_getDoNotShowNotificationValue(notification.displayOptions_identifier)
    End If

    WZ_Notification_Preferences_canShowNotification = Not suppressNotification
End Function

'--------------------------------------------------------------------------
' MARK: - In-Game Notification System Types

' void WZ_Queued_Notification::setState(WZ_Notification_Status::NotificationState newState)
Private Sub WZ_Queued_Notification_setState(ByVal request As Long, ByVal newState As Long)
    Dim onDisplay As WzFunc
    queuedPool(request - 1).status_state = newState
    queuedPool(request - 1).status_stateStartTime = realTime

    If newState = NotificationState_closed Then
        With queuedPool(request - 1)
            If WZ_Notification_isIgnorable(.notification) And Not .bWasFullyShown And (.dismissalCause <> WZ_Notification_Dismissal_Reason_PROGRAMMATIC) Then
                WZ_Notification_Preferences_incrementNotificationRuns .notification.displayOptions_identifier
            End If
            .bWasFullyShown = True
        End With
    ElseIf newState = NotificationState_shown Then
        If Not queuedPool(request - 1).bWasInitiallyShown Then
            queuedPool(request - 1).bWasInitiallyShown = True
            Set onDisplay = queuedPool(request - 1).notification.onDisplay
            If Not onDisplay Is Nothing Then
                WzFunc_call onDisplay, request
            End If
        End If
    End If
End Sub

' bool wasProgrammaticallyDismissed() const
Private Function WZ_Queued_Notification_wasProgrammaticallyDismissed(ByVal request As Long) As Boolean
    WZ_Queued_Notification_wasProgrammaticallyDismissed = (queuedPool(request - 1).dismissalCause = WZ_Notification_Dismissal_Reason_ACTION_BUTTON_CLICK _
        Or queuedPool(request - 1).dismissalCause = WZ_Notification_Dismissal_Reason_PROGRAMMATIC)
End Function

'--------------------------------------------------------------------------
' MARK: - In-Game Notification Widgets

Private Function MakeNotificationFormInit() As W_FORMINIT
    Dim sFormInit As W_FORMINIT
    W_FORMINIT_construct sFormInit
    sFormInit.formID = 0
    sFormInit.id = 0
    sFormInit.x = 0 ' always base of 0
    sFormInit.y = 0 ' always base of 0
    sFormInit.width = WZ_NOTIFICATION_WIDTH ' fixed width
    sFormInit.height = 100 ' starting height - real height is calculated later based on layout of contents
    MakeNotificationFormInit = sFormInit
End Function

' std::unique_ptr<WZ_Queued_Notification> popNextQueuedNotification(): the handle (0 = nullptr)
Private Function popNextQueuedNotification() As Long
    Dim it As Long, request As Long, num As Long, onIgnored As WzFunc
    If Not notificationPrefs Then ASSERT_FAILURE "notificationPrefs", "src/notifications.cpp:433", "popNextQueuedNotification", "Notification preferences not loaded!"
    it = 0
    Do While it < notificationQueueCount
        request = notificationQueue(it)

        If Not WZ_Notification_Preferences_canShowNotification(queuedPool(request - 1).notification) Then
            ' ignore this notification - remove from the list
            debug_ 442, LOG_GUI, "popNextQueuedNotification", "Ignoring notification: %s", queuedPool(request - 1).notification.displayOptions_identifier
            Set onIgnored = queuedPool(request - 1).notification.onIgnored
            If Not onIgnored Is Nothing Then
                WzFunc_call onIgnored, request
            End If
            notificationQueue_eraseAt it
            freeQueuedNotification request
            GoTo continue_
        End If

        ' std::min<uint32_t>(realTime - queuedTime, timeInterval)
        num = realTime - queuedPool(request - 1).status_queuedTime
        If ULT(queuedPool(request - 1).trigger_timeInterval, num) Then num = queuedPool(request - 1).trigger_timeInterval
        If num = queuedPool(request - 1).trigger_timeInterval Then
            notificationQueue_eraseAt it
            popNextQueuedNotification = request
            Exit Function
        End If

        it = it + 1
continue_:
    Loop
    popNextQueuedNotification = 0
End Function

' MARK: - W_NOTIFICATION

' ////////////////////////////////////////////////////////////////////////////
' display a notification action button
Private Sub displayNotificationAction(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim fx As Long, fy As Long, fw As Long, psBut As WIDGET, hilight As Boolean, greyOut As Boolean
    Dim isActionButton As Boolean, cache As WzText, colour As Long, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Dim fillClr As Long
    Set psBut = psWidget
    hilight = False
    greyOut = (psBut.getState() And WBUT_DISABLE) <> 0 ' if option is unavailable.
    isActionButton = (psBut.UserData = 1)

    ' Any widget using displayTextOption must have its pUserData initialized to a (DisplayTextOptionCache*)
    ' (assert(psWidget->pUserData != nullptr): the C assert, off in a release build)
    Set cache = psWidget.pUserData      ' DisplayNotificationButtonCache& cache: (VB) its wzText

    cache.setText W_BUTTON_(psBut).pText, W_BUTTON_(psBut).FontID

    If W_BUTTON_isHighlighted(psBut) Then   ' if mouse is over text then hilight.
        hilight = True
    End If

    If greyOut Then                         ' unavailable
        colour = psPalette(WZCOL_TEXT_DARK)
    Else                                    ' available
        If hilight Or isActionButton Then   ' hilight
            colour = psPalette(WZCOL_TEXT_BRIGHT)
        Else                                ' don't highlight
            colour = psPalette(WZCOL_TEXT_MEDIUM)
        End If
    End If

    If isActionButton Then
        ' "Action" buttons have a bordering box
        x0 = psBut.x() + xOffset
        y0 = psBut.y() + yOffset
        x1 = x0 + psBut.width()
        y1 = y0 + psBut.height()
        If hilight Then
            fillClr = pal_RGBA(255, 255, 255, 30)
            pie_UniTransBoxFill x0, y0, x1, y1, fillClr
        End If
        iV_Box x0, y0, x1, y1, colour
    End If

    fw = cache.width()
    fy = yOffset + psWidget.y() + (psWidget.height() - cache.lineSize()) \ 2 - cache.aboveBase()

    If (psWidget.style And WBUT_TXTCENTRE) <> 0 Then    'check for centering, calculate offset.
        fx = xOffset + psWidget.x() + ((psWidget.width() - fw) \ 2)
    Else
        fx = xOffset + psWidget.x()
    End If

    If Not greyOut Then
        cache.render fx + 1, fy + 1, pal_RGBA(0, 0, 0, 80)
    End If
    cache.render fx, fy, colour
End Sub

' bool W_NOTIFICATION::dismissNotification(float animationSpeed = 1.0f)
Private Function W_NOTIFICATION_dismissNotification(ByVal self As WIDGET, Optional ByVal animationSpeed As Single = 1!) As Boolean
    Dim request As Long
    request = W_NOTIFICATION_(self).request
    Select Case queuedPool(request - 1).status_state
    Case NotificationState_closing, NotificationState_closed
        ' do nothing
    Case Else
        queuedPool(request - 1).dismissalCause = WZ_Notification_Dismissal_Reason_PROGRAMMATIC   ' setWasProgrammaticallyDismissed()
        queuedPool(request - 1).status_animationSpeed = animationSpeed                          ' status.setAnimationSpeed(animationSpeed)
        WZ_Queued_Notification_setState request, NotificationState_closing
        W_NOTIFICATION_dismissNotification = True
        Exit Function
    End Select

    W_NOTIFICATION_dismissNotification = False
End Function

Private Sub W_NOTIFICATION_deleteThisNotificationWidget(ByVal self As WIDGET)
    Dim psStrongModalParent As WIDGET
    Set psStrongModalParent = WIDGET_fromPtr(W_NOTIFICATION_(self).psModalParent)
    If Not psStrongModalParent Is Nothing Then
        ' if modal and has a modal parent, actually submit the *parent* for removal
        psStrongModalParent.deleteLater
    Else
        self.deleteLater
    End If
End Sub

Private Sub W_NOTIFICATION_internalDismissNotification(ByVal self As WIDGET, Optional ByVal animationSpeed As Single = 1!)
    Dim request As Long
    request = W_NOTIFICATION_(self).request
    ' if notification is the one being displayed, animate it away by setting its state to closing
    Select Case queuedPool(request - 1).status_state
    Case NotificationState_waiting
        ' should not happen
    Case NotificationState_shown
        queuedPool(request - 1).status_animationSpeed = animationSpeed
        WZ_Queued_Notification_setState request, NotificationState_closing
    Case Else
        ' do nothing
    End Select
End Sub

' W_NOTIFICATION(WZ_Queued_Notification* request, W_FORMINIT init = MakeNotificationFormInit())
Public Sub W_NOTIFICATION_construct_init(ByVal self As WIDGET, ByVal request As Long, init As W_FORMINIT)
    Dim d As W_NOTIFICATION
    W_FORM_construct_init self, init
    self.kind = K_W_NOTIFICATION ' (VB) virtual calls from here on reach W_NOTIFICATION's overrides
    Set d = W_NOTIFICATION_(self)
    d.request = request
    If request <> 0 Then
        d.isModal = queuedPool(request - 1).notification.isModal
    End If
End Sub

' static std::shared_ptr<W_NOTIFICATION> make(WZ_Queued_Notification* request, W_FORMINIT init =
' MakeNotificationFormInit()): (VB) the one caller passes the default init
Private Function W_NOTIFICATION_make(ByVal request As Long) As WIDGET
    Dim init As W_FORMINIT, psNewNotificationForm As WIDGET, d As W_NOTIFICATION, isModal As Boolean
    Dim psRootFrm As WIDGET, modalRootFrm As WIDGET, notificationWidth As Long, imageSize As Long
    Dim maxTextWidth As Long, maxTitleWidth As Long, maxContentsWidth As Long, label_title As WIDGET
    Dim titleStartXPos As Long, titleStartYPos As Long, heightOfTitleLabel As Long, titleBottom As Long, titleFont As iV_fonts
    Dim label_contents As WIDGET, contentsStartYPos As Long, heightOfContentsLabel As Long, dismissLabel As String
    Dim actionLabel As String, psActionButton As WIDGET, psDismissButton As WIDGET, buttonsTop As Long
    Dim sButInit As W_BUTINIT, numTimesShown As Long, pDoNotShowAgainButton As WIDGET, minimumDimensions As Vector2i
    Dim bottom_labelContents As Long, calculatedHeight As Long, maxButtonBottom As Long, b1 As Long, b2 As Long
    If Not (request <> 0) Then ASSERT_FAILURE "request != nullptr", "src/notifications.cpp:606", "make", "Request is null"

    init = MakeNotificationFormInit()
    ' (class make_shared_enabler)
    Set psNewNotificationForm = WIDGET_alloc(K_W_NOTIFICATION)
    W_NOTIFICATION_construct_init psNewNotificationForm, request, init
    Set d = W_NOTIFICATION_(psNewNotificationForm)

    isModal = queuedPool(request - 1).notification.isModal
    Set psRootFrm = psNotificationOverlayScreen.psForm
    If isModal Then
        Set modalRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
        psRootFrm.attach modalRootFrm
        Set psRootFrm = modalRootFrm
        d.psModalParent = ObjPtr(modalRootFrm)
    End If

    psRootFrm.attach psNewNotificationForm

    ' Load the image, if specified
    If Not WZ_Notification_Image_empty(queuedPool(request - 1).notification) Then
        Set d.pImageTexture = W_NOTIFICATION_loadImage(request)
    End If

    notificationWidth = psNewNotificationForm.width()

    ' (/* Add the close button */: commented out in 4.7.0)

    ' Calculate dimensions for text area
    If Not d.pImageTexture Is Nothing Then imageSize = WZ_NOTIFICATION_IMAGE_SIZE Else imageSize = 0
    maxTextWidth = notificationWidth - (WZ_NOTIFICATION_PADDING * 2) - imageSize
    If imageSize > 0 Then maxTextWidth = maxTextWidth - WZ_NOTIFICATION_PADDING
    maxTitleWidth = maxTextWidth
    maxContentsWidth = maxTextWidth
    If isModal Then
        maxContentsWidth = notificationWidth - (WZ_NOTIFICATION_PADDING * 2)
    End If

    ' Add title
    Set label_title = W_LABEL_new()
    psNewNotificationForm.attach label_title
    titleStartXPos = WZ_NOTIFICATION_PADDING
    titleStartYPos = WZ_NOTIFICATION_PADDING
    If isModal Then
        titleStartXPos = WZ_NOTIFICATION_PADDING + imageSize
        If imageSize > 0 Then titleStartXPos = titleStartXPos + WZ_NOTIFICATION_PADDING
    End If
    label_title.setGeometry titleStartXPos, titleStartYPos, maxTitleWidth, 12
    W_LABEL_setFontColour label_title, psPalette(WZCOL_TEXT_BRIGHT)
    If isModal Then titleFont = font_medium_bold Else titleFont = font_regular_bold
    heightOfTitleLabel = W_LABEL_setFormattedString(label_title, queuedPool(request - 1).notification.contentTitle, maxTitleWidth, titleFont, WZ_NOTIFICATION_CONTENTS_LINE_SPACING)
    If isModal And (heightOfTitleLabel < imageSize) Then
        ' vertically center title text in available space
        titleStartYPos = titleStartYPos + ((imageSize - heightOfTitleLabel) \ 2)
    End If
    label_title.setGeometry titleStartXPos, titleStartYPos, maxTitleWidth, heightOfTitleLabel
    W_LABEL_setTextAlignment label_title, WLAB_ALIGNTOPLEFT
    label_title.setTransparentToMouse True
    titleBottom = label_title.y() + label_title.height()
    If isModal Then
        If label_title.height() > imageSize Then titleBottom = label_title.y() + label_title.height() Else titleBottom = label_title.y() + imageSize
    End If

    ' Add contents
    Set label_contents = W_LABEL_new()
    psNewNotificationForm.attach label_contents
    contentsStartYPos = titleBottom + WZ_NOTIFICATION_CONTENTS_TOP_PADDING
    If isModal Then
        If imageSize > 0 Then
            contentsStartYPos = titleBottom + WZ_NOTIFICATION_PADDING
        Else
            contentsStartYPos = titleBottom + (WZ_NOTIFICATION_CONTENTS_TOP_PADDING + 3)
        End If
    End If
    label_contents.setGeometry WZ_NOTIFICATION_PADDING, contentsStartYPos, maxContentsWidth, 12
    W_LABEL_setFontColour label_contents, psPalette(WZCOL_TEXT_BRIGHT)
    heightOfContentsLabel = W_LABEL_setFormattedString(label_contents, queuedPool(request - 1).notification.contentText, maxContentsWidth, font_regular, WZ_NOTIFICATION_CONTENTS_LINE_SPACING)
    label_contents.setGeometry label_contents.x(), label_contents.y(), maxContentsWidth, heightOfContentsLabel
    W_LABEL_setTextAlignment label_contents, WLAB_ALIGNTOPLEFT
    label_contents.setTransparentToMouse True

    ' Add action buttons
    dismissLabel = "Dismiss"
    actionLabel = queuedPool(request - 1).notification.action_title

    Set psActionButton = Nothing
    Set psDismissButton = Nothing

    ' Position the buttons below the text contents area
    buttonsTop = label_contents.y() + label_contents.height() + WZ_NOTIFICATION_PADDING
    If isModal Then
        buttonsTop = buttonsTop + (WZ_NOTIFICATION_PADDING \ 2)
    End If

    W_BUTINIT_construct sButInit
    sButInit.formID = 0
    If isModal Then sButInit.height = WZ_NOTIFICATION_MODAL_BUTTON_HEIGHT Else sButInit.height = WZ_NOTIFICATION_BUTTON_HEIGHT
    sButInit.y = buttonsTop
    sButInit.style = sButInit.style Or WBUT_TXTCENTRE
    sButInit.UserData = 0 ' store whether "Action" button or not
    Set sButInit.initPUserDataFunc = WzFunc_new(FN_notifButton_initPUserData)
    Set sButInit.onDelete = WzFunc_new(FN_notifButton_onDelete)
    Set sButInit.pDisplay = WzFunc_new(FN_displayNotificationAction)

    If Len(actionLabel) <> 0 Then
        ' Display both an "Action" button and a "Dismiss" button

        ' 1.) "Action" button
        sButInit.id = 2
        sButInit.width = iV_GetTextWidth(actionLabel, font_regular_bold) + 18
        sButInit.x = CInt(psNewNotificationForm.width() - WZ_NOTIFICATION_PADDING - sButInit.width)    ' (short)
        sButInit.UserData = 1 ' store "Action" state
        sButInit.FontID = font_regular_bold
        sButInit.pText = actionLabel
        Set psActionButton = W_BUTTON_new_init(sButInit)
        psNewNotificationForm.attach psActionButton
        W_BUTTON_addOnClickHandler psActionButton, WzFunc_new(FN_notifActionButton_onClick)
    End If

    If Not psActionButton Is Nothing Or queuedPool(request - 1).notification.duration = 0 Then
        ' 2.) "Dismiss" button
        dismissLabel = ChrW$(&H25B4) & " " & dismissLabel       ' u8"U+25B4 " (a small up triangle)
        sButInit.id = 3
        sButInit.FontID = font_regular
        sButInit.width = iV_GetTextWidth(dismissLabel, font_regular) + 18
        If Not psActionButton Is Nothing Then
            sButInit.x = CInt((psActionButton.x() - WZ_NOTIFICATION_BETWEEN_BUTTON_PADDING) - sButInit.width)
        Else
            sButInit.x = CInt((psNewNotificationForm.width() - WZ_NOTIFICATION_PADDING) - sButInit.width)
        End If
        sButInit.pText = dismissLabel
        sButInit.UserData = 0 ' store regular state
        Set psDismissButton = W_BUTTON_new_init(sButInit)
        W_BUTTON_addOnClickHandler psDismissButton, WzFunc_new(FN_notifDismissButton_onClick)
        psNewNotificationForm.attach psDismissButton
    End If

    If WZ_Notification_isIgnorable(queuedPool(request - 1).notification) And Not queuedPool(request - 1).notification.displayOptions_oneTime Then
        If Not notificationPrefs Then ASSERT_FAILURE "notificationPrefs", "src/notifications.cpp:781", "make", "Notification preferences not loaded!"
        numTimesShown = WZ_Notification_Preferences_getNotificationRuns(queuedPool(request - 1).notification.displayOptions_identifier)
        If Not ULT(numTimesShown, queuedPool(request - 1).notification.displayOptions_numTimesSeenBeforeDoNotShowAgainOption) Then
            ' Display "do not show again" button with checkbox
            Set pDoNotShowAgainButton = WzCheckboxButton_new()
            psNewNotificationForm.attach pDoNotShowAgainButton
            pDoNotShowAgainButton.id = WZ_NOTIFY_DONOTSHOWAGAINCB_ID
            W_BUTTON_(pDoNotShowAgainButton).pText = "Do not show again"
            W_BUTTON_(pDoNotShowAgainButton).FontID = font_small
            minimumDimensions = WzCheckboxButton_calculateDesiredDimensions(pDoNotShowAgainButton)
            If isModal Then b1 = WZ_NOTIFICATION_MODAL_BUTTON_HEIGHT Else b1 = WZ_NOTIFICATION_BUTTON_HEIGHT
            If minimumDimensions.y > b1 Then b1 = minimumDimensions.y
            pDoNotShowAgainButton.setGeometry WZ_NOTIFICATION_PADDING, buttonsTop, minimumDimensions.x, b1
            Set d.pOnDoNotShowAgainCheckbox = pDoNotShowAgainButton
        End If
    End If

    ' calculate the required height for the notification
    bottom_labelContents = label_contents.y() + label_contents.height()
    calculatedHeight = (bottom_labelContents + WZ_NOTIFICATION_PADDING) And &HFFFF&     ' uint16_t
    If Not psActionButton Is Nothing Or Not psDismissButton Is Nothing Then
        b1 = 0: b2 = 0
        If Not psActionButton Is Nothing Then b1 = psActionButton.y() + psActionButton.height()
        If Not psDismissButton Is Nothing Then b2 = psDismissButton.y() + psDismissButton.height()
        If b1 > b2 Then maxButtonBottom = b1 Else maxButtonBottom = b2
        If maxButtonBottom + WZ_NOTIFICATION_PADDING > calculatedHeight Then calculatedHeight = (maxButtonBottom + WZ_NOTIFICATION_PADDING) And &HFFFF&
    End If
    ' Also factor in the image, if one is present
    If imageSize > 0 Then
        If imageSize + (WZ_NOTIFICATION_PADDING * 2) > calculatedHeight Then calculatedHeight = (imageSize + (WZ_NOTIFICATION_PADDING * 2)) And &HFFFF&
    End If
    psNewNotificationForm.setGeometry psNewNotificationForm.x(), psNewNotificationForm.y(), psNewNotificationForm.width(), calculatedHeight

    Set W_NOTIFICATION_make = psNewNotificationForm
End Function

' W_NOTIFICATION::~W_NOTIFICATION()
Public Sub W_NOTIFICATION_destruct(ByVal self As WIDGET)
    Dim d As W_NOTIFICATION
    Set d = W_NOTIFICATION_(self)
    If Not d.pImageTexture Is Nothing Then
        Set d.pImageTexture = Nothing       ' delete pImageTexture
    End If
End Sub

' gfx_api::texture* W_NOTIFICATION::loadImage(const WZ_Notification_Image& image): the request's largeIcon
Private Function W_NOTIFICATION_loadImage(ByVal request As Long) As gfx_api_texture
    Set W_NOTIFICATION_loadImage = WZ_Notification_Image_loadImageToTexture(request)
End Function

' gfx_api::texture* WZ_Notification_Image::loadImageToTexture() const: the request's largeIcon
Private Function WZ_Notification_Image_loadImageToTexture(ByVal request As Long) As gfx_api_texture
    Dim pTexture As gfx_api_texture, ivImage As iV_Image, resultText As String, buffer() As Byte
    Set pTexture = Nothing
    If Len(queuedPool(request - 1).notification.largeIcon) <> 0 Then
        Set pTexture = gfx_api_context_loadTextureFromFile(queuedPool(request - 1).notification.largeIcon, texture_type_user_interface)
    ElseIf queuedPool(request - 1).notification.largeIcon_memoryBufferSize <> 0 Then
        Set ivImage = New iV_Image
        buffer = queuedPool(request - 1).notification.largeIcon_memoryBuffer
        resultText = iV_loadImage_PNG_mem(buffer, ivImage)
        If Len(resultText) <> 0 Then            ' !result.noError()
            debug_ 846, LOG_ERROR, "loadImageToTexture", "Failed to load image from memory buffer: %s", resultText
            Set WZ_Notification_Image_loadImageToTexture = Nothing
            Exit Function
        End If
        Set pTexture = gfx_api_context_loadTextureFromUncompressedImage(ivImage, texture_type_user_interface, "mem::notify_image")
    Else
        ' empty or unhandled case
        Set WZ_Notification_Image_loadImageToTexture = Nothing
        Exit Function
    End If
    Set WZ_Notification_Image_loadImageToTexture = pTexture
End Function

Private Function EaseOutElastic(ByVal p As Single) As Single
    EaseOutElastic = CSng(Sin(-13 * M_PI_2 * (p + 1)) * (2 ^ (-10 * CDbl(p))) + 1)
End Function

Private Function EaseOutQuint(ByVal p As Single) As Single
    Dim f As Single
    f = (p - 1)
    EaseOutQuint = f * f * f * f * f + 1
End Function

Private Function EaseInCubic(ByVal p As Single) As Single
    EaseInCubic = p * p * p
End Function

' bool W_NOTIFICATION::calculateNotificationWidgetPos()
Private Function W_NOTIFICATION_calculateNotificationWidgetPos(ByVal self As WIDGET) As Boolean
    Dim d As W_NOTIFICATION, request As Long, x As Long, y As Long, endingYPosition As Long
    Dim startTime As Long, openAnimationDuration As Single, closeAnimationDuration As Single, endTime As Long
    Dim duration As Long, percentComplete As Single, bDoNotShowAgain As Boolean, state As Long
    Set d = W_NOTIFICATION_(self)
    request = d.request
    If request = 0 Then
        W_NOTIFICATION_calculateNotificationWidgetPos = False
        Exit Function
    End If

    ' center horizontally in window
    x = (screenWidth - self.width()) \ 2
    If x < 0 Then x = 0
    y = 0 ' set below
    If Not d.isModal Then endingYPosition = WZ_NOTIFICATION_TOP_PADDING Else endingYPosition = WZ_NOTIFICATION_TOP_MODAL_PADDING

    ' calculate positioning based on state and time (the C++ switch falls through: each state's block runs on)
    state = queuedPool(request - 1).status_state
    If state = NotificationState_waiting Then
        ' first chance to display
        WZ_Queued_Notification_setState request, NotificationState_opening
        state = NotificationState_opening
        ' fallthrough
    End If
    If state = NotificationState_opening Then
        ' calculate opening progress based on the stateStartTime
        startTime = queuedPool(request - 1).status_stateStartTime
        openAnimationDuration = CSng(WZ_NOTIFICATION_OPEN_DURATION) * queuedPool(request - 1).status_animationSpeed
        endTime = startTime + U32toLongWrap(CDbl(Fix(openAnimationDuration)))
        If ULT(realTime, endTime) Then
            y = CLng(Fix((-self.height()) + (EaseOutQuint((CSng(U32(realTime)) - CSng(U32(startTime))) / openAnimationDuration) * (endingYPosition + self.height())) + 1))
            If Not (y + d.dragOffset_y >= endingYPosition) Then
                GoTo positioned
            Else
                ' factoring in the drag, the notification is already at (or past) fully open
                ' so fall through to immediately transition to the "shown" state
            End If
        End If
        WZ_Queued_Notification_setState request, NotificationState_shown
        state = NotificationState_shown
        ' fallthrough
    End If
    If state = NotificationState_shown Then
        duration = queuedPool(request - 1).notification.duration
        If duration > 0 And Not d.isInDragMode Then
            If d.dragOffset_y > 0 Then
                ' when dragging a notification more *open* (/down),
                ' ensure that the notification remains displayed for at least one additional second
                ' beyond the bounce-back from the drag release
                If ULT(queuedPool(request - 1).status_stateStartTime, realTime - duration + GAME_TICKS_PER_SEC) Then
                    queuedPool(request - 1).status_stateStartTime = realTime - duration + GAME_TICKS_PER_SEC
                End If
            End If
        End If
        If duration = 0 Or ULT(realTime, queuedPool(request - 1).status_stateStartTime + duration) Or d.isInDragMode Then
            y = endingYPosition
            If (y + d.dragOffset_y) > ((-self.height() \ 3) * 2) Then
                GoTo positioned
            End If
        End If
        WZ_Queued_Notification_setState request, NotificationState_closing
        state = NotificationState_closing
        ' fallthrough
    End If
    If state = NotificationState_closing Then
        ' calculate closing progress based on the stateStartTime
        startTime = queuedPool(request - 1).status_stateStartTime
        closeAnimationDuration = CSng(WZ_NOTIFICATION_CLOSE_DURATION) * queuedPool(request - 1).status_animationSpeed
        endTime = startTime + U32toLongWrap(CDbl(Fix(closeAnimationDuration)))
        If ULT(realTime, endTime) Then
            percentComplete = (CSng(U32(realTime)) - CSng(U32(startTime))) / closeAnimationDuration
            If d.dragOffset_y >= 0 Then
                percentComplete = EaseInCubic(percentComplete)
            End If
            y = CLng(Fix(endingYPosition - (percentComplete * (endingYPosition + self.height()))))
            If (y + d.dragOffset_y) > -self.height() Then
                GoTo positioned
            Else
                ' closed early (because of drag offset)
                ' drop through and signal "closed" state
            End If
        End If
        WZ_Queued_Notification_setState request, NotificationState_closed
        state = NotificationState_closed
        ' fallthrough
    End If
    If state = NotificationState_closed Then
        ' widget is now off-screen - get checkbox state (if present)
        bDoNotShowAgain = False
        If Not d.pOnDoNotShowAgainCheckbox Is Nothing Then
            bDoNotShowAgain = WzCheckboxButton_getIsChecked(d.pOnDoNotShowAgainCheckbox)
        End If
        ' then destroy the widget, and finalize the notification request
        removeInGameNotificationForm request
        finishedProcessingNotificationRequest request, bDoNotShowAgain ' after this, request is invalid!
        d.request = 0 ' TEMP
        W_NOTIFICATION_calculateNotificationWidgetPos = True ' processed notification
        Exit Function
    End If

positioned:
    x = x + d.dragOffset_x
    y = y + d.dragOffset_y

    self.move x, y

    W_NOTIFICATION_calculateNotificationWidgetPos = False
End Function

' (VB) a uint32_t(float) conversion's value as a Long (wrapping above 2^31)
Private Function U32toLongWrap(ByVal v As Double) As Long
    If v >= 2147483648# Then v = v - 4294967296#
    U32toLongWrap = CLng(v)
End Function

Public Function W_NOTIFICATION_capturesMouseDrag(ByVal self As WIDGET, ByVal key As WIDGET_KEY) As Boolean
    W_NOTIFICATION_capturesMouseDrag = True
End Function

Public Sub W_NOTIFICATION_mouseDragged(ByVal self As WIDGET, ByVal key As WIDGET_KEY, psStartContext As W_CONTEXT, psContext As W_CONTEXT)
    Dim d As W_NOTIFICATION, dragStartY As Long, currMouseY As Long, distanceY As Long
    Const verticalLimit As Long = 10
    Set d = W_NOTIFICATION_(self)
    dragStartY = psStartContext.my
    currMouseY = psContext.my

    ' calculate how much to respond to the drag by comparing the start to the current position
    If dragStartY > currMouseY Then
        ' dragging up (to close) - respond 1 to 1
        distanceY = dragStartY - currMouseY
        If distanceY > 0 Then d.dragOffset_y = -(distanceY) Else d.dragOffset_y = 0
    ElseIf currMouseY > dragStartY Then
        ' dragging down
        distanceY = currMouseY - dragStartY
        d.dragOffset_y = CLng(Fix(verticalLimit * (1 + CSng(Log(CSng(distanceY) / CSng(verticalLimit)) / Log(10#)))))
    Else
        d.dragOffset_y = 0
    End If
End Sub

' Run a notification widget
Public Sub W_NOTIFICATION_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_NOTIFICATION
    Const dragDecayDuration As Long = GAME_TICKS_PER_SEC * 1
    Set d = W_NOTIFICATION_(self)
    If Not d.isInDragMode Then
        If d.request <> 0 Then
            If queuedPool(d.request - 1).status_state <> NotificationState_closing Then
                ' decay drag offset
                If d.dragOffset_y <> 0 Then
                    d.dragOffset_y = d.dragOffsetEnded_y - CLng(Fix(CSng(d.dragOffsetEnded_y) * EaseOutElastic((CSng(U32(realTime)) - CSng(U32(d.dragEndedTime))) / CSng(dragDecayDuration))))
                End If
            End If
        End If
    End If

    W_NOTIFICATION_calculateNotificationWidgetPos self
End Sub

Public Sub W_NOTIFICATION_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_NOTIFICATION, dragStartMousePos As Vector2i
    Set d = W_NOTIFICATION_(self)
    If Not (d.request <> 0) Then ASSERT_FAILURE "request != nullptr", "src/notifications.cpp:1051", "clicked", "request is null": Exit Sub

    If queuedPool(d.request - 1).status_state = NotificationState_closing Then
        ' if clicked while closing, set state to shown
        queuedPool(d.request - 1).status_state = NotificationState_shown
        queuedPool(d.request - 1).status_stateStartTime = realTime
    End If

    If WzRect_contains(self.geometry(), psContext.mx, psContext.my) Then
        d.isInDragMode = True
        dragStartMousePos.x = psContext.mx
        dragStartMousePos.y = psContext.my
        d.dragStartedTime = realTime
        notificationsDidStartDragOnNotification dragStartMousePos
    End If
End Sub

Public Sub W_NOTIFICATION_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_NOTIFICATION, wasInDragMode As Boolean
    Set d = W_NOTIFICATION_(self)
    wasInDragMode = d.isInDragMode
    If d.isInDragMode Then
        d.isInDragMode = False
        d.dragEndedTime = realTime
        d.dragOffsetEnded_x = d.dragOffset_x
        d.dragOffsetEnded_y = d.dragOffset_y
        notificationDidStopDragOnNotification
    End If

    If d.request <> 0 Then
        If wasInDragMode And d.dragOffset_y < WZ_NOTIFICATION_DOWN_DRAG_DISCARD_CLICK_THRESHOLD Then
            W_NOTIFICATION_internalDismissNotification self
        End If
    End If
End Sub

Public Sub W_NOTIFICATION_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_NOTIFICATION, x0 As Long, y0 As Long, x1 As Long, y1 As Long, videoBuffWidth As Long, videoBuffHeight As Long
    Dim imageLeft As Long, imageTop As Long, image_x0 As Long, image_y0 As Long
    Set d = W_NOTIFICATION_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()
    ' Need to do a little trick here - ensure the bounds are always positive, adjust the height
    If self.y() < 0 Then
        y0 = y0 + -self.y()
    End If

    videoBuffWidth = pie_GetVideoBufferWidth()
    videoBuffHeight = pie_GetVideoBufferHeight()
    If videoBuffWidth = 0 Or videoBuffHeight = 0 Or screenWidth = 0 Or screenHeight = 0 Then
        Exit Sub
    End If

    pie_UniTransBoxFill x0, y0, x1, y1, pal_RGBA(255, 255, 255, 50)
    pie_UniTransBoxFill x0, y0, x1, y1, pal_RGBA(0, 0, 0, 50)

    pie_UniTransBoxFill x0, y0, x1, y1, psPalette(WZCOL_NOTIFICATION_BOX)
    iV_Box2 x0, y0, x1, y1, psPalette(WZCOL_FORM_DARK), psPalette(WZCOL_FORM_DARK)
    iV_Box2 x0 - 1, y0 - 1, x1 + 1, y1 + 1, pal_RGBA(255, 255, 255, 50), pal_RGBA(255, 255, 255, 50)

    ' Display the image, if present
    imageLeft = x1 - WZ_NOTIFICATION_PADDING - WZ_NOTIFICATION_IMAGE_SIZE
    imageTop = (self.y() + yOffset) + WZ_NOTIFICATION_PADDING
    If d.isModal Then
        ' Place the icon on the left side, before the title
        imageLeft = x0 + WZ_NOTIFICATION_PADDING
    End If

    If Not d.pImageTexture Is Nothing Then
        image_x0 = imageLeft
        image_y0 = imageTop

        iV_DrawImageAnisotropic d.pImageTexture, image_x0, image_y0, 0!, 0!, WZ_NOTIFICATION_IMAGE_SIZE, WZ_NOTIFICATION_IMAGE_SIZE, 0!, psPalette(WZCOL_WHITE)
    End If
End Sub

' const std::string& notificationTag() const
Private Function W_NOTIFICATION_notificationTag(ByVal self As WIDGET) As String
    W_NOTIFICATION_notificationTag = queuedPool(W_NOTIFICATION_(self).request - 1).notification.tag
End Function

' bool isActivelyBeingDragged() const
Private Function W_NOTIFICATION_isActivelyBeingDragged(ByVal self As WIDGET) As Boolean
    W_NOTIFICATION_isActivelyBeingDragged = W_NOTIFICATION_(self).isInDragMode
End Function

'--------------------------------------------------------------------------
' MARK: - In-Game Notification System Functions

Private Sub notificationsDidStartDragOnNotification(dragStartPos As Vector2i)
    notificationsStaticInit
    lastDragOnNotificationStartPos = dragStartPos
End Sub

Private Sub notificationDidStopDragOnNotification()
    notificationsStaticInit
    lastDragOnNotificationStartPos = Vector2i_(-1, -1)
End Sub

Public Function notificationsInitialize() As Boolean
    notificationsStaticInit
    WZ_Notification_Preferences_construct "notifications.json"     ' notificationPrefs = new WZ_Notification_Preferences("notifications.json");
    notificationPrefs = True

    ' Initialize the notifications overlay screen
    Set psNotificationOverlayScreen = W_SCREEN_make()
    psNotificationOverlayScreen.psForm.hide ' hiding the root form does not stop display of children, but *does* prevent it from accepting mouse over itself - i.e. basically makes it transparent
    widgRegisterOverlayScreen psNotificationOverlayScreen, &HFFFF&     ' std::numeric_limits<uint16_t>::max()

    notificationsInitialize = True
End Function

Public Sub notificationsShutDown()
    If notificationPrefs Then
        WZ_Notification_Preferences_savePreferences
        Set prefs_mRoot = Nothing               ' delete notificationPrefs;
        prefs_mFilename = ""
        notificationPrefs = False               ' notificationPrefs = nullptr;
    End If

    If Not currentInGameNotification Is Nothing Then
        widgDelete_widget currentInGameNotification
        Set currentInGameNotification = Nothing
    End If

    If Not psNotificationOverlayScreen Is Nothing Then
        widgRemoveOverlayScreen psNotificationOverlayScreen
        Set psNotificationOverlayScreen = Nothing
    End If
End Sub

Public Function isDraggingInGameNotification() As Boolean
    notificationsStaticInit
    ' right now we only support a single concurrent notification
    If Not currentInGameNotification Is Nothing Then
        If W_NOTIFICATION_isActivelyBeingDragged(currentInGameNotification) Then
            isDraggingInGameNotification = True
            Exit Function
        End If
    End If

    ' also handle the case where a drag starts, the notification dismisses itself, but the user has not released the mouse button
    ' in this case, we want to consider it an in-game notification drag *until* the mouse button is released

    isDraggingInGameNotification = (lastDragOnNotificationStartPos.x >= 0 And lastDragOnNotificationStartPos.y >= 0)
End Function

Private Function getOrCreateInGameNotificationForm(ByVal request As Long) As WIDGET
    If request = 0 Then Exit Function

    ' right now we only support a single concurrent notification
    If currentInGameNotification Is Nothing Then
        Set currentInGameNotification = W_NOTIFICATION_make(request)
    End If
    Set getOrCreateInGameNotificationForm = currentInGameNotification
End Function

Private Sub displayNotificationInGame(ByVal request As Long)
    If Not (request <> 0) Then ASSERT_FAILURE "request", "src/notifications.cpp:1221", "displayNotificationInGame", "request is null"
    getOrCreateInGameNotificationForm request
    ' NOTE: Can ignore the result of the above, because it automatically attaches it to the root notification overlay screen
End Sub

Private Sub displayNotification(ByVal request As Long)
    If Not (request <> 0) Then ASSERT_FAILURE "request", "src/notifications.cpp:1228", "displayNotification", "request is null"

    ' By default, display the notification using the in-game notification system
    displayNotificationInGame request
End Sub

' run in-game notifications queue
Public Sub runNotifications()
    ' at the moment, we only support displaying a single notification at a time

    If currentNotification = 0 Then
        If ULT(realTime - lastNotificationClosed, WZ_NOTIFICATION_MIN_DELAY_BETWEEN) Then
            ' wait to fetch a new notification till a future cycle
            Exit Sub
        End If
        ' check for a new notification to display
        currentNotification = popNextQueuedNotification()
        If currentNotification <> 0 Then
            ' display the new notification
            displayNotification currentNotification
        End If
    End If
End Sub

Private Sub removeInGameNotificationForm(ByVal request As Long)
    If request = 0 Then Exit Sub

    ' right now we only support a single concurrent notification
    W_NOTIFICATION_deleteThisNotificationWidget currentInGameNotification
    Set currentInGameNotification = Nothing
End Sub

Private Sub finishedProcessingNotificationRequest(ByVal request As Long, ByVal doNotShowAgain As Boolean)
    Dim onDismissed As WzFunc, reason As WZ_Notification_Dismissal_Reason
    ' at the moment, we only support processing a single notification at a time

    If doNotShowAgain And Not WZ_Queued_Notification_wasProgrammaticallyDismissed(request) Then
        If Not (Len(queuedPool(request - 1).notification.displayOptions_identifier) <> 0) Then ASSERT_FAILURE "!request->notification.displayOptions.uniqueNotificationIdentifier().empty()", "src/notifications.cpp:1271", "finishedProcessingNotificationRequest", "Do Not Show Again was selected, but notification has no ignore key"
        debug_ 1272, LOG_GUI, "finishedProcessingNotificationRequest", "Do Not Show Notification Again: %s", queuedPool(request - 1).notification.displayOptions_identifier
        WZ_Notification_Preferences_doNotShowNotificationAgain queuedPool(request - 1).notification.displayOptions_identifier
    End If

    Set onDismissed = queuedPool(request - 1).notification.onDismissed
    If Not onDismissed Is Nothing Then
        reason = queuedPool(request - 1).dismissalCause     ' request->dismissReason()
        WzFunc_call onDismissed, request, reason
    End If

    ' finished with this notification
    freeQueuedNotification currentNotification
    currentNotification = 0 ' currentNotification.reset(): at this point request is no longer valid!
    lastNotificationClosed = realTime
End Sub

'--------------------------------------------------------------------------
' MARK: - Base functions for handling Notifications

' void addNotification(const WZ_Notification& notification, const WZ_Notification_Trigger& trigger): trigger is
' its timeInterval (0 = WZ_Notification_Trigger::Immediate())
Public Sub addNotification(notification As WZ_Notification, ByVal trigger As Long)
    ' Verify notification properties
    If Len(notification.contentTitle) = 0 Then
        debug_ 1293, LOG_ERROR, "addNotification", "addNotification called with empty notification.contentTitle"
        Exit Sub
    End If
    If Len(notification.contentText) = 0 Then
        debug_ 1298, LOG_ERROR, "addNotification", "addNotification called with empty notification.contentText"
        Exit Sub
    End If

    ' Add the notification to the notification system's queue
    notificationQueue_push_back newQueuedNotification(notification, realTime, trigger)
End Sub

' bool removeNotificationPreferencesIf(const std::function<bool (const std::string& uniqueNotificationIdentifier)>& matchIdentifierFunc)
Public Function removeNotificationPreferencesIf(ByVal matchIdentifierFunc As WzFunc) As Boolean
    If Not notificationPrefs Then ASSERT_FAILURE "notificationPrefs", "src/notifications.cpp:1308", "removeNotificationPreferencesIf", "notificationPrefs is null": removeNotificationPreferencesIf = False: Exit Function
    removeNotificationPreferencesIf = WZ_Notification_Preferences_removeNotificationPreferencesIf(matchIdentifierFunc)
End Function

' (VB) removeNotificationPreferencesIf with the matchIdentifierFunc its callers write: the identifiers that start with
' prefix (uniqueNotificationIdentifier.rfind(prefix, 0) == 0)
Public Function removeNotificationPreferencesIfHasPrefix(ByVal prefix As String) As Boolean
    removeNotificationPreferencesIfHasPrefix = removeNotificationPreferencesIf(WzFunc_new(FN_notifIdentifierHasPrefix, prefix))
End Function

' Whether one or more notifications with the specified tag (exact match) are currently-displayed or queued
' If `scope` is `DISPLAYED_ONLY`, only currently-displayed notifications will be processed
' If `scope` is `QUEUED_ONLY`, only queued notifications will be processed
Public Function hasNotificationsWithTag(ByVal tag As String, Optional ByVal scope As NotificationScope = NotificationScope_DISPLAYED_AND_QUEUED) As Boolean
    Dim i As Long
    If (scope = NotificationScope_DISPLAYED_AND_QUEUED Or scope = NotificationScope_DISPLAYED_ONLY) And Not currentInGameNotification Is Nothing Then
        If W_NOTIFICATION_notificationTag(currentInGameNotification) = tag Then
            hasNotificationsWithTag = True
            Exit Function
        End If
    End If

    If (scope = NotificationScope_DISPLAYED_AND_QUEUED Or scope = NotificationScope_QUEUED_ONLY) Then
        For i = 0 To notificationQueueCount - 1
            If queuedPool(notificationQueue(i) - 1).notification.tag = tag Then
                hasNotificationsWithTag = True
                Exit Function
            End If
        Next
    End If
    hasNotificationsWithTag = False
End Function

' Cancel or dismiss existing notifications by tag (exact match)
' If `scope` is `DISPLAYED_ONLY`, only currently-displayed notifications will be processed
' If `scope` is `QUEUED_ONLY`, only queued notifications will be processed
'
' Returns: `true` if one or more notifications were cancelled or dismissed
Public Function cancelOrDismissNotificationsWithTag(ByVal desiredTag As String, Optional ByVal scope As NotificationScope = NotificationScope_DISPLAYED_AND_QUEUED) As Boolean
    cancelOrDismissNotificationsWithTag = cancelOrDismissNotificationIfTag(WzFunc_new(FN_notifTagEquals, desiredTag), scope)
End Function

' (VB) cancelOrDismissNotificationIfTag with the matchTagFunc its callers write: the tags that start with prefix
' (tag.rfind(prefix, 0) == 0)
Public Function cancelOrDismissNotificationIfTagHasPrefix(ByVal prefix As String, Optional ByVal scope As NotificationScope = NotificationScope_DISPLAYED_AND_QUEUED) As Boolean
    cancelOrDismissNotificationIfTagHasPrefix = cancelOrDismissNotificationIfTag(WzFunc_new(FN_notifTagHasPrefix, prefix), scope)
End Function

' Cancel or dismiss existing notifications by tag
' Accepts a `matchTagFunc` that receives each (queued / currently-displayed) notification's tag,
' and returns "true" if it should be cancelled or dismissed (if queued / currently-displayed)
' If `scope` is `DISPLAYED_ONLY`, only currently-displayed notifications will be processed
' If `scope` is `QUEUED_ONLY`, only queued notifications will be processed
'
' Returns: `true` if one or more notifications were cancelled or dismissed
Public Function cancelOrDismissNotificationIfTag(ByVal matchTagFunc As WzFunc, Optional ByVal scope As NotificationScope = NotificationScope_DISPLAYED_AND_QUEUED) As Boolean
    Dim numCancelledNotifications As Long, it As Long, request As Long
    numCancelledNotifications = 0

    If (scope = NotificationScope_DISPLAYED_AND_QUEUED Or scope = NotificationScope_DISPLAYED_ONLY) And Not currentInGameNotification Is Nothing Then
        If CBool(WzFunc_call(matchTagFunc, W_NOTIFICATION_notificationTag(currentInGameNotification))) Then
            If W_NOTIFICATION_dismissNotification(currentInGameNotification) Then
                debug_ 1367, LOG_GUI, "cancelOrDismissNotificationIfTag", "Dismissing currently-displayed notification with tag: [%s]", W_NOTIFICATION_notificationTag(currentInGameNotification)
                numCancelledNotifications = numCancelledNotifications + 1
            End If
        End If
    End If

    If (scope = NotificationScope_DISPLAYED_AND_QUEUED Or scope = NotificationScope_QUEUED_ONLY) Then
        it = 0
        Do While it < notificationQueueCount
            request = notificationQueue(it)

            If CBool(WzFunc_call(matchTagFunc, queuedPool(request - 1).notification.tag)) Then
                ' cancel this notification - remove from the queue
                debug_ 1383, LOG_GUI, "cancelOrDismissNotificationIfTag", "Cancelling queued notification: [%s]; with tag: [%s]", queuedPool(request - 1).notification.contentTitle, queuedPool(request - 1).notification.tag
                notificationQueue_eraseAt it
                freeQueuedNotification request
                numCancelledNotifications = numCancelledNotifications + 1
            Else
                it = it + 1
            End If
        Loop
    End If

    cancelOrDismissNotificationIfTag = numCancelledNotifications > 0
End Function

'--------------------------------------------------------------------------
' the lambdas

Public Function notifications_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim button As WIDGET, psNewNotificationForm As WIDGET, request As Long, onAction As WzFunc, psWidget As WIDGET
    Dim tag As String
    Select Case f.id
    Case FN_notifActionButton_onClick           ' [](W_BUTTON& button)
        Set button = args(0)
        Set psNewNotificationForm = button.parent()
        If Not WIDGET_isKindOf(psNewNotificationForm, K_W_NOTIFICATION) Then Set psNewNotificationForm = Nothing
        If Not (Not psNewNotificationForm Is Nothing) Then ASSERT_FAILURE "psNewNotificationForm != nullptr", "src/notifications.cpp:746", "make", "null parent": Exit Function
        request = W_NOTIFICATION_(psNewNotificationForm).request
        Set onAction = queuedPool(request - 1).notification.action_onAction
        If Not onAction Is Nothing Then
            WzFunc_call onAction, request
        Else
            debug_ 753, LOG_ERROR, "make", "Action defined (""%s""), but no action handler!", queuedPool(request - 1).notification.action_title
        End If
        queuedPool(request - 1).dismissalCause = WZ_Notification_Dismissal_Reason_ACTION_BUTTON_CLICK  ' setWasDismissedByActionButton()
        W_NOTIFICATION_internalDismissNotification psNewNotificationForm

    Case FN_notifDismissButton_onClick          ' [](W_BUTTON& button)
        Set button = args(0)
        Set psNewNotificationForm = button.parent()
        If Not WIDGET_isKindOf(psNewNotificationForm, K_W_NOTIFICATION) Then Set psNewNotificationForm = Nothing
        If Not (Not psNewNotificationForm Is Nothing) Then ASSERT_FAILURE "psNewNotificationForm != nullptr", "src/notifications.cpp:773", "make", "null parent": Exit Function
        W_NOTIFICATION_internalDismissNotification psNewNotificationForm

    Case FN_notifButton_initPUserData           ' []() -> void * { return new DisplayNotificationButtonCache(); }
        Set notifications_c_fn = New WzText

    Case FN_notifButton_onDelete                ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        ' (assert(psWidget->pUserData != nullptr): the C assert)
        psWidget.pUserData = Empty              ' delete static_cast<DisplayNotificationButtonCache *>(psWidget->pUserData); psWidget->pUserData = nullptr;

    Case FN_displayNotificationAction           ' (widget, xOffset, yOffset)
        displayNotificationAction args(0), args(1), args(2)

    Case FN_notifTagEquals                      ' [desiredTag](const std::string& tag)
        tag = args(0)
        notifications_c_fn = (f.cap(0) = tag)

    Case FN_notifTagHasPrefix                   ' (VB) (const std::string& tag): tag.rfind(prefix, 0) == 0
        tag = args(0)
        notifications_c_fn = (Left$(tag, Len(f.cap(0))) = f.cap(0))

    Case FN_notifIdentifierHasPrefix            ' (VB) (const std::string& uniqueNotificationIdentifier)
        tag = args(0)
        notifications_c_fn = (Left$(tag, Len(f.cap(0))) = f.cap(0))
    End Select
End Function
