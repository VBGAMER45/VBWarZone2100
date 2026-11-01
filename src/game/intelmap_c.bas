Attribute VB_Name = "intelmap_c"
' src/intelmap.cpp, intelmap.h: the Intelligence Map, an overlay screen (intelligenceOverlayScreen)
' whose root form W_INTELLIGENCEOVERLAY_FORM holds the message buttons (IntMessageButton) on a tab
' list, the paused label, the multiplayer menu and the message view (WzMessageView: the title, the
' close button, the research PIE view, the FLIC view and the message texts); their fields are the
' classes of the same names. The research model is mapdisplay_c's renderResearchToBuffer. The
' videos (seqdisp, lib/sequence), the streamed research audio and the VIEW_REPLAY /
' VIEW_RESEARCH data of a VIEWDATA (VIEWDATA_replaySeqCount, VIEWDATA_replaySeq, VIEWDATA_research:
' message.cpp's view data loading is not ported) are stand-ins. intelmap.h's intAddMessageView and
' intRemoveMessageView have no definition in 4.7.0.
' (VB) the weak pointers to the form (std::weak_ptr<W_INTELLIGENCEOVERLAY_FORM>) are locked from a
' widget the lambda holds or is called with: the form is that widget's ancestor while both live
' (intelligenceFormOf); the audio callback has none, it looks at the screen that is up
' (intelligenceFormAt).
' (VB) widget class IntMessageButton : IntFancyButton
' (VB) widget class WzMessageView : IntFormAnimated
' (VB) widget class W_INTELLIGENCEOVERLAY_FORM : W_FORM
Option Explicit

' Intelligence Map screen IDs (intelmap.h)
Public Const IDINTMAP_FORM As Long = 6000           ' The intelligence map base form
Public Const IDINTMAP_MSGVIEW As Long = 6002        ' The message 3D view for the intelligence screen

' Intelligence Map screen IDs
Private Const IDINTMAP_MSGFORM As Long = 6001       ' The intelligence map tabbed form
Private Const IDINTMAP_CLOSE As Long = 6004         ' The close button icon for the 3D view
Private Const IDINTMAP_PAUSELABEL As Long = 6005    ' The paused message
Private Const IDINITMAP_PIEVIEW As Long = 6007      ' The PIE view part of MSGVIEW
Private Const IDINTMAP_FLICVIEW As Long = 6008      ' The Flic View part of MSGVIEW
Private Const IDINTMAP_TEXTVIEW As Long = 6009      ' The Text area of MSGVIEW
Private Const IDINTMAP_TITLELABEL As Long = 6010    ' The title text

Private Const IDINTMAP_MSGSTART As Long = 6100      ' The first button on the intelligence form
Private Const IDINTMAP_MSGEND As Long = 6139        ' The last button on the intelligence form (40 MAX)

Private Const IDINTMAP_SEQTEXTSTART As Long = 6200  ' Sequence subtitle text tabs

' Intelligence Map screen positions: INTMAP_X, INTMAP_Y, INTMAP_WIDTH, INTMAP_HEIGHT, INTMAP_LABELX,
' INTMAP_MSGY (hci.h's layout, some of it depending on the screen size: functions below)
Private Const INTMAP_LABELY As Long = 10
Private Const INTMAP_LABELWIDTH As Long = 60
Private Const INTMAP_LABELHEIGHT As Long = 20

'define the 3D View sizes and positions that are required - relative to INTMAP_FORM
' (INTMAP_RESEARCHX, INTMAP_RESEARCHY: functions below)
Private Const INTMAP_RESEARCHWIDTH As Long = 440
Private Const INTMAP_RESEARCHHEIGHT As Long = 288
Private Const INTMAP_TITLEHEIGHT As Long = 25

'dimensions for FLIC view section relative to IDINTMAP_MSGVIEW
Private Const INTMAP_FLICWIDTH As Long = 192
Private Const INTMAP_FLICHEIGHT As Long = 168

'dimensions for PIE view section relative to IDINTMAP_MSGVIEW
Private Const INTMAP_PIEWIDTH As Long = 238
Private Const INTMAP_PIEHEIGHT As Long = INTMAP_FLICHEIGHT

'position for text on full screen video
Private Const VIDEO_TEXT_TOP_X As Long = 20
Private Const VIDEO_TEXT_TOP_Y As Long = 20
Private Const VIDEO_TEXT_BOTTOM_X As Long = 20
Private Const VIDEO_TEXT_BOTTOM_Y As Long = 444

Private Const TEXT_START_FRAME As Long = 0
Private Const TEXT_END_FRAME As Long = 9999

' #define PAUSE_DISPLAY_CONDITION (!bMultiPlayer)
' #define PAUSEMESSAGE_YOFFSET (0): ingameop.h's too, ingameop_c's PAUSEMESSAGE_YOFFSET

Private Const OVERLAY_MULTIMENU_FORM_Y As Long = 50   ' constexpr int

' WzFunc ids (fork G6: 12000-12099 here)
' (12000: unused; the close button's WIDGET_DISPLAY is intdisplay_c's FN_intDisplayImageHilight)
Public Const FN_intelmap_intDisplayPIEView As Long = 12001       ' intDisplayPIEView as a WIDGET_DISPLAY
Public Const FN_intelmap_intDisplayFLICView As Long = 12002      ' intDisplayFLICView as a WIDGET_DISPLAY
Public Const FN_intelmap_closeButtonCalcLayout As Long = 12003   ' WzMessageView::initialize's LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_intelmap_closeButtonOnClick As Long = 12004      ' WzMessageView::initialize's [weakParent](W_BUTTON &button)
Public Const FN_intelmap_closeViewTask As Long = 12005           ' its widgScheduleTask([weakParent]())
Public Const FN_intelmap_multiMenuFormCalcLayout As Long = 12006 ' createMultiMenuForm's [](WIDGET *form)
Public Const FN_intelmap_closeAnimFinished As Long = 12007       ' closeAnimated's [onCompleteHandler](IntFormAnimated&)
Public Const FN_intelmap_overlayFormCalcLayout As Long = 12008   ' W_INTELLIGENCEOVERLAY_FORM::make's LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_intelmap_msgFormCalcLayout As Long = 12009       ' initialize's LAMBDA_CALCLAYOUT_SIMPLE of msgForm
Public Const FN_intelmap_messageButtonOnClick As Long = 12010    ' initialize's [weakParent](IntMessageButton& widg, WIDGET_KEY mouseButton)
Public Const FN_intelmap_buttonPressedTask As Long = 12011       ' its widgScheduleTask([weakParent, messageButton]())
Public Const FN_intelmap_resetScreenTask As Long = 12012         ' run's widgScheduleTask([](){ intResetScreen(false); })
Public Const FN_intelmap_playCurrentTask As Long = 12013         ' run's widgScheduleTask([weakSelf, scheduledClickButton]())
Public Const FN_intelmap_researchAudioFinished As Long = 12014   ' intIntelButtonPressed's [weakSelf](const AUDIO_STREAM *stream, const void *)
Public Const FN_intelmap_msgDetailsViewCalcLayout As Long = 12015 ' intShowMessageView's LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_intelmap_msgDetailsViewOnClose As Long = 12016   ' intShowMessageView's onCloseFunc [weakSelf]()
Public Const FN_intelmap_removeOverlayScreen As Long = 12017     ' intRemoveIntelMap's [weakOverlayScreen]()

' the widget screen (hci_c's psWScreen)
Private intelligenceOverlayScreen As W_SCREEN    ' static std::shared_ptr<W_SCREEN>

Private immediateMessage As Boolean     ' = false

'flags whether to open the Intel Screen with a message
Private playCurrent As Boolean

'*********************** VARIABLES ****************************
' The current message being displayed
Public psCurrentMsg As Long             ' MESSAGE * = nullptr (0)

' (VB) hci.h's positions behind the INTMAP_ macros
Private Function INTMAP_X() As Long
    INTMAP_X = OBJ_BACKX
End Function

Private Function INTMAP_Y() As Long
    INTMAP_Y = OBJ_BACKY
End Function

Private Function INTMAP_WIDTH() As Long
    INTMAP_WIDTH = OBJ_BACKWIDTH
End Function

Private Function INTMAP_HEIGHT() As Long
    INTMAP_HEIGHT = OBJ_BACKHEIGHT
End Function

Private Function INTMAP_LABELX() As Long
    INTMAP_LABELX = RET_X
End Function

'tabbed message form screen positions
Private Function INTMAP_MSGY() As Long
    INTMAP_MSGY = OBJ_TABY
End Function

Private Function INTMAP_MSGHEIGHT() As Long
    INTMAP_MSGHEIGHT = OBJ_HEIGHT
End Function

Private Function INTMAP_RESEARCHX() As Long
    INTMAP_RESEARCHX = 100 + D_W
End Function

Private Function INTMAP_RESEARCHY() As Long
    INTMAP_RESEARCHY = 30 + D_H
End Function

' (VB) a RESEARCH * as a BASE_STATS * (the research index, -1 for nullptr: its handle, 0 for nullptr)
Private Function researchAsStats(ByVal psResearch As Long) As Long
    If psResearch <> -1 Then researchAsStats = STAT_RESEARCH + psResearch
End Function

' (VB) std::weak_ptr<W_INTELLIGENCEOVERLAY_FORM>::lock() from a widget the caller holds: the form is
' that widget's ancestor while both live (the parent pointers are weak: a dead form's are cleared)
Private Function intelligenceFormOf(ByVal w As WIDGET) As WIDGET
    Dim p As WIDGET
    If w Is Nothing Then Exit Function
    Set p = w.parent()
    Do While Not p Is Nothing
        If WIDGET_isKindOf(p, K_W_INTELLIGENCEOVERLAY_FORM) Then
            Set intelligenceFormOf = p
            Exit Function
        End If
        Set p = p.parent()
    Loop
End Function

' (VB) the same lock from the form's address alone (the audio callback): the root form of the
' intelligence screen that is up
Private Function intelligenceFormAt(ByVal p As Long) As WIDGET
    Dim child As WIDGET
    If intelligenceOverlayScreen Is Nothing Then Exit Function
    If intelligenceOverlayScreen.psForm.childCount() = 0 Then Exit Function
    Set child = intelligenceOverlayScreen.psForm.childAt(0)
    If ObjPtr(child) = p Then Set intelligenceFormAt = child
End Function

'--------------------------------------------------------------------------
' (VB) this module's lambdas and function pointers (WzFunc ids FN_ above): args are the C++ call's
' arguments (a display function: the widget, xOffset, yOffset; calcLayout: the widget; a W_BUTTON
' onClick: the button; a scheduled task: none)
Public Function intelmap_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, psParent As WIDGET, button As WIDGET, strongParent As WIDGET, strongSelf As WIDGET
    Dim messageButton As WIDGET, form_ As WIDGET, view As WIDGET, strongOverlayScreen As W_SCREEN
    Dim width As Long, height As Long, onCompleteHandler As WzFunc
    Select Case f.id
    Case FN_intelmap_intDisplayPIEView
        intDisplayPIEView args(0), args(1), args(2)
    Case FN_intelmap_intDisplayFLICView
        intDisplayFLICView args(0), args(1), args(2)
    Case FN_intelmap_closeButtonCalcLayout
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/intelmap.cpp:253", "initialize", "No parent": Exit Function
        psWidget.setGeometry psParent.width() - OPT_GAP - CLOSE_SIZE, OPT_GAP, CLOSE_SIZE, CLOSE_SIZE
    Case FN_intelmap_closeButtonOnClick
        ' if close button pressed on 3D View then close the view only
        ' (VB: the task holds the button; its parent is weakParent, locked in the task)
        Set button = args(0)
        widgScheduleTask WzFunc_new(FN_intelmap_closeViewTask, button)
    Case FN_intelmap_closeViewTask
        Set button = f.cap(0)
        Set strongParent = button.parent()          ' weakParent.lock()
        If Not strongParent Is Nothing Then
            If Not WIDGET_isKindOf(strongParent, K_WzMessageView) Then Set strongParent = Nothing
        End If
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/intelmap.cpp:261", "initialize", "No parent": Exit Function
        WzMessageView_close strongParent, True
        If Not WzMessageView_(strongParent).onCloseFunc Is Nothing Then
            ' (VB: called with the view, its weak capture's lock)
            WzFunc_call WzMessageView_(strongParent).onCloseFunc, strongParent
        End If
    Case FN_intelmap_multiMenuFormCalcLayout
        Set form_ = args(0)
        Set psParent = form_.parent()
        If psParent Is Nothing Then
            Exit Function
        End If
        width = form_.idealWidth()                  ' std::min((int32_t)psParent->width() - 20, form->idealWidth())
        If psParent.width() - 20 <= width Then width = psParent.width() - 20
        height = form_.idealHeight()
        form_.setGeometry (psParent.width() - width) \ 2, OVERLAY_MULTIMENU_FORM_Y, width, height
    Case FN_intelmap_closeAnimFinished
        ' Trigger onCompleteHandler once close animation is complete
        ' (VB: passing on the form, the lock of intRemoveIntelMap's weak capture)
        Set onCompleteHandler = f.cap(0)
        If Not onCompleteHandler Is Nothing Then
            WzFunc_call onCompleteHandler, args(0)
        End If
    Case FN_intelmap_overlayFormCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight
    Case FN_intelmap_msgFormCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry INTMAP_X, INTMAP_Y, INTMAP_WIDTH, INTMAP_HEIGHT
    Case FN_intelmap_messageButtonOnClick
        ' (IntMessageButton& widg, WIDGET_KEY mouseButton): the task holds the button (messageButton)
        Set messageButton = args(0)
        widgScheduleTask WzFunc_new(FN_intelmap_buttonPressedTask, messageButton)
    Case FN_intelmap_buttonPressedTask
        Set messageButton = f.cap(0)
        Set strongParent = intelligenceFormOf(messageButton)   ' weakParent.lock()
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "src/intelmap.cpp:576", "initialize", "No parent": Exit Function
        W_INTELLIGENCEOVERLAY_FORM_intIntelButtonPressed strongParent, messageButton
    Case FN_intelmap_resetScreenTask
        intResetScreen False
    Case FN_intelmap_playCurrentTask
        Set button = f.cap(0)                       ' scheduledClickButton
        Set strongSelf = intelligenceFormOf(button) ' weakSelf.lock()
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/intelmap.cpp:628", "run", "Null": Exit Function
        W_INTELLIGENCEOVERLAY_FORM_intIntelButtonPressed strongSelf, button
    Case FN_intelmap_researchAudioFinished
        ' (const AUDIO_STREAM *stream, const void *)
        Set strongSelf = intelligenceFormAt(f.cap(0))  ' weakSelf.lock()
        If Not strongSelf Is Nothing Then
            If args(0) Is W_INTELLIGENCEOVERLAY_FORM_(strongSelf).playing Then
                Set W_INTELLIGENCEOVERLAY_FORM_(strongSelf).playing = Nothing
            End If
        End If
    Case FN_intelmap_msgDetailsViewCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry INTMAP_RESEARCHX, INTMAP_RESEARCHY, INTMAP_RESEARCHWIDTH, INTMAP_RESEARCHHEIGHT
    Case FN_intelmap_msgDetailsViewOnClose
        ' (VB: called with the view, whose parent is weakSelf while both live)
        If IsObject(args(0)) Then
            Set view = args(0)
            If Not view Is Nothing Then
                Set psParent = view.parent()
                If Not psParent Is Nothing Then
                    If ObjPtr(psParent) = f.cap(0) Then Set strongSelf = psParent
                End If
            End If
        End If
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/intelmap.cpp:782", "intShowMessageView", "Null": Exit Function
        ' if close button pressed on 3D View then close the view only
        psCurrentMsg = 0
        If bMultiPlayer Then
            If Not W_INTELLIGENCEOVERLAY_FORM_(strongSelf).multiMenuForm Is Nothing Then
                W_INTELLIGENCEOVERLAY_FORM_(strongSelf).multiMenuForm.show
            End If
        End If
    Case FN_intelmap_removeOverlayScreen
        ' on animation complete: remove the overlay screen
        ' (VB: called with the closed form, whose screen pointer, weak, is weakOverlayScreen while it lives)
        If IsObject(args(0)) Then
            Set psWidget = args(0)
            If Not psWidget Is Nothing Then
                Set strongOverlayScreen = psWidget.screenPointer
                If Not strongOverlayScreen Is Nothing Then
                    If ObjPtr(strongOverlayScreen) <> f.cap(0) Then Set strongOverlayScreen = Nothing
                End If
            End If
        End If
        If Not strongOverlayScreen Is Nothing Then
            widgRemoveOverlayScreen strongOverlayScreen
        End If
    End Select
End Function

'--------------------------------------------------------------------------
' MARK: - WzMessageView

' WzMessageView(bool openAnimate = true) : IntFormAnimated(openAnimate) (protected)
Public Sub WzMessageView_construct(ByVal self As WIDGET, Optional ByVal openAnimate As Boolean = True)
    IntFormAnimated_construct self, openAnimate
    self.kind = K_WzMessageView
End Sub

' static std::shared_ptr<WzMessageView> make(MESSAGE *psMessage, bool openAnimate = true)
Public Function WzMessageView_make(ByVal psMessage As Long, Optional ByVal openAnimate As Boolean = True) As WIDGET
    Dim result As WIDGET
    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_WzMessageView)
    WzMessageView_construct result, openAnimate
    result.id = IDINTMAP_MSGVIEW
    WzMessageView_initialize result, psMessage
    Set WzMessageView_make = result
End Function

Public Sub WzMessageView_geometryChanged(ByVal self As WIDGET)
    WzMessageView_(self).closeButton.callCalcLayout
End Sub

Private Function WzMessageView_initialize(ByVal self As WIDGET, ByVal psMessage As Long) As Boolean
    Dim d As WzMessageView, title_ As WIDGET, sButInit As W_BUTINIT, grid As WIDGET, messages As WIDGET
    Dim pad As Padding, pViewData As Long, isReplay As Boolean, nSeq As Long, i As Long, j As Long
    Dim seq As SEQ_DISPLAY, message As WIDGET, sFormInit As W_FORMINIT, form3dView As WIDGET, flicView As WIDGET
    Set d = WzMessageView_(self)
    Set title_ = W_LABEL_new()
    title_.setGeometry 0, 0, INTMAP_RESEARCHWIDTH, INTMAP_TITLEHEIGHT
    W_LABEL_setFontColour title_, psPalette(WZCOL_YELLOW)
    W_LABEL_setTextAlignment title_, WLAB_ALIGNCENTRE
    W_LABEL_setFont title_, font_regular, psPalette(WZCOL_YELLOW)
    title_.setString getMessageTitle(psMessage)
    self.attach title_

    ' Add the close box
    W_BUTINIT_construct sButInit
    sButInit.id = IDINTMAP_CLOSE
    sButInit.x = 0
    sButInit.y = OPT_GAP
    sButInit.width = CLOSE_SIZE
    sButInit.height = CLOSE_SIZE
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    Set d.closeButton = W_BUTTON_new_init(sButInit)
    self.attach d.closeButton
    d.closeButton.setCalcLayout WzFunc_new(FN_intelmap_closeButtonCalcLayout)
    ' (std::weak_ptr<WzMessageView> weakParent: the button's parent, see FN_intelmap_closeViewTask)
    W_BUTTON_addOnClickHandler d.closeButton, WzFunc_new(FN_intelmap_closeButtonOnClick)

    Set grid = GridLayout_new()
    self.attach grid
    grid.setGeometry 1, INTMAP_TITLEHEIGHT, INTMAP_RESEARCHWIDTH - 2, INTMAP_RESEARCHHEIGHT - 1 - INTMAP_TITLEHEIGHT

    Set messages = ScrollableListWidget_make()
    ScrollableListWidget_setItemSpacing messages, 3
    ScrollableListWidget_setBackgroundColor messages, psPalette(WZCOL_TRANSPARENT_BOX)
    pad.top = 10: pad.right_ = 10: pad.bottom = 10: pad.left_ = 10
    ScrollableListWidget_setPadding messages, pad
    messages.setGeometry 0, 0, grid.width(), 0 ' set initial width so that grid calculations can properly calculate required height

    pViewData = aMessage(psMessage).pViewData
    isReplay = False                ' psMessage->type != MSG_RESEARCH && psMessage->pViewData->type == VIEW_RPL
    If aMessage(psMessage).typ <> MSG_RESEARCH Then
        If aViewData(pViewData).typ = VIEW_RPL Then isReplay = True
    End If
    If isReplay Then
        ' (VB) VIEW_REPLAY *psViewReplay: its seqList is VIEWDATA_replaySeqCount / VIEWDATA_replaySeq
        nSeq = VIEWDATA_replaySeqCount(pViewData)
        For i = 0 To nSeq - 1
            seq = VIEWDATA_replaySeq(pViewData, i)
            For j = 0 To seq.textMsg.size - 1
                Set message = Paragraph_new()
                Paragraph_setFontColour message, psPalette(WZCOL_TEXT_BRIGHT)
                Paragraph_addText message, seq.textMsg.items(j)
                ScrollableListWidget_addItem messages, message
            Next
        Next

        GridLayout_place grid, grid_allocation_slot_(0, 1, False), grid_allocation_slot_(0, 1, True), messages

        WzMessageView_initialize = True
        Exit Function
    End If

    If Not (aMessage(psMessage).typ <> MSG_PROXIMITY) Then ASSERT_FAILURE "psMessage->type != MSG_PROXIMITY", "src/intelmap.cpp:299", "initialize", "Invalid message type for research": WzMessageView_initialize = False: Exit Function

    'Add the PIE box
    W_FORMINIT_construct sFormInit
    sFormInit.id = IDINITMAP_PIEVIEW
    sFormInit.style = WFORM_PLAIN
    sFormInit.width = INTMAP_PIEWIDTH
    sFormInit.height = INTMAP_PIEHEIGHT
    Set sFormInit.pDisplay = WzFunc_new(FN_intelmap_intDisplayPIEView)
    sFormInit.pUserData = psMessage
    Set form3dView = W_CLICKFORM_new_init(sFormInit)

    'Add the Flic box if videos are installed, or on-demand video streaming is available
    If seq_hasVideos() Then
        W_FORMINIT_construct sFormInit  ' sFormInit = W_FORMINIT();
        sFormInit.id = IDINTMAP_FLICVIEW
        sFormInit.style = WFORM_PLAIN
        sFormInit.width = INTMAP_FLICWIDTH
        sFormInit.height = INTMAP_FLICHEIGHT
        Set sFormInit.pDisplay = WzFunc_new(FN_intelmap_intDisplayFLICView)
        sFormInit.pUserData = psMessage
        Set flicView = W_FORM_new_init(sFormInit)

        GridLayout_place grid, grid_allocation_slot_(0, 2, True), grid_allocation_slot_(0, 1, False), form3dView
        GridLayout_place grid, grid_allocation_slot_(2, 1, False), grid_allocation_slot_(0, 1, False), flicView
    Else
        GridLayout_place grid, grid_allocation_slot_(1, 1, False), grid_allocation_slot_(0, 1, False), form3dView
    End If

    'Add the text box
    For i = 0 To aViewData(pViewData).textMsg.size - 1
        Set message = Paragraph_new()
        Paragraph_setFontColour message, psPalette(WZCOL_TEXT_BRIGHT)
        Paragraph_addText message, aViewData(pViewData).textMsg.items(i)
        ScrollableListWidget_addItem messages, message
    Next

    GridLayout_place grid, grid_allocation_slot_(0, 3, True), grid_allocation_slot_(1, 1, True), messages

    WzMessageView_initialize = True
End Function

' void close(bool animated)
Public Sub WzMessageView_close(ByVal self As WIDGET, ByVal animated As Boolean)
    Dim dummy As String, strongThis As WIDGET, psParent As WIDGET
    'stop the video
    dummy = ""
    seq_RenderVideoToBuffer dummy, SEQUENCE_KILL

    If animated Then
        ' Start the window close animation.
        IntFormAnimated_closeAnimateDelete self
    Else
        'remove without the animating close window
        Set strongThis = self
        strongThis.deleteLater
        Set psParent = self.parent()
        If Not psParent Is Nothing Then
            psParent.detach strongThis
        End If
    End If
End Sub

'--------------------------------------------------------------------------
' MARK: - Intelligence screen

' W_INTELLIGENCEOVERLAY_FORM(W_FORMINIT const *init) : W_FORM(init) {} (protected)
Public Sub W_INTELLIGENCEOVERLAY_FORM_construct_init(ByVal self As WIDGET, init As W_FORMINIT)
    W_FORM_construct_init self, init
    self.kind = K_W_INTELLIGENCEOVERLAY_FORM
    W_INTELLIGENCEOVERLAY_FORM_initMembers self
End Sub

' W_INTELLIGENCEOVERLAY_FORM() : W_FORM() {} (protected)
Public Sub W_INTELLIGENCEOVERLAY_FORM_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_W_INTELLIGENCEOVERLAY_FORM
    W_INTELLIGENCEOVERLAY_FORM_initMembers self
End Sub

' the members' initialisers
Private Sub W_INTELLIGENCEOVERLAY_FORM_initMembers(ByVal self As WIDGET)
    Dim d As W_INTELLIGENCEOVERLAY_FORM
    Set d = W_INTELLIGENCEOVERLAY_FORM_(self)
    d.isClosing = False
    d.delayedPlayCurrent = False
    Set d.playing = Nothing             ' AUDIO_STREAM *playing = nullptr
End Sub

' ~W_INTELLIGENCEOVERLAY_FORM()
Public Sub W_INTELLIGENCEOVERLAY_FORM_destruct(ByVal self As WIDGET)
    Dim d As W_INTELLIGENCEOVERLAY_FORM
    Set d = W_INTELLIGENCEOVERLAY_FORM_(self)
    If sound_isStreamPlaying(d.playing) Then
        sound_StopStream d.playing
    End If
End Sub

Public Sub W_INTELLIGENCEOVERLAY_FORM_geometryChanged(ByVal self As WIDGET)
    Dim d As W_INTELLIGENCEOVERLAY_FORM
    Set d = W_INTELLIGENCEOVERLAY_FORM_(self)
    If Not d.multiMenuForm Is Nothing Then
        d.multiMenuForm.callCalcLayout
    End If
    If Not d.msgForm Is Nothing Then
        d.msgForm.callCalcLayout
    End If
    If Not d.msgDetailsView Is Nothing Then
        d.msgDetailsView.callCalcLayout
    End If
End Sub

Private Function W_INTELLIGENCEOVERLAY_FORM_createMultiMenuForm(ByVal self As WIDGET) As WIDGET
    Dim form_ As WIDGET
    Set form_ = intCreateMultiMenuForm()
    form_.id = MULTIMENU_FORM

    form_.setCalcLayout WzFunc_new(FN_intelmap_multiMenuFormCalcLayout)

    Set W_INTELLIGENCEOVERLAY_FORM_createMultiMenuForm = form_
End Function

' void closeAnimated(std::function<void ()> onCompleteHandler)
Public Sub W_INTELLIGENCEOVERLAY_FORM_closeAnimated(ByVal self As WIDGET, ByVal onCompleteHandler As WzFunc)
    Dim d As W_INTELLIGENCEOVERLAY_FORM
    Set d = W_INTELLIGENCEOVERLAY_FORM_(self)
    'remove 3dView if still there
    If Not d.msgDetailsView Is Nothing Then
        WzMessageView_close d.msgDetailsView, False
        Set d.msgDetailsView = Nothing
    End If

    ' Start the window close animation.
    If Not d.msgForm Is Nothing Then
        IntFormAnimated_closeAnimateDelete d.msgForm, WzFunc_new(FN_intelmap_closeAnimFinished, onCompleteHandler)
    End If

    'remove the text label
    widgDelete self.screenPointer, IDINTMAP_PAUSELABEL

    If bMultiPlayer Then
        If Not d.multiMenuForm Is Nothing Then
            IntFormAnimated_closeAnimateDelete d.multiMenuForm
        End If
    End If

    d.isClosing = True
End Sub

' static std::shared_ptr<W_INTELLIGENCEOVERLAY_FORM> make(bool _playCurrent, UDWORD formID = 0)
Public Function W_INTELLIGENCEOVERLAY_FORM_make(ByVal playCurrent_ As Boolean, Optional ByVal formID As Long = 0) As WIDGET
    Dim sInit As W_FORMINIT, result As WIDGET
    W_FORMINIT_construct sInit
    sInit.id = formID
    sInit.style = WFORM_PLAIN
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1
    Set sInit.calcLayout = WzFunc_new(FN_intelmap_overlayFormCalcLayout)

    ' (class make_shared_enabler)
    Set result = WIDGET_alloc(K_W_INTELLIGENCEOVERLAY_FORM)
    W_INTELLIGENCEOVERLAY_FORM_construct_init result, sInit
    W_INTELLIGENCEOVERLAY_FORM_initialize result, playCurrent_, True
    Set W_INTELLIGENCEOVERLAY_FORM_make = result
End Function

Private Sub W_INTELLIGENCEOVERLAY_FORM_initialize(ByVal self As WIDGET, ByVal playCurrent_ As Boolean, ByVal animate As Boolean)
    Dim d As W_INTELLIGENCEOVERLAY_FORM, sLabInit As W_LABINIT, pausedLabel As WIDGET, msgList As WIDGET
    Dim msgListWidth As Long, it As Long, psMessage As Long, button As WIDGET
    Set d = W_INTELLIGENCEOVERLAY_FORM_(self)
    'add message to indicate game is paused - single player mode
    If Not bMultiPlayer Then            ' PAUSE_DISPLAY_CONDITION
        W_LABINIT_construct sLabInit
        sLabInit.id = IDINTMAP_PAUSELABEL
        sLabInit.formID = 0
        sLabInit.x = INTMAP_LABELX
        sLabInit.y = INTMAP_LABELY + PAUSEMESSAGE_YOFFSET
        sLabInit.width = INTMAP_LABELWIDTH
        sLabInit.height = INTMAP_LABELHEIGHT
        sLabInit.pText = "PAUSED"
        Set pausedLabel = W_LABEL_new_init(sLabInit)
        self.attach pausedLabel
    End If

    ' (std::weak_ptr<W_INTELLIGENCEOVERLAY_FORM> weakParent: the click handlers' tasks find the form
    ' from their button, see FN_intelmap_buttonPressedTask)

    ' Add the main Intelligence Map form
    If selectedPlayer < MAX_PLAYERS Then
        Set d.msgForm = IntFormAnimated_new(animate)  ' Do not animate the opening, if the window was already open.
        self.attach d.msgForm
        d.msgForm.id = IDINTMAP_FORM
        d.msgForm.setCalcLayout WzFunc_new(FN_intelmap_msgFormCalcLayout)

        ' Add the Message form
        Set msgList = IntListTabWidget_make()
        d.msgForm.attach msgList
        msgList.id = IDINTMAP_MSGFORM
        ListTabWidget_setChildSize msgList, OBJ_BUTWIDTH, OBJ_BUTHEIGHT
        ListTabWidget_setChildSpacing msgList, OBJ_GAP, OBJ_GAP
        msgListWidth = OBJ_BUTWIDTH * 5 + OBJ_GAP * 4
        msgList.setGeometry (d.msgForm.width() - msgListWidth) \ 2, INTMAP_MSGY, msgListWidth, d.msgForm.height() - INTMAP_MSGY

        ' Add the message buttons

        'add each button
        it = apsMessages(selectedPlayer).head
        Do While it <> 0
            psMessage = lnObj(it)
            'if (psMessage->type == MSG_TUTORIAL)
            '{
            '    //tutorial cases should never happen
            '    ASSERT( false, "Tutorial message in Intelligence screen!" );
            '    continue;
            '}
            If aMessage(psMessage).typ = MSG_PROXIMITY Then
                'ignore proximity messages here (continue)
            Else
                Set button = IntMessageButton_new()
                msgList.attach button
                IntMessageButton_setMessage button, psMessage
                ListTabWidget_addWidgetToLayout msgList, button

                ' if the current message matches psSelected lock the button
                If psMessage = psCurrentMsg Then
                    Set d.selectedMsgButton = button
                    button.setState WBUT_LOCK
                    ListTabWidget_setCurrentPage msgList, ListTabWidget_pages(msgList) - 1
                End If

                ' Add the click handler
                IntMessageButton_setOnClickHandler button, WzFunc_new(FN_intelmap_messageButtonOnClick)
            End If
            it = lnNext(it)
        Loop
        'check to play current message instantly
        If playCurrent_ And (Not d.selectedMsgButton Is Nothing) Then
            If aMessage(psCurrentMsg).typ <> MSG_PROXIMITY Then
                d.delayedPlayCurrent = True
            End If
        End If
    End If

    If bMultiPlayer And (d.multiMenuForm Is Nothing) Then
        Set d.multiMenuForm = W_INTELLIGENCEOVERLAY_FORM_createMultiMenuForm(self)
        self.attach d.multiMenuForm
        d.multiMenuForm.callCalcLayout
    End If
End Sub

Public Sub W_INTELLIGENCEOVERLAY_FORM_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    ' no-op
End Sub

Public Sub W_INTELLIGENCEOVERLAY_FORM_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' no-op
End Sub

Public Sub W_INTELLIGENCEOVERLAY_FORM_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_INTELLIGENCEOVERLAY_FORM, scheduledClickButton As WIDGET
    Set d = W_INTELLIGENCEOVERLAY_FORM_(self)
    If d.isClosing Then Exit Sub
    If keyPressed(KEY_ESC) Then
        inputLoseFocus  ' clear the input buffer.
        widgScheduleTask WzFunc_new(FN_intelmap_resetScreenTask)
    End If

    'check to play current message instantly
    If d.delayedPlayCurrent And (Not d.selectedMsgButton Is Nothing) Then
        d.delayedPlayCurrent = False
        ' (std::weak_ptr weakSelf: the task finds the form from the button, see FN_intelmap_playCurrentTask)
        Set scheduledClickButton = d.selectedMsgButton
        widgScheduleTask WzFunc_new(FN_intelmap_playCurrentTask, scheduledClickButton)
    End If
End Sub

'
'deal with the actual button press - proxMsg is set to true if a proximity
'button has been pressed
'
Private Sub W_INTELLIGENCEOVERLAY_FORM_intIntelButtonPressed(ByVal self As WIDGET, ByVal button As WIDGET)
    Dim d As W_INTELLIGENCEOVERLAY_FORM, psResearch As Long, psMessage As Long, pViewData As Long, audio As String
    Set d = W_INTELLIGENCEOVERLAY_FORM_(self)

    ' message button has been pressed - clear the old button and messageView
    If Not d.selectedMsgButton Is Nothing Then
        If Not d.selectedMsgButton Is button Then
            d.selectedMsgButton.setState 0
            psCurrentMsg = 0
            Set d.selectedMsgButton = Nothing
        End If
    End If

    If button Is Nothing Then
        Exit Sub
    End If

    Set d.selectedMsgButton = button

    ' Lock the new button
    ' This means we can't click on the same movie button twice.
    d.selectedMsgButton.setState WBUT_CLICKLOCK

    ' Get the message for the new button
    psMessage = IntMessageButton_getMessage(d.selectedMsgButton)

    'deal with the message if one
    If psMessage <> 0 Then
        If Not (aMessage(psMessage).typ <> MSG_PROXIMITY) Then ASSERT_FAILURE "psMessage->type != MSG_PROXIMITY", "src/intelmap.cpp:667", "intIntelButtonPressed", "Shouldn't be able to get a proximity message!": Exit Sub

        'set the current message
        psCurrentMsg = psMessage

        'set the read flag
        aMessage(psCurrentMsg).read = True

        debug_ 675, LOG_GUI, "intIntelButtonPressed", "intIntelButtonPressed: Dealing with a new message type=%d", aMessage(psMessage).typ

        'should never have a proximity message
        If aMessage(psMessage).typ = MSG_PROXIMITY Then
            Exit Sub
        End If

        pViewData = aMessage(psMessage).pViewData
        If pViewData <> 0 Then
            ' If it's a video sequence then play it anyway
            If aViewData(pViewData).typ = VIEW_RPL Then
                If aMessage(psMessage).pViewData <> 0 Then
                    W_INTELLIGENCEOVERLAY_FORM_intShowMessageView self, psMessage
                End If
                ' attempt to show video(s)
                StartMessageSequences psMessage, True
            ElseIf aViewData(pViewData).typ = VIEW_RES Then
                psResearch = getResearchForMsg(pViewData)
                If psResearch <> -1 Then
                    ' only play the sample once, otherwise, they tend to overlap each other
                    If sound_isStreamPlaying(d.playing) Then
                        sound_StopStream d.playing
                    End If

                    audio = vbNullString    ' char const *audio = nullptr
                    Select Case asResearch(psResearch).iconID
                    Case IMAGE_RES_DROIDTECH, IMAGE_RES_CYBORGTECH
                        audio = "sequenceaudio/res_droid.ogg"
                    Case IMAGE_RES_WEAPONTECH
                        audio = "sequenceaudio/res_weapons.ogg"
                    Case IMAGE_RES_COMPUTERTECH
                        audio = "sequenceaudio/res_com.ogg"
                    Case IMAGE_RES_POWERTECH
                        audio = "sequenceaudio/res_pow.ogg"
                    Case IMAGE_RES_SYSTEMTECH
                        audio = "sequenceaudio/res_systech.ogg"
                    Case IMAGE_RES_STRUCTURETECH, IMAGE_RES_DEFENCE
                        audio = "sequenceaudio/res_strutech.ogg"
                    End Select

                    If StrPtr(audio) <> 0 Then
                        ' (std::weak_ptr weakSelf: the form's address, see FN_intelmap_researchAudioFinished)
                        Set d.playing = audio_PlayStream(audio, sound_GetUIVolume(), audio_StreamFinished_new(WzFunc_new(FN_intelmap_researchAudioFinished, ObjPtr(self))), Empty)
                    End If
                End If

                If aMessage(psMessage).pViewData <> 0 Then
                    W_INTELLIGENCEOVERLAY_FORM_intShowMessageView self, psMessage
                End If
            End If
        End If
    End If
End Sub

'Add the 3D world view for the particular message
Private Sub W_INTELLIGENCEOVERLAY_FORM_intShowMessageView(ByVal self As WIDGET, ByVal psMessage As Long)
    Dim d As W_INTELLIGENCEOVERLAY_FORM, animate As Boolean
    Set d = W_INTELLIGENCEOVERLAY_FORM_(self)
    animate = True

    ' Is the form already up?
    If Not d.msgDetailsView Is Nothing Then
        WzMessageView_close d.msgDetailsView, False
        animate = False
    End If
    If Not d.multiMenuForm Is Nothing Then
        d.multiMenuForm.hide
    End If

    Set d.msgDetailsView = WzMessageView_make(psMessage, animate)
    self.attach d.msgDetailsView
    d.msgDetailsView.setCalcLayout WzFunc_new(FN_intelmap_msgDetailsViewCalcLayout)

    ' (std::weak_ptr weakSelf: the form's address, locked from the view, see FN_intelmap_msgDetailsViewOnClose)
    Set WzMessageView_(d.msgDetailsView).onCloseFunc = WzFunc_new(FN_intelmap_msgDetailsViewOnClose, ObjPtr(self))
End Sub

Private Function createIntelligenceScreen() As W_SCREEN
    Dim result As W_SCREEN, newRootFrm As WIDGET
    ' Initialize the intelligence overlay screen
    Set result = W_SCREEN_make()
    Set newRootFrm = W_INTELLIGENCEOVERLAY_FORM_make(playCurrent)
    newRootFrm.setTransparentToMouse True
    result.psForm.hide ' hide actual root form so it doesn't get clicks
    result.psForm.attach newRootFrm

    Set createIntelligenceScreen = result
End Function


'--------------------------------------------------------------------------
' MARK: -

Public Function intAddIntelMap() As Boolean
    If Not intelligenceOverlayScreen Is Nothing Then
        ' overlay screen already up
        intAddIntelMap = True
        Exit Function
    End If

    'check playCurrent with psCurrentMsg
    If psCurrentMsg = 0 Then
        playCurrent = False
    Else
        playCurrent = True
    End If

    'set pause states before putting the interface up
    audio_StopAll
    setIntelligencePauseState

    Set intelligenceOverlayScreen = createIntelligenceScreen()
    widgRegisterOverlayScreenOnTopOfScreen intelligenceOverlayScreen, psWScreen

    intAddIntelMap = True
End Function

' Add all the Video Sequences for a message
Private Function StartMessageSequences(ByVal psMessage As Long, ByVal Start As Boolean) As Boolean
    Dim bLoop As Boolean, pViewData As Long, Sequence As Long, nSeq As Long, seq As SEQ_DISPLAY
    Dim psViewReplay As VIEW_RESEARCH
    bLoop = False

    debug_ 840, LOG_GUI, "StartMessageSequences", "StartMessageSequences: start message sequence"

    'should never have a proximity message here
    If aMessage(psMessage).typ = MSG_PROXIMITY Then
        StartMessageSequences = False
        Exit Function
    End If

    pViewData = aMessage(psMessage).pViewData
    If Not (pViewData <> 0) Then ASSERT_FAILURE "psMessage->pViewData != nullptr", "src/intelmap.cpp:848", "StartMessageSequences", "Invalid ViewData pointer": StartMessageSequences = False: Exit Function

    If aViewData(pViewData).typ = VIEW_RPL Then
        ' (VB) VIEW_REPLAY *psViewReplay: its seqList is VIEWDATA_replaySeqCount / VIEWDATA_replaySeq

        seq_ClearSeqList

        'add any sequences to the list to be played when the first one is finished
        nSeq = VIEWDATA_replaySeqCount(pViewData)
        For Sequence = 0 To nSeq - 1
            seq = VIEWDATA_replaySeq(pViewData, Sequence)
            If seq.flag = 1 Then
                bLoop = True
            Else
                bLoop = False
            End If

            seq_AddSeqToList seq.sequenceName, seq.audio, vbNullString, bLoop

            debug_ 873, LOG_GUI, "StartMessageSequences", "StartMessageSequences: sequence=%d", Sequence
            addVideoText seq, Sequence
        Next
        'play first full screen video
        If Start = True Then
            If Not seq_StartNextFullScreenVideo() Then
                StartMessageSequences = False
                Exit Function
            End If
        End If
    ElseIf aViewData(pViewData).typ = VIEW_RES Then
        'UDWORD Sequence;

        psViewReplay = VIEWDATA_research(pViewData)   ' (VIEW_RESEARCH *)psMessage->pViewData->pData

        seq_ClearSeqList
        seq_AddSeqToList psViewReplay.sequenceName, psViewReplay.audio, vbNullString, False
        'play first full screen video
        If Start = True Then
            If Not seq_StartNextFullScreenVideo() Then
                StartMessageSequences = False
                Exit Function
            End If
        End If
    Else
        StartMessageSequences = False
        Exit Function
    End If

    StartMessageSequences = True
End Function

Private Sub intCleanUpIntelMap()
    Dim removedAMessage As Boolean, it As Long, itNext As Long, psMessage As Long
    removedAMessage = False

    If selectedPlayer < MAX_PLAYERS Then
        'remove any research messages that have been read
        ' (mutating_list_iterate: the next node is taken before the function, which may erase this one)
        it = apsMessages(selectedPlayer).head
        Do While it <> 0
            itNext = lnNext(it)
            psMessage = lnObj(it)
            If aMessage(psMessage).typ = MSG_RESEARCH And aMessage(psMessage).read Then
                removeMessage psMessage, selectedPlayer
                removedAMessage = True
            End If
            ' return IterationResult::CONTINUE_ITERATION;
            it = itNext
        Loop
    End If
    If removedAMessage Then
        jsDebugMessageUpdate
    End If
    resetIntelligencePauseState
    immediateMessage = False
End Sub


' Remove the Intelligence Map widgets from the screen
Public Sub intRemoveIntelMap()
    Dim rootIntelligenceForm As WIDGET, child As WIDGET
    If intelligenceOverlayScreen Is Nothing Then
        Exit Sub
    End If

    ' Start the window close animation.
    ' (dynamic_pointer_cast of intelligenceOverlayScreen->psForm->children().front())
    If intelligenceOverlayScreen.psForm.childCount() > 0 Then
        Set child = intelligenceOverlayScreen.psForm.childAt(0)
        If WIDGET_isKindOf(child, K_W_INTELLIGENCEOVERLAY_FORM) Then Set rootIntelligenceForm = child
    End If
    If rootIntelligenceForm Is Nothing Then
        If Not (Not rootIntelligenceForm Is Nothing) Then ASSERT_FAILURE "rootIntelligenceForm != nullptr", "src/intelmap.cpp:949", "intRemoveIntelMap", "Failed to get intelligence form?"
        intRemoveIntelMapNoAnim
        Exit Sub
    End If

    ' (std::weak_ptr<W_SCREEN> weakOverlayScreen: the screen's address, see FN_intelmap_removeOverlayScreen)
    W_INTELLIGENCEOVERLAY_FORM_closeAnimated rootIntelligenceForm, WzFunc_new(FN_intelmap_removeOverlayScreen, ObjPtr(intelligenceOverlayScreen))

    Set intelligenceOverlayScreen = Nothing

    intCleanUpIntelMap
End Sub

' Remove the Intelligence Map widgets from the screen
Public Sub intRemoveIntelMapNoAnim()
    If intelligenceOverlayScreen Is Nothing Then
        Exit Sub
    End If

    widgRemoveOverlayScreen intelligenceOverlayScreen
    Set intelligenceOverlayScreen = Nothing

    intCleanUpIntelMap
End Sub

'--------------------------------------------------------------------------
' MARK: - IntMessageButton: Displays the buttons used on the intelligence map

' IntMessageButton() : IntFancyButton(), psMsg(nullptr)
Public Sub IntMessageButton_construct(ByVal self As WIDGET)
    IntFancyButton_construct self
    self.kind = K_IntMessageButton
    IntMessageButton_(self).psMsg = 0
End Sub

Public Function IntMessageButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_IntMessageButton)
    IntMessageButton_construct widget
    Set IntMessageButton_new = widget
End Function

Public Sub IntMessageButton_setMessage(ByVal self As WIDGET, ByVal msg As Long)
    IntMessageButton_(self).psMsg = msg
End Sub

Public Function IntMessageButton_getTip(ByVal self As WIDGET) As String
    Dim pMessageTitle As String
    If Not (IntMessageButton_(self).psMsg <> 0) Then ASSERT_FAILURE "psMsg != nullptr", "src/intelmap.cpp:156", "getTip", "Null message?": IntMessageButton_getTip = "": Exit Function
    pMessageTitle = getMessageTitle(IntMessageButton_(self).psMsg)
    If StrPtr(pMessageTitle) <> 0 Then IntMessageButton_getTip = pMessageTitle Else IntMessageButton_getTip = ""
End Function

Public Function IntMessageButton_getMessage(ByVal self As WIDGET) As Long
    IntMessageButton_getMessage = IntMessageButton_(self).psMsg
End Function

' void setOnClickHandler(const W_BUTTON_ONCLICK_FUNC& onClickFunc): onClickFunc is called with
' (IntMessageButton& button, WIDGET_KEY mouseButton)
Public Sub IntMessageButton_setOnClickHandler(ByVal self As WIDGET, ByVal onClickFunc As WzFunc)
    Set IntMessageButton_(self).onClickHandler = onClickFunc
End Sub

'Displays the buttons used on the intelligence map
Public Sub IntMessageButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As IntMessageButton, pResearch As Long, MovieButton As Boolean, object_ As ImdObject, image As AtlasImage
    Dim psStat As Long, compID As Long, isDownInc As Long
    Set d = IntMessageButton_(self)
    pResearch = -1                      ' RESEARCH *pResearch = nullptr
    MovieButton = False

    IntFancyButton_initDisplay self

    'shouldn't have any proximity messages here...
    If aMessage(d.psMsg).typ = MSG_PROXIMITY Then
        Exit Sub
    End If
    'set the graphic for the button
    Select Case aMessage(d.psMsg).typ
    Case MSG_RESEARCH
        pResearch = getResearchForMsg(aMessage(d.psMsg).pViewData)
        'IMDType = IMDTYPE_RESEARCH;
        'set the IMDType depending on what stat is associated with the research
        psStat = 0
        If pResearch <> -1 Then psStat = asResearch(pResearch).psStat
        If psStat <> 0 Then             ' pResearch && pResearch->psStat
            'we have a Stat associated with this research topic
            If StatIsStructure(psStat) Then
                'this defines how the button is drawn
                object_ = ImdObject_StructureStat(psStat)
            Else
                compID = StatIsComponent(psStat)
                If compID <> COMP_NUMCOMPONENTS Then
                    'this defines how the button is drawn
                    object_ = ImdObject_Component(psStat)
                Else
                    ASSERT_FAILURE "false", "src/intelmap.cpp:1027", "display", "intDisplayMessageButton: invalid stat"
                    object_ = ImdObject_Research(researchAsStats(pResearch))
                End If
            End If
        Else
            'no Stat for this research topic so use the research topic to define what is drawn
            object_ = ImdObject_Research(researchAsStats(pResearch))
        End If
    Case MSG_CAMPAIGN
        image = AtlasImage_(IntImages, IMAGE_INTEL_CAMPAIGN)
        MovieButton = True
    Case MSG_MISSION
        image = AtlasImage_(IntImages, IMAGE_INTEL_MISSION)
        MovieButton = True
    Case Else
        debug_ 1047, LOG_ERROR, "display", "Unknown message type: %i", aMessage(d.psMsg).typ
        Exit Sub
    End Select

    'if research message
    If pResearch <> -1 Then
        If asResearch(pResearch).iconID <> NO_RESEARCH_ICON Then
            image = AtlasImage_(IntImages, asResearch(pResearch).iconID)
        End If

        'do we have the same icon for the top right hand corner?
        IntFancyButton_displayIMD self, image, object_, xOffset, yOffset
    Else
        'draw buttons for mission and general messages
        ' Draw the button with the relevant image, don't add isDown() to the image ID if it's a movie button.
        If MovieButton Then
            IntFancyButton_displayImage self, image, xOffset, yOffset
        Else
            If W_CLICKFORM_isDown(self) Then isDownInc = 1 Else isDownInc = 0
            IntFancyButton_displayImage self, AtlasImage_(image.images, image.id + isDownInc), xOffset, yOffset
        End If
    End If
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

Public Sub IntMessageButton_released(ByVal self As WIDGET, context As W_CONTEXT, Optional ByVal mouseButton As WIDGET_KEY = WKEY_PRIMARY)
    Dim clickAndReleaseOnThisButton As Boolean
    clickAndReleaseOnThisButton = ((W_CLICKFORM_(self).state And WBUT_DOWN) <> 0) ' relies on W_CLICKFORM handling to properly set WBUT_DOWN

    W_CLICKFORM_released self, context, mouseButton   ' IntFancyButton::released (W_CLICKFORM's: IntFancyButton has none)

    If Not clickAndReleaseOnThisButton Then
        Exit Sub ' do nothing
    End If

    If Not IntMessageButton_(self).onClickHandler Is Nothing Then
        WzFunc_call IntMessageButton_(self).onClickHandler, self, mouseButton
    End If
End Sub

Public Function IntMessageButton_clickHeld(ByVal self As WIDGET, psContext As W_CONTEXT, ByVal key As WIDGET_KEY) As Boolean
    If key = WKEY_PRIMARY Then
        If Not IntMessageButton_(self).onClickHandler Is Nothing Then
            WzFunc_call IntMessageButton_(self).onClickHandler, self, WKEY_SECONDARY
        End If
        IntMessageButton_clickHeld = True
        Exit Function
    End If
    IntMessageButton_clickHeld = False
End Function

' displays the PIE view for the current message
Private Sub intDisplayPIEView(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim psMessage As Long, image As Long, x0 As Long, y0 As Long, width As Long, height As Long, psResearch As Long
    psMessage = psWidget.pUserData      ' (MESSAGE *)psWidget->pUserData (Empty: nullptr)
    image = -1

    ' Should not have any proximity messages here...
    If psMessage = 0 Then
        Exit Sub
    ElseIf aMessage(psMessage).typ = MSG_PROXIMITY Then
        Exit Sub
    End If

    If aMessage(psMessage).pViewData <> 0 Then
        x0 = xOffset + psWidget.x()
        y0 = yOffset + psWidget.y()
        width = psWidget.width()
        height = psWidget.height()

        'moved from after close render
        RenderWindowFrame FRAME_NORMAL, x0, y0, width, height

        If aViewData(aMessage(psMessage).pViewData).typ <> VIEW_RES Then
            ASSERT_FAILURE "false", "src/intelmap.cpp:1125", "intDisplayPIEView", "intDisplayPIEView: Invalid message type"
            Exit Sub
        End If

        'render an object
        psResearch = getResearchForMsg(aMessage(psMessage).pViewData)
        If psResearch <> -1 Then
            renderResearchToBuffer psResearch, x0 + width \ 2, y0 + height \ 2
        End If

        'draw image icon in top left of window
        If psResearch <> -1 Then
            image = asResearch(psResearch).iconID      ' (SWORD)
            If image > &H7FFF& Then image = image - &H10000
            If image > 0 Then
                iV_DrawImage IntImages, image, x0 + 3, y0 + 3
            End If
        End If
    End If
End Sub

' displays the FLIC view for the current message
Private Sub intDisplayFLICView(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim psMessage As Long, psViewResearch As VIEW_RESEARCH, x0 As Long, y0 As Long, x1 As Long, y1 As Long
    psMessage = psWidget.pUserData      ' (MESSAGE *)psWidget->pUserData (Empty: nullptr)

    'shouldn't have any proximity messages here...
    If psMessage = 0 Then
        Exit Sub
    ElseIf aMessage(psMessage).typ = MSG_PROXIMITY Then
        Exit Sub
    End If

    If aMessage(psMessage).pViewData <> 0 Then
        x0 = xOffset + psWidget.x()
        y0 = yOffset + psWidget.y()
        x1 = x0 + psWidget.width()
        y1 = y0 + psWidget.height()

        If aViewData(aMessage(psMessage).pViewData).typ <> VIEW_RES Then
            ASSERT_FAILURE "false", "src/intelmap.cpp:1169", "intDisplayFLICView", "intDisplayFLICView: Invalid message type"
            Exit Sub
        End If

        RenderWindowFrame FRAME_NORMAL, x0, y0, x1 - x0, y1 - y0
        psViewResearch = VIEWDATA_research(aMessage(psMessage).pViewData)
        ' set the dimensions to window size & position
        seq_SetDisplaySize INTMAP_FLICWIDTH, INTMAP_FLICHEIGHT, x0, y0
        'render a frame of the current movie *must* force above resolution!
        seq_RenderVideoToBuffer psViewResearch.sequenceName, SEQUENCE_HOLD
    End If
End Sub


'adds text to full screen video
Private Sub addVideoText(psSeqDisplay As SEQ_DISPLAY, ByVal sequence As Long)
    Dim i As Long, x As Long, y As Long

    If psSeqDisplay.textMsg.size <> 0 Then
        'add each message, first at the top
        ' FIXME We should perhaps get font size, and use that to calculate offset(s) ?
        x = VIDEO_TEXT_TOP_X
        y = VIDEO_TEXT_TOP_Y

        seq_AddTextForVideo psSeqDisplay.textMsg.items(0), x, y, TEXT_START_FRAME, TEXT_END_FRAME, SEQ_TEXT_POSITION 'startframe endFrame

        'add each message, the rest at the bottom
        x = VIDEO_TEXT_BOTTOM_X
        ' calculate the real bottom... NOTE, game assumes all videos to be 640x480
        y = ToInt(CDbl(pie_GetVideoBufferHeight()) / 480# * CDbl(VIDEO_TEXT_BOTTOM_Y))
        i = 1
        Do While i < psSeqDisplay.textMsg.size
            seq_AddTextForVideo psSeqDisplay.textMsg.items(i), x, y, TEXT_START_FRAME, TEXT_END_FRAME, SEQ_TEXT_POSITION 'startframe endFrame
            'initialise after the first setting
            x = 0: y = 0
            i = i + 1
        Loop
    End If
End Sub

'sets psCurrentMsg for the Intelligence screen
Public Sub setCurrentMsg()
    Dim it As Long
    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/intelmap.cpp:1215", "setCurrentMsg", "Unsupported selectedPlayer: %u", selectedPlayer: Exit Sub

    ' (rbegin() to rend(): the list backwards)
    it = apsMessages(selectedPlayer).tail
    Do While it <> 0
        If aMessage(lnObj(it)).typ <> MSG_PROXIMITY Then
            psCurrentMsg = lnObj(it)
            Exit Sub
        End If
        it = lnPrev(it)
    Loop
End Sub

'sets which states need to be paused when the intelligence screen is up
Public Sub setIntelligencePauseState()
    If Not bMultiPlayer Then
        'need to clear mission widgets from being shown on intel screen
        clearMissionWidgets
        gameTimeStop
        setGameUpdatePause True
        If Not bInTutorial Then
            ' Don't pause the scripts or the console if the tutorial is running.
            setScriptPause True
            setConsolePause True
        End If
        setScrollPause True
        screen_RestartBackDrop
    End If
End Sub

'resets the pause states
Public Sub resetIntelligencePauseState()
    If Not bMultiPlayer Then
        'put any widgets back on for the missions
        resetMissionWidgets
        setGameUpdatePause False
        If Not bInTutorial Then
            setScriptPause False
        End If
        setScrollPause False
        setConsolePause False
        gameTimeStart
        screen_StopBackDrop
    End If
End Sub

' Play an intelligence message.
' This function is used from scripts to give updates to the mission.
Public Sub displayImmediateMessage(ByVal psMessage As Long)
    '
    '    This has to be changed to support a script calling a message in the intelligence screen
    '

    ' attempt to show video(s)
    psCurrentMsg = psMessage
    If Not StartMessageSequences(psMessage, True) Then
        psCurrentMsg = 0
    End If
    ' remind the player that the message can be seen again from
    ' the intelligence screen
    addConsoleMessage "New Intelligence Report", CENTRE_JUSTIFY, SYSTEM_MESSAGE
    flashReticuleButton IDRET_INTEL_MAP
End Sub


' return whether a message is immediate
Public Function messageIsImmediate() As Boolean
    messageIsImmediate = immediateMessage
End Function

'sets the flag
Public Sub setMessageImmediate(ByVal state As Boolean)
    immediateMessage = state
End Sub

' run intel map (in the game loop)
Public Sub intRunIntelMap()
    If keyPressed(KEY_ESC) Then
        intResetScreen False
        ' clear key press so we don't enter in-game options
        inputLoseFocus
    End If
End Sub

' (VB) const char * (nullptr: vbNullString) of a MESSAGE handle
Private Function getMessageTitle(ByVal message As Long) As String
    Dim research_ As Long
    Select Case aMessage(message).typ
    Case MSG_RESEARCH
        research_ = getResearchForMsg(aMessage(message).pViewData)
        If research_ <> -1 Then getMessageTitle = asResearch(research_).name_ Else getMessageTitle = "Research Update"
    Case MSG_CAMPAIGN
        getMessageTitle = "Project Goals and Updates"
    Case MSG_MISSION
        getMessageTitle = "Current Objective"
    Case Else
        getMessageTitle = vbNullString
    End Select
End Function
