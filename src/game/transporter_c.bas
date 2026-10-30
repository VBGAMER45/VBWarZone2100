Attribute VB_Name = "transporter_c"
' src/transporter.cpp, transporter.h: code to deal with loading/unloading, interface and flight of
' transporters (whole). The interface's forms are widgets (docs/CONVENTIONS.md, "Widgets"; the lambdas'
' ids are 13500 - 13999); transInterfaceDroidList's DroidList * is the list's first node (0 for
' nullptr: the callers only walk it).
Option Explicit

Public Const IDTRANS_FORM As Long = 9000        ' The Transporter base form
Public Const IDTRANS_CONTENTFORM As Long = 9003 ' The Transporter Contents form
Public Const IDTRANS_DROIDS As Long = 9006      ' The Droid base form
Public Const IDTRANS_LAUNCH As Long = 9010      ' The Transporter Launch button
Public Const IDTRANS_CAPACITY As Long = 9500    ' The Transporter capacity label

'defines how much space is on the Transporter
Public Const TRANSPORTER_CAPACITY As Long = 10

'#define IDTRANS_FORM			9000	//The Transporter base form
Private Const IDTRANS_CLOSE As Long = 9002          ' The close button icon
Private Const IDTRANS_CONTCLOSE As Long = 9005      ' The close icon on the Contents form
Private Const IDTRANS_DROIDTAB As Long = 9007       ' The Droid tab form
Private Const IDTRANS_DROIDCLOSE As Long = 9008     ' The close icon for the Droid form

Private Const IDTRANS_START As Long = 9100          ' The first button on the Transporter tab form
Private Const IDTRANS_END As Long = 9199            ' The last button on the Transporter tab form
Private Const IDTRANS_STATSTART As Long = 9200      ' The status button for the first Transporter
Private Const IDTRANS_STATEND As Long = 9299        ' The status button for the last Transporter
Private Const IDTRANS_CONTSTART As Long = 9300      ' The first button on the Transporter contents tab form
Private Const IDTRANS_CONTEND As Long = 9399        ' The last button on the Transporter contents tab form
Private Const IDTRANS_DROIDSTART As Long = 9400     ' The first button on the Droid tab form
Private Const IDTRANS_DROIDEND As Long = 9499       ' The last button on the Droid tab form
Private Const IDTRANS_REPAIRBARSTART As Long = 9600 ' The first repair status bar on Droid button
Private Const IDTRANS_REPAIRBAREND As Long = 9699   ' The last repair status bar on Droid button

' Transporter screen positions (TRANS_X, TRANS_Y: hci_c's OBJ_BACKX, OBJ_BACKY, below)
Private Const TRANS_WIDTH As Long = OBJ_BACKWIDTH
Private Const TRANS_HEIGHT As Long = OBJ_BACKHEIGHT

'tabbed form screen positions
Private Const TRANS_TABY As Long = OBJ_TABY

'Transported contents screen positions (TRANSCONT_Y: STAT_Y, below)
Private Const TRANSCONT_X As Long = STAT_X
Private Const TRANSCONT_WIDTH As Long = STAT_WIDTH
Private Const TRANSCONT_HEIGHT As Long = STAT_HEIGHT

'contents tabbed form screen positions
Private Const TRANSCONT_TABY As Long = STAT_TABFORMY

'droid form screen positions (TRANSDROID_X: RADTLX, TRANSDROID_Y: STAT_Y, below)
Private Const TRANSDROID_WIDTH As Long = STAT_WIDTH
Private Const TRANSDROID_HEIGHT As Long = STAT_HEIGHT

'start y position of the available droids buttons
Private Const AVAIL_STARTY As Long = 0

Private Const MAX_TRANSPORT_FULL_MESSAGE_PAUSE As Long = 20000

' Maximum distance to the nearest transport (16 tiles)
Private Const MAX_NEAREST_TRANSPORT_SQ_DIST As Long = TILE_WIDTH * TILE_WIDTH * 256

Private Const TRANSPORTOUT_TIME As Long = 4& * GAME_TICKS_PER_SEC

' Static variables
Private psCurrTransporter As Long       ' DROID * = nullptr
Private onMission As Boolean
Private g_iLaunchTime As Long           ' UDWORD = 0
'used for audio message for reinforcements
Private bFirstTransporter As Boolean
'the tab positions of the DroidsAvail window
Private objMajor As Long                ' size_t = 0
' Last time the transporter is full message was displayed
Private lastTransportIsFullMsgTime As Long  ' UDWORD = 0

' (VB) the lambdas
Public Const FN_intAddTransporter_transForm_calcLayout As Long = 13500
Public Const FN_intAddTransporterContents_transContentForm_calcLayout As Long = 13501
Public Const FN_intAddTransporterContents_closeButton_calcLayout As Long = 13502
Public Const FN_intUpdateTransCapacity As Long = 13503
Public Const FN_intAddDroidsAvailForm_transDroids_calcLayout As Long = 13504
Public Const FN_intAddDroidsAvailForm_closeButton_calcLayout As Long = 13505
Public Const FN_intAddDroidsAvailForm_droidList_calcLayout As Long = 13506

' (VB) the screen positions that are functions of the screen's size
Private Function TRANS_X() As Long
    TRANS_X = OBJ_BACKX()
End Function

Private Function TRANS_Y() As Long
    TRANS_Y = OBJ_BACKY()
End Function

Private Function TRANSCONT_Y() As Long
    TRANSCONT_Y = STAT_Y()
End Function

Private Function TRANSDROID_X() As Long
    TRANSDROID_X = RADTLX()
End Function

Private Function TRANSDROID_Y() As Long
    TRANSDROID_Y = STAT_Y()
End Function

' (VB) PERCENT(a,b) (((a)*100)/(b)) on UDWORDs
Private Function PERCENT_u(ByVal a As Long, ByVal b As Long) As Long
    Dim p As Double
    p = U32(a) * 100#
    p = p - Int(p / 4294967296#) * 4294967296#
    PERCENT_u = FromU32(Int(p / U32(b)))
End Function

'initialises Transporter variables
Public Sub initTransporters()
    onMission = False
    psCurrTransporter = 0
End Sub

' Call to refresh the transporter screen, ie when a droids boards it.
'
Public Function intRefreshTransporter() As Boolean
    Dim Ret As Boolean
    ' Is the transporter screen up?
    If intMode = INT_TRANSPORTER And Not widgGetFromID(psWScreen, IDTRANS_FORM) Is Nothing Then
        ' Refresh it by re-adding it.
        Ret = intAddTransporter(psCurrTransporter, onMission)
        intMode = INT_TRANSPORTER
        intRefreshTransporter = Ret
        Exit Function
    End If

    intRefreshTransporter = True
End Function

Public Function intAddTransporter(ByVal psSelected As Long, ByVal offWorld As Boolean) As Boolean
    Dim Animate As Boolean, parent As WIDGET, transForm As WIDGET, sButInit As W_BUTINIT
    Animate = True

    onMission = offWorld
    psCurrTransporter = psSelected

    'if transporter has died - close the interface - this can only happen in
    'multiPlayer where the transporter can be killed
    ' (&&: isDead only for a transporter)
    If bMultiPlayer And psCurrTransporter <> 0 Then
        If isDead(psCurrTransporter) Then
            intRemoveTransNoAnim
            intAddTransporter = True
            Exit Function
        End If
    End If

    ' Add the main Transporter form
    ' Is the form already up?
    If Not widgGetFromID(psWScreen, IDTRANS_FORM) Is Nothing Then
        intRemoveTransNoAnim
        Animate = False
    End If

    If intIsRefreshing() Then
        Animate = False
    End If

    Set parent = psWScreen.psForm

    Set transForm = IntFormAnimated_new(Animate)    ' Do not animate the opening, if the window was already open.
    parent.attach transForm
    transForm.id = IDTRANS_FORM
    transForm.setCalcLayout WzFunc_new(FN_intAddTransporter_transForm_calcLayout)

    ' Add the close button
    W_BUTINIT_construct sButInit
    sButInit.formID = IDTRANS_FORM
    sButInit.id = IDTRANS_CLOSE
    sButInit.x = TRANS_WIDTH - CLOSE_WIDTH
    sButInit.y = 0
    sButInit.width = CLOSE_WIDTH
    sButInit.height = CLOSE_HEIGHT
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddTransporter = False
        Exit Function
    End If

    If Not intAddTransButtonForm() Then
        intAddTransporter = False
        Exit Function
    End If

    ' Add the Transporter Contents form (and buttons)
    If Not intAddTransporterContents() Then
        intAddTransporter = False
        Exit Function
    End If

    'if on a mission - add the Droids back at home base form
    ' (&&: the form only on a mission)
    If onMission Then
        If Not intAddDroidsAvailForm() Then
            intAddTransporter = False
            Exit Function
        End If
    End If

    intAddTransporter = True
End Function

' Add the main Transporter Contents Interface
Private Function intAddTransporterContents() As Boolean
    Dim Animate As Boolean, parent As WIDGET, transContentForm As WIDGET, sButInit As W_BUTINIT
    Dim sLabInit As W_LABINIT, sButFInit As W_FORMINIT, psForm As WIDGET, help As WidgetHelp
    Animate = True

    ' Is the form already up?
    If Not widgGetFromID(psWScreen, IDTRANS_CONTENTFORM) Is Nothing Then
        intRemoveTransContentNoAnim
        Animate = False
    End If

    If intIsRefreshing() Then
        Animate = False
    End If

    Set parent = psWScreen.psForm

    Set transContentForm = IntFormAnimated_new(Animate) ' Do not animate the opening, if the window was already open.
    parent.attach transContentForm
    transContentForm.id = IDTRANS_CONTENTFORM
    transContentForm.setCalcLayout WzFunc_new(FN_intAddTransporterContents_transContentForm_calcLayout)

    ' Add the close button
    W_BUTINIT_construct sButInit
    sButInit.formID = IDTRANS_CONTENTFORM
    sButInit.id = IDTRANS_CONTCLOSE
    Set sButInit.calcLayout = WzFunc_new(FN_intAddTransporterContents_closeButton_calcLayout)
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddTransporterContents = False
        Exit Function
    End If
    If bMultiPlayer Then
        'add the capacity label
        W_LABINIT_construct sLabInit
        sLabInit.formID = IDTRANS_CONTENTFORM
        sLabInit.id = IDTRANS_CAPACITY
        sLabInit.x = ToInt16(sButInit.x) - 40
        sLabInit.y = 0
        sLabInit.width = 16
        sLabInit.height = 16
        sLabInit.pText = "00/10"
        Set sLabInit.pCallback = WzFunc_new(FN_intUpdateTransCapacity)
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intAddTransporterContents = False
            Exit Function
        End If
    End If
    'add the Launch button if on a mission
    If onMission Then
        W_FORMINIT_construct sButFInit
        sButFInit.formID = IDTRANS_CONTENTFORM
        sButFInit.id = IDTRANS_LAUNCH
        sButFInit.style = WFORM_CLICKABLE Or WFORM_NOCLICKMOVE

        sButFInit.x = OBJ_STARTX
        sButFInit.y = (STAT_SLDY - 1) And &HFFFF&

        sButFInit.width = iV_GetImageWidth(IntImages, IMAGE_LAUNCHUP)
        sButFInit.height = iV_GetImageHeight(IntImages, IMAGE_LAUNCHUP)
        sButFInit.pTip = "Launch Transport"
        Set sButFInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)

        sButFInit.UserData = PACKDWORD_TRI(0, IMAGE_LAUNCHDOWN, IMAGE_LAUNCHUP)

        Set psForm = widgAddForm(psWScreen, sButFInit)
        If psForm Is Nothing Then
            intAddTransporterContents = False
            Exit Function
        End If
        Set help = New WidgetHelp
        help.setTitle "Launch Transport"
        help.addInteraction WidgetHelp_triggers(InteractionTriggers_PrimaryClick), "Launch the Transporter"
        psForm.setHelp help
    End If

    If Not intAddTransContentsForm() Then
        intAddTransporterContents = False
        Exit Function
    End If

    intAddTransporterContents = True
End Function

'This is used to display the transporter button and capacity when at the home base ONLY
Public Function intAddTransporterLaunch(ByVal psDroid As Long) As Boolean
    Dim capacity As Long, sButInit As W_FORMINIT, psForm As WIDGET, help As WidgetHelp, sLabInit As W_LABINIT
    Dim it As Long, psCurr As Long

    If bMultiPlayer Then
        intAddTransporterLaunch = True
        Exit Function
    End If

    'do this first so that if the interface is already up it syncs with this transporter
    'set up the static transporter
    psCurrTransporter = psDroid

    ' Check that neither the launch button nor the transport timer are currently up
    If Not widgGetFromID(psWScreen, IDTRANS_LAUNCH) Is Nothing _
       Or Not widgGetFromID(psWScreen, IDTRANTIMER_BUTTON) Is Nothing Then
        intAddTransporterLaunch = True
        Exit Function
    End If

    W_FORMINIT_construct sButInit       'needs to be a clickable form now
    sButInit.formID = 0
    sButInit.id = IDTRANS_LAUNCH
    sButInit.style = WFORM_CLICKABLE Or WFORM_NOCLICKMOVE
    sButInit.x = RET_X
    sButInit.y = ToInt16(TIMER_Y)
    sButInit.width = (10 + iV_GetImageWidth(IntImages, IMAGE_LAUNCHUP)) And &HFFFF&
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_LAUNCHUP)
    sButInit.pTip = "Launch Transport"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_LAUNCHDOWN, IMAGE_LAUNCHUP)
    Set psForm = widgAddForm(psWScreen, sButInit)
    If psForm Is Nothing Then
        intAddTransporterLaunch = False
        Exit Function
    End If
    Set help = New WidgetHelp
    help.setTitle "Launch Transport"
    help.addInteraction WidgetHelp_triggers(InteractionTriggers_PrimaryClick), "Launch the Transporter"
    psForm.setHelp help

    'add the capacity label
    W_LABINIT_construct sLabInit
    sLabInit.formID = IDTRANS_LAUNCH
    sLabInit.id = IDTRANS_CAPACITY
    sLabInit.x = ToInt16(sButInit.x + 40)
    sLabInit.y = 0
    sLabInit.width = 16
    sLabInit.height = 16
    sLabInit.pText = "00/10"
    Set sLabInit.pCallback = WzFunc_new(FN_intUpdateTransCapacity)
    If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
        intAddTransporterLaunch = False
        Exit Function
    End If

    'when full flash the transporter button
    ' (&&: psGroup only for a transporter)
    If psCurrTransporter <> 0 Then
        If aDroid(psCurrTransporter).psGroup <> 0 Then
            capacity = TRANSPORTER_CAPACITY
            it = aGroup(aDroid(psCurrTransporter).psGroup).psList.head
            Do While it <> 0
                psCurr = lnObj(it)
                If psCurr <> psCurrTransporter Then
                    capacity = capacity - transporterSpaceRequired(psCurr)
                End If
                it = lnNext(it)
            Loop
            If capacity = 0 Then        ' (UDWORD capacity <= 0)
                flashMissionButton IDTRANS_LAUNCH
            End If
        End If
    End If

    intAddTransporterLaunch = True
End Function

' Remove the Transporter Launch widget from the screen
Public Sub intRemoveTransporterLaunch()
    If Not widgGetFromID(psWScreen, IDTRANS_LAUNCH) Is Nothing Then
        widgDelete psWScreen, IDTRANS_LAUNCH
    End If
End Sub

' Add the Transporter Button form
Private Function intAddTransButtonForm() As Boolean
    Dim transForm As WIDGET, transList As WIDGET, objListWidth As Long, nextObjButtonId As Long, nextStatButtonId As Long
    Dim transIntDroidList As Long, it As Long, psDroid As Long, buttonHolder As WIDGET, statButton As WIDGET, objButton As WIDGET

    Set transForm = widgGetFromID(psWScreen, IDTRANS_FORM)

    ' Add the button form
    Set transList = IntListTabWidget_make()
    transForm.attach transList
    ListTabWidget_setChildSize transList, OBJ_BUTWIDTH, OBJ_BUTHEIGHT * 2
    ListTabWidget_setChildSpacing transList, OBJ_GAP, OBJ_GAP
    objListWidth = OBJ_BUTWIDTH * 5 + OBJ_GAP * 4
    transList.setGeometry (OBJ_BACKWIDTH - objListWidth) \ 2, TRANS_TABY, objListWidth, transForm.height() - TRANS_TABY

    ' Add the transporter and status buttons
    nextObjButtonId = IDTRANS_START
    nextStatButtonId = IDTRANS_STATSTART

    'add each button
    transIntDroidList = transInterfaceDroidList()
    it = transIntDroidList
    Do While it <> 0
        psDroid = lnObj(it)
        'only interested in Transporter droids
        If (DROID_isTransporter(psDroid) And (aDroid(psDroid).action = DACTION_TRANSPORTOUT Or _
            aDroid(psDroid).action = DACTION_TRANSPORTIN)) Or Not DROID_isTransporter(psDroid) Then
            GoTo continue_
        End If

        Set buttonHolder = WIDGET_new()
        transList.attach buttonHolder
        ListTabWidget_addWidgetToLayout transList, buttonHolder

        Set statButton = IntStatusButton_new()
        buttonHolder.attach statButton
        statButton.id = nextStatButtonId
        statButton.setGeometry 0, 0, OBJ_BUTWIDTH, OBJ_BUTHEIGHT

        Set objButton = IntObjectButton_new()
        buttonHolder.attach objButton
        objButton.id = nextObjButtonId
        objButton.setGeometry 0, OBJ_STARTY, OBJ_BUTWIDTH, OBJ_BUTHEIGHT

        ' Set the tip and add the button
        objButton.setTip droidGetName(psDroid)
        IntObjectButton_setObject objButton, psDroid

        'set the first Transporter to be the current one if not already set
        If psCurrTransporter = 0 Then
            psCurrTransporter = psDroid
        End If

        ' if the current droid matches psCurrTransporter lock the button
        If psDroid = psCurrTransporter Then
            objButton.setState WBUT_LOCK
            ListTabWidget_setCurrentPage transList, ListTabWidget_pages(transList) - 1
        End If

        'now do status button
        IntStatusButton_setObject statButton, 0

        ' Update the init struct for the next buttons
        nextObjButtonId = nextObjButtonId + 1
        If Not (nextObjButtonId < IDTRANS_END) Then ASSERT_FAILURE "nextObjButtonId < 9199", "src/transporter.cpp:478", "intAddTransButtonForm", "Too many Transporter buttons"

        nextStatButtonId = nextStatButtonId + 1
        If Not (nextStatButtonId < IDTRANS_STATEND) Then ASSERT_FAILURE "nextStatButtonId < 9299", "src/transporter.cpp:481", "intAddTransButtonForm", "Too many Transporter status buttons"
continue_:
        it = lnNext(it)
    Loop
    intAddTransButtonForm = True
End Function

' Add the Transporter Contents form
Private Function intAddTransContentsForm() As Boolean
    Dim contForm As WIDGET, contList As WIDGET, contListWidth As Long, nextButtonId As Long
    Dim it As Long, psDroid As Long, button As WIDGET

    Set contForm = widgGetFromID(psWScreen, IDTRANS_CONTENTFORM)

    ' Add the contents form
    Set contList = IntListTabWidget_make()
    contForm.attach contList
    ListTabWidget_setChildSize contList, OBJ_BUTWIDTH, OBJ_BUTHEIGHT
    ListTabWidget_setChildSpacing contList, OBJ_GAP, OBJ_GAP
    contListWidth = OBJ_BUTWIDTH * 2 + OBJ_GAP
    contList.setGeometry (contForm.width() - contListWidth) \ 2, TRANSCONT_TABY, contListWidth, contForm.height() - TRANSCONT_TABY

    ' Add the transporter contents buttons
    nextButtonId = IDTRANS_CONTSTART

    'add each button
    If psCurrTransporter = 0 Then
        intAddTransContentsForm = True
        Exit Function
    End If

    If Not (aDroid(psCurrTransporter).psGroup <> 0) Then ASSERT_FAILURE "psCurrTransporter->psGroup != nullptr", "src/transporter.cpp:509", "intAddTransContentsForm", "Null transporter group": intAddTransContentsForm = False: Exit Function

    it = aGroup(aDroid(psCurrTransporter).psGroup).psList.head
    Do While it <> 0
        psDroid = lnObj(it)
        If psDroid = psCurrTransporter Then
            Exit Do
        End If
        If aDroid(psDroid).selected <> 0 Then
            GoTo continue_  ' Droid is queued to be ejected from the transport, so don't display it.
        End If

        ' Set the tip and add the button
        Set button = IntTransportButton_new()
        contList.attach button
        button.id = nextButtonId
        button.setTip droidGetName(psDroid)
        IntTransportButton_setObject button, psDroid
        ListTabWidget_addWidgetToLayout contList, button

        ' Update the init struct for the next button
        nextButtonId = nextButtonId + 1
        If Not (nextButtonId < IDTRANS_CONTEND) Then ASSERT_FAILURE "nextButtonId < 9399", "src/transporter.cpp:532", "intAddTransContentsForm", "Too many Transporter Droid buttons"
continue_:
        it = lnNext(it)
    Loop
    intAddTransContentsForm = True
End Function

' Add the Droids back at home form
Private Function intAddDroidsAvailForm() As Boolean
    Dim Animate As Boolean, parent As WIDGET, transDroids As WIDGET, sButInit As W_BUTINIT, droidList As WIDGET
    Dim nextButtonId As Long, sBarInit As W_BARINIT, it As Long, psDroid As Long, button As WIDGET

    ' Is the form already up?
    Animate = True
    If Not widgGetFromID(psWScreen, IDTRANS_DROIDS) Is Nothing Then
        intRemoveTransDroidsAvailNoAnim
        Animate = False
    End If

    If intIsRefreshing() Then
        Animate = False
    End If

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/transporter.cpp:553", "intAddDroidsAvailForm", "Cannot be called for selectedPlayer: %u", selectedPlayer: intAddDroidsAvailForm = False: Exit Function

    Set parent = psWScreen.psForm

    ' Add the droids available form
    Set transDroids = IntFormAnimated_new(Animate)  ' Do not animate the opening, if the window was already open.
    parent.attach transDroids
    transDroids.id = IDTRANS_DROIDS
    transDroids.setCalcLayout WzFunc_new(FN_intAddDroidsAvailForm_transDroids_calcLayout)

    ' Add the close button
    W_BUTINIT_construct sButInit
    sButInit.formID = IDTRANS_DROIDS
    sButInit.id = IDTRANS_DROIDCLOSE
    Set sButInit.calcLayout = WzFunc_new(FN_intAddDroidsAvailForm_closeButton_calcLayout)
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddDroidsAvailForm = False
        Exit Function
    End If

    'now add the tabbed droids available form
    Set droidList = IntListTabWidget_make()
    transDroids.attach droidList
    droidList.id = IDTRANS_DROIDTAB
    droidList.setCalcLayout WzFunc_new(FN_intAddDroidsAvailForm_droidList_calcLayout)

    ' Add the droids available buttons
    nextButtonId = IDTRANS_DROIDSTART

    ' Add the state of repair bar for each droid
    W_BARINIT_construct sBarInit
    sBarInit.id = IDTRANS_REPAIRBARSTART
    sBarInit.x = STAT_TIMEBARX
    sBarInit.y = STAT_TIMEBARY
    sBarInit.width = STAT_PROGBARWIDTH
    sBarInit.height = STAT_PROGBARHEIGHT
    sBarInit.size = 50
    sBarInit.sCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MAJOR)
    sBarInit.sMinorCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MINOR)

    'add droids built before the mission
    it = mission_apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        'stop adding the buttons once IDTRANS_DROIDEND has been reached
        If nextButtonId = IDTRANS_DROIDEND Then
            Exit Do
        End If
        'don't add Transporter Droids!
        If Not DROID_isTransporter(psDroid) Then
            ' Set the tip and add the button
            Set button = IntTransportButton_new()
            droidList.attach button
            button.id = nextButtonId
            button.setTip droidGetName(psDroid)
            IntTransportButton_setObject button, psDroid
            ListTabWidget_addWidgetToLayout droidList, button

            'add bar to indicate stare of repair
            sBarInit.size = PERCENT_u(aDroid(psDroid).body, aDroid(psDroid).originalBody) And &HFFFF&
            If sBarInit.size > 100 Then
                sBarInit.size = 100
            End If

            sBarInit.formID = nextButtonId
            'sBarInit.iRange = TBAR_MAX_REPAIR;
            If widgAddBarGraph(psWScreen, sBarInit) Is Nothing Then
                intAddDroidsAvailForm = False
                Exit Function
            End If

            ' Update the init struct for the next button
            nextButtonId = nextButtonId + 1
            If Not (nextButtonId < IDTRANS_DROIDEND) Then ASSERT_FAILURE "nextButtonId < 9499", "src/transporter.cpp:642", "intAddDroidsAvailForm", "Too many Droids Built buttons"

            'and bar
            sBarInit.id = sBarInit.id + 1
        End If
        it = lnNext(it)
    Loop

    'reset which tab we were on
    ListTabWidget_setCurrentPage droidList, objMajor

    intAddDroidsAvailForm = True
End Function

'calculates how much space is remaining on the transporter - allows droids to take
'up different amount depending on their body size - currently all are set to one!
Public Function calcRemainingCapacity(ByVal psTransporter As Long) As Long
    Dim capacity As Long, it As Long, psDroid As Long, space As Long
    capacity = TRANSPORTER_CAPACITY

    ' If it's dead then just return 0.
    If isDead(psTransporter) Then
        calcRemainingCapacity = 0
        Exit Function
    End If

    ' If for some reason it doesn't have an associated psGroup (is it being recycled?), just return 0.
    If aDroid(psTransporter).psGroup = 0 Then
        calcRemainingCapacity = 0
        Exit Function
    End If

    it = aGroup(aDroid(psTransporter).psGroup).psList.head
    Do While it <> 0
        psDroid = lnObj(it)
        If psDroid = psTransporter Then
            Exit Do
        End If
        space = transporterSpaceRequired(psDroid)
        If Not (space > 0) Then ASSERT_FAILURE "space > 0", "src/transporter.cpp:680", "calcRemainingCapacity", "Invalid space required for %s", objInfo(psDroid)
        capacity = capacity - space
        it = lnNext(it)
    Loop

    If capacity < 0 Then
        capacity = 0
    End If

    calcRemainingCapacity = capacity
End Function

Public Function transporterIsEmpty(ByVal psTransporter As Long) As Boolean
    If Not (DROID_isTransporter(psTransporter)) Then ASSERT_FAILURE "psTransporter->isTransporter()", "src/transporter.cpp:694", "transporterIsEmpty", "Non-transporter droid given"

    ' Assume dead droids and non-transporter droids to be empty
    ' (||: psGroup only for a live transporter)
    If isDead(psTransporter) Then
        transporterIsEmpty = True
    ElseIf Not DROID_isTransporter(psTransporter) Then
        transporterIsEmpty = True
    ElseIf aGroup(aDroid(psTransporter).psGroup).psList.size = 0 Then
        transporterIsEmpty = True
    Else
        transporterIsEmpty = (objlist_front(aGroup(aDroid(psTransporter).psGroup).psList) = psTransporter)
    End If
End Function

Private Sub intSetTransCapacityLabel(ByVal Label As WIDGET)
    Dim capacity As Long, tmp As String
    If psCurrTransporter <> 0 Then
        capacity = calcRemainingCapacity(psCurrTransporter)

        'change round the way the remaining capacity is displayed - show 0/10 when empty now
        capacity = TRANSPORTER_CAPACITY - capacity

        tmp = Format$(capacity, "00") & "/10"   ' ssprintf(tmp, "%02d/10", capacity);
        Label.setString tmp
    End If
End Sub

'updates the capacity of the current Transporter
Public Sub intUpdateTransCapacity(ByVal psWidget As WIDGET)
    Dim Label As WIDGET
    Set Label = psWidget                    ' W_LABEL *Label = (W_LABEL *)psWidget;

    intSetTransCapacityLabel Label
End Sub

' Process return codes from the Transporter Screen
Public Sub intProcessTransporter(ByVal id As Long)
    Dim currID As Long, it As Long, psDroid As Long, psFound As Long

    If id >= IDTRANS_START And id <= IDTRANS_END Then
        ' A Transporter button has been pressed
        setCurrentTransporter id
        'refresh the Contents list
        intAddTransporterContents
    ElseIf id >= IDTRANS_CONTSTART And id <= IDTRANS_CONTEND Then
        'got to have a current transporter for this to work - and can't be flying
        ' (&&: transporterFlying only for a transporter)
        If psCurrTransporter <> 0 Then
            If Not transporterFlying(psCurrTransporter) Then
                currID = IDTRANS_CONTSTART
                ' std::find_if(groupList.begin(), groupList.end(), [&currID, id](DROID* psDroid) {...})
                psFound = 0
                it = aGroup(aDroid(psCurrTransporter).psGroup).psList.head
                Do While it <> 0
                    psDroid = lnObj(it)
                    If psDroid = psCurrTransporter Then
                        ' return false;
                    ElseIf aDroid(psDroid).selected <> 0 Then
                        ' return false;  // Already scheduled this droid for removal.
                    Else
                        If currID = id Then     ' return currID++ == id;
                            psFound = psDroid
                            Exit Do
                        End If
                        currID = currID + 1
                    End If
                    it = lnNext(it)
                Loop
                If psFound <> 0 Then
                    transporterRemoveDroid psCurrTransporter, psFound, ModeQueue
                End If
                'refresh the Contents list
                intAddTransporterContents
                If onMission Then
                    'refresh the Avail list
                    intAddDroidsAvailForm
                End If
            End If
        End If
    ElseIf id = IDTRANS_CLOSE Then
        intRemoveTransContent
        intRemoveTrans
        psCurrTransporter = 0
    ElseIf id = IDTRANS_CONTCLOSE Then
        intRemoveTransContent
    ElseIf id = IDTRANS_DROIDCLOSE Then
        intRemoveTransDroidsAvail
    ElseIf id >= IDTRANS_DROIDSTART And id <= IDTRANS_DROIDEND Then
        'got to have a current transporter for this to work - and can't be flying
        ' (&&: transporterFlying only for a transporter)
        If psCurrTransporter <> 0 Then
            If Not transporterFlying(psCurrTransporter) Then
                intTransporterAddDroid id
                'don't need to explicitly refresh here since intRefreshScreen()
                'is called by intTransporterAddDroid()
            End If
        End If
    End If
End Sub

' Remove the Transporter widgets from the screen
Public Sub intRemoveTrans(Optional ByVal skipIntModeReset As Boolean = False)
    Dim form As WIDGET

    ' Start the window close animation.
    Set form = widgGetFromID(psWScreen, IDTRANS_FORM)     ' (IntFormAnimated *)
    If Not form Is Nothing Then
        IntFormAnimated_closeAnimateDelete form
    End If

    intRemoveTransContent
    intRemoveTransDroidsAvail
    If Not skipIntModeReset Then
        intMode = INT_NORMAL
    End If
End Sub

' Remove the Transporter Content widgets from the screen w/o animation!
Public Sub intRemoveTransNoAnim(Optional ByVal skipIntModeReset As Boolean = False)
    'remove main screen
    widgDelete psWScreen, IDTRANS_FORM
    intRemoveTransContentNoAnim
    intRemoveTransDroidsAvailNoAnim
    If Not skipIntModeReset Then
        intMode = INT_NORMAL
    End If
End Sub

' Remove the Transporter Content widgets from the screen
Private Sub intRemoveTransContent()
    Dim form As WIDGET

    ' Start the window close animation.
    Set form = widgGetFromID(psWScreen, IDTRANS_CONTENTFORM)  ' (IntFormAnimated *)
    If Not form Is Nothing Then
        IntFormAnimated_closeAnimateDelete form
    End If
End Sub

' Remove the Transporter Content widgets from the screen w/o animation!
Private Sub intRemoveTransContentNoAnim()
    'remove main screen
    widgDelete psWScreen, IDTRANS_CONTENTFORM
End Sub

' Remove the Transporter Droids Avail widgets from the screen
Private Sub intRemoveTransDroidsAvail()
    Dim form As WIDGET, droidList As WIDGET

    ' Start the window close animation.
    Set form = widgGetFromID(psWScreen, IDTRANS_DROIDS)       ' (IntFormAnimated *)
    If Not form Is Nothing Then
        'remember which tab we were on
        Set droidList = widgGetFromID(psWScreen, IDTRANS_DROIDTAB)  ' (ListTabWidget *)
        objMajor = ListTabWidget_currentPage(droidList)
        IntFormAnimated_closeAnimateDelete form
    End If
End Sub

' Remove the Transporter Droids Avail widgets from the screen w/o animation!
Private Sub intRemoveTransDroidsAvailNoAnim()
    Dim form As WIDGET, droidList As WIDGET

    Set form = widgGetFromID(psWScreen, IDTRANS_DROIDS)       ' (IntFormAnimated *)
    If Not form Is Nothing Then
        'remember which tab we were on
        Set droidList = widgGetFromID(psWScreen, IDTRANS_DROIDTAB)  ' (ListTabWidget *)
        objMajor = ListTabWidget_currentPage(droidList)

        'remove main screen
        widgDelete_widget form
    End If
End Sub

'sets psCurrTransporter
Private Sub setCurrentTransporter(ByVal id As Long)
    Dim psDroid As Long, currID As Long, transIntDroidList As Long, it As Long, psCurr As Long

    psDroid = 0
    psCurrTransporter = 0
    currID = IDTRANS_START

    'loop thru all the droids to find the selected one
    transIntDroidList = transInterfaceDroidList()
    it = transIntDroidList
    Do While it <> 0
        psCurr = lnObj(it)
        If DROID_isTransporter(psCurr) And _
           (aDroid(psCurr).action <> DACTION_TRANSPORTOUT And _
            aDroid(psCurr).action <> DACTION_TRANSPORTIN) Then
            If currID = id Then
                psDroid = psCurr
                Exit Do
            End If
            currID = currID + 1
        End If
        it = lnNext(it)
    Loop
    If psDroid <> 0 Then
        psCurrTransporter = psDroid
        'set the data for the transporter timer
        widgSetUserData psWScreen, IDTRANTIMER_DISPLAY, psCurrTransporter
    End If
End Sub

'removes a droid from the group associated with the transporter
Public Sub transporterRemoveDroid(ByVal psTransport As Long, ByVal psDroid As Long, ByVal mode As QUEUE_MODE)
    Dim droidPos As Vector2i, psGroup As Long

    If Not (psTransport <> 0 And psDroid <> 0 And psTransport <> psDroid) Then ASSERT_FAILURE "psTransport != nullptr && psDroid != nullptr && psTransport != psDroid", "src/transporter.cpp:912", "transporterRemoveDroid", "Something NULL or unloading transporter from itself": Exit Sub

    If bMultiMessages And mode = ModeQueue Then
        sendDroidDisembark psTransport, psDroid
        aDroid(psDroid).selected = 1        ' Remove from interface.
        Exit Sub
    End If

    'if we're offWorld we can't pick a tile without swapping the map
    'pointers - can't be bothered so just do this...
    If onMission Then
        aDroid(psDroid).pos.x = INVALID_XY
        aDroid(psDroid).pos.y = INVALID_XY
    Else
        If bMultiPlayer Then
            'set the units next to the transporter's current location
            droidPos.x = map_coord(aDroid(psTransport).pos.x)
            droidPos.y = map_coord(aDroid(psTransport).pos.y)
        Else
            'pick a tile because save games won't remember where the droid was when it was loaded
            droidPos.x = map_coord(getLandingX(0))
            droidPos.y = map_coord(getLandingY(0))
        End If
        If Not pickATileGen_pos(droidPos, LOOK_FOR_EMPTY_TILE, PAT_zonedPAT) Then
            ASSERT_FAILURE "false", "src/transporter.cpp:943", "transporterRemoveDroid", "Unable to find a valid location"
        End If
        droidSetPosition psDroid, world_coord(droidPos.x), world_coord(droidPos.y)
        updateDroidOrientation psDroid
    End If

    ' remove it from the transporter group
    DROID_GROUP_remove aDroid(psDroid).psGroup, psDroid

    'add it back into apsDroidLists
    If onMission Then
        addDroid psDroid, mission_apsDroidLists
    Else
        ' add the droid back onto the droid list
        addDroid psDroid, apsDroidLists
    End If

    If aDroid(psDroid).pos.x <> INVALID_XY Then
        ' We can update the orders now, since everyone has been
        ' notified of the droid exiting the transporter
        updateDroidOrientation psDroid
    End If
    'initialise the movement data
    initDroidMovement psDroid
    'reset droid orders
    orderDroid psDroid, DORDER_STOP, ModeImmediate
    ' check if it is a commander
    If aDroid(psDroid).droidType = DROID_COMMAND Then
        psGroup = grpCreate()
        DROID_GROUP_add psGroup, psDroid
    End If
    aDroid(psDroid).selected = 1

    If calcRemainingCapacity(psTransport) <> 0 Then
        'make sure the button isn't flashing
        stopMissionButtonFlash IDTRANS_LAUNCH
    End If

    ' Fire off disembark event
    triggerEvent TRIGGER_TRANSPORTER_DISEMBARKED, psTransport
End Sub

'adds a droid to the current transporter via the interface
Private Sub intTransporterAddDroid(ByVal id As Long)
    Dim currID As Long, transIntDroidList As Long, droidIt As Long

    If Not (psCurrTransporter <> 0) Then ASSERT_FAILURE "psCurrTransporter != nullptr", "src/transporter.cpp:996", "intTransporterAddDroid", "intTransporterAddUnit:can't remove units"

    currID = IDTRANS_DROIDSTART
    transIntDroidList = transInterfaceDroidList()
    If transIntDroidList = 0 Then
        Exit Sub
    End If
    droidIt = transIntDroidList
    Do While droidIt <> 0
        If Not DROID_isTransporter(lnObj(droidIt)) Then
            If currID = id Then
                Exit Do
            End If
            currID = currID + 1
        End If
        droidIt = lnNext(droidIt)
    Loop
    If droidIt <> 0 Then
        transporterAddDroid psCurrTransporter, lnObj(droidIt)
    End If
End Sub

'Adds a droid to the transporter, removing it from the world
Public Sub transporterAddDroid(ByVal psTransporter As Long, ByVal psDroidToAdd As Long)
    Dim bDroidRemoved As Boolean, it As Long, psOtherDroid As Long

    If Not (psTransporter <> 0) Then ASSERT_FAILURE "psTransporter != nullptr", "src/transporter.cpp:1028", "transporterAddDroid", "Was passed a NULL transporter"
    If Not (psDroidToAdd <> 0) Then ASSERT_FAILURE "psDroidToAdd != nullptr", "src/transporter.cpp:1029", "transporterAddDroid", "Was passed a NULL droid, can't add to transporter"

    If psTransporter = 0 Or psDroidToAdd = 0 Then
        debug_ 1033, LOG_ERROR, "transporterAddDroid", "We can't add the unit to the transporter!"
        Exit Sub
    End If
    ' check for space
    If Not checkTransporterSpace(psTransporter, psDroidToAdd) Then
        If bMultiPlayer Then
            ' search for the nearest transporter if the current one is already full
            ' (&&: checkTransporterSpace (it may flash the button) and droidSqDist only for a transporter)
            it = apsDroidLists(aDroid(psTransporter).player).head
            Do While it <> 0
                psOtherDroid = lnObj(it)
                If DROID_isTransporter(psOtherDroid) Then
                    If checkTransporterSpace(psOtherDroid, psDroidToAdd) Then
                        If droidSqDist(psOtherDroid, psTransporter) < MAX_NEAREST_TRANSPORT_SQ_DIST Then
                            If aDroid(psOtherDroid).droidType = DROID_TRANSPORTER And Not DROID_isCyborg(psDroidToAdd) Then
                                GoTo continue_  ' Only send cyborgs to a cyborg transporter.
                            End If
                            orderDroidObj psDroidToAdd, DORDER_EMBARK, psOtherDroid, ModeQueue
                            Exit Sub
                        End If
                    End If
                End If
continue_:
                it = lnNext(it)
            Loop
        End If
        If Not onMission And DROID_isVtol(psDroidToAdd) Then
            moveStopDroid psDroidToAdd ' So VTOLs are put into a MOVEHOVER status from the prior moveReallyStopDroid() in order.cpp's embark section.
        End If
        If aDroid(psTransporter).player = selectedPlayer Then
            audio_PlayBuildFailedOnce
            If ULT(lastTransportIsFullMsgTime + MAX_TRANSPORT_FULL_MESSAGE_PAUSE, gameTime) Then
                addConsoleMessage "There is not enough room in the Transport!", DEFAULT_JUSTIFY, selectedPlayer
                lastTransportIsFullMsgTime = gameTime
            End If
        End If
        Exit Sub
    End If

    resetObjectAnimationState psDroidToAdd

    If onMission Then
        ' removing from droid mission list
        bDroidRemoved = droidRemove(psDroidToAdd, mission_apsDroidLists)
    Else
        ' removing from droid list
        bDroidRemoved = droidRemove(psDroidToAdd, apsDroidLists)
    End If

    If bDroidRemoved Then
        ' adding to transporter unit's group list
        DROID_GROUP_add aDroid(psTransporter).psGroup, psDroidToAdd
        aDroid(psDroidToAdd).selected = 0   ' Display in transporter interface.
    Else
        debug_ 1094, LOG_ERROR, "transporterAddDroid", "droid %d not found, so nothing added to transporter!", aDroid(psDroidToAdd).id
    End If
    If onMission Then
        visRemoveVisibilityOffWorld psDroidToAdd
    Else
        visRemoveVisibility psDroidToAdd
    End If
    fpathRemoveDroidData aDroid(psDroidToAdd).id

    ' This is called by droidRemove. But we still need to refresh after adding to the transporter group.
    intRefreshScreen
End Sub

'check to see if the droid can fit on the Transporter - return true if fits
Public Function checkTransporterSpace(ByVal psTransporter As Long, ByVal psAssigned As Long, Optional ByVal mayFlash As Boolean = True) As Boolean
    Dim capacity As Long, it As Long, psDroid As Long  ' capacity: UDWORD

    If Not (psTransporter <> 0) Then ASSERT_FAILURE "psTransporter != nullptr", "src/transporter.cpp:1115", "checkTransporterSpace", "Invalid droid pointer": checkTransporterSpace = False: Exit Function
    If Not (psAssigned <> 0) Then ASSERT_FAILURE "psAssigned != nullptr", "src/transporter.cpp:1116", "checkTransporterSpace", "Invalid droid pointer": checkTransporterSpace = False: Exit Function
    If Not (DROID_isTransporter(psTransporter)) Then ASSERT_FAILURE "psTransporter->isTransporter()", "src/transporter.cpp:1117", "checkTransporterSpace", "Droid is not a Transporter": checkTransporterSpace = False: Exit Function
    If Not (aDroid(psTransporter).psGroup <> 0) Then ASSERT_FAILURE "psTransporter->psGroup != nullptr", "src/transporter.cpp:1118", "checkTransporterSpace", "transporter doesn't have a group": checkTransporterSpace = False: Exit Function

    'work out how much space is currently left
    capacity = TRANSPORTER_CAPACITY
    it = aGroup(aDroid(psTransporter).psGroup).psList.head
    Do While it <> 0
        psDroid = lnObj(it)
        If psDroid = psTransporter Then
            Exit Do
        End If
        capacity = capacity - transporterSpaceRequired(psDroid)
        it = lnNext(it)
    Loop
    If Not ULT(capacity, transporterSpaceRequired(psAssigned)) Then     ' (unsigned capacity >= ...)
        'when full flash the transporter button
        If mayFlash And capacity - transporterSpaceRequired(psAssigned) = 0 Then
            flashMissionButton IDTRANS_LAUNCH
        End If
        checkTransporterSpace = True
    Else
        checkTransporterSpace = False
    End If
End Function

'returns the space the droid occupies on a transporter based on the body size
Public Function transporterSpaceRequired(ByVal psDroid As Long) As Long
    ' all droids are the same weight for campaign games.
    ' TODO - move this into a droid flag
    If bMultiPlayer Then
        transporterSpaceRequired = asBodyStats(DROID_getBodyStats(psDroid)).size + 1
    Else
        transporterSpaceRequired = 1
    End If
End Function

'sets which list of droids to use for the transporter interface
' (VB) DroidList *: the list's first node (0 for nullptr)
Private Function transInterfaceDroidList() As Long
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/transporter.cpp:1156", "transInterfaceDroidList", "Cannot be called for selectedPlayer: %u", selectedPlayer: transInterfaceDroidList = 0: Exit Function
    If onMission Then
        transInterfaceDroidList = mission_apsDroidLists(selectedPlayer).head
    Else
        transInterfaceDroidList = apsDroidLists(selectedPlayer).head
    End If
End Function

Public Function transporterGetLaunchTime() As Long
    transporterGetLaunchTime = g_iLaunchTime
End Function

Public Sub transporterSetLaunchTime(ByVal time_ As Long)
    g_iLaunchTime = time_
End Sub

'launches the defined transporter to the offworld map
Public Function launchTransporter(ByVal psTransporter As Long) As Boolean
    Dim iX As Long, iY As Long          ' UDWORD

    'close the interface
    intResetScreen True

    'this launches the mission if on homebase when the button is pressed
    If Not onMission Then
        'tell the transporter to move to the new offworld location
        missionGetTransporterExit aDroid(psTransporter).player, iX, iY
        orderDroidLoc psTransporter, DORDER_TRANSPORTOUT, iX, iY, ModeQueue
        'g_iLaunchTime = gameTime;
        transporterSetLaunchTime gameTime
    'otherwise just launches the Transporter
    Else
        If Not DROID_isTransporter(psTransporter) Then
            ASSERT_FAILURE "false", "src/transporter.cpp:1199", "launchTransporter", "Invalid Transporter Droid"
            launchTransporter = False
            Exit Function
        End If

        orderDroid psTransporter, DORDER_TRANSPORTIN, ModeImmediate
        ' set action transporter waits for timer
        actionDroid psTransporter, DACTION_TRANSPORTWAITTOFLYIN

        missionSetReinforcementTime gameTime
    End If

    launchTransporter = True
End Function

'checks how long the transporter has been travelling to see if it should
'have arrived - returns true when there
Public Function updateTransporter(ByVal psTransporter As Long) As Boolean
    If Not (psTransporter <> 0) Then ASSERT_FAILURE "psTransporter != nullptr", "src/transporter.cpp:1219", "updateTransporter", "Invalid droid pointer": updateTransporter = True: Exit Function
    If Not (DROID_isTransporter(psTransporter)) Then ASSERT_FAILURE "psTransporter->isTransporter()", "src/transporter.cpp:1220", "updateTransporter", "Invalid droid type": updateTransporter = True: Exit Function

    'if not moving to mission site, exit
    If aDroid(psTransporter).action <> DACTION_TRANSPORTOUT And aDroid(psTransporter).action <> DACTION_TRANSPORTIN Then
        updateTransporter = True
        Exit Function
    End If

    ' moving to a location
    ' if we're coming back for more droids then we want the transporter to
    ' fly to edge of map before turning round again
    If aDroid(psTransporter).sMove.Status = MOVEINACTIVE Or _
       aDroid(psTransporter).sMove.Status = MOVEHOVER Or _
       (aDroid(psTransporter).action = DACTION_TRANSPORTOUT And Not missionIsOffworld() And _
        (ULT(transporterGetLaunchTime() + TRANSPORTOUT_TIME, gameTime)) And _
        Not getDroidsToSafetyFlag()) Then
        audio_StopObjTrack psTransporter, ID_SOUND_BLIMP_FLIGHT
        If aDroid(psTransporter).action = DACTION_TRANSPORTIN Then
            ' !!!! GJ Hack - should be landing audio !!!!
            audio_PlayObjDynamicTrack psTransporter, ID_SOUND_BLIMP_TAKE_OFF, 0
        End If

        If Not bFirstTransporter And missionForReInforcements() And _
           aDroid(psTransporter).action = DACTION_TRANSPORTIN And _
           aDroid(psTransporter).player = selectedPlayer Then
            'play reinforcements have arrived message
            audio_QueueTrackPos ID_SOUND_TRANSPORT_LANDING, _
                                aDroid(psTransporter).pos.x, aDroid(psTransporter).pos.y, aDroid(psTransporter).pos.z
            addConsoleMessage "Reinforcements landing", LEFT_JUSTIFY, SYSTEM_MESSAGE
            'reset the data for the transporter timer
            widgSetUserData psWScreen, IDTRANTIMER_DISPLAY, Empty  ' (void *)nullptr
            updateTransporter = True
            Exit Function
        End If

        'Remove visibility so tiles are not bright around where the transporter left the map
        If aDroid(psTransporter).action <> DACTION_TRANSPORTIN Then
            visRemoveVisibility psTransporter
        End If

        ' Got to destination
        aDroid(psTransporter).action = DACTION_NONE

        'reset the flag to trigger the audio message
        bFirstTransporter = False

        updateTransporter = True
        Exit Function
    End If

    'not arrived yet...
    updateTransporter = False
End Function

'process the launch transporter button click
Public Sub processLaunchTransporter()
    Dim capacity As Long, psForm As WIDGET  ' capacity: UDWORD; psForm: W_CLICKFORM *
    capacity = TRANSPORTER_CAPACITY

    'launch the Transporter
    If psCurrTransporter <> 0 Then
        'check there is something on the transporter
        capacity = calcRemainingCapacity(psCurrTransporter)
        If capacity <> TRANSPORTER_CAPACITY Then
            'make sure the button doesn't flash once launched
            stopMissionButtonFlash IDTRANS_LAUNCH
            'disable the form so can't add any more droids into the transporter
            Set psForm = widgGetFromID(psWScreen, IDTRANS_LAUNCH)
            If Not psForm Is Nothing Then
                psForm.setState WBUT_LOCK
            End If

            'disable the form so can't add any more droids into the transporter
            Set psForm = widgGetFromID(psWScreen, IDTRANTIMER_BUTTON)
            If Not psForm Is Nothing Then
                psForm.setState WBUT_LOCK
            End If

            launchTransporter psCurrTransporter
            'set the data for the transporter timer
            widgSetUserData psWScreen, IDTRANTIMER_DISPLAY, psCurrTransporter

            ' executeFnAndProcessScriptQueuedRemovals([]() { triggerEvent(TRIGGER_TRANSPORTER_LAUNCH, psCurrTransporter); });
            triggerEvent TRIGGER_TRANSPORTER_LAUNCH, psCurrTransporter
            wzapi_processScriptQueuedObjectRemovals
        End If
    End If
End Sub

Public Function bobTransporterHeight() As Long
    ' Because 4320/12 = 360 degrees
    ' this gives us a bob frequency of 4.32 seconds.
    ' we scale amplitude to 10 (world coordinate metric).
    ' we need to use 360 degrees and not 180, as otherwise
    ' it will not 'bounce' off the top _and_ bottom of
    ' it's movemment arc.

    bobTransporterHeight = iSinSR(gameTime, 4320, 10)
End Function

'causes one of the mission buttons (Launch Button or Mission Timer) to start flashing
Public Sub flashMissionButton(ByVal buttonID As Long)
    Dim psForm As WIDGET

    'get the button from the id
    Set psForm = widgGetFromID(psWScreen, buttonID)
    If Not psForm Is Nothing Then
        Select Case buttonID
        Case IDTRANS_LAUNCH
            psForm.UserData = PACKDWORD_TRI(1, IMAGE_LAUNCHDOWN, IMAGE_LAUNCHUP)
        Case IDTIMER_FORM
            psForm.UserData = PACKDWORD_TRI(1, IMAGE_MISSION_CLOCK, IMAGE_MISSION_CLOCK_UP)
        Case Else
            'do nothing other than in debug
            ASSERT_FAILURE "false", "src/transporter.cpp:1343", "flashMissionButton", "flashMissionButton: Unknown button ID"
        End Select
    End If
End Sub

'stops one of the mission buttons (Launch Button or Mission Timer) flashing
Public Sub stopMissionButtonFlash(ByVal buttonID As Long)
    Dim psForm As WIDGET

    'get the button from the id
    Set psForm = widgGetFromID(psWScreen, buttonID)
    If Not psForm Is Nothing Then
        Select Case buttonID
        Case IDTRANS_LAUNCH
            psForm.UserData = PACKDWORD_TRI(0, IMAGE_LAUNCHDOWN, IMAGE_LAUNCHUP)
        Case IDTIMER_FORM
            psForm.UserData = PACKDWORD_TRI(0, IMAGE_MISSION_CLOCK, IMAGE_MISSION_CLOCK_UP)
        Case Else
            'do nothing other than in debug
            ASSERT_FAILURE "false", "src/transporter.cpp:1366", "stopMissionButtonFlash", "stopMissionButtonFlash: Unknown button ID"
        End Select
    End If
End Sub

'called when a Transporter has arrived back at the LZ when sending droids to safety
Public Sub resetTransporter()
    Dim psForm As WIDGET                ' W_CLICKFORM *

    'enable the form so can add more droids into the transporter
    Set psForm = widgGetFromID(psWScreen, IDTRANS_LAUNCH)
    If Not psForm Is Nothing Then
        psForm.setState 0
    End If
End Sub

'checks the order of the droid to see if its currently flying
Public Function transporterFlying(ByVal psTransporter As Long) As Boolean
    If Not (psTransporter <> 0) Then ASSERT_FAILURE "psTransporter != nullptr", "src/transporter.cpp:1388", "transporterFlying", "Invalid droid pointer": transporterFlying = False: Exit Function
    If Not (DROID_isTransporter(psTransporter)) Then ASSERT_FAILURE "psTransporter->isTransporter()", "src/transporter.cpp:1389", "transporterFlying", "Droid is not a Transporter": transporterFlying = False: Exit Function

    transporterFlying = aDroid(psTransporter).order.typ = DORDER_TRANSPORTOUT Or _
                        aDroid(psTransporter).order.typ = DORDER_TRANSPORTIN Or _
                        aDroid(psTransporter).order.typ = DORDER_TRANSPORTRETURN Or _
                        (bMultiPlayer And aDroid(psTransporter).order.typ = DORDER_MOVE) Or _
                        (bMultiPlayer And aDroid(psTransporter).order.typ = DORDER_DISEMBARK) Or _
                        (bMultiPlayer And aDroid(psTransporter).order.typ = DORDER_NONE And _
                         aDroid(psTransporter).sMove.iVertSpeed <> 0)
    ' (in multiPlayer mode the Transporter can be moved around, moved and emptied; descending still counts as flying)
End Function

'initialise the flag to indicate the first transporter has arrived - set in startMission()
Public Sub initFirstTransporterFlag()
    bFirstTransporter = True
End Sub

'--------------------------------------------------------------------------
' (VB) the lambdas and callbacks of the transporter's interface

Public Function transporter_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, droidListWidth As Long
    Select Case f.id
    Case FN_intAddTransporter_transForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry TRANS_X(), TRANS_Y() - IIf(getGroupButtonEnabled(), 80, 0), TRANS_WIDTH, TRANS_HEIGHT

    Case FN_intAddTransporterContents_transContentForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry TRANSCONT_X, TRANSCONT_Y(), TRANSCONT_WIDTH, TRANSCONT_HEIGHT

    Case FN_intAddTransporterContents_closeButton_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry STAT_WIDTH - CLOSE_WIDTH, 0, CLOSE_WIDTH, CLOSE_HEIGHT

    Case FN_intUpdateTransCapacity
        ' (psWidget, the W_CONTEXT's xOffset, yOffset, mx, my)
        intUpdateTransCapacity args(0)

    Case FN_intAddDroidsAvailForm_transDroids_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry TRANSDROID_X(), TRANSDROID_Y(), TRANSDROID_WIDTH, TRANSDROID_HEIGHT

    Case FN_intAddDroidsAvailForm_closeButton_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry TRANSDROID_WIDTH - CLOSE_WIDTH, 0, CLOSE_WIDTH, CLOSE_HEIGHT

    Case FN_intAddDroidsAvailForm_droidList_calcLayout
        Set psWidget = args(0)
        ' IntListTabWidget *droidList = static_cast<IntListTabWidget *>(psWidget); assert(droidList != nullptr);
        ListTabWidget_setChildSize psWidget, OBJ_BUTWIDTH, OBJ_BUTHEIGHT
        ListTabWidget_setChildSpacing psWidget, OBJ_GAP, OBJ_GAP
        droidListWidth = OBJ_BUTWIDTH * 2 + OBJ_GAP
        psWidget.setGeometry (TRANSDROID_WIDTH - droidListWidth) \ 2, AVAIL_STARTY + 15, droidListWidth, TRANSDROID_HEIGHT - (AVAIL_STARTY + 15)
    End Select
End Function
