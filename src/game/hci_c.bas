Attribute VB_Name = "hci_c"
' src/hci.cpp, hci.h: the in-game interface's core: its life cycle (intInitialise,
' interfaceShutDown), the reticule and its buttons (retbutstats), the power bar, the
' in-game options button (class W_INGAMEOPTIONS_BUTTON: its fields are the class of the
' same name), the frame's run and display (intRunWidgets, intDisplayWidgets), the
' structure placement glue, the notify hooks the simulation calls, the debug placement
' menu and the proximity buttons; the UI's list of templates (localTemplates) and the
' widget screen psWScreen.
' The panels (build, manufacture, research, commander, groups) are the controllers of
' src/hci/*.cpp (objects_stats_c's BaseObjectsController: interfaceController); the
' replay controller overlay (ReplayControllerWidget, createReplayControllerOverlay and
' replayOverlayScreen) is hci_replay_c's.
' std::list<DROID_TEMPLATE> localTemplates holds template_c pool handles, each its own
' copy, with the count localTemplates_size. std::vector<DROID_TEMPLATE *> apsTemplateList
' is the array of pool handles with apsTemplateList_size (apsTemplateList_clear,
' apsTemplateList_push_back); the other stats lists hold BASE_STATS * handles
' (ppsStatsList, apsComponentList, apsExtraSysList) or store indices
' (apsStructStatsList, apsFeatureStatsList), malloc being a ReDim.
' The INTMODE enum is INTMODE_ (VB would take it for the variable intMode); hci.h's
' macros that read the video buffer's size (E_W ... POW_Y) are functions; a
' playerCallbackFunc is the script's callback object (wzapi's getNamedScriptCallback:
' .invoke player), Nothing = nullptr. interfaceStructList copies the list's header
' (a read-only view of its nodes) and returns whether there is one. The tables'
' initialisers are hci_c_init's.
' (VB) widget class W_INGAMEOPTIONS_BUTTON : W_BUTTON
Option Explicit

' Reticule button indecies.
Public Enum RETBUT_INDEX
    RETBUT_CANCEL
    RETBUT_FACTORY
    RETBUT_RESEARCH
    RETBUT_BUILD
    RETBUT_DESIGN
    RETBUT_INTELMAP
    RETBUT_COMMAND
End Enum

Public Const BASE_COORDS_X As Long = 640
Public Const BASE_COORDS_Y As Long = 460

Public Const IDRET_FORM As Long = 1             ' The reticule form
Public Const IDRET_OPTIONS As Long = 2          ' option button
Public Const IDRET_BUILD As Long = 3            ' build button
Public Const IDRET_MANUFACTURE As Long = 4      ' manufacture button
Public Const IDRET_RESEARCH As Long = 5         ' research button
Public Const IDRET_INTEL_MAP As Long = 6        ' intelligence map button
Public Const IDRET_DESIGN As Long = 7           ' design droids button
Public Const IDRET_CANCEL As Long = 8           ' central cancel button
Public Const IDRET_COMMAND As Long = 9          ' command droid button
Public Const IDRET_TRANSPORTER As Long = 10     ' transporter button
Public Const IDPOW_POWERBAR_T As Long = 102     ' power bar - trough
Public Const IDTRANTIMER_BUTTON As Long = 11012 'transporter button on timer display


' Object screen IDs

Public Const IDOBJ_FORM As Long = 3000          ' The object back form for build/manufacture/research
Public Const IDOBJ_CLOSE As Long = 3001         ' The form for the close button
Public Const IDOBJ_TABFORM As Long = 3500       ' The object tab form for build/manufacture/research
Public Const IDOBJ_GROUP As Long = 3502         ' The tab for groups

Public Const IDSTAT_FORM As Long = 14000        ' The stats form for structure/droid/research type
Public Const IDSTAT_CLOSE As Long = 14003       ' The stats close box
Public Const IDSTAT_TABFORM As Long = 14004     ' The tab form with the stats buttons
Public Const IDSTAT_START As Long = 14100       ' The first stats ID
Public Const IDSTAT_END As Long = 14999         ' The last stats ID enough for 899 things
Public Const IDSTAT_TIMEBARSTART As Long = 16205000
Public Const IDSTAT_TIMEBAREND As Long = 16205999
Public Const IDSTAT_OBSOLETE_BUTTON As Long = 4406
Public Const IDSTAT_PRODSTART As Long = 4600
Public Const IDSTAT_PRODEND As Long = 4699

' Reticule position.
Public Const RET_X As Long = 6
Public Const RET_FORMWIDTH As Long = 132
Public Const RET_FORMHEIGHT As Long = 132

' chat stuff
Public Const CHAT_CONSOLEBOX As Long = &H4A001
Public Const CHAT_EDITBOX As Long = &H4A004
Public Const CHAT_TEAM As Long = &H1
Public Const CHAT_GLOB As Long = &H2
Public Const CHAT_CONSOLEBOXY As Long = 1
Public Const CHAT_CONSOLEBOXW As Long = 430
Public Const CHAT_CONSOLEBOXH As Long = 20

' Option positions
Public Const OPT_GAP As Long = 5

' Object screen position. (aka where the factories, research builds show up) [right of command retile]
Public Const BASE_GAP As Long = 6
Public Const OBJ_BACKWIDTH As Long = 320        '316		// Width of object screen back form.
Public Const OBJ_BACKHEIGHT As Long = 115       ' Height of object screen back form.

Public Const GROUP_BACKWIDTH As Long = OBJ_BACKWIDTH
Public Const GROUP_BACKHEIGHT As Long = OBJ_BACKHEIGHT - 40

' Build screen positions
Public Const OBJ_TABY As Long = 6               ' Y coord of object screen tab form.
Public Const OBJ_WIDTH As Long = 316            '312//310	// Width of object screen tab form.
Public Const OBJ_HEIGHT As Long = 112           ' Height of object screen tab form.
Public Const OBJ_GAP As Long = 2                ' Gap between buttons.
Public Const OBJ_STARTX As Long = 2             ' Offset of first obj button from left of tab form.
Public Const OBJ_STARTY As Long = 42            '44	// Offset of first obj button from top of tab form.
Public Const OBJ_STATSTARTY As Long = 0

'slider bar positions

Public Const STAT_SLDX As Long = 8              ' Slider x.
Public Const STAT_SLDY As Long = 4              ' Slider y.

Public Const STAT_SLDWIDTH As Long = 70         ' Slider width.

' Power bar position.
Public Const POW_BARWIDTH As Long = 308

' close button data
Public Const CLOSE_WIDTH As Long = 15
Public Const CLOSE_HEIGHT As Long = 15
Public Const CLOSE_SIZE As Long = 15

' Stat screen position. [aka, the 'dialog' of items you research/create]
Public Const STAT_X As Long = 6
Public Const STAT_WIDTH As Long = RET_FORMWIDTH ' Width of the tab form.
Public Const STAT_HEIGHT As Long = 322          ' Height of the tab form.
Public Const STAT_TABWIDTH As Long = 15
Public Const STAT_TABFORMX As Long = 0          ' Offset of the tab form within the main form.
Public Const STAT_TABFORMY As Long = 18         ' Offset of the tab form within the main form.

Public Const POWERPOINTS_DROIDDIV As Long = 5

Public Const OBJ_BUTWIDTH As Long = 60          ' Button width.
Public Const OBJ_BUTHEIGHT As Long = 46         ' Button height.

Public Const OBJ_TEXTX As Long = 2

Public Const OBJ_T1TEXTY As Long = 2
Public Const OBJ_T2TEXTY As Long = 14
Public Const OBJ_T3TEXTY As Long = 26
Public Const OBJ_B1TEXTY As Long = 8

Public Const STAT_SLD_OX As Long = 0            ' Stat window slider offset.
Public Const STAT_SLD_OY As Long = 0
Public Const STAT_SLDSTOPS As Long = 10         ' Slider number of stops.

Public Const STAT_PROGBARX As Long = 3
Public Const STAT_PROGBARY As Long = 36
Public Const STAT_PROGBARWIDTH As Long = OBJ_BUTWIDTH - 8
Public Const STAT_PROGBARHEIGHT As Long = 4
Public Const STAT_TIMEBARX As Long = 3
Public Const STAT_TIMEBARY As Long = OBJ_BUTHEIGHT - STAT_PROGBARHEIGHT - 3
Public Const STAT_POWERBARX As Long = 3
Public Const STAT_POWERBARY As Long = OBJ_BUTHEIGHT - STAT_PROGBARHEIGHT - 6

' maximum array sizes
Public Const MAXSTRUCTURES As Long = 65535      'bumped up from 80.  NOTE: was used for max # in build menus.
Public Const MAXRESEARCH As Long = 65535        'was 80 topic displayed   "           "
Public Const MAXFEATURES As Long = 200          'was 80.
Public Const MAXCOMPONENT As Long = 200
Public Const MAXEXTRASYS As Long = 80

Public Enum INTMODE_
    INT_NORMAL          ' Standard mode (just the reticule)

    INT_OPTION          ' Option screen
    INT_EDIT            ' Edit mode

    INT_EDITSTAT        ' Stat screen up for placing objects
    INT_OBJECT          ' Object screen
    INT_STAT            ' Object screen with stat screen
    INT_CMDORDER        ' Object screen with command droids and orders screen
    INT_DESIGN          ' Design screen
    INT_INTELMAP        ' Intelligence Map
    INT_ORDER
    INT_INGAMEOP        ' in game options.
    INT_TRANSPORTER     'Loading/unloading a Transporter
    INT_MISSIONRES      ' Results of a mission display.
    INT_MULTIMENU       ' multiplayer only, player stats etc...
    INT_CDCHANGE        ' CD Change message box
    INT_POPUPMSG        ' Adds a popup message to user

    INT_MAXMODE         'leave as last so we can start the objMode at this value
End Enum

' Which type of object screen is being displayed. Starting value is where the intMode left off
Public Enum OBJECT_MODE
    IOBJ_NONE = INT_MAXMODE ' Nothing doing.
    IOBJ_BUILD              ' The build screen
    IOBJ_BUILDSEL           ' Selecting a position for a new structure
    IOBJ_DEMOLISHSEL        ' Selecting a structure to demolish
    IOBJ_MANUFACTURE        ' The manufacture screen
    IOBJ_RESEARCH           ' The research screen
    IOBJ_COMMAND            ' the command droid screen

    IOBJ_DEBUG_STRUCTURE
    IOBJ_DEBUG_DROID

    IOBJ_MAX                ' maximum object mode,
End Enum

' Return codes for the widget interface
Public Enum INT_RETVAL
    INT_NONE            ' no key clicks have been intercepted
    INT_INTERCEPT       ' key clicks have been intercepted
    INT_QUIT            ' The game should quit
End Enum

' Empty edit window
Private secondaryWindowUp As Boolean    ' = false
' Chat dialog
Private ChatDialogUp As Boolean         ' = false

Private Const NUMRETBUTS As Long = 7    ' Number of reticule buttons.

Private Type BUTSTATE
    id As Long                          ' UDWORD
    Enabled As Boolean
    Hidden As Boolean
End Type

Private Type BUTOFFSET
    x As Long                           ' SWORD
    y As Long
End Type

Private ReticuleOffsets(0 To NUMRETBUTS - 1) As BUTOFFSET   ' Reticule button form relative positions.

Private ReticuleEnabled(0 To NUMRETBUTS - 1) As BUTSTATE    ' Reticule button enable states.

Private keyButtonMapping As Long        ' UDWORD = 0
Private ReticuleUp As Boolean           ' = false
Private Refreshing As Boolean           ' = false

'**************************************************************************************
'                  Widget ID numbers

Private Const IDPOW_FORM As Long = 100          ' power bar form

' Option screen IDs
Private Const IDOPT_FORM As Long = 1000         ' The option form
Private Const IDOPT_CLOSE As Long = 1006        ' The close button
Private Const IDOPT_LABEL As Long = 1007        ' The Option screen label
Private Const IDOPT_PLAYERFORM As Long = 1008   ' The player button form
Private Const IDOPT_PLAYERLABEL As Long = 1009  ' The player form label
Private Const IDOPT_PLAYERSTART As Long = 1010  ' The first player button
Private Const IDOPT_PLAYEREND As Long = 1030    ' The last possible player button
Private Const IDOPT_QUIT As Long = 1031         ' Quit the game
Private Const IDOPT_DROID As Long = 1037        ' The place droid button
Private Const IDOPT_STRUCT As Long = 1038       ' The place struct button
Private Const IDOPT_FEATURE As Long = 1039      ' The place feature button
Private Const IDOPT_IVISFORM As Long = 1043     ' iViS engine form
Private Const IDOPT_IVISLABEL As Long = 1044    ' iViS form label

Private Const IDPROX_START As Long = 120000     ' The first proximity button
Private Const IDPROX_END As Long = 129999       ' The last proximity button

Private Const PROX_BUTWIDTH As Long = 9
Private Const PROX_BUTHEIGHT As Long = 9
'**************************************************************************************
'                  Widget Positions

' Reticule positions
Private Const RET_BUTWIDTH As Long = 25
Private Const RET_BUTHEIGHT As Long = 28

Private Const WIDG_SPEC_SCREEN_ORDER_START As Long = 65530
Private Const WIDG_SPEC_SCREEN_ORDER_END As Long = 65532

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Public localTemplates() As Long         ' Unsynchronised list, for UI use only.
Public localTemplates_size As Long

' The widget screen
Public psWScreen As W_SCREEN            ' std::shared_ptr<W_SCREEN> psWScreen = nullptr
Public psUIHelpOverlayScreen As W_SCREEN ' std::shared_ptr<W_HELP_OVERLAY_SCREEN> = nullptr

Public intMode As INTMODE_

' Status of the positioning for the object placement
Private Enum edit_pos_mode_
    IED_NOPOS
    IED_POS
End Enum
Private editPosMode As edit_pos_mode_

Public objMode As OBJECT_MODE
Public interfaceController As BaseObjectsController ' std::shared_ptr<BaseObjectsController> = nullptr
Public groupsUI As WIDGET               ' std::shared_ptr<GroupsForum> = nullptr

' The current stats list being used by the stats screen
Private ppsStatsList() As Long          ' BASE_STATS **: (VB) a copy of the list it points to
Private ppsStatsList_size As Long

' The selected builder on the object screen when the build screen is displayed
Private psSelectedBuilder As Long       ' DROID *

' The button ID of an objects stat on the stat screen if it is locked down
Private statID As Long                  ' UDWORD

' The stats for the current getStructPos
Private psPositionStats As Long         ' BASE_STATS *

' Store a list of stats pointers from the main structure stats
Private apsStructStatsList() As Long    ' STRUCTURE_STATS **: their indices

' Store a list of Template pointers for Droids that can be built
Public apsTemplateList() As Long        ' std::vector<DROID_TEMPLATE *>: template pool handles
Public apsTemplateList_size As Long

' Store a list of Feature pointers for features to be placed on the map
Private apsFeatureStatsList() As Long   ' FEATURE_STATS **: their indices

' Store a list of component stats pointers for the design screen
Public numComponent As Long             ' UDWORD
Public apsComponentList() As Long       ' COMPONENT_STATS **: BASE_STATS * handles
Public numExtraSys As Long              ' UDWORD
Public apsExtraSysList() As Long        ' COMPONENT_STATS **: BASE_STATS * handles

' Flags to check whether the power bars are currently on the screen
Private powerBarUp As Boolean           ' = false

' Update functions to be called at an interval
Private Type IntUpdateFunc
    func As WzFunc                      ' std::function<void ()>
    intervalTicks As Long               ' uint32_t = GAME_TICKS_PER_SEC
    realTimeLastCalled As Long          ' uint32_t = 0
End Type
Private intUpdateFuncs() As IntUpdateFunc
Private intUpdateFuncs_size As Long

'**************************GAME CODE ***************************

Private Type RETBUTSTATS_
    downTime As Long                    ' = 0
    filename As String                  ' WzString
    filenameDown As String              ' WzString
    tip As String
    flashing As Boolean                 ' = false
    flashTime As Long                   ' = 0
    button As WIDGET                    ' std::shared_ptr<W_BUTTON> = nullptr
    callbackFunc As Object              ' playerCallbackFunc = nullptr
End Type
Private retbutstats(0 To NUMRETBUTS - 1) As RETBUTSTATS_

' (static std::shared_ptr<W_SCREEN> replayOverlayScreen: hci_replay_c's)

Private IntRefreshPending As Boolean    ' = false
Private IntGroupsRefreshPending As Boolean ' = false

' the lambdas (fork G1's ids)
Public Const FN_reticuleButtonHitTest As Long = 7000   ' setReticuleButtonDimensions's custom hit test (widget, x, y)
Public Const FN_spectatorOverlaysUpdate As Long = 7001 ' intInitialise's update function ()
Public Const FN_spectatorOverlayCheck As Long = 7002   ' its widgForEachOverlayScreen function (screen, order)
Public Const FN_helpOverlayClosed As Long = 7003       ' intShowWidgetHelp's close handler (screen)
Public Const FN_inGameOptionsButtonCalcLayout As Long = 7004 ' intAddInGameOptionsButton's calcLayout (widget)
Public Const FN_inGameOptionsButtonOnClick As Long = 7005    ' its onClick handler (button)
Public Const FN_inGameOptionsButtonTask As Long = 7006       ' the task the handler schedules ()
Public Const FN_reticuleFormCalcLayout As Long = 7007        ' intAddReticule's calcLayout (widget)
Public Const FN_intDisplayReticuleButton As Long = 7008      ' intDisplayReticuleButton (widget, xOffset, yOffset)
Public Const FN_powerBarCalcLayout As Long = 7009            ' intAddPower's calcLayout (widget)
Public Const FN_debugStatFormCalcLayout As Long = 7010       ' intAddDebugStatsForm's calcLayout (widget)
Public Const FN_chatScreenClosed As Long = 7011              ' chatDialog's on-close function ()

' (VB) the tables' initialisers
Public Sub hci_c_init()
    Dim i As Long
    ReticuleOffsets(RETBUT_CANCEL).x = 43: ReticuleOffsets(RETBUT_CANCEL).y = 47
    ReticuleOffsets(RETBUT_FACTORY).x = 47: ReticuleOffsets(RETBUT_FACTORY).y = 15
    ReticuleOffsets(RETBUT_RESEARCH).x = 81: ReticuleOffsets(RETBUT_RESEARCH).y = 33
    ReticuleOffsets(RETBUT_BUILD).x = 81: ReticuleOffsets(RETBUT_BUILD).y = 69
    ReticuleOffsets(RETBUT_DESIGN).x = 47: ReticuleOffsets(RETBUT_DESIGN).y = 87
    ReticuleOffsets(RETBUT_INTELMAP).x = 13: ReticuleOffsets(RETBUT_INTELMAP).y = 69
    ReticuleOffsets(RETBUT_COMMAND).x = 13: ReticuleOffsets(RETBUT_COMMAND).y = 33

    ReticuleEnabled(RETBUT_CANCEL).id = IDRET_CANCEL
    ReticuleEnabled(RETBUT_FACTORY).id = IDRET_MANUFACTURE
    ReticuleEnabled(RETBUT_RESEARCH).id = IDRET_RESEARCH
    ReticuleEnabled(RETBUT_BUILD).id = IDRET_BUILD
    ReticuleEnabled(RETBUT_DESIGN).id = IDRET_DESIGN
    ReticuleEnabled(RETBUT_INTELMAP).id = IDRET_INTEL_MAP
    ReticuleEnabled(RETBUT_COMMAND).id = IDRET_COMMAND
    For i = 0 To NUMRETBUTS - 1
        ReticuleEnabled(i).Enabled = False
        ReticuleEnabled(i).Hidden = False
    Next
End Sub

' localTemplates.push_back(design): a copy
Public Sub localTemplates_push_back(design As DROID_TEMPLATE)
    Dim h As Long, t As DROID_TEMPLATE
    t = design                              ' (VB) before the pool may grow: design may be in it
    h = newTemplate()
    apsTemplates(h) = t
    ReDim Preserve localTemplates(0 To localTemplates_size)
    localTemplates(localTemplates_size) = h
    localTemplates_size = localTemplates_size + 1
End Sub

' localTemplates.clear()
Public Sub localTemplates_clear()
    Dim k As Long
    For k = 0 To localTemplates_size - 1
        deleteTemplate localTemplates(k)
    Next
    Erase localTemplates
    localTemplates_size = 0
End Sub

' apsTemplateList.clear()
Public Sub apsTemplateList_clear()
    Erase apsTemplateList
    apsTemplateList_size = 0
End Sub

' apsTemplateList.push_back(psTemplate): a template pool handle
Public Sub apsTemplateList_push_back(ByVal psTemplate As Long)
    ReDim Preserve apsTemplateList(0 To apsTemplateList_size)
    apsTemplateList(apsTemplateList_size) = psTemplate
    apsTemplateList_size = apsTemplateList_size + 1
End Sub

' MARK: - hci.h's macros

' #define E_W (pie_GetVideoBufferWidth() - BASE_COORDS_X)
Public Function E_W() As Long
    E_W = pie_GetVideoBufferWidth() - BASE_COORDS_X
End Function

' #define E_H (pie_GetVideoBufferHeight() - BASE_COORDS_Y)
Public Function E_H() As Long
    E_H = pie_GetVideoBufferHeight() - BASE_COORDS_Y
End Function

' #define D_W ((pie_GetVideoBufferWidth() - BASE_COORDS_X)/2)
Public Function D_W() As Long
    D_W = (pie_GetVideoBufferWidth() - BASE_COORDS_X) \ 2
End Function

' #define D_H ((pie_GetVideoBufferHeight() - BASE_COORDS_Y)/2)
Public Function D_H() As Long
    D_H = (pie_GetVideoBufferHeight() - BASE_COORDS_Y) \ 2
End Function

' #define RET_Y (324+E_H)
Public Function RET_Y() As Long
    RET_Y = 324 + E_H()
End Function

' #define CHAT_CONSOLEBOXX RET_X + RET_FORMWIDTH + D_W
Public Function CHAT_CONSOLEBOXX() As Long
    CHAT_CONSOLEBOXX = RET_X + RET_FORMWIDTH + D_W()
End Function

' #define OBJ_BACKX (RET_X + RET_FORMWIDTH + BASE_GAP + D_W)	// X coord of object screen back form.
Public Function OBJ_BACKX() As Long
    OBJ_BACKX = RET_X + RET_FORMWIDTH + BASE_GAP + D_W()
End Function

' #define OBJ_BACKY RET_Y	// Y coord of object screen back form.
Public Function OBJ_BACKY() As Long
    OBJ_BACKY = RET_Y()
End Function

' #define GROUP_BACKX OBJ_BACKX
Public Function GROUP_BACKX() As Long
    GROUP_BACKX = OBJ_BACKX()
End Function

' #define GROUP_BACKY OBJ_BACKY + 40
Public Function GROUP_BACKY() As Long
    GROUP_BACKY = OBJ_BACKY() + 40
End Function

' #define POW_X OBJ_BACKX + ((OBJ_BACKWIDTH - POW_BARWIDTH) / 4)
Public Function POW_X() As Long
    POW_X = OBJ_BACKX() + ((OBJ_BACKWIDTH - POW_BARWIDTH) \ 4)
End Function

' #define POW_Y (OBJ_BACKY + OBJ_BACKHEIGHT + 6)
Public Function POW_Y() As Long
    POW_Y = OBJ_BACKY() + OBJ_BACKHEIGHT + 6
End Function

' #define STAT_Y (E_H)
Public Function STAT_Y() As Long
    STAT_Y = E_H()
End Function

' 2 16 bit values packed into a DWORD.
Public Function PACKDWORD(ByVal a As Long, ByVal b As Long) As Long
    PACKDWORD = SHL(a, 16) Or b
End Function

Public Function UNPACKDWORD_HI(ByVal a As Long) As Long
    UNPACKDWORD_HI = SHRU(a, 16)
End Function

Public Function UNPACKDWORD_LOW(ByVal a As Long) As Long
    UNPACKDWORD_LOW = a And &HFFFF&
End Function

' 3 10 bit values packed into a DWORD.
Public Function PACKDWORD_TRI(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    PACKDWORD_TRI = SHL(a And &H3FF&, 20) Or SHL(b And &H3FF&, 10) Or (c And &H3FF&)
End Function

Public Function UNPACKDWORD_TRI_A(ByVal a As Long) As Long
    UNPACKDWORD_TRI_A = SHRU(a, 20) And &H3FF&
End Function

Public Function UNPACKDWORD_TRI_B(ByVal a As Long) As Long
    UNPACKDWORD_TRI_B = SHRU(a, 10) And &H3FF&
End Function

Public Function UNPACKDWORD_TRI_C(ByVal a As Long) As Long
    UNPACKDWORD_TRI_C = a And &H3FF&
End Function

' MARK: - the reticule

Private Function buttonIsClickable(ByVal id As Long) As Boolean
    Dim enabled As Boolean
    enabled = intCheckReticuleButEnabled(id)
    buttonIsClickable = Not gamePaused() And enabled
End Function

' p: a W_BUTTON
Private Function buttonIsHighlighted(ByVal p As WIDGET) As Boolean
    Dim clickable As Boolean
    clickable = buttonIsClickable(p.id)
    buttonIsHighlighted = (p.getState() And WBUT_HIGHLIGHT) <> 0 And clickable
End Function

Public Sub setReticuleFlash(ByVal ButId As Long, ByVal flash As Boolean)
    If MissionResUp Then
        Exit Sub
    End If
    If flash <> retbutstats(ButId).flashing Then
        retbutstats(ButId).flashing = flash
        retbutstats(ButId).flashTime = 0
    End If
End Sub

' set up the button's size & hit-testing based on the dimensions of the "normal" image
' button: a W_BUTTON
Public Sub setReticuleButtonDimensions(ByVal button As WIDGET, ByVal filename As String)
    Dim image As Long, hasImage As Boolean, imageWidth As Long, imageHeight As Long
    image = 0                               ' AtlasImageDef *image = nullptr
    If Len(filename) <> 0 Then
        image = iV_GetImage(filename)
        If image <> 0 Then
            hasImage = True
            imageWidth = AtlasImageDef_Width(image)
            imageHeight = AtlasImageDef_Height(image)
        End If
    Else
        If Not (IMAGEFILE_numImages(IntImages) >= IMAGE_RETICULE_GREY) Then ASSERT_FAILURE "IntImages->imageDefs.size() >= IMAGE_RETICULE_GREY", "src/hci.cpp:357", "setReticuleButtonDimensions", "IMAGE_RETICULE_GREY isn't in IntImages?"
        If IMAGEFILE_numImages(IntImages) >= IMAGE_RETICULE_GREY Then
            hasImage = True                 ' &(IntImages->imageDefs[IMAGE_RETICULE_GREY])
            imageWidth = iV_GetImageWidth(IntImages, IMAGE_RETICULE_GREY)
            imageHeight = iV_GetImageHeight(IntImages, IMAGE_RETICULE_GREY)
        End If
    End If

    If hasImage Then
        ' set the button width/height based on the "normal" image dimensions (preserving the current x, y)
        button.setGeometry button.x(), button.y(), imageWidth \ 2, imageHeight \ 2

        ' add a custom hit-testing function that uses a tighter bounding ellipse
        button.setCustomHitTest WzFunc_new(FN_reticuleButtonHitTest)
    End If
End Sub

' the custom hit test's lambda
Private Function reticuleButtonHitTest(ByVal psWidget As WIDGET, ByVal x As Long, ByVal y As Long) As Boolean
    Dim centerX As Single, centerY As Single, axisX As Single, axisY As Single, partX As Single, partY As Single

    ' determine center of ellipse contained within the bounding rect
    centerX = CSng((psWidget.x()) + (psWidget.x() + psWidget.width())) / 2!
    centerY = CSng((psWidget.y()) + (psWidget.y() + psWidget.height())) / 2!

    ' determine semi-major axis + semi-minor axis
    axisX = CSng(psWidget.width()) / 2!
    axisY = CSng(psWidget.height()) / 2!

    ' Srivatsan (https://math.stackexchange.com/users/13425/srivatsan), Check if a point is within an ellipse, URL (version: 2011-10-27): https://math.stackexchange.com/q/76463
    partX = ((CSng(x) - centerX) * (CSng(x) - centerX)) / (axisX * axisX) ' ((x - centerX)^2) / ((axisX)^2)
    partY = ((CSng(y) - centerY) * (CSng(y) - centerY)) / (axisY * axisY) ' ((y - centerY)^2) / ((axisY)^2)
    reticuleButtonHitTest = partX + partY <= 1!
End Function

' optional<std::string>: returns has_value, the value in filename
Public Function getReticuleButtonDisplayFilename(ByVal ButId As Long, filename As String) As Boolean
    If Not ((ButId >= 0) And (ButId < NUMRETBUTS)) Then ASSERT_FAILURE "(ButId >= 0) && (ButId < NUMRETBUTS)", "src/hci.cpp:390", "getReticuleButtonDisplayFilename", "Invalid ButId: %d", ButId: getReticuleButtonDisplayFilename = False: Exit Function
    filename = retbutstats(ButId).filename
    getReticuleButtonDisplayFilename = True
End Function

' callbackFunc: playerCallbackFunc (Nothing = nullptr)
Public Sub setReticuleStats(ByVal ButId As Long, Optional ByVal tip As String = "", Optional ByVal filename As String = "", Optional ByVal filenameDown As String = "", Optional ByVal callbackFunc As Object = Nothing)
    If Not ((ButId >= 0) And (ButId < NUMRETBUTS)) Then ASSERT_FAILURE "(ButId >= 0) && (ButId < NUMRETBUTS)", "src/hci.cpp:396", "setReticuleStats", "Invalid ButId: %d", ButId: Exit Sub

    retbutstats(ButId).tip = tip
    retbutstats(ButId).filename = filename
    retbutstats(ButId).filenameDown = filenameDown
    retbutstats(ButId).downTime = 0
    retbutstats(ButId).flashing = False
    retbutstats(ButId).flashTime = 0
    Set retbutstats(ButId).callbackFunc = callbackFunc
    ReticuleEnabled(ButId).Enabled = False

    If retbutstats(ButId).button Is Nothing Then ' not quite set up yet
        Exit Sub
    End If

    setReticuleButtonDimensions retbutstats(ButId).button, retbutstats(ButId).filename

    If Len(retbutstats(ButId).tip) <> 0 Then
        retbutstats(ButId).button.setTip retbutstats(ButId).tip
    End If

    If Len(retbutstats(ButId).filename) <> 0 Then
        ReticuleEnabled(ButId).Enabled = True
        retbutstats(ButId).button.setState 0
    End If
End Sub

Public Sub setReticulesEnabled(ByVal enabled As Boolean)
    Dim i As Long, state As Long
    For i = 0 To NUMRETBUTS - 1
        If Len(retbutstats(i).filename) = 0 Then
            GoTo continue_
        End If

        ReticuleEnabled(i).Enabled = enabled

        If retbutstats(i).button Is Nothing Then
            GoTo continue_
        End If
        If enabled Then
            W_BUTTON_unlock retbutstats(i).button
        Else
            state = retbutstats(i).button.getState()
            retbutstats(i).button.setState state Or WBUT_LOCK
        End If
continue_:
    Next
End Sub

Private Sub intDisplayReticuleButton(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, DownTime As Long, flashing As Boolean, flashTime As Long
    Dim psButton As WIDGET, butDisabled As Boolean, Down As Boolean, highlighted As Boolean, u As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    u = psWidget.UserData
    DownTime = retbutstats(u).downTime
    flashing = retbutstats(u).flashing
    flashTime = retbutstats(u).flashTime
    If Not (psWidget.typ = WIDG_BUTTON) Then ASSERT_FAILURE "psWidget->type == WIDG_BUTTON", "src/hci.cpp:458", "intDisplayReticuleButton", "Not a button": Exit Sub
    Set psButton = psWidget
    butDisabled = (W_BUTTON_(psButton).state And WBUT_DISABLE) <> 0

    If Len(retbutstats(u).filename) = 0 And Not butDisabled Then
        butDisabled = True
        retbutstats(u).button.setState WBUT_DISABLE
    End If

    If butDisabled Then
        If u <> RETBUT_CANCEL Then
            iV_DrawImage2 "image_reticule_grey.png", x, y, psWidget.width(), psWidget.height()
        Else
            iV_DrawImage2 retbutstats(u).filenameDown, x, y, psWidget.width(), psWidget.height()
        End If
        Exit Sub
    End If

    Down = (W_BUTTON_(psButton).state And (WBUT_DOWN Or WBUT_CLICKLOCK)) <> 0
    If Down And buttonIsClickable(psButton.id) Then
        If (DownTime < 1) And (u <> RETBUT_CANCEL) Then
            iV_DrawImage2 "image_reticule_butdown.png", x, y, psWidget.width(), psWidget.height()
        Else
            iV_DrawImage2 retbutstats(u).filenameDown, x, y, psWidget.width(), psWidget.height()
        End If
        DownTime = DownTime + 1
        flashing = False    ' stop the reticule from flashing if it was
    Else
        If flashing And Not gamePaused() Then
            If ((realTime \ 250) Mod 2) <> 0 Then
                iV_DrawImage2 retbutstats(u).filenameDown, x, y, psWidget.width(), psWidget.height()
                flashTime = 0
            Else
                iV_DrawImage2 retbutstats(u).filename, x, y, psWidget.width(), psWidget.height()
            End If
            flashTime = flashTime + 1
        Else
            iV_DrawImage2 retbutstats(u).filename, x, y, psWidget.width(), psWidget.height()
            DownTime = 0
        End If
    End If

    highlighted = buttonIsHighlighted(psButton)
    If highlighted Then
        If u = RETBUT_CANCEL Then
            iV_DrawImage2 "image_cancel_hilight.png", x - 1, y - 1, psWidget.width() + 2, psWidget.height() + 2
        Else
            iV_DrawImage2 "image_reticule_hilight.png", x - 1, y - 1, psWidget.width() + 2, psWidget.height() + 2
        End If
    End If
    retbutstats(u).flashTime = flashTime
    retbutstats(u).flashing = flashing
    retbutstats(u).downTime = DownTime
End Sub

' MARK: - the life cycle

' Initialise the in game interface
Public Function intInitialise() As Boolean
    Dim i As Long, comp As Long, p As Long
    For i = 0 To NUMRETBUTS - 1
        ReticuleEnabled(i).Hidden = False
    Next

    widgSetTipColour psPalette(WZCOL_TOOLTIP_TEXT)

    WidgSetAudio WzFunc_new(FN_WidgetAudioCallback), ID_SOUND_BUTTON_CLICK_3, ID_SOUND_SELECT, ID_SOUND_BUILD_FAIL

    ' Create storage for Structures that can be built
    ReDim apsStructStatsList(0 To MAXSTRUCTURES - 1)

    ' Create storage for Templates that can be built
    apsTemplateList_clear

    ' Create storage for the feature list
    ReDim apsFeatureStatsList(0 To MAXFEATURES - 1)

    ' Create storage for the component list
    ReDim apsComponentList(0 To MAXCOMPONENT - 1)

    ' Create storage for the extra systems list
    ReDim apsExtraSysList(0 To MAXEXTRASYS - 1)

    psSelectedBuilder = 0

    If Not intInitialiseGraphics() Then
        debug_ 935, LOG_ERROR, "intInitialise", "Failed to initialize interface graphics"
        intInitialise = False
        Exit Function
    End If

    Set psWScreen = W_SCREEN_make()

    If GetGameMode() = GS_NORMAL Then
        If Not intAddPower() Then
            debug_ 945, LOG_ERROR, "intInitialise", "Couldn't create power Bar widget(Out of memory ?)"
            intInitialise = False
            Exit Function
        End If
    End If

    ' Note the current screen state
    intMode = INT_NORMAL

    If Not bInTutorial Then
        For comp = 0 To numStructureStats - 1
            If asStructureStats(comp).typ = REF_DEMOLISH Then
                For p = 0 To MAX_PLAYERS - 1
                    apStructTypeLists(p, comp) = AVAILABLE
                Next
            End If
        Next
    End If

    BuildController_resetShowFavorites

    If NETisReplay() And GetGameMode() = GS_NORMAL Then
        createReplayControllerOverlay
    End If

    ReDim Preserve intUpdateFuncs(0 To intUpdateFuncs_size)
    Set intUpdateFuncs(intUpdateFuncs_size).func = WzFunc_new(FN_spectatorOverlaysUpdate)
    intUpdateFuncs(intUpdateFuncs_size).intervalTicks = GAME_TICKS_PER_SEC + GAME_TICKS_PER_SEC
    intUpdateFuncs(intUpdateFuncs_size).realTimeLastCalled = 0
    intUpdateFuncs_size = intUpdateFuncs_size + 1

    ' refresh defaults for quick chat
    quickChatInitInGame

    intInitialise = True
End Function

' intInitialise's update function: widgForEachOverlayScreen's lambda (pScreen, order)
Private Function spectatorOverlayCheck(ByVal pScreen As W_SCREEN, ByVal order As Long) As Boolean
    Dim realSpec As Boolean, selSpec As Boolean
    If Not pScreen Is Nothing And order >= WIDG_SPEC_SCREEN_ORDER_START And order <= WIDG_SPEC_SCREEN_ORDER_END Then
        If realSelectedPlayer <= MAX_CONNECTED_PLAYERS Then realSpec = NetPlay_players_isSpectator(realSelectedPlayer)
        If realSpec Then
            If selectedPlayer <= MAX_CONNECTED_PLAYERS Then selSpec = NetPlay_players_isSpectator(selectedPlayer)
        End If
        If Not (realSpec And selSpec) Then
            widgRemoveOverlayScreen pScreen
            spectatorOverlayCheck = False
            Exit Function
        End If
    End If
    spectatorOverlayCheck = True
End Function

' Shut down the in game interface
Public Sub interfaceShutDown()
    Dim i As Long, blank As RETBUTSTATS_
    intRemoveIntelMapNoAnim ' always call to ensure overlay screen is destroyed

    If Not replayOverlayScreen Is Nothing Then
        widgRemoveOverlayScreen replayOverlayScreen
        Set replayOverlayScreen = Nothing
    End If

    Set groupsUI = Nothing

    Set psWScreen = Nothing

    Erase apsStructStatsList
    apsTemplateList_clear
    Erase apsFeatureStatsList
    Erase apsComponentList
    Erase apsExtraSysList
    psSelectedBuilder = 0

    Set psWScreen = Nothing

    'obviously!
    ReticuleUp = False

    For i = 0 To NUMRETBUTS - 1
        retbutstats(i) = blank
    Next

    shutdownChatScreen
    closeGuideScreen
    ChatDialogUp = False

    bAllowOtherKeyPresses = True
End Sub

' Set widget refresh pending flag.
'
Public Sub intRefreshScreen()
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    IntRefreshPending = True
End Sub

Public Function intIsRefreshing() As Boolean
    intIsRefreshing = Refreshing
End Function

Public Sub intRefreshGroupsUI()
    IntGroupsRefreshPending = True
End Sub

Public Sub intInformInterfaceObjectRemoved(ByVal psObj As Long)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    If psSelectedBuilder = psObj Then
        debug_ 1063, LOG_INFO, "intInformInterfaceObjectRemoved", "Selected builder removed"
        psSelectedBuilder = 0
    End If
    ' intDoScreenRefresh handles refreshing the backing stores for any open interfaceController
End Sub

Public Function intAddRadarWidget() As Boolean
    Dim radarWidget As WIDGET
    Set radarWidget = getRadarWidget()
    If Not (Not radarWidget Is Nothing) Then ASSERT_FAILURE "radarWidget != nullptr", "src/hci.cpp:1072", "intAddRadarWidget", "Failed to get radar widget?": intAddRadarWidget = False: Exit Function
    psWScreen.psForm.attach radarWidget, ChildZPos_Back
    intAddRadarWidget = True
End Function

' see if a delivery point is selected
' (VB) the FLAG_POSITION * (an aFlagPos handle, 0 = nullptr)
Private Function intFindSelectedDelivPoint() As Long
    Dim it As Long, psFlag As Long
    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/hci.cpp:1080", "intFindSelectedDelivPoint", "Not supported selectedPlayer: %u", selectedPlayer: intFindSelectedDelivPoint = 0: Exit Function

    it = apsFlagPosLists(selectedPlayer).head
    Do While it <> 0
        psFlag = lnObj(it)
        If aFlagPos(psFlag).selected And aFlagPos(psFlag).typ = POS_DELIVERY Then
            intFindSelectedDelivPoint = psFlag
            Exit Function
        End If
        it = lnNext(it)
    Loop

    intFindSelectedDelivPoint = 0
End Function

' Refresh widgets once per game cycle if pending flag is set.
'
Public Sub intDoScreenRefresh()
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    Dim objMajor As Long, statMajor As Long, psFlag As Long
    Dim StatsWasUp As Boolean, OrderWasUp As Boolean, psObjForm As WIDGET

    If IntGroupsRefreshPending And getGroupButtonEnabled() Then
        If Not groupsUI Is Nothing Then
            GroupsForum_updateData groupsUI
        End If
        IntGroupsRefreshPending = False
    End If

    If Not IntRefreshPending Then
        Exit Sub
    End If

    Refreshing = True

    If Not interfaceController Is Nothing Then
        interfaceController.refresh
        IntRefreshPending = False
        Refreshing = False
        Exit Sub
    End If

    objMajor = 0: statMajor = 0

    If intMode = INT_OBJECT Or _
       intMode = INT_STAT Or _
       intMode = INT_CMDORDER Or _
       intMode = INT_ORDER Or _
       intMode = INT_TRANSPORTER Then
        Set psObjForm = widgGetFromID(psWScreen, IDOBJ_FORM)
    End If
    If Not psObjForm Is Nothing Then
        If psObjForm.visible() Then
            StatsWasUp = False
            OrderWasUp = False

            ' If the stats form is up then remove it, but remember that it was up.
            If (intMode = INT_STAT) And Not widgGetFromID(psWScreen, IDSTAT_FORM) Is Nothing Then
                StatsWasUp = True
            End If

            ' store the current tab position
            If Not widgGetFromID(psWScreen, IDOBJ_TABFORM) Is Nothing Then
                objMajor = ListTabWidget_currentPage(widgGetFromID(psWScreen, IDOBJ_TABFORM))
            End If
            If StatsWasUp Then
                statMajor = ListTabWidget_currentPage(widgGetFromID(psWScreen, IDSTAT_TABFORM))
            End If
            ' now make sure the stats screen isn't up
            If Not widgGetFromID(psWScreen, IDSTAT_FORM) Is Nothing Then
                intRemoveStatsNoAnim
            End If

            ' see if there was a delivery point being positioned
            psFlag = intFindSelectedDelivPoint()

            ' see if the commander order screen is up
            If (intMode = INT_CMDORDER) And _
               (Not widgGetFromID(psWScreen, IDORDER_FORM) Is Nothing) Then
                OrderWasUp = True
            End If

            ' set the tabs again
            If Not widgGetFromID(psWScreen, IDOBJ_TABFORM) Is Nothing Then
                ListTabWidget_setCurrentPage widgGetFromID(psWScreen, IDOBJ_TABFORM), objMajor
            End If

            If Not widgGetFromID(psWScreen, IDSTAT_TABFORM) Is Nothing Then
                ListTabWidget_setCurrentPage widgGetFromID(psWScreen, IDSTAT_TABFORM), statMajor
            End If

            If psFlag <> 0 Then
                ' need to restart the delivery point position
                startDeliveryPosition psFlag
            End If

            ' make sure the commander order screen is in the right state
            If (intMode = INT_CMDORDER) And _
               Not OrderWasUp And _
               (Not widgGetFromID(psWScreen, IDORDER_FORM) Is Nothing) Then
                intRemoveOrderNoAnim
                If statID <> 0 Then
                    widgSetButtonState psWScreen, statID, 0
                End If
            End If
        End If
    End If

    ' Refresh the transporter interface.
    intRefreshTransporter

    ' Refresh the order interface.
    intRefreshOrder

    Refreshing = False

    IntRefreshPending = False
End Sub


'hides the power bar from the display
Public Sub intHidePowerBar()
    'only hides the power bar if the player has requested no power bar
    If Not powerBarUp Then
        If Not widgGetFromID(psWScreen, IDPOW_POWERBAR_T) Is Nothing Then
            widgHide psWScreen, IDPOW_POWERBAR_T
        End If
    End If
End Sub

' Reset the widget screen to just the reticule
Public Sub intResetScreen(ByVal NoAnim As Boolean, Optional ByVal skipMissionResultScreen As Boolean = False)
    Dim i As Long, startingIntMode As INTMODE_
    If getWidgetsStatus() = False Then
        NoAnim = True
    End If

    For i = 0 To NUMRETBUTS - 1
        If Not retbutstats(i).button Is Nothing And ReticuleUp Then
            W_BUTTON_unlock retbutstats(i).button
        End If
    Next

    startingIntMode = intMode

    Select Case intMode
    Case INT_DESIGN
        gInputManager_contexts_popState
        triggerEvent TRIGGER_DESIGN_QUIT
    Case INT_INTELMAP
        If Not bMultiPlayer Then
            gameTimeStart
        End If
    Case Else
    End Select

    intStopStructPosition
    If NoAnim Then
        intRemoveOrderNoAnim
        intRemoveStatsNoAnim
        intRemoveObjectNoAnim
        intCloseInGameOptionsNoAnim
        intCloseMultiMenuNoAnim
        intRemoveIntelMapNoAnim
        intRemoveTransNoAnim True
    Else
        intRemoveOrder
        intRemoveStats
        intRemoveObject
        intCloseMultiMenu
        intRemoveIntelMap
        intRemoveTrans True
    End If
    If (intMode = INT_MISSIONRES) And Not skipMissionResultScreen Then
        intRemoveMissionResultNoAnim
    End If
    If startingIntMode = INT_DESIGN Then
        intRemoveDesign
    End If
    intHidePowerBar

    If Not interfaceController Is Nothing Then
        interfaceController.prepareToClose
    End If
    Set interfaceController = Nothing
    setSecondaryWindowUp False
    intMode = INT_NORMAL
    'clearSelelection() sets IntRefreshPending = true by calling intRefreshScreen() but if we're doing this then we won't need to refresh - hopefully!
    IntRefreshPending = False
    intShowGroupSelectionMenu
End Sub

' id: OBJECT_TYPE
Public Sub intOpenDebugMenu(ByVal id As Long)
    Dim i As Long, n As Long
    If gamePaused() Then
        ' All menu tabs will not work if they are opened, and the forms will fail to cleanup properly, if the game is paused.
        Exit Sub
    End If

    intResetScreen True
    Select Case id
    Case OBJ_DROID
        apsTemplateList_clear
        For i = 0 To localTemplates_size - 1
            apsTemplateList_push_back localTemplates(i)
        Next
        ' FIXME Ugly cast, and is undefined behaviour (strict-aliasing violation) in C/C++.
        ' (VB) ppsStatsList = (BASE_STATS **)&apsTemplateList[0]: the templates' BASE_STATS * handles
        ppsStatsList_size = apsTemplateList_size
        If ppsStatsList_size > 0 Then ReDim ppsStatsList(0 To ppsStatsList_size - 1) Else Erase ppsStatsList
        For i = 0 To apsTemplateList_size - 1
            ppsStatsList(i) = STAT_TEMPLATE + apsTemplateList(i)
        Next
        objMode = IOBJ_DEBUG_DROID
        intAddDebugStatsForm ppsStatsList, apsTemplateList_size
        intMode = INT_EDITSTAT
        editPosMode = IED_NOPOS
    Case OBJ_STRUCTURE
        n = numStructureStats
        If n > MAXSTRUCTURES Then n = MAXSTRUCTURES
        For i = 0 To n - 1
            apsStructStatsList(i) = i
        Next
        ' (VB) ppsStatsList = (BASE_STATS **)apsStructStatsList: their BASE_STATS * handles
        ppsStatsList_size = n
        If n > 0 Then ReDim ppsStatsList(0 To n - 1) Else Erase ppsStatsList
        For i = 0 To n - 1
            ppsStatsList(i) = structureStatsBase(apsStructStatsList(i))
        Next
        objMode = IOBJ_DEBUG_STRUCTURE
        intAddDebugStatsForm ppsStatsList, n
        intMode = INT_EDITSTAT
        editPosMode = IED_NOPOS
    Case OBJ_FEATURE
        n = asFeatureStats_size
        If n > MAXFEATURES Then n = MAXFEATURES
        For i = 0 To n - 1
            apsFeatureStatsList(i) = i
        Next
        ' (VB) ppsStatsList = (BASE_STATS **)apsFeatureStatsList: their BASE_STATS * handles
        ppsStatsList_size = n
        If n > 0 Then ReDim ppsStatsList(0 To n - 1) Else Erase ppsStatsList
        For i = 0 To n - 1
            ppsStatsList(i) = STAT_FEATURE + apsFeatureStatsList(i)
        Next
        intAddDebugStatsForm ppsStatsList, n
        intMode = INT_EDITSTAT
        editPosMode = IED_NOPOS
    Case Else
        ASSERT_FAILURE "false", "src/hci.cpp:1342", "intOpenDebugMenu", "Unknown debug menu code"
    End Select
End Sub

' Process return codes from the object placement stats screen
Private Sub intProcessEditStats(ByVal id As Long)
    Dim debugMenuDroidDeliveryPoint As FLAG_POSITION
    If id >= IDSTAT_START And id <= IDSTAT_END Then
        ' Clicked on a stat button - need to look for a location for it
        psPositionStats = ppsStatsList(id - IDSTAT_START)
        If BASE_STATS_hasType(psPositionStats, STAT_TEMPLATE) Then
            ' Placing a droid from the debug menu, set up the flag. (This would probably be safe to do, even if we're placing something else.)
            debugMenuDroidDeliveryPoint.factoryType = REPAIR_FLAG
            debugMenuDroidDeliveryPoint.factoryInc = 0
            debugMenuDroidDeliveryPoint.player = selectedPlayer
            debugMenuDroidDeliveryPoint.selected = False
            startDeliveryPosition_flag debugMenuDroidDeliveryPoint
        Else
            intStartStructPosition psPositionStats
        End If
        editPosMode = IED_POS
    ElseIf id = IDSTAT_CLOSE Then
        intRemoveStats
        intStopStructPosition
        intMode = INT_NORMAL
        objMode = IOBJ_NONE
    End If
End Sub

Private Sub reticuleCallback(ByVal retbut As Long)
    If Not retbutstats(retbut).callbackFunc Is Nothing Then
        retbutstats(retbut).callbackFunc.invoke selectedPlayer
    End If
End Sub

' Run the widgets for the in game interface
Public Function intRunWidgets() As INT_RETVAL
    If vbDisplayOff Then intRunWidgets = INT_NONE: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    Dim quitting As Boolean, msg As String, k As Long, retID As Long
    Dim triggers() As WidgetTrigger, nTriggers As Long, retIDs() As Long, nRetIDs As Long
    Dim selectedPlayerIsSpectator As Boolean, pos As Vector2i, pos2 As Vector2i, c As Boolean
    Dim flag As FLAG_POSITION, size As Vector2i, typ As STRUCTURE_TYPE, stats As Long
    Dim lb As LineBuild, i As Long, psBuilding As Long, psTile As Long, psFeature As Long, psStructure As Long
    Dim tmp As STRUCTURE, psDroid As Long, noOrders As INITIAL_DROID_ORDERS, rot As Rotation
    Dim retCode As INT_RETVAL, sz As Long

    quitting = False

    If bLoadSaveUp Then
        c = runLoadSave(True)
        If c Then c = Len(sRequestResult) > 0
    End If
    If c Then
        If bRequestLoad Then
            loopMissionState = LMS_LOADGAME
            saveGameName = sRequestResult
        Else
            If saveGame(sRequestResult, GTYPE_SAVE_START) Then
                msg = "GAME SAVED: "
                msg = msg & savegameWithoutExtension(saveGameName)
                addConsoleMessage msg, LEFT_JUSTIFY, NOTIFY_MESSAGE

                If Not widgGetFromID(psWScreen, IDMISSIONRES_SAVE) Is Nothing Then
                    widgDelete psWScreen, IDMISSIONRES_SAVE
                End If
            Else
                ASSERT_FAILURE "false", "src/hci.cpp:1416", "intRunWidgets", "intRunWidgets: saveGame Failed"
                deleteSaveGame_classic sRequestResult
            End If
        End If
    End If

    If MissionResUp Then
        intRunMissionResult
    End If

    ' Run the current set of widgets
    nRetIDs = 0
    If Not bLoadSaveUp Then
        nTriggers = widgRunScreen(psWScreen, triggers)
        For k = 0 To nTriggers - 1
            ReDim Preserve retIDs(0 To nRetIDs)
            retIDs(nRetIDs) = triggers(k).widget.id
            nRetIDs = nRetIDs + 1
        Next
    End If

    ' Run any periodic update functions
    For k = 0 To intUpdateFuncs_size - 1
        IntUpdateFunc_callIfNeeded intUpdateFuncs(k)
    Next

    ' We may need to trigger widgets with a key press
    If keyButtonMapping <> 0 Then
        ' Set the appropriate id
        ReDim Preserve retIDs(0 To nRetIDs)
        retIDs(nRetIDs) = keyButtonMapping
        nRetIDs = nRetIDs + 1

        ' Clear it so it doesn't trigger next time around
        keyButtonMapping = 0
    End If

    ' Extra code for the power bars to deal with the shadow
    If powerBarUp Then
        intRunPower
    End If

    If OrderUp Then
        intRunOrder
    End If

    ' Extra code for the design screen to deal with the shadow bar graphs
    If intMode = INT_DESIGN Then
        setSecondaryWindowUp True
        intRunDesign
    ElseIf intMode = INT_INTELMAP Then
        intRunIntelMap
    End If

    ' Deal with any clicks.
    For k = 0 To nRetIDs - 1
        retID = retIDs(k)
        If retID >= IDPROX_START And retID <= IDPROX_END Then
            processProximityButtons retID
            intRunWidgets = INT_NONE
            Exit Function
        End If

        ' (bMultiPlayer && NetPlay.players[selectedPlayer].isSpectator) || (selectedPlayer >= MAX_PLAYERS)
        selectedPlayerIsSpectator = False
        If bMultiPlayer Then selectedPlayerIsSpectator = NetPlay_players_isSpectator(selectedPlayer)
        If Not selectedPlayerIsSpectator Then selectedPlayerIsSpectator = (selectedPlayer >= MAX_PLAYERS)

        Select Case retID
        '****************  Reticule buttons  ****************

        Case IDRET_COMMAND
            If Not selectedPlayerIsSpectator Then
                intResetScreen True
                widgSetButtonState psWScreen, IDRET_COMMAND, WBUT_CLICKLOCK
                intAddCommand
                reticuleCallback RETBUT_COMMAND
            End If

        Case IDRET_BUILD
            If Not selectedPlayerIsSpectator Then
                intResetScreen True
                widgSetButtonState psWScreen, IDRET_BUILD, WBUT_CLICKLOCK
                intAddBuild
                reticuleCallback RETBUT_BUILD
            End If

        Case IDRET_MANUFACTURE
            If Not selectedPlayerIsSpectator Then
                intResetScreen True
                widgSetButtonState psWScreen, IDRET_MANUFACTURE, WBUT_CLICKLOCK
                intAddManufacture
                reticuleCallback RETBUT_FACTORY
            End If

        Case IDRET_RESEARCH
            If Not selectedPlayerIsSpectator Then
                intResetScreen True
                widgSetButtonState psWScreen, IDRET_RESEARCH, WBUT_CLICKLOCK
                intAddResearch
                reticuleCallback RETBUT_RESEARCH
            End If

        Case IDRET_INTEL_MAP
            ' check if RMB was clicked
            If widgGetButtonKey_DEPRECATED(psWScreen) = WKEY_SECONDARY Then
                'set the current message to be the last non-proximity message added
                setCurrentMsg
                setMessageImmediate True
            Else
                psCurrentMsg = 0
            End If
            addIntelScreen
            reticuleCallback RETBUT_INTELMAP

        Case IDRET_DESIGN
            If Not selectedPlayerIsSpectator Then
                intResetScreen True
                widgSetButtonState psWScreen, IDRET_DESIGN, WBUT_CLICKLOCK
                'add the power bar - for looks!
                intShowPowerBar
                intAddDesign False
                intMode = INT_DESIGN
                gInputManager_contexts_pushState
                gInputManager_contexts_makeAllInactive
                reticuleCallback RETBUT_DESIGN
                triggerEvent TRIGGER_MENU_DESIGN_UP
            End If

        Case IDRET_CANCEL
            intResetScreen False
            psCurrentMsg = 0
            reticuleCallback RETBUT_CANCEL

        'Transporter button pressed - OFFWORLD Mission Maps ONLY ********
        Case IDTRANTIMER_BUTTON
            addTransporterInterface 0, True

        Case IDTRANS_LAUNCH
            processLaunchTransporter

        ' Catch the quit button here
        Case INTINGAMEOP_POPUP_QUIT, _
             IDMISSIONRES_QUIT, _
             INTINGAMEOP_QUIT, _
             IDOPT_QUIT
            ' mission quit, esc quit confirm, options screen quit
            intCloseInGameOptions False, False
            intResetScreen False
            quitting = True

        ' Default case passes remaining IDs to appropriate function
        Case Else
            Select Case intMode
            Case INT_EDITSTAT
                intProcessEditStats retID
            Case INT_CMDORDER
                If objMode = IOBJ_COMMAND And retID <> IDOBJ_TABFORM Then
                    intProcessOrder retID
                End If
            Case INT_ORDER
                intProcessOrder retID
            Case INT_MISSIONRES
                intProcessMissionResult retID
            Case INT_INGAMEOP
                intProcessInGameOptions retID
            Case INT_MULTIMENU
                ' no-op here
            Case INT_DESIGN
                intProcessDesign retID
            Case INT_INTELMAP
                ' no-op here
            Case INT_TRANSPORTER
                intProcessTransporter retID
            Case INT_STAT, INT_OBJECT, INT_NORMAL
            Case Else
                ASSERT_FAILURE "false", "src/hci.cpp:1632", "intRunWidgets", "intRunWidgets: unknown interface mode"
            End Select
        End Select
    Next

    If Not quitting And nRetIDs = 0 Then
        If (intMode = INT_OBJECT Or intMode = INT_STAT) And objMode = IOBJ_BUILDSEL Then
            ' See if a position for the structure has been found
            If found3DBuildLocTwo(pos, pos2) Then
                ' Send the droid off to build the structure assuming the droid
                ' can get to the location chosen

                ' Set the droid order
                c = intNumSelectedDroids(DROID_CONSTRUCT) = 0
                If c Then c = intNumSelectedDroids(DROID_CYBORG_CONSTRUCT) = 0
                If c Then c = psSelectedBuilder <> 0
                If c Then c = DROID_isConstructionDroid(psSelectedBuilder)
                If c Then
                    orderDroidStatsTwoLocDir psSelectedBuilder, DORDER_LINEBUILD, castStructureStats(psPositionStats), pos.x, pos.y, pos2.x, pos2.y, getBuildingDirection(), ModeQueue
                Else
                    orderSelectedStatsTwoLocDir selectedPlayer, DORDER_LINEBUILD, castStructureStats(psPositionStats), pos.x, pos.y, pos2.x, pos2.y, getBuildingDirection(), ctrlShiftDown()
                End If
                If Not quickQueueMode Then
                    ' Clear the object screen, only if we aren't immediately building something else
                    intResetScreen False
                End If

            ElseIf found3DBuilding(pos) Then 'found building
                'check droid hasn't died
                c = psSelectedBuilder = 0
                If Not c Then c = aDroid(psSelectedBuilder).died = 0
                If c Then
                    ' Send the droid off to build the structure assuming the droid
                    ' can get to the location chosen

                    ' Don't allow derrick to be built on burning ground.
                    If asStructureStats(castStructureStats(psPositionStats)).typ = REF_RESOURCE_EXTRACTOR Then
                        If fireOnLocation(pos.x, pos.y) Then
                            AddDerrickBurningMessage
                        End If
                    End If
                    ' Set the droid order
                    c = intNumSelectedDroids(DROID_CONSTRUCT) = 0
                    If c Then c = intNumSelectedDroids(DROID_CYBORG_CONSTRUCT) = 0
                    If c Then c = psSelectedBuilder <> 0
                    If c Then
                        orderDroidStatsLocDir psSelectedBuilder, DORDER_BUILD, castStructureStats(psPositionStats), pos.x, pos.y, getBuildingDirection(), ModeQueue
                    Else
                        orderSelectedStatsLocDir selectedPlayer, DORDER_BUILD, castStructureStats(psPositionStats), pos.x, pos.y, getBuildingDirection(), ctrlShiftDown()
                    End If
                End If

                If Not quickQueueMode Then
                    ' Clear the object screen, only if we aren't immediately building something else
                    intResetScreen False
                End If
            End If
            If buildState = BUILD3D_NONE Then
                intResetScreen False
            End If
        ElseIf intMode = INT_EDITSTAT And editPosMode = IED_POS Then
            ' Directly positioning some type of object
            pos.x = INT32_MAX: pos.y = INT32_MAX
            pos2.x = INT32_MAX: pos2.y = INT32_MAX
            size.x = 1: size.y = 1
            typ = REF_WALL
            c = sBuildDetails.psStats <> 0
            If c Then
                c = found3DBuildLocTwo(pos, pos2)
                If Not c Then c = found3DBuilding(pos)
            End If
            If c Then
                stats = castStructureStats(sBuildDetails.psStats)
                If stats <> -1 Then
                    size = STRUCTURE_STATS_size(stats, playerPos.r.y)
                    typ = asStructureStats(stats).typ
                End If
                If pos2.x = INT32_MAX Then
                    pos2 = pos
                End If
            ElseIf deliveryReposFinished_flag(flag) Then
                pos.x = flag.coords.x: pos.y = flag.coords.y
                pos2 = pos
            End If

            If pos.x <> INT32_MAX Then
                lb = calcLineBuild_size(size, typ, pos, pos2)
                For i = 0 To lb.count - 1
                    pos = LineBuild_at(lb, i)
                    ' See what type of thing is being put down
                    psBuilding = castStructureStats(psPositionStats)
                    If psBuilding <> -1 And selectedPlayer < MAX_PLAYERS Then

                        If asStructureStats(psBuilding).typ = REF_DEMOLISH Then
                            psTile = mapTile(map_coord(pos.x), map_coord(pos.y))
                            psFeature = psMapTiles(psTile).psObject
                            psStructure = psMapTiles(psTile).psObject

                            c = psStructure <> 0
                            If c Then c = SIMPLE_OBJECT_type(psMapTiles(psTile).psObject) = OBJ_STRUCTURE
                            If c Then
                                'removeStruct(psStructure, true);
                                SendDestroyStructure psStructure
                            Else
                                c = psFeature <> 0
                                If c Then c = SIMPLE_OBJECT_type(psMapTiles(psTile).psObject) = OBJ_FEATURE
                                If c Then
                                    removeFeature psFeature
                                End If
                            End If
                        Else
                            STRUCTURE_construct tmp, generateNewObjectId(), selectedPlayer
                            tmp.state = SAS_NORMAL
                            tmp.pStructureType = castStructureStats(psPositionStats)
                            tmp.pos.x = pos.x: tmp.pos.y = pos.y: tmp.pos.z = map_Height(pos.x, pos.y) + world_coord(1) \ 10

                            ' In multiplayer games be sure to send a message to the
                            ' other players, telling them a new structure has been
                            ' placed.
                            SendBuildFinished_struct tmp
                            ' Send a text message to all players, notifying them of
                            ' the fact that we're cheating ourselves a new
                            ' structure.
                            msg = va("Player %u is cheating (debug menu) him/herself a new structure: %s.", _
                                     selectedPlayer, BASE_STATS_name(structureStatsBase(tmp.pStructureType)))
                            sendInGameSystemMessage msg
                            Cheated = True
                        End If
                    ElseIf BASE_STATS_hasType(psPositionStats, STAT_FEATURE) Then
                        ' Send a text message to all players, notifying them of the fact that we're cheating ourselves a new feature.
                        msg = va("Player %u is cheating (debug menu) him/herself a new feature: %s.", _
                                 selectedPlayer, BASE_STATS_name(psPositionStats))
                        sendInGameSystemMessage msg
                        Cheated = True
                        ' Notify the other hosts that we've just built ourselves a feature
                        'sendMultiPlayerFeature(result->psStats->subType, result->pos.x, result->pos.y, result->id);
                        sendMultiPlayerFeature BASE_STATS_ref(psPositionStats), pos.x, pos.y, generateNewObjectId()
                    ElseIf BASE_STATS_hasType(psPositionStats, STAT_TEMPLATE) Then
                        msg = ""
                        psDroid = buildDroid(castDroidTemplate(psPositionStats), pos.x, pos.y, selectedPlayer, False, False, noOrders, rot)
                        cancelDeliveryRepos
                        If psDroid <> 0 Then
                            addDroid psDroid, apsDroidLists

                            ' Send a text message to all players, notifying them of
                            ' the fact that we're cheating ourselves a new droid.
                            msg = va("Player %u is cheating (debug menu) him/herself a new droid: %s.", selectedPlayer, aDroid(psDroid).aName)

                            triggerEventDroidBuilt psDroid, 0
                        Else
                            ' Send a text message to all players, notifying them of
                            ' the fact that we're cheating ourselves a new droid.
                            msg = va("Player %u is cheating (debug menu) him/herself a new droid.", selectedPlayer)
                        End If
                        sendInGameSystemMessage msg
                        Cheated = True
                    End If
                    If Not quickQueueMode Then
                        editPosMode = IED_NOPOS
                    End If
                Next
            End If
        End If
    End If

    retCode = INT_NONE
    If quitting Then
        retCode = INT_QUIT
    ElseIf nRetIDs <> 0 Or intMode = INT_EDIT Or intMode = INT_MISSIONRES Then
        retCode = INT_INTERCEPT
    ElseIf isMouseOverSomeWidget(psWScreen) Then
        retCode = INT_INTERCEPT
    End If

    c = testPlayerHasLost()
    If Not c Then c = testPlayerHasWon()
    If c Then c = Not bMultiPlayer And intMode <> INT_MISSIONRES
    If c Then c = Not gInputManager_debugManager_debugMappingsAllowed()
    If c Then
        debug_ 1838, LOG_INFO, "intRunWidgets", "PlayerHasLost Or Won"
        intResetScreen True
        retCode = INT_QUIT
    End If
    intRunWidgets = retCode
End Function

' IntUpdateFunc::callIfNeeded()
Private Sub IntUpdateFunc_callIfNeeded(u As IntUpdateFunc)
    If Not ULT(realTime - u.realTimeLastCalled, u.intervalTicks) Then
        If Not u.func Is Nothing Then WzFunc_call u.func
        u.realTimeLastCalled = realTime
    End If
End Sub

' Set the shadow for the PowerBar
Private Sub intRunPower()
    Dim highlightedStatID As Long, psStat As Long, quantity As Long
    quantity = 0

    ' Find out which button was hilited
    highlightedStatID = widgGetMouseOver(psWScreen)
    If highlightedStatID >= IDSTAT_START And highlightedStatID <= IDSTAT_END Then
        ' (VB) the C++ reads past the list (or a null one) when no debug stats form set it: nothing to show
        If highlightedStatID - IDSTAT_START >= ppsStatsList_size Then Exit Sub
        psStat = ppsStatsList(highlightedStatID - IDSTAT_START)
        If BASE_STATS_hasType(psStat, STAT_STRUCTURE) Then
            'get the structure build points
            quantity = asStructureStats(apsStructStatsList(highlightedStatID - _
                        IDSTAT_START)).powerToBuild
        ElseIf BASE_STATS_hasType(psStat, STAT_TEMPLATE) Then
            'get the template build points
            quantity = calcTemplatePower(apsTemplates(apsTemplateList(highlightedStatID - IDSTAT_START)))
        End If

        'update the power bars
        intSetShadowPower quantity
    End If
End Sub

' Set the map view point to the world coordinates x,y
Public Sub intSetMapPos(ByVal x As Long, ByVal y As Long)
    setViewPos map_coord(x), map_coord(y), True
End Sub


' Sync the interface to an object
' If psObj is NULL then reset interface displays.
'
' There should be two version of this function, one for left clicking and one got right.
'
Public Sub intObjectSelected(ByVal psObj As Long)
    Dim structure_ As Long
    If psObj <> 0 Then
        setWidgetsStatus True
        Select Case SIMPLE_OBJECT_type(psObj)
        Case OBJ_DROID
            If Not OrderUp Then
                intResetScreen False
                ' changed to a BASE_OBJECT to accommodate the factories - AB 21/04/99
                intAddOrder psObj
                intMode = INT_ORDER
            Else
                ' changed to a BASE_OBJECT to accommodate the factories - AB 21/04/99
                intAddOrder psObj
            End If

        Case OBJ_STRUCTURE
            If objMode <> IOBJ_DEMOLISHSEL Then
                structure_ = castStructure(psObj)

                If aStruct(structure_).status = SS_BUILT Then
                    If STRUCTURE_isFactory(structure_) Then
                        intAddManufacture
                        Exit Sub
                    ElseIf asStructureStats(aStruct(structure_).pStructureType).typ = REF_RESEARCH Then
                        intAddResearch
                        Exit Sub
                    End If
                End If
            End If
            intResetScreen False
        Case Else
        End Select
    Else
        intResetScreen False
    End If
End Sub


'*
' * Start location selection, where the builder will build the structure
' * structure: a STRUCTURE_STATS * (its index)
Public Sub intStartConstructionPosition(ByVal builder As Long, ByVal structure_ As Long)
    psSelectedBuilder = builder
    psPositionStats = structureStatsBase(structure_)
    intStartStructPosition structureStatsBase(structure_)
End Sub


' Start looking for a structure location
Public Sub intStartStructPosition(ByVal psStats As Long)
    init3DBuilding psStats, 0, 0
End Sub



' Stop looking for a structure location
Private Sub intStopStructPosition()
    psSelectedBuilder = 0

    ' Check there is still a struct position running
    If (intMode = INT_OBJECT Or intMode = INT_STAT) And objMode = IOBJ_BUILDSEL Then
        objMode = IOBJ_NONE
    End If
    kill3DBuilding
End Sub


' Display the widgets for the in game interface
Public Sub intDisplayWidgets()
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    Dim desiredRadarVisibility As Boolean, radarWidget As WIDGET
    'draw the background for the design screen and the Intelligence screen
    If intMode = INT_DESIGN Or intMode = INT_INTELMAP Then
        ' When will they ever learn!!!!
        If Not bMultiPlayer Then
            If Not bInTutorial Then
                screen_RestartBackDrop
            End If

            ' We need to add the console messages to the intelmap for the tutorial so that it can display messages
            If (intMode = INT_DESIGN) Or (bInTutorial And intMode = INT_INTELMAP) Then
                displayConsoleMessages
            End If
        End If
    End If

    desiredRadarVisibility = False
    If Not gameUpdatePaused() Then
        desiredRadarVisibility = radarVisible()

        cleanupOldBeaconMessages

        ' Ensure that any text messages are displayed at bottom of screen
        displayConsoleMessages
    End If

    Set radarWidget = getRadarWidget()
    If Not radarWidget Is Nothing Then
        If desiredRadarVisibility <> radarWidget.visible() Then
            If desiredRadarVisibility Then radarWidget.show Else radarWidget.hide
        End If
    End If

    widgDisplayScreen psWScreen

    If bLoadSaveUp Then
        displayLoadSave
    End If
End Sub

Public Sub intShowWidgetHelp()
    If Not (Not psWScreen Is Nothing) Then ASSERT_FAILURE "psWScreen != nullptr", "src/hci.cpp:2022", "intShowWidgetHelp", "psWScreen is null?": Exit Sub

    If psUIHelpOverlayScreen Is Nothing Then
        ' Initialize the help overlay screen
        Set psUIHelpOverlayScreen = W_HELP_OVERLAY_SCREEN_make( _
            WzFunc_new(FN_helpOverlayClosed) _
        )
        ' close handler
    End If

    If Not bMultiPlayer Or Not NetPlay_bComms Then
        If Not gamePaused() Then
            ' pause game
            setGamePauseStatus True
            setConsolePause True
            setScriptPause True
            setAudioPause True
            setScrollPause True

            ' And stop the clock
            gameTimeStop
        End If
    End If

    W_HELP_OVERLAY_SCREEN_setHelpFromWidgets psUIHelpOverlayScreen, psWScreen.psForm
    widgRegisterOverlayScreenOnTopOfScreen psUIHelpOverlayScreen, psWScreen
End Sub

' intShowWidgetHelp's close handler
Private Sub helpOverlayClosed()
    Set psUIHelpOverlayScreen = Nothing

    If Not bMultiPlayer Or Not NetPlay_bComms Then
        If gamePaused() Then
            ' unpause game

            ' Get it going again
            setGamePauseStatus False
            setConsolePause False
            setScriptPause False
            setAudioPause False
            setScrollPause False

            ' And start the clock again
            gameTimeStart
        End If
    End If
End Sub

Public Function intHelpOverlayIsUp() As Boolean
    If vbDisplayOff Then intHelpOverlayIsUp = False: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    intHelpOverlayIsUp = Not psUIHelpOverlayScreen Is Nothing
End Function

' Tell the interface when the screen has been resized
Public Sub intScreenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If psWScreen Is Nothing Then Exit Sub
    psWScreen.screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
End Sub

' Are we in build select mode
Public Function intBuildSelectMode() As Boolean
    If vbDisplayOff Then intBuildSelectMode = False: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    intBuildSelectMode = (objMode = IOBJ_BUILDSEL)
End Function

' Are we in demolish select mode
Public Function intDemolishSelectMode() As Boolean
    If vbDisplayOff Then intDemolishSelectMode = False: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    intDemolishSelectMode = (objMode = IOBJ_DEMOLISHSEL)
End Function

'Written to allow demolish order to be added to the queuing system
Public Sub intDemolishCancel()
    If objMode = IOBJ_DEMOLISHSEL Then
        objMode = IOBJ_NONE
    End If
End Sub

' Tell the interface a research facility has completed a topic
Public Sub intResearchFinished(ByVal psBuilding As Long)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    If Not (psBuilding <> 0) Then ASSERT_FAILURE "psBuilding != nullptr", "src/hci.cpp:2109", "intResearchFinished", "Invalid structure pointer": Exit Sub

    ' just do a screen refresh
    intRefreshScreen
End Sub

Public Sub intAlliedResearchChanged()
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    If (intMode = INT_OBJECT Or intMode = INT_STAT) And objMode = IOBJ_RESEARCH Then
        intRefreshScreen
    End If
End Sub

' selectedGroup: optional<UBYTE> (-1 = nullopt)
Public Sub intGroupsChanged(Optional ByVal selectedGroup As Long = -1)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    If Not groupsUI Is Nothing Then
        intRefreshGroupsUI
        If selectedGroup <> -1 And selectedGroup <> UBYTE_MAX Then
            GroupsForum_updateSelectedGroup groupsUI, selectedGroup
        End If
    End If
End Sub

' additionalDamage: uint64_t (a Double)
Public Sub intGroupDamaged(ByVal group As Long, ByVal additionalDamage As Double, ByVal unitKilled As Boolean)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    If Not groupsUI Is Nothing Then
        GroupsForum_addGroupDamageForCurrentTick groupsUI, group, additionalDamage, unitKilled
    End If
End Sub

Public Sub intCommanderGroupChanged(ByVal psCommander As Long) ' psCommander may be null!
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    intGroupsChanged ' just trigger full group change event
End Sub

Public Sub intCommanderGroupDamaged(ByVal psCommander As Long, ByVal additionalDamage As Double, ByVal unitKilled As Boolean)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    If Not groupsUI Is Nothing Then
        GroupsForum_addCommanderGroupDamageForCurrentTick groupsUI, psCommander, additionalDamage, unitKilled
    End If
End Sub

Public Function intShowGroupSelectionMenu() As Boolean
    Dim isSpectator As Boolean, groupsUIFormWidg As WIDGET
    If bMultiPlayer And selectedPlayer < NetPlay_players_size() Then isSpectator = NetPlay_players_isSpectator(selectedPlayer)
    If getGroupButtonEnabled() And Not isSpectator Then
        Set groupsUIFormWidg = widgGetFromID(psWScreen, IDOBJ_GROUP)
        If groupsUIFormWidg Is Nothing Then
            If groupsUI Is Nothing Then
                Set groupsUI = GroupsForum_make()
            End If
            psWScreen.psForm.attach groupsUI
        End If

        groupsUI.show
    Else
        widgDelete psWScreen, IDOBJ_GROUP
        Set groupsUI = Nothing
    End If
    intShowGroupSelectionMenu = True
End Function

' MARK: W_INGAMEOPTIONS_BUTTON

Public Sub W_INGAMEOPTIONS_BUTTON_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_W_INGAMEOPTIONS_BUTTON
    W_INGAMEOPTIONS_BUTTON_(self).marginX = 16
    W_INGAMEOPTIONS_BUTTON_(self).marginTop = 18
End Sub

Public Function W_INGAMEOPTIONS_BUTTON_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_W_INGAMEOPTIONS_BUTTON)
    W_INGAMEOPTIONS_BUTTON_construct widget
    W_INGAMEOPTIONS_BUTTON_initialize widget
    Set W_INGAMEOPTIONS_BUTTON_make = widget
End Function

Private Sub W_INGAMEOPTIONS_BUTTON_initialize(ByVal self As WIDGET)
    self.id = IDRET_OPTIONS
    self.setGeometry 0, 0, RET_BUTWIDTH, RET_BUTHEIGHT
End Sub

Public Function W_INGAMEOPTIONS_BUTTON_getMarginX(ByVal self As WIDGET) As Long
    W_INGAMEOPTIONS_BUTTON_getMarginX = W_INGAMEOPTIONS_BUTTON_(self).marginX
End Function

Public Function W_INGAMEOPTIONS_BUTTON_getMarginTop(ByVal self As WIDGET) As Long
    W_INGAMEOPTIONS_BUTTON_getMarginTop = W_INGAMEOPTIONS_BUTTON_(self).marginTop
End Function

' uint8_t
Private Function W_INGAMEOPTIONS_BUTTON_currentAlphaValue(ByVal self As WIDGET) As Long
    Dim alphaPercentage As Long
    alphaPercentage = war_getOptionsButtonVisibility()

    If W_INGAMEOPTIONS_BUTTON_isHighlighted(self) Then
        alphaPercentage = 100
    End If

    W_INGAMEOPTIONS_BUTTON_currentAlphaValue = ((alphaPercentage * 255) \ 100) And &HFF&
End Function

Private Function W_INGAMEOPTIONS_BUTTON_isHighlighted(ByVal self As WIDGET) As Boolean
    W_INGAMEOPTIONS_BUTTON_isHighlighted = ((self.getState() And WBUT_HIGHLIGHT) <> 0) Or InGameOpUp
End Function

Public Sub W_INGAMEOPTIONS_BUTTON_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, butDisabled As Boolean, alphaValue As Long, Down As Boolean
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    butDisabled = (self.getState() And WBUT_DISABLE) <> 0

    alphaValue = W_INGAMEOPTIONS_BUTTON_currentAlphaValue(self)

    If butDisabled Then
        iV_DrawImageTint IntImages, IMAGE_RETICULE_GREY, x0, y0, pal_RGBA(255, 255, 255, alphaValue), True, self.width(), self.height()
        Exit Sub
    End If

    Down = (self.getState() And (WBUT_DOWN Or WBUT_CLICKLOCK)) <> 0
    If Down Then
        iV_DrawImageTint IntImages, IMAGE_INGAMEOPTIONS_DOWN, x0, y0, pal_RGBA(255, 255, 255, alphaValue), True, self.width(), self.height()
    Else
        iV_DrawImageTint IntImages, IMAGE_INGAMEOPTIONS_UP, x0, y0, pal_RGBA(255, 255, 255, alphaValue), True, self.width(), self.height()
    End If

    If W_INGAMEOPTIONS_BUTTON_isHighlighted(self) Then
        iV_DrawImageTint IntImages, IMAGE_RETICULE_HILIGHT, x0, y0, pal_RGBA(255, 255, 255, alphaValue), True, self.width(), self.height()
    End If
End Sub

Public Function W_INGAMEOPTIONS_BUTTON_getTip(ByVal self As WIDGET) As String
    If Not InGameOpUp Then
        W_INGAMEOPTIONS_BUTTON_getTip = "Open In-Game Options"
    Else
        W_INGAMEOPTIONS_BUTTON_getTip = "Close In-Game Options"
    End If
End Function

Public Function intAddInGameOptionsButton() As Boolean
    Dim hiddenOptionsButton As Boolean, psExistingBut As WIDGET, oldButtonHiddenStatus As Boolean, button As WIDGET
    hiddenOptionsButton = war_getOptionsButtonVisibility() = 0

    Set psExistingBut = widgGetFromID(psWScreen, IDRET_OPTIONS)
    If Not psExistingBut Is Nothing Then
        oldButtonHiddenStatus = Not psExistingBut.visible()
        psExistingBut.show Not hiddenOptionsButton

        ' Trigger Re-position of the Replay Controller form
        If Not replayOverlayScreen Is Nothing And (oldButtonHiddenStatus <> hiddenOptionsButton) Then
            replayOverlayScreen.screenSizeDidChange screenWidth, screenHeight, screenWidth, screenHeight
        End If
        intAddInGameOptionsButton = False
        Exit Function
    End If

    Set button = W_INGAMEOPTIONS_BUTTON_make()
    If button Is Nothing Then
        debug_ 2277, LOG_ERROR, "intAddInGameOptionsButton", "Failed to create in game options button"
    Else
        psWScreen.psForm.attach button
        setReticuleButtonDimensions button, "image_ingameoptions_up.png"
        button.setCalcLayout WzFunc_new(FN_inGameOptionsButtonCalcLayout)
        W_BUTTON_addOnClickHandler button, WzFunc_new(FN_inGameOptionsButtonOnClick)

        If hiddenOptionsButton Then
            button.hide
        End If
    End If

    ' Trigger Re-position of the Replay Controller form
    If Not replayOverlayScreen Is Nothing Then
        replayOverlayScreen.screenSizeDidChange screenWidth, screenHeight, screenWidth, screenHeight
    End If

    intAddInGameOptionsButton = True
End Function

' the options button's calcLayout (LAMBDA_CALCLAYOUT_SIMPLE)
Private Sub inGameOptionsButtonCalcLayout(ByVal psWidget As WIDGET)
    Dim psIngameOptButton As WIDGET, w As Long, h As Long, x0 As Long, y0 As Long
    If WIDGET_isKindOf(psWidget, K_W_INGAMEOPTIONS_BUTTON) Then Set psIngameOptButton = psWidget
    If Not (Not psIngameOptButton Is Nothing) Then ASSERT_FAILURE "psIngameOptButton != nullptr", "src/hci.cpp:2285", "intAddInGameOptionsButton", "Wrong button type?": Exit Sub
    w = psWidget.width()
    h = psWidget.height()
    x0 = screenWidth - (w + W_INGAMEOPTIONS_BUTTON_getMarginX(psIngameOptButton))
    y0 = W_INGAMEOPTIONS_BUTTON_getMarginTop(psIngameOptButton)
    psWidget.setGeometry x0, y0, w, h
End Sub

Public Sub intHideInGameOptionsButton()
    Dim psOptionsBut As WIDGET
    Set psOptionsBut = widgGetFromID(psWScreen, IDRET_OPTIONS)
    If Not psOptionsBut Is Nothing Then
        psOptionsBut.hide
    End If
End Sub

' Add the reticule widgets to the widget screen
Public Function intAddReticule() As Boolean
    Dim parent As WIDGET, retForm As WIDGET, i As Long, sButInit As W_BUTINIT, pButton As WIDGET
    intAddInGameOptionsButton

    If ReticuleUp Then
        intAddReticule = True ' all fine
        Exit Function
    End If
    Set parent = psWScreen.psForm
    Set retForm = IntFormAnimated_new(False)
    parent.attach retForm
    retForm.id = IDRET_FORM
    retForm.setCalcLayout WzFunc_new(FN_reticuleFormCalcLayout)
    For i = 0 To NUMRETBUTS - 1
        ' Set the x,y members of a button widget initialiser
        '
        ' Default button data
        W_BUTINIT_construct sButInit
        sButInit.formID = IDRET_FORM
        sButInit.id = ReticuleEnabled(i).id
        sButInit.width = RET_BUTWIDTH
        sButInit.height = RET_BUTHEIGHT
        Set sButInit.pDisplay = WzFunc_new(FN_intDisplayReticuleButton)
        sButInit.x = ReticuleOffsets(i).x
        sButInit.y = ReticuleOffsets(i).y
        sButInit.pTip = retbutstats(i).tip
        sButInit.style = WBUT_SECONDARY
        sButInit.UserData = i
        retbutstats(i).downTime = 0
        retbutstats(i).flashing = False
        retbutstats(i).flashTime = 0
        Set pButton = widgAddButton(psWScreen, sButInit)
        Set retbutstats(i).button = Nothing
        If Not pButton Is Nothing Then
            If WIDGET_isKindOf(pButton, K_W_BUTTON) Then Set retbutstats(i).button = pButton
        End If
        If retbutstats(i).button Is Nothing Then
            debug_ 2361, LOG_ERROR, "intAddReticule", "Failed to add reticule button"
        Else
            setReticuleButtonDimensions retbutstats(i).button, retbutstats(i).filename
        End If
    Next
    ReticuleUp = True
    intAddReticule = True
End Function

Public Sub intRemoveReticule()
    If ReticuleUp = True Then
        widgDelete psWScreen, IDRET_FORM    ' remove reticule
        ReticuleUp = False
    End If
End Sub

'toggles the Power Bar display on and off
Public Sub togglePowerBar()
    'toggle the flag
    powerBarUp = Not powerBarUp

    If powerBarUp Then
        intShowPowerBar
    Else
        intHidePowerBar
    End If
End Sub

' Add the power bars to the screen
Public Function intAddPower() As Boolean
    Dim sBarInit As W_BARINIT, psBarGraph As WIDGET
    W_BARINIT_construct sBarInit

    ' Add the trough bar
    sBarInit.formID = 0     'IDPOW_FORM;
    sBarInit.id = IDPOW_POWERBAR_T
    'start the power bar off in view (default)
    sBarInit.style = WBAR_TROUGH
    sBarInit.sCol = psPalette(WZCOL_POWER_BAR)
    sBarInit.iRange = POWERBAR_SCALE

    Set psBarGraph = PowerBar_new(sBarInit)
    psBarGraph.setCalcLayout WzFunc_new(FN_powerBarCalcLayout)

    If Not (Not psWScreen Is Nothing) Then ASSERT_FAILURE "(psWScreen) != nullptr", "src/hci.cpp:2415", "intAddPower", "%s should not be nullptr", "psWScreen": intAddPower = False: Exit Function
    If Not (Not psWScreen.psForm Is Nothing) Then ASSERT_FAILURE "(psWScreen->psForm) != nullptr", "src/hci.cpp:2416", "intAddPower", "%s should not be nullptr", "psWScreen->psForm": intAddPower = False: Exit Function
    psWScreen.psForm.attach psBarGraph

    powerBarUp = True
    intAddPower = True
End Function

' Remove the build widgets from the widget screen
Public Sub intRemoveObject()
    Dim Form As WIDGET
    widgDelete psWScreen, IDOBJ_TABFORM

    ' Start the window close animation.
    Set Form = widgGetFromID(psWScreen, IDOBJ_FORM)
    If Not Form Is Nothing Then
        IntFormAnimated_closeAnimateDelete Form
    End If

    intHidePowerBar
End Sub


' Remove the build widgets from the widget screen
Public Sub intRemoveObjectNoAnim()
    widgDelete psWScreen, IDOBJ_TABFORM
    widgDelete psWScreen, IDOBJ_FORM

    intHidePowerBar
End Sub


' Remove the stats widgets from the widget screen
Public Sub intRemoveStats()
    Dim Form As WIDGET
    widgDelete psWScreen, IDSTAT_CLOSE
    widgDelete psWScreen, IDSTAT_TABFORM
    setSecondaryWindowUp False

    ' Start the window close animation.
    Set Form = widgGetFromID(psWScreen, IDSTAT_FORM)
    If Not Form Is Nothing Then
        IntFormAnimated_closeAnimateDelete Form
    End If
End Sub


' Remove the stats widgets from the widget screen
Public Sub intRemoveStatsNoAnim()
    widgDelete psWScreen, IDSTAT_CLOSE
    widgDelete psWScreen, IDSTAT_TABFORM
    widgDelete psWScreen, IDSTAT_FORM
    setSecondaryWindowUp False
End Sub

' Makes a button to toggle showing obsolete items.
Public Sub makeObsoleteButton(ByVal parent As WIDGET)
    Dim obsoleteButton As WIDGET
    Set obsoleteButton = MultipleChoiceButton_new()
    parent.attach obsoleteButton
    obsoleteButton.id = IDSTAT_OBSOLETE_BUTTON
    obsoleteButton.style = obsoleteButton.style Or WBUT_SECONDARY
    MultipleChoiceButton_setChoice obsoleteButton, IIf(includeRedundantDesigns, 1, 0)
    MultipleChoiceButton_setImages obsoleteButton, 0, W_BUTTON_Images_(AtlasImage_(IntImages, IMAGE_OBSOLETE_HIDE_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_HIDE_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_HIDE_HI), AtlasImage_())
    MultipleChoiceButton_setTip_choice obsoleteButton, 0, "Hiding Obsolete Tech"
    MultipleChoiceButton_setImages obsoleteButton, 1, W_BUTTON_Images_(AtlasImage_(IntImages, IMAGE_OBSOLETE_SHOW_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_SHOW_UP), AtlasImage_(IntImages, IMAGE_OBSOLETE_SHOW_HI), AtlasImage_())
    MultipleChoiceButton_setTip_choice obsoleteButton, 1, "Showing Obsolete Tech"
    obsoleteButton.move 4 + AtlasImage_width(IntImages, IMAGE_FDP_UP) + 4, STAT_SLDY
End Sub

' Add the stats widgets to the widget screen
' If psSelected != NULL it specifies which stat should be hilited
'   psOwner specifies which object is hilighted on the object bar for this stat
' ppsStatsList_: BASE_STATS ** (handles)
Private Function intAddDebugStatsForm(ppsStatsList_() As Long, ByVal numStats As Long) As Boolean
    Dim parent As WIDGET, statForm As WIDGET, sButInit As W_BUTINIT, statList As WIDGET, statListWidth As Long
    Dim nextButtonId As Long, sBarInit As W_BARINIT, i As Long, button As WIDGET, Stat As Long
    Dim tipString As String, powerCost As Long, bar As WIDGET, costString As String
    ' should this ever be called with psOwner == NULL?

    ' Is the form already up?
    If Not widgGetFromID(psWScreen, IDSTAT_FORM) Is Nothing Then
        intRemoveStatsNoAnim
    End If

    ' is the order form already up ?
    If Not widgGetFromID(psWScreen, IDORDER_FORM) Is Nothing Then
        intRemoveOrderNoAnim
    End If

    setSecondaryWindowUp True

    Set parent = psWScreen.psForm

    ' Create the basic form
    Set statForm = IntFormAnimated_new(False)
    parent.attach statForm
    statForm.id = IDSTAT_FORM
    statForm.setCalcLayout WzFunc_new(FN_debugStatFormCalcLayout)

    ' Add the close button
    W_BUTINIT_construct sButInit
    sButInit.formID = IDSTAT_FORM
    sButInit.id = IDSTAT_CLOSE
    sButInit.x = STAT_WIDTH - CLOSE_WIDTH
    sButInit.y = 0
    sButInit.width = CLOSE_WIDTH
    sButInit.height = CLOSE_HEIGHT
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayImageHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddDebugStatsForm = False
        Exit Function
    End If

    ' Add the tabbed form
    Set statList = IntListTabWidget_make()
    statForm.attach statList
    statList.id = IDSTAT_TABFORM
    ListTabWidget_setChildSize statList, STAT_BUTWIDTH, STAT_BUTHEIGHT
    ListTabWidget_setChildSpacing statList, STAT_GAP, STAT_GAP
    statListWidth = STAT_BUTWIDTH * 2 + STAT_GAP
    statList.setGeometry (STAT_WIDTH - statListWidth) \ 2, STAT_TABFORMY, statListWidth, STAT_HEIGHT - STAT_TABFORMY

    ' Add the stat buttons
    nextButtonId = IDSTAT_START

    W_BARINIT_construct sBarInit
    sBarInit.id = IDSTAT_TIMEBARSTART
    sBarInit.x = STAT_TIMEBARX
    sBarInit.width = STAT_PROGBARWIDTH
    sBarInit.height = STAT_PROGBARHEIGHT
    sBarInit.size = 50
    sBarInit.sCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MAJOR)
    sBarInit.sMinorCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MINOR)

    statID = 0
    For i = 0 To numStats - 1
        sBarInit.y = STAT_TIMEBARY

        If nextButtonId > IDSTAT_END Then
            'can't fit any more on the screen!
            debug_ 2564, LOG_WARNING, "intAddDebugStatsForm", "This is just a Warning! Max buttons have been allocated"
            Exit For
        End If

        Set button = IntStatsButton_new()
        statList.attach button
        button.id = nextButtonId
        button.style = button.style Or WFORM_SECONDARY
        IntStatsButton_setStats button, ppsStatsList_(i)
        ListTabWidget_addWidgetToLayout statList, button

        Stat = ppsStatsList_(i)
        tipString = BASE_STATS_name(ppsStatsList_(i))  ' getLocalizedStatsName
        powerCost = 0
        If BASE_STATS_hasType(Stat, STAT_STRUCTURE) Then ' It's a structure.
            powerCost = asStructureStats(castStructureStats(Stat)).powerToBuild
            sBarInit.size = (powerCost \ POWERPOINTS_DROIDDIV) And &HFFFF&
            If sBarInit.size > 100 Then
                sBarInit.size = 100
            End If

            sBarInit.formID = nextButtonId
            sBarInit.iRange = GAME_TICKS_PER_SEC
            Set bar = widgAddBarGraph(psWScreen, sBarInit)
            W_BARGRAPH_setBackgroundColour bar, psPalette(WZCOL_BLACK)
        ElseIf BASE_STATS_hasType(Stat, STAT_TEMPLATE) Then ' It's a droid.
            powerCost = calcTemplatePower(apsTemplates(castDroidTemplate(Stat)))
            sBarInit.size = (powerCost \ POWERPOINTS_DROIDDIV) And &HFFFF&
            If sBarInit.size > 100 Then
                sBarInit.size = 100
            End If

            sBarInit.formID = nextButtonId
            sBarInit.iRange = GAME_TICKS_PER_SEC
            Set bar = widgAddBarGraph(psWScreen, sBarInit)
            W_BARGRAPH_setBackgroundColour bar, psPalette(WZCOL_BLACK)
        End If
        costString = va("Cost: %u", powerCost)
        tipString = tipString & vbLf
        tipString = tipString & costString
        button.setTip tipString

        ' Update the init struct for the next button
        nextButtonId = nextButtonId + 1

        sBarInit.id = sBarInit.id + 1
    Next

    intAddDebugStatsForm = True
End Function

' Return the stats for a research facility
' (VB) a BASE_STATS * handle (the RESEARCH's, 0 = nullptr)
Private Function getResearchStats(ByVal psObj As Long) As Long
    Dim c As Boolean, psBuilding As Long, psSubjectPending As Long
    c = psObj <> 0
    If c Then c = SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE
    If Not (c) Then ASSERT_FAILURE "psObj != nullptr && psObj->type == OBJ_STRUCTURE", "src/hci.cpp:2624", "getResearchStats", "Invalid Structure pointer": getResearchStats = 0: Exit Function
    psBuilding = psObj
    psSubjectPending = aStruct(psBuilding).functionality.researchFacility.psSubjectPending

    If psSubjectPending <> -1 Then
        If Not IsResearchCompleted(asPlayerResList(aStruct(psBuilding).player, asResearch(psSubjectPending).index)) Then
            getResearchStats = researchStatsBase(psSubjectPending)
            Exit Function
        End If
    End If

    getResearchStats = researchStatsBase(aStruct(psBuilding).functionality.researchFacility.psSubject)
End Function

' (VB) a RESEARCH * (its index, -1 = nullptr) as a BASE_STATS * handle (0 = nullptr), as C++ converts it implicitly
Private Function researchStatsBase(ByVal psResearch As Long) As Long
    If psResearch >= 0 Then
        researchStatsBase = STAT_RESEARCH + psResearch
    Else
        researchStatsBase = 0
    End If
End Function

' controller: std::shared_ptr<BaseObjectsController>
Public Function setController(ByVal controller As BaseObjectsController, ByVal newIntMode As INTMODE_, ByVal newObjMode As OBJECT_MODE) As Boolean
    intResetScreen True

    If Not controller.showInterface() Then
        debug_ 2642, LOG_ERROR, "setController", "Failed to show interface"
        intResetScreen True
        setController = False
        Exit Function
    End If

    intMode = newIntMode
    objMode = newObjMode
    Set interfaceController = controller
    setController = True
End Function

' (VB) bMultiPlayer && NetPlay.players[selectedPlayer].isSpectator
Private Function selectedPlayerIsMultiplayerSpectator() As Boolean
    If bMultiPlayer Then selectedPlayerIsMultiplayerSpectator = NetPlay_players_isSpectator(selectedPlayer)
End Function

' Add the build widgets to the widget screen
Private Function intAddBuild() As Boolean
    If selectedPlayerIsMultiplayerSpectator() Then intAddBuild = False: Exit Function
    intAddBuild = setController(New BuildController, INT_STAT, IOBJ_BUILD)
End Function

' Add the manufacture widgets to the widget screen
Private Function intAddManufacture() As Boolean
    If selectedPlayerIsMultiplayerSpectator() Then intAddManufacture = False: Exit Function
    intAddManufacture = setController(New ManufactureController, INT_STAT, IOBJ_MANUFACTURE)
End Function

' Add the research widgets to the widget screen
Private Function intAddResearch() As Boolean
    If selectedPlayerIsMultiplayerSpectator() Then intAddResearch = False: Exit Function
    intAddResearch = setController(New ResearchController, INT_STAT, IOBJ_RESEARCH)
End Function

' Add the command droid widgets to the widget screen
Private Function intAddCommand() As Boolean
    If selectedPlayerIsMultiplayerSpectator() Then intAddCommand = False: Exit Function
    intAddCommand = setController(New CommanderController, INT_CMDORDER, IOBJ_COMMAND)
End Function

'sets up the Intelligence Screen as far as the interface is concerned
Public Sub addIntelScreen()
    If intMode = INT_INTELMAP Then
        ' screen is already up - do nothing
        Exit Sub
    End If

    intResetScreen False
    intHideGroupSelectionMenu

    intMode = INT_INTELMAP

    'lock the reticule button
    widgSetButtonState psWScreen, IDRET_INTEL_MAP, WBUT_CLICKLOCK

    'add the power bar - for looks!
    intShowPowerBar

    'add all the intelligence screen interface
    intAddIntelMap
End Sub

'sets up the Transporter Screen as far as the interface is concerned
Public Sub addTransporterInterface(ByVal psSelected As Long, ByVal onMission As Boolean)
    Dim c As Boolean
    ' if psSelected = NULL add interface but if psSelected != NULL make sure its not flying
    c = psSelected = 0
    If Not c Then c = Not transporterFlying(psSelected)
    If c Then
        intResetScreen False
        intAddTransporter psSelected, onMission
        intMode = INT_TRANSPORTER
    End If
End Sub

'sets which list of structures to use for the interface
' StructureList *: (VB) copies the list's header into lst (a view of the same nodes); False is nullptr
Public Function interfaceStructList(lst As OBJ_LIST) As Boolean
    If selectedPlayer >= MAX_PLAYERS Then
        interfaceStructList = False
        Exit Function
    End If

    If offWorldKeepLists Then
        lst = mission_apsStructLists(selectedPlayer)
    Else
        lst = apsStructLists(selectedPlayer)
    End If
    interfaceStructList = True
End Function


'causes a reticule button to start flashing
Public Sub flashReticuleButton(ByVal buttonID As Long)
    Dim psButton As WIDGET
    'get the button for the id
    Set psButton = widgGetFromID(psWScreen, buttonID)
    If Not psButton Is Nothing Then
        retbutstats(psButton.UserData).flashing = True
    End If
End Sub

' stop a reticule button flashing
Public Sub stopReticuleButtonFlash(ByVal buttonID As Long)
    Dim psButton As WIDGET
    Set psButton = widgGetFromID(psWScreen, buttonID)
    If Not psButton Is Nothing Then
        retbutstats(psButton.UserData).flashTime = 0
        retbutstats(psButton.UserData).flashing = False
    End If
End Sub

' show selected widget from reticule menu
Public Sub intShowWidget(ByVal buttonID As Long)
    Select Case buttonID
    Case RETBUT_FACTORY
        widgSetButtonState psWScreen, IDRET_MANUFACTURE, WBUT_CLICKLOCK
        intAddManufacture
        reticuleCallback RETBUT_FACTORY
    Case RETBUT_RESEARCH
        widgSetButtonState psWScreen, IDRET_RESEARCH, WBUT_CLICKLOCK
        intAddResearch
        reticuleCallback RETBUT_RESEARCH
    Case RETBUT_BUILD
        widgSetButtonState psWScreen, IDRET_BUILD, WBUT_CLICKLOCK
        intAddBuild
        reticuleCallback RETBUT_BUILD
    Case RETBUT_DESIGN
        intResetScreen True
        widgSetButtonState psWScreen, IDRET_DESIGN, WBUT_CLICKLOCK
        'add the power bar - for looks!
        intShowPowerBar
        intAddDesign False
        intMode = INT_DESIGN
        gInputManager_contexts_pushState
        gInputManager_contexts_makeAllInactive
        reticuleCallback RETBUT_DESIGN
        triggerEvent TRIGGER_MENU_DESIGN_UP
    Case RETBUT_INTELMAP
        addIntelScreen
        reticuleCallback RETBUT_INTELMAP
    Case RETBUT_COMMAND
        widgSetButtonState psWScreen, IDRET_COMMAND, WBUT_CLICKLOCK
        intAddCommand
        reticuleCallback RETBUT_COMMAND
    Case Else
        intResetScreen False
        psCurrentMsg = 0
        reticuleCallback RETBUT_CANCEL
    End Select
End Sub

Public Sub intHideGroupSelectionMenu()
    Dim isSpectator As Boolean
    If Not groupsUI Is Nothing Then
        groupsUI.hide
    End If

    If bMultiPlayer And selectedPlayer < NetPlay_players_size() Then isSpectator = NetPlay_players_isSpectator(selectedPlayer)
    If isSpectator Then
        ' skip showing groups UI if selectedPlayer is spectator
        widgDelete psWScreen, IDOBJ_GROUP
        Set groupsUI = Nothing
    End If
End Sub

'displays the Power Bar
Public Sub intShowPowerBar()
    Dim isSpectator As Boolean
    If bMultiPlayer And selectedPlayer < NetPlay_players_size() Then isSpectator = NetPlay_players_isSpectator(selectedPlayer)
    If isSpectator Then
        ' skip showing power bar if selectedPlayer is spectator
        Exit Sub
    End If

    'if its not already on display
    If Not widgGetFromID(psWScreen, IDPOW_POWERBAR_T) Is Nothing Then
        widgReveal psWScreen, IDPOW_POWERBAR_T
    End If
End Sub

'hides the power bar from the display - regardless of what player requested!
Public Sub forceHidePowerBar(Optional ByVal forceSetPowerBarUpState As Boolean = False)
    If forceSetPowerBarUpState Then
        powerBarUp = False
    End If

    If Not widgGetFromID(psWScreen, IDPOW_POWERBAR_T) Is Nothing Then
        widgHide psWScreen, IDPOW_POWERBAR_T
    End If
End Sub


' Add the Proximity message buttons
' psProxDisp: a PROXIMITY_DISPLAY * (aProxDisp handle)
Public Function intAddProximityButton(ByVal psProxDisp As Long, ByVal inc As Long) As Boolean
    Dim cnt As Long, sBFormInit As W_FORMINIT, it As Long, inUse As Boolean

    If selectedPlayer >= MAX_PLAYERS Then
        intAddProximityButton = False
        Exit Function
    End If

    W_FORMINIT_construct sBFormInit
    sBFormInit.formID = 0
    sBFormInit.id = IDPROX_START + inc

    ' loop back and find a free one!
    If sBFormInit.id >= IDPROX_END Then
        For cnt = IDPROX_START To IDPROX_END - 1
            ' go down the prox msgs and see if it's free.
            inUse = False
            it = apsProxDisp(selectedPlayer).head
            Do While it <> 0
                If aProxDisp(lnObj(it)).buttonID = cnt Then
                    inUse = True
                    Exit Do
                End If
                it = lnNext(it)
            Loop

            If Not inUse Then   ' value was unused.
                sBFormInit.id = cnt
                Exit For
            End If
        Next
        If Not (cnt <> IDPROX_END) Then ASSERT_FAILURE "cnt != IDPROX_END", "src/hci.cpp:2885", "intAddProximityButton", "Ran out of proximity displays": intAddProximityButton = False: Exit Function
    End If
    If Not (sBFormInit.id < IDPROX_END) Then ASSERT_FAILURE "sBFormInit.id < IDPROX_END", "src/hci.cpp:2887", "intAddProximityButton", "Invalid proximity message button ID %d", sBFormInit.id: intAddProximityButton = False: Exit Function
    If vbDisplayOff Then aProxDisp(psProxDisp).buttonID = sBFormInit.id: intAddProximityButton = True: Exit Function  ' (VB) the vbtest runs: statsref's stand-in (the ID, no widget)

    sBFormInit.majorID = 0
    sBFormInit.style = WFORM_CLICKABLE
    sBFormInit.width = PROX_BUTWIDTH
    sBFormInit.height = PROX_BUTHEIGHT
    'the x and y need to be set up each time the button is drawn - see intDisplayProximityBlips

    Set sBFormInit.pDisplay = WzFunc_new(FN_intDisplayProximityBlips)
    'set the data for this button
    sBFormInit.pUserData = psProxDisp

    If widgAddForm(psWScreen, sBFormInit) Is Nothing Then
        intAddProximityButton = False
        Exit Function
    End If

    'store the ID so we can detect which one has been clicked on
    aProxDisp(psProxDisp).buttonID = sBFormInit.id

    intAddProximityButton = True
End Function


'Remove a Proximity Button - when the message is deleted
Public Sub intRemoveProximityButton(ByVal psProxDisp As Long)
    If Not (aProxDisp(psProxDisp).buttonID >= IDPROX_START And aProxDisp(psProxDisp).buttonID <= IDPROX_END) Then ASSERT_FAILURE "psProxDisp->buttonID >= IDPROX_START && psProxDisp->buttonID <= IDPROX_END", "src/hci.cpp:2914", "intRemoveProximityButton", "Invalid proximity ID": Exit Sub
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in (the assert only)
    widgDelete psWScreen, aProxDisp(psProxDisp).buttonID
End Sub

'deals with the proximity message when clicked on
Private Sub processProximityButtons(ByVal id As Long)
    Dim it As Long
    If Not doWeDrawProximitys() Then
        Exit Sub
    End If

    If selectedPlayer >= MAX_PLAYERS Then Exit Sub ' no-op

    'find which proximity display this relates to
    it = apsProxDisp(selectedPlayer).head
    Do While it <> 0
        If aProxDisp(lnObj(it)).buttonID = id Then Exit Do
        it = lnNext(it)
    Loop
    If it <> 0 Then
        'if not been read - display info
        If Not aMessage(aProxDisp(lnObj(it)).psMessage).read Then
            displayProximityMessage lnObj(it)
        End If
    End If
End Sub

'	Fools the widgets by setting a key value
Public Sub setKeyButtonMapping(ByVal val As Long)
    keyButtonMapping = val
End Sub

' count the number of selected droids of a type
Private Function intNumSelectedDroids(ByVal droidType As Long) As Long
    Dim num As Long, it As Long, psDroid As Long

    If selectedPlayer >= MAX_PLAYERS Then
        intNumSelectedDroids = 0
        Exit Function
    End If

    num = 0
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected And aDroid(psDroid).droidType = droidType Then
            num = num + 1
        End If
        it = lnNext(it)
    Loop

    intNumSelectedDroids = num
End Function

'Checks to see if there are any research topics to do and flashes the button -
'only if research facility is free
Public Function intGetResearchState() As Long
    If vbDisplayOff Then intGetResearchState = 0: Exit Function  ' (VB) the vbtest runs: statsref's stand-in
    Dim resFree As Boolean, intStrList As OBJ_LIST, it As Long, psStruct As Long
    Dim count As Long, researchList() As Long, researchList_size As Long, n As Long, player As Long

    If selectedPlayer >= MAX_PLAYERS Then
        intGetResearchState = 0
        Exit Function
    End If

    resFree = False
    If interfaceStructList(intStrList) Then
        it = intStrList.head
        Do While it <> 0
            psStruct = lnObj(it)
            If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESEARCH And _
               aStruct(psStruct).status = SS_BUILT Then
                If getResearchStats(psStruct) = 0 Then
                    resFree = True
                    Exit Do
                End If
            End If
            it = lnNext(it)
        Loop
    End If

    count = 0
    If resFree Then
        'calculate the list
        researchList_size = fillResearchList(selectedPlayer, False, 0, MAXRESEARCH, researchList)
        count = researchList_size
        For n = 0 To researchList_size - 1
            For player = 0 To MAX_PLAYERS - 1
                If aiCheckAlliances(player, selectedPlayer) Then
                    If IsResearchStarted(asPlayerResList(player, researchList(n))) Then
                        count = count - 1  ' An ally is already researching this topic, so don't flash the button because of it.
                        Exit For
                    End If
                End If
            Next
        Next
    End If

    intGetResearchState = count
End Function

Public Sub intNotifyResearchButton(ByVal prevState As Long)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    Dim newState As Long
    newState = intGetResearchState()
    If newState > prevState Then
        ' Set the research reticule button to flash.
        flashReticuleButton IDRET_RESEARCH
    ElseIf newState = 0 And prevState > 0 Then
        stopReticuleButtonFlash IDRET_RESEARCH
    End If
End Sub

' see if a reticule button is enabled
Public Function intCheckReticuleButEnabled(ByVal id As Long) As Boolean
    Dim i As Long
    For i = 0 To NUMRETBUTS - 1
        If ReticuleEnabled(i).id = id Then
            intCheckReticuleButEnabled = ReticuleEnabled(i).Enabled
            Exit Function
        End If
    Next
    intCheckReticuleButEnabled = False
End Function

' Checks if a coordinate is over the build menu
Public Function CoordInBuild(ByVal x As Long, ByVal y As Long) As Boolean
    Const buildmenu_height As Long = 300
    Dim pos As Vector2f
    ' This measurement is valid for the menu, so the buildmenu_height
    ' value is used to "nudge" it all upwards from the command menu.
    ' FIXME: hardcoded value (?)

    pos.x = x - RET_X
    pos.y = y - RET_Y() + buildmenu_height ' guesstimation

    If (pos.x < 0 Or pos.y < 0 Or pos.x >= RET_FORMWIDTH Or pos.y >= buildmenu_height) Or Not isSecondaryWindowUp() Then
        CoordInBuild = False
        Exit Function
    End If

    CoordInBuild = True
End Function

' Our chat dialog for global & team communication
' \mode sets if global or team communication is wanted
Public Sub chatDialog(ByVal mode As Long, ByVal startWithQuickChatFocused As Boolean)
    Dim initialChatMode As WzChatMode
    If Not ChatDialogUp Then
        ChatDialogUp = True

        If mode = CHAT_GLOB Then initialChatMode = WzChatMode_Glob Else initialChatMode = WzChatMode_Team
        createChatScreen WzFunc_new(FN_chatScreenClosed), initialChatMode, startWithQuickChatFocused
    Else
        debug_ 3081, LOG_ERROR, "chatDialog", "Tried to throw up chat dialog when we already had one up!"
    End If
End Sub

' If chat dialog is up
Public Function isChatUp() As Boolean
    isChatUp = ChatDialogUp
End Function

' Helper call to see if we have builder/research/... window up or not.
Public Function isSecondaryWindowUp() As Boolean
    isSecondaryWindowUp = secondaryWindowUp
End Function

Public Sub setSecondaryWindowUp(ByVal value As Boolean)
    secondaryWindowUp = value
End Sub

Public Sub intSetShouldShowRedundantDesign(ByVal value As Boolean)
    includeRedundantDesigns = value
End Sub

Public Function intGetShouldShowRedundantDesign() As Boolean
    intGetShouldShowRedundantDesign = includeRedundantDesigns
End Function

Public Sub intShowInterface()
    intAddReticule
    intShowPowerBar
    intShowGroupSelectionMenu
End Sub

Public Sub intHideInterface(Optional ByVal forceHidePowerbar_ As Boolean = False)
    intRemoveReticule
    If forceHidePowerbar_ Then
        forceHidePowerBar
    Else
        intHidePowerBar
    End If
    intHideGroupSelectionMenu
End Sub

' the lambdas
Public Function hci_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET
    Select Case f.id
    Case FN_reticuleButtonHitTest
        hci_c_fn = reticuleButtonHitTest(args(0), args(1), args(2))
    Case FN_spectatorOverlaysUpdate
        widgForEachOverlayScreen WzFunc_new(FN_spectatorOverlayCheck)
    Case FN_spectatorOverlayCheck
        hci_c_fn = spectatorOverlayCheck(args(0), args(1))
    Case FN_helpOverlayClosed
        helpOverlayClosed
    Case FN_inGameOptionsButtonCalcLayout
        inGameOptionsButtonCalcLayout args(0)
    Case FN_inGameOptionsButtonOnClick
        widgScheduleTask WzFunc_new(FN_inGameOptionsButtonTask)
    Case FN_inGameOptionsButtonTask
        kf_addInGameOptions
    Case FN_reticuleFormCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry RET_X, RET_Y(), RET_FORMWIDTH, RET_FORMHEIGHT
    Case FN_intDisplayReticuleButton
        intDisplayReticuleButton args(0), args(1), args(2)
    Case FN_powerBarCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry POW_X(), POW_Y(), POW_BARWIDTH, iV_GetImageHeight(IntImages, IMAGE_PBAR_EMPTY)
    Case FN_debugStatFormCalcLayout
        Set psWidget = args(0)
        psWidget.setGeometry STAT_X, STAT_Y(), STAT_WIDTH, STAT_HEIGHT
    Case FN_chatScreenClosed
        ChatDialogUp = False
    End Select
End Function
