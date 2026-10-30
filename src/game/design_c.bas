Attribute VB_Name = "design_c"
' src/design.cpp, design.h: the design screen. char aCurrName[MAX_STR_LENGTH]
' is a String cut to MAX_STR_LENGTH - 1 characters (sstrcpy, sstrcat);
' translations are not ported (_() and gettext give the English text;
' getStatsName / getLocalizedStatsName are the stat's name, checkIfZNullStat
' an InStr of "ZNULL" in its id). intValidTemplate and checkTemplateIsVtol also
' check a template received in a message.
' (VB) A COMPONENT_STATS * (BODY_STATS *, WEAPON_STATS *, ...) is a stats handle
' (stats_c: its StatType plus its index, 0 = nullptr); a template's BASE_STATS *
' is STAT_TEMPLATE plus its template_c pool handle. A DROID_TEMPLATE * is a pool
' handle: static DROID_TEMPLATE sDefaultDesignTemplate is a pool slot of its own
' (apsTemplateList holds its address); sCurrDesign and sShadowDesign are module
' UDTs, and where C++ passes &sCurrDesign (or a local copy) as a DROID_TEMPLATE *
' to a function that takes a pool handle, designTemplateHandle copies it into the
' screen's scratch slot. (A pool element is never passed ByRef into code that may
' grow the pool: copyTemplate and localTemplates_push_back get copies.)
' The ComponentIterator (std::function<bool(std::function<bool(COMPONENT_STATS &,
' size_t)>)>), its callbacks, findMax's value functions and the function
' pointers handed to findMax / findMaxWeaponAttribute are WzFuncs (FN_ ids
' 9000-9025); componentIterator's array is its first element's handle (element i
' is psStats + i: no size). The lambdas of setTemplateStat are named procedures.
' The calcLayout lambdas and the display functions are design_c_fn's.
' DESIGN_Y and the positions built on it are functions (D_H is hci's).
' (VB) widget class DesignStatsBar : W_BARGRAPH
' (VB) widget class DesignPowerBar : DesignStatsBar
Option Explicit

' design.h
Public Const IDDES_FORM As Long = 5000          ' The base form for the design screen
Public Const IDDES_STATSFORM As Long = 5001     ' The design screen stats form
Public Const IDDES_SYSTEMFORM As Long = 5006    ' The clickable form for the weapon/ecm/sensor
Public Const IDDES_BODYFORM As Long = 5007      ' The clickable form for the body
Public Const IDDES_PROPFORM As Long = 5008      ' The clickable form for the propulsion
Public Const IDDES_3DVIEW As Long = 5009        ' The 3D view of the droid
Public Const IDDES_BIN As Long = 5011           ' The bin button
Public Const IDDES_NAMEBOX As Long = 5013       ' The Name box
Public Const IDDES_POWERFORM As Long = 5019     ' The form for the power and points bars
Public Const IDDES_TEMPLBASE As Long = 5020     ' The base form for the Template (left) form
Public Const IDDES_RIGHTBASE As Long = 5021     ' The base form for the right form!
Public Const IDDES_POWERBAR As Long = 5023      ' The power bar for the template

Public Const IDDES_WEAPONS As Long = 5024       ' The weapon button for the Component form (right)
Public Const IDDES_SYSTEMS As Long = 5025       ' The systems (sensor/ecm) button for the Component form
Public Const IDDES_COMMAND As Long = 5026       ' The command button for the Component form

Public Const IDDES_PARTFORM As Long = 5027      ' Part buttons form
Public Const IDDES_WEAPONS_A As Long = 5028     ' The weapon TURRET_A button for the Component form (right)
Public Const IDDES_WEAPONS_B As Long = 5029     ' The weapon TURRET_B button for the Component form (right)

Public Const IDDES_STOREBUTTON As Long = 5905   ' Stored template button

Public Const IDDES_BODYARMOUR_K As Long = 5100  ' The body armour bar graph for kinetic weapons
Public Const IDDES_BODYPOWER As Long = 5101     ' The body power plant graph
Public Const IDDES_BODYWEIGHT As Long = 5102    ' The body weight
Public Const IDDES_PROPROAD As Long = 5105      ' The propulsion road speed
Public Const IDDES_PROPCOUNTRY As Long = 5106   ' The propulsion cross country speed
Public Const IDDES_PROPWATER As Long = 5107     ' The propulsion water speed
Public Const IDDES_PROPAIR As Long = 5108       ' The propulsion air speed
Public Const IDDES_PROPWEIGHT As Long = 5109    ' The propulsion weight
Public Const IDDES_SENSORRANGE As Long = 5110   ' The sensor range
Public Const IDDES_SYSTEMSWEIGHT As Long = 5112 ' The systems weight
Public Const IDDES_ECMPOWER As Long = 5115      ' The ecm power
Public Const IDDES_WEAPRANGE As Long = 5120     ' The weapon range
Public Const IDDES_WEAPDAMAGE As Long = 5121    ' The weapon damage
Public Const IDDES_WEAPROF As Long = 5122       ' The weapon rate of fire
Public Const IDDES_CONSTPOINTS As Long = 5125   ' The construction build points
'extras added AB 3/9/97
Public Const IDDES_BODYPOINTS As Long = 5127    ' The body points bar graph
Public Const IDDES_BODYARMOUR_H As Long = 5128  ' The body armour bar graph for heat weapons
Public Const IDDES_REPAIRPOINTS As Long = 5129  ' The Repair points

Public Const IDDES_BODYARMOURKLAB As Long = 5200    ' The body armour (kinetic) bar graph label
Public Const IDDES_BODYPOWERLAB As Long = 5201      ' The body power plant graph label
Public Const IDDES_BODYWEIGHTLAB As Long = 5202     ' The body weight graph label
Public Const IDDES_PROPROADLAB As Long = 5205       ' The propulsion road speed label
Public Const IDDES_PROPCOUNTRYLAB As Long = 5206    ' The propulsion cross country speed label
Public Const IDDES_PROPWATERLAB As Long = 5207      ' The propulsion water speed label
Public Const IDDES_PROPAIRLAB As Long = 5208        ' The propulsion air speed label
Public Const IDDES_PROPWEIGHTLAB As Long = 5209     ' The propulsion weight label
Public Const IDDES_SENSORRANGELAB As Long = 5210    ' The sensor range label
Public Const IDDES_SYSTEMSWEIGHTLAB As Long = 5212  ' The systems weight label
Public Const IDDES_ECMPOWERLAB As Long = 5215       ' The ecm power label
Public Const IDDES_WEAPRANGELAB As Long = 5220      ' The weapon range label
Public Const IDDES_WEAPDAMAGELAB As Long = 5221     ' The weapon damage label
Public Const IDDES_WEAPROFLAB As Long = 5222        ' The weapon rate of fire label
Public Const IDDES_CONSTPOINTSLAB As Long = 5225    ' The construction build points label
'extras added AB 3/9/97
'#define IDDES_BODYPOINTSLAB 5227 // The body points label
Public Const IDDES_BODYARMOURHLAB As Long = 5228    ' The body armour (heat) bar graph label

Public Const IDDES_TEMPPOWERLAB As Long = 5229      ' The template's Power req label
Public Const IDDES_TEMPBODYLAB As Long = 5230       ' The template's Body Points label

Public Const IDDES_REPAIRPTLAB As Long = 5231       ' The Repair Points label
Public Const IDDES_REPAIRWGTLAB As Long = 5232      ' The Repair Weigth label

Public Const IDDES_TEMPLSTART As Long = 53000       ' The first design template button
Public Const IDDES_TEMPLEND As Long = 53499         ' The last design template button
Public Const IDDES_BARSTART As Long = 5400
Public Const IDDES_BAREND As Long = 5499
Public Const IDDES_COMPSTART As Long = 5500000      ' The first component button
Public Const IDDES_COMPEND As Long = 5565535        ' The last component button
Public Const IDDES_EXTRASYSSTART As Long = 5700     ' The first extra system button
Public Const IDDES_EXTRASYSEND As Long = 5899       ' The last extra system button

Public Const IDDES_SYSTEMBUTTON As Long = 5900      ' System button
Public Const IDDES_BODYBUTTON As Long = 5901        ' Body button
Public Const IDDES_PROPBUTTON As Long = 5902        ' Propulsion button
Public Const IDDES_WPABUTTON As Long = 5903         ' WeaponA button
Public Const IDDES_WPBBUTTON As Long = 5904         ' WeaponB button

' (VB) the lambdas, std::functions and function pointers of design.cpp (design_c_fn)
Public Const FN_intAddDesign_desForm_calcLayout As Long = 9000           ' intAddDesign's desForm: LAMBDA_CALCLAYOUT_SIMPLE
Public Const FN_intAddDesign_statsForm_calcLayout As Long = 9001         ' intAddDesign's statsForm
Public Const FN_intAddTemplateForm_templbaseForm_calcLayout As Long = 9002     ' intAddTemplateForm's templbaseForm
Public Const FN_intAddTemplateForm_templList_calcLayout As Long = 9003         ' intAddTemplateForm's templList
Public Const FN_intAddComponentForm_rightBase_calcLayout As Long = 9004         ' intAddComponentForm's rightBase
Public Const FN_intAddComponentForm_compList_calcLayout As Long = 9005          ' intAddComponentForm's compList
Public Const FN_intDisplayStatForm As Long = 9006           ' WIDGET_DISPLAY: args (widget, xOffset, yOffset)
Public Const FN_intDisplayViewForm As Long = 9007
Public Const FN_intDisplayDesignForm As Long = 9008
Public Const FN_componentIterator As Long = 9009            ' [=](callback): cap psStats, aAvailable (player, type), numEntries
Public Const FN_concatIterators As Long = 9010              ' [=](callback): cap the count, then the iterators
Public Const FN_findMax_callback As Long = 9011             ' [&](COMPONENT_STATS &bodyStats, size_t index): cap value, max
Public Const FN_findMaxWeight_value As Long = 9012          ' [](COMPONENT_STATS &stats): args (stats)
Public Const FN_findMaxPropulsionSpeed_value As Long = 9013 ' [=]: cap terrainType
Public Const FN_findMaxWeaponAttribute_value As Long = 9014 ' [=]: cap attributeGetter
Public Const FN_getDesignMaxBodyArmour_value As Long = 9015 ' [=]: cap weaponClass
Public Const FN_getDesignMaxEngineOutput_value As Long = 9016
Public Const FN_calcShadowBodyPoints As Long = 9017         ' the function calcShadowBodyPoints: args (stats)
Public Const FN_calcShadowPower As Long = 9018              ' the function calcShadowPower
Public Const FN_getDesignMaxSensorRange_value As Long = 9019
Public Const FN_getDesignMaxEcmRange_value As Long = 9020
Public Const FN_getDesignMaxBuildPoints_value As Long = 9021
Public Const FN_intSetSystemForm_repairPoints As Long = 9022 ' intSetSystemForm's [=](COMPONENT_STATS &stats)
Public Const FN_calculatePropulsionWeight As Long = 9023    ' the function calculatePropulsionWeight
Public Const FN_intAddComponentButtons_callback As Long = 9024 ' [&](COMPONENT_STATS &currStats, size_t index): cap compList, maxComponents, bWeapon, bVTOL, bodysize, nextButtonId, compID
Public Const FN_intAddExtraSystemButtons_callback As Long = 9025 ' [&](COMPONENT_STATS &stats, size_t i): cap compList, nextButtonId, compIndex
Public Const FN_design_proj_GetLongRange As Long = 9026     ' the function pointer proj_GetLongRange (findMaxWeaponAttribute's attributeGetter): args (stats, player)
Public Const FN_design_weaponDamage As Long = 9027          ' the function pointer weaponDamage
Public Const FN_design_weaponROF As Long = 9028             ' the function pointer weaponROF

Private Const MAX_DESIGN_COMPONENTS As Long = 40        ' Max number of stats the design screen can cope with.
Private Const MAX_SYSTEM_COMPONENTS As Long = 65535

'*************************************************************************************
'                  Max values for the design bar graphs

Private Const DBAR_TEMPLATEMAXPOINTS As Long = 8400     'maximum body points for a template
Private Const DBAR_TEMPLATEMAXPOWER As Long = 1000      'maximum power points for a template

' The maximum number of characters on the component buttons
Private Const DES_COMPBUTMAXCHAR As Long = 5

' Which type of system is displayed on the design screen
Private Enum DES_SYSMODE
    IDES_SENSOR         ' The sensor clickable is displayed
    IDES_ECM            ' The ECM clickable is displayed
    IDES_CONSTRUCT      ' The Constructor clickable is displayed
    IDES_REPAIR         ' The Repair clickable is displayed
    IDES_WEAPON         ' The Weapon clickable is displayed
    IDES_COMMAND        ' The command droid clickable is displayed
    IDES_NOSYSTEM       ' No system clickable has been displayed
End Enum

' The major component tabs on the design screen
Private Const IDES_MAINTAB As Long = 0
Private Const IDES_EXTRATAB As Long = 1

' Which component type is being selected on the design screen
'added IDES_TURRET_A,IDES_TURRET_B,changing the name of IDES_TURRET might break exist codes
Private Enum DES_COMPMODE
    IDES_SYSTEM         ' The main system for the droid (sensor, ECM, constructor)
    IDES_TURRET         ' The weapon for the droid
    IDES_BODY           ' The droid body
    IDES_PROPULSION     ' The propulsion system
    IDES_NOCOMPONENT    ' No system has been selected
    IDES_TURRET_A       ' The 2nd turret
    IDES_TURRET_B       ' The 3rd turret
End Enum

' Which type of propulsion is being selected
Private Enum DES_PROPMODE
    IDES_GROUND         ' Ground propulsion (wheeled, tracked, etc).
    IDES_AIR            ' Air propulsion
    IDES_NOPROPULSION   ' No propulsion has been selected
End Enum

Private Const STRING_BUFFER_SIZE As Long = 32 * MAX_STR_LENGTH

' Design screen positions (DESIGN_Y and the positions built on it: functions below)

Private Const DES_TABBUTGAP As Long = 2
Private Const DES_TABBUTWIDTH As Long = 60
Private Const DES_TABBUTHEIGHT As Long = 46

Private Const DES_LEFTFORMX As Long = RET_X
Private Const DES_LEFTFORMWIDTH As Long = RET_FORMWIDTH
Private Const DES_LEFTFORMHEIGHT As Long = 273
Private Const DES_LEFTFORMBUTX As Long = 2
Private Const DES_LEFTFORMBUTY As Long = 2

Private Const DES_CENTERFORMWIDTH As Long = 315
Private Const DES_CENTERFORMHEIGHT As Long = 262

Private Const DES_3DVIEWX As Long = 8
Private Const DES_3DVIEWY As Long = 25
Private Const DES_3DVIEWWIDTH As Long = 236
Private Const DES_3DVIEWHEIGHT As Long = 192

Private Const DES_STATSFORMWIDTH As Long = DES_CENTERFORMWIDTH
Private Const DES_STATSFORMHEIGHT As Long = 100

Private Const DES_PARTFORMX As Long = DES_3DVIEWX + DES_3DVIEWWIDTH + 2
Private Const DES_PARTFORMY As Long = DES_3DVIEWY
Private Const DES_PARTFORMHEIGHT As Long = DES_3DVIEWHEIGHT

Private Const DES_POWERFORMX As Long = DES_3DVIEWX
Private Const DES_POWERFORMY As Long = (DES_3DVIEWY + DES_3DVIEWHEIGHT + 2)
Private Const DES_POWERFORMWIDTH As Long = (DES_CENTERFORMWIDTH - 2 * DES_POWERFORMX)
Private Const DES_POWERFORMHEIGHT As Long = 40

Private Const DES_RIGHTFORMWIDTH As Long = (RET_FORMWIDTH + 20)
Private Const DES_RIGHTFORMHEIGHT As Long = DES_LEFTFORMHEIGHT
Private Const DES_RIGHTFORMBUTX As Long = 2
Private Const DES_RIGHTFORMBUTY As Long = 2

Private Const DES_BARFORMX As Long = 6
Private Const DES_BARFORMY As Long = 6
Private Const DES_BARFORMWIDTH As Long = 300
Private Const DES_BARFORMHEIGHT As Long = 85

Private Const DES_NAMEBOXX As Long = DES_3DVIEWX
Private Const DES_NAMEBOXY As Long = 6
Private Const DES_NAMEBOXWIDTH As Long = DES_CENTERFORMWIDTH - 2 * DES_NAMEBOXX
Private Const DES_NAMEBOXHEIGHT As Long = 14

' The central boxes on the design screen
Private Const DES_COMPBUTWIDTH As Long = 150
Private Const DES_COMPBUTHEIGHT As Long = 85

Private Const DES_POWERX As Long = 1
Private Const DES_POWERY As Long = 6
Private Const DES_POWERSEPARATIONX As Long = 4
Private Const DES_POWERSEPARATIONY As Long = 2

Private Const DES_PARTSEPARATIONX As Long = 6
Private Const DES_PARTSEPARATIONY As Long = 6

' Positions of stuff on the clickable boxes (Design screen)
Private Const DES_CLICKBARX As Long = 154
Private Const DES_CLICKBARY As Long = 7
Private Const DES_CLICKBARWIDTH As Long = 140
Private Const DES_CLICKBARHEIGHT As Long = 11
Private Const DES_CLICKGAP As Long = 9
Private Const DES_CLICKBARNAMEX As Long = 126
Private Const DES_CLICKBARNAMEWIDTH As Long = 20
Private Const DES_CLICKBARNAMEHEIGHT As Long = 19

Private Const DES_CLICKBARMAJORRED As Long = 255        '0xcc
Private Const DES_CLICKBARMAJORGREEN As Long = 235      '0
Private Const DES_CLICKBARMAJORBLUE As Long = 19        '0
Private Const DES_CLICKBARMINORRED As Long = &H55
Private Const DES_CLICKBARMINORGREEN As Long = 0
Private Const DES_CLICKBARMINORBLUE As Long = 0

Private Const DES_WEAPONBUTTON_X As Long = 26
Private Const DES_SYSTEMBUTTON_X As Long = 68
Private Const DES_SYSTEMBUTTON_Y As Long = 10

' Stat bar y positions.
Private Const DES_STATBAR_Y1 As Long = (DES_CLICKBARY)
Private Const DES_STATBAR_Y2 As Long = (DES_CLICKBARY + DES_CLICKBARHEIGHT + DES_CLICKGAP)
Private Const DES_STATBAR_Y3 As Long = (DES_CLICKBARY + (DES_CLICKBARHEIGHT + DES_CLICKGAP) * 2)
Private Const DES_STATBAR_Y4 As Long = (DES_CLICKBARY + (DES_CLICKBARHEIGHT + DES_CLICKGAP) * 3)

Private desSysMode As DES_SYSMODE
Private desCompMode As DES_COMPMODE
Private desPropMode As DES_PROPMODE

Public StringBuffer As String           ' char StringBuffer[STRING_BUFFER_SIZE]

' default droid design template
Private sDefaultDesignTemplate As Long  ' (VB) static DROID_TEMPLATE: a pool slot, made once by desSetupDesignTemplates

' The current name of the design
Private aCurrName As String

' The button id of the component that is in the design
Private desCompID As Long

Private droidTemplID As Long

' The current design being edited on the design screen
Private sCurrDesign As DROID_TEMPLATE
Private sShadowDesign As DROID_TEMPLATE

Private designScratch As Long           ' (VB) the pool slot designTemplateHandle copies a design into

' (VB) sstrcpy / sstrcat into aCurrName (char[MAX_STR_LENGTH])
Private Sub sstrcpy_aCurrName(ByVal src As String)
    aCurrName = Left$(src, MAX_STR_LENGTH - 1)
End Sub

Private Sub sstrcat_aCurrName(ByVal src As String)
    aCurrName = Left$(aCurrName & src, MAX_STR_LENGTH - 1)
End Sub

' (VB) #define DESIGN_Y (59 + D_H): the top left y value for all forms on the design screen
Private Function DESIGN_Y() As Long
    DESIGN_Y = (59 + D_H)
End Function

Private Function DES_LEFTFORMY() As Long
    DES_LEFTFORMY = DESIGN_Y
End Function

Private Function DES_CENTERFORMX() As Long
    DES_CENTERFORMX = POW_X
End Function

Private Function DES_CENTERFORMY() As Long
    DES_CENTERFORMY = DESIGN_Y
End Function

Private Function DES_STATSFORMX() As Long
    DES_STATSFORMX = POW_X
End Function

Private Function DES_STATSFORMY() As Long
    DES_STATSFORMY = (DES_CENTERFORMY + DES_CENTERFORMHEIGHT + 3)
End Function

' (VB) a PIELIGHT's byte.r (0), .g (1), .b (2) or .a (3) set to v
Private Function PIELIGHT_withByte(ByVal colour As Long, ByVal idx As Long, ByVal v As Long) As Long
    Select Case idx
    Case 0: PIELIGHT_withByte = (colour And &HFFFFFF00) Or (v And &HFF&)
    Case 1: PIELIGHT_withByte = (colour And &HFFFF00FF) Or SHL(v And &HFF&, 8)
    Case 2: PIELIGHT_withByte = (colour And &HFF00FFFF) Or SHL(v And &HFF&, 16)
    Case Else: PIELIGHT_withByte = (colour And &HFFFFFF) Or SHL(v And &HFF&, 24)
    End Select
End Function

' (VB) a DROID_TEMPLATE * for a design held outside the pool (sCurrDesign, a local copy): its copy in designScratch
Private Function designTemplateHandle(psTemplate As DROID_TEMPLATE) As Long
    apsTemplates(designScratch) = psTemplate
    designTemplateHandle = designScratch
End Function

' (VB) apsTemplateList.push_back(psTempl)
Private Sub apsTemplateList_add(ByVal psTempl As Long)
    ReDim Preserve apsTemplateList(0 To apsTemplateList_size)
    apsTemplateList(apsTemplateList_size) = psTempl
    apsTemplateList_size = apsTemplateList_size + 1
End Sub

' (VB) localTemplates.erase(i): the template's slot goes back to the pool
Private Sub localTemplates_erase_(ByVal i As Long)
    Dim k As Long
    deleteTemplate localTemplates(i)
    For k = i To localTemplates_size - 2
        localTemplates(k) = localTemplates(k + 1)
    Next
    localTemplates_size = localTemplates_size - 1
    If localTemplates_size = 0 Then Erase localTemplates
End Sub

' class DesignStatsBar: public W_BARGRAPH
Public Sub DesignStatsBar_construct(ByVal self As WIDGET, init As W_BARINIT)
    Dim d As DesignStatsBar, a As Animation_float
    W_BARGRAPH_construct_init self, init
    self.kind = K_DesignStatsBar
    Set d = DesignStatsBar_(self)
    d.maxValueTextWidth = iV_GetTextWidth("00000", font_regular)
    a = Animation_float_new(ANIMATION_TIME_REAL, EASE_IN, 200)
    Animation_float_setInitialData a, 0
    Animation_float_setFinalData a, 0
    d.minorAnimation = a
    d.PADDING_ = 3
    d.lessIsBetter = False
End Sub

Public Function DesignStatsBar_new(init As W_BARINIT) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_DesignStatsBar)
    DesignStatsBar_construct w, init
    Set DesignStatsBar_new = w
End Function

Public Function DesignStatsBar_makeLessIsBetter(init As W_BARINIT) As WIDGET
    Dim w As WIDGET
    Set w = DesignStatsBar_new(init)
    DesignStatsBar_(w).lessIsBetter = True
    Set DesignStatsBar_makeLessIsBetter = w
End Function

Public Sub DesignStatsBar_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As DesignStatsBar, b As W_BARGRAPH, x0 As Long, y0 As Long, iX As Long, iY As Long, barWidth As Long, filledWidth As Long
    Dim a As Animation_float, delta As Long, xDeltaText As Long, colour As Long, v As Single
    Set d = DesignStatsBar_(self)
    Set b = W_BARGRAPH_(self)
    x0 = xOffset + self.x() + d.PADDING_
    y0 = yOffset + self.y() + d.PADDING_

    v = CSng(b.majorValue) / CSng(b.denominator)
    d.valueText.setText va("%.*f", b.precision_, CDbl(v)), font_regular

    ' indent to allow text value
    iX = x0 + d.maxValueTextWidth
    iY = y0 + (iV_GetImageHeight(IntImages, IMAGE_DES_STATSCURR) - d.valueText.lineSize()) \ 2 - d.valueText.aboveBase()

    'draw current value section
    barWidth = self.width() - d.maxValueTextWidth - 2 * d.PADDING_
    filledWidth = MIN_(b.majorSize * barWidth \ 100, barWidth)
    iV_DrawImageRepeatX IntImages, IMAGE_DES_STATSCURR, iX, y0, filledWidth, 0, True

    d.valueText.render x0, iY, psPalette(WZCOL_TEXT_BRIGHT)

    If b.minorValue = 0 Then
        Exit Sub
    End If

    'draw the comparison value - only if not zero
    DesignStatsBar_updateAnimation self
    a = d.minorAnimation
    v = Animation_float_getCurrent(a) * barWidth
    v = v / 100
    filledWidth = MIN_(Fix(v), barWidth)
    iV_DrawImage IntImages, IMAGE_DES_STATSCOMP, iX + filledWidth, y0 - 1

    delta = b.minorValue - b.majorValue
    If delta <> 0 Then
        v = CSng(delta) / CSng(b.denominator)
        d.deltaText.setText va("%+.*f", b.precision_, CDbl(v)), font_small
        xDeltaText = xOffset + self.x() + self.width() - d.deltaText.width() - d.PADDING_
        If (delta < 0) = d.lessIsBetter Then colour = psPalette(WZCOL_LGREEN) Else colour = psPalette(WZCOL_LRED)
        d.deltaText.renderOutlined xDeltaText, iY - 1, colour, pal_RGBA(0, 0, 0, 192)
    End If
End Sub

Public Sub DesignStatsBar_updateAnimation(ByVal self As WIDGET)
    Dim d As DesignStatsBar, a As Animation_float
    Set d = DesignStatsBar_(self)
    a = d.minorAnimation
    If CSng(W_BARGRAPH_(self).minorSize) <> Animation_float_getFinalData(a) Then
        Animation_float_setInitialData a, Animation_float_getCurrent(a)
        Animation_float_setFinalData a, W_BARGRAPH_(self).minorSize
        Animation_float_start a
    End If

    Animation_float_update a
    d.minorAnimation = a
End Sub

' class DesignPowerBar: public DesignStatsBar
Public Sub DesignPowerBar_construct(ByVal self As WIDGET, init As W_BARINIT)
    DesignStatsBar_construct self, init
    self.kind = K_DesignPowerBar
End Sub

Public Function DesignPowerBar_new(init As W_BARINIT) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_DesignPowerBar)
    DesignPowerBar_construct w, init
    Set DesignPowerBar_new = w
End Function

Public Function DesignPowerBar_makeLessIsBetter(init As W_BARINIT) As WIDGET
    Dim w As WIDGET
    Set w = DesignPowerBar_new(init)
    DesignStatsBar_(w).lessIsBetter = True
    Set DesignPowerBar_makeLessIsBetter = w
End Function

Public Sub DesignPowerBar_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    iV_DrawImage IntImages, IMAGE_DES_POWERBAR_LEFT, x0, y0
    iV_DrawImage IntImages, IMAGE_DES_POWERBAR_RIGHT, x0 + self.width() - iV_GetImageWidth(IntImages, IMAGE_DES_POWERBAR_RIGHT), y0
    DesignStatsBar_display self, xOffset, yOffset
End Sub

' (VB) psStats: the first element's handle (asXStats.data()); element i is psStats + i (the size is not needed);
' aAvailable is apCompLists[aAvailablePlayer][aAvailableType]
Private Function componentIterator(ByVal psStats As Long, ByVal aAvailablePlayer As Long, ByVal aAvailableType As COMPONENT_TYPE, ByVal numEntries As Long) As WzFunc
    Set componentIterator = WzFunc_new(FN_componentIterator, psStats, aAvailablePlayer, aAvailableType, numEntries)
End Function

Private Function bodyIterator() As WzFunc
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/design.cpp:471", "bodyIterator", "selectedPlayer: %u", selectedPlayer
    Set bodyIterator = componentIterator(STAT_BODY, selectedPlayer, COMP_BODY, asBodyStats_size)
End Function

Private Function weaponIterator() As WzFunc
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/design.cpp:477", "weaponIterator", "selectedPlayer: %u", selectedPlayer
    Set weaponIterator = componentIterator(STAT_WEAPON, selectedPlayer, COMP_WEAPON, asWeaponStats_size)
End Function

Private Function propulsionIterator() As WzFunc
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/design.cpp:483", "propulsionIterator", "selectedPlayer: %u", selectedPlayer
    Set propulsionIterator = componentIterator(STAT_PROPULSION, selectedPlayer, COMP_PROPULSION, asPropulsionStats_size)
End Function

Private Function sensorIterator() As WzFunc
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/design.cpp:489", "sensorIterator", "selectedPlayer: %u", selectedPlayer
    Set sensorIterator = componentIterator(STAT_SENSOR, selectedPlayer, COMP_SENSOR, asSensorStats_size)
End Function

Private Function ecmIterator() As WzFunc
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/design.cpp:495", "ecmIterator", "selectedPlayer: %u", selectedPlayer
    Set ecmIterator = componentIterator(STAT_ECM, selectedPlayer, COMP_ECM, asECMStats_size)
End Function

Private Function constructorIterator() As WzFunc
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/design.cpp:501", "constructorIterator", "selectedPlayer: %u", selectedPlayer
    Set constructorIterator = componentIterator(STAT_CONSTRUCT, selectedPlayer, COMP_CONSTRUCT, asConstructStats_size)
End Function

Private Function repairIterator() As WzFunc
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/design.cpp:507", "repairIterator", "selectedPlayer: %u", selectedPlayer
    Set repairIterator = componentIterator(STAT_REPAIR, selectedPlayer, COMP_REPAIRUNIT, asRepairStats_size)
End Function

Private Function brainIterator() As WzFunc
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/design.cpp:513", "brainIterator", "selectedPlayer: %u", selectedPlayer
    Set brainIterator = componentIterator(STAT_BRAIN, selectedPlayer, COMP_BRAIN, asBrainStats_size)
End Function

' (VB) std::vector<ComponentIterator> iterators: up to 7 WzFuncs
Private Function concatIterators(ParamArray iterators() As Variant) As WzFunc
    Dim f As WzFunc, i As Long
    Set f = WzFunc_new(FN_concatIterators)
    f.cap(0) = UBound(iterators) + 1
    For i = 0 To UBound(iterators)
        Set f.cap(i + 1) = iterators(i)
    Next
    Set concatIterators = f
End Function

Private Function extraSystemIterator() As WzFunc
    Set extraSystemIterator = concatIterators(sensorIterator(), ecmIterator(), constructorIterator(), repairIterator(), brainIterator())
End Function

' (VB) value: std::function<uint32_t(COMPONENT_STATS &)>
Private Function findMax(ByVal componentIterator_ As WzFunc, ByVal value As WzFunc) As Long
    Dim callback As WzFunc
    ' uint32_t max = 0 (the callback's cap(1), which [&] shares)
    Set callback = WzFunc_new(FN_findMax_callback, value, 0&)

    WzFunc_call componentIterator_, callback

    findMax = callback.cap(1)
End Function

Private Function findMaxWeight(ByVal componentIterator_ As WzFunc) As Long
    findMaxWeight = findMax(componentIterator_, WzFunc_new(FN_findMaxWeight_value))
End Function

Private Function findMaxPropulsionSpeed(ByVal terrainType As Long) As Long
    findMaxPropulsionSpeed = findMax(propulsionIterator(), WzFunc_new(FN_findMaxPropulsionSpeed_value, terrainType))
End Function

' (VB) attributeGetter: std::function<uint32_t(const WEAPON_STATS&, int)>
Private Function findMaxWeaponAttribute(ByVal attributeGetter As WzFunc) As Long
    findMaxWeaponAttribute = findMax(weaponIterator(), WzFunc_new(FN_findMaxWeaponAttribute_value, attributeGetter))
End Function

Private Function getDesignMaxBodyArmour(ByVal weaponClass As WEAPON_CLASS) As Long
    getDesignMaxBodyArmour = findMax(bodyIterator(), WzFunc_new(FN_getDesignMaxBodyArmour_value, weaponClass))
End Function

Private Function getDesignMaxEngineOutput() As Long
    getDesignMaxEngineOutput = findMax(bodyIterator(), WzFunc_new(FN_getDesignMaxEngineOutput_value))
End Function

Private Function calcShadowBodyPoints(ByVal psStats As Long) As Long
    Dim designCopy As DROID_TEMPLATE
    designCopy = sCurrDesign
    setTemplateStat designCopy, psStats
    calcShadowBodyPoints = calcTemplateBody(designTemplateHandle(designCopy), selectedPlayer)
End Function

Private Function calcShadowPower(ByVal psStats As Long) As Long
    Dim designCopy As DROID_TEMPLATE
    designCopy = sCurrDesign
    setTemplateStat designCopy, psStats
    calcShadowPower = calcTemplatePower(designCopy)
End Function

Private Function getDesignMaxSensorRange() As Long
    getDesignMaxSensorRange = findMax(sensorIterator(), WzFunc_new(FN_getDesignMaxSensorRange_value))
End Function

Private Function getDesignMaxEcmRange() As Long
    getDesignMaxEcmRange = findMax(ecmIterator(), WzFunc_new(FN_getDesignMaxEcmRange_value))
End Function

Private Function getDesignMaxBuildPoints() As Long
    getDesignMaxBuildPoints = findMax(constructorIterator(), WzFunc_new(FN_getDesignMaxBuildPoints_value))
End Function

' Add the design widgets to the widget screen
Public Function intAddDesign(ByVal bShowCentreScreen As Boolean) As Boolean
    Dim sFormInit As W_FORMINIT, sLabInit As W_LABINIT, sEdInit As W_EDBINIT, sButInit As W_BUTINIT, sBarInit As W_BARINIT
    Dim parent As WIDGET, desForm As WIDGET, statsForm As WIDGET, bodyForm As WIDGET, powerForm As WIDGET, spectator As Boolean
    W_FORMINIT_construct sFormInit
    W_LABINIT_construct sLabInit
    W_EDBINIT_construct sEdInit
    W_BUTINIT_construct sButInit
    W_BARINIT_construct sBarInit

    If bMultiPlayer Then spectator = NetPlay_players_isSpectator(selectedPlayer)  ' (VB) bMultiPlayer && ..., in order
    If Not (Not spectator) Then ASSERT_FAILURE "!(bMultiPlayer && NetPlay.players[selectedPlayer].isSpectator)", "src/design.cpp:633", "intAddDesign", "Spectators can't open design mode": intAddDesign = False: Exit Function
    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/design.cpp:634", "intAddDesign", "selectedPlayer: %u", selectedPlayer: intAddDesign = False: Exit Function

    desSetupDesignTemplates

    'set which states are to be paused while design screen is up
    setDesignPauseState

    Set parent = psWScreen.psForm

    ' Add the main design form
    Set desForm = IntFormAnimated_new(False)
    desForm.id = IDDES_FORM
    parent.attach desForm
    desForm.setCalcLayout WzFunc_new(FN_intAddDesign_desForm_calcLayout)

    ' add the edit name box
    sEdInit.formID = IDDES_FORM
    sEdInit.id = IDDES_NAMEBOX
    sEdInit.x = DES_NAMEBOXX
    sEdInit.y = DES_NAMEBOXY
    sEdInit.width = DES_NAMEBOXWIDTH
    sEdInit.height = DES_NAMEBOXHEIGHT
    sEdInit.pText = "New Vehicle"
    Set sEdInit.pBoxDisplay = WzFunc_new(FN_intDisplayEditBox)
    If widgAddEditBox(psWScreen, sEdInit) Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' Initialise the current design
    apsTemplates(sDefaultDesignTemplate).droidType = DROID_ANY
    sCurrDesign = apsTemplates(sDefaultDesignTemplate)
    sCurrDesign.stored = False
    sstrcpy_aCurrName "New Vehicle"
    sCurrDesign.name_ = aCurrName

    ' Add the design templates form
    If Not intAddTemplateForm(0) Then  ' Was psCurrTemplate instead of NULL, but psCurrTemplate was always NULL. Deleted psCurrTemplate, but leaving this here, in case intAddTemplateForm(NULL) does something useful.
        intAddDesign = False
        Exit Function
    End If

    ' Add the 3D View form
    sFormInit.formID = IDDES_FORM
    sFormInit.id = IDDES_3DVIEW
    sFormInit.style = WFORM_PLAIN
    sFormInit.x = DES_3DVIEWX
    sFormInit.y = DES_3DVIEWY
    sFormInit.width = DES_3DVIEWWIDTH
    sFormInit.height = DES_3DVIEWHEIGHT
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayViewForm)
    If widgAddForm(psWScreen, sFormInit) Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' Add the part button form
    sFormInit.formID = IDDES_FORM
    sFormInit.id = IDDES_PARTFORM
    sFormInit.style = WFORM_PLAIN
    sFormInit.x = DES_PARTFORMX
    sFormInit.y = DES_PARTFORMY
    sFormInit.width = (iV_GetImageWidth(IntImages, IMAGE_DES_TURRET) + _
                       2 * DES_PARTSEPARATIONX) And &HFFFF&
    sFormInit.height = DES_PARTFORMHEIGHT
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayDesignForm)
    If widgAddForm(psWScreen, sFormInit) Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' add the body part button
    sButInit.formID = IDDES_PARTFORM
    sButInit.id = IDDES_BODYBUTTON
    sButInit.x = DES_PARTSEPARATIONX
    sButInit.y = DES_PARTSEPARATIONY
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_DES_BODY)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_BODY)
    sButInit.pTip = "Vehicle Body"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayButtonFlash)
    sButInit.UserData = PACKDWORD_TRI(1, IMAGE_DES_BODYH, IMAGE_DES_BODY)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' add the propulsion part button
    sButInit.formID = IDDES_PARTFORM
    sButInit.id = IDDES_PROPBUTTON
    sButInit.x = DES_PARTSEPARATIONX
    sButInit.y = (iV_GetImageHeight(IntImages, IMAGE_DES_PROPULSION) + _
                  2 * DES_PARTSEPARATIONY) And &HFFFF&
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_DES_PROPULSION)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_PROPULSION)
    sButInit.pTip = "Vehicle Propulsion"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayButtonFlash)
    sButInit.UserData = PACKDWORD_TRI(1, IMAGE_DES_PROPULSIONH, IMAGE_DES_PROPULSION)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' add the turret part button
    sButInit.formID = IDDES_PARTFORM
    sButInit.id = IDDES_SYSTEMBUTTON
    sButInit.x = DES_PARTSEPARATIONX
    sButInit.y = (iV_GetImageHeight(IntImages, IMAGE_DES_PROPULSION) + _
                  iV_GetImageHeight(IntImages, IMAGE_DES_BODY) + _
                  3 * DES_PARTSEPARATIONY) And &HFFFF&
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_DES_TURRET)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_TURRET)
    sButInit.pTip = "Vehicle Turret"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayButtonFlash)
    sButInit.UserData = PACKDWORD_TRI(1, IMAGE_DES_TURRETH, IMAGE_DES_TURRET)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' add the turret_a button
    sButInit.formID = IDDES_PARTFORM
    sButInit.id = IDDES_WPABUTTON
    sButInit.x = DES_PARTSEPARATIONX
    ' use BODY height for now
    sButInit.y = (iV_GetImageHeight(IntImages, IMAGE_DES_PROPULSION) + _
                  iV_GetImageHeight(IntImages, IMAGE_DES_BODY) + _
                  iV_GetImageHeight(IntImages, IMAGE_DES_BODY) + _
                  4 * DES_PARTSEPARATIONY) And &HFFFF&
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_DES_TURRET)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_TURRET)
    sButInit.pTip = "Vehicle Turret"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayButtonFlash)
    sButInit.UserData = PACKDWORD_TRI(1, IMAGE_DES_TURRETH, IMAGE_DES_TURRET)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' add the turret_b button
    sButInit.formID = IDDES_PARTFORM
    sButInit.id = IDDES_WPBBUTTON
    sButInit.x = DES_PARTSEPARATIONX
    'use body height for now
    sButInit.y = (iV_GetImageHeight(IntImages, IMAGE_DES_PROPULSION) + _
                  iV_GetImageHeight(IntImages, IMAGE_DES_BODY) + _
                  iV_GetImageHeight(IntImages, IMAGE_DES_BODY) + _
                  iV_GetImageHeight(IntImages, IMAGE_DES_BODY) + _
                  5 * DES_PARTSEPARATIONY) And &HFFFF&
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_DES_TURRET)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_TURRET)
    sButInit.pTip = "Vehicle Turret"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayButtonFlash)
    sButInit.UserData = PACKDWORD_TRI(1, IMAGE_DES_TURRETH, IMAGE_DES_TURRET)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' add the delete button
    sButInit.formID = IDDES_PARTFORM
    sButInit.id = IDDES_BIN
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_DES_BIN)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_BIN)
    sButInit.x = DES_PARTSEPARATIONX
    sButInit.y = (DES_PARTFORMHEIGHT - sButInit.height - DES_PARTSEPARATIONY) And &HFFFF&
    sButInit.pTip = "Delete Design"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayButtonHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_DES_BINH, IMAGE_DES_BIN)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' Add the store template button
    sButInit.formID = IDDES_PARTFORM
    sButInit.id = IDDES_STOREBUTTON
    sButInit.style = WBUT_PLAIN
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_DES_SAVE)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_SAVE)
    sButInit.x = DES_PARTSEPARATIONX
    sButInit.y = DES_PARTFORMHEIGHT - 2 * sButInit.height - 2 * DES_PARTSEPARATIONY
    sButInit.pTip = "Store Template"
    sButInit.FontID = font_regular
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayButtonHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_DES_SAVEH, IMAGE_DES_SAVE)

    If bMultiPlayer Then
        If widgAddButton(psWScreen, sButInit) Is Nothing Then
            intAddDesign = False
            Exit Function
        End If
    End If

    ' add central stats form
    Set statsForm = IntFormAnimated_new(False)
    parent.attach statsForm
    statsForm.id = IDDES_STATSFORM
    statsForm.setCalcLayout WzFunc_new(FN_intAddDesign_statsForm_calcLayout)

    ' Add the body form
    sFormInit.formID = IDDES_STATSFORM
    sFormInit.id = IDDES_BODYFORM
    sFormInit.style = WFORM_CLICKABLE Or WFORM_NOCLICKMOVE
    sFormInit.width = DES_BARFORMWIDTH
    sFormInit.height = DES_BARFORMHEIGHT
    sFormInit.x = DES_BARFORMX
    sFormInit.y = DES_BARFORMY
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayStatForm)
    Set bodyForm = widgAddForm(psWScreen, sFormInit)
    If bodyForm Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' Add the graphs for the Body
    sBarInit.formID = IDDES_BODYFORM
    sBarInit.id = IDDES_BODYARMOUR_K
    sBarInit.x = DES_CLICKBARX
    sBarInit.y = DES_STATBAR_Y1    'DES_CLICKBARY;
    sBarInit.width = DES_CLICKBARWIDTH
    sBarInit.height = DES_CLICKBARHEIGHT
    sBarInit.size = 50
    sBarInit.sCol = PIELIGHT_withByte(sBarInit.sCol, 0, DES_CLICKBARMAJORRED)
    sBarInit.sCol = PIELIGHT_withByte(sBarInit.sCol, 1, DES_CLICKBARMAJORGREEN)
    sBarInit.sCol = PIELIGHT_withByte(sBarInit.sCol, 2, DES_CLICKBARMAJORBLUE)
    sBarInit.sMinorCol = PIELIGHT_withByte(sBarInit.sMinorCol, 0, DES_CLICKBARMINORRED)
    sBarInit.sMinorCol = PIELIGHT_withByte(sBarInit.sMinorCol, 1, DES_CLICKBARMINORGREEN)
    sBarInit.sMinorCol = PIELIGHT_withByte(sBarInit.sMinorCol, 2, DES_CLICKBARMINORBLUE)
    sBarInit.pTip = "Kinetic Armour"
    sBarInit.iRange = getDesignMaxBodyArmour(WC_KINETIC)
    bodyForm.attach DesignStatsBar_new(sBarInit)

    sBarInit.id = IDDES_BODYARMOUR_H
    sBarInit.y = DES_STATBAR_Y2    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
    sBarInit.pTip = "Thermal Armour"
    sBarInit.iRange = getDesignMaxBodyArmour(WC_HEAT)
    bodyForm.attach DesignStatsBar_new(sBarInit)

    sBarInit.id = IDDES_BODYPOWER
    sBarInit.y = DES_STATBAR_Y3    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
    sBarInit.pTip = "Engine Output"
    sBarInit.iRange = getDesignMaxEngineOutput()
    bodyForm.attach DesignStatsBar_new(sBarInit)

    sBarInit.id = IDDES_BODYWEIGHT
    sBarInit.y = DES_STATBAR_Y4    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
    sBarInit.pTip = "Weight"
    sBarInit.iRange = findMaxWeight(bodyIterator())
    bodyForm.attach DesignStatsBar_makeLessIsBetter(sBarInit)

    ' Add the labels for the Body
    sLabInit.formID = IDDES_BODYFORM
    sLabInit.id = IDDES_BODYARMOURKLAB
    sLabInit.x = DES_CLICKBARNAMEX
    sLabInit.y = DES_CLICKBARY - DES_CLICKBARHEIGHT \ 3
    sLabInit.width = DES_CLICKBARNAMEWIDTH
    sLabInit.height = DES_CLICKBARHEIGHT
    sLabInit.pTip = "Kinetic Armour"
    Set sLabInit.pDisplay = WzFunc_new(FN_intDisplayImage)
    'just to confuse things even more - the graphics were named incorrectly!
    sLabInit.UserData = IMAGE_DES_ARMOUR_EXPLOSIVE
    If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
        intAddDesign = True
        Exit Function
    End If
    sLabInit.id = IDDES_BODYARMOURHLAB
    sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
    sLabInit.pTip = "Thermal Armour"
    Set sLabInit.pDisplay = WzFunc_new(FN_intDisplayImage)
    sLabInit.UserData = IMAGE_DES_ARMOUR_KINETIC
    If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
        intAddDesign = True
        Exit Function
    End If
    sLabInit.id = IDDES_BODYPOWERLAB
    sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
    sLabInit.pTip = "Engine Output"
    Set sLabInit.pDisplay = WzFunc_new(FN_intDisplayImage)
    sLabInit.UserData = IMAGE_DES_POWER
    If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
        intAddDesign = True
        Exit Function
    End If
    sLabInit.id = IDDES_BODYWEIGHTLAB
    sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
    sLabInit.pTip = "Weight"
    Set sLabInit.pDisplay = WzFunc_new(FN_intDisplayImage)
    sLabInit.UserData = IMAGE_DES_WEIGHT
    If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
        intAddDesign = True
        Exit Function
    End If

    ' add power/points bar subform
    W_FORMINIT_construct sFormInit
    sFormInit.formID = IDDES_FORM
    sFormInit.id = IDDES_POWERFORM
    sFormInit.style = WFORM_PLAIN
    sFormInit.x = DES_POWERFORMX
    sFormInit.y = DES_POWERFORMY
    sFormInit.width = DES_POWERFORMWIDTH
    sFormInit.height = DES_POWERFORMHEIGHT
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayDesignForm)
    Set powerForm = widgAddForm(psWScreen, sFormInit)
    If powerForm Is Nothing Then
        intAddDesign = False
        Exit Function
    End If

    ' Add the design template power bar and label
    sLabInit.formID = IDDES_POWERFORM
    sLabInit.id = IDDES_TEMPPOWERLAB
    sLabInit.x = DES_POWERX
    sLabInit.y = DES_POWERY
    sLabInit.pTip = "Total Power Required"
    Set sLabInit.pDisplay = WzFunc_new(FN_intDisplayImage)
    sLabInit.UserData = IMAGE_DES_POWER
    widgAddLabel psWScreen, sLabInit

    W_BARINIT_construct sBarInit
    sBarInit.formID = IDDES_POWERFORM
    sBarInit.id = IDDES_POWERBAR
    sBarInit.x = (DES_POWERX + DES_POWERSEPARATIONX + _
                  iV_GetImageWidth(IntImages, IMAGE_DES_BODYPOINTS))
    sBarInit.y = DES_POWERY
    sBarInit.width = (DES_POWERFORMWIDTH - 15 - _
                      iV_GetImageWidth(IntImages, IMAGE_DES_BODYPOINTS))
    sBarInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_POWERBACK)
    sBarInit.pTip = "Total Power Required"
    sBarInit.iRange = DBAR_TEMPLATEMAXPOWER    'WBAR_SCALE;
    powerForm.attach DesignPowerBar_makeLessIsBetter(sBarInit)

    ' Add the design template body points bar and label
    sLabInit.formID = IDDES_POWERFORM
    sLabInit.id = IDDES_TEMPBODYLAB
    sLabInit.x = DES_POWERX
    sLabInit.y = (DES_POWERY + DES_POWERSEPARATIONY + _
                  iV_GetImageHeight(IntImages, IMAGE_DES_BODYPOINTS))
    sLabInit.pTip = "Total Body Points"
    Set sLabInit.pDisplay = WzFunc_new(FN_intDisplayImage)
    sLabInit.UserData = IMAGE_DES_BODYPOINTS
    widgAddLabel psWScreen, sLabInit

    W_BARINIT_construct sBarInit
    sBarInit.formID = IDDES_POWERFORM
    sBarInit.id = IDDES_BODYPOINTS
    sBarInit.x = (DES_POWERX + DES_POWERSEPARATIONX + _
                  iV_GetImageWidth(IntImages, IMAGE_DES_BODYPOINTS))
    sBarInit.y = (DES_POWERY + DES_POWERSEPARATIONY + 4 + _
                  iV_GetImageHeight(IntImages, IMAGE_DES_BODYPOINTS))
    sBarInit.width = (DES_POWERFORMWIDTH - 15 - _
                      iV_GetImageWidth(IntImages, IMAGE_DES_BODYPOINTS))
    sBarInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_POWERBACK)
    sBarInit.pTip = "Total Body Points"
    sBarInit.iRange = DBAR_TEMPLATEMAXPOINTS
    powerForm.attach DesignPowerBar_new(sBarInit)

    ' Add the variable bits of the design screen and set the bar graphs
    desCompMode = IDES_NOCOMPONENT
    desSysMode = IDES_NOSYSTEM
    desPropMode = IDES_NOPROPULSION
    intSetDesignStats sCurrDesign
    intSetBodyPoints sCurrDesign
    intSetDesignPower sCurrDesign
    intSetDesignMode IDES_BODY

    ' hide design and component forms until required
    desForm.show bShowCentreScreen
    statsForm.hide
    widgHide psWScreen, IDDES_RIGHTBASE

    intAddDesign = True
End Function

' set up droid templates before going into design screen
Private Sub desSetupDesignTemplates()
    Dim k As Long, templ As Long

    ' (VB) the static sDefaultDesignTemplate (DROID_TEMPLATE()) and the scratch slot: made once
    If sDefaultDesignTemplate = 0 Then sDefaultDesignTemplate = newTemplate()
    If designScratch = 0 Then designScratch = newTemplate()

    ' init template list
    Erase apsTemplateList
    apsTemplateList_size = 0
    apsTemplateList_add sDefaultDesignTemplate
    For k = 0 To localTemplates_size - 1
        templ = localTemplates(k)
        ' add template to list if not a transporter,
        ' cyborg, person or command droid,
        If apsTemplates(templ).droidType <> DROID_TRANSPORTER And _
           apsTemplates(templ).droidType <> DROID_SUPERTRANSPORTER And _
           apsTemplates(templ).droidType <> DROID_CYBORG And _
           apsTemplates(templ).droidType <> DROID_CYBORG_SUPER And _
           apsTemplates(templ).droidType <> DROID_CYBORG_CONSTRUCT And _
           apsTemplates(templ).droidType <> DROID_CYBORG_REPAIR And _
           apsTemplates(templ).droidType <> DROID_PERSON Then
            If researchedTemplate(templ, selectedPlayer, includeRedundantDesigns) Then
                apsTemplateList_add templ
            End If
        End If
    Next
End Sub

' Add the design template form
Private Function intAddTemplateForm(ByVal psSelected As Long) As Boolean
    Dim parent As WIDGET, templbaseForm As WIDGET, templList As WIDGET

    Set parent = psWScreen.psForm

    ' add a form to place the tabbed form on
    Set templbaseForm = IntFormAnimated_new(False)
    parent.attach templbaseForm
    templbaseForm.id = IDDES_TEMPLBASE
    templbaseForm.setCalcLayout WzFunc_new(FN_intAddTemplateForm_templbaseForm_calcLayout)

    ' Add the obsolete items button.
    makeObsoleteButton templbaseForm

    ' Add the design templates form
    Set templList = IntListTabWidget_make()
    templbaseForm.attach templList
    templList.setCalcLayout WzFunc_new(FN_intAddTemplateForm_templList_calcLayout)

    ' Put the buttons on it
    intAddTemplateForm = intAddTemplateButtons(templList, psSelected)
End Function

' Add the droid template buttons to a form
Private Function intAddTemplateButtons(ByVal templList As WIDGET, ByVal psSelected As Long) As Boolean
    Dim TempString As String, nextButtonId As Long, sBarInit As W_BARINIT, k As Long, psTempl As Long, button As WIDGET

    ' Set up the button struct
    nextButtonId = IDDES_TEMPLSTART

    ' Add each button
    W_BARINIT_construct sBarInit
    sBarInit.id = IDDES_BARSTART
    sBarInit.x = STAT_TIMEBARX
    sBarInit.y = STAT_TIMEBARY
    sBarInit.width = STAT_PROGBARWIDTH
    sBarInit.height = STAT_PROGBARHEIGHT
    sBarInit.size = 50
    sBarInit.sCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MAJOR)
    sBarInit.sMinorCol = psPalette(WZCOL_ACTION_PROGRESS_BAR_MINOR)
    sBarInit.pTip = "Power Usage"

    droidTemplID = 0
    For k = 0 To apsTemplateList_size - 1
        psTempl = apsTemplateList(k)
        ' Set the tip and add the button
        Set button = IntStatsButton_new()
        templList.attach button
        button.id = nextButtonId
        IntStatsButton_setStatsAndTip button, STAT_TEMPLATE + psTempl
        ListTabWidget_addWidgetToLayout templList, button

        sBarInit.iRange = POWERPOINTS_DROIDDIV
        sBarInit.size = (calcTemplatePower(apsTemplates(psTempl)) \ POWERPOINTS_DROIDDIV) And &HFFFF&
        If sBarInit.size > WBAR_SCALE Then
            sBarInit.size = WBAR_SCALE
        End If

        TempString = Left$(va("%s - %d", "Power Usage", calcTemplatePower(apsTemplates(psTempl))), 255)
        sBarInit.pTip = TempString
        sBarInit.formID = nextButtonId
        If widgAddBarGraph(psWScreen, sBarInit) Is Nothing Then
            intAddTemplateButtons = False
            Exit Function
        End If

        ' if the current template matches psSelected lock the button
        If psTempl = psSelected Then
            droidTemplID = nextButtonId
            button.setState WBUT_LOCK
            ListTabWidget_setCurrentPage templList, ListTabWidget_pages(templList) - 1
        End If

        ' Update the init struct for the next button
        sBarInit.id = sBarInit.id + 1
        nextButtonId = nextButtonId + 1
        'check don't go over max templates that can fit on the form
        If nextButtonId >= IDDES_TEMPLEND Then
            Exit For
        End If
    Next

    intAddTemplateButtons = True
End Function

' Set the current mode of the design screen, and display the appropriate
' component lists
' added case IDES_TURRET_A,IDES_TURRET_B
Private Sub intSetDesignMode(ByVal newCompMode As DES_COMPMODE, Optional ByVal forceRefresh As Boolean = False)
    Dim weaponIndex As Long, compList As WIDGET

    If newCompMode = desCompMode And Not forceRefresh Then
        Exit Sub
    End If
    ' Have to change the component display - remove the old one
    If desCompMode <> IDES_NOCOMPONENT Then
        widgDelete psWScreen, IDDES_RIGHTBASE

        widgSetButtonState psWScreen, IDDES_BODYFORM, 0
        widgSetButtonState psWScreen, IDDES_PROPFORM, 0
        widgSetButtonState psWScreen, IDDES_SYSTEMFORM, 0
        widgHide psWScreen, IDDES_BODYFORM
        widgHide psWScreen, IDDES_PROPFORM
        widgHide psWScreen, IDDES_SYSTEMFORM

        widgSetButtonState psWScreen, IDDES_BODYBUTTON, 0
        widgSetButtonState psWScreen, IDDES_PROPBUTTON, 0
        widgSetButtonState psWScreen, IDDES_SYSTEMBUTTON, 0
        widgSetButtonState psWScreen, IDDES_WPABUTTON, 0
        widgSetButtonState psWScreen, IDDES_WPBBUTTON, 0
    End If

    ' Set up the display for the new mode
    desCompMode = newCompMode
    Select Case desCompMode
    Case IDES_NOCOMPONENT
        ' Nothing to display
    Case IDES_SYSTEM
        Set compList = intAddComponentForm()
        intAddExtraSystemButtons compList, _
                                 sCurrDesign.asParts(COMP_SENSOR), _
                                 sCurrDesign.asParts(COMP_ECM), _
                                 sCurrDesign.asParts(COMP_CONSTRUCT), _
                                 sCurrDesign.asParts(COMP_REPAIRUNIT), _
                                 sCurrDesign.asParts(COMP_BRAIN)
        intAddSystemButtons IDES_SYSTEM
        widgSetButtonState psWScreen, IDDES_SYSTEMFORM, WBUT_LOCK
        widgSetButtonState psWScreen, IDDES_SYSTEMBUTTON, WBUT_CLICKLOCK
        widgReveal psWScreen, IDDES_SYSTEMFORM
    Case IDES_TURRET
        Set compList = intAddComponentForm()
        If sCurrDesign.numWeaps > 0 Then weaponIndex = sCurrDesign.asWeaps(0) Else weaponIndex = 0
        intAddComponentButtons compList, weaponIterator(), weaponIndex, True
        intAddSystemButtons IDES_TURRET
        widgSetButtonState psWScreen, IDDES_SYSTEMFORM, WBUT_LOCK
        widgSetButtonState psWScreen, IDDES_SYSTEMBUTTON, WBUT_CLICKLOCK
        widgReveal psWScreen, IDDES_SYSTEMFORM
        intSetSystemForm STAT_WEAPON + DROID_TEMPLATE_getWeaponStats(sCurrDesign, 0) ' in case previous was a different slot
    Case IDES_BODY
        Set compList = intAddComponentForm()
        intAddComponentButtons compList, bodyIterator(), sCurrDesign.asParts(COMP_BODY), False
        widgSetButtonState psWScreen, IDDES_BODYFORM, WBUT_LOCK
        widgSetButtonState psWScreen, IDDES_BODYBUTTON, WBUT_CLICKLOCK
        widgReveal psWScreen, IDDES_BODYFORM
    Case IDES_PROPULSION
        Set compList = intAddComponentForm()
        intAddComponentButtons compList, propulsionIterator(), sCurrDesign.asParts(COMP_PROPULSION), False
        widgSetButtonState psWScreen, IDDES_PROPFORM, WBUT_LOCK
        widgSetButtonState psWScreen, IDDES_PROPBUTTON, WBUT_CLICKLOCK
        widgReveal psWScreen, IDDES_PROPFORM
        intSetPropulsionForm STAT_PROPULSION + DROID_TEMPLATE_getPropulsionStats(sCurrDesign)
    Case IDES_TURRET_A
        Set compList = intAddComponentForm()
        If sCurrDesign.numWeaps > 1 Then weaponIndex = sCurrDesign.asWeaps(1) Else weaponIndex = 0
        intAddComponentButtons compList, weaponIterator(), weaponIndex, True
        intAddSystemButtons IDES_TURRET_A
        widgSetButtonState psWScreen, IDDES_SYSTEMFORM, WBUT_LOCK
        widgSetButtonState psWScreen, IDDES_WPABUTTON, WBUT_CLICKLOCK
        widgReveal psWScreen, IDDES_SYSTEMFORM
        intSetSystemForm STAT_WEAPON + DROID_TEMPLATE_getWeaponStats(sCurrDesign, 1) ' in case previous was a different slot
        ' Stop the button flashing
        intSetButtonFlash IDDES_WPABUTTON, False
    Case IDES_TURRET_B
        Set compList = intAddComponentForm()
        If sCurrDesign.numWeaps > 2 Then weaponIndex = sCurrDesign.asWeaps(2) Else weaponIndex = 0
        intAddComponentButtons compList, weaponIterator(), weaponIndex, True
        intAddSystemButtons IDES_TURRET_B
        widgSetButtonState psWScreen, IDDES_SYSTEMFORM, WBUT_LOCK
        widgSetButtonState psWScreen, IDDES_WPBBUTTON, WBUT_CLICKLOCK
        widgReveal psWScreen, IDDES_SYSTEMFORM
        intSetSystemForm STAT_WEAPON + DROID_TEMPLATE_getWeaponStats(sCurrDesign, 2) ' in case previous was a different slot
        ' Stop the button flashing
        intSetButtonFlash IDDES_WPBBUTTON, False
    End Select
End Sub

' the COMPONENT_STATS * (a stats handle, 0: nullptr)
Private Function intChooseSystemStats(psTemplate As DROID_TEMPLATE) As Long
    Dim psStats As Long, compIndex As Long

    psStats = 0
    ' Choose correct system stats
    Select Case droidTemplateType(psTemplate)
    Case DROID_COMMAND
        compIndex = psTemplate.asParts(COMP_BRAIN)
        If Not (ULT(compIndex, asBrainStats_size)) Then ASSERT_FAILURE "compIndex < asBrainStats.size()", "src/design.cpp:1260", "intChooseSystemStats", "Invalid range referenced for numBrainStats, %d > %zu", compIndex, asBrainStats_size: intChooseSystemStats = 0: Exit Function
        psStats = STAT_BRAIN + compIndex
    Case DROID_SENSOR
        compIndex = psTemplate.asParts(COMP_SENSOR)
        If Not (ULT(compIndex, asSensorStats_size)) Then ASSERT_FAILURE "compIndex < asSensorStats.size()", "src/design.cpp:1265", "intChooseSystemStats", "Invalid range referenced for numSensorStats, %d > %zu", compIndex, asSensorStats_size: intChooseSystemStats = 0: Exit Function
        psStats = STAT_SENSOR + compIndex
    Case DROID_ECM
        compIndex = psTemplate.asParts(COMP_ECM)
        If Not (ULT(compIndex, asECMStats_size)) Then ASSERT_FAILURE "compIndex < asECMStats.size()", "src/design.cpp:1270", "intChooseSystemStats", "Invalid range referenced for numECMStats, %d > %zu", compIndex, asECMStats_size: intChooseSystemStats = 0: Exit Function
        psStats = STAT_ECM + compIndex
    Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
        compIndex = psTemplate.asParts(COMP_CONSTRUCT)
        If Not (ULT(compIndex, asConstructStats_size)) Then ASSERT_FAILURE "compIndex < asConstructStats.size()", "src/design.cpp:1276", "intChooseSystemStats", "Invalid range referenced for numConstructStats, %d > %zu", compIndex, asConstructStats_size: intChooseSystemStats = 0: Exit Function
        psStats = STAT_CONSTRUCT + compIndex
    Case DROID_REPAIR, DROID_CYBORG_REPAIR
        compIndex = psTemplate.asParts(COMP_REPAIRUNIT)
        If Not (ULT(compIndex, asRepairStats_size)) Then ASSERT_FAILURE "compIndex < asRepairStats.size()", "src/design.cpp:1282", "intChooseSystemStats", "Invalid range referenced for numRepairStats, %d > %zu", compIndex, asRepairStats_size: intChooseSystemStats = 0: Exit Function
        psStats = STAT_REPAIR + compIndex
    Case DROID_WEAPON, DROID_PERSON, DROID_CYBORG, DROID_CYBORG_SUPER, DROID_DEFAULT
        compIndex = psTemplate.asWeaps(0)
        If Not (ULT(compIndex, asWeaponStats_size)) Then ASSERT_FAILURE "compIndex < asWeaponStats.size()", "src/design.cpp:1291", "intChooseSystemStats", "Invalid range referenced for numWeaponStats, %d > %zu", compIndex, asWeaponStats_size: intChooseSystemStats = 0: Exit Function
        psStats = STAT_WEAPON + compIndex
    Case Else
        debug_ 1295, LOG_ERROR, "intChooseSystemStats", "unrecognised droid type"
        intChooseSystemStats = 0
        Exit Function
    End Select

    intChooseSystemStats = psStats
End Function

' Checks if concatenating two `char[]` will exceed `MAX_STR_LENGTH`. If the concatenated length exceeds `MAX_STR_LENGTH`, this function will log an error.
' @param string0 The `char[]` that `string1` will concatenate.
' @param string1 The `char[]` to concatenate.
Public Sub checkStringLength(ByVal string0 As String, ByVal string1 As String)
    If Len(string0) + Len(string1) > MAX_STR_LENGTH Then
        debug_ 1310, LOG_ERROR, "checkStringLength", "Name string too long %s+%s > %u", string0, string1, MAX_STR_LENGTH
        debug_ 1311, LOG_ERROR, "checkStringLength", "Please report what language you are using in the bug report!"
    End If
End Sub

Public Function GetDefaultTemplateName(psTemplate As DROID_TEMPLATE) As String
    ' NOTE:	At this time, savegames can support a max of 60. We are using MAX_STR_LENGTH (currently 256) for display
    ' We are also returning a truncated string, instead of NULL if the string is too long.
    Dim psStats As Long, compIndex As Long

    '
    '	First we check for the special cases of the Transporter & Cyborgs
    '
    If psTemplate.droidType = DROID_TRANSPORTER Then
        sstrcpy_aCurrName "Transport"
        GetDefaultTemplateName = aCurrName
        Exit Function
    End If
    If psTemplate.droidType = DROID_SUPERTRANSPORTER Then
        sstrcpy_aCurrName "Super Transport"
        GetDefaultTemplateName = aCurrName
        Exit Function
    End If

    ' For cyborgs, we don't need to add the body name nor the propulsion name. We can just use the template name.
    If psTemplate.droidType = DROID_CYBORG Or _
       psTemplate.droidType = DROID_CYBORG_CONSTRUCT Or _
       psTemplate.droidType = DROID_CYBORG_REPAIR Or _
       psTemplate.droidType = DROID_CYBORG_SUPER Then
        sstrcpy_aCurrName psTemplate.name_   ' _(name)
        GetDefaultTemplateName = aCurrName
        Exit Function
    End If

    '
    '	Now get the normal default droid name based on its components
    '
    aCurrName = "" ' Reset string to null
    psStats = intChooseSystemStats(psTemplate)
    If psTemplate.asWeaps(0) <> 0 Or _
       psTemplate.asParts(COMP_CONSTRUCT) <> 0 Or _
       psTemplate.asParts(COMP_SENSOR) <> 0 Or _
       psTemplate.asParts(COMP_ECM) <> 0 Or _
       psTemplate.asParts(COMP_REPAIRUNIT) <> 0 Or _
       psTemplate.asParts(COMP_BRAIN) <> 0 Then
        sstrcpy_aCurrName BASE_STATS_name(psStats)     ' getLocalizedStatsName(psStats)
        sstrcat_aCurrName " "
    End If

    If psTemplate.numWeaps > 1 Then
        sstrcat_aCurrName "Hydra "
    End If

    compIndex = psTemplate.asParts(COMP_BODY)
    If Not (ULT(compIndex, asBodyStats_size)) Then ASSERT_FAILURE "compIndex < asBodyStats.size()", "src/design.cpp:1369", "GetDefaultTemplateName", "Invalid range referenced for numBodyStats, %d > %zu", compIndex, asBodyStats_size: GetDefaultTemplateName = "": Exit Function
    psStats = STAT_BODY + compIndex
    If psTemplate.asParts(COMP_BODY) <> 0 Then
        checkStringLength aCurrName, BASE_STATS_name(psStats)
        sstrcat_aCurrName BASE_STATS_name(psStats)
        sstrcat_aCurrName " "
    End If

    compIndex = psTemplate.asParts(COMP_PROPULSION)
    If Not (ULT(compIndex, asPropulsionStats_size)) Then ASSERT_FAILURE "compIndex < asPropulsionStats.size()", "src/design.cpp:1379", "GetDefaultTemplateName", "Invalid range referenced for numPropulsionStats, %d > %zu", compIndex, asPropulsionStats_size: GetDefaultTemplateName = "": Exit Function
    psStats = STAT_PROPULSION + compIndex
    If psTemplate.asParts(COMP_PROPULSION) <> 0 Then
        checkStringLength aCurrName, BASE_STATS_name(psStats)
        sstrcat_aCurrName BASE_STATS_name(psStats)
    End If

    GetDefaultTemplateName = aCurrName
End Function

Private Sub intSetEditBoxTextFromTemplate(psTemplate As DROID_TEMPLATE)
    sstrcpy_aCurrName ""

    ' show component names if default template else show stat name
    If psTemplate.droidType <> DROID_DEFAULT Then
        sstrcpy_aCurrName psTemplate.name_     ' getStatsName(psTemplate)
    Else
        GetDefaultTemplateName psTemplate   ' sets aCurrName
    End If

    widgSetString psWScreen, IDDES_NAMEBOX, aCurrName
End Sub

' Set all the design bar graphs from a design template
Private Sub intSetDesignStats(psTemplate As DROID_TEMPLATE)
    Dim psStats As Long
    psStats = intChooseSystemStats(psTemplate)

    ' Set system stats
    intSetSystemForm psStats

    ' Set the body stats
    intSetBodyStats STAT_BODY + DROID_TEMPLATE_getBodyStats(psTemplate)

    ' Set the propulsion stats
    intSetPropulsionForm STAT_PROPULSION + DROID_TEMPLATE_getPropulsionStats(psTemplate)

    ' Set the name in the edit box
    intSetEditBoxTextFromTemplate psTemplate
End Sub

' Set up the system clickable form of the design screen given a set of stats
Private Function intSetSystemForm(ByVal psStats As Long) As Boolean
    Dim newSysMode As DES_SYSMODE, sFormInit As W_FORMINIT, systemForm As WIDGET, sBarInit As W_BARINIT, sLabInit As W_LABINIT

    newSysMode = 0

    ' Figure out what the new mode should be
    Select Case COMPONENT_STATS_compType(psStats)
    Case COMP_WEAPON
        newSysMode = IDES_WEAPON
    Case COMP_SENSOR
        newSysMode = IDES_SENSOR
    Case COMP_ECM
        newSysMode = IDES_ECM
    Case COMP_CONSTRUCT
        newSysMode = IDES_CONSTRUCT
    Case COMP_BRAIN
        newSysMode = IDES_COMMAND
    Case COMP_REPAIRUNIT
        newSysMode = IDES_REPAIR
    Case Else
        ASSERT_FAILURE "false", "src/design.cpp:1452", "intSetSystemForm", "Bad choice"
    End Select

    ' If the correct form is already displayed just set the stats
    If newSysMode = desSysMode Then
        intSetSystemStats psStats

        intSetSystemForm = True
        Exit Function
    End If

    ' Remove the old form if necessary
    If desSysMode <> IDES_NOSYSTEM Then
        widgDelete psWScreen, IDDES_SYSTEMFORM
    End If

    ' Set the new mode
    desSysMode = newSysMode

    ' Add the system form
    W_FORMINIT_construct sFormInit
    sFormInit.formID = IDDES_STATSFORM
    sFormInit.id = IDDES_SYSTEMFORM
    sFormInit.style = (WFORM_CLICKABLE Or WFORM_NOCLICKMOVE)
    sFormInit.x = DES_BARFORMX
    sFormInit.y = DES_BARFORMY
    sFormInit.width = DES_BARFORMWIDTH     'COMPBUTWIDTH;
    sFormInit.height = DES_BARFORMHEIGHT   'COMPBUTHEIGHT;
    sFormInit.pTip = BASE_STATS_name(psStats)  ' set form tip to stats string (getLocalizedStatsName)
    sFormInit.pUserData = psStats          ' store component stats
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayStatForm)
    Set systemForm = widgAddForm(psWScreen, sFormInit)
    If systemForm Is Nothing Then
        intSetSystemForm = False
        Exit Function
    End If

    ' Initialise the bargraph struct
    W_BARINIT_construct sBarInit
    sBarInit.formID = IDDES_SYSTEMFORM
    'sBarInit.style = WBAR_DOUBLE;
    sBarInit.x = DES_CLICKBARX
    sBarInit.y = DES_STATBAR_Y1    'DES_CLICKBARY;
    sBarInit.width = DES_CLICKBARWIDTH
    sBarInit.height = DES_CLICKBARHEIGHT
    sBarInit.sCol = PIELIGHT_withByte(sBarInit.sCol, 0, DES_CLICKBARMAJORRED)
    sBarInit.sCol = PIELIGHT_withByte(sBarInit.sCol, 1, DES_CLICKBARMAJORGREEN)
    sBarInit.sCol = PIELIGHT_withByte(sBarInit.sCol, 2, DES_CLICKBARMAJORBLUE)
    sBarInit.sMinorCol = PIELIGHT_withByte(sBarInit.sMinorCol, 0, DES_CLICKBARMINORRED)
    sBarInit.sMinorCol = PIELIGHT_withByte(sBarInit.sMinorCol, 1, DES_CLICKBARMINORGREEN)
    sBarInit.sMinorCol = PIELIGHT_withByte(sBarInit.sMinorCol, 2, DES_CLICKBARMINORBLUE)

    ' Initialise the label struct
    W_LABINIT_construct sLabInit
    sLabInit.formID = IDDES_SYSTEMFORM
    sLabInit.x = DES_CLICKBARNAMEX
    sLabInit.y = DES_CLICKBARY - DES_CLICKBARHEIGHT \ 3
    sLabInit.width = DES_CLICKBARNAMEWIDTH
    sLabInit.height = DES_CLICKBARHEIGHT
    Set sLabInit.pDisplay = WzFunc_new(FN_intDisplayImage)

    ' See what type of system stats we've got
    If BASE_STATS_hasType(psStats, STAT_SENSOR) Then
        ' Add the bar graphs
        sBarInit.id = IDDES_SENSORRANGE
        sBarInit.iRange = getDesignMaxSensorRange()
        sBarInit.pTip = "Sensor Range"
        sBarInit.denominator = TILE_UNITS
        sBarInit.precision = 1
        systemForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.denominator = 0
        sBarInit.precision = 0
        sBarInit.id = IDDES_SYSTEMSWEIGHT
        sBarInit.y = DES_STATBAR_Y2    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.iRange = findMaxWeight(extraSystemIterator())
        sBarInit.pTip = "Weight"
        systemForm.attach DesignStatsBar_makeLessIsBetter(sBarInit)

        ' Add the labels
        sLabInit.id = IDDES_SENSORRANGELAB
        sLabInit.pTip = "Sensor Range"
        sLabInit.UserData = IMAGE_DES_RANGE
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_SYSTEMSWEIGHTLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Weight"
        sLabInit.UserData = IMAGE_DES_WEIGHT
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
    ElseIf BASE_STATS_hasType(psStats, STAT_ECM) Then
        ' Add the bar graphs
        sBarInit.id = IDDES_ECMPOWER
        sBarInit.iRange = getDesignMaxEcmRange()
        sBarInit.pTip = "ECM Power"
        systemForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.id = IDDES_SYSTEMSWEIGHT
        sBarInit.y = DES_STATBAR_Y2    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.iRange = findMaxWeight(extraSystemIterator())
        sBarInit.pTip = "Weight"
        systemForm.attach DesignStatsBar_makeLessIsBetter(sBarInit)

        ' Add the labels
        sLabInit.id = IDDES_ECMPOWERLAB
        sLabInit.pTip = "ECM Power"
        sLabInit.UserData = IMAGE_DES_POWER
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_SYSTEMSWEIGHTLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Weight"
        sLabInit.UserData = IMAGE_DES_WEIGHT
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
    ElseIf BASE_STATS_hasType(psStats, STAT_CONSTRUCT) Then
        ' Add the bar graphs
        sBarInit.id = IDDES_CONSTPOINTS
        sBarInit.pTip = "Build Points"
        sBarInit.iRange = getDesignMaxBuildPoints()
        systemForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.id = IDDES_SYSTEMSWEIGHT
        sBarInit.y = DES_STATBAR_Y2    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.pTip = "Weight"
        sBarInit.iRange = findMaxWeight(extraSystemIterator())
        systemForm.attach DesignStatsBar_makeLessIsBetter(sBarInit)

        ' Add the labels
        sLabInit.id = IDDES_CONSTPOINTSLAB
        sLabInit.pTip = "Build Points"
        sLabInit.UserData = IMAGE_DES_BUILDRATE
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_SYSTEMSWEIGHTLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Weight"
        sLabInit.UserData = IMAGE_DES_WEIGHT
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
    ElseIf BASE_STATS_hasType(psStats, STAT_REPAIR) Then
        ' Add the bar graphs
        sBarInit.id = IDDES_REPAIRPOINTS
        sBarInit.pTip = "Build Points"
        sBarInit.iRange = findMax(repairIterator(), WzFunc_new(FN_intSetSystemForm_repairPoints))
        systemForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.id = IDDES_SYSTEMSWEIGHT
        sBarInit.y = DES_STATBAR_Y2    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.pTip = "Weight"
        sBarInit.iRange = findMaxWeight(extraSystemIterator())
        systemForm.attach DesignStatsBar_makeLessIsBetter(sBarInit)

        ' Add the labels
        sLabInit.id = IDDES_REPAIRPTLAB
        sLabInit.pTip = "Build Points"
        sLabInit.UserData = IMAGE_DES_BUILDRATE
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_REPAIRWGTLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Weight"
        sLabInit.UserData = IMAGE_DES_WEIGHT
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
    ElseIf BASE_STATS_hasType(psStats, STAT_WEAPON) Then
        ' Add the bar graphs
        sBarInit.id = IDDES_WEAPRANGE
        sBarInit.iRange = findMaxWeaponAttribute(WzFunc_new(FN_design_proj_GetLongRange))
        sBarInit.pTip = "Range"
        sBarInit.denominator = TILE_UNITS
        sBarInit.precision = 1
        systemForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.denominator = 1
        sBarInit.precision = 0
        sBarInit.id = IDDES_WEAPDAMAGE
        sBarInit.y = DES_STATBAR_Y2    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.iRange = findMaxWeaponAttribute(WzFunc_new(FN_design_weaponDamage))
        sBarInit.pTip = "Damage"
        systemForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.id = IDDES_WEAPROF
        sBarInit.y = DES_STATBAR_Y3    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.iRange = findMaxWeaponAttribute(WzFunc_new(FN_design_weaponROF))
        sBarInit.pTip = "Rate-of-Fire"
        systemForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.id = IDDES_SYSTEMSWEIGHT
        sBarInit.y = DES_STATBAR_Y4    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.iRange = findMaxWeight(weaponIterator())
        sBarInit.pTip = "Weight"
        systemForm.attach DesignStatsBar_makeLessIsBetter(sBarInit)

        ' Add the labels
        sLabInit.id = IDDES_WEAPRANGELAB
        sLabInit.pTip = "Range"
        sLabInit.UserData = IMAGE_DES_RANGE
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_WEAPDAMAGELAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Damage"
        sLabInit.UserData = IMAGE_DES_DAMAGE
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_WEAPROFLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Rate-of-Fire"
        sLabInit.UserData = IMAGE_DES_FIRERATE
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_SYSTEMSWEIGHTLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Weight"
        sLabInit.UserData = IMAGE_DES_WEIGHT
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetSystemForm = False
            Exit Function
        End If
    End If

    ' Add the correct component form
    Select Case desSysMode
    Case IDES_SENSOR, IDES_CONSTRUCT, IDES_ECM, IDES_REPAIR, IDES_COMMAND
        intSetDesignMode IDES_SYSTEM
    Case IDES_WEAPON
        intSetDesignMode IDES_TURRET
    Case Else
    End Select

    ' Set the stats
    intSetSystemStats psStats

    ' Lock the form down if necessary
    If desCompMode = IDES_SYSTEM Then
        widgSetButtonState psWScreen, IDDES_SYSTEMFORM, WBUT_LOCK
    End If

    intSetSystemForm = True
End Function


' Set up the propulsion clickable form of the design screen given a set of stats
Private Function intSetPropulsionForm(ByVal psStats As Long) As Boolean
    Dim newPropMode As DES_PROPMODE, sFormInit As W_FORMINIT, propulsionForm As WIDGET, sBarInit As W_BARINIT, sLabInit As W_LABINIT

    newPropMode = 0

    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:1745", "intSetPropulsionForm", "Invalid propulsion stats pointer": intSetPropulsionForm = False: Exit Function

    ' figure out what the new mode should be
    Select Case asPropulsionTypes(asPropulsionStats(psStats And &HFFFF&).propulsionType).travel
    Case GROUND
        newPropMode = IDES_GROUND
    Case AIR
        newPropMode = IDES_AIR
    End Select

    ' Remove the old form if necessary
    If desPropMode <> IDES_NOPROPULSION Then
        widgDelete psWScreen, IDDES_PROPFORM
    End If

    ' Set the new mode
    desPropMode = newPropMode

    ' Add the propulsion form
    W_FORMINIT_construct sFormInit
    sFormInit.formID = IDDES_STATSFORM
    sFormInit.id = IDDES_PROPFORM
    sFormInit.style = WFORM_CLICKABLE Or WFORM_NOCLICKMOVE
    sFormInit.x = DES_BARFORMX
    sFormInit.y = DES_BARFORMY
    sFormInit.width = DES_BARFORMWIDTH     'DES_COMPBUTWIDTH;
    sFormInit.height = DES_BARFORMHEIGHT   'DES_COMPBUTHEIGHT;
    sFormInit.pTip = BASE_STATS_name(psStats)  ' set form tip to stats string (getLocalizedStatsName)
    Set sFormInit.pDisplay = WzFunc_new(FN_intDisplayStatForm)
    Set propulsionForm = widgAddForm(psWScreen, sFormInit)
    If propulsionForm Is Nothing Then
        intSetPropulsionForm = False
        Exit Function
    End If

    ' Initialise the bargraph struct
    W_BARINIT_construct sBarInit
    sBarInit.formID = IDDES_PROPFORM
    'sBarInit.style = WBAR_DOUBLE;
    sBarInit.x = DES_CLICKBARX
    sBarInit.y = DES_STATBAR_Y1    'DES_CLICKBARY;
    sBarInit.width = DES_CLICKBARWIDTH
    sBarInit.height = DES_CLICKBARHEIGHT
    sBarInit.sCol = PIELIGHT_withByte(sBarInit.sCol, 0, DES_CLICKBARMAJORRED)
    sBarInit.sCol = PIELIGHT_withByte(sBarInit.sCol, 1, DES_CLICKBARMAJORGREEN)
    sBarInit.sCol = PIELIGHT_withByte(sBarInit.sCol, 2, DES_CLICKBARMAJORBLUE)
    sBarInit.sMinorCol = PIELIGHT_withByte(sBarInit.sMinorCol, 0, DES_CLICKBARMINORRED)
    sBarInit.sMinorCol = PIELIGHT_withByte(sBarInit.sMinorCol, 1, DES_CLICKBARMINORGREEN)
    sBarInit.sMinorCol = PIELIGHT_withByte(sBarInit.sMinorCol, 2, DES_CLICKBARMINORBLUE)

    ' Initialise the label struct
    W_LABINIT_construct sLabInit
    sLabInit.formID = IDDES_PROPFORM
    sLabInit.x = DES_CLICKBARNAMEX
    sLabInit.y = DES_CLICKBARY - DES_CLICKBARHEIGHT \ 3
    sLabInit.width = DES_CLICKBARNAMEWIDTH
    sLabInit.height = DES_CLICKBARNAMEHEIGHT   'DES_CLICKBARHEIGHT;
    Set sLabInit.pDisplay = WzFunc_new(FN_intDisplayImage)

    ' See what type of propulsion we've got
    Select Case desPropMode
    Case IDES_AIR
        ' Add the bar graphs
        sBarInit.id = IDDES_PROPAIR
        sBarInit.iRange = findMaxPropulsionSpeed(TER_ROAD)
        sBarInit.pTip = "Air Speed"
        sBarInit.denominator = TILE_UNITS
        sBarInit.precision = 2
        propulsionForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.denominator = 1
        sBarInit.precision = 0
        sBarInit.id = IDDES_PROPWEIGHT
        sBarInit.y = DES_STATBAR_Y2    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.iRange = findMax(propulsionIterator(), WzFunc_new(FN_calculatePropulsionWeight))
        sBarInit.pTip = "Weight"
        propulsionForm.attach DesignStatsBar_makeLessIsBetter(sBarInit)

        ' Add the labels
        sLabInit.id = IDDES_PROPAIRLAB
        sLabInit.pTip = "Air Speed"
        sLabInit.UserData = IMAGE_DES_HOVER
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetPropulsionForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_PROPWEIGHTLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Weight"
        sLabInit.UserData = IMAGE_DES_WEIGHT
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetPropulsionForm = False
            Exit Function
        End If
    Case IDES_GROUND
        ' Add the bar graphs
        sBarInit.id = IDDES_PROPROAD
        sBarInit.pTip = "Road Speed"
        sBarInit.iRange = findMaxPropulsionSpeed(TER_ROAD)
        sBarInit.denominator = TILE_UNITS
        sBarInit.precision = 2
        propulsionForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.id = IDDES_PROPCOUNTRY
        sBarInit.y = DES_STATBAR_Y2    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.pTip = "Off-Road Speed"
        sBarInit.iRange = findMaxPropulsionSpeed(TER_SANDYBRUSH)
        propulsionForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.id = IDDES_PROPWATER
        sBarInit.y = DES_STATBAR_Y3    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.pTip = "Water Speed"
        sBarInit.iRange = findMaxPropulsionSpeed(TER_WATER)
        propulsionForm.attach DesignStatsBar_new(sBarInit)

        sBarInit.denominator = 1
        sBarInit.precision = 0
        sBarInit.id = IDDES_PROPWEIGHT
        sBarInit.y = DES_STATBAR_Y4    '+= DES_CLICKBARHEIGHT + DES_CLICKGAP;
        sBarInit.pTip = "Weight"
        sBarInit.iRange = findMax(propulsionIterator(), WzFunc_new(FN_calculatePropulsionWeight))
        propulsionForm.attach DesignStatsBar_makeLessIsBetter(sBarInit)

        ' Add the labels
        sLabInit.id = IDDES_PROPROADLAB
        sLabInit.pTip = "Road Speed"
        sLabInit.UserData = IMAGE_DES_ROAD
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetPropulsionForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_PROPCOUNTRYLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Off-Road Speed"
        sLabInit.UserData = IMAGE_DES_CROSSCOUNTRY
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetPropulsionForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_PROPWATERLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Water Speed"
        sLabInit.UserData = IMAGE_DES_HOVER    'WATER;
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetPropulsionForm = False
            Exit Function
        End If
        sLabInit.id = IDDES_PROPWEIGHTLAB
        sLabInit.y = sLabInit.y + DES_CLICKBARHEIGHT + DES_CLICKGAP
        sLabInit.pTip = "Weight"
        sLabInit.UserData = IMAGE_DES_WEIGHT
        If widgAddLabel(psWScreen, sLabInit) Is Nothing Then
            intSetPropulsionForm = False
            Exit Function
        End If
    Case Else
    End Select

    ' Set the stats
    intSetPropulsionStats psStats

    ' Lock the form down if necessary
    If desCompMode = IDES_PROPULSION Then
        widgSetButtonState psWScreen, IDDES_PROPFORM, WBUT_LOCK
    End If

    intSetPropulsionForm = True
End Function

' Add the component tab form to the design screen
Private Function intAddComponentForm() As WIDGET
    Dim parent As WIDGET, rightBase As WIDGET, compList As WIDGET

    Set parent = psWScreen.psForm

    ' add a form to place the tabbed form on
    Set rightBase = IntFormAnimated_new(False)
    parent.attach rightBase
    rightBase.id = IDDES_RIGHTBASE
    rightBase.setCalcLayout WzFunc_new(FN_intAddComponentForm_rightBase_calcLayout)

    'now a single form
    Set compList = IntListTabWidget_make()
    rightBase.attach compList
    compList.setCalcLayout WzFunc_new(FN_intAddComponentForm_compList_calcLayout)
    Set intAddComponentForm = compList
End Function

' Add the system buttons (weapons, command droid, etc) to the design screen
Private Function intAddSystemButtons(ByVal mode As DES_COMPMODE) As Boolean
    Dim sButInit As W_BUTINIT

    ' add the weapon button
    W_BUTINIT_construct sButInit
    sButInit.formID = IDDES_RIGHTBASE
    sButInit.id = IDDES_WEAPONS
    sButInit.x = DES_WEAPONBUTTON_X
    sButInit.y = DES_SYSTEMBUTTON_Y
    sButInit.width = iV_GetImageWidth(IntImages, IMAGE_DES_WEAPONS)
    sButInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_WEAPONS)
    sButInit.pTip = "Weapons"
    Set sButInit.pDisplay = WzFunc_new(FN_intDisplayButtonHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_DES_EXTRAHI, IMAGE_DES_WEAPONS)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddSystemButtons = False
        Exit Function
    End If

    ' if currently got a VTOL proplusion attached then don't add the system buttons
    ' don't add the system button if mode is IDES_TURRET_A or IDES_TURRET_B
    If Not checkTemplateIsVtol(sCurrDesign) And mode <> IDES_TURRET_A And mode <> IDES_TURRET_B Then
        ' add the system button
        sButInit.formID = IDDES_RIGHTBASE
        sButInit.id = IDDES_SYSTEMS
        sButInit.x = DES_SYSTEMBUTTON_X
        sButInit.y = DES_SYSTEMBUTTON_Y
        sButInit.width = iV_GetImageWidth(IntImages, IMAGE_DES_SYSTEMS)
        sButInit.height = iV_GetImageHeight(IntImages, IMAGE_DES_SYSTEMS)
        sButInit.pTip = "Systems"
        Set sButInit.pDisplay = WzFunc_new(FN_intDisplayButtonHilight)
        sButInit.UserData = PACKDWORD_TRI(0, IMAGE_DES_EXTRAHI, IMAGE_DES_SYSTEMS)
        If widgAddButton(psWScreen, sButInit) Is Nothing Then
            intAddSystemButtons = False
            Exit Function
        End If
        If mode = IDES_SYSTEM Then
            widgSetButtonState psWScreen, IDDES_SYSTEMS, WBUT_LOCK
        End If
    End If

    ' lock down the correct button
    Select Case mode
    Case IDES_TURRET, IDES_TURRET_A, IDES_TURRET_B
        widgSetButtonState psWScreen, IDDES_WEAPONS, WBUT_LOCK
    Case IDES_SYSTEM
    Case Else
        ASSERT_FAILURE "!""invalid/unexpected mode""", "src/design.cpp:2007", "intAddSystemButtons", "unexpected mode"
    End Select

    intAddSystemButtons = True
End Function

' Add the component buttons to the main tab of the component form
Private Function intAddComponentButtons(ByVal compList As WIDGET, ByVal componentIterator_ As WzFunc, ByVal compID As Long, ByVal bWeapon As Boolean) As Boolean
    Dim bodysize As Long, nextButtonId As Long, maxComponents As Long, bVTOL As Boolean, psPropStats As Long, callback As WzFunc

    bodysize = SIZE_NUM

    ' Set up the button struct
    nextButtonId = IDDES_COMPSTART

    'need to set max number of buttons possible
    If bWeapon Then maxComponents = MAX_SYSTEM_COMPONENTS Else maxComponents = MAX_DESIGN_COMPONENTS

    'if adding weapons - need to check if the propulsion is a VTOL
    bVTOL = False

    If bWeapon Then
        'check if the current Template propulsion has been set
        If sCurrDesign.asParts(COMP_PROPULSION) <> 0 Then
            psPropStats = STAT_PROPULSION + DROID_TEMPLATE_getPropulsionStats(sCurrDesign)
            If Not (psPropStats <> 0) Then ASSERT_FAILURE "psPropStats != nullptr", "src/design.cpp:2034", "intAddComponentButtons", "invalid propulsion stats pointer": intAddComponentButtons = False: Exit Function

            bVTOL = bVTOL Or asPropulsionTypes(asPropulsionStats(psPropStats And &HFFFF&).propulsionType).travel = AIR
        End If
        If sCurrDesign.asParts(COMP_BODY) <> 0 Then
            bodysize = asBodyStats(DROID_TEMPLATE_getBodyStats(sCurrDesign)).size
        End If
    End If

    ' Add each button
    desCompID = 0
    numComponent = 0

    Set callback = WzFunc_new(FN_intAddComponentButtons_callback, compList, maxComponents, bWeapon, bVTOL, bodysize, nextButtonId, compID)
    WzFunc_call componentIterator_, callback

    widgSetBarRange psWScreen, IDDES_BODYPOINTS, findMax(componentIterator_, WzFunc_new(FN_calcShadowBodyPoints))
    widgSetBarRange psWScreen, IDDES_POWERBAR, findMax(componentIterator_, WzFunc_new(FN_calcShadowPower))

    intAddComponentButtons = True
End Function

' Add the component buttons to the main tab of the component form
Private Function intAddExtraSystemButtons(ByVal compList As WIDGET, ByVal sensorIndex As Long, ByVal ecmIndex As Long, ByVal constIndex As Long, ByVal repairIndex As Long, ByVal brainIndex As Long) As Boolean
    Dim compIndex As Long, nextButtonId As Long, buttonType As Long, componentIterator_ As WzFunc, callback As WzFunc

    compIndex = 0

    ' Set up the button struct
    nextButtonId = IDDES_EXTRASYSSTART

    ' Add the buttons :
    ' buttonType == 0  -  Sensor Buttons
    ' buttonType == 1  -  ECM Buttons
    ' buttonType == 2  -  Constructor Buttons
    ' buttonType == 3  -  Repair Buttons
    ' buttonType == 4  -  Brain Buttons
    numExtraSys = 0
    For buttonType = 0 To 4
        Set componentIterator_ = Nothing
        Select Case buttonType
        Case 0
            ' Sensor Buttons
            Set componentIterator_ = sensorIterator()
            compIndex = sensorIndex
        Case 1
            ' ECM Buttons
            Set componentIterator_ = ecmIterator()
            compIndex = ecmIndex
        Case 2
            ' Constructor Buttons
            Set componentIterator_ = constructorIterator()
            compIndex = constIndex
        Case 3
            ' Repair Buttons
            Set componentIterator_ = repairIterator()
            compIndex = repairIndex
        Case 4
            ' Brain Buttons
            Set componentIterator_ = brainIterator()
            compIndex = brainIndex
        End Select

        ' ([&]: nextButtonId is the callback's cap(1), read back after the loop)
        Set callback = WzFunc_new(FN_intAddExtraSystemButtons_callback, compList, nextButtonId, compIndex)
        WzFunc_call componentIterator_, callback
        nextButtonId = callback.cap(1)
    Next

    widgSetBarRange psWScreen, IDDES_BODYPOINTS, findMax(extraSystemIterator(), WzFunc_new(FN_calcShadowBodyPoints))
    widgSetBarRange psWScreen, IDDES_POWERBAR, findMax(extraSystemIterator(), WzFunc_new(FN_calcShadowPower))

    intAddExtraSystemButtons = True
End Function


' Set the bar graphs for the system clickable
Private Sub intSetSystemStats(ByVal psStats As Long)
    Dim psForm As WIDGET, tip As String

    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:2194", "intSetSystemStats", "Invalid stats pointer": Exit Sub

    ' set form tip to stats string
    If InStr(1, BASE_STATS_id(psStats), "ZNULL", vbBinaryCompare) <> 0 Then tip = "" Else tip = BASE_STATS_name(psStats)
    widgSetTip psWScreen, IDDES_SYSTEMFORM, tip

    ' set form stats for later display in intDisplayStatForm
    Set psForm = widgGetFromID(psWScreen, IDDES_SYSTEMFORM)
    If Not psForm Is Nothing Then
        psForm.pUserData = psStats
    End If

    ' Set the correct system stats
    Select Case COMPONENT_STATS_compType(psStats)
    Case COMP_SENSOR
        intSetSensorStats psStats
    Case COMP_ECM
        intSetECMStats psStats
    Case COMP_WEAPON
        intSetWeaponStats psStats
    Case COMP_CONSTRUCT
        intSetConstructStats psStats
    Case COMP_REPAIRUNIT
        intSetRepairStats psStats
    Case COMP_BRAIN
        ' ??? TBD FIXME
    Case Else
        ASSERT_FAILURE "false", "src/design.cpp:2228", "intSetSystemStats", "Bad choice"
    End Select
End Sub

' (VB) psStats && psStats->compType == compType ? psStats : nullptr
Private Function statsOfCompType(ByVal psStats As Long, ByVal compType As COMPONENT_TYPE) As Long
    If psStats <> 0 Then
        If COMPONENT_STATS_compType(psStats) = compType Then statsOfCompType = psStats
    End If
End Function

' Set the shadow bar graphs for the system clickable
Private Sub intSetSystemShadowStats(ByVal psStats As Long)
    Dim weight As Long
    Select Case desSysMode
    Case IDES_WEAPON
        intSetWeaponShadowStats statsOfCompType(psStats, COMP_WEAPON)
        Exit Sub
    Case IDES_SENSOR
        intSetSensorShadowStats statsOfCompType(psStats, COMP_SENSOR)
    Case IDES_ECM
        intSetECMShadowStats statsOfCompType(psStats, COMP_ECM)
    Case IDES_CONSTRUCT
        intSetConstructShadowStats statsOfCompType(psStats, COMP_CONSTRUCT)
    Case IDES_REPAIR
        intSetRepairShadowStats statsOfCompType(psStats, COMP_REPAIRUNIT)
    Case Else
        Exit Sub
    End Select

    If psStats <> 0 Then weight = COMPONENT_STATS_weight(psStats) Else weight = 0
    widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, weight
End Sub

' Set the bar graphs for the sensor stats
Private Sub intSetSensorStats(ByVal psStats As Long)
    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:2262", "intSetSensorStats", "Invalid stats pointer": Exit Sub
    If Not (BASE_STATS_hasType(psStats, STAT_SENSOR)) Then ASSERT_FAILURE "psStats->hasType(STAT_SENSOR)", "src/design.cpp:2263", "intSetSensorStats", "stats have wrong type": Exit Sub

    ' range
    widgSetBarSize psWScreen, IDDES_SENSORRANGE, sensorRange(asSensorStats(psStats And &HFFFF&), selectedPlayer)
    ' weight
    widgSetBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asSensorStats(psStats And &HFFFF&).weight
End Sub

' Set the shadow bar graphs for the sensor stats
Private Sub intSetSensorShadowStats(ByVal psStats As Long)
    If psStats <> 0 Then   ' (VB) psStats == nullptr || ..., in order
        If Not (BASE_STATS_hasType(psStats, STAT_SENSOR)) Then ASSERT_FAILURE "psStats == nullptr || psStats->hasType(STAT_SENSOR)", "src/design.cpp:2274", "intSetSensorShadowStats", "stats have wrong type"
    End If

    If psStats <> 0 Then
        ' range
        widgSetMinorBarSize psWScreen, IDDES_SENSORRANGE, _
                            sensorRange(asSensorStats(psStats And &HFFFF&), selectedPlayer And &HFF&)

        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asSensorStats(psStats And &HFFFF&).weight
    Else
        ' Remove the shadow bars
        widgSetMinorBarSize psWScreen, IDDES_SENSORRANGE, 0
        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, 0
    End If
End Sub


' Set the bar graphs for the ECM stats
Private Sub intSetECMStats(ByVal psStats As Long)
    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:2296", "intSetECMStats", "Invalid stats pointer": Exit Sub
    If Not (BASE_STATS_hasType(psStats, STAT_ECM)) Then ASSERT_FAILURE "psStats->hasType(STAT_ECM)", "src/design.cpp:2297", "intSetECMStats", "stats have wrong type": Exit Sub

    ' range
    widgSetBarSize psWScreen, IDDES_ECMPOWER, ecmRange(asECMStats(psStats And &HFFFF&), selectedPlayer)
    ' weight
    widgSetBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asECMStats(psStats And &HFFFF&).weight
End Sub

' Set the shadow bar graphs for the ECM stats
Private Sub intSetECMShadowStats(ByVal psStats As Long)
    If psStats <> 0 Then   ' (VB) psStats == nullptr || ..., in order
        If Not (BASE_STATS_hasType(psStats, STAT_ECM)) Then ASSERT_FAILURE "psStats == nullptr || psStats->hasType(STAT_ECM)", "src/design.cpp:2308", "intSetECMShadowStats", "stats have wrong type"
    End If

    If psStats <> 0 Then
        ' power
        widgSetMinorBarSize psWScreen, IDDES_ECMPOWER, ecmRange(asECMStats(psStats And &HFFFF&), selectedPlayer And &HFF&)
        ' weight
        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asECMStats(psStats And &HFFFF&).weight
    Else
        ' Remove the shadow bars
        widgSetMinorBarSize psWScreen, IDDES_ECMPOWER, 0
        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, 0
    End If
End Sub


' Set the bar graphs for the Constructor stats
Private Sub intSetConstructStats(ByVal psStats As Long)
    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:2329", "intSetConstructStats", "Invalid stats pointer": Exit Sub
    If Not (BASE_STATS_hasType(psStats, STAT_CONSTRUCT)) Then ASSERT_FAILURE "psStats->hasType(STAT_CONSTRUCT)", "src/design.cpp:2330", "intSetConstructStats", "stats have wrong type": Exit Sub

    ' power
    widgSetBarSize psWScreen, IDDES_CONSTPOINTS, _
                   constructorPoints(asConstructStats(psStats And &HFFFF&), selectedPlayer And &HFF&)
    ' weight
    widgSetBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asConstructStats(psStats And &HFFFF&).weight
End Sub


' Set the shadow bar graphs for the Constructor stats
Private Sub intSetConstructShadowStats(ByVal psStats As Long)
    If psStats <> 0 Then   ' (VB) psStats == nullptr || ..., in order
        If Not (BASE_STATS_hasType(psStats, STAT_CONSTRUCT)) Then ASSERT_FAILURE "psStats == nullptr || psStats->hasType(STAT_CONSTRUCT)", "src/design.cpp:2343", "intSetConstructShadowStats", "stats have wrong type"
    End If

    If psStats <> 0 Then
        ' power
        widgSetMinorBarSize psWScreen, IDDES_CONSTPOINTS, _
                            constructorPoints(asConstructStats(psStats And &HFFFF&), selectedPlayer And &HFF&)
        ' weight
        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asConstructStats(psStats And &HFFFF&).weight
    Else
        ' reset the shadow bars
        widgSetMinorBarSize psWScreen, IDDES_CONSTPOINTS, 0
        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, 0
    End If
End Sub

' Set the bar graphs for the Repair stats
Private Sub intSetRepairStats(ByVal psStats As Long)
    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:2364", "intSetRepairStats", "Invalid stats pointer": Exit Sub
    If Not (BASE_STATS_hasType(psStats, STAT_REPAIR)) Then ASSERT_FAILURE "psStats->hasType(STAT_REPAIR)", "src/design.cpp:2365", "intSetRepairStats", "stats have wrong type": Exit Sub

    ' power
    widgSetBarSize psWScreen, IDDES_REPAIRPOINTS, _
                   repairPoints(asRepairStats(psStats And &HFFFF&), selectedPlayer And &HFF&)
    ' weight
    widgSetBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asRepairStats(psStats And &HFFFF&).weight
End Sub


' Set the shadow bar graphs for the Repair stats
Private Sub intSetRepairShadowStats(ByVal psStats As Long)
    If psStats <> 0 Then   ' (VB) psStats == nullptr || ..., in order
        If Not (BASE_STATS_hasType(psStats, STAT_REPAIR)) Then ASSERT_FAILURE "psStats == nullptr || psStats->hasType(STAT_REPAIR)", "src/design.cpp:2378", "intSetRepairShadowStats", "stats have wrong type"
    End If

    If psStats <> 0 Then
        ' power
        widgSetMinorBarSize psWScreen, IDDES_REPAIRPOINTS, _
                            repairPoints(asRepairStats(psStats And &HFFFF&), selectedPlayer And &HFF&)
        ' weight
        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asRepairStats(psStats And &HFFFF&).weight
    Else
        ' reset the shadow bars
        widgSetMinorBarSize psWScreen, IDDES_REPAIRPOINTS, 0
        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, 0
    End If
End Sub


' Set the bar graphs for the Weapon stats
Private Sub intSetWeaponStats(ByVal psStats As Long)
    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:2400", "intSetWeaponStats", "Invalid stats pointer": Exit Sub
    If Not (BASE_STATS_hasType(psStats, STAT_WEAPON)) Then ASSERT_FAILURE "psStats->hasType(STAT_WEAPON)", "src/design.cpp:2401", "intSetWeaponStats", "stats have wrong type": Exit Sub

    ' range
    widgSetBarSize psWScreen, IDDES_WEAPRANGE, proj_GetLongRange(asWeaponStats(psStats And &HFFFF&), selectedPlayer)
    ' rate of fire
    widgSetBarSize psWScreen, IDDES_WEAPROF, weaponROF(asWeaponStats(psStats And &HFFFF&), selectedPlayer)
    ' damage
    widgSetBarSize psWScreen, IDDES_WEAPDAMAGE, weaponDamage(asWeaponStats(psStats And &HFFFF&), _
                   selectedPlayer And &HFF&) And &HFFFF&
    ' weight
    widgSetBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asWeaponStats(psStats And &HFFFF&).weight
End Sub

' Set the shadow bar graphs for the Weapon stats
Private Sub intSetWeaponShadowStats(ByVal psStats As Long)
    If psStats <> 0 Then   ' (VB) psStats == nullptr || ..., in order
        If Not (BASE_STATS_hasType(psStats, STAT_WEAPON)) Then ASSERT_FAILURE "psStats == nullptr || psStats->hasType(STAT_WEAPON)", "src/design.cpp:2417", "intSetWeaponShadowStats", "stats have wrong type"
    End If

    If psStats <> 0 Then
        ' range
        widgSetMinorBarSize psWScreen, IDDES_WEAPRANGE, proj_GetLongRange(asWeaponStats(psStats And &HFFFF&), selectedPlayer)
        ' rate of fire
        widgSetMinorBarSize psWScreen, IDDES_WEAPROF, weaponROF(asWeaponStats(psStats And &HFFFF&), selectedPlayer)
        ' damage
        widgSetMinorBarSize psWScreen, IDDES_WEAPDAMAGE, weaponDamage( _
                            asWeaponStats(psStats And &HFFFF&), selectedPlayer And &HFF&) And &HFFFF&
        ' weight
        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, asWeaponStats(psStats And &HFFFF&).weight
    Else
        ' Reset the shadow bars
        widgSetMinorBarSize psWScreen, IDDES_WEAPRANGE, 0
        widgSetMinorBarSize psWScreen, IDDES_WEAPROF, 0
        widgSetMinorBarSize psWScreen, IDDES_WEAPDAMAGE, 0
        widgSetMinorBarSize psWScreen, IDDES_SYSTEMSWEIGHT, 0
    End If
End Sub

' Set the bar graphs for the Body stats
Private Sub intSetBodyStats(ByVal psStats As Long)
    Dim psForm As WIDGET, tip As String

    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:2446", "intSetBodyStats", "Invalid stats pointer": Exit Sub
    If Not (BASE_STATS_hasType(psStats, STAT_BODY)) Then ASSERT_FAILURE "psStats->hasType(STAT_BODY)", "src/design.cpp:2447", "intSetBodyStats", "stats have wrong type": Exit Sub

    ' set form tip to stats string
    If InStr(1, BASE_STATS_id(psStats), "ZNULL", vbBinaryCompare) <> 0 Then tip = "" Else tip = BASE_STATS_name(psStats)
    widgSetTip psWScreen, IDDES_BODYFORM, tip

    ' armour
    'do kinetic armour
    widgSetBarSize psWScreen, IDDES_BODYARMOUR_K, bodyArmour(asBodyStats(psStats And &HFFFF&), selectedPlayer, WC_KINETIC)
    'do heat armour
    widgSetBarSize psWScreen, IDDES_BODYARMOUR_H, bodyArmour(asBodyStats(psStats And &HFFFF&), selectedPlayer, WC_HEAT)
    ' power
    widgSetBarSize psWScreen, IDDES_BODYPOWER, bodyPower(asBodyStats(psStats And &HFFFF&), selectedPlayer)
    ' weight
    widgSetBarSize psWScreen, IDDES_BODYWEIGHT, asBodyStats(psStats And &HFFFF&).weight

    ' set form stats for later display in intDisplayStatForm
    Set psForm = widgGetFromID(psWScreen, IDDES_BODYFORM)
    If Not psForm Is Nothing Then
        psForm.pUserData = psStats
    End If
End Sub

' Set the shadow bar graphs for the Body stats
Private Sub intSetBodyShadowStats(ByVal psStats As Long)
    If psStats <> 0 Then   ' (VB) psStats == nullptr || ..., in order
        If Not (BASE_STATS_hasType(psStats, STAT_BODY)) Then ASSERT_FAILURE "psStats == nullptr || psStats->hasType(STAT_BODY)", "src/design.cpp:2473", "intSetBodyShadowStats", "stats have wrong type"
    End If

    If psStats <> 0 Then
        ' armour - kinetic
        widgSetMinorBarSize psWScreen, IDDES_BODYARMOUR_K, bodyArmour(asBodyStats(psStats And &HFFFF&), selectedPlayer, WC_KINETIC)
        'armour - heat
        widgSetMinorBarSize psWScreen, IDDES_BODYARMOUR_H, bodyArmour(asBodyStats(psStats And &HFFFF&), selectedPlayer, WC_HEAT)
        ' power
        widgSetMinorBarSize psWScreen, IDDES_BODYPOWER, bodyPower(asBodyStats(psStats And &HFFFF&), selectedPlayer)
        ' weight
        widgSetMinorBarSize psWScreen, IDDES_BODYWEIGHT, asBodyStats(psStats And &HFFFF&).weight
    Else
        ' Reset the shadow bars
        widgSetMinorBarSize psWScreen, IDDES_BODYARMOUR_K, 0
        widgSetMinorBarSize psWScreen, IDDES_BODYARMOUR_H, 0
        widgSetMinorBarSize psWScreen, IDDES_BODYPOWER, 0
        widgSetMinorBarSize psWScreen, IDDES_BODYWEIGHT, 0
    End If
End Sub

' Sets the Design Power Bar for a given Template
Private Sub intSetDesignPower(psTemplate As DROID_TEMPLATE)
    ' use the same scale as PowerBar in main window so values are relative
    widgSetBarSize psWScreen, IDDES_POWERBAR, calcTemplatePower(psTemplate)
End Sub

' setTemplateStat's [&](int8_t newNumWeaps) clearWeapons
Private Sub setTemplateStat_clearWeapons(psTemplate As DROID_TEMPLATE, ByVal newNumWeaps As Long)
    Dim i As Long
    For i = newNumWeaps To MAX_WEAPONS - 1
        psTemplate.asWeaps(i) = 0
    Next
    psTemplate.numWeaps = MIN_(psTemplate.numWeaps, newNumWeaps)
End Sub

' setTemplateStat's [&] clearNonWeapons
Private Sub setTemplateStat_clearNonWeapons(psTemplate As DROID_TEMPLATE)
    ' Reset the sensor, ECM and constructor and repair
    ' - defaults will be set when OK is hit
    psTemplate.asParts(COMP_BRAIN) = 0
    psTemplate.asParts(COMP_REPAIRUNIT) = 0
    psTemplate.asParts(COMP_ECM) = 0
    psTemplate.asParts(COMP_SENSOR) = 0
    psTemplate.asParts(COMP_CONSTRUCT) = 0
End Sub

' setTemplateStat's [&] clearTurret
Private Sub setTemplateStat_clearTurret(psTemplate As DROID_TEMPLATE)
    setTemplateStat_clearNonWeapons psTemplate
    setTemplateStat_clearWeapons psTemplate, 0
End Sub

Private Sub setTemplateStat(psTemplate As DROID_TEMPLATE, ByVal psStats As Long)
    Dim i As Long, oldStats As Long

    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:2505", "setTemplateStat", "psStats not null": Exit Sub

    Select Case COMPONENT_STATS_compType(psStats)
    Case COMP_BODY
        ' auto stats = (BODY_STATS *)psStats
        psTemplate.asParts(COMP_BODY) = psStats And &HFFFF&
        If Not intCheckValidWeaponForProp(psTemplate) Then
            setTemplateStat_clearTurret psTemplate
        Else
            setTemplateStat_clearWeapons psTemplate, asBodyStats(psStats And &HFFFF&).weaponSlots
        End If
    Case COMP_BRAIN
        ' auto stats = (BRAIN_STATS *)psStats
        setTemplateStat_clearTurret psTemplate
        psTemplate.asParts(COMP_BRAIN) = psStats And &HFFFF&
        psTemplate.asWeaps(0) = asBrainStats(psStats And &HFFFF&).psWeaponStat   ' (stats->psWeaponStat - asWeaponStats.data(): the index)
        psTemplate.numWeaps = 1
    Case COMP_PROPULSION
        ' auto stats = (PROPULSION_STATS *)psStats
        oldStats = DROID_TEMPLATE_getPropulsionStats(psTemplate)
        If (asPropulsionStats(psStats And &HFFFF&).propulsionType = PROPULSION_TYPE_LIFT) <> (asPropulsionStats(oldStats).propulsionType = PROPULSION_TYPE_LIFT) Then
            setTemplateStat_clearTurret psTemplate
        End If
        psTemplate.asParts(COMP_PROPULSION) = psStats And &HFFFF&
    Case COMP_REPAIRUNIT
        setTemplateStat_clearTurret psTemplate
        psTemplate.asParts(COMP_REPAIRUNIT) = psStats And &HFFFF&
    Case COMP_ECM
        setTemplateStat_clearTurret psTemplate
        psTemplate.asParts(COMP_ECM) = psStats And &HFFFF&
    Case COMP_SENSOR
        setTemplateStat_clearTurret psTemplate
        psTemplate.asParts(COMP_SENSOR) = psStats And &HFFFF&
    Case COMP_CONSTRUCT
        setTemplateStat_clearTurret psTemplate
        psTemplate.asParts(COMP_CONSTRUCT) = psStats And &HFFFF&
    Case COMP_WEAPON
        setTemplateStat_clearNonWeapons psTemplate
        If desCompMode = IDES_TURRET_A Then
            i = 1
        ElseIf desCompMode = IDES_TURRET_B Then
            i = 2
        Else
            i = 0
        End If
        psTemplate.asWeaps(i) = psStats And &HFFFF&
        psTemplate.numWeaps = MAX_(psTemplate.numWeaps, i + 1)
    Case COMP_NUMCOMPONENTS
        ASSERT_FAILURE "false", "src/design.cpp:2587", "setTemplateStat", "COMP_NUMCOMPONENTS isn't a component type"
    End Select
End Sub

' Set the shadow bar graphs for the template power points - psStats is new hilited stats
Private Sub intSetTemplatePowerShadowStats(ByVal psStats As Long)
    If psStats = 0 Then
        ' Reset the shadow bar
        widgSetMinorBarSize psWScreen, IDDES_POWERBAR, 0
        Exit Sub
    End If

    widgSetMinorBarSize psWScreen, IDDES_POWERBAR, calcShadowPower(psStats)
End Sub

' Sets the Body Points Bar for a given Template
Private Sub intSetBodyPoints(psTemplate As DROID_TEMPLATE)
    ' If total greater than Body Bar size then scale values.
    widgSetBarSize psWScreen, IDDES_BODYPOINTS, calcTemplateBody(designTemplateHandle(psTemplate), selectedPlayer)
End Sub

' Set the shadow bar graphs for the template Body points - psStats is new hilited stats
Private Sub intSetTemplateBodyShadowStats(ByVal psStats As Long)
    If psStats = 0 Then
        ' Reset the shadow bar
        widgSetMinorBarSize psWScreen, IDDES_BODYPOINTS, 0
        Exit Sub
    End If

    widgSetMinorBarSize psWScreen, IDDES_BODYPOINTS, calcShadowBodyPoints(psStats)
End Sub


' Calculate the speed of a droid over a type of terrain
Private Function intCalcSpeed(ByVal type_ As Long, ByVal psProp As Long) As Long
    Dim psTempl As DROID_TEMPLATE, compIdx As Long, weight As Long, droidSpeed As Long

    If calcDroidWeight(sCurrDesign) = 0 Then
        intCalcSpeed = 0
        Exit Function
    End If
    psTempl = sCurrDesign
    compIdx = getCompFromID(COMP_PROPULSION, asPropulsionStats(psProp And &HFFFF&).id)
    If compIdx < 0 Then
        intCalcSpeed = 0
        Exit Function
    End If
    psTempl.asParts(COMP_PROPULSION) = compIdx
    weight = calcDroidWeight(psTempl)
    If weight = 0 Then
        intCalcSpeed = 0
        Exit Function
    End If
    'we want the design screen to show zero speed over water for all prop types except Hover and Vtol
    If type_ = TER_WATER Then
        If Not (asPropulsionStats(psProp And &HFFFF&).propulsionType = PROPULSION_TYPE_HOVER Or asPropulsionStats(psProp And &HFFFF&).propulsionType = PROPULSION_TYPE_LIFT) Then
            intCalcSpeed = 0
            Exit Function
        End If
    End If
    droidSpeed = calcDroidSpeed(calcDroidBaseSpeed(psTempl, weight, selectedPlayer), type_, psProp And &HFFFF&, 0)
    intCalcSpeed = droidSpeed
End Function


' Set the bar graphs for the Propulsion stats
Private Sub intSetPropulsionStats(ByVal psStats As Long)
    Dim psForm As WIDGET, tip As String

    If Not (psStats <> 0) Then ASSERT_FAILURE "psStats != nullptr", "src/design.cpp:2663", "intSetPropulsionStats", "Invalid stats pointer": Exit Sub
    If Not (BASE_STATS_hasType(psStats, STAT_PROPULSION)) Then ASSERT_FAILURE "psStats->hasType(STAT_PROPULSION)", "src/design.cpp:2664", "intSetPropulsionStats", "stats have wrong type": Exit Sub

    ' set form tip to stats string
    If InStr(1, BASE_STATS_id(psStats), "ZNULL", vbBinaryCompare) <> 0 Then tip = "" Else tip = BASE_STATS_name(psStats)
    widgSetTip psWScreen, IDDES_PROPFORM, tip

    ' set form stats for later display in intDisplayStatForm
    Set psForm = widgGetFromID(psWScreen, IDDES_PROPFORM)
    If Not psForm Is Nothing Then
        psForm.pUserData = psStats
    End If

    Select Case desPropMode
    Case IDES_GROUND
        ' Road speed
        widgSetBarSize psWScreen, IDDES_PROPROAD, intCalcSpeed(TER_ROAD, psStats)
        ' Cross country speed - grass
        widgSetBarSize psWScreen, IDDES_PROPCOUNTRY, intCalcSpeed(TER_SANDYBRUSH, psStats)
        ' Water speed
        widgSetBarSize psWScreen, IDDES_PROPWATER, intCalcSpeed(TER_WATER, psStats)
    Case IDES_AIR
        ' Air speed - terrain type doesn't matter, use road
        widgSetBarSize psWScreen, IDDES_PROPAIR, intCalcSpeed(TER_ROAD, psStats)
    Case Else
    End Select

    widgSetBarSize psWScreen, IDDES_PROPWEIGHT, calculatePropulsionWeight(psStats)
End Sub

Private Function calculatePropulsionWeight(ByVal propulsionStats As Long) As Long
    If sCurrDesign.asParts(COMP_BODY) = 0 Then
        calculatePropulsionWeight = 0
        Exit Function
    End If

    calculatePropulsionWeight = COMPONENT_STATS_weight(propulsionStats) * asBodyStats(DROID_TEMPLATE_getBodyStats(sCurrDesign)).weight \ 100
End Function

' Set the shadow bar graphs for the Propulsion stats
Private Sub intSetPropulsionShadowStats(ByVal psStats As Long)
    Dim travel As TRAVEL_MEDIUM, wrongType As Boolean

    If psStats <> 0 Then   ' (VB) psStats == nullptr || ..., in order
        If Not (BASE_STATS_hasType(psStats, STAT_PROPULSION)) Then ASSERT_FAILURE "psStats == nullptr || psStats->hasType(STAT_PROPULSION)", "src/design.cpp:2710", "intSetPropulsionShadowStats", "stats have wrong type"
    End If

    ' Only set the shadow stats if they are the right type
    If psStats <> 0 Then   ' (VB) psStats && (...), in order
        travel = asPropulsionTypes(asPropulsionStats(psStats And &HFFFF&).propulsionType).travel
        wrongType = (travel = GROUND And desPropMode = IDES_AIR) Or _
                    (travel = AIR And desPropMode = IDES_GROUND)
    End If
    If wrongType Then
        ' Reset the shadow bars. Prevent an incredibly trivial case where
        ' hovering over a valid propulsion and then to an invalid one to compare
        ' against (design is wheels, then hovered over half-tracks, then VTOL)
        ' causes the last shadow marker set to stay.
        If travel = GROUND And desPropMode = IDES_AIR Then
            widgSetMinorBarSize psWScreen, IDDES_PROPAIR, 0
        Else
            widgSetMinorBarSize psWScreen, IDDES_PROPROAD, 0
            widgSetMinorBarSize psWScreen, IDDES_PROPCOUNTRY, 0
            widgSetMinorBarSize psWScreen, IDDES_PROPWATER, 0
        End If
        If sCurrDesign.asParts(COMP_BODY) <> 0 Then
            widgSetMinorBarSize psWScreen, IDDES_PROPWEIGHT, calculatePropulsionWeight(psStats)
        End If
        Exit Sub
    End If

    Select Case desPropMode
    Case IDES_GROUND
        If psStats <> 0 Then
            ' Road speed
            widgSetMinorBarSize psWScreen, IDDES_PROPROAD, _
                                intCalcSpeed(TER_ROAD, psStats)
            ' Cross country speed - grass
            widgSetMinorBarSize psWScreen, IDDES_PROPCOUNTRY, _
                                intCalcSpeed(TER_SANDYBRUSH, psStats)
            ' Water speed
            widgSetMinorBarSize psWScreen, IDDES_PROPWATER, _
                                intCalcSpeed(TER_WATER, psStats)
        Else
            ' Reset the shadow bars
            widgSetMinorBarSize psWScreen, IDDES_PROPROAD, 0
            widgSetMinorBarSize psWScreen, IDDES_PROPCOUNTRY, 0
            widgSetMinorBarSize psWScreen, IDDES_PROPWATER, 0
        End If
    Case IDES_AIR
        If psStats <> 0 Then
            ' Air speed - terrain type doesn't matter, use ROAD
            widgSetMinorBarSize psWScreen, IDDES_PROPAIR, _
                                intCalcSpeed(TER_ROAD, psStats)
        Else
            ' Reset the shadow bar
            widgSetMinorBarSize psWScreen, IDDES_PROPAIR, 0
        End If
    Case Else
    End Select

    If psStats <> 0 Then
        widgSetMinorBarSize psWScreen, IDDES_PROPWEIGHT, calculatePropulsionWeight(psStats)
    Else
        ' Reset the shadow bar
        widgSetMinorBarSize psWScreen, IDDES_PROPWEIGHT, 0
    End If
End Sub

' checks whether the Template passed in is a Valid one; sets the ref, the droid type, the default system turrets and
' (newName not vbNullString: nullptr) the name
Public Function intValidTemplate(psTempl As DROID_TEMPLATE, ByVal newName As String, ByVal complain As Boolean, ByVal player As Long) As Boolean
    Dim level As code_part, bodysize As Long, isTransportTempl As Boolean, isVtolTemplate As Boolean
    Dim i As Long, weaponSize As Long, w As Long

    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/design.cpp:2797", "intValidTemplate", "Invalid player: %u", player: intValidTemplate = False: Exit Function

    If Not (ULT(psTempl.asParts(COMP_BODY), asBodyStats_size)) Then ASSERT_FAILURE "psTempl->asParts[COMP_BODY] < asBodyStats.size()", "src/design.cpp:2799", "intValidTemplate", "Invalid range referenced for numBodyStats, %d > %zu", psTempl.asParts(COMP_BODY), asBodyStats_size: intValidTemplate = False: Exit Function
    If Not (ULT(psTempl.asParts(COMP_BRAIN), asBrainStats_size)) Then ASSERT_FAILURE "psTempl->asParts[COMP_BRAIN] < asBrainStats.size()", "src/design.cpp:2800", "intValidTemplate", "Invalid range referenced for numBrainStats, %d > %zu", psTempl.asParts(COMP_BRAIN), asBrainStats_size: intValidTemplate = False: Exit Function
    If Not (ULT(psTempl.asParts(COMP_PROPULSION), asPropulsionStats_size)) Then ASSERT_FAILURE "psTempl->asParts[COMP_PROPULSION] < asPropulsionStats.size()", "src/design.cpp:2801", "intValidTemplate", "Invalid range referenced for numPropulsionStats, %d > %zu", psTempl.asParts(COMP_PROPULSION), asPropulsionStats_size: intValidTemplate = False: Exit Function
    If Not (ULT(psTempl.asParts(COMP_REPAIRUNIT), asRepairStats_size)) Then ASSERT_FAILURE "psTempl->asParts[COMP_REPAIRUNIT] < asRepairStats.size()", "src/design.cpp:2802", "intValidTemplate", "Invalid range referenced for numRepairStats, %d > %zu", psTempl.asParts(COMP_REPAIRUNIT), asRepairStats_size: intValidTemplate = False: Exit Function
    If Not (ULT(psTempl.asParts(COMP_ECM), asECMStats_size)) Then ASSERT_FAILURE "psTempl->asParts[COMP_ECM] < asECMStats.size()", "src/design.cpp:2803", "intValidTemplate", "Invalid range referenced for numECMStats, %d > %zu", psTempl.asParts(COMP_ECM), asECMStats_size: intValidTemplate = False: Exit Function
    If Not (ULT(psTempl.asParts(COMP_SENSOR), asSensorStats_size)) Then ASSERT_FAILURE "psTempl->asParts[COMP_SENSOR] < asSensorStats.size()", "src/design.cpp:2804", "intValidTemplate", "Invalid range referenced for numSensorStats, %d > %zu", psTempl.asParts(COMP_SENSOR), asSensorStats_size: intValidTemplate = False: Exit Function
    If Not (ULT(psTempl.asParts(COMP_CONSTRUCT), asConstructStats_size)) Then ASSERT_FAILURE "psTempl->asParts[COMP_CONSTRUCT] < asConstructStats.size()", "src/design.cpp:2805", "intValidTemplate", "Invalid range referenced for numConstructStats, %d > %zu", psTempl.asParts(COMP_CONSTRUCT), asConstructStats_size: intValidTemplate = False: Exit Function

    If complain Then level = LOG_ERROR Else level = LOG_NEVER
    bodysize = asBodyStats(DROID_TEMPLATE_getBodyStats(psTempl)).size

    ' set the weapon for a command droid
    If psTempl.asParts(COMP_BRAIN) <> 0 Then
        psTempl.numWeaps = 1
        psTempl.asWeaps(0) = asBrainStats(DROID_TEMPLATE_getBrainStats(psTempl)).psWeaponStat   ' (psWeaponStat - asWeaponStats.data(): the index)
    End If

    ' Check all the components have been set
    If psTempl.asParts(COMP_BODY) = 0 Then
        debug_ 2820, level, "intValidTemplate", "No body given for template"
        intValidTemplate = False
        Exit Function
    ElseIf psTempl.asParts(COMP_PROPULSION) = 0 Then
        debug_ 2825, level, "intValidTemplate", "No propulsion given for template"
        intValidTemplate = False
        Exit Function
    End If

    isTransportTempl = isTransporter_template(psTempl)

    ' Check a turret has been installed
    If psTempl.numWeaps = 0 And _
       psTempl.asParts(COMP_SENSOR) = 0 And _
       psTempl.asParts(COMP_ECM) = 0 And _
       psTempl.asParts(COMP_BRAIN) = 0 And _
       psTempl.asParts(COMP_REPAIRUNIT) = 0 And _
       psTempl.asParts(COMP_CONSTRUCT) = 0 And _
       Not isTransportTempl Then
        debug_ 2840, level, "intValidTemplate", "No turret for template"
        intValidTemplate = False
        Exit Function
    End If

    isVtolTemplate = checkTemplateIsVtol(psTempl)  ' True for transporters also.

    ' Check the weapons
    For i = 0 To psTempl.numWeaps - 1
        If Not (ULT(psTempl.asWeaps(i), asWeaponStats_size)) Then ASSERT_FAILURE "psTempl->asWeaps[i] < asWeaponStats.size()", "src/design.cpp:2849", "intValidTemplate", "Invalid range referenced for numWeaponStats, %d > %zu", psTempl.asWeaps(i), asWeaponStats_size: intValidTemplate = False: Exit Function

        w = DROID_TEMPLATE_getWeaponStats(psTempl, i)
        weaponSize = asWeaponStats(w).weaponSize

        If (weaponSize = WEAPON_SIZE_LIGHT And bodysize <> SIZE_LIGHT) _
           Or (weaponSize = WEAPON_SIZE_HEAVY And bodysize = SIZE_LIGHT) _
           Or psTempl.asWeaps(i) = 0 Then
            debug_ 2857, level, "intValidTemplate", "No weapon given for weapon droid, or wrong weapon size"
            intValidTemplate = False
            Exit Function
        End If
        If isTransportTempl And (asWeaponStats(w).flags And BIT_(WEAPON_FLAG_ALLOWED_ON_TRANSPORTER)) = 0 Then
            debug_ 2862, level, "intValidTemplate", "Transporter with invalid turret, not possible"
            intValidTemplate = False
            Exit Function
        End If
        If isVtolTemplate Then
            If asWeaponStats(w).vtolAttackRuns <= 0 Then
                debug_ 2869, level, "intValidTemplate", "VTOL with non-VTOL turret, not possible"
                intValidTemplate = False
                Exit Function
            End If
        ElseIf asWeaponStats(w).vtolAttackRuns > 0 Then
            debug_ 2875, level, "intValidTemplate", "Non-VTOL with VTOL turret, not possible"
            intValidTemplate = False
            Exit Function
        End If
    Next

    ' Check number of weapon slots
    If ULT(asBodyStats(DROID_TEMPLATE_getBodyStats(psTempl)).weaponSlots, psTempl.numWeaps) Then
        debug_ 2883, level, "intValidTemplate", "Too many weapon turrets"
        intValidTemplate = False
        Exit Function
    End If

    ' Check no mixing of systems and weapons
    If psTempl.numWeaps <> 0 And _
       ((psTempl.asParts(COMP_SENSOR) <> 0 And psTempl.asParts(COMP_SENSOR) <> aDefaultSensor(player)) Or _
        (psTempl.asParts(COMP_ECM) <> 0 And psTempl.asParts(COMP_ECM) <> aDefaultECM(player)) Or _
        (psTempl.asParts(COMP_REPAIRUNIT) <> 0 And psTempl.asParts(COMP_REPAIRUNIT) <> aDefaultRepair(player)) Or _
        psTempl.asParts(COMP_CONSTRUCT) <> 0) Then
        debug_ 2894, level, "intValidTemplate", "Cannot mix system and weapon turrets in a template!"
        intValidTemplate = False
        Exit Function
    End If
    If psTempl.numWeaps <> 1 And psTempl.asParts(COMP_BRAIN) <> 0 Then
        debug_ 2899, level, "intValidTemplate", "Commander template needs 1 weapon turret"
        intValidTemplate = False
        Exit Function
    End If

    'can only have a VTOL weapon on a VTOL propulsion
    If isVtolTemplate And Not isTransportTempl And psTempl.numWeaps = 0 Then
        debug_ 2906, level, "intValidTemplate", "VTOL with system turret, not possible"
        intValidTemplate = False
        Exit Function
    End If

    If psTempl.asParts(COMP_SENSOR) = 0 Then
        ' Set the default Sensor
        psTempl.asParts(COMP_SENSOR) = aDefaultSensor(player)
    End If

    If psTempl.asParts(COMP_ECM) = 0 Then
        ' Set the default ECM
        psTempl.asParts(COMP_ECM) = aDefaultECM(player)
    End If

    If psTempl.asParts(COMP_REPAIRUNIT) = 0 Then
        ' Set the default Repair
        psTempl.asParts(COMP_REPAIRUNIT) = aDefaultRepair(player)
    End If

    psTempl.ref = STAT_TEMPLATE

    'set the droidtype
    psTempl.droidType = droidTemplateType(psTempl)

    ' copy name into template
    If StrPtr(newName) <> 0 Then
        psTempl.name_ = newName         ' WzString::fromUtf8(newName)
    End If

    If Not designableTemplate(psTempl, player) Then
        intValidTemplate = False
        Exit Function
    End If

    psTempl.enabled = True

    intValidTemplate = True
End Function

Private Sub desCreateDefaultTemplate()
    ' set current design to default
    sCurrDesign = apsTemplates(sDefaultDesignTemplate)
    sCurrDesign.stored = False

    ' reset stats
    intSetDesignStats sCurrDesign
    widgDelete psWScreen, IDDES_SYSTEMFORM
    desSysMode = IDES_NOSYSTEM
End Sub

' Remove the design widgets from the widget screen
Public Sub intRemoveDesign()
    widgDelete psWScreen, IDDES_POWERFORM
    widgDelete psWScreen, IDDES_NAMEBOX
    widgDelete psWScreen, IDDES_TEMPLBASE
    widgDelete psWScreen, IDDES_RIGHTBASE

    widgDelete psWScreen, IDDES_BODYFORM
    widgDelete psWScreen, IDDES_PROPFORM
    widgDelete psWScreen, IDDES_SYSTEMFORM

    widgDelete psWScreen, IDDES_FORM
    widgDelete psWScreen, IDDES_STATSFORM

    resetDesignPauseState
End Sub

' set flashing flag for button
Private Sub intSetButtonFlash(ByVal id As Long, ByVal bFlash As Boolean)
    Dim psWidget As WIDGET
    Set psWidget = widgGetFromID(psWScreen, id)

    If Not (psWidget.typ = WIDG_BUTTON) Then ASSERT_FAILURE "psWidget->type == WIDG_BUTTON", "src/design.cpp:2984", "intSetButtonFlash", "Not a button": Exit Sub

    If bFlash Then
        Set psWidget.displayFunction = WzFunc_new(FN_intDisplayButtonFlash)
    Else
        Set psWidget.displayFunction = WzFunc_new(FN_intDisplayButtonHilight)
    End If
End Sub

'
' desTemplateNameCustomised
'
' Checks whether user has customised template name : template not
' customised if not complete or if generated name same as current.
'
Private Function desTemplateNameCustomised(psTemplate As DROID_TEMPLATE) As Boolean
    desTemplateNameCustomised = False
    If psTemplate.droidType <> DROID_DEFAULT Then  ' (VB) && in order
        desTemplateNameCustomised = StrComp(psTemplate.name_, GetDefaultTemplateName(psTemplate), vbBinaryCompare) <> 0
    End If
End Function

Private Function templateFromButtonId(ByVal buttonId As Long, Optional ByVal allowBlankTemplate As Boolean = False) As Long
    Dim minIndex As Long, index As Long
    If allowBlankTemplate Then minIndex = 0 Else minIndex = 1
    index = buttonId - IDDES_TEMPLSTART     ' unsigned

    If Not ULT(index, minIndex) And ULT(index, apsTemplateList_size) Then
        templateFromButtonId = apsTemplateList(index)
        Exit Function
    End If
    templateFromButtonId = 0
End Function

' Process return codes from the design screen
Public Sub intProcessDesign(ByVal id As Long)
    Dim psTempl As Long, bTemplateNameCustomised As Boolean, numWeaps As Long, maxWeaps As Long, k As Long, state As Long
    Dim obsoleteButton As WIDGET, choice As Long

    ' check template button pressed
    If id >= IDDES_TEMPLSTART And id <= IDDES_TEMPLEND Then
        ' if first template create blank design
        If id = IDDES_TEMPLSTART Then
            desCreateDefaultTemplate

            aCurrName = ""
            sCurrDesign.name_ = ""

            ' reveal body button
            widgReveal psWScreen, IDDES_BODYBUTTON
            ' hide other component buttons
            widgHide psWScreen, IDDES_SYSTEMBUTTON
            widgHide psWScreen, IDDES_PROPBUTTON
            widgHide psWScreen, IDDES_WPABUTTON
            widgHide psWScreen, IDDES_WPBBUTTON

            ' set button render routines to flash
            intSetButtonFlash IDDES_BODYBUTTON, True
            intSetButtonFlash IDDES_SYSTEMBUTTON, True
            intSetButtonFlash IDDES_PROPBUTTON, True
            intSetButtonFlash IDDES_WPABUTTON, True
            intSetButtonFlash IDDES_WPBBUTTON, True

            If bMultiPlayer Then
                widgHide psWScreen, IDDES_STOREBUTTON
            End If
        Else
            ' Find the template for the new button
            psTempl = templateFromButtonId(id, True)

            If Not (psTempl <> 0) Then ASSERT_FAILURE "psTempl != nullptr", "src/design.cpp:3052", "intProcessDesign", "template not found!": Exit Sub

            If psTempl <> 0 Then
                ' Set the new template
                sCurrDesign = apsTemplates(psTempl)
                sstrcpy_aCurrName apsTemplates(psTempl).name_      ' getStatsName(psTempl)

                ' reveal body/propulsion/turret component buttons
                widgReveal psWScreen, IDDES_BODYBUTTON
                widgReveal psWScreen, IDDES_PROPBUTTON
                widgReveal psWScreen, IDDES_SYSTEMBUTTON
                ' hide extra turrets
                widgHide psWScreen, IDDES_WPABUTTON
                widgHide psWScreen, IDDES_WPBBUTTON

                ' turn off button flashes
                intSetButtonFlash IDDES_BODYBUTTON, False
                intSetButtonFlash IDDES_SYSTEMBUTTON, False
                intSetButtonFlash IDDES_PROPBUTTON, False
                intSetButtonFlash IDDES_WPABUTTON, False
                intSetButtonFlash IDDES_WPBBUTTON, False

                ' reveal additional buttons
                If apsTemplates(psTempl).numWeaps >= 2 Then
                    widgReveal psWScreen, IDDES_WPABUTTON
                Else
                    intSetButtonFlash IDDES_WPABUTTON, True
                End If
                If apsTemplates(psTempl).numWeaps = 3 Then
                    widgReveal psWScreen, IDDES_WPBBUTTON
                Else
                    intSetButtonFlash IDDES_WPBBUTTON, True
                End If

                If bMultiPlayer Then
                    widgReveal psWScreen, IDDES_STOREBUTTON
                    updateStoreButton sCurrDesign.stored
                End If
            End If
        End If

        ' reveal design form if not already on-screen
        widgReveal psWScreen, IDDES_FORM

        ' Droid template button has been pressed - clear the old button
        If droidTemplID <> 0 Then
            widgSetButtonState psWScreen, droidTemplID, 0
        End If

        intSetDesignStats sCurrDesign

        ' show body stats only
        widgReveal psWScreen, IDDES_STATSFORM
        widgReveal psWScreen, IDDES_BODYFORM
        widgHide psWScreen, IDDES_PROPFORM
        widgHide psWScreen, IDDES_SYSTEMFORM

        'Update the Power bar stats as the power to build will have changed
        intSetDesignPower sCurrDesign
        'Update the body points
        intSetBodyPoints sCurrDesign

        ' Lock the button
        widgSetButtonState psWScreen, id, WBUT_LOCK
        droidTemplID = id

        ' Update the component form
        widgDelete psWScreen, IDDES_RIGHTBASE
        ' reset button states
        widgSetButtonState psWScreen, IDDES_SYSTEMBUTTON, 0
        widgSetButtonState psWScreen, IDDES_BODYBUTTON, 0
        widgSetButtonState psWScreen, IDDES_PROPBUTTON, 0
        widgSetButtonState psWScreen, IDDES_WPABUTTON, 0
        widgSetButtonState psWScreen, IDDES_WPBBUTTON, 0
        desCompMode = IDES_NOCOMPONENT
        intSetDesignMode IDES_BODY
    ElseIf id >= IDDES_COMPSTART And id <= IDDES_COMPEND Then
        ' check whether can change template name
        bTemplateNameCustomised = desTemplateNameCustomised(sCurrDesign)

        ' Component stats button has been pressed - clear the old button
        If desCompID <> 0 Then
            widgSetButtonState psWScreen, desCompID, 0
        End If

        ' Set the stats in the template
        Select Case desCompMode
        Case IDES_SYSTEM
            '0 weapon for utility droid
            sCurrDesign.numWeaps = 0
        Case IDES_TURRET
            setTemplateStat sCurrDesign, apsComponentList(id - IDDES_COMPSTART)
            'Watermelon:weaponslots >= 2
            If asBodyStats(DROID_TEMPLATE_getBodyStats(sCurrDesign)).weaponSlots >= 2 Then
                ' reveal turret_a button if hidden
                widgReveal psWScreen, IDDES_WPABUTTON
            End If
            ' Set the new stats on the display
            intSetSystemForm apsComponentList(id - IDDES_COMPSTART)
            ' Stop the button flashing
            intSetButtonFlash IDDES_SYSTEMBUTTON, False

            triggerEvent TRIGGER_DESIGN_WEAPON
        'Added cases for 2nd/3rd turret
        Case IDES_TURRET_A
            setTemplateStat sCurrDesign, apsComponentList(id - IDDES_COMPSTART)
            'Watermelon:weaponSlots > 2
            If asBodyStats(DROID_TEMPLATE_getBodyStats(sCurrDesign)).weaponSlots > 2 Then
                ' reveal turret_b button if hidden
                widgReveal psWScreen, IDDES_WPBBUTTON
            End If
            ' Set the new stats on the display
            intSetSystemForm apsComponentList(id - IDDES_COMPSTART)
            ' Stop the button flashing
            intSetButtonFlash IDDES_WPABUTTON, False

            triggerEvent TRIGGER_DESIGN_WEAPON
        Case IDES_TURRET_B
            setTemplateStat sCurrDesign, apsComponentList(id - IDDES_COMPSTART)
            ' Set the new stats on the display
            intSetSystemForm apsComponentList(id - IDDES_COMPSTART)
            ' Stop the button flashing
            intSetButtonFlash IDDES_WPBBUTTON, False

            triggerEvent TRIGGER_DESIGN_WEAPON
        Case IDES_BODY
            ' reveal propulsion button if hidden
            widgReveal psWScreen, IDDES_PROPBUTTON

            setTemplateStat sCurrDesign, apsComponentList(id - IDDES_COMPSTART)
            ' Set the new stats on the display
            intSetBodyStats apsComponentList(id - IDDES_COMPSTART)

            If sCurrDesign.asParts(COMP_BRAIN) <> 0 Then numWeaps = 0 Else numWeaps = sCurrDesign.numWeaps
            maxWeaps = asBodyStats(DROID_TEMPLATE_getBodyStats(sCurrDesign)).weaponSlots
            widgGetFromID(psWScreen, IDDES_WPABUTTON).show maxWeaps > 1 And numWeaps >= 1
            widgGetFromID(psWScreen, IDDES_WPBBUTTON).show maxWeaps > 2 And numWeaps >= 2
            If maxWeaps > 1 And numWeaps = 1 Then state = WBUT_FLASH Else state = 0
            widgSetButtonState psWScreen, IDDES_WPABUTTON, state
            If maxWeaps > 2 And numWeaps = 2 Then state = WBUT_FLASH Else state = 0
            widgSetButtonState psWScreen, IDDES_WPBBUTTON, state
            intSetButtonFlash IDDES_WPABUTTON, maxWeaps > 1 And numWeaps = 1
            intSetButtonFlash IDDES_WPBBUTTON, maxWeaps > 2 And numWeaps = 2
            ' Stop the button flashing
            intSetButtonFlash IDDES_BODYBUTTON, False

            triggerEvent TRIGGER_DESIGN_BODY
        Case IDES_PROPULSION
            setTemplateStat sCurrDesign, apsComponentList(id - IDDES_COMPSTART)

            ' Set the new stats on the display
            intSetPropulsionForm apsComponentList(id - IDDES_COMPSTART)

            ' Check that the weapon (if any) is valid for this propulsion
            If Not intCheckValidWeaponForProp(sCurrDesign) Then
                ' Not valid weapon so initialise the weapon stat
                widgHide psWScreen, IDDES_WPABUTTON
                widgHide psWScreen, IDDES_WPBBUTTON

                ' We need a turret again
                intSetButtonFlash IDDES_SYSTEMBUTTON, True
            End If

            ' Stop the button flashing
            intSetButtonFlash IDDES_PROPBUTTON, False

            triggerEvent TRIGGER_DESIGN_PROPULSION
        Case Else
        End Select

        ' Lock the new button
        widgSetButtonState psWScreen, id, WBUT_LOCK
        desCompID = id

        ' Update the propulsion stats as the droid weight will have changed
        intSetPropulsionStats STAT_PROPULSION + DROID_TEMPLATE_getPropulsionStats(sCurrDesign)

        'Update the Power bar stats as the power to build will have changed
        intSetDesignPower sCurrDesign
        'Update the body points
        intSetBodyPoints sCurrDesign

        ' update name if not customised
        If Not bTemplateNameCustomised Then
            sCurrDesign.name_ = GetDefaultTemplateName(sCurrDesign)
        End If

        ' Update the name in the edit box
        intSetEditBoxTextFromTemplate sCurrDesign
    ElseIf id >= IDDES_EXTRASYSSTART And id <= IDDES_EXTRASYSEND Then
        ' check whether can change template name
        bTemplateNameCustomised = desTemplateNameCustomised(sCurrDesign)

        ' Extra component stats button has been pressed - clear the old button
        If desCompID <> 0 Then
            widgSetButtonState psWScreen, desCompID, 0
        End If

        ' Now store the new stats
        setTemplateStat sCurrDesign, apsExtraSysList(id - IDDES_EXTRASYSSTART)
        widgHide psWScreen, IDDES_WPABUTTON
        widgHide psWScreen, IDDES_WPBBUTTON
        ' Set the new stats on the display
        intSetSystemForm apsExtraSysList(id - IDDES_EXTRASYSSTART)
        ' Stop the button flashing
        intSetButtonFlash IDDES_SYSTEMBUTTON, False
        ' Lock the new button
        widgSetButtonState psWScreen, id, WBUT_LOCK
        desCompID = id

        ' Update the propulsion stats as the droid weight will have changed
        intSetPropulsionStats STAT_PROPULSION + DROID_TEMPLATE_getPropulsionStats(sCurrDesign)

        ' Update the Power bar stats as the power to build will have changed
        intSetDesignPower sCurrDesign
        ' Update the body points
        intSetBodyPoints sCurrDesign

        ' update name if not customised
        If Not bTemplateNameCustomised Then
            sCurrDesign.name_ = GetDefaultTemplateName(sCurrDesign)
        End If

        ' Update the name in the edit box
        intSetEditBoxTextFromTemplate sCurrDesign

        If COMPONENT_STATS_compType(apsExtraSysList(id - IDDES_EXTRASYSSTART)) = COMP_BRAIN Then
            triggerEvent TRIGGER_DESIGN_COMMAND
        Else
            triggerEvent TRIGGER_DESIGN_SYSTEM
        End If
    Else
        Select Case id
        ' The four component clickable forms
        ' the six component clickable forms...
        Case IDDES_WEAPONS
            desCompID = 0
            intSetDesignMode IDES_TURRET
        Case IDDES_WEAPONS_A
            desCompID = 0
            intSetDesignMode IDES_TURRET_A
        Case IDDES_WEAPONS_B
            desCompID = 0
            intSetDesignMode IDES_TURRET_B
        Case IDDES_COMMAND
            desCompID = 0
        Case IDDES_SYSTEMS
            desCompID = 0
            intSetDesignMode IDES_SYSTEM
        ' The name edit box
        Case IDDES_NAMEBOX
            sCurrDesign.name_ = widgGetWzString(psWScreen, IDDES_NAMEBOX)
            sstrcpy_aCurrName sCurrDesign.name_     ' getStatsName(&sCurrDesign)
        Case IDDES_BIN
            ' Find the template for the current button
            psTempl = templateFromButtonId(droidTemplID)  ' Does not return the first template, which is the empty template.

            ' remove template if found
            If psTempl <> 0 Then
                'update player template list.
                For k = 0 To localTemplates_size - 1
                    If localTemplates(k) = psTempl Then
                        'before deleting the template, need to make sure not being used in production
                        deleteTemplateFromProduction psTempl, selectedPlayer, ModeQueue
                        ' Delete the template.
                        localTemplates_erase_ k
                        Exit For
                    End If
                Next

                ' get previous template and set as current
                psTempl = templateFromButtonId(droidTemplID - 1, True)  ' droidTemplID - 1 always valid (might be the first template), since droidTemplID is not the first template.
                If Not (psTempl <> 0) Then ASSERT_FAILURE "psTempl != nullptr", "src/design.cpp:3367", "intProcessDesign", "template not found! - unexpected!": Exit Sub ' see above comment

                ' update local list
                desSetupDesignTemplates

                ' Now update the droid template form
                widgDelete psWScreen, IDDES_TEMPLBASE
                intAddTemplateForm psTempl

                ' Set the new template
                sCurrDesign = apsTemplates(psTempl)
                sstrcpy_aCurrName apsTemplates(psTempl).name_      ' getStatsName(psTempl)

                intSetEditBoxTextFromTemplate apsTemplates(psTempl)

                intSetDesignStats sCurrDesign

                ' show body stats only
                widgReveal psWScreen, IDDES_STATSFORM
                widgReveal psWScreen, IDDES_BODYFORM
                widgHide psWScreen, IDDES_PROPFORM
                widgHide psWScreen, IDDES_SYSTEMFORM

                'Update the Power bar stats as the power to build will have changed
                intSetDesignPower sCurrDesign
                'Update the body points
                intSetBodyPoints sCurrDesign

                ' show correct body component highlight
                widgDelete psWScreen, IDDES_RIGHTBASE
                ' reset button states
                widgSetButtonState psWScreen, IDDES_SYSTEMBUTTON, 0
                widgSetButtonState psWScreen, IDDES_BODYBUTTON, 0
                widgSetButtonState psWScreen, IDDES_PROPBUTTON, 0
                widgSetButtonState psWScreen, IDDES_WPABUTTON, 0
                widgSetButtonState psWScreen, IDDES_WPBBUTTON, 0
                desCompMode = IDES_NOCOMPONENT
                intSetDesignMode IDES_BODY
            End If
        Case IDDES_STOREBUTTON
            sCurrDesign.stored = Not sCurrDesign.stored    ' Invert the current status
            saveTemplate
            storeTemplates
            updateStoreButton sCurrDesign.stored
        Case IDDES_SYSTEMBUTTON
            ' Add the correct component form
            Select Case droidTemplateType(sCurrDesign)
            Case DROID_COMMAND, DROID_SENSOR, DROID_CONSTRUCT, DROID_ECM, DROID_REPAIR
                intSetDesignMode IDES_SYSTEM
            Case Else
                intSetDesignMode IDES_TURRET
            End Select
            ' reveal components if not already onscreen
            widgReveal psWScreen, IDDES_STATSFORM
            widgReveal psWScreen, IDDES_RIGHTBASE
            widgReveal psWScreen, IDDES_SYSTEMFORM
            widgHide psWScreen, IDDES_BODYFORM
            widgHide psWScreen, IDDES_PROPFORM

        ' WPABUTTON
        Case IDDES_WPABUTTON
            ' Add the correct component form
            Select Case droidTemplateType(sCurrDesign)
            Case DROID_COMMAND, DROID_SENSOR, DROID_CONSTRUCT, DROID_ECM, DROID_REPAIR
            Case Else
                intSetDesignMode IDES_TURRET_A
            End Select
            ' reveal components if not already onscreen
            widgReveal psWScreen, IDDES_STATSFORM
            widgReveal psWScreen, IDDES_RIGHTBASE
            widgReveal psWScreen, IDDES_SYSTEMFORM
            widgHide psWScreen, IDDES_BODYFORM
            widgHide psWScreen, IDDES_PROPFORM

        ' WPBBUTTON
        Case IDDES_WPBBUTTON
            ' Add the correct component form
            Select Case droidTemplateType(sCurrDesign)
            Case DROID_COMMAND, DROID_SENSOR, DROID_CONSTRUCT, DROID_ECM, DROID_REPAIR
            Case Else
                intSetDesignMode IDES_TURRET_B
            End Select
            ' reveal components if not already onscreen
            widgReveal psWScreen, IDDES_STATSFORM
            widgReveal psWScreen, IDDES_RIGHTBASE
            widgReveal psWScreen, IDDES_SYSTEMFORM
            widgHide psWScreen, IDDES_BODYFORM
            widgHide psWScreen, IDDES_PROPFORM

        Case IDDES_BODYBUTTON
            ' reveal components if not already onscreen
            widgReveal psWScreen, IDDES_RIGHTBASE
            intSetDesignMode IDES_BODY

            widgReveal psWScreen, IDDES_STATSFORM
            widgHide psWScreen, IDDES_SYSTEMFORM
            widgReveal psWScreen, IDDES_BODYFORM
            widgHide psWScreen, IDDES_PROPFORM

        Case IDDES_PROPBUTTON
            ' reveal components if not already onscreen
            widgReveal psWScreen, IDDES_RIGHTBASE
            intSetDesignMode IDES_PROPULSION
            widgReveal psWScreen, IDDES_STATSFORM
            widgHide psWScreen, IDDES_SYSTEMFORM
            widgHide psWScreen, IDDES_BODYFORM
            widgReveal psWScreen, IDDES_PROPFORM

        Case IDSTAT_OBSOLETE_BUTTON
            includeRedundantDesigns = Not includeRedundantDesigns
            Set obsoleteButton = widgGetFromID(psWScreen, IDSTAT_OBSOLETE_BUTTON)
            If includeRedundantDesigns Then choice = 1 Else choice = 0
            MultipleChoiceButton_setChoice obsoleteButton, choice
            ' Refresh lists.
            If droidTemplID <> IDDES_TEMPLSTART Then
                intRemoveDesign
                intAddDesign False
            Else
                desSetupDesignTemplates
                widgDelete psWScreen, IDDES_TEMPLBASE
                intAddTemplateForm templateFromButtonId(droidTemplID)
                intSetDesignMode desCompMode, True
                droidTemplID = IDDES_TEMPLSTART
                widgSetButtonState psWScreen, droidTemplID, WBUT_LOCK
            End If
        End Select
    End If

    ' show body button if component button pressed and
    ' save template if valid
    If (id >= IDDES_COMPSTART And id <= IDDES_COMPEND) Or (id >= IDDES_EXTRASYSSTART And id <= IDDES_EXTRASYSEND) Then
        ' reveal body button if hidden
        widgReveal psWScreen, IDDES_BODYBUTTON

        ' save template if valid
        If saveTemplate() Then
            triggerEventDesignCreated designTemplateHandle(sCurrDesign)
        End If

        Select Case desCompMode
        Case IDES_BODY
            widgReveal psWScreen, IDDES_BODYFORM
            widgHide psWScreen, IDDES_PROPFORM
            widgHide psWScreen, IDDES_SYSTEMFORM

        Case IDES_PROPULSION
            widgHide psWScreen, IDDES_BODYFORM
            widgReveal psWScreen, IDDES_PROPFORM
            widgHide psWScreen, IDDES_SYSTEMFORM

        Case IDES_SYSTEM, IDES_TURRET, IDES_TURRET_A, IDES_TURRET_B
        ' reveals SYSTEMFORM
            widgHide psWScreen, IDDES_BODYFORM
            widgHide psWScreen, IDDES_PROPFORM
            widgReveal psWScreen, IDDES_SYSTEMFORM
        Case Else
        End Select

        widgReveal psWScreen, IDDES_STATSFORM

        ' switch automatically to next component type if initial design
        If Not intValidTemplate(sCurrDesign, aCurrName, False, selectedPlayer) Then
            ' show next component design screen
            Select Case desCompMode
            Case IDES_BODY
                intSetDesignMode IDES_PROPULSION
                widgReveal psWScreen, IDDES_PROPBUTTON

            Case IDES_PROPULSION
                intSetDesignMode IDES_TURRET
                widgReveal psWScreen, IDDES_SYSTEMBUTTON

            Case IDES_SYSTEM, IDES_TURRET
                If asBodyStats(DROID_TEMPLATE_getBodyStats(sCurrDesign)).weaponSlots > 1 And _
                   sCurrDesign.numWeaps = 1 And sCurrDesign.asParts(COMP_BRAIN) = 0 Then
                    debug_ 3591, LOG_GUI, "intProcessDesign", "intProcessDesign: First weapon selected, doing next."
                    intSetDesignMode IDES_TURRET_A
                    widgReveal psWScreen, IDDES_WPABUTTON
                Else
                    debug_ 3597, LOG_GUI, "intProcessDesign", "intProcessDesign: First weapon selected, is final."
                End If
            Case IDES_TURRET_A
                If asBodyStats(DROID_TEMPLATE_getBodyStats(sCurrDesign)).weaponSlots > 2 Then
                    debug_ 3603, LOG_GUI, "intProcessDesign", "intProcessDesign: Second weapon selected, doing next."
                    intSetDesignMode IDES_TURRET_B
                    widgReveal psWScreen, IDDES_WPBBUTTON
                Else
                    debug_ 3609, LOG_GUI, "intProcessDesign", "intProcessDesign: Second weapon selected, is final."
                End If
            Case IDES_TURRET_B
                debug_ 3613, LOG_GUI, "intProcessDesign", "intProcessDesign: Third weapon selected, is final."
            Case Else
            End Select
        End If
    End If
    'save the template if the name gets edited
    If id = IDDES_NAMEBOX Then
        saveTemplate
    End If
End Sub


' Set the shadow bar graphs for the design screen
Public Sub intRunDesign()
    Dim psStats As Long, templateButton As Boolean, statID As Long, compIndex As Long

    ' Find out which button was hilited
    templateButton = False
    statID = widgGetMouseOver(psWScreen)

    ' Somut around here is casuing a nasty crash.....
    ' If a component button is hilited get the stats for it
    If statID = desCompID Then
        ' The mouse is over the selected component - no shadow stats
        psStats = 0
    ElseIf statID >= IDDES_COMPSTART And statID <= IDDES_COMPEND Then
        compIndex = statID - IDDES_COMPSTART
        If Not (ULT(compIndex, numComponent)) Then ASSERT_FAILURE "compIndex < numComponent", "src/design.cpp:3647", "intRunDesign", "Invalid range referenced for numComponent, %d > %d", compIndex, numComponent: Exit Sub
        psStats = apsComponentList(compIndex)
    ElseIf statID >= IDDES_EXTRASYSSTART And statID <= IDDES_EXTRASYSEND Then
        compIndex = statID - IDDES_EXTRASYSSTART
        If Not (ULT(compIndex, numExtraSys)) Then ASSERT_FAILURE "compIndex < numExtraSys", "src/design.cpp:3653", "intRunDesign", "Invalid range referenced for numExtraSys, %d > %d", compIndex, numExtraSys: Exit Sub
        psStats = apsExtraSysList(compIndex)
    ElseIf statID >= IDDES_TEMPLSTART And statID <= IDDES_TEMPLEND Then
        runTemplateShadowStats statID
        templateButton = True
        psStats = 0
    Else
        ' No component button so reset the stats to nothing
        psStats = 0
    End If

    ' Now set the bar graphs for the stats - don't bother if over template
    ' since setting them all!
    If Not templateButton Then
        Select Case desCompMode
        Case IDES_SYSTEM, IDES_TURRET, IDES_TURRET_A, IDES_TURRET_B
            intSetBodyShadowStats 0
            intSetPropulsionShadowStats 0
            intSetSystemShadowStats psStats
        Case IDES_BODY
            intSetSystemShadowStats 0
            intSetPropulsionShadowStats 0
            intSetBodyShadowStats psStats
        Case IDES_PROPULSION
            intSetSystemShadowStats 0
            intSetBodyShadowStats 0
            intSetPropulsionShadowStats psStats
        Case Else
        End Select

        'set the template shadow stats
        intSetTemplateBodyShadowStats psStats
        intSetTemplatePowerShadowStats psStats
    End If

    If keyPressed(KEY_ESC) Then
        intResetScreen False
        ' clear key press so we don't enter in-game options
        inputLoseFocus
    End If
End Sub

Private Sub intDisplayStatForm(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Static iRY As Long, iRY_set As Boolean     ' static UDWORD iRY = 45
    Dim Form_ As WIDGET, x0 As Long, y0 As Long, psStats As Long, templateRadius As Long, Rotation As Vector3i, Position As Vector3i
    Dim falseScale As Long
    If Not iRY_set Then iRY = 45: iRY_set = True
    Set Form_ = psWidget        ' W_CLICKFORM *
    x0 = (xOffset + Form_.x()) And &HFFFF&
    y0 = (yOffset + Form_.y()) And &HFFFF&

    ' get stats from userdata pointer in widget stored in
    ' intSetSystemStats, intSetBodyStats, intSetPropulsionStats
    psStats = Form_.pUserData

    templateRadius = getComponentRadius(psStats)

    Rotation = Vector3i_(-30, iRY, 0)
    Position = Vector3i_(0, -templateRadius \ 4, BUTTON_DEPTH)   ' templateRadius * 12

    'scale the object around the BUTTON_RADIUS so that half size objects are draw are draw 75% the size of normal objects
    falseScale = (DESIGN_COMPONENT_SCALE * COMPONENT_RADIUS) \ templateRadius \ 2 + (DESIGN_COMPONENT_SCALE \ 2)

    iV_DrawImage IntImages, IMAGE_DES_STATBACKLEFT, x0, y0
    iV_DrawImageRepeatX IntImages, IMAGE_DES_STATBACKMID, x0 + iV_GetImageWidth(IntImages, IMAGE_DES_STATBACKLEFT), y0, _
                        Form_.width() - iV_GetImageWidth(IntImages, IMAGE_DES_STATBACKLEFT) - iV_GetImageWidth(IntImages, IMAGE_DES_STATBACKRIGHT), 0, True
    iV_DrawImage IntImages, IMAGE_DES_STATBACKRIGHT, x0 + Form_.width() - iV_GetImageWidth(IntImages, IMAGE_DES_STATBACKRIGHT), y0

    ' display current component
    pie_SetGeometricOffset xOffset + psWidget.width() \ 4, yOffset + psWidget.height() \ 2

    ' inc rotation if highlighted
    If (Form_.getState() And WBUT_HIGHLIGHT) <> 0 Then
        iRY = iRY + realTimeAdjustedAverage(BUTTONOBJ_ROTSPEED)
        iRY = iRY Mod 360
    End If

    'display component in bottom design screen window
    displayComponentButton psStats, Rotation, Position, falseScale
End Sub

' Displays the 3D view of the droid in a window on the design form
Private Sub intDisplayViewForm(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Static iRY As Long, iRY_set As Boolean     ' static UDWORD iRY = 45
    Dim Form_ As WIDGET, x0 As Long, y0 As Long, x1 As Long, y1 As Long, Rotation As Vector3i, Position As Vector3i
    Dim templateRadius As Long, falseScale As Long
    If Not iRY_set Then iRY = 45: iRY_set = True
    Set Form_ = psWidget        ' W_FORM *

    x0 = xOffset + Form_.x()
    y0 = yOffset + Form_.y()
    x1 = x0 + Form_.width()
    y1 = y0 + Form_.height()

    RenderWindowFrame FRAME_NORMAL, x0, y0, x1 - x0, y1 - y0

    pie_SetGeometricOffset (DES_CENTERFORMX + DES_3DVIEWX) + (DES_3DVIEWWIDTH \ 2), _
                           (DES_CENTERFORMY + DES_3DVIEWY) + (DES_3DVIEWHEIGHT \ 4) + 32

    Rotation = Vector3i_(-30, iRY, 0)

    ' inc rotation
    iRY = iRY + realTimeAdjustedAverage(BUTTONOBJ_ROTSPEED)
    iRY = iRY Mod 360

    'fixed depth scale the pie
    Position = Vector3i_(0, -100, BUTTON_DEPTH)

    templateRadius = getComponentDroidTemplateRadius(designTemplateHandle(sCurrDesign))
    'scale the object around the OBJECT_RADIUS so that half size objects are draw are draw 75% the size of normal objects
    falseScale = (DESIGN_DROID_SCALE * OBJECT_RADIUS) \ templateRadius

    'display large droid view in the design screen
    displayComponentButtonTemplate designTemplateHandle(sCurrDesign), Rotation, Position, falseScale
End Sub

' General display window for the design form  SOLID BACKGROUND - NOT TRANSPARENT
Private Sub intDisplayDesignForm(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long
    x0 = xOffset + psWidget.x()
    y0 = yOffset + psWidget.y()
    x1 = x0 + psWidget.width()
    y1 = y0 + psWidget.height()

    RenderWindowFrame FRAME_NORMAL, x0, y0, x1 - x0, y1 - y0
End Sub


' save the current Template if valid. Return true if stored
Private Function saveTemplate() As Boolean
    Dim psTempl As Long, blank As DROID_TEMPLATE, t As DROID_TEMPLATE

    If Not intValidTemplate(sCurrDesign, aCurrName, False, selectedPlayer) Then
        If bMultiPlayer Then
            widgHide psWScreen, IDDES_STOREBUTTON
        End If
        saveTemplate = False
        Exit Function
    End If
    If bMultiPlayer Then
        widgReveal psWScreen, IDDES_STOREBUTTON
        updateStoreButton sCurrDesign.stored   ' Change the buttons icon
    End If

    ' if first (New Design) button selected find empty template
    ' else find current button template
    If droidTemplID = IDDES_TEMPLSTART Then
        ' create empty template and point to that
        DROID_TEMPLATE_construct blank
        localTemplates_push_back blank
        psTempl = localTemplates(localTemplates_size - 1)
        sCurrDesign.multiPlayerID = generateNewObjectId()
        apsTemplateList_add psTempl

        apsTemplates(psTempl).ref = STAT_TEMPLATE

        ' set button render routines to highlight, not flash
        intSetButtonFlash IDDES_SYSTEMBUTTON, False
        intSetButtonFlash IDDES_BODYBUTTON, False
        intSetButtonFlash IDDES_PROPBUTTON, False
    Else
        ' Find the template for the current button
        psTempl = templateFromButtonId(droidTemplID)
        If psTempl = 0 Then
            debug_ 3833, LOG_ERROR, "saveTemplate", "Template not found for button"
            saveTemplate = False
            Exit Function
        End If

        ' ANY change to the template affect the production - even if the template is changed and then changed back again!
        deleteTemplateFromProduction psTempl, selectedPlayer, ModeQueue
    End If

    ' Copy the template
    apsTemplates(psTempl) = sCurrDesign

    ' Now update the droid template form
    widgDelete psWScreen, IDDES_TEMPLBASE
    intAddTemplateForm psTempl

    ' Add template to in-game template list, since localTemplates/apsTemplateList is for UI use only.
    t = apsTemplates(psTempl)      ' (VB) a copy: copyTemplate's pool may grow
    copyTemplate selectedPlayer, t

    saveTemplate = True
End Function


' Function to set the shadow bars for all the stats when the mouse is over the Template buttons
Private Sub runTemplateShadowStats(ByVal id As Long)
    Dim psTempl As Long, psStats As Long, templType As DROID_TYPE, compIndex As Long

    ' Find the template for the new button
    'we're ignoring the Blank Design so start at the second button
    psTempl = templateFromButtonId(id)

    'if we're over a different template
    If psTempl <> 0 Then   ' (VB) psTempl != &sCurrDesign: a pool template is never the design screen's own copy
        ' Now set the bar graphs for the stats
        intSetBodyShadowStats STAT_BODY + DROID_TEMPLATE_getBodyStats(apsTemplates(psTempl))
        intSetPropulsionShadowStats STAT_PROPULSION + DROID_TEMPLATE_getPropulsionStats(apsTemplates(psTempl))
        'only set the system shadow bar if the same type of droid
        psStats = 0
        templType = droidTemplateType(apsTemplates(psTempl))
        If templType = droidTemplateType(sCurrDesign) Then
            Select Case templType
            Case DROID_WEAPON
                compIndex = apsTemplates(psTempl).asWeaps(0)
                If Not (ULT(compIndex, asWeaponStats_size)) Then ASSERT_FAILURE "compIndex < asWeaponStats.size()", "src/design.cpp:3878", "runTemplateShadowStats", "Invalid range referenced for numWeaponStats, %d > %zu", compIndex, asWeaponStats_size: Exit Sub
                psStats = STAT_WEAPON + compIndex
            Case DROID_SENSOR
                compIndex = apsTemplates(psTempl).asParts(COMP_SENSOR)
                If Not (ULT(compIndex, asSensorStats_size)) Then ASSERT_FAILURE "compIndex < asSensorStats.size()", "src/design.cpp:3883", "runTemplateShadowStats", "Invalid range referenced for numSensorStats, %d > %zu", compIndex, asSensorStats_size: Exit Sub
                psStats = STAT_SENSOR + compIndex
            Case DROID_ECM
                compIndex = apsTemplates(psTempl).asParts(COMP_ECM)
                If Not (ULT(compIndex, asECMStats_size)) Then ASSERT_FAILURE "compIndex < asECMStats.size()", "src/design.cpp:3888", "runTemplateShadowStats", "Invalid range referenced for numECMStats, %d > %zu", compIndex, asECMStats_size: Exit Sub
                psStats = STAT_ECM + compIndex
            Case DROID_CONSTRUCT
                compIndex = apsTemplates(psTempl).asParts(COMP_CONSTRUCT)
                If Not (ULT(compIndex, asConstructStats_size)) Then ASSERT_FAILURE "compIndex < asConstructStats.size()", "src/design.cpp:3893", "runTemplateShadowStats", "Invalid range referenced for numConstructStats, %d > %zu", compIndex, asConstructStats_size: Exit Sub
                psStats = STAT_CONSTRUCT + compIndex
            Case DROID_REPAIR
                compIndex = apsTemplates(psTempl).asParts(COMP_REPAIRUNIT)
                If Not (ULT(compIndex, asRepairStats_size)) Then ASSERT_FAILURE "compIndex < asRepairStats.size()", "src/design.cpp:3898", "runTemplateShadowStats", "Invalid range referenced for numRepairStats, %d > %zu", compIndex, asRepairStats_size: Exit Sub
                psStats = STAT_REPAIR + compIndex
            Case Else
            End Select
        End If

        If psStats <> 0 Then
            intSetSystemShadowStats psStats
        End If
        'haven't got a stat so just do the code required here...
        widgSetMinorBarSize psWScreen, IDDES_BODYPOINTS, calcTemplateBody(psTempl, selectedPlayer)
        widgSetMinorBarSize psWScreen, IDDES_POWERBAR, calcTemplatePower(apsTemplates(psTempl))
    End If
End Sub

'sets which states need to be paused when the design screen is up
Private Sub setDesignPauseState()
    If Not bMultiPlayer And Not bInTutorial Then
        'need to clear mission widgets from being shown on design screen
        clearMissionWidgets
        intHideGroupSelectionMenu
        gameTimeStop
        setGameUpdatePause True
        setScrollPause True
        screen_RestartBackDrop
    End If
End Sub

'resets the pause states
Private Sub resetDesignPauseState()
    If Not bMultiPlayer And Not bInTutorial Then
        'put any widgets back on for the missions
        resetMissionWidgets
        intShowGroupSelectionMenu
        setGameUpdatePause False
        setScrollPause False
        gameTimeStart
        screen_StopBackDrop
    End If
End Sub

'this is called when a new propulsion type is added to the current design
'to check the weapon is 'allowed'. Check if VTOL, the weapon is direct fire.
'Also check numVTOLattackRuns for the weapon is not zero - return true if valid weapon
Private Function intCheckValidWeaponForProp(psTemplate As DROID_TEMPLATE) As Boolean
    If asPropulsionTypes(asPropulsionStats(DROID_TEMPLATE_getPropulsionStats(psTemplate)).propulsionType).travel <> AIR Then
        If psTemplate.numWeaps = 0 And _
           (psTemplate.asParts(COMP_SENSOR) <> 0 Or _
            psTemplate.asParts(COMP_REPAIRUNIT) <> 0 Or _
            psTemplate.asParts(COMP_CONSTRUCT) <> 0 Or _
            psTemplate.asParts(COMP_ECM) <> 0) Then
            ' non-AIR propulsions can have systems, too.
            intCheckValidWeaponForProp = True
            Exit Function
        End If
    End If
    intCheckValidWeaponForProp = checkValidWeaponForProp(psTemplate)
End Function

'checks if the template has PROPULSION_TYPE_LIFT propulsion attached - returns true if it does
Private Function checkTemplateIsVtol(psTemplate As DROID_TEMPLATE) As Boolean
    checkTemplateIsVtol = asPropulsionStats(DROID_TEMPLATE_getPropulsionStats(psTemplate)).propulsionType = PROPULSION_TYPE_LIFT
End Function

Private Sub updateStoreButton(ByVal isStored As Boolean)
    Dim imageset As Long

    If isStored Then
        imageset = PACKDWORD_TRI(0, IMAGE_DES_DELETEH, IMAGE_DES_DELETE)
        widgGetFromID(psWScreen, IDDES_STOREBUTTON).setTip "Do Not Store Design"
    Else
        imageset = PACKDWORD_TRI(0, IMAGE_DES_SAVEH, IMAGE_DES_SAVE)
        widgGetFromID(psWScreen, IDDES_STOREBUTTON).setTip "Store Design"
    End If

    widgSetUserData2 psWScreen, IDDES_STOREBUTTON, imageset
End Sub

' (VB) the lambdas, std::functions and function pointers of design.cpp
Public Function design_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, templList As WIDGET, templbaseForm As WIDGET, compList As WIDGET, rightBase As WIDGET
    Dim callback As WzFunc, button As WIDGET, i As Long, n As Long, psCurrStats As Long, aAvailable As Long, v As Long
    Dim templListWidth As Long, objListWidth As Long, currStats As Long, index As Long, nextButtonId As Long, w As Long

    Select Case f.id
    Case FN_intAddDesign_desForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry DES_CENTERFORMX, DES_CENTERFORMY, DES_CENTERFORMWIDTH, DES_CENTERFORMHEIGHT
    Case FN_intAddDesign_statsForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry DES_STATSFORMX, DES_STATSFORMY, DES_STATSFORMWIDTH, DES_STATSFORMHEIGHT
    Case FN_intAddTemplateForm_templbaseForm_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry RET_X, DESIGN_Y, RET_FORMWIDTH, DES_LEFTFORMHEIGHT
    Case FN_intAddTemplateForm_templList_calcLayout
        Set psWidget = args(0)
        Set templList = psWidget   ' static_cast<IntListTabWidget *>(psWidget)
        ' assert(templList != nullptr); (NDEBUG in a release build)
        ListTabWidget_setChildSize templList, DES_TABBUTWIDTH, DES_TABBUTHEIGHT
        ListTabWidget_setChildSpacing templList, DES_TABBUTGAP, DES_TABBUTGAP
        Set templbaseForm = templList.parent()
        If Not templbaseForm Is Nothing Then
            templListWidth = OBJ_BUTWIDTH * 2 + DES_TABBUTGAP
            templList.setGeometry (RET_FORMWIDTH - templListWidth) \ 2, 18, templListWidth, templbaseForm.height() - 18
        End If
    Case FN_intAddComponentForm_rightBase_calcLayout
        Set psWidget = args(0)
        psWidget.setGeometry RADTLX - 2, DESIGN_Y, RET_FORMWIDTH, DES_RIGHTFORMHEIGHT
    Case FN_intAddComponentForm_compList_calcLayout
        Set psWidget = args(0)
        Set compList = psWidget    ' static_cast<IntListTabWidget *>(psWidget)
        ' assert(compList != nullptr); (NDEBUG in a release build)
        ListTabWidget_setChildSize compList, DES_TABBUTWIDTH, DES_TABBUTHEIGHT
        ListTabWidget_setChildSpacing compList, DES_TABBUTGAP, DES_TABBUTGAP
        Set rightBase = compList.parent()
        If Not rightBase Is Nothing Then
            objListWidth = DES_TABBUTWIDTH * 2 + DES_TABBUTGAP
            compList.setGeometry (rightBase.width() - objListWidth) \ 2, 40, objListWidth, rightBase.height() - 40
        End If
    Case FN_intDisplayStatForm
        intDisplayStatForm args(0), args(1), args(2)
    Case FN_intDisplayViewForm
        intDisplayViewForm args(0), args(1), args(2)
    Case FN_intDisplayDesignForm
        intDisplayDesignForm args(0), args(1), args(2)

    Case FN_componentIterator
        ' [=](std::function<bool(COMPONENT_STATS &, size_t index)> callback): cap (psStats, aAvailable's player and type, numEntries)
        Set callback = args(0)
        n = f.cap(3)
        For i = 0 To n - 1
            psCurrStats = f.cap(0) + i

            ' Skip unavailable entries and non-design ones
            aAvailable = apCompLists(f.cap(1), f.cap(2), i)
            If (aAvailable = AVAILABLE Or (includeRedundantDesigns And aAvailable = REDUNDANT)) And COMPONENT_STATS_designable(psCurrStats) Then
                If Not WzFunc_call(callback, psCurrStats, i) Then
                    design_c_fn = False
                    Exit Function
                End If
            End If
        Next

        design_c_fn = True
    Case FN_concatIterators
        ' [=](std::function<bool(COMPONENT_STATS &, size_t index)> callback): cap (count, iterators...)
        Set callback = args(0)
        n = f.cap(0)
        For i = 1 To n
            If Not WzFunc_call(f.cap(i), callback) Then
                design_c_fn = False
                Exit Function
            End If
        Next

        design_c_fn = True
    Case FN_findMax_callback
        ' [&](COMPONENT_STATS &bodyStats, size_t index): max = std::max(max, value(bodyStats)); cap (value, max)
        v = WzFunc_call(f.cap(0), args(0))
        If ULT(f.cap(1), v) Then f.cap(1) = v
        design_c_fn = True
    Case FN_findMaxWeight_value
        design_c_fn = COMPONENT_STATS_weight(args(0))
    Case FN_findMaxPropulsionSpeed_value
        design_c_fn = intCalcSpeed(f.cap(0), args(0))
    Case FN_findMaxWeaponAttribute_value
        design_c_fn = WzFunc_call(f.cap(0), args(0), selectedPlayer)
    Case FN_getDesignMaxBodyArmour_value
        design_c_fn = bodyArmour(asBodyStats(args(0) And &HFFFF&), selectedPlayer, f.cap(0))
    Case FN_getDesignMaxEngineOutput_value
        design_c_fn = bodyPower(asBodyStats(args(0) And &HFFFF&), selectedPlayer)
    Case FN_calcShadowBodyPoints
        design_c_fn = calcShadowBodyPoints(args(0))
    Case FN_calcShadowPower
        design_c_fn = calcShadowPower(args(0))
    Case FN_getDesignMaxSensorRange_value
        design_c_fn = sensorRange(asSensorStats(args(0) And &HFFFF&), selectedPlayer)
    Case FN_getDesignMaxEcmRange_value
        design_c_fn = ecmRange(asECMStats(args(0) And &HFFFF&), selectedPlayer)
    Case FN_getDesignMaxBuildPoints_value
        design_c_fn = constructorPoints(asConstructStats(args(0) And &HFFFF&), selectedPlayer)
    Case FN_intSetSystemForm_repairPoints
        design_c_fn = repairPoints(asRepairStats(args(0) And &HFFFF&), selectedPlayer)
    Case FN_calculatePropulsionWeight
        design_c_fn = calculatePropulsionWeight(args(0))
    Case FN_design_proj_GetLongRange
        design_c_fn = proj_GetLongRange(asWeaponStats(args(0) And &HFFFF&), args(1))
    Case FN_design_weaponDamage
        design_c_fn = weaponDamage(asWeaponStats(args(0) And &HFFFF&), args(1))
    Case FN_design_weaponROF
        design_c_fn = weaponROF(asWeaponStats(args(0) And &HFFFF&), args(1))

    Case FN_intAddComponentButtons_callback
        ' [&](COMPONENT_STATS &currStats, size_t index): cap (compList, maxComponents, bWeapon, bVTOL, bodysize, nextButtonId, compID)
        currStats = args(0)
        index = args(1)
        Set compList = f.cap(0)
        ' If we are out of space in the list - stop
        If Not ULT(numComponent, f.cap(1)) Then
            design_c_fn = False
            Exit Function
        End If

        'skip indirect weapons if VTOL propulsion or numVTOLattackRuns for the weapon is zero
        If f.cap(2) Then
            w = currStats And &HFFFF&      ' auto &weapon = (WEAPON_STATS &)currStats
            If (asWeaponStats(w).vtolAttackRuns > 0) <> f.cap(3) _
               Or (asWeaponStats(w).weaponSize = WEAPON_SIZE_LIGHT And f.cap(4) <> SIZE_LIGHT) _
               Or (asWeaponStats(w).weaponSize = WEAPON_SIZE_HEAVY And f.cap(4) = SIZE_LIGHT) Then
                design_c_fn = True
                Exit Function
            End If
        End If

        ' Set the tip and add the button
        nextButtonId = f.cap(5)
        Set button = IntStatsButton_new()
        compList.attach button
        button.id = nextButtonId
        IntStatsButton_setStatsAndTip button, currStats
        ListTabWidget_addWidgetToLayout compList, button

        ' Store the stat pointer in the list
        apsComponentList(numComponent) = currStats
        numComponent = numComponent + 1

        ' If this matches the component ID lock the button
        If index = f.cap(6) Then
            desCompID = nextButtonId
            button.setState WBUT_LOCK
            ListTabWidget_setCurrentPage compList, ListTabWidget_pages(compList) - 1
        End If

        ' Update the init struct for the next button
        f.cap(5) = nextButtonId + 1

        design_c_fn = True
    Case FN_intAddExtraSystemButtons_callback
        ' [&](COMPONENT_STATS &stats, size_t i): cap (compList, nextButtonId, compIndex)
        currStats = args(0)
        index = args(1)
        Set compList = f.cap(0)
        ' If we are out of space in the list - stop
        If Not ULT(numExtraSys, MAXEXTRASYS) Then
            ASSERT_FAILURE "false", "src/design.cpp:2148", "intAddExtraSystemButtons", "Too many components for the list"
            design_c_fn = False
            Exit Function
        End If

        ' Set the tip and add the button
        nextButtonId = f.cap(1)
        Set button = IntStatsButton_new()
        compList.attach button
        button.id = nextButtonId
        IntStatsButton_setStatsAndTip button, currStats
        ListTabWidget_addWidgetToLayout compList, button

        'just use one set of buffers for mixed system form
        If COMPONENT_STATS_compType(currStats) = COMP_BRAIN Then
            w = asBrainStats(currStats And &HFFFF&).psWeaponStat   ' (WEAPON_STATS *: an index, -1 = nullptr)
            If w <> -1 Then IntStatsButton_setStats button, STAT_WEAPON + w Else IntStatsButton_setStats button, 0
        End If

        ' Store the stat pointer in the list
        apsExtraSysList(numExtraSys) = currStats
        numExtraSys = numExtraSys + 1

        ' If this matches the sensorIndex note the form and button
        If index = f.cap(2) Then
            desCompID = nextButtonId
            button.setState WBUT_LOCK
            ListTabWidget_setCurrentPage compList, ListTabWidget_pages(compList) - 1
        End If

        ' Update the init struct for the next button
        f.cap(1) = nextButtonId + 1
        design_c_fn = True
    End Select
End Function
