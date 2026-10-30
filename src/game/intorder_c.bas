Attribute VB_Name = "intorder_c"
' src/intorder.cpp, intorder.h: the droid (and factory) order form, the
' secondary orders' buttons, whole. The OrderButtons table is set by
' intorder_c_init (C++'s static initialiser; AcNumButs changes as the form is
' built). std::vector<DROID *> SelectedDroids and std::vector<AVORDER>
' AvailableOrders are arrays with their sizes: an AVORDER is its OrderIndex.
' The buttons draw with intdisplay.cpp's intDisplayImageHilight and
' intDisplayButtonHilight (fork G1's) through this module's WzFuncs.
Option Explicit

Public Const IDORDER_FORM As Long = 8000
Public Const IDORDER_CLOSE As Long = 8001

Private Const ORDER_X As Long = 6
Private Const ORDER_Y As Long = 45
' ORDER_WIDTH is RET_FORMWIDTH (hci.h)
Private Const ORDER_HEIGHT As Long = 273
Private Const ORDER_BUTX As Long = 8
Private Const ORDER_BUTY As Long = 16
Private Const ORDER_BUTGAP As Long = 4
' ORDER_BOTTOMY is 318 + E_H (hci.h's E_H: the screen's height)

Private Const MAX_DISPLAYABLE_ORDERS As Long = 11 ' Max number of displayable orders.
Private Const MAX_ORDER_BUTS As Long = 5 ' Max number of buttons for a given order.
Private Const NUM_ORDERS As Long = 13   ' Number of orders in OrderButtons list.

Private Const IDORDER_ATTACK_RANGE As Long = 8010
Private Const IDORDER_REPAIR_LEVEL As Long = 8020
Private Const IDORDER_ATTACK_LEVEL As Long = 8030
Private Const IDORDER_PATROL As Long = 8040
Private Const IDORDER_HALT_TYPE As Long = 8050
Private Const IDORDER_RETURN As Long = 8060
Private Const IDORDER_RECYCLE As Long = 8070
Private Const IDORDER_ASSIGN_PRODUCTION As Long = 8080
Private Const IDORDER_ASSIGN_CYBORG_PRODUCTION As Long = 8090
Private Const IDORDER_FIRE_DESIGNATOR As Long = 8100
Private Const IDORDER_ASSIGN_VTOL_PRODUCTION As Long = 8110
Private Const IDORDER_CIRCLE As Long = 8120
Private Const IDORDER_ACCREP As Long = 8130

Private Enum ORDBUTTONTYPE
    ORD_BTYPE_RADIO                     ' Only one state allowed.
    ORD_BTYPE_BOOLEAN                   ' Clicking on a button toggles it's state.
    ORD_BTYPE_BOOLEAN_DEPEND            ' Clicking on a button toggles it's state, button
    ' is only enabled if previous button is down.
    ORD_BTYPE_BOOLEAN_COMBINE           ' Clicking on a button toggles it's state, all
    ' the buttons states are OR'ed together to get the order state
End Enum

Private Enum ORDBUTTONCLASS
    ORDBUTCLASS_NORMAL                  ' A normal button, one order type per line.
    ORDBUTCLASS_FACTORY                 ' A factory assignment button.
    ORDBUTCLASS_CYBORGFACTORY           ' A cyborg factory assignment button.
    ORDBUTCLASS_VTOLFACTORY             ' A VTOL factory assignment button.
End Enum


'
' NOTE:
'	ORD_BTYPE_BOOLEAN_DEPEND only supports two buttons
'	ie button 1 = enable destruct, button 2 = destruct
'

Private Enum ORDBUTTONJUSTIFY
    ORD_JUSTIFY_LEFT                    ' Pretty self explanatory really.
    ORD_JUSTIFY_RIGHT
    ORD_JUSTIFY_CENTER
    ORD_JUSTIFY_COMBINE                 ' allow the button to be put on the same line
    ' as other orders with the same justify type
    ORD_NUM_JUSTIFY_TYPES
End Enum

' maximum number of orders on one line
Private Const ORD_MAX_COMBINE_BUTS As Long = 3

Private Const ORD_JUSTIFY_MASK As Long = &HF
Private Const ORD_JUSTIFY_NEWLINE As Long = &H10 ' Or with ORDBUTTONJUSTIFY enum type to specify start on new line.

Private Type ORDERBUTTONS_              ' struct ORDERBUTTONS (VB: OrderButtons is the table)
    Class_ As ORDBUTTONCLASS            ' The class of button.
    Order As SECONDARY_ORDER            ' The droid order.
    StateMask As Long                   ' UDWORD: It's state mask.
    ButType As ORDBUTTONTYPE            ' The group type.
    ButJustify As Long                  ' unsigned: Button justification. Type ORDBUTTONJUSTIFY, possibly ored with ORD_JUSTIFY_NEWLINE.
    ButBaseID As Long                   ' UDWORD: Starting widget ID for buttons
    NumButs As Long                     ' UWORD: Number of buttons ( = number of states )
    AcNumButs As Long                   ' UWORD: Actual bumber of buttons enabled.
    ButImageID(0 To MAX_ORDER_BUTS - 1) As Long     ' UWORD: Image ID's for each button ( normal ).
    ButGreyID(0 To MAX_ORDER_BUTS - 1) As Long      ' Image ID's for each button ( greyed ).
    ButHilightID(0 To MAX_ORDER_BUTS - 1) As Long   ' Image ID's for each button ( hilight overlay ).
    ButTips(0 To MAX_ORDER_BUTS - 1) As Long        ' Tip string id for each button.
    States(0 To MAX_ORDER_BUTS - 1) As Long         ' unsigned: Order state relating to each button, combination of SECONDARY_STATEs ored together.
End Type

' struct AVORDER { UWORD OrderIndex; }: Index into ORDERBUTTONS array of available orders (here the Long itself)

Private Enum STR_DORD
    STR_DORD_RANGE1
    STR_DORD_RANGE2
    STR_DORD_RANGE3
    STR_DORD_REPAIR1
    STR_DORD_REPAIR2
    STR_DORD_REPAIR3
    STR_DORD_FIRE1
    STR_DORD_FIRE2
    STR_DORD_FIRE3
    STR_DORD_PATROL
    STR_DORD_PURSUE
    STR_DORD_GUARD
    STR_DORD_HOLDPOS
    STR_DORD_RETREPAIR
    STR_DORD_RETBASE
    STR_DORD_EMBARK
    STR_DORD_ARMRECYCLE
    STR_DORD_RECYCLE
    STR_DORD_FACTORY
    STR_DORD_CYBORG_FACTORY
    STR_DORD_FIREDES
    STR_DORD_VTOL_FACTORY
    STR_DORD_CIRCLE
    STR_DORD_ACCREP
End Enum

' Define the order button groups.
Private OrderButtons(0 To NUM_ORDERS - 1) As ORDERBUTTONS_


Private SelectedDroids() As Long        ' std::vector<DROID *>: SelectedDroids(0 .. SelectedDroids_size - 1)
Private SelectedDroids_size As Long
Private psSelectedFactory As Long       ' STRUCTURE * = nullptr
Private AvailableOrders() As Long       ' std::vector<AVORDER>: AvailableOrders(0 .. AvailableOrders_size - 1)
Private AvailableOrders_size As Long


Public OrderUp As Boolean               ' = false

' (VB) the lambdas
Public Const FN_intAddOrder_layout As Long = 11020
Public Const FN_intAddOrder_newHeightLayout As Long = 11021   ' [newHeight]
Public Const FN_intorder_intDisplayImageHilight As Long = 11022 ' intdisplay.cpp's intDisplayImageHilight
Public Const FN_intorder_intDisplayButtonHilight As Long = 11023 ' intdisplay.cpp's intDisplayButtonHilight

' return translated text
Private Function getDORDDescription(ByVal id As Long) As String
    Select Case id
    Case STR_DORD_RANGE1: getDORDDescription = "Short Range"
    Case STR_DORD_RANGE2: getDORDDescription = "Long Range"
    Case STR_DORD_RANGE3: getDORDDescription = "Optimum Range"
    Case STR_DORD_REPAIR1: getDORDDescription = "Retreat at Medium Damage"
    Case STR_DORD_REPAIR2: getDORDDescription = "Retreat at Heavy Damage"
    Case STR_DORD_REPAIR3: getDORDDescription = "Do or Die!"
    Case STR_DORD_FIRE1: getDORDDescription = "Fire-At-Will"
    Case STR_DORD_FIRE2: getDORDDescription = "Return Fire"
    Case STR_DORD_FIRE3: getDORDDescription = "Hold Fire"
    Case STR_DORD_PATROL: getDORDDescription = "Patrol"
    Case STR_DORD_PURSUE: getDORDDescription = "Pursue"
    Case STR_DORD_GUARD: getDORDDescription = "Guard Position"
    Case STR_DORD_HOLDPOS: getDORDDescription = "Hold Position"
    Case STR_DORD_RETREPAIR: getDORDDescription = "Return For Repair"
    Case STR_DORD_RETBASE: getDORDDescription = "Return To HQ"
    Case STR_DORD_EMBARK: getDORDDescription = "Go to Transport"
    Case STR_DORD_ARMRECYCLE: getDORDDescription = "Return for Recycling"
    Case STR_DORD_RECYCLE: getDORDDescription = "Recycle"
    Case STR_DORD_FACTORY: getDORDDescription = "Assign Factory Production"
    Case STR_DORD_CYBORG_FACTORY: getDORDDescription = "Assign Cyborg Factory Production"
    Case STR_DORD_FIREDES: getDORDDescription = "Assign Fire Support"
    Case STR_DORD_VTOL_FACTORY: getDORDDescription = "Assign VTOL Factory Production"
    Case STR_DORD_CIRCLE: getDORDDescription = "Circle"
    Case STR_DORD_ACCREP: getDORDDescription = "Accept Retreating Units"

    Case Else: getDORDDescription = ""  ' make compiler shut up
    End Select
End Function

' (VB) one row of the OrderButtons table: the five arrays as Array(...)s
Private Sub setOrderButton(ByVal i As Long, ByVal Class_ As ORDBUTTONCLASS, ByVal Order As SECONDARY_ORDER, ByVal StateMask As Long, ByVal ButType As ORDBUTTONTYPE, ByVal ButJustify As Long, ByVal ButBaseID As Long, ByVal NumButs As Long, ByVal AcNumButs As Long, _
                           ByVal ButImageID As Variant, ByVal ButGreyID As Variant, ByVal ButHilightID As Variant, ByVal ButTips As Variant, ByVal States As Variant)
    Dim k As Long
    OrderButtons(i).Class_ = Class_
    OrderButtons(i).Order = Order
    OrderButtons(i).StateMask = StateMask
    OrderButtons(i).ButType = ButType
    OrderButtons(i).ButJustify = ButJustify
    OrderButtons(i).ButBaseID = ButBaseID
    OrderButtons(i).NumButs = NumButs
    OrderButtons(i).AcNumButs = AcNumButs
    For k = 0 To MAX_ORDER_BUTS - 1     ' the missing ones are 0
        OrderButtons(i).ButImageID(k) = 0
        OrderButtons(i).ButGreyID(k) = 0
        OrderButtons(i).ButHilightID(k) = 0
        OrderButtons(i).ButTips(k) = 0
        OrderButtons(i).States(k) = 0
        If k <= UBound(ButImageID) Then OrderButtons(i).ButImageID(k) = ButImageID(k)
        If k <= UBound(ButGreyID) Then OrderButtons(i).ButGreyID(k) = ButGreyID(k)
        If k <= UBound(ButHilightID) Then OrderButtons(i).ButHilightID(k) = ButHilightID(k)
        If k <= UBound(ButTips) Then OrderButtons(i).ButTips(k) = ButTips(k)
        If k <= UBound(States) Then OrderButtons(i).States(k) = States(k)
    Next
End Sub

' C++'s static initialisers: static ORDERBUTTONS OrderButtons[NUM_ORDERS] = { ... }
Public Sub intorder_c_init()
    setOrderButton 0, ORDBUTCLASS_NORMAL, DSO_ATTACK_RANGE, DSS_ARANGE_MASK, ORD_BTYPE_RADIO, ORD_JUSTIFY_CENTER Or ORD_JUSTIFY_NEWLINE, IDORDER_ATTACK_RANGE, 3, 0, _
        Array(IMAGE_ORD_RANGE3UP, IMAGE_ORD_RANGE1UP, IMAGE_ORD_RANGE2UP), _
        Array(IMAGE_ORD_RANGE3UP, IMAGE_ORD_RANGE1UP, IMAGE_ORD_RANGE2UP), _
        Array(IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT), _
        Array(STR_DORD_RANGE3, STR_DORD_RANGE1, STR_DORD_RANGE2), _
        Array(DSS_ARANGE_OPTIMUM, DSS_ARANGE_SHORT, DSS_ARANGE_LONG)
    setOrderButton 1, ORDBUTCLASS_NORMAL, DSO_REPAIR_LEVEL, DSS_REPLEV_MASK, ORD_BTYPE_RADIO, ORD_JUSTIFY_CENTER Or ORD_JUSTIFY_NEWLINE, IDORDER_REPAIR_LEVEL, 3, 0, _
        Array(IMAGE_ORD_REPAIR3UP, IMAGE_ORD_REPAIR2UP, IMAGE_ORD_REPAIR1UP), _
        Array(IMAGE_ORD_REPAIR3UP, IMAGE_ORD_REPAIR2UP, IMAGE_ORD_REPAIR1UP), _
        Array(IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT), _
        Array(STR_DORD_REPAIR3, STR_DORD_REPAIR2, STR_DORD_REPAIR1), _
        Array(DSS_REPLEV_NEVER, DSS_REPLEV_HIGH, DSS_REPLEV_LOW)
    setOrderButton 2, ORDBUTCLASS_NORMAL, DSO_ATTACK_LEVEL, DSS_ALEV_MASK, ORD_BTYPE_RADIO, ORD_JUSTIFY_CENTER Or ORD_JUSTIFY_NEWLINE, IDORDER_ATTACK_LEVEL, 3, 0, _
        Array(IMAGE_ORD_FATWILLUP, IMAGE_ORD_RETFIREUP, IMAGE_ORD_HOLDFIREUP), _
        Array(IMAGE_ORD_FATWILLUP, IMAGE_ORD_RETFIREUP, IMAGE_ORD_HOLDFIREUP), _
        Array(IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT), _
        Array(STR_DORD_FIRE1, STR_DORD_FIRE2, STR_DORD_FIRE3), _
        Array(DSS_ALEV_ALWAYS, DSS_ALEV_ATTACKED, DSS_ALEV_NEVER)
    setOrderButton 3, ORDBUTCLASS_NORMAL, DSO_FIRE_DESIGNATOR, DSS_FIREDES_MASK, ORD_BTYPE_BOOLEAN, ORD_JUSTIFY_COMBINE, IDORDER_FIRE_DESIGNATOR, 1, 0, _
        Array(IMAGE_ORD_FIREDES_UP, 0, 0), _
        Array(IMAGE_ORD_FIREDES_UP, 0, 0), _
        Array(IMAGE_DES_HILIGHT, 0, 0), _
        Array(STR_DORD_FIREDES, 0, 0), _
        Array(DSS_FIREDES_SET, 0, 0)
    setOrderButton 4, ORDBUTCLASS_NORMAL, DSO_ACCEPT_RETREP, DSS_ACCREP_MASK, ORD_BTYPE_BOOLEAN, ORD_JUSTIFY_COMBINE, IDORDER_ACCREP, 1, 0, _
        Array(IMAGE_ORD_ACCREP_UP, 0, 0), _
        Array(IMAGE_ORD_ACCREP_UP, 0, 0), _
        Array(IMAGE_DES_HILIGHT, 0, 0), _
        Array(STR_DORD_ACCREP, 0, 0), _
        Array(DSS_ACCREP_SET, 0, 0)
    setOrderButton 5, ORDBUTCLASS_NORMAL, DSO_PATROL, DSS_PATROL_MASK, ORD_BTYPE_BOOLEAN, ORD_JUSTIFY_COMBINE, IDORDER_PATROL, 1, 0, _
        Array(IMAGE_ORD_PATROLUP, 0, 0), _
        Array(IMAGE_ORD_PATROLUP, 0, 0), _
        Array(IMAGE_DES_HILIGHT, 0, 0), _
        Array(STR_DORD_PATROL, 0, 0), _
        Array(DSS_PATROL_SET, 0, 0)
    setOrderButton 6, ORDBUTCLASS_NORMAL, DSO_CIRCLE, DSS_CIRCLE_MASK, ORD_BTYPE_BOOLEAN, ORD_JUSTIFY_COMBINE, IDORDER_CIRCLE, 1, 0, _
        Array(IMAGE_ORD_CIRCLEUP, 0, 0), _
        Array(IMAGE_ORD_CIRCLEUP, 0, 0), _
        Array(IMAGE_DES_HILIGHT, 0, 0), _
        Array(STR_DORD_CIRCLE, 0, 0), _
        Array(DSS_CIRCLE_SET, 0, 0)
    setOrderButton 7, ORDBUTCLASS_NORMAL, DSO_HALTTYPE, DSS_HALT_MASK, ORD_BTYPE_RADIO, ORD_JUSTIFY_CENTER Or ORD_JUSTIFY_NEWLINE, IDORDER_HALT_TYPE, 3, 0, _
        Array(IMAGE_ORD_PURSUEUP, IMAGE_ORD_GUARDUP, IMAGE_ORD_HALTUP), _
        Array(IMAGE_ORD_PURSUEUP, IMAGE_ORD_GUARDUP, IMAGE_ORD_HALTUP), _
        Array(IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT), _
        Array(STR_DORD_PURSUE, STR_DORD_GUARD, STR_DORD_HOLDPOS), _
        Array(DSS_HALT_PURSUE, DSS_HALT_GUARD, DSS_HALT_HOLD)
    setOrderButton 8, ORDBUTCLASS_NORMAL, DSO_RETURN_TO_LOC, DSS_RTL_MASK, ORD_BTYPE_RADIO, ORD_JUSTIFY_CENTER Or ORD_JUSTIFY_NEWLINE, IDORDER_RETURN, 3, 0, _
        Array(IMAGE_ORD_RTRUP, IMAGE_ORD_GOTOHQUP, IMAGE_ORD_EMBARKUP), _
        Array(IMAGE_ORD_RTRUP, IMAGE_ORD_GOTOHQUP, IMAGE_ORD_EMBARKUP), _
        Array(IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT), _
        Array(STR_DORD_RETREPAIR, STR_DORD_RETBASE, STR_DORD_EMBARK), _
        Array(DSS_RTL_REPAIR, DSS_RTL_BASE, DSS_RTL_TRANSPORT)
    setOrderButton 9, ORDBUTCLASS_NORMAL, DSO_RECYCLE, DSS_RECYCLE_MASK, ORD_BTYPE_BOOLEAN_DEPEND, ORD_JUSTIFY_CENTER Or ORD_JUSTIFY_NEWLINE, IDORDER_RECYCLE, 2, 0, _
        Array(IMAGE_ORD_DESTRUCT1UP, IMAGE_ORD_DESTRUCT2UP, 0), _
        Array(IMAGE_ORD_DESTRUCT1UP, IMAGE_ORD_DESTRUCT2GREY, 0), _
        Array(IMAGE_DES_HILIGHT, IMAGE_DES_HILIGHT, 0), _
        Array(STR_DORD_ARMRECYCLE, STR_DORD_RECYCLE, 0), _
        Array(DSS_RECYCLE_SET, DSS_RECYCLE_SET, 0)
    setOrderButton 10, ORDBUTCLASS_FACTORY, DSO_ASSIGN_PRODUCTION, DSS_ASSPROD_MASK, ORD_BTYPE_BOOLEAN_COMBINE, ORD_JUSTIFY_CENTER Or ORD_JUSTIFY_NEWLINE, IDORDER_ASSIGN_PRODUCTION, 5, 0, _
        Array(IMAGE_ORD_FAC1UP, IMAGE_ORD_FAC2UP, IMAGE_ORD_FAC3UP, IMAGE_ORD_FAC4UP, IMAGE_ORD_FAC5UP), _
        Array(IMAGE_ORD_FAC1UP, IMAGE_ORD_FAC2UP, IMAGE_ORD_FAC3UP, IMAGE_ORD_FAC4UP, IMAGE_ORD_FAC5UP), _
        Array(IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE), _
        Array(STR_DORD_FACTORY, STR_DORD_FACTORY, STR_DORD_FACTORY, STR_DORD_FACTORY, STR_DORD_FACTORY), _
        Array(SHL(&H1, DSS_ASSPROD_SHIFT), SHL(&H2, DSS_ASSPROD_SHIFT), SHL(&H4, DSS_ASSPROD_SHIFT), _
              SHL(&H8, DSS_ASSPROD_SHIFT), SHL(&H10, DSS_ASSPROD_SHIFT))
    setOrderButton 11, ORDBUTCLASS_CYBORGFACTORY, DSO_ASSIGN_CYBORG_PRODUCTION, DSS_ASSPROD_MASK, ORD_BTYPE_BOOLEAN_COMBINE, ORD_JUSTIFY_CENTER Or ORD_JUSTIFY_NEWLINE, IDORDER_ASSIGN_CYBORG_PRODUCTION, 5, 0, _
        Array(IMAGE_ORD_FAC1UP, IMAGE_ORD_FAC2UP, IMAGE_ORD_FAC3UP, IMAGE_ORD_FAC4UP, IMAGE_ORD_FAC5UP), _
        Array(IMAGE_ORD_FAC1UP, IMAGE_ORD_FAC2UP, IMAGE_ORD_FAC3UP, IMAGE_ORD_FAC4UP, IMAGE_ORD_FAC5UP), _
        Array(IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE), _
        Array(STR_DORD_CYBORG_FACTORY, STR_DORD_CYBORG_FACTORY, STR_DORD_CYBORG_FACTORY, STR_DORD_CYBORG_FACTORY, STR_DORD_CYBORG_FACTORY), _
        Array(SHL(&H1, DSS_ASSPROD_CYBORG_SHIFT), SHL(&H2, DSS_ASSPROD_CYBORG_SHIFT), SHL(&H4, DSS_ASSPROD_CYBORG_SHIFT), _
              SHL(&H8, DSS_ASSPROD_CYBORG_SHIFT), SHL(&H10, DSS_ASSPROD_CYBORG_SHIFT))
    setOrderButton 12, ORDBUTCLASS_VTOLFACTORY, DSO_ASSIGN_VTOL_PRODUCTION, DSS_ASSPROD_MASK, ORD_BTYPE_BOOLEAN_COMBINE, ORD_JUSTIFY_CENTER Or ORD_JUSTIFY_NEWLINE, IDORDER_ASSIGN_VTOL_PRODUCTION, 5, 0, _
        Array(IMAGE_ORD_FAC1UP, IMAGE_ORD_FAC2UP, IMAGE_ORD_FAC3UP, IMAGE_ORD_FAC4UP, IMAGE_ORD_FAC5UP), _
        Array(IMAGE_ORD_FAC1UP, IMAGE_ORD_FAC2UP, IMAGE_ORD_FAC3UP, IMAGE_ORD_FAC4UP, IMAGE_ORD_FAC5UP), _
        Array(IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE, IMAGE_ORD_FACHILITE), _
        Array(STR_DORD_VTOL_FACTORY, STR_DORD_VTOL_FACTORY, STR_DORD_VTOL_FACTORY, STR_DORD_VTOL_FACTORY, STR_DORD_VTOL_FACTORY), _
        Array(SHL(&H1, DSS_ASSPROD_VTOL_SHIFT), SHL(&H2, DSS_ASSPROD_VTOL_SHIFT), SHL(&H4, DSS_ASSPROD_VTOL_SHIFT), _
              SHL(&H8, DSS_ASSPROD_VTOL_SHIFT), SHL(&H10, DSS_ASSPROD_VTOL_SHIFT))
End Sub

' (VB) the vectors' push_back
Private Sub SelectedDroids_push_back(ByVal psDroid As Long)
    If SelectedDroids_size = 0 Then
        ReDim SelectedDroids(0 To 15)
    ElseIf SelectedDroids_size > UBound(SelectedDroids) Then
        ReDim Preserve SelectedDroids(0 To 2 * SelectedDroids_size - 1)
    End If
    SelectedDroids(SelectedDroids_size) = psDroid
    SelectedDroids_size = SelectedDroids_size + 1
End Sub

Private Sub SelectedDroids_clear()
    Erase SelectedDroids
    SelectedDroids_size = 0
End Sub

' Build a list of currently selected droids.
' Returns true if droids were selected.
'
Private Function BuildSelectedDroidList() As Boolean
    Dim it As Long, psDroid As Long
    If selectedPlayer >= MAX_PLAYERS Then
        BuildSelectedDroidList = False
        Exit Function
    End If

    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If aDroid(psDroid).selected Then
            SelectedDroids_push_back psDroid
        End If
        it = lnNext(it)
    Loop

    BuildSelectedDroidList = SelectedDroids_size <> 0
End Function

' Build a list of orders available for the selected group of droids.
' (VB) the vector's elements to orders(0 .. n - 1), n returned
Private Function buildDroidOrderList(orders() As Long) As Long
    Dim inSet(0 To NUM_ORDERS - 1) As Boolean ' std::set<AVORDER> orders: sorted by OrderIndex, each once
    Dim d As Long, OrdIndex As Long, n As Long

    For d = 0 To SelectedDroids_size - 1
        For OrdIndex = 0 To NUM_ORDERS - 1
            ' Is the order available?
            If secondarySupported(SelectedDroids(d), OrderButtons(OrdIndex).Order) Then
                ' Sort by Order index, A bubble sort? I know but it's only
                ' a small list so what the hell.
                ' Herein lies the remains of a bubble sort.
                inSet(OrdIndex) = True  ' orders.insert(avOrder)
            End If
        Next
    Next

    ' return std::vector<AVORDER>(orders.begin(), orders.end());
    Erase orders
    For OrdIndex = 0 To NUM_ORDERS - 1
        If inSet(OrdIndex) Then
            ReDim Preserve orders(0 To n)
            orders(n) = OrdIndex
            n = n + 1
        End If
    Next
    buildDroidOrderList = n
End Function

' Build a list of orders available for the selected structure.
' (VB) the vector's elements to orders(0 .. n - 1), n returned
Private Function buildStructureOrderList(ByVal psStructure As Long, orders() As Long) As Long
    Dim isFactory As Boolean
    Erase orders
    If psStructure <> 0 Then isFactory = STRUCTURE_isFactory(psStructure)
    If Not isFactory Then ASSERT_FAILURE "psStructure && psStructure->isFactory()", "src/intorder.cpp:453", "buildStructureOrderList", "BuildStructureOrderList: structure is not a factory": Exit Function

    'this can be hard-coded!
    ReDim orders(0 To 3)
    orders(0) = 0   'DSO_ATTACK_RANGE;
    orders(1) = 1   'DSO_REPAIR_LEVEL;
    orders(2) = 2   'DSO_ATTACK_LEVEL;
    orders(3) = 6   'DSO_HALTTYPE;

    buildStructureOrderList = 4
End Function

' return the state for an order for all the units selected
' if there are multiple states then don't return a state
Private Function GetSecondaryStates(ByVal sec As SECONDARY_ORDER) As SECONDARY_STATE
    Dim state As SECONDARY_STATE, currState As SECONDARY_STATE
    Dim bFirst As Boolean, d As Long

    state = 0
    bFirst = True
    If psSelectedFactory <> 0 Then
        If getFactoryState(psSelectedFactory, sec, currState) Then
            state = currState
        End If
    Else 'droids
        For d = 0 To SelectedDroids_size - 1
            currState = secondaryGetState(SelectedDroids(d), sec, ModeQueue)
            If bFirst Then
                state = currState
                bFirst = False
            ElseIf state <> currState Then
                state = 0
            End If
        Next
    End If

    GetSecondaryStates = state
End Function


Private Function GetImageWidth(ByVal ImageFile As Long, ByVal ImageID As Long) As Long
    GetImageWidth = iV_GetImageWidth(ImageFile, ImageID And &HFFFF&)
End Function

Private Function GetImageHeight(ByVal ImageFile As Long, ByVal ImageID As Long) As Long
    GetImageHeight = iV_GetImageHeight(ImageFile, ImageID And &HFFFF&)
End Function

' std::vector<UWORD> intOrderGetButtonImageIDs(SECONDARY_ORDER): (VB) the vector to result(0 .. n - 1), n returned
Public Function intOrderGetButtonImageIDs(ByVal secondaryOrder As SECONDARY_ORDER, result() As Long) As Long
    Dim i As Long, butNum As Long
    Erase result
    i = 0
    Do While i < NUM_ORDERS
        If OrderButtons(i).Order = secondaryOrder Then
            ' found it
            Exit Do
        End If
        i = i + 1
    Loop
    If i >= NUM_ORDERS Then
        intOrderGetButtonImageIDs = 0
        Exit Function
    End If
    If OrderButtons(i).NumButs > 0 Then ReDim result(0 To OrderButtons(i).NumButs - 1)
    For butNum = 0 To OrderButtons(i).NumButs - 1
        result(butNum) = OrderButtons(i).ButImageID(butNum)
    Next
    intOrderGetButtonImageIDs = OrderButtons(i).NumButs
End Function


' Add the droid order screen.
' Returns true if the form was displayed ok.
'
'changed to a BASE_OBJECT to accommodate the factories - AB 21/04/99
Public Function intAddOrder(ByVal psObj As Long) As Boolean
    Dim Animate As Boolean
    Dim State As SECONDARY_STATE
    Dim OrdIndex As Long                ' UWORD
    Dim Height As Long, NumDisplayedOrders As Long
    Dim NumButs As Long
    Dim NumJustifyButs As Long, NumCombineButs As Long, NumCombineBefore As Long
    Dim bLastCombine As Boolean, bHidden As Boolean
    Dim Droid_ As Long
    Dim psStructure As Long
    Dim parent As WIDGET, orderForm As WIDGET, sButInit As W_BUTINIT
    Dim j As Long, i As Long, selected As Boolean, first As Boolean, newHeight As Long, butState As Long, isCommandDroid As Boolean

    Animate = True

    ' Is the form already up?
    If Not widgGetFromID(psWScreen, IDORDER_FORM) Is Nothing Then
        intRemoveOrderNoAnim
        Animate = False
    End If
    ' Is the stats window up?
    If Not widgGetFromID(psWScreen, IDSTAT_FORM) Is Nothing Then
        intRemoveStatsNoAnim
        Animate = False
    End If

    If psObj <> 0 Then
        If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
            Droid_ = psObj
            psStructure = 0
        ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Then
            Droid_ = 0
            psStructure = psObj
            psSelectedFactory = psStructure
            If Not STRUCTURE_isFactory(psSelectedFactory) Then ASSERT_FAILURE "psSelectedFactory->isFactory()", "src/intorder.cpp:578", "intAddOrder", "Trying to select a %s as a factory!", objInfo(psSelectedFactory): intAddOrder = False: Exit Function
        Else
            ASSERT_FAILURE "false", "src/intorder.cpp:583", "intAddOrder", "Invalid object type"
            Droid_ = 0
            psStructure = 0
        End If
    Else
        Droid_ = 0
        psStructure = 0
    End If

    setWidgetsStatus True

    Erase AvailableOrders
    AvailableOrders_size = 0
    SelectedDroids_clear

    ' Selected droid is a command droid?
    If Droid_ <> 0 Then isCommandDroid = (aDroid(Droid_).droidType = DROID_COMMAND) ' (the && in order)
    If isCommandDroid Then
        ' displaying for a command droid - ignore any other droids
        SelectedDroids_push_back Droid_
    ElseIf psStructure <> 0 Then
        AvailableOrders_size = buildStructureOrderList(psStructure, AvailableOrders)
        If AvailableOrders_size = 0 Then
            intAddOrder = False
            Exit Function
        End If
    ' Otherwise build a list of selected droids.
    ElseIf Not BuildSelectedDroidList() Then
        ' If no droids selected then see if we were given a specific droid.
        If Droid_ <> 0 Then
            ' and put it in the list.
            SelectedDroids_push_back Droid_
        End If
    End If

    ' Build a list of orders available for the list of selected droids. - if a factory has not been selected
    If psStructure = 0 Then
        AvailableOrders_size = buildDroidOrderList(AvailableOrders)
        If AvailableOrders_size = 0 Then
            ' If no orders then return;
            intAddOrder = False
            Exit Function
        End If
    End If

    Set parent = psWScreen.psForm

    ' Create the basic form
    Set orderForm = IntFormAnimated_new(Animate) ' Do not animate the opening, if the window was already open.
    parent.attach orderForm
    orderForm.id = IDORDER_FORM
    orderForm.setCalcLayout WzFunc_new(FN_intAddOrder_layout)

    ' Add the close button.
    W_BUTINIT_construct sButInit
    sButInit.formID = IDORDER_FORM
    sButInit.id = IDORDER_CLOSE
    sButInit.x = RET_FORMWIDTH - CLOSE_WIDTH ' ORDER_WIDTH - CLOSE_WIDTH
    sButInit.y = 0
    sButInit.width = CLOSE_WIDTH
    sButInit.height = CLOSE_HEIGHT
    sButInit.pTip = "Close"
    Set sButInit.pDisplay = WzFunc_new(FN_intorder_intDisplayImageHilight)
    sButInit.UserData = PACKDWORD_TRI(0, IMAGE_CLOSEHILIGHT, IMAGE_CLOSE)
    If widgAddButton(psWScreen, sButInit) Is Nothing Then
        intAddOrder = False
        Exit Function
    End If

    W_BUTINIT_construct sButInit        ' sButInit = W_BUTINIT();
    sButInit.formID = IDORDER_FORM
    sButInit.id = IDORDER_CLOSE + 1
    Set sButInit.pDisplay = WzFunc_new(FN_intorder_intDisplayButtonHilight)
    sButInit.y = ORDER_BUTY

    Height = 0
    NumDisplayedOrders = 0

    j = 0
    Do While j < AvailableOrders_size And NumDisplayedOrders < MAX_DISPLAYABLE_ORDERS
        OrdIndex = AvailableOrders(j)

        ' Get current order state.
        State = GetSecondaryStates(OrderButtons(OrdIndex).Order)

        ' Get number of buttons.
        NumButs = OrderButtons(OrdIndex).NumButs
        ' Set actual number of buttons.
        OrderButtons(OrdIndex).AcNumButs = NumButs

        ' Handle special case for factory -> command droid assignment buttons.
        Select Case OrderButtons(OrdIndex).Class_
        Case ORDBUTCLASS_FACTORY
            NumButs = countAssignableFactories(selectedPlayer And &HFF&, FACTORY_FLAG)
        Case ORDBUTCLASS_CYBORGFACTORY
            NumButs = countAssignableFactories(selectedPlayer And &HFF&, CYBORG_FLAG)
        Case ORDBUTCLASS_VTOLFACTORY
            NumButs = countAssignableFactories(selectedPlayer And &HFF&, VTOL_FLAG)
        Case Else
        End Select

        sButInit.id = OrderButtons(OrdIndex).ButBaseID

        NumJustifyButs = NumButs
        bLastCombine = False

        Select Case OrderButtons(OrdIndex).ButJustify And ORD_JUSTIFY_MASK
        Case ORD_JUSTIFY_LEFT
            sButInit.x = ORDER_BUTX

        Case ORD_JUSTIFY_RIGHT
            sButInit.x = orderForm.width() - ORDER_BUTX - _
                         (NumJustifyButs * GetImageWidth(IntImages, OrderButtons(OrdIndex).ButImageID(0)) + _
                          (NumJustifyButs - 1) * ORDER_BUTGAP)

        Case ORD_JUSTIFY_CENTER
            sButInit.x = (orderForm.width() - _
                          (NumJustifyButs * GetImageWidth(IntImages, OrderButtons(OrdIndex).ButImageID(0)) + _
                           (NumJustifyButs - 1) * ORDER_BUTGAP)) \ 2

        Case ORD_JUSTIFY_COMBINE
            ' see how many are on this line before the button
            NumCombineBefore = 0
            For i = 0 To j - 1
                If (OrderButtons(AvailableOrders(i)).ButJustify And ORD_JUSTIFY_MASK) _
                    = ORD_JUSTIFY_COMBINE Then
                    NumCombineBefore = NumCombineBefore + 1
                End If
            Next
            NumCombineButs = (NumCombineBefore + 1) And &HFFFF&

            ' now see how many in total
            For i = j + 1 To AvailableOrders_size - 1
                If (OrderButtons(AvailableOrders(i)).ButJustify And ORD_JUSTIFY_MASK) _
                    = ORD_JUSTIFY_COMBINE Then
                    NumCombineButs = NumCombineButs + 1
                End If
            Next

            ' get position on line
            NumCombineButs = (NumCombineButs - (NumCombineBefore - (NumCombineBefore Mod ORD_MAX_COMBINE_BUTS))) And &HFFFF&

            If NumCombineButs >= ORD_MAX_COMBINE_BUTS Then
                ' the buttons will fill the line
                sButInit.x = (ORDER_BUTX + _
                              (GetImageWidth(IntImages, OrderButtons(OrdIndex).ButImageID(0)) + ORDER_BUTGAP) * NumCombineBefore)
            Else
                ' center the buttons
                sButInit.x = orderForm.width() \ 2 - _
                             (NumCombineButs * GetImageWidth(IntImages, OrderButtons(OrdIndex).ButImageID(0)) + _
                              (NumCombineButs - 1) * ORDER_BUTGAP) \ 2
                sButInit.x = (sButInit.x + _
                              (GetImageWidth(IntImages, OrderButtons(OrdIndex).ButImageID(0)) + ORDER_BUTGAP) * NumCombineBefore)
            End If

            ' see if need to start a new line of buttons
            If (NumCombineBefore + 1) = (NumCombineButs Mod ORD_MAX_COMBINE_BUTS) Then
                bLastCombine = True
            End If

        End Select

        For i = 0 To OrderButtons(OrdIndex).AcNumButs - 1
            sButInit.pTip = getDORDDescription(OrderButtons(OrdIndex).ButTips(i))
            sButInit.width = GetImageWidth(IntImages, OrderButtons(OrdIndex).ButImageID(i)) And &HFFFF&
            sButInit.height = GetImageHeight(IntImages, OrderButtons(OrdIndex).ButImageID(i)) And &HFFFF&
            sButInit.UserData = PACKDWORD_TRI(OrderButtons(OrdIndex).ButGreyID(i), _
                                              OrderButtons(OrdIndex).ButHilightID(i), _
                                              OrderButtons(OrdIndex).ButImageID(i))
            If widgAddButton(psWScreen, sButInit) Is Nothing Then
                intAddOrder = False
                Exit Function
            End If

            ' Set the default state for the button.
            Select Case OrderButtons(OrdIndex).ButType
            Case ORD_BTYPE_RADIO, ORD_BTYPE_BOOLEAN
                selected = (State And OrderButtons(OrdIndex).StateMask) = OrderButtons(OrdIndex).States(i)
                If selected Then butState = WBUT_CLICKLOCK Else butState = 0
                widgSetButtonState psWScreen, sButInit.id, butState
            Case ORD_BTYPE_BOOLEAN_DEPEND
                selected = (State And OrderButtons(OrdIndex).StateMask) = OrderButtons(OrdIndex).States(i)
                first = i = 0
                If selected Then
                    butState = WBUT_CLICKLOCK
                ElseIf first Then
                    butState = 0
                Else
                    butState = WBUT_DISABLE
                End If
                widgSetButtonState psWScreen, sButInit.id, butState
            Case ORD_BTYPE_BOOLEAN_COMBINE
                selected = (State And OrderButtons(OrdIndex).States(i)) <> 0
                If selected Then butState = WBUT_CLICKLOCK Else butState = 0
                widgSetButtonState psWScreen, sButInit.id, butState
            End Select

            ' may not add a button if the factory doesn't exist
            bHidden = False
            Select Case OrderButtons(OrdIndex).Class_
            Case ORDBUTCLASS_FACTORY
                If Not checkFactoryExists(selectedPlayer, FACTORY_FLAG, i) Then
                    widgHide psWScreen, sButInit.id
                    bHidden = True
                End If
            Case ORDBUTCLASS_CYBORGFACTORY
                If Not checkFactoryExists(selectedPlayer, CYBORG_FLAG, i) Then
                    widgHide psWScreen, sButInit.id
                    bHidden = True
                End If
            Case ORDBUTCLASS_VTOLFACTORY
                If Not checkFactoryExists(selectedPlayer, VTOL_FLAG, i) Then
                    widgHide psWScreen, sButInit.id
                    bHidden = True
                End If
            Case Else
            End Select

            If Not bHidden Then

                sButInit.x = (sButInit.x + sButInit.width + ORDER_BUTGAP)
            End If
            sButInit.id = sButInit.id + 1
        Next

        If ((OrderButtons(OrdIndex).ButJustify And ORD_JUSTIFY_MASK) <> ORD_JUSTIFY_COMBINE) Or _
            bLastCombine Then
            sButInit.y = (sButInit.y + sButInit.height + ORDER_BUTGAP)
            Height = (Height + sButInit.height + ORDER_BUTGAP) And &HFFFF&
        End If
        NumDisplayedOrders = NumDisplayedOrders + 1
        j = j + 1
    Loop

    ' Now we know how many orders there are we can resize the form accordingly.
    newHeight = Height + CLOSE_HEIGHT + ORDER_BUTGAP
    orderForm.setCalcLayout WzFunc_new(FN_intAddOrder_newHeightLayout, newHeight)

    OrderUp = True
    setSecondaryWindowUp True

    intAddOrder = True
End Function


' Do any housekeeping for the droid order screen.
' Any droids that die now get set to NULL - John.
' No droids being selected no longer removes the screen,
' this lets the screen work with command droids - John.
Public Sub intRunOrder()
    Dim d As Long, n As Long
    ' Check to see if there all dead or unselected.
    For d = 0 To SelectedDroids_size - 1
        If aDroid(SelectedDroids(d)).died <> 0 Then
            SelectedDroids(d) = 0
        End If
    Next
    ' Remove any NULL pointers from SelectedDroids.
    n = 0
    For d = 0 To SelectedDroids_size - 1
        If SelectedDroids(d) <> 0 Then
            SelectedDroids(n) = SelectedDroids(d)
            n = n + 1
        End If
    Next
    SelectedDroids_size = n

    If psSelectedFactory <> 0 Then
        If aStruct(psSelectedFactory).died <> 0 Then
            psSelectedFactory = 0
        End If
    End If

    ' If all dead then remove the screen.
    ' If droids no longer selected then remove screen.
    If SelectedDroids_size = 0 Then
        ' might have a factory selected
        If psSelectedFactory = 0 Then
            intRemoveOrder
            Exit Sub
        End If
    End If
End Sub


' Set the secondary order state for all currently selected droids. And Factory (if one selected)
' Returns true if successful.
'
Private Function SetSecondaryState(ByVal sec As SECONDARY_ORDER, ByVal State As Long) As Boolean
    Dim d As Long
    ' This code is similar to kfsf_SetSelectedDroidsState() in keybind.cpp. Unfortunately, it seems hard to un-duplicate the code.
    For d = 0 To SelectedDroids_size - 1
        If SelectedDroids(d) <> 0 Then
            'Only set the state if it's not a transporter.
            If Not DROID_isTransporter(SelectedDroids(d)) Then
                If Not secondarySetState(SelectedDroids(d), sec, State) Then
                    SetSecondaryState = False
                    Exit Function
                End If
            End If
        End If
    Next

    ' set the Factory settings
    If psSelectedFactory <> 0 Then
        If Not setFactoryState(psSelectedFactory, sec, State) Then
            SetSecondaryState = False
            Exit Function
        End If
    End If

    SetSecondaryState = True
End Function


' Process the droid order screen.
'
Public Sub intProcessOrder(ByVal id As Long)
    Dim i As Long                       ' UWORD
    Dim OrdIndex As Long
    Dim BaseID As Long                  ' UDWORD
    Dim StateIndex As Long
    Dim CombineState As Long

    If id = IDORDER_CLOSE Then
        intRemoveOrder
        If intMode = INT_ORDER Then
            intMode = INT_NORMAL
        Else
            intMode = INT_OBJECT
        End If
        Exit Sub
    End If

    For OrdIndex = 0 To NUM_ORDERS - 1
        BaseID = OrderButtons(OrdIndex).ButBaseID

        Select Case OrderButtons(OrdIndex).ButType
        Case ORD_BTYPE_RADIO
            If (id >= BaseID) And (id < BaseID + OrderButtons(OrdIndex).AcNumButs) Then
                StateIndex = id - BaseID

                For i = 0 To OrderButtons(OrdIndex).AcNumButs - 1
                    widgSetButtonState psWScreen, BaseID + i, 0
                Next
                If SetSecondaryState(OrderButtons(OrdIndex).Order, _
                                     OrderButtons(OrdIndex).States(StateIndex) And OrderButtons(OrdIndex).StateMask) Then
                    widgSetButtonState psWScreen, id, WBUT_CLICKLOCK
                End If
            End If

        Case ORD_BTYPE_BOOLEAN
            If (id >= BaseID) And (id < BaseID + OrderButtons(OrdIndex).AcNumButs) Then
                StateIndex = id - BaseID

                If (widgGetButtonState(psWScreen, id) And WBUT_CLICKLOCK) <> 0 Then
                    widgSetButtonState psWScreen, id, 0
                    SetSecondaryState OrderButtons(OrdIndex).Order, 0
                Else
                    widgSetButtonState psWScreen, id, WBUT_CLICKLOCK
                    SetSecondaryState OrderButtons(OrdIndex).Order, _
                                      OrderButtons(OrdIndex).States(StateIndex) And OrderButtons(OrdIndex).StateMask
                End If

            End If

        Case ORD_BTYPE_BOOLEAN_DEPEND
            ' Toggle the state of this button.
            If id = BaseID Then
                If (widgGetButtonState(psWScreen, id) And WBUT_CLICKLOCK) <> 0 Then
                    widgSetButtonState psWScreen, id, 0
                    ' Disable the dependent button.
                    widgSetButtonState psWScreen, id + 1, WBUT_DISABLE
                Else
                    widgSetButtonState psWScreen, id, WBUT_CLICKLOCK
                    ' Enable the dependent button.
                    widgSetButtonState psWScreen, id + 1, 0
                End If
            End If
            If (id > BaseID) And (id < BaseID + OrderButtons(OrdIndex).AcNumButs) Then
                ' If the previous button is down ( armed )..
                If (widgGetButtonState(psWScreen, id - 1) And WBUT_CLICKLOCK) <> 0 Then
                    ' Toggle the state of this button.
                    If (widgGetButtonState(psWScreen, id) And WBUT_CLICKLOCK) <> 0 Then
                        widgSetButtonState psWScreen, id, 0
                        SetSecondaryState OrderButtons(OrdIndex).Order, 0
                    Else
                        StateIndex = id - BaseID

                        widgSetButtonState psWScreen, id, WBUT_CLICKLOCK
                        SetSecondaryState OrderButtons(OrdIndex).Order, _
                                          OrderButtons(OrdIndex).States(StateIndex) And OrderButtons(OrdIndex).StateMask
                    End If
                End If
            End If
        Case ORD_BTYPE_BOOLEAN_COMBINE
            If (id >= BaseID) And (id < BaseID + OrderButtons(OrdIndex).AcNumButs) Then
                ' Toggle the state of this button.
                If (widgGetButtonState(psWScreen, id) And WBUT_CLICKLOCK) <> 0 Then
                    widgSetButtonState psWScreen, id, 0
                Else
                    widgSetButtonState psWScreen, id, WBUT_CLICKLOCK
                End If

                ' read the state of all the buttons to get the final state
                CombineState = 0
                For StateIndex = 0 To OrderButtons(OrdIndex).AcNumButs - 1
                    If (widgGetButtonState(psWScreen, BaseID + StateIndex) And WBUT_CLICKLOCK) <> 0 Then
                        CombineState = CombineState Or OrderButtons(OrdIndex).States(StateIndex)
                    End If
                Next

                ' set the final state
                SetSecondaryState OrderButtons(OrdIndex).Order, _
                                  CombineState And OrderButtons(OrdIndex).StateMask
            End If
        End Select
    Next
End Sub


' check whether the order list has changed
Private Function CheckObjectOrderList() As Boolean
    Dim NewAvailableOrders() As Long, n As Long, k As Long

    If psSelectedFactory <> 0 Then
        n = buildStructureOrderList(psSelectedFactory, NewAvailableOrders)
    Else
        n = buildDroidOrderList(NewAvailableOrders)
    End If

    ' now check that all the orders are the same
    If n <> AvailableOrders_size Then
        CheckObjectOrderList = False
        Exit Function
    End If
    For k = 0 To n - 1
        If NewAvailableOrders(k) <> AvailableOrders(k) Then
            CheckObjectOrderList = False
            Exit Function
        End If
    Next
    CheckObjectOrderList = True
End Function

Private Function intRefreshOrderButtons() As Boolean
    Dim State As SECONDARY_STATE
    Dim OrdIndex As Long
    Dim id As Long
    Dim j As Long, i As Long, selected As Boolean, first As Boolean, butState As Long

    j = 0
    Do While j < AvailableOrders_size And j < MAX_DISPLAYABLE_ORDERS
        OrdIndex = AvailableOrders(j)

        ' Get current order state.
        State = GetSecondaryStates(OrderButtons(OrdIndex).Order)

        ' Set actual number of buttons.
        OrderButtons(OrdIndex).AcNumButs = OrderButtons(OrdIndex).NumButs

        id = OrderButtons(OrdIndex).ButBaseID
        For i = 0 To OrderButtons(OrdIndex).AcNumButs - 1
            ' Set the state for the button.
            Select Case OrderButtons(OrdIndex).ButType
            Case ORD_BTYPE_RADIO, ORD_BTYPE_BOOLEAN
                selected = (State And OrderButtons(OrdIndex).StateMask) = OrderButtons(OrdIndex).States(i)
                If selected Then butState = WBUT_CLICKLOCK Else butState = 0
                widgSetButtonState psWScreen, id, butState
            Case ORD_BTYPE_BOOLEAN_DEPEND
                selected = (State And OrderButtons(OrdIndex).StateMask) = OrderButtons(OrdIndex).States(i)
                first = i = 0
                If selected Then
                    butState = WBUT_CLICKLOCK
                ElseIf first Then
                    butState = 0
                Else
                    butState = WBUT_DISABLE
                End If
                widgSetButtonState psWScreen, id, butState
            Case ORD_BTYPE_BOOLEAN_COMBINE
                selected = (State And OrderButtons(OrdIndex).States(i)) <> 0
                If selected Then butState = WBUT_CLICKLOCK Else butState = 0
                widgSetButtonState psWScreen, id, butState
            End Select

            id = id + 1
        Next
        j = j + 1
    Loop

    intRefreshOrderButtons = True
End Function


' Call to refresh the Order screen, ie when a droids boards it.
'
Public Function intRefreshOrder() As Boolean
    Dim Ret As Boolean
    ' Is the Order screen up?
    If (intMode = INT_ORDER) And _
       (Not widgGetFromID(psWScreen, IDORDER_FORM) Is Nothing) Then

        SelectedDroids_clear

        ' check for factory selected first
        If psSelectedFactory = 0 Then
            If Not BuildSelectedDroidList() Then
                ' no units selected - quit
                intRemoveOrder
                intRefreshOrder = True
                Exit Function
            End If
        End If

        ' if the orders haven't changed, just reset the state
        If CheckObjectOrderList() Then
            Ret = intRefreshOrderButtons()
        Else
            ' Refresh it by re-adding it.
            Ret = intAddOrder(0)
            If Ret = False Then
                intMode = INT_NORMAL
            End If
        End If
        intRefreshOrder = Ret
        Exit Function
    End If

    intRefreshOrder = True
End Function


' Remove the droids order screen with animation.
'
Public Sub intRemoveOrder()
    Dim form_ As WIDGET
    widgDelete psWScreen, IDORDER_CLOSE
    setSecondaryWindowUp False

    ' Start the window close animation.
    Set form_ = widgGetFromID(psWScreen, IDORDER_FORM)  ' (IntFormAnimated *)
    If Not form_ Is Nothing Then
        IntFormAnimated_closeAnimateDelete form_
        OrderUp = False
        SelectedDroids_clear
        psSelectedFactory = 0
    End If
End Sub


' Remove the droids order screen without animation.
'
Public Sub intRemoveOrderNoAnim()
    widgDelete psWScreen, IDORDER_CLOSE
    widgDelete psWScreen, IDORDER_FORM
    setSecondaryWindowUp False
    OrderUp = False
    SelectedDroids_clear
    psSelectedFactory = 0
End Sub

'new function added to bring up the RMB order form for Factories as well as droids
Public Sub intAddFactoryOrder(ByVal psStructure As Long)
    If Not OrderUp Then
        intResetScreen False
        intAddOrder psStructure
        intMode = INT_ORDER
    Else
        intAddOrder psStructure
    End If
End Sub

'--------------------------------------------------------------------------
' (VB) the lambdas

Public Function intorder_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET
    Select Case f.id
    Case FN_intAddOrder_layout          ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        psWidget.setGeometry ORDER_X, ORDER_Y, RET_FORMWIDTH, ORDER_HEIGHT ' ORDER_WIDTH: RET_FORMWIDTH

    Case FN_intAddOrder_newHeightLayout ' [newHeight](WIDGET *psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry psWidget.x(), 318 + E_H - f.cap(0), psWidget.width(), f.cap(0) ' ORDER_BOTTOMY - newHeight

    Case FN_intorder_intDisplayImageHilight
        intDisplayImageHilight args(0), args(1), args(2)

    Case FN_intorder_intDisplayButtonHilight
        intDisplayButtonHilight args(0), args(1), args(2)
    End Select
End Function
