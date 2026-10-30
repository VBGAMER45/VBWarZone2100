Attribute VB_Name = "orderdef_h"
' src/orderdef.h: the order related structures. DroidOrder's
' constructors are DroidOrder_<form> functions; a fresh DROID_ORDER_DATA
' is DroidOrder(DORDER_NONE) once psStats is -1 (DroidOrder_construct).
Option Explicit

' All the possible droid orders.
Public Enum DroidOrderType
    DORDER_NONE             ' no order set.
    DORDER_STOP             ' stop the current order.
    DORDER_MOVE             ' 2 - move to a location.
    DORDER_ATTACK           ' attack an enemy.
    DORDER_BUILD            ' 4 - build a structure.
    DORDER_HELPBUILD        ' help to build a structure.
    DORDER_LINEBUILD        ' 6 - build a number of structures in a row (walls + bridges).
    DORDER_DEMOLISH         ' demolish a structure.
    DORDER_REPAIR           ' 8 - repair a structure.
    DORDER_OBSERVE          ' keep a target in sensor view.
    DORDER_FIRESUPPORT      ' 10 - attack whatever the linked sensor droid attacks.
    DORDER_UNUSED_4
    DORDER_UNUSED_2
    DORDER_RTB              ' return to base.
    DORDER_RTR              ' 14 - return to repair at any repair facility
    DORDER_UNUSED_5
    DORDER_EMBARK           ' 16 - board a transporter.
    DORDER_DISEMBARK        ' get off a transporter.
    DORDER_ATTACKTARGET     ' 18 - a suggestion to attack something i.e. the target was chosen because the droid could see it.
    DORDER_COMMANDERSUPPORT ' Assigns droid to the target commander.
    DORDER_BUILDMODULE      ' 20 - build a module (power, research or factory).
    DORDER_RECYCLE          ' return to factory to be recycled.
    DORDER_TRANSPORTOUT     ' 22 - offworld transporter order.
    DORDER_TRANSPORTIN      ' onworld transporter order.
    DORDER_TRANSPORTRETURN  ' 24 - transporter return after unloading.
    DORDER_GUARD            ' guard a structure.
    DORDER_DROIDREPAIR      ' 26 - repair a droid.
    DORDER_RESTORE          ' restore resistance points for a structure.
    DORDER_SCOUT            ' 28 - same as move, but stop if an enemy is seen.
    DORDER_UNUSED_3
    DORDER_UNUSED
    DORDER_PATROL           ' move between two way points.
    DORDER_REARM            ' 32 - order a vtol to rearming pad.
    DORDER_RECOVER          ' pick up an artifact.
    DORDER_UNUSED_6
    DORDER_RTR_SPECIFIED    ' return to repair at a specified repair center.
    DORDER_CIRCLE = 40      ' circles target location and engage.
    DORDER_HOLD             ' hold position until given next order.
End Enum

' All the possible secondary orders for droids.
Public Enum SECONDARY_ORDER
    DSO_UNUSED
    DSO_ATTACK_RANGE
    DSO_REPAIR_LEVEL
    DSO_ATTACK_LEVEL
    DSO_ASSIGN_PRODUCTION
    DSO_ASSIGN_CYBORG_PRODUCTION
    DSO_CLEAR_PRODUCTION
    DSO_RECYCLE
    DSO_PATROL
    DSO_HALTTYPE
    DSO_RETURN_TO_LOC
    DSO_FIRE_DESIGNATOR
    DSO_ASSIGN_VTOL_PRODUCTION
    DSO_CIRCLE
    DSO_ACCEPT_RETREP
End Enum

' All associated secondary states of the secondary orders.
Public Enum SECONDARY_STATE
    DSS_NONE = &H0
    DSS_ARANGE_SHORT = &H1
    DSS_ARANGE_LONG = &H2
    DSS_ARANGE_OPTIMUM = &H3
    DSS_REPLEV_LOW = &H4
    DSS_REPLEV_HIGH = &H8
    DSS_REPLEV_NEVER = &HC
    DSS_ALEV_ALWAYS = &H10
    DSS_ALEV_ATTACKED = &H20
    DSS_ALEV_NEVER = &H30
    DSS_HALT_HOLD = &H40
    DSS_HALT_GUARD = &H80
    DSS_HALT_PURSUE = &HC0
    DSS_RECYCLE_SET = &H100
    DSS_ASSPROD_START = &H200
    DSS_ACCREP_SET = &H400
    DSS_ASSPROD_MID = &H2000
    DSS_ASSPROD_END = &H40000
    DSS_RTL_REPAIR = &H80000
    DSS_RTL_BASE = &H100000
    DSS_RTL_TRANSPORT = &H200000
    DSS_PATROL_SET = &H400000
    DSS_CIRCLE_SET = &H400100
    DSS_FIREDES_SET = &H800000
End Enum

' masks for the secondary order state.
Public Const DSS_ARANGE_MASK As Long = &H3
Public Const DSS_REPLEV_MASK As Long = &HC
Public Const DSS_ALEV_MASK As Long = &H30
Public Const DSS_HALT_MASK As Long = &HC0
Public Const DSS_RECYCLE_MASK As Long = &H100
Public Const DSS_ASSPROD_MASK As Long = &H1F07FE00
Public Const DSS_ASSPROD_FACT_MASK As Long = &H3E00
Public Const DSS_ASSPROD_CYB_MASK As Long = &H7C000
Public Const DSS_ASSPROD_VTOL_MASK As Long = &H1F000000
Public Const DSS_ASSPROD_SHIFT As Long = 9
Public Const DSS_ASSPROD_CYBORG_SHIFT As Long = DSS_ASSPROD_SHIFT + 5
Public Const DSS_ASSPROD_VTOL_SHIFT As Long = 24
Public Const DSS_RTL_MASK As Long = &H380000
Public Const DSS_PATROL_MASK As Long = &H400000
Public Const DSS_FIREDES_MASK As Long = &H800000
Public Const DSS_CIRCLE_MASK As Long = &H400100
Public Const DSS_ACCREP_MASK As Long = &H400

Public Enum RTR_DATA_TYPE
    RTR_TYPE_NO_RESULT
    RTR_TYPE_REPAIR_FACILITY
    RTR_TYPE_DROID
    RTR_TYPE_HQ
    ' If modifying, only *add* entries at end! (or savegame loading will break)
End Enum

' Struct that stores data of an order (DroidOrder, typedef DROID_ORDER_DATA).
Public Type DROID_ORDER_DATA
    typ As DroidOrderType               ' the actual order.
    pos As Vector2i                     ' the order's position.
    pos2 As Vector2i                    ' the order's second position, in case those exist.
    direction As Long                   ' uint16_t: the order's direction, in case it exist.
    index As Long                       ' uint32_t: Module index, with DORDER_BUILDMODULE.
    rtrType As RTR_DATA_TYPE            ' specifies where to repair.
    psObj As Long                       ' BASE_OBJECT *: the order's target, in case it exist.
    psStats As Long                     ' STRUCTURE_STATS *: order structure stats (-1: nullptr).
End Type

' explicit DroidOrder(DroidOrderType type = DORDER_NONE)
Public Function DroidOrder_(Optional ByVal typ As DroidOrderType = DORDER_NONE) As DROID_ORDER_DATA
    DroidOrder_.typ = typ
    DroidOrder_.rtrType = RTR_TYPE_NO_RESULT
    DroidOrder_.psObj = 0
    DroidOrder_.psStats = -1
End Function

' (VB) a member DroidOrder's default construction
Public Sub DroidOrder_construct(o As DROID_ORDER_DATA)
    o = DroidOrder_(DORDER_NONE)
End Sub

' DroidOrder(DroidOrderType type, Vector2i pos)
Public Function DroidOrder_pos(ByVal typ As DroidOrderType, pos As Vector2i) As DROID_ORDER_DATA
    DroidOrder_pos = DroidOrder_(typ)
    DroidOrder_pos.pos = pos
End Function

' DroidOrder(DroidOrderType type, Vector2i pos, RTR_DATA_TYPE rtrType)
Public Function DroidOrder_posRtr(ByVal typ As DroidOrderType, pos As Vector2i, ByVal rtrType As RTR_DATA_TYPE) As DROID_ORDER_DATA
    DroidOrder_posRtr = DroidOrder_(typ)
    DroidOrder_posRtr.pos = pos
    DroidOrder_posRtr.rtrType = rtrType
End Function

' DroidOrder(DroidOrderType type, STRUCTURE_STATS *psStats, Vector2i pos, uint16_t direction)
Public Function DroidOrder_stats(ByVal typ As DroidOrderType, ByVal psStats As Long, pos As Vector2i, ByVal direction As Long) As DROID_ORDER_DATA
    DroidOrder_stats = DroidOrder_(typ)
    DroidOrder_stats.pos = pos
    DroidOrder_stats.direction = direction And &HFFFF&
    DroidOrder_stats.psStats = psStats
End Function

' DroidOrder(DroidOrderType type, STRUCTURE_STATS *psStats, Vector2i pos, Vector2i pos2, uint16_t direction)
Public Function DroidOrder_stats2(ByVal typ As DroidOrderType, ByVal psStats As Long, pos As Vector2i, pos2 As Vector2i, ByVal direction As Long) As DROID_ORDER_DATA
    DroidOrder_stats2 = DroidOrder_(typ)
    DroidOrder_stats2.pos = pos
    DroidOrder_stats2.pos2 = pos2
    DroidOrder_stats2.direction = direction And &HFFFF&
    DroidOrder_stats2.psStats = psStats
End Function

' DroidOrder(DroidOrderType type, BASE_OBJECT *psObj)
Public Function DroidOrder_obj(ByVal typ As DroidOrderType, ByVal psObj As Long) As DROID_ORDER_DATA
    DroidOrder_obj = DroidOrder_(typ)
    DroidOrder_obj.psObj = psObj
End Function

' DroidOrder(DroidOrderType type, BASE_OBJECT *psObj, RTR_DATA_TYPE rtrType)
Public Function DroidOrder_objRtr(ByVal typ As DroidOrderType, ByVal psObj As Long, ByVal rtrType As RTR_DATA_TYPE) As DROID_ORDER_DATA
    DroidOrder_objRtr = DroidOrder_(typ)
    DroidOrder_objRtr.psObj = psObj
    DroidOrder_objRtr.rtrType = rtrType
End Function

' DroidOrder(DroidOrderType type, BASE_OBJECT *psObj, uint32_t index)
Public Function DroidOrder_objIndex(ByVal typ As DroidOrderType, ByVal psObj As Long, ByVal index As Long) As DROID_ORDER_DATA
    DroidOrder_objIndex = DroidOrder_(typ)
    DroidOrder_objIndex.psObj = psObj
    DroidOrder_objIndex.index = index
End Function
