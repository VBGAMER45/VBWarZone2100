VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseObjectsStatsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/objects_stats.h: class BaseObjectsStatsController : public BaseStatsController, public
' BaseObjectsController. Its only member is the field highlightedObjectStats (a BASE_STATS *
' handle, 0 = nullptr), here a property every implementer keeps; its methods are objects_stats_c's
' BaseObjectsStatsController_<method>. An implementer also Implements both bases.
Option Explicit

Public Property Get highlightedObjectStats() As Long
End Property

Public Property Let highlightedObjectStats(ByVal v As Long)
End Property
