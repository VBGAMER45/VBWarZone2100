VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseStatsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/objects_stats.h: class BaseStatsController's virtual methods, an interface (see
' BaseObjectsController); displayStatsForm and scheduleDisplayStatsForm are objects_stats_c's
' BaseStatsController_<method>. getStatsAt returns a BASE_STATS * handle (0 = nullptr).
Option Explicit

Public Function statsSize() As Long
End Function

' std::shared_ptr<StatsForm>
Public Function makeStatsForm() As WIDGET
End Function

Public Function getStatsAt(ByVal index As Long) As Long
End Function
