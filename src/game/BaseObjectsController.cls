VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseObjectsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/objects_stats.h: class BaseObjectsController's virtual methods, an interface: every
' controller Implements it (objects_stats_c's header). Its non-virtual methods are objects_stats_c's
' BaseObjectsController_<method>(controller, ...). A BASE_OBJECT * is its pool handle and a
' BASE_STATS * its stats handle (0 = nullptr); size_t is a Long.
Option Explicit

Public Function objectsSize() As Long
End Function

Public Function getObjectAt(ByVal index As Long) As Long
End Function

Public Function getObjectStatsAt(ByVal index As Long) As Long
End Function

' iteration: std::function<bool (BASE_OBJECT *)> (args: the object; returns a Boolean)
Public Function findObject(ByVal iteration As WzFunc) As Boolean
End Function

Public Sub refresh()
End Sub

Public Function showInterface() As Boolean
End Function

' (VB) the C++ default (BaseObjectsController::prepareToClose) is objects_stats_c's BaseObjectsController_def_prepareToClose
Public Sub prepareToClose()
End Sub

Public Sub clearData()
End Sub

Public Function getHighlightedObject() As Long
End Function

Public Sub setHighlightedObject(ByVal object_ As Long, ByVal jumpToHighlightedStatsObject As Boolean)
End Sub

Public Function getQueuedJumpToHighlightedStatsObject() As Boolean
End Function

Public Sub clearQueuedJumpToHighlightedStatsObject()
End Sub
