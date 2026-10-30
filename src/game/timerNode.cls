VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "timerNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/qtscript.h, qtscript.cpp: struct scripting_engine::timerNode, a script
' timer. The engine keeps std::shared_ptr<timerNode>s: a VB object reference
' is the shared_ptr (updateScripts' run list keeps a removed node alive, as in
' C++). Class_Initialize is the default constructor; the other one is
' qtscript_c's timerNode_new. TimerFunc function is function_ (an ITimerFunc;
' "Function" is a VB keyword); uniqueTimerID timerID is a Double; the
' std::unique_ptr<timerAdditionalData> is an Object (Nothing = nullptr).
' ~timerNode calls additionalTimerFuncParam->onTimerDelete: the base class's
' is empty and no backend overrides it (quickjs_timer_additionaldata does
' not), so the destructor has nothing to do here. The move constructor and
' swap are not needed (a reference is moved instead).
Option Explicit

Public timerID As Double                ' uniqueTimerID
Public function_ As ITimerFunc          ' TimerFunc function
Public timerName As String
Public instance As ChakraScriptingInstance ' wzapi::scripting_instance*
Public baseobj As Long
Public baseobjtype As Long              ' OBJECT_TYPE
Public additionalTimerFuncParam As Object ' std::unique_ptr<timerAdditionalData>
Public frameTime As Long
Public ms As Long
Public player As Long
Public calls As Long
Public typ As timerType                 ' timerType type

' timerNode() : instance(nullptr), baseobjtype(OBJ_NUM_TYPES), additionalTimerFuncParam(nullptr) {}
Private Sub Class_Initialize()
    Set instance = Nothing
    baseobjtype = OBJ_NUM_TYPES
    Set additionalTimerFuncParam = Nothing
End Sub

' operator==
Public Function equals(ByVal t As timerNode) As Boolean
    equals = (timerID = t.timerID) And (timerName = t.timerName) And (player = t.player)
End Function
