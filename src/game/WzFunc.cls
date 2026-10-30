VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) a lambda, std::function or function pointer of the UI code (docs/CONVENTIONS.md,
' "Widgets"): id is a module's FN_ constant (the generated WzFunc_dispatch calls that module's
' <module>_fn with the WzFunc and the call's arguments), cap(0..7) are the captured values (an
' object, or a value copied as C++'s [=] captures it; a [&] capture of a local becomes an object
' that both sides share, or the module variable the lambda reads). Nothing is an empty
' std::function / nullptr. widgbase_h's WzFunc_new(id, captures...) makes one, WzFunc_call(f,
' args...) calls it.
Option Explicit

Public id As Long
Private cap_(0 To 7) As Variant

Public Property Get cap(ByVal i As Long) As Variant
    If IsObject(cap_(i)) Then Set cap = cap_(i) Else cap = cap_(i)
End Property

Public Property Let cap(ByVal i As Long, ByVal v As Variant)
    cap_(i) = v
End Property

Public Property Set cap(ByVal i As Long, ByVal v As Variant)
    Set cap_(i) = v
End Property
