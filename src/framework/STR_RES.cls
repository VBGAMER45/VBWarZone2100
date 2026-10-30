VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "STR_RES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/framework/strres.cpp: struct STR_RES, a string resource (opaque to
' its users: they pass it to the strres functions). Made by strresCreate.
Option Explicit

Private mIDTreap As Long                ' struct TREAP_NODE **psIDTreap: the treap to store string identifiers

Friend Property Get psIDTreap() As Long
    psIDTreap = mIDTreap
End Property

Friend Property Let psIDTreap(ByVal value As Long)
    mIDTreap = value
End Property
