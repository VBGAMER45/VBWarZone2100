Attribute VB_Name = "widgweak_vb"
' (VB) std::weak_ptr to the widgets of the lib/widget classes that C++ holds
' weakly (DropdownWidget, DropdownItemWrapper, PopoverWidget,
' PopoverMenuWidget, ScrollableTableWidget): such a widget registers itself in
' its constructor (WeakWidget_register, its serial kept in its data) and
' unregisters in its destructor; a weak pointer is the serial, and
' WeakWidget_lock(serial) is weak_ptr::lock(): the widget, or Nothing once it
' is gone (an address alone could outlive the object). Serial 0 is an empty
' weak_ptr.
Option Explicit

Private live As Collection              ' "s<serial>" -> ObjPtr of the widget
Private nextSerial As Long

Public Function WeakWidget_register(ByVal w As WIDGET) As Long
    If live Is Nothing Then Set live = New Collection
    nextSerial = nextSerial + 1
    live.Add ObjPtr(w), "s" & nextSerial
    WeakWidget_register = nextSerial
End Function

Public Sub WeakWidget_unregister(ByVal serial As Long)
    If serial = 0 Or live Is Nothing Then Exit Sub
    On Error Resume Next
    live.Remove "s" & serial
    Err.Clear
End Sub

Public Function WeakWidget_lock(ByVal serial As Long) As WIDGET
    Dim p As Long
    If serial = 0 Or live Is Nothing Then Exit Function
    On Error Resume Next
    p = live.Item("s" & serial)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    Set WeakWidget_lock = WIDGET_fromPtr(p)
End Function
