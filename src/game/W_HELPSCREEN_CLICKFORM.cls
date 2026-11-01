VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_HELPSCREEN_CLICKFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_HELPSCREEN_CLICKFORM
' src/screens/helpscreen.cpp: the fields of class W_HELPSCREEN_CLICKFORM (helpscreen_c); the initialisers are its _construct's.
Option Explicit

Public backgroundColor As Long          ' PIELIGHT = pal_RGBA(0, 0, 0, 125)
Public onClickedFunc As WzFunc          ' std::function<void ()>
Public onCancelPressed As WzFunc        ' std::function<void ()>
Public registeredWidgets As Collection  ' RegisteredHelpWidgetsMap: "p<ObjPtr(widget)>" -> Array(widget, help, helpTriggerWidget)
Public currentHelpWidget As Long        ' std::weak_ptr<WIDGET> (a HelpTriggerWidget): its widgweak_vb serial
Public currentHelpPanel As WIDGET       ' std::shared_ptr<W_FORM>
Public weakSerial As Long               ' (VB) its widgweak_vb serial (its weak pointers)
