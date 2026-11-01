VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzJoiningPasswordPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzJoiningPasswordPrompt
' src/screens/joiningscreen.cpp: the fields of class WzJoiningPasswordPrompt (joiningscreen_c)
Option Explicit

Public passwordBox As WIDGET            ' std::shared_ptr<W_EDITBOX>
Public okayButton As WIDGET             ' std::shared_ptr<WzJoiningActionButton>
Public onSubmit As WzFunc               ' OnSubmitFunc onSubmit;
Public invalidPasswordPlaceholder As Boolean ' = false

Public weakSerial As Long               ' (VB) its std::weak_ptr<WzJoiningPasswordPrompt>s: the widgweak_vb serial
