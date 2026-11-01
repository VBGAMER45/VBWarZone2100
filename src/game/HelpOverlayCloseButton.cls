VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HelpOverlayCloseButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class HelpOverlayCloseButton
' src/screens/helpscreen.cpp: the fields of class HelpOverlayCloseButton (helpscreen_c); the initialisers are its _construct's.
Option Explicit

Public downStateMask As Long            ' unsigned = WBUT_DOWN | WBUT_LOCK | WBUT_CLICKLOCK
Public greyStateMask As Long            ' unsigned = WBUT_DISABLE
Public text As WzText                   ' private: WzText text
