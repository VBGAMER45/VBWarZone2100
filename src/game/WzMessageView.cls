VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMessageView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzMessageView
' src/intelmap.cpp: WzMessageView's fields (intelmap_c).
Option Explicit

Public onCloseFunc As WzFunc            ' std::function<void ()> (VB: called with the view, see FN_intelmap_closeViewTask)
Public closeButton As WIDGET            ' std::shared_ptr<W_BUTTON>
