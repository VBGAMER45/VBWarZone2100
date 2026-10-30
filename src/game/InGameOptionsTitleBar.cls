VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InGameOptionsTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class InGameOptionsTitleBar
' src/screens/ingameopscreen.cpp: InGameOptionsTitleBar's fields (ingameopscreen_c).
Option Explicit

Public m_closeButton As WIDGET          ' std::shared_ptr<W_BUTTON> (an InGameOptionsTitleBarButton)
Public m_onCloseHandler As WzFunc       ' OnCloseFunc (std::function<void ()>)
