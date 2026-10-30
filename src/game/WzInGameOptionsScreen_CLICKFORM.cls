VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzInGameOptionsScreen_CLICKFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzInGameOptionsScreen_CLICKFORM
' src/screens/ingameopscreen.cpp: WzInGameOptionsScreen_CLICKFORM's fields (ingameopscreen_c); the
' initialisers are WzInGameOptionsScreen_CLICKFORM_initMembers'.
Option Explicit

Public backgroundColor As Long          ' PIELIGHT = pal_RGBA(0, 0, 0, 80)
Public onClickedFunc As WzFunc          ' std::function<void ()>
Public onCancelPressed As WzFunc        ' std::function<void ()>
Public wrapperForm As WIDGET            ' std::shared_ptr<WIDGET> (an InGameOptionsBaseForm)
Public edgeMargin As Long               ' const int = 10
