VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzStartHostingButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/starthostingbutton.h: the fields of class WzStartHostingButton
' (VB) data of widget class WzStartHostingButton
Option Explicit

Public multiOptionsTitleUI As Long     ' std::weak_ptr<WzMultiplayerOptionsTitleUI>: its address
Public wzText_ As WzText                ' WzText wzText;
Public optionsOverlayScreen As W_SCREEN ' std::shared_ptr<W_SCREEN>
Public optionsButton As WIDGET          ' std::shared_ptr<W_BUTTON>
