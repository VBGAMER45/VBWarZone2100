VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HostingAdvancedOptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/starthostingbutton.cpp: the fields of class HostingAdvancedOptionsForm
' (VB) data of widget class HostingAdvancedOptionsForm
Option Explicit

Public multiOptionsTitleUI As Long     ' std::weak_ptr<WzMultiplayerOptionsTitleUI>: its address
Public parentStartHostingButton As Long ' std::weak_ptr<WzStartHostingButton>: its address
Public formTitle As WIDGET              ' std::shared_ptr<W_LABEL>
Public closeButton As WIDGET            ' std::shared_ptr<WzFrontendImageButton>
Public hostOptionsForm As WIDGET        ' std::shared_ptr<OptionsForm>
