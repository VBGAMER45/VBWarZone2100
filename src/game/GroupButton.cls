VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GroupButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class GroupButton
' src/hci/groups.h, groups.cpp: GroupButton's fields (groups_c)
Option Explicit

Public controller As GroupsUIController         ' std::shared_ptr<GroupsUIController>
Public groupNumberLabel As WIDGET               ' std::shared_ptr<W_LABEL>
Public groupCountLabel As WIDGET                ' std::shared_ptr<W_LABEL>
Public groupDamagedCountLabel As WIDGET         ' std::shared_ptr<W_LABEL>
Public groupNumber As Long                      ' size_t
Public lastUpdatedGlowAlphaTime As Long         ' uint32_t = 0
