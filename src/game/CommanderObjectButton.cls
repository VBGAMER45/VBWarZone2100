VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommanderObjectButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class CommanderObjectButton
' src/hci/commander.cpp: CommanderObjectButton's fields (commander_c)
Option Explicit

Public groupSizeLabel As WIDGET                 ' std::shared_ptr<W_LABEL>
Public experienceStarsLabel As WIDGET           ' std::shared_ptr<W_LABEL>
Public controller As CommanderController        ' std::shared_ptr<CommanderController>
