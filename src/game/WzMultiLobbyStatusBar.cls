VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMultiLobbyStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/multilobbyoptions.cpp: the fields of class WzMultiLobbyStatusBar
' (VB) data of widget class WzMultiLobbyStatusBar
Option Explicit

Public leftStatus As WIDGET             ' std::shared_ptr<W_LABEL>
Public rightStatus As WIDGET            ' std::shared_ptr<W_LABEL>
Public paddingX As Long                 ' int paddingX = 16;
Public paddingY As Long                 ' int paddingY = 5;
Public isChallenge As Boolean           ' bool isChallenge = false;
