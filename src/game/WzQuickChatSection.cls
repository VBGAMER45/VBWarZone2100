VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzQuickChatSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzQuickChatSection
' src/hci/quickchat.cpp: the fields of class WzQuickChatSection (quickchat_c); std::vector<WzQuickChatMessage> messages is never used (not kept); weakSerial: (VB) its std::weak_ptrs (widgweak_vb); the initialisers are its _construct's.
Option Explicit

Public weakSerial As Long
Public messagesList As WIDGET              ' std::shared_ptr<IntListTabWidget>
Public hasFocus As Boolean                 ' = false
Public hasMouseOver As Boolean             ' = false
