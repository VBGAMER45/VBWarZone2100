VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzHostFreeChatPlayerMuteButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzHostFreeChatPlayerMuteButton
' src/hci/chatoptions.cpp: the fields of class WzHostFreeChatPlayerMuteButton (chatoptions_c); optional<uint32_t> playerIdx is has_playerIdx and the value; the std::weak_ptr parentForm is the form's widgweak_vb serial; the initialisers are its _construct's.
Option Explicit

Public has_playerIdx As Boolean            ' optional<uint32_t> playerIdx = nullopt
Public playerIdx As Long
Public parentForm As Long                  ' std::weak_ptr<WzChatOptionsForm>
