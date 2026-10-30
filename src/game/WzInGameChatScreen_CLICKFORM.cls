VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzInGameChatScreen_CLICKFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzInGameChatScreen_CLICKFORM
' src/screens/chatscreen.cpp: the fields of class WzInGameChatScreen_CLICKFORM (chatscreen_c); the initialisers are its _construct's.
Option Explicit

Public backgroundColor As Long             ' PIELIGHT = pal_RGBA(0, 0, 0, 80)
Public onClickedFunc As WzFunc             ' std::function<void ()>
Public onCancelPressed As WzFunc           ' std::function<void ()>
Public consoleBox As WIDGET                ' std::shared_ptr<WzInGameChatBoxForm>
Public quickChatForm As WIDGET             ' std::shared_ptr<W_FORM>
