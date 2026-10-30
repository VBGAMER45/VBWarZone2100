VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzInGameChatBoxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzInGameChatBoxForm
' src/screens/chatscreen.cpp: the fields of class WzInGameChatBoxForm (chatscreen_c); the initialisers are its _construct's.
Option Explicit

Public chatMode As WzChatMode
Public chatBox As WIDGET                   ' std::shared_ptr<W_EDITBOX>
Public label As WIDGET                     ' std::shared_ptr<W_LABEL>
Public optionsButton As WIDGET             ' std::shared_ptr<WzInGameChatBoxButton>
Public optionsOverlayScreen As W_SCREEN    ' std::shared_ptr<W_SCREEN>
