VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChatBoxWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multiint.cpp: the fields of class ChatBoxWidget (multiint_lobby_c)
' (VB) data of widget class ChatBoxWidget
Option Explicit

Public messages As WIDGET               ' std::shared_ptr<ScrollableListWidget>
Public quickChatButton As WIDGET        ' std::shared_ptr<ChatBoxButton>
Public optionsButton As WIDGET          ' std::shared_ptr<ChatBoxButton>
Public editBox As WIDGET                ' std::shared_ptr<W_EDITBOX>
Public handleConsoleMessage As WzFunc   ' std::shared_ptr<CONSOLE_MESSAGE_LISTENER>
Public quickChatOverlayScreen As W_SCREEN ' std::shared_ptr<W_SCREEN>
Public quickChatForm As WIDGET          ' std::shared_ptr<W_FORM>
Public optionsOverlayScreen As W_SCREEN ' std::shared_ptr<W_SCREEN>

Public currentMode As Long              ' ChatBoxSendMode = ChatBoxSendMode::ENABLED
Public hasFocus As Boolean              ' = false

Public weakSelf As Long                 ' (VB) its std::weak_ptr<ChatBoxWidget>s: the widgweak_vb serial
