VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzChatOptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzChatOptionsForm
' src/hci/chatoptions.cpp: the fields of class WzChatOptionsForm (chatoptions_c); weakSerial: (VB) its std::weak_ptrs (widgweak_vb); the initialisers are its _construct's.
Option Explicit

Public weakSerial As Long
Public updateButton As WIDGET              ' std::shared_ptr<WzChatOptionsButton>

Public showingHostingOptions As Boolean    ' = false
Public hostChatOptionsLabel As WIDGET      ' std::shared_ptr<W_LABEL>
Public globalOptionsLabel As WIDGET        ' std::shared_ptr<W_LABEL>
Public freechatEnableAllButton As WIDGET   ' std::shared_ptr<WzChatOptionsButton>
Public freechatDisableAllButton As WIDGET  ' std::shared_ptr<WzChatOptionsButton>
Public joiningClientsButton As WIDGET      ' std::shared_ptr<WzHostFreeChatPlayerMuteButton>

Public chatOptionsLabel As WIDGET          ' std::shared_ptr<W_LABEL>
Public chatOptionsDescriptionLabel As WIDGET ' std::shared_ptr<W_LABEL>
Public chatOptionsHostDisabledMsgLabel As WIDGET ' std::shared_ptr<W_LABEL>

Public playersTable As WIDGET              ' std::shared_ptr<ScrollableTableWidget>

Public onSettingsChange As WzFunc          ' std::function<void ()>
