VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_INTELLIGENCEOVERLAY_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_INTELLIGENCEOVERLAY_FORM
' src/intelmap.cpp: W_INTELLIGENCEOVERLAY_FORM's fields (intelmap_c); the initialisers are
' W_INTELLIGENCEOVERLAY_FORM_initMembers'.
Option Explicit

Public multiMenuForm As WIDGET          ' std::shared_ptr<IntFormAnimated>
Public msgForm As WIDGET                ' std::shared_ptr<IntFormAnimated>
Public selectedMsgButton As WIDGET      ' std::shared_ptr<IntMessageButton>
Public msgDetailsView As WIDGET         ' std::shared_ptr<WzMessageView>
Public isClosing As Boolean             ' = false
Public delayedPlayCurrent As Boolean    ' = false
Public playing As AUDIO_STREAM          ' AUDIO_STREAM * = nullptr (Nothing)
