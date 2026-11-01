VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_NOTIFICATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_NOTIFICATION
' src/notifications.cpp: the fields of class W_NOTIFICATION (notifications_c); the initialisers are its _construct_init's.
' Vector2i dragOffset and dragOffsetEnded are their x and y.
Option Explicit

Public pOnDoNotShowAgainCheckbox As WIDGET ' std::shared_ptr<WzCheckboxButton> = nullptr
Public request As Long                  ' WZ_Queued_Notification*: notifications_c's handle (0 = nullptr)
Public isModal As Boolean               ' = false
Public isInDragMode As Boolean          ' = false
Public dragOffset_x As Long             ' Vector2i dragOffset = {0, 0}
Public dragOffset_y As Long
Public dragOffsetEnded_x As Long        ' Vector2i dragOffsetEnded = {0, 0}
Public dragOffsetEnded_y As Long
Public dragStartedTime As Long          ' uint32_t = 0
Public dragEndedTime As Long            ' uint32_t = 0
Public pImageTexture As gfx_api_texture ' gfx_api::texture* = nullptr
Public psModalParent As Long            ' std::weak_ptr<W_FORM>: its address (WIDGET_fromPtr)
