VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChatBoxPlayerNameWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multiint.cpp: the fields of class ChatBoxPlayerNameWidget (multiint_lobby_c)
' (VB) data of widget class ChatBoxPlayerNameWidget
Option Explicit

Public player As Long                  ' std::shared_ptr<PlayerReference>: the player's index
Public font As Long                     ' iV_fonts
Public layoutName As String             ' WzString
Public cachedText As WzText             ' WzCachedText (VB: the WzText it caches)
Public horizontalPadding As Long        ' int32_t = 3
Public leftMargin As Long               ' int32_t = 3
Public ptsAboveBase As Long             ' int32_t = 0
