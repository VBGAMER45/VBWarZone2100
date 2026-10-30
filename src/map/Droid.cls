VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Droid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/include/wzmaplib/map_types.h: struct WzMap::Droid (with its
' base, struct Object). name (the template) is name_; position.x and
' position.y are position_x and position_y (uint32_t); direction is a
' uint16_t; optional<uint32_t> id is id with hasId; player is an int8_t
' (-1: the scavengers).
Option Explicit

Public name_ As String
Public position_x As Long
Public position_y As Long
Public direction As Long
Public id As Long
Public hasId As Boolean
Public player As Long
