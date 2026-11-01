VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/include/wzmaplib/map_package.h: struct WzMap::GamInfo, the basic .gam data maps use. The uint32_t
' fields are Longs (the same bits); char levelName[OLD_MAX_LEVEL_SIZE] is its text up to the first NUL.
Option Explicit

Public gameTime As Long                     ' uint32_t
Public GameType As Long                     ' uint32_t: Type of game, one of the GTYPE_... enums.
Public ScrollMinX As Long                   ' int32_t: Scroll Limits
Public ScrollMinY As Long
Public ScrollMaxX As Long                   ' uint32_t
Public ScrollMaxY As Long
Public levelName As String                  ' name of the level to load up when mid game
