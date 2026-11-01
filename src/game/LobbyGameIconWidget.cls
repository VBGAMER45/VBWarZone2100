VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LobbyGameIconWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class LobbyGameIconWidget
' src/titleui/widgets/gamebrowserform.cpp: the fields of class LobbyGameIconWidget (gamebrowserform_c)
Option Explicit

Public has_frontendImgID As Boolean     ' optional<UWORD> frontendImgID = nullopt;
Public frontendImgID As Long
Public has_customImgColor As Boolean    ' optional<PIELIGHT> customImgColor = nullopt;
Public customImgColor As Long
Public missingImage As Boolean          ' = false
Public imageWidth As Long               ' WzSize imageDimensions = WzSize(16,16);
Public imageHeight As Long
Public tipStr As String                 ' std::string tipStr;
