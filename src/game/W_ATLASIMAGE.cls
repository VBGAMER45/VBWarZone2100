VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_ATLASIMAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_ATLASIMAGE
' src/screens/guidescreen.cpp: W_ATLASIMAGE's fields (guidescreen_c); the initialisers are W_ATLASIMAGE_construct's
' and Class_Initialize's.
Option Explicit

Public ImageFile As Long                     ' IMAGEFILE * = nullptr
Public ImageIDs As Collection                ' std::vector<UWORD>
Public imageWidths As Collection             ' std::vector<int32_t>
Public imageHeights As Collection            ' std::vector<int32_t>
Public maxImageHeight As Long                ' int32_t = 0
Public totalIdealWidth As Long               ' int32_t = 0

Private Sub Class_Initialize()
    Set ImageIDs = New Collection
    Set imageWidths = New Collection
    Set imageHeights = New Collection
End Sub
