VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayRequestOptionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multimenu.cpp: struct DisplayRequestOptionData {pMapData, cache} and its
' struct DisplayRequestOptionCache (their members flattened: the cache's are
' this object's), a request list button's pUserData (multimenu_c's
' displayRequestOption). DisplayRequestOptionData(LEVEL_DATASET *pMapData =
' nullptr) is multimenu_c's DisplayRequestOptionData_new.
Option Explicit

Public pMapData As LEVEL_DATASET        ' LEVEL_DATASET * (Nothing = nullptr)

' DisplayRequestOptionCache cache:
Private hash_ As Sha256                 ' Sha256 hash
Public wzHashText As WzText

Private wzText As WzText
Private fullButString_ As String        ' std::string _fullButString
Private widgetWidth_ As Long            ' int _widgetWidth = 0

Private Sub Class_Initialize()
    Set wzHashText = New WzText
    Set wzText = New WzText
End Sub

Friend Property Get hash() As Sha256
    hash = hash_
End Property

Friend Property Let hash(v As Sha256)
    hash_ = v
End Property

' For utilizing the cached (main) display text
Public Function canUseCachedText(ByVal fullButString As String, ByVal widgetWidth As Long) As Boolean
    canUseCachedText = (widgetWidth_ = widgetWidth) And (fullButString_ = fullButString)
End Function

Public Sub setCachedText(ByVal text As String, ByVal fontID As Long, ByVal fullButString As String, ByVal widgetWidth As Long)
    wzText.setText text, fontID
    fullButString_ = fullButString
    widgetWidth_ = widgetWidth
End Sub

Public Sub renderCachedText(ByVal x As Long, ByVal y As Long, ByVal colour As Long, Optional ByVal rot_ As Single = 0!)
    wzText.render x, y, colour, rot_
End Sub
