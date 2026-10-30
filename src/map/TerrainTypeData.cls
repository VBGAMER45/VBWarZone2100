VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TerrainTypeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/include/wzmaplib/map_types.h: struct WzMap::TerrainTypeData,
' the terrain type of each tile texture (std::vector<TYPE_OF_TERRAIN>
' terrainTypes): terrainTypes(i), with terrainTypes_size, _push_back and
' _clear. TYPE_OF_TERRAIN is terrain_type.h's (MapApi).
Option Explicit

Private types() As Long
Private typeCount As Long
Private typeCap As Long

Public Function terrainTypes_size() As Long
    terrainTypes_size = typeCount
End Function

Public Property Get terrainTypes(ByVal i As Long) As TYPE_OF_TERRAIN
    terrainTypes = types(i)
End Property

Public Property Let terrainTypes(ByVal i As Long, ByVal v As TYPE_OF_TERRAIN)
    types(i) = v
End Property

Public Sub terrainTypes_push_back(ByVal v As TYPE_OF_TERRAIN)
    If typeCount = typeCap Then
        typeCap = 2 * typeCap + 16
        ReDim Preserve types(0 To typeCap - 1)
    End If
    types(typeCount) = v
    typeCount = typeCount + 1
End Sub

Public Sub terrainTypes_clear()
    typeCount = 0
End Sub

' TerrainTypeData copy = *this
Public Function clone() As TerrainTypeData
    Dim c As TerrainTypeData, i As Long
    Set c = New TerrainTypeData
    For i = 0 To typeCount - 1
        c.terrainTypes_push_back types(i)
    Next
    Set clone = c
End Function
