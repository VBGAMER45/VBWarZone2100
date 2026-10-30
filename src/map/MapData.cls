VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/include/wzmaplib/map_types.h: struct WzMap::MapData, a
' game.map: its size, its tiles (struct MapTile: height, texture, both
' uint16_t) and its gateways (struct Gateway: x1, y1, x2, y2, uint8_t).
' The vectors are mMapTiles_<field>(i) and mGateways_<field>(i), with
' _size, _push_back, _reserve and _clear. crcSumMapTiles is
' lib/wzmaplib/src/map.cpp's. width and height are uint32_t (Longs with
' the 32 bits).
Option Explicit

Public height As Long
Public width As Long

Private tileHeight() As Long
Private tileTexture() As Long
Private tileCount As Long
Private tileCap As Long
Private gwX1() As Long, gwY1() As Long, gwX2() As Long, gwY2() As Long
Private gwCount As Long
Private gwCap As Long

Public Function mMapTiles_size() As Long
    mMapTiles_size = tileCount
End Function

Public Property Get mMapTiles_height(ByVal i As Long) As Long
    mMapTiles_height = tileHeight(i)
End Property

Public Property Let mMapTiles_height(ByVal i As Long, ByVal v As Long)
    tileHeight(i) = v And &HFFFF&
End Property

Public Property Get mMapTiles_texture(ByVal i As Long) As Long
    mMapTiles_texture = tileTexture(i)
End Property

Public Property Let mMapTiles_texture(ByVal i As Long, ByVal v As Long)
    tileTexture(i) = v And &HFFFF&
End Property

Public Sub mMapTiles_reserve(ByVal n As Long)
    If n <= tileCap Then Exit Sub
    ReDim Preserve tileHeight(0 To n - 1)
    ReDim Preserve tileTexture(0 To n - 1)
    tileCap = n
End Sub

Public Sub mMapTiles_push_back(ByVal height_ As Long, ByVal texture As Long)
    If tileCount = tileCap Then mMapTiles_reserve 2 * tileCount + 16
    tileHeight(tileCount) = height_ And &HFFFF&
    tileTexture(tileCount) = texture And &HFFFF&
    tileCount = tileCount + 1
End Sub

' resize(n): new tiles are {0, 0}
Public Sub mMapTiles_resize(ByVal n As Long)
    Dim i As Long
    If n > tileCount Then
        mMapTiles_reserve n
        For i = tileCount To n - 1
            tileHeight(i) = 0
            tileTexture(i) = 0
        Next
    End If
    tileCount = n
End Sub

Public Sub mMapTiles_clear()
    tileCount = 0
End Sub

Public Function mGateways_size() As Long
    mGateways_size = gwCount
End Function

Public Property Get mGateways_x1(ByVal i As Long) As Long
    mGateways_x1 = gwX1(i)
End Property

Public Property Let mGateways_x1(ByVal i As Long, ByVal v As Long)
    gwX1(i) = v And &HFF&
End Property

Public Property Get mGateways_y1(ByVal i As Long) As Long
    mGateways_y1 = gwY1(i)
End Property

Public Property Let mGateways_y1(ByVal i As Long, ByVal v As Long)
    gwY1(i) = v And &HFF&
End Property

Public Property Get mGateways_x2(ByVal i As Long) As Long
    mGateways_x2 = gwX2(i)
End Property

Public Property Let mGateways_x2(ByVal i As Long, ByVal v As Long)
    gwX2(i) = v And &HFF&
End Property

Public Property Get mGateways_y2(ByVal i As Long) As Long
    mGateways_y2 = gwY2(i)
End Property

Public Property Let mGateways_y2(ByVal i As Long, ByVal v As Long)
    gwY2(i) = v And &HFF&
End Property

Public Sub mGateways_reserve(ByVal n As Long)
    If n <= gwCap Then Exit Sub
    ReDim Preserve gwX1(0 To n - 1)
    ReDim Preserve gwY1(0 To n - 1)
    ReDim Preserve gwX2(0 To n - 1)
    ReDim Preserve gwY2(0 To n - 1)
    gwCap = n
End Sub

Public Sub mGateways_push_back(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    If gwCount = gwCap Then mGateways_reserve 2 * gwCount + 16
    gwX1(gwCount) = x1 And &HFF&
    gwY1(gwCount) = y1 And &HFF&
    gwX2(gwCount) = x2 And &HFF&
    gwY2(gwCount) = y2 And &HFF&
    gwCount = gwCount + 1
End Sub

Public Sub mGateways_clear()
    gwCount = 0
End Sub

' MapData copy = *this
Public Function clone() As MapData
    Dim c As MapData, i As Long
    Set c = New MapData
    c.width = width
    c.height = height
    c.mMapTiles_reserve tileCount
    For i = 0 To tileCount - 1
        c.mMapTiles_push_back tileHeight(i), tileTexture(i)
    Next
    c.mGateways_reserve gwCount
    For i = 0 To gwCount - 1
        c.mGateways_push_back gwX1(i), gwY1(i), gwX2(i), gwY2(i)
    Next
    Set clone = c
End Function

' lib/wzmaplib/src/map.cpp: MapData::crcSumMapTiles
Public Function crcSumMapTiles(ByVal crc As Long) As Long
    Dim i As Long, v(0 To 0) As Long
    For i = 0 To tileCount - 1
        v(0) = tileHeight(i)
        crc = crcSumU16(crc, v, 0, 1)
        v(0) = tileTexture(i)
        crc = crcSumU16(crc, v, 0, 1)
    Next
    crcSumMapTiles = crc
End Function
