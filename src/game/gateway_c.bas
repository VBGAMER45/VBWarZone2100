Attribute VB_Name = "gateway_c"
' src/gateway.cpp, gateway.h: the routing gateways of the map.
' GATEWAY_LIST (std::list<GATEWAY *>) is an array of GATEWAY in list order
' with a count; a GATEWAY * is its index (gwFreeGateway takes the entry).
Option Explicit

Public Type GATEWAY
    x1 As Byte
    y1 As Byte
    x2 As Byte
    y2 As Byte
End Type

' the list of gateways on the current map
Private psGateways() As GATEWAY
Private psGateways_size As Long

'*****************************************************************************************************
'                   Gateway data access functions

' get the size of the map
Private Function gwMapWidth() As Long
    gwMapWidth = mapWidth
End Function

Private Function gwMapHeight() As Long
    gwMapHeight = mapHeight
End Function

' set the gateway flag on a tile
Private Sub gwSetGatewayFlag(ByVal x As Long, ByVal y As Long)
    Dim t As Long
    t = mapTile(x, y)
    psMapTiles(t).tileInfoBits = psMapTiles(t).tileInfoBits Or BITS_GATEWAY
End Sub

' clear the gateway flag on a tile
Private Sub gwClearGatewayFlag(ByVal x As Long, ByVal y As Long)
    Dim t As Long
    t = mapTile(x, y)
    psMapTiles(t).tileInfoBits = psMapTiles(t).tileInfoBits And (Not BITS_GATEWAY) And &HFF&
End Sub


'*****************************************************************************************************
'                   Gateway functions

' Initialise the gateway system
Public Function gwInitialise() As Boolean
    Erase psGateways
    psGateways_size = 0
    gwInitialise = True
End Function

' Shutdown the gateway system
Public Sub gwShutDown()
    Dim i As Long
    For i = 0 To psGateways_size - 1
        gwFreeGateway psGateways(i)
    Next
    Erase psGateways
    psGateways_size = 0
End Sub

' (VB) std::swap(mission.psGateways, gwGetGateways()) (mission.cpp): the gateway list with another
Public Sub gwSwapGateways(other() As GATEWAY, otherSize As Long)
    Dim t() As GATEWAY, tSize As Long, i As Long
    tSize = psGateways_size
    If tSize > 0 Then
        ReDim t(0 To tSize - 1)
        For i = 0 To tSize - 1
            t(i) = psGateways(i)
        Next
    End If
    Erase psGateways
    psGateways_size = otherSize
    If otherSize > 0 Then
        ReDim psGateways(0 To otherSize - 1)
        For i = 0 To otherSize - 1
            psGateways(i) = other(i)
        Next
    End If
    Erase other
    otherSize = tSize
    If tSize > 0 Then
        ReDim other(0 To tSize - 1)
        For i = 0 To tSize - 1
            other(i) = t(i)
        Next
    End If
End Sub

' Add a gateway to the system
Public Function gwNewGateway(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    Dim psNew As GATEWAY, pos As Long, temp As Long

    If Not (x1 >= 0 And x1 < gwMapWidth() And y1 >= 0 And y1 < gwMapHeight() _
            And x2 >= 0 And x2 < gwMapWidth() And y2 >= 0 And y2 < gwMapHeight() _
            And (x1 = x2 Or y1 = y2)) Then
        ASSERT_FAILURE "x1 >= 0 && x1 < gwMapWidth() && y1 >= 0 && y1 < gwMapHeight() && x2 >= 0 && x2 < gwMapWidth() && y2 >= 0 && y2 < gwMapHeight() && (x1 == x2 || y1 == y2)", _
            "src/gateway.cpp:93", "gwNewGateway", "Invalid gateway coordinates (%d, %d, %d, %d)", x1, y1, x2, y2
        gwNewGateway = False
        Exit Function
    End If

    ' make sure the first coordinate is always the smallest
    If x2 < x1 Then
        ' y is the same, swap x
        temp = x2
        x2 = x1
        x1 = temp
    ElseIf y2 < y1 Then
        ' x is the same, swap y
        temp = y2
        y2 = y1
        y1 = temp
    End If

    ' Initialise the gateway, correct out-of-map gateways (UBYTE fields)
    psNew.x1 = MAX_(3, x1) And &HFF&
    psNew.y1 = MAX_(3, y1) And &HFF&
    psNew.x2 = MIN_(x2, mapWidth - 4) And &HFF&
    psNew.y2 = MIN_(y2, mapHeight - 4) And &HFF&

    ' add the gateway to the list
    ReDim Preserve psGateways(0 To psGateways_size)
    psGateways(psGateways_size) = psNew
    psGateways_size = psGateways_size + 1

    ' set the map flags
    If psNew.x1 = psNew.x2 Then
        ' vertical gateway
        For pos = psNew.y1 To psNew.y2
            gwSetGatewayFlag psNew.x1, pos
        Next
    Else
        ' horizontal gateway
        For pos = psNew.x1 To psNew.x2
            gwSetGatewayFlag pos, psNew.y1
        Next
    End If

    gwNewGateway = True
End Function

' Return the number of gateways.
Public Function gwNumGateways() As Long
    gwNumGateways = psGateways_size
End Function

' (VB) gwGetGateways: the i-th gateway of the list
Public Sub gwGetGateway(ByVal i As Long, gw As GATEWAY)
    gw = psGateways(i)
End Sub

' Release a gateway
Private Sub gwFreeGateway(psDel As GATEWAY)
    Dim pos As Long

    If psMapTiles_allocated Then ' this lines fixes the bug where we were closing the gateways after freeing the map
        ' clear the map flags
        If psDel.x1 = psDel.x2 Then
            ' vertical gateway
            For pos = psDel.y1 To psDel.y2
                gwClearGatewayFlag psDel.x1, pos
            Next
        Else
            ' horizontal gateway
            For pos = psDel.x1 To psDel.x2
                gwClearGatewayFlag pos, psDel.y1
            Next
        End If
    End If

    ' free(psDel): the list's entry goes with the list
End Sub
