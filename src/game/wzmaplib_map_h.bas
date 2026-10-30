Attribute VB_Name = "wzmaplib_map_h"
' lib/wzmaplib/include/wzmaplib/map.h: the header's inline functions and
' its map_round macro, for the game (WzMap.dll keeps its own copies; the
' constants are MapApi's map_h_defines).
Option Explicit

Public Function TileNumber_tile(ByVal tilenumber As Long) As Long
    TileNumber_tile = tilenumber And TILE_NUMMASK
End Function

Public Function TileNumber_texture(ByVal tilenumber As Long) As Long
    TileNumber_texture = tilenumber And (Not TILE_NUMMASK) And &HFFFF&
End Function

' (uint32_t)mapCoord << TILE_SHIFT
Public Function world_coord(ByVal mapCoord As Long) As Long
    If (mapCoord And &H1000000) <> 0 Then
        world_coord = ((mapCoord And &HFFFFFF) * TILE_UNITS) Or &H80000000
    Else
        world_coord = (mapCoord And &HFFFFFF) * TILE_UNITS
    End If
End Function

' worldCoord >> TILE_SHIFT (arithmetic)
Public Function map_coord(ByVal worldCoord As Long) As Long
    map_coord = (worldCoord And (Not TILE_MASK)) \ TILE_UNITS
End Function

' Only for graphics!
Public Function map_coordf(ByVal worldCoord As Long) As Single
    map_coordf = CSng(worldCoord) / CSng(TILE_UNITS)
End Function

Public Function round_to_nearest_tile(ByVal worldCoord As Long) As Long
    round_to_nearest_tile = (worldCoord + TILE_UNITS \ 2) And (Not TILE_MASK)
End Function

' maps a position down to the corner of a tile
Public Function map_round(ByVal coord As Long) As Long
    map_round = coord And (TILE_UNITS - 1)
End Function
