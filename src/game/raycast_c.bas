Attribute VB_Name = "raycast_c"
' src/raycast.cpp, raycast.h: Raycasting routine that gives intersection
' points with map tiles: rayCast, and the camera's getBestPitchToEdgeOfGrid
' (with its getTileHeightCallback).
' (VB) A RAY_CALLBACK is a number (RAY_CALLBACK_...) and rayCallback's
' Select Case; a callback's void *data is its own module's variables, set
' before the call. The dist argument (iHypot of the point) is worked out
' only for a callback that reads it.
Option Explicit

Public Enum RAY_CALLBACK
    RAY_CALLBACK_MOVE_BLOCKING_TILE = 1         ' move.cpp's moveBlockingTileCallback (dist unused)
    RAY_CALLBACK_LOS                            ' visibility.cpp's rayLOSCallback
    RAY_CALLBACK_TILE_HEIGHT                    ' getTileHeightCallback (here)
End Enum

' struct HeightCallbackHelp_t: getBestPitchToEdgeOfGrid's data for getTileHeightCallback
Private help_height As Long             ' const int height
Private help_pitch As Long              ' uint16_t pitch

' (VB) the callback(pos, dist, data) call
Private Function rayCallback(ByVal callback As RAY_CALLBACK, pos As Vector2i, ByVal withDist As Boolean) As Boolean
    Select Case callback
    Case RAY_CALLBACK_MOVE_BLOCKING_TILE
        rayCallback = moveBlockingTileCallback(pos)
    Case RAY_CALLBACK_LOS
        If withDist Then
            rayCallback = rayLOSCallback(pos, iHypot(pos.x, pos.y))
        Else
            rayCallback = rayLOSCallback(pos, 0)
        End If
    Case RAY_CALLBACK_TILE_HEIGHT
        If withDist Then
            rayCallback = getTileHeightCallback(pos, iHypot(pos.x, pos.y))
        Else
            rayCallback = getTileHeightCallback(pos, 0)
        End If
    End Select
End Function

Private Sub initSteps(ByVal srcM As Long, ByVal dstM As Long, tile As Long, step_ As Long, cur As Long, end_ As Long)
    Dim increasing As Long
    increasing = -(srcM < dstM)
    step_ = -1 + 2 * increasing
    tile = srcM - step_
    cur = srcM + increasing
    end_ = dstM + increasing
End Sub

' Finds the next intersection of the line with a vertical grid line (or with a horizontal grid line, if called with x and y swapped).
Private Function tryStep(tile As Long, ByVal step_ As Long, cur As Long, ByVal end_ As Long, px As Long, py As Long, _
                        ByVal sx As Long, ByVal sy As Long, ByVal dx As Long, ByVal dy As Long) As Boolean
    tile = tile + step_

    If cur = end_ Then
        tryStep = False  ' No more vertical grid lines to cross before reaching the endpoint.
        Exit Function
    End If

    ' Find the point on the line with the x coordinate world_coord(cur).
    px = world_coord(cur)
    py = sy + Fix(CDbl(px - sx) * (dy - sy) / (dx - sx))

    cur = cur + step_
    tryStep = True
End Function

' Cast a ray from src to dst. The callback is called with each intersection of the ray with a tile edge, and src and dst, as
' long as it returns true. (Each tile the ray passes through is given once, with a point on that tile.)
' void rayCast(Vector2i src, Vector2i dst, RAY_CALLBACK callback, void *data)
Public Sub rayCast(src As Vector2i, dst As Vector2i, ByVal callback As RAY_CALLBACK)
    Dim srcM As Vector2i, dstM As Vector2i, step_ As Vector2i, tile As Vector2i, cur As Vector2i, end_ As Vector2i
    Dim prev As Vector2i, first As Boolean, nextX As Vector2i, nextY As Vector2i, canX As Boolean, canY As Boolean
    Dim xDist As Long, yDist As Long, sel As Vector2i, selTile As Vector2i, avg As Vector2i

    If Not rayCallback(callback, src, False) Then   ' Start at src.
        Exit Sub  ' Callback gave up after the first point, or there are no other points.
    End If
    If src.x = dst.x And src.y = dst.y Then
        Exit Sub
    End If

    srcM.x = map_coord(src.x)
    srcM.y = map_coord(src.y)
    dstM.x = map_coord(dst.x)
    dstM.y = map_coord(dst.y)

    initSteps srcM.x, dstM.x, tile.x, step_.x, cur.x, end_.x
    initSteps srcM.y, dstM.y, tile.y, step_.y, cur.y, end_.y

    first = True
    canX = tryStep(tile.x, step_.x, cur.x, end_.x, nextX.x, nextX.y, src.x, src.y, dst.x, dst.y)
    canY = tryStep(tile.y, step_.y, cur.y, end_.y, nextY.y, nextY.x, src.y, src.x, dst.y, dst.x)
    Do While canX Or canY
        xDist = Abs(nextX.x - src.x) + Abs(nextX.y - src.y)
        yDist = Abs(nextY.x - src.x) + Abs(nextY.y - src.y)
        If canX And (Not canY Or xDist < yDist) Then  ' The line crosses a vertical grid line next.
            sel = nextX
            selTile = tile
            canX = tryStep(tile.x, step_.x, cur.x, end_.x, nextX.x, nextX.y, src.x, src.y, dst.x, dst.y)
        Else  ' The line crosses a horizontal grid line next.
            Debug.Assert canY
            sel = nextY
            selTile = tile
            canY = tryStep(tile.y, step_.y, cur.y, end_.y, nextY.y, nextY.x, src.y, src.x, dst.y, dst.x)
        End If
        If Not first Then
            ' Find midpoint.
            avg.x = (prev.x + sel.x) \ 2
            avg.y = (prev.y + sel.y) \ 2
            ' But make sure it's on the right tile, since it could be off-by-one if the line passes exactly through a grid intersection.
            avg.x = MIN_(MAX_(avg.x, world_coord(selTile.x)), world_coord(selTile.x + 1) - 1)
            avg.y = MIN_(MAX_(avg.y, world_coord(selTile.y)), world_coord(selTile.y + 1) - 1)
            If Not worldOnMap(avg.x, avg.y) Then
                Exit Sub  ' Callback doesn't want any more points, or we reached the edge of the map, so return.
            End If
            If Not rayCallback(callback, avg, True) Then
                Exit Sub
            End If
        End If
        prev = sel
        first = False
    Loop

    ' Include the endpoint.
    If Not worldOnMap(dst.x, dst.y) Then
        Exit Sub  ' Stop, since reached the edge of the map.
    End If
    rayCallback callback, dst, True
End Sub

'-----------------------------------------------------------------------------------
' Will return false when we've hit the edge of the grid
Private Function getTileHeightCallback(pos As Vector2i, ByVal dist As Long) As Boolean
    Dim HasTallStructure As Boolean, height As Long, heightDiff As Long, newPitch As Long

    ' Are we still on the grid?
    If clipXY(pos.x, pos.y) Then
        HasTallStructure = (blockTile(map_coord(pos.x), map_coord(pos.y), AUX_MAP) And AIR_BLOCKED) <> 0

        If dist > TILE_UNITS Or HasTallStructure Then
            ' Only do it the current tile is > TILE_UNITS away from the starting tile. Or..
            ' there is a tall structure  on the current tile and the current tile is not the starting tile.
            ' Get height at this intersection point
            height = map_Height(pos.x, pos.y)

            If HasTallStructure Then
                height = height + TALLOBJECT_ADJUST
            End If

            If height <= help_height Then
                heightDiff = 0
            Else
                heightDiff = height - help_height
            End If

            ' Work out the angle to this point from start point
            newPitch = iAtan2(heightDiff, dist)

            ' Is this the steepest we've found?
            If angleDelta(newPitch - help_pitch) > 0 Then
                ' Yes, then keep a record of it
                help_pitch = newPitch
            End If
            '---
        End If

        ' Not at edge yet - so exit
        getTileHeightCallback = True
        Exit Function
    End If

    ' We've hit edge of grid - so exit!!
    getTileHeightCallback = False
End Function

' void getBestPitchToEdgeOfGrid(UDWORD x, UDWORD y, uint16_t direction, uint16_t *pitch)
Public Sub getBestPitchToEdgeOfGrid(ByVal x As Long, ByVal y As Long, ByVal direction As Long, pitch As Long)
    Dim src As Vector2i, dst As Vector2i
    help_height = map_Height(x, y)          ' HeightCallbackHelp_t help = {map_Height(x, y), 0};
    help_pitch = 0

    src.x = x                               ' Vector3i src(x, y, 0);
    src.y = y
    dst.x = x + iSinR(direction, 5430)      ' Vector3i delta(iSinCosR(direction, 5430), 0);
    dst.y = y + iCosR(direction, 5430)
    rayCast src, dst, RAY_CALLBACK_TILE_HEIGHT  ' FIXME Magic value

    pitch = help_pitch
End Sub
