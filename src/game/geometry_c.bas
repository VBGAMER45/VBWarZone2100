Attribute VB_Name = "geometry_c"
' src/geometry.cpp, geometry.h: calcDirection, getNearestDroid, inQuad,
' positionInQuad, objectOnScreen, QUAD, and geometry.h's getTileStructure and
' getTileFeature (getTileOccupier is not used by the port yet).
Option Explicit

Public Type QUAD
    coords(0 To 3) As Vector2i          ' = {{0,0}, {0,0}, {0,0}, {0,0}}
End Type

' uint16_t calcDirection(int32_t x0, int32_t y0, int32_t x1, int32_t y1)
Public Function calcDirection(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long) As Long
    calcDirection = iAtan2(x1 - x0, y1 - y0)
End Function


'	A useful function and one that should have been written long ago, assuming of course
'	that is hasn't been!!!! Alex M, 24th Sept, 1998. Returns the nearest unit
'	to a given world coordinate - we can choose whether we require that the unit be
'	selected or not... Makes sending the most logical unit to do something very easy.
'
'  NB*****THIS WON'T PICK A VTOL DROID*****

Public Function getNearestDroid(ByVal x As Long, ByVal y As Long, ByVal bSelected As Boolean) As Long
    Dim psBestUnit As Long, bestSoFar As Long, it As Long, psDroid As Long, dist As Long

    psBestUnit = 0
    bestSoFar = -1                      ' unsigned UDWORD_MAX

    If Not (selectedPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/geometry.cpp:53", "getNearestDroid", "Not supported selectedPlayer: %u", selectedPlayer: getNearestDroid = 0: Exit Function

    ' Go thru' all the droids  - how often have we seen this - a MACRO maybe?
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psDroid = lnObj(it)
        If Not DROID_isVtol(psDroid) Then
            ' Clever (?) bit that reads whether we're interested in droids being selected or not
            If Not bSelected Or aDroid(psDroid).selected <> 0 Then
                dist = iHypot(aDroid(psDroid).pos.x - x, aDroid(psDroid).pos.y - y)
                ' Is this the nearest one we got so far?
                If ULT(dist, bestSoFar) Then
                    ' Yes, then keep a record of the distance for comparison...
                    bestSoFar = dist
                    ' ..and store away the droid responsible
                    psBestUnit = psDroid
                End If
            End If
        End If
        it = lnNext(it)
    Loop
    getNearestDroid = psBestUnit
End Function
' -------------------------------------------------------------------------------------------

' Returns non-zero if a point is in a 4 sided polygon
' See header file for definition of QUAD
Public Function inQuad(pt As Vector2i, quad_ As QUAD) As Boolean
    Dim minX As Long, maxX As Long, minY As Long, maxY As Long, c As Boolean, i As Long, j As Long
    Dim edge As Vector2i, pos As Vector2i

    ' Early out.
    minX = minL(minL(quad_.coords(0).x, quad_.coords(1).x), minL(quad_.coords(2).x, quad_.coords(3).x))
    If pt.x < minX Then
        inQuad = False
        Exit Function
    End If
    maxX = maxL(maxL(quad_.coords(0).x, quad_.coords(1).x), maxL(quad_.coords(2).x, quad_.coords(3).x))
    If pt.x > maxX Then
        inQuad = False
        Exit Function
    End If
    minY = minL(minL(quad_.coords(0).y, quad_.coords(1).y), minL(quad_.coords(2).y, quad_.coords(3).y))
    If pt.y < minY Then
        inQuad = False
        Exit Function
    End If
    maxY = maxL(maxL(quad_.coords(0).y, quad_.coords(1).y), maxL(quad_.coords(2).y, quad_.coords(3).y))
    If pt.y > maxY Then
        inQuad = False
        Exit Function
    End If

    c = False

    i = 0
    j = 3
    Do While i < 4
        edge.x = quad_.coords(j).x - quad_.coords(i).x
        edge.y = quad_.coords(j).y - quad_.coords(i).y
        pos.x = pt.x - quad_.coords(i).x
        pos.y = pt.y - quad_.coords(i).y
        If (0 <= pos.y And pos.y < edge.y And CDbl(pos.x) * CDbl(edge.y) < CDbl(pos.y) * CDbl(edge.x)) Or _
           (edge.y <= pos.y And pos.y < 0 And CDbl(pos.x) * CDbl(edge.y) > CDbl(pos.y) * CDbl(edge.x)) Then
            c = Not c
        End If
        j = i
        i = i + 1
    Loop

    inQuad = c
End Function

Public Function positionInQuad(pt As Vector2i, quad_ As QUAD) As Vector2i
    Dim lenSq(0 To 3) As Long, ptDot(0 To 3) As Long, i As Long, j As Long, ret(0 To 1) As Long
    Dim edge As Vector2i, pos As Vector2i, posRot As Vector2i, d1 As Long, d2 As Long
    i = 0
    j = 3
    Do While i < 4
        edge.x = quad_.coords(j).x - quad_.coords(i).x
        edge.y = quad_.coords(j).y - quad_.coords(i).y
        pos.x = quad_.coords(j).x - pt.x
        pos.y = quad_.coords(j).y - pt.y
        posRot.x = pos.y
        posRot.y = -pos.x
        ' long (32 bits on Windows): dot() wraps as the C++ does
        lenSq(i) = MUL32(edge.x, edge.x) + MUL32(edge.y, edge.y)
        ptDot(i) = MUL32(posRot.x, edge.x) + MUL32(posRot.y, edge.y)
        j = i
        i = i + 1
    Loop
    For i = 0 To 1
        d1 = MUL32(ptDot(i), lenSq(i + 2))
        d2 = MUL32(ptDot(i + 2), lenSq(i))
        If CDbl(d1) + CDbl(d2) <> 0 Then
            ' (int64_t)TILE_UNITS * d1 / (d1 + d2)
            ret(i) = Fix(CDbl(TILE_UNITS) * CDbl(d1) / (CDbl(d1) + CDbl(d2)))
        Else
            ret(i) = TILE_UNITS \ 2
        End If
    Next
    positionInQuad.x = ret(0)
    positionInQuad.y = ret(1)
End Function

' (VB) std::min / std::max of two ints
Private Function minL(ByVal a As Long, ByVal b As Long) As Long
    If b < a Then minL = b Else minL = a
End Function

Private Function maxL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then maxL = b Else maxL = a
End Function

'-----------------------------------------------------------------------------------
' object: a BASE_OBJECT * (a droid, structure or feature handle)
Public Function objectOnScreen(ByVal object_ As Long, ByVal tolerance As Long) As Boolean
    Dim sDisplay As SCREEN_DISP_DATA, dX As Long, dY As Long
    Select Case objHandleType(object_)
    Case OBJ_DROID: sDisplay = aDroid(object_).sDisplay
    Case OBJ_STRUCTURE: sDisplay = aStruct(object_).sDisplay
    Case OBJ_FEATURE: sDisplay = aFeature(object_).sDisplay
    End Select
    If DrawnInLastFrame(sDisplay.frameNumber) = True Then
        dX = sDisplay.screenX
        dY = sDisplay.screenY
        ' Is it on screen
        If dX > (0 - tolerance) And dY > (0 - tolerance) _
           And dX < (pie_GetVideoBufferWidth() + tolerance) _
           And dY < (pie_GetVideoBufferHeight() + tolerance) Then
            objectOnScreen = True
            Exit Function
        End If
    End If
    objectOnScreen = False
End Function

Public Function getTileStructure(ByVal x As Long, ByVal y As Long) As Long
    Dim psObj As Long
    psObj = psMapTiles(mapTile(x, y)).psObject
    If psObj <> 0 Then
        If SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Then
            getTileStructure = psObj
        End If
    End If
End Function

Public Function getTileFeature(ByVal x As Long, ByVal y As Long) As Long
    Dim psObj As Long
    psObj = psMapTiles(mapTile(x, y)).psObject
    If psObj <> 0 Then
        If SIMPLE_OBJECT_type(psObj) = OBJ_FEATURE Then
            getTileFeature = psObj
        End If
    End If
End Function
