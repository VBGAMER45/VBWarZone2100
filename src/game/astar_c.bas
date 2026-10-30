Attribute VB_Name = "astar_c"
' src/astar.cpp, astar.h: A* path finding.
'  * First time (in a given tick) that some droid wants to pathfind to a
'    particular destination, the A* algorithm from source to destination is
'    used. The desired destination, and the nearest reachable point to the
'    destination is saved in a Context.
'  * Second time (in a given tick) that some droid wants to pathfind to a
'    particular destination, the appropriate Context is found, and the A*
'    algorithm is used to find a path from the nearest reachable point to
'    the destination (which was saved earlier), to the source.
'  * Subsequent times (in a given tick) that some droid wants to pathfind to
'    a particular destination, the path is looked up in appropriate Context.
'    If the path is not already known, the A* weights are adjusted, and the
'    previous A* pathfinding is continued until the new source is reached.
'    If the new source is not reached, the droid is on a different island
'    than the previous droid, and pathfinding is restarted from the first
'    step.
' The PathNode heap contains the priority-heap-sorted nodes which are to be
' explored. The path back is stored in the PathExploredTile 2D array of
' tiles.
' (VB) The path jobs run on the main thread (see fpath_c), so there is one
' FPathExecuteContext: this module's context list. A PathfindContext is a
' slot of ctxs; PathfindContextList's contexts are ctxs(0 .. ctxCount - 1)
' in push order and its orderedIndexes ctxOrder. A cleared list keeps its
' slots' arrays for reuse; a reused slot keeps counting its iteration from
' where it was, so no old map entry matches it, as none of a new context's
' does. A shared_ptr<PathBlockingMap> is a slot of blockingMaps (0 =
' nullptr), freed when neither fpathBlockingMaps nor a context refers to it.
' The heap of PathNodes (four arrays: nodePx, nodePy, nodeDist, nodeEst) is
' a binary max-heap under PathNode's operator < (nodeLess): that order is
' total, so which node std::pop_heap gives next does not depend on how the
' heap is laid out.
Option Explicit

' return codes for astar
Public Enum ASR_RETVAL
    ASR_OK                                  ' found a route
    ASR_FAILED                              ' no route could be found
    ASR_NEAREST                             ' found a partial route to a nearby position
End Enum

' Pathfinding blocking map
Private Type PathBlockingMap
    ' PathBlockingType type
    typeGameTime As Long                    ' uint32_t
    typePropulsion As Long                  ' PROPULSION_TYPE
    typeOwner As Long
    typeMoveType As Long                    ' FPATH_MOVETYPE
    map() As Boolean                        ' std::vector<bool>
    dangerMap() As Boolean                  ' using threatBits
    dangerMap_empty As Boolean
    used As Boolean                         ' (VB) the slot holds a map
End Type

Private Type PathNonblockingArea
    x1 As Long                              ' int16_t
    x2 As Long
    y1 As Long
    y2 As Long
End Type

' Data structures used for pathfinding, can contain cached results.
Private Type PathfindContext
    tileSx As Long                          ' PathCoord tileS: Start tile for pathfinding. (May be either source or target tile.)
    tileSy As Long
    myGameTime As Long                      ' uint32_t
    nearestX As Long                        ' PathCoord nearestCoord: Nearest reachable tile to destination.
    nearestY As Long
    iteration As Long                       ' uint16_t: Counter to implement lazy deletion from map.
    ' std::vector<PathNode> nodes: Edge of explored region of the map. PathNode's fields:
    nodePx() As Long                        ' PathCoord p (int16_t): Map coords.
    nodePy() As Long
    nodeDist() As Long                      ' unsigned: Distance so far
    nodeEst() As Long                       ' unsigned: and estimate to end.
    nodes_size As Long
    nodes_capacity As Long                  ' (VB) the arrays' size
    ' std::vector<PathExploredTile> map: Map, with paths leading back to tileS.
    mapIteration() As Long                  ' uint16_t iteration (PathExploredTile() : iteration(0xFFFF))
    mapDx() As Integer                      ' int8_t dx, dy: Offset from previous point in the route.
    mapDy() As Integer
    mapDist() As Long                       ' unsigned: Shortest known distance to tile.
    mapVisited() As Boolean
    map_size As Long
    blockingMap As Long                     ' Map of blocking tiles for the type of object which needs a path.
    dstIgnore As PathNonblockingArea        ' Area of structure at destination which should be considered nonblocking.
End Type

Private Const SIGN_BIT As Long = &H80000000

' (VB) the shared_ptr<PathBlockingMap>s
Private blockingMaps() As PathBlockingMap
Private blockingMaps_size As Long

' Lists of blocking maps from current tick.
Private fpathBlockingMaps() As Long
Private fpathBlockingMaps_size As Long
' Game time for all blocking maps in fpathBlockingMaps.
Private fpathCurrentGameTime As Long

' (VB) the FPathExecuteContextImpl's PathfindContextList fpathContexts
Private ctxs() As PathfindContext
Private ctxs_capacity As Long
Private ctxCount As Long                    ' contexts.size()
Private ctxOrder() As Long                  ' orderedIndexes

' Convert a direction into an offset
' dir 0 => x = 0, y = -1
Private aDirOffsetX(0 To 7) As Long
Private aDirOffsetY(0 To 7) As Long
Private aDirOffsetReady As Boolean

' static const uint8_t gradYLookup[99] (fpathNewNode): = sqrt(140^2 - gradientX^2), rounded to nearest integer
Private gradYLookup(0 To 98) As Long

Private Sub initTables()
    Dim v As Variant, i As Long
    v = Array(0, 1, -1, 1, -1, 0, -1, -1, 0, -1, 1, -1, 1, 0, 1, 1)
    For i = 0 To 7
        aDirOffsetX(i) = v(2 * i)
        aDirOffsetY(i) = v(2 * i + 1)
    Next
    v = Array(140, 140, 140, 140, 140, 140, 140, 140, 140, 140, 140, 140, 139, 139, 139, 139, 139, 139, 139, 139, 139, 138, 138, 138, 138, _
              138, 138, 137, 137, 137, 137, 137, 136, 136, 136, 136, 135, 135, 135, 134, 134, 134, 134, 133, 133, 133, 132, 132, 132, 131, _
              131, 130, 130, 130, 129, 129, 128, 128, 127, 127, 126, 126, 126, 125, 125, 124, 123, 123, 122, 122, 121, 121, 120, 119, 119, _
              118, 118, 117, 116, 116, 115, 114, 113, 113, 112, 111, 110, 110, 109, 108, 107, 106, 106, 105, 104, 103, 102, 101, 100)
    For i = 0 To 98
        gradYLookup(i) = v(i)
    Next
    aDirOffsetReady = True
End Sub

'--------------------------------------------------------------------------
' (VB) the blocking maps' slots

' PathBlockingMap::operator ==(PathBlockingType const &z)
Private Function blockingMapEquals(ByVal m As Long, ByVal gameTime_ As Long, ByVal propulsion As Long, ByVal owner As Long, ByVal moveType As Long) As Boolean
    blockingMapEquals = blockingMaps(m).typeGameTime = gameTime_ And _
                        fpathIsEquivalentBlocking(blockingMaps(m).typePropulsion, blockingMaps(m).typeOwner, blockingMaps(m).typeMoveType, _
                                                  propulsion, owner, moveType)
End Function

' (VB) std::make_shared<PathBlockingMap>(): a slot that nothing refers to
Private Function newBlockingMap() As Long
    Dim m As Long, i As Long, referenced As Boolean
    For m = 1 To blockingMaps_size
        If blockingMaps(m).used Then
            referenced = False
            For i = 0 To fpathBlockingMaps_size - 1
                If fpathBlockingMaps(i) = m Then referenced = True: Exit For
            Next
            If Not referenced Then
                For i = 0 To ctxCount - 1
                    If ctxs(i).blockingMap = m Then referenced = True: Exit For
                Next
            End If
            If Not referenced Then blockingMaps(m).used = False
        End If
        If Not blockingMaps(m).used Then Exit For
    Next
    If m > blockingMaps_size Then
        blockingMaps_size = blockingMaps_size + 1
        ReDim Preserve blockingMaps(1 To blockingMaps_size)
        m = blockingMaps_size
    End If
    blockingMaps(m).used = True
    blockingMaps(m).dangerMap_empty = True
    newBlockingMap = m
End Function

'--------------------------------------------------------------------------
' PathNonblockingArea and PathfindContext's members

' PathNonblockingArea(StructureBounds const &st)
Private Sub PathNonblockingArea_construct(a As PathNonblockingArea, st As StructureBounds)
    a.x1 = ToInt16(st.map.x)
    a.x2 = ToInt16(st.map.x + st.size.x)
    a.y1 = ToInt16(st.map.y)
    a.y2 = ToInt16(st.map.y + st.size.y)
End Sub

Private Function PathNonblockingArea_equals(a As PathNonblockingArea, z As PathNonblockingArea) As Boolean
    PathNonblockingArea_equals = a.x1 = z.x1 And a.x2 = z.x2 And a.y1 = z.y1 And a.y2 = z.y2
End Function

Private Function isBlocked(ByVal c As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim di As PathNonblockingArea
    di = ctxs(c).dstIgnore
    isBlocked = blockedAt(blockingMaps(ctxs(c).blockingMap).map, di, x, y)
End Function

Private Function matches(ByVal c As Long, ByVal blockingMap_ As Long, ByVal tileSx_ As Long, ByVal tileSy_ As Long, dstIgnore_ As PathNonblockingArea) As Boolean
    ' Must check myGameTime == blockingMap_->type.gameTime, otherwise blockingMap could be a deleted pointer which coincidentally compares equal to the valid pointer blockingMap_.
    matches = ctxs(c).myGameTime = blockingMaps(blockingMap_).typeGameTime And ctxs(c).blockingMap = blockingMap_ And _
              ctxs(c).tileSx = tileSx_ And ctxs(c).tileSy = tileSy_ And PathNonblockingArea_equals(ctxs(c).dstIgnore, dstIgnore_)
End Function

Private Sub assign(ByVal c As Long, ByVal blockingMap_ As Long, ByVal tileSx_ As Long, ByVal tileSy_ As Long, dstIgnore_ As PathNonblockingArea)
    Dim n As Long, i As Long
    ctxs(c).blockingMap = blockingMap_
    ctxs(c).tileSx = tileSx_
    ctxs(c).tileSy = tileSy_
    ctxs(c).dstIgnore = dstIgnore_
    ctxs(c).myGameTime = blockingMaps(ctxs(c).blockingMap).typeGameTime
    ctxs(c).nodes_size = 0

    ' Make the iteration not match any value of iteration in map.
    ctxs(c).iteration = (ctxs(c).iteration + 1) And &HFFFF&
    n = mapWidth * mapHeight
    If ctxs(c).iteration = &HFFFF& Or ctxs(c).map_size <> n Then
        ' There are no values of iteration guaranteed not to exist in map, so clear the map.
        ' (VB) and a map of another size is allocated afresh
        If ctxs(c).iteration = &HFFFF& Then ctxs(c).iteration = 0
        ReDim ctxs(c).mapIteration(0 To n - 1)
        ReDim ctxs(c).mapDx(0 To n - 1)
        ReDim ctxs(c).mapDy(0 To n - 1)
        ReDim ctxs(c).mapDist(0 To n - 1)
        ReDim ctxs(c).mapVisited(0 To n - 1)
        For i = 0 To n - 1
            ctxs(c).mapIteration(i) = &HFFFF&
        Next
        ctxs(c).map_size = n                ' Allocate space for map, if needed.
    End If
End Sub

Public Sub fpathHardTableReset()
    fpathBlockingMaps_size = 0
End Sub

' (VB) fpathShutdown's end of the path threads, and of their FPathExecuteContexts
Public Sub fpathContextsRelease()
    ctxCount = 0
End Sub

'--------------------------------------------------------------------------
' the node heap (std::push_heap, std::pop_heap, std::make_heap under PathNode::operator <)
' (VB) The A* procedures below take the context's arrays as parameters
' (nPx .. nEst: the nodes; mIter .. mVisited: the PathExploredTiles; blk,
' dng: the blocking map's map and dangerMap), so that the hot loops index
' plain arrays; fpathNewNode, fpathAStarReestimate and fpathAStarExplore
' pass a context's arrays to them. In the hot loops, PathNode::operator <,
' fpathEstimate, fpathGoodEstimate (iHypot) and context.isBlocked are
' written out where they are used: a VB6 procedure call costs more than
' any of them.

' PathNode::operator <: Sort descending est, fallback to ascending dist, fallback to sorting by position.
' (the heap procedures write it out; it is kept here as the reference)
Private Function nodeLess(ByVal aEst As Long, ByVal aDist As Long, ByVal aPx As Long, ByVal aPy As Long, _
                          ByVal zEst As Long, ByVal zDist As Long, ByVal zPx As Long, ByVal zPy As Long) As Boolean
    If aEst <> zEst Then
        nodeLess = (aEst Xor SIGN_BIT) > (zEst Xor SIGN_BIT)
    ElseIf aDist <> zDist Then
        nodeLess = (aDist Xor SIGN_BIT) < (zDist Xor SIGN_BIT)
    ElseIf aPx <> zPx Then
        nodeLess = aPx < zPx
    Else
        nodeLess = aPy < zPy
    End If
End Function

' (VB) restores the heap below hole i, of nodes 0 .. n - 1
Private Sub heapSiftDown(nPx() As Long, nPy() As Long, nDist() As Long, nEst() As Long, ByVal i As Long, ByVal n As Long)
    Dim child As Long, c2 As Long, less As Boolean, vPx As Long, vPy As Long, vDist As Long, vEst As Long
    vPx = nPx(i): vPy = nPy(i): vDist = nDist(i): vEst = nEst(i)
    Do
        child = 2 * i + 1
        If child >= n Then Exit Do
        c2 = child + 1
        If c2 < n Then
            ' nodes[child] < nodes[child + 1]
            If nEst(child) <> nEst(c2) Then
                less = (nEst(child) Xor SIGN_BIT) > (nEst(c2) Xor SIGN_BIT)
            ElseIf nDist(child) <> nDist(c2) Then
                less = (nDist(child) Xor SIGN_BIT) < (nDist(c2) Xor SIGN_BIT)
            ElseIf nPx(child) <> nPx(c2) Then
                less = nPx(child) < nPx(c2)
            Else
                less = nPy(child) < nPy(c2)
            End If
            If less Then child = c2
        End If
        ' v < nodes[child]
        If vEst <> nEst(child) Then
            less = (vEst Xor SIGN_BIT) > (nEst(child) Xor SIGN_BIT)
        ElseIf vDist <> nDist(child) Then
            less = (vDist Xor SIGN_BIT) < (nDist(child) Xor SIGN_BIT)
        ElseIf vPx <> nPx(child) Then
            less = vPx < nPx(child)
        Else
            less = vPy < nPy(child)
        End If
        If Not less Then Exit Do
        nPx(i) = nPx(child): nPy(i) = nPy(child): nDist(i) = nDist(child): nEst(i) = nEst(child)
        i = child
    Loop
    nPx(i) = vPx: nPy(i) = vPy: nDist(i) = vDist: nEst(i) = vEst
End Sub

' context.nodes.push_back(node); std::push_heap(...)
Private Sub heapPush(nPx() As Long, nPy() As Long, nDist() As Long, nEst() As Long, nSize As Long, nCap As Long, _
                     ByVal px As Long, ByVal py As Long, ByVal dist As Long, ByVal est As Long)
    Dim i As Long, parent As Long, less As Boolean
    If nSize = nCap Then
        If nCap = 0 Then nCap = 256 Else nCap = 2 * nCap
        ReDim Preserve nPx(0 To nCap - 1)
        ReDim Preserve nPy(0 To nCap - 1)
        ReDim Preserve nDist(0 To nCap - 1)
        ReDim Preserve nEst(0 To nCap - 1)
    End If
    i = nSize
    nSize = i + 1
    Do While i > 0
        parent = (i - 1) \ 2
        ' nodes[parent] < node
        If nEst(parent) <> est Then
            less = (nEst(parent) Xor SIGN_BIT) > (est Xor SIGN_BIT)
        ElseIf nDist(parent) <> dist Then
            less = (nDist(parent) Xor SIGN_BIT) < (dist Xor SIGN_BIT)
        ElseIf nPx(parent) <> px Then
            less = nPx(parent) < px
        Else
            less = nPy(parent) < py
        End If
        If Not less Then Exit Do
        nPx(i) = nPx(parent): nPy(i) = nPy(parent): nDist(i) = nDist(parent): nEst(i) = nEst(parent)
        i = parent
    Loop
    nPx(i) = px: nPy(i) = py: nDist(i) = dist: nEst(i) = est
End Sub

' Estimate the distance to the target point
' (newNode writes it out)
Private Function fpathEstimate(ByVal sx As Long, ByVal sy As Long, ByVal fx As Long, ByVal fy As Long) As Long
    Dim xDelta As Long, yDelta As Long
    ' Cost of moving horizontal/vertical = 70*2, cost of moving diagonal = 99*2, 99/70 = 1.41428571... ~ sqrt2 = 1.41421356...
    xDelta = Abs(sx - fx)
    yDelta = Abs(sy - fy)
    If xDelta < yDelta Then
        fpathEstimate = xDelta * (198 - 140) + yDelta * 140
    Else
        fpathEstimate = yDelta * (198 - 140) + xDelta * 140
    End If
End Function

Private Function fpathGoodEstimate(ByVal sx As Long, ByVal sy As Long, ByVal fx As Long, ByVal fy As Long) As Long
    Dim hx As Double, hy As Double, n As Double, r As Double
    ' Cost of moving horizontal/vertical = 70*2, cost of moving diagonal = 99*2, 99/70 = 1.41428571... ~ sqrt2 = 1.41421356...
    ' iHypot((s.x - f.x) * 140, (s.y - f.y) * 140), as trig_c's: i64Sqrt of a value below 2^53 (16-bit coordinates)
    hx = (sx - fx) * 140#
    hy = (sy - fy) * 140#
    n = hx * hx + hy * hy
    r = Int(Sqr(n))
    Do While r * r > n
        r = r - 1
    Loop
    Do While n - r * r > r * 2
        r = r + 1
    Loop
    fpathGoodEstimate = r
End Function

' Generate a new node
Private Sub fpathNewNode(ByVal c As Long, ByVal destX As Long, ByVal destY As Long, ByVal posX As Long, ByVal posY As Long, _
                         ByVal prevDist As Long, ByVal prevX As Long, ByVal prevY As Long)
    newNode ctxs(c).nodePx, ctxs(c).nodePy, ctxs(c).nodeDist, ctxs(c).nodeEst, ctxs(c).nodes_size, ctxs(c).nodes_capacity, _
            ctxs(c).mapIteration, ctxs(c).mapDx, ctxs(c).mapDy, ctxs(c).mapDist, ctxs(c).mapVisited, ctxs(c).iteration, _
            blockingMaps(ctxs(c).blockingMap).dangerMap, blockingMaps(ctxs(c).blockingMap).dangerMap_empty, _
            destX, destY, posX, posY, prevDist, prevX, prevY
End Sub

' (VB) fpathNewNode on the context's arrays
Private Sub newNode(nPx() As Long, nPy() As Long, nDist() As Long, nEst() As Long, nSize As Long, nCap As Long, _
                    mIter() As Long, mDx() As Integer, mDy() As Integer, mDist() As Long, mVisited() As Boolean, ByVal iteration As Long, _
                    dng() As Boolean, ByVal dangerEmpty As Boolean, _
                    ByVal destX As Long, ByVal destY As Long, ByVal posX As Long, ByVal posY As Long, _
                    ByVal prevDist As Long, ByVal prevX As Long, ByVal prevY As Long)
    Dim nodeDist As Long, nodeEst As Long, costFactor As Long, deltaX As Long, deltaY As Long, isDiagonal As Boolean, t As Long
    Dim deltaAx As Long, deltaAy As Long, deltaBx As Long, deltaBy As Long, deltaDeltaX As Long, deltaDeltaY As Long
    Dim distA As Long, distB As Long, tmp As Long, gradientX As Long, gradientY As Long, distP As Long
    Dim xDelta As Long, yDelta As Long, hx As Double, hy As Double, hn As Double, hr As Double
    If Not ((posX Xor SIGN_BIT) < (mapWidth Xor SIGN_BIT) And (posY Xor SIGN_BIT) < (mapHeight Xor SIGN_BIT)) Then ASSERT_FAILURE "(unsigned)pos.x < (unsigned)mapWidth && (unsigned)pos.y < (unsigned)mapHeight", "src/astar.cpp:273", "fpathNewNode", "X (%d) or Y (%d) coordinate for path finding node is out of range!", posX, posY: Exit Sub

    ' Create the node.
    t = posX + posY * mapWidth
    costFactor = 1                          ' context.isDangerous(pos.x, pos.y) ? 5 : 1
    If Not dangerEmpty Then
        If dng(t) Then costFactor = 5
    End If
    ' fpathEstimate(prevPos, pos)
    xDelta = Abs(prevX - posX)
    yDelta = Abs(prevY - posY)
    If xDelta < yDelta Then
        nodeDist = prevDist + (xDelta * (198 - 140) + yDelta * 140) * costFactor
    Else
        nodeDist = prevDist + (yDelta * (198 - 140) + xDelta * 140) * costFactor
    End If
    ' fpathGoodEstimate(pos, dest)
    hx = (posX - destX) * 140#
    hy = (posY - destY) * 140#
    hn = hx * hx + hy * hy
    hr = Int(Sqr(hn))
    Do While hr * hr > hn
        hr = hr - 1
    Loop
    Do While hn - hr * hr > hr * 2
        hr = hr + 1
    Loop
    nodeEst = nodeDist + hr

    deltaX = (posX - prevX) * 64
    deltaY = (posY - prevY) * 64
    isDiagonal = deltaX <> 0 And deltaY <> 0

    If mIter(t) = iteration Then
        If mVisited(t) Then
            Exit Sub                        ' Already visited this tile. Do nothing.
        End If
        deltaAx = deltaX
        deltaAy = deltaY
        deltaBx = mDx(t)
        deltaBy = mDy(t)
        deltaDeltaX = deltaAx - deltaBx     ' Vector pointing from current considered source tile leading to pos, to the previously considered source tile leading to pos.
        deltaDeltaY = deltaAy - deltaBy
        If Abs(deltaDeltaX) + Abs(deltaDeltaY) = 64 Then
            ' prevPos is tile A or B, and pos is tile P. We were previously called with prevPos being tile B or A, and pos tile P.
            ' We want to find the distance to tile P, taking into account that the actual shortest path involves coming from somewhere between tile A and tile B.
            ' +---+---+
            ' |   | P |
            ' +---+---+
            ' | A | B |
            ' +---+---+
            If isDiagonal Then                ' If isDiagonal, node is A and expl is B.
                distA = nodeDist - 198 * costFactor
                distB = mDist(t) - 140 * costFactor
            Else
                distA = nodeDist - 140 * costFactor
                distB = mDist(t) - 198 * costFactor
            End If
            If Not isDiagonal Then
                tmp = distA: distA = distB: distB = tmp
                tmp = deltaAx: deltaAx = deltaBx: deltaBx = tmp
                tmp = deltaAy: deltaAy = deltaBy: deltaBy = tmp
            End If
            ' int(distB - distA) / costFactor: the division is unsigned (costFactor is unsigned), so a negative
            ' difference gives no gradient in 1 .. 98
            If distB - distA > 0 Then gradientX = (distB - distA) \ costFactor Else gradientX = 0
            If gradientX > 0 And gradientX <= 98 Then   ' 98 = floor(140/sqrt2), so gradientX <= 98 is needed so that gradientX < gradientY.
                ' The distance gradient is now known to be somewhere between the direction from A to P and the direction from B to P.
                gradientY = gradYLookup(gradientX)  ' = sqrt(140^2 -  gradientX^2), rounded to nearest integer
                distP = gradientY * costFactor + distB
                nodeEst = nodeEst - (nodeDist - distP)
                nodeDist = distP
                deltaX = (deltaAx * gradientX + deltaBx * (gradientY - gradientX)) \ gradientY
                deltaY = (deltaAy * gradientX + deltaBy * (gradientY - gradientX)) \ gradientY
            End If
        End If
        If Not ((nodeDist Xor SIGN_BIT) < (mDist(t) Xor SIGN_BIT)) Then
            Exit Sub                        ' A different path to this tile is shorter.
        End If
    End If

    ' Remember where we have been, and remember the way back.
    mIter(t) = iteration
    mDx(t) = deltaX
    mDy(t) = deltaY
    mDist(t) = nodeDist
    mVisited(t) = False

    ' Add the node to the node heap.
    heapPush nPx, nPy, nDist, nEst, nSize, nCap, posX, posY, nodeDist, nodeEst  ' Add the new node to nodes, and move it to the right place in the heap.
End Sub

' Recalculates estimates to new tileF tile.
Private Sub fpathAStarReestimate(ByVal c As Long, ByVal tileFx As Long, ByVal tileFy As Long)
    reestimate ctxs(c).nodePx, ctxs(c).nodePy, ctxs(c).nodeDist, ctxs(c).nodeEst, ctxs(c).nodes_size, tileFx, tileFy
End Sub

' (VB) fpathAStarReestimate on the context's arrays
Private Sub reestimate(nPx() As Long, nPy() As Long, nDist() As Long, nEst() As Long, ByVal nSize As Long, ByVal tileFx As Long, ByVal tileFy As Long)
    Dim i As Long
    For i = 0 To nSize - 1
        nEst(i) = nDist(i) + fpathGoodEstimate(nPx(i), nPy(i), tileFx, tileFy)
    Next

    ' Changing the estimates breaks the heap ordering. Fix the heap ordering.
    For i = nSize \ 2 - 1 To 0 Step -1
        heapSiftDown nPx, nPy, nDist, nEst, i, nSize
    Next
End Sub

' Returns nearest explored tile to tileF (in nearestX, nearestY).
Private Sub fpathAStarExplore(ByVal c As Long, ByVal tileFx As Long, ByVal tileFy As Long, nearestX As Long, nearestY As Long)
    Dim di As PathNonblockingArea
    di = ctxs(c).dstIgnore
    explore ctxs(c).nodePx, ctxs(c).nodePy, ctxs(c).nodeDist, ctxs(c).nodeEst, ctxs(c).nodes_size, ctxs(c).nodes_capacity, _
            ctxs(c).mapIteration, ctxs(c).mapDx, ctxs(c).mapDy, ctxs(c).mapDist, ctxs(c).mapVisited, ctxs(c).iteration, _
            blockingMaps(ctxs(c).blockingMap).map, blockingMaps(ctxs(c).blockingMap).dangerMap, blockingMaps(ctxs(c).blockingMap).dangerMap_empty, _
            di, tileFx, tileFy, nearestX, nearestY
End Sub

' (VB) context.isBlocked(x, y) on the blocking map's array (explore writes it out)
Private Function blockedAt(blk() As Boolean, di As PathNonblockingArea, ByVal x As Long, ByVal y As Long) As Boolean
    If x >= di.x1 And x < di.x2 And y >= di.y1 And y < di.y2 Then
        blockedAt = False                   ' The path is actually blocked here by a structure, but ignore it since it's where we want to go (or where we came from).
        Exit Function
    End If
    ' Not sure whether the out-of-bounds check is needed, can only happen if pathfinding is started on a blocking tile (or off the map).
    If x < 0 Or y < 0 Or x >= mapWidth Or y >= mapHeight Then
        blockedAt = True
    Else
        blockedAt = blk(x + y * mapWidth)
    End If
End Function

' (VB) fpathAStarExplore on the context's arrays
Private Sub explore(nPx() As Long, nPy() As Long, nDist() As Long, nEst() As Long, nSize As Long, nCap As Long, _
                    mIter() As Long, mDx() As Integer, mDy() As Integer, mDist() As Long, mVisited() As Boolean, ByVal iteration As Long, _
                    blk() As Boolean, dng() As Boolean, ByVal dangerEmpty As Boolean, di As PathNonblockingArea, _
                    ByVal tileFx As Long, ByVal tileFy As Long, nearestX As Long, nearestY As Long)
    Dim nearestDist As Long, foundIt As Boolean, t As Long, dir_ As Long, n As Long, w As Long, h As Long
    Dim nodePx As Long, nodePy As Long, nodeDist As Long, nodeEst As Long, blocked As Boolean
    Dim x As Long, y As Long, x2 As Long, y2 As Long, ix1 As Long, ix2 As Long, iy1 As Long, iy2 As Long
    w = mapWidth
    h = mapHeight
    ix1 = di.x1: ix2 = di.x2: iy1 = di.y1: iy2 = di.y2
    nearestX = 0
    nearestY = 0
    nearestDist = &HFFFFFFFF

    ' search for a route
    foundIt = False
    Do While nSize > 0 And Not foundIt
        ' fpathTakeNode: find the node with the lowest distance; if equal totals, give preference to node closer to target
        nodePx = nPx(0): nodePy = nPy(0): nodeDist = nDist(0): nodeEst = nEst(0)
        ' remove the node from the list (std::pop_heap, nodes.pop_back())
        n = nSize - 1
        nSize = n
        If n > 0 Then
            nPx(0) = nPx(n): nPy(0) = nPy(n): nDist(0) = nDist(n): nEst(0) = nEst(n)
            heapSiftDown nPx, nPy, nDist, nEst, 0, n
        End If

        t = nodePx + nodePy * w
        If mVisited(t) Then
            GoTo continue_                  ' Already been here.
        End If
        mVisited(t) = True

        ' note the nearest node to the target so far
        If ((nodeEst - nodeDist) Xor SIGN_BIT) < (nearestDist Xor SIGN_BIT) Then
            nearestX = nodePx
            nearestY = nodePy
            nearestDist = nodeEst - nodeDist
        End If

        If nodePx = tileFx And nodePy = tileFy Then
            ' reached the target
            nearestX = nodePx
            nearestY = nodePy
            foundIt = True                  ' Break out of loop, but not before inserting neighbour nodes, since the neighbours may be important if the context gets reused.
        End If

        ' loop through possible moves in 8 directions to find a valid move
        For dir_ = 0 To 7
            ' Try a new location
            x = nodePx + aDirOffsetX(dir_)
            y = nodePy + aDirOffsetY(dir_)

            '   5  6  7
            '     \|/
            '   4 -I- 0
            '     /|\
            '   3  2  1
            '   odd:orthogonal-adjacent tiles even:non-orthogonal-adjacent tiles
            If (dir_ And 1) <> 0 Then
                If Not (nodePx >= ix1 And nodePx < ix2 And nodePy >= iy1 And nodePy < iy2) And _
                   Not (x >= ix1 And x < ix2 And y >= iy1 And y < iy2) Then
                    ' We cannot cut corners
                    x2 = nodePx + aDirOffsetX((dir_ + 1) And 7)
                    y2 = nodePy + aDirOffsetY((dir_ + 1) And 7)
                    ' context.isBlocked(x2, y2)
                    If x2 >= ix1 And x2 < ix2 And y2 >= iy1 And y2 < iy2 Then
                        blocked = False
                    ElseIf x2 < 0 Or y2 < 0 Or x2 >= w Or y2 >= h Then
                        blocked = True
                    Else
                        blocked = blk(x2 + y2 * w)
                    End If
                    If blocked Then
                        GoTo next_dir
                    End If
                    x2 = nodePx + aDirOffsetX((dir_ + 7) And 7)
                    y2 = nodePy + aDirOffsetY((dir_ + 7) And 7)
                    If x2 >= ix1 And x2 < ix2 And y2 >= iy1 And y2 < iy2 Then
                        blocked = False
                    ElseIf x2 < 0 Or y2 < 0 Or x2 >= w Or y2 >= h Then
                        blocked = True
                    Else
                        blocked = blk(x2 + y2 * w)
                    End If
                    If blocked Then
                        GoTo next_dir
                    End If
                End If
            End If

            ' See if the node is a blocking tile
            If x >= ix1 And x < ix2 And y >= iy1 And y < iy2 Then
                blocked = False
            ElseIf x < 0 Or y < 0 Or x >= w Or y >= h Then
                blocked = True
            Else
                blocked = blk(x + y * w)
            End If
            If blocked Then
                ' tile is blocked, skip it
                GoTo next_dir
            End If

            ' Now insert the point into the appropriate list, if not already visited.
            ' (VB) newNode returns at once, changing nothing, for a visited tile: that test is made here first
            t = x + y * w
            If mIter(t) = iteration Then
                If mVisited(t) Then GoTo next_dir
            End If
            newNode nPx, nPy, nDist, nEst, nSize, nCap, mIter, mDx, mDy, mDist, mVisited, iteration, dng, dangerEmpty, _
                    tileFx, tileFy, x, y, nodeDist, nodePx, nodePy
next_dir:
        Next
continue_:
    Loop
End Sub

Private Sub fpathInitContext(ByVal c As Long, ByVal blockingMap As Long, ByVal tileSx As Long, ByVal tileSy As Long, ByVal tileRealSx As Long, ByVal tileRealSy As Long, _
                             ByVal tileFx As Long, ByVal tileFy As Long, dstIgnore As PathNonblockingArea)
    assign c, blockingMap, tileSx, tileSy, dstIgnore

    ' Add the start point to the open list
    fpathNewNode c, tileFx, tileFy, tileRealSx, tileRealSy, 0, tileRealSx, tileRealSy
    If Not (ctxs(c).nodes_size > 0) Then ASSERT_FAILURE "!context.nodes.empty()", "src/astar.cpp:439", "fpathInitContext", "fpathNewNode failed to add node."
End Sub

'--------------------------------------------------------------------------
' PathfindContextList's members

' push_back(PathfindContext()): the new context's position in the order
Private Function contexts_push_back() As Long
    Dim c As Long
    If ctxCount = ctxs_capacity Then
        If ctxs_capacity = 0 Then ctxs_capacity = 8 Else ctxs_capacity = 2 * ctxs_capacity
        ReDim Preserve ctxs(0 To ctxs_capacity - 1)
        ReDim Preserve ctxOrder(0 To ctxs_capacity - 1)
    End If
    c = ctxCount
    ' PathfindContext() : myGameTime(0), iteration(0), blockingMap(nullptr) (the slot keeps its iteration count and arrays)
    ctxs(c).tileSx = 0
    ctxs(c).tileSy = 0
    ctxs(c).myGameTime = 0
    ctxs(c).nearestX = 0
    ctxs(c).nearestY = 0
    ctxs(c).nodes_size = 0
    ctxs(c).blockingMap = 0
    ctxs(c).dstIgnore.x1 = 0
    ctxs(c).dstIgnore.x2 = 0
    ctxs(c).dstIgnore.y1 = 0
    ctxs(c).dstIgnore.y2 = 0
    ctxOrder(ctxCount) = c
    ctxCount = ctxCount + 1
    contexts_push_back = ctxCount - 1
End Function

' moveToFront(it): std::rotate(orderedIndexes.begin(), oIt, oIt + 1)
Private Sub contexts_moveToFront(ByVal idx As Long)
    Dim i As Long, v As Long
    v = ctxOrder(idx)
    For i = idx To 1 Step -1
        ctxOrder(i) = ctxOrder(i - 1)
    Next
    ctxOrder(0) = v
End Sub

' FPathExecuteContextImpl::resetForNewGameTimeIfNeeded(job)
Private Sub resetForNewGameTimeIfNeeded(job As PATHJOB)
    If ctxCount > 0 Then
        If blockingMaps(job.blockingMap).typeGameTime <> ctxs(ctxOrder(0)).myGameTime Then
            ctxCount = 0                    ' fpathContexts.clear()
        End If
    End If
End Sub

' Use the A* algorithm to find a path
' (VB) the FPathExecuteContext is this module's
Public Function fpathAStarRoute(psMove As MOVE_CONTROL, psJob As PATHJOB) As ASR_RETVAL
    Dim retval As ASR_RETVAL, mustReverse As Boolean
    Dim tileOrigX As Long, tileOrigY As Long, tileDestX As Long, tileDestY As Long, dstIgnore As PathNonblockingArea
    Dim endX As Long, endY As Long, it As Long, c As Long, t As Long
    Dim path() As Vector2i, path_size As Long, p As Vector2i, newP As Vector2i, mapP As Vector2i
    Dim xSide As Long, ySide As Long, i As Long

    If Not aDirOffsetReady Then initTables  ' (VB) the static tables
    retval = ASR_OK

    mustReverse = True

    resetForNewGameTimeIfNeeded psJob

    tileOrigX = map_coord(psJob.origX)
    tileOrigY = map_coord(psJob.origY)
    tileDestX = map_coord(psJob.destX)
    tileDestY = map_coord(psJob.destY)
    PathNonblockingArea_construct dstIgnore, psJob.dstStructure

    ' endCoord: Either nearest coord (mustReverse = true) or orig (mustReverse = false).

    For it = 0 To ctxCount - 1
        c = ctxOrder(it)
        If Not matches(c, psJob.blockingMap, tileDestX, tileDestY, dstIgnore) Then
            ' This context is not for the same droid type and same destination.
            GoTo next_context
        End If

        ' We have tried going to tileDest before.

        t = tileOrigX + tileOrigY * mapWidth
        If ctxs(c).mapIteration(t) = ctxs(c).iteration And ctxs(c).mapVisited(t) Then
            ' Already know the path from orig to dest.
            endX = tileOrigX
            endY = tileOrigY
        Else
            ' Need to find the path from orig to dest, continue previous exploration.
            fpathAStarReestimate c, tileOrigX, tileOrigY
            fpathAStarExplore c, tileOrigX, tileOrigY, endX, endY
        End If

        If endX <> tileOrigX Or endY <> tileOrigY Then
            ' orig turned out to be on a different island than what this context was used for, so can't use this context data after all.
            GoTo next_context
        End If

        mustReverse = False                 ' We have the path from the nearest reachable tile to dest, to orig.
        Exit For                            ' Found the path! Don't search more contexts.
next_context:
    Next

    If it = ctxCount Then
        ' We did not find an appropriate context. Make one.
        it = contexts_push_back()
        c = ctxOrder(it)

        ' Init a new context, overwriting the oldest one if we are caching too many.
        ' We will be searching from orig to dest, since we don't know where the nearest reachable tile to dest is.
        fpathInitContext c, psJob.blockingMap, tileOrigX, tileOrigY, tileOrigX, tileOrigY, tileDestX, tileDestY, dstIgnore
        fpathAStarExplore c, tileDestX, tileDestY, endX, endY
        ctxs(c).nearestX = endX
        ctxs(c).nearestY = endY
    End If

    ' return the nearest route if no actual route was found
    If ctxs(c).nearestX <> tileDestX Or ctxs(c).nearestY <> tileDestY Then
        retval = ASR_NEAREST
    End If

    ' Get route, in reverse order.
    path_size = 0
    ReDim path(0 To 63)

    p.x = world_coord(endX) + TILE_UNITS \ 2
    p.y = world_coord(endY) + TILE_UNITS \ 2
    Do
        If Not (worldOnMap(p.x, p.y)) Then ASSERT_FAILURE "worldOnMap(p.x, p.y)", "src/astar.cpp:656", "fpathAStarRoute", "Assigned XY coordinates (%d, %d) not on map!", p.x, p.y: fpathAStarRoute = ASR_FAILED: Exit Function
        If Not (path_size < mapWidth * mapHeight) Then ASSERT_FAILURE "path.size() < (static_cast<size_t>(mapWidth) * static_cast<size_t>(mapHeight))", "src/astar.cpp:657", "fpathAStarRoute", "Pathfinding got in a loop.": fpathAStarRoute = ASR_FAILED: Exit Function

        If path_size > UBound(path) Then ReDim Preserve path(0 To 2 * path_size - 1)
        path(path_size) = p
        path_size = path_size + 1

        t = map_coord(p.x) + map_coord(p.y) * mapWidth
        newP.x = p.x - ctxs(c).mapDx(t) * (TILE_UNITS \ 64)
        newP.y = p.y - ctxs(c).mapDy(t) * (TILE_UNITS \ 64)
        mapP.x = map_coord(newP.x)
        mapP.y = map_coord(newP.y)
        If newP.x - world_coord(mapP.x) > TILE_UNITS \ 2 Then xSide = 1 Else xSide = -1     ' 1 if newP is on right-hand side of the tile, or -1 if newP is on the left-hand side of the tile.
        If newP.y - world_coord(mapP.y) > TILE_UNITS \ 2 Then ySide = 1 Else ySide = -1     ' 1 if newP is on bottom side of the tile, or -1 if newP is on the top side of the tile.
        If isBlocked(c, mapP.x + xSide, mapP.y) Then
            newP.x = world_coord(mapP.x) + TILE_UNITS \ 2  ' Point too close to a blocking tile on left or right side, so move the point to the middle.
        End If
        If isBlocked(c, mapP.x, mapP.y + ySide) Then
            newP.y = world_coord(mapP.y) + TILE_UNITS \ 2  ' Point too close to a blocking tile on rop or bottom side, so move the point to the middle.
        End If
        If (map_coord(p.x) = ctxs(c).tileSx And map_coord(p.y) = ctxs(c).tileSy) Or (p.x = newP.x And p.y = newP.y) Then
            Exit Do                         ' We stopped moving, because we reached the destination or the closest reachable tile to context.tileS. Give up now.
        End If
        p = newP
    Loop
    If retval = ASR_OK Then
        ' Found exact path, so use exact coordinates for last point, no reason to lose precision
        If mustReverse Then
            path(0).x = psJob.destX
            path(0).y = psJob.destY
        Else
            path(path_size - 1).x = psJob.destX
            path(path_size - 1).y = psJob.destY
        End If
    End If

    ' Allocate memory
    ReDim psMove.asPath(0 To path_size - 1)
    psMove.asPath_size = path_size

    ' get the route in the correct order
    ' If as I suspect this is to reverse the list, then it's my suspicion that
    ' we could route from destination to source as opposed to source to
    ' destination. We could then save the reversal. to risky to try now...Alex M
    '
    ' The idea is impractical, because you can't guarentee that the target is
    ' reachable. As I see it, this is the reason why psNearest got introduced.
    ' -- Dennis L.
    '
    ' If many droids are heading towards the same destination, then destination
    ' to source would be faster if reusing the information in nodeArray. --Cyp
    If mustReverse Then
        ' Copy the list, in reverse.
        For i = 0 To path_size - 1
            psMove.asPath(i) = path(path_size - 1 - i)
        Next

        If Not isBlocked(c, tileOrigX, tileOrigY) Then  ' If blocked, searching from tileDest to tileOrig wouldn't find the tileOrig tile.
            ' Next time, search starting from nearest reachable tile to the destination.
            fpathInitContext c, psJob.blockingMap, tileDestX, tileDestY, ctxs(c).nearestX, ctxs(c).nearestY, tileOrigX, tileOrigY, dstIgnore
        End If
    Else
        ' Copy the list.
        For i = 0 To path_size - 1
            psMove.asPath(i) = path(i)
        Next
    End If

    ' Move context to beginning of last recently used list.
    If it <> 0 Then                         ' Not sure whether or not the splice is a safe noop, if equal.
        contexts_moveToFront it
    End If

    psMove.destination = psMove.asPath(path_size - 1)

    fpathAStarRoute = retval
End Function

' (VB) the checksums' factor = 3 * factor + 1 (uint32_t)
Private Function nextFactor(ByVal factor As Long) As Long
    factor = MUL32(3, factor)
    If factor = &H7FFFFFFF Then nextFactor = &H80000000 Else nextFactor = factor + 1
End Function

' Call from main thread.
' Sets psJob->blockingMap for later use by pathfinding thread, generating the required map if not already generated.
Public Sub fpathSetBlockingMap(psJob As PATHJOB)
    Dim i As Long, blockMap As Long, x As Long, y As Long, n As Long
    Dim checksumMap As Long, checksumDangerMap As Long, factor As Long
    If fpathCurrentGameTime <> gameTime Then
        ' New tick, remove maps which are no longer needed.
        fpathCurrentGameTime = gameTime
        fpathBlockingMaps_size = 0
    End If

    ' Figure out which map we are looking for: PathBlockingType type = {gameTime, psJob->propulsion, psJob->owner, psJob->moveType}

    ' Find the map.
    For i = 0 To fpathBlockingMaps_size - 1
        If blockingMapEquals(fpathBlockingMaps(i), gameTime, psJob.propulsion, psJob.owner, psJob.moveType) Then Exit For
    Next
    If i = fpathBlockingMaps_size Then
        ' Didn't find the map, so i does not point to a map.
        blockMap = newBlockingMap()
        If fpathBlockingMaps_size = 0 Then
            ReDim fpathBlockingMaps(0 To 7)
        ElseIf fpathBlockingMaps_size > UBound(fpathBlockingMaps) Then
            ReDim Preserve fpathBlockingMaps(0 To 2 * fpathBlockingMaps_size - 1)
        End If
        fpathBlockingMaps(fpathBlockingMaps_size) = blockMap
        fpathBlockingMaps_size = fpathBlockingMaps_size + 1

        ' blockMap now points to an empty map with no data. Fill the map.
        blockingMaps(blockMap).typeGameTime = gameTime
        blockingMaps(blockMap).typePropulsion = psJob.propulsion
        blockingMaps(blockMap).typeOwner = psJob.owner
        blockingMaps(blockMap).typeMoveType = psJob.moveType
        n = mapWidth * mapHeight
        ReDim blockingMaps(blockMap).map(0 To n - 1)
        checksumMap = 0
        checksumDangerMap = 0
        factor = 0
        For y = 0 To mapHeight - 1
            For x = 0 To mapWidth - 1
                blockingMaps(blockMap).map(x + y * mapWidth) = fpathBaseBlockingTile(x, y, blockingMaps(blockMap).typePropulsion, _
                    blockingMaps(blockMap).typeOwner, blockingMaps(blockMap).typeMoveType)
                factor = nextFactor(factor)
                If blockingMaps(blockMap).map(x + y * mapWidth) Then checksumMap = checksumMap Xor factor
            Next
        Next
        If Not isHumanPlayer(blockingMaps(blockMap).typeOwner) And blockingMaps(blockMap).typeMoveType = FMT_MOVE Then
            ReDim blockingMaps(blockMap).dangerMap(0 To n - 1)
            blockingMaps(blockMap).dangerMap_empty = False
            For y = 0 To mapHeight - 1
                For x = 0 To mapWidth - 1
                    blockingMaps(blockMap).dangerMap(x + y * mapWidth) = (auxTile(x, y, blockingMaps(blockMap).typeOwner) And AUXBITS_THREAT) <> 0
                    factor = nextFactor(factor)
                    If blockingMaps(blockMap).dangerMap(x + y * mapWidth) Then checksumDangerMap = checksumDangerMap Xor factor
                Next
            Next
        End If
        syncDebug "fpathSetBlockingMap", "blockingMap(%d,%d,%d,%d) = %08X %08X", gameTime, psJob.propulsion, psJob.owner, psJob.moveType, checksumMap, checksumDangerMap

        psJob.blockingMap = blockMap
    Else
        syncDebug "fpathSetBlockingMap", "blockingMap(%d,%d,%d,%d) = cached", gameTime, psJob.propulsion, psJob.owner, psJob.moveType

        psJob.blockingMap = fpathBlockingMaps(i)
    End If
End Sub
