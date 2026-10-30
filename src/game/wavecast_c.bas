Attribute VB_Name = "wavecast_c"
' src/wavecast.cpp, wavecast.h: the tables of tiles within a sensor radius,
' with the angles they cover, for the visibility wave cast. The tables
' (std::map<unsigned, std::vector<WavecastTile>>) are wavecastTables, one
' per radius (wavecastIndex: radius -> its index); getWavecastTable returns
' the index of a table, whose tiles are wavecastTables(t).tiles(0 ..
' size - 1). size_t angles are Longs (they count the angles of a table).
Option Explicit

Public Type WavecastTile
    dx As Long                          ' int16_t: Tile coordinates.
    dy As Long
    invRadius As Long                   ' Arbitrary constant divided by radius.
    angBegin As Long                    ' size_t: Start and finish angles for obstruction of view. Non-linear units, for comparison purposes only.
    angEnd As Long
End Type

Public Type WAVECAST_TABLE
    tiles() As WavecastTile
    size As Long
End Type

' Angles are sorted in this order. Can only be created and compared to each other, nothing else.
' (1, 0) < (0, 1) < (-1, 0) < (0, -1) < (1, -e) < (0, 0)
' (0, 0) = special case, compares greater than all other angles.
Private Type RationalAngle
    myX As Long
    myY As Long
End Type

Public wavecastTables() As WAVECAST_TABLE
Private wavecastTables_count As Long
Private wavecastIndex As Object         ' radius -> index in wavecastTables

Private Function RationalAngle_(ByVal x As Long, ByVal y As Long) As RationalAngle
    RationalAngle_.myX = x
    RationalAngle_.myY = y
End Function

Private Function quadrant(a As RationalAngle, x As Long, y As Long) As Long
    Dim q As Long
    ' Quadrant is arbitrarily chosen for myX == 0 || myY == 0, just needs to be consistent. Quadrant is 4 for myX == 0 && myY == 0.
    If a.myY >= 0 Then
        If a.myX >= 0 Then
            If a.myX <> 0 Or a.myY <> 0 Then q = 0 Else q = 4
        Else
            q = 1
        End If
    Else
        If a.myX >= 0 Then q = 3 Else q = 2
    End If
    Select Case q
    Case 0: x = a.myX: y = a.myY
    Case 1: x = a.myY: y = -a.myX
    Case 2: x = -a.myX: y = -a.myY
    Case 3: x = -a.myY: y = a.myX
    Case 4: x = 0: y = 0    ' myX and myY are both 0, set x, y to arbitrary constant so (0, 0) compares equal to itself.
    End Select
    quadrant = q
End Function

' bool operator <(RationalAngle const &v2) const
Private Function angleLess(a As RationalAngle, b As RationalAngle) As Boolean
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, qa As Long, qb As Long
    qa = quadrant(a, x1, y1)
    qb = quadrant(b, x2, y2)
    If qa <> qb Then
        angleLess = qa < qb
    Else
        angleLess = y1 * x2 < y2 * x1  ' Same as y1/x1 < y2/x2, using real numbers. (And works even with x1 or x2 = 0.)
    End If
End Function

' bool operator ==(RationalAngle const &v2) const
Private Function angleEqual(a As RationalAngle, b As RationalAngle) As Boolean
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, qa As Long, qb As Long
    qa = quadrant(a, x1, y1)
    qb = quadrant(b, x2, y2)
    angleEqual = qa = qb And y1 * x2 = y2 * x1  ' Same as y1/x1 == y2/x2, using real numbers. (And works even with x1 or x2 = 0.)
End Function

' (VB) std::sort of the angles (the order of equal angles does not matter: they become one)
Private Sub sortAngles(a() As RationalAngle, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As RationalAngle, t As RationalAngle
    Do While hi - lo > 12
        p = a((lo + hi) \ 2)
        i = lo
        j = hi
        Do While i <= j
            Do While angleLess(a(i), p): i = i + 1: Loop
            Do While angleLess(p, a(j)): j = j - 1: Loop
            If i <= j Then
                t = a(i): a(i) = a(j): a(j) = t
                i = i + 1
                j = j - 1
            End If
        Loop
        If j - lo < hi - i Then
            sortAngles a, lo, j
            lo = i
        Else
            sortAngles a, i, hi
            hi = j
        End If
    Loop
    For i = lo + 1 To hi
        t = a(i)
        j = i - 1
        Do While j >= lo
            If Not angleLess(t, a(j)) Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next
End Sub

' std::lower_bound: the first index whose angle is not less than v
Private Function angleLowerBound(a() As RationalAngle, ByVal n As Long, v As RationalAngle) As Long
    Dim lo As Long, hi As Long, m As Long
    lo = 0
    hi = n
    Do While lo < hi
        m = (lo + hi) \ 2
        If angleLess(a(m), v) Then lo = m + 1 Else hi = m
    Loop
    angleLowerBound = lo
End Function

' Create a list of all tiles within the given radius, and the leftmost and rightmost angles of the tiles, for field of vision.
' Radius in tiles*TILE_UNITS.
' The droid is assumed to be at (dx = 0.5, dy = 0.5)*TILE_UNITS, not at the origin, since tile data is for the top left of the tile.
Private Sub generateWavecastTable(ByVal radius As Long, tbl As WAVECAST_TABLE)
    Dim tiles() As WavecastTile, nTiles As Long, tile As WavecastTile
    Dim unsortedAngles() As RationalAngle, nAngles As Long, sortedAngles() As RationalAngle, nSorted As Long
    Dim diamond As Long, quad_ As Long, s As Long, sdx As Long, sdy As Long, tileRadiusSq As Long
    Dim mcdx As Long, mcdy As Long, minAngle As RationalAngle, maxAngle As RationalAngle, i As Long
    Dim minCorner As Variant

    minCorner = Array(Array(1, 0), Array(1, 1), Array(0, 1), Array(0, 0))
    ReDim tiles(0 To 63)
    ReDim unsortedAngles(0 To 127)

    diamond = 1
    Do While U32(diamond * TILE_UNITS) < U32(radius) * 2   ' Factor is a bit more than needed to surround the circle with diamonds.
        For quad_ = 0 To 3
            For s = 0 To diamond - 1
                Select Case quad_
                Case 0: tile.dx = diamond - s: tile.dy = s + 1
                Case 1: tile.dx = -s: tile.dy = diamond - s
                Case 2: tile.dx = -diamond + s + 1: tile.dy = -s
                Case 3: tile.dx = s + 1: tile.dy = -diamond + s + 1
                End Select

                sdx = tile.dx * 2 - 1   ' 2*distance from sensor located at (0.5, 0.5)
                sdy = tile.dy * 2 - 1
                tileRadiusSq = sdx * sdx + sdy * sdy

                If ULT(MUL32(radius, radius), MUL32(tileRadiusSq, TILE_UNITS * TILE_UNITS \ 4)) Then   ' (unsigned)
                    GoTo continue_
                End If
                tile.invRadius = i64Sqrt(Int(8589934592# / tileRadiusSq)) ' Result is at most 65536, inversely proportional to the radius.

                mcdx = minCorner(quad_)(0)   ' Corner of the tile which the minimum angle.
                mcdy = minCorner(quad_)(1)
                minAngle = RationalAngle_(tile.dx - 1 + mcdx, tile.dy - 1 + mcdy)
                maxAngle = RationalAngle_(tile.dx - mcdx, tile.dy - mcdy)
                If angleLess(maxAngle, minAngle) Then
                    maxAngle = RationalAngle_(0, 0)  ' Special case, like RationalAngle(1, 0), except that it compares greater than all other angles instead of less than all other angles.
                End If
                If nAngles + 2 > UBound(unsortedAngles) + 1 Then ReDim Preserve unsortedAngles(0 To nAngles * 2 + 1)
                tile.angBegin = nAngles
                unsortedAngles(nAngles) = minAngle
                nAngles = nAngles + 1
                tile.angEnd = nAngles
                unsortedAngles(nAngles) = maxAngle
                nAngles = nAngles + 1

                If nTiles > UBound(tiles) Then ReDim Preserve tiles(0 To nTiles * 2 - 1)
                tiles(nTiles) = tile
                nTiles = nTiles + 1
continue_:
            Next
        Next
        diamond = diamond + 1
    Loop

    ' Sort the angles and remove duplicates.
    If nAngles > 0 Then
        sortedAngles = unsortedAngles
        sortAngles sortedAngles, 0, nAngles - 1
        nSorted = 1
        For i = 1 To nAngles - 1
            If Not angleEqual(sortedAngles(nSorted - 1), sortedAngles(i)) Then
                sortedAngles(nSorted) = sortedAngles(i)
                nSorted = nSorted + 1
            End If
        Next
    End If

    ' Subtitute the angle values angBegin and angEnd with ones that can be compared to each other, so that
    ' the angles can be compared without using the unsortedAngles lookup table. (And without using the
    ' sortedAngles lookup table either.)
    For i = 0 To nTiles - 1
        tiles(i).angBegin = angleLowerBound(sortedAngles, nSorted, unsortedAngles(tiles(i).angBegin))
        tiles(i).angEnd = angleLowerBound(sortedAngles, nSorted, unsortedAngles(tiles(i).angEnd))
    Next

    If nTiles > 0 Then ReDim Preserve tiles(0 To nTiles - 1)
    tbl.tiles = tiles
    tbl.size = nTiles
End Sub

' Not thread safe if someone calls with a new radius. Thread safe, otherwise.
' (VB) returns the table's index in wavecastTables
Public Function getWavecastTable(ByVal radius As Long, size As Long) As Long
    Dim t As Long
    If wavecastIndex Is Nothing Then
        Set wavecastIndex = CreateObject("Scripting.Dictionary")
    End If
    If wavecastIndex.Exists(radius) Then
        t = wavecastIndex.Item(radius)
    Else
        ' Table not generated yet.
        t = wavecastTables_count
        ReDim Preserve wavecastTables(0 To t)
        wavecastTables_count = t + 1
        generateWavecastTable radius, wavecastTables(t)
        wavecastIndex.Add radius, t
    End If
    size = wavecastTables(t).size
    getWavecastTable = t
End Function
