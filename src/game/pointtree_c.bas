Attribute VB_Name = "pointtree_c"
' src/pointtree.cpp, pointtree.h: class PointTree, of which the game has one
' (mapgrid's gridPointTree): its state is this module's, and its methods
' are PointTree_<method>. A point's data is an object handle. The uint64_t
' Morton numbers are two Longs (pointsHi, pointsLo), kept with the sign bit
' flipped so that signed comparisons order them as unsigned. A query tests
' "px >= minX && px <= maxX && py >= minY && py <= maxY" on the point's own
' x and y (pointsX, pointsY): expandX and expandY keep the order of
' x + 0x80000000 and y + 0x80000000, so the test is the same as the signed
' comparison of the coordinates. PointTree::Filter is the UDT PointTreeFilter
' and its methods PointTreeFilter_<method>.
Option Explicit

' class PointTree::Filter: Filters are invalidated when modifying the PointTree.
Public Type PointTreeFilter
    data() As Long                              ' std::vector<unsigned>
    data_size As Long
End Type

' ResultVector lastQueryResults, IndexVector lastFilteredQueryIndices
Public lastQueryResults() As Long
Public lastQueryResults_size As Long
Public lastFilteredQueryIndices() As Long
Public lastFilteredQueryIndices_size As Long
Private lastQueryResults_capacity As Long       ' (VB) the arrays' sizes
Private lastFilteredQueryIndices_capacity As Long

' Vector points: std::vector<std::pair<uint64_t, void *>>
Private pointsHi() As Long                      ' the Morton number's high 32 bits, sign bit flipped
Private pointsLo() As Long                      ' its low 32 bits, sign bit flipped
Private pointsData() As Long                    ' the object
Private pointsX() As Long                       ' (VB) the x and y it was inserted with
Private pointsY() As Long
Private points_size As Long
Private points_capacity As Long

Private expandTable(0 To 255) As Long           ' (VB) expand() of a byte: abcd efgh to 0a0b 0c0d 0e0f 0g0h
Private expandReady As Boolean

Private sortTmpHi() As Long                     ' (VB) sort's merge buffers
Private sortTmpLo() As Long
Private sortTmpData() As Long
Private sortTmpX() As Long
Private sortTmpY() As Long

' (VB) a uint64_t as two Longs (raw, not flipped)
Private Type U64
    hi As Long
    lo As Long
End Type

' struct PointTreeRange
Private Type PointTreeRange
    a As U64
    z As U64
End Type

'--------------------------------------------------------------------------
' PointTree::Filter's members

' Filter(): Must be reset before use.
Public Sub PointTreeFilter_construct(f As PointTreeFilter)
    ReDim f.data(0 To 0)
    f.data_size = 1
End Sub

Public Sub PointTreeFilter_reset(f As PointTreeFilter)
    Dim i As Long
    ' data.assign(pointTree.points.size() + 1, 0)
    If f.data_size < points_size + 1 Then
        ReDim f.data(0 To points_size)
    Else
        For i = 0 To points_size
            f.data(i) = 0
        Next
    End If
    f.data_size = points_size + 1
End Sub

' Erases the point from query results using the filter.
Public Sub PointTreeFilter_erase(f As PointTreeFilter, ByVal index As Long)
    If f.data(index) = 0 Then f.data(index) = 1     ' std::max(data[index], 1u)
End Sub

'--------------------------------------------------------------------------

' (VB) the expand() of each byte
Private Sub initExpandTable()
    Dim b As Long, r As Long, k As Long
    For b = 0 To 255
        r = 0
        For k = 0 To 7
            If b And (2 ^ k) Then r = r Or (4 ^ k)
        Next
        expandTable(b) = r
    Next
    expandReady = True
End Sub

' (VB) the 32 bits expand() makes of 16 (v), shifted left by shift (0 or 1)
Private Function expand16(ByVal v As Long, ByVal shift As Long) As Long
    Dim lo As Long, hi As Long
    lo = expandTable(v And &HFF&) * (1 + shift)
    hi = expandTable((v \ &H100&) And &HFF&) * (1 + shift)
    ' hi << 16
    If hi And &H8000& Then
        expand16 = ((hi And &H7FFF&) * &H10000) Or &H80000000 Or lo
    Else
        expand16 = (hi * &H10000) Or lo
    End If
End Function

' Expands bit pattern abcd efgh to 0a0b 0c0d 0e0f 0g0h; (VB) shifted left by shift (0 or 1)
Private Sub expand(ByVal x As Long, ByVal shift As Long, r As U64)
    If Not expandReady Then initExpandTable
    r.lo = expand16(x And &HFFFF&, shift)
    r.hi = expand16(SHRU(x, 16), shift)
End Sub

' Returns v with highest set bit and all higher bits set, and all following bits 0. Example: 0000 0110 1001 1100 -> 1111 1100 0000 0000.
Private Function findSplit(ByVal v As Long) As Long
    v = v Or SHRU(v, 1)
    v = v Or SHRU(v, 2)
    v = v Or SHRU(v, 4)
    v = v Or SHRU(v, 8)
    v = v Or SHRU(v, 16)
    findSplit = Not SHRU(v, 1)
End Function

Private Sub expandX(ByVal x As Long, r As U64)
    expand x Xor &H80000000, 1, r                  ' expand(x + 0x80000000u) << 1
End Sub

Private Sub expandY(ByVal y As Long, r As U64)
    expand y Xor &H80000000, 0, r                  ' expand(y + 0x80000000u)
End Sub

' Interleaves x and y, but after adding 0x80000000u to both, to make their ranges unsigned.
Private Sub interleave(ByVal x As Long, ByVal y As Long, r As U64)
    Dim ex As U64, ey As U64
    expandX x, ex
    expandY y, ey
    r.hi = ex.hi Or ey.hi
    r.lo = ex.lo Or ey.lo
End Sub

' Inserts a point into the point tree.
Public Sub PointTree_insert(ByVal pointData As Long, ByVal x As Long, ByVal y As Long)
    Dim k As U64
    If points_size = points_capacity Then
        If points_capacity = 0 Then points_capacity = 1024 Else points_capacity = points_capacity * 2
        ReDim Preserve pointsHi(0 To points_capacity - 1)
        ReDim Preserve pointsLo(0 To points_capacity - 1)
        ReDim Preserve pointsData(0 To points_capacity - 1)
        ReDim Preserve pointsX(0 To points_capacity - 1)
        ReDim Preserve pointsY(0 To points_capacity - 1)
    End If
    interleave x, y, k
    pointsHi(points_size) = k.hi Xor &H80000000
    pointsLo(points_size) = k.lo Xor &H80000000
    pointsData(points_size) = pointData
    pointsX(points_size) = x
    pointsY(points_size) = y
    points_size = points_size + 1
End Sub

' Clears the PointTree.
Public Sub PointTree_clear()
    points_size = 0
End Sub

' Must be done between inserting and querying, to get meaningful results.
' std::stable_sort by position only (pointTreeSortFunction), not by pointer
' address, even if two units are in the same place: (VB) a merge sort.
Public Sub PointTree_sort()
    If points_size < 2 Then Exit Sub
    ReDim sortTmpHi(0 To points_size - 1)
    ReDim sortTmpLo(0 To points_size - 1)
    ReDim sortTmpData(0 To points_size - 1)
    ReDim sortTmpX(0 To points_size - 1)
    ReDim sortTmpY(0 To points_size - 1)
    mergeSort 0, points_size
End Sub

' (VB) sorts points [first, last) stably
Private Sub mergeSort(ByVal first As Long, ByVal last As Long)
    Dim mid_ As Long, i As Long, j As Long, k As Long, takeRight As Boolean
    If last - first < 2 Then Exit Sub
    mid_ = (first + last) \ 2
    mergeSort first, mid_
    mergeSort mid_, last
    ' already in order?
    If Not keyLess(mid_, mid_ - 1) Then Exit Sub
    i = first
    j = mid_
    k = first
    Do While i < mid_ Or j < last
        If i >= mid_ Then
            takeRight = True
        ElseIf j >= last Then
            takeRight = False
        Else
            takeRight = keyLess(j, i)               ' only a strictly smaller right element goes first
        End If
        If takeRight Then
            sortTmpHi(k) = pointsHi(j): sortTmpLo(k) = pointsLo(j): sortTmpData(k) = pointsData(j)
            sortTmpX(k) = pointsX(j): sortTmpY(k) = pointsY(j)
            j = j + 1
        Else
            sortTmpHi(k) = pointsHi(i): sortTmpLo(k) = pointsLo(i): sortTmpData(k) = pointsData(i)
            sortTmpX(k) = pointsX(i): sortTmpY(k) = pointsY(i)
            i = i + 1
        End If
        k = k + 1
    Loop
    For k = first To last - 1
        pointsHi(k) = sortTmpHi(k): pointsLo(k) = sortTmpLo(k): pointsData(k) = sortTmpData(k)
        pointsX(k) = sortTmpX(k): pointsY(k) = sortTmpY(k)
    Next
End Sub

' (VB) points[a].first < points[b].first
Private Function keyLess(ByVal a As Long, ByVal b As Long) As Boolean
    If pointsHi(a) <> pointsHi(b) Then
        keyLess = pointsHi(a) < pointsHi(b)
    Else
        keyLess = pointsLo(a) < pointsLo(b)
    End If
End Function

' (VB) unsigned a < b
Private Function u64Less(a As U64, b As U64) As Boolean
    If a.hi <> b.hi Then
        u64Less = (a.hi Xor &H80000000) < (b.hi Xor &H80000000)
    Else
        u64Less = (a.lo Xor &H80000000) < (b.lo Xor &H80000000)
    End If
End Function

' (VB) a | b
Private Sub u64Or(a As U64, b As U64, r As U64)
    r.hi = a.hi Or b.hi
    r.lo = a.lo Or b.lo
End Sub

' std::lower_bound(points.begin() + first, points.end(), Point(key, nullptr), pointTreeSortFunction): the first point not less than key
Private Function lowerBound(ByVal first As Long, key As U64) As Long
    Dim hi As Long, lo As Long, count As Long, stp As Long, it As Long
    hi = key.hi Xor &H80000000
    lo = key.lo Xor &H80000000
    count = points_size - first
    Do While count > 0
        stp = count \ 2
        it = first + stp
        If pointsHi(it) < hi Or (pointsHi(it) = hi And pointsLo(it) < lo) Then
            first = it + 1
            count = count - stp - 1
        Else
            count = stp
        End If
    Loop
    lowerBound = first
End Function

' std::upper_bound(points.begin() + first, points.end(), Point(key, nullptr), pointTreeSortFunction): the first point greater than key
Private Function upperBound(ByVal first As Long, key As U64) As Long
    Dim hi As Long, lo As Long, count As Long, stp As Long, it As Long
    hi = key.hi Xor &H80000000
    lo = key.lo Xor &H80000000
    count = points_size - first
    Do While count > 0
        stp = count \ 2
        it = first + stp
        If Not (hi < pointsHi(it) Or (hi = pointsHi(it) And lo < pointsLo(it))) Then
            first = it + 1
            count = count - stp - 1
        Else
            count = stp
        End If
    Loop
    upperBound = first
End Function

' current<IsFiltered>(filterData, i): with no filter (IsFiltered false) it is i
Private Function current(f As PointTreeFilter, ByVal i As Long) As Long
    Dim ret As Long, nxt As Long
    ret = i
    Do While f.data(ret) <> 0
        ret = ret + f.data(ret)
    Loop
    Do While f.data(i) <> 0
        nxt = i + f.data(i)
        f.data(i) = ret - i
        i = nxt
    Loop

    current = ret
End Function

' queryMaybeFilter<IsFiltered>(filter, minXo, minYo, maxXo, maxYo): the
' results are lastQueryResults(0 .. lastQueryResults_size - 1)
Private Sub queryMaybeFilter(ByVal IsFiltered As Boolean, filter As PointTreeFilter, ByVal minXo As Long, ByVal minYo As Long, ByVal maxXo As Long, ByVal maxYo As Long)
    Dim minX As U64, maxX As U64, minY As U64, maxY As U64
    Dim splitXo As Long, splitYo As Long
    Dim splitX1 As U64, splitX2 As U64, splitY1 As U64, splitY2 As U64
    Dim ranges(0 To 3) As PointTreeRange, numRanges As Long, tmp As PointTreeRange, zp1 As U64
    Dim r As Long, i1 As Long, i2 As Long, i As Long, x As Long, y As Long

    expandX minXo, minX
    expandX maxXo, maxX
    expandY minYo, minY
    expandY maxYo, maxY

    splitXo = maxXo And findSplit(minXo Xor maxXo)
    splitYo = maxYo And findSplit(minYo Xor maxYo)

    expandX FromU32(CDbl(splitXo) - 1), splitX1    ' splitXo - 1 (uint32_t)
    expandX splitXo, splitX2
    expandY FromU32(CDbl(splitYo) - 1), splitY1
    expandY splitYo, splitY2

    u64Or minX, minY, ranges(0).a
    u64Or splitX1, splitY1, ranges(0).z
    u64Or splitX2, minY, ranges(1).a
    u64Or maxX, splitY1, ranges(1).z
    u64Or minX, splitY2, ranges(2).a
    u64Or splitX1, maxY, ranges(2).z
    u64Or splitX2, splitY2, ranges(3).a
    u64Or maxX, maxY, ranges(3).z
    numRanges = 4

    ' Sort ranges ready to be merged.
    If u64Less(ranges(2).a, ranges(1).a) Then
        tmp = ranges(1)
        ranges(1) = ranges(2)
        ranges(2) = tmp
    End If
    ' Merge ranges if needed.
    u64Inc ranges(2).z, zp1
    If Not u64Less(zp1, ranges(3).a) Then
        ranges(2).z = ranges(3).z
        numRanges = numRanges - 1
    End If
    u64Inc ranges(1).z, zp1
    If Not u64Less(zp1, ranges(2).a) Then
        ranges(1).z = ranges(2).z
        ranges(2) = ranges(3)
        numRanges = numRanges - 1
    End If
    u64Inc ranges(0).z, zp1
    If Not u64Less(zp1, ranges(1).a) Then
        ranges(0).z = ranges(1).z
        ranges(1) = ranges(2)
        ranges(2) = ranges(3)
        numRanges = numRanges - 1
    End If

    lastQueryResults_size = 0
    If IsFiltered Then
        lastFilteredQueryIndices_size = 0
    End If
    For r = 0 To numRanges - 1
        ' Find range of points which may be close enough. Range is [i1 ... i2 - 1]. The pointers are ignored when searching.
        i1 = lowerBound(0, ranges(r).a)
        i2 = upperBound(i1, ranges(r).z)

        If IsFiltered Then i = current(filter, i1) Else i = i1
        Do While i < i2
            x = pointsX(i)
            y = pointsY(i)
            If x >= minXo And x <= maxXo And y >= minYo And y <= maxYo Then    ' Only add point if it's at least in the desired square.
                pushResult pointsData(i)
                If IsFiltered Then
                    pushFilteredIndex i
                End If
            End If
            If IsFiltered Then i = current(filter, i + 1) Else i = i + 1
        Loop
    Next
End Sub

' (VB) r = a + 1 (wrapping)
Private Sub u64Inc(a As U64, r As U64)
    If a.lo = -1 Then
        r.lo = 0
        If a.hi = -1 Then r.hi = 0 Else r.hi = a.hi + 1
    Else
        r.lo = a.lo + 1
        r.hi = a.hi
    End If
End Sub

' (VB) lastQueryResults.push_back(p)
Private Sub pushResult(ByVal p As Long)
    If lastQueryResults_size = lastQueryResults_capacity Then
        If lastQueryResults_capacity = 0 Then lastQueryResults_capacity = 256 Else lastQueryResults_capacity = 2 * lastQueryResults_capacity
        ReDim Preserve lastQueryResults(0 To lastQueryResults_capacity - 1)
    End If
    lastQueryResults(lastQueryResults_size) = p
    lastQueryResults_size = lastQueryResults_size + 1
End Sub

' (VB) lastFilteredQueryIndices.push_back(i)
Private Sub pushFilteredIndex(ByVal i As Long)
    If lastFilteredQueryIndices_size = lastFilteredQueryIndices_capacity Then
        If lastFilteredQueryIndices_capacity = 0 Then lastFilteredQueryIndices_capacity = 256 Else lastFilteredQueryIndices_capacity = 2 * lastFilteredQueryIndices_capacity
        ReDim Preserve lastFilteredQueryIndices(0 To lastFilteredQueryIndices_capacity - 1)
    End If
    lastFilteredQueryIndices(lastFilteredQueryIndices_size) = i
    lastFilteredQueryIndices_size = lastFilteredQueryIndices_size + 1
End Sub

' Returns all points which have not been filtered away within given rectangle.
Public Sub PointTree_queryArea(ByVal x As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long)
    Dim unused As PointTreeFilter
    queryMaybeFilter False, unused, x, y, x2, y2
End Sub

' Returns all points less than or equal to radius from (x, y), possibly plus some extra nearby points.
' (More specifically, returns all objects in a square with edge length 2*radius.)
Public Sub PointTree_query(ByVal x As Long, ByVal y As Long, ByVal radius As Long)
    Dim unused As PointTreeFilter, minXo As Long, maxXo As Long, minYo As Long, maxYo As Long
    minXo = FromU32(CDbl(x) - U32(radius))
    maxXo = FromU32(CDbl(x) + U32(radius))
    minYo = FromU32(CDbl(y) - U32(radius))
    maxYo = FromU32(CDbl(y) + U32(radius))
    queryMaybeFilter False, unused, minXo, minYo, maxXo, maxYo
End Sub

' Returns all points which have not been filtered away, less than or equal to radius from (x, y), possibly plus some extra nearby points.
' (More specifically, returns objects in a square with edge length 2*radius.)
Public Sub PointTree_queryFiltered(filter As PointTreeFilter, ByVal x As Long, ByVal y As Long, ByVal radius As Long)
    Dim minXo As Long, maxXo As Long, minYo As Long, maxYo As Long
    minXo = FromU32(CDbl(x) - U32(radius))
    maxXo = FromU32(CDbl(x) + U32(radius))
    minYo = FromU32(CDbl(y) - U32(radius))
    maxYo = FromU32(CDbl(y) + U32(radius))
    queryMaybeFilter True, filter, minXo, minYo, maxXo, maxYo
End Sub
