Attribute VB_Name = "zstd_xxhash_c"
' xxHash's XXH64 as zstd 1.4.9 carries it (common/xxhash.c in ref/basis_universal/zstd/zstddeclib.c, namespaced
' ZSTD_XXH64_*): the streaming state the DCtx keeps for a frame's content checksum (XXH64_reset, XXH64_update,
' XXH64_digest). A U64 is two Longs (its 32-bit halves); a product modulo 2^64 is built from 16-bit limbs in
' Doubles (exact below 2^53). The DCtx's state (dctx->xxhState) is this module's.
Option Explicit

Private Type U64
    lo As Long
    hi As Long
End Type

' XXH64_state_t
Private Type XXH64_state_t
    total_len As U64
    v1 As U64
    v2 As U64
    v3 As U64
    v4 As U64
    mem64(0 To 31) As Byte
    memsize As Long
End Type

Private PRIME64_1 As U64
Private PRIME64_2 As U64
Private PRIME64_3 As U64
Private PRIME64_4 As U64
Private PRIME64_5 As U64
Private pow2(0 To 31) As Long
Private initialised As Boolean

Private xxhState As XXH64_state_t      ' dctx->xxhState

Private Sub initTables()
    Dim i As Long
    pow2(0) = 1
    For i = 1 To 30
        pow2(i) = pow2(i - 1) * 2
    Next i
    pow2(31) = &H80000000
    PRIME64_1.hi = &H9E3779B1: PRIME64_1.lo = &H85EBCA87    ' 11400714785074694791ULL
    PRIME64_2.hi = &HC2B2AE3D: PRIME64_2.lo = &H27D4EB4F    ' 14029467366897019727ULL
    PRIME64_3.hi = &H165667B1: PRIME64_3.lo = &H9E3779F9    '  1609587929392839161ULL
    PRIME64_4.hi = &H85EBCA77: PRIME64_4.lo = &HC2B2AE63    '  9650029242287828579ULL
    PRIME64_5.hi = &H27D4EB2F: PRIME64_5.lo = &H165667C5    '  2870177450012600261ULL
    initialised = True
End Sub

'--------------------------------------------------------------------------
' (VB) 64-bit arithmetic modulo 2^64

Private Function zSHRU(ByVal x As Long, ByVal s As Long) As Long
    If s = 0 Then zSHRU = x: Exit Function
    If x >= 0 Then
        zSHRU = x \ pow2(s)
    Else
        zSHRU = ((x And &H7FFFFFFF) \ pow2(s)) Or pow2(31 - s)
    End If
End Function

Private Function zULT(ByVal a As Long, ByVal b As Long) As Boolean
    zULT = (a Xor &H80000000) < (b Xor &H80000000)
End Function

Private Function mk64(ByVal hi As Long, ByVal lo As Long) As U64
    mk64.hi = hi
    mk64.lo = lo
End Function

Private Function u64add(a As U64, b As U64) As U64
    Dim lo As Long
    lo = a.lo + b.lo
    u64add.lo = lo
    u64add.hi = a.hi + b.hi
    If zULT(lo, a.lo) Then u64add.hi = u64add.hi + 1
End Function

Private Function u64sub(a As U64, b As U64) As U64
    u64sub.lo = a.lo - b.lo
    u64sub.hi = a.hi - b.hi
    If zULT(a.lo, b.lo) Then u64sub.hi = u64sub.hi - 1
End Function

Private Function u64xor(a As U64, b As U64) As U64
    u64xor.lo = a.lo Xor b.lo
    u64xor.hi = a.hi Xor b.hi
End Function

' the 16-bit limbs of a 32-bit half
Private Function lowLimb(ByVal x As Long) As Double
    lowLimb = x And &HFFFF&
End Function

Private Function highLimb(ByVal x As Long) As Double
    highLimb = (x And &H7FFF0000) \ &H10000
    If x < 0 Then highLimb = highLimb + 32768#
End Function

' a 32-bit half from two 16-bit limbs (0 .. 65535 each)
Private Function fromLimbs(ByVal h As Long, ByVal l As Long) As Long
    fromLimbs = ((h And &H7FFF&) * &H10000) Or l
    If (h And &H8000&) <> 0 Then fromLimbs = fromLimbs Or &H80000000
End Function

Private Function u64mul(a As U64, b As U64) As U64
    Dim a0 As Double, a1 As Double, a2 As Double, a3 As Double
    Dim b0 As Double, b1 As Double, b2 As Double, b3 As Double
    Dim t As Double, c As Double, r0 As Long, r1 As Long, r2 As Long, r3 As Long
    a0 = lowLimb(a.lo): a1 = highLimb(a.lo): a2 = lowLimb(a.hi): a3 = highLimb(a.hi)
    b0 = lowLimb(b.lo): b1 = highLimb(b.lo): b2 = lowLimb(b.hi): b3 = highLimb(b.hi)
    t = a0 * b0
    c = Int(t / 65536#)
    r0 = t - c * 65536#
    t = a0 * b1 + a1 * b0 + c
    c = Int(t / 65536#)
    r1 = t - c * 65536#
    t = a0 * b2 + a1 * b1 + a2 * b0 + c
    c = Int(t / 65536#)
    r2 = t - c * 65536#
    t = a0 * b3 + a1 * b2 + a2 * b1 + a3 * b0 + c
    c = Int(t / 65536#)
    r3 = t - c * 65536#
    u64mul.lo = fromLimbs(r1, r0)
    u64mul.hi = fromLimbs(r3, r2)
End Function

' XXH_rotl64(x, r): (x << r) | (x >> (64 - r)), 0 < r < 64
Private Function u64rotl(x As U64, ByVal r As Long) As U64
    Dim h As Long, l As Long
    If r >= 32 Then
        h = x.lo: l = x.hi
        r = r - 32
    Else
        h = x.hi: l = x.lo
    End If
    If r = 0 Then
        u64rotl.hi = h: u64rotl.lo = l
    Else
        u64rotl.hi = (h * pow2(r)) Or zSHRU(l, 32 - r)
        u64rotl.lo = (l * pow2(r)) Or zSHRU(h, 32 - r)
    End If
End Function

' x >> s, 0 < s < 64
Private Function u64shr(x As U64, ByVal s As Long) As U64
    If s >= 32 Then
        u64shr.hi = 0
        u64shr.lo = zSHRU(x.hi, s - 32)
    Else
        u64shr.hi = zSHRU(x.hi, s)
        u64shr.lo = zSHRU(x.lo, s) Or (x.hi * pow2(32 - s))
    End If
End Function

Private Function readLE32(b() As Byte, ByVal p As Long) As Long
    readLE32 = b(p) Or (CLng(b(p + 1)) * &H100&) Or (CLng(b(p + 2)) * &H10000) Or (CLng(b(p + 3)) * &H1000000)
End Function

Private Function readLE64(b() As Byte, ByVal p As Long) As U64
    readLE64.lo = readLE32(b, p)
    readLE64.hi = readLE32(b, p + 4)
End Function

'--------------------------------------------------------------------------
' xxhash.c

Private Function XXH64_round(acc As U64, inp As U64) As U64
    Dim a As U64
    a = u64add(acc, u64mul(inp, PRIME64_2))
    a = u64rotl(a, 31)
    XXH64_round = u64mul(a, PRIME64_1)
End Function

Private Function XXH64_mergeRound(acc As U64, val_ As U64) As U64
    Dim v As U64, a As U64, zero As U64
    v = XXH64_round(zero, val_)
    a = u64xor(acc, v)
    XXH64_mergeRound = u64add(u64mul(a, PRIME64_1), PRIME64_4)
End Function

' XXH64_reset(&dctx->xxhState, seed) (seed: its two halves)
Public Sub XXH64_reset(ByVal seedHi As Long, ByVal seedLo As Long)
    Dim seed As U64, i As Long
    If Not initialised Then initTables
    seed = mk64(seedHi, seedLo)
    xxhState.total_len = mk64(0, 0)
    For i = 0 To 31
        xxhState.mem64(i) = 0
    Next i
    xxhState.memsize = 0
    xxhState.v1 = u64add(u64add(seed, PRIME64_1), PRIME64_2)
    xxhState.v2 = u64add(seed, PRIME64_2)
    xxhState.v3 = seed
    xxhState.v4 = u64sub(seed, PRIME64_1)
End Sub

' XXH64_update(&dctx->xxhState, input, len): input is b(p .. p + length - 1)
Public Sub XXH64_update(b() As Byte, ByVal p As Long, ByVal length As Long)
    Dim bEnd As Long, limit As Long, i As Long, n As Long, v1 As U64, v2 As U64, v3 As U64, v4 As U64
    If Not initialised Then initTables
    bEnd = p + length

    xxhState.total_len = u64add(xxhState.total_len, mk64(0, length))

    If xxhState.memsize + length < 32 Then  ' fill in tmp buffer
        For i = 0 To length - 1
            xxhState.mem64(xxhState.memsize + i) = b(p + i)
        Next i
        xxhState.memsize = xxhState.memsize + length
        Exit Sub
    End If

    If xxhState.memsize <> 0 Then   ' tmp buffer is full
        n = 32 - xxhState.memsize
        For i = 0 To n - 1
            xxhState.mem64(xxhState.memsize + i) = b(p + i)
        Next i
        xxhState.v1 = XXH64_round(xxhState.v1, readLE64(xxhState.mem64, 0))
        xxhState.v2 = XXH64_round(xxhState.v2, readLE64(xxhState.mem64, 8))
        xxhState.v3 = XXH64_round(xxhState.v3, readLE64(xxhState.mem64, 16))
        xxhState.v4 = XXH64_round(xxhState.v4, readLE64(xxhState.mem64, 24))
        p = p + n
        xxhState.memsize = 0
    End If

    If p + 32 <= bEnd Then
        limit = bEnd - 32
        v1 = xxhState.v1
        v2 = xxhState.v2
        v3 = xxhState.v3
        v4 = xxhState.v4

        Do
            v1 = XXH64_round(v1, readLE64(b, p)): p = p + 8
            v2 = XXH64_round(v2, readLE64(b, p)): p = p + 8
            v3 = XXH64_round(v3, readLE64(b, p)): p = p + 8
            v4 = XXH64_round(v4, readLE64(b, p)): p = p + 8
        Loop While p <= limit

        xxhState.v1 = v1
        xxhState.v2 = v2
        xxhState.v3 = v3
        xxhState.v4 = v4
    End If

    If p < bEnd Then
        For i = 0 To bEnd - p - 1
            xxhState.mem64(i) = b(p + i)
        Next i
        xxhState.memsize = bEnd - p
    End If
End Sub

' (U32)XXH64_digest(&dctx->xxhState): the low half of the digest
Public Function XXH64_digestLow32() As Long
    Dim p As Long, bEnd As Long, h64 As U64, k1 As U64, zero As U64
    If Not initialised Then initTables
    p = 0
    bEnd = xxhState.memsize

    If Not zULT(xxhState.total_len.lo, 32) Or xxhState.total_len.hi <> 0 Then
        h64 = u64add(u64add(u64add(u64rotl(xxhState.v1, 1), u64rotl(xxhState.v2, 7)), u64rotl(xxhState.v3, 12)), u64rotl(xxhState.v4, 18))
        h64 = XXH64_mergeRound(h64, xxhState.v1)
        h64 = XXH64_mergeRound(h64, xxhState.v2)
        h64 = XXH64_mergeRound(h64, xxhState.v3)
        h64 = XXH64_mergeRound(h64, xxhState.v4)
    Else
        h64 = u64add(xxhState.v3, PRIME64_5)
    End If

    h64 = u64add(h64, xxhState.total_len)

    Do While p + 8 <= bEnd
        k1 = XXH64_round(zero, readLE64(xxhState.mem64, p))
        h64 = u64xor(h64, k1)
        h64 = u64add(u64mul(u64rotl(h64, 27), PRIME64_1), PRIME64_4)
        p = p + 8
    Loop

    If p + 4 <= bEnd Then
        h64 = u64xor(h64, u64mul(mk64(0, readLE32(xxhState.mem64, p)), PRIME64_1))
        h64 = u64add(u64mul(u64rotl(h64, 23), PRIME64_2), PRIME64_3)
        p = p + 4
    End If

    Do While p < bEnd
        h64 = u64xor(h64, u64mul(mk64(0, xxhState.mem64(p)), PRIME64_5))
        h64 = u64mul(u64rotl(h64, 11), PRIME64_1)
        p = p + 1
    Loop

    h64 = u64xor(h64, u64shr(h64, 33))
    h64 = u64mul(h64, PRIME64_2)
    h64 = u64xor(h64, u64shr(h64, 29))
    h64 = u64mul(h64, PRIME64_3)
    h64 = u64xor(h64, u64shr(h64, 32))

    XXH64_digestLow32 = h64.lo
End Function
