Attribute VB_Name = "zstd_entropy_c"
' zstd 1.4.9's entropy decoders as ref/basis_universal/zstd/zstddeclib.c has them (the single-file decoder Basis
' Universal ships, which 4.7.0 builds): common/bitstream.h's decoding half, common/entropy_common.c
' (FSE_readNCount, HUF_readStats), common/fse_decompress.c (the FSE decoder HUF_readStats uses for the Huffman
' weights) and decompress/huf_decompress.c (the X1 and X2 Huffman decoders, one stream and four). The 32-bit
' build's code paths (MEM_32bits(): size_t, and so the bit container, is 32 bits; no BMI2: DYNAMIC_BMI2 is 0 on
' i386; HUF_TABLELOG_MAX 12). A size_t result is a Long: an error is -ZSTD_error_xxx (zstd_decompress_c's
' ZSTD_isError). A C pointer into a buffer is the Byte array plus an offset.
'
' The Huffman decoding table is the DCtx's (dctx->entropy.hufTable, the only one ZSTD_decompress uses): this
' module's arrays, the X1 entries and the X2 entries apart (C overlays them in one HUF_DTable; DTableDesc's
' tableType says which one is valid, as in C). The FSE table and the work spaces are local arrays; their size
' checks are kept where they can fail (FSE_decompress_wksp's against HUF_readStats' 356-byte work space).
' bitstream.h's functions are Private (zstd_decompress_block_c keeps its own copy for the sequences).
Option Explicit

' error_private.h's codes (zstd_errors.h's values, as zstd_decompress_c's Public Enum ZSTD_ErrorCode)
Private Const ERR_GENERIC As Long = -1
Private Const ERR_corruption_detected As Long = -20
Private Const ERR_tableLog_tooLarge As Long = -44
Private Const ERR_maxSymbolValue_tooLarge As Long = -46
Private Const ERR_maxSymbolValue_tooSmall As Long = -48
Private Const ERR_dstSize_tooSmall As Long = -70
Private Const ERR_srcSize_wrong As Long = -72
Private Const ERROR_maxCode As Long = 120

' fse.h / huf.h
Private Const FSE_MIN_TABLELOG As Long = 5
Private Const FSE_TABLELOG_ABSOLUTE_MAX As Long = 15
Private Const FSE_MAX_SYMBOL_VALUE As Long = 255
Private Const FSE_MAX_TABLELOG As Long = 12                 ' FSE_MAX_MEMORY_USAGE (14) - 2
Private Const HUF_TABLELOG_MAX As Long = 12
Private Const HUF_SYMBOLVALUE_MAX As Long = 255
Private Const HUF_TABLELOG_ABSOLUTEMAX As Long = 15
Private Const HUF_DECOMPRESS_WORKSPACE_SIZE As Long = 2048  ' (2 << 10): the DCtx's workspace, in bytes
' HUF_READ_STATS_WORKSPACE_SIZE: FSE_DECOMPRESS_WKSP_SIZE_U32(6, HUF_TABLELOG_MAX-1) * 4 = (65 + 24) * 4
Private Const HUF_READ_STATS_WORKSPACE_SIZE As Long = 356

' bitstream.h: BIT_DStream_status
Private Const BIT_DStream_unfinished As Long = 0
Private Const BIT_DStream_endOfBuffer As Long = 1
Private Const BIT_DStream_completed As Long = 2
Private Const BIT_DStream_overflow As Long = 3

' bitstream.h: BIT_DStream_t (32-bit build: the container is a 32-bit size_t). ptr, start and limitPtr are
' offsets into the source array the caller passes along.
Private Type BIT_DStream_t
    bitContainer As Long
    bitsConsumed As Long
    ptr As Long
    start As Long
    limitPtr As Long
End Type

' fse.h: FSE_DState_t (the table is this module's fseDT* arrays)
Private Type FSE_DState_t
    state As Long
End Type

' (VB) 1 << n, BIT_mask, and the highest set bit of a byte
Private pow2(0 To 31) As Long
Private BIT_mask(0 To 31) As Long
Private hibit8(0 To 255) As Long
Private initialised As Boolean

' huf_decompress.c: the DCtx's HUF_DTable (dctx->entropy.hufTable): its DTableDesc and its entries.
' X1 (HUF_DEltX1 {byte, nbBits}) and X2 (HUF_DEltX2 {sequence (LE16), nbBits, length}), 1 << 12 of each.
Private dtd_maxTableLog As Long
Private dtd_tableType As Long
Private dtd_tableLog As Long
Private dtd_reserved As Long
Private dtX1byte(0 To 4095) As Byte
Private dtX1nbBits(0 To 4095) As Long
Private dtX2seq(0 To 4095) As Long         ' the 2 bytes copied out: low byte first
Private dtX2nbBits(0 To 4095) As Long
Private dtX2length(0 To 4095) As Long

' fse_decompress.c: the FSE_DTable FSE_decompress_wksp builds in its work space (tableLog <= 6 for the weights;
' FSE_buildDTable_internal accepts up to FSE_MAX_TABLELOG)
Private fseDT_tableLog As Long
Private fseDT_fastMode As Long
Private fseDT_newState(0 To 4095) As Long
Private fseDT_symbol(0 To 4095) As Long
Private fseDT_nbBits(0 To 4095) As Long

' huf_decompress.c: algoTime[16][3] (tableTime, decode256Time); only the single and double columns are read
Private algoTime_table(0 To 15, 0 To 2) As Long
Private algoTime_decode(0 To 15, 0 To 2) As Long

Private Sub initTables()
    Dim i As Long, v As Long, r As Long
    pow2(0) = 1
    For i = 1 To 30
        pow2(i) = pow2(i - 1) * 2
    Next i
    pow2(31) = &H80000000
    BIT_mask(0) = 0
    For i = 1 To 31
        BIT_mask(i) = pow2(i) - 1
    Next i
    For v = 1 To 255
        r = 0
        i = v
        Do While i > 1
            i = i \ 2
            r = r + 1
        Loop
        hibit8(v) = r
    Next v
    setAlgoTime 0, 0, 0, 1, 1, 2, 2
    setAlgoTime 1, 0, 0, 1, 1, 2, 2
    setAlgoTime 2, 38, 130, 1313, 74, 2151, 38
    setAlgoTime 3, 448, 128, 1353, 74, 2238, 41
    setAlgoTime 4, 556, 128, 1353, 74, 2238, 47
    setAlgoTime 5, 714, 128, 1418, 74, 2436, 53
    setAlgoTime 6, 883, 128, 1437, 74, 2464, 61
    setAlgoTime 7, 897, 128, 1515, 75, 2622, 68
    setAlgoTime 8, 926, 128, 1613, 75, 2730, 75
    setAlgoTime 9, 947, 128, 1729, 77, 3359, 77
    setAlgoTime 10, 1107, 128, 2083, 81, 4006, 84
    setAlgoTime 11, 1177, 128, 2379, 87, 4785, 88
    setAlgoTime 12, 1242, 128, 2415, 93, 5155, 84
    setAlgoTime 13, 1349, 128, 2644, 106, 5260, 106
    setAlgoTime 14, 1455, 128, 2422, 124, 4174, 124
    setAlgoTime 15, 722, 128, 1891, 145, 1936, 146
    initialised = True
End Sub

Private Sub setAlgoTime(ByVal q As Long, ByVal t0 As Long, ByVal d0 As Long, ByVal t1 As Long, ByVal d1 As Long, ByVal t2 As Long, ByVal d2 As Long)
    algoTime_table(q, 0) = t0: algoTime_decode(q, 0) = d0
    algoTime_table(q, 1) = t1: algoTime_decode(q, 1) = d1
    algoTime_table(q, 2) = t2: algoTime_decode(q, 2) = d2
End Sub

'--------------------------------------------------------------------------
' (VB) C's unsigned arithmetic on 32-bit values

' a < b for size_t / U32
Private Function zULT(ByVal a As Long, ByVal b As Long) As Boolean
    zULT = (a Xor &H80000000) < (b Xor &H80000000)
End Function

' x >> s for a U32 (the count taken modulo 32, as x86's shr does)
Private Function zSHRU(ByVal x As Long, ByVal s As Long) As Long
    s = s And 31
    If x >= 0 Then
        zSHRU = x \ pow2(s)
    Else
        zSHRU = ((x And &H7FFFFFFF) \ pow2(s)) Or pow2(31 - s)
    End If
End Function

'--------------------------------------------------------------------------
' mem.h

Private Function MEM_readLE16(b() As Byte, ByVal p As Long) As Long
    MEM_readLE16 = b(p) Or (CLng(b(p + 1)) * &H100&)
End Function

Private Function MEM_readLE32(b() As Byte, ByVal p As Long) As Long
    MEM_readLE32 = b(p) Or (CLng(b(p + 1)) * &H100&) Or (CLng(b(p + 2)) * &H10000) Or (CLng(b(p + 3)) * &H1000000)
End Function

'--------------------------------------------------------------------------
' bitstream.h (decoding)

' BIT_highbit32: the index of the highest set bit (val != 0)
Private Function BIT_highbit32(ByVal v As Long) As Long
    Dim r As Long
    If v < 0 Then BIT_highbit32 = 31: Exit Function
    If v >= &H10000 Then v = v \ &H10000: r = 16
    If v >= &H100& Then v = v \ &H100&: r = r + 8
    BIT_highbit32 = r + hibit8(v)
End Function

Private Function BIT_initDStream(bitD As BIT_DStream_t, src() As Byte, ByVal srcBuffer As Long, ByVal srcSize As Long) As Long
    Dim lastByte As Long
    If srcSize < 1 Then
        bitD.bitContainer = 0: bitD.bitsConsumed = 0: bitD.ptr = 0: bitD.start = 0: bitD.limitPtr = 0
        BIT_initDStream = ERR_srcSize_wrong
        Exit Function
    End If

    bitD.start = srcBuffer
    bitD.limitPtr = bitD.start + 4
    If srcSize >= 4 Then    ' normal case
        bitD.ptr = srcBuffer + srcSize - 4
        bitD.bitContainer = MEM_readLE32(src, bitD.ptr)
        lastByte = src(srcBuffer + srcSize - 1)
        If lastByte <> 0 Then bitD.bitsConsumed = 8 - BIT_highbit32(lastByte) Else bitD.bitsConsumed = 0 ' ensures bitsConsumed is always set
        If lastByte = 0 Then BIT_initDStream = ERR_GENERIC: Exit Function ' endMark not present
    Else
        bitD.ptr = bitD.start
        bitD.bitContainer = src(bitD.start)
        Select Case srcSize     ' (32-bit: srcSize is 1 to 3; the 5 to 7 cases are the 64-bit build's)
        Case 3
            bitD.bitContainer = bitD.bitContainer + CLng(src(srcBuffer + 2)) * &H10000
            bitD.bitContainer = bitD.bitContainer + CLng(src(srcBuffer + 1)) * &H100&
        Case 2
            bitD.bitContainer = bitD.bitContainer + CLng(src(srcBuffer + 1)) * &H100&
        End Select
        lastByte = src(srcBuffer + srcSize - 1)
        If lastByte <> 0 Then bitD.bitsConsumed = 8 - BIT_highbit32(lastByte) Else bitD.bitsConsumed = 0
        If lastByte = 0 Then BIT_initDStream = ERR_corruption_detected: Exit Function ' endMark not present
        bitD.bitsConsumed = bitD.bitsConsumed + (4 - srcSize) * 8
    End If
    BIT_initDStream = srcSize
End Function

' BIT_lookBits: BIT_getMiddleBits(bitContainer, 32 - bitsConsumed - nbBits, nbBits) (start & regMask)
Private Function BIT_lookBits(bitD As BIT_DStream_t, ByVal nbBits As Long) As Long
    BIT_lookBits = zSHRU(bitD.bitContainer, (32 - bitD.bitsConsumed - nbBits) And 31) And BIT_mask(nbBits)
End Function

' BIT_lookBitsFast: unsafe version; only works if nbBits >= 1:
' (bitContainer << (bitsConsumed & regMask)) >> (((regMask+1)-nbBits) & regMask). The Huffman decoders' loops write
' it out with the shift's divisor (dv = 1 << s) and top bit (tb = 1 << (31 - s)) computed once (s = 32 - dtLog).
Private Function BIT_lookBitsFast(bitD As BIT_DStream_t, ByVal nbBits As Long) As Long
    Dim x As Long, s As Long
    x = bitD.bitContainer * pow2(bitD.bitsConsumed And 31)
    s = (32 - nbBits) And 31
    If x >= 0 Then
        BIT_lookBitsFast = x \ pow2(s)
    Else
        BIT_lookBitsFast = ((x And &H7FFFFFFF) \ pow2(s)) Or pow2(31 - s)
    End If
End Function

Private Function BIT_readBits(bitD As BIT_DStream_t, ByVal nbBits As Long) As Long
    BIT_readBits = BIT_lookBits(bitD, nbBits)
    bitD.bitsConsumed = bitD.bitsConsumed + nbBits     ' BIT_skipBits
End Function

Private Function BIT_readBitsFast(bitD As BIT_DStream_t, ByVal nbBits As Long) As Long
    BIT_readBitsFast = BIT_lookBitsFast(bitD, nbBits)
    bitD.bitsConsumed = bitD.bitsConsumed + nbBits
End Function

Private Function BIT_reloadDStreamFast(bitD As BIT_DStream_t, src() As Byte) As Long
    If bitD.ptr < bitD.limitPtr Then BIT_reloadDStreamFast = BIT_DStream_overflow: Exit Function
    bitD.ptr = bitD.ptr - bitD.bitsConsumed \ 8
    bitD.bitsConsumed = bitD.bitsConsumed And 7
    bitD.bitContainer = MEM_readLE32(src, bitD.ptr)
    BIT_reloadDStreamFast = BIT_DStream_unfinished
End Function

Private Function BIT_reloadDStream(bitD As BIT_DStream_t, src() As Byte) As Long
    Dim nbBytes As Long, result As Long
    If bitD.bitsConsumed > 32 Then  ' overflow detected, like end of stream
        BIT_reloadDStream = BIT_DStream_overflow
        Exit Function
    End If
    If bitD.ptr >= bitD.limitPtr Then
        BIT_reloadDStream = BIT_reloadDStreamFast(bitD, src)
        Exit Function
    End If
    If bitD.ptr = bitD.start Then
        If bitD.bitsConsumed < 32 Then BIT_reloadDStream = BIT_DStream_endOfBuffer: Exit Function
        BIT_reloadDStream = BIT_DStream_completed
        Exit Function
    End If
    ' start < ptr < limitPtr
    nbBytes = bitD.bitsConsumed \ 8
    result = BIT_DStream_unfinished
    If bitD.ptr - nbBytes < bitD.start Then
        nbBytes = bitD.ptr - bitD.start ' ptr > start
        result = BIT_DStream_endOfBuffer
    End If
    bitD.ptr = bitD.ptr - nbBytes
    bitD.bitsConsumed = bitD.bitsConsumed - nbBytes * 8
    bitD.bitContainer = MEM_readLE32(src, bitD.ptr) ' reminder : srcSize > sizeof(bitD->bitContainer), otherwise bitD->ptr == bitD->start
    BIT_reloadDStream = result
End Function

' BIT_endOfDStream: 1 if DStream has _exactly_ reached its end (all bits consumed)
Private Function BIT_endOfDStream(bitD As BIT_DStream_t) As Boolean
    BIT_endOfDStream = (bitD.ptr = bitD.start) And (bitD.bitsConsumed = 32)
End Function

'--------------------------------------------------------------------------
' entropy_common.c

' FSE_ctz (val != 0)
Private Function FSE_ctz(ByVal v As Long) As Long
    FSE_ctz = BIT_highbit32(v And -v)
End Function

' size_t FSE_readNCount(short* normalizedCounter, unsigned* maxSVPtr, unsigned* tableLogPtr, const void* headerBuffer,
' size_t hbSize): normalizedCounter(0 .. *maxSVPtr); the header is headerBuffer(hb .. hb + hbSize - 1)
Public Function FSE_readNCount(normalizedCounter() As Long, maxSVPtr As Long, tableLogPtr As Long, headerBuffer() As Byte, ByVal hb As Long, ByVal hbSize As Long) As Long
    Dim istart As Long, iend As Long, ip As Long
    Dim nbBits As Long, remaining As Long, threshold As Long, bitStream As Long, bitCount As Long
    Dim charnum As Long, maxSV1 As Long, previous0 As Boolean, s As Long
    Dim buffer(0 To 7) As Byte, countSize As Long, repeats As Long, max_ As Long, count As Long

    If Not initialised Then initTables
    istart = hb
    iend = istart + hbSize
    ip = istart
    charnum = 0
    maxSV1 = maxSVPtr + 1
    previous0 = False

    If hbSize < 8 Then
        ' This function only works when hbSize >= 8
        For s = 0 To hbSize - 1
            buffer(s) = headerBuffer(hb + s)
        Next s
        countSize = FSE_readNCount(normalizedCounter, maxSVPtr, tableLogPtr, buffer, 0, 8)
        If zstdIsError(countSize) Then FSE_readNCount = countSize: Exit Function
        If countSize > hbSize Then FSE_readNCount = ERR_corruption_detected: Exit Function
        FSE_readNCount = countSize
        Exit Function
    End If

    ' init
    For s = 0 To maxSVPtr       ' all symbols not present in NCount have a frequency of 0
        normalizedCounter(s) = 0
    Next s
    bitStream = MEM_readLE32(headerBuffer, ip)
    nbBits = (bitStream And &HF) + FSE_MIN_TABLELOG ' extract tableLog
    If nbBits > FSE_TABLELOG_ABSOLUTE_MAX Then FSE_readNCount = ERR_tableLog_tooLarge: Exit Function
    bitStream = zSHRU(bitStream, 4)
    bitCount = 4
    tableLogPtr = nbBits
    remaining = pow2(nbBits) + 1
    threshold = pow2(nbBits)
    nbBits = nbBits + 1

    Do
        If previous0 Then
            ' Count the number of repeats. Each time the possible repeat code is 0b11 there is another repeat.
            repeats = FSE_ctz((Not bitStream) Or &H80000000) \ 2
            Do While repeats >= 12
                charnum = charnum + 3 * 12
                If ip <= iend - 7 Then
                    ip = ip + 3
                Else
                    bitCount = bitCount - 8 * (iend - 7 - ip)
                    bitCount = bitCount And 31
                    ip = iend - 4
                End If
                bitStream = zSHRU(MEM_readLE32(headerBuffer, ip), bitCount)
                repeats = FSE_ctz((Not bitStream) Or &H80000000) \ 2
            Loop
            charnum = charnum + 3 * repeats
            bitStream = zSHRU(bitStream, 2 * repeats)
            bitCount = bitCount + 2 * repeats

            ' Add the final repeat which isn't 0b11.
            charnum = charnum + (bitStream And 3)
            bitCount = bitCount + 2

            ' This is an error, but break and return an error at the end, because returning out of a loop makes
            ' it harder for the compiler to optimize.
            If charnum >= maxSV1 Then Exit Do

            ' We don't need to set the normalized count to 0 because we already memset the whole buffer to 0.
            If (ip <= iend - 7) Or (ip + bitCount \ 8 <= iend - 4) Then
                ip = ip + bitCount \ 8
                bitCount = bitCount And 7
            Else
                bitCount = bitCount - 8 * (iend - 4 - ip)
                bitCount = bitCount And 31
                ip = iend - 4
            End If
            bitStream = zSHRU(MEM_readLE32(headerBuffer, ip), bitCount)
        End If
        max_ = (2 * threshold - 1) - remaining
        If zULT(bitStream And (threshold - 1), max_) Then
            count = bitStream And (threshold - 1)
            bitCount = bitCount + nbBits - 1
        Else
            count = bitStream And (2 * threshold - 1)
            If count >= threshold Then count = count - max_
            bitCount = bitCount + nbBits
        End If

        count = count - 1   ' extra accuracy
        ' When it matters (small blocks), this is a predictable branch, because we don't use -1.
        If count >= 0 Then
            remaining = remaining - count
        Else
            remaining = remaining + count
        End If
        normalizedCounter(charnum) = count
        charnum = charnum + 1
        previous0 = (count = 0)

        If remaining < threshold Then
            ' This branch can be folded into the threshold update condition because we know that threshold > 1.
            If remaining <= 1 Then Exit Do
            nbBits = BIT_highbit32(remaining) + 1
            threshold = pow2(nbBits - 1)
        End If
        If charnum >= maxSV1 Then Exit Do

        If (ip <= iend - 7) Or (ip + bitCount \ 8 <= iend - 4) Then
            ip = ip + bitCount \ 8
            bitCount = bitCount And 7
        Else
            bitCount = bitCount - 8 * (iend - 4 - ip)
            bitCount = bitCount And 31
            ip = iend - 4
        End If
        bitStream = zSHRU(MEM_readLE32(headerBuffer, ip), bitCount)
    Loop
    If remaining <> 1 Then FSE_readNCount = ERR_corruption_detected: Exit Function
    ' Only possible when there are too many zeros.
    If charnum > maxSV1 Then FSE_readNCount = ERR_maxSymbolValue_tooSmall: Exit Function
    If bitCount > 32 Then FSE_readNCount = ERR_corruption_detected: Exit Function
    maxSVPtr = charnum - 1

    ip = ip + (bitCount + 7) \ 8
    FSE_readNCount = ip - istart
End Function

' (VB) ERR_isError (error_private.h): code > (size_t)-ZSTD_error_maxCode
Private Function zstdIsError(ByVal code As Long) As Boolean
    zstdIsError = (code < 0) And (code > -ERROR_maxCode)
End Function

' HUF_readStats_body: huffWeight(0 .. hwSize - 1), rankStats(0 .. HUF_TABLELOG_MAX); the source is
' src(ip0 .. ip0 + srcSize - 1); wkspSize is the work space HUF_readStats_wksp would get (its size check)
Private Function HUF_readStats_body(huffWeight() As Byte, ByVal hwSize As Long, rankStats() As Long, nbSymbolsPtr As Long, tableLogPtr As Long, src() As Byte, ByVal ip0 As Long, ByVal srcSize As Long, ByVal wkspSize As Long) As Long
    Dim weightTotal As Long, ip As Long, iSize As Long, oSize As Long, n As Long
    Dim tableLog As Long, total As Long, rest As Long, verif As Long, lastWeight As Long

    ip = ip0
    If srcSize = 0 Then HUF_readStats_body = ERR_srcSize_wrong: Exit Function
    iSize = src(ip)
    ' ZSTD_memset(huffWeight, 0, hwSize);   (is not necessary, even though some analyzer complain ...)

    If iSize >= 128 Then    ' special header
        oSize = iSize - 127
        iSize = (oSize + 1) \ 2
        If iSize + 1 > srcSize Then HUF_readStats_body = ERR_srcSize_wrong: Exit Function
        If oSize >= hwSize Then HUF_readStats_body = ERR_corruption_detected: Exit Function
        ip = ip + 1
        For n = 0 To oSize - 1 Step 2
            huffWeight(n) = src(ip + n \ 2) \ 16
            huffWeight(n + 1) = src(ip + n \ 2) And 15
        Next n
    Else    ' header compressed with FSE (normal case)
        If iSize + 1 > srcSize Then HUF_readStats_body = ERR_srcSize_wrong: Exit Function
        ' max (hwSize-1) values decoded, as last one is implied
        oSize = FSE_decompress_wksp_body(huffWeight, hwSize - 1, src, ip + 1, iSize, 6, wkspSize)
        If zstdIsError(oSize) Then HUF_readStats_body = oSize: Exit Function
    End If

    ' collect weight stats
    For n = 0 To HUF_TABLELOG_MAX
        rankStats(n) = 0
    Next n
    weightTotal = 0
    For n = 0 To oSize - 1
        If huffWeight(n) >= HUF_TABLELOG_MAX Then HUF_readStats_body = ERR_corruption_detected: Exit Function
        rankStats(huffWeight(n)) = rankStats(huffWeight(n)) + 1
        weightTotal = weightTotal + pow2(huffWeight(n)) \ 2
    Next n
    If weightTotal = 0 Then HUF_readStats_body = ERR_corruption_detected: Exit Function

    ' get last non-null symbol weight (implied, total must be 2^n)
    tableLog = BIT_highbit32(weightTotal) + 1
    If tableLog > HUF_TABLELOG_MAX Then HUF_readStats_body = ERR_corruption_detected: Exit Function
    tableLogPtr = tableLog
    ' determine last weight
    total = pow2(tableLog)
    rest = total - weightTotal
    verif = pow2(BIT_highbit32(rest))
    lastWeight = BIT_highbit32(rest) + 1
    If verif <> rest Then HUF_readStats_body = ERR_corruption_detected: Exit Function ' last value must be a clean power of 2
    huffWeight(oSize) = lastWeight
    rankStats(lastWeight) = rankStats(lastWeight) + 1

    ' check tree construction validity
    If (rankStats(1) < 2) Or ((rankStats(1) And 1) <> 0) Then HUF_readStats_body = ERR_corruption_detected: Exit Function ' by construction : at least 2 elts of rank 1, must be even

    ' results
    nbSymbolsPtr = oSize + 1
    HUF_readStats_body = iSize + 1
End Function

'--------------------------------------------------------------------------
' fse_decompress.c (as built for HUF_readStats: FSE_FUNCTION_TYPE BYTE)

' FSE_buildDTable_internal: into this module's fseDT_* table. wkspSize is the work space left (bytes).
Private Function FSE_buildDTable_internal(normalizedCounter() As Long, ByVal maxSymbolValue As Long, ByVal tableLog As Long, ByVal wkspSize As Long) As Long
    Dim maxSV1 As Long, tableSize As Long, highThreshold As Long, largeLimit As Long, s As Long
    Dim tableMask As Long, step_ As Long, position As Long, i As Long, u As Long, pos As Long, n As Long
    Dim symbol As Long, nextState As Long
    Dim symbolNext(0 To 255) As Long, spread(0 To 4095 + 8) As Long

    maxSV1 = maxSymbolValue + 1
    tableSize = pow2(tableLog)
    highThreshold = tableSize - 1

    ' Sanity Checks: FSE_BUILD_DTABLE_WKSP_SIZE(tableLog, maxSymbolValue) > wkspSize
    If 2 * (maxSymbolValue + 1) + tableSize + 8 > wkspSize Then FSE_buildDTable_internal = ERR_maxSymbolValue_tooLarge: Exit Function
    If maxSymbolValue > FSE_MAX_SYMBOL_VALUE Then FSE_buildDTable_internal = ERR_maxSymbolValue_tooLarge: Exit Function
    If tableLog > FSE_MAX_TABLELOG Then FSE_buildDTable_internal = ERR_tableLog_tooLarge: Exit Function

    ' Init, lay down lowprob symbols
    fseDT_tableLog = tableLog
    fseDT_fastMode = 1
    largeLimit = pow2(tableLog - 1)
    For s = 0 To maxSV1 - 1
        If normalizedCounter(s) = -1 Then
            fseDT_symbol(highThreshold) = s
            highThreshold = highThreshold - 1
            symbolNext(s) = 1
        Else
            If normalizedCounter(s) >= largeLimit Then fseDT_fastMode = 0
            symbolNext(s) = normalizedCounter(s)
        End If
    Next s

    ' Spread symbols
    If highThreshold = tableSize - 1 Then
        tableMask = tableSize - 1
        step_ = (tableSize \ 2) + (tableSize \ 8) + 3     ' FSE_TABLESTEP
        ' First lay down the symbols in order (C: 8 at a time, sv = s * 0x0101010101010101)
        pos = 0
        For s = 0 To maxSV1 - 1
            n = normalizedCounter(s)
            For i = 0 To n - 1
                spread(pos + i) = s
            Next i
            pos = pos + n
        Next s
        ' Now we spread those positions across the table (unrolled twice in C)
        position = 0
        For s = 0 To tableSize - 1 Step 2
            For u = 0 To 1
                fseDT_symbol((position + (u * step_)) And tableMask) = spread(s + u)
            Next u
            position = (position + (2 * step_)) And tableMask
        Next s
    Else
        tableMask = tableSize - 1
        step_ = (tableSize \ 2) + (tableSize \ 8) + 3
        position = 0
        For s = 0 To maxSV1 - 1
            For i = 0 To normalizedCounter(s) - 1
                fseDT_symbol(position) = s
                position = (position + step_) And tableMask
                Do While position > highThreshold   ' lowprob area
                    position = (position + step_) And tableMask
                Loop
            Next i
        Next s
        If position <> 0 Then FSE_buildDTable_internal = ERR_GENERIC: Exit Function ' position must reach all cells once, otherwise normalizedCounter is incorrect
    End If

    ' Build Decoding table
    For u = 0 To tableSize - 1
        symbol = fseDT_symbol(u)
        nextState = symbolNext(symbol)
        symbolNext(symbol) = (nextState + 1) And &HFFFF&     ' (U16)
        fseDT_nbBits(u) = (tableLog - BIT_highbit32(nextState)) And &HFF&
        fseDT_newState(u) = ((nextState * pow2(fseDT_nbBits(u))) - tableSize) And &HFFFF&
    Next u
    FSE_buildDTable_internal = 0
End Function

' FSE_decompress_usingDTable_generic over fseDT_*: dst(0 .. maxDstSize - 1) (the Huffman weights)
Private Function FSE_decompress_usingDTable_generic(dst() As Byte, ByVal maxDstSize As Long, cSrc() As Byte, ByVal ip As Long, ByVal cSrcSize As Long, ByVal fast As Boolean) As Long
    Dim op As Long, omax As Long, olimit As Long, r As Long
    Dim bitD As BIT_DStream_t, state1 As FSE_DState_t, state2 As FSE_DState_t

    op = 0
    omax = maxDstSize
    olimit = omax - 3

    ' Init
    r = BIT_initDStream(bitD, cSrc, ip, cSrcSize)
    If zstdIsError(r) Then FSE_decompress_usingDTable_generic = r: Exit Function

    FSE_initDState state1, bitD, cSrc
    FSE_initDState state2, bitD, cSrc

    ' 4 symbols per loop
    Do While (BIT_reloadDStream(bitD, cSrc) = BIT_DStream_unfinished) And (op < olimit)
        dst(op + 0) = FSE_getSymbol(state1, bitD, fast)
        ' (FSE_MAX_TABLELOG*2+7 > 32 is false: no reload)
        dst(op + 1) = FSE_getSymbol(state2, bitD, fast)
        ' (FSE_MAX_TABLELOG*4+7 > 32)
        If BIT_reloadDStream(bitD, cSrc) > BIT_DStream_unfinished Then op = op + 2: GoTo tail
        dst(op + 2) = FSE_getSymbol(state1, bitD, fast)
        dst(op + 3) = FSE_getSymbol(state2, bitD, fast)
        op = op + 4
    Loop

tail:
    ' note : BIT_reloadDStream(&bitD) >= FSE_DStream_partiallyFilled; Ends at exactly BIT_DStream_completed
    Do
        If op > omax - 2 Then FSE_decompress_usingDTable_generic = ERR_dstSize_tooSmall: Exit Function
        dst(op) = FSE_getSymbol(state1, bitD, fast)
        op = op + 1
        If BIT_reloadDStream(bitD, cSrc) = BIT_DStream_overflow Then
            dst(op) = FSE_getSymbol(state2, bitD, fast)
            op = op + 1
            Exit Do
        End If

        If op > omax - 2 Then FSE_decompress_usingDTable_generic = ERR_dstSize_tooSmall: Exit Function
        dst(op) = FSE_getSymbol(state2, bitD, fast)
        op = op + 1
        If BIT_reloadDStream(bitD, cSrc) = BIT_DStream_overflow Then
            dst(op) = FSE_getSymbol(state1, bitD, fast)
            op = op + 1
            Exit Do
        End If
    Loop
    FSE_decompress_usingDTable_generic = op
End Function

' FSE_initDState
Private Sub FSE_initDState(DStatePtr As FSE_DState_t, bitD As BIT_DStream_t, src() As Byte)
    DStatePtr.state = BIT_readBits(bitD, fseDT_tableLog)
    BIT_reloadDStream bitD, src
End Sub

' FSE_GETSYMBOL: fast ? FSE_decodeSymbolFast : FSE_decodeSymbol
Private Function FSE_getSymbol(DStatePtr As FSE_DState_t, bitD As BIT_DStream_t, ByVal fast As Boolean) As Byte
    Dim nbBits As Long, lowBits As Long, st As Long
    st = DStatePtr.state
    nbBits = fseDT_nbBits(st)
    If fast Then lowBits = BIT_readBitsFast(bitD, nbBits) Else lowBits = BIT_readBits(bitD, nbBits)
    DStatePtr.state = fseDT_newState(st) + lowBits
    FSE_getSymbol = fseDT_symbol(st)
End Function

' FSE_decompress_wksp_body (bmi2 0): dst(0 .. dstCapacity - 1); wkspSize is the caller's work space (bytes)
Private Function FSE_decompress_wksp_body(dst() As Byte, ByVal dstCapacity As Long, cSrc() As Byte, ByVal istart As Long, ByVal cSrcSize As Long, ByVal maxLog As Long, ByVal wkspSize As Long) As Long
    Dim ip As Long, tableLog As Long, maxSymbolValue As Long, NCountLength As Long, r As Long
    Dim counting(0 To FSE_MAX_SYMBOL_VALUE) As Long, dtSizeBytes As Long

    ip = istart
    maxSymbolValue = FSE_MAX_SYMBOL_VALUE

    ' normal FSE decoding mode
    NCountLength = FSE_readNCount(counting, maxSymbolValue, tableLog, cSrc, istart, cSrcSize)
    If zstdIsError(NCountLength) Then FSE_decompress_wksp_body = NCountLength: Exit Function
    If tableLog > maxLog Then FSE_decompress_wksp_body = ERR_tableLog_tooLarge: Exit Function
    ip = ip + NCountLength
    cSrcSize = cSrcSize - NCountLength

    ' FSE_DECOMPRESS_WKSP_SIZE(tableLog, maxSymbolValue) > wkspSize: (FSE_DTABLE_SIZE_U32 +
    ' FSE_BUILD_DTABLE_WKSP_SIZE_U32) * 4
    dtSizeBytes = (1 + pow2(tableLog)) * 4
    If dtSizeBytes + ((2 * (maxSymbolValue + 1) + pow2(tableLog) + 8 + 3) \ 4) * 4 > wkspSize Then FSE_decompress_wksp_body = ERR_tableLog_tooLarge: Exit Function
    wkspSize = wkspSize - dtSizeBytes

    r = FSE_buildDTable_internal(counting, maxSymbolValue, tableLog, wkspSize)
    If zstdIsError(r) Then FSE_decompress_wksp_body = r: Exit Function

    ' select fast mode (static)
    FSE_decompress_wksp_body = FSE_decompress_usingDTable_generic(dst, dstCapacity, cSrc, ip, cSrcSize, fseDT_fastMode <> 0)
End Function

'--------------------------------------------------------------------------
' huf_decompress.c

' (ZSTD_decompressBegin) dctx->entropy.hufTable[0] = (HUF_DTable)((HufLog)*0x1000001): maxTableLog 12, tableType 0,
' tableLog 0, reserved 12
Public Sub HUF_vbResetDTable()
    If Not initialised Then initTables
    dtd_maxTableLog = 12
    dtd_tableType = 0
    dtd_tableLog = 0
    dtd_reserved = 12
End Sub

' HUF_readDTableX1_wksp_bmi2: from src(ip0 .. ip0 + srcSize - 1); wkspSize is the caller's (bytes)
Private Function HUF_readDTableX1_wksp(src() As Byte, ByVal ip0 As Long, ByVal srcSize As Long, ByVal wkspSize As Long) As Long
    Dim tableLog As Long, nbSymbols As Long, iSize As Long
    Dim rankVal(0 To HUF_TABLELOG_ABSOLUTEMAX) As Long, rankStart(0 To HUF_TABLELOG_ABSOLUTEMAX) As Long
    Dim symbols(0 To HUF_SYMBOLVALUE_MAX) As Byte, huffWeight(0 To HUF_SYMBOLVALUE_MAX) As Byte
    Dim n As Long, nextRankStart As Long, curr As Long, w As Long, symbol As Long, rankStart_ As Long
    Dim symbolCount As Long, length_ As Long, uStart As Long, nbBits As Long, s As Long, u As Long, b As Byte

    tableLog = 0
    nbSymbols = 0
    ' sizeof(HUF_ReadDTableX1_Workspace): rankVal 16, rankStart 16, statsWksp 89 U32s, symbols 256, huffWeight 256
    If (16 + 16 + 89) * 4 + 256 + 256 > wkspSize Then HUF_readDTableX1_wksp = ERR_tableLog_tooLarge: Exit Function

    iSize = HUF_readStats_body(huffWeight, HUF_SYMBOLVALUE_MAX + 1, rankVal, nbSymbols, tableLog, src, ip0, srcSize, HUF_READ_STATS_WORKSPACE_SIZE)
    If zstdIsError(iSize) Then HUF_readDTableX1_wksp = iSize: Exit Function

    ' Table header
    If tableLog > dtd_maxTableLog + 1 Then HUF_readDTableX1_wksp = ERR_tableLog_tooLarge: Exit Function ' DTable too small, Huffman tree cannot fit in
    dtd_tableType = 0
    dtd_tableLog = tableLog

    ' Compute symbols and rankStart given rankVal (C unrolls the symbols' loop by 4: the same order)
    nextRankStart = 0
    For n = 0 To tableLog
        curr = nextRankStart
        nextRankStart = nextRankStart + rankVal(n)
        rankStart(n) = curr
    Next n
    For n = 0 To nbSymbols - 1
        w = huffWeight(n)
        symbols(rankStart(w)) = n
        rankStart(w) = rankStart(w) + 1
    Next n

    ' fill DTable: each symbol of weight w fills length = (1 << w) >> 1 cells (C: switch on the length, writing
    ' 1, 2, 4 or 8 cells at a time: the same cells)
    symbol = rankVal(0)
    rankStart_ = 0
    For w = 1 To tableLog
        symbolCount = rankVal(w)
        length_ = pow2(w) \ 2
        uStart = rankStart_
        nbBits = tableLog + 1 - w
        For s = 0 To symbolCount - 1
            b = symbols(symbol + s)
            For u = 0 To length_ - 1
                dtX1byte(uStart + u) = b
                dtX1nbBits(uStart + u) = nbBits
            Next u
            uStart = uStart + length_
        Next s
        symbol = symbol + symbolCount
        rankStart_ = rankStart_ + symbolCount * length_
    Next w
    HUF_readDTableX1_wksp = iSize
End Function

' HUF_decodeStreamX1: p .. pEnd in dst; the X1 table with dtLog
Private Function HUF_decodeStreamX1(dst() As Byte, ByVal p As Long, bitD As BIT_DStream_t, src() As Byte, ByVal pEnd As Long, ByVal dtLog As Long) As Long
    Dim pStart As Long, v As Long, x As Long, dv As Long, tb As Long
    pStart = p
    dv = pow2((32 - dtLog) And 31)  ' (BIT_lookBitsFast(bitD, dtLog) written out below)
    tb = pow2(31 - ((32 - dtLog) And 31))

    ' up to 4 symbols at a time (32-bit: HUF_DECODE_SYMBOLX1_2 is the 64-bit build's: 2 symbols)
    Do While (BIT_reloadDStream(bitD, src) = BIT_DStream_unfinished) And (p < pEnd - 3)
        x = bitD.bitContainer * pow2(bitD.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb    ' HUF_DECODE_SYMBOLX1_1
        dst(p) = dtX1byte(v)
        bitD.bitsConsumed = bitD.bitsConsumed + dtX1nbBits(v)
        p = p + 1
        x = bitD.bitContainer * pow2(bitD.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb    ' HUF_DECODE_SYMBOLX1_0
        dst(p) = dtX1byte(v)
        bitD.bitsConsumed = bitD.bitsConsumed + dtX1nbBits(v)
        p = p + 1
    Loop

    ' [0-3] symbols remaining
    ' (MEM_32bits())
    Do While (BIT_reloadDStream(bitD, src) = BIT_DStream_unfinished) And (p < pEnd)
        x = bitD.bitContainer * pow2(bitD.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(p) = dtX1byte(v)
        bitD.bitsConsumed = bitD.bitsConsumed + dtX1nbBits(v)
        p = p + 1
    Loop

    ' no more data to retrieve from bitstream, no need to reload
    Do While p < pEnd
        x = bitD.bitContainer * pow2(bitD.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(p) = dtX1byte(v)
        bitD.bitsConsumed = bitD.bitsConsumed + dtX1nbBits(v)
        p = p + 1
    Loop

    HUF_decodeStreamX1 = pEnd - pStart
End Function

' HUF_decompress1X1_usingDTable_internal_body
Private Function HUF_decompress1X1_usingDTable_internal(dst() As Byte, ByVal op As Long, ByVal dstSize As Long, cSrc() As Byte, ByVal ip As Long, ByVal cSrcSize As Long) As Long
    Dim oend As Long, bitD As BIT_DStream_t, dtLog As Long, r As Long
    oend = op + dstSize
    dtLog = dtd_tableLog

    r = BIT_initDStream(bitD, cSrc, ip, cSrcSize)
    If zstdIsError(r) Then HUF_decompress1X1_usingDTable_internal = r: Exit Function

    HUF_decodeStreamX1 dst, op, bitD, cSrc, oend, dtLog

    If Not BIT_endOfDStream(bitD) Then HUF_decompress1X1_usingDTable_internal = ERR_corruption_detected: Exit Function

    HUF_decompress1X1_usingDTable_internal = dstSize
End Function

' HUF_decompress4X1_usingDTable_internal_body
Private Function HUF_decompress4X1_usingDTable_internal(dst() As Byte, ByVal ostart As Long, ByVal dstSize As Long, cSrc() As Byte, ByVal istart As Long, ByVal cSrcSize As Long) As Long
    Dim oend As Long, olimit As Long, r As Long, v As Long
    Dim bitD1 As BIT_DStream_t, bitD2 As BIT_DStream_t, bitD3 As BIT_DStream_t, bitD4 As BIT_DStream_t
    Dim length1 As Long, length2 As Long, length3 As Long, length4 As Long
    Dim istart1 As Long, istart2 As Long, istart3 As Long, istart4 As Long
    Dim segmentSize As Long, opStart2 As Long, opStart3 As Long, opStart4 As Long
    Dim op1 As Long, op2 As Long, op3 As Long, op4 As Long, dtLog As Long, endSignal As Boolean
    Dim x As Long, dv As Long, tb As Long

    ' Check
    If cSrcSize < 10 Then HUF_decompress4X1_usingDTable_internal = ERR_corruption_detected: Exit Function ' strict minimum : jump table + 1 byte per stream

    oend = ostart + dstSize
    olimit = oend - 3
    length1 = MEM_readLE16(cSrc, istart)
    length2 = MEM_readLE16(cSrc, istart + 2)
    length3 = MEM_readLE16(cSrc, istart + 4)
    length4 = cSrcSize - (length1 + length2 + length3 + 6)
    istart1 = istart + 6    ' jumpTable
    istart2 = istart1 + length1
    istart3 = istart2 + length2
    istart4 = istart3 + length3
    segmentSize = (dstSize + 3) \ 4
    opStart2 = ostart + segmentSize
    opStart3 = opStart2 + segmentSize
    opStart4 = opStart3 + segmentSize
    op1 = ostart
    op2 = opStart2
    op3 = opStart3
    op4 = opStart4
    dtLog = dtd_tableLog
    endSignal = True
    dv = pow2((32 - dtLog) And 31)  ' (BIT_lookBitsFast(bitD, dtLog) written out below)
    tb = pow2(31 - ((32 - dtLog) And 31))

    If zULT(cSrcSize, length4) Then HUF_decompress4X1_usingDTable_internal = ERR_corruption_detected: Exit Function ' overflow
    r = BIT_initDStream(bitD1, cSrc, istart1, length1)
    If zstdIsError(r) Then HUF_decompress4X1_usingDTable_internal = r: Exit Function
    r = BIT_initDStream(bitD2, cSrc, istart2, length2)
    If zstdIsError(r) Then HUF_decompress4X1_usingDTable_internal = r: Exit Function
    r = BIT_initDStream(bitD3, cSrc, istart3, length3)
    If zstdIsError(r) Then HUF_decompress4X1_usingDTable_internal = r: Exit Function
    r = BIT_initDStream(bitD4, cSrc, istart4, length4)
    If zstdIsError(r) Then HUF_decompress4X1_usingDTable_internal = r: Exit Function

    ' up to 16 symbols per loop (4 symbols per stream) in 64-bit mode; 2 per stream here
    Do While endSignal And (op4 < olimit)
        x = bitD1.bitContainer * pow2(bitD1.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(op1) = dtX1byte(v): bitD1.bitsConsumed = bitD1.bitsConsumed + dtX1nbBits(v): op1 = op1 + 1
        x = bitD2.bitContainer * pow2(bitD2.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(op2) = dtX1byte(v): bitD2.bitsConsumed = bitD2.bitsConsumed + dtX1nbBits(v): op2 = op2 + 1
        x = bitD3.bitContainer * pow2(bitD3.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(op3) = dtX1byte(v): bitD3.bitsConsumed = bitD3.bitsConsumed + dtX1nbBits(v): op3 = op3 + 1
        x = bitD4.bitContainer * pow2(bitD4.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(op4) = dtX1byte(v): bitD4.bitsConsumed = bitD4.bitsConsumed + dtX1nbBits(v): op4 = op4 + 1
        x = bitD1.bitContainer * pow2(bitD1.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(op1) = dtX1byte(v): bitD1.bitsConsumed = bitD1.bitsConsumed + dtX1nbBits(v): op1 = op1 + 1
        x = bitD2.bitContainer * pow2(bitD2.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(op2) = dtX1byte(v): bitD2.bitsConsumed = bitD2.bitsConsumed + dtX1nbBits(v): op2 = op2 + 1
        x = bitD3.bitContainer * pow2(bitD3.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(op3) = dtX1byte(v): bitD3.bitsConsumed = bitD3.bitsConsumed + dtX1nbBits(v): op3 = op3 + 1
        x = bitD4.bitContainer * pow2(bitD4.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb
        dst(op4) = dtX1byte(v): bitD4.bitsConsumed = bitD4.bitsConsumed + dtX1nbBits(v): op4 = op4 + 1
        If BIT_reloadDStreamFast(bitD1, cSrc) <> BIT_DStream_unfinished Then endSignal = False
        If BIT_reloadDStreamFast(bitD2, cSrc) <> BIT_DStream_unfinished Then endSignal = False
        If BIT_reloadDStreamFast(bitD3, cSrc) <> BIT_DStream_unfinished Then endSignal = False
        If BIT_reloadDStreamFast(bitD4, cSrc) <> BIT_DStream_unfinished Then endSignal = False
    Loop

    ' check corruption (C: "should not be necessary", the checks of op1..op3 are commented out)

    ' finish bitStreams one by one
    HUF_decodeStreamX1 dst, op1, bitD1, cSrc, opStart2, dtLog
    HUF_decodeStreamX1 dst, op2, bitD2, cSrc, opStart3, dtLog
    HUF_decodeStreamX1 dst, op3, bitD3, cSrc, opStart4, dtLog
    HUF_decodeStreamX1 dst, op4, bitD4, cSrc, oend, dtLog

    ' check
    If Not (BIT_endOfDStream(bitD1) And BIT_endOfDStream(bitD2) And BIT_endOfDStream(bitD3) And BIT_endOfDStream(bitD4)) Then HUF_decompress4X1_usingDTable_internal = ERR_corruption_detected: Exit Function

    ' decoded size
    HUF_decompress4X1_usingDTable_internal = dstSize
End Function

' HUF_fillDTableX2Level2: into the X2 table from cell base; rankValOrigin is rankVal(row, ...); the sorted list is
' sortedSymbol/sortedWeight(sortedFirst ...)
Private Sub HUF_fillDTableX2Level2(ByVal base As Long, ByVal sizeLog As Long, ByVal consumed As Long, rankValOrigin() As Long, ByVal row As Long, ByVal minWeight As Long, sortedSymbol() As Long, sortedWeight() As Long, ByVal sortedFirst As Long, ByVal sortedListSize As Long, ByVal nbBitsBaseline As Long, ByVal baseSeq As Long)
    Dim rankVal(0 To HUF_TABLELOG_MAX) As Long, i As Long, skipSize As Long, s As Long
    Dim symbol As Long, weight As Long, nbBits As Long, length_ As Long, start As Long, end_ As Long, seq As Long

    ' get pre-calculated rankVal
    For i = 0 To HUF_TABLELOG_MAX
        rankVal(i) = rankValOrigin(row, i)
    Next i

    ' fill skipped values
    If minWeight > 1 Then
        skipSize = rankVal(minWeight)
        For i = 0 To skipSize - 1
            dtX2seq(base + i) = baseSeq
            dtX2nbBits(base + i) = consumed
            dtX2length(base + i) = 1
        Next i
    End If

    ' fill DTable
    For s = 0 To sortedListSize - 1     ' note : sortedSymbols already skipped
        symbol = sortedSymbol(sortedFirst + s)
        weight = sortedWeight(sortedFirst + s)
        nbBits = nbBitsBaseline - weight
        length_ = pow2(sizeLog - nbBits)
        start = rankVal(weight)
        i = start
        end_ = start + length_
        seq = (baseSeq + symbol * 256) And &HFFFF&
        Do  ' since length >= 1
            dtX2seq(base + i) = seq
            dtX2nbBits(base + i) = (nbBits + consumed) And &HFF&
            dtX2length(base + i) = 2
            i = i + 1
        Loop While i < end_
        rankVal(weight) = rankVal(weight) + length_
    Next s
End Sub

' HUF_fillDTableX2
Private Sub HUF_fillDTableX2(ByVal targetLog As Long, sortedSymbol() As Long, sortedWeight() As Long, ByVal sortedListSize As Long, rankStart() As Long, rankValOrigin() As Long, ByVal maxWeight As Long, ByVal nbBitsBaseline As Long)
    Dim rankVal(0 To HUF_TABLELOG_MAX) As Long, scaleLog As Long, minBits As Long, s As Long
    Dim symbol As Long, weight As Long, nbBits As Long, start As Long, length_ As Long
    Dim sortedRank As Long, minWeight As Long, u As Long, end_ As Long

    scaleLog = nbBitsBaseline - targetLog  ' note : targetLog >= srcLog, hence scaleLog <= 1
    minBits = nbBitsBaseline - maxWeight
    For s = 0 To HUF_TABLELOG_MAX
        rankVal(s) = rankValOrigin(0, s)
    Next s

    ' fill DTable
    For s = 0 To sortedListSize - 1
        symbol = sortedSymbol(s)
        weight = sortedWeight(s)
        nbBits = nbBitsBaseline - weight
        start = rankVal(weight)
        length_ = pow2(targetLog - nbBits)

        If Not zULT(targetLog - nbBits, minBits) Then  ' enough room for a second symbol
            minWeight = nbBits + scaleLog
            If minWeight < 1 Then minWeight = 1
            sortedRank = rankStart(minWeight)
            HUF_fillDTableX2Level2 start, targetLog - nbBits, nbBits, rankValOrigin, nbBits, minWeight, sortedSymbol, sortedWeight, sortedRank, sortedListSize - sortedRank, nbBitsBaseline, symbol
        Else
            end_ = start + length_
            For u = start To end_ - 1
                dtX2seq(u) = symbol
                dtX2nbBits(u) = nbBits
                dtX2length(u) = 1
            Next u
        End If
        rankVal(weight) = rankVal(weight) + length_
    Next s
End Sub

' HUF_readDTableX2_wksp: from src(ip0 .. ip0 + srcSize - 1); wkspSize is the caller's (bytes)
Private Function HUF_readDTableX2_wksp(src() As Byte, ByVal ip0 As Long, ByVal srcSize As Long, ByVal wkspSize As Long) As Long
    Dim tableLog As Long, maxW As Long, sizeOfSort As Long, nbSymbols As Long, maxTableLog As Long, iSize As Long
    Dim rankVal(0 To HUF_TABLELOG_MAX - 1, 0 To HUF_TABLELOG_MAX) As Long
    Dim rankStats(0 To HUF_TABLELOG_MAX) As Long, rankStart0(0 To HUF_TABLELOG_MAX + 1) As Long
    Dim sortedSymbol(0 To HUF_SYMBOLVALUE_MAX) As Long, sortedWeight(0 To HUF_SYMBOLVALUE_MAX) As Long
    Dim weightList(0 To HUF_SYMBOLVALUE_MAX) As Byte
    Dim w As Long, nextRankStart As Long, curr As Long, s As Long, r As Long
    Dim rescale As Long, nextRankVal As Long, minBits As Long, consumed As Long

    maxTableLog = dtd_maxTableLog
    ' spaceUsed32: rankVal 12 * 13, rankStats 13, rankStart0 14, sortedSymbol 128, weightList 64 U32s
    If (12 * 13 + 13 + 14 + 128 + 64) * 4 > wkspSize Then HUF_readDTableX2_wksp = ERR_tableLog_tooLarge: Exit Function

    ' rankStart = rankStart0 + 1: rankStart(w) is rankStart0(w + 1). (ZSTD_memset of rankStats and rankStart0:
    ' the locals start at 0.)
    If maxTableLog > HUF_TABLELOG_MAX Then HUF_readDTableX2_wksp = ERR_tableLog_tooLarge: Exit Function

    iSize = HUF_readStats_body(weightList, HUF_SYMBOLVALUE_MAX + 1, rankStats, nbSymbols, tableLog, src, ip0, srcSize, HUF_READ_STATS_WORKSPACE_SIZE)
    If zstdIsError(iSize) Then HUF_readDTableX2_wksp = iSize: Exit Function

    ' check result
    If tableLog > maxTableLog Then HUF_readDTableX2_wksp = ERR_tableLog_tooLarge: Exit Function ' DTable can't fit code depth

    ' find maxWeight
    maxW = tableLog
    Do While rankStats(maxW) = 0    ' necessarily finds a solution before 0
        maxW = maxW - 1
    Loop

    ' Get start index of each weight
    nextRankStart = 0
    For w = 1 To maxW
        curr = nextRankStart
        nextRankStart = nextRankStart + rankStats(w)
        rankStart0(w + 1) = curr
    Next w
    rankStart0(0 + 1) = nextRankStart   ' put all 0w symbols at the end of sorted list
    sizeOfSort = nextRankStart

    ' sort symbols by weight
    For s = 0 To nbSymbols - 1
        w = weightList(s)
        r = rankStart0(w + 1)
        rankStart0(w + 1) = r + 1
        sortedSymbol(r) = s
        sortedWeight(r) = w
    Next s
    rankStart0(0 + 1) = 0   ' forget 0w symbols; this is beginning of weight(1)

    ' Build rankVal
    rescale = (maxTableLog - tableLog) - 1 ' tableLog <= maxTableLog
    nextRankVal = 0
    For w = 1 To maxW
        curr = nextRankVal
        nextRankVal = nextRankVal + rankStats(w) * pow2(w + rescale)
        rankVal(0, w) = curr
    Next w
    minBits = tableLog + 1 - maxW
    For consumed = minBits To maxTableLog - minBits
        For w = 1 To maxW
            rankVal(consumed, w) = rankVal(0, w) \ pow2(consumed)
        Next w
    Next consumed

    HUF_fillDTableX2 maxTableLog, sortedSymbol, sortedWeight, sizeOfSort, rankStart0, rankVal, maxW, tableLog + 1

    dtd_tableLog = maxTableLog
    dtd_tableType = 1
    HUF_readDTableX2_wksp = iSize
End Function

' HUF_decodeLastSymbolX2
Private Sub HUF_decodeLastSymbolX2(dst() As Byte, ByVal op As Long, bitD As BIT_DStream_t, ByVal dtLog As Long)
    Dim v As Long
    v = BIT_lookBitsFast(bitD, dtLog)   ' note : dtLog >= 1
    dst(op) = dtX2seq(v) And &HFF&
    If dtX2length(v) = 1 Then
        bitD.bitsConsumed = bitD.bitsConsumed + dtX2nbBits(v)
    Else
        If bitD.bitsConsumed < 32 Then
            bitD.bitsConsumed = bitD.bitsConsumed + dtX2nbBits(v)
            If bitD.bitsConsumed > 32 Then
                ' ugly hack; works only because it's the last symbol. Note : can't easily extract nbBits from just this symbol
                bitD.bitsConsumed = 32
            End If
        End If
    End If
End Sub

' HUF_decodeSymbolX2 (copies both bytes of the cell, returns its length): p advances by the result. dv and tb are
' BIT_lookBitsFast's for dtLog (its caller's).
Private Function HUF_decodeSymbolX2(dst() As Byte, ByVal op As Long, bitD As BIT_DStream_t, ByVal dv As Long, ByVal tb As Long) As Long
    Dim v As Long, sq As Long, x As Long
    x = bitD.bitContainer * pow2(bitD.bitsConsumed And 31): If x >= 0 Then v = x \ dv Else v = ((x And &H7FFFFFFF) \ dv) Or tb   ' BIT_lookBitsFast(DStream, dtLog): dtLog >= 1
    sq = dtX2seq(v)
    dst(op) = sq And &HFF&
    dst(op + 1) = sq \ &H100&
    bitD.bitsConsumed = bitD.bitsConsumed + dtX2nbBits(v)
    HUF_decodeSymbolX2 = dtX2length(v)
End Function

' HUF_decodeStreamX2
Private Function HUF_decodeStreamX2(dst() As Byte, ByVal p As Long, bitD As BIT_DStream_t, src() As Byte, ByVal pEnd As Long, ByVal dtLog As Long) As Long
    Dim pStart As Long, dv As Long, tb As Long
    pStart = p
    dv = pow2((32 - dtLog) And 31)  ' (HUF_decodeSymbolX2's)
    tb = pow2(31 - ((32 - dtLog) And 31))

    ' up to 8 symbols at a time (32-bit: HUF_DECODE_SYMBOLX2_1 and _0)
    Do While (BIT_reloadDStream(bitD, src) = BIT_DStream_unfinished) And (p < pEnd - 3)
        p = p + HUF_decodeSymbolX2(dst, p, bitD, dv, tb)
        p = p + HUF_decodeSymbolX2(dst, p, bitD, dv, tb)
    Loop

    ' closer to end : up to 2 symbols at a time
    Do While (BIT_reloadDStream(bitD, src) = BIT_DStream_unfinished) And (p <= pEnd - 2)
        p = p + HUF_decodeSymbolX2(dst, p, bitD, dv, tb)
    Loop

    Do While p <= pEnd - 2
        p = p + HUF_decodeSymbolX2(dst, p, bitD, dv, tb)   ' no need to reload : reached the end of DStream
    Loop

    If p < pEnd Then
        HUF_decodeLastSymbolX2 dst, p, bitD, dtLog
        p = p + 1
    End If

    HUF_decodeStreamX2 = p - pStart
End Function

' HUF_decompress1X2_usingDTable_internal_body
Private Function HUF_decompress1X2_usingDTable_internal(dst() As Byte, ByVal ostart As Long, ByVal dstSize As Long, cSrc() As Byte, ByVal ip As Long, ByVal cSrcSize As Long) As Long
    Dim bitD As BIT_DStream_t, r As Long

    ' Init
    r = BIT_initDStream(bitD, cSrc, ip, cSrcSize)
    If zstdIsError(r) Then HUF_decompress1X2_usingDTable_internal = r: Exit Function

    ' decode
    HUF_decodeStreamX2 dst, ostart, bitD, cSrc, ostart + dstSize, dtd_tableLog

    ' check
    If Not BIT_endOfDStream(bitD) Then HUF_decompress1X2_usingDTable_internal = ERR_corruption_detected: Exit Function

    ' decoded size
    HUF_decompress1X2_usingDTable_internal = dstSize
End Function

' HUF_decompress4X2_usingDTable_internal_body
Private Function HUF_decompress4X2_usingDTable_internal(dst() As Byte, ByVal ostart As Long, ByVal dstSize As Long, cSrc() As Byte, ByVal istart As Long, ByVal cSrcSize As Long) As Long
    Dim oend As Long, olimit As Long, r As Long
    Dim bitD1 As BIT_DStream_t, bitD2 As BIT_DStream_t, bitD3 As BIT_DStream_t, bitD4 As BIT_DStream_t
    Dim length1 As Long, length2 As Long, length3 As Long, length4 As Long
    Dim istart1 As Long, istart2 As Long, istart3 As Long, istart4 As Long
    Dim segmentSize As Long, opStart2 As Long, opStart3 As Long, opStart4 As Long
    Dim op1 As Long, op2 As Long, op3 As Long, op4 As Long, dtLog As Long, endSignal As Boolean
    Dim dv As Long, tb As Long

    If cSrcSize < 10 Then HUF_decompress4X2_usingDTable_internal = ERR_corruption_detected: Exit Function ' strict minimum : jump table + 1 byte per stream

    oend = ostart + dstSize
    olimit = oend - 3       ' oend - (sizeof(size_t)-1)
    length1 = MEM_readLE16(cSrc, istart)
    length2 = MEM_readLE16(cSrc, istart + 2)
    length3 = MEM_readLE16(cSrc, istart + 4)
    length4 = cSrcSize - (length1 + length2 + length3 + 6)
    istart1 = istart + 6    ' jumpTable
    istart2 = istart1 + length1
    istart3 = istart2 + length2
    istart4 = istart3 + length3
    segmentSize = (dstSize + 3) \ 4
    opStart2 = ostart + segmentSize
    opStart3 = opStart2 + segmentSize
    opStart4 = opStart3 + segmentSize
    op1 = ostart
    op2 = opStart2
    op3 = opStart3
    op4 = opStart4
    endSignal = True
    dtLog = dtd_tableLog
    dv = pow2((32 - dtLog) And 31)  ' (HUF_decodeSymbolX2's)
    tb = pow2(31 - ((32 - dtLog) And 31))

    If zULT(cSrcSize, length4) Then HUF_decompress4X2_usingDTable_internal = ERR_corruption_detected: Exit Function ' overflow
    r = BIT_initDStream(bitD1, cSrc, istart1, length1)
    If zstdIsError(r) Then HUF_decompress4X2_usingDTable_internal = r: Exit Function
    r = BIT_initDStream(bitD2, cSrc, istart2, length2)
    If zstdIsError(r) Then HUF_decompress4X2_usingDTable_internal = r: Exit Function
    r = BIT_initDStream(bitD3, cSrc, istart3, length3)
    If zstdIsError(r) Then HUF_decompress4X2_usingDTable_internal = r: Exit Function
    r = BIT_initDStream(bitD4, cSrc, istart4, length4)
    If zstdIsError(r) Then HUF_decompress4X2_usingDTable_internal = r: Exit Function

    ' 16-32 symbols per loop (4-8 symbols per stream). (The #else branch, compiled by gcc and MSVC: clang's decodes
    ' each stream's symbols in another order, with reloads between; both regenerate the same bytes from a valid
    ' stream. The 32-bit build's HUF_DECODE_SYMBOLX2_1 and _0 only.)
    Do While endSignal And (op4 < olimit)
        op1 = op1 + HUF_decodeSymbolX2(dst, op1, bitD1, dv, tb)  ' HUF_DECODE_SYMBOLX2_1
        op2 = op2 + HUF_decodeSymbolX2(dst, op2, bitD2, dv, tb)
        op3 = op3 + HUF_decodeSymbolX2(dst, op3, bitD3, dv, tb)
        op4 = op4 + HUF_decodeSymbolX2(dst, op4, bitD4, dv, tb)
        op1 = op1 + HUF_decodeSymbolX2(dst, op1, bitD1, dv, tb)  ' HUF_DECODE_SYMBOLX2_0
        op2 = op2 + HUF_decodeSymbolX2(dst, op2, bitD2, dv, tb)
        op3 = op3 + HUF_decodeSymbolX2(dst, op3, bitD3, dv, tb)
        op4 = op4 + HUF_decodeSymbolX2(dst, op4, bitD4, dv, tb)
        endSignal = True
        If BIT_reloadDStreamFast(bitD1, cSrc) <> BIT_DStream_unfinished Then endSignal = False
        If BIT_reloadDStreamFast(bitD2, cSrc) <> BIT_DStream_unfinished Then endSignal = False
        If BIT_reloadDStreamFast(bitD3, cSrc) <> BIT_DStream_unfinished Then endSignal = False
        If BIT_reloadDStreamFast(bitD4, cSrc) <> BIT_DStream_unfinished Then endSignal = False
    Loop

    ' check corruption
    If op1 > opStart2 Then HUF_decompress4X2_usingDTable_internal = ERR_corruption_detected: Exit Function
    If op2 > opStart3 Then HUF_decompress4X2_usingDTable_internal = ERR_corruption_detected: Exit Function
    If op3 > opStart4 Then HUF_decompress4X2_usingDTable_internal = ERR_corruption_detected: Exit Function
    ' note : op4 already verified within main loop

    ' finish bitStreams one by one
    HUF_decodeStreamX2 dst, op1, bitD1, cSrc, opStart2, dtLog
    HUF_decodeStreamX2 dst, op2, bitD2, cSrc, opStart3, dtLog
    HUF_decodeStreamX2 dst, op3, bitD3, cSrc, opStart4, dtLog
    HUF_decodeStreamX2 dst, op4, bitD4, cSrc, oend, dtLog

    ' check
    If Not (BIT_endOfDStream(bitD1) And BIT_endOfDStream(bitD2) And BIT_endOfDStream(bitD3) And BIT_endOfDStream(bitD4)) Then HUF_decompress4X2_usingDTable_internal = ERR_corruption_detected: Exit Function

    ' decoded size
    HUF_decompress4X2_usingDTable_internal = dstSize
End Function

' HUF_selectDecoder: 0 (X1) or 1 (X2)
Private Function HUF_selectDecoder(ByVal dstSize As Long, ByVal cSrcSize As Long) As Long
    Dim Q As Long, D256 As Long, DTime0 As Long, DTime1 As Long
    ' decoder timing evaluation
    If cSrcSize >= dstSize Then Q = 15 Else Q = (cSrcSize * 16) \ dstSize ' Q < 16
    D256 = dstSize \ 256
    DTime0 = algoTime_table(Q, 0) + (algoTime_decode(Q, 0) * D256)
    DTime1 = algoTime_table(Q, 1) + (algoTime_decode(Q, 1) * D256)
    DTime1 = DTime1 + DTime1 \ 8  ' advantage to algorithm using less memory, to reduce cache eviction
    If zULT(DTime1, DTime0) Then HUF_selectDecoder = 1 Else HUF_selectDecoder = 0
End Function

' HUF_decompress1X1_DCtx_wksp_bmi2 (dctx: the DCtx's table; the work space is the DCtx's, wkspSize bytes)
Public Function HUF_decompress1X1_DCtx_wksp_bmi2(dst() As Byte, ByVal op As Long, ByVal dstSize As Long, cSrc() As Byte, ByVal ip As Long, ByVal cSrcSize As Long, ByVal wkspSize As Long) As Long
    Dim hSize As Long
    If Not initialised Then initTables
    hSize = HUF_readDTableX1_wksp(cSrc, ip, cSrcSize, wkspSize)
    If zstdIsError(hSize) Then HUF_decompress1X1_DCtx_wksp_bmi2 = hSize: Exit Function
    If hSize >= cSrcSize Then HUF_decompress1X1_DCtx_wksp_bmi2 = ERR_srcSize_wrong: Exit Function
    ip = ip + hSize: cSrcSize = cSrcSize - hSize
    HUF_decompress1X1_DCtx_wksp_bmi2 = HUF_decompress1X1_usingDTable_internal(dst, op, dstSize, cSrc, ip, cSrcSize)
End Function

' HUF_decompress4X_hufOnly_wksp_bmi2
Public Function HUF_decompress4X_hufOnly_wksp_bmi2(dst() As Byte, ByVal op As Long, ByVal dstSize As Long, cSrc() As Byte, ByVal ip As Long, ByVal cSrcSize As Long, ByVal wkspSize As Long) As Long
    Dim hSize As Long
    If Not initialised Then initTables
    ' validation checks
    If dstSize = 0 Then HUF_decompress4X_hufOnly_wksp_bmi2 = ERR_dstSize_tooSmall: Exit Function
    If cSrcSize = 0 Then HUF_decompress4X_hufOnly_wksp_bmi2 = ERR_corruption_detected: Exit Function

    If HUF_selectDecoder(dstSize, cSrcSize) <> 0 Then
        ' HUF_decompress4X2_DCtx_wksp_bmi2
        hSize = HUF_readDTableX2_wksp(cSrc, ip, cSrcSize, wkspSize)
        If zstdIsError(hSize) Then HUF_decompress4X_hufOnly_wksp_bmi2 = hSize: Exit Function
        If hSize >= cSrcSize Then HUF_decompress4X_hufOnly_wksp_bmi2 = ERR_srcSize_wrong: Exit Function
        ip = ip + hSize: cSrcSize = cSrcSize - hSize
        HUF_decompress4X_hufOnly_wksp_bmi2 = HUF_decompress4X2_usingDTable_internal(dst, op, dstSize, cSrc, ip, cSrcSize)
    Else
        ' HUF_decompress4X1_DCtx_wksp_bmi2
        hSize = HUF_readDTableX1_wksp(cSrc, ip, cSrcSize, wkspSize)
        If zstdIsError(hSize) Then HUF_decompress4X_hufOnly_wksp_bmi2 = hSize: Exit Function
        If hSize >= cSrcSize Then HUF_decompress4X_hufOnly_wksp_bmi2 = ERR_srcSize_wrong: Exit Function
        ip = ip + hSize: cSrcSize = cSrcSize - hSize
        HUF_decompress4X_hufOnly_wksp_bmi2 = HUF_decompress4X1_usingDTable_internal(dst, op, dstSize, cSrc, ip, cSrcSize)
    End If
End Function

' HUF_decompress1X_usingDTable_bmi2 (the DCtx's table, of either type)
Public Function HUF_decompress1X_usingDTable_bmi2(dst() As Byte, ByVal op As Long, ByVal maxDstSize As Long, cSrc() As Byte, ByVal ip As Long, ByVal cSrcSize As Long) As Long
    If Not initialised Then initTables
    If dtd_tableType <> 0 Then
        HUF_decompress1X_usingDTable_bmi2 = HUF_decompress1X2_usingDTable_internal(dst, op, maxDstSize, cSrc, ip, cSrcSize)
    Else
        HUF_decompress1X_usingDTable_bmi2 = HUF_decompress1X1_usingDTable_internal(dst, op, maxDstSize, cSrc, ip, cSrcSize)
    End If
End Function

' HUF_decompress4X_usingDTable_bmi2 (the DCtx's table, of either type)
Public Function HUF_decompress4X_usingDTable_bmi2(dst() As Byte, ByVal op As Long, ByVal maxDstSize As Long, cSrc() As Byte, ByVal ip As Long, ByVal cSrcSize As Long) As Long
    If Not initialised Then initTables
    If dtd_tableType <> 0 Then
        HUF_decompress4X_usingDTable_bmi2 = HUF_decompress4X2_usingDTable_internal(dst, op, maxDstSize, cSrc, ip, cSrcSize)
    Else
        HUF_decompress4X_usingDTable_bmi2 = HUF_decompress4X1_usingDTable_internal(dst, op, maxDstSize, cSrc, ip, cSrcSize)
    End If
End Function
