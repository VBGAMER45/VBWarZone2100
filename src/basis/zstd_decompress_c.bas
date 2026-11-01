Attribute VB_Name = "zstd_decompress_c"
' zstd 1.4.9's decompress/zstd_decompress.c as ref/basis_universal/zstd/zstddeclib.c has it (the single-file
' decoder Basis Universal ships; 4.7.0's KTX2 files are Zstandard supercompressed): ZSTD_decompress (single-shot)
' over its frames (zstd frames and skippable frames, one after another), each frame's header and blocks (raw,
' RLE, compressed: zstd_decompress_block_c) and its content checksum (zstd_xxhash_c); with common/
' error_private.c's ZSTD_isError / ZSTD_getErrorName. The 32-bit build: windowLog at most ZSTD_WINDOWLOG_MAX_32
' (30). zstddeclib.c is built with ZSTD_LEGACY_SUPPORT 0 (no legacy frames) and ZSTD_STRIP_ERROR_STRINGS (every
' error name is "Error strings stripped").
'
' Left out: streaming (ZSTD_decompressStream, ZSTD_decompressContinue), dictionaries (ZSTD_DDict, ZSTD_decompress_
' usingDict: a frame that names a dictionary fails with dictionary_wrong, as ZSTD_decompress's DCtx has none),
' multithreading. ZSTD_decompress's DCtx (malloc'd per call in C) is the modules' state, set up by each call.
'
' A size_t result is a Long: a size, or an error, (size_t)-ZSTD_error_xxx (-1 to -119 here: ZSTD_isError). C's
' const void* src / void* dst are a Byte array and an offset.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' zstd_errors.h: ZSTD_ErrorCode (Private: a DLL's Public Enum would go into its type library)
Private Enum ZSTD_ErrorCode
    ZSTD_error_no_error = 0
    ZSTD_error_GENERIC = 1
    ZSTD_error_prefix_unknown = 10
    ZSTD_error_version_unsupported = 12
    ZSTD_error_frameParameter_unsupported = 14
    ZSTD_error_frameParameter_windowTooLarge = 16
    ZSTD_error_corruption_detected = 20
    ZSTD_error_checksum_wrong = 22
    ZSTD_error_dictionary_corrupted = 30
    ZSTD_error_dictionary_wrong = 32
    ZSTD_error_dictionaryCreation_failed = 34
    ZSTD_error_parameter_unsupported = 40
    ZSTD_error_parameter_outOfBound = 42
    ZSTD_error_tableLog_tooLarge = 44
    ZSTD_error_maxSymbolValue_tooLarge = 46
    ZSTD_error_maxSymbolValue_tooSmall = 48
    ZSTD_error_stage_wrong = 60
    ZSTD_error_init_missing = 62
    ZSTD_error_memory_allocation = 64
    ZSTD_error_workSpace_tooSmall = 66
    ZSTD_error_dstSize_tooSmall = 70
    ZSTD_error_srcSize_wrong = 72
    ZSTD_error_dstBuffer_null = 74
    ZSTD_error_frameIndex_tooLarge = 100
    ZSTD_error_seekableIO = 102
    ZSTD_error_dstBuffer_wrong = 104
    ZSTD_error_srcBuffer_wrong = 105
    ZSTD_error_maxCode = 120
End Enum

Private Const ZSTD_MAGICNUMBER As Long = &HFD2FB528
Private Const ZSTD_MAGIC_SKIPPABLE_START As Long = &H184D2A50
Private Const ZSTD_MAGIC_SKIPPABLE_MASK As Long = &HFFFFFFF0
Private Const ZSTD_FRAMEIDSIZE As Long = 4
Private Const ZSTD_SKIPPABLEHEADERSIZE As Long = 8
Private Const ZSTD_FRAMEHEADERSIZE_PREFIX As Long = 5     ' (format ZSTD_f_zstd1)
Private Const ZSTD_FRAMEHEADERSIZE_MIN As Long = 6
Private Const ZSTD_blockHeaderSize As Long = 3
Private Const ZSTD_WINDOWLOG_ABSOLUTEMIN As Long = 10
Private Const ZSTD_WINDOWLOG_MAX As Long = 30               ' ZSTD_WINDOWLOG_MAX_32
Private Const ZSTD_BLOCKSIZE_MAX As Long = 131072

' blockType_e
Private Const bt_raw As Long = 0
Private Const bt_rle As Long = 1
Private Const bt_compressed As Long = 2
Private Const bt_reserved As Long = 3

' ZSTD_frameType_e
Private Const ZSTD_frame As Long = 0
Private Const ZSTD_skippableFrame As Long = 1

' ZSTD_frameHeader. frameContentSize is a U64 in two halves (ZSTD_CONTENTSIZE_UNKNOWN: both -1); windowSize a
' Double (a U64 compared with powers of two only)
Private Type ZSTD_frameHeader
    frameContentSizeHi As Long
    frameContentSizeLo As Long
    windowSize As Double
    blockSizeMax As Long
    frameType As Long
    headerSize As Long
    dictID As Long
    checksumFlag As Long
End Type

' ZSTD_DCtx's frame-level fields (the block-level ones are zstd_decompress_block_c's)
Private fParams As ZSTD_frameHeader
Private dictID As Long              ' 0: no dictionary
Private validateChecksum As Long
Private forceIgnoreChecksum As Long ' ZSTD_d_validateChecksum (0)
Private processedCSize As Double
Private decodedSize As Double

'--------------------------------------------------------------------------
' error_private.h / zstd_common.c

' unsigned ZSTD_isError(size_t code): code > (size_t)-ZSTD_error_maxCode
Public Function ZSTD_isError(ByVal code As Long) As Boolean
    ZSTD_isError = (code < 0) And (code > -ZSTD_error_maxCode)
End Function

' ZSTD_ErrorCode ZSTD_getErrorCode(size_t code)
Public Function ZSTD_getErrorCode(ByVal code As Long) As Long
    If Not ZSTD_isError(code) Then ZSTD_getErrorCode = 0: Exit Function
    ZSTD_getErrorCode = -code
End Function

' const char* ZSTD_getErrorName(size_t code): ERR_getErrorString as built with ZSTD_STRIP_ERROR_STRINGS
Public Function ZSTD_getErrorName(ByVal code As Long) As String
    ZSTD_getErrorName = "Error strings stripped"
End Function

'--------------------------------------------------------------------------
' (VB) C's unsigned arithmetic, mem.h

Private Function zULT(ByVal a As Long, ByVal b As Long) As Boolean
    zULT = (a Xor &H80000000) < (b Xor &H80000000)
End Function

Private Function MEM_readLE16(b() As Byte, ByVal p As Long) As Long
    MEM_readLE16 = b(p) Or (CLng(b(p + 1)) * &H100&)
End Function

Private Function MEM_readLE32(b() As Byte, ByVal p As Long) As Long
    MEM_readLE32 = b(p) Or (CLng(b(p + 1)) * &H100&) Or (CLng(b(p + 2)) * &H10000) Or (CLng(b(p + 3)) * &H1000000)
End Function

' the unsigned value of a U32
Private Function U32d(ByVal x As Long) As Double
    If x < 0 Then U32d = x + 4294967296# Else U32d = x
End Function

'--------------------------------------------------------------------------
' Frame header decoding

' ZSTD_frameHeaderSize_internal: src(ip ...), srcSize available
Private Function ZSTD_frameHeaderSize_internal(src() As Byte, ByVal ip As Long, ByVal srcSize As Long) As Long
    Dim minInputSize As Long, fhd As Long, dictID_ As Long, singleSegment As Long, fcsId As Long, r As Long
    minInputSize = ZSTD_FRAMEHEADERSIZE_PREFIX
    If srcSize < minInputSize Then ZSTD_frameHeaderSize_internal = -ZSTD_error_srcSize_wrong: Exit Function

    fhd = src(ip + minInputSize - 1)
    dictID_ = fhd And 3
    singleSegment = (fhd \ 32) And 1
    fcsId = fhd \ 64
    r = minInputSize
    If singleSegment = 0 Then r = r + 1
    r = r + Choose(dictID_ + 1, 0, 1, 2, 4)     ' ZSTD_did_fieldSize
    r = r + Choose(fcsId + 1, 0, 2, 4, 8)       ' ZSTD_fcs_fieldSize
    If (singleSegment <> 0) And (fcsId = 0) Then r = r + 1
    ZSTD_frameHeaderSize_internal = r
End Function

' ZSTD_getFrameHeader_advanced: 0, zfh filled; > 0, srcSize is too small (the size wanted); or an error
Private Function ZSTD_getFrameHeader_advanced(zfh As ZSTD_frameHeader, src() As Byte, ByVal ip As Long, ByVal srcSize As Long) As Long
    Dim minInputSize As Long, fhsize As Long, fhdByte As Long, pos As Long
    Dim dictIDSizeCode As Long, checksumFlag As Long, singleSegment As Long, fcsID As Long
    Dim windowSize As Double, dictID_ As Long, fcsHi As Long, fcsLo As Long, wlByte As Long, windowLog As Long

    minInputSize = ZSTD_FRAMEHEADERSIZE_PREFIX
    zfh.frameContentSizeHi = 0: zfh.frameContentSizeLo = 0: zfh.windowSize = 0: zfh.blockSizeMax = 0
    zfh.frameType = 0: zfh.headerSize = 0: zfh.dictID = 0: zfh.checksumFlag = 0
    If srcSize < minInputSize Then ZSTD_getFrameHeader_advanced = minInputSize: Exit Function
    ' (RETURN_ERROR_IF(src==NULL, GENERIC): src is never NULL here)

    If MEM_readLE32(src, ip) <> ZSTD_MAGICNUMBER Then
        If (MEM_readLE32(src, ip) And ZSTD_MAGIC_SKIPPABLE_MASK) = ZSTD_MAGIC_SKIPPABLE_START Then
            ' skippable frame
            If srcSize < ZSTD_SKIPPABLEHEADERSIZE Then ZSTD_getFrameHeader_advanced = ZSTD_SKIPPABLEHEADERSIZE: Exit Function ' magic number + frame length
            zfh.frameContentSizeHi = 0
            zfh.frameContentSizeLo = MEM_readLE32(src, ip + ZSTD_FRAMEIDSIZE)
            zfh.frameType = ZSTD_skippableFrame
            ZSTD_getFrameHeader_advanced = 0
            Exit Function
        End If
        ZSTD_getFrameHeader_advanced = -ZSTD_error_prefix_unknown
        Exit Function
    End If

    ' ensure there is enough `srcSize` to fully read/decode frame header
    fhsize = ZSTD_frameHeaderSize_internal(src, ip, srcSize)
    If srcSize < fhsize Then ZSTD_getFrameHeader_advanced = fhsize: Exit Function
    zfh.headerSize = fhsize

    fhdByte = src(ip + minInputSize - 1)
    pos = minInputSize
    dictIDSizeCode = fhdByte And 3
    checksumFlag = (fhdByte \ 4) And 1
    singleSegment = (fhdByte \ 32) And 1
    fcsID = fhdByte \ 64
    windowSize = 0
    dictID_ = 0
    fcsHi = -1: fcsLo = -1  ' ZSTD_CONTENTSIZE_UNKNOWN
    If (fhdByte And &H8) <> 0 Then ZSTD_getFrameHeader_advanced = -ZSTD_error_frameParameter_unsupported: Exit Function ' reserved bits, must be zero

    If singleSegment = 0 Then
        wlByte = src(ip + pos)
        pos = pos + 1
        windowLog = (wlByte \ 8) + ZSTD_WINDOWLOG_ABSOLUTEMIN
        If windowLog > ZSTD_WINDOWLOG_MAX Then ZSTD_getFrameHeader_advanced = -ZSTD_error_frameParameter_windowTooLarge: Exit Function
        windowSize = 2# ^ windowLog
        windowSize = windowSize + Int(windowSize / 8#) * (wlByte And 7)
    End If
    Select Case dictIDSizeCode
    Case 1: dictID_ = src(ip + pos): pos = pos + 1
    Case 2: dictID_ = MEM_readLE16(src, ip + pos): pos = pos + 2
    Case 3: dictID_ = MEM_readLE32(src, ip + pos): pos = pos + 4
    End Select
    Select Case fcsID
    Case 0
        If singleSegment <> 0 Then fcsHi = 0: fcsLo = src(ip + pos)
    Case 1
        fcsHi = 0: fcsLo = MEM_readLE16(src, ip + pos) + 256
    Case 2
        fcsHi = 0: fcsLo = MEM_readLE32(src, ip + pos)
    Case 3
        fcsLo = MEM_readLE32(src, ip + pos): fcsHi = MEM_readLE32(src, ip + pos + 4)
    End Select
    If singleSegment <> 0 Then windowSize = U32d(fcsHi) * 4294967296# + U32d(fcsLo)

    zfh.frameType = ZSTD_frame
    zfh.frameContentSizeHi = fcsHi
    zfh.frameContentSizeLo = fcsLo
    zfh.windowSize = windowSize
    If windowSize < ZSTD_BLOCKSIZE_MAX Then zfh.blockSizeMax = windowSize Else zfh.blockSizeMax = ZSTD_BLOCKSIZE_MAX
    zfh.dictID = dictID_
    zfh.checksumFlag = checksumFlag
    ZSTD_getFrameHeader_advanced = 0
End Function

' readSkippableFrameSize: src(ip ...), srcSize available
Private Function readSkippableFrameSize(src() As Byte, ByVal ip As Long, ByVal srcSize As Long) As Long
    Dim sizeU32 As Long, skippableSize As Long
    If srcSize < ZSTD_SKIPPABLEHEADERSIZE Then readSkippableFrameSize = -ZSTD_error_srcSize_wrong: Exit Function

    sizeU32 = MEM_readLE32(src, ip + ZSTD_FRAMEIDSIZE)
    If zULT(sizeU32 + ZSTD_SKIPPABLEHEADERSIZE, sizeU32) Then readSkippableFrameSize = -ZSTD_error_frameParameter_unsupported: Exit Function
    skippableSize = ZSTD_SKIPPABLEHEADERSIZE + sizeU32
    If zULT(srcSize, skippableSize) Then readSkippableFrameSize = -ZSTD_error_srcSize_wrong: Exit Function
    readSkippableFrameSize = skippableSize
End Function

' ZSTD_decodeFrameHeader: headerSize must be the size provided by ZSTD_frameHeaderSize()
Private Function ZSTD_decodeFrameHeader(src() As Byte, ByVal ip As Long, ByVal headerSize As Long) As Long
    Dim result As Long
    result = ZSTD_getFrameHeader_advanced(fParams, src, ip, headerSize)
    If ZSTD_isError(result) Then ZSTD_decodeFrameHeader = result: Exit Function ' invalid header
    If result > 0 Then ZSTD_decodeFrameHeader = -ZSTD_error_srcSize_wrong: Exit Function ' headerSize too small

    ' (no multiple DDict references)
    If (fParams.dictID <> 0) And (dictID <> fParams.dictID) Then ZSTD_decodeFrameHeader = -ZSTD_error_dictionary_wrong: Exit Function
    If (fParams.checksumFlag <> 0) And (forceIgnoreChecksum = 0) Then validateChecksum = 1 Else validateChecksum = 0
    If validateChecksum <> 0 Then XXH64_reset 0, 0
    processedCSize = processedCSize + headerSize
    ZSTD_decodeFrameHeader = 0
End Function

'--------------------------------------------------------------------------
' Frame decoding

' ZSTD_copyRawBlock (dst is never NULL here)
Private Function ZSTD_copyRawBlock(dst() As Byte, ByVal op As Long, ByVal dstCapacity As Long, src() As Byte, ByVal ip As Long, ByVal srcSize As Long) As Long
    If zULT(dstCapacity, srcSize) Then ZSTD_copyRawBlock = -ZSTD_error_dstSize_tooSmall: Exit Function
    If srcSize > 0 Then CopyMemory dst(op), src(ip), srcSize
    ZSTD_copyRawBlock = srcSize
End Function

' ZSTD_setRleBlock
Private Function ZSTD_setRleBlock(dst() As Byte, ByVal op As Long, ByVal dstCapacity As Long, ByVal b As Byte, ByVal regenSize As Long) As Long
    Dim i As Long
    If zULT(dstCapacity, regenSize) Then ZSTD_setRleBlock = -ZSTD_error_dstSize_tooSmall: Exit Function
    For i = 0 To regenSize - 1
        dst(op + i) = b
    Next i
    ZSTD_setRleBlock = regenSize
End Function

' ZSTD_decompressFrame: dst(dstOffset .. dstOffset + dstCapacity - 1); srcPtr / srcSizePtr progress by one frame
Private Function ZSTD_decompressFrame(dst() As Byte, ByVal dstOffset As Long, ByVal dstCapacity As Long, src() As Byte, srcPtr As Long, srcSizePtr As Long) As Long
    Dim istart As Long, ip As Long, ostart As Long, oend As Long, op As Long, remainingSrcSize As Long
    Dim frameHeaderSize As Long, r As Long, decodedSize_ As Long, cBlockSize As Long
    Dim blockType As Long, lastBlock As Long, origSize As Long, checkCalc As Long, checkRead As Long

    istart = srcPtr
    ip = istart
    ostart = dstOffset
    oend = ostart + dstCapacity
    op = ostart
    remainingSrcSize = srcSizePtr

    ' check
    If remainingSrcSize < ZSTD_FRAMEHEADERSIZE_MIN + ZSTD_blockHeaderSize Then ZSTD_decompressFrame = -ZSTD_error_srcSize_wrong: Exit Function

    ' Frame Header
    frameHeaderSize = ZSTD_frameHeaderSize_internal(src, ip, ZSTD_FRAMEHEADERSIZE_PREFIX)
    If ZSTD_isError(frameHeaderSize) Then ZSTD_decompressFrame = frameHeaderSize: Exit Function
    If remainingSrcSize < frameHeaderSize + ZSTD_blockHeaderSize Then ZSTD_decompressFrame = -ZSTD_error_srcSize_wrong: Exit Function
    r = ZSTD_decodeFrameHeader(src, ip, frameHeaderSize)
    If ZSTD_isError(r) Then ZSTD_decompressFrame = r: Exit Function
    ip = ip + frameHeaderSize
    remainingSrcSize = remainingSrcSize - frameHeaderSize

    ' Loop on each block
    Do
        cBlockSize = ZSTD_getcBlockSize(src, ip, remainingSrcSize, blockType, lastBlock, origSize)
        If ZSTD_isError(cBlockSize) Then ZSTD_decompressFrame = cBlockSize: Exit Function

        ip = ip + ZSTD_blockHeaderSize
        remainingSrcSize = remainingSrcSize - ZSTD_blockHeaderSize
        If cBlockSize > remainingSrcSize Then ZSTD_decompressFrame = -ZSTD_error_srcSize_wrong: Exit Function

        Select Case blockType
        Case bt_compressed
            decodedSize_ = ZSTD_decompressBlock_internal(dst, op, oend - op, src, ip, cBlockSize, True, fParams.windowSize)
        Case bt_raw
            decodedSize_ = ZSTD_copyRawBlock(dst, op, oend - op, src, ip, cBlockSize)
        Case bt_rle
            decodedSize_ = ZSTD_setRleBlock(dst, op, oend - op, src(ip), origSize)
        Case Else ' bt_reserved
            ZSTD_decompressFrame = -ZSTD_error_corruption_detected: Exit Function ' invalid block type
        End Select

        If ZSTD_isError(decodedSize_) Then ZSTD_decompressFrame = decodedSize_: Exit Function
        If validateChecksum <> 0 Then XXH64_update dst, op, decodedSize_
        If decodedSize_ <> 0 Then op = op + decodedSize_
        ip = ip + cBlockSize
        remainingSrcSize = remainingSrcSize - cBlockSize
        If lastBlock <> 0 Then Exit Do
    Loop

    If Not ((fParams.frameContentSizeHi = -1) And (fParams.frameContentSizeLo = -1)) Then ' != ZSTD_CONTENTSIZE_UNKNOWN
        If (fParams.frameContentSizeHi <> 0) Or (fParams.frameContentSizeLo <> op - ostart) Then ZSTD_decompressFrame = -ZSTD_error_corruption_detected: Exit Function
    End If
    If fParams.checksumFlag <> 0 Then   ' Frame content checksum verification
        If remainingSrcSize < 4 Then ZSTD_decompressFrame = -ZSTD_error_checksum_wrong: Exit Function
        If forceIgnoreChecksum = 0 Then
            checkCalc = XXH64_digestLow32()
            checkRead = MEM_readLE32(src, ip)
            If checkRead <> checkCalc Then ZSTD_decompressFrame = -ZSTD_error_checksum_wrong: Exit Function
        End If
        ip = ip + 4
        remainingSrcSize = remainingSrcSize - 4
    End If
    ' (ZSTD_DCtx_trace_end: ZSTD_TRACE is 0)
    ' Allow caller to get size read
    srcPtr = ip
    srcSizePtr = remainingSrcSize
    ZSTD_decompressFrame = op - ostart
End Function

' ZSTD_decompressBegin (the frame-level fields here; the block-level ones in zstd_decompress_block_c)
Private Function ZSTD_decompressBegin() As Long
    processedCSize = 0
    decodedSize = 0
    dictID = 0
    ZSTD_vbDecompressBeginBlock
    ZSTD_decompressBegin = 0
End Function

' ZSTD_decompressMultiFrame (no dictionary)
Private Function ZSTD_decompressMultiFrame(dst() As Byte, ByVal dstOffset As Long, ByVal dstCapacity As Long, src() As Byte, ByVal ip As Long, ByVal srcSize As Long) As Long
    Dim dststart As Long, moreThan1Frame As Boolean, magicNumber As Long, skippableSize As Long, res As Long, r As Long

    dststart = dstOffset
    moreThan1Frame = False

    Do While srcSize >= ZSTD_FRAMEHEADERSIZE_PREFIX     ' ZSTD_startingInputLength(dctx->format)
        ' (ZSTD_LEGACY_SUPPORT 0)
        magicNumber = MEM_readLE32(src, ip)
        If (magicNumber And ZSTD_MAGIC_SKIPPABLE_MASK) = ZSTD_MAGIC_SKIPPABLE_START Then
            skippableSize = readSkippableFrameSize(src, ip, srcSize)
            If ZSTD_isError(skippableSize) Then ZSTD_decompressMultiFrame = skippableSize: Exit Function ' readSkippableFrameSize failed
            ip = ip + skippableSize
            srcSize = srcSize - skippableSize
            GoTo continue_
        End If

        ' this will initialize correctly with no dict if dict == NULL
        r = ZSTD_decompressBegin()      ' ZSTD_decompressBegin_usingDict(dctx, NULL, 0)
        If ZSTD_isError(r) Then ZSTD_decompressMultiFrame = r: Exit Function
        ZSTD_checkContinuity dstOffset, dstCapacity

        res = ZSTD_decompressFrame(dst, dstOffset, dstCapacity, src, ip, srcSize)
        If (ZSTD_getErrorCode(res) = ZSTD_error_prefix_unknown) And moreThan1Frame Then
            ' At least one frame successfully completed, but following bytes are garbage: it's more likely to be a
            ' srcSize error, specifying more input bytes than size of frame(s).
            ZSTD_decompressMultiFrame = -ZSTD_error_srcSize_wrong
            Exit Function
        End If
        If ZSTD_isError(res) Then ZSTD_decompressMultiFrame = res: Exit Function
        If res <> 0 Then dstOffset = dstOffset + res
        dstCapacity = dstCapacity - res
        moreThan1Frame = True
continue_:
    Loop

    If srcSize <> 0 Then ZSTD_decompressMultiFrame = -ZSTD_error_srcSize_wrong: Exit Function ' input not entirely consumed

    ZSTD_decompressMultiFrame = dstOffset - dststart
End Function

' size_t ZSTD_decompress(void* dst, size_t dstCapacity, const void* src, size_t compressedSize): dst is
' dst(dstOffset .. dstOffset + dstCapacity - 1), src is src(srcOffset .. srcOffset + compressedSize - 1). Returns
' the number of bytes decompressed into dst (<= dstCapacity), or an error (ZSTD_isError).
Public Function ZSTD_decompress(dst() As Byte, ByVal dstOffset As Long, ByVal dstCapacity As Long, src() As Byte, ByVal srcOffset As Long, ByVal compressedSize As Long) As Long
    ' (ZSTD_HEAPMODE: ZSTD_createDCtx, ZSTD_initDCtx_internal: no dictionary, checksums validated)
    forceIgnoreChecksum = 0
    ' ZSTD_decompressDCtx -> ZSTD_decompress_usingDDict(dctx, ..., ZSTD_getDDict(dctx) = NULL)
    ZSTD_decompress = ZSTD_decompressMultiFrame(dst, dstOffset, dstCapacity, src, srcOffset, compressedSize)
End Function
