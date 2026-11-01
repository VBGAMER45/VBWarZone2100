Attribute VB_Name = "miniz_c"
' miniz.c v1.11's zlib-style API's but the compressor's and decompressor's own (miniz_tdefl_c, miniz_tinfl_c):
' mz_adler32 and mz_version, the default allocator's free (def_free_func: the stream's state is an entry of
' either module's pool, by its handle), and zlib's function names (MINIZ_NO_ZLIB_COMPATIBLE_NAMES is not set:
' "#define deflateInit mz_deflateInit" and so on) as functions with zlib's parameter order, the buffers last but
' flush (zlib's deflateInit, deflate, deflateEnd, inflateInit, inflate, inflateEnd; the rest of miniz's list is
' not needed). mz_crc32 (the zip writer's) is not ported. (VB) An empty input or output is an unallocated array
' (the C's NULL pointer) with avail 0.
Option Explicit

Private Const ADLER_BLOCK As Long = 3800                        ' (VB) the C's 5552: a VB Long holds 31 bits where the C's mz_uint32 holds 32

' (VB) ptr(ofs .. ofs + buf_len - 1); ptr unallocated: the C's NULL
Public Function mz_adler32(ByVal adler As Long, ptr() As Byte, ByVal ofs As Long, ByVal buf_len As Long) As Long
    Dim i As Long, s1 As Long, s2 As Long, block_len As Long, p As Long
    s1 = adler And &HFFFF&: s2 = SHRU(adler, 16): block_len = buf_len Mod ADLER_BLOCK
    If Not mz_bufAllocated(ptr) Then
        mz_adler32 = MZ_ADLER32_INIT
        Exit Function
    End If
    p = ofs
    Do While buf_len <> 0
        i = 0
        Do While i + 7 < block_len
            s1 = s1 + ptr(p): s2 = s2 + s1: s1 = s1 + ptr(p + 1): s2 = s2 + s1: s1 = s1 + ptr(p + 2): s2 = s2 + s1: s1 = s1 + ptr(p + 3): s2 = s2 + s1
            s1 = s1 + ptr(p + 4): s2 = s2 + s1: s1 = s1 + ptr(p + 5): s2 = s2 + s1: s1 = s1 + ptr(p + 6): s2 = s2 + s1: s1 = s1 + ptr(p + 7): s2 = s2 + s1
            i = i + 8: p = p + 8
        Loop
        Do While i < block_len
            s1 = s1 + ptr(p): s2 = s2 + s1
            i = i + 1: p = p + 1
        Loop
        s1 = s1 Mod 65521: s2 = s2 Mod 65521: buf_len = buf_len - block_len: block_len = ADLER_BLOCK
    Loop
    mz_adler32 = SHL(s2, 16) + s1
End Function

Public Function mz_version() As String
    mz_version = "9.1.11"                       ' MZ_VERSION
End Function

' static void def_free_func(void *opaque, void *address): (VB) the address is a stream's state handle, of either pool
Public Sub def_free_func(ByVal address As Long)
    If address > MZ_DEFLATE_STATE_BASE Then
        miniz_deflateStateFree address
    Else
        miniz_inflateStateFree address
    End If
End Sub

' (VB) whether b has elements (an unallocated array is the C's NULL buffer)
Public Function mz_bufAllocated(b() As Byte) As Boolean
    Dim n As Long
    On Error GoTo none
    n = UBound(b)
    mz_bufAllocated = True
none:
End Function

' ------------------- zlib-compatible names

' #define deflateInit mz_deflateInit
Public Function deflateInit(strm As z_stream, ByVal level As Long) As Long
    deflateInit = mz_deflateInit(strm, level)
End Function

' #define deflate mz_deflate (VB) strm's next_in / next_out: offsets into inBuf / outBuf
Public Function deflate(strm As z_stream, inBuf() As Byte, outBuf() As Byte, ByVal flush As Long) As Long
    deflate = mz_deflate(strm, flush, inBuf, outBuf)
End Function

' #define deflateEnd mz_deflateEnd
Public Function deflateEnd(strm As z_stream) As Long
    deflateEnd = mz_deflateEnd(strm)
End Function

' #define inflateInit mz_inflateInit
Public Function inflateInit(strm As z_stream) As Long
    inflateInit = mz_inflateInit(strm)
End Function

' #define inflate mz_inflate (VB) strm's next_in / next_out: offsets into inBuf / outBuf.
' (VB) zlib's inflate decompresses until the input or the output space runs out (or the stream ends, or an
' error); miniz's mz_inflate returns Z_OK after copying out what its 32 KB dictionary held, input and room
' still left. 4.7.0 links zlib, and its ZlibCompressionAdapter relies on zlib's: with room left after a call it
' takes all input as consumed and refills it (found by fork JT: every compressed network stream broke after
' about 32 KB). So zlib's name calls mz_inflate again while both are left: each such return has handed over
' pending bytes, so it makes progress. mz_inflate itself stays miniz's (PhysFS's zip reader loops on it).
Public Function inflate(strm As z_stream, inBuf() As Byte, outBuf() As Byte, ByVal flush As Long) As Long
    Dim ret As Long
    Do
        ret = mz_inflate(strm, flush, inBuf, outBuf)
    Loop While ret = MZ_OK And strm.avail_in > 0 And strm.avail_out > 0
    inflate = ret
End Function

' #define inflateEnd mz_inflateEnd
Public Function inflateEnd(strm As z_stream) As Long
    inflateEnd = mz_inflateEnd(strm)
End Function
