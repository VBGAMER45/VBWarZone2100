VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZlibCompressionAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/zlib_compression_adapter.cpp, zlib_compression_adapter.h and compression_adapter.h (class
' ICompressionAdapter: this is its one implementation), over the zlib streams of src/shared/miniz (cluster J's
' Contract A: the z_stream's next_in / next_out are offsets into the arrays passed on each call).
' std::vector<uint8_t> deflateOutBuf_ / inflateInBuf_ are arrays with a size beside (their capacity is the array's).
' The C++ hands out the vectors by reference (compressionOutBuffer, decompressionInBuffer); VB cannot, so
' IClientConnection's two uses are methods here: recvIntoDecompressionInBuffer (readNoInt's resize + recvImpl
' into it) and appendCompressionOutBuffer / clearCompressionOutBuffer (flush's append to the write queue, clear).
' The asserts print zlib's macros as written (4.7.0's ASSERT prints them expanded: "ret == 0"; check_debug_lines
' cannot see zlib.h).
Option Explicit

Private deflateOutBuf_() As Byte
Private deflateOutBufSize_ As Long
Private deflateOutBufCap_ As Long               ' (VB) the array's length (0: unallocated)
Private inflateInBuf_() As Byte
Private inflateInBufSize_ As Long
Private inflateInBufCap_ As Long
Private deflateStream_ As z_stream
Private inflateStream_ As z_stream
Private inflateNeedInput_ As Boolean
Private noInput_() As Byte                      ' (VB) next_in = nullptr: an unallocated array (Contract A)
Private destroyed_ As Boolean

' ZlibCompressionAdapter(): memset both streams to 0 (a new UDT is all zero)

' ~ZlibCompressionAdapter()
Friend Sub destroy_()
    If destroyed_ Then Exit Sub
    destroyed_ = True
    deflateEnd deflateStream_
    deflateEnd inflateStream_                   ' (sic: 4.7.0 ends its inflate stream with deflateEnd)
End Sub

Private Sub Class_Terminate()
    destroy_
End Sub

' (VB) std::vector<uint8_t>::resize for deflateOutBuf_ (the array keeps its capacity)
Private Sub deflateOutBufResize(ByVal n As Long)
    If n > deflateOutBufCap_ Then
        If deflateOutBufCap_ = 0 Then
            ReDim deflateOutBuf_(0 To n - 1)
        Else
            ReDim Preserve deflateOutBuf_(0 To n - 1)
        End If
        deflateOutBufCap_ = n
    End If
    deflateOutBufSize_ = n
End Sub

' virtual net::result<void> initialize() override: the result is ec
Friend Function initialize() As Long
    Dim ret As Long
    ' Init deflate stream (zalloc, zfree, opaque: Z_NULL)
    ret = deflateInit(deflateStream_, 6)
    If Not (ret = Z_OK) Then ASSERT_FAILURE "ret == Z_OK", "lib/netplay/zlib_compression_adapter.cpp:46", "initialize", "deflateInit failed! Sockets won't work."
    If ret <> Z_OK Then
        initialize = make_zlib_error_code(ret)
        Exit Function
    End If

    ' Init inflate stream
    inflateStream_.avail_in = 0
    inflateStream_.next_in = 0                  ' Z_NULL
    ret = inflateInit(inflateStream_)
    If Not (ret = Z_OK) Then ASSERT_FAILURE "ret == Z_OK", "lib/netplay/zlib_compression_adapter.cpp:59", "initialize", "deflateInit failed! Sockets won't work."
    If ret <> Z_OK Then
        initialize = make_zlib_error_code(ret)
        Exit Function
    End If

    inflateNeedInput_ = True

    initialize = 0
End Function

' void resetCompressionStreamInput(const void* src, size_t size) (zlib >= 1.2.5.2's ZLIB_CONST branch)
Private Sub resetCompressionStreamInput(ByVal ofs As Long, ByVal size As Long)
    deflateStream_.next_in = ofs
    deflateStream_.avail_in = size
End Sub

' virtual net::result<void> compress(const void* src, size_t size) override: the result is ec
Friend Function compress(src() As Byte, ByVal ofs As Long, ByVal size As Long, ec As Long) As Long
    Dim alreadyHave As Long, ret As Long
    ec = 0
    resetCompressionStreamInput ofs, size
    Do
        alreadyHave = deflateOutBufSize_
        deflateOutBufResize alreadyHave + size + 20      ' A bit more than size should be enough to always do everything in one go.
        deflateStream_.next_out = alreadyHave
        deflateStream_.avail_out = deflateOutBufSize_ - alreadyHave

        ret = deflate(deflateStream_, src, deflateOutBuf_, Z_NO_FLUSH)
        If Not (ret <> Z_STREAM_ERROR) Then ASSERT_FAILURE "ret != Z_STREAM_ERROR", "lib/netplay/zlib_compression_adapter.cpp:111", "compress", "zlib compression failed!"

        ' Remove unused part of buffer.
        deflateOutBufResize deflateOutBufSize_ - deflateStream_.avail_out
    Loop While deflateStream_.avail_out = 0

    If Not (deflateStream_.avail_in = 0) Then ASSERT_FAILURE "deflateStream_.avail_in == 0", "lib/netplay/zlib_compression_adapter.cpp:117", "compress", "zlib didn't compress everything!"

    compress = 0
End Function

' virtual net::result<void> flushCompressionStream() override: the result is ec
Friend Function flushCompressionStream() As Long
    Dim alreadyHave As Long, ret As Long
    ' Flush data out of zlib compression state.
    Do
        deflateStream_.next_in = 0                ' (Bytef*)nullptr
        deflateStream_.avail_in = 0
        alreadyHave = deflateOutBufSize_
        deflateOutBufResize alreadyHave + 1000   ' 100 bytes would probably be enough to flush the rest in one go.
        deflateStream_.next_out = alreadyHave
        deflateStream_.avail_out = deflateOutBufSize_ - alreadyHave

        ret = deflate(deflateStream_, noInput_, deflateOutBuf_, Z_PARTIAL_FLUSH)
        If Not (ret <> Z_STREAM_ERROR) Then ASSERT_FAILURE "ret != Z_STREAM_ERROR", "lib/netplay/zlib_compression_adapter.cpp:135", "flushCompressionStream", "zlib compression failed!"

        ' Remove unused part of buffer.
        deflateOutBufResize deflateOutBufSize_ - deflateStream_.avail_out
    Loop While deflateStream_.avail_out = 0

    flushCompressionStream = 0
End Function

' virtual std::vector<uint8_t>& compressionOutBuffer() override: its size
Friend Function compressionOutBufferSize() As Long
    compressionOutBufferSize = deflateOutBufSize_
End Function

' (VB) IClientConnection::flush's pwm_->append(this, [&compressionBuf](q) { q.insert(q.end(), compressionBuf...) })
Friend Sub appendCompressionOutBuffer(ByVal pwm As PendingWritesManager, ByVal conn As IClientConnection)
    pwm.append conn, deflateOutBuf_, 0, deflateOutBufSize_
End Sub

' (VB) IClientConnection::flush's compressionBuf.clear()
Friend Sub clearCompressionOutBuffer()
    deflateOutBufSize_ = 0
End Sub

' virtual net::result<void> decompress(void* dst, size_t size) override: the result is ec
Friend Function decompress(dst() As Byte, ByVal ofs As Long, ByVal size As Long, ec As Long) As Long
    Dim ret As Long, err_ As String
    ec = 0
    resetDecompressionStreamOutput ofs, size

    ret = inflate(inflateStream_, inflateInBuf_, dst, Z_NO_FLUSH)
    If Not (ret <> Z_STREAM_ERROR) Then ASSERT_FAILURE "ret != Z_STREAM_ERROR", "lib/netplay/zlib_compression_adapter.cpp:149", "decompress", "zlib inflate not working!"
    Select Case ret
    Case Z_NEED_DICT: err_ = "Z_NEED_DICT"
    Case Z_DATA_ERROR: err_ = "Z_DATA_ERROR"
    Case Z_MEM_ERROR: err_ = "Z_MEM_ERROR"
    End Select
    If ret = Z_STREAM_ERROR Then
        If Len(err_) <> 0 Then
            debug_ 161, LOG_ERROR, "decompress", "Couldn't decompress data from socket. zlib error %s", err_
        Else
            debug_ 165, LOG_ERROR, "decompress", "Couldn't decompress data from socket. Unknown zlib error"
        End If
        ec = make_zlib_error_code(ret)
        Exit Function
    End If
    decompress = 0
End Function

' (VB) IClientConnection::readNoInt's decompressInBuf.resize(size) and recvImpl(decompressInBuf.data(), size)
Friend Function recvIntoDecompressionInBuffer(ByVal conn As IClientConnection, ByVal size As Long, ec As Long) As Long
    If size > inflateInBufCap_ Then
        If inflateInBufCap_ = 0 Then
            ReDim inflateInBuf_(0 To size - 1)
        Else
            ReDim Preserve inflateInBuf_(0 To size - 1)
        End If
        inflateInBufCap_ = size
    End If
    inflateInBufSize_ = size
    recvIntoDecompressionInBuffer = conn.recvImpl(inflateInBuf_, 0, size, ec)
End Function

' virtual size_t availableSpaceToDecompress() const override
Friend Function availableSpaceToDecompress() As Long
    availableSpaceToDecompress = inflateStream_.avail_out
End Function

' virtual bool decompressionStreamConsumedAllInput() const override
Friend Function decompressionStreamConsumedAllInput() As Boolean
    decompressionStreamConsumedAllInput = (inflateStream_.avail_in = 0)
End Function

' virtual bool decompressionNeedInput() const override
Friend Function decompressionNeedInput() As Boolean
    decompressionNeedInput = inflateNeedInput_
End Function

' virtual void setDecompressionNeedInput(bool needInput) override
Friend Sub setDecompressionNeedInput(ByVal needInput As Boolean)
    inflateNeedInput_ = needInput
End Sub

' virtual void resetDecompressionStreamInputSize(size_t size) override
Friend Sub resetDecompressionStreamInputSize(ByVal size As Long)
    inflateStream_.next_in = 0                  ' inflateInBuf_.data()
    inflateStream_.avail_in = size
End Sub

' void resetDecompressionStreamOutput(void* buf, size_t size)
Private Sub resetDecompressionStreamOutput(ByVal ofs As Long, ByVal size As Long)
    inflateStream_.next_out = ofs
    inflateStream_.avail_out = size
End Sub
