VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IClientConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/client_connection.cpp, client_connection.h: class IClientConnection (cluster J's Contract C).
' PublicNotCreatable: the providers and listen sockets make it, netplay and the joining screen hold it (an
' IClientConnection* is this object, Nothing = nullptr). The derived class is WzNet's private TCPClientConnection,
' which this object holds (the base's virtual methods are the TCP ones: no other backend is ported); C++'s
' `delete conn` is destroy_ (the derived destructor closes the socket) and leaves the object invalid.
' net::result<T> returns T and sets ec (error_categories_c); a buffer is buf(), ofs, size; a size_t* is a
' ByRef Long the caller always passes. No mutexes (no threads): PendingWritesManager's lock is nothing.
' This file is not one that includes tcp/sock_error.h: EINVAL, EBADF and ETIMEDOUT are MSVC's errno values.
Option Explicit

Private Const EINVAL As Long = 22
Private Const EBADF As Long = 9
Private Const ETIMEDOUT As Long = 138

Private tcp_ As TCPClientConnection               ' (VB) the derived object
Private connProvider_ As WzConnectionProvider     ' std::weak_ptr<WzConnectionProvider> (see connProviderLock)
Private compressionProviderSet_ As Boolean        ' WzCompressionProvider* compressionProvider_ (the singleton)
Private pwm_ As PendingWritesManager
Private writeErrorSet_ As Boolean
Private writeErrorCodeHas_ As Boolean             ' optional<std::error_code> writeErrorCode_
Private writeErrorCode_ As Long
Private compressionAdapter_ As ZlibCompressionAdapter
Private readAllDescriptorSet_ As SelectDescriptorSet
Private deleteLater_ As Boolean
Private isCompressed_ As Boolean
Private destroyed_ As Boolean                     ' (VB) the C++ delete happened

' IClientConnection(WzConnectionProvider& connProvider, WzCompressionProvider& compressionProvider,
'     PendingWritesManager& pwm), as the base of the TCPClientConnection the derived construct makes
Friend Sub construct(ByVal connProvider As WzConnectionProvider, ByVal pwm As PendingWritesManager, ByVal derived As TCPClientConnection)
    ' selfConnList_({ this }) is made when used (selfConnList): a list holding this object would keep it alive
    Set connProvider_ = connProvider
    compressionProviderSet_ = True
    Set pwm_ = pwm
    Set readAllDescriptorSet_ = connProvider.newDescriptorSet(0)    ' PollEventType::READABLE
    Set tcp_ = derived
End Sub

' (VB) const std::vector<IClientConnection*> selfConnList_
Friend Function selfConnList() As Collection
    Set selfConnList = New Collection
    selfConnList.Add Me
End Function

' (VB) dynamic_cast<TCPClientConnection*>(this)
Friend Function tcpImpl() As TCPClientConnection
    Set tcpImpl = tcp_
End Function

' net::result<ssize_t> readAll(void* buf, size_t size, unsigned timeout)
Public Function readAll(buf() As Byte, ByVal ofs As Long, ByVal size As Long, ByVal timeout As Long, ec As Long) As Long
    Dim received As Long, ret As Long, msg As String, recvRes As Long, ecRecv As Long
    ec = 0
    If Not (Not isCompressed()) Then ASSERT_FAILURE "!isCompressed()", "lib/netplay/client_connection.cpp:40", "readAll", "readAll on compressed sockets not implemented.": ec = make_network_error_code(EINVAL): Exit Function

    If Not isValid() Then
        debug_ 45, LOG_ERROR, "readAll", "IClientConnection::readAll: Invalid socket (%p) (error: EBADF)", ObjPtr(Me)
        ec = make_network_error_code(EBADF)
        Exit Function
    End If

    received = 0
    Do While ULT(received, size)
        ' If a timeout is set, wait for that amount of time for data to arrive (or abort)
        If timeout <> 0 Then
            ret = checkConnectionsReadable(selfConnList(), readAllDescriptorSet_, timeout, ec)
            If ec <> 0 Then
                msg = net_errorMessage(ec)
                debug_ 59, LOG_NET, "readAll", "socket (%p) error: %s.", ObjPtr(Me), msg
                Exit Function
            ElseIf ret = 0 Then
                debug_ 64, LOG_NET, "readAll", "socket (%p) has timed out.", ObjPtr(Me)
                ec = make_network_error_code(ETIMEDOUT)
                Exit Function
            End If
            If Not (readReady()) Then ASSERT_FAILURE "readReady()", "lib/netplay/client_connection.cpp:67", "readAll", "Socket should be ready for reading"
        End If
        recvRes = recvImpl(buf, ofs + received, size - received, ecRecv)
        setReadReady False
        If ecRecv <> 0 Then
            If net_errorEqualsErrc(ecRecv, errc_resource_unavailable_try_again) Or _
               net_errorEqualsErrc(ecRecv, errc_operation_would_block) Or _
               net_errorEqualsErrc(ecRecv, errc_interrupted) Then           ' EWOULDBLOCK, EAGAIN, EINTR
                GoTo continue_
            End If
            ec = ecRecv
            readAll = recvRes
            Exit Function
        End If
        received = received + recvRes
continue_:
    Loop

    readAll = received
End Function

' net::result<ssize_t> readNoInt(void* buf, size_t max_size, size_t* rawByteCount)
Public Function readNoInt(buf() As Byte, ByVal ofs As Long, ByVal max_size As Long, rawByteCount As Long, ec As Long) As Long
    Dim recvRes As Long, received As Long
    ec = 0
    rawByteCount = 0

    If Not isValid() Then
        debug_ 97, LOG_ERROR, "readNoInt", "IClientConnection::readNoInt: Invalid socket"
        ec = make_network_error_code(EBADF)
        Exit Function
    End If

    If isCompressed() Then
        If compressionAdapter_.decompressionNeedInput() Then
            ' No input data, read some.
            ' (VB) decompressInBuf.resize(max_size + 1000); recvImpl(decompressInBuf.data(), decompressInBuf.size())
            recvRes = compressionAdapter_.recvIntoDecompressionInBuffer(Me, max_size + 1000, ec)
            If ec <> 0 Then
                readNoInt = recvRes
                Exit Function
            End If
            received = recvRes
            compressionAdapter_.resetDecompressionStreamInputSize received
            rawByteCount = received
            compressionAdapter_.setDecompressionNeedInput False
        End If

        compressionAdapter_.decompress buf, ofs, max_size, ec
        If ec <> 0 Then Exit Function

        If compressionAdapter_.availableSpaceToDecompress() <> 0 Then
            compressionAdapter_.setDecompressionNeedInput True
            If Not (compressionAdapter_.decompressionStreamConsumedAllInput()) Then ASSERT_FAILURE "compressionAdapter_->decompressionStreamConsumedAllInput()", "lib/netplay/client_connection.cpp:132", "readNoInt", "Compression algorithm impl not consuming all input!"
        End If

        readNoInt = max_size - compressionAdapter_.availableSpaceToDecompress()   ' Got some data, return how much.
        Exit Function
    End If

    recvRes = recvImpl(buf, ofs, max_size, ec)
    If ec <> 0 Then
        readNoInt = recvRes
        Exit Function
    End If

    setReadReady False
    received = recvRes
    rawByteCount = received

    readNoInt = received
End Function

' net::result<ssize_t> writeAll(const void* buf, size_t size, size_t* rawByteCount)
Public Function writeAll(buf() As Byte, ByVal ofs As Long, ByVal size As Long, rawByteCount As Long, ec As Long) As Long
    Dim writeErrHas As Long, writeErr As Long
    ec = 0
    If Not isValid() Then
        debug_ 158, LOG_ERROR, "writeAll", "IClientConnection::writeAll: Invalid socket (EBADF)"
        ec = make_network_error_code(EBADF)
        Exit Function
    End If

    writeErr = writeErrorCode(writeErrHas)
    If writeErrHas Then
        ec = writeErr
        Exit Function
    End If

    rawByteCount = 0

    If size = 0 Then
        writeAll = 0
        Exit Function
    End If

    If Not isCompressed() Then
        ' pwm_->append(this, [buf, size](writeQueue) { writeQueue.insert(writeQueue.end(), buf, buf + size); })
        pwm_.append Me, buf, ofs, size
        rawByteCount = size
    Else
        compressionAdapter_.compress buf, ofs, size, ec
        If ec <> 0 Then
            ' compress failed?
            debug_ 195, LOG_ERROR, "writeAll", "IClientConnection::writeAll: compress failed?"
            Exit Function
        End If
    End If

    writeAll = size
End Function

' net::result<void> flush(size_t* rawByteCount): the result is ec
Public Function flush(rawByteCount As Long) As Long
    Dim writeErrHas As Long, writeErr As Long, errMsg As String, ec As Long, n As Long
    If Not isValid() Then
        debug_ 207, LOG_ERROR, "flush", "IClientConnection::flush: Invalid socket (EBADF)"
        flush = make_network_error_code(EBADF)
        Exit Function
    End If

    rawByteCount = 0
    If Not isCompressed() Then
        flush = 0                                ' Not compressed, so don't mess with compression.
        Exit Function
    End If

    writeErr = writeErrorCode(writeErrHas)
    If writeErrHas Then
        errMsg = net_errorMessage(writeErr)
        debug_ 224, LOG_ERROR, "flush", "Socket write error encountered in flush: %s", errMsg
        flush = writeErr
        Exit Function
    End If

    ec = compressionAdapter_.flushCompressionStream()
    If ec <> 0 Then
        ' flushCompressionStream failed
        debug_ 232, LOG_ERROR, "flush", "Socket write error encountered flushing compression stream"
        flush = ec
        Exit Function
    End If

    n = compressionAdapter_.compressionOutBufferSize()
    If n = 0 Then
        flush = 0                                ' No data to flush out.
        Exit Function
    End If

    ' pwm_->append(this, [&compressionBuf](writeQueue) { writeQueue.insert(writeQueue.end(), compressionBuf...); })
    compressionAdapter_.appendCompressionOutBuffer pwm_, Me
    ' Data sent, don't send again.
    rawByteCount = n
    compressionAdapter_.clearCompressionOutBuffer
    flush = 0
End Function

' void enableCompression()
Public Sub enableCompression()
    Dim ec As Long, errMsg As String
    If isCompressed_ Then
        Exit Sub                                 ' Nothing to do.
    End If

    If Not (compressionProviderSet_) Then ASSERT_FAILURE "compressionProvider_ != nullptr", "lib/netplay/client_connection.cpp:263", "enableCompression", "Invalid compression provider": Exit Sub

    ' pwm_->executeUnderLock([this] { ... })
    Set compressionAdapter_ = WzCompressionProvider_newCompressionAdapter()
    ec = compressionAdapter_.initialize()
    If ec <> 0 Then
        errMsg = net_errorMessage(ec)
        debug_ 272, LOG_NET, "enableCompression", "Failed to initialize compression algorithms. Sockets won't work properly! Detailed error message: %s", errMsg
        Exit Sub
    End If
    isCompressed_ = True
End Sub

' bool isCompressed() const
Public Function isCompressed() As Boolean
    isCompressed = isCompressed_
End Function

' ICompressionAdapter& compressionAdapter()
Friend Function compressionAdapter() As ZlibCompressionAdapter
    Set compressionAdapter = compressionAdapter_
End Function

' void close() (close_: Close is a VB keyword)
Public Sub close_()
    pwm_.safeDispose Me
End Sub

' bool deleteLaterRequested() const
Public Function deleteLaterRequested() As Boolean
    deleteLaterRequested = deleteLater_
End Function

' void requestDeleteLater()
Public Sub requestDeleteLater()
    deleteLater_ = True
End Sub

' inline optional<std::error_code> writeErrorCode() const: has is set when it holds a code
Friend Function writeErrorCode(has As Long) As Long
    has = 0
    If Not writeErrorSet_ Then Exit Function
    If writeErrorCodeHas_ Then
        has = 1
        writeErrorCode = writeErrorCode_
    End If
End Function

' void setWriteErrorCode(optional<std::error_code> ec)
Friend Sub setWriteErrorCode(ByVal has As Boolean, ByVal ec As Long)
    writeErrorCodeHas_ = has
    writeErrorCode_ = ec
    writeErrorSet_ = True
End Sub

' (VB) std::weak_ptr<WzConnectionProvider> connProvider_'s lock(): the provider, or Nothing once the registry has let
' go of it (ConnectionProviderRegistry::Shutdown: the shared_ptr's owners are gone in 4.7.0's shutdown order)
Friend Function connProviderLock() As WzConnectionProvider
    If Not connProvider_ Is Nothing Then
        If Not connProvider_.isShutdownByRegistry() Then Set connProviderLock = connProvider_
    End If
End Function

' (VB) delete this: the derived destructor (~TCPClientConnection closes the socket), then the base's members
' (~ZlibCompressionAdapter ends its streams)
Friend Sub destroy_()
    If destroyed_ Then Exit Sub
    destroyed_ = True
    If Not tcp_ Is Nothing Then tcp_.destroy_
    If Not compressionAdapter_ Is Nothing Then compressionAdapter_.destroy_
    Set compressionAdapter_ = Nothing
    Set readAllDescriptorSet_ = Nothing
End Sub

' --- the virtual methods: TCPClientConnection's ---

' virtual net::result<ssize_t> sendImpl(const std::vector<uint8_t>& data) = 0
Friend Function sendImpl(data() As Byte, ByVal size As Long, ec As Long) As Long
    sendImpl = tcp_.sendImpl(data, size, ec)
End Function

' virtual net::result<ssize_t> recvImpl(char* dst, size_t maxSize) = 0
Friend Function recvImpl(dst() As Byte, ByVal ofs As Long, ByVal maxSize As Long, ec As Long) As Long
    recvImpl = tcp_.recvImpl(dst, ofs, maxSize, ec)
End Function

' virtual void setReadReady(bool ready) = 0
Public Sub setReadReady(ByVal ready As Boolean)
    tcp_.setReadReady ready
End Sub

' virtual bool readReady() const = 0
Public Function readReady() As Boolean
    readReady = tcp_.readReady()
End Function

' virtual void useNagleAlgorithm(bool enable) = 0
Public Sub useNagleAlgorithm(ByVal enable As Boolean)
    tcp_.useNagleAlgorithm enable
End Sub

' virtual std::string textAddress() const = 0
Public Function textAddress() As String
    textAddress = tcp_.textAddress()
End Function

' virtual bool isValid() const = 0
Public Function isValid() As Boolean
    If destroyed_ Then Exit Function             ' (VB) a deleted connection
    isValid = tcp_.isValid()
End Function

' virtual net::result<void> connectionStatus() const = 0: the result is ec
Public Function connectionStatus() As Long
    connectionStatus = tcp_.connectionStatus(Me)
End Function

' virtual void setConnectedTimeout(std::chrono::milliseconds timeout) = 0
Public Sub setConnectedTimeout(ByVal timeoutMs As Long)
    tcp_.setConnectedTimeout timeoutMs
End Sub
