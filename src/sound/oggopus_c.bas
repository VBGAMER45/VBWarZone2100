Attribute VB_Name = "oggopus_c"
' lib/sound/oggopus.cpp: the opusfile callbacks over a PhysFS file and over a memory buffer, and WZOpusDecoder's
' static functions (fromFilename, fromMemoryBuffer, openOpusInternal: a VB class has no static function; the
' class is WZOpusDecoder.cls). The callbacks' function pointers are OPFN_* numbers that opusfile_stream_c's
' dispatch calls (so the callbacks are Public); their void *datasource is the Variant holding the
' WzAudioDataPhysFSHandle or WZAudioDataMemoryBuffer object (codecs.h's classes). As in 4.7.0, a seek from
' SEEK_END goes to fileSize - 1 - offset, so opusfile takes the stream to end one byte early and never reads
' a track's last Ogg page (its pages end before that byte).
Option Explicit

' (VB) the function pointer numbers of the callbacks (opusfile_stream_c's op_vbCall*)
Public Const OPFN_wz_opus_physfs_read As Long = 101
Public Const OPFN_wz_opus_physfs_seek As Long = 102
Public Const OPFN_wz_opus_physfs_tell As Long = 103
Public Const OPFN_wz_opus_membuf_read As Long = 104
Public Const OPFN_wz_opus_membuf_seek As Long = 105
Public Const OPFN_wz_opus_membuf_tell As Long = 106

' MARK: - Opus PHYSFS callbacks

Public Function wz_opus_physfs_read(datasource As Variant, ptr_() As Byte, ByVal ptrOff As Long, ByVal nbytes_ As Long) As Long
    Dim fileHandle As Long, h As WzAudioDataPhysFSHandle, didread As Long
    If Not (Not datasource Is Nothing) Then ASSERT_FAILURE "datasource != nullptr", "lib/sound/oggopus.cpp:33", "wz_opus_physfs_read", "NULL decoder passed!"

    Set h = datasource
    fileHandle = h.fileHandle()
    If Not (fileHandle <> 0) Then ASSERT_FAILURE "fileHandle != nullptr", "lib/sound/oggopus.cpp:36", "wz_opus_physfs_read", "Bad PhysicsFS file handle passed in"

    ' (a negative _nbytes converted to size_t is above PHYSFS_uint32's max)
    If Not (nbytes_ >= 0) Then ASSERT_FAILURE "_nbytes <= static_cast<size_t>(std::numeric_limits<PHYSFS_uint32>::max())", "lib/sound/oggopus.cpp:38", "wz_opus_physfs_read", "readLen (%i) exceeds PHYSFS_uint32::max", nbytes_
    didread = WZ_PHYSFS_readBytes(fileHandle, ptr_, ptrOff, nbytes_)
    wz_opus_physfs_read = didread
End Function

Public Function wz_opus_physfs_seek(datasource As Variant, ByVal offset_ As Double, ByVal whence_ As Long) As Long
    Dim fileHandle As Long, newPos As Double, h As WzAudioDataPhysFSHandle, curPos As Double, fileSize As Double
    If Not (Not datasource Is Nothing) Then ASSERT_FAILURE "datasource != nullptr", "lib/sound/oggopus.cpp:47", "wz_opus_physfs_seek", "NULL decoder passed!"
    Set h = datasource
    fileHandle = h.fileHandle()
    If Not (fileHandle <> 0) Then ASSERT_FAILURE "fileHandle != nullptr", "lib/sound/oggopus.cpp:49", "wz_opus_physfs_seek", "Bad PhysicsFS file handle passed in"

    Select Case whence_
    ' Seek to absolute position
    Case SEEK_SET
        newPos = offset_

    ' Seek `offset` ahead
    Case SEEK_CUR
        curPos = PHYSFS_tell(fileHandle)
        If curPos = -1 Then
            wz_opus_physfs_seek = -1
            Exit Function
        End If

        newPos = curPos + offset_

    ' Seek backwards from the end of the file
    Case SEEK_END
        fileSize = PHYSFS_fileLength(fileHandle)
        If fileSize = -1 Then
            wz_opus_physfs_seek = -1
            Exit Function
        End If

        newPos = fileSize - 1 - offset_

    ' unrecognized seek instruction
    Case Else
        ' indicate failure
        wz_opus_physfs_seek = -1
        Exit Function
    End Select

    ' PHYSFS_seek return value of non-zero means success
    If PHYSFS_seek(fileHandle, newPos) <> 0 Then
     ' success is zero! opposite of what PHYSFS thinks
        wz_opus_physfs_seek = 0
    Else
        ' indicate failure
        wz_opus_physfs_seek = -1    ' failure
    End If
End Function

Public Function wz_opus_physfs_tell(datasource As Variant) As Double
    Dim fileHandle As Long, h As WzAudioDataPhysFSHandle, out As Double
    If Not (Not datasource Is Nothing) Then ASSERT_FAILURE "datasource != nullptr", "lib/sound/oggopus.cpp:106", "wz_opus_physfs_tell", "NULL decoder passed!"

    Set h = datasource
    fileHandle = h.fileHandle()
    If Not (fileHandle <> 0) Then ASSERT_FAILURE "fileHandle != nullptr", "lib/sound/oggopus.cpp:109", "wz_opus_physfs_tell", "Bad PhysicsFS file handle passed in"

    out = PHYSFS_tell(fileHandle)
    wz_opus_physfs_tell = out
End Function

' https://github.com/xiph/opusfile
' https://opus-codec.org/docs/opusfile_api-0.12/index.html
Private Function wz_oggOpus_physfs_callbacks() As OpusFileCallbacks
    wz_oggOpus_physfs_callbacks.read_ = OPFN_wz_opus_physfs_read
    ' this one must return 0 on success, unlike its libvorbis counterpart
    wz_oggOpus_physfs_callbacks.seek_ = OPFN_wz_opus_physfs_seek
    wz_oggOpus_physfs_callbacks.tell = OPFN_wz_opus_physfs_tell
    wz_oggOpus_physfs_callbacks.close_ = 0      ' nullptr
End Function

' MARK: - Opus memory buffer callbacks

Public Function wz_opus_membuf_read(datasource As Variant, ptr_() As Byte, ByVal ptrOff As Long, ByVal nbytes_ As Long) As Long
    Dim pMemBuf As WZAudioDataMemoryBuffer, bytesToRead As Long
    If Not (Not datasource Is Nothing) Then ASSERT_FAILURE "datasource != nullptr", "lib/sound/oggopus.cpp:131", "wz_opus_membuf_read", "NULL decoder passed!"
    Set pMemBuf = datasource

    If nbytes_ < 0 Then
        wz_opus_membuf_read = -1
        Exit Function
    End If
    bytesToRead = pMemBuf.bufferSize() - pMemBuf.currPos
    If nbytes_ < bytesToRead Then bytesToRead = nbytes_
    If bytesToRead = 0 Then
        wz_opus_membuf_read = 0
        Exit Function
    End If
    pMemBuf.vbCopyOut ptr_, ptrOff, pMemBuf.currPos, bytesToRead     ' memcpy(_ptr, &(pMemBuf->buffer[pMemBuf->currPos]), bytesToRead)
    pMemBuf.currPos = pMemBuf.currPos + bytesToRead
    wz_opus_membuf_read = bytesToRead
End Function

Public Function wz_opus_membuf_seek(datasource As Variant, ByVal offset_ As Double, ByVal whence_ As Long) As Long
    Dim pMemBuf As WZAudioDataMemoryBuffer, newPos As Double
    If Not (Not datasource Is Nothing) Then ASSERT_FAILURE "datasource != nullptr", "lib/sound/oggopus.cpp:152", "wz_opus_membuf_seek", "NULL decoder passed!"
    Set pMemBuf = datasource

    Select Case whence_
    ' Seek to absolute position
    Case SEEK_SET
        newPos = offset_

    ' Seek `offset` ahead
    Case SEEK_CUR
        If pMemBuf.error_ Then
            wz_opus_membuf_seek = -1
            Exit Function
        End If
        newPos = pMemBuf.currPos + offset_

    ' Seek backwards from the end of the file
    Case SEEK_END
        If pMemBuf.error_ Then
            wz_opus_membuf_seek = -1
            Exit Function
        End If
        newPos = pMemBuf.bufferSize() - 1 - offset_

    ' unrecognized seek instruction
    Case Else
        ' indicate failure
        wz_opus_membuf_seek = -1
        Exit Function
    End Select

    If newPos < 0 Then
        wz_opus_membuf_seek = -1 ' failure
        Exit Function
    End If
    If newPos >= pMemBuf.bufferSize() Then
        wz_opus_membuf_seek = -1 ' failure
        Exit Function
    End If

    pMemBuf.currPos = newPos
    wz_opus_membuf_seek = 0 ' success
End Function

Public Function wz_opus_membuf_tell(datasource As Variant) As Double
    Dim pMemBuf As WZAudioDataMemoryBuffer
    If Not (Not datasource Is Nothing) Then ASSERT_FAILURE "datasource != nullptr", "lib/sound/oggopus.cpp:206", "wz_opus_membuf_tell", "NULL decoder passed!"
    Set pMemBuf = datasource
    wz_opus_membuf_tell = pMemBuf.currPos
End Function

' https://github.com/xiph/opusfile
' https://opus-codec.org/docs/opusfile_api-0.12/index.html
Private Function wz_oggOpus_membuf_callbacks() As OpusFileCallbacks
    wz_oggOpus_membuf_callbacks.read_ = OPFN_wz_opus_membuf_read
    ' this one must return 0 on success, unlike its libvorbis counterpart
    wz_oggOpus_membuf_callbacks.seek_ = OPFN_wz_opus_membuf_seek
    wz_oggOpus_membuf_callbacks.tell = OPFN_wz_opus_membuf_tell
    wz_oggOpus_membuf_callbacks.close_ = 0      ' nullptr
End Function

' MARK: - WzOpusDecoder

' WZOpusDecoder::fromFilename(const char*, bool bufferEntireStream = false)
Public Function WZOpusDecoder_fromFilename(ByVal fileName As String, Optional ByVal bufferEntireStream As Boolean = False) As WZOpusDecoder
    Dim result As WZOpusDecoder, fileHandle As Long, data As WzAudioDataPhysFSHandle, mem As WZAudioDataMemoryBuffer
    Dim buf() As Byte, size As Long
    Set result = Nothing

    If Not bufferEntireStream Then
        fileHandle = PHYSFS_openRead(fileName)
        debug_ 231, LOG_WZ, "fromFilename", "Reading...[directory: %s] %s", PHYSFS_getRealDir(fileName) & "", fileName
        If fileHandle = 0 Then
            debug_ 234, LOG_ERROR, "fromFilename", "sound_LoadTrackFromFile: PHYSFS_openRead(""%s"") failed with error: %s" & vbLf, fileName, WZ_PHYSFS_getLastError()
            Set WZOpusDecoder_fromFilename = Nothing
            Exit Function
        End If

        Set data = New WzAudioDataPhysFSHandle  ' std::make_unique<WzAudioDataPhysFSHandle>(fileHandle, fileName)
        data.construct fileHandle, fileName
        Set result = openOpusInternal(data, wz_oggOpus_physfs_callbacks())
    Else
        Set mem = New WZAudioDataMemoryBuffer   ' std::make_unique<WZAudioDataMemoryBuffer>()
        If Not loadFileToBufferVector(fileName, buf, size, False, False) Then
            debug_ 246, LOG_ERROR, "fromFilename", "sound_LoadTrackFromFile: Failed to open / read file contents: %s" & vbLf, fileName
            Set WZOpusDecoder_fromFilename = Nothing
            Exit Function
        End If
        mem.vbSetBuffer buf, size               ' (data->buffer, filled by loadFileToBufferVector)
        Set result = openOpusInternal(mem, wz_oggOpus_membuf_callbacks())
    End If

    If result Is Nothing Then
        debug_ 254, LOG_ERROR, "fromFilename", "OP failed to create decoder for %s", fileName
    End If
    Set WZOpusDecoder_fromFilename = result
End Function

' WZOpusDecoder::openOpusInternal(std::unique_ptr<WZAudioDataResourceInterface> data, const OpusFileCallbacks *cb)
Private Function openOpusInternal(ByVal data As WZAudioDataResourceInterface, cb As OpusFileCallbacks) As WZOpusDecoder
    Dim error_ As Long, of_ As Long, stream_ As Variant, noData() As Byte, decoder As WZOpusDecoder
    error_ = 0
    Set stream_ = data                          ' data.get()
    of_ = op_open_callbacks(stream_, cb, noData, 0, error_)
    Select Case error_
    Case OP_EREAD
        debug_ 266, LOG_ERROR, "openOpusInternal", "An underlying read, seek, or tell operation failed when it should have succeeded, or we failed to find data in the stream we had seen before."
        Set openOpusInternal = Nothing
        Exit Function
    Case OP_EFAULT
        debug_ 269, LOG_ERROR, "openOpusInternal", "There was a memory allocation failure, or an internal library error."
        Set openOpusInternal = Nothing
        Exit Function
    Case OP_EIMPL
        debug_ 272, LOG_ERROR, "openOpusInternal", "The stream used a feature that is not implemented, such as an unsupported channel family."
        Set openOpusInternal = Nothing
        Exit Function
    Case OP_EINVAL
        debug_ 275, LOG_ERROR, "openOpusInternal", "seek() was implemented and succeeded on this source, but tell() did not, or the starting position indicator was not equal to _initial_bytes."
        Set openOpusInternal = Nothing
        Exit Function
    Case OP_ENOTFORMAT
        debug_ 278, LOG_ERROR, "openOpusInternal", "The stream contained a link that did not have any logical Opus streams in it."
        Set openOpusInternal = Nothing
        Exit Function
    Case OP_EBADHEADER
        debug_ 281, LOG_ERROR, "openOpusInternal", "A required header packet was not properly formatted, contained illegal values, or was missing altogether."
        Set openOpusInternal = Nothing
        Exit Function
    Case OP_EVERSION
        debug_ 284, LOG_ERROR, "openOpusInternal", "An ID header contained an unrecognized version number."
        Set openOpusInternal = Nothing
        Exit Function
    Case OP_EBADLINK
        debug_ 287, LOG_ERROR, "openOpusInternal", "We failed to find data we had seen before after seeking."
        Set openOpusInternal = Nothing
        Exit Function
    Case OP_EBADTIMESTAMP
        debug_ 290, LOG_ERROR, "openOpusInternal", "The first or last timestamp in a link failed basic validity checks."
        Set openOpusInternal = Nothing
        Exit Function
    Case Else
    End Select
    If Not (error_ >= 0) Then ASSERT_FAILURE "error >= 0", "lib/sound/oggopus.cpp:295", "openOpusInternal", "Unexpected libopusfile error %i", error_

    ' const OpusHead* head = op_head(of, -1): (VB) a copy, never NULL ("OP failed to read header" cannot happen);
    ' of is NULL only after an error the switch does not list, where the C's op_head would crash
    If of_ = 0 Then
        Set openOpusInternal = Nothing
        Exit Function
    End If
    Set decoder = New WZOpusDecoder             ' new WZOpusDecoder(std::move(data), of, head)
    decoder.construct data, of_
    ' Use a negative number to get the ID header information of the current link
    ' (VB: New raises an error where new returns nullptr)
    Set openOpusInternal = decoder
End Function

' WZOpusDecoder::fromMemoryBuffer(std::unique_ptr<WZAudioDataMemoryBuffer>)
Public Function WZOpusDecoder_fromMemoryBuffer(ByVal data As WZAudioDataMemoryBuffer) As WZOpusDecoder
    Dim decoder As WZOpusDecoder
    Set decoder = openOpusInternal(data, wz_oggOpus_membuf_callbacks())
    If decoder Is Nothing Then
        debug_ 319, LOG_ERROR, "fromMemoryBuffer", "OP failed to create decoder for memory buffer"
    End If
    Set WZOpusDecoder_fromMemoryBuffer = decoder
End Function
