Attribute VB_Name = "opusfile_stream_c"
' opusfile (lib/sound/3rdparty/opusfile at 9d718345), src/stream.c: the stdio and memory stream callbacks, the
' _WIN32 branches; and (VB) the dispatch of OpusFileCallbacks' function pointers.
' (VB) A function pointer of OpusFileCallbacks is an OPFN_* number: this module's, and oggopus_c's for
' oggopus.cpp's callbacks; op_vbCallRead / Seek / Tell / Close call the function by its number. A void *
' stream is a Variant: a FILE * is a Long handle of this module's FILE table (the C runtime's FILE over a
' Win32 file handle: fread is ReadFile, fgetpos / fsetpos SetFilePointer, _filelengthi64 GetFileSize, with the
' C runtime's end-of-file flag), an OpusMemStream * a Long handle of its pool (1-based; 0 = NULL), oggopus_c's
' streams are its objects. A char * path or mode is a VB String (UTF-16 already: op_fopen and op_freopen open
' it with CreateFileW as _wfopen would; op_utf8_to_utf16 is ported on UTF-8 bytes). op_fdopen: the port has no
' C runtime file descriptors, so it fails as fdopen does on a descriptor that is not open.
Option Explicit

Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const OPEN_ALWAYS As Long = 4
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_BEGIN As Long = 0
Private Const FILE_CURRENT As Long = 1
Private Const FILE_END As Long = 2
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_TYPE_DISK As Long = 1
Private Const FILE_TYPE_REMOTE As Long = &H8000&

' (VB) the function pointer numbers of stream.c's callbacks
Public Const OPFN_op_fread As Long = 1
Public Const OPFN_op_fseek As Long = 2
Public Const OPFN_op_ftell As Long = 3
Public Const OPFN_fclose As Long = 4             ' (op_close_func)fclose
Public Const OPFN_op_fseek_fail As Long = 5
Public Const OPFN_op_mem_read As Long = 6
Public Const OPFN_op_mem_seek As Long = 7
Public Const OPFN_op_mem_tell As Long = 8
Public Const OPFN_op_mem_close As Long = 9

' #define OP_MEM_SIZE_MAX (~(size_t)0>>1), OP_MEM_DIFF_MAX ((ptrdiff_t)OP_MEM_SIZE_MAX): (VB) 2^63 as a Double
Private Const OP_MEM_SIZE_MAX As Double = 4294967296# * 2147483648#
Private Const OP_MEM_DIFF_MAX As Double = OP_MEM_SIZE_MAX

' The context information needed to read from a block of memory as if it were a
'   file.
Private Type OpusMemStream
    ' The block of memory to read from. ((VB) a copy)
    data() As Byte
    ' The total size of the block.
    '  This must be at most OP_MEM_SIZE_MAX to prevent signed overflow while
    '   seeking.
    size As Double                  ' ptrdiff_t
    ' The current file position.
    '  This is allowed to be set arbitrarily greater than size (i.e., past the end
    '   of the block, though we will not read data past the end of the block), but
    '   is not allowed to be negative (i.e., before the beginning of the block).
    pos As Double                   ' ptrdiff_t
    vbUsed As Boolean               ' (VB) the slot is allocated
End Type

' (VB) the C runtime's FILE: a Win32 file handle, the end-of-file and error flags
Private Type FILE_t
    handle As Long
    eof As Boolean
    error_ As Boolean
    vbUsed As Boolean
End Type

Private memStreams() As OpusMemStream
Private memStreamsSize As Long
Private files() As FILE_t
Private filesSize As Long

' (VB) the callbacks by number (the C calls (*_of->callbacks.read)(_of->stream,...) and so on)
Public Function op_vbCallRead(ByVal fn As Long, stream_ As Variant, ptr_() As Byte, ByVal ptrOff As Long, ByVal nbytes_ As Long) As Long
    Select Case fn
    Case OPFN_op_fread
        op_vbCallRead = op_fread(stream_, ptr_, ptrOff, nbytes_)
    Case OPFN_op_mem_read
        op_vbCallRead = op_mem_read(stream_, ptr_, ptrOff, nbytes_)
    Case OPFN_wz_opus_physfs_read
        op_vbCallRead = wz_opus_physfs_read(stream_, ptr_, ptrOff, nbytes_)
    Case OPFN_wz_opus_membuf_read
        op_vbCallRead = wz_opus_membuf_read(stream_, ptr_, ptrOff, nbytes_)
    Case Else
        op_vbCallRead = OP_EREAD
    End Select
End Function

Public Function op_vbCallSeek(ByVal fn As Long, stream_ As Variant, ByVal offset_ As Double, ByVal whence_ As Long) As Long
    Select Case fn
    Case OPFN_op_fseek
        op_vbCallSeek = op_fseek(stream_, offset_, whence_)
    Case OPFN_op_fseek_fail
        op_vbCallSeek = op_fseek_fail(stream_, offset_, whence_)
    Case OPFN_op_mem_seek
        op_vbCallSeek = op_mem_seek(stream_, offset_, whence_)
    Case OPFN_wz_opus_physfs_seek
        op_vbCallSeek = wz_opus_physfs_seek(stream_, offset_, whence_)
    Case OPFN_wz_opus_membuf_seek
        op_vbCallSeek = wz_opus_membuf_seek(stream_, offset_, whence_)
    Case Else
        op_vbCallSeek = -1
    End Select
End Function

Public Function op_vbCallTell(ByVal fn As Long, stream_ As Variant) As Double
    Select Case fn
    Case OPFN_op_ftell
        op_vbCallTell = op_ftell(stream_)
    Case OPFN_op_mem_tell
        op_vbCallTell = op_mem_tell(stream_)
    Case OPFN_wz_opus_physfs_tell
        op_vbCallTell = wz_opus_physfs_tell(stream_)
    Case OPFN_wz_opus_membuf_tell
        op_vbCallTell = wz_opus_membuf_tell(stream_)
    Case Else
        op_vbCallTell = -1
    End Select
End Function

Public Function op_vbCallClose(ByVal fn As Long, stream_ As Variant) As Long
    Select Case fn
    Case OPFN_fclose
        op_vbCallClose = fclose_(stream_)
    Case OPFN_op_mem_close
        op_vbCallClose = op_mem_close(stream_)
    Case Else
        op_vbCallClose = 0
    End Select
End Function

' (VB) the C runtime's FILE functions over the handle table
Private Function fopen_w(ByVal path_ As String, ByVal mode_ As String) As Long
    Dim i As Long, h As Long, access As Long, disp As Long
    Select Case Left$(mode_, 1)
    Case "r"
        access = GENERIC_READ
        disp = OPEN_EXISTING
    Case "w"
        access = GENERIC_WRITE
        disp = CREATE_ALWAYS
    Case "a"
        access = GENERIC_WRITE
        disp = OPEN_ALWAYS
    Case Else
        Exit Function
    End Select
    If InStr(mode_, "+") > 0 Then access = GENERIC_READ Or GENERIC_WRITE
    h = CreateFileW(StrPtr(path_), access, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, disp, FILE_ATTRIBUTE_NORMAL, 0)
    If h = INVALID_HANDLE_VALUE Then Exit Function
    If Left$(mode_, 1) = "a" Then SetFilePointer h, 0, 0, FILE_END
    For i = 1 To filesSize
        If Not files(i).vbUsed Then Exit For
    Next
    If i > filesSize Then
        filesSize = filesSize + 8
        ReDim Preserve files(1 To filesSize)
    End If
    files(i).handle = h
    files(i).eof = False
    files(i).error_ = False
    files(i).vbUsed = True
    fopen_w = i
End Function

' C: fclose
Private Function fclose_(stream_ As Variant) As Long
    Dim f As Long
    f = stream_
    If f < 1 Or f > filesSize Then fclose_ = -1: Exit Function
    If Not files(f).vbUsed Then fclose_ = -1: Exit Function
    files(f).vbUsed = False
    If CloseHandle(files(f).handle) = 0 Then fclose_ = -1 Else fclose_ = 0
End Function

' C: fgetpos (0 on success)
Private Function fgetpos_(ByVal f As Long, pos As Double) As Long
    Dim lo As Long, hi As Long
    hi = 0
    lo = SetFilePointer(files(f).handle, 0, hi, FILE_CURRENT)
    If lo = -1 Then
        If GetLastError() <> 0 Then fgetpos_ = -1: Exit Function
    End If
    pos = hi * 4294967296# + IIf(lo < 0, lo + 4294967296#, CDbl(lo))
    fgetpos_ = 0
End Function

' C: fsetpos (0 on success; clears the end-of-file flag)
Private Function fsetpos_(ByVal f As Long, ByVal pos As Double) As Long
    Dim lo As Long, hi As Long, r As Long, l As Double
    hi = Int(pos / 4294967296#)
    l = pos - hi * 4294967296#
    If l >= 2147483648# Then lo = l - 4294967296# Else lo = l
    r = SetFilePointer(files(f).handle, lo, hi, FILE_BEGIN)
    If r = -1 Then
        If GetLastError() <> 0 Then fsetpos_ = -1: Exit Function
    End If
    files(f).eof = False
    fsetpos_ = 0
End Function

Private Function op_fread(stream_ As Variant, ptr_() As Byte, ByVal ptrOff As Long, ByVal buf_size_ As Long) As Long
    Dim f As Long, ret As Long
    ' Check for empty read.
    If buf_size_ <= 0 Then op_fread = 0: Exit Function
    f = stream_                                 ' stream=(FILE *)_stream
    ' ret=fread(_ptr,1,_buf_size,stream)
    ret = 0
    If ReadFile(files(f).handle, ptr_(ptrOff), buf_size_, ret, 0) = 0 Then
        files(f).error_ = True
        ret = 0
    ElseIf ret < buf_size_ Then
        files(f).eof = True
    End If
    ' (OP_ASSERT(ret<=(size_t)_buf_size))
    ' If ret==0 and !feof(stream), there was a read error.
    If ret > 0 Or files(f).eof Then op_fread = ret Else op_fread = OP_EREAD
End Function

Private Function op_fseek(stream_ As Variant, ByVal offset_ As Double, ByVal whence_ As Long) As Long
    Dim pos As Double, ret As Long, f As Long, hi As Long, lo As Long
    ' (_WIN32) _fseeki64() is not exposed until MSVCRT80.
    '  This is the default starting with MSVC 2005 (_MSC_VER>=1400), but we want
    '   to allow linking against older MSVCRT versions for compatibility back to
    '   XP without installing extra runtime libraries.
    '  i686-pc-mingw32 does not have fseeko() and requires
    '   __MSVCRT_VERSION__>=0x800 for _fseeki64(), which screws up linking with
    '   other libraries (that don't use MSVCRT80 from MSVC 2005 by default).
    '  i686-w64-mingw32 does have fseeko() and respects _FILE_OFFSET_BITS, but I
    '   don't know how to detect that at compile time.
    '  We could just use fseeko64() (which is available in both), but it's
    '   implemented using fgetpos()/fsetpos() just like this code, except without
    '   the overflow checking, so we prefer our version.
    f = stream_
    ' We don't use fpos_t directly because it might be a struct if __STDC__ is
    '   non-zero or _INTEGRAL_MAX_BITS < 64.
    ' (OP_ASSERT(sizeof(pos)==sizeof(fpos_t)))
    ' Translate the seek to an absolute one.
    If whence_ = SEEK_CUR Then
        ret = fgetpos_(f, pos)
        If ret Then op_fseek = ret: Exit Function
    ElseIf whence_ = SEEK_END Then
        lo = GetFileSize(files(f).handle, hi)       ' pos=_filelengthi64(_fileno((FILE *)_stream))
        If lo = -1 And GetLastError() <> 0 Then
            pos = -1
        Else
            pos = hi * 4294967296# + IIf(lo < 0, lo + 4294967296#, CDbl(lo))
        End If
    ElseIf whence_ = SEEK_SET Then
        pos = 0
    Else
        op_fseek = -1
        Exit Function
    End If
    ' Check for errors or overflow.
    If pos < 0 Or offset_ < -pos Or offset_ > OP_INT64_MAX - pos Then op_fseek = -1: Exit Function
    pos = pos + offset_
    op_fseek = fsetpos_(f, pos)
End Function

Private Function op_ftell(stream_ As Variant) As Double
    Dim pos As Double
    ' (_WIN32) _ftelli64() is not exposed until MSVCRT80, and ftello()/ftello64() have
    '   the same problems as fseeko()/fseeko64() in MingW.
    '  See above for a more detailed explanation.
    ' (OP_ASSERT(sizeof(pos)==sizeof(fpos_t)))
    If fgetpos_(CLng(stream_), pos) Then op_ftell = -1 Else op_ftell = pos
End Function

' (OP_FILE_CALLBACKS: op_fread, op_fseek, op_ftell, fclose)

' Windows doesn't accept UTF-8 by default, and we don't have a wchar_t API,
'   so if we just pass the path to fopen(), then there'd be no way for a user
'   of our API to open a Unicode filename.
'  Instead, we translate from UTF-8 to UTF-16 and use Windows' wchar_t API.
'  This makes this API more consistent with platforms where the character set
'   used by fopen is the same as used on disk, which is generally UTF-8, and
'   with our metadata API, which always uses UTF-8.
' (VB) _src is src_(0 ...) up to its NUL or end; returns vbNullString for NULL
Public Function op_utf8_to_utf16(src_() As Byte, ByVal len0 As Long) As String
    Dim dst() As Integer, len_ As Long, si As Long, di As Long, c0 As Long, c1 As Long, c2 As Long, c3 As Long
    Dim w As Long, s As String
    len_ = 0
    Do While len_ < len0                        ' len=strlen(_src)
        If src_(len_) = 0 Then Exit Do
        len_ = len_ + 1
    Loop
    ' Worst-case output is 1 wide character per 1 input character.
    ReDim dst(0 To len_)
    di = 0
    si = 0
    Do While si < len_
        c0 = src_(si)
        If (c0 And &H80) = 0 Then
            ' Start byte says this is a 1-byte sequence.
            dst(di) = c0
            di = di + 1
            GoTo continue_
        Else
            ' This is safe, because c0 was not 0 and _src is NUL-terminated.
            If si + 1 < len_ Then c1 = src_(si + 1) Else c1 = 0
            If (c1 And &HC0) = &H80 Then
                ' Found at least one continuation byte.
                If (c0 And &HE0) = &HC0 Then
                    ' Start byte says this is a 2-byte sequence.
                    w = (c0 And &H1F) * &H40 Or (c1 And &H3F)
                    If w >= &H80 Then
                        ' This is a 2-byte sequence that is not overlong.
                        dst(di) = w
                        di = di + 1
                        si = si + 1
                        GoTo continue_
                    End If
                Else
                    ' This is safe, because c1 was not 0 and _src is NUL-terminated.
                    If si + 2 < len_ Then c2 = src_(si + 2) Else c2 = 0
                    If (c2 And &HC0) = &H80 Then
                        ' Found at least two continuation bytes.
                        If (c0 And &HF0) = &HE0 Then
                            ' Start byte says this is a 3-byte sequence.
                            w = ((c0 And &HF) * &H1000& Or (c1 And &H3F) * &H40 Or (c2 And &H3F)) And &HFFFF&
                            If w >= &H800 And (w < &HD800& Or w >= &HE000&) And w < &HFFFE& Then
                                ' This is a 3-byte sequence that is not overlong, not a
                                '   UTF-16 surrogate pair value, and not a 'not a character'
                                '   value.
                                If w >= &H8000& Then dst(di) = w - &H10000 Else dst(di) = w
                                di = di + 1
                                si = si + 2
                                GoTo continue_
                            End If
                        Else
                            ' This is safe, because c2 was not 0 and _src is
                            '   NUL-terminated.
                            If si + 3 < len_ Then c3 = src_(si + 3) Else c3 = 0
                            If (c3 And &HC0) = &H80 Then
                                ' Found at least three continuation bytes.
                                If (c0 And &HF8) = &HF0 Then
                                    ' Start byte says this is a 4-byte sequence.
                                    ' w=(c0&7)<<18|(c1&0x3F)<<12|(c2&0x3F)<<6&(c3&0x3F) ((VB) the C's & as it is)
                                    w = (c0 And 7) * &H40000 Or (c1 And &H3F) * &H1000& Or ((c2 And &H3F) * &H40 And (c3 And &H3F))
                                    If w >= &H10000 And w < &H110000 Then
                                        ' This is a 4-byte sequence that is not overlong and not
                                        '   greater than the largest valid Unicode code point.
                                        '  Convert it to a surrogate pair.
                                        w = w - &H10000
                                        dst(di) = (&HD800& + w \ &H400&) - &H10000
                                        di = di + 1
                                        dst(di) = (&HDC00& + (w And &H3FF)) - &H10000
                                        di = di + 1
                                        si = si + 3
                                        GoTo continue_
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        ' If we got here, we encountered an illegal UTF-8 sequence.
        op_utf8_to_utf16 = vbNullString            ' _ogg_free(dst); return NULL
        Exit Function
continue_:
        si = si + 1
    Loop
    ' (OP_ASSERT(di<=len))
    ' dst[di]='\0'
    s = String$(di, vbNullChar)
    If di > 0 Then CopyMemory ByVal StrPtr(s), dst(0), di * 2
    op_utf8_to_utf16 = s
End Function

' fsetpos() internally dispatches to the win32 API call SetFilePointer().
'  According to SetFilePointer()'s documentation [0], the behavior is
'   undefined if you do not call it on "a file stored on a seeking device".
'  However, none of the MSVCRT seeking functions verify what kind of file is
'   being used before calling it (which I believe is a bug, since they are
'   supposed to fail and return an error, but it is a bug that has been there
'   for multiple decades now).
'  In practice, SetFilePointer() appears to succeed for things like stdin,
'   even when you are not just piping in a regular file, which prevents the use
'   of this API to determine whether it is possible to seek in a file at all.
'  Therefore, we take the approach recommended by the SetFilePointer()
'   documentation and confirm the type of file using GetFileType() first.
'  We do this once, when the file is opened, and return the corresponding
'   callback in order to avoid an extra win32 API call on every seek in the
'   common case.
'  Hopefully the return value of GetFileType() cannot actually change for the
'   lifetime of a file handle.
'  [0] https://docs.microsoft.com/en-us/windows/desktop/api/fileapi/nf-fileapi-setfilepointer
Private Function op_fseek_fail(stream_ As Variant, ByVal offset_ As Double, ByVal whence_ As Long) As Long
    op_fseek_fail = -1
End Function

' (OP_UNSEEKABLE_FILE_CALLBACKS: op_fread, op_fseek_fail, op_ftell, fclose)

Private Sub op_get_file_callbacks(ByVal fp_ As Long, cb_ As OpusFileCallbacks)
    Dim h_file As Long
    h_file = files(fp_).handle                  ' _get_osfhandle(_fileno(_fp))
    If h_file <> -1 Then
        If (GetFileType(h_file) And Not FILE_TYPE_REMOTE) = FILE_TYPE_DISK Then
            cb_.read_ = OPFN_op_fread           ' *_cb=OP_FILE_CALLBACKS
            cb_.seek_ = OPFN_op_fseek
            cb_.tell = OPFN_op_ftell
            cb_.close_ = OPFN_fclose
            Exit Sub
        End If
    End If
    cb_.read_ = OPFN_op_fread                   ' *_cb=OP_UNSEEKABLE_FILE_CALLBACKS
    cb_.seek_ = OPFN_op_fseek_fail
    cb_.tell = OPFN_op_ftell
    cb_.close_ = OPFN_fclose
End Sub

' (VB) returns the FILE * (0 = NULL)
Public Function op_fopen(cb_ As OpusFileCallbacks, ByVal path_ As String, ByVal mode_ As String) As Variant
    Dim fp As Long
    fp = 0
    ' (_WIN32: wpath=op_utf8_to_utf16(_path); wmode=op_utf8_to_utf16(_mode): the VB Strings are UTF-16)
    fp = fopen_w(path_, mode_)                  ' fp=_wfopen(wpath,wmode)
    If fp <> 0 Then op_get_file_callbacks fp, cb_
    op_fopen = fp
End Function

' (VB) no C runtime file descriptors: fails (fdopen's EBADF)
Public Function op_fdopen(cb_ As OpusFileCallbacks, ByVal fd_ As Long, ByVal mode_ As String) As Variant
    Dim fp As Long
    fp = 0                                      ' fp=fdopen(_fd,_mode)
    If fp <> 0 Then op_get_file_callbacks fp, cb_
    op_fdopen = fp
End Function

' (VB) _stream is a FILE * of op_fopen; the same FILE * is returned, or NULL (the old file is closed either way)
Public Function op_freopen(cb_ As OpusFileCallbacks, ByVal path_ As String, ByVal mode_ As String, stream_ As Variant) As Variant
    Dim fp As Long, old As Long, nw As Long
    fp = 0
    ' fp=_wfreopen(wpath,wmode,(FILE *)_stream)
    old = stream_
    fclose_ stream_
    nw = fopen_w(path_, mode_)
    If nw <> 0 Then
        If nw <> old Then                       ' (the same FILE: move the new handle into the old slot)
            files(old) = files(nw)
            files(nw).vbUsed = False
        End If
        fp = old
    End If
    If fp <> 0 Then op_get_file_callbacks fp, cb_
    op_freopen = fp
End Function

Private Function op_mem_read(stream_ As Variant, ptr_() As Byte, ByVal ptrOff As Long, ByVal buf_size_ As Long) As Long
    Dim stream As Long, size As Double, pos As Double
    stream = stream_
    ' Check for empty read.
    If buf_size_ <= 0 Then op_mem_read = 0: Exit Function
    size = memStreams(stream).size
    pos = memStreams(stream).pos
    ' Check for EOF.
    If pos >= size Then op_mem_read = 0: Exit Function
    ' Check for a short read.
    If size - pos < buf_size_ Then buf_size_ = size - pos
    CopyMemory ptr_(ptrOff), memStreams(stream).data(pos), buf_size_
    pos = pos + buf_size_
    memStreams(stream).pos = pos
    op_mem_read = buf_size_
End Function

Private Function op_mem_seek(stream_ As Variant, ByVal offset_ As Double, ByVal whence_ As Long) As Long
    Dim stream As Long, pos As Double, size As Double
    stream = stream_
    pos = memStreams(stream).pos
    ' (OP_ASSERT(pos>=0))
    Select Case whence_
    Case SEEK_SET
        ' Check for overflow:
        If offset_ < 0 Or offset_ > OP_MEM_DIFF_MAX Then op_mem_seek = -1: Exit Function
        pos = offset_
    Case SEEK_CUR
        ' Check for overflow:
        If offset_ < -pos Or offset_ > OP_MEM_DIFF_MAX - pos Then op_mem_seek = -1: Exit Function
        pos = pos + offset_
    Case SEEK_END
        size = memStreams(stream).size
        ' (OP_ASSERT(size>=0))
        ' Check for overflow:
        If offset_ < -size Or offset_ > OP_MEM_DIFF_MAX - size Then op_mem_seek = -1: Exit Function
        pos = size + offset_
    Case Else
        op_mem_seek = -1
        Exit Function
    End Select
    memStreams(stream).pos = pos
    op_mem_seek = 0
End Function

Private Function op_mem_tell(stream_ As Variant) As Double
    op_mem_tell = memStreams(CLng(stream_)).pos
End Function

Private Function op_mem_close(stream_ As Variant) As Long
    Dim blank As OpusMemStream, stream As Long
    stream = stream_
    memStreams(stream) = blank                  ' _ogg_free(_stream)
    op_mem_close = 0
End Function

' (OP_MEM_CALLBACKS: op_mem_read, op_mem_seek, op_mem_tell, op_mem_close)

' (VB) the stream keeps a copy of data_(0 .. size_ - 1) (the C's points at the caller's block); returns its
' handle (0 = NULL)
Public Function op_mem_stream_create(cb_ As OpusFileCallbacks, data_() As Byte, ByVal size_ As Long) As Variant
    Dim stream As Long
    If size_ > OP_MEM_SIZE_MAX Then op_mem_stream_create = 0: Exit Function
    ' stream=(OpusMemStream *)_ogg_malloc(sizeof(*stream))
    For stream = 1 To memStreamsSize
        If Not memStreams(stream).vbUsed Then Exit For
    Next
    If stream > memStreamsSize Then
        memStreamsSize = memStreamsSize + 4
        ReDim Preserve memStreams(1 To memStreamsSize)
    End If
    memStreams(stream).vbUsed = True
    cb_.read_ = OPFN_op_mem_read                ' *_cb=*&OP_MEM_CALLBACKS
    cb_.seek_ = OPFN_op_mem_seek
    cb_.tell = OPFN_op_mem_tell
    cb_.close_ = OPFN_op_mem_close
    If size_ > 0 Then
        ReDim memStreams(stream).data(0 To size_ - 1)
        CopyMemory memStreams(stream).data(0), data_(0), size_
    End If
    memStreams(stream).size = size_
    memStreams(stream).pos = 0
    op_mem_stream_create = stream
End Function
