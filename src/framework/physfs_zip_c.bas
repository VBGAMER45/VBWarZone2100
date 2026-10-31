Attribute VB_Name = "physfs_zip_c"
' PhysicsFS 3.2.0's src/physfs_archiver_zip.c, whole (ZIP support routines for PhysicsFS, read-only as the C
' is: zip64, the DOS times, symlinks and "traditional" PKWARE encryption as the C handles them), with the
' directory tree it keeps its entries in (physfs.c's __PHYSFS_DirTree*: __PHYSFS_DirTreeInit, hashPathName,
' addAncestors, __PHYSFS_DirTreeAdd, __PHYSFS_DirTreeFind, __PHYSFS_DirTreeEnumerate, __PHYSFS_DirTreeDeinit:
' PHYSFS_DirTreeAdd__ ... here) and zlib through physfs_miniz_c. physfs_win dispatches to the ZIP_ functions
' (__PHYSFS_Archiver_ZIP's table) and owns the PHYSFS_Io objects (an io is its handle; io->read(io, ...) is
' PHYSFS_Io_read(io, ...)); a ZIP io's opaque is a ZIPfileinfo of this module's pool. (VB) A ZIPinfo (the
' archiver's opaque) is a 1-based handle into zipInfos(); its tree's entries are its own array, a ZIPentry *
' an index into it (0 = NULL), the tree root entry 1. A ZIPfileinfo knows its archive besides its entry.
' Entry names are VB Strings (the zip's bytes decoded as UTF-8, as PhysFS treats them); PHYSFS_uint64 sizes
' and offsets are Doubles, the ZIPfileinfo's PHYSFS_uint32 positions too (entries under 4 GB, as the C's
' positions need). A buffer pointer is (array, offset). The hash of a tree name is computed modulo 64 (the
' bucket count): see hashPathName.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal src As Long, ByVal n As Long)
Private Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (ByVal lpTimeZoneInformation As Long, lpLocalTime As SYSTEMTIME_vb, lpUniversalTime As SYSTEMTIME_vb) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME_vb, lpFileTime As FILETIME_vb) As Long

Private Type SYSTEMTIME_vb
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type FILETIME_vb
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'
' A buffer of ZIP_READBUFSIZE is allocated for each compressed file opened,
'  and is freed when you close the file; compressed data is read into
'  this buffer, and then is decompressed into the buffer passed to
'  PHYSFS_read().
'
' Uncompressed entries in a zipfile do not allocate this buffer; they just
'  read data directly into the buffer passed to PHYSFS_read().
'
' Depending on your speed and memory requirements, you should tweak this
'  value.
'
Private Const ZIP_READBUFSIZE As Long = 16& * 1024

'
' Entries are "unresolved" until they are first opened. At that time,
'  local file headers parsed/validated, data offsets will be updated to look
'  at the actual file data instead of the header, and symlinks will be
'  followed and optimized. This means that we don't seek and read around the
'  archive until forced to do so, and after the first time, we had to do
'  less reading and parsing, which is very CD-ROM friendly.
'
' ZipResolveType
Private Const ZIP_UNRESOLVED_FILE As Long = 0
Private Const ZIP_UNRESOLVED_SYMLINK As Long = 1
Private Const ZIP_RESOLVING As Long = 2
Private Const ZIP_RESOLVED As Long = 3
Private Const ZIP_DIRECTORY As Long = 4
Private Const ZIP_BROKEN_FILE As Long = 5
Private Const ZIP_BROKEN_SYMLINK As Long = 6

'
' One ZIPentry is kept for each file in an open ZIP archive.
'
Private Type ZIPentry
    ' __PHYSFS_DirTreeEntry tree: manages directory tree
    tree_name As String                 ' char *name
    tree_hashnext As Long               ' struct __PHYSFS_DirTreeEntry *hashnext (an entry index; 0 = NULL)
    tree_children As Long
    tree_sibling As Long
    tree_isdir As Long
    symlink As Long                     ' NULL or file we symlink to
    resolved As Long                    ' Have we resolved file/symlink? (ZipResolveType)
    offset As Double                    ' offset of data in archive
    version As Long                     ' version made by
    version_needed As Long              ' version needed to extract
    general_bits As Long                ' general purpose bits
    compression_method As Long          ' compression method
    crc As Long                         ' crc-32 (its bits)
    compressed_size As Double           ' compressed size
    uncompressed_size As Double         ' uncompressed size
    last_mod_time As Double             ' last file mod time
    dos_mod_time As Long                ' original MS-DOS style mod time (its bits)
End Type

'
' One ZIPinfo is kept for each open ZIP archive.
'
Private Type ZIPinfo
    ' __PHYSFS_DirTree tree: manages directory tree.
    tree_root As Long                   ' __PHYSFS_DirTreeEntry *root (the entry 1)
    tree_hash() As Long                 ' __PHYSFS_DirTreeEntry **hash (the buckets' first entries)
    tree_hashBuckets As Long
    tree_entrylen As Long               ' (sizeof (ZIPentry): unused)
    tree_case_sensitive As Long
    tree_only_usascii As Long
    entries() As ZIPentry               ' (VB) the tree's entries (1-based): the allocator's
    entryCount As Long
    io As Long                          ' the i/o interface for this archive.
    zip64 As Long                       ' non-zero if this is a Zip64 archive.
    has_crypto As Long                  ' non-zero if any entry uses encryption.
    inUse As Boolean                    ' (VB) the pool's
End Type

'
' One ZIPfileinfo is kept for each open file in a ZIP archive.
'
Private Type ZIPfileinfo
    archive As Long                     ' (VB) the ZIPinfo the entry is in
    entry As Long                       ' Info on file.
    io As Long                          ' physical file handle.
    compressed_position As Double       ' offset in compressed data.
    uncompressed_position As Double     ' tell() position.
    buffer() As Byte                    ' decompression buffer.
    hasBuffer As Boolean                ' (VB) buffer != NULL
    crypto_keys(0 To 2) As Long         ' for "traditional" crypto.
    initial_crypto_keys(0 To 2) As Long ' for "traditional" crypto.
    stream As mz_stream                 ' zlib stream state.
    inUse As Boolean                    ' (VB) the pool's
End Type

' Magic numbers...
Private Const ZIP_LOCAL_FILE_SIG As Long = &H4034B50
Private Const ZIP_CENTRAL_DIR_SIG As Long = &H2014B50
Private Const ZIP_END_OF_CENTRAL_DIR_SIG As Long = &H6054B50
Private Const ZIP64_END_OF_CENTRAL_DIR_SIG As Long = &H6064B50
Private Const ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIG As Long = &H7064B50
Private Const ZIP64_EXTENDED_INFO_EXTRA_FIELD_SIG As Long = &H1

' compression methods...
Private Const COMPMETH_NONE As Long = 0
' ...and others...

Private Const UNIX_FILETYPE_MASK As Long = &HF000&           ' 0170000
Private Const UNIX_FILETYPE_SYMLINK As Long = &HA000&        ' 0120000

Private Const ZIP_GENERAL_BITS_TRADITIONAL_CRYPTO As Long = 1    ' (1 << 0)
Private Const ZIP_GENERAL_BITS_IGNORE_LOCAL_HEADER As Long = 8   ' (1 << 3)

Private Const TWO32 As Double = 4294967296#

' (VB) the pools: the ZIPinfo and ZIPfileinfo allocations
Private zipInfos() As ZIPinfo
Private zipInfosCount As Long
Private fileInfos() As ZIPfileinfo
Private fileInfosCount As Long

'--------------------------------------------------------------------------
' (VB) helpers

Private Function newZipInfo() As Long
    Dim i As Long, blank As ZIPinfo
    For i = 1 To zipInfosCount
        If Not zipInfos(i).inUse Then Exit For
    Next
    If i > zipInfosCount Then
        zipInfosCount = zipInfosCount + 1
        ReDim Preserve zipInfos(1 To zipInfosCount)
    End If
    zipInfos(i) = blank                                 ' memset(info, '\0', sizeof (ZIPinfo))
    zipInfos(i).inUse = True
    newZipInfo = i
End Function

Private Function newFileInfo() As Long
    Dim i As Long, blank As ZIPfileinfo
    For i = 1 To fileInfosCount
        If Not fileInfos(i).inUse Then Exit For
    Next
    If i > fileInfosCount Then
        fileInfosCount = fileInfosCount + 1
        ReDim Preserve fileInfos(1 To fileInfosCount)
    End If
    fileInfos(i) = blank                                ' memset(finfo, '\0', sizeof (ZIPfileinfo))
    fileInfos(i).inUse = True
    newFileInfo = i
End Function

Private Sub freeFileInfo(ByVal finfo As Long)
    Dim blank As ZIPfileinfo
    fileInfos(finfo) = blank                            ' allocator.Free(finfo)
End Sub

'--------------------------------------------------------------------------
' physfs.c: the directory tree (__PHYSFS_DirTree)

Private Function PHYSFS_DirTreeInit__(ByVal info As Long, ByVal case_sensitive As Long, ByVal only_usascii As Long) As Long
    Dim blank As ZIPentry
    ' (static char rootpath[2] = { '/', '\0' };)
    ' (assert(entrylen >= sizeof (__PHYSFS_DirTreeEntry));)

    ' (memset(dt, '\0', sizeof (*dt));)
    zipInfos(info).tree_case_sensitive = case_sensitive
    zipInfos(info).tree_only_usascii = only_usascii

    ReDim zipInfos(info).entries(1 To 64)              ' dt->root = allocator.Malloc(entrylen)
    zipInfos(info).entryCount = 1
    zipInfos(info).entries(1) = blank
    zipInfos(info).tree_root = 1
    zipInfos(info).entries(1).tree_name = "/"
    zipInfos(info).entries(1).tree_isdir = 1
    zipInfos(info).tree_hashBuckets = 64
    If zipInfos(info).tree_hashBuckets = 0 Then
        zipInfos(info).tree_hashBuckets = 1
    End If

    ReDim zipInfos(info).tree_hash(0 To zipInfos(info).tree_hashBuckets - 1)

    PHYSFS_DirTreeInit__ = 1
End Function

' __PHYSFS_hashString(name) % dt->hashBuckets (the zip archiver's tree is case-sensitive). (VB) With 64 buckets
' the remainder is the hash's low 6 bits, and ((hash << 5) + hash) ^ ch computes those from the low 6 bits of
' hash and ch alone, so the hash is kept modulo 64; ch is each UTF-16 code unit's low byte (an ASCII name's
' bytes, as the C's; another name may fall into another bucket than the C's, which nothing outside sees).
Private Function hashPathName(ByVal info As Long, ByVal name_ As String) As Long
    Dim b() As Byte, i As Long, hashval As Long
    hashval = 5381 And 63
    If Len(name_) > 0 Then
        b = name_
        For i = 0 To UBound(b) Step 2
            hashval = ((hashval * 33) Xor b(i)) And 63
        Next
    End If
    hashPathName = hashval Mod zipInfos(info).tree_hashBuckets
End Function

' Fill in missing parent directories.
Private Function addAncestors(ByVal info As Long, ByVal name_ As String) As Long
    Dim retval As Long, sep As Long
    retval = zipInfos(info).tree_root
    sep = InStrRev(name_, "/")

    If sep > 0 Then
        ' (*sep = '\0';  /* chop off last piece. */)
        retval = PHYSFS_DirTreeFind__(info, Left$(name_, sep - 1))

        If retval <> 0 Then
            ' (*sep = '/';)
            If zipInfos(info).entries(retval).tree_isdir = 0 Then
                PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT
                addAncestors = 0
                Exit Function
            End If
            addAncestors = retval                       ' already hashed.
            Exit Function
        End If

        ' okay, this is a new dir. Build and hash us.
        retval = PHYSFS_DirTreeAdd__(info, Left$(name_, sep - 1), 1)
        ' (*sep = '/';)
    End If

    addAncestors = retval
End Function

Private Function PHYSFS_DirTreeAdd__(ByVal info As Long, ByVal name_ As String, ByVal isdir As Long) As Long
    Dim retval As Long, hashval As Long, parent As Long, blank As ZIPentry
    retval = PHYSFS_DirTreeFind__(info, name_)
    If retval = 0 Then
        parent = addAncestors(info, name_)
        If parent = 0 Then
            PHYSFS_DirTreeAdd__ = 0
            Exit Function
        End If
        ' (assert(dt->entrylen >= sizeof (__PHYSFS_DirTreeEntry));)
        ' retval = (__PHYSFS_DirTreeEntry *) allocator.Malloc(alloclen);
        If zipInfos(info).entryCount = UBound(zipInfos(info).entries) Then
            ReDim Preserve zipInfos(info).entries(1 To 2 * zipInfos(info).entryCount)
        End If
        zipInfos(info).entryCount = zipInfos(info).entryCount + 1
        retval = zipInfos(info).entryCount
        zipInfos(info).entries(retval) = blank           ' memset(retval, '\0', dt->entrylen)
        zipInfos(info).entries(retval).tree_name = name_
        hashval = hashPathName(info, name_)
        zipInfos(info).entries(retval).tree_hashnext = zipInfos(info).tree_hash(hashval)
        zipInfos(info).tree_hash(hashval) = retval
        zipInfos(info).entries(retval).tree_sibling = zipInfos(info).entries(parent).tree_children
        zipInfos(info).entries(retval).tree_isdir = isdir
        zipInfos(info).entries(parent).tree_children = retval
    End If

    PHYSFS_DirTreeAdd__ = retval
End Function

' Find the __PHYSFS_DirTreeEntry for a path in platform-independent notation.
Private Function PHYSFS_DirTreeFind__(ByVal info As Long, ByVal path As String) As Long
    Dim cs As Long, hashval As Long, prev As Long, retval As Long, cmp As Long
    cs = zipInfos(info).tree_case_sensitive

    If Len(path) = 0 Then
        PHYSFS_DirTreeFind__ = zipInfos(info).tree_root
        Exit Function
    End If

    hashval = hashPathName(info, path)
    retval = zipInfos(info).tree_hash(hashval)
    Do While retval <> 0
        If cs <> 0 Then
            cmp = StrComp(zipInfos(info).entries(retval).tree_name, path, vbBinaryCompare)
        Else
            cmp = StrComp(zipInfos(info).entries(retval).tree_name, path, vbTextCompare)    ' PHYSFS_utf8stricmp
        End If
        If cmp = 0 Then
            If prev <> 0 Then   ' move this to the front of the list
                zipInfos(info).entries(prev).tree_hashnext = zipInfos(info).entries(retval).tree_hashnext
                zipInfos(info).entries(retval).tree_hashnext = zipInfos(info).tree_hash(hashval)
                zipInfos(info).tree_hash(hashval) = retval
            End If

            PHYSFS_DirTreeFind__ = retval
            Exit Function
        End If

        prev = retval
        retval = zipInfos(info).entries(retval).tree_hashnext
    Loop

    PHYSFS_setErrorCode PHYSFS_ERR_NOT_FOUND
    PHYSFS_DirTreeFind__ = 0
End Function

' (VB) cb: physfs_win's callback number (PHYSFS_EnumerateCallback); origdir passed on
Private Function PHYSFS_DirTreeEnumerate__(ByVal opaque As Long, ByVal dname As String, ByVal cb As Long, ByVal origdir As String) As Long
    Dim retval As Long, entry As Long, name_ As String, ptr As Long
    retval = PHYSFS_ENUM_OK
    entry = PHYSFS_DirTreeFind__(opaque, dname)
    If entry = 0 Then
        PHYSFS_setErrorCode PHYSFS_ERR_NOT_FOUND
        PHYSFS_DirTreeEnumerate__ = PHYSFS_ENUM_ERROR
        Exit Function
    End If

    entry = zipInfos(opaque).entries(entry).tree_children

    Do While (entry <> 0) And (retval = PHYSFS_ENUM_OK)
        name_ = zipInfos(opaque).entries(entry).tree_name
        ptr = InStrRev(name_, "/")
        If ptr > 0 Then
            retval = PHYSFS_EnumerateCallback_call(cb, origdir, Mid$(name_, ptr + 1))
        Else
            retval = PHYSFS_EnumerateCallback_call(cb, origdir, name_)
        End If
        If retval = PHYSFS_ENUM_ERROR Then
            PHYSFS_setErrorCode PHYSFS_ERR_APP_CALLBACK
            PHYSFS_DirTreeEnumerate__ = retval
            Exit Function
        End If
        entry = zipInfos(opaque).entries(entry).tree_sibling
    Loop

    PHYSFS_DirTreeEnumerate__ = retval
End Function

Private Sub PHYSFS_DirTreeDeinit__(ByVal info As Long)
    ' (if (!dt) return;)
    ' (assert(dt->root->sibling == NULL); the entries and the hash are freed)
    Erase zipInfos(info).entries
    zipInfos(info).entryCount = 0
    zipInfos(info).tree_root = 0
    Erase zipInfos(info).tree_hash
End Sub

'--------------------------------------------------------------------------

' support for "traditional" PKWARE encryption.
Private Function zip_entry_is_tradional_crypto(ByVal info As Long, ByVal entry As Long) As Boolean
    zip_entry_is_tradional_crypto = (zipInfos(info).entries(entry).general_bits And ZIP_GENERAL_BITS_TRADITIONAL_CRYPTO) <> 0
End Function ' zip_entry_is_traditional_crypto

Private Function zip_entry_ignore_local_header(ByVal info As Long, ByVal entry As Long) As Boolean
    zip_entry_ignore_local_header = (zipInfos(info).entries(entry).general_bits And ZIP_GENERAL_BITS_IGNORE_LOCAL_HEADER) <> 0
End Function ' zip_entry_is_traditional_crypto

Private Function zip_crypto_crc32(ByVal crc As Long, ByVal val_ As Long) As Long
    Dim i As Long, xorval As Long
    xorval = (crc Xor val_) And &HFF&
    For i = 0 To 7
        If (xorval And 1) <> 0 Then
            xorval = &HEDB88320 Xor SHRU(xorval, 1)
        Else
            xorval = SHRU(xorval, 1)
        End If
    Next
    zip_crypto_crc32 = xorval Xor SHRU(crc, 8)
End Function ' zip_crc32

Private Sub zip_update_crypto_keys(keys() As Long, ByVal val_ As Long)
    keys(0) = zip_crypto_crc32(keys(0), val_)
    keys(1) = FromU32(U32(keys(1)) + (keys(0) And &HFF&))
    keys(1) = FromU32(U32(MUL32(keys(1), 134775813)) + 1)
    keys(2) = zip_crypto_crc32(keys(2), SHRU(keys(1), 24) And &HFF&)
End Sub ' zip_update_crypto_keys

Private Function zip_decrypt_byte(keys() As Long) As Long
    Dim tmp As Long
    tmp = (keys(2) Or 2) And &HFFFF&                    ' const PHYSFS_uint16 tmp = keys[2] | 2;
    zip_decrypt_byte = SHRU(MUL32(tmp, tmp Xor 1), 8) And &HFF&
End Function ' zip_decrypt_byte

Private Function zip_read_decrypt(ByVal finfo As Long, buf() As Byte, ByVal bufOffset As Long, ByVal len_ As Long) As Long
    Dim io As Long, br As Long, i As Long, ch As Long
    io = fileInfos(finfo).io
    br = PHYSFS_Io_read(io, buf, bufOffset, len_)

    ' Decompression the new data if necessary.
    If zip_entry_is_tradional_crypto(fileInfos(finfo).archive, fileInfos(finfo).entry) And (br > 0) Then
        ' (PHYSFS_uint32 *keys = finfo->crypto_keys;)
        For i = 0 To br - 1
            ch = buf(bufOffset + i) Xor zip_decrypt_byte(fileInfos(finfo).crypto_keys)
            zip_update_crypto_keys fileInfos(finfo).crypto_keys, ch
            buf(bufOffset + i) = ch
        Next
    End If

    zip_read_decrypt = br
End Function ' zip_read_decrypt

' (VB) password: the bytes after the '$' (C: a NUL-terminated string)
Private Function zip_prep_crypto_keys(ByVal finfo As Long, crypto_header() As Byte, password() As Byte, ByVal passwordLen As Long) As Long
    ' It doesn't appear to be documented in PKWare's APPNOTE.TXT, but you
    '   need to use a different byte in the header to verify the password
    '   if general purpose bit 3 is set. Discovered this from Info-Zip.
    '   That's what the (verifier) value is doing, below.

    Dim usedate As Boolean, verifier As Long, finalbyte As Long, i As Long, c As Long, info As Long, entry As Long
    info = fileInfos(finfo).archive
    entry = fileInfos(finfo).entry
    usedate = zip_entry_ignore_local_header(info, entry)
    If usedate Then
        verifier = SHRU(zipInfos(info).entries(entry).dos_mod_time, 8) And &HFF&
    Else
        verifier = SHRU(zipInfos(info).entries(entry).crc, 24) And &HFF&
    End If
    finalbyte = 0
    i = 0

    ' initialize vector with defaults, then password, then header.
    fileInfos(finfo).crypto_keys(0) = 305419896
    fileInfos(finfo).crypto_keys(1) = 591751049
    fileInfos(finfo).crypto_keys(2) = 878082192

    For i = 0 To passwordLen - 1                        ' while (*password)
        zip_update_crypto_keys fileInfos(finfo).crypto_keys, password(i)
    Next

    For i = 0 To 11
        c = crypto_header(i) Xor zip_decrypt_byte(fileInfos(finfo).crypto_keys)
        zip_update_crypto_keys fileInfos(finfo).crypto_keys, c
        finalbyte = c
    Next

    ' you have a 1/256 chance of passing this test incorrectly. :/
    If finalbyte <> verifier Then
        PHYSFS_setErrorCode PHYSFS_ERR_BAD_PASSWORD
        zip_prep_crypto_keys = 0
        Exit Function
    End If

    ' save the initial vector for seeking purposes. Not secure!!
    For i = 0 To 2
        fileInfos(finfo).initial_crypto_keys(i) = fileInfos(finfo).crypto_keys(i)
    Next
    zip_prep_crypto_keys = 1
End Function ' zip_prep_crypto_keys

'
' Bridge physfs allocation functions to zlib's format...: (VB) zlibPhysfsAlloc and zlibPhysfsFree are
' physfs_miniz_c's pool
'

'
' Construct a new z_stream to a sane state.
'
Private Sub initializeZStream(pstr As mz_stream)
    Dim blank As mz_stream
    pstr = blank                                        ' memset(pstr, '\0', sizeof (z_stream))
    ' (pstr->zalloc = zlibPhysfsAlloc; pstr->zfree = zlibPhysfsFree; pstr->opaque = &allocator;)
End Sub ' initializeZStream

Private Function zlib_error_code(ByVal rc As Long) As Long
    Select Case rc
    Case Z_OK: zlib_error_code = PHYSFS_ERR_OK          ' not an error.
    Case Z_STREAM_END: zlib_error_code = PHYSFS_ERR_OK  ' not an error.
    Case Z_ERRNO: zlib_error_code = PHYSFS_ERR_IO
    Case Z_MEM_ERROR: zlib_error_code = PHYSFS_ERR_OUT_OF_MEMORY
    Case Else: zlib_error_code = PHYSFS_ERR_CORRUPT
    End Select
End Function ' zlib_error_string

'
' Wrap all zlib calls in this, so the physfs error state is set appropriately.
'
Private Function zlib_err(ByVal rc As Long) As Long
    PHYSFS_setErrorCode zlib_error_code(rc)
    zlib_err = rc
End Function ' zlib_err

'
' Read an unsigned 64-bit int and swap to native byte order.
'
Private Function readui64(ByVal io As Long, val_ As Double) As Long
    Dim v(0 To 7) As Byte
    If PHYSFS_readAll__(io, v, 0, 8) = 0 Then Exit Function
    val_ = U32(v(0) Or (v(1) * &H100&) Or (v(2) * &H10000) Or SHL(v(3), 24)) + _
           U32(v(4) Or (v(5) * &H100&) Or (v(6) * &H10000) Or SHL(v(7), 24)) * TWO32     ' PHYSFS_swapULE64(v)
    readui64 = 1
End Function ' readui64

'
' Read an unsigned 32-bit int and swap to native byte order.
'
Private Function readui32(ByVal io As Long, val_ As Long) As Long
    Dim v(0 To 3) As Byte
    If PHYSFS_readAll__(io, v, 0, 4) = 0 Then Exit Function
    val_ = v(0) Or (v(1) * &H100&) Or (v(2) * &H10000) Or SHL(v(3), 24)    ' PHYSFS_swapULE32(v)
    readui32 = 1
End Function ' readui32

'
' Read an unsigned 16-bit int and swap to native byte order.
'
Private Function readui16(ByVal io As Long, val_ As Long) As Long
    Dim v(0 To 1) As Byte
    If PHYSFS_readAll__(io, v, 0, 2) = 0 Then Exit Function
    val_ = v(0) Or (v(1) * &H100&)                      ' PHYSFS_swapULE16(v)
    readui16 = 1
End Function ' readui16

' (VB) buf(bufOffset ..): the caller's buffer
Public Function ZIP_read(ByVal io_ As Long, buf() As Byte, ByVal bufOffset As Long, ByVal len_ As Long) As Long
    Dim finfo As Long, info As Long, entry As Long, retval As Double, maxread As Double, avail As Double
    Dim before As Double, rc As Long, br As Double
    finfo = PHYSFS_Io_opaque(io_)
    info = fileInfos(finfo).archive
    entry = fileInfos(finfo).entry
    retval = 0
    maxread = len_
    avail = zipInfos(info).entries(entry).uncompressed_size - fileInfos(finfo).uncompressed_position

    If avail < maxread Then
        maxread = avail
    End If

    If maxread = 0 Then                                 ' quick rejection.
        ZIP_read = 0
        Exit Function
    End If

    If zipInfos(info).entries(entry).compression_method = COMPMETH_NONE Then
        retval = zip_read_decrypt(finfo, buf, bufOffset, maxread)
    Else
        fileInfos(finfo).stream.next_out = bufOffset
        fileInfos(finfo).stream.avail_out = maxread

        Do While retval < maxread
            before = fileInfos(finfo).stream.total_out      ' (PHYSFS_uint32) finfo->stream.total_out

            If fileInfos(finfo).stream.avail_in = 0 Then
                br = zipInfos(info).entries(entry).compressed_size - fileInfos(finfo).compressed_position
                If br > 0 Then
                    If br > ZIP_READBUFSIZE Then
                        br = ZIP_READBUFSIZE
                    End If

                    br = zip_read_decrypt(finfo, fileInfos(finfo).buffer, 0, br)
                    If br <= 0 Then
                        Exit Do
                    End If

                    fileInfos(finfo).compressed_position = fileInfos(finfo).compressed_position + br
                    fileInfos(finfo).stream.next_in = 0     ' finfo->buffer
                    fileInfos(finfo).stream.avail_in = br
                End If
            End If

            rc = zlib_err(mz_inflate(fileInfos(finfo).stream, Z_SYNC_FLUSH, fileInfos(finfo).buffer, buf))
            retval = retval + (fileInfos(finfo).stream.total_out - before)

            If rc <> Z_OK Then
                Exit Do
            End If
        Loop
    End If

    If retval > 0 Then
        fileInfos(finfo).uncompressed_position = fileInfos(finfo).uncompressed_position + retval
    End If

    ZIP_read = retval
End Function ' ZIP_read

Public Function ZIP_write(ByVal io As Long, b() As Byte, ByVal bOffset As Long, ByVal len_ As Long) As Long
    PHYSFS_setErrorCode PHYSFS_ERR_READ_ONLY
    ZIP_write = -1
End Function ' ZIP_write

Public Function ZIP_tell(ByVal io As Long) As Double
    ZIP_tell = fileInfos(PHYSFS_Io_opaque(io)).uncompressed_position
End Function ' ZIP_tell

Public Function ZIP_seek(ByVal io_ As Long, ByVal offset As Double) As Long
    Dim finfo As Long, info As Long, entry As Long, io As Long, encrypted As Boolean, newpos As Double
    Dim str_ As mz_stream, buf(0 To 511) As Byte, maxread As Long, i As Long
    finfo = PHYSFS_Io_opaque(io_)
    info = fileInfos(finfo).archive
    entry = fileInfos(finfo).entry
    io = fileInfos(finfo).io
    encrypted = zip_entry_is_tradional_crypto(info, entry)

    If offset > zipInfos(info).entries(entry).uncompressed_size Then
        PHYSFS_setErrorCode PHYSFS_ERR_PAST_EOF
        ZIP_seek = 0
        Exit Function
    End If

    If Not encrypted And (zipInfos(info).entries(entry).compression_method = COMPMETH_NONE) Then
        newpos = offset + zipInfos(info).entries(entry).offset
        If PHYSFS_Io_seek(io, newpos) = 0 Then
            ZIP_seek = 0
            Exit Function
        End If
        fileInfos(finfo).uncompressed_position = offset
    Else
        '
        ' If seeking backwards, we need to redecode the file
        '  from the start and throw away the compressed bits until we hit
        '  the offset we need. If seeking forward, we still need to
        '  decode, but we don't rewind first.
        '
        If offset < fileInfos(finfo).uncompressed_position Then
            ' we do a copy so state is sane if inflateInit2() fails.
            initializeZStream str_
            If zlib_err(mz_inflateInit2(str_, -MAX_WBITS)) <> Z_OK Then
                ZIP_seek = 0
                Exit Function
            End If

            If encrypted Then newpos = 12 Else newpos = 0
            If PHYSFS_Io_seek(io, zipInfos(info).entries(entry).offset + newpos) = 0 Then
                ZIP_seek = 0
                Exit Function
            End If

            mz_inflateEnd fileInfos(finfo).stream
            fileInfos(finfo).stream = str_              ' memcpy(&finfo->stream, &str, sizeof (z_stream))
            fileInfos(finfo).uncompressed_position = 0
            fileInfos(finfo).compressed_position = 0

            If encrypted Then
                For i = 0 To 2
                    fileInfos(finfo).crypto_keys(i) = fileInfos(finfo).initial_crypto_keys(i)
                Next
            End If
        End If

        Do While fileInfos(finfo).uncompressed_position <> offset
            maxread = offset - fileInfos(finfo).uncompressed_position
            If maxread > 512 Then                       ' sizeof (buf)
                maxread = 512
            End If

            If ZIP_read(io_, buf, 0, maxread) <> maxread Then
                ZIP_seek = 0
                Exit Function
            End If
        Loop
    End If

    ZIP_seek = 1
End Function ' ZIP_seek

Public Function ZIP_length(ByVal io As Long) As Double
    Dim finfo As Long
    finfo = PHYSFS_Io_opaque(io)
    ZIP_length = zipInfos(fileInfos(finfo).archive).entries(fileInfos(finfo).entry).uncompressed_size
End Function ' ZIP_length

Public Function ZIP_duplicate(ByVal io As Long) As Long
    Dim origfinfo As Long, retval As Long, finfo As Long
    origfinfo = PHYSFS_Io_opaque(io)
    ' (PHYSFS_Io *retval = allocator.Malloc(sizeof (PHYSFS_Io)): made at the end, as a copy of io)
    finfo = newFileInfo()

    fileInfos(finfo).archive = fileInfos(origfinfo).archive
    fileInfos(finfo).entry = fileInfos(origfinfo).entry
    fileInfos(finfo).io = zip_get_io(fileInfos(origfinfo).io, 0, fileInfos(finfo).archive, fileInfos(finfo).entry)
    If fileInfos(finfo).io = 0 Then GoTo failed

    initializeZStream fileInfos(finfo).stream
    If zipInfos(fileInfos(finfo).archive).entries(fileInfos(finfo).entry).compression_method <> COMPMETH_NONE Then
        ReDim fileInfos(finfo).buffer(0 To ZIP_READBUFSIZE - 1)
        fileInfos(finfo).hasBuffer = True
        If zlib_err(mz_inflateInit2(fileInfos(finfo).stream, -MAX_WBITS)) <> Z_OK Then
            GoTo failed
        End If
    End If

    retval = PHYSFS_Io_new(PHYSFS_IO_ZIP, finfo)       ' memcpy(retval, io, sizeof (PHYSFS_Io)); retval->opaque = finfo;
    ZIP_duplicate = retval
    Exit Function

failed:
    If finfo <> 0 Then
        If fileInfos(finfo).io <> 0 Then
            PHYSFS_Io_destroy fileInfos(finfo).io
        End If

        If fileInfos(finfo).hasBuffer Then
            mz_inflateEnd fileInfos(finfo).stream
        End If

        freeFileInfo finfo
    End If

    ZIP_duplicate = 0
End Function ' ZIP_duplicate

Public Function ZIP_flush(ByVal io As Long) As Long
    ZIP_flush = 1   ' no write support.
End Function

Public Sub ZIP_destroy(ByVal io As Long)
    Dim finfo As Long
    finfo = PHYSFS_Io_opaque(io)
    PHYSFS_Io_destroy fileInfos(finfo).io

    If zipInfos(fileInfos(finfo).archive).entries(fileInfos(finfo).entry).compression_method <> COMPMETH_NONE Then
        mz_inflateEnd fileInfos(finfo).stream
    End If

    ' (if (finfo->buffer != NULL) allocator.Free(finfo->buffer);)
    freeFileInfo finfo
    ' (allocator.Free(io): physfs_win's)
End Sub ' ZIP_destroy

'
' (static const PHYSFS_Io ZIP_Io: physfs_win's PHYSFS_Io_* dispatch on PHYSFS_IO_ZIP)
'

Private Function zip_find_end_of_central_dir(ByVal io As Long, len_ As Double, ByVal hasLen As Boolean) As Double
    Dim buf(0 To 255) As Byte, extra(0 To 3) As Byte, i As Long, filelen As Double, filepos As Double
    Dim maxread As Long, totalread As Long, found As Long, k As Long

    filelen = PHYSFS_Io_length(io)
    If filelen = -1 Then
        zip_find_end_of_central_dir = -1
        Exit Function
    End If

    '
    ' Jump to the end of the file and start reading backwards.
    '  The last thing in the file is the zipfile comment, which is variable
    '  length, and the field that specifies its size is before it in the
    '  file (argh!)...this means that we need to scan backwards until we
    '  hit the end-of-central-dir signature. We can then sanity check that
    '  the comment was as big as it should be to make sure we're in the
    '  right place. The comment length field is 16 bits, so we can stop
    '  searching for that signature after a little more than 64k at most,
    '  and call it a corrupted zipfile.
    '

    If 256 < filelen Then                               ' sizeof (buf) < filelen
        filepos = filelen - 256
        maxread = 256
    Else
        filepos = 0
        maxread = filelen
    End If

    Do While (totalread < filelen) And (totalread < 65557)
        If PHYSFS_Io_seek(io, filepos) = 0 Then
            zip_find_end_of_central_dir = -1
            Exit Function
        End If

        ' make sure we catch a signature between buffers.
        If totalread <> 0 Then
            If PHYSFS_readAll__(io, buf, 0, maxread - 4) = 0 Then
                zip_find_end_of_central_dir = -1
                Exit Function
            End If
            For k = 0 To 3
                buf(maxread - 4 + k) = extra(k)         ' memcpy(&buf[maxread - 4], extra, sizeof (extra))
            Next
            totalread = totalread + maxread - 4
        Else
            If PHYSFS_readAll__(io, buf, 0, maxread) = 0 Then
                zip_find_end_of_central_dir = -1
                Exit Function
            End If
            totalread = totalread + maxread
        End If

        For k = 0 To 3
            extra(k) = buf(k)                           ' memcpy(extra, buf, sizeof (extra))
        Next

        i = maxread - 4
        Do While i > 0
            If (buf(i + 0) = &H50) And _
               (buf(i + 1) = &H4B) And _
               (buf(i + 2) = &H5) And _
               (buf(i + 3) = &H6) Then
                found = 1   ' that's the signature!
                Exit Do
            End If
            i = i - 1
        Loop

        If found <> 0 Then
            Exit Do
        End If

        filepos = filepos - (maxread - 4)
        If filepos < 0 Then
            filepos = 0
        End If
    Loop

    If found = 0 Then
        PHYSFS_setErrorCode PHYSFS_ERR_UNSUPPORTED
        zip_find_end_of_central_dir = -1
        Exit Function
    End If

    If hasLen Then
        len_ = filelen
    End If

    zip_find_end_of_central_dir = filepos + i
End Function ' zip_find_end_of_central_dir

Private Function isZip(ByVal io As Long) As Long
    Dim sig As Long, retval As Long, dummy As Double
    sig = 0
    retval = 0

    '
    ' The first thing in a zip file might be the signature of the
    '  first local file record, so it makes for a quick determination.
    '
    If readui32(io, sig) <> 0 Then
        If sig = ZIP_LOCAL_FILE_SIG Then retval = 1
        If retval = 0 Then
            '
            ' No sig...might be a ZIP with data at the start
            '  (a self-extracting executable, etc), so we'll have to do
            '  it the hard way...
            '
            If zip_find_end_of_central_dir(io, dummy, False) <> -1 Then retval = 1
        End If
    End If

    isZip = retval
End Function ' isZip

' Convert paths from old, buggy DOS zippers... (VB) path: the name's bytes, path(0 .. n - 1)
Private Sub zip_convert_dos_path(ByVal entryversion As Long, path() As Byte, ByVal n As Long)
    Dim hosttype As Long, i As Long
    hosttype = SHRU(entryversion, 8) And &HFF&
    If hosttype = 0 Then    ' FS_FAT_
        For i = 0 To n - 1                              ' while (*path)
            If path(i) = 0 Then Exit For
            If path(i) = &H5C Then                      ' '\\'
                path(i) = &H2F                          ' '/'
            End If
        Next
    End If
End Sub ' zip_convert_dos_path

' (VB) path in, the expanded path out
Private Sub zip_expand_symlink_path(path As String)
    Dim ptr As Long, prevptr As Long

    ptr = 1                                             ' (1-based positions in path)
    prevptr = 1

    Do
        ptr = InStr(ptr, path, "/")                     ' ptr = strchr(ptr, '/')
        If ptr = 0 Then
            Exit Do
        End If

        If Mid$(path, ptr + 1, 1) = "." Then
            If Mid$(path, ptr + 2, 1) = "/" Then
                ' current dir in middle of string: ditch it.
                path = Left$(path, ptr - 1) & Mid$(path, ptr + 2)       ' memmove(ptr, ptr + 2, strlen(ptr + 2) + 1)

            ElseIf ptr + 2 > Len(path) Then
                ' current dir at end of string: ditch it.
                path = Left$(path, ptr - 1)                             ' *ptr = '\0'

            ElseIf Mid$(path, ptr + 2, 1) = "." Then
                If Mid$(path, ptr + 3, 1) = "/" Then
                    ' parent dir in middle: move back one, if possible.
                    path = Left$(path, prevptr - 1) & Mid$(path, ptr + 4)   ' memmove(prevptr, ptr + 4, strlen(ptr + 4) + 1)
                    ptr = prevptr
                    Do While prevptr <> 1
                        prevptr = prevptr - 1
                        If Mid$(path, prevptr, 1) = "/" Then
                            prevptr = prevptr + 1
                            Exit Do
                        End If
                    Loop
                End If

                If ptr + 3 > Len(path) Then
                    ' parent dir at end: move back one, if possible.
                    path = Left$(path, prevptr - 1)                     ' *prevptr = '\0'
                End If
            End If
        Else
            prevptr = ptr
            ptr = ptr + 1
        End If
    Loop
End Sub ' zip_expand_symlink_path

Private Function zip_find_entry(ByVal info As Long, ByVal path As String) As Long
    zip_find_entry = PHYSFS_DirTreeFind__(info, path)
End Function ' zip_find_entry

' (forward reference: zip_follow_symlink and zip_resolve call each other.)

'
' Look for the entry named by (path). If it exists, resolve it, and return
'  a pointer to that entry. If it's another symlink, keep resolving until you
'  hit a real file and then return a pointer to the final non-symlink entry.
'  If there's a problem, return NULL.
'
Private Function zip_follow_symlink(ByVal io As Long, ByVal info As Long, ByVal path As String) As Long
    Dim entry As Long

    zip_expand_symlink_path path
    entry = zip_find_entry(info, path)
    If entry <> 0 Then
        If zip_resolve(io, info, entry) = 0 Then   ' recursive!
            entry = 0
        Else
            If zipInfos(info).entries(entry).symlink <> 0 Then
                entry = zipInfos(info).entries(entry).symlink
            End If
        End If
    End If

    zip_follow_symlink = entry
End Function ' zip_follow_symlink

Private Function zip_resolve_symlink(ByVal io As Long, ByVal info As Long, ByVal entry As Long) As Long
    Dim size As Long, path() As Byte, rc As Long, complen As Long, compressed() As Byte, stream As mz_stream
    Dim pathStr As String, n As Long
    size = zipInfos(info).entries(entry).uncompressed_size
    rc = 0

    '
    ' We've already parsed the local file header of the symlink at this
    '  point. Now we need to read the actual link from the file data and
    '  follow it.
    '

    If PHYSFS_Io_seek(io, zipInfos(info).entries(entry).offset) = 0 Then
        zip_resolve_symlink = 0
        Exit Function
    End If

    ReDim path(0 To size)                               ' __PHYSFS_smallAlloc(size + 1)

    If zipInfos(info).entries(entry).compression_method = COMPMETH_NONE Then
        rc = PHYSFS_readAll__(io, path, 0, size)

    Else    ' symlink target path is compressed...
        complen = zipInfos(info).entries(entry).compressed_size
        If complen > 0 Then ReDim compressed(0 To complen - 1) Else ReDim compressed(0 To 0)
        If PHYSFS_readAll__(io, compressed, 0, complen) <> 0 Then
            initializeZStream stream
            stream.next_in = 0                          ' compressed
            stream.avail_in = complen
            stream.next_out = 0                         ' path
            stream.avail_out = size
            If zlib_err(mz_inflateInit2(stream, -MAX_WBITS)) = Z_OK Then
                rc = zlib_err(mz_inflate(stream, Z_FINISH, compressed, path))
                mz_inflateEnd stream

                ' both are acceptable outcomes...
                If (rc = Z_OK) Or (rc = Z_STREAM_END) Then rc = 1 Else rc = 0
            End If
        End If
        ' (__PHYSFS_smallFree(compressed);)
    End If

    If rc <> 0 Then
        path(size) = 0                                  ' null-terminate it.
        zip_convert_dos_path zipInfos(info).entries(entry).version, path, size
        n = 0                                           ' (the C string ends at its first NUL)
        Do While path(n) <> 0
            n = n + 1
        Loop
        pathStr = utf8_decode(path, 0, n)
        n = zip_follow_symlink(io, info, pathStr)
        zipInfos(info).entries(entry).symlink = n
    End If

    ' (__PHYSFS_smallFree(path);)

    If zipInfos(info).entries(entry).symlink <> 0 Then zip_resolve_symlink = 1
End Function ' zip_resolve_symlink

'
' Parse the local file header of an entry, and update entry->offset.
'
Private Function zip_parse_local(ByVal io As Long, ByVal info As Long, ByVal entry As Long) As Long
    Dim ui32 As Long, ui16 As Long, fnamelen As Long, extralen As Long

    '
    ' crc and (un)compressed_size are always zero if this is a "JAR"
    '  archive created with Sun's Java tools, apparently. We only
    '  consider this archive corrupted if those entries don't match and
    '  aren't zero. That seems to work well.
    ' We also ignore a mismatch if the value is 0xFFFFFFFF here, since it's
    '  possible that's a Zip64 thing.
    '

    ' !!! FIXME: apparently these are zero if general purpose bit 3 is set,
    ' !!! FIXME:  which is probably true for Jar files, fwiw, but we don't
    ' !!! FIXME:  care about these values anyhow.

    If PHYSFS_Io_seek(io, zipInfos(info).entries(entry).offset) = 0 Then Exit Function
    If readui32(io, ui32) = 0 Then Exit Function
    If ui32 <> ZIP_LOCAL_FILE_SIG Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function
    If readui16(io, ui16) = 0 Then Exit Function
    ' Windows Explorer might rewrite the entire central directory, setting
    '   this field to 2.0/MS-DOS for all files, so favor the local version,
    '   which it leaves intact if it didn't alter that specific file.
    zipInfos(info).entries(entry).version_needed = ui16
    If readui16(io, ui16) = 0 Then Exit Function   ' general bits.
    If readui16(io, ui16) = 0 Then Exit Function
    If ui16 <> zipInfos(info).entries(entry).compression_method Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function
    If readui32(io, ui32) = 0 Then Exit Function   ' date/time
    If readui32(io, ui32) = 0 Then Exit Function
    If (ui32 <> 0) And (ui32 <> zipInfos(info).entries(entry).crc) Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    If readui32(io, ui32) = 0 Then Exit Function
    If (ui32 <> 0) And (ui32 <> &HFFFFFFFF) And _
       (U32(ui32) <> zipInfos(info).entries(entry).compressed_size) Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    If readui32(io, ui32) = 0 Then Exit Function
    If (ui32 <> 0) And (ui32 <> &HFFFFFFFF) And _
       (U32(ui32) <> zipInfos(info).entries(entry).uncompressed_size) Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    If readui16(io, fnamelen) = 0 Then Exit Function
    If readui16(io, extralen) = 0 Then Exit Function

    zipInfos(info).entries(entry).offset = zipInfos(info).entries(entry).offset + fnamelen + extralen + 30
    zip_parse_local = 1
End Function ' zip_parse_local

Private Function zip_resolve(ByVal io As Long, ByVal info As Long, ByVal entry As Long) As Long
    Dim retval As Long, resolve_type As Long
    retval = 1
    resolve_type = zipInfos(info).entries(entry).resolved

    If resolve_type = ZIP_DIRECTORY Then
        zip_resolve = 1     ' we're good.
        Exit Function
    End If

    ' Don't bother if we've failed to resolve this entry before.
    If resolve_type = ZIP_BROKEN_FILE Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function
    If resolve_type = ZIP_BROKEN_SYMLINK Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' uhoh...infinite symlink loop!
    If resolve_type = ZIP_RESOLVING Then PHYSFS_setErrorCode PHYSFS_ERR_SYMLINK_LOOP: Exit Function

    '
    ' We fix up the offset to point to the actual data on the
    '  first open, since we don't want to seek across the whole file on
    '  archive open (can be SLOW on large, CD-stored files), but we
    '  need to check the local file header...not just for corruption,
    '  but since it stores offset info the central directory does not.
    '
    If resolve_type <> ZIP_RESOLVED Then
        If zipInfos(info).entries(entry).tree_isdir <> 0 Then   ' an ancestor dir that DirTree filled in?
            zipInfos(info).entries(entry).resolved = ZIP_DIRECTORY
            zip_resolve = 1
            Exit Function
        End If

        retval = zip_parse_local(io, info, entry)
        If retval <> 0 Then
            '
            ' If it's a symlink, find the original file. This will cause
            '  resolution of other entries (other symlinks and, eventually,
            '  the real file) if all goes well.
            '
            If resolve_type = ZIP_UNRESOLVED_SYMLINK Then
                retval = zip_resolve_symlink(io, info, entry)
            End If
        End If

        If resolve_type = ZIP_UNRESOLVED_SYMLINK Then
            If retval <> 0 Then zipInfos(info).entries(entry).resolved = ZIP_RESOLVED Else zipInfos(info).entries(entry).resolved = ZIP_BROKEN_SYMLINK
        ElseIf resolve_type = ZIP_UNRESOLVED_FILE Then
            If retval <> 0 Then zipInfos(info).entries(entry).resolved = ZIP_RESOLVED Else zipInfos(info).entries(entry).resolved = ZIP_BROKEN_FILE
        End If
    End If

    zip_resolve = retval
End Function ' zip_resolve

Private Function zip_entry_is_symlink(ByVal info As Long, ByVal entry As Long) As Boolean
    zip_entry_is_symlink = (zipInfos(info).entries(entry).resolved = ZIP_UNRESOLVED_SYMLINK) Or _
                           (zipInfos(info).entries(entry).resolved = ZIP_BROKEN_SYMLINK) Or _
                           (zipInfos(info).entries(entry).symlink <> 0)
End Function ' zip_entry_is_symlink

Private Function zip_version_does_symlinks(ByVal version As Long) As Long
    Dim retval As Long, hosttype As Long
    retval = 0
    hosttype = SHRU(version, 8) And &HFF&

    Select Case hosttype
        '
        ' These are the platforms that can NOT build an archive with
        '  symlinks, according to the Info-ZIP project.
        '
    Case 0, 1, 2, 4, 6, 11, 14, 13, 15, 18      ' FS_FAT_, AMIGA_, VMS_, VM_CSM_, FS_HPFS_, FS_NTFS_, FS_VFAT_, ACORN_, MVS_, THEOS_
        ' do nothing.

    Case Else   ' assume the rest to be unix-like.
        retval = 1
    End Select

    zip_version_does_symlinks = retval
End Function ' zip_version_does_symlinks

Private Function zip_has_symlink_attr(ByVal info As Long, ByVal entry As Long, ByVal extern_attr As Long) As Boolean
    Dim xattr As Long
    xattr = SHRU(extern_attr, 16) And &HFFFF&
    If zip_version_does_symlinks(zipInfos(info).entries(entry).version) <> 0 Then
        If zipInfos(info).entries(entry).uncompressed_size > 0 Then
            zip_has_symlink_attr = ((xattr And UNIX_FILETYPE_MASK) = UNIX_FILETYPE_SYMLINK)
        End If
    End If
End Function ' zip_has_symlink_attr

' (VB) C's mktime() of a struct tm with tm_isdst = -1, as the Microsoft C runtime does it: the fields
' normalised (a day 0 is the month's last day before, and so on), the local time converted to UTC by the
' time zone's rules. Its result, or -1.
Private Function mktime_vb(ByVal tm_year As Long, ByVal tm_mon As Long, ByVal tm_mday As Long, ByVal tm_hour As Long, ByVal tm_min As Long, ByVal tm_sec As Long) As Double
    Dim y As Long, m As Long, days As Double, secs As Double, era As Long, yoe As Long, doy As Long, doe As Long
    Dim mp As Long, local_ As SYSTEMTIME_vb, utc As SYSTEMTIME_vb, ft As FILETIME_vb, d As Double
    y = tm_year + 1900 + Int(tm_mon / 12)
    m = tm_mon - Int(tm_mon / 12) * 12 + 1
    ' days from 1970-01-01 to y-m-01 (proleptic Gregorian)
    If m <= 2 Then y = y - 1
    era = Int(y / 400)
    yoe = y - era * 400
    If m > 2 Then mp = m - 3 Else mp = m + 9
    doy = (153 * mp + 2) \ 5
    doe = yoe * 365 + yoe \ 4 - yoe \ 100 + doy
    days = CDbl(era) * 146097 + doe - 719468
    secs = (days + tm_mday - 1) * 86400# + CDbl(tm_hour) * 3600# + CDbl(tm_min) * 60# + tm_sec
    ' the normalised local time
    days = Int(secs / 86400#)
    d = secs - days * 86400#
    ' civil date of days
    days = days + 719468
    era = Int(days / 146097)
    doe = days - CDbl(era) * 146097
    yoe = (doe - doe \ 1460 + doe \ 36524 - doe \ 146096) \ 365
    y = yoe + era * 400
    doy = doe - (365 * yoe + yoe \ 4 - yoe \ 100)
    mp = (5 * doy + 2) \ 153
    local_.wDay = doy - (153 * mp + 2) \ 5 + 1
    If mp < 10 Then local_.wMonth = mp + 3 Else local_.wMonth = mp - 9
    If local_.wMonth <= 2 Then y = y + 1
    If y < 1601 Or y > 30827 Then
        mktime_vb = -1
        Exit Function
    End If
    local_.wYear = y
    local_.wHour = Int(d / 3600)
    local_.wMinute = Int((d - local_.wHour * 3600#) / 60)
    local_.wSecond = d - local_.wHour * 3600# - local_.wMinute * 60#
    If TzSpecificLocalTimeToSystemTime(0, local_, utc) = 0 Then
        mktime_vb = -1
        Exit Function
    End If
    If SystemTimeToFileTime(utc, ft) = 0 Then
        mktime_vb = -1
        Exit Function
    End If
    mktime_vb = Int((U32(ft.dwHighDateTime) * TWO32 + U32(ft.dwLowDateTime)) / 10000000#) - 11644473600#
End Function

Private Function zip_dos_time_to_physfs_time(ByVal dostime As Long) As Double
    Dim dosdate As Long, tm_year As Long, tm_mon As Long, tm_mday As Long, tm_hour As Long, tm_min As Long, tm_sec As Long
    ' (memset(&unixtime, '\0', sizeof (unixtime));)

    dosdate = SHRU(dostime, 16) And &HFFFF&
    dostime = dostime And &HFFFF&

    ' dissect date
    tm_year = (SHRU(dosdate, 9) And &H7F) + 80
    tm_mon = (SHRU(dosdate, 5) And &HF) - 1
    tm_mday = (dosdate And &H1F)

    ' dissect time
    tm_hour = (SHRU(dostime, 11) And &H1F)
    tm_min = (SHRU(dostime, 5) And &H3F)
    tm_sec = ((dostime * 2) And &H3E)

    ' let mktime calculate daylight savings time.
    ' (unixtime.tm_isdst = -1;)

    zip_dos_time_to_physfs_time = mktime_vb(tm_year, tm_mon, tm_mday, tm_hour, tm_min, tm_sec)
End Function ' zip_dos_time_to_physfs_time

Private Function zip_load_entry(ByVal info As Long, ByVal zip64 As Long, ByVal ofs_fixup As Double) As Long
    Dim io As Long, entry As ZIPentry, retval As Long, fnamelen As Long, extralen As Long, commentlen As Long
    Dim external_attr As Long, starting_disk As Long, offset As Double, ui16 As Long, ui32 As Long, si64 As Double
    Dim nameBytes() As Byte, name_ As String, isdir As Long, found As Long, sig As Long, len_ As Long
    Dim cs As Double, us As Double
    io = zipInfos(info).io

    ' sanity check with central directory signature...
    If readui32(io, ui32) = 0 Then Exit Function
    If ui32 <> ZIP_CENTRAL_DIR_SIG Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' (memset(&entry, '\0', sizeof (entry));)

    ' Get the pertinent parts of the record...
    If readui16(io, entry.version) = 0 Then Exit Function
    If readui16(io, entry.version_needed) = 0 Then Exit Function
    If readui16(io, entry.general_bits) = 0 Then Exit Function    ' general bits
    If readui16(io, entry.compression_method) = 0 Then Exit Function
    If readui32(io, entry.dos_mod_time) = 0 Then Exit Function
    entry.last_mod_time = zip_dos_time_to_physfs_time(entry.dos_mod_time)
    If readui32(io, entry.crc) = 0 Then Exit Function
    If readui32(io, ui32) = 0 Then Exit Function
    entry.compressed_size = U32(ui32)
    If readui32(io, ui32) = 0 Then Exit Function
    entry.uncompressed_size = U32(ui32)
    If readui16(io, fnamelen) = 0 Then Exit Function
    If readui16(io, extralen) = 0 Then Exit Function
    If readui16(io, commentlen) = 0 Then Exit Function
    If readui16(io, ui16) = 0 Then Exit Function
    starting_disk = ui16
    If readui16(io, ui16) = 0 Then Exit Function   ' internal file attribs
    If readui32(io, external_attr) = 0 Then Exit Function
    If readui32(io, ui32) = 0 Then Exit Function
    offset = U32(ui32)

    ReDim nameBytes(0 To fnamelen)                      ' __PHYSFS_smallAlloc(fnamelen + 1)
    If PHYSFS_readAll__(io, nameBytes, 0, fnamelen) = 0 Then
        ' (__PHYSFS_smallFree(name);)
        Exit Function
    End If

    If fnamelen > 0 Then                                ' (VB) the C reads name[-1] for an empty name
        If nameBytes(fnamelen - 1) = &H2F Then          ' '/'
            nameBytes(fnamelen - 1) = 0
            isdir = 1
        End If
    End If
    nameBytes(fnamelen) = 0                             ' null-terminate the filename.

    zip_convert_dos_path entry.version, nameBytes, fnamelen

    ' (VB) the C string: the bytes up to the first NUL, as UTF-8
    ui32 = 0
    Do While nameBytes(ui32) <> 0
        ui32 = ui32 + 1
    Loop
    name_ = utf8_decode(nameBytes, 0, ui32)
    retval = PHYSFS_DirTreeAdd__(info, name_, isdir)
    ' (__PHYSFS_smallFree(name);)

    If retval = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_OUT_OF_MEMORY: Exit Function

    ' It's okay to BAIL without freeing retval, because it's stored in the
    '   __PHYSFS_DirTree and will be freed later anyhow.
    If zipInfos(info).entries(retval).last_mod_time <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function   ' dupe?

    ' Move the data we already read into place in the official object.
    zipInfos(info).entries(retval).symlink = entry.symlink
    zipInfos(info).entries(retval).resolved = entry.resolved
    zipInfos(info).entries(retval).offset = entry.offset
    zipInfos(info).entries(retval).version = entry.version
    zipInfos(info).entries(retval).version_needed = entry.version_needed
    zipInfos(info).entries(retval).general_bits = entry.general_bits
    zipInfos(info).entries(retval).compression_method = entry.compression_method
    zipInfos(info).entries(retval).crc = entry.crc
    zipInfos(info).entries(retval).compressed_size = entry.compressed_size
    zipInfos(info).entries(retval).uncompressed_size = entry.uncompressed_size
    zipInfos(info).entries(retval).last_mod_time = entry.last_mod_time
    zipInfos(info).entries(retval).dos_mod_time = entry.dos_mod_time

    zipInfos(info).entries(retval).symlink = 0  ' will be resolved later, if necessary.

    If isdir <> 0 Then
        zipInfos(info).entries(retval).resolved = ZIP_DIRECTORY
    Else
        If zip_has_symlink_attr(info, retval, external_attr) Then
            zipInfos(info).entries(retval).resolved = ZIP_UNRESOLVED_SYMLINK
        Else
            zipInfos(info).entries(retval).resolved = ZIP_UNRESOLVED_FILE
        End If
    End If

    si64 = PHYSFS_Io_tell(io)
    If si64 = -1 Then Exit Function

    ' If the actual sizes didn't fit in 32-bits, look for the Zip64
    '    extended information extra field...
    If (zip64 <> 0) And _
       ((offset = 4294967295#) Or _
        (starting_disk = &HFFFFFFFF) Or _
        (zipInfos(info).entries(retval).compressed_size = 4294967295#) Or _
        (zipInfos(info).entries(retval).uncompressed_size = 4294967295#)) Then
        found = 0
        sig = 0
        len_ = 0
        Do While extralen > 4
            If readui16(io, sig) = 0 Then Exit Function
            If readui16(io, len_) = 0 Then Exit Function

            si64 = si64 + 4 + len_
            extralen = (extralen - (4 + len_)) And &HFFFF&      ' PHYSFS_uint16
            If sig <> ZIP64_EXTENDED_INFO_EXTRA_FIELD_SIG Then
                If PHYSFS_Io_seek(io, si64) = 0 Then Exit Function
                GoTo continue_extra
            End If

            found = 1
            Exit Do
continue_extra:
        Loop

        If found = 0 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

        If zipInfos(info).entries(retval).uncompressed_size = 4294967295# Then
            If len_ < 8 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function
            If readui64(io, us) = 0 Then Exit Function
            zipInfos(info).entries(retval).uncompressed_size = us
            len_ = len_ - 8
        End If

        If zipInfos(info).entries(retval).compressed_size = 4294967295# Then
            If len_ < 8 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function
            If readui64(io, cs) = 0 Then Exit Function
            zipInfos(info).entries(retval).compressed_size = cs
            len_ = len_ - 8
        End If

        If offset = 4294967295# Then
            If len_ < 8 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function
            If readui64(io, offset) = 0 Then Exit Function
            len_ = len_ - 8
        End If

        If starting_disk = &HFFFFFFFF Then
            If len_ < 8 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function
            If readui32(io, starting_disk) = 0 Then Exit Function
            len_ = len_ - 4
        End If

        If len_ <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function
    End If

    If starting_disk <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    zipInfos(info).entries(retval).offset = offset + ofs_fixup

    ' seek to the start of the next entry in the central directory...
    If PHYSFS_Io_seek(io, si64 + extralen + commentlen) = 0 Then Exit Function

    zip_load_entry = retval  ' success.
End Function ' zip_load_entry

' This leaves things allocated on error; the caller will clean up the mess.
Private Function zip_load_entries(ByVal info As Long, ByVal data_ofs As Double, ByVal central_ofs As Double, ByVal entry_count As Double) As Long
    Dim io As Long, zip64 As Long, i As Double, entry As Long
    io = zipInfos(info).io
    zip64 = zipInfos(info).zip64

    If PHYSFS_Io_seek(io, central_ofs) = 0 Then Exit Function

    i = 0
    Do While i < entry_count
        entry = zip_load_entry(info, zip64, data_ofs)
        If entry = 0 Then Exit Function
        If zip_entry_is_tradional_crypto(info, entry) Then
            zipInfos(info).has_crypto = 1
        End If
        i = i + 1
    Loop

    zip_load_entries = 1
End Function ' zip_load_entries

Private Function zip64_find_end_of_central_dir(ByVal io As Long, ByVal pos_ As Double, ByVal offset As Double) As Double
    '
    ' Naturally, the offset is useless to us; it is the offset from the
    '  start of file, which is meaningless if we've appended this .zip to
    '  a self-extracting .exe. We need to find this on our own. It should
    '  be directly before the locator record, but the record in question,
    '  like the original end-of-central-directory record, ends with a
    '  variable-length field. Unlike the original, which has to store the
    '  size of that variable-length field in a 16-bit int and thus has to be
    '  within 64k, the new one gets 64-bits.
    '
    ' Fortunately, the only currently-specified record for that variable
    '  length block is some weird proprietary thing that deals with EBCDIC
    '  and tape backups or something. So we don't seek far.
    '

    Dim ui32 As Long, pos As Double, maxbuflen As Long, len_ As Long, buf() As Byte, i As Long
    pos = pos_

    ' (assert(_pos > 0);)

    ' Try offset specified in the Zip64 end of central directory locator.
    ' This works if the entire PHYSFS_Io is the zip file.
    If PHYSFS_Io_seek(io, offset) = 0 Then zip64_find_end_of_central_dir = -1: Exit Function
    If readui32(io, ui32) = 0 Then zip64_find_end_of_central_dir = -1: Exit Function
    If ui32 = ZIP64_END_OF_CENTRAL_DIR_SIG Then
        zip64_find_end_of_central_dir = offset
        Exit Function
    End If

    ' Try 56 bytes before the Zip64 end of central directory locator.
    ' This works if the record isn't variable length and is version 1.
    If pos > 56 Then
        If PHYSFS_Io_seek(io, pos - 56) = 0 Then zip64_find_end_of_central_dir = -1: Exit Function
        If readui32(io, ui32) = 0 Then zip64_find_end_of_central_dir = -1: Exit Function
        If ui32 = ZIP64_END_OF_CENTRAL_DIR_SIG Then
            zip64_find_end_of_central_dir = pos - 56
            Exit Function
        End If
    End If

    ' Try 84 bytes before the Zip64 end of central directory locator.
    ' This works if the record isn't variable length and is version 2.
    If pos > 84 Then
        If PHYSFS_Io_seek(io, pos - 84) = 0 Then zip64_find_end_of_central_dir = -1: Exit Function
        If readui32(io, ui32) = 0 Then zip64_find_end_of_central_dir = -1: Exit Function
        If ui32 = ZIP64_END_OF_CENTRAL_DIR_SIG Then
            zip64_find_end_of_central_dir = pos - 84
            Exit Function
        End If
    End If

    ' Ok, brute force: we know it's between (offset) and (pos) somewhere.
    '  Just try moving back at most 256k. Oh well.
    If (offset < pos) And (pos > 4) Then
        maxbuflen = 256& * 1024
        If pos - offset > maxbuflen Then
            len_ = maxbuflen
        Else
            len_ = pos - offset
        End If

        ReDim buf(0 To len_ - 1)                        ' __PHYSFS_smallAlloc(len)

        If PHYSFS_Io_seek(io, pos - len_) = 0 Then
            zip64_find_end_of_central_dir = -1          ' error was set elsewhere.
            Exit Function
        End If
        If PHYSFS_readAll__(io, buf, 0, len_) = 0 Then
            zip64_find_end_of_central_dir = -1          ' error was set elsewhere.
            Exit Function
        End If

        For i = len_ - 4 To 0 Step -1
            If (buf(i) = &H50) And (buf(i + 1) = &H4B) And _
               (buf(i + 2) = &H6) And (buf(i + 3) = &H6) Then
                zip64_find_end_of_central_dir = pos - (len_ - i)
                Exit Function
            End If
        Next
    End If

    PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT              ' didn't find it.
    zip64_find_end_of_central_dir = -1
End Function ' zip64_find_end_of_central_dir

Private Function zip64_parse_end_of_central_dir(ByVal info As Long, data_start As Double, dir_ofs As Double, entry_count As Double, ByVal pos As Double) As Long
    Dim io As Long, ui64 As Double, ui32 As Long, ui16 As Long
    io = zipInfos(info).io

    ' We should be positioned right past the locator signature.

    If pos < 0 Then Exit Function
    If PHYSFS_Io_seek(io, pos) = 0 Then Exit Function

    If readui32(io, ui32) = 0 Then Exit Function
    If ui32 <> ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR_SIG Then
        zip64_parse_end_of_central_dir = -1     ' it's not a Zip64 archive. Not an error, though!
        Exit Function
    End If

    zipInfos(info).zip64 = 1

    ' number of the disk with the start of the central directory.
    If readui32(io, ui32) = 0 Then Exit Function
    If ui32 <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' offset of Zip64 end of central directory record.
    If readui64(io, ui64) = 0 Then Exit Function

    ' total number of disks
    If readui32(io, ui32) = 0 Then Exit Function
    If ui32 <> 1 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    pos = zip64_find_end_of_central_dir(io, pos, ui64)
    If pos < 0 Then
        Exit Function   ' oh well.
    End If

    '
    ' For self-extracting archives, etc, there's crapola in the file
    '  before the zipfile records; we calculate how much data there is
    '  prepended by determining how far the zip64-end-of-central-directory
    '  offset is from where it is supposed to be...the difference in bytes
    '  is how much arbitrary data is at the start of the physical file.
    '
    ' (assert(((PHYSFS_uint64) pos) >= ui64);)
    data_start = pos - ui64

    If PHYSFS_Io_seek(io, pos) = 0 Then Exit Function

    ' check signature again, just in case.
    If readui32(io, ui32) = 0 Then Exit Function
    If ui32 <> ZIP64_END_OF_CENTRAL_DIR_SIG Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' size of Zip64 end of central directory record.
    If readui64(io, ui64) = 0 Then Exit Function

    ' version made by.
    If readui16(io, ui16) = 0 Then Exit Function

    ' version needed to extract.
    If readui16(io, ui16) = 0 Then Exit Function

    ' number of this disk.
    If readui32(io, ui32) = 0 Then Exit Function
    If ui32 <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' number of disk with start of central directory record.
    If readui32(io, ui32) = 0 Then Exit Function
    If ui32 <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' total number of entries in the central dir on this disk
    If readui64(io, ui64) = 0 Then Exit Function

    ' total number of entries in the central dir
    If readui64(io, entry_count) = 0 Then Exit Function
    If ui64 <> entry_count Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' size of the central directory
    If readui64(io, ui64) = 0 Then Exit Function

    ' offset of central directory
    If readui64(io, dir_ofs) = 0 Then Exit Function

    ' Since we know the difference, fix up the central dir offset...
    dir_ofs = dir_ofs + data_start

    '
    ' There are more fields here, for encryption and feature-specific things,
    '  but we don't care about any of them at the moment.
    '

    zip64_parse_end_of_central_dir = 1  ' made it.
End Function ' zip64_parse_end_of_central_dir

Private Function zip_parse_end_of_central_dir(ByVal info As Long, data_start As Double, dir_ofs As Double, entry_count As Double) As Long
    Dim io As Long, entryCount16 As Long, offset32 As Long, ui32 As Long, ui16 As Long, len_ As Double, pos As Double, rc As Long
    io = zipInfos(info).io

    ' find the end-of-central-dir record, and seek to it.
    pos = zip_find_end_of_central_dir(io, len_, True)
    If pos = -1 Then Exit Function
    If PHYSFS_Io_seek(io, pos) = 0 Then Exit Function

    ' check signature again, just in case.
    If readui32(io, ui32) = 0 Then Exit Function
    If ui32 <> ZIP_END_OF_CENTRAL_DIR_SIG Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' Seek back to see if "Zip64 end of central directory locator" exists.
    ' this record is 20 bytes before end-of-central-dir
    rc = zip64_parse_end_of_central_dir(info, data_start, dir_ofs, entry_count, pos - 20)

    ' Error or success? Bounce out of here. Keep going if not zip64.
    If (rc = 0) Or (rc = 1) Then
        zip_parse_end_of_central_dir = rc
        Exit Function
    End If

    ' (assert(rc == -1);  /* no error, just not a Zip64 archive. */)

    ' Not Zip64? Seek back to where we were and keep processing.
    If PHYSFS_Io_seek(io, pos + 4) = 0 Then Exit Function

    ' number of this disk
    If readui16(io, ui16) = 0 Then Exit Function
    If ui16 <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' number of the disk with the start of the central directory
    If readui16(io, ui16) = 0 Then Exit Function
    If ui16 <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    ' total number of entries in the central dir on this disk
    If readui16(io, ui16) = 0 Then Exit Function

    ' total number of entries in the central dir
    If readui16(io, entryCount16) = 0 Then Exit Function
    If ui16 <> entryCount16 Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    entry_count = entryCount16

    ' size of the central directory
    If readui32(io, ui32) = 0 Then Exit Function

    ' offset of central directory
    If readui32(io, offset32) = 0 Then Exit Function
    dir_ofs = U32(offset32)
    If pos < (dir_ofs + U32(ui32)) Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    '
    ' For self-extracting archives, etc, there's crapola in the file
    '  before the zipfile records; we calculate how much data there is
    '  prepended by determining how far the central directory offset is
    '  from where it is supposed to be (start of end-of-central-dir minus
    '  sizeof central dir)...the difference in bytes is how much arbitrary
    '  data is at the start of the physical file.
    '
    data_start = pos - (dir_ofs + U32(ui32))

    ' Now that we know the difference, fix up the central dir offset...
    dir_ofs = dir_ofs + data_start

    ' zipfile comment length
    If readui16(io, ui16) = 0 Then Exit Function

    '
    ' Make sure that the comment length matches to the end of file...
    '  If it doesn't, we're either in the wrong part of the file, or the
    '  file is corrupted, but we give up either way.
    '
    If (pos + 22 + ui16) <> len_ Then PHYSFS_setErrorCode PHYSFS_ERR_CORRUPT: Exit Function

    zip_parse_end_of_central_dir = 1    ' made it.
End Function ' zip_parse_end_of_central_dir

Public Sub ZIP_closeArchive(ByVal opaque As Long)
    Dim info As Long
    info = opaque

    If info = 0 Then
        Exit Sub
    End If

    If zipInfos(info).io <> 0 Then
        PHYSFS_Io_destroy zipInfos(info).io
    End If

    PHYSFS_DirTreeDeinit__ info

    zipInfos(info).inUse = False                        ' allocator.Free(info)
    zipInfos(info).io = 0
End Sub ' ZIP_closeArchive

' (VB) name: the archive's path (unused, as in the C); claimed set to 1 when the archive is a zip
Public Function ZIP_openArchive(ByVal io As Long, ByVal name_ As String, ByVal forWriting As Long, claimed As Long) As Long
    Dim info As Long, root As Long, dstart As Double, cdir_ofs As Double, count As Double
    dstart = 0  ' data start

    ' (assert(io != NULL);  /* shouldn't ever happen. */)

    If forWriting <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_READ_ONLY: Exit Function
    If isZip(io) = 0 Then Exit Function

    claimed = 1

    info = newZipInfo()                                 ' allocator.Malloc; memset

    zipInfos(info).io = io

    If zip_parse_end_of_central_dir(info, dstart, cdir_ofs, count) = 0 Then
        GoTo ZIP_openarchive_failed
    ElseIf PHYSFS_DirTreeInit__(info, 1, 0) = 0 Then
        GoTo ZIP_openarchive_failed
    End If

    root = zipInfos(info).tree_root
    zipInfos(info).entries(root).resolved = ZIP_DIRECTORY

    If zip_load_entries(info, dstart, cdir_ofs, count) = 0 Then
        GoTo ZIP_openarchive_failed
    End If

    ' (assert(info->tree.root->sibling == NULL);)
    ZIP_openArchive = info
    Exit Function

ZIP_openarchive_failed:
    zipInfos(info).io = 0  ' don't let ZIP_closeArchive destroy (io).
    ZIP_closeArchive info
    ZIP_openArchive = 0
End Function ' ZIP_openArchive

' (VB) inf: the ZIPinfo to resolve the entry with (0 = NULL), archive: the entry's
Private Function zip_get_io(ByVal io As Long, ByVal inf As Long, ByVal archive As Long, ByVal entry As Long) As Long
    Dim success As Long, retval As Long, offset As Double
    retval = PHYSFS_Io_duplicate(io)
    If retval = 0 Then Exit Function

    ' (assert(!entry->tree.isdir); /* should have been checked before calling. */)

    ' (inf) can be NULL if we already resolved.
    If inf = 0 Then
        success = 1
    Else
        success = zip_resolve(retval, inf, entry)
    End If
    If success <> 0 Then
        If zipInfos(archive).entries(entry).symlink <> 0 Then
            offset = zipInfos(archive).entries(zipInfos(archive).entries(entry).symlink).offset
        Else
            offset = zipInfos(archive).entries(entry).offset
        End If
        success = PHYSFS_Io_seek(retval, offset)
    End If

    If success = 0 Then
        PHYSFS_Io_destroy retval
        retval = 0
    End If

    zip_get_io = retval
End Function ' zip_get_io

Public Function ZIP_openRead(ByVal opaque As Long, ByVal filename As String) As Long
    Dim retval As Long, info As Long, entry As Long, finfo As Long, io As Long, password() As Byte, passwordLen As Long
    Dim hasPassword As Boolean, ptr As Long, crypto_header(0 To 11) As Byte
    info = opaque
    entry = zip_find_entry(info, filename)

    ' if not found, see if maybe "$PASSWORD" is appended.
    If (entry = 0) And (zipInfos(info).has_crypto <> 0) Then
        ptr = InStrRev(filename, "$")
        If ptr > 0 Then
            entry = zip_find_entry(info, Left$(filename, ptr - 1))
            passwordLen = utf8_encode(Mid$(filename, ptr + 1), password)
            hasPassword = True
        End If
    End If

    If entry = 0 Then Exit Function

    If zip_resolve(zipInfos(info).io, info, entry) = 0 Then Exit Function

    If zipInfos(info).entries(entry).tree_isdir <> 0 Then PHYSFS_setErrorCode PHYSFS_ERR_NOT_A_FILE: Exit Function

    ' (retval = (PHYSFS_Io *) allocator.Malloc(sizeof (PHYSFS_Io)): made at the end)

    finfo = newFileInfo()

    io = zip_get_io(zipInfos(info).io, info, info, entry)
    If io = 0 Then GoTo ZIP_openRead_failed
    fileInfos(finfo).io = io
    fileInfos(finfo).archive = info
    If zipInfos(info).entries(entry).symlink <> 0 Then
        fileInfos(finfo).entry = zipInfos(info).entries(entry).symlink
    Else
        fileInfos(finfo).entry = entry
    End If
    initializeZStream fileInfos(finfo).stream

    If zipInfos(info).entries(fileInfos(finfo).entry).compression_method <> COMPMETH_NONE Then
        ReDim fileInfos(finfo).buffer(0 To ZIP_READBUFSIZE - 1)
        fileInfos(finfo).hasBuffer = True
        If zlib_err(mz_inflateInit2(fileInfos(finfo).stream, -MAX_WBITS)) <> Z_OK Then
            GoTo ZIP_openRead_failed
        End If
    End If

    If Not zip_entry_is_tradional_crypto(info, entry) Then
        If hasPassword Then PHYSFS_setErrorCode PHYSFS_ERR_BAD_PASSWORD: GoTo ZIP_openRead_failed
    Else
        If Not hasPassword Then PHYSFS_setErrorCode PHYSFS_ERR_BAD_PASSWORD: GoTo ZIP_openRead_failed
        If PHYSFS_Io_read(io, crypto_header, 0, 12) <> 12 Then
            GoTo ZIP_openRead_failed
        ElseIf zip_prep_crypto_keys(finfo, crypto_header, password, passwordLen) = 0 Then
            GoTo ZIP_openRead_failed
        End If
    End If

    retval = PHYSFS_Io_new(PHYSFS_IO_ZIP, finfo)       ' memcpy(retval, &ZIP_Io, sizeof (PHYSFS_Io)); retval->opaque = finfo;

    ZIP_openRead = retval
    Exit Function

ZIP_openRead_failed:
    If finfo <> 0 Then
        If fileInfos(finfo).io <> 0 Then
            PHYSFS_Io_destroy fileInfos(finfo).io
        End If

        If fileInfos(finfo).hasBuffer Then
            mz_inflateEnd fileInfos(finfo).stream
        End If

        freeFileInfo finfo
    End If

    ZIP_openRead = 0
End Function ' ZIP_openRead

Public Function ZIP_openWrite(ByVal opaque As Long, ByVal filename As String) As Long
    PHYSFS_setErrorCode PHYSFS_ERR_READ_ONLY
End Function ' ZIP_openWrite

Public Function ZIP_openAppend(ByVal opaque As Long, ByVal filename As String) As Long
    PHYSFS_setErrorCode PHYSFS_ERR_READ_ONLY
End Function ' ZIP_openAppend

Public Function ZIP_remove(ByVal opaque As Long, ByVal name_ As String) As Long
    PHYSFS_setErrorCode PHYSFS_ERR_READ_ONLY
End Function ' ZIP_remove

Public Function ZIP_mkdir(ByVal opaque As Long, ByVal name_ As String) As Long
    PHYSFS_setErrorCode PHYSFS_ERR_READ_ONLY
End Function ' ZIP_mkdir

Public Function ZIP_stat(ByVal opaque As Long, ByVal filename As String, stat As PHYSFS_Stat_) As Long
    Dim info As Long, entry As Long
    info = opaque
    entry = zip_find_entry(info, filename)

    If entry = 0 Then
        ZIP_stat = 0
        Exit Function

    ElseIf zip_resolve(zipInfos(info).io, info, entry) = 0 Then
        ZIP_stat = 0
        Exit Function

    ElseIf zipInfos(info).entries(entry).resolved = ZIP_DIRECTORY Then
        stat.filesize = 0
        stat.filetype = PHYSFS_FILETYPE_DIRECTORY

    ElseIf zip_entry_is_symlink(info, entry) Then
        stat.filesize = 0
        stat.filetype = PHYSFS_FILETYPE_SYMLINK

    Else
        stat.filesize = zipInfos(info).entries(entry).uncompressed_size
        stat.filetype = PHYSFS_FILETYPE_REGULAR
    End If

    stat.modtime = zipInfos(info).entries(entry).last_mod_time     ' ((entry) ? entry->last_mod_time : 0)
    stat.createtime = stat.modtime
    stat.accesstime = -1
    stat.readonly_ = 1   ' .zip files are always read only

    ZIP_stat = 1
End Function ' ZIP_stat

' __PHYSFS_Archiver_ZIP's enumerate: __PHYSFS_DirTreeEnumerate
Public Function ZIP_enumerate(ByVal opaque As Long, ByVal dname As String, ByVal cb As Long, ByVal origdir As String) As Long
    ZIP_enumerate = PHYSFS_DirTreeEnumerate__(opaque, dname, cb, origdir)
End Function

'
' const PHYSFS_Archiver __PHYSFS_Archiver_ZIP =
' {
'     CURRENT_PHYSFS_ARCHIVER_API_VERSION,
'     {
'         "ZIP",
'         "PkZip/WinZip/Info-Zip compatible",
'         "Ryan C. Gordon <icculus@icculus.org>",
'         "https://icculus.org/physfs/",
'         1,  /* supportsSymlinks */
'     },
'     ZIP_openArchive,
'     __PHYSFS_DirTreeEnumerate,
'     ZIP_openRead,
'     ZIP_openWrite,
'     ZIP_openAppend,
'     ZIP_remove,
'     ZIP_mkdir,
'     ZIP_stat,
'     ZIP_closeArchive
' };
' (VB: physfs_win's archiver dispatch, ARCHIVER_ZIP)
'
