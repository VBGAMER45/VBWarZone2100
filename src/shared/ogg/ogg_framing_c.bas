Attribute VB_Name = "ogg_framing_c"
' libogg 1.3.6, src/framing.c (and crctable.h): code raw packets into framed OggSquish stream and decode Ogg
' streams back into raw packets; whole, the encode and the decode sides (bitwise.c is not ported: neither
' opusfile nor the videos use oggpack). The types are ogg_h's: the C's pointers into the sync buffer and the
' stream state's body are copies here, made where the C sets the pointer (ogg_sync_pageseek and
' ogg_stream_flush_i fill og's header and body, _packetout op's packet), so they stay valid as long as the
' caller keeps them. ogg_sync_buffer returns the offset into oy.data to write at (-1 = NULL). An
' ogg_int64_t is a Double (granule positions and packet numbers), exact to 2^53. A NULL ogg_page * or
' ogg_packet * argument is the Optional ogIsNull / opIsNull. crctable.h's crc_lookup is computed at the first
' use by the generator framing.c keeps under #if 0 (_ogg_crc_init), which gives the same 2048 values.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const LONG_MAX As Long = 2147483647
Private Const INT_MAX As Long = 2147483647

' crctable.h: static const ogg_uint32_t crc_lookup[8][256] (crc_lookup[j][i] is crc_lookup(j * 256 + i))
Private crc_lookup(0 To 2047) As Long
Private crcReady As Boolean

' A complete description of Ogg framing exists in docs/framing.html

Public Function ogg_page_version(og As ogg_page) As Long
    ogg_page_version = og.header(4)
End Function

Public Function ogg_page_continued(og As ogg_page) As Long
    ogg_page_continued = og.header(5) And &H1
End Function

Public Function ogg_page_bos(og As ogg_page) As Long
    ogg_page_bos = og.header(5) And &H2
End Function

Public Function ogg_page_eos(og As ogg_page) As Long
    ogg_page_eos = og.header(5) And &H4
End Function

' (VB) the 64-bit value as a Double: the high 32 bits signed, the low 32 bits unsigned (exact to 2^53; -1 is -1)
Public Function ogg_page_granulepos(og As ogg_page) As Double
    Dim hi As Double, lo As Double
    ' unsigned char *page=og->header; ogg_uint64_t granulepos=page[13]&(0xff); granulepos=(granulepos<<8)|...
    hi = og.header(13) * 16777216# + og.header(12) * 65536# + og.header(11) * 256# + og.header(10)
    If og.header(13) >= 128 Then hi = hi - 4294967296#
    lo = og.header(9) * 16777216# + og.header(8) * 65536# + og.header(7) * 256# + og.header(6)
    ogg_page_granulepos = hi * 4294967296# + lo             ' return((ogg_int64_t)granulepos)
End Function

Public Function ogg_page_serialno(og As ogg_page) As Long
    ogg_page_serialno = le32(og.header, 14)
End Function

Public Function ogg_page_pageno(og As ogg_page) As Long
    ogg_page_pageno = le32(og.header, 18)
End Function

' (VB) the 32 bits of b(i .. i + 3), little endian (C: (ogg_uint32_t)b[i] | (ogg_uint32_t)b[i+1]<<8 | ...)
Private Function le32(b() As Byte, ByVal i As Long) As Long
    If b(i + 3) And &H80 Then
        le32 = ((b(i + 3) And &H7F) * &H1000000) Or &H80000000
    Else
        le32 = b(i + 3) * &H1000000
    End If
    le32 = le32 Or (b(i + 2) * &H10000) Or (b(i + 1) * &H100&) Or b(i)
End Function

' returns the number of packets that are completed on this page (if
'   the leading packet is begun on a previous page, but ends on this
'   page, it's counted

' NOTE:
'   If a page consists of a packet begun on a previous page, and a new
'   packet begun (but not completed) on this page, the return will be:
'     ogg_page_packets(page)   ==1,
'     ogg_page_continued(page) !=0
'
'   If a page happens to be a single packet that was begun on a
'   previous page, and spans to the next page (in the case of a three or
'   more page packet), the return will be:
'     ogg_page_packets(page)   ==0,
'     ogg_page_continued(page) !=0

Public Function ogg_page_packets(og As ogg_page) As Long
    Dim i As Long, n As Long, count As Long
    n = og.header(26)
    count = 0
    For i = 0 To n - 1
        If og.header(27 + i) < 255 Then count = count + 1
    Next
    ogg_page_packets = count
End Function

' #if 0: helper to initialize lookup for direct-table CRC (illustrative; we
'   use the static init in crctable.h). (VB) the port uses it to fill crc_lookup at the first use: it gives
'   crctable.h's values (checked with two of them).
Private Sub ogg_crc_init_()
    Dim i As Long, j As Long, polynomial As Long, crc As Long, prev As Long
    polynomial = &H4C11DB7     ' The same as the ethernet generator
                               ' polynomial, although we use an
                               ' unreflected alg and an init/final
                               ' of 0, not 0xffffffff
    For i = 0 To &HFF
        ' crc = i << 24
        If i And &H80 Then crc = ((i And &H7F) * &H1000000) Or &H80000000 Else crc = i * &H1000000

        For j = 0 To 7
            ' crc = (crc << 1) ^ (crc & (1 << 31) ? polynomial : 0)
            If crc < 0 Then
                crc = shl1(crc) Xor polynomial
            Else
                crc = shl1(crc)
            End If
        Next

        crc_lookup(i) = crc
    Next

    For i = 0 To &HFF
        For j = 1 To 7
            prev = crc_lookup((j - 1) * 256 + i)
            crc_lookup(j * 256 + i) = crc_lookup(((prev And &HFF000000) \ &H1000000) And &HFF) Xor shl8(prev)
        Next
    Next
    If crc_lookup(255) <> &HB1F740B4 Or crc_lookup(7 * 256 + 255) <> &H6760D264 Then
        debug_ 0, LOG_ERROR, "ogg_crc_init_", "(VB) the generated crc_lookup is wrong: crc_lookup[0][255] = 0x%08X", crc_lookup(255)
    End If
    crcReady = True
End Sub

' (VB) x << 1 and x << 8 for an ogg_uint32_t (the bits shifted out are lost)
Private Function shl1(ByVal x As Long) As Long
    If x And &H40000000 Then shl1 = ((x And &H3FFFFFFF) * 2) Or &H80000000 Else shl1 = (x And &H3FFFFFFF) * 2
End Function

Private Function shl8(ByVal x As Long) As Long
    If x And &H800000 Then shl8 = ((x And &H7FFFFF) * &H100&) Or &H80000000 Else shl8 = (x And &H7FFFFF) * &H100&
End Function

' init the encode/decode logical stream state

Public Function ogg_stream_init(os As ogg_stream_state, ByVal serialno As Long) As Long
    Dim blank As ogg_stream_state
    ' (VB) if(os): a UDT is never NULL
    os = blank                                  ' memset(os,0,sizeof(*os))
    os.body_storage = 16 * 1024&
    os.lacing_storage = 1024

    ReDim os.body_data(0 To os.body_storage - 1)
    ReDim os.lacing_vals(0 To os.lacing_storage - 1)
    ReDim os.granule_vals(0 To os.lacing_storage - 1)
    os.vbInit = True
    ' (VB: ReDim raises an error where _ogg_malloc returns NULL: no ogg_stream_clear and -1 here)

    os.serialno = serialno

    ogg_stream_init = 0
End Function

' async/delayed error detection for the ogg_stream_state
Public Function ogg_stream_check(os As ogg_stream_state) As Long
    If Not os.vbInit Then ogg_stream_check = -1: Exit Function     ' if(!os || !os->body_data) return -1
    ogg_stream_check = 0
End Function

' _clear does not free os, only the non-flat storage within
Public Function ogg_stream_clear(os As ogg_stream_state) As Long
    Dim blank As ogg_stream_state
    ' if(os->body_data)_ogg_free(os->body_data); ... memset(os,0,sizeof(*os))
    os = blank
    ogg_stream_clear = 0
End Function

' (VB) the struct itself is the caller's: as ogg_stream_clear
Public Function ogg_stream_destroy(os As ogg_stream_state) As Long
    ogg_stream_clear os
    ' (_ogg_free(os): the caller's variable)
    ogg_stream_destroy = 0
End Function

' Helpers for ogg_stream_encode; this keeps the structure and
'   what's happening fairly clear

Private Function os_body_expand_(os As ogg_stream_state, ByVal needed As Long) As Long
    Dim body_storage As Long
    If os.body_storage - needed <= os.body_fill Then
        If os.body_storage > LONG_MAX - needed Then
            ogg_stream_clear os
            os_body_expand_ = -1
            Exit Function
        End If
        body_storage = os.body_storage + needed
        If body_storage < LONG_MAX - 1024 Then body_storage = body_storage + 1024
        ReDim Preserve os.body_data(0 To body_storage - 1)  ' ret=_ogg_realloc(os->body_data,...)
        os.body_storage = body_storage
    End If
    os_body_expand_ = 0
End Function

Private Function os_lacing_expand_(os As ogg_stream_state, ByVal needed As Long) As Long
    Dim lacing_storage As Long
    If os.lacing_storage - needed <= os.lacing_fill Then
        If os.lacing_storage > LONG_MAX - needed Then
            ogg_stream_clear os
            os_lacing_expand_ = -1
            Exit Function
        End If
        lacing_storage = os.lacing_storage + needed
        If lacing_storage < LONG_MAX - 32 Then lacing_storage = lacing_storage + 32
        ReDim Preserve os.lacing_vals(0 To lacing_storage - 1)
        ReDim Preserve os.granule_vals(0 To lacing_storage - 1)
        os.lacing_storage = lacing_storage
    End If
    os_lacing_expand_ = 0
End Function

' checksum the page
' Direct table CRC; note that this will be faster in the future if we
'   perform the checksum simultaneously with other copies
' (VB) buffer is buf(off ...)

Private Function os_update_crc_(ByVal crc As Long, buf() As Byte, ByVal off As Long, ByVal size As Long) As Long
    Dim t As Long
    If Not crcReady Then ogg_crc_init_
    Do While size >= 8
        ' crc^=((ogg_uint32_t)buffer[0]<<24)|((ogg_uint32_t)buffer[1]<<16)|((ogg_uint32_t)buffer[2]<<8)|((ogg_uint32_t)buffer[3]);
        If buf(off) And &H80 Then t = ((buf(off) And &H7F) * &H1000000) Or &H80000000 Else t = buf(off) * &H1000000
        crc = crc Xor (t Or (buf(off + 1) * &H10000) Or (buf(off + 2) * &H100&) Or buf(off + 3))

        crc = crc_lookup(7 * 256 + (((crc And &HFF000000) \ &H1000000) And &HFF)) Xor crc_lookup(6 * 256 + ((crc And &HFF0000) \ &H10000)) Xor _
              crc_lookup(5 * 256 + ((crc And &HFF00&) \ &H100&)) Xor crc_lookup(4 * 256 + (crc And &HFF)) Xor _
              crc_lookup(3 * 256 + buf(off + 4)) Xor crc_lookup(2 * 256 + buf(off + 5)) Xor _
              crc_lookup(1 * 256 + buf(off + 6)) Xor crc_lookup(buf(off + 7))

        off = off + 8
        size = size - 8
    Loop

    Do While size > 0                           ' while (size--)
        size = size - 1
        crc = shl8(crc) Xor crc_lookup((((crc And &HFF000000) \ &H1000000) And &HFF) Xor buf(off))
        off = off + 1
    Loop
    os_update_crc_ = crc
End Function

Public Sub ogg_page_checksum_set(og As ogg_page)
    ' (VB) if(og): a UDT is never NULL
    page_checksum_set_ og.header, 0, og.header_len, og.body, 0, og.body_len
End Sub

' (VB) ogg_page_checksum_set's body on a page given as its header and body arrays and offsets (ogg_sync_pageseek
' checks the page where it lies in oy.data, as the C's temp page points there)
Private Sub page_checksum_set_(header() As Byte, ByVal headerOff As Long, ByVal header_len As Long, _
                               body() As Byte, ByVal bodyOff As Long, ByVal body_len As Long)
    Dim crc_reg As Long
    crc_reg = 0

    ' safety; needed for API behavior, but not framing code
    header(headerOff + 22) = 0
    header(headerOff + 23) = 0
    header(headerOff + 24) = 0
    header(headerOff + 25) = 0

    crc_reg = os_update_crc_(crc_reg, header, headerOff, header_len)
    crc_reg = os_update_crc_(crc_reg, body, bodyOff, body_len)

    header(headerOff + 22) = crc_reg And &HFF
    header(headerOff + 23) = (crc_reg And &HFF00&) \ &H100&
    header(headerOff + 24) = (crc_reg And &HFF0000) \ &H10000
    header(headerOff + 25) = ((crc_reg And &HFF000000) \ &H1000000) And &HFF
End Sub

' submit data to the internal buffer of the framing engine
' (VB) iov(0 .. count - 1); if(!iov) return 0: an array is never NULL
Public Function ogg_stream_iovecin(os As ogg_stream_state, iov() As ogg_iovec_t, ByVal count As Long, _
                                   ByVal e_o_s As Long, ByVal granulepos As Double) As Long
    Dim bytes As Long, lacing_vals As Long, i As Long

    bytes = 0
    If ogg_stream_check(os) Then ogg_stream_iovecin = -1: Exit Function

    For i = 0 To count - 1
        If iov(i).iov_len < 0 Then ogg_stream_iovecin = -1: Exit Function  ' if(iov[i].iov_len>LONG_MAX) return -1
        If bytes > LONG_MAX - iov(i).iov_len Then ogg_stream_iovecin = -1: Exit Function
        bytes = bytes + iov(i).iov_len
    Next
    lacing_vals = bytes \ 255 + 1

    If os.body_returned Then
        ' advance packet data according to the body_returned pointer. We
        '   had to keep it around to return a pointer into the buffer last
        '   call

        os.body_fill = os.body_fill - os.body_returned
        If os.body_fill Then CopyMemory os.body_data(0), os.body_data(os.body_returned), os.body_fill
        os.body_returned = 0
    End If

    ' make sure we have the buffer storage
    If os_body_expand_(os, bytes) Then ogg_stream_iovecin = -1: Exit Function
    If os_lacing_expand_(os, lacing_vals) Then ogg_stream_iovecin = -1: Exit Function

    ' Copy in the submitted packet.  Yes, the copy is a waste; this is
    '   the liability of overly clean abstraction for the time being.  It
    '   will actually be fairly easy to eliminate the extra copy in the
    '   future

    For i = 0 To count - 1
        If iov(i).iov_len > 0 Then CopyMemory os.body_data(os.body_fill), iov(i).iov_base(0), iov(i).iov_len
        os.body_fill = os.body_fill + iov(i).iov_len
    Next

    ' Store lacing vals for this packet
    For i = 0 To lacing_vals - 2
        os.lacing_vals(os.lacing_fill + i) = 255
        os.granule_vals(os.lacing_fill + i) = os.granulepos
    Next
    i = lacing_vals - 1
    os.lacing_vals(os.lacing_fill + i) = bytes Mod 255
    os.granule_vals(os.lacing_fill + i) = granulepos
    os.granulepos = granulepos

    ' flag the first segment as the beginning of the packet
    os.lacing_vals(os.lacing_fill) = os.lacing_vals(os.lacing_fill) Or &H100

    os.lacing_fill = os.lacing_fill + lacing_vals

    ' for the sake of completeness
    os.packetno = os.packetno + 1

    If e_o_s Then os.e_o_s = 1

    ogg_stream_iovecin = 0
End Function

Public Function ogg_stream_packetin(os As ogg_stream_state, op As ogg_packet) As Long
    Dim iov(0 To 0) As ogg_iovec_t
    If op.bytes > 0 Then iov(0).iov_base = op.packet    ' iov.iov_base = op->packet ((VB) a copy)
    iov(0).iov_len = op.bytes
    ogg_stream_packetin = ogg_stream_iovecin(os, iov, 1, op.e_o_s, op.granulepos)
End Function

' Conditionally flush a page; force==0 will only flush nominal-size
'   pages, force==1 forces us to flush a page regardless of page size
'   so long as there's any data available at all.
Private Function ogg_stream_flush_i(os As ogg_stream_state, og As ogg_page, ByVal force As Long, ByVal nfill As Long) As Long
    Dim i As Long, vals As Long, maxvals As Long, bytes As Long, acc As Long, granule_pos As Double
    Dim packets_done As Long, packet_just_done As Long, serialno As Long, pageno As Long
    vals = 0
    bytes = 0
    acc = 0
    granule_pos = -1

    If ogg_stream_check(os) Then ogg_stream_flush_i = 0: Exit Function
    If os.lacing_fill > 255 Then maxvals = 255 Else maxvals = os.lacing_fill
    If maxvals = 0 Then ogg_stream_flush_i = 0: Exit Function

    ' construct a page
    ' decide how many segments to include

    ' If this is the initial header case, the first page must only include
    '   the initial header packet
    If os.b_o_s = 0 Then  ' 'initial header page' case
        granule_pos = 0
        vals = 0
        Do While vals < maxvals
            If (os.lacing_vals(vals) And &HFF) < 255 Then
                vals = vals + 1
                Exit Do
            End If
            vals = vals + 1
        Loop
    Else

        ' The extra packets_done, packet_just_done logic here attempts to do two things:
        '   1) Don't unnecessarily span pages.
        '   2) Unless necessary, don't flush pages if there are less than four packets on
        '      them; this expands page size to reduce unnecessary overhead if incoming packets
        '      are large.
        '   These are not necessary behaviors, just 'always better than naive flushing'
        '   without requiring an application to explicitly request a specific optimized
        '   behavior. We'll want an explicit behavior setup pathway eventually as well.

        packets_done = 0
        packet_just_done = 0
        vals = 0
        Do While vals < maxvals
            If acc > nfill And packet_just_done >= 4 Then
                force = 1
                Exit Do
            End If
            acc = acc + (os.lacing_vals(vals) And &HFF)
            If (os.lacing_vals(vals) And &HFF) < 255 Then
                granule_pos = os.granule_vals(vals)
                packets_done = packets_done + 1
                packet_just_done = packets_done
            Else
                packet_just_done = 0
            End If
            vals = vals + 1
        Loop
        If vals = 255 Then force = 1
    End If

    If force = 0 Then ogg_stream_flush_i = 0: Exit Function

    ' construct the header in temp storage
    os.header(0) = 79: os.header(1) = 103: os.header(2) = 103: os.header(3) = 83    ' memcpy(os->header,"OggS",4)

    ' stream structure version
    os.header(4) = &H0

    ' continued packet flag?
    os.header(5) = &H0
    If (os.lacing_vals(0) And &H100) = 0 Then os.header(5) = os.header(5) Or &H1
    ' first page flag?
    If os.b_o_s = 0 Then os.header(5) = os.header(5) Or &H2
    ' last page flag?
    If os.e_o_s <> 0 And os.lacing_fill = vals Then os.header(5) = os.header(5) Or &H4
    os.b_o_s = 1

    ' 64 bits of PCM position
    For i = 6 To 13
        os.header(i) = granule_pos - Int(granule_pos / 256#) * 256#    ' (unsigned char)(granule_pos&0xff)
        granule_pos = Int(granule_pos / 256#)                          ' granule_pos>>=8
    Next

    ' 32 bits of stream serial number
    serialno = os.serialno
    For i = 14 To 17
        os.header(i) = serialno And &HFF
        serialno = (serialno And &HFFFFFF00) \ &H100&                  ' serialno>>=8
    Next

    ' 32 bits of page counter (we have both counter and page header
    '   because this val can roll over)
    If os.pageno = -1 Then os.pageno = 0    ' because someone called
                                            ' stream_reset; this would be a
                                            ' strange thing to do in an
                                            ' encode stream, but it has
                                            ' plausible uses
    pageno = os.pageno
    os.pageno = os.pageno + 1
    For i = 18 To 21
        os.header(i) = pageno And &HFF
        pageno = (pageno And &HFFFFFF00) \ &H100&
    Next

    ' zero for computation; filled in later
    os.header(22) = 0
    os.header(23) = 0
    os.header(24) = 0
    os.header(25) = 0

    ' segment table
    os.header(26) = vals And &HFF
    For i = 0 To vals - 1
        os.header(i + 27) = os.lacing_vals(i) And &HFF
        bytes = bytes + os.header(i + 27)
    Next

    ' set pointers in the ogg_page struct ((VB) copies of the header and the body)
    os.header_fill = vals + 27
    og.header_len = os.header_fill
    ReDim og.header(0 To og.header_len - 1)
    CopyMemory og.header(0), os.header(0), og.header_len
    og.body_len = bytes
    If bytes > 0 Then
        ReDim og.body(0 To bytes - 1)
        CopyMemory og.body(0), os.body_data(os.body_returned), bytes
    Else
        Erase og.body
    End If

    ' advance the lacing data and set the body_returned pointer

    os.lacing_fill = os.lacing_fill - vals
    If os.lacing_fill > 0 Then
        CopyMemory os.lacing_vals(0), os.lacing_vals(vals), os.lacing_fill * 4
        CopyMemory os.granule_vals(0), os.granule_vals(vals), os.lacing_fill * 8
    End If
    os.body_returned = os.body_returned + bytes

    ' calculate the checksum

    ogg_page_checksum_set og
    ' (VB) the C's og->header is os->header: the checksum is in both
    For i = 22 To 25
        os.header(i) = og.header(i)
    Next

    ' done
    ogg_stream_flush_i = 1
End Function

' This will flush remaining packets into a page (returning nonzero),
'   even if there is not enough data to trigger a flush normally
'   (undersized page). If there are no packets or partial packets to
'   flush, ogg_stream_flush returns 0.  Note that ogg_stream_flush will
'   try to flush a normal sized page like ogg_stream_pageout; a call to
'   ogg_stream_flush does not guarantee that all packets have flushed.
'   Only a return value of 0 from ogg_stream_flush indicates all packet
'   data is flushed into pages.
'
'   since ogg_stream_flush will flush the last page in a stream even if
'   it's undersized, you almost certainly want to use ogg_stream_pageout
'   (and *not* ogg_stream_flush) unless you specifically need to flush
'   a page regardless of size in the middle of a stream.

Public Function ogg_stream_flush(os As ogg_stream_state, og As ogg_page) As Long
    ogg_stream_flush = ogg_stream_flush_i(os, og, 1, 4096)
End Function

' Like the above, but an argument is provided to adjust the nominal
'   page size for applications which are smart enough to provide their
'   own delay based flushing

Public Function ogg_stream_flush_fill(os As ogg_stream_state, og As ogg_page, ByVal nfill As Long) As Long
    ogg_stream_flush_fill = ogg_stream_flush_i(os, og, 1, nfill)
End Function

' This constructs pages from buffered packet segments.  The pointers
'returned are to static buffers; do not free. The returned buffers are
'good only until the next call (using the same ogg_stream_state)

Public Function ogg_stream_pageout(os As ogg_stream_state, og As ogg_page) As Long
    Dim force As Long
    force = 0
    If ogg_stream_check(os) Then ogg_stream_pageout = 0: Exit Function

    If (os.e_o_s <> 0 And os.lacing_fill <> 0) Or _
       (os.lacing_fill <> 0 And os.b_o_s = 0) Then      ' 'were done, now flush' case / 'initial header page' case
        force = 1
    End If

    ogg_stream_pageout = ogg_stream_flush_i(os, og, force, 4096)
End Function

' Like the above, but an argument is provided to adjust the nominal
'page size for applications which are smart enough to provide their
'own delay based flushing

Public Function ogg_stream_pageout_fill(os As ogg_stream_state, og As ogg_page, ByVal nfill As Long) As Long
    Dim force As Long
    force = 0
    If ogg_stream_check(os) Then ogg_stream_pageout_fill = 0: Exit Function

    If (os.e_o_s <> 0 And os.lacing_fill <> 0) Or _
       (os.lacing_fill <> 0 And os.b_o_s = 0) Then      ' 'were done, now flush' case / 'initial header page' case
        force = 1
    End If

    ogg_stream_pageout_fill = ogg_stream_flush_i(os, og, force, nfill)
End Function

Public Function ogg_stream_eos(os As ogg_stream_state) As Long
    If ogg_stream_check(os) Then ogg_stream_eos = 1: Exit Function
    ogg_stream_eos = os.e_o_s
End Function

' DECODING PRIMITIVES: packet streaming layer *********************

' This has two layers to place more of the multi-serialno and paging
'   control in the application's hands.  First, we expose a data buffer
'   using ogg_sync_buffer().  The app either copies into the
'   buffer, or passes it directly to read(), etc.  We then call
'   ogg_sync_wrote() to tell how many bytes we just added.
'
'   Pages are returned (pointers into the buffer in ogg_sync_state)
'   by ogg_sync_pageout().  The page is then submitted to
'   ogg_stream_pagein() along with the appropriate
'   ogg_stream_state* (ie, matching serialno).  We then get raw
'   packets out calling ogg_stream_packetout() with a
'   ogg_stream_state.

' initialize the struct to a known state
Public Function ogg_sync_init(oy As ogg_sync_state) As Long
    Dim blank As ogg_sync_state
    ' (VB) if(oy): a UDT is never NULL
    oy.storage = -1 ' used as a readiness flag
    oy = blank                                  ' memset(oy,0,sizeof(*oy))
    ogg_sync_init = 0
End Function

' clear non-flat storage within
Public Function ogg_sync_clear(oy As ogg_sync_state) As Long
    Dim blank As ogg_sync_state
    ' if(oy->data)_ogg_free(oy->data); memset(oy,0,sizeof(*oy))
    oy = blank
    ogg_sync_clear = 0
End Function

' (VB) the struct itself is the caller's: as ogg_sync_clear
Public Function ogg_sync_destroy(oy As ogg_sync_state) As Long
    ogg_sync_clear oy
    ' (_ogg_free(oy): the caller's variable)
    ogg_sync_destroy = 0
End Function

Public Function ogg_sync_check(oy As ogg_sync_state) As Long
    If oy.storage < 0 Then ogg_sync_check = -1: Exit Function
    ogg_sync_check = 0
End Function

' (VB) returns the offset into oy.data where the caller writes, -1 for NULL
Public Function ogg_sync_buffer(oy As ogg_sync_state, ByVal size As Long) As Long
    Dim newsize As Long
    If ogg_sync_check(oy) Then ogg_sync_buffer = -1: Exit Function

    ' first, clear out any space that has been previously returned
    If oy.returned Then
        oy.fill = oy.fill - oy.returned
        If oy.fill > 0 Then CopyMemory oy.data(0), oy.data(oy.returned), oy.fill
        oy.returned = 0
    End If

    If size > oy.storage - oy.fill Then
        ' We need to extend the internal buffer

        If size > INT_MAX - 4096 - oy.fill Then
            ogg_sync_clear oy
            ogg_sync_buffer = -1
            Exit Function
        End If
        newsize = size + oy.fill + 4096 ' an extra page to be nice
        If oy.vbInit Then
            ReDim Preserve oy.data(0 To newsize - 1)    ' ret=_ogg_realloc(oy->data,newsize)
        Else
            ReDim oy.data(0 To newsize - 1)             ' ret=_ogg_malloc(newsize)
            oy.vbInit = True
        End If
        oy.storage = newsize
    End If

    ' expose a segment at least as large as requested at the fill mark
    ogg_sync_buffer = oy.fill
End Function

Public Function ogg_sync_wrote(oy As ogg_sync_state, ByVal bytes As Long) As Long
    If ogg_sync_check(oy) Then ogg_sync_wrote = -1: Exit Function
    If oy.fill + bytes > oy.storage Then ogg_sync_wrote = -1: Exit Function
    oy.fill = oy.fill + bytes
    ogg_sync_wrote = 0
End Function

' sync the stream.  This is meant to be useful for finding page
'   boundaries.
'
'   return values for this:
'  -n) skipped n bytes
'   0) page not ready; more data (no bytes skipped)
'   n) page synced at current location; page length n bytes
'
' (VB) page is the offset of the C's page pointer in oy.data

Public Function ogg_sync_pageseek(oy As ogg_sync_state, og As ogg_page, Optional ByVal ogIsNull As Boolean = False) As Long
    Dim page As Long, nxt As Long, bytes As Long, headerbytes As Long, i As Long
    Dim chksum(0 To 3) As Byte, differ As Boolean, lim As Long

    If ogg_sync_check(oy) Then ogg_sync_pageseek = 0: Exit Function

    page = oy.returned
    bytes = oy.fill - oy.returned

    If oy.headerbytes = 0 Then
        If bytes < 27 Then ogg_sync_pageseek = 0: Exit Function ' not enough for a header

        ' verify capture pattern
        If oy.data(page) <> 79 Then GoTo sync_fail              ' if(memcmp(page,"OggS",4))goto sync_fail
        If oy.data(page + 1) <> 103 Then GoTo sync_fail
        If oy.data(page + 2) <> 103 Then GoTo sync_fail
        If oy.data(page + 3) <> 83 Then GoTo sync_fail

        headerbytes = oy.data(page + 26) + 27
        If bytes < headerbytes Then ogg_sync_pageseek = 0: Exit Function    ' not enough for header + seg table

        ' count up body length in the segment table

        For i = 0 To oy.data(page + 26) - 1
            oy.bodybytes = oy.bodybytes + oy.data(page + 27 + i)
        Next
        oy.headerbytes = headerbytes
    End If

    If oy.bodybytes + oy.headerbytes > bytes Then ogg_sync_pageseek = 0: Exit Function

    ' The whole test page is buffered.  Verify the checksum
    ' Grab the checksum bytes, set the header field to zero

    CopyMemory chksum(0), oy.data(page + 22), 4
    oy.data(page + 22) = 0: oy.data(page + 23) = 0: oy.data(page + 24) = 0: oy.data(page + 25) = 0

    ' set up a temp page struct and recompute the checksum
    ' (VB) log.header=page; log.header_len=oy->headerbytes; log.body=page+oy->headerbytes;
    ' log.body_len=oy->bodybytes; ogg_page_checksum_set(&log): the page where it lies in oy.data
    page_checksum_set_ oy.data, page, oy.headerbytes, oy.data, page + oy.headerbytes, oy.bodybytes

    ' Compare
    differ = False
    For i = 0 To 3
        If chksum(i) <> oy.data(page + 22 + i) Then differ = True
    Next
    If differ Then
        ' D'oh.  Mismatch! Corrupt page (or miscapture and not a page
        '   at all)
        ' replace the computed checksum with the one actually read in
        CopyMemory oy.data(page + 22), chksum(0), 4

        ' (DISABLE_CRC is not defined) Bad checksum. Lose sync
        GoTo sync_fail
    End If

    ' yes, have a whole page all ready to go
    If Not ogIsNull Then
        ' (VB) og->header=page, og->body=page+oy->headerbytes: copies
        og.header_len = oy.headerbytes
        ReDim og.header(0 To og.header_len - 1)
        CopyMemory og.header(0), oy.data(page), og.header_len
        og.body_len = oy.bodybytes
        If og.body_len > 0 Then
            ReDim og.body(0 To og.body_len - 1)
            CopyMemory og.body(0), oy.data(page + oy.headerbytes), og.body_len
        Else
            Erase og.body
        End If
    End If

    oy.unsynced = 0
    bytes = oy.headerbytes + oy.bodybytes
    oy.returned = oy.returned + bytes
    oy.headerbytes = 0
    oy.bodybytes = 0
    ogg_sync_pageseek = bytes
    Exit Function

sync_fail:

    oy.headerbytes = 0
    oy.bodybytes = 0

    ' search for possible capture
    ' next=memchr(page+1,'O',bytes-1)
    nxt = -1
    lim = page + bytes - 1
    For i = page + 1 To lim
        If oy.data(i) = 79 Then nxt = i: Exit For
    Next
    If nxt < 0 Then nxt = oy.fill   ' next=oy->data+oy->fill

    oy.returned = nxt
    ogg_sync_pageseek = -(nxt - page)
End Function

' sync the stream and get a page.  Keep trying until we find a page.
'   Suppress 'sync errors' after reporting the first.
'
'   return values:
'   -1) recapture (hole in data)
'    0) need more data
'    1) page returned
'
'   Returns pointers into buffered data; invalidated by next call to
'   _stream, _clear, _init, or _buffer

Public Function ogg_sync_pageout(oy As ogg_sync_state, og As ogg_page) As Long
    Dim ret As Long

    If ogg_sync_check(oy) Then ogg_sync_pageout = 0: Exit Function

    ' all we need to do is verify a page at the head of the stream
    '   buffer.  If it doesn't verify, we look for the next potential
    '   frame

    Do
        ret = ogg_sync_pageseek(oy, og)
        If ret > 0 Then
            ' have a page
            ogg_sync_pageout = 1
            Exit Function
        End If
        If ret = 0 Then
            ' need more data
            ogg_sync_pageout = 0
            Exit Function
        End If

        ' head did not start a synced page... skipped some bytes
        If oy.unsynced = 0 Then
            oy.unsynced = 1
            ogg_sync_pageout = -1
            Exit Function
        End If

        ' loop. keep looking

    Loop
End Function

' add the incoming page to the stream state; we decompose the page
'   into packet segments here as well.
' (VB) body is og.body(bodyOff ...)

Public Function ogg_stream_pagein(os As ogg_stream_state, og As ogg_page) As Long
    Dim bodyOff As Long, bodysize As Long, segptr As Long
    Dim version As Long, continued As Long, bos As Long, eos As Long, granulepos As Double
    Dim serialno As Long, pageno As Long, segments As Long
    Dim lr As Long, br As Long, i As Long, val_ As Long, saved As Long, skip As Boolean

    bodyOff = 0                                 ' unsigned char *body=og->body
    bodysize = og.body_len
    segptr = 0

    version = ogg_page_version(og)
    continued = ogg_page_continued(og)
    bos = ogg_page_bos(og)
    eos = ogg_page_eos(og)
    granulepos = ogg_page_granulepos(og)
    serialno = ogg_page_serialno(og)
    pageno = ogg_page_pageno(og)
    segments = og.header(26)

    If ogg_stream_check(os) Then ogg_stream_pagein = -1: Exit Function

    ' clean up 'returned data'
    lr = os.lacing_returned
    br = os.body_returned

    ' body data
    If br Then
        os.body_fill = os.body_fill - br
        If os.body_fill Then CopyMemory os.body_data(0), os.body_data(br), os.body_fill
        os.body_returned = 0
    End If

    If lr Then
        ' segment table
        If os.lacing_fill - lr Then
            CopyMemory os.lacing_vals(0), os.lacing_vals(lr), (os.lacing_fill - lr) * 4
            CopyMemory os.granule_vals(0), os.granule_vals(lr), (os.lacing_fill - lr) * 8
        End If
        os.lacing_fill = os.lacing_fill - lr
        os.lacing_packet = os.lacing_packet - lr
        os.lacing_returned = 0
    End If

    ' check the serial number
    If serialno <> os.serialno Then ogg_stream_pagein = -1: Exit Function
    If version > 0 Then ogg_stream_pagein = -1: Exit Function

    If os_lacing_expand_(os, segments + 1) Then ogg_stream_pagein = -1: Exit Function

    ' are we in sequence?
    If pageno <> os.pageno Then

        ' unroll previous partial packet (if any)
        For i = os.lacing_packet To os.lacing_fill - 1
            os.body_fill = os.body_fill - (os.lacing_vals(i) And &HFF)
        Next
        os.lacing_fill = os.lacing_packet

        ' make a note of dropped data in segment table
        If os.pageno <> -1 Then
            os.lacing_vals(os.lacing_fill) = &H400
            os.lacing_fill = os.lacing_fill + 1
            os.lacing_packet = os.lacing_packet + 1
        End If
    End If

    ' are we a 'continued packet' page?  If so, we may need to skip
    '   some segments
    If continued Then
        skip = (os.lacing_fill < 1)
        If Not skip Then
            If (os.lacing_vals(os.lacing_fill - 1) And &HFF) < 255 Or os.lacing_vals(os.lacing_fill - 1) = &H400 Then skip = True
        End If
        If skip Then
            bos = 0
            Do While segptr < segments
                val_ = og.header(27 + segptr)
                bodyOff = bodyOff + val_
                bodysize = bodysize - val_
                If val_ < 255 Then
                    segptr = segptr + 1
                    Exit Do
                End If
                segptr = segptr + 1
            Loop
        End If
    End If

    If bodysize Then
        If os_body_expand_(os, bodysize) Then ogg_stream_pagein = -1: Exit Function
        CopyMemory os.body_data(os.body_fill), og.body(bodyOff), bodysize
        os.body_fill = os.body_fill + bodysize
    End If

    saved = -1
    Do While segptr < segments
        val_ = og.header(27 + segptr)
        os.lacing_vals(os.lacing_fill) = val_
        os.granule_vals(os.lacing_fill) = -1

        If bos Then
            os.lacing_vals(os.lacing_fill) = os.lacing_vals(os.lacing_fill) Or &H100
            bos = 0
        End If

        If val_ < 255 Then saved = os.lacing_fill

        os.lacing_fill = os.lacing_fill + 1
        segptr = segptr + 1

        If val_ < 255 Then os.lacing_packet = os.lacing_fill
    Loop

    ' set the granulepos on the last granuleval of the last full packet
    If saved <> -1 Then
        os.granule_vals(saved) = granulepos
    End If

    If eos Then
        os.e_o_s = 1
        If os.lacing_fill > 0 Then os.lacing_vals(os.lacing_fill - 1) = os.lacing_vals(os.lacing_fill - 1) Or &H200
    End If

    os.pageno = pageno + 1

    ogg_stream_pagein = 0
End Function

' clear things to an initial state.  Good to call, eg, before seeking
Public Function ogg_sync_reset(oy As ogg_sync_state) As Long
    If ogg_sync_check(oy) Then ogg_sync_reset = -1: Exit Function

    oy.fill = 0
    oy.returned = 0
    oy.unsynced = 0
    oy.headerbytes = 0
    oy.bodybytes = 0
    ogg_sync_reset = 0
End Function

Public Function ogg_stream_reset(os As ogg_stream_state) As Long
    If ogg_stream_check(os) Then ogg_stream_reset = -1: Exit Function

    os.body_fill = 0
    os.body_returned = 0

    os.lacing_fill = 0
    os.lacing_packet = 0
    os.lacing_returned = 0

    os.header_fill = 0

    os.e_o_s = 0
    os.b_o_s = 0
    os.pageno = -1
    os.packetno = 0
    os.granulepos = 0

    ogg_stream_reset = 0
End Function

Public Function ogg_stream_reset_serialno(os As ogg_stream_state, ByVal serialno As Long) As Long
    If ogg_stream_check(os) Then ogg_stream_reset_serialno = -1: Exit Function
    ogg_stream_reset os
    os.serialno = serialno
    ogg_stream_reset_serialno = 0
End Function

Private Function packetout_(os As ogg_stream_state, op As ogg_packet, ByVal opIsNull As Boolean, ByVal adv As Long) As Long
    Dim ptr As Long, size As Long, bytes As Long, eos As Long, bos As Long, val_ As Long

    ' The last part of decode. We have the stream broken into packet
    '   segments.  Now we need to group them into packets (or return the
    '   out of sync markers)

    ptr = os.lacing_returned

    If os.lacing_packet <= ptr Then packetout_ = 0: Exit Function

    If os.lacing_vals(ptr) And &H400 Then
        ' we need to tell the codec there's a gap; it might need to
        '   handle previous packet dependencies.
        os.lacing_returned = os.lacing_returned + 1
        os.packetno = os.packetno + 1
        packetout_ = -1
        Exit Function
    End If

    If opIsNull And adv = 0 Then packetout_ = 1: Exit Function   ' just using peek as an inexpensive way
                                                                ' to ask if there's a whole packet
                                                                ' waiting

    ' Gather the whole packet. We'll have no holes or a partial packet
    size = os.lacing_vals(ptr) And &HFF
    bytes = size
    eos = os.lacing_vals(ptr) And &H200 ' last packet of the stream?
    bos = os.lacing_vals(ptr) And &H100 ' first packet of the stream?

    Do While size = 255
        ptr = ptr + 1
        val_ = os.lacing_vals(ptr)
        size = val_ And &HFF
        If val_ And &H200 Then eos = &H200
        bytes = bytes + size
    Loop

    If Not opIsNull Then
        op.e_o_s = eos
        op.b_o_s = bos
        ' op->packet=os->body_data+os->body_returned ((VB) a copy)
        If bytes > 0 Then
            ReDim op.packet(0 To bytes - 1)
            CopyMemory op.packet(0), os.body_data(os.body_returned), bytes
        Else
            Erase op.packet
        End If
        op.packetno = os.packetno
        op.granulepos = os.granule_vals(ptr)
        op.bytes = bytes
    End If

    If adv Then
        os.body_returned = os.body_returned + bytes
        os.lacing_returned = ptr + 1
        os.packetno = os.packetno + 1
    End If
    packetout_ = 1
End Function

Public Function ogg_stream_packetout(os As ogg_stream_state, op As ogg_packet, Optional ByVal opIsNull As Boolean = False) As Long
    If ogg_stream_check(os) Then ogg_stream_packetout = 0: Exit Function
    ogg_stream_packetout = packetout_(os, op, opIsNull, 1)
End Function

Public Function ogg_stream_packetpeek(os As ogg_stream_state, op As ogg_packet, Optional ByVal opIsNull As Boolean = False) As Long
    If ogg_stream_check(os) Then ogg_stream_packetpeek = 0: Exit Function
    ogg_stream_packetpeek = packetout_(os, op, opIsNull, 0)
End Function

Public Sub ogg_packet_clear(op As ogg_packet)
    Dim blank As ogg_packet
    ' _ogg_free(op->packet); memset(op, 0, sizeof(*op))
    op = blank
End Sub

' (_V_SELFTEST, framing.c's test program, is not ported)
