Attribute VB_Name = "theora_bitpack_c"
' ref/theora (libtheora 1.2.0) lib/bitpack.c with bitpack.h: the MSb-first bit reader.
' (VB) The window holds 31 bits instead of the C's sizeof(size_t)*CHAR_BIT, so that it is never negative:
' a read is then a division by a power of two and a shift a mask and a multiplication, without overflow.
' The bytes come in at the same positions relative to the top of the window, so every read returns what the
' C returns. A refill loads whole bytes while 8 bits are free (24 to 31 bits available after it, where the
' C has 25 to 32), which serves reads of up to 24 bits; oc_pack_read splits a longer read into two (at the
' end of the data the C's eof rule gives the same result for the two parts as for the whole). The pointers
' are indices into the buffer's own copy of the bytes (buf: oc_pack_readinit copies them from the packet).
' oc_pack_bytes_left counts the same whole bytes as the C.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' typedef size_t oc_pb_window: (VB) a Long holding 31 bits (above)
' OC_PB_WINDOW_SIZE ((int)sizeof(oc_pb_window)*CHAR_BIT): (VB) 31
Public Const OC_PB_WINDOW_SIZE As Long = 31
' This is meant to be a large, positive constant that can still be efficiently loaded as an immediate (on
' platforms like ARM, for example). Even relatively modest values like 100 would work fine.
Public Const OC_LOTS_OF_BITS As Long = &H40000000

Public Type oc_pack_buf
    stop_ As Long           ' const unsigned char *stop: (VB) an index into buf
    ptr As Long             ' const unsigned char *ptr: (VB) an index into buf
    window As Long          ' oc_pb_window: (VB) 31 bits, the next bit to read at bit 30
    bits As Long
    eof As Long
    buf() As Byte           ' (VB) the bytes the pointers point into (oc_pack_readinit's copy)
    vbCapacity As Long      ' (VB) buf's size, kept between packets
End Type

' (VB) OC_PB_POW2(n) = 1 << n (n = 0..30), OC_PB_MASK(n) = (1 << n) - 1 (n = 0..31)
Public OC_PB_POW2() As Long
Public OC_PB_MASK() As Long

' (VB) the tables above (oc_vb_init)
Public Sub oc_pack_vbInit()
    Dim i As Long
    ReDim OC_PB_POW2(0 To 31)
    ReDim OC_PB_MASK(0 To 31)
    OC_PB_POW2(0) = 1
    For i = 1 To 30
        OC_PB_POW2(i) = OC_PB_POW2(i - 1) * 2
    Next
    OC_PB_POW2(31) = &H80000000         ' (VB) not a power of two as a Long: never used as one
    For i = 0 To 30
        OC_PB_MASK(i) = OC_PB_POW2(i) - 1
    Next
    OC_PB_MASK(31) = &H7FFFFFFF
End Sub

' We're 'MSb' endian; if we write a word but read individual bits, then we'll read the MSb first.

' (VB) _buf is buf(bufOffset), copied into the reader (_bytes of it)
Public Sub oc_pack_readinit(b As oc_pack_buf, buf() As Byte, ByVal bufOffset As Long, ByVal bytes As Long)
    ' memset(_b,0,sizeof(*_b))
    b.window = 0
    b.bits = 0
    b.eof = 0
    If bytes > 0 Then
        If b.vbCapacity < bytes Then
            ReDim b.buf(0 To bytes - 1)
            b.vbCapacity = bytes
        End If
        CopyMemory b.buf(0), buf(bufOffset), bytes
    End If
    b.ptr = 0
    b.stop_ = bytes
End Sub

Private Function oc_pack_refill(b As oc_pack_buf, ByVal bits_ As Long) As Long
    Dim ptr As Long
    Dim stop_ As Long
    Dim window As Long
    Dim available As Long
    Dim shift As Long
    stop_ = b.stop_
    ptr = b.ptr
    window = b.window
    available = b.bits
    shift = OC_PB_WINDOW_SIZE - available
    Do While 7 < shift
        If ptr >= stop_ Then Exit Do
        shift = shift - 8
        ' window|=(oc_pb_window)*ptr++<<shift: (VB) the bits that stay inside the window
        window = window Or ((b.buf(ptr) And OC_PB_MASK(OC_PB_WINDOW_SIZE - shift)) * OC_PB_POW2(shift))
        ptr = ptr + 1
    Loop
    b.ptr = ptr
    available = OC_PB_WINDOW_SIZE - shift
    If bits_ > available Then
        If ptr >= stop_ Then
            b.eof = 1
            available = OC_LOTS_OF_BITS
        Else
            ' window|=*ptr>>(available&7): (VB) the top shift bits of the next byte fill the window
            window = window Or (b.buf(ptr) \ OC_PB_POW2(8 - shift))
        End If
    End If
    b.bits = available
    oc_pack_refill = window
End Function

Public Function oc_pack_look1(b As oc_pack_buf) As Long
    Dim window As Long
    Dim available As Long
    window = b.window
    available = b.bits
    If available < 1 Then
        window = oc_pack_refill(b, 1)
        b.window = window
    End If
    oc_pack_look1 = window \ OC_PB_POW2(OC_PB_WINDOW_SIZE - 1)
End Function

Public Sub oc_pack_adv1(b As oc_pack_buf)
    b.window = (b.window And OC_PB_MASK(OC_PB_WINDOW_SIZE - 1)) * 2
    b.bits = b.bits - 1
End Sub

' Here we assume that 0<=_bits&&_bits<=32.
' (oc_pack_read: bitpack.h maps it to oc_pack_read_c)
Public Function oc_pack_read(b As oc_pack_buf, ByVal bits_ As Long) As Long
    Dim window As Long
    Dim available As Long
    Dim result As Long
    Dim hi As Long
    If bits_ > 24 Then
        ' (VB) two reads: the top bits_ - 16 bits, then 16 (result: the C's long of the bits)
        hi = oc_pack_read(b, bits_ - 16)
        result = oc_pack_read(b, 16)
        If (hi And &H8000&) <> 0 Then
            oc_pack_read = ((hi And &H7FFF&) * &H10000) Or &H80000000 Or result
        Else
            oc_pack_read = (hi * &H10000) Or result
        End If
        Exit Function
    End If
    window = b.window
    available = b.bits
    If bits_ = 0 Then oc_pack_read = 0: Exit Function
    If available < bits_ Then
        window = oc_pack_refill(b, bits_)
        available = b.bits
    End If
    result = window \ OC_PB_POW2(OC_PB_WINDOW_SIZE - bits_)
    available = available - bits_
    ' window<<=1; window<<=_bits-1;
    window = (window And OC_PB_MASK(OC_PB_WINDOW_SIZE - bits_)) * OC_PB_POW2(bits_)
    b.window = window
    b.bits = available
    oc_pack_read = result
End Function

' (oc_pack_read1: bitpack.h maps it to oc_pack_read1_c)
Public Function oc_pack_read1(b As oc_pack_buf) As Long
    Dim window As Long
    Dim available As Long
    Dim result As Long
    window = b.window
    available = b.bits
    If available < 1 Then
        window = oc_pack_refill(b, 1)
        available = b.bits
    End If
    result = window \ &H40000000            ' window>>OC_PB_WINDOW_SIZE-1
    available = available - 1
    window = (window And &H3FFFFFFF) * 2    ' window<<=1
    b.window = window
    b.bits = available
    oc_pack_read1 = result
End Function

' returns -1 for read beyond EOF, or the number of whole bytes available
Public Function oc_pack_bytes_left(b As oc_pack_buf) As Long
    If b.eof <> 0 Then oc_pack_bytes_left = -1: Exit Function
    oc_pack_bytes_left = b.stop_ - b.ptr + (b.bits \ 8)     ' _b->bits>>3 (bits >= 0)
End Function
