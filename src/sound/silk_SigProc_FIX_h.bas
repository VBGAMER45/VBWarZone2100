Attribute VB_Name = "silk_SigProc_FIX_h"
' libopus 1.5.2 (ref/opus/silk): SigProc_FIX.h's fixed-point macros and inline functions, with macros.h's,
' Inlines.h's, typedef.h's limits and ecintrin.h's EC_ILOG (for silk_CLZ32): the SILK decoder's integer
' arithmetic, as functions with the C's exact results. Not ported: the declarations of the signal processing
' functions (each in the module of its .c file) and the encoder-only inline helpers.
' (VB) The projects build without integer overflow checks: Long +, - and * wrap as C's int32 arithmetic does
' on x86 (two's complement), and a Long stored into an Integer keeps its low 16 bits, as C's (opus_int16).
' An arithmetic right shift of a Long is (a And mask) \ 2^n (the masked value divides exactly, so the
' quotient is the floor); a left shift is a * 2^n (wrapping). Integer * Integer and Integer + Integer are
' 16-bit in VB: C's promotion to int is a CLng (or a Long operand) in every such expression.
' The macros with 64-bit products (OPUS_FAST_INT64 or not: both give the same 32-bit results) are computed
' with the 16-bit halves of one operand: silk_SMULWB is exact in 32 bits, silk_SMULWW is silk_SMULWB plus a
' wrapping 32-bit product (macros.h's non-FAST_INT64 form), and silk_SMMUL and silk_RSHIFT_ROUND64(silk_SMULL)
' are exact in Double (silk_SMULL_RSHIFT). Hot loops write these out in place; this module is the reference.
' The shifts read tables of powers of 2 that silk_SigProc_FIX_init fills (silk_tables_init, from
' silk_InitDecoder): code that can run before a decoder exists (opus_c's packet functions) must not use them.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)

' typedef.h
Public Const silk_int32_MAX As Long = &H7FFFFFFF        ' 2^31 - 1 =  2147483647
Public Const silk_int32_MIN As Long = &H80000000        ' -2^31     = -2147483648
Public Const silk_int16_MAX As Long = &H7FFF&           ' 2^15 - 1 =  32767
Public Const silk_int16_MIN As Long = -32768            ' -2^15     = -32768
Public Const silk_int8_MAX As Long = 127                ' 2^7 - 1  =  127
Public Const silk_int8_MIN As Long = -128               ' -2^7      = -128
Public Const silk_uint8_MAX As Long = 255               ' 2^8 - 1 = 255
Public Const silk_TRUE As Long = 1
Public Const silk_FALSE As Long = 0

Public Const SILK_MAX_ORDER_LPC As Long = 24            ' max order of the LPC analysis in schur() and k2a()

' PSEUDO-RANDOM GENERATOR
Public Const RAND_MULTIPLIER As Long = 196314165
Public Const RAND_INCREMENT As Long = 907633515

' (VB) 2^n as a Long (2^31 is &H80000000) and as a Double
Private pow2(0 To 31) As Long
Private pow2d(0 To 63) As Double
Private silk_pow2_ready As Boolean

' (VB) the tables above; silk_tables_init calls it
Public Sub silk_SigProc_FIX_init()
    Dim i As Long
    If silk_pow2_ready Then Exit Sub
    pow2(0) = 1
    For i = 1 To 30
        pow2(i) = pow2(i - 1) * 2
    Next
    pow2(31) = &H80000000
    pow2d(0) = 1#
    For i = 1 To 63
        pow2d(i) = pow2d(i - 1) * 2#
    Next
    silk_pow2_ready = True
End Sub

' (VB) C's (opus_int16) of an int: the low 16 bits, sign-extended
Public Function silk_I16(ByVal a As Long) As Long
    silk_I16 = ((a And &HFFFF&) Xor &H8000&) - &H8000&
End Function

' (VB) C's (opus_int8) of an int: the low 8 bits, sign-extended
Public Function silk_I8(ByVal a As Long) As Long
    silk_I8 = ((a And &HFF&) Xor &H80&) - &H80&
End Function

' (VB) C's (opus_int32) of an opus_int64 held in a Double (an integer): modulo 2^32
Public Function silk_I32(ByVal d As Double) As Long
    If d >= -2147483648# And d <= 2147483647# Then
        silk_I32 = d
    Else
        d = d - Int(d / 4294967296#) * 4294967296#
        If d >= 2147483648# Then silk_I32 = d - 4294967296# Else silk_I32 = d
    End If
End Function

'/* (a32 * (opus_int32)((opus_int16)(b32))) >> 16 output have to be 32bit int */
Public Function silk_SMULWB(ByVal a32 As Long, ByVal b32 As Long) As Long
    Dim b As Long
    b = ((b32 And &HFFFF&) Xor &H8000&) - &H8000&
    silk_SMULWB = ((a32 And &HFFFF0000) \ &H10000) * b + ((((a32 And &HFFFF&) * b) And &HFFFF0000) \ &H10000)
End Function

'/* a32 + (b32 * (opus_int32)((opus_int16)(c32))) >> 16 output have to be 32bit int */
Public Function silk_SMLAWB(ByVal a32 As Long, ByVal b32 As Long, ByVal c32 As Long) As Long
    Dim c As Long
    c = ((c32 And &HFFFF&) Xor &H8000&) - &H8000&
    silk_SMLAWB = a32 + (((b32 And &HFFFF0000) \ &H10000) * c + ((((b32 And &HFFFF&) * c) And &HFFFF0000) \ &H10000))
End Function

'/* (a32 * (b32 >> 16)) >> 16 */
Public Function silk_SMULWT(ByVal a32 As Long, ByVal b32 As Long) As Long
    Dim b As Long
    b = (b32 And &HFFFF0000) \ &H10000
    silk_SMULWT = ((a32 And &HFFFF0000) \ &H10000) * b + ((((a32 And &HFFFF&) * b) And &HFFFF0000) \ &H10000)
End Function

'/* a32 + (b32 * (c32 >> 16)) >> 16 */
Public Function silk_SMLAWT(ByVal a32 As Long, ByVal b32 As Long, ByVal c32 As Long) As Long
    silk_SMLAWT = a32 + silk_SMULWT(b32, c32)
End Function

'/* (opus_int32)((opus_int16)(a3))) * (opus_int32)((opus_int16)(b32)) output have to be 32bit int */
Public Function silk_SMULBB(ByVal a32 As Long, ByVal b32 As Long) As Long
    silk_SMULBB = (((a32 And &HFFFF&) Xor &H8000&) - &H8000&) * (((b32 And &HFFFF&) Xor &H8000&) - &H8000&)
End Function

'/* a32 + (opus_int32)((opus_int16)(b32)) * (opus_int32)((opus_int16)(c32)) output have to be 32bit int */
Public Function silk_SMLABB(ByVal a32 As Long, ByVal b32 As Long, ByVal c32 As Long) As Long
    silk_SMLABB = a32 + (((b32 And &HFFFF&) Xor &H8000&) - &H8000&) * (((c32 And &HFFFF&) Xor &H8000&) - &H8000&)
End Function

'/* (opus_int32)((opus_int16)(a32)) * (b32 >> 16) */
Public Function silk_SMULBT(ByVal a32 As Long, ByVal b32 As Long) As Long
    silk_SMULBT = (((a32 And &HFFFF&) Xor &H8000&) - &H8000&) * ((b32 And &HFFFF0000) \ &H10000)
End Function

'/* a32 + (opus_int32)((opus_int16)(b32)) * (c32 >> 16) */
Public Function silk_SMLABT(ByVal a32 As Long, ByVal b32 As Long, ByVal c32 As Long) As Long
    silk_SMLABT = a32 + (((b32 And &HFFFF&) Xor &H8000&) - &H8000&) * ((c32 And &HFFFF0000) \ &H10000)
End Function

'/* (a32 * b32) >> 16 */
' (VB) macros.h's form without OPUS_FAST_INT64: silk_MLA(silk_SMULWB((a32), (b32)), (a32), silk_RSHIFT_ROUND((b32), 16)),
' equal modulo 2^32 to the 64-bit product shifted
Public Function silk_SMULWW(ByVal a32 As Long, ByVal b32 As Long) As Long
    Dim b As Long, r As Long
    b = ((b32 And &HFFFF&) Xor &H8000&) - &H8000&
    r = (b32 And &HFFFF8000) \ &H8000&                  ' silk_RSHIFT_ROUND( b32, 16 )
    r = ((r + 1) And &HFFFFFFFE) \ 2
    silk_SMULWW = ((a32 And &HFFFF0000) \ &H10000) * b + ((((a32 And &HFFFF&) * b) And &HFFFF0000) \ &H10000) + a32 * r
End Function

'/* a32 + ((b32 * c32) >> 16) */
Public Function silk_SMLAWW(ByVal a32 As Long, ByVal b32 As Long, ByVal c32 As Long) As Long
    silk_SMLAWW = a32 + silk_SMULWW(b32, c32)
End Function

'/* add/subtract with output saturated */
Public Function silk_ADD_SAT32(ByVal a As Long, ByVal b As Long) As Long
    Dim s As Long
    s = a + b
    If (s And &H80000000) = 0 Then
        If ((a And b) And &H80000000) <> 0 Then silk_ADD_SAT32 = silk_int32_MIN Else silk_ADD_SAT32 = s
    Else
        If ((a Or b) And &H80000000) = 0 Then silk_ADD_SAT32 = silk_int32_MAX Else silk_ADD_SAT32 = s
    End If
End Function

Public Function silk_SUB_SAT32(ByVal a As Long, ByVal b As Long) As Long
    Dim s As Long
    s = a - b
    If (s And &H80000000) = 0 Then
        If (a And (b Xor &H80000000) And &H80000000) <> 0 Then silk_SUB_SAT32 = silk_int32_MIN Else silk_SUB_SAT32 = s
    Else
        If ((a Xor &H80000000) And b And &H80000000) <> 0 Then silk_SUB_SAT32 = silk_int32_MAX Else silk_SUB_SAT32 = s
    End If
End Function

' ecintrin.h's EC_ILOG: the number of significant bits of an unsigned value (0 for 0)
Private Function EC_ILOG_(ByVal x As Long) As Long
    Dim n As Long
    If x = 0 Then Exit Function
    If x < 0 Then EC_ILOG_ = 32: Exit Function
    n = 32
    If (x And &HFFFF0000) = 0 Then n = n - 16: x = x * &H10000
    If (x And &HFF000000) = 0 Then n = n - 8: x = x * &H100&
    If (x And &HF0000000) = 0 Then n = n - 4: x = x * &H10&
    If (x And &HC0000000) = 0 Then n = n - 2: x = x * 4
    If (x And &H80000000) = 0 Then n = n - 1
    EC_ILOG_ = n
End Function

Public Function silk_CLZ16(ByVal in16 As Long) As Long
    silk_CLZ16 = 32 - EC_ILOG_(((in16 And &HFFFF&) * &H10000) Or &H8000&)
End Function

Public Function silk_CLZ32(ByVal in32 As Long) As Long
    If in32 <> 0 Then silk_CLZ32 = 32 - EC_ILOG_(in32) Else silk_CLZ32 = 32
End Function

'/* Rotate a32 right by 'rot' bits. Negative rot values result in rotating
'   left. Output is 32bit int.
'   Note: contemporary compilers recognize the C expression below and
'   compile it into a 'ror' instruction if available. No need for OPUS_INLINE ASM! */
Public Function silk_ROR32(ByVal a32 As Long, ByVal rot As Long) As Long
    Dim m As Long, r As Long
    If rot = 0 Then
        silk_ROR32 = a32
    ElseIf rot < 0 Then
        m = -rot
        silk_ROR32 = (a32 * pow2(m)) Or silk_RSHIFT_uint(a32, 32 - m)
    Else
        r = rot
        silk_ROR32 = (a32 * pow2(32 - r)) Or silk_RSHIFT_uint(a32, r)
    End If
End Function

'/* (a32 * b32) output have to be 32bit int */  silk_MUL(a32, b32) is a32 * b32 (wrapping)
'/* a32 + (b32 * c32) output have to be 32bit int */  silk_MLA(a32, b32, c32) is a32 + b32 * c32

'/* ((a32 >> 16)  * (b32 >> 16)) output have to be 32bit int */
Public Function silk_SMULTT(ByVal a32 As Long, ByVal b32 As Long) As Long
    silk_SMULTT = ((a32 And &HFFFF0000) \ &H10000) * ((b32 And &HFFFF0000) \ &H10000)
End Function

'/* a32 + ((a32 >> 16)  * (b32 >> 16)) output have to be 32bit int */
Public Function silk_SMLATT(ByVal a32 As Long, ByVal b32 As Long, ByVal c32 As Long) As Long
    silk_SMLATT = a32 + ((b32 And &HFFFF0000) \ &H10000) * ((c32 And &HFFFF0000) \ &H10000)
End Function

' (VB) silk_RSHIFT64( silk_SMULL( a32, b32 ), s ): (a32 * b32) >> s with the 64-bit product, exact for
' 9 <= s <= 63 (the product's 16-bit halves are exact in Double)
Public Function silk_SMULL_RSHIFT(ByVal a32 As Long, ByVal b32 As Long, ByVal s As Long) As Double
    Dim x As Double, y As Double
    x = CDbl(a32) * CDbl((b32 And &HFFFF0000) \ &H10000)        ' a32 * (b32 >> 16), below 2^46
    y = CDbl(a32) * CDbl(b32 And &HFFFF&)                       ' a32 * (b32 & 0xFFFF), below 2^47
    If s >= 16 Then
        silk_SMULL_RSHIFT = Int((x + Int(y / 65536#)) / pow2d(s - 16))
    Else
        silk_SMULL_RSHIFT = x * pow2d(16 - s) + Int(y / pow2d(s))
    End If
End Function

' (VB) silk_RSHIFT_ROUND64( silk_SMULL( a32, b32 ), shift ), 10 <= shift <= 64
Public Function silk_SMULL_RSHIFT_ROUND64(ByVal a32 As Long, ByVal b32 As Long, ByVal shift As Long) As Double
    silk_SMULL_RSHIFT_ROUND64 = Int((silk_SMULL_RSHIFT(a32, b32, shift - 1) + 1#) / 2#)
End Function

'/* Adds two signed 32-bit values in a way that can overflow, while not relying on undefined behaviour
'   (just standard two's complement implementation-specific behaviour) */
' silk_ADD32_ovflw(a, b), silk_SUB32_ovflw(a, b), silk_MLA_ovflw(a32, b32, c32): a + b, a - b, a32 + b32 * c32
'/* Multiply-accumulate macros that allow overflow in the addition (ie, no asserts in debug mode) */
Public Function silk_SMLABB_ovflw(ByVal a32 As Long, ByVal b32 As Long, ByVal c32 As Long) As Long
    silk_SMLABB_ovflw = a32 + (((b32 And &HFFFF&) Xor &H8000&) - &H8000&) * (((c32 And &HFFFF&) Xor &H8000&) - &H8000&)
End Function

' silk_DIV32_16(a32, b16) and silk_DIV32(a32, b32): a32 \ b (C's division truncates as VB's \ does)

Public Function silk_SAT8(ByVal a As Long) As Long
    If a > silk_int8_MAX Then
        silk_SAT8 = silk_int8_MAX
    ElseIf a < silk_int8_MIN Then
        silk_SAT8 = silk_int8_MIN
    Else
        silk_SAT8 = a
    End If
End Function

Public Function silk_SAT16(ByVal a As Long) As Long
    If a > silk_int16_MAX Then
        silk_SAT16 = silk_int16_MAX
    ElseIf a < silk_int16_MIN Then
        silk_SAT16 = silk_int16_MIN
    Else
        silk_SAT16 = a
    End If
End Function

' (VB) silk_SAT32 of an opus_int64 held in a Double
Public Function silk_SAT32(ByVal a As Double) As Long
    If a > 2147483647# Then
        silk_SAT32 = silk_int32_MAX
    ElseIf a < -2147483648# Then
        silk_SAT32 = silk_int32_MIN
    Else
        silk_SAT32 = a
    End If
End Function

Public Function silk_ADD_SAT16(ByVal a As Long, ByVal b As Long) As Long
    silk_ADD_SAT16 = silk_SAT16(a + b)
End Function

Public Function silk_SUB_SAT16(ByVal a As Long, ByVal b As Long) As Long
    silk_SUB_SAT16 = silk_SAT16(a - b)
End Function

'/* Saturation for positive input values */
Public Function silk_POS_SAT32(ByVal a As Double) As Long
    If a > 2147483647# Then silk_POS_SAT32 = silk_int32_MAX Else silk_POS_SAT32 = a
End Function

'/* Add with saturation for positive input values */
Public Function silk_ADD_POS_SAT8(ByVal a As Long, ByVal b As Long) As Long
    If ((a + b) And &H80&) <> 0 Then silk_ADD_POS_SAT8 = silk_int8_MAX Else silk_ADD_POS_SAT8 = a + b
End Function

Public Function silk_ADD_POS_SAT16(ByVal a As Long, ByVal b As Long) As Long
    If ((a + b) And &H8000&) <> 0 Then silk_ADD_POS_SAT16 = silk_int16_MAX Else silk_ADD_POS_SAT16 = a + b
End Function

Public Function silk_ADD_POS_SAT32(ByVal a As Long, ByVal b As Long) As Long
    If ((a + b) And &H80000000) <> 0 Then silk_ADD_POS_SAT32 = silk_int32_MAX Else silk_ADD_POS_SAT32 = a + b
End Function

Public Function silk_LSHIFT8(ByVal a As Long, ByVal shift As Long) As Long     '/* shift >= 0, shift < 8  */
    silk_LSHIFT8 = ((((a * pow2(shift)) And &HFF&) Xor &H80&) - &H80&)
End Function

Public Function silk_LSHIFT16(ByVal a As Long, ByVal shift As Long) As Long    '/* shift >= 0, shift < 16 */
    silk_LSHIFT16 = ((((a * pow2(shift)) And &HFFFF&) Xor &H8000&) - &H8000&)
End Function

Public Function silk_LSHIFT32(ByVal a As Long, ByVal shift As Long) As Long    '/* shift >= 0, shift < 32 */
    silk_LSHIFT32 = a * pow2(shift)
End Function

Public Function silk_LSHIFT(ByVal a As Long, ByVal shift As Long) As Long      '/* shift >= 0, shift < 32 */
    silk_LSHIFT = a * pow2(shift)
End Function

Public Function silk_RSHIFT32(ByVal a As Long, ByVal shift As Long) As Long    '/* shift >= 0, shift < 32 */
    silk_RSHIFT32 = silk_RSHIFT(a, shift)
End Function

Public Function silk_RSHIFT(ByVal a As Long, ByVal shift As Long) As Long      '/* shift >= 0, shift < 32 */
    If shift = 0 Then
        silk_RSHIFT = a
    ElseIf shift = 31 Then
        If a < 0 Then silk_RSHIFT = -1 Else silk_RSHIFT = 0
    Else
        silk_RSHIFT = (a And -pow2(shift)) \ pow2(shift)
    End If
End Function

'/* saturates before shifting */
Public Function silk_LSHIFT_SAT32(ByVal a As Long, ByVal shift As Long) As Long
    silk_LSHIFT_SAT32 = silk_LIMIT(a, silk_RSHIFT(silk_int32_MIN, shift), silk_RSHIFT(silk_int32_MAX, shift)) * pow2(shift)
End Function

Public Function silk_LSHIFT_ovflw(ByVal a As Long, ByVal shift As Long) As Long '/* shift >= 0, allowed to overflow */
    silk_LSHIFT_ovflw = a * pow2(shift)
End Function

Public Function silk_LSHIFT_uint(ByVal a As Long, ByVal shift As Long) As Long  '/* shift >= 0 */
    silk_LSHIFT_uint = a * pow2(shift)
End Function

' (VB) a logical shift: the Long holds an opus_uint32
Public Function silk_RSHIFT_uint(ByVal a As Long, ByVal shift As Long) As Long  '/* shift >= 0 */
    If shift = 0 Then
        silk_RSHIFT_uint = a
    ElseIf shift >= 32 Then
        silk_RSHIFT_uint = 0
    ElseIf a >= 0 Then
        silk_RSHIFT_uint = a \ pow2(shift)
    Else
        silk_RSHIFT_uint = ((a And &H7FFFFFFF) \ pow2(shift)) Or pow2(31 - shift)
    End If
End Function

Public Function silk_ADD_LSHIFT(ByVal a As Long, ByVal b As Long, ByVal shift As Long) As Long
    silk_ADD_LSHIFT = a + b * pow2(shift)
End Function

Public Function silk_ADD_LSHIFT32(ByVal a As Long, ByVal b As Long, ByVal shift As Long) As Long
    silk_ADD_LSHIFT32 = a + b * pow2(shift)
End Function

Public Function silk_ADD_LSHIFT_uint(ByVal a As Long, ByVal b As Long, ByVal shift As Long) As Long
    silk_ADD_LSHIFT_uint = a + b * pow2(shift)
End Function

Public Function silk_ADD_RSHIFT(ByVal a As Long, ByVal b As Long, ByVal shift As Long) As Long
    silk_ADD_RSHIFT = a + silk_RSHIFT(b, shift)
End Function

Public Function silk_ADD_RSHIFT32(ByVal a As Long, ByVal b As Long, ByVal shift As Long) As Long
    silk_ADD_RSHIFT32 = a + silk_RSHIFT(b, shift)
End Function

Public Function silk_ADD_RSHIFT_uint(ByVal a As Long, ByVal b As Long, ByVal shift As Long) As Long
    silk_ADD_RSHIFT_uint = a + silk_RSHIFT_uint(b, shift)
End Function

Public Function silk_SUB_LSHIFT32(ByVal a As Long, ByVal b As Long, ByVal shift As Long) As Long
    silk_SUB_LSHIFT32 = a - b * pow2(shift)
End Function

Public Function silk_SUB_RSHIFT32(ByVal a As Long, ByVal b As Long, ByVal shift As Long) As Long
    silk_SUB_RSHIFT32 = a - silk_RSHIFT(b, shift)
End Function

'/* Requires that shift > 0 */
Public Function silk_RSHIFT_ROUND(ByVal a As Long, ByVal shift As Long) As Long
    If shift = 1 Then
        silk_RSHIFT_ROUND = ((a And &HFFFFFFFE) \ 2) + (a And 1)
    Else
        silk_RSHIFT_ROUND = silk_RSHIFT(silk_RSHIFT(a, shift - 1) + 1, 1)
    End If
End Function

' (VB) silk_RSHIFT_ROUND64 of an opus_int64 held in a Double (an integer below 2^53)
Public Function silk_RSHIFT_ROUND64(ByVal a As Double, ByVal shift As Long) As Double
    If shift = 1 Then
        silk_RSHIFT_ROUND64 = Int(a / 2#) + (a - Int(a / 2#) * 2#)
    Else
        silk_RSHIFT_ROUND64 = Int((Int(a / pow2d(shift - 1)) + 1#) / 2#)
    End If
End Function

'/* Number of rightshift required to fit the multiplication */
Public Function silk_NSHIFT_MUL_32_32(ByVal a As Long, ByVal b As Long) As Long
    silk_NSHIFT_MUL_32_32 = -(31 - (32 - silk_CLZ32(silk_abs(a)) + (32 - silk_CLZ32(silk_abs(b)))))
End Function

Public Function silk_NSHIFT_MUL_16_16(ByVal a As Long, ByVal b As Long) As Long
    silk_NSHIFT_MUL_16_16 = -(15 - (16 - silk_CLZ16(silk_abs(a)) + (16 - silk_CLZ16(silk_abs(b)))))
End Function

Public Function silk_min(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then silk_min = a Else silk_min = b
End Function

Public Function silk_max(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then silk_max = a Else silk_max = b
End Function

'/* silk_min() versions with typecast in the function call */
Public Function silk_min_int(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then silk_min_int = a Else silk_min_int = b
End Function

Public Function silk_min_16(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then silk_min_16 = a Else silk_min_16 = b
End Function

Public Function silk_min_32(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then silk_min_32 = a Else silk_min_32 = b
End Function

'/* silk_min() versions with typecast in the function call */
Public Function silk_max_int(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then silk_max_int = a Else silk_max_int = b
End Function

Public Function silk_max_16(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then silk_max_16 = a Else silk_max_16 = b
End Function

Public Function silk_max_32(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then silk_max_32 = a Else silk_max_32 = b
End Function

Public Function silk_LIMIT(ByVal a As Long, ByVal limit1 As Long, ByVal limit2 As Long) As Long
    If limit1 > limit2 Then
        If a > limit1 Then
            silk_LIMIT = limit1
        ElseIf a < limit2 Then
            silk_LIMIT = limit2
        Else
            silk_LIMIT = a
        End If
    Else
        If a > limit2 Then
            silk_LIMIT = limit2
        ElseIf a < limit1 Then
            silk_LIMIT = limit1
        Else
            silk_LIMIT = a
        End If
    End If
End Function

' silk_LIMIT_int, silk_LIMIT_16 and silk_LIMIT_32 are silk_LIMIT

'/* Be careful, silk_abs returns wrong when input equals to silk_intXX_MIN */
Public Function silk_abs(ByVal a As Long) As Long
    If a > 0 Then silk_abs = a Else silk_abs = -a
End Function

Public Function silk_abs_int32(ByVal a As Long) As Long
    Dim s As Long
    If a < 0 Then s = -1
    silk_abs_int32 = (a Xor s) - s
End Function

Public Function silk_sign(ByVal a As Long) As Long
    If a > 0 Then
        silk_sign = 1
    ElseIf a < 0 Then
        silk_sign = -1
    End If
End Function

'/* PSEUDO-RANDOM GENERATOR                                                          */
'/* Make sure to store the result as the seed for the next call (also in between     */
'/* frames), otherwise result won't be random at all. When only using some of the    */
'/* bits, take the most significant bits by right-shifting.                          */
Public Function silk_RAND(ByVal seed As Long) As Long
    silk_RAND = RAND_INCREMENT + seed * RAND_MULTIPLIER
End Function

'/*    silk_SMMUL: Signed top word multiply.
'          ARMv6        2 instruction cycles.
'          ARMv3M+      3 instruction cycles. use SMULL and ignore LSB registers.(except xM)*/
'/* the following seems faster on x86 */
Public Function silk_SMMUL(ByVal a32 As Long, ByVal b32 As Long) As Long
    silk_SMMUL = silk_SMULL_RSHIFT(a32, b32, 32)
End Function

' Inlines.h

'/* get number of leading zeros and fractional part (the bits right after the leading one */
Public Sub silk_CLZ_FRAC(ByVal in_ As Long, lz As Long, frac_Q7 As Long)
    Dim lzeros As Long
    lzeros = silk_CLZ32(in_)
    lz = lzeros
    frac_Q7 = silk_ROR32(in_, 24 - lzeros) And &H7F&
End Sub

'/* Approximation of square root                                          */
'/* Accuracy: < +/- 10%  for output values > 15                           */
'/*           < +/- 2.5% for output values > 120                          */
Public Function silk_SQRT_APPROX(ByVal x As Long) As Long
    Dim y As Long, lz As Long, frac_Q7 As Long
    If x <= 0 Then
        silk_SQRT_APPROX = 0
        Exit Function
    End If
    silk_CLZ_FRAC x, lz, frac_Q7
    If (lz And 1) <> 0 Then
        y = 32768
    Else
        y = 46214        '/* 46214 = sqrt(2) * 32768 */
    End If
    '/* get scaling right */
    y = silk_RSHIFT(y, silk_RSHIFT(lz, 1))
    '/* increment using fractional part of input */
    y = silk_SMLAWB(y, y, silk_SMULBB(213, frac_Q7))
    silk_SQRT_APPROX = y
End Function

'/* Divide two int32 values and return result as int32 in a given Q-domain */
Public Function silk_DIV32_varQ(ByVal a32 As Long, ByVal b32 As Long, ByVal Qres As Long) As Long '/* O    returns a good approximation of "(a32 << Qres) / b32" */
    Dim a_headrm As Long, b_headrm As Long, lshift As Long
    Dim b32_inv As Long, a32_nrm As Long, b32_nrm As Long, result As Long
    '/* Compute number of bits head room and normalize inputs */
    a_headrm = silk_CLZ32(silk_abs(a32)) - 1
    a32_nrm = a32 * pow2(a_headrm)                                              '/* Q: a_headrm                  */
    b_headrm = silk_CLZ32(silk_abs(b32)) - 1
    b32_nrm = b32 * pow2(b_headrm)                                              '/* Q: b_headrm                  */
    '/* Inverse of b32, with 14 bits of precision */
    b32_inv = (silk_int32_MAX \ 4) \ silk_RSHIFT(b32_nrm, 16)                   '/* Q: 29 + 16 - b_headrm        */
    '/* First approximation */
    result = silk_SMULWB(a32_nrm, b32_inv)                                      '/* Q: 29 + a_headrm - b_headrm  */
    '/* Compute residual by subtracting product of denominator and first approximation */
    '/* It's OK to overflow because the final value of a32_nrm should always be small */
    a32_nrm = a32_nrm - silk_SMMUL(b32_nrm, result) * 8                         '/* Q: a_headrm   */
    '/* Refinement */
    result = silk_SMLAWB(result, a32_nrm, b32_inv)                              '/* Q: 29 + a_headrm - b_headrm  */
    '/* Convert to Qres domain */
    lshift = 29 + a_headrm - b_headrm - Qres
    If lshift < 0 Then
        silk_DIV32_varQ = silk_LSHIFT_SAT32(result, -lshift)
    Else
        If lshift < 32 Then
            silk_DIV32_varQ = silk_RSHIFT(result, lshift)
        Else
            '/* Avoid undefined result */
            silk_DIV32_varQ = 0
        End If
    End If
End Function

'/* Invert int32 value and return result as int32 in a given Q-domain */
Public Function silk_INVERSE32_varQ(ByVal b32 As Long, ByVal Qres As Long) As Long '/* O    returns a good approximation of "(1 << Qres) / b32" */
    Dim b_headrm As Long, lshift As Long
    Dim b32_inv As Long, b32_nrm As Long, err_Q32 As Long, result As Long
    '/* Compute number of bits head room and normalize input */
    b_headrm = silk_CLZ32(silk_abs(b32)) - 1
    b32_nrm = b32 * pow2(b_headrm)                                              '/* Q: b_headrm                */
    '/* Inverse of b32, with 14 bits of precision */
    b32_inv = (silk_int32_MAX \ 4) \ silk_RSHIFT(b32_nrm, 16)                   '/* Q: 29 + 16 - b_headrm    */
    '/* First approximation */
    result = b32_inv * &H10000                                                  '/* Q: 61 - b_headrm            */
    '/* Compute residual by subtracting product of denominator and first approximation from one */
    err_Q32 = (&H20000000 - silk_SMULWB(b32_nrm, b32_inv)) * 8                  '/* Q32                        */
    '/* Refinement */
    result = silk_SMLAWW(result, err_Q32, b32_inv)                              '/* Q: 61 - b_headrm            */
    '/* Convert to Qres domain */
    lshift = 61 - b_headrm - Qres
    If lshift <= 0 Then
        silk_INVERSE32_varQ = silk_LSHIFT_SAT32(result, -lshift)
    Else
        If lshift < 32 Then
            silk_INVERSE32_varQ = silk_RSHIFT(result, lshift)
        Else
            '/* Avoid undefined result */
            silk_INVERSE32_varQ = 0
        End If
    End If
End Function

'/* Useful Macros that can be adjusted to other platforms */
' (VB) silk_memcpy, silk_memset and silk_memmove on arrays: the C's pointers are an array and an offset; n is a
' count of elements (C: bytes). memset is only ever to 0. RtlMoveMemory handles overlaps, as memmove does.
Public Sub silk_memcpy_i16(dst() As Integer, ByVal dstOffset As Long, src() As Integer, ByVal srcOffset As Long, ByVal n As Long)
    If n > 0 Then CopyMemory dst(dstOffset), src(srcOffset), n * 2
End Sub

Public Sub silk_memcpy_i32(dst() As Long, ByVal dstOffset As Long, src() As Long, ByVal srcOffset As Long, ByVal n As Long)
    If n > 0 Then CopyMemory dst(dstOffset), src(srcOffset), n * 4
End Sub

Public Sub silk_memset_i16(dst() As Integer, ByVal dstOffset As Long, ByVal n As Long)
    If n > 0 Then ZeroMemory dst(dstOffset), n * 2
End Sub

Public Sub silk_memset_i32(dst() As Long, ByVal dstOffset As Long, ByVal n As Long)
    If n > 0 Then ZeroMemory dst(dstOffset), n * 4
End Sub
