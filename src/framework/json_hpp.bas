Attribute VB_Name = "json_hpp"
' 3rdparty/json/include/nlohmann/json.hpp (nlohmann::json 3.11.3), the parts
' Warzone uses, behind the nlohmann_json class. A third-party library, so
' like physfs_win it is rewritten for its behaviour rather than line by line:
' the lexer, the parser and the serializer give the library's values, key
' order, text, positions and exception messages (checked against the real
' library: tools/run_tests.py). dtoa_impl (Grisu2, which dump uses for
' doubles) is ported line by line, its uint64_t arithmetic in Decimal.
' Not ported: parser callbacks and comments (ignore_comments), the binary
' formats (CBOR, BSON, ...), JSON pointer and patch, diagnostics, binary
' values. (VB) strtod is cstdlib_vb's (Val), which may differ from the C
' library in the last bit for numbers with more than 15 significant digits.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private sbBuf As String
Private sbLen As Long
Private Type diyfp                      ' f * 2^e
    f As Variant                        ' std::uint64_t (Decimal)
    e As Long
End Type

Private Type boundaries
    w As diyfp
    minus As diyfp
    plus As diyfp
End Type

Private Type cached_power               ' c = f * 2^e ~= 10^k
    f As Variant
    e As Long
    k As Long
End Type

Private Const kAlpha As Long = -60
Private Const kGamma As Long = -32
Private pow2(0 To 64) As Variant        ' 2^n as Decimals
Private kCachedPowers(0 To 78) As cached_power
Private dtoaReady As Boolean
Private Const TT_UNINITIALIZED As Long = 0
Private Const TT_LITERAL_TRUE As Long = 1
Private Const TT_LITERAL_FALSE As Long = 2
Private Const TT_LITERAL_NULL As Long = 3
Private Const TT_VALUE_STRING As Long = 4
Private Const TT_VALUE_UNSIGNED As Long = 5
Private Const TT_VALUE_INTEGER As Long = 6
Private Const TT_VALUE_FLOAT As Long = 7
Private Const TT_BEGIN_ARRAY As Long = 8
Private Const TT_BEGIN_OBJECT As Long = 9
Private Const TT_END_ARRAY As Long = 10
Private Const TT_END_OBJECT As Long = 11
Private Const TT_NAME_SEPARATOR As Long = 12
Private Const TT_VALUE_SEPARATOR As Long = 13
Private Const TT_PARSE_ERROR As Long = 14
Private Const TT_END_OF_INPUT As Long = 15
Private Const TT_LITERAL_OR_VALUE As Long = 16
Private Const EOF_ As Long = -1
Private lxData() As Byte                ' the input, lxPos .. lxEnd - 1
Private lxPos As Long
Private lxEnd As Long
Private lxCurrent As Long
Private lxNextUnget As Boolean
Private lxCharsReadTotal As Long        ' position_t
Private lxCharsReadCurrentLine As Long
Private lxLinesRead As Long
Private lxTok() As Byte                 ' token_string: the raw bytes read for the token
Private lxTokLen As Long
Private lxBuf As String                 ' token_buffer, decoded (UTF-16)
Private lxBufLen As Long
Private lxErrorMessage As String
Private lxNumber As Variant             ' value_unsigned / value_integer / value_float
Private psLastToken As Long
Private psRefStack() As nlohmann_json
Private psRefCount As Long
Private psRoot As nlohmann_json
Private psObjectElement As Long          ' the slot key() made in the innermost object
Private psOrdered As Boolean

'--------------------------------------------------------------------------
' Exceptions: Err.Number is vbObjectError + id, Err.Description what()

' C: exception::name(ename, id) + what_arg
Public Sub json_throw(ByVal kind As String, ByVal id As Long, ByVal what_arg As String)
    Err.Raise vbObjectError + id, "nlohmann::json", "[json.exception." & kind & "." & id & "] " & what_arg
End Sub

' std::to_string(size_t) of an index the port holds in a Long
Public Function size_t_string(ByVal v As Long) As String
    If v < 0 Then size_t_string = CStr(CDec(v) + CDec("18446744073709551616")) Else size_t_string = CStr(v)
End Function

'--------------------------------------------------------------------------
' The output buffer of dump (output_string_adapter)

Public Sub sb_reset()
    sbLen = 0
    If Len(sbBuf) = 0 Then sbBuf = Space$(4096)
End Sub

Public Sub sb_add(ByVal s As String)
    Dim n As Long
    n = Len(s)
    If n = 0 Then Exit Sub
    If sbLen + n > Len(sbBuf) Then sbBuf = sbBuf & Space$(Len(sbBuf) + n)
    Mid$(sbBuf, sbLen + 1, n) = s
    sbLen = sbLen + n
End Sub

Public Function sb_result() As String
    sb_result = Left$(sbBuf, sbLen)
End Function

'--------------------------------------------------------------------------
' serializer

Private Function Hex4(ByVal v As Long) As String
    Hex4 = Right$("000" & LCase$(Hex$(v)), 4)
End Function

' serializer::dump_escaped: s's characters, escaped; with ensure_ascii every
' character from U+007F as \uxxxx (a surrogate pair beyond U+FFFF). A lone
' surrogate (C++: invalid UTF-8) follows error_handler.
Public Sub dump_escaped(ByVal s As String, ByVal ensure_ascii As Boolean, ByVal error_handler As error_handler_t)
    Dim i As Long, n As Long, c As Long, c2 As Long, start As Long, esc As String, used As Long, hasEsc As Boolean
    n = Len(s)
    start = 1
    i = 1
    Do While i <= n
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        used = 1
        hasEsc = True
        Select Case c
        Case &H8: esc = "\b"
        Case &H9: esc = "\t"
        Case &HA: esc = "\n"
        Case &HC: esc = "\f"
        Case &HD: esc = "\r"
        Case &H22: esc = "\"""
        Case &H5C: esc = "\\"
        Case Else
            If c <= &H1F Then
                esc = "\u" & Hex4(c)
            ElseIf c >= &HD800& And c <= &HDFFF& Then
                c2 = 0
                If c <= &HDBFF& And i < n Then c2 = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
                If c2 >= &HDC00& And c2 <= &HDFFF& Then
                    used = 2
                    If ensure_ascii Then
                        esc = "\u" & Hex4(c) & "\u" & Hex4(c2)
                    Else
                        hasEsc = False
                    End If
                Else
                    Select Case error_handler
                    Case error_handler_t_replace
                        If ensure_ascii Then esc = "�" Else esc = ChrW$(&HFFFD)
                    Case error_handler_t_ignore
                        esc = ""
                    Case Else
                        json_throw "type_error", 316, "invalid UTF-8 byte at index " & (i - 1) & ": 0x" & UCase$(Hex$(&HE0 Or (c \ &H1000)))
                    End Select
                End If
            ElseIf ensure_ascii And c >= &H7F Then
                esc = "\u" & Hex4(c)
            Else
                hasEsc = False
            End If
        End Select
        If hasEsc Then
            If i > start Then sb_add Mid$(s, start, i - start)
            sb_add esc
            start = i + used
        End If
        i = i + used
    Loop
    If n >= start Then sb_add Mid$(s, start)
End Sub

' serializer::dump_integer: every digit (a Double integer up to 2^53, a Decimal beyond)
Public Function dump_integer(ByVal v As Variant) As String
    Dim d As Double, hi As Double, lo As Double, s As String
    If VarType(v) = vbDecimal Then
        dump_integer = CStr(v)
        Exit Function
    End If
    d = v
    If Abs(d) < 1E+15 Then
        dump_integer = CStr(d)
        Exit Function
    End If
    hi = Int(Abs(d) / 100000000#)
    lo = Abs(d) - hi * 100000000#
    s = CStr(hi) & Right$("0000000" & CStr(lo), 8)
    If d < 0 Then s = "-" & s
    dump_integer = s
End Function

' serializer::dump_float: Grisu2 (to_chars); NaN and infinity as null
Public Function dump_float(ByVal x As Double) As String
    Dim hi As Long, lo As Long
    CopyMemory lo, x, 4
    CopyMemory hi, ByVal VarPtr(x) + 4, 4
    If (hi And &H7FF00000) = &H7FF00000 Then
        dump_float = "null"
        Exit Function
    End If
    dump_float = to_chars(x)
End Function

'--------------------------------------------------------------------------
' dtoa_impl (Grisu2), the library's code with uint64_t values as Decimals.

Private Function HexToDec(ByVal h As String) As Variant
    Dim i As Long, d As Variant
    d = CDec(0)
    For i = 1 To Len(h)
        d = d * 16 + CLng("&H" & Mid$(h, i, 1))
    Next
    HexToDec = d
End Function

Private Sub InitDtoa()
    Dim i As Long, rows() As String, parts() As String, t As String
    If dtoaReady Then Exit Sub
    pow2(0) = CDec(1)
    For i = 1 To 64
        pow2(i) = pow2(i - 1) * 2
    Next
    ' kCachedPowers: f, e, k
    t = ""
    t = t & "AB70FE17C79AC6CA -1060 -300;FF77B1FCBEBCDC4F -1034 -292;BE5691EF416BD60C -1007 -284;"
    t = t & "8DD01FAD907FFC3C -980 -276;D3515C2831559A83 -954 -268;9D71AC8FADA6C9B5 -927 -260;"
    t = t & "EA9C227723EE8BCB -901 -252;AECC49914078536D -874 -244;823C12795DB6CE57 -847 -236;"
    t = t & "C21094364DFB5637 -821 -228;9096EA6F3848984F -794 -220;D77485CB25823AC7 -768 -212;"
    t = t & "A086CFCD97BF97F4 -741 -204;EF340A98172AACE5 -715 -196;B23867FB2A35B28E -688 -188;"
    t = t & "84C8D4DFD2C63F3B -661 -180;C5DD44271AD3CDBA -635 -172;936B9FCEBB25C996 -608 -164;"
    t = t & "DBAC6C247D62A584 -582 -156;A3AB66580D5FDAF6 -555 -148;F3E2F893DEC3F126 -529 -140;"
    t = t & "B5B5ADA8AAFF80B8 -502 -132;87625F056C7C4A8B -475 -124;C9BCFF6034C13053 -449 -116;"
    t = t & "964E858C91BA2655 -422 -108;DFF9772470297EBD -396 -100;A6DFBD9FB8E5B88F -369 -92;"
    t = t & "F8A95FCF88747D94 -343 -84;B94470938FA89BCF -316 -76;8A08F0F8BF0F156B -289 -68;"
    t = t & "CDB02555653131B6 -263 -60;993FE2C6D07B7FAC -236 -52;E45C10C42A2B3B06 -210 -44;"
    t = t & "AA242499697392D3 -183 -36;FD87B5F28300CA0E -157 -28;BCE5086492111AEB -130 -20;"
    t = t & "8CBCCC096F5088CC -103 -12;D1B71758E219652C -77 -4;9C40000000000000 -50 4;"
    t = t & "E8D4A51000000000 -24 12;AD78EBC5AC620000 3 20;813F3978F8940984 30 28;"
    t = t & "C097CE7BC90715B3 56 36;8F7E32CE7BEA5C70 83 44;D5D238A4ABE98068 109 52;"
    t = t & "9F4F2726179A2245 136 60;ED63A231D4C4FB27 162 68;B0DE65388CC8ADA8 189 76;"
    t = t & "83C7088E1AAB65DB 216 84;C45D1DF942711D9A 242 92;924D692CA61BE758 269 100;"
    t = t & "DA01EE641A708DEA 295 108;A26DA3999AEF774A 322 116;F209787BB47D6B85 348 124;"
    t = t & "B454E4A179DD1877 375 132;865B86925B9BC5C2 402 140;C83553C5C8965D3D 428 148;"
    t = t & "952AB45CFA97A0B3 455 156;DE469FBD99A05FE3 481 164;A59BC234DB398C25 508 172;"
    t = t & "F6C69A72A3989F5C 534 180;B7DCBF5354E9BECE 561 188;88FCF317F22241E2 588 196;"
    t = t & "CC20CE9BD35C78A5 614 204;98165AF37B2153DF 641 212;E2A0B5DC971F303A 667 220;"
    t = t & "A8D9D1535CE3B396 694 228;FB9B7CD9A4A7443C 720 236;BB764C4CA7A44410 747 244;"
    t = t & "8BAB8EEFB6409C1A 774 252;D01FEF10A657842C 800 260;9B10A4E5E9913129 827 268;"
    t = t & "E7109BFBA19C0C9D 853 276;AC2820D9623BF429 880 284;80444B5E7AA7CF85 907 292;"
    t = t & "BF21E44003ACDD2D 933 300;8E679C2F5E44FF8F 960 308;D433179D9C8CB841 986 316;"
    t = t & "9E19DB92B4E31BA9 1013 324;"
    rows = Split(Left$(t, Len(t) - 1), ";")
    For i = 0 To 78
        parts = Split(rows(i), " ")
        kCachedPowers(i).f = HexToDec(parts(0))
        kCachedPowers(i).e = CLng(parts(1))
        kCachedPowers(i).k = CLng(parts(2))
    Next
    dtoaReady = True
End Sub

' x >> n (floor), x & (2^n - 1), x << n
Private Function ShrDec(ByVal x As Variant, ByVal n As Long) As Variant
    Dim q As Variant, r As Variant
    q = Int(x / pow2(n))
    r = x - q * pow2(n)
    If r < 0 Then q = q - 1
    If r >= pow2(n) Then q = q + 1
    ShrDec = q
End Function

Private Function MaskDec(ByVal x As Variant, ByVal n As Long) As Variant
    MaskDec = x - ShrDec(x, n) * pow2(n)
End Function

Private Function MakeDiyfp(ByVal f As Variant, ByVal e As Long) As diyfp
    MakeDiyfp.f = f
    MakeDiyfp.e = e
End Function

' returns x - y
Private Function diyfp_sub(x As diyfp, y As diyfp) As diyfp
    diyfp_sub = MakeDiyfp(x.f - y.f, x.e)
End Function

' returns x * y; the result is rounded (only the upper q bits are returned)
Private Function diyfp_mul(x As diyfp, y As diyfp) As diyfp
    Dim u_lo As Variant, u_hi As Variant, v_lo As Variant, v_hi As Variant
    Dim p0 As Variant, p1 As Variant, p2 As Variant, p3 As Variant
    Dim p0_hi As Variant, p1_lo As Variant, p1_hi As Variant, p2_lo As Variant, p2_hi As Variant
    Dim Q As Variant, h As Variant

    u_lo = MaskDec(x.f, 32)
    u_hi = ShrDec(x.f, 32)
    v_lo = MaskDec(y.f, 32)
    v_hi = ShrDec(y.f, 32)

    p0 = u_lo * v_lo
    p1 = u_lo * v_hi
    p2 = u_hi * v_lo
    p3 = u_hi * v_hi

    p0_hi = ShrDec(p0, 32)
    p1_lo = MaskDec(p1, 32)
    p1_hi = ShrDec(p1, 32)
    p2_lo = MaskDec(p2, 32)
    p2_hi = ShrDec(p2, 32)

    Q = p0_hi + p1_lo + p2_lo

    ' The full product might now be computed as
    ' p_hi = p3 + p2_hi + p1_hi + (Q >> 32)
    ' p_lo = p0_lo + (Q << 32)
    ' But since we only need the upper 64 bits of the product, the result is rounded:
    Q = Q + pow2(31) ' round, ties up

    h = p3 + p2_hi + p1_hi + ShrDec(Q, 32)

    diyfp_mul = MakeDiyfp(h, x.e + y.e + 64)
End Function

' normalize x such that the significand is >= 2^(q-1)
Private Function diyfp_normalize(ByVal xf As Variant, ByVal xe As Long) As diyfp
    Do While xf < pow2(63)
        xf = xf * 2
        xe = xe - 1
    Loop
    diyfp_normalize = MakeDiyfp(xf, xe)
End Function

' normalize x such that the result has the exponent E
Private Function diyfp_normalize_to(x As diyfp, ByVal target_exponent As Long) As diyfp
    Dim delta As Long
    delta = x.e - target_exponent
    diyfp_normalize_to = MakeDiyfp(x.f * pow2(delta), target_exponent)
End Function

' Compute the (normalized) diyfp representing the input number 'value' and its boundaries.
Private Function compute_boundaries(ByVal value As Double) As boundaries
    Const kBias As Long = 1075                  ' max_exponent - 1 + (kPrecision - 1)
    Const kMinExp As Long = -1074               ' 1 - kBias
    Dim hi As Long, lo As Long, e As Long, f As Variant, is_denormal As Boolean
    Dim v As diyfp, m_plus As diyfp, m_minus As diyfp, lower_boundary_is_closer As Boolean, w_plus As diyfp

    CopyMemory lo, value, 4
    CopyMemory hi, ByVal VarPtr(value) + 4, 4
    e = (hi And &H7FFFFFFF) \ &H100000          ' E = bits >> (kPrecision - 1)
    f = CDec(hi And &HFFFFF) * pow2(32) + CDec(U32(lo)) ' F = bits & (kHiddenBit - 1)

    is_denormal = (e = 0)
    If is_denormal Then
        v = MakeDiyfp(f, kMinExp)
    Else
        v = MakeDiyfp(f + pow2(52), e - kBias)
    End If

    ' Compute the boundaries m- and m+ of the floating-point value
    ' v = f * 2^e.
    lower_boundary_is_closer = (f = 0 And e > 1)
    m_plus = MakeDiyfp(2 * v.f + 1, v.e - 1)
    If lower_boundary_is_closer Then
        m_minus = MakeDiyfp(4 * v.f - 1, v.e - 2) ' (B)
    Else
        m_minus = MakeDiyfp(2 * v.f - 1, v.e - 1) ' (A)
    End If

    ' Determine the normalized w+ = m+.
    w_plus = diyfp_normalize(m_plus.f, m_plus.e)

    ' Determine w- = m- such that e_(w-) = e_(w+).
    compute_boundaries.minus = diyfp_normalize_to(m_minus, w_plus.e)
    compute_boundaries.w = diyfp_normalize(v.f, v.e)
    compute_boundaries.plus = w_plus
End Function

' For a normalized diyfp w = f * 2^e, this function returns a (normalized) cached
' power-of-ten c = f_c * 2^e_c, such that the exponent of the product w * c
' satisfies alpha <= e_c + e + q <= gamma.
Private Function get_cached_power_for_binary_exponent(ByVal e As Long) As cached_power
    Const kCachedPowersMinDecExp As Long = -300
    Const kCachedPowersDecStep As Long = 8
    Dim f As Long, k As Long, index As Long

    f = kAlpha - e - 1
    k = (f * 78913) \ 262144                    ' (f * 78913) / (1 << 18), truncated
    If f > 0 Then k = k + 1
    index = (-kCachedPowersMinDecExp + k + (kCachedPowersDecStep - 1)) \ kCachedPowersDecStep
    get_cached_power_for_binary_exponent = kCachedPowers(index)
End Function

' For n != 0, returns k, such that pow10 := 10^(k-1) <= n < 10^k.
' For n == 0, returns 1 and sets pow10 := 1.
Private Function find_largest_pow10(ByVal n As Double, pow10 As Double) As Long
    If n >= 1000000000# Then
        pow10 = 1000000000#
        find_largest_pow10 = 10
    ElseIf n >= 100000000# Then
        pow10 = 100000000#
        find_largest_pow10 = 9
    ElseIf n >= 10000000# Then
        pow10 = 10000000#
        find_largest_pow10 = 8
    ElseIf n >= 1000000# Then
        pow10 = 1000000#
        find_largest_pow10 = 7
    ElseIf n >= 100000# Then
        pow10 = 100000#
        find_largest_pow10 = 6
    ElseIf n >= 10000# Then
        pow10 = 10000#
        find_largest_pow10 = 5
    ElseIf n >= 1000# Then
        pow10 = 1000#
        find_largest_pow10 = 4
    ElseIf n >= 100# Then
        pow10 = 100#
        find_largest_pow10 = 3
    ElseIf n >= 10# Then
        pow10 = 10#
        find_largest_pow10 = 2
    Else
        pow10 = 1#
        find_largest_pow10 = 1
    End If
End Function

Private Sub grisu2_round(buf() As Long, ByVal length As Long, ByVal dist As Variant, ByVal delta As Variant, ByVal rest As Variant, ByVal ten_k As Variant)
    '               <--------------------------- delta ---->
    '                                  <---- dist --------->
    ' --------------[------------------+-------------------]--------------
    '               M-                 w                   M+
    '
    '                                  ten_k
    '                                <------>
    '                                       <---- rest ---->
    ' --------------[------------------+----+--------------]--------------
    '                                  w    V
    '                                       = buf * 10^k
    '
    ' ten_k represents a unit-in-the-last-place in the decimal representation
    ' stored in buf.
    ' Decrement buf by ten_k while this takes buf closer to w.
    Do While rest < dist And delta - rest >= ten_k And (rest + ten_k < dist Or dist - rest > rest + ten_k - dist)
        buf(length - 1) = buf(length - 1) - 1
        rest = rest + ten_k
    Loop
End Sub

' Generates V = buffer * 10^decimal_exponent, such that M- <= V <= M+.
' M- and M+ must be normalized and share the same exponent -60 <= e <= -32.
Private Sub grisu2_digit_gen(buffer() As Long, length As Long, decimal_exponent As Long, M_minus As diyfp, w As diyfp, M_plus As diyfp)
    Dim delta As Variant, dist As Variant, one As diyfp, p1 As Double, p2 As Variant
    Dim pow10 As Double, k As Long, n As Long, d As Double, r As Double, rest As Variant, ten_n As Variant
    Dim m As Long, dd As Variant, rr As Variant, ten_m As Variant

    delta = diyfp_sub(M_plus, M_minus).f ' (significand of (M+ - M-), implicit exponent is e)
    dist = diyfp_sub(M_plus, w).f        ' (significand of (M+ - w ), implicit exponent is e)

    ' Split M+ = f * 2^e into two parts p1 and p2 (note: e < 0):
    '      M+ = f * 2^e
    '         = ((f div 2^-e) * 2^-e + (f mod 2^-e)) * 2^e
    '         = ((p1        ) * 2^-e + (p2        )) * 2^e
    '         = p1 + p2 * 2^e
    one = MakeDiyfp(pow2(-M_plus.e), M_plus.e)

    p1 = CDbl(ShrDec(M_plus.f, -one.e)) ' p1 = f div 2^-e (Since -e >= 32, p1 fits into a 32-bit int.)
    p2 = MaskDec(M_plus.f, -one.e)      ' p2 = f mod 2^-e

    ' 1)
    ' Generate the digits of the integral part p1 = d[n-1]...d[1]d[0]
    k = find_largest_pow10(p1, pow10)

    '      10^(k-1) <= p1 < 10^k, pow10 = 10^(k-1)
    '
    '      p1 = (p1 div 10^(k-1)) * 10^(k-1) + (p1 mod 10^(k-1))
    '         = (d[k-1]         ) * 10^(k-1) + (p1 mod 10^(k-1))
    '
    '      M+ = p1                                             + p2 * 2^e
    '         = d[k-1] * 10^(k-1) + (p1 mod 10^(k-1))          + p2 * 2^e
    '         = d[k-1] * 10^(k-1) + ((p1 mod 10^(k-1)) * 2^-e + p2) * 2^e
    '         = d[k-1] * 10^(k-1) + (                         rest) * 2^e
    '
    ' Now generate the digits d[n] of p1 from left to right (n = k-1,...,0)
    '
    '      p1 = d[k-1]...d[n] * 10^n + d[n-1]...d[0]
    '
    ' but stop as soon as
    '
    '      rest * 2^e = (d[n-1]...d[0] * 2^-e + p2) * 2^e <= delta * 2^e

    n = k
    Do While n > 0
        ' Invariants:
        '      M+ = buffer * 10^n + (p1 + p2 * 2^e)    (buffer = 0 for n = k)
        '      pow10 = 10^(n-1) <= p1 < 10^n
        '
        d = Int(p1 / pow10)  ' d = p1 div 10^(n-1)
        r = p1 - d * pow10   ' r = p1 mod 10^(n-1)
        '
        '      M+ = buffer * 10^n + (d * 10^(n-1) + r) + p2 * 2^e
        '         = (buffer * 10 + d) * 10^(n-1) + (r + p2 * 2^e)
        '
        buffer(length) = d   ' buffer := buffer * 10 + d
        length = length + 1
        '
        '      M+ = buffer * 10^(n-1) + (r + p2 * 2^e)
        '
        p1 = r
        n = n - 1
        '
        '      M+ = buffer * 10^n + (p1 + p2 * 2^e)
        '      pow10 = 10^n
        '

        ' Now check if enough digits have been generated.
        ' Compute
        '
        '      p1 + p2 * 2^e = (p1 * 2^-e + p2) * 2^e = rest * 2^e
        '
        ' Note:
        ' Since rest and delta share the same exponent e, it suffices to
        ' compare the significands.
        rest = CDec(p1) * pow2(-one.e) + p2
        If rest <= delta Then
            ' V = buffer * 10^n, with M- <= V <= M+.

            decimal_exponent = decimal_exponent + n

            ' We may now just stop. But instead look if the buffer could be
            ' decremented to bring V closer to w.
            '
            ' pow10 = 10^n is now 1 ulp in the decimal representation V.
            ' The rounding procedure works with diyfp's with an implicit
            ' exponent of e.
            '
            '      10^n = (10^n * 2^-e) * 2^e = ulp * 2^e
            '
            ten_n = CDec(pow10) * pow2(-one.e)
            grisu2_round buffer, length, dist, delta, rest, ten_n

            Exit Sub
        End If

        pow10 = pow10 / 10#
        '
        '      pow10 = 10^(n-1) <= p1 < 10^n
        ' Invariants restored.
    Loop

    ' 2)
    '
    ' The digits of the integral part have been generated:
    '
    '      M+ = d[k-1]...d[1]d[0] + p2 * 2^e
    '         = buffer            + p2 * 2^e
    '
    ' Now generate the digits of the fractional part p2 * 2^e.
    '
    ' Note:
    ' No decimal point is generated: the exponent is adjusted instead.
    '
    ' p2 actually represents the fraction
    '
    '      p2 * 2^e
    '          = p2 / 2^-e
    '          = d[-1] / 10^1 + d[-2] / 10^2 + ...
    '
    ' Now generate the digits d[-m] of p1 from left to right (m = 1,2,...)
    '
    '      p2 * 2^e = d[-1]d[-2]...d[-m] * 10^-m
    '                      + 10^-m * (d[-m-1] / 10^1 + d[-m-2] / 10^2 + ...)
    '
    ' using
    '
    '      10^m * p2 = ((10^m * p2) div 2^-e) * 2^-e + ((10^m * p2) mod 2^-e)
    '                = (                   d) * 2^-e + (                   r)
    '
    ' or
    '      10^m * p2 * 2^e = d + r * 2^e
    '
    ' i.e.
    '
    '      M+ = buffer + p2 * 2^e
    '         = buffer + 10^-m * (d + r * 2^e)
    '         = (buffer * 10^m + d) * 10^-m + 10^-m * r * 2^e
    '
    ' and stop as soon as 10^-m * r * 2^e <= delta * 2^e

    m = 0
    Do
        ' Invariant:
        '      M+ = buffer * 10^-m + 10^-m * (d[-m-1] / 10 + d[-m-2] / 10^2 + ...) * 2^e
        '         = buffer * 10^-m + 10^-m * (p2                                 ) * 2^e
        '         = buffer * 10^-m + 10^-m * (1/10 * (10 * p2)                   ) * 2^e
        '         = buffer * 10^-m + 10^-m * (1/10 * ((10*p2 div 2^-e) * 2^-e + (10*p2 mod 2^-e)) * 2^e
        '
        p2 = p2 * 10
        dd = ShrDec(p2, -one.e)     ' d = (10 * p2) div 2^-e
        rr = MaskDec(p2, -one.e)    ' r = (10 * p2) mod 2^-e
        '
        '      M+ = buffer * 10^-m + 10^-m * (1/10 * (d * 2^-e + r) * 2^e
        '         = buffer * 10^-m + 10^-m * (1/10 * (d + r * 2^e))
        '         = (buffer * 10 + d) * 10^(-m-1) + 10^(-m-1) * r * 2^e
        '
        buffer(length) = CLng(dd) ' buffer := buffer * 10 + d
        length = length + 1
        '
        '      M+ = buffer * 10^(-m-1) + 10^(-m-1) * r * 2^e
        '
        p2 = rr
        m = m + 1
        '
        '      M+ = buffer * 10^-m + 10^-m * p2 * 2^e
        ' Invariant restored.

        ' Check if enough digits have been generated.
        '
        '      10^-m * p2 * 2^e <= delta * 2^e
        '              p2 * 2^e <= 10^m * delta * 2^e
        '                    p2 <= 10^m * delta
        delta = delta * 10
        dist = dist * 10
        If p2 <= delta Then
            Exit Do
        End If
    Loop

    ' V = buffer * 10^-m, with M- <= V <= M+.

    decimal_exponent = decimal_exponent - m

    ' 1 ulp in the decimal representation is now 10^-m.
    ' Since delta and dist are now scaled by 10^m, we need to do the
    ' same with ulp in order to keep the units in sync.
    '
    '      10^m * 10^-m = 1 = 2^-e * 2^e = ten_m * 2^e
    '
    ten_m = one.f
    grisu2_round buffer, length, dist, delta, p2, ten_m
End Sub

' v = buf * 10^decimal_exponent
' len is the length of the buffer (number of decimal digits)
' The buffer must be large enough, i.e. >= max_digits10.
Private Sub grisu2(buf() As Long, length As Long, decimal_exponent As Long, ByVal value As Double)
    Dim w As boundaries, cached As cached_power, c_minus_k As diyfp
    Dim w_ As diyfp, w_minus As diyfp, w_plus As diyfp, M_minus As diyfp, M_plus As diyfp

    ' If the neighbors (and boundaries) of 'value' are always computed for double-precision
    ' numbers, all float's can be recovered using strtod (and strtof). However, the resulting
    ' decimal representations are not exactly "short".
    w = compute_boundaries(value)

    ' --------(-----------------------+-----------------------)--------    (A)
    '         m-                      v                       m+
    '
    ' --------------------(-----------+-----------------------)--------    (B)
    '                     m-          v                       m+
    '
    ' First scale v (and m- and m+) such that the exponent is in the range
    ' [alpha, gamma].

    cached = get_cached_power_for_binary_exponent(w.plus.e)

    c_minus_k = MakeDiyfp(cached.f, cached.e) ' = c ~= 10^-k

    ' The exponent of the products is = v.e + c_minus_k.e + q and is in the range [alpha,gamma]
    w_ = diyfp_mul(w.w, c_minus_k)
    w_minus = diyfp_mul(w.minus, c_minus_k)
    w_plus = diyfp_mul(w.plus, c_minus_k)

    '  ----(---+---)---------------(---+---)---------------(---+---)----
    '          w-                      w                       w+
    '          = c*m-                  = c*v                   = c*m+
    '
    ' diyfp::mul rounds its result and c_minus_k is approximated too. w, w- and
    ' w+ are now off by a small amount.
    ' In fact:
    '
    '      w - v * 10^k < 1 ulp
    '
    ' To account for this inaccuracy, add resp. subtract 1 ulp.
    '
    '  --------+---[---------------(---+---)---------------]---+--------
    '          w-  M-                  w                   M+  w+
    '
    ' Now any number in [M-, M+] (bounds included) will round to w when input,
    ' regardless of how the input rounding algorithm breaks ties.
    '
    ' And digit_gen generates the shortest possible such number in [M-, M+].
    ' Note that this does not mean that Grisu2 always generates the shortest
    ' possible number in the interval (m-, m+).
    M_minus = MakeDiyfp(w_minus.f + 1, w_minus.e)
    M_plus = MakeDiyfp(w_plus.f - 1, w_plus.e)

    decimal_exponent = -cached.k ' = -(-k) = k

    grisu2_digit_gen buf, length, decimal_exponent, M_minus, w_, M_plus
End Sub

' appends a decimal representation of e to buf
Private Function append_exponent(ByVal e As Long) As String
    Dim s As String, k As Long
    If e < 0 Then
        e = -e
        s = "-"
    Else
        s = "+"
    End If

    k = e
    If k < 10 Then
        ' Always print at least two digits in the exponent.
        ' This is for compatibility with printf("%g").
        s = s & "0" & CStr(k)
    ElseIf k < 100 Then
        s = s & CStr(k \ 10)
        k = k Mod 10
        s = s & CStr(k)
    Else
        s = s & CStr(k \ 100)
        k = k Mod 100
        s = s & CStr(k \ 10)
        k = k Mod 10
        s = s & CStr(k)
    End If

    append_exponent = s
End Function

' prettify v = buf * 10^decimal_exponent
' If v is in the range [10^min_exp, 10^max_exp) it will be printed in fixed-point
' notation. Otherwise it will be printed in exponential notation.
Private Function format_buffer(buf() As Long, ByVal length As Long, ByVal decimal_exponent As Long, ByVal min_exp As Long, ByVal max_exp As Long) As String
    Dim k As Long, n As Long, digits As String, i As Long

    k = length
    n = length + decimal_exponent

    ' v = buf * 10^(n-k)
    ' k is the length of the buffer (number of decimal digits)
    ' n is the position of the decimal point relative to the start of the buffer.
    For i = 0 To k - 1
        digits = digits & ChrW$(48 + buf(i))
    Next

    If k <= n And n <= max_exp Then
        ' digits[000]
        ' len <= max_exp + 2
        format_buffer = digits & String$(n - k, "0") & ".0"
        Exit Function
    End If

    If 0 < n And n <= max_exp Then
        ' dig.its
        ' len <= max_digits10 + 1
        format_buffer = Left$(digits, n) & "." & Mid$(digits, n + 1)
        Exit Function
    End If

    If min_exp < n And n <= 0 Then
        ' 0.[000]digits
        ' len <= 2 + (-min_exp - 1) + max_digits10
        format_buffer = "0." & String$(-n, "0") & digits
        Exit Function
    End If

    If k = 1 Then
        ' dE+123
        ' len <= 1 + 5
        format_buffer = digits & "e" & append_exponent(n - 1)
    Else
        ' d.igitsE+123
        ' len <= max_digits10 + 1 + 5
        format_buffer = Left$(digits, 1) & "." & Mid$(digits, 2) & "e" & append_exponent(n - 1)
    End If
End Function

' generates a decimal representation of the floating-point number value
' (finite), in the format of printf's %g, as short as Grisu2 makes it
Private Function to_chars(ByVal value As Double) As String
    Const kMinExp As Long = -4
    Const kMaxExp As Long = 15                  ' std::numeric_limits<double>::digits10
    Dim hi As Long, sign_ As String, length As Long, decimal_exponent As Long, buf(0 To 31) As Long

    InitDtoa
    ' Use signbit(value) instead of (value < 0) since signbit works for -0.
    CopyMemory hi, ByVal VarPtr(value) + 4, 4
    If hi < 0 Then
        value = -value
        sign_ = "-"
    End If

    If value = 0 Then ' +-0
        ' Make it look like a floating-point number (#362, #378)
        to_chars = sign_ & "0.0"
        Exit Function
    End If

    ' Compute v = buffer * 10^decimal_exponent.
    ' The decimal digits are stored in the buffer, which needs to be interpreted
    ' as an unsigned decimal integer.
    ' len is the length of the buffer, i.e. the number of decimal digits.
    length = 0
    decimal_exponent = 0
    grisu2 buf, length, decimal_exponent, value

    to_chars = sign_ & format_buffer(buf, length, decimal_exponent, kMinExp, kMaxExp)
End Function

'--------------------------------------------------------------------------
' lexer

Private Function token_type_name(ByVal t As Long) As String
    Select Case t
    Case TT_UNINITIALIZED: token_type_name = "<uninitialized>"
    Case TT_LITERAL_TRUE: token_type_name = "true literal"
    Case TT_LITERAL_FALSE: token_type_name = "false literal"
    Case TT_LITERAL_NULL: token_type_name = "null literal"
    Case TT_VALUE_STRING: token_type_name = "string literal"
    Case TT_VALUE_UNSIGNED, TT_VALUE_INTEGER, TT_VALUE_FLOAT: token_type_name = "number literal"
    Case TT_BEGIN_ARRAY: token_type_name = "'['"
    Case TT_BEGIN_OBJECT: token_type_name = "'{'"
    Case TT_END_ARRAY: token_type_name = "']'"
    Case TT_END_OBJECT: token_type_name = "'}'"
    Case TT_NAME_SEPARATOR: token_type_name = "':'"
    Case TT_VALUE_SEPARATOR: token_type_name = "','"
    Case TT_PARSE_ERROR: token_type_name = "<parse error>"
    Case TT_END_OF_INPUT: token_type_name = "end of input"
    Case TT_LITERAL_OR_VALUE: token_type_name = "'[', '{', or a literal"
    Case Else: token_type_name = "unknown token"
    End Select
End Function

Private Sub TokPush(ByVal c As Long)
    If lxTokLen > UBound(lxTok) Then ReDim Preserve lxTok(0 To 2 * lxTokLen - 1)
    lxTok(lxTokLen) = c
    lxTokLen = lxTokLen + 1
End Sub

' add a UTF-16 code unit to token_buffer
Private Sub add(ByVal c As Long)
    If lxBufLen >= Len(lxBuf) Then lxBuf = lxBuf & Space$(Len(lxBuf) + 64)
    Mid$(lxBuf, lxBufLen + 1, 1) = ChrW$(c)
    lxBufLen = lxBufLen + 1
End Sub

' add a code point (beyond U+FFFF as a surrogate pair)
Private Sub add_codepoint(ByVal cp As Long)
    If cp >= &H10000 Then
        cp = cp - &H10000
        add &HD800& + (cp \ &H400&)
        add &HDC00& + (cp And &H3FF&)
    Else
        add cp
    End If
End Sub

Private Sub reset()
    lxBufLen = 0
    lxTokLen = 0
    TokPush lxCurrent
End Sub

Private Function get_() As Long
    lxCharsReadTotal = lxCharsReadTotal + 1
    lxCharsReadCurrentLine = lxCharsReadCurrentLine + 1

    If lxNextUnget Then
        ' just reset the next_unget variable and work with current
        lxNextUnget = False
    ElseIf lxPos < lxEnd Then
        lxCurrent = lxData(lxPos)
        lxPos = lxPos + 1
    Else
        lxCurrent = EOF_
    End If

    If lxCurrent <> EOF_ Then
        TokPush lxCurrent
    End If

    If lxCurrent = 10 Then
        lxLinesRead = lxLinesRead + 1
        lxCharsReadCurrentLine = 0
    End If

    get_ = lxCurrent
End Function

Private Sub unget()
    lxNextUnget = True

    lxCharsReadTotal = lxCharsReadTotal - 1

    ' in case we "unget" a newline, we have to also decrement the lines_read
    If lxCharsReadCurrentLine = 0 Then
        If lxLinesRead > 0 Then
            lxLinesRead = lxLinesRead - 1
        End If
    Else
        lxCharsReadCurrentLine = lxCharsReadCurrentLine - 1
    End If

    If lxCurrent <> EOF_ Then
        lxTokLen = lxTokLen - 1
    End If
End Sub

' get_token_string: the bytes read for the token, control characters as <U+XXXX>
Private Function get_token_string() As String
    Dim b() As Byte, n As Long, i As Long, s As String, c As Long
    If lxTokLen = 0 Then Exit Function
    ReDim b(0 To lxTokLen * 8)
    For i = 0 To lxTokLen - 1
        c = lxTok(i)
        If c <= &H1F Then
            s = "<U+" & Right$("000" & Hex$(c), 4) & ">"
            For c = 1 To Len(s)
                b(n) = AscW(Mid$(s, c, 1))
                n = n + 1
            Next
        Else
            b(n) = c
            n = n + 1
        End If
    Next
    get_token_string = utf8_decode(b, 0, n)
End Function

Private Function get_string() As String
    get_string = Left$(lxBuf, lxBufLen)
End Function

' get the 4 hex digits after \u; -1 when one is not a hex digit
Private Function get_codepoint() As Long
    Dim codepoint As Long, factor As Long
    For factor = 12 To 0 Step -4
        get_
        If lxCurrent >= 48 And lxCurrent <= 57 Then
            codepoint = codepoint + (lxCurrent - &H30) * (2 ^ factor)
        ElseIf lxCurrent >= 65 And lxCurrent <= 70 Then
            codepoint = codepoint + (lxCurrent - &H37) * (2 ^ factor)
        ElseIf lxCurrent >= 97 And lxCurrent <= 102 Then
            codepoint = codepoint + (lxCurrent - &H57) * (2 ^ factor)
        Else
            get_codepoint = -1
            Exit Function
        End If
    Next
    get_codepoint = codepoint
End Function

' The UTF-8 sequence's continuation bytes, each in its range (lo1, hi1, lo2, hi2, ...)
Private Function next_byte_in_range(ParamArray ranges() As Variant) As Boolean
    Dim i As Long, cp As Long, lead As Long, n As Long
    lead = lxCurrent
    n = (UBound(ranges) + 1) \ 2
    Select Case n
    Case 1: cp = lead And &H1F
    Case 2: cp = lead And &HF
    Case Else: cp = lead And &H7
    End Select
    For i = 0 To UBound(ranges) Step 2
        get_
        If ranges(i) <= lxCurrent And lxCurrent <= ranges(i + 1) Then
            cp = cp * 64 + (lxCurrent And &H3F)
        Else
            lxErrorMessage = "invalid string: ill-formed UTF-8 byte"
            Exit Function
        End If
    Next
    add_codepoint cp
    next_byte_in_range = True
End Function

Private Function ControlCharacterMessage(ByVal c As Long) As String
    Dim names As Variant, h As String, extra As String
    names = Array("NUL", "SOH", "STX", "ETX", "EOT", "ENQ", "ACK", "BEL", "BS", "HT", "LF", "VT", "FF", "CR", "SO", "SI", _
                  "DLE", "DC1", "DC2", "DC3", "DC4", "NAK", "SYN", "ETB", "CAN", "EM", "SUB", "ESC", "FS", "GS", "RS", "US")
    h = Right$("000" & Hex$(c), 4)
    Select Case c
    Case 8: extra = " or \b"
    Case 9: extra = " or \t"
    Case 10: extra = " or \n"
    Case 12: extra = " or \f"
    Case 13: extra = " or \r"
    End Select
    ControlCharacterMessage = "invalid string: control character U+" & h & " (" & names(c) & ") must be escaped to \u" & h & extra
End Function

Private Function scan_string() As Long
    Dim codepoint1 As Long, codepoint2 As Long, codepoint As Long

    ' reset token_buffer (ignore opening quote)
    reset

    Do
        ' get next character
        Select Case get_()
        ' end of file while parsing string
        Case EOF_
            lxErrorMessage = "invalid string: missing closing quote"
            scan_string = TT_PARSE_ERROR
            Exit Function

        ' closing quote
        Case &H22
            scan_string = TT_VALUE_STRING
            Exit Function

        ' escapes
        Case &H5C
            Select Case get_()
            Case &H22: add &H22         ' quotation mark
            Case &H5C: add &H5C         ' reverse solidus
            Case &H2F: add &H2F         ' solidus
            Case &H62: add &H8          ' backspace
            Case &H66: add &HC          ' form feed
            Case &H6E: add &HA          ' line feed
            Case &H72: add &HD          ' carriage return
            Case &H74: add &H9          ' tab

            ' unicode escapes
            Case &H75
                codepoint1 = get_codepoint()
                codepoint = codepoint1 ' start with codepoint1

                If codepoint1 = -1 Then
                    lxErrorMessage = "invalid string: '\u' must be followed by 4 hex digits"
                    scan_string = TT_PARSE_ERROR
                    Exit Function
                End If

                ' check if code point is a high surrogate
                If &HD800& <= codepoint1 And codepoint1 <= &HDBFF& Then
                    ' expect next \uxxxx entry
                    If get_() = &H5C Then
                        If get_() = &H75 Then
                            codepoint2 = get_codepoint()

                            If codepoint2 = -1 Then
                                lxErrorMessage = "invalid string: '\u' must be followed by 4 hex digits"
                                scan_string = TT_PARSE_ERROR
                                Exit Function
                            End If

                            ' check if codepoint2 is a low surrogate
                            If &HDC00& <= codepoint2 And codepoint2 <= &HDFFF& Then
                                ' overwrite codepoint
                                codepoint = (codepoint1 - &HD800&) * &H400& + (codepoint2 - &HDC00&) + &H10000
                            Else
                                lxErrorMessage = "invalid string: surrogate U+D800..U+DBFF must be followed by U+DC00..U+DFFF"
                                scan_string = TT_PARSE_ERROR
                                Exit Function
                            End If
                        Else
                            lxErrorMessage = "invalid string: surrogate U+D800..U+DBFF must be followed by U+DC00..U+DFFF"
                            scan_string = TT_PARSE_ERROR
                            Exit Function
                        End If
                    Else
                        lxErrorMessage = "invalid string: surrogate U+D800..U+DBFF must be followed by U+DC00..U+DFFF"
                        scan_string = TT_PARSE_ERROR
                        Exit Function
                    End If
                ElseIf &HDC00& <= codepoint1 And codepoint1 <= &HDFFF& Then
                    lxErrorMessage = "invalid string: surrogate U+DC00..U+DFFF must follow U+D800..U+DBFF"
                    scan_string = TT_PARSE_ERROR
                    Exit Function
                End If

                ' result of the above calculation yields a proper codepoint
                add_codepoint codepoint

            ' other characters after escape
            Case Else
                lxErrorMessage = "invalid string: forbidden character after backslash"
                scan_string = TT_PARSE_ERROR
                Exit Function
            End Select

        ' invalid control characters
        Case &H0 To &H1F
            lxErrorMessage = ControlCharacterMessage(lxCurrent)
            scan_string = TT_PARSE_ERROR
            Exit Function

        ' U+0020..U+007F (except U+0022 (quote) and U+005C (backspace))
        Case &H20 To &H7F
            add lxCurrent

        ' U+0080..U+07FF: bytes C2..DF 80..BF
        Case &HC2 To &HDF
            If Not next_byte_in_range(&H80, &HBF) Then
                scan_string = TT_PARSE_ERROR
                Exit Function
            End If

        ' U+0800..U+0FFF: bytes E0 A0..BF 80..BF
        Case &HE0
            If Not next_byte_in_range(&HA0, &HBF, &H80, &HBF) Then
                scan_string = TT_PARSE_ERROR
                Exit Function
            End If

        ' U+1000..U+CFFF: bytes E1..EC 80..BF 80..BF
        ' U+E000..U+FFFF: bytes EE..EF 80..BF 80..BF
        Case &HE1 To &HEC, &HEE, &HEF
            If Not next_byte_in_range(&H80, &HBF, &H80, &HBF) Then
                scan_string = TT_PARSE_ERROR
                Exit Function
            End If

        ' U+D000..U+D7FF: bytes ED 80..9F 80..BF
        Case &HED
            If Not next_byte_in_range(&H80, &H9F, &H80, &HBF) Then
                scan_string = TT_PARSE_ERROR
                Exit Function
            End If

        ' U+10000..U+3FFFF F0 90..BF 80..BF 80..BF
        Case &HF0
            If Not next_byte_in_range(&H90, &HBF, &H80, &HBF, &H80, &HBF) Then
                scan_string = TT_PARSE_ERROR
                Exit Function
            End If

        ' U+40000..U+FFFFF F1..F3 80..BF 80..BF 80..BF
        Case &HF1 To &HF3
            If Not next_byte_in_range(&H80, &HBF, &H80, &HBF, &H80, &HBF) Then
                scan_string = TT_PARSE_ERROR
                Exit Function
            End If

        ' U+100000..U+10FFFF F4 80..8F 80..BF 80..BF
        Case &HF4
            If Not next_byte_in_range(&H80, &H8F, &H80, &HBF, &H80, &HBF) Then
                scan_string = TT_PARSE_ERROR
                Exit Function
            End If

        ' remaining bytes (80..C1 and F5..FF) are ill-formed
        Case Else
            lxErrorMessage = "invalid string: ill-formed UTF-8 byte"
            scan_string = TT_PARSE_ERROR
            Exit Function
        End Select
    Loop
End Function

Private Function IsDigit(ByVal c As Long) As Boolean
    IsDigit = (c >= 48 And c <= 57)
End Function

' scan_number: the library's state machine, then strtoull / strtoll / strtod
Private Function scan_number() As Long
    Dim number_type As Long, state As Long, c As Long, s As String, v As Variant, endIdx As Long

    ' reset token_buffer to store the number's bytes
    reset

    ' the type of the parsed number; initially set to unsigned; will be
    ' changed if minus sign, decimal point or exponent is read
    number_type = TT_VALUE_UNSIGNED

    ' state (init): we just found out we need to scan a number
    c = lxCurrent
    add c
    If c = 45 Then                              ' '-'
        state = 1
    ElseIf c = 48 Then                          ' '0'
        state = 2
    Else
        state = 3
    End If

    Do
        Select Case state
        Case 1  ' scan_number_minus: we just parsed a leading minus sign
            number_type = TT_VALUE_INTEGER
            c = get_()
            If c = 48 Then
                add c
                state = 2
            ElseIf c >= 49 And c <= 57 Then
                add c
                state = 3
            Else
                lxErrorMessage = "invalid number; expected digit after '-'"
                scan_number = TT_PARSE_ERROR
                Exit Function
            End If
        Case 2  ' scan_number_zero: we just parse a zero (maybe with a leading minus sign)
            c = get_()
            If c = 46 Then
                add 46                          ' decimal_point_char
                state = 4
            ElseIf c = 101 Or c = 69 Then
                add c
                state = 6
            Else
                Exit Do
            End If
        Case 3  ' scan_number_any1: we just parsed a number 0-9 (but maybe more)
            c = get_()
            If IsDigit(c) Then
                add c
            ElseIf c = 46 Then
                add 46
                state = 4
            ElseIf c = 101 Or c = 69 Then
                add c
                state = 6
            Else
                Exit Do
            End If
        Case 4  ' scan_number_decimal1: state: we just parsed a decimal point
            number_type = TT_VALUE_FLOAT
            c = get_()
            If IsDigit(c) Then
                add c
                state = 5
            Else
                lxErrorMessage = "invalid number; expected digit after '.'"
                scan_number = TT_PARSE_ERROR
                Exit Function
            End If
        Case 5  ' scan_number_decimal2: we just parsed at least one number after a decimal point
            c = get_()
            If IsDigit(c) Then
                add c
            ElseIf c = 101 Or c = 69 Then
                add c
                state = 6
            Else
                Exit Do
            End If
        Case 6  ' scan_number_exponent: we just parsed an exponent
            number_type = TT_VALUE_FLOAT
            c = get_()
            If c = 43 Or c = 45 Then
                add c
                state = 7
            ElseIf IsDigit(c) Then
                add c
                state = 8
            Else
                lxErrorMessage = "invalid number; expected '+', '-', or digit after exponent"
                scan_number = TT_PARSE_ERROR
                Exit Function
            End If
        Case 7  ' scan_number_sign: we just parsed an exponent sign
            c = get_()
            If IsDigit(c) Then
                add c
                state = 8
            Else
                lxErrorMessage = "invalid number; expected digit after exponent sign"
                scan_number = TT_PARSE_ERROR
                Exit Function
            End If
        Case 8  ' scan_number_any2: we just parsed a number after the exponent or exponent sign
            c = get_()
            If IsDigit(c) Then
                add c
            Else
                Exit Do
            End If
        End Select
    Loop

    ' scan_number_done: unget the character after the number (we only read it to know that we are done scanning a number)
    unget

    s = get_string()

    ' try to parse integers first and fall back to floats
    If number_type = TT_VALUE_UNSIGNED Then
        If Len(s) <= 15 Then
            lxNumber = Val(s)
            scan_number = TT_VALUE_UNSIGNED
            Exit Function
        ElseIf Len(s) <= 20 Then
            v = CDec(s)
            If v <= CDec("18446744073709551615") Then
                If v <= 9007199254740992# Then lxNumber = CDbl(v) Else lxNumber = v
                scan_number = TT_VALUE_UNSIGNED
                Exit Function
            End If
        End If
    ElseIf number_type = TT_VALUE_INTEGER Then
        If Len(s) <= 16 Then
            lxNumber = Val(s)
            scan_number = TT_VALUE_INTEGER
            Exit Function
        ElseIf Len(s) <= 21 Then
            v = CDec(s)
            If v >= CDec("-9223372036854775808") Then
                If v >= -9007199254740992# Then lxNumber = CDbl(v) Else lxNumber = v
                scan_number = TT_VALUE_INTEGER
                Exit Function
            End If
        End If
    End If

    ' this code is reached if we parse a floating-point number or if an
    ' integer conversion above failed
    lxNumber = ParseFloat(s)
    scan_number = TT_VALUE_FLOAT
End Function

' strtod of a scanned number; infinity when it overflows (the parser reports that)
Private Function ParseFloat(ByVal s As String) As Double
    Dim endIdx As Long
    On Error GoTo overflow
    ParseFloat = strtod(s, endIdx)
    Exit Function
overflow:
    ParseFloat = DoubleFromBits(&H7FF00000, 0)
End Function

Private Function scan_literal(ByVal literal_text As String, ByVal return_type As Long) As Long
    Dim i As Long
    For i = 2 To Len(literal_text)
        If get_() <> AscW(Mid$(literal_text, i, 1)) Then
            lxErrorMessage = "invalid literal"
            scan_literal = TT_PARSE_ERROR
            Exit Function
        End If
    Next
    scan_literal = return_type
End Function

Private Sub skip_whitespace()
    Do
        get_
    Loop While lxCurrent = 32 Or lxCurrent = 9 Or lxCurrent = 10 Or lxCurrent = 13
End Sub

' skip the UTF-8 byte order mark: true when there is none, or a whole one
Private Function skip_bom() As Boolean
    If get_() = &HEF Then
        If get_() = &HBB Then
            skip_bom = (get_() = &HBF)
        End If
        Exit Function
    End If

    ' we did not read a BOM: unget and continue
    unget
    skip_bom = True
End Function

Private Function scan() As Long
    ' initially, skip the BOM
    If lxCharsReadTotal = 0 Then
        If Not skip_bom() Then
            lxErrorMessage = "invalid BOM; must be 0xEF 0xBB 0xBF if given"
            scan = TT_PARSE_ERROR
            Exit Function
        End If
    End If

    ' read next character and ignore whitespace
    skip_whitespace

    Select Case lxCurrent
    ' structural characters
    Case 91: scan = TT_BEGIN_ARRAY
    Case 93: scan = TT_END_ARRAY
    Case 123: scan = TT_BEGIN_OBJECT
    Case 125: scan = TT_END_OBJECT
    Case 58: scan = TT_NAME_SEPARATOR
    Case 44: scan = TT_VALUE_SEPARATOR

    ' literals
    Case 116: scan = scan_literal("true", TT_LITERAL_TRUE)
    Case 102: scan = scan_literal("false", TT_LITERAL_FALSE)
    Case 110: scan = scan_literal("null", TT_LITERAL_NULL)

    ' string
    Case 34: scan = scan_string()

    ' number
    Case 45, 48 To 57: scan = scan_number()

    ' end of input (the null byte is needed when parsing from
    ' string literals)
    Case 0, EOF_: scan = TT_END_OF_INPUT

    ' error
    Case Else
        lxErrorMessage = "invalid literal"
        scan = TT_PARSE_ERROR
    End Select
End Function

'--------------------------------------------------------------------------
' parser, with json_sax_dom_parser building the values

Private Function get_token() As Long
    psLastToken = scan()
    get_token = psLastToken
End Function

Private Function exception_message(ByVal expected As Long, ByVal context As String) As String
    Dim error_msg As String
    error_msg = "syntax error "

    If Len(context) > 0 Then
        error_msg = error_msg & "while parsing " & context & " "
    End If

    error_msg = error_msg & "- "

    If psLastToken = TT_PARSE_ERROR Then
        error_msg = error_msg & lxErrorMessage & "; last read: '" & get_token_string() & "'"
    Else
        error_msg = error_msg & "unexpected " & token_type_name(psLastToken)
    End If

    If expected <> TT_UNINITIALIZED Then
        error_msg = error_msg & "; expected " & token_type_name(expected)
    End If

    exception_message = error_msg
End Function

' parse_error::create(101, position, what): raised
Private Sub parse_error(ByVal what_arg As String)
    json_throw "parse_error", 101, "parse error at line " & (lxLinesRead + 1) & ", column " & lxCharsReadCurrentLine & ": " & what_arg
End Sub

Private Function NewValue(ByVal t As value_t) As nlohmann_json
    Set NewValue = New nlohmann_json
    NewValue.init_value t, psOrdered
End Function

' json_sax_dom_parser::handle_value: the value goes to the root, the array or the key's slot
Private Function handle_value(ByVal v As nlohmann_json) As nlohmann_json
    Dim top As nlohmann_json
    If psRefCount = 0 Then
        Set psRoot = v
    Else
        Set top = psRefStack(psRefCount - 1)
        If top.is_array() Then
            top.append_item v
        Else
            top.set_slot psObjectElement, v
        End If
    End If
    Set handle_value = v
End Function

Private Sub PushRef(ByVal v As nlohmann_json)
    If psRefCount > UBound(psRefStack) Then ReDim Preserve psRefStack(0 To 2 * psRefCount - 1)
    Set psRefStack(psRefCount) = v
    psRefCount = psRefCount + 1
End Sub

Private Sub PopRef()
    psRefCount = psRefCount - 1
    Set psRefStack(psRefCount) = Nothing
End Sub

Private Sub sax_key()
    psObjectElement = psRefStack(psRefCount - 1).object_slot(get_string())
End Sub

Private Sub sax_parse_internal()
    Dim states() As Boolean, nStates As Long, skip_to_state_evaluation As Boolean, v As nlohmann_json

    ReDim states(0 To 15)
    Do
        If Not skip_to_state_evaluation Then
            ' invariant: get_token() was called before each iteration
            Select Case psLastToken
            Case TT_BEGIN_OBJECT
                PushRef handle_value(NewValue(value_t_object))

                ' closing } -> we are done
                If get_token() = TT_END_OBJECT Then
                    PopRef
                    GoTo state_evaluation
                End If

                ' parse key
                If psLastToken <> TT_VALUE_STRING Then
                    parse_error exception_message(TT_VALUE_STRING, "object key")
                End If
                sax_key

                ' parse separator (:)
                If get_token() <> TT_NAME_SEPARATOR Then
                    parse_error exception_message(TT_NAME_SEPARATOR, "object separator")
                End If

                ' remember we are now inside an object
                If nStates > UBound(states) Then ReDim Preserve states(0 To 2 * nStates - 1)
                states(nStates) = False
                nStates = nStates + 1

                ' parse values
                get_token
                GoTo next_iteration

            Case TT_BEGIN_ARRAY
                PushRef handle_value(NewValue(value_t_array))

                ' closing ] -> we are done
                If get_token() = TT_END_ARRAY Then
                    PopRef
                    GoTo state_evaluation
                End If

                ' remember we are now inside an array
                If nStates > UBound(states) Then ReDim Preserve states(0 To 2 * nStates - 1)
                states(nStates) = True
                nStates = nStates + 1

                ' parse values (no need to call get_token)
                GoTo next_iteration

            Case TT_VALUE_FLOAT
                If (DoubleHi(lxNumber) And &H7FF00000) = &H7FF00000 Then
                    json_throw "out_of_range", 406, "number overflow parsing '" & get_token_string() & "'"
                End If
                Set v = NewValue(value_t_number_float)
                v.f_num = lxNumber
                handle_value v

            Case TT_LITERAL_FALSE
                Set v = NewValue(value_t_boolean)
                v.f_bool = False
                handle_value v

            Case TT_LITERAL_NULL
                handle_value NewValue(value_t_null)

            Case TT_LITERAL_TRUE
                Set v = NewValue(value_t_boolean)
                v.f_bool = True
                handle_value v

            Case TT_VALUE_INTEGER
                Set v = NewValue(value_t_number_integer)
                v.f_num = lxNumber
                handle_value v

            Case TT_VALUE_STRING
                Set v = NewValue(value_t_string)
                v.f_str = get_string()
                handle_value v

            Case TT_VALUE_UNSIGNED
                Set v = NewValue(value_t_number_unsigned)
                v.f_num = lxNumber
                handle_value v

            Case TT_PARSE_ERROR
                ' using "uninitialized" to avoid "expected" message
                parse_error exception_message(TT_UNINITIALIZED, "value")

            Case TT_END_OF_INPUT
                If lxCharsReadTotal = 1 Then
                    parse_error "attempting to parse an empty input; check that your input string or stream contains the expected JSON"
                End If
                parse_error exception_message(TT_LITERAL_OR_VALUE, "value")

            Case Else ' the last token was unexpected
                parse_error exception_message(TT_LITERAL_OR_VALUE, "value")
            End Select
        Else
            skip_to_state_evaluation = False
        End If

state_evaluation:
        ' we reached this line after we successfully parsed a value
        If nStates = 0 Then
            ' empty stack: we reached the end of the hierarchy: done
            Exit Sub
        End If

        If states(nStates - 1) Then ' array
            ' comma -> next value
            If get_token() = TT_VALUE_SEPARATOR Then
                ' parse a new value
                get_token
                GoTo next_iteration
            End If

            ' closing ]
            If psLastToken = TT_END_ARRAY Then
                PopRef
                ' We are done with this array. Before we can parse a
                ' new value, we need to evaluate the new state first.
                ' By setting skip_to_state_evaluation to false, we
                ' are effectively jumping to the beginning of this if.
                nStates = nStates - 1
                skip_to_state_evaluation = True
                GoTo next_iteration
            End If

            parse_error exception_message(TT_END_ARRAY, "array")
        End If

        ' states.back() is false -> object

        ' comma -> next value
        If get_token() = TT_VALUE_SEPARATOR Then
            ' parse key
            If get_token() <> TT_VALUE_STRING Then
                parse_error exception_message(TT_VALUE_STRING, "object key")
            End If
            sax_key

            ' parse separator (:)
            If get_token() <> TT_NAME_SEPARATOR Then
                parse_error exception_message(TT_NAME_SEPARATOR, "object separator")
            End If

            ' parse values
            get_token
            GoTo next_iteration
        End If

        ' closing }
        If psLastToken = TT_END_OBJECT Then
            PopRef
            ' We are done with this object. Before we can parse a
            ' new value, we need to evaluate the new state first.
            ' By setting skip_to_state_evaluation to false, we
            ' are effectively jumping to the beginning of this if.
            nStates = nStates - 1
            skip_to_state_evaluation = True
            GoTo next_iteration
        End If

        parse_error exception_message(TT_END_OBJECT, "object")
next_iteration:
    Loop
End Sub

Private Function DoubleHi(ByVal d As Double) As Long
    CopyMemory DoubleHi, ByVal VarPtr(d) + 4, 4
End Function

' basic_json::parse(first, last) of the bytes data(first .. first + size - 1): strict, with exceptions
Public Function json_parse_bytes(data() As Byte, ByVal first As Long, ByVal size As Long, Optional ByVal ordered As Boolean = False) As nlohmann_json
    lxData = data
    lxPos = first
    lxEnd = first + size
    lxCurrent = EOF_
    lxNextUnget = False
    lxCharsReadTotal = 0
    lxCharsReadCurrentLine = 0
    lxLinesRead = 0
    ReDim lxTok(0 To 63)
    lxTokLen = 0
    If Len(lxBuf) = 0 Then lxBuf = Space$(256)
    lxBufLen = 0
    lxErrorMessage = ""
    psLastToken = TT_UNINITIALIZED
    ReDim psRefStack(0 To 15)
    psRefCount = 0
    Set psRoot = Nothing
    psObjectElement = 0
    psOrdered = ordered

    ' read first token
    get_token
    sax_parse_internal

    ' in strict mode, input must be completely read
    If get_token() <> TT_END_OF_INPUT Then
        parse_error exception_message(TT_END_OF_INPUT, "value")
    End If

    Set json_parse_bytes = psRoot
    Set psRoot = Nothing
    Erase lxData
End Function

' basic_json::parse(std::string): the String as UTF-8
Public Function json_parse(ByVal s As String, Optional ByVal ordered As Boolean = False) As nlohmann_json
    Dim b() As Byte, n As Long
    n = utf8_encode(s, b)
    If n = 0 Then ReDim b(0 To 0)
    Set json_parse = json_parse_bytes(b, 0, n, ordered)
End Function

'--------------------------------------------------------------------------
' Values (the C++ constructors and factories)

' json::object()
Public Function json_object(Optional ByVal ordered As Boolean = False) As nlohmann_json
    Set json_object = New nlohmann_json
    json_object.init_value value_t_object, ordered
End Function

' json::array()
Public Function json_array(Optional ByVal ordered As Boolean = False) As nlohmann_json
    Set json_array = New nlohmann_json
    json_array.init_value value_t_array, ordered
End Function

' json(v): Empty or Null is null; Boolean a boolean; Byte an unsigned
' number; Integer, Long or Decimal an integer; Single, Double or Currency a
' float; String a string; an nlohmann_json is copied
Public Function json_value(ByVal v As Variant, Optional ByVal ordered As Boolean = False) As nlohmann_json
    Dim j As nlohmann_json
    If IsObject(v) Then
        Set json_value = v.clone()
        Exit Function
    End If
    Set j = New nlohmann_json
    Select Case VarType(v)
    Case vbEmpty, vbNull
        j.init_value value_t_null, ordered
    Case vbBoolean
        j.init_value value_t_boolean, ordered
        j.f_bool = v
    Case vbByte
        j.init_value value_t_number_unsigned, ordered
        j.f_num = CDbl(v)
    Case vbInteger, vbLong
        j.init_value value_t_number_integer, ordered
        j.f_num = CDbl(v)
    Case vbDecimal
        j.init_value value_t_number_integer, ordered
        If Abs(v) <= 9007199254740992# Then j.f_num = CDbl(v) Else j.f_num = v
    Case vbSingle, vbDouble, vbCurrency
        j.init_value value_t_number_float, ordered
        j.f_num = CDbl(v)
    Case vbString
        j.init_value value_t_string, ordered
        j.f_str = v
    Case Else
        json_throw "type_error", 302, "type must be number, but is " & TypeName(v)
    End Select
    Set json_value = j
End Function

' json(unsigned value): a Long's uint32_t bits, a Double or a Decimal
Public Function json_value_unsigned(ByVal v As Variant, Optional ByVal ordered As Boolean = False) As nlohmann_json
    Dim j As nlohmann_json
    Set j = New nlohmann_json
    j.init_value value_t_number_unsigned, ordered
    Select Case VarType(v)
    Case vbLong, vbInteger
        j.f_num = U32(v)
    Case vbDecimal
        If v <= 9007199254740992# Then j.f_num = CDbl(v) Else j.f_num = v
    Case Else
        j.f_num = CDbl(v)
    End Select
    Set json_value_unsigned = j
End Function
