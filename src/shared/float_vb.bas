Attribute VB_Name = "float_vb"
' (VB-only) float arithmetic rounded as C rounds it: each float operation is
' done in Double and rounded to Single through memory (F32), so a chain of
' operations does not keep the x87's extra precision (the projects build
' with unrounded floating point allowed). A single +, -, *, / or sqrt done in
' Double and rounded to Single gives the correctly rounded float result.
' (Moved from WzGfx's glm_vb, which keeps glm's vector operations.)
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function clearfp_ Lib "msvcrt" Alias "_clearfp" () As Long

Private f32buf(0 To 0) As Single

' a Double rounded to float (C: (float)d), through memory. VB's conversion raises
' an overflow where C's gives an infinity (from an infinity, or a value that rounds
' past FLT_MAX), so those, and the NaNs, get their bits here.
Public Function F32(ByVal d As Double) As Single
    Dim w(0 To 1) As Long, hiAbs As Long, bits As Long
    CopyMemory w(0), d, 8
    hiAbs = w(1) And &H7FFFFFFF
    If hiAbs >= &H7FF00000 Then
        ' NaN or infinity: the sign, an all-ones exponent and the top 23 bits of the mantissa.
        ' The operation that made a NaN left the x87's sticky "invalid" flag, which VB's next
        ' float to integer conversion (anywhere) would report as an overflow: cleared.
        clearfp_
        bits = (w(1) And &H80000000) Or &H7F800000 Or ((w(1) And &HFFFFF) * 8) _
            Or ((w(0) And &H60000000) \ &H20000000) Or ((w(0) < 0) And 4)
        CopyMemory F32, bits, 4
    ElseIf hiAbs > &H47EFFFFF Or (hiAbs = &H47EFFFFF And Not ULT(w(0), &HF0000000)) Then
        ' |d| >= FLT_MAX + half an ulp (0x47EFFFFFF0000000): rounds to an infinity
        bits = (w(1) And &H80000000) Or &H7F800000
        CopyMemory F32, bits, 4
    Else
        f32buf(0) = d
        F32 = f32buf(0)
    End If
End Function

' float a + b, a - b, a * b, a / b
Public Function fadd(ByVal a As Single, ByVal b As Single) As Single
    fadd = F32(CDbl(a) + CDbl(b))
End Function

Public Function fsub(ByVal a As Single, ByVal b As Single) As Single
    fsub = F32(CDbl(a) - CDbl(b))
End Function

Public Function fmul(ByVal a As Single, ByVal b As Single) As Single
    fmul = F32(CDbl(a) * CDbl(b))
End Function

Public Function fdiv(ByVal a As Single, ByVal b As Single) As Single
    If IsNaNF(a) Or IsNaNF(b) Then
        fdiv = F32(CDbl(a) + CDbl(b))   ' (a NaN operand gives that NaN)
    ElseIf b = 0 Then
        fdiv = F32(DivByZero(CDbl(a)))
    Else
        fdiv = F32(CDbl(a) / CDbl(b))
    End If
End Function

' sqrtf
Public Function fsqrt(ByVal a As Single) As Single
    If IsNaNF(a) Then
        fsqrt = a
    ElseIf a < 0 Then
        fsqrt = F32(NaN_())
    Else
        fsqrt = F32(Sqr(CDbl(a)))
    End If
End Function

' (VB) IEEE division by zero (VB's / raises an error): +-inf, or NaN for 0 / 0
Public Function DivByZero(ByVal a As Double) As Double
    Static infBits(0 To 1) As Long, nanBits(0 To 1) As Long, ready As Boolean
    Dim d As Double
    If Not ready Then
        infBits(1) = &H7FF00000
        nanBits(1) = &HFFF80000
        ready = True
    End If
    If a = 0 Or a <> a Then
        CopyMemory d, nanBits(0), 8
    Else
        CopyMemory d, infBits(0), 8
        If a < 0 Then d = -d
    End If
    DivByZero = d
End Function

' (VB) whether f is a NaN. A comparison with a NaN is unordered: VB's compiled
' comparisons do not give C's answers for it (C: all false but !=), and they set
' the x87's sticky "invalid" flag (see F32), so the code that can see a NaN tests
' for one first and compares only the numbers.
Public Function IsNaNF(ByVal f As Single) As Boolean
    Dim bits As Long
    CopyMemory bits, f, 4
    IsNaNF = ((bits And &H7F800000) = &H7F800000) And ((bits And &H7FFFFF) <> 0)
End Function

' (VB) the default NaN (0xFFF8000000000000)
Public Function NaN_() As Double
    NaN_ = DivByZero(0)
End Function
