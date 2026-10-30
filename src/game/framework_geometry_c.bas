Attribute VB_Name = "framework_geometry_c"
' lib/framework/geometry.cpp, geometry.h: class Affine3F, the fixed point 3D
' affine transformation (RotX, RotY, RotZ, Trans, Scale, InvRot,
' translation and the two products; RotZXY and friends come when the
' display needs them). The class is a UDT (Affine3F) and its methods are
' procedures taking it first (af.RotX(x) is Affine3F_RotX af, x); the
' constructor is Affine3F_construct; a method returning *this for chaining
' is a Sub. int64_t intermediates are Doubles (exact while the products
' stay below 2^53: the rotations' are below 2^48), an int64_t >> 16 is
' Int(v / 65536#), C's / is Fix, and static_cast<int32_t> is FromU32.
' (VB) In src/game rather than src/shared: Vector3i and Rotation are each
' project's own types, and only the exe uses this so far.
Option Explicit

Public Const Affine3F_FRAC_BITS As Long = 12
Public Const Affine3F_INT_BITS As Long = 31 - Affine3F_FRAC_BITS
Public Const Affine3F_FP_MULTIPLIER As Long = 4096      ' 1 << FRAC_BITS

' Compact 4x4 matrix where the 4th row is assumed to be [0 0 0 1]
' since we're representing an affine transformation
Public Type Affine3F
    m(0 To 2, 0 To 3) As Long
End Type

' Affine3F::Affine3F()
Public Sub Affine3F_construct(af As Affine3F)
    af.m(0, 0) = Affine3F_FP_MULTIPLIER: af.m(0, 1) = 0: af.m(0, 2) = 0: af.m(0, 3) = 0
    af.m(1, 0) = 0: af.m(1, 1) = Affine3F_FP_MULTIPLIER: af.m(1, 2) = 0: af.m(1, 3) = 0
    af.m(2, 0) = 0: af.m(2, 1) = 0: af.m(2, 2) = Affine3F_FP_MULTIPLIER: af.m(2, 3) = 0
End Sub

' (VB) static_cast<int32_t>(v >> 16) of an int64_t
Private Function shr16(ByVal v As Double) As Long
    shr16 = FromU32(Int(v / 65536#))
End Function

' Affine3F &Affine3F::RotX(uint16_t x)
Public Sub Affine3F_RotX(af As Affine3F, ByVal x As Long)
    Dim tmp As Long, cra As Double, sra As Double
    x = x And &HFFFF&
    If x <> 0 Then
        cra = iCos(x)
        sra = iSin(x)

        tmp = shr16(cra * af.m(0, 1) + sra * af.m(0, 2))
        af.m(0, 2) = shr16(cra * af.m(0, 2) - sra * af.m(0, 1))
        af.m(0, 1) = tmp

        tmp = shr16(cra * af.m(1, 1) + sra * af.m(1, 2))
        af.m(1, 2) = shr16(cra * af.m(1, 2) - sra * af.m(1, 1))
        af.m(1, 1) = tmp

        tmp = shr16(cra * af.m(2, 1) + sra * af.m(2, 2))
        af.m(2, 2) = shr16(cra * af.m(2, 2) - sra * af.m(2, 1))
        af.m(2, 1) = tmp
    End If
End Sub

' Affine3F &Affine3F::RotY(uint16_t y)
Public Sub Affine3F_RotY(af As Affine3F, ByVal y As Long)
    Dim tmp As Long, cra As Double, sra As Double
    y = y And &HFFFF&
    If y <> 0 Then
        cra = iCos(y)
        sra = iSin(y)

        tmp = shr16(cra * af.m(0, 0) - sra * af.m(0, 2))
        af.m(0, 2) = shr16(sra * af.m(0, 0) + cra * af.m(0, 2))
        af.m(0, 0) = tmp

        tmp = shr16(cra * af.m(1, 0) - sra * af.m(1, 2))
        af.m(1, 2) = shr16(sra * af.m(1, 0) + cra * af.m(1, 2))
        af.m(1, 0) = tmp

        tmp = shr16(cra * af.m(2, 0) - sra * af.m(2, 2))
        af.m(2, 2) = shr16(sra * af.m(2, 0) + cra * af.m(2, 2))
        af.m(2, 0) = tmp
    End If
End Sub

' Affine3F &Affine3F::RotZ(uint16_t z)
Public Sub Affine3F_RotZ(af As Affine3F, ByVal z As Long)
    Dim tmp As Long, cra As Double, sra As Double
    z = z And &HFFFF&
    If z <> 0 Then
        cra = iCos(z)
        sra = iSin(z)

        tmp = shr16(cra * af.m(0, 0) + sra * af.m(0, 1))
        af.m(0, 1) = shr16(cra * af.m(0, 1) - sra * af.m(0, 0))
        af.m(0, 0) = tmp

        tmp = shr16(cra * af.m(1, 0) + sra * af.m(1, 1))
        af.m(1, 1) = shr16(cra * af.m(1, 1) - sra * af.m(1, 0))
        af.m(1, 0) = tmp

        tmp = shr16(cra * af.m(2, 0) + sra * af.m(2, 1))
        af.m(2, 1) = shr16(cra * af.m(2, 1) - sra * af.m(2, 0))
        af.m(2, 0) = tmp
    End If
End Sub

' Affine3F &Affine3F::Trans(int32_t x, int32_t y, int32_t z): int32_t arithmetic, wrapping
Public Sub Affine3F_Trans(af As Affine3F, ByVal x As Long, ByVal y As Long, ByVal z As Long)
    af.m(0, 3) = FromU32(CDbl(af.m(0, 3)) + MUL32(x, af.m(0, 0)) + MUL32(y, af.m(0, 1)) + MUL32(z, af.m(0, 2)))
    af.m(1, 3) = FromU32(CDbl(af.m(1, 3)) + MUL32(x, af.m(1, 0)) + MUL32(y, af.m(1, 1)) + MUL32(z, af.m(1, 2)))
    af.m(2, 3) = FromU32(CDbl(af.m(2, 3)) + MUL32(x, af.m(2, 0)) + MUL32(y, af.m(2, 1)) + MUL32(z, af.m(2, 2)))
End Sub

' Affine3F &Affine3F::Trans(Vector3i tr)
Public Sub Affine3F_Trans_vec(af As Affine3F, tr As Vector3i)
    Affine3F_Trans af, tr.x, tr.y, tr.z
End Sub

' Affine3F &Affine3F::Scale(int32_t s): s is a scaled integer (scaled by FP_MULTIPLIER)
Public Sub Affine3F_Scale(af As Affine3F, ByVal s As Long)
    Dim i As Long, j As Long
    For i = 0 To 2
        For j = 0 To 2
            af.m(i, j) = FromU32(Fix(CDbl(af.m(i, j)) * s / Affine3F_FP_MULTIPLIER))
        Next j
    Next i
End Sub

' Affine3F &Affine3F::InvRot(): the swaps as written (the second, fourth, sixth and eighth undo or keep)
Public Sub Affine3F_InvRot(af As Affine3F)
    swapL af.m(0, 1), af.m(1, 0)
    swapL af.m(0, 2), af.m(2, 0)
    swapL af.m(1, 0), af.m(0, 1)
    swapL af.m(1, 2), af.m(2, 1)
    swapL af.m(2, 0), af.m(0, 2)
    swapL af.m(2, 1), af.m(1, 2)
End Sub

Private Sub swapL(a As Long, b As Long)
    Dim t As Long
    t = a
    a = b
    b = t
End Sub

' Vector3i Affine3F::translation() const
Public Function Affine3F_translation(af As Affine3F) As Vector3i
    Affine3F_translation.x = af.m(0, 3) \ Affine3F_FP_MULTIPLIER
    Affine3F_translation.y = af.m(1, 3) \ Affine3F_FP_MULTIPLIER
    Affine3F_translation.z = af.m(2, 3) \ Affine3F_FP_MULTIPLIER
End Function

' Vector3i Affine3F::operator*(const Vector3i v) const
Public Function Affine3F_mul_vec(af As Affine3F, v As Vector3i) As Vector3i
    Dim s As Vector3i
    s.x = FromU32(Fix((CDbl(v.x) * af.m(0, 0) + CDbl(v.y) * af.m(0, 1) + CDbl(v.z) * af.m(0, 2) + af.m(0, 3)) / Affine3F_FP_MULTIPLIER))
    s.y = FromU32(Fix((CDbl(v.x) * af.m(1, 0) + CDbl(v.y) * af.m(1, 1) + CDbl(v.z) * af.m(1, 2) + af.m(1, 3)) / Affine3F_FP_MULTIPLIER))
    s.z = FromU32(Fix((CDbl(v.x) * af.m(2, 0) + CDbl(v.y) * af.m(2, 1) + CDbl(v.z) * af.m(2, 2) + af.m(2, 3)) / Affine3F_FP_MULTIPLIER))
    Affine3F_mul_vec = s
End Function

' Affine3F Affine3F::operator*(const Affine3F &rhs) const
Public Function Affine3F_mul(af As Affine3F, rhs As Affine3F) As Affine3F
    Dim tr As Affine3F, i As Long, j As Long
    Affine3F_construct tr
    For i = 0 To 2
        For j = 0 To 2
            tr.m(i, j) = FromU32(Fix((CDbl(af.m(i, 0)) * rhs.m(0, j) + CDbl(af.m(i, 1)) * rhs.m(1, j) + CDbl(af.m(i, 2)) * rhs.m(2, j)) / Affine3F_FP_MULTIPLIER))
        Next j
        tr.m(i, 3) = FromU32(Fix((CDbl(af.m(i, 0)) * rhs.m(0, 3) + CDbl(af.m(i, 1)) * rhs.m(1, 3) + CDbl(af.m(i, 2)) * rhs.m(2, 3)) / Affine3F_FP_MULTIPLIER) + af.m(i, 3))
    Next i
    Affine3F_mul = tr
End Function

' Vector3i Affine3F::InvRot(const Vector3i v) const
Public Function Affine3F_InvRot_vec(af As Affine3F, v As Vector3i) As Vector3i
    Dim s As Vector3i
    s.x = FromU32(Fix((CDbl(v.x) * af.m(0, 0) + CDbl(v.y) * af.m(1, 0) + CDbl(v.z) * af.m(2, 0)) / Affine3F_FP_MULTIPLIER))
    s.y = FromU32(Fix((CDbl(v.x) * af.m(0, 1) + CDbl(v.y) * af.m(1, 1) + CDbl(v.z) * af.m(2, 1)) / Affine3F_FP_MULTIPLIER))
    s.z = FromU32(Fix((CDbl(v.x) * af.m(0, 2) + CDbl(v.y) * af.m(1, 2) + CDbl(v.z) * af.m(2, 2)) / Affine3F_FP_MULTIPLIER))
    Affine3F_InvRot_vec = s
End Function
