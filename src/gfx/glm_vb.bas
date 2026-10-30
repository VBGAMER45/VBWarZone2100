Attribute VB_Name = "glm_vb"
' (VB-only) glm's float vector operations, rounded as C++ rounds them, on
' src/shared/float_vb's float operations (F32, fadd, fmul, ...). glm's
' formulas (no SIMD): dot is (x*x + y*y) + z*z, normalize is
' v * (1 / sqrt(dot(v, v))), cross is glm's componentwise one.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Function vec3f(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Vector3f
    vec3f.x = x
    vec3f.y = y
    vec3f.z = z
End Function

Public Function v3sub(a As Vector3f, b As Vector3f) As Vector3f
    v3sub.x = fsub(a.x, b.x)
    v3sub.y = fsub(a.y, b.y)
    v3sub.z = fsub(a.z, b.z)
End Function

Public Function v3add(a As Vector3f, b As Vector3f) As Vector3f
    v3add.x = fadd(a.x, b.x)
    v3add.y = fadd(a.y, b.y)
    v3add.z = fadd(a.z, b.z)
End Function

' v * s
Public Function v3scale(a As Vector3f, ByVal s As Single) As Vector3f
    v3scale.x = fmul(a.x, s)
    v3scale.y = fmul(a.y, s)
    v3scale.z = fmul(a.z, s)
End Function

' glm::dot
Public Function v3dot(a As Vector3f, b As Vector3f) As Single
    v3dot = fadd(fadd(fmul(a.x, b.x), fmul(a.y, b.y)), fmul(a.z, b.z))
End Function

' glm::cross
Public Function v3cross(x As Vector3f, y As Vector3f) As Vector3f
    v3cross.x = fsub(fmul(x.y, y.z), fmul(y.y, x.z))
    v3cross.y = fsub(fmul(x.z, y.x), fmul(y.z, x.x))
    v3cross.z = fsub(fmul(x.x, y.y), fmul(y.x, x.y))
End Function

' glm::normalize: v * inversesqrt(dot(v, v)), inversesqrt being 1 / sqrt
Public Function v3normalize(v As Vector3f) As Vector3f
    v3normalize = v3scale(v, fdiv(1!, fsqrt(v3dot(v, v))))
End Function

' the bits of a float, as 8 hex digits (for the test dumps)
Public Function FloatHex(ByVal f As Single) As String
    Dim bits As Long
    CopyMemory bits, f, 4
    FloatHex = Right$("0000000" & Hex$(bits), 8)
End Function
