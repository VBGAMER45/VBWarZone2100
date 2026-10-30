Attribute VB_Name = "glm_mat_vb"
' (VB) glm's functions on glm_h's types, as the renderer (lib/ivis_opengl) and
' the display code (display3d, terrain, piedraw, effects, warcam,
' shadowcascades) use them: glm::mat4 operator* (glm_mat4_mul, and
' glm_mat4_mul_vec4 for mat4 * vec4), glm::translate, rotate, scale (gtx/
' transform.hpp: from the identity; the _m forms are ext/matrix_transform's
' with a matrix first), glm::ortho, frustum, perspective (right-handed, depth
' -1..1: Warzone defines no GLM_FORCE_* clip control), lookAt, inverse,
' transpose, interpolate (gtx/matrix_interpolation), and the vector
' operations they need. Written with glm 0.9.9's formulas (3rdparty/glm) in
' plain Single arithmetic: rendering need not match C++ bit for bit, so no
' float_vb calls, except that a division whose divisor can be 0 goes through
' fdiv (C's infinity where VB's / raises). glm_h's mat4 is column-major:
' glm's m[c][r] is m.m(c * 4 + r). A mat4 passed between WzGfx and the exe
' travels as its address (VarPtr): mat4_fromPtr / mat4_toPtr copy it.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

'--------------------------------------------------------------------------
' the address boundary

' the mat4 at address p (VarPtr of a mat4 or of its m(0))
Public Function mat4_fromPtr(ByVal p As Long) As mat4
    CopyMemory mat4_fromPtr.m(0), ByVal p, 64
End Function

' copies m to address p
Public Sub mat4_toPtr(m As mat4, ByVal p As Long)
    CopyMemory ByVal p, m.m(0), 64
End Sub

'--------------------------------------------------------------------------
' constructors

Public Function glm_vec2(ByVal x As Single, ByVal y As Single) As Vector2f
    glm_vec2.x = x
    glm_vec2.y = y
End Function

Public Function glm_ivec2(ByVal x As Long, ByVal y As Long) As Vector2i
    glm_ivec2.x = x
    glm_ivec2.y = y
End Function

Public Function glm_ivec3(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Vector3i
    glm_ivec3.x = x
    glm_ivec3.y = y
    glm_ivec3.z = z
End Function

Public Function glm_vec3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Vector3f
    glm_vec3.x = x
    glm_vec3.y = y
    glm_vec3.z = z
End Function

Public Function glm_vec4(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal w As Single) As Vector4f
    glm_vec4.x = x
    glm_vec4.y = y
    glm_vec4.z = z
    glm_vec4.w = w
End Function

' glm::vec4(v3, w)
Public Function glm_vec4_v3(v As Vector3f, ByVal w As Single) As Vector4f
    glm_vec4_v3.x = v.x
    glm_vec4_v3.y = v.y
    glm_vec4_v3.z = v.z
    glm_vec4_v3.w = w
End Function

Public Function glm_ivec4(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal w As Long) As Vector4i
    glm_ivec4.x = x
    glm_ivec4.y = y
    glm_ivec4.z = z
    glm_ivec4.w = w
End Function

' glm::mat4(s): s on the diagonal (glm::mat4(1.f) is the identity, glm::mat4(0.f) all zeros)
Public Function glm_mat4(ByVal s As Single) As mat4
    glm_mat4.m(0) = s
    glm_mat4.m(5) = s
    glm_mat4.m(10) = s
    glm_mat4.m(15) = s
End Function

' glm::mat4(c0, c1, c2, c3): from its four columns
Public Function glm_mat4_cols(c0 As Vector4f, c1 As Vector4f, c2 As Vector4f, c3 As Vector4f) As mat4
    Dim r As mat4
    r.m(0) = c0.x: r.m(1) = c0.y: r.m(2) = c0.z: r.m(3) = c0.w
    r.m(4) = c1.x: r.m(5) = c1.y: r.m(6) = c1.z: r.m(7) = c1.w
    r.m(8) = c2.x: r.m(9) = c2.y: r.m(10) = c2.z: r.m(11) = c2.w
    r.m(12) = c3.x: r.m(13) = c3.y: r.m(14) = c3.z: r.m(15) = c3.w
    glm_mat4_cols = r
End Function

' m[c] (a column)
Public Function glm_mat4_col(m As mat4, ByVal c As Long) As Vector4f
    glm_mat4_col.x = m.m(c * 4)
    glm_mat4_col.y = m.m(c * 4 + 1)
    glm_mat4_col.z = m.m(c * 4 + 2)
    glm_mat4_col.w = m.m(c * 4 + 3)
End Function

'--------------------------------------------------------------------------
' vectors

Public Function glm_vec3_add(a As Vector3f, b As Vector3f) As Vector3f
    glm_vec3_add.x = a.x + b.x
    glm_vec3_add.y = a.y + b.y
    glm_vec3_add.z = a.z + b.z
End Function

Public Function glm_vec3_sub(a As Vector3f, b As Vector3f) As Vector3f
    glm_vec3_sub.x = a.x - b.x
    glm_vec3_sub.y = a.y - b.y
    glm_vec3_sub.z = a.z - b.z
End Function

' v * s
Public Function glm_vec3_mul(v As Vector3f, ByVal s As Single) As Vector3f
    glm_vec3_mul.x = v.x * s
    glm_vec3_mul.y = v.y * s
    glm_vec3_mul.z = v.z * s
End Function

Public Function glm_vec4_add(a As Vector4f, b As Vector4f) As Vector4f
    glm_vec4_add.x = a.x + b.x
    glm_vec4_add.y = a.y + b.y
    glm_vec4_add.z = a.z + b.z
    glm_vec4_add.w = a.w + b.w
End Function

Public Function glm_vec4_sub(a As Vector4f, b As Vector4f) As Vector4f
    glm_vec4_sub.x = a.x - b.x
    glm_vec4_sub.y = a.y - b.y
    glm_vec4_sub.z = a.z - b.z
    glm_vec4_sub.w = a.w - b.w
End Function

' v * s
Public Function glm_vec4_mul(v As Vector4f, ByVal s As Single) As Vector4f
    glm_vec4_mul.x = v.x * s
    glm_vec4_mul.y = v.y * s
    glm_vec4_mul.z = v.z * s
    glm_vec4_mul.w = v.w * s
End Function

' v / s (componentwise: C's infinities where s is 0)
Public Function glm_vec4_div(v As Vector4f, ByVal s As Single) As Vector4f
    glm_vec4_div.x = fdiv(v.x, s)
    glm_vec4_div.y = fdiv(v.y, s)
    glm_vec4_div.z = fdiv(v.z, s)
    glm_vec4_div.w = fdiv(v.w, s)
End Function

' glm::dot (vec3): (x*x + y*y) + z*z
Public Function glm_dot3(a As Vector3f, b As Vector3f) As Single
    glm_dot3 = (a.x * b.x + a.y * b.y) + a.z * b.z
End Function

' glm::dot (vec4): (x*x + y*y) + (z*z + w*w)
Public Function glm_dot4(a As Vector4f, b As Vector4f) As Single
    glm_dot4 = (a.x * b.x + a.y * b.y) + (a.z * b.z + a.w * b.w)
End Function

' glm::cross
Public Function glm_cross(x As Vector3f, y As Vector3f) As Vector3f
    glm_cross.x = x.y * y.z - y.y * x.z
    glm_cross.y = x.z * y.x - y.z * x.x
    glm_cross.z = x.x * y.y - y.x * x.y
End Function

' glm::length (vec3): sqrt(dot(v, v))
Public Function glm_length3(v As Vector3f) As Single
    glm_length3 = Sqr(glm_dot3(v, v))
End Function

' glm::normalize (vec3): v * inversesqrt(dot(v, v)); a zero vector gives C's NaNs
' (built from their bits: 0 * inf on the x87 would leave the sticky invalid flag)
Public Function glm_normalize3(v As Vector3f) As Vector3f
    Dim d As Single, nan As Single
    d = glm_dot3(v, v)
    If d = 0! Then
        nan = F32(NaN_())
        glm_normalize3.x = nan
        glm_normalize3.y = nan
        glm_normalize3.z = nan
        Exit Function
    End If
    glm_normalize3 = glm_vec3_mul(v, 1! / Sqr(d))
End Function

' glm::normalize (vec4)
Public Function glm_normalize4(v As Vector4f) As Vector4f
    Dim d As Single, nan As Single
    d = glm_dot4(v, v)
    If d = 0! Then
        nan = F32(NaN_())
        glm_normalize4.x = nan
        glm_normalize4.y = nan
        glm_normalize4.z = nan
        glm_normalize4.w = nan
        Exit Function
    End If
    glm_normalize4 = glm_vec4_mul(v, 1! / Sqr(d))
End Function

'--------------------------------------------------------------------------
' matrices

' m1 * m2 (glm: Result[i] = m1[0] * m2[i][0] + m1[1] * m2[i][1] + m1[2] * m2[i][2] + m1[3] * m2[i][3])
Public Function glm_mat4_mul(a As mat4, b As mat4) As mat4
    Dim i As Long, c As Long, b0 As Single, b1 As Single, b2 As Single, b3 As Single
    For i = 0 To 3
        c = i * 4
        b0 = b.m(c): b1 = b.m(c + 1): b2 = b.m(c + 2): b3 = b.m(c + 3)
        glm_mat4_mul.m(c) = a.m(0) * b0 + a.m(4) * b1 + a.m(8) * b2 + a.m(12) * b3
        glm_mat4_mul.m(c + 1) = a.m(1) * b0 + a.m(5) * b1 + a.m(9) * b2 + a.m(13) * b3
        glm_mat4_mul.m(c + 2) = a.m(2) * b0 + a.m(6) * b1 + a.m(10) * b2 + a.m(14) * b3
        glm_mat4_mul.m(c + 3) = a.m(3) * b0 + a.m(7) * b1 + a.m(11) * b2 + a.m(15) * b3
    Next
End Function

' m * v (glm: (m[0] * v.x + m[1] * v.y) + (m[2] * v.z + m[3] * v.w))
Public Function glm_mat4_mul_vec4(m As mat4, v As Vector4f) As Vector4f
    glm_mat4_mul_vec4.x = (m.m(0) * v.x + m.m(4) * v.y) + (m.m(8) * v.z + m.m(12) * v.w)
    glm_mat4_mul_vec4.y = (m.m(1) * v.x + m.m(5) * v.y) + (m.m(9) * v.z + m.m(13) * v.w)
    glm_mat4_mul_vec4.z = (m.m(2) * v.x + m.m(6) * v.y) + (m.m(10) * v.z + m.m(14) * v.w)
    glm_mat4_mul_vec4.w = (m.m(3) * v.x + m.m(7) * v.y) + (m.m(11) * v.z + m.m(15) * v.w)
End Function

' m * s
Public Function glm_mat4_mul_scalar(m As mat4, ByVal s As Single) As mat4
    Dim i As Long
    For i = 0 To 15
        glm_mat4_mul_scalar.m(i) = m.m(i) * s
    Next
End Function

' glm::translate(m, v) (ext/matrix_transform): Result[3] = m[0] * v[0] + m[1] * v[1] + m[2] * v[2] + m[3]
Public Function glm_translate_m(m As mat4, v As Vector3f) As mat4
    Dim r As Long
    glm_translate_m = m
    For r = 0 To 3
        glm_translate_m.m(12 + r) = m.m(r) * v.x + m.m(4 + r) * v.y + m.m(8 + r) * v.z + m.m(12 + r)
    Next
End Function

' glm::translate(v) (gtx/transform): translate(mat4(1), v)
Public Function glm_translate(v As Vector3f) As mat4
    glm_translate = glm_mat4(1!)
    glm_translate.m(12) = v.x
    glm_translate.m(13) = v.y
    glm_translate.m(14) = v.z
End Function

' glm::rotate(m, angle, v) (ext/matrix_transform; angle in radians)
Public Function glm_rotate_m(m As mat4, ByVal angle As Single, v As Vector3f) As mat4
    Dim a As Single, c As Single, s As Single, axis As Vector3f, temp As Vector3f
    Dim r00 As Single, r01 As Single, r02 As Single
    Dim r10 As Single, r11 As Single, r12 As Single
    Dim r20 As Single, r21 As Single, r22 As Single
    Dim r As Long
    a = angle
    c = Cos(a)
    s = Sin(a)

    axis = glm_normalize3(v)
    temp = glm_vec3_mul(axis, 1! - c)

    r00 = c + temp.x * axis.x
    r01 = temp.x * axis.y + s * axis.z
    r02 = temp.x * axis.z - s * axis.y

    r10 = temp.y * axis.x - s * axis.z
    r11 = c + temp.y * axis.y
    r12 = temp.y * axis.z + s * axis.x

    r20 = temp.z * axis.x + s * axis.y
    r21 = temp.z * axis.y - s * axis.x
    r22 = c + temp.z * axis.z

    For r = 0 To 3
        glm_rotate_m.m(r) = m.m(r) * r00 + m.m(4 + r) * r01 + m.m(8 + r) * r02
        glm_rotate_m.m(4 + r) = m.m(r) * r10 + m.m(4 + r) * r11 + m.m(8 + r) * r12
        glm_rotate_m.m(8 + r) = m.m(r) * r20 + m.m(4 + r) * r21 + m.m(8 + r) * r22
        glm_rotate_m.m(12 + r) = m.m(12 + r)
    Next
End Function

' glm::rotate(angle, v) (gtx/transform): rotate(mat4(1), angle, v)
Public Function glm_rotate(ByVal angle As Single, v As Vector3f) As mat4
    Dim id As mat4
    id = glm_mat4(1!)
    glm_rotate = glm_rotate_m(id, angle, v)
End Function

' glm::scale(m, v) (ext/matrix_transform)
Public Function glm_scale_m(m As mat4, v As Vector3f) As mat4
    Dim r As Long
    For r = 0 To 3
        glm_scale_m.m(r) = m.m(r) * v.x
        glm_scale_m.m(4 + r) = m.m(4 + r) * v.y
        glm_scale_m.m(8 + r) = m.m(8 + r) * v.z
        glm_scale_m.m(12 + r) = m.m(12 + r)
    Next
End Function

' glm::scale(v) (gtx/transform): scale(mat4(1), v)
Public Function glm_scale(v As Vector3f) As mat4
    glm_scale.m(0) = v.x
    glm_scale.m(5) = v.y
    glm_scale.m(10) = v.z
    glm_scale.m(15) = 1!
End Function

' glm::ortho(left, right, bottom, top) (2D: depth -1)
Public Function glm_ortho(ByVal left_ As Single, ByVal right_ As Single, ByVal bottom As Single, ByVal top As Single) As mat4
    glm_ortho = glm_mat4(1!)
    glm_ortho.m(0) = fdiv(2!, right_ - left_)
    glm_ortho.m(5) = fdiv(2!, top - bottom)
    glm_ortho.m(10) = -1!
    glm_ortho.m(12) = -fdiv(right_ + left_, right_ - left_)
    glm_ortho.m(13) = -fdiv(top + bottom, top - bottom)
End Function

' glm::ortho(left, right, bottom, top, zNear, zFar) (orthoRH_NO)
Public Function glm_ortho6(ByVal left_ As Single, ByVal right_ As Single, ByVal bottom As Single, ByVal top As Single, ByVal zNear As Single, ByVal zFar As Single) As mat4
    glm_ortho6 = glm_mat4(1!)
    glm_ortho6.m(0) = fdiv(2!, right_ - left_)
    glm_ortho6.m(5) = fdiv(2!, top - bottom)
    glm_ortho6.m(10) = -fdiv(2!, zFar - zNear)
    glm_ortho6.m(12) = -fdiv(right_ + left_, right_ - left_)
    glm_ortho6.m(13) = -fdiv(top + bottom, top - bottom)
    glm_ortho6.m(14) = -fdiv(zFar + zNear, zFar - zNear)
End Function

' glm::frustum (frustumRH_NO)
Public Function glm_frustum(ByVal left_ As Single, ByVal right_ As Single, ByVal bottom As Single, ByVal top As Single, ByVal nearVal As Single, ByVal farVal As Single) As mat4
    glm_frustum.m(0) = fdiv(2! * nearVal, right_ - left_)
    glm_frustum.m(5) = fdiv(2! * nearVal, top - bottom)
    glm_frustum.m(8) = fdiv(right_ + left_, right_ - left_)
    glm_frustum.m(9) = fdiv(top + bottom, top - bottom)
    glm_frustum.m(10) = -fdiv(farVal + nearVal, farVal - nearVal)
    glm_frustum.m(11) = -1!
    glm_frustum.m(14) = -fdiv(2! * farVal * nearVal, farVal - nearVal)
End Function

' glm::perspective (perspectiveRH_NO; fovy in radians)
Public Function glm_perspective(ByVal fovy As Single, ByVal aspect As Single, ByVal zNear As Single, ByVal zFar As Single) As mat4
    Dim tanHalfFovy As Single
    tanHalfFovy = Tan(fovy / 2!)
    glm_perspective.m(0) = fdiv(1!, aspect * tanHalfFovy)
    glm_perspective.m(5) = fdiv(1!, tanHalfFovy)
    glm_perspective.m(10) = -fdiv(zFar + zNear, zFar - zNear)
    glm_perspective.m(11) = -1!
    glm_perspective.m(14) = -fdiv(2! * zFar * zNear, zFar - zNear)
End Function

' glm::lookAt (lookAtRH)
Public Function glm_lookAt(eye As Vector3f, center As Vector3f, up As Vector3f) As mat4
    Dim f As Vector3f, s As Vector3f, u As Vector3f
    f = glm_normalize3(glm_vec3_sub(center, eye))
    s = glm_normalize3(glm_cross(f, up))
    u = glm_cross(s, f)

    glm_lookAt = glm_mat4(1!)
    glm_lookAt.m(0) = s.x
    glm_lookAt.m(4) = s.y
    glm_lookAt.m(8) = s.z
    glm_lookAt.m(1) = u.x
    glm_lookAt.m(5) = u.y
    glm_lookAt.m(9) = u.z
    glm_lookAt.m(2) = -f.x
    glm_lookAt.m(6) = -f.y
    glm_lookAt.m(10) = -f.z
    glm_lookAt.m(12) = -glm_dot3(s, eye)
    glm_lookAt.m(13) = -glm_dot3(u, eye)
    glm_lookAt.m(14) = glm_dot3(f, eye)
End Function

' glm::transpose
Public Function glm_transpose(m As mat4) As mat4
    Dim c As Long, r As Long
    For c = 0 To 3
        For r = 0 To 3
            glm_transpose.m(c * 4 + r) = m.m(r * 4 + c)
        Next
    Next
End Function

' glm::inverse (detail::compute_inverse<4, 4>)
Public Function glm_inverse(m As mat4) As mat4
    Dim Coef00 As Single, Coef02 As Single, Coef03 As Single
    Dim Coef04 As Single, Coef06 As Single, Coef07 As Single
    Dim Coef08 As Single, Coef10 As Single, Coef11 As Single
    Dim Coef12 As Single, Coef14 As Single, Coef15 As Single
    Dim Coef16 As Single, Coef18 As Single, Coef19 As Single
    Dim Coef20 As Single, Coef22 As Single, Coef23 As Single
    Dim Fac0 As Vector4f, Fac1 As Vector4f, Fac2 As Vector4f, Fac3 As Vector4f, Fac4 As Vector4f, Fac5 As Vector4f
    Dim Vec0 As Vector4f, Vec1 As Vector4f, Vec2 As Vector4f, Vec3 As Vector4f
    Dim Inv0 As Vector4f, Inv1 As Vector4f, Inv2 As Vector4f, Inv3 As Vector4f
    Dim Inverse As mat4, Dot1 As Single, OneOverDeterminant As Single

    ' (m[c][r] is m.m(c * 4 + r))
    Coef00 = m.m(10) * m.m(15) - m.m(14) * m.m(11)
    Coef02 = m.m(6) * m.m(15) - m.m(14) * m.m(7)
    Coef03 = m.m(6) * m.m(11) - m.m(10) * m.m(7)

    Coef04 = m.m(9) * m.m(15) - m.m(13) * m.m(11)
    Coef06 = m.m(5) * m.m(15) - m.m(13) * m.m(7)
    Coef07 = m.m(5) * m.m(11) - m.m(9) * m.m(7)

    Coef08 = m.m(9) * m.m(14) - m.m(13) * m.m(10)
    Coef10 = m.m(5) * m.m(14) - m.m(13) * m.m(6)
    Coef11 = m.m(5) * m.m(10) - m.m(9) * m.m(6)

    Coef12 = m.m(8) * m.m(15) - m.m(12) * m.m(11)
    Coef14 = m.m(4) * m.m(15) - m.m(12) * m.m(7)
    Coef15 = m.m(4) * m.m(11) - m.m(8) * m.m(7)

    Coef16 = m.m(8) * m.m(14) - m.m(12) * m.m(10)
    Coef18 = m.m(4) * m.m(14) - m.m(12) * m.m(6)
    Coef19 = m.m(4) * m.m(10) - m.m(8) * m.m(6)

    Coef20 = m.m(8) * m.m(13) - m.m(12) * m.m(9)
    Coef22 = m.m(4) * m.m(13) - m.m(12) * m.m(5)
    Coef23 = m.m(4) * m.m(9) - m.m(8) * m.m(5)

    Fac0 = glm_vec4(Coef00, Coef00, Coef02, Coef03)
    Fac1 = glm_vec4(Coef04, Coef04, Coef06, Coef07)
    Fac2 = glm_vec4(Coef08, Coef08, Coef10, Coef11)
    Fac3 = glm_vec4(Coef12, Coef12, Coef14, Coef15)
    Fac4 = glm_vec4(Coef16, Coef16, Coef18, Coef19)
    Fac5 = glm_vec4(Coef20, Coef20, Coef22, Coef23)

    Vec0 = glm_vec4(m.m(4), m.m(0), m.m(0), m.m(0))
    Vec1 = glm_vec4(m.m(5), m.m(1), m.m(1), m.m(1))
    Vec2 = glm_vec4(m.m(6), m.m(2), m.m(2), m.m(2))
    Vec3 = glm_vec4(m.m(7), m.m(3), m.m(3), m.m(3))

    ' Inv0(Vec1 * Fac0 - Vec2 * Fac1 + Vec3 * Fac2) ...
    Inv0 = InverseTerm(Vec1, Fac0, Vec2, Fac1, Vec3, Fac2)
    Inv1 = InverseTerm(Vec0, Fac0, Vec2, Fac3, Vec3, Fac4)
    Inv2 = InverseTerm(Vec0, Fac1, Vec1, Fac3, Vec3, Fac5)
    Inv3 = InverseTerm(Vec0, Fac2, Vec1, Fac4, Vec2, Fac5)

    ' SignA(+1, -1, +1, -1), SignB(-1, +1, -1, +1): Inverse(Inv0 * SignA, Inv1 * SignB, Inv2 * SignA, Inv3 * SignB)
    Inverse.m(0) = Inv0.x: Inverse.m(1) = -Inv0.y: Inverse.m(2) = Inv0.z: Inverse.m(3) = -Inv0.w
    Inverse.m(4) = -Inv1.x: Inverse.m(5) = Inv1.y: Inverse.m(6) = -Inv1.z: Inverse.m(7) = Inv1.w
    Inverse.m(8) = Inv2.x: Inverse.m(9) = -Inv2.y: Inverse.m(10) = Inv2.z: Inverse.m(11) = -Inv2.w
    Inverse.m(12) = -Inv3.x: Inverse.m(13) = Inv3.y: Inverse.m(14) = -Inv3.z: Inverse.m(15) = Inv3.w

    ' Row0(Inverse[0][0], Inverse[1][0], Inverse[2][0], Inverse[3][0]); Dot0(m[0] * Row0)
    Dot1 = (m.m(0) * Inverse.m(0) + m.m(1) * Inverse.m(4)) + (m.m(2) * Inverse.m(8) + m.m(3) * Inverse.m(12))

    OneOverDeterminant = fdiv(1!, Dot1)

    glm_inverse = glm_mat4_mul_scalar(Inverse, OneOverDeterminant)
End Function

' a * fa - b * fb + c * fc (componentwise)
Private Function InverseTerm(a As Vector4f, fa As Vector4f, b As Vector4f, fb As Vector4f, c As Vector4f, fc As Vector4f) As Vector4f
    InverseTerm.x = a.x * fa.x - b.x * fb.x + c.x * fc.x
    InverseTerm.y = a.y * fa.y - b.y * fb.y + c.y * fc.y
    InverseTerm.z = a.z * fa.z - b.z * fb.z + c.z * fc.z
    InverseTerm.w = a.w * fa.w - b.w * fb.w + c.w * fc.w
End Function

' gtx/matrix_interpolation's extractMatrixRotation
Private Function extractMatrixRotation(m As mat4) As mat4
    extractMatrixRotation.m(0) = m.m(0): extractMatrixRotation.m(1) = m.m(1): extractMatrixRotation.m(2) = m.m(2)
    extractMatrixRotation.m(4) = m.m(4): extractMatrixRotation.m(5) = m.m(5): extractMatrixRotation.m(6) = m.m(6)
    extractMatrixRotation.m(8) = m.m(8): extractMatrixRotation.m(9) = m.m(9): extractMatrixRotation.m(10) = m.m(10)
    extractMatrixRotation.m(15) = 1!
End Function

' gtx/matrix_interpolation's axisAngle
Private Sub axisAngle(m As mat4, axis As Vector3f, angle As Single)
    Const epsilon As Single = 1.192093E-05     ' std::numeric_limits<float>::epsilon() * 1e2
    Const pi_f As Single = 3.141593!
    Dim nearSymmetrical As Boolean, nearIdentity As Boolean
    Dim xx As Single, yy As Single, zz As Single, xy As Single, xz As Single, yz As Single
    Dim angleCos As Single, v As Vector3f

    nearSymmetrical = Abs(m.m(1 * 4 + 0) - m.m(0 * 4 + 1)) < epsilon And _
        Abs(m.m(2 * 4 + 0) - m.m(0 * 4 + 2)) < epsilon And _
        Abs(m.m(2 * 4 + 1) - m.m(1 * 4 + 2)) < epsilon

    If nearSymmetrical Then
        nearIdentity = Abs(m.m(1 * 4 + 0) + m.m(0 * 4 + 1)) < epsilon And _
            Abs(m.m(2 * 4 + 0) + m.m(0 * 4 + 2)) < epsilon And _
            Abs(m.m(2 * 4 + 1) + m.m(1 * 4 + 2)) < epsilon And _
            Abs(m.m(0) + m.m(5) + m.m(10) - 3!) < epsilon
        If nearIdentity Then
            angle = 0!
            axis = glm_vec3(1!, 0!, 0!)
            Exit Sub
        End If
        angle = pi_f
        xx = (m.m(0) + 1!) * 0.5!
        yy = (m.m(5) + 1!) * 0.5!
        zz = (m.m(10) + 1!) * 0.5!
        xy = (m.m(1 * 4 + 0) + m.m(0 * 4 + 1)) * 0.25!
        xz = (m.m(2 * 4 + 0) + m.m(0 * 4 + 2)) * 0.25!
        yz = (m.m(2 * 4 + 1) + m.m(1 * 4 + 2)) * 0.25!
        If (xx > yy) And (xx > zz) Then
            If xx < epsilon Then
                axis.x = 0!
                axis.y = 0.7071!
                axis.z = 0.7071!
            Else
                axis.x = Sqr(xx)
                axis.y = xy / axis.x
                axis.z = xz / axis.x
            End If
        ElseIf yy > zz Then
            If yy < epsilon Then
                axis.x = 0.7071!
                axis.y = 0!
                axis.z = 0.7071!
            Else
                axis.y = Sqr(yy)
                axis.x = xy / axis.y
                axis.z = yz / axis.y
            End If
        Else
            If zz < epsilon Then
                axis.x = 0.7071!
                axis.y = 0.7071!
                axis.z = 0!
            Else
                axis.z = Sqr(zz)
                axis.x = xz / axis.z
                axis.y = yz / axis.z
            End If
        End If
        Exit Sub
    End If

    angleCos = (m.m(0) + m.m(5) + m.m(10) - 1!) * 0.5!
    If angleCos >= 1! Then
        angle = 0!
    ElseIf angleCos <= -1! Then
        angle = pi_f
    Else
        ' acos(angleCos)
        angle = Atn(-angleCos / Sqr(-angleCos * angleCos + 1#)) + 2# * Atn(1#)
    End If

    v = glm_vec3(m.m(1 * 4 + 2) - m.m(2 * 4 + 1), m.m(2 * 4 + 0) - m.m(0 * 4 + 2), m.m(0 * 4 + 1) - m.m(1 * 4 + 0))
    axis = glm_normalize3(v)
End Sub

' gtx/matrix_interpolation's axisAngleMatrix
Private Function axisAngleMatrix(axis As Vector3f, ByVal angle As Single) As mat4
    Dim c As Single, s As Single, t As Single, n As Vector3f
    c = Cos(angle)
    s = Sin(angle)
    t = 1! - c
    n = glm_normalize3(axis)

    Dim r As mat4
    r.m(0) = t * n.x * n.x + c: r.m(1) = t * n.x * n.y + n.z * s: r.m(2) = t * n.x * n.z - n.y * s: r.m(3) = 0!
    r.m(4) = t * n.x * n.y - n.z * s: r.m(5) = t * n.y * n.y + c: r.m(6) = t * n.y * n.z + n.x * s: r.m(7) = 0!
    r.m(8) = t * n.x * n.z + n.y * s: r.m(9) = t * n.y * n.z - n.x * s: r.m(10) = t * n.z * n.z + c: r.m(11) = 0!
    r.m(12) = 0!: r.m(13) = 0!: r.m(14) = 0!: r.m(15) = 1!
    axisAngleMatrix = r
End Function

' glm::interpolate(m1, m2, delta) (gtx/matrix_interpolation)
Public Function glm_interpolate(m1 As mat4, m2 As mat4, ByVal delta As Single) As mat4
    Dim m1rot As mat4, dltRotation As mat4, dltAxis As Vector3f, dltAngle As Single, m1rotT As mat4, outm As mat4, rotm As mat4
    m1rot = extractMatrixRotation(m1)
    m1rotT = glm_transpose(m1rot)
    dltRotation = glm_mat4_mul(m2, m1rotT)
    axisAngle dltRotation, dltAxis, dltAngle
    rotm = axisAngleMatrix(dltAxis, dltAngle * delta)
    outm = glm_mat4_mul(rotm, m1rot)
    outm.m(12) = m1.m(12) + delta * (m2.m(12) - m1.m(12))
    outm.m(13) = m1.m(13) + delta * (m2.m(13) - m1.m(13))
    outm.m(14) = m1.m(14) + delta * (m2.m(14) - m1.m(14))
    glm_interpolate = outm
End Function
