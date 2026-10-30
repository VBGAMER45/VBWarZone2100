Attribute VB_Name = "pienormalize_h"
' lib/ivis_opengl/pienormalize.h: the surface normal of a triangle (float
' operations through glm_vb).
Option Explicit

Public Function normalise(a As Vector3f) As Vector3f
    Dim d As Single
    d = v3dot(a, a)
    If IsNaNF(d) Then                       ' (VB) a NaN is not > 0
        normalise = a
    ElseIf d > 0! Then
        normalise = v3normalize(a)
    Else
        normalise = a
    End If
End Function

' Calculate surface normal
' Eg. if a polygon (with n points in clockwise order) normal is required,
' \c p1 = point 0, \c p2 = point 1, \c p3 = point n-1
' \param p1,p2,p3 Points for forming 2 vector for cross product
' \return Normal vector
Public Function pie_SurfaceNormal3fv(p1 As Vector3f, p2 As Vector3f, p3 As Vector3f) As Vector3f
    pie_SurfaceNormal3fv = normalise(v3cross(v3sub(p3, p1), v3sub(p2, p1)))
End Function
