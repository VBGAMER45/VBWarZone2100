Attribute VB_Name = "culling_c"
' lib/ivis_opengl/culling.cpp/.h: bounding boxes in clip space and view
' frustums as intersections of half spaces. BoundingBox (std::array<glm::vec3,
' 8>) is a UDT of 8 Vector3f. A HalfSpaceCheck (a std::function, only ever
' one of pielighting.cpp's lambdas) is a dispatch number with its captured
' values (CONVENTIONS); HalfSpaceCheck_call runs it.
Option Explicit

Public Type BoundingBox
    v(0 To 7) As Vector3f
End Type

' (VB) the lambdas a HalfSpaceCheck can be: in.x >= bound, in.x <= bound, (yAxisInverted ? -in.y : in.y) >= bound,
' (yAxisInverted ? -in.y : in.y) <= bound, in.z >= bound, in.z <= bound
Public Enum HalfSpaceCheckFn
    HalfSpace_x_ge
    HalfSpace_x_le
    HalfSpace_y_ge
    HalfSpace_y_le
    HalfSpace_z_ge
    HalfSpace_z_le
End Enum

Public Type HalfSpaceCheck
    fn As HalfSpaceCheckFn
    bound As Single                     ' (VB) the captured bound (or the lambda's constant)
    yAxisInverted As Boolean            ' (VB) the captured yAxisInverted
End Type

' Define a view frustum (as an intersection of half space
Public Type IntersectionOfHalfSpace
    checks(0 To 5) As HalfSpaceCheck
End Type

' (VB) a HalfSpaceCheck
Public Function HalfSpaceCheck_new(ByVal fn As HalfSpaceCheckFn, ByVal bound As Single, Optional ByVal yAxisInverted As Boolean = False) As HalfSpaceCheck
    HalfSpaceCheck_new.fn = fn
    HalfSpaceCheck_new.bound = bound
    HalfSpaceCheck_new.yAxisInverted = yAxisInverted
End Function

' (VB) predicate(in); a comparison with a NaN is false, as in C
Public Function HalfSpaceCheck_call(predicate As HalfSpaceCheck, in_ As Vector3f) As Boolean
    Dim value As Single
    Select Case predicate.fn
    Case HalfSpace_x_ge, HalfSpace_x_le
        value = in_.x
    Case HalfSpace_y_ge, HalfSpace_y_le
        If predicate.yAxisInverted Then value = -in_.y Else value = in_.y
    Case Else
        value = in_.z
    End Select
    If IsNaNF(value) Then
        HalfSpaceCheck_call = False
        Exit Function
    End If
    Select Case predicate.fn
    Case HalfSpace_x_ge, HalfSpace_y_ge, HalfSpace_z_ge
        HalfSpaceCheck_call = (value >= predicate.bound)
    Case Else
        HalfSpaceCheck_call = (value <= predicate.bound)
    End Select
End Function

'/ Project a bounding box in clip space
Public Function transformBoundingBox(worldViewProjectionMatrix As mat4, worldSpaceBoundingBox As BoundingBox) As BoundingBox
    Dim bboxInClipSpace As BoundingBox, i As Long, tmp As Vector4f
    For i = 0 To 7
        tmp = glm_mat4_mul_vec4(worldViewProjectionMatrix, glm_vec4_v3(worldSpaceBoundingBox.v(i), 1!))
        tmp = glm_vec4_div(tmp, tmp.w)
        bboxInClipSpace.v(i) = glm_vec3(tmp.x, tmp.y, tmp.z)
    Next
    transformBoundingBox = bboxInClipSpace
End Function

Public Function isBBoxInClipSpace(intersectionOfHalfSpace_ As IntersectionOfHalfSpace, points As BoundingBox) As Boolean
    Dim p As Long
    ' We test against the complement of the half space
    ' If all points lies in the complement a half space, it can't be part of the intersection
    For p = 0 To 5
        If CheckAllPointsInSpace(points, intersectionOfHalfSpace_.checks(p)) Then
            isBBoxInClipSpace = False
            Exit Function
        End If
    Next
    isBBoxInClipSpace = True
End Function

' the lambda CheckAllPointsInSpace: std::all_of(points, !predicate(v))
Private Function CheckAllPointsInSpace(points As BoundingBox, predicate As HalfSpaceCheck) As Boolean
    Dim i As Long
    For i = 0 To 7
        If HalfSpaceCheck_call(predicate, points.v(i)) Then
            CheckAllPointsInSpace = False
            Exit Function
        End If
    Next
    CheckAllPointsInSpace = True
End Function
