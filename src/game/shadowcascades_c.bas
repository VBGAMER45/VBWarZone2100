Attribute VB_Name = "shadowcascades_c"
' src/shadowcascades.cpp, shadowcascades.h: the shadow map cascades' split
' depths and light matrices (portions derived from Sascha Willems'
' shadowmappingcascade.cpp, MIT licence, as the C++). output is the
' std::vector<Cascade>: output (0-based) and output_size.
Option Explicit

Public Type Cascade
    splitDepth As Single                ' = 0.f
    viewMatrix As mat4
    projectionMatrix As mat4
End Type

Public cascadeSplitLambda As Single     ' = 0.3f

' (VB) the C++ initialisers (Sub Main calls it)
Public Sub shadowcascades_c_init()
    cascadeSplitLambda = 0.3!
End Sub

Public Sub calculateShadowCascades(player As iView, ByVal terrainDistance As Single, baseViewMatrix As mat4, lightInvDir As Vector3f, ByVal SHADOW_MAP_CASCADE_COUNT As Long, output_() As Cascade, output_size As Long)
    Dim pitch As Single, adjustedZFar As Single, base_ As Double
    Dim cascadeSplits() As Single, nearClip As Single, farClip As Single, clipRange As Single
    Dim minZ As Single, maxZ As Single, range As Single, ratio As Single
    Dim i As Long, iSplit As Long, p As Single, log_ As Single, uniform As Single, d As Single
    Dim lastSplitDist As Single, splitDist As Single
    Dim frustumCorners(0 To 7) As Vector3f, invCam As mat4, persp As mat4, view_ As mat4, invCorner As Vector4f
    Dim dist As Vector3f, frustumCenter As Vector3f, radius As Single, distance2 As Single
    Dim maxExtents As Vector3f, minExtents As Vector3f, lightDir As Vector3f
    Dim lightViewMatrix As mat4, lightProjectionMatrix As mat4, eye As Vector3f

    output_size = 0                     ' output.clear()
    If SHADOW_MAP_CASCADE_COUNT <= 0 Then
        Exit Sub
    End If
    ReDim output_(0 To SHADOW_MAP_CASCADE_COUNT - 1)
    output_size = SHADOW_MAP_CASCADE_COUNT ' output.resize(SHADOW_MAP_CASCADE_COUNT)

    pitch = UNDEG(-player.r.x)

    ' (VB) pow of a negative base is NaN in C++; VB would raise an error (and a NaN poisons the FPU)
    base_ = (1.57! - pitch) / 1.57!
    If base_ >= 0# Then
        adjustedZFar = CSng((base_ ^ 0.9!) * 4000! + 2000! + terrainDistance)
    Else
        adjustedZFar = 2000! + terrainDistance
    End If

    ReDim cascadeSplits(0 To SHADOW_MAP_CASCADE_COUNT - 1)

    nearClip = 330!

    farClip = adjustedZFar 'pie_GetPerspectiveZFar();
    clipRange = farClip - nearClip

    minZ = nearClip
    maxZ = nearClip + clipRange

    range = maxZ - minZ
    ratio = maxZ / minZ

    ' Calculate split depths based on view camera frustum
    ' Based on method presented in https://developer.nvidia.com/gpugems/GPUGems3/gpugems3_ch10.html
    For i = 0 To SHADOW_MAP_CASCADE_COUNT - 1
        p = (i + 1) / CSng(SHADOW_MAP_CASCADE_COUNT)
        log_ = minZ * (ratio ^ p)
        uniform = minZ + range * p
        d = cascadeSplitLambda * (log_ - uniform) + uniform
        cascadeSplits(i) = (d - nearClip) / clipRange
    Next

    ' Calculate orthographic projection matrix for each cascade
    lastSplitDist = 0!
    For iSplit = 0 To SHADOW_MAP_CASCADE_COUNT - 1
        splitDist = cascadeSplits(iSplit)

        frustumCorners(0) = glm_vec3(-1!, 1!, -1!) ' opengl
        frustumCorners(1) = glm_vec3(1!, 1!, -1!)
        frustumCorners(2) = glm_vec3(1!, -1!, -1!)
        frustumCorners(3) = glm_vec3(-1!, -1!, -1!)
        frustumCorners(4) = glm_vec3(-1!, 1!, 1!)
        frustumCorners(5) = glm_vec3(1!, 1!, 1!)
        frustumCorners(6) = glm_vec3(1!, -1!, 1!)
        frustumCorners(7) = glm_vec3(-1!, -1!, 1!)

        ' Project frustum corners into world space
        pie_PerspectiveGetConstrained nearClip, adjustedZFar, VarPtr(persp)
        view_ = glm_mat4_mul(baseViewMatrix, glm_translate(glm_vec3(-player.p.x, 0!, player.p.z)))
        invCam = glm_inverse(glm_mat4_mul(persp, view_))
        For i = 0 To 7
            invCorner = glm_mat4_mul_vec4(invCam, glm_vec4_v3(frustumCorners(i), 1!))
            frustumCorners(i) = glm_vec3(invCorner.x / invCorner.w, invCorner.y / invCorner.w, invCorner.z / invCorner.w)
        Next

        For i = 0 To 3
            dist = glm_vec3_sub(frustumCorners(i + 4), frustumCorners(i))
            frustumCorners(i + 4) = glm_vec3_add(frustumCorners(i), glm_vec3_mul(dist, splitDist))
            frustumCorners(i) = glm_vec3_add(frustumCorners(i), glm_vec3_mul(dist, lastSplitDist))
        Next

        ' Get frustum center
        frustumCenter = glm_vec3(0!, 0!, 0!)
        For i = 0 To 7
            frustumCenter = glm_vec3_add(frustumCenter, frustumCorners(i))
        Next
        frustumCenter = glm_vec3_mul(frustumCenter, 1! / 8!)

        radius = 0!
        For i = 0 To 7
            distance2 = glm_length3(glm_vec3_sub(frustumCorners(i), frustumCenter))
            If distance2 > radius Then radius = distance2
        Next
        radius = -Int(-(radius * 16!)) / 16!     ' std::ceil

        maxExtents = glm_vec3(radius, radius, radius)
        minExtents = glm_vec3(-maxExtents.x, -maxExtents.y, -maxExtents.z)

        lightDir = glm_normalize3(lightInvDir) 'normalize(-lightPos);
        eye = glm_vec3_add(frustumCenter, lightDir)
        lightViewMatrix = glm_lookAt(eye, frustumCenter, glm_vec3(0!, 1!, 0!))
        lightProjectionMatrix = glm_mat4_mul(glm_ortho6(minExtents.x, maxExtents.x, minExtents.y, maxExtents.y, minExtents.z, maxExtents.z), glm_scale(glm_vec3(1!, 1!, -1!)))

        ' Store split distance and matrix in cascade
        output_(iSplit).splitDepth = (-nearClip + splitDist * clipRange)
        output_(iSplit).viewMatrix = lightViewMatrix
        output_(iSplit).projectionMatrix = lightProjectionMatrix

        lastSplitDist = cascadeSplits(iSplit)
    Next
End Sub
