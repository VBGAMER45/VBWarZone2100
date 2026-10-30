Attribute VB_Name = "shadows_h"
' lib/ivis_opengl/shadows.h: ShadowCascadesInfo, which the exe's display code
' (shadowcascades.cpp, display3d.cpp) fills and piedraw_c reads: compiled into
' WzGfx and the exe (src/shared/gfx), passed to pie_DrawAllMeshes by address
' (VarPtr). WZ_MAX_SHADOW_CASCADES is gfx_api_h's.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Type ShadowCascadesInfo
    shadowMVPMatrix(0 To WZ_MAX_SHADOW_CASCADES - 1) As mat4 ' = {glm::mat4(1.f)}: ShadowCascadesInfo_new
    shadowCascadeSplit(0 To WZ_MAX_SHADOW_CASCADES - 1) As Single ' = {1.f}
    shadowMapSize As Long               ' = 0
End Type

' (VB) ShadowCascadesInfo's default member initialisers: the first matrix the identity, the
' first split 1 (the others value-initialised: zero)
Public Function ShadowCascadesInfo_new() As ShadowCascadesInfo
    Dim r As ShadowCascadesInfo
    r.shadowMVPMatrix(0).m(0) = 1!
    r.shadowMVPMatrix(0).m(5) = 1!
    r.shadowMVPMatrix(0).m(10) = 1!
    r.shadowMVPMatrix(0).m(15) = 1!
    r.shadowCascadeSplit(0) = 1!
    ShadowCascadesInfo_new = r
End Function

' (VB) the ShadowCascadesInfo at address p (VarPtr of the caller's; it holds no strings or dynamic arrays)
Public Sub shadowCascadesInfo_fromPtr(ByVal p As Long, out As ShadowCascadesInfo)
    CopyMemory out, ByVal p, LenB(out)
End Sub
