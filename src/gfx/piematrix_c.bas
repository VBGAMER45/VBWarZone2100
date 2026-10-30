Attribute VB_Name = "piematrix_c"
' lib/ivis_opengl/piematrix.cpp: matrix manipulation functions.
' The perspective getters return the mat4 (the exe gets it through an
' address: GfxApi); pie_PerspectiveGetCache's modifyGeneratedMatrix is a
' Boolean: its three callers pass the same lambda.
Option Explicit


' static float perspectiveZClose = 330.f; static float perspectiveZFar = 45000.f; (never changed)
Private Const perspectiveZClose As Single = 330!
Private Const perspectiveZFar As Single = 45000!

'***************************************************************************
'
'	Local Definitions
'
'***************************************************************************

Private Enum PerspectiveModes
    PerspectiveModes_Camera
    PerspectiveModes_Skybox
    PerspectiveModes_ScreenUI
End Enum
Private Const Num_Perspective_Modes As Long = PerspectiveModes_ScreenUI + 1

Private Type PerspectiveCache
    currentPerspectiveMatrix As mat4    ' = glm::mat4()
    width_ As Single                    ' _width = 0.f
    height_ As Single                   ' _height = 0.f
    rendSurface_xcentre_ As Long        ' _rendSurface_xcentre = 0
    rendSurface_ycentre_ As Long        ' _rendSurface_ycentre = 0
    zfar_ As Single                     ' _zfar = 0.f
End Type
' Since computing the perspective matrix is a semi-costly operation, store a cache here
Private perspectiveCaches(0 To Num_Perspective_Modes - 1) As PerspectiveCache

'!
' * 3D vector perspective projection
' * Projects 3D vector into 2D screen space
' * \param v3d       3D vector to project
' * \param perspectiveViewMatrix     precomputed perspective multipled by view matrix
' * \param[out] v2d  resulting 2D vector
' * \return projected z component of v2d
Public Function pie_RotateProjectWithPerspective(v3d As Vector3i, perspectiveViewMatrix As mat4, v2d As Vector2i) As Long
    Dim hackScaleFactor As Single, v As Vector4f, xx As Single, yy As Single
    hackScaleFactor = 1! / (3 * 330)     ' HACK: This seems to work by experimentation, not sure why.

    '
    ' * v = curMatrix . v3d
    '
    v = glm_mat4_mul_vec4(perspectiveViewMatrix, glm_vec4(v3d.x, v3d.y, v3d.z, 1!))

    ' (VB) xx = v.x / v.w and yy = v.y / v.w are used only where v.w >= 256 * hackScaleFactor (> 0):
    ' computed there, as VB's / raises on a division by zero

    If v.w < 256 * hackScaleFactor Then
        v2d.x = LONG_WAY 'just along way off screen
        v2d.y = LONG_WAY
    Else
        xx = v.x / v.w
        yy = v.y / v.w
        v2d.x = ToInt((0.5 + 0.5 * xx) * pie_GetVideoBufferWidth())
        v2d.y = ToInt((0.5 - 0.5 * yy) * pie_GetVideoBufferHeight())
    End If

    pie_RotateProjectWithPerspective = ToInt(v.w)
End Function

' (VB) pie_RotateProjectWithPerspective for the exe: the vectors as their fields, the matrix by address
Public Function pie_RotateProjectWithPerspective_p(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal pPerspectiveViewMatrix As Long, v2dx As Long, v2dy As Long) As Long
    Dim v3d As Vector3i, v2d As Vector2i, m As mat4
    v3d.x = x: v3d.y = y: v3d.z = z
    m = mat4_fromPtr(pPerspectiveViewMatrix)
    v2d.x = v2dx: v2d.y = v2dy
    pie_RotateProjectWithPerspective_p = pie_RotateProjectWithPerspective(v3d, m, v2d)
    v2dx = v2d.x: v2dy = v2d.y
End Function

Public Function pie_PerspectiveGetConstrained(ByVal zNear As Single, ByVal zFar As Single) As mat4
    Dim width As Single, height As Single, xangle As Single, yangle As Single
    Dim fr As mat4, sc As mat4
    width = MaxL(pie_GetVideoBufferWidth(), 1)  ' Require width > 0 && height > 0, to avoid glScalef(1, 1, -1) crashing in some graphics drivers.
    height = MaxL(pie_GetVideoBufferHeight(), 1)

    ' can be a semi-costly operation
    xangle = width / 6!
    yangle = height / 6!
'   glm::translate(glm::vec3((2.f * rendSurface.xcentre - width) / width, (height - 2.f * rendSurface.ycentre) / height, 0.f)) *
    fr = glm_frustum(-xangle, xangle, -yangle, yangle, zNear, zFar)
    sc = glm_scale(glm_vec3(1!, 1!, -1!))
    pie_PerspectiveGetConstrained = glm_mat4_mul(fr, sc)
End Function

' modifyGeneratedMatrix: (VB) whether to apply the lambda the callers pass:
' matrix = glm::translate(glm::vec3((2.f * rendSurface.xcentre - width) / width, (height - 2.f * rendSurface.ycentre) / height, 0.f)) * matrix;
Private Function pie_PerspectiveGetCache(ByVal mode As PerspectiveModes, ByVal zNear As Single, ByVal zFar As Single, ByVal modifyGeneratedMatrix As Boolean) As mat4
    Dim width As Single, height As Single, t As mat4
    width = MaxL(pie_GetVideoBufferWidth(), 1)  ' Require width > 0 && height > 0, to avoid glScalef(1, 1, -1) crashing in some graphics drivers.
    height = MaxL(pie_GetVideoBufferHeight(), 1)

    If width <> perspectiveCaches(mode).width_ Or height <> perspectiveCaches(mode).height_ Or rendSurface.xcentre <> perspectiveCaches(mode).rendSurface_xcentre_ Or rendSurface.ycentre <> perspectiveCaches(mode).rendSurface_ycentre_ Or zFar <> perspectiveCaches(mode).zfar_ Then
        ' update the current perspective matrix (can be a semi-costly operation)
        perspectiveCaches(mode).currentPerspectiveMatrix = pie_PerspectiveGetConstrained(zNear, zFar)
        If modifyGeneratedMatrix Then
            t = glm_translate(glm_vec3((2! * rendSurface.xcentre - width) / width, (height - 2! * rendSurface.ycentre) / height, 0!))
            perspectiveCaches(mode).currentPerspectiveMatrix = glm_mat4_mul(t, perspectiveCaches(mode).currentPerspectiveMatrix)
        End If
        perspectiveCaches(mode).width_ = width
        perspectiveCaches(mode).height_ = height
        perspectiveCaches(mode).rendSurface_xcentre_ = rendSurface.xcentre
        perspectiveCaches(mode).rendSurface_ycentre_ = rendSurface.ycentre
        perspectiveCaches(mode).zfar_ = perspectiveZFar
    End If

    pie_PerspectiveGetCache = perspectiveCaches(mode).currentPerspectiveMatrix
End Function

Public Function pie_SkyboxPerspectiveGet() As mat4
    pie_SkyboxPerspectiveGet = pie_PerspectiveGetCache(PerspectiveModes_Skybox, 330!, 100000!, True)
End Function

Public Function pie_PerspectiveGet() As mat4
    pie_PerspectiveGet = pie_PerspectiveGetCache(PerspectiveModes_Camera, perspectiveZClose, perspectiveZFar, True)
End Function

Public Function pie_UIPerspectiveGet() As mat4
    pie_UIPerspectiveGet = pie_PerspectiveGetCache(PerspectiveModes_ScreenUI, 330!, 100000!, True)
End Function

Public Function pie_getPerspectiveZFar() As Single
    pie_getPerspectiveZFar = perspectiveZFar
End Function

Public Sub pie_Begin3DScene()
    gfx_api_context_set_depth_range 0!, 1!
End Sub

Public Sub pie_BeginInterface()
    gfx_api_context_set_depth_range 0!, 0.1!
End Sub

Public Sub pie_SetGeometricOffset(ByVal x As Long, ByVal y As Long)
    rendSurface.xcentre = x
    rendSurface.ycentre = y
End Sub

' std::max(int, int)
Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MaxL = b Else MaxL = a
End Function
