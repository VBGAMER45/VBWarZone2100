Attribute VB_Name = "piefunc_c"
' lib/ivis_opengl/piefunc.cpp: extended render routines for 3D rendering
' (the radar's viewing window, the coloured triangle, the skybox). The
' exe's arrays and matrices come by address in the _p forms (GfxApi).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private skyboxGfx As GFX                ' = nullptr
Private radarViewGfx(0 To 1) As GFX     ' = { nullptr, nullptr }

Private Const VW_VERTICES As Long = 5

' v: v(0 .. 3) (C: Vector3i *v)
Public Sub pie_SetViewingWindow(v() As Vector3i, ByVal colour As Long)
    Dim pieVrts_insidefill(0 To VW_VERTICES - 2) As Vector3i
    Dim pieVrts_outline(0 To VW_VERTICES - 1) As Vector3i
    Dim vert(0 To VW_VERTICES * 2 - 1) As Single
    Dim cols(0 To VW_VERTICES * 4 - 1) As Byte
    Dim i As Long

    If radarViewGfx(0) Is Nothing Then
        Set radarViewGfx(0) = GFX_new(GFX_COLOUR, 2)
        Set radarViewGfx(1) = GFX_new(GFX_COLOUR, 2)
    End If

    ' static_assert(VW_VERTICES == 5, "Assumption that VW_VERTICES == 5 invalid. Update the following code.");
    pieVrts_outline(0) = v(1)
    pieVrts_outline(1) = v(0)
    pieVrts_outline(2) = v(2)
    pieVrts_outline(3) = v(3)
    pieVrts_outline(4) = v(1)
    pieVrts_insidefill(0) = v(2)
    pieVrts_insidefill(1) = v(3)
    pieVrts_insidefill(2) = v(0)
    pieVrts_insidefill(3) = v(1)

    ' Initialize the RGB values for both, and the alpha for the inside fill
    For i = 0 To VW_VERTICES * 4 - 1 Step 4
        cols(i + 0) = PIELIGHT_r(colour)
        cols(i + 1) = PIELIGHT_g(colour)
        cols(i + 2) = PIELIGHT_b(colour)
        cols(i + 3) = PIELIGHT_a(colour) \ 2   ' >> 1
    Next

    ' Inside fill
    For i = 0 To (VW_VERTICES - 1) - 1
        vert(i * 2 + 0) = pieVrts_insidefill(i).x
        vert(i * 2 + 1) = pieVrts_insidefill(i).y
    Next
    radarViewGfx(0).buffers VW_VERTICES - 1, VarPtr(vert(0)), VarPtr(cols(0))

    ' Outline
    For i = 0 To VW_VERTICES - 1
        vert(i * 2 + 0) = pieVrts_outline(i).x
        vert(i * 2 + 1) = pieVrts_outline(i).y
    Next
    For i = 0 To VW_VERTICES * 4 - 1 Step 4 ' set the proper alpha for the outline
        cols(i + 3) = PIELIGHT_a(colour)
    Next
    radarViewGfx(1).buffers VW_VERTICES, VarPtr(vert(0)), VarPtr(cols(0))
End Sub

' (VB) pie_SetViewingWindow for the exe: pV is the address of 4 Vector3i
Public Sub pie_SetViewingWindow_p(ByVal pV As Long, ByVal colour As Long)
    Dim v(0 To 3) As Vector3i
    CopyMemory v(0), ByVal pV, 4 * LenB(v(0))
    pie_SetViewingWindow v, colour
End Sub

Public Sub pie_DrawViewingWindow(modelViewProjectionMatrix As mat4)
    Dim cb As constant_buffer_type_SHADER_GFX_COLOUR
    cb.transform_matrix = modelViewProjectionMatrix
    gfx_api_pso_get(RadarViewInsideFillPSO).bind
    gfx_api_pso_get(RadarViewInsideFillPSO).bind_constants VarPtr(cb), LenB(cb)
    radarViewGfx(0).draw RadarViewInsideFillPSO, modelViewProjectionMatrix
    gfx_api_pso_get(RadarViewOutlinePSO).bind
    gfx_api_pso_get(RadarViewOutlinePSO).bind_constants VarPtr(cb), LenB(cb)
    radarViewGfx(1).draw RadarViewOutlinePSO, modelViewProjectionMatrix
End Sub

Public Sub pie_ViewingWindow_Shutdown()
    If Not radarViewGfx(0) Is Nothing Then
        Set radarViewGfx(0) = Nothing
    End If
    If Not radarViewGfx(1) Is Nothing Then
        Set radarViewGfx(1) = Nothing
    End If
End Sub

' vrt: vrt(0 .. 2) (C: const std::array<Vector3f, 3> &)
Public Sub pie_TransColouredTriangle(vrt() As Vector3f, ByVal c As Long, modelViewMatrix As mat4)
    Dim color As Vector4f, cb As constant_buffer_type_SHADER_GENERIC_COLOR
    color = glm_vec4(PIELIGHT_r(c) / 255!, PIELIGHT_g(c) / 255!, PIELIGHT_b(c) / 255!, 128! / 255!)

    gfx_api_pso_get(TransColouredTrianglePSO).bind
    cb.transform_matrix = glm_mat4_mul(pie_UIPerspectiveGet(), modelViewMatrix)
    cb.unused = glm_vec2(0!, 0!)
    cb.unused2 = glm_vec2(0!, 0!)
    cb.colour = color
    gfx_api_pso_get(TransColouredTrianglePSO).bind_constants VarPtr(cb), LenB(cb)
    gfx_api_context_bind_streamed_vertex_buffers VarPtr(vrt(0)), 3 * LenB(vrt(0))
    gfx_api_pso_get(TransColouredTrianglePSO).draw 3, 0
    gfx_api_context_disable_all_vertex_buffers
End Sub

' (VB) pie_TransColouredTriangle for the exe: pVrt is the address of 3 Vector3f, pModelViewMatrix of the mat4
Public Sub pie_TransColouredTriangle_p(ByVal pVrt As Long, ByVal c As Long, ByVal pModelViewMatrix As Long)
    Dim vrt(0 To 2) As Vector3f, m As mat4
    CopyMemory vrt(0), ByVal pVrt, 3 * LenB(vrt(0))
    m = mat4_fromPtr(pModelViewMatrix)
    pie_TransColouredTriangle vrt, c, m
End Sub

Public Sub pie_Skybox_Init()
    Dim narrowingOfTop As Single, top As Single, middle As Single, baseline As Single, bottom As Single
    Dim narrowCoordinate As Single
    Dim northWestBottom As Vector3f, northWestMiddle As Vector3f, northWestBaseline As Vector3f, northWestTop As Vector3f
    Dim northEastBottom As Vector3f, northEastTop As Vector3f, northEastMiddle As Vector3f, northEastBaseline As Vector3f
    Dim southEastBottom As Vector3f, southEastTop As Vector3f, southEastMiddle As Vector3f, southEastBaseline As Vector3f
    Dim southWestBottom As Vector3f, southWestMiddle As Vector3f, southWestBaseline As Vector3f, southWestTop As Vector3f
    Dim vertex(0 To 77) As Vector3f, n As Long
    Dim uvTopLeft As Vector2f, uvMiddleLeft As Vector2f, uvBaselineLeft As Vector2f
    Dim uvTopRight As Vector2f, uvMiddleRight As Vector2f, uvBaselineRight As Vector2f, uvNull As Vector2f
    Dim texc(0 To 77) As Vector2f, q As Long

    narrowingOfTop = 0.15!
    top = 1
    middle = 0.15!
    baseline = -0.15! - 0.3! ' Additional lowering of baseline for where fog starts fading, to reduce certain artifacts
    bottom = -1

    ' Skybox looks like this from the side:
    '
    '             _____      top, slightly narrower to give perspective
    '            /     .
    '           /       .
    '           +-------+    middle
    '           |       |
    '           +-------+    baseline of map
    '           |       |
    '           +-------+    bottom
    ' bottom cap ---^

    ' The skybox is shown from baseline to top. below the baseline, the color fades to black (if fog enabled - otherwise a color of the skybox bottom).
    ' We extended the skybox downwards to be able to paint more of the screen with fog.

    narrowCoordinate = 1 - narrowingOfTop
    northWestBottom = glm_vec3(-1, bottom, 1) ' nw
    northWestMiddle = glm_vec3(-1, middle, 1)
    northWestBaseline = glm_vec3(-1, baseline, 1)
    northWestTop = glm_vec3(-narrowCoordinate, top, narrowCoordinate)
    northEastBottom = glm_vec3(1, bottom, 1) ' ne
    northEastTop = glm_vec3(narrowCoordinate, top, narrowCoordinate)
    northEastMiddle = glm_vec3(1, middle, 1)
    northEastBaseline = glm_vec3(1, baseline, 1)
    southEastBottom = glm_vec3(1, bottom, -1) ' se
    southEastTop = glm_vec3(narrowCoordinate, top, -narrowCoordinate)
    southEastMiddle = glm_vec3(1, middle, -1)
    southEastBaseline = glm_vec3(1, baseline, -1)
    southWestBottom = glm_vec3(-1, bottom, -1) ' sw
    southWestMiddle = glm_vec3(-1, middle, -1)
    southWestBaseline = glm_vec3(-1, baseline, -1)
    southWestTop = glm_vec3(-narrowCoordinate, top, -narrowCoordinate)

    ' const std::array<Vector3f, 78> vertex{
    ' North quads
    Put3 vertex, n, northWestMiddle, northWestTop, northEastMiddle
    Put3 vertex, n, northEastTop, northEastMiddle, northWestTop
    Put3 vertex, n, northWestBaseline, northWestMiddle, northEastBaseline
    Put3 vertex, n, northEastMiddle, northEastBaseline, northWestMiddle
    Put3 vertex, n, northWestBottom, northWestBaseline, northEastBottom
    Put3 vertex, n, northEastBaseline, northEastBottom, northWestBaseline
    ' East quads
    Put3 vertex, n, northEastMiddle, northEastTop, southEastMiddle
    Put3 vertex, n, southEastTop, southEastMiddle, northEastTop
    Put3 vertex, n, northEastBaseline, northEastMiddle, southEastBaseline
    Put3 vertex, n, southEastMiddle, southEastBaseline, northEastMiddle
    Put3 vertex, n, northEastBottom, northEastBaseline, southEastBottom
    Put3 vertex, n, southEastBaseline, southEastBottom, northEastBaseline
    ' South quads
    Put3 vertex, n, southEastMiddle, southEastTop, southWestMiddle
    Put3 vertex, n, southWestTop, southWestMiddle, southEastTop
    Put3 vertex, n, southEastBaseline, southEastMiddle, southWestBaseline
    Put3 vertex, n, southWestMiddle, southWestBaseline, southEastMiddle
    Put3 vertex, n, southEastBottom, southEastBaseline, southWestBottom
    Put3 vertex, n, southWestBaseline, southWestBottom, southEastBaseline
    ' West quads
    Put3 vertex, n, southWestMiddle, southWestTop, northWestMiddle
    Put3 vertex, n, northWestTop, northWestMiddle, southWestTop
    Put3 vertex, n, southWestBaseline, southWestMiddle, northWestBaseline
    Put3 vertex, n, northWestMiddle, northWestBaseline, southWestMiddle
    Put3 vertex, n, southWestBottom, southWestBaseline, northWestBottom
    Put3 vertex, n, northWestBaseline, northWestBottom, southWestBaseline
    ' Bottom quads
    Put3 vertex, n, southWestBottom, northWestBottom, northEastBottom
    Put3 vertex, n, northEastBottom, southEastBottom, southWestBottom

    uvTopLeft = glm_vec2(0, 0.01!) ' fractions of 0 and 1 to avoid mipmap bleeding
    uvMiddleLeft = glm_vec2(0, 1 - middle)
    uvBaselineLeft = glm_vec2(0, 0.99!)
    uvTopRight = glm_vec2(2, 0.01!)     ' 2 = each side of the skybox contains 2 repeats of the same texture
    uvMiddleRight = glm_vec2(2, 1 - middle)
    uvBaselineRight = glm_vec2(2, 0.99!)
    uvNull = glm_vec2(0, 0.99!) ' = single color, blend in with sides.

    ' const std::array<Vector2f, 78> texc =
    For q = 0 To 3
        ' North, East, West, South quad
        Put2 texc, q * 18, uvMiddleLeft, uvTopLeft, uvMiddleRight
        Put2 texc, q * 18 + 3, uvTopRight, uvMiddleRight, uvTopLeft
        Put2 texc, q * 18 + 6, uvBaselineLeft, uvMiddleLeft, uvBaselineRight
        Put2 texc, q * 18 + 9, uvMiddleRight, uvBaselineRight, uvMiddleLeft
        Put2 texc, q * 18 + 12, uvNull, uvNull, uvNull ' bottom part has no color.
        Put2 texc, q * 18 + 15, uvNull, uvNull, uvNull
    Next
    ' Bottom quad
    Put2 texc, 72, uvNull, uvNull, uvNull ' bottom cap has no color
    Put2 texc, 75, uvNull, uvNull, uvNull

    gfx_api_context_debugStringMarker "Initializing skybox"
    Set skyboxGfx = GFX_new(GFX_TEXTURE, 3)
    skyboxGfx.buffers 78, VarPtr(vertex(0)), VarPtr(texc(0))
End Sub

' (VB) three entries of the skybox's vertex list, from n on
Private Sub Put3(arr() As Vector3f, n As Long, a As Vector3f, b As Vector3f, c As Vector3f)
    arr(n) = a
    arr(n + 1) = b
    arr(n + 2) = c
    n = n + 3
End Sub

' (VB) three entries of the skybox's texture coordinates, at i
Private Sub Put2(arr() As Vector2f, ByVal i As Long, a As Vector2f, b As Vector2f, c As Vector2f)
    arr(i) = a
    arr(i + 1) = b
    arr(i + 2) = c
End Sub

Public Sub pie_Skybox_Texture(ByVal filename As String)
    skyboxGfx.loadTexture filename, texture_type_game_texture
End Sub

Public Sub pie_Skybox_Shutdown()
    Set skyboxGfx = Nothing
End Sub

Public Sub pie_DrawSkybox(ByVal scale_ As Single, projectionMatrix As mat4, viewMatrix As mat4)
    Dim modelViewProjectionMatrix As mat4, renderState As RENDER_STATE, fogColor As Vector4f
    Dim cb As constant_buffer_type_SHADER_SKYBOX
    modelViewProjectionMatrix = glm_mat4_mul(glm_mat4_mul(projectionMatrix, viewMatrix), glm_scale(glm_vec3(scale_, scale_ / 2!, scale_)))

    renderState = getCurrentRenderState()
    fogColor = pielightToRGBAVec4(renderState.fogColour)

    gfx_api_pso_get(SkyboxPSO).bind
    cb.transform_matrix = modelViewProjectionMatrix
    cb.color = glm_vec4(1!, 1!, 1!, 1!)
    cb.fog_color = fogColor
    cb.fog_enabled = -renderState.fogEnabled
    gfx_api_pso_get(SkyboxPSO).bind_constants VarPtr(cb), LenB(cb)
    skyboxGfx.draw SkyboxPSO, modelViewProjectionMatrix
End Sub
