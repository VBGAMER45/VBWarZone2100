VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/pieblitfunc.cpp/.h: class GFX, a texture (or a colour
' buffer) and the vertex buffers of a 2D/3D primitive. Made by GFX_new
' (pieblitfunc_c). Private to WzGfx: only the renderer's modules use it
' (lib/sequence, its other user, is not ported). The vertex data are
' addresses (gfx_api::gfxFloat / gfxByte arrays). draw<PSO> takes the
' pso_type (a Friend method: it takes a mat4); which branch of the template
' runs depends on whether the PSO has textures (pso_has_textures).
' The destructor is Class_Terminate (C++'s delete).
Option Explicit

Private mType As GFXTYPE
Private mCoordsPerVertex As Long
Private mBuffers(0 To VBO_COUNT - 1) As gfx_api_buffer  ' = { nullptr }
Private mTexture As gfx_api_texture     ' = nullptr
Private mSize As Long

' GFX(GFXTYPE type, int coordsPerVertex = 3) : mType(type), mCoordsPerVertex(coordsPerVertex), mSize(0)
Friend Sub init(ByVal type_ As GFXTYPE, ByVal coordsPerVertex As Long)
    mType = type_
    mCoordsPerVertex = coordsPerVertex
    mSize = 0
End Sub

Public Sub loadTexture(ByVal filename As String, ByVal textureType As texture_type, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1)
    Dim extension As String, p As Long
    If Not (mType = GFX_TEXTURE) Then ASSERT_FAILURE "mType == GFX_TEXTURE", "lib/ivis_opengl/pieblitfunc.cpp:80", "loadTexture", "Wrong GFX type"
    p = InStrRev(filename, ".")             ' determine the filetype
    If p > 0 Then extension = Mid$(filename, p)
    If p = 0 Or (extension <> ".png") Then
        debug_ 88, LOG_ERROR, "loadTexture", "Bad image filename: %s", filename
        Exit Sub
    End If
    If Not mTexture Is Nothing Then Set mTexture = Nothing
    Set mTexture = gfx_api_context_loadTextureFromFile(filename, textureType, maxWidth, maxHeight)
End Sub

' void loadTexture(iV_Image&& bitmap, gfx_api::texture_type textureType, const std::string& filename, int maxWidth = -1, int maxHeight = -1)
Public Sub loadTexture_image(ByVal bitmap As iV_Image, ByVal textureType As texture_type, ByVal filename As String, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1)
    If Not (mType = GFX_TEXTURE) Then ASSERT_FAILURE "mType == GFX_TEXTURE", "lib/ivis_opengl/pieblitfunc.cpp:98", "loadTexture", "Wrong GFX type"
    If Not mTexture Is Nothing Then Set mTexture = Nothing
    Set mTexture = gfx_api_context_loadTextureFromUncompressedImage(bitmap, textureType, filename, maxWidth, maxHeight)
End Sub

Public Sub makeTexture(ByVal width As Long, ByVal height As Long, Optional ByVal format As pixel_format = pixel_format_FORMAT_RGBA8_UNORM_PACK8, Optional ByVal debugName As String = "")
    If Not (mType = GFX_TEXTURE) Then ASSERT_FAILURE "mType == GFX_TEXTURE", "lib/ivis_opengl/pieblitfunc.cpp:106", "makeTexture", "Wrong GFX type"
    If Not mTexture Is Nothing Then Set mTexture = Nothing
    Set mTexture = Nothing
    If width > 0 And height > 0 Then
        Set mTexture = gfx_api_context_create_texture(1, width, height, format, debugName)
    End If
End Sub

' image: Nothing = nullptr
Public Sub makeCompatibleTexture(ByVal image As iV_Image, ByVal filename As String)
    If Not (mType = GFX_TEXTURE) Then ASSERT_FAILURE "mType == GFX_TEXTURE", "lib/ivis_opengl/pieblitfunc.cpp:118", "makeCompatibleTexture", "Wrong GFX type"
    If Not mTexture Is Nothing Then Set mTexture = Nothing
    Set mTexture = Nothing
    If image Is Nothing Then
        Exit Sub
    End If
    Set mTexture = gfx_api_context_createTextureForCompatibleImageUploads(1, image, filename)
End Sub

Public Sub updateTexture(ByVal image As iV_Image)
    If Not (mType = GFX_TEXTURE) Then ASSERT_FAILURE "mType == GFX_TEXTURE", "lib/ivis_opengl/pieblitfunc.cpp:131", "updateTexture", "Wrong GFX type"
    If Not (Not mTexture Is Nothing) Then ASSERT_FAILURE "mTexture != nullptr", "lib/ivis_opengl/pieblitfunc.cpp:132", "updateTexture", "Null texture??": Exit Sub
    mTexture.upload 0, image
End Sub

' vertBuf, auxBuf: addresses of the vertex data (gfxFloat x mCoordsPerVertex per vertex) and of the
' texture coordinates (2 gfxFloat) or colours (4 gfxByte) per vertex
Public Sub buffers(ByVal vertices As Long, ByVal vertBuf As Long, ByVal auxBuf As Long)
    If mBuffers(VBO_VERTEX) Is Nothing Then Set mBuffers(VBO_VERTEX) = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_static_draw, "GFX::VBO_VERTEX")
    mBuffers(VBO_VERTEX).upload vertices * mCoordsPerVertex * 4, vertBuf

    If mType = GFX_TEXTURE Then
        If mBuffers(VBO_TEXCOORD) Is Nothing Then Set mBuffers(VBO_TEXCOORD) = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_static_draw, "GFX::VBO_TEXCOORD")
        mBuffers(VBO_TEXCOORD).upload vertices * 2 * 4, auxBuf
    ElseIf mType = GFX_COLOUR Then
        ' reusing texture buffer for colours for now
        If mBuffers(VBO_TEXCOORD) Is Nothing Then Set mBuffers(VBO_TEXCOORD) = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_static_draw, "GFX::VBO_TEXCOORD")
        mBuffers(VBO_TEXCOORD).upload vertices * 4 * 1, auxBuf
    End If
    mSize = vertices
End Sub

' template<typename PSO> draw(const glm::mat4 &modelViewProjectionMatrix): the two enable_if branches
Friend Sub draw(ByVal pso As pso_type, modelViewProjectionMatrix As mat4)
    Dim helper As pipeline_state_helper, cb As constant_buffer_type_SHADER_GFX_COLOUR
    Set helper = gfx_api_pso_get(pso)
    If pso_has_textures(pso) Then
        If mTexture Is Nothing Then Exit Sub
        helper.bind_textures mTexture
        helper.bind_vertex_buffers mBuffers(VBO_VERTEX), mBuffers(VBO_TEXCOORD)
        helper.draw mSize, 0
        helper.unbind_vertex_buffers mBuffers(VBO_VERTEX), mBuffers(VBO_TEXCOORD)
    Else
        cb.transform_matrix = modelViewProjectionMatrix
        helper.bind_constants VarPtr(cb), LenB(cb)
        helper.bind_vertex_buffers mBuffers(VBO_VERTEX), mBuffers(VBO_TEXCOORD)
        helper.draw mSize, 0
        helper.unbind_vertex_buffers mBuffers(VBO_VERTEX), mBuffers(VBO_TEXCOORD)
    End If
End Sub

' (VB) std::tuple_size<typename PSO::texture_tuple>::value != 0, for the PSOs a GFX draws (gfx_api.h:
' the GFX<...> instantiations and SkyboxPSO); the textureless ones use SHADER_GFX_COLOUR's constants
Private Function pso_has_textures(ByVal pso As pso_type) As Boolean
    Select Case pso
    Case RadarViewInsideFillPSO, RadarViewOutlinePSO
        pso_has_textures = False
    Case Else                           ' VideoPSO, BackDropPSO, SkyboxPSO, RadarPSO
        pso_has_textures = True
    End Select
End Function

Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To VBO_COUNT - 1
        Set mBuffers(i) = Nothing
    Next
    If Not mTexture Is Nothing Then Set mTexture = Nothing
End Sub
