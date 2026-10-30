Attribute VB_Name = "gfx_api_h"
' lib/ivis_opengl/gfx_api.h: the constant buffers (constant_buffer_type<SHADER_X>
' is constant_buffer_type_SHADER_X; the named uniform structs keep their names)
' and the interleaved vertex and instance data. Compiled into WzGfx and the exe
' (src/shared/gfx): the exe fills them and hands WzGfx their address and size
' (pipeline_state_helper.bind_constants VarPtr(cb), LenB(cb)); the GL backend
' copies them into its own copy of the same type. Every field is 4 bytes wide,
' so VB lays them out as C++ does (no padding): the vertex structs' offsets
' are the offsetof values of the vertex descriptions (gfx_api_pso_c).
' NOTE: Be very careful changing these constant_buffer_type structs;
'       they must match std140 layout rules (see: the Vulkan shaders)
Option Explicit

Public Const WZ_MAX_SHADOW_CASCADES As Long = 3      ' shadows.h

Public Const gfx_api_max_lights As Long = 128
Public Const gfx_api_max_indexed_lights As Long = 512
Public Const gfx_api_bucket_dimension As Long = 8

' Only change once per frame
Public Type Draw3DShapeGlobalUniforms
    ProjectionMatrix As mat4
    ViewMatrix As mat4
    ShadowMapMVPMatrix As mat4
    sunPos As Vector4f
    sceneColor As Vector4f
    ambient As Vector4f
    diffuse As Vector4f
    specular As Vector4f
    fogColour As Vector4f
    fogEnd As Single
    fogBegin As Single
    timeState As Single                 ' graphicsCycle
    fogEnabled As Long
End Type

' Only change per mesh
Public Type Draw3DShapePerMeshUniforms
    tcmask As Long
    normalMap As Long
    specularMap As Long
    hasTangents As Long
End Type

' Change per instance of mesh
Public Type Draw3DShapePerInstanceUniforms
    ' instead of passing the modelMatrix, pre-calculate the ModelViewMatrix and NormalMatrix CPU-side
    ModelViewMatrix As mat4
    NormalMatrix As mat4
    colour As Vector4f
    teamcolour As Vector4f
    shaderStretch As Single
    animFrameNumber As Single
    ecmState As Long
    alphaTest As Long
End Type

' Only change once per frame
Public Type Draw3DShapeInstancedGlobalUniforms
    ProjectionMatrix As mat4
    ViewMatrix As mat4
    ModelUVLightmapMatrix As mat4
    ShadowMapMVPMatrix(0 To WZ_MAX_SHADOW_CASCADES - 1) As mat4
    cameraPos As Vector4f               ' in modelSpace
    sunPos As Vector4f
    sceneColor As Vector4f
    ambient As Vector4f
    diffuse As Vector4f
    specular As Vector4f
    fogColour As Vector4f
    ShadowMapCascadeSplits As Vector4f  ' Can't use float[4] (because of std140 layout alignment rules, which don't match C/C++ and waste a lot of space)
    ShadowMapSize As Long
    fogEnd As Single
    fogBegin As Single
    timeState As Single                 ' graphicsCycle
    fogEnabled As Long
    viewportWidth As Long
    viewportheight As Long
    unused2 As Single
    PointLightsPosition(0 To gfx_api_max_lights - 1) As Vector4f
    PointLightsColorAndEnergy(0 To gfx_api_max_lights - 1) As Vector4f
    bucketOffsetAndSize(0 To gfx_api_bucket_dimension * gfx_api_bucket_dimension - 1) As Vector4i
    indexed_lights(0 To gfx_api_max_indexed_lights - 1) As Vector4i
    bucketDimensionUsed As Long
End Type

' Only change per mesh
Public Type Draw3DShapeInstancedPerMeshUniforms
    tcmask As Long
    normalMap As Long
    specularMap As Long
    hasTangents As Long
    shieldEffect As Long
End Type

' interleaved vertex data
Public Type Draw3DShapePerInstanceInterleavedData
    ModelViewMatrix As mat4             ' 16 bytes * 4 = 64 bytes
    ' glm::mat4 NormalMatrix; // can be calculated in the shader
    shaderStretch_ecmState_alphaTest_animFrameNumber As Vector4f ' 16 bytes
    colour As Long                      ' 4 bytes (stored as u8x4_norm)
    teamcolour As Long                  ' 4 bytes (stored as u8x4_norm)
    ' (+ 8 bytes padding for overall struct)
    unused_padding_(0 To 1) As Long
End Type

Public Type Draw3DShapeInstancedDepthOnlyGlobalUniforms
    ProjectionMatrix As mat4
    ViewMatrix As mat4
End Type

Public Type constant_buffer_type_SHADER_GENERIC_COLOR
    transform_matrix As mat4
    unused As Vector2f
    unused2 As Vector2f
    colour As Vector4f
End Type

Public Type constant_buffer_type_SHADER_TERRAIN_DEPTH
    transform_matrix As mat4
    paramX As Vector4f
    paramY As Vector4f
    paramXLight As Vector4f
    paramYLight As Vector4f
    unused As mat4
    texture_matrix As mat4
    fog_colour As Vector4f
    fog_enabled As Long
    fog_begin As Single
    fog_end As Single
    texture0 As Long
    texture1 As Long
End Type

Public Type constant_buffer_type_SHADER_TERRAIN_DEPTHMAP
    transform_matrix As mat4
    fog_enabled As Long
    fog_begin As Single
    fog_end As Single
End Type

Public Type TerrainDecalVertex
    pos As Vector3f                     ' = Vector3f(0.f, 0.f, 0.f)
    decalUv As Vector2f                 ' = Vector2f(0.f, 0.f)
    normal As Vector3f                  ' = Vector3f(0.f, 1.f, 0.f)
    decalTangent As Vector4f            ' = glm::vec4(1.f, 0.f, 0.f, 1.f); applicable only to decal texture!
    decalNo As Long                     ' = 0; tile #
    grounds As Long                     ' uint32_t: ground texture indexes for splatting. encoded as rgba bytes.
    groundWeights As Long               ' PIELIGHT: weights of corresponding ground textures. encoded as rgba floats
End Type

Public Type TerrainCombinedUniforms
    ModelViewProjectionMatrix As mat4
    ViewMatrix As mat4
    ModelUVLightmapMatrix As mat4
    ShadowMapMVPMatrix(0 To WZ_MAX_SHADOW_CASCADES - 1) As mat4
    groundScale As mat4                 ' array of scales for ground textures, encoded in mat4. scale_i = groundScale[i/4][i%4]
    cameraPos As Vector4f               ' in modelSpace
    sunPos As Vector4f                  ' in modelSpace
    emissiveLight As Vector4f           ' light colors/intensity
    ambientLight As Vector4f
    diffuseLight As Vector4f
    specularLight As Vector4f
    fog_colour As Vector4f
    ShadowMapCascadeSplits As Vector4f  ' Can't use float[4] (because of std140 layout alignment rules, which don't match C/C++ and waste a lot of space)
    ShadowMapSize As Long
    fog_enabled As Long
    fog_begin As Single
    fog_end As Single
    quality As Long
    viewportWidth As Long
    viewportheight As Long
    unused2 As Single
    PointLightsPosition(0 To gfx_api_max_lights - 1) As Vector4f
    PointLightsColorAndEnergy(0 To gfx_api_max_lights - 1) As Vector4f
    bucketOffsetAndSize(0 To gfx_api_bucket_dimension * gfx_api_bucket_dimension - 1) As Vector4i
    indexed_lights(0 To gfx_api_max_indexed_lights - 1) As Vector4i
    bucketDimensionUsed As Long
End Type

Public Type constant_buffer_type_SHADER_WATER
    ModelViewProjectionMatrix As mat4
    ModelUVLightmapMatrix As mat4
    ModelUV1Matrix As mat4
    ModelUV2Matrix As mat4
    cameraPos As Vector4f               ' in modelSpace
    sunPos As Vector4f                  ' in modelSpace
    emissiveLight As Vector4f           ' light colors/intensity
    ambientLight As Vector4f
    diffuseLight As Vector4f
    specularLight As Vector4f
    fog_colour As Vector4f
    fog_enabled As Long
    fog_begin As Single
    fog_end As Single
    timeSec As Single
End Type

Public Type constant_buffer_type_SHADER_WATER_HIGH
    ModelViewProjectionMatrix As mat4
    ViewMatrix As mat4
    ModelUVLightmapMatrix As mat4
    ShadowMapMVPMatrix(0 To WZ_MAX_SHADOW_CASCADES - 1) As mat4
    cameraPos As Vector4f               ' in modelSpace
    sunPos As Vector4f                  ' in modelSpace
    emissiveLight As Vector4f           ' light colors/intensity
    ambientLight As Vector4f
    diffuseLight As Vector4f
    specularLight As Vector4f
    fog_colour As Vector4f
    ShadowMapCascadeSplits As Vector4f  ' Can't use float[4] (because of std140 layout alignment rules, which don't match C/C++ and waste a lot of space)
    ShadowMapSize As Long
    fog_enabled As Long
    fog_begin As Single
    fog_end As Single
    timeSec As Single
End Type

Public Type constant_buffer_type_SHADER_WATER_CLASSIC
    ModelViewProjectionMatrix As mat4
    ModelUVLightmapMatrix As mat4
    ShadowMapMVPMatrix As mat4
    ModelUV1Matrix As mat4
    ModelUV2Matrix As mat4
    cameraPos As Vector4f               ' in modelSpace
    sunPos As Vector4f                  ' in modelSpace
    fog_colour As Vector4f
    fog_enabled As Long
    fog_begin As Single
    fog_end As Single
    timeSec As Single
End Type

Public Type constant_buffer_type_SHADER_GFX_TEXT
    transform_matrix As mat4
    offset As Vector2f
    size As Vector2f
    color As Vector4f
    texture As Long                     ' IGNORED
End Type

Public Type constant_buffer_type_SHADER_SKYBOX
    transform_matrix As mat4
    color As Vector4f
    fog_color As Vector4f
    fog_enabled As Long
End Type

Public Type constant_buffer_type_SHADER_GFX_COLOUR
    transform_matrix As mat4
End Type

Public Type constant_buffer_type_SHADER_TEXT
    transform_matrix As mat4
    offset As Vector2f
    size As Vector2f
    color As Vector4f
    texture As Long                     ' IGNORED
End Type

Public Type constant_buffer_type_SHADER_RECT
    transform_matrix As mat4
    offset As Vector2f
    size As Vector2f
    colour As Vector4f
End Type

Public Type constant_buffer_type_SHADER_TEXRECT
    transform_matrix As mat4
    offset As Vector2f
    size As Vector2f
    color As Vector4f
    texture As Long                     ' IGNORED
End Type

Public Type constant_buffer_type_SHADER_RECT_INSTANCED
    ProjectionMatrix As mat4
End Type

' interleaved vertex data
Public Type MultiRectPerInstanceInterleavedData
    TransformationMatrix As mat4        ' 16 bytes * 4 = 64 bytes
    offset_scale As Vector4f            ' 16 bytes (2 vec2)
    colour As Long                      ' 4 bytes (stored as u8x4_norm)
    ' (+ 12 bytes padding for overall struct)
    unused_padding_(0 To 2) As Long
End Type

Public Type constant_buffer_type_SHADER_LINE
    mat As mat4
    p0 As Vector2f
    p1 As Vector2f
    colour As Vector4f
End Type

Public Type constant_buffer_type_SHADER_DEBUG_TEXTURE2D_QUAD
    transform_matrix As mat4
    uv_transform_matrix As mat4
    swizzle As Vector4i
    color As Vector4f
    texture As Long
End Type

Public Type constant_buffer_type_SHADER_DEBUG_TEXTURE2DARRAY_QUAD
    transform_matrix As mat4
    uv_transform_matrix As mat4
    swizzle As Vector4i
    color As Vector4f
    layer As Long
    texture As Long
End Type

Public Type constant_buffer_type_SHADER_WORLD_TO_SCREEN
    gamma As Single
End Type
