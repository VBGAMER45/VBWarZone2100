Attribute VB_Name = "gfx_api_pso_c"
' lib/ivis_opengl/gfx_api.h: the pipeline descriptions (state_description,
' vertex_buffer, texture_input, pipeline_create_info, lighting_constants and
' their enums) and the pipeline_state_helper instantiations. C++ builds each
' instantiation's pipeline_create_info from its template arguments
' (rasterizer_state<...>, primitive, index, the uniform, vertex buffer and
' texture tuples, the shader); here pso_create_info(pso) builds it from the
' pso_type (GfxApi), and gfx_api_pso_get(pso) returns the instantiation's one
' object (its static get()).
Option Explicit

Public Enum polygon_offset
    polygon_offset_enabled
    polygon_offset_disabled
End Enum

Public Enum stencil_mode
    stencil_mode_stencil_shadow_silhouette
    stencil_mode_stencil_shadow_quad
    stencil_mode_stencil_disabled
End Enum

Public Enum cull_mode
    cull_mode_back
    cull_mode_front
    cull_mode_shadow_mapping
    cull_mode_none
End Enum

Public Enum vertex_attribute_type
    vertex_attribute_type_float2
    vertex_attribute_type_float3
    vertex_attribute_type_float4
    ' mat4,
    vertex_attribute_type_u8x4_norm
    ' not guaranteed to be supported on all systems:
    vertex_attribute_type_u8x4_uint
    vertex_attribute_type_int1
End Enum

Public Enum vertex_attribute_input_rate
    vertex_attribute_input_rate_vertex
    vertex_attribute_input_rate_instance
End Enum

Public Enum sampler_type
    sampler_type_bilinear
    sampler_type_bilinear_repeat
    sampler_type_bilinear_border
    sampler_type_anisotropic
    sampler_type_nearest_clamped
    sampler_type_nearest_border
    sampler_type_anisotropic_repeat
End Enum

Public Enum border_color
    border_color_none
    border_color_transparent_black
    border_color_opaque_black
    border_color_opaque_white
End Enum

' (VB) the uniform block types (C++ tells them apart by std::type_index): what
' pipeline_create_info.uniform_blocks lists and gl_pipeline_state_object's
' set_constants dispatches on
Public Enum uniform_block_type
    ub_Draw3DShapeGlobalUniforms
    ub_Draw3DShapePerMeshUniforms
    ub_Draw3DShapePerInstanceUniforms
    ub_Draw3DShapeInstancedGlobalUniforms
    ub_Draw3DShapeInstancedPerMeshUniforms
    ub_Draw3DShapeInstancedDepthOnlyGlobalUniforms
    ub_SHADER_TERRAIN_DEPTH
    ub_SHADER_TERRAIN_DEPTHMAP
    ub_TerrainCombinedUniforms
    ub_SHADER_WATER
    ub_SHADER_WATER_HIGH
    ub_SHADER_WATER_CLASSIC
    ub_SHADER_RECT
    ub_SHADER_TEXRECT
    ub_SHADER_GFX_COLOUR
    ub_SHADER_GFX_TEXT
    ub_SHADER_SKYBOX
    ub_SHADER_GENERIC_COLOR
    ub_SHADER_RECT_INSTANCED
    ub_SHADER_LINE
    ub_SHADER_TEXT
    ub_SHADER_DEBUG_TEXTURE2D_QUAD
    ub_SHADER_DEBUG_TEXTURE2DARRAY_QUAD
    ub_SHADER_WORLD_TO_SCREEN
End Enum

' the vertex attribute ids (gfx_api.h's constexpr std::size_t gfx_api::position, ...)
Public Const gfx_api_position As Long = 0
Public Const gfx_api_texcoord As Long = 1
Public Const gfx_api_color As Long = 2
Public Const gfx_api_normal As Long = 3
Public Const gfx_api_tangent As Long = 4
' for new terrain shader
Public Const gfx_api_terrain_tileNo As Long = 5
Public Const gfx_api_terrain_grounds As Long = 6
Public Const gfx_api_terrain_groundWeights As Long = 7
' for instanced model rendering
Public Const gfx_api_instance_modelMatrix As Long = 5
Public Const gfx_api_instance_packedValues As Long = 9
Public Const gfx_api_instance_Colour As Long = 10
Public Const gfx_api_instance_TeamColour As Long = 11

Public Type state_description
    blend_state As REND_MODE
    depth_mode As DEPTH_MODE
    output_mask As Long                 ' uint8_t
    offset As Boolean                   ' _polygon_offset == polygon_offset::enabled
    stencil As stencil_mode
    cull As cull_mode
End Type

Public Type vertex_buffer_input
    id As Long
    type_ As vertex_attribute_type
    offset As Long
End Type

Public Type vertex_buffer
    stride As Long
    rate As vertex_attribute_input_rate
    attributes() As vertex_buffer_input
    attributes_size As Long             ' (VB) attributes.size()
End Type

Public Type texture_input
    id As Long
    sampler As sampler_type
    target As pixel_format_target
    border As border_color
End Type

Public Type pipeline_create_info
    state_desc As state_description
    shader_mode As SHADER_MODE
    primitive As primitive_type
    uniform_blocks() As uniform_block_type
    uniform_blocks_size As Long
    texture_desc() As texture_input
    texture_desc_size As Long
    attribute_descriptions() As vertex_buffer
    attribute_descriptions_size As Long
    index As index_type                 ' (VB) the instantiation's index_type (a template argument)
End Type

Public Type lighting_constants
    shadowMode As Long                  ' uint32_t = 1
    shadowFilterSize As Long            ' uint32_t = 5
    shadowCascadesCount As Long         ' uint32_t = WZ_MAX_SHADOW_CASCADES
    isPointLightPerPixelEnabled As Boolean ' = false
End Type

' gfx_api::X::get()'s function-local statics: one pipeline_state_helper per pso_type, made on first use
Private psoHelpers() As pipeline_state_helper
Private psoHelpersReady As Boolean

' the pipeline_create_info that pso_create_info is writing
Private ci As pipeline_create_info

' gfx_api::lighting_constants() (the member initialisers)
Public Function lighting_constants_new() As lighting_constants
    lighting_constants_new.shadowMode = 1
    lighting_constants_new.shadowFilterSize = 5
    lighting_constants_new.shadowCascadesCount = WZ_MAX_SHADOW_CASCADES
    lighting_constants_new.isPointLightPerPixelEnabled = False
End Function

' lighting_constants::operator==
Public Function lighting_constants_equal(lhs As lighting_constants, rhs As lighting_constants) As Boolean
    lighting_constants_equal = lhs.shadowMode = rhs.shadowMode _
        And lhs.shadowFilterSize = rhs.shadowFilterSize _
        And lhs.shadowCascadesCount = rhs.shadowCascadesCount _
        And lhs.isPointLightPerPixelEnabled = rhs.isPointLightPerPixelEnabled
End Function

' gfx_api::X::get() (pipeline_state_helper<...>::get(): static pipeline_state_helper object; return object;)
Public Function gfx_api_pso_get(ByVal pso As pso_type) As pipeline_state_helper
    If Not psoHelpersReady Then
        ReDim psoHelpers(0 To pso_type_count - 1)
        psoHelpersReady = True
    End If
    If psoHelpers(pso) Is Nothing Then
        Set psoHelpers(pso) = New pipeline_state_helper
        psoHelpers(pso).init pso
    End If
    Set gfx_api_pso_get = psoHelpers(pso)
End Function

'--------------------------------------------------------------------------
' The template arguments, written out: rasterizer_state<...>::get(), the
' primitive and index types, the uniform, vertex buffer and texture tuples
' (untuple_typeinfo, untuple<vertex_buffer>, untuple<texture_input>) and the
' shader. Each helper appends to ci what the template of its name describes.

' rasterizer_state<render_mode, depth_mode, output_mask, offset, stencil, cull>::get()
Private Sub rasterizer_state(ByVal render_mode As REND_MODE, ByVal depth_mode_ As DEPTH_MODE, ByVal output_mask As Long, ByVal offset As polygon_offset, ByVal stencil As stencil_mode, ByVal cull As cull_mode)
    ci.state_desc.blend_state = render_mode
    ci.state_desc.depth_mode = depth_mode_
    ci.state_desc.output_mask = output_mask
    ci.state_desc.offset = (offset = polygon_offset_enabled)
    ci.state_desc.stencil = stencil
    ci.state_desc.cull = cull
End Sub

' one type of the uniform_inputs tuple
Private Sub uniform_input(ByVal block As uniform_block_type)
    ReDim Preserve ci.uniform_blocks(0 To ci.uniform_blocks_size)
    ci.uniform_blocks(ci.uniform_blocks_size) = block
    ci.uniform_blocks_size = ci.uniform_blocks_size + 1
End Sub

' vertex_buffer_description<stride, rate, ...>: starts a vertex buffer; its
' vertex_attribute_descriptions follow
Private Sub vertex_buffer_description(ByVal stride As Long, ByVal rate As vertex_attribute_input_rate)
    ReDim Preserve ci.attribute_descriptions(0 To ci.attribute_descriptions_size)
    ci.attribute_descriptions(ci.attribute_descriptions_size).stride = stride
    ci.attribute_descriptions(ci.attribute_descriptions_size).rate = rate
    ci.attribute_descriptions(ci.attribute_descriptions_size).attributes_size = 0
    ci.attribute_descriptions_size = ci.attribute_descriptions_size + 1
End Sub

' vertex_attribute_description<id, type, offset> (of the last vertex buffer)
Private Sub vertex_attribute_description(ByVal id As Long, ByVal type_ As vertex_attribute_type, ByVal offset As Long)
    Dim b As Long, n As Long
    b = ci.attribute_descriptions_size - 1
    n = ci.attribute_descriptions(b).attributes_size
    ReDim Preserve ci.attribute_descriptions(b).attributes(0 To n)
    ci.attribute_descriptions(b).attributes(n).id = id
    ci.attribute_descriptions(b).attributes(n).type_ = type_
    ci.attribute_descriptions(b).attributes(n).offset = offset
    ci.attribute_descriptions(b).attributes_size = n + 1
End Sub

' texture_description<texture_unit, sampler, target = texture_2d, border = none>
Private Sub texture_description(ByVal texture_unit As Long, ByVal sampler As sampler_type, Optional ByVal target As pixel_format_target = pixel_format_target_texture_2d, Optional ByVal border As border_color = border_color_none)
    ReDim Preserve ci.texture_desc(0 To ci.texture_desc_size)
    ci.texture_desc(ci.texture_desc_size).id = texture_unit
    ci.texture_desc(ci.texture_desc_size).sampler = sampler
    ci.texture_desc(ci.texture_desc_size).target = target
    ci.texture_desc(ci.texture_desc_size).border = border
    ci.texture_desc_size = ci.texture_desc_size + 1
End Sub

' the remaining template arguments: primitive, index, shader
Private Sub pipeline_args(ByVal primitive As primitive_type, ByVal index As index_type, ByVal shader As SHADER_MODE)
    ci.primitive = primitive
    ci.index = index
    ci.shader_mode = shader
End Sub

' the instance data vertex buffer of Draw3DShapeInstanced and Draw3DShapeDepthOnly_Instanced
Private Sub Draw3DShapePerInstanceInterleavedData_description()
    Dim d As Draw3DShapePerInstanceInterleavedData, v4 As Vector4f
    vertex_buffer_description LenB(d), vertex_attribute_input_rate_instance
    vertex_attribute_description gfx_api_instance_modelMatrix, vertex_attribute_type_float4, 0
    vertex_attribute_description gfx_api_instance_modelMatrix + 1, vertex_attribute_type_float4, LenB(v4)
    vertex_attribute_description gfx_api_instance_modelMatrix + 2, vertex_attribute_type_float4, LenB(v4) * 2
    vertex_attribute_description gfx_api_instance_modelMatrix + 3, vertex_attribute_type_float4, LenB(v4) * 3
    vertex_attribute_description gfx_api_instance_packedValues, vertex_attribute_type_float4, VarPtr(d.shaderStretch_ecmState_alphaTest_animFrameNumber) - VarPtr(d)
    vertex_attribute_description gfx_api_instance_Colour, vertex_attribute_type_u8x4_norm, VarPtr(d.colour) - VarPtr(d)
    vertex_attribute_description gfx_api_instance_TeamColour, vertex_attribute_type_u8x4_norm, VarPtr(d.teamcolour) - VarPtr(d)
End Sub

' template<REND_MODE render_mode, SHADER_MODE shader, DEPTH_MODE depth_mode> using Draw3DShape
Private Sub Draw3DShape_template(ByVal render_mode As REND_MODE, ByVal shader As SHADER_MODE, ByVal depth_mode_ As DEPTH_MODE)
    rasterizer_state render_mode, depth_mode_, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
    uniform_input ub_Draw3DShapeGlobalUniforms
    uniform_input ub_Draw3DShapePerMeshUniforms
    uniform_input ub_Draw3DShapePerInstanceUniforms
    vertex_buffer_description 12, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_position, vertex_attribute_type_float3, 0
    vertex_buffer_description 12, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_normal, vertex_attribute_type_float3, 0
    vertex_buffer_description 16, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_texcoord, vertex_attribute_type_float4, 0
    vertex_buffer_description 16, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_tangent, vertex_attribute_type_float4, 0
    texture_description 0, sampler_type_anisotropic   ' diffuse
    texture_description 1, sampler_type_bilinear      ' team color mask
    texture_description 2, sampler_type_anisotropic   ' normal map
    texture_description 3, sampler_type_anisotropic   ' specular map
    pipeline_args primitive_type_triangles, index_type_u16, shader
End Sub

' template<REND_MODE render_mode, SHADER_MODE shader, DEPTH_MODE depth_mode> using Draw3DShapeInstanced
Private Sub Draw3DShapeInstanced_template(ByVal render_mode As REND_MODE, ByVal shader As SHADER_MODE, ByVal depth_mode_ As DEPTH_MODE)
    rasterizer_state render_mode, depth_mode_, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
    uniform_input ub_Draw3DShapeInstancedGlobalUniforms
    uniform_input ub_Draw3DShapeInstancedPerMeshUniforms
    vertex_buffer_description 12, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_position, vertex_attribute_type_float3, 0
    vertex_buffer_description 12, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_normal, vertex_attribute_type_float3, 0
    vertex_buffer_description 16, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_texcoord, vertex_attribute_type_float4, 0
    vertex_buffer_description 16, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_tangent, vertex_attribute_type_float4, 0
    ' instance data
    Draw3DShapePerInstanceInterleavedData_description
    texture_description 0, sampler_type_anisotropic   ' diffuse
    texture_description 1, sampler_type_bilinear      ' team color mask
    texture_description 2, sampler_type_anisotropic   ' normal map
    texture_description 3, sampler_type_anisotropic   ' specular map
    texture_description 4, sampler_type_bilinear_border, pixel_format_target_depth_map, border_color_opaque_white ' depth / shadow map
    texture_description 5, sampler_type_bilinear      ' lightmap
    pipeline_args primitive_type_triangles, index_type_u16, shader
End Sub

' using TerrainVertexVBODescription
Private Sub TerrainVertexVBODescription()
    Dim v3 As Vector3f
    vertex_buffer_description LenB(v3), vertex_attribute_input_rate_vertex   ' sizeof(glm::vec3) /*+ sizeof(glm::vec3)*/
    vertex_attribute_description gfx_api_position, vertex_attribute_type_float3, 0
End Sub

' template<REND_MODE render_mode, SHADER_MODE shader> using TerrainCombinedTemplate
Private Sub TerrainCombinedTemplate(ByVal render_mode As REND_MODE, ByVal shader As SHADER_MODE)
    Dim v As TerrainDecalVertex
    rasterizer_state render_mode, DEPTH_CMP_LEQ_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
    uniform_input ub_TerrainCombinedUniforms
    vertex_buffer_description LenB(v), vertex_attribute_input_rate_vertex  ' TerrainDecalVertex struct
    vertex_attribute_description gfx_api_position, vertex_attribute_type_float3, 0
    vertex_attribute_description gfx_api_texcoord, vertex_attribute_type_float2, VarPtr(v.decalUv) - VarPtr(v)
    vertex_attribute_description gfx_api_normal, vertex_attribute_type_float3, VarPtr(v.normal) - VarPtr(v)
    vertex_attribute_description gfx_api_tangent, vertex_attribute_type_float4, VarPtr(v.decalTangent) - VarPtr(v)
    vertex_attribute_description gfx_api_terrain_tileNo, vertex_attribute_type_int1, VarPtr(v.decalNo) - VarPtr(v)
    vertex_attribute_description gfx_api_terrain_grounds, vertex_attribute_type_u8x4_uint, VarPtr(v.grounds) - VarPtr(v)
    vertex_attribute_description gfx_api_terrain_groundWeights, vertex_attribute_type_u8x4_norm, VarPtr(v.groundWeights) - VarPtr(v)
    texture_description 0, sampler_type_bilinear      ' lightmap
    texture_description 1, sampler_type_anisotropic_repeat, pixel_format_target_texture_2d_array ' ground
    texture_description 2, sampler_type_anisotropic_repeat, pixel_format_target_texture_2d_array ' ground normal
    texture_description 3, sampler_type_anisotropic_repeat, pixel_format_target_texture_2d_array ' ground specular
    texture_description 4, sampler_type_anisotropic_repeat, pixel_format_target_texture_2d_array ' ground height
    texture_description 5, sampler_type_anisotropic, pixel_format_target_texture_2d_array ' decal
    texture_description 6, sampler_type_anisotropic, pixel_format_target_texture_2d_array ' decal normal
    texture_description 7, sampler_type_anisotropic, pixel_format_target_texture_2d_array ' decal specular
    texture_description 8, sampler_type_anisotropic, pixel_format_target_texture_2d_array ' decal height
    texture_description 9, sampler_type_bilinear_border, pixel_format_target_depth_map, border_color_opaque_white ' depth / shadow map
    pipeline_args primitive_type_triangles, index_type_u16, shader
End Sub

' the water PSOs' one vertex buffer: WaterVertex, w is depth
Private Sub WaterVertex_description()
    vertex_buffer_description 16, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_position, vertex_attribute_type_float4, 0
End Sub

' using gfx_tc, gfx_colour (gfx_colour_: GFXTYPE's GFX_COLOUR differs only in case), gfx_vtx2, gfx_vtx3
Private Sub gfx_tc()
    vertex_buffer_description 8, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_texcoord, vertex_attribute_type_float2, 0
End Sub

Private Sub gfx_colour_()
    vertex_buffer_description 4, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_color, vertex_attribute_type_u8x4_norm, 0
End Sub

Private Sub gfx_vtx2()
    vertex_buffer_description 8, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_position, vertex_attribute_type_float2, 0
End Sub

Private Sub gfx_vtx3()
    vertex_buffer_description 12, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_position, vertex_attribute_type_float3, 0
End Sub

' template<REND_MODE rm, DEPTH_MODE dm, primitive_type primitive, typename VTX, typename Second, SHADER_MODE shader, typename texture> using GFX
' (VTX is gfx_vtx2 in every use; secondColour: Second is gfx_colour, else gfx_tc; the texture tuple follows the call)
Private Sub GFX_template(ByVal rm As REND_MODE, ByVal dm As DEPTH_MODE, ByVal primitive As primitive_type, ByVal secondColour As Boolean, ByVal shader As SHADER_MODE, ByVal block As uniform_block_type)
    rasterizer_state rm, dm, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
    uniform_input block                                ' constant_buffer_type<shader>
    gfx_vtx2
    If secondColour Then
        gfx_colour_
    Else
        gfx_tc
    End If
    pipeline_args primitive, index_type_u16, shader
End Sub

' the 2D PSOs' one vertex buffer: position as u8x4_norm
Private Sub u8x4_position_description()
    vertex_buffer_description 4, vertex_attribute_input_rate_vertex
    vertex_attribute_description gfx_api_position, vertex_attribute_type_u8x4_norm, 0
End Sub

' (VB) the pipeline_create_info of the instantiation pso (its template arguments)
Public Function pso_create_info(ByVal pso As pso_type) As pipeline_create_info
    Dim blank As pipeline_create_info, m As MultiRectPerInstanceInterleavedData, v4 As Vector4f
    ci = blank
    Select Case pso
    Case Draw3DShapeOpaque: Draw3DShape_template REND_OPAQUE, SHADER_COMPONENT, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeAlpha: Draw3DShape_template REND_ALPHA, SHADER_COMPONENT, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapePremul: Draw3DShape_template REND_PREMULTIPLIED, SHADER_COMPONENT, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeAdditive: Draw3DShape_template REND_ADDITIVE, SHADER_COMPONENT, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeNoLightOpaque: Draw3DShape_template REND_OPAQUE, SHADER_NOLIGHT, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeNoLightAlpha: Draw3DShape_template REND_ALPHA, SHADER_NOLIGHT, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeNoLightPremul: Draw3DShape_template REND_PREMULTIPLIED, SHADER_NOLIGHT, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeNoLightAdditive: Draw3DShape_template REND_ADDITIVE, SHADER_NOLIGHT, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeAlphaNoDepthWRT: Draw3DShape_template REND_ALPHA, SHADER_COMPONENT, DEPTH_CMP_LEQ_WRT_OFF
    Case Draw3DShapeNoLightAlphaNoDepthWRT: Draw3DShape_template REND_ALPHA, SHADER_NOLIGHT, DEPTH_CMP_LEQ_WRT_OFF

    Case Draw3DShapeOpaque_Instanced: Draw3DShapeInstanced_template REND_OPAQUE, SHADER_COMPONENT_INSTANCED, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeAlpha_Instanced: Draw3DShapeInstanced_template REND_ALPHA, SHADER_COMPONENT_INSTANCED, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapePremul_Instanced: Draw3DShapeInstanced_template REND_PREMULTIPLIED, SHADER_COMPONENT_INSTANCED, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeAdditive_Instanced: Draw3DShapeInstanced_template REND_ADDITIVE, SHADER_COMPONENT_INSTANCED, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeNoLightOpaque_Instanced: Draw3DShapeInstanced_template REND_OPAQUE, SHADER_NOLIGHT_INSTANCED, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeNoLightAlpha_Instanced: Draw3DShapeInstanced_template REND_ALPHA, SHADER_NOLIGHT_INSTANCED, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeNoLightPremul_Instanced: Draw3DShapeInstanced_template REND_PREMULTIPLIED, SHADER_NOLIGHT_INSTANCED, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeNoLightAdditive_Instanced: Draw3DShapeInstanced_template REND_ADDITIVE, SHADER_NOLIGHT_INSTANCED, DEPTH_CMP_LEQ_WRT_ON
    Case Draw3DShapeAlphaNoDepthWRT_Instanced: Draw3DShapeInstanced_template REND_ALPHA, SHADER_COMPONENT_INSTANCED, DEPTH_CMP_LEQ_WRT_OFF
    Case Draw3DShapeNoLightAlphaNoDepthWRT_Instanced: Draw3DShapeInstanced_template REND_ALPHA, SHADER_NOLIGHT_INSTANCED, DEPTH_CMP_LEQ_WRT_OFF
    Case Draw3DShapeAdditiveNoDepthWRT_Instanced: Draw3DShapeInstanced_template REND_ADDITIVE, SHADER_COMPONENT_INSTANCED, DEPTH_CMP_LEQ_WRT_OFF
    Case Draw3DShapeNoLightAdditiveNoDepthWRT_Instanced: Draw3DShapeInstanced_template REND_ADDITIVE, SHADER_NOLIGHT_INSTANCED, DEPTH_CMP_LEQ_WRT_OFF
    Case Draw3DShapePremulNoDepthWRT_Instanced: Draw3DShapeInstanced_template REND_PREMULTIPLIED, SHADER_COMPONENT_INSTANCED, DEPTH_CMP_LEQ_WRT_OFF
    Case Draw3DShapeNoLightPremulNoDepthWRT_Instanced: Draw3DShapeInstanced_template REND_PREMULTIPLIED, SHADER_NOLIGHT_INSTANCED, DEPTH_CMP_LEQ_WRT_OFF

    Case Draw3DShapeDepthOnly_Instanced
        rasterizer_state REND_OPAQUE, DEPTH_CMP_LEQ_WRT_ON, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_shadow_mapping
        uniform_input ub_Draw3DShapeInstancedDepthOnlyGlobalUniforms
        vertex_buffer_description 12, vertex_attribute_input_rate_vertex
        vertex_attribute_description gfx_api_position, vertex_attribute_type_float3, 0
        vertex_buffer_description 12, vertex_attribute_input_rate_vertex
        vertex_attribute_description gfx_api_normal, vertex_attribute_type_float3, 0
        ' instance data
        Draw3DShapePerInstanceInterleavedData_description
        ' notexture
        pipeline_args primitive_type_triangles, index_type_u16, SHADER_COMPONENT_DEPTH_INSTANCED

    Case TransColouredTrianglePSO
        rasterizer_state REND_ADDITIVE, DEPTH_CMP_LEQ_WRT_ON, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_GENERIC_COLOR
        vertex_buffer_description 12, vertex_attribute_input_rate_vertex
        vertex_attribute_description gfx_api_position, vertex_attribute_type_float3, 0
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_GENERIC_COLOR

    Case DrawStencilShadow
        rasterizer_state REND_OPAQUE, DEPTH_CMP_LEQ_WRT_OFF, 0, polygon_offset_disabled, stencil_mode_stencil_shadow_silhouette, cull_mode_none
        uniform_input ub_SHADER_GENERIC_COLOR
        vertex_buffer_description 12, vertex_attribute_input_rate_vertex
        vertex_attribute_description gfx_api_position, vertex_attribute_type_float3, 0
        pipeline_args primitive_type_triangles, index_type_u16, SHADER_GENERIC_COLOR

    Case TerrainDepth
        rasterizer_state REND_OPAQUE, DEPTH_CMP_LEQ_WRT_ON, 0, polygon_offset_enabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_TERRAIN_DEPTH
        TerrainVertexVBODescription
        texture_description 0, sampler_type_bilinear_repeat
        pipeline_args primitive_type_triangles, index_type_u32, SHADER_TERRAIN_DEPTH

    Case TerrainDepthOnlyForDepthMap
        rasterizer_state REND_OPAQUE, DEPTH_CMP_LEQ_WRT_ON, 0, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_none
        uniform_input ub_SHADER_TERRAIN_DEPTHMAP
        TerrainVertexVBODescription
        texture_description 0, sampler_type_bilinear_repeat
        pipeline_args primitive_type_triangles, index_type_u32, SHADER_TERRAIN_DEPTHMAP

    Case TerrainCombined_Classic: TerrainCombinedTemplate REND_ALPHA, SHADER_TERRAIN_COMBINED_CLASSIC
    Case TerrainCombined_Medium: TerrainCombinedTemplate REND_OPAQUE, SHADER_TERRAIN_COMBINED_MEDIUM
    Case TerrainCombined_High: TerrainCombinedTemplate REND_OPAQUE, SHADER_TERRAIN_COMBINED_HIGH

    Case WaterPSO
        rasterizer_state REND_MULTIPLICATIVE, DEPTH_CMP_LEQ_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_WATER
        WaterVertex_description
        texture_description 0, sampler_type_anisotropic_repeat ' tex1
        texture_description 1, sampler_type_anisotropic_repeat ' tex2
        texture_description 2, sampler_type_bilinear      ' lightmap
        pipeline_args primitive_type_triangles, index_type_u32, SHADER_WATER

    Case WaterHighPSO
        rasterizer_state REND_ALPHA, DEPTH_CMP_LEQ_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_WATER_HIGH
        WaterVertex_description
        texture_description 0, sampler_type_anisotropic_repeat, pixel_format_target_texture_2d_array ' textures
        texture_description 1, sampler_type_anisotropic_repeat, pixel_format_target_texture_2d_array ' normal maps
        texture_description 2, sampler_type_anisotropic_repeat, pixel_format_target_texture_2d_array ' specular maps
        texture_description 3, sampler_type_bilinear      ' lightmap
        texture_description 4, sampler_type_bilinear_border, pixel_format_target_depth_map, border_color_opaque_white ' depth / shadow map
        pipeline_args primitive_type_triangles, index_type_u32, SHADER_WATER_HIGH

    Case WaterClassicPSO
        rasterizer_state REND_OPAQUE, DEPTH_CMP_LEQ_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_WATER_CLASSIC
        WaterVertex_description
        texture_description 0, sampler_type_bilinear      ' lightmap
        texture_description 1, sampler_type_anisotropic_repeat, pixel_format_target_texture_2d ' water decal texture
        pipeline_args primitive_type_triangles, index_type_u32, SHADER_WATER_CLASSIC

    Case VideoPSO
        GFX_template REND_OPAQUE, DEPTH_CMP_ALWAYS_WRT_OFF, primitive_type_triangle_strip, False, SHADER_GFX_TEXT, ub_SHADER_GFX_TEXT
        texture_description 0, sampler_type_bilinear
    Case BackDropPSO
        GFX_template REND_OPAQUE, DEPTH_CMP_ALWAYS_WRT_OFF, primitive_type_triangle_strip, False, SHADER_GFX_TEXT, ub_SHADER_GFX_TEXT
        texture_description 0, sampler_type_nearest_clamped
    Case SkyboxPSO
        rasterizer_state REND_ALPHA, DEPTH_CMP_LEQ_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_SKYBOX
        gfx_vtx3
        gfx_tc
        texture_description 0, sampler_type_bilinear_repeat
        pipeline_args primitive_type_triangles, index_type_u16, SHADER_SKYBOX
    Case RadarPSO
        GFX_template REND_ALPHA, DEPTH_CMP_ALWAYS_WRT_OFF, primitive_type_triangle_strip, False, SHADER_GFX_TEXT, ub_SHADER_GFX_TEXT
        texture_description 0, sampler_type_nearest_clamped
    Case RadarViewInsideFillPSO
        GFX_template REND_ALPHA, DEPTH_CMP_ALWAYS_WRT_OFF, primitive_type_triangle_strip, True, SHADER_GFX_COLOUR, ub_SHADER_GFX_COLOUR
    Case RadarViewOutlinePSO
        GFX_template REND_ALPHA, DEPTH_CMP_ALWAYS_WRT_OFF, primitive_type_line_strip, True, SHADER_GFX_COLOUR, ub_SHADER_GFX_COLOUR

    Case DrawImageTextPSO
        rasterizer_state REND_TEXT, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_none
        uniform_input ub_SHADER_TEXT
        u8x4_position_description
        texture_description 0, sampler_type_bilinear
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_TEXT
    Case ShadowBox2DPSO
        rasterizer_state REND_OPAQUE, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_RECT
        u8x4_position_description
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_RECT
    Case UniTransBoxPSO
        rasterizer_state REND_ALPHA, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_RECT
        u8x4_position_description
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_RECT
    Case DrawImagePSO
        rasterizer_state REND_ALPHA, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_TEXRECT
        u8x4_position_description
        texture_description 0, sampler_type_bilinear
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_TEXRECT
    Case DrawImageAnisotropicPSO
        rasterizer_state REND_ALPHA, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_TEXRECT
        u8x4_position_description
        texture_description 0, sampler_type_anisotropic
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_TEXRECT
    Case BoxFillPSO
        rasterizer_state REND_OPAQUE, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_RECT
        u8x4_position_description
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_RECT
    Case BoxFillAlphaPSO
        rasterizer_state REND_ALPHA, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_shadow_quad, cull_mode_back
        uniform_input ub_SHADER_RECT
        u8x4_position_description
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_RECT
    Case BoxFillPSO_Instanced
        rasterizer_state REND_OPAQUE, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_RECT_INSTANCED
        u8x4_position_description
        ' instance data
        vertex_buffer_description LenB(m), vertex_attribute_input_rate_instance
        vertex_attribute_description gfx_api_instance_modelMatrix, vertex_attribute_type_float4, 0
        vertex_attribute_description gfx_api_instance_modelMatrix + 1, vertex_attribute_type_float4, LenB(v4)
        vertex_attribute_description gfx_api_instance_modelMatrix + 2, vertex_attribute_type_float4, LenB(v4) * 2
        vertex_attribute_description gfx_api_instance_modelMatrix + 3, vertex_attribute_type_float4, LenB(v4) * 3
        vertex_attribute_description gfx_api_instance_packedValues, vertex_attribute_type_float4, VarPtr(m.offset_scale) - VarPtr(m)
        vertex_attribute_description gfx_api_instance_Colour, vertex_attribute_type_u8x4_norm, VarPtr(m.colour) - VarPtr(m)
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_RECT_INSTANCED
    Case LinePSO
        rasterizer_state REND_ALPHA, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_back
        uniform_input ub_SHADER_LINE
        u8x4_position_description
        pipeline_args primitive_type_lines, index_type_u16, SHADER_LINE
    Case DebugDrawTexture2DToQuad
        rasterizer_state REND_TEXT, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_none
        uniform_input ub_SHADER_DEBUG_TEXTURE2D_QUAD
        u8x4_position_description
        texture_description 0, sampler_type_bilinear
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_DEBUG_TEXTURE2D_QUAD
    Case DebugDrawTexture2DArrayToQuad
        rasterizer_state REND_TEXT, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_none
        uniform_input ub_SHADER_DEBUG_TEXTURE2DARRAY_QUAD
        u8x4_position_description
        texture_description 0, sampler_type_bilinear, pixel_format_target_texture_2d_array
        pipeline_args primitive_type_triangle_strip, index_type_u16, SHADER_DEBUG_TEXTURE2DARRAY_QUAD
    Case WorldToScreenPSO
        rasterizer_state REND_OPAQUE, DEPTH_CMP_ALWAYS_WRT_OFF, 255, polygon_offset_disabled, stencil_mode_stencil_disabled, cull_mode_none
        uniform_input ub_SHADER_WORLD_TO_SCREEN
        vertex_buffer_description 2 * 4, vertex_attribute_input_rate_vertex  ' 2 * sizeof(gfxFloat)
        vertex_attribute_description gfx_api_position, vertex_attribute_type_float2, 0
        texture_description 0, sampler_type_bilinear, pixel_format_target_texture_2d
        pipeline_args primitive_type_triangles, index_type_u16, SHADER_WORLD_TO_SCREEN
    End Select
    pso_create_info = ci
End Function
