Attribute VB_Name = "piedraw_c"
' lib/ivis_opengl/piedraw.cpp, piedraw.h, piedef.h: render routines for 3D
' coloured and shaded transparency rendering: the lighting values, the shadow
' modes, the model draw queue and its instanced and old (non-instanced)
' drawing, and the fallback stencil shadows. A shape is an iIMDShape handle
' (imdload_c's imdShapes; its textures iIMDShape_getTextures). There is one
' InstancedMeshRenderer (instancedMeshRenderer): its members are this
' module's imr_* variables and its methods InstancedMeshRenderer_<name>. The
' templates over PSO types take pso_type values (which PSO a branch binds);
' ShaderOnce's set of types is a Boolean per pso_type. The std::unordered_map
' and InsertionOrderMap of ShapeVectors are ShapeVectorMaps: the keys in
' insertion order, each key's shapes chained in push order (C++ iterates the
' unordered_map in hash order: the opaque models' draw order differs, which
' the depth test hides). The memory pool and its allocator are not needed.
' The ShadowCache's two maps are one Scripting.Dictionary keyed by the shape
' and the draw parameters (the light's float bits). A uint64_t frame count
' is a Double.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const SHADOW_END_DISTANCE As Long = 64000000 ' (8000*8000) Keep in sync with lighting.c:FOG_END
Private Const pie_SHIELD_FACTOR As Single = 1.125!  ' pietypes.h

' InstancedMeshRenderer::DrawParts
Private Const DrawParts_ShadowCastingShapes As Long = &H1
Private Const DrawParts_TranslucentShapes As Long = &H2
Private Const DrawParts_AdditiveShapes As Long = &H4
Private Const DrawParts_OldShadows As Long = &H8
Private Const DrawParts_All As Long = DrawParts_ShadowCastingShapes Or DrawParts_TranslucentShapes Or DrawParts_AdditiveShapes

' enum DrawShadowResult
Private Const DRAW_SUCCESS_CACHED As Long = 0
Private Const DRAW_SUCCESS_UNCACHED As Long = 1

Private Type ShadowcastingShape
    modelViewMatrix As mat4
    shape As Long                       ' const iIMDShape *
    flag As Long
    flag_data As Long
    light_ As Vector4f
End Type

' struct SHAPE (SHAPE_: VB would read the many variables named shape as the type)
Private Type SHAPE_
    modelMatrix As mat4
    shape As Long                       ' const iIMDShape *
    frame As Long
    colour As Long                      ' PIELIGHT
    teamcolour As Long                  ' PIELIGHT
    flag As Long
    flag_data As Long
    stretch As Single
End Type

Private Type templatedState
    shader As SHADER_MODE               ' = SHADER_NONE
    shape As Long                       ' const iIMDShape * = nullptr
    pieFlag As Long                     ' = 0
End Type

' Ensure the uniforms are set for a shader once until reset
Private Type ShaderOnce
    performed_once(0 To pso_type_count - 1) As Boolean ' std::unordered_set<std::type_index>
End Type

' (VB) std::unordered_map<templatedState, ShapeVector> / InsertionOrderMap<templatedState, ShapeVector>
Private Type ShapeVectorMap
    keys() As templatedState            ' in insertion order
    firstShape() As Long                ' each key's shapes: a chain through nextShape (-1 ends it)
    lastShape() As Long
    counts() As Long                    ' ShapeVector::size()
    keys_size As Long
    shapes() As SHAPE_
    nextShape() As Long
    shapes_size As Long
    shapes_cap As Long                  ' (VB) the shapes arrays' capacity
    slots() As Long                     ' (VB) the hash table: key index + 1 (0 = empty slot)
    slotGen() As Long                   ' a slot is used when its generation is the map's
    gen As Long
    slotMask As Long
End Type

Private Type InstancedDrawCall
    state As templatedState
    instance_count As Long
    startingIdxInInstancesBuffer As Long ' = 0
End Type

Private Type CachedShadowData
    lastQueriedFrameCount As Double     ' uint64_t = 0
    vertexes() As Vector3f
    vertexes_size As Long
    key As String                       ' (VB) its key in shapeMap
End Type

Private Type ShadowCache
    currentFrame_ As Double             ' uint64_t _currentFrame = 0
    shapeMap As Object                  ' (VB) the shape and draw parameters -> index in entries
    entries() As CachedShadowData
    entries_size As Long
    freeEntries() As Long               ' (VB) the unused entries
    freeEntries_size As Long
    freeEntries_cap As Long
    vertexes() As Vector3f              ' std::vector<Vector3f> vertexes
    vertexes_size As Long
End Type

'
'  Local Variables
'

Private pieCount As Long                ' size_t = 0
Private polyCount As Long
Private drawCallsCount As Long
Private shadows As Boolean              ' = false
Private shadowsHasBeenInit As Boolean   ' = false
Private shadowMode_ As ShadowMode       ' = ShadowMode::Shadow_Mapping (shadowMode_: VB's names ignore case)
Private numShadowCascades As Long       ' uint32_t = WZ_MAX_SHADOW_CASCADES
Private lighting0(0 To LIGHT_MAX - 1, 0 To 3) As Single
Private lightingDefault(0 To LIGHT_MAX - 1, 0 To 3) As Single

Private currentSunPosition As Vector3f  ' (0.f, 0.f, 0.f)

Private scshapes() As ShadowcastingShape
Private scshapes_size As Long
Private scshapes_cap As Long
Private pZeroedVertexBuffer As gfx_api_buffer ' = nullptr
Private zeroedVertexBufferCurrentSize As Long ' getZeroedVertexBuffer's static size_t currentSize = 0

' pie_DrawShadow's statics, to save allocations
Private edgelist() As EDGE
Private edgelist_size As Long
Private edgelist_cap As Long
Private edgelistFlipped() As EDGE
Private edgelistFlipped_size As Long
Private edgelistFiltered() As EDGE
Private edgelistFiltered_size As Long

' static InstancedMeshRenderer instancedMeshRenderer: its members
Private imr_useInstancedRendering As Boolean ' = false
Private imr_instanceMeshes As ShapeVectorMap
Private imr_instanceTranslucentMeshes As ShapeVectorMap
Private imr_instanceTranslucentMeshesNoDepthWrite As ShapeVectorMap
Private imr_instanceAdditiveMeshes As ShapeVectorMap
Private imr_instancesCount As Long      ' size_t = 0
Private imr_translucentInstancesCount As Long
Private imr_additiveInstancesCount As Long
Private imr_tshapes() As SHAPE_
Private imr_tshapes_size As Long
Private imr_tshapes_cap As Long
Private imr_shapes() As SHAPE_
Private imr_shapes_size As Long
Private imr_shapes_cap As Long
Private imr_instancesData() As Draw3DShapePerInstanceInterleavedData
Private imr_instancesData_size As Long
Private imr_instanceDataBuffers() As gfx_api_buffer
Private imr_instanceDataBuffers_size As Long
Private imr_currInstanceBufferIdx As Long ' = 0
Private imr_lightmapTexture As gfx_api_texture ' = nullptr
Private imr_modelUVLightmapMatrix As mat4 ' = glm::mat4()
Private imr_finalizedDrawCalls() As InstancedDrawCall
Private imr_finalizedDrawCalls_size As Long
Private imr_finalizedDrawCalls_cap As Long
Private imr_startIdxTranslucentDrawCalls As Long
Private imr_startIdxTranslucentNoDepthWriteDrawCalls As Long
Private imr_startIdxAdditiveDrawCalls As Long

Private shadowCache_ As ShadowCache     ' static ShadowCache shadowCache
Private perFrameUniformsShaderOnce As ShaderOnce

Private piedrawReady As Boolean         ' (VB) the C++ initialisers have run

'
'  Source
'

' (VB) the globals' C++ initialisers (a VB module has none)
Private Sub InitPiedraw()
    If piedrawReady Then Exit Sub
    piedrawReady = True
    shadowMode_ = ShadowMode_Shadow_Mapping
    numShadowCascades = WZ_MAX_SHADOW_CASCADES
    Set shadowCache_.shapeMap = CreateObject("Scripting.Dictionary")
    shadowCache_.shapeMap.CompareMode = vbBinaryCompare
End Sub

' (VB) for GfxApi's forwarders: the Vector3f at address p; a Vector4f written to address p
Public Sub piedraw_vec3_fromPtr(ByVal p As Long, out As Vector3f)
    CopyMemory out, ByVal p, 12
End Sub

Public Sub piedraw_vec4_toPtr(v As Vector4f, ByVal p As Long)
    CopyMemory ByVal p, v, 16
End Sub

Public Sub pie_InitLighting()
    Dim i As Long, j As Long
    ' set scene color, ambient, diffuse and specular light intensities of sun
    ' diffuse lighting is turned off because players dislike it
    Dim defaultLight(0 To LIGHT_MAX - 1, 0 To 3) As Single
    defaultLight(LIGHT_EMISSIVE, 0) = 0!: defaultLight(LIGHT_EMISSIVE, 1) = 0!: defaultLight(LIGHT_EMISSIVE, 2) = 0!: defaultLight(LIGHT_EMISSIVE, 3) = 1!  ' LIGHT_EMISSIVE
    defaultLight(LIGHT_AMBIENT, 0) = 0.5!: defaultLight(LIGHT_AMBIENT, 1) = 0.5!: defaultLight(LIGHT_AMBIENT, 2) = 0.5!: defaultLight(LIGHT_AMBIENT, 3) = 1!  ' LIGHT_AMBIENT
    defaultLight(LIGHT_DIFFUSE, 0) = 1!: defaultLight(LIGHT_DIFFUSE, 1) = 1!: defaultLight(LIGHT_DIFFUSE, 2) = 1!: defaultLight(LIGHT_DIFFUSE, 3) = 1!  ' LIGHT_DIFFUSE
    defaultLight(LIGHT_SPECULAR, 0) = 1!: defaultLight(LIGHT_SPECULAR, 1) = 1!: defaultLight(LIGHT_SPECULAR, 2) = 1!: defaultLight(LIGHT_SPECULAR, 3) = 1!  ' LIGHT_SPECULAR
    For i = 0 To LIGHT_MAX - 1
        For j = 0 To 3
            lighting0(i, j) = defaultLight(i, j)
            lightingDefault(i, j) = defaultLight(i, j)
        Next
    Next
End Sub

Public Function pie_GetLighting0(ByVal entry As LIGHTING_TYPE) As Vector4f
    pie_GetLighting0.x = lighting0(entry, 0)
    pie_GetLighting0.y = lighting0(entry, 1)
    pie_GetLighting0.z = lighting0(entry, 2)
    pie_GetLighting0.w = lighting0(entry, 3)
End Function

' value: float[4]
Public Sub pie_Lighting0(ByVal entry As LIGHTING_TYPE, value() As Single)
    lighting0(entry, 0) = value(0)
    lighting0(entry, 1) = value(1)
    lighting0(entry, 2) = value(2)
    lighting0(entry, 3) = value(3)
End Sub

Private Function isShadowMappingEnabled() As Boolean
    InitPiedraw
    isShadowMappingEnabled = (shadows And shadowMode_ = ShadowMode_Shadow_Mapping)
End Function

Private Sub refreshShadowShaders()
    Dim shadowConstants As lighting_constants, bShadowMappingEnabled As Boolean, actualNumCascadesAndPasses As Long
    If gfx_api_context_isInitialized() Then
        shadowConstants = gfx_api_context_getShadowConstants()
        bShadowMappingEnabled = isShadowMappingEnabled()
        If bShadowMappingEnabled Then actualNumCascadesAndPasses = numShadowCascades Else actualNumCascadesAndPasses = 0
        If bShadowMappingEnabled Then
            shadowConstants.shadowMode = 1 ' Possible future TODO: Could get this from the config file to allow testing alternative filter methods
            shadowConstants.shadowCascadesCount = actualNumCascadesAndPasses
        Else
            shadowConstants.shadowMode = 0 ' Disable shader-drawn / shadow-mapping shadows
        End If

        ' Preserve existing shadowFilterSize

        gfx_api_context_setShadowConstants shadowConstants

        ' Trigger depth pass buffer free / rebuild if needed
        gfx_api_context_setDepthPassProperties actualNumCascadesAndPasses, gfx_api_context_getDepthPassDimensions(0)
    End If
End Sub

Public Sub pie_setShadows(ByVal drawShadows As Boolean)
    InitPiedraw
    shadows = drawShadows
    refreshShadowShaders
End Sub

' optional<bool> pie_supportsShadowMapping(): has_value returned, the value in value
Public Function pie_supportsShadowMapping(value As Boolean) As Boolean
    If Not gfx_api_context_isInitialized() Then
        ' can't determine support yet
        pie_supportsShadowMapping = False
        Exit Function
    End If

    ' double-check that current system supports instanced rendering - *only* if instanced rendering is possible does WZ support shadow mapping
    value = gfx_api_context_supportsInstancedRendering()
    pie_supportsShadowMapping = True
End Function

Public Function pie_setShadowMapResolution(ByVal resolution As Long) As Boolean
    Dim bShadowMappingEnabled As Boolean
    InitPiedraw
    If Not (resolution <> 0 And (resolution And (resolution - 1)) = 0) Then ASSERT_FAILURE "resolution && !(resolution & (resolution - 1))", "lib/ivis_opengl/piedraw.cpp:169", "pie_setShadowMapResolution", "Expecting power-of-2 resolution, received: %u", resolution: Exit Function
    bShadowMappingEnabled = isShadowMappingEnabled()
    If bShadowMappingEnabled Then
        pie_setShadowMapResolution = gfx_api_context_setDepthPassProperties(numShadowCascades, resolution)
    Else
        pie_setShadowMapResolution = gfx_api_context_setDepthPassProperties(0, resolution)
    End If
End Function

Public Function pie_getShadowMapResolution() As Long
    pie_getShadowMapResolution = gfx_api_context_getDepthPassDimensions(0)
End Function

Public Function pie_setShadowMode(ByVal mode As ShadowMode) As Boolean
    Dim successfulChangeToInputMode As Boolean, supported As Boolean
    InitPiedraw
    If mode = shadowMode_ And shadowsHasBeenInit Then
        pie_setShadowMode = True
        Exit Function
    End If
    If Not gfx_api_context_isInitialized() Then ASSERT_FAILURE "gfx_api::context::isInitialized()", "lib/ivis_opengl/piedraw.cpp:185", "pie_setShadowMode", "Can't set shadow mode until gfx backend is initialized": Exit Function
    successfulChangeToInputMode = True
    If mode = ShadowMode_Shadow_Mapping Then
        ' double-check that current system supports instanced rendering - *only* if instanced rendering is possible does WZ support shadow mapping
        If Not pie_supportsShadowMapping(supported) Then supported = False ' value_or(false)
        If Not supported Then
            ' for now, only the fallback stencil shadows are supported
            mode = ShadowMode_Fallback_Stencil_Shadows
            successfulChangeToInputMode = False
        End If
        ' double check that all conditions pass for the InstancedMeshRenderer to use instanced rendering (this may include additional checks)
        If Not canInstancedMeshRendererUseInstancedRendering(True) Then
            ' if the InstancedMeshRenderer can't use instanced rendering, use only the fallback stencil shadows
            mode = ShadowMode_Fallback_Stencil_Shadows
            successfulChangeToInputMode = False
        End If
    End If
    shadowMode_ = mode
    refreshShadowShaders
    shadowsHasBeenInit = True
    pie_setShadowMode = successfulChangeToInputMode
End Function

Public Function pie_getShadowMode() As ShadowMode
    InitPiedraw
    pie_getShadowMode = shadowMode_
End Function

Public Function pie_setShadowCascades(ByVal newValue As Long) As Boolean
    InitPiedraw
    If newValue = 0 Or newValue > WZ_MAX_SHADOW_CASCADES Then
        pie_setShadowCascades = False
        Exit Function
    End If
    If numShadowCascades = newValue Then
        pie_setShadowCascades = True
        Exit Function
    End If
    numShadowCascades = newValue
    refreshShadowShaders
    pie_setShadowCascades = True
End Function

Public Function pie_getShadowCascades() As Long
    InitPiedraw
    pie_getShadowCascades = numShadowCascades
End Function

Public Sub pie_BeginLighting(light_ As Vector3f)
    currentSunPosition = light_
End Sub

'**************************************************************************
' * pie_Draw3dShape
' *
' * Project and render a pumpkin image to render surface
' * Will support zbuffering, texturing, coloured lighting and alpha effects
' * Avoids recalculating vertex projections for every poly
' **************************************************************************

Private Function getZeroedVertexBuffer(ByVal size As Long) As gfx_api_buffer
    Dim tempZeroes() As Byte
    If pZeroedVertexBuffer Is Nothing Or (zeroedVertexBufferCurrentSize < size) Then
        If Not pZeroedVertexBuffer Is Nothing Then
            Set pZeroedVertexBuffer = Nothing ' delete pZeroedVertexBuffer
        End If
        Set pZeroedVertexBuffer = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_static_draw, "zeroedVertexBuffer")
        If size > 0 Then
            ReDim tempZeroes(0 To size - 1)
            pZeroedVertexBuffer.upload size, VarPtr(tempZeroes(0))
        Else
            pZeroedVertexBuffer.upload 0, 0
        End If
        zeroedVertexBufferCurrentSize = size
    End If
    Set getZeroedVertexBuffer = pZeroedVertexBuffer
End Function

' (VB) &pie_Texture(page) where the page may be iV_TEX_INVALID (then nullptr)
Private Function pageTextureOrNothing(ByVal page As Long) As gfx_api_texture
    If page <> iV_TEX_INVALID Then Set pageTextureOrNothing = pie_Texture(page)
End Function

Private Function boolToInt(ByVal b As Boolean) As Long
    If b Then boolToInt = 1
End Function

Public Sub pie_Draw3DButton(ByVal shape As Long, ByVal teamcolour As Long, modelMatrix As mat4, viewMatrix As mat4)
    Dim textures As iIMDShapeTextures, tcmask As gfx_api_texture, normalmap As gfx_api_texture, specularmap As gfx_api_texture
    Dim pTangentBuffer As gfx_api_buffer, colour As Long, pso As pipeline_state_helper
    Dim sceneColor As Vector4f, ambient As Vector4f, diffuse As Vector4f, specular As Vector4f
    Dim globalUniforms As Draw3DShapeGlobalUniforms, meshUniforms As Draw3DShapePerMeshUniforms, instanceUniforms As Draw3DShapePerInstanceUniforms

    textures = iIMDShape_getTextures(shape)
    Set tcmask = pageTextureOrNothing(textures.tcmaskpage)
    Set normalmap = pageTextureOrNothing(textures.normalpage)
    Set specularmap = pageTextureOrNothing(textures.specularpage)

    If Not imdShapes(shape).buffers(VBO_TANGENT) Is Nothing Then
        Set pTangentBuffer = imdShapes(shape).buffers(VBO_TANGENT)
    Else
        Set pTangentBuffer = getZeroedVertexBuffer(imdShapes(shape).vertexCount * 4 * 4)
    End If

    colour = psPalette(WZCOL_WHITE)
    sceneColor = glm_vec4(lightingDefault(LIGHT_EMISSIVE, 0), lightingDefault(LIGHT_EMISSIVE, 1), lightingDefault(LIGHT_EMISSIVE, 2), lightingDefault(LIGHT_EMISSIVE, 3))
    ambient = glm_vec4(lightingDefault(LIGHT_AMBIENT, 0), lightingDefault(LIGHT_AMBIENT, 1), lightingDefault(LIGHT_AMBIENT, 2), lightingDefault(LIGHT_AMBIENT, 3))
    diffuse = glm_vec4(lightingDefault(LIGHT_DIFFUSE, 0), lightingDefault(LIGHT_DIFFUSE, 1), lightingDefault(LIGHT_DIFFUSE, 2), lightingDefault(LIGHT_DIFFUSE, 3))
    specular = glm_vec4(lightingDefault(LIGHT_SPECULAR, 0), lightingDefault(LIGHT_SPECULAR, 1), lightingDefault(LIGHT_SPECULAR, 2), lightingDefault(LIGHT_SPECULAR, 3))
    Set pso = gfx_api_pso_get(Draw3DShapeOpaque)
    pso.bind

    globalUniforms.ProjectionMatrix = pie_UIPerspectiveGet()
    globalUniforms.ViewMatrix = viewMatrix
    globalUniforms.ShadowMapMVPMatrix = glm_mat4(1!)
    globalUniforms.sunPos = glm_vec4_v3(getDefaultSunPosition(), 0!)
    globalUniforms.sceneColor = sceneColor
    globalUniforms.ambient = ambient
    globalUniforms.diffuse = diffuse
    globalUniforms.specular = specular
    ' fogColour glm::vec4(0.f), fogEnd, fogBegin, timeState 0.f, fogEnabled 0

    meshUniforms.tcmask = boolToInt(Not tcmask Is Nothing)
    meshUniforms.normalMap = boolToInt(Not normalmap Is Nothing)
    meshUniforms.specularMap = boolToInt(Not specularmap Is Nothing)
    meshUniforms.hasTangents = boolToInt(Not imdShapes(shape).buffers(VBO_TANGENT) Is Nothing)

    instanceUniforms.ModelViewMatrix = glm_mat4_mul(viewMatrix, modelMatrix)
    instanceUniforms.NormalMatrix = glm_transpose(glm_inverse(modelMatrix))
    instanceUniforms.colour = pal_PIELIGHTtoVec4(colour)
    instanceUniforms.teamcolour = pal_PIELIGHTtoVec4(teamcolour)
    ' shaderStretch 0.f, animFrameNumber 0.f, ecmState 0
    instanceUniforms.alphaTest = boolToInt((imdShapes(shape).flags And pie_PREMULTIPLIED) = 0)

    ' TODO: Maybe use a separate shader that doesn't bother querying the shadow map?

    pso.set_uniforms VarPtr(globalUniforms), LenB(globalUniforms), VarPtr(meshUniforms), LenB(meshUniforms), VarPtr(instanceUniforms), LenB(instanceUniforms)

    pso.bind_textures pie_Texture(textures.texpage), tcmask, normalmap, specularmap
    pso.bind_vertex_buffers imdShapes(shape).buffers(VBO_VERTEX), imdShapes(shape).buffers(VBO_NORMAL), imdShapes(shape).buffers(VBO_TEXCOORD), pTangentBuffer
    gfx_api_context_bind_index_buffer imdShapes(shape).buffers(VBO_INDEX), index_type_u16
    pso.draw_elements imdShapes(shape).polys_size * 3, 0
    polyCount = polyCount + imdShapes(shape).polys_size
    pso.unbind_vertex_buffers imdShapes(shape).buffers(VBO_VERTEX), imdShapes(shape).buffers(VBO_NORMAL), imdShapes(shape).buffers(VBO_TEXCOORD), pTangentBuffer
    gfx_api_context_unbind_index_buffer imdShapes(shape).buffers(VBO_INDEX)
End Sub

' templatedState::operator==
Private Function templatedState_equal(a As templatedState, b As templatedState) As Boolean
    templatedState_equal = (a.shader = b.shader) And (a.shape = b.shape) And (a.pieFlag = b.pieFlag)
End Function

Private Function templatedState_new(ByVal shader As SHADER_MODE, ByVal shape As Long, ByVal pieFlag As Long) As templatedState
    templatedState_new.shader = shader
    templatedState_new.shape = shape
    templatedState_new.pieFlag = pieFlag
End Function

' (VB) ShaderOnce::perform_once<T>: whether func is to run now (and records that it ran)
Private Function ShaderOnce_perform_once(once As ShaderOnce, ByVal pso As pso_type) As Boolean
    If once.performed_once(pso) Then
        Exit Function
    End If
    once.performed_once(pso) = True
    ShaderOnce_perform_once = True
End Function

Private Sub ShaderOnce_reset(once As ShaderOnce)
    Dim blank As ShaderOnce
    once = blank
End Sub

' template<SHADER_MODE shader, typename AdditivePSO, typename AlphaPSO, typename AlphaNoDepthWRTPSO, typename PremultipliedPSO, typename OpaquePSO>
' (VB) the PSO types are the pso_type values; each branch of the C++ binds its own PSO the same way
Private Sub draw3dShapeTemplated(ByVal shader As SHADER_MODE, ByVal AdditivePSO As pso_type, ByVal AlphaPSO As pso_type, ByVal AlphaNoDepthWRTPSO As pso_type, ByVal PremultipliedPSO As pso_type, ByVal OpaquePSO As pso_type, lastState As templatedState, globalsOnce As ShaderOnce, globalUniforms As Draw3DShapeGlobalUniforms, ByVal colour As Long, ByVal teamcolour As Long, ByVal stretch As Single, ByVal ecmState As Long, modelViewMatrix As mat4, ByVal shape As Long, ByVal pieFlag As Long, ByVal frame As Long)
    Dim currentState As templatedState, textures As iIMDShapeTextures
    Dim tcmask As gfx_api_texture, normalmap As gfx_api_texture, specularmap As gfx_api_texture
    Dim meshUniforms As Draw3DShapePerMeshUniforms, instanceUniforms As Draw3DShapePerInstanceUniforms
    Dim pTangentBuffer As gfx_api_buffer, psoType As pso_type, pso As pipeline_state_helper
    currentState = templatedState_new(shader, shape, pieFlag)

    textures = iIMDShape_getTextures(shape)
    Set tcmask = pageTextureOrNothing(textures.tcmaskpage)
    Set normalmap = pageTextureOrNothing(textures.normalpage)
    Set specularmap = pageTextureOrNothing(textures.specularpage)

    meshUniforms.tcmask = boolToInt(Not tcmask Is Nothing)
    meshUniforms.normalMap = boolToInt(Not normalmap Is Nothing)
    meshUniforms.specularMap = boolToInt(Not specularmap Is Nothing)
    meshUniforms.hasTangents = boolToInt(Not imdShapes(shape).buffers(VBO_TANGENT) Is Nothing)

    instanceUniforms.ModelViewMatrix = modelViewMatrix
    instanceUniforms.NormalMatrix = glm_transpose(glm_inverse(modelViewMatrix))
    instanceUniforms.colour = pal_PIELIGHTtoVec4(colour)
    instanceUniforms.teamcolour = pal_PIELIGHTtoVec4(teamcolour)
    instanceUniforms.shaderStretch = stretch
    instanceUniforms.animFrameNumber = CSng(frame)
    instanceUniforms.ecmState = ecmState
    instanceUniforms.alphaTest = boolToInt((pieFlag And pie_PREMULTIPLIED) = 0)

    If Not imdShapes(shape).buffers(VBO_TANGENT) Is Nothing Then
        Set pTangentBuffer = imdShapes(shape).buffers(VBO_TANGENT)
    Else
        Set pTangentBuffer = getZeroedVertexBuffer(imdShapes(shape).vertexCount * 4 * 4)
    End If

    ' Set tranlucency
    If (pieFlag And pie_ADDITIVE) <> 0 Then
        psoType = AdditivePSO
    ElseIf (pieFlag And pie_TRANSLUCENT) <> 0 Then
        If (pieFlag And pie_NODEPTHWRITE) = 0 Then
            psoType = AlphaPSO
        Else
            psoType = AlphaNoDepthWRTPSO
        End If
    ElseIf (pieFlag And pie_PREMULTIPLIED) <> 0 Then
        psoType = PremultipliedPSO
    Else
        psoType = OpaquePSO
    End If
    Set pso = gfx_api_pso_get(psoType)
    pso.bind
    If ShaderOnce_perform_once(globalsOnce, psoType) Then
        pso.set_uniforms_at 0, VarPtr(globalUniforms), LenB(globalUniforms)
    End If
    If Not templatedState_equal(currentState, lastState) Then
        pso.set_uniforms_at 1, VarPtr(meshUniforms), LenB(meshUniforms)
        pso.bind_vertex_buffers imdShapes(shape).buffers(VBO_VERTEX), imdShapes(shape).buffers(VBO_NORMAL), imdShapes(shape).buffers(VBO_TEXCOORD), pTangentBuffer
        pso.bind_textures pie_Texture(textures.texpage), tcmask, normalmap, specularmap
        gfx_api_context_bind_index_buffer imdShapes(shape).buffers(VBO_INDEX), index_type_u16
    End If
    pso.set_uniforms_at 2, VarPtr(instanceUniforms), LenB(instanceUniforms)
    pso.draw_elements imdShapes(shape).polys_size * 3, 0
'   PSO::get().unbind_vertex_buffers(shape->buffers[VBO_VERTEX], shape->buffers[VBO_NORMAL], shape->buffers[VBO_TEXCOORD]);
End Sub

' PIELIGHT straightAlphaToPremultiplied(PIELIGHT colour)
Public Function straightAlphaToPremultiplied(ByVal colour As Long) As Long
    Dim a As Long, r As Long, g As Long, b As Long
    a = PIELIGHT_a(colour)
    r = (PIELIGHT_r(colour) * a) \ 255
    g = (PIELIGHT_g(colour) * a) \ 255
    b = (PIELIGHT_b(colour) * a) \ 255
    straightAlphaToPremultiplied = pal_RGBA(r, g, b, a)
End Function

Private Function GenerateInstanceData(ByVal frame As Long, ByVal colour As Long, ByVal teamcolour As Long, ByVal pieFlag As Long, ByVal pieFlagData As Long, modelMatrix As mat4, ByVal stretchDepth As Single) As Draw3DShapePerInstanceInterleavedData
    Dim ecmState As Long, r As Draw3DShapePerInstanceInterleavedData
    If (pieFlag And pie_ECM) <> 0 Then ecmState = 1 Else ecmState = 0

    ' Set translucency
    If (pieFlag And pie_ADDITIVE) <> 0 Then
        colour = PIELIGHT_setA(colour, pieFlagData)
    ElseIf (pieFlag And pie_TRANSLUCENT) <> 0 Then
        colour = PIELIGHT_setA(colour, pieFlagData)
    End If

    If (pieFlag And pie_PREMULTIPLIED) <> 0 Then
        colour = straightAlphaToPremultiplied(colour)
        teamcolour = straightAlphaToPremultiplied(teamcolour)
    End If

    r.ModelViewMatrix = modelMatrix
    r.shaderStretch_ecmState_alphaTest_animFrameNumber = glm_vec4(stretchDepth, CSng(ecmState), CSng(boolToInt((pieFlag And pie_PREMULTIPLIED) = 0)), CSng(frame))
    r.colour = colour                   ' colour.rgba()
    r.teamcolour = teamcolour
    GenerateInstanceData = r
End Function

Private Function pie_Draw3DShape2(lastState As templatedState, globalsOnce As ShaderOnce, globalUniforms As Draw3DShapeGlobalUniforms, ByVal shape As Long, ByVal frame As Long, ByVal colour As Long, ByVal teamcolour As Long, ByVal pieFlag As Long, ByVal pieFlagData As Long, modelMatrix As mat4, ByVal stretchDepth As Single) As templatedState
    Dim light_ As Boolean, ecmState As Long, currentState As templatedState, numFrames As Long, modelViewMatrix As mat4
    light_ = True
    ecmState = 0

    drawCallsCount = drawCallsCount + 1

    ' Set fog status
    If (pieFlag And pie_FORCE_FOG) = 0 And ((pieFlag And pie_ADDITIVE) <> 0 Or (pieFlag And pie_TRANSLUCENT) <> 0 Or (pieFlag And pie_PREMULTIPLIED) <> 0) Then
        pie_SetFogStatus False
    Else
        pie_SetFogStatus True
    End If

    ' Set translucency
    If (pieFlag And pie_ADDITIVE) <> 0 Then
        colour = PIELIGHT_setA(colour, pieFlagData)
        light_ = False
    ElseIf (pieFlag And pie_TRANSLUCENT) <> 0 Then
        colour = PIELIGHT_setA(colour, pieFlagData)
        light_ = False
    ElseIf (pieFlag And pie_PREMULTIPLIED) <> 0 Then
        light_ = False
    End If

    If (pieFlag And pie_ECM) <> 0 Then
        light_ = True
        ecmState = 1
    End If

    numFrames = imdShapes(shape).numFrames
    If numFrames < 1 Then numFrames = 1
    frame = frame Mod numFrames

    If light_ Then
        currentState = templatedState_new(SHADER_COMPONENT, shape, pieFlag)
    Else
        currentState = templatedState_new(SHADER_NOLIGHT, shape, pieFlag)
    End If

    modelViewMatrix = glm_mat4_mul(globalUniforms.ViewMatrix, modelMatrix)
    If light_ Then
        draw3dShapeTemplated SHADER_COMPONENT, Draw3DShapeAdditive, Draw3DShapeAlpha, Draw3DShapeAlphaNoDepthWRT, Draw3DShapePremul, Draw3DShapeOpaque, lastState, globalsOnce, globalUniforms, colour, teamcolour, stretchDepth, ecmState, modelViewMatrix, shape, pieFlag, frame
    Else
        draw3dShapeTemplated SHADER_NOLIGHT, Draw3DShapeNoLightAdditive, Draw3DShapeNoLightAlpha, Draw3DShapeNoLightAlphaNoDepthWRT, Draw3DShapeNoLightPremul, Draw3DShapeNoLightOpaque, lastState, globalsOnce, globalUniforms, colour, teamcolour, stretchDepth, ecmState, modelViewMatrix, shape, pieFlag, frame
    End If

    polyCount = polyCount + imdShapes(shape).polys_size

    pie_Draw3DShape2 = currentState
End Function

Private Function edgeLessThan(e1 As EDGE, e2 As EDGE) As Boolean
    edgeLessThan = e1.sort_key < e2.sort_key
End Function

' static_cast<uint64_t>(from) << 32 | to (a Double: exact, the indices being below 65536)
Private Function edgeSortKey(ByVal from As Long, ByVal to_ As Long) As Double
    edgeSortKey = CDbl(from) * 4294967296# + CDbl(to_)
End Function

Private Sub flipEdge(e As EDGE)
    Dim t As Long
    t = e.from: e.from = e.to_: e.to_ = t
    e.sort_key = edgeSortKey(e.from, e.to_)
End Sub

' scale the height according to the flags
Private Function scale_y(ByVal y As Single, ByVal flag As Long, ByVal flag_data As Long) As Single
    Dim tempY As Single
    tempY = y
    If (flag And pie_RAISE) <> 0 Then
        tempY = y - flag_data
        If y - flag_data < 0 Then
            tempY = 0
        End If
    ElseIf (flag And pie_HEIGHT_SCALED) <> 0 Then
        If y > 0 Then
            tempY = (y * flag_data) / pie_RAISE_SCALE
        End If
    End If
    scale_y = tempY
End Function

' (VB) the key of ShadowCache's two maps: the shape, then ShadowDrawParameters (flag, flag_data, light)
Private Function shadowCacheKey(ByVal shape As Long, ByVal flag As Long, ByVal flag_data As Long, light_ As Vector4f) As String
    shadowCacheKey = shape & "|" & flag & "|" & flag_data & "|" & FloatHex(light_.x) & FloatHex(light_.y) & FloatHex(light_.z) & FloatHex(light_.w)
End Function

' ShadowCache::findCacheForShadowDraw: the entry's index, or -1 (nullptr)
Private Function ShadowCache_findCacheForShadowDraw(cache As ShadowCache, ByVal shape As Long, ByVal flag As Long, ByVal flag_data As Long, light_ As Vector4f) As Long
    Dim key As String, idx As Long
    key = shadowCacheKey(shape, flag, flag_data, light_)
    If Not cache.shapeMap.Exists(key) Then
        ShadowCache_findCacheForShadowDraw = -1
        Exit Function
    End If
    idx = cache.shapeMap(key)
    ' update the frame in which we requested this shadow cache
    cache.entries(idx).lastQueriedFrameCount = cache.currentFrame_
    ShadowCache_findCacheForShadowDraw = idx
End Function

Private Function ShadowCache_createCacheForShadowDraw(cache As ShadowCache, ByVal shape As Long, ByVal flag As Long, ByVal flag_data As Long, light_ As Vector4f) As Long
    Dim key As String, idx As Long, blank As CachedShadowData
    key = shadowCacheKey(shape, flag, flag_data, light_)
    If cache.shapeMap.Exists(key) Then
        idx = cache.shapeMap(key)
    Else
        If cache.freeEntries_size > 0 Then
            cache.freeEntries_size = cache.freeEntries_size - 1
            idx = cache.freeEntries(cache.freeEntries_size)
        Else
            If cache.entries_size = 0 Then
                ReDim cache.entries(0 To 63)
            ElseIf cache.entries_size > UBound(cache.entries) Then
                ReDim Preserve cache.entries(0 To 2 * cache.entries_size - 1)
            End If
            idx = cache.entries_size
            cache.entries_size = cache.entries_size + 1
        End If
        cache.entries(idx) = blank
        cache.entries(idx).key = key
        cache.shapeMap.Add key, idx
    End If
    cache.entries(idx).lastQueriedFrameCount = cache.currentFrame_
    ShadowCache_createCacheForShadowDraw = idx
End Function

Private Sub ShadowCache_addPremultipliedVertexes(cache As ShadowCache, ByVal cachedData As Long, modelViewMatrix As mat4)
    Dim mat_a As Single, mat_b As Single, mat_c As Single, mat_d As Single
    Dim mat_e As Single, mat_f As Single, mat_g As Single, mat_h As Single
    Dim mat_i As Single, mat_j As Single, mat_k As Single, mat_l As Single
    Dim premult_x As Single, premult_y As Single, premult_z As Single, i As Long, n As Long
    mat_a = modelViewMatrix.m(0 * 4 + 0)
    mat_b = modelViewMatrix.m(1 * 4 + 0)
    mat_c = modelViewMatrix.m(2 * 4 + 0)
    mat_d = modelViewMatrix.m(3 * 4 + 0)
    mat_e = modelViewMatrix.m(0 * 4 + 1)
    mat_f = modelViewMatrix.m(1 * 4 + 1)
    mat_g = modelViewMatrix.m(2 * 4 + 1)
    mat_h = modelViewMatrix.m(3 * 4 + 1)
    mat_i = modelViewMatrix.m(0 * 4 + 2)
    mat_j = modelViewMatrix.m(1 * 4 + 2)
    mat_k = modelViewMatrix.m(2 * 4 + 2)
    mat_l = modelViewMatrix.m(3 * 4 + 2)
    n = cache.entries(cachedData).vertexes_size
    If n = 0 Then Exit Sub
    If cache.vertexes_size + n - 1 > ubVec3(cache.vertexes) Then
        ReDim Preserve cache.vertexes(0 To 2 * (cache.vertexes_size + n))
    End If
    For i = 0 To n - 1
        premult_x = cache.entries(cachedData).vertexes(i).x * mat_a + cache.entries(cachedData).vertexes(i).y * mat_b + cache.entries(cachedData).vertexes(i).z * mat_c + mat_d
        premult_y = cache.entries(cachedData).vertexes(i).x * mat_e + cache.entries(cachedData).vertexes(i).y * mat_f + cache.entries(cachedData).vertexes(i).z * mat_g + mat_h
        premult_z = cache.entries(cachedData).vertexes(i).x * mat_i + cache.entries(cachedData).vertexes(i).y * mat_j + cache.entries(cachedData).vertexes(i).z * mat_k + mat_l
        cache.vertexes(cache.vertexes_size) = glm_vec3(premult_x, premult_y, premult_z)
        cache.vertexes_size = cache.vertexes_size + 1
    Next
End Sub

Private Sub ShadowCache_clearPremultipliedVertexes(cache As ShadowCache)
    cache.vertexes_size = 0
End Sub

Private Sub ShadowCache_setCurrentFrame(cache As ShadowCache, ByVal currentFrame As Double)
    cache.currentFrame_ = currentFrame
End Sub

Private Function ShadowCache_removeUnused(cache As ShadowCache) As Long
    Dim oldItemsRemoved As Long, idx As Long, k As Variant, unused() As Long, n As Long, i As Long
    Dim blank As CachedShadowData
    If cache.shapeMap.count = 0 Then Exit Function
    ReDim unused(0 To cache.shapeMap.count - 1)
    For Each k In cache.shapeMap.keys
        idx = cache.shapeMap(k)
        If cache.entries(idx).lastQueriedFrameCount <> cache.currentFrame_ Then
            unused(n) = idx
            n = n + 1
        End If
    Next
    For i = 0 To n - 1
        idx = unused(i)
        cache.shapeMap.Remove cache.entries(idx).key
        cache.entries(idx) = blank
        If cache.freeEntries_size >= cache.freeEntries_cap Then
            If cache.freeEntries_cap = 0 Then cache.freeEntries_cap = 64 Else cache.freeEntries_cap = 2 * cache.freeEntries_cap
            ReDim Preserve cache.freeEntries(0 To cache.freeEntries_cap - 1)
        End If
        cache.freeEntries(cache.freeEntries_size) = idx
        cache.freeEntries_size = cache.freeEntries_size + 1
        oldItemsRemoved = oldItemsRemoved + 1
    Next
    ShadowCache_removeUnused = oldItemsRemoved
End Function

' (VB) UBound of a Vector3f array that may be empty (-1)
Private Function ubVec3(a() As Vector3f) As Long
    On Error GoTo empty_
    ubVec3 = UBound(a)
    Exit Function
empty_:
    ubVec3 = -1
End Function

' (VB) shape->pShadowPoints->data()[i]
Private Function shadowPoint(ByVal shape As Long, ByVal i As Long) As Vector3f
    If imdShapes(shape).shadowUsesAlt Then
        shadowPoint = imdShapes(shape).altShadowPoints(i)
    Else
        shadowPoint = imdShapes(shape).points(i)
    End If
End Function

' (VB) edgelist.push_back (cap: the array's capacity)
Private Sub pushEdge(list() As EDGE, size As Long, cap As Long, ByVal from As Long, ByVal to_ As Long)
    If size >= cap Then
        If cap = 0 Then cap = 256 Else cap = 2 * cap
        ReDim Preserve list(0 To cap - 1)
    End If
    list(size).from = from
    list(size).to_ = to_
    list(size).sort_key = edgeSortKey(from, to_)
    size = size + 1
End Sub

' (VB) ska_sort(begin, end, [](const EDGE & edge) { return edge.sort_key; }): by sort_key
Private Sub sortEdges(a() As EDGE, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Double, t As EDGE
    Do While hi - lo > 12
        pivot = a((lo + hi) \ 2).sort_key
        i = lo
        j = hi
        Do While i <= j
            Do While a(i).sort_key < pivot
                i = i + 1
            Loop
            Do While a(j).sort_key > pivot
                j = j - 1
            Loop
            If i <= j Then
                t = a(i): a(i) = a(j): a(j) = t
                i = i + 1
                j = j - 1
            End If
        Loop
        ' recurse into the smaller part, loop on the larger
        If j - lo < hi - i Then
            If lo < j Then sortEdges a, lo, j
            lo = i
        Else
            If i < hi Then sortEdges a, i, hi
            hi = j
        End If
    Loop
    ' insertion sort for the short rest
    For i = lo + 1 To hi
        t = a(i)
        j = i - 1
        Do While j >= lo
            If a(j).sort_key <= t.sort_key Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next
End Sub

' Draw the shadow for a shape
' Prequisite:
'		Caller must call the following before all calls to pie_DrawShadow():
'			const auto &program = pie_ActivateShader(SHADER_GENERIC_COLOR, pie_PerspectiveGet(), glm::vec4());
'			glEnableVertexAttribArray(program.locVertex);
'		and must call the following after all calls to pie_DrawShadow():
'			glDisableVertexAttribArray(program.locVertex);
'			pie_DeactivateShader();
'		The only place this is currently called is pie_ShadowDrawLoop(), which handles this properly.
Private Function pie_DrawShadow(cache As ShadowCache, ByVal shape As Long, ByVal flag As Long, ByVal flag_data As Long, light_ As Vector4f, modelViewMatrix As mat4) As Long
    Dim useShapeEdgeList As Boolean          ' EDGE *drawlist: shape->shadowEdgeList (else edgelistFiltered)
    Dim edge_count As Long, result As Long, pCached As Long, j As Long, n As Long, i As Long
    Dim current As Long, p(0 To 2) As Vector3f, pv As Vector3f, lightv As Vector3f, polys_size As Long
    Dim from As Long, to_ As Long, i1 As Long, i2 As Long, a As Long, b As Long
    Dim v1 As Vector3f, v3 As Vector3f, va As Vector3f, vb As Vector3f
    Dim vertexes() As Vector3f, vertexes_size As Long, poly As iIMDPoly

    ' Find cached data (if available)
    ' Note: The modelViewMatrix is not used for calculating the sorted / filtered vertices, so it's not included
    pCached = ShadowCache_findCacheForShadowDraw(cache, shape, flag, flag_data, light_)
    If pCached = -1 Then
        ' (shape->shadowEdgeList != nullptr: it is allocated only with nShadowEdges > 0)
        If (flag And pie_STATIC_SHADOW) <> 0 And imdShapes(shape).nShadowEdges > 0 Then
            useShapeEdgeList = True
            edge_count = imdShapes(shape).nShadowEdges
        Else
            edgelist_size = 0
            lightv = glm_vec3(light_.x, light_.y, light_.z)
            If imdShapes(shape).shadowUsesAlt Then polys_size = imdShapes(shape).altShadowPolys_size Else polys_size = imdShapes(shape).polys_size
            For i = 0 To polys_size - 1
                If imdShapes(shape).shadowUsesAlt Then poly = imdShapes(shape).altShadowPolys(i) Else poly = imdShapes(shape).polys(i)
                For j = 0 To 2
                    current = poly.pindex(j)
                    pv = shadowPoint(shape, current)
                    p(j) = glm_vec3(pv.x, scale_y(pv.y, flag, flag_data), pv.z)
                Next
                If glm_dot3(glm_cross(glm_vec3_sub(p(2), p(0)), glm_vec3_sub(p(1), p(0))), lightv) > 0! Then
                    For n = 0 To 2
                        ' Add the edges
                        from = poly.pindex(n)
                        to_ = poly.pindex((n + 1) Mod 3)
                        pushEdge edgelist, edgelist_size, edgelist_cap, from, to_
                    Next
                End If
            Next

            ' Remove duplicate pairs from the edge list. For example, in the list ((1 2), (2 6), (6 2), (3, 4)), remove (2 6) and (6 2).
            edgelistFlipped_size = edgelist_size
            If edgelist_size > 0 Then
                edgelistFlipped = edgelist
                For i = 0 To edgelist_size - 1
                    flipEdge edgelistFlipped(i)
                Next
                sortEdges edgelist, 0, edgelist_size - 1
                sortEdges edgelistFlipped, 0, edgelistFlipped_size - 1
            End If
            ' std::set_difference(edgelist, edgelistFlipped, edgelistFiltered, edgeLessThan)
            edgelistFiltered_size = 0
            If edgelist_size > 0 Then ReDim edgelistFiltered(0 To edgelist_size - 1)
            i1 = 0
            i2 = 0
            Do While i1 < edgelist_size
                If i2 >= edgelistFlipped_size Then
                    edgelistFiltered(edgelistFiltered_size) = edgelist(i1)
                    edgelistFiltered_size = edgelistFiltered_size + 1
                    i1 = i1 + 1
                ElseIf edgeLessThan(edgelist(i1), edgelistFlipped(i2)) Then
                    edgelistFiltered(edgelistFiltered_size) = edgelist(i1)
                    edgelistFiltered_size = edgelistFiltered_size + 1
                    i1 = i1 + 1
                Else
                    If Not edgeLessThan(edgelistFlipped(i2), edgelist(i1)) Then
                        i1 = i1 + 1
                    End If
                    i2 = i2 + 1
                End If
            Loop

            edge_count = edgelistFiltered_size
            useShapeEdgeList = False
            'debug(LOG_WARNING, "we have %i edges", edge_count);

            If (flag And pie_STATIC_SHADOW) <> 0 Then
                ' then store it in the imd
                imdShapes(shape).nShadowEdges = edge_count
                If edge_count > 0 Then
                    ReDim imdShapes(shape).shadowEdgeList(0 To imdShapes(shape).nShadowEdges - 1)
                    For i = 0 To edge_count - 1
                        imdShapes(shape).shadowEdgeList(i) = edgelistFiltered(i)
                    Next
                End If
            End If
        End If

        vertexes_size = 0
        If edge_count > 0 Then ReDim vertexes(0 To edge_count * 6 - 1)
        For i = 0 To edge_count - 1
            If useShapeEdgeList Then
                a = imdShapes(shape).shadowEdgeList(i).from: b = imdShapes(shape).shadowEdgeList(i).to_
            Else
                a = edgelistFiltered(i).from: b = edgelistFiltered(i).to_
            End If
            va = shadowPoint(shape, a)
            vb = shadowPoint(shape, b)

            v1 = glm_vec3(vb.x, scale_y(vb.y, flag, flag_data), vb.z)
            v3 = glm_vec3(va.x + light_.x, scale_y(va.y, flag, flag_data) + light_.y, va.z + light_.z)

            vertexes(vertexes_size) = v1
            vertexes(vertexes_size + 1) = glm_vec3(vb.x + light_.x, scale_y(vb.y, flag, flag_data) + light_.y, vb.z + light_.z) 'v2
            vertexes(vertexes_size + 2) = v3

            vertexes(vertexes_size + 3) = v3
            vertexes(vertexes_size + 4) = glm_vec3(va.x, scale_y(va.y, flag, flag_data), va.z) 'v4
            vertexes(vertexes_size + 5) = v1
            vertexes_size = vertexes_size + 6
        Next

        pCached = ShadowCache_createCacheForShadowDraw(cache, shape, flag, flag_data, light_)
        cache.entries(pCached).vertexes_size = vertexes_size
        If vertexes_size > 0 Then
            cache.entries(pCached).vertexes = vertexes ' std::move(vertexes)
        Else
            Erase cache.entries(pCached).vertexes
        End If
        result = DRAW_SUCCESS_UNCACHED
    Else
        result = DRAW_SUCCESS_CACHED
    End If

    ' Aggregate the vertexes (pre-computed with the modelViewMatrix)
    ShadowCache_addPremultipliedVertexes cache, pCached, modelViewMatrix

    pie_DrawShadow = result
End Function

Private Function canInstancedMeshRendererUseInstancedRendering(Optional ByVal quiet As Boolean = False) As Boolean
    Dim max_vertex_attribs As Long, max_vertex_output_components As Long, maxInstancedMeshShaderVertexAttribs As Long
    Const min_required_vertex_output_components As Long = 12 * 4 ' from the shaders - see tcmask_instanced.vert / nolight_instanced.vert
    If Not gfx_api_context_supportsInstancedRendering() Then
        canInstancedMeshRendererUseInstancedRendering = False
        Exit Function
    End If

    ' check for minimum required vertex attributes (and vertex outputs) for the instanced shaders
    max_vertex_attribs = gfx_api_context_get_context_value(context_value_MAX_VERTEX_ATTRIBS)
    max_vertex_output_components = gfx_api_context_get_context_value(context_value_MAX_VERTEX_OUTPUT_COMPONENTS)
    ' std::max({gfx_api::instance_modelMatrix, gfx_api::instance_packedValues, gfx_api::instance_Colour, gfx_api::instance_TeamColour}) + 1
    maxInstancedMeshShaderVertexAttribs = gfx_api_instance_modelMatrix
    If gfx_api_instance_packedValues > maxInstancedMeshShaderVertexAttribs Then maxInstancedMeshShaderVertexAttribs = gfx_api_instance_packedValues
    If gfx_api_instance_Colour > maxInstancedMeshShaderVertexAttribs Then maxInstancedMeshShaderVertexAttribs = gfx_api_instance_Colour
    If gfx_api_instance_TeamColour > maxInstancedMeshShaderVertexAttribs Then maxInstancedMeshShaderVertexAttribs = gfx_api_instance_TeamColour
    maxInstancedMeshShaderVertexAttribs = maxInstancedMeshShaderVertexAttribs + 1
    If max_vertex_attribs < maxInstancedMeshShaderVertexAttribs Then
        If Not quiet Then
            debug_ 946, LOG_INFO, "canInstancedMeshRendererUseInstancedRendering", "Disabling instanced rendering - Insufficient MAX_VERTEX_ATTRIBS (%i; need: %zu)", max_vertex_attribs, maxInstancedMeshShaderVertexAttribs
        End If
        canInstancedMeshRendererUseInstancedRendering = False
        Exit Function
    End If
    If max_vertex_output_components < min_required_vertex_output_components Then
        If Not quiet Then
            debug_ 954, LOG_INFO, "canInstancedMeshRendererUseInstancedRendering", "Disabling instanced rendering - Insufficient MAX_VERTEX_OUTPUT_COMPONENTS (%i; need: %zu)", max_vertex_output_components, min_required_vertex_output_components
        End If
        canInstancedMeshRendererUseInstancedRendering = False
        Exit Function
    End If

    ' passed all context checks
    canInstancedMeshRendererUseInstancedRendering = True
End Function

'--------------------------------------------------------------------------
' (VB) ShapeVectorMap: std::unordered_map<MeshInstanceKey, ShapeVector> and InsertionOrderMap<MeshInstanceKey, ShapeVector>

Private Function ShapeVectorMap_hash(key As templatedState, ByVal mask As Long) As Long
    ShapeVectorMap_hash = (key.shape Xor (key.pieFlag * 4099) Xor (key.shader * 65537)) And mask
End Function

Private Sub ShapeVectorMap_clear(m As ShapeVectorMap)
    m.keys_size = 0
    m.shapes_size = 0
    m.gen = m.gen + 1
    If m.gen = &H7FFFFFFF Then
        ReDim m.slotGen(0 To m.slotMask)
        m.gen = 1
    End If
End Sub

' the key's index (inserted at the end if new)
Private Function ShapeVectorMap_index(m As ShapeVectorMap, key As templatedState) As Long
    Dim s As Long, k As Long, i As Long
    If m.slotMask = 0 Then
        m.slotMask = 255
        ReDim m.slots(0 To m.slotMask)
        ReDim m.slotGen(0 To m.slotMask)
        If m.gen = 0 Then m.gen = 1
        ReDim m.keys(0 To 63)
        ReDim m.firstShape(0 To 63)
        ReDim m.lastShape(0 To 63)
        ReDim m.counts(0 To 63)
    End If
    s = ShapeVectorMap_hash(key, m.slotMask)
    Do While m.slotGen(s) = m.gen
        k = m.slots(s) - 1
        If templatedState_equal(m.keys(k), key) Then
            ShapeVectorMap_index = k
            Exit Function
        End If
        s = (s + 1) And m.slotMask
    Loop
    ' a new key
    k = m.keys_size
    If k > UBound(m.keys) Then
        ReDim Preserve m.keys(0 To 2 * k - 1)
        ReDim Preserve m.firstShape(0 To 2 * k - 1)
        ReDim Preserve m.lastShape(0 To 2 * k - 1)
        ReDim Preserve m.counts(0 To 2 * k - 1)
    End If
    m.keys(k) = key
    m.firstShape(k) = -1
    m.lastShape(k) = -1
    m.counts(k) = 0
    m.keys_size = k + 1
    m.slots(s) = k + 1
    m.slotGen(s) = m.gen
    ' keep the table at most half full
    If 2 * m.keys_size > m.slotMask Then
        m.slotMask = 2 * m.slotMask + 1
        ReDim m.slots(0 To m.slotMask)
        ReDim m.slotGen(0 To m.slotMask)
        For i = 0 To m.keys_size - 1
            s = ShapeVectorMap_hash(m.keys(i), m.slotMask)
            Do While m.slotGen(s) = m.gen
                s = (s + 1) And m.slotMask
            Loop
            m.slots(s) = i + 1
            m.slotGen(s) = m.gen
        Next
    End If
    ShapeVectorMap_index = k
End Function

' m[key].push_back(shp)
Private Sub ShapeVectorMap_push(m As ShapeVectorMap, key As templatedState, shp As SHAPE_)
    Dim k As Long, i As Long
    k = ShapeVectorMap_index(m, key)
    i = m.shapes_size
    If i >= m.shapes_cap Then
        If m.shapes_cap = 0 Then m.shapes_cap = 256 Else m.shapes_cap = 2 * m.shapes_cap
        ReDim Preserve m.shapes(0 To m.shapes_cap - 1)
        ReDim Preserve m.nextShape(0 To m.shapes_cap - 1)
    End If
    m.shapes(i) = shp
    m.nextShape(i) = -1
    If m.lastShape(k) = -1 Then m.firstShape(k) = i Else m.nextShape(m.lastShape(k)) = i
    m.lastShape(k) = i
    m.counts(k) = m.counts(k) + 1
    m.shapes_size = i + 1
End Sub

' (VB) ShapeVector::push_back (cap: the array's capacity)
Private Sub pushShape(v() As SHAPE_, size As Long, cap As Long, shp As SHAPE_)
    If size >= cap Then
        If cap = 0 Then cap = 256 Else cap = 2 * cap
        ReDim Preserve v(0 To cap - 1)
    End If
    v(size) = shp
    size = size + 1
End Sub

'--------------------------------------------------------------------------
' class InstancedMeshRenderer (the one instance's methods)

Private Function InstancedMeshRenderer_initialize() As Boolean
    Dim i As Long
    InstancedMeshRenderer_reset

    If Not canInstancedMeshRendererUseInstancedRendering(False) Then
        imr_useInstancedRendering = False
        InstancedMeshRenderer_initialize = True
        Exit Function
    End If

    imr_instanceDataBuffers_size = gfx_api_context_maxFramesInFlight() + 1
    ReDim imr_instanceDataBuffers(0 To imr_instanceDataBuffers_size - 1)
    For i = 0 To imr_instanceDataBuffers_size - 1
        Set imr_instanceDataBuffers(i) = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_stream_draw, "InstancedMeshRenderer::instanceDataBuffer[" & CStr(i) & "]")
    Next
    imr_useInstancedRendering = True
    InstancedMeshRenderer_initialize = True
End Function

Private Sub InstancedMeshRenderer_clear()
    Dim blank As mat4
    ShapeVectorMap_clear imr_instanceMeshes
    ShapeVectorMap_clear imr_instanceTranslucentMeshes
    ShapeVectorMap_clear imr_instanceTranslucentMeshesNoDepthWrite
    ShapeVectorMap_clear imr_instanceAdditiveMeshes
    imr_instancesCount = 0
    imr_translucentInstancesCount = 0
    imr_additiveInstancesCount = 0
    imr_tshapes_size = 0
    imr_shapes_size = 0
    Set imr_lightmapTexture = Nothing
    imr_modelUVLightmapMatrix = blank   ' glm::mat4()
End Sub

Private Sub InstancedMeshRenderer_reset()
    Dim i As Long
    InstancedMeshRenderer_clear
    For i = 0 To imr_instanceDataBuffers_size - 1
        Set imr_instanceDataBuffers(i) = Nothing ' delete buffer
    Next
    Erase imr_instanceDataBuffers
    imr_instanceDataBuffers_size = 0
End Sub

Private Sub InstancedMeshRenderer_setLightmap(ByVal lightmapTexture_ As gfx_api_texture, modelUVLightmapMatrix_ As mat4)
    Set imr_lightmapTexture = lightmapTexture_
    imr_modelUVLightmapMatrix = modelUVLightmapMatrix_
End Sub

' Queues a mesh for drawing
Private Function InstancedMeshRenderer_Draw3DShape(ByVal shape As Long, ByVal frame As Long, ByVal teamcolour As Long, ByVal colour As Long, ByVal pieFlag As Long, ByVal pieFlagData As Long, modelMatrix As mat4, viewMatrix As mat4, ByVal stretchDepth As Single) As Boolean
    Dim light_ As Boolean, numFrames As Long, currentState As templatedState, tshape As SHAPE_, shader As SHADER_MODE
    Dim distance As Single, scshape As ShadowcastingShape, pos_light0 As Vector4f, invmat As mat4
    light_ = True

    If (pieFlag And pie_ADDITIVE) <> 0 Then
        light_ = False
    ElseIf (pieFlag And pie_TRANSLUCENT) <> 0 Then
        light_ = False
    ElseIf (pieFlag And pie_PREMULTIPLIED) <> 0 Then
        light_ = False
    End If

    If (pieFlag And pie_ECM) <> 0 Then
        light_ = True
    End If

    If (pieFlag And pie_FORCELIGHT) <> 0 Then
        light_ = True
    End If

    numFrames = imdShapes(shape).numFrames
    If numFrames < 1 Then numFrames = 1
    frame = frame Mod numFrames

    If light_ Then
        If imr_useInstancedRendering Then shader = SHADER_COMPONENT_INSTANCED Else shader = SHADER_COMPONENT
    Else
        If imr_useInstancedRendering Then shader = SHADER_NOLIGHT_INSTANCED Else shader = SHADER_NOLIGHT
    End If
    currentState = templatedState_new(shader, shape, pieFlag)

    tshape.shape = shape
    tshape.frame = frame
    tshape.colour = colour
    tshape.teamcolour = teamcolour
    tshape.flag = pieFlag
    tshape.flag_data = pieFlagData
    tshape.stretch = stretchDepth
    tshape.modelMatrix = modelMatrix

    If (pieFlag And pie_SHIELD) <> 0 Then
        tshape.modelMatrix = glm_scale_m(tshape.modelMatrix, glm_vec3(pie_SHIELD_FACTOR, pie_SHIELD_FACTOR, pie_SHIELD_FACTOR))
    End If

    If (pieFlag And pie_HEIGHT_SCALED) <> 0 Then ' construct
        tshape.modelMatrix = glm_scale_m(tshape.modelMatrix, glm_vec3(1!, CSng(pieFlagData) / CSng(pie_RAISE_SCALE), 1!))
    End If
    If (pieFlag And pie_RAISE) <> 0 Then     ' collapse
        tshape.modelMatrix = glm_translate_m(tshape.modelMatrix, glm_vec3(1!, (-imdShapes(shape).max.y * (pie_RAISE_SCALE - pieFlagData)) * (1! / pie_RAISE_SCALE), 1!))
    End If

    If (pieFlag And (pie_ADDITIVE Or pie_PREMULTIPLIED)) <> 0 Then
        If imr_useInstancedRendering Then
            ShapeVectorMap_push imr_instanceAdditiveMeshes, currentState, tshape
        Else
            pushShape imr_tshapes, imr_tshapes_size, imr_tshapes_cap, tshape
        End If
        imr_additiveInstancesCount = imr_additiveInstancesCount + 1
    ElseIf (pieFlag And pie_TRANSLUCENT) <> 0 Then
        If imr_useInstancedRendering Then
            If (pieFlag And pie_NODEPTHWRITE) <> 0 Then
                ShapeVectorMap_push imr_instanceTranslucentMeshesNoDepthWrite, currentState, tshape
            Else
                ShapeVectorMap_push imr_instanceTranslucentMeshes, currentState, tshape
            End If
        Else
            pushShape imr_tshapes, imr_tshapes_size, imr_tshapes_cap, tshape
        End If
        imr_translucentInstancesCount = imr_translucentInstancesCount + 1
    Else
        If shadows And ((pieFlag And pie_SHADOW) <> 0 Or (pieFlag And pie_STATIC_SHADOW) <> 0) And (shadowMode_ = ShadowMode_Fallback_Stencil_Shadows) Then
            ' draw a shadow
            scshape.modelViewMatrix = glm_mat4_mul(viewMatrix, modelMatrix)
            distance = scshape.modelViewMatrix.m(3 * 4 + 0) * scshape.modelViewMatrix.m(3 * 4 + 0)
            distance = distance + scshape.modelViewMatrix.m(3 * 4 + 1) * scshape.modelViewMatrix.m(3 * 4 + 1)
            distance = distance + scshape.modelViewMatrix.m(3 * 4 + 2) * scshape.modelViewMatrix.m(3 * 4 + 2)

            ' if object is too far in the fog don't generate a shadow.
            If distance < SHADOW_END_DISTANCE Then
                ' Calculate the light position relative to the object
                pos_light0 = glm_vec4_v3(currentSunPosition, 0!)
                invmat = glm_inverse(scshape.modelViewMatrix)

                scshape.light_ = glm_mat4_mul_vec4(invmat, pos_light0)
                scshape.shape = shape
                scshape.flag = pieFlag
                scshape.flag_data = pieFlagData

                If scshapes_size >= scshapes_cap Then
                    If scshapes_cap = 0 Then scshapes_cap = 64 Else scshapes_cap = 2 * scshapes_cap
                    ReDim Preserve scshapes(0 To scshapes_cap - 1)
                End If
                scshapes(scshapes_size) = scshape
                scshapes_size = scshapes_size + 1
            End If
        End If
        If imr_useInstancedRendering Then
            ShapeVectorMap_push imr_instanceMeshes, currentState, tshape
        Else
            pushShape imr_shapes, imr_shapes_size, imr_shapes_cap, tshape
        End If

        imr_instancesCount = imr_instancesCount + 1
    End If

    InstancedMeshRenderer_Draw3DShape = True
End Function

Public Sub pie_InitializeInstancedRenderer()
    InitPiedraw
    InstancedMeshRenderer_initialize
End Sub

Public Sub pie_CleanUp()
    InstancedMeshRenderer_reset
    scshapes_size = 0                   ' scshapes.clear()
    Erase scshapes
    scshapes_cap = 0
    If Not pZeroedVertexBuffer Is Nothing Then
        Set pZeroedVertexBuffer = Nothing ' delete pZeroedVertexBuffer
    End If
End Sub

' bool pie_Draw3DShape(const iIMDShape *shape, int frame, int team, PIELIGHT colour, int pieFlag, int pieFlagData,
'                      const glm::mat4 &modelMatrix, const glm::mat4 &viewMatrix, float stretchDepth = 0.f, bool onlySingleLevel = false)
Public Function pie_Draw3DShape(ByVal shape As Long, ByVal frame As Long, ByVal team As Long, ByVal colour As Long, ByVal pieFlag As Long, ByVal pieFlagData As Long, modelMatrix As mat4, viewMatrix As mat4, Optional ByVal stretchDepth As Single = 0!, Optional ByVal onlySingleLevel As Boolean = False) As Boolean
    Dim retVal As Boolean, drawAllLevels As Boolean, teamcolour As Long, pCurrShape As Long
    InitPiedraw
    pieCount = pieCount + 1

    If Not (frame >= 0) Then ASSERT_FAILURE "frame >= 0", "lib/ivis_opengl/piedraw.cpp:1302", "pie_Draw3DShape", "Negative frame %d", frame
    If Not (team >= 0) Then ASSERT_FAILURE "team >= 0", "lib/ivis_opengl/piedraw.cpp:1303", "pie_Draw3DShape", "Negative team %d", team

    retVal = False
    drawAllLevels = (imdShapes(shape).modelLevel = 0) And Not onlySingleLevel
    teamcolour = iIMDShape_getTeamColourForModel(shape, team)

    pCurrShape = shape
    Do
        If (pieFlag And pie_BUTTON) <> 0 Then
            pie_Draw3DButton pCurrShape, teamcolour, modelMatrix, viewMatrix
            retVal = True
        Else
            retVal = InstancedMeshRenderer_Draw3DShape(pCurrShape, frame, teamcolour, colour, pieFlag, pieFlagData, modelMatrix, viewMatrix, stretchDepth)
        End If

        pCurrShape = imdShapes(pCurrShape).nxt

    Loop While drawAllLevels And pCurrShape <> 0 And retVal

    pie_Draw3DShape = retVal
End Function

Private Sub pie_ShadowDrawLoop(cache As ShadowCache, projectionMatrix As mat4)
    Dim i As Long, pso As pipeline_state_helper, cb As constant_buffer_type_SHADER_GENERIC_COLOR
    Dim startingIndex As Long, vertex_count As Long, n As Long
    Const SHADOW_BATCH_MAX As Long = 8192& * 3 ' must be divisible by 3
'	size_t cachedShadowDraws = 0;
'	size_t uncachedShadowDraws = 0;
    For i = 0 To scshapes_size - 1
        '/*DrawShadowResult result =*/
        pie_DrawShadow cache, scshapes(i).shape, scshapes(i).flag, scshapes(i).flag_data, scshapes(i).light_, scshapes(i).modelViewMatrix
    Next

    If cache.vertexes_size > 0 Then
        ' Draw the shadow volume
        Set pso = gfx_api_pso_get(DrawStencilShadow)
        pso.bind
        ' The vertexes returned by shadowCache.getPremultipliedVertexes() are pre-multiplied by the modelViewMatrix
        ' Thus we only need to include the perspective matrix
        cb.transform_matrix = projectionMatrix
        pso.bind_constants VarPtr(cb), LenB(cb)
        gfx_api_context_bind_streamed_vertex_buffers VarPtr(cache.vertexes(0)), 12 * cache.vertexes_size

        ' Batch into glDrawArrays calls of <= SHADOW_BATCH_MAX
        vertex_count = cache.vertexes_size
        For startingIndex = 0 To vertex_count - 1 Step SHADOW_BATCH_MAX
            n = vertex_count - startingIndex
            If n > SHADOW_BATCH_MAX Then n = SHADOW_BATCH_MAX
            pso.draw n, startingIndex
        Next

        gfx_api_context_disable_all_vertex_buffers
    End If

    ShadowCache_clearPremultipliedVertexes cache

'	debug(LOG_INFO, "Cached shadow draws: %lu, uncached shadow draws: %lu", cachedShadowDraws, uncachedShadowDraws);
End Sub

Private Sub pie_DrawShadows(ByVal currentGameFrame As Double, projectionMatrix As mat4)
    Dim width As Long, height As Long, grey As Long
    width = pie_GetVideoBufferWidth()
    height = pie_GetVideoBufferHeight()
    ShadowCache_setCurrentFrame shadowCache_, currentGameFrame

    pie_ShadowDrawLoop shadowCache_, projectionMatrix

    grey = pal_RGBA(0, 0, 0, 128)       ' grey.byte = { 0, 0, 0, 128 };
    pie_BoxFill_alpha 0, 0, width, height, grey

    scshapes_size = 0                   ' scshapes.resize(0)
    ShadowCache_removeUnused shadowCache_
End Sub

Public Sub pie_StartMeshes()
    InstancedMeshRenderer_clear
End Sub

Public Sub pie_UpdateLightmap(ByVal lightmapTexture As gfx_api_texture, modelUVLightmapMatrix As mat4)
    InstancedMeshRenderer_setLightmap lightmapTexture, modelUVLightmapMatrix
End Sub

Public Sub pie_FinalizeMeshes(ByVal currentGameFrame As Double)
    InstancedMeshRenderer_FinalizeInstances
End Sub

Public Sub pie_DrawAllMeshes(ByVal currentGameFrame As Double, projectionMatrix As mat4, viewMatrix As mat4, cameraPos As Vector3f, shadowMVPMatrix As ShadowCascadesInfo, ByVal depthPass As Boolean)
    Dim drawParts As Long
    InitPiedraw
    drawParts = DrawParts_All
    If shadowMode_ = ShadowMode_Fallback_Stencil_Shadows Then
        drawParts = drawParts Or DrawParts_OldShadows
    End If
    If depthPass Then
        drawParts = 0
        If shadows Then
            drawParts = drawParts Or DrawParts_ShadowCastingShapes
        End If
    End If

    If drawParts = 0 Then
        Exit Sub
    End If
    InstancedMeshRenderer_DrawAll currentGameFrame, projectionMatrix, viewMatrix, cameraPos, shadowMVPMatrix, drawParts, depthPass
End Sub

' (VB) the four loops of FinalizeInstances: one map's instances, key by key
Private Sub appendMapInstances(m As ShapeVectorMap)
    Dim k As Long, i As Long, startingIdxInInstancesBuffer As Long
    For k = 0 To m.keys_size - 1
        startingIdxInInstancesBuffer = imr_instancesData_size
        i = m.firstShape(k)
        Do While i <> -1
            imr_instancesData(imr_instancesData_size) = GenerateInstanceData(m.shapes(i).frame, m.shapes(i).colour, m.shapes(i).teamcolour, m.shapes(i).flag, m.shapes(i).flag_data, m.shapes(i).modelMatrix, m.shapes(i).stretch)
            imr_instancesData_size = imr_instancesData_size + 1
            i = m.nextShape(i)
        Loop
        ' finalizedDrawCalls.emplace_back(mesh.first, meshInstances.size(), startingIdxInInstancesBuffer)
        If imr_finalizedDrawCalls_size >= imr_finalizedDrawCalls_cap Then
            If imr_finalizedDrawCalls_cap = 0 Then imr_finalizedDrawCalls_cap = 64 Else imr_finalizedDrawCalls_cap = 2 * imr_finalizedDrawCalls_cap
            ReDim Preserve imr_finalizedDrawCalls(0 To imr_finalizedDrawCalls_cap - 1)
        End If
        imr_finalizedDrawCalls(imr_finalizedDrawCalls_size).state = m.keys(k)
        imr_finalizedDrawCalls(imr_finalizedDrawCalls_size).instance_count = m.counts(k)
        imr_finalizedDrawCalls(imr_finalizedDrawCalls_size).startingIdxInInstancesBuffer = startingIdxInInstancesBuffer
        imr_finalizedDrawCalls_size = imr_finalizedDrawCalls_size + 1
    Next
End Sub

' Finalizes queued meshes, ready for one or more DrawAll calls
' (After this is called, Draw3DShape should not be called until the InstancedMeshRenderer is clear()-ed)
Private Function InstancedMeshRenderer_FinalizeInstances() As Boolean
    Dim total As Long
    If Not imr_useInstancedRendering Then
        ' nothing to do - without instanced rendering we just issue *tons* of draw calls...
        InstancedMeshRenderer_FinalizeInstances = True
        Exit Function
    End If

    imr_instancesData_size = 0
    imr_finalizedDrawCalls_size = 0

    total = imr_instancesCount + imr_translucentInstancesCount + imr_additiveInstancesCount
    If total = 0 Then
        InstancedMeshRenderer_FinalizeInstances = True
        Exit Function
    End If

    ' instancesData.reserve(instancesCount + translucentInstancesCount + additiveInstancesCount)
    If ubInstances() < total - 1 Then ReDim imr_instancesData(0 To total - 1)

    appendMapInstances imr_instanceMeshes

    imr_startIdxTranslucentDrawCalls = imr_finalizedDrawCalls_size

    appendMapInstances imr_instanceTranslucentMeshes

    imr_startIdxTranslucentNoDepthWriteDrawCalls = imr_finalizedDrawCalls_size

    appendMapInstances imr_instanceTranslucentMeshesNoDepthWrite

    imr_startIdxAdditiveDrawCalls = imr_finalizedDrawCalls_size

    appendMapInstances imr_instanceAdditiveMeshes

    ' Upload buffer
    imr_currInstanceBufferIdx = imr_currInstanceBufferIdx + 1
    If imr_currInstanceBufferIdx >= imr_instanceDataBuffers_size Then
        imr_currInstanceBufferIdx = 0
    End If
    imr_instanceDataBuffers(imr_currInstanceBufferIdx).upload imr_instancesData_size * LenB(imr_instancesData(0)), VarPtr(imr_instancesData(0))

    imr_instancesData_size = 0          ' instancesData.clear()

    InstancedMeshRenderer_FinalizeInstances = True
End Function

Private Function ubInstances() As Long
    On Error GoTo empty_
    ubInstances = UBound(imr_instancesData)
    Exit Function
empty_:
    ubInstances = -1
End Function

' Draws all queued meshes, given a projection + view matrix
Private Function InstancedMeshRenderer_DrawAll(ByVal currentGameFrame As Double, projectionMatrix As mat4, viewMatrix As mat4, cameraPos As Vector3f, shadowCascades As ShadowCascadesInfo, ByVal drawParts As Long, ByVal depthPass As Boolean) As Boolean
    Dim sceneColor As Vector4f, ambient As Vector4f, diffuse As Vector4f, specular As Vector4f
    Dim renderState As RENDER_STATE, fogColor As Vector4f, dimW As Long, dimH As Long
    ' (VB) Static: VB allows a procedure 64 KB of locals; these are filled in full before each use
    Static bucketLight As PointLightBuckets, globalUniformsInst As Draw3DShapeInstancedGlobalUniforms
    Dim globalUniformsOld As Draw3DShapeGlobalUniforms
    ShaderOnce_reset perFrameUniformsShaderOnce

    ' Generate global (per-frame) uniforms
    sceneColor = glm_vec4(lighting0(LIGHT_EMISSIVE, 0), lighting0(LIGHT_EMISSIVE, 1), lighting0(LIGHT_EMISSIVE, 2), lighting0(LIGHT_EMISSIVE, 3))
    ambient = glm_vec4(lighting0(LIGHT_AMBIENT, 0), lighting0(LIGHT_AMBIENT, 1), lighting0(LIGHT_AMBIENT, 2), lighting0(LIGHT_AMBIENT, 3))
    diffuse = glm_vec4(lighting0(LIGHT_DIFFUSE, 0), lighting0(LIGHT_DIFFUSE, 1), lighting0(LIGHT_DIFFUSE, 2), lighting0(LIGHT_DIFFUSE, 3))
    specular = glm_vec4(lighting0(LIGHT_SPECULAR, 0), lighting0(LIGHT_SPECULAR, 1), lighting0(LIGHT_SPECULAR, 2), lighting0(LIGHT_SPECULAR, 3))

    renderState = getCurrentRenderState()
    If renderState.fogEnabled Then fogColor = pielightToRGBAVec4(renderState.fogColour) ' else glm::vec4(0.f)

    If imr_useInstancedRendering Then

        bucketLight = ILightingManager_getPointLightBuckets(getCurrentLightingManager())
        gfx_api_context_getDrawableDimensions dimW, dimH
        ' (VB) the aggregate initialiser's fields, in the struct's order
        globalUniformsInst.ProjectionMatrix = projectionMatrix
        globalUniformsInst.ViewMatrix = viewMatrix
        globalUniformsInst.ModelUVLightmapMatrix = imr_modelUVLightmapMatrix
        globalUniformsInst.ShadowMapMVPMatrix(0) = shadowCascades.shadowMVPMatrix(0)
        globalUniformsInst.ShadowMapMVPMatrix(1) = shadowCascades.shadowMVPMatrix(1)
        globalUniformsInst.ShadowMapMVPMatrix(2) = shadowCascades.shadowMVPMatrix(2)
        globalUniformsInst.cameraPos = glm_vec4_v3(cameraPos, 0!)
        globalUniformsInst.sunPos = glm_vec4_v3(currentSunPosition, 0!)
        globalUniformsInst.sceneColor = sceneColor
        globalUniformsInst.ambient = ambient
        globalUniformsInst.diffuse = diffuse
        globalUniformsInst.specular = specular
        globalUniformsInst.fogColour = fogColor
        globalUniformsInst.ShadowMapCascadeSplits = glm_vec4(shadowCascades.shadowCascadeSplit(0), shadowCascades.shadowCascadeSplit(1), shadowCascades.shadowCascadeSplit(2), pie_getPerspectiveZFar())
        globalUniformsInst.ShadowMapSize = shadowCascades.shadowMapSize
        ' (C++ gives fogBegin to the struct's fogEnd and fogEnd to its fogBegin: the initialiser's order is not the struct's)
        globalUniformsInst.fogEnd = renderState.fogBegin
        globalUniformsInst.fogBegin = renderState.fogEnd
        globalUniformsInst.timeState = pie_GetShaderTime()
        globalUniformsInst.fogEnabled = boolToInt(renderState.fogEnabled)
        globalUniformsInst.viewportWidth = dimW
        globalUniformsInst.viewportheight = dimH
        globalUniformsInst.unused2 = 0!
        CopyMemory globalUniformsInst.PointLightsPosition(0), bucketLight.positions(0), gfx_api_max_lights * 16
        CopyMemory globalUniformsInst.PointLightsColorAndEnergy(0), bucketLight.colorAndEnergy(0), gfx_api_max_lights * 16
        CopyMemory globalUniformsInst.bucketOffsetAndSize(0), bucketLight.bucketOffsetAndSize(0), gfx_api_bucket_dimension * gfx_api_bucket_dimension * 16
        CopyMemory globalUniformsInst.indexed_lights(0), bucketLight.light_index(0), gfx_api_max_indexed_lights * 16
        globalUniformsInst.bucketDimensionUsed = bucketLight.bucketDimensionUsed
        InstancedMeshRenderer_Draw3DShapes_Instanced currentGameFrame, perFrameUniformsShaderOnce, globalUniformsInst, drawParts, depthPass
    Else
        globalUniformsOld.ProjectionMatrix = projectionMatrix
        globalUniformsOld.ViewMatrix = viewMatrix
        globalUniformsOld.ShadowMapMVPMatrix = shadowCascades.shadowMVPMatrix(0)
        globalUniformsOld.sunPos = glm_vec4_v3(currentSunPosition, 0!)
        globalUniformsOld.sceneColor = sceneColor
        globalUniformsOld.ambient = ambient
        globalUniformsOld.diffuse = diffuse
        globalUniformsOld.specular = specular
        globalUniformsOld.fogColour = fogColor
        ' (the same order as above)
        globalUniformsOld.fogEnd = renderState.fogBegin
        globalUniformsOld.fogBegin = renderState.fogEnd
        globalUniformsOld.timeState = pie_GetShaderTime()
        globalUniformsOld.fogEnabled = boolToInt(renderState.fogEnabled)
        InstancedMeshRenderer_Draw3DShapes_Old currentGameFrame, perFrameUniformsShaderOnce, globalUniformsOld, drawParts
    End If

    InstancedMeshRenderer_DrawAll = True
End Function

' template<SHADER_MODE shader, typename Draw3DInstancedPSO>
Private Sub drawInstanced3dShapeTemplated_Inner(ByVal shader As SHADER_MODE, ByVal Draw3DInstancedPSO As pso_type, globalsOnce As ShaderOnce, globalUniforms As Draw3DShapeInstancedGlobalUniforms, ByVal shape As Long, ByVal instanceDataBuffer As gfx_api_buffer, ByVal instanceBufferOffset As Long, ByVal instance_count As Long, ByVal lightmapTexture As gfx_api_texture, ByVal shieldEffect As Boolean)
    Dim textures As iIMDShapeTextures, tcmask As gfx_api_texture, normalmap As gfx_api_texture, specularmap As gfx_api_texture
    Dim meshUniforms As Draw3DShapeInstancedPerMeshUniforms, pTangentBuffer As gfx_api_buffer, pso As pipeline_state_helper
    Dim buffers(0 To 4) As gfx_api_buffer, offsets(0 To 4) As Long
    textures = iIMDShape_getTextures(shape)
    Set tcmask = pageTextureOrNothing(textures.tcmaskpage)
    Set normalmap = pageTextureOrNothing(textures.normalpage)
    Set specularmap = pageTextureOrNothing(textures.specularpage)

    meshUniforms.tcmask = boolToInt(Not tcmask Is Nothing)
    meshUniforms.normalMap = boolToInt(Not normalmap Is Nothing)
    meshUniforms.specularMap = boolToInt(Not specularmap Is Nothing)
    meshUniforms.hasTangents = boolToInt(Not imdShapes(shape).buffers(VBO_TANGENT) Is Nothing)
    meshUniforms.shieldEffect = boolToInt(shieldEffect)

    If Not imdShapes(shape).buffers(VBO_TANGENT) Is Nothing Then
        Set pTangentBuffer = imdShapes(shape).buffers(VBO_TANGENT)
    Else
        Set pTangentBuffer = getZeroedVertexBuffer(imdShapes(shape).vertexCount * 4 * 4)
    End If

    Set pso = gfx_api_pso_get(Draw3DInstancedPSO)
    pso.bind
    gfx_api_context_bind_index_buffer imdShapes(shape).buffers(VBO_INDEX), index_type_u16
    If ShaderOnce_perform_once(globalsOnce, Draw3DInstancedPSO) Then
        pso.set_uniforms_at 0, VarPtr(globalUniforms), LenB(globalUniforms)
    End If

    pso.set_uniforms_at 1, VarPtr(meshUniforms), LenB(meshUniforms)
    Set buffers(0) = imdShapes(shape).buffers(VBO_VERTEX)
    Set buffers(1) = imdShapes(shape).buffers(VBO_NORMAL)
    Set buffers(2) = imdShapes(shape).buffers(VBO_TEXCOORD)
    Set buffers(3) = pTangentBuffer
    Set buffers(4) = instanceDataBuffer
    offsets(4) = instanceBufferOffset
    gfx_api_context_bind_vertex_buffers 0, buffers, offsets, 5
    pso.bind_textures pie_Texture(textures.texpage), tcmask, normalmap, specularmap, gfx_api_context_getDepthTexture(), lightmapTexture

    pso.draw_elements_instanced imdShapes(shape).polys_size * 3, 0, instance_count
'   Draw3DInstancedPSO::get().unbind_vertex_buffers(shape->buffers[VBO_VERTEX], shape->buffers[VBO_NORMAL], shape->buffers[VBO_TEXCOORD]);
End Sub

Private Sub drawInstanced3dShapeDepthOnly(globalsOnce As ShaderOnce, globalUniforms As Draw3DShapeInstancedGlobalUniforms, ByVal shape As Long, ByVal pieFlag As Long, ByVal instanceDataBuffer As gfx_api_buffer, ByVal instanceBufferOffset As Long, ByVal instance_count As Long)
    Dim pso As pipeline_state_helper, depthGlobals As Draw3DShapeInstancedDepthOnlyGlobalUniforms
    Dim buffers(0 To 2) As gfx_api_buffer, offsets(0 To 2) As Long
    If (pieFlag And pie_NODEPTHWRITE) <> 0 Then
        Exit Sub
    End If

    Set pso = gfx_api_pso_get(Draw3DShapeDepthOnly_Instanced)
    pso.bind
    gfx_api_context_bind_index_buffer imdShapes(shape).buffers(VBO_INDEX), index_type_u16
    If ShaderOnce_perform_once(globalsOnce, Draw3DShapeDepthOnly_Instanced) Then
        depthGlobals.ProjectionMatrix = globalUniforms.ProjectionMatrix
        depthGlobals.ViewMatrix = globalUniforms.ViewMatrix
        pso.set_uniforms_at 0, VarPtr(depthGlobals), LenB(depthGlobals)
    End If

'	gfx_api::Draw3DShapeDepthOnly_Instanced::get().set_uniforms_at(1, meshUniforms);
    Set buffers(0) = imdShapes(shape).buffers(VBO_VERTEX)
    Set buffers(1) = imdShapes(shape).buffers(VBO_NORMAL)
'		std::make_tuple(shape->buffers[VBO_TEXCOORD], 0),
'		std::make_tuple(pTangentBuffer, 0),
    Set buffers(2) = instanceDataBuffer
    offsets(2) = instanceBufferOffset
    gfx_api_context_bind_vertex_buffers 0, buffers, offsets, 3
'	gfx_api::Draw3DShapeDepthOnly_Instanced::get().bind_textures(&pie_Texture(shape->texpage), tcmask, normalmap, specularmap);

    pso.draw_elements_instanced imdShapes(shape).polys_size * 3, 0, instance_count
'	Draw3DInstancedPSO::get().unbind_vertex_buffers(shape->buffers[VBO_VERTEX], shape->buffers[VBO_NORMAL], shape->buffers[VBO_TEXCOORD]);
End Sub

' template<SHADER_MODE shader, typename AdditivePSO, typename AdditiveNoDepthWRTPSO, typename AlphaPSO, typename AlphaNoDepthWRTPSO,
'          typename PremultipliedPSO, typename PremultipliedNoDepthWRTPSO, typename OpaquePSO>
Private Sub drawInstanced3dShapeTemplated(ByVal shader As SHADER_MODE, ByVal AdditivePSO As pso_type, ByVal AdditiveNoDepthWRTPSO As pso_type, ByVal AlphaPSO As pso_type, ByVal AlphaNoDepthWRTPSO As pso_type, ByVal PremultipliedPSO As pso_type, ByVal PremultipliedNoDepthWRTPSO As pso_type, ByVal OpaquePSO As pso_type, globalsOnce As ShaderOnce, globalUniforms As Draw3DShapeInstancedGlobalUniforms, ByVal shape As Long, ByVal pieFlag As Long, ByVal instanceDataBuffer As gfx_api_buffer, ByVal instanceBufferOffset As Long, ByVal instance_count As Long, ByVal lightmapTexture As gfx_api_texture)
    Dim shieldEffect As Boolean
    shieldEffect = (pieFlag And pie_SHIELD) <> 0

    ' Set tranlucency
    If (pieFlag And pie_ADDITIVE) <> 0 Then
        If (pieFlag And pie_NODEPTHWRITE) = 0 Then
            drawInstanced3dShapeTemplated_Inner shader, AdditivePSO, globalsOnce, globalUniforms, shape, instanceDataBuffer, instanceBufferOffset, instance_count, lightmapTexture, shieldEffect
        Else
            drawInstanced3dShapeTemplated_Inner shader, AdditiveNoDepthWRTPSO, globalsOnce, globalUniforms, shape, instanceDataBuffer, instanceBufferOffset, instance_count, lightmapTexture, shieldEffect
        End If
    ElseIf (pieFlag And pie_TRANSLUCENT) <> 0 Then
        If (pieFlag And pie_NODEPTHWRITE) = 0 Then
            drawInstanced3dShapeTemplated_Inner shader, AlphaPSO, globalsOnce, globalUniforms, shape, instanceDataBuffer, instanceBufferOffset, instance_count, lightmapTexture, shieldEffect
        Else
            drawInstanced3dShapeTemplated_Inner shader, AlphaNoDepthWRTPSO, globalsOnce, globalUniforms, shape, instanceDataBuffer, instanceBufferOffset, instance_count, lightmapTexture, shieldEffect
        End If
    ElseIf (pieFlag And pie_PREMULTIPLIED) <> 0 Then
        If (pieFlag And pie_NODEPTHWRITE) = 0 Then
            drawInstanced3dShapeTemplated_Inner shader, PremultipliedPSO, globalsOnce, globalUniforms, shape, instanceDataBuffer, instanceBufferOffset, instance_count, lightmapTexture, shieldEffect
        Else
            drawInstanced3dShapeTemplated_Inner shader, PremultipliedNoDepthWRTPSO, globalsOnce, globalUniforms, shape, instanceDataBuffer, instanceBufferOffset, instance_count, lightmapTexture, shieldEffect
        End If
    Else
        drawInstanced3dShapeTemplated_Inner shader, OpaquePSO, globalsOnce, globalUniforms, shape, instanceDataBuffer, instanceBufferOffset, instance_count, lightmapTexture, shieldEffect
    End If
End Sub

Private Sub pie_Draw3DShape2_Instanced(globalsOnce As ShaderOnce, globalUniforms As Draw3DShapeInstancedGlobalUniforms, ByVal shape As Long, ByVal pieFlag As Long, ByVal instanceDataBuffer As gfx_api_buffer, ByVal instanceBufferOffset As Long, ByVal instance_count As Long, ByVal depthPass As Boolean, ByVal lightmapTexture As gfx_api_texture)
    Dim light_ As Boolean
    light_ = True

    drawCallsCount = drawCallsCount + 1

    ' Set fog status
    If (pieFlag And pie_FORCE_FOG) = 0 And ((pieFlag And pie_ADDITIVE) <> 0 Or (pieFlag And pie_TRANSLUCENT) <> 0 Or (pieFlag And pie_PREMULTIPLIED) <> 0) Then
        pie_SetFogStatus False
    Else
        pie_SetFogStatus True
    End If

    ' Set translucency
    If (pieFlag And pie_ADDITIVE) <> 0 Then
        light_ = False
    ElseIf (pieFlag And pie_TRANSLUCENT) <> 0 Then
        light_ = False
    ElseIf (pieFlag And pie_PREMULTIPLIED) <> 0 Then
        light_ = False
    End If

    If (pieFlag And pie_ECM) <> 0 Then
        light_ = True
    End If

    If (pieFlag And pie_FORCELIGHT) <> 0 Then
        light_ = True
    End If

'	gfx_api::context::get().bind_index_buffer(*shape->buffers[VBO_INDEX], gfx_api::index_type::u16);

    If depthPass Then
        drawInstanced3dShapeDepthOnly globalsOnce, globalUniforms, shape, pieFlag, instanceDataBuffer, instanceBufferOffset, instance_count
    ElseIf light_ Then
        drawInstanced3dShapeTemplated SHADER_COMPONENT_INSTANCED, Draw3DShapeAdditive_Instanced, Draw3DShapeAdditiveNoDepthWRT_Instanced, Draw3DShapeAlpha_Instanced, Draw3DShapeAlphaNoDepthWRT_Instanced, Draw3DShapePremul_Instanced, Draw3DShapePremulNoDepthWRT_Instanced, Draw3DShapeOpaque_Instanced, globalsOnce, globalUniforms, shape, pieFlag, instanceDataBuffer, instanceBufferOffset, instance_count, lightmapTexture
    Else
        drawInstanced3dShapeTemplated SHADER_NOLIGHT_INSTANCED, Draw3DShapeNoLightAdditive_Instanced, Draw3DShapeNoLightAdditiveNoDepthWRT_Instanced, Draw3DShapeNoLightAlpha_Instanced, Draw3DShapeNoLightAlphaNoDepthWRT_Instanced, Draw3DShapeNoLightPremul_Instanced, Draw3DShapeNoLightPremulNoDepthWRT_Instanced, Draw3DShapeNoLightOpaque_Instanced, globalsOnce, globalUniforms, shape, pieFlag, instanceDataBuffer, instanceBufferOffset, instance_count, lightmapTexture
    End If

    polyCount = polyCount + imdShapes(shape).polys_size
End Sub

' (VB) the instance offset of a finalized draw call: sizeof(Draw3DShapePerInstanceInterleavedData) * startingIdxInInstancesBuffer
Private Function instanceOffsetOf(ByVal i As Long) As Long
    Dim d As Draw3DShapePerInstanceInterleavedData
    instanceOffsetOf = LenB(d) * imr_finalizedDrawCalls(i).startingIdxInInstancesBuffer
End Function

' New, instanced rendering
Private Sub InstancedMeshRenderer_Draw3DShapes_Instanced(ByVal currentGameFrame As Double, globalsOnce As ShaderOnce, globalUniforms As Draw3DShapeInstancedGlobalUniforms, ByVal drawParts As Long, ByVal depthPass As Boolean)
    Dim i As Long, shape As Long, pieFlag As Long, instanceBufferOffset As Long
    If imr_finalizedDrawCalls_size = 0 Then
        Exit Sub
    End If

    If (drawParts And DrawParts_ShadowCastingShapes) = DrawParts_ShadowCastingShapes Then
        ' Draw opaque models
        gfx_api_context_debugStringMarker "Remaining passes - opaque models"
        For i = 0 To imr_startIdxTranslucentDrawCalls - 1
            shape = imr_finalizedDrawCalls(i).state.shape
            pieFlag = imr_finalizedDrawCalls(i).state.pieFlag
            If depthPass And Not ((pieFlag And pie_SHADOW) <> 0 Or (pieFlag And pie_STATIC_SHADOW) <> 0) Then
                GoTo continue_opaque
            End If
            instanceBufferOffset = instanceOffsetOf(i)
            pie_Draw3DShape2_Instanced perFrameUniformsShaderOnce, globalUniforms, shape, imr_finalizedDrawCalls(i).state.pieFlag, imr_instanceDataBuffers(imr_currInstanceBufferIdx), instanceBufferOffset, imr_finalizedDrawCalls(i).instance_count, depthPass, imr_lightmapTexture
continue_opaque:
        Next
        If imr_startIdxTranslucentDrawCalls > 0 Then
            ' unbind last index buffer bound inside pie_Draw3DShape2
            gfx_api_context_unbind_index_buffer imdShapes(imr_finalizedDrawCalls(imr_startIdxTranslucentDrawCalls - 1).state.shape).buffers(VBO_INDEX)
        End If
    End If

    If (drawParts And DrawParts_OldShadows) = DrawParts_OldShadows Then
        ' Draw shadows
        gfx_api_context_debugStringMarker "Remaining passes - shadows"
        If shadows Then
            pie_DrawShadows currentGameFrame, globalUniforms.ProjectionMatrix
        End If
    End If

    If (drawParts And DrawParts_TranslucentShapes) = DrawParts_TranslucentShapes Then
        ' Draw translucent models next
        gfx_api_context_debugStringMarker "Remaining passes - translucent models"
        For i = imr_startIdxTranslucentDrawCalls To imr_startIdxTranslucentNoDepthWriteDrawCalls - 1
            shape = imr_finalizedDrawCalls(i).state.shape
            instanceBufferOffset = instanceOffsetOf(i)
            pie_Draw3DShape2_Instanced perFrameUniformsShaderOnce, globalUniforms, shape, imr_finalizedDrawCalls(i).state.pieFlag, imr_instanceDataBuffers(imr_currInstanceBufferIdx), instanceBufferOffset, imr_finalizedDrawCalls(i).instance_count, depthPass, imr_lightmapTexture
        Next
        If imr_startIdxTranslucentDrawCalls < imr_startIdxTranslucentNoDepthWriteDrawCalls Then
            ' unbind last index buffer bound inside pie_Draw3DShape2
            gfx_api_context_unbind_index_buffer imdShapes(imr_finalizedDrawCalls(imr_startIdxTranslucentNoDepthWriteDrawCalls - 1).state.shape).buffers(VBO_INDEX)
        End If

        ' Draw translucent (no depth write) models after
        gfx_api_context_debugStringMarker "Remaining passes - translucent (no depth write) models"
        For i = imr_startIdxTranslucentNoDepthWriteDrawCalls To imr_startIdxAdditiveDrawCalls - 1
            shape = imr_finalizedDrawCalls(i).state.shape
            instanceBufferOffset = instanceOffsetOf(i)
            pie_Draw3DShape2_Instanced perFrameUniformsShaderOnce, globalUniforms, shape, imr_finalizedDrawCalls(i).state.pieFlag, imr_instanceDataBuffers(imr_currInstanceBufferIdx), instanceBufferOffset, imr_finalizedDrawCalls(i).instance_count, depthPass, imr_lightmapTexture
        Next
        If imr_startIdxTranslucentNoDepthWriteDrawCalls < imr_startIdxAdditiveDrawCalls Then
            ' unbind last index buffer bound inside pie_Draw3DShape2
            gfx_api_context_unbind_index_buffer imdShapes(imr_finalizedDrawCalls(imr_startIdxAdditiveDrawCalls - 1).state.shape).buffers(VBO_INDEX)
        End If
    End If

    If (drawParts And DrawParts_AdditiveShapes) = DrawParts_AdditiveShapes Then
        ' Draw additive models last
        gfx_api_context_debugStringMarker "Remaining passes - additive models"
        For i = imr_startIdxAdditiveDrawCalls To imr_finalizedDrawCalls_size - 1
            shape = imr_finalizedDrawCalls(i).state.shape
            instanceBufferOffset = instanceOffsetOf(i)
            pie_Draw3DShape2_Instanced perFrameUniformsShaderOnce, globalUniforms, shape, imr_finalizedDrawCalls(i).state.pieFlag, imr_instanceDataBuffers(imr_currInstanceBufferIdx), instanceBufferOffset, imr_finalizedDrawCalls(i).instance_count, depthPass, imr_lightmapTexture
        Next
        If imr_startIdxAdditiveDrawCalls < imr_finalizedDrawCalls_size Then
            ' unbind last index buffer bound inside pie_Draw3DShape2
            gfx_api_context_unbind_index_buffer imdShapes(imr_finalizedDrawCalls(imr_finalizedDrawCalls_size - 1).state.shape).buffers(VBO_INDEX)
        End If
    End If

    If drawParts <> 0 Then
        gfx_api_context_disable_all_vertex_buffers
    End If

    gfx_api_context_debugStringMarker "Remaining passes - done"
End Sub

' (VB) std::sort(shapes.begin(), shapes.end(), less_than_shape()): by the shape pointer
Private Sub sortShapesByShape(v() As SHAPE_, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Long, t As SHAPE_
    Do While hi - lo > 12
        pivot = v((lo + hi) \ 2).shape
        i = lo
        j = hi
        Do While i <= j
            Do While v(i).shape < pivot
                i = i + 1
            Loop
            Do While v(j).shape > pivot
                j = j - 1
            Loop
            If i <= j Then
                t = v(i): v(i) = v(j): v(j) = t
                i = i + 1
                j = j - 1
            End If
        Loop
        If j - lo < hi - i Then
            If lo < j Then sortShapesByShape v, lo, j
            lo = i
        Else
            If i < hi Then sortShapesByShape v, i, hi
            hi = j
        End If
    Loop
    For i = lo + 1 To hi
        t = v(i)
        j = i - 1
        Do While j >= lo
            If v(j).shape <= t.shape Then Exit Do
            v(j + 1) = v(j)
            j = j - 1
        Loop
        v(j + 1) = t
    Next
End Sub

' Old, non-instanced rendering
Private Sub InstancedMeshRenderer_Draw3DShapes_Old(ByVal currentGameFrame As Double, globalsOnce As ShaderOnce, globalUniforms As Draw3DShapeGlobalUniforms, ByVal drawParts As Long)
    Dim lastState As templatedState, blank As templatedState, i As Long
    If (drawParts And DrawParts_ShadowCastingShapes) = DrawParts_ShadowCastingShapes Then
        ' Draw models
        ' sort list to reduce state changes
        If imr_shapes_size > 1 Then sortShapesByShape imr_shapes, 0, imr_shapes_size - 1
        gfx_api_context_debugStringMarker "Remaining passes - opaque models"
        For i = 0 To imr_shapes_size - 1
            lastState = pie_Draw3DShape2(lastState, perFrameUniformsShaderOnce, globalUniforms, imr_shapes(i).shape, imr_shapes(i).frame, imr_shapes(i).colour, imr_shapes(i).teamcolour, imr_shapes(i).flag, imr_shapes(i).flag_data, imr_shapes(i).modelMatrix, imr_shapes(i).stretch)
        Next
        gfx_api_context_disable_all_vertex_buffers
        If imr_shapes_size > 0 Then
            ' unbind last index buffer bound inside pie_Draw3DShape2
            gfx_api_context_unbind_index_buffer imdShapes(imr_shapes(imr_shapes_size - 1).shape).buffers(VBO_INDEX)
        End If
    End If

    If (drawParts And DrawParts_OldShadows) = DrawParts_OldShadows Then
        gfx_api_context_debugStringMarker "Remaining passes - shadows"
        ' Draw shadows
        If shadows Then
            pie_DrawShadows currentGameFrame, globalUniforms.ProjectionMatrix
        End If
    End If

    If (drawParts And DrawParts_TranslucentShapes) = DrawParts_TranslucentShapes Then
        ' Draw translucent models last
        ' TODO, sort list by Z order to do translucency correctly
        gfx_api_context_debugStringMarker "Remaining passes - translucent models"
        lastState = blank               ' templatedState()
        For i = 0 To imr_tshapes_size - 1
            lastState = pie_Draw3DShape2(lastState, perFrameUniformsShaderOnce, globalUniforms, imr_tshapes(i).shape, imr_tshapes(i).frame, imr_tshapes(i).colour, imr_tshapes(i).teamcolour, imr_tshapes(i).flag, imr_tshapes(i).flag_data, imr_tshapes(i).modelMatrix, imr_tshapes(i).stretch)
        Next
        gfx_api_context_disable_all_vertex_buffers
        If imr_tshapes_size > 0 Then
            ' unbind last index buffer bound inside pie_Draw3DShape2
            gfx_api_context_unbind_index_buffer imdShapes(imr_tshapes(imr_tshapes_size - 1).shape).buffers(VBO_INDEX)
        End If
    End If

    gfx_api_context_debugStringMarker "Remaining passes - done"
End Sub

' void pie_GetResetCounts(size_t *pPieCount, size_t *pPolyCount)
Public Sub pie_GetResetCounts(pPieCount As Long, pPolyCount As Long)
    pPieCount = pieCount
    pPolyCount = polyCount

    pieCount = 0
    polyCount = 0
    drawCallsCount = 0
End Sub
