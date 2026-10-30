VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gl_pipeline_state_object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/gfx_api_gl.cpp: gl_pipeline_state_object (gfx_api_gl.h), a shader program and
' its fixed state. The constructor is construct (gfx_api_gl_c's gl_pipeline_state_object_new);
' the gl_context& it takes is gfx_api_gl_c itself (its gles flag is a parameter). The uniform
' bind functions (std::function, picked by the uniform struct's std::type_index) are the
' uniform_block_type of each block (-1 = nullptr): set_constants copies the block from its
' address into a local of its type and calls that type's set_constants_<type>. setUniforms'
' overloads carry what they take: setUniforms_vec4, setUniforms_mat4, setUniforms_mat4_count (an
' address and a count), ... std::regex is VBScript.RegExp (late bound; ECMAScript-like).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private desc As state_description
Public program As Long                  ' GLuint program = 0
Public vertexShader As Long             ' = 0
Public fragmentShader As Long           ' = 0
Private vertex_buffer_desc() As vertex_buffer
Private vertex_buffer_desc_size As Long
Private locations() As Long
Private locations_size As Long
Private duplicateFragmentUniformLocations() As Long
Private duplicateFragmentUniformLocations_size As Long
Public hasSpecializationConstant_ShadowConstants As Boolean ' = false
Public hasSpecializationConstants_PointLights As Boolean    ' = false
Private uniform_bind_functions() As Long ' uniform_block_type, -1 = nullptr
Private uniform_bind_functions_size As Long
Public broken As Boolean                ' (gfx_api::pipeline_state_object) = false

Private Const glsl_whitespace_chars As String = " " & vbTab & vbVerticalTab & vbFormFeed & vbCr & vbLf

' (VB) current_program->vertex_buffer_desc[i] and .size(), for gl_context
Friend Function vertex_buffer_desc_at(ByVal i As Long) As vertex_buffer
    vertex_buffer_desc_at = vertex_buffer_desc(i)
End Function

Friend Function vertex_buffer_desc_count() As Long
    vertex_buffer_desc_count = vertex_buffer_desc_size
End Function

' gl_pipeline_state_object(gl_context& ctx, bool fragmentHighpFloatAvailable, bool fragmentHighpIntAvailable, bool patchFragmentShaderMipLodBias, const gfx_api::pipeline_create_info& createInfo, optional<float> mipLodBias, const gfx_api::lighting_constants& shadowConstants)
Friend Sub construct(ByVal ctx_gles As Boolean, ByVal fragmentHighpFloatAvailable As Boolean, ByVal fragmentHighpIntAvailable As Boolean, ByVal patchFragmentShaderMipLodBias As Boolean, createInfo As pipeline_create_info, ByVal has_mipLodBias As Boolean, ByVal mipLodBias As Single, shadowConstants As lighting_constants)
    Dim vertexShaderHeader As String, fragmentShaderHeader As String, shaderVersionStr As String
    Dim pd As program_data, i As Long, uniform_block As Long
    desc = createInfo.state_desc
    vertex_buffer_desc_size = createInfo.attribute_descriptions_size
    If vertex_buffer_desc_size > 0 Then
        ReDim vertex_buffer_desc(0 To vertex_buffer_desc_size - 1)
        For i = 0 To vertex_buffer_desc_size - 1
            vertex_buffer_desc(i) = createInfo.attribute_descriptions(i)
        Next
    End If

    If Not ctx_gles Then
        ' Determine the shader version directive we should use by examining the current OpenGL context
        ' (The built-in shaders support (and have been tested with) VERSION_120, VERSION_150_CORE, VERSION_330_CORE)
        shaderVersionStr = shaderVersionString(getMaximumShaderVersionForCurrentGLContext(VERSION_120, VERSION_330_CORE))

        vertexShaderHeader = shaderVersionStr
        fragmentShaderHeader = shaderVersionStr
    Else
        ' Determine the shader version directive we should use by examining the current OpenGL ES context
        ' (The built-in shaders support (and have been tested with) VERSION_ES_100 and VERSION_ES_300)
        shaderVersionStr = shaderVersionString_es(getMaximumShaderVersionForCurrentGLESContext(VERSION_ES_100, VERSION_ES_300))

        vertexShaderHeader = shaderVersionStr
        fragmentShaderHeader = shaderVersionStr
        ' OpenGL ES Shading Language - 4. Variables and Types - pp. 35-36
        ' https://www.khronos.org/registry/gles/specs/2.0/GLSL_ES_Specification_1.0.17.pdf?#page=41
        '
        ' > The fragment language has no default precision qualifier for floating point types.
        ' > Hence for float, floating point vector and matrix variable declarations, either the
        ' > declaration must include a precision qualifier or the default float precision must
        ' > have been previously declared.
        fragmentShaderHeader = fragmentShaderHeader & "#if GL_FRAGMENT_PRECISION_HIGH" & vbLf & "precision highp float;" & vbLf & "precision highp int;" & vbLf & "#else" & vbLf & "precision mediump float;" & vbLf & "#endif" & vbLf
        fragmentShaderHeader = fragmentShaderHeader & "#if __VERSION__ >= 300 || defined(GL_EXT_texture_array)" & vbLf & "precision lowp sampler2DArray;" & vbLf & "#endif" & vbLf
        fragmentShaderHeader = fragmentShaderHeader & "#if __VERSION__ >= 300" & vbLf & "precision lowp sampler2DShadow;" & vbLf & "precision lowp sampler2DArrayShadow;" & vbLf & "#endif" & vbLf
    End If

    ' shader_to_file_table.at(createInfo.shader_mode)
    shader_to_file_table_at createInfo.shader_mode, pd
    build_program fragmentHighpFloatAvailable, fragmentHighpIntAvailable, patchFragmentShaderMipLodBias, _
                  pd.friendly_name, _
                  vertexShaderHeader, _
                  pd.vertex_file, _
                  fragmentShaderHeader, _
                  pd.fragment_file, _
                  pd.uniform_names, _
                  pd, _
                  has_mipLodBias, mipLodBias, shadowConstants

    ' uniforms_bind_table: every uniform_block_type has its set_constants
    uniform_bind_functions_size = createInfo.uniform_blocks_size
    If uniform_bind_functions_size > 0 Then ReDim uniform_bind_functions(0 To uniform_bind_functions_size - 1)
    For i = 0 To createInfo.uniform_blocks_size - 1
        uniform_block = createInfo.uniform_blocks(i)
        If uniform_block < ub_Draw3DShapeGlobalUniforms Or uniform_block > ub_SHADER_WORLD_TO_SCREEN Then
            ASSERT_FAILURE "false", "lib/ivis_opengl/gfx_api_gl.cpp:1236", "gl_pipeline_state_object", "Missing mapping for uniform block type: %s", CStr(uniform_block)
            uniform_bind_functions(i) = -1
            GoTo continue_
        End If
        uniform_bind_functions(i) = uniform_block
continue_:
    Next
End Sub

' the uniform bind function of uniform_block (C++'s lambdas of uniform_binding_entry / uniform_setting_func)
Private Sub uniform_bind(ByVal uniform_block As Long, ByVal buffer As Long, ByVal buflen As Long)
    Select Case uniform_block
    Case ub_Draw3DShapeGlobalUniforms
        Static c0 As Draw3DShapeGlobalUniforms
        If Not (buflen = LenB(c0)) Then ASSERT_FAILURE "buflen == sizeof(const T)", "lib/ivis_opengl/gfx_api_gl.cpp:1147", "uniform_setting_func", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c0): Exit Sub
        CopyMemory c0, ByVal buffer, buflen
        set_constants_Draw3DShapeGlobalUniforms c0
    Case ub_Draw3DShapePerMeshUniforms
        Static c1 As Draw3DShapePerMeshUniforms
        If Not (buflen = LenB(c1)) Then ASSERT_FAILURE "buflen == sizeof(const T)", "lib/ivis_opengl/gfx_api_gl.cpp:1147", "uniform_setting_func", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c1): Exit Sub
        CopyMemory c1, ByVal buffer, buflen
        set_constants_Draw3DShapePerMeshUniforms c1
    Case ub_Draw3DShapePerInstanceUniforms
        Static c2 As Draw3DShapePerInstanceUniforms
        If Not (buflen = LenB(c2)) Then ASSERT_FAILURE "buflen == sizeof(const T)", "lib/ivis_opengl/gfx_api_gl.cpp:1147", "uniform_setting_func", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c2): Exit Sub
        CopyMemory c2, ByVal buffer, buflen
        set_constants_Draw3DShapePerInstanceUniforms c2
    Case ub_Draw3DShapeInstancedGlobalUniforms
        Static c3 As Draw3DShapeInstancedGlobalUniforms
        If Not (buflen = LenB(c3)) Then ASSERT_FAILURE "buflen == sizeof(const T)", "lib/ivis_opengl/gfx_api_gl.cpp:1147", "uniform_setting_func", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c3): Exit Sub
        CopyMemory c3, ByVal buffer, buflen
        set_constants_Draw3DShapeInstancedGlobalUniforms c3
    Case ub_Draw3DShapeInstancedPerMeshUniforms
        Static c4 As Draw3DShapeInstancedPerMeshUniforms
        If Not (buflen = LenB(c4)) Then ASSERT_FAILURE "buflen == sizeof(const T)", "lib/ivis_opengl/gfx_api_gl.cpp:1147", "uniform_setting_func", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c4): Exit Sub
        CopyMemory c4, ByVal buffer, buflen
        set_constants_Draw3DShapeInstancedPerMeshUniforms c4
    Case ub_Draw3DShapeInstancedDepthOnlyGlobalUniforms
        Static c5 As Draw3DShapeInstancedDepthOnlyGlobalUniforms
        If Not (buflen = LenB(c5)) Then ASSERT_FAILURE "buflen == sizeof(const T)", "lib/ivis_opengl/gfx_api_gl.cpp:1147", "uniform_setting_func", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c5): Exit Sub
        CopyMemory c5, ByVal buffer, buflen
        set_constants_Draw3DShapeInstancedDepthOnlyGlobalUniforms c5
    Case ub_SHADER_TERRAIN_DEPTH
        Static c6 As constant_buffer_type_SHADER_TERRAIN_DEPTH
        If Not (buflen = LenB(c6)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c6): Exit Sub
        CopyMemory c6, ByVal buffer, buflen
        set_constants_SHADER_TERRAIN_DEPTH c6
    Case ub_SHADER_TERRAIN_DEPTHMAP
        Static c7 As constant_buffer_type_SHADER_TERRAIN_DEPTHMAP
        If Not (buflen = LenB(c7)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c7): Exit Sub
        CopyMemory c7, ByVal buffer, buflen
        set_constants_SHADER_TERRAIN_DEPTHMAP c7
    Case ub_TerrainCombinedUniforms
        Static c8 As TerrainCombinedUniforms
        If Not (buflen = LenB(c8)) Then ASSERT_FAILURE "buflen == sizeof(const T)", "lib/ivis_opengl/gfx_api_gl.cpp:1147", "uniform_setting_func", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c8): Exit Sub
        CopyMemory c8, ByVal buffer, buflen
        set_constants_TerrainCombinedUniforms c8
    Case ub_SHADER_WATER
        Static c9 As constant_buffer_type_SHADER_WATER
        If Not (buflen = LenB(c9)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c9): Exit Sub
        CopyMemory c9, ByVal buffer, buflen
        set_constants_SHADER_WATER c9
    Case ub_SHADER_WATER_HIGH
        Static c10 As constant_buffer_type_SHADER_WATER_HIGH
        If Not (buflen = LenB(c10)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c10): Exit Sub
        CopyMemory c10, ByVal buffer, buflen
        set_constants_SHADER_WATER_HIGH c10
    Case ub_SHADER_WATER_CLASSIC
        Static c11 As constant_buffer_type_SHADER_WATER_CLASSIC
        If Not (buflen = LenB(c11)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c11): Exit Sub
        CopyMemory c11, ByVal buffer, buflen
        set_constants_SHADER_WATER_CLASSIC c11
    Case ub_SHADER_RECT
        Static c12 As constant_buffer_type_SHADER_RECT
        If Not (buflen = LenB(c12)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c12): Exit Sub
        CopyMemory c12, ByVal buffer, buflen
        set_constants_SHADER_RECT c12
    Case ub_SHADER_TEXRECT
        Static c13 As constant_buffer_type_SHADER_TEXRECT
        If Not (buflen = LenB(c13)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c13): Exit Sub
        CopyMemory c13, ByVal buffer, buflen
        set_constants_SHADER_TEXRECT c13
    Case ub_SHADER_GFX_COLOUR
        Static c14 As constant_buffer_type_SHADER_GFX_COLOUR
        If Not (buflen = LenB(c14)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c14): Exit Sub
        CopyMemory c14, ByVal buffer, buflen
        set_constants_SHADER_GFX_COLOUR c14
    Case ub_SHADER_GFX_TEXT
        Static c15 As constant_buffer_type_SHADER_GFX_TEXT
        If Not (buflen = LenB(c15)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c15): Exit Sub
        CopyMemory c15, ByVal buffer, buflen
        set_constants_SHADER_GFX_TEXT c15
    Case ub_SHADER_SKYBOX
        Static c16 As constant_buffer_type_SHADER_SKYBOX
        If Not (buflen = LenB(c16)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c16): Exit Sub
        CopyMemory c16, ByVal buffer, buflen
        set_constants_SHADER_SKYBOX c16
    Case ub_SHADER_GENERIC_COLOR
        Static c17 As constant_buffer_type_SHADER_GENERIC_COLOR
        If Not (buflen = LenB(c17)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c17): Exit Sub
        CopyMemory c17, ByVal buffer, buflen
        set_constants_SHADER_GENERIC_COLOR c17
    Case ub_SHADER_RECT_INSTANCED
        Static c18 As constant_buffer_type_SHADER_RECT_INSTANCED
        If Not (buflen = LenB(c18)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c18): Exit Sub
        CopyMemory c18, ByVal buffer, buflen
        set_constants_SHADER_RECT_INSTANCED c18
    Case ub_SHADER_LINE
        Static c19 As constant_buffer_type_SHADER_LINE
        If Not (buflen = LenB(c19)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c19): Exit Sub
        CopyMemory c19, ByVal buffer, buflen
        set_constants_SHADER_LINE c19
    Case ub_SHADER_TEXT
        Static c20 As constant_buffer_type_SHADER_TEXT
        If Not (buflen = LenB(c20)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c20): Exit Sub
        CopyMemory c20, ByVal buffer, buflen
        set_constants_SHADER_TEXT c20
    Case ub_SHADER_DEBUG_TEXTURE2D_QUAD
        Static c21 As constant_buffer_type_SHADER_DEBUG_TEXTURE2D_QUAD
        If Not (buflen = LenB(c21)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c21): Exit Sub
        CopyMemory c21, ByVal buffer, buflen
        set_constants_SHADER_DEBUG_TEXTURE2D_QUAD c21
    Case ub_SHADER_DEBUG_TEXTURE2DARRAY_QUAD
        Static c22 As constant_buffer_type_SHADER_DEBUG_TEXTURE2DARRAY_QUAD
        If Not (buflen = LenB(c22)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c22): Exit Sub
        CopyMemory c22, ByVal buffer, buflen
        set_constants_SHADER_DEBUG_TEXTURE2DARRAY_QUAD c22
    Case ub_SHADER_WORLD_TO_SCREEN
        Static c23 As constant_buffer_type_SHADER_WORLD_TO_SCREEN
        If Not (buflen = LenB(c23)) Then ASSERT_FAILURE "buflen == sizeof(const gfx_api::constant_buffer_type<shader>)", "lib/ivis_opengl/gfx_api_gl.cpp:1138", "uniform_binding_entry", "Unexpected buffer size; received %zu, expecting %zu", buflen, LenB(c23): Exit Sub
        CopyMemory c23, ByVal buffer, buflen
        set_constants_SHADER_WORLD_TO_SCREEN c23
    End Select
End Sub

Public Sub set_constants(ByVal buffer As Long, ByVal size As Long)
    ' uniform_bind_functions[0](buffer, size)
    If uniform_bind_functions_size = 0 Then Exit Sub
    uniform_bind uniform_bind_functions(0), buffer, size
End Sub

' uniform_blocks: blocks(i) (an address, 0 = nullptr) of sizes(i) bytes, i < count
Public Sub set_uniforms(ByVal first As Long, blocks() As Long, sizes() As Long, ByVal count As Long)
    Dim i As Long, e As Long, uniform_bind_function As Long, buffer As Long
    e = count
    i = 0
    Do While i < e And (first + i) < uniform_bind_functions_size
        uniform_bind_function = uniform_bind_functions(first + i)
        buffer = blocks(i)
        If buffer = 0 Then
            GoTo continue_
        End If
        uniform_bind uniform_bind_function, buffer, sizes(i)
continue_:
        i = i + 1
    Loop
End Sub

Public Sub bind()
    glUseProgram program
    Select Case desc.blend_state
    Case REND_OPAQUE
        glDisable GL_BLEND

    Case REND_ALPHA
        glEnable GL_BLEND
        glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA

    Case REND_ADDITIVE
        glEnable GL_BLEND
        glBlendFunc GL_SRC_ALPHA, GL_ONE

    Case REND_MULTIPLICATIVE
        glEnable GL_BLEND
        glBlendFunc GL_DST_COLOR, GL_ONE_MINUS_SRC_ALPHA

    Case REND_PREMULTIPLIED
        glEnable GL_BLEND
        glBlendFunc GL_ONE, GL_ONE_MINUS_SRC_ALPHA

    Case REND_TEXT
        glEnable GL_BLEND
        glBlendFunc GL_ONE, GL_ONE_MINUS_SRC_ALPHA ' Should be GL_ONE_MINUS_SRC1_COLOR, if supported. Also, gl_FragData[1] then needs to be set in text.frag.
    End Select

    Select Case desc.depth_mode
    Case DEPTH_CMP_LEQ_WRT_OFF
        glEnable GL_DEPTH_TEST
        glDepthFunc GL_LEQUAL
        glDepthMask GL_FALSE
    Case DEPTH_CMP_LEQ_WRT_ON
        glEnable GL_DEPTH_TEST
        glDepthFunc GL_LEQUAL
        glDepthMask GL_TRUE
    Case DEPTH_CMP_ALWAYS_WRT_ON
        glDisable GL_DEPTH_TEST
        glDepthMask GL_TRUE

    Case DEPTH_CMP_ALWAYS_WRT_OFF
        glDisable GL_DEPTH_TEST
        glDepthMask GL_FALSE
    End Select

    If desc.output_mask = 0 Then
        glColorMask GL_FALSE, GL_FALSE, GL_FALSE, GL_FALSE
    Else
        glColorMask GL_TRUE, GL_TRUE, GL_TRUE, GL_TRUE
    End If

    If desc.offset Then
        glEnable GL_POLYGON_OFFSET_FILL
    Else
        glDisable GL_POLYGON_OFFSET_FILL
    End If

    Select Case desc.stencil
    Case stencil_mode_stencil_shadow_quad
        glEnable GL_STENCIL_TEST
        glStencilOp GL_KEEP, GL_KEEP, GL_KEEP
        glStencilMask Not 0
        glStencilFunc GL_LESS, 0, Not 0
    Case stencil_mode_stencil_shadow_silhouette
        glEnable GL_STENCIL_TEST
        glStencilMask Not 0
        glStencilOpSeparate GL_BACK, GL_KEEP, GL_KEEP, GL_INCR_WRAP
        glStencilOpSeparate GL_FRONT, GL_KEEP, GL_KEEP, GL_DECR_WRAP
        glStencilFunc GL_ALWAYS, 0, Not 0

    Case stencil_mode_stencil_disabled
        glDisable GL_STENCIL_TEST
        'glDisable(GL_STENCIL_TEST_TWO_SIDE_EXT);
    End Select

    Select Case desc.cull
    Case cull_mode_back
        glEnable GL_CULL_FACE
        glCullFace GL_FRONT
    Case cull_mode_shadow_mapping, cull_mode_front
        glEnable GL_CULL_FACE
        glCullFace GL_BACK
    Case cull_mode_none
        glDisable GL_CULL_FACE
    End Select
End Sub

' Read shader into text buffer
Private Function readShaderBuf(ByVal name_ As String, ancestorIncludePaths As StringVector) As String
    Dim fp As Long, filesize As Long, buffer() As Byte, n As Long, strResult As String
    Dim ancestorIncludesPlusMe As StringVector

    fp = PHYSFS_openRead(name_)
    If fp = 0 Then
        debug_ 1381, LOG_FATAL, "readShaderBuf", "Failed to read required shader file: %s\nPlease remove any mods and/or reinstall Warzone 2100.", name_
        Exit Function
    End If
    debug_ 1384, LOG_3D, "readShaderBuf", "Reading...[directory: %s] %s", PHYSFS_getRealDir(name_), name_
    filesize = CLng(PHYSFS_fileLength(fp))
    If filesize > 0 Then
        ReDim buffer(0 To filesize - 1)
        WZ_PHYSFS_readBytes fp, buffer, 0, filesize
    End If
    PHYSFS_close fp

    If filesize > 0 Then
        ' strResult = buffer (a C string: up to the first NUL)
        n = 0
        Do While n < filesize
            If buffer(n) = 0 Then Exit Do
            n = n + 1
        Loop
        If n < filesize Then ReDim Preserve buffer(0 To n)
        strResult = StrConv(buffer, vbUnicode)
        If n < filesize Then strResult = Left$(strResult, n)
    End If

    ancestorIncludesPlusMe = ancestorIncludePaths
    StringVector_push_back ancestorIncludesPlusMe, name_
    patchShaderHandleIncludes strResult, ancestorIncludesPlusMe

    readShaderBuf = strResult
End Function

Private Sub patchShaderHandleIncludes(shaderStr As String, ancestorIncludePaths As StringVector)
    Dim re As Object, matches As Object, m As Object, result As String, lastEnd As Long
    Dim parentPath As String, relativeIncludePath As String, fullIncludePath As String, replacement As String
    If Not (ancestorIncludePaths.size > 0) Then ASSERT_FAILURE "!ancestorIncludePaths.empty()", "lib/ivis_opengl/gfx_api_gl.cpp:1440", "patchShaderHandleIncludes", "No ancestors?": Exit Sub

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\s*#include\s+(\"".*\"")\s*"
    re.Global = True

    ' compute parent path basedir
    parentPath = WzPathInfo_path(ancestorIncludePaths.items(ancestorIncludePaths.size - 1))

    ' regex_replace_func(shaderStr, re, [&parentPathInfo, &ancestorIncludePaths](const std::smatch& match) -> std::string {...})
    Set matches = re.Execute(shaderStr)
    If matches.Count = 0 Then Exit Sub
    lastEnd = 1
    For Each m In matches
        result = result & Mid$(shaderStr, lastEnd, m.FirstIndex + 1 - lastEnd)
        relativeIncludePath = m.SubMatches(0)

        ' remove any surrounding " "
        Do While Left$(relativeIncludePath, 1) = """"
            relativeIncludePath = Mid$(relativeIncludePath, 2)
        Loop
        Do While Right$(relativeIncludePath, 1) = """"
            relativeIncludePath = Left$(relativeIncludePath, Len(relativeIncludePath) - 1)
        Loop

        fullIncludePath = parentPath & "/" & relativeIncludePath

        If ancestorIncludePaths.size > 5 Then
            debug_ 1462, LOG_ERROR, "patchShaderHandleIncludes", "Nested includes depth > 5 - not loading: %s", fullIncludePath
            replacement = vbLf
        Else
            replacement = vbLf & readShaderBuf(fullIncludePath, ancestorIncludePaths) & vbLf
        End If
        result = result & replacement
        lastEnd = m.FirstIndex + m.Length + 1
    Next
    result = result & Mid$(shaderStr, lastEnd)
    shaderStr = result
End Sub

' lib/framework/wzpaths.cpp's WzPathInfo::path (fromPlatformIndependentPath(file))
Private Function WzPathInfo_path(ByVal file As String) As String
    Dim lastSlashPos As Long
    lastSlashPos = InStrRev(file, "/")
    If lastSlashPos = 0 Then
        ' no slash - return "." as the path
        WzPathInfo_path = "."
        Exit Function
    End If
    ' file.substr(0, std::max<size_t>(lastSlashPos, (size_t)1))
    If lastSlashPos - 1 < 1 Then WzPathInfo_path = Left$(file, 1) Else WzPathInfo_path = Left$(file, lastSlashPos - 1)
End Function

' Retrieve shader compilation errors
Private Sub printShaderInfoLog(ByVal part As code_part, ByVal shader As Long)
    Dim infologLen As Long, charsWritten As Long, infoLog() As Byte

    glGetShaderiv shader, GL_INFO_LOG_LENGTH, VarPtr(infologLen)
    If infologLen > 0 Then
        charsWritten = 0
        ReDim infoLog(0 To infologLen - 1)

        glGetShaderInfoLog shader, infologLen, VarPtr(charsWritten), VarPtr(infoLog(0))

        debug_ 1488, part, "printShaderInfoLog", "Shader info log: %s", gl_cstr(VarPtr(infoLog(0)))
    End If
End Sub

' Retrieve shader linkage errors
Private Sub printProgramInfoLog(ByVal part As code_part, ByVal program_ As Long)
    Dim infologLen As Long, charsWritten As Long, infoLog() As Byte

    glGetProgramiv program_, GL_INFO_LOG_LENGTH, VarPtr(infologLen)
    If infologLen > 0 Then
        charsWritten = 0
        ReDim infoLog(0 To infologLen - 1)

        glGetProgramInfoLog program_, infologLen, VarPtr(charsWritten), VarPtr(infoLog(0))
        debug_ 1505, part, "printProgramInfoLog", "Program info log: %s", gl_cstr(VarPtr(infoLog(0)))
    End If
End Sub

' samplersToBind: pd's additional_samplers
Private Sub getLocs(pd As program_data)
    Dim locTex(0 To 3) As Long, i As Long, loc As Long
    glUseProgram program

    ' Uniforms, these never change.
    locTex(0) = glGetUniformLocation(program, "Texture")
    locTex(1) = glGetUniformLocation(program, "TextureTcmask")
    locTex(2) = glGetUniformLocation(program, "TextureNormal")
    locTex(3) = glGetUniformLocation(program, "TextureSpecular")

    For i = 0 To 3
        If locTex(i) <> -1 Then
            glUniform1i locTex(i), i
        End If
    Next

    ' additional sampler uniforms
    For i = 0 To pd.additional_samplers_size - 1
        loc = glGetUniformLocation(program, pd.additional_sampler_names(i))
        If loc <> -1 Then
            glUniform1i loc, pd.additional_sampler_units(i)
        Else
            debug_ 1539, LOG_3D, "getLocs", "Missing expected sampler uniform: %s", pd.additional_sampler_names(i)
        End If
    Next

End Sub

' std::unordered_set<std::string>: a Scripting.Dictionary of the names
Private Function getUniformNamesFromSource(ByVal shaderContents As String) As Object
    Dim uniformNames As Object, re As Object, m As Object, uniformLineStr As String, uniformName As String
    Dim lastWhitespaceIdx As Long, p As Long, q As Long

    Set uniformNames = CreateObject("Scripting.Dictionary")
    uniformNames.CompareMode = vbBinaryCompare
    Set getUniformNamesFromSource = uniformNames
    If StrPtr(shaderContents) = 0 Then
        debug_ 1549, LOG_INFO, "getUniformNamesFromSource", "shaderContents is null"
        Exit Function
    End If

    ' White space: the space character, horizontal tab, vertical tab, form feed, carriage-return, and line-feed.

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "uniform.*(?![^\{]*\})(?=[=\[;])"
    re.Global = True

    For Each m In re.Execute(shaderContents)
        uniformLineStr = m.Value

        ' trim glsl whitespace chars from beginning / end
        p = 1
        Do While p <= Len(uniformLineStr)
            If InStr(glsl_whitespace_chars, Mid$(uniformLineStr, p, 1)) = 0 Then Exit Do
            p = p + 1
        Loop
        uniformLineStr = Mid$(uniformLineStr, p)
        q = Len(uniformLineStr)
        Do While q >= 1
            If InStr(glsl_whitespace_chars, Mid$(uniformLineStr, q, 1)) = 0 Then Exit Do
            q = q - 1
        Loop
        uniformLineStr = Left$(uniformLineStr, q)

        ' size_t lastWhitespaceIdx = uniformLineStr.find_last_of(glsl_whitespace_chars) (1-based here; 0 = npos)
        lastWhitespaceIdx = 0
        For q = Len(uniformLineStr) To 1 Step -1
            If InStr(glsl_whitespace_chars, Mid$(uniformLineStr, q, 1)) > 0 Then
                lastWhitespaceIdx = q
                Exit For
            End If
        Next
        If lastWhitespaceIdx <> 0 Then
            uniformName = Mid$(m.Value, lastWhitespaceIdx + 1)
            If Len(uniformName) > 0 Then
                uniformNames(uniformName) = True
            End If
        End If
    Next
End Function

' returns the modified fragment shader source; duplicateFragmentUniformNameMap: a Scripting.Dictionary
Private Function renameDuplicateFragmentShaderUniforms(ByVal vertexShaderContents As String, ByVal fragmentShaderContents As String, duplicateFragmentUniformNameMap As Object) As String
    Dim vertexUniformNames As Object, fragmentUniformNames As Object, fragmentUniform As Variant
    Dim modifiedFragmentShaderSource As String, replacementUniformName As String, re As Object

    Set duplicateFragmentUniformNameMap = CreateObject("Scripting.Dictionary")
    duplicateFragmentUniformNameMap.CompareMode = vbBinaryCompare

    Set vertexUniformNames = getUniformNamesFromSource(vertexShaderContents)
    Set fragmentUniformNames = getUniformNamesFromSource(fragmentShaderContents)

    modifiedFragmentShaderSource = fragmentShaderContents
    For Each fragmentUniform In fragmentUniformNames.Keys
        If vertexUniformNames.Exists(fragmentUniform) Then
            ' duplicate uniform name found - rename it!
            replacementUniformName = "wzfix_frag_" & fragmentUniform
            duplicateFragmentUniformNameMap(fragmentUniform) = replacementUniformName

            ' and replace it in the fragment shader source
            Set re = CreateObject("VBScript.RegExp")
            re.Pattern = "\b" & fragmentUniform & "\b"
            re.Global = True
            modifiedFragmentShaderSource = re.Replace(modifiedFragmentShaderSource, replacementUniformName)
        End If
    Next

    renameDuplicateFragmentShaderUniforms = modifiedFragmentShaderSource
End Function

Private Function regex_replace_wrapper(input_ As String, ByVal pattern As String, ByVal replace_ As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.Global = True
    If Not re.Test(input_) Then
        regex_replace_wrapper = False
        Exit Function
    End If

    input_ = re.Replace(input_, replace_)

    regex_replace_wrapper = True
End Function

Private Sub patchFragmentShaderTextureLodBias(fragmentShaderStr As String, ByVal mipLodBias As Single)
    Dim re As Object, floatAsString As String, lastNon0Pos As Long, i As Long
    ' Look for:
    ' #define WZ_MIP_LOAD_BIAS 0.f
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "#define WZ_MIP_LOAD_BIAS .*"
    re.Global = True

    floatAsString = va("%f", mipLodBias)
    lastNon0Pos = 0
    For i = Len(floatAsString) To 1 Step -1
        If Mid$(floatAsString, i, 1) <> "0" Then
            lastNon0Pos = i
            Exit For
        End If
    Next
    If lastNon0Pos <> 0 Then
        floatAsString = Left$(floatAsString, lastNon0Pos)
    Else
        ' only 0?
        Exit Sub
    End If
    For i = 1 To Len(floatAsString)
        If InStr("0123456789.-", Mid$(floatAsString, i, 1)) = 0 Then
            ASSERT_FAILURE "floatAsString.find_last_not_of(""0123456789.-"") == std::string::npos", "lib/ivis_opengl/gfx_api_gl.cpp:1661", "patchFragmentShaderTextureLodBias", "Found unexpected / invalid character in: %s", floatAsString
            Exit For
        End If
    Next

    fragmentShaderStr = re.Replace(fragmentShaderStr, va("#define WZ_MIP_LOAD_BIAS %s", floatAsString))
End Sub

Private Function patchFragmentShaderPointLightsDefines(fragmentShaderStr As String, lightingConstants As lighting_constants) As Boolean
    Dim defines(0 To 3) As String, values(0 To 3) As Long, i As Long, foundAndReplaced_PointLightsDefine As Boolean, r As Boolean
    defines(0) = "WZ_MAX_POINT_LIGHTS": values(0) = gfx_api_max_lights
    defines(1) = "WZ_MAX_INDEXED_POINT_LIGHTS": values(1) = gfx_api_max_indexed_lights
    defines(2) = "WZ_BUCKET_DIMENSION": values(2) = gfx_api_bucket_dimension
    defines(3) = "WZ_POINT_LIGHT_ENABLED": values(3) = IIf(lightingConstants.isPointLightPerPixelEnabled, 1, 0)

    ' replacer(define, value): regex_replace_wrapper(fragmentShaderStr, "#define {} .*", "#define {} {}")
    foundAndReplaced_PointLightsDefine = False
    For i = 0 To 3
        r = regex_replace_wrapper(fragmentShaderStr, "#define " & defines(i) & " .*", "#define " & defines(i) & " " & CStr(values(i)))
        foundAndReplaced_PointLightsDefine = r Or foundAndReplaced_PointLightsDefine
    Next
    patchFragmentShaderPointLightsDefines = foundAndReplaced_PointLightsDefine
End Function

Private Function patchFragmentShaderShadowConstants(fragmentShaderStr As String, shadowConstants As lighting_constants) As Boolean
    Dim foundAndReplaced_shadowMode As Boolean, foundAndReplaced_shadowFilterSize As Boolean, foundAndReplaced_shadowCascadesCount As Boolean
    ' #define WZ_SHADOW_MODE <number>
    foundAndReplaced_shadowMode = regex_replace_wrapper(fragmentShaderStr, "#define WZ_SHADOW_MODE .*", va("#define WZ_SHADOW_MODE %u", shadowConstants.shadowMode))

    ' #define WZ_SHADOW_FILTER_SIZE <number>
    foundAndReplaced_shadowFilterSize = regex_replace_wrapper(fragmentShaderStr, "#define WZ_SHADOW_FILTER_SIZE .*", va("#define WZ_SHADOW_FILTER_SIZE %u", shadowConstants.shadowFilterSize))

    ' #define WZ_SHADOW_CASCADES_COUNT <number>
    foundAndReplaced_shadowCascadesCount = regex_replace_wrapper(fragmentShaderStr, "#define WZ_SHADOW_CASCADES_COUNT .*", va("#define WZ_SHADOW_CASCADES_COUNT %u", shadowConstants.shadowCascadesCount))

    patchFragmentShaderShadowConstants = foundAndReplaced_shadowMode Or foundAndReplaced_shadowFilterSize Or foundAndReplaced_shadowCascadesCount
End Function

' the attribute location of index is bound only if a vertex buffer description uses it
Private Sub bindVertexAttribLocationIfUsed(expectedVertexAttribLoc() As Boolean, ByVal prg As Long, ByVal index As Long, ByVal name_ As String)
    If expectedVertexAttribLoc(index) Then
        glBindAttribLocation prg, index, name_
    End If
End Sub

' uniformNames: pd.uniform_names; samplersToBind: pd.additional_samplers
Private Sub build_program(ByVal fragmentHighpFloatAvailable As Boolean, ByVal fragmentHighpIntAvailable As Boolean, _
                          ByVal patchFragmentShaderMipLodBias As Boolean, _
                          ByVal programName As String, _
                          ByVal vertex_header As String, ByVal vertexPath As String, _
                          ByVal fragment_header As String, ByVal fragmentPath As String, _
                          uniformNames As StringVector, _
                          pd As program_data, _
                          ByVal has_mipLodBias As Boolean, ByVal mipLodBias As Single, lightingConstants As lighting_constants)
    Dim status As Long, success As Boolean, i As Long, j As Long, shader As Long
    Dim expectedVertexAttribLoc(0 To 63) As Boolean
    Dim vertexShaderContents As String, fragmentShaderStr As String, noAncestors As StringVector
    Dim duplicateFragmentUniformNames As StringVector, fragmentUniformNameChanges As Object, it As Variant
    Dim originalUniformName As String, replacementUniformName As String, uniformIdx As Long
    success = True ' Assume overall success

    For i = 0 To vertex_buffer_desc_size - 1
        For j = 0 To vertex_buffer_desc(i).attributes_size - 1
            expectedVertexAttribLoc(vertex_buffer_desc(i).attributes(j).id) = True
        Next
    Next

    program = glCreateProgram()
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, 0, "vertex"
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, 1, "vertexTexCoord"
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, 2, "vertexColor"
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, 3, "vertexNormal"
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, 4, "vertexTangent"
    ' only needed for instanced mesh rendering
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, gfx_api_instance_modelMatrix, "instanceModelMatrix" ' uses 4 slots
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, gfx_api_instance_packedValues, "instancePackedValues"
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, gfx_api_instance_Colour, "instanceColour"
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, gfx_api_instance_TeamColour, "instanceTeamColour"
    If Not (program <> 0) Then ASSERT_FAILURE "program", "lib/ivis_opengl/gfx_api_gl.cpp:1744", "build_program", "Could not create shader program!": Exit Sub
    ' only needed for new terrain renderer
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, gfx_api_terrain_tileNo, "tileNo"
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, gfx_api_terrain_grounds, "grounds"
    bindVertexAttribLocationIfUsed expectedVertexAttribLoc, program, gfx_api_terrain_groundWeights, "groundWeights"

    If Len(vertexPath) > 0 Then
        success = False ' Assume failure before reading shader file

        vertexShaderContents = readShaderBuf(vertexPath, noAncestors)
        If Len(vertexShaderContents) > 0 Then
            shader = glCreateShader(GL_VERTEX_SHADER)
            vertexShader = shader

            shaderSource2 shader, vertex_header, vertexShaderContents
            glCompileShader shader

            ' Check for compilation errors
            glGetShaderiv shader, GL_COMPILE_STATUS, VarPtr(status)
            If status = 0 Then
                debug_ 1771, LOG_ERROR, "build_program", "Vertex shader compilation has failed [%s]", vertexPath
                printShaderInfoLog LOG_ERROR, shader
            Else
                printShaderInfoLog LOG_3D, shader
                glAttachShader program, shader
                success = True
            End If
            gl_context_wzGLObjectLabel GL_SHADER, shader, -1, vertexPath
        End If
    End If

    If success And Len(fragmentPath) > 0 Then
        success = False ' Assume failure before reading shader file

        fragmentShaderStr = readShaderBuf(fragmentPath, noAncestors)
        If Len(fragmentShaderStr) > 0 Then
            shader = glCreateShader(GL_FRAGMENT_SHADER)
            fragmentShader = shader

            If Not fragmentHighpFloatAvailable Or Not fragmentHighpIntAvailable Then
                ' rename duplicate uniforms
                ' to avoid conflicting uniform precision issues
                fragmentShaderStr = renameDuplicateFragmentShaderUniforms(vertexShaderContents, fragmentShaderStr, fragmentUniformNameChanges)

                ' duplicateFragmentUniformNames.resize(uniformNames.size())
                duplicateFragmentUniformNames.size = uniformNames.size
                If uniformNames.size > 0 Then ReDim duplicateFragmentUniformNames.items(0 To uniformNames.size - 1)
                For Each it In fragmentUniformNameChanges.Keys
                    originalUniformName = it
                    replacementUniformName = fragmentUniformNameChanges(it)

                    debug_ 1811, LOG_3D, "build_program", " - Found duplicate uniform name in fragment shader: %s", originalUniformName

                    ' find uniform index in global uniforms array
                    uniformIdx = -1
                    For i = 0 To uniformNames.size - 1
                        If uniformNames.items(i) = originalUniformName Then
                            uniformIdx = i
                            Exit For
                        End If
                    Next
                    If uniformIdx <> -1 Then
                        duplicateFragmentUniformNames.items(uniformIdx) = replacementUniformName

                        debug_ 1820, LOG_3D, "build_program", "  - Renaming ""%s"" -> ""%s"" in fragment shader", originalUniformName, replacementUniformName
                    End If
                Next
            End If

            If patchFragmentShaderMipLodBias And has_mipLodBias Then
                patchFragmentShaderTextureLodBias fragmentShaderStr, mipLodBias
            End If
            hasSpecializationConstant_ShadowConstants = patchFragmentShaderShadowConstants(fragmentShaderStr, lightingConstants)
            hasSpecializationConstants_PointLights = patchFragmentShaderPointLightsDefines(fragmentShaderStr, lightingConstants)

            shaderSource2 shader, fragment_header, fragmentShaderStr
            glCompileShader shader

            ' Check for compilation errors
            glGetShaderiv shader, GL_COMPILE_STATUS, VarPtr(status)
            If status = 0 Then
                debug_ 1841, LOG_ERROR, "build_program", "Fragment shader compilation has failed [%s]", fragmentPath
                printShaderInfoLog LOG_ERROR, shader
            Else
                printShaderInfoLog LOG_3D, shader
                glAttachShader program, shader
                success = True
            End If
            gl_context_wzGLObjectLabel GL_SHADER, shader, -1, fragmentPath
        End If
    End If

    If success Then
        glLinkProgram program

        ' Check for linkage errors
        glGetProgramiv program, GL_LINK_STATUS, VarPtr(status)
        If status = 0 Then
            debug_ 1864, LOG_ERROR, "build_program", "Shader program linkage has failed [%s, %s]", vertexPath, fragmentPath
            printProgramInfoLog LOG_ERROR, program
            success = False
        Else
            printProgramInfoLog LOG_3D, program
        End If
        gl_context_wzGLObjectLabel GL_PROGRAM, program, -1, programName
    End If
    fetch_uniforms uniformNames, duplicateFragmentUniformNames, programName
    getLocs pd
    broken = broken Or Not success
End Sub

' glShaderSource(shader, 2, { header, source }, nullptr)
Private Sub shaderSource2(ByVal shader As Long, ByVal header As String, ByVal source As String)
    Dim b0() As Byte, b1() As Byte, ShaderStrings(0 To 1) As Long
    b0 = StrConv(header & vbNullChar, vbFromUnicode)
    b1 = StrConv(source & vbNullChar, vbFromUnicode)
    ShaderStrings(0) = VarPtr(b0(0))
    ShaderStrings(1) = VarPtr(b1(0))
    glShaderSource shader, 2, VarPtr(ShaderStrings(0)), 0
End Sub

Private Sub fetch_uniforms(uniformNames As StringVector, duplicateFragmentUniformNames As StringVector, ByVal programName As String)
    Dim i As Long, result As Long
    locations_size = uniformNames.size
    If locations_size > 0 Then ReDim locations(0 To locations_size - 1)
    For i = 0 To uniformNames.size - 1
        result = glGetUniformLocation(program, uniformNames.items(i))
        If result = -1 Then
            debug_ 1890, LOG_3D, "fetch_uniforms", "[%s]: Did not find uniform: %s", programName, uniformNames.items(i)
        End If
        locations(i) = result
    Next
    If duplicateFragmentUniformNames.size > 0 Then
        duplicateFragmentUniformLocations_size = duplicateFragmentUniformNames.size
        ReDim duplicateFragmentUniformLocations(0 To duplicateFragmentUniformLocations_size - 1)
        For i = 0 To duplicateFragmentUniformNames.size - 1
            If Len(duplicateFragmentUniformNames.items(i)) = 0 Then
                duplicateFragmentUniformLocations(i) = -1
            Else
                duplicateFragmentUniformLocations(i) = glGetUniformLocation(program, duplicateFragmentUniformNames.items(i))
            End If
        Next
    Else
        duplicateFragmentUniformLocations_size = uniformNames.size
        If duplicateFragmentUniformLocations_size > 0 Then ReDim duplicateFragmentUniformLocations(0 To duplicateFragmentUniformLocations_size - 1)
        For i = 0 To duplicateFragmentUniformLocations_size - 1
            duplicateFragmentUniformLocations(i) = -1
        Next
    End If
End Sub

Private Sub setUniforms_vec4(ByVal uniformIdx As Long, v As Vector4f)
    glUniform4f locations(uniformIdx), v.x, v.y, v.z, v.w
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniform4f duplicateFragmentUniformLocations(uniformIdx), v.x, v.y, v.z, v.w
    End If
End Sub

Private Sub setUniforms_mat4(ByVal uniformIdx As Long, m As mat4)
    glUniformMatrix4fv locations(uniformIdx), 1, GL_FALSE, VarPtr(m)
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniformMatrix4fv duplicateFragmentUniformLocations(uniformIdx), 1, GL_FALSE, VarPtr(m)
    End If
End Sub

' m: the address of the first glm::mat4
Private Sub setUniforms_mat4_count(ByVal uniformIdx As Long, ByVal m As Long, ByVal count As Long)
    glUniformMatrix4fv locations(uniformIdx), count, GL_FALSE, m
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniformMatrix4fv duplicateFragmentUniformLocations(uniformIdx), count, GL_FALSE, m
    End If
End Sub

' m: the address of the first glm::vec4
Private Sub setUniforms_vec4_count(ByVal uniformIdx As Long, ByVal m As Long, ByVal count As Long)
    glUniform4fv locations(uniformIdx), count, m
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniform4fv duplicateFragmentUniformLocations(uniformIdx), count, m
    End If
End Sub

' m: the address of the first glm::ivec4
Private Sub setUniforms_ivec4_count(ByVal uniformIdx As Long, ByVal m As Long, ByVal count As Long)
    glUniform4iv locations(uniformIdx), count, m
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniform4iv duplicateFragmentUniformLocations(uniformIdx), count, m
    End If
End Sub

' v: the address of the first float
Private Sub setUniforms_float_count(ByVal uniformIdx As Long, ByVal v As Long, ByVal count As Long)
    glUniform1fv locations(uniformIdx), count, v
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniform1fv duplicateFragmentUniformLocations(uniformIdx), count, v
    End If
End Sub

Private Sub setUniforms_ivec4(ByVal uniformIdx As Long, v As Vector4i)
    glUniform4i locations(uniformIdx), v.x, v.y, v.z, v.w
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniform4i duplicateFragmentUniformLocations(uniformIdx), v.x, v.y, v.z, v.w
    End If
End Sub

Private Sub setUniforms_ivec2(ByVal uniformIdx As Long, v As Vector2i)
    glUniform2i locations(uniformIdx), v.x, v.y
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniform2i duplicateFragmentUniformLocations(uniformIdx), v.x, v.y
    End If
End Sub

Private Sub setUniforms_vec2(ByVal uniformIdx As Long, v As Vector2f)
    glUniform2f locations(uniformIdx), v.x, v.y
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniform2f duplicateFragmentUniformLocations(uniformIdx), v.x, v.y
    End If
End Sub

Private Sub setUniforms_int(ByVal uniformIdx As Long, ByVal v As Long)
    glUniform1i locations(uniformIdx), v
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniform1i duplicateFragmentUniformLocations(uniformIdx), v
    End If
End Sub

Private Sub setUniforms_float(ByVal uniformIdx As Long, ByVal v As Single)
    glUniform1f locations(uniformIdx), v
    If duplicateFragmentUniformLocations(uniformIdx) <> -1 Then
        glUniform1f duplicateFragmentUniformLocations(uniformIdx), v
    End If
End Sub

' MARK: -

Private Sub Class_Terminate()
    If Me.vertexShader <> 0 Then glDeleteShader Me.vertexShader
    If Me.fragmentShader <> 0 Then glDeleteShader Me.fragmentShader
    glDeleteProgram Me.program
End Sub

Private Sub set_constants_Draw3DShapeGlobalUniforms(cbuf As Draw3DShapeGlobalUniforms)
    setUniforms_mat4 0, cbuf.ProjectionMatrix
    setUniforms_mat4 1, cbuf.ViewMatrix
    setUniforms_mat4 2, cbuf.ShadowMapMVPMatrix
    setUniforms_vec4 3, cbuf.sunPos
    setUniforms_vec4 4, cbuf.sceneColor
    setUniforms_vec4 5, cbuf.ambient
    setUniforms_vec4 6, cbuf.diffuse
    setUniforms_vec4 7, cbuf.specular
    setUniforms_vec4 8, cbuf.fogColour
    setUniforms_float 9, cbuf.fogEnd
    setUniforms_float 10, cbuf.fogBegin
    setUniforms_float 11, cbuf.timeState
    setUniforms_int 12, cbuf.fogEnabled
End Sub

Private Sub set_constants_Draw3DShapePerMeshUniforms(cbuf As Draw3DShapePerMeshUniforms)
    setUniforms_int 13, cbuf.tcmask
    setUniforms_int 14, cbuf.normalMap
    setUniforms_int 15, cbuf.specularMap
    setUniforms_int 16, cbuf.hasTangents
End Sub

Private Sub set_constants_Draw3DShapePerInstanceUniforms(cbuf As Draw3DShapePerInstanceUniforms)
    setUniforms_mat4 17, cbuf.ModelViewMatrix
    setUniforms_mat4 18, cbuf.NormalMatrix
    setUniforms_vec4 19, cbuf.colour
    setUniforms_vec4 20, cbuf.teamcolour
    setUniforms_float 21, cbuf.shaderStretch
    setUniforms_float 22, cbuf.animFrameNumber
    setUniforms_int 23, cbuf.ecmState
    setUniforms_int 24, cbuf.alphaTest
End Sub

Private Sub set_constants_Draw3DShapeInstancedGlobalUniforms(cbuf As Draw3DShapeInstancedGlobalUniforms)
    setUniforms_mat4 0, cbuf.ProjectionMatrix
    setUniforms_mat4 1, cbuf.ViewMatrix
    setUniforms_mat4 2, cbuf.ModelUVLightmapMatrix
    setUniforms_mat4_count 3, VarPtr(cbuf.ShadowMapMVPMatrix(0)), WZ_MAX_SHADOW_CASCADES
    setUniforms_vec4 4, cbuf.cameraPos
    setUniforms_vec4 5, cbuf.sunPos
    setUniforms_vec4 6, cbuf.sceneColor
    setUniforms_vec4 7, cbuf.ambient
    setUniforms_vec4 8, cbuf.diffuse
    setUniforms_vec4 9, cbuf.specular
    setUniforms_vec4 10, cbuf.fogColour
    setUniforms_vec4 11, cbuf.ShadowMapCascadeSplits
    setUniforms_int 12, cbuf.ShadowMapSize
    setUniforms_float 13, cbuf.fogEnd
    setUniforms_float 14, cbuf.fogBegin
    setUniforms_float 15, cbuf.timeState
    setUniforms_int 16, cbuf.fogEnabled
    setUniforms_vec4_count 17, VarPtr(cbuf.PointLightsPosition(0)), gfx_api_max_lights
    setUniforms_vec4_count 18, VarPtr(cbuf.PointLightsColorAndEnergy(0)), gfx_api_max_lights
    setUniforms_ivec4_count 19, VarPtr(cbuf.bucketOffsetAndSize(0)), gfx_api_bucket_dimension * gfx_api_bucket_dimension
    setUniforms_ivec4_count 20, VarPtr(cbuf.indexed_lights(0)), gfx_api_max_indexed_lights
    setUniforms_int 21, cbuf.bucketDimensionUsed
    setUniforms_int 22, cbuf.viewportWidth
    setUniforms_int 23, cbuf.viewportheight
End Sub

Private Sub set_constants_Draw3DShapeInstancedPerMeshUniforms(cbuf As Draw3DShapeInstancedPerMeshUniforms)
    ' IMPORTANT: uniformIdx continues incrementing from Draw3DShapeInstancedGlobalUniforms above
    setUniforms_int 24, cbuf.tcmask
    setUniforms_int 25, cbuf.normalMap
    setUniforms_int 26, cbuf.specularMap
    setUniforms_int 27, cbuf.hasTangents
    setUniforms_int 28, cbuf.shieldEffect
End Sub

Private Sub set_constants_Draw3DShapeInstancedDepthOnlyGlobalUniforms(cbuf As Draw3DShapeInstancedDepthOnlyGlobalUniforms)
    setUniforms_mat4 0, cbuf.ProjectionMatrix
    setUniforms_mat4 1, cbuf.ViewMatrix
End Sub

Private Sub set_constants_SHADER_TERRAIN_DEPTH(cbuf As constant_buffer_type_SHADER_TERRAIN_DEPTH)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_vec4 1, cbuf.paramX
    setUniforms_vec4 2, cbuf.paramY
    setUniforms_int 3, cbuf.texture0
    setUniforms_vec4 4, cbuf.paramXLight
    setUniforms_vec4 5, cbuf.paramYLight
    setUniforms_int 6, cbuf.fog_enabled
    setUniforms_float 7, cbuf.fog_begin
    setUniforms_float 8, cbuf.fog_end
End Sub

Private Sub set_constants_SHADER_TERRAIN_DEPTHMAP(cbuf As constant_buffer_type_SHADER_TERRAIN_DEPTHMAP)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_int 1, cbuf.fog_enabled
    setUniforms_float 2, cbuf.fog_begin
    setUniforms_float 3, cbuf.fog_end
End Sub

Private Sub set_constants_TerrainCombinedUniforms(cbuf As TerrainCombinedUniforms)
    Dim i As Long
    i = 0
    setUniforms_mat4 i, cbuf.ModelViewProjectionMatrix: i = i + 1
    setUniforms_mat4 i, cbuf.ViewMatrix: i = i + 1
    setUniforms_mat4 i, cbuf.ModelUVLightmapMatrix: i = i + 1
    setUniforms_mat4_count i, VarPtr(cbuf.ShadowMapMVPMatrix(0)), WZ_MAX_SHADOW_CASCADES: i = i + 1
    setUniforms_mat4 i, cbuf.groundScale: i = i + 1
    setUniforms_vec4 i, cbuf.cameraPos: i = i + 1
    setUniforms_vec4 i, cbuf.sunPos: i = i + 1
    setUniforms_vec4 i, cbuf.emissiveLight: i = i + 1
    setUniforms_vec4 i, cbuf.ambientLight: i = i + 1
    setUniforms_vec4 i, cbuf.diffuseLight: i = i + 1
    setUniforms_vec4 i, cbuf.specularLight: i = i + 1
    setUniforms_vec4 i, cbuf.fog_colour: i = i + 1
    setUniforms_vec4 i, cbuf.ShadowMapCascadeSplits: i = i + 1
    setUniforms_int i, cbuf.ShadowMapSize: i = i + 1
    setUniforms_int i, cbuf.fog_enabled: i = i + 1
    setUniforms_float i, cbuf.fog_begin: i = i + 1
    setUniforms_float i, cbuf.fog_end: i = i + 1
    setUniforms_int i, cbuf.quality: i = i + 1
    setUniforms_vec4_count i, VarPtr(cbuf.PointLightsPosition(0)), gfx_api_max_lights: i = i + 1
    setUniforms_vec4_count i, VarPtr(cbuf.PointLightsColorAndEnergy(0)), gfx_api_max_lights: i = i + 1
    setUniforms_ivec4_count i, VarPtr(cbuf.bucketOffsetAndSize(0)), gfx_api_bucket_dimension * gfx_api_bucket_dimension: i = i + 1
    setUniforms_ivec4_count i, VarPtr(cbuf.indexed_lights(0)), gfx_api_max_indexed_lights: i = i + 1
    setUniforms_int i, cbuf.bucketDimensionUsed: i = i + 1
    setUniforms_int i, cbuf.viewportWidth: i = i + 1
    setUniforms_int i, cbuf.viewportheight: i = i + 1
    setUniforms_int i, 0: i = i + 1 ' lightmap_tex
    setUniforms_int i, 1: i = i + 1 ' ground
    setUniforms_int i, 2: i = i + 1 ' groundNormal
    setUniforms_int i, 3: i = i + 1 ' groundSpecular
    setUniforms_int i, 4: i = i + 1 ' groundHeight
    setUniforms_int i, 5: i = i + 1 ' decal
    setUniforms_int i, 6: i = i + 1 ' decalNormal
    setUniforms_int i, 7: i = i + 1 ' decalSpecular
    setUniforms_int i, 8: i = i + 1 ' decalHeight
    setUniforms_int i, 9: i = i + 1 ' shadowMap
End Sub

Private Sub set_constants_SHADER_WATER(cbuf As constant_buffer_type_SHADER_WATER)
    Dim i As Long
    i = 0
    setUniforms_mat4 i, cbuf.ModelViewProjectionMatrix: i = i + 1
    setUniforms_mat4 i, cbuf.ModelUVLightmapMatrix: i = i + 1
    setUniforms_mat4 i, cbuf.ModelUV1Matrix: i = i + 1
    setUniforms_mat4 i, cbuf.ModelUV2Matrix: i = i + 1
    setUniforms_vec4 i, cbuf.cameraPos: i = i + 1
    setUniforms_vec4 i, cbuf.sunPos: i = i + 1
    setUniforms_vec4 i, cbuf.emissiveLight: i = i + 1
    setUniforms_vec4 i, cbuf.ambientLight: i = i + 1
    setUniforms_vec4 i, cbuf.diffuseLight: i = i + 1
    setUniforms_vec4 i, cbuf.specularLight: i = i + 1
    setUniforms_vec4 i, cbuf.fog_colour: i = i + 1
    setUniforms_int i, cbuf.fog_enabled: i = i + 1
    setUniforms_float i, cbuf.fog_begin: i = i + 1
    setUniforms_float i, cbuf.fog_end: i = i + 1
    setUniforms_float i, cbuf.timeSec: i = i + 1
     ' textures:
    setUniforms_int i, 0: i = i + 1
    setUniforms_int i, 1: i = i + 1
    setUniforms_int i, 2: i = i + 1 ' lightmap_tex
End Sub

Private Sub set_constants_SHADER_WATER_HIGH(cbuf As constant_buffer_type_SHADER_WATER_HIGH)
    Dim i As Long
    i = 0
    setUniforms_mat4 i, cbuf.ModelViewProjectionMatrix: i = i + 1
    setUniforms_mat4 i, cbuf.ViewMatrix: i = i + 1
    setUniforms_mat4 i, cbuf.ModelUVLightmapMatrix: i = i + 1
    setUniforms_mat4_count i, VarPtr(cbuf.ShadowMapMVPMatrix(0)), WZ_MAX_SHADOW_CASCADES: i = i + 1
    setUniforms_vec4 i, cbuf.cameraPos: i = i + 1
    setUniforms_vec4 i, cbuf.sunPos: i = i + 1
    setUniforms_vec4 i, cbuf.emissiveLight: i = i + 1
    setUniforms_vec4 i, cbuf.ambientLight: i = i + 1
    setUniforms_vec4 i, cbuf.diffuseLight: i = i + 1
    setUniforms_vec4 i, cbuf.specularLight: i = i + 1
    setUniforms_vec4 i, cbuf.fog_colour: i = i + 1
    setUniforms_vec4 i, cbuf.ShadowMapCascadeSplits: i = i + 1
    setUniforms_int i, cbuf.ShadowMapSize: i = i + 1
    setUniforms_int i, cbuf.fog_enabled: i = i + 1
    setUniforms_float i, cbuf.fog_begin: i = i + 1
    setUniforms_float i, cbuf.fog_end: i = i + 1
    setUniforms_float i, cbuf.timeSec: i = i + 1
     ' textures:
    setUniforms_int i, 0: i = i + 1
    setUniforms_int i, 1: i = i + 1
    setUniforms_int i, 2: i = i + 1
    setUniforms_int i, 3: i = i + 1 ' lightmap_tex
End Sub

Private Sub set_constants_SHADER_WATER_CLASSIC(cbuf As constant_buffer_type_SHADER_WATER_CLASSIC)
    Dim i As Long
    i = 0
    setUniforms_mat4 i, cbuf.ModelViewProjectionMatrix: i = i + 1
    setUniforms_mat4 i, cbuf.ModelUVLightmapMatrix: i = i + 1
    setUniforms_mat4 i, cbuf.ShadowMapMVPMatrix: i = i + 1
    setUniforms_mat4 i, cbuf.ModelUV1Matrix: i = i + 1
    setUniforms_mat4 i, cbuf.ModelUV2Matrix: i = i + 1
    setUniforms_vec4 i, cbuf.cameraPos: i = i + 1
    setUniforms_vec4 i, cbuf.sunPos: i = i + 1
    setUniforms_vec4 i, cbuf.fog_colour: i = i + 1
    setUniforms_int i, cbuf.fog_enabled: i = i + 1
    setUniforms_float i, cbuf.fog_begin: i = i + 1
    setUniforms_float i, cbuf.fog_end: i = i + 1
    setUniforms_float i, cbuf.timeSec: i = i + 1
     ' textures:
    setUniforms_int i, 0: i = i + 1
    setUniforms_int i, 1: i = i + 1
End Sub

Private Sub set_constants_SHADER_RECT(cbuf As constant_buffer_type_SHADER_RECT)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_vec4 1, cbuf.colour
End Sub

Private Sub set_constants_SHADER_TEXRECT(cbuf As constant_buffer_type_SHADER_TEXRECT)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_vec2 1, cbuf.offset
    setUniforms_vec2 2, cbuf.size
    setUniforms_vec4 3, cbuf.color
End Sub

Private Sub set_constants_SHADER_GFX_COLOUR(cbuf As constant_buffer_type_SHADER_GFX_COLOUR)
    setUniforms_mat4 0, cbuf.transform_matrix
End Sub

Private Sub set_constants_SHADER_GFX_TEXT(cbuf As constant_buffer_type_SHADER_GFX_TEXT)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_vec4 1, cbuf.color
End Sub

Private Sub set_constants_SHADER_SKYBOX(cbuf As constant_buffer_type_SHADER_SKYBOX)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_vec4 1, cbuf.color
    setUniforms_vec4 2, cbuf.fog_color
    setUniforms_int 3, cbuf.fog_enabled
End Sub

Private Sub set_constants_SHADER_GENERIC_COLOR(cbuf As constant_buffer_type_SHADER_GENERIC_COLOR)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_vec4 1, cbuf.colour
End Sub

Private Sub set_constants_SHADER_RECT_INSTANCED(cbuf As constant_buffer_type_SHADER_RECT_INSTANCED)
    setUniforms_mat4 0, cbuf.ProjectionMatrix
End Sub

Private Sub set_constants_SHADER_LINE(cbuf As constant_buffer_type_SHADER_LINE)
    setUniforms_vec2 0, cbuf.p0
    setUniforms_vec2 1, cbuf.p1
    setUniforms_vec4 2, cbuf.colour
    setUniforms_mat4 3, cbuf.mat
End Sub

Private Sub set_constants_SHADER_TEXT(cbuf As constant_buffer_type_SHADER_TEXT)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_vec2 1, cbuf.offset
    setUniforms_vec2 2, cbuf.size
    setUniforms_vec4 3, cbuf.color
End Sub

Private Sub set_constants_SHADER_DEBUG_TEXTURE2D_QUAD(cbuf As constant_buffer_type_SHADER_DEBUG_TEXTURE2D_QUAD)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_mat4 1, cbuf.uv_transform_matrix
    setUniforms_ivec4 2, cbuf.swizzle
    setUniforms_vec4 3, cbuf.color
    setUniforms_int 4, cbuf.texture
End Sub

Private Sub set_constants_SHADER_DEBUG_TEXTURE2DARRAY_QUAD(cbuf As constant_buffer_type_SHADER_DEBUG_TEXTURE2DARRAY_QUAD)
    setUniforms_mat4 0, cbuf.transform_matrix
    setUniforms_mat4 1, cbuf.uv_transform_matrix
    setUniforms_ivec4 2, cbuf.swizzle
    setUniforms_vec4 3, cbuf.color
    setUniforms_int 4, cbuf.layer
    setUniforms_int 5, cbuf.texture
End Sub

Private Sub set_constants_SHADER_WORLD_TO_SCREEN(cbuf As constant_buffer_type_SHADER_WORLD_TO_SCREEN)
    setUniforms_float 0, cbuf.gamma
End Sub
