Attribute VB_Name = "gfx_api_null_c"
' lib/ivis_opengl/gfx_api_null.cpp, gfx_api_null.h: null_context, the null
' (headless) backend, and null_pipeline_state_object. There is one context, so
' null_context's members are this module's variables and its methods are
' null_context_<name>, called by gfx_api_context_vb's dispatch with its
' signatures (see its header for the adaptations). null_texture,
' null_texture_array and null_buffer are classes. A null_pipeline_state_object
' is a gfx_api_pipeline_state_object whose psoID is its 1-based index in
' null_pipeline_state_objects (its desc and vertex_buffer_desc).
' backend_Null_Impl is WzGfx's interface INullBackend (WzPlatform's
' win_Null_Impl implements it).
Option Explicit

' null_pipeline_state_object's fields
Private Type null_pipeline_state_object
    desc As state_description
    vertex_buffer_desc() As vertex_buffer
    vertex_buffer_desc_size As Long
End Type

Private null_pipeline_state_objects() As null_pipeline_state_object
Private null_pipeline_state_objects_size As Long

' null_context's members
Private backend_impl As INullBackend
Private current_program As Long         ' null_pipeline_state_object* (an index, 0 = nullptr)
Private formattedRendererInfoString As String ' = "Null backend (Headless mode)"
Private frameNum As Long                ' size_t = 0

' MARK: null_buffer

' null_buffer::null_buffer(usage, hint)
Public Function null_buffer_new(ByVal usage As buffer_usage, ByVal hint As buffer_storage_hint) As null_buffer
    Dim b As null_buffer
    Set b = New null_buffer
    b.usage = usage
    b.hint = hint
    ' no-op
    Set null_buffer_new = b
End Function

' MARK: null_pipeline_state_object

Private Function null_pipeline_state_object_new(desc As state_description, vertex_buffer_desc() As vertex_buffer, ByVal vertex_buffer_desc_size As Long) As gfx_api_pipeline_state_object
    Dim pso As gfx_api_pipeline_state_object, i As Long
    null_pipeline_state_objects_size = null_pipeline_state_objects_size + 1
    ReDim Preserve null_pipeline_state_objects(1 To null_pipeline_state_objects_size)
    null_pipeline_state_objects(null_pipeline_state_objects_size).desc = desc
    null_pipeline_state_objects(null_pipeline_state_objects_size).vertex_buffer_desc_size = vertex_buffer_desc_size
    If vertex_buffer_desc_size > 0 Then
        ReDim null_pipeline_state_objects(null_pipeline_state_objects_size).vertex_buffer_desc(0 To vertex_buffer_desc_size - 1)
        For i = 0 To vertex_buffer_desc_size - 1
            null_pipeline_state_objects(null_pipeline_state_objects_size).vertex_buffer_desc(i) = vertex_buffer_desc(i)
        Next i
    End If
    ' no-op
    Set pso = New gfx_api_pipeline_state_object
    pso.psoID = null_pipeline_state_objects_size
    Set null_pipeline_state_object_new = pso
End Function

' MARK: null_context

' null_context::null_context(bool _debug) {} and the member initialisers
Private Sub null_context_construct()
    current_program = 0
    formattedRendererInfoString = "Null backend (Headless mode)"
    frameNum = 0
End Sub

' null_context::~null_context(): nothing - any cleanup should occur in null_context::shutdown()

Public Function null_context_create_texture(ByVal mipmap_count As Long, ByVal width As Long, ByVal height As Long, ByVal internal_format As pixel_format, Optional ByVal filename As String = "") As gfx_api_texture
    Dim new_texture As null_texture
    Set new_texture = New null_texture
    new_texture.internal_format = internal_format
    new_texture.tex_width = width
    new_texture.tex_height = height
    Set null_context_create_texture = new_texture
End Function

Public Function null_context_create_texture_array(ByVal mipmap_count As Long, ByVal layer_count As Long, ByVal width As Long, ByVal height As Long, ByVal internal_format As pixel_format, Optional ByVal filename As String = "") As gfx_api_texture_array
    Dim new_texture As null_texture_array
    If Not (mipmap_count > 0) Then ASSERT_FAILURE "mipmap_count > 0", "lib/ivis_opengl/gfx_api_null.cpp:197", "create_texture_array", "mipmap_count must be > 0"
    Set new_texture = New null_texture_array
    new_texture.internal_format = internal_format
    Set null_context_create_texture_array = new_texture
End Function

Public Function null_context_create_buffer_object(ByVal usage As buffer_usage, Optional ByVal hint As buffer_storage_hint = buffer_storage_hint_static_draw, Optional ByVal debugName As String = "") As gfx_api_buffer
    Set null_context_create_buffer_object = null_buffer_new(usage, hint)
End Function

Public Function null_context_build_pipeline(ByVal existing_pso As gfx_api_pipeline_state_object, createInfo As pipeline_create_info) As gfx_api_pipeline_state_object
    Set null_context_build_pipeline = null_pipeline_state_object_new(createInfo.state_desc, createInfo.attribute_descriptions, createInfo.attribute_descriptions_size)
End Function

Public Sub null_context_bind_pipeline(ByVal pso As gfx_api_pipeline_state_object, ByVal notextures As Boolean)
    Dim new_program As Long
    If Not pso Is Nothing Then new_program = pso.psoID
    If current_program <> new_program Then
        current_program = new_program
    End If
End Sub

Public Sub null_context_bind_vertex_buffers(ByVal first As Long, buffers() As gfx_api_buffer, offsets() As Long, ByVal count As Long)
    Dim i As Long, e As Long, buffer As null_buffer
    If Not (current_program <> 0) Then ASSERT_FAILURE "current_program != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:224", "bind_vertex_buffers", "current_program == NULL": Exit Sub
    e = count
    i = 0
    Do While i < e
        If Not ((first + i) < null_pipeline_state_objects(current_program).vertex_buffer_desc_size) Then Exit Do
        If Not buffers(i) Is Nothing Then
            Set buffer = buffers(i)
            If Not (buffer.usage = buffer_usage_vertex_buffer) Then ASSERT_FAILURE "buffer->usage == gfx_api::buffer::usage::vertex_buffer", "lib/ivis_opengl/gfx_api_null.cpp:232", "bind_vertex_buffers", "bind_vertex_buffers called with non-vertex-buffer"
            ' no-op
        End If
        i = i + 1
    Loop
End Sub

Public Sub null_context_unbind_vertex_buffers(ByVal first As Long, buffers() As gfx_api_buffer, offsets() As Long, ByVal count As Long)
    If Not (current_program <> 0) Then ASSERT_FAILURE "current_program != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:239", "unbind_vertex_buffers", "current_program == NULL": Exit Sub
    ' no-op
End Sub

Public Sub null_context_disable_all_vertex_buffers()
    ' no-op
End Sub

Public Sub null_context_bind_streamed_vertex_buffers(ByVal data As Long, ByVal size As Long)
    If Not (current_program <> 0) Then ASSERT_FAILURE "current_program != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:250", "bind_streamed_vertex_buffers", "current_program == NULL": Exit Sub
    If Not (size > 0) Then ASSERT_FAILURE "size > 0", "lib/ivis_opengl/gfx_api_null.cpp:251", "bind_streamed_vertex_buffers", "bind_streamed_vertex_buffers called with size 0"
    ' no-op
End Sub

Public Sub null_context_bind_index_buffer(ByVal buffer_ As gfx_api_buffer, ByVal index As index_type)
    Dim buffer As null_buffer
    If Not (current_program <> 0) Then ASSERT_FAILURE "current_program != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:257", "bind_index_buffer", "current_program == NULL": Exit Sub
    Set buffer = buffer_
    If Not (buffer.usage = buffer_usage_index_buffer) Then ASSERT_FAILURE "buffer.usage == gfx_api::buffer::usage::index_buffer", "lib/ivis_opengl/gfx_api_null.cpp:259", "bind_index_buffer", "Passed gfx_api::buffer is not an index buffer"
    ' no-op
End Sub

Public Sub null_context_unbind_index_buffer(ByVal buffer As gfx_api_buffer)
    ' no-op
End Sub

Public Sub null_context_bind_textures(texture_descriptions() As texture_input, ByVal descriptionCount As Long, textures() As gfx_api_abstract_texture, ByVal textureCount As Long)
    If Not (current_program <> 0) Then ASSERT_FAILURE "current_program != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:270", "bind_textures", "current_program == NULL": Exit Sub
    If Not (textureCount <= descriptionCount) Then ASSERT_FAILURE "textures.size() <= texture_descriptions.size()", "lib/ivis_opengl/gfx_api_null.cpp:271", "bind_textures", "Received more textures than expected"
    ' no-op
End Sub

Public Sub null_context_set_constants(ByVal buffer As Long, ByVal size As Long)
    If Not (current_program <> 0) Then ASSERT_FAILURE "current_program != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:277", "set_constants", "current_program == NULL": Exit Sub
    ' no-op
End Sub

Public Sub null_context_set_uniforms(ByVal first As Long, blocks() As Long, sizes() As Long, ByVal count As Long)
    If Not (current_program <> 0) Then ASSERT_FAILURE "current_program != nullptr", "lib/ivis_opengl/gfx_api_null.cpp:283", "set_uniforms", "current_program == NULL": Exit Sub
    ' no-op
End Sub

Public Sub null_context_draw(ByVal offset As Long, ByVal count As Long, ByVal primitive As primitive_type)
    ' no-op
End Sub

Public Sub null_context_draw_instanced(ByVal offset As Long, ByVal count As Long, ByVal primitive As primitive_type, ByVal instance_count As Long)
    ' no-op
End Sub

Public Sub null_context_draw_elements(ByVal offset As Long, ByVal count As Long, ByVal primitive As primitive_type, ByVal index As index_type)
    ' no-op
End Sub

Public Sub null_context_draw_elements_instanced(ByVal offset As Long, ByVal count As Long, ByVal primitive As primitive_type, ByVal index As index_type, ByVal instance_count As Long)
    ' no-op
End Sub

Public Sub null_context_set_polygon_offset(ByVal offset As Single, ByVal slope As Single)
    ' no-op
End Sub

Public Sub null_context_set_depth_range(ByVal min As Single, ByVal max As Single)
    ' no-op
End Sub

Public Function null_context_get_context_value(ByVal property_ As context_value) As Long
    ' provide some fake, large-enough values to avoid issues
    Select Case property_
    Case context_value_MAX_ELEMENTS_VERTICES
        null_context_get_context_value = 32000
        Exit Function
    Case context_value_MAX_ELEMENTS_INDICES
        null_context_get_context_value = 32000
        Exit Function
    Case context_value_MAX_TEXTURE_SIZE
        null_context_get_context_value = 4096
        Exit Function
    Case context_value_MAX_SAMPLES
        null_context_get_context_value = 0
        Exit Function
    Case context_value_MAX_ARRAY_TEXTURE_LAYERS
        null_context_get_context_value = 2048
        Exit Function
    Case context_value_MAX_VERTEX_ATTRIBS
        null_context_get_context_value = 16
        Exit Function
    Case context_value_MAX_VERTEX_OUTPUT_COMPONENTS
        null_context_get_context_value = 64
        Exit Function
    End Select
    debug_ 337, LOG_FATAL, "get_context_value", "Unsupported property"
    null_context_get_context_value = 0
End Function

Public Function null_context_get_estimated_vram_mb(ByVal dedicatedOnly As Boolean) As Double
    null_context_get_estimated_vram_mb = 0
End Function

' MARK: null_context - debug

Public Sub null_context_debugStringMarker(ByVal str_ As String)
    ' no-op
End Sub

Public Sub null_context_debugSceneBegin(ByVal descr As String)
    ' no-op
End Sub

Public Sub null_context_debugSceneEnd(ByVal descr As String)
    ' no-op
End Sub

Public Function null_context_debugPerfAvailable() As Boolean
    null_context_debugPerfAvailable = False
End Function

Public Function null_context_debugPerfStart(ByVal sample As Long) As Boolean
    null_context_debugPerfStart = False
End Function

Public Sub null_context_debugPerfStop()
    ' no-op
End Sub

Public Sub null_context_debugPerfBegin(ByVal pp As PERF_POINT, ByVal descr As String)
    ' no-op
End Sub

Public Sub null_context_debugPerfEnd(ByVal pp As PERF_POINT)
    ' no-op
End Sub

Public Function null_context_debugGetPerfValue(ByVal pp As PERF_POINT) As Double
    null_context_debugGetPerfValue = 0
End Function

Public Function null_context_getBackendGameInfo() As Object
    Dim backendGameInfo As Object
    Set backendGameInfo = CreateObject("Scripting.Dictionary")
    backendGameInfo.CompareMode = vbBinaryCompare
    backendGameInfo("null_gfx_backend") = "true"
    Set null_context_getBackendGameInfo = backendGameInfo
End Function

Public Function null_context_getFormattedRendererInfoString() As String
    null_context_getFormattedRendererInfoString = formattedRendererInfoString
End Function

Public Function null_context_getScreenshot(image As iV_Image) As Boolean
    null_context_getScreenshot = False
End Function

Public Function null_context_initialize_(ByVal impl As Object, ByVal antialiasing As Long, ByVal mode As swap_interval_mode, ByVal has_mipLodBias As Boolean, ByVal mipLodBias As Single, ByVal depthMapResolution As Long) As Boolean
    null_context_construct
    ' obtain backend_Null_Impl from impl
    Set backend_impl = Nothing
    If Not impl Is Nothing Then
        If TypeOf impl Is INullBackend Then Set backend_impl = impl ' impl.createNullBackendImpl()
    End If
    If backend_impl Is Nothing Then
        debug_ 416, LOG_ERROR, "_initialize", "Failed to get Null / headless backend implementation"
        null_context_initialize_ = False
        Exit Function
    End If

    If Not setSwapIntervalInternal(mode) Then
        ' default to vsync on
        debug_ 423, LOG_3D, "_initialize", "Failed to set swap interval: %d; defaulting to vsync on", to_int(mode)
        setSwapIntervalInternal swap_interval_mode_vsync
    End If

    null_context_initialize_ = True
End Function

' gfx_api.h: static inline int to_int(gfx_api::context::swap_interval_mode mode)
Private Function to_int(ByVal mode As swap_interval_mode) As Long
    to_int = mode
End Function

Public Sub null_context_beginRenderPass()
    ' no-op
End Sub

Public Sub null_context_endRenderPass()
    frameNum = frameNum + 1
    If frameNum < 1 Then frameNum = 1   ' std::max<size_t>(frameNum + 1, 1)

    ' Backend is expected to handle throttling / sleeping
    backend_impl.swapWindow

    current_program = 0
End Sub

Public Sub null_context_handleWindowSizeChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ' no-op
End Sub

Public Sub null_context_getDrawableDimensions(w As Long, h As Long)
    w = 0
    h = 0
End Sub

Public Function null_context_isYAxisInverted() As Boolean
    null_context_isYAxisInverted = False
End Function

Public Function null_context_shouldDraw() As Boolean
    null_context_shouldDraw = False
End Function

Public Sub null_context_shutdown()
    ' no-op
End Sub

Public Function null_context_current_FrameNum() As Long
    null_context_current_FrameNum = frameNum
End Function

Private Function setSwapIntervalInternal(ByVal mode As swap_interval_mode) As Boolean
    setSwapIntervalInternal = backend_impl.setSwapInterval(mode)
End Function

' (VB) the completion handler: the caller runs it after a True return
Public Function null_context_setSwapInterval(ByVal mode As swap_interval_mode) As Boolean
    Dim success As Boolean
    success = setSwapIntervalInternal(mode)
    null_context_setSwapInterval = success
End Function

Public Function null_context_getSwapInterval() As swap_interval_mode
    null_context_getSwapInterval = backend_impl.getSwapInterval()
End Function

Public Function null_context_textureFormatIsSupported(ByVal target As pixel_format_target, ByVal format As pixel_format, ByVal usage As pixel_format_usage) As Boolean
    ' no matter what the input is, return true (since this null backend doesn't care and no-ops whatever it gets)
    null_context_textureFormatIsSupported = True
End Function

Public Function null_context_supportsMipLodBias() As Boolean
    null_context_supportsMipLodBias = True
End Function

Public Function null_context_supports2DTextureArrays() As Boolean
    null_context_supports2DTextureArrays = True
End Function

Public Function null_context_supportsIntVertexAttributes() As Boolean
    null_context_supportsIntVertexAttributes = True
End Function

Public Function null_context_maxFramesInFlight() As Long
    null_context_maxFramesInFlight = 1
End Function

Public Function null_context_getShadowConstants() As lighting_constants
    null_context_getShadowConstants = lighting_constants_new()
End Function

Public Function null_context_setShadowConstants(newValues As lighting_constants) As Boolean
    null_context_setShadowConstants = True
End Function

Public Function null_context_debugRecompileAllPipelines() As Boolean
    null_context_debugRecompileAllPipelines = True
End Function

Public Function null_context_supportsInstancedRendering() As Boolean
    null_context_supportsInstancedRendering = False
End Function
