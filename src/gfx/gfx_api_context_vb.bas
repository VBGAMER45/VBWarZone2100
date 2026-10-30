Attribute VB_Name = "gfx_api_context_vb"
' (VB) gfx_api::context's virtual methods (gfx_api.h). There is one context,
' chosen by gfx_api::context::initialize (gfx_api_c), so the backends are
' modules, not classes: gl_context's methods are gfx_api_gl_c's gl_context_<name>
' and null_context's are gfx_api_null_c's null_context_<name>, and each
' procedure here is the virtual call: a Select Case on the current backend
' (a VB class's public methods could take no UDT). context::get().X(...) is
' gfx_api_context_X(...); the context's non-virtual members (loadTextureFromFile,
' ...) are gfx_api_c's gfx_api_context_<name>.
' Adaptations of the signatures:
' * std::vector parameters are arrays (0-based) and a count;
'   std::tuple<buffer*, size_t> lists are parallel arrays of buffers and offsets;
'   set_uniforms' blocks are parallel arrays of addresses and sizes.
' * build_pipeline takes the pipeline_create_info (gfx_api_pso_c).
' * getScreenshot's callback gets the image at once in both backends: the
'   image comes back in the ByRef parameter, and the caller runs the callback's
'   body after a True return (screen_c). setSwapInterval's completion handler
'   likewise: the caller runs it after a True return (gl_context calls it only
'   on success).
' * getBackendGameInfo's std::map<std::string, std::string> is a
'   Scripting.Dictionary.
' * optional<float> mipLodBias is has_mipLodBias and mipLodBias.
' * a const char * descr that may be nullptr is vbNullString.
Option Explicit

' gfx_api.cpp: static gfx_api::context* current_backend_context = nullptr;
Public Const CONTEXT_NONE As Long = 0
Public Const CONTEXT_NULL As Long = 1   ' null_context (gfx_api_null_c)
Public Const CONTEXT_GL As Long = 2     ' gl_context (gfx_api_gl_c)
Public current_backend_context As Long

Public Function gfx_api_context_create_texture(ByVal mipmap_count As Long, ByVal width As Long, ByVal height As Long, ByVal internal_format As pixel_format, Optional ByVal filename As String = "") As gfx_api_texture
    Select Case current_backend_context
    Case CONTEXT_GL: Set gfx_api_context_create_texture = gl_context_create_texture(mipmap_count, width, height, internal_format, filename)
    Case CONTEXT_NULL: Set gfx_api_context_create_texture = null_context_create_texture(mipmap_count, width, height, internal_format, filename)
    End Select
End Function

Public Function gfx_api_context_create_texture_array(ByVal mipmap_count As Long, ByVal layer_count As Long, ByVal width As Long, ByVal height As Long, ByVal internal_format As pixel_format, Optional ByVal filename As String = "") As gfx_api_texture_array
    Select Case current_backend_context
    Case CONTEXT_GL: Set gfx_api_context_create_texture_array = gl_context_create_texture_array(mipmap_count, layer_count, width, height, internal_format, filename)
    Case CONTEXT_NULL: Set gfx_api_context_create_texture_array = null_context_create_texture_array(mipmap_count, layer_count, width, height, internal_format, filename)
    End Select
End Function

Public Function gfx_api_context_create_buffer_object(ByVal usage As buffer_usage, Optional ByVal hint As buffer_storage_hint = buffer_storage_hint_static_draw, Optional ByVal debugName As String = "") As gfx_api_buffer
    Select Case current_backend_context
    Case CONTEXT_GL: Set gfx_api_context_create_buffer_object = gl_context_create_buffer_object(usage, hint, debugName)
    Case CONTEXT_NULL: Set gfx_api_context_create_buffer_object = null_context_create_buffer_object(usage, hint, debugName)
    End Select
End Function

Public Function gfx_api_context_build_pipeline(ByVal existing_pso As gfx_api_pipeline_state_object, createInfo As pipeline_create_info) As gfx_api_pipeline_state_object
    Select Case current_backend_context
    Case CONTEXT_GL: Set gfx_api_context_build_pipeline = gl_context_build_pipeline(existing_pso, createInfo)
    Case CONTEXT_NULL: Set gfx_api_context_build_pipeline = null_context_build_pipeline(existing_pso, createInfo)
    End Select
End Function

Public Sub gfx_api_context_bind_pipeline(ByVal pso As gfx_api_pipeline_state_object, ByVal notextures As Boolean)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_bind_pipeline pso, notextures
    Case CONTEXT_NULL: null_context_bind_pipeline pso, notextures
    End Select
End Sub

Public Sub gfx_api_context_bind_index_buffer(ByVal buffer As gfx_api_buffer, ByVal index As index_type)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_bind_index_buffer buffer, index
    Case CONTEXT_NULL: null_context_bind_index_buffer buffer, index
    End Select
End Sub

Public Sub gfx_api_context_unbind_index_buffer(ByVal buffer As gfx_api_buffer)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_unbind_index_buffer buffer
    Case CONTEXT_NULL: null_context_unbind_index_buffer buffer
    End Select
End Sub

' vertex_buffers_offset: buffers(i) (Nothing = nullptr) at offsets(i), i < count
Public Sub gfx_api_context_bind_vertex_buffers(ByVal first As Long, buffers() As gfx_api_buffer, offsets() As Long, ByVal count As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_bind_vertex_buffers first, buffers, offsets, count
    Case CONTEXT_NULL: null_context_bind_vertex_buffers first, buffers, offsets, count
    End Select
End Sub

Public Sub gfx_api_context_unbind_vertex_buffers(ByVal first As Long, buffers() As gfx_api_buffer, offsets() As Long, ByVal count As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_unbind_vertex_buffers first, buffers, offsets, count
    Case CONTEXT_NULL: null_context_unbind_vertex_buffers first, buffers, offsets, count
    End Select
End Sub

Public Sub gfx_api_context_disable_all_vertex_buffers()
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_disable_all_vertex_buffers
    Case CONTEXT_NULL: null_context_disable_all_vertex_buffers
    End Select
End Sub

Public Sub gfx_api_context_bind_streamed_vertex_buffers(ByVal data As Long, ByVal size As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_bind_streamed_vertex_buffers data, size
    Case CONTEXT_NULL: null_context_bind_streamed_vertex_buffers data, size
    End Select
End Sub

' textures(i) (Nothing = nullptr), i < textureCount
Public Sub gfx_api_context_bind_textures(texture_descriptions() As texture_input, ByVal descriptionCount As Long, textures() As gfx_api_abstract_texture, ByVal textureCount As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_bind_textures texture_descriptions, descriptionCount, textures, textureCount
    Case CONTEXT_NULL: null_context_bind_textures texture_descriptions, descriptionCount, textures, textureCount
    End Select
End Sub

' buffer: the address of the constant buffer (a gfx_api_h UDT), size: its LenB
Public Sub gfx_api_context_set_constants(ByVal buffer As Long, ByVal size As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_set_constants buffer, size
    Case CONTEXT_NULL: null_context_set_constants buffer, size
    End Select
End Sub

' uniform_blocks: blocks(i) (an address) of sizes(i) bytes, i < count
Public Sub gfx_api_context_set_uniforms(ByVal first As Long, blocks() As Long, sizes() As Long, ByVal count As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_set_uniforms first, blocks, sizes, count
    Case CONTEXT_NULL: null_context_set_uniforms first, blocks, sizes, count
    End Select
End Sub

Public Sub gfx_api_context_draw(ByVal offset As Long, ByVal count As Long, ByVal primitive As primitive_type)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_draw offset, count, primitive
    Case CONTEXT_NULL: null_context_draw offset, count, primitive
    End Select
End Sub

Public Sub gfx_api_context_draw_elements(ByVal offset As Long, ByVal count As Long, ByVal primitive As primitive_type, ByVal index As index_type)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_draw_elements offset, count, primitive, index
    Case CONTEXT_NULL: null_context_draw_elements offset, count, primitive, index
    End Select
End Sub

Public Sub gfx_api_context_set_polygon_offset(ByVal offset As Single, ByVal slope As Single)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_set_polygon_offset offset, slope
    Case CONTEXT_NULL: null_context_set_polygon_offset offset, slope
    End Select
End Sub

Public Sub gfx_api_context_set_depth_range(ByVal min As Single, ByVal max As Single)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_set_depth_range min, max
    Case CONTEXT_NULL: null_context_set_depth_range min, max
    End Select
End Sub

Public Function gfx_api_context_get_context_value(ByVal property_ As context_value) As Long
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_get_context_value = gl_context_get_context_value(property_)
    Case CONTEXT_NULL: gfx_api_context_get_context_value = null_context_get_context_value(property_)
    End Select
End Function

' uint64_t (a Double)
Public Function gfx_api_context_get_estimated_vram_mb(ByVal dedicatedOnly As Boolean) As Double
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_get_estimated_vram_mb = gl_context_get_estimated_vram_mb(dedicatedOnly)
    Case CONTEXT_NULL: gfx_api_context_get_estimated_vram_mb = null_context_get_estimated_vram_mb(dedicatedOnly)
    End Select
End Function

' the context's virtual methods with a default body (return 0 / false / nullptr, do nothing):
' gl_context overrides them, null_context does not
Public Function gfx_api_context_numDepthPasses() As Long
    If current_backend_context = CONTEXT_GL Then gfx_api_context_numDepthPasses = gl_context_numDepthPasses()
End Function

Public Function gfx_api_context_setDepthPassProperties(ByVal numDepthPasses_ As Long, ByVal depthBufferResolution As Long) As Boolean
    If current_backend_context = CONTEXT_GL Then gfx_api_context_setDepthPassProperties = gl_context_setDepthPassProperties(numDepthPasses_, depthBufferResolution)
End Function

Public Sub gfx_api_context_beginDepthPass(ByVal idx As Long)
    If current_backend_context = CONTEXT_GL Then gl_context_beginDepthPass idx
End Sub

Public Function gfx_api_context_getDepthPassDimensions(ByVal idx As Long) As Long
    If current_backend_context = CONTEXT_GL Then gfx_api_context_getDepthPassDimensions = gl_context_getDepthPassDimensions(idx)
End Function

Public Sub gfx_api_context_endCurrentDepthPass()
    If current_backend_context = CONTEXT_GL Then gl_context_endCurrentDepthPass
End Sub

Public Function gfx_api_context_getDepthTexture() As gfx_api_abstract_texture
    If current_backend_context = CONTEXT_GL Then Set gfx_api_context_getDepthTexture = gl_context_getDepthTexture()
End Function

Public Sub gfx_api_context_beginSceneRenderPass()
    If current_backend_context = CONTEXT_GL Then gl_context_beginSceneRenderPass
End Sub

Public Sub gfx_api_context_endSceneRenderPass()
    If current_backend_context = CONTEXT_GL Then gl_context_endSceneRenderPass
End Sub

Public Function gfx_api_context_getSceneTexture() As gfx_api_abstract_texture
    If current_backend_context = CONTEXT_GL Then Set gfx_api_context_getSceneTexture = gl_context_getSceneTexture()
End Function

Public Sub gfx_api_context_beginRenderPass()
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_beginRenderPass
    Case CONTEXT_NULL: null_context_beginRenderPass
    End Select
End Sub

Public Sub gfx_api_context_endRenderPass()
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_endRenderPass
    Case CONTEXT_NULL: null_context_endRenderPass
    End Select
End Sub

Public Sub gfx_api_context_debugStringMarker(ByVal str_ As String)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_debugStringMarker str_
    Case CONTEXT_NULL: null_context_debugStringMarker str_
    End Select
End Sub

Public Sub gfx_api_context_debugSceneBegin(ByVal descr As String)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_debugSceneBegin descr
    Case CONTEXT_NULL: null_context_debugSceneBegin descr
    End Select
End Sub

Public Sub gfx_api_context_debugSceneEnd(ByVal descr As String)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_debugSceneEnd descr
    Case CONTEXT_NULL: null_context_debugSceneEnd descr
    End Select
End Sub

Public Function gfx_api_context_debugPerfAvailable() As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_debugPerfAvailable = gl_context_debugPerfAvailable()
    Case CONTEXT_NULL: gfx_api_context_debugPerfAvailable = null_context_debugPerfAvailable()
    End Select
End Function

Public Function gfx_api_context_debugPerfStart(ByVal sample As Long) As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_debugPerfStart = gl_context_debugPerfStart(sample)
    Case CONTEXT_NULL: gfx_api_context_debugPerfStart = null_context_debugPerfStart(sample)
    End Select
End Function

Public Sub gfx_api_context_debugPerfStop()
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_debugPerfStop
    Case CONTEXT_NULL: null_context_debugPerfStop
    End Select
End Sub

Public Sub gfx_api_context_debugPerfBegin(ByVal pp As PERF_POINT, ByVal descr As String)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_debugPerfBegin pp, descr
    Case CONTEXT_NULL: null_context_debugPerfBegin pp, descr
    End Select
End Sub

Public Sub gfx_api_context_debugPerfEnd(ByVal pp As PERF_POINT)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_debugPerfEnd pp
    Case CONTEXT_NULL: null_context_debugPerfEnd pp
    End Select
End Sub

' uint64_t (a Double)
Public Function gfx_api_context_debugGetPerfValue(ByVal pp As PERF_POINT) As Double
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_debugGetPerfValue = gl_context_debugGetPerfValue(pp)
    Case CONTEXT_NULL: gfx_api_context_debugGetPerfValue = null_context_debugGetPerfValue(pp)
    End Select
End Function

' std::map<std::string, std::string>: a Scripting.Dictionary
Public Function gfx_api_context_getBackendGameInfo() As Object
    Select Case current_backend_context
    Case CONTEXT_GL: Set gfx_api_context_getBackendGameInfo = gl_context_getBackendGameInfo()
    Case CONTEXT_NULL: Set gfx_api_context_getBackendGameInfo = null_context_getBackendGameInfo()
    End Select
End Function

Public Function gfx_api_context_getFormattedRendererInfoString() As String
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_getFormattedRendererInfoString = gl_context_getFormattedRendererInfoString()
    Case CONTEXT_NULL: gfx_api_context_getFormattedRendererInfoString = null_context_getFormattedRendererInfoString()
    End Select
End Function

' (VB) the callback's std::unique_ptr<iV_Image> comes back in image (Nothing = nullptr)
Public Function gfx_api_context_getScreenshot(image As iV_Image) As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_getScreenshot = gl_context_getScreenshot(image)
    Case CONTEXT_NULL: gfx_api_context_getScreenshot = null_context_getScreenshot(image)
    End Select
End Function

Public Sub gfx_api_context_handleWindowSizeChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_handleWindowSizeChange oldWidth, oldHeight, newWidth, newHeight
    Case CONTEXT_NULL: null_context_handleWindowSizeChange oldWidth, oldHeight, newWidth, newHeight
    End Select
End Sub

' std::pair<uint32_t, uint32_t>: (w, h)
Public Sub gfx_api_context_getDrawableDimensions(w As Long, h As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_getDrawableDimensions w, h
    Case CONTEXT_NULL: null_context_getDrawableDimensions w, h
    End Select
End Sub

Public Function gfx_api_context_isYAxisInverted() As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_isYAxisInverted = gl_context_isYAxisInverted()
    Case CONTEXT_NULL: gfx_api_context_isYAxisInverted = null_context_isYAxisInverted()
    End Select
End Function

Public Function gfx_api_context_shouldDraw() As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_shouldDraw = gl_context_shouldDraw()
    Case CONTEXT_NULL: gfx_api_context_shouldDraw = null_context_shouldDraw()
    End Select
End Function

Public Sub gfx_api_context_shutdown()
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_shutdown
    Case CONTEXT_NULL: null_context_shutdown
    End Select
End Sub

Public Function gfx_api_context_current_FrameNum() As Long
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_current_FrameNum = gl_context_current_FrameNum()
    Case CONTEXT_NULL: gfx_api_context_current_FrameNum = null_context_current_FrameNum()
    End Select
End Function

' (VB) the completion handler: the caller runs it after a True return
Public Function gfx_api_context_setSwapInterval(ByVal mode As swap_interval_mode) As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_setSwapInterval = gl_context_setSwapInterval(mode)
    Case CONTEXT_NULL: gfx_api_context_setSwapInterval = null_context_setSwapInterval(mode)
    End Select
End Function

Public Function gfx_api_context_getSwapInterval() As swap_interval_mode
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_getSwapInterval = gl_context_getSwapInterval()
    Case CONTEXT_NULL: gfx_api_context_getSwapInterval = null_context_getSwapInterval()
    End Select
End Function

Public Function gfx_api_context_textureFormatIsSupported(ByVal target As pixel_format_target, ByVal format As pixel_format, ByVal usage As pixel_format_usage) As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_textureFormatIsSupported = gl_context_textureFormatIsSupported(target, format, usage)
    Case CONTEXT_NULL: gfx_api_context_textureFormatIsSupported = null_context_textureFormatIsSupported(target, format, usage)
    End Select
End Function

Public Function gfx_api_context_supportsMipLodBias() As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_supportsMipLodBias = gl_context_supportsMipLodBias()
    Case CONTEXT_NULL: gfx_api_context_supportsMipLodBias = null_context_supportsMipLodBias()
    End Select
End Function

Public Function gfx_api_context_supports2DTextureArrays() As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_supports2DTextureArrays = gl_context_supports2DTextureArrays()
    Case CONTEXT_NULL: gfx_api_context_supports2DTextureArrays = null_context_supports2DTextureArrays()
    End Select
End Function

Public Function gfx_api_context_supportsIntVertexAttributes() As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_supportsIntVertexAttributes = gl_context_supportsIntVertexAttributes()
    Case CONTEXT_NULL: gfx_api_context_supportsIntVertexAttributes = null_context_supportsIntVertexAttributes()
    End Select
End Function

Public Function gfx_api_context_maxFramesInFlight() As Long
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_maxFramesInFlight = gl_context_maxFramesInFlight()
    Case CONTEXT_NULL: gfx_api_context_maxFramesInFlight = null_context_maxFramesInFlight()
    End Select
End Function

Public Function gfx_api_context_getShadowConstants() As lighting_constants
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_getShadowConstants = gl_context_getShadowConstants()
    Case CONTEXT_NULL: gfx_api_context_getShadowConstants = null_context_getShadowConstants()
    End Select
End Function

Public Function gfx_api_context_setShadowConstants(values As lighting_constants) As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_setShadowConstants = gl_context_setShadowConstants(values)
    Case CONTEXT_NULL: gfx_api_context_setShadowConstants = null_context_setShadowConstants(values)
    End Select
End Function

' instanced rendering APIs
Public Function gfx_api_context_supportsInstancedRendering() As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_supportsInstancedRendering = gl_context_supportsInstancedRendering()
    Case CONTEXT_NULL: gfx_api_context_supportsInstancedRendering = null_context_supportsInstancedRendering()
    End Select
End Function

Public Sub gfx_api_context_draw_instanced(ByVal offset As Long, ByVal count As Long, ByVal primitive As primitive_type, ByVal instance_count As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_draw_instanced offset, count, primitive, instance_count
    Case CONTEXT_NULL: null_context_draw_instanced offset, count, primitive, instance_count
    End Select
End Sub

Public Sub gfx_api_context_draw_elements_instanced(ByVal offset As Long, ByVal count As Long, ByVal primitive As primitive_type, ByVal index As index_type, ByVal instance_count As Long)
    Select Case current_backend_context
    Case CONTEXT_GL: gl_context_draw_elements_instanced offset, count, primitive, index, instance_count
    Case CONTEXT_NULL: null_context_draw_elements_instanced offset, count, primitive, index, instance_count
    End Select
End Sub

' debug apis for recompiling pipelines
Public Function gfx_api_context_debugRecompileAllPipelines() As Boolean
    Select Case current_backend_context
    Case CONTEXT_GL: gfx_api_context_debugRecompileAllPipelines = gl_context_debugRecompileAllPipelines()
    Case CONTEXT_NULL: gfx_api_context_debugRecompileAllPipelines = null_context_debugRecompileAllPipelines()
    End Select
End Function

' _initialize: the backend's own; backend: which one context::initialize created (gfx_api_c)
' impl: the backend_Impl_Factory's product for it (an IGLBackend or an INullBackend)
Public Function gfx_api_context_initialize_(ByVal backend As Long, ByVal impl As Object, ByVal antialiasing As Long, ByVal mode As swap_interval_mode, ByVal has_mipLodBias As Boolean, ByVal mipLodBias As Single, ByVal depthMapResolution As Long) As Boolean
    Select Case backend
    Case CONTEXT_GL: gfx_api_context_initialize_ = gl_context_initialize_(impl, antialiasing, mode, has_mipLodBias, mipLodBias, depthMapResolution)
    Case CONTEXT_NULL: gfx_api_context_initialize_ = null_context_initialize_(impl, antialiasing, mode, has_mipLodBias, mipLodBias, depthMapResolution)
    End Select
End Function
