VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GfxApi"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' WzGfx.dll's public interface. GlobalMultiUse: the exe and the other DLLs
' call these procedures as if they were their own. lib/ivis_opengl is ported
' into this DLL's modules (until cluster D: the stand-ins of stubs_gfx); this
' class only forwards to them.
Option Explicit

' gfx_api.h: gfx_api::context::swap_interval_mode
Public Enum swap_interval_mode
    swap_interval_mode_adaptive_vsync = -1
    swap_interval_mode_immediate = 0
    swap_interval_mode_vsync = 1
End Enum

' gfx_api.h: gfx_api::backend_type
Public Enum backend_type
    backend_type_null_backend
    backend_type_opengl_backend
    backend_type_vulkan_backend
End Enum

' pietypes.h
Public Enum LIGHTING_TYPE
    LIGHT_EMISSIVE
    LIGHT_AMBIENT
    LIGHT_DIFFUSE
    LIGHT_SPECULAR
    LIGHT_MAX
End Enum

Public Enum REND_MODE
    REND_ALPHA
    REND_ADDITIVE
    REND_OPAQUE
    REND_MULTIPLICATIVE
    REND_PREMULTIPLIED
    REND_TEXT
End Enum

Public Enum DEPTH_MODE
    DEPTH_CMP_LEQ_WRT_ON
    DEPTH_CMP_LEQ_WRT_OFF
    DEPTH_CMP_ALWAYS_WRT_ON
    DEPTH_CMP_ALWAYS_WRT_OFF
End Enum

Public Enum TEXPAGE_TYPE
    TEXPAGE_NONE = -1
    TEXPAGE_EXTERN = -2
End Enum

Public Enum SHADER_MODE
    SHADER_NONE
    SHADER_COMPONENT
    SHADER_COMPONENT_INSTANCED
    SHADER_COMPONENT_DEPTH_INSTANCED
    SHADER_NOLIGHT
    SHADER_NOLIGHT_INSTANCED
    SHADER_TERRAIN_DEPTH
    SHADER_TERRAIN_DEPTHMAP
    SHADER_WATER
    SHADER_RECT
    SHADER_RECT_INSTANCED
    SHADER_TEXRECT
    SHADER_GFX_COLOUR
    SHADER_GFX_TEXT
    SHADER_SKYBOX
    SHADER_GENERIC_COLOR
    SHADER_LINE
    SHADER_TEXT
    SHADER_TERRAIN_COMBINED_CLASSIC
    SHADER_TERRAIN_COMBINED_MEDIUM
    SHADER_TERRAIN_COMBINED_HIGH
    SHADER_WATER_CLASSIC
    SHADER_WATER_HIGH
    SHADER_WORLD_TO_SCREEN
    SHADER_DEBUG_TEXTURE2D_QUAD
    SHADER_DEBUG_TEXTURE2DARRAY_QUAD
    SHADER_MAX
End Enum

' pietypes.h: iV_Image::ColorOrder
Public Enum ColorOrder
    ColorOrder_RGB
    ColorOrder_BGR
End Enum

' gfx_api_formats_def.h: gfx_api::pixel_format
Public Enum pixel_format
    pixel_format_invalid
    pixel_format_FORMAT_RGBA8_UNORM_PACK8
    pixel_format_FORMAT_BGRA8_UNORM_PACK8
    pixel_format_FORMAT_RGB8_UNORM_PACK8
    pixel_format_FORMAT_RG8_UNORM       ' not guaranteed support
    pixel_format_FORMAT_R8_UNORM
    pixel_format_FORMAT_RGB_BC1_UNORM   ' BC1 / DXT1
    pixel_format_FORMAT_RGBA_BC2_UNORM  ' BC2 / DXT3
    pixel_format_FORMAT_RGBA_BC3_UNORM  ' BC3 / DXT5
    pixel_format_FORMAT_R_BC4_UNORM     ' BC4 / RED_RGTC1 (Interpolated greyscale)
    pixel_format_FORMAT_RG_BC5_UNORM    ' BC5 / RG_RGTC2 (Interpolated two-channel)
    pixel_format_FORMAT_RGBA_BPTC_UNORM ' BC7 - higher quality, not quite as widely available (not available at all on macOS)
    pixel_format_FORMAT_RGB8_ETC1       ' Available on most OpenGL ES 2.0+ - RGB only and not great quality
    pixel_format_FORMAT_RGB8_ETC2       ' Compresses RGB888 data (successor to ETC1)
    pixel_format_FORMAT_RGBA8_ETC2_EAC  ' Compresses RGBA8888 data with full alpha support
    pixel_format_FORMAT_R11_EAC         ' one channel unsigned data
    pixel_format_FORMAT_RG11_EAC        ' two channel unsigned data
    pixel_format_FORMAT_ASTC_4x4_UNORM
End Enum

' gfx_api::pixel_format_usage::flags
Public Enum pixel_format_usage
    pixel_format_usage_none = 0                         ' 0 (if the format is unsupported)
    pixel_format_usage_sampled_image = 1                ' 00001 == 1
    pixel_format_usage_storage_image = 2                ' 00010 == 2
    pixel_format_usage_depth_stencil_attachment = 4     ' 00100 == 4
End Enum

Public Enum pixel_format_target
    pixel_format_target_texture_2d = 0
    pixel_format_target_texture_2d_array = 1
    pixel_format_target_depth_map = 2
End Enum

Public Enum texture_type
    texture_type_user_interface
    texture_type_game_texture           ' a RGB / RGBA texture, possibly stored in a compressed format
    texture_type_alpha_mask             ' a single-channel texture, containing the alpha values
    texture_type_normal_map
    texture_type_specular_map           ' a single-channel texture, containing the specular / luma value
    texture_type_height_map             ' a single-channel texture, containing the height values
End Enum

' gfx_api.h: gfx_api::buffer::usage, gfx_api::buffer::update_flag
Public Enum buffer_usage
    buffer_usage_vertex_buffer
    buffer_usage_index_buffer
End Enum

Public Enum update_flag
    ' Default behavior
    update_flag_none
    ' This flag disables asserts caused by multiple updates to a single buffer instance in a single frame
    ' *ONLY* use if you are certain the updates are non-overlapping
    update_flag_non_overlapping_updates_promise
End Enum

Public Enum primitive_type
    primitive_type_lines
    primitive_type_line_strip
    primitive_type_triangles
    primitive_type_triangle_strip
    ' NOTE: Do *NOT* support triangle_fan, for portability reasons
End Enum

Public Enum index_type
    index_type_u16
    index_type_u32
End Enum

' gfx_api::context::buffer_storage_hint, gfx_api::context::context_value
Public Enum buffer_storage_hint
    buffer_storage_hint_static_draw
    buffer_storage_hint_stream_draw
    buffer_storage_hint_dynamic_draw
End Enum

Public Enum context_value
    context_value_MAX_ELEMENTS_VERTICES
    context_value_MAX_ELEMENTS_INDICES
    context_value_MAX_TEXTURE_SIZE
    context_value_MAX_SAMPLES           ' max antialiasing
    context_value_MAX_ARRAY_TEXTURE_LAYERS
    context_value_MAX_VERTEX_ATTRIBS
    context_value_MAX_VERTEX_OUTPUT_COMPONENTS
End Enum

' (VB) the pipeline_state_helper instantiations gfx_api.h names (gfx_api::X::get() is
' gfx_api_pso_get(X)); gfx_api_pso_c holds their template arguments
Public Enum pso_type
    Draw3DShapeOpaque
    Draw3DShapeAlpha
    Draw3DShapePremul
    Draw3DShapeAdditive
    Draw3DShapeNoLightOpaque
    Draw3DShapeNoLightAlpha
    Draw3DShapeNoLightPremul
    Draw3DShapeNoLightAdditive
    Draw3DShapeAlphaNoDepthWRT
    Draw3DShapeNoLightAlphaNoDepthWRT
    Draw3DShapeOpaque_Instanced
    Draw3DShapeAlpha_Instanced
    Draw3DShapePremul_Instanced
    Draw3DShapeAdditive_Instanced
    Draw3DShapeNoLightOpaque_Instanced
    Draw3DShapeNoLightAlpha_Instanced
    Draw3DShapeNoLightPremul_Instanced
    Draw3DShapeNoLightAdditive_Instanced
    Draw3DShapeAlphaNoDepthWRT_Instanced
    Draw3DShapeNoLightAlphaNoDepthWRT_Instanced
    Draw3DShapeAdditiveNoDepthWRT_Instanced
    Draw3DShapeNoLightAdditiveNoDepthWRT_Instanced
    Draw3DShapePremulNoDepthWRT_Instanced
    Draw3DShapeNoLightPremulNoDepthWRT_Instanced
    Draw3DShapeDepthOnly_Instanced
    TransColouredTrianglePSO
    DrawStencilShadow
    TerrainDepth
    TerrainDepthOnlyForDepthMap
    TerrainCombined_Classic
    TerrainCombined_Medium
    TerrainCombined_High
    WaterPSO
    WaterHighPSO
    WaterClassicPSO
    VideoPSO
    BackDropPSO
    SkyboxPSO
    RadarPSO
    RadarViewInsideFillPSO
    RadarViewOutlinePSO
    DrawImageTextPSO
    ShadowBox2DPSO
    UniTransBoxPSO
    DrawImagePSO
    DrawImageAnisotropicPSO
    BoxFillPSO
    BoxFillAlphaPSO
    BoxFillPSO_Instanced
    LinePSO
    DebugDrawTexture2DToQuad
    DebugDrawTexture2DArrayToQuad
    WorldToScreenPSO
    pso_type_count                      ' (VB) the number of them
End Enum

' screen.h
Public Enum PERF_POINT
    PERF_START_FRAME
    PERF_EFFECTS
    PERF_TERRAIN
    PERF_SKYBOX
    PERF_MODEL_INIT
    PERF_PARTICLES
    PERF_WATER
    PERF_MODELS
    PERF_MISC
    PERF_GUI
    PERF_COUNT
End Enum

' textdraw.h
Public Enum iV_fonts
    font_regular
    font_large
    font_medium
    font_small
    font_bar
    font_scaled
    font_regular_bold
    font_medium_bold
    font_count
End Enum

' textdraw.h: FTEXT_*
Public Enum FTEXT_JUSTIFY
    FTEXT_LEFTJUSTIFY                   ' Left justify.
    FTEXT_CENTRE                        ' Centre justify.
    FTEXT_RIGHTJUSTIFY                  ' Right justify.
End Enum

' pieblitfunc.h
Public Enum GFXTYPE
    GFX_TEXTURE
    GFX_COLOUR
    GFX_COUNT
End Enum

' (VB) the GenerateDefaultTextureFunc lambdas the game passes to loadTextureArrayFromFiles
' (terrain.cpp, texture.cpp); gfx_api_c runs them
Public Enum GenerateDefaultTextureFunc
    GenerateDefaultTexture_none = 0         ' nullptr
    GenerateDefaultTexture_zeroed = 1       ' allocate(width, height, channels, true): the default specular / height map, the blank decal
    GenerateDefaultTexture_normalMap = 2    ' the default normal map: (0,0,1) (0x7f, blue 0xff)
End Enum

' (VB) the progress callbacks the game passes to loadTextureArrayFromFiles
Public Enum ProgressCallbackFunc
    ProgressCallback_none = 0               ' nullptr
    ProgressCallback_resDoResLoadCallback = 1 ' []() { resDoResLoadCallback(); }
End Enum

' piepalette.h: the palette's indices (C++'s WZCOL_X is psPalette[n]: here psPalette(WZCOL_X))
Public Enum WZCOL
    WZCOL_BLACK = 0
    WZCOL_WHITE = 1
    WZCOL_RELOAD_BACKGROUND = 2
    WZCOL_RELOAD_BAR = 3
    WZCOL_HEALTH_HIGH = 4
    WZCOL_HEALTH_MEDIUM = 5
    WZCOL_HEALTH_LOW = 6
    WZCOL_GREEN = 7
    WZCOL_RED = 8
    WZCOL_YELLOW = 9
    WZCOL_MENU_BACKGROUND = 10
    WZCOL_MENU_BORDER = 11
    WZCOL_MENU_LOAD_BORDER = 12
    WZCOL_CURSOR = 13
    WZCOL_MENU_SCORES_INTERIOR = 14
    WZCOL_MENU_SEPARATOR = 15
    WZCOL_TEXT_BRIGHT = 16
    WZCOL_TEXT_MEDIUM = 17
    WZCOL_TEXT_DARK = 18
    WZCOL_SCORE_BOX_BORDER = 19
    WZCOL_SCORE_BOX = 20
    WZCOL_TOOLTIP_TEXT = 21
    WZCOL_UNIT_SELECT_BORDER = 22
    WZCOL_UNIT_SELECT_BOX = 23
    WZCOL_RADAR_BACKGROUND = 24
    WZCOL_MAP_OUTLINE_OK = 25
    WZCOL_MAP_OUTLINE_BAD = 26
    WZCOL_KEYMAP_ACTIVE = 27
    WZCOL_KEYMAP_FIXED = 28
    WZCOL_MENU_SCORE_LOSS = 29
    WZCOL_MENU_SCORE_DESTROYED = 30
    WZCOL_MENU_SCORE_BUILT = 31
    WZCOL_MENU_SCORE_RANK = 32
    WZCOL_FRAME_BORDER_NORMAL = 33
    WZCOL_CONS_TEXT_SYSTEM = 34
    WZCOL_CONS_TEXT_USER = 35
    WZCOL_CONS_TEXT_USER_ALLY = 36
    WZCOL_CONS_TEXT_USER_ENEMY = 37
    WZCOL_CONS_TEXT_DEBUG = 38
    WZCOL_GREY = 39
    WZCOL_MAP_PREVIEW_AIPLAYER = 40
    WZCOL_MENU_SHADOW = 41
    WZCOL_DBLUE = 42
    WZCOL_LBLUE = 43
    WZCOL_LGREEN = 44
    WZCOL_LRED = 45
    WZCOL_BLUEPRINT_PLANNED = 46
    WZCOL_HEALTH_HIGH_SHADOW = 47
    WZCOL_HEALTH_MEDIUM_SHADOW = 48
    WZCOL_HEALTH_LOW_SHADOW = 49
    WZCOL_HEALTH_RESISTANCE = 50
    WZCOL_TEAM1 = 51
    WZCOL_TEAM2 = 52
    WZCOL_TEAM3 = 53
    WZCOL_TEAM4 = 54
    WZCOL_TEAM5 = 55
    WZCOL_TEAM6 = 56
    WZCOL_TEAM7 = 57
    WZCOL_TEAM8 = 58
    WZCOL_FORM_BACKGROUND = 59
    WZCOL_FORM_TEXT = 60
    WZCOL_FORM_LIGHT = 61
    WZCOL_FORM_DARK = 62
    WZCOL_FORM_HILITE = 63
    WZCOL_FORM_CURSOR = 64
    WZCOL_FORM_TIP_BACKGROUND = 65
    WZCOL_FORM_DISABLE = 66
    WZCOL_DESIGN_POWER_FORM_BACKGROUND = 67
    WZCOL_POWER_BAR = 68
    WZCOL_ACTION_PROGRESS_BAR_MAJOR = 69
    WZCOL_ACTION_PROGRESS_BAR_MINOR = 70
    WZCOL_ACTION_PRODUCTION_RUN_TEXT = 71
    WZCOL_ACTION_PRODUCTION_RUN_BACKGROUND = 72
    WZCOL_LOADING_BAR_BACKGROUND = 73
    WZCOL_MAP_PREVIEW_HQ = 74
    WZCOL_MAP_PREVIEW_OIL = 75
    WZCOL_FOG = 76
    WZCOL_TEAM9 = 77
    WZCOL_TEAM10 = 78
    WZCOL_TEAM11 = 79
    WZCOL_TEAM12 = 80
    WZCOL_TEAM13 = 81
    WZCOL_TEAM14 = 82
    WZCOL_TEAM15 = 83
    WZCOL_TEAM16 = 84
    WZCOL_BLUEPRINT_PLANNED_BY_ALLY = 85
    WZCOL_CONSTRUCTION_BARTEXT = 86
    WZCOL_POWERQUEUE_BARTEXT = 87
    WZCOL_MAP_PREVIEW_BARREL = 88
    WZCOL_CONSTRUCTOR_BEAM = 89
    WZCOL_DEMOLISH_BEAM = 90
    WZCOL_FORM_PLAYER_NOPING = 91
    WZCOL_FORM_PLAYER_UNKNOWN = 92
    WZCOL_FORM_PLAYER_KNOWN = 93
    WZCOL_FORM_PLAYER_KNOWN_BY_SERVER = 94
    WZCOL_TRANSPARENT_BOX = 95
    WZCOL_CONS_TEXT_INFO = 96
    WZCOL_NOTIFICATION_BOX = 97
    WZCOL_FOG_ARIZONA = 98
    WZCOL_FOG_URBAN = 99
    WZCOL_FOG_ROCKIE = 100
    WZCOL_MAX = 101
End Enum

' pieblitfunc.h
Public Enum SCREENTYPE
    SCREEN_RANDOMBDROP
    SCREEN_MISSIONEND
End Enum

' piedef.h: enum class ShadowMode (piedraw.cpp, cluster E; piestate_c's pie_LoadShaders uses it)
Public Enum ShadowMode
    ShadowMode_Fallback_Stencil_Shadows
    ShadowMode_Shadow_Mapping
End Enum

' pietypes.h: the render style flags for all pie draw functions, pie_RAISE_SCALE, LONG_WAY, BUTTON_DEPTH
' (#defines; pie_SHIELD_FACTOR, a float, is piedraw_c's own Private Const)
Public Enum pietypes_defines
    pie_ECM = &H1
    pie_TRANSLUCENT = &H2
    pie_ADDITIVE = &H4
    pie_FORCE_FOG = &H8
    pie_HEIGHT_SCALED = &H10
    pie_RAISE = &H20
    pie_BUTTON = &H40
    pie_SHADOW = &H80
    pie_STATIC_SHADOW = &H100
    pie_PREMULTIPLIED = &H200
    pie_NODEPTHWRITE = &H400
    pie_FORCELIGHT = &H800
    pie_SHIELD = &H1000
    pie_RAISE_SCALE = 256
    LONG_WAY = &H8000&                  ' (1<<15)
    BUTTON_DEPTH = 2000                 ' will be stretched to 16000
End Enum

' Name and version of this DLL and of the WzFramework it uses, for the log header.
Public Function GfxBuildInfo() As String
    GfxBuildInfo = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision & " (uses " & FrameworkBuildInfo() & ")"
End Function

' impl stands for C++'s backend_Impl_Factory: the implementation the factory
' would create for backendType (an IGLBackend for opengl_backend, an
' INullBackend for null_backend).
Public Function gfx_api_context_initialize(ByVal impl As Object, ByVal antialiasing As Long, ByVal mode As swap_interval_mode, ByVal mipLodBias As Single, ByVal depthMapResolution As Long, ByVal backendType As backend_type) As Boolean
    ' (VB) optional<float> mipLodBias: WzPlatform passes 0 for nullopt (wzMainScreenSetup sets it only for a percentage <> 0)
    gfx_api_context_initialize = gfx_api_c.gfx_api_context_initialize(impl, antialiasing, mode, mipLodBias <> 0!, mipLodBias, depthMapResolution, backendType)
End Function

Public Sub gfx_api_context_shutdown()
    gfx_api_context_vb.gfx_api_context_shutdown
End Sub

Public Sub gfx_api_context_handleWindowSizeChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    gfx_api_context_vb.gfx_api_context_handleWindowSizeChange oldWidth, oldHeight, newWidth, newHeight
End Sub

Public Sub pie_SetVideoBufferWidth(ByVal width As Long)
    pieclip_c.pie_SetVideoBufferWidth width
End Sub

Public Sub pie_SetVideoBufferHeight(ByVal height As Long)
    pieclip_c.pie_SetVideoBufferHeight height
End Sub

Public Function pie_GetVideoBufferWidth() As Long
    pie_GetVideoBufferWidth = pieclip_c.pie_GetVideoBufferWidth()
End Function

Public Function pie_GetVideoBufferHeight() As Long
    pie_GetVideoBufferHeight = pieclip_c.pie_GetVideoBufferHeight()
End Function

Public Sub pie_UpdateSurfaceGeometry()
    piemode_c.pie_UpdateSurfaceGeometry
End Sub

Public Sub screen_updateGeometry()
    screen_c.screen_updateGeometry
End Sub

Public Sub pie_ScreenFrameRenderBegin()
    piemode_c.pie_ScreenFrameRenderBegin
End Sub

Public Sub pie_ScreenFrameRenderEnd()
    piemode_c.pie_ScreenFrameRenderEnd
End Sub

'--------------------------------------------------------------------------
' piestate.cpp, screen.cpp

Public Sub pie_EnableFog(ByVal val As Boolean)
    piestate_c.pie_EnableFog val
End Sub

Public Function pie_GetFogEnabled() As Boolean
    pie_GetFogEnabled = piestate_c.pie_GetFogEnabled()
End Function

Public Function pie_GetFogStatus() As Boolean
    pie_GetFogStatus = piestate_c.pie_GetFogStatus()
End Function

Public Sub pie_SetFogStatus(ByVal val As Boolean)
    piestate_c.pie_SetFogStatus val
End Sub

Public Property Get wz_texture_compression() As Boolean
    wz_texture_compression = screen_c.get_wz_texture_compression()
End Property

Public Property Let wz_texture_compression(ByVal value As Boolean)
    screen_c.set_wz_texture_compression value
End Property

'--------------------------------------------------------------------------
' gfx_api.cpp, png_util.cpp

Public Property Get uses_gfx_debug() As Boolean
    uses_gfx_debug = gfx_api_c.uses_gfx_debug
End Property

Public Property Let uses_gfx_debug(ByVal value As Boolean)
    gfx_api_c.uses_gfx_debug = value
End Property

' (REPLACES the existing forwarder to stubs_gfx)
Public Function iV_LoadAndSavePNG_AsLumaSingleChannel(ByVal inputFilename As String, ByVal outputFilename As String, Optional ByVal check As Boolean = False) As Boolean
    iV_LoadAndSavePNG_AsLumaSingleChannel = png_util_gdip.iV_LoadAndSavePNG_AsLumaSingleChannel(inputFilename, outputFilename, check)
End Function

' (the result is dropped: init.cpp ignores it)
Public Sub gfx_api_loadTextureCompressionOverrides()
    gfx_api_c.gfx_api_loadTextureCompressionOverrides
End Sub

' gfx_api::imageLoadFilenameFromInputFilename
Public Function gfx_api_imageLoadFilenameFromInputFilename(ByVal filename As String) As String
    gfx_api_imageLoadFilenameFromInputFilename = gfx_api_c.gfx_api_imageLoadFilenameFromInputFilename(filename)
End Function

'--------------------------------------------------------------------------
' stand-ins (stubs_gfx): piemode.h, pieblitfunc.h, textdraw.h, screen.h, piepalette.h

Public Function pie_Initialise() As Boolean
    pie_Initialise = piemode_c.pie_Initialise()
End Function

Public Function pie_InitRadar() As Boolean
    pie_InitRadar = pieblitfunc_c.pie_InitRadar()
End Function

Public Function pie_ShutdownRadar() As Boolean
    pie_ShutdownRadar = pieblitfunc_c.pie_ShutdownRadar()
End Function

Public Sub iV_TextInit(ByVal horizScalePercentage As Long, ByVal vertScalePercentage As Long)
    textdraw_gdi.iV_TextInit horizScalePercentage, vertScalePercentage
End Sub

' descr vbNullString: nullptr
Public Sub wzSceneBegin(ByVal descr As String)
    screen_c.wzSceneBegin descr
End Sub

Public Sub wzSceneEnd(ByVal descr As String)
    screen_c.wzSceneEnd descr
End Sub

' (REPLACES the existing forwarder to stubs_gfx)
Public Sub pal_ShutDown()
    piepalette_c.pal_ShutDown
End Sub

Public Sub screenShutDown()
    screen_c.screenShutDown
End Sub

'--------------------------------------------------------------------------
' imd.h (imdload_c). A model is a handle: an iIMDBaseShape * or iIMDShape *
' (0 = nullptr); the iIMDBaseShape_ procedures read a base shape's fields.

Public Sub modelShutdown()
    imdload_c.modelShutdown
End Sub

Public Sub modelUpdateTilesetIdx(ByVal tilesetIdx As Long)
    imdload_c.modelUpdateTilesetIdx tilesetIdx
End Sub

' enumerateLoadedModels: the number of models, and the i-th (0-based) one's name and base shape
Public Function enumerateLoadedModels_count() As Long
    enumerateLoadedModels_count = imdload_c.enumerateLoadedModels_count()
End Function

Public Function enumerateLoadedModels_at(ByVal i As Long, modelName_ As String) As Long
    enumerateLoadedModels_at = imdload_c.enumerateLoadedModels_at(i, modelName_)
End Function

' model: an iIMDShape *
Public Function modelName(ByVal model As Long) As String
    modelName = imdload_c.modelName(model)
End Function

' an iIMDBaseShape *
Public Function modelGet(ByVal filename As String) As Long
    modelGet = imdload_c.modelGet(filename)
End Function

Public Function debugReloadDisplayModelsForBaseModel(ByVal baseModel As Long) As Boolean
    debugReloadDisplayModelsForBaseModel = imdload_c.debugReloadDisplayModelsForBaseModel(baseModel)
End Function

Public Sub debugReloadAllDisplayModels()
    imdload_c.debugReloadAllDisplayModels
End Sub

Public Function getModelLoadingErrorCount() As Long
    getModelLoadingErrorCount = imdload_c.getModelLoadingErrorCount()
End Function

Public Function getModelTextureLoadingFailuresCount() As Long
    getModelTextureLoadingFailuresCount = imdload_c.getModelTextureLoadingFailuresCount()
End Function

' iIMDBaseShape::displayModel(): an iIMDShape *
Public Function iIMDBaseShape_displayModel(ByVal baseShape As Long) As Long
    iIMDBaseShape_displayModel = imdload_c.iIMDBaseShape_displayModel(baseShape)
End Function

Public Function iIMDBaseShape_filename(ByVal baseShape As Long) As String
    iIMDBaseShape_filename = imdload_c.iIMDBaseShape_filename(baseShape)
End Function

' iIMDShape::objanimpie[animEvent] (an iIMDBaseShape *), objanimcycles, objanimtime
Public Function iIMDShape_objanimpie(ByVal shape As Long, ByVal animEvent As Long) As Long
    iIMDShape_objanimpie = imdload_c.iIMDShape_objanimpie(shape, animEvent)
End Function

Public Function iIMDShape_objanimcycles(ByVal shape As Long) As Long
    iIMDShape_objanimcycles = imdload_c.iIMDShape_objanimcycles(shape)
End Function

Public Function iIMDShape_objanimtime(ByVal shape As Long) As Long
    iIMDShape_objanimtime = imdload_c.iIMDShape_objanimtime(shape)
End Function

' iIMDShape::objanimframes, next (an iIMDShape *, 0 = nullptr), points.size(), points[i]
Public Function iIMDShape_objanimframes(ByVal shape As Long) As Long
    iIMDShape_objanimframes = imdload_c.iIMDShape_objanimframes(shape)
End Function

Public Function iIMDShape_next(ByVal shape As Long) As Long
    iIMDShape_next = imdload_c.iIMDShape_next(shape)
End Function

Public Function iIMDShape_points_size(ByVal shape As Long) As Long
    iIMDShape_points_size = imdload_c.iIMDShape_points_size(shape)
End Function

Public Sub iIMDShape_point(ByVal shape As Long, ByVal i As Long, x As Single, y As Single, z As Single)
    imdload_c.iIMDShape_point shape, i, x, y, z
End Sub

Public Function iIMDBaseShape_path(ByVal baseShape As Long) As String
    iIMDBaseShape_path = imdload_c.iIMDBaseShape_path(baseShape)
End Function

Public Function iIMDBaseShape_sradius(ByVal baseShape As Long) As Long
    iIMDBaseShape_sradius = imdload_c.iIMDBaseShape_sradius(baseShape)
End Function

Public Function iIMDBaseShape_radius(ByVal baseShape As Long) As Long
    iIMDBaseShape_radius = imdload_c.iIMDBaseShape_radius(baseShape)
End Function

Public Sub iIMDBaseShape_min(ByVal baseShape As Long, x As Long, y As Long, z As Long)
    imdload_c.iIMDBaseShape_min baseShape, x, y, z
End Sub

Public Sub iIMDBaseShape_max(ByVal baseShape As Long, x As Long, y As Long, z As Long)
    imdload_c.iIMDBaseShape_max baseShape, x, y, z
End Sub

Public Sub iIMDBaseShape_ocen(ByVal baseShape As Long, x As Single, y As Single, z As Single)
    imdload_c.iIMDBaseShape_ocen baseShape, x, y, z
End Sub

Public Function iIMDBaseShape_connectors_size(ByVal baseShape As Long) As Long
    iIMDBaseShape_connectors_size = imdload_c.iIMDBaseShape_connectors_size(baseShape)
End Function

Public Sub iIMDBaseShape_connector(ByVal baseShape As Long, ByVal i As Long, x As Long, y As Long, z As Long)
    imdload_c.iIMDBaseShape_connector baseShape, i, x, y, z
End Sub

' (VB) the model as text, for the tests
Public Function modelDump(ByVal baseShape As Long) As String
    modelDump = imdload_c.modelDump(baseShape)
End Function

'--------------------------------------------------------------------------
' cluster D (fork A): gfx_api.cpp, gfx_api.h, gfx_api_null.cpp

Public Function gfx_api_context_isInitialized() As Boolean
    gfx_api_context_isInitialized = gfx_api_c.gfx_api_context_isInitialized()
End Function

' gfx_api::X::get() for the pipeline_state_helper instantiation X
Public Function gfx_api_pso_get(ByVal pso As pso_type) As pipeline_state_helper
    Set gfx_api_pso_get = gfx_api_pso_c.gfx_api_pso_get(pso)
End Function

Public Function gfx_api_checkImageFilesWouldLoadFromSameParentMountPath(filenames() As String, ByVal filenames_size As Long, ByVal ignoreNotFound As Boolean) As Boolean
    gfx_api_checkImageFilesWouldLoadFromSameParentMountPath = gfx_api_c.gfx_api_checkImageFilesWouldLoadFromSameParentMountPath(filenames, filenames_size, ignoreNotFound)
End Function

' the context's high-level loaders (Nothing = nullptr)
Public Function gfx_api_context_loadTextureFromFile(ByVal filename As String, ByVal textureType As texture_type, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1, Optional ByVal quiet As Boolean = False) As gfx_api_texture
    Set gfx_api_context_loadTextureFromFile = gfx_api_c.gfx_api_context_loadTextureFromFile(filename, textureType, maxWidth, maxHeight, quiet)
End Function

' image: the iV_Image&& (ownership passes: the caller drops it)
Public Function gfx_api_context_loadTextureFromUncompressedImage(ByVal image As iV_Image, ByVal textureType As texture_type, ByVal filename As String, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1) As gfx_api_texture
    Set gfx_api_context_loadTextureFromUncompressedImage = gfx_api_c.gfx_api_context_loadTextureFromUncompressedImage(image, textureType, filename, maxWidth, maxHeight)
End Function

' filenames: filenames_size names (0-based; "" = the default texture for that layer)
Public Function gfx_api_context_loadTextureArrayFromFiles(filenames() As String, ByVal filenames_size As Long, ByVal textureType As texture_type, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1, Optional ByVal defaultTextureGenerator As GenerateDefaultTextureFunc = GenerateDefaultTexture_none, Optional ByVal progressCallback As ProgressCallbackFunc = ProgressCallback_none, Optional ByVal debugName As String = "") As gfx_api_texture_array
    Set gfx_api_context_loadTextureArrayFromFiles = gfx_api_c.gfx_api_context_loadTextureArrayFromFiles(filenames, filenames_size, textureType, maxWidth, maxHeight, defaultTextureGenerator, progressCallback, debugName)
End Function

' the context's virtual methods (gfx_api_context_vb)
Public Function gfx_api_context_create_texture(ByVal mipmap_count As Long, ByVal width As Long, ByVal height As Long, ByVal internal_format As pixel_format, Optional ByVal filename As String = "") As gfx_api_texture
    Set gfx_api_context_create_texture = gfx_api_context_vb.gfx_api_context_create_texture(mipmap_count, width, height, internal_format, filename)
End Function

Public Function gfx_api_context_create_buffer_object(ByVal usage As buffer_usage, Optional ByVal hint As buffer_storage_hint = buffer_storage_hint_static_draw, Optional ByVal debugName As String = "") As gfx_api_buffer
    Set gfx_api_context_create_buffer_object = gfx_api_context_vb.gfx_api_context_create_buffer_object(usage, hint, debugName)
End Function

Public Sub gfx_api_context_bind_index_buffer(ByVal buffer As gfx_api_buffer, ByVal index As index_type)
    gfx_api_context_vb.gfx_api_context_bind_index_buffer buffer, index
End Sub

Public Sub gfx_api_context_unbind_index_buffer(ByVal buffer As gfx_api_buffer)
    gfx_api_context_vb.gfx_api_context_unbind_index_buffer buffer
End Sub

Public Sub gfx_api_context_set_polygon_offset(ByVal offset As Single, ByVal slope As Single)
    gfx_api_context_vb.gfx_api_context_set_polygon_offset offset, slope
End Sub

Public Function gfx_api_context_get_context_value(ByVal property_ As context_value) As Long
    gfx_api_context_get_context_value = gfx_api_context_vb.gfx_api_context_get_context_value(property_)
End Function

' uint64_t (a Double)
Public Function gfx_api_context_get_estimated_vram_mb(ByVal dedicatedOnly As Boolean) As Double
    gfx_api_context_get_estimated_vram_mb = gfx_api_context_vb.gfx_api_context_get_estimated_vram_mb(dedicatedOnly)
End Function

Public Function gfx_api_context_numDepthPasses() As Long
    gfx_api_context_numDepthPasses = gfx_api_context_vb.gfx_api_context_numDepthPasses()
End Function

Public Sub gfx_api_context_beginDepthPass(ByVal idx As Long)
    gfx_api_context_vb.gfx_api_context_beginDepthPass idx
End Sub

Public Function gfx_api_context_getDepthPassDimensions(ByVal idx As Long) As Long
    gfx_api_context_getDepthPassDimensions = gfx_api_context_vb.gfx_api_context_getDepthPassDimensions(idx)
End Function

Public Sub gfx_api_context_endCurrentDepthPass()
    gfx_api_context_vb.gfx_api_context_endCurrentDepthPass
End Sub

Public Function gfx_api_context_getDepthTexture() As gfx_api_abstract_texture
    Set gfx_api_context_getDepthTexture = gfx_api_context_vb.gfx_api_context_getDepthTexture()
End Function

Public Sub gfx_api_context_beginSceneRenderPass()
    gfx_api_context_vb.gfx_api_context_beginSceneRenderPass
End Sub

Public Sub gfx_api_context_endSceneRenderPass()
    gfx_api_context_vb.gfx_api_context_endSceneRenderPass
End Sub

Public Function gfx_api_context_getSceneTexture() As gfx_api_abstract_texture
    Set gfx_api_context_getSceneTexture = gfx_api_context_vb.gfx_api_context_getSceneTexture()
End Function

Public Sub gfx_api_context_debugStringMarker(ByVal str_ As String)
    gfx_api_context_vb.gfx_api_context_debugStringMarker str_
End Sub

' std::map<std::string, std::string>: a Scripting.Dictionary
Public Function gfx_api_context_getBackendGameInfo() As Object
    Set gfx_api_context_getBackendGameInfo = gfx_api_context_vb.gfx_api_context_getBackendGameInfo()
End Function

Public Function gfx_api_context_getFormattedRendererInfoString() As String
    gfx_api_context_getFormattedRendererInfoString = gfx_api_context_vb.gfx_api_context_getFormattedRendererInfoString()
End Function

' std::pair<uint32_t, uint32_t>: (w, h)
Public Sub gfx_api_context_getDrawableDimensions(w As Long, h As Long)
    gfx_api_context_vb.gfx_api_context_getDrawableDimensions w, h
End Sub

Public Function gfx_api_context_shouldDraw() As Boolean
    gfx_api_context_shouldDraw = gfx_api_context_vb.gfx_api_context_shouldDraw()
End Function

' (VB) the completion handler: the caller runs it after a True return
Public Function gfx_api_context_setSwapInterval(ByVal mode As swap_interval_mode) As Boolean
    gfx_api_context_setSwapInterval = gfx_api_context_vb.gfx_api_context_setSwapInterval(mode)
End Function

Public Function gfx_api_context_getSwapInterval() As swap_interval_mode
    gfx_api_context_getSwapInterval = gfx_api_context_vb.gfx_api_context_getSwapInterval()
End Function

Public Function gfx_api_context_supports2DTextureArrays() As Boolean
    gfx_api_context_supports2DTextureArrays = gfx_api_context_vb.gfx_api_context_supports2DTextureArrays()
End Function

Public Function gfx_api_context_supportsIntVertexAttributes() As Boolean
    gfx_api_context_supportsIntVertexAttributes = gfx_api_context_vb.gfx_api_context_supportsIntVertexAttributes()
End Function

' gfx_api::lighting_constants travels as its fields
Public Sub gfx_api_context_getShadowConstants(shadowMode_ As Long, shadowFilterSize As Long, shadowCascadesCount As Long, isPointLightPerPixelEnabled As Boolean)
    Dim values As lighting_constants
    values = gfx_api_context_vb.gfx_api_context_getShadowConstants()
    shadowMode_ = values.shadowMode
    shadowFilterSize = values.shadowFilterSize
    shadowCascadesCount = values.shadowCascadesCount
    isPointLightPerPixelEnabled = values.isPointLightPerPixelEnabled
End Sub

Public Function gfx_api_context_setShadowConstants(ByVal shadowMode_ As Long, ByVal shadowFilterSize As Long, ByVal shadowCascadesCount As Long, ByVal isPointLightPerPixelEnabled As Boolean) As Boolean
    Dim values As lighting_constants
    values.shadowMode = shadowMode_
    values.shadowFilterSize = shadowFilterSize
    values.shadowCascadesCount = shadowCascadesCount
    values.isPointLightPerPixelEnabled = isPointLightPerPixelEnabled
    gfx_api_context_setShadowConstants = gfx_api_context_vb.gfx_api_context_setShadowConstants(values)
End Function

Public Function gfx_api_context_debugRecompileAllPipelines() As Boolean
    gfx_api_context_debugRecompileAllPipelines = gfx_api_context_vb.gfx_api_context_debugRecompileAllPipelines()
End Function

'--------------------------------------------------------------------------
' cluster D (fork I): pieimage.cpp, png_util, tex.cpp, bitimage.cpp, piepalette.cpp

Public Function iV_loadImage_PNG(ByVal fileName As String, ByVal image As iV_Image) As Boolean
    iV_loadImage_PNG = png_util_gdip.iV_loadImage_PNG(fileName, image)
End Function

Public Function iV_loadImage_PNG2(ByVal fileName As String, ByVal image As iV_Image, Optional ByVal forceRGBA8 As Boolean = False, Optional ByVal quietOnOpenFail As Boolean = False) As Boolean
    iV_loadImage_PNG2 = png_util_gdip.iV_loadImage_PNG2(fileName, image, forceRGBA8, quietOnOpenFail)
End Function

' IMGSaveError iV_loadImage_PNG(const std::vector<unsigned char>& memoryBuffer, iV_Image *image): memoryBuffer(0 .. UBound)
Public Function iV_loadImage_PNG_mem(memoryBuffer() As Byte, ByVal image As iV_Image) As String
    iV_loadImage_PNG_mem = png_util_gdip.iV_loadImage_PNG_mem(memoryBuffer, image)
End Function

Public Function iV_saveImage_PNG(ByVal fileName As String, ByVal image As iV_Image) As String
    iV_saveImage_PNG = png_util_gdip.iV_saveImage_PNG(fileName, image)
End Function

Public Function iV_saveImage_PNG_Gray(ByVal fileName As String, ByVal image As iV_Image) As String
    iV_saveImage_PNG_Gray = png_util_gdip.iV_saveImage_PNG_Gray(fileName, image)
End Function

Public Sub pie_TexInit()
    tex_c.pie_TexInit
End Sub

Public Sub pie_TexShutDown()
    tex_c.pie_TexShutDown
End Sub

Public Function pie_NumberOfPages() As Long
    pie_NumberOfPages = tex_c.pie_NumberOfPages()
End Function

Public Function replaceTexture(ByVal oldfile As String, ByVal newfile As String) As Boolean
    replaceTexture = tex_c.replaceTexture(oldfile, newfile)
End Function

Public Function iV_LoadImageFile(ByVal fileName As String) As Long
    iV_LoadImageFile = bitimage_c.iV_LoadImageFile(fileName)
End Function

Public Sub iV_FreeImageFile(ByVal imgFile As Long)
    bitimage_c.iV_FreeImageFile imgFile
End Sub

Public Function iV_GetImage(ByVal filename As String) As Long
    iV_GetImage = bitimage_c.iV_GetImage(filename)
End Function

Public Function iV_GetImageWidth(ByVal imgFile As Long, ByVal ID As Long) As Long
    iV_GetImageWidth = bitimage_c.iV_GetImageWidth(imgFile, ID)
End Function

Public Function iV_GetImageHeight(ByVal imgFile As Long, ByVal ID As Long) As Long
    iV_GetImageHeight = bitimage_c.iV_GetImageHeight(imgFile, ID)
End Function

Public Function iV_GetImageXOffset(ByVal imgFile As Long, ByVal ID As Long) As Long
    iV_GetImageXOffset = bitimage_c.iV_GetImageXOffset(imgFile, ID)
End Function

Public Function iV_GetImageYOffset(ByVal imgFile As Long, ByVal ID As Long) As Long
    iV_GetImageYOffset = bitimage_c.iV_GetImageYOffset(imgFile, ID)
End Function

' IMAGEFILE::numImages(), IMAGEFILE::find(name)
Public Function IMAGEFILE_numImages(ByVal imgFile As Long) As Long
    IMAGEFILE_numImages = bitimage_c.IMAGEFILE_numImages(imgFile)
End Function

Public Function IMAGEFILE_find(ByVal imgFile As Long, ByVal filename As String) As Long
    IMAGEFILE_find = bitimage_c.IMAGEFILE_find(imgFile, filename)
End Function

Public Function AtlasImageDef_Width(ByVal def As Long) As Long
    AtlasImageDef_Width = bitimage_c.AtlasImageDef_Width(def)
End Function

Public Function AtlasImageDef_Height(ByVal def As Long) As Long
    AtlasImageDef_Height = bitimage_c.AtlasImageDef_Height(def)
End Function

Public Function AtlasImageDef_XOffset(ByVal def As Long) As Long
    AtlasImageDef_XOffset = bitimage_c.AtlasImageDef_XOffset(def)
End Function

Public Function AtlasImageDef_YOffset(ByVal def As Long) As Long
    AtlasImageDef_YOffset = bitimage_c.AtlasImageDef_YOffset(def)
End Function

Public Function AtlasImage_width(ByVal images As Long, ByVal id As Long) As Long
    AtlasImage_width = bitimage_c.AtlasImage_width(images, id)
End Function

Public Function AtlasImage_height(ByVal images As Long, ByVal id As Long) As Long
    AtlasImage_height = bitimage_c.AtlasImage_height(images, id)
End Function

Public Function AtlasImage_xOffset(ByVal images As Long, ByVal id As Long) As Long
    AtlasImage_xOffset = bitimage_c.AtlasImage_xOffset(images, id)
End Function

Public Function AtlasImage_yOffset(ByVal images As Long, ByVal id As Long) As Long
    AtlasImage_yOffset = bitimage_c.AtlasImage_yOffset(images, id)
End Function

Public Property Get psPalette(ByVal i As WZCOL) As Long
    psPalette = piepalette_c.psPalette(i)
End Property

Public Property Let psPalette(ByVal i As WZCOL, ByVal value As Long)
    piepalette_c.psPalette(i) = value
End Property

Public Sub pal_Init()
    piepalette_c.pal_Init
End Sub

Public Function pal_GetTeamColour(ByVal team As Long) As Long
    pal_GetTeamColour = piepalette_c.pal_GetTeamColour(team)
End Function

'--------------------------------------------------------------------------
' cluster D (fork C): piestate, piemode, piematrix, pieclip, culling, piefunc, pielighting, pieblitfunc, screen

Public Function pie_GetResScalingFactor() As Long
    pie_GetResScalingFactor = piemode_c.pie_GetResScalingFactor()
End Function

' rendSurface's fields
Public Property Get rendSurface_width() As Long
    rendSurface_width = piemode_c.rendSurface.width
End Property

Public Property Get rendSurface_height() As Long
    rendSurface_height = piemode_c.rendSurface.height
End Property

Public Property Get rendSurface_xcentre() As Long
    rendSurface_xcentre = piemode_c.rendSurface.xcentre
End Property

Public Property Get rendSurface_ycentre() As Long
    rendSurface_ycentre = piemode_c.rendSurface.ycentre
End Property

' colour: a PIELIGHT (its rgba())
Public Sub pie_SetFogColour(ByVal colour As Long)
    piestate_c.pie_SetFogColour colour
End Sub

Public Function pie_GetFogColour() As Long
    pie_GetFogColour = piestate_c.pie_GetFogColour()
End Function

Public Sub pie_UpdateFogDistance(ByVal begin As Single, ByVal end_ As Single)
    piestate_c.pie_UpdateFogDistance begin, end_
End Sub

' RENDER_STATE getCurrentRenderState(): its fields
Public Sub getCurrentRenderState(fogEnabled As Boolean, fog As Boolean, fogColour As Long, fogBegin As Single, fogEnd As Single, texPage As Long, rendMode As REND_MODE)
    piestate_c.getCurrentRenderState_fields fogEnabled, fog, fogColour, fogBegin, fogEnd, texPage, rendMode
End Sub

Public Function pie_GetMaxAntialiasing() As Long
    pie_GetMaxAntialiasing = piestate_c.pie_GetMaxAntialiasing()
End Function

' const Vector3f& getDefaultSunPosition(): its fields
Public Sub getDefaultSunPosition(x As Single, y As Single, z As Single)
    Dim v As Vector3f
    v = piestate_c.getDefaultSunPosition()
    x = v.x: y = v.y: z = v.z
End Sub

Public Sub pie_SetShaderTime(ByVal shaderTime As Long)
    piestate_c.pie_SetShaderTime shaderTime
End Sub

Public Function pie_GetShaderTime() As Single
    pie_GetShaderTime = piestate_c.pie_GetShaderTime()
End Function

Public Function pie_LoadShaders(ByVal shadowFilterSize As Long, ByVal pointLightEnabled As Boolean) As Boolean
    pie_LoadShaders = piestate_c.pie_LoadShaders(shadowFilterSize, pointLightEnabled)
End Function

Public Sub pie_FreeShaders()
    piestate_c.pie_FreeShaders
End Sub

Public Function pie_getPerspectiveZFar() As Single
    pie_getPerspectiveZFar = piematrix_c.pie_getPerspectiveZFar()
End Function

Public Sub pie_Begin3DScene()
    piematrix_c.pie_Begin3DScene
End Sub

Public Sub pie_BeginInterface()
    piematrix_c.pie_BeginInterface
End Sub

Public Sub pie_SetGeometricOffset(ByVal x As Long, ByVal y As Long)
    piematrix_c.pie_SetGeometricOffset x, y
End Sub

' const glm::mat4& pie_PerspectiveGet(): written to pOut
Public Sub pie_PerspectiveGet(ByVal pOut As Long)
    mat4_toPtr piematrix_c.pie_PerspectiveGet(), pOut
End Sub

Public Sub pie_SkyboxPerspectiveGet(ByVal pOut As Long)
    mat4_toPtr piematrix_c.pie_SkyboxPerspectiveGet(), pOut
End Sub

Public Sub pie_UIPerspectiveGet(ByVal pOut As Long)
    mat4_toPtr piematrix_c.pie_UIPerspectiveGet(), pOut
End Sub

Public Sub pie_PerspectiveGetConstrained(ByVal zNear As Single, ByVal zFar As Single, ByVal pOut As Long)
    mat4_toPtr piematrix_c.pie_PerspectiveGetConstrained(zNear, zFar), pOut
End Sub

' int32_t pie_RotateProjectWithPerspective(const Vector3i *v3d, const glm::mat4 &perspectiveViewMatrix, Vector2i *v2d)
Public Function pie_RotateProjectWithPerspective(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal pPerspectiveViewMatrix As Long, v2dx As Long, v2dy As Long) As Long
    pie_RotateProjectWithPerspective = piematrix_c.pie_RotateProjectWithPerspective_p(x, y, z, pPerspectiveViewMatrix, v2dx, v2dy)
End Function

' pV: the address of 4 Vector3i
Public Sub pie_SetViewingWindow(ByVal pV As Long, ByVal colour As Long)
    piefunc_c.pie_SetViewingWindow_p pV, colour
End Sub

Public Sub pie_DrawViewingWindow(ByVal pModelViewProjectionMatrix As Long)
    piefunc_c.pie_DrawViewingWindow mat4_fromPtr(pModelViewProjectionMatrix)
End Sub

' pVrt: the address of 3 Vector3f
Public Sub pie_TransColouredTriangle(ByVal pVrt As Long, ByVal c As Long, ByVal pModelViewMatrix As Long)
    piefunc_c.pie_TransColouredTriangle_p pVrt, c, pModelViewMatrix
End Sub

Public Sub pie_DrawSkybox(ByVal scale_ As Single, ByVal pProjectionMatrix As Long, ByVal pViewMatrix As Long)
    piefunc_c.pie_DrawSkybox scale_, mat4_fromPtr(pProjectionMatrix), mat4_fromPtr(pViewMatrix)
End Sub

Public Sub pie_Skybox_Texture(ByVal filename As String)
    piefunc_c.pie_Skybox_Texture filename
End Sub

Public Sub iV_ShadowBox(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal pad As Long, ByVal first As Long, ByVal second As Long, ByVal fill As Long)
    pieblitfunc_c.iV_ShadowBox x0, y0, x1, y1, pad, first, second, fill
End Sub

Public Sub iV_Line(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal colour As Long)
    pieblitfunc_c.iV_Line x0, y0, x1, y1, colour
End Sub

' pLines: the address of count glm::ivec4 (x0, y0, x1, y1)
Public Sub iV_Lines(ByVal pLines As Long, ByVal count As Long, ByVal colour As Long)
    pieblitfunc_c.iV_Lines_p pLines, count, colour
End Sub

Public Sub iV_Box2(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal first As Long, ByVal second As Long)
    pieblitfunc_c.iV_Box2 x0, y0, x1, y1, first, second
End Sub

Public Sub iV_Box(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal first As Long)
    pieblitfunc_c.iV_Box x0, y0, x1, y1, first
End Sub

Public Sub pie_BoxFill(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal colour As Long)
    pieblitfunc_c.pie_BoxFill x0, y0, x1, y1, colour
End Sub

Public Sub pie_BoxFillf(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal colour As Long)
    pieblitfunc_c.pie_BoxFillf x0, y0, x1, y1, colour
End Sub

Public Sub pie_BoxFill_alpha(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal colour As Long)
    pieblitfunc_c.pie_BoxFill_alpha x0, y0, x1, y1, colour
End Sub

' pRects: the address of count PIERECT_DrawRequest (x0, y0, x1, y1, color: 5 Longs each)
Public Sub pie_DrawMultiRect(ByVal pRects As Long, ByVal count As Long)
    pieblitfunc_c.pie_DrawMultiRect_p pRects, count
End Sub

Public Sub iV_TransBoxFill(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single)
    pieblitfunc_c.iV_TransBoxFill x0, y0, x1, y1
End Sub

Public Sub pie_UniTransBoxFill(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal colour As Long)
    pieblitfunc_c.pie_UniTransBoxFill x0, y0, x1, y1, colour
End Sub

Public Function assertValidImage(ByVal imageFile_ As Long, ByVal id As Long) As Boolean
    assertValidImage = pieblitfunc_c.assertValidImage(imageFile_, id)
End Function

' angle in radians
Public Sub iV_DrawImageAnisotropic(ByVal TextureID As gfx_api_texture, ByVal PositionX As Long, ByVal PositionY As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal sizeX As Single, ByVal sizeY As Single, ByVal angle As Single, ByVal colour As Long)
    pieblitfunc_c.iV_DrawImageAnisotropic TextureID, glm_ivec2(PositionX, PositionY), glm_vec2(offsetX, offsetY), glm_vec2(sizeX, sizeY), angle, colour
End Sub

Public Sub iV_DrawImage(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, Optional ByVal pModelViewProjection As Long = 0, Optional ByVal pBatchedRequests As BatchedImageDrawRequests, Optional ByVal alpha As Long = 255)
    pieblitfunc_c.iV_DrawImage imageFile_, ID, x, y, defaultProjectionMatrix_or(pModelViewProjection), pBatchedRequests, alpha
End Sub

' size: optional<Vector2f> (has_size, sizeX, sizeY)
Public Sub iV_DrawImageTint(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Single, ByVal y As Single, ByVal color As Long, Optional ByVal has_size As Boolean = False, Optional ByVal sizeX As Single = 0!, Optional ByVal sizeY As Single = 0!, Optional ByVal pModelViewProjection As Long = 0, Optional ByVal pBatchedRequests As BatchedImageDrawRequests)
    pieblitfunc_c.iV_DrawImageTint imageFile_, ID, x, y, color, has_size, glm_vec2(sizeX, sizeY), defaultProjectionMatrix_or(pModelViewProjection), pBatchedRequests
End Sub

Public Sub iV_DrawImageFileAnisotropic(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, ByVal sizeX As Single, ByVal sizeY As Single, Optional ByVal pModelViewProjection As Long = 0, Optional ByVal alpha As Long = 255)
    pieblitfunc_c.iV_DrawImageFileAnisotropic imageFile_, ID, x, y, glm_vec2(sizeX, sizeY), defaultProjectionMatrix_or(pModelViewProjection), alpha
End Sub

Public Sub iV_DrawImageFileAnisotropicTint(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, ByVal sizeX As Single, ByVal sizeY As Single, ByVal color As Long, Optional ByVal pModelViewProjection As Long = 0)
    pieblitfunc_c.iV_DrawImageFileAnisotropicTint imageFile_, ID, x, y, glm_vec2(sizeX, sizeY), color, defaultProjectionMatrix_or(pModelViewProjection)
End Sub

Public Sub iV_DrawImage2(ByVal filename As String, ByVal x As Single, ByVal y As Single, Optional ByVal width As Single = 0!, Optional ByVal height As Single = 0!)
    pieblitfunc_c.iV_DrawImage2 filename, x, y, width, height
End Sub

' image: an AtlasImageDef handle (iV_GetImage)
Public Sub iV_DrawImage2_image(ByVal image As Long, ByVal x As Single, ByVal y As Single, Optional ByVal width As Single = 0!, Optional ByVal height As Single = 0!)
    pieblitfunc_c.iV_DrawImage2_image image, x, y, width, height
End Sub

' iV_DrawImageTc(AtlasImage image, AtlasImage imageTc, int x, int y, PIELIGHT colour, const glm::mat4 &modelViewProjection = defaultProjectionMatrix())
Public Sub iV_DrawImageTc(ByVal image_images As Long, ByVal image_id As Long, ByVal imageTc_images As Long, ByVal imageTc_id As Long, ByVal x As Long, ByVal y As Long, ByVal colour As Long, Optional ByVal pModelViewProjection As Long = 0)
    pieblitfunc_c.iV_DrawImageTc image_images, image_id, imageTc_images, imageTc_id, x, y, colour, defaultProjectionMatrix_or(pModelViewProjection)
End Sub

' the inline iV_DrawImageTc(IMAGEFILE *imageFile, unsigned id, unsigned idTc, int x, int y, PIELIGHT colour)
Public Sub iV_DrawImageTc_ids(ByVal imageFile_ As Long, ByVal id As Long, ByVal idTc As Long, ByVal x As Long, ByVal y As Long, ByVal colour As Long)
    pieblitfunc_c.iV_DrawImageTc_ids imageFile_, id, idTc, x, y, colour
End Sub

' iV_DrawImageImage(AtlasImage image, int x, int y, uint8_t alpha = 255)
Public Sub iV_DrawImageImage(ByVal image_images As Long, ByVal image_id As Long, ByVal x As Long, ByVal y As Long, Optional ByVal alpha As Long = 255)
    pieblitfunc_c.iV_DrawImageImage image_images, image_id, x, y, alpha
End Sub

Public Sub iV_DrawImageRepeatX(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, Optional ByVal pModelViewProjection As Long = 0, Optional ByVal enableHorizontalTilingSeamWorkaround As Boolean = False, Optional ByVal pBatchedRequests As BatchedImageDrawRequests)
    pieblitfunc_c.iV_DrawImageRepeatX imageFile_, ID, x, y, width, defaultProjectionMatrix_or(pModelViewProjection), enableHorizontalTilingSeamWorkaround, pBatchedRequests
End Sub

Public Sub iV_DrawImageRepeatY(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, ByVal height As Long, Optional ByVal pModelViewProjection As Long = 0, Optional ByVal pBatchedRequests As BatchedImageDrawRequests)
    pieblitfunc_c.iV_DrawImageRepeatY imageFile_, ID, x, y, height, defaultProjectionMatrix_or(pModelViewProjection), pBatchedRequests
End Sub

' BatchedImageDrawRequests(bool deferRender = false)
Public Function BatchedImageDrawRequests_new(Optional ByVal deferRender As Boolean = False) As BatchedImageDrawRequests
    Set BatchedImageDrawRequests_new = pieblitfunc_c.BatchedImageDrawRequests_new(deferRender)
End Function

Public Sub pie_SetRadar(ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal twidth As Long, ByVal theight As Long)
    pieblitfunc_c.pie_SetRadar x, y, width, height, twidth, theight
End Sub

Public Sub pie_DownLoadRadar(ByVal bitmap As iV_Image)
    pieblitfunc_c.pie_DownLoadRadar bitmap
End Sub

Public Sub pie_RenderRadar(ByVal pModelViewProjectionMatrix As Long)
    pieblitfunc_c.pie_RenderRadar mat4_fromPtr(pModelViewProjectionMatrix)
End Sub

Public Sub pie_LoadBackDrop(ByVal screenType_ As SCREENTYPE)
    pieblitfunc_c.pie_LoadBackDrop screenType_
End Sub

' glm::mat4 defaultProjectionMatrix(): written to pOut
Public Sub defaultProjectionMatrix(ByVal pOut As Long)
    mat4_toPtr pieblitfunc_c.defaultProjectionMatrix(), pOut
End Sub

Public Function screenInitialise() As Boolean
    screenInitialise = screen_c.screenInitialise()
End Function

Public Sub screen_SetRandomBackdrop(ByVal dirname As String, ByVal basename As String)
    screen_c.screen_SetRandomBackdrop dirname, basename
End Sub

Public Sub screen_SetBackDropFromFile(ByVal filename As String)
    screen_c.screen_SetBackDropFromFile filename
End Sub

Public Sub screen_StopBackDrop()
    screen_c.screen_StopBackDrop
End Sub

Public Function screen_RestartBackDrop() As Boolean
    screen_RestartBackDrop = screen_c.screen_RestartBackDrop()
End Function

Public Function screen_GetBackDrop() As Boolean
    screen_GetBackDrop = screen_c.screen_GetBackDrop()
End Function

Public Sub screen_FlipIfBackDropTransition()
    screen_c.screen_FlipIfBackDropTransition
End Sub

' newBackdropImage: iV_Image&& (drop it after the call)
Public Sub screen_Upload(ByVal newBackdropImage As iV_Image)
    screen_c.screen_Upload newBackdropImage
End Sub

Public Sub screen_Display()
    screen_c.screen_Display
End Sub

Public Sub screenDumpToDisk(ByVal path As String, ByVal level As String)
    screen_c.screenDumpToDisk path, level
End Sub

' pPlayerpositions: the address of MAX_PLAYERS Vector2i
Public Sub screen_enableMapPreview(ByVal width As Long, ByVal height As Long, ByVal pPlayerpositions As Long)
    screen_c.screen_enableMapPreview_p width, height, pPlayerpositions
End Sub

Public Sub screen_disableMapPreview()
    screen_c.screen_disableMapPreview
End Sub

Public Sub wzPerfBegin(ByVal pp As PERF_POINT, ByVal descr As String)
    screen_c.wzPerfBegin pp, descr
End Sub

Public Sub wzPerfEnd(ByVal pp As PERF_POINT)
    screen_c.wzPerfEnd pp
End Sub

Public Sub wzPerfStart()
    screen_c.wzPerfStart
End Sub

Public Sub wzPerfShutdown()
    screen_c.wzPerfShutdown
End Sub

Public Function wzPerfAvailable() As Boolean
    wzPerfAvailable = screen_c.wzPerfAvailable()
End Function

Public Sub setLightingManager(ByVal manager As ILightingManager)
    pielighting_c.setLightingManager manager
End Sub

Public Function getCurrentLightingManager() As ILightingManager
    Set getCurrentLightingManager = pielighting_c.getCurrentLightingManager()
End Function

' std::make_unique<renderingNew::LightingManager>()
Public Function renderingNew_LightingManager_new() As ILightingManager
    Set renderingNew_LightingManager_new = pielighting_c.renderingNew_LightingManager_new()
End Function

Public Sub ILightingManager_SetFrameStart(ByVal manager As ILightingManager)
    pielighting_c.ILightingManager_SetFrameStart manager
End Sub

' getPointLightBuckets(): its arrays copied to the four addresses; returns bucketDimensionUsed
Public Function ILightingManager_getPointLightBuckets(ByVal manager As ILightingManager, ByVal pPositions As Long, ByVal pColorAndEnergy As Long, ByVal pBucketOffsetAndSize As Long, ByVal pLightIndex As Long) As Long
    ILightingManager_getPointLightBuckets = pielighting_c.ILightingManager_getPointLightBuckets_p(manager, pPositions, pColorAndEnergy, pBucketOffsetAndSize, pLightIndex)
End Function

' getCurrentLightingData().lights
Public Sub LightingData_lights_clear()
    pielighting_c.LightingData_lights_clear
End Sub

Public Sub LightingData_lights_push_back(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal range As Long, ByVal colour As Long)
    pielighting_c.LightingData_lights_push_back_fields x, y, z, range, colour
End Sub

Public Function LightingData_lights_size() As Long
    LightingData_lights_size = pielighting_c.LightingData_lights_size()
End Function

Public Sub LightingData_lights_at(ByVal i As Long, x As Long, y As Long, z As Long, range As Long, colour As Long)
    pielighting_c.LightingData_lights_at_fields i, x, y, z, range, colour
End Sub

' getCurrentLightmapData()(x, y) (the const operator(): LightMap_get_const), .reset(w, h), the data (row after row)
Public Function LightMap_get(ByVal x As Long, ByVal y As Long) As Long
    LightMap_get = pielighting_c.LightMap_get(x, y)
End Function

Public Function LightMap_get_const(ByVal x As Long, ByVal y As Long) As Long
    LightMap_get_const = pielighting_c.LightMap_get_const(x, y)
End Function

Public Sub LightMap_set(ByVal x As Long, ByVal y As Long, ByVal colour As Long)
    pielighting_c.LightMap_set x, y, colour
End Sub

Public Sub LightMap_reset(ByVal width As Long, ByVal height As Long)
    pielighting_c.LightMap_reset width, height
End Sub

Public Function LightMap_dataAddress() As Long
    LightMap_dataAddress = pielighting_c.LightMap_dataAddress()
End Function

'--------------------------------------------------------------------------
' cluster D (fork T): textdraw.cpp (on GDI)

Public Sub iV_TextUpdateScaleFactor(ByVal horizScalePercentage As Long, ByVal vertScalePercentage As Long)
    textdraw_gdi.iV_TextUpdateScaleFactor horizScalePercentage, vertScalePercentage
End Sub

Public Sub iV_TextShutdown()
    textdraw_gdi.iV_TextShutdown
End Sub

Public Sub iV_font(ByVal fontName As String, ByVal fontFace As String, ByVal fontFaceBold As String)
    textdraw_gdi.iV_font fontName, fontFace, fontFaceBold
End Sub

Public Function iV_GetEllipsisWidth(ByVal fontID As iV_fonts) As Long
    iV_GetEllipsisWidth = textdraw_gdi.iV_GetEllipsisWidth(fontID)
End Function

' position: Vector2f (x, y)
Public Sub iV_DrawEllipsis(ByVal fontID As iV_fonts, ByVal x As Single, ByVal y As Single, ByVal colour As Long, Optional ByVal rotation_ As Single = 0!)
    textdraw_gdi.iV_DrawEllipsis fontID, x, y, colour, rotation_
End Sub

Public Function iV_GetTextAboveBase(ByVal fontID As iV_fonts) As Long
    iV_GetTextAboveBase = textdraw_gdi.iV_GetTextAboveBase(fontID)
End Function

Public Function iV_GetTextBelowBase(ByVal fontID As iV_fonts) As Long
    iV_GetTextBelowBase = textdraw_gdi.iV_GetTextBelowBase(fontID)
End Function

Public Function iV_GetTextLineSize(ByVal fontID As iV_fonts) As Long
    iV_GetTextLineSize = textdraw_gdi.iV_GetTextLineSize(fontID)
End Function

Public Function iV_GetTextWidth(ByVal string_ As String, ByVal fontID As iV_fonts) As Long
    iV_GetTextWidth = textdraw_gdi.iV_GetTextWidth(string_, fontID)
End Function

Public Function iV_GetCountedTextWidth(ByVal string_ As String, ByVal string_length As Long, ByVal fontID As iV_fonts) As Long
    iV_GetCountedTextWidth = textdraw_gdi.iV_GetCountedTextWidth(string_, string_length, fontID)
End Function

Public Function iV_GetCharWidth(ByVal charCode As Long, ByVal fontID As iV_fonts) As Long
    iV_GetCharWidth = textdraw_gdi.iV_GetCharWidth(charCode, fontID)
End Function

Public Function iV_GetTextHeight(ByVal string_ As String, ByVal fontID As iV_fonts) As Long
    iV_GetTextHeight = textdraw_gdi.iV_GetTextHeight(string_, fontID)
End Function

Public Sub iV_SetTextColour(ByVal colour As Long)
    textdraw_gdi.iV_SetTextColour colour
End Sub

' optional<iV_fonts>: returns has_value, the font in result
Public Function iV_FontModifyBold(ByVal fontID As iV_fonts, ByVal bold As Boolean, result As iV_fonts) As Boolean
    iV_FontModifyBold = textdraw_gdi.iV_FontModifyBold(fontID, bold, result)
End Function

Public Function iV_ShrinkFont(ByVal fontID As iV_fonts, result As iV_fonts) As Boolean
    iV_ShrinkFont = textdraw_gdi.iV_ShrinkFont(fontID, result)
End Function

' std::vector<WzTextRun>: the runs' fields in arrays (0-based); returns their number
Public Function iV_SplitTextParagraphIntoRuns(ByVal string_ As String, ByVal fontID As iV_fonts, startOffsets() As Long, endOffsets() As Long, rightToLeft() As Boolean) As Long
    iV_SplitTextParagraphIntoRuns = textdraw_gdi.iV_SplitTextParagraphIntoRuns_arrays(string_, fontID, startOffsets, endOffsets, rightToLeft)
End Function

Public Function iV_GetMaxTextRunLenForWidth(ByVal textRun_ As String, ByVal fontID As iV_fonts, ByVal maxWidthInPoints As Long, ByVal rightToLeft As Boolean) As Long
    iV_GetMaxTextRunLenForWidth = textdraw_gdi.iV_GetMaxTextRunLenForWidth(textRun_, fontID, maxWidthInPoints, rightToLeft)
End Function

' std::vector<TextLine>: the lines' text, dimensions (x, y) and offset (x, y) in arrays (0-based); returns their number
Public Function iV_FormatText(ByVal String_ As String, ByVal MaxWidth As Long, ByVal Justify As FTEXT_JUSTIFY, ByVal fontID As iV_fonts, ByVal ignoreNewlines As Boolean, lineText() As String, lineWidth() As Long, lineHeight() As Long, lineOffsetX() As Long, lineOffsetY() As Long) As Long
    iV_FormatText = textdraw_gdi.iV_FormatText_arrays(String_, MaxWidth, Justify, fontID, ignoreNewlines, lineText, lineWidth, lineHeight, lineOffsetX, lineOffsetY)
End Function

Public Sub iV_DrawTextRotated(ByVal string_ As String, ByVal XPos As Single, ByVal YPos As Single, ByVal rotation_ As Single, ByVal fontID As iV_fonts)
    textdraw_gdi.iV_DrawTextRotated string_, XPos, YPos, rotation_, fontID
End Sub

Public Sub iV_DrawText(ByVal string_ As String, ByVal x As Single, ByVal y As Single, ByVal fontID As iV_fonts)
    textdraw_gdi.iV_DrawText string_, x, y, fontID
End Sub

' WzText(const WzString &text, iV_fonts fontID)
Public Function WzText_new(ByVal text As String, ByVal fontID As iV_fonts) As WzText
    Set WzText_new = textdraw_gdi.WzText_new(text, fontID)
End Function

'--------------------------------------------------------------------------
' cluster E (fork P): piedraw.cpp, imdload.cpp's display shapes

' bool pie_Draw3DShape(const iIMDShape *shape, int frame, int team, PIELIGHT colour, int pieFlag, int pieFlagData,
'   const glm::mat4 &modelMatrix, const glm::mat4 &viewMatrix, float stretchDepth = 0.f, bool onlySingleLevel = false)
' shape: an iIMDShape handle; pModelMatrix, pViewMatrix: the mat4s' addresses
Public Function pie_Draw3DShape(ByVal shape As Long, ByVal frame As Long, ByVal team As Long, ByVal colour As Long, ByVal pieFlag As Long, ByVal pieFlagData As Long, ByVal pModelMatrix As Long, ByVal pViewMatrix As Long, Optional ByVal stretchDepth As Single = 0!, Optional ByVal onlySingleLevel As Boolean = False) As Boolean
    Dim modelMatrix As mat4, viewMatrix As mat4
    modelMatrix = mat4_fromPtr(pModelMatrix)
    viewMatrix = mat4_fromPtr(pViewMatrix)
    pie_Draw3DShape = piedraw_c.pie_Draw3DShape(shape, frame, team, colour, pieFlag, pieFlagData, modelMatrix, viewMatrix, stretchDepth, onlySingleLevel)
End Function

' void pie_Draw3DButton(const iIMDShape *shape, PIELIGHT teamcolour, const glm::mat4 &modelMatrix, const glm::mat4 &viewMatrix)
Public Sub pie_Draw3DButton(ByVal shape As Long, ByVal teamcolour As Long, ByVal pModelMatrix As Long, ByVal pViewMatrix As Long)
    Dim modelMatrix As mat4, viewMatrix As mat4
    modelMatrix = mat4_fromPtr(pModelMatrix)
    viewMatrix = mat4_fromPtr(pViewMatrix)
    piedraw_c.pie_Draw3DButton shape, teamcolour, modelMatrix, viewMatrix
End Sub

' void pie_GetResetCounts(size_t *pPieCount, size_t *pPolyCount)
Public Sub pie_GetResetCounts(pPieCount As Long, pPolyCount As Long)
    piedraw_c.pie_GetResetCounts pPieCount, pPolyCount
End Sub

' void pie_BeginLighting(const Vector3f &light): pLight, its address
Public Sub pie_BeginLighting(ByVal pLight As Long)
    Dim light_ As Vector3f
    piedraw_c.piedraw_vec3_fromPtr pLight, light_
    piedraw_c.pie_BeginLighting light_
End Sub

Public Sub pie_setShadows(ByVal drawShadows As Boolean)
    piedraw_c.pie_setShadows drawShadows
End Sub

Public Sub pie_InitLighting()
    piedraw_c.pie_InitLighting
End Sub

' void pie_Lighting0(LIGHTING_TYPE entry, const float value[4]): value(0 To 3)
Public Sub pie_Lighting0(ByVal entry As LIGHTING_TYPE, value() As Single)
    piedraw_c.pie_Lighting0 entry, value
End Sub

' glm::vec4 pie_GetLighting0(LIGHTING_TYPE entry): written to pOut (the address of a Vector4f)
Public Sub pie_GetLighting0(ByVal entry As LIGHTING_TYPE, ByVal pOut As Long)
    Dim v As Vector4f
    v = piedraw_c.pie_GetLighting0(entry)
    piedraw_c.piedraw_vec4_toPtr v, pOut
End Sub

Public Function pie_setShadowMode(ByVal mode As ShadowMode) As Boolean
    pie_setShadowMode = piedraw_c.pie_setShadowMode(mode)
End Function

Public Function pie_getShadowMode() As ShadowMode
    pie_getShadowMode = piedraw_c.pie_getShadowMode()
End Function

Public Function pie_setShadowCascades(ByVal newValue As Long) As Boolean
    pie_setShadowCascades = piedraw_c.pie_setShadowCascades(newValue)
End Function

Public Function pie_getShadowCascades() As Long
    pie_getShadowCascades = piedraw_c.pie_getShadowCascades()
End Function

' optional<bool> pie_supportsShadowMapping(): has_value returned, the value in value
Public Function pie_supportsShadowMapping(value As Boolean) As Boolean
    pie_supportsShadowMapping = piedraw_c.pie_supportsShadowMapping(value)
End Function

Public Function pie_setShadowMapResolution(ByVal resolution As Long) As Boolean
    pie_setShadowMapResolution = piedraw_c.pie_setShadowMapResolution(resolution)
End Function

Public Function pie_getShadowMapResolution() As Long
    pie_getShadowMapResolution = piedraw_c.pie_getShadowMapResolution()
End Function

Public Sub pie_StartMeshes()
    piedraw_c.pie_StartMeshes
End Sub

' void pie_UpdateLightmap(gfx_api::texture* lightmapTexture, const glm::mat4& modelUVLightmapMatrix)
Public Sub pie_UpdateLightmap(ByVal lightmapTexture As gfx_api_texture, ByVal pModelUVLightmapMatrix As Long)
    Dim modelUVLightmapMatrix As mat4
    modelUVLightmapMatrix = mat4_fromPtr(pModelUVLightmapMatrix)
    piedraw_c.pie_UpdateLightmap lightmapTexture, modelUVLightmapMatrix
End Sub

' void pie_FinalizeMeshes(uint64_t currentGameFrame)
Public Sub pie_FinalizeMeshes(ByVal currentGameFrame As Double)
    piedraw_c.pie_FinalizeMeshes currentGameFrame
End Sub

' void pie_DrawAllMeshes(uint64_t currentGameFrame, const glm::mat4 &projectionMatrix, const glm::mat4 &viewMatrix,
'   const Vector3f &cameraPos, const ShadowCascadesInfo& shadowMVPMatrix, bool depthPass)
' pCameraPos: a Vector3f's address; pShadowCascadesInfo: the address of a src/shared/gfx/shadows_h ShadowCascadesInfo
Public Sub pie_DrawAllMeshes(ByVal currentGameFrame As Double, ByVal pProjectionMatrix As Long, ByVal pViewMatrix As Long, ByVal pCameraPos As Long, ByVal pShadowCascadesInfo As Long, ByVal depthPass As Boolean)
    Dim projectionMatrix As mat4, viewMatrix As mat4, cameraPos As Vector3f, shadowMVPMatrix As ShadowCascadesInfo
    projectionMatrix = mat4_fromPtr(pProjectionMatrix)
    viewMatrix = mat4_fromPtr(pViewMatrix)
    piedraw_c.piedraw_vec3_fromPtr pCameraPos, cameraPos
    shadowCascadesInfo_fromPtr pShadowCascadesInfo, shadowMVPMatrix
    piedraw_c.pie_DrawAllMeshes currentGameFrame, projectionMatrix, viewMatrix, cameraPos, shadowMVPMatrix, depthPass
End Sub

' PIELIGHT straightAlphaToPremultiplied(PIELIGHT colour)
Public Function straightAlphaToPremultiplied(ByVal colour As Long) As Long
    straightAlphaToPremultiplied = piedraw_c.straightAlphaToPremultiplied(colour)
End Function

Public Sub iIMDShape_min(ByVal shape As Long, x As Long, y As Long, z As Long)
    imdload_c.iIMDShape_min shape, x, y, z
End Sub

Public Sub iIMDShape_max(ByVal shape As Long, x As Long, y As Long, z As Long)
    imdload_c.iIMDShape_max shape, x, y, z
End Sub

Public Function iIMDShape_sradius(ByVal shape As Long) As Long
    iIMDShape_sradius = imdload_c.iIMDShape_sradius(shape)
End Function

Public Function iIMDShape_radius(ByVal shape As Long) As Long
    iIMDShape_radius = imdload_c.iIMDShape_radius(shape)
End Function

Public Sub iIMDShape_ocen(ByVal shape As Long, x As Single, y As Single, z As Single)
    imdload_c.iIMDShape_ocen shape, x, y, z
End Sub

Public Function iIMDShape_connectors_size(ByVal shape As Long) As Long
    iIMDShape_connectors_size = imdload_c.iIMDShape_connectors_size(shape)
End Function

Public Sub iIMDShape_connector(ByVal shape As Long, ByVal i As Long, x As Long, y As Long, z As Long)
    imdload_c.iIMDShape_connector shape, i, x, y, z
End Sub

Public Function iIMDShape_flags(ByVal shape As Long) As Long
    iIMDShape_flags = imdload_c.iIMDShape_flags(shape)
End Function

Public Function iIMDShape_numFrames(ByVal shape As Long) As Long
    iIMDShape_numFrames = imdload_c.iIMDShape_numFrames(shape)
End Function

Public Function iIMDShape_animInterval(ByVal shape As Long) As Long
    iIMDShape_animInterval = imdload_c.iIMDShape_animInterval(shape)
End Function

Public Function iIMDShape_interpolate(ByVal shape As Long) As Long
    iIMDShape_interpolate = imdload_c.iIMDShape_interpolate(shape)
End Function

Public Function iIMDShape_modelLevel(ByVal shape As Long) As Long
    iIMDShape_modelLevel = imdload_c.iIMDShape_modelLevel(shape)
End Function

Public Function iIMDShape_polys_size(ByVal shape As Long) As Long
    iIMDShape_polys_size = imdload_c.iIMDShape_polys_size(shape)
End Function

' objanimdata[i]: scale, pos, rot
' objanimdata[i] copied to the ANIMFRAME at pOut (36 bytes: scale, pos, rot)
Public Sub iIMDShape_objanimdata_p(ByVal shape As Long, ByVal i As Long, ByVal pOut As Long)
    imdload_c.iIMDShape_objanimdata_p shape, i, pOut
End Sub

Public Function iIMDShape_texpage(ByVal shape As Long) As Long
    iIMDShape_texpage = imdload_c.iIMDShape_texpage(shape)
End Function

Public Sub iIMDShape_objanimdata(ByVal shape As Long, ByVal i As Long, scaleX As Single, scaleY As Single, scaleZ As Single, posX As Single, posY As Single, posZ As Single, rotDirection As Long, rotPitch As Long, rotRoll As Long)
    imdload_c.iIMDShape_objanimdata shape, i, scaleX, scaleY, scaleZ, posX, posY, posZ, rotDirection, rotPitch, rotRoll
End Sub

' PIELIGHT iIMDShape::getTeamColourForModel(int team) const
Public Function iIMDShape_getTeamColourForModel(ByVal shape As Long, ByVal team As Long) As Long
    iIMDShape_getTeamColourForModel = imdload_c.iIMDShape_getTeamColourForModel(shape, team)
End Function

' imd.h: void modelReloadAllModelTextures()
Public Sub modelReloadAllModelTextures()
    imdload_c.modelReloadAllModelTextures
End Sub
