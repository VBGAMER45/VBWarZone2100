Attribute VB_Name = "graphicsoptions_c"
' src/titleui/options/graphicsoptions.cpp: the Graphics Options form (makeGraphicsOptionsForm) and
' its availability conditions, whole. (lib/ivis_opengl/gfx_api.h's static inline to_swap_mode: a
' copy, as main_c's.)
Option Explicit

Public Const FN_TerrainShadingQualityAvailable As Long = 4300 ' OptionInfo::AvailabilityResult TerrainShadingQualityAvailable(const OptionInfo&)
Public Const FN_ShadowsEnabled As Long = 4301               ' OptionInfo::AvailabilityResult ShadowsEnabled(const OptionInfo&)
Public Const FN_SupportsShadowMapping As Long = 4302        ' OptionInfo::AvailabilityResult SupportsShadowMapping(const OptionInfo&)
Public Const FN_PerPixelLightingAvailable As Long = 4303    ' OptionInfo::AvailabilityResult PerPixelLightingAvailable(const OptionInfo&)
Public Const FN_gfx_terrainAppearance_populate As Long = 4304 ' makeGraphicsOptionsForm, "gfx.terrainAppearance"
Public Const FN_gfx_terrainAppearance_select As Long = 4305
Public Const FN_gfx_terrainShadingQuality_populate As Long = 4306 ' "gfx.terrainShadingQuality"
Public Const FN_gfx_terrainShadingQuality_select As Long = 4307
Public Const FN_gfx_drawShadows_populate As Long = 4308     ' "gfx.drawShadows"
Public Const FN_gfx_drawShadows_select As Long = 4309
Public Const FN_gfx_drawShadowsTerrain_populate As Long = 4310 ' "gfx.drawShadowsTerrain"
Public Const FN_gfx_drawShadowsTerrain_select As Long = 4311
Public Const FN_gfx_shadowResolution_populate As Long = 4312 ' "gfx.shadowResolution"
Public Const FN_gfx_shadowResolution_select As Long = 4313
Public Const FN_gfx_shadowFiltering_populate As Long = 4314 ' "gfx.shadowFiltering"
Public Const FN_gfx_shadowFiltering_select As Long = 4315
Public Const FN_gfx_pointLights_populate As Long = 4316     ' "gfx.pointLights"
Public Const FN_gfx_pointLights_select As Long = 4317
Public Const FN_gfx_fog_populate As Long = 4318             ' "gfx.fog"
Public Const FN_gfx_fog_select As Long = 4319
Public Const FN_gfx_screenShake_populate As Long = 4320     ' "gfx.screenShake"
Public Const FN_gfx_screenShake_select As Long = 4321
Public Const FN_gfx_textureSizeLimit_populate As Long = 4322 ' "gfx.textureSizeLimit"
Public Const FN_gfx_textureSizeLimit_select As Long = 4323
Public Const FN_gfx_lodDistance_populate As Long = 4324     ' "gfx.lodDistance"
Public Const FN_gfx_lodDistance_select As Long = 4325
Public Const FN_gfx_backend_populate As Long = 4326         ' "gfx.backend"
Public Const FN_gfx_backend_select As Long = 4327
Public Const FN_gfx_antialiasing_populate As Long = 4328    ' "gfx.antialiasing"
Public Const FN_gfx_antialiasing_select As Long = 4329
Public Const FN_gfx_vsync_populate As Long = 4330           ' "gfx.vsync"
Public Const FN_gfx_vsync_select As Long = 4331

' lib/ivis_opengl/gfx_api.h: static inline gfx_api::context::swap_interval_mode to_swap_mode(int value)
Private Function to_swap_mode(ByVal value As Long) As swap_interval_mode
    Select Case value
    Case swap_interval_mode_immediate
        to_swap_mode = swap_interval_mode_immediate
        Exit Function
    Case swap_interval_mode_vsync
        to_swap_mode = swap_interval_mode_vsync
        Exit Function
    Case swap_interval_mode_adaptive_vsync
        to_swap_mode = swap_interval_mode_adaptive_vsync
        Exit Function
    Case Else
        debug_ 1352, LOG_WARNING, "to_swap_mode", "Invalid vsync value (%d); defaulting to vsync ON", value
    End Select
    to_swap_mode = swap_interval_mode_vsync
End Function

' (VB) pie_supportsShadowMapping().value_or(false)
Private Function supportsShadowMappingOrFalse() As Boolean
    Dim v As Boolean
    If Not pie_supportsShadowMapping(v) Then v = False
    supportsShadowMappingOrFalse = v
End Function

' MARK: -

Public Function TerrainShadingQualityAvailable(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = (getTerrainShaderQuality() = NORMAL_MAPPING)
    result.localizedUnavailabilityReason = "Terrain Shading Quality only applies when using terrain appearance: Remastered (HQ)"
    Set TerrainShadingQualityAvailable = result
End Function

Public Function ShadowsEnabled(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = getDrawShadows()
    result.localizedUnavailabilityReason = "Shadows are not enabled."
    Set ShadowsEnabled = result
End Function

Public Function SupportsShadowMapping(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult, bShadowMappingSupported As Boolean
    Set result = New OptionInfo_AvailabilityResult
    bShadowMappingSupported = supportsShadowMappingOrFalse()
    result.available_ = bShadowMappingSupported
    result.localizedUnavailabilityReason = "Shadow mapping not available on this system."
    Set SupportsShadowMapping = result
End Function

Public Function PerPixelLightingAvailable(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = (getTerrainShaderQuality() = NORMAL_MAPPING)
    result.localizedUnavailabilityReason = "Advanced lighting is only available when using terrain appearance: Remastered (HQ)"
    Set PerPixelLightingAvailable = result
End Function

' MARK: -

Public Function makeGraphicsOptionsForm() As WIDGET
    Dim result As WIDGET, optionInfo As OptionInfo, valueChanger As WIDGET, lodDistanceString As String
    Set result = OptionsForm_make()

    ' Terrain:
    OptionsForm_addSection result, OptionsSection_("Terrain", ""), True
    Set optionInfo = OptionInfo_new("gfx.terrainAppearance", "Terrain Appearance", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_terrainAppearance_populate), WzFunc_new(FN_gfx_terrainAppearance_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("gfx.terrainShadingQuality", "Terrain Shading Quality", "The quality of terrain shading (normal maps, specular highlights, etc). Higher settings improve terrain details at the expense of performance and memory usage.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    optionInfo.addAvailabilityCondition WzFunc_new(FN_TerrainShadingQualityAvailable)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_terrainShadingQuality_populate), WzFunc_new(FN_gfx_terrainShadingQuality_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Shadows:
    OptionsForm_addSection result, OptionsSection_("Shadows", ""), True
    Set optionInfo = OptionInfo_new("gfx.drawShadows", "Draw Shadows", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_drawShadows_populate), WzFunc_new(FN_gfx_drawShadows_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("gfx.drawShadowsTerrain", "Terrain Shadows", "Whether terrain casts shadows.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_ShadowsEnabled)
    optionInfo.addAvailabilityCondition WzFunc_new(FN_SupportsShadowMapping)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_drawShadowsTerrain_populate), WzFunc_new(FN_gfx_drawShadowsTerrain_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True, 1

    Set optionInfo = OptionInfo_new("gfx.shadowResolution", "Shadow Resolution", "The internal resolution used to render shadow maps. Higher values improve shadow appearance at the expense of performance and memory usage.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_ShadowsEnabled)
    optionInfo.addAvailabilityCondition WzFunc_new(FN_SupportsShadowMapping)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_shadowResolution_populate), WzFunc_new(FN_gfx_shadowResolution_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("gfx.shadowFiltering", "Shadow Filtering", "How much filtering / smoothing is applied to shadow appearance. Higher settings smooth shadows more at the expense of performance.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_ShadowsEnabled)
    optionInfo.addAvailabilityCondition WzFunc_new(FN_SupportsShadowMapping)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_shadowFiltering_populate), WzFunc_new(FN_gfx_shadowFiltering_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Lighting:
    OptionsForm_addSection result, OptionsSection_("Lighting", ""), True
    Set optionInfo = OptionInfo_new("gfx.pointLights", "Point Lights", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    optionInfo.addAvailabilityCondition WzFunc_new(FN_PerPixelLightingAvailable)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_pointLights_populate), WzFunc_new(FN_gfx_pointLights_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Effects:
    OptionsForm_addSection result, OptionsSection_("Effects", ""), True
    Set optionInfo = OptionInfo_new("gfx.fog", "Fog", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_fog_populate), WzFunc_new(FN_gfx_fog_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("gfx.screenShake", "Screen Shake", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_screenShake_populate), WzFunc_new(FN_gfx_screenShake_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Textures:
    OptionsForm_addSection result, OptionsSection_("Textures", ""), True
    Set optionInfo = OptionInfo_new("gfx.textureSizeLimit", "Texture Size Limit", "Textures larger than the configured limit will be downsampled. Unless running on a system with limited available memory, this should usually be set to the maximum value.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_textureSizeLimit_populate), WzFunc_new(FN_gfx_textureSizeLimit_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' TRANSLATORS: "LOD" = "Level of Detail" - this setting is used to describe how level of detail (in textures) is preserved as distance increases (examples: "Default", "High", etc)
    lodDistanceString = "LOD Distance"
    Set optionInfo = OptionInfo_new("gfx.lodDistance", lodDistanceString, "How level of detail (in textures) is preserved as distance increases.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    optionInfo.setRequiresRestart True
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_lodDistance_populate), WzFunc_new(FN_gfx_lodDistance_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Advanced:
    OptionsForm_addSection result, OptionsSection_("Advanced", ""), True
    Set optionInfo = OptionInfo_new("gfx.backend", "Graphics Backend", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    optionInfo.setRequiresRestart True
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_backend_populate), WzFunc_new(FN_gfx_backend_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("gfx.antialiasing", "Antialiasing", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    optionInfo.setRequiresRestart True
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_antialiasing_populate), WzFunc_new(FN_gfx_antialiasing_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("gfx.vsync", "Vertical Sync", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_gfx_vsync_populate), WzFunc_new(FN_gfx_vsync_select), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set makeGraphicsOptionsForm = result
End Function

' (VB) OptionChoices' "Custom" item: result.choices.push_back({WzString::format("(Custom: %" PRIu32 ")", currValue), "", currValue}); result.currentIdx = result.choices.size() - 1;
Private Sub addCustomChoice(ByVal result As OptionChoices, ByVal currValue As Long)
    result.addChoice "(Custom: " & CStr(currValue) & ")", "", currValue
    result.currentIdx = result.count - 1
End Sub

' MARK: - the lambdas

Public Function graphicsoptions_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim result As OptionChoices, currentValue As Long, currValue As Long, newValue As Long, i As Long
    Dim shadowMode_ As Long, shadowFilterSize As Long, shadowCascadesCount As Long, isPointLightPerPixelEnabled As Boolean
    Dim availableBackends() As Long, availableBackendsCount As Long, maxAA As Long, value As Long, canChange As Boolean
    Select Case f.id
    Case FN_TerrainShadingQualityAvailable
        Set graphicsoptions_c_fn = TerrainShadingQualityAvailable(args(0))
    Case FN_ShadowsEnabled
        Set graphicsoptions_c_fn = ShadowsEnabled(args(0))
    Case FN_SupportsShadowMapping
        Set graphicsoptions_c_fn = SupportsShadowMapping(args(0))
    Case FN_PerPixelLightingAvailable
        Set graphicsoptions_c_fn = PerPixelLightingAvailable(args(0))

    Case FN_gfx_terrainAppearance_populate      ' OptionChoices<TerrainShaderQuality> ()
        Set result = New OptionChoices
        result.addChoice to_display_string(CLASSIC), "The classic terrain appearance from the initial release. Low-resolution, pixel-art textures.", CLASSIC, Not isSupportedTerrainShaderQualityOption(CLASSIC)
        result.addChoice to_display_string(MEDIUM), "The normal terrain appearance used as the default for more than a decade. Higher-resolution, recreated textures.", MEDIUM, Not isSupportedTerrainShaderQualityOption(MEDIUM)
        result.addChoice to_display_string(NORMAL_MAPPING), "The remastered, high-quality terrain appearance. High-resolution, with modern bump + specular mapping. Requires more memory and a better GPU.", NORMAL_MAPPING, Not isSupportedTerrainShaderQualityOption(NORMAL_MAPPING)
        currentValue = getTerrainShaderQuality()
        If currentValue >= 0 Then result.currentIdx = currentValue Else result.currentIdx = 1
        Set graphicsoptions_c_fn = result

    Case FN_gfx_terrainAppearance_select        ' [](const auto& newValue) -> bool
        newValue = CLng(args(0))
        If Not isSupportedTerrainShaderQualityOption(newValue) Then
            graphicsoptions_c_fn = False
            Exit Function
        End If
        If Not setTerrainShaderQuality(newValue) Then
            debug_ 98, LOG_ERROR, "makeGraphicsOptionsForm", "Failed to set terrain shader quality: %s", to_display_string(newValue)
            graphicsoptions_c_fn = False
            Exit Function
        End If
        graphicsoptions_c_fn = True

    Case FN_gfx_terrainShadingQuality_populate  ' OptionChoices<int32_t> ()
        Set result = New OptionChoices
        result.addChoice "Medium Quality", "", 512
        result.addChoice "High Quality", "", 1024
        result.setCurrentIdxForValue getTerrainMappingTexturesMaxSize()
        Set graphicsoptions_c_fn = result

    Case FN_gfx_terrainShadingQuality_select
        newValue = CLng(args(0))
        If Not setTerrainMappingTexturesMaxSize(newValue) Then
            debug_ 123, LOG_ERROR, "makeGraphicsOptionsForm", "Failed to set terrain mapping texture quality: %d", newValue
            graphicsoptions_c_fn = False
            Exit Function
        End If
        graphicsoptions_c_fn = True

    Case FN_gfx_drawShadows_populate            ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        If getDrawShadows() Then result.currentIdx = 1 Else result.currentIdx = 0
        Set graphicsoptions_c_fn = result

    Case FN_gfx_drawShadows_select
        setDrawShadows CBool(args(0))
        graphicsoptions_c_fn = True

    Case FN_gfx_drawShadowsTerrain_populate     ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue getDrawTerrainShadows()
        Set graphicsoptions_c_fn = result

    Case FN_gfx_drawShadowsTerrain_select
        setDrawTerrainShadows CBool(args(0))
        graphicsoptions_c_fn = True

    Case FN_gfx_shadowResolution_populate       ' OptionChoices<uint32_t> ()
        Set result = New OptionChoices
        result.addChoice "Normal" & " (2048)", "", 2048
        result.addChoice "High" & " (4096)", "", 4096
        currValue = pie_getShadowMapResolution()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add "Custom" item
            addCustomChoice result, currValue
        End If
        Set graphicsoptions_c_fn = result

    Case FN_gfx_shadowResolution_select         ' [](const auto& newResolution) -> bool
        newValue = CLng(args(0))
        If Not supportsShadowMappingOrFalse() Then
            graphicsoptions_c_fn = False
            Exit Function
        End If
        If Not pie_setShadowMapResolution(newValue) Then
            debug_ 201, LOG_ERROR, "makeGraphicsOptionsForm", "Failed to set map resolution: %u", newValue
            graphicsoptions_c_fn = False
            Exit Function
        End If
        war_setShadowMapResolution newValue
        graphicsoptions_c_fn = True

    Case FN_gfx_shadowFiltering_populate        ' OptionChoices<uint32_t> ()
        Set result = New OptionChoices
        result.addChoice "Low", "", 3
        result.addChoice "High", "", 5
        result.addChoice "Ultra", "", 7
        gfx_api_context_getShadowConstants shadowMode_, shadowFilterSize, shadowCascadesCount, isPointLightPerPixelEnabled
        currValue = shadowFilterSize
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add "Custom" item
            addCustomChoice result, currValue
        End If
        Set graphicsoptions_c_fn = result

    Case FN_gfx_shadowFiltering_select          ' [](const auto& newFilterSize) -> bool
        newValue = CLng(args(0))
        If Not supportsShadowMappingOrFalse() Then
            graphicsoptions_c_fn = False
            Exit Function
        End If
        gfx_api_context_getShadowConstants shadowMode_, shadowFilterSize, shadowCascadesCount, isPointLightPerPixelEnabled
        shadowFilterSize = newValue
        If Not gfx_api_context_setShadowConstants(shadowMode_, shadowFilterSize, shadowCascadesCount, isPointLightPerPixelEnabled) Then
            debug_ 240, LOG_ERROR, "makeGraphicsOptionsForm", "Failed to set shadow filter size: %u", newValue
            graphicsoptions_c_fn = False
            Exit Function
        End If
        war_setShadowFilterSize newValue
        graphicsoptions_c_fn = True

    Case FN_gfx_pointLights_populate            ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Lightmap", "Use a classic lightmap to calculate point lights impacting objects and tiles. (Classic & Normal default. Fastest performance.)", False
        result.addChoice "Per-Pixel (HQ)", "Use advanced per-pixel point-lighting to improve visuals. (May impact performance. Requires a more powerful GPU.)", True, (getTerrainShaderQuality() <> NORMAL_MAPPING)
        result.setCurrentIdxForValue war_getPointLightPerPixelLighting()
        Set graphicsoptions_c_fn = result

    Case FN_gfx_pointLights_select              ' [](const auto& newValue) -> bool
        If getTerrainShaderQuality() <> NORMAL_MAPPING Then
            ' point light per pixel lighting is only supported in normal_mapping mode
            If CBool(args(0)) Then
                graphicsoptions_c_fn = False
                Exit Function
            End If
        End If
        gfx_api_context_getShadowConstants shadowMode_, shadowFilterSize, shadowCascadesCount, isPointLightPerPixelEnabled
        isPointLightPerPixelEnabled = CBool(args(0))
        If Not gfx_api_context_setShadowConstants(shadowMode_, shadowFilterSize, shadowCascadesCount, isPointLightPerPixelEnabled) Then
            debug_ 279, LOG_ERROR, "makeGraphicsOptionsForm", "Failed to set per pixel point lighting value: %d", IIf(CBool(args(0)), 1, 0)
            graphicsoptions_c_fn = False
            Exit Function
        End If
        war_setPointLightPerPixelLighting CBool(args(0))
        graphicsoptions_c_fn = True

    Case FN_gfx_fog_populate                    ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue pie_GetFogEnabled()
        Set graphicsoptions_c_fn = result

    Case FN_gfx_fog_select
        If pie_GetFogEnabled() Then
            pie_SetFogStatus False
            pie_EnableFog False
        Else
            pie_EnableFog True
        End If
        graphicsoptions_c_fn = True

    Case FN_gfx_screenShake_populate            ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue getShakeStatus()
        Set graphicsoptions_c_fn = result

    Case FN_gfx_screenShake_select
        setShakeStatus CBool(args(0))
        graphicsoptions_c_fn = True

    Case FN_gfx_textureSizeLimit_populate       ' OptionChoices<uint32_t> ()
        Set result = New OptionChoices
        result.addChoice "512", "", 512
        result.addChoice "1024", "", 1024
        result.addChoice "2048", "", 2048
        '    { "4096", "", 4096 },
        currValue = getTextureSize()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add "Custom" item
            addCustomChoice result, currValue
        End If
        Set graphicsoptions_c_fn = result

    Case FN_gfx_textureSizeLimit_select         ' [](const auto& newTexSize) -> bool
        setTextureSize CLng(args(0))
        graphicsoptions_c_fn = True

    Case FN_gfx_lodDistance_populate            ' OptionChoices<int> ()
        Set result = New OptionChoices
        result.addChoice "Default", "The system default (no supplied bias, may result in less sharp visuals).", 0  ' the *system* default (no supplied bias)
        result.addChoice "High", "Use LOD Distance bias to increase visual sharpness.", WZ_LODDISTANCEPERCENTAGE_HIGH
        currValue = war_getLODDistanceBiasPercentage()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add "Custom" item (C++ formats the int with PRIu32)
            result.addChoice "(Custom: " & CStr(U32(currValue)) & ")", "", currValue
            result.currentIdx = result.count - 1
        End If
        Set graphicsoptions_c_fn = result

    Case FN_gfx_lodDistance_select
        war_setLODDistanceBiasPercentage CLng(args(0))
        graphicsoptions_c_fn = True

    Case FN_gfx_backend_populate                ' OptionChoices<video_backend> ()
        availableBackendsCount = wzAvailableGfxBackends(availableBackends)
        Set result = New OptionChoices
        For i = 0 To availableBackendsCount - 1
            result.addChoice to_display_string_video_backend(availableBackends(i)), "", availableBackends(i)
        Next
        currValue = war_getGfxBackend()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' somehow the current graphics backend isn't in the available list... add it
            result.addChoice to_display_string_video_backend(currValue), "", currValue
            result.currentIdx = result.count - 1
        End If
        Set graphicsoptions_c_fn = result

    Case FN_gfx_backend_select
        war_setGfxBackend CLng(args(0))
        graphicsoptions_c_fn = True

    Case FN_gfx_antialiasing_populate           ' OptionChoices<int> ()
        maxAA = pie_GetMaxAntialiasing()
        Set result = New OptionChoices
        result.addChoice "Off", "", 0
        value = 2
        Do While value <= maxAA
            result.addChoice CStr(value) & "x", "", value
            value = value * 2
        Loop
        currValue = war_getAntialiasing()
        If Not result.setCurrentIdxForValue(currValue) Then
            result.addChoice CStr(currValue) & "x", "", currValue
            result.currentIdx = result.count - 1
        End If
        Set graphicsoptions_c_fn = result

    Case FN_gfx_antialiasing_select
        war_setAntialiasing CLng(args(0))
        graphicsoptions_c_fn = True

    Case FN_gfx_vsync_populate                  ' OptionChoices<gfx_api::context::swap_interval_mode> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", swap_interval_mode_immediate
        result.addChoice "On", "", swap_interval_mode_vsync
        result.addChoice "Adaptive", "", swap_interval_mode_adaptive_vsync
        result.setCurrentIdxForValue to_swap_mode(war_GetVsync())
        Set graphicsoptions_c_fn = result

    Case FN_gfx_vsync_select                    ' [](const auto& newValue) -> bool
        newValue = CLng(args(0))
        ' (VB) gfx_api::context::get().setSwapInterval(newValue, []() { wzPostChangedSwapInterval(); }): WzGfx's
        ' setSwapInterval takes no callback; it is called here once the change succeeded
        canChange = gfx_api_context_setSwapInterval(newValue)
        If canChange Then wzPostChangedSwapInterval
        If Not canChange Then
            debug_ 477, LOG_INFO, "makeGraphicsOptionsForm", "Vsync mode not supported by system: %d", newValue
            graphicsoptions_c_fn = False
            Exit Function
        End If
        war_SetVsync newValue
        graphicsoptions_c_fn = True
    End Select
End Function
