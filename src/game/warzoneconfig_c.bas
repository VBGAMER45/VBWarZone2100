Attribute VB_Name = "warzoneconfig_c"
' src/warzoneconfig.cpp, with the enums of warzoneconfig.h. WARZONE_GLOBALS
' is a UDT whose defaults war_SetDefaultStates sets (C++ initialises the
' struct's members). C++'s overloaded to_string becomes
' to_string_JS_BACKEND and to_string_ConnectionProviderType.
Option Explicit

Private Const MAX_REPLAY_FILES As Long = 36
Private Const MAX_OLD_LOGS As Long = 50

Public Const CAMERASPEED_MAX As Long = 5000
Public Const CAMERASPEED_MIN As Long = 100
Public Const CAMERASPEED_DEFAULT As Long = 2500
Public Const CAMERASPEED_STEP As Long = 100
Public Const MIN_MPINACTIVITY_MINUTES As Long = 4
Public Const MIN_MPGAMETIMELIMIT_MINUTES As Long = 30
Public Const WZ_LODDISTANCEPERCENTAGE_HIGH As Long = -50

' enum FMV_MODE
Public Const FMV_FULLSCREEN As Long = 0
Public Const FMV_1X As Long = 1
Public Const FMV_2X As Long = 2
Public Const FMV_MAX As Long = 3

' enum class JS_BACKEND
Public Const JS_BACKEND_quickjs As Long = 0
Public Const JS_BACKEND_num_backends As Long = 1 ' Must be last!

' enum class TrapCursorMode : uint8_t
Public Const TrapCursorMode_Disabled As Long = 0
Public Const TrapCursorMode_Enabled As Long = 1
Public Const TrapCursorMode_Automatic As Long = 2

Private Type WARZONE_GLOBALS
    FMVmode As Long
    width As Long
    height As Long
    videoBufferDepth As Long
    displayScale As Long
    screen As Long
    SPcolor As Long                 ' int8_t
    MPcolour As Long
    antialiasing As Long
    Fullscreen As Long              ' WINDOW_MODE
    soundEnabled As Boolean
    trapCursor As Long
    vsync As Long
    pauseOnFocusLoss As Boolean
    ColouredCursor As Boolean
    MusicEnabled As Boolean
    hrtfMode As Long
    mapZoom As Long
    mapZoomRate As Long
    radarZoom As Long
    cameraSpeed As Long
    radarJump As Boolean
    gfxBackend As Long
    jsBackend As Long
    autoAdjustDisplayScale As Boolean
    autoLagKickSeconds As Long
    autoLagKickAggressiveness As Long
    autoDesyncKickSeconds As Long
    autoNotReadyKickSeconds As Long
    disableReplayRecording As Boolean
    maxReplaysSaved As Long
    oldLogsLimit As Long
    MPinactivityMinutes As Long
    MPgameTimeLimitMinutes As Long
    MPopenSpectatorSlots As Long
    MPplayerLeaveMode As Long
    lastIPConnectServerName As String
    fogStart As Long
    fogEnd As Long
    lodDistanceBiasPercentage As Long
    minimizeOnFocusLoss As Long
    ' fullscreen mode settings
    fullscreenModeWidth As Long
    fullscreenModeHeight As Long
    fullscreenModeScreen As Long
    fullscreenPixelDensity As Single
    fullscreenRefreshRate As Single
    toggleFullscreenMode As Long
    cursorScale As Long
    ' shadow mapping settings
    shadowFilteringMode As Long
    shadowFilterSize As Long
    shadowMapResolution As Long
    pointLightLighting As Boolean
    ' UI config
    groupsMenuEnabled As Boolean
    optionsButtonVisibility As Long
    ' run-time only settings (not persisted to config!)
    allowVulkanImplicitLayers As Boolean
    ' Connection provider used for hosting games
    hostProviderType As Long
    ' lobby
    lobbyDisableIPv6 As Boolean
    lobbyFilterIPv6Only As Boolean
    ' audio cues
    playAudioCue_GroupReporting As Boolean
End Type

Private warGlobs As WARZONE_GLOBALS

' std::to_string of a bool: "1" or "0"
Private Function BoolStr(ByVal b As Boolean) As String
    If b Then BoolStr = "1" Else BoolStr = "0"
End Function

Private Function js_backend_names(ByVal i As Long) As String
    Select Case i
    Case JS_BACKEND_quickjs: js_backend_names = "quickjs"
    Case Else: js_backend_names = "invalid" ' Must be last!
    End Select
End Function

Public Function js_backend_from_str(ByVal str_ As String, output_backend As Long) As Boolean
    Dim i As Long
    For i = 0 To JS_BACKEND_num_backends - 1
        If StrComp(js_backend_names(i), str_, vbTextCompare) = 0 Then
            output_backend = i
            js_backend_from_str = True
            Exit Function
        End If
    Next
End Function

Public Function to_string_JS_BACKEND(ByVal backend As Long) As String
    to_string_JS_BACKEND = js_backend_names(backend)
End Function

'**************************************************************************

Public Sub war_SetDefaultStates()
    Dim fresh As WARZONE_GLOBALS
    warGlobs = fresh
    With warGlobs
        .FMVmode = FMV_FULLSCREEN
        .width = 1024
        .height = 768
        .videoBufferDepth = 32
        .displayScale = 100
        .screen = 0
        .SPcolor = 0
        .MPcolour = -1
        .antialiasing = 0
        .Fullscreen = WINDOW_MODE_windowed ' Leave this to windowed, some system will fail and they can't see the system popup dialog!
        .soundEnabled = True
        .trapCursor = TrapCursorMode_Automatic
        .vsync = 1
        .pauseOnFocusLoss = False
        .ColouredCursor = True
        .MusicEnabled = True
        .hrtfMode = HRTFMode_Auto
        .mapZoom = STARTDISTANCE
        .mapZoomRate = MAP_ZOOM_RATE_DEFAULT
        .radarZoom = DEFAULT_RADARZOOM
        .cameraSpeed = CAMERASPEED_DEFAULT
        .radarJump = False
        .gfxBackend = video_backend_opengl ' the actual default value is determined in loadConfig()
        .jsBackend = 0
        .autoAdjustDisplayScale = True
        .autoLagKickSeconds = 60
        .autoLagKickAggressiveness = 3
        .autoDesyncKickSeconds = 10
        .autoNotReadyKickSeconds = 0
        .disableReplayRecording = False
        .maxReplaysSaved = MAX_REPLAY_FILES
        .oldLogsLimit = MAX_OLD_LOGS
        .MPinactivityMinutes = 5
        .MPgameTimeLimitMinutes = 0 ' default to unlimited
        .MPopenSpectatorSlots = 0
        .MPplayerLeaveMode = PLAYER_LEAVE_MODE_DEFAULT
        .lastIPConnectServerName = ""
        .fogStart = 4000
        .fogEnd = 8000
        .lodDistanceBiasPercentage = WZ_LODDISTANCEPERCENTAGE_HIGH ' default to "High" to best match prior version behavior
        .minimizeOnFocusLoss = -1 ' see enum class MinimizeOnFocusLossBehavior
        .fullscreenModeWidth = 0 ' current display default
        .fullscreenModeHeight = 0 ' current display default
        .fullscreenModeScreen = -1
        .fullscreenPixelDensity = 1!
        .fullscreenRefreshRate = 0! ' current display default
        .toggleFullscreenMode = 0 ' 0 = the backend default
        .cursorScale = 100
        .shadowFilteringMode = 1
        .shadowFilterSize = 5
        .shadowMapResolution = 0 ' this defaults to 0, which causes the gfx backend to figure out a recommended default based on the system properties
        .pointLightLighting = False
        .groupsMenuEnabled = True
        .optionsButtonVisibility = 100
        .allowVulkanImplicitLayers = False
        .hostProviderType = ConnectionProviderType_TCP_DIRECT
        .lobbyDisableIPv6 = False
        .lobbyFilterIPv6Only = False
        .playAudioCue_GroupReporting = True
    End With
End Sub

Public Sub war_SetSPcolor(ByVal color As Long)
    If color >= 1 And color <= 3 Then ' only 0,4,5,6,7 are allowed for SP games, AI uses the other colors.
        color = 0
    End If
    warGlobs.SPcolor = ToInt8(color)
    setPlayerColour 0, color
    ActivityManager_changedSetting "SPcolor", CStr(color)
End Sub

Public Function war_GetSPcolor() As Long
    war_GetSPcolor = warGlobs.SPcolor
End Function

Public Sub war_setMPcolour(ByVal colour As Long)
    warGlobs.MPcolour = colour
    ActivityManager_changedSetting "MPcolour", CStr(colour)
End Sub

Public Function war_getMPcolour() As Long
    war_getMPcolour = warGlobs.MPcolour
End Function

Public Sub war_setWindowMode(ByVal b As Long)
    warGlobs.Fullscreen = b
End Sub

Public Function war_getWindowMode() As WINDOW_MODE
    war_getWindowMode = warGlobs.Fullscreen
End Function

Public Sub war_setAntialiasing(ByVal antialiasing As Long)
    If antialiasing > 16 Then
        debug_ 203, LOG_WARNING, "war_setAntialiasing", "Antialising set to value > 16, which can cause crashes."
    End If
    warGlobs.antialiasing = antialiasing
    ActivityManager_changedSetting "antialiasing", CStr(antialiasing)
End Sub

Public Function war_getAntialiasing() As Long
    war_getAntialiasing = warGlobs.antialiasing
End Function

Public Sub war_SetTrapCursor(ByVal v As Long)
    warGlobs.trapCursor = v
    ActivityManager_changedSetting "trapCursor", CStr(v)
End Sub

Public Function war_GetTrapCursor() As Long
    war_GetTrapCursor = warGlobs.trapCursor
End Function

Public Sub war_SetVsync(ByVal value As Long)
    warGlobs.vsync = value
End Sub

Public Function war_GetVsync() As Long
    war_GetVsync = warGlobs.vsync
End Function

Public Function war_GetDisplayScale() As Long
    war_GetDisplayScale = warGlobs.displayScale
End Function

Public Sub war_SetDisplayScale(ByVal scale_ As Long)
    warGlobs.displayScale = scale_
    ActivityManager_changedSetting "displayScale", Format$(U32(scale_), "0")
End Sub

Public Sub war_setCursorScale(ByVal scale_ As Long)
    If U32(scale_) < 100 Then scale_ = 100
    warGlobs.cursorScale = scale_
End Sub

Public Function war_getCursorScale() As Long
    war_getCursorScale = warGlobs.cursorScale
End Function

Public Sub war_SetWidth(ByVal width As Long)
    warGlobs.width = width
End Sub

Public Function war_GetWidth() As Long
    war_GetWidth = warGlobs.width
End Function

Public Sub war_SetHeight(ByVal height As Long)
    warGlobs.height = height
End Sub

Public Function war_GetHeight() As Long
    war_GetHeight = warGlobs.height
End Function

Public Sub war_SetVideoBufferDepth(ByVal videoBufferDepth As Long)
    warGlobs.videoBufferDepth = videoBufferDepth
End Sub

Public Function war_GetVideoBufferDepth() As Long
    war_GetVideoBufferDepth = warGlobs.videoBufferDepth
End Function

Public Sub war_SetScreen(ByVal screen As Long)
    warGlobs.screen = screen
End Sub

Public Function war_GetScreen() As Long
    war_GetScreen = warGlobs.screen
End Function

Public Sub war_SetFMVmode(ByVal mode As Long)
    warGlobs.FMVmode = mode Mod FMV_MAX
End Sub

Public Function war_GetFMVmode() As Long
    war_GetFMVmode = warGlobs.FMVmode
End Function

Public Sub war_setScanlineMode(ByVal mode As Long)
    debug_ 308, LOG_VIDEO, "war_setScanlineMode", "%d", mode
    seq_setScanlineMode mode
End Sub

Public Function war_getScanlineMode() As Long
    debug_ 314, LOG_VIDEO, "war_getScanlineMode", "%d", seq_getScanlineMode()
    war_getScanlineMode = seq_getScanlineMode()
End Function

Public Sub war_SetPauseOnFocusLoss(ByVal enabled As Boolean)
    warGlobs.pauseOnFocusLoss = enabled
    ActivityManager_changedSetting "pauseOnFocusLoss", BoolStr(enabled)
End Sub

Public Function war_GetPauseOnFocusLoss() As Boolean
    war_GetPauseOnFocusLoss = warGlobs.pauseOnFocusLoss
End Function

Public Sub war_SetColouredCursor(ByVal enabled As Boolean)
    warGlobs.ColouredCursor = enabled
    ActivityManager_changedSetting "ColouredCursor", BoolStr(enabled)
End Sub

Public Function war_GetColouredCursor() As Boolean
    war_GetColouredCursor = warGlobs.ColouredCursor
End Function

Public Sub war_setSoundEnabled(ByVal soundEnabled As Boolean)
    warGlobs.soundEnabled = soundEnabled
    ActivityManager_changedSetting "soundEnabled", BoolStr(soundEnabled)
End Sub

Public Function war_getSoundEnabled() As Boolean
    war_getSoundEnabled = warGlobs.soundEnabled
End Function

Public Function war_GetMusicEnabled() As Boolean
    war_GetMusicEnabled = warGlobs.MusicEnabled
End Function

Public Sub war_SetMusicEnabled(ByVal enabled As Boolean)
    warGlobs.MusicEnabled = enabled
    ActivityManager_changedSetting "musicEnabled", BoolStr(enabled)
End Sub

Public Function war_GetHRTFMode() As Long
    war_GetHRTFMode = warGlobs.hrtfMode
End Function

Public Sub war_SetHRTFMode(ByVal mode As Long)
    warGlobs.hrtfMode = mode
    ActivityManager_changedSetting "hrtfMode", CStr(mode)
End Sub

Public Function war_GetMapZoom() As Long
    war_GetMapZoom = warGlobs.mapZoom
End Function

Public Sub war_SetMapZoom(ByVal mapZoom As Long)
    If mapZoom Mod MAP_ZOOM_CONFIG_STEP = 0 And Not (mapZoom < MINDISTANCE_CONFIG Or mapZoom > MAXDISTANCE) Then
        warGlobs.mapZoom = mapZoom
        ActivityManager_changedSetting "mapZoom", CStr(mapZoom)
    End If
End Sub

Public Function war_GetMapZoomRate() As Long
    war_GetMapZoomRate = warGlobs.mapZoomRate
End Function

Public Sub war_SetMapZoomRate(ByVal mapZoomRate As Long)
    If mapZoomRate Mod MAP_ZOOM_RATE_STEP = 0 And Not (mapZoomRate < MAP_ZOOM_RATE_MIN Or mapZoomRate > MAP_ZOOM_RATE_MAX) Then
        warGlobs.mapZoomRate = mapZoomRate
        ActivityManager_changedSetting "mapZoomRate", CStr(mapZoomRate)
    End If
End Sub

Public Function war_GetRadarZoom() As Long
    war_GetRadarZoom = warGlobs.radarZoom
End Function

Public Sub war_SetRadarZoom(ByVal radarZoom As Long)
    If radarZoom Mod RADARZOOM_STEP = 0 And Not (radarZoom < MIN_RADARZOOM Or radarZoom > MAX_RADARZOOM) Then
        warGlobs.radarZoom = radarZoom
        ActivityManager_changedSetting "radarZoom", CStr(radarZoom)
    End If
End Sub

Public Function war_GetCameraSpeed() As Long
    war_GetCameraSpeed = warGlobs.cameraSpeed
End Function

Public Sub war_SetCameraSpeed(ByVal cameraSpeed As Long)
    If cameraSpeed Mod CAMERASPEED_STEP = 0 And Not (cameraSpeed < CAMERASPEED_MIN Or cameraSpeed > CAMERASPEED_MAX) Then
        warGlobs.cameraSpeed = cameraSpeed
        ActivityManager_changedSetting "cameraSpeed", CStr(cameraSpeed)
    End If
End Sub

Public Function war_GetRadarJump() As Boolean
    war_GetRadarJump = warGlobs.radarJump
End Function

Public Sub war_SetRadarJump(ByVal radarJump As Boolean)
    warGlobs.radarJump = radarJump
    ActivityManager_changedSetting "radarJump", BoolStr(radarJump)
End Sub

Public Function war_getLODDistanceBiasPercentage() As Long
    war_getLODDistanceBiasPercentage = warGlobs.lodDistanceBiasPercentage
End Function

Public Sub war_setLODDistanceBiasPercentage(ByVal bias As Long)
    If bias > 200 Then bias = 200
    If bias < -200 Then bias = -200
    warGlobs.lodDistanceBiasPercentage = bias
End Sub

Public Function war_getMinimizeOnFocusLoss() As Long
    war_getMinimizeOnFocusLoss = warGlobs.minimizeOnFocusLoss
End Function

Public Sub war_setMinimizeOnFocusLoss(ByVal val_ As Long)
    warGlobs.minimizeOnFocusLoss = val_
End Sub

Public Function war_getGfxBackend() As Long
    war_getGfxBackend = warGlobs.gfxBackend
End Function

Public Sub war_setGfxBackend(ByVal backend As Long)
    warGlobs.gfxBackend = backend
End Sub

Public Function war_getJSBackend() As Long
    war_getJSBackend = warGlobs.jsBackend
End Function

Public Sub war_setJSBackend(ByVal backend As Long)
    warGlobs.jsBackend = backend
End Sub

Public Function war_getAutoAdjustDisplayScale() As Boolean
    war_getAutoAdjustDisplayScale = warGlobs.autoAdjustDisplayScale
End Function

Public Sub war_setAutoAdjustDisplayScale(ByVal autoAdjustDisplayScale As Boolean)
    warGlobs.autoAdjustDisplayScale = autoAdjustDisplayScale
End Sub

Public Function war_getAutoLagKickSeconds() As Long
    war_getAutoLagKickSeconds = warGlobs.autoLagKickSeconds
End Function

Public Sub war_setAutoLagKickSeconds(ByVal seconds As Long)
    If seconds < 0 Then seconds = 0
    If seconds > 0 Then
        If seconds < 15 Then seconds = 15
    End If
    warGlobs.autoLagKickSeconds = seconds
End Sub

Public Function war_getAutoLagKickAggressiveness() As Long
    war_getAutoLagKickAggressiveness = warGlobs.autoLagKickAggressiveness
End Function

Public Sub war_setAutoLagKickAggressiveness(ByVal aggressiveness As Long)
    If aggressiveness < 1 Then aggressiveness = 1
    If aggressiveness > 10 Then aggressiveness = 10
    warGlobs.autoLagKickAggressiveness = aggressiveness
End Sub

Public Function war_getAutoDesyncKickSeconds() As Long
    war_getAutoDesyncKickSeconds = warGlobs.autoDesyncKickSeconds
End Function

Public Sub war_setAutoDesyncKickSeconds(ByVal seconds As Long)
    If seconds < 0 Then seconds = 0
    If seconds > 0 Then
        If seconds < 10 Then seconds = 10
    End If
    warGlobs.autoDesyncKickSeconds = seconds
End Sub

Public Function war_getAutoNotReadyKickSeconds() As Long
    war_getAutoNotReadyKickSeconds = warGlobs.autoNotReadyKickSeconds
End Function

Public Sub war_setAutoNotReadyKickSeconds(ByVal seconds As Long)
    If seconds < 0 Then seconds = 0
    If seconds > 0 Then
        If seconds < 15 Then seconds = 15
    End If
    warGlobs.autoNotReadyKickSeconds = seconds
End Sub

Public Function war_getDisableReplayRecording() As Boolean
    war_getDisableReplayRecording = warGlobs.disableReplayRecording
End Function

Public Sub war_setDisableReplayRecording(ByVal disable As Boolean)
    warGlobs.disableReplayRecording = disable
End Sub

Public Function war_getMaxReplaysSaved() As Long
    war_getMaxReplaysSaved = warGlobs.maxReplaysSaved
End Function

Public Sub war_setMaxReplaysSaved(ByVal maxReplaysSaved As Long)
    warGlobs.maxReplaysSaved = maxReplaysSaved
End Sub

Public Function war_getOldLogsLimit() As Long
    war_getOldLogsLimit = warGlobs.oldLogsLimit
End Function

Public Sub war_setOldLogsLimit(ByVal oldLogsLimit As Long)
    warGlobs.oldLogsLimit = oldLogsLimit
End Sub

Public Function war_getMPInactivityMinutes() As Long
    war_getMPInactivityMinutes = warGlobs.MPinactivityMinutes
End Function

Public Sub war_setMPInactivityMinutes(ByVal minutes As Long)
    If minutes <> 0 And ULT(minutes, MIN_MPINACTIVITY_MINUTES) Then
        minutes = MIN_MPINACTIVITY_MINUTES
    End If
    warGlobs.MPinactivityMinutes = minutes
End Sub

Public Function war_getMPGameTimeLimitMinutes() As Long
    war_getMPGameTimeLimitMinutes = warGlobs.MPgameTimeLimitMinutes
End Function

Public Sub war_setMPGameTimeLimitMinutes(ByVal minutes As Long)
    If minutes <> 0 And ULT(minutes, MIN_MPGAMETIMELIMIT_MINUTES) Then
        minutes = MIN_MPGAMETIMELIMIT_MINUTES
    End If
    warGlobs.MPgameTimeLimitMinutes = minutes
End Sub

Public Function war_getMPopenSpectatorSlots() As Long
    war_getMPopenSpectatorSlots = warGlobs.MPopenSpectatorSlots
End Function

' spectatorSlots: uint16_t; stored as uint8_t
Public Sub war_setMPopenSpectatorSlots(ByVal spectatorSlots As Long)
    spectatorSlots = spectatorSlots And &HFFFF&
    If spectatorSlots > MAX_SPECTATOR_SLOTS Then spectatorSlots = MAX_SPECTATOR_SLOTS
    warGlobs.MPopenSpectatorSlots = spectatorSlots And &HFF&
End Sub

Public Function war_getMPPlayerLeaveMode() As Long
    war_getMPPlayerLeaveMode = warGlobs.MPplayerLeaveMode
End Function

Public Sub war_setMPPlayerLeaveMode(ByVal mode As Long)
    warGlobs.MPplayerLeaveMode = mode
End Sub

Public Sub war_setLastIpServerConnect(ByVal serverName As String)
    warGlobs.lastIPConnectServerName = serverName
End Sub

Public Function war_getLastIpServerConnect() As String
    war_getLastIpServerConnect = warGlobs.lastIPConnectServerName
End Function

Public Function war_getFogEnd() As Long
    war_getFogEnd = warGlobs.fogEnd
End Function

Public Function war_getFogStart() As Long
    war_getFogStart = warGlobs.fogStart
End Function

Public Sub war_setFogEnd(ByVal end_ As Long)
    warGlobs.fogEnd = end_
End Sub

Public Sub war_setFogStart(ByVal start As Long)
    warGlobs.fogStart = start
End Sub

Public Sub war_SetFullscreenModeWidth(ByVal width As Long)
    warGlobs.fullscreenModeWidth = width
End Sub

Public Function war_GetFullscreenModeWidth() As Long
    war_GetFullscreenModeWidth = warGlobs.fullscreenModeWidth
End Function

Public Sub war_SetFullscreenModeHeight(ByVal height As Long)
    warGlobs.fullscreenModeHeight = height
End Sub

Public Function war_GetFullscreenModeHeight() As Long
    war_GetFullscreenModeHeight = warGlobs.fullscreenModeHeight
End Function

Public Sub war_SetFullscreenModeScreen(ByVal screen As Long)
    warGlobs.fullscreenModeScreen = screen
End Sub

Public Function war_GetFullscreenModeScreen() As Long
    war_GetFullscreenModeScreen = warGlobs.fullscreenModeScreen
End Function

Public Function war_GetFullscreenModePixelDensity() As Single
    war_GetFullscreenModePixelDensity = warGlobs.fullscreenPixelDensity
End Function

Public Sub war_SetFullscreenModePixelDensity(ByVal pixelDensity As Single)
    warGlobs.fullscreenPixelDensity = pixelDensity
End Sub

Public Function war_GetFullscreenModeRefreshRate() As Single
    war_GetFullscreenModeRefreshRate = warGlobs.fullscreenRefreshRate
End Function

Public Sub war_SetFullscreenModeRefreshRate(ByVal refreshRate As Single)
    warGlobs.fullscreenRefreshRate = refreshRate
End Sub

Public Sub war_setToggleFullscreenMode(ByVal mode As Long)
    warGlobs.toggleFullscreenMode = mode
End Sub

Public Function war_getToggleFullscreenMode() As Long
    war_getToggleFullscreenMode = warGlobs.toggleFullscreenMode
End Function

Public Function war_getShadowFilterSize() As Long
    war_getShadowFilterSize = warGlobs.shadowFilterSize
End Function

Public Sub war_setShadowFilterSize(ByVal filterSize As Long)
    If ULT(7, filterSize) Then
        debug_ 726, LOG_INFO, "war_setShadowFilterSize", "Shadow filter size %u may not have the desired effect", filterSize
    End If
    warGlobs.shadowFilterSize = filterSize
End Sub

Public Function war_getShadowMapResolution() As Long
    war_getShadowMapResolution = warGlobs.shadowMapResolution
End Function

Public Sub war_setShadowMapResolution(ByVal resolution As Long)
    If resolution <> 0 And ((resolution And (resolution - 1)) <> 0) Then
        debug_ 740, LOG_ERROR, "war_setShadowMapResolution", "Shadow map resolution %u is unsupported: must be a power of 2", resolution
        Exit Sub
    End If
    If resolution <> 0 And (ULT(resolution, 2048) Or ULT(4096, resolution)) Then ' 0 is a special case that maps to "figure out a decent default for this system"
        debug_ 745, LOG_INFO, "war_setShadowMapResolution", "Shadow map resolution %u may not have the desired effect", resolution
    End If
    warGlobs.shadowMapResolution = resolution
End Sub

Public Function war_getPointLightPerPixelLighting() As Boolean
    war_getPointLightPerPixelLighting = warGlobs.pointLightLighting
End Function

Public Sub war_setPointLightPerPixelLighting(ByVal perPixelEnabled As Boolean)
    warGlobs.pointLightLighting = perPixelEnabled
End Sub

Public Function war_getGroupsMenuEnabled() As Boolean
    war_getGroupsMenuEnabled = warGlobs.groupsMenuEnabled
End Function

Public Sub war_setGroupsMenuEnabled(ByVal enabled As Boolean)
    warGlobs.groupsMenuEnabled = enabled
End Sub

Public Sub war_runtimeOnlySetAllowVulkanImplicitLayers(ByVal allowed As Boolean) ' not persisted to config
    warGlobs.allowVulkanImplicitLayers = allowed
End Sub

Public Function war_getOptionsButtonVisibility() As Long
    war_getOptionsButtonVisibility = warGlobs.optionsButtonVisibility
End Function

' val_: uint8_t
Public Sub war_setOptionsButtonVisibility(ByVal val_ As Long)
    val_ = val_ And &HFF&
    If val_ > 100 Then val_ = 100
    If val_ > 0 Then
        If val_ < 50 Then val_ = 50
    End If
    warGlobs.optionsButtonVisibility = val_
End Sub

Public Function war_getAllowVulkanImplicitLayers() As Boolean
    war_getAllowVulkanImplicitLayers = warGlobs.allowVulkanImplicitLayers
End Function

Public Sub war_setHostConnectionProvider(ByVal pt As Long)
    warGlobs.hostProviderType = pt
End Sub

Public Function war_getHostConnectionProvider() As Long
    war_getHostConnectionProvider = warGlobs.hostProviderType
End Function

Public Function net_backend_from_str(ByVal str_ As String, pt As Long) As Boolean
    If StrComp(str_, "tcp", vbTextCompare) = 0 Then
        pt = ConnectionProviderType_TCP_DIRECT
        net_backend_from_str = True
        Exit Function
    End If
End Function

Public Function to_string_ConnectionProviderType(ByVal pt As Long) As String
    Select Case pt
    Case ConnectionProviderType_TCP_DIRECT
        to_string_ConnectionProviderType = "tcp"
        Exit Function
    End Select
    ASSERT_FAILURE "false", "src/warzoneconfig.cpp:833", "to_string", "Invalid connection provider type enumeration value: %d", pt ' silence GCC warning
End Function

Public Function war_getLobbyDisableIPv6() As Boolean
    war_getLobbyDisableIPv6 = warGlobs.lobbyDisableIPv6
End Function

Public Sub war_setLobbyDisableIPv6(ByVal enabled As Boolean)
    warGlobs.lobbyDisableIPv6 = enabled
End Sub

Public Function war_getLobbyFilterIPv6Only() As Boolean
    war_getLobbyFilterIPv6Only = warGlobs.lobbyFilterIPv6Only
End Function

Public Sub war_setLobbyFilterIPv6Only(ByVal enabled As Boolean)
    warGlobs.lobbyFilterIPv6Only = enabled
End Sub

Public Function war_getPlayAudioCue_GroupReporting() As Boolean
    war_getPlayAudioCue_GroupReporting = warGlobs.playAudioCue_GroupReporting
End Function

Public Sub war_setPlayAudioCue_GroupReporting(ByVal val_ As Boolean)
    warGlobs.playAudioCue_GroupReporting = val_
End Sub
