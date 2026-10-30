Attribute VB_Name = "configuration_c"
' src/configuration.cpp. An optional<T> result is a Boolean (has_value)
' with the value ByRef; an optional<MOUSE_KEY_CODE> is a Long, 0 = nullopt.
' loadConfig's lambdas are the module's iniGet* functions on iniGeneral.
' C++ quirk kept: loadConfig's iniGetFloat lambda takes and returns
' optional<int>, so the float settings it reads are truncated to integers.
' The std::stoi / std::stof exception texts are MSVC's. Translations (_())
' are not ported: the English text is used.
Option Explicit

Private Const GAMESERVERPORT As Long = 2100
Private Const BASECONFVERSION As Long = 1
Private Const CURRCONFVERSION As Long = 4

' configuration.h
Public Const DEFAULTSKIRMISHMAP As String = "Sk-Rush"
Public Const DEFAULTSKIRMISHMAPMAXPLAYERS As Long = 4

Private Const fileName As String = "config"
Private Const MAX_CONFIG_FILE_SIZE As Double = 2097152#     ' 2 MB seems like enough...

Private iniGeneral As IniSection        ' what loadConfig's lambdas capture

'////////////////////////////////////////////////////////////////////////////
' PhysFS helpers

Private Function WZ_PHYSFS_getRealPath(ByVal filename_ As String) As String
    Dim fullPath As String
    fullPath = PHYSFS_getRealDir(filename_)
    If Len(fullPath) = 0 Then
        WZ_PHYSFS_getRealPath = fullPath
        Exit Function
    End If
    fullPath = fullPath & PHYSFS_getDirSeparator()
    fullPath = fullPath & filename_
    WZ_PHYSFS_getRealPath = fullPath
End Function

' C: optional<uint64_t> WZ_PHYSFS_getFileSize(const char *filename)
Private Function WZ_PHYSFS_getFileSize(ByVal filename_ As String, size As Double) As Boolean
    Dim filesize As Double, modtime As Double, filetype As Long, ro As Long
    If PHYSFS_stat(filename_, filesize, modtime, filetype, ro) = 0 Then
        Exit Function ' failed to get file info
    End If
    If filesize < 0 Then
        ' unknown filesize
        Exit Function
    End If
    size = filesize
    WZ_PHYSFS_getFileSize = True
End Function

'////////////////////////////////////////////////////////////////////////////

' C: std::stoi: base 10, leading white space, at least one digit; invalid_argument / out_of_range.
Private Function stoi(ByVal s As String, value As Long, errText As String) As Boolean
    Dim endIdx As Long, d As Double
    d = strtoll(s, endIdx, 10)
    If endIdx = 0 Then
        errText = "invalid stoi argument"
        Exit Function
    End If
    If d > 2147483647# Or d < -2147483648# Then
        errText = "stoi argument out of range"
        Exit Function
    End If
    value = d
    stoi = True
End Function

' C: optional<int> iniSectionGetInteger(iniSection, key, optional<int> defaultValue = nullopt)
Private Function iniSectionGetInteger(ByVal iniSection As IniSection, ByVal key As String, value As Long, Optional ByVal hasDefault As Boolean = False, Optional ByVal defaultValue As Long = 0) As Boolean
    Dim valueStr As String, valueInt As Long, errText As String
    If Not iniSection.HasValue(key) Then
        value = defaultValue
        iniSectionGetInteger = hasDefault
        Exit Function
    End If
    valueStr = iniSection.Get_(key, "")
    If stoi(valueStr, valueInt, errText) Then
        value = valueInt
        iniSectionGetInteger = True
    Else
        debug_ 120, LOG_ERROR, "iniSectionGetInteger", "Failed to convert value for key: ""%s"" to integer; error: %s", key, errText
        value = defaultValue
        iniSectionGetInteger = hasDefault
    End If
End Function

Private Sub iniSectionSetInteger(ByVal iniSection As IniSection, ByVal key As String, ByVal value As Long)
    iniSection.SetString key, CStr(value)
End Sub

' C: optional<float> iniSectionGetFloat(iniSection, key, optional<float> defaultValue = nullopt)
Private Function iniSectionGetFloat(ByVal iniSection As IniSection, ByVal key As String, value As Single, Optional ByVal hasDefault As Boolean = False, Optional ByVal defaultValue As Single = 0) As Boolean
    Dim valueStr As String, endIdx As Long, d As Double
    If Not iniSection.HasValue(key) Then
        value = defaultValue
        iniSectionGetFloat = hasDefault
        Exit Function
    End If
    valueStr = iniSection.Get_(key, "")
    d = strtod(valueStr, endIdx)
    If endIdx = 0 Then
        debug_ 143, LOG_ERROR, "iniSectionGetFloat", "Failed to convert value for key: ""%s"" to float; error: %s", key, "invalid stof argument"
        value = defaultValue
        iniSectionGetFloat = hasDefault
    ElseIf Abs(d) > 3.40282346638529E+38 Then
        debug_ 143, LOG_ERROR, "iniSectionGetFloat", "Failed to convert value for key: ""%s"" to float; error: %s", key, "stof argument out of range"
        value = defaultValue
        iniSectionGetFloat = hasDefault
    Else
        value = d
        iniSectionGetFloat = True
    End If
End Function

Private Sub iniSectionSetFloat(ByVal iniSection As IniSection, ByVal key As String, ByVal value As Single)
    iniSection.SetString key, va("%g", CDbl(value))      ' fmt::format("{:g}", value)
End Sub

' C: optional<bool> iniSectionGetBool(iniSection, key, optional<bool> defaultValue = nullopt)
Private Function iniSectionGetBool(ByVal iniSection As IniSection, ByVal key As String, value As Boolean, Optional ByVal hasDefault As Boolean = False, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim valueStr As String, valueInt As Long
    If Not iniSection.HasValue(key) Then
        value = defaultValue
        iniSectionGetBool = hasDefault
        Exit Function
    End If
    valueStr = LCase$(iniSection.Get_(key, ""))
    ' first check if it's equal to "true" or "false" (case-insensitive)
    If valueStr = "true" Then
        value = True
        iniSectionGetBool = True
        Exit Function
    ElseIf valueStr = "false" Then
        value = False
        iniSectionGetBool = True
        Exit Function
    Else
        ' check for 1 or 0
        If iniSectionGetInteger(iniSection, key, valueInt) Then
            If valueInt = 1 Then
                value = True
                iniSectionGetBool = True
                Exit Function
            ElseIf valueInt = 0 Then
                value = False
                iniSectionGetBool = True
                Exit Function
            End If
        End If
    End If
    value = defaultValue
    iniSectionGetBool = hasDefault
End Function

Private Sub iniSectionSetBool(ByVal iniSection As IniSection, ByVal key As String, ByVal value As Boolean)
    If value Then iniSection.SetString key, "true" Else iniSection.SetString key, "false"
End Sub

' C: optional<std::string> iniSectionGetString(iniSection, key, optional<std::string> defaultValue = nullopt)
Private Function iniSectionGetString(ByVal iniSection As IniSection, ByVal key As String, value As String, Optional ByVal hasDefault As Boolean = False, Optional ByVal defaultValue As String = "") As Boolean
    Dim result As String
    If Not iniSection.HasValue(key) Then
        value = defaultValue
        iniSectionGetString = hasDefault
        Exit Function
    End If
    result = iniSection.Get_(key, "")
    ' To support prior INI files written by QSettings, strip surrounding "" if present
    If Len(result) > 0 Then
        If Left$(result, 1) = """" And Right$(result, 1) = """" Then
            If Len(result) <= 2 Then
                value = ""
                iniSectionGetString = True
                Exit Function
            End If
            result = Mid$(result, 2, Len(result) - 2)
        End If
    End If
    value = result
    iniSectionGetString = True
End Function

Public Function saveIniFile(ByVal outputPath As String, ByVal ini As INIReaderWriter) As Boolean
    Dim iniOutput As String, b() As Byte, n As Long
    ' write out ini file changes
    iniOutput = ini.dump()
    n = utf8_encode(iniOutput, b)
    If n = 0 Then ReDim b(0 To 0)
    If Not saveFile(outputPath, b, n) Then
        debug_ 217, LOG_ERROR, "saveIniFile", "Ini write failed"
        Exit Function
    End If
    saveIniFile = True
End Function

Private Function loadConfigIniFile(ByVal inputFile As String) As INIReaderWriter
    Dim fileStatSize As Double, file As Long, filesize As Double, length_read As Long
    Dim fileContentsBuffer() As Byte, bufferSize As Long, result As INIReaderWriter

    If Not WZ_PHYSFS_getFileSize(inputFile, fileStatSize) Then fileStatSize = 0
    If fileStatSize > MAX_CONFIG_FILE_SIZE Then
        debug_ 230, LOG_ERROR, "loadConfigIniFile", "Could not read existing configuration file ""%s""; filesize (%u) exceeds max", inputFile, CDec(fileStatSize)
        Set loadConfigIniFile = New INIReaderWriter
        Exit Function
    End If

    ' load in the existing ini file
    file = PHYSFS_openRead(inputFile)
    If file = 0 Then
        debug_ 241, LOG_WZ, "loadConfigIniFile", "Could not read existing configuration file ""%s""", inputFile
        Set loadConfigIniFile = New INIReaderWriter
        Exit Function
    End If

    ' (from here every way out of the block closes the file: C++'s gsl::finally)
    debug_ 254, LOG_WZ, "loadConfigIniFile", "Reading configuration from: %s", WZ_PHYSFS_getRealPath(inputFile)

    ' get file size from the open file handle
    filesize = PHYSFS_fileLength(file)
    If filesize <= 0 Then
        ' File size could not be determined. Is a directory?
        debug_ 261, LOG_INFO, "loadConfigIniFile", "Could not read existing configuration file ""%s""; filesize returned (%i)", inputFile, CDec(filesize)
        CloseConfigFile file, inputFile
        Set loadConfigIniFile = New INIReaderWriter
        Exit Function
    End If
    If filesize > MAX_CONFIG_FILE_SIZE Then
        debug_ 266, LOG_ERROR, "loadConfigIniFile", "Could not read existing configuration file ""%s""; filesize (%u) exceeds max", inputFile, CDec(filesize)
        CloseConfigFile file, inputFile
        Set loadConfigIniFile = New INIReaderWriter
        Exit Function
    End If
    If Not (filesize < 2147483647#) Then
        ASSERT_FAILURE "filesize < static_cast<PHYSFS_sint64>(std::numeric_limits<PHYSFS_sint32>::max())", "src/configuration.cpp:269", "loadConfigIniFile", """%s"" filesize >= std::numeric_limits<PHYSFS_sint32>::max()", inputFile
        CloseConfigFile file, inputFile
        Set loadConfigIniFile = New INIReaderWriter
        Exit Function
    End If
    If Not (filesize < 4294967295#) Then
        ASSERT_FAILURE "static_cast<PHYSFS_uint64>(filesize) < static_cast<PHYSFS_uint64>(std::numeric_limits<size_t>::max())", "src/configuration.cpp:270", "loadConfigIniFile", """%s"" filesize >= std::numeric_limits<size_t>::max()", inputFile
        CloseConfigFile file, inputFile
        Set loadConfigIniFile = New INIReaderWriter
        Exit Function
    End If

    ReDim fileContentsBuffer(0 To filesize)
    bufferSize = filesize + 1

    ' Load the file data
    length_read = PHYSFS_readBytes(file, fileContentsBuffer, 0, filesize)
    If length_read <> filesize Then
        bufferSize = 0

        debug_ 280, LOG_ERROR, "loadConfigIniFile", "Reading %s short: %s", inputFile, WZ_PHYSFS_getLastError()
        PHYSFS_close file
        CloseConfigFile file, inputFile     ' the finally closes it a second time, as in C++ (and logs the failure)
        Set loadConfigIniFile = New INIReaderWriter
        Exit Function
    End If
    CloseConfigFile file, inputFile

    If bufferSize = 0 Then
        debug_ 288, LOG_INFO, "loadConfigIniFile", "Nothing read from configuration file ""%s""?", inputFile
        Set loadConfigIniFile = New INIReaderWriter
        Exit Function
    End If

    ' append null
    fileContentsBuffer(bufferSize - 1) = 0

    ' load IniReaderWriter from the buffer
    Set result = New INIReaderWriter
    result.InitFromData fileContentsBuffer, bufferSize - 1
    If result.ParseError() <> 0 Then
        debug_ 299, LOG_ERROR, "loadConfigIniFile", "Could not read existing configuration file ""%s""; error parsing line: %d", inputFile, result.ParseError()
        Set loadConfigIniFile = New INIReaderWriter
        Exit Function
    End If
    Set loadConfigIniFile = result
End Function

' loadConfigIniFile's gsl::finally
Private Sub CloseConfigFile(file As Long, ByVal inputFile As String)
    If PHYSFS_close(file) = 0 Then
        debug_ 248, LOG_ERROR, "loadConfigIniFile", "Error closing %s: %s", inputFile, WZ_PHYSFS_getLastError()
        ' but continue on...
    End If
    file = 0
End Sub

'////////////////////////////////////////////////////////////////////////////
' loadConfig's lambdas

Private Function iniGetInteger(ByVal key As String, ByVal defaultValue As Long) As Long
    iniSectionGetInteger iniGeneral, key, iniGetInteger, True, defaultValue
End Function

' takes and returns optional<int>: the float is truncated
Private Function iniGetFloat(ByVal key As String, ByVal defaultValue As Long) As Long
    Dim f As Single
    iniSectionGetFloat iniGeneral, key, f, True, CSng(defaultValue)
    iniGetFloat = ToInt(f)
End Function

Private Function iniGetIntegerOpt(ByVal key As String, value As Long) As Boolean
    iniGetIntegerOpt = iniSectionGetInteger(iniGeneral, key, value)
End Function

Private Function iniGetBool(ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    iniSectionGetBool iniGeneral, key, iniGetBool, True, defaultValue
End Function

Private Function iniGetBoolOpt(ByVal key As String, value As Boolean) As Boolean
    iniGetBoolOpt = iniSectionGetBool(iniGeneral, key, value)
End Function

Private Function iniGetString(ByVal key As String, ByVal defaultValue As String) As String
    iniSectionGetString iniGeneral, key, iniGetString, True, defaultValue
End Function

' optional<MOUSE_KEY_CODE>: 0 = nullopt
Private Function iniGetMouseKeyCode(ByVal key As String, ByVal defaultValue As Long) As Long
    Dim intVal As Long
    If Not iniSectionGetInteger(iniGeneral, key, intVal) Then
        iniGetMouseKeyCode = defaultValue
        Exit Function
    End If
    If intVal = 0 Then
        iniGetMouseKeyCode = 0 ' 0 represents "disabled"
        Exit Function
    End If
    If intVal >= MOUSE_LMB And intVal <= MOUSE_X2 Then ' deliberately exclude mouse MOUSE_WUP + MOUSE_WDN
        iniGetMouseKeyCode = intVal
    Else
        debug_ 354, LOG_WARNING, "loadConfig", "Unsupported / invalid MOUSE_KEY_CODE value: %d", intVal
        iniGetMouseKeyCode = defaultValue
    End If
End Function

Private Function iniGetPlayerLeaveMode(ByVal key As String, ByVal defaultValue As Long) As Long
    Dim intVal As Long
    If Not iniSectionGetInteger(iniGeneral, key, intVal) Then
        iniGetPlayerLeaveMode = defaultValue
        Exit Function
    End If
    If intVal >= 0 And intVal <= PLAYER_LEAVE_MODE_MAX Then
        iniGetPlayerLeaveMode = intVal
    Else
        debug_ 371, LOG_WARNING, "loadConfig", "Unsupported / invalid PLAYER_LEAVE_MODE value: %d; defaulting to: %d", intVal, defaultValue
        iniGetPlayerLeaveMode = defaultValue
    End If
End Function

'////////////////////////////////////////////////////////////////////////////
Public Function loadConfig() As Boolean
    Dim ini As INIReaderWriter, createdConfigFile As Boolean, configVersion As Long
    Dim value As Long, bValue As Boolean, v As Long, hrtfmode_int As Long
    Dim oldMiddleClickRotateValue As Boolean, level As Long
    Dim fullscreenmode_int As Long, oldBoolValue As Boolean, intTrapCursorValue As Long
    Dim displayScale As Long, width As Long, height As Long, screen As Long
    Dim fullscreenWidth As Long, fullscreenHeight As Long, fullscreenPixelDensity As Single
    Dim fullscreenRefreshRate As Single, fullscreenScreen As Long, ignoreOldFullscreenValues As Boolean
    Dim gfxBackend As Long, gfxbackendStr As String, js_backend As Long, jsbackendStr As String
    Dim fogEnabled As Boolean, openSpecSlotsIntValue As Long, intValue As Long, defAI As String
    Dim hostConnProvider As Long, netBackend As String, defConnProviderStr As String
    Dim botProt As Long, proxyIPs As Long, hostingIPs As Long, rotateKey As Long

    Set ini = loadConfigIniFile(fileName)
    createdConfigFile = Not ini.LoadedFromData()

    Set iniGeneral = ini.GetSection("General")

    ActivityManager_beginLoadingSettings

    configVersion = iniGetInteger("configVersion", BASECONFVERSION)

    If iniGetIntegerOpt("voicevol", value) Then
        sound_SetUIVolume CSng(value) / 100!
    End If
    If iniGetIntegerOpt("fxvol", value) Then
        sound_SetEffectsVolume CSng(value) / 100!
    End If
    If iniGetIntegerOpt("cdvol", value) Then
        sound_SetMusicVolume CSng(value) / 100!
    End If
    If iniGetBoolOpt("music_enabled", bValue) Then
        war_SetMusicEnabled bValue
    End If
    If iniGetIntegerOpt("hrtf", value) Then
        hrtfmode_int = value
        If hrtfmode_int >= HRTFMode_Disabled And hrtfmode_int <= HRTFMode_Auto Then   ' MIN_VALID_HRTFMode .. MAX_VALID_HRTFMode
            war_SetHRTFMode hrtfmode_int
        End If
    End If
    If iniGetIntegerOpt("mapZoom", value) Then
        v = value
        If v Mod MAP_ZOOM_CONFIG_STEP <> 0 Then war_SetMapZoom STARTDISTANCE Else war_SetMapZoom v
    End If
    If iniGetIntegerOpt("mapZoomRate", value) Then
        v = value
        If v Mod MAP_ZOOM_RATE_STEP <> 0 Then war_SetMapZoomRate MAP_ZOOM_RATE_DEFAULT Else war_SetMapZoomRate v
    End If
    If iniGetIntegerOpt("radarZoom", value) Then
        v = value
        If v Mod RADARZOOM_STEP <> 0 Then war_SetRadarZoom DEFAULT_RADARZOOM Else war_SetRadarZoom v
    End If
    If iniGeneral.has("language") Then
        setLanguage iniGetString("language", "")
    End If
    If iniGetBoolOpt("nomousewarp", bValue) Then
        setMouseWarp bValue
    End If
    wz_texture_compression = iniGetBool("textureCompression", True)
    showFPS = iniGetBool("showFPS", False)
    showUNITCOUNT = iniGetBool("showUNITCOUNT", False)
    If iniGetIntegerOpt("cameraSpeed", value) Then
        v = value
        If v Mod CAMERASPEED_STEP <> 0 Then war_SetCameraSpeed CAMERASPEED_DEFAULT Else war_SetCameraSpeed v
    End If
    setShakeStatus iniGetBool("shake", False)
    war_setGroupsMenuEnabled iniGetBool("groupmenu", True)
    setGroupButtonEnabled war_getGroupsMenuEnabled()
    war_setOptionsButtonVisibility iniGetInteger("optionsButtonVisibility", war_getOptionsButtonVisibility())
    setCameraAccel iniGetBool("cameraAccel", True)
    setDrawShadows iniGetBool("shadows", True)
    war_setSoundEnabled iniGetBool("sound", True)
    setInvertMouseStatus iniGetBool("mouseflip", True)
    setRightClickOrders iniGetBool("RightClickOrders", False)
    setPanMouseKey iniGetMouseKeyCode("mouseKeyPan", getPanMouseKey())
    If Not createdConfigFile And configVersion < 3 Then
        ' Upgrade old MiddleClickRotate value
        oldMiddleClickRotateValue = iniGetBool("MiddleClickRotate", False)
        If oldMiddleClickRotateValue Then rotateKey = MOUSE_MMB Else rotateKey = MOUSE_RMB
        If Not setRotateMouseKey(rotateKey) And Not oldMiddleClickRotateValue Then
            ' if unable to set (because of conflict, presumably with RightClickOrders), default to MMB
            setRotateMouseKey MOUSE_MMB
        End If
    Else
        setRotateMouseKey iniGetMouseKeyCode("mouseKeyRotate", getRotateMouseKey())
    End If
    setEdgeScrollOutsideWindowBounds iniGetBool("edgeScrollOutsideWindow", getEdgeScrollOutsideWindowBounds())
    If iniGetIntegerOpt("cursorScale", value) Then
        war_setCursorScale value
    End If
    If iniGetBoolOpt("radarJump", bValue) Then
        war_SetRadarJump bValue
    End If
    If iniGetIntegerOpt("lodDistanceBias", value) Then
        war_setLODDistanceBiasPercentage value
    End If
    rotateRadar = iniGetBool("rotateRadar", True)
    radarRotationArrow = iniGetBool("radarRotationArrow", True)
    hostQuitConfirmation = iniGetBool("hostQuitConfirmation", True)
    war_SetPauseOnFocusLoss iniGetBool("PauseOnFocusLoss", False)
    NETsetLobbyserverAddress iniGetString("lobbyserver", netlobby_GetDefaultLobbyAddress())
    war_setLastIpServerConnect iniGetString("server_name", "")
    If Not netGameserverPortOverride Then ' do not load the config port setting if there's a command-line override
        NETsetGameserverPort iniGetInteger("gameserver_port", GAMESERVERPORT)
    End If
    NETsetJoinPreferenceIPv6 iniGetBool("prefer_ipv6", True)
    NETsetDefaultMPHostFreeChatPreference iniGetBool("hostingChatDefault", NETgetDefaultMPHostFreeChatPreference())
    NETsetEnableTCPNoDelay iniGetBool("tcp_nodelay", NETgetEnableTCPNoDelay())
    NETgetLobbyHostJoinOptions botProt, proxyIPs, hostingIPs
    If iniGetIntegerOpt("lobbyHostJoinOpts_botProt", value) Then
        botProt = Clamp(value, HostJoinOptionValue_BlockAll, HostJoinOptionValue_Allow)
    End If
    If iniGetIntegerOpt("lobbyHostJoinOpts_proxyIPs", value) Then
        proxyIPs = Clamp(value, HostJoinOptionValue_BlockAll, HostJoinOptionValue_Allow)
    End If
    If iniGetIntegerOpt("lobbyHostJoinOpts_hostingIPs", value) Then
        hostingIPs = Clamp(value, HostJoinOptionValue_BlockAll, HostJoinOptionValue_Allow)
    End If
    NETsetLobbyHostJoinOptions botProt, proxyIPs, hostingIPs
    setPublicIPv4LookupService iniGetString("publicIPv4LookupService_Url", WZ_DEFAULT_PUBLIC_IPv4_LOOKUP_SERVICE_URL), iniGetString("publicIPv4LookupService_JSONKey", WZ_DEFAULT_PUBLIC_IPv4_LOOKUP_SERVICE_JSONKEY)
    setPublicIPv6LookupService iniGetString("publicIPv6LookupService_Url", WZ_DEFAULT_PUBLIC_IPv6_LOOKUP_SERVICE_URL), iniGetString("publicIPv6LookupService_JSONKey", WZ_DEFAULT_PUBLIC_IPv6_LOOKUP_SERVICE_JSONKEY)
    war_SetFMVmode iniGetInteger("FMVmode", war_GetFMVmode())
    war_setScanlineMode iniGetInteger("scanlines", SCANLINES_OFF)
    seq_SetSubtitles iniGetBool("subtitles", True)
    setDifficultyLevel iniGetInteger("difficulty", DL_NORMAL)
    If Not createdConfigFile And configVersion < 2 Then
        level = getDifficultyLevel() + 1
        If level > AIDifficulty_INSANE Then
            level = AIDifficulty_INSANE
        ElseIf level < AIDifficulty_SUPEREASY Then
            level = AIDifficulty_SUPEREASY
        End If
        setDifficultyLevel level
    End If
    war_SetSPcolor iniGetInteger("colour", 0)  ' default is green (0)
    war_setMPcolour iniGetInteger("colourMP", -1)  ' default is random (-1)
    game.name_ = Left$(iniGetString("gameName", "My Game"), 127)
    sPlayer = Left$(iniGetString("playerName", "Player"), 127)

    ' Set a default map to prevent hosting games without a map.
    game.map = Left$(DEFAULTSKIRMISHMAP, 127)
    Sha256_setZero game.hash
    game.maxPlayers = DEFAULTSKIRMISHMAPMAXPLAYERS

    game.techLevel = 1

    game.power = iniGetInteger("powerLevel", LEV_MED)
    game.base = iniGetInteger("base", CAMP_BASE) And &HFF&
    game.alliance = iniGetInteger("alliance", NO_ALLIANCES) And &HFF&
    game.scavengers = iniGetInteger("newScavengers", SCAVENGERS) And &HFF&
    war_setMPInactivityMinutes iniGetInteger("inactivityMinutesMP", war_getMPInactivityMinutes())
    game.inactivityMinutes = war_getMPInactivityMinutes()
    war_setMPGameTimeLimitMinutes iniGetInteger("gameTimeLimitMinutesMP", war_getMPGameTimeLimitMinutes())
    game.gameTimeLimitMinutes = war_getMPGameTimeLimitMinutes()
    war_setMPPlayerLeaveMode iniGetPlayerLeaveMode("playerLeaveModeMP", war_getMPPlayerLeaveMode())
    game.playerLeaveMode = war_getMPPlayerLeaveMode()
    bEnemyAllyRadarColor = iniGetBool("radarObjectMode", False)
    radarDrawMode = iniGetInteger("radarTerrainMode", RADAR_MODE_DEFAULT)
    If NUM_RADAR_MODES - 1 < radarDrawMode Then radarDrawMode = NUM_RADAR_MODES - 1 ' restrict to allowed values
    If iniGetIntegerOpt("textureSize", value) Then
        setTextureSize value
    End If
    ' First try the new setting called "portMapping", otherwise try the legacy "UPnP" option.
    NetPlay_isPortMappingEnabled = iniGetBool("portMapping", iniGetBool("UPnP", True))
    If iniGetIntegerOpt("antialiasing", value) Then
        war_setAntialiasing value
    End If
    If iniGetIntegerOpt("fullscreen", value) Then
        fullscreenmode_int = value
        ' Handle porting from old fullscreen values
        If fullscreenmode_int <> WINDOW_MODE_windowed Then
            fullscreenmode_int = WINDOW_MODE_fullscreen
        End If
        If fullscreenmode_int >= WINDOW_MODE_windowed And fullscreenmode_int <= WINDOW_MODE_fullscreen Then  ' MIN_VALID_WINDOW_MODE .. MAX_VALID_WINDOW_MODE
            war_setWindowMode fullscreenmode_int
        End If
    End If
    If Not createdConfigFile And configVersion < 3 Then
        ' Upgrade old bool value to TrapCursorMode
        ' - map true -> Enabled and upgrade false -> Automatic
        oldBoolValue = iniGetBool("trapCursor", False)
        If oldBoolValue Then war_SetTrapCursor TrapCursorMode_Enabled Else war_SetTrapCursor TrapCursorMode_Automatic
    Else
        intTrapCursorValue = iniGetInteger("trapCursor", TrapCursorMode_Automatic)
        If intTrapCursorValue < TrapCursorMode_Disabled Or intTrapCursorValue > TrapCursorMode_Automatic Then
            intTrapCursorValue = TrapCursorMode_Automatic
        End If
        war_SetTrapCursor intTrapCursorValue
    End If
    war_SetColouredCursor iniGetBool("coloredCursor", True)
    ' this should be enabled on all systems by default
    war_SetVsync iniGetInteger("vsync", 1)
    ' the default (and minimum) display scale is 100 (%)
    displayScale = iniGetInteger("displayScale", war_GetDisplayScale())
    If displayScale < 100 Then
        displayScale = 100
    End If
    If displayScale > 500 Then ' Maximum supported display scale of 500%. (Further testing needed for anything greater.)
        displayScale = 500
    End If
    war_SetDisplayScale displayScale
    war_setAutoAdjustDisplayScale iniGetBool("autoAdjustDisplayScale", True)
    ' 640x480 is minimum that we will support, but default to something more sensible
    width = iniGetInteger("width", war_GetWidth())
    height = iniGetInteger("height", war_GetHeight())
    screen = iniGetInteger("screen", 0)
    If width < 640 Or height < 480 Then  ' sanity check
        width = 640
        height = 480
    End If
    pie_SetVideoBufferWidth width
    pie_SetVideoBufferHeight height
    war_SetWidth width
    war_SetHeight height
    war_SetScreen screen

    fullscreenWidth = iniGetInteger("fullscreenWidth", war_GetFullscreenModeWidth())
    fullscreenHeight = iniGetInteger("fullscreenHeight", war_GetFullscreenModeHeight())
    fullscreenPixelDensity = iniGetFloat("fullscreenPixelDensity", ToInt(war_GetFullscreenModePixelDensity()))
    fullscreenRefreshRate = iniGetFloat("fullscreenRefreshRate", ToInt(war_GetFullscreenModeRefreshRate()))
    fullscreenScreen = iniGetInteger("fullscreenScreen", 0)
    ignoreOldFullscreenValues = (Not createdConfigFile And configVersion <= 3)
    If ignoreOldFullscreenValues Or (fullscreenWidth <> 0 And fullscreenWidth < 640) Or (fullscreenHeight <> 0 And fullscreenHeight < 480) Then  ' sanity check
        ' set to special value (0x0) that reverts to the default for this display
        fullscreenWidth = 0
        fullscreenHeight = 0
        fullscreenRefreshRate = 0!
    End If
    war_SetFullscreenModeWidth fullscreenWidth
    war_SetFullscreenModeHeight fullscreenHeight
    war_SetFullscreenModeScreen fullscreenScreen
    war_SetFullscreenModePixelDensity fullscreenPixelDensity
    war_SetFullscreenModeRefreshRate fullscreenRefreshRate

    If iniGetIntegerOpt("bpp", value) Then
        war_SetVideoBufferDepth value
    End If

    If iniGeneral.has("gfxbackend") Then
        gfxbackendStr = iniGetString("gfxbackend", "")
        If Not video_backend_from_str(gfxbackendStr, gfxBackend) Then
            gfxBackend = wzGetDefaultGfxBackendForCurrentSystem()
            debug_ 648, LOG_WARNING, "loadConfig", "Unsupported / invalid gfxbackend value: %s; defaulting to: %s", gfxbackendStr, to_string_video_backend(gfxBackend)
        End If
    Else
        gfxBackend = wzGetDefaultGfxBackendForCurrentSystem()
    End If
    war_setGfxBackend gfxBackend
    If iniGetIntegerOpt("minimizeOnFocusLoss", value) Then
        war_setMinimizeOnFocusLoss value
    End If
    If iniGetIntegerOpt("altEnterToggleMode", value) Then
        war_setToggleFullscreenMode value
    End If
    js_backend = war_getJSBackend()
    If iniGeneral.has("jsbackend") Then
        jsbackendStr = iniGetString("jsbackend", "")
        If Not js_backend_from_str(jsbackendStr, js_backend) Then
            js_backend = 0 ' use the first available option, whatever it is
            debug_ 671, LOG_WARNING, "loadConfig", "Unsupported / invalid jsbackend value: %s; defaulting to: %s", jsbackendStr, to_string_JS_BACKEND(js_backend)
        End If
    End If
    BlueprintTrackAnimationSpeed = iniGetInteger("BlueprintTrackAnimationSpeed", 20)
    lockCameraScrollWhileRotating = iniGetBool("lockCameraScrollWhileRotating", False)
    autosaveEnabled = iniGetBool("autosaveEnabled", True)
    fogEnabled = iniGetBool("fog", False)
    If fogEnabled Then
        pie_EnableFog True
    Else
        pie_SetFogStatus False
        pie_EnableFog False
    End If
    war_setAutoLagKickSeconds iniGetInteger("hostAutoLagKickSeconds", war_getAutoLagKickSeconds())
    war_setAutoLagKickAggressiveness iniGetInteger("hostAutoLagKickAggressiveness", war_getAutoLagKickAggressiveness())
    war_setAutoDesyncKickSeconds iniGetInteger("hostAutoDesyncKickSeconds", war_getAutoDesyncKickSeconds())
    war_setAutoNotReadyKickSeconds iniGetInteger("hostAutoNotReadyKickSeconds", war_getAutoNotReadyKickSeconds())
    war_setDisableReplayRecording iniGetBool("disableReplayRecord", war_getDisableReplayRecording())
    war_setMaxReplaysSaved iniGetInteger("maxReplaysSaved", war_getMaxReplaysSaved())
    war_setOldLogsLimit iniGetInteger("oldLogsLimit", war_getOldLogsLimit())
    openSpecSlotsIntValue = iniGetInteger("openSpectatorSlotsMP", war_getMPopenSpectatorSlots())
    war_setMPopenSpectatorSlots Clamp(openSpecSlotsIntValue, 0, MAX_SPECTATOR_SLOTS) And &HFFFF&
    war_setFogEnd iniGetInteger("fogEnd", 8000)
    war_setFogStart iniGetInteger("fogStart", 4000)
    If iniGetIntegerOpt("terrainMode", value) Then
        intValue = value
        If intValue >= 0 And intValue <= TerrainShaderQuality_MAX Then
            setTerrainShaderQuality intValue
        Else
            debug_ 707, LOG_WARNING, "loadConfig", "Unsupported / invalid terrainMode value: %d; using default", intValue
        End If
    End If
    If iniGetIntegerOpt("terrainShadingQuality", value) Then
        intValue = value
        If Not setTerrainMappingTexturesMaxSize(intValue) Then
            debug_ 715, LOG_WARNING, "loadConfig", "Unsupported / invalid terrainShadingQuality value: %d; using default", intValue
        End If
    End If
    setDrawTerrainShadows iniGetBool("terrainShadows", True)
    war_setShadowFilterSize iniGetInteger("shadowFilterSize", war_getShadowFilterSize())
    If iniGetIntegerOpt("shadowMapResolution", value) Then
        war_setShadowMapResolution value
    End If

    If iniGetBoolOpt("pointLightsPerpixel", bValue) Then
        war_setPointLightPerPixelLighting bValue
    Else
        war_setPointLightPerPixelLighting False
    End If

    defAI = iniGetString("defaultSkirmishAI", DEFAULT_SKIRMISH_AI_SCRIPT_NAME)
    setDefaultSkirmishAI defAI

    war_setPlayAudioCue_GroupReporting iniGetBool("audioCueGroupReporting", war_getPlayAudioCue_GroupReporting())

    hostConnProvider = war_getHostConnectionProvider()
    If iniGeneral.has("hostConnectionProvider") Then
        netBackend = iniGetString("hostConnectionProvider", "tcp")
        If Not net_backend_from_str(netBackend, hostConnProvider) Then
            hostConnProvider = ConnectionProviderType_TCP_DIRECT ' fall back to using TCP_DIRECT
            defConnProviderStr = to_string_ConnectionProviderType(hostConnProvider)
            debug_ 743, LOG_WARNING, "loadConfig", "Unsupported / invalid network backend: %s; defaulting to: %s", _
                netBackend, defConnProviderStr
        End If
        war_setHostConnectionProvider hostConnProvider
    End If

    war_setLobbyDisableIPv6 iniGetBool("lobbyDisableIPv6", war_getLobbyDisableIPv6())
    war_setLobbyFilterIPv6Only iniGetBool("lobbyFilterIPv6Only", war_getLobbyFilterIPv6Only())

    ActivityManager_endLoadingSettings
    Set iniGeneral = Nothing
    loadConfig = True
End Function

' C: std::clamp<int>
Private Function Clamp(ByVal v As Long, ByVal lo As Long, ByVal hi As Long) As Long
    If v < lo Then
        Clamp = lo
    ElseIf v > hi Then
        Clamp = hi
    Else
        Clamp = v
    End If
End Function

' saveConfig's iniSetFromCString: the text up to its terminator, at most maxLength - 1 characters
Private Sub iniSetFromCString(ByVal iniSection As IniSection, ByVal key As String, ByVal value As String, ByVal maxLength As Long)
    Dim strVal As String, length_ As Long
    length_ = InStr(value, vbNullChar) - 1
    If length_ < 0 Then length_ = Len(value)
    If length_ > maxLength Then length_ = maxLength
    If Not (length_ < maxLength) Then
        ASSERT_FAILURE "len < maxLength", "src/configuration.cpp:789", "saveConfig", "Input c-string value (for key: %s) appears to be missing null-terminator?", key
    End If
    strVal = Left$(value, length_)
    iniSection.SetString key, strVal
End Sub

' saveConfig's iniSetMouseKeyOpt: 0 for nullopt
Private Sub iniSetMouseKeyOpt(ByVal iniSection As IniSection, ByVal key As String, ByVal code As Long)
    iniSectionSetInteger iniSection, key, code
End Sub

'////////////////////////////////////////////////////////////////////////////
Public Function saveConfig() As Boolean
    Dim ini As INIReaderWriter, fullConfigFilePath As String, g As IniSection, currentSpectatorSlotInfo As Long

    Set ini = loadConfigIniFile(fileName)

    If StrPtr(PHYSFS_getWriteDir()) <> 0 Then
        fullConfigFilePath = fullConfigFilePath & PHYSFS_getWriteDir()
        fullConfigFilePath = fullConfigFilePath & "/"
    End If
    fullConfigFilePath = fullConfigFilePath & fileName
    debug_ 768, LOG_WZ, "saveConfig", "Writing configuration to: ""%s""", fullConfigFilePath

    Set g = ini.GetSection("General")

    '//////////////////////////
    ' voicevol, fxvol and cdvol
    iniSectionSetInteger g, "voicevol", ToInt(CDbl(sound_GetUIVolume()) * 100#)
    iniSectionSetInteger g, "fxvol", ToInt(CDbl(sound_GetEffectsVolume()) * 100#)
    iniSectionSetInteger g, "cdvol", ToInt(CDbl(sound_GetMusicVolume()) * 100#)
    iniSectionSetBool g, "music_enabled", war_GetMusicEnabled()
    iniSectionSetInteger g, "hrtf", war_GetHRTFMode()
    iniSectionSetInteger g, "mapZoom", war_GetMapZoom()
    iniSectionSetInteger g, "mapZoomRate", war_GetMapZoomRate()
    iniSectionSetInteger g, "radarZoom", war_GetRadarZoom()
    iniSectionSetInteger g, "width", war_GetWidth()
    iniSectionSetInteger g, "height", war_GetHeight()
    iniSectionSetInteger g, "screen", war_GetScreen()
    iniSectionSetInteger g, "fullscreenWidth", war_GetFullscreenModeWidth()
    iniSectionSetInteger g, "fullscreenHeight", war_GetFullscreenModeHeight()
    iniSectionSetInteger g, "fullscreenScreen", war_GetFullscreenModeScreen()
    iniSectionSetFloat g, "fullscreenPixelDensity", war_GetFullscreenModePixelDensity()
    iniSectionSetFloat g, "fullscreenRefreshRate", war_GetFullscreenModeRefreshRate()
    iniSectionSetInteger g, "bpp", war_GetVideoBufferDepth()
    iniSectionSetInteger g, "fullscreen", war_getWindowMode()
    g.SetString "language", getLanguage()
    iniSectionSetInteger g, "difficulty", getDifficultyLevel()      ' level
    iniSectionSetInteger g, "cameraSpeed", war_GetCameraSpeed()    ' camera speed
    iniSectionSetBool g, "radarJump", war_GetRadarJump()        ' radar jump
    iniSectionSetInteger g, "lodDistanceBias", war_getLODDistanceBiasPercentage()
    iniSectionSetBool g, "cameraAccel", getCameraAccel()        ' camera acceleration
    iniSectionSetInteger g, "shake", -CLng(getShakeStatus())        ' screenshake
    iniSectionSetInteger g, "groupmenu", -CLng(war_getGroupsMenuEnabled())        ' groups menu
    iniSectionSetInteger g, "optionsButtonVisibility", war_getOptionsButtonVisibility()
    iniSectionSetInteger g, "mouseflip", -CLng(getInvertMouseStatus())   ' flipmouse
    iniSectionSetInteger g, "nomousewarp", -CLng(getMouseWarp())        ' mouse warp
    iniSectionSetInteger g, "coloredCursor", -CLng(war_GetColouredCursor())
    iniSectionSetInteger g, "RightClickOrders", -CLng(getRightClickOrders())
    iniSetMouseKeyOpt g, "mouseKeyPan", getPanMouseKey()
    iniSetMouseKeyOpt g, "mouseKeyRotate", getRotateMouseKey()
    iniSectionSetInteger g, "edgeScrollOutsideWindow", -CLng(getEdgeScrollOutsideWindowBounds())
    iniSectionSetInteger g, "cursorScale", war_getCursorScale()
    If wz_texture_compression Then iniSectionSetInteger g, "textureCompression", 1 Else iniSectionSetInteger g, "textureCompression", 0
    iniSectionSetInteger g, "showFPS", -CLng(showFPS)
    iniSectionSetInteger g, "showUNITCOUNT", -CLng(showUNITCOUNT)
    iniSectionSetInteger g, "shadows", -CLng(getDrawShadows())  ' shadows
    iniSectionSetInteger g, "sound", -CLng(war_getSoundEnabled())
    iniSectionSetInteger g, "FMVmode", war_GetFMVmode()     ' sequences
    iniSectionSetInteger g, "scanlines", war_getScanlineMode()
    iniSectionSetInteger g, "subtitles", -CLng(seq_GetSubtitles())      ' subtitles
    iniSectionSetInteger g, "radarObjectMode", -CLng(bEnemyAllyRadarColor)   ' enemy/allies radar view
    iniSectionSetInteger g, "radarTerrainMode", radarDrawMode
    iniSectionSetInteger g, "trapCursor", war_GetTrapCursor()
    iniSectionSetInteger g, "vsync", war_GetVsync()
    iniSectionSetInteger g, "displayScale", war_GetDisplayScale()
    iniSectionSetBool g, "autoAdjustDisplayScale", war_getAutoAdjustDisplayScale()
    iniSectionSetInteger g, "textureSize", getTextureSize()
    iniSectionSetInteger g, "antialiasing", war_getAntialiasing()
    iniSectionSetInteger g, "portMapping", -CLng(NetPlay_isPortMappingEnabled)
    iniSectionSetBool g, "rotateRadar", rotateRadar
    iniSectionSetBool g, "radarRotationArrow", radarRotationArrow
    iniSectionSetBool g, "hostQuitConfirmation", hostQuitConfirmation
    iniSectionSetBool g, "PauseOnFocusLoss", war_GetPauseOnFocusLoss()
    g.SetString "lobbyserver", NETgetLobbyserverAddress()
    g.SetString "server_name", war_getLastIpServerConnect()
    If Not netGameserverPortOverride Then ' do not save the config port setting if there's a command-line override
        iniSectionSetInteger g, "gameserver_port", NETgetGameserverPort()
    End If
    iniSectionSetBool g, "prefer_ipv6", NETgetJoinPreferenceIPv6()
    If NETgetDefaultMPHostFreeChatPreference() Then iniSectionSetInteger g, "hostingChatDefault", 1 Else iniSectionSetInteger g, "hostingChatDefault", 0
    If NETgetEnableTCPNoDelay() Then iniSectionSetInteger g, "tcp_nodelay", 1 Else iniSectionSetInteger g, "tcp_nodelay", 0

    g.SetString "publicIPv4LookupService_Url", getPublicIPv4LookupServiceUrl()
    g.SetString "publicIPv4LookupService_JSONKey", getPublicIPv4LookupServiceJSONKey()
    g.SetString "publicIPv6LookupService_Url", getPublicIPv6LookupServiceUrl()
    g.SetString "publicIPv6LookupService_JSONKey", getPublicIPv6LookupServiceJSONKey()
    If Not bMultiPlayer Then
        iniSectionSetInteger g, "colour", war_GetSPcolor()         ' favourite colour.
    Else
        If NetPlay_isHost And ingame.localJoiningInProgress Then
            If bMultiPlayer And NetPlay_bComms Then
                iniSetFromCString g, "gameName", game.name_, 128        '  last hosted game
                war_setMPInactivityMinutes game.inactivityMinutes
                war_setMPGameTimeLimitMinutes game.gameTimeLimitMinutes
                war_setMPPlayerLeaveMode game.playerLeaveMode

                ' remember number of spectator slots in MP games
                currentSpectatorSlotInfo = SpectatorInfo_currentNetPlayState()   ' SpectatorInfo::toUint32()
                war_setMPopenSpectatorSlots currentSpectatorSlotInfo And &HFFFF&     ' .totalSpectatorSlots
            End If
            iniSetFromCString g, "mapName", game.map, 128              '  map name
            g.SetString "mapHash", Sha256_toString(game.hash)          '  map hash
            iniSectionSetInteger g, "maxPlayers", game.maxPlayers      ' maxPlayers
            iniSectionSetInteger g, "powerLevel", game.power               ' power
            iniSectionSetInteger g, "base", game.base               ' size of base
            iniSectionSetInteger g, "alliance", game.alliance       ' allow alliances
            iniSectionSetInteger g, "newScavengers", game.scavengers
        End If
        iniSetFromCString g, "playerName", sPlayer, 128      ' player name
    End If
    iniSectionSetInteger g, "colourMP", war_getMPcolour()
    iniSectionSetInteger g, "inactivityMinutesMP", war_getMPInactivityMinutes()
    iniSectionSetInteger g, "gameTimeLimitMinutesMP", war_getMPGameTimeLimitMinutes()
    iniSectionSetInteger g, "playerLeaveModeMP", war_getMPPlayerLeaveMode()
    iniSectionSetInteger g, "openSpectatorSlotsMP", war_getMPopenSpectatorSlots()
    g.SetString "gfxbackend", to_string_video_backend(war_getGfxBackend())
    iniSectionSetInteger g, "minimizeOnFocusLoss", war_getMinimizeOnFocusLoss()
    iniSectionSetInteger g, "altEnterToggleMode", war_getToggleFullscreenMode()
    g.SetString "jsbackend", to_string_JS_BACKEND(war_getJSBackend())
    iniSectionSetInteger g, "BlueprintTrackAnimationSpeed", BlueprintTrackAnimationSpeed
    iniSectionSetBool g, "lockCameraScrollWhileRotating", lockCameraScrollWhileRotating
    iniSectionSetBool g, "autosaveEnabled", autosaveEnabled
    iniSectionSetBool g, "fog", pie_GetFogEnabled()
    iniSectionSetInteger g, "hostAutoLagKickSeconds", war_getAutoLagKickSeconds()
    iniSectionSetInteger g, "hostAutoDesyncKickSeconds", war_getAutoDesyncKickSeconds()
    iniSectionSetInteger g, "hostAutoNotReadyKickSeconds", war_getAutoNotReadyKickSeconds()
    iniSectionSetBool g, "disableReplayRecord", war_getDisableReplayRecording()
    iniSectionSetInteger g, "maxReplaysSaved", war_getMaxReplaysSaved()
    iniSectionSetInteger g, "oldLogsLimit", war_getOldLogsLimit()
    iniSectionSetInteger g, "fogEnd", war_getFogEnd()
    iniSectionSetInteger g, "fogStart", war_getFogStart()
    iniSectionSetInteger g, "terrainMode", getTerrainShaderQuality()
    iniSectionSetInteger g, "terrainShadingQuality", getTerrainMappingTexturesMaxSize()
    iniSectionSetInteger g, "terrainShadows", -CLng(getDrawTerrainShadows())
    iniSectionSetInteger g, "shadowFilterSize", war_getShadowFilterSize()
    iniSectionSetInteger g, "shadowMapResolution", war_getShadowMapResolution()
    iniSectionSetBool g, "pointLightsPerpixel", war_getPointLightPerPixelLighting()
    g.SetString "defaultSkirmishAI", getDefaultSkirmishAI()
    iniSectionSetBool g, "audioCueGroupReporting", war_getPlayAudioCue_GroupReporting()

    iniSectionSetBool g, "lobbyDisableIPv6", war_getLobbyDisableIPv6()
    iniSectionSetBool g, "lobbyFilterIPv6Only", war_getLobbyFilterIPv6Only()

    iniSectionSetInteger g, "configVersion", CURRCONFVERSION

    ' write out ini file changes
    saveConfig = saveIniFile(fileName, ini)
End Function

Public Function saveGfxConfig() As Boolean
    Dim ini As INIReaderWriter, fullConfigFilePath As String, g As IniSection

    Set ini = loadConfigIniFile(fileName)

    If StrPtr(PHYSFS_getWriteDir()) <> 0 Then
        fullConfigFilePath = fullConfigFilePath & PHYSFS_getWriteDir()
        fullConfigFilePath = fullConfigFilePath & "/"
    End If
    fullConfigFilePath = fullConfigFilePath & fileName
    debug_ 955, LOG_WZ, "saveGfxConfig", "Writing gfx configuration to: ""%s""", fullConfigFilePath

    Set g = ini.GetSection("General")

    ' only change specific gfx entries
    g.SetString "gfxbackend", to_string_video_backend(war_getGfxBackend())
    iniSectionSetInteger g, "antialiasing", war_getAntialiasing()

    ' write out ini file changes
    saveGfxConfig = saveIniFile(fileName, ini)
End Function

' Saves and loads the relevant part of the config files for MP games
' Ensures that others' games don't change our own configuration settings
Public Function reloadMPConfig() As Boolean
    Dim ini As INIReaderWriter, g As IniSection, s As String, v As Long

    Set ini = loadConfigIniFile(fileName)
    If Not ini.LoadedFromData() Then
        debug_ 979, LOG_INFO, "reloadMPConfig", "Could not read existing configuration file ""%s""", fileName
    End If

    Set g = ini.GetSection("General")

    debug_ 984, LOG_WZ, "reloadMPConfig", "Reloading mp config"

    ' If we're in-game, we already have our own configuration set, so no need to reload it.
    If NetPlay_isHost And Not ingame.localJoiningInProgress Then
        If bMultiPlayer And Not NetPlay_bComms Then
            ' one-player skirmish mode sets game name to "One Player Skirmish", so
            ' reset the name
            If g.has("gameName") Then
                iniSectionGetString g, "gameName", s, True, ""
                game.name_ = Left$(s, 127)
            End If
        End If
        reloadMPConfig = True
        Exit Function
    End If

    ' If we're host and not in-game, we can safely save our settings and return.
    If NetPlay_isHost And ingame.localJoiningInProgress Then
        If bMultiPlayer And NetPlay_bComms Then
            ReloadSetFromCString g, "gameName", game.name_, 128          '  last hosted game
        Else
            ' One-player skirmish mode sets game name to "One Player Skirmish", so
            ' reset the name
            If g.has("gameName") Then
                iniSectionGetString g, "gameName", s, True, ""
                game.name_ = Left$(s, 127)
            End If
        End If

        ' Set a default map to prevent hosting games without a map.
        game.map = Left$(DEFAULTSKIRMISHMAP, 127)
        Sha256_setZero game.hash
        game.maxPlayers = DEFAULTSKIRMISHMAPMAXPLAYERS

        iniSectionSetInteger g, "powerLevel", game.power               ' power
        iniSectionSetInteger g, "base", game.base               ' size of base
        iniSectionSetInteger g, "alliance", game.alliance       ' allow alliances

        ' write out ini file changes
        reloadMPConfig = saveIniFile(fileName, ini)
        Exit Function
    End If

    ' We're not host, so let's get rid of the host's game settings and restore our own.

    ' game name
    If g.has("gameName") Then
        iniSectionGetString g, "gameName", s, True, ""
        game.name_ = Left$(s, 127)
    End If

    ' Set a default map to prevent hosting games without a map.
    game.map = Left$(DEFAULTSKIRMISHMAP, 127)
    Sha256_setZero game.hash
    game.maxPlayers = DEFAULTSKIRMISHMAPMAXPLAYERS

    iniSectionGetInteger g, "powerLevel", v, True, LEV_MED
    game.power = v
    iniSectionGetInteger g, "base", v, True, CAMP_BASE
    game.base = v And &HFF&
    iniSectionGetInteger g, "alliance", v, True, NO_ALLIANCES
    game.alliance = v And &HFF&
    game.inactivityMinutes = war_getMPInactivityMinutes()
    game.gameTimeLimitMinutes = war_getMPGameTimeLimitMinutes()
    game.playerLeaveMode = war_getMPPlayerLeaveMode()
    game.blindMode = BLIND_MODE_NONE

    ' restore group menus enabled setting (as tutorial may override it)
    setGroupButtonEnabled war_getGroupsMenuEnabled()

    reloadMPConfig = True
End Function

' reloadMPConfig's iniSetFromCString lambda (its ASSERT is at line 991)
Private Sub ReloadSetFromCString(ByVal iniSection As IniSection, ByVal key As String, ByVal value As String, ByVal maxLength As Long)
    Dim strVal As String, length_ As Long
    length_ = InStr(value, vbNullChar) - 1
    If length_ < 0 Then length_ = Len(value)
    If length_ > maxLength Then length_ = maxLength
    If Not (length_ < maxLength) Then
        ASSERT_FAILURE "len < maxLength", "src/configuration.cpp:991", "reloadMPConfig", "Input c-string value (for key: %s) appears to be missing null-terminator?", key
    End If
    strVal = Left$(value, length_)
    iniSection.SetString key, strVal
End Sub
