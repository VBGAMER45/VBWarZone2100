Attribute VB_Name = "interfaceoptions_c"
' src/titleui/options/interfaceoptions.cpp: the Interface Options form (makeInterfaceOptionsForm),
' whole. (VB) std::shared_ptr<IMAGEFILE> flagsImagesCopy is the image file's handle (0 = nullptr);
' getLocales() fills arrays of the codes and names.
Option Explicit

Public Const FN_interface_language_populate As Long = 4450  ' makeInterfaceOptionsForm, "interface.language": [flagsImagesCopy]()
Public Const FN_interface_language_select As Long = 4451    ' [languageDidChangeHandler](const auto& newValue) -> bool
Public Const FN_interface_displayscale_populate As Long = 4452 ' "interface.displayscale": []()
Public Const FN_interface_displayscale_select As Long = 4453 ' [weakOptionsForm](const auto& newValue) -> bool
Public Const FN_interface_displayscale_task As Long = 4454  ' widgScheduleTask([weakOptionsForm]() {...})
Public Const FN_interface_showFPS_populate As Long = 4455   ' "interface.game.showFPS"
Public Const FN_interface_showFPS_select As Long = 4456
Public Const FN_interface_showCounts_populate As Long = 4457 ' "interface.game.showCounts"
Public Const FN_interface_showCounts_select As Long = 4458
Public Const FN_interface_groupsMenu_populate As Long = 4459 ' "interface.game.groupsMenu"
Public Const FN_interface_groupsMenu_select As Long = 4460
Public Const FN_interface_optionsButton_populate As Long = 4461 ' "interface.game.optionsButton"
Public Const FN_interface_optionsButton_select As Long = 4462 ' [inGame](const auto& newValue) -> bool
Public Const FN_interface_rotateRadar_populate As Long = 4463 ' "interface.game.rotateRadar"
Public Const FN_interface_rotateRadar_select As Long = 4464
Public Const FN_interface_radarJump_populate As Long = 4465 ' "interface.game.radarJump"
Public Const FN_interface_radarJump_select As Long = 4466
Public Const FN_interface_radarZoom_get As Long = 4467      ' "interface.game.radarZoom"
Public Const FN_interface_radarZoom_set As Long = 4468
Public Const FN_interface_videoSize_populate As Long = 4469 ' "interface.video.videoSize"
Public Const FN_interface_videoSize_select As Long = 4470
Public Const FN_interface_scanlines_populate As Long = 4471 ' "interface.video.scanlines"
Public Const FN_interface_scanlines_select As Long = 4472
Public Const FN_interface_subtitles_populate As Long = 4473 ' "interface.video.subtitles"
Public Const FN_interface_subtitles_select As Long = 4474

' MARK: -

' static std::map<WzString, const char *> languageIconsMap(): (VB) iconsMap.find(code): True and the
' icon file name when the code is in the map
Private Function languageIconsMap_find(ByVal code As String, icon As String) As Boolean
    Select Case code
    Case "": icon = "icon-system.png"
    Case "ar_SA": icon = "lang-AR.png"
    Case "bg": icon = "flag-BG.png"
    Case "ca_ES": icon = "flag-ES.png"
    Case "cs_CZ": icon = "flag-CZ.png"
    Case "da": icon = "flag-DK.png"
    Case "de": icon = "flag-DE.png"
    Case "el_GR": icon = "flag-GR.png"
    Case "en": icon = "flag-US.png"
    Case "en_GB": icon = "flag-GB.png"
    Case "es": icon = "flag-ES.png"
    Case "et_EE": icon = "flag-EE.png"
    Case "eu_ES": icon = "flag-ES.png"
    Case "fi": icon = "flag-FI.png"
    Case "fr": icon = "flag-FR.png"
    Case "fy_NL": icon = "flag-NL.png"
    Case "ga_IE": icon = "flag-IE.png"
    Case "hr": icon = "flag-HR.png"
    Case "hu": icon = "flag-HU.png"
    Case "id": icon = "flag-ID.png"
    Case "it": icon = "flag-IT.png"
    Case "ko_KR": icon = "flag-KR.png"
    Case "la": icon = "flag-VA.png"
    Case "lt": icon = "flag-LT.png"
    Case "lv": icon = "flag-LV.png"
    Case "nb_NO": icon = "flag-NO.png"
    Case "nn_NO": icon = "flag-NO.png"
    Case "nl": icon = "flag-NL.png"
    Case "pl": icon = "flag-PL.png"
    Case "pt_BR": icon = "flag-BR.png"
    Case "pt": icon = "flag-PT.png"
    Case "ro": icon = "flag-RO.png"
    Case "ru": icon = "flag-RU.png"
    Case "sk": icon = "flag-SK.png"
    Case "sl_SI": icon = "flag-SI.png"
    Case "sr_RS": icon = "flag-RS.png"
    Case "sv_SE": icon = "flag-SE.png"
    Case "sv": icon = "flag-SE.png"
    Case "tr": icon = "flag-TR.png"
    Case "uz": icon = "flag-UZ.png"
    Case "uk_UA": icon = "flag-UA.png"
    Case "zh_CN": icon = "flag-CN.png"
    Case "zh_TW": icon = "flag-TW.png"
    Case Else
        languageIconsMap_find = False
        Exit Function
    End Select
    languageIconsMap_find = True
End Function

' std::shared_ptr<OptionsForm> makeInterfaceOptionsForm(bool inGame, const std::function<void()> languageDidChangeHandler = nullptr)
Public Function makeInterfaceOptionsForm(ByVal inGame_ As Boolean, Optional ByVal languageDidChangeHandler As WzFunc = Nothing) As WIDGET
    Dim result As WIDGET, optionInfo As OptionInfo, valueChanger As WIDGET, flagsImagesCopy As Long, weakOptionsForm As Long
    Set result = OptionsForm_make()
    OptionsForm_setRefreshOptionsOnScreenSizeChange result, True

    OptionsForm_addSection result, OptionsSection_("General", ""), True
    Set optionInfo = OptionInfo_new("interface.language", "Language", "")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_IsNotInGame)
    ' NOTE: By capturing a copy of pFlagsImages in the populate handler, we ensure it stays around until the OptionsDropdown is deleted
    If GetGameMode() <> GS_NORMAL Then flagsImagesCopy = getFlagsImages() Else flagsImagesCopy = 0
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_language_populate, flagsImagesCopy), WzFunc_new(FN_interface_language_select, languageDidChangeHandler), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    weakOptionsForm = OptionsForm_(result).weakSerial
    Set optionInfo = OptionInfo_new("interface.displayscale", "Display Scale", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_displayscale_populate), WzFunc_new(FN_interface_displayscale_select, weakOptionsForm), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    OptionsForm_addSection result, OptionsSection_("Game UI", ""), True
    Set optionInfo = OptionInfo_new("interface.game.showFPS", "Show FPS", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_showFPS_populate), WzFunc_new(FN_interface_showFPS_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("interface.game.showCounts", "Show Unit Counts", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_showCounts_populate), WzFunc_new(FN_interface_showCounts_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("interface.game.groupsMenu", "Groups Menu", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_groupsMenu_populate), WzFunc_new(FN_interface_groupsMenu_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("interface.game.optionsButton", "Options Button", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_optionsButton_populate), WzFunc_new(FN_interface_optionsButton_select, inGame_), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    OptionsForm_addSection result, OptionsSection_("Radar", ""), True
    Set optionInfo = OptionInfo_new("interface.game.rotateRadar", "Radar Orientation", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_rotateRadar_populate), WzFunc_new(FN_interface_rotateRadar_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("interface.game.radarJump", "Radar Jump", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_radarJump_populate), WzFunc_new(FN_interface_radarJump_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("interface.game.radarZoom", "Radar Zoom", "")
    Set valueChanger = OptionsSlider_make(MIN_RADARZOOM, MAX_RADARZOOM, RADARZOOM_STEP, WzFunc_new(FN_interface_radarZoom_get), WzFunc_new(FN_interface_radarZoom_set), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    OptionsForm_addSection result, OptionsSection_("Video Playback", ""), True
    Set optionInfo = OptionInfo_new("interface.video.videoSize", "Video Size", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_videoSize_populate), WzFunc_new(FN_interface_videoSize_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("interface.video.scanlines", "Scanlines", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_scanlines_populate), WzFunc_new(FN_interface_scanlines_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("interface.video.subtitles", "Subtitles", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_interface_subtitles_populate), WzFunc_new(FN_interface_subtitles_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set makeInterfaceOptionsForm = result
End Function

' MARK: - the lambdas

Public Function interfaceoptions_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim result As OptionChoices, bIsInGame As Boolean, locales_code() As String, locales_name() As String, localesCount As Long
    Dim i As Long, icon As Long, iconFile As String, flagsImagesCopy As Long, currentCode As String, languageDidChangeHandler As WzFunc
    Dim displayScales() As Long, displayScalesCount As Long, maximumDisplayScale As Long, displayScale As Long, currValue As Long
    Dim newValue As Long, strongOptionsForm As WIDGET
    Select Case f.id
    Case FN_interface_language_populate         ' [flagsImagesCopy]()
        flagsImagesCopy = f.cap(0)
        bIsInGame = (GetGameMode() = GS_NORMAL)
        If bIsInGame Then
            ' shortcut to avoid loading and processing lots of strings - just populate with the currrent language (can't be changed in-game anyway)
            Set result = New OptionChoices
            result.addChoice getLanguageName(), "", getLanguage(), False
            result.currentIdx = 0
            Set interfaceoptions_c_fn = result
            Exit Function
        End If
        localesCount = getLocales(locales_code, locales_name)
        Set result = New OptionChoices
        For i = 0 To localesCount - 1
            icon = 0
            If languageIconsMap_find(locales_code(i), iconFile) Then
                If flagsImagesCopy <> 0 Then icon = IMAGEFILE_find(flagsImagesCopy, iconFile) Else icon = 0
            End If
            result.addChoice locales_name(i), "", locales_code(i), False, icon
        Next
        currentCode = getLanguage()
        result.setCurrentIdxForValue currentCode
        Set interfaceoptions_c_fn = result

    Case FN_interface_language_select           ' [languageDidChangeHandler](const auto& newValue) -> bool
        If Not setLanguage(CStr(args(0))) Then
            interfaceoptions_c_fn = False
            Exit Function
        End If
        ' hack to update translations of AI names and tooltips
        readAIs
        ' call the language change handler
        If IsObject(f.cap(0)) Then Set languageDidChangeHandler = f.cap(0)
        If Not languageDidChangeHandler Is Nothing Then
            WzFunc_call languageDidChangeHandler
        End If
        interfaceoptions_c_fn = True

    Case FN_interface_displayscale_populate     ' OptionChoices<unsigned int> ()
        displayScalesCount = availableDisplayScalesSorted(displayScales)
        maximumDisplayScale = wzGetMaximumDisplayScaleForCurrentWindowSize()
        Set result = New OptionChoices
        For i = 0 To displayScalesCount - 1
            displayScale = displayScales(i)
            If displayScale > maximumDisplayScale Then
                GoTo continue_
            End If
            result.addChoice CStr(displayScale) & "%", "", displayScale
continue_:
        Next
        currValue = war_GetDisplayScale()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add "Custom" item
            result.addChoice CStr(currValue) & "%", "", currValue
            result.currentIdx = result.count - 1
        End If
        Set interfaceoptions_c_fn = result

    Case FN_interface_displayscale_select       ' [weakOptionsForm](const auto& newValue) -> bool
        newValue = CLng(args(0))
        If Not wzChangeDisplayScale(newValue) Then
            debug_ 177, LOG_WARNING, "makeInterfaceOptionsForm", "Failed to change display scale from: %d to: %d", war_GetDisplayScale(), newValue
            interfaceoptions_c_fn = False
            Exit Function
        End If
        war_SetDisplayScale newValue
        widgScheduleTask WzFunc_new(FN_interface_displayscale_task, f.cap(0))
        interfaceoptions_c_fn = True

    Case FN_interface_displayscale_task         ' [weakOptionsForm]()
        Set strongOptionsForm = WeakWidget_lock(f.cap(0))
        If Not strongOptionsForm Is Nothing Then
            ' trigger layout recalc because of the display scaling change
            OptionsForm_refreshOptions strongOptionsForm, True
        End If

    Case FN_interface_showFPS_populate          ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue showFPS
        Set interfaceoptions_c_fn = result
    Case FN_interface_showFPS_select
        showFPS = CBool(args(0))
        interfaceoptions_c_fn = True

    Case FN_interface_showCounts_populate       ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue showUNITCOUNT
        Set interfaceoptions_c_fn = result
    Case FN_interface_showCounts_select
        showUNITCOUNT = CBool(args(0))
        interfaceoptions_c_fn = True

    Case FN_interface_groupsMenu_populate       ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue getGroupButtonEnabled()
        Set interfaceoptions_c_fn = result
    Case FN_interface_groupsMenu_select
        setGroupButtonEnabled CBool(args(0))
        war_setGroupsMenuEnabled getGroupButtonEnabled() ' persist
        interfaceoptions_c_fn = True

    Case FN_interface_optionsButton_populate    ' OptionChoices<uint8_t> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", 0
        result.addChoice "Opacity: 50%", "", 50
        result.addChoice "On", "", 100
        currValue = war_getOptionsButtonVisibility()
        If Not result.setCurrentIdxForValue(currValue) Then
            ' add "Custom" item
            result.addChoice "(Custom: " & CStr(currValue) & ")", "", currValue
            result.currentIdx = result.count - 1
        End If
        Set interfaceoptions_c_fn = result
    Case FN_interface_optionsButton_select      ' [inGame](const auto& newValue) -> bool
        war_setOptionsButtonVisibility CLng(args(0))
        If CBool(f.cap(0)) Then
            intAddInGameOptionsButton   ' properly refresh button state
        End If
        interfaceoptions_c_fn = True

    Case FN_interface_rotateRadar_populate      ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Fixed", "", False
        result.addChoice "Rotating", "", True
        result.setCurrentIdxForValue rotateRadar
        Set interfaceoptions_c_fn = result
    Case FN_interface_rotateRadar_select
        rotateRadar = CBool(args(0))
        interfaceoptions_c_fn = True

    Case FN_interface_radarJump_populate        ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Instant", "", True
        result.addChoice "Tracked", "", False
        result.setCurrentIdxForValue war_GetRadarJump()
        Set interfaceoptions_c_fn = result
    Case FN_interface_radarJump_select
        war_SetRadarJump CBool(args(0))
        interfaceoptions_c_fn = True

    Case FN_interface_radarZoom_get             ' []() { return war_GetRadarZoom(); }
        interfaceoptions_c_fn = war_GetRadarZoom()
    Case FN_interface_radarZoom_set             ' [](int32_t newValue)
        war_SetRadarZoom CLng(args(0))

    Case FN_interface_videoSize_populate        ' OptionChoices<FMV_MODE> ()
        Set result = New OptionChoices
        result.addChoice "1" & ChrW$(&HD7), "", FMV_1X
        result.addChoice "2" & ChrW$(&HD7), "", FMV_2X
        result.addChoice "Fullscreen", "", FMV_FULLSCREEN
        result.setCurrentIdxForValue war_GetFMVmode()
        Set interfaceoptions_c_fn = result
    Case FN_interface_videoSize_select
        war_SetFMVmode CLng(args(0))
        interfaceoptions_c_fn = True

    Case FN_interface_scanlines_populate        ' OptionChoices<SCANLINE_MODE> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", SCANLINES_OFF
        result.addChoice "50%", "", SCANLINES_50
        result.addChoice "Black", "", SCANLINES_BLACK
        result.setCurrentIdxForValue war_getScanlineMode()
        Set interfaceoptions_c_fn = result
    Case FN_interface_scanlines_select
        war_setScanlineMode CLng(args(0))
        interfaceoptions_c_fn = True

    Case FN_interface_subtitles_populate        ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue seq_GetSubtitles()
        Set interfaceoptions_c_fn = result
    Case FN_interface_subtitles_select
        seq_SetSubtitles CBool(args(0))
        interfaceoptions_c_fn = True
    End Select
End Function
