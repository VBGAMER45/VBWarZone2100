Attribute VB_Name = "audiooptions_c"
' src/titleui/options/audiooptions.cpp: the Audio Options form (makeAudioOptionsForm) and its
' OptionsButton widget class, whole. The widget class follows docs/CONVENTIONS.md, "Widgets" (its
' fields: OptionsButton.cls); OptionValueChangerInterface is optionsform_c's ovc_ virtuals.
Option Explicit

' (VB) widget class OptionsButton : W_BUTTON

Public Const FN_HRTFSupported As Long = 4220                ' OptionInfo::AvailabilityResult HRTFSupported(const OptionInfo&)
Public Const FN_audio_voiceVol_get As Long = 4221           ' makeAudioOptionsForm, "audio.voiceVol": []() { return ...; }
Public Const FN_audio_voiceVol_set As Long = 4222           ' makeAudioOptionsForm, "audio.voiceVol": [](int32_t newValue)
Public Const FN_audio_fxVol_get As Long = 4223              ' makeAudioOptionsForm, "audio.fxVol"
Public Const FN_audio_fxVol_set As Long = 4224
Public Const FN_audio_musicVol_get As Long = 4225           ' makeAudioOptionsForm, "audio.musicVol"
Public Const FN_audio_musicVol_set As Long = 4226
Public Const FN_audio_hrtf_populate As Long = 4227          ' makeAudioOptionsForm, "audio.hrtf"
Public Const FN_audio_hrtf_select As Long = 4228
Public Const FN_audio_subtitles_populate As Long = 4229     ' makeAudioOptionsForm, "interface.video.subtitles"
Public Const FN_audio_subtitles_select As Long = 4230
Public Const FN_audio_musicManager_update As Long = 4231    ' makeAudioOptionsForm, "audio.music.manager": [](OptionsButton& but)
Public Const FN_audio_musicManager_click As Long = 4232     ' makeAudioOptionsForm, "audio.music.manager": [](W_BUTTON& but)
Public Const FN_audio_groupReporting_populate As Long = 4233 ' makeAudioOptionsForm, "audio.cues.groupReporting"
Public Const FN_audio_groupReporting_select As Long = 4234
Public Const FN_doOpenMusicManager_onClose As Long = 4235   ' doOpenMusicManager, []() { // no-op on close function }

' (VB) PIELIGHT's byte.a, read and set
Private Function pielightA(ByVal c As Long) As Long
    pielightA = SHRU(c, 24) And &HFF&
End Function

Private Function pielightWithA(ByVal c As Long, ByVal a As Long) As Long
    pielightWithA = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' OptionsButton::OptionsButton() (with the member initialisers)
Public Sub OptionsButton_construct(ByVal self As WIDGET)
    Dim d As OptionsButton
    W_BUTTON_construct self
    self.kind = K_OptionsButton
    Set d = OptionsButton_(self)
    Set d.wzText = WzCachedText_new()
    d.FontID = font_regular
    d.bindingColor = psPalette(WZCOL_FORM_TEXT)
    d.cachedIdealTextWidth = 0
    d.lastWidgetWidth = 0
    d.verticalPadding = 10
    d.horizontalPadding = 14
    d.isTruncated = False
    d.highlightBackgroundColor = psPalette(WZCOL_MENU_SCORE_BUILT)
End Sub

' static std::shared_ptr<OptionsButton> make(UpdateButtonFunc updateButtonFunc, bool needsStateUpdates)
' (UpdateButtonFunc: void (OptionsButton&))
Public Function OptionsButton_make(ByVal updateButtonFunc As WzFunc, ByVal needsStateUpdates As Boolean) As WIDGET
    Dim result As WIDGET, d As OptionsButton
    If Not (Not updateButtonFunc Is Nothing) Then ASSERT_FAILURE "updateButtonFunc != nullptr", "src/titleui/options/audiooptions.cpp:86", "make", "Missing required updateButtonFunc": Exit Function

    Set result = WIDGET_alloc(K_OptionsButton)
    OptionsButton_construct result
    Set d = OptionsButton_(result)
    Set d.updateButtonFunc = updateButtonFunc
    d.needsStateUpdates = needsStateUpdates

    WzFunc_call updateButtonFunc, result

    Set OptionsButton_make = result
End Function

Public Sub OptionsButton_setString(ByVal self As WIDGET, ByVal string_ As String)
    Dim d As OptionsButton
    W_BUTTON_setString self, string_
    Set d = OptionsButton_(self)
    d.cachedIdealTextWidth = iV_GetTextWidth(W_BUTTON_(self).pText, d.FontID)
End Sub

Public Sub OptionsButton_setFont(ByVal self As WIDGET, ByVal font As iV_fonts)
    Dim d As OptionsButton
    Set d = OptionsButton_(self)
    If font = d.FontID Then
        Exit Sub
    End If
    d.FontID = font
    d.cachedIdealTextWidth = iV_GetTextWidth(W_BUTTON_(self).pText, d.FontID)
End Sub

Public Sub OptionsButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As OptionsButton, pText As String, x0 As Long, y0 As Long, w As Long, h As Long
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, fontID As iV_fonts, shrunk As iV_fonts
    Dim maxTextDisplayableWidth As Long, backX0 As Long, backY0 As Long, backX1 As Long, backY1 As Long
    Dim drawHighlightBorder As Boolean, backgroundColor As Long, textColor As Long
    Dim maxDisplayableMainTextWidth As Long, textX0 As Long, textY0 As Long
    Set d = OptionsButton_(self)
    pText = W_BUTTON_(self).pText
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = (self.getState() And WBUT_HIGHLIGHT) <> 0

    fontID = d.wzText.getFontID()
    If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzText.getText() <> pText Then
        fontID = d.FontID
        d.wzText.setText pText, fontID
    End If

    maxTextDisplayableWidth = w - (d.horizontalPadding * 2)
    If d.wzText.deref().width() > maxTextDisplayableWidth Then
        ' text would exceed the bounds of the button area
        ' try to shrink font so it fits
        Do
            If fontID = font_small Or fontID = font_bar Then
                Exit Do
            End If
            If Not iV_ShrinkFont(fontID, shrunk) Then
                Exit Do
            End If
            fontID = shrunk
            d.wzText.setText pText, fontID
        Loop While d.wzText.deref().width() > maxTextDisplayableWidth
    End If
    d.lastWidgetWidth = self.width()

    ' Draw background
    backX0 = x0
    backY0 = y0
    backX1 = x0 + w - 1
    backY1 = y0 + h - 1
    drawHighlightBorder = isHighlight
    If drawHighlightBorder Then
        pie_UniTransBoxFill backX0, backY0, backX1, backY1, pal_RGBA(255, 255, 255, 225)
    End If
    backX0 = backX0 + 1
    backY0 = backY0 + 1
    backX1 = backX1 - 1
    backY1 = backY1 - 1
    If isDown Then
        backgroundColor = psPalette(WZCOL_FORM_DARK)
    ElseIf isHighlight And pielightA(d.highlightBackgroundColor) <> 0 Then     ' !isTransparent()
        backgroundColor = d.highlightBackgroundColor
    Else
        backgroundColor = psPalette(WZCOL_FORM_DARK)
        backgroundColor = pielightWithA(backgroundColor, pielightA(backgroundColor) \ 3)
    End If
    pie_UniTransBoxFill backX0, backY0, backX1, backY1, backgroundColor

    ' Draw the main text
    If Not isDisabled Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isDisabled Then
        textColor = pielightWithA(textColor, pielightA(textColor) \ 2)
    End If

    maxDisplayableMainTextWidth = maxTextDisplayableWidth
    d.isTruncated = maxDisplayableMainTextWidth < d.wzText.deref().width()
    If d.isTruncated Then
        maxDisplayableMainTextWidth = maxDisplayableMainTextWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
    End If
    textX0 = x0 + (w - MIN_(maxDisplayableMainTextWidth, d.wzText.deref().width())) \ 2
    textY0 = y0 + (h - d.wzText.deref().lineSize()) \ 2 - d.wzText.deref().aboveBase()
    d.wzText.deref().render textX0, textY0, textColor, 0!, maxDisplayableMainTextWidth
    If d.isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxDisplayableMainTextWidth + 2, textY0, textColor
    End If
End Sub

Public Sub OptionsButton_geometryChanged(ByVal self As WIDGET)
    ' currently no-op
End Sub

Public Function OptionsButton_idealWidth(ByVal self As WIDGET) As Long
    Dim d As OptionsButton
    Set d = OptionsButton_(self)
    OptionsButton_idealWidth = (d.horizontalPadding * 2) + d.cachedIdealTextWidth
End Function

Public Function OptionsButton_idealHeight(ByVal self As WIDGET) As Long
    Dim d As OptionsButton
    Set d = OptionsButton_(self)
    OptionsButton_idealHeight = (d.verticalPadding * 2) + iV_GetTextLineSize(d.FontID)
End Function

Public Function OptionsButton_ovc_is(ByVal self As WIDGET) As Boolean
    OptionsButton_ovc_is = True
End Function

Public Sub OptionsButton_ovc_update(ByVal self As WIDGET, Optional ByVal force As Boolean = False)
    Dim d As OptionsButton
    Set d = OptionsButton_(self)
    If Not force And Not d.needsStateUpdates Then
        Exit Sub
    End If

    WzFunc_call d.updateButtonFunc, self
End Sub

Public Sub OptionsButton_ovc_informAvailable(ByVal self As WIDGET, ByVal isAvailable As Boolean)
    If Not isAvailable Then self.setState WBUT_DISABLE Else self.setState 0
End Sub

Public Sub OptionsButton_ovc_addOnChangeHandler(ByVal self As WIDGET, ByVal handler As WzFunc)
    ' no-op for buttons - expectation is that an onclick-handler will be added
End Sub

' MARK: -

Public Function HRTFSupported(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = (sound_GetHRTFMode() <> HRTFMode_Unsupported)
    result.localizedUnavailabilityReason = "HRTF is not supported on your device / system / OpenAL library"
    Set HRTFSupported = result
End Function

Private Sub doOpenMusicManager(ByVal currentScreen As W_SCREEN)
    Select Case GetGameMode()
    Case GS_NORMAL
        showInGameOptionsScreen currentScreen, makeMusicManagerForm(True), WzFunc_new(FN_doOpenMusicManager_onClose)
    Case GS_TITLE_SCREEN
        changeTitleUI WzMusicManagerTitleUI_new(wzTitleUICurrent)
    Case Else
    End Select
End Sub

' MARK: -

Public Function makeAudioOptionsForm() As WIDGET
    Dim result As WIDGET, optionInfo As OptionInfo, valueChanger As WIDGET
    Set result = OptionsForm_make()

    ' Volume:
    OptionsForm_addSection result, OptionsSection_("Volume", ""), True
    Set optionInfo = OptionInfo_new("audio.voiceVol", "Voice Volume", "")
    Set valueChanger = OptionsSlider_make(0, AUDIO_VOL_MAX, 1, WzFunc_new(FN_audio_voiceVol_get), WzFunc_new(FN_audio_voiceVol_set), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("audio.fxVol", "FX Volume", "")
    Set valueChanger = OptionsSlider_make(0, AUDIO_VOL_MAX, 1, WzFunc_new(FN_audio_fxVol_get), WzFunc_new(FN_audio_fxVol_set), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("audio.musicVol", "Music Volume", "")
    Set valueChanger = OptionsSlider_make(0, AUDIO_VOL_MAX, 1, WzFunc_new(FN_audio_musicVol_get), WzFunc_new(FN_audio_musicVol_set), False)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Output:
    OptionsForm_addSection result, OptionsSection_("Output", ""), True
    Set optionInfo = OptionInfo_new("audio.hrtf", "HRTF", "HRTF can enable enhanced spatialization when used with headphones or some stereo output.")
    optionInfo.addAvailabilityCondition WzFunc_new(FN_HRTFSupported)
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_audio_hrtf_populate), WzFunc_new(FN_audio_hrtf_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set optionInfo = OptionInfo_new("interface.video.subtitles", "Subtitles", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_audio_subtitles_populate), WzFunc_new(FN_audio_subtitles_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Music:
    OptionsForm_addSection result, OptionsSection_("Music", ""), True
    Set optionInfo = OptionInfo_new("audio.music.manager", "Music Manager", "")
    Set valueChanger = OptionsButton_make(WzFunc_new(FN_audio_musicManager_update), False)
    W_BUTTON_addOnClickHandler valueChanger, WzFunc_new(FN_audio_musicManager_click)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' Audio Cues:
    OptionsForm_addSection result, OptionsSection_("Audio Cues", ""), True
    Set optionInfo = OptionInfo_new("audio.cues.groupReporting", "Group Reporting", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_audio_groupReporting_populate), WzFunc_new(FN_audio_groupReporting_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set makeAudioOptionsForm = result
End Function

' MARK: - the lambdas

Public Function audiooptions_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim result As OptionChoices, currentMode As Long, isAutoSetting As Boolean, currentModeStr As String, but As WIDGET
    Dim v As Long
    Select Case f.id
    Case FN_HRTFSupported
        Set audiooptions_c_fn = HRTFSupported(args(0))

    Case FN_audio_voiceVol_get                  ' []() { return static_cast<int32_t>(sound_GetUIVolume() * 100.0f); }
        audiooptions_c_fn = CLng(Fix(CSng(sound_GetUIVolume() * 100!)))
    Case FN_audio_voiceVol_set                  ' [](int32_t newValue)
        sound_SetUIVolume CSng(CSng(CLng(args(0))) / 100!)
    Case FN_audio_fxVol_get
        audiooptions_c_fn = CLng(Fix(CSng(sound_GetEffectsVolume() * 100!)))
    Case FN_audio_fxVol_set
        sound_SetEffectsVolume CSng(CSng(CLng(args(0))) / 100!)
    Case FN_audio_musicVol_get
        audiooptions_c_fn = CLng(Fix(CSng(sound_GetMusicVolume() * 100!)))
    Case FN_audio_musicVol_set
        sound_SetMusicVolume CSng(CSng(CLng(args(0))) / 100!)

    Case FN_audio_hrtf_populate                 ' OptionChoices<HRTFMode> ()
        Set result = New OptionChoices
        ' note: sound_GetHRTFMode() returns the current actual status (and never auto, even if the request was "auto")
        currentMode = sound_GetHRTFMode()
        If currentMode = HRTFMode_Unsupported Then
            result.addChoice "Unsupported", "", HRTFMode_Unsupported
            result.currentIdx = 0
            Set audiooptions_c_fn = result
            Exit Function
        End If
        result.addChoice "Disabled", "", HRTFMode_Disabled
        result.addChoice "Enabled", "", HRTFMode_Enabled
        isAutoSetting = war_GetHRTFMode() = HRTFMode_Auto
        If isAutoSetting Then
            Select Case currentMode
            Case HRTFMode_Disabled
                currentModeStr = "Disabled"
            Case HRTFMode_Enabled
                currentModeStr = "Enabled"
            Case Else
            End Select
            result.addChoice "Auto" & " (" & currentModeStr & ")", "", HRTFMode_Auto
        Else
            result.addChoice "Auto", "", HRTFMode_Auto
        End If
        If isAutoSetting Then v = HRTFMode_Auto Else v = currentMode
        result.setCurrentIdxForValue v
        Set audiooptions_c_fn = result

    Case FN_audio_hrtf_select                   ' [](const auto& newValue) -> bool
        If Not sound_SetHRTFMode(CLng(args(0))) Then
            audiooptions_c_fn = False
            Exit Function
        End If
        war_SetHRTFMode CLng(args(0))
        audiooptions_c_fn = True

    Case FN_audio_subtitles_populate            ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue seq_GetSubtitles()
        Set audiooptions_c_fn = result

    Case FN_audio_subtitles_select
        seq_SetSubtitles CBool(args(0))
        audiooptions_c_fn = True

    Case FN_audio_musicManager_update           ' [](OptionsButton& but)
        Set but = args(0)
        but.setString "Open Music Manager"

    Case FN_audio_musicManager_click            ' [](W_BUTTON& but)
        Set but = args(0)
        doOpenMusicManager but.screenPointer

    Case FN_audio_groupReporting_populate       ' OptionChoices<bool> ()
        Set result = New OptionChoices
        result.addChoice "Off", "", False
        result.addChoice "On", "", True
        result.setCurrentIdxForValue war_getPlayAudioCue_GroupReporting()
        Set audiooptions_c_fn = result

    Case FN_audio_groupReporting_select
        war_setPlayAudioCue_GroupReporting CBool(args(0))
        audiooptions_c_fn = True

    Case FN_doOpenMusicManager_onClose          ' []() { }
        ' no-op on close function
    End Select
End Function
