Attribute VB_Name = "windowoptions_c"
' src/titleui/options/windowoptions.cpp: the Window Options form (makeWindowOptionsForm), whole.
' class ScreenResolutionsModel is ScreenResolutionsModel.cls (an object the lambdas share, as the
' C++ std::shared_ptr); its static functions are here (ScreenResolutionsModel_<name>), with an
' optional<screeninfo> as (has_value, info).
Option Explicit

Public Const FN_NotEditableCondition As Long = 4400         ' OptionInfo::AvailabilityResult NotEditableCondition(const OptionInfo&)
Public Const FN_window_mode_populate As Long = 4401         ' makeWindowOptionsForm, "window.mode"
Public Const FN_window_mode_select As Long = 4402
Public Const FN_fullscreen_mode_populate As Long = 4403     ' makeWindowOptionsForm, "fullscreen.mode": [screenResolutionsModel]()
Public Const FN_fullscreen_mode_select As Long = 4404       ' [screenResolutionsModel](const auto& newResolutionIdx) -> bool
Public Const FN_window_minimizeOnFocusLoss_populate As Long = 4405 ' makeWindowOptionsForm, "window.minimizeOnFocusLoss"
Public Const FN_window_minimizeOnFocusLoss_select As Long = 4406

' MARK: -

Public Function NotEditableCondition(ByVal info As OptionInfo) As OptionInfo_AvailabilityResult
    Dim result As OptionInfo_AvailabilityResult
    Set result = New OptionInfo_AvailabilityResult
    result.available_ = False
    Set NotEditableCondition = result
End Function

' MARK: - ScreenResolutionsModel's static functions

Public Function ScreenResolutionsModel_currentResolutionString() As String
    Dim has_info As Boolean, info As screeninfo
    has_info = ScreenResolutionsModel_getCurrentResolution(info)
    ScreenResolutionsModel_currentResolutionString = ScreenResolutionsModel_resolutionString(has_info, info)
End Function

' static WzString resolutionString(const optional<screeninfo> &optInfo)
Public Function ScreenResolutionsModel_resolutionString(ByVal has_optInfo As Boolean, optInfo As screeninfo) As String
    Dim str_ As String, dpiStr As String
    If Not has_optInfo Then
        ScreenResolutionsModel_resolutionString = "Desktop Full"
        Exit Function
    End If
    str_ = va("[%d] %d " & ChrW$(&HD7) & " %d @ %.2fhz", optInfo.screen, optInfo.width, optInfo.height, CDbl(optInfo.refresh_rate))
    If optInfo.pixel_density > 1! Then
        dpiStr = va("%g", CDbl(optInfo.pixel_density))     ' fmt::format("{:g}", info.pixel_density)
        str_ = str_ & va(" [%sx DPI]", dpiStr)
    End If
    ScreenResolutionsModel_resolutionString = str_
End Function

Public Function ScreenResolutionsModel_getCurrentWindowedResolution() As screeninfo
    Dim screen As Long, windowWidth As Long, windowHeight As Long, info As screeninfo
    screen = 0
    windowWidth = 0
    windowHeight = 0
    wzGetWindowResolution screen, windowWidth, windowHeight
    info.screen = screen
    info.width = windowWidth
    info.height = windowHeight
    info.pixel_density = 0!     ' Unused.
    info.refresh_rate = -1!     ' Unused.
    ScreenResolutionsModel_getCurrentWindowedResolution = info
End Function

' static optional<screeninfo> getCurrentResolution(): has_value, the value in info
Public Function ScreenResolutionsModel_getCurrentResolution(info As screeninfo) As Boolean
    ScreenResolutionsModel_getCurrentResolution = wzGetCurrentFullscreenDisplayMode(info)
End Function

' static std::tuple<int, int, int, float, float> compareKey(const screeninfo &x): a < b of the tuples is
' compareKeyLess (screen, width, height, refresh_rate, pixel_density, in order)
Private Function compareKeyLess(a As screeninfo, b As screeninfo) As Boolean
    If a.screen <> b.screen Then compareKeyLess = ULT(a.screen, b.screen): Exit Function
    If a.width <> b.width Then compareKeyLess = a.width < b.width: Exit Function
    If a.height <> b.height Then compareKeyLess = a.height < b.height: Exit Function
    If a.refresh_rate <> b.refresh_rate Then compareKeyLess = a.refresh_rate < b.refresh_rate: Exit Function
    compareKeyLess = a.pixel_density < b.pixel_density
End Function

Private Function compareKeyEq(a As screeninfo, b As screeninfo) As Boolean
    compareKeyEq = a.screen = b.screen And a.width = b.width And a.height = b.height And a.refresh_rate = b.refresh_rate And a.pixel_density = b.pixel_density
End Function

Public Function ScreenResolutionsModel_compareLess(ByVal a_has As Boolean, a As screeninfo, ByVal b_has As Boolean, b As screeninfo) As Boolean
    If Not a_has Then
        ScreenResolutionsModel_compareLess = b_has
        Exit Function
    End If
    If Not b_has Then
        ScreenResolutionsModel_compareLess = False
        Exit Function
    End If
    ScreenResolutionsModel_compareLess = compareKeyLess(a, b)
End Function

Public Function ScreenResolutionsModel_compareEq(ByVal a_has As Boolean, a As screeninfo, ByVal b_has As Boolean, b As screeninfo) As Boolean
    If Not a_has Or Not b_has Then
        ScreenResolutionsModel_compareEq = (a_has = b_has)
        Exit Function
    End If
    ScreenResolutionsModel_compareEq = compareKeyEq(a, b)
End Function

' MARK: -

Public Function makeWindowOptionsForm() As WIDGET
    Dim result As WIDGET, optionInfo As OptionInfo, valueChanger As WIDGET, screenResolutionsModel As ScreenResolutionsModel
    Set result = OptionsForm_make()
    OptionsForm_setRefreshOptionsOnScreenSizeChange result, True

    Set screenResolutionsModel = New ScreenResolutionsModel
    screenResolutionsModel.construct

    OptionsForm_addSection result, OptionsSection_("Window", ""), True
    Set optionInfo = OptionInfo_new("window.mode", "Window Mode", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_window_mode_populate), WzFunc_new(FN_window_mode_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    OptionsForm_addSection result, OptionsSection_("Fullscreen", ""), True
    Set optionInfo = OptionInfo_new("fullscreen.mode", "Fullscreen Mode", "")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_fullscreen_mode_populate, screenResolutionsModel), WzFunc_new(FN_fullscreen_mode_select, screenResolutionsModel), True, 15)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    ' (#if !defined(__EMSCRIPTEN__))
    OptionsForm_addSection result, OptionsSection_("Behaviors", ""), True
    Set optionInfo = OptionInfo_new("window.minimizeOnFocusLoss", "Minimize on Focus Loss", "Whether the window should auto-minimize on focus loss")
    Set valueChanger = OptionsDropdown_make(WzFunc_new(FN_window_minimizeOnFocusLoss_populate), WzFunc_new(FN_window_minimizeOnFocusLoss_select), True)
    OptionsForm_addOption result, optionInfo, valueChanger, True

    Set makeWindowOptionsForm = result
End Function

' MARK: - the lambdas

Public Function windowoptions_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim result As OptionChoices, filtered As OptionChoices, i As Long, newMode As Long, model As ScreenResolutionsModel
    Dim closestResolution As Long
    Select Case f.id
    Case FN_NotEditableCondition
        Set windowoptions_c_fn = NotEditableCondition(args(0))

    Case FN_window_mode_populate                ' OptionChoices<WINDOW_MODE> ()
        Set result = New OptionChoices
        result.addChoice "Windowed", "", WINDOW_MODE_windowed
        result.addChoice "Fullscreen", "", WINDOW_MODE_fullscreen
        ' result.choices.erase(std::remove_if(..., [](const auto& item) -> bool { return !wzIsSupportedWindowMode(item.value); }), ...)
        Set filtered = New OptionChoices
        For i = 0 To result.count - 1
            If wzIsSupportedWindowMode(CLng(result.value(i))) Then
                filtered.addChoice result.displayName(i), result.helpDescription(i), result.value(i), result.disabled(i), result.icon(i)
            End If
        Next
        Set result = filtered
        result.setCurrentIdxForValue war_getWindowMode()
        Set windowoptions_c_fn = result

    Case FN_window_mode_select                  ' [](const auto& newMode) -> bool
        newMode = CLng(args(0))
        If newMode = wzGetCurrentWindowMode() Then
            windowoptions_c_fn = True
            Exit Function
        End If
        If Not wzChangeWindowMode(newMode) Then
            ' unable to change to this fullscreen mode, so disable the widget
            debug_ 218, LOG_ERROR, "makeWindowOptionsForm", "Failed to set fullscreen mode: %s", to_display_string_WINDOW_MODE(newMode)
            windowoptions_c_fn = False
            Exit Function
        End If
        war_setWindowMode newMode   ' persist config change
        windowoptions_c_fn = True

    Case FN_fullscreen_mode_populate            ' OptionChoices<size_t> ()
        Set model = f.cap(0)
        Set result = New OptionChoices
        For i = 0 To model.count - 1
            result.addChoice model.resolutionStringAt(i), "", i
        Next
        closestResolution = model.findResolutionClosestToCurrent()
        If closestResolution <> model.count Then
            result.currentIdx = closestResolution
        End If
        Set windowoptions_c_fn = result

    Case FN_fullscreen_mode_select              ' [screenResolutionsModel](const auto& newResolutionIdx) -> bool
        Set model = f.cap(0)
        windowoptions_c_fn = model.selectAt(CLng(args(0)))

    Case FN_window_minimizeOnFocusLoss_populate ' OptionChoices<MinimizeOnFocusLossBehavior> ()
        Set result = New OptionChoices
        result.addChoice "Auto", "", MinimizeOnFocusLossBehavior_Auto
        result.addChoice "Off", "", MinimizeOnFocusLossBehavior_Off
        result.addChoice "On (Fullscreen)", "", MinimizeOnFocusLossBehavior_On_Fullscreen
        result.setCurrentIdxForValue wzGetCurrentMinimizeOnFocusLossBehavior()
        Set windowoptions_c_fn = result

    Case FN_window_minimizeOnFocusLoss_select
        wzSetMinimizeOnFocusLoss CLng(args(0))
        war_setMinimizeOnFocusLoss CLng(args(0))
        windowoptions_c_fn = True
    End Select
End Function
