Attribute VB_Name = "wzapp_vb"
' (VB) lib/framework/wzapp.h's functions whose values are a screeninfo, over WzPlatform's PlatformApi (a
' screeninfo crosses as its fields): wzAvailableResolutions, wzGetCurrentFullscreenDisplayMode,
' wzChangeFullscreenDisplayMode. wzapp.h's other display, window, clipboard and system functions are
' PlatformApi's own (lib/sdl/main_sdl.cpp, cluster S).
Option Explicit

' struct screeninfo
Public Type screeninfo
    width As Long
    height As Long
    pixel_density As Single
    refresh_rate As Single
    screen As Long                          ' uint32_t
End Type

' std::vector<optional<screeninfo>> wzAvailableResolutions(): the count, modes(i) with has_(i) (False: nullopt)
Public Function wzAvailableResolutions(modes() As screeninfo, has_() As Boolean) As Long
    Dim screens() As Long, widths() As Long, heights() As Long, pixelDensities() As Single, refreshRates() As Single
    Dim n As Long, i As Long
    n = wzAvailableResolutionsList(screens, widths, heights, pixelDensities, refreshRates, has_)
    If n > 0 Then
        ReDim modes(0 To n - 1)
        For i = 0 To n - 1
            modes(i).screen = screens(i)
            modes(i).width = widths(i)
            modes(i).height = heights(i)
            modes(i).pixel_density = pixelDensities(i)
            modes(i).refresh_rate = refreshRates(i)
        Next
    End If
    wzAvailableResolutions = n
End Function

' optional<screeninfo> wzGetCurrentFullscreenDisplayMode(): has_value, the value in info
Public Function wzGetCurrentFullscreenDisplayMode(info As screeninfo) As Boolean
    wzGetCurrentFullscreenDisplayMode = wzGetCurrentFullscreenDisplayModeInfo(info.screen, info.width, info.height, info.pixel_density, info.refresh_rate)
End Function

' bool wzChangeFullscreenDisplayMode(optional<screeninfo> config)
Public Function wzChangeFullscreenDisplayMode(ByVal has_config As Boolean, config As screeninfo) As Boolean
    wzChangeFullscreenDisplayMode = wzChangeFullscreenDisplayModeTo(has_config, config.screen, config.width, config.height, config.pixel_density, config.refresh_rate)
End Function
