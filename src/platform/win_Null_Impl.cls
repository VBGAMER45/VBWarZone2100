VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "win_Null_Impl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sdl/gfx_api_null_sdl.cpp: sdl_Null_Impl, the headless backend's
' window-system side: no window; swapWindow throttles to
' NULLBACKEND_VSYNC_MAX_FPS "frames" per second when vsync is on.
' std::chrono::steady_clock is QueryPerformanceCounter (Currency: the count
' / 10000), std::this_thread::sleep_until a Sleep of the whole milliseconds
' left, then a spin to the end time.
Option Explicit

Implements INullBackend

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const NULLBACKEND_VSYNC_MAX_FPS As Long = 120

Private swapMode As swap_interval_mode
Private lastSwapEndTime As Double       ' seconds of the performance counter
Private minFrameDuration As Double      ' seconds

Private Sub Class_Initialize()
    ' no-op
    swapMode = swap_interval_mode_vsync
    minFrameDuration = 1# / NULLBACKEND_VSYNC_MAX_FPS
End Sub

Private Function steadyNow() As Double
    Dim c As Currency, f As Currency
    QueryPerformanceCounter c
    QueryPerformanceFrequency f
    steadyNow = CDbl(c) / CDbl(f)
End Function

Private Sub INullBackend_swapWindow()
    Dim swapEndTime As Double, minSwapEndTime As Double, ms As Long
    ' Handle throttling / sleeping
    Select Case swapMode
    Case swap_interval_mode_vsync, swap_interval_mode_adaptive_vsync
        ' simulate pseudo-vsync - try to restrict to NULLBACKEND_VSYNC_FPS "frames" per sec
        swapEndTime = steadyNow()
        minSwapEndTime = lastSwapEndTime + minFrameDuration
        If swapEndTime < minSwapEndTime Then
            ms = Int((minSwapEndTime - swapEndTime) * 1000#)
            If ms > 0 Then Sleep ms
            Do While steadyNow() < minSwapEndTime
            Loop
            swapEndTime = minSwapEndTime
        End If
        lastSwapEndTime = swapEndTime
    Case swap_interval_mode_immediate
    End Select
End Sub

Private Function INullBackend_setSwapInterval(ByVal mode As swap_interval_mode) As Boolean
    swapMode = mode
    INullBackend_setSwapInterval = True
End Function

Private Function INullBackend_getSwapInterval() As swap_interval_mode
    INullBackend_getSwapInterval = swapMode
End Function
