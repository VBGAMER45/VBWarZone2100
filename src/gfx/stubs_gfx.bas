Attribute VB_Name = "stubs_gfx"
' Stand-ins for what WzGfx calls that is not ported yet (cluster D's merge). Each
' keeps the C++ name and signature of what it stands for, so its callers need
' no change: piedraw.cpp (cluster E), the crash handler (not ported), and the
' exe's and WzPlatform's functions WzGfx reaches with no callback yet.
Option Explicit

Private screenWidthSet_ As Long
Private screenHeightSet_ As Long

'--------------------------------------------------------------------------
' (fork B)
' Fork B (gl): stand-ins for WzGfx's stubs_gfx.bas. Each keeps the C++ name and signature.

' lib/exceptionhandler/dumpinfo.h (not ported): the crash dump's extra lines
Public Sub addDumpInfo(ByVal inbuffer As String)
End Sub

' lib/framework/wzapp.h, implemented by lib/sdl/main_sdl.cpp (WzPlatform): resets the graphics
' settings that may cause an OpenGL failure (MSAA, ...). Stand-in until WzPlatform gets a callback
' into WzGfx (IGLBackend could carry it, or WzFramework's IFrameworkPlatform).
Public Sub wzResetGfxSettingsOnFailure()
End Sub

' lib/framework/wzapp.h, implemented by lib/sdl/main_sdl.cpp (WzPlatform): the fatal error dialog
' for a graphics backend failure. Stand-in: the log only.
Public Sub wzDisplayFatalGfxBackendFailure(Optional ByVal additionalErrorDetails As String = "")
    debug_ 0, LOG_ERROR, "wzDisplayFatalGfxBackendFailure", "(stand-in) Fatal graphics backend failure: %s", additionalErrorDetails
End Sub


'--------------------------------------------------------------------------
' (fork C)
' ===== Fork C: stand-ins for WzGfx's stubs_gfx.bas (what fork C's modules call that nobody provides yet) =====

'--------------------------------------------------------------------------
' lib/ivis_opengl/piedraw.cpp (cluster E). ShadowMode is GfxApi's enum (see d_C_gfxapi.txt).





'--------------------------------------------------------------------------
' lib/sdl/main_sdl.cpp: unsigned int screenWidth = 0; unsigned int screenHeight = 0; (WzPlatform's, which
' WzGfx cannot reach: WzPlatform uses WzGfx). Stand-in: the value screen_c last stored, else the video
' buffer's size (the platform sets both from the same screenWidth/screenHeight). Suggested real fix: the
' platform hands WzGfx its screenWidth/screenHeight whenever they change, as it does
' pie_SetVideoBufferWidth/Height (a GfxApi Property Let), and screenInitialise's MAX(.., 640) is handed back.

Public Property Get screenWidth() As Long
    If screenWidthSet_ <> 0 Then screenWidth = screenWidthSet_ Else screenWidth = pie_GetVideoBufferWidth()
End Property

Public Property Let screenWidth(ByVal value As Long)
    screenWidthSet_ = value
End Property

Public Property Get screenHeight() As Long
    If screenHeightSet_ <> 0 Then screenHeight = screenHeightSet_ Else screenHeight = pie_GetVideoBufferHeight()
End Property

Public Property Let screenHeight(ByVal value As Long)
    screenHeightSet_ = value
End Property

'--------------------------------------------------------------------------
' src/console.cpp: bool addConsoleMessage(const char *Text, CONSOLE_TEXT_JUSTIFICATION jusType, SDWORD player,
' bool team = false, UDWORD duration = DEFAULT_CONSOLE_MESSAGE_DURATION) (the exe's: needs a callback from
' WzGfx to the exe; screen_c's screenshot messages)
Public Function addConsoleMessage(ByVal Text As String, ByVal jusType As Long, ByVal player As Long, Optional ByVal team As Boolean = False, Optional ByVal duration As Long = -1) As Boolean
    addConsoleMessage = True
End Function


'--------------------------------------------------------------------------
' (fork T)
' Fork T's stand-ins for stubs_gfx.bas (WzGfx). textdraw_gdi's iV_TextInit calls
' these when the base fonts fail to load (then LOG_FATAL and abort).

' src/init.cpp's (the exe's init_c has them): WzGfx has no callback into the exe for them yet
Public Sub debugOutputSearchPaths()
End Sub

Public Sub debugOutputSearchPathMountErrors()
End Sub
