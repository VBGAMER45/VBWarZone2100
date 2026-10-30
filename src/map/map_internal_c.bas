Attribute VB_Name = "map_internal_c"
' lib/wzmaplib/src/map_internal.cpp, map_internal.h: _printLog, called by
' the library's debug(logger, level, ...) macro, is
'   printLog_ pCustomLogger, <C++ line>, LogLevel_<level>, "<function>", fmt, args
' with the macro's __LINE__ and __FUNCTION__ written out (debug_'s
' convention, docs/CONVENTIONS.md). LOG_INFO_VERBOSE, LOG_INFO,
' LOG_WARNING / LOG_SYNTAX_WARNING and LOG_ERROR are LogLevel_Info_Verbose,
' LogLevel_Info, LogLevel_Warning and LogLevel_Error (LoggingProtocol.cls).
' The message is formatted into the static 2048-byte buffer of vssprintf:
' cut to 2047 bytes of UTF-8. strEndsWith is string_ext_h's (the same
' code); win_utf8ToUtf16 and win_utf16toUtf8 are not needed (a VB String
' is UTF-16 already).
Option Explicit

Public Sub printLog_(ByVal logger As LoggingProtocol, ByVal line_ As Long, ByVal level As LogLevel, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant, outputBuffer As String, b() As Byte, n As Long
    a = args
    outputBuffer = vsprintf(str_, a)
    If Len(outputBuffer) > 682 Then             ' 2047 / 3: shorter text is under 2048 bytes
        n = utf8_encode(outputBuffer, b)
        If n > 2047 Then outputBuffer = utf8_decode(b, 0, 2047)
    End If
    If Not logger Is Nothing Then
        logger.printLog level, function_, line_, outputBuffer
    End If
End Sub
