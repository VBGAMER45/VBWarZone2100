Attribute VB_Name = "plum_log_c"
' libplum: Copyright (c) 2020-2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/log.c, log.h, as built for 4.7.0 (RELEASE: no "file:line: " in front of a message). The mutex is
' left out (one thread). The macros PLUM_LOG_VERBOSE(...) .. PLUM_LOG_FATAL(...) are
' plum_log_write PLUM_LOG_LEVEL_<x>, fmt, args... (the format through printf_vb's vsprintf: C's escapes in the
' format string); PLUM_LOG_<x>_ENABLED is plum_log_is_enabled(PLUM_LOG_LEVEL_<x>). A plum_log_callback_t is a
' function number (0: NULL): FN_PlumLogCallback (port_mapping_manager_impl_libplum_c).
Option Explicit

Private Const BUFFER_SIZE As Long = 4096

Private log_level_ As Long                      ' (VB) log_level - PLUM_LOG_LEVEL_NONE: 0 is the initializer's NONE
Private log_callback As Long                    ' static plum_log_callback_t log_callback = NULL;

' static const char *log_level_names[]
Private Function log_level_names(ByVal level As Long) As String
    log_level_names = Choose(level + 1, "VERBOSE", "DEBUG", "INFO", "WARN", "ERROR", "FATAL")
End Function

' void plum_log_init()
Public Sub plum_log_init()
    ' mutex_init(&log_mutex, 0): (VB) no mutex
End Sub

' void plum_log_cleanup()
Public Sub plum_log_cleanup()
    ' mutex_destroy(&log_mutex): (VB) no mutex
End Sub

' void plum_set_log_level(plum_log_level_t level)
Public Sub plum_set_log_level(ByVal level As Long)
    log_level_ = level - PLUM_LOG_LEVEL_NONE
End Sub

' void plum_set_log_handler(plum_log_callback_t callback)
Public Sub plum_set_log_handler(ByVal callback As Long)
    log_callback = callback
End Sub

' bool plum_log_is_enabled(plum_log_level_t level)
Public Function plum_log_is_enabled(ByVal level As Long) As Boolean
    plum_log_is_enabled = level <> PLUM_LOG_LEVEL_NONE And level >= log_level_ + PLUM_LOG_LEVEL_NONE
End Function

' void plum_log_write(plum_log_level_t level, const char *file, int line, const char *fmt, ...)
Public Sub plum_log_write(ByVal level As Long, ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant, message As String, lt As tm, buffer As String
    If Not plum_log_is_enabled(level) Then Exit Sub

    a = args
    If log_callback <> 0 Then
        message = vsprintf(fmt, a)                   ' vsnprintf(message + len, BUFFER_SIZE - len, fmt, args)
        If Len(message) > BUFFER_SIZE - 1 Then message = Left$(message, BUFFER_SIZE - 1)
        plum_log_callback_call log_callback, level, message
    Else
        If localtime_s(lt, time_()) = 0 Then buffer = strftime("%H:%M:%S", lt) Else buffer = ""

        ' (use_color() is false on Windows: no colour codes)
        fprintf STDOUT_FILE, "%s %-7s ", buffer, log_level_names(level)
        fputs vsprintf(fmt, a), STDOUT_FILE
        fprintf STDOUT_FILE, "\n"
        fflush STDOUT_FILE
    End If
End Sub

' (VB) log_callback(level, message): the function number's function
Private Sub plum_log_callback_call(ByVal callback As Long, ByVal level As Long, ByVal message As String)
    Select Case callback
    Case FN_PlumLogCallback: PlumLogCallback level, message
    End Select
End Sub
