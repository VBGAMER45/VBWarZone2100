Attribute VB_Name = "wztime_c"
' lib/framework/wztime.cpp (compiled into every project: small and without
' state, like the C runtime it wraps). Ported: getUtcTime, getLocalTimeOpt,
' getLocalTime, formatLocalDateTime (std::put_time is strftime). Missing:
' getAscTime.
' optional<tm> is a Boolean result with the tm ByRef.
Option Explicit

' C: tm getUtcTime(std::time_t const &timer)
Public Function getUtcTime(ByVal timer As Double) As tm
    Dim timeinfo As tm
    If gmtime_s(timeinfo, timer) <> 0 Then
        debug_ 33, LOG_ERROR, "getUtcTime", "Invalid time_t argument"
    End If
    getUtcTime = timeinfo
End Function

' C: optional<tm> getLocalTimeOpt(std::time_t const &timer, bool quiet = false)
Public Function getLocalTimeOpt(timeinfo As tm, ByVal timer As Double, Optional ByVal quiet As Boolean = False) As Boolean
    Dim result As Long
    result = localtime_s(timeinfo, timer)
    If result <> 0 Then
        If Not quiet Then
            debug_ 58, LOG_ERROR, "getLocalTimeOpt", "localtime_s failed with error: %s", "Invalid argument"
        End If
        Exit Function
    End If
    getLocalTimeOpt = True
End Function

' C: tm getLocalTime(std::time_t const &timer, bool quiet = false)
Public Function getLocalTime(ByVal timer As Double, Optional ByVal quiet As Boolean = False) As tm
    Dim result As tm, zeroResult As tm
    If Not getLocalTimeOpt(result, timer, quiet) Then
        getLocalTime = zeroResult
        Exit Function
    End If
    getLocalTime = result
End Function

' static std::string const formatLocalDateTime_Internal(char const *format, tm const &tmVal, optional<std::locale>)
Private Function formatLocalDateTime_Internal(ByVal format_ As String, tmVal As tm) As String
    formatLocalDateTime_Internal = strftime(format_, tmVal)
End Function

' std::string const formatLocalDateTime(char const *format, std::time_t const &timer); timer missing:
' formatLocalDateTime(char const *format) (std::time(nullptr))
Public Function formatLocalDateTime(ByVal format_ As String, Optional ByVal timer As Variant) As String
    Dim timeinfo As tm
    If IsMissing(timer) Then
        timeinfo = getLocalTime(time_())
    Else
        timeinfo = getLocalTime(CDbl(timer))
    End If
    formatLocalDateTime = formatLocalDateTime_Internal(format_, timeinfo)
End Function
