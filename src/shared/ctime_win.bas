Attribute VB_Name = "ctime_win"
' C's <time.h> for the port, on Win32 (no C++ counterpart; compiled into
' every project). A time_t is a Double: seconds since 1970-01-01 UTC, exact
' for any date Warzone handles. struct tm keeps C's fields and ranges
' (tm_mon 0..11, tm_year since 1900, tm_wday 0 = Sunday).
Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Sub GetSystemTimeAsFileTime Lib "kernel32" (lpSystemTimeAsFileTime As FILETIME)
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (ByVal lpTimeZone As Long, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long

Public Type tm
    tm_sec As Long
    tm_min As Long
    tm_hour As Long
    tm_mday As Long
    tm_mon As Long
    tm_year As Long
    tm_wday As Long
    tm_yday As Long
    tm_isdst As Long
End Type

Private Const EPOCH_DIFF As Double = 11644473600#    ' seconds from 1601-01-01 to 1970-01-01

Private Function UnsignedLong(ByVal l As Long) As Double
    If l < 0 Then UnsignedLong = l + 4294967296# Else UnsignedLong = l
End Function

' C: time (NULL)
Public Function time_() As Double
    Dim ft As FILETIME
    GetSystemTimeAsFileTime ft
    time_ = Int((UnsignedLong(ft.dwHighDateTime) * 4294967296# + UnsignedLong(ft.dwLowDateTime)) / 10000000#) - EPOCH_DIFF
End Function

Private Function TimeToFileTime(ByVal timer As Double, ft As FILETIME) As Boolean
    Dim t As Double, hi As Double, lo As Double
    t = (timer + EPOCH_DIFF) * 10000000#
    If t < 0 Or t >= 9.22337203685478E+18 Then Exit Function
    hi = Int(t / 4294967296#)
    lo = t - hi * 4294967296#
    If hi >= 2147483648# Then hi = hi - 4294967296#
    If lo >= 2147483648# Then lo = lo - 4294967296#
    ft.dwHighDateTime = hi
    ft.dwLowDateTime = lo
    TimeToFileTime = True
End Function

Private Sub SystemTimeToTm(st As SYSTEMTIME, t As tm)
    Static days(0 To 11) As Long
    Dim leap As Boolean
    If days(1) = 0 Then
        days(0) = 0: days(1) = 31: days(2) = 59: days(3) = 90: days(4) = 120: days(5) = 151
        days(6) = 181: days(7) = 212: days(8) = 243: days(9) = 273: days(10) = 304: days(11) = 334
    End If
    t.tm_sec = st.wSecond
    t.tm_min = st.wMinute
    t.tm_hour = st.wHour
    t.tm_mday = st.wDay
    t.tm_mon = st.wMonth - 1
    t.tm_year = st.wYear - 1900
    t.tm_wday = st.wDayOfWeek
    leap = ((st.wYear Mod 4 = 0) And (st.wYear Mod 100 <> 0)) Or (st.wYear Mod 400 = 0)
    t.tm_yday = days(st.wMonth - 1) + st.wDay - 1
    If leap And st.wMonth > 2 Then t.tm_yday = t.tm_yday + 1
    t.tm_isdst = 0
End Sub

' C: gmtime_s (&t, &timer); returns 0 on success, like the Microsoft C library.
Public Function gmtime_s(t As tm, ByVal timer As Double) As Long
    Dim ft As FILETIME, st As SYSTEMTIME, blank As tm
    t = blank
    If Not TimeToFileTime(timer, ft) Then
        gmtime_s = 22                          ' EINVAL
        Exit Function
    End If
    If FileTimeToSystemTime(ft, st) = 0 Then
        gmtime_s = 22
        Exit Function
    End If
    SystemTimeToTm st, t
End Function

' C: localtime_s (&t, &timer); returns 0 on success.
Public Function localtime_s(t As tm, ByVal timer As Double) As Long
    Dim ft As FILETIME, utc As SYSTEMTIME, loc_ As SYSTEMTIME, blank As tm
    t = blank
    If Not TimeToFileTime(timer, ft) Then
        localtime_s = 22
        Exit Function
    End If
    If FileTimeToSystemTime(ft, utc) = 0 Then
        localtime_s = 22
        Exit Function
    End If
    If SystemTimeToTzSpecificLocalTime(0, utc, loc_) = 0 Then
        localtime_s = 22
        Exit Function
    End If
    SystemTimeToTm loc_, t
End Function

' C: strftime, the conversions the port uses so far: %H %M %S %d %m %Y %y %j %%.
Public Function strftime(ByVal fmt As String, t As tm) As String
    Dim i As Long, c As String, out As String
    i = 1
    Do While i <= Len(fmt)
        c = Mid$(fmt, i, 1)
        If c = "%" And i < Len(fmt) Then
            i = i + 1
            Select Case Mid$(fmt, i, 1)
            Case "H": out = out & Format$(t.tm_hour, "00")
            Case "M": out = out & Format$(t.tm_min, "00")
            Case "S": out = out & Format$(t.tm_sec, "00")
            Case "d": out = out & Format$(t.tm_mday, "00")
            Case "m": out = out & Format$(t.tm_mon + 1, "00")
            Case "Y": out = out & CStr(t.tm_year + 1900)
            Case "y": out = out & Format$((t.tm_year + 1900) Mod 100, "00")
            Case "j": out = out & Format$(t.tm_yday + 1, "000")
            Case "%": out = out & "%"
            Case Else: out = out & "%" & Mid$(fmt, i, 1)
            End Select
        Else
            out = out & c
        End If
        i = i + 1
    Loop
    strftime = out
End Function
