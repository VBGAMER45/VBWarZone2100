Attribute VB_Name = "timer_c"
' lib/sequence/timer.cpp, timer.h: the video player's timer (gettimeofday's microseconds). WZ_OS_WIN: the
' file's own gettimeofday over GetSystemTimeAsFileTime, its uint64_t arithmetic in Decimal (exact, as the
' C's); winsock2.h's struct timeval has longs (32 bits). gettimeofday is this module's (timer.h declares it,
' nobody else calls it).
Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

' winsock2.h: struct timeval
Private Type timeval
    tv_sec As Long                          ' long
    tv_usec As Long                         ' long
End Type

Private Declare Sub GetSystemTimeAsFileTime Lib "kernel32" (lpSystemTimeAsFileTime As FILETIME)

Private startTimeInMicroSec As Double       ' starting time in microseconds
Private endTimeInMicroSec As Double         ' ending time in microseconds
Private stopped As Boolean                  ' stop flag

Private startCount As timeval
Private endCount As timeval

' (VB) a DWORD as a Decimal (its unsigned value)
Private Function u32dec(ByVal x As Long) As Variant
    If x < 0 Then
        u32dec = CDec(x) + CDec(4294967296#)
    Else
        u32dec = CDec(x)
    End If
End Function

'*
' * The difference between the FAT32 and Unix epoch.
' *
' * The FAT32 epoch starts at 1 January 1601 while the Unix epoch starts at 1
' * January 1970. And apparently we gained 3.25 days in that time period.
' *
' * Thus the amount of micro seconds passed between these dates can be computed
' * as follows:
' * \f[((1970 - 1601) \cdot 365.25 + 3.25) \cdot 86400 \cdot 1000000\f]
' *
' * Use 1461 and 13 instead of 365.25 and 3.25 respectively because we can't use
' * floating point math here.
' (VB) static const uint64_t usecs_between_fat32_and_unix_epoch: a Decimal
Private Function usecs_between_fat32_and_unix_epoch() As Variant
    usecs_between_fat32_and_unix_epoch = CDec((1970 - 1601) * 1461& + 13) * CDec(86400) / CDec(4) * CDec(1000000)
End Function

' int gettimeofday(struct timeval *tv, struct timezone *tz): tz is always NULL here (tzIsNull)
Private Function gettimeofday(tv As timeval, ByVal tzIsNull As Boolean) As Long
    Dim ft As FILETIME, systime As Variant, usec As Variant, s As Variant
    If Not (tzIsNull) Then ASSERT_FAILURE "tz == NULL", "lib/sequence/timer.cpp:53", "gettimeofday", "This gettimeofday implementation doesn't provide timezone info."

    ' (if (tv): tv is always given)
    ' Retrieve the current time expressed as 100 nano-second
    ' intervals since 1 January 1601 (UTC).
    GetSystemTimeAsFileTime ft
    systime = u32dec(ft.dwHighDateTime) * CDec(4294967296#) + u32dec(ft.dwLowDateTime)

    ' Convert to micro seconds since 1 January 1970 (UTC).
    usec = Int(systime / CDec(10)) - usecs_between_fat32_and_unix_epoch()

    s = Int(usec / CDec(1000000))
    tv.tv_sec = CLng(s)                     ' usec / (uint64_t)1000000
    tv.tv_usec = CLng(usec - s * CDec(1000000))  ' usec % (uint64_t)1000000

    gettimeofday = 0
End Function

' Uses the highest resolution timers avail on windows & linux
Public Sub Timer_Init()
    startCount.tv_sec = 0: startCount.tv_usec = 0
    endCount.tv_sec = 0: endCount.tv_usec = 0

    stopped = False
    startTimeInMicroSec = 0
    endTimeInMicroSec = 0
End Sub

' start timer
Public Sub Timer_start()
    stopped = False ' reset stop flag
    gettimeofday startCount, True
End Sub

' stop the timer
Public Sub Timer_stop()
    stopped = True ' set timer stopped flag

    gettimeofday endCount, True
End Sub

' get elapsed time in microseconds
Public Function Timer_getElapsedMicroSecs() As Double
    If Not stopped Then
        gettimeofday endCount, True
    End If
    startTimeInMicroSec = (startCount.tv_sec * 1000000#) + startCount.tv_usec
    endTimeInMicroSec = (endCount.tv_sec * 1000000#) + endCount.tv_usec

    Timer_getElapsedMicroSecs = endTimeInMicroSec - startTimeInMicroSec
End Function

' get elapsed time in milliseconds
Public Function Timer_getElapsedMilliSecs() As Double
    Timer_getElapsedMilliSecs = Timer_getElapsedMicroSecs() * 0.001
End Function
