Attribute VB_Name = "fixedpoint_h"
' lib/framework/fixedpoint.h: fixed point degrees (DEG_360 is a full
' turn). The DEG macro is an integer function here; RADIANS and UNDEG are
' the float versions (display only).
Option Explicit

' PSX-style float emulation: 12 digit semi-floats stored in an int
' FIXME!
Public Const FP12_SHIFT As Long = 12
Public Const FP12_MULTIPLIER As Long = 4096     ' (1 << FP12_SHIFT)

' Global Definitions (CONSTANTS)
Public Const DEG_360 As Long = 65536
Public Const DEG_1 As Single = 65536! / 360!

Public Function RADIANS(ByVal degrees As Single) As Single
    RADIANS = degrees * (3.141592! / 180!)
End Function

Public Function UNDEG(ByVal angle As Long) As Single
    UNDEG = CSng(angle And &HFFFF&) * (360! / 65536!) * (3.141592! / 180!)
End Function

' Should be a macro (or two separate functions), since we can't do function overloading for float and int, and we don't want to use the float version for anything game-state related.
' 65536 / 360 = 8192 / 45, with a bit less overflow risk.
Public Function DEG(ByVal degrees As Long) As Long
    DEG = degrees * 8192 \ 45
End Function
