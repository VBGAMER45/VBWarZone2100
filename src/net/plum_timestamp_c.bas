Attribute VB_Name = "plum_timestamp_c"
' libplum: Copyright (c) 2020 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/timestamp.c, timestamp.h (the _WIN32 branch): timestamp_t and timediff_t (int64_t) are Doubles
' (milliseconds).
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

' timestamp_t current_timestamp()
Public Function current_timestamp() As Double
    current_timestamp = U32(GetTickCount())         ' (timestamp_t)GetTickCount()
End Function
