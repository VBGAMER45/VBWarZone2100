Attribute VB_Name = "plum_random_c"
' libplum: Copyright (c) 2020 Paul-Louis Ageneau; 2020-2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/random.c, random.h (the _WIN32 branch: BCryptGenRandom, rand() as the fallback). The mutex is left
' out (one thread). A buffer is a byte array and the offset of its first byte. plum_rand64 is not used.
Option Explicit

Private Declare Function BCryptGenRandom Lib "bcrypt" (ByVal hAlgorithm As Long, pbBuffer As Any, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = 2

' static int random_bytes(void *buf, size_t size)
Private Function random_bytes(buf() As Byte, ByVal offset As Long, ByVal size As Long) As Long
    Dim status As Long
    ' Requires Windows 7 or later
    status = BCryptGenRandom(0, buf(offset), size, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
    If status = 0 Then random_bytes = 0 Else random_bytes = -1
End Function

' static unsigned int generate_seed()
Private Function generate_seed() As Long
    generate_seed = GetTickCount()
End Function

' void plum_random_init()
Public Sub plum_random_init()
    ' mutex_init(&rand_mutex, 0): (VB) no mutex
    srand generate_seed()
End Sub

' void plum_random_cleanup()
Public Sub plum_random_cleanup()
    ' mutex_destroy(&rand_mutex): (VB) no mutex
End Sub

' void plum_random(void *buf, size_t size)
Public Sub plum_random(buf() As Byte, ByVal offset As Long, ByVal size As Long)
    Dim i As Long
    If random_bytes(buf, offset, size) = 0 Then Exit Sub

    ' RAND_MAX is guaranteed to be at least 2^15 - 1
    For i = 0 To size - 1
        buf(offset + i) = (rand() And &H7F80&) \ &H80&
    Next
End Sub

' uint32_t plum_rand32(void): the bits of the uint32_t (U32 for its value)
Public Function plum_rand32() As Long
    Dim r(0 To 3) As Byte, v As Long
    plum_random r, 0, 4
    CopyMemory v, r(0), 4
    plum_rand32 = v
End Function
