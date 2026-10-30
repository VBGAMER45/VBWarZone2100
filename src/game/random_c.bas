Attribute VB_Name = "random_c"
' src/random.cpp, random.h: the game's pseudorandom numbers. A
' MersenneTwister is a UDT; MersenneTwister_construct is its constructor
' (seed 42 by default). random_c_init constructs the game's generator (the
' C++ static initialiser). The uint32_t arithmetic goes through math_vb
' (MUL32, SHRU, SHL, U32).
Option Explicit

' Pseudorandom number generator with 19937 bit state and period of 2**19937 - 1. Equidistributed in up to 623 dimensions.
' See http://en.wikipedia.org/wiki/Mersenne_twister
' If all clients use the same seed, they will generate the same pseudorandom number sequence.
Public Type MersenneTwister
    offset As Long                      ' Index in state.
    state(0 To 623) As Long             ' 19937 bit state, rounded up to next 32 bits.
End Type

Private gamePseudorandomNumberGenerator As MersenneTwister
Private lastSeed As Long

Public Sub random_c_init()
    MersenneTwister_construct gamePseudorandomNumberGenerator
End Sub

Public Sub MersenneTwister_construct(mt As MersenneTwister, Optional ByVal seed As Long = 42)
    Dim i As Long
    mt.offset = 624
    ' Set state to something pseudorandom. Exact constants aren't important here.
    mt.state(0) = seed
    For i = 1 To 623
        mt.state(i) = FromU32(U32(MUL32(&H6C078965, mt.state(i - 1) Xor SHRU(mt.state(i - 1), 30))) + i)
    Next
End Sub

' Generates a random number in the interval [0...UINT32_MAX].
Public Function MersenneTwister_u32(mt As MersenneTwister) As Long
    Dim ret As Long
    If mt.offset = 624 Then
        MersenneTwister_generate mt
    End If

    ret = mt.state(mt.offset)
    mt.offset = mt.offset + 1

    ' Give better than 624 dimensional equidistribution for lower bits.
    ' Don't change the constants unless you know what you're doing.
    ' Can skip this step if it noticeably affects performance (which it probably doesn't), since we aren't doing Monte-Carlo simulations.
    ret = ret Xor SHRU(ret, 11)
    ret = ret Xor (SHL(ret, 7) And &H9D2C5680)
    ret = ret Xor (SHL(ret, 15) And &HEFC60000)
    ret = ret Xor SHRU(ret, 18)

    MersenneTwister_u32 = ret
End Function

' Generates more random numbers.
Private Sub MersenneTwister_generate(mt As MersenneTwister)
    Dim i As Long, v As Long
    mt.offset = 0

    ' Don't change these constants, unless you have verified that the period is still 2**19937 - 1.
    ' Discards the lower 31 bits of state[0], since the state is 19937 = 624*32 - 31 bits.
    ' Loop tripled, to avoid using %624 everywhere.
    For i = 0 To 226
        v = (mt.state(i) And &H80000000) Or (mt.state(i + 1) And &H7FFFFFFF)
        mt.state(i) = mt.state(i + 397) Xor SHRU(v, 1) Xor ((v And 1) * &H9908B0DF)
    Next
    For i = 227 To 622
        v = (mt.state(i) And &H80000000) Or (mt.state(i + 1) And &H7FFFFFFF)
        mt.state(i) = mt.state(i + 397 - 624) Xor SHRU(v, 1) Xor ((v And 1) * &H9908B0DF)
    Next
    For i = 623 To 623                  ' Very short loop.
        v = (mt.state(i) And &H80000000) Or (mt.state(i + 1 - 624) And &H7FFFFFFF)
        mt.state(i) = mt.state(i + 397 - 624) Xor SHRU(v, 1) Xor ((v And 1) * &H9908B0DF)
    Next
End Sub

' Seeds the random number generator. The seed is sent over the network, such that all clients generate the same number sequence, without the number sequence being the same each game.
Public Sub gameSRand(ByVal seed As Long)
    lastSeed = seed
    MersenneTwister_construct gamePseudorandomNumberGenerator, seed
End Sub

' Get the seed used to seed the game's random number generator
Public Function gameRand_GetSeed() As Long
    gameRand_GetSeed = lastSeed
End Function

' Generates a random number in the interval [0...UINT32_MAX].
' Must not be called from graphics routines, only for making game decisions.
Public Function gameRandU32() As Long
    gameRandU32 = MersenneTwister_u32(gamePseudorandomNumberGenerator)
End Function

' Generates a random number in the interval [0...limit - 1]. Not equidistributed for large non-powers of 2.
' Must not be called from graphics routines, only for making game decisions.
Public Function gameRand(ByVal limit As Long) As Long
    Dim u As Double
    syncDebug "gameRand", "Used a random number."

    If limit = 0 Then
        gameRand = 0
        Exit Function
    End If

    u = U32(MersenneTwister_u32(gamePseudorandomNumberGenerator))
    gameRand = FromU32(u - Int(u / U32(limit)) * U32(limit))
End Function
