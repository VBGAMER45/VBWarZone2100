Attribute VB_Name = "difficulty_c"
' src/difficulty.cpp, difficulty.h: Handes the difficulty level effects on
' gameplay (separate modifiers for enemy and player damage).
Option Explicit

Public Enum DIFFICULTY_LEVEL
    DL_SUPER_EASY
    DL_EASY
    DL_NORMAL
    DL_HARD
    DL_INSANE
End Enum

' ------------------------------------------------------------------------------------

Private presDifLevel As DIFFICULTY_LEVEL

Private fDifPlayerModifier As Long
Private fDifEnemyModifier As Long

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub difficulty_c_init()
    presDifLevel = DL_NORMAL
End Sub

Public Sub setDamageModifiers(ByVal playerModifier As Long, ByVal enemyModifier As Long)
    fDifPlayerModifier = MAX_(1, playerModifier)
    If Not bMultiPlayer And getCamTweakOption_PS1Modifiers() Then
        fDifEnemyModifier = MAX_(1, enemyModifier \ 3)
    Else
        fDifEnemyModifier = MAX_(1, enemyModifier)
    End If
End Sub

' ------------------------------------------------------------------------------------
' Sets the game difficulty level
Public Sub setDifficultyLevel(ByVal lev As DIFFICULTY_LEVEL)
    Select Case lev
    Case DL_SUPER_EASY
        setDamageModifiers 150, 70
    Case DL_EASY
        setDamageModifiers 125, 85
    Case DL_NORMAL
        setDamageModifiers 100, 100
    Case DL_HARD
        setDamageModifiers 90, 110
    Case DL_INSANE
        setDamageModifiers 80, 120
    End Select
    presDifLevel = lev
End Sub

' ------------------------------------------------------------------------------------
' Returns the difficulty level
Public Function getDifficultyLevel() As DIFFICULTY_LEVEL
    getDifficultyLevel = presDifLevel
End Function

Public Function modifyForDifficultyLevel(ByVal basicVal As Long, ByVal IsPlayer As Boolean) As Long
    If IsPlayer Then
        modifyForDifficultyLevel = basicVal * fDifPlayerModifier \ 100
    Else
        modifyForDifficultyLevel = basicVal * fDifEnemyModifier \ 100
    End If
End Function

' reset damage modifiers changed by "double up" or "biffer baker" cheat and
' prevent campaign difficulty from influencing skirmish and multiplayer games
Public Sub resetDamageModifiers()
    If bMultiPlayer Then
        setDamageModifiers 100, 100
    Else
        setDifficultyLevel getDifficultyLevel()
    End If
End Sub
