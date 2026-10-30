VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScoreDataToScreenCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/scores.h: struct ScoreDataToScreenCache, the pUserData of the mission result screen's back
' form (mission.cpp's intDisplayMissionBackDrop: scores.cpp's scoreDataToScreen draws with it). A
' new one is C++'s new ScoreDataToScreenCache(); the form's onDelete drops it (pUserData = Empty).
Option Explicit

Public wzLabelText_UnitLosses As WzText
Public wzLabelText_StructureLosses As WzText
Public wzLabelText_ForceInformation As WzText

Public wzInfoBarText As Collection      ' std::vector<WzText>: its WzTexts in order

Public wzInfoText_ArtifactsFound As WzText
Public wzInfoText_MissionTime As WzText
Public wzInfoText_TotalGameTime As WzText
Public wzInfoText_Cheated As WzText

Private Sub Class_Initialize()
    Set wzLabelText_UnitLosses = New WzText
    Set wzLabelText_StructureLosses = New WzText
    Set wzLabelText_ForceInformation = New WzText
    Set wzInfoBarText = New Collection
    Set wzInfoText_ArtifactsFound = New WzText
    Set wzInfoText_MissionTime = New WzText
    Set wzInfoText_TotalGameTime = New WzText
    Set wzInfoText_Cheated = New WzText
End Sub
