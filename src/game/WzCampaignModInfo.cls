VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignModInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/modinfo.h: class WzCampaignModInfo : public WZModInfo (WZModInfo's fields first; WZModInfo has no other
' subclass in the game, so the two are one class). modinfo_c (modinfo.cpp) fills it: loadCampaignModInfoFromFile;
' getModCompatibility is modinfo.cpp's WZModInfo::getModCompatibility. A WzJsonLocalizedString is its
' getLocalizedString() (see modinfo_c); optional<std::string> is has<name> and <name>; the
' std::unordered_set<DIFFICULTY_LEVEL> is difficultyLevels ("d<level>" -> level); the std::vectors of
' WzCampaignTweakOption are Collections of WzCampaignTweakOption objects.
Option Explicit

Public name_ As String                     ' std::string name
Public author As String
Public description As String               ' WzJsonLocalizedString
Public type_ As WzModType
Public license As String
Public hasVersion As Boolean               ' optional<std::string> version
Public version As String
Public hasUpdatesUrl As Boolean            ' optional<std::string> updatesUrl
Public updatesUrl As String
' compatibility
Public minVersionSupported As String
Public maxVersionTested As String
' from mod file itself
Public modFilename As String
Public modPath As String
Private modBannerPNGData_() As Byte        ' std::vector<unsigned char> modBannerPNGData
Private modBannerPNGData_size As Long
' WzCampaignModInfo
Private campaignFiles_() As CAMPAIGN_FILE  ' std::vector<CAMPAIGN_FILE> campaignFiles
Public campaignFiles_size As Long
Public difficultyLevels As Collection      ' std::unordered_set<DIFFICULTY_LEVEL>
Public camTweakOptions As Collection       ' std::vector<WzCampaignTweakOption>
Public customTweakOptions As Collection    ' std::vector<WzCampaignTweakOption>
Public universe As WzCampaignUniverse      ' = WzCampaignUniverse::WZ2100Extended

Private Sub Class_Initialize()
    Set difficultyLevels = New Collection
    Set camTweakOptions = New Collection
    Set customTweakOptions = New Collection
    universe = WzCampaignUniverse_WZ2100Extended
End Sub

' WzModCompatibilityResult WZModInfo::getModCompatibility() const (modinfo.cpp)
Public Function getModCompatibility() As WzModCompatibilityResult
    Dim self As WzCampaignModInfo
    Set self = Me
    getModCompatibility = isWZModVersionCompatible(self)
End Function

' difficultyLevels.count(lvl) > 0
Public Function hasDifficultyLevel(ByVal lvl As Long) As Boolean
    Dim v As Variant
    On Error GoTo missing
    v = difficultyLevels.Item("d" & lvl)
    hasDifficultyLevel = True
    Exit Function
missing:
    hasDifficultyLevel = False
End Function

' modBannerPNGData: its bytes in out(0 .. n - 1); n
Friend Function getModBannerPNGData(out() As Byte) As Long
    If modBannerPNGData_size > 0 Then out = modBannerPNGData_
    getModBannerPNGData = modBannerPNGData_size
End Function

Friend Sub setModBannerPNGData(data() As Byte, ByVal n As Long)
    Dim i As Long
    modBannerPNGData_size = n
    If n > 0 Then
        ReDim modBannerPNGData_(0 To n - 1)
        For i = 0 To n - 1
            modBannerPNGData_(i) = data(i)
        Next
    Else
        Erase modBannerPNGData_
    End If
End Sub

' campaignFiles: in out(0 .. n - 1); n
Friend Function getCampaignFiles(out() As CAMPAIGN_FILE) As Long
    Dim i As Long
    If campaignFiles_size > 0 Then
        ReDim out(0 To campaignFiles_size - 1)
        For i = 0 To campaignFiles_size - 1
            out(i) = campaignFiles_(i)
        Next
    End If
    getCampaignFiles = campaignFiles_size
End Function

Friend Sub setCampaignFiles(list() As CAMPAIGN_FILE, ByVal n As Long)
    Dim i As Long
    campaignFiles_size = n
    If n > 0 Then
        ReDim campaignFiles_(0 To n - 1)
        For i = 0 To n - 1
            campaignFiles_(i) = list(i)
        Next
    Else
        Erase campaignFiles_
    End If
End Sub
