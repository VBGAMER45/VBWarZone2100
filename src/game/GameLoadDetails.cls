VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GameLoadDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/game.h, src/game.cpp: struct GameLoadDetails, what to load (a save game, a
' level file or a map package). Made by GameLoadDetails_makeUserSaveGameLoad,
' _makeMapPackageLoad and _makeLevelFileLoad (game_c; the constructor is
' Class_Initialize plus those). A map package is a .wz archive (WzMapZipIO)
' read by WzMap's MapPackage; getGamInfoFromPackage's GamInfo is an Object.
Option Explicit

Public loadType As Long                 ' GameLoadType
Public filePath As String

Private m_logger As LoggingProtocol     ' std::shared_ptr<WzMap::LoggingProtocol>
Private m_loadedPackage As MapPackage   ' mutable std::shared_ptr<WzMap::MapPackage>

Private Sub Class_Initialize()
    Set m_logger = New WzMapDebugLogger
End Sub

Public Function setLogger(ByVal logger As LoggingProtocol) As GameLoadDetails
    Set m_logger = logger
    Set setLogger = Me
End Function

Public Function getMapFolderPath() As String
    Dim result As String
    result = filePath
    If strEndsWith(result, ".gam") Then
        result = Left$(result, Len(result) - 4)
    End If
    If Len(result) = 0 Then
        result = result & "/"
    ElseIf Right$(result, 1) <> "/" Then
        result = result & "/"
    End If
    getMapFolderPath = result
End Function

Public Function getMap(ByVal mapSeed As Long) As Map
    Dim package As Object
    Select Case loadType
        Case GameLoadType_UserSaveGame
            Set getMap = Map_loadFromPath(getMapFolderPath(), getWzMapType(True), game.maxPlayers, mapSeed, m_logger, New WzMapPhysFSIO)
            Exit Function
        Case GameLoadType_MapPackage
            Set package = getMapPackage()
            If package Is Nothing Then
                Set getMap = Nothing
                Exit Function
            End If
            Set getMap = m_loadedPackage.loadMap(mapSeed)
            Exit Function
        Case GameLoadType_Level
            Set getMap = Map_loadFromPath(getMapFolderPath(), getWzMapType(False), game.maxPlayers, mapSeed, m_logger, New WzMapPhysFSIO)
            Exit Function
    End Select

    Set getMap = Nothing                ' silence compiler warning
End Function

' the package of a .wz map (Nothing: nullptr)
Private Function getMapPackage() As MapPackage
    Dim mapZipIO As WzMapZipIO, zipReadSource As WzZipIOPHYSFSSourceReadProvider
    If Not (loadType = GameLoadType_MapPackage) Then ASSERT_FAILURE "loadType == GameLoadType::MapPackage", "src/game.cpp:2609", "getMapPackage", "Not a map package": Set getMapPackage = Nothing: Exit Function
    If Not m_loadedPackage Is Nothing Then
        Set getMapPackage = m_loadedPackage
        Exit Function
    End If
    ' std::shared_ptr<WzMapZipIO> mapZipIO = getSpecialInMemoryMapArchive(filePath): (VB) the in-memory map
    ' archive is not ported: there is none
    If mapZipIO Is Nothing Then
        Set zipReadSource = WzZipIOPHYSFSSourceReadProvider_make(filePath)
        Set mapZipIO = WzMapZipIO_openZipArchiveReadIOProvider(zipReadSource)
    End If
    If mapZipIO Is Nothing Then
        debug_ 2622, LOG_ERROR, "getMapPackage", "Failed to open map package: %s", filePath
        Set getMapPackage = Nothing
        Exit Function
    End If
    Set m_loadedPackage = MapPackage_loadPackage("", m_logger, mapZipIO)
    Set getMapPackage = m_loadedPackage
End Function

' const WzMap::GamInfo* (an Object; Nothing = nullptr)
Public Function getGamInfoFromPackage() As Object
    Dim package As Object
    Select Case loadType
        Case GameLoadType_UserSaveGame
            Set getGamInfoFromPackage = Nothing
            Exit Function
        Case GameLoadType_MapPackage
            Set package = getMapPackage()
            If package Is Nothing Then
                Set getGamInfoFromPackage = Nothing
                Exit Function
            End If
            Set getGamInfoFromPackage = package.getGamInfo()
            Exit Function
        Case GameLoadType_Level
            Set getGamInfoFromPackage = Nothing
            Exit Function
    End Select
    Set getGamInfoFromPackage = Nothing ' silence compiler warning
End Function
