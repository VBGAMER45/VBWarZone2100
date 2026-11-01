VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/src/map_package.cpp, include/wzmaplib/map_package.h: class WzMap::MapPackage, a map package (a map
' folder with its level details: the new level.json format, or a classic <map>.addon.lev / .xplayers.lev file)
' read through an IOProvider. Made by MapApi.MapPackage_loadPackage (map_package_c); a std::shared_ptr<MapPackage>
' is a MapPackage object (Nothing: nullptr). The optionals are a value with a has-flag. The std::function
' enumerations (modTypesEnumerate) return the mod types in types(0 .. count - 1).
' Missing (no caller in the game): the constructor from a LevelDetails and a Map, exportMapPackageFiles,
' calculateMapStats (map_stats.cpp).
Option Explicit

Private m_pathToMapPackage As String
Private m_loadedLevelFormat As LevelFormat
Private m_hasLoadedLevelFormat As Boolean
Private m_levelDetails As LevelDetails
Private m_gamInfo As GamInfo                    ' optional<GamInfo> (Nothing: nullopt)
Private m_mapType As MapType
Private m_packageType As MapPackageType
Private m_hasPackageType As Boolean
Private m_flatMapPackage As Boolean
Private m_modTypes As Long                      ' optional<uint64_t> (the flags fit a Long)
Private m_hasModTypes As Boolean
Private m_originalGenerator As String
Private m_hasOriginalGenerator As Boolean
Private m_mapIO As IOProvider                   ' std::shared_ptr<WzMap::IOProvider>
Private m_logger As LoggingProtocol             ' std::shared_ptr<LoggingProtocol>
Private m_loadedMap As Map                      ' std::shared_ptr<Map>

Private Sub Class_Initialize()
    Set m_levelDetails = New LevelDetails
End Sub

' (VB) what loadPackage sets on the new package (result->m_...)
Friend Sub initLoaded(ByVal details As LevelDetails, ByVal format As LevelFormat, ByVal pathToMapPackage As String, _
        ByVal mapIO As IOProvider, ByVal logger As LoggingProtocol, ByVal flatMapPackage As Boolean)
    m_levelDetails.copyFrom details
    m_loadedLevelFormat = format
    m_hasLoadedLevelFormat = True
    m_mapType = m_levelDetails.type_
    m_pathToMapPackage = pathToMapPackage
    Set m_mapIO = mapIO
    Set m_logger = logger
    m_flatMapPackage = flatMapPackage
End Sub

Friend Sub setOriginalGenerator(ByVal originalGenerator As String)
    m_originalGenerator = originalGenerator
    m_hasOriginalGenerator = True
End Sub

Public Function levelDetails() As LevelDetails
    Set levelDetails = m_levelDetails
End Function

Public Sub updateLevelDetails(ByVal newLevelDetails As LevelDetails)
    m_levelDetails.copyFrom newLevelDetails
End Sub

' Get the loaded level details format
' Note: Returns a value only if the MapPackage was loaded (i.e. via loadPackage)
Public Function loadedLevelDetailsFormat(format As LevelFormat) As Boolean
    format = m_loadedLevelFormat
    loadedLevelDetailsFormat = m_hasLoadedLevelFormat
End Function

Public Function isScriptGeneratedMap() As Boolean
    Dim fullPathToMapFolder As String, gameJSPath As String
    If Not m_loadedMap Is Nothing Then
        ' get this from the loaded map
        isScriptGeneratedMap = m_loadedMap.wasScriptGenerated()
        Exit Function
    End If

    ' No map is loaded yet - verify this was instantiated with a map IO provider
    If m_mapIO Is Nothing Then
        ' Should not happen
        printLog_ m_logger, 1609, LogLevel_Error, "isScriptGeneratedMap", "Missing IOProvider - no map can be loaded"
        isScriptGeneratedMap = False
        Exit Function
    End If

    ' Instead of calling Map::loadFromPath (like MapPackage::loadMap does), which would load and run the map script (if present), check for existence of the game.js file
    fullPathToMapFolder = IOProvider_pathJoin(m_mapIO, m_pathToMapPackage, m_levelDetails.mapFolderPath)
    gameJSPath = IOProvider_pathJoin(m_mapIO, fullPathToMapFolder, "game.js")
    isScriptGeneratedMap = m_mapIO.fileExists(gameJSPath)
End Function

' Get the map data
' Returns nullptr if the loading failed
Public Function loadMap(ByVal seed As Long, Optional ByVal logger As LoggingProtocol = Nothing) As Map
    Dim fullPathToMapFolder As String, scriptSeed As Long
    If Not m_loadedMap Is Nothing Then
        If Not m_loadedMap.scriptGeneratedMapSeed(scriptSeed) Then scriptSeed = seed    ' value_or(seed)
        If scriptSeed = seed Then
            Set loadMap = m_loadedMap
            Exit Function
        Else
            ' Request for a map load of a script-generated map with a different seed
            ' This must generate a new Map instance
            If Not m_loadedMap.wasScriptGenerated() Then
                ' Should not happen
                printLog_ logger, 1636, LogLevel_Error, "loadMap", "Map has a script generation seed (so supposedly script-generated), but flag for wasScriptGenerated is false??"
                Set loadMap = Nothing
                Exit Function
            End If
            Set loadMap = m_loadedMap.generateFromExistingScriptMap(seed, logger)
            Exit Function
        End If
    Else
        ' No map is loaded yet - verify this was instantiated with a map IO provider
        If m_mapIO Is Nothing Then
            ' Should not happen
            printLog_ logger, 1648, LogLevel_Error, "loadMap", "Missing IOProvider - no map can be loaded"
            Set loadMap = Nothing
            Exit Function
        End If
    End If

    fullPathToMapFolder = IOProvider_pathJoin(m_mapIO, m_pathToMapPackage, m_levelDetails.mapFolderPath)
    Set m_loadedMap = Map_loadFromPath(fullPathToMapFolder, m_levelDetails.type_, m_levelDetails.players, seed, logger, m_mapIO)
    Set loadMap = m_loadedMap
End Function

Public Function packageType() As MapPackageType
    Dim foundUnapprovedPath As Boolean, multiplay As String, maps As String, pathSeparator As String
    Dim names() As String, n As Long, i As Long, enumSuccess As Boolean
    If m_hasPackageType Then
        packageType = m_packageType
        Exit Function
    End If

    If m_hasModTypes And m_modTypes <> 0 Then
        setPackageType MapPackageType_Map_Mod
        packageType = m_packageType
        Exit Function
    End If

    If m_mapIO Is Nothing Then
        setPackageType MapPackageType_Map_Plain
        packageType = m_packageType
        Exit Function
    End If

    ' First, check if this is a "flat" modern map package
    ' i.e. if it has "level.json" in its root
    ' (if so, we can skip all of the older checks because it won't be mounted in a path that conflicts with core game files)
    If m_flatMapPackage Then
        setPackageType MapPackageType_Map_Plain
        packageType = m_packageType
        Exit Function
    End If

    ' otherwise, check package (older checks, non-flat map packages)
    multiplay = "multiplay"
    maps = "maps"
    pathSeparator = m_mapIO.pathSeparator()

    ' - First pass (root) looks for any folder that isn't "multiplay"
    enumSuccess = m_mapIO.enumerateFolders(m_pathToMapPackage, names, n)
    For i = 0 To n - 1
        If Len(names(i)) > 0 Then
            If Not (names(i) = multiplay Or names(i) = multiplay & pathSeparator) Then
                foundUnapprovedPath = True
                Exit For                        ' found a non-map path
            End If
        End If
    Next
    If foundUnapprovedPath Then
        setPackageType MapPackageType_Map_Mod
        packageType = m_packageType
        Exit Function
    End If
    If Not enumSuccess Then
        ' failed to enumerate??
        printLog_ m_logger, 1745, LogLevel_Error, "packageType", "Failed enumerating package"
    End If

    ' - Second pass (multiplay) looks for any folder that isn't "maps"
    enumSuccess = m_mapIO.enumerateFolders(IOProvider_pathJoin(m_mapIO, m_pathToMapPackage, multiplay), names, n)
    For i = 0 To n - 1
        If Len(names(i)) > 0 Then
            If Not (names(i) = maps Or names(i) = maps & pathSeparator) Then
                foundUnapprovedPath = True
                Exit For                        ' found a non-map path
            End If
        End If
    Next
    If foundUnapprovedPath Then
        setPackageType MapPackageType_Map_Mod
        packageType = m_packageType
        Exit Function
    End If
    If Not enumSuccess Then
        ' failed to enumerate??
        printLog_ m_logger, 1778, LogLevel_Error, "packageType", "Failed enumerating package (2)"
    End If

    ' - Third pass (root) look for specific files that might affect game behavior
    enumSuccess = m_mapIO.enumerateFiles(m_pathToMapPackage, names, n)
    For i = 0 To n - 1
        If names(i) = "gamedesc.lev" Or names(i) = "palette.txt" Then
            ' Found one!
            foundUnapprovedPath = True
            Exit For
        End If
    Next
    If foundUnapprovedPath Then
        setPackageType MapPackageType_Map_Mod
        packageType = m_packageType
        Exit Function
    End If
    If Not enumSuccess Then
        ' failed to enumerate??
        printLog_ m_logger, 1799, LogLevel_Error, "packageType", "Failed enumerating package (3)"
    End If

    setPackageType MapPackageType_Map_Plain
    packageType = m_packageType
End Function

Private Sub setPackageType(ByVal t As MapPackageType)
    m_packageType = t
    m_hasPackageType = True
End Sub

Public Function isFlatMapPackage() As Boolean
    isFlatMapPackage = m_flatMapPackage
End Function

' (#define dirCmp(a, b) safeDirectoryPathCompare(a, b, strlen(b), pathSeparator))
Private Function dirCmp(ByVal folderPath As String, ByVal desiredPath As String, ByVal pathSeparator As String) As Boolean
    dirCmp = safeDirectoryPathCompare(folderPath, desiredPath, Len(desiredPath), pathSeparator)
End Function

Public Function baseModificationTypes() As Long
    Dim modTypesValue As Long, pathSeparator As String, names() As String, n As Long, i As Long, enumSuccess As Boolean
    Dim name_ As String
    If m_hasModTypes Then
        baseModificationTypes = m_modTypes
        Exit Function
    End If

    If m_hasPackageType And m_packageType = MapPackageType_Map_Plain Then
        m_modTypes = 0
        m_hasModTypes = True
        baseModificationTypes = 0
        Exit Function
    End If

    If m_mapIO Is Nothing Then
        m_modTypes = 0
        m_hasModTypes = True
        baseModificationTypes = 0
        Exit Function
    End If

    modTypesValue = 0

    pathSeparator = m_mapIO.pathSeparator()

    ' - First pass looks in the root folder
    enumSuccess = m_mapIO.enumerateFolders(m_pathToMapPackage, names, n)
    For i = 0 To n - 1
        name_ = names(i)
        If Len(name_) = 0 Then GoTo continue1_
        If dirCmp(name_, "components", pathSeparator) Or dirCmp(name_, "features", pathSeparator) Or dirCmp(name_, "structs", pathSeparator) Or dirCmp(name_, "misc", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_GameModels
        ElseIf dirCmp(name_, "effects", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Effects
        ElseIf dirCmp(name_, "shaders", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Shaders
        ElseIf dirCmp(name_, "stats", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Stats
        ElseIf dirCmp(name_, "texpages", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Textures
        ElseIf dirCmp(name_, "script", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Scripts
        ElseIf dirCmp(name_, "messages", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Messages
        ElseIf dirCmp(name_, "fonts", pathSeparator) Or dirCmp(name_, "icons", pathSeparator) Or dirCmp(name_, "images", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_MiscUI
        ElseIf dirCmp(name_, "audio", pathSeparator) Or dirCmp(name_, "sequenceaudio", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Sound
        ElseIf dirCmp(name_, "tileset", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Tilesets
        ElseIf dirCmp(name_, "wrf", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Datasets
        End If
continue1_:
    Next
    If Not enumSuccess Then
        ' failed to enumerate??
        printLog_ m_logger, 1896, LogLevel_Error, "baseModificationTypes", "Failed enumerating package"
    End If

    ' - Second pass (multiplay folder)
    enumSuccess = m_mapIO.enumerateFolders(IOProvider_pathJoin(m_mapIO, m_pathToMapPackage, "multiplay"), names, n)
    For i = 0 To n - 1
        name_ = names(i)
        If Len(name_) = 0 Then GoTo continue2_
        If dirCmp(name_, "script", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_Scripts
        ElseIf dirCmp(name_, "skirmish", pathSeparator) Then
            modTypesValue = modTypesValue Or ModTypes_SkirmishAI
        End If
continue2_:
    Next
    If Not enumSuccess Then
        ' failed to enumerate??
        printLog_ m_logger, 1920, LogLevel_Error, "baseModificationTypes", "Failed enumerating package (2)"
    End If

    m_modTypes = modTypesValue
    m_hasModTypes = True
    baseModificationTypes = modTypesValue
End Function

' modTypesEnumerate(func) over baseModificationTypes(): the types in types(0 .. count - 1); False when there are none
Public Function modTypesEnumerate(types() As Long, count As Long) As Boolean
    modTypesEnumerate = modTypesEnumerateValue(baseModificationTypes(), types, count)
End Function

Public Function modTypesEnumerateValue(ByVal modTypesValue As Long, types() As Long, count As Long) As Boolean
    Dim i As Long, modTypeCurrent As Long
    count = 0
    If modTypesValue = 0 Then
        modTypesEnumerateValue = False
        Exit Function
    End If
    i = 0
    Do While (2 ^ i) <= ModTypes_LastModType
        modTypeCurrent = 2 ^ i
        If (modTypesValue And modTypeCurrent) = modTypeCurrent Then
            ReDim Preserve types(0 To count)
            types(count) = modTypeCurrent
            count = count + 1
        End If
        i = i + 1
    Loop
    modTypesEnumerateValue = True
End Function

' NOTE: There may be clients that rely on the specific output for each of these ModTypes
' DO NOT EDIT the string output for existing entries!!
Public Function to_string(ByVal modType As ModTypes) As String
    Select Case modType
        Case ModTypes_GameModels: to_string = "gamemodels"
        Case ModTypes_Effects: to_string = "effects"
        Case ModTypes_Shaders: to_string = "shaders"
        Case ModTypes_Stats: to_string = "stats"
        Case ModTypes_Textures: to_string = "textures"
        Case ModTypes_Scripts: to_string = "scripts"
        Case ModTypes_SkirmishAI: to_string = "skirmishai"
        Case ModTypes_Messages: to_string = "messages"
        Case ModTypes_MiscUI: to_string = "miscui"
        Case ModTypes_Sound: to_string = "sound"
        Case ModTypes_Tilesets: to_string = "tilesets"
        Case ModTypes_Datasets: to_string = "datasets"
        Case Else: to_string = ""      ' silence warning
    End Select
End Function

Public Function getGamInfo() As GamInfo
    Dim fullPathToMapFolder As String, loadedGamInfo As GamInfo
    If Not m_gamInfo Is Nothing Then
        Set getGamInfo = m_gamInfo
        Exit Function
    End If

    If m_mapIO Is Nothing Then
        Set m_gamInfo = New GamInfo
        Set getGamInfo = m_gamInfo
        Exit Function
    End If

    fullPathToMapFolder = IOProvider_pathJoin(m_mapIO, m_pathToMapPackage, m_levelDetails.mapFolderPath)
    Set loadedGamInfo = loadGamFile_JSON(IOProvider_pathJoin(m_mapIO, fullPathToMapFolder, "gam.json"), m_mapIO, m_logger)
    If loadedGamInfo Is Nothing Then
        ' fall back to older <map>.gam file
        Set loadedGamInfo = loadGamFile_OldBinary(IOProvider_pathJoin(m_mapIO, m_pathToMapPackage, m_levelDetails.gamFilePath()), m_mapIO, m_logger)
    End If
    If loadedGamInfo Is Nothing Then
        Set loadedGamInfo = New GamInfo
    End If

    Set m_gamInfo = loadedGamInfo
    Set getGamInfo = m_gamInfo
End Function
