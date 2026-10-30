VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/include/wzmaplib/map.h, lib/wzmaplib/src/map.cpp: class
' WzMap::Map, a map folder whose files load on first request (mapData(),
' mapStructures()...). Made by Map_loadFromPath, or Map_new for an empty Map
' (MapApi; the static members are map_c's Map_<name>). shared_ptr members
' are object references (Nothing for nullptr); the vectors of objects are
' ObjectVectors. m_fileVersions (an unordered_map keyed by MapFile) is an
' array with an entry per MapFile. optional<> results are a Boolean with
' the value ByRef. Missing: calculateMapStats (map_stats.cpp).
Option Explicit

' (enum class LoadedFormat is MapApi's)

' struct LoadedFileVersion: enum FileType
Private Const FileType_BinaryBJO As Long = 0
Private Const FileType_JSON As Long = 1
Private Const FileType_ScriptGenerated As Long = 2

' enum class MapFile
Private Const MapFile_MapData As Long = 0
Private Const MapFile_Structures As Long = 1
Private Const MapFile_Droids As Long = 2
Private Const MapFile_Features As Long = 3
Private Const MapFile_TerrainTypes As Long = 4

Private Const VERSION_40 As Long = 40

Private m_mapFolderPath As String
Private m_mapType As MapType
Private m_mapMaxPlayers As Long
Private m_logger As LoggingProtocol
Private m_mapIO As IOProvider
Private m_scriptGeneratedMapSeed As Long
Private m_wasScriptGenerated As Boolean
Private m_mapScriptContents() As Byte        ' std::shared_ptr<std::vector<char>>
Private m_mapScriptContentsSize As Long
Private m_hasMapScriptContents As Boolean

Private m_mapData As MapData
Private m_structures As ObjectVector
Private m_droids As ObjectVector
Private m_features As ObjectVector
Private m_terrainTypes As TerrainTypeData

' m_fileVersions: for each MapFile, whether it has an entry, and the entry (LoadedFileVersion: fileType, fileVersion)
Private m_hasFileVersion(0 To 4) As Boolean
Private m_fileType(0 To 4) As Long
Private m_fileVersion(0 To 4) As Long

Private Sub Class_Initialize()
    m_mapMaxPlayers = 8
End Sub

' Construct an empty Map, for modification
Friend Sub construct_empty()
    Set m_mapData = New MapData
    Set m_structures = New ObjectVector
    Set m_droids = New ObjectVector
    Set m_features = New ObjectVector
    Set m_terrainTypes = New TerrainTypeData
End Sub

' Map(mapFolderPath, mapType, mapMaxPlayers, logger, mapIO)
Friend Sub construct(ByVal mapFolderPath As String, ByVal mapType_ As MapType, ByVal mapMaxPlayers As Long, ByVal logger As LoggingProtocol, ByVal mapIO As IOProvider)
    m_mapFolderPath = mapFolderPath
    m_mapType = mapType_
    m_mapMaxPlayers = mapMaxPlayers
    Set m_logger = logger
    Set m_mapIO = mapIO
End Sub

' What loadFromPath and generateFromExistingScriptMap set on the Map a script returned
Friend Sub setScriptGenerated(ByVal mapFolderPath As String, ByVal mapType_ As MapType, ByVal mapMaxPlayers As Long, ByVal mapIO As IOProvider, ByVal logger As LoggingProtocol, ByVal seed As Long, contents() As Byte, ByVal contentsSize As Long, ByVal terrainTypes As TerrainTypeData)
    m_mapFolderPath = mapFolderPath
    m_mapType = mapType_
    m_mapMaxPlayers = mapMaxPlayers
    Set m_mapIO = mapIO
    Set m_logger = logger
    Set m_terrainTypes = terrainTypes          ' Nothing: the ttypes.ttp file is loaded on-demand
    m_scriptGeneratedMapSeed = seed
    m_wasScriptGenerated = True
    If contentsSize > 0 Then m_mapScriptContents = contents
    m_mapScriptContentsSize = contentsSize
    m_hasMapScriptContents = True
End Sub

Private Sub setFileVersion(ByVal file As Long, ByVal fileType As Long, ByVal version As Long)
    m_hasFileVersion(file) = True
    m_fileType(file) = fileType
    m_fileVersion(file) = version
End Sub

' Get the map data
' Returns Nothing if constructed for loading and the loading failed
Public Function mapData() As MapData
    Dim loadResult As MapData, version As Long
    If Not m_mapData Is Nothing Then
        Set mapData = m_mapData
        Exit Function
    End If

    ' otherwise, load the map data on first request
    If m_mapIO Is Nothing Then Exit Function
    If loadMapData_Internal(IOProvider_pathJoin(m_mapIO, m_mapFolderPath, "game.map"), m_mapIO, m_logger, loadResult, version) Then
        Set m_mapData = loadResult
        setFileVersion MapFile_MapData, FileType_BinaryBJO, version
        Set mapData = m_mapData
    End If
End Function

' Get the structures
' Returns Nothing if constructed for loading and the loading failed
Public Function mapStructures() As ObjectVector
    Dim fileType As Long, ok As Boolean, objects As ObjectVector, version As Long
    If Not m_structures Is Nothing Then
        Set mapStructures = m_structures
        Exit Function
    End If

    ' Otherwise, load the data on first request
    If m_mapIO Is Nothing Then Exit Function
    ' Try JSON first
    fileType = FileType_JSON
    ok = loadJsonStructureInit(IOProvider_pathJoin(m_mapIO, m_mapFolderPath, "struct.json"), m_mapType, m_mapIO, m_logger, objects, version)
    If Not ok Then
        ' Fallback to .bjo (old binary format)
        fileType = FileType_BinaryBJO
        ok = loadBJOStructureInit(IOProvider_pathJoin(m_mapIO, m_mapFolderPath, "struct.bjo"), m_mapMaxPlayers, m_mapIO, m_logger, objects, version)
    End If
    If ok Then
        Set m_structures = objects
        setFileVersion MapFile_Structures, fileType, version
        Set mapStructures = m_structures
    End If
End Function

' Get the droids
' Returns Nothing if constructed for loading and the loading failed
Public Function mapDroids() As ObjectVector
    Dim fileType As Long, ok As Boolean, objects As ObjectVector, version As Long
    If Not m_droids Is Nothing Then
        Set mapDroids = m_droids
        Exit Function
    End If

    ' Otherwise, load the data on first request
    If m_mapIO Is Nothing Then Exit Function
    ' Try JSON first
    fileType = FileType_JSON
    ok = loadJsonDroidInit(IOProvider_pathJoin(m_mapIO, m_mapFolderPath, "droid.json"), m_mapType, m_mapIO, m_logger, objects, version)
    If Not ok Then
        ' Fallback to .bjo (old binary format)
        fileType = FileType_BinaryBJO
        ok = loadBJODroidInit(IOProvider_pathJoin(m_mapIO, m_mapFolderPath, "dinit.bjo"), m_mapMaxPlayers, m_mapIO, m_logger, objects, version)
    End If
    If ok Then
        Set m_droids = objects
        setFileVersion MapFile_Droids, fileType, version
        Set mapDroids = m_droids
    End If
End Function

' Get the features
' Returns Nothing if constructed for loading and the loading failed
Public Function mapFeatures() As ObjectVector
    Dim fileType As Long, ok As Boolean, objects As ObjectVector, version As Long
    If Not m_features Is Nothing Then
        Set mapFeatures = m_features
        Exit Function
    End If

    ' Otherwise, load the data on first request
    If m_mapIO Is Nothing Then Exit Function
    ' Try JSON first
    fileType = FileType_JSON
    ok = loadJsonFeatureInit(IOProvider_pathJoin(m_mapIO, m_mapFolderPath, "feature.json"), m_mapType, m_mapIO, m_logger, objects, version)
    If Not ok Then
        ' Fallback to .bjo (old binary format)
        fileType = FileType_BinaryBJO
        ok = loadBJOFeatureInit(IOProvider_pathJoin(m_mapIO, m_mapFolderPath, "feat.bjo"), m_mapMaxPlayers, m_mapIO, m_logger, objects, version)
    End If
    If ok Then
        Set m_features = objects
        setFileVersion MapFile_Features, fileType, version
        Set mapFeatures = m_features
    End If
End Function

' Get the terrain type map
Public Function mapTerrainTypes() As TerrainTypeData
    Dim loadResult As TerrainTypeData
    If Not m_terrainTypes Is Nothing Then
        Set mapTerrainTypes = m_terrainTypes
        Exit Function
    End If

    ' Otherwise, load the data on first request
    If m_mapIO Is Nothing Then Exit Function
    Set loadResult = loadTerrainTypes(IOProvider_pathJoin(m_mapIO, m_mapFolderPath, "ttypes.ttp"), m_mapIO, m_logger)
    If Not loadResult Is Nothing Then
        Set m_terrainTypes = loadResult
        Set mapTerrainTypes = m_terrainTypes
    End If
End Function

' Obtaining CRC values
Public Function crcSumMapTiles(ByVal crc As Long) As Long
    Dim pMapData As MapData
    Set pMapData = mapData()
    If pMapData Is Nothing Then
        crcSumMapTiles = crc
        Exit Function
    End If
    crcSumMapTiles = pMapData.crcSumMapTiles(crc)
End Function

' The CRC steps of the members the three kinds of object share: id (when
' set), name, position, direction
Private Function crcSumObjectBase(ByVal crc As Long, ByVal o As Object) As Long
    Dim v(0 To 0) As Long, b() As Byte, n As Long
    If o.hasId Then
        v(0) = o.id
        crc = crcSumU32(crc, v, 0, 1)
    End If
    n = utf8_encode(o.name_, b)
    If n > 0 Then crc = crcSum(crc, b, 0, n)
    crc = crcSumWorldPos(crc, o.position_x, o.position_y)
    v(0) = o.direction
    crc = crcSumU16(crc, v, 0, 1)
    crcSumObjectBase = crc
End Function

Public Function crcSumStructures(ByVal crc As Long) As Long
    Dim pStructures As ObjectVector, i As Long, o As Structure, b(0 To 0) As Byte
    Set pStructures = mapStructures()
    If pStructures Is Nothing Then
        crcSumStructures = crc
        Exit Function
    End If
    For i = 0 To pStructures.size() - 1
        Set o = pStructures.at(i)
        crc = crcSumObjectBase(crc, o)
        b(0) = o.modules And &HFF&
        crc = crcSum(crc, b, 0, 1)
        b(0) = o.player And &HFF&
        crc = crcSum(crc, b, 0, 1)
    Next
    crcSumStructures = crc
End Function

Public Function crcSumDroids(ByVal crc As Long) As Long
    Dim pDroids As ObjectVector, i As Long, o As Droid, b(0 To 0) As Byte
    Set pDroids = mapDroids()
    If pDroids Is Nothing Then
        crcSumDroids = crc
        Exit Function
    End If
    For i = 0 To pDroids.size() - 1
        Set o = pDroids.at(i)
        crc = crcSumObjectBase(crc, o)
        b(0) = o.player And &HFF&
        crc = crcSum(crc, b, 0, 1)
    Next
    crcSumDroids = crc
End Function

Public Function crcSumFeatures(ByVal crc As Long) As Long
    Dim pFeatures As ObjectVector, i As Long, o As Feature, b(0 To 0) As Byte
    Set pFeatures = mapFeatures()
    If pFeatures Is Nothing Then
        crcSumFeatures = crc
        Exit Function
    End If
    For i = 0 To pFeatures.size() - 1
        Set o = pFeatures.at(i)
        crc = crcSumObjectBase(crc, o)
        If o.hasPlayer Then
            b(0) = o.player And &HFF&
            crc = crcSum(crc, b, 0, 1)
        End If
    Next
    crcSumFeatures = crc
End Function

' Obtain the map format of a loaded map
' Returns false if loading the map data failed, the format can't be determined, or if this WzMap instance was not created by loading a map
Public Function loadedMapFormat(format As LoadedFormat) As Boolean
    Dim numFilesOfType(0 To 2) As Long, hasNumFilesOfType(0 To 2) As Boolean, numTypes As Long
    Dim lastJSONFileVersion As Long, hasLastJSONFileVersion As Boolean, file As Long, fileType As Long, t As Long

    ' make sure we have all the pieces of a valid map
    If mapData() Is Nothing Then Exit Function
    If mapTerrainTypes() Is Nothing Then Exit Function
    If mapDroids() Is Nothing Then Exit Function
    If mapFeatures() Is Nothing Then Exit Function
    If mapStructures() Is Nothing Then Exit Function

    If m_wasScriptGenerated Then
        format = LoadedFormat_SCRIPT_GENERATED
        loadedMapFormat = True
        Exit Function
    End If

    ' process the loaded map object file versions
    For file = 0 To 4
        If Not m_hasFileVersion(file) Then GoTo continue_
        ' skip MapData and TerrainData, which are currently always binary file types
        If file = MapFile_MapData Or file = MapFile_TerrainTypes Then GoTo continue_
        fileType = m_fileType(file)
        numFilesOfType(fileType) = numFilesOfType(fileType) + 1
        hasNumFilesOfType(fileType) = True
        ' check if there are any other file types
        For t = 0 To 2
            If hasNumFilesOfType(t) And t <> fileType And numFilesOfType(t) > 0 Then
                format = LoadedFormat_MIXED
                loadedMapFormat = True
                Exit Function
            End If
        Next
        If fileType = FileType_JSON Then
            ' Need to check for consistent version
            If hasLastJSONFileVersion Then
                If lastJSONFileVersion <> m_fileVersion(file) Then
                    ' File version is different from all prior file versions - JSON maps are expected to have files of the same version
                    format = LoadedFormat_MIXED
                    loadedMapFormat = True
                    Exit Function
                End If
            Else
                lastJSONFileVersion = m_fileVersion(file)
                hasLastJSONFileVersion = True
            End If
        End If
continue_:
    Next

    ' if we reached here, we should have all map object files of the same FileType
    For t = 0 To 2
        If hasNumFilesOfType(t) Then
            numTypes = numTypes + 1
            fileType = t
        End If
    Next
    If numTypes <> 1 Then
        ' or not? (presumably this isn't a loaded map)
        Exit Function
    End If

    Select Case fileType
    Case FileType_BinaryBJO
        ' for now, just return that everything is in binary file format
        format = LoadedFormat_BINARY_OLD
        loadedMapFormat = True
    Case FileType_JSON
        If Not hasLastJSONFileVersion Then
            ' should not happen
            Exit Function
        End If
        If Not m_hasFileVersion(MapFile_MapData) Then
            ' should not happen
            Exit Function
        End If
        If lastJSONFileVersion = 1 Then
            If m_fileType(MapFile_MapData) = FileType_BinaryBJO And ULT(m_fileVersion(MapFile_MapData), VERSION_40) Then format = LoadedFormat_JSON_v1 Else format = LoadedFormat_MIXED
            loadedMapFormat = True
        ElseIf lastJSONFileVersion = 2 Then
            If m_fileType(MapFile_MapData) = FileType_BinaryBJO And m_fileVersion(MapFile_MapData) = VERSION_40 Then format = LoadedFormat_JSON_v2 Else format = LoadedFormat_MIXED
            loadedMapFormat = True
        Else
            ' unknown JSON file version?
            Exit Function
        End If
    Case FileType_ScriptGenerated
        format = LoadedFormat_SCRIPT_GENERATED
        loadedMapFormat = True
    End Select
End Function

' Obtain the script from a loaded script-generated map: false for nullptr
Public Function scriptMapContents(contents() As Byte, contentsSize As Long) As Boolean
    If Not m_hasMapScriptContents Then Exit Function
    If m_mapScriptContentsSize > 0 Then contents = m_mapScriptContents
    contentsSize = m_mapScriptContentsSize
    scriptMapContents = True
End Function

' Returns a new Map instance, from an existing script-generated map instance, generated using a new seed
' Fails if:
'   - The current Map is not a script-generated map (wasScriptGenerated() == false)
Public Function generateFromExistingScriptMap(ByVal seed As Long, Optional ByVal logger As LoggingProtocol = Nothing) As Map
    Dim gameJSPath As String, result As Map, terrainTypes As TerrainTypeData
    If Not m_wasScriptGenerated Or Not m_hasMapScriptContents Then
        Exit Function
    End If

    gameJSPath = IOProvider_pathJoin(m_mapIO, m_mapFolderPath, "game.js")
    ' Load script map, which actually loads everything
    Set result = runMapScript(m_mapScriptContents, m_mapScriptContentsSize, gameJSPath, seed, False, logger)
    If Not result Is Nothing Then
        If Not m_terrainTypes Is Nothing Then
            ' deep-copy over the current terrainTypes data
            Set terrainTypes = m_terrainTypes.clone()
        Else
            Set terrainTypes = Nothing  ' clear this so the ttypes.ttp file is loaded on-demand (if possible)
        End If
        result.setScriptGenerated m_mapFolderPath, m_mapType, m_mapMaxPlayers, m_mapIO, logger, seed, m_mapScriptContents, m_mapScriptContentsSize, terrainTypes
    End If
    Set generateFromExistingScriptMap = result
End Function

' Other Map instance data
Public Function wasScriptGenerated() As Boolean
    wasScriptGenerated = m_wasScriptGenerated
End Function

Public Function scriptGeneratedMapSeed(seed As Long) As Boolean
    If Not m_wasScriptGenerated Then
        Exit Function
    End If
    seed = m_scriptGeneratedMapSeed
    scriptGeneratedMapSeed = True
End Function

Public Function mapFolderPath() As String
    mapFolderPath = m_mapFolderPath
End Function
