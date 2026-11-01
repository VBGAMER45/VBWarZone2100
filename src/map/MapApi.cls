VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapApi"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' WzMap.dll's public interface. GlobalMultiUse: the exe and the other DLLs
' call these procedures as if they were their own. lib/wzmaplib is ported into
' this DLL's modules; this class only forwards to them. The library's
' classes (Map, MapData, Structure, Droid, Feature, TerrainTypeData,
' ObjectVector) and interfaces (LoggingProtocol, IOProvider,
' BinaryIOStream) are public classes of the DLL; the enums of map.h,
' map_types.h and terrain_type.h are here.
Option Explicit

' map_types.h: enum class OutputFormat (LatestOutputFormat = VER3)
Public Enum OutputFormat
    OutputFormat_VER1_BINARY_OLD = 0    ' FlaME-compatible / binary .bjo
    OutputFormat_VER2 = 1               ' JSON format (WZ 3.4+?)
    OutputFormat_VER3 = 2               ' JSONv2 format + full-range game.map (WZ 4.1+)
End Enum

' map_types.h: enum class MapType
Public Enum MapType
    MapType_CAMPAIGN = 0
    MapType_SAVEGAME = 1
    MapType_SKIRMISH = 2
End Enum

' map.h: enum class Map::LoadedFormat, Map Format information (for a loaded map)
Public Enum LoadedFormat
    LoadedFormat_MIXED = 0              ' Doesn't fully fit a map format - may have a mix of files from different formats / versions (loadable, but not ideal)
    LoadedFormat_BINARY_OLD = 1         ' Old binary .bjo
    LoadedFormat_JSON_v1 = 2            ' JSON (v1) format (WZ 3.4+?)
    LoadedFormat_SCRIPT_GENERATED = 3   ' Script-generated map (WZ 4.0+)
    LoadedFormat_JSON_v2 = 4            ' JSON (v2) format + full-range game.map (WZ 4.1+)
End Enum

' terrain_type.h: The different types of terrain as far as the game is concerned
Public Enum TYPE_OF_TERRAIN
    TER_SAND = 0
    TER_SANDYBRUSH = 1
    TER_BAKEDEARTH = 2
    TER_GREENMUD = 3
    TER_REDBRUSH = 4
    TER_PINKROCK = 5
    TER_ROAD = 6
    TER_WATER = 7
    TER_CLIFFFACE = 8
    TER_RUBBLE = 9
    TER_SHEETICE = 10
    TER_SLUSH = 11
    TER_MAX = 12
End Enum

' terrain_type.h: enum class MAP_TILESET
Public Enum MAP_TILESET
    MAP_TILESET_ARIZONA = 0
    MAP_TILESET_URBAN = 1
    MAP_TILESET_ROCKIES = 2
End Enum

' map_package.h: enum class LevelFormat
Public Enum LevelFormat
    LevelFormat_LEV = 0                 ' the old (classic) .lev file format
    LevelFormat_JSON = 1                ' the new (in-folder) level.json format
End Enum

' map_package.h: enum class MapPackage::MapPackageType
Public Enum MapPackageType
    MapPackageType_Map_Plain = 0        ' a package that contains only a map
    MapPackageType_Map_Mod = 1          ' a package that contains a map and (potentially) other modifications to core game files / textures / data (i.e. a "map mod")
End Enum

' map_package.h: enum MapPackage::ModTypes (flags)
Public Enum ModTypes
    ModTypes_GameModels = 1             ' .pie models (for components, features, structs, misc)
    ModTypes_Effects = 2                ' .pie models (for effects)
    ModTypes_Shaders = 4                ' shaders
    ModTypes_Stats = 8                  ' stats modifications
    ModTypes_Textures = 16              ' texpages
    ModTypes_Scripts = 32               ' script / rule modifications
    ModTypes_SkirmishAI = 64            ' skirmish AI modifications
    ModTypes_Messages = 128             ' messages
    ModTypes_MiscUI = 256               ' fonts, icons, images
    ModTypes_Sound = 512                ' audio, sequenceaudio
    ModTypes_Tilesets = 1024            ' tileset
    ModTypes_Datasets = 2048            ' wrf
    ModTypes_LastModType = 2048         ' static constexpr ModTypes LastModType = ModTypes::Datasets
End Enum

' map.h: Various defines needed by both maplib and the game's map-handling code
Public Enum map_h_defines
    MAP_MAXWIDTH = 256
    MAP_MAXHEIGHT = 256
    MAP_MAXAREA = 65536                 ' (256*256)
    ELEVATION_SCALE = 2                 ' Multiplier for the tile height - used for older map formats (which stored tile height as a byte)
    TILE_MAX_HEIGHT = 510               ' (255 * ELEVATION_SCALE)
    TILE_MIN_HEIGHT = 0
    PLAYER_SCAVENGERS = -1              ' The player number for scavengers (as loaded by this library)
    TILE_XFLIP = &H8000&                ' Flags for whether texture tiles are flipped in X and Y or rotated
    TILE_YFLIP = &H4000&
    TILE_ROTMASK = &H3000&
    TILE_ROTSHIFT = 12
    TILE_TRIFLIP = &H800&               ' This bit describes the direction the tile is split into 2 triangles (same as triangleFlip)
    TILE_NUMMASK = &H1FF&
    TILE_SHIFT = 7                      ' The shift on a world coordinate to get the tile coordinate
    TILE_MASK = &H7F&                   ' The mask to get internal tile coords from a full coordinate
    TILE_UNITS = 128                    ' The number of units across a tile (1<<TILE_SHIFT)
End Enum

' Name and version of this DLL, for the log header.
Public Function MapBuildInfo() As String
    MapBuildInfo = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
End Function

' map_version.h
Public Function wzmaplib_version_string() As String
    wzmaplib_version_string = map_version_c.wzmaplib_version_string()
End Function

Public Function wzmaplib_version_major() As Long
    wzmaplib_version_major = map_version_c.wzmaplib_version_major()
End Function

Public Function wzmaplib_version_minor() As Long
    wzmaplib_version_minor = map_version_c.wzmaplib_version_minor()
End Function

Public Function wzmaplib_version_rev() As Long
    wzmaplib_version_rev = map_version_c.wzmaplib_version_rev()
End Function

' map.h: class Map's constructor Map() (an empty Map, for modification)
Public Function Map_new() As Map
    Set Map_new = map_c.Map_new()
End Function

' map.h: Map::loadFromPath; mapIO omitted: a new StdIOProvider
Public Function Map_loadFromPath(ByVal mapFolderPath As String, ByVal mapType_ As MapType, ByVal mapMaxPlayers As Long, ByVal seed As Long, Optional ByVal logger As LoggingProtocol = Nothing, Optional ByVal mapIO As Variant) As Map
    Set Map_loadFromPath = map_c.Map_loadFromPath(mapFolderPath, mapType_, mapMaxPlayers, seed, logger, mapIO)
End Function

' map.h: Map::exportMapToPath; mapIO omitted: a new StdIOProvider
Public Function Map_exportMapToPath(ByVal map_ As Map, ByVal mapFolderPath As String, ByVal mapType_ As MapType, ByVal mapMaxPlayers As Long, ByVal format As OutputFormat, Optional ByVal logger As LoggingProtocol = Nothing, Optional ByVal mapIO As Variant) As Boolean
    Map_exportMapToPath = map_c.Map_exportMapToPath(map_, mapFolderPath, mapType_, mapMaxPlayers, format, logger, mapIO)
End Function

' map.h: Map::expectedFileNames(optional<LoadedFormat>): the names in names(0 .. count - 1)
Public Function Map_expectedFileNames(names() As String, Optional ByVal hasFormat As Boolean = False, Optional ByVal format As LoadedFormat = LoadedFormat_MIXED) As Long
    Map_expectedFileNames = map_c.Map_expectedFileNames(names, hasFormat, format)
End Function

' map.h: loadMapData (Nothing when it fails)
Public Function loadMapData(ByVal mapFile As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As MapData
    Set loadMapData = map_c.loadMapData(mapFile, mapIO, pCustomLogger)
End Function

' map.h: writeMapData
Public Function writeMapData(ByVal map_ As MapData, ByVal filename As String, ByVal mapIO As IOProvider, ByVal format As OutputFormat, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    writeMapData = map_c.writeMapData(map_, filename, mapIO, format, pCustomLogger)
End Function

' map_terrain_types.h: loadTerrainTypes (Nothing when it fails)
Public Function loadTerrainTypes(ByVal filename As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As TerrainTypeData
    Set loadTerrainTypes = map_terrain_types_c.loadTerrainTypes(filename, mapIO, pCustomLogger)
End Function

' (VB) WzMap::TerrainTypeData's construction, for the exe (the class is PublicNotCreatable): game.cpp's
' writeTerrainTypeMapFile fills one
Public Function TerrainTypeData_new() As TerrainTypeData
    Set TerrainTypeData_new = New TerrainTypeData
End Function

' map_terrain_types.h: writeTerrainTypes
Public Function writeTerrainTypes(ByVal ttypeData As TerrainTypeData, ByVal filename As String, ByVal mapIO As IOProvider, ByVal format As OutputFormat, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    writeTerrainTypes = map_terrain_types_c.writeTerrainTypes(ttypeData, filename, mapIO, format, pCustomLogger)
End Function

' map_io.h: a new StdIOProvider (the default IOProvider)
Public Function StdIOProvider_new() As IOProvider
    Set StdIOProvider_new = New StdIOProvider
End Function

' map_io.h: IOProvider::pathJoin, pathDirName, pathBaseName
Public Function IOProvider_pathJoin(ByVal io As IOProvider, ByVal a As String, ByVal b As String) As String
    IOProvider_pathJoin = map_io_c.IOProvider_pathJoin(io, a, b)
End Function

Public Function IOProvider_pathDirName(ByVal io As IOProvider, ByVal fullPath As String) As String
    IOProvider_pathDirName = map_io_c.IOProvider_pathDirName(io, fullPath)
End Function

Public Function IOProvider_pathBaseName(ByVal io As IOProvider, ByVal fullPath As String) As String
    IOProvider_pathBaseName = map_io_c.IOProvider_pathBaseName(io, fullPath)
End Function

' map_package.h: trimTechLevelFromMapName (in place)
Public Sub trimTechLevelFromMapName(mapName As String)
    map_package_c.trimTechLevelFromMapName mapName
End Sub

' map_package.h: MapPackage::loadPackage (pMapIO Nothing: a StdIOProvider); Nothing when it fails
Public Function MapPackage_loadPackage(ByVal pathToMapPackage As String, Optional ByVal logger As LoggingProtocol = Nothing, Optional ByVal pMapIO As IOProvider = Nothing) As MapPackage
    Set MapPackage_loadPackage = map_package_c.MapPackage_loadPackage(pathToMapPackage, logger, pMapIO)
End Function

' map_package.h: loadLevelDetails, loadLevelDetails_LEV, loadLevelDetails_JSON (Nothing: nullopt)
Public Function loadLevelDetails(ByVal levelFile As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As LevelDetails
    Set loadLevelDetails = map_package_c.loadLevelDetails(levelFile, mapIO, pCustomLogger)
End Function

Public Function loadLevelDetails_LEV(ByVal levelFile As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As LevelDetails
    Set loadLevelDetails_LEV = map_package_c.loadLevelDetails_LEV(levelFile, mapIO, pCustomLogger)
End Function

Public Function loadLevelDetails_JSON(ByVal levelJsonFile As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As LevelDetails
    Set loadLevelDetails_JSON = map_package_c.loadLevelDetails_JSON(levelJsonFile, mapIO, pCustomLogger)
End Function

' map_package.h: loadGamFile (Nothing: nullopt)
Public Function loadGamFile(ByVal gamFile As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As GamInfo
    Set loadGamFile = map_package_c.loadGamFile(gamFile, mapIO, pCustomLogger)
End Function
