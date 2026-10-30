Attribute VB_Name = "game_c"
' src/game.cpp, src/game.h, src/gamedef.h (ported in part): the start of a
' game from a level, a replay or a save game's .gam (loadGameInit, gameLoad,
' deserializeSaveGameHeader, readGamJson, parseJsonFile, gameLoadV7 with
' deserializeSaveGameV7Data_json, loadRulesetJson, loadMissionExtras),
' loadGame whole, setMapScroll, getPlayerNames, allocatePlayers,
' getWzMapType, getSaveGameType, the GameLoadDetails factories (the class is
' GameLoadDetails.cls), the terrain type overrides of the built-in maps
' (shouldLoadTerrainTypeOverrides, loadTerrainTypeMapOverride) and the map's
' objects (loadWzMapFeature, loadWzMapStructure, loadWzMapDroidInit,
' RemapWzMapPlayerNumber).
' std::array<std::unordered_map<UDWORD, UDWORD>, MAX_PLAYER_SLOTS>&
' moduleToBuilding is an array of Scripting.Dictionary objects, and
' std::unordered_map<UDWORD, UDWORD>& fixedMapIdToGeneratedId a
' Scripting.Dictionary of Longs; loadGame's std::map<WzString,
' PerPlayer...Lists *> is a Dictionary of the lists' numbers (LISTS_...),
' iterated in key order. optional<nlohmann::json> is a Boolean (has_value)
' and the value. SAVE_GAME (saveGameData) has the fields loadGame reads.
' gameLoadV: the .gam of version 39 (its gam.json, or the binary of the
' (de)serialisation API); the raw binary structs of versions 10 to 34 (Warzone
' 2.x saves) are refused (4.7.0 reads its SAVE_GAME_V33 and V34 raw over a
' NETPLAY and a MULTIPLAYERGAME that hold std containers now). game_main_c has
' the SAVE_GAME versions' (de)serialisers, the main file (loadMainFile,
' writeMainFile), saveGame and the writers; game_objects_c the object files.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' gamedef.h: the version number used in save games
Private Const VERSION_7 As Long = 7     ' string ID names saved for comps an objects
Private Const VERSION_8 As Long = 8     ' research status saved out for user saved games
Private Const VERSION_9 As Long = 9     ' power and experience saved out for user saved games
Private Const VERSION_10 As Long = 10   ' includes gateways and zones in .map file.
Private Const VERSION_11 As Long = 11   ' newstyle save game with extending structure checked in 13 Nov.
Private Const VERSION_12 As Long = 12   ' mission and order stuff checked in 20 Nov.
Private Const VERSION_14 As Long = 14
Private Const VERSION_15 As Long = 15
Private Const VERSION_16 As Long = 16   ' beta save game
Private Const VERSION_17 As Long = 17   ' objId and new struct stats included
Private Const VERSION_18 As Long = 18   ' droid name savegame validity stamps
Private Const VERSION_19 As Long = 19   ' alliances, colours, radar zoom
Private Const VERSION_20 As Long = 20   ' MAX_NAME_ SIZE replaced by MAX_SAVE_NAME_SIZE
Private Const VERSION_21 As Long = 21   ' timeStartHold saved out for research facilities
Private Const VERSION_22 As Long = 22   ' asRundData
Private Const VERSION_23 As Long = 23   ' limbo droids and camstart droids saved properly, no cluster save
Private Const VERSION_24 As Long = 24   ' reinforceTime, droid move, droid resistance
Private Const VERSION_25 As Long = 25   ' limbo droid, research module hold fixed, cleaned by Alex
Private Const VERSION_26 As Long = 26   ' reArm pads
Private Const VERSION_27 As Long = 27   ' unit not the "d" word, experience and repair facility currentPtsAdded
Private Const VERSION_28 As Long = 28   ' rearm pads currentPtsAdded saved
Private Const VERSION_29 As Long = 29   ' mission scroll limits saved
Private Const VERSION_30 As Long = 30   ' script external variables saved
Private Const VERSION_31 As Long = 31   ' mission cheat time saved
Private Const VERSION_32 As Long = 32   ' factory secondary order saved
Private Const VERSION_33 As Long = 33   ' skirmish save
Private Const VERSION_34 As Long = 34   ' saves AI names for multiplayer (.gam file)
Private Const VERSION_35 As Long = 35   ' uses the (de)serialization API for saving/loading games and is big-endian instead of little-endian
Private Const VERSION_36 As Long = 36   ' saves beacon properly
Private Const VERSION_38 As Long = 38   ' mod list!
Private Const VERSION_39 As Long = 39   ' lots of changes, breaking everything

Private Const CURRENT_VERSION_NUM As Long = VERSION_39

Private Const SAVEKEY_ONMISSION As Long = &H100

'used in the loadGame
Public Const KEEPOBJECTS As Boolean = True
Public Const FREEMEM As Boolean = True

' general save definitions
Private Const MAX_LEVEL_SIZE As Long = 20
Private Const SAVE_GAME_V7_SIZE As Long = 4& * 6 + MAX_LEVEL_SIZE   ' sizeof(SAVE_GAME_V7)

' (VB) the lists a loadGame std::map<WzString, PerPlayer...Lists *> points to
Private Const LISTS_MAIN As Long = 0    ' apsDroidLists / apsStructLists
Private Const LISTS_MISSION As Long = 1 ' mission.apsDroidLists / mission.apsStructLists
Private Const LISTS_LIMBO As Long = 2   ' apsLimboDroids

Private Type GAME_SAVEHEADER
    aFileType(0 To 3) As Byte           ' char aFileType[4]
    version As Long                     ' uint32_t
End Type

Private Type SAVE_GAME_V7
    gameTime As Long                    ' uint32_t
    GameType As Long                    ' uint32_t: Type of game , one of the GTYPE_... enums.
    ScrollMinX As Long                  ' int32_t: Scroll Limits
    ScrollMinY As Long
    ScrollMaxX As Long                  ' uint32_t
    ScrollMaxY As Long
    levelName As String                 ' char levelName[MAX_LEVEL_SIZE]: name of the level to load up when mid game
End Type

' (SAVE_GAME, its PLAYER and NETPLAY: game_main_c's)

Private saveGameVersion As Long         ' UDWORD
Private saveGameOnMission As Boolean
Private saveGameData As SAVE_GAME          ' (game_main_c's Type)

Private savedGameTime As Long           ' UDWORD
Private savedObjId As Long              ' UDWORD
Private startX As Long, startY As Long  ' SDWORD
Private width_ As Long, height_ As Long ' UDWORD width, height
Private gameType As Long                ' GAME_TYPE
Private IsScenario As Boolean

' Maps built into the game with custom tile types need to be excluded for overrides.
Public Function shouldLoadTerrainTypeOverrides(ByVal name_ As String) As Boolean
    Dim excludedMaps As Variant, i As Long
    excludedMaps = Array("Monocot")

    For i = 0 To UBound(excludedMaps)
        If excludedMaps(i) = name_ Then
            shouldLoadTerrainTypeOverrides = False
            Exit Function
        End If
    Next
    shouldLoadTerrainTypeOverrides = True
End Function

Public Function loadTerrainTypeMapOverride(ByVal tileSet As MAP_TILESET) As Boolean
    Dim iniName As String, ini As WzConfig, tileTypeKey As String, list() As String, n As Long, i As Long
    Dim counter As Long, keys() As String, nKeys As Long, j As Long, tileType As Long

    resForceBaseDir "/data/base/"
    iniName = "tileset/tileTypes.json"
    If PHYSFS_exists(iniName) = 0 Then
        loadTerrainTypeMapOverride = False
        Exit Function
    End If

    Set ini = WzConfig_new(iniName, WzConfig_ReadOnly)

    Select Case tileSet
        Case MAP_TILESET_ARIZONA
            tileTypeKey = "Arizona"
        Case MAP_TILESET_URBAN
            tileTypeKey = "Urban"
        Case MAP_TILESET_ROCKIES
            tileTypeKey = "Rockies"
'		default:
'			debug(LOG_ERROR, "Unknown tile type");
'			resForceBaseDir("");
'			return false;
    End Select

    n = ini.childGroups(list)
    For i = 0 To n - 1
        If list(i) = tileTypeKey Then
            ini.beginGroup list(i)

            debug_ 7378, LOG_TERRAIN, "loadTerrainTypeMapOverride", "Looking at tileset type: %s", tileTypeKey
            counter = 0

            nKeys = ini.childKeys(keys)
            For j = 0 To nKeys - 1
                tileType = ini.value(keys(j)).toUInt()

                If ULT(TER_MAX, tileType) Then
                    debug_ 7388, LOG_ERROR, "loadTerrainTypeMapOverride", "loadTerrainTypeMapOverride: terrain type out of range"
                    resForceBaseDir ""
                    loadTerrainTypeMapOverride = False
                    Exit Function
                End If
                ' Workaround for fugly map editor bug, since we can't fix the map editor
                If ULT(MAX_TILE_TEXTURES - 1, counter) Then
                    debug_ 7395, LOG_ERROR, "loadTerrainTypeMapOverride", "loadTerrainTypeMapOverride: too many textures!"
                    resForceBaseDir ""
                    loadTerrainTypeMapOverride = False
                    Exit Function
                End If
                ' Log the output for the override value.
                If terrainTypes(counter) <> tileType Then
                    debug_ 7402, LOG_TERRAIN, "loadTerrainTypeMapOverride", "Upgrading map tile %d (type %d) to type %d", counter, CLng(terrainTypes(counter)), tileType
                End If
                terrainTypes(counter) = tileType And &HFF&
                counter = counter + 1
                debug_ 7406, LOG_TERRAIN, "loadTerrainTypeMapOverride", "Tile %d at value: %d", counter - 1, tileType
            Next
            ini.endGroup
        End If
    Next

    resForceBaseDir ""

    loadTerrainTypeMapOverride = True
End Function

' (VB) fixedMapIdToGeneratedId.insert(value_type(mapId, newID)).second
Private Function fixedMapIdInsert(ByVal fixedMapIdToGeneratedId As Object, ByVal mapId As Long, ByVal newID As Long) As Boolean
    If fixedMapIdToGeneratedId.Exists(mapId) Then Exit Function
    fixedMapIdToGeneratedId.Add mapId, newID
    fixedMapIdInsert = True
End Function

Public Function loadWzMapFeature(ByVal wzMap As Map, ByVal fixedMapIdToGeneratedId As Object) As Boolean
    Dim pFeatures As ObjectVector, i As Long, feature_ As WzMap.Feature, psStats As Long, pFeature As Long
    Dim newID As Long, addedMapping As Boolean

    Set pFeatures = wzMap.mapFeatures()
    If pFeatures Is Nothing Then
        loadWzMapFeature = False
        Exit Function
    End If

    For i = 0 To pFeatures.size() - 1
        Set feature_ = pFeatures.at(i)
        For psStats = 0 To asFeatureStats_size - 1
            If asFeatureStats(psStats).id = feature_.name_ Then Exit For
        Next
        If psStats = asFeatureStats_size Then
            debug_ 7104, LOG_ERROR, "loadWzMapFeature", "Feature type ""%s"" unknown", feature_.name_
            GoTo continue_  ' ignore this
        End If
        ' Create the Feature
        pFeature = 0
        'restore values && create Feature
        newID = generateSynchronisedObjectId()
        If feature_.hasId Then
            If feature_.id <> 0 Then   ' (uint32_t > 0)
                addedMapping = fixedMapIdInsert(fixedMapIdToGeneratedId, feature_.id, newID)
                If addedMapping Then
                    debug_ 7116, LOG_MAP, "loadWzMapFeature", "Fixed map object ID: %u -> generated id: %u", feature_.id, newID
                Else
                    debug_ 7120, LOG_ERROR, "loadWzMapFeature", "Found duplicate hard-coded object ID in map data: %u", feature_.id
                End If
            End If
        End If
        pFeature = buildFeature(psStats, feature_.position_x, feature_.position_y, True, newID)
        If pFeature = 0 Then
            debug_ 7126, LOG_ERROR, "loadWzMapFeature", "Unable to create feature %s", feature_.name_
            GoTo continue_
        End If
        If asFeatureStats(aFeature(pFeature).psStats).subType = FEAT_OIL_RESOURCE Then
            scriptSetDerrickPos aFeature(pFeature).pos.x, aFeature(pFeature).pos.y
        End If

        aFeature(pFeature).rot.direction = feature_.direction And &HFFFF&
        If feature_.hasPlayer Then
            aFeature(pFeature).player = feature_.player And &HFF&
        Else
            aFeature(pFeature).player = PLAYER_FEATURE
        End If
continue_:
    Next

    loadWzMapFeature = True
End Function

Private Function RemapWzMapPlayerNumber(ByVal oldNumber As Long) As Long
    Dim i As Long
    If oldNumber < 0 Then
        game.mapHasScavengers = True
        RemapWzMapPlayerNumber = scavengerSlot()
        Exit Function
    End If

    If game.typ = LEVEL_TYPE_CAMPAIGN Then  ' don't remap for SP games
        RemapWzMapPlayerNumber = oldNumber
        Exit Function
    End If

    For i = 0 To MAX_PLAYERS - 1
        If oldNumber = NetPlay_players_position(i) Then
            game.mapHasScavengers = game.mapHasScavengers Or i = scavengerSlot()
            RemapWzMapPlayerNumber = i
            Exit Function
        End If
    Next
    ASSERT_FAILURE "false", "src/game.cpp:5264", "RemapWzMapPlayerNumber", "Found no player position for player %d", oldNumber
    RemapWzMapPlayerNumber = 0
End Function

Public Function loadWzMapStructure(ByVal wzMap As Map, ByVal fixedMapIdToGeneratedId As Object, moduleToBuilding() As Object) As Boolean
    Dim NumberOfSkippedStructures As Long, pStructures As ObjectVector, k As Long, structure_ As WzMap.Structure
    Dim psStats As Long, psStructure As Long, player As Long, newID As Long, addedMapping As Boolean
    Dim moduleStat As Long, i As Long, idStr As Long

    NumberOfSkippedStructures = 0
    Set pStructures = wzMap.mapStructures()
    If pStructures Is Nothing Then
        loadWzMapStructure = False
        Exit Function
    End If

    For k = 0 To pStructures.size() - 1
        Set structure_ = pStructures.at(k)
        For psStats = 0 To numStructureStats - 1
            If asStructureStats(psStats).id = structure_.name_ Then Exit For
        Next
        If psStats = numStructureStats Then
            debug_ 6328, LOG_ERROR, "loadWzMapStructure", "Structure type ""%s"" unknown", structure_.name_
            GoTo continue_  ' ignore this
        End If
        'for modules - need to check the base structure exists
        If IsStatExpansionModule(psStats) Then
            psStructure = getTileStructure(map_coord(structure_.position_x), map_coord(structure_.position_y))
            If psStructure = 0 Then
                debug_ 6337, LOG_ERROR, "loadWzMapStructure", "No owning structure for module - %s for player - %d", structure_.name_, structure_.player
                GoTo continue_ ' ignore this module
            End If
        End If
        'check not trying to build too near the edge
        If map_coord(structure_.position_x) < TOO_NEAR_EDGE Or map_coord(structure_.position_x) > mapWidth - TOO_NEAR_EDGE _
         Or map_coord(structure_.position_y) < TOO_NEAR_EDGE Or map_coord(structure_.position_y) > mapHeight - TOO_NEAR_EDGE Then
            debug_ 6345, LOG_ERROR, "loadWzMapStructure", "Structure %s, coord too near the edge of the map", structure_.name_
            GoTo continue_ ' skip it
        End If
        player = RemapWzMapPlayerNumber(structure_.player)
        If Not ULT(player, MAX_PLAYERS) Then
            player = MAX_PLAYERS - 1
            NumberOfSkippedStructures = NumberOfSkippedStructures + 1
        End If
        psStructure = 0
        If structure_.hasId Then idStr = structure_.id Else idStr = -1
        debug_ 6355, LOG_NEVER, "loadWzMapStructure", "trying to build structure %i;%i;%s;%i;%i", idStr, player, _
               structure_.name_, map_coord(structure_.position_x), map_coord(structure_.position_y)

        newID = generateSynchronisedObjectId()
        If structure_.hasId Then
            If structure_.id <> 0 Then   ' (uint32_t > 0)
                addedMapping = fixedMapIdInsert(fixedMapIdToGeneratedId, structure_.id, newID)
                If addedMapping Then
                    debug_ 6364, LOG_MAP, "loadWzMapStructure", "Fixed map object ID: %u -> generated id: %u", structure_.id, newID
                Else
                    debug_ 6368, LOG_ERROR, "loadWzMapStructure", "Found duplicate hard-coded object ID in map data: %u", structure_.id
                End If
            End If
        End If
        psStructure = buildStructureDir(psStats, structure_.position_x, structure_.position_y, structure_.direction, player, True, newID)
        If psStructure = 0 Then
            debug_ 6374, LOG_ERROR, "loadWzMapStructure", "Structure %s couldn't be built (probably on top of another structure).", structure_.name_
            GoTo continue_
        End If

        If IsStatExpansionModule(psStats) Then
            ' backward compatibility:
            ' JS scripts try to find buildings by their modules ids, which isn't possible anymore
            ' (because SIMPLE_OBJECT.id is const),
            ' so we maintain a map from module_id to buildings, so that qtscripts.cpp::loadLabels
            ' can replace module_id with building_id
            ' *Not* doing so will break campaign (no artifacts being spawned from modules)
            If moduleToBuilding(player).Exists(newID) Then
                If structure_.hasId Then idStr = structure_.id Else idStr = newID
                ASSERT_FAILURE "mappingInsertionResult.second", "src/game.cpp:6387", "loadWzMapStructure", "Duplicate input module id: %u - ignoring mapping module to building", idStr
            Else
                moduleToBuilding(player).Add newID, aStruct(psStructure).id
            End If
        End If

        ' Previously, we would override building's ID with module's ID
        ' now, "id" is const, we can't do that.
        ' this may break some mods which look up structures by theirs module id.
        If structure_.modules > 0 Then
            moduleStat = getModuleStat(psStructure)
            If moduleStat = -1 Then
                debug_ 6398, LOG_ERROR, "loadWzMapStructure", "Structure %s can't have modules.", structure_.name_
                GoTo continue_
            End If
            For i = 0 To structure_.modules - 1
                buildStructure moduleStat, structure_.position_x, structure_.position_y, player, True
            Next
        End If
        buildingComplete psStructure
        If asStructureStats(aStruct(psStructure).pStructureType).typ = REF_HQ Then
            scriptSetStartPos player, aStruct(psStructure).pos.x, aStruct(psStructure).pos.y
        ElseIf asStructureStats(aStruct(psStructure).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then
            scriptSetDerrickPos aStruct(psStructure).pos.x, aStruct(psStructure).pos.y
        End If
continue_:
    Next

    If NumberOfSkippedStructures > 0 Then
        debug_ 6419, LOG_ERROR, "loadWzMapStructure", "structureLoad: invalid player number in %d structures ... assigned to the last player!\n\n", NumberOfSkippedStructures
        loadWzMapStructure = False
        Exit Function
    End If

    loadWzMapStructure = True
End Function

Public Function loadWzMapDroidInit(ByVal wzMap As Map, ByVal fixedMapIdToGeneratedId As Object) As Boolean
    Dim NumberOfSkippedDroids As Long, pDroids As ObjectVector, k As Long, droid_ As WzMap.Droid
    Dim player As Long, psTemplate As Long, psDroid As Long, newID As Long, addedMapping As Boolean
    Dim startpos As Vector2i

    NumberOfSkippedDroids = 0
    Set pDroids = wzMap.mapDroids()
    If Not (Not pDroids Is Nothing) Then ASSERT_FAILURE "pDroids.get() != nullptr", "src/game.cpp:5272", "loadWzMapDroidInit", "No data.": loadWzMapDroidInit = False: Exit Function

    For k = 0 To pDroids.size() - 1
        Set droid_ = pDroids.at(k)
        player = RemapWzMapPlayerNumber(droid_.player)
        If Not ULT(player, MAX_PLAYERS) Then
            player = MAX_PLAYERS - 1  ' now don't lose any droids ... force them to be the last player
            NumberOfSkippedDroids = NumberOfSkippedDroids + 1
        End If
        psTemplate = getTemplateFromTranslatedNameNoPlayer(droid_.name_)
        If psTemplate = 0 Then
            debug_ 5285, LOG_ERROR, "loadWzMapDroidInit", "Unable to find template for %s for player %d -- unit skipped", droid_.name_, player
            GoTo continue_
        End If
        turnOffMultiMsg True
        psDroid = 0
        newID = generateSynchronisedObjectId()
        If droid_.hasId Then
            If droid_.id <> 0 Then   ' (uint32_t > 0)
                addedMapping = fixedMapIdInsert(fixedMapIdToGeneratedId, droid_.id, newID)
                If addedMapping Then
                    debug_ 5296, LOG_MAP, "loadWzMapDroidInit", "Fixed map object ID: %u -> generated id: %u", droid_.id, newID
                Else
                    debug_ 5300, LOG_ERROR, "loadWzMapDroidInit", "Found duplicate hard-coded object ID in map data: %u", droid_.id
                End If
            End If
        End If
        psDroid = reallyBuildDroid(psTemplate, Vector3i_(droid_.position_x, droid_.position_y, 0), player, False, Rotation_(droid_.direction, 0, 0), newID)
        turnOffMultiMsg False
        If psDroid = 0 Then
            debug_ 5307, LOG_ERROR, "loadWzMapDroidInit", "Failed to build unit %s", droid_.name_
            GoTo continue_
        End If
        If Not (aDroid(psDroid).id <> 0) Then ASSERT_FAILURE "psDroid->id != 0", "src/game.cpp:5310", "loadWzMapDroidInit", "Droid ID should never be zero here"

        ' HACK!!
        startpos = getPlayerStartPosition(player)
        If aDroid(psDroid).droidType = DROID_CONSTRUCT And startpos.x = 0 And startpos.y = 0 Then
            scriptSetStartPos aDroid(psDroid).player, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y  ' set map start position, FIXME - save properly elsewhere!
        End If

        addDroid psDroid, apsDroidLists
continue_:
    Next
    If NumberOfSkippedDroids <> 0 Then
        debug_ 5323, LOG_ERROR, "loadWzMapDroidInit", "Bad Player number in %d unit(s)... assigned to the last player!", NumberOfSkippedDroids
        loadWzMapDroidInit = False
        Exit Function
    End If

    loadWzMapDroidInit = True
End Function

' (VB) a little-endian uint32_t from a byte buffer, as a struct read raw on x86 has it
Private Function bytesLE32(b() As Byte, ByVal i As Long) As Long
    CopyMemory bytesLE32, b(i), 4
End Function

' (VB) PHYSFS_swapUBE32 on a little-endian machine: the bytes reversed
Private Function swapUBE32(ByVal v As Long) As Long
    Dim b(0 To 3) As Byte, r(0 To 3) As Byte
    CopyMemory b(0), v, 4
    r(0) = b(3)
    r(1) = b(2)
    r(2) = b(1)
    r(3) = b(0)
    CopyMemory swapUBE32, r(0), 4
End Function

' (VB) a char[n] read raw: the text up to its terminating zero
Private Function charArrayString(b() As Byte, ByVal first As Long, ByVal n As Long) As String
    Dim i As Long, s As String
    For i = first To first + n - 1
        If b(i) = 0 Then Exit For
        s = s & Chr$(b(i))
    Next
    charArrayString = s
End Function

' (VB) std::list::reverse
Private Sub objlistReverse(l As OBJ_LIST)
    Dim objs() As Long, n As Long, it As Long, i As Long
    If l.size = 0 Then Exit Sub
    ReDim objs(0 To l.size - 1)
    it = l.head
    Do While it <> 0
        objs(n) = lnObj(it)
        n = n + 1
        it = lnNext(it)
    Loop
    objlist_clear l
    For i = n - 1 To 0 Step -1
        objlist_push_back l, objs(i)
    Next
End Sub

' (VB) the keys of a std::map<WzString, ...> (a Dictionary) in its order: WzString's operator< compares the UTF-8 bytes
Private Function sortedKeys(ByVal m As Object, keys() As String) As Long
    Dim k As Variant, n As Long, i As Long, j As Long, t As String
    n = m.Count
    If n = 0 Then Exit Function
    ReDim keys(0 To n - 1)
    i = 0
    For Each k In m.keys
        keys(i) = k
        i = i + 1
    Next
    For i = 1 To n - 1
        t = keys(i)
        j = i - 1
        Do While j >= 0
            If StrComp(keys(j), t, vbBinaryCompare) <= 0 Then Exit Do
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        keys(j + 1) = t
    Next
    sortedKeys = n
End Function

Private Function deserializeSaveGameHeader(ByVal fileHandle As Long, serializeHeader As GAME_SAVEHEADER) As Boolean
    Dim b() As Byte, k As Long

    ' Read in the header from the file
    ReDim b(0 To 3)
    If WZ_PHYSFS_readBytes(fileHandle, b, 0, 4) <> 4 Then
        deserializeSaveGameHeader = False
        Exit Function
    End If
    For k = 0 To 3
        serializeHeader.aFileType(k) = b(k)
    Next
    If WZ_PHYSFS_readBytes(fileHandle, b, 0, 4) <> 4 Then
        deserializeSaveGameHeader = False
        Exit Function
    End If
    serializeHeader.version = bytesLE32(b, 0)

    ' All save game file versions below version 35 (i.e. _not_ version 35 itself)
    ' have their version numbers stored as little endian. Versions from 35 and
    ' onward use big-endian. This basically means that, because of endian
    ' swapping, numbers from 35 and onward will be ridiculously high if a
    ' little-endian byte-order is assumed.

    ' Convert from little endian to native byte-order and check if we get a
    ' ridiculously high number (endian_udword: nothing to do on x86)

    If Not ULT(VERSION_34, serializeHeader.version) Then
        ' Apparently we don't get a ridiculously high number if we assume
        ' little-endian, so lets assume our version number is 34 at max and return
        debug_ 247, LOG_SAVE, "deserializeSaveGameHeader", "Version = %u (little-endian)", serializeHeader.version

        deserializeSaveGameHeader = True
        Exit Function
    Else
        ' Apparently we get a larger number than expected if using little-endian.
        ' So assume we have a version of 35 and onward

        ' Reverse the little-endian decoding (endian_udword: nothing to do on x86)
    End If

    ' Considering that little-endian didn't work we now use big-endian instead
    serializeHeader.version = swapUBE32(serializeHeader.version)
    debug_ 262, LOG_SAVE, "deserializeSaveGameHeader", "Version %u = (big-endian)", serializeHeader.version

    deserializeSaveGameHeader = True
End Function

Private Sub deserializeSaveGameV7Data_json(ByVal o As nlohmann_json, serializeGame As SAVE_GAME_V7)
    serializeGame.gameTime = o.at("gameTime").get_uint()
    serializeGame.GameType = o.at("GameType").get_uint()
    serializeGame.ScrollMinX = o.at("ScrollMinX").get_int()
    serializeGame.ScrollMinY = o.at("ScrollMinY").get_int()
    serializeGame.ScrollMaxX = o.at("ScrollMaxX").get_uint()
    serializeGame.ScrollMaxY = o.at("ScrollMaxY").get_uint()
    serializeGame.levelName = Left$(o.at("levelName").get_string(), MAX_LEVEL_SIZE - 1)  ' sstrcpy
End Sub

' (VB) GameLoadDetails::GameLoadDetails(GameLoadType loadType, const std::string& filePath) (its
' m_logger is made by Class_Initialize)
Private Function GameLoadDetails_new(ByVal loadType_ As Long, ByVal filePath_ As String) As GameLoadDetails
    Dim d As GameLoadDetails
    Set d = New GameLoadDetails
    d.loadType = loadType_
    d.filePath = filePath_
    Set GameLoadDetails_new = d
End Function

Public Function GameLoadDetails_makeUserSaveGameLoad(ByVal saveGame As String) As GameLoadDetails
    Set GameLoadDetails_makeUserSaveGameLoad = GameLoadDetails_new(GameLoadType_UserSaveGame, saveGame)
End Function

Public Function GameLoadDetails_makeMapPackageLoad(ByVal mapPackageFilePath As String) As GameLoadDetails
    Set GameLoadDetails_makeMapPackageLoad = GameLoadDetails_new(GameLoadType_MapPackage, mapPackageFilePath)
End Function

Public Function GameLoadDetails_makeLevelFileLoad(ByVal levelFileName As String) As GameLoadDetails
    Set GameLoadDetails_makeLevelFileLoad = GameLoadDetails_new(GameLoadType_Level, levelFileName)
End Function

'This just loads up the .gam file to determine which level data to set up - split up
'so can be called in levLoadData when starting a game from a load save game

' -----------------------------------------------------------------------------------------
Public Function loadGameInit(ByVal gameToLoad As GameLoadDetails) As Boolean
    Dim pGamInfo As Object, optionsHandler As WZGameReplayOptionsHandler

    If Not (Len(gameToLoad.filePath) <> 0) Then ASSERT_FAILURE "!gameToLoad.filePath.empty()", "src/game.cpp:2314", "loadGameInit", "filePath is empty??": loadGameInit = False: Exit Function

    If gameToLoad.loadType = GameLoadType_MapPackage Then
        Set pGamInfo = gameToLoad.getGamInfoFromPackage()
        If pGamInfo Is Nothing Then
            ' Failed to load map package
            debug_ 2322, LOG_ERROR, "loadGameInit", "Failed to load map package: %s", gameToLoad.filePath
            loadGameInit = False
            Exit Function
        End If

        loadRulesetJson

        saveGameVersion = 0
        savedGameTime = pGamInfo.gameTime
        startX = pGamInfo.ScrollMinX
        startY = pGamInfo.ScrollMinY
        width_ = pGamInfo.ScrollMaxX - pGamInfo.ScrollMinX
        height_ = pGamInfo.ScrollMaxY - pGamInfo.ScrollMinY
        gameType = pGamInfo.GameType
        'set IsScenario to true if not a user saved game
        If gameType = GTYPE_SAVE_START Then
            debug_ 2338, LOG_FATAL, "loadGameInit", "Should not be called with gameType GTYPE_SAVE_START"
        End If
        IsScenario = True
        loadGameInit = True
        Exit Function
    End If

    ' Otherwise, for level load or savegame load cases:

    If strEndsWith(gameToLoad.filePath, ".wzrp") Then
        If Not (gameToLoad.loadType = GameLoadType_UserSaveGame) Then ASSERT_FAILURE "(gameToLoad.loadType == GameLoadDetails::GameLoadType::UserSaveGame)", "src/game.cpp:2348", "loadGameInit", "Invalid load type": loadGameInit = False: Exit Function

        SetGameMode GS_TITLE_SCREEN ' hack - the caller sets this to GS_NORMAL but we actually want to proceed with normal startGameLoop

        ' if it ends in .wzrp, try to load the replay!
        Set optionsHandler = New WZGameReplayOptionsHandler
        If Not NETloadReplay(gameToLoad.filePath, optionsHandler) Then
            loadGameInit = False
            Exit Function
        End If

        bMultiPlayer = True
        bMultiMessages = True
        changeTitleMode STARTGAME
    ElseIf Not gameLoad(gameToLoad.filePath) Then
        debug_ 2365, LOG_ERROR, "loadGameInit", "Corrupted / unsupported savegame file %s, Unable to load!", gameToLoad.filePath
        ' NOTE: why do we start the game clock on a *failed* load?
        ' Start the game clock
        gameTimeStart

        loadGameInit = False
        Exit Function
    End If

    loadGameInit = True
End Function

' -----------------------------------------------------------------------------------------
' Load a file from a save game into the psx.
' This is divided up into 2 parts ...
'
' if it is a level loaded up from CD then UserSaveGame will by false
' UserSaveGame ... Extra stuff to load after scripts
Public Function loadMissionExtras(ByVal pGameToLoad As String, ByVal levelType As Long) As Boolean
    Dim aFileName As String, fileExten As Long

    aFileName = Left$(pGameToLoad, 256 - 1)         ' sstrcpy(aFileName, pGameToLoad) (char[256])
    fileExten = Len(pGameToLoad) - 3
    aFileName = Left$(aFileName, fileExten - 1)     ' aFileName[fileExten - 1] = '\0'
    aFileName = aFileName & "/"

    If saveGameVersion >= VERSION_11 Then
        'if user save game then load up the messages AFTER any droids or structures are loaded
        If gameType = GTYPE_SAVE_START Or gameType = GTYPE_SAVE_MIDMISSION Then
            'load in the message list file
            aFileName = Left$(aFileName, fileExten) & "messtate.json"
            If Not loadSaveMessage(aFileName, levelType) Then
                debug_ 2403, LOG_ERROR, "loadMissionExtras", "Failed to load mission extras from %s", aFileName
                loadMissionExtras = False
                Exit Function
            End If
        End If
    End If

    loadMissionExtras = True
End Function

Private Sub allocatePlayers()
    Dim i As Long
    ' DebugInputManager& dbgInputManager = gInputManager.debugManager(); (DebugInputManager_...)
    For i = 0 To MAX_PLAYERS - 1
        NetPlay_players_team(i) = saveGameData.sNetPlay.players(i).team
        NetPlay_players_ai(i) = saveGameData.sNetPlay.players(i).ai
        NetPlay_players_difficulty(i) = saveGameData.sNetPlay.players(i).difficulty
'		NetPlay.players[i].faction; // read and initialized by loadMainFile
        setPlayerName i, saveGameData.sNetPlay.players(i).name_
        NetPlay_players_position(i) = saveGameData.sNetPlay.players(i).position
        If NetPlay_players_difficulty(i) = AIDifficulty_HUMAN Or (game.typ = LEVEL_TYPE_CAMPAIGN And i = 0) Then
            NetPlay_players_allocated(i) = True
            'processDebugMappings ensures game does not start in DEBUG mode
            DebugInputManager_setPlayerWantsDebugMappings i, False
        Else
            NetPlay_players_allocated(i) = False
        End If
    Next
End Sub

Private Sub getPlayerNames()
    Dim i As Long
    ' Get human and AI players names
    If saveGameVersion >= VERSION_34 Then
        For i = 0 To MAX_PLAYERS - 1
            setPlayerName i, saveGameData.sPlayerName(i)
        Next
    End If
End Sub

' (VB) Public: GameLoadDetails.getMap calls it
Public Function getWzMapType(ByVal UserSaveGame As Boolean) As MapType
    If UserSaveGame Then
        getWzMapType = MapType_SAVEGAME
    Else
        If game.typ = LEVEL_TYPE_CAMPAIGN Then getWzMapType = MapType_CAMPAIGN Else getWzMapType = MapType_SKIRMISH
    End If
End Function

' (VB) loadGame's droidMap: loadSaveDroidPointers on the lists a number stands for
Private Sub loadSaveDroidPointersOf(ByVal key As String, ByVal lists As Long)
    Select Case lists
    Case LISTS_MAIN: loadSaveDroidPointers key, apsDroidLists
    Case LISTS_MISSION: loadSaveDroidPointers key, mission_apsDroidLists
    Case LISTS_LIMBO: loadSaveDroidPointers key, apsLimboDroids
    End Select
End Sub

' (VB) loadGame's structMap: loadSaveStructurePointers on the lists a number stands for
Private Sub loadSaveStructurePointersOf(ByVal key As String, ByVal lists As Long)
    Select Case lists
    Case LISTS_MAIN: loadSaveStructurePointers key, apsStructLists
    Case LISTS_MISSION: loadSaveStructurePointers key, mission_apsStructLists
    End Select
End Sub

' -----------------------------------------------------------------------------------------
' UserSaveGame ... this is true when you are loading a players save game
Public Function loadGame(ByVal gameToLoad As GameLoadDetails, ByVal keepObjects_ As Boolean, ByVal freeMem_ As Boolean) As Boolean
    Dim data As Map, droidMap As Object, structMap As Object, fixedMapIdToGeneratedId As Object
    Dim moduleToBuilding(0 To MAX_PLAYER_SLOTS - 1) As Object
    Dim UserSaveGame As Boolean, aFileName As String, fileExten As Long, fileSize As Long
    Dim player As Long, inc As Long, i As Long, j As Long
    Dim missionScrollMinX As Long, missionScrollMinY As Long, missionScrollMaxX As Long, missionScrollMaxY As Long
    Dim pl As Long, it As Long, psStr As Long, psDroid As Long, psCurr As Long, newMax As Long, seed As Long
    Dim plStats As PLAYERSTATS_, scav As Long, mapData_ As MapData, keys() As String, nKeys As Long, k As Long

    ' only populated when loading maps, *not* savegames
    Set fixedMapIdToGeneratedId = CreateObject("Scripting.Dictionary")
    ' JS compatibility: remembers which module_id maps to which building
    ' - why "MAX_PLAYER_SLOTS" instead of "MAX_PLAYERS"?
    '    because in Campaign, we have objects with player==12
    '   I don't what's the reason but we definitely need to handle more than MAX_PLAYERS
    For i = 0 To MAX_PLAYER_SLOTS - 1
        Set moduleToBuilding(i) = CreateObject("Scripting.Dictionary")
    Next
    Set droidMap = CreateObject("Scripting.Dictionary")
    Set structMap = CreateObject("Scripting.Dictionary")

    UserSaveGame = (gameToLoad.loadType = GameLoadType_UserSaveGame)
    missionScrollMinX = 0
    missionScrollMinY = 0
    missionScrollMaxX = 0
    missionScrollMaxY = 0

    ' Stop the game clock
    gameTimeStop

    If (gameType = GTYPE_SAVE_START) Or _
       (gameType = GTYPE_SAVE_MIDMISSION) Then
        gameTimeReset savedGameTime     'added 14 may 98 JPS to solve kev's problem with no firing droids
    End If

    ' Clear all the objects off the map and free up the map memory
    proj_FreeAllProjectiles     'always clear this
    If freeMem_ Then
        'clear out the audio
        audio_StopAll

        freeAllDroids
        freeAllStructs
        freeAllFeatures

        'clear all the messages?
        releaseAllProxDisp
    End If

    If Not keepObjects_ Then
        'initialise the lists
        For player = 0 To MAX_PLAYERS - 1
            objlist_clear apsDroidLists(player)
            objlist_clear apsStructLists(player)
            objlist_clear apsFlagPosLists(player)
            'clear all the messages?
            objlist_clear apsProxDisp(player)
            objlist_clear apsExtractorLists(player)
        Next
        objlist_clear apsFeatureList(0)
        objlist_clear apsOilList(0)
        objlist_clear apsSensorList(0)
        initFactoryNumFlag
    End If

    If UserSaveGame Then 'always !keepObjects
        'initialise the lists
        For player = 0 To MAX_PLAYERS - 1
            objlist_clear apsLimboDroids(player)
            objlist_clear mission_apsDroidLists(player)
            objlist_clear mission_apsStructLists(player)
            objlist_clear mission_apsFlagPosLists(player)
            objlist_clear mission_apsExtractorLists(player)
        Next
        objlist_clear mission_apsFeatureList(0)
        objlist_clear mission_apsOilList(0)
        objlist_clear mission_apsSensorList(0)

        ' Stuff added after level load to avoid being reset or initialised during load
        ' always !keepObjects

        If saveGameVersion >= VERSION_33 Then
            bMultiPlayer = (saveGameData.multiPlayer <> 0)
        End If

        If saveGameVersion >= VERSION_12 Then
            mission_startTime = saveGameData.missionTime
        End If

        'set the scroll variables
        If saveGameData.ScrollMinX >= 0 And ULT(saveGameData.ScrollMaxX, saveGameData.ScrollMinX) Then
            debug_ 2748, LOG_ERROR, "loadGame", "ScrollMinX (%i) exceeds ScrollMaxX %u", saveGameData.ScrollMinX, saveGameData.ScrollMaxX
            newMax = saveGameData.ScrollMinX
            saveGameData.ScrollMinX = saveGameData.ScrollMaxX
            saveGameData.ScrollMaxX = newMax
        End If
        If saveGameData.ScrollMinY >= 0 And ULT(saveGameData.ScrollMaxY, saveGameData.ScrollMinY) Then
            debug_ 2755, LOG_ERROR, "loadGame", "ScrollMinY (%i) exceeds ScrollMaxY %u", saveGameData.ScrollMinY, saveGameData.ScrollMaxY
            newMax = saveGameData.ScrollMinY
            saveGameData.ScrollMinY = saveGameData.ScrollMaxY
            saveGameData.ScrollMaxY = newMax
        End If
        startX = saveGameData.ScrollMinX
        startY = saveGameData.ScrollMinY
        width_ = saveGameData.ScrollMaxX - saveGameData.ScrollMinX
        height_ = saveGameData.ScrollMaxY - saveGameData.ScrollMinY
        gameType = saveGameData.GameType

        If saveGameVersion >= VERSION_14 Then
            'mission data
            mission_time = saveGameData.missionOffTime
            mission_ETA = saveGameData.missionETA
            mission_homeLZ_X = saveGameData.missionHomeLZ_X
            mission_homeLZ_Y = saveGameData.missionHomeLZ_Y
            mission_playerX = saveGameData.missionPlayerX
            mission_playerY = saveGameData.missionPlayerY
            'mission data
            For player = 0 To MAX_PLAYERS - 1
                aDefaultSensor(player) = saveGameData.aDefaultSensor(player)
                aDefaultECM(player) = saveGameData.aDefaultECM(player)
                aDefaultRepair(player) = saveGameData.aDefaultRepair(player)
                'check for self repair having been set
                If aDefaultRepair(player) <> 0 Then
                    If asRepairStats(aDefaultRepair(player)).location = LOC_DEFAULT Then
                        enableSelfRepair player And &HFF&
                    End If
                End If
                mission_iTranspEntryTileX(player) = saveGameData.iTranspEntryTileX(player)
                mission_iTranspEntryTileY(player) = saveGameData.iTranspEntryTileY(player)
                mission_iTranspExitTileX(player) = saveGameData.iTranspExitTileX(player)
                mission_iTranspExitTileY(player) = saveGameData.iTranspExitTileY(player)
            Next
        End If

        If saveGameVersion >= VERSION_15 Then 'V21
            offWorldKeepLists = (saveGameData.offWorldKeepLists <> 0)
            setRubbleTile saveGameData.RubbleTile
            setUnderwaterTile saveGameData.WaterTile
        End If
        If saveGameVersion >= VERSION_19 Then 'V21
            For i = 0 To MAX_PLAYERS - 1
                For j = 0 To MAX_PLAYERS - 1
                    alliances(i, j) = saveGameData.alliances(i, j)
                    If i = j Then
                        alliances(i, j) = ALLIANCE_FORMED   ' hack to fix old savegames
                    End If
                    If bMultiPlayer And alliancesSharedVision(game.alliance) And alliances(i, j) = ALLIANCE_FORMED Then
                        alliancebits(i) = alliancebits(i) Or BIT_(j)
                    End If
                Next
            Next
            For i = 0 To MAX_PLAYERS - 1
                setPlayerColour i, saveGameData.playerColour(i)
            Next
            SetRadarZoom saveGameData.radarZoom
        End If

        If saveGameVersion >= VERSION_20 Then 'V21
            setDroidsToSafetyFlag saveGameData.bDroidsToSafetyFlag <> 0
        End If

        If saveGameVersion >= VERSION_24 Then 'V24
            missionSetReinforcementTime saveGameData.reinforceTime
            ' horrible hack to catch savegames that were saving garbage into these fields
            If saveGameData.bPlayCountDown <= 1 Then
                setPlayCountDown saveGameData.bPlayCountDown
            End If
            If saveGameData.bPlayerHasWon <= 1 Then
                setPlayerHasWon saveGameData.bPlayerHasWon <> 0
            End If
            If saveGameData.bPlayerHasLost <= 1 Then
                setPlayerHasLost saveGameData.bPlayerHasLost <> 0
            End If
        End If

        'extra code added for the first patch (v1.1) to save out if mission time is not being counted
        If saveGameVersion >= VERSION_31 Then
            'mission data
            mission_cheatTime = saveGameData.missionCheatTime
        End If

        ' skirmish saves.
        If saveGameVersion >= VERSION_33 Then
            scav = game.scavengers

            game = saveGameData.sGame
            game.scavengers = scav  ' ok, so this is butt ugly. but i'm just getting inspiration from the rest of the code around here. ok? - per
            productionPlayer = selectedPlayer
            bMultiMessages = bMultiPlayer

            NetPlay_bComms = saveGameData.sNetPlay.bComms

            allocatePlayers

            If bMultiPlayer Then
                loadMultiStats saveGameData.sPName, plStats             ' stats stuff
                setMultiStats selectedPlayer, plStats, False
                setMultiStats selectedPlayer, plStats, True
            End If
        End If
    End If

    getPlayerNames

    'clear the player Power structs
    If (gameType <> GTYPE_SAVE_START) And (gameType <> GTYPE_SAVE_MIDMISSION) And _
       (Not keepObjects_) Then
        clearPlayerPower
    End If

    'before loading the data - turn power off so don't get any power low warnings
    powerCalculated = False

    ' Load in the chosen file data
    aFileName = Left$(gameToLoad.getMapFolderPath(), 256 - 1)  ' sstrcpy (char[256])
    fileExten = Len(aFileName)

    ' construct the WzMap object for loading map data
    aFileName = Left$(aFileName, fileExten)
    Set data = gameToLoad.getMap(gameRandU32())

    If Not data Is Nothing Then
        If data.wasScriptGenerated() Then
            ' Log the random seed used to generate this instance of the map
            If Not data.scriptGeneratedMapSeed(seed) Then seed = 0   ' value_or(0)
            debug_ 2901, LOG_INFO, "loadGame", "Loaded script-generated map ""%s"" with random seed: %u", aFileName, seed
        End If
    End If

    'the terrain type WILL only change with Campaign changes (well at the moment!)
    If gameType <> GTYPE_SCENARIO_EXPAND Or UserSaveGame Then
        If data Is Nothing Then
            debug_ 2909, LOG_ERROR, "loadGame", "Failed to load map from path: %s", aFileName
            loadGame = False
            Exit Function
        End If
        'load the terrain type data
        If Not loadTerrainTypeMap(data.mapTerrainTypes()) Then
            debug_ 2915, LOG_ERROR, "loadGame", "Failed loading terrain types: %s/ttypes.ttp", aFileName
            GoTo error_
        End If
    End If

    ' decide if we have to create teams, ONLY in multiplayer mode!
    If bMultiPlayer And alliancesSharedVision(game.alliance) Then
        createTeamAlliances

        ' Update ally vision for pre-placed structures and droids
        For pl = 0 To MAX_PLAYERS - 1
            If pl <> selectedPlayer Then
                ' Structures
                it = apsStructLists(pl).head
                Do While it <> 0
                    psStr = lnObj(it)
                    If selectedPlayer < MAX_PLAYERS Then
                        If aiCheckAlliances(aStruct(psStr).player, selectedPlayer) Then
                            visTilesUpdate psStr
                        End If
                    End If
                    it = lnNext(it)
                Loop

                ' Droids
                it = apsDroidLists(pl).head
                Do While it <> 0
                    psDroid = lnObj(it)
                    If selectedPlayer < MAX_PLAYERS Then
                        If aiCheckAlliances(aDroid(psDroid).player, selectedPlayer) Then
                            visTilesUpdate psDroid
                        End If
                    End If
                    it = lnNext(it)
                Loop
            End If
        Next
    End If

    'load up the Droid Templates BEFORE any structures are loaded
    If IsScenario = False Then
        If bMultiPlayer Then
            droidTemplateShutDown
        Else
            clearTemplates 0
            localTemplates_clear
        End If

        aFileName = Left$(aFileName, fileExten) & "templates.json"
        If Not loadSaveTemplate(aFileName) Then
            debug_ 2968, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
            GoTo error_
        End If
    End If

    If (saveGameVersion >= VERSION_15) And UserSaveGame Then
        aFileName = Left$(aFileName, fileExten) & "limits.json"

        'load the data into apsStructLists
        If Not loadSaveLimits(aFileName) Then
            debug_ 2981, LOG_ERROR, "loadGame", "failed to load %s", aFileName
            GoTo error_
        End If
    End If

    'if user save game then load up the research BEFORE any droids or structures are loaded
    If gameType = GTYPE_SAVE_START Or gameType = GTYPE_SAVE_MIDMISSION Then
        'load in the research list file
        aFileName = Left$(aFileName, fileExten) & "resstate.json"
        If Not loadSaveResearch(aFileName) Then
            debug_ 2994, LOG_ERROR, "loadGame", "Failed to load research data from %s", aFileName
            GoTo error_
        End If
    End If

    If saveGameOnMission And UserSaveGame Then
        'the scroll limits for the mission map have already been written
        If saveGameVersion >= VERSION_29 Then
            missionScrollMinX = mission_scrollMinX And &HFFFF&
            missionScrollMinY = mission_scrollMinY And &HFFFF&
            missionScrollMaxX = mission_scrollMaxX And &HFFFF&
            missionScrollMaxY = mission_scrollMaxY And &HFFFF&
        End If

        'load the map and the droids then swap pointers

        'load in the map file
        aFileName = Left$(aFileName, fileExten) & "mission.map"
        If Not mapLoad(aFileName) Then
            debug_ 3017, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
            loadGame = False
            Exit Function
        End If

        'load in the visibility file
        aFileName = Left$(aFileName, fileExten) & "misvis.bjo"

        ' Load in the visibility data from the chosen file
        If Not readVisibilityData(aFileName) Then
            debug_ 3028, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
            GoTo error_
        End If

        ' reload the objects that were in the mission list
        'load in the features -do before the structures
        aFileName = Left$(aFileName, fileExten) & "mfeature.json"

        'load the data into apsFeatureList
        If Not loadSaveFeature2(aFileName) Then
            aFileName = Left$(aFileName, fileExten) & "mfeat.bjo"
            ' Load in the chosen file data
            If Not loadFileToBuffer(aFileName, fileLoadBuffer, FILE_LOAD_BUFFER_SIZE, fileSize) Then
                debug_ 3046, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
                GoTo error_
            End If
            If Not loadSaveFeature(fileLoadBuffer, fileSize) Then
                debug_ 3051, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
                GoTo error_
            End If
        End If

        initStructLimits
        aFileName = Left$(aFileName, fileExten) & "mstruct.json"

        'load in the mission structures
        If Not loadSaveStructure2(aFileName) Then
            aFileName = Left$(aFileName, fileExten) & "mstruct.bjo"
            ' Load in the chosen file data
            If Not loadFileToBuffer(aFileName, fileLoadBuffer, FILE_LOAD_BUFFER_SIZE, fileSize) Then
                debug_ 3069, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
                GoTo error_
            End If
            'load the data into apsStructLists
            If Not loadSaveStructure(fileLoadBuffer, fileSize) Then
                debug_ 3075, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
                GoTo error_
            End If
        Else
            structMap.Item(aFileName) = LISTS_MISSION   ' we swap pointers below
        End If

        ' load in the mission droids, if any
        aFileName = Left$(aFileName, fileExten) & "mdroid.json"
        If loadSaveDroid(aFileName, apsDroidLists) Then
            droidMap.Item(aFileName) = LISTS_MISSION    ' need to swap here to read correct list later
        End If

        ' after we've loaded in the units we need to redo the orientation because
        ' the direction may have been saved - we need to do it outside of the loop
        ' whilst the current map is valid for the units
        For player = 0 To MAX_PLAYERS - 1
            it = apsDroidLists(player).head
            Do While it <> 0
                psCurr = lnObj(it)
                ' (C++: a commented-out "&& psCurr->droidType != DROID_CYBORG" before !psCurr->isCyborg())
                If aDroid(psCurr).droidType <> DROID_PERSON _
                   And Not DROID_isCyborg(psCurr) _
                   And (Not DROID_isTransporter(psCurr)) _
                   And aDroid(psCurr).pos.x <> INVALID_XY Then
                    updateDroidOrientation psCurr
                End If
                it = lnNext(it)
            Loop
        Next

        swapMissionPointers

        'once the mission map has been loaded reset the mission scroll limits
        If saveGameVersion >= VERSION_29 Then
            mission_scrollMinX = missionScrollMinX
            mission_scrollMinY = missionScrollMinY
            mission_scrollMaxX = missionScrollMaxX
            mission_scrollMaxY = missionScrollMaxY
        End If
    End If

    'if Campaign Expand then don't load in another map
    If gameType <> GTYPE_SCENARIO_EXPAND Then
        Erase psMapTiles                ' psMapTiles = nullptr
        psMapTiles_allocated = False
        ' load in the map file
        If data Is Nothing Then
            debug_ 3130, LOG_ERROR, "loadGame", "Failed to load map from path: %s", aFileName
            loadGame = False
            Exit Function
        End If
        Set mapData_ = data.mapData()
        If mapData_ Is Nothing Then
            debug_ 3136, LOG_ERROR, "loadGame", "Failed to load map data from path: %s", aFileName
            loadGame = False
            Exit Function
        End If
        If data.wasScriptGenerated() Then
            syncDebug "loadGame", "mapSize = [%d, %d]", mapData_.width, mapData_.height
            syncDebug "loadGame", "crc(maptiles) = 0x%08x", mapData_.crcSumMapTiles(0)
            syncDebug "loadGame", "crc(structures) = 0x%08x", data.crcSumStructures(0)
            syncDebug "loadGame", "crc(droids) = 0x%08x", data.crcSumDroids(0)
            syncDebug "loadGame", "crc(features) = 0x%08x", data.crcSumFeatures(0)
        End If
        If Not mapLoadFromWzMapData(mapData_) Then
            debug_ 3149, LOG_ERROR, "loadGame", "Failed to process map data from path: %s", aFileName
            loadGame = False
            Exit Function
        End If
    End If

    ' FIXME THIS FILE IS A HUGE MESS, this code should probably appear at another position...
    If saveGameVersion > VERSION_12 Then
        'if user save game then load up the FX
        If (gameType = GTYPE_SAVE_START) Or _
           (gameType = GTYPE_SAVE_MIDMISSION) Then
            'load in the message list file
            aFileName = Left$(aFileName, fileExten) & "fxstate.json"

            ' load the fx data from the file
            If Not readFXData(aFileName) Then
                debug_ 3168, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
                GoTo error_
            End If
        End If
    End If

    'save game stuff added after map load

    If saveGameVersion >= VERSION_16 Then
        For inc = 0 To MAX_NOGO_AREAS - 1
            setNoGoArea saveGameData.sLandingZone(inc).x1, saveGameData.sLandingZone(inc).y1, _
                        saveGameData.sLandingZone(inc).x2, saveGameData.sLandingZone(inc).y2, inc And &HFF&
        Next
    End If

    'adjust the scroll range for the new map or the expanded map
    setMapScroll

    If IsScenario Then
        'load in the droids
        aFileName = Left$(aFileName, fileExten) & "droid.json"

        'load the data into apsDroidLists
        If Not UserSaveGame Then
            If Not (Not data Is Nothing) Then ASSERT_FAILURE "data != nullptr", "src/game.cpp:3197", "loadGame", "Expecting WzMap::Map instance"
            If loadWzMapDroidInit(data, fixedMapIdToGeneratedId) Then
                debug_ 3200, LOG_SAVE, "loadGame", "Loaded new style droids"
                ' *maps* should not be added to droidMap to have loadSaveDroidPointers called later
            Else
                aFileName = Left$(aFileName, fileExten)
                debug_ 3206, LOG_ERROR, "loadGame", "Failed to load map droid init from map directory: %s", aFileName
                GoTo error_
            End If
        Else
            If loadSaveDroid(aFileName, apsDroidLists) Then
                debug_ 3214, LOG_SAVE, "loadGame", "Loaded new style droids"
                droidMap.Item(aFileName) = LISTS_MAIN   ' load pointers later
            End If
        End If
    Else
        'load in the droids
        aFileName = Left$(aFileName, fileExten) & "droid.json"

        'load the data into apsDroidLists
        If Not loadSaveDroid(aFileName, apsDroidLists) Then
            debug_ 3228, LOG_ERROR, "loadGame", "failed to load %s", aFileName
            GoTo error_
        End If
        droidMap.Item(aFileName) = LISTS_MAIN   ' load pointers later

        ' after we've loaded in the units we need to redo the orientation because
        ' the direction may have been saved - we need to do it outside of the loop
        ' whilst the current map is valid for the units
        For player = 0 To MAX_PLAYERS - 1
            it = apsDroidLists(player).head
            Do While it <> 0
                psCurr = lnObj(it)
                If aDroid(psCurr).droidType <> DROID_PERSON _
                   And Not DROID_isCyborg(psCurr) _
                   And Not DROID_isTransporter(psCurr) _
                   And aDroid(psCurr).pos.x <> INVALID_XY Then
                    updateDroidOrientation psCurr
                End If
                it = lnNext(it)
            Loop
        Next
        If Not saveGameOnMission Then
            'load in the mission droids
            aFileName = Left$(aFileName, fileExten) & "mdroid.json"

            ' load the data into mission.apsDroidLists, if any
            If loadSaveDroid(aFileName, mission_apsDroidLists) Then
                droidMap.Item(aFileName) = LISTS_MISSION
            End If
        End If
    End If

    If saveGameVersion >= VERSION_23 Then
        ' load in the limbo droids, if any
        aFileName = Left$(aFileName, fileExten) & "limbo.json"
        If loadSaveDroid(aFileName, apsLimboDroids) Then
            droidMap.Item(aFileName) = LISTS_LIMBO
        End If
    End If

    'load in the features -do before the structures
    aFileName = Left$(aFileName, fileExten)
    If Not UserSaveGame Then
        If Not (Not data Is Nothing) Then ASSERT_FAILURE "data != nullptr", "src/game.cpp:3279", "loadGame", "Expecting WzMap::Map instance"
        If Not loadWzMapFeature(data, fixedMapIdToGeneratedId) Then
            debug_ 3282, LOG_ERROR, "loadGame", "Failed to load map feature init from map directory: %s", aFileName
            GoTo error_
        End If
    Else
        aFileName = Left$(aFileName, fileExten) & "feature.json"
        If Not loadSaveFeature2(aFileName) Then
            debug_ 3292, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
            GoTo error_
        End If
    End If

    'load in the structures
    initStructLimits
    aFileName = Left$(aFileName, fileExten) & "struct.json"
    If Not UserSaveGame Then
        If Not (Not data Is Nothing) Then ASSERT_FAILURE "data != nullptr", "src/game.cpp:3303", "loadGame", "Expecting WzMap::Map instance"
        If game.typ <> LEVEL_TYPE_CAMPAIGN Then
            freeAllFlagPositions        'clear any flags put in during level loads
        End If
        If Not loadWzMapStructure(data, fixedMapIdToGeneratedId, moduleToBuilding) Then
            aFileName = Left$(aFileName, fileExten)
            debug_ 3311, LOG_ERROR, "loadGame", "Failed to load map structure init from map directory: %s", aFileName
            GoTo error_
        End If
        If game.typ <> LEVEL_TYPE_CAMPAIGN Then
            resetFactoryNumFlag         'reset flags into the masks
        End If
    Else
        If Not loadSaveStructure2(aFileName) Then
            debug_ 3323, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
            GoTo error_
        End If
        structMap.Item(aFileName) = LISTS_MAIN
    End If

    'if user save game then load up the current level for structs and components
    If gameType = GTYPE_SAVE_START Or gameType = GTYPE_SAVE_MIDMISSION Then
        'load in the component list file
        aFileName = Left$(aFileName, fileExten) & "complist.json"
        If Not loadSaveCompList(aFileName) Then
            debug_ 3337, LOG_ERROR, "loadGame", "failed to load %s", aFileName
            GoTo error_
        End If
        'load in the structure type list file
        aFileName = Left$(aFileName, fileExten) & "strtype.json"
        If Not loadSaveStructTypeList(aFileName) Then
            debug_ 3345, LOG_ERROR, "loadGame", "failed to load %s", aFileName
            GoTo error_
        End If

        ' load in the game guide topics
        aFileName = Left$(aFileName, fileExten) & "guidetopics.json"
        If Not loadSaveGuideTopics(aFileName) Then
            debug_ 3354, LOG_ERROR, "loadGame", "failed to load %s", aFileName
            ' currently not fatal if this fails
        End If
    End If

    If saveGameVersion >= VERSION_11 Then
        'if user save game then load up the Visibility
        If (gameType = GTYPE_SAVE_START) Or _
           (gameType = GTYPE_SAVE_MIDMISSION) Then
            'load in the visibility file
            aFileName = Left$(aFileName, fileExten) & "visstate.bjo"

            ' Load in the visibility data from the chosen file
            If Not readVisibilityData(aFileName) Then
                debug_ 3372, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
                GoTo error_
            End If
        End If
    End If

    If saveGameVersion >= VERSION_16 Then
        'if user save game then load up the FX
        If (gameType = GTYPE_SAVE_START) Or _
           (gameType = GTYPE_SAVE_MIDMISSION) Then
            aFileName = Left$(aFileName, fileExten) & "score.json"

            ' Load the fx data from the chosen file
            If Not readScoreData(aFileName) Then
                debug_ 3390, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
                GoTo error_
            End If
        End If
    End If

    If saveGameVersion >= VERSION_21 Then
        'rebuild the apsCommandDesignation AFTER all droids and structures are loaded
        If (gameType = GTYPE_SAVE_START) Or _
           (gameType = GTYPE_SAVE_MIDMISSION) Then
            'load in the command list file
            aFileName = Left$(aFileName, fileExten) & "firesupport.json"

            If Not readFiresupportDesignators(aFileName) Then
                debug_ 3408, LOG_ERROR, "loadGame", "Failed with: %s", aFileName
                GoTo error_
            End If
        End If
    End If

    If (saveGameVersion >= VERSION_15) And UserSaveGame Then
        setCurrentStructQuantity False
    Else
        setCurrentStructQuantity True
    End If

    'check that delivery points haven't been put down in invalid location
    checkDeliveryPoints saveGameVersion

    'turn power on for rest of game
    powerCalculated = True

    If Not keepObjects_ Then 'only reset the pointers if they were set
        ' Reset the object pointers in the droid target lists
        ' Should only be called for savegames
        nKeys = sortedKeys(droidMap, keys)
        For k = 0 To nKeys - 1
            loadSaveDroidPointersOf keys(k), droidMap.Item(keys(k))
        Next
        nKeys = sortedKeys(structMap, keys)
        For k = 0 To nKeys - 1
            loadSaveStructurePointersOf keys(k), structMap.Item(keys(k))
        Next
    End If

    If UserSaveGame Then
        loadFixupResearchPendingStates
    End If

    ' Load labels
    aFileName = Left$(aFileName, fileExten) & "labels.json"
    loadLabels aFileName, fixedMapIdToGeneratedId, moduleToBuilding, UserSaveGame

    'if user save game then reset the time - BEWARE IF YOU USE IT
    If (gameType = GTYPE_SAVE_START) Or _
       (gameType = GTYPE_SAVE_MIDMISSION) Then
        If Not (gameTime = savedGameTime) Then ASSERT_FAILURE "gameTime == savedGameTime", "src/game.cpp:3461", "loadGame", "loadGame; game time modified during load"
        gameTimeReset savedGameTime     'added 14 may 98 JPS to solve kev's problem with no firing droids

        'reset the objId for new objects
        If saveGameVersion >= VERSION_17 Then
            unsynchObjID = FromU32(Int(U32(FromU32(U32(savedObjId) + 1)) / 2))  ' Make new object ID start at savedObjId*8.
            synchObjID = MUL32(savedObjId, 4)       ' Make new object ID start at savedObjId*8.
        End If

        If getDroidsToSafetyFlag() Then
            'The droids lists are "reversed" as they are loaded in loadSaveDroid().
            'Which later causes issues in saveCampaignData() which tries to extract
            'the first transporter group sent off at Beta-end by reversing this very list.
            If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/game.cpp:3476", "loadGame", "selectedPlayer is out of bounds: %u", selectedPlayer
            objlistReverse mission_apsDroidLists(selectedPlayer)
        End If
    End If

    'check the research button isn't flashing unnecessarily
    'cancel first
    stopReticuleButtonFlash IDRET_RESEARCH
    'then see if needs to be set
    intNotifyResearchButton 0

    'set up the mission countdown flag
    setMissionCountDown

    ' Start the game clock
    gameTimeStart

    'check if limbo_expand mission has changed to an expand mission for user save game (mid-mission)
    If gameType = GTYPE_SAVE_MIDMISSION And missionLimboExpand() Then
        ' when all the units have moved from the mission.apsDroidList then the
        ' campaign has been reset to an EXPAND type - OK so there should have
        ' been another flag to indicate this state has changed but its late in
        ' the day excuses...excuses...excuses
        If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/game.cpp:3501", "loadGame", "selectedPlayer is out of bounds: %u", selectedPlayer
        If mission_apsDroidLists(selectedPlayer).size = 0 Then
            'set the mission type
            startMissionSave LEVEL_TYPE_LDS_EXPAND
        End If
    End If

    'set this if come into a save game mid mission
    If gameType = GTYPE_SAVE_MIDMISSION Then
        setScriptWinLoseVideo PLAY_NONE
    End If

    'need to clear before setting up
    clearMissionWidgets
    'put any widgets back on for the missions
    resetMissionWidgets

    debug_ 3520, LOG_NEVER, "loadGame", "Done loading"

    loadGame = True
    Exit Function

error_:
    debug_ 3525, LOG_ERROR, "loadGame", "Game load failed for %s, FS:%s, params=%s,%s,%s", gameToLoad.filePath, PHYSFS_getRealDir(gameToLoad.filePath) & "", _
          IIf(keepObjects_, "true", "false"), IIf(freeMem_, "true", "false"), IIf(UserSaveGame, "true", "false")

    ' Clear all the objects off the map and free up the map memory
    freeAllDroids
    freeAllStructs
    freeAllFeatures
    droidTemplateShutDown
    Erase psMapTiles                    ' psMapTiles = nullptr
    psMapTiles_allocated = False

    ' Start the game clock
    gameTimeStart

    loadGame = False
End Function

Private Function loadRulesetJson() As Boolean
    Dim ruleset As WzConfig, tag As String, k As Long, c As String, valid As Boolean

    ' Prior to getting here, the directory structure has been set to wherever the
    ' map or savegame is loaded from, so we will get the right ruleset file.
    If PHYSFS_exists("ruleset.json") = 0 Then
        debug_ 3879, LOG_ERROR, "loadRulesetJson", "ruleset.json not found! User generated data will not work."
        rulesettag = ""                 ' memset(rulesettag, 0, sizeof(rulesettag))
        loadRulesetJson = False
        Exit Function
    Else
        Set ruleset = WzConfig_new("ruleset.json", WzConfig_ReadOnly)
        If Not ruleset.contains("tag") Then
            debug_ 3888, LOG_ERROR, "loadRulesetJson", "ruleset tag not found in ruleset.json!"  ' fall-through
        End If
        tag = ruleset.value("tag", json_variant_new("[]")).toWzString()
        rulesettag = Left$(tag, 40 - 1) ' sstrcpy(rulesettag, tag.toUtf8().c_str()) (char rulesettag[40])
        ' strspn(rulesettag, "abcdefghijklmnopqrstuvwxyz") != strlen(rulesettag)
        valid = True
        For k = 1 To Len(rulesettag)
            c = Mid$(rulesettag, k, 1)
            If c < "a" Or c > "z" Then valid = False
        Next
        If Not valid Then ' for safety
            debug_ 3894, LOG_ERROR, "loadRulesetJson", "ruleset.json userdata tag contains invalid characters!"
            debug_ 3895, LOG_ERROR, "loadRulesetJson", "User generated data will not work."
            rulesettag = ""             ' memset(rulesettag, 0, sizeof(rulesettag))
        End If
    End If

    loadRulesetJson = True
End Function

' -----------------------------------------------------------------------------------------
Private Function gameLoad(ByVal fileName As String) As Boolean
    Dim CurrentFileName As String, fileHeader As GAME_SAVEHEADER, gamJsonSave As nlohmann_json, hasGamJsonSave As Boolean
    Dim fileHandle As Long, retVal As Boolean

    CurrentFileName = fileName          ' strcpy(CurrentFileName, fileName)
    ' GAME_SAVEHEADER fileHeader = {};
    hasGamJsonSave = readGamJson(fileName, gamJsonSave)
    debug_ 3910, LOG_INFO, "gameLoad", "loading: %s", fileName
    fileHandle = openLoadFile(fileName, False)
    If Not hasGamJsonSave And fileHandle <> 0 Then
        ' haven't converted .gam to .json yet!
        ' Read the header from the file
        If Not deserializeSaveGameHeader(fileHandle, fileHeader) Then
            debug_ 3918, LOG_ERROR, "gameLoad", "gameLoad: error while reading header from file (%s): %s", fileName, WZ_PHYSFS_getLastError()
            PHYSFS_close fileHandle
            gameLoad = False
            Exit Function
        End If
    ElseIf hasGamJsonSave Then
        ' be compatible with .gam logic
        fileHeader.version = gamJsonSave.at("version").get_uint()
        fileHeader.aFileType(0) = Asc("g")
        fileHeader.aFileType(1) = Asc("a")
        fileHeader.aFileType(2) = Asc("m")
        fileHeader.aFileType(3) = Asc("e")
    End If
    ' exit if no .gam neither .json was found
    If fileHandle = 0 And Not hasGamJsonSave Then
        ' Failure to open the file is a failure to load the specified savegame
        gameLoad = False
        Exit Function
    End If
    debug_ 3938, LOG_WZ, "gameLoad", "gameLoad"

    ' Check the header to see if we've been given a file of the right type
    If Not (fileHeader.aFileType(0) = Asc("g") And fileHeader.aFileType(1) = Asc("a") And fileHeader.aFileType(2) = Asc("m") And fileHeader.aFileType(3) = Asc("e")) Then
        debug_ 3943, LOG_ERROR, "gameLoad", "gameLoad: Weird file type found? Has header letters - '%c' '%c' '%c' '%c' (should be 'g' 'a' 'm' 'e')", _
              CLng(fileHeader.aFileType(0)), _
              CLng(fileHeader.aFileType(1)), _
              CLng(fileHeader.aFileType(2)), _
              CLng(fileHeader.aFileType(3))

        PHYSFS_close fileHandle

        gameLoad = False
        Exit Function
    End If

    debug_ 3954, LOG_NEVER, "gameLoad", "gl .gam file is version %u\n", fileHeader.version

    ' Prior to getting here, the directory structure has been set to wherever the
    ' map or savegame is loaded from, so we will get the right ruleset file.
    loadRulesetJson

    'set main version Id from game file
    saveGameVersion = fileHeader.version
    debug_ 3962, LOG_SAVE, "gameLoad", "file version is %u, (%s)", fileHeader.version, fileName
    ' Check the file version
    If ULT(fileHeader.version, VERSION_7) Then
        debug_ 3966, LOG_ERROR, "gameLoad", "gameLoad: unsupported save format version %d", fileHeader.version
        PHYSFS_close fileHandle

        gameLoad = False
        Exit Function
    ElseIf ULT(fileHeader.version, VERSION_9) Then
        retVal = gameLoadV7(fileHandle, hasGamJsonSave, gamJsonSave)
        PHYSFS_close fileHandle
        gameLoad = retVal
        Exit Function
    ElseIf Not ULT(CURRENT_VERSION_NUM, fileHeader.version) Then
        'The in-game load menu was clearing level data AFTER loading in some savegame data.
        'Hacking this in here so things make a little more sense and so that data loaded
        'in from main.json is somewhat safe from being reset by mistake.
        If Not levReleaseAll() Then
            debug_ 3984, LOG_ERROR, "gameLoad", "Failed to unload old data. Attempting to load anyway"
        End If
        challengeFileName = ""

        'remove the file extension
        CurrentFileName = Left$(CurrentFileName, Len(CurrentFileName) - 4)
        loadMainFile CurrentFileName & "/main.json"

        retVal = gameLoadV(fileHandle, fileHeader.version, hasGamJsonSave, gamJsonSave)
        PHYSFS_close fileHandle

        loadMainFileFinal CurrentFileName & "/main.json"

        gameLoad = retVal
        Exit Function
    Else
        debug_ 4001, LOG_ERROR, "gameLoad", "Unsupported main save format version %u", fileHeader.version
        PHYSFS_close fileHandle

        gameLoad = False
    End If
End Function

' -----------------------------------------------------------------------------------------
' code specific to version 7 of a save game
Private Function gameLoadV7(ByVal fileHandle As Long, ByVal hasGamJson As Boolean, ByVal gamJson As nlohmann_json) As Boolean
    Dim saveGame As SAVE_GAME_V7, b() As Byte, newMax As Long

    If hasGamJson Then
        ' this seems to be still used by maps/mission loading routines
        deserializeSaveGameV7Data_json gamJson, saveGame
    Else
        ReDim b(0 To SAVE_GAME_V7_SIZE - 1)
        If WZ_PHYSFS_readBytes(fileHandle, b, 0, SAVE_GAME_V7_SIZE) <> SAVE_GAME_V7_SIZE Then
            debug_ 4143, LOG_ERROR, "gameLoadV7", "gameLoadV7: error while reading file: %s", WZ_PHYSFS_getLastError()

            gameLoadV7 = False
            Exit Function
        End If
        ' (VB) the struct as read raw
        saveGame.gameTime = bytesLE32(b, 0)
        saveGame.GameType = bytesLE32(b, 4)
        saveGame.ScrollMinX = bytesLE32(b, 8)
        saveGame.ScrollMinY = bytesLE32(b, 12)
        saveGame.ScrollMaxX = bytesLE32(b, 16)
        saveGame.ScrollMaxY = bytesLE32(b, 20)
        saveGame.levelName = charArrayString(b, 24, MAX_LEVEL_SIZE)

        ' GAME_SAVE_V7 (endian_udword, endian_sdword: nothing to do on x86)
    End If

    savedGameTime = saveGame.gameTime

    'set the scroll variables
    If saveGame.ScrollMinX >= 0 And ULT(saveGame.ScrollMaxX, saveGame.ScrollMinX) Then
        debug_ 4162, LOG_ERROR, "gameLoadV7", "ScrollMinX (%i) exceeds ScrollMaxX %u", saveGame.ScrollMinX, saveGame.ScrollMaxX
        newMax = saveGame.ScrollMinX
        saveGame.ScrollMinX = saveGame.ScrollMaxX
        saveGame.ScrollMaxX = newMax
    End If
    If saveGame.ScrollMinY >= 0 And ULT(saveGame.ScrollMaxY, saveGame.ScrollMinY) Then
        debug_ 4169, LOG_ERROR, "gameLoadV7", "ScrollMinY (%i) exceeds ScrollMaxY %u", saveGame.ScrollMinY, saveGame.ScrollMaxY
        newMax = saveGame.ScrollMinY
        saveGame.ScrollMinY = saveGame.ScrollMaxY
        saveGame.ScrollMaxY = newMax
    End If
    startX = saveGame.ScrollMinX
    startY = saveGame.ScrollMinY
    width_ = saveGame.ScrollMaxX - saveGame.ScrollMinX
    height_ = saveGame.ScrollMaxY - saveGame.ScrollMinY
    gameType = saveGame.GameType
    'set IsScenario to true if not a user saved game
    If gameType = GTYPE_SAVE_START Then
        debug_ 4182, LOG_FATAL, "gameLoadV7", "Should not be called with gameType GTYPE_SAVE_START"
    End If
    IsScenario = True

    gameLoadV7 = True
End Function

' -----------------------------------------------------------------------------------------
' non specific version of a save game
' (VB) the versions 9 to 34, which the C++ reads as raw structs (sizeof(SAVE_GAME_V10) ... sizeof(SAVE_GAME_V34)),
' are refused: Warzone 2.x saves (and 4.7.0 reads its V33 and V34 over a NETPLAY and a MULTIPLAYERGAME that
' hold std containers now); endian_SaveGameV, which only they need, is not ported
Public Function gameLoadV(ByVal fileHandle As Long, ByVal version As Long, ByVal hasGamJson As Boolean, ByVal gamJson As nlohmann_json) As Boolean
    Dim i As Long, j As Long, player As Long, newMax As Long, playerStats As PLAYERSTATS_, scav As Long
    Static powerSaved(0 To MAX_PLAYERS - 1) As SAVE_POWER

    debug_ 4197, LOG_WZ, "gameLoadV", "gameLoadV: version %u", version

    ' Version 7 and earlier are loaded separately in gameLoadV7

    'size is now variable so only check old save games
    If Not ULT(VERSION_34, version) Then
        debug_ 0, LOG_ERROR, "gameLoadV", "(VB) gameLoadV: binary save games of version %u (before version 35) are not supported", version

        gameLoadV = False
        Exit Function
    ElseIf ULT(version, VERSION_39) Then
        debug_ 4366, LOG_ERROR, "gameLoadV", "Unsupported savegame version"
        gameLoadV = False
        Exit Function
    ElseIf Not ULT(CURRENT_VERSION_NUM, version) Then
        If hasGamJson Then
            debug_ 4373, LOG_SAVEGAME, "gameLoadV", "gam json found, loading that."
            If Not deserializeSaveGameData_json(gamJson, saveGameData) Then
                debug_ 4376, LOG_ERROR, "gameLoadV", "failed to load gamjson"
                gameLoadV = False
                Exit Function
            End If
        Else
            debug_ 4382, LOG_SAVEGAME, "gameLoadV", "no gam json found, falling back to .gam"
            If Not deserializeSaveGameData(fileHandle, saveGameData) Then
                debug_ 4385, LOG_ERROR, "gameLoadV", "gameLoadV: error while reading data from file for deserialization (with version number %u): %s", version, WZ_PHYSFS_getLastError()
                gameLoadV = False
                Exit Function
            End If
        End If
    Else
        debug_ 4392, LOG_ERROR, "gameLoadV", "Unsupported version number (%u) for savegame", version

        gameLoadV = False
        Exit Function
    End If

    debug_ 4397, LOG_SAVE, "gameLoadV", "Savegame is of type: %u", saveGameData.sGame.typ And &HFF&
    game.typ = saveGameData.sGame.typ
    ' Test mod list
    If Not ULT(version, VERSION_38) Then
        setOverrideMods saveGameData.modList
    End If

    If game.typ = LEVEL_TYPE_CAMPAIGN And hasOverrideMods() Then
        ' When loading campaign saves, if there are mods loaded (such as campaign mods),
        ' we *must* reload the search paths now (to ensure that levels in campaign mods can be found)
        debug_ 4409, LOG_SAVE, "gameLoadV", "Rebuilding level list with override mods loaded"

        ' Clear & reload level list, rebuilding search paths to include the override mods
        levShutDown
        levInitialise
        rebuildSearchPath mod_campaign, True
        pal_Init ' Update palettes.
        If Not buildMapList(True) Then
            debug_ 4418, LOG_FATAL, "gameLoadV", "Failed to build map / level list?"
        End If
    End If

    ' All savegames from version 34 or before are little endian so swap them. All
    ' from version 35, and onward, are already swapped to the native byte-order
    ' by the (de)serialization API (VB: those versions are refused above)

    savedGameTime = saveGameData.gameTime

    If Not ULT(version, VERSION_12) Then
        mission_startTime = saveGameData.missionTime
        If (saveGameData.saveKey And SAVEKEY_ONMISSION) <> 0 Then
            saveGameOnMission = True
        Else
            saveGameOnMission = False
        End If

    Else
        saveGameOnMission = False
    End If

    'set the scroll variables
    If saveGameData.ScrollMinX >= 0 And ULT(saveGameData.ScrollMaxX, saveGameData.ScrollMinX) Then
        debug_ 4453, LOG_ERROR, "gameLoadV", "ScrollMinX (%i) exceeds ScrollMaxX %u", saveGameData.ScrollMinX, saveGameData.ScrollMaxX
        newMax = saveGameData.ScrollMinX
        saveGameData.ScrollMinX = saveGameData.ScrollMaxX
        saveGameData.ScrollMaxX = newMax
    End If
    If saveGameData.ScrollMinY >= 0 And ULT(saveGameData.ScrollMaxY, saveGameData.ScrollMinY) Then
        debug_ 4460, LOG_ERROR, "gameLoadV", "ScrollMinY (%i) exceeds ScrollMaxY %u", saveGameData.ScrollMinY, saveGameData.ScrollMaxY
        newMax = saveGameData.ScrollMinY
        saveGameData.ScrollMinY = saveGameData.ScrollMaxY
        saveGameData.ScrollMaxY = newMax
    End If
    startX = saveGameData.ScrollMinX
    startY = saveGameData.ScrollMinY
    width_ = saveGameData.ScrollMaxX - saveGameData.ScrollMinX
    height_ = saveGameData.ScrollMaxY - saveGameData.ScrollMinY
    gameType = saveGameData.GameType

    If Not ULT(version, VERSION_11) Then
        'camera position
        disp3d_setView saveGameData.currentPlayerPos
    Else
        disp3d_oldView
    End If

    'load mission data from save game these values reloaded after load game

    If Not ULT(version, VERSION_14) Then
        'mission data
        mission_time = saveGameData.missionOffTime
        mission_ETA = saveGameData.missionETA
        mission_homeLZ_X = saveGameData.missionHomeLZ_X
        mission_homeLZ_Y = saveGameData.missionHomeLZ_Y
        mission_playerX = saveGameData.missionPlayerX
        mission_playerY = saveGameData.missionPlayerY

        For player = 0 To MAX_PLAYERS - 1
            mission_iTranspEntryTileX(player) = saveGameData.iTranspEntryTileX(player)
            mission_iTranspEntryTileY(player) = saveGameData.iTranspEntryTileY(player)
            mission_iTranspExitTileX(player) = saveGameData.iTranspExitTileX(player)
            mission_iTranspExitTileY(player) = saveGameData.iTranspExitTileY(player)
            aDefaultSensor(player) = saveGameData.aDefaultSensor(player)
            aDefaultECM(player) = saveGameData.aDefaultECM(player)
            aDefaultRepair(player) = saveGameData.aDefaultRepair(player)
        Next
    End If

    If Not ULT(version, VERSION_15) Then
        offWorldKeepLists = saveGameData.offWorldKeepLists <> 0
        setRubbleTile saveGameData.RubbleTile
        setUnderwaterTile saveGameData.WaterTile
    End If

    If Not ULT(version, VERSION_17) Then
        unsynchObjID = FromU32(Int(U32(FromU32(U32(saveGameData.objId) + 1)) / 2))  ' Make new object ID start at savedObjId*8.
        synchObjID = MUL32(saveGameData.objId, 4)       ' Make new object ID start at savedObjId*8.
        savedObjId = saveGameData.objId
    End If

    If Not ULT(version, VERSION_19) Then 'version 19
        For i = 0 To MAX_PLAYERS - 1
            For j = 0 To MAX_PLAYERS - 1
                alliances(i, j) = saveGameData.alliances(i, j)
            Next
        Next
        For i = 0 To MAX_PLAYERS - 1
            setPlayerColour i, saveGameData.playerColour(i)
        Next
    End If

    If Not ULT(version, VERSION_20) Then 'version 20
        setDroidsToSafetyFlag saveGameData.bDroidsToSafetyFlag <> 0
    End If

    If Not ULT(saveGameVersion, VERSION_24) Then 'V24
        missionSetReinforcementTime saveGameData.reinforceTime

        ' horrible hack to catch savegames that were saving garbage into these fields
        If saveGameData.bPlayCountDown <= 1 Then
            setPlayCountDown saveGameData.bPlayCountDown
        End If
        If saveGameData.bPlayerHasWon <= 1 Then
            setPlayerHasWon saveGameData.bPlayerHasWon <> 0
        End If
        If saveGameData.bPlayerHasLost <= 1 Then
            setPlayerHasLost saveGameData.bPlayerHasLost <> 0
        End If
    End If

    If Not ULT(saveGameVersion, VERSION_29) Then
        mission_scrollMinX = saveGameData.missionScrollMinX
        mission_scrollMinY = saveGameData.missionScrollMinY
        mission_scrollMaxX = saveGameData.missionScrollMaxX
        mission_scrollMaxY = saveGameData.missionScrollMaxY
    End If

    If Not ULT(saveGameVersion, VERSION_31) Then
        mission_cheatTime = saveGameData.missionCheatTime
    End If

    droidInit

    'set IsScenario to true if not a user saved game
    If (gameType = GTYPE_SAVE_START) Or _
       (gameType = GTYPE_SAVE_MIDMISSION) Then
        For i = 0 To MAX_PLAYERS - 1
            powerSaved(i).currentPower = saveGameData.power(i).currentPower
            powerSaved(i).extractedPower = saveGameData.power(i).extractedPower
        Next

        allocatePlayers

        IsScenario = False
        'copy the level name across
        aLevelName = Left$(saveGameData.levelName, 255)   ' sstrcpy(aLevelName, saveGameData.levelName)
        'load up the level dataset
        ' Not sure what aLevelName is, in relation to game.map. But need to use aLevelName here, to be able to start the right map for campaign, and need game.hash, to start the right non-campaign map, if there are multiple identically named maps.
        If Not levLoadData(aLevelName, saveGameData.sGame.hash, True, saveGameName, gameType) Then
            gameLoadV = False
            Exit Function
        End If

        If Not ULT(saveGameVersion, VERSION_33) Then
            scav = game.scavengers  ' loaded earlier, keep it over struct copy below

            bMultiPlayer = saveGameData.multiPlayer <> 0
            bMultiMessages = bMultiPlayer
            productionPlayer = selectedPlayer
            game = saveGameData.sGame   ' why do this again????
            game.scavengers = scav
            NetPlay_bComms = saveGameData.sNetPlay.bComms
            If bMultiPlayer Then
                loadMultiStats saveGameData.sPName, playerStats             ' stats stuff
                setMultiStats selectedPlayer, playerStats, False
                setMultiStats selectedPlayer, playerStats, True
            End If
        End If
    Else
        IsScenario = True
    End If

    getPlayerNames

    clearPlayerPower
    'don't adjust any power if a camStart (gameType is set to GTYPE_SCENARIO_START when a camChange saveGame is loaded)
    If gameType <> GTYPE_SCENARIO_START Then
        'set the players power
        For i = 0 To MAX_PLAYERS - 1
            'only overwrite selectedPlayer's power on a startMission save game
            If gameType = GTYPE_SAVE_MIDMISSION Or i = selectedPlayer Then
                setPower i, powerSaved(i).currentPower
            End If
        Next
    End If
    radarPermitted = powerSaved(0).extractedPower <> 0  ' nice hack, eh? don't want to break savegames now...
    allowDesign = powerSaved(1).extractedPower <> 0     ' nice hack, eh? don't want to break savegames now...

    gameLoadV = True
End Function

' (VB) game_main_c's saveGame and writers: the module's static UDWORD saveGameVersion and bool saveGameOnMission
Public Function saveGameVersionNow() As Long
    saveGameVersionNow = saveGameVersion
End Function

Public Function saveGameOnMissionNow() As Boolean
    saveGameOnMissionNow = saveGameOnMission
End Function

Public Sub setSaveGameOnMission(ByVal value As Boolean)
    saveGameOnMission = value
End Sub

' optional<nlohmann::json>: has_value, the value in result
Private Function readGamJson(ByVal filenameWithGamExtension As String, result As nlohmann_json) As Boolean
    Dim len_ As Long, lastSep As Long, lastSegment As String, commonPath As String, containingFolder As String
    Dim gamJson As String, oldGamJsonPath As String, hasResult As Boolean

    ' TODO: use SaveGamePath_t instead of char* ...
    ' but that's a lot of change, so not doing it now
    len_ = Len(filenameWithGamExtension)
    If Not (Right$(filenameWithGamExtension, 4) = ".gam") Then ASSERT_FAILURE "strcmp(filenameWithGamExtension + (len - 4), "".gam"") == 0", "src/game.cpp:5197", "readGamJson", "hmm... not .gam?"
    lastSep = InStrRev(filenameWithGamExtension, "/") - 1   ' rfind("/"): -1 for std::string::npos
    If Not (lastSep > 0 Or lastSep = -1) Then ASSERT_FAILURE "lastSep > 0", "src/game.cpp:5200", "readGamJson", "unexpected filename format: '%s'", filenameWithGamExtension  ' (npos > 0)
    lastSegment = Mid$(filenameWithGamExtension, lastSep + 2, len_ - lastSep - 5)
    commonPath = Left$(filenameWithGamExtension, len_ - Len(lastSegment) - 4)
    containingFolder = commonPath & lastSegment & "/"
    ' "gam.json" in the containingFolder
    gamJson = containingFolder & "gam.json"
    hasResult = parseJsonFile(gamJson, result)
    If hasResult Then
        debug_ 5209, LOG_SAVEGAME, "readGamJson", "last segment was %s, common path %s, save-info %s", lastSegment, commonPath, gamJson
    Else
        ' FALLBACK for older saves: prior behavior was to save "<lastSegment>.json" in the containingFolder
        oldGamJsonPath = containingFolder & lastSegment & ".json"
        debug_ 5215, LOG_SAVEGAME, "readGamJson", "last segment was %s, common path %s, save-info %s", lastSegment, commonPath, oldGamJsonPath
        hasResult = parseJsonFile(oldGamJsonPath, result)
    End If
    readGamJson = hasResult
End Function

' optional<nlohmann::json>: has_value, the value in result
Public Function parseJsonFile(ByVal filename As String, result As nlohmann_json) As Boolean
    Dim pFileSize As Long, ppFileData() As Byte

    debug_ 5225, LOG_SAVEGAME, "parseJsonFile", "starting deserialize %s", filename
    If Not loadFile(filename, ppFileData, pFileSize, False) Then
        debug_ 5228, LOG_SAVE, "parseJsonFile", "No %s found, sad", filename
        Set result = Nothing
        parseJsonFile = False
        Exit Function
    End If
    On Error GoTo failed
    Set result = json_parse_bytes(ppFileData, 0, pFileSize)
    On Error GoTo 0
    parseJsonFile = True
    Exit Function

failed:
    ' catch (const std::exception &e)
    ASSERT_FAILURE "false", "src/game.cpp:5236", "parseJsonFile", "JSON document from %s is invalid: %s", filename, Err.Description
    Set result = Nothing
    parseJsonFile = False
End Function

' -----------------------------------------------------------------------------------------
' set the global scroll values to use for the save game
Private Sub setMapScroll()
    'if loading in a pre version5 then scroll values will not have been set up so set to max poss
    If width_ = 0 And height_ = 0 Then
        scrollMinX = 0
        scrollMaxX = mapWidth
        scrollMinY = 0
        scrollMaxY = mapHeight
        Exit Sub
    End If
    scrollMinX = startX
    scrollMinY = startY
    scrollMaxX = startX + width_
    scrollMaxY = startY + height_
    'check not going beyond width/height of map
    If scrollMaxX > mapWidth Then
        scrollMaxX = mapWidth
        debug_ 8228, LOG_NEVER, "setMapScroll", "scrollMaxX was too big - It has been set to map width"
    End If
    If scrollMaxY > mapHeight Then
        scrollMaxY = mapHeight
        debug_ 8233, LOG_NEVER, "setMapScroll", "scrollMaxY was too big - It has been set to map height"
    End If
    ' check for invalid minimum values (fixes some broken maps)
    If scrollMinX >= scrollMaxX Then
        ASSERT_FAILURE "false", "src/game.cpp:8238", "setMapScroll", "scrollMinX was >= scrollMaxX - min has been set to 0, max has been set to mapWidth"
        scrollMinX = 0
        scrollMaxX = mapWidth
    End If
    If scrollMinY >= scrollMaxY Then
        ASSERT_FAILURE "false", "src/game.cpp:8244", "setMapScroll", "scrollMinY was >= scrollMaxY - min has been set to 0, max has been set to mapHeight"
        scrollMinY = 0
        scrollMaxY = mapHeight
    End If
End Sub

' -----------------------------------------------------------------------------------------
'returns the current type of save game being loaded
Public Function getSaveGameType() As Long
    getSaveGameType = gameType
End Function
