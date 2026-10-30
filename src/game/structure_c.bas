Attribute VB_Name = "structure_c"
' src/structure.cpp, structure.h: so far the structure stats: their store
' (asStructureStats, numStructureStats), the loaders of structure.json and
' the strength modifiers, the lookups by name, and the favourites. A
' STRUCTURE_STATS * is its index in asStructureStats (-1 = nullptr);
' asStructureStats_allocated stands for asStructureStats != nullptr. The
' name lookup (std::unordered_map) is a Scripting.Dictionary of indices.
' optional<int> structureDamageBaseExperienceLevel is the value plus
' has_structureDamageBaseExperienceLevel.
' The structures themselves: the container (GlobalStructContainer),
' STRUCTURE's constructor and destructor, checkStructure, the aux bits,
' structure limits and factory flags, walls, foundations,
' buildStructureDir (with modules), setFunctionality, validLocation,
' removeStruct, assembly points, the derrick and power generator links,
' buildingComplete, the module stats, body and resistance, hold and release
' of production, requestOpenGate; for the droids' actions and orders:
' structureBuild, structureDemolish, structureRepair, gateCurrentOpenHeight,
' structGetDemolishStat, calcStructureMuzzleBaseLocation,
' validStructResistance, the sensor tests, the rearm pads
' (findNearestReArmPad, ensureRearmPadClear, clearRearmPad,
' vtolOnRearmPad), isDamaged, structPowerToBuildOrAddNextModule,
' canStructureHaveAModuleAdded and calcLineBuild (LineBuild's operator[]
' is LineBuild_at; the (size, type, ...) overload is calcLineBuild_size;
' begin and step are begin_ and step_). A local STRUCTURE (STRUCTURE building(id, player)) is a
' STRUCTURE UDT, copied into the pool by GlobalStructContainer_emplace.
' std::vector<bool> factoryNumFlag[MAX_PLAYERS][NUM_FLAG_TYPES] is a fixed
' Boolean array with factoryNumFlag_size. The damage (structureDamage,
' destroyStruct, electronicDamage, the muzzles). The structures' update
' (structureUpdate, aiUpdateStructure, the repair facility's
' aiUpdateRepair_*, the rearm pads), factory production
' (structSetManufacture; the production runs: std::vector<ProductionRun>
' asProductionRun[NUM_FACTORY_TYPES] is a ProductionRunList per factory
' type, see structuredef_h; cancelProduction, doNextProduction,
' factoryProdUpdate, factoryProdAdjust, getProduction), droid placement
' (placeDroid, structPlaceDroid: only std::sort's least element is used,
' found by a linear search), factory commanders, the unit limits. A
' FACTORY *, REPAIR_FACILITY *, REARM_PAD * or RESEARCH_FACILITY * is its
' structure's handle; aiUpdateStructure's BASE_STATS *pSubject is the
' RESEARCH * (-1 = nullptr) or DROID_TEMPLATE * (0 = nullptr) plus
' hasSubject. The gifts and captures: transferFixupFunctionality, the
' rewards (electronicReward, factoryReward, repairFacilityReward, hqReward)
' and giftSingleStructure.
' Not ported yet: the blueprint wall code (structChooseWallTypeBlueprint,
' structFindWallBlueprints, buildBlueprint), fillStructureList,
' printStructureInfo, loadFavoriteStructsFile and
' writeFavoriteStructsFile (the JSON file of favourites: favoriteStructs
' stays empty).
Option Explicit

' Extra z padding for assembly points
Public Const ASSEMBLY_POINT_Z_PADDING As Long = 10

Public Type LineBuild
    begin_ As Vector2i
    step_ As Vector2i
    count As Long
End Type

Public Const STRUCTURE_DAMAGE_SCALING As Long = 400

'production loop max
Public Const INFINITE_PRODUCTION As Long = 9 '10

'This should correspond to the structLimits!
Public Const MAX_FACTORY As Long = 5

'used to flag when the Factory is ready to start building
Public Const ACTION_START_TIME As Long = 0

'Value is stored for easy access to this structure stat
Public factoryModuleStat As Long
Public powerModuleStat As Long
Public researchModuleStat As Long

'holder for all StructureStats
Public asStructureStats() As STRUCTURE_STATS
Public asStructureStats_allocated As Boolean    ' (VB) asStructureStats != nullptr
Public numStructureStats As Long
Private lookupStructStatPtr As Object           ' std::unordered_map<WzString, STRUCTURE_STATS *>
Private structureDamageBaseExperienceLevel As Long
Private has_structureDamageBaseExperienceLevel As Boolean

'used to hold the modifiers cross refd by weapon effect and structureStrength
Public asStructStrengthModifier(0 To WE_NUMEFFECTS - 1, 0 To NUM_STRUCT_STRENGTH - 1) As Long  ' STRUCTSTRENGTH_MODIFIER

' destroy building construction droid stat pointer
Private g_psStatDestroyStruct As Long           ' STRUCTURE_STATS *

Private favoriteStructs As StringVector

Private Const MAX_INCLINE As Long = 50 '80'40

' droid construction smoke cloud constants
Private Const DROID_CONSTRUCTION_SMOKE_OFFSET As Long = 30
Private Const DROID_CONSTRUCTION_SMOKE_HEIGHT As Long = 20

'specifies which numbers have been allocated for the assembly points for the factories
Private factoryNumFlag(0 To MAX_PLAYERS - 1, 0 To NUM_FLAG_TYPES - 1, 0 To 255) As Boolean
Private factoryNumFlag_size(0 To MAX_PLAYERS - 1, 0 To NUM_FLAG_TYPES - 1) As Long

'stores which player the production list has been set up for
Public productionPlayer As Long         ' SBYTE

' the structure that was last hit
Public psLastStructHit As Long

'flag for drawing all sat uplink sees
Private satUplinkExists(0 To MAX_PLAYERS - 1) As Long  ' UBYTE
'flag for when the player has one built - either completely or partially
Private lasSatExists(0 To MAX_PLAYERS - 1) As Long     ' UBYTE

' last time the maximum units message was displayed
Private lastMaxUnitMessage As Long

' max number of units
Private droidLimit(0 To MAX_PLAYERS - 1) As Long
' max number of commanders
Private commanderLimit(0 To MAX_PLAYERS - 1) As Long
' max number of constructors
Private constructorLimit(0 To MAX_PLAYERS - 1) As Long

'used to calculate how often to increase the resistance level of a structure
Private Const RESISTANCE_INTERVAL As Long = 2000

' the number of different (types of) droids that can be put into a production run
Private Const MAX_IN_RUN As Long = 9

'the list of what to build - only for selectedPlayer
Public asProductionRun(0 To NUM_FACTORY_TYPES - 1) As ProductionRunList

Private Const MAX_UNIT_MESSAGE_PAUSE As Long = 40000

Private Const VERSION_19 As Long = 19   ' game.h

' All this wall type code is horrible, but I really can't think of a better way to do it.  John.
' (enum WallOrientation)
Private Const WallConnectNone As Long = 0
Private Const WallConnectLeft As Long = 1
Private Const WallConnectRight As Long = 2
Private Const WallConnectUp As Long = 4
Private Const WallConnectDown As Long = 8

' enum STRUCTURE_PACKABILITY
Private Const PACKABILITY_EMPTY As Long = 0
Private Const PACKABILITY_DEFENSE As Long = 1
Private Const PACKABILITY_NORMAL As Long = 2
Private Const PACKABILITY_REPAIR As Long = 3

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub structure_c_init()
    g_psStatDestroyStruct = -1
End Sub

Public Function getStructureDamageBaseExperienceLevel() As Long
    ' COMPAT NOTES:
    '
    ' Default / compat structure damage handling (the only option for many years - from at least 2.0.10-4.4.2):
    '
    ' This causes the game to treat structures at a base experience level of 1 instead of 0 when calculating damage to them,
    ' yielding actualDamage at 94% of the base damage value. Or, in other words, structures are a bit tougher
    ' than the raw numbers in the stats files would suggest, and get a hidden experience level boost.
    '
    ' However, structure.json created and tested during this long period may be expecting this outcome / behavior,
    ' So unless it's explicitly specified in the special `_config_` dict, it always defaults to `1`.

    If has_structureDamageBaseExperienceLevel Then
        getStructureDamageBaseExperienceLevel = structureDamageBaseExperienceLevel
    Else
        getStructureDamageBaseExperienceLevel = 1
    End If
End Function

' (VB) map_STRUCTURE_TYPE and map_STRUCT_STRENGTH as dictionaries (std::map<WzString, ...>)
Private Function map_STRUCTURE_TYPE() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbBinaryCompare
    m.Add "HQ", REF_HQ
    m.Add "FACTORY", REF_FACTORY
    m.Add "FACTORY MODULE", REF_FACTORY_MODULE
    m.Add "RESEARCH", REF_RESEARCH
    m.Add "RESEARCH MODULE", REF_RESEARCH_MODULE
    m.Add "POWER GENERATOR", REF_POWER_GEN
    m.Add "POWER MODULE", REF_POWER_MODULE
    m.Add "RESOURCE EXTRACTOR", REF_RESOURCE_EXTRACTOR
    m.Add "DEFENSE", REF_DEFENSE
    m.Add "WALL", REF_WALL
    m.Add "CORNER WALL", REF_WALLCORNER
    m.Add "REPAIR FACILITY", REF_REPAIR_FACILITY
    m.Add "COMMAND RELAY", REF_COMMAND_CONTROL
    m.Add "DEMOLISH", REF_DEMOLISH
    m.Add "CYBORG FACTORY", REF_CYBORG_FACTORY
    m.Add "VTOL FACTORY", REF_VTOL_FACTORY
    m.Add "LAB", REF_LAB
    m.Add "GENERIC", REF_GENERIC
    m.Add "REARM PAD", REF_REARM_PAD
    m.Add "MISSILE SILO", REF_MISSILE_SILO
    m.Add "SAT UPLINK", REF_SAT_UPLINK
    m.Add "GATE", REF_GATE
    m.Add "LASSAT", REF_LASSAT
    m.Add "FORTRESS", REF_FORTRESS
    Set map_STRUCTURE_TYPE = m
End Function

Private Function map_STRUCT_STRENGTH() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbBinaryCompare
    m.Add "SOFT", STRENGTH_SOFT
    m.Add "MEDIUM", STRENGTH_MEDIUM
    m.Add "HARD", STRENGTH_HARD
    m.Add "BUNKER", STRENGTH_BUNKER
    Set map_STRUCT_STRENGTH = m
End Function

Private Sub initModuleStats(ByVal i As Long, ByVal typ As STRUCTURE_TYPE)
    'need to work out the stats for the modules - HACK! - But less hacky than what was here before.
    Select Case typ
    Case REF_FACTORY_MODULE
        'store the stat for easy access later on
        factoryModuleStat = i

    Case REF_RESEARCH_MODULE
        'store the stat for easy access later on
        researchModuleStat = i

    Case REF_POWER_MODULE
        'store the stat for easy access later on
        powerModuleStat = i
    Case Else
    End Select
End Sub

' load the structure stats from the ini file
Public Function loadStructureStats(ByVal ini As WzConfig) As Boolean
    Const CONFIG_DICT_KEY_STR As String = "_config_"     ' STRUCTURE_JSON_CONFIG_DICT_KEY
    Dim structType As Object, structStrength As Object
    Dim list() As String, n As Long, statWriteIdx As Long, readIdx As Long, i As Long, j As Long, k As Long
    Dim convValueSuccess As Variant, baseStructDamageExpLevel As Long, psStats As Long, typ As String
    Dim limits0 As Long, limits1 As Long, limits2 As Long, flags() As String, nFlags As Long, strength As String
    Dim models() As String, nModels As Long, imd As Long, baseModel As String, ecm As Long, sensor As Long
    Dim weapons() As String, nWeapons As Long, weaponsID As String, weapon_ As Long, types As Long, iID As Long

    Set structType = map_STRUCTURE_TYPE()
    Set structStrength = map_STRUCT_STRENGTH()
    If lookupStructStatPtr Is Nothing Then
        Set lookupStructStatPtr = CreateObject("Scripting.Dictionary")
        lookupStructStatPtr.CompareMode = vbBinaryCompare
    End If

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/structure.cpp:505", "loadStructureStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    Erase asStructureStats
    If n > 0 Then ReDim asStructureStats(0 To n - 1)
    For k = 0 To n - 1
        STRUCTURE_STATS_construct asStructureStats(k)
    Next
    asStructureStats_allocated = True
    numStructureStats = 0
    statWriteIdx = 0
    For readIdx = 0 To n - 1
        If list(readIdx) = CONFIG_DICT_KEY_STR Then
            ' handle the special config dict
            ini.beginGroup list(readIdx)

            ' baseStructDamageExpLevel
            convValueSuccess = False
            baseStructDamageExpLevel = ini.value("baseStructDamageExpLevel", json_variant_new(1)).toInt(convValueSuccess)
            If Not CBool(convValueSuccess) Then
                baseStructDamageExpLevel = 1 ' reset to old default
            End If
            If baseStructDamageExpLevel >= 0 And baseStructDamageExpLevel < 10 Then
                If Not has_structureDamageBaseExperienceLevel Then
                    structureDamageBaseExperienceLevel = baseStructDamageExpLevel
                    has_structureDamageBaseExperienceLevel = True
                Else
                    If structureDamageBaseExperienceLevel <> baseStructDamageExpLevel Then
                        debug_ 534, LOG_ERROR, "loadStructureStats", "Non-matching structure JSON baseStructDamageExpLevel"
                        debug_ 535, LOG_INFO, "loadStructureStats", "Structure JSON file ""%s"" has specified a baseStructDamageExpLevel (""%d"") that does not match the first loaded structure JSON's baseStructDamageExpLevel (""%d"")", ini.fileName(), baseStructDamageExpLevel, structureDamageBaseExperienceLevel
                    End If
                End If
            Else
                ASSERT_FAILURE "false", "src/structure.cpp:541", "loadStructureStats", "Invalid _config_ ""baseStructDamageExpLevel"" value: ""%d""", baseStructDamageExpLevel
                loadStructureStats = False
                Exit Function
            End If

            ini.endGroup
            GoTo continue_
        End If

        ini.beginGroup list(readIdx)
        psStats = statWriteIdx
        loadStructureStats_BaseStats ini, psStats, statWriteIdx
        asStructureStats(psStats).ref = STAT_STRUCTURE + statWriteIdx

        ' set structure type
        typ = ini.value("type", json_variant_new("")).toWzString()
        If Not structType.Exists(typ) Then
            ASSERT_FAILURE "false", "src/structure.cpp:557", "loadStructureStats", "Invalid type '%s' of structure '%s'", typ, asStructureStats(psStats).id
            ' nothing we can really do here except skip this stat
            unloadStructureStats_BaseStats asStructureStats(psStats) ' must be called to properly clear from lookup tables (that loadStructureStats_BaseStats adds it to)
            ini.endGroup
            GoTo continue_
        End If
        asStructureStats(psStats).typ = structType(typ)

        ' save indexes of special structures for futher use
        initModuleStats statWriteIdx, asStructureStats(psStats).typ  ' This function looks like a hack. But slightly less hacky than before.

        If ini.contains("userLimits") Then
            ini.vector3i "userLimits", limits0, limits1, limits2
            asStructureStats(psStats).minLimit = limits0
            asStructureStats(psStats).maxLimit = limits2
            asStructureStats(psStats).base.limit = limits1
        Else
            asStructureStats(psStats).minLimit = 0
            asStructureStats(psStats).maxLimit = LOTS_OF
            asStructureStats(psStats).base.limit = LOTS_OF
        End If
        asStructureStats(psStats).base.research = ini.value("researchPoints", json_variant_new(0)).toInt()
        asStructureStats(psStats).base.moduleResearch = ini.value("moduleResearchPoints", json_variant_new(0)).toInt()
        asStructureStats(psStats).base.production = ini.value("productionPoints", json_variant_new(0)).toInt()
        asStructureStats(psStats).base.moduleProduction = ini.value("moduleProductionPoints", json_variant_new(0)).toInt()
        asStructureStats(psStats).base.repair = ini.value("repairPoints", json_variant_new(0)).toInt()
        asStructureStats(psStats).base.power = ini.value("powerPoints", json_variant_new(0)).toInt()
        asStructureStats(psStats).base.modulePower = ini.value("modulePowerPoints", json_variant_new(0)).toInt()
        asStructureStats(psStats).base.rearm = ini.value("rearmPoints", json_variant_new(0)).toInt()
        asStructureStats(psStats).base.resistance = ini.value("resistance", json_variant_new(0)).toUInt()
        asStructureStats(psStats).base.hitpoints = ini.value("hitpoints", json_variant_new(1)).toUInt()
        asStructureStats(psStats).base.armour = ini.value("armour", json_variant_new(0)).toUInt()
        asStructureStats(psStats).base.thermal = ini.value("thermal", json_variant_new(0)).toUInt()
        For i = 0 To MAX_PLAYERS - 1
            asStructureStats(psStats).upgrade(i).limit = asStructureStats(psStats).base.limit
            asStructureStats(psStats).upgrade(i).research = asStructureStats(psStats).base.research
            asStructureStats(psStats).upgrade(i).moduleResearch = asStructureStats(psStats).base.moduleResearch
            asStructureStats(psStats).upgrade(i).power = asStructureStats(psStats).base.power
            asStructureStats(psStats).upgrade(i).modulePower = asStructureStats(psStats).base.modulePower
            asStructureStats(psStats).upgrade(i).repair = asStructureStats(psStats).base.repair
            asStructureStats(psStats).upgrade(i).production = asStructureStats(psStats).base.production
            asStructureStats(psStats).upgrade(i).moduleProduction = asStructureStats(psStats).base.moduleProduction
            asStructureStats(psStats).upgrade(i).rearm = asStructureStats(psStats).base.rearm
            asStructureStats(psStats).upgrade(i).resistance = ini.value("resistance", json_variant_new(0)).toUInt()
            asStructureStats(psStats).upgrade(i).hitpoints = ini.value("hitpoints", json_variant_new(1)).toUInt()
            asStructureStats(psStats).upgrade(i).armour = ini.value("armour", json_variant_new(0)).toUInt()
            asStructureStats(psStats).upgrade(i).thermal = ini.value("thermal", json_variant_new(0)).toUInt()
        Next

        asStructureStats(psStats).flags = 0
        nFlags = ini.value("flags").toWzStringList(flags)
        For i = 0 To nFlags - 1
            If flags(i) = "Connected" Then
                asStructureStats(psStats).flags = asStructureStats(psStats).flags Or STRUCTURE_CONNECTED
            End If
        Next

        ' set structure strength
        strength = ini.value("strength", json_variant_new("")).toWzString()
        If Not structStrength.Exists(strength) Then
            ASSERT_FAILURE "false", "src/structure.cpp:624", "loadStructureStats", "Invalid strength '%s' of structure '%s'", strength, asStructureStats(psStats).id
            strength = "SOFT" ' just default to the first strength - stat file creator needs to fix the json!
        End If
        asStructureStats(psStats).strength = structStrength(strength)

        ' set baseWidth
        asStructureStats(psStats).baseWidth = ini.value("width", json_variant_new(0)).toUInt()
        If ULT(100, asStructureStats(psStats).baseWidth) Then
            ASSERT_FAILURE "false", "src/structure.cpp:633", "loadStructureStats", "Invalid width '%d' for structure '%s'", asStructureStats(psStats).baseWidth, asStructureStats(psStats).id
            asStructureStats(psStats).baseWidth = 100
        End If

        ' set baseBreadth
        asStructureStats(psStats).baseBreadth = ini.value("breadth", json_variant_new(0)).toUInt()
        If Not ULT(asStructureStats(psStats).baseBreadth, 100) Then
            ASSERT_FAILURE "false", "src/structure.cpp:641", "loadStructureStats", "Invalid breadth '%d' for structure '%s'", asStructureStats(psStats).baseBreadth, asStructureStats(psStats).id
            asStructureStats(psStats).baseBreadth = 99
        End If

        asStructureStats(psStats).height = ini.value("height").toUInt()
        asStructureStats(psStats).powerToBuild = ini.value("buildPower").toUInt()
        asStructureStats(psStats).buildPoints = ini.value("buildPoints").toUInt()

        ' set structure models
        nModels = ini.value("structureModel").toWzStringList(models)
        For j = 0 To nModels - 1
            imd = modelGet(WzString_trimmed(models(j)))
            If Not (imd <> 0) Then ASSERT_FAILURE "imd != nullptr", "src/structure.cpp:654", "loadStructureStats", "Cannot find the PIE structureModel '%s' for structure '%s'", models(j), asStructureStats(psStats).id
            ReDim Preserve asStructureStats(psStats).pIMD(0 To asStructureStats(psStats).pIMD_size)
            asStructureStats(psStats).pIMD(asStructureStats(psStats).pIMD_size) = imd
            asStructureStats(psStats).pIMD_size = asStructureStats(psStats).pIMD_size + 1
        Next

        ' set base model
        baseModel = ini.value("baseModel", json_variant_new("")).toWzString()
        If baseModel <> "" Then
            imd = modelGet(baseModel)
            If Not (imd <> 0) Then ASSERT_FAILURE "imd != nullptr", "src/structure.cpp:663", "loadStructureStats", "Cannot find the PIE baseModel '%s' for structure '%s'", baseModel, asStructureStats(psStats).id
            asStructureStats(psStats).pBaseIMD = imd
        End If

        ecm = getCompFromName(COMP_ECM, ini.value("ecmID", json_variant_new("ZNULLECM")).toWzString())
        If ecm >= 0 Then
            asStructureStats(psStats).pECM = ecm
        Else
            If Not (ecm >= 0) Then ASSERT_FAILURE "ecm >= 0", "src/structure.cpp:674", "loadStructureStats", "Invalid ECM found for '%s'", asStructureStats(psStats).id
        End If

        sensor = getCompFromName(COMP_SENSOR, ini.value("sensorID", json_variant_new("ZNULLSENSOR")).toWzString())
        If sensor >= 0 Then
            asStructureStats(psStats).pSensor = sensor
        Else
            If Not (sensor >= 0) Then ASSERT_FAILURE "sensor >= 0", "src/structure.cpp:684", "loadStructureStats", "Invalid sensor found for structure '%s'", asStructureStats(psStats).id
        End If

        ' set list of weapons
        For j = 0 To MAX_WEAPONS - 1
            asStructureStats(psStats).psWeapStat(j) = -1
        Next
        nWeapons = ini.value("weapons").toWzStringList(weapons)
        If nWeapons > MAX_WEAPONS Then
            ASSERT_FAILURE "false", "src/structure.cpp:692", "loadStructureStats", "Too many weapons are attached to structure '%s'. Maximum is %d", asStructureStats(psStats).id, MAX_WEAPONS
            nWeapons = MAX_WEAPONS
        End If
        asStructureStats(psStats).numWeaps = nWeapons
        For j = 0 To asStructureStats(psStats).numWeaps - 1
            weaponsID = WzString_trimmed(weapons(j))
            weapon_ = getCompFromName(COMP_WEAPON, weaponsID)
            If weapon_ >= 0 Then
                asStructureStats(psStats).psWeapStat(j) = weapon_
            Else
                ' couldn't find the weapon - truncate weapons at this item
                ASSERT_FAILURE "false", "src/structure.cpp:708", "loadStructureStats", "Invalid item '%s' in list of weapons of structure '%s' ", weaponsID, asStructureStats(psStats).id
                asStructureStats(psStats).numWeaps = j
                Exit For
            End If
        Next

        ' check used structure turrets
        types = 0
        If asStructureStats(psStats).numWeaps <> 0 Then types = types + 1
        If asStructureStats(psStats).pECM <> -1 Then
            If asECMStats(asStructureStats(psStats).pECM).location = LOC_TURRET Then types = types + 1
        End If
        If asStructureStats(psStats).pSensor <> -1 Then
            If asSensorStats(asStructureStats(psStats).pSensor).location = LOC_TURRET Then types = types + 1
        End If
        If Not (types <= 1) Then ASSERT_FAILURE "types <= 1", "src/structure.cpp:719", "loadStructureStats", "Too many turret types for structure '%s'", asStructureStats(psStats).id

        asStructureStats(psStats).combinesWithWall = ini.value("combinesWithWall", json_variant_new(False)).toBool()

        ini.endGroup

        If Not lookupStructStatPtr.Exists(asStructureStats(psStats).id) Then lookupStructStatPtr.Add asStructureStats(psStats).id, psStats
        statWriteIdx = statWriteIdx + 1
continue_:
    Next
    numStructureStats = statWriteIdx
    parseFavoriteStructs

    ' get global dummy stat pointer - GJ
    g_psStatDestroyStruct = -1
    For iID = 0 To numStructureStats - 1
        If asStructureStats(iID).typ = REF_DEMOLISH Then
            g_psStatDestroyStruct = iID
            Exit For
        End If
    Next
    If Not (g_psStatDestroyStruct <> -1) Then ASSERT_FAILURE "g_psStatDestroyStruct", "src/structure.cpp:741", "loadStructureStats", "Destroy structure stat not found": loadStructureStats = False: Exit Function

    loadStructureStats = True
End Function

'Load the Structure Strength Modifiers from the file exported from Access
Public Function loadStructureStrengthModifiers(ByVal ini As WzConfig) As Boolean
    Dim i As Long, j As Long, list() As String, n As Long, keys() As String, nKeys As Long
    Dim effectInc As WEAPON_EFFECT, strength As String, modifier As Long

    'initialise to 100%
    For i = 0 To WE_NUMEFFECTS - 1
        For j = 0 To NUM_STRUCT_STRENGTH - 1
            asStructStrengthModifier(i, j) = 100
        Next
    Next
    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/structure.cpp:780", "loadStructureStrengthModifiers", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        If Not getWeaponEffect(list(i), effectInc) Then
            debug_ 788, LOG_FATAL, "loadStructureStrengthModifiers", "Invalid Weapon Effect - %s", list(i)
            ini.endGroup
            GoTo continue_
        End If
        nKeys = ini.childKeys(keys)
        For j = 0 To nKeys - 1
            strength = keys(j)
            modifier = ini.value(strength).toInt()
            ' FIXME - add support for dynamic categories
            If strength = "SOFT" Then
                asStructStrengthModifier(effectInc, 0) = modifier And &HFFFF&
            ElseIf strength = "MEDIUM" Then
                asStructStrengthModifier(effectInc, 1) = modifier And &HFFFF&
            ElseIf strength = "HARD" Then
                asStructStrengthModifier(effectInc, 2) = modifier And &HFFFF&
            ElseIf strength = "BUNKER" Then
                asStructStrengthModifier(effectInc, 3) = modifier And &HFFFF&
            Else
                debug_ 816, LOG_ERROR, "loadStructureStrengthModifiers", "Unsupported structure strength %s", strength
            End If
        Next
        ini.endGroup
continue_:
    Next
    loadStructureStrengthModifiers = True
End Function

Public Function structureStatsShutDown() As Boolean
    Dim i As Long
    packFavoriteStructs
    If asStructureStats_allocated Then
        For i = 0 To numStructureStats - 1
            unloadStructureStats_BaseStats asStructureStats(i)
        Next
    End If
    If Not lookupStructStatPtr Is Nothing Then lookupStructStatPtr.RemoveAll
    Erase asStructureStats
    asStructureStats_allocated = False
    numStructureStats = 0
    structureStatsShutDown = True
End Function

' gets a structure stat from its name - relies on the name being unique (or it will
'return the first one it finds!!
Public Function getStructStatFromName(ByVal name_ As String) As Long
    Dim psStat As Long
    psStat = getStructStatsFromName(name_)
    If psStat <> -1 Then
        getStructStatFromName = asStructureStats(psStat).index
        Exit Function
    End If
    getStructStatFromName = -1
End Function

' the STRUCTURE_STATS's index, or -1
Public Function getStructStatsFromName(ByVal name_ As String) As Long
    Dim psStat As Long
    psStat = -1
    If Not lookupStructStatPtr Is Nothing Then
        If lookupStructStatPtr.Exists(name_) Then
            psStat = lookupStructStatPtr(name_)
        End If
    End If
    getStructStatsFromName = psStat
End Function

Public Function STRUCTURE_isIdle(ByVal this As Long) As Boolean
    Dim pSubject As Long, hasSubject As Boolean

    checkStructure this, "src/structure.cpp:4869", "isIdle", max_check_object_recursion

    If aStruct(this).pFunctionality = 0 Then
        STRUCTURE_isIdle = True
        Exit Function
    End If

    'determine the Subject
    Select Case asStructureStats(aStruct(this).pStructureType).typ
    Case REF_RESEARCH
        pSubject = aStruct(this).functionality.researchFacility.psSubject
        hasSubject = (pSubject <> -1)       ' RESEARCH *: -1 = nullptr
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        pSubject = aStruct(this).functionality.factory.psSubject
        hasSubject = (pSubject <> 0)        ' DROID_TEMPLATE *: 0 = nullptr
    Case Else
    End Select

    If hasSubject Then
        STRUCTURE_isIdle = False
        Exit Function
    End If

    STRUCTURE_isIdle = True
End Function

Private Sub parseFavoriteStructs()
    Dim i As Long
    If Not asStructureStats_allocated Then
        debug_ 7234, LOG_WARNING, "parseFavoriteStructs", "asStructureStats was null?"
        Exit Sub
    End If

    For i = 0 To numStructureStats - 1
        If StringVector_contains(favoriteStructs, asStructureStats(i).id) Then
            asStructureStats(i).isFavorite = True
        Else
            asStructureStats(i).isFavorite = False
        End If
    Next
End Sub

Private Sub packFavoriteStructs()
    Dim i As Long
    If Not asStructureStats_allocated Then
        debug_ 7255, LOG_WARNING, "packFavoriteStructs", "asStructureStats was null?"
        Exit Sub
    End If

    StringVector_clear favoriteStructs

    For i = 0 To numStructureStats - 1
        If asStructureStats(i).isFavorite Then
            If Len(asStructureStats(i).id) = 0 Then
                ASSERT_FAILURE "false", "src/structure.cpp:7267", "packFavoriteStructs", "Invalid struct stats - empty id"
                GoTo continue_
            End If
            StringVector_push_back favoriteStructs, asStructureStats(i).id
        End If
continue_:
    Next
End Sub

' (VB) GlobalStructContainer().emplace(std::move(building))
Public Function GlobalStructContainer_emplace(building As STRUCTURE) As Long
    Dim psBuilding As Long
    psBuilding = StructContainer_alloc()
    aStruct(psBuilding) = building
    GlobalStructContainer_emplace = psBuilding
End Function

' STRUCTURE::STRUCTURE(uint32_t id, unsigned player), with BASE_OBJECT's and SIMPLE_OBJECT's constructors
Public Sub STRUCTURE_construct(s As STRUCTURE, ByVal id As Long, ByVal player As Long)
    Dim i As Long
    ' SIMPLE_OBJECT(type, id, player)
    s.typ = OBJ_STRUCTURE
    s.id = id
    s.player = player And &HFF&
    s.born = gameTime
    s.died = 0
    s.time_ = 0
    ' BASE_OBJECT(type, id, player)
    s.group = 0
    s.repairGroup = UBYTE_MAX
    s.selected = 0
    s.lastEmission = 0
    s.lastHitWeapon = WSC_NUM_WEAPON_SUBCLASSES  ' No such weapon.
    s.timeLastHit = UDWORD_MAX
    s.body = 0
    s.periodicalDamageStart = 0
    s.periodicalDamage = 0
    s.timeAnimationStarted = 0
    s.animationEvent = ANIM_EVENT_NONE
    For i = 0 To MAX_PLAYERS - 1
        s.visible(i) = 0
    Next
    s.sDisplay.imd = 0
    s.sDisplay.frameNumber = 0
    s.sDisplay.screenX = 0
    s.sDisplay.screenY = 0
    s.sDisplay.screenR = 0
    s.hasExtraFunction = False
    ' STRUCTURE
    s.pStructureType = -1
    s.pFunctionality = 0
    s.buildRate = 1  ' Initialise to 1 instead of 0, to make sure we don't get destroyed first tick due to inactivity.
    s.lastBuildRate = 0
    s.prebuiltImd = 0
    For i = 0 To MAX_WEAPONS - 1
        s.psTarget(i) = 0
    Next
    s.productToGroup = UBYTE_MAX

    s.pos = Vector3i_(0, 0, 0)
    s.rot = Rotation_(0, 0, 0)
    s.capacity = 0
End Sub

' Release all resources associated with a structure
' (STRUCTURE::~STRUCTURE, with ~BASE_OBJECT and ~SIMPLE_OBJECT after it)
Public Sub STRUCTURE_destruct(ByVal this As Long)
    Dim psBuilding As Long
    psBuilding = this

    ' free up the space used by the functionality array
    aStruct(psBuilding).pFunctionality = 0

    ' ~BASE_OBJECT
    visRemoveVisibility this
    Erase aStruct(this).watchedTiles
    aStruct(this).watchedTiles_size = 0
    ' ~SIMPLE_OBJECT
    aStruct(this).typ = aStruct(this).typ + 1000000000      ' Hopefully this will trigger an assert              if someone uses the freed object.
    aStruct(this).player = (aStruct(this).player + 100) And &HFF&  ' Hopefully this will trigger an assert and/or crash if someone uses the freed object.
End Sub

' (VB) psBuilding->pFunctionality = (FUNCTIONALITY *)calloc(1, sizeof(*psBuilding->pFunctionality))
Public Sub STRUCTURE_calloc_functionality(ByVal psBuilding As Long)
    Dim f As FUNCTIONALITY
    f.researchFacility.psSubject = -1
    f.researchFacility.psSubjectPending = -1
    f.researchFacility.psBestTopic = -1
    aStruct(psBuilding).functionality = f
    aStruct(psBuilding).pFunctionality = psBuilding
End Sub

Public Function IsStatExpansionModule(ByVal psStats As Long) As Boolean
    ' If the stat is any of the 3 expansion types ... then return true
    Select Case asStructureStats(psStats).typ
    Case REF_POWER_MODULE, REF_FACTORY_MODULE, REF_RESEARCH_MODULE
        IsStatExpansionModule = True
    End Select
End Function

'gives the attacking player a reward based on the type of structure that has
'been attacked
Private Function electronicReward(ByVal psStructure As Long, ByVal attackPlayer As Long) As Boolean
    Dim bRewarded As Boolean

    If Not bMultiPlayer Then
        electronicReward = False 'campaign should not give rewards (especially to the player)
        Exit Function
    End If

    If Not (attackPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "attackPlayer < MAX_PLAYERS", "src/structure.cpp:5911", "electronicReward", "Invalid player id %d", attackPlayer: electronicReward = False: Exit Function

    bRewarded = False

    Select Case asStructureStats(aStruct(psStructure).pStructureType).typ
    Case REF_RESEARCH
        researchReward aStruct(psStructure).player, attackPlayer
        bRewarded = True
    Case REF_FACTORY, REF_VTOL_FACTORY, REF_CYBORG_FACTORY
        factoryReward aStruct(psStructure).player, attackPlayer
        bRewarded = True
    Case REF_HQ
        hqReward aStruct(psStructure).player, attackPlayer
        If attackPlayer = selectedPlayer Then
            addConsoleMessage "Electronic Reward - Visibility Report", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
        End If
        bRewarded = True
    Case REF_REPAIR_FACILITY
        repairFacilityReward aStruct(psStructure).player, attackPlayer
        bRewarded = True
    Case Else
        bRewarded = False
    End Select

    electronicReward = bRewarded
End Function


'find the 'best' prop/body/weapon component the losing player has and
''give' it to the reward player
Private Sub factoryReward(ByVal losingPlayer As Long, ByVal rewardPlayer As Long)
    Dim comp As Long, inc As Long
    comp = 0

    If Not (losingPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "losingPlayer < MAX_PLAYERS", "src/structure.cpp:5953", "factoryReward", "Invalid losingPlayer id %d", losingPlayer: Exit Sub
    If Not (rewardPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "rewardPlayer < MAX_PLAYERS", "src/structure.cpp:5954", "factoryReward", "Invalid rewardPlayer id %d", rewardPlayer: Exit Sub

    'search through the propulsions first
    For inc = 0 To asPropulsionStats_size - 1
        If apCompLists(losingPlayer, COMP_PROPULSION, inc) = AVAILABLE And _
           apCompLists(rewardPlayer, COMP_PROPULSION, inc) <> AVAILABLE Then
            If asPropulsionStats(inc).buildPower > asPropulsionStats(comp).buildPower Then
                comp = inc
            End If
        End If
    Next
    If comp <> 0 Then
        apCompLists(rewardPlayer, COMP_PROPULSION, comp) = AVAILABLE
        If rewardPlayer = selectedPlayer Then
            console "%s :- %s", "Factory Reward - Propulsion", asPropulsionStats(comp).name_
        End If
        Exit Sub
    End If

    'haven't found a propulsion - look for a body
    For inc = 0 To asBodyStats_size - 1
        If apCompLists(losingPlayer, COMP_BODY, inc) = AVAILABLE And _
           apCompLists(rewardPlayer, COMP_BODY, inc) <> AVAILABLE Then
            If asBodyStats(inc).buildPower > asBodyStats(comp).buildPower Then
                comp = inc
            End If
        End If
    Next
    If comp <> 0 Then
        apCompLists(rewardPlayer, COMP_BODY, comp) = AVAILABLE
        If rewardPlayer = selectedPlayer Then
            console "%s :- %s", "Factory Reward - Body", asBodyStats(comp).name_
        End If
        Exit Sub
    End If

    'haven't found a body - look for a weapon
    For inc = 0 To asWeaponStats_size - 1
        If apCompLists(losingPlayer, COMP_WEAPON, inc) = AVAILABLE And _
           apCompLists(rewardPlayer, COMP_WEAPON, inc) <> AVAILABLE Then
            If asWeaponStats(inc).buildPower > asWeaponStats(comp).buildPower Then
                comp = inc
            End If
        End If
    Next
    If comp <> 0 Then
        apCompLists(rewardPlayer, COMP_WEAPON, comp) = AVAILABLE
        If rewardPlayer = selectedPlayer Then
            console "%s :- %s", "Factory Reward - Weapon", asWeaponStats(comp).name_
        End If
        Exit Sub
    End If

    'losing Player hasn't got anything better so don't gain anything!
    If rewardPlayer = selectedPlayer Then
        addConsoleMessage "Factory Reward - Nothing", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub

'find the 'best' repair component the losing player has and
''give' it to the reward player
Private Sub repairFacilityReward(ByVal losingPlayer As Long, ByVal rewardPlayer As Long)
    Dim comp As Long, inc As Long
    comp = 0

    If Not (losingPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "losingPlayer < MAX_PLAYERS", "src/structure.cpp:6035", "repairFacilityReward", "Invalid losingPlayer id %d", losingPlayer: Exit Sub
    If Not (rewardPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "rewardPlayer < MAX_PLAYERS", "src/structure.cpp:6036", "repairFacilityReward", "Invalid rewardPlayer id %d", rewardPlayer: Exit Sub

    'search through the repair stats
    For inc = 0 To asRepairStats_size - 1
        If apCompLists(losingPlayer, COMP_REPAIRUNIT, inc) = AVAILABLE And _
           apCompLists(rewardPlayer, COMP_REPAIRUNIT, inc) <> AVAILABLE Then
            If asRepairStats(inc).buildPower > asRepairStats(comp).buildPower Then
                comp = inc
            End If
        End If
    Next
    If comp <> 0 Then
        apCompLists(rewardPlayer, COMP_REPAIRUNIT, comp) = AVAILABLE
        If rewardPlayer = selectedPlayer Then
            console "%s :- %s", "Repair Facility Award - Repair", asRepairStats(comp).name_
        End If
        Exit Sub
    End If
    If rewardPlayer = selectedPlayer Then
        addConsoleMessage "Repair Facility Award - Nothing", DEFAULT_JUSTIFY, SYSTEM_MESSAGE
    End If
End Sub


'makes the losing players tiles/structures/features visible to the reward player
Public Sub hqReward(ByVal losingPlayer As Long, ByVal rewardPlayer As Long)
    Dim x As Long, y As Long, psTile As Long, i As Long, it As Long, psStruct As Long, psDroid As Long, psFeat As Long
    If Not (losingPlayer < MAX_PLAYERS And rewardPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "losingPlayer < MAX_PLAYERS && rewardPlayer < MAX_PLAYERS", "src/structure.cpp:6069", "hqReward", "losingPlayer (%u), rewardPlayer (%u) must both be < MAXPLAYERS", losingPlayer, rewardPlayer: Exit Sub

    ' share exploration info - pretty useless but perhaps a nice touch?
    For y = 0 To mapHeight - 1
        For x = 0 To mapWidth - 1
            psTile = mapTile(x, y)
            If TEST_TILE_VISIBLE(losingPlayer, psTile) Then
                psMapTiles(psTile).tileExploredBits = psMapTiles(psTile).tileExploredBits Or alliancebits(rewardPlayer)
            End If
        Next
    Next

    'struct
    For i = 0 To MAX_PLAYERS - 1
        it = apsStructLists(i).head
        Do While it <> 0
            psStruct = lnObj(it)
            If aStruct(psStruct).visible(losingPlayer) <> 0 And aStruct(psStruct).died = 0 Then
                aStruct(psStruct).visible(rewardPlayer) = aStruct(psStruct).visible(losingPlayer)
            End If
            it = lnNext(it)
        Loop

        'droids.
        it = apsDroidLists(i).head
        Do While it <> 0
            psDroid = lnObj(it)
            If aDroid(psDroid).visible(losingPlayer) <> 0 Or aDroid(psDroid).player = losingPlayer Then
                aDroid(psDroid).visible(rewardPlayer) = UBYTE_MAX
            End If
            it = lnNext(it)
        Loop
    Next

    'feature
    it = apsFeatureList(0).head
    Do While it <> 0
        psFeat = lnObj(it)
        If aFeature(psFeat).visible(losingPlayer) <> 0 Then
            aFeature(psFeat).visible(rewardPlayer) = aFeature(psFeat).visible(losingPlayer)
        End If
        it = lnNext(it)
    Loop
End Sub

' Return true if structure is a factory of any type.
Public Function STRUCTURE_isFactory(ByVal this As Long) As Boolean
    If Not (aStruct(this).pStructureType <> -1) Then ASSERT_FAILURE "pStructureType != nullptr", "src/structure.cpp:6120", "isFactory", "Invalid structureType!": Exit Function

    If aStruct(this).typ = OBJ_STRUCTURE Then
        Select Case asStructureStats(aStruct(this).pStructureType).typ
        Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
            STRUCTURE_isFactory = True
        End Select
    End If
End Function

' structure.h's setStructureTarget (_setStructureTarget)
Public Sub setStructureTarget(ByVal psBuilding As Long, ByVal psNewTarget As Long, ByVal idx As Long, ByVal targetOrigin As TARGET_ORIGIN)
    If Not (idx < MAX_WEAPONS) Then ASSERT_FAILURE "idx < MAX_WEAPONS", "src/structure.h:360", "_setStructureTarget", "Bad index": Exit Sub
    If psNewTarget <> 0 Then
        If Not (SIMPLE_OBJECT_died(psNewTarget) = 0) Then ASSERT_FAILURE "psNewTarget == nullptr || !psNewTarget->died", "src/structure.h:361", "_setStructureTarget", "setStructureTarget set dead target": Exit Sub
    End If
    aStruct(psBuilding).psTarget(idx) = psNewTarget
    aStruct(psBuilding).asWeaps(idx).origin = targetOrigin
End Sub

' ---- the structures

' structure.h: True iff object is a structure; STRUCTURE * if structure or NULL if not.
Public Function castStructure(ByVal psObject As Long) As Long
    If psObject <> 0 Then
        If SIMPLE_OBJECT_type(psObject) = OBJ_STRUCTURE Then castStructure = psObject
    End If
End Function

Public Function getBuildingResearchPoints(ByVal psStruct As Long) As Long
    With asStructureStats(aStruct(psStruct).pStructureType).upgrade(aStruct(psStruct).player)
        getBuildingResearchPoints = .research + .moduleResearch * aStruct(psStruct).capacity
    End With
End Function

Public Function getBuildingProductionPoints(ByVal psStruct As Long) As Long
    With asStructureStats(aStruct(psStruct).pStructureType).upgrade(aStruct(psStruct).player)
        getBuildingProductionPoints = .production + .moduleProduction * aStruct(psStruct).capacity
    End With
End Function

Public Function getBuildingPowerPoints(ByVal psStruct As Long) As Long
    With asStructureStats(aStruct(psStruct).pStructureType).upgrade(aStruct(psStruct).player)
        getBuildingPowerPoints = .power + .modulePower * aStruct(psStruct).capacity
    End With
End Function

Public Function getBuildingRepairPoints(ByVal psStruct As Long) As Long
    getBuildingRepairPoints = asStructureStats(aStruct(psStruct).pStructureType).upgrade(aStruct(psStruct).player).repair
End Function

Public Function getBuildingRearmPoints(ByVal psStruct As Long) As Long
    getBuildingRearmPoints = asStructureStats(aStruct(psStruct).pStructureType).upgrade(aStruct(psStruct).player).rearm
End Function

' Ask a gate to open: returns the time until it is open (0 if it can't open, is open, or is busy)
Public Function requestOpenGate(ByVal psStructure As Long) As Long
    If aStruct(psStructure).status <> SS_BUILT Or asStructureStats(aStruct(psStructure).pStructureType).typ <> REF_GATE Then
        requestOpenGate = 0  ' Can't open.
        Exit Function
    End If

    Select Case aStruct(psStructure).state
    Case SAS_NORMAL
        aStruct(psStructure).lastStateTime = gameTime
        aStruct(psStructure).state = SAS_OPENING
    Case SAS_OPEN
        aStruct(psStructure).lastStateTime = gameTime
        requestOpenGate = 0  ' Already open.
        Exit Function
    Case SAS_OPENING
    Case SAS_CLOSING
        aStruct(psStructure).lastStateTime = 2 * gameTime - aStruct(psStructure).lastStateTime - SAS_OPEN_SPEED
        aStruct(psStructure).state = SAS_OPENING
        requestOpenGate = 0 ' Busy
        Exit Function
    End Select

    requestOpenGate = aStruct(psStructure).lastStateTime + SAS_OPEN_SPEED - gameTime
End Function

' syncDebugStructure(psStruct, ch) is syncDebugStructure_ function, psStruct, ch
Public Sub syncDebugStructure_(ByVal function_ As String, ByVal psStruct As Long, ByVal ch As Long)
    Dim ref As Long, refChr As Long, list(0 To 11) As Long
    If aStruct(psStruct).typ <> OBJ_STRUCTURE Then
        ASSERT_FAILURE "false", "src/structure.cpp:3669", "_syncDebugStructure", "%c Broken psStruct->type %u!", ch, aStruct(psStruct).typ
        syncDebug "_syncDebugStructure", "Broken psStruct->type %u!", aStruct(psStruct).typ
        Exit Sub
    End If
    ref = 0
    refChr = Asc(" ")

    ' Print what the structure is producing (or repairing), too.
    Select Case asStructureStats(aStruct(psStruct).pStructureType).typ
    Case REF_RESEARCH
        If aStruct(psStruct).pFunctionality <> 0 Then
            If aStruct(psStruct).functionality.researchFacility.psSubject <> -1 Then
                ref = asResearch(aStruct(psStruct).functionality.researchFacility.psSubject).ref
                refChr = Asc("r")
            End If
        End If
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        If aStruct(psStruct).pFunctionality <> 0 Then
            If aStruct(psStruct).functionality.factory.psSubject <> 0 Then
                ref = apsTemplates(aStruct(psStruct).functionality.factory.psSubject).multiPlayerID
                refChr = Asc("p")
            End If
        End If
    Case REF_REPAIR_FACILITY
        If aStruct(psStruct).pFunctionality <> 0 Then
            If aStruct(psStruct).functionality.repairFacility.psObj <> 0 Then
                ref = SIMPLE_OBJECT_id(aStruct(psStruct).functionality.repairFacility.psObj)
                refChr = Asc("+")
            End If
        End If
    End Select

    list(0) = ch

    list(1) = aStruct(psStruct).id

    list(2) = aStruct(psStruct).player
    list(3) = aStruct(psStruct).pos.x
    list(4) = aStruct(psStruct).pos.y
    list(5) = aStruct(psStruct).pos.z
    list(6) = aStruct(psStruct).status
    list(7) = asStructureStats(aStruct(psStruct).pStructureType).typ
    list(8) = refChr
    list(9) = ref
    list(10) = aStruct(psStruct).currentBuildPts
    list(11) = aStruct(psStruct).body
    syncDebugIntList_ function_, "%c structure%d = p%d;pos(%d,%d,%d),status%d,type%d,%c%.0d,bld%d,body%d", list, 12
End Sub

' structure.h: CHECK_STRUCTURE(object) is checkStructure object, "<file>:<line>", "<function>", max_check_object_recursion.
' (VB) Its ASSERT_HELPERs log under the caller's location, as in C++; the [function:line] of a failure's log
' line is the caller's here (C++: checkStructure's own line).
Public Sub checkStructure(ByVal psStructure As Long, ByVal location_description As String, ByVal function_ As String, ByVal recurse As Long)
    Dim i As Long
    If recurse < 0 Then
        Exit Sub
    End If

    If Not (psStructure <> 0) Then ASSERT_HELPER_FAILURE 7136, "psStructure != nullptr", location_description, function_, "CHECK_STRUCTURE: NULL pointer": Exit Sub
    If Not (aStruct(psStructure).id <> 0) Then ASSERT_HELPER_FAILURE 7137, "psStructure->id != 0", location_description, function_, "CHECK_STRUCTURE: Structure with ID 0"
    If Not (aStruct(psStructure).typ = OBJ_STRUCTURE) Then ASSERT_HELPER_FAILURE 7138, "psStructure->type == OBJ_STRUCTURE", location_description, function_, "CHECK_STRUCTURE: No structure (type num %u)", aStruct(psStructure).typ
    If Not (aStruct(psStructure).player < MAX_PLAYERS) Then ASSERT_HELPER_FAILURE 7139, "psStructure->player < MAX_PLAYERS", location_description, function_, "CHECK_STRUCTURE: Out of bound player num (%u)", aStruct(psStructure).player
    If Not (asStructureStats(aStruct(psStructure).pStructureType).typ < NUM_DIFF_BUILDINGS) Then ASSERT_HELPER_FAILURE 7140, "psStructure->pStructureType->type < NUM_DIFF_BUILDINGS", location_description, function_, "CHECK_STRUCTURE: Out of bound structure type (%u)", asStructureStats(aStruct(psStructure).pStructureType).typ
    If Not (aStruct(psStructure).numWeaps <= MAX_WEAPONS) Then ASSERT_HELPER_FAILURE 7141, "psStructure->numWeaps <= MAX_WEAPONS", location_description, function_, "CHECK_STRUCTURE: Out of bound weapon count (%u)", aStruct(psStructure).numWeaps

    ' (checkObject of the targets: not ported yet)
End Sub

Private Sub auxStructureNonblocking(ByVal psStructure As Long)
    Dim b As StructureBounds, i As Long, j As Long, x As Long, y As Long, psTile As Long
    b = getStructureBounds_struct(psStructure)

    For i = 0 To b.size.x - 1
        For j = 0 To b.size.y - 1
            x = b.map.x + i
            y = b.map.y + j
            psTile = mapTile(x, y)
            If psMapTiles(psTile).psObject = psStructure Then
                auxClearAll x, y, AUXBITS_BLOCKING Or AUXBITS_OUR_BUILDING Or AUXBITS_NONPASSABLE
            Else
                ' Likely a script-queued object removal for a position where the script immediately replaced the old struct - just log
                debug_ 181, LOG_WZ, "auxStructureNonblocking", "Skipping blocking bit clear - structure %u is not the recorded tile object at (%d, %d)", aStruct(psStructure).id, x, y
            End If
        Next
    Next
End Sub

Private Sub auxStructureBlocking(ByVal psStructure As Long)
    Dim b As StructureBounds, i As Long, j As Long
    b = getStructureBounds_struct(psStructure)

    For i = 0 To b.size.x - 1
        For j = 0 To b.size.y - 1
            auxSetAllied b.map.x + i, b.map.y + j, aStruct(psStructure).player, AUXBITS_OUR_BUILDING
            auxSetAll b.map.x + i, b.map.y + j, AUXBITS_BLOCKING Or AUXBITS_NONPASSABLE
        Next
    Next
End Sub

Private Sub auxStructureOpenGate(ByVal psStructure As Long)
    Dim b As StructureBounds, i As Long, j As Long
    b = getStructureBounds_struct(psStructure)

    For i = 0 To b.size.x - 1
        For j = 0 To b.size.y - 1
            auxClearAll b.map.x + i, b.map.y + j, AUXBITS_BLOCKING
        Next
    Next
End Sub

Private Sub auxStructureClosedGate(ByVal psStructure As Long)
    Dim b As StructureBounds, i As Long, j As Long
    b = getStructureBounds_struct(psStructure)

    For i = 0 To b.size.x - 1
        For j = 0 To b.size.y - 1
            auxSetEnemy b.map.x + i, b.map.y + j, aStruct(psStructure).player, AUXBITS_NONPASSABLE
            auxSetAll b.map.x + i, b.map.y + j, AUXBITS_BLOCKING
        Next
    Next
End Sub

Private Function numStructureModules(ByVal psStruct As Long) As Long
    numStructureModules = aStruct(psStruct).capacity
End Function

Public Function STRUCTURE_isBlueprint(ByVal this As Long) As Boolean
    Select Case aStruct(this).status
    Case SS_BLUEPRINT_VALID, SS_BLUEPRINT_INVALID, SS_BLUEPRINT_PLANNED, SS_BLUEPRINT_PLANNED_BY_ALLY
        STRUCTURE_isBlueprint = True
    End Select
End Function

Public Function isBlueprint(ByVal psObject As Long) As Boolean
    If psObject <> 0 Then
        If SIMPLE_OBJECT_type(psObject) = OBJ_STRUCTURE Then isBlueprint = STRUCTURE_isBlueprint(psObject)
    End If
End Function

' Add smoke effect to cover the droid's emergence from the factory or when building structures.
' DISPLAY ONLY - does not affect game state.
Private Sub displayConstructionCloud(pos As Vector3i)
    Dim psTile As Long, iVecEffect As Vector3i
    psTile = mapTile(map_coord(pos.x), map_coord(pos.y))
    If Not tileIsClearlyVisible(psTile) Then
        Exit Sub
    End If

    iVecEffect.x = pos.x
    iVecEffect.y = map_Height(pos.x, pos.y) + DROID_CONSTRUCTION_SMOKE_HEIGHT
    iVecEffect.z = pos.y
    addEffect iVecEffect, EFFECT_CONSTRUCTION, CONSTRUCTION_TYPE_DRIFTING, False, 0, 0, gameTime - deltaGameTime + 1
    iVecEffect.x = pos.x - DROID_CONSTRUCTION_SMOKE_OFFSET
    iVecEffect.z = pos.y - DROID_CONSTRUCTION_SMOKE_OFFSET
    addEffect iVecEffect, EFFECT_CONSTRUCTION, CONSTRUCTION_TYPE_DRIFTING, False, 0, 0, gameTime - deltaGameTime + 1
    iVecEffect.z = pos.y + DROID_CONSTRUCTION_SMOKE_OFFSET
    addEffect iVecEffect, EFFECT_CONSTRUCTION, CONSTRUCTION_TYPE_DRIFTING, False, 0, 0, gameTime - deltaGameTime + 1
    iVecEffect.x = pos.x + DROID_CONSTRUCTION_SMOKE_OFFSET
    addEffect iVecEffect, EFFECT_CONSTRUCTION, CONSTRUCTION_TYPE_DRIFTING, False, 0, 0, gameTime - deltaGameTime + 1
    iVecEffect.z = pos.y - DROID_CONSTRUCTION_SMOKE_OFFSET
    addEffect iVecEffect, EFFECT_CONSTRUCTION, CONSTRUCTION_TYPE_DRIFTING, False, 0, 0, gameTime - deltaGameTime + 1
End Sub

Public Sub initStructLimits()
    Dim i As Long, p As Long
    For i = 0 To numStructureStats - 1
        For p = 0 To MAX_PLAYERS - 1
            asStructureStats(i).curCount(p) = 0
        Next
    Next
End Sub

Public Sub structureInitVars()
    Dim i As Long, j As Long
    If Not (Not asStructureStats_allocated) Then ASSERT_FAILURE "asStructureStats == nullptr", "src/structure.cpp:292", "structureInitVars", "Failed to cleanup prior asStructureStats?"

    Erase asStructureStats
    asStructureStats_allocated = False
    If Not lookupStructStatPtr Is Nothing Then lookupStructStatPtr.RemoveAll
    numStructureStats = 0
    factoryModuleStat = 0
    powerModuleStat = 0
    researchModuleStat = 0
    lastMaxUnitMessage = 0
    has_structureDamageBaseExperienceLevel = False

    initStructLimits
    For i = 0 To MAX_PLAYERS - 1
        droidLimit(i) = INT16_MAX
        commanderLimit(i) = INT16_MAX
        constructorLimit(i) = INT16_MAX
        For j = 0 To NUM_FLAG_TYPES - 1
            factoryNumFlag_size(i, j) = 0
        Next
    Next
    For i = 0 To MAX_PLAYERS - 1
        satUplinkExists(i) = 0
        lasSatExists(i) = 0
    Next
    'initialise the selectedPlayer's production run
    For i = 0 To NUM_FACTORY_TYPES - 1
        Erase asProductionRun(i).runs
        asProductionRun(i).size = 0
    Next
    'set up at beginning of game which player will have a production list
    productionPlayer = selectedPlayer And &HFF&
    If productionPlayer > 127 Then productionPlayer = productionPlayer - 256    ' (SBYTE)
End Sub

'initialises the flag before a new data set is loaded up
Public Sub initFactoryNumFlag()
    Dim i As Long, j As Long
    For i = 0 To MAX_PLAYERS - 1
        'initialise the flag
        For j = 0 To NUM_FLAG_TYPES - 1
            factoryNumFlag_size(i, j) = 0
        Next
    Next
End Sub

'called at start of missions
Public Sub resetFactoryNumFlag()
    Dim i As Long, typ As Long, it As Long, psStruct As Long, inc As Long, k As Long
    For i = 0 To MAX_PLAYERS - 1
        For typ = 0 To NUM_FLAG_TYPES - 1
            ' reset them all to false
            factoryNumFlag_size(i, typ) = 0
        Next
        'look through the list of structures to see which have been used
        it = apsStructLists(i).head
        Do While it <> 0
            psStruct = lnObj(it)
            Select Case asStructureStats(aStruct(psStruct).pStructureType).typ
            Case REF_FACTORY: typ = FACTORY_FLAG
            Case REF_CYBORG_FACTORY: typ = CYBORG_FLAG
            Case REF_VTOL_FACTORY: typ = VTOL_FLAG
            Case REF_REPAIR_FACILITY: typ = REPAIR_FLAG
            Case Else: GoTo continue_
            End Select

            inc = -1
            If typ = REPAIR_FLAG Then
                If aStruct(psStruct).functionality.repairFacility.psDeliveryPoint <> 0 Then
                    inc = aFlagPos(aStruct(psStruct).functionality.repairFacility.psDeliveryPoint).factoryInc
                End If
            Else
                If aStruct(psStruct).functionality.factory.psAssemblyPoint <> 0 Then
                    inc = aFlagPos(aStruct(psStruct).functionality.factory.psAssemblyPoint).factoryInc
                End If
            End If
            If inc >= 0 Then
                ' resize(std::max<size_t>(size, inc + 1), false)
                For k = factoryNumFlag_size(i, typ) To inc
                    factoryNumFlag(i, typ, k) = False
                Next
                If factoryNumFlag_size(i, typ) < inc + 1 Then factoryNumFlag_size(i, typ) = inc + 1
                factoryNumFlag(i, typ, inc) = True
            End If
continue_:
            it = lnNext(it)
        Loop
    Next
End Sub

Public Function getStructureDamage(ByVal psStructure As Long) As Long
    Dim maxBody As Long, health As Double

    checkStructure psStructure, "src/structure.cpp:882", "getStructureDamage", max_check_object_recursion

    maxBody = structureBodyBuilt(psStructure)

    ' (int64_t)65536 * psStructure->body / MAX(1, maxBody) (unsigned)
    If ULT(maxBody, 1) Then maxBody = 1
    health = Int(65536# * U32(aStruct(psStructure).body) / U32(maxBody))
    If health < 0 Then health = 0
    If health > 65536 Then health = 65536

    getStructureDamage = 65536 - health
End Function

Public Function structureBuildPointsToCompletion(ByVal structure_ As Long) As Long
    Dim moduleStat As Long
    If structureHasModules(structure_) Then
        moduleStat = getModuleStat(structure_)

        If moduleStat <> -1 Then
            structureBuildPointsToCompletion = asStructureStats(moduleStat).buildPoints
            Exit Function
        End If
    End If

    structureBuildPointsToCompletion = asStructureStats(aStruct(structure_).pStructureType).buildPoints
End Function

Public Function structureCompletionProgress(ByVal structure_ As Long) As Single
    Dim f As Single
    f = fdiv(F32(U32(aStruct(structure_).currentBuildPts)), F32(U32(structureBuildPointsToCompletion(structure_))))
    ' MIN(1, f)
    If IsNaNF(f) Then
        structureCompletionProgress = f
    ElseIf 1! < f Then
        structureCompletionProgress = 1!
    Else
        structureCompletionProgress = f
    End If
End Function

Private Function structureHasModules(ByVal psStruct As Long) As Boolean
    structureHasModules = aStruct(psStruct).capacity <> 0
End Function

Private Function wallDir(ByVal orient As Long) As Long
    Dim dirs As Variant
    '                 0       1       2       3       4       5       6       7       8       9       A       B       C       D       E       F
    dirs = Array(DEG(0), DEG(0), DEG(180), DEG(0), DEG(270), DEG(0), DEG(270), DEG(0), DEG(90), DEG(90), DEG(180), DEG(180), DEG(270), DEG(90), DEG(270), DEG(0))  ' DEG(90) = rotate ccw, DEG(270) = rotate cw
    wallDir = dirs(orient) And &HFFFF&
End Function

Private Function wallType(ByVal orient As Long) As Long
    Dim types As Variant
    '             0  1  2  3  4  5  6  7  8  9  A  B  C  D  E  F
    types = Array(0, 0, 0, 0, 0, 3, 3, 2, 0, 3, 3, 2, 0, 2, 2, 1)
    wallType = types(orient)
End Function

' look at where other walls are to decide what type of wall to build
Private Function structWallScan(aWallPresent() As Boolean, ByVal x As Long, ByVal y As Long) As Long
    Dim l As Long, r As Long, u As Long, d As Long
    If aWallPresent(x - 1, y) Then l = WallConnectLeft
    If aWallPresent(x + 1, y) Then r = WallConnectRight
    If aWallPresent(x, y - 1) Then u = WallConnectUp
    If aWallPresent(x, y + 1) Then d = WallConnectDown
    structWallScan = l Or r Or u Or d
End Function

Private Function isWallCombiningStructureType(ByVal pStructureType As Long) As Boolean
    Select Case asStructureStats(pStructureType).typ
    Case REF_WALL, REF_GATE, REF_WALLCORNER
        isWallCombiningStructureType = True
    Case Else
        isWallCombiningStructureType = asStructureStats(pStructureType).combinesWithWall  ' hardpoints and fortresses by default
    End Select
End Function

Public Function isWall(ByVal typ As STRUCTURE_TYPE) As Boolean
    isWall = typ = REF_WALL Or typ = REF_WALLCORNER
End Function

Public Function isBuildableOnWalls(ByVal typ As STRUCTURE_TYPE) As Boolean
    isBuildableOnWalls = typ = REF_DEFENSE Or typ = REF_GATE
End Function

Private Sub structFindWalls(ByVal player As Long, ByVal mapX As Long, ByVal mapY As Long, aWallPresent() As Boolean, apsStructs() As Long)
    Dim x As Long, y As Long, psStruct As Long
    For y = -2 To 2
        For x = -2 To 2
            psStruct = castStructure(psMapTiles(mapTile(mapX + x, mapY + y)).psObject)
            If psStruct <> 0 Then
                If isWallCombiningStructureType(aStruct(psStruct).pStructureType) And ULT(player, MAX_PLAYERS) Then
                    If aiCheckAlliances(player, aStruct(psStruct).player) Then
                        aWallPresent(x + 2, y + 2) = True
                        apsStructs(x + 2, y + 2) = psStruct
                    End If
                End If
            End If
        Next
    Next
    ' add in the wall about to be built
    aWallPresent(2, 2) = True
End Sub

Private Function wallBlockingTerrainJoin(ByVal mapX As Long, ByVal mapY As Long) As Boolean
    Dim psTile As Long
    psTile = mapTile(mapX, mapY)
    wallBlockingTerrainJoin = terrainType(psTile) = TER_WATER Or terrainType(psTile) = TER_CLIFFFACE Or psMapTiles(psTile).psObject <> 0
End Function

Private Function structWallScanTerrain(aWallPresent() As Boolean, ByVal mapX As Long, ByVal mapY As Long) As Long
    Dim orientation As Long
    orientation = structWallScan(aWallPresent, 2, 2)

    If orientation = WallConnectNone Then
        ' If neutral, try choosing horizontal or vertical based on terrain, but don't change to corner type.
        aWallPresent(2, 1) = wallBlockingTerrainJoin(mapX, mapY - 1)
        aWallPresent(2, 3) = wallBlockingTerrainJoin(mapX, mapY + 1)
        aWallPresent(1, 2) = wallBlockingTerrainJoin(mapX - 1, mapY)
        aWallPresent(3, 2) = wallBlockingTerrainJoin(mapX + 1, mapY)
        orientation = structWallScan(aWallPresent, 2, 2)
        If (orientation And (WallConnectLeft Or WallConnectRight)) <> 0 And (orientation And (WallConnectUp Or WallConnectDown)) <> 0 Then
            orientation = WallConnectNone
        End If
    End If

    structWallScanTerrain = orientation
End Function

Private Sub structFindWallBlueprints(ByVal mapX As Long, ByVal mapY As Long, aWallPresent() As Boolean)
    Dim x As Long, y As Long, stats As Long
    For y = -2 To 2
        For x = -2 To 2
            stats = getTileBlueprintStats(mapX + x, mapY + y)
            If stats <> -1 Then
                If isWallCombiningStructureType(stats) Then
                    aWallPresent(x + 2, y + 2) = True
                End If
            End If
        Next
    Next
End Sub

Private Function structChooseWallTypeBlueprint(ByVal mapX As Long, ByVal mapY As Long) As Long
    Dim aWallPresent(0 To 4, 0 To 4) As Boolean, apsStructs(0 To 4, 0 To 4) As Long

    ' scan around the location looking for walls
    ' (memset(aWallPresent, 0, ...): a VB array starts zeroed)
    structFindWalls selectedPlayer, mapX, mapY, aWallPresent, apsStructs
    structFindWallBlueprints mapX, mapY, aWallPresent

    ' finally return the type for this wall
    structChooseWallTypeBlueprint = structWallScanTerrain(aWallPresent, mapX, mapY)
End Function

' Choose a type of wall for a location - and update any neighbouring walls
Private Function structChooseWallType(ByVal player As Long, ByVal mapX As Long, ByVal mapY As Long) As Long
    Dim aWallPresent(0 To 4, 0 To 4) As Boolean, psStruct As Long, apsStructs(0 To 4, 0 To 4) As Long
    Dim x As Long, y As Long, scanType As Long, n As Long

    ' scan around the location looking for walls
    structFindWalls player, mapX, mapY, aWallPresent, apsStructs

    ' now make sure that all the walls around this one are OK
    For x = 1 To 3
        For y = 1 To 3
            ' do not look at walls diagonally from this wall
            If ((x = 2 And y <> 2) Or (x <> 2 And y = 2)) And aWallPresent(x, y) Then
                ' figure out what type the wall currently is
                psStruct = apsStructs(x, y)
                If asStructureStats(aStruct(psStruct).pStructureType).typ <> REF_WALL And asStructureStats(aStruct(psStruct).pStructureType).typ <> REF_GATE Then
                    ' do not need to adjust anything apart from walls
                    GoTo continue_
                End If

                ' see what type the wall should be
                scanType = structWallScan(aWallPresent, x, y)

                ' Got to change the wall
                If scanType <> WallConnectNone Then
                    aStruct(psStruct).functionality.wall.typ = wallType(scanType)
                    aStruct(psStruct).rot.direction = wallDir(scanType)
                    n = asStructureStats(aStruct(psStruct).pStructureType).pIMD_size - 1
                    If ULT(aStruct(psStruct).functionality.wall.typ, n) Then n = aStruct(psStruct).functionality.wall.typ
                    aStruct(psStruct).sDisplay.imd = asStructureStats(aStruct(psStruct).pStructureType).pIMD(n)
                End If
            End If
continue_:
        Next
    Next

    ' finally return the type for this wall
    structChooseWallType = structWallScanTerrain(aWallPresent, mapX, mapY)
End Function

' For now all this does is work out what height the terrain needs to be set to
' An actual foundation structure may end up being placed down
' The x and y passed in are the CENTRE of the structure
Private Function foundationHeight(ByVal psStruct As Long) As Long
    Dim b As StructureBounds, foundationMin As Long, foundationMax As Long, breadth As Long, width As Long, height As Long
    b = getStructureBounds_struct(psStruct)

    'check the terrain is the correct type return -1 if not

    'may also have to check that overlapping terrain can be set to the average height
    'eg water - don't want it to 'flow' into the structure if this effect is coded!

    'initialise the starting values so they get set in loop
    foundationMin = INT32_MAX
    foundationMax = INT32_MIN

    For breadth = 0 To b.size.y
        For width = 0 To b.size.x
            height = map_TileHeight(b.map.x + width, b.map.y + breadth)
            foundationMin = MIN_(foundationMin, height)
            foundationMax = MAX_(foundationMax, height)
        Next
    Next
    'return the average of max/min height
    foundationHeight = (foundationMin + foundationMax) \ 2
End Function

Private Sub buildFlatten(ByVal pStructure As Long, ByVal h As Long)
    Dim b As StructureBounds, breadth As Long, width As Long
    b = getStructureBounds_struct(pStructure)

    For breadth = 0 To b.size.y
        For width = 0 To b.size.x
            setTileHeight b.map.x + width, b.map.y + breadth, h
            ' We need to raise features on raised tiles to the new height
            If TileHasFeature(mapTile(b.map.x + width, b.map.y + breadth)) Then
                aFeature(getTileFeature(b.map.x + width, b.map.y + breadth)).pos.z = h
            End If
        Next
    Next
End Sub

Private Function isPulledToTerrain(ByVal psBuilding As Long) As Boolean
    Select Case asStructureStats(aStruct(psBuilding).pStructureType).typ
    Case REF_DEFENSE, REF_GATE, REF_WALL, REF_WALLCORNER, REF_REARM_PAD
        isPulledToTerrain = True
    End Select
End Function

Public Sub alignStructure(ByVal psBuilding As Long)
    Dim mapH As Long, b As StructureBounds, breadth As Long, width As Long, neighbourStructure As Long
    Dim s As Long, dirX As Long, dirY As Long, p1x As Long, p1y As Long, p2x As Long, p2y As Long
    Dim maxX As Long, maxY As Long, maxZ As Long, minX As Long, minY As Long, minZ As Long
    Dim h1 As Long, h2 As Long, h3 As Long, h4 As Long, minH As Long, maxH As Long

    ' DEFENSIVE structures are pulled to the terrain
    If Not isPulledToTerrain(psBuilding) Then
        mapH = foundationHeight(psBuilding)

        buildFlatten psBuilding, mapH
        aStruct(psBuilding).pos.z = mapH
        aStruct(psBuilding).foundationDepth = aStruct(psBuilding).pos.z

        ' Align surrounding structures.
        b = getStructureBounds_struct(psBuilding)
        syncDebug "alignStructure", "Flattened (%d+%d, %d+%d) to %d for %d(p%d)", b.map.x, b.size.x, b.map.y, b.size.y, mapH, aStruct(psBuilding).id, aStruct(psBuilding).player
        For breadth = -1 To b.size.y
            For width = -1 To b.size.x
                neighbourStructure = castStructure(psMapTiles(mapTile(b.map.x + width, b.map.y + breadth)).psObject)
                If neighbourStructure <> 0 Then
                    If isPulledToTerrain(neighbourStructure) Then
                        alignStructure neighbourStructure  ' Recursive call, but will go to the else case, so will not re-recurse.
                    End If
                End If
            Next
        Next
    Else
        ' TODO: Does this need iIMDBaseShape, or is this truly visual only?
        ' Sample points around the structure to find a good depth for the foundation
        s = aStruct(psBuilding).sDisplay.imd

        aStruct(psBuilding).pos.z = TILE_MIN_HEIGHT
        aStruct(psBuilding).foundationDepth = TILE_MAX_HEIGHT

        ' iSinCosR(direction, 1)
        dirX = iSinR(aStruct(psBuilding).rot.direction, 1)
        dirY = iCosR(aStruct(psBuilding).rot.direction, 1)
        ' Rotate s->max.{x, z} and s->min.{x, z} by angle rot.direction.
        iIMDBaseShape_max s, maxX, maxY, maxZ
        iIMDBaseShape_min s, minX, minY, minZ
        p1x = maxX * dirY - maxZ * dirX
        p1y = maxX * dirX + maxZ * dirY
        p2x = minX * dirY - minZ * dirX
        p2y = minX * dirX + minZ * dirY

        h1 = map_Height(aStruct(psBuilding).pos.x + p1x, aStruct(psBuilding).pos.y + p2y)
        h2 = map_Height(aStruct(psBuilding).pos.x + p1x, aStruct(psBuilding).pos.y + p1y)
        h3 = map_Height(aStruct(psBuilding).pos.x + p2x, aStruct(psBuilding).pos.y + p1y)
        h4 = map_Height(aStruct(psBuilding).pos.x + p2x, aStruct(psBuilding).pos.y + p2y)
        minH = MIN_(MIN_(MIN_(h1, h2), h3), h4)
        maxH = MAX_(MAX_(MAX_(h1, h2), h3), h4)
        aStruct(psBuilding).pos.z = MAX_(aStruct(psBuilding).pos.z, maxH)
        If minH < aStruct(psBuilding).foundationDepth Then aStruct(psBuilding).foundationDepth = F32(minH)
        syncDebug "alignStructure", "minH=%d,maxH=%d,pointHeight=%d", minH, maxH, aStruct(psBuilding).pos.z  ' s->max is based on floats! If this causes desynchs, need to fix!
    End If
End Sub

' STRUCTURE *buildStructure(STRUCTURE_STATS *pStructureType, UDWORD x, UDWORD y, UDWORD player, bool FromSave)
Public Function buildStructure(ByVal pStructureType As Long, ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal FromSave As Boolean) As Long
    buildStructure = buildStructureDir(pStructureType, x, y, 0, player, FromSave, generateSynchronisedObjectId())
End Function

' The two overloads of buildStructureDir: the id missing, a new synchronised one, as the six-argument one takes.
Public Function buildStructureDir(ByVal pStructureType As Long, ByVal x As Long, ByVal y As Long, ByVal direction As Long, ByVal player As Long, ByVal FromSave As Boolean, Optional ByVal id As Variant) As Long
    Dim psBuilding As Long, size As Vector2i, preScrollMinX As Long, preScrollMinY As Long, preScrollMaxX As Long, preScrollMaxY As Long
    Dim max As Long, i As Long, wallOrientation As Long, dy As Long, dx As Long, posX As Long, posY As Long
    Dim building As STRUCTURE, mapX As Long, mapY As Long, psFeature As Long, tileX As Long, tileY As Long, psTile As Long
    Dim weapon_ As Long, vPlayer As Long, bounds As StructureBounds, playerNum As Long, it As Long, psStruct As Long, fp As Long
    Dim fpX As Long, fpY As Long, bUpgraded As Boolean, bodyDiff As Long, prevResearchState As Long, imdIndex As Long
    Dim humanStr As String

    If IsMissing(id) Then id = generateSynchronisedObjectId()

    size = STRUCTURE_STATS_size(pStructureType, direction)

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:1509", "buildStructureDir", "Cannot build structure for player %u (>= MAX_PLAYERS)", player: Exit Function
    If pStructureType = -1 Then ASSERT_FAILURE "pStructureType && pStructureType->type != REF_DEMOLISH", "src/structure.cpp:1510", "buildStructureDir", "You cannot build demolition!": Exit Function
    If Not (asStructureStats(pStructureType).typ <> REF_DEMOLISH) Then ASSERT_FAILURE "pStructureType && pStructureType->type != REF_DEMOLISH", "src/structure.cpp:1510", "buildStructureDir", "You cannot build demolition!": Exit Function

    If IsStatExpansionModule(pStructureType) = False Then
        preScrollMinX = 0: preScrollMinY = 0: preScrollMaxX = 0: preScrollMaxY = 0
        max = pStructureType

        If Not (Not ULT(numStructureStats, max)) Then ASSERT_FAILURE "max <= numStructureStats", "src/structure.cpp:1518", "buildStructureDir", "Invalid structure type": Exit Function

        ' Don't allow more than interface limits
        If ULT(asStructureStats(max).upgrade(player).limit, asStructureStats(max).curCount(player) + 1) Then
            debug_ 1523, LOG_ERROR, "buildStructureDir", "Player %u: Building %s could not be built due to building limits (has %u, max %u)!", _
                   player, asStructureStats(pStructureType).name_, asStructureStats(max).curCount(player), _
                   asStructureStats(max).upgrade(player).limit
            Exit Function
        End If

        ' snap the coords to a tile
        x = (x And Not TILE_MASK) + (size.x Mod 2) * TILE_UNITS \ 2
        y = (y And Not TILE_MASK) + (size.y Mod 2) * TILE_UNITS \ 2

        'check not trying to build too near the edge
        If map_coord(x) < TOO_NEAR_EDGE Or map_coord(x) > (mapWidth - TOO_NEAR_EDGE) Then
            debug_ 1536, LOG_WARNING, "buildStructureDir", "attempting to build too closely to map-edge, x coord (%u) too near edge (req. distance is %u)", x, TOO_NEAR_EDGE
            Exit Function
        End If
        If map_coord(y) < TOO_NEAR_EDGE Or map_coord(y) > (mapHeight - TOO_NEAR_EDGE) Then
            debug_ 1542, LOG_WARNING, "buildStructureDir", "attempting to build too closely to map-edge, y coord (%u) too near edge (req. distance is %u)", y, TOO_NEAR_EDGE
            Exit Function
        End If

        wallOrientation = WallConnectNone
        If Not FromSave And isWallCombiningStructureType(pStructureType) Then
            For dy = 0 To size.y - 1
                For dx = 0 To size.x - 1
                    ' map_coord(Vector2i(x, y) - size * TILE_UNITS / 2) + Vector2i(dx, dy)
                    posX = map_coord(x - size.x * TILE_UNITS \ 2) + dx
                    posY = map_coord(y - size.y * TILE_UNITS \ 2) + dy
                    wallOrientation = structChooseWallType(player, posX, posY)  ' This makes neighbouring walls match us, even if we're a hardpoint, not a wall.
                Next
            Next
        End If

        ' initialize the structure object
        STRUCTURE_construct building, CLng(id), player

        'fill in other details
        building.pStructureType = pStructureType

        building.pos.x = x
        building.pos.y = y
        building.rot.direction = snapDirection(direction)
        building.rot.pitch = 0
        building.rot.roll = 0

        'This needs to be done before the functionality bit...
        'load into the map data and structure list if not an upgrade
        mapX = map_coord(x) - size.x \ 2
        mapY = map_coord(y) - size.y \ 2

        'set up the imd to use for the display
        building.sDisplay.imd = asStructureStats(pStructureType).pIMD(0)

        building.state = SAS_NORMAL
        building.lastStateTime = gameTime

        ' if resource extractor - need to remove oil feature first, but do not do any
        ' consistency checking here - save games do not have any feature to remove
        ' to remove when placing oil derricks!
        If asStructureStats(pStructureType).typ = REF_RESOURCE_EXTRACTOR Then
            psFeature = getTileFeature(map_coord(x), map_coord(y))

            If psFeature <> 0 Then
                If asFeatureStats(aFeature(psFeature).psStats).subType = FEAT_OIL_RESOURCE Then
                    If fireOnLocation(aFeature(psFeature).pos.x, aFeature(psFeature).pos.y) Then
                        ' Can't build on burning oil resource
                        Exit Function
                    End If
                    ' remove it from the map
                    turnOffMultiMsg True ' don't send this one!
                    removeFeature psFeature
                    turnOffMultiMsg False
                End If
            End If
        End If

        For tileY = mapY To mapY + size.y - 1
            For tileX = mapX To mapX + size.x - 1
                psTile = mapTile(tileX, tileY)

                ' Remove any walls underneath the building. You can build defense buildings on top
                ' of walls, you see. This is not the place to test whether we own it!
                If isBuildableOnWalls(asStructureStats(pStructureType).typ) And TileHasWall(psTile) Then
                    removeStruct psMapTiles(psTile).psObject, True
                ElseIf TileHasStructure(psTile) Then
                    If Not wzapi_scriptIsObjectQueuedForRemoval(psMapTiles(psTile).psObject) Then
                        If isHumanPlayer(player) Then humanStr = "Human" Else humanStr = "AI"
                        debug_ 1619, LOG_ERROR, "buildStructureDir", "Player %u (%s): is building %s at (%d, %d) but found %s already at (%d, %d)", _
                               player, humanStr, asStructureStats(pStructureType).name_, mapX, mapY, _
                               asStructureStats(aStruct(getTileStructure(tileX, tileY)).pStructureType).name_, tileX, tileY
                        Exit Function
                    End If
                End If
            Next
        Next
        ' Emplace the structure being built in the global storage to obtain stable address.
        psBuilding = GlobalStructContainer_emplace(building)
        For tileY = mapY To mapY + size.y - 1
            For tileX = mapX To mapX + size.x - 1
                ' We now know the previous loop didn't return early, so it is safe to save references to `stableBuilding` now.
                psTile = mapTile(tileX, tileY)
                psMapTiles(psTile).psObject = psBuilding

                ' if it's a tall structure then flag it in the map.
                If aStruct(psBuilding).sDisplay.imd <> 0 Then
                    If imdMaxY(aStruct(psBuilding).sDisplay.imd) > TALLOBJECT_YMAX Then
                        auxSetBlocking tileX, tileY, AIR_BLOCKED
                    End If
                End If
            Next
        Next

        Select Case asStructureStats(pStructureType).typ
        Case REF_REARM_PAD
            ' Not blocking.
        Case Else
            auxStructureBlocking psBuilding
        End Select

        'set up the rest of the data
        For i = 0 To MAX_WEAPONS - 1
            aStruct(psBuilding).asWeaps(i).rot.direction = 0
            aStruct(psBuilding).asWeaps(i).rot.pitch = 0
            aStruct(psBuilding).asWeaps(i).rot.roll = 0
            aStruct(psBuilding).asWeaps(i).prevRot = aStruct(psBuilding).asWeaps(i).rot
            aStruct(psBuilding).asWeaps(i).origin = ORIGIN_UNKNOWN
            aStruct(psBuilding).psTarget(i) = 0
        Next

        aStruct(psBuilding).periodicalDamageStart = 0
        aStruct(psBuilding).periodicalDamage = 0

        aStruct(psBuilding).status = SS_BEING_BUILT
        aStruct(psBuilding).currentBuildPts = 0

        alignStructure psBuilding

        ' Store the weapons
        aStruct(psBuilding).numWeaps = 0
        If asStructureStats(pStructureType).numWeaps > 0 Then
            For weapon_ = 0 To asStructureStats(pStructureType).numWeaps - 1
                If asStructureStats(pStructureType).psWeapStat(weapon_) <> -1 Then
                    aStruct(psBuilding).asWeaps(weapon_).lastFired = 0
                    aStruct(psBuilding).asWeaps(weapon_).shotsFired = 0
                    'in multiPlayer make the Las-Sats require re-loading from the start
                    If bMultiPlayer Then
                        If asWeaponStats(asStructureStats(pStructureType).psWeapStat(0)).weaponSubClass = WSC_LAS_SAT Then
                            aStruct(psBuilding).asWeaps(0).lastFired = gameTime
                        End If
                    End If
                    aStruct(psBuilding).asWeaps(weapon_).nStat = asStructureStats(pStructureType).psWeapStat(weapon_)
                    aStruct(psBuilding).asWeaps(weapon_).ammo = asWeaponStats(aStruct(psBuilding).asWeaps(weapon_).nStat).upgrade(aStruct(psBuilding).player).numRounds
                    aStruct(psBuilding).numWeaps = aStruct(psBuilding).numWeaps + 1
                End If
            Next
        Else
            If asStructureStats(pStructureType).psWeapStat(0) <> -1 Then
                aStruct(psBuilding).asWeaps(0).lastFired = 0
                aStruct(psBuilding).asWeaps(0).shotsFired = 0
                'in multiPlayer make the Las-Sats require re-loading from the start
                If bMultiPlayer Then
                    If asWeaponStats(asStructureStats(pStructureType).psWeapStat(0)).weaponSubClass = WSC_LAS_SAT Then
                        aStruct(psBuilding).asWeaps(0).lastFired = gameTime
                    End If
                End If
                aStruct(psBuilding).asWeaps(0).nStat = asStructureStats(pStructureType).psWeapStat(0)
                aStruct(psBuilding).asWeaps(0).ammo = asWeaponStats(aStruct(psBuilding).asWeaps(0).nStat).upgrade(aStruct(psBuilding).player).numRounds
            End If
        End If

        aStruct(psBuilding).resistance = structureResistance(pStructureType, player And &HFF&) And &HFFFF&
        aStruct(psBuilding).lastResistance = ACTION_START_TIME

        ' Do the visibility stuff before setFunctionality - so placement of DP's can work
        For vPlayer = 0 To MAX_PLAYERS - 1
            aStruct(psBuilding).seenThisTick(vPlayer) = 0
        Next

        ' Structure is visible to anyone with shared vision.
        For vPlayer = 0 To MAX_PLAYERS - 1
            If hasSharedVision(vPlayer, player) Then aStruct(psBuilding).visible(vPlayer) = UINT8_MAX Else aStruct(psBuilding).visible(vPlayer) = 0
        Next

        ' Reveal any tiles that can be seen by the structure
        visTilesUpdate psBuilding

        'if we're coming from a SAVEGAME and we're on an Expand_Limbo mission,
        'any factories that were built previously for the selectedPlayer will
        'have DP's in an invalid location - the scroll limits will have been
        'changed to not include them. This is the only HACK I can think of to
        'enable them to be loaded up. So here goes...
        If FromSave And player = selectedPlayer And missionLimboExpand() Then
            'save the current values
            preScrollMinX = scrollMinX
            preScrollMinY = scrollMinY
            preScrollMaxX = scrollMaxX
            preScrollMaxY = scrollMaxY
            'set the current values to mapWidth/mapHeight
            scrollMinX = 0
            scrollMinY = 0
            scrollMaxX = mapWidth
            scrollMaxY = mapHeight
            ' NOTE: resizeRadar() may be required here, since we change scroll limits?
        End If
        'set the functionality dependent on the type of structure
        If Not setFunctionality(psBuilding, asStructureStats(pStructureType).typ) Then
            removeStructFromMap psBuilding
            objmemDestroy psBuilding, False
            'better reset these if you couldn't build the structure!
            If FromSave And player = selectedPlayer And missionLimboExpand() Then
                'reset the current values
                scrollMinX = preScrollMinX
                scrollMinY = preScrollMinY
                scrollMaxX = preScrollMaxX
                scrollMaxY = preScrollMaxY
                ' NOTE: resizeRadar() may be required here, since we change scroll limits?
            End If
            Exit Function
        End If

        'reset the scroll values if adjusted
        If FromSave And player = selectedPlayer And missionLimboExpand() Then
            'reset the current values
            scrollMinX = preScrollMinX
            scrollMinY = preScrollMinY
            scrollMaxX = preScrollMaxX
            scrollMaxY = preScrollMaxY
            ' NOTE: resizeRadar() may be required here, since we change scroll limits?
        End If

        ' rotate a wall if necessary
        If Not FromSave And (asStructureStats(pStructureType).typ = REF_WALL Or asStructureStats(pStructureType).typ = REF_GATE) Then
            aStruct(psBuilding).functionality.wall.typ = wallType(wallOrientation)
            If wallOrientation <> WallConnectNone Then
                aStruct(psBuilding).rot.direction = wallDir(wallOrientation)
                imdIndex = asStructureStats(aStruct(psBuilding).pStructureType).pIMD_size - 1
                If ULT(aStruct(psBuilding).functionality.wall.typ, imdIndex) Then imdIndex = aStruct(psBuilding).functionality.wall.typ
                aStruct(psBuilding).sDisplay.imd = asStructureStats(aStruct(psBuilding).pStructureType).pIMD(imdIndex)
            End If
        End If

        aStruct(psBuilding).body = STRUCTURE_structureBody(psBuilding) And &HFFFF&   ' (UWORD)
        aStruct(psBuilding).expectedDamage = 0  ' Begin life optimistically.

        'add the structure to the list - this enables it to be drawn whilst being built
        addStructure psBuilding

        asStructureStats(max).curCount(player) = asStructureStats(max).curCount(player) + 1

        If isLasSat(aStruct(psBuilding).pStructureType) Then
            aStruct(psBuilding).asWeaps(0).ammo = 1 ' ready to trigger the fire button
        End If

        ' Move any delivery points under the new structure.
        bounds = getStructureBounds_struct(psBuilding)
        For playerNum = 0 To MAX_PLAYERS - 1
            it = apsStructLists(playerNum).head
            Do While it <> 0
                psStruct = lnObj(it)
                If psStruct <> 0 Then
                    fp = 0
                    If STRUCTURE_isFactory(psStruct) Then
                        fp = aStruct(psStruct).functionality.factory.psAssemblyPoint
                    ElseIf asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REPAIR_FACILITY Then
                        fp = aStruct(psStruct).functionality.repairFacility.psDeliveryPoint
                    End If
                    If fp <> 0 Then
                        fpX = map_coord(aFlagPos(fp).coords.x)
                        fpY = map_coord(aFlagPos(fp).coords.y)
                        If ULT(fpX - bounds.map.x, bounds.size.x) And ULT(fpY - bounds.map.y, bounds.size.y) Then
                            ' Delivery point fp is under the new structure. Need to move it.
                            setAssemblyPoint fp, aFlagPos(fp).coords.x, aFlagPos(fp).coords.y, playerNum, True
                        End If
                    End If
                End If
                it = lnNext(it)
            Loop
        Next

        If Not FromSave Then
            displayConstructionCloud aStruct(psBuilding).pos
        End If
    Else 'its an upgrade
        bUpgraded = False
        bodyDiff = 0

        'don't create the Structure use existing one
        psBuilding = getTileStructure(map_coord(x), map_coord(y))

        If psBuilding = 0 Then
            Exit Function
        End If

        If Not (aStruct(psBuilding).player = player) Then ASSERT_FAILURE "psBuilding->player == player", "src/structure.cpp:1851", "buildStructureDir", "Trying to upgrade player %u building with player %u module?", aStruct(psBuilding).player, player

        prevResearchState = intGetResearchState()

        If asStructureStats(pStructureType).typ = REF_FACTORY_MODULE Then
            If asStructureStats(aStruct(psBuilding).pStructureType).typ <> REF_FACTORY And _
               asStructureStats(aStruct(psBuilding).pStructureType).typ <> REF_VTOL_FACTORY Then
                debug_ 1860, LOG_INFO, "buildStructureDir", "Factory module %u trying to upgrade structure of type %u - ignoring", CLng(id), asStructureStats(aStruct(psBuilding).pStructureType).typ
                Exit Function
            End If
            'increment the capacity and output for the owning structure
            If aStruct(psBuilding).capacity < SIZE_SUPER_HEAVY Then
                'store the % difference in body points before upgrading
                bodyDiff = 65536 - getStructureDamage(psBuilding)

                aStruct(psBuilding).capacity = (aStruct(psBuilding).capacity + 1) And &HFF&
                bUpgraded = True
                'put any production on hold
                holdProduction psBuilding, ModeImmediate
            End If
        End If

        If asStructureStats(pStructureType).typ = REF_RESEARCH_MODULE Then
            If asStructureStats(aStruct(psBuilding).pStructureType).typ <> REF_RESEARCH Then
                debug_ 1880, LOG_INFO, "buildStructureDir", "Research module %u trying to upgrade structure of type %u - ignoring", CLng(id), asStructureStats(aStruct(psBuilding).pStructureType).typ
                Exit Function
            End If
            'increment the capacity and research points for the owning structure
            If aStruct(psBuilding).capacity = 0 Then
                'store the % difference in body points before upgrading
                bodyDiff = 65536 - getStructureDamage(psBuilding)

                aStruct(psBuilding).capacity = (aStruct(psBuilding).capacity + 1) And &HFF&
                bUpgraded = True
                'cancel any research - put on hold now
                If aStruct(psBuilding).functionality.researchFacility.psSubject <> -1 Then
                    'cancel the topic
                    holdResearch psBuilding, ModeImmediate
                End If
            End If
        End If

        If asStructureStats(pStructureType).typ = REF_POWER_MODULE Then
            If asStructureStats(aStruct(psBuilding).pStructureType).typ <> REF_POWER_GEN Then
                debug_ 1904, LOG_INFO, "buildStructureDir", "Power module %u trying to upgrade structure of type %u - ignoring", CLng(id), asStructureStats(aStruct(psBuilding).pStructureType).typ
                Exit Function
            End If
            'increment the capacity and research points for the owning structure
            If aStruct(psBuilding).capacity = 0 Then
                'store the % difference in body points before upgrading
                bodyDiff = 65536 - getStructureDamage(psBuilding)

                'increment the power output, multiplier and capacity
                aStruct(psBuilding).capacity = (aStruct(psBuilding).capacity + 1) And &HFF&
                bUpgraded = True

                'need to inform any res Extr associated that not digging until complete
                releasePowerGen psBuilding
            End If
        End If

        If bUpgraded Then
            ' *2-1 because even-numbered IMDs are structures, odd-numbered IMDs are just the modules, and we want just the module since we cache the fully-built part of the building in psBuilding->prebuiltImd.
            imdIndex = MIN_(aStruct(psBuilding).capacity * 2, asStructureStats(aStruct(psBuilding).pStructureType).pIMD_size - 1) - 1
            If imdIndex < 0 Then
                ' Looks like we don't have a model for this structure's upgrade
                ' Log it and default to the base model (to avoid a crash)
                debug_ 1930, LOG_ERROR, "buildStructureDir", "No upgraded structure model to draw."
                imdIndex = 0
            End If
            aStruct(psBuilding).prebuiltImd = aStruct(psBuilding).sDisplay.imd
            aStruct(psBuilding).sDisplay.imd = asStructureStats(aStruct(psBuilding).pStructureType).pIMD(imdIndex)

            'calculate the new body points of the owning structure
            aStruct(psBuilding).body = FromU32(Int(U32(STRUCTURE_structureBody(psBuilding)) * bodyDiff / 65536#))

            'initialise the build points
            aStruct(psBuilding).currentBuildPts = 0
            'start building again
            aStruct(psBuilding).status = SS_BEING_BUILT
            aStruct(psBuilding).buildRate = 1  ' Don't abandon the structure first tick, so set to nonzero.

            If Not FromSave Then
                triggerEventStructureUpgradeStarted psBuilding

                If aStruct(psBuilding).player = selectedPlayer Then
                    intRefreshScreen
                End If
            End If
        Else
            debug_ 1957, LOG_INFO, "buildStructureDir", "Did not use upgrade %u trying to upgrade structure of type %u (either reached capacity or not compatible)", CLng(id), asStructureStats(aStruct(psBuilding).pStructureType).typ
        End If
        intNotifyResearchButton prevResearchState
    End If
    If asStructureStats(pStructureType).typ <> REF_WALL And asStructureStats(pStructureType).typ <> REF_WALLCORNER Then
        If player = selectedPlayer Then
            scoreUpdateVar WD_STR_BUILT
        End If
    End If

    ' why is this necessary - it makes tiles under the structure visible
    setUnderTilesVis psBuilding, player

    aStruct(psBuilding).prevTime = gameTime - deltaGameTime  ' Structure hasn't been updated this tick, yet.
    aStruct(psBuilding).time_ = aStruct(psBuilding).prevTime - 1  ' -1, so the times are different, even before updating.

    buildStructureDir = psBuilding
End Function

Private Function defaultAssemblyPointPos(ByVal psBuilding As Long) As Vector2i
    Dim size As Vector2i, pos As Vector2i
    size = STRUCTURE_size(psBuilding)  ' Adding Vector2i(1, 1) to select the middle of the tile outside the building instead of the corner.
    size.x = size.x + 1
    size.y = size.y + 1
    pos.x = aStruct(psBuilding).pos.x
    pos.y = aStruct(psBuilding).pos.y
    Select Case snapDirection(aStruct(psBuilding).rot.direction)
    Case &H0&: defaultAssemblyPointPos = Vector2i_(pos.x + (TILE_UNITS \ 2) * size.x, pos.y + (TILE_UNITS \ 2) * size.y)
    Case &H4000&: defaultAssemblyPointPos = Vector2i_(pos.x + (TILE_UNITS \ 2) * size.x, pos.y + (TILE_UNITS \ 2) * -size.y)
    Case &H8000&: defaultAssemblyPointPos = Vector2i_(pos.x + (TILE_UNITS \ 2) * -size.x, pos.y + (TILE_UNITS \ 2) * -size.y)
    Case &HC000&: defaultAssemblyPointPos = Vector2i_(pos.x + (TILE_UNITS \ 2) * -size.x, pos.y + (TILE_UNITS \ 2) * size.y)
    End Select
End Function

Private Function setFunctionality(ByVal psBuilding As Long, ByVal functionType As STRUCTURE_TYPE) As Boolean
    Dim pos As Vector2i

    If Not (psBuilding <> 0) Then ASSERT_FAILURE "psBuilding != nullptr", "src/structure.cpp:2078", "setFunctionality", "Invalid pointer": Exit Function
    checkStructure psBuilding, "src/structure.cpp:2079", "setFunctionality", max_check_object_recursion

    Select Case functionType
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY, REF_RESEARCH, REF_POWER_GEN, REF_RESOURCE_EXTRACTOR, _
         REF_REPAIR_FACILITY, REF_REARM_PAD, REF_WALL, REF_GATE
        ' Allocate space for the buildings functionality
        STRUCTURE_calloc_functionality psBuilding
    Case Else
        aStruct(psBuilding).pFunctionality = 0
    End Select

    Select Case functionType
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        aStruct(psBuilding).functionality.factory.psSubject = 0

        ' Default the secondary order - AB 22/04/99
        aStruct(psBuilding).functionality.factory.secondaryOrder = DSS_ARANGE_LONG Or DSS_REPLEV_NEVER Or DSS_ALEV_ALWAYS Or DSS_HALT_GUARD

        ' Create the assembly point for the factory
        If Not createFlagPosition(aStruct(psBuilding).functionality.factory.psAssemblyPoint, aStruct(psBuilding).player) Then
            Exit Function
        End If

        ' Set the assembly point
        pos = defaultAssemblyPointPos(psBuilding)
        setAssemblyPoint aStruct(psBuilding).functionality.factory.psAssemblyPoint, pos.x, pos.y, aStruct(psBuilding).player, True

        ' Add the flag to the list
        addFlagPosition aStruct(psBuilding).functionality.factory.psAssemblyPoint
        Select Case functionType
        Case REF_FACTORY
            setFlagPositionInc aStruct(psBuilding).pFunctionality, aStruct(psBuilding).player, FACTORY_FLAG
        Case REF_CYBORG_FACTORY
            setFlagPositionInc aStruct(psBuilding).pFunctionality, aStruct(psBuilding).player, CYBORG_FLAG
        Case REF_VTOL_FACTORY
            setFlagPositionInc aStruct(psBuilding).pFunctionality, aStruct(psBuilding).player, VTOL_FLAG
        End Select
    Case REF_POWER_GEN, REF_HQ, REF_REARM_PAD
    Case REF_RESOURCE_EXTRACTOR
        ' Make the structure inactive
        aStruct(psBuilding).functionality.resourceExtractor.psPowerGen = 0
    Case REF_REPAIR_FACILITY
        aStruct(psBuilding).functionality.repairFacility.psObj = 0

        ' Create an assembly point for repaired droids
        If Not createFlagPosition(aStruct(psBuilding).functionality.repairFacility.psDeliveryPoint, aStruct(psBuilding).player) Then
            Exit Function
        End If

        ' Set the assembly point
        pos = defaultAssemblyPointPos(psBuilding)
        setAssemblyPoint aStruct(psBuilding).functionality.repairFacility.psDeliveryPoint, pos.x, pos.y, aStruct(psBuilding).player, True

        ' Add the flag (triangular marker on the ground) at the delivery point
        addFlagPosition aStruct(psBuilding).functionality.repairFacility.psDeliveryPoint
        setFlagPositionInc aStruct(psBuilding).pFunctionality, aStruct(psBuilding).player, REPAIR_FLAG
    ' Structure types without a FUNCTIONALITY
    End Select

    setFunctionality = True
End Function

Private Function canPack(ByVal a As Long, ByVal b As Long) As Boolean
    canPack = a + b <= 3  ' Defense can be put next to anything except repair facilities, normal base structures can't be put next to each other, and anything goes next to empty tiles.
End Function

Private Function baseStructureTypePackability(ByVal typ As STRUCTURE_TYPE) As Long
    Select Case typ
    Case REF_DEFENSE, REF_WALL, REF_WALLCORNER, REF_GATE, REF_REARM_PAD, REF_MISSILE_SILO
        baseStructureTypePackability = PACKABILITY_DEFENSE
    Case REF_REPAIR_FACILITY
        baseStructureTypePackability = PACKABILITY_REPAIR
    Case Else
        baseStructureTypePackability = PACKABILITY_NORMAL
    End Select
End Function

Private Function baseObjectPackability(ByVal psObject As Long) As Long
    If psObject = 0 Then
        baseObjectPackability = PACKABILITY_EMPTY
        Exit Function
    End If
    Select Case SIMPLE_OBJECT_type(psObject)
    Case OBJ_STRUCTURE: baseObjectPackability = baseStructureTypePackability(asStructureStats(aStruct(psObject).pStructureType).typ)
    Case OBJ_FEATURE
        If asFeatureStats(aFeature(psObject).psStats).subType = FEAT_OIL_RESOURCE Then baseObjectPackability = PACKABILITY_NORMAL Else baseObjectPackability = PACKABILITY_EMPTY
    Case Else: baseObjectPackability = PACKABILITY_EMPTY
    End Select
End Function

' psStats: a BASE_STATS * handle (DUMMY_STATS_REF0 for findAssemblyPointPosition's local stats)
Public Function validLocation(ByVal psStats As Long, pos As Vector2i, ByVal direction As Long, ByVal player As Long, ByVal bCheckBuildQueue As Boolean) As Boolean
    Dim b As StructureBounds, it As Long, psFlag As Long, flagX As Long, flagY As Long
    Dim psBuilding As Long, psTemplate As Long, i As Long, j As Long, psTile As Long, maxH As Long, minH As Long
    Dim packThis As Long, packObj As Long, object As Long, structure_ As Long, connection As Boolean, psStruct As Long, psFeat As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:4276", "validLocation", "player (%u) >= MAX_PLAYERS", player: Exit Function

    b = getStructureBounds_stats(psStats, pos, direction)

    'make sure we are not too near map edge and not going to go over it
    If b.map.x < scrollMinX + TOO_NEAR_EDGE Or b.map.x + b.size.x > scrollMaxX - TOO_NEAR_EDGE Or _
       b.map.y < scrollMinY + TOO_NEAR_EDGE Or b.map.y + b.size.y > scrollMaxY - TOO_NEAR_EDGE Then
        Exit Function
    End If

    If bCheckBuildQueue Then
        ' cant place on top of a delivery point...
        it = apsFlagPosLists(selectedPlayer).head
        Do While it <> 0
            psFlag = lnObj(it)
            If Not (aFlagPos(psFlag).coords.x <> Not 0) Then ASSERT_FAILURE "psFlag->coords.x != ~0", "src/structure.cpp:4292", "validLocation", "flag has invalid position": Exit Function
            flagX = map_coord(aFlagPos(psFlag).coords.x)
            flagY = map_coord(aFlagPos(psFlag).coords.y)
            If flagX >= b.map.x And flagX < b.map.x + b.size.x And flagY >= b.map.y And flagY < b.map.y + b.size.y Then
                Exit Function
            End If
            it = lnNext(it)
        Loop
    End If

    psBuilding = castStructureStats(psStats)
    psTemplate = castDroidTemplate(psStats)
    If psBuilding <> -1 Then
        For j = 0 To b.size.y - 1
            For i = 0 To b.size.x - 1
                ' Don't allow building structures (allow delivery points, though) outside visible area in single-player with debug mode off. (Why..?)
                If Not bMultiPlayer And Not debugMappingsAllowed() Then
                    If Not TEST_TILE_VISIBLE(player, mapTile(b.map.x + i, b.map.y + j)) Then
                        Exit Function
                    End If
                End If
            Next
        Next

        Select Case asStructureStats(psBuilding).typ
        Case REF_DEMOLISH
        Case NUM_DIFF_BUILDINGS, REF_BRIDGE
            ASSERT_FAILURE "!""invalid structure type""", "src/structure.cpp:4322", "validLocation", "Bad structure type %u", asStructureStats(psBuilding).typ
        Case REF_HQ, REF_FACTORY, REF_LAB, REF_RESEARCH, REF_POWER_GEN, REF_WALL, REF_WALLCORNER, REF_GATE, REF_DEFENSE, _
             REF_REPAIR_FACILITY, REF_COMMAND_CONTROL, REF_CYBORG_FACTORY, REF_VTOL_FACTORY, REF_GENERIC, REF_REARM_PAD, _
             REF_MISSILE_SILO, REF_SAT_UPLINK, REF_LASSAT, REF_FORTRESS
            'need to check each tile the structure will sit on is not water
            For j = 0 To b.size.y - 1
                For i = 0 To b.size.x - 1
                    psTile = mapTile(b.map.x + i, b.map.y + j)
                    If (terrainType(psTile) = TER_WATER) Or (terrainType(psTile) = TER_CLIFFFACE) Then
                        Exit Function
                    End If
                Next
            Next
            'check not within landing zone
            For j = 0 To b.size.y - 1
                For i = 0 To b.size.x - 1
                    If withinLandingZone(b.map.x + i, b.map.y + j) Then
                        Exit Function
                    End If
                Next
            Next

            'walls/defensive structures can be built along any ground
            Select Case asStructureStats(psBuilding).typ
            Case REF_REPAIR_FACILITY, REF_DEFENSE, REF_GATE, REF_WALL
            Case Else
                'cannot build on ground that is too steep
                For j = 0 To b.size.y - 1
                    For i = 0 To b.size.x - 1
                        getTileMaxMin b.map.x + i, b.map.y + j, maxH, minH
                        If maxH - minH > MAX_INCLINE Then
                            Exit Function
                        End If
                    Next
                Next
            End Select
            'don't bother checking if already found a problem
            packThis = baseStructureTypePackability(asStructureStats(psBuilding).typ)

            ' skirmish AIs don't build nondefensives next to anything. (route hack)
            If packThis = PACKABILITY_NORMAL And bMultiPlayer And game.typ = LEVEL_TYPE_SKIRMISH Then
                If Not isHumanPlayer(player) Then packThis = PACKABILITY_REPAIR
            End If

            ' need to check there is one tile between buildings
            For j = -1 To b.size.y
                For i = -1 To b.size.x
                    'skip the actual area the structure will cover
                    If i < 0 Or i >= b.size.x Or j < 0 Or j >= b.size.y Then
                        object = psMapTiles(mapTile(b.map.x + i, b.map.y + j)).psObject
                        structure_ = castStructure(object)
                        If structure_ <> 0 Then
                            If aStruct(structure_).visible(player) = 0 And Not aiCheckAlliances(player, aStruct(structure_).player) Then
                                GoTo continue1  ' Ignore structures we can't see.
                            End If
                        End If

                        packObj = baseObjectPackability(object)

                        If Not canPack(packThis, packObj) Then
                            Exit Function
                        End If
                    End If
continue1:
                Next
            Next
            If (asStructureStats(psBuilding).flags And STRUCTURE_CONNECTED) <> 0 Then
                connection = False
                For j = -1 To b.size.y
                    For i = -1 To b.size.x
                        'skip the actual area the structure will cover
                        If i < 0 Or i >= b.size.x Or j < 0 Or j >= b.size.y Then
                            psStruct = getTileStructure(b.map.x + i, b.map.y + j)
                            If psStruct <> 0 Then
                                If aStruct(psStruct).player = player And aStruct(psStruct).status = SS_BUILT Then
                                    connection = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                Next
                If Not connection Then
                    Exit Function ' needed to be connected to another building
                End If
            End If

            'need to check each tile the structure will sit on
            For j = 0 To b.size.y - 1
                For i = 0 To b.size.x - 1
                    psTile = mapTile(b.map.x + i, b.map.y + j)
                    If TileIsKnownOccupied(psTile, player) Then
                        If TileHasWall(psTile) And (asStructureStats(psBuilding).typ = REF_DEFENSE Or asStructureStats(psBuilding).typ = REF_GATE Or asStructureStats(psBuilding).typ = REF_WALL) Then
                            psStruct = getTileStructure(b.map.x + i, b.map.y + j)
                            If psStruct <> 0 Then
                                If aStruct(psStruct).player <> player Then
                                    Exit Function
                                End If
                            End If
                        Else
                            Exit Function
                        End If
                    End If
                Next
            Next
        Case REF_FACTORY_MODULE
            If Not validLocationModule(pos, player, REF_FACTORY, REF_VTOL_FACTORY) Then Exit Function
        Case REF_RESEARCH_MODULE
            If Not validLocationModule(pos, player, REF_RESEARCH, REF_RESEARCH) Then Exit Function
        Case REF_POWER_MODULE
            If Not validLocationModule(pos, player, REF_POWER_GEN, REF_POWER_GEN) Then Exit Function
        Case REF_RESOURCE_EXTRACTOR
            If Not TileHasFeature(worldTile(pos.x, pos.y)) Then Exit Function
            psFeat = getTileFeature(map_coord(pos.x), map_coord(pos.y))
            If psFeat = 0 Then Exit Function
            If asFeatureStats(aFeature(psFeat).psStats).subType <> FEAT_OIL_RESOURCE Then Exit Function
        End Select
        'if setting up a build queue need to check against future sites as well - AB 4/5/99
        If ctrlShiftDown() And player = selectedPlayer And bCheckBuildQueue Then
            If anyBlueprintTooClose(psBuilding, pos, direction) Then
                Exit Function
            End If
        End If
    ElseIf psTemplate <> 0 Then
        If fpathBlockingTile(b.map.x, b.map.y, asPropulsionStats(DROID_TEMPLATE_getPropulsionStats(apsTemplates(psTemplate))).propulsionType) Then
            Exit Function
        End If
    Else
        ' not positioning a structure or droid, ie positioning a feature
        If fpathBlockingTile(b.map.x, b.map.y, PROPULSION_TYPE_WHEELED) Then
            Exit Function
        End If
    End If

    validLocation = True
End Function

' (VB) validLocation's three module cases: the structure under pos is a built baseType1 or baseType2 of an ally with room for a module
Private Function validLocationModule(pos As Vector2i, ByVal player As Long, ByVal baseType1 As STRUCTURE_TYPE, ByVal baseType2 As STRUCTURE_TYPE) As Boolean
    Dim psStruct As Long, t As STRUCTURE_TYPE
    If TileHasStructure(worldTile(pos.x, pos.y)) Then
        psStruct = getTileStructure(map_coord(pos.x), map_coord(pos.y))
        If psStruct <> 0 Then
            t = asStructureStats(aStruct(psStruct).pStructureType).typ
            If (t = baseType1 Or t = baseType2) And aStruct(psStruct).status = SS_BUILT Then
                If aiCheckAlliances(player, aStruct(psStruct).player) Then
                    validLocationModule = nextModuleToBuild(psStruct, -1) > 0
                End If
            End If
        End If
    End If
End Function

'remove a structure from the map
Private Sub removeStructFromMap(ByVal psStruct As Long)
    Dim b As StructureBounds, i As Long, j As Long, psTile As Long
    auxStructureNonblocking psStruct

    ' set tiles drawing
    b = getStructureBounds_struct(psStruct)
    For j = 0 To b.size.y - 1
        For i = 0 To b.size.x - 1
            psTile = mapTile(b.map.x + i, b.map.y + j)
            If psMapTiles(psTile).psObject = psStruct Then
                psMapTiles(psTile).psObject = 0
                auxClearBlocking b.map.x + i, b.map.y + j, AIR_BLOCKED
            End If
        Next
    Next
End Sub

' remove a structure from a game without any visible effects
' bDestroy = true if the object is to be destroyed
' (for example used to change the type of wall at a location)
Public Function removeStruct(ByVal psDel As Long, ByVal bDestroy As Boolean) As Boolean
    Dim resourceFound As Boolean, psAssemblyPoint As Long, prevResearchState As Long, psOil As Long, p As Long, typ As STRUCTURE_TYPE

    If Not (psDel <> 0) Then ASSERT_FAILURE "psDel != nullptr", "src/structure.cpp:4570", "removeStruct", "Invalid structure pointer": Exit Function

    prevResearchState = intGetResearchState()

    If bDestroy Then
        removeStructFromMap psDel
    End If

    typ = asStructureStats(aStruct(psDel).pStructureType).typ
    If bDestroy Then
        'if the structure is a resource extractor, need to put the resource back in the map
        'ONLY IF ANY POWER LEFT - HACK HACK HACK!!!! OIL POOLS NEED TO KNOW
        'HOW MUCH IS THERE && NOT RES EXTRACTORS
        If typ = REF_RESOURCE_EXTRACTOR Then
            psOil = buildFeature(oilResFeature, aStruct(psDel).pos.x, aStruct(psDel).pos.y, False)
            For p = 0 To MAX_PLAYERS - 1
                aFeature(psOil).seenThisTick(p) = aStruct(psDel).visible(p)
            Next
            resourceFound = True
        End If
    End If

    If typ = REF_RESOURCE_EXTRACTOR Then
        'tell associated Power Gen
        releaseResExtractor psDel
        'tell keybind that this is going away (to prevent dangling pointer in kf_JumpToResourceExtractor)
        keybindInformResourceExtractorRemoved psDel
    End If

    If typ = REF_POWER_GEN Then
        'tell associated Res Extractors
        releasePowerGen psDel
    End If

    'check for a research topic currently under way
    If typ = REF_RESEARCH Then
        If aStruct(psDel).functionality.researchFacility.psSubject <> -1 Then
            'cancel the topic
            cancelResearch psDel, ModeImmediate
        End If
    End If

    If typ = REF_REPAIR_FACILITY Then
        If aStruct(psDel).pFunctionality <> 0 Then
            If aStruct(psDel).functionality.repairFacility.state = RepairState_Repairing Then
                If aStruct(psDel).functionality.repairFacility.psObj <> 0 Then
                    droidRepairStopped castDroid(aStruct(psDel).functionality.repairFacility.psObj), psDel
                End If
            End If
        End If
    End If

    'subtract one from the structLimits list so can build another - don't allow to go less than zero!
    If asStructureStats(aStruct(psDel).pStructureType).curCount(aStruct(psDel).player) <> 0 Then
        asStructureStats(aStruct(psDel).pStructureType).curCount(aStruct(psDel).player) = asStructureStats(aStruct(psDel).pStructureType).curCount(aStruct(psDel).player) - 1
    End If

    'if it is a factory - need to reset the factoryNumFlag
    If STRUCTURE_isFactory(psDel) Then
        'need to initialise the production run as well
        cancelProduction psDel, ModeImmediate

        psAssemblyPoint = aStruct(psDel).functionality.factory.psAssemblyPoint
    ElseIf typ = REF_REPAIR_FACILITY Then
        psAssemblyPoint = aStruct(psDel).functionality.repairFacility.psDeliveryPoint
    End If

    If psAssemblyPoint <> 0 Then
        If aFlagPos(psAssemblyPoint).factoryInc < factoryNumFlag_size(aStruct(psDel).player, aFlagPos(psAssemblyPoint).factoryType) Then
            factoryNumFlag(aStruct(psDel).player, aFlagPos(psAssemblyPoint).factoryType, aFlagPos(psAssemblyPoint).factoryInc) = False
        End If

        'need to cancel the repositioning of the DP if selectedPlayer and currently moving
        If aStruct(psDel).player = selectedPlayer And aFlagPos(psAssemblyPoint).selected Then
            cancelDeliveryRepos
        End If
    End If

    If bDestroy Then
        debug_ 4664, LOG_DEATH, "removeStruct", "Killing off %s id %d (%p)", objInfo(psDel), aStruct(psDel).id, psDel
        killStruct psDel
    End If

    If aStruct(psDel).player = selectedPlayer Then
        intRefreshScreen
    End If

    delPowerRequest psDel

    intNotifyResearchButton prevResearchState

    removeStruct = resourceFound
End Function

Private Sub findAssemblyPointPosition(pX As Long, pY As Long, ByVal player As Long)
    Dim passes As Long, i As Long, j As Long, startX As Long, endX As Long, startY As Long, endY As Long

    'set up a dummy stat pointer (STRUCTURE_STATS sStats with ref 0, 1 x 1: the handle DUMMY_STATS_REF0)

    ' Initial box dimensions and set iteration count to zero
    startX = pX: endX = pX: startY = pY: endY = pY
    passes = 0

    'if the value passed in is not a valid location - find one!
    If Not validLocation(DUMMY_STATS_REF0, Vector2i_(world_coord(pX), world_coord(pY)), 0, player, False) Then
        ' Keep going until we get a tile or we exceed distance
        Do While passes < LOOK_FOR_EMPTY_TILE
            ' Process whole box
            For i = startX To endX
                For j = startY To endY
                    ' Test only perimeter as internal tested previous iteration
                    If i = startX Or i = endX Or j = startY Or j = endY Then
                        ' Good enough?
                        If validLocation(DUMMY_STATS_REF0, Vector2i_(world_coord(i), world_coord(j)), 0, player, False) Then
                            ' Set exit conditions and get out NOW
                            pX = i
                            pY = j
                            'jump out of the loop
                            Exit Sub
                        End If
                    End If
                Next
            Next
            ' Expand the box out in all directions - off map handled by validLocation()
            startX = startX - 1: startY = startY - 1: endX = endX + 1: endY = endY + 1: passes = passes + 1
        Loop
    Else
        'the first location was valid
        Exit Sub
    End If
    ' If we got this far, then we failed - passed in values will be unchanged
    ASSERT_FAILURE "!""unable to find a valid location""", "src/structure.cpp:4976", "findAssemblyPointPosition", "unable to find a valid location!"
End Sub

'sets the point new droids go to - x/y in world coords for a Factory
'bCheck is set to true for initial placement of the Assembly Point
Public Sub setAssemblyPoint(ByVal psAssemblyPoint As Long, ByVal x As Long, ByVal y As Long, ByVal player As Long, ByVal bCheck As Boolean)
    If Not (psAssemblyPoint <> 0) Then ASSERT_FAILURE "psAssemblyPoint != nullptr", "src/structure.cpp:4985", "setAssemblyPoint", "invalid AssemblyPoint pointer": Exit Sub

    'check its valid
    x = map_coord(x)
    y = map_coord(y)
    If bCheck Then
        findAssemblyPointPosition x, y, player
    End If
    'add half a tile so the centre is in the middle of the tile
    x = world_coord(x) + TILE_UNITS \ 2
    y = world_coord(y) + TILE_UNITS \ 2

    aFlagPos(psAssemblyPoint).coords.x = x
    aFlagPos(psAssemblyPoint).coords.y = y

    ' Deliv Point sits at the height of the tile it's centre is on + arbitrary amount!
    aFlagPos(psAssemblyPoint).coords.z = map_Height(x, y) + ASSEMBLY_POINT_Z_PADDING
End Sub

'sets the factory Inc for the Assembly Point
' (pFunctionality: the handle of the structure that holds it)
Private Sub setFlagPositionInc(ByVal pFunctionality As Long, ByVal player As Long, ByVal factoryType As Long)
    Dim inc As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:5009", "setFlagPositionInc", "invalid player number": Exit Sub

    'find the first vacant slot
    inc = 0
    Do While inc < factoryNumFlag_size(player, factoryType)
        If Not factoryNumFlag(player, factoryType, inc) Then Exit Do
        inc = inc + 1
    Loop
    If inc = factoryNumFlag_size(player, factoryType) Then
        ' first time init for this factory flag slot, set it to false
        factoryNumFlag(player, factoryType, inc) = False
        factoryNumFlag_size(player, factoryType) = inc + 1
    End If

    If factoryType = REPAIR_FLAG Then
        ' this is a special case, there are no flag numbers for this "factory"
        aFlagPos(aStruct(pFunctionality).functionality.repairFacility.psDeliveryPoint).factoryInc = 0
        aFlagPos(aStruct(pFunctionality).functionality.repairFacility.psDeliveryPoint).factoryType = factoryType
        ' factoryNumFlag[player][factoryType][inc] = true;
    Else
        aFlagPos(aStruct(pFunctionality).functionality.factory.psAssemblyPoint).factoryInc = inc And &HFF&
        aFlagPos(aStruct(pFunctionality).functionality.factory.psAssemblyPoint).factoryType = factoryType
        factoryNumFlag(player, factoryType, inc) = True
    End If
End Sub

'sets the flag to indicate a SatUplink Exists - so draw everything!
Public Sub setSatUplinkExists(ByVal state As Boolean, ByVal player As Long)
    satUplinkExists(player) = -CLng(state)
    If state Then
        satuplinkbits = satuplinkbits Or BIT_(player)
    Else
        satuplinkbits = satuplinkbits And Not BIT_(player)
    End If
End Sub

'returns the status of the flag
Public Function getSatUplinkExists(ByVal player As Long) As Boolean
    getSatUplinkExists = satUplinkExists(player) <> 0
End Function

'sets the flag to indicate a Las Sat Exists - ONLY EVER WANT ONE
Public Sub setLasSatExists(ByVal state As Boolean, ByVal player As Long)
    lasSatExists(player) = -CLng(state)
End Sub

Public Sub checkForResExtractors(ByVal psBuilding As Long)
    Dim prio(0 To NUM_POWER_MODULES) As Long, ders(0 To NUM_POWER_MODULES) As Long, n As Long
    Dim it As Long, currExtractor As Long, priority As Long, d As Long, k As Long, i As Long, derrick As Long

    If Not (asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_POWER_GEN) Then ASSERT_FAILURE "psBuilding->pStructureType->type == REF_POWER_GEN", "src/structure.cpp:5179", "checkForResExtractors", "invalid structure type": Exit Sub

    ' Find derricks, sorted by unused first, then ones attached to power generators without modules.
    ' (Derricks: std::vector<std::pair<int, STRUCTURE *>>, at most NUM_POWER_MODULES kept)
    it = apsExtractorLists(aStruct(psBuilding).player).head
    Do While it <> 0
        currExtractor = lnObj(it)

        If aStruct(currExtractor).status <> SS_BUILT Then
            GoTo continue_  ' Derrick not complete.
        End If
        If aStruct(currExtractor).functionality.resourceExtractor.psPowerGen <> 0 Then
            priority = aStruct(aStruct(currExtractor).functionality.resourceExtractor.psPowerGen).capacity
        Else
            priority = -1
        End If
        d = 0
        Do While d < n
            If Not (prio(d) <= priority) Then Exit Do
            d = d + 1
        Loop
        ' derricks.insert(d, Derrick(priority, currExtractor))
        For k = n To d + 1 Step -1
            prio(k) = prio(k - 1)
            ders(k) = ders(k - 1)
        Next
        prio(d) = priority
        ders(d) = currExtractor
        n = n + 1
        If n > NUM_POWER_MODULES Then n = NUM_POWER_MODULES  ' No point remembering more derricks than this.
continue_:
        it = lnNext(it)
    Loop

    ' Attach derricks.
    d = 0
    For i = 0 To NUM_POWER_MODULES - 1
        If aStruct(psBuilding).functionality.powerGenerator.apResExtractors(i) <> 0 Then
            GoTo continue2  ' Slot full.
        End If

        priority = aStruct(psBuilding).capacity
        If d = n Then GoTo continue2  ' No more derricks to transfer to this power generator.
        If prio(d) >= priority Then GoTo continue2

        derrick = ders(d)
        If aStruct(derrick).functionality.resourceExtractor.psPowerGen <> 0 Then
            informPowerGen derrick  ' Remove the derrick from the previous power generator.
        End If
        ' Assign the derrick to the power generator.
        aStruct(psBuilding).functionality.powerGenerator.apResExtractors(i) = derrick
        aStruct(derrick).functionality.resourceExtractor.psPowerGen = psBuilding

        d = d + 1
continue2:
    Next
End Sub

Public Function countPlayerUnusedDerricks() As Long
    Dim total As Long, it As Long, psStruct As Long

    If Not ULT(selectedPlayer, MAX_PLAYERS) Then Exit Function

    it = apsExtractorLists(selectedPlayer).head
    Do While it <> 0
        psStruct = lnObj(it)
        If aStruct(psStruct).status = SS_BUILT And asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then
            If aStruct(psStruct).functionality.resourceExtractor.psPowerGen = 0 Then
                total = total + 1
            End If
        End If
        it = lnNext(it)
    Loop

    countPlayerUnusedDerricks = total And &HFFFF&
End Function

'Looks through the list of structures to see if there are any Power Gens
'with available slots for the new Res Ext
Public Sub checkForPowerGen(ByVal psBuilding As Long)
    Dim bestPowerGen As Long, bestSlot As Long, it As Long, psCurr As Long, i As Long

    If Not (asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESOURCE_EXTRACTOR) Then ASSERT_FAILURE "psBuilding->pStructureType->type == REF_RESOURCE_EXTRACTOR", "src/structure.cpp:5258", "checkForPowerGen", "invalid structure type": Exit Sub

    If aStruct(psBuilding).functionality.resourceExtractor.psPowerGen <> 0 Then
        Exit Sub
    End If

    ' Find a power generator, if possible with a power module.
    bestPowerGen = 0
    bestSlot = 0
    it = apsStructLists(aStruct(psBuilding).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If asStructureStats(aStruct(psCurr).pStructureType).typ = REF_POWER_GEN And aStruct(psCurr).status = SS_BUILT Then
            If bestPowerGen <> 0 Then
                If aStruct(bestPowerGen).capacity >= aStruct(psCurr).capacity Then
                    GoTo continue_  ' Power generator not better.
                End If
            End If

            For i = 0 To NUM_POWER_MODULES - 1
                If aStruct(psCurr).functionality.powerGenerator.apResExtractors(i) = 0 Then
                    bestPowerGen = psCurr
                    bestSlot = i
                    Exit For
                End If
            Next
        End If
continue_:
        it = lnNext(it)
    Loop

    If bestPowerGen <> 0 Then
        ' Attach the derrick to the power generator.
        aStruct(bestPowerGen).functionality.powerGenerator.apResExtractors(bestSlot) = psBuilding
        aStruct(psBuilding).functionality.resourceExtractor.psPowerGen = bestPowerGen
    End If
End Sub

'initialise the slot the Resource Extractor filled in the owning Power Gen
Private Sub informPowerGen(ByVal psStruct As Long)
    Dim i As Long, psPowerGen As Long

    If asStructureStats(aStruct(psStruct).pStructureType).typ <> REF_RESOURCE_EXTRACTOR Then
        ASSERT_FAILURE "!""invalid structure type""", "src/structure.cpp:5309", "informPowerGen", "invalid structure type"
        Exit Sub
    End If

    'get the owning power generator
    psPowerGen = aStruct(psStruct).functionality.resourceExtractor.psPowerGen
    If psPowerGen <> 0 Then
        For i = 0 To NUM_POWER_MODULES - 1
            If aStruct(psPowerGen).functionality.powerGenerator.apResExtractors(i) = psStruct Then
                'initialise the 'slot'
                aStruct(psPowerGen).functionality.powerGenerator.apResExtractors(i) = 0
                Exit For
            End If
        Next
    End If
End Sub

'called when a Res extractor is destroyed or runs out of power or is disconnected
'adjusts the owning Power Gen so that it can link to a different Res Extractor if one
'is available
Public Sub releaseResExtractor(ByVal psRelease As Long)
    Dim it As Long, psCurr As Long

    If asStructureStats(aStruct(psRelease).pStructureType).typ <> REF_RESOURCE_EXTRACTOR Then
        ASSERT_FAILURE "!""invalid structure type""", "src/structure.cpp:5337", "releaseResExtractor", "Invalid structure type"
        Exit Sub
    End If

    'tell associated Power Gen
    If aStruct(psRelease).functionality.resourceExtractor.psPowerGen <> 0 Then
        informPowerGen psRelease
    End If

    aStruct(psRelease).functionality.resourceExtractor.psPowerGen = 0

    'there may be spare resource extractors
    it = apsExtractorLists(aStruct(psRelease).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        'check not connected and power left and built!
        If psCurr <> psRelease And aStruct(psCurr).functionality.resourceExtractor.psPowerGen = 0 And aStruct(psCurr).status = SS_BUILT Then
            checkForPowerGen psCurr
        End If
        it = lnNext(it)
    Loop
End Sub

'called when a Power Gen is destroyed or is disconnected
'adjusts the associated Res Extractors so that they can link to different Power
'Gens if any are available
Public Sub releasePowerGen(ByVal psRelease As Long)
    Dim i As Long, it As Long, psCurr As Long

    If asStructureStats(aStruct(psRelease).pStructureType).typ <> REF_POWER_GEN Then
        ASSERT_FAILURE "!""invalid structure type""", "src/structure.cpp:5371", "releasePowerGen", "Invalid structure type"
        Exit Sub
    End If

    'go through list of res extractors, setting them to inactive
    For i = 0 To NUM_POWER_MODULES - 1
        If aStruct(psRelease).functionality.powerGenerator.apResExtractors(i) <> 0 Then
            aStruct(aStruct(psRelease).functionality.powerGenerator.apResExtractors(i)).functionality.resourceExtractor.psPowerGen = 0
            aStruct(psRelease).functionality.powerGenerator.apResExtractors(i) = 0
        End If
    Next
    'may have a power gen with spare capacity
    it = apsStructLists(aStruct(psRelease).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If asStructureStats(aStruct(psCurr).pStructureType).typ = REF_POWER_GEN And _
           psCurr <> psRelease And aStruct(psCurr).status = SS_BUILT Then
            checkForResExtractors psCurr
        End If
        it = lnNext(it)
    Loop
End Sub

'this is called whenever a structure has finished building
Public Sub buildingComplete(ByVal psBuilding As Long)
    Dim prevState As Long, imdIndex As Long

    checkStructure psBuilding, "src/structure.cpp:5400", "buildingComplete", max_check_object_recursion

    prevState = 0
    If asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESEARCH Then
        prevState = intGetResearchState()
    End If

    aStruct(psBuilding).currentBuildPts = structureBuildPointsToCompletion(psBuilding)
    aStruct(psBuilding).status = SS_BUILT

    visTilesUpdate psBuilding

    If aStruct(psBuilding).prebuiltImd <> 0 Then
        ' We finished building a module, now use the combined IMD.
        imdIndex = MIN_(numStructureModules(psBuilding) * 2, asStructureStats(aStruct(psBuilding).pStructureType).pIMD_size - 1) ' *2 because even-numbered IMDs are structures, odd-numbered IMDs are just the modules.
        aStruct(psBuilding).prebuiltImd = 0
        aStruct(psBuilding).sDisplay.imd = asStructureStats(aStruct(psBuilding).pStructureType).pIMD(imdIndex)
    End If

    Select Case asStructureStats(aStruct(psBuilding).pStructureType).typ
    Case REF_POWER_GEN
        checkForResExtractors psBuilding
    Case REF_RESOURCE_EXTRACTOR
        checkForPowerGen psBuilding
    Case REF_RESEARCH
        'this deals with research facilities that are upgraded whilst mid-research
        releaseResearch psBuilding, ModeImmediate
        intNotifyResearchButton prevState
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        'this deals with factories that are upgraded whilst mid-production
        releaseProduction psBuilding, ModeImmediate
    Case REF_SAT_UPLINK
        revealAll aStruct(psBuilding).player
    Case REF_GATE
        auxStructureNonblocking psBuilding  ' Clear outdated flags.
        auxStructureClosedGate psBuilding  ' Don't block for the sake of allied pathfinding.
    Case Else
        'do nothing
    End Select
End Sub

'for a given structure, return a pointer to its module stat
Public Function getModuleStat(ByVal psStruct As Long) As Long
    If Not (psStruct <> 0) Then ASSERT_FAILURE "psStruct != nullptr", "src/structure.cpp:5458", "getModuleStat", "Invalid structure pointer": getModuleStat = -1: Exit Function

    Select Case asStructureStats(aStruct(psStruct).pStructureType).typ
    Case REF_POWER_GEN
        getModuleStat = powerModuleStat
    Case REF_FACTORY, REF_VTOL_FACTORY
        getModuleStat = factoryModuleStat
    Case REF_RESEARCH
        getModuleStat = researchModuleStat
    Case Else
        'no other structures can have modules attached
        getModuleStat = -1
    End Select
End Function

Public Function structureBodyBuilt(ByVal psStructure As Long) As Long
    Dim maxBody As Long, unbuiltBody As Long, deltaBody As Long, f As Single

    maxBody = STRUCTURE_structureBody(psStructure)

    If aStruct(psStructure).status = SS_BEING_BUILT Then
        ' Calculate the body points the structure would have, if not damaged.
        unbuiltBody = FromU32(Int((U32(maxBody) + 9) / 10))  ' See droidStartBuild() in droid.cpp.
        ' static_cast<unsigned>(maxBody * 9 * structureCompletionProgress(*psStructure) / 10)
        f = fdiv(fmul(F32(U32(MUL32(maxBody, 9))), structureCompletionProgress(psStructure)), 10!)
        deltaBody = FromU32(Fix(f))  ' See structureBuild() in structure.cpp.
        maxBody = unbuiltBody + deltaBody
    End If

    structureBodyBuilt = maxBody
End Function

'Access functions for the upgradeable stats of a structure
Public Function STRUCTURE_structureBody(ByVal this As Long) As Long
    STRUCTURE_structureBody = asStructureStats(aStruct(this).pStructureType).upgrade(aStruct(this).player).hitpoints
End Function

Public Function structureResistance(ByVal psStats As Long, ByVal player As Long) As Long
    structureResistance = asStructureStats(psStats).upgrade(player).resistance
End Function

' structure.h
Public Function isLasSat(ByVal pStructureType As Long) As Boolean
    If Not (pStructureType <> -1) Then ASSERT_FAILURE "pStructureType != nullptr", "src/structure.h:166", "isLasSat", "LasSat is invalid?": Exit Function

    If asStructureStats(pStructureType).psWeapStat(0) <> -1 Then
        isLasSat = asWeaponStats(asStructureStats(pStructureType).psWeapStat(0)).weaponSubClass = WSC_LAS_SAT
    End If
End Function

' structure.h's setStatusPendingHold and setStatusPendingRelease, for a FACTORY
Private Sub setStatusPendingHold_factory(func_ As FACTORY)
    If func_.psSubjectPending = 0 Then
        func_.psSubjectPending = func_.psSubject
    End If
    func_.statusPending = FACTORY_HOLD_PENDING
    func_.pendingCount = func_.pendingCount + 1
End Sub

Private Sub setStatusPendingRelease_factory(func_ As FACTORY)
    If func_.psSubjectPending = 0 And func_.statusPending <> FACTORY_CANCEL_PENDING Then
        func_.psSubjectPending = func_.psSubject
    End If
    If func_.psSubjectPending <> 0 Then
        func_.statusPending = FACTORY_START_PENDING
    End If
    func_.pendingCount = func_.pendingCount + 1
End Sub

'set a factory's production run to hold
Public Sub holdProduction(ByVal psBuilding As Long, ByVal mode As QUEUE_MODE)
    If psBuilding = 0 Then ASSERT_FAILURE "psBuilding && psBuilding->isFactory()", "src/structure.cpp:6237", "holdProduction", "structure not a factory": Exit Sub
    If Not STRUCTURE_isFactory(psBuilding) Then ASSERT_FAILURE "psBuilding && psBuilding->isFactory()", "src/structure.cpp:6237", "holdProduction", "structure not a factory": Exit Sub

    If mode = ModeQueue Then
        sendStructureInfo psBuilding, STRUCTUREINFO_HOLDPRODUCTION, 0
        setStatusPendingHold_factory aStruct(psBuilding).functionality.factory

        Exit Sub
    End If

    If aStruct(psBuilding).functionality.factory.psSubject <> 0 Then
        'set the time the factory was put on hold
        aStruct(psBuilding).functionality.factory.timeStartHold = gameTime
        'play audio to indicate on hold
        If aStruct(psBuilding).player = selectedPlayer Then
            audio_PlayTrack ID_SOUND_WINDOWCLOSE
        End If
    End If

    delPowerRequest psBuilding
End Sub

'release a factory's production run from hold
Public Sub releaseProduction(ByVal psBuilding As Long, ByVal mode As QUEUE_MODE)
    If psBuilding = 0 Then ASSERT_FAILURE "psBuilding && psBuilding->isFactory()", "src/structure.cpp:6266", "releaseProduction", "structure not a factory": Exit Sub
    If Not STRUCTURE_isFactory(psBuilding) Then ASSERT_FAILURE "psBuilding && psBuilding->isFactory()", "src/structure.cpp:6266", "releaseProduction", "structure not a factory": Exit Sub

    If mode = ModeQueue Then
        sendStructureInfo psBuilding, STRUCTUREINFO_RELEASEPRODUCTION, 0
        setStatusPendingRelease_factory aStruct(psBuilding).functionality.factory

        Exit Sub
    End If

    If aStruct(psBuilding).functionality.factory.psSubject <> 0 And aStruct(psBuilding).functionality.factory.timeStartHold <> 0 Then
        'adjust the start time for the current subject
        If aStruct(psBuilding).functionality.factory.timeStarted <> ACTION_START_TIME Then
            aStruct(psBuilding).functionality.factory.timeStarted = aStruct(psBuilding).functionality.factory.timeStarted + (gameTime - aStruct(psBuilding).functionality.factory.timeStartHold)
        End If
        aStruct(psBuilding).functionality.factory.timeStartHold = 0
    End If
End Sub
' ---- B7: building, demolishing and repairing by droids; sensors; rearm pads

' Add buildPoints to the structures currentBuildPts, due to construction work by the droid
' Also can deconstruct (demolish) a building if passed negative buildpoints
Public Sub structureBuild(ByVal psStruct As Long, ByVal psDroid As Long, ByVal buildPoints As Long, Optional ByVal buildRate As Long = 1)
    Dim checkResearchButton As Boolean, prevResearchState As Long, player As Long, it As Long, psCurr As Long
    Dim haveEnoughPower As Boolean, newBuildPoints As Long, deltaBody As Long, prevStatus As STRUCT_STATES
    Dim psIter As Long, ok As Boolean, t As DroidOrderType

    checkResearchButton = aStruct(psStruct).status = SS_BUILT  ' We probably just started demolishing, if this is true.
    prevResearchState = 0
    If checkResearchButton Then
        prevResearchState = intGetResearchState()
    End If

    ok = psDroid <> 0
    If ok Then ok = Not aiCheckAlliances(aStruct(psStruct).player, aDroid(psDroid).player)
    If ok Then
        ' Enemy structure
        Exit Sub
    ElseIf asStructureStats(aStruct(psStruct).pStructureType).typ <> REF_FACTORY_MODULE Then
        For player = 0 To MAX_PLAYERS - 1
            it = apsDroidLists(player).head
            Do While it <> 0
                psCurr = lnObj(it)
                ' An enemy droid is blocking it
                If orderStateObj(psCurr, DORDER_BUILD) = psStruct Then
                    If Not aiCheckAlliances(aStruct(psStruct).player, aDroid(psCurr).player) Then
                        Exit Sub
                    End If
                End If
                it = lnNext(it)
            Loop
        Next
    End If
    aStruct(psStruct).buildRate = aStruct(psStruct).buildRate + buildRate  ' buildRate = buildPoints/GAME_UPDATES_PER_SEC, but might be rounded up or down each tick, so can't use buildPoints to get a stable number.
    If aStruct(psStruct).currentBuildPts = 0 And buildPoints > 0 Then   ' (uint32_t) currentBuildPts <= 0
        ' Just starting to build structure, need power for it.
        haveEnoughPower = requestPowerFor(psStruct, structPowerToBuildOrAddNextModule(psStruct))
        If Not haveEnoughPower Then
            buildPoints = 0  ' No power to build.
        End If
    End If

    newBuildPoints = FromU32(CDbl(aStruct(psStruct).currentBuildPts) + buildPoints)
    If Not (newBuildPoints <= 1 + 3 * structureBuildPointsToCompletion(psStruct)) Then ASSERT_FAILURE "newBuildPoints <= 1 + 3 * (int)structureBuildPointsToCompletion(*psStruct)", "src/structure.cpp:953", "structureBuild", "unsigned int underflow?"
    CLIP newBuildPoints, 0, structureBuildPointsToCompletion(psStruct)

    If aStruct(psStruct).currentBuildPts <> 0 And newBuildPoints <= 0 Then
        ' Demolished structure, return some power.
        addPower aStruct(psStruct).player, structureTotalReturn(psStruct)
    End If

    deltaBody = quantiseFraction(MUL32(9, STRUCTURE_structureBody(psStruct)), MUL32(10, structureBuildPointsToCompletion(psStruct)), newBuildPoints, aStruct(psStruct).currentBuildPts)
    aStruct(psStruct).currentBuildPts = newBuildPoints
    aStruct(psStruct).body = MAX_(FromU32(CDbl(aStruct(psStruct).body) + deltaBody), 1)

    'check if structure is built
    If buildPoints > 0 And Not ULT(aStruct(psStruct).currentBuildPts, structureBuildPointsToCompletion(psStruct)) Then
        buildingComplete psStruct

        'only play the sound if selected player
        If psDroid <> 0 Then
            If aStruct(psStruct).player = selectedPlayer _
               And (aDroid(psDroid).order.typ <> DORDER_LINEBUILD _
                    Or (map_coord(aDroid(psDroid).order.pos.x) = map_coord(aDroid(psDroid).order.pos2.x) And map_coord(aDroid(psDroid).order.pos.y) = map_coord(aDroid(psDroid).order.pos2.y))) Then
                audio_QueueTrackPos ID_SOUND_STRUCTURE_COMPLETED, _
                                    aStruct(psStruct).pos.x, aStruct(psStruct).pos.y, aStruct(psStruct).pos.z
                intRefreshScreen        ' update any open interface bars.
            End If
        End If

        ' must reset here before the callback, droid must have DACTION_NONE
        '     in order to be able to start a new built task, doubled in actionUpdateDroid()
        If psDroid <> 0 Then
            ' Clear all orders for helping hands. Needed for AI script which runs next frame.
            it = apsDroidLists(aDroid(psDroid).player).head
            Do While it <> 0
                psIter = lnObj(it)
                t = aDroid(psIter).order.typ
                ok = (t = DORDER_BUILD Or t = DORDER_HELPBUILD Or t = DORDER_LINEBUILD) And aDroid(psIter).order.psObj = psStruct
                If ok And t = DORDER_LINEBUILD Then
                    ok = map_coord(aDroid(psIter).order.pos.x) = map_coord(aDroid(psIter).order.pos2.x) And map_coord(aDroid(psIter).order.pos.y) = map_coord(aDroid(psIter).order.pos2.y)
                End If
                If ok Then
                    objTrace aDroid(psIter).id, "structureBuild", "Construction order %s complete (%d, %d -> %d, %d)", getDroidOrderName(aDroid(psDroid).order.typ), _
                             aDroid(psIter).order.pos2.x, aDroid(psIter).order.pos.y, aDroid(psIter).order.pos2.x, aDroid(psIter).order.pos2.y
                    aDroid(psIter).action = DACTION_NONE
                    aDroid(psIter).order = DroidOrder_(DORDER_NONE)
                    setDroidActionTarget psIter, 0, 0
                End If
                it = lnNext(it)
            Loop

            audio_StopObjTrack psDroid, ID_SOUND_CONSTRUCTION_LOOP
        End If
        triggerEventStructBuilt psStruct, psDroid
        checkPlayerBuiltHQ psStruct
    Else
        prevStatus = aStruct(psStruct).status
        aStruct(psStruct).status = SS_BEING_BUILT
        If prevStatus = SS_BUILT Then
            ' Starting to demolish.
            triggerEventStructDemolish psStruct, psDroid
            If aStruct(psStruct).player = selectedPlayer Then
                intRefreshScreen
            End If

            Select Case asStructureStats(aStruct(psStruct).pStructureType).typ
            Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
                If aStruct(psStruct).pFunctionality <> 0 Then
                    If aStruct(psStruct).functionality.factory.psCommander <> 0 Then
                        'remove the commander from the factory
                        syncDebugDroid_ "structureBuild", aStruct(psStruct).functionality.factory.psCommander, Asc("-")
                        assignFactoryCommandDroid psStruct, 0
                    End If
                End If
            Case REF_POWER_GEN
                releasePowerGen psStruct
            Case REF_RESOURCE_EXTRACTOR
                releaseResExtractor psStruct
            Case REF_REPAIR_FACILITY
                If aStruct(psStruct).pFunctionality <> 0 Then
                    If aStruct(psStruct).functionality.repairFacility.psObj <> 0 Then
                        If aStruct(psStruct).functionality.repairFacility.state = RepairState_Repairing Then
                            droidRepairStopped castDroid(aStruct(psStruct).functionality.repairFacility.psObj), psStruct
                        End If
                        aStruct(psStruct).functionality.repairFacility.psObj = 0
                        aStruct(psStruct).functionality.repairFacility.state = RepairState_Idle
                    End If
                End If
            Case REF_REARM_PAD
                If aStruct(psStruct).pFunctionality <> 0 Then
                    If aStruct(psStruct).functionality.rearmPad.psObj <> 0 Then
                        ' Possible TODO: Need to do anything with the droid? (order it to find a new place to rearm?)
                        aStruct(psStruct).functionality.rearmPad.psObj = 0
                    End If
                End If
            Case Else
            End Select
        End If
    End If
    If buildPoints < 0 And aStruct(psStruct).currentBuildPts = 0 Then
        triggerEvent TRIGGER_OBJECT_RECYCLED, psStruct
        removeStruct psStruct, True
    End If

    If checkResearchButton Then
        intNotifyResearchButton prevResearchState
    End If
End Sub

' Power returned on demolish, which is half the power taken to build the structure and any modules
Private Function structureTotalReturn(ByVal psStruct As Long) As Long
    Dim power As Long, moduleStats As Long
    power = asStructureStats(aStruct(psStruct).pStructureType).powerToBuild

    moduleStats = getModuleStat(psStruct)

    If moduleStats <> -1 Then
        power = power + aStruct(psStruct).capacity * asStructureStats(moduleStats).powerToBuild
    End If

    structureTotalReturn = power \ 2
End Function

Public Sub structureDemolish(ByVal psStruct As Long, ByVal psDroid As Long, ByVal buildPoints As Long)
    structureBuild psStruct, psDroid, -buildPoints
End Sub

Public Sub structureRepair(ByVal psStruct As Long, ByVal psDroid As Long, ByVal buildRate As Long)
    Dim repairAmount As Long, v As Long
    repairAmount = gameTimeAdjustedAverage(MUL32(buildRate, STRUCTURE_structureBody(psStruct)), asStructureStats(aStruct(psStruct).pStructureType).buildPoints)
    '   (droid construction power * current max hitpoints [incl. upgrades])
    '       / construction power that was necessary to build structure in the first place
    '
    ' => to repair a building from 1HP to full health takes as much time as building it.
    ' => if buildPoints = 1 and structureBody < buildPoints, repairAmount might get truncated to zero.
    '   This happens with expensive, but weak buildings like mortar pits. In this case, do nothing
    '   and notify the caller (read: droid) of your idleness by returning false.
    ' clip<UDWORD>(body + repairAmount, 0, structureBody())
    v = FromU32(CDbl(aStruct(psStruct).body) + repairAmount)
    If ULT(STRUCTURE_structureBody(psStruct), v) Then v = STRUCTURE_structureBody(psStruct)
    aStruct(psStruct).body = v
End Sub

' int gateCurrentOpenHeight(const STRUCTURE *psStructure, uint32_t time, int minimumStub)
' Returns how far open the gate is, or 0 if the structure is not a gate.
Public Function gateCurrentOpenHeight(ByVal psStructure As Long, ByVal time_ As Long, ByVal minimumStub As Long) As Long
    Dim psStructureStats As Long, height As Long, openHeight As Long
    psStructureStats = aStruct(psStructure).pStructureType
    If asStructureStats(psStructureStats).typ = REF_GATE Then
        height = imdMaxY(aStruct(psStructure).sDisplay.imd)
        Select Case aStruct(psStructure).state
        Case SAS_OPEN
            openHeight = height
        Case SAS_OPENING
            openHeight = MUL32(height, MAX_(FromU32(CDbl(time_) + GAME_TICKS_PER_UPDATE - aStruct(psStructure).lastStateTime), 0)) \ SAS_OPEN_SPEED
        Case SAS_CLOSING
            openHeight = height - MUL32(height, MAX_(FromU32(CDbl(time_) - aStruct(psStructure).lastStateTime), 0)) \ SAS_OPEN_SPEED
        Case Else
            gateCurrentOpenHeight = 0
            Exit Function
        End Select
        gateCurrentOpenHeight = MAX_(MIN_(openHeight, height - minimumStub), 0)
        Exit Function
    End If
    gateCurrentOpenHeight = 0
End Function

Public Function structGetDemolishStat() As Long
    If Not (g_psStatDestroyStruct <> -1) Then ASSERT_FAILURE "g_psStatDestroyStruct != nullptr", "src/structure.cpp:5038", "structGetDemolishStat", "Demolish stat not initialised": structGetDemolishStat = -1: Exit Function
    structGetDemolishStat = g_psStatDestroyStruct
End Function

' calculate muzzle base location in 3d world
Public Function calcStructureMuzzleBaseLocation(ByVal psStructure As Long, muzzle As Vector3i, ByVal weapon_slot As Long) As Boolean
    Dim psShape As Long, barrel As Vector3i, af As Affine3F, cx As Long, cy As Long, cz As Long, r As Vector3i, ok As Boolean

    psShape = asStructureStats(aStruct(psStructure).pStructureType).pIMD(0)

    checkStructure psStructure, "src/structure.cpp:5077", "calcStructureMuzzleBaseLocation", max_check_object_recursion

    ok = psShape <> 0
    If ok Then ok = iIMDBaseShape_connectors_size(psShape) <> 0
    If ok Then
        barrel.x = 0: barrel.y = 0: barrel.z = 0

        Affine3F_construct af

        Affine3F_Trans af, aStruct(psStructure).pos.x, -aStruct(psStructure).pos.z, aStruct(psStructure).pos.y

        'matrix = the center of droid
        Affine3F_RotY af, aStruct(psStructure).rot.direction
        Affine3F_RotX af, aStruct(psStructure).rot.pitch
        Affine3F_RotZ af, -aStruct(psStructure).rot.roll
        iIMDBaseShape_connector psShape, weapon_slot, cx, cy, cz
        Affine3F_Trans af, cx, -cz, -cy 'note y and z flipped

        r = Affine3F_mul_vec(af, barrel)
        muzzle.x = r.x  ' xzy()
        muzzle.y = r.z
        muzzle.z = r.y
        muzzle.z = -muzzle.z
    Else
        muzzle.x = aStruct(psStructure).pos.x
        muzzle.y = aStruct(psStructure).pos.y
        muzzle.z = aStruct(psStructure).pos.z + imdMaxY(aStruct(psStructure).sDisplay.imd)
    End If

    calcStructureMuzzleBaseLocation = True
End Function

' EW works differently in multiplayer mode compared with single player.
Public Function validStructResistance(ByVal psStruct As Long) As Boolean
    Dim bTarget As Boolean

    bTarget = False

    If Not (psStruct <> 0) Then ASSERT_FAILURE "psStruct != nullptr", "src/structure.cpp:5840", "validStructResistance", "Invalid structure pointer": validStructResistance = False: Exit Function

    If asStructureStats(aStruct(psStruct).pStructureType).upgrade(aStruct(psStruct).player).resistance <> 0 Then
        ' certain structures will only provide rewards in multiplayer so
        ' before they can become valid targets their resistance must be at least
        ' half the base value
        If bMultiPlayer Then
            Select Case asStructureStats(aStruct(psStruct).pStructureType).typ
            Case REF_RESEARCH, REF_FACTORY, REF_VTOL_FACTORY, REF_CYBORG_FACTORY, REF_HQ, REF_REPAIR_FACILITY
                If Not ULT(aStruct(psStruct).resistance, structureResistance(aStruct(psStruct).pStructureType, aStruct(psStruct).player) \ 2) Then
                    bTarget = True
                End If
            Case Else
                bTarget = True
            End Select
        Else
            bTarget = True
        End If
    End If

    validStructResistance = bTarget
End Function

' compares the structure sensor type with the droid weapon type to see if the
' FIRE_SUPPORT order can be assigned
Public Function structSensorDroidWeapon(ByVal psStruct As Long, ByVal psDroid As Long) As Boolean
    Dim weaponStats As Long
    'another crash when nStat is marked as 0xcd... FIXME: Why is nStat not initialized properly?
    'Added a safety check: Only units with weapons can be assigned.
    If aDroid(psDroid).numWeaps > 0 Then
        weaponStats = DROID_getWeaponStats(psDroid, 0)
        'Standard Sensor Tower + indirect weapon droid (non VTOL)
        If structStandardSensor(psStruct) And (aDroid(psDroid).asWeaps(0).nStat <> 0 And _
                                               Not proj_Direct(weaponStats)) And _
           Not DROID_isVtol(psDroid) Then
            structSensorDroidWeapon = True
            Exit Function
        'CB Sensor Tower + indirect weapon droid (non VTOL)
        ElseIf structCBSensor(psStruct) And (aDroid(psDroid).asWeaps(0).nStat <> 0 And _
                                             Not proj_Direct(weaponStats)) And _
               Not DROID_isVtol(psDroid) Then
            structSensorDroidWeapon = True
            Exit Function
        'VTOL Intercept Sensor Tower + any weapon VTOL droid
        ElseIf structVTOLSensor(psStruct) And aDroid(psDroid).asWeaps(0).nStat <> 0 And _
               DROID_isVtol(psDroid) Then
            structSensorDroidWeapon = True
            Exit Function
        'VTOL CB Sensor Tower + any weapon VTOL droid
        ElseIf structVTOLCBSensor(psStruct) And aDroid(psDroid).asWeaps(0).nStat <> 0 And _
               DROID_isVtol(psDroid) Then
            structSensorDroidWeapon = True
            Exit Function
        End If
    End If

    'case not matched
    structSensorDroidWeapon = False
End Function

' (VB) the four sensor tests below: the structure's sensor is of one of the two types and on a turret
Private Function structSensorIs(ByVal psStruct As Long, ByVal type1 As SENSOR_TYPE, ByVal type2 As SENSOR_TYPE) As Boolean
    Dim s As Long
    s = asStructureStats(aStruct(psStruct).pStructureType).pSensor
    If s <> -1 Then
        If (asSensorStats(s).typ = type1 Or asSensorStats(s).typ = type2) And asSensorStats(s).location = LOC_TURRET Then
            structSensorIs = True
        End If
    End If
End Function

' checks if the structure has a Counter Battery sensor attached - returns
' true if it has
Public Function structCBSensor(ByVal psStruct As Long) As Boolean
    ' Super Sensor works as any type
    structCBSensor = structSensorIs(psStruct, INDIRECT_CB_SENSOR, SUPER_SENSOR)
End Function

' checks if the structure has a Standard Turret sensor attached - returns
' true if it has
Public Function structStandardSensor(ByVal psStruct As Long) As Boolean
    ' Super Sensor works as any type
    structStandardSensor = structSensorIs(psStruct, STANDARD_SENSOR, SUPER_SENSOR)
End Function

' checks if the structure has a VTOL Intercept sensor attached - returns
' true if it has
Public Function structVTOLSensor(ByVal psStruct As Long) As Boolean
    ' Super Sensor works as any type
    structVTOLSensor = structSensorIs(psStruct, VTOL_INTERCEPT_SENSOR, SUPER_SENSOR)
End Function

' checks if the structure has a VTOL Counter Battery sensor attached - returns
' true if it has
Public Function structVTOLCBSensor(ByVal psStruct As Long) As Boolean
    ' Super Sensor works as any type
    structVTOLCBSensor = structSensorIs(psStruct, VTOL_CB_SENSOR, SUPER_SENSOR)
End Function

' check whether a rearm pad is clear
Public Function clearRearmPad(ByVal psStruct As Long) As Boolean
    If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REARM_PAD Then
        If aStruct(psStruct).functionality.rearmPad.psObj = 0 Then
            clearRearmPad = True
        Else
            clearRearmPad = vtolHappy(aStruct(psStruct).functionality.rearmPad.psObj)
        End If
    End If
End Function

' return the nearest rearm pad
' psTarget can be NULL
Public Function findNearestReArmPad(ByVal psDroid As Long, ByVal psTarget As Long, ByVal bClear As Boolean) As Long
    Dim psNearest As Long, psTotallyClear As Long, xdiff As Long, ydiff As Long, mindist As Long, currdist As Long, totallyDist As Long
    Dim cx As Long, cy As Long, it As Long, psStruct As Long, cond As Boolean

    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/structure.cpp:6765", "findNearestReArmPad", "No droid was passed.": findNearestReArmPad = 0: Exit Function

    If psTarget <> 0 Then
        If Not vtolOnRearmPad(psTarget, psDroid) Then
            findNearestReArmPad = psTarget
            Exit Function
        End If
        cx = aStruct(psTarget).pos.x
        cy = aStruct(psTarget).pos.y
    Else
        cx = aDroid(psDroid).pos.x
        cy = aDroid(psDroid).pos.y
    End If
    mindist = SDWORD_MAX
    totallyDist = SDWORD_MAX
    psNearest = 0
    psTotallyClear = 0
    it = apsStructLists(aDroid(psDroid).player).head
    Do While it <> 0
        psStruct = lnObj(it)
        cond = asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REARM_PAD
        If cond And bClear Then cond = clearRearmPad(psStruct)
        If cond Then
            xdiff = aStruct(psStruct).pos.x - cx
            ydiff = aStruct(psStruct).pos.y - cy
            currdist = FromU32(CDbl(xdiff) * xdiff + CDbl(ydiff) * ydiff)
            cond = bClear
            If cond Then cond = Not vtolOnRearmPad(psStruct, psDroid)
            If cond Then
                If currdist < totallyDist Then
                    totallyDist = currdist
                    psTotallyClear = psStruct
                End If
            Else
                If currdist < mindist Then
                    mindist = currdist
                    psNearest = psStruct
                End If
            End If
        End If
        it = lnNext(it)
    Loop
    If bClear And (psTotallyClear <> 0) Then
        psNearest = psTotallyClear
    End If
    If psNearest = 0 Then
        If bClear Then
            objTrace aDroid(psDroid).id, "findNearestReArmPad", "Failed to find a landing spot (%s)!", "req clear"
        Else
            objTrace aDroid(psDroid).id, "findNearestReArmPad", "Failed to find a landing spot (%s)!", "any"
        End If
    End If
    findNearestReArmPad = psNearest
End Function

' clear a rearm pad for a droid to land on it
Public Sub ensureRearmPadClear(ByVal psStruct As Long, ByVal psDroid As Long)
    Dim tx As Long, ty As Long, i As Long, it As Long, psCurr As Long
    tx = map_coord(aStruct(psStruct).pos.x)
    ty = map_coord(aStruct(psStruct).pos.y)

    For i = 0 To MAX_PLAYERS - 1
        If aiCheckAlliances(aStruct(psStruct).player, i) Then
            it = apsDroidLists(i).head
            Do While it <> 0
                psCurr = lnObj(it)
                If psCurr <> psDroid _
                   And map_coord(aDroid(psCurr).pos.x) = tx _
                   And map_coord(aDroid(psCurr).pos.y) = ty Then
                    If DROID_isVtol(psCurr) Then
                        actionDroid_obj psCurr, DACTION_CLEARREARMPAD, psStruct
                    End If
                End If
                it = lnNext(it)
            Loop
        End If
    Next
End Sub

' return whether a rearm pad has a vtol on it
Public Function vtolOnRearmPad(ByVal psStruct As Long, ByVal psDroid As Long) As Boolean
    Dim tx As Long, ty As Long, it As Long, psCurr As Long

    tx = map_coord(aStruct(psStruct).pos.x)
    ty = map_coord(aStruct(psStruct).pos.y)

    it = apsDroidLists(aStruct(psStruct).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If psCurr <> psDroid _
           And map_coord(aDroid(psCurr).pos.x) = tx _
           And map_coord(aDroid(psCurr).pos.y) = ty Then
            vtolOnRearmPad = True
            Exit Function
        End If
        it = lnNext(it)
    Loop

    vtolOnRearmPad = False
End Function

' Just returns true if the structure's present body points aren't as high as the original
Public Function STRUCTURE_isDamaged(ByVal this As Long) As Boolean
    STRUCTURE_isDamaged = ULT(aStruct(this).body, STRUCTURE_structureBody(this))
End Function

' give a structure from one player to another - used in Electronic Warfare
'returns pointer to the new structure
Public Function giftSingleStructure(ByVal psStructure As Long, ByVal attackPlayer As Long, Optional ByVal electronic_warfare As Boolean = True) As Long
    Dim psNewStruct As Long
    Dim psType As Long, psModule As Long
    Dim x As Long, y As Long
    Dim capacity As Long, originalPlayer As Long
    Dim buildPoints As Long
    Dim bPowerOn As Boolean
    Dim direction As Long
    Dim prevState As Long, reward As Boolean, maxIndex As Long, it As Long, psCurr As Long, i As Long, psStruct As Long
    capacity = 0
    buildPoints = 0

    If Not (attackPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "attackPlayer < MAX_PLAYERS", "src/structure.cpp:6887", "giftSingleStructure", "attackPlayer (%u) must be < MAX_PLAYERS", attackPlayer: giftSingleStructure = 0: Exit Function
    checkStructure psStructure, "src/structure.cpp:6888", "giftSingleStructure", max_check_object_recursion
    visRemoveVisibility psStructure

    prevState = intGetResearchState()
    reward = electronicReward(psStructure, attackPlayer)

    If bMultiPlayer Then
        'certain structures give specific results - the rest swap sides!
        If Not electronic_warfare Or Not reward Then
            originalPlayer = aStruct(psStructure).player
            'tell the system the structure no longer exists
            removeStruct psStructure, False

            ' remove structure from one list
            removeStructureFromList psStructure, apsStructLists

            aStruct(psStructure).selected = 0

            ' change player id
            aStruct(psStructure).player = attackPlayer

            'restore the resistance value
            aStruct(psStructure).resistance = structureResistance(aStruct(psStructure).pStructureType, aStruct(psStructure).player) And &HFFFF&

            ' add to other list.
            addStructure psStructure

            ' increment structure count for new owner
            maxIndex = aStruct(psStructure).pStructureType     ' UDWORD max
            asStructureStats(maxIndex).curCount(attackPlayer) = asStructureStats(maxIndex).curCount(attackPlayer) + 1

            'check through the 'attackPlayer' players list of droids to see if any are targetting it
            it = apsDroidLists(attackPlayer).head
            Do While it <> 0
                psCurr = lnObj(it)
                If aDroid(psCurr).order.psObj = psStructure Then
                    orderDroid psCurr, DORDER_STOP, ModeImmediate
                    Exit Do
                End If
                For i = 0 To aDroid(psCurr).numWeaps - 1
                    If aDroid(psCurr).psActionTarget(i) = psStructure Then
                        orderDroid psCurr, DORDER_STOP, ModeImmediate
                        Exit For
                    End If
                Next
                'check through order list
                orderClearTargetFromDroidList psCurr, psStructure
                it = lnNext(it)
            Loop

            'check through the 'attackPlayer' players list of structures to see if any are targetting it
            it = apsStructLists(attackPlayer).head
            Do While it <> 0
                psStruct = lnObj(it)
                If aStruct(psStruct).psTarget(0) = psStructure Then
                    setStructureTarget psStruct, 0, 0, ORIGIN_UNKNOWN
                End If
                it = lnNext(it)
            Loop

            transferFixupFunctionality psStructure, asStructureStats(aStruct(psStructure).pStructureType).typ, originalPlayer

            If aStruct(psStructure).status = SS_BUILT Then
                buildingComplete psStructure
            End If
            'since the structure isn't being rebuilt, the visibility code needs to be adjusted
            'make sure this structure is visible to selectedPlayer
            aStruct(psStructure).visible(attackPlayer) = UINT8_MAX
            triggerEventObjectTransfer psStructure, originalPlayer
        End If
        intNotifyResearchButton prevState
        giftSingleStructure = 0
        Exit Function
    End If

    'save info about the structure
    psType = aStruct(psStructure).pStructureType
    x = aStruct(psStructure).pos.x
    y = aStruct(psStructure).pos.y
    direction = aStruct(psStructure).rot.direction
    originalPlayer = aStruct(psStructure).player
    'save how complete the build process is
    If aStruct(psStructure).status = SS_BEING_BUILT Then
        buildPoints = aStruct(psStructure).currentBuildPts And &HFFFF&     ' SWORD
        If buildPoints >= &H8000& Then buildPoints = buildPoints - &H10000
    End If
    'check module not attached
    psModule = getModuleStat(psStructure)
    'get rid of the structure
    removeStruct psStructure, True

    'make sure power is not used to build
    bPowerOn = powerCalculated
    powerCalculated = False
    'build a new one for the attacking player - set last element to true so it doesn't adjust x/y
    psNewStruct = buildStructure(psType, x, y, attackPlayer, True)
    capacity = aStruct(psStructure).capacity
    If psNewStruct <> 0 Then
        aStruct(psNewStruct).rot.direction = direction
        If capacity <> 0 Then
            Select Case asStructureStats(psType).typ
            Case REF_POWER_GEN, REF_RESEARCH
                'build the module for powerGen and research
                buildStructure psModule, aStruct(psNewStruct).pos.x, aStruct(psNewStruct).pos.y, attackPlayer, False
            Case REF_FACTORY, REF_VTOL_FACTORY
                'build the appropriate number of modules
                Do While capacity <> 0
                    buildStructure psModule, aStruct(psNewStruct).pos.x, aStruct(psNewStruct).pos.y, attackPlayer, False
                    capacity = capacity - 1
                Loop
            Case Else
            End Select
        End If
        If buildPoints <> 0 Then
            aStruct(psNewStruct).status = SS_BEING_BUILT
            aStruct(psNewStruct).currentBuildPts = buildPoints
        Else
            aStruct(psNewStruct).status = SS_BUILT
            buildingComplete psNewStruct
            triggerEventStructBuilt psNewStruct, 0
            checkPlayerBuiltHQ psNewStruct
        End If

        If Not bMultiPlayer Then
            If originalPlayer = selectedPlayer Then
                'make sure this structure is visible to selectedPlayer if the structure used to be selectedPlayers'
                If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < 11", "src/structure.cpp:7030", "giftSingleStructure", "selectedPlayer (%u) must be < MAX_PLAYERS", selectedPlayer
                aStruct(psNewStruct).visible(selectedPlayer) = UBYTE_MAX
            End If
            If Not electronic_warfare Or Not reward Then
                triggerEventObjectTransfer psNewStruct, originalPlayer
            End If
        End If
    End If
    powerCalculated = bPowerOn
    intNotifyResearchButton prevState
    giftSingleStructure = psNewStruct
End Function

' returns the power cost to build this structure, or to add its next module
Public Function structPowerToBuildOrAddNextModule(ByVal psStruct As Long) As Long
    Dim psStats As Long
    If aStruct(psStruct).capacity <> 0 Then
        psStats = getModuleStat(psStruct)
        If Not (psStats <> -1) Then ASSERT_FAILURE "psStats != nullptr", "src/structure.cpp:7050", "structPowerToBuildOrAddNextModule", "getModuleStat returned null"
        If psStats <> -1 Then
            structPowerToBuildOrAddNextModule = asStructureStats(psStats).powerToBuild ' return the cost to build the module
            Exit Function
        End If
    End If
    ' no module attached so building the base structure
    structPowerToBuildOrAddNextModule = asStructureStats(aStruct(psStruct).pStructureType).powerToBuild
End Function

' This follows the logic in droid.cpp nextModuleToBuild()
Public Function canStructureHaveAModuleAdded(ByVal structure_ As Long) As Boolean
    If structure_ = 0 Then
        canStructureHaveAModuleAdded = False
        Exit Function
    End If
    If aStruct(structure_).pStructureType = -1 Then
        canStructureHaveAModuleAdded = False
        Exit Function
    End If
    If aStruct(structure_).status <> SS_BUILT Then
        canStructureHaveAModuleAdded = False
        Exit Function
    End If

    Select Case asStructureStats(aStruct(structure_).pStructureType).typ
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        canStructureHaveAModuleAdded = aStruct(structure_).capacity < NUM_FACTORY_MODULES

    Case REF_POWER_GEN, REF_RESEARCH
        canStructureHaveAModuleAdded = aStruct(structure_).capacity = 0

    Case Else
        canStructureHaveAModuleAdded = False
    End Select
End Function

' LineBuild calcLineBuild(Vector2i size, STRUCTURE_TYPE type, Vector2i worldPos, Vector2i worldPos2)
Public Function calcLineBuild_size(size As Vector2i, ByVal typ As STRUCTURE_TYPE, worldPos As Vector2i, worldPos2 As Vector2i) As LineBuild
    Dim packed As Boolean, tile As Vector2i, padding_ As Vector2i, paddedSize As Vector2i, worldSize As Vector2i, worldPaddedSize As Vector2i
    Dim lb As LineBuild, delta As Vector2i, count As Vector2i

    If Not (size.x > 0 And size.y > 0) Then ASSERT_FAILURE "size.x > 0 && size.y > 0", "src/structure.cpp:7301", "calcLineBuild", "Zero-size building": Exit Function

    packed = typ = REF_RESOURCE_EXTRACTOR Or baseStructureTypePackability(typ) <= PACKABILITY_DEFENSE
    tile.x = TILE_UNITS: tile.y = TILE_UNITS
    If packed Then
        padding_.x = 0: padding_.y = 0
    Else
        padding_.x = 1: padding_.y = 1
    End If
    paddedSize.x = size.x + padding_.x
    paddedSize.y = size.y + padding_.y
    worldSize.x = world_coord(size.x)
    worldSize.y = world_coord(size.y)
    worldPaddedSize.x = world_coord(paddedSize.x)
    worldPaddedSize.y = world_coord(paddedSize.y)

    lb.begin_.x = round_to_nearest_tile(worldPos.x - worldSize.x \ 2) + worldSize.x \ 2
    lb.begin_.y = round_to_nearest_tile(worldPos.y - worldSize.y \ 2) + worldSize.y \ 2

    delta.x = worldPos2.x - lb.begin_.x
    delta.y = worldPos2.y - lb.begin_.y
    count.x = (Abs(delta.x) + worldPaddedSize.x \ 2) \ paddedSize.x + tile.x
    count.y = (Abs(delta.y) + worldPaddedSize.y \ 2) \ paddedSize.y + tile.y
    lb.count = map_coord(MAX_(count.x, count.y))
    If lb.count <= 1 Then
        lb.step_.x = 0: lb.step_.y = 0
    ElseIf count.x > count.y Then
        If delta.x < 0 Then lb.step_.x = -worldPaddedSize.x Else lb.step_.x = worldPaddedSize.x
        lb.step_.y = round_to_nearest_tile(delta.y \ (lb.count - 1))
    Else
        lb.step_.x = round_to_nearest_tile(delta.x \ (lb.count - 1))
        If delta.y < 0 Then lb.step_.y = -worldPaddedSize.y Else lb.step_.y = worldPaddedSize.y
    End If

    calcLineBuild_size = lb
End Function

' LineBuild calcLineBuild(STRUCTURE_STATS const *stats, uint16_t direction, Vector2i pos, Vector2i pos2)
Public Function calcLineBuild(ByVal stats As Long, ByVal direction As Long, pos As Vector2i, pos2 As Vector2i) As LineBuild
    Dim size As Vector2i
    size = STRUCTURE_STATS_size(stats, direction)
    calcLineBuild = calcLineBuild_size(size, asStructureStats(stats).typ, pos, pos2)
End Function

' Vector2i LineBuild::operator [](int i) const
Public Function LineBuild_at(lb As LineBuild, ByVal i As Long) As Vector2i
    LineBuild_at.x = lb.begin_.x + i * lb.step_.x
    LineBuild_at.y = lb.begin_.y + i * lb.step_.y
End Function

' Vector2i LineBuild::back() const
Public Function LineBuild_back(lb As LineBuild) As Vector2i
    LineBuild_back = LineBuild_at(lb, lb.count - 1)
End Function

' Deals damage to a Structure.
' \param psStructure structure to deal damage to
' \param damage amount of damage to deal
' \param weaponClass the class of the weapon that deals the damage
' \param weaponSubClass the subclass of the weapon that deals the damage
' \return < 0 when the dealt damage destroys the structure, > 0 when the structure survives
Public Function structureDamage(ByVal psStructure As Long, ByVal psProjectile As Long, ByVal damage As Long, ByVal weaponClass As WEAPON_CLASS, ByVal weaponSubClass As WEAPON_SUBCLASS, ByVal impactTime As Long, ByVal isDamagePerSecond As Boolean, ByVal minDamage As Long, ByVal empRadiusHit As Boolean) As Long
    Dim relativeDamage As Long

    checkStructure psStructure, "src/structure.cpp:858", "structureDamage", max_check_object_recursion

    debug_ 860, LOG_ATTACK, "structureDamage", "structure id %d, body %d, armour %d, damage: %d", _
           aStruct(psStructure).id, aStruct(psStructure).body, objArmour(psStructure, weaponClass), damage

    relativeDamage = objDamage(psStructure, psProjectile, damage, STRUCTURE_structureBody(psStructure), weaponClass, weaponSubClass, isDamagePerSecond, minDamage, empRadiusHit)

    ' If the shell did sufficient damage to destroy the structure
    If relativeDamage < 0 Then
        debug_ 868, LOG_ATTACK, "structureDamage", "Structure (id %d) DESTROYED", aStruct(psStructure).id
        destroyStruct psStructure, impactTime
    Else
        ' Survived
        checkStructure psStructure, "src/structure.cpp:874", "structureDamage", max_check_object_recursion
    End If

    structureDamage = relativeDamage
End Function

' Remove a structure
Public Function destroyStruct(ByVal psDel As Long, ByVal impactTime As Long) As Boolean
    Const burnDurationWall As Long = 1000
    Const burnDurationOilWell As Long = 60000
    Const burnDurationOther As Long = 10000
    Dim widthScatter As Long, breadthScatter As Long, heightScatter As Long, typ As STRUCTURE_TYPE
    Dim bMinor As Boolean, bDerrick As Boolean, bPowerGen As Boolean, burnDuration As Long, pos As Vector3i, i As Long
    Dim resourceFound As Boolean, b As StructureBounds, breadth As Long, width As Long, psTile As Long

    checkStructure psDel, "src/structure.cpp:4689", "destroyStruct", max_check_object_recursion
    If Not (Not ULT(impactTime, FromU32(U32(gameTime) - U32(deltaGameTime)))) Then ASSERT_FAILURE "gameTime - deltaGameTime <= impactTime", "src/structure.cpp:4690", "destroyStruct", "Expected %u <= %u, gameTime = %u, bad impactTime", FromU32(U32(gameTime) - U32(deltaGameTime)), impactTime, gameTime

    ' Firstly, are we dealing with a wall section
    typ = asStructureStats(aStruct(psDel).pStructureType).typ
    bMinor = typ = REF_WALL Or typ = REF_WALLCORNER
    bDerrick = typ = REF_RESOURCE_EXTRACTOR
    bPowerGen = typ = REF_POWER_GEN
    If bMinor Then
        burnDuration = burnDurationWall
    ElseIf bDerrick Then
        burnDuration = burnDurationOilWell
    Else
        burnDuration = burnDurationOther
    End If
    If aStruct(psDel).status = SS_BEING_BUILT Then
        burnDuration = FromU32(Fix(CDbl(fmul(F32(U32(burnDuration)), structureCompletionProgress(psDel)))))
    End If

    ' Only add if visible
    If BASE_OBJECT_visibleForLocalDisplay(psDel) <> 0 Then
        ' Set off some explosions, but not for walls
        ' First Explosions
        widthScatter = TILE_UNITS
        breadthScatter = TILE_UNITS
        heightScatter = TILE_UNITS
        For i = 0 To IIf(bMinor, 2, 4) - 1  ' only add two for walls - gets crazy otherwise
            pos.x = aStruct(psDel).pos.x + widthScatter - rand() Mod (2 * widthScatter)
            pos.z = aStruct(psDel).pos.y + breadthScatter - rand() Mod (2 * breadthScatter)
            pos.y = aStruct(psDel).pos.z + 32 + rand() Mod heightScatter
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_MEDIUM, False, 0, 0, impactTime
        Next

        ' Get coordinates for everybody!
        pos.x = aStruct(psDel).pos.x
        pos.z = aStruct(psDel).pos.y  ' z = y [sic] intentional
        pos.y = map_Height(pos.x, pos.z)

        ' Set off a fire, provide dimensions for the fire
        If bMinor Then
            effectGiveAuxVar world_coord(asStructureStats(aStruct(psDel).pStructureType).baseWidth) \ 4
        Else
            effectGiveAuxVar world_coord(asStructureStats(aStruct(psDel).pStructureType).baseWidth) \ 3
        End If
        ' Give a duration
        effectGiveAuxVarSec burnDuration
        If bDerrick Then  ' oil resources
            ' Oil resources burn AND puff out smoke AND for longer
            addEffect pos, EFFECT_FIRE, FIRE_TYPE_SMOKY, False, 0, 0, impactTime
        Else  ' everything else
            addEffect pos, EFFECT_FIRE, FIRE_TYPE_LOCALISED, False, 0, 0, impactTime
        End If

        ' Power stations have their own destruction sequence
        If bPowerGen Then
            addEffect pos, EFFECT_DESTRUCTION, DESTRUCTION_TYPE_POWER_STATION, False, 0, 0, impactTime
            pos.y = pos.y + SHOCK_WAVE_HEIGHT
            addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_SHOCKWAVE, False, 0, 0, impactTime
        ' As do wall sections
        ElseIf bMinor Then
            addEffect pos, EFFECT_DESTRUCTION, DESTRUCTION_TYPE_WALL_SECTION, False, 0, 0, impactTime
        Else ' and everything else goes here.....
            addEffect pos, EFFECT_DESTRUCTION, DESTRUCTION_TYPE_STRUCTURE, False, 0, 0, impactTime
        End If

        ' shake the screen if we're near enough and it is explosive in nature
        If clipXY(pos.x, pos.z) Then
            Select Case typ
            ' These are the types that would cause a explosive outcome if destoryed
            Case REF_HQ, REF_POWER_GEN, REF_MISSILE_SILO  ' REF_MISSILE_SILO: for campaign
                shakeStart 1500
            Case REF_COMMAND_CONTROL, REF_VTOL_FACTORY, REF_CYBORG_FACTORY, REF_FACTORY
                shakeStart 750
            Case REF_RESOURCE_EXTRACTOR
                shakeStart 400
            Case Else
            End Select
        End If

        ' and add a sound effect
        audio_PlayStaticTrack aStruct(psDel).pos.x, aStruct(psDel).pos.y, ID_SOUND_EXPLOSION
    End If

    ' Actually set the tiles on fire - even if the effect is not visible.
    tileSetFire aStruct(psDel).pos.x, aStruct(psDel).pos.y, burnDuration

    resourceFound = removeStruct(psDel, True)
    aStruct(psDel).died = impactTime

    ' Leave burn marks in the ground where building once stood
    If BASE_OBJECT_visibleForLocalDisplay(psDel) <> 0 And Not resourceFound And Not bMinor Then
        b = getStructureBounds(psDel)
        For breadth = 0 To b.size.y - 1
            For width = 0 To b.size.x - 1
                psTile = mapTile(b.map.x + width, b.map.y + breadth)
                If TEST_TILE_VISIBLE_TO_SELECTEDPLAYER(psTile) Then
                    psMapTiles(psTile).illumination = psMapTiles(psTile).illumination \ 2
                    psMapTiles(psTile).ambientOcclusion = psMapTiles(psTile).ambientOcclusion \ 2
                End If
            Next
        Next
    End If

    If bMultiPlayer Then
        technologyGiveAway psDel  ' Drop an artifact, if applicable.
    End If

    ' updates score stats only if not wall
    If Not bMinor Then
        If aStruct(psDel).player = selectedPlayer Then
            scoreUpdateVar WD_STR_LOST
        ' only counts as a kill if structure doesn't belong to our ally
        ElseIf ULT(selectedPlayer, MAX_PLAYERS) Then
            If Not aiCheckAlliances(aStruct(psDel).player, selectedPlayer) Then
                scoreUpdateVar WD_STR_KILLED
            End If
        End If
    End If

    destroyStruct = True
End Function

' calculate muzzle tip location in 3d world
Public Function calcStructureMuzzleLocation(ByVal psStructure As Long, muzzle As Vector3i, ByVal weapon_slot As Long) As Boolean
    Dim psShape As Long, barrel As Vector3i, nWeaponStat As Long, psWeaponImd As Long, psMountImd As Long
    Dim af As Affine3F, cx As Long, cy As Long, cz As Long, r As Vector3i, connector_num As Long, nConnectors As Long

    psShape = asStructureStats(aStruct(psStructure).pStructureType).pIMD(0)

    checkStructure psStructure, "src/structure.cpp:5111", "calcStructureMuzzleLocation", max_check_object_recursion

    nConnectors = 0
    If psShape <> 0 Then nConnectors = iIMDBaseShape_connectors_size(psShape)
    If nConnectors <> 0 Then
        barrel.x = 0: barrel.y = 0: barrel.z = 0
        nWeaponStat = aStruct(psStructure).asWeaps(weapon_slot).nStat
        psWeaponImd = 0
        psMountImd = 0

        If nWeaponStat <> 0 Then
            psWeaponImd = asWeaponStats(nWeaponStat).pIMD
            psMountImd = asWeaponStats(nWeaponStat).pMountGraphic
        End If

        Affine3F_construct af

        Affine3F_Trans af, aStruct(psStructure).pos.x, -aStruct(psStructure).pos.z, aStruct(psStructure).pos.y

        'matrix = the center of droid
        Affine3F_RotY af, aStruct(psStructure).rot.direction
        Affine3F_RotX af, aStruct(psStructure).rot.pitch
        Affine3F_RotZ af, -aStruct(psStructure).rot.roll
        iIMDBaseShape_connector psShape, weapon_slot, cx, cy, cz
        Affine3F_Trans af, cx, -cz, -cy 'note y and z flipped

        'matrix = the weapon[slot] mount on the body
        Affine3F_RotY af, aStruct(psStructure).asWeaps(weapon_slot).rot.direction  ' +ve anticlockwise

        ' process turret mount
        If psMountImd <> 0 Then
            If iIMDBaseShape_connectors_size(psMountImd) <> 0 Then
                iIMDBaseShape_connector psMountImd, 0, cx, cy, cz
                Affine3F_Trans af, cx, -cz, -cy
            End If
        End If

        'matrix = the turret connector for the gun
        Affine3F_RotX af, aStruct(psStructure).asWeaps(weapon_slot).rot.pitch  ' +ve up

        'process the gun
        If psWeaponImd <> 0 Then
            nConnectors = iIMDBaseShape_connectors_size(psWeaponImd)
            If nConnectors <> 0 Then
                connector_num = 0

                ' which barrel is firing if model have multiple muzzle connectors?
                If aStruct(psStructure).asWeaps(weapon_slot).shotsFired <> 0 And (nConnectors > 1) Then
                    ' shoot first, draw later - substract one shot to get correct results
                    connector_num = FromU32(U32(FromU32(U32(aStruct(psStructure).asWeaps(weapon_slot).shotsFired) - 1)) - Int((U32(FromU32(U32(aStruct(psStructure).asWeaps(weapon_slot).shotsFired) - 1))) / nConnectors) * nConnectors)
                End If

                iIMDBaseShape_connector psWeaponImd, connector_num, cx, cy, cz
                barrel = Vector3i_(cx, -cz, -cy)
            End If
        End If

        r = Affine3F_mul_vec(af, barrel)
        muzzle.x = r.x  ' xzy()
        muzzle.y = r.z
        muzzle.z = r.y
        muzzle.z = -muzzle.z
    Else
        muzzle.x = aStruct(psStructure).pos.x
        muzzle.y = aStruct(psStructure).pos.y
        muzzle.z = aStruct(psStructure).pos.z + 0 + imdMaxY(aStruct(psStructure).sDisplay.imd)
    End If

    calcStructureMuzzleLocation = True
End Function

' calculates the damage caused to the resistance levels of structures - returns
' true when captured
Public Function electronicDamage(ByVal psTarget As Long, ByVal psAttacker As Long, ByVal damage As Long, ByVal attackPlayer As Long, ByVal tpCapture As Boolean) As Boolean
    Dim psStructure As Long, psDroid As Long, bCompleted As Boolean, pos As Vector3i, lastHit As Long, i As Long
    Dim giftPos As Vector2i, attackerPos As Vector3i

    bCompleted = True

    If Not (attackPlayer < MAX_PLAYERS) Then ASSERT_FAILURE "attackPlayer < MAX_PLAYERS", "src/structure.cpp:5722", "electronicDamage", "Invalid player id %d", attackPlayer: electronicDamage = False: Exit Function
    If Not (psTarget <> 0) Then ASSERT_FAILURE "psTarget != nullptr", "src/structure.cpp:5723", "electronicDamage", "Target is NULL": electronicDamage = False: Exit Function

    'structure electronic damage
    If SIMPLE_OBJECT_type(psTarget) = OBJ_STRUCTURE Then
        psStructure = psTarget
        bCompleted = False

        If asStructureStats(aStruct(psStructure).pStructureType).upgrade(aStruct(psStructure).player).resistance = 0 Then
            electronicDamage = False  ' this structure type cannot be taken over
            Exit Function
        End If

        'if resistance is already less than 0 don't do any more
        If aStruct(psStructure).resistance < 0 Then
            bCompleted = True
        Else
            'store the time it was hit
            lastHit = aStruct(psStructure).timeLastHit
            aStruct(psStructure).timeLastHit = gameTime

            aStruct(psStructure).lastHitWeapon = WSC_ELECTRONIC

            triggerEventAttacked psStructure, psAttacker, lastHit

            aStruct(psStructure).resistance = ToInt16(aStruct(psStructure).resistance - damage)

            If aStruct(psStructure).resistance < 0 Then
                'add a console message for the selected Player
                If aStruct(psStructure).player = selectedPlayer Then
                    console "%s - Electronically Damaged", asStructureStats(aStruct(psStructure).pStructureType).name_
                End If
                bCompleted = True
                'give the structure to the attacking player
                giftSingleStructure psStructure, attackPlayer
            End If
        End If
    'droid electronic damage
    ElseIf SIMPLE_OBJECT_type(psTarget) = OBJ_DROID Then
        psDroid = psTarget
        bCompleted = False
        lastHit = aDroid(psDroid).timeLastHit
        aDroid(psDroid).timeLastHit = gameTime
        aDroid(psDroid).lastHitWeapon = WSC_ELECTRONIC

        'in multiPlayer cannot attack a Transporter with EW
        If bMultiPlayer Then
            If Not (Not DROID_isTransporter(psDroid)) Then ASSERT_FAILURE "!psDroid->isTransporter()", "src/structure.cpp:5779", "electronicDamage", "Cannot attack a Transporter in multiPlayer": electronicDamage = True: Exit Function
        End If

        If aDroid(psDroid).resistance = ACTION_START_TIME Then
            'need to set the current resistance level since not been previously attacked (by EW)
            aDroid(psDroid).resistance = DROID_droidResistance(psDroid)
        End If

        If aDroid(psDroid).resistance < 0 Then
            bCompleted = True
        Else
            triggerEventAttacked psDroid, psAttacker, lastHit

            aDroid(psDroid).resistance = ToInt16(aDroid(psDroid).resistance - damage)

            If aDroid(psDroid).resistance <= 0 Then
                'add a console message for the selected Player
                If aDroid(psDroid).player = selectedPlayer Then
                    console "%s - Electronically Damaged", "Unit"
                End If
                bCompleted = True

                'give the droid to the attacking player

                If BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 Then ' display-only check for adding effect
                    For i = 0 To 4
                        pos.x = aDroid(psDroid).pos.x + (30 - rand() Mod 60)
                        pos.z = aDroid(psDroid).pos.y + (30 - rand() Mod 60)
                        pos.y = aDroid(psDroid).pos.z + (rand() Mod 8)
                        effectGiveAuxVar 80
                        addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_FLAMETHROWER, False, 0, 0, gameTime - deltaGameTime
                    Next
                End If
                If psAttacker <> 0 And tpCapture Then
                    attackerPos = SIMPLE_OBJECT_pos(psAttacker)
                    giftPos = Vector2i_(attackerPos.x, attackerPos.y)
                Else
                    giftPos = Vector2i_(aDroid(psDroid).pos.x, aDroid(psDroid).pos.y)
                End If
                If Not isDead(psDroid) Then
                    If giftSingleDroid(psDroid, attackPlayer, True, giftPos) = 0 Then
                        ' droid limit reached, recycle
                        ' don't check for transporter/mission coz multiplayer only issue.
                        recycleDroid psDroid
                    End If
                End If
            End If
        End If
    End If

    electronicDamage = bCompleted
End Function
' ---- B9: the structures' update, factory production and droid placement

'Initialise the production list and set up the production player
Public Sub changeProductionPlayer(ByVal player As Long)
    Dim typ As Long
    'clear the production run
    For typ = 0 To NUM_FACTORY_TYPES - 1
        Erase asProductionRun(typ).runs
        asProductionRun(typ).size = 0
    Next
    'set this player to have the production list
    productionPlayer = ToInt8(player)
End Sub

' set the current number of structures of each type built
Public Sub setCurrentStructQuantity(ByVal displayError As Boolean)
    Dim player As Long, inc As Long, it As Long
    For player = 0 To MAX_PLAYERS - 1
        For inc = 0 To numStructureStats - 1
            asStructureStats(inc).curCount(player) = 0
        Next
        it = apsStructLists(player).head
        Do While it <> 0
            inc = aStruct(lnObj(it)).pStructureType
            asStructureStats(inc).curCount(player) = asStructureStats(inc).curCount(player) + 1
            If displayError Then
                'check quantity never exceeds the limit
                If ULT(asStructureStats(inc).upgrade(player).limit, asStructureStats(inc).curCount(player)) Then ASSERT_FAILURE "asStructureStats[inc].curCount[player] <= asStructureStats[inc].upgrade[player].limit", "src/structure.cpp:762", "setCurrentStructQuantity", _
                       "There appears to be too many %s on this map!", asStructureStats(inc).name_     ' getStatsName(&asStructureStats[inc])
            End If
            it = lnNext(it)
        Loop
    Next
End Sub

' TODO: The abandoned code needs to be factored out, see: saveMissionData
Public Sub handleAbandonedStructures()
    ' TODO: do something here
End Sub

Private Sub refundFactoryBuildPower(ByVal psBuilding As Long)
    Dim psFactory As Long
    If psBuilding = 0 Then ASSERT_FAILURE "psBuilding && psBuilding->isFactory()", "src/structure.cpp:1131", "refundFactoryBuildPower", "structure not a factory": Exit Sub
    If Not STRUCTURE_isFactory(psBuilding) Then ASSERT_FAILURE "psBuilding && psBuilding->isFactory()", "src/structure.cpp:1131", "refundFactoryBuildPower", "structure not a factory": Exit Sub
    psFactory = psBuilding  ' FACTORY *

    If aStruct(psFactory).functionality.factory.psSubject <> 0 Then
        If aStruct(psFactory).functionality.factory.buildPointsRemaining < calcTemplateBuild(apsTemplates(aStruct(psFactory).functionality.factory.psSubject)) Then
            ' We started building, so give the power back that was used.
            addPower aStruct(psBuilding).player, calcTemplatePower(apsTemplates(aStruct(psFactory).functionality.factory.psSubject))
        End If

    End If
End Sub

' structure.h's setStatusPendingStart and setStatusPendingCancel, for a FACTORY
Private Sub setStatusPendingStart_factory(func_ As FACTORY, ByVal subject As Long)
    func_.psSubjectPending = subject
    func_.statusPending = FACTORY_START_PENDING
    func_.pendingCount = func_.pendingCount + 1
End Sub

Private Sub setStatusPendingCancel_factory(func_ As FACTORY)
    func_.psSubjectPending = 0
    func_.statusPending = FACTORY_CANCEL_PENDING
    func_.pendingCount = func_.pendingCount + 1
End Sub

' Set the type of droid for a factory to build
Public Function structSetManufacture(ByVal psStruct As Long, ByVal psTempl As Long, ByVal mode As QUEUE_MODE) As Boolean
    Dim psFact As Long, ok As Boolean

    checkStructure psStruct, "src/structure.cpp:1148", "structSetManufacture", max_check_object_recursion

    If Not (STRUCTURE_isFactory(psStruct)) Then ASSERT_FAILURE "psStruct->isFactory()", "src/structure.cpp:1150", "structSetManufacture", "structure not a factory": structSetManufacture = False: Exit Function

    ' psTempl might be NULL if the build is being cancelled in the middle
    ok = psTempl = 0
    If Not ok Then
        ok = validTemplateForFactory(psTempl, psStruct, True)
        If ok Then ok = researchedTemplate(psTempl, aStruct(psStruct).player, True, True)
    End If
    If Not ok Then ok = aStruct(psStruct).player = scavengerPlayer()
    If Not ok Then ok = Not bMultiPlayer
    If Not ok Then ASSERT_FAILURE "!psTempl || (validTemplateForFactory(psTempl, psStruct, true) && researchedTemplate(psTempl, psStruct->player, true, true)) || psStruct->player == scavengerPlayer() || !bMultiPlayer", "src/structure.cpp:1153", "structSetManufacture", _
                                  "Wrong template for player %d factory, type %d.", aStruct(psStruct).player, asStructureStats(aStruct(psStruct).pStructureType).typ: structSetManufacture = False: Exit Function

    psFact = psStruct  ' FACTORY *

    If mode = ModeQueue Then
        sendStructureInfo psStruct, STRUCTUREINFO_MANUFACTURE, psTempl
        setStatusPendingStart_factory aStruct(psFact).functionality.factory, psTempl
        structSetManufacture = True  ' Wait for our message before doing anything.
        Exit Function
    End If

    'assign it to the Factory
    refundFactoryBuildPower psStruct
    aStruct(psFact).functionality.factory.psSubject = psTempl

    'set up the start time and build time
    If psTempl <> 0 Then
        'only use this for non selectedPlayer
        If aStruct(psStruct).player <> selectedPlayer Then
            'set quantity to produce
            aStruct(psFact).functionality.factory.productionLoops = 1
        End If

        aStruct(psFact).functionality.factory.timeStarted = ACTION_START_TIME 'gameTime;
        aStruct(psFact).functionality.factory.timeStartHold = 0

        aStruct(psFact).functionality.factory.buildPointsRemaining = calcTemplateBuild(apsTemplates(psTempl))
        'check for zero build time - usually caused by 'silly' data! If so, set to 1 build point - ie very fast!
        aStruct(psFact).functionality.factory.buildPointsRemaining = MAX_(aStruct(psFact).functionality.factory.buildPointsRemaining, 1)
    End If
    structSetManufacture = True
End Function

Private Function transferFixupFunctionality(ByVal psBuilding As Long, ByVal functionType As STRUCTURE_TYPE, ByVal priorPlayer As Long) As Boolean
    Dim psAssemblyPoint As Long

    If Not (psBuilding <> 0) Then ASSERT_FAILURE "psBuilding != nullptr", "src/structure.cpp:2189", "transferFixupFunctionality", "Invalid pointer": transferFixupFunctionality = False: Exit Function
    checkStructure psBuilding, "src/structure.cpp:2190", "transferFixupFunctionality", max_check_object_recursion

    Select Case functionType
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY, REF_RESEARCH, REF_POWER_GEN, _
         REF_RESOURCE_EXTRACTOR, REF_REPAIR_FACILITY, REF_REARM_PAD, REF_WALL, REF_GATE
        ' Structure should already have functionality
        If Not (aStruct(psBuilding).pFunctionality <> 0) Then ASSERT_FAILURE "psBuilding->pFunctionality", "src/structure.cpp:2205", "transferFixupFunctionality", "Structure does not already have functionality pointer?": transferFixupFunctionality = False: Exit Function

    Case Else
        aStruct(psBuilding).pFunctionality = 0
    End Select

    Select Case functionType
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        ' (VB) FACTORY *psFactory is aStruct(psBuilding).functionality.factory

        ' Reset factoryNumFlag for prior player
        psAssemblyPoint = aStruct(psBuilding).functionality.factory.psAssemblyPoint
        If psAssemblyPoint <> 0 Then
            If aFlagPos(psAssemblyPoint).factoryInc < factoryNumFlag_size(priorPlayer, aFlagPos(psAssemblyPoint).factoryType) Then
                factoryNumFlag(priorPlayer, aFlagPos(psAssemblyPoint).factoryType, aFlagPos(psAssemblyPoint).factoryInc) = False
            End If

            'need to cancel the repositioning of the DP if selectedPlayer and currently moving
            If priorPlayer = selectedPlayer And aFlagPos(psAssemblyPoint).selected Then
                cancelDeliveryRepos
            End If
        End If

        ' Transfer / fix-up factory assembly point, and number
        transferFlagPositionToPlayer aStruct(psBuilding).functionality.factory.psAssemblyPoint, priorPlayer, aStruct(psBuilding).player

        Select Case functionType
        Case REF_FACTORY
            setFlagPositionInc aStruct(psBuilding).pFunctionality, aStruct(psBuilding).player, FACTORY_FLAG
        Case REF_CYBORG_FACTORY
            setFlagPositionInc aStruct(psBuilding).pFunctionality, aStruct(psBuilding).player, CYBORG_FLAG
        Case REF_VTOL_FACTORY
            setFlagPositionInc aStruct(psBuilding).pFunctionality, aStruct(psBuilding).player, VTOL_FLAG
        Case Else
            ASSERT_FAILURE "false", "src/structure.cpp:2252", "transferFixupFunctionality", "Invalid factory type": transferFixupFunctionality = False: Exit Function
        End Select
    Case REF_POWER_GEN, REF_HQ, REF_REARM_PAD
    Case REF_RESOURCE_EXTRACTOR
        ' (VB) RES_EXTRACTOR *psResExtracter is aStruct(psBuilding).functionality.resourceExtractor

        ' Make the structure inactive
        aStruct(psBuilding).functionality.resourceExtractor.psPowerGen = 0
    Case REF_REPAIR_FACILITY
        ' (VB) REPAIR_FACILITY *psRepairFac is aStruct(psBuilding).functionality.repairFacility

        ' POSSIBLE TODO: Do something about the group? (Or can we just keep it?)

        ' Reset factoryNumFlag for prior player
        psAssemblyPoint = aStruct(psBuilding).functionality.repairFacility.psDeliveryPoint
        If psAssemblyPoint <> 0 Then
            If aFlagPos(psAssemblyPoint).factoryInc < factoryNumFlag_size(priorPlayer, aFlagPos(psAssemblyPoint).factoryType) Then
                factoryNumFlag(priorPlayer, aFlagPos(psAssemblyPoint).factoryType, aFlagPos(psAssemblyPoint).factoryInc) = False
            End If

            'need to cancel the repositioning of the DP if selectedPlayer and currently moving
            If priorPlayer = selectedPlayer And aFlagPos(psAssemblyPoint).selected Then
                cancelDeliveryRepos
            End If
        End If

        ' Transfer / fix-up factory assembly point, and number
        transferFlagPositionToPlayer aStruct(psBuilding).functionality.repairFacility.psDeliveryPoint, priorPlayer, aStruct(psBuilding).player

        setFlagPositionInc aStruct(psBuilding).pFunctionality, aStruct(psBuilding).player, REPAIR_FLAG
    ' Structure types without a FUNCTIONALITY
    Case Else
    End Select

    transferFixupFunctionality = True
End Function

' Set the command droid that factory production should go to
Public Sub assignFactoryCommandDroid(ByVal psStruct As Long, ByVal psCommander As Long)
    Dim psFact As Long, factoryInc As Long, typeFlag As Long, player As Long, flagPosIt As Long, flagPosItNext As Long, flagPos As Long

    checkStructure psStruct, "src/structure.cpp:2312", "assignFactoryCommandDroid", max_check_object_recursion
    If Not (STRUCTURE_isFactory(psStruct)) Then ASSERT_FAILURE "psStruct->isFactory()", "src/structure.cpp:2313", "assignFactoryCommandDroid", "structure not a factory": Exit Sub

    psFact = psStruct  ' FACTORY *

    Select Case asStructureStats(aStruct(psStruct).pStructureType).typ
    Case REF_FACTORY
        typeFlag = FACTORY_FLAG
    Case REF_VTOL_FACTORY
        typeFlag = VTOL_FLAG
    Case REF_CYBORG_FACTORY
        typeFlag = CYBORG_FLAG
    Case Else
        ASSERT_FAILURE "!""unknown factory type""", "src/structure.cpp:2329", "assignFactoryCommandDroid", "unknown factory type"
        typeFlag = FACTORY_FLAG
    End Select

    ' removing a commander from a factory
    If aStruct(psFact).functionality.factory.psCommander <> 0 Then
        If typeFlag = FACTORY_FLAG Then
            secondarySetState aStruct(psFact).functionality.factory.psCommander, DSO_CLEAR_PRODUCTION, _
                              BIT_(aFlagPos(aStruct(psFact).functionality.factory.psAssemblyPoint).factoryInc + DSS_ASSPROD_SHIFT)
        ElseIf typeFlag = CYBORG_FLAG Then
            secondarySetState aStruct(psFact).functionality.factory.psCommander, DSO_CLEAR_PRODUCTION, _
                              BIT_(aFlagPos(aStruct(psFact).functionality.factory.psAssemblyPoint).factoryInc + DSS_ASSPROD_CYBORG_SHIFT)
        Else
            secondarySetState aStruct(psFact).functionality.factory.psCommander, DSO_CLEAR_PRODUCTION, _
                              BIT_(aFlagPos(aStruct(psFact).functionality.factory.psAssemblyPoint).factoryInc + DSS_ASSPROD_VTOL_SHIFT)
        End If

        aStruct(psFact).functionality.factory.psCommander = 0
        ' TODO: Synchronise .psCommander.
        'syncDebug("Removed commander from factory %d", psStruct->id);
        If Not missionIsOffworld() Then
            addFlagPosition aStruct(psFact).functionality.factory.psAssemblyPoint ' add the assembly point back into the list
        Else
            addFlagPositionToList aStruct(psFact).functionality.factory.psAssemblyPoint, mission_apsFlagPosLists
        End If
    End If

    If psCommander <> 0 Then
        If Not (Not missionIsOffworld()) Then ASSERT_FAILURE "!missionIsOffworld()", "src/structure.cpp:2368", "assignFactoryCommandDroid", "cannot assign a commander to a factory when off world": Exit Sub

        factoryInc = aFlagPos(aStruct(psFact).functionality.factory.psAssemblyPoint).factoryInc

        player = aStruct(psStruct).player   ' auto& flagPosList = apsFlagPosLists[psStruct->player]
        flagPosIt = apsFlagPosLists(player).head
        Do While flagPosIt <> 0
            flagPosItNext = lnNext(flagPosIt)
            flagPos = lnObj(flagPosIt)
            If aFlagPos(flagPos).factoryInc = factoryInc And aFlagPos(flagPos).factoryType = typeFlag Then
                If flagPos <> aStruct(psFact).functionality.factory.psAssemblyPoint Then
                    removeFlagPosition flagPos
                Else
                    ' need to keep the assembly point(s) for the factory
                    ' but remove it(the primary) from the list so it doesn't get
                    ' displayed
                    objlist_erase apsFlagPosLists(player), flagPosIt
                End If
            End If
            flagPosIt = flagPosItNext
        Loop
        aStruct(psFact).functionality.factory.psCommander = psCommander
        syncDebug "assignFactoryCommandDroid", "Assigned commander %d to factory %d", aDroid(psCommander).id, aStruct(psStruct).id
    End If
End Sub

' remove all factories from a command droid
Public Sub clearCommandDroidFactory(ByVal psDroid As Long)
    Dim it As Long, psCurr As Long, t As Long

    If Not (ULT(selectedPlayer, MAX_PLAYERS)) Then ASSERT_FAILURE "selectedPlayer < MAX_PLAYERS", "src/structure.cpp:2403", "clearCommandDroidFactory", "invalid selectedPlayer: %u", selectedPlayer: Exit Sub

    it = apsStructLists(selectedPlayer).head
    Do While it <> 0
        psCurr = lnObj(it)
        t = asStructureStats(aStruct(psCurr).pStructureType).typ
        If (t = REF_FACTORY) Or _
           (t = REF_CYBORG_FACTORY) Or _
           (t = REF_VTOL_FACTORY) Then
            If aStruct(psCurr).functionality.factory.psCommander = psDroid Then
                assignFactoryCommandDroid psCurr, 0
            End If
        End If
        it = lnNext(it)
    Loop
    it = mission_apsStructLists(selectedPlayer).head
    Do While it <> 0
        psCurr = lnObj(it)
        t = asStructureStats(aStruct(psCurr).pStructureType).typ
        If (t = REF_FACTORY) Or _
           (t = REF_CYBORG_FACTORY) Or _
           (t = REF_VTOL_FACTORY) Then
            If aStruct(psCurr).functionality.factory.psCommander = psDroid Then
                assignFactoryCommandDroid psCurr, 0
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

' Check that a tile is vacant for a droid to be placed
Private Function structClearTile(ByVal x As Long, ByVal y As Long, ByVal propulsion As Long) As Boolean
    Dim player As Long, it As Long, psCurr As Long, p As Long

    ' Check for a structure
    ' NOTE: Substitute PROPULSION_TYPE_WHEELED for PROPULSION_TYPE_LIFT, as flying droids are initially placed on the ground
    If propulsion <> PROPULSION_TYPE_LIFT Then p = propulsion Else p = PROPULSION_TYPE_WHEELED
    If fpathBlockingTile(x, y, p) Then
        debug_ 2440, LOG_NEVER, "structClearTile", "failed - blocked"
        structClearTile = False
        Exit Function
    End If

    ' Check for a droid
    For player = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(player).head
        Do While it <> 0
            psCurr = lnObj(it)
            If map_coord(aDroid(psCurr).pos.x) = x _
               And map_coord(aDroid(psCurr).pos.y) = y Then
                debug_ 2452, LOG_NEVER, "structClearTile", "failed - not vacant"
                structClearTile = False
                Exit Function
            End If
            it = lnNext(it)
        Loop
    Next

    debug_ 2458, LOG_NEVER, "structClearTile", "succeeded"
    structClearTile = True
End Function

' An auxiliary function for std::stable_sort in placeDroid
Private Function comparePlacementPoints(a As Vector2i, b As Vector2i) As Boolean
    Dim dist_a As Long, dist_b As Long
    ' Compare by Manhattan distance first
    dist_a = Abs(a.x) + Abs(a.y)
    dist_b = Abs(b.x) + Abs(b.y)

    If dist_a <> dist_b Then
        comparePlacementPoints = dist_a < dist_b ' Sort by Manhattan distance
    Else
        If a.x <> b.x Then
            comparePlacementPoints = a.x < b.x ' Compare x-coordinates
            Exit Function
        End If
        comparePlacementPoints = a.y < b.y ' Compare y-coordinates
    End If
End Function

' Find a location near to a structure to start the droid of
Public Function placeDroid(ByVal psStructure As Long, ByVal psTempl As Long, droidX As Long, droidY As Long) As Boolean
    Dim bounds As StructureBounds, xmin As Long, xmax As Long, ymin As Long, ymax As Long, direction As Long
    Dim sx As Long, sy As Long, tiles() As Vector2i, tiles_size As Long, x As Long, y As Long, best As Long, i As Long
    Dim wx As Long, wy As Long

    checkStructure psStructure, "src/structure.cpp:2486", "placeDroid", max_check_object_recursion

    ' Find the four corners of the square
    bounds = getStructureBounds_struct(psStructure)
    xmin = MAX_(bounds.map.x - 1, 0)
    xmax = MIN_(bounds.map.x + bounds.size.x, mapWidth)
    ymin = MAX_(bounds.map.y - 1, 0)
    ymax = MIN_(bounds.map.y + bounds.size.y, mapHeight)

    ' Round direction to nearest 90 degrees.
    direction = snapDirection(aStruct(psStructure).rot.direction)

    ' We sort all adjacent tiles by their Manhattan distance to the
    ' target droid exit tile, misplaced by (1/3, 1/4) tiles.
    ' Since only whole coordinates are sorted, this makes sure sorting
    ' is deterministic. Target coordinates, multiplied by 12 to eliminate
    ' floats, are stored in (sx, sy).

    If direction = &H0& Then
        sx = 12 * (xmin + 1) + 4
        sy = 12 * ymax + 3
    ElseIf direction = &H4000& Then
        sx = 12 * xmax + 3
        sy = 12 * (ymax - 1) - 4
    ElseIf direction = &H8000& Then
        sx = 12 * (xmax - 1) - 4
        sy = 12 * ymin - 3
    Else
        sx = 12 * xmin - 3
        sy = 12 * (ymin + 1) + 4
    End If

    ' std::vector<Vector2i> tiles
    tiles_size = 0
    For y = ymin To ymax
        For x = xmin To xmax
            If structClearTile(x, y, asPropulsionStats(DROID_TEMPLATE_getPropulsionStats(apsTemplates(psTempl))).propulsionType) Then
                ReDim Preserve tiles(0 To tiles_size)
                tiles(tiles_size) = Vector2i_(12 * x - sx, 12 * y - sy)
                tiles_size = tiles_size + 1
            End If
        Next
    Next

    If tiles_size = 0 Then
        placeDroid = False
        Exit Function
    End If

    ' std::sort(tiles.begin(), tiles.end(), comparePlacementPoints): only tiles[0], the least (the points
    ' differ, so it is unique), is used
    best = 0
    For i = 1 To tiles_size - 1
        If comparePlacementPoints(tiles(i), tiles(best)) Then best = i
    Next

    ' Store best tile coordinates in (sx, sy),
    ' which are also map coordinates of its north-west corner.
    ' Store world coordinates of this tile's center in (wx, wy)
    sx = (tiles(best).x + sx) \ 12
    sy = (tiles(best).y + sy) \ 12
    wx = world_coord(sx) + TILE_UNITS \ 2
    wy = world_coord(sy) + TILE_UNITS \ 2

    ' Finally, find world coordinates of the structure point closest to (mx, my).
    ' For simplicity, round to grid vertices.
    If 2 * sx <= xmin + xmax Then
        wx = wx + (TILE_UNITS \ 2 - 1)
    End If
    If 2 * sx >= xmin + xmax Then
        wx = wx - (TILE_UNITS \ 2 - 1)
    End If
    If 2 * sy <= ymin + ymax Then
        wy = wy + (TILE_UNITS \ 2 - 1)
    End If
    If 2 * sy >= ymin + ymax Then
        wy = wy - (TILE_UNITS \ 2 - 1)
    End If

    droidX = wx
    droidY = wy
    placeDroid = True
End Function

'Set the factory secondary orders to a droid
Public Sub setFactorySecondaryState(ByVal psDroid As Long, ByVal psStructure As Long)
    Dim newState As Long, diff As Long

    checkStructure psStructure, "src/structure.cpp:2580", "setFactorySecondaryState", max_check_object_recursion
    If Not (STRUCTURE_isFactory(psStructure)) Then ASSERT_FAILURE "psStructure->isFactory()", "src/structure.cpp:2581", "setFactorySecondaryState", "structure not a factory": Exit Sub

    If myResponsibility(aStruct(psStructure).player) Then
        newState = aStruct(psStructure).functionality.factory.secondaryOrder
        diff = newState Xor aDroid(psDroid).secondaryOrder
        If (diff And DSS_ARANGE_MASK) <> 0 Then
            secondarySetState psDroid, DSO_ATTACK_RANGE, newState And DSS_ARANGE_MASK
        End If
        If (diff And DSS_REPLEV_MASK) <> 0 Then
            secondarySetState psDroid, DSO_REPAIR_LEVEL, newState And DSS_REPLEV_MASK
        End If
        If (diff And DSS_ALEV_MASK) <> 0 Then
            secondarySetState psDroid, DSO_ATTACK_LEVEL, newState And DSS_ALEV_MASK
        End If
        If (diff And DSS_CIRCLE_MASK) <> 0 Then
            secondarySetState psDroid, DSO_CIRCLE, newState And DSS_CIRCLE_MASK
        End If
        If (diff And DSS_HALT_MASK) <> 0 Then
            secondarySetState psDroid, DSO_HALTTYPE, newState And DSS_HALT_MASK
        End If
    End If
End Sub

' Place a newly manufactured droid next to a factory  and then send if off
' to the assembly point, returns true if droid was placed successfully
Private Function structPlaceDroid(ByVal psStructure As Long, ByVal psTempl As Long, ppsDroid As Long) As Boolean
    Dim x As Long, y As Long, placed As Boolean, psNewDroid As Long, psFact As Long, factoryType As Long
    Dim assignCommander As Boolean, initialOrders As INITIAL_DROID_ORDERS, rot As Rotation, hasCommander As Boolean
    Dim psAssemblyPoint As Long, psFlag As Long, it As Long, pos As Vector2i, droidPos As Vector3i

    checkStructure psStructure, "src/structure.cpp:2621", "structPlaceDroid", max_check_object_recursion

    placed = placeDroid(psStructure, psTempl, x, y)

    If placed Then
        psAssemblyPoint = aStruct(psStructure).functionality.factory.psAssemblyPoint
        initialOrders.secondaryOrder = aStruct(psStructure).functionality.factory.secondaryOrder
        initialOrders.moveToX = aFlagPos(psAssemblyPoint).coords.x
        initialOrders.moveToY = aFlagPos(psAssemblyPoint).coords.y
        initialOrders.factoryId = aStruct(psStructure).id
        'create a droid near to the structure
        syncDebug "structPlaceDroid", "Placing new droid at (%d,%d)", x, y
        turnOffMultiMsg True
        rot = aStruct(psStructure).rot
        psNewDroid = buildDroid(psTempl, x, y, aStruct(psStructure).player, False, True, initialOrders, rot)
        turnOffMultiMsg False
        If psNewDroid = 0 Then
            ppsDroid = 0
            structPlaceDroid = False
            Exit Function
        End If
        psFact = psStructure  ' FACTORY *
        hasCommander = aStruct(psFact).functionality.factory.psCommander <> 0 And myResponsibility(aStruct(psStructure).player)
        ' assign a group to the manufactured droid
        If aStruct(psStructure).productToGroup <> UBYTE_MAX Then
            aDroid(psNewDroid).group = aStruct(psStructure).productToGroup
            If Not hasCommander Or DROID_isConstructionDroid(psNewDroid) Then    ' isConstructionDroid(psNewDroid)
                intGroupsChanged aDroid(psNewDroid).group ' update groups UI
                SelectGroupDroid psNewDroid
            End If
        End If
        setFactorySecondaryState psNewDroid, psStructure
        droidPos = aDroid(psNewDroid).pos
        displayConstructionCloud droidPos
        ' add the droid to the list
        addDroid psNewDroid, apsDroidLists
        ppsDroid = psNewDroid
        If aDroid(psNewDroid).player = selectedPlayer Then
            audio_QueueTrack ID_SOUND_DROID_COMPLETED
            intRefreshScreen ' update any interface implications.
        End If

        ' update the droid counts
        adjustDroidCount psNewDroid, 1

        ' if we've built a command droid - make sure that it isn't assigned to another commander
        assignCommander = False
        If (aDroid(psNewDroid).droidType = DROID_COMMAND) And _
           (aStruct(psFact).functionality.factory.psCommander <> 0) Then
            assignFactoryCommandDroid psStructure, 0
            assignCommander = True
        End If

        If DROID_isVtol(psNewDroid) And Not DROID_isTransporter(psNewDroid) Then
            moveToRearm psNewDroid
        End If
        If hasCommander Then
            ' TODO: Should synchronise .psCommander in all cases.
            'syncDebug("Has commander.");
            If DROID_isTransporter(psNewDroid) Then
                ' Transporters can't be assigned to commanders, due to abuse of .psGroup. Try to land on the commander instead. Hopefully the transport is heavy enough to crush the commander.
                orderDroidLoc psNewDroid, DORDER_MOVE, aDroid(aStruct(psFact).functionality.factory.psCommander).pos.x, aDroid(aStruct(psFact).functionality.factory.psCommander).pos.y, ModeQueue
            ElseIf DROID_isVtol(psNewDroid) Then
                orderDroidObj psNewDroid, DORDER_FIRESUPPORT, aStruct(psFact).functionality.factory.psCommander, ModeQueue
                'moveToRearm(psNewDroid);
            ElseIf Not DROID_isConstructionDroid(psNewDroid) Then
                orderDroidObj psNewDroid, DORDER_COMMANDERSUPPORT, aStruct(psFact).functionality.factory.psCommander, ModeQueue
            End If
        Else
            'check flag against factory type
            factoryType = FACTORY_FLAG
            If asStructureStats(aStruct(psStructure).pStructureType).typ = REF_CYBORG_FACTORY Then
                factoryType = CYBORG_FLAG
            ElseIf asStructureStats(aStruct(psStructure).pStructureType).typ = REF_VTOL_FACTORY Then
                factoryType = VTOL_FLAG
            End If
            'find flag in question.
            psFlag = 0
            it = apsFlagPosLists(aFlagPos(aStruct(psFact).functionality.factory.psAssemblyPoint).player).head
            Do While it <> 0
                If aFlagPos(lnObj(it)).factoryInc = aFlagPos(aStruct(psFact).functionality.factory.psAssemblyPoint).factoryInc _
                   And aFlagPos(lnObj(it)).factoryType = factoryType Then   ' correct fact. / correct type
                    psFlag = lnObj(it)
                    Exit Do
                End If
                it = lnNext(it)
            Loop
            If Not (psFlag <> 0) Then ASSERT_FAILURE "psFlag != flagPosList.end()", "src/structure.cpp:2715", "structPlaceDroid", "No flag found for %s at (%d, %d)", objInfo(psStructure), aStruct(psStructure).pos.x, aStruct(psStructure).pos.y
            'if vtol droid - send it to ReArm Pad if one exists
            If psFlag <> 0 And DROID_isVtol(psNewDroid) Then
                pos = Vector2i_(aFlagPos(psFlag).coords.x, aFlagPos(psFlag).coords.y)
                'find a suitable location near the delivery point
                actionVTOLLandingPos psNewDroid, pos
                orderDroidLoc psNewDroid, DORDER_MOVE, pos.x, pos.y, ModeQueue
            ElseIf psFlag <> 0 Then
                orderDroidLoc psNewDroid, DORDER_MOVE, aFlagPos(psFlag).coords.x, aFlagPos(psFlag).coords.y, ModeQueue
            End If
        End If
        If assignCommander Then
            assignFactoryCommandDroid psStructure, psNewDroid
        End If
        structPlaceDroid = True
        Exit Function
    Else
        syncDebug "structPlaceDroid", "Droid placement failed"
        ppsDroid = 0
    End If
    structPlaceDroid = False
End Function

' psFactory: a FACTORY * (its structure's handle)
Private Function IsFactoryCommanderGroupFull(ByVal psFactory As Long) As Boolean
    Dim DroidsInGroup As Long, psCommander As Long

    If bMultiPlayer Then
        ' TODO: Synchronise .psCommander. Have to return false here, to avoid desynch.
        IsFactoryCommanderGroupFull = False
        Exit Function
    End If

    ' If we don't have a commander return false (group not full)
    If aStruct(psFactory).functionality.factory.psCommander = 0 Then
        IsFactoryCommanderGroupFull = False
        Exit Function
    End If

    ' allow any number of IDF droids
    If templateIsIDF(aStruct(psFactory).functionality.factory.psSubject) Then
        IsFactoryCommanderGroupFull = False
        Exit Function
    End If
    If asPropulsionStats(DROID_TEMPLATE_getPropulsionStats(apsTemplates(aStruct(psFactory).functionality.factory.psSubject))).propulsionType = PROPULSION_TYPE_LIFT Then
        IsFactoryCommanderGroupFull = False
        Exit Function
    End If

    ' Get the number of droids in the commanders group
    psCommander = aStruct(psFactory).functionality.factory.psCommander
    If aDroid(psCommander).psGroup <> 0 Then DroidsInGroup = DROID_GROUP_getNumMembers(aDroid(psCommander).psGroup) Else DroidsInGroup = 0

    ' if the number in group is less than the maximum allowed then return false (group not full)
    If ULT(DroidsInGroup, cmdDroidMaxGroup(psCommander)) Then
        IsFactoryCommanderGroupFull = False
        Exit Function
    End If

    ' the number in group has reached the maximum
    IsFactoryCommanderGroupFull = True
End Function

' Check if a player has a certain structure. Optionally, checks if there is
' at least one that is built.
Public Function structureExists(ByVal player As Long, ByVal typ As STRUCTURE_TYPE, ByVal built As Boolean, ByVal isMission As Boolean) As Boolean
    Dim found_ As Boolean, it As Long, psCurr As Long

    found_ = False

    If Not (player >= 0) Then ASSERT_FAILURE "player >= 0", "src/structure.cpp:2785", "structureExists", "invalid player: %d", player: structureExists = False: Exit Function
    If player >= MAX_PLAYERS Then
        structureExists = False
        Exit Function
    End If

    ' StructureList* pList = isMission ? &mission.apsStructLists[player] : &apsStructLists[player]
    If isMission Then it = mission_apsStructLists(player).head Else it = apsStructLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If asStructureStats(aStruct(psCurr).pStructureType).typ = typ And (Not built Or (built And aStruct(psCurr).status = SS_BUILT)) Then
            found_ = True
            Exit Do
        End If
        it = lnNext(it)
    Loop

    structureExists = found_
End Function

' Disallow manufacture of units once these limits are reached,
' doesn't mean that these numbers can't be exceeded if units are
' put down in the editor or by the scripts.

Public Sub setMaxDroids(ByVal player As Long, ByVal value As Long)
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:2810", "setMaxDroids", "player = %u", player: Exit Sub
    droidLimit(player) = value
End Sub

Public Sub setMaxCommanders(ByVal player As Long, ByVal value As Long)
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:2816", "setMaxCommanders", "player = %u", player: Exit Sub
    commanderLimit(player) = value
End Sub

Public Sub setMaxConstructors(ByVal player As Long, ByVal value As Long)
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:2822", "setMaxConstructors", "player = %u", player: Exit Sub
    constructorLimit(player) = value
End Sub

Public Function getMaxDroids(ByVal player As Long) As Long
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:2828", "getMaxDroids", "player = %u", player: getMaxDroids = 0: Exit Function
    getMaxDroids = droidLimit(player)
End Function

Public Function getMaxCommanders(ByVal player As Long) As Long
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:2834", "getMaxCommanders", "player = %u", player: getMaxCommanders = 0: Exit Function
    getMaxCommanders = commanderLimit(player)
End Function

Public Function getMaxConstructors(ByVal player As Long) As Long
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:2840", "getMaxConstructors", "player = %u", player: getMaxConstructors = 0: Exit Function
    getMaxConstructors = constructorLimit(player)
End Function

Public Function IsPlayerDroidLimitReached(ByVal player As Long) As Boolean
    Dim numDroids As Long
    numDroids = getNumDroids(player) + getNumMissionDroids(player) + getNumTransporterDroids(player)
    IsPlayerDroidLimitReached = numDroids >= getMaxDroids(player)
End Function

' Check for max number of units reached and halt production.
Private Function checkHaltOnMaxUnitsReached(ByVal psStructure As Long, ByVal isMission As Boolean) As Boolean
    Dim limitMsg As String, isLimit As Boolean, player As Long, templ As Long

    checkStructure psStructure, "src/structure.cpp:2853", "checkHaltOnMaxUnitsReached", max_check_object_recursion

    isLimit = False
    player = aStruct(psStructure).player

    templ = aStruct(psStructure).functionality.factory.psSubject

    ' if the players that owns the factory has reached his (or hers) droid limit
    ' then put production on hold & return - we need a message to be displayed here !!!!!!!
    If IsPlayerDroidLimitReached(player) Then
        isLimit = True
        limitMsg = "Can't build any more units, Unit Limit Reached " & ChrW$(&H2014) & " Production Halted"
    Else
        Select Case droidTemplateType(apsTemplates(templ))
        Case DROID_COMMAND
            If Not structureExists(player, REF_COMMAND_CONTROL, True, isMission) Then
                isLimit = True
                limitMsg = va("Can't build ""%s"" without a Command Relay Center " & ChrW$(&H2014) & " Production Halted", apsTemplates(templ).name_)
            ElseIf Not ULT(getNumCommandDroids(player), getMaxCommanders(player)) Then
                isLimit = True
                limitMsg = va("Can't build ""%s"", Commander Limit Reached " & ChrW$(&H2014) & " Production Halted", apsTemplates(templ).name_)
            End If
        Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
            If Not ULT(getNumConstructorDroids(player), getMaxConstructors(player)) Then
                isLimit = True
                limitMsg = va("Can't build any more ""%s"", Construction Unit Limit Reached " & ChrW$(&H2014) & " Production Halted", apsTemplates(templ).name_)
            End If
        Case Else
        End Select
    End If

    If isLimit And player = selectedPlayer Then
        If lastMaxUnitMessage = 0 Or Not ULT(gameTime, FromU32(U32(lastMaxUnitMessage) + MAX_UNIT_MESSAGE_PAUSE)) Then
            addConsoleMessage limitMsg, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
            lastMaxUnitMessage = gameTime
        End If
    End If

    checkHaltOnMaxUnitsReached = isLimit
End Function

' station: the repair facility (STRUCTURE &)
Public Sub aiUpdateRepair_handleState(ByVal station As Long)
    Dim psStructure As Long, psRepairFac As Long, psDroid As Long, iVecEffect As Vector3i

    psStructure = station
    psRepairFac = station  ' const REPAIR_FACILITY *
    psDroid = aStruct(psRepairFac).functionality.repairFacility.psObj
    ' apply logic for current state
    Select Case aStruct(psRepairFac).functionality.repairFacility.state
    Case RepairState_Idle
        actionAlignTurret psStructure, 0
        Exit Sub
    Case RepairState_Repairing
        If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/structure.cpp:2918", "aiUpdateRepair_handleState", "invalid droid pointer": Exit Sub
        aDroid(psDroid).body = FromU32(U32(aDroid(psDroid).body) + gameTimeAdjustedAverage(getBuildingRepairPoints(psStructure)))
        If ULT(aDroid(psDroid).originalBody, aDroid(psDroid).body) Then aDroid(psDroid).body = aDroid(psDroid).originalBody   ' MIN (unsigned)
        actionTargetTurret psStructure, psDroid, 0
        If BASE_OBJECT_visibleForLocalDisplay(psStructure) <> 0 And BASE_OBJECT_visibleForLocalDisplay(psDroid) <> 0 Then ' display only - does not impact simulation state
            ' add plasma repair effect whilst being repaired
            iVecEffect.x = aDroid(psDroid).pos.x + (10 - rand() Mod 20)
            iVecEffect.y = aDroid(psDroid).pos.z + (10 - rand() Mod 20)
            iVecEffect.z = aDroid(psDroid).pos.y + (10 - rand() Mod 20)
            effectSetSize 100
            addEffect iVecEffect, EFFECT_EXPLOSION, EXPLOSION_TYPE_SPECIFIED, True, getDisplayImdFromIndex(MI_FLAME), 0, gameTime - deltaGameTime + 1
        End If
        Exit Sub
    Case RepairState_Invalid
        debug_ 2936, LOG_ERROR, "aiUpdateRepair_handleState", "repair state is invalid"
        Exit Sub
    End Select
End Sub

' station: the repair facility (const STRUCTURE &); psDroidOut: DROID **
Public Function aiUpdateRepair_obtainEvents(ByVal station As Long, psDroidOut As Long) As RepairEvents
    Dim psStructure As Long, psRepairFac As Long, xdiff As Long, ydiff As Long, psDroid As Long

    psStructure = station
    psRepairFac = station  ' const REPAIR_FACILITY *
    Select Case aStruct(psRepairFac).functionality.repairFacility.state
    Case RepairState_Idle
        psDroid = findSomeoneToRepair(psStructure, (TILE_UNITS * 5 \ 2))
        psDroidOut = psDroid
        If psDroid <> 0 Then aiUpdateRepair_obtainEvents = RepairEvents_RepairTargetFound Else aiUpdateRepair_obtainEvents = RepairEvents_NoEvents
        Exit Function
    Case RepairState_Repairing
        psDroid = aStruct(psRepairFac).functionality.repairFacility.psObj
        If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid", "src/structure.cpp:2958", "aiUpdateRepair_obtainEvents", "repairing a null object??"
        If aDroid(psDroid).died <> 0 Then aiUpdateRepair_obtainEvents = RepairEvents_UnitDied: Exit Function

        xdiff = aDroid(psDroid).pos.x - aStruct(psStructure).pos.x
        ydiff = aDroid(psDroid).pos.y - aStruct(psStructure).pos.y
        If FromU32(CDbl(xdiff) * xdiff + CDbl(ydiff) * ydiff) > (TILE_UNITS * 5 \ 2) * (TILE_UNITS * 5 \ 2) Then
            aiUpdateRepair_obtainEvents = RepairEvents_UnitMovedAway
            Exit Function
        End If
        If Not ULT(aDroid(psDroid).body, aDroid(psDroid).originalBody) Then
            aiUpdateRepair_obtainEvents = RepairEvents_UnitReachedMaxHP
            Exit Function
        End If
        aiUpdateRepair_obtainEvents = RepairEvents_NoEvents
        Exit Function
    Case RepairState_Invalid
        debug_ 2975, LOG_ERROR, "aiUpdateRepair_obtainEvents", "Inavlid repair station (id=%i player=%i) state", aStruct(station).id, aStruct(station).player
        aiUpdateRepair_obtainEvents = RepairEvents_NoEvents
        Exit Function
    Case Else
        debug_ 2980, LOG_ERROR, "aiUpdateRepair_obtainEvents", "Unknown repair station (id=%i player=%i) state", aStruct(station).id, aStruct(station).player
        aiUpdateRepair_obtainEvents = RepairEvents_NoEvents
        Exit Function
    End Select
End Function

' station: the repair facility (STRUCTURE &)
Public Function aiUpdateRepair_handleEvents(ByVal station As Long, ByVal ev As RepairEvents, ByVal found_ As Long) As RepairState
    Dim psRepairFac As Long, psStructure As Long, psDroid As Long

    psRepairFac = station  ' REPAIR_FACILITY *
    psStructure = station
    'don't do anything if the resistance is low in multiplayer
    If bMultiPlayer And aStruct(psStructure).resistance < structureResistance(aStruct(psStructure).pStructureType, aStruct(psStructure).player) Then
        objTrace aStruct(psStructure).id, "aiUpdateRepair_handleEvents", "Resistance too low for repair"
        If aStruct(psRepairFac).functionality.repairFacility.psObj <> 0 Then
            droidRepairStopped castDroid(aStruct(psRepairFac).functionality.repairFacility.psObj), station
            aStruct(psRepairFac).functionality.repairFacility.psObj = 0
        End If
        aiUpdateRepair_handleEvents = RepairState_Idle
        Exit Function
    End If
    Select Case ev
    Case RepairEvents_NoEvents
        aiUpdateRepair_handleEvents = aStruct(psRepairFac).functionality.repairFacility.state
        Exit Function
    Case RepairEvents_RepairTargetFound
        If Not (found_ <> 0) Then ASSERT_FAILURE "found != nullptr", "src/structure.cpp:3009", "aiUpdateRepair_handleEvents", "Bug! found droid, but it was null?"
        aStruct(psRepairFac).functionality.repairFacility.psObj = found_
        droidRepairStarted found_, station
        aiUpdateRepair_handleEvents = RepairState_Repairing
        Exit Function
    Case RepairEvents_UnitReachedMaxHP
        psDroid = aStruct(psRepairFac).functionality.repairFacility.psObj
        aStruct(psRepairFac).functionality.repairFacility.psObj = 0
        objTrace aStruct(psStructure).id, "aiUpdateRepair_handleEvents", "Repair complete of droid %d", aDroid(psDroid).id
        aDroid(psDroid).body = aDroid(psDroid).originalBody
        ' if completely repaired reset order
        objTrace aDroid(psDroid).id, "aiUpdateRepair_handleEvents", "was fully repaired by RF"
        droidRepairStopped psDroid, station
        droidWasFullyRepaired psDroid, psRepairFac
        ' only call "secondarySetState" *after* triggering "droidWasFullyRepaired"
        ' because in some cases calling it would modify primary order
        ' thus, loosing information that we actually had a RTR|RTR_SPECIFIED before
        secondarySetState psDroid, DSO_RETURN_TO_LOC, DSS_NONE
        aiUpdateRepair_handleEvents = RepairState_Idle
        Exit Function
    Case RepairEvents_UnitDied
        psDroid = aStruct(psRepairFac).functionality.repairFacility.psObj
        syncDebugDroid_ "aiUpdateRepair_handleEvents", psDroid, Asc("-")
        droidRepairStopped psDroid, station
        aStruct(psRepairFac).functionality.repairFacility.psObj = 0
        aiUpdateRepair_handleEvents = RepairState_Idle
        Exit Function
    Case RepairEvents_UnitMovedAway
        droidRepairStopped castDroid(aStruct(psRepairFac).functionality.repairFacility.psObj), station
        aStruct(psRepairFac).functionality.repairFacility.psObj = 0
        aiUpdateRepair_handleEvents = RepairState_Idle
        Exit Function
    End Select
    debug_ 3045, LOG_ERROR, "aiUpdateRepair_handleEvents", "Unknown event passed"
    aiUpdateRepair_handleEvents = RepairState_Invalid
End Function

' station: the repair facility (STRUCTURE &)
Public Sub aiUpdateRepairStation(ByVal station As Long)
    Dim found_ As Long, topEvent As RepairEvents, nextState As RepairState

    found_ = 0
    topEvent = aiUpdateRepair_obtainEvents(station, found_)
    nextState = aiUpdateRepair_handleEvents(station, topEvent, found_)
    If Not (nextState <> RepairState_Invalid) Then ASSERT_FAILURE "nextState != RepairState::Invalid", "src/structure.cpp:3054", "aiUpdateRepairStation", "Bug! invalid state received."
    aStruct(station).functionality.repairFacility.state = nextState
    aiUpdateRepair_handleState station
End Sub

Private Sub aiUpdateStructure(ByVal psStructure As Long, ByVal isMission As Boolean)
    Dim structureMode As Long, psDroid As Long, psChosenObjs(0 To MAX_WEAPONS - 1) As Long, psChosenObj As Long
    Dim psFactory As Long, bDroidPlaced As Boolean, psWStats As Long, bDirect As Boolean, tmpOrigin As TARGET_ORIGIN
    Dim count1 As Long, count2 As Long, it As Long, i As Long, c As Boolean, pSubject As Long, hasSubject As Boolean
    Dim psReArmPad As Long, psCurr As Long, psResFacility As Long, researchIndex As Long, player As Long, pResearch As Long
    Dim pointsToAdd As Long, d As Long, shareProgress As Long, shareIsFinished As Boolean, haveEnoughPower As Boolean
    Dim prevState As Long, progress As Long, pointsAlreadyAdded As Long, ammoToAddThisTime As Long, rearmPoints As Long

    structureMode = 0
    psChosenObj = 0
    bDroidPlaced = False
    bDirect = False
    tmpOrigin = ORIGIN_UNKNOWN

    checkStructure psStructure, "src/structure.cpp:3071", "aiUpdateStructure", max_check_object_recursion

    If aStruct(psStructure).time_ = gameTime Then
        ' This isn't supposed to happen, and really shouldn't be possible - if this happens, maybe a structure is being updated twice?
        count1 = 0
        count2 = 0
        it = apsStructLists(aStruct(psStructure).player).head
        Do While it <> 0
            If lnObj(it) = psStructure Then count1 = count1 + 1
            it = lnNext(it)
        Loop
        it = mission_apsStructLists(aStruct(psStructure).player).head
        Do While it <> 0
            If lnObj(it) = psStructure Then count2 = count2 + 1
            it = lnNext(it)
        Loop
        debug_ 3085, LOG_ERROR, "aiUpdateStructure", "psStructure->prevTime = %u, psStructure->time = %u, gameTime = %u, count1 = %d, count2 = %d", aStruct(psStructure).prevTime, aStruct(psStructure).time_, gameTime, count1, count2
        aStruct(psStructure).time_ = aStruct(psStructure).time_ - 1
    End If
    aStruct(psStructure).prevTime = aStruct(psStructure).time_
    aStruct(psStructure).time_ = gameTime
    For i = 0 To MAX_(1, aStruct(psStructure).numWeaps) - 1
        aStruct(psStructure).asWeaps(i).prevRot = aStruct(psStructure).asWeaps(i).rot
    Next

    If isMission Then
        Select Case asStructureStats(aStruct(psStructure).pStructureType).typ
        Case REF_RESEARCH, REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        Case Else
            Exit Sub ' nothing to do
        End Select
    End If

    ' Will go out into a building EVENT stats/text file
    ' Spin round yer sensors!
    If aStruct(psStructure).numWeaps = 0 Then
        If (aStruct(psStructure).asWeaps(0).nStat = 0) And _
           (asStructureStats(aStruct(psStructure).pStructureType).typ <> REF_REPAIR_FACILITY) Then

            '//////
            ' - radar should rotate every three seconds ... 'cause we timed it at Heathrow !
            ' gameTime is in milliseconds - one rotation every 3 seconds = 1 rotation event 3000 millisecs
            aStruct(psStructure).asWeaps(0).rot.direction = ((FromU32(Int(U32(gameTime) * 65536# / 3000#)) And &HFFFF&) + ((aStruct(psStructure).pos.x + aStruct(psStructure).pos.y) Mod 10) * 6550) And &HFFFF&  ' Randomize by hashing position as seed for rotating 1/10th turns. Cast wrapping intended.
            aStruct(psStructure).asWeaps(0).rot.pitch = 0
        End If
    End If

    ' Check lassat
    c = isLasSat(aStruct(psStructure).pStructureType)
    If c Then c = ULT(weaponFirePause(asWeaponStats(BASE_OBJECT_getWeaponStats(psStructure, 0)), aStruct(psStructure).player), FromU32(U32(gameTime) - U32(aStruct(psStructure).asWeaps(0).lastFired)))
    If c Then c = aStruct(psStructure).asWeaps(0).ammo <> 0
    If c Then
        If aStruct(psStructure).player = selectedPlayer Then
            addConsoleMessage "Laser Satellite is ready to fire!", CENTRE_JUSTIFY, SYSTEM_MESSAGE
            audio_PlayTrack ID_SOUND_UPLINK ' This is about the best sound available for this.
        End If
        triggerEventStructureReady psStructure
        aStruct(psStructure).asWeaps(0).ammo = 0 ' do not fire more than once
    End If

    ' See if there is an enemy to attack
    If aStruct(psStructure).numWeaps > 0 Then
        'structures always update their targets
        For i = 0 To aStruct(psStructure).numWeaps - 1
            bDirect = proj_Direct(BASE_OBJECT_getWeaponStats(psStructure, i))
            c = aStruct(psStructure).asWeaps(i).nStat <> 0  ' nStat > 0 (unsigned)
            If c Then c = asWeaponStats(BASE_OBJECT_getWeaponStats(psStructure, i)).weaponSubClass <> WSC_LAS_SAT
            If c Then
                If aiChooseTarget(psStructure, psChosenObjs(i), i, True, True, tmpOrigin) Then
                    objTrace aStruct(psStructure).id, "aiUpdateStructure", "Weapon %d is targeting %d at (%d, %d)", i, SIMPLE_OBJECT_id(psChosenObjs(i)), _
                             SIMPLE_OBJECT_pos(psChosenObjs(i)).x, SIMPLE_OBJECT_pos(psChosenObjs(i)).y
                    setStructureTarget psStructure, psChosenObjs(i), i, tmpOrigin
                Else
                    If i > 0 Then
                        If psChosenObjs(0) <> 0 Then
                            objTrace aStruct(psStructure).id, "aiUpdateStructure", "Weapon %d is supporting main weapon: %d at (%d, %d)", i, _
                                     SIMPLE_OBJECT_id(psChosenObjs(0)), SIMPLE_OBJECT_pos(psChosenObjs(0)).x, SIMPLE_OBJECT_pos(psChosenObjs(0)).y
                            setStructureTarget psStructure, psChosenObjs(0), i, tmpOrigin
                            psChosenObjs(i) = psChosenObjs(0)
                        Else
                            setStructureTarget psStructure, 0, i, ORIGIN_UNKNOWN
                            psChosenObjs(i) = 0
                        End If
                    Else
                        setStructureTarget psStructure, 0, i, ORIGIN_UNKNOWN
                        psChosenObjs(i) = 0
                    End If
                End If

                c = psChosenObjs(i) <> 0
                If c Then c = Not aiObjectIsProbablyDoomed(psChosenObjs(i), bDirect)
                If c Then
                    ' get the weapon stat to see if there is a visible turret to rotate
                    psWStats = BASE_OBJECT_getWeaponStats(psStructure, i)

                    'if were going to shoot at something move the turret first then fire when locked on
                    If asWeaponStats(psWStats).pMountGraphic = 0 Then 'no turret so lock on whatever
                        aStruct(psStructure).asWeaps(i).rot.direction = calcDirection(aStruct(psStructure).pos.x, aStruct(psStructure).pos.y, SIMPLE_OBJECT_pos(psChosenObjs(i)).x, SIMPLE_OBJECT_pos(psChosenObjs(i)).y)
                        combFire aStruct(psStructure).asWeaps(i), psStructure, psChosenObjs(i), i
                    ElseIf actionTargetTurret(psStructure, psChosenObjs(i), i) Then
                        combFire aStruct(psStructure).asWeaps(i), psStructure, psChosenObjs(i), i
                    End If
                Else
                    ' realign the turret
                    If (aStruct(psStructure).asWeaps(i).rot.direction Mod DEG(90)) <> 0 Or aStruct(psStructure).asWeaps(i).rot.pitch <> 0 Then
                        actionAlignTurret psStructure, i
                    End If
                End If
            End If
        Next

    ' See if there is an enemy to attack for Sensor Towers that have weapon droids attached
    ElseIf asStructureStats(aStruct(psStructure).pStructureType).pSensor <> -1 Then
        c = structStandardSensor(psStructure)
        If Not c Then c = structVTOLSensor(psStructure)
        If Not c Then c = objRadarDetector(psStructure)
        If c Then
            If aiChooseSensorTarget(psStructure, psChosenObj) Then
                objTrace aStruct(psStructure).id, "aiUpdateStructure", "Sensing (%d)", SIMPLE_OBJECT_id(psChosenObj)
                If objRadarDetector(psStructure) Then
                    setStructureTarget psStructure, psChosenObj, 0, ORIGIN_RADAR_DETECTOR
                Else
                    setStructureTarget psStructure, psChosenObj, 0, ORIGIN_SENSOR
                End If
            Else
                setStructureTarget psStructure, 0, 0, ORIGIN_UNKNOWN
            End If
            psChosenObj = aStruct(psStructure).psTarget(0)
        Else
            psChosenObj = aStruct(psStructure).psTarget(0)
        End If
    End If
    'only interested if the Structure "does" something!
    If aStruct(psStructure).pFunctionality = 0 Then
        Exit Sub
    End If

    ' Process the functionality according to type
    ' determine the subject stats (for research or manufacture)
    ' or base object (for repair) or update power levels for resourceExtractor
    ' (VB) BASE_STATS *pSubject: a RESEARCH * (-1 = nullptr) or a DROID_TEMPLATE * (0 = nullptr); hasSubject = pSubject != nullptr
    pSubject = 0
    hasSubject = False
    Select Case asStructureStats(aStruct(psStructure).pStructureType).typ
    Case REF_RESEARCH
        pSubject = aStruct(psStructure).functionality.researchFacility.psSubject
        hasSubject = pSubject <> -1
        structureMode = REF_RESEARCH
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        pSubject = aStruct(psStructure).functionality.factory.psSubject
        hasSubject = pSubject <> 0
        structureMode = REF_FACTORY
        'check here to see if the factory's commander has died
        If aStruct(psStructure).functionality.factory.psCommander <> 0 Then
            If aDroid(aStruct(psStructure).functionality.factory.psCommander).died <> 0 Then
                'remove the commander from the factory
                syncDebugDroid_ "aiUpdateStructure", aStruct(psStructure).functionality.factory.psCommander, Asc("-")
                assignFactoryCommandDroid psStructure, 0
            End If
        End If
    Case REF_REPAIR_FACILITY
        aiUpdateRepairStation psStructure
    Case REF_REARM_PAD
        psReArmPad = psStructure  ' REARM_PAD *

        psChosenObj = aStruct(psReArmPad).functionality.rearmPad.psObj
        structureMode = REF_REARM_PAD
        psDroid = 0

        ' select next droid if none being rearmed
        If psChosenObj = 0 Then
            objTrace aStruct(psStructure).id, "aiUpdateStructure", "Rearm pad idle - look for victim"
            it = apsDroidLists(aStruct(psStructure).player).head
            Do While it <> 0
                psCurr = lnObj(it)
                ' move next droid waiting on ground to rearm pad
                If vtolReadyToRearm(psCurr, psStructure) Then
                    c = psChosenObj = 0
                    If Not c Then c = ULT(aDroid(psCurr).actionStarted, aDroid(psChosenObj).actionStarted)
                    If c Then
                        objTrace aDroid(psCurr).id, "aiUpdateStructure", "rearm pad candidate"
                        objTrace aStruct(psStructure).id, "aiUpdateStructure", "we found %s to rearm", objInfo(psCurr)
                        psChosenObj = psCurr
                    End If
                End If
                it = lnNext(it)
            Loop
            ' None available? Try allies.
            For i = 0 To MAX_PLAYERS - 1
                If psChosenObj <> 0 Then Exit For
                If aiCheckAlliances(i, aStruct(psStructure).player) And i <> aStruct(psStructure).player Then
                    it = apsDroidLists(i).head
                    Do While it <> 0
                        psCurr = lnObj(it)
                        ' move next droid waiting on ground to rearm pad
                        If vtolReadyToRearm(psCurr, psStructure) Then
                            psChosenObj = psCurr
                            objTrace aDroid(psCurr).id, "aiUpdateStructure", "allied rearm pad candidate"
                            objTrace aStruct(psStructure).id, "aiUpdateStructure", "we found allied %s to rearm", objInfo(psDroid)
                            Exit Do
                        End If
                        it = lnNext(it)
                    Loop
                End If
            Next
            psDroid = psChosenObj
            If psDroid <> 0 Then
                actionDroid_obj psDroid, DACTION_MOVETOREARMPOINT, psStructure
            End If
        Else
            psDroid = psChosenObj
            If (aDroid(psDroid).sMove.Status = MOVEINACTIVE Or _
                aDroid(psDroid).sMove.Status = MOVEHOVER) And _
               aDroid(psDroid).action = DACTION_WAITFORREARM Then
                objTrace aDroid(psDroid).id, "aiUpdateStructure", "supposed to go to rearm but not on our way -- fixing" ' this should never happen...
                actionDroid_obj psDroid, DACTION_MOVETOREARMPOINT, psStructure
            End If
        End If

        ' if found a droid to rearm assign it to the rearm pad
        If psDroid <> 0 Then
            ' set chosen object
            psChosenObj = psDroid
            aStruct(psReArmPad).functionality.rearmPad.psObj = psChosenObj
            If aDroid(psDroid).action = DACTION_MOVETOREARMPOINT Then
                ' reset rearm started
                aStruct(psReArmPad).functionality.rearmPad.timeStarted = ACTION_START_TIME
                aStruct(psReArmPad).functionality.rearmPad.timeLastUpdated = 0
            End If
            auxStructureBlocking psStructure
        Else
            auxStructureNonblocking psStructure
        End If
    Case Else
    End Select

    If structureMode = REF_RESEARCH Then
        incrementMultiStatsResearchPotential aStruct(psStructure).player
    End If

    ' check subject stats (for research or manufacture)
    If hasSubject Then
        'if subject is research...
        If structureMode = REF_RESEARCH Then
            psResFacility = psStructure  ' RESEARCH_FACILITY *

            'if on hold don't do anything
            If aStruct(psResFacility).functionality.researchFacility.timeStartHold <> 0 Then
                delPowerRequest psStructure
                Exit Sub
            End If

            'electronic warfare affects the functionality of some structures in multiPlayer
            If bMultiPlayer And aStruct(psStructure).resistance < structureResistance(aStruct(psStructure).pStructureType, aStruct(psStructure).player) Then
                Exit Sub
            End If

            incrementMultiStatsResearchPerformance aStruct(psStructure).player

            researchIndex = asResearch(pSubject).ref - STAT_RESEARCH

            ' PLAYER_RESEARCH *pPlayerRes = &asPlayerResList[psStructure->player][researchIndex]
            player = aStruct(psStructure).player
            'check research has not already been completed by another structure
            If Not IsResearchCompleted(asPlayerResList(player, researchIndex)) Then
                pResearch = pSubject

                pointsToAdd = gameTimeAdjustedAverage(getBuildingResearchPoints(psStructure))
                d = FromU32(asResearch(pResearch).researchPoints - U32(asPlayerResList(player, researchIndex).currentPoints))
                If Not ULT(pointsToAdd, d) Then pointsToAdd = d     ' MIN (unsigned)

                shareProgress = asPlayerResList(player, researchIndex).currentPoints  ' Share old research progress instead of new one, so it doesn't get sped up by multiple players researching.
                shareIsFinished = False

                If pointsToAdd <> 0 And asPlayerResList(player, researchIndex).currentPoints = 0 Then
                    haveEnoughPower = requestPowerFor(psStructure, asResearch(pResearch).researchPower)
                    If haveEnoughPower Then
                        shareProgress = 1  ' Share research payment, to avoid double payment even if starting research in the same game tick.
                    Else
                        pointsToAdd = 0
                    End If
                End If

                If pointsToAdd <> 0 And asResearch(pResearch).researchPoints > 0 Then ' might be a "free" research
                    asPlayerResList(player, researchIndex).currentPoints = FromU32(U32(asPlayerResList(player, researchIndex).currentPoints) + U32(pointsToAdd))
                End If
                syncDebug "aiUpdateStructure", "Research at %u/%u.", asPlayerResList(player, researchIndex).currentPoints, asResearch(pResearch).researchPoints

                'check if Research is complete
                If Not ULT(asPlayerResList(player, researchIndex).currentPoints, asResearch(pResearch).researchPoints) Then
                    prevState = intGetResearchState()

                    'store the last topic researched - if its the best
                    If aStruct(psResFacility).functionality.researchFacility.psBestTopic = -1 Then
                        aStruct(psResFacility).functionality.researchFacility.psBestTopic = aStruct(psResFacility).functionality.researchFacility.psSubject
                    Else
                        If asResearch(pResearch).researchPoints > asResearch(aStruct(psResFacility).functionality.researchFacility.psBestTopic).researchPoints Then
                            aStruct(psResFacility).functionality.researchFacility.psBestTopic = aStruct(psResFacility).functionality.researchFacility.psSubject
                        End If
                    End If
                    aStruct(psResFacility).functionality.researchFacility.psSubject = -1
                    intResearchFinished psStructure
                    researchResult researchIndex, aStruct(psStructure).player, True, psStructure, True

                    shareIsFinished = True

                    'check if this result has enabled another topic
                    intNotifyResearchButton prevState
                End If

                ' Update allies research accordingly
                If game.typ = LEVEL_TYPE_SKIRMISH And alliancesSharedResearch(game.alliance) Then
                    For i = 0 To MAX_PLAYERS - 1
                        If alliances(i, aStruct(psStructure).player) = ALLIANCE_FORMED Then
                            If Not IsResearchCompleted(asPlayerResList(i, researchIndex)) Then
                                ' Share the research for that player.
                                ' auto &allyProgress = asPlayerResList[i][researchIndex].currentPoints
                                If ULT(asPlayerResList(i, researchIndex).currentPoints, shareProgress) Then asPlayerResList(i, researchIndex).currentPoints = shareProgress   ' std::max
                                If shareIsFinished Then
                                    researchResult researchIndex, i, False, 0, True
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                'cancel this Structure's research since now complete
                aStruct(psResFacility).functionality.researchFacility.psSubject = -1
                intResearchFinished psStructure
                syncDebug "aiUpdateStructure", "Research completed elsewhere."
            End If
        'check for manufacture
        ElseIf structureMode = REF_FACTORY Then
            psFactory = psStructure  ' FACTORY *

            'if on hold don't do anything
            If aStruct(psFactory).functionality.factory.timeStartHold <> 0 Then
                Exit Sub
            End If

            'electronic warfare affects the functionality of some structures in multiPlayer
            If bMultiPlayer And aStruct(psStructure).resistance < structureResistance(aStruct(psStructure).pStructureType, aStruct(psStructure).player) Then
                Exit Sub
            End If

            If aStruct(psFactory).functionality.factory.timeStarted = ACTION_START_TIME Then
                ' also need to check if a command droid's group is full

                ' If the factory commanders group is full - return
                If IsFactoryCommanderGroupFull(psFactory) Then
                    Exit Sub
                End If
                If checkHaltOnMaxUnitsReached(psStructure, isMission) Then
                    Exit Sub
                End If

                'set the time started
                aStruct(psFactory).functionality.factory.timeStarted = gameTime
            End If

            If aStruct(psFactory).functionality.factory.buildPointsRemaining > 0 Then
                progress = gameTimeAdjustedAverage(getBuildingProductionPoints(psStructure))
                If aStruct(psFactory).functionality.factory.buildPointsRemaining = calcTemplateBuild(apsTemplates(aStruct(psFactory).functionality.factory.psSubject)) And progress > 0 Then
                    ' We're just starting to build, check for power.
                    haveEnoughPower = requestPowerFor(psStructure, calcTemplatePower(apsTemplates(aStruct(psFactory).functionality.factory.psSubject)))
                    If Not haveEnoughPower Then
                        progress = 0
                    End If
                End If
                aStruct(psFactory).functionality.factory.buildPointsRemaining = aStruct(psFactory).functionality.factory.buildPointsRemaining - progress
            End If

            'check for manufacture to be complete
            c = aStruct(psFactory).functionality.factory.buildPointsRemaining <= 0
            If c Then c = Not IsFactoryCommanderGroupFull(psFactory)
            If c Then c = Not checkHaltOnMaxUnitsReached(psStructure, isMission)
            If c Then
                If isMission Then
                    ' put it in the mission list
                    psDroid = buildMissionDroid(pSubject, _
                                                aStruct(psStructure).pos.x, aStruct(psStructure).pos.y, _
                                                aStruct(psStructure).player)
                    If psDroid <> 0 Then
                        aDroid(psDroid).secondaryOrder = aStruct(psFactory).functionality.factory.secondaryOrder
                        aDroid(psDroid).secondaryOrderPending = aDroid(psDroid).secondaryOrder
                        setFactorySecondaryState psDroid, psStructure
                        setDroidBase psDroid, psStructure
                        bDroidPlaced = True
                    End If
                Else
                    ' place it on the map
                    bDroidPlaced = structPlaceDroid(psStructure, pSubject, psDroid)
                End If

                'script callback, must be called after factory was flagged as idle
                If bDroidPlaced Then
                    'reset the start time
                    aStruct(psFactory).functionality.factory.timeStarted = ACTION_START_TIME
                    aStruct(psFactory).functionality.factory.psSubject = 0

                    doNextProduction psStructure, pSubject, ModeImmediate

                    cbNewDroid psStructure, psDroid
                End If
            End If
        End If
    End If

    ' check base object (for rearm)
    If psChosenObj <> 0 Then
        If structureMode = REF_REARM_PAD Then
            'check for rearming
            psReArmPad = psStructure  ' REARM_PAD *

            psDroid = psChosenObj
            If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/structure.cpp:3570", "aiUpdateStructure", "invalid droid pointer": Exit Sub
            If Not (DROID_isVtol(psDroid)) Then ASSERT_FAILURE "psDroid->isVtol()", "src/structure.cpp:3571", "aiUpdateStructure", "invalid droid type": Exit Sub

            'check hasn't died whilst waiting to be rearmed
            ' also clear out any previously repaired droid
            If aDroid(psDroid).died <> 0 Or (aDroid(psDroid).action <> DACTION_MOVETOREARMPOINT And aDroid(psDroid).action <> DACTION_WAITDURINGREARM) Then
                syncDebugDroid_ "aiUpdateStructure", psDroid, Asc("-")
                aStruct(psReArmPad).functionality.rearmPad.psObj = 0
                objTrace aDroid(psDroid).id, "aiUpdateStructure", "VTOL has wrong action or is dead"
                Exit Sub
            End If
            If aDroid(psDroid).action = DACTION_WAITDURINGREARM And aDroid(psDroid).sMove.Status = MOVEINACTIVE Then
                If aStruct(psReArmPad).functionality.rearmPad.timeStarted = ACTION_START_TIME Then
                    'set the time started and last updated
                    aStruct(psReArmPad).functionality.rearmPad.timeStarted = gameTime
                    aStruct(psReArmPad).functionality.rearmPad.timeLastUpdated = gameTime
                End If
                rearmPoints = getBuildingRearmPoints(psStructure)
                pointsToAdd = FromU32(Fix(U32(MUL32(rearmPoints, FromU32(U32(gameTime) - U32(aStruct(psReArmPad).functionality.rearmPad.timeStarted)))) / GAME_TICKS_PER_SEC))
                rearmPoints = getBuildingRearmPoints(psStructure)
                pointsAlreadyAdded = FromU32(Fix(U32(MUL32(rearmPoints, FromU32(U32(aStruct(psReArmPad).functionality.rearmPad.timeLastUpdated) - U32(aStruct(psReArmPad).functionality.rearmPad.timeStarted)))) / GAME_TICKS_PER_SEC))
                If Not ULT(pointsToAdd, aDroid(psDroid).weight) Then ' amount required is a factor of the droid weight
                    ' We should be fully loaded by now.
                    fillVtolDroid psDroid
                    objTrace aDroid(psDroid).id, "aiUpdateStructure", "fully loaded"
                Else
                    For i = 0 To aDroid(psDroid).numWeaps - 1      ' rearm one weapon at a time
                        ' Make sure it's a rearmable weapon (and so we don't divide by zero)
                        c = aDroid(psDroid).asWeaps(i).usedAmmo <> 0
                        If c Then c = asWeaponStats(DROID_getWeaponStats(psDroid, i)).upgrade(aDroid(psDroid).player).numRounds > 0
                        If c Then
                            ' Do not "simplify" this formula.
                            ' It is written this way to prevent rounding errors.
                            ammoToAddThisTime = _
                                FromU32(Fix(U32(MUL32(pointsToAdd, getNumAttackRuns(psDroid, i))) / U32(aDroid(psDroid).weight)) - _
                                        Fix(U32(MUL32(pointsAlreadyAdded, getNumAttackRuns(psDroid, i))) / U32(aDroid(psDroid).weight)))
                            ' usedAmmo -= std::min<unsigned>(ammoToAddThisTime, usedAmmo)
                            If ULT(ammoToAddThisTime, aDroid(psDroid).asWeaps(i).usedAmmo) Then
                                aDroid(psDroid).asWeaps(i).usedAmmo = FromU32(U32(aDroid(psDroid).asWeaps(i).usedAmmo) - U32(ammoToAddThisTime))
                            Else
                                aDroid(psDroid).asWeaps(i).usedAmmo = 0
                            End If
                            If ammoToAddThisTime <> 0 Then
                                ' reset ammo and lastFired
                                aDroid(psDroid).asWeaps(i).ammo = asWeaponStats(DROID_getWeaponStats(psDroid, i)).upgrade(aDroid(psDroid).player).numRounds
                                aDroid(psDroid).asWeaps(i).lastFired = 0
                                Exit For
                            End If
                        End If
                    Next
                End If
                If DROID_isDamaged(psDroid) Then ' do repairs
                    aDroid(psDroid).body = FromU32(U32(aDroid(psDroid).body) + gameTimeAdjustedAverage(getBuildingRepairPoints(psStructure)))
                    If Not ULT(aDroid(psDroid).body, aDroid(psDroid).originalBody) Then
                        aDroid(psDroid).body = aDroid(psDroid).originalBody
                    End If
                End If
                aStruct(psReArmPad).functionality.rearmPad.timeLastUpdated = gameTime

                'check for fully armed and fully repaired
                If vtolHappy(psDroid) Then
                    'clear the rearm pad
                    aDroid(psDroid).action = DACTION_NONE
                    aStruct(psReArmPad).functionality.rearmPad.psObj = 0
                    auxStructureNonblocking psStructure
                    triggerEventDroidIdle psDroid
                    objTrace aDroid(psDroid).id, "aiUpdateStructure", "VTOL happy and ready for action!"
                End If
            End If
        End If
    End If
End Sub

' Decides whether a structure should emit smoke when it's damaged
Private Function canSmoke(ByVal psStruct As Long) As Boolean
    If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_WALL Or asStructureStats(aStruct(psStruct).pStructureType).typ = REF_WALLCORNER _
       Or aStruct(psStruct).status = SS_BEING_BUILT Or asStructureStats(aStruct(psStruct).pStructureType).typ = REF_GATE Then
        canSmoke = False
    Else
        canSmoke = True
    End If
End Function

Private Function CalcStructureSmokeInterval(ByVal damage As Single) As Single
    CalcStructureSmokeInterval = F32((((1# - damage) + 0.1) * 10) * STRUCTURE_DAMAGE_SCALING)
End Function

' The main update routine for all Structures
Public Sub structureUpdate(ByVal psBuilding As Long, ByVal bMission As Boolean)
    Dim widthScatter As Long, breadthScatter As Long, emissionInterval As Long, iPointsToAdd As Long, iPointsRequired As Long
    Dim dv As Vector3i, i As Long, found_ As Boolean, gridList() As Long, gridList_size As Long, gi As Long, c As Boolean
    Dim strFirstBaseImd As Long, strFirstImd As Long, strImd As Long, damage As Long, effectTime As Long, size As Vector2i
    Dim position As Vector3i, pointX As Single, pointY As Single, pointZ As Single, realY As Long, pointIndex As Long
    Dim pDisplayModel As Long, psTile As Long, t As Long

    syncDebugStructure_ "structureUpdate", psBuilding, Asc("<")

    If (aStruct(psBuilding).flags And BIT_(OBJECT_FLAG_DIRTY)) <> 0 And Not bMission Then
        visTilesUpdate psBuilding
        aStruct(psBuilding).flags = aStruct(psBuilding).flags And Not BIT_(OBJECT_FLAG_DIRTY)
    End If

    If asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_GATE Then
        If aStruct(psBuilding).state = SAS_OPEN And ULT(FromU32(U32(aStruct(psBuilding).lastStateTime) + SAS_STAY_OPEN_TIME), gameTime) Then
            found_ = False

            ' static GridList gridList;  // static to avoid allocations.
            gridList_size = gridStartIterate(aStruct(psBuilding).pos.x, aStruct(psBuilding).pos.y, TILE_UNITS, gridList)
            For gi = 0 To gridList_size - 1
                If found_ Then Exit For
                found_ = isDroid(gridList(gi))
            Next

            If Not found_ Then ' no droids on our tile, safe to close
                aStruct(psBuilding).state = SAS_CLOSING
                auxStructureClosedGate psBuilding     ' closed
                aStruct(psBuilding).lastStateTime = gameTime ' reset timer
            End If
        ElseIf aStruct(psBuilding).state = SAS_OPENING And ULT(FromU32(U32(aStruct(psBuilding).lastStateTime) + SAS_OPEN_SPEED), gameTime) Then
            aStruct(psBuilding).state = SAS_OPEN
            auxStructureOpenGate psBuilding       ' opened
            aStruct(psBuilding).lastStateTime = gameTime ' reset timer
        ElseIf aStruct(psBuilding).state = SAS_CLOSING And ULT(FromU32(U32(aStruct(psBuilding).lastStateTime) + SAS_OPEN_SPEED), gameTime) Then
            aStruct(psBuilding).state = SAS_NORMAL
            aStruct(psBuilding).lastStateTime = gameTime ' reset timer
        End If
    ElseIf asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESOURCE_EXTRACTOR Then
        If aStruct(psBuilding).functionality.resourceExtractor.psPowerGen = 0 _
           And aStruct(psBuilding).animationEvent = ANIM_EVENT_ACTIVE Then ' no power generator connected
            resetObjectAnimationState psBuilding
        ElseIf aStruct(psBuilding).functionality.resourceExtractor.psPowerGen <> 0 _
               And aStruct(psBuilding).animationEvent = ANIM_EVENT_NONE _
               And aStruct(psBuilding).sDisplay.imd <> 0 Then ' we have a power generator, but no animation
            aStruct(psBuilding).animationEvent = ANIM_EVENT_ACTIVE

            strFirstBaseImd = iIMDShape_objanimpie(iIMDBaseShape_displayModel(aStruct(psBuilding).sDisplay.imd), aStruct(psBuilding).animationEvent)
            If strFirstBaseImd <> 0 Then strFirstImd = iIMDBaseShape_displayModel(strFirstBaseImd) Else strFirstImd = 0
            c = strFirstImd <> 0
            If c Then c = iIMDShape_next(strFirstImd) <> 0
            If c Then
                strImd = iIMDShape_next(strFirstImd) ' first imd isn't animated
                aStruct(psBuilding).timeAnimationStarted = gameTime + (rand() Mod (iIMDShape_objanimframes(strImd) * iIMDShape_objanimtime(strImd))) ' vary animation start time
            Else
                If Not c Then ASSERT_FAILURE "strFirstImd != nullptr && strFirstImd->next != nullptr", "src/structure.cpp:3845", "structureUpdate", "Unexpected objanimpie"
                aStruct(psBuilding).timeAnimationStarted = gameTime  ' so start animation
            End If
        End If
    End If

    ' Remove invalid targets. This must be done each frame.
    For i = 0 To MAX_WEAPONS - 1
        If aStruct(psBuilding).psTarget(i) <> 0 Then
            If SIMPLE_OBJECT_died(aStruct(psBuilding).psTarget(i)) <> 0 Then
                syncDebugObject_ "structureUpdate", aStruct(psBuilding).psTarget(i), Asc("-")
                setStructureTarget psBuilding, 0, i, ORIGIN_UNKNOWN
            End If
        End If
    Next

    'update the manufacture/research of the building once complete
    If aStruct(psBuilding).status = SS_BUILT Then
        aiUpdateStructure psBuilding, bMission
    End If

    If aStruct(psBuilding).status <> SS_BUILT Then
        If aStruct(psBuilding).selected <> 0 Then
            aStruct(psBuilding).selected = 0
        End If
    End If

    If Not bMission Then
        If aStruct(psBuilding).status = SS_BEING_BUILT And aStruct(psBuilding).buildRate = 0 And Not structureHasModules(psBuilding) Then
            If asStructureStats(aStruct(psBuilding).pStructureType).powerToBuild = 0 Then
                ' Building is free, and not currently being built, so deconstruct slowly over 1 minute.
                aStruct(psBuilding).currentBuildPts = aStruct(psBuilding).currentBuildPts - MIN_(aStruct(psBuilding).currentBuildPts, gameTimeAdjustedAverage(structureBuildPointsToCompletion(psBuilding), 60))
            End If

            If aStruct(psBuilding).currentBuildPts = 0 Then
                removeStruct psBuilding, True  ' If giving up on building something, remove the structure (and remove it from the power queue).
            End If
        End If
        aStruct(psBuilding).lastBuildRate = aStruct(psBuilding).buildRate
        aStruct(psBuilding).buildRate = 0  ' Reset to 0, each truck building us will add to our buildRate.
    End If

    ' Only add smoke if they're visible and they can 'burn'
    c = Not bMission
    If c Then c = BASE_OBJECT_visibleForLocalDisplay(psBuilding) <> 0
    If c Then c = canSmoke(psBuilding)
    If c Then
        damage = getStructureDamage(psBuilding)

        ' Is there any damage?
        If damage > 0 Then
            emissionInterval = Fix(CalcStructureSmokeInterval(fdiv(CSng(damage), 65536!)))
            ' unsigned effectTime = std::max(gameTime - deltaGameTime + 1, psBuilding->lastEmission + emissionInterval)
            effectTime = FromU32(U32(gameTime) - U32(deltaGameTime) + 1)
            t = FromU32(U32(aStruct(psBuilding).lastEmission) + U32(emissionInterval))
            If ULT(effectTime, t) Then effectTime = t
            If Not ULT(gameTime, effectTime) Then
                size = STRUCTURE_size(psBuilding)
                widthScatter = size.x * TILE_UNITS \ 2 \ 3
                breadthScatter = size.y * TILE_UNITS \ 2 \ 3
                dv.x = aStruct(psBuilding).pos.x + widthScatter - rand() Mod (2 * widthScatter)
                dv.z = aStruct(psBuilding).pos.y + breadthScatter - rand() Mod (2 * breadthScatter)
                dv.y = aStruct(psBuilding).pos.z
                dv.y = dv.y + (imdMaxY(aStruct(psBuilding).sDisplay.imd) * 3) \ 4
                addEffect dv, EFFECT_SMOKE, SMOKE_TYPE_DRIFTING_HIGH, False, 0, 0, effectTime
                aStruct(psBuilding).lastEmission = effectTime
            End If
        End If
    End If

    ' Update the fire damage data
    If aStruct(psBuilding).periodicalDamageStart <> 0 And aStruct(psBuilding).periodicalDamageStart <> gameTime - deltaGameTime Then ' -deltaGameTime, since projectiles are updated after structures.
        ' The periodicalDamageStart has been set, but is not from the previous tick, so we must be out of the fire.
        aStruct(psBuilding).periodicalDamage = 0  ' Reset burn damage done this tick.
        ' Finished burning.
        aStruct(psBuilding).periodicalDamageStart = 0
    End If

    'check the resistance level of the structure
    iPointsRequired = structureResistance(aStruct(psBuilding).pStructureType, aStruct(psBuilding).player)
    If aStruct(psBuilding).resistance < ToInt16(iPointsRequired) Then
        'start the resistance increase
        If aStruct(psBuilding).lastResistance = ACTION_START_TIME Then
            aStruct(psBuilding).lastResistance = gameTime
        End If
        'increase over time if low
        If ULT(RESISTANCE_INTERVAL, FromU32(U32(gameTime) - U32(aStruct(psBuilding).lastResistance))) Then
            aStruct(psBuilding).resistance = ToInt16(aStruct(psBuilding).resistance + 1)

            'in multiplayer, certain structures do not function whilst low resistance
            If bMultiPlayer Then
                resetResistanceLag psBuilding
            End If

            aStruct(psBuilding).lastResistance = gameTime
            'once the resistance is back up reset the last time increased
            If aStruct(psBuilding).resistance >= ToInt16(iPointsRequired) Then
                aStruct(psBuilding).lastResistance = ACTION_START_TIME
            End If
        End If
    Else
        'if selfrepair has been researched then check the health level of the
        'structure once resistance is fully up
        iPointsRequired = STRUCTURE_structureBody(psBuilding)
        c = selfRepairEnabled(aStruct(psBuilding).player)
        If c Then c = ULT(aStruct(psBuilding).body, iPointsRequired)
        If c Then c = aStruct(psBuilding).status <> SS_BEING_BUILT
        If c Then
            'start the self repair off
            If aStruct(psBuilding).lastResistance = ACTION_START_TIME Then
                aStruct(psBuilding).lastResistance = gameTime
            End If

            'since self repair, then add half repair points depending on the time delay for the stat
            iPointsToAdd = MUL32(repairPoints(asRepairStats(aDefaultRepair( _
                                 aStruct(psBuilding).player)), aStruct(psBuilding).player) \ 4, FromU32(Fix(U32(FromU32(U32(gameTime) - _
                                         U32(aStruct(psBuilding).lastResistance))) / U32(asRepairStats( _
                                                 aDefaultRepair(aStruct(psBuilding).player)).time_))))

            'add the blue flashing effect for multiPlayer
            c = bMultiPlayer
            If c Then c = (rand() Mod 10 = 0)   ' ONEINTEN
            If c Then c = Not bMission And aStruct(psBuilding).sDisplay.imd <> 0
            If c Then
                ' since this is a visual effect, it should be based on the *display* model
                pDisplayModel = iIMDBaseShape_displayModel(aStruct(psBuilding).sDisplay.imd)
                pointIndex = rand() Mod (iIMDShape_points_size(pDisplayModel) - 1)
                iIMDShape_point pDisplayModel, pointIndex, pointX, pointY, pointZ   ' point = &(pDisplayModel->points.at(pointIndex))
                position.x = ToInt(fadd(CSng(aStruct(psBuilding).pos.x), pointX))
                realY = ToInt(fmul(structHeightScale(psBuilding), pointY))
                position.y = aStruct(psBuilding).pos.z + realY
                position.z = ToInt(fsub(CSng(aStruct(psBuilding).pos.y), pointZ))
                psTile = mapTile(map_coord(position.x), map_coord(position.y))
                If tileIsClearlyVisible(psTile) Then
                    effectSetSize 30
                    addEffect position, EFFECT_EXPLOSION, EXPLOSION_TYPE_SPECIFIED, True, getDisplayImdFromIndex(MI_PLASMA), 0, gameTime - deltaGameTime + rand() Mod deltaGameTime
                End If
            End If

            If iPointsToAdd <> 0 Then
                aStruct(psBuilding).body = FromU32(U32(aStruct(psBuilding).body) + U32(iPointsToAdd)) And &HFFFF&   ' (UWORD)
                aStruct(psBuilding).lastResistance = gameTime
                If ULT(iPointsRequired, aStruct(psBuilding).body) Then
                    aStruct(psBuilding).body = iPointsRequired And &HFFFF&   ' (UWORD)
                    aStruct(psBuilding).lastResistance = ACTION_START_TIME
                End If
            End If
        End If
    End If

    syncDebugStructure_ "structureUpdate", psBuilding, Asc(">")

    checkStructure psBuilding, "src/structure.cpp:4014", "structureUpdate", max_check_object_recursion
End Sub

'checks to see if a specific structure type exists -as opposed to a structure
'stat type
Public Function checkSpecificStructExists(ByVal structInc As Long, ByVal player As Long) As Boolean
    Dim it As Long, psStructure As Long

    If Not (ULT(structInc, numStructureStats)) Then ASSERT_FAILURE "structInc < numStructureStats", "src/structure.cpp:4908", "checkSpecificStructExists", "Invalid structure inc": checkSpecificStructExists = False: Exit Function

    it = apsStructLists(player).head
    Do While it <> 0
        psStructure = lnObj(it)
        If aStruct(psStructure).status = SS_BUILT Then
            If asStructureStats(aStruct(psStructure).pStructureType).ref - STAT_STRUCTURE = structInc Then
                checkSpecificStructExists = True
                Exit Function
            End If
        End If
        it = lnNext(it)
    Loop
    checkSpecificStructExists = False
End Function

' Count the artillery and VTOL droids assigned to a structure.
Private Function countAssignedDroids(ByVal psStructure As Long) As Long
    Dim num As Long, it As Long, psCurr As Long, weapontype As Long

    checkStructure psStructure, "src/structure.cpp:5482", "countAssignedDroids", max_check_object_recursion

    ' For non-debug builds
    If psStructure = 0 Then
        countAssignedDroids = 0
        Exit Function
    End If

    If Not ULT(selectedPlayer, MAX_PLAYERS) Then
        countAssignedDroids = 0
        Exit Function
    End If

    num = 0
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aDroid(psCurr).order.psObj <> 0 Then
            If SIMPLE_OBJECT_id(aDroid(psCurr).order.psObj) = aStruct(psStructure).id _
               And aDroid(psCurr).player = aStruct(psStructure).player Then
                weapontype = asWeaponStats(DROID_getWeaponStats(psCurr, 0)).movementModel

                If weapontype = MM_INDIRECT _
                   Or weapontype = MM_HOMINGINDIRECT _
                   Or DROID_isVtol(psCurr) Then
                    num = num + 1
                End If
            End If
        End If
        it = lnNext(it)
    Loop

    countAssignedDroids = num
End Function

'Checks the template type against the factory type - returns false
'if not a good combination!
Public Function validTemplateForFactory(ByVal psTemplate As Long, ByVal psFactory As Long, ByVal complain As Boolean) As Boolean
    Dim level As Long, c As Boolean

    If Not (psTemplate <> 0) Then ASSERT_FAILURE "psTemplate", "src/structure.cpp:5651", "validTemplateForFactory", "Invalid template!": validTemplateForFactory = False: Exit Function
    If complain Then level = LOG_ERROR Else level = LOG_NEVER

    'not in multiPlayer! - AB 26/5/99
    'ignore Transporter Droids
    If Not bMultiPlayer And isTransporter_template(apsTemplates(psTemplate)) Then
        debug_ 5658, level, "validTemplateForFactory", "Cannot build transporter in campaign."
        validTemplateForFactory = False
        Exit Function
    End If

    'check if droid is a cyborg
    If apsTemplates(psTemplate).droidType = DROID_CYBORG Or _
       apsTemplates(psTemplate).droidType = DROID_CYBORG_SUPER Or _
       apsTemplates(psTemplate).droidType = DROID_CYBORG_CONSTRUCT Or _
       apsTemplates(psTemplate).droidType = DROID_CYBORG_REPAIR Then
        If asStructureStats(aStruct(psFactory).pStructureType).typ <> REF_CYBORG_FACTORY Then
            debug_ 5670, level, "validTemplateForFactory", "Cannot build cyborg except in cyborg factory, not in %s.", objInfo(psFactory)
            validTemplateForFactory = False
            Exit Function
        End If
    'check for VTOL droid
    ElseIf apsTemplates(psTemplate).asParts(COMP_PROPULSION) <> 0 And _
           (asPropulsionStats(DROID_TEMPLATE_getPropulsionStats(apsTemplates(psTemplate))).propulsionType = PROPULSION_TYPE_LIFT) Then
        If asStructureStats(aStruct(psFactory).pStructureType).typ <> REF_VTOL_FACTORY Then
            debug_ 5680, level, "validTemplateForFactory", "Cannot build vtol except in vtol factory, not in %s.", objInfo(psFactory)
            validTemplateForFactory = False
            Exit Function
        End If
    End If

    'check if cyborg factory
    If asStructureStats(aStruct(psFactory).pStructureType).typ = REF_CYBORG_FACTORY Then
        If Not (apsTemplates(psTemplate).droidType = DROID_CYBORG Or _
                apsTemplates(psTemplate).droidType = DROID_CYBORG_SUPER Or _
                apsTemplates(psTemplate).droidType = DROID_CYBORG_CONSTRUCT Or _
                apsTemplates(psTemplate).droidType = DROID_CYBORG_REPAIR) Then
            debug_ 5693, level, "validTemplateForFactory", "Can only build cyborg in cyborg factory, not droidType %d in %s.", _
                   apsTemplates(psTemplate).droidType, objInfo(psFactory)
            validTemplateForFactory = False
            Exit Function
        End If
    'check if vtol factory
    ElseIf asStructureStats(aStruct(psFactory).pStructureType).typ = REF_VTOL_FACTORY Then
        c = apsTemplates(psTemplate).asParts(COMP_PROPULSION) = 0
        If Not c Then c = (asPropulsionStats(DROID_TEMPLATE_getPropulsionStats(apsTemplates(psTemplate))).propulsionType <> PROPULSION_TYPE_LIFT)
        If c Then
            debug_ 5704, level, "validTemplateForFactory", "Can only build vtol in vtol factory, not in %s.", objInfo(psFactory)
            validTemplateForFactory = False
            Exit Function
        End If
    End If

    'got through all the tests...
    validTemplateForFactory = True
End Function

' Return true if flag is a delivery point for a factory.
'
Public Function FlagIsFactory(ByVal psCurrFlag As Long) As Boolean
    If (aFlagPos(psCurrFlag).factoryType = FACTORY_FLAG) Or (aFlagPos(psCurrFlag).factoryType = CYBORG_FLAG) Or _
       (aFlagPos(psCurrFlag).factoryType = VTOL_FLAG) Then
        FlagIsFactory = True
        Exit Function
    End If

    FlagIsFactory = False
End Function

' Find a structure's delivery point , only if it's a factory.
' Returns NULL if not found or the structure isn't a factory.
'
Public Function FindFactoryDelivery(ByVal struct_ As Long) As Long
    Dim it As Long, psFlag As Long, psAssemblyPoint As Long
    If struct_ <> 0 Then
        If STRUCTURE_isFactory(struct_) Then
            ' Find the factories delivery point.
            psAssemblyPoint = aStruct(struct_).functionality.factory.psAssemblyPoint
            it = apsFlagPosLists(aStruct(struct_).player).head
            Do While it <> 0
                psFlag = lnObj(it)
                If FlagIsFactory(psFlag) _
                   And aFlagPos(psAssemblyPoint).factoryInc = aFlagPos(psFlag).factoryInc _
                   And aFlagPos(psAssemblyPoint).factoryType = aFlagPos(psFlag).factoryType Then
                    FindFactoryDelivery = psFlag
                    Exit Function
                End If
                it = lnNext(it)
            Loop
        End If
    End If

    FindFactoryDelivery = 0
End Function

'Find the factory associated with the delivery point - returns NULL if none exist
Public Function findDeliveryFactory(ByVal psDelPoint As Long) As Long
    Dim it As Long, psCurr As Long, psFactory As Long, psRepair As Long

    it = apsStructLists(aFlagPos(psDelPoint).player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If psCurr <> 0 Then
            If STRUCTURE_isFactory(psCurr) Then
                psFactory = psCurr  ' FACTORY *
                If aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryInc = aFlagPos(psDelPoint).factoryInc And _
                   aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryType = aFlagPos(psDelPoint).factoryType Then
                    findDeliveryFactory = psCurr
                    Exit Function
                End If
            ElseIf asStructureStats(aStruct(psCurr).pStructureType).typ = REF_REPAIR_FACILITY Then
                psRepair = psCurr  ' REPAIR_FACILITY *
                If aStruct(psRepair).functionality.repairFacility.psDeliveryPoint = psDelPoint Then
                    findDeliveryFactory = psCurr
                    Exit Function
                End If
            End If
        End If
        it = lnNext(it)
    Loop
    findDeliveryFactory = 0
End Function

'cancels the production run for the factory and returns any power that was
'accrued but not used
Public Sub cancelProduction(ByVal psBuilding As Long, ByVal mode As QUEUE_MODE, Optional ByVal mayClearProductionRun As Boolean = True)
    Dim psFactory As Long, factoryType As Long, factoryInc As Long

    If psBuilding = 0 Then ASSERT_FAILURE "psBuilding && psBuilding->isFactory()", "src/structure.cpp:6203", "cancelProduction", "structure not a factory": Exit Sub
    If Not STRUCTURE_isFactory(psBuilding) Then ASSERT_FAILURE "psBuilding && psBuilding->isFactory()", "src/structure.cpp:6203", "cancelProduction", "structure not a factory": Exit Sub

    psFactory = psBuilding  ' FACTORY *

    If aStruct(psBuilding).player = productionPlayer And mayClearProductionRun Then
        'clear the production run for this factory
        factoryType = aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryType
        factoryInc = aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryInc
        If factoryInc < asProductionRun(factoryType).size Then
            productionRunClear asProductionRun(factoryType).runs(factoryInc)
        End If
        aStruct(psFactory).functionality.factory.productionLoops = 0
    End If

    If mode = ModeQueue Then
        sendStructureInfo psBuilding, STRUCTUREINFO_CANCELPRODUCTION, 0
        setStatusPendingCancel_factory aStruct(psFactory).functionality.factory

        Exit Sub
    End If

    'clear the factory's subject
    refundFactoryBuildPower psBuilding
    aStruct(psFactory).functionality.factory.psSubject = 0
    delPowerRequest psBuilding
End Sub

Public Sub doNextProduction(ByVal psStructure As Long, ByVal current As Long, ByVal mode As QUEUE_MODE)
    Dim psNextTemplate As Long
    psNextTemplate = factoryProdUpdate(psStructure, current)

    If psNextTemplate <> 0 Then
        structSetManufacture psStructure, psNextTemplate, ModeQueue  ' ModeQueue instead of mode, since production lists aren't currently synchronised.
    Else
        cancelProduction psStructure, mode
    End If
End Sub

' bool ProductionRunEntry::operator ==(DROID_TEMPLATE *t) const
Public Function ProductionRunEntry_equals(e As ProductionRunEntry, ByVal t As Long) As Boolean
    ProductionRunEntry_equals = apsTemplates(e.psTemplate).multiPlayerID = apsTemplates(t).multiPlayerID
End Function

' (VB) std::find(productionRun.begin(), productionRun.end(), psTemplate): the entry's index, or productionRun.size() (end())
Private Function productionRunFind(prodRun As ProductionRun, ByVal psTemplate As Long) As Long
    Dim i As Long
    For i = 0 To prodRun.size - 1
        If ProductionRunEntry_equals(prodRun.items(i), psTemplate) Then Exit For
    Next
    productionRunFind = i
End Function

' (VB) productionRun.erase(productionRun.begin() + entry)
Private Sub productionRunErase(prodRun As ProductionRun, ByVal entry As Long)
    Dim i As Long
    For i = entry To prodRun.size - 2
        prodRun.items(i) = prodRun.items(i + 1)
    Next
    prodRun.size = prodRun.size - 1
End Sub

' (VB) productionRun.push_back(e)
Private Sub productionRunPushBack(prodRun As ProductionRun, e As ProductionRunEntry)
    If prodRun.size = 0 Then
        ReDim prodRun.items(0 To 0)
    ElseIf prodRun.size > UBound(prodRun.items) Then
        ReDim Preserve prodRun.items(0 To prodRun.size)
    End If
    prodRun.items(prodRun.size) = e
    prodRun.size = prodRun.size + 1
End Sub

' (VB) productionRun.clear()
Private Sub productionRunClear(prodRun As ProductionRun)
    Erase prodRun.items
    prodRun.size = 0
End Sub

' (VB) asProductionRun[type].resize(n), growing: the new production runs are empty
Private Sub productionRunListResize(list As ProductionRunList, ByVal n As Long)
    ReDim Preserve list.runs(0 To n - 1)
    list.size = n
End Sub

'this is called when a factory produces a droid. The Template returned is the next
'one to build - if any
Public Function factoryProdUpdate(ByVal psStructure As Long, ByVal psTemplate As Long) As Long
    Dim psFactory As Long, factoryType As Long, factoryInc As Long, entry As Long, inc As Long

    checkStructure psStructure, "src/structure.cpp:6312", "factoryProdUpdate", max_check_object_recursion
    If aStruct(psStructure).player <> productionPlayer Then
        factoryProdUpdate = 0  ' Production lists not currently synchronised.
        Exit Function
    End If

    psFactory = psStructure  ' FACTORY *
    factoryType = aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryType
    factoryInc = aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryInc
    If factoryInc >= asProductionRun(factoryType).size Then
        factoryProdUpdate = 0  ' Don't even have a production list.
        Exit Function
    End If
    ' ProductionRun &productionRun = asProductionRun[factoryType][factoryInc]

    If psTemplate <> 0 Then
        'find the entry in the array for this template
        entry = productionRunFind(asProductionRun(factoryType).runs(factoryInc), psTemplate)
        If entry <> asProductionRun(factoryType).runs(factoryInc).size Then
            asProductionRun(factoryType).runs(factoryInc).items(entry).built = MIN_(asProductionRun(factoryType).runs(factoryInc).items(entry).built + 1, asProductionRun(factoryType).runs(factoryInc).items(entry).quantity)
            If Not ProductionRunEntry_isComplete(asProductionRun(factoryType).runs(factoryInc).items(entry)) Then
                factoryProdUpdate = psTemplate  ' Build another of the same type.
                Exit Function
            End If
            If aStruct(psFactory).functionality.factory.productionLoops = 0 Then
                productionRunErase asProductionRun(factoryType).runs(factoryInc), entry
            End If
        End If
    End If
    'find the next template to build - this just looks for the first uncompleted run
    For inc = 0 To asProductionRun(factoryType).runs(factoryInc).size - 1
        If Not ProductionRunEntry_isComplete(asProductionRun(factoryType).runs(factoryInc).items(inc)) Then
            factoryProdUpdate = asProductionRun(factoryType).runs(factoryInc).items(inc).psTemplate
            Exit Function
        End If
    Next
    ' Check that we aren't looping doing nothing.
    If asProductionRun(factoryType).runs(factoryInc).size = 0 Then
        If aStruct(psFactory).functionality.factory.productionLoops <> INFINITE_PRODUCTION Then
            aStruct(psFactory).functionality.factory.productionLoops = 0  ' Reset number of loops, unless set to infinite.
        End If
    ElseIf aStruct(psFactory).functionality.factory.productionLoops <> 0 Then  'If you've got here there's nothing left to build unless factory is on loop production
        'reduce the loop count if not infinite
        If aStruct(psFactory).functionality.factory.productionLoops <> INFINITE_PRODUCTION Then
            aStruct(psFactory).functionality.factory.productionLoops = aStruct(psFactory).functionality.factory.productionLoops - 1
        End If

        'need to reset the quantity built for each entry in the production list
        For inc = 0 To asProductionRun(factoryType).runs(factoryInc).size - 1
            ProductionRunEntry_restart asProductionRun(factoryType).runs(factoryInc).items(inc)
        Next

        'get the first to build again
        factoryProdUpdate = asProductionRun(factoryType).runs(factoryInc).items(0).psTemplate
        Exit Function
    End If
    'if got to here then nothing left to produce so clear the array
    productionRunClear asProductionRun(factoryType).runs(factoryInc)
    factoryProdUpdate = 0
End Function

'adjust the production run for this template type
Public Sub factoryProdAdjust(ByVal psStructure As Long, ByVal psTemplate As Long, ByVal add As Boolean)
    Dim psFactory As Long, factoryType As Long, factoryInc As Long, entry As Long, tmplEntry As ProductionRunEntry
    Dim production As ProductionRunEntry

    checkStructure psStructure, "src/structure.cpp:6381", "factoryProdAdjust", max_check_object_recursion
    If Not (aStruct(psStructure).player = productionPlayer) Then ASSERT_FAILURE "psStructure->player == productionPlayer", "src/structure.cpp:6382", "factoryProdAdjust", "called for incorrect player": Exit Sub
    If Not (psTemplate <> 0) Then ASSERT_FAILURE "psTemplate != nullptr", "src/structure.cpp:6383", "factoryProdAdjust", "NULL template": Exit Sub

    psFactory = psStructure  ' FACTORY *
    factoryType = aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryType
    factoryInc = aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryInc
    If factoryInc >= asProductionRun(factoryType).size Then
        productionRunListResize asProductionRun(factoryType), factoryInc + 1  ' Don't have a production list, create it.
    End If
    ' ProductionRun &productionRun = asProductionRun[factoryType][factoryInc]

    'see if the template is already in the list
    entry = productionRunFind(asProductionRun(factoryType).runs(factoryInc), psTemplate)

    If entry <> asProductionRun(factoryType).runs(factoryInc).size Then
        If aStruct(psFactory).functionality.factory.productionLoops = 0 Then
            ProductionRunEntry_removeComplete asProductionRun(factoryType).runs(factoryInc).items(entry) ' We are not looping, so remove the built droids from the list, so that quantity corresponds to the displayed number.
        End If

        'adjust the prod run
        If add Then
            asProductionRun(factoryType).runs(factoryInc).items(entry).quantity = asProductionRun(factoryType).runs(factoryInc).items(entry).quantity + 1
        Else
            asProductionRun(factoryType).runs(factoryInc).items(entry).quantity = asProductionRun(factoryType).runs(factoryInc).items(entry).quantity - 1
        End If
        asProductionRun(factoryType).runs(factoryInc).items(entry).built = MIN_(asProductionRun(factoryType).runs(factoryInc).items(entry).built, asProductionRun(factoryType).runs(factoryInc).items(entry).quantity)

        ' Allows us to queue up more units up to MAX_IN_RUN instead of ignoring how many we have built from that queue
        ' check to see if user canceled all orders in queue
        If asProductionRun(factoryType).runs(factoryInc).items(entry).quantity <= 0 Or asProductionRun(factoryType).runs(factoryInc).items(entry).quantity > MAX_IN_RUN Then
            productionRunErase asProductionRun(factoryType).runs(factoryInc), entry  ' Entry empty, so get rid of it.
        End If
    Else
        'start off a new template
        tmplEntry.psTemplate = psTemplate
        If add Then tmplEntry.quantity = 1 Else tmplEntry.quantity = MAX_IN_RUN 'wrap around to max value
        tmplEntry.built = 0
        productionRunPushBack asProductionRun(factoryType).runs(factoryInc), tmplEntry
    End If
    'if nothing is allocated then the current factory may have been cancelled
    If asProductionRun(factoryType).runs(factoryInc).size = 0 Then
        'must have cancelled everything - so tell the struct
        If aStruct(psFactory).functionality.factory.productionLoops <> INFINITE_PRODUCTION Then
            aStruct(psFactory).functionality.factory.productionLoops = 0  ' Reset number of loops, unless set to infinite.
        End If
    End If

    'need to check if this was the template that was mid-production
    production = getProduction(psStructure, FactoryGetTemplate(psFactory))
    If ProductionRunEntry_numRemaining(production) = 0 Then
        doNextProduction psStructure, FactoryGetTemplate(psFactory), ModeQueue
    ElseIf Not StructureIsManufacturingPending(psStructure) Then
        structSetManufacture psStructure, psTemplate, ModeQueue
    End If

    If StructureIsOnHoldPending(psStructure) Then
        releaseProduction psStructure, ModeQueue
    End If
End Sub

' checks the status of the production of a template
Public Function getProduction(ByVal psStructure As Long, ByVal psTemplate As Long) As ProductionRunEntry
    Dim psFactory As Long, factoryType As Long, factoryInc As Long, entry As Long, c As Boolean

    c = psStructure = 0
    If Not c Then c = aStruct(psStructure).player <> productionPlayer
    If Not c Then c = psTemplate = 0
    If Not c Then c = Not STRUCTURE_isFactory(psStructure)
    If c Then
        Exit Function  ' Not producing any NULL pointers. (ProductionRunEntry())
    End If

    psFactory = psStructure  ' FACTORY *
    c = aStruct(psFactory).functionality.factory.psAssemblyPoint = 0
    If Not c Then c = aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryInc >= asProductionRun(aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryType).size
    If c Then
        Exit Function  ' Don't have a production list. (ProductionRunEntry())
    End If
    factoryType = aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryType
    factoryInc = aFlagPos(aStruct(psFactory).functionality.factory.psAssemblyPoint).factoryInc
    ' ProductionRun &productionRun = asProductionRun[factoryType][factoryInc]

    'see if the template is in the list
    entry = productionRunFind(asProductionRun(factoryType).runs(factoryInc), psTemplate)
    If entry <> asProductionRun(factoryType).runs(factoryInc).size Then
        getProduction = asProductionRun(factoryType).runs(factoryInc).items(entry)
        Exit Function
    End If

    'not in the list so none being produced
    ' (ProductionRunEntry())
End Function

' Count number of factories assignable to a command droid.
'
Public Function countAssignableFactories(ByVal player As Long, ByVal factoryType As Long) As Long
    Dim quantity As Long, factoryInc As Long

    quantity = 0

    If Not (player = selectedPlayer) Then ASSERT_FAILURE "player == selectedPlayer", "src/structure.cpp:6481", "countAssignableFactories", "%s should only be called for selectedPlayer", "countAssignableFactories": countAssignableFactories = 0: Exit Function

    If player >= MAX_PLAYERS Then countAssignableFactories = 0: Exit Function

    For factoryInc = 0 To factoryNumFlag_size(player, factoryType) - 1
        'check to see if there is a factory
        If factoryNumFlag(player, factoryType, factoryInc) Then
            quantity = (quantity + 1) And &HFF&   ' UBYTE
        End If
    Next

    countAssignableFactories = quantity
End Function

' check whether a factory of a certain number and type exists
Public Function checkFactoryExists(ByVal player As Long, ByVal factoryType As Long, ByVal inc As Long) As Boolean
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/structure.cpp:6501", "checkFactoryExists", "Invalid player": checkFactoryExists = False: Exit Function
    If Not (ULT(factoryType, NUM_FACTORY_TYPES)) Then ASSERT_FAILURE "factoryType < NUM_FACTORY_TYPES", "src/structure.cpp:6502", "checkFactoryExists", "Invalid factoryType": checkFactoryExists = False: Exit Function

    If ULT(inc, factoryNumFlag_size(player, factoryType)) Then checkFactoryExists = factoryNumFlag(player, factoryType, inc)
End Function

'check that delivery points haven't been put down in invalid location
Public Sub checkDeliveryPoints(ByVal version As Long)
    Dim inc As Long, psFactory As Long, psRepair As Long, x As Long, y As Long, it As Long, psStruct As Long, psFlag As Long

    'find any factories
    For inc = 0 To MAX_PLAYERS - 1
        'don't bother checking selectedPlayer's - causes problems when try and use
        'validLocation since it finds that the DP is on itself! And validLocation
        'will have been called to put in down in the first place
        If inc <> selectedPlayer Then
            it = apsStructLists(inc).head
            Do While it <> 0
                psStruct = lnObj(it)
                If psStruct <> 0 Then
                    If STRUCTURE_isFactory(psStruct) Then
                        'check the DP
                        psFactory = psStruct  ' FACTORY *
                        If aStruct(psFactory).functionality.factory.psAssemblyPoint = 0 Then 'need to add one
                            If Not (aStruct(psFactory).functionality.factory.psAssemblyPoint <> 0) Then ASSERT_FAILURE "psFactory->psAssemblyPoint != nullptr", "src/structure.cpp:6536", "checkDeliveryPoints", "no delivery point for factory": Exit Sub
                        Else
                            psFlag = aStruct(psFactory).functionality.factory.psAssemblyPoint
                            setAssemblyPoint psFlag, aFlagPos(psFlag).coords.x, aFlagPos(psFlag).coords.y, inc, True
                        End If
                    ElseIf asStructureStats(aStruct(psStruct).pStructureType).typ = REF_REPAIR_FACILITY Then
                        psRepair = psStruct  ' REPAIR_FACILITY *

                        If aStruct(psRepair).functionality.repairFacility.psDeliveryPoint = 0 Then 'need to add one
                            If Not ULT(version, VERSION_19) Then
                                If Not (aStruct(psRepair).functionality.repairFacility.psDeliveryPoint <> 0) Then ASSERT_FAILURE "psRepair->psDeliveryPoint != nullptr", "src/structure.cpp:6552", "checkDeliveryPoints", "no delivery point for repair facility": Exit Sub
                            Else
                                ' add an assembly point
                                If Not createFlagPosition(aStruct(psRepair).functionality.repairFacility.psDeliveryPoint, aStruct(psStruct).player) Then
                                    ASSERT_FAILURE "!""can't create new delivery point for repair facility""", "src/structure.cpp:6559", "checkDeliveryPoints", "unable to create new delivery point for repair facility"
                                    Exit Sub
                                End If
                                addFlagPosition aStruct(psRepair).functionality.repairFacility.psDeliveryPoint
                                setFlagPositionInc aStruct(psStruct).pFunctionality, aStruct(psStruct).player, REPAIR_FLAG
                                'initialise the assembly point position
                                x = map_coord(aStruct(psStruct).pos.x + 256)
                                y = map_coord(aStruct(psStruct).pos.y + 256)
                                ' Belt and braces - shouldn't be able to build too near edge
                                setAssemblyPoint aStruct(psRepair).functionality.repairFacility.psDeliveryPoint, world_coord(x), _
                                                 world_coord(y), inc, True
                            End If
                        Else 'check existing one
                            psFlag = aStruct(psRepair).functionality.repairFacility.psDeliveryPoint
                            setAssemblyPoint psFlag, aFlagPos(psFlag).coords.x, aFlagPos(psFlag).coords.y, inc, True
                        End If
                    End If
                End If
                it = lnNext(it)
            Loop
        End If
    Next
End Sub

'adjust the loop quantity for this factory
Public Sub factoryLoopAdjust(ByVal psStruct As Long, ByVal add As Boolean)
    Dim psFactory As Long

    If psStruct = 0 Then ASSERT_FAILURE "psStruct && psStruct->isFactory()", "src/structure.cpp:6589", "factoryLoopAdjust", "structure is not a factory": Exit Sub
    If Not STRUCTURE_isFactory(psStruct) Then ASSERT_FAILURE "psStruct && psStruct->isFactory()", "src/structure.cpp:6589", "factoryLoopAdjust", "structure is not a factory": Exit Sub
    If Not (aStruct(psStruct).player = selectedPlayer) Then ASSERT_FAILURE "psStruct->player == selectedPlayer", "src/structure.cpp:6590", "factoryLoopAdjust", "should only be called for selectedPlayer": Exit Sub

    psFactory = psStruct  ' FACTORY *

    If add Then
        'check for wrapping to infinite production
        If aStruct(psFactory).functionality.factory.productionLoops = MAX_IN_RUN Then
            aStruct(psFactory).functionality.factory.productionLoops = 0
        Else
            'increment the count
            aStruct(psFactory).functionality.factory.productionLoops = aStruct(psFactory).functionality.factory.productionLoops + 1
            'check for limit - this caters for when on infinite production and want to wrap around
            If aStruct(psFactory).functionality.factory.productionLoops > MAX_IN_RUN Then
                aStruct(psFactory).functionality.factory.productionLoops = INFINITE_PRODUCTION
            End If
        End If
    Else
        'decrement the count
        If aStruct(psFactory).functionality.factory.productionLoops = 0 Then
            aStruct(psFactory).functionality.factory.productionLoops = INFINITE_PRODUCTION
        Else
            aStruct(psFactory).functionality.factory.productionLoops = aStruct(psFactory).functionality.factory.productionLoops - 1
        End If
    End If
End Sub

'Used for determining how much of the structure to draw as being built or demolished
Public Function structHeightScale(ByVal psStruct As Long) As Single
    Dim progress As Single
    progress = structureCompletionProgress(psStruct)
    If progress > 0.05! Then structHeightScale = progress Else structHeightScale = 0.05!   ' MAX
End Function

'for MULTIPLAYER ONLY
'this adjusts the time the relevant action started if the building is attacked by EW weapon
Private Sub resetResistanceLag(ByVal psBuilding As Long)
    Dim psFactory As Long
    If bMultiPlayer Then
        Select Case asStructureStats(aStruct(psBuilding).pStructureType).typ
        Case REF_RESEARCH
        Case REF_FACTORY, REF_VTOL_FACTORY, REF_CYBORG_FACTORY
            psFactory = psBuilding  ' FACTORY *

            'if working on a unit
            If aStruct(psFactory).functionality.factory.psSubject <> 0 Then
                'adjust the start time for the current subject
                If aStruct(psFactory).functionality.factory.timeStarted <> ACTION_START_TIME Then
                    aStruct(psFactory).functionality.factory.timeStarted = FromU32(U32(aStruct(psFactory).functionality.factory.timeStarted) + (U32(gameTime) - U32(aStruct(psBuilding).lastResistance)))
                End If
            End If
            ' (falls through to default: do nothing)
        Case Else 'do nothing
        End Select
    End If
End Sub

'checks the structure passed in is a Las Sat structure which is currently
'selected - returns true if valid
Public Function lasSatStructSelected(ByVal psStruct As Long) As Boolean
    Dim c As Boolean
    c = aStruct(psStruct).selected <> 0
    If Not c Then c = bMultiPlayer And Not isHumanPlayer(aStruct(psStruct).player)
    If c Then c = aStruct(psStruct).asWeaps(0).nStat <> 0
    If c Then c = (asWeaponStats(BASE_OBJECT_getWeaponStats(psStruct, 0)).weaponSubClass = WSC_LAS_SAT)
    If c Then
        lasSatStructSelected = True
        Exit Function
    End If

    lasSatStructSelected = False
End Function

' Call CALL_NEWDROID script callback
Public Sub cbNewDroid(ByVal psFactory As Long, ByVal psDroid As Long)
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/structure.cpp:7111", "cbNewDroid", "no droid assigned for CALL_NEWDROID callback": Exit Sub
    triggerEventDroidBuilt psDroid, psFactory
End Sub

'--------------------------------------------------------------------------
' structure.cpp:4257 bool isBlueprintTooClose(STRUCTURE_STATS const *stats1, Vector2i pos1, uint16_t dir1,
'                                             STRUCTURE_STATS const *stats2, Vector2i pos2, uint16_t dir2)
Public Function isBlueprintTooClose(ByVal stats1 As Long, pos1 As Vector2i, ByVal dir1 As Long, ByVal stats2 As Long, pos2 As Vector2i, ByVal dir2 As Long) As Boolean
    Dim packable As Boolean, minDist As Long, b1 As StructureBounds, b2 As StructureBounds
    Dim delta12 As Vector2i, delta21 As Vector2i, dist As Long
    If stats1 = stats2 And pos1.x = pos2.x And pos1.y = pos2.y And dir1 = dir2 Then
        isBlueprintTooClose = False  ' Same blueprint, so ignore it.
        Exit Function
    End If

    packable = canPack(baseStructureTypePackability(asStructureStats(stats1).typ), baseStructureTypePackability(asStructureStats(stats2).typ))
    If packable Then minDist = 0 Else minDist = 1
    b1 = getStructureBounds_structStats(stats1, pos1, dir1)
    b2 = getStructureBounds_structStats(stats2, pos2, dir2)
    delta12.x = b2.map.x - (b1.map.x + b1.size.x)
    delta12.y = b2.map.y - (b1.map.y + b1.size.y)
    delta21.x = b1.map.x - (b2.map.x + b2.size.x)
    delta21.y = b1.map.y - (b2.map.y + b2.size.y)
    dist = delta12.x
    If delta21.x > dist Then dist = delta21.x
    If delta12.y > dist Then dist = delta12.y
    If delta21.y > dist Then dist = delta21.y
    isBlueprintTooClose = dist < minDist
End Function

' optional<STRUCTURE>: a structure handle, 0 = nullopt. (VB) The STRUCTURE C++ returns by value is the
' structure pool's scratch slot 0 (paged_entity_container_h's StructContainer_scratch): outside every
' list, valid until the next call.
Public Function buildBlueprint(ByVal psStats As Long, pos As Vector3i, ByVal direction As Long, ByVal moduleIndex As Long, ByVal state As STRUCT_STATES, ByVal ownerPlayer As Long) As Long
    Dim rot As Rotation, moduleNumber As Long, pIMDStats As Long, baseStruct As Long, baseModuleNumber As Long
    Dim blueprint_ As Long, scanType As Long, typ As Long, p As Vector3i, n As Long
    If Not (psStats <> -1) Then ASSERT_FAILURE "psStats != nullptr", "src/structure.cpp:1980", "buildBlueprint", "No blueprint stats": Exit Function
    If Not (asStructureStats(psStats).pIMD_size > 0) Then ASSERT_FAILURE "psStats->pIMD[0] != nullptr", "src/structure.cpp:1981", "buildBlueprint", "No blueprint model for %s", asStructureStats(psStats).name_: Exit Function
    If Not (asStructureStats(psStats).pIMD(0) <> 0) Then ASSERT_FAILURE "psStats->pIMD[0] != nullptr", "src/structure.cpp:1981", "buildBlueprint", "No blueprint model for %s", asStructureStats(psStats).name_: Exit Function
    If Not ULT(ownerPlayer, MAX_PLAYERS) Then ASSERT_FAILURE "ownerPlayer < MAX_PLAYERS", "src/structure.cpp:1982", "buildBlueprint", "invalid ownerPlayer: %u", ownerPlayer: Exit Function

    rot.direction = direction           ' Rotation rot(direction, 0, 0);
    rot.pitch = 0
    rot.roll = 0
    p = pos

    moduleNumber = 0
    pIMDStats = psStats                 ' std::vector<iIMDBaseShape *> const *pIMD = &psStats->pIMD;
    If IsStatExpansionModule(psStats) Then
        baseStruct = castStructure(psMapTiles(worldTile(p.x, p.y)).psObject)
        If baseStruct <> 0 Then
            If moduleIndex = 0 Then
                moduleIndex = nextModuleToBuild(baseStruct, 0)
            End If
            baseModuleNumber = moduleIndex * 2 - 1 ' *2-1 because even-numbered IMDs are structures, odd-numbered IMDs are just the modules.
            If ULT(baseModuleNumber, asStructureStats(aStruct(baseStruct).pStructureType).pIMD_size) Then
                ' Draw the module.
                moduleNumber = baseModuleNumber
                pIMDStats = aStruct(baseStruct).pStructureType
                p = aStruct(baseStruct).pos
                rot = aStruct(baseStruct).rot
            End If
        End If
    End If

    blueprint_ = StructContainer_scratch(0)
    STRUCTURE_construct aStruct(blueprint_), 0, ownerPlayer
    aStruct(blueprint_).state = SAS_NORMAL
    ' construct the fake structure
    aStruct(blueprint_).pStructureType = psStats ' Couldn't be bothered to fix const correctness everywhere.
    If ULT(selectedPlayer, MAX_PLAYERS) Then
        aStruct(blueprint_).visible(selectedPlayer) = UBYTE_MAX
    End If
    n = asStructureStats(pIMDStats).pIMD_size - 1
    If moduleNumber < n Then n = moduleNumber
    aStruct(blueprint_).sDisplay.imd = asStructureStats(pIMDStats).pIMD(n)
    aStruct(blueprint_).pos = p
    aStruct(blueprint_).rot = rot
    aStruct(blueprint_).selected = 0

    aStruct(blueprint_).numWeaps = 0
    aStruct(blueprint_).asWeaps(0).nStat = 0

    ' give defensive structures a weapon
    If asStructureStats(psStats).psWeapStat(0) <> -1 Then
        aStruct(blueprint_).asWeaps(0).nStat = asStructureStats(psStats).psWeapStat(0)
    End If
    ' things with sensors or ecm (or repair facilities) need these set, even if they have no official weapon
    aStruct(blueprint_).numWeaps = 0
    aStruct(blueprint_).asWeaps(0).lastFired = 0
    aStruct(blueprint_).asWeaps(0).rot.pitch = 0
    aStruct(blueprint_).asWeaps(0).rot.direction = 0
    aStruct(blueprint_).asWeaps(0).rot.roll = 0
    aStruct(blueprint_).asWeaps(0).prevRot = aStruct(blueprint_).asWeaps(0).rot

    aStruct(blueprint_).expectedDamage = 0

    ' Times must be different, but don't otherwise matter.
    aStruct(blueprint_).time_ = 23
    aStruct(blueprint_).prevTime = 42

    aStruct(blueprint_).status = state

    ' Rotate wall if needed.
    If asStructureStats(aStruct(blueprint_).pStructureType).typ = REF_WALL Or asStructureStats(aStruct(blueprint_).pStructureType).typ = REF_GATE Then
        scanType = structChooseWallTypeBlueprint(map_coord(aStruct(blueprint_).pos.x), map_coord(aStruct(blueprint_).pos.y))
        typ = wallType(scanType)
        If scanType <> WallConnectNone Then
            aStruct(blueprint_).rot.direction = wallDir(scanType)
            n = asStructureStats(aStruct(blueprint_).pStructureType).pIMD_size - 1
            If typ < n Then n = typ
            aStruct(blueprint_).sDisplay.imd = asStructureStats(aStruct(blueprint_).pStructureType).pIMD(n)
        End If
    End If

    buildBlueprint = blueprint_
End Function
'/*
'fills the list with Structure that can be built. There is a limit on how many can
'be built at any one time. Pass back the number available.
'There is now a limit of how many of each type of structure are allowed per mission
'*/
' (VB) std::vector<STRUCTURE_STATS *>: the STRUCTURE_STATS indexes in structureList(0 .. n - 1), n returned
Public Function fillStructureList(ByVal selectedPlayer_ As Long, ByVal limit As Long, ByVal showFavorites As Boolean, structureList() As Long) As Long
    Dim structureList_size As Long, inc As Long, psBuilding As Long, it As Long, psCurr As Long
    Dim researchLabCurrMax(0 To 1) As Long, factoriesCurrMax(0 To 1) As Long, vtolFactoriesCurrMax(0 To 1) As Long, powerGenCurrMax(0 To 1) As Long
    Dim factoryModules As Long, powerGenModules As Long, researchModules As Long
    Dim counter As Long                 ' int8_t *counter: 1 research labs, 2 factories, 3 VTOL factories, 4 power generators

    Erase structureList
    structureList_size = 0

    If Not (selectedPlayer_ < MAX_PLAYERS) Then ASSERT_FAILURE "_selectedPlayer < MAX_PLAYERS", "src/structure.cpp:4052", "fillStructureList", "_selectedPlayer = %u", selectedPlayer_: fillStructureList = structureList_size: Exit Function

    ' counters for current nb of buildings, max buildings, current nb modules (int8_t)
    factoryModules = 0
    powerGenModules = 0
    researchModules = 0

    'if currently on a mission can't build factory/research/power/derricks
    If Not missionIsOffworld() Then
        it = apsStructLists(selectedPlayer_).head
        Do While it <> 0
            psCurr = lnObj(it)
            If asStructureStats(aStruct(psCurr).pStructureType).typ = REF_RESEARCH And aStruct(psCurr).status = SS_BUILT Then
                researchModules = int8(researchModules + aStruct(psCurr).capacity)
            ElseIf asStructureStats(aStruct(psCurr).pStructureType).typ = REF_FACTORY And aStruct(psCurr).status = SS_BUILT Then
                factoryModules = int8(factoryModules + aStruct(psCurr).capacity)
            ElseIf asStructureStats(aStruct(psCurr).pStructureType).typ = REF_POWER_GEN And aStruct(psCurr).status = SS_BUILT Then
                powerGenModules = int8(powerGenModules + aStruct(psCurr).capacity)
            ElseIf asStructureStats(aStruct(psCurr).pStructureType).typ = REF_VTOL_FACTORY And aStruct(psCurr).status = SS_BUILT Then
                ' same as REF_FACTORY
                factoryModules = int8(factoryModules + aStruct(psCurr).capacity)
            End If
            it = lnNext(it)
        Loop
    End If

    ' find maximum allowed limits (current built numbers already available, just grab them)
    For inc = 0 To numStructureStats - 1
        If apStructTypeLists(selectedPlayer_, inc) = AVAILABLE Or (includeRedundantDesigns And apStructTypeLists(selectedPlayer_, inc) = REDUNDANT) Then
            If asStructureStats(inc).typ = REF_RESEARCH Then
                counter = 1
            ElseIf asStructureStats(inc).typ = REF_FACTORY Then
                counter = 2
            ElseIf asStructureStats(inc).typ = REF_VTOL_FACTORY Then
                counter = 3
            ElseIf asStructureStats(inc).typ = REF_POWER_GEN Then
                counter = 4
            Else
                GoTo continue_limits
            End If
            Select Case counter
            Case 1
                researchLabCurrMax(0) = int8(asStructureStats(inc).curCount(selectedPlayer_))
                researchLabCurrMax(1) = int8(asStructureStats(inc).upgrade(selectedPlayer_).limit)
            Case 2
                factoriesCurrMax(0) = int8(asStructureStats(inc).curCount(selectedPlayer_))
                factoriesCurrMax(1) = int8(asStructureStats(inc).upgrade(selectedPlayer_).limit)
            Case 3
                vtolFactoriesCurrMax(0) = int8(asStructureStats(inc).curCount(selectedPlayer_))
                vtolFactoriesCurrMax(1) = int8(asStructureStats(inc).upgrade(selectedPlayer_).limit)
            Case 4
                powerGenCurrMax(0) = int8(asStructureStats(inc).curCount(selectedPlayer_))
                powerGenCurrMax(1) = int8(asStructureStats(inc).upgrade(selectedPlayer_).limit)
            End Select

        End If
continue_limits:
    Next

    debug_ 4120, LOG_NEVER, "fillStructureList", "structures: RL %i/%i (%i), F %i/%i (%i), VF %i/%i, PG %i/%i (%i)", _
                    researchLabCurrMax(0), researchLabCurrMax(1), researchModules, _
                    factoriesCurrMax(0), factoriesCurrMax(1), factoryModules, _
                    vtolFactoriesCurrMax(0), vtolFactoriesCurrMax(1), _
                    powerGenCurrMax(0), powerGenCurrMax(1), powerGenModules

    'set the list of Structures to build
    For inc = 0 To numStructureStats - 1
        'if the structure is flagged as available, add it to the list
        If apStructTypeLists(selectedPlayer_, inc) = AVAILABLE Or (includeRedundantDesigns And apStructTypeLists(selectedPlayer_, inc) = REDUNDANT) Then
            'check not built the maximum allowed already
            If asStructureStats(inc).curCount(selectedPlayer_) < asStructureStats(inc).upgrade(selectedPlayer_).limit Then
                psBuilding = inc

                'don't want corner wall to appear in list
                If asStructureStats(psBuilding).typ = REF_WALLCORNER Then
                    GoTo continue_list
                End If

                ' Remove the demolish stat from the list for tutorial
                ' tjc 4-dec-98  ...
                If bInTutorial Then
                    If asStructureStats(psBuilding).typ = REF_DEMOLISH Then
                        GoTo continue_list
                    End If
                End If

                'can't build list when offworld
                If missionIsOffworld() Then
                    Select Case asStructureStats(psBuilding).typ
                    Case REF_FACTORY, REF_COMMAND_CONTROL, REF_HQ, REF_POWER_GEN, REF_RESOURCE_EXTRACTOR, REF_RESEARCH, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
                        GoTo continue_list
                    End Select
                End If

                If asStructureStats(psBuilding).typ = REF_RESEARCH_MODULE Then
                    'don't add to list if Research Facility not presently built
                    'or if all labs already have a module
                    If researchLabCurrMax(0) = 0 Or researchModules >= researchLabCurrMax(1) Then
                        GoTo continue_list
                    End If
                ElseIf asStructureStats(psBuilding).typ = REF_FACTORY_MODULE Then
                    'don't add to list if Factory not presently built
                    'or if all factories already have all possible modules
                    If factoriesCurrMax(0) = 0 Or (factoryModules >= (factoriesCurrMax(1) + vtolFactoriesCurrMax(1)) * 2) Then
                        GoTo continue_list
                    End If
                ElseIf asStructureStats(psBuilding).typ = REF_POWER_MODULE Then
                    'don't add to list if Power Gen not presently built
                    'or if all generators already have a module
                    If powerGenCurrMax(0) = 0 Or (powerGenModules >= powerGenCurrMax(1)) Then
                        GoTo continue_list
                    End If
                End If
                ' show only favorites?
                If showFavorites And Not asStructureStats(inc).isFavorite Then
                    GoTo continue_list
                End If

                debug_ 4202, LOG_NEVER, "fillStructureList", "adding %s (%x)", asStructureStats(psBuilding).name_, apStructTypeLists(selectedPlayer_, inc)
                ReDim Preserve structureList(0 To structureList_size)
                structureList(structureList_size) = psBuilding
                structureList_size = structureList_size + 1
                If structureList_size = limit Then
                    fillStructureList = structureList_size
                    Exit Function
                End If
            End If
        End If
continue_list:
    Next
    fillStructureList = structureList_size
End Function

' (VB) an int8_t's wrap-around
Private Function int8(ByVal v As Long) As Long
    v = v And &HFF&
    If v >= &H80& Then v = v - &H100&
    int8 = v
End Function
