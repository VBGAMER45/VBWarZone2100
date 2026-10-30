Attribute VB_Name = "stats_c"
' src/stats.cpp, stats.h: the stores of the droid components' stats, the
' propulsion types, the terrain table and the weapon modifiers, their
' loaders, the lookups by name and the upgradeable values. A
' std::vector<X_STATS> asXStats is the array asXStats with its count
' asXStats_size. apCompLists[player][type] is apCompLists(player, type,
' comp), one array with the largest list's length (mechanics_c allocates
' it), and apStructTypeLists[player] is apStructTypeLists(player, stat).
' The lookups (std::unordered_map by name) are Scripting.Dictionary
' objects of handles.
' (VB) A BASE_STATS * or COMPONENT_STATS * is a handle: the stat's StatType
' plus its index in its array (0 = nullptr), which for every stat but a
' template equals its ref; a template's is STAT_TEMPLATE plus its index in
' template_c's pool. The BASE_STATS_<field> and COMPONENT_STATS_<field>
' properties at the end read and write a field through a handle (a
' RESEARCH's BASE_STATS::index is its base_index).
' Overloads: getWeaponSubClass(WEAPON_SUBCLASS) is getWeaponSubClass_name,
' getWeaponEffect(WEAPON_EFFECT) is getWeaponEffect_name.
' Not ported yet: loadImageForWeapSubclass (the notifications' images),
' map_WEAPON_EFFECT (StringToEnumMap, for the scripts).
Option Explicit

Private Const WEAPON_TIME As Long = 100

Private Const DEFAULT_DROID_RESISTANCE As Long = 150

'Values to fill apCompLists and apStructTypeLists. Not a bitfield, values are in case that helps with savegame compatibility.
Public Enum ItemAvailability
    AVAILABLE = 1               ' This item can be used to design droids.
    UNAVAILABLE = 2             ' The player does not know about this item.
    FOUND = 4                   ' This item has been found, but is unresearched.
    REDUNDANT = 10              ' The player no longer needs this item.
    REDUNDANT_UNAVAILABLE = 11  ' The player no longer needs this item.
    REDUNDANT_FOUND = 12        ' The player no longer needs this item.
End Enum

' The stores for the different stats
Public asBodyStats() As BODY_STATS
Public asBodyStats_size As Long
Public asBrainStats() As BRAIN_STATS
Public asBrainStats_size As Long
Public asPropulsionStats() As PROPULSION_STATS
Public asPropulsionStats_size As Long
Public asSensorStats() As SENSOR_STATS
Public asSensorStats_size As Long
Public asECMStats() As ECM_STATS
Public asECMStats_size As Long
Public asRepairStats() As REPAIR_STATS
Public asRepairStats_size As Long
Public asWeaponStats() As WEAPON_STATS
Public asWeaponStats_size As Long
Public asConstructStats() As CONSTRUCT_STATS
Public asConstructStats_size As Long
Public asPropulsionTypes() As PROPULSION_TYPES
Public asPropulsionTypes_size As Long
Private asTerrainTable() As Long            ' int *: [terrainType * PROPULSION_TYPE_NUM + propulsionType]

'used to hold the modifiers cross refd by weapon effect and propulsion type
Public asWeaponModifier(0 To WE_NUMEFFECTS - 1, 0 To PROPULSION_TYPE_NUM - 1) As Long    ' WEAPON_MODIFIER (uint16_t)
Public asWeaponModifierBody(0 To WE_NUMEFFECTS - 1, 0 To SIZE_NUM - 1) As Long

'stores for each players component states - can be either UNAVAILABLE, REDUNDANT, FOUND or AVAILABLE
Public apCompLists() As Byte                ' (player, type, comp)
Public apCompLists_count(0 To COMP_NUMCOMPONENTS - 1) As Long    ' each type's length; 0: nullptr

'store for each players Structure states
Public apStructTypeLists() As Byte          ' (player, stat)
Public apStructTypeLists_allocated As Boolean

Private lookupStatPtr As Object             ' std::unordered_map<WzString, BASE_STATS *>
Private lookupCompStatPtr As Object         ' std::unordered_map<WzString, COMPONENT_STATS *>
Private statModelLoadingFailures As Long

Private deprecatedModelNames As Object      ' deprecatedModelUpgrade's static map

' (VB) the lookups, made on first use
Private Sub InitLookups()
    If lookupStatPtr Is Nothing Then
        Set lookupStatPtr = CreateObject("Scripting.Dictionary")
        lookupStatPtr.CompareMode = vbBinaryCompare
        Set lookupCompStatPtr = CreateObject("Scripting.Dictionary")
        lookupCompStatPtr.CompareMode = vbBinaryCompare
    End If
End Sub

'**********************************************************************************
'*	Dealloc the extra storage tables
'**********************************************************************************
'Deallocate the storage assigned for the Propulsion Types table
Private Sub deallocPropulsionTypes()
    Erase asPropulsionTypes
    asPropulsionTypes_size = 0
End Sub

'dealloc the storage assigned for the terrain table
Private Sub deallocTerrainTable()
    Erase asTerrainTable
End Sub

'******************************************************************************
'*		Generic stats macros/functions
'******************************************************************************

'Deallocate all the stats assigned from input data
Public Function statsShutDown() As Boolean
    InitLookups
    lookupStatPtr.RemoveAll
    lookupCompStatPtr.RemoveAll

    Erase asWeaponStats: asWeaponStats_size = 0
    Erase asBrainStats: asBrainStats_size = 0
    Erase asPropulsionStats: asPropulsionStats_size = 0
    Erase asRepairStats: asRepairStats_size = 0
    Erase asConstructStats: asConstructStats_size = 0
    Erase asECMStats: asECMStats_size = 0
    Erase asSensorStats: asSensorStats_size = 0
    Erase asBodyStats: asBodyStats_size = 0
    deallocPropulsionTypes
    deallocTerrainTable

    statModelLoadingFailures = 0

    statsShutDown = True
End Function

Public Function getStatModelLoadingFailures() As Long
    getStatModelLoadingFailures = statModelLoadingFailures
End Function

'******************************************************************************
'*		Allocate stats functions
'******************************************************************************
' (ALLOC_STATS: the list cleared and resized, its new elements constructed)

' Allocate Weapon stats
Public Function statsAllocWeapons(ByVal numStats As Long) As Boolean
    Dim i As Long
    If Not (ULT(numStats, &H10000)) Then ASSERT_FAILURE "(numStats) < ~STAT_MASK + 1", "src/stats.cpp:138", "statsAllocWeapons", "Number of stats entries too large for WEAPON_STATS"
    Erase asWeaponStats
    asWeaponStats_size = numStats
    If numStats > 0 Then ReDim asWeaponStats(0 To numStats - 1)
    For i = 0 To numStats - 1
        WEAPON_STATS_construct asWeaponStats(i)
    Next
    statsAllocWeapons = True
End Function

' Allocate Body Stats
Public Function statsAllocBody(ByVal numStats As Long) As Boolean
    Dim i As Long
    If Not (ULT(numStats, &H10000)) Then ASSERT_FAILURE "(numStats) < ~STAT_MASK + 1", "src/stats.cpp:143", "statsAllocBody", "Number of stats entries too large for BODY_STATS"
    Erase asBodyStats
    asBodyStats_size = numStats
    If numStats > 0 Then ReDim asBodyStats(0 To numStats - 1)
    For i = 0 To numStats - 1
        BODY_STATS_construct asBodyStats(i)
    Next
    statsAllocBody = True
End Function

' Allocate Brain Stats
Public Function statsAllocBrain(ByVal numStats As Long) As Boolean
    Dim i As Long
    If Not (ULT(numStats, &H10000)) Then ASSERT_FAILURE "(numStats) < ~STAT_MASK + 1", "src/stats.cpp:148", "statsAllocBrain", "Number of stats entries too large for BRAIN_STATS"
    Erase asBrainStats
    asBrainStats_size = numStats
    If numStats > 0 Then ReDim asBrainStats(0 To numStats - 1)
    For i = 0 To numStats - 1
        BRAIN_STATS_construct asBrainStats(i)
    Next
    statsAllocBrain = True
End Function

' Allocate Propulsion Stats
Public Function statsAllocPropulsion(ByVal numStats As Long) As Boolean
    Dim i As Long
    If Not (ULT(numStats, &H10000)) Then ASSERT_FAILURE "(numStats) < ~STAT_MASK + 1", "src/stats.cpp:153", "statsAllocPropulsion", "Number of stats entries too large for PROPULSION_STATS"
    Erase asPropulsionStats
    asPropulsionStats_size = numStats
    If numStats > 0 Then ReDim asPropulsionStats(0 To numStats - 1)
    For i = 0 To numStats - 1
        PROPULSION_STATS_construct asPropulsionStats(i)
    Next
    statsAllocPropulsion = True
End Function

' Allocate Sensor Stats
Public Function statsAllocSensor(ByVal numStats As Long) As Boolean
    Dim i As Long
    If Not (ULT(numStats, &H10000)) Then ASSERT_FAILURE "(numStats) < ~STAT_MASK + 1", "src/stats.cpp:158", "statsAllocSensor", "Number of stats entries too large for SENSOR_STATS"
    Erase asSensorStats
    asSensorStats_size = numStats
    If numStats > 0 Then ReDim asSensorStats(0 To numStats - 1)
    For i = 0 To numStats - 1
        SENSOR_STATS_construct asSensorStats(i)
    Next
    statsAllocSensor = True
End Function

' Allocate Ecm Stats
Public Function statsAllocECM(ByVal numStats As Long) As Boolean
    Dim i As Long
    If Not (ULT(numStats, &H10000)) Then ASSERT_FAILURE "(numStats) < ~STAT_MASK + 1", "src/stats.cpp:163", "statsAllocECM", "Number of stats entries too large for ECM_STATS"
    Erase asECMStats
    asECMStats_size = numStats
    If numStats > 0 Then ReDim asECMStats(0 To numStats - 1)
    For i = 0 To numStats - 1
        ECM_STATS_construct asECMStats(i)
    Next
    statsAllocECM = True
End Function

' Allocate Repair Stats
Public Function statsAllocRepair(ByVal numStats As Long) As Boolean
    Dim i As Long
    If Not (ULT(numStats, &H10000)) Then ASSERT_FAILURE "(numStats) < ~STAT_MASK + 1", "src/stats.cpp:169", "statsAllocRepair", "Number of stats entries too large for REPAIR_STATS"
    Erase asRepairStats
    asRepairStats_size = numStats
    If numStats > 0 Then ReDim asRepairStats(0 To numStats - 1)
    For i = 0 To numStats - 1
        REPAIR_STATS_construct asRepairStats(i)
    Next
    statsAllocRepair = True
End Function

' Allocate Construct Stats
Public Function statsAllocConstruct(ByVal numStats As Long) As Boolean
    Dim i As Long
    If Not (ULT(numStats, &H10000)) Then ASSERT_FAILURE "(numStats) < ~STAT_MASK + 1", "src/stats.cpp:175", "statsAllocConstruct", "Number of stats entries too large for CONSTRUCT_STATS"
    Erase asConstructStats
    asConstructStats_size = numStats
    If numStats > 0 Then ReDim asConstructStats(0 To numStats - 1)
    For i = 0 To numStats - 1
        CONSTRUCT_STATS_construct asConstructStats(i)
    Next
    statsAllocConstruct = True
End Function

'******************************************************************************
'*		Load stats functions
'******************************************************************************

' optional<WzString>: True and the new name (empty: no replacement) when the model is deprecated
Private Function deprecatedModelUpgrade(filename As String, result As String) As Boolean
    Dim name_ As String, pairs As Variant, k As Long
    If deprecatedModelNames Is Nothing Then
        Set deprecatedModelNames = CreateObject("Scripting.Dictionary")
        deprecatedModelNames.CompareMode = vbBinaryCompare
        pairs = Array( _
            "mibnkdrl.pie", "mibnkdr.pie", "mibnkdrr.pie", "", "prhlhtr3.pie", "prhhtr3.pie", "prhrhtr3.pie", "", _
            "prhltrk3.pie", "prhtrk3.pie", "prhrtrk3.pie", "", "prhlvtl1.pie", "prhvtl1.pie", "prhrvtl1.pie", "", _
            "prhlvtl1.pie", "prhvtl1.pie", "prhrvtl1.pie", "", "prhlvtl2.pie", "prhvtl2.pie", "prhrvtl2.pie", "", _
            "prhlvtl3.pie", "prhvtl3.pie", "prhrvtl3.pie", "", "prhlvtl4.pie", "prhvtl4.pie", "prhrvtl4.pie", "", _
            "prhlwhl1.pie", "prhwhl1.pie", "prhrwhl1.pie", "", "prllhtr1.pie", "prlhtr1.pie", "prlrhtr1.pie", "", _
            "prlltrk1.pie", "prltrk1.pie", "prlrtrk1.pie", "", "prllvtl1.pie", "prlvtl1.pie", "prlrvtl1.pie", "", _
            "prllvtl2.pie", "prlvtl2.pie", "prlrvtl2.pie", "", "prllvtl3.pie", "prlvtl3.pie", "prlrvtl3.pie", "", _
            "prllwhl1.pie", "prlwhl1.pie", "prlrwhl1.pie", "", "prmlhtr2.pie", "prmhtr2.pie", "prmrhtr2.pie", "", _
            "prmltrk2.pie", "prmtrk2.pie", "prmrtrk2.pie", "", "prmlvtl1.pie", "prmvtl1.pie", "prmrvtl1.pie", "", _
            "prmlwhl1.pie", "prmwhl1.pie", "prmrwhl1.pie", "", _
            "prslwhl1.pie", "prswhl1.pie", "prsrwhl1.pie", "", "prslvtl1.pie", "prsvtl1.pie", "prsrvtl1.pie", "", _
            "prsltrk4.pie", "prstrk4.pie", "prsrtrk4.pie", "", "prslhtr4.pie", "prshtr4.pie", "prsrhtr4.pie", "")
        ' base (the first 76 names), then mp; a repeated key is ignored, as by the initializer list
        For k = 0 To UBound(pairs) - 1 Step 2
            If Not deprecatedModelNames.Exists(pairs(k)) Then deprecatedModelNames.Add pairs(k), pairs(k + 1)
        Next
    End If

    name_ = WzString_toLower(filename)
    If deprecatedModelNames.Exists(name_) Then
        debug_ 239, LOG_INFO, "deprecatedModelUpgrade", "Deprecation notice: The model ""%s"" is referenced in a mod's stats JSON file, and is deprecated. Please check the latest base stats models for how to update this mod for future compatibility.", name_
        result = deprecatedModelNames(name_)
        deprecatedModelUpgrade = True
        Exit Function
    End If
    deprecatedModelUpgrade = False
End Function

' psStats: a BASE_STATS * handle; the model handle, 0: nullptr
Private Function statsGetIMD(ByVal json As WzConfig, ByVal psStats As Long, ByVal key As String, Optional ByVal key2 As String = "") As Long
    Dim retval As Long, value As nlohmann_json, obj As nlohmann_json, filename As String, upgraded As String
    retval = 0
    If json.contains(key) Then
        Set value = json.json(key)
        If value.is_object() Then
            If Not (Len(key2) > 0) Then ASSERT_FAILURE "!key2.isEmpty()", "src/stats.cpp:253", "statsGetIMD", "Cannot look up a JSON object with an empty key!": statsGetIMD = 0: Exit Function
            Set obj = value
            If obj.find(key2) Is Nothing Then
                statsGetIMD = 0
                Exit Function
            End If
            Set value = obj.item(key2)
        End If
        filename = json_variant_new(value).toWzString()
        deprecatedModelUpgrade filename, upgraded ' FUTURE TODO: Use output of this to auto-upgrade old model references? (Check for empty string)
        retval = modelGet(filename)
        If retval = 0 Then
            statModelLoadingFailures = statModelLoadingFailures + 1
        End If
        If Not (retval <> 0) Then ASSERT_FAILURE "retval != nullptr", "src/stats.cpp:268", "statsGetIMD", "Cannot find the PIE model %s for stat %s in %s", _
            filename, BASE_STATS_name(psStats), json.fileName()
    End If
    statsGetIMD = retval
End Function

' psStats: a BASE_STATS * handle
Private Sub loadStats(ByVal json As WzConfig, ByVal psStats As Long, ByVal index As Long)
    Dim id As String
    InitLookups
    BASE_STATS_id(psStats) = json.group()
    BASE_STATS_name(psStats) = json.string_("name")
    BASE_STATS_index(psStats) = index
    id = BASE_STATS_id(psStats)
    If lookupStatPtr.Exists(id) Then ASSERT_FAILURE "lookupStatPtr.find(psStats->id) == lookupStatPtr.end()", "src/stats.cpp:279", "loadStats", "Duplicate ID found! (%s)", id
    If Not lookupStatPtr.Exists(id) Then lookupStatPtr.Add id, psStats      ' insert: an existing key stays
End Sub

' psStats: the STRUCTURE_STATS's index
Public Sub loadStructureStats_BaseStats(ByVal json As WzConfig, ByVal psStats As Long, ByVal index As Long)
    loadStats json, STAT_STRUCTURE + psStats, index
End Sub

Public Sub unloadStructureStats_BaseStats(psStats As STRUCTURE_STATS)
    InitLookups
    If lookupStatPtr.Exists(psStats.id) Then lookupStatPtr.Remove psStats.id
End Sub

' psStats: a COMPONENT_STATS * handle
Private Sub loadCompStats(ByVal json As WzConfig, ByVal psStats As Long, ByVal index As Long)
    Dim useType As String, dtype As String, id As String
    loadStats json, psStats, index
    id = BASE_STATS_id(psStats)
    If Not lookupCompStatPtr.Exists(id) Then lookupCompStatPtr.Add id, psStats
    COMPONENT_STATS_buildPower(psStats) = json.value("buildPower", json_variant_new(0)).toUInt()
    COMPONENT_STATS_buildPoints(psStats) = json.value("buildPoints", json_variant_new(0)).toUInt()
    COMPONENT_STATS_designable(psStats) = json.value("designable", json_variant_new(False)).toBool()
    COMPONENT_STATS_weight(psStats) = json.value("weight", json_variant_new(0)).toUInt()
    useType = json.value("usageClass", json_variant_new("")).toWzString()
    If InStr(useType, "SuperCyborg") <> 0 Then
        COMPONENT_STATS_usageClass(psStats) = UsageClass_SuperCyborg
    ElseIf InStr(useType, "Cyborg") <> 0 Then
        COMPONENT_STATS_usageClass(psStats) = UsageClass_Cyborg
    End If
    COMPONENT_STATS_getBase_hitpoints(psStats) = json.value("hitpoints", json_variant_new(0)).toUInt()
    COMPONENT_STATS_getBase_hitpointPct(psStats) = json.value("hitpointPct", json_variant_new(100)).toUInt()

    dtype = json.value("droidType", json_variant_new("DROID")).toWzString()
    COMPONENT_STATS_droidTypeOverride(psStats) = DROID_ANY
    If dtype = "PERSON" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_PERSON
    ElseIf dtype = "TRANSPORTER" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_TRANSPORTER
    ElseIf dtype = "SUPERTRANSPORTER" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_SUPERTRANSPORTER
    ElseIf dtype = "CYBORG" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_CYBORG
    ElseIf dtype = "CYBORG_SUPER" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_CYBORG_SUPER
    ElseIf dtype = "CYBORG_CONSTRUCT" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_CYBORG_CONSTRUCT
    ElseIf dtype = "CYBORG_REPAIR" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_CYBORG_REPAIR
    ElseIf dtype = "DROID_CONSTRUCT" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_CONSTRUCT
    ElseIf dtype = "DROID_ECM" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_ECM
    ElseIf dtype = "DROID_COMMAND" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_COMMAND
    ElseIf dtype = "DROID_SENSOR" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_SENSOR
    ElseIf dtype = "DROID_REPAIR" Then
        COMPONENT_STATS_droidTypeOverride(psStats) = DROID_REPAIR
    ElseIf dtype <> "DROID" Then
        debug_ 365, LOG_ERROR, "loadCompStats", "Unrecognized droidType %s", dtype
    End If
End Sub

' (VB) std::find of name in list(0 .. n - 1): its index, or n (end())
Private Function FindInList(list() As String, ByVal n As Long, ByVal name_ As String) As Long
    Dim k As Long
    For k = 0 To n - 1
        If list(k) = name_ Then Exit For
    Next
    FindInList = k
End Function

' (VB) std::iter_swap(list.begin() + k, list.begin())
Private Sub SwapToFront(list() As String, ByVal k As Long)
    Dim t As String
    t = list(k)
    list(k) = list(0)
    list(0) = t
End Sub

'Load the weapon stats from the file exported from Access
Public Function loadWeaponStats(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, nullweapon As Long, i As Long, j As Long
    Dim flags_raw() As String, nFlags As Long, flags As StringVector, weaponSize As Long
    Dim periodicalDamageWeaponClass As String, periodicalDamageWeaponSubClass As String, periodicalDamageWeaponEffect As String
    Dim shortRange As Long, longRange As Long, shortHit As Long, longHit As Long
    Dim weaponSoundID As Long, explosionSoundID As Long, szWeaponWav As String, szExplosionWav As String, result As Boolean

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:372", "loadWeaponStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    statsAllocWeapons n
    ' Hack to make sure ZNULLWEAPON is always first in list
    nullweapon = FindInList(list, n, "ZNULLWEAPON")
    If Not (nullweapon <> n) Then ASSERT_FAILURE "nullweapon != list.end()", "src/stats.cpp:377", "loadWeaponStats", "ZNULLWEAPON is mandatory": loadWeaponStats = False: Exit Function
    SwapToFront list, nullweapon
    For i = 0 To n - 1
        StringVector_clear flags

        ini.beginGroup list(i)
        loadCompStats ini, STAT_WEAPON + i, i
        asWeaponStats(i).compType = COMP_WEAPON

        asWeaponStats(i).radiusLife = ini.value("radiusLife", json_variant_new(0)).toUInt()

        asWeaponStats(i).base.shortRange = ini.value("shortRange").toUInt()
        asWeaponStats(i).base.maxRange = ini.value("longRange").toUInt()
        asWeaponStats(i).base.minRange = ini.value("minRange", json_variant_new(0)).toUInt()
        asWeaponStats(i).base.hitChance = ini.value("longHit", json_variant_new(100)).toUInt()
        asWeaponStats(i).base.shortHitChance = ini.value("shortHit", json_variant_new(100)).toUInt()
        asWeaponStats(i).base.firePause = ini.value("firePause").toUInt()
        asWeaponStats(i).base.numRounds = ini.value("numRounds").toUInt() And &HFF&
        asWeaponStats(i).base.reloadTime = ini.value("reloadTime").toUInt()
        asWeaponStats(i).base.damage = ini.value("damage").toUInt()
        asWeaponStats(i).base.minimumDamage = ini.value("minimumDamage", json_variant_new(0)).toInt()
        asWeaponStats(i).base.empRadius = ini.value("empRadius", json_variant_new(0)).toUInt()
        asWeaponStats(i).base.radius = ini.value("radius", json_variant_new(0)).toUInt()
        asWeaponStats(i).base.radiusDamage = ini.value("radiusDamage", json_variant_new(0)).toUInt()
        asWeaponStats(i).base.periodicalDamageTime = ini.value("periodicalDamageTime", json_variant_new(0)).toUInt()
        asWeaponStats(i).base.periodicalDamage = ini.value("periodicalDamage", json_variant_new(0)).toUInt()
        asWeaponStats(i).base.periodicalDamageRadius = ini.value("periodicalDamageRadius", json_variant_new(0)).toUInt()
        ' multiply time stats
        asWeaponStats(i).base.firePause = asWeaponStats(i).base.firePause * WEAPON_TIME
        asWeaponStats(i).base.periodicalDamageTime = asWeaponStats(i).base.periodicalDamageTime * WEAPON_TIME
        asWeaponStats(i).radiusLife = asWeaponStats(i).radiusLife * WEAPON_TIME
        asWeaponStats(i).base.reloadTime = asWeaponStats(i).base.reloadTime * WEAPON_TIME
        ' copy for upgrades
        For j = 0 To MAX_PLAYERS - 1
            asWeaponStats(i).upgrade(j) = asWeaponStats(i).base
        Next

        asWeaponStats(i).numExplosions = ini.value("numExplosions").toUInt()
        asWeaponStats(i).flightSpeed = ini.value("flightSpeed", json_variant_new(1)).toUInt()
        asWeaponStats(i).rotate = ToInt16(ini.value("rotate").toUInt())
        asWeaponStats(i).minElevation = ToInt16(ini.value("minElevation").toInt())
        asWeaponStats(i).maxElevation = ToInt16(ini.value("maxElevation").toInt())
        asWeaponStats(i).recoilValue = ini.value("recoilValue").toUInt()
        asWeaponStats(i).effectSize = ini.value("effectSize").toUInt() And &HFFFF&
        nFlags = ini.value("flags", json_variant_new(0)).toWzStringList(flags_raw)
        ' convert flags entries to lowercase
        For j = 0 To nFlags - 1
            StringVector_push_back flags, WzString_toLower(flags_raw(j))
        Next
        asWeaponStats(i).vtolAttackRuns = ToInt16(ini.value("numAttackRuns", json_variant_new(0)).toUInt())
        asWeaponStats(i).penetrate = ini.value("penetrate", json_variant_new(False)).toBool()
        ' weapon size limitation
        weaponSize = ini.value("weaponSize", json_variant_new(WEAPON_SIZE_ANY)).toInt()
        If Not (weaponSize <= WEAPON_SIZE_ANY) Then ASSERT_FAILURE "weaponSize <= WEAPON_SIZE_ANY", "src/stats.cpp:436", "loadWeaponStats", "Bad weapon size for %s", list(i)
        asWeaponStats(i).weaponSize = weaponSize

        If Not (asWeaponStats(i).flightSpeed <> 0) Then ASSERT_FAILURE "psStats->flightSpeed > 0", "src/stats.cpp:439", "loadWeaponStats", "Invalid flight speed for %s", list(i)

        asWeaponStats(i).ref = STAT_WEAPON + i

        'get the IMD for the component
        asWeaponStats(i).pIMD = statsGetIMD(ini, STAT_WEAPON + i, "model")
        asWeaponStats(i).pMountGraphic = statsGetIMD(ini, STAT_WEAPON + i, "mountModel")
        If GetGameMode() = GS_NORMAL Then
            asWeaponStats(i).pMuzzleGraphic = statsGetIMD(ini, STAT_WEAPON + i, "muzzleGfx")
            asWeaponStats(i).pInFlightGraphic = statsGetIMD(ini, STAT_WEAPON + i, "flightGfx")
            asWeaponStats(i).pTargetHitGraphic = statsGetIMD(ini, STAT_WEAPON + i, "hitGfx")
            asWeaponStats(i).pTargetMissGraphic = statsGetIMD(ini, STAT_WEAPON + i, "missGfx")
            asWeaponStats(i).pWaterHitGraphic = statsGetIMD(ini, STAT_WEAPON + i, "waterGfx")
            asWeaponStats(i).pTrailGraphic = statsGetIMD(ini, STAT_WEAPON + i, "trailGfx")
        End If
        asWeaponStats(i).fireOnMove = ini.value("fireOnMove", json_variant_new(True)).toBool()

        'set the weapon class
        If Not getWeaponClass(ini.value("weaponClass").toWzString(), asWeaponStats(i).weaponClass) Then
            debug_ 460, LOG_ERROR, "loadWeaponStats", "Invalid weapon class for weapon %s - assuming KINETIC", asWeaponStats(i).name_
            asWeaponStats(i).weaponClass = WC_KINETIC
        End If

        'set the subClass
        If Not getWeaponSubClass(ini.value("weaponSubClass").toWzString(), asWeaponStats(i).weaponSubClass) Then
            loadWeaponStats = False
            Exit Function
        End If

        ' set max extra weapon range on misses, make this modifiable one day by mod makers
        If asWeaponStats(i).weaponSubClass = WSC_MGUN Or asWeaponStats(i).weaponSubClass = WSC_COMMAND Then
            asWeaponStats(i).distanceExtensionFactor = 120
        ElseIf asWeaponStats(i).weaponSubClass = WSC_AAGUN Then
            asWeaponStats(i).distanceExtensionFactor = 100
        Else ' default
            asWeaponStats(i).distanceExtensionFactor = 150
        End If

        'set the weapon effect
        If Not getWeaponEffect(ini.value("weaponEffect").toWzString(), asWeaponStats(i).weaponEffect) Then
            debug_ 487, LOG_FATAL, "loadWeaponStats", "loadWepaonStats: Invalid weapon effect for weapon %s", asWeaponStats(i).name_
            loadWeaponStats = False
            Exit Function
        End If

        'set periodical damage weapon class
        periodicalDamageWeaponClass = ini.value("periodicalDamageWeaponClass", json_variant_new("")).toWzString()
        If periodicalDamageWeaponClass = "" Then
            'was not setted in ini - use default value
            asWeaponStats(i).periodicalDamageWeaponClass = asWeaponStats(i).weaponClass
        ElseIf Not getWeaponClass(periodicalDamageWeaponClass, asWeaponStats(i).periodicalDamageWeaponClass) Then
            debug_ 500, LOG_ERROR, "loadWeaponStats", "Invalid periodicalDamageWeaponClass for weapon %s - assuming same class as weapon", asWeaponStats(i).name_
            asWeaponStats(i).periodicalDamageWeaponClass = asWeaponStats(i).weaponClass
        End If

        'set periodical damage weapon subclass
        periodicalDamageWeaponSubClass = ini.value("periodicalDamageWeaponSubClass", json_variant_new("")).toWzString()
        If periodicalDamageWeaponSubClass = "" Then
            'was not setted in ini - use default value
            asWeaponStats(i).periodicalDamageWeaponSubClass = asWeaponStats(i).weaponSubClass
        ElseIf Not getWeaponSubClass(periodicalDamageWeaponSubClass, asWeaponStats(i).periodicalDamageWeaponSubClass) Then
            debug_ 513, LOG_ERROR, "loadWeaponStats", "Invalid periodicalDamageWeaponSubClass for weapon %s - assuming same subclass as weapon", asWeaponStats(i).name_
            asWeaponStats(i).periodicalDamageWeaponSubClass = asWeaponStats(i).weaponSubClass
        End If

        'set periodical damage weapon effect
        periodicalDamageWeaponEffect = ini.value("periodicalDamageWeaponEffect", json_variant_new("")).toWzString()
        If periodicalDamageWeaponEffect = "" Then
            'was not setted in ini - use default value
            asWeaponStats(i).periodicalDamageWeaponEffect = asWeaponStats(i).weaponEffect
        ElseIf Not getWeaponEffect(periodicalDamageWeaponEffect, asWeaponStats(i).periodicalDamageWeaponEffect) Then
            debug_ 526, LOG_ERROR, "loadWeaponStats", "Invalid periodicalDamageWeaponEffect for weapon %s - assuming same effect as weapon", asWeaponStats(i).name_
            asWeaponStats(i).periodicalDamageWeaponEffect = asWeaponStats(i).weaponEffect
        End If

        'set the movement model
        If Not getMovementModel(ini.value("movement").toWzString(), asWeaponStats(i).movementModel) Then
            loadWeaponStats = False
            Exit Function
        End If

        shortRange = asWeaponStats(i).upgrade(0).shortRange
        longRange = asWeaponStats(i).upgrade(0).maxRange
        shortHit = asWeaponStats(i).upgrade(0).shortHitChance
        longHit = asWeaponStats(i).upgrade(0).hitChance
        If ULT(longRange, shortRange) Then
            debug_ 542, LOG_ERROR, "loadWeaponStats", "%s, Short range (%d) is greater than long range (%d)", asWeaponStats(i).name_, shortRange, longRange
        End If
        If shortRange = longRange And shortHit <> longHit Then
            debug_ 546, LOG_ERROR, "loadWeaponStats", "%s, shortHit and longHit should be equal if the ranges are the same", asWeaponStats(i).name_
        End If

        ' set the face Player value
        asWeaponStats(i).facePlayer = -CLng(ini.value("facePlayer", json_variant_new(False)).toBool())

        ' set the In flight face Player value
        asWeaponStats(i).faceInFlight = -CLng(ini.value("faceInFlight", json_variant_new(False)).toBool())

        'set the light world value
        asWeaponStats(i).lightWorld = ini.value("lightWorld", json_variant_new(False)).toBool()

        ' interpret flags
        asWeaponStats(i).surfaceToAir = SHOOT_ON_GROUND ' default
        If StringVector_contains(flags, "aironly") Then ' "AirOnly"
            asWeaponStats(i).surfaceToAir = SHOOT_IN_AIR
        ElseIf StringVector_contains(flags, "shootair") Then ' "ShootAir"
            asWeaponStats(i).surfaceToAir = asWeaponStats(i).surfaceToAir Or SHOOT_IN_AIR
        End If
        If StringVector_contains(flags, "nofriendlyfire") Then ' "NoFriendlyFire"
            asWeaponStats(i).flags = asWeaponStats(i).flags Or BIT_(WEAPON_FLAG_NO_FRIENDLY_FIRE)
        End If
        If StringVector_contains(flags, "allowedontransporter") Then ' "AllowedOnTransporter"
            asWeaponStats(i).flags = asWeaponStats(i).flags Or BIT_(WEAPON_FLAG_ALLOWED_ON_TRANSPORTER)
        End If
        If StringVector_contains(flags, "teleportcapture") Then ' "TeleportCapture"
            asWeaponStats(i).flags = asWeaponStats(i).flags Or BIT_(WEAPON_FLAG_TELEPORT_CAPTURE)
        End If

        'set the weapon sounds to default value
        asWeaponStats(i).iAudioFireID = NO_SOUND
        asWeaponStats(i).iAudioImpactID = NO_SOUND

        ' load sounds
        szWeaponWav = ini.value("weaponWav", json_variant_new("-1")).toWzString()
        szExplosionWav = ini.value("explosionWav", json_variant_new("-1")).toWzString()
        result = statsGetAudioIDFromString(list(i), szWeaponWav, weaponSoundID)
        If Not (result) Then ASSERT_FAILURE "result", "src/stats.cpp:590", "loadWeaponStats", "Weapon sound %s not found for %s", szWeaponWav, asWeaponStats(i).name_: loadWeaponStats = False: Exit Function
        result = statsGetAudioIDFromString(list(i), szExplosionWav, explosionSoundID)
        If Not (result) Then ASSERT_FAILURE "result", "src/stats.cpp:592", "loadWeaponStats", "Explosion sound %s not found for %s", szExplosionWav, asWeaponStats(i).name_: loadWeaponStats = False: Exit Function
        asWeaponStats(i).iAudioFireID = weaponSoundID
        asWeaponStats(i).iAudioImpactID = explosionSoundID

        ini.endGroup
    Next

    loadWeaponStats = True
End Function

' (VB) std::vector<iIMDBaseShape *>::resize(n, nullptr)
Private Sub ResizeIMDList(list() As Long, size As Long, ByVal n As Long)
    Dim k As Long
    If n = 0 Then
        Erase list
    Else
        ReDim Preserve list(0 To n - 1)
        For k = size To n - 1
            list(k) = 0
        Next
    End If
    size = n
End Sub

Public Function loadBodyStats(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, nullbody As Long, i As Long, j As Long, numStats As Long
    Dim psBodyStat As Long, keys() As String, nKeys As Long

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:604", "loadBodyStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    statsAllocBody n
    ' Hack to make sure ZNULLBODY is always first in list
    nullbody = FindInList(list, n, "ZNULLBODY")
    If Not (nullbody <> n) Then ASSERT_FAILURE "nullbody != list.end()", "src/stats.cpp:609", "loadBodyStats", "ZNULLBODY is mandatory": loadBodyStats = False: Exit Function
    SwapToFront list, nullbody
    For i = 0 To n - 1
        ini.beginGroup list(i)
        loadCompStats ini, STAT_BODY + i, i
        asBodyStats(i).compType = COMP_BODY

        asBodyStats(i).weaponSlots = ini.value("weaponSlots").toInt()
        asBodyStats(i).bodyClass = ini.value("class").toWzString()
        asBodyStats(i).base.thermal = ini.value("armourHeat").toInt()
        asBodyStats(i).base.armour = ini.value("armourKinetic").toInt()
        asBodyStats(i).base.power = ini.value("powerOutput").toInt()
        asBodyStats(i).base.resistance = ini.value("resistance", json_variant_new(DEFAULT_DROID_RESISTANCE)).toInt()
        For j = 0 To MAX_PLAYERS - 1
            asBodyStats(i).upgrade(j) = asBodyStats(i).base
        Next
        asBodyStats(i).ref = STAT_BODY + i
        If Not getBodySize(ini.value("size").toWzString(), asBodyStats(i).size) Then
            ASSERT_FAILURE "false", "src/stats.cpp:632", "loadBodyStats", "Unknown body size for %s", asBodyStats(i).name_
            loadBodyStats = False
            Exit Function
        End If
        asBodyStats(i).pIMD = statsGetIMD(ini, STAT_BODY + i, "model")

        ini.endGroup
    Next

    ' now get the extra stuff ... hack it together with above later, moved here from
    ' separate function

    ' allocate space
    For numStats = 0 To asBodyStats_size - 1
        ResizeIMDList asBodyStats(numStats).ppIMDList, asBodyStats(numStats).ppIMDList_size, asPropulsionStats_size * NUM_PROP_SIDES
        ResizeIMDList asBodyStats(numStats).ppMoveIMDList, asBodyStats(numStats).ppMoveIMDList_size, asPropulsionStats_size * NUM_PROP_SIDES
        ResizeIMDList asBodyStats(numStats).ppStillIMDList, asBodyStats(numStats).ppStillIMDList_size, asPropulsionStats_size * NUM_PROP_SIDES
    Next
    For i = 0 To n - 1
        psBodyStat = -1

        ini.beginGroup list(i)
        If Not ini.contains("propulsionExtraModels") Then
            ini.endGroup
            GoTo continue_
        End If
        ini.beginGroup "propulsionExtraModels"
        'get the body stats
        For numStats = 0 To asBodyStats_size - 1
            psBodyStat = numStats
            If list(i) = asBodyStats(psBodyStat).id Then
                Exit For
            End If
        Next
        If numStats = asBodyStats_size Then ' not found
            debug_ 675, LOG_FATAL, "loadBodyStats", "Invalid body name %s", list(i)
            loadBodyStats = False
            Exit Function
        End If
        nKeys = ini.childKeys(keys)
        For j = 0 To nKeys - 1
            For numStats = 0 To asPropulsionStats_size - 1
                If keys(j) = asPropulsionStats(numStats).id Then
                    Exit For
                End If
            Next
            If numStats = asPropulsionStats_size Then
                debug_ 691, LOG_FATAL, "loadBodyStats", "Invalid propulsion name %s", keys(j)
                loadBodyStats = False
                Exit Function
            End If
            'allocate the left and right propulsion IMDs + movement and standing still animations
            ' NOTE: New base stats files + IMDs use a single combined model (in the "left" key) for both sides...
            '       however, older mods / stats files may still rely on having both left + right
            asBodyStats(psBodyStat).ppIMDList(numStats * NUM_PROP_SIDES + LEFT_PROP) = statsGetIMD(ini, STAT_BODY + psBodyStat, keys(j), "left")
            asBodyStats(psBodyStat).ppIMDList(numStats * NUM_PROP_SIDES + RIGHT_PROP) = statsGetIMD(ini, STAT_BODY + psBodyStat, keys(j), "right")
            asBodyStats(psBodyStat).ppMoveIMDList(numStats) = statsGetIMD(ini, STAT_BODY + psBodyStat, keys(j), "moving")
            asBodyStats(psBodyStat).ppStillIMDList(numStats) = statsGetIMD(ini, STAT_BODY + psBodyStat, keys(j), "still")
        Next
        ini.endGroup
        ini.endGroup
continue_:
    Next

    loadBodyStats = True
End Function

'Load the Brain stats from the file exported from Access
Public Function loadBrainStats(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, nullbrain As Long, i As Long, j As Long, k As Long, retVal As Boolean
    Dim cmdExpRange As nlohmann_json, rankNames As nlohmann_json, rankThresholds As nlohmann_json, weapon_ As Long

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:712", "loadBrainStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    statsAllocBrain n
    ' Hack to make sure ZNULLBRAIN is always first in list
    nullbrain = FindInList(list, n, "ZNULLBRAIN")
    If Not (nullbrain <> n) Then ASSERT_FAILURE "nullbrain != list.end()", "src/stats.cpp:717", "loadBrainStats", "ZNULLBRAIN is mandatory": loadBrainStats = False: Exit Function
    SwapToFront list, nullbrain
    retVal = True
    For i = 0 To n - 1
        ini.beginGroup list(i)
        loadCompStats ini, STAT_BRAIN + i, i
        asBrainStats(i).compType = COMP_BRAIN

        asBrainStats(i).weight = ini.value("weight", json_variant_new(0)).toInt()
        asBrainStats(i).base.maxDroids = ini.value("maxDroids").toInt()
        asBrainStats(i).base.maxDroidsMult = ini.value("maxDroidsMult").toInt()

        Set cmdExpRange = ini.json("cmdExpRange")
        If Not cmdExpRange.is_null() Then
            If Not (cmdExpRange.is_array()) Then ASSERT_FAILURE "cmdExpRange.is_array()", "src/stats.cpp:735", "loadBrainStats", "cmdExpRange is not an array"
            For k = 0 To cmdExpRange.size() - 1
                ReDim Preserve asBrainStats(i).cmdExpRange(0 To asBrainStats(i).cmdExpRange_size)
                asBrainStats(i).cmdExpRange(asBrainStats(i).cmdExpRange_size) = cmdExpRange.value_at(k).get_int()
                asBrainStats(i).cmdExpRange_size = asBrainStats(i).cmdExpRange_size + 1
            Next
        End If

        Set rankNames = ini.json("ranks")
        If Not (rankNames.is_array()) Then ASSERT_FAILURE "rankNames.is_array()", "src/stats.cpp:743", "loadBrainStats", "ranks is not an array"
        For k = 0 To rankNames.size() - 1
            StringVector_push_back asBrainStats(i).rankNames, rankNames.value_at(k).get_string()
        Next
        Set rankThresholds = ini.json("thresholds")
        For k = 0 To rankThresholds.size() - 1
            ReDim Preserve asBrainStats(i).base.rankThresholds(0 To asBrainStats(i).base.rankThresholds_size)
            asBrainStats(i).base.rankThresholds(asBrainStats(i).base.rankThresholds_size) = rankThresholds.value_at(k).get_int()
            asBrainStats(i).base.rankThresholds_size = asBrainStats(i).base.rankThresholds_size + 1
        Next
        asBrainStats(i).ref = STAT_BRAIN + i

        For j = 0 To MAX_PLAYERS - 1
            asBrainStats(i).upgrade(j) = asBrainStats(i).base
        Next

        ' check weapon attached
        asBrainStats(i).psWeaponStat = -1
        If ini.contains("turret") Then
            weapon_ = getCompFromName(COMP_WEAPON, ini.value("turret").toWzString())
            If weapon_ >= 0 Then
                asBrainStats(i).psWeaponStat = weapon_
            Else
                If Not (weapon_ >= 0) Then ASSERT_FAILURE "weapon >= 0", "src/stats.cpp:771", "loadBrainStats", "Unable to find weapon for brain %s", asBrainStats(i).name_
                retVal = False
            End If
        End If
        asBrainStats(i).designable = ini.value("designable", json_variant_new(False)).toBool()
        ini.endGroup
    Next
    loadBrainStats = retVal
End Function


'returns the propulsion type based on the string name passed in
Public Function getPropulsionType(ByVal typeName As String, typ As PROPULSION_TYPE) As Boolean
    If typeName = "Wheeled" Then
        typ = PROPULSION_TYPE_WHEELED
    ElseIf typeName = "Tracked" Then
        typ = PROPULSION_TYPE_TRACKED
    ElseIf typeName = "Legged" Then
        typ = PROPULSION_TYPE_LEGGED
    ElseIf typeName = "Hover" Then
        typ = PROPULSION_TYPE_HOVER
    ElseIf typeName = "Lift" Then
        typ = PROPULSION_TYPE_LIFT
    ElseIf typeName = "Propellor" Then
        typ = PROPULSION_TYPE_PROPELLOR
    ElseIf typeName = "Half-Tracked" Then
        typ = PROPULSION_TYPE_HALF_TRACKED
    Else
        debug_ 815, LOG_ERROR, "getPropulsionType", "getPropulsionType: Invalid Propulsion type %s - assuming Hover", typeName
        typ = PROPULSION_TYPE_HOVER

        getPropulsionType = False
        Exit Function
    End If

    getPropulsionType = True
End Function

'Load the Propulsion stats from the file exported from Access
Public Function loadPropulsionStats(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, nullprop As Long, i As Long, j As Long, retVal As Boolean

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:827", "loadPropulsionStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    statsAllocPropulsion n
    ' Hack to make sure ZNULLPROP is always first in list
    nullprop = FindInList(list, n, "ZNULLPROP")
    If Not (nullprop <> n) Then ASSERT_FAILURE "nullprop != list.end()", "src/stats.cpp:832", "loadPropulsionStats", "ZNULLPROP is mandatory": loadPropulsionStats = False: Exit Function
    SwapToFront list, nullprop
    retVal = True
    For i = 0 To n - 1
        ini.beginGroup list(i)
        loadCompStats ini, STAT_PROPULSION + i, i
        asPropulsionStats(i).compType = COMP_PROPULSION

        asPropulsionStats(i).base.hitpointPctOfBody = ini.value("hitpointPctOfBody", json_variant_new(0)).toInt()
        asPropulsionStats(i).maxSpeed = ini.value("speed").toInt()
        asPropulsionStats(i).ref = STAT_PROPULSION + i
        asPropulsionStats(i).turnSpeed = ini.value("turnSpeed", json_variant_new(DEG(1) \ 3)).toInt()
        If Not (asPropulsionStats(i).turnSpeed <> 0) Then ASSERT_FAILURE "psStats->turnSpeed != 0", "src/stats.cpp:847", "loadPropulsionStats", """%s"".""turnSpeed"" is 0", asPropulsionStats(i).id
        asPropulsionStats(i).spinSpeed = ini.value("spinSpeed", json_variant_new(DEG(3) \ 4)).toInt()
        If Not (asPropulsionStats(i).spinSpeed <> 0) Then ASSERT_FAILURE "psStats->spinSpeed != 0", "src/stats.cpp:849", "loadPropulsionStats", """%s"".""spinSpeed"" is 0", asPropulsionStats(i).id
        asPropulsionStats(i).spinAngle = ini.value("spinAngle", json_variant_new(180)).toInt()
        If Not (asPropulsionStats(i).spinAngle <> 0) Then ASSERT_FAILURE "psStats->spinAngle != 0", "src/stats.cpp:851", "loadPropulsionStats", """%s"".""spinAngle"" is 0", asPropulsionStats(i).id
        asPropulsionStats(i).acceleration = ini.value("acceleration", json_variant_new(250)).toInt()
        If Not (asPropulsionStats(i).acceleration <> 0) Then ASSERT_FAILURE "psStats->acceleration != 0", "src/stats.cpp:853", "loadPropulsionStats", """%s"".""acceleration"" is 0", asPropulsionStats(i).id
        asPropulsionStats(i).deceleration = ini.value("deceleration", json_variant_new(800)).toInt()
        asPropulsionStats(i).skidDeceleration = ini.value("skidDeceleration", json_variant_new(600)).toInt()
        asPropulsionStats(i).pIMD = 0
        asPropulsionStats(i).pIMD = statsGetIMD(ini, STAT_PROPULSION + i, "model")
        For j = 0 To MAX_PLAYERS - 1
            asPropulsionStats(i).upgrade(j) = asPropulsionStats(i).base
        Next
        If Not getPropulsionType(ini.value("type").toWzString(), asPropulsionStats(i).propulsionType) Then
            debug_ 864, LOG_FATAL, "loadPropulsionStats", "loadPropulsionStats: Invalid Propulsion type for %s", asPropulsionStats(i).name_
            retVal = False
        End If
        ini.endGroup
    Next

    loadPropulsionStats = retVal
End Function

'Load the Sensor stats from the file exported from Access
Public Function loadSensorStats(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, nullsensor As Long, i As Long, j As Long, location As String, typ As String

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:876", "loadSensorStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    statsAllocSensor n
    ' Hack to make sure ZNULLSENSOR is always first in list
    nullsensor = FindInList(list, n, "ZNULLSENSOR")
    If Not (nullsensor <> n) Then ASSERT_FAILURE "nullsensor != list.end()", "src/stats.cpp:881", "loadSensorStats", "ZNULLSENSOR is mandatory": loadSensorStats = False: Exit Function
    SwapToFront list, nullsensor
    For i = 0 To n - 1
        ini.beginGroup list(i)
        loadCompStats ini, STAT_SENSOR + i, i
        asSensorStats(i).compType = COMP_SENSOR

        asSensorStats(i).base.range = ini.value("range").toInt()
        For j = 0 To MAX_PLAYERS - 1
            asSensorStats(i).upgrade(j) = asSensorStats(i).base
        Next

        asSensorStats(i).ref = STAT_SENSOR + i

        location = ini.value("location").toWzString()
        If location = "DEFAULT" Then
            asSensorStats(i).location = LOC_DEFAULT
        ElseIf location = "TURRET" Then
            asSensorStats(i).location = LOC_TURRET
        Else
            ASSERT_FAILURE "false", "src/stats.cpp:910", "loadSensorStats", "Invalid Sensor location: %s", location
        End If
        typ = ini.value("type").toWzString()
        If typ = "STANDARD" Then
            asSensorStats(i).typ = STANDARD_SENSOR
        ElseIf typ = "INDIRECT CB" Then
            asSensorStats(i).typ = INDIRECT_CB_SENSOR
        ElseIf typ = "VTOL CB" Then
            asSensorStats(i).typ = VTOL_CB_SENSOR
        ElseIf typ = "VTOL INTERCEPT" Then
            asSensorStats(i).typ = VTOL_INTERCEPT_SENSOR
        ElseIf typ = "SUPER" Then
            asSensorStats(i).typ = SUPER_SENSOR
        ElseIf typ = "RADAR DETECTOR" Then
            asSensorStats(i).typ = RADAR_DETECTOR_SENSOR
        Else
            ASSERT_FAILURE "false", "src/stats.cpp:939", "loadSensorStats", "Invalid Sensor type: %s", typ
        End If

        'get the IMD for the component
        asSensorStats(i).pIMD = statsGetIMD(ini, STAT_SENSOR + i, "sensorModel")
        asSensorStats(i).pMountGraphic = statsGetIMD(ini, STAT_SENSOR + i, "mountModel")

        ini.endGroup
    Next
    loadSensorStats = True
End Function

'Load the ECM stats from the file exported from Access
Public Function loadECMStats(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, nullecm As Long, i As Long, j As Long, location As String

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:954", "loadECMStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    statsAllocECM n
    ' Hack to make sure ZNULLECM is always first in list
    nullecm = FindInList(list, n, "ZNULLECM")
    If Not (nullecm <> n) Then ASSERT_FAILURE "nullecm != list.end()", "src/stats.cpp:959", "loadECMStats", "ZNULLECM is mandatory": loadECMStats = False: Exit Function
    SwapToFront list, nullecm
    For i = 0 To n - 1
        ini.beginGroup list(i)
        loadCompStats ini, STAT_ECM + i, i
        asECMStats(i).compType = COMP_ECM

        asECMStats(i).base.range = ini.value("range").toInt()
        For j = 0 To MAX_PLAYERS - 1
            asECMStats(i).upgrade(j) = asECMStats(i).base
        Next

        asECMStats(i).ref = STAT_ECM + i

        location = ini.value("location").toWzString()
        If location = "DEFAULT" Then
            asECMStats(i).location = LOC_DEFAULT
        ElseIf location = "TURRET" Then
            asECMStats(i).location = LOC_TURRET
        Else
            ASSERT_FAILURE "false", "src/stats.cpp:988", "loadECMStats", "Invalid ECM location: %s", location
        End If

        'get the IMD for the component
        asECMStats(i).pIMD = statsGetIMD(ini, STAT_ECM + i, "sensorModel")
        asECMStats(i).pMountGraphic = statsGetIMD(ini, STAT_ECM + i, "mountModel")

        ini.endGroup
    Next
    loadECMStats = True
End Function

'Load the Repair stats from the file exported from Access
Public Function loadRepairStats(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, nullrepair As Long, i As Long, j As Long, location As String, retVal As Boolean

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:1003", "loadRepairStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    statsAllocRepair n
    ' Hack to make sure ZNULLREPAIR is always first in list
    nullrepair = FindInList(list, n, "ZNULLREPAIR")
    If Not (nullrepair <> n) Then ASSERT_FAILURE "nullrepair != list.end()", "src/stats.cpp:1008", "loadRepairStats", "ZNULLREPAIR is mandatory": loadRepairStats = False: Exit Function
    SwapToFront list, nullrepair
    retVal = True
    For i = 0 To n - 1
        ini.beginGroup list(i)
        loadCompStats ini, STAT_REPAIR + i, i
        asRepairStats(i).compType = COMP_REPAIRUNIT

        asRepairStats(i).base.repairPoints = ini.value("repairPoints").toInt()
        For j = 0 To MAX_PLAYERS - 1
            asRepairStats(i).upgrade(j) = asRepairStats(i).base
        Next
        asRepairStats(i).time_ = ini.value("time", json_variant_new(0)).toInt() * WEAPON_TIME

        asRepairStats(i).ref = STAT_REPAIR + i

        location = ini.value("location").toWzString()
        If location = "DEFAULT" Then
            asRepairStats(i).location = LOC_DEFAULT
        ElseIf location = "TURRET" Then
            asRepairStats(i).location = LOC_TURRET
        Else
            ASSERT_FAILURE "false", "src/stats.cpp:1039", "loadRepairStats", "Invalid Repair location: %s", location
        End If

        'check its not 0 since we will be dividing by it at a later stage
        If Not (asRepairStats(i).time_ <> 0) Then
            If Not (asRepairStats(i).time_ <> 0) Then ASSERT_FAILURE "psStats->time > 0", "src/stats.cpp:1045", "loadRepairStats", "Repair delay cannot be zero for %s", asRepairStats(i).name_
            asRepairStats(i).time_ = 1
            retVal = False
        End If

        'get the IMD for the component
        asRepairStats(i).pIMD = statsGetIMD(ini, STAT_REPAIR + i, "model")
        asRepairStats(i).pMountGraphic = statsGetIMD(ini, STAT_REPAIR + i, "mountModel")

        ini.endGroup
    Next
    loadRepairStats = retVal
End Function

'Load the Construct stats from the file exported from Access
Public Function loadConstructStats(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, nullconstruct As Long, i As Long, j As Long

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:1062", "loadConstructStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    statsAllocConstruct n
    ' Hack to make sure ZNULLCONSTRUCT is always first in list
    nullconstruct = FindInList(list, n, "ZNULLCONSTRUCT")
    If Not (nullconstruct <> n) Then ASSERT_FAILURE "nullconstruct != list.end()", "src/stats.cpp:1067", "loadConstructStats", "ZNULLCONSTRUCT is mandatory": loadConstructStats = False: Exit Function
    SwapToFront list, nullconstruct
    For i = 0 To n - 1
        ini.beginGroup list(i)
        loadCompStats ini, STAT_CONSTRUCT + i, i
        asConstructStats(i).compType = COMP_CONSTRUCT

        asConstructStats(i).base.constructPoints = ini.value("constructPoints").toInt()
        For j = 0 To MAX_PLAYERS - 1
            asConstructStats(i).upgrade(j) = asConstructStats(i).base
        Next
        asConstructStats(i).ref = STAT_CONSTRUCT + i

        'get the IMD for the component
        asConstructStats(i).pIMD = statsGetIMD(ini, STAT_CONSTRUCT + i, "sensorModel")
        asConstructStats(i).pMountGraphic = statsGetIMD(ini, STAT_CONSTRUCT + i, "mountModel")

        ini.endGroup
    Next
    loadConstructStats = True
End Function


'Load the Propulsion Types from the file exported from Access
Public Function loadPropulsionTypes(ByVal ini As WzConfig) As Boolean
    Const NumTypes As Long = PROPULSION_TYPE_NUM
    Dim list() As String, n As Long, i As Long, typ As PROPULSION_TYPE, multiplier As Long, flightName As String

    'allocate storage for the stats
    ReDim Preserve asPropulsionTypes(0 To NumTypes - 1)
    asPropulsionTypes_size = NumTypes
    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:1101", "loadPropulsionTypes", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)

    For i = 0 To NumTypes - 1
        ini.beginGroup list(i)
        multiplier = ini.value("multiplier").toUInt()

        'set the pointer for this record based on the name
        If Not getPropulsionType(list(i), typ) Then
            debug_ 1114, LOG_FATAL, "loadPropulsionTypes", "Invalid Propulsion type - %s", list(i)
            loadPropulsionTypes = False
            Exit Function
        End If

        flightName = ini.value("flightName").toWzString()
        If flightName = "GROUND" Then
            asPropulsionTypes(typ).travel = GROUND
        ElseIf flightName = "AIR" Then
            asPropulsionTypes(typ).travel = AIR
        Else
            ASSERT_FAILURE "false", "src/stats.cpp:1131", "loadPropulsionTypes", "Invalid travel type for Propulsion: %s", flightName
        End If

        'don't care about this anymore! AB FRIDAY 13/11/98
        'want it back again! AB 27/11/98
        If ULT(&HFFFF&, multiplier) Then                 ' multiplier > UWORD_MAX
            ASSERT_FAILURE "false", "src/stats.cpp:1138", "loadPropulsionTypes", "loadPropulsionTypes: power Ratio multiplier too high"
            'set to a default value since not life threatening!
            multiplier = 100
        End If
        asPropulsionTypes(typ).powerRatioMult = multiplier And &HFFFF&

        'initialise all the sound variables
        asPropulsionTypes(typ).startID = NO_SOUND
        asPropulsionTypes(typ).idleID = NO_SOUND
        asPropulsionTypes(typ).moveOffID = NO_SOUND
        asPropulsionTypes(typ).moveID = NO_SOUND
        asPropulsionTypes(typ).hissID = NO_SOUND
        asPropulsionTypes(typ).shutDownID = NO_SOUND

        ini.endGroup
    Next

    loadPropulsionTypes = True
End Function

Public Function loadTerrainTable(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, i As Long, terrainType As Long

    ReDim asTerrainTable(0 To PROPULSION_TYPE_NUM * TER_MAX - 1)    ' malloc: C leaves the entries no group sets undefined
    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:1161", "loadTerrainTable", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        terrainType = ini.value("id").toInt()
        ini.beginGroup "speedFactor"
        asTerrainTable(terrainType * PROPULSION_TYPE_NUM + PROPULSION_TYPE_WHEELED) = ini.value("wheeled", json_variant_new(100)).toUInt()
        asTerrainTable(terrainType * PROPULSION_TYPE_NUM + PROPULSION_TYPE_TRACKED) = ini.value("tracked", json_variant_new(100)).toUInt()
        asTerrainTable(terrainType * PROPULSION_TYPE_NUM + PROPULSION_TYPE_LEGGED) = ini.value("legged", json_variant_new(100)).toUInt()
        asTerrainTable(terrainType * PROPULSION_TYPE_NUM + PROPULSION_TYPE_HOVER) = ini.value("hover", json_variant_new(100)).toUInt()
        asTerrainTable(terrainType * PROPULSION_TYPE_NUM + PROPULSION_TYPE_LIFT) = ini.value("lift", json_variant_new(100)).toUInt()
        asTerrainTable(terrainType * PROPULSION_TYPE_NUM + PROPULSION_TYPE_PROPELLOR) = ini.value("propellor", json_variant_new(100)).toUInt()
        asTerrainTable(terrainType * PROPULSION_TYPE_NUM + PROPULSION_TYPE_HALF_TRACKED) = ini.value("half-tracked", json_variant_new(100)).toUInt()
        ini.endGroup
        ini.endGroup
    Next
    loadTerrainTable = True
End Function

Private Function statsGetAudioIDFromString(ByVal szStatName As String, ByVal szWavName As String, piWavID As Long) As Boolean
    If szWavName = "-1" Then
        piWavID = NO_SOUND
    Else
        piWavID = audio_GetIDFromStr(szWavName)
        If piWavID = NO_SOUND Then
            debug_ 1189, LOG_FATAL, "statsGetAudioIDFromString", "Could not get ID %d for sound %s", piWavID, szWavName
            statsGetAudioIDFromString = False
            Exit Function
        End If
    End If
    If (piWavID < 0 Or piWavID > ID_MAX_SOUND) And piWavID <> NO_SOUND Then
        debug_ 1194, LOG_FATAL, "statsGetAudioIDFromString", "Invalid ID - %d for sound %s", piWavID, szStatName
        statsGetAudioIDFromString = False
        Exit Function
    End If
    statsGetAudioIDFromString = True
End Function

Public Function loadWeaponModifiers(ByVal ini As WzConfig) As Boolean
    Dim i As Long, j As Long, list() As String, n As Long, keys() As String, nKeys As Long
    Dim effectInc As WEAPON_EFFECT, propInc As PROPULSION_TYPE, modifier As Long, body As BODY_SIZE

    'initialise to 100%
    For i = 0 To WE_NUMEFFECTS - 1
        For j = 0 To PROPULSION_TYPE_NUM - 1
            asWeaponModifier(i, j) = 100
        Next
        For j = 0 To SIZE_NUM - 1
            asWeaponModifierBody(i, j) = 100
        Next
    Next
    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/stats.cpp:1214", "loadWeaponModifiers", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        'get the weapon effect inc
        If Not getWeaponEffect(list(i), effectInc) Then
            debug_ 1225, LOG_FATAL, "loadWeaponModifiers", "Invalid Weapon Effect - %s", list(i)
            GoTo continue_                      ' (the group stays open, as in C++)
        End If
        nKeys = ini.childKeys(keys)
        For j = 0 To nKeys - 1
            modifier = ini.value(keys(j)).toInt()
            If Not getPropulsionType(keys(j), propInc) Then
                ' If not propulsion, must be body
                body = SIZE_NUM
                If Not getBodySize(keys(j), body) Then
                    debug_ 1238, LOG_FATAL, "loadWeaponModifiers", "Invalid Propulsion or Body type - %s", keys(j)
                    GoTo continue_inner
                End If
                asWeaponModifierBody(effectInc, body) = modifier And &HFFFF&
            Else ' is propulsion
                asWeaponModifier(effectInc, propInc) = modifier And &HFFFF&
            End If
continue_inner:
        Next
        ini.endGroup
continue_:
    Next
    loadWeaponModifiers = True
End Function

'Load the propulsion type sounds from the file exported from Access
Public Function loadPropulsionSounds(ByVal pFileName As String) As Boolean
    Dim i As Long, startID As Long, idleID As Long, moveOffID As Long, moveID As Long, hissID As Long, shutDownID As Long
    Dim typ As PROPULSION_TYPE, ini As WzConfig, list() As String, n As Long

    If Not (asPropulsionTypes_size <> 0) Then ASSERT_FAILURE "asPropulsionTypes.size() != 0", "src/stats.cpp:1259", "loadPropulsionSounds", "loadPropulsionSounds: Propulsion type stats not loaded"

    Set ini = WzConfig_new(pFileName, WzConfig_ReadOnlyAndRequired)
    n = ini.childGroups(list)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        If Not statsGetAudioIDFromString(list(i), ini.value("szStart").toWzString(), startID) Then
            loadPropulsionSounds = False
            Exit Function
        End If
        If Not statsGetAudioIDFromString(list(i), ini.value("szIdle").toWzString(), idleID) Then
            loadPropulsionSounds = False
            Exit Function
        End If
        If Not statsGetAudioIDFromString(list(i), ini.value("szMoveOff").toWzString(), moveOffID) Then
            loadPropulsionSounds = False
            Exit Function
        End If
        If Not statsGetAudioIDFromString(list(i), ini.value("szMove").toWzString(), moveID) Then
            loadPropulsionSounds = False
            Exit Function
        End If
        If Not statsGetAudioIDFromString(list(i), ini.value("szHiss").toWzString(), hissID) Then
            loadPropulsionSounds = False
            Exit Function
        End If
        If Not statsGetAudioIDFromString(list(i), ini.value("szShutDown").toWzString(), shutDownID) Then
            loadPropulsionSounds = False
            Exit Function
        End If
        If Not getPropulsionType(list(i), typ) Then
            debug_ 1292, LOG_FATAL, "loadPropulsionSounds", "Invalid Propulsion type - %s", list(i)
            loadPropulsionSounds = False
            Exit Function
        End If
        asPropulsionTypes(typ).startID = ToInt16(startID)
        asPropulsionTypes(typ).idleID = ToInt16(idleID)
        asPropulsionTypes(typ).moveOffID = ToInt16(moveOffID)
        asPropulsionTypes(typ).moveID = ToInt16(moveID)
        asPropulsionTypes(typ).hissID = ToInt16(hissID)
        asPropulsionTypes(typ).shutDownID = ToInt16(shutDownID)

        ini.endGroup
    Next

    loadPropulsionSounds = True
End Function

'get the speed factor for a given terrain type and propulsion type
Public Function getSpeedFactor(ByVal typ As Long, ByVal propulsionType As Long) As Long
    If Not (ULT(propulsionType, PROPULSION_TYPE_NUM)) Then ASSERT_FAILURE "propulsionType < PROPULSION_TYPE_NUM", "src/stats.cpp:1312", "getSpeedFactor", "The propulsion type is too large"
    getSpeedFactor = asTerrainTable(typ * PROPULSION_TYPE_NUM + propulsionType)
End Function

Public Function getCompFromName(ByVal compType As COMPONENT_TYPE, ByVal name_ As String) As Long
    getCompFromName = getCompFromID(compType, name_)
End Function

Public Function getCompFromID(ByVal compType As COMPONENT_TYPE, ByVal name_ As String) As Long
    Dim psComp As Long
    InitLookups
    psComp = 0
    If lookupCompStatPtr.Exists(name_) Then
        psComp = lookupCompStatPtr(name_)
    End If
    If Not (psComp <> 0) Then ASSERT_FAILURE "psComp", "src/stats.cpp:1329", "getCompFromID", "No such component ID [%s] found", name_: getCompFromID = -1: Exit Function
    If Not (compType = COMPONENT_STATS_compType(psComp)) Then ASSERT_FAILURE "compType == psComp->compType", "src/stats.cpp:1330", "getCompFromID", "Wrong component type for ID %s", name_: getCompFromID = -1: Exit Function
    ' (psComp->index <= INT_MAX always holds)
    getCompFromID = BASE_STATS_index(psComp)
End Function

' Get the component for a stat based on the name alone.
' Returns NULL if record not found
Public Function getCompStatsFromName(ByVal name_ As String) As Long
    Dim psComp As Long
    InitLookups
    psComp = 0
    If lookupCompStatPtr.Exists(name_) Then
        psComp = lookupCompStatPtr(name_)
    End If
    getCompStatsFromName = psComp
End Function

Public Function getBaseStatsFromName(ByVal name_ As String) As Long
    Dim psStat As Long
    InitLookups
    psStat = 0
    If lookupStatPtr.Exists(name_) Then
        psStat = lookupStatPtr(name_)
    End If
    getBaseStatsFromName = psStat
End Function

'sets the store to the body size based on the name passed in - returns false
'if doesn't compare with any
Public Function getBodySize(ByVal size As String, pStore As BODY_SIZE) As Boolean
    If size = "LIGHT" Then
        pStore = SIZE_LIGHT
        getBodySize = True
        Exit Function
    ElseIf size = "MEDIUM" Then
        pStore = SIZE_MEDIUM
        getBodySize = True
        Exit Function
    ElseIf size = "HEAVY" Then
        pStore = SIZE_HEAVY
        getBodySize = True
        Exit Function
    ElseIf size = "SUPER HEAVY" Then
        pStore = SIZE_SUPER_HEAVY
        getBodySize = True
        Exit Function
    End If

    ASSERT_FAILURE "false", "src/stats.cpp:1392", "getBodySize", "Invalid size - %s", size
    getBodySize = False
End Function

'returns the weapon sub class based on the string name passed in
Public Function getWeaponSubClass(ByVal subClass As String, wclass As WEAPON_SUBCLASS) As Boolean
    If subClass = "CANNON" Then
        wclass = WSC_CANNON
    ElseIf subClass = "MORTARS" Then
        wclass = WSC_MORTARS
    ElseIf subClass = "MISSILE" Then
        wclass = WSC_MISSILE
    ElseIf subClass = "ROCKET" Then
        wclass = WSC_ROCKET
    ElseIf subClass = "ENERGY" Then
        wclass = WSC_ENERGY
    ElseIf subClass = "GAUSS" Then
        wclass = WSC_GAUSS
    ElseIf subClass = "FLAME" Then
        wclass = WSC_FLAME
    ElseIf subClass = "HOWITZERS" Then
        wclass = WSC_HOWITZERS
    ElseIf subClass = "MACHINE GUN" Then
        wclass = WSC_MGUN
    ElseIf subClass = "ELECTRONIC" Then
        wclass = WSC_ELECTRONIC
    ElseIf subClass = "A-A GUN" Then
        wclass = WSC_AAGUN
    ElseIf subClass = "SLOW MISSILE" Then
        wclass = WSC_SLOWMISSILE
    ElseIf subClass = "SLOW ROCKET" Then
        wclass = WSC_SLOWROCKET
    ElseIf subClass = "LAS_SAT" Then
        wclass = WSC_LAS_SAT
    ElseIf subClass = "BOMB" Then
        wclass = WSC_BOMB
    ElseIf subClass = "COMMAND" Then
        wclass = WSC_COMMAND
    ElseIf subClass = "EMP" Then
        wclass = WSC_EMP
    Else
        ASSERT_FAILURE "!""Invalid weapon sub class""", "src/stats.cpp:1469", "getWeaponSubClass", "Invalid weapon sub class: %s", subClass
        getWeaponSubClass = False
        Exit Function
    End If

    getWeaponSubClass = True
End Function

'returns the weapon sub class based on the string name passed in
' (C++: the overload getWeaponSubClass(WEAPON_SUBCLASS))
Public Function getWeaponSubClass_name(ByVal wclass As WEAPON_SUBCLASS) As String
    Select Case wclass
    Case WSC_CANNON: getWeaponSubClass_name = "CANNON": Exit Function
    Case WSC_MORTARS: getWeaponSubClass_name = "MORTARS": Exit Function
    Case WSC_MISSILE: getWeaponSubClass_name = "MISSILE": Exit Function
    Case WSC_ROCKET: getWeaponSubClass_name = "ROCKET": Exit Function
    Case WSC_ENERGY: getWeaponSubClass_name = "ENERGY": Exit Function
    Case WSC_GAUSS: getWeaponSubClass_name = "GAUSS": Exit Function
    Case WSC_FLAME: getWeaponSubClass_name = "FLAME": Exit Function
    Case WSC_HOWITZERS: getWeaponSubClass_name = "HOWITZERS": Exit Function
    Case WSC_MGUN: getWeaponSubClass_name = "MACHINE GUN": Exit Function
    Case WSC_ELECTRONIC: getWeaponSubClass_name = "ELECTRONIC": Exit Function
    Case WSC_AAGUN: getWeaponSubClass_name = "A-A GUN": Exit Function
    Case WSC_SLOWMISSILE: getWeaponSubClass_name = "SLOW MISSILE": Exit Function
    Case WSC_SLOWROCKET: getWeaponSubClass_name = "SLOW ROCKET": Exit Function
    Case WSC_LAS_SAT: getWeaponSubClass_name = "LAS_SAT": Exit Function
    Case WSC_BOMB: getWeaponSubClass_name = "BOMB": Exit Function
    Case WSC_COMMAND: getWeaponSubClass_name = "COMMAND": Exit Function
    Case WSC_EMP: getWeaponSubClass_name = "EMP": Exit Function
    Case WSC_NUM_WEAPON_SUBCLASSES
    End Select
    ASSERT_FAILURE "false", "src/stats.cpp:1500", "getWeaponSubClass", "No such weapon subclass"
    getWeaponSubClass_name = "Bad weapon subclass"
End Function

' returns the translated weapon sub class name - for local display purposes only
Public Function getWeaponSubClassDisplayName(ByVal wclass As WEAPON_SUBCLASS, ByVal shortForm As Boolean) As String
    Select Case wclass
    Case WSC_CANNON
        getWeaponSubClassDisplayName = "Cannons": Exit Function
    Case WSC_MORTARS
        getWeaponSubClassDisplayName = "Mortars": Exit Function
    Case WSC_MISSILE
        getWeaponSubClassDisplayName = "Missiles": Exit Function
    Case WSC_ROCKET
        getWeaponSubClassDisplayName = "Rockets": Exit Function
    Case WSC_ENERGY
        getWeaponSubClassDisplayName = "Energy": Exit Function
    Case WSC_GAUSS
        getWeaponSubClassDisplayName = "Gauss": Exit Function
    Case WSC_FLAME
        ' TRANSLATORS: "Flame" weapons class
        getWeaponSubClassDisplayName = "Flame": Exit Function
    Case WSC_HOWITZERS
        getWeaponSubClassDisplayName = "Howitzers": Exit Function
    Case WSC_MGUN
        If shortForm Then
            ' TRANSLATORS: A short-form for "Machine Guns" (should be shorter than the translated full form - if not, just use the full translation for "Machine Guns")
            getWeaponSubClassDisplayName = "MG": Exit Function
        Else
            getWeaponSubClassDisplayName = "Machine Guns": Exit Function
        End If
    Case WSC_ELECTRONIC
        ' TRANSLATORS: "Electronic" weapon class
        getWeaponSubClassDisplayName = "Electronic": Exit Function
    Case WSC_AAGUN
        If shortForm Then
            ' TRANSLATORS: A short-form for "Anti-Air" weapons class (should be shorter than the translated full form - if not, just use the full translation for "Anti-Air")
            getWeaponSubClassDisplayName = "A-A": Exit Function
        Else
            getWeaponSubClassDisplayName = "Anti-Air": Exit Function
        End If
    Case WSC_SLOWMISSILE
        getWeaponSubClassDisplayName = "Slow Missile": Exit Function
    Case WSC_SLOWROCKET
        getWeaponSubClassDisplayName = "Slow Rocket": Exit Function
    Case WSC_LAS_SAT
        If shortForm Then
            ' TRANSLATORS: A short-form for "Laser Satellite" weapons class - some just use "Las-Sat" or "LasSat"
            getWeaponSubClassDisplayName = "Las-Sat": Exit Function
        Else
            getWeaponSubClassDisplayName = "Laser Satellite": Exit Function
        End If
    Case WSC_BOMB
        getWeaponSubClassDisplayName = "Bombs": Exit Function
    Case WSC_COMMAND
        getWeaponSubClassDisplayName = "Command": Exit Function
    Case WSC_EMP
        ' TRANSLATORS: "EMP" (electromagnetic pulse) weapons class - probably should keep it as "EMP" (or the respective acronym in your language)
        getWeaponSubClassDisplayName = "EMP": Exit Function
    Case WSC_NUM_WEAPON_SUBCLASSES
    End Select
    ASSERT_FAILURE "false", "src/stats.cpp:1572", "getWeaponSubClassDisplayName", "No such weapon subclass"
    getWeaponSubClassDisplayName = ""
End Function

'returns the movement model based on the string name passed in
Private Function getMovementModel(ByVal movementModel As String, model As MOVEMENT_MODEL) As Boolean
    If movementModel = "DIRECT" Then
        model = MM_DIRECT
    ElseIf movementModel = "INDIRECT" Then
        model = MM_INDIRECT
    ElseIf movementModel = "HOMING-DIRECT" Then
        model = MM_HOMINGDIRECT
    ElseIf movementModel = "HOMING-INDIRECT" Then
        model = MM_HOMINGINDIRECT
    Else
        ' We've got problem if we got here
        ASSERT_FAILURE "!""Invalid movement model""", "src/stats.cpp:1598", "getMovementModel", "Invalid movement model: %s", movementModel
        getMovementModel = False
        Exit Function
    End If

    getMovementModel = True
End Function

Public Function getWeaponEffect(ByVal weaponEffect As String, effect_ As WEAPON_EFFECT) As Boolean
    If weaponEffect = "ANTI PERSONNEL" Then
        effect_ = WE_ANTI_PERSONNEL
    ElseIf weaponEffect = "ANTI TANK" Then
        effect_ = WE_ANTI_TANK
    ElseIf weaponEffect = "BUNKER BUSTER" Then
        effect_ = WE_BUNKER_BUSTER
    ElseIf weaponEffect = "ARTILLERY ROUND" Then
        effect_ = WE_ARTILLERY_ROUND
    ElseIf weaponEffect = "FLAMER" Then
        effect_ = WE_FLAMER
    ElseIf weaponEffect = "ANTI AIRCRAFT" Or weaponEffect = "ALL ROUNDER" Then
        effect_ = WE_ANTI_AIRCRAFT
    Else
        ASSERT_FAILURE "!""Invalid weapon effect""", "src/stats.cpp:1645", "getWeaponEffect", "Invalid weapon effect: %s", weaponEffect
        getWeaponEffect = False
        Exit Function
    End If

    getWeaponEffect = True
End Function

'returns the weapon effect string based on the enum passed in
' (C++: the overload getWeaponEffect(WEAPON_EFFECT))
Public Function getWeaponEffect_name(ByVal effect_ As WEAPON_EFFECT) As String
    Select Case effect_
    Case WE_ANTI_PERSONNEL: getWeaponEffect_name = "ANTI PERSONNEL": Exit Function
    Case WE_ANTI_TANK: getWeaponEffect_name = "ANTI TANK": Exit Function
    Case WE_BUNKER_BUSTER: getWeaponEffect_name = "BUNKER BUSTER": Exit Function
    Case WE_ARTILLERY_ROUND: getWeaponEffect_name = "ARTILLERY ROUND": Exit Function
    Case WE_FLAMER: getWeaponEffect_name = "FLAMER": Exit Function
    Case WE_ANTI_AIRCRAFT: getWeaponEffect_name = "ANTI AIRCRAFT": Exit Function
    Case WE_NUMEFFECTS
    End Select
    ASSERT_FAILURE "false", "src/stats.cpp:1665", "getWeaponEffect", "No such weapon effect"
    getWeaponEffect_name = "Bad weapon effect"
End Function

Public Function getWeaponClass(ByVal weaponClassStr As String, weaponClass As WEAPON_CLASS) As Boolean
    If weaponClassStr = "KINETIC" Then
        weaponClass = WC_KINETIC
    ElseIf weaponClassStr = "HEAT" Then
        weaponClass = WC_HEAT
    Else
        ASSERT_FAILURE "false", "src/stats.cpp:1681", "getWeaponClass", "Bad weapon class %s", weaponClassStr
        getWeaponClass = False
        Exit Function
    End If
    getWeaponClass = True
End Function

' (ASSERT_PLAYER_OR_RETURN(retVal, player): ASSERT_OR_RETURN(retVal, player >= 0 && player < MAX_PLAYERS, "Invalid player: %" PRIu32 "", player))

'Access functions for the upgradeable stats of a weapon
Public Function weaponFirePause(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1693", "weaponFirePause", "Invalid player: %u", player: weaponFirePause = 0: Exit Function
    weaponFirePause = psStats.upgrade(player).firePause
End Function

' Reload time is reduced for weapons with salvo fire
Public Function weaponReloadTime(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1700", "weaponReloadTime", "Invalid player: %u", player: weaponReloadTime = 0: Exit Function
    weaponReloadTime = psStats.upgrade(player).reloadTime
End Function

Public Function weaponLongHit(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1706", "weaponLongHit", "Invalid player: %u", player: weaponLongHit = 0: Exit Function
    weaponLongHit = psStats.upgrade(player).hitChance
End Function

Public Function weaponShortHit(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1712", "weaponShortHit", "Invalid player: %u", player: weaponShortHit = 0: Exit Function
    weaponShortHit = psStats.upgrade(player).shortHitChance
End Function

Public Function weaponDamage(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1718", "weaponDamage", "Invalid player: %u", player: weaponDamage = 0: Exit Function
    weaponDamage = psStats.upgrade(player).damage
End Function

Public Function weaponRadDamage(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1724", "weaponRadDamage", "Invalid player: %u", player: weaponRadDamage = 0: Exit Function
    weaponRadDamage = psStats.upgrade(player).radiusDamage
End Function

Public Function weaponPeriodicalDamage(psStats As WEAPON_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1730", "weaponPeriodicalDamage", "Invalid player: %u", player: weaponPeriodicalDamage = 0: Exit Function
    weaponPeriodicalDamage = psStats.upgrade(player).periodicalDamage
End Function

Public Function sensorRange(psStats As SENSOR_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1736", "sensorRange", "Invalid player: %u", player: sensorRange = 0: Exit Function
    sensorRange = psStats.upgrade(player).range
End Function

Public Function ecmRange(psStats As ECM_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1742", "ecmRange", "Invalid player: %u", player: ecmRange = 0: Exit Function
    ecmRange = psStats.upgrade(player).range
End Function

Public Function repairPoints(psStats As REPAIR_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1748", "repairPoints", "Invalid player: %u", player: repairPoints = 0: Exit Function
    repairPoints = psStats.upgrade(player).repairPoints
End Function

Public Function constructorPoints(psStats As CONSTRUCT_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1754", "constructorPoints", "Invalid player: %u", player: constructorPoints = 0: Exit Function
    constructorPoints = psStats.upgrade(player).constructPoints
End Function

Public Function bodyPower(psStats As BODY_STATS, ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1760", "bodyPower", "Invalid player: %u", player: bodyPower = 0: Exit Function
    bodyPower = psStats.upgrade(player).power
End Function

Public Function bodyArmour(psStats As BODY_STATS, ByVal player As Long, ByVal weaponClass As WEAPON_CLASS) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1766", "bodyArmour", "Invalid player: %u", player: bodyArmour = 0: Exit Function
    Select Case weaponClass
    Case WC_KINETIC
        bodyArmour = psStats.upgrade(player).armour
        Exit Function
    Case WC_HEAT
        bodyArmour = psStats.upgrade(player).thermal
        Exit Function
    Case WC_NUM_WEAPON_CLASSES
    End Select
    ASSERT_FAILURE "false", "src/stats.cpp:1776", "bodyArmour", "Unknown weapon class"
    bodyArmour = 0 ' Should never get here.
End Function

'calculates the weapons ROF based on the fire pause and the salvos
Public Function weaponROF(psStat As WEAPON_STATS, ByVal player As Long) As Long
    Dim rof As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/stats.cpp:1783", "weaponROF", "Invalid player: %u", player: weaponROF = 0: Exit Function
    rof = 0
    ' if there are salvos
    If player >= 0 Then
        If psStat.upgrade(player).numRounds <> 0 And psStat.upgrade(player).reloadTime <> 0 Then
            ' Rounds per salvo multiplied with the number of salvos per minute
            rof = psStat.upgrade(player).numRounds * 60 * GAME_TICKS_PER_SEC \ weaponReloadTime(psStat, player)
        End If
    End If

    If rof = 0 Then
        rof = weaponFirePause(psStat, player)
        If rof <> 0 Then
            rof = (60 * GAME_TICKS_PER_SEC \ rof) And &HFFFF&
        End If
        'else leave it at 0
    End If
    weaponROF = rof
End Function

' (VB) a bad BASE_STATS * or COMPONENT_STATS * handle: a porting error
Private Sub BadStatsHandle(ByVal psStats As Long)
    Err.Raise vbObjectError + 513, "stats_c", "(VB) not a stats handle: " & Hex$(psStats)
End Sub

' (VB) BASE_STATS::hasType: the type of the stat's ref
Public Function BASE_STATS_hasType(ByVal psStats As Long, ByVal typ As StatType) As Boolean
    BASE_STATS_hasType = (BASE_STATS_ref(psStats) And STAT_MASK) = typ
End Function

' (VB) the fields of a stat through its BASE_STATS * or COMPONENT_STATS * handle

Public Property Get BASE_STATS_id(ByVal psStats As Long) As String
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: BASE_STATS_id = asBodyStats(i).id
    Case STAT_BRAIN: BASE_STATS_id = asBrainStats(i).id
    Case STAT_PROPULSION: BASE_STATS_id = asPropulsionStats(i).id
    Case STAT_SENSOR: BASE_STATS_id = asSensorStats(i).id
    Case STAT_ECM: BASE_STATS_id = asECMStats(i).id
    Case STAT_REPAIR: BASE_STATS_id = asRepairStats(i).id
    Case STAT_WEAPON: BASE_STATS_id = asWeaponStats(i).id
    Case STAT_CONSTRUCT: BASE_STATS_id = asConstructStats(i).id
    Case STAT_RESEARCH: BASE_STATS_id = asResearch(i).id
    Case STAT_TEMPLATE: BASE_STATS_id = apsTemplates(i).id
    Case STAT_STRUCTURE: BASE_STATS_id = asStructureStats(i).id
    Case STAT_FEATURE: BASE_STATS_id = asFeatureStats(i).id
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let BASE_STATS_id(ByVal psStats As Long, ByVal value As String)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).id = value
    Case STAT_BRAIN: asBrainStats(i).id = value
    Case STAT_PROPULSION: asPropulsionStats(i).id = value
    Case STAT_SENSOR: asSensorStats(i).id = value
    Case STAT_ECM: asECMStats(i).id = value
    Case STAT_REPAIR: asRepairStats(i).id = value
    Case STAT_WEAPON: asWeaponStats(i).id = value
    Case STAT_CONSTRUCT: asConstructStats(i).id = value
    Case STAT_RESEARCH: asResearch(i).id = value
    Case STAT_TEMPLATE: apsTemplates(i).id = value
    Case STAT_STRUCTURE: asStructureStats(i).id = value
    Case STAT_FEATURE: asFeatureStats(i).id = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get BASE_STATS_name(ByVal psStats As Long) As String
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: BASE_STATS_name = asBodyStats(i).name_
    Case STAT_BRAIN: BASE_STATS_name = asBrainStats(i).name_
    Case STAT_PROPULSION: BASE_STATS_name = asPropulsionStats(i).name_
    Case STAT_SENSOR: BASE_STATS_name = asSensorStats(i).name_
    Case STAT_ECM: BASE_STATS_name = asECMStats(i).name_
    Case STAT_REPAIR: BASE_STATS_name = asRepairStats(i).name_
    Case STAT_WEAPON: BASE_STATS_name = asWeaponStats(i).name_
    Case STAT_CONSTRUCT: BASE_STATS_name = asConstructStats(i).name_
    Case STAT_RESEARCH: BASE_STATS_name = asResearch(i).name_
    Case STAT_TEMPLATE: BASE_STATS_name = apsTemplates(i).name_
    Case STAT_STRUCTURE: BASE_STATS_name = asStructureStats(i).name_
    Case STAT_FEATURE: BASE_STATS_name = asFeatureStats(i).name_
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let BASE_STATS_name(ByVal psStats As Long, ByVal value As String)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).name_ = value
    Case STAT_BRAIN: asBrainStats(i).name_ = value
    Case STAT_PROPULSION: asPropulsionStats(i).name_ = value
    Case STAT_SENSOR: asSensorStats(i).name_ = value
    Case STAT_ECM: asECMStats(i).name_ = value
    Case STAT_REPAIR: asRepairStats(i).name_ = value
    Case STAT_WEAPON: asWeaponStats(i).name_ = value
    Case STAT_CONSTRUCT: asConstructStats(i).name_ = value
    Case STAT_RESEARCH: asResearch(i).name_ = value
    Case STAT_TEMPLATE: apsTemplates(i).name_ = value
    Case STAT_STRUCTURE: asStructureStats(i).name_ = value
    Case STAT_FEATURE: asFeatureStats(i).name_ = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get BASE_STATS_ref(ByVal psStats As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: BASE_STATS_ref = asBodyStats(i).ref
    Case STAT_BRAIN: BASE_STATS_ref = asBrainStats(i).ref
    Case STAT_PROPULSION: BASE_STATS_ref = asPropulsionStats(i).ref
    Case STAT_SENSOR: BASE_STATS_ref = asSensorStats(i).ref
    Case STAT_ECM: BASE_STATS_ref = asECMStats(i).ref
    Case STAT_REPAIR: BASE_STATS_ref = asRepairStats(i).ref
    Case STAT_WEAPON: BASE_STATS_ref = asWeaponStats(i).ref
    Case STAT_CONSTRUCT: BASE_STATS_ref = asConstructStats(i).ref
    Case STAT_RESEARCH: BASE_STATS_ref = asResearch(i).ref
    Case STAT_TEMPLATE: BASE_STATS_ref = apsTemplates(i).ref
    Case STAT_STRUCTURE: BASE_STATS_ref = asStructureStats(i).ref
    Case STAT_FEATURE: BASE_STATS_ref = asFeatureStats(i).ref
    Case DUMMY_STATS_REF0: BASE_STATS_ref = 0
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let BASE_STATS_ref(ByVal psStats As Long, ByVal value As Long)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).ref = value
    Case STAT_BRAIN: asBrainStats(i).ref = value
    Case STAT_PROPULSION: asPropulsionStats(i).ref = value
    Case STAT_SENSOR: asSensorStats(i).ref = value
    Case STAT_ECM: asECMStats(i).ref = value
    Case STAT_REPAIR: asRepairStats(i).ref = value
    Case STAT_WEAPON: asWeaponStats(i).ref = value
    Case STAT_CONSTRUCT: asConstructStats(i).ref = value
    Case STAT_RESEARCH: asResearch(i).ref = value
    Case STAT_TEMPLATE: apsTemplates(i).ref = value
    Case STAT_STRUCTURE: asStructureStats(i).ref = value
    Case STAT_FEATURE: asFeatureStats(i).ref = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get BASE_STATS_index(ByVal psStats As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: BASE_STATS_index = asBodyStats(i).index
    Case STAT_BRAIN: BASE_STATS_index = asBrainStats(i).index
    Case STAT_PROPULSION: BASE_STATS_index = asPropulsionStats(i).index
    Case STAT_SENSOR: BASE_STATS_index = asSensorStats(i).index
    Case STAT_ECM: BASE_STATS_index = asECMStats(i).index
    Case STAT_REPAIR: BASE_STATS_index = asRepairStats(i).index
    Case STAT_WEAPON: BASE_STATS_index = asWeaponStats(i).index
    Case STAT_CONSTRUCT: BASE_STATS_index = asConstructStats(i).index
    Case STAT_RESEARCH: BASE_STATS_index = asResearch(i).base_index
    Case STAT_TEMPLATE: BASE_STATS_index = apsTemplates(i).index
    Case STAT_STRUCTURE: BASE_STATS_index = asStructureStats(i).index
    Case STAT_FEATURE: BASE_STATS_index = asFeatureStats(i).index
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let BASE_STATS_index(ByVal psStats As Long, ByVal value As Long)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).index = value
    Case STAT_BRAIN: asBrainStats(i).index = value
    Case STAT_PROPULSION: asPropulsionStats(i).index = value
    Case STAT_SENSOR: asSensorStats(i).index = value
    Case STAT_ECM: asECMStats(i).index = value
    Case STAT_REPAIR: asRepairStats(i).index = value
    Case STAT_WEAPON: asWeaponStats(i).index = value
    Case STAT_CONSTRUCT: asConstructStats(i).index = value
    Case STAT_RESEARCH: asResearch(i).base_index = value
    Case STAT_TEMPLATE: apsTemplates(i).index = value
    Case STAT_STRUCTURE: asStructureStats(i).index = value
    Case STAT_FEATURE: asFeatureStats(i).index = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_pIMD(ByVal psStats As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_pIMD = asBodyStats(i).pIMD
    Case STAT_BRAIN: COMPONENT_STATS_pIMD = asBrainStats(i).pIMD
    Case STAT_PROPULSION: COMPONENT_STATS_pIMD = asPropulsionStats(i).pIMD
    Case STAT_SENSOR: COMPONENT_STATS_pIMD = asSensorStats(i).pIMD
    Case STAT_ECM: COMPONENT_STATS_pIMD = asECMStats(i).pIMD
    Case STAT_REPAIR: COMPONENT_STATS_pIMD = asRepairStats(i).pIMD
    Case STAT_WEAPON: COMPONENT_STATS_pIMD = asWeaponStats(i).pIMD
    Case STAT_CONSTRUCT: COMPONENT_STATS_pIMD = asConstructStats(i).pIMD
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_pIMD(ByVal psStats As Long, ByVal value As Long)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).pIMD = value
    Case STAT_BRAIN: asBrainStats(i).pIMD = value
    Case STAT_PROPULSION: asPropulsionStats(i).pIMD = value
    Case STAT_SENSOR: asSensorStats(i).pIMD = value
    Case STAT_ECM: asECMStats(i).pIMD = value
    Case STAT_REPAIR: asRepairStats(i).pIMD = value
    Case STAT_WEAPON: asWeaponStats(i).pIMD = value
    Case STAT_CONSTRUCT: asConstructStats(i).pIMD = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_buildPower(ByVal psStats As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_buildPower = asBodyStats(i).buildPower
    Case STAT_BRAIN: COMPONENT_STATS_buildPower = asBrainStats(i).buildPower
    Case STAT_PROPULSION: COMPONENT_STATS_buildPower = asPropulsionStats(i).buildPower
    Case STAT_SENSOR: COMPONENT_STATS_buildPower = asSensorStats(i).buildPower
    Case STAT_ECM: COMPONENT_STATS_buildPower = asECMStats(i).buildPower
    Case STAT_REPAIR: COMPONENT_STATS_buildPower = asRepairStats(i).buildPower
    Case STAT_WEAPON: COMPONENT_STATS_buildPower = asWeaponStats(i).buildPower
    Case STAT_CONSTRUCT: COMPONENT_STATS_buildPower = asConstructStats(i).buildPower
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_buildPower(ByVal psStats As Long, ByVal value As Long)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).buildPower = value
    Case STAT_BRAIN: asBrainStats(i).buildPower = value
    Case STAT_PROPULSION: asPropulsionStats(i).buildPower = value
    Case STAT_SENSOR: asSensorStats(i).buildPower = value
    Case STAT_ECM: asECMStats(i).buildPower = value
    Case STAT_REPAIR: asRepairStats(i).buildPower = value
    Case STAT_WEAPON: asWeaponStats(i).buildPower = value
    Case STAT_CONSTRUCT: asConstructStats(i).buildPower = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_buildPoints(ByVal psStats As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_buildPoints = asBodyStats(i).buildPoints
    Case STAT_BRAIN: COMPONENT_STATS_buildPoints = asBrainStats(i).buildPoints
    Case STAT_PROPULSION: COMPONENT_STATS_buildPoints = asPropulsionStats(i).buildPoints
    Case STAT_SENSOR: COMPONENT_STATS_buildPoints = asSensorStats(i).buildPoints
    Case STAT_ECM: COMPONENT_STATS_buildPoints = asECMStats(i).buildPoints
    Case STAT_REPAIR: COMPONENT_STATS_buildPoints = asRepairStats(i).buildPoints
    Case STAT_WEAPON: COMPONENT_STATS_buildPoints = asWeaponStats(i).buildPoints
    Case STAT_CONSTRUCT: COMPONENT_STATS_buildPoints = asConstructStats(i).buildPoints
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_buildPoints(ByVal psStats As Long, ByVal value As Long)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).buildPoints = value
    Case STAT_BRAIN: asBrainStats(i).buildPoints = value
    Case STAT_PROPULSION: asPropulsionStats(i).buildPoints = value
    Case STAT_SENSOR: asSensorStats(i).buildPoints = value
    Case STAT_ECM: asECMStats(i).buildPoints = value
    Case STAT_REPAIR: asRepairStats(i).buildPoints = value
    Case STAT_WEAPON: asWeaponStats(i).buildPoints = value
    Case STAT_CONSTRUCT: asConstructStats(i).buildPoints = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_weight(ByVal psStats As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_weight = asBodyStats(i).weight
    Case STAT_BRAIN: COMPONENT_STATS_weight = asBrainStats(i).weight
    Case STAT_PROPULSION: COMPONENT_STATS_weight = asPropulsionStats(i).weight
    Case STAT_SENSOR: COMPONENT_STATS_weight = asSensorStats(i).weight
    Case STAT_ECM: COMPONENT_STATS_weight = asECMStats(i).weight
    Case STAT_REPAIR: COMPONENT_STATS_weight = asRepairStats(i).weight
    Case STAT_WEAPON: COMPONENT_STATS_weight = asWeaponStats(i).weight
    Case STAT_CONSTRUCT: COMPONENT_STATS_weight = asConstructStats(i).weight
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_weight(ByVal psStats As Long, ByVal value As Long)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).weight = value
    Case STAT_BRAIN: asBrainStats(i).weight = value
    Case STAT_PROPULSION: asPropulsionStats(i).weight = value
    Case STAT_SENSOR: asSensorStats(i).weight = value
    Case STAT_ECM: asECMStats(i).weight = value
    Case STAT_REPAIR: asRepairStats(i).weight = value
    Case STAT_WEAPON: asWeaponStats(i).weight = value
    Case STAT_CONSTRUCT: asConstructStats(i).weight = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_compType(ByVal psStats As Long) As COMPONENT_TYPE
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_compType = asBodyStats(i).compType
    Case STAT_BRAIN: COMPONENT_STATS_compType = asBrainStats(i).compType
    Case STAT_PROPULSION: COMPONENT_STATS_compType = asPropulsionStats(i).compType
    Case STAT_SENSOR: COMPONENT_STATS_compType = asSensorStats(i).compType
    Case STAT_ECM: COMPONENT_STATS_compType = asECMStats(i).compType
    Case STAT_REPAIR: COMPONENT_STATS_compType = asRepairStats(i).compType
    Case STAT_WEAPON: COMPONENT_STATS_compType = asWeaponStats(i).compType
    Case STAT_CONSTRUCT: COMPONENT_STATS_compType = asConstructStats(i).compType
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_compType(ByVal psStats As Long, ByVal value As COMPONENT_TYPE)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).compType = value
    Case STAT_BRAIN: asBrainStats(i).compType = value
    Case STAT_PROPULSION: asPropulsionStats(i).compType = value
    Case STAT_SENSOR: asSensorStats(i).compType = value
    Case STAT_ECM: asECMStats(i).compType = value
    Case STAT_REPAIR: asRepairStats(i).compType = value
    Case STAT_WEAPON: asWeaponStats(i).compType = value
    Case STAT_CONSTRUCT: asConstructStats(i).compType = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_droidTypeOverride(ByVal psStats As Long) As DROID_TYPE
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_droidTypeOverride = asBodyStats(i).droidTypeOverride
    Case STAT_BRAIN: COMPONENT_STATS_droidTypeOverride = asBrainStats(i).droidTypeOverride
    Case STAT_PROPULSION: COMPONENT_STATS_droidTypeOverride = asPropulsionStats(i).droidTypeOverride
    Case STAT_SENSOR: COMPONENT_STATS_droidTypeOverride = asSensorStats(i).droidTypeOverride
    Case STAT_ECM: COMPONENT_STATS_droidTypeOverride = asECMStats(i).droidTypeOverride
    Case STAT_REPAIR: COMPONENT_STATS_droidTypeOverride = asRepairStats(i).droidTypeOverride
    Case STAT_WEAPON: COMPONENT_STATS_droidTypeOverride = asWeaponStats(i).droidTypeOverride
    Case STAT_CONSTRUCT: COMPONENT_STATS_droidTypeOverride = asConstructStats(i).droidTypeOverride
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_droidTypeOverride(ByVal psStats As Long, ByVal value As DROID_TYPE)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).droidTypeOverride = value
    Case STAT_BRAIN: asBrainStats(i).droidTypeOverride = value
    Case STAT_PROPULSION: asPropulsionStats(i).droidTypeOverride = value
    Case STAT_SENSOR: asSensorStats(i).droidTypeOverride = value
    Case STAT_ECM: asECMStats(i).droidTypeOverride = value
    Case STAT_REPAIR: asRepairStats(i).droidTypeOverride = value
    Case STAT_WEAPON: asWeaponStats(i).droidTypeOverride = value
    Case STAT_CONSTRUCT: asConstructStats(i).droidTypeOverride = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_designable(ByVal psStats As Long) As Boolean
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_designable = asBodyStats(i).designable
    Case STAT_BRAIN: COMPONENT_STATS_designable = asBrainStats(i).designable
    Case STAT_PROPULSION: COMPONENT_STATS_designable = asPropulsionStats(i).designable
    Case STAT_SENSOR: COMPONENT_STATS_designable = asSensorStats(i).designable
    Case STAT_ECM: COMPONENT_STATS_designable = asECMStats(i).designable
    Case STAT_REPAIR: COMPONENT_STATS_designable = asRepairStats(i).designable
    Case STAT_WEAPON: COMPONENT_STATS_designable = asWeaponStats(i).designable
    Case STAT_CONSTRUCT: COMPONENT_STATS_designable = asConstructStats(i).designable
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_designable(ByVal psStats As Long, ByVal value As Boolean)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).designable = value
    Case STAT_BRAIN: asBrainStats(i).designable = value
    Case STAT_PROPULSION: asPropulsionStats(i).designable = value
    Case STAT_SENSOR: asSensorStats(i).designable = value
    Case STAT_ECM: asECMStats(i).designable = value
    Case STAT_REPAIR: asRepairStats(i).designable = value
    Case STAT_WEAPON: asWeaponStats(i).designable = value
    Case STAT_CONSTRUCT: asConstructStats(i).designable = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_usageClass(ByVal psStats As Long) As UsageClass
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_usageClass = asBodyStats(i).usageClass
    Case STAT_BRAIN: COMPONENT_STATS_usageClass = asBrainStats(i).usageClass
    Case STAT_PROPULSION: COMPONENT_STATS_usageClass = asPropulsionStats(i).usageClass
    Case STAT_SENSOR: COMPONENT_STATS_usageClass = asSensorStats(i).usageClass
    Case STAT_ECM: COMPONENT_STATS_usageClass = asECMStats(i).usageClass
    Case STAT_REPAIR: COMPONENT_STATS_usageClass = asRepairStats(i).usageClass
    Case STAT_WEAPON: COMPONENT_STATS_usageClass = asWeaponStats(i).usageClass
    Case STAT_CONSTRUCT: COMPONENT_STATS_usageClass = asConstructStats(i).usageClass
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_usageClass(ByVal psStats As Long, ByVal value As UsageClass)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).usageClass = value
    Case STAT_BRAIN: asBrainStats(i).usageClass = value
    Case STAT_PROPULSION: asPropulsionStats(i).usageClass = value
    Case STAT_SENSOR: asSensorStats(i).usageClass = value
    Case STAT_ECM: asECMStats(i).usageClass = value
    Case STAT_REPAIR: asRepairStats(i).usageClass = value
    Case STAT_WEAPON: asWeaponStats(i).usageClass = value
    Case STAT_CONSTRUCT: asConstructStats(i).usageClass = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_getBase_hitpoints(ByVal psStats As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_getBase_hitpoints = asBodyStats(i).base.hitpoints
    Case STAT_BRAIN: COMPONENT_STATS_getBase_hitpoints = asBrainStats(i).base.hitpoints
    Case STAT_PROPULSION: COMPONENT_STATS_getBase_hitpoints = asPropulsionStats(i).base.hitpoints
    Case STAT_SENSOR: COMPONENT_STATS_getBase_hitpoints = asSensorStats(i).base.hitpoints
    Case STAT_ECM: COMPONENT_STATS_getBase_hitpoints = asECMStats(i).base.hitpoints
    Case STAT_REPAIR: COMPONENT_STATS_getBase_hitpoints = asRepairStats(i).base.hitpoints
    Case STAT_WEAPON: COMPONENT_STATS_getBase_hitpoints = asWeaponStats(i).base.hitpoints
    Case STAT_CONSTRUCT: COMPONENT_STATS_getBase_hitpoints = asConstructStats(i).base.hitpoints
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_getBase_hitpoints(ByVal psStats As Long, ByVal value As Long)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).base.hitpoints = value
    Case STAT_BRAIN: asBrainStats(i).base.hitpoints = value
    Case STAT_PROPULSION: asPropulsionStats(i).base.hitpoints = value
    Case STAT_SENSOR: asSensorStats(i).base.hitpoints = value
    Case STAT_ECM: asECMStats(i).base.hitpoints = value
    Case STAT_REPAIR: asRepairStats(i).base.hitpoints = value
    Case STAT_WEAPON: asWeaponStats(i).base.hitpoints = value
    Case STAT_CONSTRUCT: asConstructStats(i).base.hitpoints = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_getBase_hitpointPct(ByVal psStats As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_getBase_hitpointPct = asBodyStats(i).base.hitpointPct
    Case STAT_BRAIN: COMPONENT_STATS_getBase_hitpointPct = asBrainStats(i).base.hitpointPct
    Case STAT_PROPULSION: COMPONENT_STATS_getBase_hitpointPct = asPropulsionStats(i).base.hitpointPct
    Case STAT_SENSOR: COMPONENT_STATS_getBase_hitpointPct = asSensorStats(i).base.hitpointPct
    Case STAT_ECM: COMPONENT_STATS_getBase_hitpointPct = asECMStats(i).base.hitpointPct
    Case STAT_REPAIR: COMPONENT_STATS_getBase_hitpointPct = asRepairStats(i).base.hitpointPct
    Case STAT_WEAPON: COMPONENT_STATS_getBase_hitpointPct = asWeaponStats(i).base.hitpointPct
    Case STAT_CONSTRUCT: COMPONENT_STATS_getBase_hitpointPct = asConstructStats(i).base.hitpointPct
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Let COMPONENT_STATS_getBase_hitpointPct(ByVal psStats As Long, ByVal value As Long)
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: asBodyStats(i).base.hitpointPct = value
    Case STAT_BRAIN: asBrainStats(i).base.hitpointPct = value
    Case STAT_PROPULSION: asPropulsionStats(i).base.hitpointPct = value
    Case STAT_SENSOR: asSensorStats(i).base.hitpointPct = value
    Case STAT_ECM: asECMStats(i).base.hitpointPct = value
    Case STAT_REPAIR: asRepairStats(i).base.hitpointPct = value
    Case STAT_WEAPON: asWeaponStats(i).base.hitpointPct = value
    Case STAT_CONSTRUCT: asConstructStats(i).base.hitpointPct = value
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_getUpgrade_hitpoints(ByVal psStats As Long, ByVal player As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_getUpgrade_hitpoints = asBodyStats(i).upgrade(player).hitpoints
    Case STAT_BRAIN: COMPONENT_STATS_getUpgrade_hitpoints = asBrainStats(i).upgrade(player).hitpoints
    Case STAT_PROPULSION: COMPONENT_STATS_getUpgrade_hitpoints = asPropulsionStats(i).upgrade(player).hitpoints
    Case STAT_SENSOR: COMPONENT_STATS_getUpgrade_hitpoints = asSensorStats(i).upgrade(player).hitpoints
    Case STAT_ECM: COMPONENT_STATS_getUpgrade_hitpoints = asECMStats(i).upgrade(player).hitpoints
    Case STAT_REPAIR: COMPONENT_STATS_getUpgrade_hitpoints = asRepairStats(i).upgrade(player).hitpoints
    Case STAT_WEAPON: COMPONENT_STATS_getUpgrade_hitpoints = asWeaponStats(i).upgrade(player).hitpoints
    Case STAT_CONSTRUCT: COMPONENT_STATS_getUpgrade_hitpoints = asConstructStats(i).upgrade(player).hitpoints
    Case Else: BadStatsHandle psStats
    End Select
End Property

Public Property Get COMPONENT_STATS_getUpgrade_hitpointPct(ByVal psStats As Long, ByVal player As Long) As Long
    Dim i As Long
    i = psStats And &HFFFF&
    Select Case psStats And STAT_MASK
    Case STAT_BODY: COMPONENT_STATS_getUpgrade_hitpointPct = asBodyStats(i).upgrade(player).hitpointPct
    Case STAT_BRAIN: COMPONENT_STATS_getUpgrade_hitpointPct = asBrainStats(i).upgrade(player).hitpointPct
    Case STAT_PROPULSION: COMPONENT_STATS_getUpgrade_hitpointPct = asPropulsionStats(i).upgrade(player).hitpointPct
    Case STAT_SENSOR: COMPONENT_STATS_getUpgrade_hitpointPct = asSensorStats(i).upgrade(player).hitpointPct
    Case STAT_ECM: COMPONENT_STATS_getUpgrade_hitpointPct = asECMStats(i).upgrade(player).hitpointPct
    Case STAT_REPAIR: COMPONENT_STATS_getUpgrade_hitpointPct = asRepairStats(i).upgrade(player).hitpointPct
    Case STAT_WEAPON: COMPONENT_STATS_getUpgrade_hitpointPct = asWeaponStats(i).upgrade(player).hitpointPct
    Case STAT_CONSTRUCT: COMPONENT_STATS_getUpgrade_hitpointPct = asConstructStats(i).upgrade(player).hitpointPct
    Case Else: BadStatsHandle psStats
    End Select
End Property
' Check if an object has a weapon
Public Function objHasWeapon(ByVal psObj As Long) As Boolean
    'check if valid type
    If SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        If aDroid(psObj).numWeaps > 0 Then
            objHasWeapon = True
            Exit Function
        End If
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Then
        If aStruct(psObj).numWeaps > 0 Then
            objHasWeapon = True
            Exit Function
        End If
    End If

    objHasWeapon = False
End Function

' SENSOR_STATS *objActiveRadar(const BASE_OBJECT *psObj): the asSensorStats index (-1: nullptr)
Public Function objActiveRadar(ByVal psObj As Long) As Long
    Dim psStats As Long, compIndex As Long

    psStats = -1

    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        If aDroid(psObj).droidType <> DROID_SENSOR And aDroid(psObj).droidType <> DROID_COMMAND Then
            objActiveRadar = -1
            Exit Function
        End If
        compIndex = aDroid(psObj).asBits(COMP_SENSOR)
        If Not (compIndex < asSensorStats_size) Then ASSERT_FAILURE "compIndex < asSensorStats.size()", "src/stats.cpp:1841", "objActiveRadar", "Invalid range referenced for numSensorStats, %d > %zu", compIndex, asSensorStats_size: objActiveRadar = -1: Exit Function
        psStats = compIndex
    Case OBJ_STRUCTURE
        psStats = asStructureStats(aStruct(psObj).pStructureType).pSensor
        If psStats = -1 Then
            objActiveRadar = -1
            Exit Function
        End If
        If asSensorStats(psStats).location <> LOC_TURRET Or aStruct(psObj).status <> SS_BUILT Then
            objActiveRadar = -1
            Exit Function
        End If
    Case Else
    End Select
    objActiveRadar = psStats
End Function

Public Function objRadarDetector(ByVal psObj As Long) As Boolean
    Dim psStruct As Long, psDroid As Long, psSensor As Long
    If SIMPLE_OBJECT_type(psObj) = OBJ_STRUCTURE Then
        psStruct = psObj

        If aStruct(psStruct).status = SS_BUILT And asStructureStats(aStruct(psStruct).pStructureType).pSensor <> -1 Then
            objRadarDetector = asSensorStats(asStructureStats(aStruct(psStruct).pStructureType).pSensor).typ = RADAR_DETECTOR_SENSOR
        End If
        Exit Function
    ElseIf SIMPLE_OBJECT_type(psObj) = OBJ_DROID Then
        psDroid = psObj
        psSensor = DROID_getSensorStats(psDroid)

        objRadarDetector = (psSensor <> -1)
        If objRadarDetector Then objRadarDetector = asSensorStats(psSensor).typ = RADAR_DETECTOR_SENSOR
        Exit Function
    End If
    objRadarDetector = False
End Function
