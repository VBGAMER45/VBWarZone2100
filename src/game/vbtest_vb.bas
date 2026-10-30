Attribute VB_Name = "vbtest_vb"
' (VB-only) --vbtest=<command file>: the test commands of tools\run_tests.py,
' which realmain runs after the search path is set up and before the
' window opens; then the process ends. The commands, their output and the
' dump are tools\statsref's (tools\statsref\statsref.cpp), which runs the
' same file on the real C++: one command per line, its words separated by
' tabs; each command writes "> " and its line, the log lines it caused
' (debug_c's capture callback) and its results. "searchpath <file>" writes
' this process's search path into <file> (for statsref's own searchpath
' command); "output <file>" starts the file the next commands write to.
' The map commands (overrides, mapload, mapreloadground, mapqueries,
' mapshutdown), the object commands (mapplayers, mapobjects, objshutdown)
' and the simulation's (syncreset, synclog, mapgrid, mappower, mappaths, mapmove, maporders, mapcombat,
' mapbase) are listed in
' tools\statsref\statsref.cpp.
Option Explicit

Public vbTestCommandFile As String

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private outChunks() As String           ' the current output (joined when written)
Private outCount As Long
Private outFile As String
Private terrainLoaded As Boolean
Private testMap As Map                  ' the last map mapload loaded (its objects for mapobjects)
Private syncDumpAll As Boolean          ' syncdumpall: mapmove, maporders, mapcombat, mapbase dump every tick's sync log
' (mapbase) what research and the test change in the stats the maps share, restored after the test
Private savedResList() As PLAYER_RESEARCH
Private savedCompLists() As Byte
Private savedStructTypeLists() As Byte
Private savedDefaultSensor(0 To MAX_PLAYERS - 1) As Long
Private savedDefaultECM(0 To MAX_PLAYERS - 1) As Long
Private savedDefaultRepair(0 To MAX_PLAYERS - 1) As Long
Private savedWeaponStats() As WEAPON_STATS
Private savedBodyStats() As BODY_STATS
Private savedBrainStats() As BRAIN_STATS
Private savedPropulsionStats() As PROPULSION_STATS
Private savedSensorStats() As SENSOR_STATS
Private savedECMStats() As ECM_STATS
Private savedRepairStats() As REPAIR_STATS
Private savedConstructStats() As CONSTRUCT_STATS
Private savedStructureStats() As STRUCTURE_STATS
Private moduleToBuilding(0 To MAX_PLAYER_SLOTS - 1) As Object
Private fixedIds As Object              ' loadGame's fixedMapIdToGeneratedId
Private netTicks As Long                ' mapnet: what wzGetTicks gives (it only moves in mapnet, 100 ms a frame, across the maps)
Private netReplayFile As String         ' maprecord's replay, for mapplay
Private Const NET_MODE_NET As Long = 0     ' netRun's modes: mapnet,
Private Const NET_MODE_RECORD As Long = 1  ' maprecord (mapnet's game, recorded as a replay),
Private Const NET_MODE_PLAY As Long = 2    ' mapplay (that replay played back)
Private Const MAX_SYNC_HISTORY As Long = 12                         ' sync_debug.cpp
Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS

Private Sub putOut(ByVal s As String)
    If outCount > UBound(outChunks) Then ReDim Preserve outChunks(0 To outCount * 2)
    outChunks(outCount) = s
    outCount = outCount + 1
End Sub

Private Sub putCaptured()
    putOut debug_capture_take()
End Sub

Private Sub flushOutput()
    Dim f As Long
    If Len(outFile) = 0 Then Exit Sub
    f = fopen(outFile, "wb")
    If f <> 0 Then
        If outCount > 0 Then
            ReDim Preserve outChunks(0 To outCount - 1)
            fputs Join(outChunks, ""), f
        End If
        fclose f
    End If
End Sub

Private Sub resetOutput()
    ReDim outChunks(0 To 255)
    outCount = 0
End Sub

' an iIMDBaseShape *: its file name
Private Function modelStr(ByVal m As Long) As String
    If m = 0 Then modelStr = "null" Else modelStr = iIMDBaseShape_filename(m)
End Function

' a BASE_STATS * as "<id>/<ref>"
Private Function statRef(ByVal s As Long) As String
    If s = 0 Then
        statRef = "null"
    Else
        statRef = va("%s/%x", BASE_STATS_id(s), BASE_STATS_ref(s))
    End If
End Function

' a pointer into a stats array (its index, -1 = nullptr)
Private Function indexStr(ByVal p As Long) As String
    If p = -1 Then indexStr = "null" Else indexStr = CStr(p)
End Function

Private Sub putCommon(ByVal kind As String, ByVal i As Long, ByVal h As Long)
    putOut va("%s %u|%s|%s|ref=%x|index=%u|imd=%s|power=%u|points=%u|weight=%u|comp=%d|dtype=%d|design=%d|usage=%d\n", _
        kind, i, BASE_STATS_id(h), BASE_STATS_name(h), BASE_STATS_ref(h), BASE_STATS_index(h), modelStr(COMPONENT_STATS_pIMD(h)), _
        COMPONENT_STATS_buildPower(h), COMPONENT_STATS_buildPoints(h), COMPONENT_STATS_weight(h), COMPONENT_STATS_compType(h), _
        COMPONENT_STATS_droidTypeOverride(h), -CLng(COMPONENT_STATS_designable(h)), COMPONENT_STATS_usageClass(h))
End Sub

' " base<line>" and " upgrade <j><line>" for each player's upgrade that differs from the base
Private Sub putUpgrades(baseLine As String, upgradeLines() As String)
    Dim j As Long
    putOut " base" & baseLine & vbLf
    For j = 0 To MAX_PLAYERS - 1
        If upgradeLines(j) <> baseLine Then putOut va(" upgrade %d", j) & upgradeLines(j) & vbLf
    Next
End Sub

Private Function weaponUpgrade(u As WEAPON_STATS_UPGRADE) As String
    weaponUpgrade = va("|hp=%u|hpPct=%d|short=%u|max=%u|min=%u|hit=%u|shortHit=%u|pause=%u|rounds=%u|reload=%u|damage=%u|emp=%u|radius=%u|radDamage=%u|pd=%u|pdRadius=%u|pdTime=%u|minDamage=%u", _
        u.hitpoints, u.hitpointPct, u.shortRange, u.maxRange, u.minRange, u.hitChance, u.shortHitChance, u.firePause, u.numRounds, _
        u.reloadTime, u.damage, u.empRadius, u.radius, u.radiusDamage, u.periodicalDamage, u.periodicalDamageRadius, u.periodicalDamageTime, u.minimumDamage)
End Function

Private Function bodyUpgrade(u As BODY_STATS_UPGRADE) As String
    bodyUpgrade = va("|hp=%u|hpPct=%d|power=%u|armour=%u|thermal=%d|resistance=%d", u.hitpoints, u.hitpointPct, u.power, u.armour, u.thermal, u.resistance)
End Function

Private Function brainUpgrade(u As BRAIN_STATS_UPGRADE) As String
    Dim t As String, k As Long
    t = va("|hp=%u|hpPct=%d|maxDroids=%d|maxDroidsMult=%d|thresholds", u.hitpoints, u.hitpointPct, u.maxDroids, u.maxDroidsMult)
    For k = 0 To u.rankThresholds_size - 1
        t = t & va(" %d", u.rankThresholds(k))
    Next
    brainUpgrade = t
End Function

Private Function structureUpgrade(u As STRUCTURE_STATS_UPGRADE) As String
    structureUpgrade = va("|research=%u|moduleResearch=%u|repair=%u|power=%u|modulePower=%u|production=%u|moduleProduction=%u|rearm=%u|armour=%u|thermal=%u|hp=%u|resistance=%u|limit=%u", _
        u.research, u.moduleResearch, u.repair, u.power, u.modulePower, u.production, u.moduleProduction, u.rearm, u.armour, u.thermal, u.hitpoints, _
        u.resistance, u.limit)
End Function

Private Sub putList(ByVal label As String, list() As Long, ByVal size As Long)
    Dim k As Long, t As String
    t = va(" %s %u", label, size)
    For k = 0 To size - 1
        If list(k) <> 0 Then t = t & va(" %u:%s", k, modelStr(list(k)))
    Next
    putOut t & vbLf
End Sub

Private Sub putTemplate(t As DROID_TEMPLATE)
    Dim k As Long, s As String
    s = va("template %s|%s|ref=%x|mpid=%u|type=%d|parts", t.id, t.name_, t.ref, t.multiPlayerID, t.droidType)
    For k = 0 To DROID_MAXCOMP - 1
        s = s & va(" %u", t.asParts(k))
    Next
    s = s & va("|weapons %d", t.numWeaps)
    For k = 0 To MAX_WEAPONS - 1
        s = s & va(" %u", t.asWeaps(k))
    Next
    s = s & va("|prefab=%d|stored=%d|enabled=%d\n", -CLng(t.prefab), -CLng(t.stored), -CLng(t.enabled))
    putOut s
End Sub

Private Sub dump()
    Dim i As Long, j As Long, k As Long, t As Long, p As Long, c As Long, n As Long, s As String
    Dim ups(0 To MAX_PLAYERS - 1) As String

    putOut va("== weapons %u\n", asWeaponStats_size)
    For i = 0 To asWeaponStats_size - 1
        putCommon "weapon", i, STAT_WEAPON + i
        For j = 0 To MAX_PLAYERS - 1
            ups(j) = weaponUpgrade(asWeaponStats(i).upgrade(j))
        Next
        putUpgrades weaponUpgrade(asWeaponStats(i).base), ups
        With asWeaponStats(i)
            putOut va(" pdClass=%d|pdSub=%d|pdEffect=%d|class=%d|sub=%d|movement=%d|effect=%d|size=%d|flight=%u|recoil=%u|ext=%d|rotate=%d|maxElev=%d|minElev=%d|effectSize=%u|runs=%d|face=%u|faceFlight=%u|air=%u|light=%d|penetrate=%d|fireOnMove=%d|flags=%u|radiusLife=%u|explosions=%u\n", _
                .periodicalDamageWeaponClass, .periodicalDamageWeaponSubClass, .periodicalDamageWeaponEffect, .weaponClass, .weaponSubClass, _
                .movementModel, .weaponEffect, .weaponSize, .flightSpeed, .recoilValue, .distanceExtensionFactor, .rotate, .maxElevation, _
                .minElevation, .effectSize, .vtolAttackRuns, .facePlayer, .faceInFlight, .surfaceToAir, -CLng(.lightWorld), _
                -CLng(.penetrate), -CLng(.fireOnMove), .flags, .radiusLife, .numExplosions)
            putOut va(" mount=%s|muzzle=%s|flight=%s|hit=%s|miss=%s|water=%s|trail=%s|fire=%d|impact=%d\n", _
                modelStr(.pMountGraphic), modelStr(.pMuzzleGraphic), modelStr(.pInFlightGraphic), modelStr(.pTargetHitGraphic), _
                modelStr(.pTargetMissGraphic), modelStr(.pWaterHitGraphic), modelStr(.pTrailGraphic), .iAudioFireID, .iAudioImpactID)
        End With
    Next
    putOut va("== bodies %u\n", asBodyStats_size)
    For i = 0 To asBodyStats_size - 1
        putCommon "body", i, STAT_BODY + i
        For j = 0 To MAX_PLAYERS - 1
            ups(j) = bodyUpgrade(asBodyStats(i).upgrade(j))
        Next
        putUpgrades bodyUpgrade(asBodyStats(i).base), ups
        putOut va(" size=%d|slots=%u|class=%s\n", asBodyStats(i).size, asBodyStats(i).weaponSlots, asBodyStats(i).bodyClass)
        putList "imd", asBodyStats(i).ppIMDList, asBodyStats(i).ppIMDList_size
        putList "move", asBodyStats(i).ppMoveIMDList, asBodyStats(i).ppMoveIMDList_size
        putList "still", asBodyStats(i).ppStillIMDList, asBodyStats(i).ppStillIMDList_size
    Next
    putOut va("== brains %u\n", asBrainStats_size)
    For i = 0 To asBrainStats_size - 1
        putCommon "brain", i, STAT_BRAIN + i
        For j = 0 To MAX_PLAYERS - 1
            ups(j) = brainUpgrade(asBrainStats(i).upgrade(j))
        Next
        putUpgrades brainUpgrade(asBrainStats(i).base), ups
        s = va(" weapon=%s|ranks", indexStr(asBrainStats(i).psWeaponStat))
        For k = 0 To asBrainStats(i).rankNames.size - 1
            s = s & " " & asBrainStats(i).rankNames.items(k)
        Next
        s = s & "|cmdExpRange"
        For k = 0 To asBrainStats(i).cmdExpRange_size - 1
            s = s & va(" %d", asBrainStats(i).cmdExpRange(k))
        Next
        putOut s & vbLf
    Next
    putOut va("== propulsion %u\n", asPropulsionStats_size)
    For i = 0 To asPropulsionStats_size - 1
        putCommon "propulsion", i, STAT_PROPULSION + i
        For j = 0 To MAX_PLAYERS - 1
            ups(j) = va("|hp=%u|hpPct=%d|pctOfBody=%d", asPropulsionStats(i).upgrade(j).hitpoints, asPropulsionStats(i).upgrade(j).hitpointPct, asPropulsionStats(i).upgrade(j).hitpointPctOfBody)
        Next
        putUpgrades va("|hp=%u|hpPct=%d|pctOfBody=%d", asPropulsionStats(i).base.hitpoints, asPropulsionStats(i).base.hitpointPct, asPropulsionStats(i).base.hitpointPctOfBody), ups
        With asPropulsionStats(i)
            putOut va(" type=%d|speed=%u|turn=%u|spin=%u|spinAngle=%u|skid=%u|decel=%u|accel=%u\n", .propulsionType, .maxSpeed, .turnSpeed, .spinSpeed, _
                .spinAngle, .skidDeceleration, .deceleration, .acceleration)
        End With
    Next
    putOut va("== sensors %u\n", asSensorStats_size)
    For i = 0 To asSensorStats_size - 1
        putCommon "sensor", i, STAT_SENSOR + i
        For j = 0 To MAX_PLAYERS - 1
            ups(j) = va("|hp=%u|hpPct=%d|range=%u", asSensorStats(i).upgrade(j).hitpoints, asSensorStats(i).upgrade(j).hitpointPct, asSensorStats(i).upgrade(j).range)
        Next
        putUpgrades va("|hp=%u|hpPct=%d|range=%u", asSensorStats(i).base.hitpoints, asSensorStats(i).base.hitpointPct, asSensorStats(i).base.range), ups
        putOut va(" mount=%s|location=%u|type=%d\n", modelStr(asSensorStats(i).pMountGraphic), asSensorStats(i).location, asSensorStats(i).typ)
    Next
    putOut va("== ecm %u\n", asECMStats_size)
    For i = 0 To asECMStats_size - 1
        putCommon "ecm", i, STAT_ECM + i
        For j = 0 To MAX_PLAYERS - 1
            ups(j) = va("|hp=%u|hpPct=%d|range=%u", asECMStats(i).upgrade(j).hitpoints, asECMStats(i).upgrade(j).hitpointPct, asECMStats(i).upgrade(j).range)
        Next
        putUpgrades va("|hp=%u|hpPct=%d|range=%u", asECMStats(i).base.hitpoints, asECMStats(i).base.hitpointPct, asECMStats(i).base.range), ups
        putOut va(" mount=%s|location=%u\n", modelStr(asECMStats(i).pMountGraphic), asECMStats(i).location)
    Next
    putOut va("== repair %u\n", asRepairStats_size)
    For i = 0 To asRepairStats_size - 1
        putCommon "repair", i, STAT_REPAIR + i
        For j = 0 To MAX_PLAYERS - 1
            ups(j) = va("|hp=%u|hpPct=%d|points=%u", asRepairStats(i).upgrade(j).hitpoints, asRepairStats(i).upgrade(j).hitpointPct, asRepairStats(i).upgrade(j).repairPoints)
        Next
        putUpgrades va("|hp=%u|hpPct=%d|points=%u", asRepairStats(i).base.hitpoints, asRepairStats(i).base.hitpointPct, asRepairStats(i).base.repairPoints), ups
        putOut va(" mount=%s|location=%u|time=%u\n", modelStr(asRepairStats(i).pMountGraphic), asRepairStats(i).location, asRepairStats(i).time_)
    Next
    putOut va("== construct %u\n", asConstructStats_size)
    For i = 0 To asConstructStats_size - 1
        putCommon "construct", i, STAT_CONSTRUCT + i
        For j = 0 To MAX_PLAYERS - 1
            ups(j) = va("|hp=%u|hpPct=%d|points=%u", asConstructStats(i).upgrade(j).hitpoints, asConstructStats(i).upgrade(j).hitpointPct, asConstructStats(i).upgrade(j).constructPoints)
        Next
        putUpgrades va("|hp=%u|hpPct=%d|points=%u", asConstructStats(i).base.hitpoints, asConstructStats(i).base.hitpointPct, asConstructStats(i).base.constructPoints), ups
        putOut va(" mount=%s\n", modelStr(asConstructStats(i).pMountGraphic))
    Next
    putOut va("== propulsion types %u\n", asPropulsionTypes_size)
    For i = 0 To asPropulsionTypes_size - 1
        With asPropulsionTypes(i)
            putOut va("proptype %u|travel=%d|mult=%u|start=%d|idle=%d|moveOff=%d|move=%d|hiss=%d|shutDown=%d\n", i, .travel, .powerRatioMult, _
                .startID, .idleID, .moveOffID, .moveID, .hissID, .shutDownID)
        End With
    Next
    If terrainLoaded Then
        putOut "== terrain" & vbLf
        For t = 0 To TER_MAX - 1
            s = va("terrain %d", t)
            For p = 0 To PROPULSION_TYPE_NUM - 1
                s = s & va(" %u", getSpeedFactor(t, p))
            Next
            putOut s & vbLf
        Next
    End If
    putOut "== modifiers" & vbLf
    For i = 0 To WE_NUMEFFECTS - 1
        s = va("effect %d|prop", i)
        For p = 0 To PROPULSION_TYPE_NUM - 1
            s = s & va(" %u", asWeaponModifier(i, p))
        Next
        s = s & "|body"
        For k = 0 To SIZE_NUM - 1
            s = s & va(" %u", asWeaponModifierBody(i, k))
        Next
        s = s & "|struct"
        For k = 0 To NUM_STRUCT_STRENGTH - 1
            s = s & va(" %u", asStructStrengthModifier(i, k))
        Next
        putOut s & vbLf
    Next
    putOut "== component lists" & vbLf
    For t = 0 To COMP_NUMCOMPONENTS - 1
        If apCompLists_count(t) = 0 Then GoTo next_type
        Select Case t
        Case COMP_BODY: n = asBodyStats_size
        Case COMP_BRAIN: n = asBrainStats_size
        Case COMP_PROPULSION: n = asPropulsionStats_size
        Case COMP_REPAIRUNIT: n = asRepairStats_size
        Case COMP_ECM: n = asECMStats_size
        Case COMP_SENSOR: n = asSensorStats_size
        Case COMP_CONSTRUCT: n = asConstructStats_size
        Case COMP_WEAPON: n = asWeaponStats_size
        End Select
        For p = 0 To MAX_PLAYERS - 1
            s = va("comp %d player %d:", t, p)
            For c = 0 To n - 1
                s = s & va(" %u", apCompLists(p, t, c))
            Next
            putOut s & vbLf
        Next
next_type:
    Next
    If apStructTypeLists_allocated Then
        For p = 0 To MAX_PLAYERS - 1
            s = va("structlist player %d:", p)
            For c = 0 To numStructureStats - 1
                s = s & va(" %u", apStructTypeLists(p, c))
            Next
            putOut s & vbLf
        Next
    End If
    putOut va("== structures %u|factoryModule=%u|powerModule=%u|researchModule=%u|baseExp=%d\n", numStructureStats, factoryModuleStat, powerModuleStat, _
        researchModuleStat, getStructureDamageBaseExperienceLevel())
    For i = 0 To numStructureStats - 1
        With asStructureStats(i)
            s = va("structure %u|%s|%s|ref=%x|index=%u|type=%d|strength=%d|width=%u|breadth=%u|points=%u|height=%u|power=%u|imds", i, .id, _
                .name_, .ref, .index, .typ, .strength, .baseWidth, .baseBreadth, .buildPoints, .height, .powerToBuild)
            For k = 0 To .pIMD_size - 1
                s = s & " " & modelStr(.pIMD(k))
            Next
            s = s & va("|base=%s|ecm=%s|sensor=%s|weapons=%u", modelStr(.pBaseIMD), indexStr(.pECM), indexStr(.pSensor), .numWeaps)
            For k = 0 To MAX_WEAPONS - 1
                s = s & " " & indexStr(.psWeapStat(k))
            Next
            s = s & va("|flags=%u|walls=%d|minLimit=%u|maxLimit=%u|favorite=%d\n", .flags, -CLng(.combinesWithWall), .minLimit, .maxLimit, -CLng(.isFavorite))
        End With
        putOut s
        For j = 0 To MAX_PLAYERS - 1
            ups(j) = structureUpgrade(asStructureStats(i).upgrade(j))
        Next
        putUpgrades structureUpgrade(asStructureStats(i).base), ups
    Next
    putOut va("== features %u|oil=%s\n", asFeatureStats_size, indexStr(oilResFeature))
    For i = 0 To asFeatureStats_size - 1
        With asFeatureStats(i)
            putOut va("feature %u|%s|%s|ref=%x|index=%u|type=%d|imd=%s|width=%u|breadth=%u|tile=%d|los=%d|visible=%d|damageable=%d|body=%u|armour=%u\n", i, _
                .id, .name_, .ref, .index, .subType, modelStr(.psImd), .baseWidth, _
                .baseBreadth, -CLng(.tileDraw), -CLng(.allowLOS), -CLng(.visibleAtStart), -CLng(.damageable), .body, .armourValue)
        End With
    Next
    putOut va("== research %u|mode=%d\n", asResearch_size, getResearchUpgradeCalcMode())
    For i = 0 To asResearch_size - 1
        With asResearch(i)
            s = va("research %u|%s|%s|ref=%x|baseIndex=%u|index=%d|tech=%u|subGroup=%u|points=%u|power=%u|key=%u|noCheat=%u|disabled=%u|icon=%u|stat=%s|imd=%s|imd2=%s|view=%s|category=%s", _
                i, .id, .name_, .ref, .base_index, .index, .techCode, _
                .subGroup, .researchPoints, .researchPower, .keyTopic, .excludeFromCheats, .disabledWhen, _
                .iconID, statRef(.psStat), modelStr(.pIMD), modelStr(.pIMD2), IIf(.pViewData <> 0, "set", "null"), .category)
            If Len(.category) > 0 Then
                s = s & va(" %u/%u", .categoryProgress, .categoryMax)
            End If
            s = s & vbLf & " pre"
            For k = 0 To .pPRList_size - 1
                s = s & va(" %u", .pPRList(k))
            Next
            s = s & "|reqStruct"
            For k = 0 To .pStructList_size - 1
                s = s & va(" %u", .pStructList(k))
            Next
            s = s & "|redStruct"
            For k = 0 To .pRedStructs_size - 1
                s = s & va(" %u", .pRedStructs(k))
            Next
            s = s & "|resStruct"
            For k = 0 To .pStructureResults_size - 1
                s = s & va(" %u", .pStructureResults(k))
            Next
            s = s & vbLf & " redComp"
            For k = 0 To .pRedArtefacts_size - 1
                s = s & " " & statRef(.pRedArtefacts(k))
            Next
            s = s & "|resComp"
            For k = 0 To .componentResults_size - 1
                s = s & " " & statRef(.componentResults(k))
            Next
            s = s & "|replace"
            For k = 0 To .componentReplacement_size - 1
                s = s & " " & statRef(.componentReplacement(k).pOldComponent) & ">" & statRef(.componentReplacement(k).pNewComponent)
            Next
            s = s & vbLf & " results " & .results.dump() & vbLf
        End With
        putOut s
    Next
    For p = 0 To MAX_PLAYERS - 1
        putOut va("playerres %d %u\n", p, asPlayerResList_size(p))
    Next
    For p = 0 To MAX_PLAYERS - 1
        putOut va("== templates player %d: %u\n", p, templateCount(p))
        For k = 0 To templateCount(p) - 1
            putTemplate apsTemplates(droidTemplates_at(p, k))
        Next
    Next
    putOut va("== local templates %u\n", localTemplates_size)
    For k = 0 To localTemplates_size - 1
        putTemplate apsTemplates(localTemplates(k))
    Next
    putOut va("== object ids %u\n", unsynchObjID)
End Sub

' data.cpp's buffer loaders (without the data hash: bMultiPlayer is false)
'--------------------------------------------------------------------------
' the map (map.cpp)

' loadGame's map part: the map folder through wzmaplib, its terrain types, then mapLoadFromWzMapData
' (tileset "-": tilesetDir stays nullptr)
Private Function vbtestMapLoad(ByVal folder As String, ByVal tileset As String, ByVal players As Long, ByVal kind As String) As Boolean
    Dim m As Map, io As WzMapPhysFSIO, logger As WzMapDebugLogger, data As MapData, typ As MapType
    If tileset = "-" Then tilesetDir = vbNullString Else tilesetDir = tileset
    If kind = "campaign" Then typ = MapType_CAMPAIGN Else typ = MapType_SKIRMISH
    Set io = New WzMapPhysFSIO
    Set logger = New WzMapDebugLogger
    Set m = Map_loadFromPath(folder, typ, players, 0, logger, io)
    Set testMap = m
    If m Is Nothing Then
        putCaptured
        putOut "map=null" & vbLf
        Exit Function
    End If
    If Not loadTerrainTypeMap(m.mapTerrainTypes()) Then
        putCaptured
        putOut "ttypes failed" & vbLf
        Exit Function
    End If
    Set data = m.mapData()
    If data Is Nothing Then
        putCaptured
        putOut "mapdata=null" & vbLf
        Exit Function
    End If
    vbtestMapLoad = mapLoadFromWzMapData(data)
End Function

Private Sub mapDump()
    Dim i As Long, x As Long, y As Long, g As GROUND_TYPE, s As String, gw As GATEWAY, t As Long
    If StrPtr(tilesetDir) = 0 Then s = "(null)" Else s = tilesetDir
    putOut va("map width=%d height=%d scroll=%d %d %d %d tileset=%s current=%d groundTypes=%u\n", mapWidth, mapHeight, _
        scrollMinX, scrollMinY, scrollMaxX, scrollMaxY, s, currentMapTileset, getNumGroundTypes())
    For i = 0 To getNumGroundTypes() - 1
        g = getGroundType(i)
        putOut va("ground %u: %s %s %s|%s|%s %d\n", i, g.textureName, FloatBits(g.textureSize), g.normalMapTextureName, _
            g.specularMapTextureName, g.heightMapTextureName, -CLng(g.highQualityTextures))
    Next
    s = "ttypes:"
    For i = 0 To MAX_TILE_TEXTURES - 1
        s = s & " " & CStr(terrainTypes(i))
    Next
    putOut s & vbLf
    putOut va("gateways %u:", gwNumGateways())
    For i = 0 To gwNumGateways() - 1
        gwGetGateway i, gw
        putOut va(" %u,%u,%u,%u", CLng(gw.x1), CLng(gw.y1), CLng(gw.x2), CLng(gw.y2))
    Next
    putOut vbLf
    For y = 0 To mapHeight - 1
        putOut CStr(y) & ":"
        For x = 0 To mapWidth - 1
            t = x + y * mapWidth
            With psMapTiles(t)
                putOut " " & CStr(.texture) & "/" & CStr(.height) & "/" & CStr(.waterLevel) & "/" & CStr(.ground) & "/" & Hex$(.tileInfoBits) _
                    & "/" & CStr(.limitedContinent) & "/" & CStr(.hoverContinent) & "/" & Hex$(blockTile(x, y, 0))
            End With
        Next
        putOut vbLf
    Next
End Sub

' the objects after mapfeatures: the id counters, the feature list, the
' oil list, the derricks and start positions, and each tile's object and
' blocking bits
Private Sub objectsDump()
    Dim it As Long, f As Long, x As Long, y As Long, s As String, t As Long, n As Long
    putOut va("ids unsynch=%u synch=%u destroyed=%u\n", unsynchObjID, synchObjID, psDestroyedObj.size)
    putOut va("features %u\n", apsFeatureList(0).size)
    it = apsFeatureList(0).head
    Do While it <> 0
        f = lnObj(it)
        With aFeature(f)
            putOut va(" f id=%u stat=%s pos=%d,%d,%d rot=%u,%u,%u player=%u body=%u fd=%s born=%u died=%u imd=%s type=%d extra=%d\n", _
                .id, asFeatureStats(.psStats).id, .pos.x, .pos.y, .pos.z, .rot.direction, .rot.pitch, .rot.roll, .player, .body, _
                FloatBits(.foundationDepth), .born, .died, modelStr(.sDisplay.imd), .typ, -CLng(.hasExtraFunction))
        End With
        it = lnNext(it)
    Loop
    s = va("oil %u:", apsOilList(0).size)
    it = apsOilList(0).head
    Do While it <> 0
        s = s & " " & CStr(aFeature(lnObj(it)).id)
        it = lnNext(it)
    Loop
    putOut s & vbLf
    s = va("derricks %u:", derricks_count())
    For t = 0 To 65535
        If derricks_contains(t) Then s = s & va(" %u,%u", t \ 256, t And 255)
    Next
    putOut s & vbLf
    s = "starts:"
    For t = 0 To MAX_PLAYERS - 1
        s = s & va(" %d,%d", positions(t).x, positions(t).y)
    Next
    putOut s & vbLf
    droidsDump
    structuresDump
    For y = 0 To mapHeight - 1
        s = CStr(y) & ":"
        n = 0
        For x = 0 To mapWidth - 1
            t = x + y * mapWidth
            If psMapTiles(t).psObject <> 0 Then
                s = s & va(" %d=%u", x, SIMPLE_OBJECT_id(psMapTiles(t).psObject))
                n = n + 1
            End If
        Next
        s = s & " |"
        For x = 0 To mapWidth - 1
            s = s & Right$("0" & Hex$(blockTile(x, y, 0)), 2)
        Next
        putOut s & vbLf
    Next
End Sub

' the droids after mapobjects, player by player
Private Sub droidsDump()
    Dim p As Long, it As Long, d As Long, s As String, i As Long
    For p = 0 To MAX_PLAYERS - 1
        If apsDroidLists(p).size > 0 Then putOut va("droids p%d %u\n", p, apsDroidLists(p).size)
        it = apsDroidLists(p).head
        Do While it <> 0
            d = lnObj(it)
            With aDroid(d)
                s = va(" d id=%u name=%s type=%d player=%u pos=%d,%d,%d rot=%u,%u,%u body=%u orig=%u weight=%u speed=%u exp=%u kills=%u shield=%d resist=%d", _
                    .id, .aName, .droidType, .player, .pos.x, .pos.y, .pos.z, .rot.direction, .rot.pitch, .rot.roll, .body, .originalBody, _
                    .weight, .baseSpeed, .experience, .kills, .shieldPoints, .resistance)
                s = s & " bits="
                For i = 0 To DROID_MAXCOMP - 1
                    s = s & CStr(.asBits(i)) & "/"
                Next
                s = s & va(" weaps=%u", .numWeaps)
                For i = 0 To MAX_WEAPONS - 1
                    s = s & va(" w%d=%u/%u/%u", i, .asWeaps(i).nStat, .asWeaps(i).ammo, .asWeaps(i).lastFired)
                Next
                s = s & va(" blocked=%x imd=%s t=%u,%u born=%u died=%u watched=%u flags=%d sec=%x secp=%x action=%d order=%d group=%u illum=%u", _
                    .blockedBits, modelStr(.sDisplay.imd), .time_, .prevSpacetime.time_, .born, .died, .watchedTiles_size, .flags, _
                    .secondaryOrder, .secondaryOrderPending, .action, .order.typ, .group, .illumination)
                If .psGroup <> 0 Then
                    s = s & va(" grp=%d/%d/%d", aGroup(.psGroup).id, aGroup(.psGroup).typ, aGroup(.psGroup).refCount)
                Else
                    s = s & " grp=null"
                End If
                s = s & " vis="
                For i = 0 To MAX_PLAYERS - 1
                    s = s & Right$("0" & Hex$(.visible(i)), 2)
                Next
            End With
            putOut s & vbLf
            it = lnNext(it)
        Loop
    Next
End Sub

' the structures, the function lists and the flags after mapobjects; each tile's
' explored bits, and its watchers and sensors where any is set
Private Sub structuresDump()
    Dim p As Long, it As Long, st As Long, s As String, i As Long, x As Long, y As Long, t As Long, fp As Long
    For p = 0 To MAX_PLAYERS - 1
        If apsStructLists(p).size > 0 Then putOut va("structures p%d %u\n", p, apsStructLists(p).size)
        it = apsStructLists(p).head
        Do While it <> 0
            st = lnObj(it)
            With aStruct(st)
                s = va(" s id=%u stat=%s player=%u pos=%d,%d,%d rot=%u,%u,%u status=%d body=%u resist=%d cap=%u pts=%u imd=%s pre=%s fd=%s state=%d t=%u,%u,%u born=%u extra=%d flags=%d watched=%u weaps=%u", _
                    .id, asStructureStats(.pStructureType).id, .player, .pos.x, .pos.y, .pos.z, .rot.direction, .rot.pitch, .rot.roll, _
                    .status, .body, .resistance, .capacity, .currentBuildPts, modelStr(.sDisplay.imd), modelStr(.prebuiltImd), _
                    FloatBits(.foundationDepth), .state, .lastStateTime, .prevTime, .time_, .born, -CLng(.hasExtraFunction), .flags, _
                    .watchedTiles_size, .numWeaps)
                For i = 0 To MAX_WEAPONS - 1
                    s = s & va(" w%d=%u/%u/%u/%d", i, .asWeaps(i).nStat, .asWeaps(i).ammo, .asWeaps(i).lastFired, .asWeaps(i).origin)
                Next
                s = s & " vis="
                For i = 0 To MAX_PLAYERS - 1
                    s = s & Right$("0" & Hex$(.visible(i)), 2)
                Next
                If .pFunctionality <> 0 Then
                    Select Case asStructureStats(.pStructureType).typ
                    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
                        fp = .functionality.factory.psAssemblyPoint
                        s = s & va(" factory order=%x", .functionality.factory.secondaryOrder)
                        If fp <> 0 Then s = s & va(" ap=%u/%u/%d,%d,%d", aFlagPos(fp).factoryInc, aFlagPos(fp).factoryType, aFlagPos(fp).coords.x, aFlagPos(fp).coords.y, aFlagPos(fp).coords.z)
                    Case REF_REPAIR_FACILITY
                        fp = .functionality.repairFacility.psDeliveryPoint
                        If fp <> 0 Then s = s & va(" dp=%u/%u/%d,%d,%d", aFlagPos(fp).factoryInc, aFlagPos(fp).factoryType, aFlagPos(fp).coords.x, aFlagPos(fp).coords.y, aFlagPos(fp).coords.z)
                    Case REF_POWER_GEN
                        s = s & " gen="
                        For i = 0 To NUM_POWER_MODULES - 1
                            If .functionality.powerGenerator.apResExtractors(i) <> 0 Then s = s & va("%u,", aStruct(.functionality.powerGenerator.apResExtractors(i)).id) Else s = s & "0,"
                        Next
                    Case REF_RESOURCE_EXTRACTOR
                        If .functionality.resourceExtractor.psPowerGen <> 0 Then s = s & va(" pg=%u", aStruct(.functionality.resourceExtractor.psPowerGen).id) Else s = s & " pg=0"
                    Case REF_WALL, REF_GATE
                        s = s & va(" wall=%u", .functionality.wall.typ)
                    End Select
                End If
            End With
            putOut s & vbLf
            it = lnNext(it)
        Loop
        s = va("extractors p%d %u:", p, apsExtractorLists(p).size)
        it = apsExtractorLists(p).head
        Do While it <> 0
            s = s & " " & CStr(aStruct(lnObj(it)).id)
            it = lnNext(it)
        Loop
        If apsExtractorLists(p).size > 0 Then putOut s & vbLf
        s = va("flags p%d %u:", p, apsFlagPosLists(p).size)
        it = apsFlagPosLists(p).head
        Do While it <> 0
            fp = lnObj(it)
            s = s & va(" %u/%u/%d,%d,%d", aFlagPos(fp).factoryInc, aFlagPos(fp).factoryType, aFlagPos(fp).coords.x, aFlagPos(fp).coords.y, aFlagPos(fp).coords.z)
            it = lnNext(it)
        Loop
        If apsFlagPosLists(p).size > 0 Then putOut s & vbLf
    Next
    s = va("sensors %u:", apsSensorList(0).size)
    it = apsSensorList(0).head
    Do While it <> 0
        s = s & " " & CStr(SIMPLE_OBJECT_id(lnObj(it)))
        it = lnNext(it)
    Loop
    putOut s & vbLf
    For y = 0 To mapHeight - 1
        s = CStr(y) & " vis:"
        For x = 0 To mapWidth - 1
            t = x + y * mapWidth
            s = s & Right$("000" & Hex$(psMapTiles(t).tileExploredBits), 4)
        Next
        For x = 0 To mapWidth - 1
            t = x + y * mapWidth
            For p = 0 To MAX_PLAYERS - 1
                If psMapTiles(t).watchers(p) <> 0 Or psMapTiles(t).sensors(p) <> 0 Then
                    s = s & va(" %d:%d=%u/%u", x, p, psMapTiles(t).watchers(p), psMapTiles(t).sensors(p))
                End If
            Next
            If psMapTiles(t).sensorBits <> 0 Or psMapTiles(t).jammerBits <> 0 Then s = s & va(" %d:s%x/%x", x, psMapTiles(t).sensorBits, psMapTiles(t).jammerBits)
        Next
        putOut s & vbLf
        s = CStr(y) & " aux:"
        For x = 0 To mapWidth - 1
            t = x + y * mapWidth
            For p = 0 To MAX_PLAYERS - 1
                If psAuxMap(p, t) <> 0 Then s = s & va(" %d:%d=%x", x, p, CLng(psAuxMap(p, t)))
            Next
        Next
        putOut s & vbLf
    Next
End Sub

' the bits of a float, as %08X
Private Function FloatBits(ByVal f As Single) As String
    Dim b As Long
    CopyMemory b, f, 4
    FloatBits = Right$("0000000" & Hex$(b), 8)
End Function

' Park and Miller's minimal standard generator (48271), as tools\statsref's
Private Function testRand(st As Double) As Long
    st = st * 48271#
    st = st - Int(st / 2147483647#) * 2147483647#
    testRand = st
End Function

' testRand modulo n (n > 0)
Private Function testRandMod(st As Double, ByVal n As Long) As Long
    testRandMod = testRand(st) Mod n
End Function

' the ids of a grid query's objects
Private Function gridIds(list() As Long, ByVal n As Long) As String
    Dim s As String, i As Long
    For i = 0 To n - 1
        If i = 0 Then s = CStr(U32(SIMPLE_OBJECT_id(list(i)))) Else s = s & "," & CStr(U32(SIMPLE_OBJECT_id(list(i))))
    Next
    gridIds = s
End Function

Private Sub mapGrid()
    Dim st As Double, wx As Long, wy As Long, all_() As Long, nAll As Long, p As Long, it As Long
    Dim i As Long, x As Long, y As Long, r As Long, kind As Long, player As Long, s As String, psObj As Long
    Dim list() As Long, n As Long, k As Long
    st = 7 + mapWidth * 13 + mapHeight
    wx = world_coord(mapWidth)
    wy = world_coord(mapHeight)
    ' the objects in gridReset's order, to centre half the queries on
    ReDim all_(0 To 15)
    For p = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(p).head
        Do While it <> 0
            If nAll > UBound(all_) Then ReDim Preserve all_(0 To 2 * nAll - 1)
            all_(nAll) = lnObj(it)
            nAll = nAll + 1
            it = lnNext(it)
        Loop
        it = apsStructLists(p).head
        Do While it <> 0
            If nAll > UBound(all_) Then ReDim Preserve all_(0 To 2 * nAll - 1)
            all_(nAll) = lnObj(it)
            nAll = nAll + 1
            it = lnNext(it)
        Loop
    Next
    it = apsFeatureList(0).head
    Do While it <> 0
        If nAll > UBound(all_) Then ReDim Preserve all_(0 To 2 * nAll - 1)
        all_(nAll) = lnObj(it)
        nAll = nAll + 1
        it = lnNext(it)
    Loop
    gridInitialise
    gridReset
    putCaptured
    For i = 0 To 79
        If i = 40 Then
            gridReset
        End If
        If nAll > 0 And testRandMod(st, 2) = 0 Then
            psObj = all_(testRandMod(st, nAll))
            x = SIMPLE_OBJECT_pos(psObj).x + testRandMod(st, 2 * TILE_UNITS + 1) - TILE_UNITS
            y = SIMPLE_OBJECT_pos(psObj).y + testRandMod(st, 2 * TILE_UNITS + 1) - TILE_UNITS
        Else
            x = testRandMod(st, wx + 2 * TILE_UNITS) - TILE_UNITS
            y = testRandMod(st, wy + 2 * TILE_UNITS) - TILE_UNITS
        End If
        r = testRandMod(st, 12 * TILE_UNITS)
        kind = testRandMod(st, 5)
        player = testRandMod(st, MAX_PLAYERS)
        Select Case kind
        Case 0: n = gridStartIterate(x, y, r, list)
        Case 1: n = gridStartIterateArea(x, y, x + r, y + r, list)
        Case 2: n = gridStartIterateDroidsByPlayer(x, y, r, player, list)
        Case 3: n = gridStartIterateRepairCandidates(x, y, r, player, list)
        Case Else
            n = gridStartIterateUnseen(x, y, r, player, list)
            ' every other object found is seen: the player's later queries skip it
            For k = 0 To n - 1 Step 2
                Select Case list(k) \ OBJ_HANDLE_UNIT
                Case OBJ_DROID: aDroid(list(k)).seenThisTick(player) = UINT8_MAX
                Case OBJ_STRUCTURE: aStruct(list(k)).seenThisTick(player) = UINT8_MAX
                Case OBJ_FEATURE: aFeature(list(k)).seenThisTick(player) = UINT8_MAX
                End Select
            Next
        End Select
        s = gridIds(list, n)
        putCaptured
        putOut va("grid %d: %d %d,%d,%d p%d = %s\n", i, kind, x, y, r, player, s)
    Next
    gridShutDown
End Sub

Private Function powerState(ByVal p As Long) As String
    powerState = va(" p%d=%d/%lld/%lld/%lld/%d/%d/%s", p, getPower(p), getPrecisePower(p), getExtractedPower(p), _
        getWastedPower(p), getQueuedPower(p), getPowerMinusQueued(p), getApproxPowerGeneratedPerSecForDisplay(p))
End Function

Private Sub mapPower()
    Dim st As Double, p As Long, t As Long, s As String, k As Long, it As Long, psStruct As Long, amount As Long, ok As Boolean
    st = 11 + mapWidth * 3 + mapHeight * 5
    allocPlayerPower
    For p = 0 To MAX_PLAYERS - 1
        If p Mod 3 = 1 Then
            setPowerModifier p, 50 + 25 * (p Mod 5)
        End If
        setPower p, testRandMod(st, 3000)
    Next
    setPowerMaxStorage 2, 2500
    putCaptured
    For t = 1 To 40
        gameTime = t * GAME_TICKS_PER_UPDATE
        For p = 0 To MAX_PLAYERS - 1
            If t Mod 3 = 0 Then updatePlayerPower p, 3 Else updatePlayerPower p, 1
        Next
        s = va("power t%d:", t)
        For p = 0 To MAX_PLAYERS - 1
            If apsStructLists(p).size > 0 Then
                ' a request from one of the player's structures
                k = testRandMod(st, apsStructLists(p).size)
                it = apsStructLists(p).head
                Do While k > 0
                    it = lnNext(it)
                    k = k - 1
                Loop
                psStruct = lnObj(it)
                amount = testRandMod(st, 900) - 100
                If t Mod 4 = 0 Then
                    ok = requestPrecisePowerFor(psStruct, CDbl(amount) * 1000003#)
                Else
                    ok = requestPowerFor(psStruct, amount)
                End If
                s = s & va(" r%u,%d=%d/%d", aStruct(psStruct).id, amount, -CLng(ok), checkPowerRequest(psStruct))
                If t Mod 9 = 0 Then
                    delPowerRequest psStruct
                End If
            End If
        Next
        If t Mod 7 = 0 Then
            For p = 0 To MAX_PLAYERS - 1
                addPower p, testRandMod(st, 600) - 100
                usePower p, testRandMod(st, 400)
            Next
        End If
        If t = 20 Then
            ' a derrick of each player dies: its power generator lets go of it
            For p = 0 To MAX_PLAYERS - 1
                If apsExtractorLists(p).size > 0 Then
                    aStruct(lnObj(apsExtractorLists(p).head)).died = gameTime
                End If
            Next
        End If
        For p = 0 To MAX_PLAYERS - 1
            s = s & powerState(p)
        Next
        putCaptured
        putOut s & vbLf
    Next
    gameTime = 0
End Sub

' a path's points
Private Function pathPoints(m As MOVE_CONTROL) As String
    Dim i As Long, s As String
    For i = 0 To m.asPath_size - 1
        s = s & " " & CStr(m.asPath(i).x) & "," & CStr(m.asPath(i).y)
    Next
    pathPoints = s
End Function

Private Sub mapPaths()
    Dim st As Double, wx As Long, wy As Long, tick As Long, d As Long, o As Long, tX As Long, tY As Long, prop As Long, owner As Long
    Dim moveType As Long, id As Long, sX As Long, sY As Long, m As MOVE_CONTROL, emptyMove As MOVE_CONTROL, r1 As Long, r2 As Long
    Dim structs() As Long, nStructs As Long, p As Long, it As Long, n As Long, psDroid As Long, psStruct As Long, s As String, i As Long
    Dim a As Vector3i, b As Vector3i, nobounds As StructureBounds
    st = 3 + mapWidth * 17 + mapHeight * 3
    wx = world_coord(mapWidth)
    wy = world_coord(mapHeight)
    fpathInitialise
    putCaptured
    nobounds = getStructureBounds(0)
    For tick = 1 To 6
        gameTime = tick * GAME_TICKS_PER_UPDATE
        ' each destination wanted from several origins: the contexts are reused within a tick
        For d = 0 To 2
            tX = testRandMod(st, wx)
            tY = testRandMod(st, wy)
            prop = testRandMod(st, PROPULSION_TYPE_NUM)
            owner = testRandMod(st, MAX_PLAYERS)
            moveType = testRandMod(st, 3)
            For o = 0 To 3
                id = 1000 * tick + 10 * d + o
                sX = testRandMod(st, wx)
                sY = testRandMod(st, wy)
                m = emptyMove
                r1 = fpathRoute_test(m, id, sX, sY, tX, tY, prop, DROID_WEAPON, moveType, owner, o Mod 2 = 0, nobounds)
                m.Status = MOVEWAITROUTE
                r2 = fpathRoute_test(m, id, sX, sY, tX, tY, prop, DROID_WEAPON, moveType, owner, o Mod 2 = 0, nobounds)
                putCaptured
                putOut va("path %u: %d,%d->%d,%d prop%d p%d m%d = %d/%d %d dest=%d,%d n=%d:%s\n", id, sX, sY, tX, tY, prop, owner, moveType, r1, r2, _
                    m.Status, m.destination.x, m.destination.y, m.asPath_size, pathPoints(m))
            Next
        Next
    Next
    ' the droids, to points and to structures
    gameTime = 7 * GAME_TICKS_PER_UPDATE
    ReDim structs(0 To 15)
    For p = 0 To MAX_PLAYERS - 1
        it = apsStructLists(p).head
        Do While it <> 0
            If nStructs > UBound(structs) Then ReDim Preserve structs(0 To 2 * nStructs - 1)
            structs(nStructs) = lnObj(it)
            nStructs = nStructs + 1
            it = lnNext(it)
        Loop
    Next
    n = 0
    For p = 0 To MAX_PLAYERS - 1
        If n >= 30 Then Exit For
        it = apsDroidLists(p).head
        Do While it <> 0
            If n >= 30 Then n = n + 1: Exit Do
            n = n + 1
            psDroid = lnObj(it)
            If nStructs > 0 And testRandMod(st, 2) = 0 Then
                psStruct = structs(testRandMod(st, nStructs))
                tX = aStruct(psStruct).pos.x
                tY = aStruct(psStruct).pos.y
            Else
                tX = testRandMod(st, wx)
                tY = testRandMod(st, wy)
            End If
            r1 = fpathDroidRoute(psDroid, tX, tY, FMT_MOVE)
            aDroid(psDroid).sMove.Status = MOVEWAITROUTE
            r2 = fpathDroidRoute(psDroid, tX, tY, FMT_MOVE)
            putCaptured
            putOut va("droid %u: ->%d,%d = %d/%d %d dest=%d,%d n=%d:%s\n", aDroid(psDroid).id, tX, tY, r1, r2, aDroid(psDroid).sMove.Status, _
                aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y, aDroid(psDroid).sMove.asPath_size, pathPoints(aDroid(psDroid).sMove))
            it = lnNext(it)
        Loop
    Next
    s = "check:"
    For i = 0 To 39
        a = Vector3i_(testRandMod(st, wx + 2 * TILE_UNITS) - TILE_UNITS, testRandMod(st, wy + 2 * TILE_UNITS) - TILE_UNITS, 0)
        b = Vector3i_(testRandMod(st, wx), testRandMod(st, wy), 0)
        prop = testRandMod(st, PROPULSION_TYPE_NUM)
        s = s & va(" %d", -CLng(fpathCheck(a, b, prop)))
    Next
    putCaptured
    putOut s & vbLf
    fpathShutdown
    putCaptured
    gameTime = 0
End Sub

' a droid's movement state
Private Function moveDroidLine(ByVal psDroid As Long) As String
    Dim s As String, f As Long
    s = va("droid %u: p%d %s pos(%d,%d,%d) rot(%d,%d,%d) speed%d moveDir%d path%d/%d src(%d,%d) target(%d,%d) dest(%d,%d)", _
           aDroid(psDroid).id, aDroid(psDroid).player, moveDescription(aDroid(psDroid).sMove.Status), _
           aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, aDroid(psDroid).pos.z, _
           aDroid(psDroid).rot.direction, aDroid(psDroid).rot.pitch, aDroid(psDroid).rot.roll, aDroid(psDroid).sMove.speed, aDroid(psDroid).sMove.moveDir, _
           aDroid(psDroid).sMove.pathIndex, aDroid(psDroid).sMove.asPath_size, aDroid(psDroid).sMove.src.x, aDroid(psDroid).sMove.src.y, _
           aDroid(psDroid).sMove.target.x, aDroid(psDroid).sMove.target.y, aDroid(psDroid).sMove.destination.x, aDroid(psDroid).sMove.destination.y)
    s = s & va(" bump(%u,%d,%d,%d,%d,%d) shuffle%u tol%u vert%d height%d exp%u anim%d", _
               aDroid(psDroid).sMove.bumpTime, aDroid(psDroid).sMove.lastBump, aDroid(psDroid).sMove.pauseTime, aDroid(psDroid).sMove.bumpDir, _
               aDroid(psDroid).sMove.bumpPos.x, aDroid(psDroid).sMove.bumpPos.y, aDroid(psDroid).sMove.shuffleStart, _
               aDroid(psDroid).sMove.tolerance, aDroid(psDroid).sMove.iVertSpeed, aDroid(psDroid).heightAboveMap, aDroid(psDroid).experience, _
               aDroid(psDroid).animationEvent)
    f = aDroid(psDroid).sMove.psFormation
    If f <> 0 Then
        s = s & va(" frm(%d,%d,%d,%d,%u,%d,%d)", aFormation(f).x, aFormation(f).y, aFormation(f).direction, aFormation(f).refCount, _
                   aFormation(f).iSpeed, aFormation(f).size, aFormation(f).rankDist)
    End If
    moveDroidLine = s & vbLf
End Function

' the droids of every player, in list order (the count)
Private Function allDroids(v() As Long) As Long
    Dim p As Long, it As Long, n As Long
    ReDim v(0 To 15)
    For p = 0 To MAX_PLAYERS - 1
        it = apsDroidLists(p).head
        Do While it <> 0
            If n > UBound(v) Then ReDim Preserve v(0 To 2 * n - 1)
            v(n) = lnObj(it)
            n = n + 1
            it = lnNext(it)
        Loop
    Next
    allDroids = n
End Function

Private Sub mapMove()
    Dim st As Double, wx As Long, wy As Long, droids() As Long, nDroids As Long, i As Long, psDroid As Long
    Dim n As Long, tX As Long, tY As Long, rx As Long, ry As Long, ok As Boolean, how As String, tick As Long
    Dim alive() As Long, nAlive As Long, crc As Long, x As Long, y As Long, s As String, p As Long, vtolDirect As Boolean
    st = 5 + mapWidth * 13 + mapHeight * 7
    wx = world_coord(mapWidth)
    wy = world_coord(mapHeight)
    fpathInitialise
    gridInitialise
    gameTime = 1000
    deltaGameTime = GAME_TICKS_PER_UPDATE
    nDroids = allDroids(droids)
    For i = 0 To nDroids - 1
        moveReallyStopDroid droids(i)
        initDroidMovement droids(i)
    Next
    gridReset
    putCaptured
    ' the orders: three droids to a destination, near the first or anywhere
    n = 0
    For i = 0 To nDroids - 1
        If n >= 48 Then Exit For
        psDroid = droids(i)
        If n Mod 3 = 0 Then
            If testRandMod(st, 4) = 0 Then
                tX = testRandMod(st, wx)
                tY = testRandMod(st, wy)
            Else
                rx = testRandMod(st, 24 * TILE_UNITS + 1) - 12 * TILE_UNITS
                ry = testRandMod(st, 24 * TILE_UNITS + 1) - 12 * TILE_UNITS
                tX = clip_(aDroid(psDroid).pos.x + rx, TILE_UNITS, wx - TILE_UNITS - 1)
                tY = clip_(aDroid(psDroid).pos.y + ry, TILE_UNITS, wy - TILE_UNITS - 1)
            End If
        End If
        ok = True
        vtolDirect = False
        If DROID_isVtol(psDroid) Then vtolDirect = testRandMod(st, 2) = 0
        If vtolDirect Then
            moveDroidToDirect psDroid, tX, tY
            how = "direct"
        ElseIf n Mod 5 = 4 Then
            ok = moveDroidToNoFormation(psDroid, tX, tY)
            how = "noformation"
        Else
            ok = moveDroidTo(psDroid, tX, tY)
            how = "to"
        End If
        putCaptured
        putOut va("order %u: %s ->%d,%d = %d %s\n", aDroid(psDroid).id, how, tX, tY, -CLng(ok), moveDescription(aDroid(psDroid).sMove.Status))
        n = n + 1
    Next
    For tick = 1 To 80
        gameTime = gameTime + GAME_TICKS_PER_UPDATE
        gridReset
        fpathUpdate
        If tick = 30 Or tick = 50 Then
            ' stops, turns and new destinations on the way
            For i = 0 To nDroids - 1
                If i >= 48 Then Exit For
                psDroid = droids(i)
                If aDroid(psDroid).died <> 0 Then GoTo next_i
                If i Mod 7 = 3 Then
                    moveStopDroid psDroid
                    putOut va("stop %u: %s\n", aDroid(psDroid).id, moveDescription(aDroid(psDroid).sMove.Status))
                ElseIf i Mod 11 = 5 Then
                    x = testRandMod(st, wx)
                    y = testRandMod(st, wy)
                    moveTurnDroid psDroid, x, y
                    putOut va("turn %u: ->%d,%d %s\n", aDroid(psDroid).id, x, y, moveDescription(aDroid(psDroid).sMove.Status))
                ElseIf i Mod 13 = 6 Then
                    moveReallyStopDroid psDroid
                    putOut va("reallystop %u\n", aDroid(psDroid).id)
                ElseIf i Mod 4 = 1 Then
                    x = testRandMod(st, wx)
                    y = testRandMod(st, wy)
                    ok = moveDroidTo(psDroid, x, y)
                    putCaptured
                    putOut va("reorder %u: ->%d,%d = %d %s\n", aDroid(psDroid).id, x, y, -CLng(ok), moveDescription(aDroid(psDroid).sMove.Status))
                End If
next_i:
            Next
        End If
        nAlive = allDroids(alive)
        For i = 0 To nAlive - 1
            psDroid = alive(i)
            If aDroid(psDroid).died = 0 Then
                syncDebugDroid_ "mapMove", psDroid, Asc("M")
                moveUpdateDroid psDroid
                syncDebugDroid_ "mapMove", psDroid, Asc(">")
            End If
        Next
        putCaptured
        If tick = 1 Or tick = 80 Or syncDumpAll Then
            putOut dumpCurrentDebugSyncLog()
        End If
        crc = nextDebugSync(gameTime)
        putCaptured
        putOut va("tick %d: crc=%04X\n", tick, crc)
        If tick Mod 20 = 0 Then
            For i = 0 To nDroids - 1
                If i >= 48 Then Exit For
                If aDroid(droids(i)).died = 0 Then
                    putOut moveDroidLine(droids(i))
                End If
            Next
        End If
    Next
    putOut va("features %d\n", apsFeatureList(0).size)
    s = "power:"
    For p = 0 To MAX_PLAYERS - 1
        s = s & va(" %d", getPower(p))
    Next
    putOut s & vbLf
    gridShutDown
    fpathShutdown
    putCaptured
    gameTime = 0
    deltaGameTime = 0
End Sub

' a droid's order and action state
Private Function orderDroidLine(ByVal psDroid As Long) As String
    Dim t As Long, o As Long, base As Long
    t = aDroid(psDroid).psActionTarget(0)
    If t <> 0 Then t = SIMPLE_OBJECT_id(t)
    o = aDroid(psDroid).order.psObj
    If o <> 0 Then o = SIMPLE_OBJECT_id(o)
    base = aDroid(psDroid).psBaseStruct
    If base <> 0 Then base = aStruct(base).id
    orderDroidLine = va("droid %u: p%d t%d %s pos(%d,%d,%d) body%u/%u order%d(%d,%d)(%d,%d)dir%d obj%u stats%d rtr%d list%d/%d", _
                        aDroid(psDroid).id, aDroid(psDroid).player, aDroid(psDroid).droidType, moveDescription(aDroid(psDroid).sMove.Status), _
                        aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, aDroid(psDroid).pos.z, aDroid(psDroid).body, aDroid(psDroid).originalBody, _
                        aDroid(psDroid).order.typ, aDroid(psDroid).order.pos.x, aDroid(psDroid).order.pos.y, aDroid(psDroid).order.pos2.x, aDroid(psDroid).order.pos2.y, _
                        aDroid(psDroid).order.direction, o, aDroid(psDroid).order.psStats, aDroid(psDroid).order.rtrType, aDroid(psDroid).listSize, aDroid(psDroid).asOrderList_size) & _
                     va(" action%d(%d,%d) tgt%u start%u pts%u sec%X base%u grp%d/%d wpn(%d,%d) ur%d", _
                        aDroid(psDroid).action, aDroid(psDroid).actionPos.x, aDroid(psDroid).actionPos.y, t, aDroid(psDroid).actionStarted, aDroid(psDroid).actionPoints, _
                        aDroid(psDroid).secondaryOrder, base, aDroid(psDroid).group, aDroid(psDroid).repairGroup, _
                        aDroid(psDroid).asWeaps(0).rot.direction, aDroid(psDroid).asWeaps(0).rot.pitch, aDroid(psDroid).underRepair) & vbLf
End Function

' a structure's building state
Private Function orderStructLine(ByVal s As Long) As String
    orderStructLine = va("struct %u: p%d %s status%d body%u/%u pts%d rate%d cap%d res%d pos(%d,%d,%d)", aStruct(s).id, aStruct(s).player, _
                         asStructureStats(aStruct(s).pStructureType).id, aStruct(s).status, aStruct(s).body, STRUCTURE_structureBody(s), _
                         aStruct(s).currentBuildPts, aStruct(s).buildRate, aStruct(s).capacity, aStruct(s).resistance, _
                         aStruct(s).pos.x, aStruct(s).pos.y, aStruct(s).pos.z) & vbLf
End Function

Private Function randomStructure(st As Double, ByVal player As Long) As Long
    Dim n As Long, k As Long, it As Long
    n = apsStructLists(player).size
    If n = 0 Then
        randomStructure = 0
        Exit Function
    End If
    k = testRandMod(st, n)
    it = apsStructLists(player).head
    Do While it <> 0
        If k = 0 Then
            randomStructure = lnObj(it)
            Exit Function
        End If
        k = k - 1
        it = lnNext(it)
    Loop
    randomStructure = 0
End Function

Private Function randomDroid(st As Double, ByVal player As Long) As Long
    Dim n As Long, k As Long, it As Long
    n = apsDroidLists(player).size
    If n = 0 Then
        randomDroid = 0
        Exit Function
    End If
    k = testRandMod(st, n)
    it = apsDroidLists(player).head
    Do While it <> 0
        If k = 0 Then
            randomDroid = lnObj(it)
            Exit Function
        End If
        k = k - 1
        it = lnNext(it)
    Loop
    randomDroid = 0
End Function

' a droid or structure of a player not allied to player
Private Function randomEnemy(st As Double, ByVal player As Long) As Long
    Dim q As Long, i As Long
    q = testRandMod(st, MAX_PLAYERS)
    For i = 0 To MAX_PLAYERS - 1
        If q = player Then GoTo next_i
        If aiCheckAlliances(q, player) Then GoTo next_i
        If apsDroidLists(q).size = 0 And apsStructLists(q).size = 0 Then GoTo next_i
        If apsStructLists(q).size = 0 Then
            randomEnemy = randomDroid(st, q)
            Exit Function
        End If
        If apsDroidLists(q).size <> 0 Then
            If testRandMod(st, 2) = 0 Then
                randomEnemy = randomDroid(st, q)
                Exit Function
            End If
        End If
        randomEnemy = randomStructure(st, q)
        Exit Function
next_i:
        q = (q + 1) Mod MAX_PLAYERS
    Next
    randomEnemy = 0
End Function

' a point up to tiles away from pos, on the map
Private Function nearPoint(st As Double, pos As Vector3i, ByVal tiles As Long, ByVal wx As Long, ByVal wy As Long) As Vector2i
    Dim rx As Long, ry As Long
    rx = testRandMod(st, 2 * tiles * TILE_UNITS + 1) - tiles * TILE_UNITS
    ry = testRandMod(st, 2 * tiles * TILE_UNITS + 1) - tiles * TILE_UNITS
    nearPoint = Vector2i_(clip_(pos.x + rx, TILE_UNITS, wx - TILE_UNITS - 1), clip_(pos.y + ry, TILE_UNITS, wy - TILE_UNITS - 1))
End Function

' an order for a droid, as a player could give it (what it was); with ModeQueue, through the game messages
Private Function giveOrder(st As Double, ByVal psDroid As Long, avail() As Long, ByVal nAvail As Long, ByVal wallStat As Long, ByVal wx As Long, ByVal wy As Long, Optional ByVal mode As QUEUE_MODE = ModeImmediate) As String
    Dim r As Long, p As Long, t As Vector2i, e As Long, s As Long, x As Long, y As Long, dir_ As Long, d As Long
    Dim length As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, r2 As Long, o As Long, ord As DROID_ORDER_DATA
    Dim t2 As Vector2i, t3 As Vector2i, q1 As DROID_ORDER_DATA, q2 As DROID_ORDER_DATA
    r = testRandMod(st, 100)
    p = aDroid(psDroid).player
    If DROID_isVtol(psDroid) Then
        t = nearPoint(st, aDroid(psDroid).pos, 10, wx, wy)
        If r < 30 Then
            orderDroidLoc psDroid, DORDER_CIRCLE, t.x, t.y, mode
            giveOrder = va("circle %d,%d", t.x, t.y)
            Exit Function
        End If
        If r < 60 Then
            orderDroidLoc psDroid, DORDER_PATROL, t.x, t.y, mode
            giveOrder = va("patrol %d,%d", t.x, t.y)
            Exit Function
        End If
        If r < 85 Then
            e = randomEnemy(st, p)
            If e <> 0 Then
                orderDroidObj psDroid, DORDER_ATTACK, e, mode
                giveOrder = va("attack %u", SIMPLE_OBJECT_id(e))
                Exit Function
            End If
        End If
        orderDroid psDroid, DORDER_RTB, mode
        giveOrder = "rtb"
        Exit Function
    End If
    Select Case aDroid(psDroid).droidType
    Case DROID_CONSTRUCT, DROID_CYBORG_CONSTRUCT
        If r < 40 And nAvail > 0 Then
            s = avail(testRandMod(st, nAvail))
            t = nearPoint(st, aDroid(psDroid).pos, 3, wx, wy)
            x = world_coord(map_coord(t.x)) + TILE_UNITS \ 2
            y = world_coord(map_coord(t.y)) + TILE_UNITS \ 2
            dir_ = (testRandMod(st, 4) * DEG(90)) And &HFFFF&
            orderDroidStatsLocDir psDroid, DORDER_BUILD, s, x, y, dir_, mode
            giveOrder = va("build %s %d,%d %d", asStructureStats(s).id, x, y, dir_)
            Exit Function
        End If
        If r < 50 And wallStat >= 0 Then
            t = nearPoint(st, aDroid(psDroid).pos, 3, wx, wy)
            length = testRandMod(st, 5) + 1
            x1 = world_coord(map_coord(t.x)) + TILE_UNITS \ 2
            y1 = world_coord(map_coord(t.y)) + TILE_UNITS \ 2
            x2 = x1
            y2 = y1
            If testRandMod(st, 2) = 0 Then
                x2 = clip_(x1 + length * TILE_UNITS, TILE_UNITS, wx - TILE_UNITS - 1)
            Else
                y2 = clip_(y1 + length * TILE_UNITS, TILE_UNITS, wy - TILE_UNITS - 1)
            End If
            orderDroidStatsTwoLocDir psDroid, DORDER_LINEBUILD, wallStat, x1, y1, x2, y2, 0, mode
            giveOrder = va("linebuild %d,%d %d,%d", x1, y1, x2, y2)
            Exit Function
        End If
        If r < 60 Then
            s = randomStructure(st, p)
            If s <> 0 Then
                orderDroidObj psDroid, DORDER_DEMOLISH, s, mode
                giveOrder = va("demolish %u", aStruct(s).id)
                Exit Function
            End If
        ElseIf r < 75 Then
            s = randomStructure(st, p)
            If s <> 0 Then
                orderDroidObj psDroid, DORDER_REPAIR, s, mode
                giveOrder = va("repair %u", aStruct(s).id)
                Exit Function
            End If
        End If
    Case DROID_REPAIR, DROID_CYBORG_REPAIR
        If r < 50 Then
            d = randomDroid(st, p)
            If d <> 0 And d <> psDroid Then
                orderDroidObj psDroid, DORDER_DROIDREPAIR, d, mode
                giveOrder = va("droidrepair %u", aDroid(d).id)
                Exit Function
            End If
        End If
    Case DROID_SENSOR
        If r < 50 Then
            e = randomEnemy(st, p)
            If e <> 0 Then
                orderDroidObj psDroid, DORDER_OBSERVE, e, mode
                giveOrder = va("observe %u", SIMPLE_OBJECT_id(e))
                Exit Function
            End If
        End If
    Case Else
        If r < 25 Then
            e = randomEnemy(st, p)
            If e <> 0 Then
                orderDroidObj psDroid, DORDER_ATTACK, e, mode
                giveOrder = va("attack %u", SIMPLE_OBJECT_id(e))
                Exit Function
            End If
        ElseIf r < 35 Then
            s = randomStructure(st, p)
            If s <> 0 Then
                orderDroidObj psDroid, DORDER_GUARD, s, mode
                giveOrder = va("guard %u", aStruct(s).id)
                Exit Function
            End If
        End If
    End Select
    ' the orders every droid takes
    r2 = testRandMod(st, 100)
    t = nearPoint(st, aDroid(psDroid).pos, 10, wx, wy)
    If r2 < 30 Then
        orderDroidLoc psDroid, DORDER_MOVE, t.x, t.y, mode
        giveOrder = va("move %d,%d", t.x, t.y)
    ElseIf r2 < 45 Then
        orderDroidLoc psDroid, DORDER_SCOUT, t.x, t.y, mode
        giveOrder = va("scout %d,%d", t.x, t.y)
    ElseIf r2 < 55 Then
        orderDroidLoc psDroid, DORDER_PATROL, t.x, t.y, mode
        giveOrder = va("patrol %d,%d", t.x, t.y)
    ElseIf r2 < 62 Then
        orderDroid psDroid, DORDER_HOLD, mode
        giveOrder = "hold"
    ElseIf r2 < 70 Then
        orderDroid psDroid, DORDER_RTB, mode
        giveOrder = "rtb"
    ElseIf r2 < 78 Then
        orderDroid psDroid, DORDER_RTR, mode
        giveOrder = "rtr"
    ElseIf r2 < 83 Then
        orderDroid psDroid, DORDER_STOP, mode
        giveOrder = "stop"
    ElseIf r2 < 87 Then
        orderDroid psDroid, DORDER_RECYCLE, mode
        giveOrder = "recycle"
    ElseIf r2 < 95 Then
        ' a click on an object (chooseOrderObj, as orderSelectedObjAdd gives it)
        o = 0
        If testRandMod(st, 2) = 0 Then
            o = randomEnemy(st, p)
        Else
            o = randomStructure(st, p)
        End If
        If o = 0 Then
            giveOrder = "click none"
            Exit Function
        End If
        ord = chooseOrderObj(psDroid, o, False)
        If ord.typ <> DORDER_NONE Then
            If validOrderForObj(ord.typ) Then
                orderDroidObj psDroid, ord.typ, ord.psObj, mode
            End If
        End If
        giveOrder = va("click %u: %s", SIMPLE_OBJECT_id(o), getDroidOrderName(ord.typ))
    Else
        ' a move, then two queued
        orderDroidLoc psDroid, DORDER_MOVE, t.x, t.y, mode
        t2 = nearPoint(st, aDroid(psDroid).pos, 10, wx, wy)
        q1 = DroidOrder_pos(DORDER_SCOUT, t2)
        If mode = ModeQueue Then sendDroidInfo psDroid, q1, True Else orderDroidAdd psDroid, q1
        t3 = nearPoint(st, aDroid(psDroid).pos, 10, wx, wy)
        q2 = DroidOrder_pos(DORDER_MOVE, t3)
        If mode = ModeQueue Then sendDroidInfo psDroid, q2, True Else orderDroidAdd psDroid, q2
        giveOrder = va("queue %d,%d %d,%d %d,%d", t.x, t.y, t2.x, t2.y, t3.x, t3.y)
    End If
End Function

Private Sub mapOrders()
    Dim st As Double, wx As Long, wy As Long, buildable As Variant, avail() As Long, nAvail As Long, i As Long, p As Long
    Dim wallStat As Long, droids() As Long, nDroids As Long, k As Long, it As Long, s As Long, psDroid As Long, how As String
    Dim tick As Long, nxt As Long, crc As Long, line_ As String, c As Boolean
    st = 11 + mapWidth * 17 + mapHeight * 3
    wx = world_coord(mapWidth)
    wy = world_coord(mapHeight)
    fpathInitialise
    gridInitialise
    gameTime = 1000
    deltaGameTime = GAME_TICKS_PER_UPDATE
    allocPlayerPower
    ' the structures the trucks may build, buildable by every player
    buildable = Array("A0LightFactory", "A0PowerGenerator", "A0ResearchFacility", "A0HardcreteMk1Wall", "WallTower01", _
                      "A0RepairCentre3", "Sys-SensoTower02", "A0CommandCentre", "A0ResourceExtractor", "A0FacMod1", _
                      "A0PowMod1", "A0ResearchModule1", "PillBox1", "A0CyborgFactory", "A0VTolFactory1", "A0VtolPad")
    ReDim avail(0 To UBound(buildable))
    For k = 0 To UBound(buildable)
        i = getStructStatFromName(buildable(k))
        If i >= 0 Then
            avail(nAvail) = i
            nAvail = nAvail + 1
            For p = 0 To MAX_PLAYERS - 1
                apStructTypeLists(p, i) = AVAILABLE
            Next
        End If
    Next
    wallStat = getStructStatFromName("A0HardcreteMk1Wall")
    For p = 0 To MAX_PLAYERS - 1
        setPower p, 1000 + 500 * (p Mod 3)
    Next
    ' damage: a third of the droids and a quarter of the built structures
    nDroids = allDroids(droids)
    For i = 0 To nDroids - 1
        ' DROID's constructor leaves actionStarted and actionPoints uninitialised: 0, as statsref sets them
        aDroid(droids(i)).actionStarted = 0
        aDroid(droids(i)).actionPoints = 0
        If i Mod 3 = 1 Then
            aDroid(droids(i)).body = MAX_(1, aDroid(droids(i)).body * 2 \ 5)
        End If
    Next
    k = 0
    For p = 0 To MAX_PLAYERS - 1
        it = apsStructLists(p).head
        Do While it <> 0
            s = lnObj(it)
            If aStruct(s).status = SS_BUILT Then
                c = k Mod 4 = 2
                k = k + 1
                If c Then aStruct(s).body = MAX_(1, aStruct(s).body \ 2)
            End If
            it = lnNext(it)
        Loop
    Next
    ' secondary orders
    For i = 0 To nDroids - 1
        If i >= 64 Then Exit For
        psDroid = droids(i)
        If i Mod 7 = 3 Then
            If i Mod 2 <> 0 Then
                secondarySetState psDroid, DSO_HALTTYPE, DSS_HALT_HOLD, ModeImmediate
            Else
                secondarySetState psDroid, DSO_HALTTYPE, DSS_HALT_PURSUE, ModeImmediate
            End If
        ElseIf i Mod 7 = 5 Then
            secondarySetState psDroid, DSO_ATTACK_RANGE, DSS_ARANGE_SHORT, ModeImmediate
        ElseIf i Mod 11 = 7 Then
            secondarySetState psDroid, DSO_REPAIR_LEVEL, DSS_REPLEV_HIGH, ModeImmediate
        Else
            GoTo next_secondary
        End If
        putCaptured
        putOut va("secondary %u: %X %d\n", aDroid(psDroid).id, aDroid(psDroid).secondaryOrder, aDroid(psDroid).order.typ)
next_secondary:
    Next
    gridReset
    putCaptured
    ' the orders
    For i = 0 To nDroids - 1
        If i >= 64 Then Exit For
        psDroid = droids(i)
        If aDroid(psDroid).died = 0 Then
            how = giveOrder(st, psDroid, avail, nAvail, wallStat, wx, wy)
            putCaptured
            putOut va("order %u: %s\n", aDroid(psDroid).id, how)
            putOut orderDroidLine(psDroid)
        End If
    Next
    For tick = 1 To 90
        gameTime = gameTime + GAME_TICKS_PER_UPDATE
        gridReset
        fpathUpdate
        If tick = 20 Or tick = 40 Then
            ' new orders on the way
            For i = 0 To nDroids - 1
                If i >= 64 Then Exit For
                psDroid = droids(i)
                If aDroid(psDroid).died = 0 And i Mod 3 = tick \ 20 - 1 Then
                    how = giveOrder(st, psDroid, avail, nAvail, wallStat, wx, wy)
                    putCaptured
                    putOut va("reorder %u: %s\n", aDroid(psDroid).id, how)
                End If
            Next
        End If
        For p = 0 To MAX_PLAYERS - 1
            ' mutating_list_iterate
            it = apsDroidLists(p).head
            Do While it <> 0
                nxt = lnNext(it)
                droidUpdate lnObj(it)
                it = nxt
            Loop
        Next
        For p = 0 To MAX_PLAYERS - 1
            updatePlayerPower p
        Next
        putCaptured
        If tick = 1 Or tick = 90 Or syncDumpAll Then
            putOut dumpCurrentDebugSyncLog()
        End If
        crc = nextDebugSync(gameTime)
        putCaptured
        putOut va("tick %d: crc=%04X\n", tick, crc)
        If tick Mod 30 = 0 Then
            For i = 0 To nDroids - 1
                If i >= 64 Then Exit For
                If aDroid(droids(i)).died = 0 Then
                    putOut orderDroidLine(droids(i))
                End If
            Next
        End If
    Next
    line_ = "count:"
    For p = 0 To MAX_PLAYERS - 1
        line_ = line_ & va(" %d/%d", apsDroidLists(p).size, apsStructLists(p).size)
        it = apsStructLists(p).head
        Do While it <> 0
            s = lnObj(it)
            If aStruct(s).status <> SS_BUILT Or aStruct(s).body <> STRUCTURE_structureBody(s) Then
                putOut orderStructLine(s)
            End If
            it = lnNext(it)
        Loop
    Next
    putOut line_ & vbLf
    putOut va("features %d\n", apsFeatureList(0).size)
    line_ = "power:"
    For p = 0 To MAX_PLAYERS - 1
        line_ = line_ & va(" %d", getPower(p))
    Next
    putOut line_ & vbLf
    gridShutDown
    fpathShutdown
    putCaptured
    gameTime = 0
    deltaGameTime = 0
End Sub

' a droid's combat state
Private Function combatDroidLine(ByVal psDroid As Long) As String
    Dim t As Long, o As Long
    t = aDroid(psDroid).psActionTarget(0)
    If t <> 0 Then t = SIMPLE_OBJECT_id(t)
    o = aDroid(psDroid).order.psObj
    If o <> 0 Then o = SIMPLE_OBJECT_id(o)
    With aDroid(psDroid)
        combatDroidLine = va("droid %u: p%d t%d %s pos(%d,%d,%d) body%u/%u order%d obj%u action%d tgt%u exp%u kills%u", _
                             .id, .player, .droidType, moveDescription(.sMove.Status), .pos.x, .pos.y, .pos.z, .body, .originalBody, _
                             .order.typ, o, .action, t, .experience, .kills) & _
                          va(" w(%u,%u,%u,%u,%d,%d) hit(%u,%d) vis%d,%d,%d,%d", _
                             .asWeaps(0).nStat, .asWeaps(0).ammo, .asWeaps(0).lastFired, .asWeaps(0).shotsFired, .asWeaps(0).rot.direction, _
                             .asWeaps(0).rot.pitch, .timeLastHit, .lastHitWeapon, .visible(0), .visible(1), .visible(2), .visible(3)) & vbLf
    End With
End Function

' a structure of a player not allied to player
Private Function randomEnemyStructure(st As Double, ByVal player As Long) As Long
    Dim q As Long, i As Long
    q = testRandMod(st, MAX_PLAYERS)
    For i = 0 To MAX_PLAYERS - 1
        If q <> player Then
            If Not aiCheckAlliances(q, player) Then
                If apsStructLists(q).size <> 0 Then
                    randomEnemyStructure = randomStructure(st, q)
                    Exit Function
                End If
            End If
        End If
        q = (q + 1) Mod MAX_PLAYERS
    Next
    randomEnemyStructure = 0
End Function

Private Sub mapCombat()
    Dim st As Double, wx As Long, wy As Long, droids() As Long, nDroids As Long, fighters() As Long, nFighters As Long
    Dim i As Long, k As Long, psDroid As Long, psTarget As Long, prop As Long, x As Long, y As Long, found_ As Boolean
    Dim tries As Long, ang As Long, r As Long, how As String, tick As Long, p As Long, it As Long, nxt As Long
    Dim nProj As Long, pr As Long, crc As Long, s As String, isFighter As Boolean
    st = 13 + mapWidth * 19 + mapHeight * 11
    wx = world_coord(mapWidth)
    wy = world_coord(mapHeight)
    fpathInitialise
    gridInitialise
    gameTime = 1000
    deltaGameTime = GAME_TICKS_PER_UPDATE
    resetDamageModifiers
    proj_InitSystem
    ' up to 40 armed droids moved near an enemy structure, and ordered
    nDroids = allDroids(droids)
    ReDim fighters(0 To 39)
    For i = 0 To nDroids - 1
        psDroid = droids(i)
        If nFighters < 40 And aDroid(psDroid).numWeaps > 0 Then
            If aDroid(psDroid).asWeaps(0).nStat > 0 And Not DROID_isVtol(psDroid) And Not DROID_isTransporter(psDroid) Then
                fighters(nFighters) = psDroid
                nFighters = nFighters + 1
            End If
        End If
    Next
    For i = 0 To nFighters - 1
        psDroid = fighters(i)
        psTarget = randomEnemyStructure(st, aDroid(psDroid).player)
        If psTarget = 0 Then GoTo next_fighter
        prop = asPropulsionStats(DROID_getPropulsionStats(psDroid)).propulsionType
        x = 0
        y = 0
        found_ = False
        For tries = 0 To 19
            If found_ Then Exit For
            ang = testRandMod(st, 65536)
            r = (3 + testRandMod(st, 4)) * TILE_UNITS
            x = aStruct(psTarget).pos.x + iSinR(ang, r)
            y = aStruct(psTarget).pos.y + iCosR(ang, r)
            found_ = x >= TILE_UNITS And y >= TILE_UNITS And x < wx - TILE_UNITS And y < wy - TILE_UNITS
            If found_ Then found_ = Not fpathBlockingTile(map_coord(x), map_coord(y), prop)
            If found_ Then found_ = Not TileIsOccupied(mapTile(map_coord(x), map_coord(y)))
        Next
        If Not found_ Then
            putOut va("fighter %u: no place near %u\n", aDroid(psDroid).id, aStruct(psTarget).id)
            GoTo next_fighter
        End If
        droidSetPosition psDroid, x, y
        Select Case testRandMod(st, 4)
        Case 0
            orderDroidObj psDroid, DORDER_ATTACK, psTarget, ModeImmediate
            how = "attack"
        Case 1
            orderDroidLoc psDroid, DORDER_SCOUT, aStruct(psTarget).pos.x, aStruct(psTarget).pos.y, ModeImmediate
            how = "scout"
        Case 2
            orderDroid psDroid, DORDER_HOLD, ModeImmediate
            how = "hold"
        Case Else
            orderDroid psDroid, DORDER_STOP, ModeImmediate
            how = "stop"
        End Select
        putCaptured
        putOut va("fighter %u: ->%d,%d near %u %s\n", aDroid(psDroid).id, x, y, aStruct(psTarget).id, how)
next_fighter:
    Next
    gridReset
    putCaptured
    For tick = 1 To 60
        gameTime = gameTime + GAME_TICKS_PER_UPDATE
        visUpdateLevel
        gridReset
        processVisibility
        fpathUpdate
        cmdDroidUpdate
        For p = 0 To MAX_PLAYERS - 1
            updatePlayerPower p
            ' mutating_list_iterate
            it = apsDroidLists(p).head
            Do While it <> 0
                nxt = lnNext(it)
                droidUpdate lnObj(it)
                it = nxt
            Loop
        Next
        proj_UpdateAll
        objmemUpdate
        putCaptured
        If tick = 1 Or tick = 60 Or syncDumpAll Then
            putOut dumpCurrentDebugSyncLog()
        End If
        nProj = 0
        pr = proj_GetFirst()
        Do While pr <> 0
            nProj = nProj + 1
            pr = proj_GetNext()
        Loop
        crc = nextDebugSync(gameTime)
        putCaptured
        putOut va("tick %d: crc=%04X proj=%d\n", tick, crc, nProj)
        If tick Mod 20 = 0 Then
            nDroids = allDroids(droids)
            For i = 0 To nDroids - 1
                isFighter = False
                For k = 0 To nFighters - 1
                    If fighters(k) = droids(i) Then isFighter = True: Exit For
                Next
                If isFighter Then putOut combatDroidLine(droids(i))
            Next
        End If
    Next
    s = "count:"
    For p = 0 To MAX_PLAYERS - 1
        s = s & va(" %d/%d", apsDroidLists(p).size, apsStructLists(p).size)
        it = apsStructLists(p).head
        Do While it <> 0
            If aStruct(lnObj(it)).body <> STRUCTURE_structureBody(lnObj(it)) Then
                putOut orderStructLine(lnObj(it))
            End If
            it = lnNext(it)
        Loop
    Next
    putOut s & vbLf
    putOut va("features %d\n", apsFeatureList(0).size)
    proj_FreeAllProjectiles
    gridShutDown
    fpathShutdown
    putCaptured
    gameTime = 0
    deltaGameTime = 0
End Sub

' (mapbase) save (or restore) the research state and the stats' upgrades
Private Sub baseSnapshot(ByVal save As Boolean)
    Dim p As Long, i As Long
    If save Then
        savedResList = asPlayerResList
        savedCompLists = apCompLists
        savedStructTypeLists = apStructTypeLists
        For p = 0 To MAX_PLAYERS - 1
            savedDefaultSensor(p) = aDefaultSensor(p)
            savedDefaultECM(p) = aDefaultECM(p)
            savedDefaultRepair(p) = aDefaultRepair(p)
        Next
        savedWeaponStats = asWeaponStats
        savedBodyStats = asBodyStats
        savedBrainStats = asBrainStats
        savedPropulsionStats = asPropulsionStats
        savedSensorStats = asSensorStats
        savedECMStats = asECMStats
        savedRepairStats = asRepairStats
        savedConstructStats = asConstructStats
        savedStructureStats = asStructureStats
        research_testState True
    Else
        asPlayerResList = savedResList
        apCompLists = savedCompLists
        apStructTypeLists = savedStructTypeLists
        For p = 0 To MAX_PLAYERS - 1
            aDefaultSensor(p) = savedDefaultSensor(p)
            aDefaultECM(p) = savedDefaultECM(p)
            aDefaultRepair(p) = savedDefaultRepair(p)
        Next
        For p = 0 To MAX_PLAYERS - 1
            For i = 0 To asWeaponStats_size - 1
                asWeaponStats(i).upgrade(p) = savedWeaponStats(i).upgrade(p)
            Next
            For i = 0 To asBodyStats_size - 1
                asBodyStats(i).upgrade(p) = savedBodyStats(i).upgrade(p)
            Next
            For i = 0 To asBrainStats_size - 1
                asBrainStats(i).upgrade(p) = savedBrainStats(i).upgrade(p)
            Next
            For i = 0 To asPropulsionStats_size - 1
                asPropulsionStats(i).upgrade(p) = savedPropulsionStats(i).upgrade(p)
            Next
            For i = 0 To asSensorStats_size - 1
                asSensorStats(i).upgrade(p) = savedSensorStats(i).upgrade(p)
            Next
            For i = 0 To asECMStats_size - 1
                asECMStats(i).upgrade(p) = savedECMStats(i).upgrade(p)
            Next
            For i = 0 To asRepairStats_size - 1
                asRepairStats(i).upgrade(p) = savedRepairStats(i).upgrade(p)
            Next
            For i = 0 To asConstructStats_size - 1
                asConstructStats(i).upgrade(p) = savedConstructStats(i).upgrade(p)
            Next
            For i = 0 To numStructureStats - 1
                asStructureStats(i).upgrade(p) = savedStructureStats(i).upgrade(p)
            Next
        Next
        research_testState False
    End If
End Sub

' structure.h's popStatusPending, for a RESEARCH_FACILITY
Private Sub popStatusPending_research(f As RESEARCH_FACILITY)
    If f.pendingCount = 0 Then
        f.pendingCount = f.pendingCount + 1
    End If
    f.pendingCount = f.pendingCount - 1
    If f.pendingCount = 0 Then
        ' Subject is now synchronised, remove pending.
        f.psSubjectPending = -1
        f.statusPending = FACTORY_NOTHING_PENDING
    End If
End Sub

' a topic the player has neither started nor completed, from a random start; -1 if none
Private Function chooseResearch(st As Double, ByVal player As Long) As Long
    Dim n As Long, start As Long, k As Long, i As Long
    chooseResearch = -1
    n = asResearch_size
    If n = 0 Then Exit Function
    start = testRandMod(st, n)
    For k = 0 To n - 1
        i = (start + k) Mod n
        If Not IsResearchStarted(asPlayerResList(player, i)) Then
            If Not IsResearchCompleted(asPlayerResList(player, i)) Then
                chooseResearch = i
                Exit Function
            End If
        End If
    Next
End Function

' recvResearchStatus's start of a topic (the GAME_RESEARCHSTATUS message, bStart)
Private Sub startResearch(ByVal psBuilding As Long, ByVal index As Long)
    Dim player As Long, psOtherBuilding As Long
    player = aStruct(psBuilding).player
    ResetPendingResearchStatus asPlayerResList(player, index)
    popStatusPending_research aStruct(psBuilding).functionality.researchFacility
    If aStruct(psBuilding).functionality.researchFacility.psSubject <> -1 Then
        cancelResearch psBuilding, ModeImmediate
    End If
    If IsResearchStarted(asPlayerResList(player, index)) Then
        psOtherBuilding = findResearchingFacilityByResearchIndex(player, index)
        If psOtherBuilding <> 0 Then
            cancelResearch psOtherBuilding, ModeImmediate
        End If
    End If
    aStruct(psBuilding).functionality.researchFacility.psSubject = index
    MakeResearchStarted asPlayerResList(player, index)
    aStruct(psBuilding).functionality.researchFacility.timeStartHold = 0
End Sub

' a template of the player's for the factory, from a random start; 0 if none
Private Function chooseTemplate(st As Double, ByVal psFactory As Long) As Long
    Dim n As Long, start As Long, k As Long, psTempl As Long, player As Long
    chooseTemplate = 0
    player = aStruct(psFactory).player
    n = templateCount(player)
    If n = 0 Then Exit Function
    start = testRandMod(st, n)
    For k = 0 To n - 1
        psTempl = droidTemplates_at(player, (start + k) Mod n)
        If validTemplateForFactory(psTempl, psFactory, False) Then
            chooseTemplate = psTempl
            Exit Function
        End If
    Next
End Function

' work for the idle labs and factories
Private Sub baseGiveWork(st As Double)
    Dim p As Long, it As Long, s As Long, index As Long, psTempl As Long, ok As Boolean
    For p = 0 To MAX_PLAYERS - 1
        it = apsStructLists(p).head
        Do While it <> 0
            s = lnObj(it)
            If aStruct(s).status = SS_BUILT And aStruct(s).pFunctionality <> 0 Then
                If asStructureStats(aStruct(s).pStructureType).typ = REF_RESEARCH Then
                    If aStruct(s).functionality.researchFacility.psSubject = -1 Then
                        index = chooseResearch(st, p)
                        If index >= 0 Then
                            ok = researchAvailable(index, p, ModeImmediate)
                            startResearch s, index
                            putCaptured
                            putOut va("lab %u: research %d avail%d\n", aStruct(s).id, index, -CLng(ok))
                        End If
                    End If
                ElseIf STRUCTURE_isFactory(s) Then
                    If aStruct(s).functionality.factory.psSubject = 0 Then
                        psTempl = chooseTemplate(st, s)
                        If psTempl <> 0 Then
                            ok = structSetManufacture(s, psTempl, ModeImmediate)
                            putCaptured
                            putOut va("factory %u: %s ok%d\n", aStruct(s).id, apsTemplates(psTempl).id, -CLng(ok))
                        End If
                    End If
                End If
            End If
            it = lnNext(it)
        Loop
    Next
End Sub

' an object's id, 0 for nullptr
Private Function objId0(ByVal h As Long) As Long
    If h <> 0 Then objId0 = SIMPLE_OBJECT_id(h)
End Function

Private Function baseStructLine(ByVal s As Long) As String
    Dim subj As Long, w As String
    w = va("struct %u: p%d %s status%d body%u res%d tgt%u", aStruct(s).id, aStruct(s).player, asStructureStats(aStruct(s).pStructureType).id, _
           aStruct(s).status, aStruct(s).body, aStruct(s).resistance, objId0(aStruct(s).psTarget(0)))
    If aStruct(s).pFunctionality <> 0 Then
        Select Case asStructureStats(aStruct(s).pStructureType).typ
        Case REF_RESEARCH
            subj = aStruct(s).functionality.researchFacility.psSubject
            If subj <> -1 Then
                w = w & va(" research %d pts%u", subj, asPlayerResList(aStruct(s).player, subj).currentPoints)
            Else
                w = w & " research none"
            End If
        Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
            subj = aStruct(s).functionality.factory.psSubject
            If subj <> 0 Then
                w = w & va(" making %s left%d started%u", apsTemplates(subj).id, aStruct(s).functionality.factory.buildPointsRemaining, _
                           aStruct(s).functionality.factory.timeStarted)
            Else
                w = w & " making none"
            End If
        Case REF_REPAIR_FACILITY
            subj = aStruct(s).functionality.repairFacility.psObj
            w = w & va(" repairing %u state%d", objId0(subj), aStruct(s).functionality.repairFacility.state)
        Case REF_REARM_PAD
            subj = aStruct(s).functionality.rearmPad.psObj
            w = w & va(" rearming %u", objId0(subj))
        End Select
    End If
    baseStructLine = w & vbLf
End Function

Private Sub mapBase()
    Dim st As Double, i As Long, p As Long, tick As Long, it As Long, crc As Long, nProj As Long, pr As Long
    Dim s As String, nDroids As Long, droids() As Long, psDroid As Long, t As Long, c As Long, n As Long
    Dim limits(0 To MAX_PLAYERS - 1, 0 To 2) As Long
    st = 17 + mapWidth * 23 + mapHeight * 13
    fpathInitialise
    gridInitialise
    ' after the earlier blocks' times (the droids' lastFired)
    gameTime = 20000
    deltaGameTime = GAME_TICKS_PER_UPDATE
    game.map = "mapbase"
    resetDamageModifiers
    proj_InitSystem
    mapInit
    baseSnapshot True
    ' research and production 50 times faster; every component available (the templates researched)
    For i = 0 To numStructureStats - 1
        For p = 0 To MAX_PLAYERS - 1
            With asStructureStats(i).upgrade(p)
                .research = .research * 50
                .moduleResearch = .moduleResearch * 50
                .production = .production * 50
                .moduleProduction = .moduleProduction * 50
            End With
        Next
    Next
    For p = 0 To MAX_PLAYERS - 1
        For t = 0 To COMP_NUMCOMPONENTS - 1
            For c = 0 To apCompLists_count(t) - 1
                apCompLists(p, t, c) = AVAILABLE
            Next
        Next
        setPower p, 3000 + 1000 * (p Mod 3)
        ' rules.js's droid limits
        limits(p, 0) = getMaxDroids(p)
        limits(p, 1) = getMaxCommanders(p)
        limits(p, 2) = getMaxConstructors(p)
        setMaxDroids p, 150
        setMaxCommanders p, 10
        setMaxConstructors p, 15
    Next
    ' damaged droids go for repair, VTOLs to rearm
    nDroids = allDroids(droids)
    For i = 0 To nDroids - 1
        psDroid = droids(i)
        If DROID_isVtol(psDroid) Then
            moveToRearm psDroid
            putCaptured
            putOut va("rearm %u: action%d\n", aDroid(psDroid).id, aDroid(psDroid).action)
        ElseIf DROID_isDamaged(psDroid) Then
            If testRandMod(st, 2) = 0 Then
                orderDroid psDroid, DORDER_RTR, ModeImmediate
                putCaptured
                putOut va("rtr %u: order%d\n", aDroid(psDroid).id, aDroid(psDroid).order.typ)
            End If
        End If
    Next
    baseGiveWork st
    countUpdate False
    gridReset
    putCaptured
    For tick = 1 To 100
        gameTime = gameTime + GAME_TICKS_PER_UPDATE
        deltaGameTime = GAME_TICKS_PER_UPDATE
        gameStateUpdate
        putCaptured
        If tick = 1 Or tick = 100 Or syncDumpAll Then
            putOut dumpCurrentDebugSyncLog()
        End If
        nProj = 0
        pr = proj_GetFirst()
        Do While pr <> 0
            nProj = nProj + 1
            pr = proj_GetNext()
        Loop
        crc = nextDebugSync(gameTime)
        putCaptured
        nDroids = allDroids(droids)
        putOut va("tick %d: crc=%04X proj=%d droids=%d\n", tick, crc, nProj, nDroids)
        If tick Mod 25 = 0 Then
            For p = 0 To MAX_PLAYERS - 1
                it = apsStructLists(p).head
                Do While it <> 0
                    If aStruct(lnObj(it)).pFunctionality <> 0 Or aStruct(lnObj(it)).numWeaps > 0 Then
                        putOut baseStructLine(lnObj(it))
                    End If
                    it = lnNext(it)
                Loop
            Next
            baseGiveWork st
        End If
    Next
    s = "count:"
    For p = 0 To MAX_PLAYERS - 1
        n = 0
        For i = 0 To asResearch_size - 1
            If IsResearchCompleted(asPlayerResList(p, i)) Then n = n + 1
        Next
        s = s & va(" %d/%d/%d/%d", apsDroidLists(p).size, apsStructLists(p).size, n, getPower(p))
    Next
    putOut s & vbLf
    putOut va("features %d\n", apsFeatureList(0).size)
    For p = 0 To MAX_PLAYERS - 1
        setMaxDroids p, limits(p, 0)
        setMaxCommanders p, limits(p, 1)
        setMaxConstructors p, limits(p, 2)
    Next
    baseSnapshot False
    proj_FreeAllProjectiles
    gridShutDown
    fpathShutdown
    putCaptured
    gameTime = 0
    deltaGameTime = 0
End Sub

' the header line of the sync log nextDebugSync closed last
Private Function netLastLogHeader() As String
    Dim s As String, k As Long
    s = dumpLocalDebugSyncLog((syncDebugNextIndex() + MAX_SYNC_HISTORY - 1) Mod MAX_SYNC_HISTORY)
    k = InStr(s, vbLf)
    If k > 0 Then s = Left$(s, k)
    netLastLogHeader = s
End Function

' a topic the player may start (researchAvailable, ModeQueue, as wzapi's pursueResearch asks), from a random start; -1 if none
Private Function chooseAvailableResearch(st As Double, ByVal player As Long) As Long
    Dim n As Long, start As Long, k As Long, i As Long
    chooseAvailableResearch = -1
    n = asResearch_size
    If n = 0 Then Exit Function
    start = testRandMod(st, n)
    For k = 0 To n - 1
        i = (start + k) Mod n
        If researchAvailable(i, player, ModeQueue) Then
            chooseAvailableResearch = i
            Exit Function
        End If
    Next
End Function

' work for the labs and factories through the game messages, as the interface and the scripts give it: a topic for
' each idle lab (sendResearchStatus, as wzapi's pursueResearch), a template for each idle factory (structSetManufacture,
' ModeQueue); the first busy lab and factory of each player held, released or cancelled (by round)
Private Sub netGiveWork(st As Double, ByVal round_ As Long)
    Dim p As Long, it As Long, s As Long, index As Long, psTempl As Long, ok As Boolean, labDone As Boolean, facDone As Boolean
    For p = 0 To MAX_PLAYERS - 1
        labDone = False
        facDone = False
        it = apsStructLists(p).head
        Do While it <> 0
            s = lnObj(it)
            If aStruct(s).status = SS_BUILT And aStruct(s).pFunctionality <> 0 Then
                If asStructureStats(aStruct(s).pStructureType).typ = REF_RESEARCH Then
                    If aStruct(s).functionality.researchFacility.psSubject = -1 Then
                        index = chooseAvailableResearch(st, p)
                        If index >= 0 Then
                            sendResearchStatus s, index, p, True
                            putCaptured
                            putOut va("lab %u: research %d\n", aStruct(s).id, index)
                        End If
                    ElseIf Not labDone Then
                        labDone = True
                        Select Case round_ Mod 3
                        Case 0
                            holdResearch s, ModeQueue
                        Case 1
                            releaseResearch s, ModeQueue
                        Case Else
                            cancelResearch s, ModeQueue
                        End Select
                        putCaptured
                        putOut va("lab %u: %d\n", aStruct(s).id, round_ Mod 3)
                    End If
                ElseIf STRUCTURE_isFactory(s) Then
                    If aStruct(s).functionality.factory.psSubject = 0 Then
                        psTempl = chooseTemplate(st, s)
                        If psTempl <> 0 Then
                            ok = structSetManufacture(s, psTempl, ModeQueue)
                            putCaptured
                            putOut va("factory %u: %s ok%d\n", aStruct(s).id, apsTemplates(psTempl).id, -CLng(ok))
                        End If
                    ElseIf Not facDone Then
                        facDone = True
                        Select Case round_ Mod 3
                        Case 0
                            holdProduction s, ModeQueue
                        Case 1
                            releaseProduction s, ModeQueue
                        Case Else
                            cancelProduction s, ModeQueue
                        End Select
                        putCaptured
                        putOut va("factory %u: %d\n", aStruct(s).id, round_ Mod 3)
                    End If
                End If
            End If
            it = lnNext(it)
        Loop
    Next
End Sub

' orders through the game messages for up to 24 droids (a third of them, from a random start), secondary orders for
' some others
Private Sub netOrders(st As Double, avail() As Long, ByVal nAvail As Long, ByVal wallStat As Long, ByVal wx As Long, ByVal wy As Long)
    Dim droids() As Long, nDroids As Long, i As Long, k As Long, n As Long, psDroid As Long, how As String
    nDroids = allDroids(droids)
    If nDroids = 0 Then Exit Sub
    k = testRandMod(st, 3)
    For i = 0 To nDroids - 1
        psDroid = droids(i)
        If (i + k) Mod 3 = 0 Then
            If n < 24 Then
                n = n + 1
                how = giveOrder(st, psDroid, avail, nAvail, wallStat, wx, wy, ModeQueue)
                putCaptured
                putOut va("order %u: %s\n", aDroid(psDroid).id, how)
            End If
        ElseIf i Mod 5 = 1 Then
            If i Mod 2 <> 0 Then
                secondarySetState psDroid, DSO_ATTACK_RANGE, DSS_ARANGE_SHORT, ModeQueue
            Else
                secondarySetState psDroid, DSO_HALTTYPE, DSS_HALT_HOLD, ModeQueue
            End If
            putCaptured
            putOut va("secondary %u: %X\n", aDroid(psDroid).id, aDroid(psDroid).secondaryOrderPending)
        End If
    Next
End Sub

' the debug mode's messages: refused in a skirmish (debugMappingsAllowed is false), done in the campaign
Private Sub netDebugMessages(st As Double)
    Dim droids() As Long, nDroids As Long, psDroid As Long, s As Long, p As Long, psTempl As Long, noOrders As INITIAL_DROID_ORDERS
    nDroids = allDroids(droids)
    If nDroids > 0 Then
        psDroid = droids(testRandMod(st, nDroids))
        SendDestroyDroid psDroid
        putCaptured
        putOut va("debug: destroy droid %u\n", aDroid(psDroid).id)
        psDroid = droids(testRandMod(st, nDroids))
        p = aDroid(psDroid).player
        psTempl = 0
        If templateCount(p) > 0 Then psTempl = droidTemplates_at(p, testRandMod(st, templateCount(p)))
        If psTempl <> 0 Then
            SendDroid psTempl, aDroid(psDroid).pos.x, aDroid(psDroid).pos.y, p, generateNewObjectId(), False, noOrders
            putCaptured
            putOut va("debug: droid %s for p%d\n", apsTemplates(psTempl).id, p)
        End If
    End If
    s = randomStructure(st, testRandMod(st, MAX_PLAYERS))
    If s <> 0 Then
        SendDestroyStructure s
        putCaptured
        putOut va("debug: destroy structure %u\n", aStruct(s).id)
    End If
    s = randomStructure(st, testRandMod(st, MAX_PLAYERS))
    If s <> 0 Then
        SendBuildFinished s
        putCaptured
        putOut va("debug: build finished %u\n", aStruct(s).id)
    End If
End Sub

Private Sub netRun(ByVal mode As Long)
    Dim st As Double, wx As Long, wy As Long, i As Long, p As Long, t As Long, c As Long, n As Long, k As Long, it As Long
    Dim buildable As Variant, avail() As Long, nAvail As Long, wallStat As Long, frame As Long, nProj As Long, pr As Long
    Dim s As String, droids() As Long, nDroids As Long, savedMultiMessages As Boolean, savedIsHost As Boolean
    Dim savedAllocated(0 To MAX_CONNECTED_PLAYERS - 1) As Boolean
    Dim limits(0 To MAX_PLAYERS - 1, 0 To 2) As Long
    Dim nFrames As Long, workEvery As Long, savedSelected As Long, handler As ReplayOptionsHandler, ver As Long
    st = 29 + mapWidth * 31 + mapHeight * 7 + 1000 * mode
    If mode = NET_MODE_NET Then
        nFrames = 80
        workEvery = 20
    Else
        nFrames = 40
        workEvery = 10
    End If
    wx = world_coord(mapWidth)
    wy = world_coord(mapHeight)
    fpathInitialise
    gridInitialise
    game.map = "mapnet"
    resetDamageModifiers
    proj_InitSystem
    mapInit
    baseSnapshot True
    ' research and production 50 times faster; every component available (the templates researched)
    For i = 0 To numStructureStats - 1
        For p = 0 To MAX_PLAYERS - 1
            With asStructureStats(i).upgrade(p)
                .research = .research * 50
                .moduleResearch = .moduleResearch * 50
                .production = .production * 50
                .moduleProduction = .moduleProduction * 50
            End With
        Next
    Next
    For p = 0 To MAX_PLAYERS - 1
        For t = 0 To COMP_NUMCOMPONENTS - 1
            For c = 0 To apCompLists_count(t) - 1
                apCompLists(p, t, c) = AVAILABLE
            Next
        Next
        setPower p, 4000 + 1000 * (p Mod 2)
        ' rules.js's droid limits
        limits(p, 0) = getMaxDroids(p)
        limits(p, 1) = getMaxCommanders(p)
        limits(p, 2) = getMaxConstructors(p)
        setMaxDroids p, 150
        setMaxCommanders p, 10
        setMaxConstructors p, 15
    Next
    ' the structures the trucks may build (as maporders)
    buildable = Array("A0LightFactory", "A0PowerGenerator", "A0ResearchFacility", "A0HardcreteMk1Wall", "WallTower01", _
                      "A0RepairCentre3", "Sys-SensoTower02", "A0CommandCentre", "A0ResourceExtractor", "A0FacMod1", _
                      "A0PowMod1", "A0ResearchModule1", "PillBox1", "A0CyborgFactory", "A0VTolFactory1", "A0VtolPad")
    ReDim avail(0 To UBound(buildable))
    For k = 0 To UBound(buildable)
        i = getStructStatFromName(buildable(k))
        If i >= 0 Then
            avail(nAvail) = i
            nAvail = nAvail + 1
            For p = 0 To MAX_PLAYERS - 1
                apStructTypeLists(p, i) = AVAILABLE
            Next
        End If
    Next
    wallStat = getStructStatFromName("A0HardcreteMk1Wall")
    ' a skirmish host without a network: the messages go through the game queues (init.cpp's stageTwoInitialise);
    ' the clock is gtime's, on a clock that moves 100 ms a frame
    savedMultiMessages = bMultiMessages
    savedIsHost = NetPlay_isHost
    savedSelected = selectedPlayer
    bMultiMessages = bMultiPlayer           ' as a game sets it: SPinit (the campaign) false, a skirmish host true
    NetPlay_isHost = True
    realSelectedPlayer = selectedPlayer
    ' the selected player the only human (the others AIs, which the host is responsible for)
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        savedAllocated(i) = NetPlay_players_allocated(i)
        NetPlay_players_allocated(i) = i = selectedPlayer
    Next
    If netTicks = 0 Then netTicks = 1000
    wzSetTestTicks True, netTicks
    Set handler = New ReplayTestHandler
    If mode = NET_MODE_PLAY Then
        ' a replay viewer (game.cpp's load of a .wzrp: NETloadReplay, then the game's start: gameTimeInit), on one
        ' tick a frame (main.cpp: no fast forward for replays outside headless and autogame runs)
        NetPlay_isHost = False
        If Not NETloadReplay(netReplayFile, handler) Then putOut "replay not loaded" & vbLf
        setMaxFastForwardTicks True, 0, True
    End If
    gameTimeInit
    ' after the earlier blocks' times (the droids' lastFired)
    setGameTime 40000
    If mode <> NET_MODE_PLAY Then
        sendQueuedDroidInfo
        For i = 0 To MAX_CONNECTED_PLAYERS - 1
            NETinitQueue NETgameQueue(i)
            If Not myResponsibility(i) Or Not NetPlay_bComms Then
                NETsetNoSendOverNetwork NETgameQueue(i)
            End If
        Next
    End If
    If mode = NET_MODE_RECORD Then
        ' the lobby's options the replay saves, as loadConfig sets them without a config file (statsref has none)
        game.name_ = "My Game"
        game.techLevel = 1
        game.power = LEV_MED
        game.base = CAMP_BASE
        game.inactivityMinutes = 5
        game.gameTimeLimitMinutes = 0
        game.playerLeaveMode = 1                ' PLAYER_LEAVE_MODE::SPLIT_WITH_TEAM
        ' the game's seed, as at a game's start (the replay restores it)
        gameSRand 4242
        netReplayFile = NETreplaySaveStart("skirmish", handler, 0, False)
        If Len(netReplayFile) = 0 Then putOut "replay not started" & vbLf
    End If
    countUpdate False
    gridReset
    putCaptured
    If mode <> NET_MODE_PLAY Then
        netGiveWork st, 0
        netOrders st, avail, nAvail, wallStat, wx, wy
    End If
    For frame = 1 To nFrames
        netTicks = netTicks + GAME_TICKS_PER_UPDATE
        wzSetTestTicks True, netTicks
        realTimeUpdate
        gameLoop
        putCaptured
        If frame = 2 Or frame = nFrames Or syncDumpAll Then
            putOut dumpLocalDebugSyncLog((syncDebugNextIndex() + MAX_SYNC_HISTORY - 1) Mod MAX_SYNC_HISTORY)
        End If
        nProj = 0
        pr = proj_GetFirst()
        Do While pr <> 0
            nProj = nProj + 1
            pr = proj_GetNext()
        Loop
        nDroids = allDroids(droids)
        putOut va("frame %d: gameTime=%u proj=%d droids=%d ", frame, gameTime, nProj, nDroids) & netLastLogHeader()
        If frame Mod workEvery = 0 Then
            For p = 0 To MAX_PLAYERS - 1
                it = apsStructLists(p).head
                Do While it <> 0
                    If aStruct(lnObj(it)).pFunctionality <> 0 Then
                        putOut baseStructLine(lnObj(it))
                    End If
                    it = lnNext(it)
                Loop
            Next
            If frame < nFrames And mode <> NET_MODE_PLAY Then
                netGiveWork st, frame \ workEvery
                netOrders st, avail, nAvail, wallStat, wx, wy
                If frame = 2 * workEvery Then netDebugMessages st
            End If
        End If
    Next
    If mode = NET_MODE_RECORD Then
        If Not NETreplaySaveStop(handler) Then putOut "replay not saved" & vbLf
        putCaptured
        putOut replaySummary(netReplayFile)
    ElseIf mode = NET_MODE_PLAY Then
        NETshutdownReplay
        NetPlay_players_resize MAX_CONNECTED_PLAYERS
        selectedPlayer = savedSelected
        setMaxFastForwardTicks False, 0, True
    End If
    s = "count:"
    For p = 0 To MAX_PLAYERS - 1
        n = 0
        For i = 0 To asResearch_size - 1
            If IsResearchCompleted(asPlayerResList(p, i)) Then n = n + 1
        Next
        s = s & va(" %d/%d/%d/%d", apsDroidLists(p).size, apsStructLists(p).size, n, getPower(p))
    Next
    putOut s & vbLf
    putOut va("features %d\n", apsFeatureList(0).size)
    For p = 0 To MAX_PLAYERS - 1
        setMaxDroids p, limits(p, 0)
        setMaxCommanders p, limits(p, 1)
        setMaxConstructors p, limits(p, 2)
    Next
    baseSnapshot False
    sendQueuedDroidInfo
    NETdeleteQueue
    bMultiMessages = savedMultiMessages
    NetPlay_isHost = savedIsHost
    realSelectedPlayer = selectedPlayer
    For i = 0 To MAX_CONNECTED_PLAYERS - 1
        NetPlay_players_allocated(i) = savedAllocated(i)
    Next
    proj_FreeAllProjectiles
    gridShutDown
    fpathShutdown
    setGameTime 0
    putCaptured
End Sub
' a replay file as text: its header, its settings (the JSON as written), the embedded map's size, the messages'
' count by type with a CRC of their bytes, and the end-of-game JSON
Private Function replaySummary(ByVal filename As String) As String
    Dim f As Long, data() As Byte, size As Long, pos As Long, n As Long, crc As Long, out As String, i As Long
    Dim counts(0 To 255) As Long, typ As Long, length As Long, start As Long
    f = PHYSFS_openRead(filename)
    If f = 0 Then
        replaySummary = "replay file missing" & vbLf
        Exit Function
    End If
    size = PHYSFS_fileLength(f)
    If size > 0 Then
        ReDim data(0 To size - 1)
        PHYSFS_readBytes f, data, 0, size
    End If
    PHYSFS_close f
    out = va("replay %d bytes, magic %08X\n", size, replayU32(data, 0))
    pos = 4
    n = replayU32(data, pos)
    out = out & "settings: " & utf8_decode(data, pos + 4, n) & vbLf
    pos = pos + 4 + n
    out = out & va("map data v%u, %u bytes\n", replayU32(data, pos), replayU32(data, pos + 4))
    pos = pos + 8 + replayU32(data, pos + 4)
    start = pos
    n = 0
    Do While pos + 4 <= size
        typ = data(pos + 1)
        length = CLng(data(pos + 2)) * 256 + data(pos + 3)
        counts(typ) = counts(typ) + 1
        n = n + 1
        pos = pos + 4 + length
        If typ = 255 Then Exit Do               ' REPLAY_ENDED
    Loop
    crc = wz_crc_update(wz_crc_init(), data, start, pos - start)
    out = out & va("messages %d, crc %08X:", n, crc)
    For i = 0 To 255
        If counts(i) > 0 Then out = out & va(" %s=%d", messageTypeToString(i), counts(i))
    Next
    out = out & vbLf
    n = replayU32(data, pos)
    out = out & "end: " & utf8_decode(data, pos + 4, n) & va(" (%u)\n", replayU32(data, pos + 4 + n))
    replaySummary = out
End Function

' a big-endian uint32_t of a replay file
Private Function replayU32(data() As Byte, ByVal pos As Long) As Long
    replayU32 = FromU32(((CDbl(data(pos)) * 256# + data(pos + 1)) * 256# + data(pos + 2)) * 256# + data(pos + 3))
End Function

Private Sub mapNet()
    netRun NET_MODE_NET
End Sub

' a replay read as the game reads one (NETreplayLoadStart with the tests' handler, then every message): the options as
' restored (saved again by the game's own handler), the messages' count by player and type, a CRC of their bytes, and
' the first ten in hex
Private Sub replayInfo(ByVal filename As String)
    Dim handler As ReplayOptionsHandler, gameHandler As ReplayOptionsHandler, ver As Long, msg As NetMessage, player As Long
    Dim j As nlohmann_json, n As Long, crc As Long, data() As Byte, size As Long, i As Long, p As Long, s As String
    Dim counts(0 To 255, 0 To MAX_GAMEQUEUE_SLOTS - 1) As Long, savedSelected As Long
    savedSelected = selectedPlayer
    Set handler = New ReplayTestHandler
    If Not NETreplayLoadStart(filename, handler, ver) Then
        putCaptured
        putOut "replay not loaded" & vbLf
        Exit Sub
    End If
    putCaptured
    putOut va("format %d, selectedPlayer %d, players %d\n", ver, selectedPlayer, NetPlay_players_size())
    Set gameHandler = New WZGameReplayOptionsHandler
    Set j = json_object()
    gameHandler.saveOptions j
    putOut "options: " & j.dump() & vbLf
    crc = wz_crc_init()
    Do While NETreplayLoadNetMessage(msg, player)
        data = msg.rawData()
        size = UBound(data) + 1
        If player >= 0 And player < MAX_GAMEQUEUE_SLOTS Then counts(msg.type_(), player) = counts(msg.type_(), player) + 1
        crc = wz_crc_update(crc, data, 0, size)
        If n < 10 Then
            s = va("message %d: p%d", n, player)
            For i = 0 To size - 1
                s = s & va(" %02X", data(i))
            Next
            putOut s & vbLf
        End If
        n = n + 1
        If msg.type_() = REPLAY_ENDED Then Exit Do
    Loop
    NETreplayLoadStop
    putCaptured
    s = va("messages %d, crc %08X:", n, crc)
    For p = 0 To MAX_GAMEQUEUE_SLOTS - 1
        For i = 0 To 255
            If counts(i, p) > 0 Then s = s & va(" p%d %s=%d", p, messageTypeToString(i), counts(i, p))
        Next
    Next
    putOut s & vbLf
    NETdeleteQueue
    NetPlay_players_resize MAX_CONNECTED_PLAYERS
    selectedPlayer = savedSelected
    realSelectedPlayer = selectedPlayer
    putCaptured
End Sub

Private Sub mapQueries()
    Dim st As Double, i As Long, x As Long, y As Long, wx As Long, wy As Long, h As Long, mx As Long, mn As Long
    Dim sx As Long, sy As Long, sz As Long, dx As Long, dy As Long, dz As Long, r As Long
    Dim cx As Long, cy As Long, vx As Long, vy As Long, ix As Long, iy As Long, ok As Boolean, s As String, t As Long
    st = 1 + mapWidth * 7 + mapHeight
    wx = world_coord(mapWidth)
    wy = world_coord(mapHeight)
    s = "heights:"
    For i = 1 To 100
        x = testRandMod(st, wx + TILE_UNITS) - TILE_UNITS \ 2
        y = testRandMod(st, wy + TILE_UNITS) - TILE_UNITS \ 2
        s = s & " " & CStr(x) & "," & CStr(y) & "=" & CStr(map_Height(x, y))
    Next
    putCaptured
    putOut s & vbLf
    s = "maxmin:"
    For i = 1 To 100
        x = testRandMod(st, mapWidth + 2) - 1
        y = testRandMod(st, mapHeight + 2) - 1
        getTileMaxMin x, y, mx, mn
        s = s & " " & CStr(x) & "," & CStr(y) & "=" & CStr(mx) & "," & CStr(mn)
    Next
    putCaptured
    putOut s & vbLf
    s = "lines:"
    For i = 1 To 100
        sx = testRandMod(st, wx)
        sy = testRandMod(st, wy)
        sz = map_Height(sx, sy) + testRandMod(st, 400) - 100
        dx = testRandMod(st, wx)
        dy = testRandMod(st, wy)
        dz = testRandMod(st, 1200) - 200
        r = map_LineIntersect(sx, sy, sz, dx, dy, dz, 65536)
        s = s & " " & CStr(sx) & "," & CStr(sy) & "," & CStr(sz) & ">" & CStr(dx) & "," & CStr(dy) & "," & CStr(dz) & "=" & CStr(U32(r))
    Next
    putCaptured
    putOut s & vbLf
    If wx > 2 * TILE_UNITS And wy > 2 * TILE_UNITS Then
        s = "intersect:"
        For i = 1 To 100
            cx = testRandMod(st, wx - 2 * TILE_UNITS) + TILE_UNITS
            cy = testRandMod(st, wy - 2 * TILE_UNITS) + TILE_UNITS
            vx = testRandMod(st, 401) - 200
            vy = testRandMod(st, 401) - 200
            ix = 0: iy = 0
            s = s & " " & CStr(cx) & "," & CStr(cy) & "," & CStr(vx) & "," & CStr(vy)
            ok = map_Intersect(cx, cy, vx, vy, ix, iy)
            s = s & "=" & CStr(-CLng(ok)) & "," & CStr(cx) & "," & CStr(cy) & "," & CStr(vx) & "," & CStr(vy) & "," & CStr(ix) & "," & CStr(iy)
        Next
        putCaptured
        putOut s & vbLf
    End If
    s = "blocking:"
    For i = 1 To 100
        x = testRandMod(st, mapWidth + 2) - 1
        y = testRandMod(st, mapHeight + 2) - 1
        s = s & " " & CStr(x) & "," & CStr(y) & "=" & CStr(-CLng(fpathBlockingTile(x, y, PROPULSION_TYPE_WHEELED))) _
            & CStr(-CLng(fpathBlockingTile(x, y, PROPULSION_TYPE_HOVER))) & CStr(-CLng(fpathBlockingTile(x, y, PROPULSION_TYPE_LIFT))) _
            & CStr(-CLng(fpathBlockingTile(x, y, PROPULSION_TYPE_PROPELLOR)))
    Next
    putCaptured
    putOut s & vbLf
    s = "fire:"
    For i = 1 To 20
        gameTime = testRand(st)
        x = testRandMod(st, wx)
        y = testRandMod(st, wy)
        r = testRandMod(st, 5000)
        tileSetFire x, y, r
        t = worldTile(x, y)
        s = s & " " & CStr(x) & "," & CStr(y) & "," & CStr(r) & "=" & Hex$(psMapTiles(t).tileInfoBits) & "," & CStr(psMapTiles(t).fireEndTime) _
            & "," & CStr(-CLng(fireOnLocation(x, y)))
    Next
    gameTime = 0
    putCaptured
    putOut s & vbLf
End Sub

Private Function loadKind(ByVal kind As String, ByVal file As String) As Boolean
    Dim ini As WzConfig
    If kind = "propsounds" Then
        loadKind = loadPropulsionSounds(file)
        Exit Function
    End If
    If kind = "templates" Then
        loadKind = loadDroidTemplates(file)
        Exit Function
    End If
    Set ini = WzConfig_new(file, WzConfig_ReadOnlyAndRequired)
    Select Case kind
    Case "weapons": loadKind = loadWeaponStats(ini): If loadKind Then loadKind = allocComponentList(COMP_WEAPON, asWeaponStats_size)
    Case "body": loadKind = loadBodyStats(ini): If loadKind Then loadKind = allocComponentList(COMP_BODY, asBodyStats_size)
    Case "brain": loadKind = loadBrainStats(ini): If loadKind Then loadKind = allocComponentList(COMP_BRAIN, asBrainStats_size)
    Case "propulsion": loadKind = loadPropulsionStats(ini): If loadKind Then loadKind = allocComponentList(COMP_PROPULSION, asPropulsionStats_size)
    Case "sensor": loadKind = loadSensorStats(ini): If loadKind Then loadKind = allocComponentList(COMP_SENSOR, asSensorStats_size)
    Case "ecm": loadKind = loadECMStats(ini): If loadKind Then loadKind = allocComponentList(COMP_ECM, asECMStats_size)
    Case "repair": loadKind = loadRepairStats(ini): If loadKind Then loadKind = allocComponentList(COMP_REPAIRUNIT, asRepairStats_size)
    Case "construct": loadKind = loadConstructStats(ini): If loadKind Then loadKind = allocComponentList(COMP_CONSTRUCT, asConstructStats_size)
    Case "proptypes": loadKind = loadPropulsionTypes(ini)
    Case "terrain"
        terrainLoaded = True
        loadKind = loadTerrainTable(ini)
    Case "weapmods": loadKind = loadWeaponModifiers(ini)
    Case "structures": loadKind = loadStructureStats(ini): If loadKind Then loadKind = allocStructLists()
    Case "structmods": loadKind = loadStructureStrengthModifiers(ini)
    Case "features": loadKind = loadFeatureStats(ini)
    Case "research"
        If asResearch_size <> 0 Then
            ResearchShutDown
        End If
        loadKind = loadResearch(ini)
    Case Else
        putOut va("unknown kind %s\n", kind)
        loadKind = False
    End Select
End Function

' runs the commands; the process's exit code
Public Function vbtest_run(ByVal commandFile As String) As Long
    Dim f As Integer, data() As Byte, content As String, lines() As String, li As Long, ln As String, w() As String
    Dim i As Long, mask As Long, ok As Boolean, s As String, sp() As String, nsp As Long, sp2() As Long

    debug_register_callback CB_debug_callback_capture, CB_NULL, CB_NULL, Empty
    ' statsref's effects and display entry points are stand-ins that do nothing (effects_c, display_c)
    vbEffectsOff = True
    vbDisplayOff = True
    ' the script engine ready, with no script loaded, as the game is after prepareScripts (statsref's
    ' triggerEvent* are stand-ins that do nothing; the engine's would assert "Scripts not initialized yet")
    prepareScripts False
    debug_capture_take
    resetOutput

    f = FreeFile
    Open commandFile For Binary Access Read As #f
    If LOF(f) > 0 Then
        ReDim data(0 To LOF(f) - 1)
        Get #f, , data
        content = utf8_decode(data, 0, UBound(data) + 1)
    End If
    Close #f
    lines = Split(Replace(content, vbCr, ""), vbLf)

    For li = 0 To UBound(lines)
        ln = lines(li)
        If Len(ln) = 0 Then GoTo next_line
        w = Split(ln, vbTab)
        Select Case w(0)
        Case "searchpath"
            ' (this process's search path, for statsref's searchpath command)
            nsp = PHYSFS_getSearchPath(sp)
            s = ""
            For i = 0 To nsp - 1
                If i > 0 Then s = s & "|"
                s = s & sp(i)
            Next
            i = fopen(w(1), "wb")
            fputs s, i
            fclose i
            GoTo next_line
        Case "output"
            flushOutput
            resetOutput
            debug_capture_take
            outFile = w(1)
            GoTo next_line
        End Select
        putOut "> " & ln & vbLf
        Select Case w(0)
        Case "debug"
            debug_enable_switch w(1)
        Case "selected"
            selectedPlayer = CLng(w(1))
        Case "players"
            mask = CLng(w(1))
            For i = 0 To MAX_PLAYERS - 1
                NetPlay_players_allocated(i) = (mask And BIT_(i)) <> 0
            Next
        Case "gamemode"
            SetGameMode CLng(w(1))
        Case "reset"
            freeComponentLists
            statsShutDown
            freeStructureLists
            structureStatsShutDown
            featureStatsShutDown
            featureInitVars
            ResearchShutDown
            researchInitVars
            droidTemplateShutDown
            objmemInitialise
            terrainLoaded = False
        Case "strings"
            ' data.cpp's dataStrResLoad
            ok = True
            If psStringRes Is Nothing Then
                ok = stringsInitialise()
            End If
            If ok Then
                ok = strresLoad(psStringRes, w(1))
            End If
            putCaptured
            putOut va("result=%d\n", -CLng(ok))
            GoTo next_line
        Case "releasestrings"
            ' data.cpp's dataStrResRelease
            If Not psStringRes Is Nothing Then
                strresDestroy psStringRes
                Set psStringRes = Nothing
            End If
        Case "getstring"
            If psStringRes Is Nothing Then s = vbNullString Else s = strresGetString(psStringRes, w(1))
            putCaptured
            If StrPtr(s) = 0 Then putOut "string=(null)" & vbLf Else putOut "string=" & s & vbLf
            GoTo next_line
        Case "load"
            ok = loadKind(w(1), w(2))
            putCaptured
            putOut va("result=%d\n", -CLng(ok))
            GoTo next_line
        Case "dump"
            putCaptured
            dump
            GoTo next_line
        Case "model"
            ' imdload.cpp's modelGet, and the model
            i = modelGet(w(1))
            putCaptured
            putOut va("errors=%u\n", getModelLoadingErrorCount())
            If i = 0 Then putOut "model=null" & vbLf Else putOut modelDump(i)
            GoTo next_line
        Case "playerfaction"
            NetPlay_players_faction(CLng(w(1))) = CLng(w(2))
        Case "factions"
            ' getEnabledFactions (in id order)
            mask = getEnabledFactions(CLng(w(1)) <> 0, sp2)
            s = "factions:"
            For i = 0 To mask - 1
                s = s & " " & CStr(sp2(i))
            Next
            putCaptured
            putOut s & vbLf
            GoTo next_line
        Case "factionmodel"
            If getFactionModelName(CLng(w(1)), w(2), s) Then putOut "model=" & s & vbLf Else putOut "model=nullopt" & vbLf
        Case "overrides"
            useTerrainOverrides = (CLng(w(1)) <> 0)
        Case "mapload"
            ' <folder> <tilesetDir or -> <players> <skirmish|campaign>
            ok = vbtestMapLoad(w(1), w(2), CLng(w(3)), w(4))
            putCaptured
            putOut va("result=%d\n", -CLng(ok))
            If psMapTiles_allocated Then mapDump
            GoTo next_line
        Case "mapreloadground"
            ok = mapReloadGroundTypes()
            putCaptured
            putOut va("result=%d\n", -CLng(ok))
            If psMapTiles_allocated Then mapDump
            GoTo next_line
        Case "mapqueries"
            putCaptured
            If psMapTiles_allocated Then mapQueries Else putOut "no map" & vbLf
            GoTo next_line
        Case "mapshutdown"
            gwShutDown
            mapShutdown
        Case "mapplayers"
            ' a skirmish or campaign game of n players: game.type, game.maxPlayers, bMultiPlayer, the start positions
            If w(1) = "campaign" Then game.typ = LEVEL_TYPE_CAMPAIGN Else game.typ = LEVEL_TYPE_SKIRMISH
            game.maxPlayers = CLng(w(2))
            game.scavengers = NO_SCAVENGERS
            game.alliance = NO_ALLIANCES
            bMultiPlayer = (w(1) <> "campaign")
            For i = 0 To MAX_PLAYERS - 1
                If i < CLng(w(2)) Then NetPlay_players_position(i) = i Else NetPlay_players_position(i) = -1
            Next
        Case "mapobjects"
            ' loadGame's object set-up, then loadWzMapFeature and loadWzMapStructure on the map mapload loaded, as loadGame orders them
            gameSRand 12345
            objInitialise
            grpInitialise
            aiInitialise
            scriptInit
            visInitialise
            initFactoryNumFlag
            droidInit
            putCaptured
            If testMap Is Nothing Or Not psMapTiles_allocated Then
                putOut "no map" & vbLf
            Else
                Set fixedIds = CreateObject("Scripting.Dictionary")
                ok = loadWzMapDroidInit(testMap, fixedIds)
                putCaptured
                putOut va("droids result=%d\n", -CLng(ok))
                ok = loadWzMapFeature(testMap, fixedIds)
                putCaptured
                putOut va("features result=%d\n", -CLng(ok))
                initStructLimits
                If game.typ <> LEVEL_TYPE_CAMPAIGN Then freeAllFlagPositions
                For i = 0 To MAX_PLAYER_SLOTS - 1
                    Set moduleToBuilding(i) = CreateObject("Scripting.Dictionary")
                Next
                ok = loadWzMapStructure(testMap, fixedIds, moduleToBuilding)
                If game.typ <> LEVEL_TYPE_CAMPAIGN Then resetFactoryNumFlag
                putCaptured
                putOut va("structures result=%d\n", -CLng(ok))
                objectsDump
            End If
            GoTo next_line
        Case "syncreset"
            resetSyncDebug
        Case "syncdumpall"
            syncDumpAll = UBound(w) >= 1
            If syncDumpAll Then syncDumpAll = w(1) = "1"
        Case "synclog"
            ' the sync log since the last synclog, then nextDebugSync
            putCaptured
            putOut dumpCurrentDebugSyncLog()
            i = nextDebugSync(gameTime)
            putCaptured
            putOut va("crc=%04X\n", i)
            GoTo next_line
        Case "mapgrid"
            putCaptured
            If psMapTiles_allocated Then mapGrid Else putOut "no map" & vbLf
            GoTo next_line
        Case "mappaths"
            putCaptured
            If psMapTiles_allocated Then mapPaths Else putOut "no map" & vbLf
            GoTo next_line
        Case "mappower"
            putCaptured
            mapPower
            GoTo next_line
        Case "mapmove"
            putCaptured
            If psMapTiles_allocated Then mapMove Else putOut "no map" & vbLf
            GoTo next_line
        Case "maporders"
            putCaptured
            If psMapTiles_allocated Then mapOrders Else putOut "no map" & vbLf
            GoTo next_line
        Case "mapcombat"
            putCaptured
            If psMapTiles_allocated Then mapCombat Else putOut "no map" & vbLf
            GoTo next_line
        Case "mapbase"
            putCaptured
            If psMapTiles_allocated Then mapBase Else putOut "no map" & vbLf
            GoTo next_line
        Case "mapnet"
            putCaptured
            If psMapTiles_allocated Then mapNet Else putOut "no map" & vbLf
            GoTo next_line
        Case "replayinfo"
            putCaptured
            replayInfo w(1)
            GoTo next_line
        Case "maprecord"
            putCaptured
            If psMapTiles_allocated Then netRun NET_MODE_RECORD Else putOut "no map" & vbLf
            GoTo next_line
        Case "mapplay"
            putCaptured
            If psMapTiles_allocated Then netRun NET_MODE_PLAY Else putOut "no map" & vbLf
            GoTo next_line
        Case "objshutdown"
            ' the messages first (their proximity displays name the objects), as a level's start frees them
            freeMessages
            freeAllDroids
            freeAllStructs
            freeAllFeatures
            objlist_clear apsOilList(0)
            objlist_clear apsSensorList(0)
            For i = 0 To MAX_PLAYERS - 1
                objlist_clear apsExtractorLists(i)
            Next
            freeAllFlagPositions
            objlist_clear psDestroyedObj
            objmemShutdown
        Case "reload"
            ' debugReloadDisplayModelsForBaseModel of a loaded model, and the model
            i = modelGet(w(1))
            If i <> 0 Then ok = debugReloadDisplayModelsForBaseModel(i)
            putCaptured
            If i = 0 Then
                putOut "model=null" & vbLf
            Else
                putOut va("result=%d\n", -CLng(ok))
                putOut modelDump(i)
            End If
            GoTo next_line
        Case Else
            putOut "unknown command" & vbLf
        End Select
        putCaptured
next_line:
    Next
    putCaptured
    flushOutput
    vbtest_run = 0
End Function
