Attribute VB_Name = "research_c"
' src/research.cpp, research.h: so far the research store (asResearch with
' its count asResearch_size), the players' research lists, the loader of
' research.json with its checks (the dependency cycle, the categories),
' and the lookups. A RESEARCH * is its index in asResearch (-1 = nullptr).
' asPlayerResList[player] is asPlayerResList(player, i) with
' asPlayerResList_size(player), one array whose last dimension is the
' longest list. optional<ResearchUpgradeCalculationMode> is the value plus
' has_researchUpgradeCalcMode. resCategories (std::unordered_map) is a
' Scripting.Dictionary of the categories, in first-appearance order (C++:
' the hash order, which only orders the log lines of invalid categories),
' and their members. std::stable_sort of a category is an insertion sort:
' the Microsoft library's stable_sort on up to 32 elements (the largest
' category has 12).
' The research results: researchResult, with the upgrades
' (eventResearchedHandleUpgrades, through wzapi_c's Stats and Upgrades
' objects, cached in cachedStatsObject (Nothing: null) and
' cachedPerPlayerUpgrades), cancelResearch, researchAvailable, the
' component replacements, listAllyResearch. The std::unordered_maps of
' strings (cachedPerPlayerRawUpgradeChange, whose two levels are one key
' "component" & vbNullChar & "parameter", and PlayerUpgradeCounts') are
' Scripting.Dictionary objects; eventResearchedHandleUpgrades accumulates
' its counts in the order the classes first appear (C++: the hash order,
' which orders only its LOG_RESEARCH lines). The function pointers of
' researchAvailable are IsResearchCancelledFunc and IsResearchStartedFunc,
' which take the mode. listAllyResearch's std::map is two arrays in the
' keys' order, and its result is copied into the caller's array (its count
' returned); fillResearchList's optional topic is has_topic plus topic.
' A lambda logs under its function (replaceComponentInTemplate).
Option Explicit

Public Const NO_RESEARCH_ICON As Long = 0
'max 'research complete' console message length
Public Const MAX_RESEARCH_MSG_SIZE As Long = 200

Private Const MPFLAGS_MAX As Long = &H3F        ' multiplay.h
Private Const UWORD_MAX_ As Long = &HFFFF&      ' (types.h's UWORD_MAX)

'used for loading in the research stats into the appropriate list
Public Const REQ_LIST As Long = 0
Public Const RED_LIST As Long = 1
Public Const RES_LIST As Long = 2

' (research.h's anonymous enum)
Public Enum research_h_RID
    RID_ROCKET
    RID_CANNON
    RID_HOVERCRAFT
    RID_ECM
    RID_PLASCRETE
    RID_TRACKS
    RID_DROIDTECH
    RID_WEAPONTECH
    RID_COMPUTERTECH
    RID_POWERTECH
    RID_SYSTEMTECH
    RID_STRUCTURETECH
    RID_CYBORGTECH
    RID_DEFENCE
    RID_QUESTIONMARK
    RID_GRPACC
    RID_GRPUPG
    RID_GRPREP
    RID_GRPROF
    RID_GRPDAM
    RID_MAXRID
End Enum

' (enum class ResearchUpgradeCalculationMode)
Public Enum ResearchUpgradeCalculationMode
    ' Default / compat cumulative upgrade handling (the only option for many years - from at least 3.x/(3.2+?)-4.4.2?)
    ' This can accumulate noticeable error, especially if repeatedly upgrading small values by small percentages (commonly impacted: armour, thermal)
    ' However, research.json created and tested during this long period may be expecting this outcome / behavior
    ResearchUpgradeCalculationMode_Compat

    ' "Improved" cumulative upgrade handling (significantly reduces accumulated errors)
    ResearchUpgradeCalculationMode_Improved
End Enum

' (enum class BodyClass)
Public Enum BodyClass
    BodyClass_Tank
    BodyClass_Cyborg
End Enum

Private Type RES_CATEGORY_MEMBERS               ' (VB) a resCategories value: std::vector<size_t>
    members() As Long
    size As Long
End Type

' The stores for the research stats
Public asResearch() As RESEARCH
Public asResearch_size As Long
Private researchUpgradeCalcMode As ResearchUpgradeCalculationMode
Private has_researchUpgradeCalcMode As Boolean
Private resCategories As Object                 ' category -> index into resCategoryMembers
Private resCategoryMembers() As RES_CATEGORY_MEMBERS

'used for Callbacks to say which topic was last researched
Public psCBLastResearch As Long                 ' RESEARCH *
Public psCBLastResStructure As Long             ' STRUCTURE *
Public CBResFacilityOwner As Long

'List of pointers to arrays of PLAYER_RESEARCH[numResearch] for each player
Public asPlayerResList() As PLAYER_RESEARCH     ' (player, i)
Public asPlayerResList_size(0 To MAX_PLAYERS - 1) As Long
Private asPlayerResListCap As Long

' Default level of sensor, Repair and ECM
Public aDefaultSensor(0 To MAX_PLAYERS - 1) As Long
Public aDefaultECM(0 To MAX_PLAYERS - 1) As Long
Public aDefaultRepair(0 To MAX_PLAYERS - 1) As Long

'flag that indicates whether the player can self repair
Private bSelfRepair(0 To MAX_PLAYERS - 1) As Long

' CycleDetection's sets (by research index)
Private cdVisited() As Boolean
Private cdExploring() As Boolean

' wzapi's Stats object and the players' Upgrades, made by the first research result
Private cachedStatsObject As nlohmann_json              ' Nothing: nlohmann::json(nullptr)
Private cachedPerPlayerUpgrades() As PerPlayerUpgrades
Private cachedPerPlayerUpgrades_size As Long

' RawResearchUpgradeChangeValues per player: int64_t totals (Double) by component name & vbNullChar & parameter
Private cachedPerPlayerRawUpgradeChange(0 To MAX_PLAYERS - 1) As Object

' Per-player statistics about research upgrades
Private Type PLAYER_UPGRADE_COUNTS              ' struct PlayerUpgradeCounts
    numBodyClassArmourUpgrades As Object        ' std::unordered_map<std::string, uint32_t>
    numBodyClassThermalUpgrades As Object
    numWeaponImpactClassUpgrades As Object
End Type
Private playerUpgradeCounts() As PLAYER_UPGRADE_COUNTS
Private playerUpgradeCounts_size As Long

' (VB, for the tests: research_testState) the saved state
Private testSelfRepair(0 To MAX_PLAYERS - 1) As Long
Private testRawUpgradeChange(0 To MAX_PLAYERS - 1) As Object
Private testUpgradeCounts() As PLAYER_UPGRADE_COUNTS
Private testUpgradeCounts_size As Long

Public Type AllyResearch
    player As Long                              ' unsigned
    completion As Long
    powerNeeded As Long
    timeToResearch As Long
    active As Boolean
End Type

' listAllyResearch's statics
Private Type ALLY_RESEARCH_LIST                 ' (VB) std::vector<AllyResearch>
    items() As AllyResearch
    size As Long
End Type
Private allyLastGameTime As Long                ' static uint32_t lastGameTime = ~0
Private allyResearchKeys() As Long              ' static std::map<unsigned, std::vector<AllyResearch>> researches: the keys, ascending
Private allyResearchValues() As ALLY_RESEARCH_LIST
Private allyResearch_size As Long

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub research_c_init()
    psCBLastResearch = -1
    allyLastGameTime = -1
End Sub

Public Function researchInitVars() As Boolean
    Dim i As Long
    psCBLastResearch = -1
    psCBLastResStructure = 0
    CBResFacilityOwner = -1
    Erase asResearch
    asResearch_size = 0
    has_researchUpgradeCalcMode = False
    ClearResCategories
    Set cachedStatsObject = Nothing
    Erase cachedPerPlayerUpgrades
    cachedPerPlayerUpgrades_size = 0
    ResetPlayerUpgradeCounts

    For i = 0 To MAX_PLAYERS - 1
        bSelfRepair(i) = False
        aDefaultSensor(i) = 0
        aDefaultECM(i) = 0
        aDefaultRepair(i) = 0
    Next

    researchInitVars = True
End Function

Public Function getResearchUpgradeCalcMode() As ResearchUpgradeCalculationMode
    ' Default to ResearchUpgradeCalculationMode::Compat, unless otherwise specified
    If has_researchUpgradeCalcMode Then
        getResearchUpgradeCalcMode = researchUpgradeCalcMode
    Else
        getResearchUpgradeCalcMode = ResearchUpgradeCalculationMode_Compat
    End If
End Function

' (VB) a std::unordered_map<std::string, ...>
Private Function NewStringMap() As Object
    Set NewStringMap = CreateObject("Scripting.Dictionary")
    NewStringMap.CompareMode = vbBinaryCompare
End Function

' (VB) playerUpgradeCounts = std::vector<PlayerUpgradeCounts>(MAX_PLAYERS)
Private Sub ResetPlayerUpgradeCounts()
    Dim i As Long
    ReDim playerUpgradeCounts(0 To MAX_PLAYERS - 1)
    For i = 0 To MAX_PLAYERS - 1
        Set playerUpgradeCounts(i).numBodyClassArmourUpgrades = NewStringMap()
        Set playerUpgradeCounts(i).numBodyClassThermalUpgrades = NewStringMap()
        Set playerUpgradeCounts(i).numWeaponImpactClassUpgrades = NewStringMap()
    Next
    playerUpgradeCounts_size = MAX_PLAYERS
End Sub

Private Function PlayerUpgradeCounts_getNumWeaponImpactClassUpgrades(this As PLAYER_UPGRADE_COUNTS, ByVal subClass As WEAPON_SUBCLASS) As Long
    Dim subClassStr As String
    subClassStr = getWeaponSubClass_name(subClass)
    If Not this.numWeaponImpactClassUpgrades.Exists(subClassStr) Then
        PlayerUpgradeCounts_getNumWeaponImpactClassUpgrades = 0
        Exit Function
    End If
    PlayerUpgradeCounts_getNumWeaponImpactClassUpgrades = this.numWeaponImpactClassUpgrades.Item(subClassStr)
End Function

Private Function bodyClassToStr(ByVal bc As BodyClass) As String
    Dim bodyClassStr As String
    Select Case bc
    Case BodyClass_Tank
        bodyClassStr = "Droids"
    Case BodyClass_Cyborg
        bodyClassStr = "Cyborgs"
    End Select
    bodyClassToStr = bodyClassStr
End Function

Private Function PlayerUpgradeCounts_getNumBodyClassArmourUpgrades(this As PLAYER_UPGRADE_COUNTS, ByVal bc As BodyClass) As Long
    Dim bodyClassStr As String
    bodyClassStr = bodyClassToStr(bc)
    If Not this.numBodyClassArmourUpgrades.Exists(bodyClassStr) Then
        PlayerUpgradeCounts_getNumBodyClassArmourUpgrades = 0
        Exit Function
    End If
    PlayerUpgradeCounts_getNumBodyClassArmourUpgrades = this.numBodyClassArmourUpgrades.Item(bodyClassStr)
End Function

Private Function PlayerUpgradeCounts_getNumBodyClassThermalArmourUpgrades(this As PLAYER_UPGRADE_COUNTS, ByVal bc As BodyClass) As Long
    Dim bodyClassStr As String
    bodyClassStr = bodyClassToStr(bc)
    If Not this.numBodyClassThermalUpgrades.Exists(bodyClassStr) Then
        PlayerUpgradeCounts_getNumBodyClassThermalArmourUpgrades = 0
        Exit Function
    End If
    PlayerUpgradeCounts_getNumBodyClassThermalArmourUpgrades = this.numBodyClassThermalUpgrades.Item(bodyClassStr)
End Function

Public Function getNumWeaponImpactClassUpgrades(ByVal player As Long, ByVal subClass As WEAPON_SUBCLASS) As Long
    If Not (ULT(player, playerUpgradeCounts_size)) Then ASSERT_FAILURE "player < playerUpgradeCounts.size()", "src/research.cpp:186", "getNumWeaponImpactClassUpgrades", "Out of bounds player: %u", player: getNumWeaponImpactClassUpgrades = 0: Exit Function
    getNumWeaponImpactClassUpgrades = PlayerUpgradeCounts_getNumWeaponImpactClassUpgrades(playerUpgradeCounts(player), subClass)
End Function

Public Function getNumBodyClassArmourUpgrades(ByVal player As Long, ByVal bc As BodyClass) As Long
    If Not (ULT(player, playerUpgradeCounts_size)) Then ASSERT_FAILURE "player < playerUpgradeCounts.size()", "src/research.cpp:192", "getNumBodyClassArmourUpgrades", "Out of bounds player: %u", player: getNumBodyClassArmourUpgrades = 0: Exit Function
    getNumBodyClassArmourUpgrades = PlayerUpgradeCounts_getNumBodyClassArmourUpgrades(playerUpgradeCounts(player), bc)
End Function

Public Function getNumBodyClassThermalArmourUpgrades(ByVal player As Long, ByVal bc As BodyClass) As Long
    If Not (ULT(player, playerUpgradeCounts_size)) Then ASSERT_FAILURE "player < playerUpgradeCounts.size()", "src/research.cpp:198", "getNumBodyClassThermalArmourUpgrades", "Out of bounds player: %u", player: getNumBodyClassThermalArmourUpgrades = 0: Exit Function
    getNumBodyClassThermalArmourUpgrades = PlayerUpgradeCounts_getNumBodyClassThermalArmourUpgrades(playerUpgradeCounts(player), bc)
End Function

' (VB) resCategories.clear()
Private Sub ClearResCategories()
    If resCategories Is Nothing Then
        Set resCategories = CreateObject("Scripting.Dictionary")
        resCategories.CompareMode = vbBinaryCompare
    End If
    resCategories.RemoveAll
    Erase resCategoryMembers
End Sub

' CycleDetection::explore: True and the cycle in cycle(0 .. n - 1) when one is found
Private Function CycleDetection_explore(ByVal research_ As Long, cycle() As Long, n As Long) As Boolean
    Dim k As Long, requirement As Long, m As Long
    If cdVisited(research_) Then
        CycleDetection_explore = False
        Exit Function
    End If

    If cdExploring(research_) Then
        ReDim cycle(0 To 0)
        cycle(0) = research_
        n = 1
        CycleDetection_explore = True
        Exit Function
    End If

    cdExploring(research_) = True

    For k = 0 To asResearch(research_).pPRList_size - 1
        requirement = asResearch(research_).pPRList(k)
        If CycleDetection_explore(requirement, cycle, n) Then
            ' cycle->push_front(research)
            ReDim Preserve cycle(0 To n)
            For m = n To 1 Step -1
                cycle(m) = cycle(m - 1)
            Next
            cycle(0) = research_
            n = n + 1
            CycleDetection_explore = True
            Exit Function
        End If
    Next

    cdExploring(research_) = False
    cdVisited(research_) = True
    CycleDetection_explore = False
End Function

' CycleDetection::detectCycle
Private Function CycleDetection_detectCycle(cycle() As Long, n As Long) As Boolean
    Dim research_ As Long
    If asResearch_size > 0 Then
        ReDim cdVisited(0 To asResearch_size - 1)
        ReDim cdExploring(0 To asResearch_size - 1)
    End If

    For research_ = 0 To asResearch_size - 1
        If CycleDetection_explore(research_, cycle, n) Then
            CycleDetection_detectCycle = True
            Exit Function
        End If
    Next

    CycleDetection_detectCycle = False
End Function

Private Function isResAPrereqForResB(ByVal resAIndex As Long, ByVal resBIndex As Long) As Boolean
    Dim stack() As Long, stackSize As Long, pCurr As Long, k As Long, prereqIndex As Long
    If resAIndex = resBIndex Then
        isResAPrereqForResB = False
        Exit Function
    End If
    ReDim stack(0 To 15)
    stack(0) = resBIndex
    stackSize = 1
    Do While stackSize > 0
        pCurr = stack(stackSize - 1)
        stackSize = stackSize - 1
        For k = 0 To asResearch(pCurr).pPRList_size - 1
            prereqIndex = asResearch(pCurr).pPRList(k)
            If prereqIndex = resAIndex Then
                isResAPrereqForResB = True
                Exit Function
            End If
            If stackSize > UBound(stack) Then ReDim Preserve stack(0 To stackSize * 2)
            stack(stackSize) = prereqIndex
            stackSize = stackSize + 1
        Next
    Loop
    isResAPrereqForResB = False
End Function

Private Function isResearchListValidConnectedDependencyGraph(researchIndexes() As Long, ByVal count As Long) As Boolean
    Dim a As Long, b As Long, in_degree() As Long, source_nodes As Long

    If count = 0 Then
        isResearchListValidConnectedDependencyGraph = False
        Exit Function
    End If

    ' std::unordered_map<size_t, size_t> in_degree: by position (the indices are distinct)
    ReDim in_degree(0 To count - 1)

    For a = 0 To count - 1
        For b = 0 To count - 1
            If isResAPrereqForResB(researchIndexes(a), researchIndexes(b)) Then
                in_degree(a) = in_degree(a) + 1
            End If
        Next
    Next

    For a = 0 To count - 1
        If in_degree(a) = 0 Then
            source_nodes = source_nodes + 1
        End If
    Next

    If source_nodes <> 1 Then
        debug_ 312, LOG_WZ, "isResearchListValidConnectedDependencyGraph", "Category graph has %zu potential source / edge nodes.", source_nodes
        isResearchListValidConnectedDependencyGraph = False
        Exit Function
    End If

    isResearchListValidConnectedDependencyGraph = True
End Function

' optional<ResearchUpgradeCalculationMode>: True and the mode, or False (nullopt)
Private Function resCalcModeStringToValue(ByVal calcModeStr As String, mode As ResearchUpgradeCalculationMode) As Boolean
    If calcModeStr = "compat" Then
        mode = ResearchUpgradeCalculationMode_Compat
        resCalcModeStringToValue = True
    ElseIf calcModeStr = "improved" Then
        mode = ResearchUpgradeCalculationMode_Improved
        resCalcModeStringToValue = True
    Else
        resCalcModeStringToValue = False
    End If
End Function

Private Function resCalcModeToString(ByVal mode As ResearchUpgradeCalculationMode) As String
    Select Case mode
    Case ResearchUpgradeCalculationMode_Compat
        resCalcModeToString = "compat"
        Exit Function
    Case ResearchUpgradeCalculationMode_Improved
        resCalcModeToString = "improved"
        Exit Function
    End Select
    resCalcModeToString = "invalid"
End Function

' (VB) asPlayerResList[player].push_back(item) for every player
Private Sub PushPlayerResearch(item As PLAYER_RESEARCH)
    Dim j As Long
    For j = 0 To MAX_PLAYERS - 1
        If asPlayerResList_size(j) >= asPlayerResListCap Then
            If asPlayerResListCap = 0 Then
                asPlayerResListCap = 64
                ReDim asPlayerResList(0 To MAX_PLAYERS - 1, 0 To asPlayerResListCap - 1)
            Else
                asPlayerResListCap = asPlayerResListCap * 2
                ReDim Preserve asPlayerResList(0 To MAX_PLAYERS - 1, 0 To asPlayerResListCap - 1)
            End If
        End If
        asPlayerResList(j, asPlayerResList_size(j)) = item
        asPlayerResList_size(j) = asPlayerResList_size(j) + 1
    Next
End Sub

' (VB) push_back on a UWORD or pointer list of a RESEARCH
Private Sub PushLong(list() As Long, size As Long, ByVal value As Long)
    ReDim Preserve list(0 To size)
    list(size) = value
    size = size + 1
End Sub

' Load the research stats
Public Function loadResearch(ByVal ini As WzConfig) As Boolean
    Const CONFIG_DICT_KEY_STR As String = "_config_"    ' RESEARCH_JSON_CONFIG_DICT_KEY
    Dim list() As String, n As Long, dummy As PLAYER_RESEARCH, preResearch() As StringVector
    Dim inc As Long, j As Long, k As Long, calcModeStr As String, calcModeParsed As ResearchUpgradeCalculationMode
    Dim research_ As RESEARCH, blank As RESEARCH, subGroup As String, keyTopic As Long, excludeFromCheats As Long
    Dim techCode As Long, disabledWhen As Long, iconID As String, statID As String, imdName As String, imdName2 As String
    Dim msgName As String, resPoints As Long, resPower As Long, items() As String, nItems As Long
    Dim compID As String, pComp As Long, pair() As String, nPair As Long, oldCompID As String, newCompID As String
    Dim oldComp As Long, newComp As Long, replItem As RES_COMP_REPLACEMENT, strucID As String, structIndex As Long
    Dim resID As String, preResItem As Long, cycle() As Long, nCycle As Long, cat As String, slot As Long
    Dim keys As Variant, membersOfCategory() As Long, nMembers As Long, prog As Long, categorySize As Long
    Dim val_ As Long, hole As Long

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/research.cpp:352", "loadResearch", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    ' PLAYER_RESEARCH dummy: memset to 0
    If n > 0 Then ReDim preResearch(0 To n - 1)
    For inc = 0 To n - 1
        If list(inc) = CONFIG_DICT_KEY_STR Then
            ' handle the special config dict
            ini.beginGroup list(inc)

            ' calculationMode
            calcModeStr = ini.value("calculationMode", json_variant_new(resCalcModeToString(ResearchUpgradeCalculationMode_Compat))).toWzString()
            If resCalcModeStringToValue(calcModeStr, calcModeParsed) Then
                If Not has_researchUpgradeCalcMode Then
                    researchUpgradeCalcMode = calcModeParsed
                    has_researchUpgradeCalcMode = True
                Else
                    If researchUpgradeCalcMode <> calcModeParsed Then
                        debug_ 379, LOG_ERROR, "loadResearch", "Non-matching research JSON calculationModes"
                        debug_ 380, LOG_INFO, "loadResearch", "Research JSON file ""%s"" has specified a calculationMode (""%s"") that does not match the first loaded research JSON's calculationMode (""%s"")", ini.fileName(), calcModeStr, resCalcModeToString(researchUpgradeCalcMode)
                    End If
                End If
            Else
                ASSERT_FAILURE "false", "src/research.cpp:386", "loadResearch", "Invalid _config_ ""calculationMode"" value: ""%s""", calcModeStr
                loadResearch = False
                Exit Function
            End If

            ini.endGroup
            GoTo continue_
        End If

        ' HACK FIXME: the code assumes we have empty PLAYER_RESEARCH entries to throw around
        PushPlayerResearch dummy

        ini.beginGroup list(inc)
        research_ = blank
        RESEARCH_construct research_
        research_.index = inc
        research_.name_ = ini.string_("name")
        research_.category = ini.string_("category")
        research_.id = list(inc)

        'check the name hasn't been used already
        If Not checkResearchName(research_, inc) Then ASSERT_FAILURE "checkResearchName(&research, inc)", "src/research.cpp:407", "loadResearch", "Research name '%s' used already", research_.name_: loadResearch = False: Exit Function

        research_.ref = STAT_RESEARCH + inc

        Set research_.results = ini.json("results", json_array())

        'set subGroup icon
        subGroup = ini.value("subgroupIconID", json_variant_new("")).toWzString()
        If subGroup <> "" Then
            research_.subGroup = setIconID(subGroup, research_.name_)
        Else
            research_.subGroup = NO_RESEARCH_ICON
        End If

        'set key topic
        keyTopic = ini.value("keyTopic", json_variant_new(0)).toUInt()
        If ULT(1, keyTopic) Then ASSERT_FAILURE "keyTopic <= 1", "src/research.cpp:426", "loadResearch", "Invalid keyTopic for research topic - '%s' ", research_.name_
        If Not ULT(1, keyTopic) Then
            research_.keyTopic = ini.value("keyTopic", json_variant_new(0)).toUInt() And &HFF&
        Else
            research_.keyTopic = 0
        End If

        'special flag to not reveal research from "give all" and not to research with "research all" cheats
        excludeFromCheats = ini.value("excludeFromCheats", json_variant_new(0)).toUInt()
        If ULT(1, excludeFromCheats) Then ASSERT_FAILURE "excludeFromCheats <= 1", "src/research.cpp:438", "loadResearch", "Invalid excludeFromCheats for research topic - '%s' ", research_.name_
        If Not ULT(1, excludeFromCheats) Then
            research_.excludeFromCheats = ini.value("excludeFromCheats", json_variant_new(0)).toUInt() And &HFF&
        Else
            research_.excludeFromCheats = 0
        End If

        'set tech code
        techCode = ini.value("techCode", json_variant_new(0)).toUInt() And &HFF&
        If Not (techCode <= 1) Then ASSERT_FAILURE "techCode <= 1", "src/research.cpp:450", "loadResearch", "Invalid tech code for research topic - '%s' ", research_.name_
        If techCode = 0 Then
            research_.techCode = TC_MAJOR
        Else
            research_.techCode = TC_MINOR
        End If

        'get flags when to disable tech
        disabledWhen = ini.value("disabledWhen", json_variant_new(0)).toUInt() And &HFF&
        If Not (disabledWhen <= MPFLAGS_MAX) Then ASSERT_FAILURE "disabledWhen <= 0x3f", "src/research.cpp:462", "loadResearch", "Invalid disabled tech flag for research topic - '%s' ", research_.name_
        research_.disabledWhen = disabledWhen

        'set the iconID
        iconID = ini.value("iconID", json_variant_new("")).toWzString()
        If iconID <> "" Then
            research_.iconID = setIconID(iconID, research_.name_)
        Else
            research_.iconID = NO_RESEARCH_ICON
        End If

        'get the IMDs used in the interface
        statID = ini.value("statID", json_variant_new("")).toWzString()
        research_.psStat = 0
        If statID <> "" Then
            'try find the stat with given name
            research_.psStat = getBaseStatsFromName(statID)
            If Not (research_.psStat <> 0) Then ASSERT_FAILURE "research.psStat", "src/research.cpp:483", "loadResearch", "Could not find stats for %s research %s", statID, research_.name_: loadResearch = False: Exit Function
        End If

        imdName = ini.value("imdName", json_variant_new("")).toWzString()
        If imdName <> "" Then
            research_.pIMD = modelGet(imdName)
            If Not (research_.pIMD <> 0) Then ASSERT_FAILURE "research.pIMD != nullptr", "src/research.cpp:490", "loadResearch", "Cannot find the research PIE '%s' for record '%s'", imdName, research_.name_
        End If

        imdName2 = ini.value("imdName2", json_variant_new("")).toWzString()
        If imdName2 <> "" Then
            research_.pIMD2 = modelGet(imdName2)
            If Not (research_.pIMD2 <> 0) Then ASSERT_FAILURE "research.pIMD2 != nullptr", "src/research.cpp:497", "loadResearch", "Cannot find the 2nd research '%s' PIE for record '%s'", imdName2, research_.name_
        End If

        msgName = ini.value("msgName", json_variant_new("")).toWzString()
        If msgName <> "" Then
            'check its a major tech code
            If Not (research_.techCode = TC_MAJOR) Then ASSERT_FAILURE "research.techCode == TC_MAJOR", "src/research.cpp:504", "loadResearch", "This research should not have a message associated with it, '%s' the message will be ignored!", research_.name_
            If research_.techCode = TC_MAJOR Then
                research_.pViewData = getViewData(msgName)
            End If
        End If

        'set the researchPoints
        resPoints = ini.value("researchPoints", json_variant_new(0)).toUInt()
        If ULT(UWORD_MAX_, resPoints) Then ASSERT_FAILURE "resPoints <= UWORD_MAX", "src/research.cpp:513", "loadResearch", "Research Points too high for research topic - '%s' ", research_.name_: loadResearch = False: Exit Function
        research_.researchPoints = resPoints

        'set the research power
        resPower = ini.value("researchPower", json_variant_new(0)).toUInt()
        If ULT(UWORD_MAX_, resPower) Then ASSERT_FAILURE "resPower <= UWORD_MAX", "src/research.cpp:518", "loadResearch", "Research Power too high for research topic - '%s' ", research_.name_: loadResearch = False: Exit Function
        research_.researchPower = resPower

        'remember research pre-requisites for futher checking
        nItems = ini.value("requiredResearch").toWzStringList(items)
        StringVector_clear preResearch(inc)
        For j = 0 To nItems - 1
            StringVector_push_back preResearch(inc), items(j)
        Next

        'set components results
        nItems = ini.value("resultComponents").toWzStringList(items)
        For j = 0 To nItems - 1
            compID = WzString_trimmed(items(j))
            pComp = getCompStatsFromName(compID)
            If pComp <> 0 Then
                PushLong research_.componentResults, research_.componentResults_size, pComp
            Else
                ASSERT_FAILURE "false", "src/research.cpp:536", "loadResearch", "Invalid item '%s' in list of result components of research '%s' ", compID, research_.name_
            End If
        Next

        'set replaced components
        nItems = ini.value("replacedComponents").toWzStringList(items)
        For j = 0 To nItems - 1
            'read pair of components oldComponent:newComponent
            nPair = WzString_split(items(j), ":", pair)
            If Not (nPair = 2) Then ASSERT_FAILURE "pair.size() == 2", "src/research.cpp:546", "loadResearch", "Invalid item '%s' in list of replaced components of research '%s'. Required format: 'oldItem:newItem, item1:item2'", items(j), research_.name_
            If nPair <> 2 Then
                GoTo continue_replaced 'skip invalid entries
            End If
            oldCompID = WzString_trimmed(pair(0))
            newCompID = WzString_trimmed(pair(1))
            oldComp = getCompStatsFromName(oldCompID)
            If oldComp = 0 Then
                ASSERT_FAILURE "false", "src/research.cpp:556", "loadResearch", "Invalid item '%s' in list of replaced components of research '%s'. Wrong component code.", oldCompID, research_.name_
                GoTo continue_replaced
            End If
            newComp = getCompStatsFromName(newCompID)
            If newComp = 0 Then
                ASSERT_FAILURE "false", "src/research.cpp:562", "loadResearch", "Invalid item '%s' in list of replaced components of research '%s'. Wrong component code.", newCompID, research_.name_
                GoTo continue_replaced
            End If
            replItem.pOldComponent = oldComp
            replItem.pNewComponent = newComp
            ReDim Preserve research_.componentReplacement(0 To research_.componentReplacement_size)
            research_.componentReplacement(research_.componentReplacement_size) = replItem
            research_.componentReplacement_size = research_.componentReplacement_size + 1
continue_replaced:
        Next

        'set redundant components
        nItems = ini.value("redComponents").toWzStringList(items)
        For j = 0 To nItems - 1
            compID = WzString_trimmed(items(j))
            pComp = getCompStatsFromName(compID)
            If pComp = 0 Then
                ASSERT_FAILURE "false", "src/research.cpp:579", "loadResearch", "Invalid item '%s' in list of redundant components of research '%s' ", compID, research_.name_
            Else
                PushLong research_.pRedArtefacts, research_.pRedArtefacts_size, pComp
            End If
        Next

        'set result structures
        nItems = ini.value("resultStructures").toWzStringList(items)
        For j = 0 To nItems - 1
            strucID = WzString_trimmed(items(j))
            structIndex = getStructStatFromName(strucID)
            If Not (structIndex >= 0) Then ASSERT_FAILURE "structIndex >= 0", "src/research.cpp:593", "loadResearch", "Invalid item '%s' in list of result structures of research '%s' ", strucID, research_.name_
            If structIndex >= 0 Then
                PushLong research_.pStructureResults, research_.pStructureResults_size, structIndex And &HFFFF&
            End If
        Next

        'set required structures
        nItems = ini.value("requiredStructures").toWzStringList(items)
        For j = 0 To nItems - 1
            strucID = WzString_trimmed(items(j))
            structIndex = getStructStatFromName(strucID)
            If Not (structIndex >= 0) Then ASSERT_FAILURE "structIndex >= 0", "src/research.cpp:606", "loadResearch", "Invalid item '%s' in list of required structures of research '%s' ", strucID, research_.name_
            If structIndex >= 0 Then
                PushLong research_.pStructList, research_.pStructList_size, structIndex And &HFFFF&
            End If
        Next

        'set redundant structures
        nItems = ini.value("redStructures").toWzStringList(items)
        For j = 0 To nItems - 1
            strucID = WzString_trimmed(items(j))
            structIndex = getStructStatFromName(strucID)
            If Not (structIndex >= 0) Then ASSERT_FAILURE "structIndex >= 0", "src/research.cpp:619", "loadResearch", "Invalid item '%s' in list of redundant structures of research '%s' ", strucID, research_.name_
            If structIndex >= 0 Then
                PushLong research_.pRedStructs, research_.pRedStructs_size, structIndex And &HFFFF&
            End If
        Next

        ' asResearch.push_back(research)
        ReDim Preserve asResearch(0 To asResearch_size)
        asResearch(asResearch_size) = research_
        Set asResearch(asResearch_size).results = research_.results.clone()
        asResearch_size = asResearch_size + 1
        ini.endGroup
continue_:
    Next

    'Load and check research pre-requisites (need do it AFTER loading research items)
    For inc = 0 To asResearch_size - 1
        For j = 0 To preResearch(inc).size - 1
            resID = WzString_trimmed(preResearch(inc).items(j))
            preResItem = getResearch(resID)
            If Not (preResItem <> -1) Then ASSERT_FAILURE "preResItem != nullptr", "src/research.cpp:638", "loadResearch", "Invalid item '%s' in list of pre-requisites of research '%s' ", resID, asResearch(inc).name_
            If preResItem <> -1 Then
                PushLong asResearch(inc).pPRList, asResearch(inc).pPRList_size, asResearch(preResItem).index And &HFFFF&
            End If
        Next
    Next

    If CycleDetection_detectCycle(cycle, nCycle) Then
        debug_ 648, LOG_ERROR, "loadResearch", "A cycle was detected in the research dependency graph:"
        For k = 0 To nCycle - 1
            debug_ 651, LOG_ERROR, "loadResearch", "\t-> %s", asResearch(cycle(k)).id
        Next
        loadResearch = False
        Exit Function
    End If

    ' populate research category info
    ClearResCategories ' must clear because we re-process the entire asResearch list if loading more than one research file
    For inc = 0 To asResearch_size - 1
        cat = asResearch(inc).category
        If Len(cat) = 0 Then
            GoTo continue_cat
        End If
        If Not resCategories.Exists(cat) Then
            slot = resCategories.count
            ReDim Preserve resCategoryMembers(0 To slot)
            resCategories.Add cat, slot
        End If
        slot = resCategories(cat)
        PushLong resCategoryMembers(slot).members, resCategoryMembers(slot).size, inc
continue_cat:
    Next
    keys = resCategories.keys()
    For k = 0 To resCategories.count - 1
        cat = keys(k)
        slot = resCategories(cat)
        nMembers = resCategoryMembers(slot).size
        membersOfCategory = resCategoryMembers(slot).members

        If Not isResearchListValidConnectedDependencyGraph(membersOfCategory, nMembers) Then
            debug_ 673, LOG_ERROR, "loadResearch", "Research category items do not exist on a valid connected dependency graph: ""%s""", cat
            For j = 0 To nMembers - 1
                asResearch(membersOfCategory(j)).category = ""
            Next
            ' it = resCategories.erase(it)
            resCategories.Remove cat
            GoTo continue_erase
        End If

        ' std::stable_sort(..., isResAPrereqForResB): the library's insertion sort
        For j = 1 To nMembers - 1
            val_ = membersOfCategory(j)
            If isResAPrereqForResB(val_, membersOfCategory(0)) Then
                For hole = j To 1 Step -1
                    membersOfCategory(hole) = membersOfCategory(hole - 1)
                Next
                membersOfCategory(0) = val_
            Else
                hole = j
                Do While isResAPrereqForResB(val_, membersOfCategory(hole - 1))
                    membersOfCategory(hole) = membersOfCategory(hole - 1)
                    hole = hole - 1
                Loop
                membersOfCategory(hole) = val_
            End If
        Next
        resCategoryMembers(slot).members = membersOfCategory
        prog = 1
        categorySize = nMembers
        For j = 0 To nMembers - 1
            asResearch(membersOfCategory(j)).categoryProgress = prog And &HFFFF&
            asResearch(membersOfCategory(j)).categoryMax = categorySize And &HFFFF&
            prog = prog + 1
        Next
continue_erase:
    Next

    ' If the first research json file does not explicitly set calculationMode, default to compat
    If Not has_researchUpgradeCalcMode Then
        researchUpgradeCalcMode = ResearchUpgradeCalculationMode_Compat
        has_researchUpgradeCalcMode = True
    End If

    loadResearch = True
End Function

' (VB) researchAvailable's IsResearchCancelledFunc: IsResearchCancelledPending, or IsResearchCancelled in ModeImmediate
Private Function IsResearchCancelledFunc(ByVal mode As QUEUE_MODE, x As PLAYER_RESEARCH) As Boolean
    If mode = ModeImmediate Then
        IsResearchCancelledFunc = IsResearchCancelled(x)
    Else
        IsResearchCancelledFunc = IsResearchCancelledPending(x)
    End If
End Function

' (VB) researchAvailable's IsResearchStartedFunc: IsResearchStartedPending, or IsResearchStarted in ModeImmediate
Private Function IsResearchStartedFunc(ByVal mode As QUEUE_MODE, x As PLAYER_RESEARCH) As Boolean
    If mode = ModeImmediate Then
        IsResearchStartedFunc = IsResearchStarted(x)
    Else
        IsResearchStartedFunc = IsResearchStartedPending(x)
    End If
End Function

Public Function researchAvailable(ByVal inc As Long, ByVal playerID As Long, ByVal mode As QUEUE_MODE) As Boolean
    Dim incPR As Long, incS As Long, bPRFound As Boolean, bStructFound As Boolean, researchStarted As Boolean, psBuilding As Long

    If Not ULT(playerID, MAX_PLAYERS) Then
        researchAvailable = False
        Exit Function
    End If

    ' Decide whether to use IsResearchCancelledPending/IsResearchStartedPending or IsResearchCancelled/IsResearchStarted.
    ' (VB) IsResearchCancelledFunc and IsResearchStartedFunc decide by the mode

    ' if its a cancelled topic - add to list
    If IsResearchCancelledFunc(mode, asPlayerResList(playerID, inc)) Then
        researchAvailable = True
        Exit Function
    End If
    ' Ignore disabled
    If IsResearchDisabled(asPlayerResList(playerID, inc)) Then
        researchAvailable = False
        Exit Function
    End If
    ' if the topic is possible and has not already been researched - add to list
    If IsResearchPossible(asPlayerResList(playerID, inc)) Then
        If Not IsResearchCompleted(asPlayerResList(playerID, inc)) _
           And Not IsResearchStartedFunc(mode, asPlayerResList(playerID, inc)) Then
            researchAvailable = True
            Exit Function
        End If
    End If

    ' if single player mode and key topic, then ignore cos can't do it!
    If Not bMultiPlayer And asResearch(inc).keyTopic <> 0 Then
        researchAvailable = False
        Exit Function
    End If

    researchStarted = IsResearchStartedFunc(mode, asPlayerResList(playerID, inc))
    If researchStarted Then
        psBuilding = findResearchingFacilityByResearchIndex(playerID, inc)  ' May fail to find the structure here, if the research is merely pending, not actually started.
        If psBuilding <> 0 Then
            If aStruct(psBuilding).status = SS_BEING_BUILT Then
                researchStarted = False  ' Although research is started, the facility is currently being upgraded or demolished, so we want to be able to research this elsewhere.
            End If
        End If
    End If

    ' make sure that the research is not completed  or started by another researchfac
    If Not IsResearchCompleted(asPlayerResList(playerID, inc)) And Not researchStarted Then
        ' Research is not completed  ... also  it has not been started by another researchfac

        ' if there aren't any PR's - go to next topic
        If asResearch(inc).pPRList_size = 0 Then
            researchAvailable = False
            Exit Function
        End If

        ' check for pre-requisites
        bPRFound = True
        For incPR = 0 To asResearch(inc).pPRList_size - 1
            If Not IsResearchCompleted(asPlayerResList(playerID, asResearch(inc).pPRList(incPR))) Then
                ' if haven't pre-requisite - quit checking rest
                bPRFound = False
                Exit For
            End If
        Next
        If Not bPRFound Then
            ' if haven't pre-requisites, skip the rest of the checks
            researchAvailable = False
            Exit Function
        End If

        ' check for structure effects
        bStructFound = True
        For incS = 0 To asResearch(inc).pStructList_size - 1
            If Not checkSpecificStructExists(asResearch(inc).pStructList(incS), playerID) Then
                'if not built, quit checking
                bStructFound = False
                Exit For
            End If
        Next
        If Not bStructFound Then
            ' if haven't all structs built, skip to next topic
            researchAvailable = False
            Exit Function
        End If
        researchAvailable = True
        Exit Function
    End If
    researchAvailable = False
End Function

' Function to check what can be researched for a particular player at any one
' instant (see research.cpp): the topics in list(0 ..), and their count.
' NOTE by AJL may 99 - skirmish now has it's own version of this, skTopicAvail.
Public Function fillResearchList(ByVal playerID As Long, ByVal has_topic As Boolean, ByVal topic As Long, ByVal limit As Long, list() As Long) As Long
    Dim inc As Long, n As Long, add As Boolean

    For inc = 0 To asResearch_size - 1
        ' if the inc matches the 'topic' - automatically add to the list
        add = False
        If has_topic Then add = (inc = topic)
        If Not add Then add = researchAvailable(inc, playerID, ModeQueue)
        If add Then
            ReDim Preserve list(0 To n)
            list(n) = inc
            n = n + 1
            If n = limit Then
                fillResearchList = n
                Exit Function
            End If
        End If
    Next

    fillResearchList = n
End Function

' (int64_t: Double)
Private Function iDivCeil(ByVal dividend As Double, ByVal divisor As Double) As Double
    Dim hasPosQuotient As Boolean, q As Double
    If Not (divisor <> 0) Then ASSERT_FAILURE "divisor != 0", "src/research.cpp:859", "iDivCeil", "Divide by 0": iDivCeil = 0: Exit Function
    hasPosQuotient = (dividend >= 0) = (divisor >= 0)
    ' C++11 defines the behavior of % to be truncated
    q = Fix(dividend / divisor)
    If dividend - q * divisor <> 0 And hasPosQuotient Then q = q + 1
    iDivCeil = q
End Function

Private Function iDivFloor(ByVal dividend As Double, ByVal divisor As Double) As Double
    Dim hasNegQuotient As Boolean, q As Double
    If Not (divisor <> 0) Then ASSERT_FAILURE "divisor != 0", "src/research.cpp:867", "iDivFloor", "Divide by 0": iDivFloor = 0: Exit Function
    hasNegQuotient = (dividend >= 0) <> (divisor >= 0)
    ' C++11 defines the behavior of % to be truncated
    q = Fix(dividend / divisor)
    If dividend - q * divisor <> 0 And hasNegQuotient Then q = q - 1
    iDivFloor = q
End Function

' (VB) a std::unordered_map<std::string, uint32_t>'s operator[](key)++
Private Sub IncrementCount(ByVal m As Object, ByVal key As String)
    If m.Exists(key) Then
        m.Item(key) = U32Add1(m.Item(key))
    Else
        m.Add key, 1&
    End If
End Sub

' (VB) a uint32_t's ++
Private Function U32Add1(ByVal v As Long) As Long
    If v = -1 Then
        U32Add1 = 0
    ElseIf v = &H7FFFFFFF Then
        U32Add1 = &H80000000
    Else
        U32Add1 = v + 1
    End If
End Function

Private Sub eventResearchedHandleUpgrades(ByVal psResearch As Long, ByVal psStruct As Long, ByVal player As Long)
    Dim upgradeCalcMode As ResearchUpgradeCalculationMode, playerRawUpgradeChangeTotals As Object
    Dim tempStats As PLAYER_UPGRADE_COUNTS, i As Long, k As Long, x As Long, v As nlohmann_json
    Dim it_ctype As nlohmann_json, it_parameter As nlohmann_json, it_value As nlohmann_json
    Dim ctype As String, parameter As String, value As Double, it_filterparam As nlohmann_json, it_filtervalue As nlohmann_json
    Dim filterparamStr As String, filtervalStr As String, pPlayerEntityClass As Long, isBodyClass As Boolean, isWeaponClass As Boolean
    Dim cname As GameEntityRulesPair, statsEntityClassObj As nlohmann_json, statsEntityObj As nlohmann_json
    Dim filterparam As String, pStatsFilterParameterValue As nlohmann_json, pStatsParameterValue As nlohmann_json
    Dim dst As nlohmann_json, statsOriginalValueX As nlohmann_json, currentUpgradesValue_json As nlohmann_json
    Dim currentUpgradesValue As Double, scaledChange As Double, newUpgradesChange As Double, newUpgradesValue As Double
    Dim statsOriginalValue As Double, cumulativeUpgradeScaledChange As Double, rawKey As String
    Dim impactClass As nlohmann_json, bodyClass_ As nlohmann_json, keys As Variant, clsName As String

    If cachedStatsObject Is Nothing Then Set cachedStatsObject = wzapi_constructStatsObject()
    If cachedPerPlayerUpgrades_size = 0 Then cachedPerPlayerUpgrades_size = wzapi_getUpgradesObject(cachedPerPlayerUpgrades)
    ' internal_execution_context_base temp_no_throw_context: (VB) wzapi_c's procedures take no context

    debug_ 879, LOG_RESEARCH, "eventResearchedHandleUpgrades", "RESEARCH : %s(%s) for %d", asResearch(psResearch).name_, asResearch(psResearch).id, player

    If Not (player >= 0 And player < cachedPerPlayerUpgrades_size) Then ASSERT_FAILURE "player >= 0 && player < cachedPerPlayerUpgrades.size()", "src/research.cpp:881", "eventResearchedHandleUpgrades", "Player %d does not exist in per-player upgrades?", player: Exit Sub
    If cachedPerPlayerRawUpgradeChange(player) Is Nothing Then Set cachedPerPlayerRawUpgradeChange(player) = NewStringMap()
    Set playerRawUpgradeChangeTotals = cachedPerPlayerRawUpgradeChange(player)
    upgradeCalcMode = getResearchUpgradeCalcMode()

    Set tempStats.numBodyClassArmourUpgrades = NewStringMap()
    Set tempStats.numBodyClassThermalUpgrades = NewStringMap()
    Set tempStats.numWeaponImpactClassUpgrades = NewStringMap()

    ' iterate over all research results
    For i = 0 To asResearch(psResearch).results.size() - 1
        Set v = asResearch(psResearch).results.at_idx(i)
        ' Required members of research upgrades: "class", "parameter", "value"
        ' (VB) RS_GET_REQUIRED_RESULT_PROPERTY(it_ctype, "class", is_string)
        Set it_ctype = v.find("class")
        If it_ctype Is Nothing Then
            ASSERT_FAILURE "false", "src/research.cpp:904", "eventResearchedHandleUpgrades", "Research(""%s"").results[%zu]: Missing required parameter: ""%s""", asResearch(psResearch).id, i, "class"
            GoTo continue_result
        End If
        If Not it_ctype.is_string() Then
            ASSERT_FAILURE "false", "src/research.cpp:904", "eventResearchedHandleUpgrades", "Research(""%s"").results[%zu][""%s""]: Unexpected value type: ""%s""", asResearch(psResearch).id, i, "class", it_ctype.type_name()
            GoTo continue_result
        End If
        ' RS_GET_REQUIRED_RESULT_PROPERTY(it_parameter, "parameter", is_string)
        Set it_parameter = v.find("parameter")
        If it_parameter Is Nothing Then
            ASSERT_FAILURE "false", "src/research.cpp:905", "eventResearchedHandleUpgrades", "Research(""%s"").results[%zu]: Missing required parameter: ""%s""", asResearch(psResearch).id, i, "parameter"
            GoTo continue_result
        End If
        If Not it_parameter.is_string() Then
            ASSERT_FAILURE "false", "src/research.cpp:905", "eventResearchedHandleUpgrades", "Research(""%s"").results[%zu][""%s""]: Unexpected value type: ""%s""", asResearch(psResearch).id, i, "parameter", it_parameter.type_name()
            GoTo continue_result
        End If
        ' RS_GET_REQUIRED_RESULT_PROPERTY(it_value, "value", is_number_integer)
        Set it_value = v.find("value")
        If it_value Is Nothing Then
            ASSERT_FAILURE "false", "src/research.cpp:906", "eventResearchedHandleUpgrades", "Research(""%s"").results[%zu]: Missing required parameter: ""%s""", asResearch(psResearch).id, i, "value"
            GoTo continue_result
        End If
        If Not it_value.is_number_integer() Then
            ASSERT_FAILURE "false", "src/research.cpp:906", "eventResearchedHandleUpgrades", "Research(""%s"").results[%zu][""%s""]: Unexpected value type: ""%s""", asResearch(psResearch).id, i, "value", it_value.type_name()
            GoTo continue_result
        End If
        ctype = it_ctype.get_string()
        parameter = it_parameter.get_string()
        value = it_value.get_int64()
        Set it_filterparam = v.find("filterParameter") ' optional
        Set it_filtervalue = v.find("filterValue") ' required if "filterParameter" is specified
        If Not it_filterparam Is Nothing Then
            If Not it_filterparam.is_string() Then
                ASSERT_FAILURE "false", "src/research.cpp:916", "eventResearchedHandleUpgrades", "Research(""%s"").results[%zu][""%s""]: Unexpected value type: ""%s""", asResearch(psResearch).id, i, "filterParameter", it_parameter.type_name()
                GoTo continue_result
            End If
            If it_filtervalue Is Nothing Then
                ' ERROR: Supplied a filterParameter but not a filterValue
                ASSERT_FAILURE "false", "src/research.cpp:922", "eventResearchedHandleUpgrades", "Research(""%s"").results[%zu]: Missing ""%s"" property (required when ""filterParameter"" is specified)", asResearch(psResearch).id, i, "filterParameter"
                GoTo continue_result
            End If
        End If
        filterparamStr = ""
        If Not it_filterparam Is Nothing Then filterparamStr = it_filterparam.get_string()
        filtervalStr = ""
        If Not it_filtervalue Is Nothing Then filtervalStr = it_filtervalue.dump()
        debug_ 926, LOG_RESEARCH, "eventResearchedHandleUpgrades", "    RESULT : class=""%s"" parameter=""%s"" value=%i filter=""%s"" filterval=%s", ctype, parameter, value, filterparamStr, filtervalStr

        pPlayerEntityClass = PerPlayerUpgrades_find(cachedPerPlayerUpgrades(player), ctype)
        If pPlayerEntityClass = -1 Then
            ASSERT_FAILURE "pPlayerEntityClass", "src/research.cpp:931", "eventResearchedHandleUpgrades", "Unknown entity class: %s", ctype
            GoTo continue_result
        End If
        isBodyClass = ctype = "Body"
        isWeaponClass = ctype = "Weapon"
        For k = 0 To cachedPerPlayerUpgrades(player).upgrades_values(pPlayerEntityClass).rules_size - 1 ' iterate over all components of this type
            cname = cachedPerPlayerUpgrades(player).upgrades_values(pPlayerEntityClass).rules(k)
            Set statsEntityClassObj = cachedStatsObject.find(ctype)
            If statsEntityClassObj Is Nothing Then
                ASSERT_FAILURE "false", "src/research.cpp:941", "eventResearchedHandleUpgrades", "Parameter ""%s"" does not exist in Stats[%s][%s] ?", parameter, ctype, cname.first
                GoTo continue_cname
            End If
            Set statsEntityObj = statsEntityClassObj.find(cname.first)
            If statsEntityObj Is Nothing Then
                ASSERT_FAILURE "false", "src/research.cpp:947", "eventResearchedHandleUpgrades", "Parameter ""%s"" does not exist in Stats[%s][%s] ?", parameter, ctype, cname.first
                GoTo continue_cname
            End If

            If Not it_filterparam Is Nothing Then
                ' more specific filter
                filterparam = it_filterparam.get_string()
                Set pStatsFilterParameterValue = statsEntityObj.find(filterparam)
                If pStatsFilterParameterValue Is Nothing Then
                    ' Did not find filter parameter
                    GoTo continue_cname
                End If
                If Not pStatsFilterParameterValue.equals(it_filtervalue) Then
                    ' Non-matching filter parameter
                    GoTo continue_cname
                End If
            End If

            Set pStatsParameterValue = statsEntityObj.find(parameter)
            If pStatsParameterValue Is Nothing Then
                ' Did not find it??
                ASSERT_FAILURE "false", "src/research.cpp:972", "eventResearchedHandleUpgrades", "Parameter ""%s"" does not exist in Stats[%s][%s] ?", parameter, ctype, cname.first
                GoTo continue_cname
            End If

            If pStatsParameterValue.is_array() Then ' (ex. modifying "RankThresholds")
                Set dst = GameEntityRules_getPropertyValue(cname.second, parameter)
                If Not dst.is_array() Then
                    ' The Upgrades parameter unexpectedly is not an array, or not the same array size
                    ASSERT_FAILURE "false", "src/research.cpp:982", "eventResearchedHandleUpgrades", "Upgrades parameter ""%s"" value (type %s) does not match Stats[%s][%s] value type (%s) or size (%zu)", parameter, dst.type_name(), ctype, cname.first, pStatsParameterValue.type_name(), pStatsParameterValue.size()
                    GoTo continue_cname
                ElseIf dst.size() <> pStatsParameterValue.size() Then
                    ASSERT_FAILURE "false", "src/research.cpp:982", "eventResearchedHandleUpgrades", "Upgrades parameter ""%s"" value (type %s) does not match Stats[%s][%s] value type (%s) or size (%zu)", parameter, dst.type_name(), ctype, cname.first, pStatsParameterValue.type_name(), pStatsParameterValue.size()
                    GoTo continue_cname
                End If
                For x = 0 To dst.size() - 1
                    Set statsOriginalValueX = pStatsParameterValue.at_idx(x)
                    If Not statsOriginalValueX.is_number_integer() Then
                        ASSERT_FAILURE "false", "src/research.cpp:990", "eventResearchedHandleUpgrades", "Unexpected parameter ""%s[%zu]"" value type (%s) in Stats[%s][%s]", parameter, x, pStatsParameterValue.type_name(), ctype, cname.first
                        GoTo continue_x
                    End If
                    Set currentUpgradesValue_json = dst.item_idx(x)
                    If Not currentUpgradesValue_json.is_number_integer() Then
                        ' The Upgrades parameter unexpectedly is not an integer
                        ASSERT_FAILURE "false", "src/research.cpp:997", "eventResearchedHandleUpgrades", "Upgrades parameter ""%s[%zu]"" value type (%s) does not match Stats[%s][%s] value type (%s)", parameter, x, currentUpgradesValue_json.type_name(), ctype, cname.first, pStatsParameterValue.type_name()
                        GoTo continue_x
                    End If
                    currentUpgradesValue = currentUpgradesValue_json.get_int64()
                    scaledChange = (statsOriginalValueX.get_int64() * value)
                    If value < 0 Then
                        newUpgradesChange = iDivFloor(scaledChange, 100)
                    Else
                        newUpgradesChange = iDivCeil(scaledChange, 100)
                    End If
                    newUpgradesValue = (currentUpgradesValue + newUpgradesChange)
                    If currentUpgradesValue_json.is_number_unsigned() Then
                        ' original was unsigned integer - round anything less than 0 up to 0
                        If newUpgradesValue < 0 Then newUpgradesValue = 0
                        dst.set_item_idx x, json_value_unsigned(newUpgradesValue)
                    Else
                        dst.set_item_idx x, json_value(CDec(newUpgradesValue))
                    End If
continue_x:
                Next
                GameEntityRules_setPropertyValue cname.second, parameter, dst
                debug_ 1016, LOG_RESEARCH, "eventResearchedHandleUpgrades", "    upgraded to : %s", dst.dump()
            ElseIf pStatsParameterValue.is_number_integer() Then
                statsOriginalValue = pStatsParameterValue.get_int64()
                If statsOriginalValue <= 0 Then ' only applies if stat has above zero value already
                    GoTo continue_cname
                End If
                Set currentUpgradesValue_json = GameEntityRules_getPropertyValue(cname.second, parameter)
                If Not currentUpgradesValue_json.is_number_integer() Then
                    ' The Upgrades parameter unexpectedly is not an integer
                    ASSERT_FAILURE "false", "src/research.cpp:1029", "eventResearchedHandleUpgrades", "Upgrades parameter ""%s"" value type (%s) does not match Stats[%s][%s] value type (%s)", parameter, currentUpgradesValue_json.type_name(), ctype, cname.first, pStatsParameterValue.type_name()
                    GoTo continue_cname
                End If
                currentUpgradesValue = currentUpgradesValue_json.get_int64()
                scaledChange = (statsOriginalValue * value)
                newUpgradesChange = 0
                newUpgradesValue = 0
                Select Case upgradeCalcMode
                Case ResearchUpgradeCalculationMode_Compat
                    ' Default / compat cumulative upgrade handling (the only option for many years - from at least 3.x/(3.2+?)-4.4.2?)
                    ' This can accumulate noticeable error, especially if repeatedly upgrading small values by small percentages (commonly impacted: armour, thermal)
                    ' However, research.json created and tested during this long period may be expecting this outcome / behavior
                    If value < 0 Then
                        newUpgradesChange = iDivFloor(scaledChange, 100)
                    Else
                        newUpgradesChange = iDivCeil(scaledChange, 100)
                    End If
                    newUpgradesValue = (currentUpgradesValue + newUpgradesChange)
                Case ResearchUpgradeCalculationMode_Improved
                    ' "Improved" cumulative upgrade handling (significantly reduces accumulated errors)
                    rawKey = cname.first & vbNullChar & parameter
                    If playerRawUpgradeChangeTotals.Exists(rawKey) Then
                        cumulativeUpgradeScaledChange = playerRawUpgradeChangeTotals.Item(rawKey)
                    Else
                        cumulativeUpgradeScaledChange = 0
                    End If
                    cumulativeUpgradeScaledChange = cumulativeUpgradeScaledChange + scaledChange
                    playerRawUpgradeChangeTotals.Item(rawKey) = cumulativeUpgradeScaledChange
                    If cumulativeUpgradeScaledChange < 0 Then
                        newUpgradesValue = statsOriginalValue + iDivFloor(cumulativeUpgradeScaledChange, 100)
                    Else
                        newUpgradesValue = statsOriginalValue + iDivCeil(cumulativeUpgradeScaledChange, 100)
                    End If
                    newUpgradesChange = newUpgradesValue - currentUpgradesValue
                End Select
                If currentUpgradesValue_json.is_number_unsigned() Then
                    ' original was unsigned integer - round anything less than 0 up to 0
                    If newUpgradesValue < 0 Then newUpgradesValue = 0
                    GameEntityRules_setPropertyValue cname.second, parameter, json_value_unsigned(newUpgradesValue)
                Else
                    GameEntityRules_setPropertyValue cname.second, parameter, json_value(CDec(newUpgradesValue))
                End If
                debug_ 1066, LOG_RESEARCH, "eventResearchedHandleUpgrades", "      upgraded ""%s"" to %i by %i", cname.first, newUpgradesValue, newUpgradesChange
                If isWeaponClass Then
                    Set impactClass = statsEntityObj.find("ImpactClass")
                    If Not impactClass Is Nothing Then
                        IncrementCount tempStats.numWeaponImpactClassUpgrades, impactClass.get_string()
                    Else
                        ASSERT_FAILURE "false", "src/research.cpp:1076", "eventResearchedHandleUpgrades", "Did not find expected ""ImpactClass"" member in Stats[%s][%s]", ctype, cname.first
                    End If
                ElseIf isBodyClass And parameter = "Armour" Then
                    Set bodyClass_ = statsEntityObj.find("BodyClass")
                    If Not bodyClass_ Is Nothing Then
                        IncrementCount tempStats.numBodyClassArmourUpgrades, bodyClass_.get_string()
                    Else
                        ASSERT_FAILURE "false", "src/research.cpp:1088", "eventResearchedHandleUpgrades", "Did not find expected ""BodyClass"" member in Stats[%s][%s]", ctype, cname.first
                    End If
                ElseIf isBodyClass And parameter = "Thermal" Then
                    Set bodyClass_ = statsEntityObj.find("BodyClass")
                    If Not bodyClass_ Is Nothing Then
                        IncrementCount tempStats.numBodyClassThermalUpgrades, bodyClass_.get_string()
                    Else
                        ASSERT_FAILURE "false", "src/research.cpp:1100", "eventResearchedHandleUpgrades", "Did not find expected ""BodyClass"" member in Stats[%s][%s]", ctype, cname.first
                    End If
                End If
            Else
                ' unexpected type
                ' Research stats / upgrades are not supposed to expose non-integer types, as the core stats / game state calculations must use integer arithmetic
                ASSERT_FAILURE "false", "src/research.cpp:1108", "eventResearchedHandleUpgrades", "Unexpected parameter ""%s"" value type (%s) in Stats[%s][%s]", parameter, pStatsParameterValue.type_name(), ctype, cname.first
                GoTo continue_cname
            End If
continue_cname:
        Next
continue_result:
    Next

    ' accumulate stats
    keys = tempStats.numBodyClassArmourUpgrades.keys()
    For k = 0 To tempStats.numBodyClassArmourUpgrades.count - 1
        clsName = keys(k)
        If tempStats.numBodyClassArmourUpgrades.Item(clsName) <> 0 Then
            IncrementCount playerUpgradeCounts(player).numBodyClassArmourUpgrades, clsName
            debug_ 1121, LOG_RESEARCH, "eventResearchedHandleUpgrades", "  Player[%d], Armour[%s] grade: %u", player, clsName, playerUpgradeCounts(player).numBodyClassArmourUpgrades.Item(clsName)
        End If
    Next
    keys = tempStats.numBodyClassThermalUpgrades.keys()
    For k = 0 To tempStats.numBodyClassThermalUpgrades.count - 1
        clsName = keys(k)
        If tempStats.numBodyClassThermalUpgrades.Item(clsName) <> 0 Then
            IncrementCount playerUpgradeCounts(player).numBodyClassThermalUpgrades, clsName
            debug_ 1130, LOG_RESEARCH, "eventResearchedHandleUpgrades", "  Player[%d], Thermal[%s] grade: %u", player, clsName, playerUpgradeCounts(player).numBodyClassThermalUpgrades.Item(clsName)
        End If
    Next
    keys = tempStats.numWeaponImpactClassUpgrades.keys()
    For k = 0 To tempStats.numWeaponImpactClassUpgrades.count - 1
        clsName = keys(k)
        If tempStats.numWeaponImpactClassUpgrades.Item(clsName) <> 0 Then
            IncrementCount playerUpgradeCounts(player).numWeaponImpactClassUpgrades, clsName
            debug_ 1139, LOG_RESEARCH, "eventResearchedHandleUpgrades", "  Player[%d], Weapon[%s] grade: %u", player, clsName, playerUpgradeCounts(player).numWeaponImpactClassUpgrades.Item(clsName)
        End If
    Next
End Sub

Private Sub makeComponentRedundant(state As Byte)
    Select Case state
    Case AVAILABLE
        state = REDUNDANT
    Case UNAVAILABLE
        state = REDUNDANT_UNAVAILABLE
    Case FOUND
        state = REDUNDANT_FOUND
    End Select
End Sub

Private Sub makeComponentAvailable(state As Byte)
    Select Case state
    Case UNAVAILABLE, FOUND
        state = AVAILABLE
    Case REDUNDANT_UNAVAILABLE, REDUNDANT_FOUND
        state = REDUNDANT
    End Select
End Sub

' process the results of a completed research topic
Public Sub researchResult(ByVal researchIndex As Long, ByVal player As Long, ByVal bDisplay As Boolean, ByVal psResearchFacility As Long, ByVal bTrigger As Boolean)
    Dim pResearch As Long, pMessage As Long, consoleMsg As String, k As Long, pOldComp As Long
    Dim typ As COMPONENT_TYPE, compInc As Long, componentResult As Long, pRedArtefact As Long

    If Not (ULT(researchIndex, asResearch_size)) Then ASSERT_FAILURE "researchIndex < asResearch.size()", "src/research.cpp:1178", "researchResult", "Invalid research index %u", researchIndex: Exit Sub
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/research.cpp:1179", "researchResult", "invalid player: %u", player: Exit Sub

    pResearch = researchIndex
    'the message gets sent to console

    syncDebug "researchResult", "researchResult(%u, %u, " & ChrW$(&H2026) & ")", researchIndex, player

    MakeResearchCompleted asPlayerResList(player, researchIndex)

    'check for structures to be made available
    For k = 0 To asResearch(pResearch).pStructureResults_size - 1
        makeComponentAvailable apStructTypeLists(player, asResearch(pResearch).pStructureResults(k))
    Next

    'check for structures to be made redundant
    For k = 0 To asResearch(pResearch).pRedStructs_size - 1
        makeComponentRedundant apStructTypeLists(player, asResearch(pResearch).pRedStructs(k))
    Next

    'check for component replacement
    If asResearch(pResearch).componentReplacement_size <> 0 Then
        For k = 0 To asResearch(pResearch).componentReplacement_size - 1
            pOldComp = asResearch(pResearch).componentReplacement(k).pOldComponent
            replaceComponent asResearch(pResearch).componentReplacement(k).pNewComponent, pOldComp, player
            makeComponentRedundant apCompLists(player, COMPONENT_STATS_compType(pOldComp), BASE_STATS_index(pOldComp))
        Next
    End If

    'check for artefacts to be made available
    For k = 0 To asResearch(pResearch).componentResults_size - 1
        componentResult = asResearch(pResearch).componentResults(k)
        'determine the type of artefact
        typ = COMPONENT_STATS_compType(componentResult)
        'set the component state to AVAILABLE
        compInc = BASE_STATS_index(componentResult)
        makeComponentAvailable apCompLists(player, typ, compInc)
        ' (VB) C++'s && does not index the other types' arrays
        Select Case typ
        Case COMP_SENSOR
            'check for default sensor
            If asSensorStats(compInc).location = LOC_DEFAULT Then
                aDefaultSensor(player) = compInc
            End If
        Case COMP_ECM
            'check for default ECM
            If asECMStats(compInc).location = LOC_DEFAULT Then
                aDefaultECM(player) = compInc
            End If
        Case COMP_REPAIRUNIT
            'check for default Repair
            If asRepairStats(compInc).location = LOC_DEFAULT Then
                aDefaultRepair(player) = compInc
                enableSelfRepair player
            End If
        End Select
    Next

    'check for artefacts to be made redundant
    For k = 0 To asResearch(pResearch).pRedArtefacts_size - 1
        pRedArtefact = asResearch(pResearch).pRedArtefacts(k)
        typ = COMPONENT_STATS_compType(pRedArtefact)
        makeComponentRedundant apCompLists(player, typ, BASE_STATS_index(pRedArtefact))
    Next

    'Add message to player's list if Major Topic
    If (asResearch(pResearch).techCode = TC_MAJOR) And bDisplay Then
        'only play sound if major topic
        If player = selectedPlayer Then
            audio_QueueTrack ID_SOUND_MAJOR_RESEARCH
        End If

        'check there is viewdata for the research topic - just don't add message if not!
        If asResearch(pResearch).pViewData <> 0 Then
            pMessage = addMessage(MSG_RESEARCH, False, player)
            If pMessage <> 0 Then
                MESSAGE_pViewData(pMessage) = asResearch(pResearch).pViewData
                jsDebugMessageUpdate
            End If
        End If
    ElseIf player = selectedPlayer And bDisplay Then
        audio_QueueTrack ID_SOUND_RESEARCH_COMPLETED
    End If

    If player = selectedPlayer And bDisplay Then
        'add console text message
        consoleMsg = Left$(va("Research completed: %s", asResearch(pResearch).name_), MAX_RESEARCH_MSG_SIZE - 1)
        addConsoleMessage consoleMsg, LEFT_JUSTIFY, SYSTEM_MESSAGE
    End If

    If psResearchFacility <> 0 Then
        aStruct(psResearchFacility).functionality.researchFacility.psSubject = -1      ' Make sure topic is cleared
    End If

    eventResearchedHandleUpgrades pResearch, psResearchFacility, player

    triggerEventResearched pResearch, psResearchFacility, player
End Sub

'This function is called when the research files are reloaded
Public Function ResearchShutDown() As Boolean
    ResearchRelease
    ResearchShutDown = True
End Function

'This function is called when a game finishes
Public Sub ResearchRelease()
    Dim i As Long
    Erase asResearch
    asResearch_size = 0
    has_researchUpgradeCalcMode = False
    ClearResCategories
    For i = 0 To MAX_PLAYERS - 1
        asPlayerResList_size(i) = 0
    Next
    Erase asPlayerResList
    asPlayerResListCap = 0
    Set cachedStatsObject = Nothing
    Erase cachedPerPlayerUpgrades
    cachedPerPlayerUpgrades_size = 0
    For i = 0 To MAX_PLAYERS - 1
        Set cachedPerPlayerRawUpgradeChange(i) = Nothing
    Next
    ResetPlayerUpgradeCounts
End Sub

'set the iconID based on the name read in in the stats
Private Function setIconID(ByVal pIconName As String, ByVal pName As String) As Long
    'compare the names with those created in 'Framer'
    Select Case pIconName
    Case "IMAGE_ROCKET": setIconID = IMAGE_ROCKET
    Case "IMAGE_CANNON": setIconID = IMAGE_CANNON
    Case "IMAGE_HOVERCRAFT": setIconID = IMAGE_HOVERCRAFT
    Case "IMAGE_ECM": setIconID = IMAGE_ECM
    Case "IMAGE_PLASCRETE": setIconID = IMAGE_PLASCRETE
    Case "IMAGE_TRACKS": setIconID = IMAGE_TRACKS
    Case "IMAGE_RES_DROIDTECH": setIconID = IMAGE_RES_DROIDTECH
    Case "IMAGE_RES_WEAPONTECH": setIconID = IMAGE_RES_WEAPONTECH
    Case "IMAGE_RES_COMPUTERTECH": setIconID = IMAGE_RES_COMPUTERTECH
    Case "IMAGE_RES_POWERTECH": setIconID = IMAGE_RES_POWERTECH
    Case "IMAGE_RES_SYSTEMTECH": setIconID = IMAGE_RES_SYSTEMTECH
    Case "IMAGE_RES_STRUCTURETECH": setIconID = IMAGE_RES_STRUCTURETECH
    Case "IMAGE_RES_CYBORGTECH": setIconID = IMAGE_RES_CYBORGTECH
    Case "IMAGE_RES_DEFENCE": setIconID = IMAGE_RES_DEFENCE
    Case "IMAGE_RES_QUESTIONMARK": setIconID = IMAGE_RES_QUESTIONMARK
    Case "IMAGE_RES_GRPACC": setIconID = IMAGE_RES_GRPACC
    Case "IMAGE_RES_GRPUPG": setIconID = IMAGE_RES_GRPUPG
    Case "IMAGE_RES_GRPREP": setIconID = IMAGE_RES_GRPREP
    Case "IMAGE_RES_GRPROF": setIconID = IMAGE_RES_GRPROF
    Case "IMAGE_RES_GRPDAM": setIconID = IMAGE_RES_GRPDAM
    Case Else
        ' Add more names as images are created
        ASSERT_FAILURE "false", "src/research.cpp:1551", "setIconID", "Invalid icon graphic %s for topic %s", pIconName, pName

        setIconID = NO_RESEARCH_ICON ' Should never get here.
    End Select
End Function

Public Function mapIconToRID(ByVal iconID As Long) As Long
    Select Case iconID
    Case IMAGE_ROCKET
        mapIconToRID = RID_ROCKET
    Case IMAGE_CANNON
        mapIconToRID = RID_CANNON
    Case IMAGE_HOVERCRAFT
        mapIconToRID = RID_HOVERCRAFT
    Case IMAGE_ECM
        mapIconToRID = RID_ECM
    Case IMAGE_PLASCRETE
        mapIconToRID = RID_PLASCRETE
    Case IMAGE_TRACKS
        mapIconToRID = RID_TRACKS
    Case IMAGE_RES_DROIDTECH
        mapIconToRID = RID_DROIDTECH
    Case IMAGE_RES_WEAPONTECH
        mapIconToRID = RID_WEAPONTECH
    Case IMAGE_RES_COMPUTERTECH
        mapIconToRID = RID_COMPUTERTECH
    Case IMAGE_RES_POWERTECH
        mapIconToRID = RID_POWERTECH
    Case IMAGE_RES_SYSTEMTECH
        mapIconToRID = RID_SYSTEMTECH
    Case IMAGE_RES_STRUCTURETECH
        mapIconToRID = RID_STRUCTURETECH
    Case IMAGE_RES_CYBORGTECH
        mapIconToRID = RID_CYBORGTECH
    Case IMAGE_RES_DEFENCE
        mapIconToRID = RID_DEFENCE
    Case IMAGE_RES_QUESTIONMARK
        mapIconToRID = RID_QUESTIONMARK
    Case IMAGE_RES_GRPACC
        mapIconToRID = RID_GRPACC
    Case IMAGE_RES_GRPUPG
        mapIconToRID = RID_GRPUPG
    Case IMAGE_RES_GRPREP
        mapIconToRID = RID_GRPREP
    Case IMAGE_RES_GRPROF
        mapIconToRID = RID_GRPROF
    Case IMAGE_RES_GRPDAM
        mapIconToRID = RID_GRPDAM
    Case Else
        mapIconToRID = -1 'pass back a value that can never have been set up
    End Select
End Function

'return a pointer to a research topic based on the name (its index, or -1)
Public Function getResearch(ByVal pName As String) As Long
    Dim inc As Long
    For inc = 0 To asResearch_size - 1
        If asResearch(inc).id = pName Then
            getResearch = inc
            Exit Function
        End If
    Next
    debug_ 1636, LOG_WARNING, "getResearch", "Unknown research - %s", pName
    getResearch = -1
End Function

' (VB) replaceComponent's lambda replaceComponentInTemplate, with its captures
Private Function replaceComponentInTemplate(ByVal psTemplates As Long, ByVal oldType As COMPONENT_TYPE, ByVal oldCompInc As Long, ByVal newCompInc As Long) As Boolean
    Dim inc As Long
    Select Case oldType
    Case COMP_BODY, COMP_BRAIN, COMP_PROPULSION, COMP_REPAIRUNIT, COMP_ECM, COMP_SENSOR, COMP_CONSTRUCT
        If apsTemplates(psTemplates).asParts(oldType) = oldCompInc Then
            apsTemplates(psTemplates).asParts(oldType) = newCompInc And &HFF&
        End If
    Case COMP_WEAPON
        For inc = 0 To apsTemplates(psTemplates).numWeaps - 1
            If apsTemplates(psTemplates).asWeaps(inc) = oldCompInc Then
                apsTemplates(psTemplates).asWeaps(inc) = newCompInc
            End If
        Next
    Case Else
        'unknown comp type
        debug_ 1687, LOG_ERROR, "replaceComponent", "Unknown component type - invalid Template"
        replaceComponentInTemplate = True
        Exit Function
    End Select
    replaceComponentInTemplate = True
End Function

' looks through the players lists of structures and droids to see if any are using
' the old component - if any then replaces them with the new component
Private Sub replaceComponent(ByVal pNewComponent As Long, ByVal pOldComponent As Long, ByVal player As Long)
    Dim oldType As COMPONENT_TYPE, oldCompInc As Long, newType As COMPONENT_TYPE, newCompInc As Long
    Dim i As Long, it As Long, psCBuilding As Long

    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/research.cpp:1645", "replaceComponent", "invalid player: %u", player: Exit Sub

    oldType = COMPONENT_STATS_compType(pOldComponent)
    oldCompInc = BASE_STATS_index(pOldComponent)
    newType = COMPONENT_STATS_compType(pNewComponent)
    newCompInc = BASE_STATS_index(pNewComponent)

    'check old and new type are the same
    If oldType <> newType Then
        Exit Sub
    End If

    replaceDroidComponent apsDroidLists(player), oldType, oldCompInc, newCompInc
    replaceDroidComponent mission_apsDroidLists(player), oldType, oldCompInc, newCompInc
    replaceDroidComponent apsLimboDroids(player), oldType, oldCompInc, newCompInc
    'check thru the templates (enumerateTemplates)
    For i = 0 To templateCount(player) - 1
        If Not replaceComponentInTemplate(droidTemplates_at(player, i), oldType, oldCompInc, newCompInc) Then
            Exit For
        End If
    Next
    ' also check build queues
    it = apsStructLists(player).head
    Do While it <> 0
        psCBuilding = lnObj(it)
        Select Case asStructureStats(aStruct(psCBuilding).pStructureType).typ
        Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
            If aStruct(psCBuilding).functionality.factory.psSubject <> 0 Then
                replaceComponentInTemplate aStruct(psCBuilding).functionality.factory.psSubject, oldType, oldCompInc, newCompInc
            End If
        End Select
        it = lnNext(it)
    Loop
    replaceStructureComponent apsStructLists(player), oldType, oldCompInc, newCompInc, player
    replaceStructureComponent mission_apsStructLists(player), oldType, oldCompInc, newCompInc, player
End Sub

'Looks through all the currently allocated stats to check the name is not
'a duplicate
Private Function checkResearchName(psResearch As RESEARCH, ByVal numStats As Long) As Boolean
    Dim inc As Long
    For inc = 0 To numStats - 1
        If Not (asResearch(inc).id <> psResearch.id) Then ASSERT_FAILURE "asResearch[inc].id.compare(psResearch->id) != 0", "src/research.cpp:1716", "checkResearchName", _
            "Research name has already been used - %s", psResearch.name_: checkResearchName = False: Exit Function
    Next
    checkResearchName = True
End Function

' structure.h's setStatusPendingHold and setStatusPendingRelease, for a RESEARCH_FACILITY
Private Sub setStatusPendingHold_research(func_ As RESEARCH_FACILITY)
    If func_.psSubjectPending = -1 Then
        func_.psSubjectPending = func_.psSubject
    End If
    func_.statusPending = FACTORY_HOLD_PENDING
    func_.pendingCount = func_.pendingCount + 1
End Sub

Private Sub setStatusPendingCancel_research(func_ As RESEARCH_FACILITY)
    func_.psSubjectPending = -1
    func_.statusPending = FACTORY_CANCEL_PENDING
    func_.pendingCount = func_.pendingCount + 1
End Sub

Private Sub setStatusPendingRelease_research(func_ As RESEARCH_FACILITY)
    If func_.psSubjectPending = -1 And func_.statusPending <> FACTORY_CANCEL_PENDING Then
        func_.psSubjectPending = func_.psSubject
    End If
    If func_.psSubjectPending <> -1 Then
        func_.statusPending = FACTORY_START_PENDING
    End If
    func_.pendingCount = func_.pendingCount + 1
End Sub
' structure.h's setStatusPendingStart (Functions for the GUI to know what's pending, before it's
' synchronised), for a RESEARCH_FACILITY (hci/research.cpp's startResearch): subject a RESEARCH *
Public Sub setStatusPendingStart_research(func_ As RESEARCH_FACILITY, ByVal subject As Long)
    func_.psSubjectPending = subject
    func_.statusPending = FACTORY_START_PENDING
    func_.pendingCount = func_.pendingCount + 1
End Sub

'set a research facility to hold
Public Sub holdResearch(ByVal psBuilding As Long, ByVal mode As QUEUE_MODE)
    If Not (asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESEARCH) Then ASSERT_FAILURE "psBuilding->pStructureType->type == REF_RESEARCH", "src/research.cpp:1317", "holdResearch", "structure not a research facility": Exit Sub

    If mode = ModeQueue Then
        sendStructureInfo psBuilding, STRUCTUREINFO_HOLDRESEARCH, 0
        setStatusPendingHold_research aStruct(psBuilding).functionality.researchFacility
        Exit Sub
    End If

    If aStruct(psBuilding).functionality.researchFacility.psSubject <> -1 Then
        'set the time the research facility was put on hold
        aStruct(psBuilding).functionality.researchFacility.timeStartHold = gameTime
        'play audio to indicate on hold
        If aStruct(psBuilding).player = selectedPlayer Then
            audio_PlayTrack ID_SOUND_WINDOWCLOSE
        End If
    End If

    delPowerRequest psBuilding
End Sub

'release a research facility from hold
Public Sub releaseResearch(ByVal psBuilding As Long, ByVal mode As QUEUE_MODE)
    If Not (asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESEARCH) Then ASSERT_FAILURE "psBuilding->pStructureType->type == REF_RESEARCH", "src/research.cpp:1345", "releaseResearch", "structure not a research facility": Exit Sub

    If mode = ModeQueue Then
        sendStructureInfo psBuilding, STRUCTUREINFO_RELEASERESEARCH, 0
        setStatusPendingRelease_research aStruct(psBuilding).functionality.researchFacility
        Exit Sub
    End If

    If aStruct(psBuilding).functionality.researchFacility.psSubject <> -1 And aStruct(psBuilding).functionality.researchFacility.timeStartHold <> 0 Then
        'adjust the start time for the current subject
        aStruct(psBuilding).functionality.researchFacility.timeStartHold = 0
    End If
End Sub

' Cancel All Research for player 0
Public Sub CancelAllResearch(ByVal pl As Long)
    Dim it As Long, psCurr As Long
    If Not ULT(pl, MAX_PLAYERS) Then Exit Sub

    it = apsStructLists(pl).head
    Do While it <> 0
        psCurr = lnObj(it)
        If asStructureStats(aStruct(psCurr).pStructureType).typ = REF_RESEARCH Then
            If aStruct(psCurr).pFunctionality <> 0 Then
                If aStruct(psCurr).functionality.researchFacility.psSubject <> -1 Then
                    debug_ 1382, LOG_NEVER, "CancelAllResearch", "canceling research for %p\n", psCurr
                    cancelResearch psCurr, ModeQueue
                End If
            End If
        End If
        it = lnNext(it)
    Loop
End Sub

' Sets the status of the topic to cancelled and stores the current research points accquired
Public Sub cancelResearch(ByVal psBuilding As Long, ByVal mode As QUEUE_MODE)
    Dim topicInc As Long, ok As Boolean, pl As Long

    ok = aStruct(psBuilding).pStructureType <> -1
    If ok Then ok = asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESEARCH
    If Not ok Then ASSERT_FAILURE "psBuilding->pStructureType && psBuilding->pStructureType->type == REF_RESEARCH", "src/research.cpp:1396", "cancelResearch", "Structure not a research facility": Exit Sub

    ' RESEARCH_FACILITY *psResFac = &psBuilding->pFunctionality->researchFacility
    If aStruct(psBuilding).functionality.researchFacility.psSubject = -1 Then
        debug_ 1401, LOG_SYNC, "cancelResearch", "Invalid research topic"
        Exit Sub
    End If
    topicInc = asResearch(aStruct(psBuilding).functionality.researchFacility.psSubject).index
    If Not (Not ULT(asResearch_size, topicInc)) Then ASSERT_FAILURE "topicInc <= asResearch.size()", "src/research.cpp:1405", "cancelResearch", "Invalid research topic %u (max %d)", topicInc, asResearch_size: Exit Sub
    pl = aStruct(psBuilding).player     ' pPlayerRes = &asPlayerResList[psBuilding->player][topicInc]
    If asStructureStats(aStruct(psBuilding).pStructureType).typ = REF_RESEARCH Then
        If mode = ModeQueue Then
            ' Tell others that we want to stop researching something.
            sendResearchStatus psBuilding, topicInc, aStruct(psBuilding).player, False
            ' Immediately tell the UI that we can research this now. (But don't change the game state.)
            MakeResearchCancelledPending asPlayerResList(pl, topicInc)
            setStatusPendingCancel_research aStruct(psBuilding).functionality.researchFacility
            Exit Sub  ' Wait for our message before doing anything. (Whatever this function does...)
        End If

        'check if waiting to accrue power
        If asPlayerResList(pl, topicInc).currentPoints = 0 Then
            ' Reset this topic as not having been researched
            ResetResearchStatus asPlayerResList(pl, topicInc)
        Else
            ' Set the researched flag
            MakeResearchCancelled asPlayerResList(pl, topicInc)
        End If

        ' Initialise the research facility's subject
        aStruct(psBuilding).functionality.researchFacility.psSubject = -1

        delPowerRequest psBuilding
    End If
End Sub

' For a given view data get the research this is related to (its index, or -1)
Public Function getResearchForMsg(ByVal pViewData As Long) As Long
    Dim inc As Long
    For inc = 0 To asResearch_size - 1
        If asResearch(inc).pViewData = pViewData Then   ' compare the pointer
            getResearchForMsg = inc
            Exit Function
        End If
    Next
    getResearchForMsg = -1
End Function

' Sets the 'possible' flag for a player's research so the topic will appear in
' the research list next time the Research Facility is selected
Public Function enableResearch(ByVal psResearch As Long, ByVal player As Long) As Boolean
    Dim inc As Long, prevState As Long

    If Not (psResearch <> -1) Then ASSERT_FAILURE "psResearch", "src/research.cpp:1728", "enableResearch", "No such research topic": enableResearch = False: Exit Function
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/research.cpp:1729", "enableResearch", "invalid player: %u", player: enableResearch = False: Exit Function

    inc = asResearch(psResearch).index
    If ULT(asResearch_size, inc) Then
        ASSERT_FAILURE "false", "src/research.cpp:1734", "enableResearch", "enableResearch: Invalid research topic - %s", asResearch(psResearch).name_
        enableResearch = False
        Exit Function
    End If

    prevState = intGetResearchState()

    'found, so set the flag
    MakeResearchPossible asPlayerResList(player, inc)

    If player = selectedPlayer Then
        'set the research reticule button to flash if research facility is free
        intNotifyResearchButton prevState
    End If

    enableResearch = True
End Function

' find the last research topic of importance that the losing player did and
' 'give' the results to the reward player
Public Sub researchReward(ByVal losingPlayer As Long, ByVal rewardPlayer As Long)
    Dim topicIndex As Long, researchPoints As Long, rewardID As Long, it As Long, psStruct As Long, psBest As Long

    'look through the losing players structures to find a research facility
    it = apsStructLists(losingPlayer).head
    Do While it <> 0
        psStruct = lnObj(it)
        If asStructureStats(aStruct(psStruct).pStructureType).typ = REF_RESEARCH Then
            ' RESEARCH_FACILITY *psFacility = (RESEARCH_FACILITY *)psStruct->pFunctionality
            psBest = aStruct(psStruct).functionality.researchFacility.psBestTopic
            If psBest <> -1 Then
                topicIndex = asResearch(psBest).ref - STAT_RESEARCH
                If topicIndex <> 0 Then
                    If Not IsResearchCompleted(asPlayerResList(rewardPlayer, topicIndex)) Then
                        'if it cost more - it is better (or should be)
                        If researchPoints < asResearch(topicIndex).researchPoints Then
                            'store the 'best' topic
                            researchPoints = asResearch(topicIndex).researchPoints
                            rewardID = topicIndex
                        End If
                    End If
                End If
            End If
        End If
        it = lnNext(it)
    Loop

    'if a topic was found give the reward player the results of that research
    If rewardID <> 0 Then
        researchResult rewardID, rewardPlayer, True, 0, True
        If rewardPlayer = selectedPlayer Then
            'name the actual reward
            CONPRINTF "%s :- %s", "Research Award", asResearch(rewardID).name_
        End If
    End If
End Sub

' flag self repair so droids can start when idle
Public Sub enableSelfRepair(ByVal player As Long)
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/research.cpp:1798", "enableSelfRepair", "invalid player: %u", player: Exit Sub
    bSelfRepair(player) = True
End Sub

' check to see if any research has been completed that enables self repair
Public Function selfRepairEnabled(ByVal player As Long) As Boolean
    If Not (player < MAX_PLAYERS) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/research.cpp:1805", "selfRepairEnabled", "invalid player: %u", player: selfRepairEnabled = False: Exit Function
    If bSelfRepair(player) Then
        selfRepairEnabled = True
    Else
        selfRepairEnabled = False
    End If
End Function

' for a given list of droids, replace the old component if exists
Public Sub replaceDroidComponent(pList As OBJ_LIST, ByVal oldType As Long, ByVal oldCompInc As Long, ByVal newCompInc As Long)
    Dim it As Long, psDroid As Long
    'check thru the droids
    it = pList.head
    Do While it <> 0
        psDroid = lnObj(it)
        switchComponent psDroid, oldType, oldCompInc, newCompInc
        ' Need to replace the units inside the transporter
        If DROID_isTransporter(psDroid) Then
            replaceTransDroidComponents psDroid, oldType, oldCompInc, newCompInc
        End If
        it = lnNext(it)
    Loop
End Sub

' replaces any components necessary for units that are inside a transporter
Public Sub replaceTransDroidComponents(ByVal psTransporter As Long, ByVal oldType As Long, ByVal oldCompInc As Long, ByVal newCompInc As Long)
    Dim it As Long, psCurr As Long
    If Not (DROID_isTransporter(psTransporter)) Then ASSERT_FAILURE "psTransporter->isTransporter()", "src/research.cpp:1836", "replaceTransDroidComponents", "invalid unit type"

    it = aGroup(aDroid(psTransporter).psGroup).psList.head
    Do While it <> 0
        psCurr = lnObj(it)
        If psCurr <> psTransporter Then
            switchComponent psCurr, oldType, oldCompInc, newCompInc
        End If
        it = lnNext(it)
    Loop
End Sub

Public Sub replaceStructureComponent(pList As OBJ_LIST, ByVal oldType As Long, ByVal oldCompInc As Long, ByVal newCompInc As Long, ByVal player As Long)
    Dim inc As Long, it As Long, psStructure As Long

    ' If the type is not one we are interested in, then don't bother checking
    If Not (oldType = COMP_ECM Or oldType = COMP_SENSOR Or oldType = COMP_WEAPON) Then
        Exit Sub
    End If

    'check thru the structures
    it = pList.head
    Do While it <> 0
        psStructure = lnObj(it)
        Select Case oldType
        Case COMP_WEAPON
            For inc = 0 To aStruct(psStructure).numWeaps - 1
                If aStruct(psStructure).asWeaps(inc).nStat > 0 Then
                    If aStruct(psStructure).asWeaps(inc).nStat = oldCompInc Then
                        aStruct(psStructure).asWeaps(inc).nStat = newCompInc
                    End If
                End If
            Next
        Case Else
            'ignore all other component types
        End Select
        it = lnNext(it)
    Loop
End Sub

' swaps the old component for the new one for a specific droid
Private Sub switchComponent(ByVal psDroid As Long, ByVal oldType As Long, ByVal oldCompInc As Long, ByVal newCompInc As Long)
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/research.cpp:1886", "switchComponent", "Invalid droid pointer": Exit Sub

    Select Case oldType
    Case COMP_BODY, COMP_BRAIN, COMP_PROPULSION, COMP_REPAIRUNIT, COMP_ECM, COMP_SENSOR, COMP_CONSTRUCT
        If aDroid(psDroid).asBits(oldType) = oldCompInc Then
            aDroid(psDroid).asBits(oldType) = newCompInc And &HFF&
        End If
    Case COMP_WEAPON
        ' Can only be one weapon now
        If aDroid(psDroid).asWeaps(0).nStat > 0 Then
            If aDroid(psDroid).asWeaps(0).nStat = oldCompInc Then
                aDroid(psDroid).asWeaps(0).nStat = newCompInc
            End If
        End If
    Case Else
        'unknown comp type
        debug_ 1914, LOG_ERROR, "switchComponent", "Unknown component type - invalid droid"
        exit_ 3                         ' abort()
        Exit Sub
    End Select
End Sub

' (std::sort's comparison: a goes before b)
Private Function allyResearchSortFunction(a As AllyResearch, b As AllyResearch) As Boolean
    If a.active <> b.active Then
        allyResearchSortFunction = a.active
        Exit Function
    End If
    If a.timeToResearch <> b.timeToResearch Then
        allyResearchSortFunction = ULT(a.timeToResearch, b.timeToResearch)     ' Unsigned cast = sort -1 as infinite.
        Exit Function
    End If
    If a.powerNeeded <> b.powerNeeded Then
        allyResearchSortFunction = ULT(a.powerNeeded, b.powerNeeded)
        Exit Function
    End If
    If a.completion <> b.completion Then
        allyResearchSortFunction = a.completion > b.completion
        Exit Function
    End If
    allyResearchSortFunction = ULT(a.player, b.player)
End Function

' (VB) researches[ref]: its index in allyResearchKeys/allyResearchValues, inserted if missing
Private Function AllyResearchSlot(ByVal ref As Long) As Long
    Dim i As Long, k As Long
    i = 0
    Do While i < allyResearch_size
        If Not ULT(allyResearchKeys(i), ref) Then Exit Do
        i = i + 1
    Loop
    If i < allyResearch_size Then
        If allyResearchKeys(i) = ref Then
            AllyResearchSlot = i
            Exit Function
        End If
    End If
    ReDim Preserve allyResearchKeys(0 To allyResearch_size)
    ReDim Preserve allyResearchValues(0 To allyResearch_size)
    For k = allyResearch_size To i + 1 Step -1
        allyResearchKeys(k) = allyResearchKeys(k - 1)
        allyResearchValues(k) = allyResearchValues(k - 1)
    Next
    allyResearchKeys(i) = ref
    allyResearchValues(i).size = 0
    Erase allyResearchValues(i).items
    allyResearch_size = allyResearch_size + 1
    AllyResearchSlot = i
End Function

' std::vector<AllyResearch> const &listAllyResearch(unsigned ref): copied into result(0 ..), the count returned
Public Function listAllyResearch(ByVal ref As Long, result() As AllyResearch) As Long
    Dim player As Long, it As Long, psStruct As Long, subject As Long, cRef As Long, r As AllyResearch
    Dim slot As Long, i As Long, j As Long, tmp As AllyResearch, n As Long

    If Not ULT(selectedPlayer, MAX_PLAYERS) Then
        listAllyResearch = 0    ' noAllyResearch
        Exit Function
    End If

    If gameTime <> allyLastGameTime Then
        allyLastGameTime = gameTime
        Erase allyResearchKeys
        Erase allyResearchValues
        allyResearch_size = 0

        For player = 0 To MAX_PLAYERS - 1
            If player = selectedPlayer Then GoTo continue_player
            If Not aiCheckAlliances(selectedPlayer, player) Then GoTo continue_player
            If Not alliancesSharedResearch(game.alliance) Then GoTo continue_player  ' Skip this player, not an ally.

            ' Check each research facility to see if they are doing this topic. (As opposed to having started the topic, but stopped researching it.)
            it = apsStructLists(player).head
            Do While it <> 0
                psStruct = lnObj(it)
                ' RESEARCH_FACILITY *res = (RESEARCH_FACILITY *)psStruct->pFunctionality
                If asStructureStats(aStruct(psStruct).pStructureType).typ <> REF_RESEARCH Then GoTo continue_struct
                If aStruct(psStruct).functionality.researchFacility.psSubject = -1 Then GoTo continue_struct  ' Not a researching research facility.

                subject = aStruct(psStruct).functionality.researchFacility.psSubject
                ' PLAYER_RESEARCH const &playerRes = asPlayerResList[player][subject.index]
                cRef = asResearch(subject).ref

                r.player = player
                r.completion = asPlayerResList(player, asResearch(subject).index).currentPoints
                r.powerNeeded = checkPowerRequest(psStruct)
                r.timeToResearch = -1
                If r.powerNeeded = -1 Then
                    r.timeToResearch = FromU32(Fix(U32(asResearch(subject).researchPoints - asPlayerResList(player, asResearch(subject).index).currentPoints) / U32(MAX_(getBuildingResearchPoints(psStruct), 1))))
                End If
                r.active = aStruct(psStruct).status = SS_BUILT
                slot = AllyResearchSlot(cRef)
                ReDim Preserve allyResearchValues(slot).items(0 To allyResearchValues(slot).size)
                allyResearchValues(slot).items(allyResearchValues(slot).size) = r
                allyResearchValues(slot).size = allyResearchValues(slot).size + 1
continue_struct:
                it = lnNext(it)
            Loop
continue_player:
        Next
        ' std::sort(..., allyResearchSortFunction): (VB) an insertion sort; elements that compare equal are equal
        For slot = 0 To allyResearch_size - 1
            For i = 1 To allyResearchValues(slot).size - 1
                tmp = allyResearchValues(slot).items(i)
                j = i - 1
                Do While j >= 0
                    If Not allyResearchSortFunction(tmp, allyResearchValues(slot).items(j)) Then Exit Do
                    allyResearchValues(slot).items(j + 1) = allyResearchValues(slot).items(j)
                    j = j - 1
                Loop
                allyResearchValues(slot).items(j + 1) = tmp
            Next
        Next
    End If

    For slot = 0 To allyResearch_size - 1
        If allyResearchKeys(slot) = ref Then
            n = allyResearchValues(slot).size
            If n > 0 Then
                ReDim result(0 To n - 1)
                For i = 0 To n - 1
                    result(i) = allyResearchValues(slot).items(i)
                Next
            End If
            listAllyResearch = n
            Exit Function
        End If
    Next
    listAllyResearch = 0    ' noAllyResearch
End Function

' Recursively disable research for all players
Private Sub RecursivelyDisableResearchByID(ByVal index As Long)
    Dim player As Long, inc As Long, prereq As Long
    If IsResearchDisabled(asPlayerResList(0, index)) Then
        Exit Sub
    End If

    For player = 0 To MAX_PLAYERS - 1
        DisableResearch asPlayerResList(player, index)
    Next

    For inc = 0 To asResearch_size - 1
        For prereq = 0 To asResearch(inc).pPRList_size - 1
            If asResearch(inc).pPRList(prereq) = index Then
                RecursivelyDisableResearchByID inc
            End If
        Next
    Next
End Sub

Public Sub RecursivelyDisableResearchByFlags(ByVal flags As Long)
    Dim inc As Long
    For inc = 0 To asResearch_size - 1
        If (asResearch(inc).disabledWhen And flags) <> 0 Then
            RecursivelyDisableResearchByID inc
        End If
    Next
End Sub

' (VB, for the tests) a copy of a string map (Nothing stays Nothing)
Private Function CopyStringMap(ByVal d As Object) As Object
    Dim k As Variant
    If d Is Nothing Then Exit Function
    Set CopyStringMap = NewStringMap()
    For Each k In d.Keys
        CopyStringMap.Add k, d(k)
    Next
End Function

' (VB, for the tests: statsref's statsref_researchState) save, or restore, what
' research changes here besides the players' lists and the stats: bSelfRepair,
' playerUpgradeCounts and cachedPerPlayerRawUpgradeChange
Public Sub research_testState(ByVal save As Boolean)
    Dim i As Long
    If save Then
        For i = 0 To MAX_PLAYERS - 1
            testSelfRepair(i) = bSelfRepair(i)
            Set testRawUpgradeChange(i) = CopyStringMap(cachedPerPlayerRawUpgradeChange(i))
        Next
        testUpgradeCounts_size = playerUpgradeCounts_size
        If playerUpgradeCounts_size > 0 Then
            ReDim testUpgradeCounts(0 To playerUpgradeCounts_size - 1)
            For i = 0 To playerUpgradeCounts_size - 1
                Set testUpgradeCounts(i).numBodyClassArmourUpgrades = CopyStringMap(playerUpgradeCounts(i).numBodyClassArmourUpgrades)
                Set testUpgradeCounts(i).numBodyClassThermalUpgrades = CopyStringMap(playerUpgradeCounts(i).numBodyClassThermalUpgrades)
                Set testUpgradeCounts(i).numWeaponImpactClassUpgrades = CopyStringMap(playerUpgradeCounts(i).numWeaponImpactClassUpgrades)
            Next
        End If
    Else
        For i = 0 To MAX_PLAYERS - 1
            bSelfRepair(i) = testSelfRepair(i)
            Set cachedPerPlayerRawUpgradeChange(i) = CopyStringMap(testRawUpgradeChange(i))
        Next
        playerUpgradeCounts_size = testUpgradeCounts_size
        If testUpgradeCounts_size > 0 Then
            ReDim playerUpgradeCounts(0 To testUpgradeCounts_size - 1)
            For i = 0 To testUpgradeCounts_size - 1
                Set playerUpgradeCounts(i).numBodyClassArmourUpgrades = CopyStringMap(testUpgradeCounts(i).numBodyClassArmourUpgrades)
                Set playerUpgradeCounts(i).numBodyClassThermalUpgrades = CopyStringMap(testUpgradeCounts(i).numBodyClassThermalUpgrades)
                Set playerUpgradeCounts(i).numWeaponImpactClassUpgrades = CopyStringMap(testUpgradeCounts(i).numWeaponImpactClassUpgrades)
            Next
        Else
            Erase playerUpgradeCounts
        End If
    End If
End Sub
