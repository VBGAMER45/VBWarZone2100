Attribute VB_Name = "template_c"
' src/template.cpp, template.h: so far the template store and the loader of
' templates.json. The templates live in a pool, apsTemplates (a
' DROID_TEMPLATE * is its 1-based index, 0 = nullptr; free slots chain
' through multiPlayerID); std::map<UDWORD, std::unique_ptr<DROID_TEMPLATE>>
' droidTemplates[player] is a list of (multiPlayerID, template) pairs kept
' in the map's order (unsigned ascending), and replacedDroidTemplates
' keeps the replaced ones alive, as in C++. enumerateTemplates
' (std::function) is a loop over droidTemplates_at(player, i) for i below
' templateCount(player). The loadTemplateCommon overload for a WzConfig is
' loadTemplateCommon_ini. The DROID_TEMPLATE methods are
' DROID_TEMPLATE_<method>(template) and return the stat's index.
' saveTemplateCommon builds an nlohmann_json; storeTemplates writes the
' stored designs through a WzConfig (saved when the object goes). The
' vector fillTemplateList returns is its pList(0 .. count - 1) with the
' count as its value.
' Not ported yet: listTemplates (a debug listing).
Option Explicit

Private Type TEMPLATE_MAP                   ' (VB) std::map<UDWORD, std::unique_ptr<DROID_TEMPLATE>>
    keys() As Long                          ' multiPlayerID, ascending (unsigned)
    values() As Long                        ' DROID_TEMPLATE *
    size As Long
End Type

Private Type TEMPLATE_LIST                  ' (VB) std::vector<std::unique_ptr<DROID_TEMPLATE>>
    items() As Long
    size As Long
End Type

' (VB) the pool of templates
Public apsTemplates() As DROID_TEMPLATE
Private apsTemplatesUsed() As Boolean
Private apsTemplatesCap As Long
Private freeTemplates As Long               ' first free slot (chained through multiPlayerID)

' Template storage
Private droidTemplates(0 To MAX_PLAYERS - 1) As TEMPLATE_MAP
Private replacedDroidTemplates(0 To MAX_PLAYERS - 1) As TEMPLATE_LIST

Public allowDesign As Boolean
Public includeRedundantDesigns As Boolean
Public playerBuiltHQ As Boolean

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub template_c_init()
    allowDesign = True
End Sub

' (VB) a new template in the pool, constructed (DROID_TEMPLATE::DROID_TEMPLATE())
Public Function newTemplate() As Long
    Dim h As Long, i As Long, blank As DROID_TEMPLATE
    If freeTemplates = 0 Then
        ReDim Preserve apsTemplates(1 To apsTemplatesCap * 2 + 16)
        ReDim Preserve apsTemplatesUsed(1 To apsTemplatesCap * 2 + 16)
        For i = apsTemplatesCap + 1 To UBound(apsTemplates)
            If i < UBound(apsTemplates) Then apsTemplates(i).multiPlayerID = i + 1 Else apsTemplates(i).multiPlayerID = 0
        Next
        freeTemplates = apsTemplatesCap + 1
        apsTemplatesCap = UBound(apsTemplates)
    End If
    h = freeTemplates
    freeTemplates = apsTemplates(h).multiPlayerID
    apsTemplates(h) = blank
    DROID_TEMPLATE_construct apsTemplates(h)
    apsTemplatesUsed(h) = True
    newTemplate = h
End Function

' (VB) a template's destructor: its slot back to the pool
Public Sub deleteTemplate(ByVal h As Long)
    Dim blank As DROID_TEMPLATE
    apsTemplates(h) = blank
    apsTemplatesUsed(h) = False
    apsTemplates(h).multiPlayerID = freeTemplates
    freeTemplates = h
End Sub

' DROID_TEMPLATE::DROID_TEMPLATE()  // This constructor replaces a memset in scrAssembleWeaponTemplate(), not needed elsewhere.
Public Sub DROID_TEMPLATE_construct(t As DROID_TEMPLATE)
    Dim i As Long
    t.ref = STAT_TEMPLATE                   ' BASE_STATS(STAT_TEMPLATE)
    t.numWeaps = 0
    t.droidType = DROID_WEAPON
    t.multiPlayerID = 0
    t.prefab = False
    t.stored = False
    t.enabled = False
    For i = 0 To DROID_MAXCOMP - 1
        t.asParts(i) = 0
    Next
    For i = 0 To MAX_WEAPONS - 1
        t.asWeaps(i) = 0
    Next
End Sub

Private Function researchedItem(ByVal player As Long, ByVal partIndex As COMPONENT_TYPE, ByVal part As Long, ByVal allowZero As Boolean, ByVal allowRedundant As Boolean) As Boolean
    Dim availability As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:57", "researchedItem", "Invalid player: %u", player: researchedItem = False: Exit Function
    If allowZero And part <= 0 Then
        researchedItem = True
        Exit Function
    End If
    availability = apCompLists(player, partIndex, part)
    researchedItem = availability = AVAILABLE Or (allowRedundant And availability = REDUNDANT)
End Function

Private Function researchedPart(psCurr As DROID_TEMPLATE, ByVal player As Long, ByVal partIndex As COMPONENT_TYPE, ByVal allowZero As Boolean, ByVal allowRedundant As Boolean) As Boolean
    researchedPart = researchedItem(player, partIndex, psCurr.asParts(partIndex), allowZero, allowRedundant)
End Function

Private Function researchedWeap(psCurr As DROID_TEMPLATE, ByVal player As Long, ByVal weapIndex As Long, ByVal allowRedundant As Boolean) As Boolean
    Dim availability As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:73", "researchedWeap", "Invalid player: %u", player: researchedWeap = False: Exit Function
    availability = apCompLists(player, COMP_WEAPON, psCurr.asWeaps(weapIndex))
    researchedWeap = availability = AVAILABLE Or (allowRedundant And availability = REDUNDANT)
End Function

' psCurr: a DROID_TEMPLATE *
Public Function researchedTemplate(ByVal psCurr As Long, ByVal player As Long, Optional ByVal allowRedundant As Boolean = False, Optional ByVal verbose As Boolean = False) As Boolean
    Dim resBody As Boolean, resBrain As Boolean, resProp As Boolean, resSensor As Boolean, resEcm As Boolean
    Dim resRepair As Boolean, resConstruct As Boolean, researchedEverything As Boolean, weapIndex As Long
    If Not (psCurr <> 0) Then ASSERT_FAILURE "psCurr", "src/template.cpp:80", "researchedTemplate", "Given a null template": researchedTemplate = False: Exit Function
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:81", "researchedTemplate", "Invalid player: %u", player: researchedTemplate = False: Exit Function
    resBody = researchedPart(apsTemplates(psCurr), player, COMP_BODY, False, allowRedundant)
    resBrain = researchedPart(apsTemplates(psCurr), player, COMP_BRAIN, True, allowRedundant)
    resProp = researchedPart(apsTemplates(psCurr), player, COMP_PROPULSION, False, allowRedundant)
    resSensor = researchedPart(apsTemplates(psCurr), player, COMP_SENSOR, True, allowRedundant)
    resEcm = researchedPart(apsTemplates(psCurr), player, COMP_ECM, True, allowRedundant)
    resRepair = researchedPart(apsTemplates(psCurr), player, COMP_REPAIRUNIT, True, allowRedundant)
    resConstruct = researchedPart(apsTemplates(psCurr), player, COMP_CONSTRUCT, True, allowRedundant)
    researchedEverything = resBody And resBrain And resProp And resSensor And resEcm And resRepair And resConstruct
    If verbose And Not researchedEverything Then
        debug_ 92, LOG_ERROR, "researchedTemplate", "%s : not researched : body=%d brai=%d prop=%d sensor=%d ecm=%d rep=%d con=%d", apsTemplates(psCurr).name_, _
            -CLng(resBody), -CLng(resBrain), -CLng(resProp), -CLng(resSensor), -CLng(resEcm), -CLng(resRepair), -CLng(resConstruct)
    End If
    weapIndex = 0
    Do While weapIndex < apsTemplates(psCurr).numWeaps And researchedEverything
        researchedEverything = researchedWeap(apsTemplates(psCurr), player, weapIndex, allowRedundant)
        If Not researchedEverything And verbose Then
            debug_ 100, LOG_ERROR, "researchedTemplate", "%s : not researched weapon %u", apsTemplates(psCurr).name_, weapIndex
        End If
        weapIndex = weapIndex + 1
    Loop
    researchedTemplate = researchedEverything
End Function

Public Function droidTemplate_LoadPartByName(ByVal compType As COMPONENT_TYPE, ByVal name_ As String, outputTemplate As DROID_TEMPLATE) As Boolean
    Dim index As Long
    index = getCompFromName(compType, name_)
    If index < 0 Then
        debug_ 111, LOG_ERROR, "droidTemplate_LoadPartByName", "Stored template contains an unknown (type: %d) component: %s.", compType, name_
        droidTemplate_LoadPartByName = False
        Exit Function
    End If
    If index > 255 Then
        ' returned index exceeds uint8_t max - consider changing type of asParts in DROID_TEMPLATE?
        debug_ 117, LOG_ERROR, "droidTemplate_LoadPartByName", "Stored template contains a (type: %d) component (%s) index that exceeds UINT8_MAX: %d", compType, name_, index
        droidTemplate_LoadPartByName = False
        Exit Function
    End If
    outputTemplate.asParts(compType) = index
    droidTemplate_LoadPartByName = True
End Function

Public Function droidTemplate_LoadWeapByName(ByVal destIndex As Long, ByVal name_ As String, outputTemplate As DROID_TEMPLATE) As Boolean
    Dim index As Long
    index = getCompFromName(COMP_WEAPON, name_)
    If index < 0 Then
        debug_ 129, LOG_ERROR, "droidTemplate_LoadWeapByName", "Stored template contains an unknown (type: %d) component: %s.", COMP_WEAPON, name_
        droidTemplate_LoadWeapByName = False
        Exit Function
    End If
    ' (INT_MAX > UINT32_MAX is false)
    outputTemplate.asWeaps(destIndex) = index
    droidTemplate_LoadWeapByName = True
End Function

' loadTemplateCommon(WzConfig &ini, DROID_TEMPLATE &outputTemplate)
Public Function loadTemplateCommon_ini(ByVal ini As WzConfig, outputTemplate As DROID_TEMPLATE) As Boolean
    loadTemplateCommon_ini = loadTemplateCommon(ini.currentJsonValue(), outputTemplate)
End Function

Public Function jsonGetAsWzString(ByVal obj As nlohmann_json, ByVal key As String, Optional ByVal defaultValue As String = "") As String
    Dim it As nlohmann_json, ok As Variant, stringValue As String
    Set it = obj.find(key)
    If it Is Nothing Then
        jsonGetAsWzString = defaultValue
    Else
        ' Use json_variant to support conversion from other value types to string
        ok = False
        stringValue = json_variant_new(it).toWzString(ok)
        If Not CBool(ok) Then
            ASSERT_FAILURE "false", "src/template.cpp:163", "jsonGetAsWzString", "Failed to convert value of key ""%s"" to string", key
            jsonGetAsWzString = defaultValue
            Exit Function
        End If
        jsonGetAsWzString = stringValue
    End If
End Function

Public Function loadTemplateCommon(ByVal obj As nlohmann_json, outputTemplate As DROID_TEMPLATE) As Boolean
    Dim droidType As String, it As nlohmann_json, weapons() As String, nWeapons As Long, w As String
    If Not (obj.is_object()) Then ASSERT_FAILURE "obj.is_object()", "src/template.cpp:172", "loadTemplateCommon", "Not an object?": loadTemplateCommon = False: Exit Function
    ' (DROID_TEMPLATE &design = outputTemplate)
    outputTemplate.name_ = jsonGetAsWzString(obj, "name")
    droidType = jsonGetAsWzString(obj, "type")

    If droidType = "ECM" Then
        outputTemplate.droidType = DROID_ECM
    ElseIf droidType = "SENSOR" Then
        outputTemplate.droidType = DROID_SENSOR
    ElseIf droidType = "CONSTRUCT" Then
        outputTemplate.droidType = DROID_CONSTRUCT
    ElseIf droidType = "WEAPON" Then
        outputTemplate.droidType = DROID_WEAPON
    ElseIf droidType = "PERSON" Then
        outputTemplate.droidType = DROID_PERSON
    ElseIf droidType = "CYBORG" Then
        outputTemplate.droidType = DROID_CYBORG
    ElseIf droidType = "CYBORG_SUPER" Then
        outputTemplate.droidType = DROID_CYBORG_SUPER
    ElseIf droidType = "CYBORG_CONSTRUCT" Then
        outputTemplate.droidType = DROID_CYBORG_CONSTRUCT
    ElseIf droidType = "CYBORG_REPAIR" Then
        outputTemplate.droidType = DROID_CYBORG_REPAIR
    ElseIf droidType = "TRANSPORTER" Then
        outputTemplate.droidType = DROID_TRANSPORTER
    ElseIf droidType = "SUPERTRANSPORTER" Then
        outputTemplate.droidType = DROID_SUPERTRANSPORTER
    ElseIf droidType = "DROID" Then
        outputTemplate.droidType = DROID_DEFAULT
    ElseIf droidType = "DROID_COMMAND" Then
        outputTemplate.droidType = DROID_COMMAND
    ElseIf droidType = "REPAIR" Then
        outputTemplate.droidType = DROID_REPAIR
    Else
        ASSERT_FAILURE "false", "src/template.cpp:235", "loadTemplateCommon", "No such droid type ""%s"" for %s", droidType, outputTemplate.id
    End If

    If Not droidTemplate_LoadPartByName(COMP_BODY, jsonGetAsWzString(obj, "body"), outputTemplate) Then loadTemplateCommon = False: Exit Function
    If Not droidTemplate_LoadPartByName(COMP_BRAIN, jsonGetAsWzString(obj, "brain", "ZNULLBRAIN"), outputTemplate) Then loadTemplateCommon = False: Exit Function
    If Not droidTemplate_LoadPartByName(COMP_PROPULSION, jsonGetAsWzString(obj, "propulsion", "ZNULLPROP"), outputTemplate) Then loadTemplateCommon = False: Exit Function
    If Not droidTemplate_LoadPartByName(COMP_REPAIRUNIT, jsonGetAsWzString(obj, "repair", "ZNULLREPAIR"), outputTemplate) Then loadTemplateCommon = False: Exit Function
    If Not droidTemplate_LoadPartByName(COMP_ECM, jsonGetAsWzString(obj, "ecm", "ZNULLECM"), outputTemplate) Then loadTemplateCommon = False: Exit Function
    If Not droidTemplate_LoadPartByName(COMP_SENSOR, jsonGetAsWzString(obj, "sensor", "ZNULLSENSOR"), outputTemplate) Then loadTemplateCommon = False: Exit Function
    If Not droidTemplate_LoadPartByName(COMP_CONSTRUCT, jsonGetAsWzString(obj, "construct", "ZNULLCONSTRUCT"), outputTemplate) Then loadTemplateCommon = False: Exit Function

    Set it = obj.find("weapons")
    If Not it Is Nothing Then
        nWeapons = json_variant_new(it).toWzStringList(weapons)
    End If
    If Not (nWeapons <= MAX_WEAPONS) Then ASSERT_FAILURE "weapons.size() <= 3", "src/template.cpp:252", "loadTemplateCommon", "Number of weapons (%zu) exceeds MAX_WEAPONS (%d)", nWeapons, MAX_WEAPONS
    If nWeapons <= MAX_WEAPONS Then outputTemplate.numWeaps = nWeapons Else outputTemplate.numWeaps = MAX_WEAPONS
    If nWeapons >= 1 Then w = weapons(0) Else w = "ZNULLWEAPON"
    If Not droidTemplate_LoadWeapByName(0, w, outputTemplate) Then loadTemplateCommon = False: Exit Function
    If nWeapons >= 2 Then w = weapons(1) Else w = "ZNULLWEAPON"
    If Not droidTemplate_LoadWeapByName(1, w, outputTemplate) Then loadTemplateCommon = False: Exit Function
    If nWeapons >= 3 Then w = weapons(2) Else w = "ZNULLWEAPON"
    If Not droidTemplate_LoadWeapByName(2, w, outputTemplate) Then loadTemplateCommon = False: Exit Function

    loadTemplateCommon = True
End Function

' designableTemplate's lambda designablePart (a component's designable flag and name): records the part that is not designable
Private Function designablePart(ByVal designable As Boolean, ByVal componentName As String, ByVal part As String, failPart As String, failPartName As String) As Boolean
    If Not designable Then
        failPart = part
        failPartName = componentName
    End If
    designablePart = designable
End Function

' (VB) designableTemplate's (asParts[x] == 0 || designablePart(*getXStats(), part))
Private Function partDesignable(ByVal partIndex As Long, ByVal designable As Boolean, ByVal componentName As String, ByVal part As String, failPart As String, failPartName As String) As Boolean
    If partIndex = 0 Then
        partDesignable = True
    Else
        partDesignable = designablePart(designable, componentName, part, failPart, failPartName)
    End If
End Function

' (VB) designableTemplate's ((asParts[x] == 0) || (!designablePart(*getXStats(), part) && getXStats()->usageClass == UsageClass::Cyborg))
Private Function cyborgPartDesignable(ByVal partIndex As Long, ByVal designable As Boolean, ByVal componentName As String, ByVal usage As UsageClass, ByVal part As String, failPart As String, failPartName As String) As Boolean
    If partIndex = 0 Then
        cyborgPartDesignable = True
    ElseIf Not designablePart(designable, componentName, part, failPart, failPartName) Then
        cyborgPartDesignable = usage = UsageClass_Cyborg
    End If
End Function

' (VB) designableTemplate's cyborg weapon slot: ((numWeaps <= slot) || [slot 0: (getBrainStats()->psWeaponStat == getWeaponStats(0)) ||]
' (!designablePart(*getWeaponStats(slot), part) && ((!isSuperCyborg && usageClass == Cyborg) || (isSuperCyborg && usageClass == SuperCyborg))))
Private Function cyborgWeaponDesignable(psTempl As DROID_TEMPLATE, ByVal slot As Long, ByVal isSuperCyborg As Boolean, ByVal part As String, failPart As String, failPartName As String) As Boolean
    Dim w As Long
    If psTempl.numWeaps <= slot Then
        cyborgWeaponDesignable = True
        Exit Function
    End If
    w = DROID_TEMPLATE_getWeaponStats(psTempl, slot)
    If slot = 0 Then
        If asBrainStats(DROID_TEMPLATE_getBrainStats(psTempl)).psWeaponStat = w Then
            cyborgWeaponDesignable = True
            Exit Function
        End If
    End If
    If Not designablePart(asWeaponStats(w).designable, asWeaponStats(w).name_, part, failPart, failPartName) Then
        If Not isSuperCyborg Then
            cyborgWeaponDesignable = asWeaponStats(w).usageClass = UsageClass_Cyborg
        Else
            cyborgWeaponDesignable = asWeaponStats(w).usageClass = UsageClass_SuperCyborg
        End If
    End If
End Function

' A way to check if a design is something someone could legitimately have in multiplayer
' (VB) the && / || chains are evaluated in C's order, since designablePart records the part that failed
Public Function designableTemplate(psTempl As DROID_TEMPLATE, ByVal player As Long) As Boolean
    Dim failPart As String, failPartName As String, designable As Boolean, repair As Boolean, isSuperCyborg As Boolean
    Dim transporter As Boolean, b As Long, pr As Long, r As Long, br As Long, e As Long, sn As Long, c As Long, w As Long
    If Not bMultiPlayer Or Not isHumanPlayer(player) Then
        designableTemplate = True ' Don't care about AIs
        Exit Function
    End If

    failPart = vbNullString
    b = DROID_TEMPLATE_getBodyStats(psTempl)
    pr = DROID_TEMPLATE_getPropulsionStats(psTempl)
    r = DROID_TEMPLATE_getRepairStats(psTempl)
    br = DROID_TEMPLATE_getBrainStats(psTempl)
    e = DROID_TEMPLATE_getECMStats(psTempl)
    sn = DROID_TEMPLATE_getSensorStats(psTempl)
    c = DROID_TEMPLATE_getConstructStats(psTempl)

    designable = False

    If psTempl.droidType = DROID_CYBORG Or _
       psTempl.droidType = DROID_CYBORG_SUPER Or _
       psTempl.droidType = DROID_CYBORG_CONSTRUCT Or _
       psTempl.droidType = DROID_CYBORG_REPAIR Then
        repair = psTempl.asParts(COMP_REPAIRUNIT) = 0
        If Not repair Then
            If Not designablePart(asRepairStats(r).designable, asRepairStats(r).name_, "Repair unit", failPart, failPartName) Then
                repair = asRepairStats(r).usageClass = UsageClass_Cyborg
            End If
        End If
        If Not repair Then
            If psTempl.asParts(COMP_REPAIRUNIT) <> 0 Then repair = selfRepairEnabled(player)
        End If
        isSuperCyborg = asBodyStats(b).usageClass = UsageClass_SuperCyborg

        designable = Not designablePart(asBodyStats(b).designable, asBodyStats(b).name_, "Body", failPart, failPartName)
        If designable Then designable = StrComp(asBodyStats(b).bodyClass, "Cyborgs", vbBinaryCompare) = 0
        If designable Then designable = Not designablePart(asPropulsionStats(pr).designable, asPropulsionStats(pr).name_, "Propulsion", failPart, failPartName)
        If designable Then designable = asPropulsionStats(pr).propulsionType = PROPULSION_TYPE_LEGGED
        If designable Then designable = asPropulsionStats(pr).usageClass = UsageClass_Cyborg
        If designable Then designable = repair
        If designable Then designable = cyborgPartDesignable(psTempl.asParts(COMP_BRAIN), asBrainStats(br).designable, asBrainStats(br).name_, asBrainStats(br).usageClass, "Brain", failPart, failPartName)
        If designable Then designable = cyborgPartDesignable(psTempl.asParts(COMP_ECM), asECMStats(e).designable, asECMStats(e).name_, asECMStats(e).usageClass, "ECM", failPart, failPartName)
        If designable Then designable = cyborgPartDesignable(psTempl.asParts(COMP_SENSOR), asSensorStats(sn).designable, asSensorStats(sn).name_, asSensorStats(sn).usageClass, "Sensor", failPart, failPartName)
        If designable Then designable = cyborgPartDesignable(psTempl.asParts(COMP_CONSTRUCT), asConstructStats(c).designable, asConstructStats(c).name_, asConstructStats(c).usageClass, "Construction part", failPart, failPartName)
        If designable Then designable = cyborgWeaponDesignable(psTempl, 0, isSuperCyborg, "Weapon 0", failPart, failPartName)
        If designable Then designable = cyborgWeaponDesignable(psTempl, 1, isSuperCyborg, "Weapon 1", failPart, failPartName)
        If designable Then designable = cyborgWeaponDesignable(psTempl, 2, isSuperCyborg, "Weapon 2", failPart, failPartName)
    Else
        repair = psTempl.asParts(COMP_REPAIRUNIT) = 0
        If Not repair Then repair = designablePart(asRepairStats(r).designable, asRepairStats(r).name_, "Repair unit", failPart, failPartName)
        If Not repair Then
            If psTempl.asParts(COMP_REPAIRUNIT) <> 0 Then repair = selfRepairEnabled(player)
        End If
        transporter = (psTempl.droidType = DROID_TRANSPORTER) Or (psTempl.droidType = DROID_SUPERTRANSPORTER)

        designable = transporter
        If Not designable Then designable = designablePart(asBodyStats(b).designable, asBodyStats(b).name_, "Body", failPart, failPartName)
        If designable Then designable = designablePart(asPropulsionStats(pr).designable, asPropulsionStats(pr).name_, "Propulsion", failPart, failPartName)
        If designable Then designable = partDesignable(psTempl.asParts(COMP_BRAIN), asBrainStats(br).designable, asBrainStats(br).name_, "Brain", failPart, failPartName)
        If designable Then designable = repair
        If designable Then designable = partDesignable(psTempl.asParts(COMP_ECM), asECMStats(e).designable, asECMStats(e).name_, "ECM", failPart, failPartName)
        If designable Then designable = partDesignable(psTempl.asParts(COMP_SENSOR), asSensorStats(sn).designable, asSensorStats(sn).name_, "Sensor", failPart, failPartName)
        If designable Then designable = partDesignable(psTempl.asParts(COMP_CONSTRUCT), asConstructStats(c).designable, asConstructStats(c).name_, "Construction part", failPart, failPartName)
        If designable Then
            If psTempl.numWeaps > 0 Then
                w = DROID_TEMPLATE_getWeaponStats(psTempl, 0)
                If asBrainStats(br).psWeaponStat <> w Then
                    designable = designablePart(asWeaponStats(w).designable, asWeaponStats(w).name_, "Weapon 0", failPart, failPartName)
                End If
            End If
        End If
        If designable Then
            If psTempl.numWeaps > 1 Then
                w = DROID_TEMPLATE_getWeaponStats(psTempl, 1)
                designable = designablePart(asWeaponStats(w).designable, asWeaponStats(w).name_, "Weapon 1", failPart, failPartName)
            End If
        End If
        If designable Then
            If psTempl.numWeaps > 2 Then
                w = DROID_TEMPLATE_getWeaponStats(psTempl, 2)
                designable = designablePart(asWeaponStats(w).designable, asWeaponStats(w).name_, "Weapon 2", failPart, failPartName)
            End If
        End If

        If transporter And designable Then
            designable = asPropulsionStats(pr).propulsionType = PROPULSION_TYPE_LIFT
            If designable Then designable = Not designablePart(asBodyStats(b).designable, asBodyStats(b).name_, "Body", failPart, failPartName)
            If designable Then
                designable = StrComp(asBodyStats(b).bodyClass, "Transports", vbBinaryCompare) = 0
            End If
        End If
    End If

    If Not designable Then
        debug_ 341, LOG_ERROR, "designableTemplate", "%s ""%s"" for ""%s"" cannot be designed", failPart, failPartName, psTempl.name_
    End If

    designableTemplate = designable
End Function

Public Function initTemplates() As Boolean
    Dim ini As WzConfig, version As Long, design As DROID_TEMPLATE, blank As DROID_TEMPLATE, loadCommonSuccess As Boolean
    Dim valid As Boolean, psDestTemplate As Long, k As Long

    If Not ULT(selectedPlayer, MAX_PLAYERS) Then initTemplates = False: Exit Function

    Set ini = WzConfig_new("userdata/" & rulesettag & "/templates.json", WzConfig_ReadOnly)
    If Not ini.status() Then
        debug_ 354, LOG_WZ, "initTemplates", "Could not open %s", ini.fileName()
        initTemplates = False
        Exit Function
    End If
    version = ini.value("version", json_variant_new(0)).toInt()
    If version = 0 Then
        initTemplates = True ' too old version
        Exit Function
    End If
    ini.beginArray "templates"
    Do While ini.remainingArrayItems() > 0
        design = blank
        DROID_TEMPLATE_construct design
        loadCommonSuccess = loadTemplateCommon_ini(ini, design)
        design.multiPlayerID = generateNewObjectId()
        design.prefab = False       ' not AI template
        design.stored = True

        If Not loadCommonSuccess Then
            debug_ 372, LOG_ERROR, "initTemplates", "Stored template ""%s"" contains an unknown component.", design.name_
            GoTo continue_
        End If
        valid = intValidTemplate(design, ini.value("name").toWzString(), False, selectedPlayer)
        If Not valid Then
            debug_ 378, LOG_ERROR, "initTemplates", "Invalid template ""%s"" from stored templates", design.name_
            GoTo continue_
        End If
        psDestTemplate = 0
        For k = 0 To droidTemplates(selectedPlayer).size - 1
            psDestTemplate = droidTemplates(selectedPlayer).values(k)
            ' Check if template is identical to a loaded template
            If apsTemplates(psDestTemplate).droidType = design.droidType _
                And StrComp(apsTemplates(psDestTemplate).name_, design.name_, vbBinaryCompare) = 0 _
                And apsTemplates(psDestTemplate).numWeaps = design.numWeaps _
                And apsTemplates(psDestTemplate).asWeaps(0) = design.asWeaps(0) _
                And apsTemplates(psDestTemplate).asWeaps(1) = design.asWeaps(1) _
                And apsTemplates(psDestTemplate).asWeaps(2) = design.asWeaps(2) _
                And apsTemplates(psDestTemplate).asParts(COMP_BODY) = design.asParts(COMP_BODY) _
                And apsTemplates(psDestTemplate).asParts(COMP_PROPULSION) = design.asParts(COMP_PROPULSION) _
                And apsTemplates(psDestTemplate).asParts(COMP_REPAIRUNIT) = design.asParts(COMP_REPAIRUNIT) _
                And apsTemplates(psDestTemplate).asParts(COMP_ECM) = design.asParts(COMP_ECM) _
                And apsTemplates(psDestTemplate).asParts(COMP_SENSOR) = design.asParts(COMP_SENSOR) _
                And apsTemplates(psDestTemplate).asParts(COMP_CONSTRUCT) = design.asParts(COMP_CONSTRUCT) _
                And apsTemplates(psDestTemplate).asParts(COMP_BRAIN) = design.asParts(COMP_BRAIN) Then
                Exit For
            End If
            psDestTemplate = 0
        Next
        If psDestTemplate <> 0 Then
            apsTemplates(psDestTemplate).stored = True ' assimilate it
            GoTo continue_ ' next!
        End If
        design.enabled = allowDesign
        copyTemplate selectedPlayer, design
        localTemplates_push_back design
continue_:
        ini.nextArrayItem
    Loop
    ini.endArray
    initTemplates = True
End Function

Public Function saveTemplateCommon(ByVal psCurr As Long) As nlohmann_json
    Dim templateObj As nlohmann_json, weapons As nlohmann_json, j As Long, s As String
    Set templateObj = json_object()
    templateObj.set_item "name", json_value(apsTemplates(psCurr).name_)
    s = vbNullString
    Select Case apsTemplates(psCurr).droidType
    Case DROID_ECM: s = "ECM"
    Case DROID_SENSOR: s = "SENSOR"
    Case DROID_CONSTRUCT: s = "CONSTRUCT"
    Case DROID_WEAPON: s = "WEAPON"
    Case DROID_PERSON: s = "PERSON"
    Case DROID_CYBORG: s = "CYBORG"
    Case DROID_CYBORG_SUPER: s = "CYBORG_SUPER"
    Case DROID_CYBORG_CONSTRUCT: s = "CYBORG_CONSTRUCT"
    Case DROID_CYBORG_REPAIR: s = "CYBORG_REPAIR"
    Case DROID_TRANSPORTER: s = "TRANSPORTER"
    Case DROID_SUPERTRANSPORTER: s = "SUPERTRANSPORTER"
    Case DROID_COMMAND: s = "DROID_COMMAND"
    Case DROID_REPAIR: s = "REPAIR"
    Case DROID_DEFAULT: s = "DROID"
    Case Else: ASSERT_FAILURE "false", "src/template.cpp:437", "saveTemplateCommon", "No such droid type ""%d"" for %s", apsTemplates(psCurr).droidType, apsTemplates(psCurr).name_
    End Select
    If StrPtr(s) <> 0 Then templateObj.set_item "type", json_value(s)
    If Not (ULT(apsTemplates(psCurr).asParts(COMP_BODY), asBodyStats_size)) Then ASSERT_FAILURE "psCurr->asParts[COMP_BODY] < asBodyStats.size()", "src/template.cpp:439", "saveTemplateCommon", "asParts[COMP_BODY] (%d) exceeds numBodyStats (%zu)", apsTemplates(psCurr).asParts(COMP_BODY), asBodyStats_size
    templateObj.set_item "body", json_value(asBodyStats(DROID_TEMPLATE_getBodyStats(apsTemplates(psCurr))).id)
    If Not (ULT(apsTemplates(psCurr).asParts(COMP_PROPULSION), asPropulsionStats_size)) Then ASSERT_FAILURE "psCurr->asParts[COMP_PROPULSION] < asPropulsionStats.size()", "src/template.cpp:441", "saveTemplateCommon", "asParts[COMP_PROPULSION] (%d) exceeds numPropulsionStats (%zu)", apsTemplates(psCurr).asParts(COMP_PROPULSION), asPropulsionStats_size
    templateObj.set_item "propulsion", json_value(asPropulsionStats(DROID_TEMPLATE_getPropulsionStats(apsTemplates(psCurr))).id)
    If apsTemplates(psCurr).asParts(COMP_BRAIN) <> 0 Then
        If Not (ULT(apsTemplates(psCurr).asParts(COMP_BRAIN), asBrainStats_size)) Then ASSERT_FAILURE "psCurr->asParts[COMP_BRAIN] < asBrainStats.size()", "src/template.cpp:445", "saveTemplateCommon", "asParts[COMP_BRAIN] (%d) exceeds numBrainStats (%zu)", apsTemplates(psCurr).asParts(COMP_BRAIN), asBrainStats_size
        templateObj.set_item "brain", json_value(asBrainStats(DROID_TEMPLATE_getBrainStats(apsTemplates(psCurr))).id)
    End If
    If asRepairStats(DROID_TEMPLATE_getRepairStats(apsTemplates(psCurr))).location = LOC_TURRET Then ' avoid auto-repair...
        If Not (ULT(apsTemplates(psCurr).asParts(COMP_REPAIRUNIT), asRepairStats_size)) Then ASSERT_FAILURE "psCurr->asParts[COMP_REPAIRUNIT] < asRepairStats.size()", "src/template.cpp:450", "saveTemplateCommon", "asParts[COMP_REPAIRUNIT] (%d) exceeds numRepairStats (%zu)", apsTemplates(psCurr).asParts(COMP_REPAIRUNIT), asRepairStats_size
        templateObj.set_item "repair", json_value(asRepairStats(DROID_TEMPLATE_getRepairStats(apsTemplates(psCurr))).id)
    End If
    If asECMStats(DROID_TEMPLATE_getECMStats(apsTemplates(psCurr))).location = LOC_TURRET Then
        If Not (ULT(apsTemplates(psCurr).asParts(COMP_ECM), asECMStats_size)) Then ASSERT_FAILURE "psCurr->asParts[COMP_ECM] < asECMStats.size()", "src/template.cpp:455", "saveTemplateCommon", "asParts[COMP_ECM] (%d) exceeds numECMStats (%zu)", apsTemplates(psCurr).asParts(COMP_ECM), asECMStats_size
        templateObj.set_item "ecm", json_value(asECMStats(DROID_TEMPLATE_getECMStats(apsTemplates(psCurr))).id)
    End If
    If asSensorStats(DROID_TEMPLATE_getSensorStats(apsTemplates(psCurr))).location = LOC_TURRET Then
        If Not (ULT(apsTemplates(psCurr).asParts(COMP_SENSOR), asSensorStats_size)) Then ASSERT_FAILURE "psCurr->asParts[COMP_SENSOR] < asSensorStats.size()", "src/template.cpp:460", "saveTemplateCommon", "asParts[COMP_SENSOR] (%d) exceeds numSensorStats (%zu)", apsTemplates(psCurr).asParts(COMP_SENSOR), asSensorStats_size
        templateObj.set_item "sensor", json_value(asSensorStats(DROID_TEMPLATE_getSensorStats(apsTemplates(psCurr))).id)
    End If
    If apsTemplates(psCurr).asParts(COMP_CONSTRUCT) <> 0 Then
        If Not (ULT(apsTemplates(psCurr).asParts(COMP_CONSTRUCT), asConstructStats_size)) Then ASSERT_FAILURE "psCurr->asParts[COMP_CONSTRUCT] < asConstructStats.size()", "src/template.cpp:465", "saveTemplateCommon", "asParts[COMP_CONSTRUCT] (%d) exceeds numConstructStats (%zu)", apsTemplates(psCurr).asParts(COMP_CONSTRUCT), asConstructStats_size
        templateObj.set_item "construct", json_value(asConstructStats(DROID_TEMPLATE_getConstructStats(apsTemplates(psCurr))).id)
    End If
    Set weapons = json_array()
    For j = 0 To apsTemplates(psCurr).numWeaps - 1
        If Not (ULT(apsTemplates(psCurr).asWeaps(j), asWeaponStats_size)) Then ASSERT_FAILURE "psCurr->asWeaps[j] < asWeaponStats.size()", "src/template.cpp:471", "saveTemplateCommon", "psCurr->asWeaps[%d] (%d) exceeds numWeaponStats (%zu)", j, apsTemplates(psCurr).asWeaps(j), asWeaponStats_size
        weapons.push_back json_value(asWeaponStats(DROID_TEMPLATE_getWeaponStats(apsTemplates(psCurr), j)).id)
    Next
    If Not weapons.empty_() Then
        templateObj.set_item_ref "weapons", weapons
    End If
    Set saveTemplateCommon = templateObj
End Function

Public Function storeTemplates() As Boolean
    Dim ini As WzConfig, k As Long, psCurr As Long

    If Not ULT(selectedPlayer, MAX_PLAYERS) Then storeTemplates = False: Exit Function

    ' Write stored templates (back) to file
    Set ini = WzConfig_new("userdata/" & rulesettag & "/templates.json", WzConfig_ReadAndWrite)
    If Not ini.status() Or Not ini.isWritable() Then
        debug_ 489, LOG_ERROR, "storeTemplates", "Could not open %s", ini.fileName()
        storeTemplates = False
        Exit Function
    End If
    ini.setValue "version", json_value(1) ' for breaking backwards compatibility in a nice way
    ini.beginArray "templates"
    For k = 0 To droidTemplates(selectedPlayer).size - 1
        psCurr = droidTemplates(selectedPlayer).values(k)
        If apsTemplates(psCurr).stored Then
            ini.currentJsonValue().assign saveTemplateCommon(psCurr)
            ini.nextArrayItem
        End If
    Next
    ini.endArray
    storeTemplates = True
End Function

Public Function shutdownTemplates() As Boolean
    shutdownTemplates = storeTemplates()
End Function

' DROID_TEMPLATE::getBodyStats() etc.: the component's index in its array
Public Function DROID_TEMPLATE_getBodyStats(t As DROID_TEMPLATE) As Long
    DROID_TEMPLATE_getBodyStats = t.asParts(COMP_BODY)
End Function

Public Function DROID_TEMPLATE_getBrainStats(t As DROID_TEMPLATE) As Long
    DROID_TEMPLATE_getBrainStats = t.asParts(COMP_BRAIN)
End Function

Public Function DROID_TEMPLATE_getPropulsionStats(t As DROID_TEMPLATE) As Long
    DROID_TEMPLATE_getPropulsionStats = t.asParts(COMP_PROPULSION)
End Function

Public Function DROID_TEMPLATE_getSensorStats(t As DROID_TEMPLATE) As Long
    DROID_TEMPLATE_getSensorStats = t.asParts(COMP_SENSOR)
End Function

Public Function DROID_TEMPLATE_getECMStats(t As DROID_TEMPLATE) As Long
    DROID_TEMPLATE_getECMStats = t.asParts(COMP_ECM)
End Function

Public Function DROID_TEMPLATE_getRepairStats(t As DROID_TEMPLATE) As Long
    DROID_TEMPLATE_getRepairStats = t.asParts(COMP_REPAIRUNIT)
End Function

Public Function DROID_TEMPLATE_getConstructStats(t As DROID_TEMPLATE) As Long
    DROID_TEMPLATE_getConstructStats = t.asParts(COMP_CONSTRUCT)
End Function

Public Function DROID_TEMPLATE_getWeaponStats(t As DROID_TEMPLATE, ByVal weaponSlot As Long) As Long
    DROID_TEMPLATE_getWeaponStats = t.asWeaps(weaponSlot)
End Function

Public Function loadDroidTemplates(ByVal filename As String) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, design As DROID_TEMPLATE, blank As DROID_TEMPLATE
    Dim available_ As Boolean, droidResourceName As String, playerIdx As Long, it As Long, prefabStr As String

    Set ini = WzConfig_new(filename, WzConfig_ReadOnlyAndRequired)
    n = ini.childGroups(list)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        design = blank
        DROID_TEMPLATE_construct design
        If Not loadTemplateCommon_ini(ini, design) Then
            debug_ 577, LOG_ERROR, "loadDroidTemplates", "Stored template ""%s"" contains an unknown component.", ini.string_("name")
            GoTo continue_                  ' (the group stays open, as in C++)
        End If
        design.id = list(i)
        design.name_ = ini.string_("name")
        design.multiPlayerID = generateNewObjectId()
        design.prefab = True
        design.stored = False
        design.enabled = True
        available_ = ini.value("available", json_variant_new(False)).toBool()
        droidResourceName = getDroidResourceName(list(i))
        If StrPtr(droidResourceName) <> 0 Then
            design.name_ = droidResourceName
        Else
            design.name_ = GetDefaultTemplateName(design)
        End If
        ini.endGroup

        For playerIdx = 0 To MAX_PLAYERS - 1
            ' Give those meant for humans to all human players.
            If NetPlay_players_allocated(playerIdx) And available_ Then
                design.prefab = False
                copyTemplate playerIdx, design

                ' This sets up the UI templates for display purposes ONLY--we still only use droidTemplates for making them.
                ' FIXME: Why are we doing this here, and not on demand ?
                ' Only add unique designs to the UI list (Note, perhaps better to use std::map instead?)
                For it = 0 To localTemplates_size - 1
                    If apsTemplates(localTemplates(it)).multiPlayerID = design.multiPlayerID Then
                        debug_ 608, LOG_WARNING, "loadDroidTemplates", "Design id:%d (%s) *NOT* added to UI list (duplicate), player= %d", design.multiPlayerID, design.name_, playerIdx
                        Exit For
                    End If
                Next
                If it = localTemplates_size Then
                    debug_ 614, LOG_NEVER, "loadDroidTemplates", "Design id:%d (%s) added to UI list, player =%d", design.multiPlayerID, design.name_, playerIdx
                    localTemplates_push_back design
                End If
            ElseIf Not NetPlay_players_allocated(playerIdx) Then ' AI template
                design.prefab = True ' prefabricated templates referenced from VLOs
                copyTemplate playerIdx, design
            End If
        Next
        If design.prefab Then prefabStr = "yes" Else prefabStr = "no"
        debug_ 624, LOG_NEVER, "loadDroidTemplates", "Droid template found, Name: %s, MP ID: %d, ref: %u, ID: %s, prefab: %s, type:%d (loading)", _
            design.name_, design.multiPlayerID, design.ref, design.id, prefabStr, design.droidType
continue_:
    Next

    loadDroidTemplates = True
End Function

' Make a duplicate of template given by pointer and store it. Then return pointer to copy.
Public Function copyTemplate(ByVal player As Long, psTemplate As DROID_TEMPLATE) As Long
    Dim dup As Long, t As DROID_TEMPLATE
    t = psTemplate                          ' (VB) before the pool may grow: psTemplate may be in it
    dup = newTemplate()
    apsTemplates(dup) = t
    copyTemplate = addTemplate(player, dup)
End Function

' Take ownership of template given by pointer.
' Returns a new usable DROID_TEMPLATE *
Public Function addTemplate(ByVal player As Long, ByVal psTemplate As Long) As Long
    Dim multiPlayerID As Long, k As Long, m As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:639", "addTemplate", "Invalid player: %u", player: addTemplate = 0: Exit Function
    multiPlayerID = apsTemplates(psTemplate).multiPlayerID
    k = FindTemplateKey(droidTemplates(player), multiPlayerID)
    If k < droidTemplates(player).size Then
        If droidTemplates(player).keys(k) = multiPlayerID Then
            ' replacing existing template
            m = replacedDroidTemplates(player).size
            ReDim Preserve replacedDroidTemplates(player).items(0 To m)
            replacedDroidTemplates(player).items(m) = droidTemplates(player).values(k)
            replacedDroidTemplates(player).size = m + 1
            droidTemplates(player).values(k) = psTemplate
            addTemplate = psTemplate
            Exit Function
        End If
    End If
    ' new template
    m = droidTemplates(player).size
    ReDim Preserve droidTemplates(player).keys(0 To m)
    ReDim Preserve droidTemplates(player).values(0 To m)
    Do While m > k
        droidTemplates(player).keys(m) = droidTemplates(player).keys(m - 1)
        droidTemplates(player).values(m) = droidTemplates(player).values(m - 1)
        m = m - 1
    Loop
    droidTemplates(player).keys(k) = multiPlayerID
    droidTemplates(player).values(k) = psTemplate
    droidTemplates(player).size = droidTemplates(player).size + 1
    addTemplate = psTemplate
End Function

' (VB) std::map::lower_bound: the first position whose key is not below id (unsigned)
Private Function FindTemplateKey(m As TEMPLATE_MAP, ByVal id As Long) As Long
    Dim lo As Long, hi As Long, mid_ As Long
    lo = 0
    hi = m.size
    Do While lo < hi
        mid_ = (lo + hi) \ 2
        If ULT(m.keys(mid_), id) Then lo = mid_ + 1 Else hi = mid_
    Loop
    FindTemplateKey = lo
End Function

' (VB) enumerateTemplates: the i-th template of droidTemplates[player] in the map's order
Public Function droidTemplates_at(ByVal player As Long, ByVal i As Long) As Long
    droidTemplates_at = droidTemplates(player).values(i)
End Function

Public Function findPlayerTemplateById(ByVal player As Long, ByVal templateId As Long) As Long
    Dim k As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:671", "findPlayerTemplateById", "Invalid player: %u", player: findPlayerTemplateById = 0: Exit Function
    k = FindTemplateKey(droidTemplates(player), templateId)
    If k < droidTemplates(player).size Then
        If droidTemplates(player).keys(k) = templateId Then
            findPlayerTemplateById = droidTemplates(player).values(k)
            Exit Function
        End If
    End If
    findPlayerTemplateById = 0
End Function

Public Function templateCount(ByVal player As Long) As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:682", "templateCount", "Invalid player: %u", player: templateCount = 0: Exit Function
    templateCount = droidTemplates(player).size
End Function

Public Sub clearTemplates(ByVal player As Long)
    Dim k As Long
    If Not (player >= 0 And player < MAX_PLAYERS) Then ASSERT_FAILURE "player >= 0 && player < MAX_PLAYERS", "src/template.cpp:688", "clearTemplates", "Invalid player: %u", player: Exit Sub
    For k = 0 To droidTemplates(player).size - 1
        deleteTemplate droidTemplates(player).values(k)
    Next
    droidTemplates(player).size = 0
    Erase droidTemplates(player).keys
    Erase droidTemplates(player).values
    For k = 0 To replacedDroidTemplates(player).size - 1
        deleteTemplate replacedDroidTemplates(player).items(k)
    Next
    replacedDroidTemplates(player).size = 0
    Erase replacedDroidTemplates(player).items
End Sub

'free the storage for the droid templates
Public Function droidTemplateShutDown() As Boolean
    Dim player As Long
    For player = 0 To MAX_PLAYERS - 1
        clearTemplates player
    Next
    localTemplates_clear
    droidTemplateShutDown = True
End Function

' Get a static template from its name. This is used from scripts. These templates must
' never be changed or deleted.
' \param pName Template name
' \pre pName has to be the unique, untranslated name!
Public Function getTemplateFromTranslatedNameNoPlayer(ByVal pName As String) As Long
    Dim player As Long, k As Long
    For player = 0 To MAX_PLAYERS - 1
        For k = 0 To droidTemplates(player).size - 1
            If apsTemplates(droidTemplates(player).values(k)).id = pName Then
                getTemplateFromTranslatedNameNoPlayer = droidTemplates(player).values(k)
                Exit Function
            End If
        Next
    Next
    getTemplateFromTranslatedNameNoPlayer = 0
End Function

'getTemplatefFromMultiPlayerID gets template for unique ID  searching all lists
Public Function getTemplateFromMultiPlayerID(ByVal multiPlayerID As Long) As Long
    Dim player As Long, h As Long
    For player = 0 To MAX_PLAYERS - 1
        h = findPlayerTemplateById(player, multiPlayerID)
        If h <> 0 Then
            getTemplateFromMultiPlayerID = h
            Exit Function
        End If
    Next
    getTemplateFromMultiPlayerID = 0
End Function

'called when a Template is deleted in the Design screen
Public Sub deleteTemplateFromProduction(ByVal psTemplate As Long, ByVal player As Long, ByVal mode As QUEUE_MODE)
    Dim i As Long, it As Long, psStruct As Long, psFactory As Long, psAssemblyPoint As Long, factoryType As Long, factoryInc As Long, inc As Long

    If Not (psTemplate <> 0) Then ASSERT_FAILURE "psTemplate != nullptr", "src/template.cpp:741", "deleteTemplateFromProduction", "Null psTemplate": Exit Sub
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/template.cpp:742", "deleteTemplateFromProduction", "Invalid player: %u", player: Exit Sub

    'see if any factory is currently using the template
    For i = 0 To 1
        ' StructureList* psList (VB: the list's first node)
        Select Case i
        Case 0
            it = apsStructLists(player).head
        Case 1
            it = mission_apsStructLists(player).head
        End Select
        Do While it <> 0
            psStruct = lnObj(it)
            If psStruct <> 0 Then
                If STRUCTURE_isFactory(psStruct) Then
                    If aStruct(psStruct).pFunctionality = 0 Then
                        GoTo continue_
                    End If
                    psFactory = psStruct    ' FACTORY *

                    psAssemblyPoint = aStruct(psFactory).functionality.factory.psAssemblyPoint
                    If psAssemblyPoint <> 0 Then
                        If aFlagPos(psAssemblyPoint).factoryType < NUM_FACTORY_TYPES Then
                            If aFlagPos(psAssemblyPoint).factoryInc < asProductionRun(aFlagPos(psAssemblyPoint).factoryType).size Then
                                ' ProductionRun &productionRun = asProductionRun[psFactory->psAssemblyPoint->factoryType][psFactory->psAssemblyPoint->factoryInc]
                                factoryType = aFlagPos(psAssemblyPoint).factoryType
                                factoryInc = aFlagPos(psAssemblyPoint).factoryInc
                                inc = 0
                                Do While inc < asProductionRun(factoryType).runs(factoryInc).size
                                    If asProductionRun(factoryType).runs(factoryInc).items(inc).psTemplate <> 0 Then
                                        If apsTemplates(asProductionRun(factoryType).runs(factoryInc).items(inc).psTemplate).multiPlayerID = apsTemplates(psTemplate).multiPlayerID And mode = ModeQueue Then
                                            'just need to erase this production run entry
                                            productionRun_erase asProductionRun(factoryType).runs(factoryInc), inc
                                            inc = inc - 1
                                        End If
                                    End If
                                    inc = inc + 1
                                Loop
                            End If
                        End If
                    End If

                    If aStruct(psFactory).functionality.factory.psSubject = 0 Then
                        GoTo continue_
                    End If

                    ' check not being built in the factory for the template player
                    If apsTemplates(psTemplate).multiPlayerID = apsTemplates(aStruct(psFactory).functionality.factory.psSubject).multiPlayerID And mode = ModeImmediate Then
                        syncDebugStructure_ "deleteTemplateFromProduction", psStruct, Asc("<")
                        syncDebug "deleteTemplateFromProduction", "Clearing production"

                        ' Clear the factory's subject, and returns power.
                        cancelProduction psStruct, ModeImmediate, False
                        ' Check to see if anything left to produce. (Also calls cancelProduction again, if nothing left to produce, which is a no-op. But if other things are left to produce, doesn't call cancelProduction, so wouldn't return power without the explicit cancelProduction call above.)
                        doNextProduction psStruct, 0, ModeImmediate

                        syncDebugStructure_ "deleteTemplateFromProduction", psStruct, Asc(">")
                    End If
                End If
            End If
continue_:
            it = lnNext(it)
        Loop
    Next
End Sub

' (VB) productionRun.erase(productionRun.begin() + inc)
Private Sub productionRun_erase(prodRun As ProductionRun, ByVal inc As Long)
    Dim i As Long
    For i = inc To prodRun.size - 2
        prodRun.items(i) = prodRun.items(i + 1)
    Next
    prodRun.size = prodRun.size - 1
End Sub

Public Sub checkPlayerBuiltHQ(ByVal psStruct As Long)
    If selectedPlayer = aStruct(psStruct).player And asStructureStats(aStruct(psStruct).pStructureType).typ = REF_HQ Then
        playerBuiltHQ = True
    End If
End Sub

' return whether a template is for an IDF droid
Public Function templateIsIDF(ByVal psTemplate As Long) As Boolean
    'add Cyborgs
    If Not (apsTemplates(psTemplate).droidType = DROID_WEAPON Or apsTemplates(psTemplate).droidType = DROID_CYBORG Or apsTemplates(psTemplate).droidType = DROID_CYBORG_SUPER) Then
        templateIsIDF = False
        Exit Function
    End If

    If proj_Direct(DROID_TEMPLATE_getWeaponStats(apsTemplates(psTemplate), 0)) Then
        templateIsIDF = False
        Exit Function
    End If

    templateIsIDF = True
End Function

'
'fills the list with Templates that can be manufactured
'in the Factory - based on size. There is a limit on how many can be manufactured
'at any one time.
'
' (VB) std::vector<DROID_TEMPLATE *>: the templates in pList(0 .. count - 1), the count as the value
Public Function fillTemplateList(ByVal psFactory As Long, pList() As Long) As Long
    Dim n As Long, player As Long, iCapacity As BODY_SIZE, k As Long, psCurr As Long, production As ProductionRunEntry
    Dim bodySize As BODY_SIZE

    Erase pList
    n = 0
    player = aStruct(psFactory).player

    iCapacity = aStruct(psFactory).capacity

    ' Add the templates to the list
    For k = 0 To localTemplates_size - 1
        psCurr = localTemplates(k)
        ' Must add droids if currently in production.
        production = getProduction(psFactory, psCurr)
        If production.quantity = 0 Then
            If Not apsTemplates(psCurr).enabled Then GoTo continue_
            If bMultiPlayer And Not playerBuiltHQ And (apsTemplates(psCurr).droidType <> DROID_CONSTRUCT And apsTemplates(psCurr).droidType <> DROID_CYBORG_CONSTRUCT) Then GoTo continue_
            If Not validTemplateForFactory(psCurr, psFactory, False) Then GoTo continue_
            If Not researchedTemplate(psCurr, player, includeRedundantDesigns) Then GoTo continue_
        End If

        'check the factory can cope with this sized body
        bodySize = asBodyStats(DROID_TEMPLATE_getBodyStats(apsTemplates(psCurr))).size
        If bodySize <= iCapacity Then
            ReDim Preserve pList(0 To n)
            pList(n) = psCurr
            n = n + 1
        ElseIf bMultiPlayer And (iCapacity = SIZE_HEAVY) Then
            ' Special case for Super heavy bodyies (Super Transporter)
            If bodySize = SIZE_SUPER_HEAVY Then
                ReDim Preserve pList(0 To n)
                pList(n) = psCurr
                n = n + 1
            End If
        End If
continue_:
    Next

    fillTemplateList = n
End Function
