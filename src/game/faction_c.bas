Attribute VB_Name = "faction_c"
' src/faction.cpp, faction.h, factionid.h: the factions and their
' replacement models. A FACTION * is the faction's index in factions
' (which is its FactionID), so the two getFactionModelName overloads are
' one procedure; std::map<WzString, WzString> ReplaceIMDMap is a
' Scripting.Dictionary. optional<WzString> is the Boolean result and the
' string ByRef. getEnabledFactions returns the factions in id order
' (std::unordered_set: hash order). faction_c_init is the C++ static
' initialiser of factions. _() (gettext) is not ported:
' to_localized_string_FactionID gives the English names.
Option Explicit

' factionid.h: enum FactionID : uint8_t
Public Const FACTION_NORMAL As Long = 0
Public Const FACTION_NEXUS As Long = 1
Public Const FACTION_COLLECTIVE As Long = 2
Public Const MAX_FACTION_ID As Long = FACTION_COLLECTIVE

Public Const NUM_FACTIONS As Long = 3

Public Type FACTION
    name_ As String
    replaceIMD As Object                ' ReplaceIMDMap: Scripting.Dictionary (BinaryCompare) of normal model -> faction model
End Type

Private factions(0 To NUM_FACTIONS - 1) As FACTION

' factionid.h: optional<FactionID> uintToFactionID(uint8_t value): false for nullopt
Public Function uintToFactionID(ByVal value As Long, factionId As Long) As Boolean
    If (value And &HFF&) > MAX_FACTION_ID Then
        uintToFactionID = False
        Exit Function
    End If
    factionId = value And &HFF&
    uintToFactionID = True
End Function

' (VB) a FACTION { name, { {normal, replacement}, ... } } of the mapping table
Private Function newFaction(ByVal name_ As String, ByVal pairs As Variant) As FACTION
    Dim f As FACTION, i As Long
    f.name_ = name_
    Set f.replaceIMD = CreateObject("Scripting.Dictionary")
    f.replaceIMD.CompareMode = vbBinaryCompare
    For i = 0 To UBound(pairs) - 1 Step 2
        If Not f.replaceIMD.Exists(pairs(i)) Then f.replaceIMD.Add pairs(i), pairs(i + 1)
    Next
    newFaction = f
End Function

' getInitialFactionsMappingTable, into factions
Private Sub getInitialFactionsMappingTable()
    factions(0) = newFaction("Normal", Array())
    factions(1) = newFaction("NEXUS", Array( _
        "blwallc1.pie", "blwallc_nex.pie", _
        "blwallh.pie", "blwallh_nex.pie", _
        "blwallh_t.pie", "blwallh_t_nex.pie", _
        "blwallh_l.pie", "blwallh_l_nex.pie", _
        "blhq.pie", "blhq_nex.pie", _
        "blguard1.pie", "blguard_nex.pie", _
        "blguardr.pie", "blguardr_nex.pie", _
        "blresch0.pie", "blresch0_nex.pie", _
        "blresch4.pie", "blresch4_nex.pie", _
        "research_module4.pie", "research_module4_nex.pie", _
        "blbresch.pie", "blbresch_nex.pie", _
        "blbpower.pie", "blbpower_nex.pie", _
        "blbhq.pie", "blbhq_nex.pie", _
        "blbdrdcm.pie", "blbdrdcm_nex.pie", _
        "blpower0.pie", "blpower0_nex.pie", _
        "blpower4.pie", "blpower4_nex.pie", _
        "blpilbox.pie", "blpilbox_nex.pie", _
        "blfact0.pie", "blfact0_nex.pie", _
        "factory_module1.pie", "factory_module1_nex.pie", _
        "blfact1.pie", "blfact1_nex.pie", _
        "factory_module2.pie", "factory_module2_nex.pie", _
        "blfact2.pie", "blfact2_nex.pie") _
        )
    AddPairs factions(1), Array( _
        "blbfact.pie", "blbfact_nex.pie", _
        "bldrdcm0.pie", "bldrdcm0_nex.pie", _
        "blcfact1.pie", "blcfact1_nex.pie", _
        "blbcfact.pie", "blbcfact_nex.pie", _
        "blcanpil.pie", "blcanpil_nex.pie", _
        "blhowmnt.pie", "blhowmnt_nex.pie", _
        "blguardm.pie", "blguardm_nex.pie", _
        "blvfact0.pie", "blvfact0_nex.pie", _
        "blvfact1.pie", "blvfact1_nex.pie", _
        "blvfact2.pie", "blvfact2_nex.pie", _
        "vtolfactory_module1.pie", "vtolfactory_module1_nex.pie", _
        "vtolfactory_module2.pie", "vtolfactory_module2_nex.pie", _
        "blderik_anim.pie", "blderik_anim_nex.pie", _
        "blderik.pie", "blderik_nex.pie", _
        "blmrtpit.pie", "blmrtpit_nex.pie", _
        "stwpfcan.pie", "stwpfcan_nex.pie")
    factions(2) = newFaction("Collective", Array( _
        "blwallc1.pie", "blwallc_col.pie", _
        "blwallh.pie", "blwallh_col.pie", _
        "blwallh_t.pie", "blwallh_t_col.pie", _
        "blwallh_l.pie", "blwallh_l_col.pie", _
        "blhq.pie", "blhq_col.pie", _
        "blguard1.pie", "blguard_col.pie", _
        "blguardr.pie", "blguardr_col.pie", _
        "blresch0.pie", "blresch0_col.pie", _
        "blresch4.pie", "blresch4_col.pie", _
        "research_module4.pie", "research_module4_col.pie", _
        "blbresch.pie", "blbresch_col.pie", _
        "blbpower.pie", "blbpower_col.pie", _
        "blbhq.pie", "blbhq_col.pie", _
        "blbdrdcm.pie", "blbdrdcm_col.pie", _
        "blpower0.pie", "blpower0_col.pie", _
        "blpower4.pie", "blpower4_col.pie", _
        "blpilbox.pie", "blpilbox_col.pie", _
        "blfact0.pie", "blfact0_col.pie", _
        "factory_module1.pie", "factory_module1_col.pie", _
        "blfact1.pie", "blfact1_col.pie", _
        "factory_module2.pie", "factory_module2_col.pie", _
        "blfact2.pie", "blfact2_col.pie") _
        )
    AddPairs factions(2), Array( _
        "blbfact.pie", "blbfact_col.pie", _
        "bldrdcm0.pie", "bldrdcm0_col.pie", _
        "blcfact1.pie", "blcfact1_col.pie", _
        "blbcfact.pie", "blbcfact_col.pie", _
        "blcanpil.pie", "blcanpil_col.pie", _
        "blvfact0.pie", "blvfact0_col.pie", _
        "blvfact1.pie", "blvfact1_col.pie", _
        "blvfact2.pie", "blvfact2_col.pie", _
        "vtolfactory_module1.pie", "vtolfactory_module1_col.pie", _
        "vtolfactory_module2.pie", "vtolfactory_module2_col.pie", _
        "blderik_anim.pie", "blderik_anim_col.pie", _
        "blderik.pie", "blderik_col.pie", _
        "blmrtpit.pie", "blmrtpit_col.pie", _
        "stwpfcan.pie", "stwpfcan_col.pie")
    '			{"blhowmnt.pie", "blhowmnt_col.pie"},
    '			{"blguardm.pie", "blguardm_col.pie"},
End Sub

' (VB) the rest of a faction's table (a statement allows 24 line continuations)
Private Sub AddPairs(f As FACTION, ByVal pairs As Variant)
    Dim i As Long
    For i = 0 To UBound(pairs) - 1 Step 2
        If Not f.replaceIMD.Exists(pairs(i)) Then f.replaceIMD.Add pairs(i), pairs(i + 1)
    Next
End Sub

Public Sub faction_c_init()
    getInitialFactionsMappingTable
End Sub

Public Sub reinitFactionsMapping()
    getInitialFactionsMappingTable
End Sub

' faction_: a FACTION * (its index in factions) or a FactionID; the faction model's name in factionModelName
Public Function getFactionModelName(ByVal faction_ As Long, ByVal normalFactionName As String, factionModelName As String) As Boolean
    If Not factions(faction_).replaceIMD.Exists(normalFactionName) Then
        getFactionModelName = False
        Exit Function
    Else
        factionModelName = factions(faction_).replaceIMD(normalFactionName)
        getFactionModelName = True
    End If
End Function

' DISPLAY ONLY: imd an iIMDShape *
Public Function getFactionDisplayIMD(ByVal faction_ As Long, ByVal imd As Long) As Long
    Dim factionModelName As String, baseModel As Long
    If Not getFactionModelName(faction_, modelName(imd), factionModelName) Then
        getFactionDisplayIMD = imd
        Exit Function
    Else
        baseModel = modelGet(factionModelName)
        If baseModel <> 0 Then getFactionDisplayIMD = iIMDBaseShape_displayModel(baseModel) Else getFactionDisplayIMD = imd
    End If
End Function

' a FACTION *
Public Function getPlayerFaction(ByVal player As Long) As Long
    getPlayerFaction = NetPlay_players_faction(player And &HFF&)
End Function

' a FACTION *
Public Function getFactionByID(ByVal faction_ As Long) As Long
    getFactionByID = faction_ And &HFF&
End Function


' (VB) const FACTION *faction: faction->name
Public Function FACTION_name(ByVal faction_ As Long) As String
    FACTION_name = factions(faction_).name_
End Function

' the factions in enabledFactions(0 .. count - 1)
Public Function getEnabledFactions(ByVal ignoreNormalFaction As Boolean, enabledFactions() As Long) As Long
    Dim f_id As Long, faction_ As Long, player As Long, count As Long, k As Long, found_ As Boolean
    ReDim enabledFactions(0 To NUM_FACTIONS - 1)
    For f_id = 0 To NUM_FACTIONS - 1
        faction_ = f_id
        If ignoreNormalFaction And (faction_ = FACTION_NORMAL) Then
            GoTo continue_
        End If
        For player = 0 To NetPlay_players_size() - 1
            If NetPlay_players_faction(player) = faction_ Then
                ' enabledFactions.insert(faction)
                found_ = False
                For k = 0 To count - 1
                    If enabledFactions(k) = faction_ Then found_ = True
                Next
                If Not found_ Then
                    enabledFactions(count) = faction_
                    count = count + 1
                End If
            End If
        Next
continue_:
    Next
    getEnabledFactions = count
End Function

Public Function to_string_FactionID(ByVal faction_ As Long) As String
    Select Case faction_
        Case FACTION_NORMAL
            to_string_FactionID = "Normal"
            Exit Function
        Case FACTION_NEXUS
            to_string_FactionID = "NEXUS"
            Exit Function
        Case FACTION_COLLECTIVE
            to_string_FactionID = "Collective"
            Exit Function
    End Select
    to_string_FactionID = "" ' silence warning - switch above should be complete
End Function

Public Function to_localized_string_FactionID(ByVal faction_ As Long) As String
    Select Case faction_
        Case FACTION_NORMAL
            ' TRANSLATORS: "Normal" Faction
            to_localized_string_FactionID = "Normal"
            Exit Function
        Case FACTION_NEXUS
            ' TRANSLATORS: "NEXUS" Faction
            to_localized_string_FactionID = "NEXUS"
            Exit Function
        Case FACTION_COLLECTIVE
            ' TRANSLATORS: "Collective" Faction
            to_localized_string_FactionID = "Collective"
            Exit Function
    End Select
    to_localized_string_FactionID = "" ' silence warning - switch above should be complete
End Function
