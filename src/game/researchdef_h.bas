Attribute VB_Name = "researchdef_h"
' src/researchdef.h: the research structs. RESEARCH's base class
' BASE_STATS gives its first fields (see statsdef_h); RESEARCH declares its
' own int index, which hides BASE_STATS::index, so BASE_STATS's is
' base_index here (0: the loader sets only RESEARCH's). A RESEARCH * is its
' index in asResearch, -1 = nullptr; COMPONENT_STATS * and BASE_STATS *
' are stats handles (stats_c), a VIEWDATA * a handle (0 = nullptr). The
' std::vectors are arrays with a _size. The inline functions take the
' PLAYER_RESEARCH ByRef. Not ported yet: RecursivelyDisableResearchByFlags.
Option Explicit

' Research struct type definitions
Public Enum TECH_CODE
    TC_MAJOR
    TC_MINOR
End Enum

Public Type RES_COMP_REPLACEMENT
    pOldComponent As Long               ' COMPONENT_STATS *
    pNewComponent As Long               ' COMPONENT_STATS *
End Type

Public Type RESEARCH
    ' BASE_STATS
    id As String                        ' Text id (i.e. short language-independent name)
    name_ As String                     ' Full / real name of the item
    ref As Long                         ' Unique ID of the item
    base_index As Long                  ' BASE_STATS::index (hidden by RESEARCH::index)

    techCode As Long                    ' UBYTE
    subGroup As Long                    ' UWORD: Subgroup of the item - an iconID from 'Framer' to depict in the button

    researchPoints As Long              ' UWORD: Number of research points required to complete the research
    researchPower As Long               ' Power cost to research
    keyTopic As Long                    ' UBYTE: Flag to indicate whether in single player this topic must be explicitly enabled
    excludeFromCheats As Long           ' UBYTE: Flag to prevent this topic from appearing and being researched with cheats
    disabledWhen As Long                ' UBYTE: flags when to disable tech
    pPRList() As Long                   ' UWORD: List of research pre-requisites
    pPRList_size As Long
    pStructList() As Long               ' UWORD: List of structures that when built would enable this research
    pStructList_size As Long
    pRedStructs() As Long               ' UWORD: List of Structures that become redundant
    pRedStructs_size As Long
    pRedArtefacts() As Long             ' COMPONENT_STATS *: List of Artefacts that become redundant
    pRedArtefacts_size As Long
    pStructureResults() As Long         ' UWORD: List of Structures that are possible after this research
    pStructureResults_size As Long
    componentResults() As Long          ' COMPONENT_STATS *: List of Components that are possible after this research
    componentResults_size As Long
    componentReplacement() As RES_COMP_REPLACEMENT  ' List of Components that are automatically replaced with new onew after research
    componentReplacement_size As Long
    results As nlohmann_json            ' Research upgrades
    pViewData As Long                   ' VIEWDATA *: Data used to display a message in the Intelligence Screen
    iconID As Long                      ' UWORD: the ID from 'Framer' for which graphic to draw in interface
    psStat As Long                      ' BASE_STATS *: A stat used to define which graphic is drawn instead of the two fields below
    pIMD As Long                        ' the IMD to draw for this research topic
    pIMD2 As Long                       ' the 2nd IMD for base plates/turrets
    category As String                  ' Category name as in json (for progression numbering)
    categoryProgress As Long            ' uint16_t: Category this/max (gets filled on load)
    categoryMax As Long                 ' uint16_t: Category progress/this (gets filled on load)
    index As Long                       ' Unique index for this research, set incrementally
End Type

Public Type PLAYER_RESEARCH
    currentPoints As Long               ' If the research has been suspended then this value contains the number of points generated at the suspension/cancel point
    ' normally it is null

    ResearchStatus As Long              ' UBYTE: Bit flags   ...  see below

    possible As Long                    ' UBYTE: is the research possible ... so can enable topics vis scripts
End Type

Public Const STARTED_RESEARCH As Long = &H1             ' research in progress
Public Const CANCELLED_RESEARCH As Long = &H2           ' research has been canceled
Public Const RESEARCHED As Long = &H4                   ' research is complete
Public Const CANCELLED_RESEARCH_PENDING As Long = &H8   ' research almost cancelled, waiting for GAME_RESEARCHSTATUS message to be processed.
Public Const STARTED_RESEARCH_PENDING As Long = &H10    ' research almost in progress, waiting for GAME_RESEARCHSTATUS message to be processed.
Public Const RESBITS As Long = STARTED_RESEARCH Or CANCELLED_RESEARCH Or RESEARCHED
Public Const RESBITS_PENDING_ONLY As Long = STARTED_RESEARCH_PENDING Or CANCELLED_RESEARCH_PENDING
Public Const RESBITS_ALL As Long = RESBITS Or RESBITS_PENDING_ONLY

Public Const RESEARCH_IMPOSSIBLE As Long = &H0          ' research is (temporarily) not possible
Public Const RESEARCH_POSSIBLE As Long = &H1            ' research is possible
Public Const RESEARCH_DISABLED As Long = &H2            ' research is disabled (e.g. most VTOL research in no-VTOL games)

' (VB) RESEARCH's default constructor: RESEARCH() : pViewData(nullptr), iconID(0), psStat(nullptr), pIMD(nullptr), pIMD2(nullptr)
' (the fields without an initialiser are the loader's to set)
Public Sub RESEARCH_construct(r As RESEARCH)
    r.pViewData = 0
    r.iconID = 0
    r.psStat = 0
    r.pIMD = 0
    r.pIMD2 = 0
End Sub

Public Function IsResearchPossible(research_ As PLAYER_RESEARCH) As Boolean
    IsResearchPossible = research_.possible = RESEARCH_POSSIBLE
End Function

Public Function IsResearchDisabled(research_ As PLAYER_RESEARCH) As Boolean
    IsResearchDisabled = research_.possible = RESEARCH_DISABLED
End Function

Public Sub MakeResearchPossible(research_ As PLAYER_RESEARCH)
    If research_.possible = RESEARCH_IMPOSSIBLE Then
        research_.possible = RESEARCH_POSSIBLE
    End If
End Sub

Public Sub DisableResearch(research_ As PLAYER_RESEARCH)
    research_.possible = RESEARCH_DISABLED
End Sub

Public Function GetResearchPossible(research_ As PLAYER_RESEARCH) As Long
    GetResearchPossible = research_.possible
End Function

Public Sub SetResearchPossible(research_ As PLAYER_RESEARCH, ByVal possible As Long)
    research_.possible = possible And &HFF&
End Sub

Public Function IsResearchCompleted(x As PLAYER_RESEARCH) As Boolean
    IsResearchCompleted = (x.ResearchStatus And RESEARCHED) <> 0
End Function

Public Function IsResearchCancelled(x As PLAYER_RESEARCH) As Boolean
    IsResearchCancelled = (x.ResearchStatus And CANCELLED_RESEARCH) <> 0
End Function

Public Function IsResearchStarted(x As PLAYER_RESEARCH) As Boolean
    IsResearchStarted = (x.ResearchStatus And STARTED_RESEARCH) <> 0
End Function

' Pending means not yet synchronised, so only permitted to affect the UI, not the game state.
Public Function IsResearchCancelledPending(x As PLAYER_RESEARCH) As Boolean
    If (x.ResearchStatus And RESBITS_PENDING_ONLY) <> 0 Then
        IsResearchCancelledPending = (x.ResearchStatus And CANCELLED_RESEARCH_PENDING) <> 0
    Else
        IsResearchCancelledPending = (x.ResearchStatus And CANCELLED_RESEARCH) <> 0
    End If
End Function

Public Function IsResearchStartedPending(x As PLAYER_RESEARCH) As Boolean
    If (x.ResearchStatus And RESBITS_PENDING_ONLY) <> 0 Then
        IsResearchStartedPending = (x.ResearchStatus And STARTED_RESEARCH_PENDING) <> 0
    Else
        IsResearchStartedPending = (x.ResearchStatus And STARTED_RESEARCH) <> 0
    End If
End Function

Public Sub MakeResearchCompleted(x As PLAYER_RESEARCH)
    x.ResearchStatus = x.ResearchStatus And Not RESBITS_ALL
    x.ResearchStatus = x.ResearchStatus Or RESEARCHED
End Sub

Public Sub MakeResearchCancelled(x As PLAYER_RESEARCH)
    x.ResearchStatus = x.ResearchStatus And Not RESBITS_ALL
    x.ResearchStatus = x.ResearchStatus Or CANCELLED_RESEARCH
End Sub

Public Sub MakeResearchStarted(x As PLAYER_RESEARCH)
    x.ResearchStatus = x.ResearchStatus And Not RESBITS_ALL
    x.ResearchStatus = x.ResearchStatus Or STARTED_RESEARCH
End Sub

' Pending means not yet synchronised, so only permitted to affect the UI, not the game state.
Public Sub MakeResearchCancelledPending(x As PLAYER_RESEARCH)
    x.ResearchStatus = x.ResearchStatus And Not RESBITS_PENDING_ONLY
    x.ResearchStatus = x.ResearchStatus Or CANCELLED_RESEARCH_PENDING
End Sub

Public Sub MakeResearchStartedPending(x As PLAYER_RESEARCH)
    x.ResearchStatus = x.ResearchStatus And Not RESBITS_PENDING_ONLY
    x.ResearchStatus = x.ResearchStatus Or STARTED_RESEARCH_PENDING
End Sub

Public Sub ResetPendingResearchStatus(x As PLAYER_RESEARCH)
    x.ResearchStatus = x.ResearchStatus And Not RESBITS_PENDING_ONLY
End Sub

' clear all bits in the status except for the possible bit
Public Sub ResetResearchStatus(x As PLAYER_RESEARCH)
    x.ResearchStatus = x.ResearchStatus And Not RESBITS_ALL
End Sub
