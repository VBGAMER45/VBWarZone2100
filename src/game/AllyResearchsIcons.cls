VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AllyResearchsIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/research.cpp: class AllyResearchsIcons, not a widget (a member of ResearchObjectButton and
' ResearchOptionButton: their data classes' allyResearchIcons). Its static displayAllyIcon is
' research_ui_c's (the WzFunc FN_AllyResearchsIcons_displayAllyIcon). update takes a module UDT
' array, hence Friend.
Option Explicit

Private Const MAX_ALLY_ICONS As Long = 4
Private allyIcons(0 To MAX_ALLY_ICONS - 1) As WIDGET ' std::shared_ptr<W_LABEL> allyIcons[MAX_ALLY_ICONS]

Public Sub initialize(ByVal parent As WIDGET)
    Dim init As W_LABINIT, i As Long
    For i = 0 To MAX_ALLY_ICONS - 1
        W_LABINIT_construct init
        init.width = iV_GetImageWidth(IntImages, IMAGE_ALLY_RESEARCH)
        init.height = iV_GetImageHeight(IntImages, IMAGE_ALLY_RESEARCH)
        init.x = STAT_BUTWIDTH - (init.width + 2) * i - init.width - 2
        init.y = STAT_BUTHEIGHT - init.height - 3 - STAT_PROGBARHEIGHT
        Set init.pDisplay = WzFunc_new(FN_AllyResearchsIcons_displayAllyIcon)
        Set allyIcons(i) = W_LABEL_new_init(init)
        parent.attach allyIcons(i)
    Next
End Sub

Public Sub hide()
    Dim i As Long
    For i = 0 To MAX_ALLY_ICONS - 1
        allyIcons(i).hide
    Next
End Sub

' allyResearches: std::vector<AllyResearch> (its count allyResearches_size)
Friend Sub update(allyResearches() As AllyResearch, ByVal allyResearches_size As Long)
    Dim i As Long
    i = 0
    Do While i < MAX_ALLY_ICONS And i < allyResearches_size
        allyIcons(i).UserData = getPlayerColour(allyResearches(i).player)
        allyIcons(i).setTip getPlayerName(allyResearches(i).player)
        allyIcons(i).show
        i = i + 1
    Loop
End Sub

Public Sub setTop(ByVal y As Long)
    Dim i As Long
    For i = 0 To MAX_ALLY_ICONS - 1
        allyIcons(i).move allyIcons(i).x(), y
    Next
End Sub
