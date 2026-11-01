Attribute VB_Name = "mapdisplay_c"
' src/mapdisplay.cpp, mapdisplay.h: MapDisplay - Renders the world view necessary for the intelligence map
' (Alex McLean, Pumpkin Studios, EIDOS Interactive, 1997): renderResearchToBuffer, the research topic's model
' turning in the intelligence screen's view. A RESEARCH * is the research index (-1: nullptr); a BASE_STATS *
' a stats handle (STAT_RESEARCH + index for a research topic, as intelmap_c's researchAsStats).
Option Explicit

Private Const ROTATE_TIME As Long = 2 * GAME_TICKS_PER_SEC

' renders the Research IMDs into the surface - used by message display in
' Intelligence Map
' void renderResearchToBuffer(RESEARCH *psResearch, UDWORD OriginX, UDWORD OriginY, bool rotate = true)
Public Sub renderResearchToBuffer(ByVal psResearch As Long, ByVal OriginX As Long, ByVal OriginY As Long, Optional ByVal rotate As Boolean = True)
    Dim psResGraphic As Long, IMDType_ As ImdType, scale_ As Long, basePlateSize As Long, pDisplayModel As Long
    Dim compID As Long, Radius As Long, angle As Long, Position As Vector3i, rotationV As Vector3i, t As Double
    scale_ = 0

    If psResearch = -1 Then
        Exit Sub
    End If

    pie_SetGeometricOffset OriginX, OriginY + 10

    ' Rotate round
    ' full rotation once every 2 seconds..
    If rotate Then
        t = U32(realTime)                   ' (realTime % ROTATE_TIME) * 360 / ROTATE_TIME: a UDWORD
        angle = ((t - Int(t / ROTATE_TIME) * ROTATE_TIME) * 360) \ ROTATE_TIME
    Else
        angle = 45
    End If
    Position = Vector3i_(0, 0, BUTTON_DEPTH)
    rotationV = Vector3i_(-30, angle, 0)

    'draw the IMD for the research
    If asResearch(psResearch).psStat <> 0 Then
        'we have a Stat associated with this research topic
        If StatIsStructure(asResearch(psResearch).psStat) Then
            'this defines how the button is drawn
            IMDType_ = IMDTYPE_STRUCTURESTAT
            psResGraphic = asResearch(psResearch).psStat
            'set up the scale
            basePlateSize = getStructureStatSizeMax(castStructureStats(asResearch(psResearch).psStat))
            If basePlateSize = 1 Then
                scale_ = RESEARCH_COMPONENT_SCALE \ 2
                'HACK HACK HACK!
                'if its a 'tall thin (ie tower)' structure stat with something on
                'the top - offset the position to show the object on top
                pDisplayModel = iIMDBaseShape_displayModel(asStructureStats(castStructureStats(asResearch(psResearch).psStat)).pIMD(0))
                If iIMDShape_connectors_size(pDisplayModel) <> 0 And _
                   getStructureStatHeight(castStructureStats(asResearch(psResearch).psStat)) > TOWER_HEIGHT Then
                    Position.y = Position.y - 30
                End If
            ElseIf basePlateSize = 2 Then
                scale_ = RESEARCH_COMPONENT_SCALE \ 4
            Else
                scale_ = RESEARCH_COMPONENT_SCALE \ 5
            End If
        Else
            compID = StatIsComponent(asResearch(psResearch).psStat)
            If compID <> COMP_NUMCOMPONENTS Then
                'this defines how the button is drawn
                IMDType_ = IMDTYPE_COMPONENT
                psResGraphic = asResearch(psResearch).psStat
                ' FIXME: Another kludge to deal with the superTransport to make it "fit" the display.
                If asResearch(psResearch).id = "R-SuperTransport" Then
                    scale_ = RESEARCH_COMPONENT_SCALE \ 3
                ElseIf asResearch(psResearch).id = "R-Cyborg-Transport" Then
                    scale_ = RESEARCH_COMPONENT_SCALE \ 2
                Else
                    scale_ = RESEARCH_COMPONENT_SCALE
                End If
            Else
                ASSERT_FAILURE "false", "src/mapdisplay.cpp:117", "renderResearchToBuffer", "intDisplayMessageButton: invalid stat"
                IMDType_ = IMDTYPE_RESEARCH
                psResGraphic = STAT_RESEARCH + psResearch   ' (BASE_STATS *)psResearch
            End If
        End If
    Else
        'no Stat for this research topic so use the research topic to define what is drawn
        psResGraphic = STAT_RESEARCH + psResearch       ' (BASE_STATS *)psResearch
        IMDType_ = IMDTYPE_RESEARCH
    End If

    'scale the research according to size of IMD
    If IMDType_ = IMDTYPE_RESEARCH Then
        Radius = getResearchRadius(psResGraphic)
        If Radius <= 100 Then
            scale_ = RESEARCH_COMPONENT_SCALE \ 2
        ElseIf Radius <= 128 Then
            scale_ = RESEARCH_COMPONENT_SCALE \ 3
        ElseIf Radius <= 256 Then
            scale_ = RESEARCH_COMPONENT_SCALE \ 4
        Else
            scale_ = RESEARCH_COMPONENT_SCALE \ 5
        End If
    End If

    ' display the IMDs
    If IMDType_ = IMDTYPE_COMPONENT Then
        displayComponentButton psResGraphic, rotationV, Position, scale_
    ElseIf IMDType_ = IMDTYPE_RESEARCH Then
        displayResearchButton psResGraphic, rotationV, Position, scale_
    ElseIf IMDType_ = IMDTYPE_STRUCTURESTAT Then
        displayStructureStatButton castStructureStats(psResGraphic), rotationV, Position, scale_
    Else
        ASSERT_FAILURE "false", "src/mapdisplay.cpp:167", "renderResearchToBuffer", "Unknown PIEType"
    End If
End Sub
