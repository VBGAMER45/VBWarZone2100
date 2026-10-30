Attribute VB_Name = "intdisplay_c"
' src/intdisplay.cpp, intdisplay.h: the interface's display and callback functions: the bar
' graph formats (formatTime, formatPower), the power bar (class PowerBar : W_BARGRAPH), the
' buttons that render a model (class IntFancyButton : W_CLICKFORM and its IntObjectButton,
' IntStatusButton, IntStatsButton, IntTransportButton), the forms IntFormTransparent and
' IntFormAnimated, the widgets' display functions (intDisplayImage ... intDisplayEditBox,
' intDisplayProximityBlips), the queries on droids, structures and stats the panels use, the
' widgets' audio callback, the radar's proximity blips and the quantity sliders' callback.
' Each widget class's fields are the class of the same name (IntFancyButton's anonymous
' struct model: its model_ fields). An ImdObject's const void *ptr is a handle: a BASE_OBJECT *
' (IMDTYPE_DROID, IMDTYPE_STRUCTURE), a BASE_STATS * (the stats types: a template's is its
' BASE_STATS * handle, STAT_TEMPLATE plus its pool handle) or an iIMDShape * (IMDTYPE_FEATURE),
' 0 = nullptr. A FACTORY * and a RESEARCH_FACILITY * are their structure's handle; the
' templates _structureIsManufacturingPending and structureIsOnHoldPending have a version per
' functionality. PowerBar::display's statics and WidgetAudioCallback's are module variables;
' drawRadarBlips's static image tables are radarBlipImage.
' (VB) widget class IntFormTransparent : W_FORM
' (VB) widget class IntFormAnimated : W_FORM
' (VB) widget class PowerBar : W_BARGRAPH
' (VB) widget class IntFancyButton : W_CLICKFORM
' (VB) widget class IntObjectButton : IntFancyButton
' (VB) widget class IntStatusButton : IntObjectButton
' (VB) widget class IntStatsButton : IntFancyButton
' (VB) widget class IntTransportButton : IntFancyButton
Option Explicit

' Power levels are divided by this for power bar display. The extra factor has
' been included so that the levels appear the same for the power bar as for the
' power values in the buttons
Public Const POWERBAR_SCALE As Long = 5 * WBAR_SCALE \ STAT_PROGBARWIDTH

Public Const BUTTONOBJ_ROTSPEED As Long = 90    ' Speed to rotate objects rendered in
' buttons ( degrees per second )

Public Enum ImdType
    IMDTYPE_NONE
    IMDTYPE_DROID
    IMDTYPE_DROIDTEMPLATE
    IMDTYPE_COMPONENT
    IMDTYPE_STRUCTURE
    IMDTYPE_RESEARCH
    IMDTYPE_STRUCTURESTAT
    IMDTYPE_FEATURE
End Enum

' struct ImdObject (ImdObject() : ptr(nullptr), type(IMDTYPE_NONE) is a blank UDT)
Public Type ImdObject
    ptr As Long                         ' const void *: a handle (see above), 0 = nullptr
    typ As ImdType
End Type

'the two types of button used in the object display (bottom bar)
Public Enum IntFancyButton_ButtonType
    TOPBUTTON = 0
    BTMBUTTON = 1
End Enum

Private Const FORM_OPEN_ANIM_DURATION As Long = GAME_TICKS_PER_SEC \ 6 ' Time duration for form open/close anims.

'the loop default value
Private Const DEFAULT_LOOP As Long = 1

Private FormOpenAudioID As Long         ' ID of sfx to play when form opens.
Private FormCloseAudioID As Long        ' ID of sfx to play when form closes.
Private FormOpenCount As Long           ' Count used to ensure only one sfx played when two forms opening.
Private FormCloseCount As Long          ' Count used to ensure only one sfx played when two forms closeing.

Private Const DEFAULT_BUTTON_ROTATION As Long = 45

Private ManuPower As Long               ' UDWORD = 0: Power required to manufacture the current item.

Private displayPower As Double          ' PowerBar::display's static double displayPower
Private lastRealTime As Long            ' and static unsigned lastRealTime

Private Const WIDGETBEEPGAP As Long = 200   ' 200 milliseconds between each beep please
Private LastTimeAudio As Long           ' WidgetAudioCallback's static SDWORD LastTimeAudio

' the display functions and callbacks as WzFuncs (fork G1's ids)
Public Const FN_intDisplayImage As Long = 7100           ' (widget, xOffset, yOffset)
Public Const FN_intDisplayMissionClock As Long = 7101    ' (widget, xOffset, yOffset)
Public Const FN_intDisplayImageHilight As Long = 7102    ' (widget, xOffset, yOffset)
Public Const FN_intDisplayButtonHilight As Long = 7103   ' (widget, xOffset, yOffset)
Public Const FN_intDisplayButtonFlash As Long = 7104     ' (widget, xOffset, yOffset)
Public Const FN_intDisplayEditBox As Long = 7105         ' (widget, xOffset, yOffset)
Public Const FN_WidgetAudioCallback As Long = 7106       ' WIDGET_AUDIOCALLBACK (audioID)
Public Const FN_intDisplayProximityBlips As Long = 7107  ' (widget, xOffset, yOffset)

' MARK: - ImdObject's named constructors

Public Function ImdObject_Droid(ByVal p As Long) As ImdObject
    ImdObject_Droid.ptr = p
    ImdObject_Droid.typ = IMDTYPE_DROID
End Function

Public Function ImdObject_DroidTemplate(ByVal p As Long) As ImdObject
    ImdObject_DroidTemplate.ptr = p
    ImdObject_DroidTemplate.typ = IMDTYPE_DROIDTEMPLATE
End Function

Public Function ImdObject_Component(ByVal p As Long) As ImdObject
    ImdObject_Component.ptr = p
    ImdObject_Component.typ = IMDTYPE_COMPONENT
End Function

Public Function ImdObject_Structure(ByVal p As Long) As ImdObject
    ImdObject_Structure.ptr = p
    ImdObject_Structure.typ = IMDTYPE_STRUCTURE
End Function

Public Function ImdObject_Research(ByVal p As Long) As ImdObject
    ImdObject_Research.ptr = p
    ImdObject_Research.typ = IMDTYPE_RESEARCH
End Function

Public Function ImdObject_StructureStat(ByVal p As Long) As ImdObject
    ImdObject_StructureStat.ptr = p
    ImdObject_StructureStat.typ = IMDTYPE_STRUCTURESTAT
End Function

' p: a FEATURE_STATS's BASE_STATS * handle
Public Function ImdObject_Feature(ByVal p As Long) As ImdObject
    Dim fStat As Long
    fStat = p And &HFFFF&               ' FEATURE_STATS *fStat = (FEATURE_STATS *)p
    If asFeatureStats(fStat).psImd <> 0 Then ImdObject_Feature.ptr = iIMDBaseShape_displayModel(asFeatureStats(fStat).psImd)
    ImdObject_Feature.typ = IMDTYPE_FEATURE
End Function

Public Function ImdObject_empty(o As ImdObject) As Boolean
    ImdObject_empty = o.ptr = 0
End Function

' (VB) ivisdef.h's static inline safeGetDisplayModelFromBase
Private Function safeGetDisplayModelFromBase(ByVal pBaseIMD As Long) As Long
    If pBaseIMD <> 0 Then safeGetDisplayModelFromBase = iIMDBaseShape_displayModel(pBaseIMD)
End Function

' (VB) a DROID_TEMPLATE * (its pool handle, 0 = nullptr) as a BASE_STATS * handle, as C++ converts it implicitly
Private Function templateStatsBase(ByVal psTemplate As Long) As Long
    If psTemplate <> 0 Then templateStatsBase = STAT_TEMPLATE + psTemplate
End Function

' Set audio IDs for form opening/closing anims.
' Use -1 to dissable audio.
'
Public Sub SetFormAudioIDs(ByVal OpenID As Long, ByVal CloseID As Long)
    FormOpenAudioID = OpenID
    FormCloseAudioID = CloseID
    FormOpenCount = 0
    FormCloseCount = 0
End Sub

' barGraph: a W_BARGRAPH
Private Sub setBarGraphValue(ByVal barGraph As WIDGET, ByVal colour As Long, ByVal value As Long, ByVal range As Long)
    If Not (range <> 0) Then ASSERT_FAILURE "range != 0", "src/intdisplay.cpp:110", "setBarGraphValue", "range is 0": Exit Sub
    W_BARGRAPH_(barGraph).majorCol = colour
    W_BARGRAPH_(barGraph).majorSize = ((clip_(value, 0, range) * WBAR_SCALE) \ range) And &HFFFF&    ' PERNUM(WBAR_SCALE, clip(value, 0, range), range)
    barGraph.show
End Sub

Private Sub formatEmpty(ByVal barGraph As WIDGET)
    W_BARGRAPH_(barGraph).text = ""
    setBarGraphValue barGraph, psPalette(WZCOL_BLACK), 0, 1
End Sub

Private Sub formatTimeText(ByVal barGraph As WIDGET, ByVal time_ As Long)
    Dim timeText As String
    timeText = va("%d:%02d", time_ \ 60, time_ Mod 60)
    W_BARGRAPH_(barGraph).text = timeText
    W_BARGRAPH_(barGraph).textCol = psPalette(WZCOL_CONSTRUCTION_BARTEXT)
End Sub

Public Sub formatTime(ByVal barGraph As WIDGET, ByVal buildPointsDone As Long, ByVal buildPointsTotal As Long, ByVal buildRate As Long, ByVal toolTip As String)
    Dim timeToBuild As Long
    barGraph.setTip toolTip

    If buildRate <> 0 Then
        timeToBuild = (buildPointsTotal - buildPointsDone) \ buildRate

        formatTimeText barGraph, timeToBuild
    Else
        W_BARGRAPH_(barGraph).text = ""
    End If

    setBarGraphValue barGraph, psPalette(WZCOL_YELLOW), buildPointsDone, buildPointsTotal
End Sub

Private Sub formatPowerText(ByVal barGraph As WIDGET, ByVal neededPower As Long)
    Dim powerText As String
    powerText = va("%d", neededPower)
    W_BARGRAPH_(barGraph).text = powerText
    W_BARGRAPH_(barGraph).textCol = psPalette(WZCOL_POWERQUEUE_BARTEXT)
End Sub

Public Sub formatPower(ByVal barGraph As WIDGET, ByVal neededPower As Long, ByVal powerToBuild As Long)
    If Not (Not barGraph Is Nothing) Then ASSERT_FAILURE "(barGraph) != nullptr", "src/intdisplay.cpp:158", "formatPower", "%s should not be nullptr", "barGraph": Exit Sub
    If neededPower = -1 Or powerToBuild = 0 Then
        formatEmpty barGraph
        Exit Sub
    End If

    barGraph.setTip "Waiting for Power"
    formatPowerText barGraph, neededPower
    setBarGraphValue barGraph, psPalette(WZCOL_GREEN), powerToBuild - neededPower, powerToBuild
End Sub

' MARK: - PowerBar

' PowerBar(W_BARINIT* init): W_BARGRAPH(init) {}
Public Sub PowerBar_construct(ByVal self As WIDGET, init As W_BARINIT)
    W_BARGRAPH_construct_init self, init
    self.kind = K_PowerBar
End Sub

Public Function PowerBar_new(init As W_BARINIT) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_PowerBar)
    PowerBar_construct widget, init
    Set PowerBar_new = widget
End Function

Public Function PowerBar_getTip(ByVal self As WIDGET) As String
    Dim income As String
    income = getApproxPowerGeneratedPerSecForDisplay(selectedPlayer)
    PowerBar_getTip = va("%s\n%s: %s", "Power", "Power Per Second", income)
End Function

' !!!!!!!!!!!!!!!!!!!!!!ONLY WORKS ON A SIDEWAYS POWERBAR!!!!!!!!!!!!!!!!!
Public Sub PowerBar_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim Avail As Long, ManPow As Long, realPower As Long
    Dim Empty_ As Long
    Dim BarWidth As Long, textWidth As Long
    Dim iX As Long, iY As Long
    Dim desiredPower As Double, elapsed As Double, d As PowerBar
    Dim x0 As Long, y0 As Long, top_x0 As Long, top_y0 As Long, imageDrawBatch As BatchedImageDrawRequests
    Dim unusedDerricks As Long, showNeedMessage As Boolean, unusedText As String
    Dim needTextWidth As Long, textX As Long, colour As Long
    Set d = PowerBar_(self)
    textWidth = 0

    desiredPower = getPowerMinusQueued(selectedPlayer)
    elapsed = CDbl(realTime) - CDbl(lastRealTime)       ' (unsigned)
    If elapsed < 0 Then elapsed = elapsed + 4294967296#
    displayPower = desiredPower + (displayPower - desiredPower) * Exp(elapsed / -80#) ' If realTime < lastRealTime, then exp() returns 0 due to unsigned overflow.
    lastRealTime = realTime

    ManPow = UDIV_(ManuPower, POWERBAR_SCALE)
    Avail = Fix((displayPower + 0.00000001) / POWERBAR_SCALE)
    realPower = Fix((displayPower + 0.00000001) - U32(ManuPower))
    ManuPower = 0

    BarWidth = self.width()

    d.wzText_.setText CStr(realPower), font_regular

    textWidth = d.wzText_.width()
    BarWidth = BarWidth - textWidth

    If ManPow > Avail Then
        Empty_ = BarWidth - ManPow
    Else
        Empty_ = BarWidth - Avail
    End If

    If Avail > BarWidth Then
        ManPow = (ManPow * BarWidth) \ Avail ' PERNUM(BarWidth, ManPow, Avail)
        Avail = BarWidth
        Empty_ = 0
    End If

    If ManPow > BarWidth Then
        ManPow = BarWidth
        Avail = 0
        Empty_ = 0
    End If

    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    pie_SetFogStatus False

    top_x0 = x0
    top_y0 = y0

    iX = x0 + 3
    iY = y0 + 10

    x0 = x0 + iV_GetImageWidth(IntImages, IMAGE_PBAR_TOP)

    Set imageDrawBatch = BatchedImageDrawRequests_new(True) ' defer drawing

    'fill in the empty section behind text
    If textWidth > 0 Then
        iV_DrawImageRepeatX IntImages, IMAGE_PBAR_EMPTY, x0 - 1, y0, textWidth + 1, 0, True, imageDrawBatch ' Overdraw by 1 to reduce seam with left-beginning-piece when scaling
        x0 = x0 + textWidth
    End If

    'draw the left-most beginning tip
    'to reduce a visible seam when scaling, this must come *after* the empty / text section above
    iV_DrawImage IntImages, IMAGE_PBAR_TOP, top_x0, top_y0, 0, imageDrawBatch

    'draw required section
    If ManPow > Avail Then
        'draw the required in red
        iV_DrawImageRepeatX IntImages, IMAGE_PBAR_USED, x0, y0, ManPow, 0, True, imageDrawBatch
    Else
        iV_DrawImageRepeatX IntImages, IMAGE_PBAR_REQUIRED, x0, y0, ManPow, 0, True, imageDrawBatch
    End If
    x0 = x0 + ManPow

    'draw the available section if any!
    If Avail - ManPow > 0 Then
        iV_DrawImageRepeatX IntImages, IMAGE_PBAR_AVAIL, x0, y0, Avail - ManPow, 0, True, imageDrawBatch
        x0 = x0 + Avail - ManPow
    End If

    'fill in the rest with empty section
    If Empty_ > 0 Then
        iV_DrawImageRepeatX IntImages, IMAGE_PBAR_EMPTY, x0, y0, Empty_ + 1, 0, True, imageDrawBatch ' Overdraw by 1 to reduce seam with right-end-piece when scaling
        x0 = x0 + Empty_
    End If

    iV_DrawImage IntImages, IMAGE_PBAR_BOTTOM, x0, y0, 0, imageDrawBatch

    imageDrawBatch.draw True

    unusedDerricks = countPlayerUnusedDerricks()

    showNeedMessage = True
    If unusedDerricks > 0 Then
        unusedText = va("%u derrick(s) inactive", unusedDerricks)
        d.wzNeedText.setText unusedText, font_small
    ElseIf Avail < 0 Then
        d.wzNeedText.setText "Need more resources!", font_small
    Else
        showNeedMessage = False
    End If

    If showNeedMessage Then
        needTextWidth = d.wzNeedText.width()
        textX = iX + (self.width() - needTextWidth) \ 2
        pie_UniTransBoxFill textX - 3, y0 + 1, textX + needTextWidth + 3, y0 + self.height() - 1, psPalette(WZCOL_TRANSPARENT_BOX)
        If ((realTime \ 1250) Mod 5) <> 0 Then colour = psPalette(WZCOL_WHITE) Else colour = psPalette(WZCOL_RED)
        d.wzNeedText.render textX, iY - 1, colour
    End If

    ' draw text value
    If Avail < 0 Then colour = psPalette(WZCOL_RED) Else colour = psPalette(WZCOL_TEXT_BRIGHT)
    d.wzText_.render iX, iY, colour
End Sub

' (VB) an unsigned division of a UDWORD by a positive int
Private Function UDIV_(ByVal a As Long, ByVal b As Long) As Long
    UDIV_ = Fix(U32(a) / b)
End Function

' MARK: - IntFancyButton

Public Sub IntFancyButton_construct(ByVal self As WIDGET)
    Dim d As IntFancyButton, v As Vector3i
    W_CLICKFORM_construct self
    self.kind = K_IntFancyButton
    Set d = IntFancyButton_(self)
    d.buttonType = TOPBUTTON
    d.buttonBackgroundEmpty = False
    v.x = 0
    v.y = 0
    v.z = BUTTON_DEPTH
    d.model_position = v
    v.x = -30
    v.y = DEFAULT_BUTTON_ROTATION
    v.z = 0
    d.model_rotation = v
    d.model_rate = 0
End Sub

Public Function IntFancyButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_IntFancyButton)
    IntFancyButton_construct widget
    Set IntFancyButton_new = widget
End Function

Public Sub IntFancyButton_initDisplay(ByVal self As WIDGET)
    Dim d As IntFancyButton, rot As Vector3i, delta As Long
    Set d = IntFancyButton_(self)
    rot = d.model_rotation
    If W_CLICKFORM_isHighlighted(self) Then
        If W_CLICKFORM_isHighlighted(self) Then delta = 2 * BUTTONOBJ_ROTSPEED Else delta = -4 * BUTTONOBJ_ROTSPEED
        d.model_rate = d.model_rate + realTimeAdjustedAverage(delta)
        d.model_rate = clip_(d.model_rate, 0, BUTTONOBJ_ROTSPEED)
        rot.y = rot.y + realTimeAdjustedAverage(d.model_rate)
        rot.y = rot.y Mod 360
    ElseIf rot.y <> DEFAULT_BUTTON_ROTATION Then
        ' return to default rotation using the nearest direction
        If rot.y > (180 + DEFAULT_BUTTON_ROTATION) Then
            d.model_rate = d.model_rate + realTimeAdjustedAverage(5 * BUTTONOBJ_ROTSPEED)
            d.model_rate = clip_(d.model_rate, 0, 5 * BUTTONOBJ_ROTSPEED)
            rot.y = rot.y + realTimeAdjustedAverage(d.model_rate)
            If rot.y > 360 Then
                rot.y = rot.y Mod 360
                rot.y = MIN_(rot.y, DEFAULT_BUTTON_ROTATION)
            End If
        ElseIf rot.y < DEFAULT_BUTTON_ROTATION Then
            d.model_rate = d.model_rate + realTimeAdjustedAverage(5 * BUTTONOBJ_ROTSPEED)
            d.model_rate = clip_(d.model_rate, 0, 5 * BUTTONOBJ_ROTSPEED)
            rot.y = rot.y + realTimeAdjustedAverage(d.model_rate)
            rot.y = MIN_(rot.y, DEFAULT_BUTTON_ROTATION)
        Else
            d.model_rate = d.model_rate - realTimeAdjustedAverage(7 * BUTTONOBJ_ROTSPEED)
            d.model_rate = clip_(d.model_rate, -5 * BUTTONOBJ_ROTSPEED, BUTTONOBJ_ROTSPEED)
            rot.y = rot.y + realTimeAdjustedAverage(d.model_rate)
            rot.y = MAX_(rot.y, DEFAULT_BUTTON_ROTATION)
        End If
    End If
    d.model_rotation = rot
End Sub

Public Sub IntFancyButton_displayIfHighlight(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    If W_CLICKFORM_isHighlighted(self) Then
        If IntFancyButton_(self).buttonType = TOPBUTTON Then
            iV_DrawImage IntImages, IMAGE_BUT_HILITE, xOffset + self.x(), yOffset + self.y()
        Else
            iV_DrawImage IntImages, IMAGE_BUTB_HILITE, xOffset + self.x(), yOffset + self.y()
        End If
    End If
End Sub

' MARK: - IntStatusButton

Public Sub IntStatusButton_construct(ByVal self As WIDGET)
    IntObjectButton_construct self
    self.kind = K_IntStatusButton
    IntStatusButton_(self).theStats = 0
    IntFancyButton_(self).buttonType = TOPBUTTON
End Sub

Public Function IntStatusButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_IntStatusButton)
    IntStatusButton_construct widget
    Set IntStatusButton_new = widget
End Function

' object: a BASE_OBJECT *
Public Sub IntStatusButton_setObject(ByVal self As WIDGET, ByVal object_ As Long)
    IntObjectButton_(self).psObj = object_
    IntStatusButton_(self).theStats = 0
End Sub

' object: a BASE_OBJECT *, stats: a BASE_STATS *
Public Sub IntStatusButton_setObjectAndStats(ByVal self As WIDGET, ByVal object_ As Long, ByVal stats As Long)
    IntObjectButton_(self).psObj = object_
    IntStatusButton_(self).theStats = stats
End Sub

' Widget callback to display a rendered status button, ie the progress of a manufacturing or  building task.
Public Sub IntStatusButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim Structure_ As Long, Droid_ As Long, Stats As Long, psResGraphic As Long
    Dim compID As Long, bOnHold As Boolean, object_ As ImdObject, image As AtlasImage
    Dim shape As Long, psObj As Long

    bOnHold = False

    IntFancyButton_initDisplay self

    psObj = IntObjectButton_(self).psObj
    If psObj <> 0 Then
        If isDead(psObj) Then
            ' this may catch this horrible crash bug we've been having,
            ' who knows?.... Shipping tomorrow, la de da :-)
            IntObjectButton_(self).psObj = 0
            psObj = 0
            intRefreshScreen
        End If
    End If

    If psObj <> 0 Then
        Select Case SIMPLE_OBJECT_type(psObj)
        Case OBJ_DROID                  ' If it's a droid...
            Droid_ = psObj

            If DroidIsBuilding(Droid_) Then
                Structure_ = DroidGetBuildStructure(Droid_)
                If Structure_ <> 0 Then
                    object_ = ImdObject_Structure(Structure_)
                End If
            ElseIf DroidGoingToBuild(Droid_) Then
                Stats = DroidGetBuildStats(Droid_)
                If Not (Stats <> 0) Then ASSERT_FAILURE "Stats != nullptr", "src/intdisplay.cpp:418", "display", "NULL Stats pointer."
                object_ = ImdObject_StructureStat(Stats)
            ElseIf orderState(Droid_, DORDER_DEMOLISH) Then
                Stats = structureStatsBase(structGetDemolishStat())
                If Not (Stats <> 0) Then ASSERT_FAILURE "Stats != nullptr", "src/intdisplay.cpp:424", "display", "NULL Stats pointer."
                object_ = ImdObject_StructureStat(Stats)
            ElseIf aDroid(Droid_).droidType = DROID_COMMAND Then
                Structure_ = droidGetCommandFactory(Droid_)
                If Structure_ <> 0 Then
                    object_ = ImdObject_Structure(Structure_)
                End If
            End If

        Case OBJ_STRUCTURE              ' If it's a structure...
            Structure_ = psObj
            Select Case asStructureStats(aStruct(Structure_).pStructureType).typ
            Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
                If StructureIsManufacturingPending(Structure_) Then
                    object_ = ImdObject_DroidTemplate(templateStatsBase(FactoryGetTemplate(StructureGetFactory(Structure_))))
                    bOnHold = StructureIsOnHoldPending(Structure_)
                End If

            Case REF_RESEARCH
                If structureIsResearchingPending(Structure_) Then
                    Stats = IntStatusButton_(self).theStats
                    If Stats = 0 Then
                        GoTo researchEnd
                    End If
                    bOnHold = StructureIsOnHoldPending(Structure_)
                    StatGetResearchImage Stats, image, shape, psResGraphic, False
                    If psResGraphic <> 0 Then
                        ' we have a Stat associated with this research topic
                        If StatIsStructure(psResGraphic) Then
                            ' overwrite the Object pointer
                            object_ = ImdObject_StructureStat(psResGraphic)
                        Else
                            compID = StatIsComponent(psResGraphic)
                            If compID <> COMP_NUMCOMPONENTS Then
                                ' overwrite the Object pointer
                                object_ = ImdObject_Component(psResGraphic)
                            Else
                                ASSERT_FAILURE "false", "src/intdisplay.cpp:481", "display", "Invalid Stat for research button"
                                object_ = ImdObject_Research(0)
                            End If
                        End If
                    Else
                        ' no Stat for this research topic so just use the graphic provided
                        ' if Object != NULL the there must be a IMD so set the object to
                        ' equal the Research stat
                        If shape <> 0 Then
                            object_ = ImdObject_Research(Stats)
                        End If
                    End If
                End If
researchEnd:
            Case Else
            End Select

        Case Else
            ASSERT_FAILURE "false", "src/intdisplay.cpp:504", "display", "Invalid structure type"
        End Select
    End If

    ' Render the object into the button.
    IntFancyButton_displayIMD self, image, object_, xOffset, yOffset

    'need to flash the button if a factory is on hold production
    If bOnHold Then
        If ((realTime \ 250) Mod 2) = 0 Then
            iV_DrawImage IntImages, IMAGE_BUT0_DOWN, xOffset + self.x(), yOffset + self.y()
        Else
            iV_DrawImage IntImages, IMAGE_BUT_HILITE, xOffset + self.x(), yOffset + self.y()
        End If
    Else
        IntFancyButton_displayIfHighlight self, xOffset, yOffset
    End If
End Sub

' MARK: - IntObjectButton

Public Sub IntObjectButton_construct(ByVal self As WIDGET)
    IntFancyButton_construct self
    self.kind = K_IntObjectButton
    IntObjectButton_(self).psObj = 0
    IntFancyButton_(self).buttonType = BTMBUTTON
End Sub

Public Function IntObjectButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_IntObjectButton)
    IntObjectButton_construct widget
    Set IntObjectButton_new = widget
End Function

' object: a BASE_OBJECT *
Public Sub IntObjectButton_setObject(ByVal self As WIDGET, ByVal object_ As Long)
    IntObjectButton_(self).psObj = object_
End Sub

Public Function IntObjectButton_clearData(ByVal self As WIDGET) As Boolean
    Dim ret As Boolean
    ret = IntObjectButton_(self).psObj <> 0
    IntObjectButton_(self).psObj = 0
    IntObjectButton_clearData = ret
End Function

' Widget callback to display a rendered object button.
Public Sub IntObjectButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim object_ As ImdObject, psObj As Long, noImage As AtlasImage
    IntFancyButton_initDisplay self

    psObj = IntObjectButton_(self).psObj
    If psObj <> 0 Then
        If isDead(psObj) Then
            ' this may catch this horrible crash bug we've been having,
            ' who knows?.... Shipping tomorrow, la de da :-)
            IntObjectButton_(self).psObj = 0
            psObj = 0
            intRefreshScreen
        End If
    End If

    If psObj <> 0 Then
        Select Case SIMPLE_OBJECT_type(psObj)
        Case OBJ_DROID                  ' If it's a droid...
            object_ = ImdObject_Droid(psObj)
        Case OBJ_STRUCTURE              ' If it's a structure...
            object_ = ImdObject_Structure(psObj)
        Case Else
            ASSERT_FAILURE "false", "src/intdisplay.cpp:555", "display", "Invalid structure type"
        End Select
    End If

    IntFancyButton_displayIMD self, noImage, object_, xOffset, yOffset
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

' MARK: - IntStatsButton

Public Sub IntStatsButton_construct(ByVal self As WIDGET)
    IntFancyButton_construct self
    self.kind = K_IntStatsButton
    IntStatsButton_(self).Stat = 0
End Sub

Public Function IntStatsButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_IntStatsButton)
    IntStatsButton_construct widget
    Set IntStatsButton_new = widget
End Function

' stats: a BASE_STATS *
Public Sub IntStatsButton_setStats(ByVal self As WIDGET, ByVal stats As Long)
    IntStatsButton_(self).Stat = stats
End Sub

Public Sub IntStatsButton_setStatsAndTip(ByVal self As WIDGET, ByVal stats As Long)
    IntStatsButton_setStats self, stats
    self.setTip BASE_STATS_name(stats)  ' getLocalizedStatsName(stats)
End Sub

' Widget callback to display a rendered stats button, ie the job selection window buttons.
'
Public Sub IntStatsButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim psResGraphic As Long, compID As Long, object_ As ImdObject, image As AtlasImage, Stat As Long, shape As Long

    IntFancyButton_initDisplay self

    Stat = IntStatsButton_(self).Stat
    If Stat <> 0 Then
        If StatIsStructure(Stat) Then
            object_ = ImdObject_StructureStat(Stat)
        ElseIf StatIsTemplate(Stat) Then
            object_ = ImdObject_DroidTemplate(Stat)
        ElseIf StatIsFeature(Stat) Then
            object_ = ImdObject_Feature(Stat)
        Else
            compID = StatIsComponent(Stat) ' This fails for viper body.
            If compID <> COMP_NUMCOMPONENTS Then
                object_ = ImdObject_Component(Stat)
            ElseIf StatIsResearch(Stat) Then
                shape = 0
                StatGetResearchImage Stat, image, shape, psResGraphic, True
                If psResGraphic <> 0 Then
                    'we have a Stat associated with this research topic
                    If StatIsStructure(psResGraphic) Then
                        object_ = ImdObject_StructureStat(psResGraphic)
                    Else
                        compID = StatIsComponent(psResGraphic)
                        If compID <> COMP_NUMCOMPONENTS Then
                            'overwrite the Object pointer
                            object_ = ImdObject_Component(psResGraphic)
                        Else
                            ASSERT_FAILURE "false", "src/intdisplay.cpp:622", "display", "Invalid Stat for research button"
                            object_ = ImdObject_Research(0)
                        End If
                    End If
                Else
                    'no Stat for this research topic so just use the graphic provided
                    'if Object != NULL the there must be a IMD so set the object to
                    'equal the Research stat
                    If shape <> 0 Then
                        object_ = ImdObject_Research(Stat)
                    End If
                End If
            End If
        End If
    Else
        'BLANK button for now - AB 9/1/98
        object_ = ImdObject_Component(0)
    End If

    IntFancyButton_displayIMD self, image, object_, xOffset, yOffset
    IntFancyButton_displayIfHighlight self, xOffset, yOffset
End Sub

' MARK: - IntFormTransparent: Form which only acts as a glass container.

Public Sub IntFormTransparent_construct(ByVal self As WIDGET)
    W_FORM_construct self
    self.kind = K_IntFormTransparent
End Sub

Public Function IntFormTransparent_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_IntFormTransparent)
    IntFormTransparent_construct widget
    Set IntFormTransparent_new = widget
End Function

Public Sub IntFormTransparent_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
End Sub

' MARK: - IntFormAnimated: Form which animates opening/closing.

Public Sub IntFormAnimated_construct(ByVal self As WIDGET, Optional ByVal openAnimate As Boolean = True)
    W_FORM_construct self
    self.kind = K_IntFormAnimated
    With IntFormAnimated_(self)
        .startTime = 0
        If openAnimate Then .currentAction = 0 Else .currentAction = 2
    End With
    W_FORM_(self).disableChildren = openAnimate
End Sub

Public Function IntFormAnimated_new(Optional ByVal openAnimate As Boolean = True) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_IntFormAnimated)
    IntFormAnimated_construct widget, openAnimate
    Set IntFormAnimated_new = widget
End Function

' Animates the form closing, and deletes itself when done.
' _onCloseAnimFinished: W_ANIMATED_ON_CLOSE_FUNC (void (IntFormAnimated&)), Nothing = nullptr
Public Sub IntFormAnimated_closeAnimateDelete(ByVal self As WIDGET, Optional ByVal onCloseAnimFinished_ As WzFunc = Nothing)
    IntFormAnimated_(self).currentAction = 3
    W_FORM_(self).disableChildren = True
    Set IntFormAnimated_(self).onCloseAnimFinished = onCloseAnimFinished_
End Sub

Public Function IntFormAnimated_isClosing(ByVal self As WIDGET) As Boolean
    IntFormAnimated_isClosing = IntFormAnimated_(self).currentAction >= 3
End Function

Public Sub IntFormAnimated_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As IntFormAnimated, aOpen As WzRect, aClosed As WzRect, aBegin As WzRect, aEnd As WzRect, aCur As WzRect
    Dim den As Long, num As Long, elapsed As Long
    Set d = IntFormAnimated_(self)
    aOpen.x = xOffset + self.x(): aOpen.y = yOffset + self.y(): aOpen.w = self.width(): aOpen.h = self.height()
    aClosed.x = aOpen.x + aOpen.w \ 4: aClosed.y = aOpen.y + aOpen.h \ 2 - 4: aClosed.w = aOpen.w \ 2: aClosed.h = 8
    Select Case d.currentAction
    Case 1: FormOpenCount = 0
    Case 4: FormCloseCount = 0
    End Select
    Select Case d.currentAction
    Case 0, 1
        If d.currentAction = 0 Then ' Start opening.
            If FormOpenAudioID >= 0 And FormOpenCount = 0 Then
                audio_PlayTrack FormOpenAudioID
                FormOpenCount = FormOpenCount + 1
            End If
            d.startTime = realTime
            d.currentAction = d.currentAction + 1
            ' fallthrough
        End If
        ' Continue opening.
        aBegin = aClosed
        aEnd = aOpen
    Case 2  ' Open.
        aBegin = aOpen
        aEnd = aOpen
        d.startTime = realTime
    Case 3, 4
        If d.currentAction = 3 Then ' Start closing.
            If FormCloseAudioID >= 0 And FormCloseCount = 0 Then
                audio_PlayTrack FormCloseAudioID
                FormCloseCount = FormCloseCount + 1
            End If
            d.startTime = realTime
            d.currentAction = d.currentAction + 1
            ' fallthrough
        End If
        ' Continue closing.
        aBegin = aOpen
        aEnd = aClosed
    End Select
    den = FORM_OPEN_ANIM_DURATION
    elapsed = realTime - d.startTime    ' std::min<unsigned>(realTime - startTime, den)
    If elapsed < 0 Or elapsed > den Then num = den Else num = elapsed
    If num = den Then
        d.currentAction = d.currentAction + 1
        Select Case d.currentAction
        Case 2: W_FORM_(self).disableChildren = False
        Case 5
            self.deleteLater
            If Not d.onCloseAnimFinished Is Nothing Then
                WzFunc_call d.onCloseAnimFinished, self
            End If
        End Select
    End If

    aCur.x = aBegin.x + (aEnd.x - aBegin.x) * num \ den
    aCur.y = aBegin.y + (aEnd.y - aBegin.y) * num \ den
    aCur.w = aBegin.w + (aEnd.w - aBegin.w) * num \ den
    aCur.h = aBegin.h + (aEnd.h - aBegin.h) * num \ den

    RenderWindowFrame FRAME_NORMAL, aCur.x, aCur.y, aCur.w, aCur.h
End Sub

' MARK: - the display functions

' Display an image for a widget.
'
Public Sub intDisplayImage(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    iV_DrawImage IntImages, psWidget.UserData, x, y
End Sub


'draws the mission clock - flashes when below a predefined time
Public Sub intDisplayMissionClock(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, flash As Boolean
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    ' Draw the background image
    iV_DrawImage IntImages, UNPACKDWORD_TRI_B(psWidget.UserData), x, y
    ' Need to flash the timer when < 5 minutes remaining, but > 4 minutes
    flash = UNPACKDWORD_TRI_A(psWidget.UserData) <> 0
    If flash And ((realTime \ 250) Mod 2) = 0 Then
        iV_DrawImage IntImages, UNPACKDWORD_TRI_C(psWidget.UserData), x, y
    End If
End Sub


' Display one of two images depending on if the widget is hilighted by the mouse.
'
Public Sub intDisplayImageHilight(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, ImageID As Long, Hilight As Boolean, flash As Boolean
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()
    Hilight = False

    Select Case psWidget.typ
    Case WIDG_FORM
        Hilight = (psWidget.getState() And WBUT_HIGHLIGHT) <> 0

    Case WIDG_BUTTON
        Hilight = (psWidget.getState() And WBUT_HIGHLIGHT) <> 0     ' buttonIsHilite(psWidget)

    Case WIDG_EDITBOX
        If (W_EDITBOX_(psWidget).state And WEDBS_HILITE) <> 0 Then
            Hilight = True
        End If

    Case WIDG_SLIDER
        If W_SLIDER_isHighlighted(psWidget) Then
            Hilight = True
        End If

    Case Else
        Hilight = False
    End Select

    ImageID = UNPACKDWORD_TRI_C(psWidget.UserData)


    'need to flash the button if Full Transporter
    flash = UNPACKDWORD_TRI_A(psWidget.UserData) <> 0
    If flash And psWidget.id = IDTRANS_LAUNCH Then
        If ((realTime \ 250) Mod 2) = 0 Then
            iV_DrawImage IntImages, UNPACKDWORD_TRI_B(psWidget.UserData), x, y
        Else
            iV_DrawImage IntImages, ImageID, x, y
        End If
    Else
        iV_DrawImage IntImages, ImageID, x, y
        If Hilight Then
            iV_DrawImage IntImages, UNPACKDWORD_TRI_B(psWidget.UserData), x, y
        End If
    End If
End Sub

' Display one of two images depending on whether the widget is highlighted by the mouse.
Public Sub intDisplayButtonHilight(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, ImageID As Long, state As Long, grey As Boolean, down As Boolean, highlight As Boolean
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    state = psWidget.getState()
    grey = (state And WBUT_DISABLE) <> 0
    down = (state And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    highlight = (state And WBUT_HIGHLIGHT) <> 0
    If grey Then
        ImageID = UNPACKDWORD_TRI_A(psWidget.UserData)
        highlight = False
    Else
        ImageID = UNPACKDWORD_TRI_C(psWidget.UserData)
        If down Then ImageID = ImageID + 1
    End If

    iV_DrawImage IntImages, ImageID, x, y
    If highlight Then
        iV_DrawImage IntImages, UNPACKDWORD_TRI_B(psWidget.UserData), x, y
    End If

End Sub

' Flash one of two images, regardless of whether or not it is highlighted
' Commented-out portions are retained because I am planning on making the intensity of the
' flash depend on whether or not the button is highlighted.
Public Sub intDisplayButtonFlash(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x As Long, y As Long, ImageID As Long
    x = xOffset + psWidget.x()
    y = yOffset + psWidget.y()

    If Not (psWidget.typ = WIDG_BUTTON) Then ASSERT_FAILURE "psWidget->type == WIDG_BUTTON", "src/intdisplay.cpp:879", "intDisplayButtonFlash", "Not a button"

    If (realTime \ 250) Mod 2 = 0 Then
        ImageID = UNPACKDWORD_TRI_B(psWidget.UserData)
    Else
        ImageID = UNPACKDWORD_TRI_C(psWidget.UserData)
    End If

    iV_DrawImage IntImages, ImageID, x, y
End Sub

' display highlighted edit box from left, middle and end edit box graphics
Public Sub intDisplayEditBox(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim psEditBox As WIDGET, iImageIDLeft As Long, iImageIDMid As Long, iImageIDRight As Long
    Dim iX As Long, iY As Long, iXRight As Long, iXLeft As Long, iYLeft As Long

    Set psEditBox = psWidget
    iXLeft = xOffset + psWidget.x()
    iYLeft = yOffset + psWidget.y()

    If (W_EDITBOX_(psEditBox).state And WEDBS_HILITE) <> 0 Then
        iImageIDLeft = IMAGE_DES_EDITBOXLEFTH
        iImageIDMid = IMAGE_DES_EDITBOXMIDH
        iImageIDRight = IMAGE_DES_EDITBOXRIGHTH
    Else
        iImageIDLeft = IMAGE_DES_EDITBOXLEFT
        iImageIDMid = IMAGE_DES_EDITBOXMID
        iImageIDRight = IMAGE_DES_EDITBOXRIGHT
    End If

    ' draw left side of bar
    iX = iXLeft
    iY = iYLeft
    iV_DrawImage IntImages, iImageIDLeft, iX, iY

    ' draw middle of bar
    iX = iX + iV_GetImageWidth(IntImages, iImageIDLeft)
    iXRight = xOffset + psWidget.width() - iV_GetImageWidth(IntImages, iImageIDRight)
    If ULT(iX, iXRight) Then
        iV_DrawImageRepeatX IntImages, iImageIDMid, iX, iY, (iXRight - iX) + 3, 0, True
    End If

    ' draw right side of bar
    iV_DrawImage IntImages, iImageIDRight, iXRight, iY
End Sub

' Initialise all the surfaces,graphics etc. used by the interface.
'
Public Function intInitialiseGraphics() As Boolean
    ' Initialise any bitmaps used by the interface.
    If Not imageInitBitmaps() Then
        intInitialiseGraphics = False
        Exit Function
    End If

    intInitialiseGraphics = True
End Function

' Clear a button bitmap. ( copy the button background ).
Public Sub IntFancyButton_displayClear(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim buttonId As Long
    buttonId = 0
    Select Case IntFancyButton_(self).buttonType
    Case TOPBUTTON
        If IntFancyButton_(self).buttonBackgroundEmpty Then
            If W_CLICKFORM_isDown(self) Then buttonId = IMAGE_BUT_EMPTY_DOWN Else buttonId = IMAGE_BUT_EMPTY_UP
        Else
            If W_CLICKFORM_isDown(self) Then buttonId = IMAGE_BUT0_DOWN Else buttonId = IMAGE_BUT0_UP
        End If
    Case BTMBUTTON
        If W_CLICKFORM_isDown(self) Then buttonId = IMAGE_BUTB0_DOWN Else buttonId = IMAGE_BUTB0_UP
    End Select
    iV_DrawImage IntImages, buttonId, xOffset + self.x(), yOffset + self.y()
End Sub

' (VB) pie_SetGeometricOffset for a composite droid: the down image's centre, or the up image's
Private Sub IntFancyButton_setModelOffset(ByVal self As WIDGET, ByVal ButXPos As Long, ByVal ButYPos As Long)
    If W_CLICKFORM_isDown(self) Then
        'the top button is smaller than the bottom button
        If IntFancyButton_(self).buttonType = TOPBUTTON Then
            pie_SetGeometricOffset _
                (ButXPos + iV_GetImageWidth(IntImages, IMAGE_BUT0_DOWN) \ 2) + 2, _
                (ButYPos + iV_GetImageHeight(IntImages, IMAGE_BUT0_DOWN) \ 2) + 2 + 8
        Else
            pie_SetGeometricOffset _
                (ButXPos + iV_GetImageWidth(IntImages, IMAGE_BUTB0_DOWN) \ 2) + 2, _
                (ButYPos + iV_GetImageHeight(IntImages, IMAGE_BUTB0_DOWN) \ 2) + 2 + 12
        End If
    Else
        'the top button is smaller than the bottom button
        If IntFancyButton_(self).buttonType = TOPBUTTON Then
            pie_SetGeometricOffset _
                (ButXPos + iV_GetImageWidth(IntImages, IMAGE_BUT0_UP) \ 2), _
                (ButYPos + iV_GetImageHeight(IntImages, IMAGE_BUT0_UP) \ 2) + 8
        Else
            pie_SetGeometricOffset _
                (ButXPos + iV_GetImageWidth(IntImages, IMAGE_BUT0_UP) \ 2), _
                (ButYPos + iV_GetImageHeight(IntImages, IMAGE_BUTB0_UP) \ 2) + 12
        End If
    End If
End Sub

' (VB) pie_SetGeometricOffset for a single IMD: as IntFancyButton_setModelOffset, but an up bottom
' button's x is IMAGE_BUTB0_UP's (the droids' is IMAGE_BUT0_UP's)
Private Sub IntFancyButton_setSingleModelOffset(ByVal self As WIDGET, ByVal ButXPos As Long, ByVal ButYPos As Long)
    If W_CLICKFORM_isDown(self) Then
        If IntFancyButton_(self).buttonType = TOPBUTTON Then
            pie_SetGeometricOffset _
                (ButXPos + iV_GetImageWidth(IntImages, IMAGE_BUT0_DOWN) \ 2) + 2, _
                (ButYPos + iV_GetImageHeight(IntImages, IMAGE_BUT0_DOWN) \ 2) + 2 + 8
        Else
            pie_SetGeometricOffset _
                (ButXPos + iV_GetImageWidth(IntImages, IMAGE_BUTB0_DOWN) \ 2) + 2, _
                (ButYPos + iV_GetImageHeight(IntImages, IMAGE_BUTB0_DOWN) \ 2) + 2 + 12
        End If
    Else
        If IntFancyButton_(self).buttonType = TOPBUTTON Then
            pie_SetGeometricOffset _
                (ButXPos + iV_GetImageWidth(IntImages, IMAGE_BUT0_UP) \ 2), _
                (ButYPos + iV_GetImageHeight(IntImages, IMAGE_BUT0_UP) \ 2) + 8
        Else
            pie_SetGeometricOffset _
                (ButXPos + iV_GetImageWidth(IntImages, IMAGE_BUTB0_UP) \ 2), _
                (ButYPos + iV_GetImageHeight(IntImages, IMAGE_BUTB0_UP) \ 2) + 12
        End If
    End If
End Sub

' Create a button by rendering an IMD object into it.
Public Sub IntFancyButton_displayIMD(ByVal self As WIDGET, image As AtlasImage, imdObject_ As ImdObject, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim ButXPos As Long, ButYPos As Long, ox As Long, oy As Long, Radius As Long, basePlateSize As Long
    Dim IMDType_ As ImdType, Object_ As Long, d As IntFancyButton, rot As Vector3i, posv As Vector3i, imdRadius As Long

    If ImdObject_empty(imdObject_) Then
        IntFancyButton_displayImage self, image, xOffset, yOffset
        Exit Sub
    End If
    Set d = IntFancyButton_(self)
    ButXPos = xOffset + self.x()
    ButYPos = yOffset + self.y()

    If W_CLICKFORM_isDown(self) Then
        ox = 2: oy = 2
    Else
        ox = 0: oy = 0
    End If

    IMDType_ = imdObject_.typ
    Object_ = imdObject_.ptr
    If Not (Object_ <> 0) Then ASSERT_FAILURE "Object != nullptr", "src/intdisplay.cpp:994", "displayIMD", "imdObject.ptr is null?": Exit Sub
    If IMDType_ = IMDTYPE_DROID Or IMDType_ = IMDTYPE_DROIDTEMPLATE Then
        ' The case where we have to render a composite droid.
        IntFancyButton_setModelOffset self, ButXPos, ButYPos

        If IMDType_ = IMDTYPE_DROID Then
            Radius = getComponentDroidRadius(Object_)
        Else
            Radius = getComponentDroidTemplateRadius(castDroidTemplate(Object_))
        End If

        d.model_scale = DROID_BUT_SCALE
        If Not (Radius <= 128) Then ASSERT_FAILURE "Radius <= 128", "src/intdisplay.cpp:1041", "displayIMD", "create PIE button big component found"

        IntFancyButton_displayClear self, xOffset, yOffset

        If IMDType_ = IMDTYPE_DROID Then
            If DROID_isTransporter(Object_) Then
                If aDroid(Object_).droidType = DROID_TRANSPORTER Then
                    d.model_scale = DROID_BUT_SCALE \ 2
                Else
                    d.model_scale = DROID_BUT_SCALE \ 3
                End If
            End If
        Else '(IMDType_ == IMDTYPE_DROIDTEMPLATE)
            If isTransporter_template(apsTemplates(castDroidTemplate(Object_))) Then
                If apsTemplates(castDroidTemplate(Object_)).droidType = DROID_TRANSPORTER Then
                    d.model_scale = DROID_BUT_SCALE \ 2
                Else
                    d.model_scale = DROID_BUT_SCALE \ 3
                End If
            End If
        End If

        rot = d.model_rotation
        posv = d.model_position
        'lefthand display droid buttons
        If IMDType_ = IMDTYPE_DROID Then
            displayComponentButtonObject Object_, rot, posv, d.model_scale
        Else
            displayComponentButtonTemplate castDroidTemplate(Object_), rot, posv, d.model_scale
        End If
    Else
        ' Just drawing a single IMD.
        IntFancyButton_setSingleModelOffset self, ButXPos, ButYPos

        ' Decide which button grid size to use.
        If IMDType_ = IMDTYPE_COMPONENT Then
            Radius = getComponentRadius(Object_)
            d.model_scale = rescaleButtonObject(Radius, COMP_BUT_SCALE, COMPONENT_RADIUS)
            ' NOTE: The Super transport is huge, and is considered a component type, so refit it to inside the button.
            If BASE_STATS_id(Object_) = "SuperTransportBody" Then
                d.model_scale = Fix(CSng(d.model_scale) * 0.4!)
            ElseIf BASE_STATS_id(Object_) = "TransporterBody" Then
                d.model_scale = Fix(CSng(d.model_scale) * 0.6!)
            End If
        ElseIf IMDType_ = IMDTYPE_RESEARCH Then
            Radius = getResearchRadius(Object_)
            If Radius <= 100 Then
                d.model_scale = rescaleButtonObject(Radius, COMP_BUT_SCALE, COMPONENT_RADIUS)
            ElseIf Radius <= 128 Then
                d.model_scale = SMALL_STRUCT_SCALE
            ElseIf Radius <= 256 Then
                d.model_scale = MED_STRUCT_SCALE
            Else
                d.model_scale = LARGE_STRUCT_SCALE
            End If
        ElseIf IMDType_ = IMDTYPE_STRUCTURE Then
            basePlateSize = getStructureSizeMax(Object_)
            If basePlateSize = 1 Then
                d.model_scale = SMALL_STRUCT_SCALE
            ElseIf basePlateSize = 2 Then
                d.model_scale = MED_STRUCT_SCALE
            Else
                d.model_scale = LARGE_STRUCT_SCALE
            End If
        ElseIf IMDType_ = IMDTYPE_STRUCTURESTAT Then
            basePlateSize = getStructureStatSizeMax(castStructureStats(Object_))
            If basePlateSize = 1 Then
                d.model_scale = SMALL_STRUCT_SCALE
            ElseIf basePlateSize = 2 Then
                d.model_scale = MED_STRUCT_SCALE
            Else
                d.model_scale = LARGE_STRUCT_SCALE
            End If
        ElseIf IMDType_ = IMDTYPE_FEATURE Then
            imdRadius = iIMDShape_sradius(Object_)

            If imdRadius <= 40 Then
                d.model_scale = ULTRA_SMALL_FEATURE_SCALE
            ElseIf imdRadius <= 64 Then
                d.model_scale = REALLY_SMALL_FEATURE_SCALE
            ElseIf imdRadius <= 128 Then
                d.model_scale = SMALL_FEATURE_SCALE
            ElseIf imdRadius <= 256 Then
                d.model_scale = MED_FEATURE_SCALE
            Else
                d.model_scale = LARGE_FEATURE_SCALE
            End If
        Else
            Radius = iIMDShape_sradius(Object_)

            If Radius <= 128 Then
                d.model_scale = SMALL_STRUCT_SCALE
            ElseIf Radius <= 256 Then
                d.model_scale = MED_STRUCT_SCALE
            Else
                d.model_scale = LARGE_STRUCT_SCALE
            End If
        End If

        IntFancyButton_displayClear self, xOffset, yOffset

        If Not AtlasImage_isNull(image) Then
            iV_DrawImageImage image.images, image.id, ButXPos + ox, ButYPos + oy
        End If

        rot = d.model_rotation
        posv = d.model_position
        ' all non droid buttons
        If IMDType_ = IMDTYPE_COMPONENT Then
            displayComponentButton Object_, rot, posv, d.model_scale
        ElseIf IMDType_ = IMDTYPE_RESEARCH Then
            displayResearchButton Object_, rot, posv, d.model_scale
        ElseIf IMDType_ = IMDTYPE_STRUCTURE Then
            displayStructureButton Object_, rot, posv, d.model_scale
        ElseIf IMDType_ = IMDTYPE_STRUCTURESTAT Then
            displayStructureStatButton castStructureStats(Object_), rot, posv, d.model_scale
        ElseIf IMDType_ = IMDTYPE_FEATURE Then
            displayIMDButton Object_, rot, posv, d.model_scale
        Else
            displayIMDButton Object_, rot, posv, d.model_scale
        End If
    End If
End Sub

' Create a button by rendering an image into it.
Public Sub IntFancyButton_displayImage(ByVal self As WIDGET, image As AtlasImage, ByVal xOffset As Long, ByVal yOffset As Long)
    If AtlasImage_isNull(image) Then
        IntFancyButton_displayBlank self, xOffset, yOffset
        Exit Sub
    End If

    IntFancyButton_displayClear self, xOffset, yOffset
    iV_DrawImageImage image.images, image.id, xOffset + self.x(), yOffset + self.y()
End Sub

' Create a blank button.
Public Sub IntFancyButton_displayBlank(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long, Optional ByVal withQuestionMark As Boolean = True)
    Dim ox As Long, oy As Long

    If W_CLICKFORM_isDown(self) Then
        ox = 1: oy = 1
    Else
        ox = 0: oy = 0
    End If

    IntFancyButton_displayClear self, xOffset, yOffset

    If withQuestionMark Then
        ' Draw a question mark, bit of quick hack this.
        iV_DrawImage IntImages, IMAGE_QUESTION_MARK, xOffset + self.x() + ox + 10, yOffset + self.y() + oy + 3
    End If
End Sub

' MARK: - the droids, structures and stats

' Returns true if the droid is currently building something.
'
Public Function DroidIsBuilding(ByVal Droid_ As Long) As Boolean
    Dim Stats As Long
    If Not (Droid_ <> 0) Then ASSERT_FAILURE "(Droid) != nullptr", "src/intdisplay.cpp:1306", "DroidIsBuilding", "%s should not be nullptr", "Droid": DroidIsBuilding = False: Exit Function

    If Not (aDroid(Droid_).droidType = DROID_CONSTRUCT Or _
            aDroid(Droid_).droidType = DROID_CYBORG_CONSTRUCT) Then
        DroidIsBuilding = False
        Exit Function
    End If

    If orderStateStatsLoc(Droid_, DORDER_BUILD, Stats) Then
        ' Moving to build location?
        DroidIsBuilding = False
        Exit Function
    ElseIf orderStateObj(Droid_, DORDER_BUILD) <> 0 _
         Or orderStateObj(Droid_, DORDER_HELPBUILD) <> 0 Then ' Is building or helping?
        DroidIsBuilding = True
        Exit Function
    End If

    DroidIsBuilding = False
End Function


' Returns true if the droid has been ordered build something ( but hasn't started yet )
'
Public Function DroidGoingToBuild(ByVal Droid_ As Long) As Boolean
    Dim Stats As Long
    If Not (Droid_ <> 0) Then ASSERT_FAILURE "(Droid) != nullptr", "src/intdisplay.cpp:1334", "DroidGoingToBuild", "%s should not be nullptr", "Droid": DroidGoingToBuild = False: Exit Function

    If Not (aDroid(Droid_).droidType = DROID_CONSTRUCT Or _
            aDroid(Droid_).droidType = DROID_CYBORG_CONSTRUCT) Then
        DroidGoingToBuild = False
        Exit Function
    End If

    If orderStateStatsLoc(Droid_, DORDER_BUILD, Stats) Then ' Moving to build location?
        DroidGoingToBuild = True
        Exit Function
    End If

    DroidGoingToBuild = False
End Function


' Get the structure for a structure which a droid is currently building.
'
Public Function DroidGetBuildStructure(ByVal Droid_ As Long) As Long
    Dim actionTarget As Long
    If orderStateObj(Droid_, DORDER_BUILD) <> 0 Then
        DroidGetBuildStructure = orderStateObj(Droid_, DORDER_HELPBUILD)
        Exit Function
    End If

    If aDroid(Droid_).action = DACTION_BUILD Then
        actionTarget = aDroid(Droid_).psActionTarget(0)
        If actionTarget <> 0 Then
            If SIMPLE_OBJECT_type(actionTarget) = OBJ_STRUCTURE Then
                DroidGetBuildStructure = actionTarget
                Exit Function
            End If
        End If
    End If

    DroidGetBuildStructure = 0
End Function

' (VB) droidGetCommandFactory's search of the player's structures: the factory of the type whose
' assembly point is the inc'th
Private Function droidGetCommandFactory_find(ByVal player As Long, ByVal typ As STRUCTURE_TYPE, ByVal inc As Long) As Long
    Dim it As Long, psCurr As Long
    it = apsStructLists(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If (asStructureStats(aStruct(psCurr).pStructureType).typ = typ) Then
            If aFlagPos(aStruct(psCurr).functionality.factory.psAssemblyPoint).factoryInc = inc Then
                droidGetCommandFactory_find = psCurr
                Exit Function
            End If
        End If
        it = lnNext(it)
    Loop
End Function

' Get the first factory assigned to a command droid
Public Function droidGetCommandFactory(ByVal psDroid As Long) As Long
    Dim inc As Long, psCurr As Long

    For inc = 0 To MAX_FACTORY - 1
        If (aDroid(psDroid).secondaryOrder And SHL(1, inc + DSS_ASSPROD_SHIFT)) <> 0 Then
            ' found an assigned factory - look for it in the lists
            psCurr = droidGetCommandFactory_find(aDroid(psDroid).player, REF_FACTORY, inc)
            If psCurr <> 0 Then
                droidGetCommandFactory = psCurr
                Exit Function
            End If
        End If
        If (aDroid(psDroid).secondaryOrder And SHL(1, inc + DSS_ASSPROD_CYBORG_SHIFT)) <> 0 Then
            ' found an assigned factory - look for it in the lists
            psCurr = droidGetCommandFactory_find(aDroid(psDroid).player, REF_CYBORG_FACTORY, inc)
            If psCurr <> 0 Then
                droidGetCommandFactory = psCurr
                Exit Function
            End If
        End If
        If (aDroid(psDroid).secondaryOrder And SHL(1, inc + DSS_ASSPROD_VTOL_SHIFT)) <> 0 Then
            ' found an assigned factory - look for it in the lists
            psCurr = droidGetCommandFactory_find(aDroid(psDroid).player, REF_VTOL_FACTORY, inc)
            If psCurr <> 0 Then
                droidGetCommandFactory = psCurr
                Exit Function
            End If
        End If
    Next

    droidGetCommandFactory = 0
End Function

' Get the stats for a structure which a droid is going to ( but not yet ) building.
' (VB) a BASE_STATS * handle
Public Function DroidGetBuildStats(ByVal Droid_ As Long) As Long
    Dim Stats As Long

    If orderStateStatsLoc(Droid_, DORDER_BUILD, Stats) Then ' Moving to build location?
        DroidGetBuildStats = structureStatsBase(Stats)
        Exit Function
    End If

    DroidGetBuildStats = 0
End Function

' iIMDBaseShape *
Public Function DroidGetIMD(ByVal Droid_ As Long) As Long
    DroidGetIMD = aDroid(Droid_).sDisplay.imd
End Function

Private Function structureIsManufacturingPending_factory(func_ As FACTORY) As Boolean
    If func_.statusPending <> FACTORY_NOTHING_PENDING Then
        structureIsManufacturingPending_factory = func_.statusPending = FACTORY_START_PENDING Or func_.statusPending = FACTORY_HOLD_PENDING
        Exit Function
    End If
    structureIsManufacturingPending_factory = func_.psSubject <> 0
End Function

Private Function structureIsManufacturingPending_research(func_ As RESEARCH_FACILITY) As Boolean
    If func_.statusPending <> FACTORY_NOTHING_PENDING Then
        structureIsManufacturingPending_research = func_.statusPending = FACTORY_START_PENDING Or func_.statusPending = FACTORY_HOLD_PENDING
        Exit Function
    End If
    structureIsManufacturingPending_research = func_.psSubject <> -1
End Function

Public Function StructureIsManufacturingPending(ByVal structure_ As Long) As Boolean
    If Not (structure_ <> 0) Then ASSERT_FAILURE "(structure) != nullptr", "src/intdisplay.cpp:1454", "StructureIsManufacturingPending", "%s should not be nullptr", "structure": StructureIsManufacturingPending = False: Exit Function
    Select Case asStructureStats(aStruct(structure_).pStructureType).typ
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        StructureIsManufacturingPending = structureIsManufacturingPending_factory(aStruct(structure_).functionality.factory)
    Case Else
        StructureIsManufacturingPending = False
    End Select
End Function

' FACTORY *: the structure's handle
Public Function StructureGetFactory(ByVal Structure_ As Long) As Long
    If Not (Structure_ <> 0) Then ASSERT_FAILURE "(Structure) != nullptr", "src/intdisplay.cpp:1468", "StructureGetFactory", "%s should not be nullptr", "Structure": StructureGetFactory = 0: Exit Function
    StructureGetFactory = Structure_       ' &Structure->pFunctionality->factory
End Function

Public Function structureIsResearchingPending(ByVal structure_ As Long) As Boolean
    If Not (structure_ <> 0) Then ASSERT_FAILURE "(structure) != nullptr", "src/intdisplay.cpp:1474", "structureIsResearchingPending", "%s should not be nullptr", "structure": structureIsResearchingPending = False: Exit Function
    If asStructureStats(aStruct(structure_).pStructureType).typ = REF_RESEARCH Then
        structureIsResearchingPending = structureIsManufacturingPending_research(aStruct(structure_).functionality.researchFacility)
    End If
End Function

Private Function structureIsOnHoldPending_factory(func_ As FACTORY) As Boolean
    If func_.statusPending <> FACTORY_NOTHING_PENDING Then
        structureIsOnHoldPending_factory = func_.statusPending = FACTORY_HOLD_PENDING
        Exit Function
    End If
    structureIsOnHoldPending_factory = func_.timeStartHold <> 0
End Function

Private Function structureIsOnHoldPending_research(func_ As RESEARCH_FACILITY) As Boolean
    If func_.statusPending <> FACTORY_NOTHING_PENDING Then
        structureIsOnHoldPending_research = func_.statusPending = FACTORY_HOLD_PENDING
        Exit Function
    End If
    structureIsOnHoldPending_research = func_.timeStartHold <> 0
End Function

Public Function StructureIsOnHoldPending(ByVal structure_ As Long) As Boolean
    If Not (structure_ <> 0) Then ASSERT_FAILURE "(structure) != nullptr", "src/intdisplay.cpp:1490", "StructureIsOnHoldPending", "%s should not be nullptr", "structure": StructureIsOnHoldPending = False: Exit Function
    Select Case asStructureStats(aStruct(structure_).pStructureType).typ
    Case REF_FACTORY, REF_CYBORG_FACTORY, REF_VTOL_FACTORY
        StructureIsOnHoldPending = structureIsOnHoldPending_factory(aStruct(structure_).functionality.factory)
    Case REF_RESEARCH
        StructureIsOnHoldPending = structureIsOnHoldPending_research(aStruct(structure_).functionality.researchFacility)
    Case Else
        ASSERT_FAILURE "false", "src/intdisplay.cpp:1500", "StructureIsOnHoldPending", "Huh?"
        StructureIsOnHoldPending = False
    End Select
End Function

' RESEARCH_FACILITY *: the structure's handle
Public Function StructureGetResearch(ByVal Structure_ As Long) As Long
    If Not (Structure_ <> 0) Then ASSERT_FAILURE "(Structure) != nullptr", "src/intdisplay.cpp:1507", "StructureGetResearch", "%s should not be nullptr", "Structure": StructureGetResearch = 0: Exit Function
    If Not (aStruct(Structure_).pFunctionality <> 0) Then ASSERT_FAILURE "(Structure->pFunctionality) != nullptr", "src/intdisplay.cpp:1508", "StructureGetResearch", "%s should not be nullptr", "Structure->pFunctionality": StructureGetResearch = 0: Exit Function
    StructureGetResearch = Structure_      ' &Structure->pFunctionality->researchFacility
End Function


' Factory: the FACTORY *, its structure's handle
Public Function FactoryGetTemplate(ByVal Factory_ As Long) As Long
    If Not (Factory_ <> 0) Then ASSERT_FAILURE "(Factory) != nullptr", "src/intdisplay.cpp:1515", "FactoryGetTemplate", "%s should not be nullptr", "Factory": FactoryGetTemplate = 0: Exit Function
    If aStruct(Factory_).functionality.factory.psSubjectPending <> 0 Then
        FactoryGetTemplate = aStruct(Factory_).functionality.factory.psSubjectPending
        Exit Function
    End If

    FactoryGetTemplate = aStruct(Factory_).functionality.factory.psSubject
End Function

' Stat: a BASE_STATS * handle
Public Function StatIsStructure(ByVal Stat As Long) As Boolean
    StatIsStructure = BASE_STATS_hasType(Stat, STAT_STRUCTURE)
End Function

Public Function StatIsFeature(ByVal Stat As Long) As Boolean
    StatIsFeature = BASE_STATS_hasType(Stat, STAT_FEATURE)
End Function

' iIMDBaseShape *
Public Function StatGetStructureIMD(ByVal Stat As Long, ByVal playerNum As Long) As Long ' playerNum: Player (unused)
    StatGetStructureIMD = asStructureStats(Stat And &HFFFF&).pIMD(0)   ' ((STRUCTURE_STATS *)Stat)->pIMD[0]
End Function

Public Function StatIsTemplate(ByVal Stat As Long) As Boolean
    StatIsTemplate = BASE_STATS_hasType(Stat, STAT_TEMPLATE)
End Function

Public Function StatIsComponent(ByVal Stat As Long) As COMPONENT_TYPE
    Select Case BASE_STATS_ref(Stat) And STAT_MASK
    Case STAT_BODY: StatIsComponent = COMP_BODY
    Case STAT_BRAIN: StatIsComponent = COMP_BRAIN
    Case STAT_PROPULSION: StatIsComponent = COMP_PROPULSION
    Case STAT_WEAPON: StatIsComponent = COMP_WEAPON
    Case STAT_SENSOR: StatIsComponent = COMP_SENSOR
    Case STAT_ECM: StatIsComponent = COMP_ECM
    Case STAT_CONSTRUCT: StatIsComponent = COMP_CONSTRUCT
    Case STAT_REPAIR: StatIsComponent = COMP_REPAIRUNIT
    Case Else: StatIsComponent = COMP_NUMCOMPONENTS
    End Select
End Function

' CompIMD, MountIMD: the const iIMDShape * results (0 = nullptr)
Public Function StatGetComponentIMD(ByVal Stat As Long, ByVal compID As Long, CompIMD As Long, MountIMD As Long) As Boolean ' DISPLAY ONLY
    Dim psWStat As Long, i As Long

    CompIMD = 0
    MountIMD = 0
    i = Stat And &HFFFF&

    Select Case compID
    Case COMP_BODY
        CompIMD = safeGetDisplayModelFromBase(COMPONENT_STATS_pIMD(Stat))
        StatGetComponentIMD = True
        Exit Function

    Case COMP_BRAIN
        psWStat = asBrainStats(i).psWeaponStat
        MountIMD = safeGetDisplayModelFromBase(asWeaponStats(psWStat).pMountGraphic)
        CompIMD = safeGetDisplayModelFromBase(asWeaponStats(psWStat).pIMD)
        StatGetComponentIMD = True
        Exit Function

    Case COMP_WEAPON
        MountIMD = safeGetDisplayModelFromBase(asWeaponStats(i).pMountGraphic)
        CompIMD = safeGetDisplayModelFromBase(COMPONENT_STATS_pIMD(Stat))
        StatGetComponentIMD = True
        Exit Function

    Case COMP_SENSOR
        MountIMD = safeGetDisplayModelFromBase(asSensorStats(i).pMountGraphic)
        CompIMD = safeGetDisplayModelFromBase(COMPONENT_STATS_pIMD(Stat))
        StatGetComponentIMD = True
        Exit Function

    Case COMP_ECM
        MountIMD = safeGetDisplayModelFromBase(asECMStats(i).pMountGraphic)
        CompIMD = safeGetDisplayModelFromBase(COMPONENT_STATS_pIMD(Stat))
        StatGetComponentIMD = True
        Exit Function

    Case COMP_CONSTRUCT
        MountIMD = safeGetDisplayModelFromBase(asConstructStats(i).pMountGraphic)
        CompIMD = safeGetDisplayModelFromBase(COMPONENT_STATS_pIMD(Stat))
        StatGetComponentIMD = True
        Exit Function

    Case COMP_PROPULSION
        CompIMD = safeGetDisplayModelFromBase(COMPONENT_STATS_pIMD(Stat))
        StatGetComponentIMD = True
        Exit Function

    Case COMP_REPAIRUNIT
        MountIMD = safeGetDisplayModelFromBase(asRepairStats(i).pMountGraphic)
        CompIMD = safeGetDisplayModelFromBase(COMPONENT_STATS_pIMD(Stat))
        StatGetComponentIMD = True
        Exit Function

    Case COMP_NUMCOMPONENTS
        ASSERT_FAILURE "false", "src/intdisplay.cpp:1610", "StatGetComponentIMD", "Unknown component"
    End Select

    StatGetComponentIMD = False
End Function


Public Function StatIsResearch(ByVal Stat As Long) As Boolean
    StatIsResearch = BASE_STATS_hasType(Stat, STAT_RESEARCH)
End Function

' psStat: a RESEARCH's BASE_STATS * handle; Shape: the const iIMDShape *, ppGraphicData: a BASE_STATS * (0 = nullptr)
Private Sub StatGetResearchImage(ByVal psStat As Long, image As AtlasImage, Shape As Long, ppGraphicData As Long, ByVal drawTechIcon As Boolean) ' DISPLAY ONLY
    Dim i As Long
    i = psStat And &HFFFF&              ' (const RESEARCH *)psStat
    If drawTechIcon And asResearch(i).iconID <> NO_RESEARCH_ICON Then
        image = AtlasImage_(IntImages, asResearch(i).iconID)
    End If
    'if the research has a Stat associated with it - use this as display in the button
    If asResearch(i).psStat <> 0 Then
        ppGraphicData = asResearch(i).psStat
        'make sure the IMDShape is initialised
        Shape = 0
    Else
        'no stat so just just the IMD associated with the research
        Shape = safeGetDisplayModelFromBase(asResearch(i).pIMD)
        'make sure the stat is initialised
        ppGraphicData = 0
    End If
End Sub

' Widget callback function to play an audio track.
'
Public Sub WidgetAudioCallback(ByVal AudioID As Long)
    Dim TimeSinceLastWidgetBeep As Long
    If AudioID >= 0 Then
        ' Don't allow a widget beep if one was made in the last WIDGETBEEPGAP milliseconds
        ' This stops double beeps happening (which seems to happen all the time)
        TimeSinceLastWidgetBeep = realTime - LastTimeAudio
        If TimeSinceLastWidgetBeep < 0 Or TimeSinceLastWidgetBeep > WIDGETBEEPGAP Then
            LastTimeAudio = realTime
            audio_PlayTrack AudioID
        End If
    End If
End Sub

' MARK: - IntTransportButton

Public Sub IntTransportButton_construct(ByVal self As WIDGET)
    IntFancyButton_construct self
    self.kind = K_IntTransportButton
    IntTransportButton_(self).psDroid = 0
End Sub

Public Function IntTransportButton_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_IntTransportButton)
    IntTransportButton_construct widget
    Set IntTransportButton_new = widget
End Function

' object: a DROID *
Public Sub IntTransportButton_setObject(ByVal self As WIDGET, ByVal object_ As Long)
    IntTransportButton_(self).psDroid = object_
End Sub

' Widget callback to display a contents button for the Transporter
Public Sub IntTransportButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim psDroid As Long, gfxId As Long, noImage As AtlasImage, object_ As ImdObject
    psDroid = IntTransportButton_(self).psDroid
    ' There should always be a droid associated with the button
    If Not (psDroid <> 0) Then ASSERT_FAILURE "psDroid != nullptr", "src/intdisplay.cpp:1674", "display", "Invalid droid pointer"

    IntFancyButton_initDisplay self
    object_ = ImdObject_Droid(psDroid)
    IntFancyButton_displayIMD self, noImage, object_, xOffset, yOffset
    IntFancyButton_displayIfHighlight self, xOffset, yOffset

    If psDroid <> 0 Then
        ' Add the experience level for each droid
        gfxId = getDroidRankGraphic(psDroid)
        If gfxId <> UDWORD_MAX Then
            ' Render the rank graphic at the correct location
            iV_DrawImage IntImages, gfxId, xOffset + self.x() + 50, yOffset + self.y() + 10
        End If
    End If
End Sub

' MARK: - the radar's blips

' (VB) drawRadarBlips's static tables: images[i] of imagesEnemy (0), imagesResource (1), imagesArtifact (2)
' (imagesProxTypes[proxType] is table proxType) and imagesBurningResource (3)
Private Function radarBlipImage(ByVal table As Long, ByVal i As Long) As Long
    Select Case table
    Case 0: radarBlipImage = Choose(i + 1, IMAGE_RAD_ENMREAD, IMAGE_RAD_ENM1, IMAGE_RAD_ENM2, IMAGE_RAD_ENM3)
    Case 1: radarBlipImage = Choose(i + 1, IMAGE_RAD_RESREAD, IMAGE_RAD_RES1, IMAGE_RAD_RES2, IMAGE_RAD_RES3)
    Case 2: radarBlipImage = Choose(i + 1, IMAGE_RAD_ARTREAD, IMAGE_RAD_ART1, IMAGE_RAD_ART2, IMAGE_RAD_ART3)
    Case 3: radarBlipImage = Choose(i + 1, IMAGE_RAD_BURNRESREAD, IMAGE_RAD_BURNRES1, IMAGE_RAD_BURNRES2, IMAGE_RAD_BURNRES3, IMAGE_RAD_BURNRES4, IMAGE_RAD_BURNRES5, IMAGE_RAD_BURNRES6)
    End Select
End Function

' Draws blips on radar to represent Proximity Display and damaged structures
Public Sub drawRadarBlips(ByVal radarX As Long, ByVal radarY As Long, ByVal pixSizeH As Single, ByVal pixSizeV As Single, modelViewProjection As mat4)
    Dim imageID As Long, delay As Long, width_ As Long, height_ As Long, x As Long, y As Long
    Dim it As Long, psProxDisp As Long, psMessage As Long, animationLength As Long, images As Long
    Dim psFeature As Long, psObj As Long, objPos As Vector3i, strobe As Long, c As Boolean
    delay = 150
    x = 0: y = 0

    ' store the width & height of the radar/mini-map
    width_ = scrollMaxX - scrollMinX
    height_ = scrollMaxY - scrollMinY

    ' Go through all the proximity Displays
    If selectedPlayer < MAX_PLAYERS Then
        it = apsProxDisp(selectedPlayer).head
        Do While it <> 0
            psProxDisp = lnObj(it)
            psMessage = aProxDisp(psProxDisp).psMessage
            animationLength = 4 - 1     ' ARRAY_SIZE(imagesEnemy) - 1: Same size as imagesResource and imagesArtifact.

            If aMessage(psMessage).player <> selectedPlayer Then
                GoTo continue_
            End If

            If aProxDisp(psProxDisp).typ = POS_PROXDATA Then
                If aMessage(psMessage).pViewData = 0 Then
                    GoTo continue_ ' if no data - ignore message
                End If
                images = aViewData(aMessage(psMessage).pViewData).prox_proxType    ' imagesProxTypes[proxType]
            Else
                psFeature = castFeature(aMessage(psMessage).psObj)

                c = psFeature <> 0
                If c Then c = aFeature(psFeature).psStats >= 0
                If Not (c) Then ASSERT_FAILURE "psFeature && psFeature->psStats", "src/intdisplay.cpp:1735", "drawRadarBlips", "Bad feature message": Exit Sub
                If asFeatureStats(aFeature(psFeature).psStats).subType = FEAT_OIL_RESOURCE Then
                    images = 1              ' imagesResource
                    If fireOnLocation(aFeature(psFeature).pos.x, aFeature(psFeature).pos.y) Then
                        images = 3          ' imagesBurningResource
                        animationLength = 7 - 1 ' Longer animation for burning oil wells.
                    End If
                Else
                    images = 2              ' imagesArtifact
                End If
            End If

            ' Draw the 'blips' on the radar - use same timings as radar blips if the message is read - don't animate
            If aMessage(psMessage).read Then
                imageID = radarBlipImage(images, 0)
            Else
                ' Draw animated
                If ULT(delay, realTime - aProxDisp(psProxDisp).timeLastDrawn) Then
                    aProxDisp(psProxDisp).strobe = aProxDisp(psProxDisp).strobe + 1
                    aProxDisp(psProxDisp).timeLastDrawn = realTime
                End If
                aProxDisp(psProxDisp).strobe = aProxDisp(psProxDisp).strobe Mod animationLength
                imageID = radarBlipImage(images, 1 + aProxDisp(psProxDisp).strobe)
            End If

            If aProxDisp(psProxDisp).typ = POS_PROXDATA Then
                ' const VIEW_PROXIMITY* psViewProx
                x = Fix((aViewData(aMessage(psMessage).pViewData).prox_x \ TILE_UNITS - scrollMinX) * pixSizeH)
                y = Fix((aViewData(aMessage(psMessage).pViewData).prox_y \ TILE_UNITS - scrollMinY) * pixSizeV)
            ElseIf aProxDisp(psProxDisp).typ = POS_PROXOBJ Then
                psObj = aMessage(psMessage).psObj
                objPos = SIMPLE_OBJECT_pos(psObj)
                x = Fix((objPos.x \ TILE_UNITS - scrollMinX) * pixSizeH)
                y = Fix((objPos.y \ TILE_UNITS - scrollMinY) * pixSizeV)
            Else
                ASSERT_FAILURE "false", "src/intdisplay.cpp:1782", "drawRadarBlips", "Bad message type"
                GoTo continue_
            End If

            ' NOTE:  On certain missions (limbo & expand), there is still valid data that is stored outside the
            ' normal radar/mini-map view.  We must now calculate the radar/mini-map's bounding box, and clip
            ' everything outside the box.
            If (x + radarX) < width_ * pixSizeV / 2 And (x + radarX) > -width_ * pixSizeV / 2 _
               And (y + radarY) < height_ * pixSizeH / 2 And (y + radarY) > -height_ * pixSizeH / 2 Then
                ' Draw the 'blip'
                iV_DrawImage IntImages, imageID, x + radarX, y + radarY, VarPtr(modelViewProjection)
            End If
continue_:
            it = lnNext(it)
        Loop
    End If
    If audio_GetPreviousQueueTrackRadarBlipPos(x, y) Then
        animationLength = 4 - 1
        strobe = (realTime \ delay) Mod animationLength
        x = Fix((x \ TILE_UNITS - scrollMinX) * pixSizeH)
        y = Fix((y \ TILE_UNITS - scrollMinY) * pixSizeV)
        imageID = radarBlipImage(0, strobe)

        ' NOTE:  On certain missions (limbo & expand), there is still valid data that is stored outside the
        ' normal radar/mini-map view.  We must now calculate the radar/mini-map's bounding box, and clip
        ' everything outside the box.
        If (x + radarX) < width_ * pixSizeV / 2 And (x + radarX) > -width_ * pixSizeV / 2 _
           And (y + radarY) < height_ * pixSizeH / 2 And (y + radarY) > -height_ * pixSizeH / 2 Then
            ' Draw the 'blip'
            iV_DrawImage IntImages, imageID, x + radarX, y + radarY, VarPtr(modelViewProjection)
        End If
    End If
End Sub


'Displays the proximity messages blips over the world
Public Sub intDisplayProximityBlips(ByVal psWidget As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim psButton As WIDGET, psProxDisp As Long, psMsg As Long, x As Long, y As Long, objPos As Vector3i
    Set psButton = psWidget
    psProxDisp = psButton.pUserData
    psMsg = aProxDisp(psProxDisp).psMessage
    x = 0: y = 0

    If Not (aMessage(psMsg).typ = MSG_PROXIMITY) Then ASSERT_FAILURE "psMsg->type == MSG_PROXIMITY", "src/intdisplay.cpp:1826", "intDisplayProximityBlips", "Invalid message type"

    'if no data - ignore message
    If aMessage(psMsg).pViewData = 0 Or aMessage(psMsg).player <> selectedPlayer Then
        Exit Sub
    End If
    If aProxDisp(psProxDisp).typ = POS_PROXDATA Then
        x = aViewData(aMessage(psMsg).pViewData).prox_x
        y = aViewData(aMessage(psMsg).pViewData).prox_y
    ElseIf aProxDisp(psProxDisp).typ = POS_PROXOBJ Then
        objPos = SIMPLE_OBJECT_pos(aMessage(psMsg).psObj)
        x = objPos.x
        y = objPos.y
    End If

    'if not within view ignore message
    If Not clipXY(x, y) Then
        Exit Sub
    End If

    'if the message is read - don't draw
    If Not aMessage(psMsg).read Then
        'set the button's x/y so that can be clicked on
        psButton.move aProxDisp(psProxDisp).screenX - psButton.width() \ 2, aProxDisp(psProxDisp).screenY - psButton.height() \ 2
    End If
End Sub

' MARK: - the quantity sliders

Private Function sliderMouseUnit(ByVal Slider As WIDGET) As Long
    Dim posStops As Long
    posStops = (W_SLIDER_(Slider).numStops \ 20) And &HFFFF&

    If posStops = 0 Or W_SLIDER_(Slider).pos = 0 Or W_SLIDER_(Slider).pos = W_SLIDER_(Slider).numStops Then
        sliderMouseUnit = 1
        Exit Function
    End If

    If W_SLIDER_(Slider).pos < posStops Then
        sliderMouseUnit = W_SLIDER_(Slider).pos
        Exit Function
    End If

    If W_SLIDER_(Slider).pos > (W_SLIDER_(Slider).numStops - posStops) Then
        sliderMouseUnit = (W_SLIDER_(Slider).numStops - W_SLIDER_(Slider).pos) And &HFFFF&
        Exit Function
    End If
    sliderMouseUnit = posStops
End Function

' psWidget: the W_SLIDER
Public Sub intUpdateQuantitySlider(ByVal psWidget As WIDGET, psContext As W_CONTEXT)
    Dim Slider As WIDGET
    Set Slider = psWidget

    If W_SLIDER_isHighlighted(Slider) Then
        If keyDown(KEY_LEFTARROW) Then
            If W_SLIDER_(Slider).pos > 0 Then
                W_SLIDER_(Slider).pos = (W_SLIDER_(Slider).pos - sliderMouseUnit(Slider)) And &HFFFF&
            End If
        ElseIf keyDown(KEY_RIGHTARROW) Then
            If W_SLIDER_(Slider).pos < W_SLIDER_(Slider).numStops Then
                W_SLIDER_(Slider).pos = (W_SLIDER_(Slider).pos + sliderMouseUnit(Slider)) And &HFFFF&
            End If
        End If
    End If
End Sub

' psBar: a W_BARGRAPH; researches: listAllyResearch's (the first researches_size)
Public Sub intDisplayUpdateAllyBar(ByVal psBar As WIDGET, research_ As RESEARCH, researches() As AllyResearch, ByVal researches_size As Long)
    Const researchNotStarted As Long = 3600000
    Dim bestCompletion As Long, bestPowerNeeded As Long, bestTimeToResearch As Long, researchPowerCost As Long, k As Long
    bestCompletion = 0
    bestPowerNeeded = researchNotStarted
    bestTimeToResearch = researchNotStarted
    researchPowerCost = researchNotStarted
    For k = 0 To researches_size - 1
        If ULT(bestCompletion, researches(k).completion) Then
            bestCompletion = researches(k).completion
            W_BARGRAPH_(psBar).majorCol = pal_GetTeamColour(getPlayerColour(researches(k).player))
        End If
        If Not researches(k).active Then
            GoTo continue_  ' Don't show remaining time/power, if the facility is currently being upgraded.
        End If
        If researches(k).powerNeeded = -1 Then
            If ULT(researches(k).timeToResearch, bestTimeToResearch) Then bestTimeToResearch = researches(k).timeToResearch ' std::min<unsigned>
        Else
            If ULT(researches(k).powerNeeded, bestPowerNeeded) Then bestPowerNeeded = researches(k).powerNeeded ' std::min<unsigned>
            researchPowerCost = research_.researchPower
        End If
continue_:
    Next

    setBarGraphValue psBar, W_BARGRAPH_(psBar).majorCol, bestCompletion, research_.researchPoints
    If bestTimeToResearch <> researchNotStarted Then
        ' Show research progress.
        formatTimeText psBar, bestTimeToResearch
    ElseIf bestCompletion > 0 Then
        ' Waiting for module...
        W_BARGRAPH_(psBar).text = ChrW$(&H2014) & "*" & ChrW$(&H2014)  ' "—*—"
    ElseIf bestPowerNeeded <> researchNotStarted Then
        ' Show how much power is needed, before research can start.
        formatPowerText psBar, bestPowerNeeded
        setBarGraphValue psBar, W_BARGRAPH_(psBar).majorCol, researchPowerCost - bestPowerNeeded, researchPowerCost
    End If
End Sub

' Set the shadow power for the selected player
Public Sub intSetShadowPower(ByVal quantity As Long)
    ManuPower = quantity
End Sub

' the display functions and callbacks
Public Function intdisplay_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_intDisplayImage
        intDisplayImage args(0), args(1), args(2)
    Case FN_intDisplayMissionClock
        intDisplayMissionClock args(0), args(1), args(2)
    Case FN_intDisplayImageHilight
        intDisplayImageHilight args(0), args(1), args(2)
    Case FN_intDisplayButtonHilight
        intDisplayButtonHilight args(0), args(1), args(2)
    Case FN_intDisplayButtonFlash
        intDisplayButtonFlash args(0), args(1), args(2)
    Case FN_intDisplayEditBox
        intDisplayEditBox args(0), args(1), args(2)
    Case FN_WidgetAudioCallback
        WidgetAudioCallback args(0)
    Case FN_intDisplayProximityBlips
        intDisplayProximityBlips args(0), args(1), args(2)
    End Select
End Function
