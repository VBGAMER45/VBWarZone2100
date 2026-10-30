Attribute VB_Name = "radar_c"
' src/radar.cpp, radar.h: the minimap. The bitmap is an iV_Image (WzGfx);
' its pixels are built in radarPixels (a Long per pixel, the PIELIGHT's
' rgba, r in the low byte: the bytes C++ writes into radarBitmap.bmp_w())
' and copied into the image before pie_DownLoadRadar. The RadarWidget
' class's fields are a UDT (pRadarWidget) with its methods as procedures
' (RadarWidgetUDT_display, _hitTest, _clicked, _released, _highlight,
' _highlightLost, _capturesMouseDrag, _mouseDragged) and the WIDGET
' geometry it reads; the widget class RadarWidget (pRadarWidgetObj, the
' WIDGET getRadarWidget returns) forwards to them.
' The globals' C++ initialisers are radar_c_init's.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Const MAX_RADARZOOM As Long = 64
Public Const MIN_RADARZOOM As Long = 8
Public Const DEFAULT_RADARZOOM As Long = 24
Public Const RADARZOOM_STEP As Long = 4

' Different mini-map draw modes.
Public Enum RADAR_DRAW_MODE
    RADAR_MODE_TERRAIN                  ' Draw terrain map
    RADAR_MODE_DEFAULT = RADAR_MODE_TERRAIN ' Default is terrain map
    RADAR_MODE_HEIGHT_MAP               ' Draw height map
    RADAR_MODE_COMBINED
    RADAR_MODE_NO_TERRAIN               ' Only display objects
    NUM_RADAR_MODES
End Enum

' display3ddef.h
Public Const RADWIDTH As Long = 128
Public Const RADHEIGHT As Long = 128

Private Const HIT_NOTIFICATION As Long = (GAME_TICKS_PER_SEC * 2)
Private Const RADAR_FRAME_SKIP As Long = 10

Public bEnemyAllyRadarColor As Boolean  ' = false: Enemy/ally radar color.
Public radarDrawMode As RADAR_DRAW_MODE ' = RADAR_MODE_DEFAULT: Current mini-map mode.
Public rotateRadar As Boolean           ' = true: Rotate the radar?
Public radarRotationArrow As Boolean    ' = true: display arrow when radar rotation enabled?

Private colRadarAlly As Long, colRadarMe As Long, colRadarEnemy As Long ' PIELIGHT
Private tileColours(0 To MAX_TILES - 1) As Long ' PIELIGHT
Private radarBitmap As iV_Image
Private radarPixels() As Long           ' (VB) radarBitmap's pixels (see above)
Private radarOverlayBuffer() As Long    ' UDWORD *: = nullptr
Private radarOverlayAllocated As Boolean
Private playerpos_ As Vector3i          ' = {0, 0, 0} (VB: the radar's own playerpos, not display3d's playerPos)

' (VB) class RadarWidget : public WIDGET (see above)
Private Type RadarWidget_
    m_clickPosition(0 To 2) As Vector2i ' std::array<optional<Vector2i>, 3>
    has_clickPosition(0 To 2) As Boolean
    m_hasHighlight As Boolean           ' = false
    m_handlingDrag As Boolean           ' = false
    x As Long                           ' (VB) the WIDGET's geometry (setGeometry)
    y As Long
    width As Long
    height As Long
End Type

Private pRadarWidget As RadarWidget_    ' static std::shared_ptr<RadarWidget> pRadarWidget;
Private hasRadarWidget As Boolean       ' (VB) pRadarWidget != nullptr
Private pRadarWidgetObj As WIDGET       ' (VB) pRadarWidget's WIDGET (Nothing = nullptr)

Public clanColours(0 To 15) As Long     ' PIELIGHT: see frontend2.png for team color order.
Private flashColours(0 To 15) As Long   ' PIELIGHT

Private radarWidth As Long, radarHeight As Long, radarTexWidth As Long, radarTexHeight As Long ' size_t
Private radarCenterX As Long, radarCenterY As Long
Private RadarZoom As Long               ' uint8_t
Private RadarZoomMultiplier As Single   ' = 1.0f
Private radarBufferSize As Long         ' size_t = 0
Private frameSkip As Long               ' = 0
Private lastBlink As Long               ' UDWORD = 0
Private Const BLINK_INTERVAL As Long = GAME_TICKS_PER_SEC \ 1
Private Const BLINK_HALF_INTERVAL As Long = BLINK_INTERVAL \ 2
Private Const OVERLAY_OPACITY As Single = 0.5!


' How far the mouse has to move to start a drag
Private Const DRAG_THRESHOLD As Long = 5

' (VB) the C++ initialisers of radar.cpp's globals
Public Sub radar_c_init()
    Dim i As Long
    bEnemyAllyRadarColor = False
    radarDrawMode = RADAR_MODE_DEFAULT
    rotateRadar = True
    radarRotationArrow = True
    RadarZoomMultiplier = 1!
    clanColours(0) = pal_RGBA(0, 255, 0, 255)       ' green  Player 0
    clanColours(1) = pal_RGBA(255, 192, 40, 255)    ' orange Player 1
    clanColours(2) = pal_RGBA(255, 255, 255, 255)   ' grey   Player 2
    clanColours(3) = pal_RGBA(0, 0, 0, 255)         ' black  Player 3
    clanColours(4) = pal_RGBA(255, 0, 0, 255)       ' red    Player 4
    clanColours(5) = pal_RGBA(20, 20, 255, 255)     ' blue   Player 5
    clanColours(6) = pal_RGBA(255, 0, 192, 255)     ' pink   Player 6
    clanColours(7) = pal_RGBA(0, 255, 255, 255)     ' cyan   Player 7
    clanColours(8) = pal_RGBA(255, 255, 0, 255)     ' yellow Player 8
    clanColours(9) = pal_RGBA(144, 0, 255, 255)     ' purple Player 9
    clanColours(10) = pal_RGBA(200, 255, 255, 255)  ' white  Player A (Should be brighter than grey, but grey is already maximum.)
    clanColours(11) = pal_RGBA(128, 128, 255, 255)  ' bright blue Player B
    clanColours(12) = pal_RGBA(128, 255, 128, 255)  ' neon green  Player C
    clanColours(13) = pal_RGBA(128, 0, 0, 255)      ' infrared    Player D
    clanColours(14) = pal_RGBA(64, 0, 128, 255)     ' ultraviolet Player E
    clanColours(15) = pal_RGBA(128, 128, 0, 255)    ' brown       Player F
    'right now the flash color is all bright red
    For i = 0 To 15
        flashColours(i) = pal_RGBA(254, 37, 37, 200)
    Next
    Set radarBitmap = New iV_Image
End Sub

' (VB) a PIELIGHT's bytes
Private Function PL_r(ByVal c As Long) As Long
    PL_r = c And &HFF&
End Function

Private Function PL_g(ByVal c As Long) As Long
    PL_g = SHRU(c, 8) And &HFF&
End Function

Private Function PL_b(ByVal c As Long) As Long
    PL_b = SHRU(c, 16) And &HFF&
End Function

Private Function PL_a(ByVal c As Long) As Long
    PL_a = SHRU(c, 24) And &HFF&
End Function

' taken from https://en.wikipedia.org/wiki/Alpha_compositing
Private Function mix(ByVal over As Long, ByVal base As Long) As Long
    Dim overAlpha As Single, overRed As Single, overGreen As Single, overBlue As Single
    Dim baseAlpha As Single, baseRed As Single, baseGreen As Single, baseBlue As Single
    Dim newAlpha As Single, newRed As Single, newGreen As Single, newBlue As Single
    overAlpha = PL_a(over) / 255!
    overRed = PL_r(over) / 255!
    overGreen = PL_g(over) / 255!
    overBlue = PL_b(over) / 255!

    baseAlpha = PL_a(base) / 255!
    baseRed = PL_r(base) / 255!
    baseGreen = PL_g(base) / 255!
    baseBlue = PL_b(base) / 255!

    newAlpha = overAlpha + baseAlpha * (1 - overAlpha)
    newRed = (overRed * overAlpha + baseRed * baseAlpha * (1 - overAlpha)) / newAlpha
    newGreen = (overGreen * overAlpha + baseGreen * baseAlpha * (1 - overAlpha)) / newAlpha
    newBlue = (overBlue * overAlpha + baseBlue * baseAlpha * (1 - overAlpha)) / newAlpha

    mix = pal_RGBA(clip255(newRed * 255!), clip255(newGreen * 255!), clip255(newBlue * 255!), clip255(newAlpha * 255!))
End Function

' (VB) static_cast<UBYTE>(clip<int>(static_cast<int>(v), 0, 255))
Private Function clip255(ByVal v As Single) As Long
    Dim i As Long
    i = Fix(v)
    If i < 0 Then i = 0
    If i > 255 Then i = 255
    clip255 = i
End Function

Private Function applyAlpha(ByVal color As Long, ByVal alpha As Single) As Long
    applyAlpha = (color And &HFFFFFF) Or SHL(Fix(PL_a(color) * alpha) And &HFF&, 24)
End Function

Private Sub radarSize(ByVal ZoomLevel As Long)
    Dim zoom As Single, radarDiag As Long, radarX0 As Long, radarY0 As Long, m As Long
    zoom = CSng(ZoomLevel) * RadarZoomMultiplier / 16!
    radarWidth = Fix(radarTexWidth * zoom)
    radarHeight = Fix(radarTexHeight * zoom)
    If rotateRadar Then
        If radarHeight > radarWidth Then m = radarHeight Else m = radarWidth
        radarCenterX = pie_GetVideoBufferWidth() - BASE_GAP * 4 - m \ 2
        radarCenterY = pie_GetVideoBufferHeight() - BASE_GAP * 4 - m \ 2
    Else
        radarCenterX = pie_GetVideoBufferWidth() - BASE_GAP * 4 - radarWidth \ 2
        radarCenterY = pie_GetVideoBufferHeight() - BASE_GAP * 4 - radarHeight \ 2
    End If
    debug_ 223, LOG_WZ, "radarSize", "radar=(%u,%u) tex=(%zu,%zu) size=(%zu,%zu)", radarCenterX, radarCenterY, radarTexWidth, radarTexHeight, radarWidth, radarHeight

    If hasRadarWidget Then
        radarDiag = -Int(-Sqr(CDbl((radarWidth + 1) * (radarWidth + 1)) + CDbl((radarHeight + 1) * (radarHeight + 1))))
        radarX0 = radarCenterX - (radarDiag \ 2)
        radarY0 = radarCenterY - (radarDiag \ 2)
        pRadarWidget.x = radarX0         ' pRadarWidget->setGeometry(radarX0, radarY0, radarDiag, radarDiag)
        pRadarWidget.y = radarY0
        pRadarWidget.width = radarDiag
        pRadarWidget.height = radarDiag
        pRadarWidgetObj.setGeometry radarX0, radarY0, radarDiag, radarDiag
    End If
End Sub

Public Sub radarInitVars()
    radarTexWidth = 0
    radarTexHeight = 0
    RadarZoom = war_GetRadarZoom() And &HFF&
    debug_ 239, LOG_WZ, "radarInitVars", "Resetting radar zoom to %u", RadarZoom
    radarSize RadarZoom
    playerpos_.x = -1: playerpos_.y = -1: playerpos_.z = -1
    frameSkip = 0
End Sub

Public Function InitRadar() As Boolean
    ' Ally/enemy/me colors
    colRadarAlly = psPalette(WZCOL_YELLOW)
    colRadarEnemy = psPalette(WZCOL_RED)
    colRadarMe = psPalette(WZCOL_WHITE)
    If Not hasRadarWidget Then
        Dim fresh As RadarWidget_
        pRadarWidget = fresh            ' std::make_shared<RadarWidget>()
        Set pRadarWidgetObj = RadarWidget_new()
        hasRadarWidget = True
    End If
    InitRadar = True
End Function

Public Function resizeRadar() As Boolean
    Dim m As Long
    radarBitmap.clear
    If radarOverlayAllocated Then
        Erase radarOverlayBuffer
        radarOverlayAllocated = False
    End If
    radarTexWidth = Abs(scrollMaxX - scrollMinX)
    radarTexHeight = Abs(scrollMaxY - scrollMinY)
    radarBufferSize = radarTexWidth * radarTexHeight * 4    ' sizeof(UDWORD)
    radarBitmap.allocate radarTexWidth, radarTexHeight, 4, True
    If radarTexWidth * radarTexHeight > 0 Then
        ReDim radarOverlayBuffer(0 To radarTexWidth * radarTexHeight - 1)  ' malloc, memset 0
        ReDim radarPixels(0 To radarTexWidth * radarTexHeight - 1)
        radarOverlayAllocated = True
    End If
    frameSkip = 0
    If rotateRadar Then
        m = radarTexWidth
        If radarTexHeight > m Then m = radarTexHeight
        If m < 1 Then m = 1
        RadarZoomMultiplier = CSng(IIf(RADWIDTH > RADHEIGHT, RADWIDTH, RADHEIGHT)) / m
    Else
        RadarZoomMultiplier = 1!
    End If
    debug_ 280, LOG_WZ, "resizeRadar", "Setting radar zoom to %u", RadarZoom
    radarSize RadarZoom
    pie_SetRadar -CSng(radarWidth) / 2! - 1, -CSng(radarHeight) / 2! - 1, CSng(radarWidth), CSng(radarHeight), radarTexWidth, radarTexHeight

    resizeRadar = True
End Function

Public Function ShutdownRadar() As Boolean
    radarBitmap.clear
    If radarOverlayAllocated Then
        Erase radarOverlayBuffer
        Erase radarPixels
        radarOverlayAllocated = False
    End If
    frameSkip = 0
    If hasRadarWidget Then
        ' psWScreen->psForm->detach(pRadarWidget): (VB) from its parent, the screen's form (none when the
        ' interface never attached it)
        If Not pRadarWidgetObj.parent Is Nothing Then pRadarWidgetObj.parent.detach pRadarWidgetObj
        Set pRadarWidgetObj = Nothing
        hasRadarWidget = False          ' pRadarWidget.reset()
    End If
    ShutdownRadar = True
End Function

Public Sub SetRadarZoom(ByVal ZoomLevel As Long)
    ZoomLevel = ZoomLevel And &HFF&
    If ZoomLevel > MAX_RADARZOOM Then
        ZoomLevel = MAX_RADARZOOM
    End If
    If ZoomLevel < MIN_RADARZOOM Then
        ZoomLevel = MIN_RADARZOOM
    End If
    debug_ 311, LOG_WZ, "SetRadarZoom", "Setting radar zoom to %u from %u", ZoomLevel, RadarZoom
    RadarZoom = ZoomLevel
    radarSize RadarZoom
    frameSkip = 0
    resizeRadar
End Sub

Public Function GetRadarZoom() As Long
    GetRadarZoom = RadarZoom
End Function

' Calculate the radar pixel sizes. Returns pixels per tile.
Private Sub CalcRadarPixelSize(SizeH As Single, SizeV As Single)
    Dim h As Long, w As Long
    h = radarTexHeight: If h < 1 Then h = 1
    w = radarTexWidth: If w < 1 Then w = 1
    SizeH = CSng(radarHeight) / h
    SizeV = CSng(radarWidth) / w
End Sub

' (VB) vector.h's Vector2f_Rotate2f(v, angle): v rotated by the 16-bit angle
Private Function Vector2f_Rotate2f(v As Vector2f, ByVal angle As Long) As Vector2f
    Dim c As Single, s As Single
    c = iCos(angle) / 65536!
    s = iSin(angle) / 65536!
    Vector2f_Rotate2f.x = v.x * c - v.y * s
    Vector2f_Rotate2f.y = v.x * s + v.y * c
End Function

' Given a position within the radar, return a world coordinate.
Public Sub CalcRadarPosition(ByVal mX As Long, ByVal mY As Long, PosX As Long, PosY As Long)
    Dim sPosX As Long, sPosY As Long, pixSizeH As Single, pixSizeV As Single, pos As Vector2f

    pos.x = mX - radarCenterX
    pos.y = mY - radarCenterY
    If rotateRadar Then
        pos = Vector2f_Rotate2f(pos, -playerPos.r.y)
    End If
    pos.x = pos.x + radarWidth / 2!
    pos.y = pos.y + radarHeight / 2!

    CalcRadarPixelSize pixSizeH, pixSizeV
    sPosX = Fix(pos.x / pixSizeH)   ' adjust for pixel size
    sPosY = Fix(pos.y / pixSizeV)
    sPosX = sPosX + scrollMinX      ' adjust for scroll limits
    sPosY = sPosY + scrollMinY

    ' old safety code -- still necessary?
    If sPosX < scrollMinX Then sPosX = scrollMinX
    If sPosX > scrollMaxX - 1 Then sPosX = scrollMaxX - 1
    If sPosY < scrollMinY Then sPosY = scrollMinY
    If sPosY > scrollMaxY - 1 Then sPosY = scrollMaxY - 1

    PosX = sPosX
    PosY = sPosY
End Sub

Public Sub drawRadar()
    Dim pixSizeH As Single, pixSizeV As Single, radarMatrix As mat4, orthoMatrix As mat4, m As mat4, ext As mat4
    CalcRadarPixelSize pixSizeH, pixSizeV

    If Not (radarBitmap.bmp_w() <> 0) Then ASSERT_FAILURE "radarBitmap.bmp_w()", "src/radar.cpp:371", "drawRadar", "No radar buffer allocated": Exit Sub
    If Not radarOverlayAllocated Then ASSERT_FAILURE "radarOverlayBuffer", "src/radar.cpp:371", "drawRadar", "No radar buffer allocated": Exit Sub

    setViewingWindow
    playerpos_ = playerPos.p ' cache position

    If frameSkip <= 0 Then
        DrawRadarTiles
        DrawRadarObjects
        applyMinimapOverlay
        CopyMemory ByVal radarBitmap.bmp_w(), radarPixels(0), radarTexWidth * radarTexHeight * 4
        pie_DownLoadRadar radarBitmap
        frameSkip = RADAR_FRAME_SKIP
    End If
    frameSkip = frameSkip - 1
    radarMatrix = glm_translate(glm_vec3(radarCenterX, radarCenterY, 0))
    orthoMatrix = glm_ortho(0!, CSng(pie_GetVideoBufferWidth()), CSng(pie_GetVideoBufferHeight()), 0!)
    If rotateRadar Then
        ' rotate the map
        radarMatrix = glm_mat4_mul(radarMatrix, glm_rotate(UNDEG(playerPos.r.y), glm_vec3(0!, 0!, 1!)))
        If radarRotationArrow Then
            m = glm_mat4_mul(orthoMatrix, radarMatrix)
            DrawNorth m
        End If
    End If

    m = glm_mat4_mul(orthoMatrix, radarMatrix)
    pie_RenderRadar VarPtr(m)
    ext = glm_mat4_mul(m, glm_translate(glm_vec3(-CSng(radarWidth) / 2! - 1!, -CSng(radarHeight) / 2! - 1!, 0!)))
    DrawRadarExtras ext
    drawRadarBlips Fix(-radarWidth / 2! - 1), Fix(-radarHeight / 2! - 1), pixSizeH, pixSizeV, m
End Sub

Private Sub DrawNorth(modelViewProjectionMatrix As mat4)
    iV_DrawImage IntImages, RADAR_NORTH, Fix(-((radarWidth / 2!) + iV_GetImageWidth(IntImages, RADAR_NORTH) + 1)), Fix(-(radarHeight / 2!)), VarPtr(modelViewProjectionMatrix)
End Sub

' (VB) col.byte.x = static_cast<uint8_t>(sqrtf(col.byte.x * v)) for each of r, g, b
Private Function radarScaled(ByVal col As Long, ByVal v As Single) As Long
    radarScaled = pal_RGBA(Fix(Sqr(PL_r(col) * v)) And &HFF&, Fix(Sqr(PL_g(col) * v)) And &HFF&, Fix(Sqr(PL_b(col) * v)) And &HFF&, PL_a(col))
End Function

' (VB) col.byte.r/g/b = col.byte.x * num / den (integer)
Private Function radarDimmed(ByVal col As Long, ByVal num As Long, ByVal den As Long) As Long
    radarDimmed = pal_RGBA((PL_r(col) * num) \ den, (PL_g(col) * num) \ den, (PL_b(col) * num) \ den, PL_a(col))
End Function

Private Function appliedRadarColour(ByVal drawMode As RADAR_DRAW_MODE, ByVal WTile As Long) As Long
    Dim WScr As Long, col As Long, h As Long
    WScr = psPalette(WZCOL_BLACK)   ' squelch warning

    ' draw radar on/off feature
    If Not getRevealStatus() And Not TEST_TILE_VISIBLE_TO_SELECTEDPLAYER(WTile) Then
        appliedRadarColour = psPalette(WZCOL_TRANSPARENT_BOX)
        Exit Function
    End If

    Select Case drawMode
    Case RADAR_MODE_TERRAIN
        ' draw radar terrain on/off feature
        col = tileColours(TileNumber_tile(psMapTiles(WTile).texture))

        col = radarScaled(col, psMapTiles(WTile).illumination)
        If terrainType(WTile) = TER_CLIFFFACE Then
            col = radarDimmed(col, 1, 2)
        End If
        If Not hasSensorOnTile(WTile, selectedPlayer) Then
            col = radarDimmed(col, 2, 3)
        End If
        If Not TEST_TILE_VISIBLE_TO_SELECTEDPLAYER(WTile) Then
            col = radarDimmed(col, 1, 2)
        End If
        WScr = col
    Case RADAR_MODE_COMBINED
        ' draw radar terrain on/off feature
        col = tileColours(TileNumber_tile(psMapTiles(WTile).texture))

        col = radarScaled(col, (psMapTiles(WTile).illumination + psMapTiles(WTile).height \ ELEVATION_SCALE) \ 2)
        If terrainType(WTile) = TER_CLIFFFACE Then
            col = radarDimmed(col, 1, 2)
        End If
        If Not hasSensorOnTile(WTile, selectedPlayer) Then
            col = radarDimmed(col, 2, 3)
        End If
        If Not TEST_TILE_VISIBLE_TO_SELECTEDPLAYER(WTile) Then
            col = radarDimmed(col, 1, 2)
        End If
        WScr = col
    Case RADAR_MODE_HEIGHT_MAP
        h = (psMapTiles(WTile).height \ ELEVATION_SCALE) And &HFF&
        WScr = pal_RGBA(h, h, h, PL_a(WScr))
    Case RADAR_MODE_NO_TERRAIN
        WScr = psPalette(WZCOL_RADAR_BACKGROUND)
    Case NUM_RADAR_MODES
        ' assert(false); (NDEBUG in a release build)
    End Select
    appliedRadarColour = WScr
End Function

' Draw the map tiles on the radar.
Private Sub DrawRadarTiles()
    Dim x As Long, y As Long, psTile As Long, pixelStartPos As Long, radarBufferSize2 As Long
    radarBufferSize2 = radarBitmap.size_in_bytes()
    For x = scrollMinX To scrollMaxX - 1
        For y = scrollMinY To scrollMaxY - 1
            psTile = mapTile(x, y)
            pixelStartPos = (radarTexWidth * (y - scrollMinY) + (x - scrollMinX)) ' (VB) in pixels: * 4 in bytes

            If Not (pixelStartPos * 4 < radarBufferSize2) Then ASSERT_FAILURE "pixelStartPos < radarBufferSize2", "src/radar.cpp:511", "DrawRadarTiles", "Buffer overrun"

            If y = scrollMinY Or x = scrollMinX Or y = scrollMaxY - 1 Or x = scrollMaxX - 1 Then
                radarPixels(pixelStartPos) = psPalette(WZCOL_BLACK)
            Else
                radarPixels(pixelStartPos) = appliedRadarColour(radarDrawMode, psTile)
            End If
        Next
    Next
End Sub

' (VB) psObj->visibleForLocalDisplay() || (bMultiPlayer && alliancesSharedVision(game.alliance)
'      && selectedPlayer < MAX_PLAYERS && aiCheckAlliances(selectedPlayer, psObj->player))
Private Function radarShows(ByVal psObj As Long) As Boolean
    If BASE_OBJECT_visibleForLocalDisplay(psObj) <> 0 Then
        radarShows = True
    ElseIf bMultiPlayer Then
        If alliancesSharedVision(game.alliance) And selectedPlayer < MAX_PLAYERS Then
            radarShows = aiCheckAlliances(selectedPlayer, SIMPLE_OBJECT_player(psObj))
        End If
    End If
End Function

' Draw the droids and structure positions on the radar.
Private Sub DrawRadarObjects()
    Dim clan As Long, playerCol As Long, flashCol As Long, blinkState As Boolean, it As Long, psDroid As Long
    Dim x As Long, y As Long, pos As Long, psTile As Long, psStruct As Long, i As Long
    For i = 0 To radarTexWidth * radarTexHeight - 1
        radarOverlayBuffer(i) = 0       ' memset(radarOverlayBuffer, 0, radarBufferSize)
    Next
    blinkState = (Fix(U32(gameTime - lastBlink) / BLINK_HALF_INTERVAL) <> 0)

    ' Show droids on map - go through all players
    For clan = 0 To MAX_PLAYERS - 1
        'see if have to draw enemy/ally color
        If bEnemyAllyRadarColor Then
            If clan = selectedPlayer Then
                playerCol = colRadarMe
            ElseIf selectedPlayer < MAX_PLAYERS Then
                If aiCheckAlliances(selectedPlayer, clan) Then playerCol = colRadarAlly Else playerCol = colRadarEnemy
            Else
                playerCol = colRadarEnemy
            End If
        Else
            'original 8-color mode
            playerCol = clanColours(getPlayerColour(clan))
        End If

        flashCol = flashColours(getPlayerColour(clan))

        ' Go through all droids
        it = apsDroidLists(clan).head
        Do While it <> 0
            psDroid = lnObj(it)
            If aDroid(psDroid).pos.x < world_coord(scrollMinX) Or aDroid(psDroid).pos.y < world_coord(scrollMinY) _
               Or aDroid(psDroid).pos.x >= world_coord(scrollMaxX) Or aDroid(psDroid).pos.y >= world_coord(scrollMaxY) Then
                GoTo continue_droid
            End If
            If radarShows(psDroid) Then
                x = aDroid(psDroid).pos.x \ TILE_UNITS
                y = aDroid(psDroid).pos.y \ TILE_UNITS
                pos = (x - scrollMinX) + (y - scrollMinY) * radarTexWidth

                If Not (pos * 4 < radarBufferSize) Then ASSERT_FAILURE "pos * sizeof(*radarOverlayBuffer) < radarBufferSize", "src/radar.cpp:579", "DrawRadarObjects", "Buffer overrun"
                If clan = selectedPlayer And U32(gameTime) > HIT_NOTIFICATION And U32(gameTime - aDroid(psDroid).timeLastHit) < HIT_NOTIFICATION Then
                    If aDroid(psDroid).selected And Not blinkState Then
                        radarOverlayBuffer(pos) = applyAlpha(flashCol, OVERLAY_OPACITY)
                    Else
                        radarOverlayBuffer(pos) = flashCol
                    End If
                Else
                    If aDroid(psDroid).selected And Not blinkState Then
                        radarOverlayBuffer(pos) = applyAlpha(playerCol, OVERLAY_OPACITY)
                    Else
                        radarOverlayBuffer(pos) = playerCol
                    End If
                End If
            End If
continue_droid:
            it = lnNext(it)
        Loop
    Next

    ' Do the same for structures
    For x = scrollMinX To scrollMaxX - 1
        For y = scrollMinY To scrollMaxY - 1
            psTile = mapTile(x, y)
            pos = (x - scrollMinX) + (y - scrollMinY) * radarTexWidth

            If Not (pos * 4 < radarBufferSize) Then ASSERT_FAILURE "pos * sizeof(*radarOverlayBuffer) < radarBufferSize", "src/radar.cpp:607", "DrawRadarObjects", "Buffer overrun"
            If Not TileHasStructure(psTile) Then
                GoTo continue_tile
            End If
            psStruct = psMapTiles(psTile).psObject
            clan = aStruct(psStruct).player

            'see if have to draw enemy/ally color
            If bEnemyAllyRadarColor Then
                If clan = selectedPlayer Then
                    playerCol = colRadarMe
                ElseIf selectedPlayer < MAX_PLAYERS Then
                    If aiCheckAlliances(selectedPlayer, clan) Then playerCol = colRadarAlly Else playerCol = colRadarEnemy
                Else
                    playerCol = colRadarEnemy
                End If
            Else
                'original 8-color mode
                playerCol = clanColours(getPlayerColour(clan))
            End If
            flashCol = flashColours(getPlayerColour(clan))

            If radarShows(psStruct) Then
                If clan = selectedPlayer And U32(gameTime) > HIT_NOTIFICATION And U32(gameTime - aStruct(psStruct).timeLastHit) < HIT_NOTIFICATION Then
                    If aStruct(psStruct).player = selectedPlayer And aStruct(psStruct).selected And Not blinkState Then
                        radarOverlayBuffer(pos) = applyAlpha(flashCol, OVERLAY_OPACITY)
                    Else
                        radarOverlayBuffer(pos) = flashCol
                    End If
                Else
                    If aStruct(psStruct).player = selectedPlayer And aStruct(psStruct).selected And Not blinkState Then
                        radarOverlayBuffer(pos) = applyAlpha(playerCol, OVERLAY_OPACITY)
                    Else
                        radarOverlayBuffer(pos) = playerCol
                    End If
                End If
            End If
continue_tile:
        Next
    Next
    If U32(gameTime - lastBlink) >= BLINK_INTERVAL Then
        lastBlink = gameTime
    End If
End Sub

Private Sub applyMinimapOverlay()
    Dim radarTexCount As Long, i As Long
    radarTexCount = radarTexWidth * radarTexHeight
    If Not (radarTexCount * radarBitmap.channels() <= radarBitmap.size_in_bytes()) Then ASSERT_FAILURE "radarTexCount * static_cast<size_t>(radarBitmap.channels()) <= radarBufferSize2", "src/radar.cpp:664", "applyMinimapOverlay", "Buffer overrun"
    If Not (radarTexCount * radarBitmap.channels() <= radarBufferSize) Then ASSERT_FAILURE "radarTexCount * static_cast<size_t>(radarBitmap.channels()) <= radarBufferSize", "src/radar.cpp:665", "applyMinimapOverlay", "Buffer overrun"
    For i = 0 To radarTexCount - 1
        If radarOverlayBuffer(i) = 0 Then
            GoTo continue_
        End If
        ' baseColor = pal_RGBA(the pixel's bytes), overColor = PLfromUDWORD(radarOverlayBuffer[i]): the same Longs here
        radarPixels(i) = mix(radarOverlayBuffer(i), radarPixels(i))
continue_:
    Next
End Sub

' Rotate an array of 2d vectors about a given angle, also translates them after rotating.
Private Sub RotateVector2D(Vector() As Vector3i, TVector() As Vector3i, Pos As Vector3i, ByVal Angle As Long, ByVal count As Long)
    Dim Cos_ As Double, Sin_ As Double, ox As Long, oy As Long, i As Long
    Cos_ = iCos(Angle)
    Sin_ = iSin(Angle)

    ox = Pos.x
    oy = Pos.y

    For i = 0 To count - 1
        ' (int64_t products) >> 16: floor division by 65536
        TVector(i).x = Int((Vector(i).x * Cos_ + Vector(i).y * Sin_) / 65536#) + ox
        TVector(i).y = Int((Vector(i).y * Cos_ - Vector(i).x * Sin_) / 65536#) + oy
    Next
End Sub

Private Function getDistanceAdjust() As Long
    Dim dif As Long
    dif = Fix(MAXDISTANCE - getViewDistance())
    If dif < 0 Then dif = 0

    getDistanceAdjust = dif \ 100
End Function

Private Function getLengthAdjust() As Long
    Dim pitch As Long, lookingDown As Long, lookingFar As Long, dif As Long
    pitch = Fix(360 - (playerPos.r.x / DEG_1))

    ' Max at
    lookingDown = (0 - MIN_PLAYER_X_ANGLE)
    lookingFar = (0 - MAX_PLAYER_X_ANGLE)

    dif = pitch - lookingFar
    If dif < 0 Then dif = 0
    If dif > (lookingDown - lookingFar) Then
        dif = (lookingDown - lookingFar)
    End If

    getLengthAdjust = dif \ 2
End Function

' Draws a Myth/FF7 style viewing window
Private Sub setViewingWindow()
    Dim pixSizeH As Single, pixSizeV As Single, v(0 To 3) As Vector3i, tv(0 To 3) As Vector3i, centre As Vector3i
    Dim shortX As Long, longX As Long, yDrop As Long, yDropVar As Long, dif As Long, dif2 As Long, colour As Long
    Dim x As Long, y As Long
    dif = getDistanceAdjust()
    dif2 = getLengthAdjust()
    CalcRadarPixelSize pixSizeH, pixSizeV
    x = Fix(playerPos.p.x * pixSizeH / TILE_UNITS)
    y = Fix(playerPos.p.z * pixSizeV / TILE_UNITS)

    shortX = Fix(((visibleTiles.x \ 4) - (dif \ 6)) * pixSizeH)
    longX = Fix(((visibleTiles.x \ 2) - (dif \ 4)) * pixSizeH)
    yDropVar = Fix(((visibleTiles.y \ 2) - (dif2 \ 3)) * pixSizeV)
    yDrop = Fix(((visibleTiles.y \ 2) - dif2 \ 3) * pixSizeV)

    v(0).x = longX
    v(0).y = -yDropVar

    v(1).x = -longX
    v(1).y = -yDropVar

    v(2).x = shortX
    v(2).y = yDrop

    v(3).x = -shortX
    v(3).y = yDrop

    centre.x = Fix(x - scrollMinX * pixSizeH)
    centre.y = Fix(y - scrollMinY * pixSizeV)

    RotateVector2D v, tv, centre, playerPos.r.y, 4

    Select Case getCampaignNumber()
    Case 1, 2
        ' white
        colour = pal_RGBA(255, 255, 255, &H3F)
    Case 3
        ' greenish
        colour = pal_RGBA(&H3F, 255, &H3F, &H3F)
    Case Else
        ' black
        colour = pal_RGBA(0, 0, 0, &H3F)
    End Select

    ' Send the four points to the draw routine and the clip box params
    pie_SetViewingWindow VarPtr(tv(0)), colour
End Sub

Private Sub DrawRadarExtras(modelViewProjectionMatrix As mat4)
    pie_DrawViewingWindow VarPtr(modelViewProjectionMatrix)
    RenderWindowFrame FRAME_RADAR, -1, -1, radarWidth + 2, radarHeight + 2, VarPtr(modelViewProjectionMatrix)
End Sub

' Does a screen coordinate lie within the radar area?
Public Function CoordInRadar(ByVal x As Long, ByVal y As Long) As Boolean
    Dim pos As Vector2f
    pos.x = x - radarCenterX
    pos.y = y - radarCenterY
    If rotateRadar Then
        pos = Vector2f_Rotate2f(pos, -playerPos.r.y)
    End If
    pos.x = pos.x + radarWidth / 2!
    pos.y = pos.y + radarHeight / 2!

    If pos.x < 0 Or pos.y < 0 Or pos.x >= radarWidth Or pos.y >= radarHeight Then
        CoordInRadar = False
        Exit Function
    End If
    CoordInRadar = True
End Function

Public Sub radarColour(ByVal tileNumber As Long, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    tileColours(tileNumber) = pal_RGBA(r, g, b, 255)
End Sub

' MARK: - RadarWidget

' std::shared_ptr<WIDGET> getRadarWidget() (Nothing = nullptr)
Public Function getRadarWidget() As WIDGET
    Set getRadarWidget = pRadarWidgetObj
End Function

' #define WKEY_ORDER (getRightClickOrders()?WKEY_SECONDARY:WKEY_PRIMARY)
Private Function WKEY_ORDER() As Long
    If getRightClickOrders() Then WKEY_ORDER = WKEY_SECONDARY Else WKEY_ORDER = WKEY_PRIMARY
End Function

' #define WKEY_SELECT (getRightClickOrders()?WKEY_PRIMARY:WKEY_SECONDARY)
Private Function WKEY_SELECT() As Long
    If getRightClickOrders() Then WKEY_SELECT = WKEY_PRIMARY Else WKEY_SELECT = WKEY_SECONDARY
End Function

Public Function RadarWidgetUDT_capturesMouseDrag(ByVal wkey As Long) As Boolean
    RadarWidgetUDT_capturesMouseDrag = (wkey = WKEY_SELECT())
End Function

' psStartContext: its mx, my; psContext: its mx, my and the screen context's (convertToScreenContext) mx, my
Public Sub RadarWidgetUDT_mouseDragged(ByVal wkey As Long, ByVal startMx As Long, ByVal startMy As Long, ByVal mx As Long, ByVal my As Long, ByVal screenMx As Long, ByVal screenMy As Long)
    Dim x As Long, y As Long, PosX As Long, PosY As Long
    If wkey <> WKEY_SELECT() Then
        Exit Sub
    End If

    If Not getRotActive() Then
        If Not pRadarWidget.m_handlingDrag Then
            If Abs(startMx - mx) > DRAG_THRESHOLD Or _
               Abs(startMy - my) > DRAG_THRESHOLD Then
                ' start processing as drag
                pRadarWidget.m_handlingDrag = True
            End If
        End If

        If pRadarWidget.m_handlingDrag Then
            x = screenMx
            y = screenMy
            CalcRadarPosition x, y, PosX, PosY
            setViewPos PosX, PosY, True
            If ctrlShiftDown() Then
                playerPos.r.y = 0
            End If
        End If
    End If
End Sub

Public Sub RadarWidgetUDT_display(ByVal xOffset As Long, ByVal yOffset As Long)
    If Not radarVisible() Then Exit Sub

    gfx_api_context_debugStringMarker "Draw 3D scene - radar"
    drawRadar
End Sub

Public Function RadarWidgetUDT_hitTest(ByVal x As Long, ByVal y As Long) As Boolean
    ' NOTE: CoordInRadar expects x, y in *screen* coordinates - this is _currently_ the case as a byproduct of how the radar widget is added to the UI screen
    ' (VB) WIDGET::hitTest: inside the geometry
    If x >= pRadarWidget.x And y >= pRadarWidget.y And x < pRadarWidget.x + pRadarWidget.width And y < pRadarWidget.y + pRadarWidget.height Then
        RadarWidgetUDT_hitTest = CoordInRadar(x, y)
    End If
End Function

Public Sub RadarWidgetUDT_clicked(ByVal xOffset As Long, ByVal yOffset As Long, ByVal mx As Long, ByVal my As Long, ByVal key As Long)
    Dim x As Long, y As Long
    x = xOffset + mx
    y = yOffset + my

    pRadarWidget.m_clickPosition(key).x = x
    pRadarWidget.m_clickPosition(key).y = y
    pRadarWidget.has_clickPosition(key) = True
End Sub

Public Sub RadarWidgetUDT_highlight()
    pRadarWidget.m_hasHighlight = True
End Sub

Public Sub RadarWidgetUDT_highlightLost()
    Dim k As Long
    pRadarWidget.m_hasHighlight = False
    For k = 0 To 2
        pRadarWidget.has_clickPosition(k) = False
    Next
End Sub

Public Sub RadarWidgetUDT_released(ByVal xOffset As Long, ByVal yOffset As Long, ByVal mx As Long, ByVal my As Long, ByVal key As Long)
    Dim PosX As Long, PosY As Long, x As Long, y As Long
    x = xOffset + mx
    y = yOffset + my

    If key = WKEY_ORDER() Then
        If pRadarWidget.has_clickPosition(WKEY_ORDER()) Then
            x = pRadarWidget.m_clickPosition(WKEY_ORDER()).x
            y = pRadarWidget.m_clickPosition(WKEY_ORDER()).y

            ' If we're tracking a droid, then cancel that
            CalcRadarPosition x, y, PosX, PosY
            If selectedPlayer < MAX_PLAYERS Then
                ' MARKER
                ' Send all droids to that location
                orderSelectedLoc selectedPlayer, (PosX * TILE_UNITS) + TILE_UNITS \ 2, _
                                 (PosY * TILE_UNITS) + TILE_UNITS \ 2, ctrlShiftDown() ' ctrlShiftDown() = ctrl clicked a destination, add an order
            End If
            CheckScrollLimits
            audio_PlayTrack ID_SOUND_MESSAGEEND
        End If
    End If

    If key = WKEY_SELECT() And Not pRadarWidget.m_handlingDrag Then
        If pRadarWidget.has_clickPosition(WKEY_SELECT()) Then
            x = pRadarWidget.m_clickPosition(WKEY_SELECT()).x
            y = pRadarWidget.m_clickPosition(WKEY_SELECT()).y

            CalcRadarPosition x, y, PosX, PosY

            If war_GetRadarJump() Then
                ' Go instantly
                setViewPos PosX, PosY, True
            Else
                ' Pan to it
                requestRadarTrack PosX * TILE_UNITS, PosY * TILE_UNITS
            End If
        End If
    End If

    pRadarWidget.has_clickPosition(key) = False
    pRadarWidget.m_handlingDrag = False
End Sub

Public Function isMouseOverRadar() As Boolean
    If Not hasRadarWidget Then Exit Function
    isMouseOverRadar = pRadarWidget.m_hasHighlight
End Function

Public Function isRadarDragging() As Boolean
    If Not hasRadarWidget Then Exit Function
    isRadarDragging = pRadarWidget.m_handlingDrag
End Function

' MARK: - RadarWidget's widget class
' (VB) widget class RadarWidget : WIDGET

' RadarWidget() : WIDGET()
Public Sub RadarWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_RadarWidget
End Sub

Public Function RadarWidget_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_RadarWidget)
    RadarWidget_construct widget
    Set RadarWidget_new = widget
End Function

Public Function RadarWidget_capturesMouseDrag(ByVal self As WIDGET, ByVal key As WIDGET_KEY) As Boolean
    RadarWidget_capturesMouseDrag = RadarWidgetUDT_capturesMouseDrag(key)
End Function

Public Sub RadarWidget_mouseDragged(ByVal self As WIDGET, ByVal key As WIDGET_KEY, start As W_CONTEXT, current As W_CONTEXT)
    Dim screenContext As W_CONTEXT
    screenContext = W_CONTEXT_convertToScreenContext(current)
    RadarWidgetUDT_mouseDragged key, start.mx, start.my, current.mx, current.my, screenContext.mx, screenContext.my
End Sub

Public Sub RadarWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    RadarWidgetUDT_display xOffset, yOffset
End Sub

Public Function RadarWidget_hitTest(ByVal self As WIDGET, ByVal x As Long, ByVal y As Long) As Boolean
    RadarWidget_hitTest = RadarWidgetUDT_hitTest(x, y)
End Function

Public Sub RadarWidget_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    RadarWidgetUDT_clicked psContext.xOffset, psContext.yOffset, psContext.mx, psContext.my, key
End Sub

Public Sub RadarWidget_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    RadarWidgetUDT_highlight
End Sub

Public Sub RadarWidget_highlightLost(ByVal self As WIDGET)
    RadarWidgetUDT_highlightLost
End Sub

Public Sub RadarWidget_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    RadarWidgetUDT_released psContext.xOffset, psContext.yOffset, psContext.mx, psContext.my, key
End Sub
