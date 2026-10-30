Attribute VB_Name = "advvis_c"
' src/advvis.cpp, advvis.h: smooth transitions for terrain visibility (the
' tiles' level, for the terrain's light map; display only).
Option Explicit

' ------------------------------------------------------------------------------------
Private Const FADE_IN_TIME As Long = GAME_TICKS_PER_SEC \ 10
Private Const START_DIVIDE As Long = 8
Private Const MIN_ILLUM As Single = 45!

' Whether unexplored tiles should be shown as just darker fog. Left here as a future option
' for scripts, since campaign may still want total darkness on unexplored tiles.
Private bRevealActive As Boolean        ' = true

' (VB) the C++ initialisers (Sub Main calls it)
Public Sub advvis_c_init()
    bRevealActive = True
End Sub

' For display only (*NOT* for use in game state calculations)
Private Function getTileIllumination(ByVal psTile As Long) As Single
    getTileIllumination = psMapTiles(psTile).ambientOcclusion ' sunlight is handled by shaders so only AO needed for lightmap
End Function

' ------------------------------------------------------------------------------------
Public Sub avUpdateTiles()
    Dim len_ As Long, playermask As Long, i As Long, maxLevel As Single, increment As Single
    Dim playerAllianceBits As Long
    len_ = mapHeight * mapWidth
    playermask = BIT_(selectedPlayer)
    increment = graphicsTimeAdjustedIncrement(FADE_IN_TIME)    ' call once per frame

    If selectedPlayer < MAX_PLAYER_SLOTS Then playerAllianceBits = alliancebits(selectedPlayer) Else playerAllianceBits = 0

    ' Go through the tiles
    For i = 0 To len_ - 1
        maxLevel = getTileIllumination(i)

        If psMapTiles(i).level > MIN_ILLUM Or (psMapTiles(i).tileExploredBits And playermask) <> 0 Then ' seen
            ' If we are not omniscient, and we are not seeing the tile, and none of our allies see the tile...
            If Not godMode And (playerAllianceBits And (satuplinkbits Or psMapTiles(i).sensorBits)) = 0 Then
                maxLevel = maxLevel / 2!
            End If
            If psMapTiles(i).level > maxLevel Then
                psMapTiles(i).level = psMapTiles(i).level - increment
                If psMapTiles(i).level < maxLevel Then psMapTiles(i).level = maxLevel
            ElseIf psMapTiles(i).level < maxLevel Then
                psMapTiles(i).level = psMapTiles(i).level + increment
                If psMapTiles(i).level > maxLevel Then psMapTiles(i).level = maxLevel
            End If
        End If
    Next
End Sub

' ------------------------------------------------------------------------------------
Public Function avGetObjLightLevel(ByVal psObj As Long, ByVal origLevel As Long) As Long
    Dim div As Single, lowest As Long, newLevel As Long
    div = CSng(BASE_OBJECT_visibleForLocalDisplay(psObj)) / 255!
    lowest = origLevel \ START_DIVIDE
    newLevel = CLng(Fix(div * origLevel))

    If newLevel < lowest Then
        newLevel = lowest
    End If

    avGetObjLightLevel = newLevel
End Function

' ------------------------------------------------------------------------------------
Public Function getRevealStatus() As Boolean
    getRevealStatus = bRevealActive
End Function

' ------------------------------------------------------------------------------------
Public Sub setRevealStatus(ByVal val_ As Boolean)
    Dim s As String
    If val_ Then s = "ON" Else s = "OFF"
    debug_ 108, LOG_FOG, "setRevealStatus", "avSetRevealStatus: Setting reveal to %s", s
    bRevealActive = val_
End Sub

' ------------------------------------------------------------------------------------
Public Sub preProcessVisibility()
    Dim i As Long, j As Long, psTile As Long, v As Single
    For i = 0 To mapWidth - 1
        For j = 0 To mapHeight - 1
            psTile = mapTile(i, j)
            If bRevealActive Then
                v = getTileIllumination(psTile) / 4!
                If v > MIN_ILLUM Then v = MIN_ILLUM
                psMapTiles(psTile).level = v
            Else
                psMapTiles(psTile).level = 0!
            End If

            If TEST_TILE_VISIBLE_TO_SELECTEDPLAYER(psTile) Then
                psMapTiles(psTile).level = getTileIllumination(psTile)
            End If
        Next
    Next
End Sub
' ------------------------------------------------------------------------------------
