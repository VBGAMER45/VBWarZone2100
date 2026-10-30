Attribute VB_Name = "feature_c"
' src/feature.cpp, feature.h: the feature stats (asFeatureStats with its
' count asFeatureStats_size) and their loader, and the features: the
' container (GlobalFeatureContainer), buildFeature, removeFeature,
' destroyFeature, featureUpdate, the sync log. A FEATURE_STATS * is its
' index, -1 = nullptr; a FEATURE * a handle into aFeature (see basedef_h).
' The two buildFeature overloads are one procedure whose id is Optional
' (missing: a new synchronised id, as the four-argument one does).
' Not ported yet: featureDamage (with combat.cpp's objDamage).
Option Explicit

' The statistics for the features
Public asFeatureStats() As FEATURE_STATS
Public asFeatureStats_size As Long

'Value is stored for easy access to this feature in destroyDroid()/destroyStruct()
Public oilResFeature As Long            ' FEATURE_STATS *

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub feature_c_init()
    oilResFeature = -1
End Sub

Public Sub featureInitVars()
    Erase asFeatureStats
    asFeatureStats_size = 0
    oilResFeature = -1
End Sub

' Load the feature stats
Public Function loadFeatureStats(ByVal ini As WzConfig) As Boolean
    Dim list() As String, n As Long, i As Long, p As Long, subType As String

    If Not (ini.isAtDocumentRoot()) Then ASSERT_FAILURE "ini.isAtDocumentRoot()", "src/feature.cpp:70", "loadFeatureStats", "WzConfig instance is in the middle of traversal"
    n = ini.childGroups(list)
    ' (reserve: the array grows by n now)
    If asFeatureStats_size + n > 0 Then ReDim Preserve asFeatureStats(0 To asFeatureStats_size + n - 1)
    For i = 0 To n - 1
        ini.beginGroup list(i)
        ' emplace_back(STAT_FEATURE + i)
        FEATURE_STATS_construct asFeatureStats(asFeatureStats_size), STAT_FEATURE + i
        asFeatureStats_size = asFeatureStats_size + 1
        p = i                               ' FEATURE_STATS& p = asFeatureStats[i]
        asFeatureStats(p).name_ = ini.string_("name")
        asFeatureStats(p).id = list(i)
        subType = ini.value("type").toWzString()
        If subType = "TANK WRECK" Then
            asFeatureStats(p).subType = FEAT_TANK
        ElseIf subType = "GENERIC ARTEFACT" Then
            asFeatureStats(p).subType = FEAT_GEN_ARTE
        ElseIf subType = "OIL RESOURCE" Then
            asFeatureStats(p).subType = FEAT_OIL_RESOURCE
        ElseIf subType = "BOULDER" Then
            asFeatureStats(p).subType = FEAT_BOULDER
        ElseIf subType = "VEHICLE" Then
            asFeatureStats(p).subType = FEAT_VEHICLE
        ElseIf subType = "BUILDING" Then
            asFeatureStats(p).subType = FEAT_BUILDING
        ElseIf subType = "OIL DRUM" Then
            asFeatureStats(p).subType = FEAT_OIL_DRUM
        ElseIf subType = "TREE" Then
            asFeatureStats(p).subType = FEAT_TREE
        ElseIf subType = "SKYSCRAPER" Then
            asFeatureStats(p).subType = FEAT_SKYSCRAPER
        Else
            ASSERT_FAILURE "false", "src/feature.cpp:119", "loadFeatureStats", "Unknown feature type: %s", subType
        End If
        asFeatureStats(p).psImd = modelGet(ini.value("model").toWzString())
        asFeatureStats(p).baseWidth = ini.value("width", json_variant_new(1)).toInt() And &HFFFF&
        asFeatureStats(p).baseBreadth = ini.value("breadth", json_variant_new(1)).toInt() And &HFFFF&
        asFeatureStats(p).tileDraw = ini.value("tileDraw", json_variant_new(1)).toInt() <> 0
        asFeatureStats(p).allowLOS = ini.value("lineOfSight", json_variant_new(1)).toInt() <> 0
        asFeatureStats(p).visibleAtStart = ini.value("startVisible", json_variant_new(1)).toInt() <> 0
        asFeatureStats(p).damageable = ini.value("damageable", json_variant_new(1)).toInt() <> 0
        asFeatureStats(p).body = ini.value("hitpoints", json_variant_new(1)).toInt()
        asFeatureStats(p).armourValue = ini.value("armour", json_variant_new(1)).toInt()

        'and the oil resource - assumes only one!
        If asFeatureStats(p).subType = FEAT_OIL_RESOURCE Then
            oilResFeature = p
        End If
        ini.endGroup
    Next

    loadFeatureStats = True
End Function

' Release the feature stats memory
Public Sub featureStatsShutDown()
    Erase asFeatureStats
    asFeatureStats_size = 0
End Sub

' (VB) GlobalFeatureContainer().emplace(id, psStats)
Public Function GlobalFeatureContainer_emplace(ByVal id As Long, ByVal psStats As Long) As Long
    Dim psFeature As Long
    psFeature = FeatureContainer_alloc()
    FEATURE_construct psFeature, id, psStats
    GlobalFeatureContainer_emplace = psFeature
End Function

' Get pitch and roll from direction and tile data
Private Sub updateFeatureOrientation(ByVal psFeature As Long)
    Dim hx0 As Long, hx1 As Long, hy0 As Long, hy1 As Long
    Dim newPitch As Long, deltaPitch As Long
    Dim dzdx As Long, dzdy As Long, dzdv As Long, dzdw As Long
    Const d As Long = 20 ' from updateDroidOrientation() - a magic number distance from the feature position
    Dim vX As Long, vY As Long

    ' Find the height of 4 points around the feature center.
    '    hy0
    ' hx0 * hx1      (* = feature)
    '    hy1
    hx1 = map_Height(aFeature(psFeature).pos.x + d, aFeature(psFeature).pos.y)
    hx0 = map_Height(MAX_(0, aFeature(psFeature).pos.x - d), aFeature(psFeature).pos.y)
    hy1 = map_Height(aFeature(psFeature).pos.x, aFeature(psFeature).pos.y + d)
    hy0 = map_Height(aFeature(psFeature).pos.x, MAX_(0, aFeature(psFeature).pos.y - d))

    'update height in case in the bottom of a trough
    aFeature(psFeature).pos.z = MAX_(aFeature(psFeature).pos.z, (hx0 + hx1) \ 2)
    aFeature(psFeature).pos.z = MAX_(aFeature(psFeature).pos.z, (hy0 + hy1) \ 2)

    Select Case asFeatureStats(aFeature(psFeature).psStats).subType
    Case FEAT_TREE, FEAT_SKYSCRAPER, FEAT_BUILDING
        ' Do not rotate or pitch - trees + buildings look weird if they aren't pointing up
        Exit Sub
    End Select

    ' Vector of length 65536 pointing in direction feature is facing.
    vX = iSin(aFeature(psFeature).rot.direction)
    vY = iCos(aFeature(psFeature).rot.direction)

    ' Calculate pitch of ground.
    dzdx = hx1 - hx0                                ' 2*d*dz(x, y)/dx       of ground
    dzdy = hy1 - hy0                                ' 2*d*dz(x, y)/dy       of ground
    dzdv = dzdx * vX + dzdy * vY                    ' 2*d*dz(x, y)/dv << 16 of ground, where v is the direction the droid is facing.
    newPitch = iAtan2(dzdv, (2 * d) * 65536)        ' pitch = atan(dz(x, y)/dv)/2pi << 16

    deltaPitch = angleDelta(newPitch - aFeature(psFeature).rot.pitch)

    ' Update pitch.
    aFeature(psFeature).rot.pitch = (aFeature(psFeature).rot.pitch + deltaPitch) And &HFFFF&

    ' Calculate and update roll of ground (not taking pitch into account, but good enough).
    dzdw = dzdx * vY - dzdy * vX                    ' 2*d*dz(x, y)/dw << 16 of ground, where w is at right angles to the direction the droid is facing.
    aFeature(psFeature).rot.roll = iAtan2(dzdw, (2 * d) * 65536) And &HFFFF&   ' pitch = atan(dz(x, y)/dw)/2pi << 16
End Sub

' Create a feature on the map
' (the id missing: buildFeature(psStats, x, y, FromSave), which takes a new synchronised id)
Public Function buildFeature(ByVal psStats As Long, ByVal x As Long, ByVal y As Long, ByVal FromSave As Boolean, Optional ByVal id As Variant) As Long
    Dim psFeature As Long, b As StructureBounds
    Dim foundationMin As Long, foundationMax As Long, breadth As Long, width As Long, h As Long, height As Long
    Dim psTile As Long, psBlock As Long

    If IsMissing(id) Then id = generateSynchronisedObjectId()

    'try and create the Feature, obtain stable address.
    psFeature = GlobalFeatureContainer_emplace(CLng(id), psStats)

    'add the feature to the list - this enables it to be drawn whilst being built
    addFeature psFeature

    ' snap the coords to a tile
    If Not FromSave Then
        x = (x And Not TILE_MASK) + (asFeatureStats(psStats).baseWidth Mod 2) * TILE_UNITS \ 2
        y = (y And Not TILE_MASK) + (asFeatureStats(psStats).baseBreadth Mod 2) * TILE_UNITS \ 2
    Else
        If (x And TILE_MASK) <> (asFeatureStats(psStats).baseWidth Mod 2) * TILE_UNITS \ 2 Or _
           (y And TILE_MASK) <> (asFeatureStats(psStats).baseBreadth Mod 2) * TILE_UNITS \ 2 Then
            debug_ 255, LOG_WARNING, "buildFeature", "Feature not aligned. position (%d,%d), size (%d,%d)", x, y, asFeatureStats(psStats).baseWidth, asFeatureStats(psStats).baseBreadth
        End If
    End If

    aFeature(psFeature).pos.x = x
    aFeature(psFeature).pos.y = y

    b = getStructureBounds(psFeature)

    ' get the terrain average height
    foundationMin = INT32_MAX
    foundationMax = INT32_MIN
    For breadth = 0 To b.size.y
        For width = 0 To b.size.x
            h = map_TileHeight(b.map.x + width, b.map.y + breadth)
            foundationMin = MIN_(foundationMin, h)
            foundationMax = MAX_(foundationMax, h)
        Next
    Next
    'return the average of max/min height
    height = (foundationMin + foundationMax) \ 2

    If asFeatureStats(psStats).subType = FEAT_TREE Then
        aFeature(psFeature).rot.direction = gameRand(DEG_360) And &HFFFF&
    Else
        aFeature(psFeature).rot.direction = 0
    End If
    aFeature(psFeature).body = asFeatureStats(psStats).body
    aFeature(psFeature).periodicalDamageStart = 0
    aFeature(psFeature).periodicalDamage = 0
    If foundationMin < TILE_MAX_HEIGHT Then aFeature(psFeature).foundationDepth = foundationMin Else aFeature(psFeature).foundationDepth = TILE_MAX_HEIGHT

    ' it has never been drawn
    aFeature(psFeature).sDisplay.frameNumber = 0

    For h = 0 To MAX_PLAYERS - 1
        aFeature(psFeature).seenThisTick(h) = 0
        aFeature(psFeature).visible(h) = 0
    Next

    ' set up the imd for the feature
    aFeature(psFeature).sDisplay.imd = asFeatureStats(psStats).psImd

    If Not (aFeature(psFeature).sDisplay.imd <> 0) Then ASSERT_FAILURE "psFeature->sDisplay.imd", "src/feature.cpp:301", "buildFeature", "No IMD for feature": buildFeature = 0: Exit Function       ' make sure we have an imd.

    For breadth = 0 To b.size.y - 1
        For width = 0 To b.size.x - 1
            psTile = mapTile(b.map.x + width, b.map.y + breadth)

            'check not outside of map - for load save game
            If Not (b.map.x + width < mapWidth) Then ASSERT_FAILURE "b.map.x + width < mapWidth", "src/feature.cpp:310", "buildFeature", "x coord bigger than map width - %s, id = %d", asFeatureStats(aFeature(psFeature).psStats).name_, aFeature(psFeature).id: buildFeature = 0: Exit Function
            If Not (b.map.y + breadth < mapHeight) Then ASSERT_FAILURE "b.map.y + breadth < mapHeight", "src/feature.cpp:311", "buildFeature", "y coord bigger than map height - %s, id = %d", asFeatureStats(aFeature(psFeature).psStats).name_, aFeature(psFeature).id: buildFeature = 0: Exit Function

            If width <> asFeatureStats(psStats).baseWidth And breadth <> asFeatureStats(psStats).baseBreadth Then
                If TileHasFeature(psTile) Then
                    psBlock = psMapTiles(psTile).psObject

                    debug_ 319, LOG_ERROR, "buildFeature", "%s(%d) already placed at (%d+%d, %d+%d) when trying to place %s(%d) at (%d+%d, %d+%d) - removing it", _
                           asFeatureStats(aFeature(psBlock).psStats).name_, aFeature(psBlock).id, map_coord(aFeature(psBlock).pos.x), asFeatureStats(aFeature(psBlock).psStats).baseWidth, map_coord(aFeature(psBlock).pos.y), _
                           asFeatureStats(aFeature(psBlock).psStats).baseBreadth, asFeatureStats(aFeature(psFeature).psStats).name_, aFeature(psFeature).id, b.map.x, b.size.x, b.map.y, b.size.y

                    removeFeature psBlock
                End If

                psMapTiles(psTile).psObject = psFeature

                ' if it's a tall feature then flag it in the map.
                If imdMaxY(aFeature(psFeature).sDisplay.imd) > TALLOBJECT_YMAX Then
                    auxSetBlocking b.map.x + width, b.map.y + breadth, AIR_BLOCKED
                End If

                If asFeatureStats(psStats).subType <> FEAT_GEN_ARTE And asFeatureStats(psStats).subType <> FEAT_OIL_DRUM Then
                    auxSetBlocking b.map.x + width, b.map.y + breadth, FEATURE_BLOCKED
                End If
            End If

            If (Not asFeatureStats(psStats).tileDraw) And (FromSave = False) Then
                psMapTiles(psTile).height = height
            End If
        Next
    Next
    aFeature(psFeature).pos.z = map_TileHeight(aFeature(psFeature).pos.x, aFeature(psFeature).pos.y) 'jps 18july97
    updateFeatureOrientation psFeature

    buildFeature = psFeature
End Function

' FEATURE::FEATURE(uint32_t id, FEATURE_STATS const *psStats), with BASE_OBJECT's and SIMPLE_OBJECT's constructors
Public Sub FEATURE_construct(ByVal this As Long, ByVal id As Long, ByVal psStats As Long)
    Dim i As Long
    ' SIMPLE_OBJECT(OBJ_FEATURE, id, PLAYER_FEATURE)  ' Set the default player out of range to avoid targeting confusions
    aFeature(this).typ = OBJ_FEATURE
    aFeature(this).id = id
    aFeature(this).player = PLAYER_FEATURE
    aFeature(this).born = gameTime
    aFeature(this).died = 0
    aFeature(this).time_ = 0
    ' BASE_OBJECT(OBJ_FEATURE, id, PLAYER_FEATURE)
    aFeature(this).group = 0
    aFeature(this).repairGroup = UBYTE_MAX
    aFeature(this).selected = 0
    aFeature(this).lastEmission = 0
    aFeature(this).lastHitWeapon = WSC_NUM_WEAPON_SUBCLASSES  ' No such weapon.
    aFeature(this).timeLastHit = UDWORD_MAX
    aFeature(this).body = 0
    aFeature(this).periodicalDamageStart = 0
    aFeature(this).periodicalDamage = 0
    aFeature(this).timeAnimationStarted = 0
    aFeature(this).animationEvent = ANIM_EVENT_NONE
    For i = 0 To MAX_PLAYERS - 1
        aFeature(this).visible(i) = 0
    Next
    aFeature(this).sDisplay.imd = 0
    aFeature(this).sDisplay.frameNumber = 0
    aFeature(this).sDisplay.screenX = 0
    aFeature(this).sDisplay.screenY = 0
    aFeature(this).sDisplay.screenR = 0
    aFeature(this).hasExtraFunction = False
    ' FEATURE
    aFeature(this).psStats = psStats
    aFeature(this).foundationDepth = 0!
End Sub

' Release the resources associated with a feature
' (FEATURE::~FEATURE, with ~BASE_OBJECT and ~SIMPLE_OBJECT after it)
Public Sub FEATURE_destruct(ByVal this As Long)
    ' ~BASE_OBJECT
    visRemoveVisibility this
    Erase aFeature(this).watchedTiles
    aFeature(this).watchedTiles_size = 0
    ' ~SIMPLE_OBJECT
    aFeature(this).typ = aFeature(this).typ + 1000000000    ' Hopefully this will trigger an assert              if someone uses the freed object.
    aFeature(this).player = (aFeature(this).player + 100) And &HFF&    ' Hopefully this will trigger an assert and/or crash if someone uses the freed object.
End Sub

' syncDebugFeature(psFeature, ch) is syncDebugFeature_ function, psFeature, ch
Public Sub syncDebugFeature_(ByVal function_ As String, ByVal psFeature As Long, ByVal ch As Long)
    Dim list(0 To 8) As Long
    If aFeature(psFeature).typ <> OBJ_FEATURE Then
        ASSERT_FAILURE "false", "src/feature.cpp:366", "_syncDebugFeature", "%c Broken psFeature->type %u!", ch, aFeature(psFeature).typ
        syncDebug "_syncDebugFeature", "Broken psFeature->type %u!", aFeature(psFeature).typ
    End If
    list(0) = ch

    list(1) = aFeature(psFeature).id

    list(2) = aFeature(psFeature).player
    list(3) = aFeature(psFeature).pos.x
    list(4) = aFeature(psFeature).pos.y
    list(5) = aFeature(psFeature).pos.z
    list(6) = asFeatureStats(aFeature(psFeature).psStats).subType
    list(7) = -CLng(asFeatureStats(aFeature(psFeature).psStats).damageable)
    list(8) = aFeature(psFeature).body
    syncDebugIntList_ function_, "%c feature%d = p%d;pos(%d,%d,%d),subtype%d,damageable%d,body%d", list, 9
End Sub

' Update routine for features
Public Sub featureUpdate(ByVal psFeat As Long)
    syncDebugFeature_ "featureUpdate", psFeat, Asc("<")

    ' Update the periodical damage data
    If aFeature(psFeat).periodicalDamageStart <> 0 And aFeature(psFeat).periodicalDamageStart <> gameTime - deltaGameTime Then  ' -deltaGameTime, since projectiles are updated after features.
        ' The periodicalDamageStart has been set, but is not from the previous tick, so we must be out of the periodical damage.
        aFeature(psFeat).periodicalDamage = 0  ' Reset periodical damage done this tick.
        ' Finished periodical damaging
        aFeature(psFeat).periodicalDamageStart = 0
    End If

    syncDebugFeature_ "featureUpdate", psFeat, Asc(">")
End Sub

' free up a feature with no visual effects
Public Function removeFeature(ByVal psDel As Long) As Boolean
    Dim psMessage As Long, pos As Vector3i, b As StructureBounds
    Dim breadth As Long, width As Long, psTile As Long, player As Long, removedAMessage As Boolean

    If Not (psDel <> 0) Then ASSERT_FAILURE "psDel != nullptr", "src/feature.cpp:408", "removeFeature", "Invalid feature pointer": Exit Function
    If Not (aFeature(psDel).died = 0) Then ASSERT_FAILURE "!psDel->died", "src/feature.cpp:409", "removeFeature", "Feature already dead": Exit Function

    'remove from the map data
    b = getStructureBounds(psDel)
    For breadth = 0 To b.size.y - 1
        For width = 0 To b.size.x - 1
            If tileOnMap(b.map.x + width, b.map.y + breadth) Then
                psTile = mapTile(b.map.x + width, b.map.y + breadth)

                If psMapTiles(psTile).psObject = psDel Then
                    psMapTiles(psTile).psObject = 0
                    auxClearBlocking b.map.x + width, b.map.y + breadth, FEATURE_BLOCKED Or AIR_BLOCKED
                End If
            End If
        Next
    Next

    If asFeatureStats(aFeature(psDel).psStats).subType = FEAT_GEN_ARTE Or asFeatureStats(aFeature(psDel).psStats).subType = FEAT_OIL_DRUM Then
        pos.x = aFeature(psDel).pos.x
        pos.z = aFeature(psDel).pos.y
        pos.y = map_Height(pos.x, pos.z) + 30
        addEffect pos, EFFECT_EXPLOSION, EXPLOSION_TYPE_DISCOVERY, False, 0, 0, gameTime - deltaGameTime + 1
        If asFeatureStats(aFeature(psDel).psStats).subType = FEAT_GEN_ARTE Then
            scoreUpdateVar WD_ARTEFACTS_FOUND
            intRefreshScreen
        End If
    End If

    removedAMessage = False
    If asFeatureStats(aFeature(psDel).psStats).subType = FEAT_GEN_ARTE Or asFeatureStats(aFeature(psDel).psStats).subType = FEAT_OIL_RESOURCE Then
        For player = 0 To MAX_PLAYERS - 1
            psMessage = findMessage(psDel, MSG_PROXIMITY, player)
            Do While psMessage <> 0
                removeMessage psMessage, player
                removedAMessage = True
                psMessage = findMessage(psDel, MSG_PROXIMITY, player)
            Loop
        Next
    End If
    If removedAMessage Then
        jsDebugMessageUpdate
    End If

    debug_ 462, LOG_DEATH, "removeFeature", "Killing off feature %s id %d (%p)", objInfo(psDel), aFeature(psDel).id, psDel
    killFeature psDel

    removeFeature = True
End Function

' Remove a Feature and free it's memory
Public Function destroyFeature(ByVal psDel As Long, ByVal impactTime As Long) As Boolean
    Dim widthScatter As Long, breadthScatter As Long, heightScatter As Long, i As Long
    Dim explosionSize As EFFECT_TYPE, pos As Vector3i
    Dim b As StructureBounds, isUrban As Boolean, breadth As Long, width As Long, x As Long, y As Long, psTile As Long

    If Not (psDel <> 0) Then ASSERT_FAILURE "psDel != nullptr", "src/feature.cpp:475", "destroyFeature", "Invalid feature pointer": Exit Function
    If Not (ULT(gameTime - deltaGameTime, impactTime)) Then ASSERT_FAILURE "gameTime - deltaGameTime < impactTime", "src/feature.cpp:476", "destroyFeature", "Expected %u < %u, gameTime = %u, bad impactTime", gameTime - deltaGameTime, impactTime, gameTime

    ' Only add if visible and damageable
    If BASE_OBJECT_visibleForLocalDisplay(psDel) <> 0 And asFeatureStats(aFeature(psDel).psStats).damageable Then
        ' Set off a destruction effect
        ' First Explosions
        widthScatter = TILE_UNITS \ 2
        breadthScatter = TILE_UNITS \ 2
        heightScatter = TILE_UNITS \ 4
        'set which explosion to use based on size of feature
        If asFeatureStats(aFeature(psDel).psStats).baseWidth < 2 And asFeatureStats(aFeature(psDel).psStats).baseBreadth < 2 Then
            explosionSize = EXPLOSION_TYPE_SMALL
        ElseIf asFeatureStats(aFeature(psDel).psStats).baseWidth < 3 And asFeatureStats(aFeature(psDel).psStats).baseBreadth < 3 Then
            explosionSize = EXPLOSION_TYPE_MEDIUM
        Else
            explosionSize = EXPLOSION_TYPE_LARGE
        End If
        For i = 0 To 3
            ' (UDWORD arithmetic: rand() % (2 * scatter) is converted to unsigned)
            pos.x = aFeature(psDel).pos.x + widthScatter - rand() Mod (2 * widthScatter)
            pos.z = aFeature(psDel).pos.y + breadthScatter - rand() Mod (2 * breadthScatter)
            pos.y = aFeature(psDel).pos.z + 32 + rand() Mod heightScatter
            addEffect pos, EFFECT_EXPLOSION, explosionSize, False, 0, 0, impactTime
        Next

        If asFeatureStats(aFeature(psDel).psStats).subType = FEAT_SKYSCRAPER Then
            pos.x = aFeature(psDel).pos.x
            pos.z = aFeature(psDel).pos.y
            pos.y = aFeature(psDel).pos.z
            addEffect pos, EFFECT_DESTRUCTION, DESTRUCTION_TYPE_SKYSCRAPER, True, iIMDBaseShape_displayModel(aFeature(psDel).sDisplay.imd), 0, impactTime
            initPerimeterSmoke iIMDBaseShape_displayModel(aFeature(psDel).sDisplay.imd), pos

            shakeStart 250 ' small shake
        End If

        ' Then a sequence of effects
        pos.x = aFeature(psDel).pos.x
        pos.z = aFeature(psDel).pos.y
        pos.y = map_Height(pos.x, pos.z)
        addEffect pos, EFFECT_DESTRUCTION, DESTRUCTION_TYPE_FEATURE, False, 0, 0, impactTime

        'play sound
        ' ffs gj
        If asFeatureStats(aFeature(psDel).psStats).subType = FEAT_SKYSCRAPER Then
            audio_PlayStaticTrack aFeature(psDel).pos.x, aFeature(psDel).pos.y, ID_SOUND_BUILDING_FALL
        Else
            audio_PlayStaticTrack aFeature(psDel).pos.x, aFeature(psDel).pos.y, ID_SOUND_EXPLOSION
        End If
    End If

    If asFeatureStats(aFeature(psDel).psStats).subType = FEAT_SKYSCRAPER Then
        ' ----- Flip all the tiles under the skyscraper to a rubble tile
        ' smoke effect should disguise this happening
        b = getStructureBounds(psDel)
        isUrban = (currentMapTileset = MAP_TILESET_URBAN)
        For breadth = 0 To b.size.y - 1
            For width = 0 To b.size.x - 1
                x = b.map.x + width
                y = b.map.y + breadth
                psTile = mapTile(x, y)
                If psMapTiles(psTile).psObject <> psDel Then
                    GoTo continue_
                End If
                ' stops water texture changing for underwater features
                If terrainType(psTile) <> TER_WATER Then
                    If terrainType(psTile) <> TER_CLIFFFACE Then
                        ' Clear feature bits
                        If isUrban Then
                            makeTileRubbleTexture psTile, x, y, RUBBLE_TILE
                        End If
                        auxClearBlocking x, y, AUXBITS_ALL
                    Else
                        ' This remains a blocking tile
                        psMapTiles(psTile).psObject = 0
                        auxClearBlocking x, y, AIR_BLOCKED  ' Shouldn't remain blocking for air units, however.
                        If isUrban Then
                            makeTileRubbleTexture psTile, x, y, BLOCKING_RUBBLE_TILE
                        End If
                    End If
                End If
continue_:
            Next
        Next
    End If

    removeFeature psDel
    aFeature(psDel).died = impactTime
    destroyFeature = True
End Function

' feature.h: True iff object is a feature.
Public Function isFeature(ByVal psObject As Long) As Boolean
    If psObject <> 0 Then isFeature = SIMPLE_OBJECT_type(psObject) = OBJ_FEATURE
End Function

' feature.h: Returns FEATURE * if feature or NULL if not.
Public Function castFeature(ByVal psObject As Long) As Long
    If isFeature(psObject) Then castFeature = psObject
End Function

Public Function getFeatureStatFromName(ByVal name_ As String) As Long
    Dim inc As Long
    For inc = 0 To asFeatureStats_size - 1
        If asFeatureStats(inc).id = name_ Then
            getFeatureStatFromName = inc
            Exit Function
        End If
    Next
    getFeatureStatFromName = -1
End Function

' Damage a feature
' \param psFeature feature to deal damage to
' \param damage amount of damage to deal
' \param weaponClass,weaponSubClass the class and subclass of the weapon that deals the damage
' \return < 0 never, >= 0 always
Public Function featureDamage(ByVal psFeature As Long, ByVal damage As Long, ByVal weaponClass As WEAPON_CLASS, ByVal weaponSubClass As WEAPON_SUBCLASS, ByVal impactTime As Long, ByVal isDamagePerSecond As Boolean, ByVal minDamage As Long, ByVal empRadiusHit As Boolean) As Long
    Dim relativeDamage As Long

    If Not (psFeature <> 0) Then ASSERT_FAILURE "psFeature != nullptr", "src/feature.cpp:158", "featureDamage", "Invalid feature pointer": featureDamage = 0: Exit Function

    debug_ 160, LOG_ATTACK, "featureDamage", "feature (id %d): body %d armour %d damage: %d", _
           aFeature(psFeature).id, aFeature(psFeature).body, asFeatureStats(aFeature(psFeature).psStats).armourValue, damage

    relativeDamage = objDamage(psFeature, 0, damage, asFeatureStats(aFeature(psFeature).psStats).body, weaponClass, weaponSubClass, isDamagePerSecond, minDamage, empRadiusHit)

    ' If the shell did sufficient damage to destroy the feature
    If relativeDamage < 0 Then
        debug_ 168, LOG_ATTACK, "featureDamage", "feature (id %d) DESTROYED", aFeature(psFeature).id
        destroyFeature psFeature, impactTime
        featureDamage = relativeDamage * -1
        Exit Function
    Else
        featureDamage = relativeDamage
        Exit Function
    End If
End Function
