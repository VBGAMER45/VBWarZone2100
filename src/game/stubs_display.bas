Attribute VB_Name = "stubs_display"
' Cluster E's merge: what the 3D display's modules (display3d_c and its parts, display_c,
' terrain_c, effects_c, radar_c, warcam_c, ...) call that no module provides yet. Stand-ins keep
' the C++ name and signature. Some are real ports of headers' inline functions or of one
' function of a module not ported yet; their comments say where they belong ("REAL, FOR",
' "->", "P:"), and they move there when that module comes.
Option Explicit





Private Declare Function GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX_) As Long

Private Type MEMORYSTATUSEX_
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As Currency
    ullAvailPhys As Currency
    ullTotalPageFile As Currency
    ullAvailPageFile As Currency
    ullTotalVirtual As Currency
    ullAvailVirtual As Currency
    ullAvailExtendedVirtual As Currency
End Type

' lib/framework/cursors.h: enum CURSOR (stubs_game's Const CURSOR_DEFAULT = 8 goes)
Public Enum CURSOR
    CURSOR_ARROW
    CURSOR_DEST
    CURSOR_SIGHT
    CURSOR_TARGET
    CURSOR_LARROW
    CURSOR_RARROW
    CURSOR_DARROW
    CURSOR_UARROW
    CURSOR_DEFAULT
    CURSOR_EDGEOFMAP
    CURSOR_ATTACH
    CURSOR_ATTACK
    CURSOR_BOMB
    CURSOR_BRIDGE
    CURSOR_BUILD
    CURSOR_EMBARK
    CURSOR_DISEMBARK
    CURSOR_FIX
    CURSOR_GUARD
    CURSOR_JAM
    CURSOR_LOCKON
    CURSOR_SCOUT
    CURSOR_MENU
    CURSOR_MOVE
    CURSOR_NOTPOSSIBLE
    CURSOR_PICKUP
    CURSOR_SEEKREPAIR
    CURSOR_SELECT
    CURSOR_MAX
End Enum




' src/intimage.h (cluster G): enum FRAMETYPE


' src/hci.h (cluster G): IMAGEFILE *IntImages (intfac.img; loaded by hci's intInitialise): WzGfx's image
' sheet handle, 0 = nullptr (S: nothing loads it yet, so the radar's north arrow, the rank and group
' images are not drawn until cluster G)
Public IntImages As Long


' src/ingameop.h (cluster G)

' src/console.h: const DEBOUNCED_MESSAGE CANNOT_BUILD_BURNING = {2500}; (the message's debounce time)

' lib/framework/input's InputContext (cluster G: the key mappings)



'--------------------------------------------------------------------------
' REAL, FOR src/game/map_h.bas: map.h's
' WZ_DECL_ALWAYS_INLINE static inline bool hasSensorOnTile(MAPTILE *psTile, unsigned player)
' psTile: psMapTiles' index (mapTile's)
Public Function hasSensorOnTile(ByVal psTile As Long, ByVal player As Long) As Boolean
    If player = selectedPlayer And godMode Then
        hasSensorOnTile = True
        Exit Function
    End If
    ' (player < MAX_PLAYER_SLOTS) && (alliancebits[selectedPlayer] & (satuplinkbits | psTile->sensorBits))
    If ULT(player, MAX_PLAYER_SLOTS) Then
        hasSensorOnTile = (alliancebits(selectedPlayer) And (satuplinkbits Or psMapTiles(psTile).sensorBits)) <> 0
    End If
End Function






'--------------------------------------------------------------------------
' multistat.cpp:1219 uint32_t getSelectedPlayerUnitsKilled(): the multiplayer kill count outside the campaign
' (getMultiPlayUnitsKilled: multistat.cpp's per-player counts, not ported yet). Stand-in: missionData's.
Public Function getSelectedPlayerUnitsKilled() As Long
    getSelectedPlayerUnitsKilled = missionData.unitsKilled
End Function


' init.cpp: std::vector<TerrainShaderQuality> getAvailableTerrainShaderQualityTextures() (init_c has no
' port of it yet: it needs init_c's searchPathRegistry). Stand-in: every quality, as the data folder
' ships both texture packs (data/terrain_overrides/classic and high). The values in result (0-based),
' their number returned.
Public Function getAvailableTerrainShaderQualityTextures(result() As Long) As Long
    ReDim result(0 To 2)
    result(0) = CLASSIC
    result(1) = MEDIUM
    result(2) = NORMAL_MAPPING
    getAvailableTerrainShaderQualityTextures = 3
End Function

' lib/framework/wzapp.h, lib/sdl/main_sdl.cpp: uint64_t wzGetCurrentSystemRAM() // gets the system RAM in MiB
' (SDL_GetSystemRAM; WzPlatform's when it gets it). A uint64_t is a Double.
Public Function wzGetCurrentSystemRAM() As Double
    Dim m As MEMORYSTATUSEX_
    m.dwLength = LenB(m)
    If GlobalMemoryStatusEx(m) = 0 Then
        wzGetCurrentSystemRAM = 0
        Exit Function
    End If
    ' ullTotalPhys in bytes (a Currency holds the 64 bits scaled by 10000)
    wzGetCurrentSystemRAM = Fix(CDbl(m.ullTotalPhys) * 10000# / 1048576#)
End Function

' src/wzcrashhandlingproviders.h (the crash handler is not ported): no provider takes the tag
Public Function crashHandlingProviderSetTag(ByVal key As String, ByVal value As String) As Boolean
    crashHandlingProviderSetTag = False
End Function

' src/wzcrashhandlingproviders.h (the crash handler is not ported); additionalData: nullptr
Public Function crashHandlingProviderCaptureException(ByVal type_ As String, ByVal value As String, ByVal description As String, ByVal captureStackTrace As Boolean, Optional ByVal handled As Boolean = True) As Boolean
    crashHandlingProviderCaptureException = False
End Function

Private Function justBeenHitByEW(ByVal psObj As Long) As Boolean
    If gamePaused() Then
        justBeenHitByEW = False
        Exit Function
    End If

    Select Case SIMPLE_OBJECT_type(psObj)
    Case OBJ_DROID
        If ULT(gameTime - aDroid(psObj).timeLastHit, ELEC_DAMAGE_DURATION) _
           And aDroid(psObj).lastHitWeapon = WSC_ELECTRONIC Then
            justBeenHitByEW = True
            Exit Function
        End If

    Case OBJ_FEATURE
        If ULT(gameTime - aFeature(psObj).timeLastHit, ELEC_DAMAGE_DURATION) Then
            justBeenHitByEW = True
            Exit Function
        End If

    Case OBJ_STRUCTURE
        If ULT(gameTime - aStruct(psObj).timeLastHit, ELEC_DAMAGE_DURATION) _
           And aStruct(psObj).lastHitWeapon = WSC_ELECTRONIC Then
            justBeenHitByEW = True
            Exit Function
        End If

    Case Else
        ASSERT_FAILURE "false", "src/projectile.cpp:1800", "justBeenHitByEW", "Unknown or invalid object for EW: %s", objInfo(psObj)
        justBeenHitByEW = False
        Exit Function
    End Select

    justBeenHitByEW = False
End Function

' glm::mat4 objectShimmy(BASE_OBJECT *psObj) (display3ddef.h's SKY_SHIMMY: (DEG(1)*SKY_MULT)/2 - rand()%(2*...);
' the three rand() calls taken in the order of the rotations)
Public Function objectShimmy(ByVal psObj As Long) As mat4
    Dim rotations As mat4, sx As Long, sy As Long, sz As Long, tx As Long, tz As Long
    Const SKY_SHIMMY_BASE As Long = 182
    If justBeenHitByEW(psObj) Then
        sx = SKY_SHIMMY_BASE - (rand() Mod (2 * SKY_SHIMMY_BASE))
        sy = SKY_SHIMMY_BASE - (rand() Mod (2 * SKY_SHIMMY_BASE))
        sz = SKY_SHIMMY_BASE - (rand() Mod (2 * SKY_SHIMMY_BASE))
        rotations = glm_mat4_mul(glm_mat4_mul( _
            glm_rotate(UNDEG(sx), glm_vec3(1!, 0!, 0!)), _
            glm_rotate(UNDEG(sy), glm_vec3(0!, 1!, 0!))), _
            glm_rotate(UNDEG(sz), glm_vec3(0!, 0!, 1!)))
        If SIMPLE_OBJECT_type(psObj) <> OBJ_DROID Then
            objectShimmy = rotations
            Exit Function
        End If
        tx = 1 - rand() Mod 3
        tz = 1 - rand() Mod 3
        objectShimmy = glm_mat4_mul(rotations, glm_translate(glm_vec3(CSng(tx), 0!, CSng(tz))))
        Exit Function
    End If
    objectShimmy = glm_mat4(1!)
End Function

' droid.cpp: #define DEFAULT_RECOIL_TIME (GAME_TICKS_PER_SEC/4)
Public Function getRecoil(weapon_ As WEAPON) As Long
    Const DEFAULT_RECOIL_TIME As Long = GAME_TICKS_PER_SEC \ 4
    Dim recoilTime As Long, recoilAmount As Long, maxRecoil As Long
    If weapon_.nStat <> 0 Then
        ' We have a weapon.
        If Not ULT(graphicsTime, weapon_.lastFired) And ULT(graphicsTime, weapon_.lastFired + DEFAULT_RECOIL_TIME) Then
            recoilTime = graphicsTime - weapon_.lastFired
            recoilAmount = DEFAULT_RECOIL_TIME \ 2 - Abs(recoilTime - DEFAULT_RECOIL_TIME \ 2)
            maxRecoil = asWeaponStats(weapon_.nStat).recoilValue  ' Max recoil is 1/10 of this value.
            getRecoil = (maxRecoil * recoilAmount) \ ((DEFAULT_RECOIL_TIME \ 2) * 10)
            Exit Function
        End If
        ' Recoil effect is over.
    End If
    getRecoil = 0
End Function

Public Function getInterpolatedWeaponRotation(ByVal psDroid As Long, ByVal weaponSlot As Long, ByVal time_ As Long) As Rotation
    getInterpolatedWeaponRotation = interpolateRot(aDroid(psDroid).asWeaps(weaponSlot).prevRot, aDroid(psDroid).asWeaps(weaponSlot).rot, aDroid(psDroid).prevSpacetime.time_, aDroid(psDroid).time_, time_)
End Function

Public Function structureGetInterpolatedWeaponRotation(ByVal psStructure As Long, ByVal weaponSlot As Long, ByVal time_ As Long) As Rotation
    structureGetInterpolatedWeaponRotation = interpolateRot(aStruct(psStructure).asWeaps(weaponSlot).prevRot, aStruct(psStructure).asWeaps(weaponSlot).rot, aStruct(psStructure).prevTime, aStruct(psStructure).time_, time_)
End Function



' safeGetDisplayModelFromBase (ivisdef.h)
Private Function safeGetDisplayModelFromBase(ByVal pBaseIMD As Long) As Long
    If pBaseIMD <> 0 Then safeGetDisplayModelFromBase = iIMDBaseShape_displayModel(pBaseIMD)
End Function


' src/basedef.h: the SIMPLE_OBJECT / BASE_OBJECT fields display3d's overlays read through a BASE_OBJECT *
' (P: basedef_h, next to BASE_OBJECT_sDisplay_imd)
Public Property Get BASE_OBJECT_sDisplay_frameNumber(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_sDisplay_frameNumber = aDroid(psObj).sDisplay.frameNumber
    Case OBJ_STRUCTURE: BASE_OBJECT_sDisplay_frameNumber = aStruct(psObj).sDisplay.frameNumber
    Case OBJ_FEATURE: BASE_OBJECT_sDisplay_frameNumber = aFeature(psObj).sDisplay.frameNumber
    End Select
End Property

Public Property Get BASE_OBJECT_sDisplay_screenX(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_sDisplay_screenX = aDroid(psObj).sDisplay.screenX
    Case OBJ_STRUCTURE: BASE_OBJECT_sDisplay_screenX = aStruct(psObj).sDisplay.screenX
    Case OBJ_FEATURE: BASE_OBJECT_sDisplay_screenX = aFeature(psObj).sDisplay.screenX
    End Select
End Property

Public Property Get BASE_OBJECT_sDisplay_screenY(ByVal psObj As Long) As Long
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_sDisplay_screenY = aDroid(psObj).sDisplay.screenY
    Case OBJ_STRUCTURE: BASE_OBJECT_sDisplay_screenY = aStruct(psObj).sDisplay.screenY
    Case OBJ_FEATURE: BASE_OBJECT_sDisplay_screenY = aFeature(psObj).sDisplay.screenY
    End Select
End Property

Public Property Get BASE_OBJECT_selected(ByVal psObj As Long) As Boolean
    Select Case psObj \ OBJ_HANDLE_UNIT
    Case OBJ_DROID: BASE_OBJECT_selected = aDroid(psObj).selected
    Case OBJ_STRUCTURE: BASE_OBJECT_selected = aStruct(psObj).selected
    Case OBJ_FEATURE: BASE_OBJECT_selected = aFeature(psObj).selected
    End Select
End Property

' src/geometry.h (P: geometry_c)
'/ WARNING: Returns NULL if tile not visible to selectedPlayer.
'/ Must *NOT* be used for anything game-state/simulation-calculation related
Public Function getTileOccupier(ByVal x As Long, ByVal y As Long) As Long
    Dim psTile As Long
    psTile = mapTile(x, y)

    If TEST_TILE_VISIBLE_TO_SELECTEDPLAYER(psTile) Then
        getTileOccupier = psMapTiles(mapTile(x, y)).psObject
    Else
        getTileOccupier = 0
    End If
End Function




' src/raycast.cpp: void getBestPitchToEdgeOfGrid(UDWORD x, UDWORD y, uint16_t direction, uint16_t *pitch) (S: the
' camera's pitch is 0 until raycast_c gets the getTileHeightCallback ray (a RAY_CALLBACK number))
Public Sub getBestPitchToEdgeOfGrid(ByVal x As Long, ByVal y As Long, ByVal direction As Long, pitch As Long)
    pitch = 0
End Sub

' src/oprint.cpp: void printDroidInfo(const DROID *psDroid) (S)
Public Sub printDroidInfo(ByVal psDroid As Long)
End Sub

' src/hci.cpp, src/oprint.cpp: void printStructureInfo(STRUCTURE *psStructure) (S: hci, cluster G)
Public Sub printStructureInfo(ByVal psStructure As Long)
End Sub



















' src/display3d.cpp (fork O's renderDeliveryPoint): (VB) renderDeliveryPoint for a FLAG_POSITION that is not in
' aFlagPos (display.cpp's flagPos copy): if fork O's takes a handle, copy flagPos into a spare aFlagPos slot (S)
Public Sub renderDeliveryPoint_flag(psPosition As FLAG_POSITION, ByVal blueprint_ As Boolean, viewMatrix As mat4, perspectiveViewMatrix As mat4)
End Sub
