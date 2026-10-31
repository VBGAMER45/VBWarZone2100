VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundApi"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' WzSound.dll's public interface. GlobalMultiUse: the exe and the other DLLs
' call these procedures as if they were their own. lib/sound is ported into
' this DLL's modules; this class only forwards to them.
Option Explicit

' sounddefs.h: enum class HRTFMode : int
Public Enum HRTFMode
    HRTFMode_Unsupported = -1
    HRTFMode_Disabled = 0
    HRTFMode_Enabled = 1
    HRTFMode_Auto = 2 ' Can be used for setting - will never be returned by sound_GetHRTFMode
End Enum

' audio_id.h: the sound effects (asAudioID's order, audio_id_c)
Public Enum INGAME_AUDIO
    NO_SOUND = -1

    ' Beeps

    ID_SOUND_WINDOWCLOSE
    ID_SOUND_WINDOWOPEN
    ID_SOUND_BUTTON_CLICK_3
    ID_SOUND_SELECT
    ID_SOUND_BUTTON_CLICK_5
    FE_AUDIO_MESSAGEEND
    ID_SOUND_ZOOM_ON_RADAR
    ID_SOUND_BUILD_FAIL
    ID_SOUND_MESSAGEEND
    ID_SOUND_GAME_SHUTDOWN

    ' Design

    ID_SOUND_TURRET_SELECTED
    ID_SOUND_BODY_SELECTED
    ID_SOUND_PROPULSION_SELECTED
    ID_SOUND_DESIGN_COMPLETED

    ' Structures

    ID_SOUND_CONSTRUCTION_STARTED
    ID_SOUND_STRUCTURE_COMPLETED
    ID_SOUND_STRUCTURE_UNDER_ATTACK
    ID_SOUND_STRUCTURE_REPAIR_IN_PROGRESS
    ID_SOUND_STRUCTURE_DEMOLISHED

    ' Power

    ID_SOUND_POWER_GENERATOR_UNDER_ATTACK
    ID_SOUND_POWER_GENERATOR_DESTROYED
    ID_SOUND_POWER_LOW
    ID_SOUND_RESOURCE_HERE
    ID_SOUND_DERRICK_UNDER_ATTACK
    ID_SOUND_DERRICK_DESTROYED
    ID_SOUND_RESOURCE_DEPLETED
    ID_SOUND_POWER_TRANSFER_IN_PROGRESS
    ID_SOUND_POWER_GENERATOR_REQUIRED

    ' Research

    ID_SOUND_RESEARCH_FACILITY_REQUIRED
    ID_SOUND_ARTIFACT
    ID_SOUND_ARTIFACT_RECOVERED
    ID_SOUND_NEW_RESEARCH_PROJ_AVAILABLE
    ID_SOUND_NEW_STRUCTURE_AVAILABLE
    ID_SOUND_NEW_COMPONENT_AVAILABLE
    ID_SOUND_NEW_CYBORG_AVAILABLE
    ID_SOUND_RESEARCH_COMPLETED
    ID_SOUND_MAJOR_RESEARCH
    ID_SOUND_STRUCTURE_RESEARCH_COMPLETED
    ID_SOUND_POWER_RESEARCH_COMPLETED
    ID_SOUND_COMPUTER_RESEARCH_COMPLETED
    ID_SOUND_VEHICLE_RESEARCH_COMPLETED
    ID_SOUND_SYSTEMS_RESEARCH_COMPLETED
    ID_SOUND_WEAPON_RESEARCH_COMPLETED
    ID_SOUND_CYBORG_RESEARCH_COMPLETED

    ' Production

    ID_SOUND_PRODUCTION_STARTED
    ID_SOUND_DROID_COMPLETED
    ID_SOUND_PRODUCTION_PAUSED
    ID_SOUND_PRODUCTION_CANCELLED
    ID_SOUND_DELIVERY_POINT_ASSIGNED
    ID_SOUND_DELIVERY_POINT_ASSIGNED_TO

    ' Repair

    ID_SOUND_UNIT_REPAIRED

    ' Detection

    ID_SOUND_SCAVENGERS_DETECTED
    ID_SOUND_SCAVENGER_BASE_DETECTED
    ID_SOUND_SCAVENGER_OUTPOST_DETECTED
    ID_SOUND_POWER_RESOURCE
    ID_SOUND_ARTEFACT_DISC
    ID_SOUND_ENEMY_UNIT_DETECTED
    ID_SOUND_ENEMY_BASE_DETECTED
    ID_SOUND_ALLY_DETECTED
    ID_SOUND_ENEMY_TRANSPORT_DETECTED
    ID_SOUND_ENEMY_LZ_DETECTED
    ID_SOUND_FRIENDLY_LZ_DETECTED
    ID_SOUND_NEXUS_TOWER_DETECTED
    ID_SOUND_NEXUS_TURRET_DETECTED
    ID_SOUND_NEXUS_UNIT_DETECTED
    ID_SOUND_ENEMY_BATTERY_DETECTED
    ID_SOUND_ENEMY_VTOLS_DETECTED

    ' Status

    ID_SOUND_SCAVENGER_BASE
    ID_SOUND_SCAVENGER_OUTPOST
    ID_SOUND_SCAVENGER_OUTPOST_ERADICATED
    ID_SOUND_SCAVENGER_BASE_ERADICATED
    ID_SOUND_ENEMY_BASE
    ID_SOUND_ENEMY_BASE_ERADICATED
    ID_SOUND_INCOMING_ENEMY_TRANSPORT
    ID_SOUND_ENEMY_LZ
    ID_SOUND_LZ1
    ID_SOUND_LZ2

    ' Combat

    ID_SOUND_UNIT_UNDER_ATTACK
    ID_SOUND_UNIT_DESTROYED
    ID_SOUND_UNIT_RETREATING
    ID_SOUND_UNIT_RETURNING_FOR_REPAIR

    ' Artillery Batteries
    ID_SOUND_ASSIGNED_TO_SENSOR
    ID_SOUND_SENSOR_LOCKED_ON
    ID_SOUND_ASSIGNED_TO_COUNTER_RADAR
    ID_SOUND_ENEMY_BATTERY_LOCATED
    ID_SOUND_BATTERY_FIRING_COUNTER_ATTACK

    ' Vtols

    ID_SOUND_INTERCEPTORS_LAUNCHED
    ID_SOUND_REARMING
    ID_SOUND_VTOLS_ENGAGING
    ID_SOUND_ASSIGNED
    ID_SOUND_INTERCEPTORS_ASSIGNED

    ' Command Console

    ID_SOUND_COMMAND_CONSOLE_ACTIVATED
    ID_SOUND_SHORT_RANGE
    ID_SOUND_LONG_RANGE
    ID_SOUND_OPTIMUM_RANGE
    ID_SOUND_RETREAT_AT_MEDIUM_DAMAGE
    ID_SOUND_RETREAT_AT_HEAVY_DAMAGE
    ID_SOUND_NO_RETREAT
    ID_SOUND_FIRE_AT_WILL
    ID_SOUND_RETURN_FIRE
    ID_SOUND_CEASEFIRE
    ID_SOUND_HOLD_POSITION
    ID_SOUND_GUARD
    ID_SOUND_PURSUE
    ID_SOUND_PATROL
    ID_SOUND_RETURN_TO_LZ
    ID_SOUND_RECYCLING
    ID_SOUND_SCATTER

    ' Tutorial Stuff

    ID_SOUND_NOT_POSSIBLE_TRY_AGAIN
    ID_SOUND_NO
    ID_SOUND_THAT_IS_INCORRECT
    ID_SOUND_WELL_DONE
    ID_SOUND_EXCELLENT

    ' Group and Commander Assignment

    ID_SOUND_ASSIGNED_TO_COMMANDER
    ID_SOUND_GROUP_REPORTING
    ID_SOUND_COMMANDER_REPORTING

    ' Routing

    ID_SOUND_ROUTE_OBSTRUCTED
    ID_SOUND_NO_ROUTE_AVAILABLE

    ' Transports and LZS

    ID_SOUND_REINFORCEMENTS_AVAILABLE
    ID_SOUND_REINFORCEMENTS_IN_TRANSIT
    ID_SOUND_TRANSPORT_LANDING
    ID_SOUND_TRANSPORT_UNDER_ATTACK
    ID_SOUND_TRANSPORT_REPAIRING
    ID_SOUND_LZ_COMPROMISED
    ID_SOUND_LZ_CLEAR
    ID_SOUND_TRANSPORT_RETURNING_TO_BASE
    ID_SOUND_TRANSPORT_UNABLE_TO_LAND

    ' Mission Messages

    ID_SOUND_MISSION_OBJECTIVE
    ID_SOUND_MISSION_UPDATE
    ID_SOUND_WARZONE_PAUSED
    ID_SOUND_WARZONE_ACTIVE
    ID_SOUND_MISSION_RESULTS
    ID_SOUND_RESEARCH_STOLEN
    ID_SOUND_TECHNOLOGY_TAKEN
    ID_SOUND_INCOMING_TRANSMISSION
    ID_SOUND_INCOMING_INTELLIGENCE_REPORT
    ID_SOUND_MISSION_FAILED
    ID_SOUND_MISSION_SUCCESSFUL
    ID_SOUND_OBJECTIVE_ACCOMPLISHED
    ID_SOUND_OBJECTIVE_FAILED
    ID_SOUND_MISSION_TIMER_ACTIVATED
    ID_SOUND_10_MINUTES_REMAINING
    ID_SOUND_5_MINUTES_REMAINING
    ID_SOUND_3_MINUTES_REMAINING
    ID_SOUND_2_MINUTES_REMAINING
    ID_SOUND_1_MINUTE_REMAINING
    ID_SOUND_UNIT_CAPTURED
    ID_SOUND_SYSTEM_FAILURE_IMMINENT
    ID_SOUND_YOU_ARE_DEFEATED
    ID_SOUND_MISSILE_CODES_DECIPHERED
    ID_SOUND_1ST_MISSILE_CODES_DECIPHERED
    ID_SOUND_2ND_MISSILE_CODES_DECIPHERED
    ID_SOUND_3RD_MISSILE_CODES_DECIPHERED
    ID_SOUND_MISSILE_CODES_CRACKED
    ID_SOUND_ENTERING_WARZONE
    ID_ALLIANCE_ACC
    ID_ALLIANCE_BRO
    ID_ALLIANCE_OFF
    ID_CLAN_ENTER
    ID_CLAN_EXIT
    ID_GIFT
    ID_POWER_TRANSMIT
    ID_SENSOR_DOWNLOAD
    ID_TECHNOLOGY_TRANSFER
    ID_UNITS_TRANSFER

    ' Group and Commander Voices - Male
    ID_SOUND_GROUP
    ID_SOUND_GROUP_0
    ID_SOUND_GROUP_1
    ID_SOUND_GROUP_2
    ID_SOUND_GROUP_3
    ID_SOUND_GROUP_4
    ID_SOUND_GROUP_5
    ID_SOUND_GROUP_6
    ID_SOUND_GROUP_7
    ID_SOUND_GROUP_8
    ID_SOUND_GROUP_9
    ID_SOUND_REPORTING
    ID_SOUND_COMMANDER
    ID_SOUND_COM_SCAVS_DETECTED
    ID_SOUND_COM_SCAV_BASE_DETECTED
    ID_SOUND_COM_SCAV_OUTPOST_DETECTED
    ID_SOUND_COM_RESOURCE_DETECTED
    ID_SOUND_COM_ARTEFACT_DETECTED
    ID_SOUND_COM_ENEMY_DETECTED
    ID_SOUND_COM_ENEMY_BASE_DETECTED
    ID_SOUND_COM_ALLY_DETECTED
    ID_SOUND_COM_ENEMY_TRANSPORT_DETECTED
    ID_SOUND_COM_ENEMY_LZ_DETECTED
    ID_SOUND_COM_FRIENDLY_LZ_DETECTED
    ID_SOUND_COM_NEXUS_TOWER_DETECTED
    ID_SOUND_COM_NEXUS_TURRET_DETECTED
    ID_SOUND_COM_NEXUS_DETECTED
    ID_SOUND_COM_ENEMY_BATTERY_DETECTED
    ID_SOUND_COM_ENEMY_VTOLS_DETECTED
    ID_SOUND_COM_ROUTE_OBSTRUCTED
    ID_SOUND_COM_NO_ROUTE_AVAILABLE
    ID_SOUND_COM_UNABLE_TO_COMPLY
    ID_SOUND_COM_RETURNING_FOR_REPAIR
    ID_SOUND_COM_HEADING_FOR_RALLY_POINT

    ' Radio Clicks

    ID_SOUND_RADIOCLICK_1
    ID_SOUND_RADIOCLICK_2
    ID_SOUND_RADIOCLICK_3
    ID_SOUND_RADIOCLICK_4
    ID_SOUND_RADIOCLICK_5
    ID_SOUND_RADIOCLICK_6

    ' Transport Pilots
    ID_SOUND_APPROACHING_LZ
    ID_SOUND_ALRIGHT_BOYS
    ID_SOUND_GREEN_LIGHT_IN_5
    ID_SOUND_GREEN_LIGHT_IN_4
    ID_SOUND_GREEN_LIGHT_IN_3
    ID_SOUND_GREEN_LIGHT_IN_2
    ID_SOUND_GO_GO_GO
    ID_SOUND_PREPARE_FOR_DUST_OFF

    ' VTol Pilots
    ' Ver-1
    ID_SOUND_ENEMY_LOCATED1
    ID_SOUND_ON_OUR_WAY1
    ID_SOUND_RETURNING_TO_BASE1
    ID_SOUND_LOCKED_ON1
    ID_SOUND_COMMENCING_ATTACK_RUN1
    ID_SOUND_ABORTING_ATTACK_RUN1

    ' Ver-2
    ID_SOUND_ENEMY_LOCATED2
    ID_SOUND_ON_OUR_WAY2
    ID_SOUND_RETURNING_TO_BASE2
    ID_SOUND_LOCKED_ON2
    ID_SOUND_COMMENCING_ATTACK_RUN2
    ID_SOUND_ABORTING_ATTACK_RUN2

    ' Ver-3
    ID_SOUND_ENEMY_LOCATED3
    ID_SOUND_ON_OUR_WAY3
    ID_SOUND_RETURNING_TO_BASE3
    ID_SOUND_LOCKED_ON3
    ID_SOUND_COMMENCING_ATTACK_RUN3
    ID_SOUND_ABORTING_ATTACK_RUN3

    ' The Collective
    ID_SOUND_COLL_CLEANSE_AND_DESTROY
    ID_SOUND_COLL_DESTROYING_BIOLOGICALS
    ID_SOUND_COLL_ATTACK
    ID_SOUND_COLL_FIRE
    ID_SOUND_COLL_ENEMY_DETECTED
    ID_SOUND_COLL_ENGAGING
    ID_SOUND_COLL_STARTING_ATTACK_RUN
    ID_SOUND_COLL_DIE
    ID_SOUND_COLL_INTERCEPT_AND_DESTROY
    ID_SOUND_COLL_ENEMY_DESTROYED

    ' SFX

    ' Weapon Sounds

    ID_SOUND_ROCKET
    ID_SOUND_ROTARY_LASER
    ID_SOUND_GAUSSGUN
    ID_SOUND_LARGE_CANNON
    ID_SOUND_SMALL_CANNON
    ID_SOUND_MEDIUM_CANNON
    ID_SOUND_FLAME_THROWER
    ID_SOUND_PULSE_LASER
    ID_SOUND_BEAM_LASER
    ID_SOUND_MORTAR
    ID_SOUND_HOWITZ_FLIGHT
    ID_SOUND_BABA_MG_1
    ID_SOUND_BABA_MG_2
    ID_SOUND_BABA_MG_3
    ID_SOUND_BABA_MG_HEAVY
    ID_SOUND_BABA_MG_TOWER
    ID_SOUND_SPLASH
    ID_SOUND_ASSAULT_MG
    ID_SOUND_RAPID_CANNON
    ID_SOUND_HIVEL_CANNON
    ID_SOUND_NEXUS_TOWER

    ' Construction sounds
    ID_SOUND_WELD_1
    ID_SOUND_WELD_2
    ID_SOUND_CONSTRUCTION_START
    ID_SOUND_CONSTRUCTION_LOOP
    ID_SOUND_CONSTRUCTION_1
    ID_SOUND_CONSTRUCTION_2
    ID_SOUND_CONSTRUCTION_3
    ID_SOUND_CONSTRUCTION_4

    ' Explosions

    ID_SOUND_EXPLOSION_SMALL
    ID_SOUND_EXPLOSION_LASER
    ID_SOUND_EXPLOSION
    ID_SOUND_EXPLOSION_ANTITANK
    ID_SOUND_RICOCHET_1
    ID_SOUND_RICOCHET_2
    ID_SOUND_RICOCHET_3
    ID_SOUND_BARB_SQUISH
    ID_SOUND_BUILDING_FALL
    ID_SOUND_NEXUS_EXPLOSION

    ' Droid Engine Noises
    ID_SOUND_CONSTRUCTOR_MOVEOFF
    ID_SOUND_CONSTRUCTOR_MOVE
    ID_SOUND_CONSTRUCTOR_SHUTDOWN

    ' Transports
    ID_SOUND_BLIMP_FLIGHT
    ID_SOUND_BLIMP_IDLE
    ID_SOUND_BLIMP_LAND
    ID_SOUND_BLIMP_TAKE_OFF

    ' Vtols
    ID_SOUND_VTOL_LAND
    ID_SOUND_VTOL_OFF
    ID_SOUND_VTOL_MOVE

    ' Treads
    ID_SOUND_TREAD

    ' Hovers
    ID_SOUND_HOVER_MOVE
    ID_SOUND_HOVER_START
    ID_SOUND_HOVER_STOP

    ' Cyborgs
    ID_SOUND_CYBORG_MOVE

    ' Buildings

    ID_SOUND_OIL_PUMP_2
    ID_SOUND_POWER_HUM
    ID_SOUND_POWER_SPARK
    ID_SOUND_STEAM
    ID_SOUND_ECM_TOWER
    ID_SOUND_FIRE_ROAR

    ' Misc

    ID_SOUND_HELP
    ID_SOUND_BARB_SCREAM
    ID_SOUND_BARB_SCREAM2
    ID_SOUND_BARB_SCREAM3
    ID_SOUND_OF_SILENCE
    ID_SOUND_SHIELD_HIT

    ' Extra

    ID_SOUND_LANDING_ZONE
    ID_SOUND_SATELLITE_UPLINK
    ID_SOUND_NASDA_CENTRAL
    ID_SOUND_NUCLEAR_REACTOR
    ID_SOUND_SAM_SITE
    ID_SOUND_MISSILE_SILO
    ID_SOUND_MISSILE_NME_DETECTED
    ID_SOUND_STRUCTURE_CAPTURED
    ID_SOUND_CIVILIAN_RESCUED
    ID_SOUND_CIVILIANS_RESCUED
    ID_SOUND_UNITS_RESCUED
    ID_SOUND_GROUP_RESCUED
    ID_SOUND_GROUP_CAPTURED
    ID_SOUND_OBJECTIVE_CAPTURED
    ID_SOUND_OBJECTIVE_DESTROYED
    ID_SOUND_STRUCTURE_INFECTED
    ID_SOUND_GROUP_INFECTED
    ID_SOUND_OUT_OF_TIME
    ID_SOUND_ENEMY_ESCAPED
    ID_SOUND_ENEMY_ESCAPING
    ID_SOUND_ENEMY_TRANSPORT_LANDING
    ID_SOUND_TEAM_ALPHA_ERADICATED
    ID_SOUND_TEAM_BETA_ERADICATED
    ID_SOUND_TEAM_GAMMA_ERADICATED
    ID_SOUND_TEAM_ALPHA_RESCUED
    ID_SOUND_TEAM_BETA_RESCUED
    ID_SOUND_TEAM_GAMMA_RESCUED
    ID_SOUND_LASER_SATELLITE_FIRING
    ID_SOUND_INCOMING_LASER_SAT_STRIKE

    ' Nexus

    ID_SOUND_NEXUS_DEFENCES_ABSORBED
    ID_SOUND_NEXUS_DEFENCES_NEUTRALISED
    ID_SOUND_NEXUS_LAUGH1
    ID_SOUND_NEXUS_LAUGH2
    ID_SOUND_NEXUS_LAUGH3
    ID_SOUND_NEXUS_PRODUCTION_COMPLETED
    ID_SOUND_NEXUS_RESEARCH_ABSORBED
    ID_SOUND_NEXUS_STRUCTURE_ABSORBED
    ID_SOUND_NEXUS_STRUCTURE_NEUTRALISED
    ID_SOUND_NEXUS_SYNAPTIC_LINK
    ID_SOUND_NEXUS_UNIT_ABSORBED
    ID_SOUND_NEXUS_UNIT_NEUTRALISED

    ' multiplayer sfx

    ID_SOUND_CYBORG_GROUND
    ID_SOUND_CYBORG_HEAVY
    ID_SOUND_EMP
    ID_SOUND_LASER_HEAVY
    ID_SOUND_PLASMA_FLAMER
    ID_SOUND_PARTICLE_GUN
    ID_SOUND_UPLINK

    ' added for upgrade - 2/9/99 AB
    ID_SOUND_LAS_SAT_COUNTDOWN

    ID_SOUND_BEACON

    ' Last ID
    ID_SOUND_NEXT    ' Thanks to this dummy we don't have to redefine ID_MAX_SOUND every time in terms of the preceding enum value
    ID_MAX_SOUND = ID_SOUND_NEXT - 1
End Enum

' track.h: the #defines the game uses
Public Enum track_h
    SAMPLE_NOT_ALLOCATED = -1
    SAMPLE_NOT_FOUND = -3
    SAMPLE_COORD_INVALID = -5
    AUDIO_VOL_MAX = 100                 ' 100L
End Enum

' cdaudio.h: enum SONG_CONTEXT
Public Enum SONG_CONTEXT
    SONG_FRONTEND
    SONG_INGAME
End Enum

' cdaudio.h: enum class MusicGameMode
Public Enum MusicGameMode
    MusicGameMode_MENUS = -1
    MusicGameMode_CAMPAIGN = 0
    MusicGameMode_CHALLENGE
    MusicGameMode_SKIRMISH
    MusicGameMode_MULTIPLAYER
End Enum

' Name and version of this DLL and of the WzFramework it uses, for the log header.
Public Function SoundBuildInfo() As String
    SoundBuildInfo = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision & " (uses " & FrameworkBuildInfo() & ")"
End Function

' (VB) lib/sound's calls into the game (aud.h, the game time, the AUDIO_CALLBACKs): the exe's ISoundGame,
' handed over before audio_Init and taken back (Nothing) before the process ends
Public Sub SoundSetGame(ByVal game As ISoundGame)
    Set soundgame_vb.soundGame = game
End Sub

'--------------------------------------------------------------------------
' audio.h (audio_c). An AUDIO_CALLBACK is the game's number (0 = nullptr), a SIMPLE_OBJECT * the game's
' object handle (0 = nullptr).

Public Function loadAudioEffectFileData(ByVal ini As WzConfig) As Boolean
    loadAudioEffectFileData = audio_c.loadAudioEffectFileData(ini)
End Function

' pStopTrackCallback: the game's number for its AUDIO_CALLBACK
Public Function audio_Init(ByVal pStopTrackCallback As Long, ByVal hrtf As HRTFMode, ByVal really_init As Boolean) As Boolean
    audio_Init = audio_c.audio_Init(pStopTrackCallback, hrtf, really_init)
End Function

Public Sub audio_Update()
    audio_c.audio_Update
End Sub

Public Function audio_Shutdown() As Boolean
    audio_Shutdown = audio_c.audio_Shutdown()
End Function

Public Function audio_Disabled() As Boolean
    audio_Disabled = audio_c.audio_Disabled()
End Function

Public Function audio_Paused() As Boolean
    audio_Paused = audio_c.audio_Paused()
End Function

Public Function audio_SetTrackVals(ByVal fileName As String, ByVal loop_ As Boolean, ByVal volume As Long, ByVal audibleRadius As Long) As Long
    audio_SetTrackVals = audio_c.audio_SetTrackVals(fileName, loop_, volume, audibleRadius)
End Function

Public Sub audio_PlayBuildFailedOnce()
    audio_c.audio_PlayBuildFailedOnce
End Sub

Public Function audio_PlayStaticTrack(ByVal iX As Long, ByVal iY As Long, ByVal iTrack As Long) As Boolean
    audio_PlayStaticTrack = audio_c.audio_PlayStaticTrack(iX, iY, iTrack)
End Function

Public Function audio_PlayObjStaticTrack(ByVal psObj As Long, ByVal iTrack As Long) As Boolean
    audio_PlayObjStaticTrack = audio_c.audio_PlayObjStaticTrack(psObj, iTrack)
End Function

Public Function audio_PlayObjStaticTrackCallback(ByVal psObj As Long, ByVal iTrack As Long, ByVal pUserCallback As Long) As Boolean
    audio_PlayObjStaticTrackCallback = audio_c.audio_PlayObjStaticTrackCallback(psObj, iTrack, pUserCallback)
End Function

Public Function audio_PlayObjDynamicTrack(ByVal psObj As Long, ByVal iTrack As Long, ByVal pUserCallback As Long) As Boolean
    audio_PlayObjDynamicTrack = audio_c.audio_PlayObjDynamicTrack(psObj, iTrack, pUserCallback)
End Function

Public Sub audio_StopObjTrack(ByVal psObj As Long, ByVal iTrack As Long)
    audio_c.audio_StopObjTrack psObj, iTrack
End Sub

Public Sub audio_PlayTrack(ByVal iTrack As Long, Optional ByVal playIfPaused As Boolean = True)
    audio_c.audio_PlayTrack iTrack, playIfPaused
End Sub

' onFinished: the std::function (Nothing = empty); user_data: the const void * handed back to it
Public Function audio_PlayStream(ByVal fileName As String, ByVal volume As Single, ByVal onFinished As IStreamFinished, ByVal user_data As Variant) As AUDIO_STREAM
    Set audio_PlayStream = audio_c.audio_PlayStream(fileName, volume, onFinished, user_data)
End Function

Public Sub audio_QueueTrack(ByVal iTrack As Long)
    audio_c.audio_QueueTrack iTrack
End Sub

Public Sub audio_QueueTrackMinDelay(ByVal iTrack As Long, ByVal iMinDelay As Long)
    audio_c.audio_QueueTrackMinDelay iTrack, iMinDelay
End Sub

Public Sub audio_QueueTrackMinDelayPos(ByVal iTrack As Long, ByVal iMinDelay As Long, ByVal iX As Long, ByVal iY As Long, ByVal iZ As Long)
    audio_c.audio_QueueTrackMinDelayPos iTrack, iMinDelay, iX, iY, iZ
End Sub

Public Sub audio_QueueTrackPos(ByVal iTrack As Long, ByVal iX As Long, ByVal iY As Long, ByVal iZ As Long)
    audio_c.audio_QueueTrackPos iTrack, iX, iY, iZ
End Sub

Public Function audio_GetPreviousQueueTrackPos(iX As Long, iY As Long, iZ As Long) As Boolean
    audio_GetPreviousQueueTrackPos = audio_c.audio_GetPreviousQueueTrackPos(iX, iY, iZ)
End Function

Public Function audio_GetPreviousQueueTrackRadarBlipPos(iX As Long, iY As Long) As Boolean
    audio_GetPreviousQueueTrackRadarBlipPos = audio_c.audio_GetPreviousQueueTrackRadarBlipPos(iX, iY)
End Function

Public Sub audio_setPause(ByVal state As Boolean)
    audio_c.audio_setPause state
End Sub

Public Sub audio_StopAll()
    audio_c.audio_StopAll
End Sub

Public Function audio_GetTrackID(ByVal fileName As String) As Long
    audio_GetTrackID = audio_c.audio_GetTrackID(fileName)
End Function

Public Sub audio_RemoveObj(ByVal psObj As Long)
    audio_c.audio_RemoveObj psObj
End Sub

Public Function audio_GetSampleQueueCount() As Long
    audio_GetSampleQueueCount = audio_c.audio_GetSampleQueueCount()
End Function

Public Function audio_GetSampleListCount() As Long
    audio_GetSampleListCount = audio_c.audio_GetSampleListCount()
End Function

'--------------------------------------------------------------------------
' track.h (track_c, openal_track_c)

' a loaded sound effect, the data of the game's "WAV" resources (Nothing on failure)
Public Function sound_LoadTrackFromFile(ByVal fileName As String) As TRACK
    Set sound_LoadTrackFromFile = openal_track_c.sound_LoadTrackFromFile(fileName)
End Function

' psTrack: the resource's data (a TRACK; Empty when the sound was off as it loaded: nullptr)
Public Sub sound_ReleaseTrack(ByVal psTrack As Variant)
    If IsObject(psTrack) Then track_c.sound_ReleaseTrack psTrack
End Sub

Public Sub sound_CheckAllUnloaded()
    track_c.sound_CheckAllUnloaded
End Sub

Public Function sound_GetActiveSamplesCount() As Long
    sound_GetActiveSamplesCount = openal_track_c.sound_GetActiveSamplesCount()
End Function

Public Function sound_isStreamPlaying(ByVal stream As AUDIO_STREAM) As Boolean
    sound_isStreamPlaying = openal_track_c.sound_isStreamPlaying(stream)
End Function

Public Sub sound_StopStream(ByVal stream As AUDIO_STREAM)
    openal_track_c.sound_StopStream stream
End Sub

Public Sub sound_PauseStream(ByVal stream As AUDIO_STREAM)
    openal_track_c.sound_PauseStream stream
End Sub

Public Sub sound_ResumeStream(ByVal stream As AUDIO_STREAM)
    openal_track_c.sound_ResumeStream stream
End Sub

Public Function sound_GetStreamVolume(ByVal stream As AUDIO_STREAM) As Single
    sound_GetStreamVolume = openal_track_c.sound_GetStreamVolume(stream)
End Function

Public Sub sound_SetStreamVolume(ByVal stream As AUDIO_STREAM, ByVal volume As Single)
    openal_track_c.sound_SetStreamVolume stream, volume
End Sub

Public Function sound_GetStreamTotalTime(ByVal stream As AUDIO_STREAM) As Double
    sound_GetStreamTotalTime = openal_track_c.sound_GetStreamTotalTime(stream)
End Function

'--------------------------------------------------------------------------
' tracklib.h (openal_track_c)

Public Function sound_GetHRTFMode() As HRTFMode
    sound_GetHRTFMode = openal_track_c.sound_GetHRTFMode()
End Function

Public Function sound_SetHRTFMode(ByVal mode As HRTFMode) As Boolean
    sound_SetHRTFMode = openal_track_c.sound_SetHRTFMode(mode)
End Function

' (VB) test aid (--vbsndcapture): before audio_Init, the mix goes to this WAV file and no device opens
Public Sub sound_vbSetCapture(ByVal wavPath As String)
    openal_track_c.sound_vbSetCapture wavPath
End Sub

'--------------------------------------------------------------------------
' mixer.h: volumes 0.0 - 1.0 (openal_track_c, cdaudio_c)

Public Function sound_GetUIVolume() As Single
    sound_GetUIVolume = openal_track_c.sound_GetUIVolume()
End Function

Public Sub sound_SetUIVolume(ByVal volume As Single)
    openal_track_c.sound_SetUIVolume volume
End Sub

Public Function sound_GetEffectsVolume() As Single
    sound_GetEffectsVolume = openal_track_c.sound_GetEffectsVolume()
End Function

Public Sub sound_SetEffectsVolume(ByVal volume As Single)
    openal_track_c.sound_SetEffectsVolume volume
End Sub

Public Function sound_GetMusicVolume() As Single
    sound_GetMusicVolume = cdaudio_c.sound_GetMusicVolume()
End Function

Public Sub sound_SetMusicVolume(ByVal volume As Single)
    cdaudio_c.sound_SetMusicVolume volume
End Sub

'--------------------------------------------------------------------------
' cdaudio.h (cdaudio_c). A std::shared_ptr<const WZ_TRACK> is a WZ_TRACK (Nothing = nullptr).

' std::string to_string(MusicGameMode mode)
Public Function MusicGameMode_to_string(ByVal mode As MusicGameMode) As String
    MusicGameMode_to_string = cdaudio_c.MusicGameMode_to_string(mode)
End Function

Public Function cdAudio_Open(ByVal user_musicdir As String) As Boolean
    cdAudio_Open = cdaudio_c.cdAudio_Open(user_musicdir)
End Function

Public Sub cdAudio_Close()
    cdaudio_c.cdAudio_Close
End Sub

Public Function cdAudio_PlayTrack(ByVal context As SONG_CONTEXT) As Boolean
    cdAudio_PlayTrack = cdaudio_c.cdAudio_PlayTrack(context)
End Function

Public Function cdAudio_PlaySpecificTrack(ByVal track As WZ_TRACK) As Boolean
    cdAudio_PlaySpecificTrack = cdaudio_c.cdAudio_PlaySpecificTrack(track)
End Function

Public Function cdAudio_CurrentSongContext() As SONG_CONTEXT
    cdAudio_CurrentSongContext = cdaudio_c.cdAudio_CurrentSongContext()
End Function

Public Sub cdAudio_SetGameMode(ByVal mode As MusicGameMode)
    cdaudio_c.cdAudio_SetGameMode mode
End Sub

Public Sub cdAudio_NextTrack()
    cdaudio_c.cdAudio_NextTrack
End Sub

Public Sub cdAudio_Stop()
    cdaudio_c.cdAudio_Stop
End Sub

Public Sub cdAudio_Pause()
    cdaudio_c.cdAudio_Pause
End Sub

Public Sub cdAudio_Resume()
    cdaudio_c.cdAudio_Resume
End Sub

Public Function cdAudio_GetCurrentTrack() As WZ_TRACK
    Set cdAudio_GetCurrentTrack = cdaudio_c.cdAudio_GetCurrentTrack()
End Function

Public Sub cdAudio_RegisterForEvents(ByVal musicEventSink As ICDAudioEventSink)
    cdaudio_c.cdAudio_RegisterForEvents musicEventSink
End Sub

'--------------------------------------------------------------------------
' playlist.h (playlist_c): what the game's music manager reads and sets

' const std::vector<std::shared_ptr<const WZ_TRACK>>& PlayList_GetFullTrackList(): its size and its tracks (0-based)
Public Function PlayList_GetFullTrackList_size() As Long
    PlayList_GetFullTrackList_size = playlist_c.PlayList_GetFullTrackList_size()
End Function

Public Function PlayList_GetFullTrackList_at(ByVal i As Long) As WZ_TRACK
    Set PlayList_GetFullTrackList_at = playlist_c.PlayList_GetFullTrackList_at(i)
End Function

Public Function PlayList_IsTrackEnabledForMusicMode(ByVal track As WZ_TRACK, ByVal mode As MusicGameMode) As Boolean
    PlayList_IsTrackEnabledForMusicMode = playlist_c.PlayList_IsTrackEnabledForMusicMode(track, mode)
End Function

Public Sub PlayList_SetTrackMusicMode(ByVal track As WZ_TRACK, ByVal mode As MusicGameMode, ByVal enabled As Boolean)
    playlist_c.PlayList_SetTrackMusicMode track, mode, enabled
End Sub

Public Function PlayList_GetCurrentMusicMode() As MusicGameMode
    PlayList_GetCurrentMusicMode = playlist_c.PlayList_GetCurrentMusicMode()
End Function

' (VB) C's rand() has one state per process, src/shared's cstdlib_vb one per project: the game's srand (main.cpp's
' srand((unsigned int)time(NULL))) seeds WzSound's too, whose one rand() call is PlayList_RandomizeCurrentSong's
' (the first in-game track; unseeded it would be the same track in every run)
Public Sub sound_srand(ByVal seed As Long)
    cstdlib_vb.srand seed
End Sub

'--------------------------------------------------------------------------
' audio_id.cpp

Public Function audio_GetIDFromStr(ByVal fileName As String) As INGAME_AUDIO
    audio_GetIDFromStr = audio_id_c.audio_GetIDFromStr(fileName)
End Function

'--------------------------------------------------------------------------
' (VB) For lib/sequence (WzSequence, cluster K): libvorbis's synthesis API (codec.h), over WzSound's stb_vorbis
' port (stb_vorbis_c's "libvorbis's synthesis API"). vorbis_info, vorbis_comment, vorbis_dsp_state and
' vorbis_block are handles (0 = a cleared struct); the _init functions write the new handle, the _clear ones
' free it and write 0. An ogg_packet is passed as its fields (packet: its bytes from index 0, as libogg's
' ogg_stream_packetout gives them; the array may be empty when bytes is 0). vorbis_synthesis_pcmout's
' float ***pcm is a copy: pcm(channel, i) = pcm[channel][i] for i below the count returned.

Public Sub vorbis_info_init(vi As Long)
    stb_vorbis_c.vorbis_info_init vi
End Sub

Public Sub vorbis_info_clear(vi As Long)
    stb_vorbis_c.vorbis_info_clear vi
End Sub

' vi->channels
Public Function vorbis_info_channels(ByVal vi As Long) As Long
    vorbis_info_channels = stb_vorbis_c.vorbis_info_channels(vi)
End Function

' vi->rate
Public Function vorbis_info_rate(ByVal vi As Long) As Long
    vorbis_info_rate = stb_vorbis_c.vorbis_info_rate(vi)
End Function

Public Sub vorbis_comment_init(vc As Long)
    stb_vorbis_c.vorbis_comment_init vc
End Sub

Public Sub vorbis_comment_clear(vc As Long)
    stb_vorbis_c.vorbis_comment_clear vc
End Sub

Public Function vorbis_synthesis_headerin(ByVal vi As Long, ByVal vc As Long, packet() As Byte, ByVal bytes As Long, _
        ByVal b_o_s As Long, ByVal e_o_s As Long, ByVal granulepos As Double, ByVal packetno As Double) As Long
    vorbis_synthesis_headerin = stb_vorbis_c.vorbis_synthesis_headerin(vi, vc, packet, bytes, b_o_s, e_o_s, granulepos, packetno)
End Function

Public Function vorbis_synthesis_init(vd As Long, ByVal vi As Long) As Long
    vorbis_synthesis_init = stb_vorbis_c.vorbis_synthesis_init(vd, vi)
End Function

Public Function vorbis_block_init(ByVal vd As Long, vb As Long) As Long
    vorbis_block_init = stb_vorbis_c.vorbis_block_init(vd, vb)
End Function

Public Function vorbis_synthesis(ByVal vb As Long, packet() As Byte, ByVal bytes As Long, ByVal b_o_s As Long, _
        ByVal e_o_s As Long, ByVal granulepos As Double, ByVal packetno As Double) As Long
    vorbis_synthesis = stb_vorbis_c.vorbis_synthesis(vb, packet, bytes, b_o_s, e_o_s, granulepos, packetno)
End Function

Public Function vorbis_synthesis_blockin(ByVal vd As Long, ByVal vb As Long) As Long
    vorbis_synthesis_blockin = stb_vorbis_c.vorbis_synthesis_blockin(vd, vb)
End Function

Public Function vorbis_synthesis_pcmout(ByVal vd As Long, pcm() As Single) As Long
    vorbis_synthesis_pcmout = stb_vorbis_c.vorbis_synthesis_pcmout(vd, pcm)
End Function

Public Function vorbis_synthesis_read(ByVal vd As Long, ByVal samples As Long) As Long
    vorbis_synthesis_read = stb_vorbis_c.vorbis_synthesis_read(vd, samples)
End Function

' vd->granulepos
Public Function vorbis_dsp_granulepos(ByVal vd As Long) As Double
    vorbis_dsp_granulepos = stb_vorbis_c.vorbis_dsp_granulepos(vd)
End Function

Public Function vorbis_block_clear(vb As Long) As Long
    vorbis_block_clear = stb_vorbis_c.vorbis_block_clear(vb)
End Function

Public Sub vorbis_dsp_clear(vd As Long)
    stb_vorbis_c.vorbis_dsp_clear vd
End Sub

'--------------------------------------------------------------------------
' (VB) For lib/sequence (cluster K): OpenAL (al_win) as sequence.cpp calls it, with al.h's names and
' semantics; a pointer to ALuint names is an array and the index of its first element. al.h's constants are
' not here: an enum member named like one of al_win's constants is ambiguous inside WzSound (measured), so
' the caller keeps al.h's values.

Public Sub alGenSources(ByVal n As Long, sources() As Long, Optional ByVal first As Long = 0)
    al_win.alGenSources n, sources, first
End Sub

Public Sub alDeleteSources(ByVal n As Long, sources() As Long, Optional ByVal first As Long = 0)
    al_win.alDeleteSources n, sources, first
End Sub

Public Sub alGenBuffers(ByVal n As Long, buffers() As Long, Optional ByVal first As Long = 0)
    al_win.alGenBuffers n, buffers, first
End Sub

Public Sub alDeleteBuffers(ByVal n As Long, buffers() As Long, Optional ByVal first As Long = 0)
    al_win.alDeleteBuffers n, buffers, first
End Sub

' data: the bytes from index 0 (16-bit samples in little-endian order for the 16-bit formats)
Public Sub alBufferData(ByVal buffer As Long, ByVal fmt_ As Long, data() As Byte, ByVal size As Long, ByVal freq As Long)
    al_win.alBufferData buffer, fmt_, data, size, freq
End Sub

Public Sub alSourceQueueBuffers(ByVal source As Long, ByVal nb As Long, buffers() As Long, Optional ByVal first As Long = 0)
    al_win.alSourceQueueBuffers source, nb, buffers, first
End Sub

Public Sub alSourceUnqueueBuffers(ByVal source As Long, ByVal nb As Long, buffers() As Long, Optional ByVal first As Long = 0)
    al_win.alSourceUnqueueBuffers source, nb, buffers, first
End Sub

Public Sub alGetSourcei(ByVal source As Long, ByVal param As Long, value As Long)
    al_win.alGetSourcei source, param, value
End Sub

Public Sub alGetBufferi(ByVal buffer As Long, ByVal param As Long, value As Long)
    al_win.alGetBufferi buffer, param, value
End Sub

Public Sub alSourcef(ByVal source As Long, ByVal param As Long, ByVal value As Single)
    al_win.alSourcef source, param, value
End Sub

Public Sub alSourcePlay(ByVal source As Long)
    al_win.alSourcePlay source
End Sub

' openal_error.h: the macro sound_GetError() (location_description: the call's "file:line", the macro's AT_MACRO)
Public Function sound_GetError(ByVal location_description As String) As Long
    sound_GetError = openal_error_c.sound_GetError(location_description)
End Function

' (VB) al_win mixes only when polled (OpenAL Soft: in a thread of its own); lib/sequence polls the current
' context while a video plays (the full screen videos run outside the loops that call sound_Update)
Public Sub al_vbProcessCurrent()
    al_win.al_vbProcessCurrent
End Sub
