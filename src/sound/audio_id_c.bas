Attribute VB_Name = "audio_id_c"
' lib/sound/audio_id.cpp: the table of the sound effects' file names
' (asAudioID) and audio_GetIDFromStr. The enum INGAME_AUDIO (audio_id.h)
' is SoundApi's. The table is filled on the first call (VB has no static
' initialisers).
Option Explicit

Private Type AUDIO_ID_MAP
    ID As Long                  ' INGAME_AUDIO
    fileName As String
End Type

Private asAudioID(0 To ID_SOUND_NEXT - 1) As AUDIO_ID_MAP
Private asAudioIDCount As Long

Private Sub AddID(ByVal ID As Long, ByVal fileName As String)
    asAudioID(asAudioIDCount).ID = ID
    asAudioID(asAudioIDCount).fileName = fileName
    asAudioIDCount = asAudioIDCount + 1
End Sub

Private Sub InitAudioID()
    ' Beeps

    AddID ID_SOUND_WINDOWCLOSE, "beep1.ogg"
    AddID ID_SOUND_WINDOWOPEN, "beep2.ogg"
    AddID ID_SOUND_BUTTON_CLICK_3, "beep3.ogg"
    AddID ID_SOUND_SELECT, "beep4.ogg"
    AddID ID_SOUND_BUTTON_CLICK_5, "beep5.ogg"
    AddID FE_AUDIO_MESSAGEEND, "beep6.ogg"
    AddID ID_SOUND_ZOOM_ON_RADAR, "beep7.ogg"
    AddID ID_SOUND_BUILD_FAIL, "beep8.ogg"
    AddID ID_SOUND_MESSAGEEND, "beep9.ogg"
    AddID ID_SOUND_GAME_SHUTDOWN, "gmeshtdn.ogg"

    ' Design Sequence

    AddID ID_SOUND_TURRET_SELECTED, "pcv331.ogg"
    AddID ID_SOUND_BODY_SELECTED, "pcv332.ogg"
    AddID ID_SOUND_PROPULSION_SELECTED, "pcv333.ogg"
    AddID ID_SOUND_DESIGN_COMPLETED, "pcv334.ogg"

    ' Structures

    AddID ID_SOUND_CONSTRUCTION_STARTED, "pcv335.ogg"
    AddID ID_SOUND_STRUCTURE_COMPLETED, "pcv336.ogg"
    AddID ID_SOUND_STRUCTURE_UNDER_ATTACK, "pcv337.ogg"
    AddID ID_SOUND_STRUCTURE_REPAIR_IN_PROGRESS, "pcv339.ogg"
    AddID ID_SOUND_STRUCTURE_DEMOLISHED, "pcv340.ogg"

    ' Power

    AddID ID_SOUND_POWER_GENERATOR_UNDER_ATTACK, "pcv341.ogg"
    AddID ID_SOUND_POWER_GENERATOR_DESTROYED, "pcv342.ogg"
    AddID ID_SOUND_POWER_LOW, "pcv343.ogg"
    AddID ID_SOUND_RESOURCE_HERE, "pcv344.ogg"
    AddID ID_SOUND_DERRICK_UNDER_ATTACK, "pcv345.ogg"
    AddID ID_SOUND_DERRICK_DESTROYED, "pcv346.ogg"
    AddID ID_SOUND_RESOURCE_DEPLETED, "pcv347.ogg"
    AddID ID_SOUND_POWER_TRANSFER_IN_PROGRESS, "pcv348.ogg"
    AddID ID_SOUND_POWER_GENERATOR_REQUIRED, "pcv349.ogg"

    ' Research

    AddID ID_SOUND_RESEARCH_FACILITY_REQUIRED, "pcv350.ogg"
    AddID ID_SOUND_ARTIFACT, "pcv351.ogg"
    AddID ID_SOUND_ARTIFACT_RECOVERED, "pcv352.ogg"
    AddID ID_SOUND_NEW_RESEARCH_PROJ_AVAILABLE, "pcv353.ogg"
    AddID ID_SOUND_NEW_STRUCTURE_AVAILABLE, "pcv354.ogg"
    AddID ID_SOUND_NEW_COMPONENT_AVAILABLE, "pcv355.ogg"
    AddID ID_SOUND_NEW_CYBORG_AVAILABLE, "pcv356.ogg"
    AddID ID_SOUND_RESEARCH_COMPLETED, "pcv357.ogg"
    AddID ID_SOUND_MAJOR_RESEARCH, "pcv358.ogg"
    AddID ID_SOUND_STRUCTURE_RESEARCH_COMPLETED, "pcv359.ogg"
    AddID ID_SOUND_POWER_RESEARCH_COMPLETED, "pcv360.ogg"
    AddID ID_SOUND_COMPUTER_RESEARCH_COMPLETED, "pcv361.ogg"
    AddID ID_SOUND_VEHICLE_RESEARCH_COMPLETED, "pcv362.ogg"
    AddID ID_SOUND_SYSTEMS_RESEARCH_COMPLETED, "pcv363.ogg"
    AddID ID_SOUND_WEAPON_RESEARCH_COMPLETED, "pcv364.ogg"
    AddID ID_SOUND_CYBORG_RESEARCH_COMPLETED, "pcv365.ogg"

    ' Production

    AddID ID_SOUND_PRODUCTION_STARTED, "pcv366.ogg"
    AddID ID_SOUND_DROID_COMPLETED, "pcv367.ogg"
    AddID ID_SOUND_PRODUCTION_PAUSED, "pcv368.ogg"
    AddID ID_SOUND_PRODUCTION_CANCELLED, "pcv369.ogg"
    AddID ID_SOUND_DELIVERY_POINT_ASSIGNED, "pcv370.ogg"
    AddID ID_SOUND_DELIVERY_POINT_ASSIGNED_TO, "pcv371.ogg"

    ' Repair

    AddID ID_SOUND_UNIT_REPAIRED, "pcv372.ogg"

    ' Detection

    AddID ID_SOUND_SCAVENGERS_DETECTED, "pcv373.ogg"
    AddID ID_SOUND_SCAVENGER_BASE_DETECTED, "pcv374.ogg"
    AddID ID_SOUND_SCAVENGER_OUTPOST_DETECTED, "pcv375.ogg"
    AddID ID_SOUND_POWER_RESOURCE, "pcv376.ogg"
    AddID ID_SOUND_ARTEFACT_DISC, "pcv377.ogg"
    AddID ID_SOUND_ENEMY_UNIT_DETECTED, "pcv378.ogg"
    AddID ID_SOUND_ENEMY_BASE_DETECTED, "pcv379.ogg"
    AddID ID_SOUND_ALLY_DETECTED, "pcv380.ogg"
    AddID ID_SOUND_ENEMY_TRANSPORT_DETECTED, "pcv381.ogg"
    AddID ID_SOUND_ENEMY_LZ_DETECTED, "pcv382.ogg"
    AddID ID_SOUND_FRIENDLY_LZ_DETECTED, "pcv383.ogg"
    AddID ID_SOUND_NEXUS_TOWER_DETECTED, "pcv384.ogg"
    AddID ID_SOUND_NEXUS_TURRET_DETECTED, "pcv385.ogg"
    AddID ID_SOUND_NEXUS_UNIT_DETECTED, "pcv386.ogg"
    AddID ID_SOUND_ENEMY_BATTERY_DETECTED, "pcv387.ogg"
    AddID ID_SOUND_ENEMY_VTOLS_DETECTED, "pcv388.ogg"

    ' Status

    AddID ID_SOUND_SCAVENGER_BASE, "pcv389.ogg"
    AddID ID_SOUND_SCAVENGER_OUTPOST, "pcv390.ogg"
    AddID ID_SOUND_SCAVENGER_OUTPOST_ERADICATED, "pcv391.ogg"
    AddID ID_SOUND_SCAVENGER_BASE_ERADICATED, "pcv392.ogg"
    AddID ID_SOUND_ENEMY_BASE, "pcv393.ogg"
    AddID ID_SOUND_ENEMY_BASE_ERADICATED, "pcv394.ogg"
    AddID ID_SOUND_INCOMING_ENEMY_TRANSPORT, "pcv395.ogg"
    AddID ID_SOUND_ENEMY_LZ, "pcv396.ogg"
    AddID ID_SOUND_LZ1, "pcv397.ogg"
    AddID ID_SOUND_LZ2, "pcv398.ogg"

    ' Combat

    AddID ID_SOUND_UNIT_UNDER_ATTACK, "pcv399.ogg"
    AddID ID_SOUND_UNIT_DESTROYED, "pcv400.ogg"
    AddID ID_SOUND_UNIT_RETREATING, "pcv401.ogg"
    AddID ID_SOUND_UNIT_RETURNING_FOR_REPAIR, "pcv402.ogg"

    ' Artillery Batteries
    AddID ID_SOUND_ASSIGNED_TO_SENSOR, "pcv403.ogg"
    AddID ID_SOUND_SENSOR_LOCKED_ON, "pcv404.ogg"
    AddID ID_SOUND_ASSIGNED_TO_COUNTER_RADAR, "pcv405.ogg"
    AddID ID_SOUND_ENEMY_BATTERY_LOCATED, "pcv406.ogg"
    AddID ID_SOUND_BATTERY_FIRING_COUNTER_ATTACK, "pcv407.ogg"

    ' Vtols

    AddID ID_SOUND_INTERCEPTORS_LAUNCHED, "pcv408.ogg"
    AddID ID_SOUND_REARMING, "pcv409.ogg"
    AddID ID_SOUND_VTOLS_ENGAGING, "pcv410.ogg"
    AddID ID_SOUND_ASSIGNED, "pcv411.ogg"
    AddID ID_SOUND_INTERCEPTORS_ASSIGNED, "pcv412.ogg"

    ' Command Console
    AddID ID_SOUND_COMMAND_CONSOLE_ACTIVATED, "pcv413.ogg"
    AddID ID_SOUND_SHORT_RANGE, "pcv414.ogg"
    AddID ID_SOUND_LONG_RANGE, "pcv415.ogg"
    AddID ID_SOUND_OPTIMUM_RANGE, "pcv416.ogg"
    AddID ID_SOUND_RETREAT_AT_MEDIUM_DAMAGE, "pcv417.ogg"
    AddID ID_SOUND_RETREAT_AT_HEAVY_DAMAGE, "pcv418.ogg"
    AddID ID_SOUND_NO_RETREAT, "pcv419.ogg"
    AddID ID_SOUND_FIRE_AT_WILL, "pcv420.ogg"
    AddID ID_SOUND_RETURN_FIRE, "pcv421.ogg"
    AddID ID_SOUND_CEASEFIRE, "pcv422.ogg"
    AddID ID_SOUND_HOLD_POSITION, "pcv423.ogg"
    AddID ID_SOUND_GUARD, "pcv424.ogg"
    AddID ID_SOUND_PURSUE, "pcv425.ogg"
    AddID ID_SOUND_PATROL, "pcv426.ogg"
    AddID ID_SOUND_RETURN_TO_LZ, "pcv427.ogg"
    AddID ID_SOUND_RECYCLING, "pcv428.ogg"
    AddID ID_SOUND_SCATTER, "pcv429.ogg"

    ' Tutorial Stuff

    AddID ID_SOUND_NOT_POSSIBLE_TRY_AGAIN, "pcv430.ogg"
    AddID ID_SOUND_NO, "pcv431.ogg"
    AddID ID_SOUND_THAT_IS_INCORRECT, "pcv432.ogg"
    AddID ID_SOUND_WELL_DONE, "pcv433.ogg"
    AddID ID_SOUND_EXCELLENT, "pcv434.ogg"

    ' Group and Commander Assignment

    AddID ID_SOUND_ASSIGNED_TO_COMMANDER, "pcv435.ogg"
    AddID ID_SOUND_GROUP_REPORTING, "pcv436.ogg"
    AddID ID_SOUND_COMMANDER_REPORTING, "pcv437.ogg"

    ' Routing

    AddID ID_SOUND_ROUTE_OBSTRUCTED, "pcv438.ogg"
    AddID ID_SOUND_NO_ROUTE_AVAILABLE, "pcv439.ogg"

    ' Transports and LZS
    AddID ID_SOUND_REINFORCEMENTS_AVAILABLE, "pcv440.ogg"
    AddID ID_SOUND_REINFORCEMENTS_IN_TRANSIT, "pcv441.ogg"
    AddID ID_SOUND_TRANSPORT_LANDING, "pcv442.ogg"
    AddID ID_SOUND_TRANSPORT_UNDER_ATTACK, "pcv443.ogg"
    AddID ID_SOUND_TRANSPORT_REPAIRING, "pcv444.ogg"
    AddID ID_SOUND_LZ_COMPROMISED, "pcv445.ogg"
    AddID ID_SOUND_LZ_CLEAR, "lz-clear.ogg"
    AddID ID_SOUND_TRANSPORT_RETURNING_TO_BASE, "pcv446.ogg"
    AddID ID_SOUND_TRANSPORT_UNABLE_TO_LAND, "pcv447.ogg"

    ' Mission Messages

    AddID ID_SOUND_MISSION_OBJECTIVE, "pcv448.ogg"
    AddID ID_SOUND_MISSION_UPDATE, "pcv449.ogg"
    AddID ID_SOUND_WARZONE_PAUSED, "pcv450.ogg"
    AddID ID_SOUND_WARZONE_ACTIVE, "pcv451.ogg"
    AddID ID_SOUND_MISSION_RESULTS, "pcv452.ogg"
    AddID ID_SOUND_RESEARCH_STOLEN, "pcv453.ogg"
    AddID ID_SOUND_TECHNOLOGY_TAKEN, "pcv454.ogg"
    AddID ID_SOUND_INCOMING_TRANSMISSION, "pcv455.ogg"
    AddID ID_SOUND_INCOMING_INTELLIGENCE_REPORT, "pcv456.ogg"
    AddID ID_SOUND_MISSION_FAILED, "pcv458.ogg"
    AddID ID_SOUND_MISSION_SUCCESSFUL, "pcv459.ogg"
    AddID ID_SOUND_OBJECTIVE_ACCOMPLISHED, "pcv460.ogg"
    AddID ID_SOUND_OBJECTIVE_FAILED, "pcv461.ogg"
    AddID ID_SOUND_MISSION_TIMER_ACTIVATED, "pcv462.ogg"
    AddID ID_SOUND_10_MINUTES_REMAINING, "pcv463.ogg"
    AddID ID_SOUND_5_MINUTES_REMAINING, "pcv464.ogg"
    AddID ID_SOUND_3_MINUTES_REMAINING, "pcv465.ogg"
    AddID ID_SOUND_2_MINUTES_REMAINING, "pcv466.ogg"
    AddID ID_SOUND_1_MINUTE_REMAINING, "pcv467.ogg"
    AddID ID_SOUND_UNIT_CAPTURED, "pcv468.ogg"
    AddID ID_SOUND_SYSTEM_FAILURE_IMMINENT, "pcv469.ogg"
    AddID ID_SOUND_YOU_ARE_DEFEATED, "pcv470.ogg"
    AddID ID_SOUND_MISSILE_CODES_DECIPHERED, "pcv471.ogg"
    AddID ID_SOUND_1ST_MISSILE_CODES_DECIPHERED, "pcv472.ogg"
    AddID ID_SOUND_2ND_MISSILE_CODES_DECIPHERED, "pcv473.ogg"
    AddID ID_SOUND_3RD_MISSILE_CODES_DECIPHERED, "pcv474.ogg"
    AddID ID_SOUND_MISSILE_CODES_CRACKED, "pcv475.ogg"
    AddID ID_SOUND_ENTERING_WARZONE, "pcv476.ogg"
    AddID ID_ALLIANCE_ACC, "pcv477.ogg"
    AddID ID_ALLIANCE_BRO, "pcv478.ogg"
    AddID ID_ALLIANCE_OFF, "pcv479.ogg"
    AddID ID_CLAN_ENTER, "pcv480.ogg"
    AddID ID_CLAN_EXIT, "pcv481.ogg"
    AddID ID_GIFT, "pcv482.ogg"
    AddID ID_POWER_TRANSMIT, "power-transferred.ogg"
    AddID ID_SENSOR_DOWNLOAD, "pcv484.ogg"
    AddID ID_TECHNOLOGY_TRANSFER, "pcv485.ogg"
    AddID ID_UNITS_TRANSFER, "pcv486.ogg"

    ' Group and Commander Voices - Male
    AddID ID_SOUND_GROUP, "group.ogg"
    AddID ID_SOUND_GROUP_0, "0.ogg"
    AddID ID_SOUND_GROUP_1, "1.ogg"
    AddID ID_SOUND_GROUP_2, "2.ogg"
    AddID ID_SOUND_GROUP_3, "3.ogg"
    AddID ID_SOUND_GROUP_4, "4.ogg"
    AddID ID_SOUND_GROUP_5, "5.ogg"
    AddID ID_SOUND_GROUP_6, "6.ogg"
    AddID ID_SOUND_GROUP_7, "7.ogg"
    AddID ID_SOUND_GROUP_8, "8.ogg"
    AddID ID_SOUND_GROUP_9, "9.ogg"
    AddID ID_SOUND_REPORTING, "reprting.ogg"
    AddID ID_SOUND_COMMANDER, "commnder.ogg"
    AddID ID_SOUND_COM_SCAVS_DETECTED, "com021.ogg"
    AddID ID_SOUND_COM_SCAV_BASE_DETECTED, "com022.ogg"
    AddID ID_SOUND_COM_SCAV_OUTPOST_DETECTED, "com023.ogg"
    AddID ID_SOUND_COM_RESOURCE_DETECTED, "com024.ogg"
    AddID ID_SOUND_COM_ARTEFACT_DETECTED, "com025.ogg"
    AddID ID_SOUND_COM_ENEMY_DETECTED, "com026.ogg"
    AddID ID_SOUND_COM_ENEMY_BASE_DETECTED, "com027.ogg"
    AddID ID_SOUND_COM_ALLY_DETECTED, "com028.ogg"
    AddID ID_SOUND_COM_ENEMY_TRANSPORT_DETECTED, "com029.ogg"
    AddID ID_SOUND_COM_ENEMY_LZ_DETECTED, "com030.ogg"
    AddID ID_SOUND_COM_FRIENDLY_LZ_DETECTED, "com031.ogg"
    AddID ID_SOUND_COM_NEXUS_TOWER_DETECTED, "com032.ogg"
    AddID ID_SOUND_COM_NEXUS_TURRET_DETECTED, "com033.ogg"
    AddID ID_SOUND_COM_NEXUS_DETECTED, "com034.ogg"
    AddID ID_SOUND_COM_ENEMY_BATTERY_DETECTED, "com035.ogg"
    AddID ID_SOUND_COM_ENEMY_VTOLS_DETECTED, "com036.ogg"
    AddID ID_SOUND_COM_ROUTE_OBSTRUCTED, "com037.ogg"
    AddID ID_SOUND_COM_NO_ROUTE_AVAILABLE, "com038.ogg"
    AddID ID_SOUND_COM_UNABLE_TO_COMPLY, "com039.ogg"
    AddID ID_SOUND_COM_RETURNING_FOR_REPAIR, "com040.ogg"
    AddID ID_SOUND_COM_HEADING_FOR_RALLY_POINT, "com041.ogg"

    ' Radio Clicks
    AddID ID_SOUND_RADIOCLICK_1, "radclik1.ogg"
    AddID ID_SOUND_RADIOCLICK_2, "radclik2.ogg"
    AddID ID_SOUND_RADIOCLICK_3, "radclik3.ogg"
    AddID ID_SOUND_RADIOCLICK_4, "radclik4.ogg"
    AddID ID_SOUND_RADIOCLICK_5, "radclik5.ogg"
    AddID ID_SOUND_RADIOCLICK_6, "radclik6.ogg"

    ' Transport Pilots
    AddID ID_SOUND_APPROACHING_LZ, "t-aprolz.ogg"
    AddID ID_SOUND_ALRIGHT_BOYS, "t-arboys.ogg"
    AddID ID_SOUND_GREEN_LIGHT_IN_5, "t-grnli5.ogg"
    AddID ID_SOUND_GREEN_LIGHT_IN_4, "t-grnli4.ogg"
    AddID ID_SOUND_GREEN_LIGHT_IN_3, "t-grnli3.ogg"
    AddID ID_SOUND_GREEN_LIGHT_IN_2, "t-grnli2.ogg"
    AddID ID_SOUND_GO_GO_GO, "t-gogogo.ogg"
    AddID ID_SOUND_PREPARE_FOR_DUST_OFF, "t-dustof.ogg"

    ' VTol Pilots

    ' Ver-1
    AddID ID_SOUND_ENEMY_LOCATED1, "v-eloc1.ogg"
    AddID ID_SOUND_ON_OUR_WAY1, "v-onway1.ogg"
    AddID ID_SOUND_RETURNING_TO_BASE1, "v-retba1.ogg"
    AddID ID_SOUND_LOCKED_ON1, "v-locon1.ogg"
    AddID ID_SOUND_COMMENCING_ATTACK_RUN1, "v-atkrn1.ogg"
    AddID ID_SOUND_ABORTING_ATTACK_RUN1, "v-abtrn1.ogg"

    ' Ver-2
    AddID ID_SOUND_ENEMY_LOCATED2, "v-eloc2.ogg"
    AddID ID_SOUND_ON_OUR_WAY2, "v-onway2.ogg"
    AddID ID_SOUND_RETURNING_TO_BASE2, "v-retba2.ogg"
    AddID ID_SOUND_LOCKED_ON2, "v-locon2.ogg"
    AddID ID_SOUND_COMMENCING_ATTACK_RUN2, "v-atkrn2.ogg"
    AddID ID_SOUND_ABORTING_ATTACK_RUN2, "v-abtrn2.ogg"

    ' Ver-3
    AddID ID_SOUND_ENEMY_LOCATED3, "v-eloc3.ogg"
    AddID ID_SOUND_ON_OUR_WAY3, "v-onway3.ogg"
    AddID ID_SOUND_RETURNING_TO_BASE3, "v-retba3.ogg"
    AddID ID_SOUND_LOCKED_ON3, "v-locon3.ogg"
    AddID ID_SOUND_COMMENCING_ATTACK_RUN3, "v-atkrn3.ogg"
    AddID ID_SOUND_ABORTING_ATTACK_RUN3, "v-abtrn3.ogg"

    ' The Collective
    AddID ID_SOUND_COLL_CLEANSE_AND_DESTROY, "col011a.ogg"
    AddID ID_SOUND_COLL_DESTROYING_BIOLOGICALS, "col012a.ogg"
    AddID ID_SOUND_COLL_ATTACK, "col013a.ogg"
    AddID ID_SOUND_COLL_FIRE, "col014a.ogg"
    AddID ID_SOUND_COLL_ENEMY_DETECTED, "col015a.ogg"
    AddID ID_SOUND_COLL_ENGAGING, "col016a.ogg"
    AddID ID_SOUND_COLL_STARTING_ATTACK_RUN, "col017a.ogg"
    AddID ID_SOUND_COLL_DIE, "col018a.ogg"
    AddID ID_SOUND_COLL_INTERCEPT_AND_DESTROY, "col019a.ogg"
    AddID ID_SOUND_COLL_ENEMY_DESTROYED, "col020a.ogg"

    ' SFX
    ' Weapon Sounds
    AddID ID_SOUND_ROCKET, "rocket.ogg"
    AddID ID_SOUND_ROTARY_LASER, "rotlsr.ogg"
    AddID ID_SOUND_GAUSSGUN, "gaussgun.ogg"
    AddID ID_SOUND_LARGE_CANNON, "lrgcan.ogg"
    AddID ID_SOUND_SMALL_CANNON, "smlcan.ogg"
    AddID ID_SOUND_MEDIUM_CANNON, "medcan.ogg"
    AddID ID_SOUND_FLAME_THROWER, "flmthrow.ogg"
    AddID ID_SOUND_PULSE_LASER, "plslsr.ogg"
    AddID ID_SOUND_BEAM_LASER, "bemlsr.ogg"
    AddID ID_SOUND_MORTAR, "mortar.ogg"
    AddID ID_SOUND_HOWITZ_FLIGHT, "hwtzflgt.ogg"
    AddID ID_SOUND_BABA_MG_1, "mgbar1.ogg"
    AddID ID_SOUND_BABA_MG_2, "mgbar2.ogg"
    AddID ID_SOUND_BABA_MG_3, "mgbar3.ogg"
    AddID ID_SOUND_BABA_MG_HEAVY, "mgheavy.ogg"
    AddID ID_SOUND_BABA_MG_TOWER, "mgtower.ogg"
    AddID ID_SOUND_SPLASH, "splash.ogg"
    AddID ID_SOUND_ASSAULT_MG, "asltmg.ogg"
    AddID ID_SOUND_RAPID_CANNON, "rapdcan.ogg"
    AddID ID_SOUND_HIVEL_CANNON, "hivelcan.ogg"
    AddID ID_SOUND_NEXUS_TOWER, "nxstower.ogg"

    ' Construction sounds
    AddID ID_SOUND_WELD_1, "weld-1.ogg"
    AddID ID_SOUND_WELD_2, "weld-2.ogg"
    AddID ID_SOUND_CONSTRUCTION_START, "bldstart.ogg"
    AddID ID_SOUND_CONSTRUCTION_LOOP, "bldloop.ogg"
    AddID ID_SOUND_CONSTRUCTION_1, "build1.ogg"
    AddID ID_SOUND_CONSTRUCTION_2, "build2.ogg"
    AddID ID_SOUND_CONSTRUCTION_3, "build3.ogg"
    AddID ID_SOUND_CONSTRUCTION_4, "build4.ogg"

    ' Explosions

    AddID ID_SOUND_EXPLOSION_SMALL, "smlexpl.ogg"
    AddID ID_SOUND_EXPLOSION_LASER, "lsrexpl.ogg"
    AddID ID_SOUND_EXPLOSION, "lrgexpl.ogg"
    AddID ID_SOUND_EXPLOSION_ANTITANK, "atnkexpl.ogg"
    AddID ID_SOUND_RICOCHET_1, "richet1.ogg"
    AddID ID_SOUND_RICOCHET_2, "richet2.ogg"
    AddID ID_SOUND_RICOCHET_3, "richet3.ogg"
    AddID ID_SOUND_BARB_SQUISH, "squish.ogg"
    AddID ID_SOUND_BUILDING_FALL, "bldfall.ogg"
    AddID ID_SOUND_NEXUS_EXPLOSION, "nxsexpld.ogg"

    ' Droid Engine Noises
    AddID ID_SOUND_CONSTRUCTOR_MOVEOFF, "con-move-off.ogg"
    AddID ID_SOUND_CONSTRUCTOR_MOVE, "con-move.ogg"
    AddID ID_SOUND_CONSTRUCTOR_SHUTDOWN, "con-shut-down.ogg"

    ' Transports
    AddID ID_SOUND_BLIMP_FLIGHT, "tflight.ogg"
    AddID ID_SOUND_BLIMP_IDLE, "thover.ogg"
    AddID ID_SOUND_BLIMP_LAND, "tland.ogg"
    AddID ID_SOUND_BLIMP_TAKE_OFF, "tstart.ogg"

    ' Vtols
    AddID ID_SOUND_VTOL_LAND, "vtolland.ogg"
    AddID ID_SOUND_VTOL_OFF, "vtoloff.ogg"
    AddID ID_SOUND_VTOL_MOVE, "vtol-move.ogg"

    ' Treads
    AddID ID_SOUND_TREAD, "tread.ogg"

    ' Hovers
    AddID ID_SOUND_HOVER_MOVE, "hovmove.ogg"
    AddID ID_SOUND_HOVER_START, "hovstart.ogg"
    AddID ID_SOUND_HOVER_STOP, "hovstop.ogg"

    ' Cyborgs
    AddID ID_SOUND_CYBORG_MOVE, "cyber-move.ogg"

    ' Buildings

    AddID ID_SOUND_OIL_PUMP_2, "oilpump.ogg"
    AddID ID_SOUND_POWER_HUM, "powerhum.ogg"
    AddID ID_SOUND_POWER_SPARK, "powerspk.ogg"
    AddID ID_SOUND_STEAM, "steam.ogg"
    AddID ID_SOUND_ECM_TOWER, "ecmtower.ogg"
    AddID ID_SOUND_FIRE_ROAR, "freroar.ogg"

    ' Misc
    AddID ID_SOUND_HELP, "help.ogg"
    AddID ID_SOUND_BARB_SCREAM, "scream.ogg"
    AddID ID_SOUND_BARB_SCREAM2, "scream2.ogg"
    AddID ID_SOUND_BARB_SCREAM3, "scream3.ogg"
    AddID ID_SOUND_OF_SILENCE, "silence.ogg"
    AddID ID_SOUND_SHIELD_HIT, "shield-hit.ogg"

    ' Extra
    AddID ID_SOUND_LANDING_ZONE, "lndgzne.ogg"
    AddID ID_SOUND_SATELLITE_UPLINK, "pcv652.ogg"
    AddID ID_SOUND_NASDA_CENTRAL, "pcv653.ogg"
    AddID ID_SOUND_NUCLEAR_REACTOR, "pcv654.ogg"
    AddID ID_SOUND_SAM_SITE, "pcv655.ogg"
    AddID ID_SOUND_MISSILE_SILO, "pcv656.ogg"
    AddID ID_SOUND_MISSILE_NME_DETECTED, "nmedeted.ogg"
    AddID ID_SOUND_STRUCTURE_CAPTURED, "pcv611.ogg"
    AddID ID_SOUND_CIVILIAN_RESCUED, "pcv612.ogg"
    AddID ID_SOUND_CIVILIANS_RESCUED, "pcv613.ogg"
    AddID ID_SOUND_UNITS_RESCUED, "pcv615.ogg"
    AddID ID_SOUND_GROUP_RESCUED, "pcv616.ogg"
    AddID ID_SOUND_GROUP_CAPTURED, "pcv618.ogg"
    AddID ID_SOUND_OBJECTIVE_CAPTURED, "pcv621.ogg"
    AddID ID_SOUND_OBJECTIVE_DESTROYED, "pcv622.ogg"
    AddID ID_SOUND_STRUCTURE_INFECTED, "pcv623.ogg"
    AddID ID_SOUND_GROUP_INFECTED, "pcv625.ogg"
    AddID ID_SOUND_OUT_OF_TIME, "pcv629.ogg"
    AddID ID_SOUND_ENEMY_ESCAPED, "pcv631.ogg"
    AddID ID_SOUND_ENEMY_ESCAPING, "pcv632.ogg"
    AddID ID_SOUND_ENEMY_TRANSPORT_LANDING, "pcv633.ogg"
    AddID ID_SOUND_TEAM_ALPHA_ERADICATED, "pcv635.ogg"
    AddID ID_SOUND_TEAM_BETA_ERADICATED, "pcv636.ogg"
    AddID ID_SOUND_TEAM_GAMMA_ERADICATED, "pcv637.ogg"
    AddID ID_SOUND_TEAM_ALPHA_RESCUED, "pcv638.ogg"
    AddID ID_SOUND_TEAM_BETA_RESCUED, "pcv639.ogg"
    AddID ID_SOUND_TEAM_GAMMA_RESCUED, "pcv640.ogg"
    AddID ID_SOUND_LASER_SATELLITE_FIRING, "pcv650.ogg"
    AddID ID_SOUND_INCOMING_LASER_SAT_STRIKE, "pcv657.ogg"

    ' Nexus
    AddID ID_SOUND_NEXUS_DEFENCES_ABSORBED, "defabsrd.ogg"
    AddID ID_SOUND_NEXUS_DEFENCES_NEUTRALISED, "defnut.ogg"
    AddID ID_SOUND_NEXUS_LAUGH1, "laugh1.ogg"
    AddID ID_SOUND_NEXUS_LAUGH2, "laugh2.ogg"
    AddID ID_SOUND_NEXUS_LAUGH3, "laugh3.ogg"
    AddID ID_SOUND_NEXUS_PRODUCTION_COMPLETED, "pordcomp.ogg"
    AddID ID_SOUND_NEXUS_RESEARCH_ABSORBED, "resabsrd.ogg"
    AddID ID_SOUND_NEXUS_STRUCTURE_ABSORBED, "strutabs.ogg"
    AddID ID_SOUND_NEXUS_STRUCTURE_NEUTRALISED, "strutnut.ogg"
    AddID ID_SOUND_NEXUS_SYNAPTIC_LINK, "synplnk.ogg"
    AddID ID_SOUND_NEXUS_UNIT_ABSORBED, "untabsrd.ogg"
    AddID ID_SOUND_NEXUS_UNIT_NEUTRALISED, "untnut.ogg"

    ' extra multiplayer sounds
    AddID ID_SOUND_CYBORG_GROUND, "cybgrnd.ogg"
    AddID ID_SOUND_CYBORG_HEAVY, "hvcybmov.ogg"
    AddID ID_SOUND_EMP, "emp.ogg"
    AddID ID_SOUND_LASER_HEAVY, "hevlsr.ogg"
    AddID ID_SOUND_PLASMA_FLAMER, "plasflm.ogg"
    AddID ID_SOUND_PARTICLE_GUN, "partgun.ogg"
    AddID ID_SOUND_UPLINK, "uplink.ogg"
    AddID ID_SOUND_LAS_SAT_COUNTDOWN, "lasstrk.ogg"
    AddID ID_SOUND_BEACON, "beacon.ogg"
End Sub

'**************************************************************************

Public Function audio_GetIDFromStr(ByVal fileName As String) As Long
    Dim i As Long
    If asAudioIDCount = 0 Then InitAudioID

    For i = 0 To ID_SOUND_NEXT - 1
        If fileName = asAudioID(i).fileName Then
            If Not (i = asAudioID(i).ID) Then ASSERT_FAILURE "i == asAudioID[i].ID", "lib/sound/audio_id.cpp:511", "audio_GetIDFromStr", "audioID_GetIDFromStr: %s stored IDs don't match", fileName

            audio_GetIDFromStr = asAudioID(i).ID
            Exit Function
        End If
    Next

    audio_GetIDFromStr = NO_SOUND
End Function
