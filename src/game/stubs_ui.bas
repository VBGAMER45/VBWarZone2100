Attribute VB_Name = "stubs_ui"
' Clusters F and G's merge: what the widgets, the front end and the in-game interface call that
' no module provides yet (load and save, the campaign, the in-game multiplayer menu, videos, the
' network lobby, ...). Stand-ins keep the C++ name and signature; a few are real ports of a function
' of a module not ported yet, and say where they belong.
Option Explicit

Private Const MAX_CONNECTED_PLAYERS As Long = MAX_PLAYER_SLOTS + MAX_SPECTATOR_SLOTS


' src/transporter.h: #define IDTRANS_LAUNCH 9010 //The Transporter Launch button








' src/loadsave.cpp: extern bool lastSaveMP; (the last save was a skirmish/multiplayer one)

' enum class MinimizeOnFocusLossBehavior
Public Const MinimizeOnFocusLossBehavior_Auto As Long = -1

Public Const MinimizeOnFocusLossBehavior_Off As Long = 0

Public Const MinimizeOnFocusLossBehavior_On_Fullscreen As Long = 1

' enum class KeyMappingSlot



' enum class KeyMappingType



' enum class KeyMappingInputSource


' enum class KeyAction



' netplay.h: #define password_string_size 64 (netplay_c's copy is Private)
Public Const password_string_size As Long = 64

' multiplay.h: enum TechLevel
Public Const TECH_1 As Long = 1

Public Const TECH_2 As Long = 2

Public Const TECH_3 As Long = 3

Public Const TECH_4 As Long = 4

' (stand-in state) NETgetLobbyHostJoinOptions / NETgetLobbyHostDirectJoinOption
Private hostJoinOpts_set As Boolean

Private hostJoinOpts_botProt As Long

Private hostJoinOpts_proxyIPs As Long

Private hostJoinOpts_hostingIPs As Long













' src/display.cpp: startDeliveryPosition(FLAG_POSITION *psFlag) for a flag that is not in aFlagPos (the debug menu's
' local FLAG_POSITION) (S: nothing; display_c's takes an aFlagPos handle)
Public Sub startDeliveryPosition_flag(psFlag As FLAG_POSITION)
End Sub

' src/multistruct.cpp: bool SendBuildFinished(STRUCTURE *psStruct) for a STRUCTURE that is not in the pool (the debug
' menu's local STRUCTURE tmp) (S: nothing sent; multistruct_c's takes an aStruct handle)
Public Function SendBuildFinished_struct(psStruct As STRUCTURE) As Boolean
    SendBuildFinished_struct = False
End Function

' --- droid.cpp (droid_c): bool checkValidWeaponForProp(const DROID_TEMPLATE *psTemplate) ---
' (droid.h:281; design.cpp's intCheckValidWeaponForProp calls it.) A DROID_TEMPLATE UDT ByRef, as droid_c's
' droidTemplateType / isTransporter_template take it.
Public Function checkValidWeaponForProp(psTemplate As DROID_TEMPLATE) As Boolean
    Dim psPropStats As Long

    'check propulsion stat for vtol
    psPropStats = DROID_TEMPLATE_getPropulsionStats(psTemplate)

    ' ASSERT_OR_RETURN(false, psPropStats != nullptr, "invalid propulsion stats pointer"): an index is never nullptr

    ' if there are no weapons, then don't even bother continuing
    If psTemplate.numWeaps = 0 Then
        checkValidWeaponForProp = False
        Exit Function
    End If

    If asPropulsionTypes(asPropulsionStats(psPropStats).propulsionType).travel = AIR Then
        'check weapon stat for indirect
        If Not proj_Direct(DROID_TEMPLATE_getWeaponStats(psTemplate, 0)) _
           Or asWeaponStats(DROID_TEMPLATE_getWeaponStats(psTemplate, 0)).vtolAttackRuns = 0 Then
            checkValidWeaponForProp = False
            Exit Function
        End If
    Else
        ' VTOL weapons do not go on non-AIR units.
        If asWeaponStats(DROID_TEMPLATE_getWeaponStats(psTemplate, 0)).vtolAttackRuns <> 0 Then
            checkValidWeaponForProp = False
            Exit Function
        End If
    End If

    'also checks that there is no other system component
    If psTemplate.asParts(COMP_BRAIN) <> 0 _
       And asWeaponStats(DROID_TEMPLATE_getWeaponStats(psTemplate, 0)).weaponSubClass <> WSC_COMMAND Then
        ' assert(false); (NDEBUG in a release build)
        checkValidWeaponForProp = False
        Exit Function
    End If

    checkValidWeaponForProp = True
End Function

' --- display3ddef.h: #define RADTLX (OBJ_BACKX + OBJ_BACKWIDTH + BASE_GAP + 1 +D_W) // Paul's settings (492+12) ---
' (a function: OBJ_BACKX and D_W depend on the screen size; OBJ_BACKX, OBJ_BACKWIDTH, BASE_GAP and D_W are
' hci.h's, fork G1's. mission.cpp and transporter.cpp use it too.)
Public Function RADTLX() As Long
    RADTLX = (OBJ_BACKX + OBJ_BACKWIDTH + BASE_GAP + 1 + D_W)
End Function


' src/mechanics.cpp: void makeAllAvailable() (S: the cheat makes nothing available yet)
Public Sub makeAllAvailable()
End Sub






' src/activity.cpp: ActivityManager::instance().completedMission(result, collectEndGameStatsData(), cheatsUsed)
' (S: the statistics argument left out, as main_c's ActivityManager_quitGame)
Public Sub ActivityManager_completedMission(ByVal result As Boolean, ByVal cheatsUsed As Boolean)
End Sub






















' main.cpp / stdinreader: void cmdInterfaceLogChatMsg(const NetworkTextMessage& message, const char* log_prefix)
Public Sub cmdInterfaceLogChatMsg(message As NetworkTextMessage, ByVal log_prefix As String)
End Sub

' activity.h: ActivityManager::instance().hostKickPlayer(NetPlay.players[player_id], type, reason)
Public Sub ActivityManager_hostKickPlayer(ByVal player As Long, ByVal typ As LOBBY_ERROR_TYPES, ByVal reason As String)
End Sub

' ActivityManager::instance().hostLobbyQuit(errorResult), joinedLobbyQuit(errorResult)
Public Sub ActivityManager_hostLobbyQuit(ByVal errorResult As LOBBY_ERROR_TYPES)
End Sub

Public Sub ActivityManager_joinedLobbyQuit(ByVal errorResult As LOBBY_ERROR_TYPES)
End Sub

' ActivityManager::instance().updateMultiplayGameData(game, ingame, NETGameIsLocked())
Public Sub ActivityManager_updateMultiplayGameData()
End Sub







' (R) lib/framework/string_ext.h: size_t nthOccurrenceOfChar(const std::string& str, const char c, int n): the
' 0-based position, -1 (std::string::npos) when there are fewer
Public Function nthOccurrenceOfChar(ByVal str_ As String, ByVal c As String, ByVal n As Long) As Long
    Dim i As Long, count As Long
    For i = 1 To Len(str_)
        If Mid$(str_, i, 1) = c Then
            count = count + 1
            If count = n Then
                nthOccurrenceOfChar = i - 1
                Exit Function
            End If
        End If
    Next
    nthOccurrenceOfChar = -1
End Function




































