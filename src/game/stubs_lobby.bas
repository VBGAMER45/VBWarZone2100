Attribute VB_Name = "stubs_lobby"
' Stand-ins that cluster J's scope keeps (docs/PLAN.md, section 36) for the joining screen (joiningscreen_c), the
' game browser (gamebrowserform_c, gamebrowser_c, WzGameBrowserTitleUI) and multijoin_helpers_c:
' * activity.cpp's join events (the activity sinks: not ported); updatemanager.cpp's version check;
' * multiint.cpp's parseKickRedirectInfo (a host's kick redirect: not ported) with multiplay.h's KickRedirectInfo;
' * MSVC's std::errc values the joining screen compares error codes with (WzNet's error_categories_c has them,
'   not public).
' (lib/netplay/netlobby.h's types and functions: netlobby_c, cluster J part 2)
Option Explicit

'==========================================================================================================
' std::errc and errno values (lib/netplay/error_categories.h)

' (enum class ConnectionProviderType: NetApi's, ConnectionProviderType_TCP_DIRECT)

' std::errc (MSVC's values): the ones the joining screen compares with
Public Const errc_connection_refused As Long = 107  ' MSVC's std::errc (WzNet's error_categories_c)
Public Const errc_connection_reset As Long = 108
Public Const errc_timed_out As Long = 138
Public Const ETIMEDOUT As Long = 138            ' errno.h (MSVC): joiningscreen.cpp's make_network_error_code(ETIMEDOUT)

'==========================================================================================================

'==========================================================================================================
' multiint.cpp's kick redirect (not ported)


'==========================================================================================================
' Staying stand-ins: activity.cpp's sinks, updatemanager.cpp (not ported)

'----------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------
' (procedures) multiint.cpp's kick redirect (not ported)


'----------------------------------------------------------------------------------------------------------
' (procedures) Staying stand-ins: activity.cpp's sinks, updatemanager.cpp (not ported)

' ActivityManager::instance().joinGameFailed()
Public Sub ActivityManager_joinGameFailed()
End Sub

' ActivityManager::instance().joinGameSucceeded(const char *host, uint32_t port)
Public Sub ActivityManager_joinGameSucceeded(ByVal host As String, ByVal port As Long)
End Sub

' ActivityManager::instance().joinLobbyGameSucceeded(const std::string& lobbyAddress, const std::string& gameId, const char *host, uint32_t port)
Public Sub ActivityManager_joinLobbyGameSucceeded(ByVal lobbyAddress As String, ByVal gameId As String, ByVal host As String, ByVal port As Long)
End Sub

' updatemanager.h: optional<bool> getVersionCheckNewVersionAvailable() (has_value: the result; the update check is
' not ported: nullopt)
Public Function getVersionCheckNewVersionAvailable(value As Boolean) As Boolean
    value = False
    getVersionCheckNewVersionAvailable = False
End Function
