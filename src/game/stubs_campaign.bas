Attribute VB_Name = "stubs_campaign"
' Stand-ins from cluster I's forks (the campaign, save and load): what their ports call that no module
' provides yet (the game guide, quick chat, mods, the multiplayer player data, ...). Each names the C++
' function it stands for; a real port elsewhere replaces it (check_names finds the duplicate).
Option Explicit























' lib/netplay (cluster J): NetPlay.scriptSetPlayerDataStrings[idx] = <the object's std::unordered_map<std::string, std::string>>
' (the scripts' player data strings are not kept before cluster J)
Public Sub NetPlay_scriptSetPlayerDataStrings_assign(ByVal idx As Long, ByVal obj As nlohmann_json)
End Sub

' src/screens/guidescreen.cpp: std::vector<std::string> saveLoadedGuideTopics() (S: no topics loaded)
' (VB) the topics in topicIds(0 .. n - 1); returns n
Public Function saveLoadedGuideTopics(topicIds() As String) As Long
    saveLoadedGuideTopics = 0
End Function

' src/screens/guidescreen.cpp: bool restoreLoadedGuideTopics(const std::vector<std::string>& topicIds) (S)
' (VB) topicIds(0 .. n - 1)
Public Function restoreLoadedGuideTopics(topicIds() As String, ByVal n As Long) As Boolean
    restoreLoadedGuideTopics = True
End Function

' src/screens/guidescreen.cpp: bool getGameGuideDisableTopicPopups() (S: the popups are enabled)
Public Function getGameGuideDisableTopicPopups() As Boolean
    getGameGuideDisableTopicPopups = False
End Function

' src/screens/guidescreen.cpp: void setGameGuideDisableTopicPopups(bool disablePopups) (S)
Public Sub setGameGuideDisableTopicPopups(ByVal disablePopups As Boolean)
End Sub

' activity.cpp (the activity sinks, cluster J): ActivityManager::instance().cheatUsed(cheat_name)
Public Sub ActivityManager_cheatUsed(ByVal cheatName As String)
End Sub

' template.cpp: void listTemplates() (a debug listing; template_c has no port of it yet). REAL, FOR template_c:
' the templates of the selected player are droidTemplates_<...> in template_c; here the log only says so.
Public Sub listTemplates()
End Sub

' REAL, FOR src/game/droid_c.bas: droid.cpp:2537 UDWORD getNumDroidsForLevel(uint32_t player, UDWORD level)
Public Function getNumDroidsForLevel(ByVal player As Long, ByVal level As Long) As Long
    Dim idx As Long, count As Long, it As Long, it2 As Long, psDroid As Long, psCurr As Long, haveList As Boolean
    idx = 0
    count = 0

    If Not ULT(player, MAX_PLAYERS) Then getNumDroidsForLevel = 0: Exit Function

    Do
        ' (VB) DroidList* dList: 0 apsDroidLists, 1 apsLimboDroids (nullptr: not taken)
        haveList = False
        Select Case idx
        Case 0: it = apsDroidLists(selectedPlayer).head: haveList = True
        Case 1
            If prevMissionType = LEVEL_TYPE_LDS_MKEEP_LIMBO Then
                it = apsLimboDroids(selectedPlayer).head: haveList = True
            End If
        End Select
        If Not haveList Then
            GoTo continue_
        End If
        Do While it <> 0
            psDroid = lnObj(it)
            it = lnNext(it)
            If psDroid = 0 Then GoTo continueDroid
            If isDead(psDroid) Then GoTo continueDroid
            If U32(getDroidLevel(psDroid)) = U32(level) Then
                count = count + 1
            End If
            If DROID_isTransporter(psDroid) Then
                If aDroid(psDroid).psGroup = 0 Then
                    GoTo continueDroid
                End If
                it2 = aGroup(aDroid(psDroid).psGroup).psList.head
                Do While it2 <> 0
                    psCurr = lnObj(it2)
                    it2 = lnNext(it2)
                    If psCurr <> psDroid Then
                        If U32(getDroidLevel(psCurr)) = U32(level) Then
                            count = count + 1
                        End If
                    End If
                Loop
            End If
continueDroid:
        Loop
continue_:
        idx = idx + 1
    Loop While idx < 2

    getNumDroidsForLevel = count
End Function

' --- lib/framework/wzapp.h: void wzAsyncExecOnMainThread(const std::function<void ()> &func) (the backend's,
' not ported: platform's main_win.bas). REAL enough: the functions queue (a Static Collection) and
' wzAsyncExecOnMainThread_runPending runs them, which main_c's mainLoop calls first (where the SDL backend's event
' loop runs its WZ_SDL_MAINTHREAD_EXEC events, before mainLoop) ---
Private Function wzAsyncExecOnMainThread_queue() As Collection
    Static q As Collection
    If q Is Nothing Then Set q = New Collection
    Set wzAsyncExecOnMainThread_queue = q
End Function

Public Sub wzAsyncExecOnMainThread(ByVal func As WzFunc)
    wzAsyncExecOnMainThread_queue().Add func
End Sub

' (VB) runs the queued functions, in their order (a function queued while they run waits for the next call)
Public Sub wzAsyncExecOnMainThread_runPending()
    Dim q As Collection, n As Long, f As WzFunc
    Set q = wzAsyncExecOnMainThread_queue()
    n = q.Count
    Do While n > 0
        Set f = q.Item(1)
        q.Remove 1
        n = n - 1
        WzFunc_call f
    Loop
End Sub

' --- src/urlrequest.h (curl: not ported) ---
' AsyncURLRequestHandle urlRequestData(const URLDataRequest& request) (the request's url only): no handle
Public Function urlRequestData(ByVal url As String) As Long
    urlRequestData = 0
End Function

' void urlRequestSetCancelFlag(AsyncURLRequestHandle requestHandle)
Public Sub urlRequestSetCancelFlag(ByVal requestHandle As Long)
End Sub
