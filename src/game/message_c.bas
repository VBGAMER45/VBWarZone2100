Attribute VB_Name = "message_c"
' src/message.cpp, message.h, messagedef.h (ported in part): the message
' and proximity display lists (apsMessages, apsProxDisp) with their set-up
' and release (messageInitVars, initViewData, initMessage, freeMessages,
' releaseAllProxDisp), the messages themselves (createMessage,
' addMessageToList, removeMessageFromList, releaseAllMessages,
' addBeaconMessage, addProximityDisplay, addMessage, removeProxDisp,
' removeMessage, both findMessage) and the VIEWDATA a beacon makes
' (multiplay.cpp's CreateBeaconViewData).
' A MESSAGE *, PROXIMITY_DISPLAY * and VIEWDATA * are 1-based handles into
' the heaps aMessage, aProxDisp and aViewData (new and delete: pools with
' free lists), 0 = nullptr; a list holds them (objlist_vb).
' findMessage(const VIEWDATA *, ...) is findMessage_viewData (a handle
' each: the overloads would clash).
' A VIEWDATA's VIEW_BASE *pData is its VIEW_PROXIMITY (fields prox_*,
' hasProxData), VIEW_RESEARCH (res, hasResData) or VIEW_REPLAY (its seqList
' replay_seqList, hasReplayData). The view data the files hold (apsViewData:
' a name-sorted array, getViewData, loadViewData and the JSON loaders,
' viewDataShutDown) and the display of the proximity messages came with
' cluster I.
Option Explicit

' message.h
Public Const NO_AUDIO_MSG As Long = -1

' messagedef.h
Public Enum VIEW_TYPE
    VIEW_RES            ' research view
    VIEW_RPL            ' full screen view sequence - flic
    VIEW_PROX           ' proximity view - no view really!
    VIEW_RPLX           ' full screen view sequence - flic.	extended format
    VIEW_BEACON         ' Beacon message
    VIEW_SIZE
End Enum

Public Enum PROX_TYPE
    PROX_ENEMY          ' enemy proximity message
    PROX_RESOURCE       ' resource proximity message
    PROX_ARTEFACT       ' artefact proximity message
    PROX_TYPES
End Enum

Public Enum MSG_DATA_TYPE
    MSG_DATA_DEFAULT
    MSG_DATA_BEACON
End Enum

' struct MESSAGE
Public Type MESSAGE_
    typ As MESSAGE_TYPE                 ' The type of message
    id As Long                          ' UDWORD: ID number of the message
    pViewData As Long                   ' VIEWDATA *: Pointer to view data - if any - should be some!
    psObj As Long                       ' BASE_OBJECT *
    read As Boolean                     ' flag to indicate whether message has been read
    player As Long                      ' UDWORD: which player this message belongs to
    dataType As MSG_DATA_TYPE
    nextFree As Long                    ' (VB) the heap's free list
End Type

' struct PROXIMITY_DISPLAY : public OBJECT_POSITION
Public Type PROXIMITY_DISPLAY_
    typ As POSITION_TYPE                ' OBJECT_POSITION's
    frameNumber As Long
    screenX As Long
    screenY As Long
    screenR As Long
    player As Long
    selected As Boolean
    psMessage As Long                   ' MESSAGE *: message associated with this 'button'
    timeLastDrawn As Long               ' UDWORD: stores the time the 'button' was last drawn for animation
    strobe As Long                      ' UDWORD: id of image last used
    buttonID As Long                    ' UDWORD: id of the button for the interface
    nextFree As Long                    ' (VB) the heap's free list
End Type

' struct VIEW_RESEARCH : VIEW_BASE
Public Type VIEW_RESEARCH
    pIMD As Long                            ' iIMDBaseShape * = nullptr
    pIMD2 As Long                           ' iIMDBaseShape * = nullptr: allows base plates and turrets to be drawn as well
    sequenceName As String                  ' which windowed flic to display
    audio As String                         ' name of audio track to play (for this seq)
End Type

' struct SEQ_DISPLAY
Public Type SEQ_DISPLAY
    sequenceName As String
    flag As Long                            ' UBYTE: flag data to control video playback 1 = loop till audio finish
    textMsg As StringVector                 ' std::vector<WzString>: Text messages - if any
    audio As String                         ' name of audio track to play (for this seq)
End Type

' struct VIEWDATA (with its VIEW_BASE *pData: a VIEW_PROXIMITY when hasProxData, a VIEW_RESEARCH (res) when
' hasResData, a VIEW_REPLAY (its seqList: replay_seqList) when hasReplayData)
Public Type VIEWDATA_
    name As String                      ' name ID of the message - used for loading in and identifying
    typ As VIEW_TYPE                    ' the type of view
    textMsg As StringVector             ' Text messages, if any
    hasResData As Boolean               ' pData is a VIEW_RESEARCH
    res As VIEW_RESEARCH
    hasReplayData As Boolean            ' pData is a VIEW_REPLAY
    replay_seqList() As SEQ_DISPLAY     ' VIEW_REPLAY::seqList: replay_seqList(0 .. replay_seqList_size - 1)
    replay_seqList_size As Long
    hasProxData As Boolean              ' VIEW_BASE *pData != nullptr (a VIEW_PROXIMITY)
    prox_x As Long                      ' UDWORD: world coords for position of Proximity message
    prox_y As Long
    prox_z As Long
    prox_proxType As PROX_TYPE
    prox_audioID As Long                ' SDWORD: ID of the audio track to play - if any
    prox_sender As Long                 ' SDWORD: user who sent this msg
    prox_timeAdded As Long              ' SDWORD: remember when was added, so can remove after certain period of time
    fileName As String                  ' file it came from, for piecemeal destruction (pretty lame reason)
    nextFree As Long                    ' (VB) the heap's free list
End Type

' static std::map<WzString, VIEWDATA *> apsViewData: the names in their byte order (binary compare: a VB
' Collection's keys ignore case) and the handles, apsViewData_size of them
Private apsViewData_names() As String
Private apsViewData_ptrs() As Long
Private apsViewData_size As Long

' The id number for the next message allocated
' Each message will have a unique id number irrespective of type
Private msgID As Long                   ' UDWORD

Private currentNumProxDisplays As Long

Public apsMessages(0 To MAX_PLAYERS - 1) As OBJ_LIST    ' PerPlayerMessageLists

' The list of proximity displays allocated
Public apsProxDisp(0 To MAX_PLAYERS - 1) As OBJ_LIST    ' PerPlayerProximityDisplayLists

' The IMD to use for the proximity messages
Public pProximityMsgIMD As Long         ' iIMDBaseShape *

' (VB) the heaps
Public aMessage() As MESSAGE_
Private aMessage_cap As Long
Private aMessage_free As Long
Public aProxDisp() As PROXIMITY_DISPLAY_
Private aProxDisp_cap As Long
Private aProxDisp_free As Long
Public aViewData() As VIEWDATA_
Private aViewData_cap As Long
Private aViewData_free As Long

' (VB) new MESSAGE: a slot of aMessage, value-initialised
Private Function messageNew() As Long
    Dim i As Long, n As Long, blank As MESSAGE_
    If aMessage_free = 0 Then
        n = aMessage_cap
        If aMessage_cap = 0 Then aMessage_cap = 64 Else aMessage_cap = aMessage_cap * 2
        ReDim Preserve aMessage(1 To aMessage_cap)
        For i = aMessage_cap To n + 1 Step -1
            aMessage(i).nextFree = aMessage_free
            aMessage_free = i
        Next
    End If
    messageNew = aMessage_free
    aMessage_free = aMessage(messageNew).nextFree
    aMessage(messageNew) = blank
End Function

' (VB) delete MESSAGE
Private Sub messageDelete(ByVal psMsg As Long)
    Dim blank As MESSAGE_
    aMessage(psMsg) = blank
    aMessage(psMsg).nextFree = aMessage_free
    aMessage_free = psMsg
End Sub

' (VB) new PROXIMITY_DISPLAY
Private Function proxDispNew() As Long
    Dim i As Long, n As Long, blank As PROXIMITY_DISPLAY_
    If aProxDisp_free = 0 Then
        n = aProxDisp_cap
        If aProxDisp_cap = 0 Then aProxDisp_cap = 64 Else aProxDisp_cap = aProxDisp_cap * 2
        ReDim Preserve aProxDisp(1 To aProxDisp_cap)
        For i = aProxDisp_cap To n + 1 Step -1
            aProxDisp(i).nextFree = aProxDisp_free
            aProxDisp_free = i
        Next
    End If
    proxDispNew = aProxDisp_free
    aProxDisp_free = aProxDisp(proxDispNew).nextFree
    aProxDisp(proxDispNew) = blank
End Function

' (VB) delete PROXIMITY_DISPLAY
Private Sub proxDispDelete(ByVal psDisp As Long)
    Dim blank As PROXIMITY_DISPLAY_
    aProxDisp(psDisp) = blank
    aProxDisp(psDisp).nextFree = aProxDisp_free
    aProxDisp_free = psDisp
End Sub

' (VB) new VIEWDATA (multiplay.cpp's CreateBeaconViewData makes them; nothing deletes a beacon's, as in C++)
Public Function viewDataNew() As Long
    Dim i As Long, n As Long, blank As VIEWDATA_
    If aViewData_free = 0 Then
        n = aViewData_cap
        If aViewData_cap = 0 Then aViewData_cap = 64 Else aViewData_cap = aViewData_cap * 2
        ReDim Preserve aViewData(1 To aViewData_cap)
        For i = aViewData_cap To n + 1 Step -1
            aViewData(i).nextFree = aViewData_free
            aViewData_free = i
        Next
    End If
    viewDataNew = aViewData_free
    aViewData_free = aViewData(viewDataNew).nextFree
    aViewData(viewDataNew) = blank
End Function

' Creating a new message
' new is a pointer to a pointer to the new message
' type is the type of the message
Private Function createMessage(ByVal msgType As MESSAGE_TYPE, ByVal player As Long) As Long
    Dim newMsg As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/message.cpp:72", "createMessage", "Bad player": createMessage = 0: Exit Function
    If Not (msgType < MSG_TYPES) Then ASSERT_FAILURE "msgType < MSG_TYPES", "src/message.cpp:73", "createMessage", "Bad message": createMessage = 0: Exit Function

    ' Allocate memory for the message, and on failure return a NULL pointer
    newMsg = messageNew()
    aMessage(newMsg).typ = msgType
    aMessage(newMsg).dataType = MSG_DATA_DEFAULT
    aMessage(newMsg).id = MUL32(msgID, 8) Or selectedPlayer   ' (msgID << 3) | selectedPlayer
    aMessage(newMsg).player = player
    msgID = FromU32(U32(msgID) + 1)

    createMessage = newMsg
End Function

' (VB) list.emplace(it, msg): msg inserted before the list's node it (0: at the end)
Private Sub messageListInsertBefore(l As OBJ_LIST, ByVal it As Long, ByVal msg As Long)
    Dim items() As Long, n As Long, k As Long, node As Long
    ReDim items(0 To l.size)
    node = l.head
    Do While node <> 0
        If node = it Then
            items(n) = msg                      ' msg goes before this element
            n = n + 1
        End If
        items(n) = lnObj(node)
        n = n + 1
        node = lnNext(node)
    Loop
    If it = 0 Then
        items(n) = msg                          ' end(): at the back
        n = n + 1
    End If
    objlist_clear l
    For k = 0 To n - 1
        objlist_push_back l, items(k)
    Next
End Sub

' Add the message to the BOTTOM of the list
' list is a pointer to the message list
' Order is now CAMPAIGN, MISSION, RESEARCH/PROXIMITY
Private Sub addMessageToList(list() As OBJ_LIST, ByVal msg As Long, ByVal player As Long)
    Dim it As Long, t As MESSAGE_TYPE

    If Not (msg <> 0) Then ASSERT_FAILURE "msg != nullptr", "src/message.cpp:92", "addMessageToList", "Invalid message pointer": Exit Sub
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/message.cpp:93", "addMessageToList", "Bad player": Exit Sub

    ' If there is no message list, create one
    If list(player).size = 0 Then
        objlist_push_front list(player), msg
        Exit Sub
    End If

    Select Case aMessage(msg).typ
    Case MSG_CAMPAIGN
        'add it before the first mission/research/prox message
        it = list(player).head
        Do While it <> 0
            t = aMessage(lnObj(it)).typ
            If t = MSG_MISSION Or t = MSG_RESEARCH Or t = MSG_PROXIMITY Then Exit Do
            it = lnNext(it)
        Loop
        messageListInsertBefore list(player), it, msg
    Case MSG_MISSION
        'add it before the first research/prox message
        it = list(player).head
        Do While it <> 0
            t = aMessage(lnObj(it)).typ
            If t = MSG_RESEARCH Or t = MSG_PROXIMITY Then Exit Do
            it = lnNext(it)
        Loop
        messageListInsertBefore list(player), it, msg
    Case MSG_RESEARCH, MSG_PROXIMITY
        'add it to the bottom of the list
        objlist_push_back list(player), msg
    Case Else
        debug_ 133, LOG_ERROR, "addMessageToList", "unknown message type"
    End Select
End Sub

' Remove a message from the list
' list is a pointer to the message list
' del is a pointer to the message to remove
Private Sub removeMessageFromList(list() As OBJ_LIST, ByVal del As Long, ByVal player As Long)
    Dim it As Long

    If Not (del <> 0) Then ASSERT_FAILURE "del != nullptr", "src/message.cpp:146", "removeMessageFromList", "Invalid message pointer": Exit Sub
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/message.cpp:147", "removeMessageFromList", "Bad player": Exit Sub

    it = objlist_find(list(player), del)
    If Not (it <> 0) Then ASSERT_FAILURE "it != list[player].end()", "src/message.cpp:150", "removeMessageFromList", "Message %p not found in list", del: Exit Sub
    objlist_erase list(player), it
    messageDelete del
End Sub

Private Sub releaseAllMessages(list() As OBJ_LIST)
    Dim pl As Long, it As Long
    ' Iterate through all players' message lists
    For pl = 0 To MAX_PLAYERS - 1
        ' Iterate through all messages in list
        it = list(pl).head
        Do While it <> 0
            messageDelete lnObj(it)
            it = lnNext(it)
        Loop
        objlist_clear list(pl)
    Next
End Sub

Public Function messageInitVars() As Boolean
    Dim i As Long

    msgID = 0
    currentNumProxDisplays = 0

    For i = 0 To MAX_PLAYERS - 1
        objlist_clear apsMessages(i)
        objlist_clear apsProxDisp(i)
    Next

    pProximityMsgIMD = 0

    messageInitVars = True
End Function

'allocates the viewdata heap
Public Function initViewData() As Boolean
    initViewData = True
End Function

Public Function addBeaconMessage(ByVal msgType As MESSAGE_TYPE, ByVal proxPos As Boolean, ByVal player As Long) As Long
    Dim psBeaconMsgToAdd As Long
    psBeaconMsgToAdd = addMessage(msgType, proxPos, player)

    If Not (psBeaconMsgToAdd <> 0) Then ASSERT_FAILURE "psBeaconMsgToAdd", "src/message.cpp:198", "addBeaconMessage", "createMessage failed": addBeaconMessage = 0: Exit Function

    ' remember we are storing beacon data in this message
    aMessage(psBeaconMsgToAdd).dataType = MSG_DATA_BEACON

    addBeaconMessage = psBeaconMsgToAdd
End Function

' adds a proximity display - holds variables that enable the message to be
' displayed in the Intelligence Screen
Private Sub addProximityDisplay(ByVal psMessage As Long, ByVal proxPos As Boolean, ByVal player As Long)
    Dim psToAdd As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/message.cpp:212", "addProximityDisplay", "Bad player": Exit Sub
    debug_ 213, LOG_MSG, "addProximityDisplay", "Added prox display for player %u (proxPos=%d)", player, -CLng(proxPos)

    'create the proximity display
    psToAdd = proxDispNew()

    If proxPos Then
        aProxDisp(psToAdd).typ = POS_PROXOBJ
    Else
        aProxDisp(psToAdd).typ = POS_PROXDATA
    End If
    aProxDisp(psToAdd).psMessage = psMessage
    aProxDisp(psToAdd).screenX = 0
    aProxDisp(psToAdd).screenY = 0
    aProxDisp(psToAdd).screenR = 0
    aProxDisp(psToAdd).player = player
    aProxDisp(psToAdd).timeLastDrawn = 0
    aProxDisp(psToAdd).frameNumber = 0
    aProxDisp(psToAdd).selected = False
    aProxDisp(psToAdd).strobe = 0

    'add a button to the interface
    If intAddProximityButton(psToAdd, currentNumProxDisplays) Then
        ' Now add it to the top of the list. Be aware that this
        ' check means that messages and proximity displays can
        ' become out of sync - but this should never happen.
        objlist_push_front apsProxDisp(player), psToAdd
        currentNumProxDisplays = currentNumProxDisplays + 1
    Else
        proxDispDelete psToAdd  ' clean up
    End If
End Sub

'Add a message to the list
Public Function addMessage(ByVal msgType As MESSAGE_TYPE, ByVal proxPos As Boolean, ByVal player As Long) As Long
    Dim psMsgToAdd As Long
    'first create a message of the required type
    psMsgToAdd = createMessage(msgType, player)

    debug_ 257, LOG_MSG, "addMessage", "adding message for player %d, type is %d, proximity is %d", player, msgType, -CLng(proxPos)

    If Not (psMsgToAdd <> 0) Then ASSERT_FAILURE "psMsgToAdd", "src/message.cpp:259", "addMessage", "createMessage failed"
    If psMsgToAdd = 0 Then
        addMessage = 0
        Exit Function
    End If
    'then add to the players' list
    addMessageToList apsMessages, psMsgToAdd, player

    'add a proximity display
    If msgType = MSG_PROXIMITY Then
        addProximityDisplay psMsgToAdd, proxPos, player
    End If

    addMessage = psMsgToAdd
End Function

' remove a proximity display
Private Sub removeProxDisp(ByVal psMessage As Long, ByVal player As Long)
    Dim it As Long, psDisp As Long

    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/message.cpp:279", "removeProxDisp", "Bad player": Exit Sub
    If Not (psMessage <> 0) Then ASSERT_FAILURE "psMessage != nullptr", "src/message.cpp:280", "removeProxDisp", "Bad message": Exit Sub

    If apsProxDisp(player).size = 0 Then
        Exit Sub    ' no corresponding proximity display
    End If

    'find the proximity display for this message
    it = apsProxDisp(player).head
    Do While it <> 0
        If aProxDisp(lnObj(it)).psMessage = psMessage Then Exit Do
        it = lnNext(it)
    Loop
    If it <> 0 Then
        psDisp = lnObj(it)
        objlist_erase apsProxDisp(player), it
        intRemoveProximityButton psDisp
        proxDispDelete psDisp
    End If
End Sub

'remove a message
Public Sub removeMessage(ByVal psDel As Long, ByVal player As Long)
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/message.cpp:304", "removeMessage", "Bad player": Exit Sub
    If Not (psDel <> 0) Then ASSERT_FAILURE "psDel != nullptr", "src/message.cpp:305", "removeMessage", "Bad message": Exit Sub
    debug_ 306, LOG_MSG, "removeMessage", "removing message for player %d", player

    If aMessage(psDel).typ = MSG_PROXIMITY Then
        removeProxDisp psDel, player
    End If
    removeMessageFromList apsMessages, psDel, player
End Sub

' Remove all Messages
Public Sub freeMessages()
    releaseAllProxDisp
    releaseAllMessages apsMessages
    jsDebugMessageUpdate
End Sub

' removes all the proximity displays
Public Sub releaseAllProxDisp()
    Dim removedAMessage As Boolean, player As Long, it As Long, psCurr As Long

    removedAMessage = False

    For player = 0 To MAX_PLAYERS - 1
        ' mutating_list_iterate: the next node is taken before the display's message (and the display) is removed
        it = apsProxDisp(player).head
        Do While it <> 0
            psCurr = lnObj(it)
            it = lnNext(it)
            'remove message associated with this display
            removeMessage aProxDisp(psCurr).psMessage, player
            removedAMessage = True
        Loop
        objlist_clear apsProxDisp(player)
    Next
    're-initialise variables
    currentNumProxDisplays = 0
    If removedAMessage Then
        jsDebugMessageUpdate
    End If
End Sub

' Initialise the message heaps
Public Function initMessage() As Boolean
    'set up the imd used for proximity messages
    pProximityMsgIMD = modelGet("arrow.pie")
    If pProximityMsgIMD = 0 Then
        ASSERT_FAILURE "false", "src/message.cpp:355", "initMessage", "Unable to load Proximity Message PIE"
        initMessage = False
        Exit Function
    End If

    initMessage = True
End Function

' Looks through the players list of messages to find one with the same viewdata ptr and which is the same type
' (MESSAGE *findMessage(const VIEWDATA *pViewData, MESSAGE_TYPE type, UDWORD player))
Public Function findMessage_viewData(ByVal pViewData As Long, ByVal typ As MESSAGE_TYPE, ByVal player As Long) As Long
    Dim it As Long, psCurr As Long
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/message.cpp:1018", "findMessage", "Bad player": findMessage_viewData = 0: Exit Function
    If Not (typ < MSG_TYPES) Then ASSERT_FAILURE "type < MSG_TYPES", "src/message.cpp:1019", "findMessage", "Bad message type": findMessage_viewData = 0: Exit Function

    it = apsMessages(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aMessage(psCurr).typ = typ And aMessage(psCurr).pViewData = pViewData Then
            findMessage_viewData = psCurr
            Exit Function
        End If
        it = lnNext(it)
    Loop

    'not found the message so return NULL
    findMessage_viewData = 0
End Function

' Looks through the players list of messages to find one with the same object ptr and which is the same type
Public Function findMessage(ByVal psObj As Long, ByVal typ As MESSAGE_TYPE, ByVal player As Long) As Long
    Dim it As Long, psCurr As Long
    If Not (ULT(player, MAX_PLAYERS)) Then ASSERT_FAILURE "player < MAX_PLAYERS", "src/message.cpp:1035", "findMessage", "Bad player": findMessage = 0: Exit Function
    If Not (typ < MSG_TYPES) Then ASSERT_FAILURE "type < MSG_TYPES", "src/message.cpp:1036", "findMessage", "Bad message type": findMessage = 0: Exit Function

    it = apsMessages(player).head
    Do While it <> 0
        psCurr = lnObj(it)
        If aMessage(psCurr).typ = typ And aMessage(psCurr).psObj = psObj Then
            findMessage = psCurr
            Exit Function
        End If
        it = lnNext(it)
    Loop

    findMessage = 0
End Function

' (VB) delete VIEWDATA (and its pData)
Private Sub viewDataDelete(ByVal psViewData As Long)
    Dim blank As VIEWDATA_
    aViewData(psViewData) = blank
    aViewData(psViewData).nextFree = aViewData_free
    aViewData_free = psViewData
End Sub

' (VB) apsViewData.find(name): the index in the names' order, or -1 (end()); insertAt is where it would go
Private Function apsViewData_find(ByVal name_ As String, Optional insertAt As Long) As Long
    Dim lo As Long, hi As Long, mid_ As Long, c As Long
    lo = 0
    hi = apsViewData_size - 1
    Do While lo <= hi
        mid_ = (lo + hi) \ 2
        c = StrComp(apsViewData_names(mid_), name_, vbBinaryCompare)
        If c = 0 Then
            apsViewData_find = mid_
            insertAt = mid_
            Exit Function
        ElseIf c < 0 Then
            lo = mid_ + 1
        Else
            hi = mid_ - 1
        End If
    Loop
    insertAt = lo
    apsViewData_find = -1
End Function

' (VB) apsViewData[name] = psViewData
Private Sub apsViewData_set(ByVal name_ As String, ByVal psViewData As Long)
    Dim i As Long, at_ As Long
    i = apsViewData_find(name_, at_)
    If i <> -1 Then
        apsViewData_ptrs(i) = psViewData
        Exit Sub
    End If
    If apsViewData_size = 0 Then
        ReDim apsViewData_names(0 To 0)
        ReDim apsViewData_ptrs(0 To 0)
    Else
        ReDim Preserve apsViewData_names(0 To apsViewData_size)
        ReDim Preserve apsViewData_ptrs(0 To apsViewData_size)
    End If
    For i = apsViewData_size To at_ + 1 Step -1
        apsViewData_names(i) = apsViewData_names(i - 1)
        apsViewData_ptrs(i) = apsViewData_ptrs(i - 1)
    Next
    apsViewData_names(at_) = name_
    apsViewData_ptrs(at_) = psViewData
    apsViewData_size = apsViewData_size + 1
End Sub

' (VB) apsViewData.erase(it): the entry at index i
Private Sub apsViewData_erase(ByVal i As Long)
    Dim k As Long
    For k = i To apsViewData_size - 2
        apsViewData_names(k) = apsViewData_names(k + 1)
        apsViewData_ptrs(k) = apsViewData_ptrs(k + 1)
    Next
    apsViewData_size = apsViewData_size - 1
End Sub

' Return the number of newlines in a file buffer
Private Function numCR(pFileBuffer() As Byte, ByVal fileSize As Long) As Long
    Dim lines As Long, i As Long
    lines = 0

    For i = 0 To fileSize - 1
        If pFileBuffer(i) = 10 Then          ' '\n'
            lines = lines + 1
        End If
    Next

    numCR = lines
End Function

'load the view data for the messages from the file
' WzString *loadViewData(const char *pViewMsgData, UDWORD bufferSize): (VB) the WzString in ptr, False for nullptr.
' The buffer is a String (s) and pViewMsgData an index into it (p, 1-based); each sscanf reads Mid$(s, p) and %n
' moves p.
Public Function loadViewData(pViewMsgData() As Byte, ByVal bufferSize As Long, ptr As String) As Boolean
    Dim numData As Long, count As Long, dummy As Long, name_ As String, imdName As String, string_ As String, imdName2 As String
    Dim audioName As String, LocX As Long, LocY As Long, LocZ As Long, audioID As Long, proxType As Long, cnt As Long
    Dim filename As String, i As Long, psViewData As Long, numText As Long, readint As Long, dataInc As Long, seqInc As Long
    Dim s As String, p As Long, str_ As String, numSeqText As Long, count2 As Long, tmp As Long, k As Long, z As SEQ_DISPLAY
    Dim nm As String

    filename = GetLastResourceFilename()
    nm = "%255[^,'" & vbCr & vbLf & "]"     ' the C's %255[^,'\r\n]

    s = utf8_decode(pViewMsgData, 0, bufferSize)
    p = 1
    numData = numCR(pViewMsgData, bufferSize)
    For i = 0 To numData - 1
        psViewData = viewDataNew()

        aViewData(psViewData).typ = VIEW_SIZE
        aViewData(psViewData).fileName = filename
        name_ = ""

        'read the data into the storage - the data is delimited using comma's
        cnt = 0
        sscanf Mid$(s, p), nm & ",%u%n", name_, numText, cnt
        p = p + cnt

        'allocate storage for the name
        aViewData(psViewData).name = name_
        debug_ 412, LOG_MSG, "loadViewData", "Loaded %s", aViewData(psViewData).name

        'read in the data for the text strings
        For dataInc = 0 To numText - 1
            name_ = ""
            cnt = 0
            sscanf Mid$(s, p), "," & nm & "%n", name_, cnt
            p = p + cnt

            ' Get the string from the ID string
            str_ = strresGetString(psStringRes, name_)
            If Not (StrPtr(str_) <> 0) Then ASSERT_FAILURE "str", "src/message.cpp:423", "loadViewData", "Cannot find the view data string with id ""%s""", name_
            StringVector_push_back aViewData(psViewData).textMsg, str_
        Next

        cnt = 0
        sscanf Mid$(s, p), ",%d%n", readint, cnt
        aViewData(psViewData).typ = readint
        p = p + cnt

        'allocate data according to type
        Select Case aViewData(psViewData).typ
        Case VIEW_RES
            aViewData(psViewData).hasResData = True     ' psViewData->pData = new VIEW_RESEARCH;
            imdName = ""
            imdName2 = ""
            string_ = ""
            audioName = ""
            cnt = 0
            sscanf Mid$(s, p), "," & nm & "," & nm & "," & nm & "," & nm & ",%d%n", _
                   imdName, imdName2, string_, audioName, dummy, cnt
            p = p + cnt
            aViewData(psViewData).res.pIMD = modelGet(imdName)
            If aViewData(psViewData).res.pIMD = 0 Then
                ASSERT_FAILURE "false", "src/message.cpp:448", "loadViewData", "Cannot find the PIE for message %s", name_
                viewDataDelete psViewData
                loadViewData = False
                Exit Function
            End If
            If imdName2 <> "0" Then
                aViewData(psViewData).res.pIMD2 = modelGet(imdName2)
                If aViewData(psViewData).res.pIMD2 = 0 Then
                    ASSERT_FAILURE "false", "src/message.cpp:457", "loadViewData", "Cannot find the 2nd PIE for message %s", name_
                    viewDataDelete psViewData
                    loadViewData = False
                    Exit Function
                End If
            Else
                aViewData(psViewData).res.pIMD2 = 0
            End If
            aViewData(psViewData).res.sequenceName = string_
            'get the audio text string
            If audioName <> "0" Then
                'allocate space
                aViewData(psViewData).res.audio = audioName
            End If
        Case VIEW_RPL, VIEW_RPLX
            ' This is now also used for the stream playing on the PSX
            ' NOTE: on the psx the last entry (audioID) is used as the number of frames in the stream
            aViewData(psViewData).hasReplayData = True  ' psViewData->pData = new VIEW_REPLAY;

            'read in number of sequences for this message
            cnt = 0
            sscanf Mid$(s, p), ",%u%n", count, cnt
            p = p + cnt

            ' psViewReplay->seqList.resize(count);
            If count > 0 Then
                ReDim aViewData(psViewData).replay_seqList(0 To count - 1)
                For k = 0 To count - 1
                    aViewData(psViewData).replay_seqList(k) = z
                Next
            End If
            aViewData(psViewData).replay_seqList_size = count

            'read in the data for the sequences
            For dataInc = 0 To aViewData(psViewData).replay_seqList_size - 1
                numSeqText = 0

                name_ = ""
                'load extradat for extended type only
                If aViewData(psViewData).typ = VIEW_RPL Then
                    cnt = 0
                    sscanf Mid$(s, p), "," & nm & ",%u%n", name_, count, cnt
                    p = p + cnt
                    'set the flag to default
                    aViewData(psViewData).replay_seqList(dataInc).flag = 0
                    numSeqText = count
                Else 'extended type
                    cnt = 0
                    sscanf Mid$(s, p), "," & nm & ",%u,%d%n", name_, count, count2, cnt
                    p = p + cnt
                    aViewData(psViewData).replay_seqList(dataInc).flag = count And &HFF&
                    numSeqText = count2
                End If
                aViewData(psViewData).replay_seqList(dataInc).sequenceName = name_

                'get the text strings for this sequence - if any
                For seqInc = 0 To numSeqText - 1
                    name_ = ""
                    cnt = 0
                    sscanf Mid$(s, p), "," & nm & "%n", name_, cnt
                    p = p + cnt

                    ' Get the string from the ID string
                    str_ = strresGetString(psStringRes, name_)
                    If Not (StrPtr(str_) <> 0) Then ASSERT_FAILURE "str", "src/message.cpp:521", "loadViewData", "Cannot find the view data string with id ""%s""", name_
                    StringVector_push_back aViewData(psViewData).replay_seqList(dataInc).textMsg, str_
                Next
                'get the audio text string
                cnt = 0
                sscanf Mid$(s, p), "," & nm & ",%d%n", audioName, dummy, cnt
                p = p + cnt

                If audioName <> "0" Then
                    'allocate space
                    ASSERT_FAILURE "false", "src/message.cpp:532", "loadViewData", "Unexpected separate audio track provided (%s), for %s in %s - this is deprecated!", audioName, aViewData(psViewData).replay_seqList(dataInc).sequenceName, filename
                    aViewData(psViewData).replay_seqList(dataInc).audio = audioName
                End If
            Next
            aViewData(psViewData).typ = VIEW_RPL 'no longer need to know if it is extended type

        Case VIEW_PROX
            aViewData(psViewData).hasProxData = True    ' psViewData->pData = new VIEW_PROXIMITY;
            audioName = ""
            cnt = 0
            sscanf Mid$(s, p), ", %d,%d,%d," & nm & ",%d%n", LocX, LocY, LocZ, audioName, tmp, cnt
            proxType = tmp
            p = p + cnt

            'allocate audioID
            If audioName = "0" Then
                audioID = NO_SOUND
            Else
                audioID = audio_GetIDFromStr(audioName)
                If audioID = NO_SOUND Then
                    ASSERT_FAILURE "false", "src/message.cpp:560", "loadViewData", "couldn't get ID %d for weapon sound %s", audioID, audioName
                    loadViewData = False
                    Exit Function
                End If

                If (audioID < 0 Or audioID > ID_MAX_SOUND) And audioID <> NO_SOUND Then
                    ASSERT_FAILURE "false", "src/message.cpp:568", "loadViewData", "Invalid Weapon Sound ID - %d for weapon %s", audioID, audioName
                    loadViewData = False
                    Exit Function
                End If
            End If

            aViewData(psViewData).prox_audioID = audioID

            If LocX < 0 Then
                ASSERT_FAILURE "false", "src/message.cpp:578", "loadViewData", "Negative X coord for prox message - %s", name_
                loadViewData = False
                Exit Function
            End If
            aViewData(psViewData).prox_x = LocX
            If LocY < 0 Then
                ASSERT_FAILURE "false", "src/message.cpp:584", "loadViewData", "Negative Y coord for prox message - %s", name_
                loadViewData = False
                Exit Function
            End If
            aViewData(psViewData).prox_y = LocY
            If LocZ < 0 Then
                ASSERT_FAILURE "false", "src/message.cpp:590", "loadViewData", "Negative Z coord for prox message - %s", name_
                loadViewData = False
                Exit Function
            End If
            aViewData(psViewData).prox_z = LocZ

            If proxType > PROX_TYPES Then
                ASSERT_FAILURE "false", "src/message.cpp:597", "loadViewData", "Invalid proximity message sub type - %s", name_
                loadViewData = False
                Exit Function
            End If
            aViewData(psViewData).prox_proxType = proxType
        Case Else
            ASSERT_FAILURE "false", "src/message.cpp:603", "loadViewData", "Unknown ViewData type"
            loadViewData = False
            Exit Function
        End Select
        'increment the pointer to the start of the next record
        k = InStr(p, s, vbLf, vbBinaryCompare)
        If k = 0 Then p = Len(s) + 1 Else p = k + 1

        apsViewData_set aViewData(psViewData).name, psViewData
    Next

    ptr = filename ' so that cleanup function will be called for correct data
    loadViewData = True
End Function

' WzString *loadResearchViewData(const char *fileName): (VB) the WzString in ptr, False for nullptr
Public Function loadResearchViewData(ByVal fileName As String, ptr As String) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, v As Long, j As Long, texts() As String, nTexts As Long, t As String
    If Not (PHYSFS_exists(fileName) <> 0) Then ASSERT_FAILURE "PHYSFS_exists(fileName)", "src/message.cpp:617", "loadResearchViewData", "%s not found", fileName: loadResearchViewData = False: Exit Function
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    n = ini.childGroups(list)
    For i = 0 To n - 1
        v = viewDataNew()

        aViewData(v).name = list(i)
        aViewData(v).fileName = fileName

        ini.beginGroup list(i)

        nTexts = ini.value("text").toWzStringList(texts)
        StringVector_clear aViewData(v).textMsg
        For j = 0 To nTexts - 1
            t = Replace(texts(j), vbTab, "")                ' v->textMsg[j].remove("\t");
            ' (v->textMsg[j] = WzString::fromUtf8(_(...)): no translations)
            t = Replace(t, "%%", "%")
            StringVector_push_back aViewData(v).textMsg, t
        Next
        aViewData(v).typ = VIEW_RES
        aViewData(v).hasResData = True                      ' v->pData = r;
        If ini.contains("imdName") Then
            aViewData(v).res.pIMD = modelGet(ini.value("imdName").toWzString())
        End If
        If ini.contains("imdName2") Then
            aViewData(v).res.pIMD2 = modelGet(ini.value("imdName2").toWzString())
        End If
        If ini.contains("sequenceName") Then
            aViewData(v).res.sequenceName = ini.value("sequenceName").toWzString()
        End If
        If ini.contains("audioName") Then
            aViewData(v).res.audio = ini.value("audioName").toWzString()
        End If

        ini.endGroup
        apsViewData_set aViewData(v).name, v
    Next
    ptr = fileName ' so that cleanup function will be called on right data
    loadResearchViewData = True
End Function

' WzString *loadProximityViewData(const char *fileName): (VB) the WzString in ptr, False for nullptr
Public Function loadProximityViewData(ByVal fileName As String, ptr As String) As Boolean
    Dim ini As WzConfig, list() As String, n As Long, i As Long, v As Long, array_ As nlohmann_json, a As nlohmann_json, k As Long
    Dim msg As String, str_ As String, proxType As Long, audioID As Long, audioName As String
    If Not (PHYSFS_exists(fileName) <> 0) Then ASSERT_FAILURE "PHYSFS_exists(fileName)", "src/message.cpp:665", "loadProximityViewData", "%s not found", fileName: loadProximityViewData = False: Exit Function
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)

    ' "type": "wz2100.proxmsgs.v1"
    If Not ini.contains("type") Then
        debug_ 671, LOG_ERROR, "loadProximityViewData", "Missing required ""type"" key: %s", fileName
        loadProximityViewData = False
        Exit Function
    End If
    If ini.value("type").toWzString() <> "wz2100.proxmsgs.v1" Then
        debug_ 676, LOG_ERROR, "loadProximityViewData", "Unexpected ""type"" - expecting ""wz2100.proxmsgs.v1"": %s", fileName
        loadProximityViewData = False
        Exit Function
    End If

    ' msgs
    If Not ini.contains("msgs") Then
        debug_ 683, LOG_ERROR, "loadProximityViewData", "Missing required ""msgs"" key: %s", fileName
        loadProximityViewData = False
        Exit Function
    End If
    If Not ini.beginGroup("msgs") Then
        debug_ 688, LOG_ERROR, "loadProximityViewData", "Missing valid ""msgs"" value: %s", fileName
        loadProximityViewData = False
        Exit Function
    End If
    n = ini.childGroups(list)
    For i = 0 To n - 1
        ' Proximity viewdata init
        v = viewDataNew()
        aViewData(v).hasProxData = True         ' v->pData = r;
        aViewData(v).name = list(i)
        aViewData(v).fileName = fileName
        aViewData(v).typ = VIEW_PROX

        ini.beginGroup list(i)

        ' Set the message string when this Proximity blip gets clicked on
        If ini.contains("message") Then
            Set array_ = ini.json("message")
            If Not array_.is_null() And array_.is_array() Then
                For k = 0 To array_.size() - 1
                    Set a = array_.at_idx(k)
                    If Not a.is_string() Then
                        debug_ 714, LOG_ERROR, "loadProximityViewData", """message"" value is not a string: %s", fileName
                        GoTo continueMsg
                    End If
                    msg = a.get_string()
                    str_ = strresGetString(psStringRes, msg)
                    If Not (StrPtr(str_) <> 0) Then ASSERT_FAILURE "str", "src/message.cpp:719", "loadProximityViewData", "Cannot find the view data string with id ""%s""", msg
                    StringVector_push_back aViewData(v).textMsg, str_
continueMsg:
                Next
            Else
                msg = ini.value("message").toWzString()
                str_ = strresGetString(psStringRes, msg)
                If Not (StrPtr(str_) <> 0) Then ASSERT_FAILURE "str", "src/message.cpp:727", "loadProximityViewData", "Cannot find the view data string with id ""%s""", msg
                StringVector_push_back aViewData(v).textMsg, str_
            End If
        End If

        ' Read in the rest of the Proximity data
        If ini.contains("x") Then aViewData(v).prox_x = ini.value("x").toUInt()
        If ini.contains("y") Then aViewData(v).prox_y = ini.value("y").toUInt()
        If ini.contains("z") Then aViewData(v).prox_z = ini.value("z").toUInt()
        If ini.contains("type") Then
            proxType = ini.value("type").toUInt()
            If ULT(PROX_TYPES, proxType) Then
                ASSERT_FAILURE "false", "src/message.cpp:741", "loadProximityViewData", "Invalid proximity message sub type - %s", aViewData(v).name
                loadProximityViewData = False
                Exit Function
            End If
            aViewData(v).prox_proxType = proxType
        End If
        If ini.contains("audio") Then
            audioID = 0
            audioName = ini.value("audio").toWzString()
            If audioName = "0" Then
                audioID = NO_SOUND
            Else
                audioID = audio_GetIDFromStr(audioName)
                If audioID = NO_SOUND Then
                    ASSERT_FAILURE "false", "src/message.cpp:758", "loadProximityViewData", "couldn't get ID %d for weapon sound %s", audioID, audioName
                    loadProximityViewData = False
                    Exit Function
                End If
                If (audioID < 0 Or audioID > ID_MAX_SOUND) And audioID <> NO_SOUND Then
                    ASSERT_FAILURE "false", "src/message.cpp:763", "loadProximityViewData", "Invalid Weapon Sound ID - %d for weapon %s", audioID, audioName
                    loadProximityViewData = False
                    Exit Function
                End If
            End If
            aViewData(v).prox_audioID = audioID
        End If

        ini.endGroup
        apsViewData_set aViewData(v).name, v
    Next
    ini.endGroup ' "msgs"

    ptr = fileName ' so that cleanup function will be called on right data
    loadProximityViewData = True
End Function

' inline void from_json(const nlohmann::json& j, SEQ_DISPLAY& v) (the C++'s j["key"] on a const json: at)
Private Sub SEQ_DISPLAY_from_json(ByVal j As nlohmann_json, v As SEQ_DISPLAY)
    Dim subtitles As nlohmann_json, a As nlohmann_json, k As Long, msg As String, str_ As String
    v.sequenceName = j.at("video").get_string()
    debug_ 781, LOG_WZ, "from_json", "Sequence name: %s", v.sequenceName
    v.flag = j.at("loop").get_uint()
    debug_ 783, LOG_WZ, "from_json", "Sequence loop: %d", v.flag
    ' Set the subtitle string for the sequence.
    Set subtitles = j.at("subtitles")
    If Not subtitles.is_null() And subtitles.is_array() Then
        For k = 0 To subtitles.size() - 1
            Set a = subtitles.at_idx(k)
            msg = a.get_string()
            If Len(msg) <> 0 Then
                str_ = strresGetString(psStringRes, msg)
                If Not (StrPtr(str_) <> 0) Then ASSERT_FAILURE "str", "src/message.cpp:794", "from_json", "Cannot find the view data string with id ""%s""", msg
                StringVector_push_back v.textMsg, str_
                debug_ 796, LOG_WZ, "from_json", "Sequence subtitle array: %s", msg
            End If
        Next
    Else
        msg = subtitles.get_string()
        If Len(msg) <> 0 Then
            str_ = strresGetString(psStringRes, msg)
            If Not (StrPtr(str_) <> 0) Then ASSERT_FAILURE "str", "src/message.cpp:806", "from_json", "Cannot find the view data string with id ""%s""", msg
            StringVector_push_back v.textMsg, str_
            debug_ 808, LOG_WZ, "from_json", "Sequence subtitle string: %s", msg
        End If
    End If
End Sub

' (VB) the try block around briefObj["sequences"].get<std::vector<SEQ_DISPLAY>>(): False when it throws
Private Function seqListFromJson_try(ByVal arr As nlohmann_json, seqList() As SEQ_DISPLAY, n As Long) As Boolean
    Dim k As Long, z As SEQ_DISPLAY
    On Error GoTo threw
    n = 0
    If arr.size() > 0 Then ReDim seqList(0 To arr.size() - 1)
    For k = 0 To arr.size() - 1
        seqList(k) = z
        SEQ_DISPLAY_from_json arr.at_idx(k), seqList(k)
        n = n + 1
    Next
    seqListFromJson_try = True
    Exit Function
threw:
    seqListFromJson_try = False
End Function

' WzString *loadFlicViewData(const char *fileName): (VB) the WzString in ptr, False for nullptr
Public Function loadFlicViewData(ByVal fileName As String, ptr As String) As Boolean
    Dim jsonObj As nlohmann_json, it As nlohmann_json, briefsObj As nlohmann_json, briefObj As nlohmann_json, briefIt As nlohmann_json
    Dim k As Long, key As String, seqList() As SEQ_DISPLAY, nSeq As Long, v As Long
    If Not (PHYSFS_exists(fileName) <> 0) Then ASSERT_FAILURE "PHYSFS_exists(fileName)", "src/message.cpp:820", "loadFlicViewData", "%s not found", fileName: loadFlicViewData = False: Exit Function

    Set jsonObj = wzLoadJsonObjectFromFile(fileName)
    If jsonObj Is Nothing Then
        debug_ 825, LOG_ERROR, "loadFlicViewData", "Failed to load JSON: %s", fileName
        loadFlicViewData = False
        Exit Function
    End If

    ' "type": "wz2100.briefs.v1"
    Set it = jsonObj.find("type")
    If it Is Nothing Then
        debug_ 834, LOG_ERROR, "loadFlicViewData", "Missing required ""type"" key: %s", fileName
        loadFlicViewData = False
        Exit Function
    End If
    If Not it.is_string() Then
        debug_ 839, LOG_ERROR, "loadFlicViewData", "Unexpected ""type"" - expecting ""wz2100.briefs.v1"": %s", fileName
        loadFlicViewData = False
        Exit Function
    End If
    If it.get_string() <> "wz2100.briefs.v1" Then
        debug_ 839, LOG_ERROR, "loadFlicViewData", "Unexpected ""type"" - expecting ""wz2100.briefs.v1"": %s", fileName
        loadFlicViewData = False
        Exit Function
    End If

    ' "briefs"
    Set it = jsonObj.find("briefs")
    If it Is Nothing Then
        debug_ 847, LOG_ERROR, "loadFlicViewData", "Missing required ""briefs"" key: %s", fileName
        loadFlicViewData = False
        Exit Function
    End If
    If Not it.is_object() Then
        debug_ 852, LOG_ERROR, "loadFlicViewData", """briefs"" value is not an object: %s", fileName
        loadFlicViewData = False
        Exit Function
    End If
    Set briefsObj = it

    For k = 0 To briefsObj.size() - 1
        key = briefsObj.key_at(k)
        debug_ 859, LOG_WZ, "loadFlicViewData", "Sequence video set: %s", key

        Set briefObj = briefsObj.value_at(k)
        If Not briefObj.is_object() Then
            debug_ 864, LOG_ERROR, "loadFlicViewData", """briefs""[%s] value is not an object: %s", key, fileName
            GoTo continue_
        End If

        Set briefIt = briefObj.find("sequences")
        If briefIt Is Nothing Then
            debug_ 871, LOG_ERROR, "loadFlicViewData", """briefs""[%s] lacks a ""sequences"" property: %s", key, fileName
            GoTo continue_
        End If
        If Not briefIt.is_array() Then
            debug_ 876, LOG_ERROR, "loadFlicViewData", """briefs""[%s] ""sequences"" property must have an array value: %s", key, fileName
            GoTo continue_
        End If
        Erase seqList
        If Not seqListFromJson_try(briefIt, seqList, nSeq) Then
            debug_ 885, LOG_ERROR, "loadFlicViewData", "Failed to parse ""briefs""[%s] ""sequences"" property value: %s", key, fileName
            GoTo continue_
        End If
        debug_ 888, LOG_WZ, "loadFlicViewData", "Sequence list size: %zu", nSeq

        ' Replay viewdata init
        v = viewDataNew()
        aViewData(v).hasReplayData = True       ' v->pData = r;
        aViewData(v).fileName = fileName
        aViewData(v).typ = VIEW_RPL

        aViewData(v).name = key
        aViewData(v).replay_seqList = seqList   ' r->seqList = std::move(seqList);
        aViewData(v).replay_seqList_size = nSeq

        apsViewData_set aViewData(v).name, v
continue_:
    Next
    ptr = fileName ' so that cleanup function will be called on right data
    loadFlicViewData = True
End Function

' Get the view data identified by the name
' VIEWDATA *getViewData(const WzString &name): (VB) the handle, 0 for nullptr
Public Function getViewData(ByVal name_ As String) As Long
    Dim i As Long, ptr As Long, iter As Long
    If vbDisplayOff Then getViewData = 0: Exit Function  ' (VB) the vbtest runs: statsref's stand-in (no view data loaded)
    i = apsViewData_find(name_)
    If i <> -1 Then ptr = apsViewData_ptrs(i) Else ptr = 0
    If ptr = 0 Then ' dump for debugging
        For iter = 0 To apsViewData_size - 1
            debug_ 920, LOG_WZ, "getViewData", vbTab & "%s", aViewData(apsViewData_ptrs(iter)).name
        Next
    End If
    If Not (ptr <> 0) Then ASSERT_FAILURE "ptr", "src/message.cpp:924", "getViewData", "Message %s not found, run with --debug=wz to get a list of all known messages", name_
    getViewData = ptr
End Function

' std::vector<WzString> getViewDataKeys(): (VB) the names in keys(0 .. n - 1), n returned
Public Function getViewDataKeys(keys() As String) As Long
    Dim i As Long
    If apsViewData_size > 0 Then
        ReDim keys(0 To apsViewData_size - 1)
        For i = 0 To apsViewData_size - 1
            keys(i) = apsViewData_names(i)
        Next
    Else
        Erase keys
    End If
    getViewDataKeys = apsViewData_size
End Function

' Release the message heaps
Public Function messageShutdown() As Boolean
    freeMessages
    messageShutdown = True
End Function

'check for any messages using this viewdata and remove them
Private Sub checkMessages(ByVal psViewData As Long)
    Dim removedAMessage As Boolean, i As Long, it As Long, psCurr As Long
    removedAMessage = False
    For i = 0 To MAX_PLAYERS - 1
        ' mutating_list_iterate: the next node is taken before the message is removed
        it = apsMessages(i).head
        Do While it <> 0
            psCurr = lnObj(it)
            it = lnNext(it)
            If aMessage(psCurr).pViewData = psViewData Then
                removeMessage psCurr, i
                removedAMessage = True
            End If
        Loop
    Next

    If removedAMessage Then
        jsDebugMessageUpdate
    End If
End Sub

' void releaseAllFlicMessages(PerPlayerMessageLists& list)
Public Sub releaseAllFlicMessages(list() As OBJ_LIST)
    Dim i As Long, it As Long, psCurr As Long
    ' Iterate through all players' message lists
    For i = 0 To MAX_PLAYERS - 1
        ' Iterate through all messages in list
        it = list(i).head
        Do While it <> 0
            psCurr = lnObj(it)
            it = lnNext(it)
            If aMessage(psCurr).typ = MSG_MISSION Or aMessage(psCurr).typ = MSG_CAMPAIGN Then
                removeMessage psCurr, i
            End If
        Loop
    Next
End Sub

' Release the viewdata memory
Public Sub viewDataShutDown(ByVal fileName As String)
    Dim iter As Long, psViewData As Long
    debug_ 989, LOG_MSG, "viewDataShutDown", "calling shutdown for %s", fileName
    iter = 0
    Do While iter < apsViewData_size
        psViewData = apsViewData_ptrs(iter)

        If StrComp(aViewData(psViewData).fileName, fileName, vbBinaryCompare) <> 0 Then
            iter = iter + 1
            GoTo continue_ ' do not delete this now
        End If
        If Not releaseObjectives And (aViewData(psViewData).typ = VIEW_RPL Or aViewData(psViewData).typ = VIEW_RPLX) Then
            iter = iter + 1
            GoTo continue_
        End If

        ' check for any messages using this viewdata
        checkMessages psViewData
        viewDataDelete psViewData           ' delete psViewData->pData; delete psViewData;
        apsViewData_erase iter              ' iter = apsViewData.erase(iter);
continue_:
    Loop
End Sub

' 'displays' a proximity display
Public Sub displayProximityMessage(ByVal psProxDisp As Long)
    Dim psViewData As Long, i As Long, psFeature As Long, psMessage As Long
    psMessage = aProxDisp(psProxDisp).psMessage
    If aProxDisp(psProxDisp).typ = POS_PROXDATA Then
        If aMessage(psMessage).pViewData = 0 Then
            Exit Sub ' if no data - ignore message
        End If

        psViewData = aMessage(psMessage).pViewData
        ' (const VIEW_PROXIMITY *psViewProx = (VIEW_PROXIMITY *)psViewData->pData: its prox_* fields)

        'display text - if any
        If aViewData(psViewData).textMsg.size <> 0 And aViewData(psViewData).typ <> VIEW_BEACON Then
            For i = 0 To aViewData(psViewData).textMsg.size - 1
                addConsoleMessage aViewData(psViewData).textMsg.items(i), DEFAULT_JUSTIFY, SYSTEM_MESSAGE
            Next
        End If

        'play message - if any
        If aViewData(psViewData).prox_audioID <> NO_AUDIO_MSG Then
            audio_QueueTrackPos aViewData(psViewData).prox_audioID, aViewData(psViewData).prox_x, aViewData(psViewData).prox_y, aViewData(psViewData).prox_z
        End If
    ElseIf aProxDisp(psProxDisp).typ = POS_PROXOBJ Then
        If Not (aMessage(psMessage).psObj <> 0) Then ASSERT_FAILURE "psProxDisp->psMessage->psObj", "src/message.cpp:1079", "displayProximityMessage", "Invalid proxobj - null object": Exit Sub
        If Not (isFeature(aMessage(psMessage).psObj)) Then ASSERT_FAILURE "psProxDisp->psMessage->psObj->type == OBJ_FEATURE", "src/message.cpp:1080", "displayProximityMessage", "Invalid proxobj - must be feature": Exit Sub
        psFeature = aMessage(psMessage).psObj

        If asFeatureStats(aFeature(psFeature).psStats).subType = FEAT_OIL_RESOURCE Then
            'play default audio message for oil resource
            audio_QueueTrackPos ID_SOUND_RESOURCE_HERE, aFeature(psFeature).pos.x, aFeature(psFeature).pos.y, aFeature(psFeature).pos.z
        ElseIf asFeatureStats(aFeature(psFeature).psStats).subType = FEAT_GEN_ARTE Then
            'play default audio message for artefact
            audio_QueueTrackPos ID_SOUND_ARTIFACT, aFeature(psFeature).pos.x, aFeature(psFeature).pos.y, aFeature(psFeature).pos.z
        End If
    End If

    'set the read flag
    aMessage(psMessage).read = True
End Sub

Public Sub cleanupOldBeaconMessages()
    Dim removedAMessage As Boolean, i As Long, it As Long, psProxDisp As Long, psCurrMsg As Long, pViewData As Long
    ' Check if it's time to remove beacons
    removedAMessage = False
    For i = 0 To MAX_PLAYERS - 1
        ' Go through all the proximity Displays
        ' (mutating_list_iterate: the next node is taken before the display may be removed)
        it = apsProxDisp(i).head
        Do While it <> 0
            psProxDisp = lnObj(it)
            it = lnNext(it)
            If aMessage(aProxDisp(psProxDisp).psMessage).dataType = MSG_DATA_BEACON Then
                psCurrMsg = aProxDisp(psProxDisp).psMessage
                pViewData = aMessage(psCurrMsg).pViewData

                If Not (pViewData <> 0) Then ASSERT_FAILURE "pViewData != nullptr", "src/message.cpp:1113", "cleanupOldBeaconMessages", "Message without data!": GoTo continue_

                If aViewData(pViewData).typ = VIEW_BEACON Then
                    If Not (aViewData(pViewData).hasProxData) Then ASSERT_FAILURE "pViewData->pData != nullptr", "src/message.cpp:1117", "cleanupOldBeaconMessages", "Help message without data!": GoTo continue_
                    If aViewData(pViewData).hasProxData Then
                        If U32(aViewData(pViewData).prox_timeAdded) + 60000# <= U32(gameTime) Then
                            debug_ 1120, LOG_MSG, "cleanupOldBeaconMessages", "blip timeout for %d, from %d", i, aViewData(pViewData).prox_sender
                            removeMessage psCurrMsg, i  'remove beacon
                            removedAMessage = True
                            Exit Do                     ' return IterationResult::BREAK_ITERATION; //there can only be 1 beacon per player
                        End If
                    End If
                End If
            End If
continue_:
        Loop
    Next
    If removedAMessage Then
        jsDebugMessageUpdate
    End If
End Sub

' (VIEW_RESEARCH *)psViewData->pData (an empty VIEW_RESEARCH when it is none)
Public Function VIEWDATA_research(ByVal psViewData As Long) As VIEW_RESEARCH
    If aViewData(psViewData).hasResData Then VIEWDATA_research = aViewData(psViewData).res
End Function

' ((VIEW_REPLAY *)psViewData->pData)->seqList.size() (0 when it is none)
Public Function VIEWDATA_replaySeqCount(ByVal psViewData As Long) As Long
    If aViewData(psViewData).hasReplayData Then VIEWDATA_replaySeqCount = aViewData(psViewData).replay_seqList_size
End Function

' ((VIEW_REPLAY *)psViewData->pData)->seqList.at(i)
Public Function VIEWDATA_replaySeq(ByVal psViewData As Long, ByVal i As Long) As SEQ_DISPLAY
    If Not (ULT(i, aViewData(psViewData).replay_seqList_size)) Then Err.Raise vbObjectError + 400, "std::out_of_range", "vector::_M_range_check"
    VIEWDATA_replaySeq = aViewData(psViewData).replay_seqList(i)
End Function

'--------------------------------------------------------------------------
' (VB) the fields read and written through a MESSAGE *, PROXIMITY_DISPLAY *, VIEWDATA * and its VIEW_PROXIMITY *

Public Property Get MESSAGE_psObj(ByVal psMessage As Long) As Long
    MESSAGE_psObj = aMessage(psMessage).psObj
End Property

Public Property Let MESSAGE_psObj(ByVal psMessage As Long, ByVal v As Long)
    aMessage(psMessage).psObj = v
End Property

Public Property Get MESSAGE_pViewData(ByVal psMessage As Long) As Long
    MESSAGE_pViewData = aMessage(psMessage).pViewData
End Property

Public Property Let MESSAGE_pViewData(ByVal psMessage As Long, ByVal v As Long)
    aMessage(psMessage).pViewData = v
End Property

Public Property Get PROXIMITY_DISPLAY_psMessage(ByVal psProxDisp As Long) As Long
    PROXIMITY_DISPLAY_psMessage = aProxDisp(psProxDisp).psMessage
End Property

Public Property Get VIEWDATA_name(ByVal psViewData As Long) As String
    VIEWDATA_name = aViewData(psViewData).name
End Property

Public Property Get VIEW_PROXIMITY_x(ByVal psViewData As Long) As Long
    VIEW_PROXIMITY_x = aViewData(psViewData).prox_x
End Property

Public Property Get VIEW_PROXIMITY_y(ByVal psViewData As Long) As Long
    VIEW_PROXIMITY_y = aViewData(psViewData).prox_y
End Property

Public Property Get VIEW_PROXIMITY_z(ByVal psViewData As Long) As Long
    VIEW_PROXIMITY_z = aViewData(psViewData).prox_z
End Property

Public Property Let VIEW_PROXIMITY_z(ByVal psViewData As Long, ByVal v As Long)
    aViewData(psViewData).prox_z = v
End Property
