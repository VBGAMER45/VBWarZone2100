Attribute VB_Name = "console_c"
' src/console.cpp, console.h: the in-game console (the messages at the top of
' the screen, the history, the info line), whole. The std::deque<CONSOLE_MESSAGE>
' lists are arrays with their sizes; a CONSOLE_MESSAGE holds its text as a
' WzText (GfxApi). CONSOLE_CALC_LAYOUT_FUNC and CONSOLE_MESSAGE_LISTENER are
' WzFuncs: a listener is called with the ConsoleMessage's fields (sender,
' text, justification, team, duration: multiint_lobby_c reads (sender, text)). A DEBOUNCED_MESSAGE is its debounce
' time (CANNOT_BUILD_BURNING = 2500), which also tells the messages apart: the
' only one there is. CONPRINTF's and console()'s format and arguments go
' through printf_vb's vsprintf.
' (VB) vbDisplayOff (display_c; the vbtest runs set it): addConsoleMessage
' returns true at once, as statsref's stand-in does.
Option Explicit

Public Const MAX_CONSOLE_MESSAGES As Long = 64
Public Const MAX_CONSOLE_STRING_LENGTH As Long = 255
Public Const MAX_CONSOLE_TMP_STRING_LENGTH As Long = 255

Public Enum CONSOLE_TEXT_JUSTIFICATION
    LEFT_JUSTIFY
    DEFAULT_JUSTIFY = LEFT_JUSTIFY
    RIGHT_JUSTIFY
    CENTRE_JUSTIFY
End Enum

' const DEBOUNCED_MESSAGE CANNOT_BUILD_BURNING = {2500}; (its debounceTime)
Public Const CANNOT_BUILD_BURNING As Long = 2500

' ID to use for addConsoleMessage() in case of a system message
Public Const SYSTEM_MESSAGE As Long = -1
Public Const NOTIFY_MESSAGE As Long = -2        ' mainly used for lobby & error messages
Public Const INFO_MESSAGE As Long = -3          ' This type is not stored, it is used for simple messages
Public Const MAX_CONSOLE_MESSAGE_DURATION As Long = -1  ' ((UDWORD)-1)
Public Const DEFAULT_CONSOLE_MESSAGE_DURATION As Long = 0

' FIXME: When we switch over to full JS, use class version of this file

Private Const DEFAULT_MESSAGE_DURATION As Long = GAME_TICKS_PER_SEC * 5
Private Const DEFAULT_MESSAGE_DURATION_CAMPAIGN As Long = GAME_TICKS_PER_SEC * 12
' Chat/history "window"
Private Const CON_BORDER_WIDTH As Long = 4
Private Const CON_BORDER_HEIGHT As Long = 4
' HISTORYBOX_X is RET_X, HISTORYBOX_Y is RET_Y - 80 (hci.h's: RET_Y depends on the screen's height)
Private Const NumDisplayLines As Long = 4

Private Type CONSOLE_                 ' struct CONSOLE (VB: console() is the Sub)
    topX As Long                        ' UDWORD
    topY As Long
    width As Long
    textDepth As Long
    permanent As Boolean
End Type

Private Type CONSOLE_MESSAGE
    display As WzText
    timeAdded As Long                   ' UDWORD: When was it added to our list?
    duration As Long                    ' UDWORD
    JustifyType As CONSOLE_TEXT_JUSTIFICATION ' text justification
    player As Long                      ' Player who sent this message or SYSTEM_MESSAGE
End Type

' (VB) the deques: msg_ActiveMessages(0 .. msg_ActiveMessages_size - 1) and so on
Private ActiveMessages() As CONSOLE_MESSAGE     ' we add all messages to this container
Private ActiveMessages_size As Long
Private TeamMessages() As CONSOLE_MESSAGE       ' history of team/private communications
Private TeamMessages_size As Long
Private HistoryMessages() As CONSOLE_MESSAGE    ' history of all other communications
Private HistoryMessages_size As Long
Private InfoMessages() As CONSOLE_MESSAGE
Private InfoMessages_size As Long
Private bConsoleDropped As Boolean      ' = false: Is the console history on or off?
Private HistoryMode As Boolean          ' = false: toggle between team & global history
Private updatepos As Long               ' = 0: if user wants to scroll back the history log
Private linePitch As Long               ' = 0: the pitch of a line
Private showBackgroundColor As Boolean  ' = false: if user wants to add more contrast to the history display
Private mainConsole As CONSOLE_          ' Stores the console dimensions and states
Private historyConsole As CONSOLE_       ' Stores the History console dimensions and states
Private messageDuration As Long         ' UDWORD: How long do messages last for?
Private bTextBoxActive As Boolean       ' = false: Is there a box under the console text?
Private bConsoleDisplayEnabled As Boolean ' = false: Is the console being displayed?
Private consoleVisibleLines As Long     ' UDWORD: How many lines are displayed?
Private allowNewMessages As Long        ' int: Whether new messages are allowed to be added
Private messageListeners As Collection  ' std::set<std::shared_ptr<CONSOLE_MESSAGE_LISTENER>>: WzFuncs

Public ConsoleString As String          ' char ConsoleString[MAX_CONSOLE_TMP_STRING_LENGTH]: Globally available string for new console messages.

Private calcLayoutFunc As WzFunc        ' static CONSOLE_CALC_LAYOUT_FUNC calcLayoutFunc;

' addConsoleMessageDebounced's static std::map<const DEBOUNCED_MESSAGE *, std::chrono::steady_clock::time_point>
Private lastAllowedMessageTimes As Collection ' key: the message's debounce time; item: widgSteadyNowMs's time

' (VB) the lambdas
Public Const FN_initConsoleMessages_layout As Long = 11000

'--------------------------------------------------------------------------
' (VB) the deques' operations

Private Sub msgs_push_back(msgs() As CONSOLE_MESSAGE, n As Long, ByVal text As String, ByVal fontID As iV_fonts, ByVal time_ As Long, ByVal duration As Long, ByVal justify As CONSOLE_TEXT_JUSTIFICATION, ByVal plr As Long)
    If n = 0 Then
        ReDim msgs(0 To 15)
    ElseIf n > UBound(msgs) Then
        ReDim Preserve msgs(0 To 2 * n - 1)
    End If
    ' CONSOLE_MESSAGE(const std::string &text, iV_fonts fontID, UDWORD time, UDWORD duration, CONSOLE_TEXT_JUSTIFICATION justify, int plr)
    Set msgs(n).display = WzText_new(text, fontID)
    msgs(n).timeAdded = time_
    msgs(n).duration = duration
    msgs(n).JustifyType = justify
    msgs(n).player = plr
    n = n + 1
End Sub

Private Sub msgs_erase(msgs() As CONSOLE_MESSAGE, n As Long, ByVal i As Long)
    Dim k As Long
    For k = i To n - 2
        msgs(k) = msgs(k + 1)
    Next
    Set msgs(n - 1).display = Nothing
    n = n - 1
End Sub

Private Sub msgs_clear(msgs() As CONSOLE_MESSAGE, n As Long)
    Erase msgs
    n = 0
End Sub

'--------------------------------------------------------------------------

Public Sub consoleAddMessageListener(ByVal listener As WzFunc)
    Dim f As WzFunc
    If messageListeners Is Nothing Then Set messageListeners = New Collection
    For Each f In messageListeners
        If f Is listener Then Exit Sub  ' a std::set
    Next
    messageListeners.Add listener
End Sub

Public Sub consoleRemoveMessageListener(ByVal listener As WzFunc)
    Dim i As Long
    If messageListeners Is Nothing Then Exit Sub
    For i = messageListeners.count To 1 Step -1
        If messageListeners(i) Is listener Then messageListeners.Remove i
    Next
End Sub

'
'	Specify how long messages will stay on screen.
'
Private Sub setConsoleMessageDuration(ByVal time_ As Long)
    messageDuration = time_
End Sub

Public Sub setConsoleCalcLayout(ByVal layoutFunc As WzFunc)
    Set calcLayoutFunc = layoutFunc
    If Not calcLayoutFunc Is Nothing Then
        WzFunc_call calcLayoutFunc
    End If
End Sub

' Sets the system up
Public Sub initConsoleMessages()
    Dim duration As Long
    If game.typ = LEVEL_TYPE_SKIRMISH Then duration = DEFAULT_MESSAGE_DURATION Else duration = DEFAULT_MESSAGE_DURATION_CAMPAIGN
    linePitch = iV_GetTextLineSize(font_regular)
    bConsoleDropped = False
    setConsoleMessageDuration duration                      ' Setup how long messages are displayed for
    setConsoleBackdropStatus True                           ' No box under the text
    enableConsoleDisplay True                               ' Turn on the console display

    '	Set up the main console size and position x,y,width
    setConsoleCalcLayout WzFunc_new(FN_initConsoleMessages_layout)
    setConsoleLineInfo MAX_CONSOLE_MESSAGES \ 4 + 4
    setConsolePermanence False, True                        ' We're not initially having permanent messages
    permitNewConsoleMessages True                           ' Allow new messages
End Sub

Public Sub shutdownConsoleMessages()
    permitNewConsoleMessages False
    flushConsoleMessages
    clearInfoMessages
End Sub

Public Sub consoleScreenDidChangeSize(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If calcLayoutFunc Is Nothing Then Exit Sub
    WzFunc_call calcLayoutFunc
End Sub

' toggle between team & global history
Public Sub setHistoryMode(ByVal mode As Boolean)
    HistoryMode = mode
End Sub

' Open the console when it's closed and close it when it's open.
Public Sub toggleConsoleDrop()
    If Not bConsoleDropped Then
        ' it was closed, so play open sound
        bConsoleDropped = True
        audio_PlayTrack ID_SOUND_WINDOWOPEN
    Else
        ' play closing sound
        audio_PlayTrack ID_SOUND_WINDOWCLOSE
        bConsoleDropped = False
    End If
End Sub

' debouncedMessage: the DEBOUNCED_MESSAGE (its debounceTime, e.g. CANNOT_BUILD_BURNING)
Public Function addConsoleMessageDebounced(ByVal Text As String, ByVal jusType As CONSOLE_TEXT_JUSTIFICATION, ByVal player As Long, ByVal debouncedMessage As Long, Optional ByVal team As Boolean = False, Optional ByVal duration As Long = DEFAULT_CONSOLE_MESSAGE_DURATION) As Boolean
    ' Messages are debounced individually - one debounced message won't stop a different one from appearing
    Dim DEBOUNCE_TIME As Double, now As Double, lastAllowedMessageTime As Double, haveTime As Boolean, key As String

    DEBOUNCE_TIME = debouncedMessage    ' std::chrono::milliseconds DEBOUNCE_TIME(message.debounceTime)
    now = widgSteadyNowMs()             ' std::chrono::steady_clock::now()

    If lastAllowedMessageTimes Is Nothing Then Set lastAllowedMessageTimes = New Collection
    key = "m" & CStr(debouncedMessage)
    On Error Resume Next
    lastAllowedMessageTime = lastAllowedMessageTimes(key)
    haveTime = (Err.Number = 0)
    On Error GoTo 0

    If Not haveTime Or Fix(now - lastAllowedMessageTime) >= DEBOUNCE_TIME Then
        If haveTime Then lastAllowedMessageTimes.Remove key
        lastAllowedMessageTimes.Add now, key
        addConsoleMessageDebounced = addConsoleMessage(Text, jusType, player, team, duration)
    Else
        addConsoleMessageDebounced = False
    End If
End Function

' Add a string to the console.
Public Function addConsoleMessage(ByVal Text As String, ByVal jusType As CONSOLE_TEXT_JUSTIFICATION, ByVal player As Long, Optional ByVal team As Boolean = False, Optional ByVal duration As Long = DEFAULT_CONSOLE_MESSAGE_DURATION) As Boolean
    Dim listener As WzFunc, lines() As String, li As Long, FitText As String, pixelWidth As Long, newMsgDuration As Long
    If vbDisplayOff Then addConsoleMessage = True: Exit Function ' (VB) the vbtest runs: statsref's stand-in

    ' ConsoleMessage message = { Text, jusType, player, team, duration };
    If Not messageListeners Is Nothing Then
        For Each listener In messageListeners
            WzFunc_call listener, player, Text, jusType, team, duration
        Next
    End If

    If allowNewMessages = 0 Then
        addConsoleMessage = False       ' Don't allow it to be added if we've disabled adding of new messages
        Exit Function
    End If

    ' std::istringstream stream(Text); while (std::getline(stream, lines)): the lines split at \n (a last
    ' empty one, after a final \n, is not a line)
    lines = Split(Text, vbLf)
    For li = 0 To UBound(lines)
        If li = UBound(lines) And Len(lines(li)) = 0 And li > 0 Then Exit For
        ' We got one "line" from the total string, now we must check
        ' to see if it fits, if not, we truncate it. (Full text is in the logs)
        ' NOTE: may want to break up line into multi-line so it matches the log
        FitText = lines(li)
        Do While Len(FitText) > 0
            pixelWidth = iV_GetTextWidth(FitText, font_regular)
            If Not ULT(mainConsole.width, pixelWidth) Then
                Exit Do
            End If
            FitText = Left$(FitText, Len(FitText) - 1)  ' Erase last char.
        Loop

        debug_ 251, LOG_CONSOLE, "addConsoleMessage", "(to player %d): %s", player, FitText

        If duration <> DEFAULT_CONSOLE_MESSAGE_DURATION Then newMsgDuration = duration Else newMsgDuration = messageDuration

        If player = INFO_MESSAGE Then
            msgs_push_back InfoMessages, InfoMessages_size, FitText, font_regular, realTime, newMsgDuration, jusType, player
        Else
            msgs_push_back ActiveMessages, ActiveMessages_size, FitText, font_regular, realTime, newMsgDuration, jusType, player ' everything gets logged here for a specific period of time
            If team Then
                msgs_push_back TeamMessages, TeamMessages_size, FitText, font_regular, realTime, newMsgDuration, jusType, player ' persistent team specific logs
            End If
            msgs_push_back HistoryMessages, HistoryMessages_size, FitText, font_regular, realTime, newMsgDuration, jusType, player ' persistent messages (all types)
        End If
    Next

    addConsoleMessage = True
End Function

' \return The number of active console messages
Public Function getNumberConsoleMessages() As Long
    getNumberConsoleMessages = ActiveMessages_size
End Function

' Update the console messages.
'	This function will remove messages that are overdue.
Public Sub updateConsoleMessages()
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    Dim i As Long
    ' If there are no messages or we're on permanent (usually for scripts) then exit
    If ((getNumberConsoleMessages() = 0) And InfoMessages_size = 0) Or mainConsole.permanent Then
        Exit Sub
    End If
    i = 0
    Do While i < InfoMessages_size
        If (InfoMessages(i).duration <> MAX_CONSOLE_MESSAGE_DURATION) And ULT(InfoMessages(i).duration, FromU32(U32(realTime) - U32(InfoMessages(i).timeAdded))) Then
            msgs_erase InfoMessages, InfoMessages_size, i
        Else
            i = i + 1
        End If
    Loop
    ' Time to kill all expired ones
    i = 0
    Do While i < ActiveMessages_size
        If (ActiveMessages(i).duration <> MAX_CONSOLE_MESSAGE_DURATION) And ULT(ActiveMessages(i).duration, FromU32(U32(realTime) - U32(ActiveMessages(i).timeAdded))) Then
            msgs_erase ActiveMessages, ActiveMessages_size, i
        Else
            i = i + 1
        End If
    Loop
End Sub

'
'	Remove the top message on screen.
'	This and setConsoleMessageDuration should be sufficient to allow
'	us to put up messages that stay there until we remove them
'	ourselves - be sure and reset message duration afterwards
'
Public Sub removeTopConsoleMessage()
    If getNumberConsoleMessages() <> 0 Then
        msgs_erase ActiveMessages, ActiveMessages_size, 0
    End If
End Sub

' Clears just Active console messages
Public Sub clearActiveConsole()
    msgs_clear ActiveMessages, ActiveMessages_size
End Sub

' Clears all console messages
Public Sub flushConsoleMessages()
    msgs_clear ActiveMessages, ActiveMessages_size
    msgs_clear TeamMessages, TeamMessages_size
    msgs_clear HistoryMessages, HistoryMessages_size
End Sub

' Sets console text color depending on message type
Public Function getConsoleTextColor(ByVal player As Long) As Long
    ' System messages
    If player = SYSTEM_MESSAGE Then
        getConsoleTextColor = psPalette(WZCOL_CONS_TEXT_SYSTEM)
    ElseIf player = NOTIFY_MESSAGE Then
        getConsoleTextColor = psPalette(WZCOL_YELLOW)
    ElseIf player = INFO_MESSAGE Then
        getConsoleTextColor = psPalette(WZCOL_CONS_TEXT_INFO)
    ElseIf isSpectatorPlayer(player) Then
        getConsoleTextColor = psPalette(WZCOL_TEXT_MEDIUM)
    Else
        ' Only use friend-foe colors if we are (potentially) a player
        If ULT(selectedPlayer, MAX_PLAYERS) Then
            ' Don't use friend-foe colors in the lobby
            If bEnemyAllyRadarColor And (GetGameMode() = GS_NORMAL) Then
                If aiCheckAlliances(player, selectedPlayer) Then
                    If selectedPlayer = player Then
                        getConsoleTextColor = psPalette(WZCOL_TEXT_BRIGHT)
                        Exit Function
                    End If
                    getConsoleTextColor = psPalette(WZCOL_CONS_TEXT_USER_ALLY)
                    Exit Function
                Else
                    getConsoleTextColor = psPalette(WZCOL_CONS_TEXT_USER_ENEMY)
                    Exit Function
                End If
            End If
        End If

        getConsoleTextColor = psPalette(WZCOL_TEXT_BRIGHT)
    End If
End Function

' (VB) player >= 0 && player < NetPlay.players.size() && NetPlay.players[player].isSpectator, its && in order
Private Function isSpectatorPlayer(ByVal player As Long) As Boolean
    If player >= 0 Then
        If player < NetPlay_players_size() Then
            isSpectatorPlayer = NetPlay_players_isSpectator(player)
        End If
    End If
End Function

Private Sub console_drawtext(ByVal display As WzText, ByVal colour As Long, ByVal x As Long, ByVal y As Long, ByVal justify As CONSOLE_TEXT_JUSTIFICATION, ByVal width As Long)
    Select Case justify
    Case LEFT_JUSTIFY
        ' do nothing
    Case RIGHT_JUSTIFY
        x = x + width - display.width()
    Case CENTRE_JUSTIFY
        x = x + (width - display.width()) \ 2
    End Select
    display.render x, y, colour
End Sub

' Show global (mode=false) or team (mode=true) history messages
Public Sub displayOldMessages(ByVal mode As Boolean)
    Dim startpos As Long, count As Long, i As Long, nudgeright As Long, TextYpos As Long, colour As Long
    Dim n As Long

    ' std::deque<CONSOLE_MESSAGE> *WhichMessages: TeamMessages (mode) or HistoryMessages
    If mode Then n = TeamMessages_size Else n = HistoryMessages_size
    If n <> 0 Then
        count = n                       ' total number of messages
        If count > NumDisplayLines Then ' if we have more than we can display
            startpos = count - NumDisplayLines  ' show last X lines
            startpos = startpos + updatepos ' unless user wants to start at something else
            If startpos < 0 Then        ' don't underflow
                startpos = 0
                updatepos = (count - NumDisplayLines) * -1 ' reset back to first entry
                count = NumDisplayLines
            ElseIf Not ULT(count, FromU32(U32(count) + U32(updatepos))) Then
                count = FromU32(U32(count) + U32(updatepos)) ' user may want something different
            Else
                ' reset all, we got overflow
                count = n
                updatepos = 0
                startpos = count - NumDisplayLines
            End If
        End If

        nudgeright = 0
        TextYpos = historyConsole.topY + linePitch - 2

        If isSecondaryWindowUp() Then   ' see if (build/research/...)window is up
            nudgeright = RET_FORMWIDTH + 2 ' move text over
        End If
        ' if user wants to add a bit more contrast to the text
        If showBackgroundColor Then
            iV_TransBoxFill historyConsole.topX + nudgeright - CON_BORDER_WIDTH, historyConsole.topY - historyConsole.textDepth - CON_BORDER_HEIGHT, _
                            historyConsole.topX + historyConsole.width, historyConsole.topY + (NumDisplayLines * linePitch) + CON_BORDER_HEIGHT
        End If
        For i = startpos To count - 1
            If mode Then
                colour = psPalette(WZCOL_CONS_TEXT_USER_ALLY)
                console_drawtext TeamMessages(i).display, colour, historyConsole.topX + nudgeright, TextYpos, TeamMessages(i).JustifyType, historyConsole.width
                TextYpos = TextYpos + TeamMessages(i).display.lineSize()
            Else
                colour = getConsoleTextColor(HistoryMessages(i).player)
                console_drawtext HistoryMessages(i).display, colour, historyConsole.topX + nudgeright, TextYpos, HistoryMessages(i).JustifyType, historyConsole.width
                TextYpos = TextYpos + HistoryMessages(i).display.lineSize()
            End If
        Next
    End If
End Sub

' Displays all the console messages
Public Sub displayConsoleMessages()
    Dim i As Long, tmp As Long, TextYpos As Long
    ' Check if we have any messages we want to show
    If ActiveMessages_size = 0 And Not bConsoleDropped And InfoMessages_size = 0 Then
        Exit Sub
    End If

    ' scripts can disable the console
    If Not bConsoleDisplayEnabled And InfoMessages_size = 0 Then
        Exit Sub
    End If

    pie_SetFogStatus False

    If bConsoleDropped Then
        displayOldMessages HistoryMode
    End If

    If InfoMessages_size <> 0 Then
        i = InfoMessages_size - 1       ' we can only show the last one...
        tmp = pie_GetVideoBufferWidth()
        drawBlueBox 0, 0, tmp, 18
        tmp = tmp - InfoMessages(i).display.width()
        console_drawtext InfoMessages(i).display, getConsoleTextColor(InfoMessages(i).player), tmp - 6, linePitch - 2, InfoMessages(i).JustifyType, InfoMessages(i).display.width()
    End If

    If ActiveMessages_size <> 0 Then
        TextYpos = mainConsole.topY
        ' Draw the blue background for the text (only in game, not lobby)
        If bTextBoxActive And GetGameMode() = GS_NORMAL Then
            iV_TransBoxFill mainConsole.topX - CON_BORDER_WIDTH, _
                            mainConsole.topY - mainConsole.textDepth - CON_BORDER_HEIGHT, _
                            mainConsole.topX + mainConsole.width, _
                            mainConsole.topY + (getNumberConsoleMessages() * linePitch) + CON_BORDER_HEIGHT - linePitch
        End If
        For i = 0 To ActiveMessages_size - 1
            console_drawtext ActiveMessages(i).display, getConsoleTextColor(ActiveMessages(i).player), mainConsole.topX, _
                             TextYpos, ActiveMessages(i).JustifyType, mainConsole.width
            TextYpos = TextYpos + ActiveMessages(i).display.lineSize()
        Next
    End If
End Sub

' destroy CONPRINTF messages *
Public Sub clearInfoMessages()
    msgs_clear InfoMessages, InfoMessages_size
End Sub

' Allows toggling of the box under the console text
Public Sub setConsoleBackdropStatus(ByVal state As Boolean)
    bTextBoxActive = state
End Sub

'
'	Turns on and off display of console. It's worth
'	noting that this is just the display so if you want
'	to make sure that when it's turned back on again, there
'	are no messages, the call flushConsoleMessages first.
'
Public Sub enableConsoleDisplay(ByVal state As Boolean)
    bConsoleDisplayEnabled = state
End Sub

' Allows positioning of the console on screen
Public Sub setConsoleSizePos(ByVal x As Long, ByVal y As Long, ByVal width As Long)
    mainConsole.topX = x
    mainConsole.topY = y
    mainConsole.width = width

    ' Should be done below
    mainConsole.textDepth = 8

    ' Do not call flushConsoleMessages() when simply changing console size/position -
    ' it is possible for the console size/pos to change during display.
End Sub

'	Establishes whether the console messages stay there
Public Sub setConsolePermanence(ByVal state As Boolean, ByVal bClearOld As Boolean)
    If mainConsole.permanent = True And state = False Then
        If bClearOld Then
            flushConsoleMessages
        End If
        mainConsole.permanent = False
    Else
        If bClearOld Then
            flushConsoleMessages
        End If
        mainConsole.permanent = state
    End If
End Sub

' Check if mouse is over the Active console 'window' area
Public Function mouseOverConsoleBox() As Boolean
    Dim gotMessages As Long
    gotMessages = getNumberConsoleMessages()
    ' (the && chain in order, unsigned compares)
    If gotMessages <> 0 Then
        If ULT(mainConsole.topX, mouseX()) Then
            If ULT(mainConsole.topY, mouseY()) Then
                If ULT(mouseX(), mainConsole.topX + mainConsole.width) Then
                    If ULT(mouseY(), (mainConsole.topY + 4 + linePitch * gotMessages)) Then
                        mouseOverConsoleBox = True
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    mouseOverConsoleBox = False
End Function

' Check if mouse is over the History console 'window' area
Public Function mouseOverHistoryConsoleBox() As Boolean
    Dim nudgeright As Long, inside As Boolean
    nudgeright = 0
    If isSecondaryWindowUp() Then
        ' if a build/research/... is up, we need to move text over by this much
        nudgeright = RET_FORMWIDTH
    End If
    ' (#if 0: the hitbox of the history console window drawn)
    ' check to see if mouse is in the area when console is enabled
    If bConsoleDropped Then
        If ULT(historyConsole.topX + nudgeright, mouseX()) Then
            If ULT(historyConsole.topY, mouseY()) Then
                If ULT(mouseX(), historyConsole.topX + historyConsole.width) Then
                    inside = ULT(mouseY(), (historyConsole.topY + 4 + linePitch * NumDisplayLines))
                End If
            End If
        End If
    End If
    If inside Then
        If mousePressed(MOUSE_WUP) Then
            updatepos = updatepos - 1
        ElseIf mousePressed(MOUSE_WDN) Then
            updatepos = updatepos + 1
        End If
        If keyDown(KEY_LCTRL) Then
            showBackgroundColor = True
        Else
            showBackgroundColor = False
        End If
        mouseOverHistoryConsoleBox = True
    Else
        mouseOverHistoryConsoleBox = False
    End If
End Function

' Sets up how many lines are allowed and how many are visible
Public Sub setConsoleLineInfo(ByVal vis As Long)
    If Not (Not ULT(MAX_CONSOLE_MESSAGES, vis)) Then ASSERT_FAILURE "vis <= (64)", "src/console.cpp:646", "setConsoleLineInfo", "Request for more visible lines in the console than exist"
    consoleVisibleLines = vis
End Sub

' get how many lines are allowed and how many are visible
Public Function getConsoleLineInfo() As Long
    getConsoleLineInfo = consoleVisibleLines
End Function

' Set if new messages may be added to the console
Public Sub permitNewConsoleMessages(ByVal allow As Boolean)
    allowNewMessages = IIf(allow, 1, 0)
End Sub

' \return the visibility of the console
Public Function getConsoleDisplayStatus() As Boolean
    getConsoleDisplayStatus = bConsoleDisplayEnabled
End Function

' Use console() for when you want to display a console message,
'    and keep it in the history logs.
'
'	Use the macro CONPRINTF if you don't want it to be in the history logs.
Public Sub console(ByVal pFormat As String, ParamArray args() As Variant)
    If vbDisplayOff Then Exit Sub  ' (VB) the vbtest runs: statsref's stand-in
    Dim aBuffer As String, a As Variant ' char aBuffer[500]: Output string buffer
    a = args

    ' Print out the string
    aBuffer = vsprintf(pFormat, a)
    If Len(aBuffer) > 499 Then aBuffer = Left$(aBuffer, 499)

    ' Output it
    addConsoleMessage aBuffer, DEFAULT_JUSTIFY, SYSTEM_MESSAGE
End Sub

' console.h's CONPRINTF(format, params...): snprintf into ConsoleString, then an INFO_MESSAGE
Public Sub CONPRINTF(ByVal format As String, ParamArray params() As Variant)
    Dim a As Variant
    a = params
    ConsoleString = vsprintf(format, a)
    If Len(ConsoleString) > MAX_CONSOLE_TMP_STRING_LENGTH - 1 Then ConsoleString = Left$(ConsoleString, MAX_CONSOLE_TMP_STRING_LENGTH - 1)
    addConsoleMessage ConsoleString, DEFAULT_JUSTIFY, INFO_MESSAGE
End Sub

'--------------------------------------------------------------------------
' (VB) the lambdas

Public Function console_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_initConsoleMessages_layout
        ' []() { setConsoleSizePos(...); historyConsole... }
        If Not challengeActive And (game.typ = LEVEL_TYPE_SKIRMISH) Then
            setConsoleSizePos 16, 32, pie_GetVideoBufferWidth() - 32
        Else
            setConsoleSizePos 16, (32 + TIMER_Y), pie_GetVideoBufferWidth() - 32
        End If

        historyConsole.topX = RET_X     ' HISTORYBOX_X
        historyConsole.topY = RET_Y - 80 ' HISTORYBOX_Y
        historyConsole.width = pie_GetVideoBufferWidth() - 32
    End Select
End Function
