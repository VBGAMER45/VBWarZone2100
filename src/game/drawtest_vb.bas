Attribute VB_Name = "drawtest_vb"
' (VB) --vbdrawtest=<name>: cluster D's screenshot test. On the title screen
' each frame draws, over the backdrop, boxes and lines, a line of text in
' every font and colour, a WzText, rotated text, and the frontend image
' sheet's images; at frame DRAWTEST_SHOT_FRAME it asks screen_c for a
' screenshot (screenshots/wz2100-<time>-<name>.png in the write dir), and
' DRAWTEST_SHOT_FRAME + 5 frames later it quits. tools/draw_test.py runs it.
Option Explicit

Public vbDrawTest As String             ' the screenshot's name ("" = no test)
' (VB) --vbwatchshot=<gameTime ms>,<name>: cluster E's test. In a game, each rendered frame
' (loop_c's gameLoop) counts and every 100th logs the frame rate; the first frame at or after
' the game time asks screen_c for a screenshot named so, and the run quits 3 frames later.
Public vbWatchShotTime As Long          ' 0 = no test
Public vbWatchShotName As String
Public vbWatchRealtime As Boolean     ' a field "rt": the game at its normal speed (gameTimeSetMod 1 on the first frame; --autogame sets 500)
Public vbWatchCameraPlayer As Long    ' a field "p<N>": the camera to player N's home on the first frame (-1: none)

Private Const DRAWTEST_SHOT_FRAME As Long = 30
Private frameNo As Long
Private wzt As WzText
Private watchFrames As Long
Private watchShotFrame As Long
Private watchLastTicks As Long
' (VB) --vbscript=<file>: a test script run on the game's frames (the title screen's and the game's):
' one command a line, each run when the previous one is done: wait <frames>, move <x> <y>,
' click <x> <y>, rclick <x> <y>, key <virtual key code>, keydown <virtual key code>, keyup <virtual key
' code> (a key held across the next commands: keydown 17, key 86, keyup 17 is Ctrl+V), text <characters>,
' shot <name> (screen_c's screenshot, screenshots/wz2100-<time>-<name>.png), log <text>, quit. The mouse and keys are messages
' posted to the game's window, as Windows would send them; '#' starts a comment line.
Public vbScriptFile As String
Private Declare Function FindWindowExW Lib "user32" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function PostMessageW Lib "user32" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MapVirtualKeyW Lib "user32" (ByVal uCode As Long, ByVal uMapType As Long) As Long
Private Const MAPVK_VK_TO_VSC As Long = 0
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const MK_LBUTTON As Long = 1
Private Const MK_RBUTTON As Long = 2
Private scriptLines() As String
Private scriptCount As Long
Private scriptPos As Long
Private scriptLoaded As Boolean
Private scriptWait As Long
Private scriptWaitUntil As Long         ' (wzGetTicks) a wait's end: N frames and N / 60 s at least
Private scriptPhase As Long             ' a click's frames: 0 move, 1 down, 2 up
Private scriptQuiet As Long             ' quiet: the frames in a row without a full-screen video

Public Sub drawtest_frame()
    Dim i As Long, x As Long, y As Long, w As Long, h As Long, n As Long, rowH As Long
    Dim colours(0 To 5) As Long, img As Long
    If Len(vbDrawTest) = 0 Then Exit Sub
    frameNo = frameNo + 1

    ' boxes and lines
    pie_BoxFill 20, 20, 220, 120, pal_RGBA(40, 80, 160, 255)
    pie_UniTransBoxFill 120, 60, 320, 160, pal_RGBA(200, 40, 40, 128)
    iV_Box 20, 20, 220, 120, psPalette(WZCOL_WHITE)
    iV_ShadowBox 340, 20, 540, 120, 0, psPalette(WZCOL_FORM_LIGHT), psPalette(WZCOL_FORM_DARK), psPalette(WZCOL_FORM_BACKGROUND)
    For i = 0 To 10
        iV_Line 560, 20 + i * 10, 760, 120 - i * 10, pal_RGBA(255, 255 - i * 20, i * 20, 255)
    Next

    ' every font
    colours(0) = psPalette(WZCOL_TEXT_BRIGHT)
    colours(1) = psPalette(WZCOL_YELLOW)
    colours(2) = psPalette(WZCOL_GREEN)
    colours(3) = psPalette(WZCOL_RED)
    colours(4) = psPalette(WZCOL_WHITE)
    colours(5) = psPalette(WZCOL_LBLUE)
    y = 190
    For i = font_regular To font_medium_bold
        iV_SetTextColour colours(i Mod 6)
        iV_DrawText "Font " & i & ": The quick brown fox jumps over the lazy dog 0123456789", 20, y, i
        y = y + iV_GetTextLineSize(i) + 4
    Next
    If wzt Is Nothing Then Set wzt = WzText_new("WzText: Warzone 2100 VB", font_large)
    wzt.render 20, y + 20, psPalette(WZCOL_YELLOW)
    iV_SetTextColour psPalette(WZCOL_WHITE)
    iV_DrawTextRotated "Rotated text", 900, 200, 90!, font_regular

    ' the frontend image sheet's images, in rows
    img = 0
    If Not IsEmpty(FrontImages) Then img = FrontImages
    If img <> 0 Then
        n = IMAGEFILE_numImages(img)
        x = 20
        y = y + 70
        rowH = 0
        For i = 0 To n - 1
            w = iV_GetImageWidth(img, i)
            h = iV_GetImageHeight(img, i)
            If w > 0 And h > 0 And w <= 300 And h <= 200 Then
                If x + w > pie_GetVideoBufferWidth() - 20 Then
                    x = 20
                    y = y + rowH + 4
                    rowH = 0
                End If
                If y + h > pie_GetVideoBufferHeight() Then Exit For
                iV_DrawImage img, i, x, y
                x = x + w + 4
                If h > rowH Then rowH = h
            End If
        Next
    End If

    If frameNo = DRAWTEST_SHOT_FRAME Then
        screenDumpToDisk "screenshots/", vbDrawTest
    ElseIf frameNo = DRAWTEST_SHOT_FRAME + 5 Then
        wzQuit 0
    End If
End Sub

Public Sub watchtest_frame()
    Dim t As Long
    If vbWatchShotTime = 0 Then Exit Sub
    watchFrames = watchFrames + 1
    If watchFrames Mod 100 = 1 Then
        t = wzGetTicks()
        If watchFrames > 1 Then
            debug_ 0, LOG_INFO, "watchtest_frame", "(VB) watch: frame %d, gameTime %u, %.1f fps", watchFrames, gameTime, 100000# / IIf(t - watchLastTicks > 0, t - watchLastTicks, 1)
        Else
            debug_ 0, LOG_INFO, "watchtest_frame", "(VB) watch: first frame, gameTime %u", gameTime
            If vbWatchRealtime Then gameTimeSetMod 1, 1
            If vbWatchCameraPlayer >= 0 Then cameraToHome vbWatchCameraPlayer, False, False
        End If
        watchLastTicks = t
    End If
    If watchShotFrame = 0 Then
        If Not ULT(gameTime, vbWatchShotTime) Then
            debug_ 0, LOG_INFO, "watchtest_frame", "(VB) watch: screenshot at frame %d, gameTime %u", watchFrames, gameTime
            screenDumpToDisk "screenshots/", vbWatchShotName
            watchShotFrame = watchFrames
        End If
    ElseIf watchFrames = watchShotFrame + 3 Then
        wzQuit 0
    End If
End Sub

' This process's game window: with two copies running (the LAN test), the first window of the class may be the
' other copy's
Private Function gameWindow() As Long
    Static hWnd As Long
    Dim cls As String, h As Long, pid As Long
    If hWnd = 0 Then
        cls = "WarzoneVB"
        h = FindWindowExW(0, 0, StrPtr(cls), 0)
        Do While h <> 0
            GetWindowThreadProcessId h, pid
            If pid = GetCurrentProcessId() Then
                hWnd = h
                Exit Do
            End If
            h = FindWindowExW(0, h, StrPtr(cls), 0)
        Loop
    End If
    gameWindow = hWnd
End Function

Private Function mouseLParam(ByVal x As Long, ByVal y As Long) As Long
    mouseLParam = (x And &HFFFF&) Or SHL(y And &HFFFF&, 16)
End Function

' (VB) the script's embark command: the selected player's units (all but the transporters) board the
' first transporter on the map, as each one clicked onto it would (DORDER_EMBARK)
Private Sub scriptEmbark()
    Dim it As Long, psTransporter As Long, n As Long
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        If DROID_isTransporter(lnObj(it)) Then psTransporter = lnObj(it): Exit Do
        it = lnNext(it)
    Loop
    If psTransporter = 0 Then
        debug_ 0, LOG_INFO, "vbscript_frame", "(VB) script: embark: no transporter"
        Exit Sub
    End If
    it = apsDroidLists(selectedPlayer).head
    Do While it <> 0
        If Not DROID_isTransporter(lnObj(it)) Then
            orderDroidObj lnObj(it), DORDER_EMBARK, psTransporter, ModeImmediate
            n = n + 1
        End If
        it = lnNext(it)
    Loop
    debug_ 0, LOG_INFO, "vbscript_frame", "(VB) script: embark: %d units ordered", n
End Sub

' (VB) the script's widgets command: the widget tree of psWScreen in the log, one widget a line (its
' id, class, geometry on the screen, hidden or not, number of children)
Private Sub dumpWidgets(ByVal w As WIDGET, ByVal depth As Long)
    Dim i As Long
    debug_ 0, LOG_INFO, "dumpWidgets", "(VB) %s%d %s (%d,%d %dx%d)%s %d", Space$(depth * 2), w.id, WIDGET_kindName(w.kind), w.screenPosX(), w.screenPosY(), w.width(), w.height(), IIf(w.visible(), "", " hidden"), w.childCount()
    For i = 0 To w.childCount() - 1
        dumpWidgets w.childAt(i), depth + 1
    Next
End Sub

' one step of the --vbscript test script, each frame
Public Sub vbscript_frame()
    Dim f As Integer, content As String, cmd As String, w() As String, x As Long, y As Long, i As Long
    If Len(vbScriptFile) = 0 Then Exit Sub
    If Not scriptLoaded Then
        scriptLoaded = True
        f = FreeFile
        Open vbScriptFile For Binary Access Read As #f
        content = Space$(LOF(f))
        Get #f, , content
        Close #f
        content = Replace(content, vbCr, "")
        scriptLines = Split(content, vbLf)
        scriptCount = UBound(scriptLines) + 1
        scriptPos = 0
    End If
    If scriptWait > 0 Or wzGetTicks() - scriptWaitUntil < 0 Then
        If scriptWait > 0 Then scriptWait = scriptWait - 1
        Exit Sub
    End If
    Do While scriptPos < scriptCount
        cmd = Trim$(scriptLines(scriptPos))
        If Len(cmd) > 0 And Left$(cmd, 1) <> "#" Then Exit Do
        scriptPos = scriptPos + 1
    Loop
    If scriptPos >= scriptCount Then Exit Sub
    w = Split(cmd, " ")
    Select Case LCase$(w(0))
    Case "wait"
        scriptWait = CLng(w(1))
        scriptWaitUntil = wzGetTicks() + CLng(w(1)) * 1000 \ 60   ' (the frames of a 60 Hz screen at least)
        scriptPos = scriptPos + 1
    Case "move"
        PostMessageW gameWindow(), WM_MOUSEMOVE, 0, mouseLParam(CLng(w(1)), CLng(w(2)))
        scriptPos = scriptPos + 1
    Case "click", "rclick"
        x = CLng(w(1))
        y = CLng(w(2))
        ' a frame apart each: the move, the press, the release
        Select Case scriptPhase
        Case 0
            PostMessageW gameWindow(), WM_MOUSEMOVE, 0, mouseLParam(x, y)
            scriptPhase = 1
        Case 1
            If LCase$(w(0)) = "click" Then
                PostMessageW gameWindow(), WM_LBUTTONDOWN, MK_LBUTTON, mouseLParam(x, y)
            Else
                PostMessageW gameWindow(), WM_RBUTTONDOWN, MK_RBUTTON, mouseLParam(x, y)
            End If
            scriptPhase = 2
        Case Else
            If LCase$(w(0)) = "click" Then
                PostMessageW gameWindow(), WM_LBUTTONUP, 0, mouseLParam(x, y)
            Else
                PostMessageW gameWindow(), WM_RBUTTONUP, 0, mouseLParam(x, y)
            End If
            scriptPhase = 0
            scriptPos = scriptPos + 1
        End Select
    Case "key"
        ' WzPlatform reads the key's scan code from lParam's bits 16-23 (as SDL does): the press, then a
        ' frame later the release
        x = MapVirtualKeyW(CLng(w(1)), MAPVK_VK_TO_VSC) And &HFF&
        x = x * &H10000
        If UBound(w) >= 2 Then
            If LCase$(w(2)) = "ext" Then x = x Or &H1000000     ' (VB) "key <vk> ext": an extended key (lParam bit 24)
        End If
        If scriptPhase = 0 Then
            PostMessageW gameWindow(), WM_KEYDOWN, CLng(w(1)), 1 Or x
            scriptPhase = 1
        Else
            PostMessageW gameWindow(), WM_KEYUP, CLng(w(1)), &HC0000001 Or x
            scriptPhase = 0
            scriptPos = scriptPos + 1
        End If
    Case "keydown", "keyup"
        ' a key held across the next commands (the same messages as key's, one at a time)
        x = MapVirtualKeyW(CLng(w(1)), MAPVK_VK_TO_VSC) And &HFF&
        If LCase$(w(0)) = "keydown" Then
            PostMessageW gameWindow(), WM_KEYDOWN, CLng(w(1)), 1 Or x * &H10000
        Else
            PostMessageW gameWindow(), WM_KEYUP, CLng(w(1)), &HC0000001 Or x * &H10000
        End If
        scriptPos = scriptPos + 1
    Case "text"
        For i = Len(w(0)) + 2 To Len(cmd)
            PostMessageW gameWindow(), WM_CHAR, AscW(Mid$(cmd, i, 1)), 1
        Next
        scriptPos = scriptPos + 1
    Case "shot"
        screenDumpToDisk "screenshots/", w(1)
        scriptPos = scriptPos + 1
    Case "log"
        debug_ 0, LOG_INFO, "vbscript_frame", "(VB) script: %s", Mid$(cmd, 5)
        scriptPos = scriptPos + 1
    Case "embark"
        scriptEmbark
        scriptPos = scriptPos + 1
    Case "novideo"
        ' (VB) holds the script while a full-screen video plays (loop.cpp's video mode: the campaign's videos)
        If Not loop_GetVideoStatus() Then scriptPos = scriptPos + 1
    Case "quiet"
        ' (VB) holds the script until N frames in a row have passed without a full-screen video (a level's
        ' briefing videos start a few seconds in, one after another)
        If loop_GetVideoStatus() Then scriptQuiet = 0 Else scriptQuiet = scriptQuiet + 1
        If scriptQuiet >= CLng(w(1)) Then
            scriptQuiet = 0
            scriptPos = scriptPos + 1
        End If
    Case "waitlevel"
        ' (VB) holds the script until this level is the current one, in the game and without a video
        If StrComp(getLevelName(), w(1), vbTextCompare) = 0 And GetGameMode() = GS_NORMAL And Not loop_GetVideoStatus() Then
            scriptPos = scriptPos + 1
        End If
    Case "closeguide"
        ' (VB) closes the Game Guide when it is open (a campaign topic's pop-up pauses the game until closed)
        closeGuideScreen
        scriptPos = scriptPos + 1
    Case "widgets"
        dumpWidgets psWScreen.psForm, 0
        scriptPos = scriptPos + 1
    Case "quit"
        wzQuit 0
        scriptPos = scriptPos + 1
    Case Else
        debug_ 0, LOG_ERROR, "vbscript_frame", "(VB) script: unknown command %s", cmd
        scriptPos = scriptPos + 1
    End Select
End Sub
