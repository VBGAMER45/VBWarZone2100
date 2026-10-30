Attribute VB_Name = "tip_c"
' lib/widget/tip.cpp, tip.h: the tool tip display system. tipSourceWidget (a
' WIDGET *, only compared) is the widget's address.
Option Explicit

' Time (in milliseconds) after pointing widget required to show the tool tip
Private Const TIP_PAUSE As Long = 200

' Time (in milliseconds) required for a tip to be refreshed since the last refresh
Private Const TIP_REFRESH_COOLDOWN As Long = 500

' Size of border around tip text
Private Const TIP_HGAP As Long = 6
Private Const TIP_VGAP As Long = 3

Private Enum TIP_STATE
    TIP_INACTIVE
    TIP_ACTIVE
    TIP_CLICKED
End Enum
Private tipState As TIP_STATE

' struct TipDisplayCache {std::vector<WzText> wzTip; std::vector<int> wzTipYStart;} displayCache
Private displayCache_wzTip() As WzText
Private displayCache_wzTipYStart() As Long
Private displayCache_count As Long

Private startTime As Long               ' int32_t
Private lastMouseCoords As Vector2i
Private tipRect As WzRect
Private textOffset As Vector2i
Private tipText As String               ' std::string
Private tipSourceWidget As Long         ' WIDGET *: its address
Private TipColour As Long               ' PIELIGHT
Private has_lastRefreshTime As Boolean  ' nonstd::optional<uint32_t> lastRefreshTime
Private lastRefreshTime As Long

' Initialise the tool tip module
Public Sub tipInitialise()
    tipState = TIP_INACTIVE
    TipColour = psPalette(WZCOL_WHITE)
End Sub

Public Sub tipShutdown()
    displayCacheClear
End Sub

' Set the global toop tip text colour.
Public Sub widgSetTipColour(ByVal colour As Long)
    TipColour = colour
End Sub

' (VB) displayCache.wzTip.clear() (and wzTipYStart's)
Private Sub displayCacheClear()
    Erase displayCache_wzTip
    Erase displayCache_wzTipYStart
    displayCache_count = 0
End Sub

' (VB) displayCache.wzTip.resize(n), displayCache.wzTipYStart.resize(n): the WzTexts kept, new ones made
Private Sub displayCacheResize(ByVal n As Long)
    Dim i As Long
    If n = 0 Then
        displayCacheClear
        Exit Sub
    End If
    If displayCache_count = 0 Then
        ReDim displayCache_wzTip(0 To n - 1)
        ReDim displayCache_wzTipYStart(0 To n - 1)
    Else
        ReDim Preserve displayCache_wzTip(0 To n - 1)
        ReDim Preserve displayCache_wzTipYStart(0 To n - 1)
    End If
    For i = displayCache_count To n - 1
        Set displayCache_wzTip(i) = New WzText
        displayCache_wzTipYStart(i) = 0
    Next
    displayCache_count = n
End Sub

' static std::vector<std::string> splitString(const std::string& str, char delimiter): strings(0 .. n - 1), n
' returned (std::getline: no last empty piece after a final delimiter)
Private Function splitString(ByVal str_ As String, ByVal delimiter As String, strings() As String) As Long
    Dim n As Long, p As Long, q As Long
    Erase strings
    p = 1
    Do While p <= Len(str_)
        q = InStr(p, str_, delimiter)
        If q = 0 Then q = Len(str_) + 1
        If n = 0 Then
            ReDim strings(0 To 3)
        ElseIf n > UBound(strings) Then
            ReDim Preserve strings(0 To 2 * n - 1)
        End If
        strings(n) = Mid$(str_, p, q - p)
        n = n + 1
        p = q + 1
    Loop
    splitString = n
End Function

Private Sub tipStop()
    tipState = TIP_INACTIVE
    displayCacheClear
    tipText = ""
    startTime = wzGetTicks()
    tipSourceWidget = 0
End Sub

Private Function isHighlightedChanged(ByVal mouseOverWidget As WIDGET) As Boolean
    If mouseOverWidget Is Nothing Then
        isHighlightedChanged = True
        Exit Function
    End If
    isHighlightedChanged = tipSourceWidget <> ObjPtr(mouseOverWidget)
End Function

Private Sub refreshTip(ByVal mouseOverWidget As WIDGET)
    Dim newTip As String, fontId As iV_fonts, lockedScreen As W_SCREEN, maxLineWidth As Long, lines() As String
    Dim nLines As Long, totalLineSize As Long, n As Long, width_ As Long, height_ As Long, x As Long, y As Long, yOffset As Long
    If isHighlightedChanged(mouseOverWidget) Then
        tipStop
        Exit Sub
    End If

    If has_lastRefreshTime Then
        If wzGetTicks() - lastRefreshTime < TIP_REFRESH_COOLDOWN Then
            Exit Sub
        End If
    End If

    newTip = mouseOverWidget.getTip()
    lastRefreshTime = wzGetTicks()
    has_lastRefreshTime = True
    If newTip = tipText Then
        Exit Sub
    End If

    tipText = newTip
    If tipText = "" Then
        Exit Sub
    End If

    fontId = font_regular
    Set lockedScreen = mouseOverWidget.screenPointer
    If Not lockedScreen Is Nothing Then
        fontId = lockedScreen.TipFontID
    End If

    maxLineWidth = 0
    nLines = splitString(tipText, vbLf, lines)
    totalLineSize = 0
    displayCacheResize nLines
    For n = 0 To nLines - 1
        displayCache_wzTip(n).setText lines(n), fontId
        maxLineWidth = MAX_(maxLineWidth, displayCache_wzTip(n).width())
        displayCache_wzTipYStart(n) = totalLineSize
        totalLineSize = totalLineSize + displayCache_wzTip(n).lineSize()
    Next

    ' Position the tip box
    width_ = maxLineWidth + TIP_HGAP * 2
    height_ = TIP_VGAP * 2 + totalLineSize
    If nLines <> 0 Then
        height_ = height_ + displayCache_wzTip(nLines - 1).belowBase()
    End If
    x = clip_(mouseOverWidget.screenPosX() + mouseOverWidget.width() \ 2, 0, screenWidth - width_ - 1)
    y = MIN_(mouseOverWidget.screenPosY() - height_ - TIP_VGAP, screenHeight)
    If y < 0 Then
        ' Position the tip below the button
        y = MAX_(mouseOverWidget.screenPosY() + mouseOverWidget.height() + TIP_VGAP, 0)
    End If
    tipRect = WzRect_(x - 1, y - 1, width_ + 2, height_ + 2)

    ' Position the text
    yOffset = (height_ - totalLineSize) \ 2
    If nLines <> 0 Then
        yOffset = yOffset - displayCache_wzTip(0).aboveBase()
    End If
    textOffset = Vector2i_(x + TIP_HGAP, y + yOffset)
End Sub

Private Sub handleTipInactive()
    Dim mouseCoords As Vector2i, mouseOverWidget As WIDGET, changed As Boolean
    mouseCoords = Vector2i_(mouseX(), mouseY())
    changed = (mouseCoords.x <> lastMouseCoords.x Or mouseCoords.y <> lastMouseCoords.y)
    If Not changed Then
        Set mouseOverWidget = getMouseOverWidget()
        changed = isHighlightedChanged(mouseOverWidget)
    End If
    If changed Then
        lastMouseCoords = mouseCoords
        startTime = wzGetTicks()
        tipSourceWidget = ObjPtr(mouseOverWidget)
        Exit Sub
    End If

    If wzGetTicks() - startTime > TIP_PAUSE Then
        tipState = TIP_ACTIVE
        has_lastRefreshTime = False
        refreshTip mouseOverWidget
    End If
End Sub

Private Sub handleTipActive()
    Dim n As Long
    If mousePressed(MOUSE_LMB) Then
        tipState = TIP_CLICKED
        Exit Sub
    End If

    refreshTip getMouseOverWidget()

    If Len(tipText) = 0 Then
        Exit Sub
    End If

    ' Draw the tool tip
    iV_ShadowBox tipRect.x, tipRect.y, WzRect_right(tipRect), WzRect_bottom(tipRect), 1, psPalette(WZCOL_FORM_LIGHT), psPalette(WZCOL_FORM_DARK), psPalette(WZCOL_FORM_TIP_BACKGROUND)

    For n = 0 To displayCache_count - 1
        displayCache_wzTip(n).render textOffset.x, textOffset.y + displayCache_wzTipYStart(n), TipColour
    Next
End Sub

Private Sub handleTipClicked()
    If isHighlightedChanged(getMouseOverWidget()) Then
        tipStop
    End If
End Sub

' Update and possibly display the tip
Public Sub tipDisplay()
    Select Case tipState
    Case TIP_INACTIVE: handleTipInactive
    Case TIP_ACTIVE: handleTipActive
    Case TIP_CLICKED: handleTipClicked
    End Select
End Sub
