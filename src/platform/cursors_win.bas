Attribute VB_Name = "cursors_win"
' lib/sdl/cursors_sdl.cpp and lib/sdl/cursor_sdl_helpers.cpp, on Win32 instead of SDL: the game's mouse
' cursors (the monochrome 32x32 images of cursors_win_data, the coloured ones from images/intfac/*.png), scaled
' by the cursor scale, made into Windows cursors as SDL's WIN_CreateCursor makes them (a 32-bit top-down DIB
' section with alpha and a mask from the alpha, CreateIconIndirect). SDL_SetCursor's place is main_win's
' win_setCursor (the window procedure keeps the cursor on WM_SETCURSOR, as SDL's does).
' The CURSOR enum (lib/framework/cursors.h) is the exe's; its values are these constants.
Option Explicit

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pbmi As BITMAPINFOHEADER, ByVal iUsage As Long, ppvBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' lib/framework/cursors.h: enum CURSOR
Private Const CURSOR_ARROW As Long = 0
Private Const CURSOR_DEST As Long = 1
Private Const CURSOR_SIGHT As Long = 2
Private Const CURSOR_TARGET As Long = 3
Private Const CURSOR_LARROW As Long = 4
Private Const CURSOR_RARROW As Long = 5
Private Const CURSOR_DARROW As Long = 6
Private Const CURSOR_UARROW As Long = 7
Private Const CURSOR_DEFAULT As Long = 8
Private Const CURSOR_EDGEOFMAP As Long = 9
Private Const CURSOR_ATTACH As Long = 10
Private Const CURSOR_ATTACK As Long = 11
Private Const CURSOR_BOMB As Long = 12
Private Const CURSOR_BRIDGE As Long = 13
Private Const CURSOR_BUILD As Long = 14
Private Const CURSOR_EMBARK As Long = 15
Private Const CURSOR_DISEMBARK As Long = 16
Private Const CURSOR_FIX As Long = 17
Private Const CURSOR_GUARD As Long = 18
Private Const CURSOR_JAM As Long = 19
Private Const CURSOR_LOCKON As Long = 20
Private Const CURSOR_SCOUT As Long = 21
Private Const CURSOR_MENU As Long = 22
Private Const CURSOR_MOVE As Long = 23
Private Const CURSOR_NOTPOSSIBLE As Long = 24
Private Const CURSOR_PICKUP As Long = 25
Private Const CURSOR_SEEKREPAIR As Long = 26
Private Const CURSOR_SELECT As Long = 27
Private Const CURSOR_MAX As Long = 28

Private currentCursor As Long           ' = CURSOR_MAX
Private lastAppliedCursor As Long       ' = CURSOR_MAX
Private aCursors(0 To CURSOR_MAX - 1) As Long    ' SDL_Cursor *: HCURSORs (0: nullptr)
Private monoCursor As Boolean
Private cursorsEnabled As Boolean
Private initialised As Boolean          ' (VB) the two CURSOR_MAX initialisers

Private Sub initState()
    If initialised Then Exit Sub
    currentCursor = CURSOR_MAX
    lastAppliedCursor = CURSOR_MAX
    initialised = True
End Sub

Private Sub scaleCursorImageForUpload(ByVal cursorImage As iV_Image, hot_x As Long, hot_y As Long)
    Dim horizontalScalePercentage As Long, verticalScalePercentage As Long
    Dim horizontalScaleFactor As Single, verticalScaleFactor As Single
    Dim scaledWidth As Long, scaledHeight As Long, contentScaleFactor As Single
    horizontalScalePercentage = platformGame().warGetCursorScale()
    verticalScalePercentage = horizontalScalePercentage

    ' Must manually scale the cursor image based on the content scale factor on Windows w/ SDL <= 3.2.x
    ' (Fixed in https://github.com/libsdl-org/SDL/pull/14285)
    contentScaleFactor = wzGetDisplayContentScale()
    horizontalScalePercentage = Fix(horizontalScalePercentage * contentScaleFactor)
    verticalScalePercentage = Fix(verticalScalePercentage * contentScaleFactor)

    If horizontalScalePercentage = 100 And verticalScalePercentage = 100 Then
        Exit Sub
    End If
    horizontalScaleFactor = horizontalScalePercentage / 100!
    verticalScaleFactor = verticalScalePercentage / 100!
    scaledWidth = Fix(cursorImage.width() * horizontalScaleFactor)
    scaledHeight = Fix(cursorImage.height() * verticalScaleFactor)
    If scaledWidth > cursorImage.width() And scaledHeight > cursorImage.height() Then
        debug_ 1295, LOG_GUI, "scaleCursorImageForUpload", "Scaling cursor image from (%u x %u) to (%u x %u)", cursorImage.width(), cursorImage.height(), scaledWidth, scaledHeight
        cursorImage.resize scaledWidth, scaledHeight
        hot_x = Fix(hot_x * horizontalScaleFactor)
        hot_y = Fix(hot_y * verticalScaleFactor)
    End If
End Sub

' SDL_Cursor* init_cursor_from_image(iV_Image&& sprite, int hot_x, int hot_y): the HCURSOR
' (SDL_CreateRGBSurfaceFrom + SDL_CreateColorCursor: SDL's WIN_CreateCursor)
Private Function init_cursor_from_image(ByVal sprite As iV_Image, ByVal hot_x As Long, ByVal hot_y As Long) As Long
    Dim w As Long, h As Long, bi As BITMAPINFOHEADER, ii As ICONINFO, pixels As Long, src() As Byte, dst() As Byte
    Dim maskbits() As Byte, maskPitch As Long, x As Long, y As Long, s As Long, d As Long, pointer As Long

    scaleCursorImageForUpload sprite, hot_x, hot_y

    w = sprite.width()
    h = sprite.height()
    If w <= 0 Or h <= 0 Or sprite.channels() <> 4 Then
        debug_ 1329, LOG_FATAL, "init_cursor_from_image", "Failed to create cursor because: %s", "(VB) not a 4-channel image"
        Exit Function
    End If
    ReDim src(0 To w * h * 4 - 1)
    CopyMemory src(0), ByVal sprite.bmp(), w * h * 4

    ' the colour: a top-down 32-bit DIB (ARGB8888: B, G, R, A bytes), the sprite's R, G, B, A
    bi.biSize = LenB(bi)
    bi.biWidth = w
    bi.biHeight = -h                    ' Invert height to make the top-down DIB.
    bi.biPlanes = 1
    bi.biBitCount = 32
    ii.hbmColor = CreateDIBSection(0, bi, 0, pixels, 0, 0)     ' DIB_RGB_COLORS
    ' the mask: 1 where the pixel is transparent (alpha 0), rows of whole 16-bit words
    maskPitch = ((w + 15) \ 16) * 2
    ReDim maskbits(0 To maskPitch * h - 1)
    ReDim dst(0 To w * h * 4 - 1)
    For y = 0 To h - 1
        For x = 0 To w - 1
            s = (y * w + x) * 4
            dst(s) = src(s + 2)
            dst(s + 1) = src(s + 1)
            dst(s + 2) = src(s)
            dst(s + 3) = src(s + 3)
            If src(s + 3) = 0 Then
                d = y * maskPitch + (x \ 8)
                maskbits(d) = maskbits(d) Or (&H80 \ (2 ^ (x Mod 8)))
            End If
        Next
    Next
    If ii.hbmColor = 0 Or pixels = 0 Then
        debug_ 1329, LOG_FATAL, "init_cursor_from_image", "Failed to create cursor because: %s", "CreateDIBSection failed (" & Err.LastDllError & ")"
        Exit Function
    End If
    CopyMemory ByVal pixels, dst(0), w * h * 4
    ii.hbmMask = CreateBitmap(w, h, 1, 1, maskbits(0))
    ii.fIcon = 0
    ii.xHotspot = hot_x
    ii.yHotspot = hot_y
    pointer = CreateIconIndirect(ii)
    DeleteObject ii.hbmColor
    DeleteObject ii.hbmMask
    If pointer = 0 Then
        debug_ 1335, LOG_FATAL, "init_cursor_from_image", "Could not create cursor because %s", "CreateIconIndirect failed (" & Err.LastDllError & ")"
        exit_ -1
    End If

    ' free up surface & image data
    sprite.clear

    init_cursor_from_image = pointer
End Function

' init_system_ColorCursor()-- Create a colored mouse cursor image
Private Function init_system_ColorCursor(ByVal cur As Long, ByVal fileName As String) As Long
    Dim sprite As New iV_Image, hot_x As Long, hot_y As Long
    If Not iV_loadImage_PNG2(fileName, sprite, True) Then
        debug_ 1354, LOG_FATAL, "init_system_ColorCursor", "Could not load cursor sprite. Exiting."
        exit_ -1
    End If

    ' We center the hotspot for all (FIXME ?)
    hot_x = sprite.width() \ 2
    hot_y = sprite.height() \ 2

    init_system_ColorCursor = init_cursor_from_image(sprite, hot_x, hot_y)
End Function

' init_system_cursor32()-- Create a monochrome mouse cursor image
Private Function init_system_cursor32(ByVal cur As Long) As Long
    Dim i As Long, row As Long, col As Long, data(0 To 4 * 32 - 1) As Byte, mask(0 To 4 * 32 - 1) As Byte
    Dim hot_x As Long, hot_y As Long, image() As String, cursorImage As iV_Image, hot() As String
    If Not (cur < CURSOR_MAX) Then ASSERT_FAILURE "cur < CURSOR_MAX", "lib/sdl/cursors_sdl.cpp:1375", "init_system_cursor32", "Attempting to load non-existent cursor: %u", cur
    image = Split(cursors_image(cur), vbLf)

    i = -1
    For row = 0 To 31
        For col = 0 To 31
            If (col Mod 8) <> 0 Then
                data(i) = (CLng(data(i)) * 2) And &HFF
                mask(i) = (CLng(mask(i)) * 2) And &HFF
            Else
                i = i + 1
                data(i) = 0
                mask(i) = 0
            End If
            Select Case Mid$(image(4 + row), col + 1, 1)
            Case "X"
                data(i) = data(i) Or 1
                mask(i) = mask(i) Or 1
            Case "."
                mask(i) = mask(i) Or 1
            Case " "
            End Select
        Next
    Next

    ' sscanf(image[4 + row], "%d,%d", &hot_x, &hot_y)
    hot = Split(image(4 + 32), ",")
    hot_x = Val(hot(0))
    hot_y = Val(hot(1))

    Set cursorImage = New iV_Image
    If Not ivImageFromMonoCursorDataSDLCompat(data, mask, 32, 32, cursorImage) Then
        debug_ 1416, LOG_FATAL, "init_system_cursor32", "Failed to convert data to cursor!"
        exit_ 3                         ' abort()
    End If

    init_system_cursor32 = init_cursor_from_image(cursorImage, hot_x, hot_y)
End Function

' cursor_sdl_helpers.cpp: the monochrome data and mask as SDL_CreateCursor reads them (a set mask bit: black or
' white; a clear one: black where the data bit is set, else transparent), into an RGBA image
Private Function ivImageFromMonoCursorDataSDLCompat(data() As Byte, mask() As Byte, ByVal w As Long, ByVal h As Long, ByVal image As iV_Image) As Boolean
    Dim x As Long, y As Long, p As Long, datab As Long, maskb As Long, px() As Byte, k As Long, isBlack As Boolean, isWhite As Boolean
    If Not image.allocate(w, h, 4, True) Then
        Exit Function
    End If
    ReDim px(0 To w * h * 4 - 1)
    For y = 0 To h - 1
        p = y * w * 4
        For x = 0 To w - 1
            If (x Mod 8) = 0 Then
                datab = data(k)
                maskb = mask(k)
                k = k + 1
            End If
            isBlack = False
            isWhite = False
            If (maskb And &H80) <> 0 Then
                If (datab And &H80) <> 0 Then isBlack = True Else isWhite = True
            Else
                If (datab And &H80) <> 0 Then isBlack = True     ' else transparent
            End If
            If isBlack Then
                px(p + 3) = &HFF                                ' 0xFF000000: R, G, B 0, A 255
            ElseIf isWhite Then
                px(p) = &HFF
                px(p + 1) = &HFF
                px(p + 2) = &HFF
                px(p + 3) = &HFF
            End If
            p = p + 4
            datab = (datab * 2) And &HFF
            maskb = (maskb * 2) And &HFF
        Next
    Next
    CopyMemory ByVal image.bmp_w(), px(0), w * h * 4
    ivImageFromMonoCursorDataSDLCompat = True
End Function

' wzSetCursor()-- Set the current cursor. Cursor is actually applied in wzApplyCursor() in mainLoop()
Public Sub wzSetCursor(ByVal cur As Long)
    initState
    If Not (cur < CURSOR_MAX) Then ASSERT_FAILURE "cur < CURSOR_MAX", "lib/sdl/cursors_sdl.cpp:1428", "wzSetCursor", "Specified cursor(%d) is over our limit of (%d)!", cur, CURSOR_MAX

    currentCursor = cur
End Sub

Public Sub wzSDLReinitCursors()
    initState
    sdlFreeCursors
    If platformGame().warGetColouredCursor() Then sdlInitColoredCursors Else sdlInitCursors
    If currentCursor < CURSOR_MAX Then win_setCursor aCursors(currentCursor) Else win_setCursor 0
End Sub

Public Sub wzQueueRefreshCursor()
    initState
    lastAppliedCursor = CURSOR_MAX
End Sub

Public Sub wzApplyCursor()
    initState
    If Not cursorsEnabled Then Exit Sub

    ' If mouse cursor options change, change cursors (used to only work on mouse options screen for some reason)
    If Not (platformGame().warGetColouredCursor() Xor monoCursor) Then
        wzSDLReinitCursors
        lastAppliedCursor = currentCursor
        Exit Sub
    End If

    ' If we are already using this cursor then return
    If currentCursor = lastAppliedCursor Then
        Exit Sub
    End If

    If currentCursor < CURSOR_MAX Then win_setCursor aCursors(currentCursor) Else win_setCursor 0
    lastAppliedCursor = currentCursor
End Sub

' sdlInitColoredCursors() --Setup the mouse cursor with bitmaps
Public Sub sdlInitColoredCursors()
    initState
    monoCursor = False
    aCursors(CURSOR_ARROW) = init_system_cursor32(CURSOR_ARROW)
    aCursors(CURSOR_DEST) = init_system_ColorCursor(CURSOR_DEST, "images/intfac/image_cursor_dest.png")
    aCursors(CURSOR_SIGHT) = init_system_cursor32(CURSOR_SIGHT)
    aCursors(CURSOR_TARGET) = init_system_cursor32(CURSOR_TARGET)
    aCursors(CURSOR_LARROW) = init_system_ColorCursor(CURSOR_LARROW, "images/intfac/image_cursor_larrow.png")
    aCursors(CURSOR_RARROW) = init_system_ColorCursor(CURSOR_RARROW, "images/intfac/image_cursor_rarrow.png")
    aCursors(CURSOR_DARROW) = init_system_ColorCursor(CURSOR_DARROW, "images/intfac/image_cursor_darrow.png")
    aCursors(CURSOR_UARROW) = init_system_ColorCursor(CURSOR_UARROW, "images/intfac/image_cursor_uarrow.png")
    aCursors(CURSOR_DEFAULT) = init_system_ColorCursor(CURSOR_DEFAULT, "images/intfac/image_cursor_default.png")
    aCursors(CURSOR_EDGEOFMAP) = init_system_cursor32(CURSOR_EDGEOFMAP)
    aCursors(CURSOR_ATTACH) = init_system_ColorCursor(CURSOR_ATTACH, "images/intfac/image_cursor_attach.png")
    aCursors(CURSOR_ATTACK) = init_system_ColorCursor(CURSOR_ATTACK, "images/intfac/image_cursor_attack.png")
    aCursors(CURSOR_BOMB) = init_system_ColorCursor(CURSOR_BOMB, "images/intfac/image_cursor_bomb.png")
    aCursors(CURSOR_BRIDGE) = init_system_ColorCursor(CURSOR_BRIDGE, "images/intfac/image_cursor_bridge.png")
    aCursors(CURSOR_BUILD) = init_system_ColorCursor(CURSOR_BUILD, "images/intfac/image_cursor_build.png")
    aCursors(CURSOR_EMBARK) = init_system_ColorCursor(CURSOR_EMBARK, "images/intfac/image_cursor_embark.png")
    aCursors(CURSOR_DISEMBARK) = init_system_ColorCursor(CURSOR_DISEMBARK, "images/intfac/image_cursor_disembark.png")
    aCursors(CURSOR_FIX) = init_system_ColorCursor(CURSOR_FIX, "images/intfac/image_cursor_fix.png")
    aCursors(CURSOR_GUARD) = init_system_ColorCursor(CURSOR_GUARD, "images/intfac/image_cursor_guard.png")
    aCursors(CURSOR_JAM) = init_system_ColorCursor(CURSOR_JAM, "images/intfac/image_cursor_ecm.png")
    aCursors(CURSOR_LOCKON) = init_system_ColorCursor(CURSOR_LOCKON, "images/intfac/image_cursor_lockon.png")
    aCursors(CURSOR_SCOUT) = init_system_ColorCursor(CURSOR_SCOUT, "images/intfac/image_cursor_scout.png")
    aCursors(CURSOR_MENU) = init_system_cursor32(CURSOR_MENU)
    aCursors(CURSOR_MOVE) = init_system_ColorCursor(CURSOR_MOVE, "images/intfac/image_cursor_move.png")
    aCursors(CURSOR_NOTPOSSIBLE) = init_system_ColorCursor(CURSOR_NOTPOSSIBLE, "images/intfac/image_cursor_notpos.png")
    aCursors(CURSOR_PICKUP) = init_system_ColorCursor(CURSOR_PICKUP, "images/intfac/image_cursor_pickup.png")
    aCursors(CURSOR_SEEKREPAIR) = init_system_ColorCursor(CURSOR_SEEKREPAIR, "images/intfac/image_cursor_repair.png")
    aCursors(CURSOR_SELECT) = init_system_ColorCursor(CURSOR_SELECT, "images/intfac/image_cursor_select.png")
    cursorsEnabled = True
End Sub

' sdlInitCursors() --Setup the mouse cursor with monochrome data
Public Sub sdlInitCursors()
    Dim i As Long
    initState
    monoCursor = True
    For i = CURSOR_ARROW To CURSOR_SELECT
        aCursors(i) = init_system_cursor32(i)
    Next
    cursorsEnabled = True
End Sub

Public Sub sdlFreeCursors()
    Dim i As Long
    initState
    win_setCursor 0                     ' SDL_SetCursor(SDL_GetDefaultCursor())
    For i = 0 To CURSOR_MAX - 1
        If aCursors(i) <> 0 Then
            DestroyCursor aCursors(i)
            aCursors(i) = 0
        End If
    Next
    cursorsEnabled = False
End Sub
