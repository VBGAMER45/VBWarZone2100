Attribute VB_Name = "screen_c"
' lib/ivis_opengl/screen.cpp: basic double buffered display (the backdrop,
' the map preview, screenshots, the GPU performance counters, the debug
' scenes). The screenshot is taken through gfx_api_context_getScreenshot,
' whose callback runs here after it returns; the save "thread" runs at once
' (no threads: CONVENTIONS), and so do the calls it dispatches to the main
' thread. screenWidth and screenHeight are lib/sdl's (WzPlatform): reached
' through Property procedures (stand-ins until the callback to the platform
' exists). A const char * that may be nullptr is vbNullString.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' console.h
Private Const LEFT_JUSTIFY As Long = 0
Private Const SYSTEM_MESSAGE As Long = -1
Private Const INFO_MESSAGE As Long = -3

' global used to indicate preferred internal OpenGL format
Private wz_texture_compression As Boolean   ' = true

Private bBackDrop As Boolean            ' = false
Private screendump_filename As String
Private screendump_required As Boolean  ' = false

Private backdropGfx As GFX              ' = nullptr
Private backdropIsMapPreview As Boolean ' = false

Private perfStarted As Boolean          ' = false
Private Type PERF_STORE
    counters(0 To PERF_COUNT - 1) As Double ' uint64_t
End Type
Private perfList() As PERF_STORE
Private perfList_size As Long
Private queryActive As PERF_POINT       ' = PERF_COUNT

Private preview_width As Long           ' = 0
Private preview_height As Long          ' = 0
Private player_pos(0 To MAX_PLAYERS - 1) As Vector2i
Private player_Text(0 To MAX_PLAYERS - 1) As WzText
Private mappreview As Boolean           ' = false

Private sceneActive As String           ' static const char *sceneActive = nullptr;

Private hadBackDrop As Boolean          ' screen_FlipIfBackDropTransition's static auto hadBackDrop = false;

Private screenReady As Boolean          ' (VB) whether InitScreen ran

' (VB) the globals' C++ initialisers (a VB module has none): bool wz_texture_compression = true;
' static PERF_POINT queryActive = PERF_COUNT; static WzText player_Text[MAX_PLAYERS];
Private Sub InitScreen()
    Dim i As Long
    If screenReady Then Exit Sub
    screenReady = True
    wz_texture_compression = True
    queryActive = PERF_COUNT
    sceneActive = vbNullString
    For i = 0 To MAX_PLAYERS - 1
        Set player_Text(i) = New WzText
    Next
End Sub

Public Function get_wz_texture_compression() As Boolean
    InitScreen
    get_wz_texture_compression = wz_texture_compression
End Function

Public Sub set_wz_texture_compression(ByVal value As Boolean)
    InitScreen
    wz_texture_compression = value
End Sub

' Initialise the double buffered display
Public Function screenInitialise() As Boolean
    InitScreen
    If screenWidth < 640 Then screenWidth = 640       ' MAX(screenWidth, 640)
    If screenHeight < 480 Then screenHeight = 480     ' MAX(screenHeight, 480)

    pie_Skybox_Init

    ' Generate backdrop render
    Set backdropGfx = GFX_new(GFX_TEXTURE, 2)

    screenInitialise = True
End Function

Public Function wzPerfAvailable() As Boolean
    wzPerfAvailable = gfx_api_context_debugPerfAvailable()
End Function

Public Sub wzPerfStart()
    InitScreen
    perfStarted = gfx_api_context_debugPerfStart(perfList_size)
End Sub

Private Sub wzPerfWriteOut(list() As PERF_STORE, ByVal list_size As Long, ByVal outfile As String)
    Dim fileHandle As Long, fileHeader() As Byte, line_ As String, bytes() As Byte, i As Long, j As Long
    fileHandle = PHYSFS_openWrite(outfile)
    If fileHandle <> 0 Then
        ' const char fileHeader[] = "...\n"; (sizeof includes the terminating NUL)
        fileHeader = StrConv("START, EFF, TERRAIN, SKY, LOAD, PRTCL, WATER, MODELS, MISC, GUI" & vbLf & vbNullChar, vbFromUnicode)
        If WZ_PHYSFS_writeBytes(fileHandle, fileHeader, 0, UBound(fileHeader) + 1) <> UBound(fileHeader) + 1 Then
            ' Failed to write header to file
            debug_ 119, LOG_ERROR, "wzPerfWriteOut", "wzPerfWriteOut: could not write header to %s; PHYSFS error: %s", outfile, WZ_PHYSFS_getLastError()
            PHYSFS_close fileHandle
            Exit Sub
        End If
        For i = 0 To list_size - 1
            line_ = ""
            line_ = line_ & va("%llu", list(i).counters(PERF_START_FRAME))
            For j = 1 To PERF_COUNT - 1
                line_ = line_ & ", " & va("%llu", list(i).counters(j))
            Next
            line_ = line_ & vbLf
            If Not (Len(line_) <= 4294967295#) Then ASSERT_FAILURE "line.toUtf8().length() <= static_cast<size_t>(std::numeric_limits<PHYSFS_uint32>::max())", "lib/ivis_opengl/screen.cpp:132", "wzPerfWriteOut", "Line length exceeds PHYSFS_uint32::max"
            bytes = StrConv(line_, vbFromUnicode)
            If WZ_PHYSFS_writeBytes(fileHandle, bytes, 0, UBound(bytes) + 1) <> UBound(bytes) + 1 Then
                ' Failed to write line to file
                debug_ 136, LOG_ERROR, "wzPerfWriteOut", "wzPerfWriteOut: could not line to %s; PHYSFS error: %s", outfile, WZ_PHYSFS_getLastError()
                PHYSFS_close fileHandle
                Exit Sub
            End If
        Next
        PHYSFS_close fileHandle
    Else
        debug_ 145, LOG_ERROR, "wzPerfWriteOut", "%s could not be opened: %s", outfile, WZ_PHYSFS_getLastError()
        ' assert(!"wzPerfWriteOut: couldn't open file for writing"); (a release build has none)
    End If
End Sub

Public Sub wzPerfShutdown()
    InitScreen
    If perfList_size = 0 Then
        Exit Sub
    End If

    ' write performance counter list to file
    wzPerfWriteOut perfList, perfList_size, "gfx-performance.csv"

    ' all done, clear data
    gfx_api_context_debugPerfStop: perfStarted = False
    perfList_size = 0
    queryActive = PERF_COUNT
End Sub

' call after swap buffers
Public Sub wzPerfFrame()
    Dim store As PERF_STORE, i As Long, aclock As Double, t As tm
    InitScreen
    If Not perfStarted Then
        Exit Sub ' not started yet
    End If
    If Not (queryActive = PERF_COUNT) Then ASSERT_FAILURE "queryActive == PERF_COUNT", "lib/ivis_opengl/screen.cpp:173", "wzPerfFrame", "Missing wfPerfEnd() call"
    For i = 0 To PERF_COUNT - 1
        store.counters(i) = gfx_api_context_debugGetPerfValue(i)
    Next
    If perfList_size = 0 Then
        ReDim perfList(0 To 15)
    ElseIf perfList_size > UBound(perfList) Then
        ReDim Preserve perfList(0 To 2 * perfList_size - 1)
    End If
    perfList(perfList_size) = store
    perfList_size = perfList_size + 1
    gfx_api_context_debugPerfStop: perfStarted = False

    ' Make a screenshot to document sample content

    aclock = time_()        ' Get time in seconds
    t = getLocalTime(aclock)  ' Convert time to struct

    screendump_filename = va("screenshots/wz2100-perf-sample-%02zu-%04d%02d%02d_%02d%02d%02d.png", perfList_size - 1, _
             t.tm_year + 1900, t.tm_mon + 1, t.tm_mday, t.tm_hour, t.tm_min, t.tm_sec)
    screendump_required = True
    gfx_api_context_debugStringMarker "Performance sample complete"
End Sub

' descr: vbNullString = nullptr
Public Sub wzSceneBegin(ByVal descr As String)
    InitScreen
    If Not (StrPtr(sceneActive) = 0) Then ASSERT_FAILURE "sceneActive == nullptr", "lib/ivis_opengl/screen.cpp:198", "wzSceneBegin", "Out of order scenes: Wanted to start %s, was in %s", descr, sceneActive
    gfx_api_context_debugSceneBegin descr
    sceneActive = descr
End Sub

Public Sub wzSceneEnd(ByVal descr As String)
    InitScreen
    If Not (StrPtr(descr) = 0 Or StrComp(descr, sceneActive, vbBinaryCompare) = 0) Then ASSERT_FAILURE "descr == nullptr || strcmp(descr, sceneActive) == 0", "lib/ivis_opengl/screen.cpp:205", "wzSceneEnd", "Out of order scenes: Wanted to stop %s, was in %s", descr, sceneActive
    gfx_api_context_debugSceneEnd descr
    sceneActive = vbNullString
End Sub

Public Sub wzPerfBegin(ByVal pp As PERF_POINT, ByVal descr As String)
    InitScreen
    If Not (queryActive = PERF_COUNT Or pp > queryActive) Then ASSERT_FAILURE "queryActive == PERF_COUNT || pp > queryActive", "lib/ivis_opengl/screen.cpp:212", "wzPerfBegin", "Out of order timer query call"
    queryActive = pp
    gfx_api_context_debugPerfBegin pp, descr
End Sub

Public Sub wzPerfEnd(ByVal pp As PERF_POINT)
    InitScreen
    If Not (queryActive = pp) Then ASSERT_FAILURE "queryActive == pp", "lib/ivis_opengl/screen.cpp:219", "wzPerfEnd", "Mismatched wzPerfBegin...End"
    queryActive = PERF_COUNT
    gfx_api_context_debugPerfEnd pp
End Sub

Public Sub screenShutDown()
    Dim i As Long
    InitScreen
    pie_ShutDown
    pie_TexShutDown
    iV_TextShutdown

    pie_Skybox_Shutdown

    Set backdropGfx = Nothing

    Set pie_internal_rectBuffer = Nothing

    For i = 0 To MAX_PLAYERS - 1
        Set player_Text(i) = New WzText
    Next

    iV_ImageFileShutdown
End Sub

'/ Display a random backdrop from files in dirname starting with basename.
'/ dirname must have a trailing slash.
Public Sub screen_SetRandomBackdrop(ByVal dirname As String, ByVal basename As String)
    Dim names() As String, names_size As Long, files() As String, files_count As Long, i As Long
    Dim length As Long, searchPath As String, ran As Long, full_path As String
    If Not (StrPtr(dirname) <> 0) Then ASSERT_FAILURE "dirname != nullptr", "lib/ivis_opengl/screen.cpp:250", "screen_SetRandomBackdrop", "dirname is null": Exit Sub
    ' vector to hold the strings we want
    ' all the files in dirname
    ' Walk thru the files in our dir, adding the ones that start with basename to our vector of strings
    length = Len(basename)
    If WZ_PHYSFS_enumerateFiles(dirname, files, files_count) Then
        If files_count > 0 Then ReDim names(0 To files_count - 1)
        For i = 0 To files_count - 1
            ' does our filename start with basename?
            If StrComp(Left$(files(i), length), basename, vbBinaryCompare) = 0 Then
                names(names_size) = files(i)
                names_size = names_size + 1
            End If
        Next
    End If

    If names_size = 0 Then
        searchPath = dirname
        If Len(searchPath) > 0 And Right$(searchPath, 1) <> "/" Then
            searchPath = searchPath & "/"
        End If
        debug_ 271, LOG_FATAL, "screen_SetRandomBackdrop", "Missing files: ""%s%s*"" - data files / folders may be corrupt.", searchPath, basename
        exit_ 3                         ' abort()
        Exit Sub
    End If

    ' pick a random name from our vector of names
    ran = rand() Mod names_size
    full_path = dirname & names(ran)

    screen_SetBackDropFromFile full_path
End Sub

Public Sub screen_SetBackDropFromFile(ByVal filename As String)
    Dim maxTextureSize As Long
    maxTextureSize = gfx_api_context_get_context_value(context_value_MAX_TEXTURE_SIZE)
    backdropGfx.loadTexture filename, texture_type_user_interface, maxTextureSize, maxTextureSize
    backdropIsMapPreview = False
    ' Generate coordinates and put them into VBOs
    screen_GenerateCoordinatesAndVBOs
End Sub

Public Sub screen_StopBackDrop()
    bBackDrop = False  'checking [movie]
End Sub

Public Function screen_RestartBackDrop() As Boolean
    Dim changedValue As Boolean
    changedValue = Not bBackDrop
    bBackDrop = True
    screen_RestartBackDrop = changedValue
End Function

Public Function screen_GetBackDrop() As Boolean
    screen_GetBackDrop = bBackDrop
End Function

Public Sub screen_Display()
    Dim modelViewProjectionMatrix As mat4, cb As constant_buffer_type_SHADER_GFX_TEXT
    Dim s1 As Long, s2 As Long, scale_ As Long, w As Long, h As Long, text As String
    Dim i As Long, x As Long, y As Long
    InitScreen
    ' Draw backdrop
    modelViewProjectionMatrix = glm_ortho(0!, CSng(pie_GetVideoBufferWidth()), CSng(pie_GetVideoBufferHeight()), 0!)
    gfx_api_pso_get(BackDropPSO).bind
    cb.transform_matrix = modelViewProjectionMatrix
    cb.offset = glm_vec2(0!, 0!)
    cb.size = glm_vec2(0!, 0!)
    cb.color = glm_vec4(1!, 1!, 1!, 1!)
    cb.texture = 0
    gfx_api_pso_get(BackDropPSO).bind_constants VarPtr(cb), LenB(cb)
    backdropGfx.draw BackDropPSO, modelViewProjectionMatrix

    If mappreview Then
        s1 = screenWidth \ preview_width
        s2 = screenHeight \ preview_height
        If s1 < s2 Then scale_ = s1 Else scale_ = s2   ' MIN(s1, s2)
        w = preview_width * scale_
        h = preview_height * scale_

        For i = 0 To MAX_PLAYERS - 1
            x = player_pos(i).x
            y = player_pos(i).y

            If x = &H77777777 Then
                GoTo continue_
            End If

            x = screenWidth \ 2 - w \ 2 + x * scale_
            y = screenHeight \ 2 - h \ 2 + y * scale_
            text = CStr(i)
            player_Text(i).setText text, font_large
            player_Text(i).render x - 1, y - 1, psPalette(0)    ' WZCOL_BLACK
            player_Text(i).render x + 1, y - 1, psPalette(0)    ' WZCOL_BLACK
            player_Text(i).render x - 1, y + 1, psPalette(0)    ' WZCOL_BLACK
            player_Text(i).render x + 1, y + 1, psPalette(0)    ' WZCOL_BLACK
            player_Text(i).render x, y, psPalette(1)            ' WZCOL_WHITE
continue_:
        Next
    End If
End Sub

'******************************************************************
'slight hack to display maps (or whatever) in background.
'bitmap MUST be (BACKDROP_HACK_WIDTH * BACKDROP_HACK_HEIGHT) for now.
Private Sub screen_GenerateCoordinatesAndVBOs()
    Dim x1 As Single, x2 As Single, y1 As Single, y2 As Single, tx As Single, ty As Single
    Dim scale_ As Long, w As Long, h As Long, aspect As Single, backdropAspect As Single, offset As Long
    Dim s1 As Long, s2 As Long, texcoords(0 To 7) As Single, vertices(0 To 7) As Single
    ' assert(backdropGfx != nullptr); (a release build has none)

    x1 = 0: x2 = screenWidth: y1 = 0: y2 = screenHeight
    tx = 1: ty = 1
    scale_ = 0: w = 0: h = 0
    aspect = screenWidth / CSng(screenHeight)
    backdropAspect = 4 / CSng(3)

    If aspect < backdropAspect Then
        offset = ToInt((screenWidth - screenHeight * backdropAspect) / 2)
        x1 = x1 + offset
        x2 = x2 - offset
    Else
        offset = ToInt((screenHeight - screenWidth / backdropAspect) / 2)
        y1 = y1 + offset
        y2 = y2 - offset
    End If

    If backdropIsMapPreview Then ' preview
        s1 = screenWidth \ preview_width
        s2 = screenHeight \ preview_height
        If s1 < s2 Then scale_ = s1 Else scale_ = s2   ' MIN(s1, s2)

        w = preview_width * scale_
        h = preview_height * scale_
        x1 = screenWidth \ 2 - w \ 2
        x2 = screenWidth \ 2 + w \ 2
        y1 = screenHeight \ 2 - h \ 2
        y2 = screenHeight \ 2 + h \ 2

        tx = preview_width / CSng(BACKDROP_HACK_WIDTH)
        ty = preview_height / CSng(BACKDROP_HACK_HEIGHT)
    End If

    ' Generate coordinates and put them into VBOs
    texcoords(0) = 0!: texcoords(1) = 0!: texcoords(2) = tx: texcoords(3) = 0!
    texcoords(4) = 0!: texcoords(5) = ty: texcoords(6) = tx: texcoords(7) = ty
    vertices(0) = x1: vertices(1) = y1: vertices(2) = x2: vertices(3) = y1
    vertices(4) = x1: vertices(5) = y2: vertices(6) = x2: vertices(7) = y2
    backdropGfx.buffers 4, VarPtr(vertices(0)), VarPtr(texcoords(0))
End Sub

' newBackdropImage: iV_Image&& (the caller drops it)
Public Sub screen_Upload(ByVal newBackdropImage As iV_Image)
    ' Slight hack to display maps previews in background.
    ' Bitmap MUST be (BACKDROP_HACK_WIDTH * BACKDROP_HACK_HEIGHT) for now.
    backdropGfx.loadTexture_image newBackdropImage, texture_type_user_interface, "mem::generated_map_preview"
    backdropIsMapPreview = True

    ' Generate coordinates and put them into VBOs
    screen_GenerateCoordinatesAndVBOs
End Sub

Public Sub screen_updateGeometry()
    If Not backdropGfx Is Nothing Then
        screen_GenerateCoordinatesAndVBOs
    End If
End Sub

' playerpositions: playerpositions(0 .. MAX_PLAYERS - 1)
Public Sub screen_enableMapPreview(ByVal width As Long, ByVal height As Long, playerpositions() As Vector2i)
    Dim i As Long
    mappreview = True
    preview_width = width
    preview_height = height
    For i = 0 To MAX_PLAYERS - 1
        player_pos(i).x = playerpositions(i).x
        player_pos(i).y = playerpositions(i).y
    Next
End Sub

' (VB) screen_enableMapPreview for the exe: pPlayerpositions is the address of MAX_PLAYERS Vector2i
Public Sub screen_enableMapPreview_p(ByVal width As Long, ByVal height As Long, ByVal pPlayerpositions As Long)
    Dim positions(0 To MAX_PLAYERS - 1) As Vector2i
    CopyMemory positions(0), ByVal pPlayerpositions, MAX_PLAYERS * 8
    screen_enableMapPreview width, height, positions
End Sub

Public Sub screen_disableMapPreview()
    mappreview = False
End Sub

' Screenshot code goes below this

'* This runs in a separate thread
' (VB) the ScreenshotSaveRequest's fields are its arguments; it runs at once
Private Function saveScreenshotThreadFunc(ByVal fileName As String, ByVal image As iV_Image) As Long
    Dim pngerror As String, ConsoleString As String
    pngerror = iV_saveImage_PNG(fileName, image)

'   //iV_saveImage_JPEG is *NOT* thread-safe, and cannot be safely called from another thread
'   iV_saveImage_JPEG(fileName, &image);

    If Len(pngerror) <> 0 Then
        ' dispatch logging the error to the main thread
        debug_ 472, LOG_ERROR, "saveScreenshotThreadFunc", "%s" & vbLf, pngerror
    Else
        ' display message to user about screenshot
        ' this must be dispatched back to the main thread
        ConsoleString = va("Screenshot %s saved!", fileName)
        addConsoleMessage ConsoleString, LEFT_JUSTIFY, SYSTEM_MESSAGE
    End If

    ' delete pSaveRequest: the onDelete lambda
    If Not image Is Nothing Then
        image.clear
    End If
    saveScreenshotThreadFunc = 0
End Function

'* Writes a screenshot of the current frame to file.
' *
' *  Performs the actual work of writing the frame currently displayed on screen
' *  to the filename specified by screenDumpToDisk().
' *
' *  @NOTE This function will only dump a screenshot to file if it was requested
' *        by screenDumpToDisk().
' *
' *  \sa screenDumpToDisk()
Public Sub screenDoDumpToDiskIfRequired()
    Dim fileName As String, bSentRequest As Boolean, image As iV_Image, ConsoleString As String
'   iV_Image image = { 0, 0, 8, nullptr };

    If Not screendump_required Then
        Exit Sub
    End If

    fileName = screendump_filename
    debug_ 513, LOG_3D, "screenDoDumpToDiskIfRequired", "Saving screenshot %s", fileName
    screendump_required = False

    bSentRequest = gfx_api_context_getScreenshot(image)

    If Not bSentRequest Then
        ' Backend / system does not support saving a screenshot (yet)
        debug_ 555, LOG_3D, "screenDoDumpToDiskIfRequired", "Failed to save screenshot %s", fileName
        Exit Sub
    End If

    ' the callback
    If image Is Nothing Then
        debug_ 520, LOG_3D, "screenDoDumpToDiskIfRequired", "Failed to get screenshot %s", fileName
        Exit Sub
    End If

    ' Dispatch encoding and saving screenshot to a background thread (since this is fairly costly)
    ConsoleString = va("Saving screenshot %s ...", fileName)
    addConsoleMessage ConsoleString, LEFT_JUSTIFY, INFO_MESSAGE
    ' (VB) wzThreadCreate(saveScreenshotThreadFunc, pSaveRequest): the thread function runs here
    saveScreenshotThreadFunc fileName, image
End Sub

'* Registers the currently displayed frame for making a screen shot.
' *
' *  The filename will be suffixed with a number, such that no files are
' *  overwritten.
' *
' *  \param path The directory path to save the screenshot in.
Public Sub screenDumpToDisk(ByVal path As String, ByVal level As String)
    Dim screendump_num As Long, aclock As Double, t As tm
    screendump_num = 0

    aclock = time_()        ' Get time in seconds
    t = getLocalTime(aclock)  ' Convert time to struct

    screendump_filename = va("%swz2100-%04d%02d%02d_%02d%02d%02d-%s.png", path, t.tm_year + 1900, t.tm_mon + 1, t.tm_mday, t.tm_hour, t.tm_min, t.tm_sec, level)

    Do While PHYSFS_exists(screendump_filename) <> 0
        screendump_num = screendump_num + 1
        screendump_filename = va("%swz2100-%04d%02d%02d_%02d%02d%02d-%s-%d.png", path, t.tm_year + 1900, t.tm_mon + 1, t.tm_mday, t.tm_hour, t.tm_min, t.tm_sec, level, screendump_num)
    Loop
    screendump_required = True
End Sub

Public Sub screen_FlipIfBackDropTransition()
    If hadBackDrop <> screen_GetBackDrop() Then
        pie_ScreenFrameRenderEnd
        pie_ScreenFrameRenderBegin
        hadBackDrop = screen_GetBackDrop()
    End If
End Sub
